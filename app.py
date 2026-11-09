import io
import os
import re
import html
import hashlib
import traceback
from xml.sax.saxutils import escape as xml_escape

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.exceptions import HTTPException
import psycopg2
from psycopg2 import errors as pg_errors
from psycopg2.extras import RealDictCursor, execute_values
import docx
import pdfplumber
import groq as groq_module
from groq import Groq
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

app = Flask(__name__)
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

DATABASE_URL = os.getenv("DATABASE_URL")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
# llama-3.1-8b-instant was shut down on 2026-08-16. Groq recommends this replacement.
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-20b")

groq_client = Groq(api_key=GROQ_API_KEY, timeout=60, max_retries=1) if GROQ_API_KEY else None

MAX_QUESTIONS = 80
MAX_PDF_PAGES = 80
MAX_NOTE_CHARS_PER_FILE = 200000


# ───────────────────────── database helpers ─────────────────────────

def get_conn():
    return psycopg2.connect(DATABASE_URL, connect_timeout=15)


def fetch_all(sql, params=()):
    conn = get_conn()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    finally:
        conn.close()


def fetch_one(sql, params=()):
    rows = fetch_all(sql, params)
    return rows[0] if rows else None


def execute(sql, params=(), returning=False):
    conn = get_conn()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            row = cur.fetchone() if returning else None
        conn.commit()
        return row
    finally:
        conn.close()


def init_db():
    try:
        conn = get_conn()
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    email TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS subjects (
                    id SERIAL PRIMARY KEY,
                    user_id INTEGER NOT NULL REFERENCES users(id),
                    name TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS questions (
                    id SERIAL PRIMARY KEY,
                    subject_id INTEGER NOT NULL REFERENCES subjects(id),
                    question_text TEXT NOT NULL,
                    marks INTEGER
                )""")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS answers (
                    id SERIAL PRIMARY KEY,
                    question_id INTEGER NOT NULL REFERENCES questions(id),
                    answer_text TEXT,
                    word_count INTEGER
                )""")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS note_chunks (
                    id SERIAL PRIMARY KEY,
                    subject_id INTEGER NOT NULL REFERENCES subjects(id),
                    source_file TEXT,
                    chunk_text TEXT NOT NULL
                )""")
        conn.commit()
        conn.close()
        print("Database initialized")
    except Exception as e:
        print(f"DB init error: {e}")


init_db()


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def to_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


# ───────────────────────── parsing ─────────────────────────

HEADER_WORDS = [
    "subject name", "subject code", "question bank", "department", "semester",
    "module", "unit", "section", "part a", "part b", "sl no", "sl.no", "q.no",
    "course code", "course name", "faculty", "year", "batch", "college",
    "ia1", "ia2", "ia3", "internal assessment", "answer all", "answer any",
    "max marks", "total marks", "instructions", "note:",
]
HEADER_RE = re.compile(r"\b(?:" + "|".join(re.escape(w) for w in HEADER_WORDS) + r")", re.I)
MARKS_RE = re.compile(r"\b(\d+)\s*[Mm]arks?\b|\[(\d+)\]|\((\d+)\)|\b(\d+)[Mm]\b")


def detect_marks(text, default_marks):
    m = MARKS_RE.search(text)
    if m:
        value = int(next(g for g in m.groups() if g))
        if 0 < value <= 20:
            return value
    return default_marks


def looks_like_question(text, seen):
    if len(text) < 12 or text in seen:
        return False
    words = text.split()
    if len(words) < 3:
        return False
    if HEADER_RE.match(text):
        return False
    if len(words) < 12 and HEADER_RE.search(text):
        return False
    return True


def parse_questions(streams, default_marks):
    questions, seen = [], set()
    for stream in streams:
        d = docx.Document(stream)
        texts = [p.text.strip() for p in d.paragraphs]
        for table in d.tables:
            for row in table.rows:
                for cell in row.cells:
                    texts.append(cell.text.strip())
        for text in texts:
            if looks_like_question(text, seen):
                seen.add(text)
                questions.append({"text": text, "marks": detect_marks(text, default_marks)})
    return questions[:MAX_QUESTIONS]


def extract_pdf_text(stream):
    parts, total = [], 0
    with pdfplumber.open(stream) as pdf:
        for i, page in enumerate(pdf.pages):
            if i >= MAX_PDF_PAGES or total >= MAX_NOTE_CHARS_PER_FILE:
                break
            text = page.extract_text() or ""
            if text.strip():
                parts.append(text)
                total += len(text)
            if hasattr(page, "flush_cache"):
                page.flush_cache()
    return "\n\n".join(parts)[:MAX_NOTE_CHARS_PER_FILE]


def chunk_text(text, size=1200):
    text = re.sub(r"[ \t]+", " ", text)
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    chunks, buf = [], ""
    for p in paras:
        while len(p) > size:
            cut = p.rfind(" ", 0, size)
            cut = cut if cut > 200 else size
            if buf:
                chunks.append(buf)
                buf = ""
            chunks.append(p[:cut].strip())
            p = p[cut:].strip()
        if not p:
            continue
        if buf and len(buf) + len(p) + 1 > size:
            chunks.append(buf)
            buf = p
        else:
            buf = (buf + "\n" + p).strip()
    if buf:
        chunks.append(buf)
    return chunks


# ───────────────────────── answer generation ─────────────────────────

STOPWORDS = set("""the a an of and or to in on for with is are was were be by as at from that this these those it
its what which who whom how why explain describe define discuss write short note notes brief briefly about give list
differentiate between compare mention state different types type example examples using use also can will""".split())


def keywords(text):
    words = re.findall(r"[a-zA-Z][a-zA-Z0-9\-]{2,}", text.lower())
    return {w for w in words if w not in STOPWORDS}


def pick_chunks(question_text, chunks, k=3, budget=4500):
    q = keywords(question_text)
    if not q:
        return ""
    scored = []
    for c in chunks:
        score = len(q & keywords(c))
        if score:
            scored.append((score, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    picked, used = [], 0
    for _, c in scored[:k]:
        if used + len(c) > budget:
            break
        picked.append(c)
        used += len(c)
    return "\n---\n".join(picked)


SYSTEM_PROMPT = ("You are an exam answer writer for engineering students. "
                 "Write only the final answer, directly, with no preamble or commentary.")


def ask_groq(question_text, marks, context):
    marks = marks if marks and marks > 0 else 10
    if marks <= 2:
        guide = "Write 3-5 lines. Give a clear definition and one key point."
    elif marks <= 5:
        guide = "Write 8-12 lines. Include a short explanation and one relevant example."
    else:
        guide = ("Write 20-30 lines. Include an introduction, detailed explanation in points, "
                 "examples, and a short conclusion.")
    notes_block = context or "(No relevant notes found. Answer from standard textbook knowledge.)"
    prompt = f"""Answer this exam question for an engineering student.

Reference notes (use these as the primary source):
{notes_block}

Question: {question_text}
Marks: {marks}
{guide}

Rules:
- Write ONLY the answer. No preamble, no thinking, no commentary.
- Use **bold** for headings.
- For comparisons, use a markdown table.
- For SQL or code, put each statement on its own line."""

    resp = groq_client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        max_tokens=3500,
        temperature=0.3,
    )
    return (resp.choices[0].message.content or "").strip()


def convert_tables(text):
    out, rows = [], []

    def flush():
        if not rows:
            return
        cells_html = []
        for i, cells in enumerate(rows):
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            tag = "th" if i == 0 else "td"
            cells_html.append("<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>")
        out.append("<table>" + "".join(cells_html) + "</table>")
        rows.clear()

    for line in text.split("\n"):
        s = line.strip()
        if s.startswith("|") and s.endswith("|") and len(s) > 2:
            rows.append([c.strip() for c in s.strip("|").split("|")])
        else:
            flush()
            out.append(line)
    flush()
    return "\n".join(out)


def clean_answer(raw):
    text = re.sub(r"<think>.*?</think>", "", raw or "", flags=re.DOTALL).strip()
    text = html.escape(text, quote=False)
    text = re.sub(r"^#{1,6}\s*(.+)$", r"<strong>\1</strong>", text, flags=re.MULTILINE)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<![\*\w])\*(?!\s)([^*\n]+?)(?<!\s)\*(?![\*\w])", r"<em>\1</em>", text)
    text = convert_tables(text)
    text = text.replace("\n\n", "<br><br>").replace("\n", "<br>")
    return text.strip()


# ───────────────────────── PDF export ─────────────────────────

def answer_to_pdf_markup(answer_html):
    t = answer_html or ""
    t = t.replace("</tr>", "<br/>")
    t = re.sub(r"</t[hd]>", " | ", t)
    t = t.replace("<br>", "<br/>")
    t = t.replace("<strong>", "<b>").replace("</strong>", "</b>")
    t = t.replace("<em>", "<i>").replace("</em>", "</i>")
    return re.sub(r"<(?!/?(?:b|i)>|br/>)[^>]*>", "", t)


def build_pdf(subject_name, rows, out):
    doc = SimpleDocTemplate(out, pagesize=A4, title=f"{subject_name} answers")
    styles = getSampleStyleSheet()
    story = [Paragraph(xml_escape(subject_name) + " - Answers", styles["Title"]), Spacer(1, 12)]
    for i, r in enumerate(rows, 1):
        marks = r["marks"] if r["marks"] else 10
        q = xml_escape(r["question_text"]).replace("\n", "<br/>")
        story.append(Paragraph(f"Q{i} ({marks} marks): {q}", styles["Heading3"]))
        story.append(Spacer(1, 4))
        markup = answer_to_pdf_markup(r["answer_text"]) or "<i>Answer not generated.</i>"
        try:
            story.append(Paragraph(markup, styles["Normal"]))
        except Exception:
            plain = xml_escape(re.sub(r"<[^>]+>", " ", html.unescape(markup)))
            story.append(Paragraph(plain, styles["Normal"]))
        story.append(Spacer(1, 14))
    doc.build(story)


# ───────────────────────── error handling ─────────────────────────

@app.errorhandler(Exception)
def handle_error(e):
    if isinstance(e, HTTPException):
        return jsonify(error=e.description), e.code
    traceback.print_exc()
    return jsonify(error="Server error: " + str(e)[:200]), 500


# ───────────────────────── routes ─────────────────────────

@app.route("/health")
def health():
    return jsonify(status="ok")


@app.route("/api/signup", methods=["POST"])
def signup():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip().lower()
    password = (data.get("password") or "").strip()
    if not email or not password:
        return jsonify(error="Email and password are required"), 400
    try:
        execute("INSERT INTO users (email, password_hash) VALUES (%s, %s)",
                (email, hash_password(password)))
    except pg_errors.UniqueViolation:
        return jsonify(error="Email already exists"), 400
    return jsonify(message="Signup successful"), 201


@app.route("/api/login", methods=["POST"])
def login():
    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip().lower()
    password = (data.get("password") or "").strip()
    if not email or not password:
        return jsonify(error="Email and password are required"), 400
    user = fetch_one("SELECT id FROM users WHERE lower(email)=%s AND password_hash=%s",
                     (email, hash_password(password)))
    if not user:
        return jsonify(error="Invalid email or password"), 401
    return jsonify(user_id=user["id"], message="Login successful")


@app.route("/api/subjects", methods=["GET"])
def list_subjects():
    user_id = to_int(request.args.get("user_id"))
    if not user_id:
        return jsonify(error="user_id required"), 400
    rows = fetch_all("SELECT id, name FROM subjects WHERE user_id=%s ORDER BY created_at DESC, id DESC",
                     (user_id,))
    return jsonify(subjects=rows)


@app.route("/api/subjects", methods=["POST"])
def create_subject():
    data = request.get_json(silent=True) or {}
    user_id = to_int(data.get("user_id"))
    name = (data.get("name") or "").strip()
    if not user_id or not name:
        return jsonify(error="user_id and name required"), 400
    row = execute("INSERT INTO subjects (user_id, name) VALUES (%s, %s) RETURNING id",
                  (user_id, name[:120]), returning=True)
    return jsonify(subject_id=row["id"], message="Subject created"), 201


@app.route("/api/subjects/<int:subject_id>", methods=["GET"])
def get_subject(subject_id):
    subject = fetch_one("SELECT id, name FROM subjects WHERE id=%s", (subject_id,))
    if not subject:
        return jsonify(error="Subject not found"), 404
    qa = fetch_all("""SELECT q.id, q.question_text, q.marks, a.answer_text
                      FROM questions q LEFT JOIN answers a ON a.question_id = q.id
                      WHERE q.subject_id=%s ORDER BY q.id""", (subject_id,))
    notes = fetch_one("SELECT COUNT(*) AS n FROM note_chunks WHERE subject_id=%s", (subject_id,))
    return jsonify(subject=subject, qa_pairs=qa, notes_count=notes["n"])


@app.route("/api/subjects/<int:subject_id>/questions", methods=["POST"])
def upload_questions(subject_id):
    files = [f for f in request.files.getlist("qbank") if f and f.filename]
    if not files:
        return jsonify(error="Please choose a question bank (.docx)"), 400
    default_marks = to_int(request.form.get("default_marks")) or 10
    streams = []
    for f in files[:3]:
        if not f.filename.lower().endswith(".docx"):
            return jsonify(error=f"{f.filename} is not a .docx file"), 400
        streams.append(io.BytesIO(f.read()))
    try:
        questions = parse_questions(streams, default_marks)
    except Exception as e:
        return jsonify(error=f"Could not read the question bank: {str(e)[:150]}"), 400
    if not questions:
        return jsonify(error="No questions found in the file"), 400

    conn = get_conn()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""DELETE FROM answers WHERE question_id IN
                           (SELECT id FROM questions WHERE subject_id=%s)""", (subject_id,))
            cur.execute("DELETE FROM questions WHERE subject_id=%s", (subject_id,))
            saved = []
            for q in questions:
                cur.execute("""INSERT INTO questions (subject_id, question_text, marks)
                               VALUES (%s, %s, %s) RETURNING id""",
                            (subject_id, q["text"], q["marks"]))
                saved.append({"id": cur.fetchone()["id"], "marks": q["marks"]})
        conn.commit()
    finally:
        conn.close()
    return jsonify(count=len(saved), questions=saved)


@app.route("/api/subjects/<int:subject_id>/notes", methods=["POST"])
def upload_notes(subject_id):
    f = request.files.get("notes")
    if not f or not f.filename:
        return jsonify(error="Please choose a PDF"), 400
    reset = request.form.get("reset") == "true"
    try:
        text = extract_pdf_text(io.BytesIO(f.read()))
    except Exception as e:
        return jsonify(error=f"Could not read {f.filename}: {str(e)[:150]}"), 400
    chunks = chunk_text(text)

    conn = get_conn()
    try:
        with conn.cursor() as cur:
            if reset:
                cur.execute("DELETE FROM note_chunks WHERE subject_id=%s", (subject_id,))
            if chunks:
                execute_values(cur,
                               "INSERT INTO note_chunks (subject_id, source_file, chunk_text) VALUES %s",
                               [(subject_id, f.filename[:200], c) for c in chunks])
        conn.commit()
    finally:
        conn.close()
    return jsonify(chunks=len(chunks), chars=len(text))


@app.route("/api/questions/<int:question_id>/generate", methods=["POST"])
def generate(question_id):
    if not groq_client:
        return jsonify(error="GROQ_API_KEY is not set on the server"), 500
    q = fetch_one("SELECT id, subject_id, question_text, marks FROM questions WHERE id=%s", (question_id,))
    if not q:
        return jsonify(error="Question not found"), 404
    chunks = [r["chunk_text"] for r in
              fetch_all("SELECT chunk_text FROM note_chunks WHERE subject_id=%s", (q["subject_id"],))]
    context = pick_chunks(q["question_text"], chunks)
    try:
        raw = ask_groq(q["question_text"], q["marks"], context)
    except Exception as e:
        rate_cls = getattr(groq_module, "RateLimitError", None)
        is_rate = (rate_cls is not None and isinstance(e, rate_cls)) or "429" in str(e)
        print(f"Groq error on question {question_id}: {e}")
        return jsonify(error=str(e)[:300]), (429 if is_rate else 502)

    answer = clean_answer(raw)
    if not answer:
        return jsonify(error="The model returned an empty answer. Please retry."), 502

    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM answers WHERE question_id=%s", (question_id,))
            cur.execute("INSERT INTO answers (question_id, answer_text, word_count) VALUES (%s, %s, %s)",
                        (question_id, answer, len(answer.split())))
        conn.commit()
    finally:
        conn.close()
    return jsonify(answer=answer)


@app.route("/api/download/<int:subject_id>", methods=["GET"])
def download(subject_id):
    subject = fetch_one("SELECT name FROM subjects WHERE id=%s", (subject_id,))
    if not subject:
        return jsonify(error="Subject not found"), 404
    rows = fetch_all("""SELECT q.question_text, q.marks, a.answer_text
                        FROM questions q LEFT JOIN answers a ON a.question_id = q.id
                        WHERE q.subject_id=%s ORDER BY q.id""", (subject_id,))
    if not rows:
        return jsonify(error="No questions yet"), 404
    buf = io.BytesIO()
    build_pdf(subject["name"], rows, buf)
    buf.seek(0)
    filename = re.sub(r"[^\w\-]+", "_", subject["name"]
