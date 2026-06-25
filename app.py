# -*- coding: utf-8 -*-
"""
Upgraded Technical Standards Search Portal
Features:
- Flask web app
- Login + role-based access
- SQLite database
- Upload PDF/TXT/CSV/XLSX/DOCX
- Deduplication by SHA256
- Requirement-level parsing
- Better Dutch-English search
- OCR fallback for scanned PDFs
- PDF preview and page jump (with inline page image)
- Table preview for CSV/XLSX
- Requirement full block display

Run:
    python app.py

Default URL:
    [127.0.0.1](http://127.0.0.1:5000)

Initial admin:
    set ADMIN_USERNAME and ADMIN_PASSWORD before first start

For production, also set:
    APP_ENV=production
    APP_SECRET_KEY=<strong-random-secret>
"""

import os
import re
import io
import csv
import math
import json
import time
import uuid
import hashlib
import secrets
import zipfile
import mimetypes
from datetime import datetime
from collections import defaultdict, Counter
from pathlib import Path

import pdfplumber
import pandas as pd

from flask import (
    Flask, request, render_template_string, redirect, url_for,
    send_file, flash, abort, session, jsonify
)
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

from flask_sqlalchemy import SQLAlchemy  # pyright: ignore[reportMissingImports]
from flask_login import (  # pyright: ignore[reportMissingImports]
    LoginManager, login_user, logout_user, login_required,
    current_user, UserMixin
)

from PIL import Image

def extract_text_from_image(path):
    if not TESSERACT_AVAILABLE:
        return "", [(0, 0, 1)], []

    try:
        image = Image.open(path)
        text = pytesseract.image_to_string(image)
        return text, [(0, len(text), 1)], []
    except Exception as e:
        print(f"Image OCR error: {path} -> {e}")
        return "", [(0, 0, 1)], []


# Optional packages
DOCX_AVAILABLE = True
try:
    from docx import Document as DocxDocument
except Exception:
    DOCX_AVAILABLE = False

OCR_AVAILABLE = True
try:
    import fitz  # PyMuPDF
except Exception:
    OCR_AVAILABLE = False

TESSERACT_AVAILABLE = True
PDF2IMAGE_AVAILABLE = False
try:
    import pytesseract
except Exception:
    TESSERACT_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False

SEMANTIC_AVAILABLE = True
try:
    from sentence_transformers import SentenceTransformer
except Exception:
    SEMANTIC_AVAILABLE = False


# =========================================================
# Base paths
# =========================================================
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

APP_ENV = os.environ.get("APP_ENV", "development").strip().lower() or "development"
IS_PRODUCTION = APP_ENV == "production"


def env_flag(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


SEMANTIC_SEARCH_ENABLED = env_flag("ENABLE_SEMANTIC_SEARCH", default=False)


DEFAULT_DATA_DIR = BASE_DIR / "data"
DATA_DIR = Path(os.environ.get("DATA_DIR", str(DEFAULT_DATA_DIR))).expanduser()
DOC_FOLDER = DATA_DIR / "documents"
PREVIEW_FOLDER = DATA_DIR / "previews"
DATABASE_PATH = Path(
    os.environ.get("DATABASE_PATH", str(DATA_DIR / "search_portal.db"))
).expanduser()

for path in (DATA_DIR, DOC_FOLDER, PREVIEW_FOLDER, DATABASE_PATH.parent):
    path.mkdir(parents=True, exist_ok=True)


def build_database_uri():
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if database_url:
        return database_url
    return f"sqlite:///{DATABASE_PATH}"


APP_SECRET_KEY = os.environ.get("APP_SECRET_KEY", "").strip()
if not APP_SECRET_KEY:
    if IS_PRODUCTION:
        raise RuntimeError("APP_SECRET_KEY must be set when APP_ENV=production.")
    APP_SECRET_KEY = "dev-secret-key-change-me"


# =========================================================
# Flask config
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = APP_SECRET_KEY
app.config["SQLALCHEMY_DATABASE_URI"] = build_database_uri()
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["REMEMBER_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
if IS_PRODUCTION:
    app.config["SESSION_COOKIE_SECURE"] = True
    app.config["REMEMBER_COOKIE_SECURE"] = True

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message_category = "warning"

# =========================================================
# Semantic model lazy load
# =========================================================
semantic_model = None

def get_semantic_model():
    global semantic_model
    if not SEMANTIC_SEARCH_ENABLED:
        return None
    if semantic_model is None and SEMANTIC_AVAILABLE:
        try:
            semantic_model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
        except Exception:
            semantic_model = None
    return semantic_model


# =========================================================
# Settings
# =========================================================
ALLOWED_EXTENSIONS = {
    "pdf", "txt", "csv", "xlsx", "docx",
    "png", "jpg", "jpeg"
}
if DOCX_AVAILABLE:
    ALLOWED_EXTENSIONS.add("docx")

ADMIN_ROLE = "admin"
USER_ROLE = "user"

STOPWORDS = {
    "a","an","and","are","as","at","be","by","for","from","has","he","in","is","it","its",
    "of","on","that","the","to","was","were","will","with","i","you","we","they","this",
    "these","those","or","but","so","not","can","could","would","should","may","might",
    "do","does","did","done","have","had","having","been","being","if","then","than",
    "there","here","about","into","over","under","de","het","een","en","van","op","te",
    "met","voor","dat","die","is","zijn","wordt","moet","dient","bij","of","in","aan",
    "tot","als","uit","door","niet","welke","welk"
}

MULTI_LANG_SYNONYMS = {
    "aarding": ["aarding", "aard", "geaard", "aardrail", "aardverbinding", "grounding", "earthing", "earth"],
    "grounding": ["grounding", "earthing", "earth", "aarding", "aard", "geaard", "aardrail", "aardverbinding"],
    "earthing": ["earthing", "grounding", "earth", "aarding", "aard", "geaard", "aardrail", "aardverbinding"],
    "earth": ["earth", "earthing", "grounding", "aarding", "aard", "geaard"],
    "resistance": ["resistance", "weerstand"],
    "weerstand": ["weerstand", "resistance"],
    "voltage": ["voltage", "spanning"],
    "spanning": ["spanning", "voltage"],
    "transformer": ["transformer", "transformator"],
    "transformator": ["transformator", "transformer"],
    "requirement": ["requirement", "requirements", "req", "eis", "eisen"],
    "req": ["req", "requirement", "eis"],
    "eis": ["eis", "requirement", "req"],
    "kabel": ["kabel", "cable"],
    "cable": ["cable", "kabel"],
    "bescherming": ["bescherming", "protection"],
    "protection": ["protection", "bescherming"],
    "installatie": ["installatie", "installation"],
    "installation": ["installation", "installatie"],
    "stroom": ["stroom", "current"],
    "current": ["current", "stroom"],
    "frequentie": ["frequentie", "frequency"],
    "frequency": ["frequency", "frequentie"],
}

REQ_ID_REGEX = re.compile(r'\b([A-Z]{1,10}-Req-\d+(?:\.\d+)?)\b', re.IGNORECASE)

# =========================================================
# Database models
# =========================================================
class User(db.Model, UserMixin):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default=USER_ROLE)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_admin(self):
        return self.role == ADMIN_ROLE


class DocumentRecord(db.Model):
    __tablename__ = "documents"

    id = db.Column(db.Integer, primary_key=True)
    original_filename = db.Column(db.String(500), nullable=False)
    stored_filename = db.Column(db.String(500), nullable=False, unique=True)
    extension = db.Column(db.String(20), nullable=False)
    file_hash = db.Column(db.String(64), nullable=False, unique=True, index=True)
    size_bytes = db.Column(db.Integer, nullable=False)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_ocr = db.Column(db.Boolean, default=False)
    text_preview = db.Column(db.Text, default="")
    status = db.Column(db.String(50), default="indexed")

    requirements = db.relationship("RequirementBlock", backref="document", cascade="all, delete-orphan")
    tables = db.relationship("TablePreview", backref="document", cascade="all, delete-orphan")


class RequirementBlock(db.Model):
    __tablename__ = "requirement_blocks"

    id = db.Column(db.Integer, primary_key=True)
    document_id = db.Column(db.Integer, db.ForeignKey("documents.id"), nullable=False, index=True)

    requirement_id = db.Column(db.String(120), index=True)
    title = db.Column(db.String(500), default="")
    section = db.Column(db.String(300), default="")
    page = db.Column(db.Integer, default=1)
    category = db.Column(db.String(120), default="General")
    definition = db.Column(db.Text, default="")
    summary = db.Column(db.Text, default="")
    full_text = db.Column(db.Text, nullable=False)

    token_blob = db.Column(db.Text, default="")
    text_hash = db.Column(db.String(64), index=True)
    semantic_text = db.Column(db.Text, default="")


class TablePreview(db.Model):
    __tablename__ = "table_previews"

    id = db.Column(db.Integer, primary_key=True)
    document_id = db.Column(db.Integer, db.ForeignKey("documents.id"), nullable=False, index=True)

    sheet_name = db.Column(db.String(255), default="")
    page = db.Column(db.Integer, default=1)
    table_format = db.Column(db.String(20), default="csv")
    html_table = db.Column(db.Text, default="")
    csv_text = db.Column(db.Text, default="")
    preview_title = db.Column(db.String(255), default="")


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# =========================================================
# Helpers
# =========================================================
def is_admin():
    return current_user.is_authenticated and current_user.is_admin

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def sha256_of_file(path, chunk_size=1024 * 1024):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()

def preprocess(text):
    text = (text or "").lower()
    tokens = re.findall(r"[a-zA-ZÀ-ÿ0-9\-_\.]+", text)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

def expand_query_tokens(tokens):
    out = []
    for t in tokens:
        out.extend(MULTI_LANG_SYNONYMS.get(t, [t]))
    return list(dict.fromkeys(out))

def detect_category(text, title=""):
    combined = f"{title} {text}".lower()
    if "emc" in combined and "aard" in combined:
        return "EMC grounding"
    if "aardrail" in combined:
        return "Earth rail"
    if "aardverbinding" in combined:
        return "Earth connection"
    if "aarding" in combined or "geaard" in combined or "grounding" in combined or "earthing" in combined:
        return "Grounding"
    if "klem" in combined and "aard" in combined:
        return "Grounding terminal"
    if "kast" in combined and ("aard" in combined or "geaard" in combined):
        return "Cabinet grounding"
    if "definitie" in combined or "definition" in combined:
        return "Definition"
    return "General"

def strip_html(text):
    return re.sub(r"<[^>]+>", "", text or "")

def highlight_terms(text, terms):
    if not text:
        return ""
    safe = text
    if terms:
        pattern = re.compile("|".join(re.escape(t) for t in sorted(set(terms), key=len, reverse=True)), re.IGNORECASE)
        safe = pattern.sub(lambda m: f'<span class="highlight">{m.group(0)}</span>', safe)
    return safe

def get_page_number(page_offsets, char_pos):
    for start, end, page_num in page_offsets:
        if start <= char_pos <= end:
            return page_num
    return 1

def file_ext(filename):
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""

def current_time_str():
    return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

# =========================================================
# PDF page image helper
# =========================================================
def get_page_image(document_path, page_num=1):
    """
    Convert a single PDF page to PNG image bytes.
    Returns (bytes, mimetype) or None if conversion fails.
    """
    if not PDF2IMAGE_AVAILABLE:
        return None
    try:
        # Convert only the requested page (page_num is 1-based)
        images = convert_from_path(
            document_path,
            first_page=page_num,
            last_page=page_num,
            dpi=150  # reasonable quality for preview
        )
        if not images:
            return None
        img_io = io.BytesIO()
        images[0].save(img_io, format='PNG')
        img_io.seek(0)
        return img_io.getvalue(), 'image/png'
    except Exception as e:
        print(f"Page image error: {document_path} page {page_num} -> {e}")
        return None
# =========================================================
# Text extraction
# =========================================================
def extract_text_from_pdf(path):
    full_text = ""
    page_offsets = []
    current = 0

    try:
        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                start = current
                full_text += text + "\n"
                current += len(text) + 1
                end = current
                page_offsets.append((start, end, page_num))
    except Exception as e:
        print(f"PDF extract error: {path} -> {e}")
        return "", [(0, 0, 1)]

    return full_text, page_offsets

def extract_text_from_txt(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()
    return text, [(0, len(text), 1)]

def extract_text_from_docx(path):
    if not DOCX_AVAILABLE:
        return "", [(0, 0, 1)]
    doc = DocxDocument(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    return text, [(0, len(text), 1)]

def extract_text_from_csv(path):
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
    except UnicodeDecodeError:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="latin1")
    text = df.fillna("").astype(str).to_string(index=False)
    return text, [(0, len(text), 1)], [("CSV", df)]
    
def extract_text_from_xlsx(path):
    sheets = pd.read_excel(path, sheet_name=None)
    all_parts = []
    tables = []
    for sheet_name, df in sheets.items():
        tables.append((sheet_name, df))
        all_parts.append(f"[Sheet: {sheet_name}]")
        all_parts.append(df.fillna("").astype(str).to_string(index=False))
    text = "\n".join(all_parts)
    return text, [(0, len(text), 1)], tables

def ocr_pdf(path):
    if not TESSERACT_AVAILABLE:
        return "", [(0, 0, 1)]

    full_text = ""
    page_offsets = []
    current = 0

    try:
        images = convert_from_path(path)
        for i, image in enumerate(images, start=1):
            text = pytesseract.image_to_string(image)
            start = current
            full_text += text + "\n"
            current += len(text) + 1
            end = current
            page_offsets.append((start, end, i))
    except Exception as e:
        print(f"OCR error: {path} -> {e}")
        return "", [(0, 0, 1)]

    return full_text, page_offsets

def extract_text_by_extension(path):
    ext = file_ext(path)

    if ext == "pdf":
        text, page_offsets = extract_text_from_pdf(path)
        ocr_text, ocr_offsets = ocr_pdf(path)

        used_ocr = False
        merged_text = text

        if ocr_text.strip():
            used_ocr = True
            merged_text = text + "\n\n[OCR_LAYER]\n" + ocr_text if text.strip() else ocr_text

            # if using merged text, page mapping can just use OCR offsets
            # or keep PDF offsets if native text is dominant
            if len(ocr_text) > len(text):
                page_offsets = ocr_offsets

        return merged_text, page_offsets, [], used_ocr


    if ext == "txt":
        text, page_offsets = extract_text_from_txt(path)
        return text, page_offsets, [], False

    if ext == "docx":
        text, page_offsets = extract_text_from_docx(path)
        return text, page_offsets, [], False

    if ext == "csv":
        text, page_offsets, tables = extract_text_from_csv(path)
        return text, page_offsets, tables, False

    if ext == "xlsx":
        text, page_offsets, tables = extract_text_from_xlsx(path)
        return text, page_offsets, tables, False

    if ext in {"png", "jpg", "jpeg"}:
        text, page_offsets, tables = extract_text_from_image(path)
        return text, page_offsets, tables, True

    return "", [(0, 0, 1)], [], False


# =========================================================
# Requirement parsing
# =========================================================
def find_definition_nearby(lines, start_index):
    for i in range(start_index, min(start_index + 8, len(lines))):
        line = lines[i].strip()
        low = line.lower()
        if low.startswith("definition") or low.startswith("definitie"):
            if i + 1 < len(lines):
                return lines[i + 1].strip()
        if "definition:" in low or "definitie:" in low:
            parts = re.split(r"definition:|definitie:", line, flags=re.IGNORECASE)
            if len(parts) > 1:
                return parts[1].strip()
    return ""

def parse_requirement_blocks(content, page_offsets):
    matches = list(REQ_ID_REGEX.finditer(content))
    blocks = []

    if matches:
        for i, m in enumerate(matches):
            start = m.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
            block_text = content[start:end].strip()
            req_id = m.group(1)
            page = get_page_number(page_offsets, start)

            lines = [line.strip() for line in block_text.splitlines() if line.strip()]
            title = ""
            summary = ""
            section = ""
            definition = ""

            if len(lines) >= 2:
                if req_id.lower() in lines[0].lower():
                    title = lines[1]
                else:
                    title = lines[0]

            for idx, line in enumerate(lines):
                low = line.lower()
                if re.match(r"^\d+(\.\d+)*\s+", line):
                    section = line
                    break
                if low.startswith(("de ", "het ", "een ", "opdrachtnemer ", "the ")):
                    summary = line
                    definition = find_definition_nearby(lines, idx)
                    break

            if not summary:
                for line in lines[1:6]:
                    if len(line) > 20:
                        summary = line
                        break

            if not definition:
                definition = find_definition_nearby(lines, 0)

            category = detect_category(block_text, title)
            blocks.append({
                "requirement_id": req_id,
                "title": title,
                "section": section,
                "page": page,
                "category": category,
                "definition": definition,
                "summary": summary,
                "full_text": block_text,
                "token_blob": " ".join(preprocess(block_text)),
            })
        return blocks

    # Fallback: paragraph blocks
    parts = re.split(r"\n\s*\n+", content)
    cursor = 0
    for part in parts:
        text = part.strip()
        if not text:
            continue

        pos = content.find(part, cursor)
        if pos == -1:
            pos = cursor
        page = get_page_number(page_offsets, pos)
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        title = lines[0][:160] if lines else ""
        category = detect_category(text, title)

        blocks.append({
            "requirement_id": "",
            "title": title,
            "section": "",
            "page": page,
            "category": category,
            "definition": "",
            "summary": lines[1][:500] if len(lines) > 1 else title,
            "full_text": text,
            "token_blob": " ".join(preprocess(text)),
        })
        cursor = pos + len(part)

    return blocks


# =========================================================
# Database indexing
# =========================================================
def create_default_admin():
    admin_username = os.environ.get("ADMIN_USERNAME", "admin").strip() or "admin"
    admin_password = os.environ.get("ADMIN_PASSWORD", "").strip()
    admin_auto_create = env_flag("ADMIN_AUTO_CREATE", default=True)

    if not admin_auto_create:
        print("Admin bootstrap skipped because ADMIN_AUTO_CREATE is disabled.")
        return

    if not admin_password:
        print(
            "Admin bootstrap skipped. Set ADMIN_USERNAME and ADMIN_PASSWORD "
            "before first start to create the initial admin account."
        )
        return

    user = User.query.filter_by(username=admin_username).first()
    if not user:
        user = User(username=admin_username, role=ADMIN_ROLE)
        user.set_password(admin_password)
        db.session.add(user)
        db.session.commit()
        print(f"Created initial admin user '{admin_username}'.")

def save_table_previews(doc_id, tables):
    for sheet_name, df in tables:
        if df is None:
            continue
        preview_df = df.head(30).copy()
        html_table = preview_df.to_html(index=False, classes="data-table", border=0)
        csv_buf = io.StringIO()
        preview_df.to_csv(csv_buf, index=False)
        row = TablePreview(
            document_id=doc_id,
            sheet_name=str(sheet_name),
            page=1,
            table_format="xlsx" if sheet_name != "CSV" else "csv",
            html_table=html_table,
            csv_text=csv_buf.getvalue(),
            preview_title=f"{sheet_name} preview"
        )
        db.session.add(row)

def remove_existing_blocks(doc_id):
    RequirementBlock.query.filter_by(document_id=doc_id).delete()
    TablePreview.query.filter_by(document_id=doc_id).delete()
    db.session.commit()

def index_document_record(doc_record):
    full_path = os.path.join(DOC_FOLDER, doc_record.stored_filename)
    text, page_offsets, tables, used_ocr = extract_text_by_extension(full_path)

    doc_record.is_ocr = used_ocr
    doc_record.text_preview = text[:2000]
    doc_record.status = "indexed"
    db.session.commit()

    remove_existing_blocks(doc_record.id)

    blocks = parse_requirement_blocks(text, page_offsets)
    seen_hashes = set()

    for b in blocks:
        th = hashlib.sha256(
    b["full_text"].encode("utf-8", errors="ignore")
).hexdigest()
        if th in seen_hashes:
            continue
        seen_hashes.add(th)

        row = RequirementBlock(
            document_id=doc_record.id,
            requirement_id=b["requirement_id"],
            title=b["title"],
            section=b["section"],
            page=b["page"],
            category=b["category"],
            definition=b["definition"],
            summary=b["summary"],
            full_text=b["full_text"],
            token_blob=b["token_blob"],
            text_hash=th,
            semantic_text=" ".join([b["title"], b["summary"], b["definition"], b["full_text"][:2000]]),
        )
        db.session.add(row)

    save_table_previews(doc_record.id, tables)
    db.session.commit()

def dedupe_lookup(file_hash):
    return DocumentRecord.query.filter_by(file_hash=file_hash).first()


# =========================================================
# Search
# =========================================================
def lexical_score(block, query, expanded_tokens):
    score = 0.0
    q = query.lower().strip()

    doc = block.document
    filename = (doc.original_filename or "").lower()
    stored_filename = (doc.stored_filename or "").lower()
    filename_no_ext = filename.rsplit(".", 1)[0] if "." in filename else filename

    title = (block.title or "").lower()
    summary = (block.summary or "").lower()
    definition = (block.definition or "").lower()
    full_text = (block.full_text or "").lower()
    req_id = (block.requirement_id or "").lower()
    section = (block.section or "").lower()
    category = (block.category or "").lower()
    tokens = set((block.token_blob or "").split())

    # Strong exact filename matches
    if q and q == filename:
        score += 100
    if q and q == stored_filename:
        score += 100
    if q and q == filename_no_ext:
        score += 95

    # Partial filename matches
    if q and q in filename:
        score += 40
    if q and q in stored_filename:
        score += 35
    if q and q in filename_no_ext:
        score += 35

    # Existing scoring
    if q and q == req_id:
        score += 30
    if q and q in title:
        score += 16
    if q and q in summary:
        score += 12
    if q and q in definition:
        score += 12
    if q and q in section:
        score += 8
    if q and q in category:
        score += 6
    if q and q in full_text:
        score += 8

    for term in set(expanded_tokens):
        if term in filename:
            score += 10
        if term in filename_no_ext:
            score += 10
        if term in title:
            score += 6
        if term in summary:
            score += 5
        if term in definition:
            score += 5
        if term in category:
            score += 3
        if term in tokens:
            score += 2

    if any(t.startswith("aard") for t in expanded_tokens) and "ground" in category:
        score += 3

    return score

def semantic_score(query, block):
    model = get_semantic_model()
    if not model:
        return 0.0
    try:
        q_emb = model.encode([query])[0]
        b_emb = model.encode([block.semantic_text[:3000]])[0]

        dot = float(sum(a * b for a, b in zip(q_emb, b_emb)))
        q_norm = math.sqrt(sum(a * a for a in q_emb))
        b_norm = math.sqrt(sum(a * a for a in b_emb))
        if q_norm == 0 or b_norm == 0:
            return 0.0
        return (dot / (q_norm * b_norm)) * 20.0
    except Exception:
        return 0.0

def get_result_snippet(block, expanded_tokens):
    text = block.full_text or ""
    text_lower = text.lower()
    hit_pos = -1

    for t in expanded_tokens:
        pos = text_lower.find(t.lower())
        if pos != -1 and (hit_pos == -1 or pos < hit_pos):
            hit_pos = pos

    window = 900
    if hit_pos == -1:
        snippet = text[:window]
        if len(text) > window:
            snippet += "..."
    else:
        start = max(0, hit_pos - 250)
        end = min(len(text), hit_pos + 650)
        snippet = text[start:end]
        if start > 0:
            snippet = "..." + snippet
        if end < len(text):
            snippet = snippet + "..."

    return highlight_terms(snippet, expanded_tokens)
    

    # First: exact/near filename matches
    direct_docs = DocumentRecord.query.all()
    matched_doc_ids = []

    for d in direct_docs:
        original_name = (d.original_filename or "").lower()
        stored_name = (d.stored_filename or "").lower()
        base_name = original_name.rsplit(".", 1)[0] if "." in original_name else original_name

        if q in {original_name, stored_name, base_name}:
            matched_doc_ids.append(d.id)

    rows = RequirementBlock.query.all()

    scored = []
    for row in rows:
        score = lexical_score(row, query, expanded_tokens)

        if row.document_id in matched_doc_ids:
            score += 120

        score += semantic_score(query, row)

        if score > 0:
            scored.append((row, score))

    scored.sort(key=lambda x: x[1], reverse=True)
    results = []

    for row, score in scored[:top_k]:
        doc = row.document
        results.append({
            "block_id": row.id,
            "document_id": doc.id,
            "filename": doc.original_filename,
            "stored_filename": doc.stored_filename,
            "page": row.page,
            "requirement_id": row.requirement_id,
            "title": row.title,
            "section": row.section,
            "category": row.category,
            "definition": row.definition,
            "summary": row.summary,
            "full_text": row.full_text,
            "snippet": get_result_snippet(row, expanded_tokens),
            "relevance": round(score, 2),
            "is_pdf": doc.extension.lower() == "pdf",
            "has_table": doc.extension.lower() in {"csv", "xlsx"},
            "ocr_used": doc.is_ocr,
        })

    return results, expanded_tokens

def grouped_search_summary(results):
    category_counts = Counter(r["category"] for r in results if r.get("category"))
    top_categories = category_counts.most_common(5)
    req_ids = [r["requirement_id"] for r in results if r.get("requirement_id")]
    return {
        "top_categories": top_categories,
        "top_requirement_ids": req_ids[:8]
    }

def search_requirements(query, top_k=30):
    if not query.strip():
        return [], []

    query_tokens = preprocess(query)
    expanded_tokens = expand_query_tokens(query_tokens)
    q = query.lower().strip()

    direct_docs = DocumentRecord.query.all()
    matched_doc_ids = []

    for d in direct_docs:
        original_name = (d.original_filename or "").lower()
        stored_name = (d.stored_filename or "").lower()
        base_name = original_name.rsplit(".", 1)[0] if "." in original_name else original_name

        if q in {original_name, stored_name, base_name}:
            matched_doc_ids.append(d.id)

    rows = RequirementBlock.query.all()

    scored = []
    for row in rows:
        score = lexical_score(row, query, expanded_tokens)

        if row.document_id in matched_doc_ids:
            score += 120

        score += semantic_score(query, row)

        if score > 0:
            scored.append((row, score))

    scored.sort(key=lambda x: x[1], reverse=True)
    results = []

    for row, score in scored[:top_k]:
        doc = row.document
        results.append({
            "block_id": row.id,
            "document_id": doc.id,
            "filename": doc.original_filename,
            "stored_filename": doc.stored_filename,
            "page": row.page,
            "requirement_id": row.requirement_id,
            "title": row.title,
            "section": row.section,
            "category": row.category,
            "definition": row.definition,
            "summary": row.summary,
            "full_text": row.full_text,
            "snippet": get_result_snippet(row, expanded_tokens),
            "relevance": round(score, 2),
            "is_pdf": doc.extension.lower() == "pdf",
            "has_table": doc.extension.lower() in {"csv", "xlsx"},
            "ocr_used": doc.is_ocr,
            "is_image": doc.extension.lower() in {"png", "jpg", "jpeg"},
        })

    return results, expanded_tokens
   

def search_documents(query, top_k=10):
    q = query.lower().strip()
    docs = DocumentRecord.query.all()
    scored = []

    for d in docs:
        filename = (d.original_filename or "").lower()
        stored = (d.stored_filename or "").lower()
        base = filename.rsplit(".", 1)[0] if "." in filename else filename

        score = 0
        if q == filename:
            score += 100
        if q == stored:
            score += 100
        if q == base:
            score += 95
        if q in filename:
            score += 40
        if q in stored:
            score += 35
        if q in base:
            score += 35

        if score > 0:
            scored.append((d, score))

    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:top_k]


# =========================================================
# Templates
# =========================================================
BASE_CSS = """
<style>
body { font-family: Arial, sans-serif; margin: 24px; background:#f5f7fb; color:#222; }
.container { max-width: 1280px; margin:auto; background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }
.topbar { display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; }
.actions { display:flex; gap:8px; flex-wrap:wrap; }
input[type=text], input[type=password], input[type=file], select { padding:10px; font-size:15px; }
input[type=text] { width:68%; min-width:280px; }
button, .btn { padding:10px 16px; border:none; border-radius:6px; text-decoration:none; color:white; cursor:pointer; display:inline-block; }
button { background:#0069d9; }
.btn-green { background:#28a745; }
.btn-purple { background:#6f42c1; }
.btn-orange { background:#fd7e14; }
.btn-gray { background:#6c757d; }
.notice, .flash, .warning, .info { padding:12px; border-radius:6px; margin:12px 0; }
.notice { background:#fff3cd; border-left:4px solid #ffc107; }
.flash { background:#e9f7ef; border-left:4px solid #28a745; }
.warning { background:#fdecea; border-left:4px solid #dc3545; }
.info { background:#eef6ff; border-left:4px solid #339af0; }
.result-item { margin-top:16px; padding:16px; border-left:4px solid #0069d9; background:#fafafa; border-radius:6px; }
.filename { font-size:18px; font-weight:bold; color:#b00020; }
.meta { color:#555; margin:8px 0; font-size:14px; }
.summary-box { margin-top:16px; padding:12px; background:#f8f9fa; border:1px solid #ececec; border-radius:6px; }
.snippet, .full-block { background:#fff; padding:12px; white-space:pre-wrap; border:1px solid #eee; border-radius:6px; font-family:Consolas, monospace; }
.highlight { color:green; font-weight:bold; background:#eaf7ea; padding:1px 2px; border-radius:2px; }
.grid { display:grid; grid-template-columns: 1fr 1fr; gap:16px; }
.data-table table, table.data-table { width:100%; border-collapse:collapse; }
.data-table th, .data-table td, table.data-table th, table.data-table td { border:1px solid #ddd; padding:6px 8px; font-size:14px; vertical-align:top; }
a { color:#0069d9; text-decoration:none; }
a:hover { text-decoration:underline; }
.small { font-size:13px; color:#666; }
.badge { display:inline-block; background:#eef; color:#334; padding:3px 8px; border-radius:12px; font-size:12px; margin-right:6px; }
.page-img-preview { max-width:200px; max-height:200px; border:1px solid #ccc; margin-right:16px; float:left; cursor:pointer; transition: transform 0.2s; }
.page-img-preview:hover { transform: scale(1.02); }
.clearfix::after { content:""; display:table; clear:both; }
@media (max-width: 900px) { .grid { grid-template-columns: 1fr; } input[type=text] { width:100%; } }
</style>
"""

HOME_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Technical Standards Search Portal</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <div class="topbar">
        <div>
            <h1>Technical Standards Search Portal</h1>
            <div class="small">Requirement-level search for PDF, TXT, CSV, XLSX, DOCX</div>
        </div>
        <div class="actions">
            {% if current_user.is_authenticated %}
                <a class="btn btn-gray" href="{{ url_for('logout') }}">Logout ({{ current_user.username }})</a>
                {% if current_user.is_admin %}
                    <a class="btn btn-purple" href="{{ url_for('upload_files') }}">Upload</a>
                    <a class="btn btn-orange" href="{{ url_for('admin_documents') }}">Documents</a>
                    <a class="btn btn-green" href="{{ url_for('admin_users') }}">Users</a>
                {% endif %}
            {% else %}
                <a class="btn btn-gray" href="{{ url_for('login') }}">Login</a>
            {% endif %}
        </div>
    </div>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}
          <div class="flash">{{ msg }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="notice">
        Examples: <strong>aarding</strong>, <strong>grounding</strong>, <strong>earthing resistance</strong>, <strong>AM-Req-6165</strong>
    </div>

    <form method="GET" action="/">
        <input type="text" name="q" value="{{ query }}" placeholder="Search requirement ID, Dutch/English keyword, section, definition">
        <button type="submit">Search</button>
    </form>

    <div class="summary-box">
        <strong>Indexed documents:</strong> {{ doc_count }} |
        <strong>Requirement blocks:</strong> {{ block_count }} |
        <strong>Tables:</strong> {{ table_count }}
    </div>

    {% if query %}
        <div class="summary-box">
            <strong>Expanded terms:</strong> {{ expanded_terms|join(', ') if expanded_terms else '-' }}
            {% if summary.top_categories %}
            <br><strong>Top categories:</strong>
            {% for cat, cnt in summary.top_categories %}
                <span class="badge">{{ cat }} ({{ cnt }})</span>
            {% endfor %}
            {% endif %}
            {% if summary.top_requirement_ids %}
            <br><strong>Top requirement IDs:</strong> {{ summary.top_requirement_ids|join(', ') }}
            {% endif %}
        </div>

        <h3>Results for "{{ query }}"</h3>

<div style="margin:12px 0 18px 0;">
    <a href="{{ url_for('export_search_csv', q=query) }}"
       style="display:inline-block; padding:10px 16px; background:#198754; color:#fff; text-decoration:none; border-radius:8px; font-weight:700; margin-right:8px;">
        Export CSV
    </a>
    <a href="{{ url_for('export_search_xlsx', q=query) }}"
       style="display:inline-block; padding:10px 16px; background:#0d6efd; color:#fff; text-decoration:none; border-radius:8px; font-weight:700;">
        Export Excel
    </a>
</div>

        {% if results %}
            {% for res in results %}
                <div class="result-item">
                    <div class="filename">
                        {% if res.is_pdf %}
                            <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">
                                {{ res.requirement_id or '-' }} - {{ res.title or res.filename }}
                            </a>
                        {% else %}
                            {{ res.requirement_id or '-' }} - {{ res.title or res.filename }}
                        {% endif %}
                    </div>

                    <div class="meta">
                        File: <strong>{{ res.filename }}</strong>
                        | Page: <strong>{{ res.page or '?' }}</strong>
                        | Category: <strong>{{ res.category }}</strong>
                        | Relevance: <strong>{{ res.relevance }}</strong>
                        {% if res.ocr_used %}| <strong>OCR used</strong>{% endif %}
                    </div>

                    {% if res.definition %}
                        <div><strong>Definition:</strong> {{ res.definition }}</div>
                    {% endif %}
                    {% if res.summary %}
                        <div><strong>Summary:</strong> {{ res.summary }}</div>
                    {% endif %}
                    {% if res.section %}
                        <div><strong>Section:</strong> {{ res.section }}</div>
                    {% endif %}

                    <div class="snippet">{{ res.snippet|safe }}</div>

                    <div class="meta">
                        <a href="{{ url_for('requirement_detail', block_id=res.block_id) }}">Open full requirement</a>
                        {% if res.is_pdf %}
                            | <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open PDF at page {{ res.page or 1 }}</a>
                        {% endif %}
                        {% if res.has_table %}
                            | <a href="{{ url_for('table_preview', document_id=res.document_id) }}">Open table preview</a>
                        {% endif %}
                    </div>
                </div>
            {% endfor %}
        {% else %}
            <div class="warning">No matching results found.</div>
        {% endif %}
    {% endif %}
</div>
</body>
</html>
"""

# Requirement detail template also gets image preview
REQUIREMENT_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Requirement detail</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>{{ block.requirement_id or 'Requirement block' }}</h2>

    <div class="summary-box">
        <div><strong>Title:</strong> {{ block.title or '-' }}</div>
        <div><strong>Document:</strong> {{ block.document.original_filename }}</div>
        <div><strong>Page:</strong> {{ block.page }}</div>
        <div><strong>Category:</strong> {{ block.category }}</div>
        {% if block.section %}<div><strong>Section:</strong> {{ block.section }}</div>{% endif %}
        {% if block.definition %}<div><strong>Definition:</strong> {{ block.definition }}</div>{% endif %}
        {% if block.summary %}<div><strong>Summary:</strong> {{ block.summary }}</div>{% endif %}
    </div>

    <!-- NEW: page image preview for PDFs in requirement detail -->
    {% if block.document.extension == 'pdf' and pdf2image_available %}
        <div style="margin: 16px 0;">
            <a href="{{ url_for('pdf_viewer', document_id=block.document.id, page=block.page or 1) }}" target="_blank">
                <img src="{{ url_for('page_image', document_id=block.document.id, page=block.page or 1) }}"
                     style="max-width:100%; max-height:500px; border:1px solid #ccc;"
                     alt="Page {{ block.page or 1 }} preview"
                     onerror="this.style.display='none'">
            </a>
        </div>
    {% endif %}

    <h3>Full requirement text</h3>
    <div class="full-block">{{ block.full_text }}</div>

    <div style="margin-top:16px;">
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
        {% if block.document.extension == 'pdf' %}
            <a class="btn btn-purple" href="{{ url_for('pdf_viewer', document_id=block.document.id, page=block.page or 1) }}" target="_blank">Open PDF at page {{ block.page or 1 }}</a>
        {% endif %}
    </div>
</div>
</body>
</html>
"""

#LOGIN_TEMPLATE
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Login</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container" style="max-width: 520px;">
    <h2>Login</h2>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}
          <div class="warning">{{ msg }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <form method="POST">
        <div style="margin-bottom:12px;">
            <label>Username</label><br>
            <input type="text" name="username" required style="width:100%;">
        </div>
        <div style="margin-bottom:12px;">
            <label>Password</label><br>
            <input type="password" name="password" required style="width:100%;">
        </div>
        <button type="submit">Login</button>
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
    </form>

    <div class="notice" style="margin-top:16px;">
        Create the first admin user by setting <code>ADMIN_USERNAME</code> and
        <code>ADMIN_PASSWORD</code> before the first start.
    </div>
</div>
</body>
</html>
"""

UPLOAD_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Upload files</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>Upload files</h2>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}
          <div class="flash">{{ msg }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="info">
        Allowed types: {{ allowed_extensions }}<br>
        Max size: 100 MB per request
    </div>

    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="files" multiple required>
        <br><br>
        <button type="submit">Upload and index</button>
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
    </form>
</div>
</body>
</html>
"""

DOCS_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Documents</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>Indexed documents</h2>

    <div style="margin-bottom:16px;">
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
    </div>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}
          <div class="flash">{{ msg }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <form method="POST" action="{{ url_for('bulk_delete_documents') }}" onsubmit="return confirmBulkDelete();">
        <div style="margin-bottom:12px; display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
            <button type="submit" class="btn btn-orange">Delete Selected</button>
            <button type="button" class="btn btn-gray" onclick="toggleAll(true)">Select All</button>
            <button type="button" class="btn btn-gray" onclick="toggleAll(false)">Clear All</button>
            <span class="small">Select multiple documents and delete them in one action.</span>
        </div>

        <table class="data-table" style="margin-top:16px; width:100%;">
            <tr>
                <th style="width:40px;">
                    <input type="checkbox" onclick="toggleFromHeader(this)">
                </th>
                <th>ID</th>
                <th>Filename</th>
                <th>Type</th>
                <th>Size</th>
                <th>OCR</th>
                <th>Status</th>
                <th>Uploaded</th>
                <th>Actions</th>
            </tr>
            {% for d in docs %}
            <tr>
                <td>
                    <input type="checkbox" name="document_ids" value="{{ d.id }}" class="doc-checkbox">
                </td>
                <td>{{ d.id }}</td>
                <td>{{ d.original_filename }}</td>
                <td>{{ d.extension }}</td>
                <td>{{ d.size_bytes }}</td>
                <td>{{ 'Yes' if d.is_ocr else 'No' }}</td>
                <td>{{ d.status }}</td>
                <td>{{ d.uploaded_at }}</td>
                <td>
                    <a href="{{ url_for('reindex_document', document_id=d.id) }}">Reindex</a>
                    |
                    <a href="{{ url_for('delete_document', document_id=d.id) }}" onclick="return confirm('Delete this document?');">Delete</a>
                    {% if d.extension == 'pdf' %}
                    | <a href="{{ url_for('pdf_viewer', document_id=d.id, page=1) }}" target="_blank">View PDF</a>
                    {% endif %}
                    {% if d.extension in ['csv','xlsx'] %}
                    | <a href="{{ url_for('table_preview', document_id=d.id) }}">Preview table</a>
                    {% endif %}
                </td>
            </tr>
            {% endfor %}
        </table>
    </form>
</div>

<script>
function toggleAll(checked) {
    const boxes = document.querySelectorAll('.doc-checkbox');
    boxes.forEach(box => box.checked = checked);
}

function toggleFromHeader(headerCheckbox) {
    toggleAll(headerCheckbox.checked);
}

function confirmBulkDelete() {
    const checked = document.querySelectorAll('.doc-checkbox:checked');
    if (checked.length === 0) {
        alert('Please select at least one document to delete.');
        return false;
    }
    return confirm('Delete ' + checked.length + ' selected document(s)? This cannot be undone.');
}
</script>
</body>
</html>
"""

   

USERS_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Users</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>User management</h2>

    <form method="POST" style="margin-bottom:24px;">
        <div class="grid">
            <div>
                <label>Username</label><br>
                <input type="text" name="username" required style="width:100%;">
            </div>
            <div>
                <label>Password</label><br>
                <input type="password" name="password" required style="width:100%;">
            </div>
        </div>
        <div style="margin-top:12px;">
            <label>Role</label><br>
            <select name="role">
                <option value="user">user</option>
                <option value="admin">admin</option>
            </select>
        </div>
        <br>
        <button type="submit">Create user</button>
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
    </form>

    <table class="data-table">
        <tr>
            <th>ID</th>
            <th>Username</th>
            <th>Role</th>
            <th>Created</th>
        </tr>
        {% for u in users %}
        <tr>
            <td>{{ u.id }}</td>
            <td>{{ u.username }}</td>
            <td>{{ u.role }}</td>
            <td>{{ u.created_at }}</td>
        </tr>
        {% endfor %}
    </table>
</div>
</body>
</html>
"""

REQUIREMENT_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Requirement detail</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>{{ block.requirement_id or 'Requirement block' }}</h2>

    <div class="summary-box">
        <div><strong>Title:</strong> {{ block.title or '-' }}</div>
        <div><strong>Document:</strong> {{ block.document.original_filename }}</div>
        <div><strong>Page:</strong> {{ block.page }}</div>
        <div><strong>Category:</strong> {{ block.category }}</div>
        {% if block.section %}<div><strong>Section:</strong> {{ block.section }}</div>{% endif %}
        {% if block.definition %}<div><strong>Definition:</strong> {{ block.definition }}</div>{% endif %}
        {% if block.summary %}<div><strong>Summary:</strong> {{ block.summary }}</div>{% endif %}
    </div>

    <h3>Full requirement text</h3>
    <div class="full-block">{{ block.full_text }}</div>

    <div style="margin-top:16px;">
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
        {% if block.document.extension == 'pdf' %}
            <a class="btn btn-purple" href="{{ url_for('pdf_viewer', document_id=block.document.id, page=block.page or 1) }}" target="_blank">Open PDF at page {{ block.page or 1 }}</a>
        {% endif %}
    </div>
</div>
</body>
</html>
"""

PDF_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>PDF Viewer</title>
""" + BASE_CSS + """
<style>
body { margin:0; }
.toolbar { padding:12px 16px; background:white; border-bottom:1px solid #ddd; display:flex; justify-content:space-between; align-items:center; }
.viewer-container { height: calc(100vh - 60px); }
iframe { width:100%; height:100%; border:none; }
</style>
</head>
<body>
    <div class="toolbar">
        <div><strong>{{ filename }}</strong> | Page {{ page }}</div>
        <div>
            <a href="{{ url_for('home') }}">Back</a>
        </div>
    </div>
    <div class="viewer-container">
        <iframe src="{{ pdf_url }}#page={{ page }}"></iframe>
    </div>
</body>
</html>
"""

TABLE_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Table preview</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>Table preview - {{ doc.original_filename }}</h2>
    <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>

    {% if tables %}
        {% for t in tables %}
            <div class="result-item">
                <div class="filename">{{ t.preview_title }}</div>
                <div class="meta">Format: {{ t.table_format }} | Sheet: {{ t.sheet_name }}</div>
                <div class="data-table">{{ t.html_table|safe }}</div>
            </div>
        {% endfor %}
    {% else %}
        <div class="warning">No table preview available.</div>
    {% endif %}


</div>
</body>
</html>
"""


# =========================================================
# Routes
# =========================================================
@app.route("/")
def home():
    query = request.args.get("q", "").strip()
    results = []
    expanded_terms = []
    summary = {"top_categories": [], "top_requirement_ids": []}
    document_matches = []

    if query:
        document_matches = search_documents(query)
        results, expanded_terms = search_requirements(query)
        summary = grouped_search_summary(results)

    return render_template_string(
        HOME_TEMPLATE,
        query=query,
        results=results,
        document_matches=document_matches,
        expanded_terms=expanded_terms,
        summary=summary,
        doc_count=DocumentRecord.query.count(),
        block_count=RequirementBlock.query.count(),
        table_count=TablePreview.query.count(),
    )


@app.route("/healthz")
def healthz():
    return jsonify(
        {
            "status": "ok",
            "environment": APP_ENV,
            "data_dir": str(DATA_DIR),
            "database_path": str(DATABASE_PATH),
        }
    )

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter_by(username=username).first()
        if not user or not user.check_password(password):
            flash("Invalid username or password.")
            return redirect(url_for("login"))

        login_user(user)
        flash("Login successful.")
        return redirect(url_for("home"))

    return render_template_string(LOGIN_TEMPLATE)


@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Logged out.")
    return redirect(url_for("home"))


@app.route("/admin/users", methods=["GET", "POST"])
@login_required
def admin_users():
    if not is_admin():
        abort(403)

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", USER_ROLE)

        if not username or not password:
            flash("Username and password are required.")
        elif User.query.filter_by(username=username).first():
            flash("Username already exists.")
        else:
            new_user = User(username=username, role=role)
            new_user.set_password(password)
            db.session.add(new_user)
            db.session.commit()
            flash(f"User '{username}' created.")

        return redirect(url_for("admin_users"))

    users = User.query.order_by(User.id).all()
    return render_template_string(USERS_TEMPLATE, users=users)

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload_files():
    if not is_admin():
        abort(403)

    if request.method == "POST":
        files = request.files.getlist("files")
        if not files or all(f.filename == "" for f in files):
            flash("No files selected.")
            return redirect(url_for("upload_files"))

        uploaded_count = 0
        skipped_count = 0

        for file in files:
            if not file or file.filename == "":
                continue
            if not allowed_file(file.filename):
                flash(f"File type not allowed: {file.filename}")
                continue

            original_filename = secure_filename(file.filename)
            ext = file_ext(original_filename)

            # Save file temporarily to compute hash
            temp_path = os.path.join(DOC_FOLDER, f"temp_{uuid.uuid4().hex}.{ext}")
            file.save(temp_path)
            file_hash = sha256_of_file(temp_path)

            existing = dedupe_lookup(file_hash)
            if existing:
                os.remove(temp_path)
                flash(f"Duplicate file skipped: {original_filename}")
                skipped_count += 1
                continue

            # Move to permanent storage
            stored_filename = f"{uuid.uuid4().hex}.{ext}"
            final_path = os.path.join(DOC_FOLDER, stored_filename)
            os.rename(temp_path, final_path)

            doc = DocumentRecord(
                original_filename=original_filename,
                stored_filename=stored_filename,
                extension=ext,
                file_hash=file_hash,
                size_bytes=os.path.getsize(final_path),
                uploaded_by=current_user.id,
            )
            db.session.add(doc)
            db.session.commit()
            try:
                index_document_record(doc)
            except Exception as e:
                flash(f"Indexing error for {original_filename}: {str(e)}")

            uploaded_count += 1

        flash(f"Uploaded {uploaded_count} file(s), skipped {skipped_count} duplicate(s).")
        return redirect(url_for("upload_files"))

    return render_template_string(
        UPLOAD_TEMPLATE,
        allowed_extensions=", ".join(sorted(ALLOWED_EXTENSIONS))
    )

@app.route("/admin/documents")
@login_required
def admin_documents():
    if not is_admin():
        abort(403)

    docs = DocumentRecord.query.order_by(DocumentRecord.uploaded_at.desc()).all()
    return render_template_string(DOCS_TEMPLATE, docs=docs)

@app.route("/reindex/<int:document_id>")
@login_required
def reindex_document(document_id):
    if not is_admin():
        abort(403)

    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        flash("Document not found.")
        return redirect(url_for("admin_documents"))

    try:
        index_document_record(doc)
        flash(f"Reindexed: {doc.original_filename}")
    except Exception as e:
        flash(f"Reindex error: {str(e)}")

    return redirect(url_for("admin_documents"))

@app.route("/delete/<int:document_id>")
@login_required
def delete_document(document_id):
    if not is_admin():
        abort(403)

    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        flash("Document not found.")
        return redirect(url_for("admin_documents"))

    # Delete physical file
    file_path = os.path.join(DOC_FOLDER, doc.stored_filename)
    if os.path.exists(file_path):
        os.remove(file_path)

    # DB cascade will remove requirements & tables
    db.session.delete(doc)
    db.session.commit()
    flash(f"Deleted: {doc.original_filename}")
    return redirect(url_for("admin_documents"))

@app.route("/requirement/<int:block_id>")
def requirement_detail(block_id):
    block = db.session.get(RequirementBlock, block_id)
    if not block:
        abort(404)
    return render_template_string(
        REQUIREMENT_TEMPLATE,
        block=block,
        pdf2image_available=PDF2IMAGE_AVAILABLE
    )

@app.route("/table/<int:document_id>")
def table_preview(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        abort(404)

    tables = TablePreview.query.filter_by(document_id=document_id).all()
    return render_template_string(TABLE_TEMPLATE, doc=doc, tables=tables)

@app.route("/pdf/<int:document_id>")
def pdf_viewer(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)

    page = request.args.get("page", 1, type=int)
    pdf_url = url_for("serve_pdf", document_id=document_id)
    return render_template_string(
        PDF_TEMPLATE,
        filename=doc.original_filename,
        page=page,
        pdf_url=pdf_url
    )

@app.route("/serve-pdf/<int:document_id>")
def serve_pdf(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)

    file_path = os.path.join(DOC_FOLDER, doc.stored_filename)
    if not os.path.exists(file_path):
        abort(404)

    return send_file(file_path, mimetype="application/pdf")

@app.route("/page-image/<int:document_id>")
def page_image(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)
    page = request.args.get("page", 1, type=int)
    file_path = os.path.join(DOC_FOLDER, doc.stored_filename)
    if not os.path.exists(file_path):
        abort(404)
    image_data = get_page_image(file_path, page)
    if image_data is None:
        # fallback: return a 1x1 transparent PNG (or 404)
        from flask import Response
        return Response(status=404)
    return send_file(io.BytesIO(image_data[0]), mimetype=image_data[1])

@app.route("/export-search-csv")
def export_search_csv():
    query = request.args.get("q", "").strip()
    if not query:
        flash("Please enter a search query first.")
        return redirect(url_for("home"))

    results, expanded_terms = search_requirements(query, top_k=1000)

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "query",
        "expanded_terms",
        "relevance",
        "requirement_id",
        "title",
        "summary",
        "section",
        "category",
        "definition",
        "page",
        "filename",
        "full_text",
        "snippet",
        "is_pdf",
        "has_table",
        "ocr_used"
    ])

    for res in results:
        writer.writerow([
            query,
            ", ".join(expanded_terms),
            res.get("relevance", ""),
            res.get("requirement_id", ""),
            res.get("title", ""),
            res.get("summary", ""),
            res.get("section", ""),
            res.get("category", ""),
            res.get("definition", ""),
            res.get("page", ""),
            res.get("filename", ""),
            res.get("full_text", ""),
            re.sub(r"<[^>]+>", "", res.get("snippet", "")),  # remove highlight html
            "yes" if res.get("is_pdf") else "no",
            "yes" if res.get("has_table") else "no",
            "yes" if res.get("ocr_used") else "no",
        ])

    csv_bytes = io.BytesIO(output.getvalue().encode("utf-8-sig"))
    output.close()

    safe_query = re.sub(r"[^a-zA-Z0-9_-]+", "_", query)[:80] or "search"
    filename = f"search_results_{safe_query}.csv"

    return send_file(
        csv_bytes,
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=filename
    )

@app.route("/export-search-xlsx")
def export_search_xlsx():
    query = request.args.get("q", "").strip()
    if not query:
        flash("Please enter a search query first.")
        return redirect(url_for("home"))

    results, expanded_terms = search_requirements(query, top_k=5000)

    rows = []
    for res in results:
        rows.append({
            "query": query,
            "expanded_terms": ", ".join(expanded_terms),
            "relevance": res.get("relevance", ""),
            "requirement_id": res.get("requirement_id", ""),
            "title": res.get("title", ""),
            "summary": res.get("summary", ""),
            "section": res.get("section", ""),
            "category": res.get("category", ""),
            "definition": res.get("definition", ""),
            "page": res.get("page", ""),
            "filename": res.get("filename", ""),
            "full_text": res.get("full_text", ""),
            "snippet": re.sub(r"<[^>]+>", "", res.get("snippet", "")),
            "is_pdf": "yes" if res.get("is_pdf") else "no",
            "has_table": "yes" if res.get("has_table") else "no",
            "ocr_used": "yes" if res.get("ocr_used") else "no",
            "is_image": "yes" if res.get("is_image") else "no",
        })

    df = pd.DataFrame(rows)

    xlsx_buffer = io.BytesIO()
    with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="search_results")

    xlsx_buffer.seek(0)

    safe_query = re.sub(r"[^a-zA-Z0-9_-]+", "_", query)[:80] or "search"
    filename = f"search_results_{safe_query}.xlsx"

    return send_file(
        xlsx_buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

@app.route("/delete-multiple", methods=["POST"])
@login_required
def bulk_delete_documents():
    if not is_admin():
        abort(403)

    raw_ids = request.form.getlist("document_ids")
    if not raw_ids:
        flash("No documents selected.")
        return redirect(url_for("admin_documents"))

    deleted_count = 0
    missing_count = 0
    error_count = 0

    for raw_id in raw_ids:
        try:
            document_id = int(raw_id)
        except ValueError:
            error_count += 1
            continue

        doc = db.session.get(DocumentRecord, document_id)
        if not doc:
            missing_count += 1
            continue

        try:
            file_path = os.path.join(DOC_FOLDER, doc.stored_filename)
            if os.path.exists(file_path):
                os.remove(file_path)

            db.session.delete(doc)
            db.session.commit()
            deleted_count += 1
        except Exception as e:
            db.session.rollback()
            print(f"Bulk delete error for document {document_id}: {e}")
            error_count += 1

    flash(
        f"Bulk delete completed. Deleted: {deleted_count}, Missing: {missing_count}, Errors: {error_count}."
    )
    return redirect(url_for("admin_documents"))



# =========================================================
# App initialization
# =========================================================
with app.app_context():
    db.create_all()
    create_default_admin()

def _running_in_notebook():
    try:
        from IPython import get_ipython
        return get_ipython() is not None
    except Exception:
        return False


in_notebook = _running_in_notebook()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    debug = env_flag("FLASK_DEBUG", default=not IS_PRODUCTION and not in_notebook)
    app.run(
        host="0.0.0.0",
        port=port,
        debug=debug,
        use_reloader=debug,
    )
