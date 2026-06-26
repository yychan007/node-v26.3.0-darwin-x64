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

import storage
import translation

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


def env_int(name, default):
    raw = os.environ.get(name)
    if raw is None or str(raw).strip() == "":
        return default
    return int(raw)


SEMANTIC_SEARCH_ENABLED = env_flag("ENABLE_SEMANTIC_SEARCH", default=False)
OCR_ENABLED = env_flag("ENABLE_OCR", default=False)
INDEX_BATCH_SIZE = int(os.environ.get("INDEX_BATCH_SIZE", "50"))
MAX_PDF_PAGES = env_int("MAX_PDF_PAGES", 100 if IS_PRODUCTION else 0)
MAX_INDEX_FILE_MB = env_int("MAX_INDEX_FILE_MB", 40 if IS_PRODUCTION else 0)
TEXT_PREVIEW_LIMIT = env_int("TEXT_PREVIEW_LIMIT", 150000)
TRANSLATE_MAX_CELLS = env_int("TRANSLATE_MAX_CELLS", 180)


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
        if database_url.startswith("postgres://"):
            database_url = database_url.replace(
                "postgres://", "postgresql+psycopg://", 1
            )
        elif database_url.startswith("postgresql://"):
            database_url = database_url.replace(
                "postgresql://", "postgresql+psycopg://", 1
            )
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


@app.context_processor
def inject_translation_context():
    return {
        "translation_enabled": translation.translation_enabled(),
        "translation_provider_info": get_translation_status(),
    }


def get_translation_status():
    status_fn = getattr(translation, "provider_status", None)
    if callable(status_fn):
        return status_fn()
    provider_name = get_translation_provider()
    return {
        "enabled": translation.translation_enabled(),
        "provider": provider_name,
        "configured": True,
        "model": "google-translate",
    }


def get_translation_provider():
    provider_fn = getattr(translation, "translation_provider", None)
    return provider_fn() if callable(provider_fn) else "google"

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
    "kabel": ["kabel", "cable", "cables"],
    "cable": ["cable", "cables", "kabel", "kabels"],
    "cables": ["cables", "cable", "kabel", "kabels"],
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
    page_offsets_json = db.Column(db.Text, default="")
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
    char_start = db.Column(db.Integer, default=0)
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
    html_table_en = db.Column(db.Text, default="")
    csv_text = db.Column(db.Text, default="")
    preview_title = db.Column(db.String(255), default="")


class TranslationCache(db.Model):
    __tablename__ = "translation_cache"

    text_hash = db.Column(db.String(64), primary_key=True)
    source_text = db.Column(db.Text, default="")
    translated_text = db.Column(db.Text, default="")
    provider = db.Column(db.String(32), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class DocumentPageTranslation(db.Model):
    __tablename__ = "document_page_translations"
    __table_args__ = (
        db.UniqueConstraint("document_id", "page", name="uq_document_page_translation"),
    )

    id = db.Column(db.Integer, primary_key=True)
    document_id = db.Column(
        db.Integer, db.ForeignKey("documents.id", ondelete="CASCADE"), nullable=False
    )
    page = db.Column(db.Integer, nullable=False)
    source_text = db.Column(db.Text, default="")
    translated_text = db.Column(db.Text, default="")
    provider = db.Column(db.String(32), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


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
        variants = list(MULTI_LANG_SYNONYMS.get(t, [t]))
        if len(t) > 3 and t.endswith("s"):
            variants.append(t[:-1])
        elif len(t) > 2 and not t.endswith("s"):
            variants.append(t + "s")
        out.extend(variants)
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
    if "cable" in combined or "kabel" in combined:
        return "Cables"
    return "General"

def strip_html(text):
    return re.sub(r"<[^>]+>", "", text or "")


def get_cached_translation(text_hash):
    row = TranslationCache.query.get(text_hash)
    return row.translated_text if row else None


def store_cached_translation(text_hash, source_text, translated_text, provider=None):
    if not translated_text:
        return
    row = TranslationCache.query.get(text_hash)
    if row:
        return
    db.session.add(
        TranslationCache(
            text_hash=text_hash,
            source_text=source_text[:5000],
            translated_text=translated_text,
            provider=provider or get_translation_provider(),
        )
    )
    try:
        db.session.commit()
    except Exception:
        db.session.rollback()


def extract_single_page_text(doc, page_num):
    offsets = load_page_offsets(doc)
    text = doc.text_preview or ""
    if offsets and text:
        for start, end, page in offsets:
            if page == page_num:
                return text[start:end].strip()

    if doc.extension != "pdf" or not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        return ""

    if not OCR_AVAILABLE:
        return ""

    try:
        with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
            with fitz.open(file_path) as pdf_doc:
                if page_num < 1 or page_num > len(pdf_doc):
                    return ""
                return (pdf_doc.load_page(page_num - 1).get_text("text") or "").strip()
    except Exception as exc:
        print(f"Page text extract error for doc {doc.id} page {page_num}: {exc}")
        return ""


def get_or_translate_page(doc, page_num):
    row = DocumentPageTranslation.query.filter_by(
        document_id=doc.id, page=page_num
    ).first()
    if row and row.translated_text:
        return {
            "page": page_num,
            "source": row.source_text,
            "translation": row.translated_text,
            "provider": row.provider,
            "cached": True,
        }

    source = extract_single_page_text(doc, page_num)
    if not source:
        return {
            "page": page_num,
            "source": "",
            "translation": "",
            "provider": get_translation_provider(),
            "cached": False,
            "error": "No extractable text on this page.",
        }

    translated = translate_to_english(source)
    provider = get_translation_provider()
    if translated:
        if row:
            row.source_text = source[:20000]
            row.translated_text = translated
            row.provider = provider
        else:
            db.session.add(
                DocumentPageTranslation(
                    document_id=doc.id,
                    page=page_num,
                    source_text=source[:20000],
                    translated_text=translated,
                    provider=provider,
                )
            )
        try:
            db.session.commit()
        except Exception:
            db.session.rollback()

    return {
        "page": page_num,
        "source": source,
        "translation": translated,
        "provider": provider,
        "cached": False,
    }


def count_pdf_pages(doc):
    offsets = load_page_offsets(doc)
    if offsets:
        return max(page for _, _, page in offsets)

    if doc.extension != "pdf" or not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        return 0

    if not OCR_AVAILABLE:
        return 0

    try:
        with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
            with fitz.open(file_path) as pdf_doc:
                return len(pdf_doc)
    except Exception as exc:
        print(f"PDF page count error for doc {doc.id}: {exc}")
        return 0


def _split_text_for_pages(text, max_chars=3200):
    text = text or ""
    if len(text) <= max_chars:
        return [text] if text else [""]

    parts = []
    start = 0
    while start < len(text):
        end = min(len(text), start + max_chars)
        if end < len(text):
            split_at = text.rfind("\n", start, end)
            if split_at > start + 400:
                end = split_at + 1
        parts.append(text[start:end])
        start = end
    return parts


def _append_translation_text(pdf_doc, header, text):
    page_width, page_height = fitz.paper_size("a4")
    margin_x = 50
    margin_top = 70
    margin_bottom = 50
    body_rect = fitz.Rect(
        margin_x, margin_top, page_width - margin_x, page_height - margin_bottom
    )
    chunks = _split_text_for_pages(text)

    for index, chunk in enumerate(chunks):
        page = pdf_doc.new_page(width=page_width, height=page_height)
        if index == 0 and header:
            page.insert_text(
                (margin_x, 42),
                header,
                fontsize=11,
                fontname="helv",
            )
        page.insert_textbox(
            body_rect,
            chunk,
            fontsize=10,
            fontname="helv",
            align=fitz.TEXT_ALIGN_LEFT,
        )


def build_translated_pdf_bytes(doc):
    if not OCR_AVAILABLE:
        raise RuntimeError("PyMuPDF is required to export translated PDFs.")

    if not translation.translation_enabled():
        raise RuntimeError("Translation is disabled.")

    total_pages = count_pdf_pages(doc)
    if total_pages <= 0:
        raise ValueError("Could not determine PDF page count.")

    export_limit = MAX_PDF_PAGES if MAX_PDF_PAGES > 0 else total_pages
    pages_to_export = min(total_pages, export_limit)
    truncated = pages_to_export < total_pages

    pdf_out = fitz.open()
    source_name = doc.original_filename or doc.stored_filename

    try:
        if truncated:
            notice = (
                f"English translation export for {source_name}\n"
                f"Pages 1-{pages_to_export} of {total_pages}\n\n"
            )
            _append_translation_text(pdf_out, "Export notice", notice)

        for page_num in range(1, pages_to_export + 1):
            result = get_or_translate_page(doc, page_num)
            translation_text = (result.get("translation") or "").strip()
            if not translation_text:
                translation_text = "(No translation available for this page.)"

            header = f"{source_name} - Page {page_num} (English)"
            _append_translation_text(pdf_out, header, translation_text)

        if pdf_out.page_count == 0:
            raise ValueError("No translated pages were generated.")

        buffer = io.BytesIO()
        pdf_out.save(buffer, garbage=4, deflate=True)
        buffer.seek(0)
        return buffer.getvalue()
    finally:
        pdf_out.close()


def translate_to_english(text):
    return translation.translate_text(
        text,
        cache_get=get_cached_translation,
        cache_set=store_cached_translation,
    )


def build_english_snippet(snippet_html):
    plain = strip_html(snippet_html)
    translated = translate_to_english(plain)
    if not translated or translated == plain:
        return ""
    return highlight_terms(translated, [])


def enrich_result_with_translation(result_dict):
    result_dict["snippet_en"] = build_english_snippet(result_dict.get("snippet", ""))
    return result_dict


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
    if not page_offsets:
        return 1
    for start, end, page_num in page_offsets:
        if start <= char_pos < end:
            return page_num
    if char_pos >= page_offsets[-1][0]:
        return page_offsets[-1][2]
    return 1


def load_page_offsets(doc):
    raw = getattr(doc, "page_offsets_json", None) or ""
    if not raw:
        return []
    try:
        data = json.loads(raw)
        return [(int(a), int(b), int(c)) for a, b, c in data]
    except Exception:
        return []


def find_hit_position(text, expanded_tokens, query=""):
    text_lower = (text or "").lower()
    q = (query or "").lower().strip()
    if q and q in text_lower:
        return text_lower.find(q)

    hit_pos = -1
    for term in expanded_tokens:
        pos = text_lower.find(term.lower())
        if pos != -1 and (hit_pos == -1 or pos < hit_pos):
            hit_pos = pos
    return hit_pos


def resolve_result_page(doc, char_start, body_text, expanded_tokens, query, fallback_page=1):
    offsets = load_page_offsets(doc)
    if not offsets:
        return fallback_page or 1

    hit_in_body = find_hit_position(body_text, expanded_tokens, query)
    anchor = (char_start or 0) + hit_in_body if hit_in_body >= 0 else (char_start or 0)
    if anchor < 0:
        return fallback_page or 1
    return get_page_number(offsets, anchor)

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
def ensure_indexable_file(path):
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Stored file not found: {file_path}")

    if MAX_INDEX_FILE_MB > 0:
        size_mb = file_path.stat().st_size / (1024 * 1024)
        if size_mb > MAX_INDEX_FILE_MB:
            raise ValueError(
                f"File is {size_mb:.1f} MB; this server limits indexing to "
                f"{MAX_INDEX_FILE_MB} MB. Split the PDF or upgrade the hosting plan."
            )


def extract_text_from_pdf_fitz(path, max_pages=0):
    if not OCR_AVAILABLE:
        return None

    full_text = ""
    page_offsets = []
    current = 0

    try:
        with fitz.open(path) as doc:
            total_pages = len(doc)
            page_limit = total_pages if max_pages <= 0 else min(total_pages, max_pages)
            for page_num in range(page_limit):
                page = doc.load_page(page_num)
                text = page.get_text("text") or ""
                start = current
                full_text += text + "\n"
                current += len(text) + 1
                page_offsets.append((start, current, page_num + 1))
        return full_text, page_offsets
    except Exception as e:
        print(f"PyMuPDF extract error: {path} -> {e}")
        return None


def extract_text_from_pdf(path):
    fitz_result = extract_text_from_pdf_fitz(path, MAX_PDF_PAGES)
    if fitz_result is not None:
        return fitz_result

    full_text = ""
    page_offsets = []
    current = 0

    try:
        with pdfplumber.open(path) as pdf:
            pages = pdf.pages
            if MAX_PDF_PAGES > 0:
                pages = pages[:MAX_PDF_PAGES]
            for page_num, page in enumerate(pages, start=1):
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
        used_ocr = False
        merged_text = text

        if not OCR_ENABLED:
            return merged_text, page_offsets, [], used_ocr

        ocr_text, ocr_offsets = ocr_pdf(path)
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

def normalize_requirement_text(content):
    """Repair common PDF line-break splits inside requirement IDs."""
    text = content or ""
    text = re.sub(
        r"([A-Z]{1,10}-Req-)\s+(\d+(?:\.\d+)?)",
        r"\1\2",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"([A-Z]{1,10}-Req-\d+(?:\.\d+)?)\s+-\s+",
        r"\1 - ",
        text,
        flags=re.IGNORECASE,
    )
    return text


SECTION_HEADER_REGEX = re.compile(r"(?m)^(\d+(?:\.\d+)+)\s+(.+)$")


def parse_section_blocks(content, page_offsets):
    matches = list(SECTION_HEADER_REGEX.finditer(content))
    if not matches:
        return []

    blocks = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        block_text = content[start:end].strip()
        if len(block_text) < 40:
            continue

        section_num = match.group(1)
        section_title = match.group(2).strip()
        page = get_page_number(page_offsets, start)
        lines = [line.strip() for line in block_text.splitlines() if line.strip()]
        summary = ""
        for line in lines[1:6]:
            if len(line) > 25:
                summary = line
                break

        blocks.append({
            "requirement_id": f"Section-{section_num}",
            "title": section_title,
            "section": f"{section_num} {section_title}",
            "page": page,
            "char_start": start,
            "category": detect_category(block_text, section_title),
            "definition": "",
            "summary": summary,
            "full_text": block_text,
            "token_blob": " ".join(preprocess(block_text)),
        })

    return blocks


def parse_requirement_blocks(content, page_offsets):
    content = normalize_requirement_text(content)
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
                "char_start": start,
                "category": category,
                "definition": definition,
                "summary": summary,
                "full_text": block_text,
                "token_blob": " ".join(preprocess(block_text)),
            })
        return blocks

    section_blocks = parse_section_blocks(content, page_offsets)
    if section_blocks:
        return section_blocks

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
            "char_start": pos,
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
        html_table_en = ""
        if translation.translation_enabled():
            try:
                en_df = translation.translate_dataframe_values(
                    preview_df,
                    cache_get=get_cached_translation,
                    cache_set=store_cached_translation,
                    max_cells=TRANSLATE_MAX_CELLS,
                )
                html_table_en = en_df.to_html(index=False, classes="data-table", border=0)
            except Exception as exc:
                print(f"Table translation skipped for sheet {sheet_name}: {exc}")
        csv_buf = io.StringIO()
        preview_df.to_csv(csv_buf, index=False)
        row = TablePreview(
            document_id=doc_id,
            sheet_name=str(sheet_name),
            page=1,
            table_format="xlsx" if sheet_name != "CSV" else "csv",
            html_table=html_table,
            html_table_en=html_table_en,
            csv_text=csv_buf.getvalue(),
            preview_title=f"{sheet_name} preview"
        )
        db.session.add(row)


def regenerate_table_previews(doc_record):
    ext = doc_record.extension.lower()
    if ext not in {"csv", "xlsx"}:
        return []

    if not storage.document_exists(doc_record.stored_filename, DOC_FOLDER):
        raise FileNotFoundError(
            f"Stored file not found: {doc_record.stored_filename}"
        )

    with storage.open_document_local_path(
        doc_record.stored_filename, DOC_FOLDER
    ) as full_path:
        ensure_indexable_file(full_path)
        if ext == "xlsx":
            _, _, tables = extract_text_from_xlsx(full_path)
        else:
            _, _, tables = extract_text_from_csv(full_path)

    if not tables:
        return []

    TablePreview.query.filter_by(document_id=doc_record.id).delete()
    save_table_previews(doc_record.id, tables)
    db.session.commit()
    return TablePreview.query.filter_by(document_id=doc_record.id).all()

def remove_existing_blocks(doc_id):
    RequirementBlock.query.filter_by(document_id=doc_id).delete()
    TablePreview.query.filter_by(document_id=doc_id).delete()
    db.session.commit()

def index_document_record(doc_record):
    if not storage.document_exists(doc_record.stored_filename, DOC_FOLDER):
        raise FileNotFoundError(
            f"Stored file not found: {doc_record.stored_filename}. "
            "The upload may have been lost after a service restart."
        )

    with storage.open_document_local_path(
        doc_record.stored_filename, DOC_FOLDER
    ) as full_path:
        ensure_indexable_file(full_path)
        text, page_offsets, tables, used_ocr = extract_text_by_extension(full_path)

    doc_record.is_ocr = used_ocr
    doc_record.text_preview = text[:TEXT_PREVIEW_LIMIT]
    doc_record.page_offsets_json = json.dumps(page_offsets)
    doc_record.status = "indexed"
    db.session.commit()

    remove_existing_blocks(doc_record.id)

    blocks = parse_requirement_blocks(text, page_offsets)
    seen_hashes = set()
    pending = 0

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
            char_start=b.get("char_start", 0),
            category=b["category"],
            definition=b["definition"],
            summary=b["summary"],
            full_text=b["full_text"],
            token_blob=b["token_blob"],
            text_hash=th,
            semantic_text=" ".join([b["title"], b["summary"], b["definition"], b["full_text"][:2000]]),
        )
        db.session.add(row)
        pending += 1
        if pending >= INDEX_BATCH_SIZE:
            db.session.commit()
            pending = 0

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
        score += 20
    if q and len(q.split()) > 1 and q in title:
        score += 15
    if q and q in summary:
        score += 12
    if q and q in definition:
        score += 12
    if q and q in section:
        score += 10
    if q and q in category:
        score += 6
    if q and q in full_text:
        score += 10
    if q and len(q.split()) > 1 and q in full_text:
        score += 12

    query_tokens = preprocess(query)
    if len(query_tokens) > 1 and all(t in full_text for t in query_tokens):
        score += 10
    if len(query_tokens) > 1 and all(t in title for t in query_tokens):
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


def get_text_snippet(text, expanded_tokens, query=""):
    text = text or ""
    text_lower = text.lower()
    hit_pos = find_hit_position(text, expanded_tokens, query)

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
            snippet += "..."

    return highlight_terms(snippet, expanded_tokens)


def score_document_text(doc, query, expanded_tokens):
    q = query.lower().strip()
    text = doc.text_preview or ""
    if not text.strip():
        return 0

    text_lower = text.lower()
    filename = (doc.original_filename or "").lower()
    base_name = filename.rsplit(".", 1)[0] if "." in filename else filename
    score = 0.0

    if q and q in text_lower:
        score += 18
    if q and q in filename:
        score += 25
    if q and q in base_name:
        score += 20

    matched_terms = 0
    for term in set(expanded_tokens):
        if term.lower() in text_lower:
            matched_terms += 1
            score += 5
        if term.lower() in filename or term.lower() in base_name:
            score += 4

    query_tokens = preprocess(query)
    if len(query_tokens) > 1 and all(t in text_lower for t in query_tokens):
        score += 14

    return score


def build_document_fallback_results(query, expanded_tokens, seen_keys, top_k=10):
    results = []
    for doc in DocumentRecord.query.all():
        score = score_document_text(doc, query, expanded_tokens)
        if score <= 0:
            continue

        page = resolve_result_page(
            doc,
            0,
            doc.text_preview,
            expanded_tokens,
            query,
            1,
        )
        key = (doc.id, page, doc.original_filename)
        if key in seen_keys:
            continue
        seen_keys.add(key)

        results.append(enrich_result_with_translation({
            "block_id": None,
            "document_id": doc.id,
            "filename": doc.original_filename,
            "stored_filename": doc.stored_filename,
            "page": page,
            "requirement_id": "",
            "title": doc.original_filename,
            "section": "Document match",
            "category": "Document",
            "definition": "",
            "summary": (doc.text_preview or "")[:240],
            "full_text": doc.text_preview or "",
            "snippet": get_text_snippet(doc.text_preview, expanded_tokens, query),
            "relevance": round(score, 2),
            "is_pdf": doc.extension.lower() == "pdf",
            "has_table": doc.extension.lower() in {"csv", "xlsx"},
            "ocr_used": doc.is_ocr,
            "is_image": doc.extension.lower() in {"png", "jpg", "jpeg"},
            "is_document_fallback": True,
        }))

    results.sort(key=lambda r: r["relevance"], reverse=True)
    return results[:top_k]


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
    seen_keys = set()

    for row, score in scored[:top_k]:
        doc = row.document
        seen_keys.add((doc.id, row.page or 1, row.requirement_id or row.title))
        display_page = resolve_result_page(
            doc,
            row.char_start or 0,
            row.full_text,
            expanded_tokens,
            query,
            row.page,
        )
        results.append(enrich_result_with_translation({
            "block_id": row.id,
            "document_id": doc.id,
            "filename": doc.original_filename,
            "stored_filename": doc.stored_filename,
            "page": display_page,
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
            "is_document_fallback": False,
        }))

    if len(results) < top_k:
        fallback = build_document_fallback_results(
            query, expanded_tokens, seen_keys, top_k=top_k - len(results)
        )
        results.extend(fallback)

    results.sort(key=lambda r: r["relevance"], reverse=True)
    return results[:top_k], expanded_tokens


def search_documents(query, top_k=10):
    q = query.lower().strip()
    query_tokens = preprocess(query)
    expanded_tokens = expand_query_tokens(query_tokens)
    scored = []

    for doc in DocumentRecord.query.all():
        score = score_document_text(doc, query, expanded_tokens)
        filename = (doc.original_filename or "").lower()
        stored = (doc.stored_filename or "").lower()
        base = filename.rsplit(".", 1)[0] if "." in filename else filename

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
            scored.append((doc, score))

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
                {% if translation_enabled %}
                <a class="btn btn-green" href="{{ url_for('ai_translate_tool') }}">AI Translate</a>
                {% endif %}
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

    {% if doc_count == 0 %}
        <div class="warning">
            No documents are indexed yet. Log in as admin, upload PDF files, then use <strong>Reindex</strong> in Documents.
        </div>
    {% endif %}

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

                    <div class="snippet">
                        <strong>Table 1 / Dutch (Original)</strong>
                        <div>{{ res.snippet|safe }}</div>
                    </div>
                    {% if res.snippet_en %}
                    <div class="snippet" style="margin-top:10px; border-left:4px solid #28a745;">
                        <strong>Table 2 / English (Translation)</strong>
                        <div>{{ res.snippet_en|safe }}</div>
                    </div>
                    {% endif %}

                    <div class="meta">
                        {% if res.is_document_fallback %}
                            <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open PDF at page {{ res.page or 1 }}</a>
                        {% else %}
                            <a href="{{ url_for('requirement_detail', block_id=res.block_id) }}">Open full requirement</a>
                            {% if res.is_pdf %}
                                | <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open PDF at page {{ res.page or 1 }}</a>
                            {% endif %}
                            {% if res.has_table %}
                                | <a href="{{ url_for('table_preview', document_id=res.document_id) }}">Open table preview</a>
                            {% endif %}
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
                    | <a href="{{ url_for('export_translation_pdf', document_id=d.id) }}">Export EN PDF</a>
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
    <div class="summary-box"><strong>Dutch / Original</strong></div>
    <div class="full-block">{{ block.full_text }}</div>
    {% if block.full_text_en %}
    <div class="summary-box" style="margin-top:16px;"><strong>English (Translation)</strong></div>
    <div class="full-block">{{ block.full_text_en }}</div>
    {% endif %}

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
.toolbar { padding:12px 16px; background:white; border-bottom:1px solid #ddd; display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; }
.viewer-layout { display:flex; height: calc(100vh - 60px); }
.pdf-pane { flex: 1 1 58%; min-width: 320px; border-right:1px solid #ddd; position:relative; }
.translation-pane { flex: 1 1 42%; min-width: 280px; overflow:auto; background:#f8fafc; padding:16px; }
iframe { width:100%; height:100%; border:none; }
.pdf-missing { padding:24px; color:#842029; background:#f8d7da; border:1px solid #f5c2c7; margin:16px; border-radius:8px; line-height:1.6; }
.translation-box { background:white; border:1px solid #e2e8f0; border-left:4px solid #28a745; border-radius:8px; padding:14px; white-space:pre-wrap; line-height:1.55; margin-bottom:14px; }
.translation-box h4 { margin:0 0 8px 0; color:#334155; }
.translation-status { color:#64748b; font-size:13px; margin-bottom:12px; }
.btn-small { padding:6px 10px; font-size:13px; }
.page-nav { display:flex; gap:6px; align-items:center; flex-wrap:wrap; }
</style>
</head>
<body>
    <div class="toolbar">
        <div>
            <strong>{{ filename }}</strong>
            | Page {{ page }} / {{ total_pages }}
        </div>
        <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
            <div class="page-nav">
                {% if page > 1 %}
                <a class="btn btn-gray btn-small" href="{{ url_for('pdf_viewer', document_id=document_id, page=page-1) }}">Prev</a>
                {% endif %}
                {% if page < total_pages %}
                <a class="btn btn-gray btn-small" href="{{ url_for('pdf_viewer', document_id=document_id, page=page+1) }}">Next</a>
                {% endif %}
            </div>
            {% if translation_enabled %}
            <button class="btn btn-green btn-small" id="reload-translation">Refresh all</button>
            <button class="btn btn-purple btn-small" id="export-translation" disabled>Export page (.txt)</button>
            <a class="btn btn-orange btn-small" href="{{ url_for('export_translation_pdf', document_id=document_id) }}">Export full PDF (EN)</a>
            {% endif %}
            <a href="{{ url_for('home') }}">Back</a>
        </div>
    </div>
    <div class="viewer-layout">
        <div class="pdf-pane">
            {% if pdf_available %}
            <iframe src="{{ pdf_url }}#page={{ page }}"></iframe>
            {% else %}
            <div class="pdf-missing">
                <strong>Original PDF not found on server.</strong><br>
                The file may have been lost after a restart before Supabase Storage was configured.<br>
                Please go to <strong>Documents</strong>, re-upload this PDF, then click <strong>Reindex</strong>.
                <br><br>
                English translations on the right may still work from cached index data.
            </div>
            {% endif %}
        </div>
        {% if translation_enabled %}
        <div class="translation-pane">
            <div class="translation-status" id="translation-status">Loading all translations...</div>
            <div id="translated-pages"></div>
        </div>
        {% endif %}
    </div>
    {% if translation_enabled %}
    <script>
    const documentId = {{ document_id }};
    const pageNum = {{ page }};
    const totalPages = {{ total_pages }};
    const fileBase = {{ filename|tojson }};
    const statusEl = document.getElementById("translation-status");
    const pagesEl = document.getElementById("translated-pages");
    const exportBtn = document.getElementById("export-translation");
    let currentTranslation = "";

    function escapeHtml(value) {
        return (value || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }

    function renderPageBlock(page, translation, provider, cached) {
        const box = document.createElement("div");
        box.className = "translation-box";
        box.id = "translation-page-" + page;
        if (page === pageNum) {
            box.style.boxShadow = "0 0 0 2px #6f42c1";
        }
        const tag = cached ? "cached" : "new";
        box.innerHTML =
            "<h4>Page " + page + " / " + totalPages + " (" + provider + ", " + tag + ")</h4>" +
            "<div class='page-text'>" + escapeHtml(translation || "(Translation unavailable)") + "</div>";
        return box;
    }

    async function loadAllTranslations() {
        statusEl.textContent = "Loading translations for all pages...";
        pagesEl.innerHTML = "";
        exportBtn.disabled = true;
        currentTranslation = "";

        try {
            const resp = await fetch(`/api/document/${documentId}/translate-all`);
            const data = await resp.json();
            if (!resp.ok) {
                throw new Error(data.error || "Failed to load translations");
            }

            pagesEl.innerHTML = "";
            for (const item of data.pages) {
                pagesEl.appendChild(
                    renderPageBlock(item.page, item.translation, item.provider || "unknown", item.cached)
                );
                if (item.page === pageNum) {
                    currentTranslation = item.translation || "";
                }
            }

            const loaded = data.loaded_pages || data.pages.length;
            statusEl.textContent = data.truncated
                ? `Loaded ${loaded} of ${data.total_pages} pages (server limit). Use Export full PDF for file download.`
                : `Loaded all ${loaded} pages. Current highlight: page ${pageNum}.`;

            const currentBlock = document.getElementById("translation-page-" + pageNum);
            if (currentBlock) {
                currentBlock.scrollIntoView({ behavior: "smooth", block: "start" });
            }
            exportBtn.disabled = !currentTranslation.trim();
        } catch (err) {
            statusEl.textContent = "Translation error: " + err.message;
        }
    }

    function exportTranslation() {
        if (!currentTranslation.trim()) {
            return;
        }
        const safeName = fileBase.replace(/\\.[^.]+$/, "") || "document";
        const blob = new Blob([currentTranslation], { type: "text/plain;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `${safeName}_page_${pageNum}_en.txt`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
    }

    document.getElementById("reload-translation").addEventListener("click", loadAllTranslations);
    exportBtn.addEventListener("click", exportTranslation);
    loadAllTranslations();
    </script>
    {% endif %}
</body>
</html>
"""

AI_TRANSLATE_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>AI Translation Tool</title>
""" + BASE_CSS + """
</head>
<body>
<div class="container">
    <h2>AI Translation Tool</h2>
    <p class="small">
        Provider: <strong>{{ provider_info.provider }}</strong>
        | Model: <strong>{{ provider_info.model }}</strong>
        | Status: <strong>{{ "Ready" if provider_info.configured else "Not configured" }}</strong>
    </p>

    <form method="post">
        <label><strong>Source text (Dutch / mixed)</strong></label>
        <textarea name="source_text" rows="10" style="width:100%; margin-top:8px;">{{ source_text }}</textarea>
        <div style="margin-top:12px;">
            <button class="btn btn-green" type="submit">Translate to English</button>
            <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
        </div>
    </form>

    {% if translated_text %}
    <div class="summary-box" style="margin-top:18px;">
        <strong>English translation</strong>
        <div class="full-block" style="margin-top:10px;">{{ translated_text }}</div>
    </div>
    {% endif %}
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
                <h3>Table 1 - Dutch (Original)</h3>
                <div class="data-table">{{ t.html_table|safe }}</div>
                {% if t.html_table_en %}
                <h3 style="margin-top:18px;">Table 2 - English (Translation)</h3>
                <div class="data-table">{{ t.html_table_en|safe }}</div>
                {% else %}
                <div class="notice" style="margin-top:12px;">English translation not available yet. Reindex the document to generate it.</div>
                {% endif %}
            </div>
        {% endfor %}
    {% else %}
        <div class="warning">No table preview available.</div>
        {% if error_msg %}
        <div class="notice" style="margin-top:12px;">{{ error_msg }}</div>
        {% else %}
        <div class="notice" style="margin-top:12px;">
            The table was not indexed yet, or the source file was lost after a server restart.
            {% if current_user.is_authenticated and current_user.is_admin %}
            Go to <a href="{{ url_for('admin_documents') }}">Documents</a> and click <strong>Reindex</strong> for this file.
            {% else %}
            Ask an admin to reindex this document.
            {% endif %}
        </div>
        {% endif %}
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
            "storage_backend": storage.storage_backend_name(),
            "data_dir": str(DATA_DIR),
            "database_path": str(DATABASE_PATH),
            "translation": get_translation_status(),
        }
    )


@app.route("/ai-translate", methods=["GET", "POST"])
@login_required
def ai_translate_tool():
    source_text = ""
    translated_text = ""
    provider_info = get_translation_status()

    if request.method == "POST":
        source_text = request.form.get("source_text", "").strip()
        if source_text:
            translated_text = translate_to_english(source_text)

    return render_template_string(
        AI_TRANSLATE_TEMPLATE,
        source_text=source_text,
        translated_text=translated_text,
        provider_info=provider_info,
    )


@app.route("/api/translate", methods=["POST"])
@login_required
def api_translate_text():
    payload = request.get_json(silent=True) or {}
    source_text = (payload.get("text") or "").strip()
    if not source_text:
        return jsonify({"error": "text is required"}), 400
    if not translation.translation_enabled():
        return jsonify({"error": "translation is disabled"}), 503

    translated = translate_to_english(source_text)
    return jsonify(
        {
            "source": source_text,
            "translation": translated,
            "provider": get_translation_provider(),
        }
    )


@app.route("/api/document/<int:document_id>/translate-page")
def api_translate_page(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not translation.translation_enabled():
        return jsonify({"error": "translation is disabled"}), 503

    page = request.args.get("page", 1, type=int)
    result = get_or_translate_page(doc, page)
    if result.get("error") and not result.get("translation"):
        return jsonify(result), 404
    return jsonify(result)


@app.route("/api/document/<int:document_id>/translate-all")
def api_translate_all(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        return jsonify({"error": "Document not found"}), 404
    if not translation.translation_enabled():
        return jsonify({"error": "translation is disabled"}), 503

    total_pages = count_pdf_pages(doc)
    if total_pages <= 0:
        return jsonify({"error": "No pages found for this document."}), 404

    export_limit = MAX_PDF_PAGES if MAX_PDF_PAGES > 0 else total_pages
    pages_to_load = min(total_pages, export_limit)
    results = []
    for page_num in range(1, pages_to_load + 1):
        result = get_or_translate_page(doc, page_num)
        results.append(
            {
                "page": page_num,
                "translation": result.get("translation") or "",
                "provider": result.get("provider") or get_translation_provider(),
                "cached": bool(result.get("cached")),
                "error": result.get("error", ""),
            }
        )

    return jsonify(
        {
            "total_pages": total_pages,
            "loaded_pages": pages_to_load,
            "truncated": pages_to_load < total_pages,
            "pages": results,
        }
    )


@app.route("/document/<int:document_id>/export-translation-pdf")
@login_required
def export_translation_pdf(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)

    if not translation.translation_enabled():
        flash("Translation is disabled.")
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        flash("Source PDF not found. Re-upload the file and try again.")
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    try:
        pdf_bytes = build_translated_pdf_bytes(doc)
    except Exception as exc:
        print(f"Translated PDF export error for document {document_id}: {exc}")
        flash(f"Export failed: {exc}")
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    base_name = secure_filename(
        Path(doc.original_filename).stem or f"document_{document_id}"
    )
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{base_name}_english.pdf",
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

            stored_filename = f"{uuid.uuid4().hex}.{ext}"
            size_bytes = storage.save_document(stored_filename, temp_path, DOC_FOLDER)

            doc = DocumentRecord(
                original_filename=original_filename,
                stored_filename=stored_filename,
                extension=ext,
                file_hash=file_hash,
                size_bytes=size_bytes,
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
        block_count = RequirementBlock.query.filter_by(document_id=doc.id).count()
        flash(f"Reindexed: {doc.original_filename} ({block_count} requirement blocks)")
    except Exception as e:
        db.session.rollback()
        print(f"Reindex error for document {document_id}: {e}")
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

    storage.delete_document(doc.stored_filename, DOC_FOLDER)

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
    block_view = {
        "requirement_id": block.requirement_id,
        "title": block.title,
        "section": block.section,
        "page": block.page,
        "category": block.category,
        "definition": block.definition,
        "summary": block.summary,
        "full_text": block.full_text,
        "full_text_en": translate_to_english(block.full_text),
        "document": block.document,
    }
    return render_template_string(
        REQUIREMENT_TEMPLATE,
        block=block_view,
        pdf2image_available=PDF2IMAGE_AVAILABLE
    )

@app.route("/table/<int:document_id>")
def table_preview(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        abort(404)

    tables = TablePreview.query.filter_by(document_id=document_id).all()
    error_msg = ""
    if not tables and doc.extension.lower() in {"csv", "xlsx"}:
        try:
            tables = regenerate_table_previews(doc)
            if not tables:
                error_msg = "The spreadsheet has no readable sheets or table data."
        except FileNotFoundError:
            error_msg = (
                "Source file not found. It may have been lost after a restart. "
                "Configure Supabase Storage and upload the file again."
            )
        except Exception as exc:
            print(f"Table preview regeneration error for document {document_id}: {exc}")
            error_msg = f"Could not build table preview: {exc}"

    return render_template_string(
        TABLE_TEMPLATE, doc=doc, tables=tables, error_msg=error_msg
    )

@app.route("/pdf/<int:document_id>")
def pdf_viewer(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)

    page = request.args.get("page", 1, type=int)
    pdf_available = storage.document_exists(doc.stored_filename, DOC_FOLDER)
    total_pages = count_pdf_pages(doc)
    if total_pages <= 0:
        total_pages = 1
    page = max(1, min(page, total_pages))
    pdf_url = url_for("serve_pdf", document_id=document_id) if pdf_available else ""
    return render_template_string(
        PDF_TEMPLATE,
        document_id=document_id,
        filename=doc.original_filename,
        page=page,
        total_pages=total_pages,
        pdf_url=pdf_url,
        pdf_available=pdf_available,
        translation_enabled=translation.translation_enabled(),
    )

@app.route("/serve-pdf/<int:document_id>")
def serve_pdf(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)

    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        abort(404)

    pdf_bytes = storage.read_document_bytes(doc.stored_filename, DOC_FOLDER)
    return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf")

@app.route("/page-image/<int:document_id>")
def page_image(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)
    page = request.args.get("page", 1, type=int)
    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        abort(404)
    with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
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
            storage.delete_document(doc.stored_filename, DOC_FOLDER)

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



def ensure_schema():
    from sqlalchemy import inspect, text

    inspector = inspect(db.engine)
    doc_cols = {col["name"] for col in inspector.get_columns("documents")}
    if "page_offsets_json" not in doc_cols:
        db.session.execute(text("ALTER TABLE documents ADD COLUMN page_offsets_json TEXT"))
        db.session.commit()

    block_cols = {col["name"] for col in inspector.get_columns("requirement_blocks")}
    if "char_start" not in block_cols:
        db.session.execute(text("ALTER TABLE requirement_blocks ADD COLUMN char_start INTEGER DEFAULT 0"))
        db.session.commit()

    table_cols = {col["name"] for col in inspector.get_columns("table_previews")}
    if "html_table_en" not in table_cols:
        db.session.execute(text("ALTER TABLE table_previews ADD COLUMN html_table_en TEXT"))
        db.session.commit()

    cache_cols = set()
    if "translation_cache" in inspector.get_table_names():
        cache_cols = {col["name"] for col in inspector.get_columns("translation_cache")}
    if cache_cols and "provider" not in cache_cols:
        db.session.execute(text("ALTER TABLE translation_cache ADD COLUMN provider VARCHAR(32)"))
        db.session.commit()

    if "document_page_translations" not in inspector.get_table_names():
        db.create_all()


# =========================================================
# App initialization
# =========================================================
with app.app_context():
    db.create_all()
    ensure_schema()
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
