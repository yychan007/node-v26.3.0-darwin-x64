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
import tempfile
import mimetypes
from datetime import datetime
from collections import defaultdict, Counter
from pathlib import Path
from typing import Any, cast

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

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
    if not TESSERACT_AVAILABLE or pytesseract is None:
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
DocxDocument = None
try:
    from docx import Document as DocxDocument
except Exception:
    DOCX_AVAILABLE = False

OCR_AVAILABLE = True
fitz = None
try:
    import fitz  # PyMuPDF
except Exception:
    OCR_AVAILABLE = False

TESSERACT_AVAILABLE = True
PDF2IMAGE_AVAILABLE = False
pytesseract = None
try:
    import pytesseract
except Exception:
    TESSERACT_AVAILABLE = False

convert_from_path = None
try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False

SEMANTIC_AVAILABLE = True
SentenceTransformer = None
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
TRANSLATE_MAX_CELLS = env_int("TRANSLATE_MAX_CELLS", 500)


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
# Avoid static type issues with Flask-Login's stubs.
setattr(login_manager, "login_view", "login")
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
    if semantic_model is None and SEMANTIC_AVAILABLE and SentenceTransformer is not None:
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
    major_section = db.Column(db.String(300), default="")
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
    html_table_es = db.Column(db.Text, default="")
    csv_text = db.Column(db.Text, default="")
    csv_text_en = db.Column(db.Text, default="")
    csv_text_es = db.Column(db.Text, default="")
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
    translated_text_es = db.Column(db.Text, default="")
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


def get_all_expanded_terms(query):
    q = (query or "").strip().lower()
    terms = []
    if q:
        terms.append(q)
    for token in preprocess(query):
        terms.extend(expand_query_tokens([token]))
    return list(dict.fromkeys(terms))


def default_exact_search_terms(query):
    tokens = preprocess(query)
    if tokens:
        return tokens
    q = (query or "").strip().lower()
    return [q] if q else []


def resolve_active_search_terms(query, selected_terms=None):
    all_terms = get_all_expanded_terms(query)
    if not selected_terms:
        return default_exact_search_terms(query)

    allowed = {term.lower() for term in all_terms}
    active = []
    for raw in selected_terms:
        value = (raw or "").strip().lower()
        if value not in allowed:
            continue
        for candidate in all_terms:
            if candidate.lower() == value:
                active.append(candidate)
                break
    return list(dict.fromkeys(active)) or default_exact_search_terms(query)


def term_matches_text(term, text):
    value = (term or "").strip().lower()
    haystack = (text or "").lower()
    if not value or not haystack:
        return False
    if " " in value:
        return value in haystack
    return re.search(r"\b" + re.escape(value) + r"\b", haystack) is not None


def count_exact_term_hits(block, query, exact_terms):
    combined = " ".join(
        [
            block.title or "",
            block.summary or "",
            block.definition or "",
            block.full_text or "",
            block.requirement_id or "",
        ]
    )
    hits = 0
    q = (query or "").strip().lower()
    if q and q in combined.lower():
        hits += 3
    for term in exact_terms:
        if term_matches_text(term, combined):
            hits += 1
    return hits

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

def strip_html_legacy(text):
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
    cache_cls = cast(Any, TranslationCache)
    db.session.add(
        cache_cls(
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


def get_document_full_text(doc):
    text = (doc.text_preview or "").strip()
    if text:
        return text

    ext = doc.extension.lower()
    if ext not in {"pdf", "docx", "txt"}:
        return ""
    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        return ""

    try:
        with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
            if ext == "docx":
                text, _ = extract_text_from_docx(file_path)
            elif ext == "txt":
                text, _ = extract_text_from_txt(file_path)
            elif ext == "pdf" and OCR_AVAILABLE and fitz is not None:
                with fitz.open(file_path) as pdf_doc:
                    # Convert to str explicitly to keep the type checker happy.
                    text_parts = []
                    for i in range(len(pdf_doc)):
                        page_text = pdf_doc.load_page(i).get_text("text") or ""
                        text_parts.append(str(page_text))
                    text = "\n".join(text_parts)
            else:
                text = ""
    except Exception as exc:
        print(f"Full text extract error for doc {doc.id}: {exc}")
        text = ""
    return (text or "").strip()


def split_text_into_virtual_pages(text, max_chars=3500):
    text = (text or "").strip()
    if not text:
        return [""]
    if len(text) <= max_chars:
        return [text]

    parts = []
    start = 0
    while start < len(text):
        end = min(len(text), start + max_chars)
        if end < len(text):
            split_at = text.rfind("\n\n", start, end)
            if split_at <= start + 400:
                split_at = text.rfind("\n", start, end)
            if split_at > start + 400:
                end = split_at + 1
        parts.append(text[start:end].strip())
        start = end
    return [part for part in parts if part] or [""]


def extract_page_text_from_preview(doc, page_num):
    if page_num < 1:
        return ""

    ext = doc.extension.lower()
    if ext in {"docx", "txt"}:
        chunks = split_text_into_virtual_pages(get_document_full_text(doc))
        if page_num > len(chunks):
            return ""
        return chunks[page_num - 1]

    offsets = load_page_offsets(doc)
    text = doc.text_preview or ""
    if offsets and text:
        for start, end, page in offsets:
            if page == page_num:
                return text[start:end].strip()
    return ""


def pdf_source_file_missing(doc):
    return (
        doc.extension.lower() == "pdf"
        and not storage.document_exists(doc.stored_filename, DOC_FOLDER)
    )


def extract_single_page_text(doc, page_num):
    ext = doc.extension.lower()
    preview_text = extract_page_text_from_preview(doc, page_num)
    if preview_text:
        return preview_text

    if ext in {"docx", "txt"}:
        return ""

    if ext != "pdf" or not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        return ""

    if not OCR_AVAILABLE or fitz is None:
        return ""

    try:
        with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
            with fitz.open(file_path) as pdf_doc:
                if page_num < 1 or page_num > len(pdf_doc):
                    return ""
                page_text = pdf_doc.load_page(page_num - 1).get_text("text") or ""
                return str(page_text).strip()
    except Exception as exc:
        print(f"Page text extract error for doc {doc.id} page {page_num}: {exc}")
        return ""


def normalize_translation_lang(lang):
    return "en"


def get_or_translate_page(doc, page_num, target_lang="en"):
    target_lang = normalize_translation_lang(target_lang)
    row = DocumentPageTranslation.query.filter_by(
        document_id=doc.id, page=page_num
    ).first()

    cached_translation = ""
    if row:
        cached_translation = (
            row.translated_text_es if target_lang == "es" else row.translated_text
        ) or ""
    if cached_translation:
        cached_source = row.source_text if row else ""
        cached_provider = row.provider if row else get_translation_provider()
        return {
            "page": page_num,
            "lang": target_lang,
            "source": cached_source,
            "translation": cached_translation,
            "provider": cached_provider,
            "cached": True,
        }

    source = extract_single_page_text(doc, page_num)
    if not source:
        if pdf_source_file_missing(doc):
            error = (
                "Source PDF not found on server. "
                "Re-upload the file in Documents, then click Reindex."
            )
        else:
            error = "No extractable text on this page."
        return {
            "page": page_num,
            "lang": target_lang,
            "source": "",
            "translation": "",
            "provider": get_translation_provider(),
            "cached": False,
            "error": error,
        }

    translated = translate_to_language(source, target_lang)
    provider = get_translation_provider()
    if translated:
        if row:
            row.source_text = source[:20000]
            if target_lang == "es":
                row.translated_text_es = translated
            else:
                row.translated_text = translated
            row.provider = provider
        else:
            dpt_cls = cast(Any, DocumentPageTranslation)
            db.session.add(
                dpt_cls(
                    document_id=doc.id,
                    page=page_num,
                    source_text=source[:20000],
                    translated_text=translated if target_lang == "en" else "",
                    translated_text_es=translated if target_lang == "es" else "",
                    provider=provider,
                )
            )
        try:
            db.session.commit()
        except Exception:
            db.session.rollback()

    return {
        "page": page_num,
        "lang": target_lang,
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

    if not OCR_AVAILABLE or fitz is None:
        return 0

    try:
        with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
            with fitz.open(file_path) as pdf_doc:
                return len(pdf_doc)
    except Exception as exc:
        print(f"PDF page count error for doc {doc.id}: {exc}")
        return 0


def count_document_pages(doc):
    ext = doc.extension.lower()
    if ext in {"docx", "txt"}:
        return max(1, len(split_text_into_virtual_pages(get_document_full_text(doc))))
    if ext == "pdf":
        total = count_pdf_pages(doc)
        return total if total > 0 else 0
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
    if fitz is None:
        raise RuntimeError("PyMuPDF is required to export translated PDFs.")
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


def build_translated_pdf_bytes(doc, target_lang="en"):
    if not OCR_AVAILABLE:
        raise RuntimeError("PyMuPDF is required to export translated PDFs.")
    if fitz is None:
        raise RuntimeError("PyMuPDF is required to export translated PDFs.")

    if not translation.translation_enabled():
        raise RuntimeError("Translation is disabled.")

    target_lang = normalize_translation_lang(target_lang)
    lang_label = translation.language_label(target_lang)

    total_pages = count_document_pages(doc)
    if total_pages <= 0:
        raise ValueError("Could not determine document page count.")

    export_limit = MAX_PDF_PAGES if MAX_PDF_PAGES > 0 else total_pages
    pages_to_export = min(total_pages, export_limit)
    truncated = pages_to_export < total_pages

    pdf_out = fitz.open()
    source_name = doc.original_filename or doc.stored_filename

    try:
        if truncated:
            notice = (
                f"{lang_label} translation export for {source_name}\n"
                f"Pages 1-{pages_to_export} of {total_pages}\n\n"
            )
            _append_translation_text(pdf_out, "Export notice", notice)

        for page_num in range(1, pages_to_export + 1):
            result = get_or_translate_page(doc, page_num, target_lang=target_lang)
            translation_text = (result.get("translation") or "").strip()
            if not translation_text:
                translation_text = "(No translation available for this page.)"

            page_word = "Page" if doc.extension == "pdf" else "Section"
            header = f"{source_name} - {page_word} {page_num} ({lang_label})"
            _append_translation_text(pdf_out, header, translation_text)

        if pdf_out.page_count == 0:
            raise ValueError("No translated pages were generated.")

        buffer = io.BytesIO()
        pdf_out.save(buffer, garbage=4, deflate=True)
        buffer.seek(0)
        return buffer.getvalue()
    finally:
        pdf_out.close()


def translate_to_language(text, target_lang="en"):
    return translation.translate_text(
        text,
        target_lang=normalize_translation_lang(target_lang),
        cache_get=get_cached_translation,
        cache_set=store_cached_translation,
    )


def translate_to_english(text):
    return translate_to_language(text, "en")


def build_translated_snippet(snippet_html, target_lang="en"):
    plain = strip_html(snippet_html)
    translated = translate_to_language(plain, target_lang)
    if not translated or translated == plain:
        return ""
    return highlight_terms(translated, [])


def build_english_snippet(snippet_html):
    return build_translated_snippet(snippet_html, "en")


def enrich_result_with_translation(result_dict):
    result_dict["snippet_en"] = build_translated_snippet(result_dict.get("snippet", ""), "en")
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

    exact_terms = default_exact_search_terms(query)
    for term in exact_terms:
        match = re.search(r"\b" + re.escape(term.lower()) + r"\b", text_lower)
        if match:
            return match.start()

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
    # First try pdf2image (requires poppler).
    if PDF2IMAGE_AVAILABLE and convert_from_path is not None:
        try:
            images = convert_from_path(
                document_path,
                first_page=page_num,
                last_page=page_num,
                dpi=150,
            )
            if images:
                img_io = io.BytesIO()
                images[0].save(img_io, format="PNG")
                img_io.seek(0)
                return img_io.getvalue(), "image/png"
        except Exception as e:
            print(f"pdf2image page render error: {document_path} page {page_num} -> {e}")

    # Fallback: render with PyMuPDF, which doesn't need poppler.
    if OCR_AVAILABLE and fitz is not None:
        try:
            with fitz.open(document_path) as pdf_doc:
                if page_num < 1 or page_num > len(pdf_doc):
                    return None
                page = pdf_doc.load_page(page_num - 1)
                pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
                return pix.tobytes("png"), "image/png"
        except Exception as e:
            print(f"fitz page render error: {document_path} page {page_num} -> {e}")

    return None


def list_docx_image_entries(document_path):
    try:
        with zipfile.ZipFile(document_path) as docx_zip:
            image_entries = [
                name
                for name in docx_zip.namelist()
                if name.startswith("word/media/")
                and file_ext(name) in {"png", "jpg", "jpeg", "gif", "bmp", "webp"}
            ]
        return sorted(image_entries)
    except Exception as exc:
        print(f"DOCX image list error: {document_path} -> {exc}")
        return []


def get_docx_image(document_path, image_index=0):
    image_entries = list_docx_image_entries(document_path)
    if image_index < 0 or image_index >= len(image_entries):
        return None

    entry_name = image_entries[image_index]
    try:
        with zipfile.ZipFile(document_path) as docx_zip:
            image_bytes = docx_zip.read(entry_name)
        mime_type = mimetypes.guess_type(entry_name)[0] or "application/octet-stream"
        return image_bytes, mime_type
    except Exception as exc:
        print(f"DOCX image read error: {document_path} {entry_name} -> {exc}")
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
    if not OCR_AVAILABLE or fitz is None:
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
                text = str(page.get_text("text") or "")
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
    if not DOCX_AVAILABLE or DocxDocument is None:
        return "", [(0, 0, 1)]
    doc = DocxDocument(path)
    parts = []
    for paragraph in doc.paragraphs:
        value = (paragraph.text or "").strip()
        if value:
            parts.append(value)
    for table in doc.tables:
        for row in table.rows:
            cells = [(cell.text or "").strip() for cell in row.cells]
            row_text = " | ".join(cell for cell in cells if cell)
            if row_text:
                parts.append(row_text)
    text = "\n".join(parts)
    return text, [(0, len(text), 1)]

def extract_text_from_csv(path):
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
    except UnicodeDecodeError:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="latin1")
    df = clean_dataframe_for_display(df, max_rows=200)
    text = df.to_string(index=False)
    return text, [(0, len(text), 1)], [("CSV", df)]
    
def clean_dataframe_for_display(df, max_rows=50):
    if df is None:
        return pd.DataFrame()

    cleaned = df.copy()
    cleaned = cleaned.fillna("")
    cleaned = cleaned.astype(str)
    for bad in ("nan", "NaN", "None", "<NA>"):
        cleaned = cleaned.replace(bad, "")

    cleaned.columns = [
        (str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {idx + 1}")
        for idx, col in enumerate(cleaned.columns)
    ]

    if not cleaned.empty:
        row_keep = cleaned.apply(
            lambda row: any(str(value).strip() for value in row), axis=1
        )
        cleaned = cleaned.loc[row_keep]
        col_keep = cleaned.apply(
            lambda col: any(str(value).strip() for value in col), axis=0
        )
        cleaned = cleaned.loc[:, col_keep]

    return cleaned.head(max_rows).reset_index(drop=True)


def extract_text_from_xlsx(path):
    workbook = pd.ExcelFile(path)
    all_parts = []
    tables = []
    for sheet_name in workbook.sheet_names:
        raw_df = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
        df = normalize_excel_sheet(raw_df)
        if df.empty:
            continue
        tables.append((sheet_name, df))
        all_parts.append(f"[Sheet: {sheet_name}]")
        all_parts.append(df.to_string(index=False))
    text = "\n".join(all_parts)
    return text, [(0, len(text), 1)], tables


def normalize_excel_sheet(raw_df, max_rows=200):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame()

    sheet = raw_df.fillna("").astype(str)
    for bad in ("nan", "NaN", "None", "<NA>"):
        sheet = sheet.replace(bad, "")

    header_row_idx = None
    for idx in range(min(len(sheet), 25)):
        row_values = [str(value).strip() for value in sheet.iloc[idx].tolist()]
        non_empty = [value for value in row_values if value]
        if len(non_empty) >= 2:
            header_row_idx = idx
            break

    if header_row_idx is None:
        return clean_dataframe_for_display(sheet, max_rows=max_rows)

    headers = []
    for col_idx, value in enumerate(sheet.iloc[header_row_idx].tolist()):
        label = str(value).strip()
        headers.append(label or f"Column {col_idx + 1}")

    body = sheet.iloc[header_row_idx + 1 :].copy()
    body = body.iloc[:, : len(headers)]
    body.columns = headers
    body = clean_dataframe_for_display(body, max_rows=max_rows)
    if body.empty:
        return body

    title_lines = []
    for idx in range(header_row_idx):
        row_text = " ".join(
            str(value).strip()
            for value in sheet.iloc[idx].tolist()
            if str(value).strip()
        )
        if row_text:
            title_lines.append(row_text)

    if title_lines:
        title_row = {col: "" for col in body.columns}
        title_row[body.columns[0]] = " | ".join(title_lines)
        body = pd.concat([pd.DataFrame([title_row]), body], ignore_index=True)

    return body.head(max_rows)

def ocr_pdf(path):
    if not TESSERACT_AVAILABLE or pytesseract is None or convert_from_path is None:
        return "", [(0, 0, 1)]

    full_text = ""
    page_offsets = []
    current = 0

    try:
        images = convert_from_path(path)
        for i, image in enumerate(images, start=1):
            text = pytesseract.image_to_string(image) or ""
            text = str(text)
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


SECTION_HEADER_REGEX = re.compile(
    r"(?m)^(\d+(?:\.\d+)*)\s+(.+)$"
)


def clean_section_title(title):
    value = (title or "").strip()
    value = re.sub(r"\.{2,}.*$", "", value).strip()
    value = re.sub(r"\s+\d+$", "", value).strip()
    return value


def parse_document_sections(content):
    sections = []
    for match in SECTION_HEADER_REGEX.finditer(content or ""):
        num = match.group(1).strip()
        title = clean_section_title(match.group(2))
        if not title or len(title) < 3 or len(title) > 200:
            continue
        if title.lower().startswith(("page ", "pagina ")):
            continue
        sections.append(
            {
                "char_start": match.start(),
                "num": num,
                "title": title,
                "label": f"{num} {title}",
            }
        )
    return sections


def get_major_section_for_position(sections, char_pos):
    major = ""
    for sec in sections:
        if sec["char_start"] > char_pos:
            break
        if re.fullmatch(r"\d+", sec["num"]):
            major = sec["label"]
    return major


def get_nearest_section_for_position(sections, char_pos):
    nearest = ""
    for sec in sections:
        if sec["char_start"] > char_pos:
            break
        nearest = sec["label"]
    return nearest


def parse_section_blocks(content, page_offsets):
    matches = list(SECTION_HEADER_REGEX.finditer(content))
    if not matches:
        return []

    doc_sections = parse_document_sections(content)
    blocks = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        block_text = content[start:end].strip()
        if len(block_text) < 40:
            continue

        section_num = match.group(1)
        section_title = clean_section_title(match.group(2))
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
            "major_section": get_major_section_for_position(doc_sections, start),
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
    doc_sections = parse_document_sections(content)
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
            section = get_nearest_section_for_position(doc_sections, start)
            major_section = get_major_section_for_position(doc_sections, start)
            definition = ""

            if len(lines) >= 2:
                if req_id.lower() in lines[0].lower():
                    title = lines[1]
                else:
                    title = lines[0]

            for idx, line in enumerate(lines):
                low = line.lower()
                if not section and re.match(r"^\d+(?:\.\d+)*\s+", line):
                    section = clean_section_title(line)
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
                "major_section": major_section,
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
    doc_sections = parse_document_sections(content)
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
            "section": get_nearest_section_for_position(doc_sections, pos),
            "major_section": get_major_section_for_position(doc_sections, pos),
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
        user_cls = cast(Any, User)
        user = user_cls(username=admin_username, role=ADMIN_ROLE)
        user.set_password(admin_password)
        db.session.add(user)
        db.session.commit()
        print(f"Created initial admin user '{admin_username}'.")

def build_table_html(df):
    preview_df = clean_dataframe_for_display(df, max_rows=40)
    if preview_df.empty:
        return "", pd.DataFrame()
    html_table = preview_df.to_html(
        index=False, classes="data-table", border=0, na_rep=""
    )
    return html_table, preview_df


def translate_preview_df(preview_df, target_lang="en"):
    if preview_df is None or preview_df.empty or not translation.translation_enabled():
        return "", ""
    try:
        translated_df = translation.translate_dataframe_values(
            preview_df,
            target_lang=target_lang,
            cache_get=get_cached_translation,
            cache_set=store_cached_translation,
            max_cells=TRANSLATE_MAX_CELLS,
        )
        translated_df = clean_dataframe_for_display(translated_df, max_rows=40)
        csv_buf = io.StringIO()
        translated_df.to_csv(csv_buf, index=False)
        html_table = translated_df.to_html(
            index=False, classes="data-table", border=0, na_rep=""
        )
        return html_table, csv_buf.getvalue()
    except Exception as exc:
        print(f"Table translation skipped ({target_lang}): {exc}")
        return "", ""


def translate_preview_df_to_html(preview_df, target_lang="en"):
    html_table, _ = translate_preview_df(preview_df, target_lang)
    return html_table


def save_table_previews(doc_id, tables):
    for sheet_name, df in tables:
        if df is None:
            continue
        html_table, preview_df = build_table_html(df)
        if preview_df.empty:
            continue
        html_table_en = ""
        csv_text_en = ""
        if translation.translation_enabled():
            html_table_en, csv_text_en = translate_preview_df(preview_df, "en")
        csv_buf = io.StringIO()
        preview_df.to_csv(csv_buf, index=False)
        table_preview_cls = cast(Any, TablePreview)
        row = table_preview_cls(
            document_id=doc_id,
            sheet_name=str(sheet_name),
            page=1,
            table_format="xlsx" if sheet_name != "CSV" else "csv",
            html_table=html_table,
            html_table_en=html_table_en,
            html_table_es="",
            csv_text=csv_buf.getvalue(),
            csv_text_en=csv_text_en,
            csv_text_es="",
            preview_title=f"{sheet_name} preview"
        )
        db.session.add(row)


def table_preview_needs_regenerate(tables):
    if not tables:
        return True
    for row in tables:
        html = row.html_table or ""
        if "NaN" in html or "Unnamed:" in html:
            return True
    return False


def table_preview_needs_translation(tables):
    if not translation.translation_enabled():
        return False
    for row in tables:
        if not (row.html_table_en or "").strip():
            return True
    return False


def refresh_table_translations_from_cache(tables):
    updated = False
    for row in tables:
        needs_en = not (row.html_table_en or "").strip()
        if (
            not needs_en
            and "NaN" not in (row.html_table or "")
        ):
            continue
        if not (row.csv_text or "").strip():
            continue
        try:
            df = pd.read_csv(
                io.StringIO(row.csv_text), dtype=str, keep_default_na=False
            )
            html_table, preview_df = build_table_html(df)
            if preview_df.empty:
                continue
            row.html_table = html_table
            if needs_en:
                row.html_table_en, row.csv_text_en = translate_preview_df(preview_df, "en")
            updated = True
        except Exception as exc:
            print(f"Cached table translation failed for preview {row.id}: {exc}")
    if updated:
        db.session.commit()
    return updated


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
    try:
        _index_document_record_impl(doc_record)
    except Exception as exc:
        db.session.rollback()
        raise


def _index_document_record_impl(doc_record):
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

        req_block_cls = cast(Any, RequirementBlock)
        row = req_block_cls(
            document_id=doc_record.id,
            requirement_id=b["requirement_id"],
            title=b["title"],
            section=b["section"],
            major_section=b.get("major_section", ""),
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


def restore_document_file_from_temp(doc, temp_path):
    size_bytes = storage.save_document(doc.stored_filename, temp_path, DOC_FOLDER)
    storage.verify_document_stored(doc.stored_filename, DOC_FOLDER)
    doc.size_bytes = size_bytes
    db.session.commit()
    return size_bytes


# =========================================================
# Search
# =========================================================
def lexical_score(block, query, active_terms, exact_terms=None):
    score = 0.0
    q = query.lower().strip()
    exact_terms = exact_terms or default_exact_search_terms(query)
    synonym_terms = [t for t in active_terms if t.lower() not in {e.lower() for e in exact_terms}]

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
    major_section = (getattr(block, "major_section", "") or "").lower()
    category = (block.category or "").lower()
    combined = " ".join([title, summary, definition, full_text, req_id, section, major_section, category])

    matched_active = any(term_matches_text(term, combined) for term in active_terms)
    if not matched_active and not any(
        term_matches_text(term, " ".join([filename, stored_filename, filename_no_ext]))
        for term in active_terms
    ):
        return 0.0

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

    # Exact query phrase / token matches (highest priority in content)
    if q and q in title:
        score += 45
    if q and q in summary:
        score += 35
    if q and q in definition:
        score += 35
    if q and q in full_text:
        score += 40
    if q and len(q.split()) > 1 and q in full_text:
        score += 15

    for term in exact_terms:
        if term_matches_text(term, title):
            score += 28
        if term_matches_text(term, summary):
            score += 22
        if term_matches_text(term, definition):
            score += 22
        if term_matches_text(term, full_text):
            score += 30
        if term_matches_text(term, section):
            score += 18
        if term_matches_text(term, major_section):
            score += 24
        if term_matches_text(term, req_id):
            score += 25
        if term_matches_text(term, category):
            score += 10

    # Existing ID / metadata matches for exact query
    if q and q == req_id:
        score += 30
    if q and q in section:
        score += 10
    if q and q in major_section:
        score += 14
    if q and q in category:
        score += 6

    query_tokens = preprocess(query)
    if len(query_tokens) > 1 and all(t in full_text for t in query_tokens):
        score += 10
    if len(query_tokens) > 1 and all(t in title for t in query_tokens):
        score += 8

    # Selected synonym / expanded terms (lower weight)
    for term in set(synonym_terms):
        if term_matches_text(term, filename) or term_matches_text(term, filename_no_ext):
            score += 8
        if term_matches_text(term, title):
            score += 5
        if term_matches_text(term, summary):
            score += 4
        if term_matches_text(term, definition):
            score += 4
        if term_matches_text(term, full_text):
            score += 4
        if term_matches_text(term, category):
            score += 2

    if any(t.startswith("aard") for t in active_terms) and "ground" in category:
        score += 2

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

def get_result_snippet(block, active_terms, query=""):
    text = block.full_text or ""
    hit_pos = find_hit_position(text, active_terms, query)

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

    highlight = list(dict.fromkeys([query] + active_terms))
    return highlight_terms(snippet, highlight)


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

    return highlight_terms(snippet, list(dict.fromkeys([query] + expanded_tokens)))


def score_document_text(doc, query, active_terms, exact_terms=None):
    q = query.lower().strip()
    exact_terms = exact_terms or default_exact_search_terms(query)
    synonym_terms = [t for t in active_terms if t.lower() not in {e.lower() for e in exact_terms}]
    text = doc.text_preview or ""
    if not text.strip():
        return 0

    text_lower = text.lower()
    filename = (doc.original_filename or "").lower()
    base_name = filename.rsplit(".", 1)[0] if "." in filename else filename
    score = 0.0

    if not any(
        term_matches_text(term, text)
        or term_matches_text(term, filename)
        or term_matches_text(term, base_name)
        for term in active_terms
    ):
        return 0.0

    if q and q in text_lower:
        score += 35
    if q and q in filename:
        score += 25
    if q and q in base_name:
        score += 20

    for term in exact_terms:
        if term_matches_text(term, text_lower):
            score += 20
        if term_matches_text(term, filename) or term_matches_text(term, base_name):
            score += 12

    matched_terms = 0
    for term in set(synonym_terms):
        if term_matches_text(term, text_lower):
            matched_terms += 1
            score += 4
        if term_matches_text(term, filename) or term_matches_text(term, base_name):
            score += 3

    query_tokens = preprocess(query)
    if len(query_tokens) > 1 and all(t in text_lower for t in query_tokens):
        score += 14

    return score


def count_exact_term_hits_simple(result_item, query, exact_terms):
    combined = " ".join(
        [
            result_item.get("title") or "",
            result_item.get("summary") or "",
            result_item.get("full_text") or "",
            result_item.get("filename") or "",
        ]
    )
    hits = 0
    q = (query or "").strip().lower()
    if q and q in combined.lower():
        hits += 3
    for term in exact_terms:
        if term_matches_text(term, combined):
            hits += 1
    return hits


def build_document_fallback_results(query, active_terms, seen_keys, top_k=10):
    exact_terms = default_exact_search_terms(query)
    results = []
    for doc in DocumentRecord.query.all():
        score = score_document_text(doc, query, active_terms, exact_terms)
        if score <= 0:
            continue

        page = resolve_result_page(
            doc,
            0,
            doc.text_preview,
            active_terms,
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
            "snippet": get_text_snippet(doc.text_preview, active_terms, query),
            "relevance": round(score, 2),
            "is_pdf": doc.extension.lower() == "pdf",
            "is_docx": doc.extension.lower() == "docx",
            "is_txt": doc.extension.lower() == "txt",
            "has_table": doc.extension.lower() in {"csv", "xlsx"},
            "ocr_used": doc.is_ocr,
            "is_image": doc.extension.lower() in {"png", "jpg", "jpeg"},
            "is_document_fallback": True,
        }))

    results.sort(
        key=lambda item: (
            count_exact_term_hits_simple(item, query, exact_terms),
            item["relevance"],
        ),
        reverse=True,
    )
    return results[:top_k]


def grouped_search_summary(results):
    category_counts = Counter(r["category"] for r in results if r.get("category"))
    top_categories = category_counts.most_common(5)
    req_ids = [r["requirement_id"] for r in results if r.get("requirement_id")]
    return {
        "top_categories": top_categories,
        "top_requirement_ids": req_ids[:8]
    }

def search_requirements(query, top_k=30, active_terms=None):
    if not query.strip():
        return [], []

    all_expanded_terms = get_all_expanded_terms(query)
    exact_terms = default_exact_search_terms(query)
    if active_terms is None:
        active_terms = exact_terms
    else:
        active_terms = resolve_active_search_terms(query, active_terms)

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
        score = lexical_score(row, query, active_terms, exact_terms)

        if row.document_id in matched_doc_ids:
            score += 120

        score += semantic_score(query, row)

        if score > 0:
            exact_hits = count_exact_term_hits(row, query, exact_terms)
            scored.append((row, score, exact_hits))

    scored.sort(key=lambda x: (x[2], x[1]), reverse=True)
    results = []
    seen_keys = set()

    for row, score, _exact_hits in scored[:top_k]:
        doc = row.document
        seen_keys.add((doc.id, row.page or 1, row.requirement_id or row.title))
        display_page = resolve_result_page(
            doc,
            row.char_start or 0,
            row.full_text,
            active_terms,
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
            "major_section": getattr(row, "major_section", "") or "",
            "category": row.category,
            "definition": row.definition,
            "summary": row.summary,
            "full_text": row.full_text,
            "snippet": get_result_snippet(row, active_terms, query),
            "relevance": round(score, 2),
            "is_pdf": doc.extension.lower() == "pdf",
            "is_docx": doc.extension.lower() == "docx",
            "is_txt": doc.extension.lower() == "txt",
            "has_table": doc.extension.lower() in {"csv", "xlsx"},
            "ocr_used": doc.is_ocr,
            "is_image": doc.extension.lower() in {"png", "jpg", "jpeg"},
            "is_document_fallback": False,
        }))

    if len(results) < top_k:
        fallback = build_document_fallback_results(
            query, active_terms, seen_keys, top_k=top_k - len(results)
        )
        results.extend(fallback)

    results.sort(
        key=lambda r: (
            count_exact_term_hits_simple(r, query, exact_terms),
            r["relevance"],
        ),
        reverse=True,
    )
    return results[:top_k], all_expanded_terms


def search_documents(query, top_k=10, active_terms=None):
    q = query.lower().strip()
    exact_terms = default_exact_search_terms(query)
    if active_terms is None:
        active_terms = exact_terms
    else:
        active_terms = resolve_active_search_terms(query, active_terms)
    scored = []

    for doc in DocumentRecord.query.all():
        score = score_document_text(doc, query, active_terms, exact_terms)
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

    results = []
    for doc, score in scored[:top_k]:
        page = resolve_result_page(
            doc,
            0,
            doc.text_preview or "",
            active_terms,
            query,
            1,
        )
        image_pages = []
        image_indexes = []
        if doc.extension.lower() == "pdf":
            offsets = load_page_offsets(doc)
            total_pages = max((p for _s, _e, p in offsets), default=page or 1)
            center = page or 1
            candidates = [center, center + 1, center - 1]
            image_pages = [
                p for p in candidates if isinstance(p, int) and p >= 1 and p <= total_pages
            ]
            # keep order and remove duplicates
            seen_pages = set()
            image_pages = [p for p in image_pages if not (p in seen_pages or seen_pages.add(p))]
        if doc.extension.lower() == "docx" and storage.document_exists(doc.stored_filename, DOC_FOLDER):
            try:
                with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
                    image_indexes = list(range(min(3, len(list_docx_image_entries(file_path)))))
            except Exception as exc:
                print(f"DOCX image preview error for doc {doc.id}: {exc}")
        results.append(
            {
                "document_id": doc.id,
                "filename": doc.original_filename,
                "page": page,
                "snippet": get_text_snippet(doc.text_preview or "", active_terms, query),
                "relevance": round(score, 2),
                "is_pdf": doc.extension.lower() == "pdf",
                "is_docx": doc.extension.lower() == "docx",
                "is_txt": doc.extension.lower() == "txt",
                "is_image": doc.extension.lower() in {"png", "jpg", "jpeg"},
                "image_pages": image_pages,
                "image_indexes": image_indexes,
            }
        )
    return results


# =========================================================
# Templates
# =========================================================
BASE_CSS = """
<style>
body { font-family: Arial, sans-serif; margin: 24px; background:#f5f7fb; color:#222; }
.container { max-width: 1280px; margin:auto; background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }
.topbar { display:flex; justify-content:space-between; align-items:flex-start; gap:16px; flex-wrap:nowrap; }
.topbar-brand { flex:1 1 auto; min-width:0; }
.topbar-brand h1 { margin:0 0 6px 0; line-height:1.2; }
.actions { display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; align-items:flex-start; flex:0 0 auto; margin-left:auto; }
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
.stats-panel, .search-meta-panel, .search-filter-panel {
  margin-top:16px;
  padding:16px 18px;
  background:#eef1f5;
  border:1px solid #d8dee6;
  border-radius:10px;
}
.panel-title {
  font-size:15px;
  font-weight:700;
  color:#334155;
  margin:0 0 12px 0;
  padding-bottom:8px;
  border-bottom:1px solid #d8dee6;
}
.panel-section { margin-top:14px; }
.panel-section:first-of-type { margin-top:0; }
.panel-label {
  font-size:13px;
  font-weight:700;
  color:#475569;
  margin-bottom:8px;
  letter-spacing:0.02em;
  text-transform:uppercase;
}
.stats-grid {
  display:flex;
  flex-wrap:wrap;
  gap:12px;
}
.stat-item {
  flex:1 1 180px;
  background:#fff;
  border:1px solid #dbe1ea;
  border-radius:8px;
  padding:12px 14px;
}
.stat-label { display:block; font-size:12px; color:#64748b; margin-bottom:4px; }
.stat-value { display:block; font-size:22px; font-weight:700; color:#0f172a; }
.term-chip-list { display:flex; flex-wrap:wrap; gap:8px; }
.term-chip {
  display:inline-block;
  background:#fff;
  border:1px solid #cbd5e1;
  color:#334155;
  padding:5px 10px;
  border-radius:999px;
  font-size:13px;
}
.term-chip.active {
  background:#dbeafe;
  border-color:#93c5fd;
  color:#1d4ed8;
  font-weight:700;
}
.term-chip.exact {
  border-color:#86efac;
  background:#ecfdf5;
}
.meta-line { margin-top:6px; line-height:1.6; color:#334155; }
.btn-small { padding:7px 12px; font-size:13px; }
.search-filter-panel .term-filter-row {
  display:flex;
  flex-wrap:wrap;
  gap:10px 16px;
  margin-bottom:12px;
}
.search-filter-panel label {
  display:inline-flex;
  align-items:center;
  gap:6px;
  cursor:pointer;
  background:#fff;
  border:1px solid #dbe1ea;
  border-radius:999px;
  padding:6px 12px;
}
.search-filter-panel label:has(input:checked) {
  background:#dbeafe;
  border-color:#93c5fd;
}
.search-filter-actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:4px; }
.snippet, .full-block { background:#fff; padding:12px; white-space:pre-wrap; border:1px solid #eee; border-radius:6px; font-family:Consolas, monospace; }
.snippet-columns {
  display:grid;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
  gap:12px;
  align-items:start;
  margin-top:10px;
}
.snippet-panel {
  min-width:0;
}
.snippet-panel .snippet {
  margin-top:8px;
  height:100%;
  box-sizing:border-box;
}
.snippet-panel-original .snippet {
  border-left:4px solid #0069d9;
}
.snippet-panel-translation .snippet {
  border-left:4px solid #28a745;
}
@media (max-width: 900px) {
  .snippet-columns {
    grid-template-columns: 1fr;
  }
}
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
.result-page-img img {
  max-width: 240px;
  max-height: 240px;
  border: 1px solid #ccc;
  border-radius: 6px;
}
.result-page-img {
  margin: 10px 0 12px 0;
}
.document-match-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 16px;
  margin-top: 12px;
}
.document-match-card {
  background: #fff;
  border: 1px solid #dbe1ea;
  border-radius: 10px;
  padding: 12px;
}
.document-match-card img {
  width: 100%;
  max-height: 260px;
  object-fit: contain;
  border: 1px solid #ddd;
  border-radius: 6px;
  background: #fff;
}
.document-match-card .doc-title {
  font-weight: 700;
  margin-bottom: 8px;
  line-height: 1.4;
}
.document-match-card .doc-snippet {
  margin-top: 10px;
  font-size: 13px;
}
@media (max-width: 900px) {
  .result-page-img img {
    max-width: 100%;
    height: auto;
  }
}
.clearfix::after { content:""; display:table; clear:both; }
@media (max-width: 900px) {
  .grid { grid-template-columns: 1fr; }
  input[type=text] { width:100%; }
  .topbar { flex-wrap:wrap; }
  .actions { width:100%; justify-content:flex-end; }
}
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
        <div class="topbar-brand">
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

    <form method="GET" action="/" id="search-form">
        <input type="text" name="q" value="{{ query }}" placeholder="Search requirement ID, Dutch/English keyword, section, definition">
        <button type="submit">Search</button>

        {% if query and all_expanded_terms %}
        <div class="search-filter-panel">
            <div class="panel-title">Term filters</div>
            <div class="panel-section">
                <div class="panel-label">Select search terms</div>
                <div class="small" style="margin-bottom:10px;">Exact search word is checked by default. Choose additional related terms if needed.</div>
                <div class="term-filter-row">
                    {% for term in all_expanded_terms %}
                    <label>
                        <input type="checkbox" name="term" value="{{ term }}"
                            {% if term in selected_terms %}checked{% endif %}>
                        <span{% if term in exact_terms %} style="font-weight:700;"{% endif %}>{{ term }}</span>
                    </label>
                    {% endfor %}
                </div>
                <div class="search-filter-actions">
                    <button type="submit" class="btn btn-green btn-small">Apply term filters</button>
                    <button type="button" class="btn btn-gray btn-small" onclick="selectAllTerms(true)">Select all</button>
                    <button type="button" class="btn btn-gray btn-small" onclick="selectAllTerms(false)">Clear all</button>
                    <button type="button" class="btn btn-gray btn-small" onclick="selectExactTermsOnly()">Exact only</button>
                </div>
                {% if selected_terms %}
                <div class="meta-line small" style="margin-top:10px;">
                    <strong>Active terms:</strong> {{ selected_terms|join(', ') }}
                </div>
                {% endif %}
            </div>
        </div>
        {% endif %}
    </form>

    <script>
    function selectAllTerms(checked) {
        document.querySelectorAll('#search-form input[name="term"]').forEach(function(el) {
            el.checked = checked;
        });
    }
    function selectExactTermsOnly() {
        const exact = new Set({{ exact_terms|tojson }});
        document.querySelectorAll('#search-form input[name="term"]').forEach(function(el) {
            el.checked = exact.has(el.value);
        });
    }
    </script>

    {% if doc_count == 0 %}
        <div class="warning">
            No documents are indexed yet. Log in as admin, upload PDF files, then use <strong>Reindex</strong> in Documents.
        </div>
    {% endif %}

    {% if query %}
        <div class="search-meta-panel">
            <div class="panel-title">Search analysis</div>

            {% if all_expanded_terms %}
            <div class="panel-section">
                <div class="panel-label">Expanded terms</div>
                <div class="term-chip-list">
                    {% for term in all_expanded_terms %}
                    <span class="term-chip{% if term in selected_terms %} active{% endif %}{% if term in exact_terms %} exact{% endif %}">{{ term }}</span>
                    {% endfor %}
                </div>
            </div>
            {% endif %}

            {% if summary.top_categories %}
            <div class="panel-section">
                <div class="panel-label">Top categories</div>
                <div class="meta-line">
                    {% for cat, cnt in summary.top_categories %}
                        <span class="badge">{{ cat }} ({{ cnt }})</span>
                    {% endfor %}
                </div>
            </div>
            {% endif %}

            {% if summary.top_requirement_ids %}
            <div class="panel-section">
                <div class="panel-label">Top requirement IDs</div>
                <div class="meta-line">{{ summary.top_requirement_ids|join(', ') }}</div>
            </div>
            {% endif %}
        </div>

        {% if document_matches %}
        <div class="search-meta-panel">
            <div class="panel-title">Related pages / drawings</div>
            <div class="document-match-grid">
                {% for doc in document_matches %}
                <div class="document-match-card">
                    <div class="doc-title">
                        {% if doc.is_pdf %}
                        <a href="{{ url_for('pdf_viewer', document_id=doc.document_id, page=doc.page or 1) }}" target="_blank">
                            {{ doc.filename }}
                        </a>
                        {% else %}
                            {{ doc.filename }}
                        {% endif %}
                    </div>
                    <div class="meta">
                        Page: <strong>{{ doc.page or 1 }}</strong>
                        | Relevance: <strong>{{ doc.relevance }}</strong>
                    </div>
                    {% if doc.is_pdf and doc.image_pages %}
                    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr)); gap:8px;">
                        {% for page_num in doc.image_pages %}
                        <a href="{{ url_for('pdf_viewer', document_id=doc.document_id, page=page_num) }}" target="_blank">
                            <img
                                src="{{ url_for('page_image', document_id=doc.document_id, page=page_num) }}"
                                loading="lazy"
                                alt="PDF page {{ page_num }} preview"
                                onerror="this.style.display='none';"
                            />
                        </a>
                        {% endfor %}
                    </div>
                    {% endif %}
                    {% if doc.is_docx and doc.image_indexes %}
                    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr)); gap:8px;">
                        {% for image_index in doc.image_indexes %}
                        <a href="{{ url_for('docx_viewer', document_id=doc.document_id, page=1) }}" target="_blank">
                            <img
                                src="{{ url_for('docx_image', document_id=doc.document_id, image_index=image_index) }}"
                                loading="lazy"
                                alt="DOCX drawing {{ image_index + 1 }}"
                                onerror="this.style.display='none';"
                            />
                        </a>
                        {% endfor %}
                    </div>
                    {% endif %}
                    <div class="snippet doc-snippet">{{ doc.snippet|safe }}</div>
                </div>
                {% endfor %}
            </div>
        </div>
        {% endif %}

        <h3>Results for "{{ query }}"</h3>

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
                        {% if res.major_section or res.section %}
                        | Section: <strong>{{ res.major_section or res.section }}</strong>
                        {% endif %}
                        | Category: <strong>{{ res.category }}</strong>
                        {% if res.ocr_used %}| <strong>OCR used</strong>{% endif %}
                    </div>

                    {% if res.is_pdf %}
                    <div class="result-page-img">
                        <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">
                            <img
                                src="{{ url_for('page_image', document_id=res.document_id, page=res.page or 1) }}"
                                loading="lazy"
                                alt="Page {{ res.page or 1 }} preview"
                                onerror="this.style.display='none';"
                            />
                        </a>
                    </div>
                    {% endif %}

                    {% if res.definition %}
                        <div><strong>Definition:</strong> {{ res.definition }}</div>
                    {% endif %}
                    {% if res.summary %}
                        <div><strong>Summary:</strong> {{ res.summary }}</div>
                    {% endif %}
                    {% if res.section and res.major_section and res.section != res.major_section %}
                        <div><strong>Subsection:</strong> {{ res.section }}</div>
                    {% elif res.section and not res.major_section %}
                        <div><strong>Section:</strong> {{ res.section }}</div>
                    {% endif %}

                    <div class="snippet-columns">
                        <div class="snippet-panel snippet-panel-original">
                            <strong>Table 1 / Dutch (Original)</strong>
                            <div class="snippet">{{ res.snippet|safe }}</div>
                        </div>
                        {% if res.snippet_en %}
                        <div class="snippet-panel snippet-panel-translation">
                            <strong>Table 2 / English (Translation)</strong>
                            <div class="snippet">{{ res.snippet_en|safe }}</div>
                        </div>
                        {% endif %}
                    </div>

                    <div class="meta">
                        {% if res.is_document_fallback %}
                            {% if res.is_pdf %}
                            <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open PDF at page {{ res.page or 1 }}</a>
                            {% elif res.is_docx %}
                            <a href="{{ url_for('docx_viewer', document_id=res.document_id, page=1) }}" target="_blank">Open DOCX translation</a>
                            {% elif res.is_txt %}
                            <a href="{{ url_for('document_view', document_id=res.document_id, page=1) }}" target="_blank">Open TXT translation</a>
                            {% endif %}
                        {% else %}
                            <a href="{{ url_for('requirement_detail', block_id=res.block_id) }}">Open full requirement</a>
                            {% if res.is_pdf %}
                                | <a href="{{ url_for('pdf_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open PDF at page {{ res.page or 1 }}</a>
                            {% elif res.is_docx %}
                                | <a href="{{ url_for('docx_viewer', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open DOCX translation</a>
                            {% elif res.is_txt %}
                                | <a href="{{ url_for('document_view', document_id=res.document_id, page=res.page or 1) }}" target="_blank">Open TXT translation</a>
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
        {% if block.major_section %}<div><strong>Section:</strong> {{ block.major_section }}</div>{% endif %}
        {% if block.section and block.section != block.major_section %}<div><strong>Subsection:</strong> {{ block.section }}</div>{% endif %}
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
        {% elif block.document.extension == 'docx' %}
            <a class="btn btn-purple" href="{{ url_for('docx_viewer', document_id=block.document.id, page=1) }}" target="_blank">Open DOCX translation</a>
        {% elif block.document.extension == 'txt' %}
            <a class="btn btn-purple" href="{{ url_for('document_view', document_id=block.document.id, page=1) }}" target="_blank">Open TXT translation</a>
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
        Max size: 100 MB per request<br>
        Storage backend: <strong>{{ storage_backend }}</strong>
        {% if storage_status.persistent %}
        <br>Bucket: <strong>{{ storage_status.bucket }}</strong>
        {% if storage_status.key_prefix and storage_status.key_prefix != '(root)' %}
        <br>Key prefix: <strong>{{ storage_status.key_prefix }}</strong>
        {% endif %}
        {% endif %}
        {% if not storage_persistent %}
        <br><span style="color:#b45309;"><strong>Warning:</strong> Files are stored on temporary disk and will be lost after a Render restart. Configure Supabase S3 for persistent PDF storage.</span>
        {% endif %}
    </div>

    {% if storage_status.persistent and not storage_status.ok %}
    <div class="flash" style="background:#fdecea; border-left-color:#dc3545;">
        <strong>Storage not ready:</strong> {{ storage_status.message }}<br>
        {{ storage_status.hint }}
    </div>
    {% endif %}

    <form method="POST" enctype="multipart/form-data"{% if storage_status.persistent and not storage_status.ok %} onsubmit="alert('Fix Supabase storage configuration before uploading.'); return false;"{% endif %}>
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

    <div class="stats-panel">
        <div class="panel-title">Index statistics</div>
        <div class="stats-grid">
            <div class="stat-item">
                <span class="stat-label">Indexed documents</span>
                <span class="stat-value">{{ doc_count }}</span>
            </div>
            <div class="stat-item">
                <span class="stat-label">Requirement blocks</span>
                <span class="stat-value">{{ block_count }}</span>
            </div>
            <div class="stat-item">
                <span class="stat-label">Tables</span>
                <span class="stat-value">{{ table_count }}</span>
            </div>
        </div>
    </div>

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
                <th>File</th>
                <th>Uploaded</th>
                <th>Actions</th>
            </tr>
            {% for row in doc_rows %}
            {% set d = row.record %}
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
                <td>
                    {% if row.file_available %}
                    <span style="color:#198754; font-weight:700;">OK</span>
                    {% else %}
                    <span style="color:#dc3545; font-weight:700;">Missing</span>
                    {% endif %}
                </td>
                <td>{{ d.uploaded_at }}</td>
                <td>
                    <a href="{{ url_for('reindex_document', document_id=d.id) }}">Reindex</a>
                    |
                    <a href="{{ url_for('delete_document', document_id=d.id) }}" onclick="return confirm('Delete this document?');">Delete</a>
                    {% if d.extension == 'pdf' %}
                    | <a href="{{ url_for('pdf_viewer', document_id=d.id, page=1) }}" target="_blank">View PDF</a>
                    | <a href="{{ url_for('export_translation_pdf', document_id=d.id) }}">Export EN PDF</a>
                    {% elif d.extension == 'docx' %}
                    | <a href="{{ url_for('docx_viewer', document_id=d.id, page=1) }}" target="_blank">View DOCX</a>
                    | <a href="{{ url_for('export_translation_pdf', document_id=d.id) }}">Export EN PDF</a>
                    {% elif d.extension == 'txt' %}
                    | <a href="{{ url_for('document_view', document_id=d.id, page=1) }}" target="_blank">View TXT</a>
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
        {% if block.major_section %}<div><strong>Section:</strong> {{ block.major_section }}</div>{% endif %}
        {% if block.section and block.section != block.major_section %}<div><strong>Subsection:</strong> {{ block.section }}</div>{% endif %}
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
        {% elif block.document.extension == 'docx' %}
            <a class="btn btn-purple" href="{{ url_for('docx_viewer', document_id=block.document.id, page=1) }}" target="_blank">Open DOCX translation</a>
        {% elif block.document.extension == 'txt' %}
            <a class="btn btn-purple" href="{{ url_for('document_view', document_id=block.document.id, page=1) }}" target="_blank">Open TXT translation</a>
        {% endif %}
    </div>
</div>
</body>
</html>
"""

TRANSLATABLE_EXTENSIONS = {"pdf", "docx", "txt"}


def build_document_view_context(doc, page=1):
    ext = doc.extension.lower()
    if ext not in TRANSLATABLE_EXTENSIONS:
        return None

    file_available = storage.document_exists(doc.stored_filename, DOC_FOLDER)
    total_pages = count_document_pages(doc)
    if total_pages <= 0:
        total_pages = 1
    page = max(1, min(page, total_pages))

    original_text = ""
    if ext in {"docx", "txt"}:
        chunks = split_text_into_virtual_pages(get_document_full_text(doc))
        original_text = chunks[page - 1] if chunks else ""
    elif ext == "pdf" and not file_available:
        original_text = extract_page_text_from_preview(doc, page)

    return {
        "document_id": doc.id,
        "filename": doc.original_filename,
        "extension": ext,
        "page": page,
        "total_pages": total_pages,
        "view_mode": "pdf" if ext == "pdf" else "text",
        "file_available": file_available,
        "pdf_available": file_available and ext == "pdf",
        "pdf_file_missing": pdf_source_file_missing(doc) if ext == "pdf" else False,
        "pdf_url": url_for("serve_document", document_id=doc.id) if ext == "pdf" and file_available else "",
        "download_url": url_for("serve_document", document_id=doc.id) if file_available else "",
        "original_text": original_text,
        "translation_enabled": translation.translation_enabled(),
        "page_label": "Page" if ext == "pdf" else "Section",
        "viewer_route": "document_view",
    }


DOCUMENT_VIEWER_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Document Viewer</title>
""" + BASE_CSS + """
<style>
body { margin:0; }
.toolbar { padding:12px 16px; background:white; border-bottom:1px solid #ddd; display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; }
.viewer-layout { display:flex; height: calc(100vh - 60px); }
.original-pane { flex: 1 1 58%; min-width: 320px; border-right:1px solid #ddd; position:relative; overflow:auto; }
.translation-pane { flex: 1 1 42%; min-width: 280px; overflow:auto; background:#f8fafc; padding:16px; }
iframe { width:100%; height:100%; border:none; }
.file-missing { padding:24px; color:#842029; background:#f8d7da; border:1px solid #f5c2c7; margin:16px; border-radius:8px; line-height:1.6; }
.original-text-box { margin:16px; padding:16px; background:white; border:1px solid #e2e8f0; border-radius:8px; white-space:pre-wrap; line-height:1.55; }
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
            | {{ page_label }} {{ page }} / {{ total_pages }}
        </div>
        <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
            <div class="page-nav">
                {% if page > 1 %}
                <a class="btn btn-gray btn-small" href="{{ url_for(viewer_route, document_id=document_id, page=page-1) }}">Prev</a>
                {% endif %}
                {% if page < total_pages %}
                <a class="btn btn-gray btn-small" href="{{ url_for(viewer_route, document_id=document_id, page=page+1) }}">Next</a>
                {% endif %}
            </div>
            {% if download_url %}
            <a class="btn btn-gray btn-small" href="{{ download_url }}">Download original</a>
            {% endif %}
            {% if translation_enabled %}
            <button class="btn btn-green btn-small" id="reload-translation">Refresh</button>
            <button class="btn btn-purple btn-small" id="export-translation" disabled>Export this page (.txt)</button>
            <a class="btn btn-orange btn-small" href="{{ url_for('export_translation_pdf', document_id=document_id, lang='en') }}">Export full EN PDF</a>
            {% endif %}
            <a href="{{ url_for('home') }}">Back</a>
        </div>
    </div>
    <div class="viewer-layout">
        <div class="original-pane">
            {% if view_mode == 'pdf' %}
                {% if pdf_available %}
                <iframe src="{{ pdf_url }}#page={{ page }}"></iframe>
                {% elif original_text %}
                <div style="padding:16px;">
                    <h3 style="margin-top:0;">Original text — {{ page_label }} {{ page }}</h3>
                    <div class="notice" style="margin-bottom:12px;">
                        <strong>PDF file not found on server.</strong>
                        Re-upload in <strong>Documents</strong> and click <strong>Reindex</strong> to restore the PDF preview.
                        Showing indexed text below.
                    </div>
                    <div class="original-text-box">{{ original_text }}</div>
                </div>
                {% else %}
                <div class="file-missing">
                    <strong>Original PDF not found on server.</strong><br>
                    Re-upload the file in <strong>Documents</strong>, then click <strong>Reindex</strong>.
                </div>
                {% endif %}
            {% else %}
                <div style="padding:16px;">
                    <h3 style="margin-top:0;">Original (Dutch)</h3>
                    {% if original_text %}
                    <div class="original-text-box">{{ original_text }}</div>
                    {% else %}
                    <div class="file-missing">
                        <strong>Original text not available.</strong><br>
                        Re-upload the file and click <strong>Reindex</strong>.
                    </div>
                    {% endif %}
                </div>
            {% endif %}
        </div>
        {% if translation_enabled %}
        <div class="translation-pane">
            <div class="translation-status" id="translation-status">Loading translation for {{ page_label|lower }} {{ page }}...</div>
            <div class="translation-box">
                <h4>English translation — {{ page_label }} {{ page }} / {{ total_pages }}</h4>
                <div id="translated-text">...</div>
            </div>
        </div>
        {% endif %}
    </div>
    {% if translation_enabled %}
    <script>
    const documentId = {{ document_id }};
    const pageNum = {{ page }};
    const totalPages = {{ total_pages }};
    const pageLabel = {{ page_label|tojson }};
    const fileBase = {{ filename|tojson }};
    const statusEl = document.getElementById("translation-status");
    const translatedEl = document.getElementById("translated-text");
    const exportBtn = document.getElementById("export-translation");
    let currentTranslation = "";

    async function readJsonResponse(resp) {
        const raw = await resp.text();
        if (!raw.trim()) {
            throw new Error(`Empty server response (${resp.status}).`);
        }
        try {
            return JSON.parse(raw);
        } catch (parseErr) {
            const preview = raw.replace(/\\s+/g, " ").slice(0, 160);
            throw new Error(
                preview.startsWith("<")
                    ? `Server returned HTML instead of JSON (${resp.status}). Try logging in again or redeploying the app.`
                    : `Invalid server response (${resp.status}).`
            );
        }
    }

    async function loadCurrentTranslation() {
        statusEl.textContent = `Translating ${pageLabel.toLowerCase()} ${pageNum} of ${totalPages}...`;
        translatedEl.textContent = "...";
        currentTranslation = "";
        exportBtn.disabled = true;

        try {
            const resp = await fetch(`/api/document/${documentId}/translate-page?page=${pageNum}`);
            const data = await readJsonResponse(resp);
            if (!resp.ok) {
                throw new Error(data.error || "Translation request failed");
            }

            currentTranslation = data.translation || "";
            translatedEl.textContent = currentTranslation || "(Translation unavailable)";
            const provider = data.provider || "unknown";
            statusEl.textContent = data.cached
                ? `${pageLabel} ${pageNum} — cached (${provider}). Use Next to translate the following page.`
                : `${pageLabel} ${pageNum} — translated with ${provider}. Use Next for the next page.`;
            if (data.error) {
                statusEl.textContent = data.error;
            }
            exportBtn.disabled = !currentTranslation.trim();
        } catch (err) {
            statusEl.textContent = "Translation error: " + err.message;
            translatedEl.textContent = "";
            exportBtn.disabled = true;
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
        link.download = `${safeName}_${pageLabel.toLowerCase()}_${pageNum}_en.txt`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
    }

    document.getElementById("reload-translation").addEventListener("click", loadCurrentTranslation);
    exportBtn.addEventListener("click", exportTranslation);
    loadCurrentTranslation();
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
    <div style="margin:12px 0; display:flex; gap:8px; flex-wrap:wrap;">
        <a class="btn btn-gray" href="{{ url_for('home') }}">Back</a>
        {% if current_user.is_authenticated and current_user.is_admin %}
        <a class="btn btn-green" href="{{ url_for('refresh_table_preview', document_id=doc.id) }}">Refresh table &amp; translate</a>
        {% endif %}
        {% if tables %}
        <a class="btn btn-purple" href="{{ url_for('export_table_xlsx', document_id=doc.id) }}">Export Excel</a>
        {% endif %}
    </div>

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
    all_expanded_terms = []
    selected_terms = []
    exact_terms = []
    summary = {"top_categories": [], "top_requirement_ids": []}
    document_matches = []

    if query:
        all_expanded_terms = get_all_expanded_terms(query)
        exact_terms = default_exact_search_terms(query)
        if request.args.getlist("term"):
            selected_terms = resolve_active_search_terms(query, request.args.getlist("term"))
        else:
            selected_terms = exact_terms

        document_matches = search_documents(query, active_terms=selected_terms)
        results, _ = search_requirements(query, active_terms=selected_terms)
        summary = grouped_search_summary(results)

    return render_template_string(
        HOME_TEMPLATE,
        query=query,
        results=results,
        document_matches=document_matches,
        all_expanded_terms=all_expanded_terms,
        selected_terms=selected_terms,
        exact_terms=exact_terms,
        summary=summary,
        doc_count=DocumentRecord.query.count(),
        block_count=RequirementBlock.query.count(),
        table_count=TablePreview.query.count(),
    )


@app.route("/healthz")
def healthz():
    storage_status = storage.get_storage_status()
    return jsonify(
        {
            "status": "ok" if storage_status.get("ok", True) else "degraded",
            "environment": APP_ENV,
            "storage_backend": storage.storage_backend_name(),
            "storage": storage_status,
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
    target_lang = normalize_translation_lang(payload.get("lang", "en"))
    if not source_text:
        return jsonify({"error": "text is required"}), 400
    if not translation.translation_enabled():
        return jsonify({"error": "translation is disabled"}), 503

    translated = translate_to_language(source_text, target_lang)
    return jsonify(
        {
            "source": source_text,
            "translation": translated,
            "lang": target_lang,
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
    target_lang = normalize_translation_lang(request.args.get("lang", "en"))
    try:
        result = get_or_translate_page(doc, page, target_lang=target_lang)
    except Exception as exc:
        print(f"Translate page API error for doc {document_id} page {page}: {exc}")
        return jsonify(
            {
                "page": page,
                "lang": target_lang,
                "translation": "",
                "error": f"Translation failed: {exc}",
            }
        ), 500

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

    total_pages = count_document_pages(doc)
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
    if not doc or doc.extension.lower() not in TRANSLATABLE_EXTENSIONS:
        abort(404)

    if not translation.translation_enabled():
        flash("Translation is disabled.")
        if doc.extension == "docx":
            return redirect(url_for("docx_viewer", document_id=document_id, page=1))
        if doc.extension == "txt":
            return redirect(url_for("document_view", document_id=document_id, page=1))
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    if not storage.document_exists(doc.stored_filename, DOC_FOLDER) and doc.extension == "pdf":
        flash("Source PDF not found. Re-upload the file and try again.")
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    try:
        target_lang = normalize_translation_lang(request.args.get("lang", "en"))
        pdf_bytes = build_translated_pdf_bytes(doc, target_lang=target_lang)
    except Exception as exc:
        print(f"Translated PDF export error for document {document_id}: {exc}")
        flash(f"Export failed: {exc}")
        if doc.extension == "docx":
            return redirect(url_for("docx_viewer", document_id=document_id, page=1))
        return redirect(url_for("pdf_viewer", document_id=document_id, page=1))

    base_name = secure_filename(
        Path(doc.original_filename).stem or f"document_{document_id}"
    )
    lang_suffix = "english"
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{base_name}_{lang_suffix}.pdf",
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
            user_cls = cast(Any, User)
            new_user = user_cls(username=username, role=role)
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
        try:
            return _handle_upload_post()
        except Exception as exc:
            db.session.rollback()
            print(f"Upload route error: {exc}")
            flash(f"Upload failed: {exc}")
            return redirect(url_for("upload_files"))

    return render_template_string(
        UPLOAD_TEMPLATE,
        allowed_extensions=", ".join(sorted(ALLOWED_EXTENSIONS)),
        storage_backend=storage.storage_backend_name(),
        storage_persistent=storage.object_storage_enabled(),
        storage_status=storage.get_storage_status(),
    )


def _handle_upload_post():
    storage_status = storage.get_storage_status()
    if storage.object_storage_enabled() and not storage_status.get("ok"):
        flash(storage_status.get("hint") or storage_status.get("message") or "Storage is not configured.")
        return redirect(url_for("upload_files"))

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        flash("No files selected.")
        return redirect(url_for("upload_files"))

    uploaded_count = 0
    skipped_count = 0
    restored_count = 0

    for file in files:
        if not file or file.filename == "":
            continue
        if not allowed_file(file.filename):
            flash(f"File type not allowed: {file.filename}")
            continue

        original_filename = secure_filename(file.filename or "")
        if not original_filename:
            flash("Invalid filename.")
            continue
        ext = file_ext(original_filename)

        temp_fd, temp_path = tempfile.mkstemp(suffix=f".{ext}")
        os.close(temp_fd)
        try:
            file.save(temp_path)
            file_hash = sha256_of_file(temp_path)

            existing = dedupe_lookup(file_hash)
            if existing:
                if not storage.document_exists(existing.stored_filename, DOC_FOLDER):
                    try:
                        restore_document_file_from_temp(existing, temp_path)
                        index_document_record(existing)
                        flash(
                            f"Restored missing file and reindexed: {original_filename}"
                        )
                        restored_count += 1
                    except Exception as exc:
                        db.session.rollback()
                        flash(
                            f"Could not restore duplicate file {original_filename}: {exc}"
                        )
                else:
                    flash(f"Duplicate file skipped: {original_filename}")
                    skipped_count += 1
                continue

            stored_filename = f"{uuid.uuid4().hex}.{ext}"
            try:
                size_bytes = storage.save_document(
                    stored_filename, temp_path, DOC_FOLDER
                )
                storage.verify_document_stored(stored_filename, DOC_FOLDER)
            except Exception as exc:
                flash(f"Upload failed for {original_filename}: {exc}")
                continue

            doc_cls = cast(Any, DocumentRecord)
            doc = doc_cls(
                original_filename=original_filename,
                stored_filename=stored_filename,
                extension=ext,
                file_hash=file_hash,
                size_bytes=size_bytes,
                uploaded_by=current_user.id,
            )
            db.session.add(doc)
            try:
                db.session.commit()
            except Exception as exc:
                db.session.rollback()
                storage.delete_document(stored_filename, DOC_FOLDER)
                flash(f"Database error for {original_filename}: {exc}")
                continue

            try:
                index_document_record(doc)
            except Exception as exc:
                db.session.rollback()
                print(f"Indexing error for {original_filename}: {exc}")
                flash(f"File saved but indexing failed for {original_filename}: {exc}")

            uploaded_count += 1
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    summary_parts = []
    if uploaded_count:
        summary_parts.append(f"uploaded {uploaded_count}")
    if restored_count:
        summary_parts.append(f"restored {restored_count}")
    if skipped_count:
        summary_parts.append(f"skipped {skipped_count} duplicate(s)")
    flash(
        "Upload complete: " + ", ".join(summary_parts) + "."
        if summary_parts
        else "No files were uploaded."
    )
    return redirect(url_for("upload_files"))

@app.route("/admin/documents")
@login_required
def admin_documents():
    if not is_admin():
        abort(403)

    docs = DocumentRecord.query.order_by(DocumentRecord.uploaded_at.desc()).all()
    doc_rows = [
        {
            "record": doc,
            "file_available": storage.document_exists(doc.stored_filename, DOC_FOLDER),
        }
        for doc in docs
    ]
    return render_template_string(
        DOCS_TEMPLATE,
        doc_rows=doc_rows,
        doc_count=DocumentRecord.query.count(),
        block_count=RequirementBlock.query.count(),
        table_count=TablePreview.query.count(),
    )

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
        "major_section": getattr(block, "major_section", "") or "",
        "page": block.page,
        "category": block.category,
        "definition": block.definition,
        "summary": block.summary,
        "full_text": block.full_text,
        "full_text_en": translate_to_english(block.full_text),
        "document": getattr(block, "document", None),
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
    file_available = storage.document_exists(doc.stored_filename, DOC_FOLDER)

    if (
        file_available
        and doc.extension.lower() in {"csv", "xlsx"}
        and table_preview_needs_regenerate(tables)
    ):
        try:
            tables = regenerate_table_previews(doc)
            if not tables:
                error_msg = "The spreadsheet has no readable sheets or table data."
        except FileNotFoundError:
            file_available = False
        except Exception as exc:
            print(f"Table preview regeneration error for document {document_id}: {exc}")
            error_msg = f"Could not build table preview: {exc}"

    if tables and table_preview_needs_translation(tables):
        refresh_table_translations_from_cache(tables)
        tables = TablePreview.query.filter_by(document_id=document_id).all()

    if not tables and not file_available:
        error_msg = (
            f"Source file '{doc.original_filename}' was not found in storage. "
            "Please go to Documents, delete this entry, upload the Excel file again, "
            "then click Reindex."
        )
    elif not tables and file_available:
        error_msg = error_msg or "No readable table data was found in this spreadsheet."

    return render_template_string(
        TABLE_TEMPLATE, doc=doc, tables=tables, error_msg=error_msg
    )


@app.route("/table/<int:document_id>/refresh")
@login_required
def refresh_table_preview(document_id):
    if not is_admin():
        abort(403)

    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        flash("Document not found.")
        return redirect(url_for("home"))

    if doc.extension.lower() not in {"csv", "xlsx"}:
        flash("This document does not contain a spreadsheet table.")
        return redirect(url_for("home"))

    try:
        tables = regenerate_table_previews(doc)
        if tables:
            flash(f"Table refreshed with translation ({len(tables)} sheet(s)).")
        else:
            flash("No readable table data found in this spreadsheet.")
    except FileNotFoundError:
        flash("Source file not found. Re-upload the spreadsheet and try again.")
    except Exception as exc:
        flash(f"Table refresh failed: {exc}")

    return redirect(url_for("table_preview", document_id=document_id))

@app.route("/document-view/<int:document_id>")
def document_view(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        abort(404)
    page = request.args.get("page", 1, type=int)
    context = build_document_view_context(doc, page=page)
    if not context:
        abort(404)
    return render_template_string(DOCUMENT_VIEWER_TEMPLATE, **context)


@app.route("/pdf/<int:document_id>")
def pdf_viewer(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "pdf":
        abort(404)
    page = request.args.get("page", 1, type=int)
    context = build_document_view_context(doc, page=page)
    if not context:
        abort(404)
    context["viewer_route"] = "pdf_viewer"
    return render_template_string(DOCUMENT_VIEWER_TEMPLATE, **context)


@app.route("/docx/<int:document_id>")
def docx_viewer(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "docx":
        abort(404)
    page = request.args.get("page", 1, type=int)
    context = build_document_view_context(doc, page=page)
    if not context:
        abort(404)
    context["viewer_route"] = "docx_viewer"
    return render_template_string(DOCUMENT_VIEWER_TEMPLATE, **context)


@app.route("/serve-document/<int:document_id>")
def serve_document(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension.lower() not in TRANSLATABLE_EXTENSIONS:
        abort(404)

    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        abort(404)

    file_bytes = storage.read_document_bytes(doc.stored_filename, DOC_FOLDER)
    mime_map = {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "txt": "text/plain",
    }
    return send_file(
        io.BytesIO(file_bytes),
        mimetype=mime_map.get(doc.extension.lower(), "application/octet-stream"),
        as_attachment=False,
        download_name=doc.original_filename,
    )


@app.route("/serve-pdf/<int:document_id>")
def serve_pdf(document_id):
    return serve_document(document_id)

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


@app.route("/docx-image/<int:document_id>/<int:image_index>")
def docx_image(document_id, image_index):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc or doc.extension != "docx":
        abort(404)
    if not storage.document_exists(doc.stored_filename, DOC_FOLDER):
        abort(404)
    with storage.open_document_local_path(doc.stored_filename, DOC_FOLDER) as file_path:
        image_data = get_docx_image(file_path, image_index)
    if image_data is None:
        abort(404)
    return send_file(io.BytesIO(image_data[0]), mimetype=image_data[1])

def safe_excel_sheet_name(name, used_names=None) -> str:
    value = re.sub(r"[\\/*?:\[\]]+", "_", str(name or "Sheet")).strip(" .")
    if not value:
        value = "Sheet"
    value = value[:31]
    if used_names is not None:
        base = value
        suffix = 1
        while value in used_names:
            tail = f"_{suffix}"
            value = f"{base[:31 - len(tail)]}{tail}"
            suffix += 1
        used_names.add(value)
    return value


def csv_text_to_dataframe(csv_text):
    if not (csv_text or "").strip():
        return None
    try:
        df = pd.read_csv(io.StringIO(csv_text), dtype=str, keep_default_na=False)
        for bad in ("nan", "NaN", "None", "<NA>"):
            df = df.replace(bad, "")
        return df
    except Exception as exc:
        print(f"CSV table parse error: {exc}")
        return None


def html_table_to_dataframe(html):
    if not (html or "").strip():
        return None
    try:
        frames = pd.read_html(io.StringIO(html))
        if not frames:
            return None
        df = frames[0].fillna("").astype(str)
        for bad in ("nan", "NaN", "None", "<NA>"):
            df = df.replace(bad, "")
        return df
    except Exception as exc:
        print(f"HTML table parse error: {exc}")
        return None


def table_preview_to_dataframe(table_row, lang="nl"):
    lang = (lang or "nl").strip().lower()
    csv_field = {
        "nl": "csv_text",
        "en": "csv_text_en",
        "es": "csv_text_es",
    }.get(lang, "csv_text")
    df = csv_text_to_dataframe(getattr(table_row, csv_field, "") or "")
    if df is not None and not df.empty:
        return df

    html_field = {
        "nl": "html_table",
        "en": "html_table_en",
        "es": "html_table_es",
    }.get(lang, "html_table")
    df = html_table_to_dataframe(getattr(table_row, html_field, "") or "")
    if df is not None and not df.empty:
        return df

    if lang == "nl" or not translation.translation_enabled():
        return df

    nl_df = csv_text_to_dataframe(table_row.csv_text)
    if nl_df is None or nl_df.empty:
        return None
    try:
        return translation.translate_dataframe_values(
            nl_df,
            target_lang=lang,
            cache_get=get_cached_translation,
            cache_set=store_cached_translation,
            max_cells=TRANSLATE_MAX_CELLS,
        )
    except Exception as exc:
        print(f"On-demand table translation failed ({lang}): {exc}")
        return None


def _coerce_excel_cell_value(value):
    if value in (None, "", "nan", "NaN", "None", "<NA>"):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    if re.fullmatch(r"-?\d+\.\d+", text):
        return float(text)
    return text


def _row_has_single_banner_cell(row_values):
    filled = [value for value in row_values if str(value).strip()]
    if len(filled) != 1:
        return False
    text = str(filled[0]).strip()
    return len(text) >= 20 or "|" in text or "Drawing" in text or "SPE." in text


_EXCEL_THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
_EXCEL_WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")


def _format_excel_table_block(ws, min_row, max_row, min_col, max_col):
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = _EXCEL_THIN_BORDER
            cell.alignment = _EXCEL_WRAP_ALIGN


def _auto_fit_excel_columns(ws, min_col, max_col, min_row, max_row):
    for col in range(min_col, max_col + 1):
        max_len = 10
        for row in range(min_row, max_row + 1):
            value = ws.cell(row=row, column=col).value
            if value is not None:
                max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 12), 55)


def _append_dataframe_to_sheet(ws, df, start_row, title=None):
    row_cursor = start_row
    if title:
        title_cell = ws.cell(row=row_cursor, column=1, value=title)
        title_cell.font = Font(bold=True, size=12)
        row_cursor += 1

    if df is None or df.empty:
        return row_cursor

    ncols = len(df.columns)
    header_row = row_cursor
    for col_offset, col_name in enumerate(df.columns, start=1):
        header_cell = ws.cell(row=header_row, column=col_offset, value=col_name)
        header_cell.font = Font(bold=True)
        header_cell.alignment = _EXCEL_WRAP_ALIGN
    row_cursor += 1

    data_start = row_cursor
    for row_offset, row_values in enumerate(df.values):
        for col_offset, value in enumerate(row_values, start=1):
            ws.cell(
                row=row_cursor + row_offset,
                column=col_offset,
                value=_coerce_excel_cell_value(value),
            )

    data_end = row_cursor + len(df) - 1
    if ncols > 0:
        for row_idx in range(data_start, data_end + 1):
            row_values = [
                ws.cell(row=row_idx, column=col_idx).value
                for col_idx in range(1, ncols + 1)
            ]
            if _row_has_single_banner_cell(row_values):
                ws.merge_cells(
                    start_row=row_idx,
                    start_column=1,
                    end_row=row_idx,
                    end_column=ncols,
                )
                merged_cell = ws.cell(row=row_idx, column=1)
                merged_cell.alignment = _EXCEL_WRAP_ALIGN

        _format_excel_table_block(ws, header_row, data_end, 1, ncols)
        _auto_fit_excel_columns(ws, 1, ncols, header_row, data_end)

    return data_end + 3


@app.route("/table/<int:document_id>/export-xlsx")
def export_table_xlsx(document_id):
    doc = db.session.get(DocumentRecord, document_id)
    if not doc:
        abort(404)

    tables = TablePreview.query.filter_by(document_id=document_id).all()
    if not tables:
        flash("No table data to export.")
        return redirect(url_for("table_preview", document_id=document_id))

    wb = Workbook()
    ws = wb.active
    ws.title = str(  # pyright: ignore[reportOptionalMemberAccess]
        safe_excel_sheet_name(
            Path(doc.original_filename).stem or f"document_{document_id}"
        )
    )

    row_cursor = 1
    wrote_content = False
    single_table = len(tables) == 1
    section = ""

    for index, table_row in enumerate(tables, start=1):
        if not single_table:
            section = table_row.sheet_name or table_row.preview_title or f"Table {index}"
            section_cell = cast(Any, ws.cell(row=row_cursor, column=1, value=section))  # pyright: ignore[reportOptionalMemberAccess]
            # pyright: ignore[reportOptionalMemberAccess]
            section_cell.font = Font(bold=True, size=13)
            row_cursor += 2

        nl_df = table_preview_to_dataframe(table_row, "nl")
        if nl_df is not None and not nl_df.empty:
            nl_title = "Table 1 - Dutch (Original)" if single_table else f"{section} - Dutch (Original)"
            row_cursor = _append_dataframe_to_sheet(ws, nl_df, row_cursor, nl_title)
            wrote_content = True

        en_df = table_preview_to_dataframe(table_row, "en")
        if en_df is not None and not en_df.empty:
            en_title = "Table 2 - English (Translation)" if single_table else f"{section} - English (Translation)"
            row_cursor = _append_dataframe_to_sheet(ws, en_df, row_cursor, en_title)
            wrote_content = True

    if not wrote_content:
        flash("No table content available to export.")
        return redirect(url_for("table_preview", document_id=document_id))

    xlsx_buffer = io.BytesIO()
    wb.save(xlsx_buffer)
    xlsx_buffer.seek(0)
    base_name = secure_filename(Path(doc.original_filename).stem or f"document_{document_id}")
    return send_file(
        xlsx_buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{base_name}_tables.xlsx",
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
    if "major_section" not in block_cols:
        db.session.execute(text("ALTER TABLE requirement_blocks ADD COLUMN major_section VARCHAR(300)"))
        db.session.commit()

    table_cols = {col["name"] for col in inspector.get_columns("table_previews")}
    if "html_table_en" not in table_cols:
        db.session.execute(text("ALTER TABLE table_previews ADD COLUMN html_table_en TEXT"))
        db.session.commit()
    if "html_table_es" not in table_cols:
        db.session.execute(text("ALTER TABLE table_previews ADD COLUMN html_table_es TEXT"))
        db.session.commit()
    if "csv_text_en" not in table_cols:
        db.session.execute(text("ALTER TABLE table_previews ADD COLUMN csv_text_en TEXT"))
        db.session.commit()
    if "csv_text_es" not in table_cols:
        db.session.execute(text("ALTER TABLE table_previews ADD COLUMN csv_text_es TEXT"))
        db.session.commit()

    page_translation_cols = set()
    if "document_page_translations" in inspector.get_table_names():
        page_translation_cols = {
            col["name"] for col in inspector.get_columns("document_page_translations")
        }
    if page_translation_cols and "translated_text_es" not in page_translation_cols:
        db.session.execute(
            text("ALTER TABLE document_page_translations ADD COLUMN translated_text_es TEXT")
        )
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
