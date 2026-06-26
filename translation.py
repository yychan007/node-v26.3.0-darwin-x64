import hashlib
import os
import re

_translator = None
_enabled = None


def translation_enabled():
    global _enabled
    if _enabled is None:
        value = os.environ.get("ENABLE_TRANSLATION", "true").strip().lower()
        _enabled = value in {"1", "true", "yes", "on"}
    return _enabled


def _get_translator():
    global _translator
    if _translator is None:
        from deep_translator import GoogleTranslator

        _translator = GoogleTranslator(source="auto", target="en")
    return _translator


def should_skip_translation(text):
    value = (text or "").strip()
    if not value or len(value) < 2:
        return True
    if re.fullmatch(r"[\d\W_]+", value):
        return True
    if re.fullmatch(r"[A-Za-z0-9._\-/ ]{1,12}", value):
        return True
    return False


def _chunk_text(text, chunk_size=4500):
    text = text or ""
    if len(text) <= chunk_size:
        return [text]
    parts = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        if end < len(text):
            split_at = text.rfind("\n", start, end)
            if split_at > start + 500:
                end = split_at
        parts.append(text[start:end])
        start = end
    return parts


def translate_text(text, cache_get=None, cache_set=None):
    if not translation_enabled():
        return ""

    text = (text or "").strip()
    if not text or should_skip_translation(text):
        return text

    cache_key = hashlib.sha256(text.encode("utf-8")).hexdigest()
    if cache_get:
        cached = cache_get(cache_key)
        if cached is not None:
            return cached

    try:
        translator = _get_translator()
        translated_parts = []
        for chunk in _chunk_text(text):
            chunk = chunk.strip()
            if not chunk:
                continue
            if should_skip_translation(chunk):
                translated_parts.append(chunk)
                continue
            translated_parts.append(translator.translate(chunk))
        translated = "\n".join(part for part in translated_parts if part).strip()
    except Exception as exc:
        print(f"Translation error: {exc}")
        return ""

    if cache_set and translated:
        cache_set(cache_key, text, translated)
    return translated


def translate_dataframe_values(df, cache_get=None, cache_set=None, max_cells=180):
    if df is None:
        return df

    translated = df.copy().astype(str)
    seen = 0
    for col in translated.columns:
        for idx in translated.index:
            if seen >= max_cells:
                return translated
            value = translated.at[idx, col]
            if should_skip_translation(value):
                continue
            translated.at[idx, col] = translate_text(value, cache_get, cache_set)
            seen += 1
    return translated
