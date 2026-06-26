import hashlib
import json
import os
import re
import urllib.error
import urllib.request

_translator = None
_enabled = None

SYSTEM_PROMPT = (
    "You are a professional technical translator for electrical engineering and "
    "building standards documents (often Dutch NEN/IEC standards). "
    "Translate accurately into English. Preserve requirement IDs, section numbers, "
    "table references, units, and normative wording. "
    "Return only the translation without explanations."
)


def translation_enabled():
    global _enabled
    if _enabled is None:
        value = os.environ.get("ENABLE_TRANSLATION", "true").strip().lower()
        _enabled = value in {"1", "true", "yes", "on"}
    return _enabled


def translation_provider():
    provider = os.environ.get("TRANSLATION_PROVIDER", "google").strip().lower()
    if provider in {"openai", "gpt", "chatgpt"}:
        return "openai"
    if provider in {"anthropic", "claude"}:
        return "anthropic"
    return "google"


def provider_configured():
    provider = translation_provider()
    if provider == "openai":
        return bool(os.environ.get("OPENAI_API_KEY", "").strip())
    if provider == "anthropic":
        return bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())
    return True


def provider_status():
    provider = translation_provider()
    return {
        "enabled": translation_enabled(),
        "provider": provider,
        "configured": provider_configured(),
        "model": _provider_model(provider),
    }


def _provider_model(provider):
    if provider == "openai":
        return os.environ.get("OPENAI_MODEL", "gpt-4o-mini").strip() or "gpt-4o-mini"
    if provider == "anthropic":
        return (
            os.environ.get("ANTHROPIC_MODEL", "claude-3-5-haiku-latest").strip()
            or "claude-3-5-haiku-latest"
        )
    return "google-translate"


def _get_google_translator():
    global _translator
    if _translator is None:
        from deep_translator import GoogleTranslator

        source = os.environ.get("TRANSLATION_SOURCE_LANG", "auto").strip() or "auto"
        target = os.environ.get("TRANSLATION_TARGET_LANG", "en").strip() or "en"
        _translator = GoogleTranslator(source=source, target=target)
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


def _http_json_post(url, headers, payload, timeout=90):
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _translate_openai(text):
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set.")

    model = _provider_model("openai")
    payload = {
        "model": model,
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": text},
        ],
    }
    result = _http_json_post(
        "https://api.openai.com/v1/chat/completions",
        {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        payload,
    )
    return (result["choices"][0]["message"]["content"] or "").strip()


def _translate_anthropic(text):
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    model = _provider_model("anthropic")
    payload = {
        "model": model,
        "max_tokens": 4096,
        "temperature": 0.2,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": text}],
    }
    result = _http_json_post(
        "https://api.anthropic.com/v1/messages",
        {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        },
        payload,
    )
    parts = []
    for block in result.get("content", []):
        if block.get("type") == "text":
            parts.append(block.get("text", ""))
    return "\n".join(parts).strip()


def _translate_google(text):
    translator = _get_google_translator()
    return translator.translate(text)


def _translate_chunk(text):
    provider = translation_provider()
    if provider == "openai" and provider_configured():
        return _translate_openai(text)
    if provider == "anthropic" and provider_configured():
        return _translate_anthropic(text)
    return _translate_google(text)


def translate_text(text, cache_get=None, cache_set=None, cache_provider=None):
    if not translation_enabled():
        return ""

    text = (text or "").strip()
    if not text or should_skip_translation(text):
        return text

    provider = translation_provider()
    cache_key = hashlib.sha256(f"{provider}:{text}".encode("utf-8")).hexdigest()
    if cache_get:
        cached = cache_get(cache_key)
        if cached is not None:
            return cached

    try:
        translated_parts = []
        for chunk in _chunk_text(text):
            chunk = chunk.strip()
            if not chunk:
                continue
            if should_skip_translation(chunk):
                translated_parts.append(chunk)
                continue
            try:
                translated_parts.append(_translate_chunk(chunk))
            except Exception as exc:
                print(f"Primary translation error ({provider}): {exc}")
                if provider != "google":
                    translated_parts.append(_translate_google(chunk))
                else:
                    raise
        translated = "\n".join(part for part in translated_parts if part).strip()
    except Exception as exc:
        print(f"Translation error: {exc}")
        return ""

    if cache_set and translated:
        cache_set(cache_key, text, translated, provider if cache_provider else None)
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
