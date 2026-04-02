"""
Document text extraction: images (OCR), PDFs, Word documents.
- Offline: RapidOCR (ONNX Runtime) -- compact, no PyTorch needed.
- Online: Vision API (GPT-4o / DeepSeek) -- far more accurate for complex layouts.
"""

import os
import re
import json
import base64
import unicodedata

_ocr_engine = None

RAPIDOCR_LANG_MAP = {
    "cn": "ch",
    "en": "en",
    "jp": "japan",
    "kr": "korean",
    "vi": "vi",
    "th": "thai",
}


def _get_ocr(lang: str = "jp"):
    """Lazy-init RapidOCR engine for the given language."""
    global _ocr_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
        return _ocr_engine
    except ImportError:
        raise RuntimeError(
            "rapidocr-onnxruntime is not installed.\n"
            "Run: pip install rapidocr-onnxruntime"
        )


def read_image(path: str, lang: str = "jp") -> str:
    """OCR an image file, return extracted text."""
    ocr = _get_ocr(lang)
    result, _ = ocr(path)
    if not result:
        return ""
    lines = [item[1] for item in result]
    return "\n".join(lines)


def read_pdf(path: str, lang: str = "jp") -> str:
    """Extract text from PDF. Falls back to OCR for scanned pages."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise RuntimeError(
            "PyMuPDF is not installed.\n"
            "Run: pip install PyMuPDF"
        )

    doc = fitz.open(path)
    all_text = []

    for page in doc:
        text = page.get_text().strip()
        if len(text) > 20:
            all_text.append(text)
        else:
            pix = page.get_pixmap(dpi=300)
            img_path = os.path.join(
                os.environ.get("TEMP", "/tmp"),
                f"_mcv2_page_{page.number}.png"
            )
            pix.save(img_path)
            try:
                ocr_text = read_image(img_path, lang)
                if ocr_text:
                    all_text.append(ocr_text)
            finally:
                if os.path.exists(img_path):
                    os.remove(img_path)

    doc.close()
    return "\n".join(all_text)


def read_word(path: str) -> str:
    """Extract text from a .docx file."""
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError(
            "python-docx is not installed.\n"
            "Run: pip install python-docx"
        )

    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                paragraphs.append("\t".join(cells))

    return "\n".join(paragraphs)


def read_excel(path: str) -> str:
    """Extract text from Excel, each row as tab-separated line."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RuntimeError(
            "openpyxl is not installed.\n"
            "Run: pip install openpyxl"
        )

    wb = load_workbook(path, read_only=True, data_only=True)
    lines = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, values_only=True):
            cells = [str(c) if c is not None else "" for c in row]
            line = "\t".join(cells).strip()
            if line:
                lines.append(line)
    wb.close()
    return "\n".join(lines)


def read_file(path: str, lang: str = "jp") -> str:
    """Auto-detect file type and extract text (offline OCR)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"):
        return read_image(path, lang)
    elif ext == ".pdf":
        return read_pdf(path, lang)
    elif ext in (".docx", ".doc"):
        return read_word(path)
    elif ext in (".xlsx", ".xls"):
        return read_excel(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


# ──────────────────────────────────────────────────────────────
# Vision API: send images directly to GPT-4o / DeepSeek Vision
# ──────────────────────────────────────────────────────────────

VISION_SYSTEM_PROMPT = """You are a menu data extractor. Analyze the menu image and extract ALL menu items.

Return ONLY a JSON array, no other text. Each element:
{"name": "item name (original language)", "name_alt": "item name in other language if visible", "price": number, "category": "category/section name"}

Rules:
- Extract EVERY menu item visible, do not skip any
- "name" = the primary name in the menu's main language
- "name_alt" = secondary name if shown (e.g., Chinese name below Japanese name), or "" if none
- "price" = numeric price (no currency symbol). If no price visible, use 0
- "category" = the section header this item belongs to (e.g., "前菜", "肉類", "Appetizer")
- Items with a number prefix like "45. 鶏肉の唐辛子炒め" → name = "鶏肉の唐辛子炒め" (strip the number)
- Ignore decorative text, disclaimers, and restaurant descriptions
- If a set/course menu has a single price for the whole set, category = the set name, price = set price for each item"""

VISION_PROVIDERS = {
    "OpenAI": {
        "base_url": None,
        "model": "gpt-4o-mini",
    },
    "DeepSeek": {
        "base_url": "https://api.deepseek.com",
        "model": "deepseek-chat",
    },
    "Doubao": {
        "base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "model": "doubao-1.5-thinking-vision-pro",
    },
    "Claude": {
        "base_url": "https://api.anthropic.com/v1/",
        "model": "claude-sonnet-4-20250514",
    },
}


def _encode_image_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def _image_media_type(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    return {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".bmp": "image/bmp",
        ".tiff": "image/tiff", ".tif": "image/tiff",
        ".webp": "image/webp",
    }.get(ext, "image/png")


def read_image_vision(path: str, api_key: str, provider: str = "OpenAI") -> list[dict]:
    """
    Use Vision API to extract structured menu items from an image.
    Returns: [{"name": ..., "name_alt": ..., "price": ..., "category": ...}]
    """
    try:
        from openai import OpenAI
    except ImportError:
        raise RuntimeError("openai is not installed.\nRun: pip install openai")

    cfg = VISION_PROVIDERS.get(provider, VISION_PROVIDERS["OpenAI"])
    client_kwargs = {"api_key": api_key}
    if cfg["base_url"]:
        client_kwargs["base_url"] = cfg["base_url"]

    client = OpenAI(**client_kwargs)
    b64 = _encode_image_base64(path)
    media = _image_media_type(path)

    resp = client.chat.completions.create(
        model=cfg["model"],
        messages=[
            {"role": "system", "content": VISION_SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {
                    "url": f"data:{media};base64,{b64}"
                }},
                {"type": "text", "text": "Extract all menu items from this image."},
            ]},
        ],
        temperature=0.1,
        max_tokens=4096,
    )

    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

    try:
        items = json.loads(raw)
        if isinstance(items, list):
            return items
    except json.JSONDecodeError:
        pass

    print(f"[WARN] Vision API returned non-JSON:\n{raw[:500]}")
    return []


def read_pdf_vision(path: str, api_key: str, provider: str = "OpenAI",
                    lang: str = "jp") -> list[dict]:
    """Extract menu items from PDF pages via Vision API."""
    try:
        import fitz
    except ImportError:
        raise RuntimeError("PyMuPDF is not installed.\nRun: pip install PyMuPDF")

    doc = fitz.open(path)
    all_items = []

    for page in doc:
        pix = page.get_pixmap(dpi=200)
        img_path = os.path.join(
            os.environ.get("TEMP", "/tmp"),
            f"_mcv2_vision_p{page.number}.png"
        )
        pix.save(img_path)
        try:
            items = read_image_vision(img_path, api_key, provider)
            all_items.extend(items)
        finally:
            if os.path.exists(img_path):
                os.remove(img_path)

    doc.close()
    return all_items


def read_file_vision(path: str, api_key: str, provider: str = "OpenAI",
                     lang: str = "jp") -> list[dict]:
    """
    Read file using Vision API (for images/PDFs).
    Falls back to text extraction for Word/Excel.
    Returns: list of {"name", "name_alt", "price", "category"} dicts.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"):
        return read_image_vision(path, api_key, provider)
    elif ext == ".pdf":
        return read_pdf_vision(path, api_key, provider, lang)
    elif ext in (".docx", ".doc", ".xlsx", ".xls"):
        return None
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def detect_language(text: str) -> str:
    """
    Heuristic language detection based on Unicode character ranges.
    Returns: 'jp', 'cn', 'kr', 'th', 'vi', or 'en'
    """
    if not text:
        return "en"

    counters = {"jp": 0, "cn": 0, "kr": 0, "th": 0, "cjk": 0, "latin": 0}

    for ch in text:
        cp = ord(ch)
        if 0x3040 <= cp <= 0x309F or 0x30A0 <= cp <= 0x30FF:
            counters["jp"] += 1
        elif 0xAC00 <= cp <= 0xD7AF or 0x1100 <= cp <= 0x11FF:
            counters["kr"] += 1
        elif 0x0E00 <= cp <= 0x0E7F:
            counters["th"] += 1
        elif 0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF:
            counters["cjk"] += 1
        elif cp < 0x0250:
            counters["latin"] += 1

    if counters["jp"] > 5:
        return "jp"
    if counters["kr"] > 5:
        return "kr"
    if counters["th"] > 5:
        return "th"
    if counters["cjk"] > 5:
        if counters["jp"] > 0:
            return "jp"
        return "cn"

    viet_chars = set("ăâđêôơưàảãáạằẳẵắặầẩẫấậèẻẽéẹềểễếệìỉĩíịòỏõóọồổỗốộờởỡớợùủũúụừửữứựỳỷỹýỵ")
    viet_count = sum(1 for ch in text.lower() if ch in viet_chars)
    if viet_count > 3:
        return "vi"

    return "en"
