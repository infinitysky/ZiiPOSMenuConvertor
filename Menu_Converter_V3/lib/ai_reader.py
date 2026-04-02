"""
Document text extraction: images (OCR), PDFs, Word documents.
- Offline: RapidOCR 3.8+ with rapid-table for structured table recognition.
- Online: Vision API (GPT-4o / DeepSeek) -- far more accurate for complex layouts.
"""

import os
import re
import json
import base64
import unicodedata
import logging

log = logging.getLogger(__name__)

_ocr_engine = None
_table_engine = None

RAPIDOCR_LANG_MAP = {
    "cn": "ch",
    "en": "en",
    "jp": "japan",
    "kr": "korean",
    "vi": "vi",
    "th": "thai",
}


def _get_ocr(lang: str = "jp"):
    """Lazy-init RapidOCR engine with full PP-OCRv5 (rapidocr >= 3.8)."""
    global _ocr_engine
    if _ocr_engine is not None:
        return _ocr_engine
    try:
        from rapidocr import (
            RapidOCR, EngineType, LangDet, LangRec,
            ModelType, OCRVersion,
        )
        _ocr_engine = RapidOCR(params={
            "Det.engine_type": EngineType.ONNXRUNTIME,
            "Det.lang_type": LangDet.CH,
            "Det.model_type": ModelType.MOBILE,
            "Det.ocr_version": OCRVersion.PPOCRV5,
            "Cls.engine_type": EngineType.ONNXRUNTIME,
            "Cls.lang_type": LangDet.CH,
            "Cls.model_type": ModelType.MOBILE,
            "Cls.ocr_version": OCRVersion.PPOCRV5,
            "Rec.engine_type": EngineType.ONNXRUNTIME,
            "Rec.lang_type": LangRec.CH,
            "Rec.model_type": ModelType.MOBILE,
            "Rec.ocr_version": OCRVersion.PPOCRV5,
        })
        return _ocr_engine
    except ImportError:
        raise RuntimeError(
            "rapidocr is not installed.\n"
            "Run: pip install rapidocr>=3.8.0"
        )


def _get_table_engine():
    """Lazy-init RapidTable engine for structured table recognition."""
    global _table_engine
    if _table_engine is not None:
        return _table_engine
    try:
        from rapid_table import RapidTable
        _table_engine = RapidTable()
        return _table_engine
    except ImportError:
        log.warning("rapid-table not installed; table structure recognition disabled")
        return None


def read_image(path: str, lang: str = "jp") -> str:
    """OCR an image file, return extracted text."""
    ocr = _get_ocr(lang)
    result = ocr(path)
    if result is None or len(result) == 0:
        return ""
    return "\n".join(result.txts)


def read_image_with_boxes(path: str, lang: str = "jp") -> list[dict]:
    """
    OCR an image, return list of {'text', 'box', 'score'} dicts.
    box = [[x1,y1],[x2,y2],[x3,y3],[x4,y4]] (4-point polygon).
    """
    ocr = _get_ocr(lang)
    result = ocr(path)
    if result is None or len(result) == 0:
        return []
    items = []
    for i, txt in enumerate(result.txts):
        items.append({
            "text": txt,
            "box": result.boxes[i].tolist() if result.boxes is not None else [],
            "score": float(result.scores[i]) if result.scores is not None else 0.0,
        })
    return items


def read_image_table(path: str, lang: str = "jp") -> str | None:
    """
    Use rapid-table to recognize table structure from an image.
    RapidTable runs its own internal OCR to avoid box format mismatches.
    Returns HTML table string, or None if no table engine or no table found.
    """
    table_engine = _get_table_engine()
    if table_engine is None:
        return None
    try:
        table_output = table_engine(path)
        if table_output and table_output.pred_htmls:
            html = table_output.pred_htmls[0]
            if "<td>" in html or "<td " in html:
                return html
    except Exception as e:
        log.warning("rapid-table failed: %s", e)
    return None


def read_image_structured(path: str, lang: str = "jp") -> dict:
    """
    Best-effort structured read of an image:
    1. Try rapid-table for table structure -> parse HTML -> items
    2. Fall back to OCR with bounding boxes for spatial parsing
    3. Fall back to plain OCR text
    Returns {'mode': 'table'|'ocr', 'items': [...], 'text': str, 'boxes': [...]}
    """
    try:
        table_html = read_image_table(path, lang)
        if table_html:
            try:
                from .menu_parser import parse_table_html
            except ImportError:
                from menu_parser import parse_table_html
            items = parse_table_html(table_html)
            if items:
                return {"mode": "table", "items": items, "text": "", "boxes": []}
    except Exception as e:
        log.warning("Table recognition failed, falling back to OCR: %s", e)

    try:
        box_data = read_image_with_boxes(path, lang)
        text = "\n".join(d["text"] for d in box_data)
        return {"mode": "ocr", "items": [], "text": text, "boxes": box_data}
    except Exception as e:
        log.warning("OCR with boxes failed, falling back to plain OCR: %s", e)

    text = read_image(path, lang)
    return {"mode": "ocr", "items": [], "text": text, "boxes": []}


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
    """Use Vision API to extract structured menu items from an image."""
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


# ──────────────────────────────────────────────────────────────
# LLM Post-Refine: text-only structuring of OCR results
# ──────────────────────────────────────────────────────────────

OCR_REFINE_PROMPT = """You are a menu data structuring assistant. You receive raw OCR text blocks extracted from a restaurant menu image, each with approximate position coordinates.

Your task: parse these into a clean JSON array of menu items.

Return ONLY a JSON array. Each element:
{"name": "item name (primary language, usually Chinese/Japanese)", "name_alt": "item name in secondary language (usually English) or empty string", "price": number, "category": "section/category name"}

Rules:
- Each real menu item typically has: an optional number prefix (like "07"), a name in CJK characters, sometimes an English name, and a price
- Prices may appear as "$28.80" or "$28.80/份 $26.80/份" (regular/member) -- use the FIRST (regular) price
- Lines with no price and decorative styling (like "金风玉露·凉菜篇 COLD DISHES") are CATEGORY HEADERS, not items
- Strip item number prefixes (e.g., "07", "08") from names
- If an item has both Chinese and English names, put CJK in "name" and English in "name_alt"
- Ignore disclaimers, allergen markers (like ⑥⑧), restaurant info, and decorative text
- price should be a number (not string), 0 if unknown
- Do NOT skip any menu items"""


def refine_ocr_with_llm(box_data: list[dict], api_key: str,
                         provider: str = "OpenAI") -> list[dict]:
    """
    Send OCR text blocks (with positions) to an LLM for structured extraction.
    This is a TEXT-ONLY call (cheap, fast) -- not a vision call.
    Returns list of {'name', 'name_alt', 'price', 'category'} dicts.
    """
    try:
        from openai import OpenAI
    except ImportError:
        log.warning("openai not installed, skipping LLM refine")
        return []

    cfg = VISION_PROVIDERS.get(provider, VISION_PROVIDERS["OpenAI"])
    client_kwargs = {"api_key": api_key}
    if cfg["base_url"]:
        client_kwargs["base_url"] = cfg["base_url"]

    text_model = cfg.get("model", "gpt-4o-mini")

    lines = []
    for d in box_data:
        text = d.get("text", "").strip()
        if not text:
            continue
        box = d.get("box", [])
        if box and len(box) >= 4:
            x = int(min(pt[0] for pt in box))
            y = int(min(pt[1] for pt in box))
            lines.append(f"[x={x},y={y}] {text}")
        else:
            lines.append(text)

    if not lines:
        return []

    ocr_text = "\n".join(lines)
    if len(ocr_text) > 12000:
        ocr_text = ocr_text[:12000] + "\n... (truncated)"

    client = OpenAI(**client_kwargs)
    try:
        resp = client.chat.completions.create(
            model=text_model,
            messages=[
                {"role": "system", "content": OCR_REFINE_PROMPT},
                {"role": "user", "content": f"Here are the OCR text blocks from a menu:\n\n{ocr_text}\n\nExtract all menu items as JSON."},
            ],
            temperature=0.1,
            max_tokens=4096,
        )
    except Exception as e:
        log.warning("LLM refine API call failed: %s", e)
        return []

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

    log.warning("LLM refine returned non-JSON: %s", raw[:300])
    return []


def detect_language(text: str) -> str:
    """Heuristic language detection based on Unicode character ranges."""
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
