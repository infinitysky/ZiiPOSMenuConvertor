"""
Rule-based menu text parser.
Extracts structured menu items (itemcode, name, price, category) from OCR text.
"""

import re

CURRENCY_SYMBOLS = r'[¥￥$＄€£₩฿]'
PRICE_PATTERN = re.compile(
    r'(?:' + CURRENCY_SYMBOLS + r')\s*([0-9,]+(?:\.[0-9]{1,2})?)'
    r'|([0-9,]+(?:\.[0-9]{1,2})?)\s*(?:円|元|won|baht)'
    r'|(?:^|\s)([0-9,]{2,}(?:\.[0-9]{1,2})?)(?:\s*$|\s{2,})',
    re.IGNORECASE
)

SHARED_PRICE_PATTERN = re.compile(
    r'(?:' + CURRENCY_SYMBOLS + r')\s*([0-9,]+(?:\.[0-9]{1,2})?)'
    r'|([0-9,]+(?:\.[0-9]{1,2})?)\s*(?:円|元)',
    re.IGNORECASE
)

CATEGORY_HINT_CHARS = re.compile(r'^[=＝─━\-\*\#■□●○◆◇【】\[\]]+')
BULLET_PREFIX = re.compile(r'^[○◯◎●・\-\*]\s*')


def _parse_price(text: str) -> float:
    """Convert a price string to float."""
    cleaned = text.replace(",", "").replace("，", "").strip()
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0


def _extract_price(line: str) -> tuple:
    """Try to extract a price from a line. Returns (name, price) or (line, None)."""
    m = PRICE_PATTERN.search(line)
    if m:
        price_str = m.group(1) or m.group(2) or m.group(3)
        price = _parse_price(price_str)
        if price > 0:
            name = line[:m.start()].strip()
            name = re.sub(r'[\.\…·‧]+\s*$', '', name).strip()
            if not name:
                name = line[m.end():].strip()
            return name, price
    return line.strip(), None


def _is_category_header(line: str, has_items_after: bool = True) -> bool:
    """Heuristic: is this line a category/section header?"""
    stripped = line.strip()
    if not stripped:
        return False
    if CATEGORY_HINT_CHARS.search(stripped):
        cleaned = CATEGORY_HINT_CHARS.sub('', stripped).strip()
        if cleaned:
            return True
    if stripped.isupper() and len(stripped) > 2 and not any(c.isdigit() for c in stripped):
        return True
    _, price = _extract_price(stripped)
    if price is not None:
        return False
    if len(stripped) < 20 and has_items_after:
        return True
    return False


def detect_shared_price(text: str) -> float | None:
    """
    Detect if the document has a single shared price
    (e.g., 飲み放題 ¥1,800 -- all items share this price).
    """
    lines = text.strip().split("\n")
    prices_found = []

    for line in lines[:10]:
        for m in SHARED_PRICE_PATTERN.finditer(line):
            price_str = m.group(1) or m.group(2)
            p = _parse_price(price_str)
            if p >= 100:
                prices_found.append(p)

    item_prices = []
    for line in lines:
        _, price = _extract_price(line)
        if price is not None and price > 0:
            item_prices.append(price)

    if prices_found and len(item_prices) <= len(prices_found) + 2:
        return prices_found[0]

    return None


def _try_parse_tabular(text: str) -> list[dict] | None:
    """
    Try to parse tab-separated tabular data (e.g., from Excel).
    Returns items list if successful, None if text is not tabular.
    """
    lines = text.strip().split("\n")
    lines = [l for l in lines if l.strip()]
    if not lines:
        return None

    tab_lines = [l for l in lines if "\t" in l]
    if len(tab_lines) < len(lines) * 0.5:
        return None

    first_cols = lines[0].split("\t")
    num_cols = len(first_cols)
    if num_cols < 2:
        return None

    name_col = -1
    price_col = -1
    category_col = -1
    code_col = -1
    header_row = None

    HEADER_NAME = {
        "name", "名称", "名前", "品名", "description", "description1",
        "名称1", "item", "itemname", "menu", "商品名",
    }
    HEADER_PRICE = {
        "price", "価格", "金額", "单价", "價格", "price1", "价格",
        "価格1", "价格1",
    }
    HEADER_CAT = {
        "category", "分类", "カテゴリ", "分類", "类别", "類別", "group",
    }
    HEADER_CODE = {
        "itemcode", "code", "产品代码", "商品コード", "コード", "编码", "代码",
    }

    norm_headers = [c.strip().lower().replace(" ", "") for c in first_cols]
    for i, h in enumerate(norm_headers):
        if h in HEADER_NAME and name_col < 0:
            name_col = i
        elif h in HEADER_PRICE and price_col < 0:
            price_col = i
        elif h in HEADER_CAT and category_col < 0:
            category_col = i
        elif h in HEADER_CODE and code_col < 0:
            code_col = i

    if name_col >= 0 and price_col >= 0:
        header_row = 0

    if header_row is None:
        for i, h in enumerate(norm_headers):
            try:
                float(h.replace(",", ""))
                if price_col < 0:
                    price_col = i
            except ValueError:
                if name_col < 0 and len(h) > 0:
                    name_col = i

    if name_col < 0:
        return None

    data_start = 1 if header_row is not None else 0
    items = []
    counter = 1

    for line in lines[data_start:]:
        cols = line.split("\t")
        if len(cols) <= name_col:
            continue

        name = cols[name_col].strip()
        if not name:
            continue

        price = 0.0
        if price_col >= 0 and price_col < len(cols):
            price = _parse_price(cols[price_col])

        category = "Default"
        if category_col >= 0 and category_col < len(cols):
            cat = cols[category_col].strip()
            if cat:
                category = cat

        itemcode = "%04d" % counter
        if code_col >= 0 and code_col < len(cols):
            c = cols[code_col].strip()
            if c:
                itemcode = c

        items.append({
            "itemcode": itemcode,
            "name": name,
            "price": price,
            "category": category,
        })
        counter += 1

    return items if items else None


def parse_menu_text(text: str, source_lang: str = "auto") -> list[dict]:
    """
    Parse OCR/extracted text into structured menu items.
    Returns: [{"itemcode": "0001", "name": "...", "price": 1800, "category": "..."}]
    """
    tabular = _try_parse_tabular(text)
    if tabular is not None:
        return tabular

    lines = text.strip().split("\n")
    lines = [l.strip() for l in lines if l.strip()]

    if not lines:
        return []

    shared_price = detect_shared_price(text)
    shared_price_line_seen = False

    items = []
    current_category = "Default"
    item_counter = 1

    for i, line in enumerate(lines):
        raw = line.strip()
        had_bullet = bool(BULLET_PREFIX.match(raw))
        cleaned = BULLET_PREFIX.sub('', raw).strip()
        if not cleaned:
            continue

        if shared_price and not shared_price_line_seen and not had_bullet:
            _, lp = _extract_price(cleaned)
            if lp is not None and lp == shared_price:
                shared_price_line_seen = True
                continue

        if had_bullet:
            name, price = _extract_price(cleaned)
            if not name:
                continue
            if price is None and shared_price is not None:
                price = shared_price
            elif price is None:
                price = 0
            items.append({
                "itemcode": "%04d" % item_counter,
                "name": name,
                "price": price,
                "category": current_category,
            })
            item_counter += 1
            continue

        has_items_after = i < len(lines) - 1
        if _is_category_header(cleaned, has_items_after):
            cat = CATEGORY_HINT_CHARS.sub('', cleaned).strip()
            cat = cat.strip('=＝─━-*# ')
            if cat:
                current_category = cat
            continue

        name, price = _extract_price(cleaned)
        if not name:
            continue

        if price is None and shared_price is not None:
            price = shared_price
        elif price is None:
            price = 0

        items.append({
            "itemcode": "%04d" % item_counter,
            "name": name,
            "price": price,
            "category": current_category,
        })
        item_counter += 1

    return items
