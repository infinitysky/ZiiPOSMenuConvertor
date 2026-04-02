"""
Rule-based menu text parser.
Extracts structured menu items (itemcode, name, price, category) from OCR text.
Supports: plain text, tab-separated, HTML tables (from rapid-table), spatial boxes.
"""

import re
from html.parser import HTMLParser

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
    cleaned = text.replace(",", "").replace("，", "").strip()
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0


def _extract_price(line: str) -> tuple:
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


# ──────────────────────────────────────────────────────────────
# HTML table parser (for rapid-table output)
# ──────────────────────────────────────────────────────────────

class _TableHTMLParser(HTMLParser):
    """Minimal HTML table parser that extracts rows of cells."""

    def __init__(self):
        super().__init__()
        self.rows: list[list[str]] = []
        self._current_row: list[str] = []
        self._current_cell: list[str] = []
        self._in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._current_row = []
        elif tag in ("td", "th"):
            self._in_cell = True
            self._current_cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self._in_cell = False
            self._current_row.append("".join(self._current_cell).strip())
        elif tag == "tr":
            if self._current_row:
                self.rows.append(self._current_row)

    def handle_data(self, data):
        if self._in_cell:
            self._current_cell.append(data)


def parse_table_html(html: str) -> list[dict]:
    """
    Parse an HTML table (from rapid-table) into menu items.
    Tries to detect name/price/category columns by header or content heuristics.
    """
    parser = _TableHTMLParser()
    parser.feed(html)
    rows = parser.rows
    if not rows or len(rows) < 2:
        return []

    HEADER_NAME = {
        "name", "名称", "名前", "品名", "description", "description1",
        "名称1", "item", "itemname", "menu", "商品名", "品目", "メニュー",
    }
    HEADER_PRICE = {
        "price", "価格", "金額", "单价", "單價", "price1", "价格",
        "値段", "料金",
    }
    HEADER_CAT = {
        "category", "分类", "カテゴリ", "分類", "类别", "類別", "group",
    }

    name_col = -1
    price_col = -1
    category_col = -1
    data_start = 0

    norm_headers = [c.strip().lower().replace(" ", "") for c in rows[0]]
    for i, h in enumerate(norm_headers):
        if h in HEADER_NAME and name_col < 0:
            name_col = i
        elif h in HEADER_PRICE and price_col < 0:
            price_col = i
        elif h in HEADER_CAT and category_col < 0:
            category_col = i

    if name_col >= 0:
        data_start = 1
    else:
        num_cols = max(len(r) for r in rows)
        price_counts = [0] * num_cols
        text_lengths = [0] * num_cols

        for r in rows:
            for ci, cell in enumerate(r):
                try:
                    float(cell.replace(",", "").replace("，", "").strip())
                    price_counts[ci] += 1
                except ValueError:
                    text_lengths[ci] += len(cell)

        best_price = -1
        best_price_count = 0
        for ci in range(num_cols):
            if price_counts[ci] > best_price_count:
                best_price_count = price_counts[ci]
                best_price = ci

        best_name = -1
        best_text_len = 0
        for ci in range(num_cols):
            if ci != best_price and text_lengths[ci] > best_text_len:
                best_text_len = text_lengths[ci]
                best_name = ci

        if best_name >= 0:
            name_col = best_name
        if best_price >= 0 and best_price_count >= len(rows) * 0.3:
            price_col = best_price

    if name_col < 0:
        return []

    items = []
    counter = 1
    for row in rows[data_start:]:
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if not name:
            continue

        price = 0.0
        if 0 <= price_col < len(row):
            price = _parse_price(row[price_col])

        category = "Default"
        if 0 <= category_col < len(row):
            cat = row[category_col].strip()
            if cat:
                category = cat

        items.append({
            "itemcode": "%04d" % counter,
            "name": name,
            "price": price,
            "category": category,
        })
        counter += 1

    return items


# ──────────────────────────────────────────────────────────────
# Spatial box parser (name-price pairing using OCR bounding boxes)
# Multi-column aware, with pattern matching and bilingual merging.
# ──────────────────────────────────────────────────────────────

_BOX_PRICE_RE = re.compile(
    r'^[¥￥$＄€£₩฿]?\s*[0-9,]+(?:\.[0-9]{1,2})?\s*(?:円|元|/份|/个|/位)?$'
)
_BOX_PRICE_PAIR_RE = re.compile(
    r'[¥￥$＄]?\s*([0-9,]+(?:\.[0-9]{1,2})?)\s*[/／]\s*(?:份|个|位|会员)'
)
_ITEM_NUMBER_RE = re.compile(r'^(\d{1,3})\s+(.+)')
_MEMBER_TAG_RE = re.compile(r'【会员】|【會員】|\[会员\]|会员价|member', re.IGNORECASE)
_CATEGORY_DECO_RE = re.compile(
    r'[■□●○◆◇★☆·•▪▫►▸◄◂▲△▼▽※❖🔸🔹]'
    r'|^[=＝─━\-]{2,}'
)


def _is_cjk(text: str) -> bool:
    cjk = sum(1 for ch in text if 0x4E00 <= ord(ch) <= 0x9FFF
              or 0x3040 <= ord(ch) <= 0x30FF
              or 0x3400 <= ord(ch) <= 0x4DBF)
    return cjk > len(text) * 0.3


def _is_latin(text: str) -> bool:
    latin = sum(1 for ch in text if ch.isascii() and ch.isalpha())
    return latin > len(text) * 0.5


def _classify_box(text: str) -> str:
    """Classify a text block: 'price', 'number', 'category', 'name', 'member_tag'."""
    stripped = text.strip()
    if not stripped:
        return "empty"
    if _MEMBER_TAG_RE.search(stripped):
        return "member_tag"
    if _BOX_PRICE_RE.match(stripped):
        return "price"
    clean = re.sub(r'[¥￥$＄€£₩฿]', '', stripped).replace(",", "").strip()
    if _BOX_PRICE_PAIR_RE.search(stripped):
        return "price"
    try:
        v = float(clean)
        if 0.5 <= v <= 99999:
            return "price"
    except ValueError:
        pass
    if re.match(r'^\d{1,3}$', stripped):
        return "number"
    return "name"


def _extract_first_price(text: str) -> float:
    """Extract the first numeric price from a text block."""
    text = re.sub(r'【.*?】|\[.*?\]', '', text)
    text = re.sub(r'[¥￥$＄€£₩฿]', '', text)
    m = re.search(r'([0-9,]+\.[0-9]{1,2})', text)
    if m:
        return _parse_price(m.group(1))
    m = re.search(r'([0-9]{2,})', text)
    if m:
        return _parse_price(m.group(1))
    return 0.0


def _is_box_category_header(text: str, box_height: float, avg_height: float) -> bool:
    """Detect if a text block is a category/section header."""
    stripped = text.strip()
    if not stripped:
        return False
    if _CATEGORY_DECO_RE.search(stripped):
        return True
    if '·' in stripped or '•' in stripped:
        parts = re.split(r'[·•]', stripped)
        if len(parts) >= 2 and all(len(p.strip()) > 0 for p in parts[:2]):
            return True
    has_cn = _is_cjk(stripped)
    has_en_upper = bool(re.search(r'[A-Z]{3,}', stripped))
    if has_cn and has_en_upper:
        return True
    if box_height > avg_height * 1.3 and len(stripped) < 25:
        return True
    return False


def _detect_columns(entries: list[dict], img_width: float) -> list[list[dict]]:
    """
    Detect columns by clustering x_left of NON-PRICE text blocks
    (prices are always right-aligned within a column, so they'd create
    false splits if included). Then assign all boxes to nearest column.
    """
    if not entries or img_width <= 0:
        return [entries]

    name_entries = [e for e in entries if _classify_box(e["text"]) != "price"]
    if len(name_entries) < 4:
        return [entries]

    x_lefts = sorted(e["x_left"] for e in name_entries)

    gaps = []
    for i in range(1, len(x_lefts)):
        gaps.append((x_lefts[i] - x_lefts[i - 1], i))
    gaps.sort(reverse=True)

    min_gap = img_width * 0.10

    split_points = []
    for gap_size, idx in gaps:
        if gap_size < min_gap:
            break
        boundary = (x_lefts[idx - 1] + x_lefts[idx]) / 2
        too_close = False
        for sp in split_points:
            if abs(boundary - sp) < img_width * 0.08:
                too_close = True
                break
        if not too_close:
            split_points.append(boundary)
        if len(split_points) >= 3:
            break

    if not split_points:
        return [entries]

    split_points.sort()
    boundaries = [0] + split_points + [img_width + 1]

    col_ranges = []
    for ci in range(len(boundaries) - 1):
        col_entries = [e for e in name_entries
                       if boundaries[ci] <= e["x_left"] < boundaries[ci + 1]]
        if col_entries:
            max_right = max(e["x_right"] for e in col_entries)
            col_ranges.append((ci, boundaries[ci], max_right))

    full_boundaries = [0]
    for i in range(len(col_ranges) - 1):
        mid = (col_ranges[i][2] + col_ranges[i + 1][1]) / 2
        full_boundaries.append(mid)
    full_boundaries.append(img_width + 1)

    columns: list[list[dict]] = [[] for _ in range(len(col_ranges))]
    for e in entries:
        ex = e["x_left"]
        placed = False
        for ci in range(len(full_boundaries) - 1):
            if full_boundaries[ci] <= ex < full_boundaries[ci + 1]:
                if ci < len(columns):
                    columns[ci].append(e)
                placed = True
                break
        if not placed and columns:
            columns[-1].append(e)

    columns = [c for c in columns if len(c) >= 2]
    for c in columns:
        c.sort(key=lambda e: (e["y"], e["x_left"]))

    return columns if columns else [entries]


def _group_into_rows(entries: list[dict], y_threshold: float) -> list[list[dict]]:
    """Group entries into rows by Y-coordinate proximity."""
    if not entries:
        return []
    entries_sorted = sorted(entries, key=lambda e: (e["y"], e["x_left"]))
    rows: list[list[dict]] = []
    current_row = [entries_sorted[0]]
    for e in entries_sorted[1:]:
        row_y = sum(r["y"] for r in current_row) / len(current_row)
        if abs(e["y"] - row_y) < y_threshold:
            current_row.append(e)
        else:
            rows.append(sorted(current_row, key=lambda r: r["x_left"]))
            current_row = [e]
    rows.append(sorted(current_row, key=lambda r: r["x_left"]))
    return rows


def _parse_column_rows(rows: list[list[dict]], avg_height: float) -> list[dict]:
    """
    Parse rows within a single column into structured menu items.
    Handles: item numbers, name-price pairing, category headers, bilingual merging.
    """
    raw_items: list[dict] = []
    current_category = "Default"
    i = 0

    while i < len(rows):
        row = rows[i]
        classified = [(e, _classify_box(e["text"])) for e in row]

        names = [(e, c) for e, c in classified if c == "name"]
        prices = [(e, c) for e, c in classified if c == "price"]
        numbers = [(e, c) for e, c in classified if c == "number"]

        all_text = " ".join(e["text"] for e in row)

        if not prices and not numbers and len(names) <= 2:
            combined = " ".join(e["text"].strip() for e, _ in names)
            h = max((e["box_h"] for e in row), default=avg_height)
            if _is_box_category_header(combined, h, avg_height):
                cat = re.sub(r'^[■□●○◆◇★☆·•▪▫]+\s*', '', combined).strip()
                cat = re.sub(r'\s+(COLD DISHES|HOT DISHES|APPETIZER|SOUPS?|SEAFOOD|'
                             r'STAPLES|DESSERTS?|BEVERAGES?|RECOMMENDED DISHES|'
                             r'HOT BEVERAGES|TRADITIONAL BEIJING ROAST DUCK)$',
                             '', cat, flags=re.IGNORECASE).strip()
                if cat:
                    current_category = cat
                i += 1
                continue

        if not names and not prices:
            i += 1
            continue

        item_name = ""
        item_name_alt = ""
        item_price = 0.0
        item_number = ""

        if numbers:
            item_number = numbers[0][0]["text"].strip()

        for e, c in names:
            txt = e["text"].strip()
            m = _ITEM_NUMBER_RE.match(txt)
            if m:
                if not item_number:
                    item_number = m.group(1)
                txt = m.group(2).strip()
            if _is_cjk(txt):
                if item_name:
                    item_name += " " + txt
                else:
                    item_name = txt
            elif _is_latin(txt):
                if item_name_alt:
                    item_name_alt += " " + txt
                else:
                    item_name_alt = txt
            else:
                if item_name:
                    item_name += " " + txt
                else:
                    item_name = txt

        if prices:
            price_vals = []
            for e, _ in prices:
                p = _extract_first_price(e["text"])
                if p > 0:
                    price_vals.append(p)
            if price_vals:
                item_price = max(price_vals)

        if not item_name and item_name_alt:
            item_name = item_name_alt
            item_name_alt = ""

        if not item_name:
            i += 1
            continue

        if i + 1 < len(rows):
            next_row = rows[i + 1]
            next_classified = [(e, _classify_box(e["text"])) for e in next_row]
            next_names = [(e, c) for e, c in next_classified if c == "name"]
            next_prices = [(e, c) for e, c in next_classified if c == "price"]

            if next_names and not next_prices:
                next_text = " ".join(e["text"].strip() for e, _ in next_names)
                next_h = max((e["box_h"] for e in next_row), default=avg_height)
                if not _is_box_category_header(next_text, next_h, avg_height):
                    if _is_cjk(item_name) and _is_latin(next_text):
                        item_name_alt = next_text
                        i += 1
                    elif _is_latin(item_name) and _is_cjk(next_text):
                        item_name_alt = item_name
                        item_name = next_text
                        i += 1

        raw_items.append({
            "name": item_name,
            "name_alt": item_name_alt,
            "price": item_price,
            "category": current_category,
            "number": item_number,
        })
        i += 1

    return raw_items


def parse_menu_from_boxes(box_data: list[dict]) -> list[dict]:
    """
    Parse menu items from OCR bounding boxes with:
    - Multi-column detection
    - Row grouping per column
    - Pattern-based classification (number, name, price, category)
    - Bilingual name merging (CJK + Latin on adjacent rows)

    box_data: list of {'text': str, 'box': [[x1,y1],...], 'score': float}
    """
    if not box_data:
        return []

    entries = []
    heights = []
    max_x = 0
    for d in box_data:
        box = d.get("box", [])
        if not box or len(box) < 4:
            continue
        y_center = sum(pt[1] for pt in box) / 4
        x_left = min(pt[0] for pt in box)
        x_right = max(pt[0] for pt in box)
        box_h = abs(box[2][1] - box[0][1])
        text = d["text"].strip()
        if not text:
            continue
        entries.append({
            "text": text,
            "y": y_center,
            "x_left": x_left,
            "x_right": x_right,
            "box_h": box_h,
            "score": d.get("score", 0),
        })
        heights.append(box_h)
        if x_right > max_x:
            max_x = x_right

    if not entries:
        return []

    avg_height = sum(heights) / len(heights)
    y_threshold = avg_height * 0.7

    columns = _detect_columns(entries, max_x)

    all_raw_items: list[dict] = []
    for col_entries in columns:
        rows = _group_into_rows(col_entries, y_threshold)
        raw_items = _parse_column_rows(rows, avg_height)
        all_raw_items.extend(raw_items)

    items = []
    counter = 1
    for raw in all_raw_items:
        items.append({
            "itemcode": "%04d" % counter,
            "name": raw["name"],
            "name_alt": raw.get("name_alt", ""),
            "price": raw["price"],
            "category": raw["category"],
        })
        counter += 1

    return items
