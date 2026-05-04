import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import re
import io
from datetime import datetime

# ─── PAGE CONFIG ────────────────────────────────────────────
st.set_page_config(
    page_title="Pro Standard — PO Generator",
    page_icon="🏆",
    layout="wide"
)

# ─── STYLES ─────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
    
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    
    .main-title {
        font-family: 'Bebas Neue', sans-serif;
        font-size: 3rem;
        letter-spacing: 4px;
        color: #e8c84a;
        margin-bottom: 0;
    }
    .sub-title {
        font-family: 'Bebas Neue', sans-serif;
        font-size: 1.2rem;
        letter-spacing: 3px;
        color: #888;
        margin-top: 0;
    }
    .section-title {
        font-family: 'Bebas Neue', sans-serif;
        font-size: 1.1rem;
        letter-spacing: 2px;
        color: #aaa;
    }
    .field-box {
        background: #1a1a1a;
        border-radius: 6px;
        padding: 10px 14px;
        margin: 4px 0;
    }
    .field-label { font-size: 0.7rem; letter-spacing: 1.5px; text-transform: uppercase; color: #666; }
    .field-value { font-size: 0.95rem; color: #f5f5f0; }
    .warn-box {
        background: rgba(255,165,0,0.1);
        border: 1px solid rgba(255,165,0,0.4);
        border-radius: 6px;
        padding: 10px 14px;
        font-size: 0.85rem;
        color: #ffa500;
    }
    .success-box {
        background: rgba(74,222,128,0.1);
        border: 1px solid rgba(74,222,128,0.4);
        border-radius: 6px;
        padding: 10px 14px;
        font-size: 0.85rem;
        color: #4ade80;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #444;
        border-radius: 8px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ─── GLOSARIO ───────────────────────────────────────────────
GLOSSARY = {
    "po_number": [
        "PO#", "PO NUMBER", "PO NO", "PURCHASE ORDER NUMBER",
        "PURCHASE ORDER", "ORDER NUMBER", "ORDER #", "P.O. NUMBER", "P.O.#"
    ],
    "po_date": [
        "PO DATE", "ORDER DATE", "PO CREATION DATE", "CREATION DATE",
        "ISSUE DATE", "CREATED ON", "DATE", "START DATE"
    ],
    "ship_date": [
        "SHIP DATE", "START SHIP", "SHIP WINDOW", "DELIVERY DATE",
        "IN HANDS DATE", "IN-HANDS DATE", "IN DC DATE", "DISPATCH DATE",
        "SHIP BY", "DELIVER BY", "NEED BY DATE", "ON FLOOR"
    ],
    "cancel_date": [
        "CANCEL DATE", "CANCEL BY", "CANCELLATION DATE", "TERMINATION DATE",
        "DO NOT SHIP AFTER", "DNSA", "EXPIRATION DATE", "LATEST SHIP DATE",
        "FECHA LIMITE", "CANCEL"
    ],
    "customer_name": [
        "CUSTOMER NAME", "SOLD TO", "ACCOUNT NAME", "BILL TO", "BILL-TO",
        "BILLED TO", "CLIENT", "BUYER", "ACCOUNT", "COMPANY"
    ],
    "customer_code": [
        "CUSTOMER CODE", "ACCOUNT NUMBER", "ACCOUNT #", "CUSTOMER ID",
        "BUYER CODE", "ACCOUNT NO", "CLIENT CODE"
    ],
    "ship_to": [
        "SHIP TO", "SHIP-TO", "SHIPPING ADDRESS", "DELIVER TO",
        "DELIVERY ADDRESS", "CONSIGNEE", "DC ADDRESS", "STORE ADDRESS"
    ],
    "bill_to": [
        "BILL TO", "BILL-TO", "INVOICE ADDRESS", "ACCOUNTS PAYABLE",
        "AP ADDRESS", "BILLING ADDRESS"
    ],
    "terms": [
        "TERMS", "PAYMENT TERMS", "NET TERMS", "NET"
    ],
    # Product columns
    "col_style": [
        "VENDOR STYLE", "STYLE ID", "STYLE #", "STYLE NO", "STYLE NUMBER",
        "STOCK", "ITEM #", "ITEM NUMBER", "ITEM NO", "SKU", "SKU#",
        "PRODUCT CODE", "STYLE", "VENDOR ITEM", "ARTICLE"
    ],
    "col_desc": [
        "ITEM DESCRIPTION", "ITEM DESC", "DESCRIPTION", "DESC",
        "PRODUCT NAME", "PRODUCT DESCRIPTION", "ARTICLE DESCRIPTION"
    ],
    "col_color": [
        "COLOR NAME", "COLOR", "COLOUR", "COLOR STORY", "COLORWAY"
    ],
    "col_size_break": [
        "SIZE SCALE", "SIZE BREAK", "SIZE RUN", "SIZES", "SIZE DISTRIBUTION",
        "PACK RATIO", "SIZE CURVE"
    ],
    "col_qty": [
        "QTY", "QUANTITY", "PURCH QTY", "ORDER QTY", "UNITS",
        "# UNITS", "PIECES", "PCS", "TOTAL QTY", "TOTAL UNITS", "TTL UNITS"
    ],
    "col_cost": [
        "COST", "UNIT COST", "UNIT PRICE", "PRICE PER UNIT", "WHOLESALE",
        "WS", "PRICE", "NET COST", "COST WITH DISCOUNT"
    ],
    "col_msrp": [
        "MSRP", "RETAIL", "RETAIL PRICE", "SRP", "SUGGESTED RETAIL",
        "SELL PRICE", "TAG PRICE"
    ],
    "col_total_cost": [
        "TOTAL COST", "EXT COST", "EXTENDED COST", "LINE TOTAL",
        "TOTAL DOLLARS", "SUBTOTAL"
    ],
    "col_total_retail": [
        "TOTAL RETAIL", "EXT RETAIL", "EXTENDED RETAIL", "RETAIL VALUE"
    ],
    "col_discount": [
        "DISCOUNT", "DISC", "DISC %", "DISCOUNT %", "OFF %"
    ],
    "col_upc": [
        "UPC", "UPC CODE", "BARCODE", "GTIN", "EAN"
    ],
}

SIZE_ORDER = ["OS", "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]

SIZE_ALIASES = {
    "OSFA": "OS", "OSFM": "OS", "ONE SIZE": "OS", "O/S": "OS",
    "XXS": "XXS", "2XS": "XXS",
    "XS": "XS",
    "S": "S", "SM": "S", "SML": "S", "S/P": "S",
    "M": "M", "MD": "M", "MED": "M", "M/M": "M",
    "L": "L", "LG": "L", "LRG": "L", "L/G": "L",
    "XL": "XL", "X-LARGE": "XL", "X/L": "XL",
    "XXL": "2XL", "2XL": "2XL", "2X": "2XL", "XX": "2XL", "2X-LARGE": "2XL",
    "XXXL": "3XL", "3XL": "3XL", "3X": "3XL", "3X-LARGE": "3XL",
}

COLOR_CODES = {
    "BLK": "BLACK", "WHT": "WHITE", "NVY": "NAVY", "RED": "RED",
    "BLU": "BLUE", "GRY": "GREY", "GRN": "GREEN", "GLD": "GOLD",
    "WBK": "WASHED BLACK", "EGG": "EGGSHELL", "CRM": "CREAM",
    "HGR": "HEATHER GREY", "DHG": "DARK HEATHER GREY", "LGY": "LIGHT GREY",
    "MDN": "MIDNIGHT NAVY", "RYB": "ROYAL BLUE", "PUR": "PURPLE",
    "WNE": "WINE", "TEL": "TEAL", "FOR": "FOREST", "KGN": "KELLY GREEN",
    "CRD": "CRIMSON RED", "NSF": "NIGHT SKY FADE", "DBL": "DODGER BLUE",
    "CNC": "CHARCOAL",
}

# ─── HELPER FUNCTIONS ────────────────────────────────────────

def fmtDate(v):
    """Format date value to MM/DD/YYYY string."""
    if not v:
        return ""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime('%m/%d/%Y')
    s = str(v).strip()
    # Clean double slashes
    s = re.sub(r'/+', '/', s)
    return s


def cell_str(v):
    return str(v).upper().strip() if v is not None else ""

def matches_any(text, keywords):
    text = text.upper().strip()
    return any(text == kw or kw in text for kw in keywords)

def find_header_value(rows, keywords, max_rows=25):
    """Search first max_rows for a label matching keywords, return adjacent value."""
    for i, row in enumerate(rows[:max_rows]):
        for j, cell in enumerate(row):
            c = cell_str(cell)
            if matches_any(c, keywords):
                # Try value after colon in same cell
                raw = str(cell).strip()
                if ":" in raw:
                    val = raw.split(":", 1)[1].strip()
                    if val and val != "0":
                        return val
                # Try next cells in same row
                for k in range(j+1, min(j+4, len(row))):
                    val = str(row[k]).strip()
                    if val and val not in ("", "0", "None"):
                        return val
    return ""

def normalize_size(s):
    s = str(s).upper().strip()
    return SIZE_ALIASES.get(s, s)

def parse_size_break(cell_value):
    """
    Parse compressed size break like 'S - 3X   18/30/42/30/18/6'
    Splits on 2+ spaces to separate range from quantities.
    Returns dict {S:18, M:30, L:42, XL:30, 2XL:18, 3XL:6}
    """
    sizes = {s: 0 for s in SIZE_ORDER}
    if not cell_value:
        return sizes

    val = str(cell_value).strip()

    # Split range from quantities on 2+ spaces or tab
    parts = re.split(r'\s{2,}|\t', val, maxsplit=1)
    range_part   = parts[0].strip()
    numbers_part = parts[1].strip() if len(parts) > 1 else val

    # Parse size range
    range_match = re.match(r'([A-Z0-9]+)\s*[-]+\s*([A-Z0-9]+)', range_part.upper())
    if range_match:
        start_size = normalize_size(range_match.group(1))
        end_size   = normalize_size(range_match.group(2))
        try:
            start_idx  = SIZE_ORDER.index(start_size)
            end_idx    = SIZE_ORDER.index(end_size)
            size_slice = SIZE_ORDER[start_idx:end_idx+1]
        except ValueError:
            size_slice = ["S","M","L","XL","2XL","3XL"]
    else:
        size_slice = ["S","M","L","XL","2XL","3XL"]

    # Parse quantities from numbers_part only (avoids mixing range digits)
    numbers = [int(n) for n in re.findall(r'\d+', numbers_part)]
    for i, sz in enumerate(size_slice):
        if i < len(numbers):
            sizes[sz] = numbers[i]

    return sizes

def parse_style_color(style_code):
    """
    Split 'FSS1411937-WBK' into stock='FSS1411937', color_code='WBK'
    """
    style_code = str(style_code).strip()
    if "-" in style_code:
        parts = style_code.rsplit("-", 1)
        stock = parts[0].strip()
        color_code = parts[1].strip().upper()
        color_name = COLOR_CODES.get(color_code, color_code)
        return stock, color_code, color_name
    return style_code, "", ""

def detect_col(header_row, keywords):
    """Find column index matching keywords in header row."""
    for i, cell in enumerate(header_row):
        if matches_any(cell_str(cell), keywords):
            return i
    return None

# ─── PARSE EXCEL PO ─────────────────────────────────────────

def parse_po_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    result = {
        "po_number": "", "po_date": "", "ship_date": "", "cancel_date": "",
        "customer_name": "", "customer_code": "", "ship_to": "", "bill_to": "",
        "terms": "", "currency": "USD", "lines": [], "warnings": []
    }

    # ── Extract header fields ──

    # PO Number — try standard lookup first, then special patterns
    po_number = find_header_value(rows, GLOSSARY["po_number"])
    if not po_number:
        # Pattern 1: label "PO NAME" / "PO NUMBER" with value in adjacent cell
        for row in rows[:5]:
            for j, cell in enumerate(row):
                c = cell_str(cell)
                if c in ("PO NAME", "PO NUMBER", "PO#"):
                    # Value is in next cell on same row OR last non-empty cell in row
                    for k in range(j+1, len(row)):
                        v = str(row[k] or "").strip()
                        if v and v not in ("", "None"):
                            po_number = v
                            break
                    if po_number:
                        break
                # Pattern 2: "PO# 26064  Pro Standard" embedded in text
                if cell and "PO" in str(cell).upper() and not po_number:
                    m = re.search(r'PO#?\s*([A-Z0-9][A-Z0-9\s\-]+)', str(cell), re.IGNORECASE)
                    if m:
                        candidate = m.group(1).strip()
                        if len(candidate) > 2 and candidate.upper() not in ("NAME","NUMBER","DATE"):
                            po_number = candidate
            if po_number:
                break
    result["po_number"] = po_number

    # PO Date
    po_date_raw = find_header_value(rows, GLOSSARY["po_date"])
    result["po_date"] = fmtDate(po_date_raw)

    # Ship Date — clean double slashes e.g. "9/15//2026" → "9/15/2026"
    ship_date_raw = find_header_value(rows, GLOSSARY["ship_date"])
    result["ship_date"] = re.sub(r'/+', '/', str(ship_date_raw)) if ship_date_raw else ""

    # Cancel Date
    cancel_raw = find_header_value(rows, GLOSSARY["cancel_date"])
    result["cancel_date"] = fmtDate(cancel_raw)

    # Customer Name — try standard, then look for "BILL TO:" label and grab value below/beside
    customer_name = find_header_value(rows, GLOSSARY["customer_name"])
    if not customer_name:
        for i, row in enumerate(rows[:10]):
            for j, cell in enumerate(row):
                c = cell_str(cell)
                if c in ("BILL TO:", "BILL TO", "SOLD TO:", "SOLD TO"):
                    # Value is in same cell after colon, or next cell, or next row same col
                    raw = str(cell).strip()
                    if ":" in raw:
                        after = raw.split(":",1)[1].strip()
                        if after and len(after) > 2:
                            customer_name = after
                            break
                    # Try next row, same column
                    if i+1 < len(rows) and j < len(rows[i+1]):
                        v = str(rows[i+1][j] or "").strip()
                        if v and len(v) > 2:
                            customer_name = v
                            break
            if customer_name:
                break
    result["customer_name"] = customer_name
    result["customer_code"] = find_header_value(rows, GLOSSARY["customer_code"])

    # Ship To — try standard, then look for "SHIP TO:" and collect lines below
    ship_to = find_header_value(rows, GLOSSARY["ship_to"])
    if not ship_to:
        for i, row in enumerate(rows[:10]):
            for j, cell in enumerate(row):
                c = cell_str(cell)
                if c in ("SHIP TO:", "SHIP TO", "SHIPPING ADDRESS:", "DELIVERY ADDRESS:"):
                    # Collect next 3 non-empty rows at same column (stop at product table)
                    parts = []
                    for k in range(i+1, min(i+5, len(rows))):
                        if j < len(rows[k]):
                            v = str(rows[k][j] or "").strip()
                            # Stop if we hit a product table header keyword
                            if v and v.upper() in ("DESCRIPTION","VENDOR STYLE NO.","STYLE","QTY","UPC"):
                                break
                            if v:
                                parts.append(v)
                    if parts:
                        ship_to = " | ".join(parts)
                    break
            if ship_to:
                break
    result["ship_to"] = ship_to

    result["bill_to"] = find_header_value(rows, GLOSSARY["bill_to"])
    result["terms"]   = find_header_value(rows, GLOSSARY["terms"])

    # ── Find product table header row ──
    data_header_row = None
    col_map = {}

    for i, row in enumerate(rows):
        row_str = " | ".join(cell_str(c) for c in row)
        has_style = matches_any(row_str, GLOSSARY["col_style"])
        has_qty   = matches_any(row_str, GLOSSARY["col_qty"]) or matches_any(row_str, GLOSSARY["col_size_break"])
        if has_style and has_qty:
            data_header_row = i
            col_map["style"]       = detect_col(row, GLOSSARY["col_style"])
            col_map["desc"]        = detect_col(row, GLOSSARY["col_desc"])
            col_map["color"]       = detect_col(row, GLOSSARY["col_color"])
            col_map["size_break"]  = detect_col(row, GLOSSARY["col_size_break"])
            col_map["size"]        = detect_col(row, ["SIZE", "SIZES"])
            col_map["qty"]         = detect_col(row, GLOSSARY["col_qty"])
            col_map["cost"]        = detect_col(row, GLOSSARY["col_cost"])
            col_map["msrp"]        = detect_col(row, GLOSSARY["col_msrp"])
            col_map["discount"]    = detect_col(row, GLOSSARY["col_discount"])
            col_map["total_cost"]  = detect_col(row, GLOSSARY["col_total_cost"])
            col_map["total_retail"]= detect_col(row, GLOSSARY["col_total_retail"])
            break

    if data_header_row is None:
        result["warnings"].append("⚠️ No se encontró la tabla de productos.")
        return result

    # ── Read product rows ──
    lines = []
    last_valid_style = ""
    last_ship_date = result["ship_date"]

    for row in rows[data_header_row + 1:]:
        # Check if row has meaningful data
        if not any(c for c in row if c is not None and str(c).strip()):
            continue

        # Check for ship date embedded in rows (like "Ship Date: 10/1/2026-10/15/2026")
        row_text = " ".join(cell_str(c) for c in row if c)
        if matches_any(row_text, GLOSSARY["ship_date"]) and not result["ship_date"]:
            for c in row:
                if c and matches_any(cell_str(c), GLOSSARY["ship_date"]):
                    raw = str(c).strip()
                    if ":" in raw:
                        val = raw.split(":", 1)[1].strip()
                        if val:
                            result["ship_date"] = val
                            break
            continue

        # Get style — carry forward last valid style for row-per-size format
        style_raw = ""
        if col_map.get("style") is not None and col_map["style"] < len(row):
            style_raw = str(row[col_map["style"]] or "").strip()

        if style_raw:
            # Skip total/summary rows
            if matches_any(style_raw, ["TOTAL", "SUBTOTAL", "GRAND TOTAL"]):
                continue
            # Must look like a real style code: letters+digits
            if not re.match(r'^[A-Za-z]{2,}\d+', style_raw):
                continue
            # Valid — update carry-forward
            last_valid_style = style_raw
        else:
            # Empty style cell — use carry-forward if we have a size value (row-per-size)
            if last_valid_style and col_map.get("size") is not None:
                style_raw = last_valid_style
            else:
                continue

        # Parse style + color
        stock, color_code, color_name = parse_style_color(style_raw)

        # Description
        desc = ""
        if col_map.get("desc") is not None and col_map["desc"] < len(row):
            desc = str(row[col_map["desc"]] or "").strip()

        # Color override from dedicated column
        if col_map.get("color") is not None and col_map["color"] < len(row):
            color_val = str(row[col_map["color"]] or "").strip()
            if color_val:
                color_name = color_val

        # Cost
        cost = 0.0
        if col_map.get("cost") is not None and col_map["cost"] < len(row):
            try:
                cost = float(row[col_map["cost"]] or 0)
            except (ValueError, TypeError):
                cost = 0.0

        # MSRP
        msrp = 0.0
        if col_map.get("msrp") is not None and col_map["msrp"] < len(row):
            try:
                msrp = float(row[col_map["msrp"]] or 0)
            except (ValueError, TypeError):
                msrp = 0.0

        # Discount — sometimes stored as decimal (0.525) sometimes as % (52.5)
        # Also check last non-empty cell if no explicit discount column
        discount = 0.0
        disc_col = col_map.get("discount")
        if disc_col is None:
            # Check last column for decimal values that look like discounts
            for idx in range(len(row)-1, -1, -1):
                v = row[idx]
                if v is not None:
                    try:
                        d = float(v)
                        if 0 < d <= 1:
                            disc_col = idx
                    except (ValueError, TypeError):
                        pass
                    break
        if disc_col is not None and disc_col < len(row):
            try:
                d = float(row[disc_col] or 0)
                discount = d if d <= 1 else d / 100
            except (ValueError, TypeError):
                discount = 0.0

        # Sizes
        sizes = {s: 0 for s in SIZE_ORDER}
        total_units = 0
        size_val = ""
        qty_val  = 0

        if col_map.get("size_break") is not None and col_map["size_break"] < len(row):
            # Compressed size break in one cell e.g. "S - 3X   18/30/42/30/18/6"
            sb_val = row[col_map["size_break"]]
            if sb_val:
                sizes = parse_size_break(sb_val)
                total_units = sum(sizes.values())
        elif col_map.get("size") is not None and col_map.get("qty") is not None:
            # One row per size — just record size+qty, group later
            size_val = str(row[col_map["size"]] or "").strip()
            try:
                qty_val = int(float(row[col_map["qty"]] or 0))
            except (ValueError, TypeError):
                qty_val = 0
            # Handle "OSFM / BLU" combined size+color
            if "/" in size_val:
                parts = size_val.split("/")
                size_val = parts[0].strip()
                if not color_code:
                    color_code = parts[1].strip()[:3]
            norm_size = normalize_size(size_val) if size_val else "OS"
            sizes[norm_size] = qty_val
            total_units = qty_val
        elif col_map.get("qty") is not None and col_map["qty"] < len(row):
            try:
                total_units = int(float(row[col_map["qty"]] or 0))
                sizes["OS"] = total_units
            except (ValueError, TypeError):
                pass

        if total_units == 0 and not desc:
            continue

        # In row-per-size format, skip rows with no size value (likely total rows)
        if col_map.get("size") is not None and not size_val and not col_map.get("size_break"):
            continue

        line_cost    = cost
        total_cost   = line_cost * total_units
        total_retail = msrp * total_units

        # Check if we can merge with previous line (same style+color, row-per-size format)
        key = stock + "|" + color_code
        merged = False
        if col_map.get("size") is not None and lines and size_val:
            for existing in reversed(lines):
                if existing["stock"] == stock and existing["color_code"] == color_code:
                    # Merge sizes
                    norm_size = normalize_size(size_val) if size_val else "OS"
                    existing["sizes"][norm_size] = existing["sizes"].get(norm_size, 0) + qty_val
                    existing["total_units"] += qty_val
                    existing["total_cost"]   = existing["line_cost"] * existing["total_units"]
                    existing["total_retail"] = existing["msrp"] * existing["total_units"]
                    merged = True
                    break

        if not merged:
            lines.append({
                "style_raw":    style_raw,
                "stock":        stock,
                "color_code":   color_code,
                "color_name":   color_name,
                "description":  desc,
                "cost":         cost,
                "msrp":         msrp,
                "discount":     discount,
                "line_cost":    line_cost,
                "sizes":        sizes,
                "total_units":  total_units,
                "total_cost":   total_cost,
                "total_retail": total_retail,
            })

    result["lines"] = lines
    return result

# ─── GENERATE ORDER FORM EXCEL ───────────────────────────────

def generate_order_form(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ORDER FORM"

    # Colors
    YELLOW = PatternFill("solid", fgColor="FFD700")
    HEADER_FILL = PatternFill("solid", fgColor="1A1A1A")
    ACCENT_FILL = PatternFill("solid", fgColor="2A2A2A")

    header_font   = Font(name="Calibri", bold=True, size=9)
    label_font    = Font(name="Calibri", bold=True, size=9, color="888888")
    value_font    = Font(name="Calibri", size=9)
    col_hdr_font  = Font(name="Calibri", bold=True, size=8, color="FFFFFF")

    totU = sum(l["total_units"] for l in data["lines"])
    totC = sum(l["total_cost"]  for l in data["lines"])
    totR = sum(l["total_retail"] for l in data["lines"])

    # ── Header block ──
    def write_label_value(row, col_label, label, col_value, value):
        c = ws.cell(row=row, column=col_label, value=label)
        c.font = label_font
        v = ws.cell(row=row, column=col_value, value=value)
        v.font = value_font

    write_label_value(1, 1, "Customer code:", 2, data.get("customer_code",""))
    ws.cell(row=1, column=5, value="Vendor:").font = label_font
    ws.cell(row=1, column=6, value="Maxima Apparel Corp SAPI de CV").font = value_font
    ws.cell(row=1, column=8, value="Brand:").font = label_font
    ws.cell(row=1, column=9, value="PRO STANDARD").font = Font(name="Calibri", bold=True, size=9)

    write_label_value(3, 1, "Customer name:", 2, data.get("customer_name",""))
    write_label_value(3, 5, "Ship to:", 6, data.get("ship_to",""))

    write_label_value(5, 1, "PO Number:", 2, data.get("po_number",""))
    write_label_value(6, 1, "PO Creation Date:", 2, data.get("po_date",""))

    write_label_value(8, 1,  "Ship Date:",   2, data.get("ship_date",""))
    write_label_value(8, 5,  "Terms:",       6, data.get("terms",""))
    write_label_value(9, 1,  "Cancel Date:", 2, data.get("cancel_date",""))
    ws.cell(row=9, column=5, value="Total cost:").font = label_font
    ws.cell(row=9, column=6, value=round(totC, 2)).font = value_font

    ws.cell(row=10, column=5, value="Total qty:").font  = label_font
    ws.cell(row=10, column=6, value=totU).font = value_font

    write_label_value(11, 1, "Buyer:", 2, data.get("buyer",""))
    ws.cell(row=11, column=5, value="Currency:").font = label_font
    ws.cell(row=11, column=6, value=data.get("currency","USD")).font = value_font

    # ── Column headers (row 13) ──
    COLS = [
        "BRAND","STOCK","COLOR CODE","COLOR NAME","DESCRIPTION","CURRENCY",
        "LINE COST","MSRP","OS","XXS","XS","S","M","L","XL","2XL","3XL",
        "TOTAL UNITS","TOTAL COST","TOTAL RETAIL"
    ]
    for j, col_name in enumerate(COLS, start=1):
        c = ws.cell(row=13, column=j, value=col_name)
        c.font = col_hdr_font
        c.fill = PatternFill("solid", fgColor="222222")
        c.alignment = Alignment(horizontal="center")

    # ── Data rows ──
    for i, line in enumerate(data["lines"], start=14):
        s = line["sizes"]
        row_data = [
            "PRO STANDARD",
            line["stock"],
            line["color_code"],
            line["color_name"],
            line["description"],
            data.get("currency","USD"),
            round(line["line_cost"], 2),
            round(line["msrp"], 2),
            s.get("OS",0), s.get("XXS",0), s.get("XS",0),
            s.get("S",0), s.get("M",0), s.get("L",0),
            s.get("XL",0), s.get("2XL",0), s.get("3XL",0),
            line["total_units"],
            round(line["total_cost"], 2),
            round(line["total_retail"], 2),
        ]
        for j, val in enumerate(row_data, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.font = value_font
            c.alignment = Alignment(horizontal="center" if j > 5 else "left")

    # ── Column widths ──
    widths = [14, 18, 11, 14, 42, 9, 10, 8, 6, 6, 6, 6, 6, 6, 6, 6, 6, 12, 12, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─── MAIN APP ────────────────────────────────────────────────

st.markdown('<p class="main-title">PURCHASE ORDER</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">PRO STANDARD — GENERATOR</p>', unsafe_allow_html=True)
st.divider()

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown('<p class="section-title">📋 Orden del Cliente</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Arrastra o selecciona el archivo",
        type=["xlsx", "xls"],
        label_visibility="collapsed"
    )

if uploaded_file:
    with st.spinner("Leyendo archivo..."):
        file_bytes = uploaded_file.read()
        data = parse_po_excel(file_bytes)

    # ── Show warnings ──
    for w in data.get("warnings", []):
        st.markdown(f'<div class="warn-box">{w}</div>', unsafe_allow_html=True)

    if data["lines"]:
        st.markdown('<div class="success-box">✓ Orden leída correctamente — ' +
                    str(len(data["lines"])) + ' líneas detectadas</div>', unsafe_allow_html=True)
        st.write("")

        # ── Header preview ──
        st.markdown('<p class="section-title">Datos del Header</p>', unsafe_allow_html=True)
        h_cols = st.columns(4)
        fields = [
            ("PO Number",    data["po_number"]),
            ("Customer",     data["customer_name"]),
            ("Ship Date",    data["ship_date"]),
            ("Cancel Date",  data["cancel_date"]),
            ("PO Date",      data["po_date"]),
            ("Ship To",      data["ship_to"]),
            ("Terms",        data["terms"]),
            ("Currency",     data["currency"]),
        ]
        for idx, (label, value) in enumerate(fields):
            with h_cols[idx % 4]:
                st.markdown(
                    f'<div class="field-box"><div class="field-label">{label}</div>'
                    f'<div class="field-value">{value or "—"}</div></div>',
                    unsafe_allow_html=True
                )

        st.write("")

        # ── Lines preview ──
        st.markdown('<p class="section-title">Líneas de la Orden</p>', unsafe_allow_html=True)
        preview_rows = []
        for l in data["lines"]:
            s = l["sizes"]
            preview_rows.append({
                "STOCK":       l["stock"],
                "COLOR":       l["color_name"] or l["color_code"],
                "DESCRIPTION": l["description"],
                "S":  s.get("S",0), "M": s.get("M",0), "L": s.get("L",0),
                "XL": s.get("XL",0), "2XL": s.get("2XL",0), "3XL": s.get("3XL",0),
                "UNITS":      l["total_units"],
                "LINE COST":  f'${l["line_cost"]:.2f}',
                "MSRP":       f'${l["msrp"]:.2f}',
                "TOTAL COST": f'${l["total_cost"]:.2f}',
            })

        st.dataframe(
            pd.DataFrame(preview_rows),
            use_container_width=True,
            hide_index=True
        )

        totU = sum(l["total_units"] for l in data["lines"])
        totC = sum(l["total_cost"]  for l in data["lines"])
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Estilos", len(data["lines"]))
        m2.metric("Total Unidades", f"{totU:,}")
        m3.metric("Total Costo", f"${totC:,.2f}")

        st.write("")

        # ── Generate button ──
        if st.button("⬇️  Generar ORDER FORM", type="primary", use_container_width=True):
            with st.spinner("Generando Excel..."):
                output = generate_order_form(data)
                customer = (data["customer_name"] or data["po_number"] or "ORDER").replace("/","")[:30]
                filename = f"ORDER FORM - {customer}.xlsx"

            st.download_button(
                label="📥 Descargar ORDER FORM",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.markdown('<div class="warn-box">⚠️ No se detectaron líneas de producto. Verifica el archivo.</div>',
                    unsafe_allow_html=True)

else:
    with col1:
        st.markdown("""
        <div style="color:#555; font-size:0.85rem; line-height:1.8; margin-top:8px;">
        Formatos soportados:<br>
        • Excel (.xlsx / .xls)<br>
        • PDF — próximamente<br>
        • Imagen — próximamente
        </div>
        """, unsafe_allow_html=True)
