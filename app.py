import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import re
import io
from datetime import datetime, date

# ─── PAGE CONFIG ────────────────────────────────────────────
st.set_page_config(page_title="Pro Standard — PO Generator", page_icon="🏆", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main-title{font-family:'Bebas Neue',sans-serif;font-size:3rem;letter-spacing:4px;color:#e8c84a;}
.sub-title{font-family:'Bebas Neue',sans-serif;font-size:1.2rem;letter-spacing:3px;color:#888;}
.section-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:2px;color:#aaa;}
.field-box{background:#1a1a1a;border-radius:6px;padding:10px 14px;margin:4px 0;}
.field-label{font-size:0.7rem;letter-spacing:1.5px;text-transform:uppercase;color:#666;}
.field-value{font-size:0.95rem;color:#f5f5f0;}
.warn-box{background:rgba(255,165,0,0.1);border:1px solid rgba(255,165,0,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#ffa500;}
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#4ade80;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# GLOSARIO COMPLETO
# Basado en Glosario_PO.docx + POs reales: Kings, Fitters, Fanatics, Dallas Cowboys, Pro Standard HO
# ══════════════════════════════════════════════════════════════
GLOSSARY = {
    # ── Encabezado ────────────────────────────────────────────
    "po_number": [
        "PO#","PO NUMBER","PO NO","PO NO.","PO NAME",
        "PURCHASE ORDER NUMBER","PURCHASE ORDER","ORDER NUMBER",
        "ORDER #","ORDER NO","P.O. NUMBER","P.O.#","PO:",
    ],
    "po_date": [
        "PO DATE","ORDER DATE","PO CREATION DATE","CREATION DATE",
        "ISSUE DATE","CREATED ON","DATE ISSUED","DATE",
        "FECHA DE EMISION","FECHA DE LA ORDEN","FECHA DE GENERACION",
    ],
    "ship_date": [
        # Glosario: Start Ship, Ship Window Open, In-DC Date, Delivery Date, Dispatch Date
        # Fitters usa "Start Date", Kings usa "Ship Date", Dallas Cowboys usa "Delivery Date"
        "SHIP DATE","START DATE","START SHIP","SHIP WINDOW OPEN",
        "IN-DC DATE","IN DC DATE","DELIVERY DATE","DISPATCH DATE",
        "SHIP BY","DELIVER BY","IN HANDS DATE","IN-HANDS DATE",
        "IN HANDS","NEED BY DATE","ON FLOOR",
        "FECHA DE EMBARQUE","FECHA DE ENVIO","FECHA PROGRAMADA",
    ],
    "cancel_date": [
        # Glosario: Cancel-by Date, DNSA, Expiration Date, Latest Ship Date
        "CANCEL DATE","CANCEL BY","CANCEL-BY DATE","CANCELLATION DATE",
        "TERMINATION DATE","DO NOT SHIP AFTER","DNSA",
        "EXPIRATION DATE","LATEST SHIP DATE",
        "FECHA LIMITE","FECHA DE CANCELACION","FECHA TOPE",
    ],
    "customer_name": [
        # Glosario: Buyer, Bill-to Name, Account Name, Cliente
        # NOTE: "BILL TO" intentionally excluded here — handled by special block in parser
        # to avoid false positives (Kings PO has "BILL TO:" with PO name on same row)
        "CUSTOMER NAME","SOLD TO","ACCOUNT NAME",
        "BILLED TO","CLIENT","BUYER NAME","BILL TO NAME",
        "RAZON SOCIAL","COMPRADOR","CONSIGNATARIO",
    ],
    "customer_code": [
        # Glosario: Account Number, Account #, Customer ID, Buyer Code
        "CUSTOMER CODE","ACCOUNT NUMBER","ACCOUNT #","ACCOUNT NO",
        "CUSTOMER ID","BUYER CODE","CLIENT CODE",
        "CUENTA DEL CLIENTE","ID DE CLIENTE","NO. DE CUENTA",
    ],
    "ship_to": [
        # Glosario: Delivery Address, Consignee, DC Address, Store Address
        "SHIP TO","SHIP-TO","SHIPPING ADDRESS","DELIVER TO",
        "DELIVERY ADDRESS","CONSIGNEE","DC ADDRESS","STORE ADDRESS",
        "DROP ADDRESS","DIRECCION DE ENVIO","DESTINATARIO",
    ],
    "bill_to": [
        # Glosario: Invoice Address, Sold to, Accounts Payable Address
        "BILL TO","BILL-TO","INVOICE ADDRESS","ACCOUNTS PAYABLE",
        "AP ADDRESS","BILLING ADDRESS","INVOICE TO","FACTURAR A",
    ],
    "terms": ["TERMS","PAYMENT TERMS","NET TERMS","NET"],

    # ── Columnas de tabla ─────────────────────────────────────
    "col_style": [
        # Glosario: Style Number, Style Code, Vendor Item Number, Model Number
        # Visto: Vendor Style No. (Kings), Style Number (Fitters), Vendor Item Number (Caesars)
        "VENDOR STYLE NO.","VENDOR STYLE NO","VENDOR STYLE NUMBER",
        "VENDOR STYLE","STYLE ID","STYLE #","STYLE NO","STYLE NO.",
        "STYLE NUMBER","STYLE CODE","STOCK","ITEM #","ITEM NUMBER",
        "ITEM NO","SKU","SKU#","PRODUCT CODE","STYLE",
        "VENDOR ITEM NUMBER","VENDOR ITEM","ARTICLE",
        "MODEL NUMBER","MANUFACTURER CODE",
        "ESTILO","REFERENCIA DEL PROVEEDOR",
    ],
    "col_desc": [
        # Glosario: Item Description, Product Name, Article Description
        "ITEM DESCRIPTION","ITEM DESC","DESCRIPTION","DESC",
        "PRODUCT NAME","PRODUCT DESCRIPTION","ARTICLE DESCRIPTION",
        "DETALLE DEL PRODUCTO","DESCRIPCION DEL ARTICULO",
    ],
    "col_color": [
        # Glosario: Color, Colorway, Hue, Color ID, Color Name
        "COLOR NAME","COLOR STORY","COLOUR NAME","COLOR",
        "COLOUR","COLORWAY","HUE","COLOR ID","CODIGO DE COLOR",
    ],
    "col_size_break": [
        # Glosario: Size scale, Size run, Size distribution, Pack ratio
        # Visto en Pro Standard HO: "SIZE SCALE"
        "SIZE SCALE","SIZE BREAK","SIZE RUN","SIZE DISTRIBUTION",
        "PACK RATIO","SIZE CURVE","SIZES",
        "QUEBRA DE TALLAS","CURVA DE TALLAS","DISTRIBUCION POR TALLA",
    ],
    "col_qty": [
        # Glosario: Total Qty, Total Pieces, Total Units, Sum of Qty
        "QTY","QUANTITY","PURCH QTY","ORDER QTY",
        "UNITS","# UNITS","PIECES","PCS",
        "TOTAL QTY","TOTAL UNITS","TTL UNITS","TTL QTY",
        "TOTAL PIECES","SUM OF QTY",
    ],
    "col_cost": [
        # Glosario: Net cost, Unit cost, Price Per Unit, Wholesale net
        # Visto: "Price Per Unit" (Fitters), "Cost W/ Disc" (Kings), "Unit Cost" (Caesars), "Cost" (Dallas Cowboys)
        "COST W/ DISC","COST W/DISC","COST WITH DISCOUNT",
        "NET COST","NET UNIT COST","UNIT COST","UNIT PRICE",
        "PRICE PER UNIT","WHOLESALE NET","DISCOUNTED PRICE",
        "WHOLESALE","COST","WS","PRICE",
        "COSTO NETO","PRECIO NETO","COSTO UNITARIO","MAYOREO NETO",
    ],
    "col_msrp": [
        # Glosario: Suggested Retail Price, Retail Price, List Price, Tag Price
        # Visto: "MSRP / Retail" (Fitters), "Retail Price" (Caesars), "Retail" (Kings)
        "MSRP","MSRP / RETAIL","RETAIL","RETAIL PRICE",
        "SUGGESTED RETAIL PRICE","SRP","LIST PRICE",
        "TAG PRICE","TICKET PRICE","PRECIO SUGERIDO","PVP",
    ],
    "col_total_cost": [
        # Glosario: Extended Cost, Ext Cost, Line Total, Total Dollars
        "TOTAL DOLLARS","EXT COST","EXTENDED COST","LINE TOTAL",
        "TOTAL COST","SUBTOTAL","COSTO EXTENDIDO","COSTO POR RENGLON",
    ],
    "col_total_retail": [
        # Glosario: Ext Retail, Retail Total, Sales Value
        "EXT RETAIL","EXTENDED RETAIL","RETAIL VALUE",
        "TOTAL RETAIL","RETAIL TOTAL","SALES VALUE","VALOR DE VENTA",
    ],
    "col_discount": [
        "DISC %","DISCOUNT %","DISC","DISCOUNT","OFF %","DESCUENTO",
    ],
    "col_upc": [
        # Glosario: Barcode, GTIN, EAN, Item barcode, Scan code
        "UPC","UPC CODE","BARCODE","GTIN","EAN",
        "EAN-13","SCAN CODE","ITEM BARCODE","CODIGO DE BARRAS",
    ],

    # ── Tallas individuales (formato matriz / columnas separadas) ──
    # Glosario: "Puede aparecer como matriz horizontal (XS,S,M,L,XL,2X,3X)"
    # Visto en Fitters y Saleh Sportswear
    "size_os":  ["OS","OSFA","OSFM","ONE SIZE","O/S","ONE-SIZE","ONE SIZE FITS ALL"],
    "size_xxs": ["XXS","2XS"],
    "size_xs":  ["XS"],
    "size_s":   ["S","SM","SML","S/P","SMALL"],
    "size_m":   ["M","MD","MED","M/M","MEDIUM"],
    "size_l":   ["L","LG","LRG","L/G","LARGE"],
    "size_xl":  ["XL","X-LARGE","X/L","XLARGE"],
    "size_2xl": ["2XL","XXL","2X","XX","2X-LARGE","2X LARGE","DOUBLE XL"],
    "size_3xl": ["3XL","XXXL","3X","3X-LARGE","3X LARGE","TRIPLE XL"],
}

SIZE_ORDER   = ["OS","XXS","XS","S","M","L","XL","2XL","3XL"]
SIZE_KEY_MAP = {
    "size_os":"OS","size_xxs":"XXS","size_xs":"XS",
    "size_s":"S","size_m":"M","size_l":"L",
    "size_xl":"XL","size_2xl":"2XL","size_3xl":"3XL",
}
SIZE_NORMALIZE = {
    "OSFA":"OS","OSFM":"OS","ONE SIZE":"OS","O/S":"OS","ONE-SIZE":"OS",
    "XXS":"XXS","2XS":"XXS",
    "XS":"XS",
    "S":"S","SM":"S","SML":"S","S/P":"S","SMALL":"S",
    "M":"M","MD":"M","MED":"M","M/M":"M","MEDIUM":"M",
    "L":"L","LG":"L","LRG":"L","L/G":"L","LARGE":"L",
    "XL":"XL","X-LARGE":"XL","X/L":"XL","XLARGE":"XL","EXTRA LARGE":"XL",
    "2XL":"2XL","XXL":"2XL","2X":"2XL","XX":"2XL","2X-LARGE":"2XL","2X LARGE":"2XL",
    "3XL":"3XL","XXXL":"3XL","3X":"3XL","3X-LARGE":"3XL","3X LARGE":"3XL",
}
COLOR_CODES = {
    "BLK":"BLACK","WHT":"WHITE","NVY":"NAVY","RED":"RED","BLU":"BLUE",
    "GRY":"GREY","WBK":"WASHED BLACK","EGG":"EGGSHELL","CRM":"CREAM",
    "DBL":"DODGER BLUE","HGR":"HEATHER GREY","DHG":"DARK HEATHER GREY",
    "LGY":"LIGHT GREY","MDN":"MIDNIGHT NAVY","RYB":"ROYAL BLUE",
    "NSF":"NIGHT SKY FADE","GLD":"GOLD","PUR":"PURPLE","FOR":"FOREST",
    "TEL":"TEAL","KGN":"KELLY GREEN","CRD":"CRIMSON RED","CNC":"CHARCOAL",
}

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def cell_str(v):
    return str(v).upper().strip() if v is not None else ""

def matches_any(text, keywords):
    t = text.upper().strip()
    for kw in keywords:
        k = kw.upper().strip()
        if t == k or t.startswith(k) or k in t:
            return True
    return False

def detect_col(header_row, keywords, exact_only=False):
    """Find column index. Prefers exact match over contains match.
    If exact_only=True, only returns exact matches (used for size columns)."""
    exact_match = None
    contains_match = None
    for i, cell in enumerate(header_row):
        c = cell_str(cell)
        if not c:
            continue
        for kw in keywords:
            k = kw.upper().strip()
            if c == k and exact_match is None:
                exact_match = i
                break
            elif not exact_only and k in c and contains_match is None:
                contains_match = i
    return exact_match if exact_match is not None else (None if exact_only else contains_match)

def normalize_size(s):
    u = str(s).upper().strip()
    return SIZE_NORMALIZE.get(u, u)

def fmtDate(v):
    if not v:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime('%m/%d/%Y')
    s = str(v).strip()
    s = re.sub(r'\s+\d{2}:\d{2}:\d{2}.*', '', s)
    s = re.sub(r'/+', '/', s)
    return s

def parse_style_color(style_code):
    s = str(style_code).strip()
    if "-" in s:
        parts = s.rsplit("-", 1)
        stock = parts[0].strip()
        cc    = parts[1].strip().upper()
        cn    = COLOR_CODES.get(cc, cc)
        return stock, cc, cn
    return s, "", ""

def parse_size_break_compressed(cell_value):
    sizes = {s: 0 for s in SIZE_ORDER}
    if not cell_value:
        return sizes
    val = str(cell_value).strip()
    parts = re.split(r'\s{2,}|\t', val, maxsplit=1)
    range_part   = parts[0].strip()
    numbers_part = parts[1].strip() if len(parts) > 1 else val
    range_match  = re.match(r'([A-Z0-9/\-]+?)\s*[-]+\s*([A-Z0-9]+)', range_part.upper())
    if range_match:
        start = normalize_size(range_match.group(1).split('/')[0].strip())
        end   = normalize_size(range_match.group(2).strip())
        try:
            si = SIZE_ORDER.index(start)
            ei = SIZE_ORDER.index(end)
            size_slice = SIZE_ORDER[si:ei+1]
        except ValueError:
            size_slice = ["S","M","L","XL","2XL","3XL"]
    else:
        size_slice = ["S","M","L","XL","2XL","3XL"]
    numbers = [int(n) for n in re.findall(r'\d+', numbers_part)]
    for i, sz in enumerate(size_slice):
        if i < len(numbers):
            sizes[sz] = numbers[i]
    return sizes

# Labels that should never be returned as values
_LABEL_WORDS = {
    "PO NUMBER","PO NAME","PO DATE","PO#","SHIP DATE","CANCEL DATE",
    "BILL TO","SHIP TO","VENDOR","BRAND","BUYER","TERMS","CURRENCY",
    "DESCRIPTION","STYLE","QTY","COST","RETAIL","MSRP","UPC",
}

def find_header_raw(rows, keywords, max_rows=25):
    for row in rows[:max_rows]:
        for j, cell in enumerate(row):
            c = cell_str(cell)
            if c and matches_any(c, keywords):
                raw = str(cell).strip()
                if ":" in raw:
                    after = raw.split(":", 1)[1].strip()
                    if after and after not in ("", "0", "None"):
                        return after
                for k in range(j+1, min(j+5, len(row))):
                    v = row[k]
                    if v is None:
                        continue
                    vs = str(v).strip()
                    if vs in ("", "0", "None"):
                        continue
                    # Skip values that look like column labels
                    if vs.upper() in _LABEL_WORDS:
                        continue
                    return v
    return None

def find_header_value(rows, keywords, max_rows=25):
    v = find_header_raw(rows, keywords, max_rows)
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v
    return str(v).strip()

# ══════════════════════════════════════════════════════════════
# PARSER PRINCIPAL
# ══════════════════════════════════════════════════════════════

def parse_po_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    result = {
        "po_number":"","po_date":"","ship_date":"","cancel_date":"",
        "customer_name":"","customer_code":"","ship_to":"","bill_to":"",
        "terms":"","currency":"USD","lines":[],"warnings":[]
    }

    # ── 1. Header ────────────────────────────────────────────
    po_number = find_header_value(rows, GLOSSARY["po_number"])
    if not po_number:
        for row in rows[:5]:
            for cell in row:
                if cell and "PO" in str(cell).upper():
                    m = re.search(r'PO#?\s*([A-Z0-9][A-Z0-9\s\-]+)', str(cell), re.IGNORECASE)
                    if m:
                        candidate = m.group(1).strip()
                        if len(candidate) > 2 and candidate.upper() not in ("NAME","NUMBER","DATE","NO"):
                            po_number = candidate
                            break
            if po_number:
                break
    result["po_number"]   = str(po_number).strip() if po_number else ""
    result["po_date"]     = fmtDate(find_header_value(rows, GLOSSARY["po_date"]))
    result["ship_date"]   = fmtDate(find_header_value(rows, GLOSSARY["ship_date"]))
    result["cancel_date"] = fmtDate(find_header_value(rows, GLOSSARY["cancel_date"]))
    result["terms"]       = find_header_value(rows, GLOSSARY["terms"])

    # Customer name — first try generic labels, then special BILL TO block
    customer_name = find_header_value(rows, GLOSSARY["customer_name"])
    if not customer_name:
        for i, row in enumerate(rows[:15]):
            for j, cell in enumerate(row):
                c = cell_str(cell)
                if c in ("BILL TO:","BILL TO","SOLD TO:","SOLD TO"):
                    # Value is on the NEXT ROW at same column (Kings format)
                    # or after colon on same cell
                    raw = str(cell).strip()
                    if ":" in raw:
                        after = raw.split(":",1)[1].strip()
                        # Make sure it's not just another label
                        if after and len(after) > 2 and after.upper() not in _LABEL_WORDS:
                            customer_name = after; break
                    # Next row same column
                    if i+1 < len(rows) and j < len(rows[i+1]):
                        v = str(rows[i+1][j] or "").strip()
                        if v and len(v) > 2 and v.upper() not in _LABEL_WORDS:
                            customer_name = v; break
            if customer_name:
                break
    result["customer_name"] = customer_name
    result["customer_code"] = find_header_value(rows, GLOSSARY["customer_code"])

    ship_to = find_header_value(rows, GLOSSARY["ship_to"])
    if not ship_to:
        for i, row in enumerate(rows[:15]):
            for j, cell in enumerate(row):
                c = cell_str(cell)
                if matches_any(c, ["SHIP TO:","SHIP TO","SHIPPING ADDRESS:","DELIVERY ADDRESS:"]):
                    parts = []
                    for k in range(i+1, min(i+5, len(rows))):
                        if j < len(rows[k]):
                            v = str(rows[k][j] or "").strip()
                            if v.upper() in ("DESCRIPTION","VENDOR STYLE NO.","STYLE","QTY","UPC","STYLE NUMBER"):
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

    # ── 2. Encontrar tabla de productos ──────────────────────
    data_header_row = None
    col_map = {}

    for i, row in enumerate(rows):
        row_str = " | ".join(cell_str(c) for c in row)
        has_style = matches_any(row_str, GLOSSARY["col_style"])
        has_qty   = (matches_any(row_str, GLOSSARY["col_qty"]) or
                     matches_any(row_str, GLOSSARY["col_size_break"]) or
                     any(matches_any(row_str, GLOSSARY[k]) for k in SIZE_KEY_MAP))
        if has_style and has_qty:
            data_header_row = i
            col_map["style"]        = detect_col(row, GLOSSARY["col_style"])
            col_map["desc"]         = detect_col(row, GLOSSARY["col_desc"])
            col_map["color"]        = detect_col(row, GLOSSARY["col_color"])
            col_map["size_break"]   = detect_col(row, GLOSSARY["col_size_break"])
            col_map["size"]         = detect_col(row, ["SIZE"])
            col_map["qty"]          = detect_col(row, GLOSSARY["col_qty"])
            col_map["cost"]         = detect_col(row, GLOSSARY["col_cost"])
            col_map["msrp"]         = detect_col(row, GLOSSARY["col_msrp"])
            col_map["discount"]     = detect_col(row, GLOSSARY["col_discount"])
            col_map["total_cost"]   = detect_col(row, GLOSSARY["col_total_cost"])
            col_map["total_retail"] = detect_col(row, GLOSSARY["col_total_retail"])
            for sz_key in SIZE_KEY_MAP:
                col_map[sz_key] = detect_col(row, GLOSSARY[sz_key], exact_only=True)
            break

    if data_header_row is None:
        result["warnings"].append("⚠️ No se encontró la tabla de productos.")
        return result

    has_matrix     = any(col_map.get(k) is not None for k in SIZE_KEY_MAP)
    has_size_col   = col_map.get("size") is not None
    has_size_break = col_map.get("size_break") is not None

    # ── 3. Leer filas de productos ───────────────────────────
    lines = []
    last_valid_style = ""

    for row in rows[data_header_row + 1:]:
        row_text = " ".join(cell_str(c) for c in row if c)

        # Detectar ship_date embebida en fila (ej. "Ship Date: 10/1/2026")
        if matches_any(row_text, GLOSSARY["ship_date"]) and not result["ship_date"]:
            for c in row:
                if c and matches_any(cell_str(c), GLOSSARY["ship_date"]):
                    raw = str(c).strip()
                    if ":" in raw:
                        val = raw.split(":",1)[1].strip()
                        if val:
                            result["ship_date"] = fmtDate(val)
                            break
            continue

        # Obtener style
        style_raw = ""
        if col_map.get("style") is not None and col_map["style"] < len(row):
            style_raw = str(row[col_map["style"]] or "").strip()

        if style_raw:
            if matches_any(style_raw, ["TOTAL","SUBTOTAL","GRAND TOTAL","SUB-TOTAL"]):
                continue
            if re.match(r'^[A-Za-z]{2,}\d+', style_raw):
                last_valid_style = style_raw
            else:
                continue
        else:
            if last_valid_style and (has_size_col or has_matrix):
                style_raw = last_valid_style
            else:
                continue

        stock, color_code, color_name = parse_style_color(style_raw)

        desc = ""
        if col_map.get("desc") is not None and col_map["desc"] < len(row):
            desc = str(row[col_map["desc"]] or "").strip()

        if col_map.get("color") is not None and col_map["color"] < len(row):
            cv = str(row[col_map["color"]] or "").strip()
            if cv:
                color_name = cv

        cost = 0.0
        if col_map.get("cost") is not None and col_map["cost"] < len(row):
            try: cost = float(row[col_map["cost"]] or 0)
            except: pass

        msrp = 0.0
        if col_map.get("msrp") is not None and col_map["msrp"] < len(row):
            try: msrp = float(row[col_map["msrp"]] or 0)
            except: pass

        discount = 0.0
        disc_col = col_map.get("discount")
        if disc_col is None:
            for idx in range(len(row)-1, -1, -1):
                v = row[idx]
                if v is not None:
                    try:
                        d = float(v)
                        if 0 < d <= 1:
                            disc_col = idx
                    except: pass
                    break
        if disc_col is not None and disc_col < len(row):
            try:
                d = float(row[disc_col] or 0)
                discount = d if d <= 1 else d / 100
            except: pass

        # ── Tallas ───────────────────────────────────────────
        sizes = {s: 0 for s in SIZE_ORDER}
        total_units = 0
        size_val = ""
        qty_val  = 0

        if has_size_break and col_map.get("size_break") is not None and col_map["size_break"] < len(row):
            sb = row[col_map["size_break"]]
            if sb:
                sizes = parse_size_break_compressed(sb)
                total_units = sum(sizes.values())

        elif has_matrix:
            for sz_key, sz_name in SIZE_KEY_MAP.items():
                ci = col_map.get(sz_key)
                if ci is not None and ci < len(row):
                    try: sizes[sz_name] = int(float(row[ci] or 0))
                    except: pass
            total_units = sum(sizes.values())

        elif has_size_col and col_map.get("size") is not None:
            size_val = str(row[col_map["size"]] or "").strip()
            if col_map.get("qty") is not None and col_map["qty"] < len(row):
                try: qty_val = int(float(row[col_map["qty"]] or 0))
                except: pass
            if "/" in size_val:
                parts = size_val.split("/")
                size_val = parts[0].strip()
                if not color_code:
                    color_code = parts[1].strip()[:3]
            if size_val:
                norm = normalize_size(size_val)
                sizes[norm] = qty_val
                total_units = qty_val

        elif col_map.get("qty") is not None and col_map["qty"] < len(row):
            try:
                total_units = int(float(row[col_map["qty"]] or 0))
                sizes["OS"] = total_units
            except: pass

        if total_units == 0 and not desc:
            continue
        if has_size_col and not has_matrix and not size_val and not has_size_break:
            continue

        line_cost    = cost * (1 - discount) if discount else cost
        total_cost   = line_cost * total_units
        total_retail = msrp * total_units

        # Merge con línea existente (formato fila-por-talla)
        merged = False
        if has_size_col and size_val:
            for existing in reversed(lines):
                if existing["stock"] == stock and existing["color_code"] == color_code:
                    norm = normalize_size(size_val)
                    existing["sizes"][norm] = existing["sizes"].get(norm, 0) + qty_val
                    existing["total_units"] += qty_val
                    if not existing["cost"] and cost: existing["cost"] = cost
                    if not existing["msrp"] and msrp: existing["msrp"] = msrp
                    if not existing["discount"] and discount: existing["discount"] = discount
                    existing["line_cost"]    = existing["cost"] * (1 - existing["discount"]) if existing["discount"] else existing["cost"]
                    existing["total_cost"]   = existing["line_cost"] * existing["total_units"]
                    existing["total_retail"] = existing["msrp"] * existing["total_units"]
                    merged = True
                    break

        if not merged:
            lines.append({
                "style_raw": style_raw, "stock": stock,
                "color_code": color_code, "color_name": color_name,
                "description": desc, "cost": cost, "msrp": msrp,
                "discount": discount, "line_cost": line_cost,
                "sizes": sizes, "total_units": total_units,
                "total_cost": total_cost, "total_retail": total_retail,
            })

    result["lines"] = lines
    return result

# ══════════════════════════════════════════════════════════════
# GENERAR ORDER FORM EXCEL
# ══════════════════════════════════════════════════════════════

def generate_order_form(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ORDER FORM"

    lf = Font(name="Calibri", bold=True, size=9, color="888888")
    vf = Font(name="Calibri", size=9)
    hf = Font(name="Calibri", bold=True, size=8, color="FFFFFF")
    bf = Font(name="Calibri", bold=True, size=9)

    totU = sum(l["total_units"] for l in data["lines"])
    totC = sum(l["total_cost"]  for l in data["lines"])

    def wlv(r, cl, label, cv, value):
        ws.cell(row=r, column=cl, value=label).font = lf
        ws.cell(row=r, column=cv, value=value).font = vf

    # Get discount % from first line that has one
    disc_pct = 0.0
    for l in data["lines"]:
        if l.get("discount"):
            disc_pct = l["discount"]
            break
    disc_label = f'{disc_pct*100:.0f}%' if disc_pct else ""
    cost_w_disc_header = f"COST W/DISCOUNT {disc_label}".strip()

    # Match ORDER_FORM.xlsx template exactly
    wlv(1,1,"Customer code:",2,data.get("customer_code",""))
    ws.cell(row=1,column=5,value="Vendor:").font=lf
    ws.cell(row=1,column=6,value="Maxima Apparel Corp SAPI de CV").font=vf
    ws.cell(row=1,column=8,value="Brand:").font=lf
    ws.cell(row=1,column=9,value="PRO STANDARD").font=bf

    wlv(3,1,"Customer name:",2,data.get("customer_name",""))
    wlv(3,5,"Ship to:",6,data.get("ship_to",""))
    wlv(5,1,"PO Number:",2,data.get("po_number",""))
    wlv(6,1,"PO Creation Date:",2,data.get("po_date",""))
    wlv(8,1,"Ship Date:",2,data.get("ship_date",""))
    wlv(9,1,"Cancel Date:",2,data.get("cancel_date",""))
    ws.cell(row=9,column=5,value="Total cost:").font=lf
    ws.cell(row=9,column=6,value=round(totC,2)).font=vf
    ws.cell(row=10,column=5,value="Total qty:").font=lf
    ws.cell(row=10,column=6,value=totU).font=vf
    ws.cell(row=11,column=5,value="Currency:").font=lf
    ws.cell(row=11,column=6,value=data.get("currency","USD")).font=vf
    ws.cell(row=12,column=5,value="Discount %:").font=lf
    ws.cell(row=12,column=6,value=disc_label).font=vf

    COLS = ["BRAND","STOCK","COLOR CODE","COLOR NAME","DESCRIPTION","CURRENCY",
            "LINE COST", cost_w_disc_header, "MSRP",
            "OS","XXS","XS","S","M","L","XL","2XL","3XL",
            "TOTAL UNITS","TOTAL COST","TOTAL RETAIL"]
    for j, cn in enumerate(COLS, start=1):
        c = ws.cell(row=13, column=j, value=cn)
        c.font = hf
        c.fill = PatternFill("solid", fgColor="222222")
        c.alignment = Alignment(horizontal="center")

    for i, line in enumerate(data["lines"], start=14):
        s = line["sizes"]
        rd = ["PRO STANDARD", line["stock"], line["color_code"], line["color_name"],
              line["description"], data.get("currency","USD"),
              round(line["cost"],2),        # LINE COST = costo bruto
              round(line["line_cost"],2),   # COST W/DISCOUNT = costo neto
              round(line["msrp"],2),
              s.get("OS",0), s.get("XXS",0), s.get("XS",0),
              s.get("S",0), s.get("M",0), s.get("L",0),
              s.get("XL",0), s.get("2XL",0), s.get("3XL",0),
              line["total_units"], round(line["total_cost"],2), round(line["total_retail"],2)]
        for j, val in enumerate(rd, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.font = vf
            c.alignment = Alignment(horizontal="center" if j > 5 else "left")

    for i, w in enumerate([14,20,11,14,42,9,10,14,8,6,6,6,6,6,6,6,6,6,12,12,12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

st.markdown('<p class="main-title">PURCHASE ORDER</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">PRO STANDARD — GENERATOR</p>', unsafe_allow_html=True)
st.divider()

col1, col2 = st.columns([1, 2])
with col1:
    st.markdown('<p class="section-title">📋 Orden del Cliente</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Arrastra o selecciona", type=["xlsx","xls"], label_visibility="collapsed")

if uploaded_file:
    with st.spinner("Leyendo archivo..."):
        data = parse_po_excel(uploaded_file.read())

    for w in data.get("warnings", []):
        st.markdown(f'<div class="warn-box">{w}</div>', unsafe_allow_html=True)

    if data["lines"]:
        st.markdown('<div class="success-box">✓ ' + str(len(data["lines"])) + ' estilos detectados</div>', unsafe_allow_html=True)
        st.write("")

        st.markdown('<p class="section-title">Datos del Header</p>', unsafe_allow_html=True)
        hc = st.columns(4)
        for idx, (label, value) in enumerate([
            ("PO Number",data["po_number"]),("Cliente",data["customer_name"]),
            ("Ship Date",data["ship_date"]),("Cancel Date",data["cancel_date"]),
            ("PO Date",data["po_date"]),("Ship To",data["ship_to"]),
            ("Terms",data["terms"]),("Currency",data["currency"]),
        ]):
            with hc[idx % 4]:
                st.markdown(f'<div class="field-box"><div class="field-label">{label}</div><div class="field-value">{value or "—"}</div></div>', unsafe_allow_html=True)

        st.write("")
        st.markdown('<p class="section-title">Líneas de la Orden</p>', unsafe_allow_html=True)
        preview = []
        for l in data["lines"]:
            s = l["sizes"]
            preview.append({
                "STOCK":l["stock"],"COLOR":l["color_name"] or l["color_code"],
                "DESCRIPTION":l["description"],
                "OS":s.get("OS",0),"S":s.get("S",0),"M":s.get("M",0),
                "L":s.get("L",0),"XL":s.get("XL",0),"2XL":s.get("2XL",0),"3XL":s.get("3XL",0),
                "UNITS":l["total_units"],
                "LINE COST":f'${l["line_cost"]:.2f}',"MSRP":f'${l["msrp"]:.2f}',
                "TOTAL COST":f'${l["total_cost"]:.2f}',
            })
        st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

        totU = sum(l["total_units"] for l in data["lines"])
        totC = sum(l["total_cost"]  for l in data["lines"])
        m1,m2,m3 = st.columns(3)
        m1.metric("Estilos", len(data["lines"]))
        m2.metric("Unidades", f"{totU:,}")
        m3.metric("Total Costo", f"${totC:,.2f}")

        st.write("")
        if st.button("⬇️  Generar ORDER FORM", type="primary", use_container_width=True):
            with st.spinner("Generando..."):
                output   = generate_order_form(data)
                customer = (data["customer_name"] or data["po_number"] or "ORDER").replace("/","")[:30]
                filename = f"ORDER FORM - {customer}.xlsx"
            st.download_button("📥 Descargar ORDER FORM", data=output, file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
    else:
        st.markdown('<div class="warn-box">⚠️ No se detectaron líneas. Verifica el archivo.</div>', unsafe_allow_html=True)
else:
    with col1:
        st.markdown("""
        <div style="color:#555;font-size:0.85rem;line-height:1.8;margin-top:8px;">
        Formatos soportados:<br>
        • Tallas en filas (Kings / Caesars)<br>
        • Tallas en columnas (Fitters / Saleh)<br>
        • Size break comprimido (Pro Standard HO)<br>
        • Cualquier formato con encabezado en Excel
        </div>""", unsafe_allow_html=True)
