import streamlit as st
import pandas as pd
import fitz  # pymupdf
import re
import io
import os
import psycopg2
from datetime import date
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

DB_HOST     = "db.maximaapparel.com"
DB_PORT     = 5432
DB_NAME     = "maxima_reporting"
DB_USER     = "agent_ro"
DB_PASSWORD = "R3@d1234!1"
TABLE       = "public.sss_upc_report"

CURRENCY_MAP = {
    "USD": ("wholesale_usd", "msrp_usd"),
    "CAD": ("wholesale_cad", "msrp_cad"),
    "GBP": ("wholesale_gbp", "msrp_gbp"),
    "EUR": ("wholesale_eur", "msrp_eur"),
    "AED": ("wholesale_aed", "msrp_aed"),
    "MXN": ("wholesale_mxn", "msrp_mxn"),
    "BRL": ("wholesale_brl", "msrp_brl"),
    "CLP": ("wholesale_clp", "msrp_clp"),
    "AUD": ("wholesale_aud", "msrp_aud"),
    "NZD": ("wholesale_nzd", "msrp_nzd"),
    "RMB": ("wholesale_rmb", "msrp_rmb"),
    "ARS": ("wholesale_ars", "msrp_ars"),
    "ECU": ("wholesale_ecu", "msrp_ecu"),
    "BOB": ("wholesale_bob", "msrp_bob"),
    "PEN": ("wholesale_pen", "msrp_pen"),
}

CURRENCY_SYMBOLS = {
    "USD": "$",    "CAD": "CA$",  "GBP": "£",    "EUR": "€",
    "AED": "AED ", "MXN": "MX$", "BRL": "R$",   "CLP": "CLP$",
    "AUD": "AU$",  "NZD": "NZ$", "RMB": "¥",    "ARS": "AR$",
    "ECU": "$",    "BOB": "Bs.", "PEN": "S/.",
}

PRICE_RE = re.compile(r'WS:\s*\$?([\d,]+\.?\d*)\s*/\s*MSRP:\s*\$?([\d,]+\.?\d*)', re.IGNORECASE)
STYLE_RE = re.compile(r'^[A-Z]{2,5}\d{6,}-[A-Z0-9]{2,5}$')

st.set_page_config(page_title="PDF → PPT Converter", page_icon="📑", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main-title{font-family:'Bebas Neue',sans-serif;font-size:3rem;letter-spacing:4px;color:#e8c84a;}
.sub-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:3px;color:#888;}
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:12px 16px;font-size:0.9rem;color:#4ade80;}
.info-box{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);border-radius:6px;padding:12px 16px;font-size:0.85rem;color:#ccc;margin-bottom:8px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">PDF → PPT CONVERTER</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MAXIMA APPAREL — PRICE CONVERTER</div>', unsafe_allow_html=True)
st.markdown("---")


def fetch_prices(style_codes):
    if not style_codes:
        return {}
    price_cols = [col for pair in CURRENCY_MAP.values() for col in pair]
    try:
        conn = psycopg2.connect(
            host=DB_HOST, port=DB_PORT, dbname=DB_NAME,
            user=DB_USER, password=DB_PASSWORD, connect_timeout=15
        )
        cols_str = ", ".join(["ivnum"] + price_cols)
        df = pd.read_sql(
            f"SELECT DISTINCT {cols_str} FROM {TABLE} WHERE ivnum = ANY(%s)",
            conn, params=(list(style_codes),)
        )
        conn.close()
        return {row["ivnum"]: row for _, row in df.iterrows()}
    except Exception as e:
        st.error(f"Error DB: {e}")
        return {}


def sample_bg(page, rect):
    """Sample background color under a rect before redaction."""
    clip = fitz.Rect(rect)
    pix = page.get_pixmap(matrix=fitz.Matrix(1, 1), clip=clip, alpha=False)
    if pix.width > 0 and pix.height > 0:
        cx, cy = max(0, pix.width // 2 - 1), max(0, pix.height // 2 - 1)
        s = pix.samples
        idx = (cy * pix.width + cx) * pix.n
        r, g, b = s[idx] / 255, s[idx + 1] / 255, s[idx + 2] / 255
        return (r, g, b)
    return (1.0, 1.0, 1.0)


def int_to_rgb(color_int):
    r = ((color_int >> 16) & 0xFF) / 255
    g = ((color_int >> 8) & 0xFF) / 255
    b = (color_int & 0xFF) / 255
    return (r, g, b)


def fmt_price(val, currency):
    sym = CURRENCY_SYMBOLS.get(currency, "$")
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "N/A"
    if currency == "CLP":
        return f"{sym}{int(float(val)):,}"
    return f"{sym}{float(val):,.2f}"


def process_pdf(pdf_bytes, price_data, currency):
    """Replace prices in PDF for given currency, return list of PNG bytes (non-empty pages only)."""
    ws_col, ms_col = CURRENCY_MAP[currency]
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    result_pngs = []

    for page_num in range(len(doc)):
        page = doc[page_num]

        # Skip empty pages
        if not page.get_text().strip():
            continue

        # Collect all text spans in reading order
        spans = []
        for block in page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span["text"].strip():
                        spans.append(span)

        # Find (style_code, price_span) pairs
        replacements = []
        current_style = None
        for span in spans:
            txt = span["text"].strip()
            if STYLE_RE.match(txt):
                current_style = txt
            elif PRICE_RE.search(txt) and current_style:
                row = price_data.get(current_style)
                if row is not None:
                    ws_val = row.get(ws_col)
                    ms_val = row.get(ms_col)
                    new_text = f"WS: {fmt_price(ws_val, currency)} / MSRP: {fmt_price(ms_val, currency)}"
                    replacements.append({
                        "rect":     fitz.Rect(span["bbox"]),
                        "new_text": new_text,
                        "color":    int_to_rgb(span["color"]),
                        "size":     span["size"],
                    })
                current_style = None

        # Sample backgrounds, then redact
        bg_colors = [sample_bg(page, r["rect"]) for r in replacements]

        for rep, bg in zip(replacements, bg_colors):
            page.add_redact_annot(rep["rect"], fill=bg)
        page.apply_redactions()

        # Re-insert new price text
        for rep, bg in zip(replacements, bg_colors):
            rect = rep["rect"]
            page.insert_text(
                (rect.x0, rect.y1 - 1),
                rep["new_text"],
                fontsize=rep["size"],
                color=rep["color"],
                fontname="helv",
            )

        # Render page → PNG at 2× for quality
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        result_pngs.append(pix.tobytes("png"))

    doc.close()
    return result_pngs


def build_pptx(pngs_by_currency):
    """One slide per (currency, page). Currencies grouped in order."""
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for currency, pngs in pngs_by_currency.items():
        for png_bytes in pngs:
            slide = prs.slides.add_slide(blank)
            img = Image.open(io.BytesIO(png_bytes))
            w, h = img.size
            aspect = w / h
            sw, sh = prs.slide_width, prs.slide_height
            if aspect > sw / sh:
                pw, ph = sw, int(sw / aspect)
            else:
                ph, pw = sh, int(sh * aspect)
            left = (sw - pw) // 2
            top  = (sh - ph) // 2
            slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, pw, ph)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("#### Paso 1 — Sube los PDFs")
uploaded_pdfs = st.file_uploader(
    "Arrastra o selecciona los PDFs de presentación",
    type=["pdf"],
    accept_multiple_files=True,
)

st.markdown("#### Paso 2 — Selecciona moneda(s)")
selected_currencies = st.multiselect(
    "Moneda(s)",
    options=list(CURRENCY_MAP.keys()),
    default=["USD"],
    placeholder="Selecciona moneda(s)...",
)

if uploaded_pdfs and selected_currencies:
    if st.button("🔄 CONVERTIR Y GENERAR PPT", type="primary", use_container_width=True):
        st.session_state.pop("ppt_output", None)

        # 1. Extract all style codes from uploaded PDFs
        all_styles = set()
        pdf_store = {}
        for f in uploaded_pdfs:
            pdf_bytes = f.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for pn in range(len(doc)):
                for line in doc[pn].get_text().splitlines():
                    if STYLE_RE.match(line.strip()):
                        all_styles.add(line.strip())
            doc.close()
            pdf_store[f.name] = pdf_bytes

        st.info(f"📦 {len(all_styles)} estilos detectados en {len(uploaded_pdfs)} PDF(s)")

        # 2. Fetch prices from DB
        with st.spinner("Consultando base de datos..."):
            price_data = fetch_prices(all_styles)

        found = len(price_data)
        missing = all_styles - set(price_data.keys())
        st.info(f"✅ {found} encontrados en DB" + (f" | ⚠️ {len(missing)} no encontrados: {', '.join(sorted(missing))}" if missing else ""))

        # 3. Process each PDF for each currency
        pngs_by_currency = {curr: [] for curr in selected_currencies}
        for pdf_name, pdf_bytes in pdf_store.items():
            for currency in selected_currencies:
                with st.spinner(f"Procesando {pdf_name} → {currency}..."):
                    pngs = process_pdf(pdf_bytes, price_data, currency)
                    pngs_by_currency[currency].extend(pngs)

        total_slides = sum(len(v) for v in pngs_by_currency.values())

        # 4. Build PPT
        with st.spinner("Generando PPT..."):
            pptx_bytes = build_pptx(pngs_by_currency)

        curr_label = "_".join(selected_currencies)
        fname = f"Presentacion_{curr_label}_{date.today().strftime('%m.%d.%Y')}.pptx"

        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "UPCs_generados")
        os.makedirs(out_dir, exist_ok=True)
        try:
            with open(os.path.join(out_dir, fname), "wb") as fh:
                fh.write(pptx_bytes)
        except Exception:
            pass

        st.session_state["ppt_output"] = {"fname": fname, "bytes": pptx_bytes, "slides": total_slides}

out = st.session_state.get("ppt_output")
if out:
    st.markdown(f'<div class="success-box">✅ {out["slides"]} diapositiva(s) generada(s)</div>', unsafe_allow_html=True)
    st.download_button(
        label=f"⬇ Descargar {out['fname']}",
        data=out["bytes"],
        file_name=out["fname"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
