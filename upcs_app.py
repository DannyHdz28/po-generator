import streamlit as st
import pandas as pd
import io
import os
from datetime import date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from pbi_downloader import run_download

st.set_page_config(page_title="UPCs Generator", page_icon="🏷️", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main-title{font-family:'Bebas Neue',sans-serif;font-size:3rem;letter-spacing:4px;color:#e8c84a;}
.sub-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:3px;color:#888;}
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:12px 16px;font-size:0.9rem;color:#4ade80;}
.step-box{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);border-radius:6px;padding:12px 16px;font-size:0.85rem;color:#ccc;margin-bottom:8px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">UPCs GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MAXIMA APPAREL</div>', unsafe_allow_html=True)
st.markdown("---")

SIZES = ["2T", "3T", "4", "4T", "5", "6", "7", "OS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]

st.markdown("#### Paso 1 — Obtén los datos")

tab_auto, tab_manual = st.tabs(["🔄 Automático (Base de datos)", "📁 Manual (subir archivo)"])

with tab_auto:
    st.markdown("""
    <div class="step-box">
    Conecta directamente a la base de datos de Maxima Apparel y descarga los UPCs automáticamente.
    </div>
    """, unsafe_allow_html=True)

    if "db_df" not in st.session_state:
        st.session_state.db_df = None
    if "db_log" not in st.session_state:
        st.session_state.db_log = []

    if st.button("🔌 Conectar y Descargar", use_container_width=True):
        st.session_state.db_df = None
        st.session_state.db_log = []
        st.session_state.pop("output", None)
        log_box = st.empty()

        def progress(msg):
            st.session_state.db_log.append(msg)
            log_box.markdown("\n\n".join(f"• {m}" for m in st.session_state.db_log))

        with st.spinner("Conectando..."):
            result = run_download(progress_fn=progress)
        st.session_state.db_df = result if result is not None and not isinstance(result, list) else None

    if st.session_state.db_log:
        with st.expander("Log de conexión", expanded=True):
            for msg in st.session_state.db_log:
                color = "#f87171" if "ERROR" in msg.upper() else "#4ade80" if "obtenidos" in msg.lower() else "#ccc"
                st.markdown(f'<div style="font-size:0.85rem;color:{color};">• {msg}</div>', unsafe_allow_html=True)

with tab_manual:
    st.markdown("""
    <div class="step-box">
    Exporta manualmente desde Power BI y sube el archivo aquí.
    </div>
    """, unsafe_allow_html=True)

uploaded = st.file_uploader("Sube el archivo exportado de Power BI", type=["xlsx", "csv"], accept_multiple_files=True)

st.markdown("#### Paso 2 — Genera el archivo")
file_date = st.date_input("Fecha", value=date.today())

BRANDS = ["Pro Standard", "Off White / L/AB"]
selected_brands = st.multiselect(
    "Filtrar por marca (vacío = todas las marcas)",
    options=BRANDS,
    default=[],
    placeholder="Selecciona una o más marcas..."
)


def build_base_from_df(df, brands=None):
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    rename = {}
    for col in df.columns:
        if col in ("ivstyle", "style", "estilo"):
            rename[col] = "Style"
        elif col in ("size", "talla"):
            rename[col] = "Size"
        elif col == "upc":
            rename[col] = "UPC"
        elif col in ("description", "ivdesc", "descripcion", "descripción"):
            rename[col] = "DESCRIPTION"
        elif col in ("wholesale_usd", "wholesale", "mayoreo"):
            rename[col] = "WHOLESALE"
        elif col in ("msrp_usd", "msrp"):
            rename[col] = "MSRP"
        elif col in ("brand_name", "reporting_brand_name"):
            rename[col] = col  # conservar para filtrar
    df = df.rename(columns=rename)

    # Filtro de marca (si se seleccionó alguna)
    if brands:
        # reporting_brand_name: "Pro Standard" agrupa todas las ligas (NBA,NFL,MLB,etc.)
        # "Off-White Division" agrupa Off-White y L/AB
        brand_map = {
            "Pro Standard": ["Pro Standard"],
            "Off White / L/AB": ["Off-White Division"],
        }
        allowed = []
        for b in brands:
            allowed.extend(brand_map.get(b, [b]))
        if "reporting_brand_name" in df.columns:
            df = df[df["reporting_brand_name"].astype(str).str.strip().isin(allowed)]

    needed = [c for c in ["Style", "Size", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"] if c in df.columns]
    df = df[needed].copy()
    df["UPC"] = df["UPC"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df = df[df["UPC"].str.len() > 3]
    df = df[df["UPC"].str.lower() != "nan"]
    if "Size" in df.columns:
        df = df[df["Size"].astype(str).str.strip().isin(SIZES)]
    df = df[df["Style"].notna() & (df["Style"].astype(str).str.strip() != "")]
    df = df.drop_duplicates(subset=["Style", "Size"], keep="first")
    return df.reset_index(drop=True)


def build_base(files):
    all_rows = []
    for f in files:
        for header_row in [2, 0, 1, 3]:
            try:
                df = pd.read_excel(f, header=header_row, dtype=str)
                df.columns = [str(c).strip() for c in df.columns]

                rename = {}
                for col in df.columns:
                    cl = col.lower().strip()
                    if cl in ["ivnum", "estilo", "style"]:
                        rename[col] = "Style"
                    elif cl in ["size", "talla"]:
                        rename[col] = "Size"
                    elif cl == "upc":
                        rename[col] = "UPC"
                    elif cl in ["description", "descripcion", "descripción"]:
                        rename[col] = "DESCRIPTION"
                    elif cl in ["wholesale", "mayoreo"]:
                        rename[col] = "WHOLESALE"
                    elif cl == "msrp":
                        rename[col] = "MSRP"
                df = df.rename(columns=rename)

                if "Style" not in df.columns or "UPC" not in df.columns:
                    continue

                needed = [c for c in ["Style", "Size", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"] if c in df.columns]
                df = df[needed].copy()
                df["UPC"] = df["UPC"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(12)
                df = df[df["UPC"].str.strip().str.len() > 3]
                df = df[df["UPC"].str.strip() != "nan"]

                if "Size" in df.columns:
                    df = df[df["Size"].isin(SIZES)]

                df = df[df["Style"].notna() & (df["Style"].str.strip() != "")]
                all_rows.append(df)
                break
            except Exception:
                continue

    if not all_rows:
        return pd.DataFrame(columns=["Style", "Size", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"])

    base = pd.concat(all_rows, ignore_index=True)
    if "Style" in base.columns and "Size" in base.columns:
        base = base.drop_duplicates(subset=["Style", "Size"], keep="first")
    return base.reset_index(drop=True)


def build_upc_sheet(base):
    rows = []
    for _, r in base.iterrows():
        rows.append({
            "Styles": r.get("Style", ""),
            "Size": r.get("Size", ""),
            "CONCAT": f"{r.get('Style','')}{r.get('Size','')}",
            "UPC": r.get("UPC", ""),
            "DESCRIPTION": r.get("DESCRIPTION", ""),
            "WholeSale": r.get("WHOLESALE", ""),
            "MSRP": r.get("MSRP", ""),
        })
    return pd.DataFrame(rows, columns=["Styles", "Size", "CONCAT", "UPC", "DESCRIPTION", "WholeSale", "MSRP"])


HEADER_FILL = PatternFill("solid", fgColor="006699")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY_FMT   = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'


def style_header(ws):
    for cell in ws[1]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def add_excel_table(ws, table_name):
    """Add a real Excel Table so [@ColumnName] structured references work."""
    max_col = get_column_letter(ws.max_column)
    max_row = ws.max_row
    tab = Table(displayName=table_name, ref=f"A1:{max_col}{max_row}")
    # No built-in style so our custom blue header colours are preserved
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=False, showColumnStripes=False,
    )
    ws.add_table(tab)


def autofit(ws):
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 10), 50)


def build_output_xlsx(upc_df, base):
    wb = Workbook()
    styles_list = sorted(base["Style"].dropna().unique().tolist())

    # ── Styles ──
    ws_styles = wb.active
    ws_styles.title = "Styles"
    ws_styles.append(["Styles"])
    for s in styles_list:
        ws_styles.append([s])
    style_header(ws_styles)
    autofit(ws_styles)
    add_excel_table(ws_styles, "StylesList")

    # ── UPC (fórmulas idénticas al manual) ──
    ws_upc = wb.create_sheet("UPC")
    ws_upc.append(["Styles", "Size", "CONCAT", "UPC", "DESCRIPTION", "WholeSale", "MSRP"])
    for i, (_, r) in enumerate(upc_df.iterrows(), start=2):
        ws_upc.cell(row=i, column=1, value=r.get("Styles", ""))
        ws_upc.cell(row=i, column=2, value=r.get("Size", ""))
        ws_upc.cell(row=i, column=3, value=f"=A{i}&B{i}")
        cell_upc = ws_upc.cell(row=i, column=4, value=str(r.get("UPC", "")))
        cell_upc.number_format = "@"
        ws_upc.cell(row=i, column=5, value=f'=IFERROR(VLOOKUP(A{i},BASE!$A:$F,4,0),"N/A")')
        ws_upc.cell(row=i, column=6, value=f'=IFERROR(VLOOKUP(A{i},BASE!$A:$F,5,0),"N/A")')
        ws_upc.cell(row=i, column=6).number_format = MONEY_FMT
        ws_upc.cell(row=i, column=7, value=f'=IFERROR(VLOOKUP(A{i},BASE!$A:$F,6,0),"N/A")')
        ws_upc.cell(row=i, column=7).number_format = MONEY_FMT
    style_header(ws_upc)
    autofit(ws_upc)
    add_excel_table(ws_upc, "UPCData")

    # ── BASE ──
    ws_base = wb.create_sheet("BASE")
    for row in dataframe_to_rows(base, index=False, header=True):
        ws_base.append(row)
    style_header(ws_base)
    add_excel_table(ws_base, "BASEData")
    for cell in ws_base["C"][1:]:
        cell.number_format = "@"
    for row in ws_base.iter_rows(min_row=2, max_row=ws_base.max_row):
        for cell in row:
            if cell.column in (5, 6):
                cell.number_format = MONEY_FMT
    autofit(ws_base)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


db_df = st.session_state.get("db_df", None)
has_data = db_df is not None or (uploaded and len(uploaded) > 0)

if has_data:
    if st.button("GENERAR UPCs", type="primary", use_container_width=True):
        st.session_state.pop("output", None)
        with st.spinner("Procesando... (puede tardar 1-2 min con muchos datos)"):
            if db_df is not None:
                base_df = build_base_from_df(db_df, brands=selected_brands)
            else:
                base_df = build_base(list(uploaded) if uploaded else [])

        if base_df.empty:
            st.error("No se encontraron datos. Verifica que el archivo sea el correcto.")
        else:
            upc_df = build_upc_sheet(base_df)
            xlsx_bytes = build_output_xlsx(upc_df, base_df)
            fname = f"UPCS_{file_date.strftime('%m.%d.%Y')}.xlsx"

            # Guardar directo en disco (confiable aunque sea pesado)
            out_dir = os.path.join(os.getcwd(), "UPCs_generados")
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, fname)
            try:
                with open(out_path, "wb") as fh:
                    fh.write(xlsx_bytes)
                saved_path = out_path
            except Exception:
                saved_path = None

            st.session_state["output"] = {
                "fname": fname,
                "bytes": xlsx_bytes,
                "count": len(base_df),
                "path": saved_path,
            }

# Mostrar resultado (persiste aunque la app se recargue)
out = st.session_state.get("output")
if out:
    st.markdown(
        f'<div class="success-box">Listo — {out["count"]:,} registros procesados</div>',
        unsafe_allow_html=True,
    )
    if out.get("path"):
        st.success(f"✅ Guardado automáticamente en:\n\n`{out['path']}`")
    st.download_button(
        label=f"⬇ Descargar {out['fname']}",
        data=out["bytes"],
        file_name=out["fname"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
elif not has_data:
    st.info("Conecta a la base de datos o sube un archivo de Power BI para continuar.")
else:
    st.info("Conecta a la base de datos o sube un archivo de Power BI para continuar.")
