import streamlit as st
import pandas as pd
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="UPCs Generator", page_icon="🏷️", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main-title{font-family:'Bebas Neue',sans-serif;font-size:3rem;letter-spacing:4px;color:#e8c84a;}
.sub-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:3px;color:#888;}
.warn-box{background:rgba(255,165,0,0.1);border:1px solid rgba(255,165,0,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#ffa500;}
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#4ade80;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">UPCs GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MAXIMA APPAREL</div>', unsafe_allow_html=True)
st.markdown("---")

TARGET_COLS = ["IVNUM", "SIZE", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"]


def read_pbi_export(file) -> pd.DataFrame:
    df = pd.read_excel(file, header=2, dtype={"UPC": str})
    df = df[[c for c in TARGET_COLS if c in df.columns]].copy()
    df["UPC"] = df["UPC"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(12)
    return df


def build_base(files) -> pd.DataFrame:
    frames = []
    for f in files:
        try:
            frames.append(read_pbi_export(f))
        except Exception as e:
            st.error(f"No pude leer {f.name}: {e}")
    if not frames:
        return pd.DataFrame(columns=TARGET_COLS)
    base = pd.concat(frames, ignore_index=True)
    base = base.drop_duplicates(subset=["IVNUM", "SIZE"], keep="first").reset_index(drop=True)
    base = base.rename(columns={"IVNUM": "Style", "SIZE": "Size"})
    return base[["Style", "Size", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"]]


def build_upc_sheet(base: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in base.iterrows():
        rows.append({
            "Styles": r["Style"],
            "Size": r["Size"],
            "CONCAT": f"{r['Style']}{r['Size']}",
            "UPC": r["UPC"],
            "DESCRIPTION": r["DESCRIPTION"],
            "WholeSale": r["WHOLESALE"],
            "MSRP": r["MSRP"],
        })
    return pd.DataFrame(rows, columns=["Styles", "Size", "CONCAT", "UPC", "DESCRIPTION", "WholeSale", "MSRP"])


def build_output_xlsx(upc_df: pd.DataFrame, base: pd.DataFrame) -> bytes:
    wb = Workbook()

    styles_list = sorted(base["Style"].unique().tolist())
    ws_styles = wb.active
    ws_styles.title = "Styles"
    ws_styles.append(["Styles"])
    for s in styles_list:
        ws_styles.append([s])

    ws_upc = wb.create_sheet("UPC")
    for row in dataframe_to_rows(upc_df, index=False, header=True):
        ws_upc.append(row)
    for cell in ws_upc["D"][1:]:
        cell.number_format = "@"

    base_out = base.copy()
    ws_base = wb.create_sheet("BASE")
    for row in dataframe_to_rows(base_out, index=False, header=True):
        ws_base.append(row)
    for cell in ws_base["C"][1:]:
        cell.number_format = "@"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── UI ────────────────────────────────────────────────────
file_date = st.date_input("Fecha", value=date.today())

# Archivos PBI ocultos en expander (temporal hasta tener conexión automática)
with st.expander("Archivos de Power BI", expanded=True):
    uploaded = st.file_uploader(
        "Sube los archivos descargados de PBI",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

base_df = None
if uploaded:
    base_df = build_base(uploaded)
    st.markdown(f'<div class="success-box">{len(base_df):,} registros listos</div>',
                unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
generate = st.button("GENERAR UPCs", type="primary", use_container_width=True,
                     disabled=base_df is None)

if generate:
    with st.spinner("Generando archivo..."):
        upc_df = build_upc_sheet(base_df)
        xlsx_bytes = build_output_xlsx(upc_df, base_df)

    st.markdown('<div class="success-box">Archivo listo!</div>', unsafe_allow_html=True)

    fname = f"UPCS_{file_date.strftime('%m.%d.%Y')}.xlsx"
    st.download_button(
        label=f"Descargar {fname}",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
