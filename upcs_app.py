import streamlit as st
import pandas as pd
import io
import threading
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
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#4ade80;}
.warn-box{background:rgba(255,165,0,0.1);border:1px solid rgba(255,165,0,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#ffa500;}
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
        except Exception:
            pass
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

    ws_base = wb.create_sheet("BASE")
    for row in dataframe_to_rows(base, index=False, header=True):
        ws_base.append(row)
    for cell in ws_base["C"][1:]:
        cell.number_format = "@"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── UI ────────────────────────────────────────────────────────
file_date = st.date_input("Fecha", value=date.today())
st.markdown("<br>", unsafe_allow_html=True)

generate = st.button("GENERAR UPCs", type="primary", use_container_width=True)

if generate:
    status = st.empty()
    progress = st.progress(0)
    messages = []

    def update(msg):
        messages.append(msg)
        status.info(messages[-1])

    update("Iniciando descarga desde Power BI...")

    try:
        from pbi_downloader import run_download
        files = run_download(progress_fn=update)
        progress.progress(60)

        update("Procesando archivos...")
        base_df = build_base(files)
        progress.progress(80)

        update("Generando archivo final...")
        upc_df = build_upc_sheet(base_df)
        xlsx_bytes = build_output_xlsx(upc_df, base_df)
        progress.progress(100)

        st.markdown(f'<div class="success-box">Listo — {len(base_df):,} registros procesados</div>',
                    unsafe_allow_html=True)

        fname = f"UPCS_{file_date.strftime('%m.%d.%Y')}.xlsx"
        st.download_button(
            label=f"⬇ Descargar {fname}",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Verifica que tengas conexión a la red de Maxima y que Power BI esté accesible.")
