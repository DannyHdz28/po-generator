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
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#4ade80;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">UPCs GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MAXIMA APPAREL</div>', unsafe_allow_html=True)
st.markdown("---")

SIZES = ["2T", "3T", "4", "4T", "5", "6", "7", "OS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]


def build_base(files):
    all_rows = []
    for f in files:
        try:
            df = pd.read_excel(f, header=2, dtype=str)
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


def build_output_xlsx(upc_df, base):
    wb = Workbook()
    styles_list = sorted(base["Style"].dropna().unique().tolist())

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


file_date = st.date_input("Fecha", value=date.today())
st.markdown("<br>", unsafe_allow_html=True)

generate = st.button("GENERAR UPCs", type="primary", use_container_width=True)

if generate:
    status = st.empty()
    progress = st.progress(0)

    def update(msg):
        status.info(msg)

    update("Iniciando descarga desde Power BI...")

    try:
        from pbi_downloader import run_download
        files = run_download(progress_fn=update)
        progress.progress(60)

        update("Procesando archivos...")
        base_df = build_base(files)
        progress.progress(80)

        if base_df.empty:
            st.error("No se encontraron datos en los archivos descargados.")
        else:
            update("Generando archivo final...")
            upc_df = build_upc_sheet(base_df)
            xlsx_bytes = build_output_xlsx(upc_df, base_df)
            progress.progress(100)

            st.markdown(f'<div class="success-box">Listo — {len(base_df):,} registros procesados</div>',
                        unsafe_allow_html=True)

            fname = f"UPCS_{file_date.strftime('%m.%d.%Y')}.xlsx"
            st.download_button(
                label=f"Descargar {fname}",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Verifica que tengas conexion a la red de Maxima y que Power BI este accesible.")
