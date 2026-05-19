import streamlit as st
import pandas as pd
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="UPCs Generator", page_icon="🏷️", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.main-title{font-family:'Bebas Neue',sans-serif;font-size:3rem;letter-spacing:4px;color:#e8c84a;}
.sub-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:3px;color:#888;}
.section-title{font-family:'Bebas Neue',sans-serif;font-size:1.1rem;letter-spacing:2px;color:#aaa;margin-top:1rem;}
.warn-box{background:rgba(255,165,0,0.1);border:1px solid rgba(255,165,0,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#ffa500;}
.success-box{background:rgba(74,222,128,0.1);border:1px solid rgba(74,222,128,0.4);border-radius:6px;padding:10px 14px;font-size:0.85rem;color:#4ade80;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">UPCs GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MAXIMA APPAREL · STYLE UPC AUTOMATION</div>', unsafe_allow_html=True)

TARGET_COLS = ["IVNUM", "SIZE", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"]


def read_pbi_export(file) -> pd.DataFrame:
    """Read one Power BI export. Header is on row 3 (row 1 = filter banner, row 2 = blank)."""
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
    base = base.rename(columns={"IVNUM": "Style", "SIZE": "Size", "WHOLESALE": "WHOLESALE", "MSRP": "MSRP"})
    return base[["Style", "Size", "UPC", "DESCRIPTION", "WHOLESALE", "MSRP"]]


def parse_styles_input(text: str) -> list[str]:
    return [s.strip() for s in text.replace(",", "\n").splitlines() if s.strip()]


def build_upc_sheet(styles: list[str], base: pd.DataFrame) -> pd.DataFrame:
    base["CONCAT"] = base["Style"].astype(str) + base["Size"].astype(str)
    rows = []
    for style in styles:
        matches = base[base["Style"] == style]
        if matches.empty:
            rows.append({"Styles": style, "Size": "", "CONCAT": "", "UPC": "NOT FOUND",
                         "DESCRIPTION": "", "WholeSale": "", "MSRP": ""})
            continue
        for _, r in matches.iterrows():
            rows.append({
                "Styles": style,
                "Size": r["Size"],
                "CONCAT": f"{style}{r['Size']}",
                "UPC": r["UPC"],
                "DESCRIPTION": r["DESCRIPTION"],
                "WholeSale": r["WHOLESALE"],
                "MSRP": r["MSRP"],
            })
    return pd.DataFrame(rows, columns=["Styles", "Size", "CONCAT", "UPC", "DESCRIPTION", "WholeSale", "MSRP"])


def build_output_xlsx(styles: list[str], upc_df: pd.DataFrame, base: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws_styles = wb.active
    ws_styles.title = "Styles"
    ws_styles.append(["Styles"])
    for s in styles:
        ws_styles.append([s])

    ws_upc = wb.create_sheet("UPC")
    for row in dataframe_to_rows(upc_df, index=False, header=True):
        ws_upc.append(row)
    for cell in ws_upc["D"][1:]:
        cell.number_format = "@"

    base_out = base.drop(columns=[c for c in ["CONCAT"] if c in base.columns])
    ws_base = wb.create_sheet("BASE")
    for row in dataframe_to_rows(base_out, index=False, header=True):
        ws_base.append(row)
    for cell in ws_base["C"][1:]:
        cell.number_format = "@"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── Power BI auto-download (placeholder) ──────────────────
st.markdown('<div class="section-title">1 · POWER BI</div>', unsafe_allow_html=True)
with st.expander("Descargar automáticamente desde Power BI (requiere credenciales)"):
    st.markdown("""
<div class="warn-box">
Esta sección queda lista para activarse cuando tengas credenciales de Azure AD (Tenant ID,
Client ID, Client Secret) o usuario/password de Power BI sin MFA. Mientras tanto, usa la
sección de abajo para subir los archivos descargados manualmente.
</div>
""", unsafe_allow_html=True)
    st.text_input("Tenant ID", key="pbi_tenant", disabled=True, placeholder="pendiente de credenciales")
    st.text_input("Workspace ID", key="pbi_ws", disabled=True, placeholder="pendiente de URL del reporte")
    st.text_input("Report ID", key="pbi_report", disabled=True, placeholder="pendiente de URL del reporte")
    st.button("Descargar 7 tallas desde Power BI", disabled=True)

# ─── Manual upload ─────────────────────────────────────────
st.markdown('<div class="section-title">2 · SUBIR EXCELS DE POWER BI</div>', unsafe_allow_html=True)
st.caption("Arrastra los 7 archivos descargados de PBI (uno por talla) o los que tengas a la mano.")
uploaded = st.file_uploader(
    "Archivos data_XX.xlsx",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

base_df = None
if uploaded:
    base_df = build_base(uploaded)
    st.markdown(f'<div class="success-box">BASE consolidada: {len(base_df):,} filas únicas '
                f'(de {len(uploaded)} archivos)</div>', unsafe_allow_html=True)
    with st.expander("Ver vista previa de BASE"):
        st.dataframe(base_df.head(50), use_container_width=True)

# ─── Styles input ──────────────────────────────────────────
st.markdown('<div class="section-title">3 · ESTILOS A BUSCAR</div>', unsafe_allow_html=True)
styles_text = st.text_area(
    "Pega la lista de Styles (uno por línea, formato IVNUM como LPHT1315927-RYB)",
    height=180,
    placeholder="LPHT1315927-RYB\nLNYT1316089-MDN\nFPEL7411912-BG",
)
styles = parse_styles_input(styles_text)
if styles:
    st.caption(f"{len(styles)} estilos detectados")

# ─── Generate ──────────────────────────────────────────────
st.markdown('<div class="section-title">4 · GENERAR ARCHIVO</div>', unsafe_allow_html=True)
col1, col2 = st.columns([1, 3])
with col1:
    file_date = st.date_input("Fecha del archivo", value=date.today())
generate = st.button("Generar UPCs", type="primary", use_container_width=True,
                     disabled=not (base_df is not None and styles))

if generate:
    upc_df = build_upc_sheet(styles, base_df)
    missing = upc_df[upc_df["UPC"] == "NOT FOUND"]["Styles"].unique().tolist()
    xlsx_bytes = build_output_xlsx(styles, upc_df, base_df)

    if missing:
        st.markdown(f'<div class="warn-box">Sin UPC en BASE: {", ".join(missing)}</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="success-box">Todos los estilos encontrados en BASE</div>',
                    unsafe_allow_html=True)

    st.dataframe(upc_df, use_container_width=True)

    fname = f"UPCS_{file_date.strftime('%m.%d.%Y')}.xlsx"
    st.download_button(
        label=f"Descargar {fname}",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
