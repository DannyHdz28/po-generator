import io
import re
from pathlib import Path

import streamlit as st

# ─── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(page_title="Image Search — Pro Standard", page_icon="🔍", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Bebas+Neue&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main-title { font-family:'Bebas Neue',sans-serif; font-size:2.2rem;
              letter-spacing:5px; color:#111827; margin-bottom:0; }
.sub-title  { font-size:0.72rem; letter-spacing:3px; color:#6b7280;
              text-transform:uppercase; margin-top:2px; }
.section-h  { font-size:0.7rem; font-weight:600; letter-spacing:2px;
              text-transform:uppercase; color:#6b7280;
              border-bottom:1px solid #e5e7eb; padding-bottom:6px;
              margin-top:1.8rem; margin-bottom:0.8rem; }
.pill-ok    { display:inline-block; background:#d1fae5; border:1px solid #6ee7b7;
              color:#065f46; padding:3px 10px; border-radius:4px;
              font-size:0.72rem; font-weight:600; letter-spacing:0.5px; }
.pill-warn  { display:inline-block; background:#fee2e2; border:1px solid #fca5a5;
              color:#991b1b; padding:3px 10px; border-radius:4px;
              font-size:0.72rem; font-weight:600; letter-spacing:0.5px; }
.pill-info  { display:inline-block; background:#eff6ff; border:1px solid #bfdbfe;
              color:#1e40af; padding:3px 10px; border-radius:4px;
              font-size:0.72rem; font-weight:600; letter-spacing:0.5px; }
.img-label  { font-size:0.7rem; color:#6b7280; text-align:center;
              margin-top:4px; word-break:break-all; }
.img-path   { font-size:0.65rem; color:#9ca3af; text-align:center; }

section[data-testid="stSidebar"] { background:#1e293b !important; }
section[data-testid="stSidebar"] * { color:#e2e8f0 !important; }
section[data-testid="stSidebar"] input { background:#334155 !important;
                                          border-color:#475569 !important; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────
SERVER_DEFAULT = r"\\10.0.1.30\Design\DESIGN\Pro Standard"
MAX_RESULTS    = 100   # límite para no colgar la UI

# ─── HELPERS ──────────────────────────────────────────────────
@st.cache_data(ttl=120, show_spinner=False)
def search_images(base: str, query: str) -> list[str]:
    """Busca recursivamente PNGs cuyo nombre contenga `query`."""
    base_path = Path(base)
    pattern   = f"*{query.upper()}*.png"
    try:
        results = sorted(
            str(p) for p in base_path.rglob(pattern)
            if p.is_file()
        )
        return results[:MAX_RESULTS]
    except PermissionError:
        return []
    except Exception:
        return []

def read_image(path: str) -> bytes | None:
    try:
        return Path(path).read_bytes()
    except Exception:
        return None

def short_path(path: str, base: str) -> str:
    """Muestra solo la ruta relativa al base."""
    try:
        return str(Path(path).relative_to(Path(base)))
    except Exception:
        return path

# ─── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### CONFIG")
    server_base = st.text_input(
        "Ruta del servidor",
        value=SERVER_DEFAULT,
        help=r"Ruta UNC, ej: \\10.0.1.30\Design\DESIGN\Pro Standard",
    )
    st.markdown("---")
    st.markdown("### FILTROS")
    only_pdp = st.checkbox("Solo carpeta PDP", value=False,
                            help="Muestra solo imágenes dentro de carpetas llamadas PDP")
    cols_count = st.slider("Columnas", min_value=2, max_value=6, value=4)

# ─── HEADER ───────────────────────────────────────────────────
st.markdown('<div class="main-title">IMAGE SEARCH</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">PRO STANDARD · BUSCADOR DE IMÁGENES</div>', unsafe_allow_html=True)
st.write("")

# ─── VERIFICAR SERVIDOR ───────────────────────────────────────
base_path = Path(server_base)
if not base_path.exists():
    st.error(f"No se puede acceder al servidor: `{server_base}`\n\nVerifica que estés conectado a la red.")
    st.stop()

# ─── BÚSQUEDA ─────────────────────────────────────────────────
st.markdown('<div class="section-h">Buscar por estilo</div>', unsafe_allow_html=True)

col_input, col_btn = st.columns([4, 1])
with col_input:
    query = st.text_input(
        "Número de estilo",
        placeholder="Ej: HBS567746",
        label_visibility="collapsed",
    ).strip()
with col_btn:
    buscar = st.button("Buscar →", type="primary", use_container_width=True)

# ─── BÚSQUEDA MÚLTIPLE (batch) ────────────────────────────────
with st.expander("Buscar varios estilos a la vez", expanded=False):
    batch_input = st.text_area(
        "Un estilo por línea",
        placeholder="HBS567746\nBDP5516969\nBDP6516978",
        height=120,
        label_visibility="collapsed",
    )
    batch_btn = st.button("Buscar todos →", key="batch_btn")

# ─── RESULTADOS ───────────────────────────────────────────────
def show_results(results: list[str], label: str):
    if only_pdp:
        results = [r for r in results if "PDP" in Path(r).parts]

    if not results:
        st.markdown('<span class="pill-warn">⚠ Sin resultados</span>', unsafe_allow_html=True)
        st.caption("Verifica el número de estilo o que el servidor tenga imágenes para ese estilo.")
        return

    count = len(results)
    badge = f'<span class="pill-ok">✓ {count} IMAGEN{"ES" if count != 1 else ""} · {label}</span>'
    if count == MAX_RESULTS:
        badge += f' &nbsp; <span class="pill-info">mostrando primeros {MAX_RESULTS}</span>'
    st.markdown(badge, unsafe_allow_html=True)
    st.write("")

    cols = st.columns(cols_count)
    for i, img_path in enumerate(results):
        with cols[i % cols_count]:
            img_bytes = read_image(img_path)
            if img_bytes:
                st.image(img_bytes, use_container_width=True)
            else:
                st.markdown("_(sin acceso)_")
            st.markdown(
                f'<div class="img-label">{Path(img_path).name}</div>'
                f'<div class="img-path">{short_path(img_path, server_base)}</div>',
                unsafe_allow_html=True,
            )
            # Botón copiar path completo
            st.code(img_path, language=None)


if buscar and query:
    with st.spinner(f"Buscando {query.upper()} en el servidor..."):
        results = search_images(server_base, query.upper())
    show_results(results, query.upper())

elif batch_btn and batch_input.strip():
    styles = [s.strip().upper() for s in batch_input.strip().splitlines() if s.strip()]
    st.markdown(f'<div class="section-h">Resultados — {len(styles)} estilo(s)</div>', unsafe_allow_html=True)
    for style in styles:
        st.markdown(f"**{style}**")
        with st.spinner(f"Buscando {style}..."):
            results = search_images(server_base, style)
        show_results(results, style)
        st.divider()
