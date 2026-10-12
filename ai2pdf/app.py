import io
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st
from pypdf import PdfReader

# ─── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(page_title="AI → PDF — Pro Standard", page_icon="🎨", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.main-title  { font-family: 'Bebas Neue', sans-serif; font-size: 3rem;   letter-spacing: 4px; color: #e8c84a; margin-bottom: 0; }
.sub-title   { font-family: 'Bebas Neue', sans-serif; font-size: 1.1rem; letter-spacing: 3px; color: #888;     margin-top: 0; }
.section-h   { font-family: 'Bebas Neue', sans-serif; font-size: 1.4rem; letter-spacing: 2px; color: #f5f5f0;  margin-top: 1.5rem; }
.pill-ok     { background: rgba(74,222,128,0.12); border:1px solid rgba(74,222,128,0.4); color:#4ade80; padding:2px 10px; border-radius:3px; font-size:0.75rem; letter-spacing:1px; }
.pill-warn   { background: rgba(248,113,113,0.12); border:1px solid rgba(248,113,113,0.4); color:#f87171; padding:2px 10px; border-radius:3px; font-size:0.75rem; letter-spacing:1px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────
# Ruta base por defecto del servidor de diseño
SERVER_BASE_DEFAULT = r"\\10.0.1.30\Design\DESIGN\Pro Standard"

# Nombre de cada nivel de carpeta debajo de la ruta base.
# Ej: 2027 \ Q3 2027 \ Q3 ELITE TEAM \ MLB \ HOUSTON ASTROS
LEVEL_LABELS = ["Año", "Quarter", "Cápsula", "Liga", "Equipo"]

AI_EXTS      = {".ai"}
NONE_OPTION  = "—"

# Carpetas con versiones viejas que se ignoran (OLD, _OLD, OLD FILES, OLDS...)
OLD_DIR_RE   = re.compile(r"^[_\-\s]*OLD(S)?\b", re.IGNORECASE)
MAX_LIST     = 1000  # tope de archivos a listar en búsqueda recursiva

# Texto que Illustrator pone en la página "placeholder" cuando el .ai
# se guardó SIN "Create PDF Compatible File"
NO_PDF_MARKERS = ("saved without pdf content", "without pdf content")

# Precios tipo $45 · $45.00 · $ 1,299.99 · USD 45
PRICE_RE      = re.compile(r"(?:\$|USD)\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?")
MIN_TEXT_CHARS = 10  # menos que esto = prácticamente sin texto (contornos)


# ─── HELPERS ──────────────────────────────────────────────────
def level_label(i: int) -> str:
    return LEVEL_LABELS[i] if i < len(LEVEL_LABELS) else f"Subcarpeta {i + 1}"


def clean_name(name: str) -> str:
    """Para mostrar: quita el '_' inicial de carpetas como '_DALLAS COWBOYS'."""
    return name.lstrip("_").strip() or name


def is_old_dir(name: str) -> bool:
    return bool(OLD_DIR_RE.match(name))


@st.cache_data(ttl=300, show_spinner=False)
def list_dirs(path_str: str, skip_old: bool = True) -> list[str]:
    """Subcarpetas ordenadas (ignora ocultas, del sistema y, opcional, las OLD)."""
    with os.scandir(path_str) as it:
        names = [
            e.name for e in it
            if e.is_dir() and not e.name.startswith((".", "~", "$"))
            and not (skip_old and is_old_dir(e.name))
        ]
    return sorted(names, key=lambda n: clean_name(n).upper())


@st.cache_data(ttl=300, show_spinner=False)
def list_ai_files(path_str: str, recursive: bool, skip_old: bool = True) -> tuple[list[dict], bool]:
    """Lista los .ai de la carpeta. Regresa (archivos, se_truncó)."""
    root  = Path(path_str)
    found = []

    def walk(folder: str):
        try:
            with os.scandir(folder) as it:
                entries = sorted(it, key=lambda e: e.name.upper())
        except (PermissionError, FileNotFoundError, OSError):
            return
        for e in entries:
            if len(found) >= MAX_LIST:
                return
            if e.is_file() and Path(e.name).suffix.lower() in AI_EXTS and not e.name.startswith(("._", "~")):
                st_ = e.stat()
                found.append({
                    "path":     e.path,
                    "name":     e.name,
                    "subdir":   Path(e.path).parent.relative_to(root).as_posix() if Path(e.path).parent != root else "",
                    "size_mb":  round(st_.st_size / 1_048_576, 1),
                    "modified": datetime.fromtimestamp(st_.st_mtime).strftime("%Y-%m-%d %H:%M"),
                })
            elif (recursive and e.is_dir() and not e.name.startswith((".", "~", "$"))
                  and not (skip_old and is_old_dir(e.name))):
                walk(e.path)

    walk(path_str)
    return found, len(found) >= MAX_LIST


def ai_to_pdf(path_str: str) -> tuple[bytes | None, str]:
    """
    Convierte un .ai a PDF.
    Un .ai guardado con "Create PDF Compatible File" (default de Illustrator)
    ya ES un PDF por dentro, así que se valida y se regresa tal cual.
    Regresa (bytes_pdf | None, mensaje).
    """
    try:
        data = Path(path_str).read_bytes()
    except Exception as e:
        return None, f"No se pudo leer el archivo: {e}"

    head = data[:1024]
    if b"%PDF" not in head:
        if head.startswith(b"%!PS"):
            return None, "AI antiguo (PostScript) — exportar a PDF desde Illustrator"
        return None, "No es un PDF compatible — exportar a PDF desde Illustrator"

    try:
        reader = PdfReader(io.BytesIO(data))
        n_pages = len(reader.pages)
        if n_pages == 0:
            return None, "El PDF interno no tiene páginas"
        first_text = (reader.pages[0].extract_text() or "").lower()
        if any(m in first_text for m in NO_PDF_MARKERS):
            return None, "Guardado SIN 'PDF Compatible File' — exportar a PDF desde Illustrator"
    except Exception as e:
        return None, f"PDF interno dañado: {e}"

    return data, f"OK · {n_pages} página(s)"


def analyze_text(pdf_bytes: bytes) -> dict:
    """
    Revisa si el PDF tiene texto real (editable) o si el texto se convirtió
    a contornos. También detecta los precios que encuentre.
    """
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text   = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception as e:
        return {"status": "error", "msg": f"No se pudo leer el texto: {e}", "prices": []}

    chars  = len(re.sub(r"\s+", "", text))
    prices = list(dict.fromkeys(m.group(0).replace(" ", "") for m in PRICE_RE.finditer(text)))

    if chars < MIN_TEXT_CHARS:
        return {"status": "outlines", "prices": [],
                "msg": "SIN TEXTO — el texto está en contornos (no editable)"}
    if prices:
        shown = ", ".join(prices[:8]) + (" …" if len(prices) > 8 else "")
        return {"status": "editable", "prices": prices,
                "msg": f"TEXTO EDITABLE · {len(prices)} precio(s): {shown}"}
    return {"status": "no_prices", "prices": [],
            "msg": "Texto editable, pero NO se encontraron precios como texto (¿precios en contornos?)"}


def pdf_name(file: dict) -> str:
    """Nombre de salida: misma ruta relativa, extensión .pdf"""
    base = re.sub(r"\.ai$", ".pdf", file["name"], flags=re.IGNORECASE)
    return f"{file['subdir']}/{base}" if file["subdir"] else base


# ═══════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="main-title">AI → PDF</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">PRO STANDARD · DESIGN SERVER EXPORT</div>', unsafe_allow_html=True)
st.write("")

with st.expander("Cómo funciona", expanded=False):
    st.markdown("""
1. Navega las carpetas del servidor de diseño: **Año → Quarter → Cápsula → Liga → Equipo**.
2. La app lista los archivos de Illustrator (`.ai`) de esa carpeta.
3. Selecciona los que quieras y da clic en **Convertir a PDF**. Se descargan ya como `.pdf`.

Los `.ai` guardados con **"Create PDF Compatible File"** (opción por default de Illustrator)
se convierten directo. Si alguno se guardó sin esa opción, aparece marcado y hay que
exportarlo a PDF desde Illustrator.
    """)

# ─── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### SERVIDOR")
    server_base = st.text_input(
        "Ruta base",
        value=SERVER_BASE_DEFAULT,
        help="Ruta UNC al servidor de diseño, ej: \\\\10.0.1.30\\Design\\DESIGN\\Pro Standard",
    )
    recursive = st.checkbox(
        "Incluir subcarpetas",
        value=True,
        help="Busca .ai también dentro de las subcarpetas de la carpeta seleccionada",
    )
    skip_old = st.checkbox(
        "Ignorar carpetas OLD",
        value=True,
        help="No muestra ni busca dentro de carpetas OLD / _OLD / OLD FILES (versiones viejas)",
    )
    if st.button("↻ Re-scan"):
        st.cache_data.clear()
        st.session_state.pop("results", None)
        st.rerun()

# ─── STEP 1: NAVEGAR ──────────────────────────────────────────
st.markdown('<div class="section-h">1 · SELECCIONA LA CARPETA</div>', unsafe_allow_html=True)

base_path = Path(server_base)
if not base_path.is_dir():
    st.error(f"No se puede acceder a la ruta: `{server_base}`\n\nVerifica que estés conectado a la red de la empresa.")
    st.stop()

current = base_path
level   = 0
cols    = st.columns(len(LEVEL_LABELS))

while True:
    try:
        subdirs = list_dirs(str(current), skip_old)
    except PermissionError:
        st.error(f"Sin permisos para leer: {current}")
        break
    except Exception as e:
        st.error(f"Error leyendo carpeta: {e}")
        break
    if not subdirs:
        break

    col = cols[level % len(cols)]
    with col:
        choice = st.selectbox(
            level_label(level),
            [NONE_OPTION] + subdirs,
            format_func=lambda n: n if n == NONE_OPTION else clean_name(n),
            key=f"lvl_{level}_{current}",
        )
    if choice == NONE_OPTION:
        break
    current = current / choice
    level  += 1

st.caption(f"📂 `{current}`")

# Si cambia la carpeta, se descartan resultados anteriores
if st.session_state.get("results_dir") != str(current):
    st.session_state.pop("results", None)

# ─── STEP 2: ARCHIVOS .AI ─────────────────────────────────────
st.markdown('<div class="section-h">2 · ARCHIVOS ILLUSTRATOR</div>', unsafe_allow_html=True)

if level == 0:
    st.caption("Selecciona al menos el año para buscar archivos.")
    st.stop()

with st.spinner("Buscando archivos .ai..."):
    files, truncated = list_ai_files(str(current), recursive, skip_old)

if not files:
    st.markdown('<span class="pill-warn">⚠ SIN ARCHIVOS .AI</span>', unsafe_allow_html=True)
    st.caption("No hay archivos de Illustrator en esta carpeta" + (" ni en sus subcarpetas." if recursive else "."))
    st.stop()

st.markdown(f'<span class="pill-ok">✓ {len(files)} ARCHIVO(S) .AI</span>', unsafe_allow_html=True)
if truncated:
    st.warning(f"Se muestran solo los primeros {MAX_LIST}. Baja un nivel más (ej. hasta Equipo) para ver todos.")

select_all = st.checkbox("Seleccionar todos", value=True)
table = st.data_editor(
    [
        {
            "Descargar":  select_all,
            "Archivo":    f["name"],
            "Subcarpeta": f["subdir"] or "—",
            "MB":         f["size_mb"],
            "Modificado": f["modified"],
        }
        for f in files
    ],
    hide_index=True,
    width="stretch",
    disabled=["Archivo", "Subcarpeta", "MB", "Modificado"],
    key=f"table_{current}_{select_all}",
)
selected = [f for f, row in zip(files, table) if row["Descargar"]]

# ─── STEP 3: CONVERTIR Y DESCARGAR ────────────────────────────
st.markdown('<div class="section-h">3 · CONVERTIR A PDF</div>', unsafe_allow_html=True)

if not selected:
    st.caption("Selecciona al menos un archivo.")
    st.stop()

total_mb = sum(f["size_mb"] for f in selected)
st.caption(f"{len(selected)} archivo(s) seleccionados · {total_mb:.1f} MB")

if st.button("Convertir a PDF →", type="primary"):
    results  = []
    progress = st.progress(0.0, text="Convirtiendo...")
    for i, f in enumerate(selected, start=1):
        pdf_bytes, msg = ai_to_pdf(f["path"])
        text_info = analyze_text(pdf_bytes) if pdf_bytes else None
        results.append({"file": f, "pdf": pdf_bytes, "msg": msg, "text": text_info})
        progress.progress(i / len(selected), text=f"Convirtiendo {i}/{len(selected)} · {f['name']}")
    progress.empty()
    st.session_state["results"]     = results
    st.session_state["results_dir"] = str(current)

results = st.session_state.get("results")
if results:
    ok   = [r for r in results if r["pdf"]]
    fail = [r for r in results if not r["pdf"]]

    if ok:
        st.markdown(f'<span class="pill-ok">✓ {len(ok)} PDF(S) LISTOS</span>', unsafe_allow_html=True)
    if fail:
        st.markdown(f'<span class="pill-warn">✗ {len(fail)} NO SE PUDIERON CONVERTIR</span>', unsafe_allow_html=True)

    # Resumen de texto editable vs contornos
    editable = sum(1 for r in ok if r["text"]["status"] == "editable")
    no_text  = sum(1 for r in ok if r["text"]["status"] in ("outlines", "no_prices", "error"))
    if ok:
        st.write("")
        if editable:
            st.markdown(f'<span class="pill-ok">✓ {editable} CON PRECIOS EDITABLES</span>', unsafe_allow_html=True)
        if no_text:
            st.markdown(f'<span class="pill-warn">⚠ {no_text} SIN PRECIOS COMO TEXTO</span>', unsafe_allow_html=True)

    with st.expander("Detalle", expanded=True):
        text_colors = {"editable": "#4ade80", "no_prices": "#e8c84a", "outlines": "#f87171", "error": "#f87171"}
        for r in results:
            color = "#4ade80" if r["pdf"] else "#f87171"
            icon  = "✓" if r["pdf"] else "✗"
            st.markdown(
                f'<div style="font-size:0.8rem;color:{color};">{icon} {pdf_name(r["file"])} — {r["msg"]}</div>',
                unsafe_allow_html=True,
            )
            if r["text"]:
                st.markdown(
                    f'<div style="font-size:0.75rem;color:{text_colors[r["text"]["status"]]};'
                    f'margin:0 0 6px 18px;">↳ {r["text"]["msg"]}</div>',
                    unsafe_allow_html=True,
                )

    folder_name = clean_name(current.name)
    if len(ok) == 1:
        st.download_button(
            "⬇ Descargar PDF",
            data=ok[0]["pdf"],
            file_name=Path(pdf_name(ok[0]["file"])).name,
            mime="application/pdf",
            type="primary",
        )
    elif ok:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in ok:
                zf.writestr(pdf_name(r["file"]), r["pdf"])
        buf.seek(0)
        st.download_button(
            f"⬇ Descargar {len(ok)} PDFs (.zip)",
            data=buf,
            file_name=f"{folder_name} — PDF.zip",
            mime="application/zip",
            type="primary",
        )
