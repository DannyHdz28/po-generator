import io
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import pymupdf
import streamlit as st

import currency as cu

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
        with pymupdf.open(stream=data, filetype="pdf") as doc:
            n_pages = doc.page_count
            if n_pages == 0:
                return None, "El PDF interno no tiene páginas"
            first_text = doc[0].get_text().lower()
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
        with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
            text   = "\n".join(pg.get_text() for pg in doc)
            prices = [p["text"] for pg in doc for p in cu.find_prices(pg)]
    except Exception as e:
        return {"status": "error", "msg": f"No se pudo leer el texto: {e}", "prices": []}

    chars  = len(re.sub(r"\s+", "", text))
    prices = list(dict.fromkeys(p.replace(" ", "") for p in prices))

    if chars < MIN_TEXT_CHARS:
        return {"status": "outlines", "prices": [],
                "msg": "SIN TEXTO — el texto está en contornos (no editable)"}
    if prices:
        shown = ", ".join(prices[:8]) + (" …" if len(prices) > 8 else "")
        return {"status": "editable", "prices": prices,
                "msg": f"TEXTO EDITABLE · {len(prices)} precio(s): {shown}"}
    return {"status": "no_prices", "prices": [],
            "msg": "Texto editable, pero NO se encontraron precios como texto (¿precios en contornos?)"}


def out_base(file: dict) -> str:
    """Ruta relativa de salida sin extensión."""
    base = re.sub(r"\.ai$", "", file["name"], flags=re.IGNORECASE)
    return f"{file['subdir']}/{base}" if file["subdir"] else base


def add_outputs(outputs: list, folder: str, file: dict, pdf_bytes: bytes,
                formats: list[str], dpi: int):
    """Agrega PDF y/o imágenes de un archivo a la lista de salidas (ruta_zip, bytes)."""
    prefix = f"{folder}/" if folder else ""
    base   = out_base(file)
    if "PDF" in formats:
        outputs.append((f"{prefix}{base}.pdf", pdf_bytes))
    if "JPG" in formats:
        imgs = cu.render_images(pdf_bytes, dpi, "jpg")
        for i, img in enumerate(imgs, start=1):
            suffix = f"_p{i}" if len(imgs) > 1 else ""
            outputs.append((f"{prefix}{base}{suffix}.jpg", img))


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
3. Selecciona los que quieras, elige monedas y formato (PDF / JPG) y da clic en **Generar**.

**Monedas** (opcional):
- **Una moneda por archivo** → una carpeta por moneda (`CAD/`, `MXN/`…).
- **Todas en un archivo** → cada precio muestra todas las monedas elegidas.
- Si no eliges monedas, se descargan los originales en USD.

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


# ─── STEP 3: MONEDAS Y FORMATO ────────────────────────────────
st.markdown('<div class="section-h">3 · MONEDAS Y FORMATO</div>', unsafe_allow_html=True)

if not selected:
    st.caption("Selecciona al menos un archivo.")
    st.stop()

total_mb = sum(f["size_mb"] for f in selected)
st.caption(f"{len(selected)} archivo(s) seleccionados · {total_mb:.1f} MB")

with st.expander("💱 Tipos de cambio (provisional)", expanded=False):
    st.caption(
        "Valores de EJEMPLO — se reemplazan cuando se conecte la fuente oficial de precios. "
        "Puedes editarlos o agregar monedas."
    )
    cur_rows = st.data_editor(
        [dict(c) for c in cu.DEFAULT_CURRENCIES],
        num_rows="dynamic",
        hide_index=True,
        width="stretch",
        column_config={
            "code":     st.column_config.TextColumn("Moneda", required=True),
            "prefix":   st.column_config.TextColumn("Prefijo"),
            "symbol":   st.column_config.TextColumn("Símbolo"),
            "rate":     st.column_config.NumberColumn("1 USD =", min_value=0.0, format="%.4f"),
            "decimals": st.column_config.NumberColumn("Decimales", min_value=0, max_value=4, step=1),
        },
        key="currency_table",
    )

cur_map = {}
for r in cur_rows:
    code = (r.get("code") or "").strip().upper()
    if not code or not r.get("rate"):
        continue
    cur_map[code] = {
        "code":     code,
        "prefix":   (r.get("prefix") or "").strip(),
        "symbol":   (r.get("symbol") or "").strip(),
        "rate":     float(r["rate"]),
        "decimals": int(r.get("decimals") if r.get("decimals") is not None else 2),
    }
usd = cur_map.get("USD", {"code": "USD", "prefix": "USD", "symbol": "$", "rate": 1.0, "decimals": 2})

c1, c2 = st.columns(2)
with c1:
    sel_codes = st.multiselect(
        "Monedas",
        [k for k in cur_map if k != "USD"],
        placeholder="Sin monedas = originales en USD",
    )
with c2:
    formats = st.multiselect("Formato de salida", ["PDF", "JPG"], default=["PDF"])
    dpi = 200
    if "JPG" in formats:
        dpi = st.select_slider("Calidad JPG (DPI)", options=[100, 150, 200, 300], value=200)

mode_single = mode_multi = include_orig = False
layout_multi, include_usd, align = cu.LAYOUT_LINES, True, "left"
if sel_codes:
    m1, m2 = st.columns(2)
    with m1:
        mode_single = st.checkbox("Una moneda por archivo", value=True,
                                  help="Una carpeta por moneda: CAD/, MXN/, …")
        include_orig = st.checkbox("Incluir también los originales (USD)", value=False)
    with m2:
        mode_multi = st.checkbox("Todas las monedas en un archivo", value=len(sel_codes) > 1,
                                 help="Cada precio muestra todas las monedas elegidas")
        if mode_multi:
            layout_multi = cu.LAYOUT_LINES if st.radio(
                "Formato", ["En líneas (una debajo de otra)", "En una sola línea"], horizontal=True,
            ).startswith("En líneas") else cu.LAYOUT_INLINE
            include_usd = st.checkbox("Incluir USD", value=True)
    align = {"Izquierda": "left", "Centro": "center", "Derecha": "right"}[st.radio(
        "Alineación del precio", ["Izquierda", "Centro", "Derecha"], horizontal=True,
        help="Cómo se acomoda el precio nuevo respecto al original (cuando el precio está solo en su línea)",
    )]

if not formats:
    st.warning("Elige al menos un formato de salida (PDF o JPG).")
    st.stop()
if sel_codes and not (mode_single or mode_multi):
    st.warning("Elige al menos un modo: una moneda por archivo y/o todas en un archivo.")
    st.stop()

if st.button("Generar →", type="primary"):
    jobs = []   # (carpeta, monedas, layout)
    if mode_single:
        jobs += [(code, [cur_map[code]], cu.LAYOUT_SINGLE) for code in sel_codes]
    if mode_multi:
        curs = ([usd] if include_usd else []) + [cur_map[c] for c in sel_codes]
        jobs.append(("MULTI " + "-".join(c["code"] for c in curs), curs, layout_multi))

    outputs, report, changes, previews = [], [], [], []
    progress = st.progress(0.0, text="Procesando...")
    for i, f in enumerate(selected, start=1):
        progress.progress((i - 1) / len(selected), text=f"Procesando {i}/{len(selected)} · {f['name']}")
        pdf_bytes, msg = ai_to_pdf(f["path"])
        text_info      = analyze_text(pdf_bytes) if pdf_bytes else None
        report.append({"file": f, "ok": bool(pdf_bytes), "msg": msg, "text": text_info, "skipped": False})
        if not pdf_bytes:
            continue

        if not sel_codes or include_orig:
            add_outputs(outputs, "USD (original)" if sel_codes else "", f, pdf_bytes, formats, dpi)
        if not jobs:
            continue
        if text_info["status"] != "editable":
            report[-1]["skipped"] = True
            continue

        for folder, curs, lay in jobs:
            try:
                new_pdf, ch = cu.apply_currencies(pdf_bytes, curs, lay, align)
            except Exception as e:
                changes.append({"Archivo": out_base(f), "Versión": folder, "Pág": "", "Original": "",
                                "Nuevo": "", "Estado": "✗", "Nota": f"Error: {e}"})
                continue
            add_outputs(outputs, folder, f, new_pdf, formats, dpi)
            for c in ch:
                changes.append({
                    "Archivo": out_base(f), "Versión": folder, "Pág": c["page"],
                    "Original": c["original"], "Nuevo": c["new"],
                    "Estado": {"ok": "✓", "warn": "⚠", "err": "✗"}[c["status"]], "Nota": c["note"],
                })
            if len(previews) < 6:
                previews.append((f"{folder} · {f['name']}", cu.render_images(new_pdf, 72, "png")[0]))
    progress.empty()

    st.session_state["results"] = {
        "report": report, "outputs": outputs, "changes": changes,
        "previews": previews, "tag": "-".join(sel_codes) or "PDF",
    }
    st.session_state["results_dir"] = str(current)

# ─── RESULTADOS ───────────────────────────────────────────────
results = st.session_state.get("results")
if results:
    report  = results["report"]
    ok      = [r for r in report if r["ok"]]
    fail    = [r for r in report if not r["ok"]]
    skipped = [r for r in report if r["skipped"]]

    st.write("")
    if ok:
        st.markdown(f'<span class="pill-ok">✓ {len(ok)} ARCHIVO(S) LEÍDOS</span>', unsafe_allow_html=True)
    if fail:
        st.markdown(f'<span class="pill-warn">✗ {len(fail)} NO SE PUDIERON CONVERTIR</span>', unsafe_allow_html=True)
    if skipped:
        st.markdown(
            f'<span class="pill-warn">⚠ {len(skipped)} SIN PRECIOS EDITABLES — SIN VERSIÓN EN OTRAS MONEDAS</span>',
            unsafe_allow_html=True,
        )

    with st.expander("Detalle por archivo", expanded=bool(fail or skipped)):
        text_colors = {"editable": "#4ade80", "no_prices": "#e8c84a", "outlines": "#f87171", "error": "#f87171"}
        for r in report:
            color = "#4ade80" if r["ok"] else "#f87171"
            icon  = "✓" if r["ok"] else "✗"
            st.markdown(
                f'<div style="font-size:0.8rem;color:{color};">{icon} {r["file"]["name"]} — {r["msg"]}</div>',
                unsafe_allow_html=True,
            )
            if r["text"]:
                st.markdown(
                    f'<div style="font-size:0.75rem;color:{text_colors[r["text"]["status"]]};'
                    f'margin:0 0 6px 18px;">↳ {r["text"]["msg"]}</div>',
                    unsafe_allow_html=True,
                )

    if results["changes"]:
        n_warn = sum(1 for c in results["changes"] if c["Estado"] != "✓")
        st.markdown('<div class="section-h">REVISIÓN DE PRECIOS</div>', unsafe_allow_html=True)
        if n_warn:
            st.markdown(f'<span class="pill-warn">⚠ {n_warn} PRECIO(S) PARA REVISAR</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="pill-ok">✓ TODOS LOS PRECIOS CAMBIADOS SIN AVISOS</span>', unsafe_allow_html=True)
        st.dataframe(results["changes"], hide_index=True, width="stretch")

    if results["previews"]:
        st.markdown('<div class="section-h">PREVIEW</div>', unsafe_allow_html=True)
        pcols = st.columns(3)
        for i, (label, img) in enumerate(results["previews"]):
            with pcols[i % 3]:
                st.image(img, caption=label, width="stretch")

    outputs     = results["outputs"]
    folder_name = clean_name(current.name)
    if len(outputs) == 1:
        name, data = outputs[0]
        st.download_button(
            "⬇ Descargar",
            data=data,
            file_name=Path(name).name,
            mime="application/pdf" if name.endswith(".pdf") else "image/jpeg",
            type="primary",
        )
    elif outputs:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in outputs:
                zf.writestr(name, data)
        buf.seek(0)
        st.download_button(
            f"⬇ Descargar {len(outputs)} archivo(s) (.zip)",
            data=buf,
            file_name=f"{folder_name} — {results['tag']}.zip",
            mime="application/zip",
            type="primary",
        )
