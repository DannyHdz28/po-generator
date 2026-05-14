import io
import re
from collections import defaultdict
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

# ─── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(page_title="Deck Builder — Pro Standard", page_icon="🏆", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.main-title  { font-family: 'Bebas Neue', sans-serif; font-size: 3rem;   letter-spacing: 4px; color: #e8c84a; margin-bottom: 0; }
.sub-title   { font-family: 'Bebas Neue', sans-serif; font-size: 1.1rem; letter-spacing: 3px; color: #888;     margin-top: 0; }
.section-h   { font-family: 'Bebas Neue', sans-serif; font-size: 1.4rem; letter-spacing: 2px; color: #f5f5f0;  margin-top: 1.5rem; }
.pill-ok     { background: rgba(74,222,128,0.12); border:1px solid rgba(74,222,128,0.4); color:#4ade80; padding:2px 10px; border-radius:3px; font-size:0.75rem; letter-spacing:1px; }
.pill-warn   { background: rgba(248,113,113,0.12); border:1px solid rgba(248,113,113,0.4); color:#f87171; padding:2px 10px; border-radius:3px; font-size:0.75rem; letter-spacing:1px; }
.capsule-card { background:#1a1a1a; border:1px solid #2a2a2a; border-radius:4px; padding:10px 14px; }
.capsule-card.match { border-color:#4ade80; }
.capsule-card h4 { font-family:'Bebas Neue',sans-serif; font-size:1rem; letter-spacing:1.5px; color:#f5f5f0; margin:0; }
.capsule-card p  { font-size:0.75rem; color:#888; margin:2px 0 0 0; }
.team-badge { display:inline-block; padding:6px 14px; background:rgba(232,200,74,0.1);
              border:1px solid rgba(232,200,74,0.3); border-radius:3px; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────
GENDER_CODES = {"M", "W", "K"}
IMAGE_EXTS   = {".jpg", ".jpeg", ".png"}

SERVER_BASE_DEFAULT = r"\\10.0.1.30\Sales Toolkits\PROSTANDARD\USA_CANADA\CATALOGS"

CAPSULES_BY_QUARTER = {
    "Q1": [
        "TEAM CITY", "FLAGSHIP", "SPRING BREAK", "ULTIMATE FAN",
        "PRO FILE", "PROPERTY OF", "HERITAGE HUSTLE",
        "HERITAGE & HUSTLE", "REFLECTION", "FLORAL SPORT",
    ],
    "Q2": [],
    "Q3": [],
    "Q4": [],
}

INSERT_LEFT   = Emu(795600)
INSERT_TOP    = Emu(0)
INSERT_WIDTH  = Emu(10598400)
INSERT_HEIGHT = Emu(6858000)


# ─── PARSERS ──────────────────────────────────────────────────
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[&\-]", " ", s.upper())).strip()


def parse_merch_name(filename: str):
    """Parsea LIGA_EQUIPO_CAPSULA[_GENERO][_NNN].ext
    Soporta tanto M_001 como M-00 como sufijo de género+número."""
    base = re.sub(r"\.[^.]+$", "", filename)
    parts = base.split("_")
    if len(parts) < 3:
        return None
    liga    = parts[0].upper().strip()
    end_idx = len(parts) - 1
    gender  = None

    # Detecta sufijo tipo "M-00", "W-01", "K-002"
    gender_num = re.match(r'^([MWK])-(\d+)$', parts[end_idx], re.IGNORECASE)
    if gender_num:
        gender  = gender_num.group(1).upper()
        end_idx -= 1
    elif parts[end_idx].isdigit():
        end_idx -= 1
        if end_idx >= 0 and parts[end_idx].upper() in GENDER_CODES:
            gender  = parts[end_idx].upper()
            end_idx -= 1
    elif parts[end_idx].upper() in GENDER_CODES:
        gender  = parts[end_idx].upper()
        end_idx -= 1

    if end_idx < 1:
        return None
    team    = parts[1].strip()
    capsule = " ".join(parts[2:end_idx + 1]).upper().strip()
    if not capsule:
        return None
    key = capsule + (f" {gender}" if gender else "")
    return {"liga": liga, "team": team, "capsule": capsule, "gender": gender, "key": key}


def parse_note_line(line: str):
    n = norm(line).strip()
    n = re.sub(r"\s+\d+$", "", n).strip()
    if len(n) < 2:
        return None
    parts = n.split(" ")
    last  = parts[-1]
    if last in GENDER_CODES:
        return {"key": n, "capsule": " ".join(parts[:-1]), "gender": last}
    return {"key": n, "capsule": n, "gender": None}


def find_note_match(notes_text: str, merch_map: dict):
    if not notes_text:
        return None
    for raw in notes_text.split("\n"):
        p = parse_note_line(raw)
        if not p:
            continue
        if p["key"] in merch_map:
            return p["key"]
        if p["capsule"] in merch_map:
            return p["capsule"]
        norm_p = norm(p["key"])
        for k in merch_map:
            if norm(k) == norm_p:
                return k
    return None


# ─── SERVER HELPERS ───────────────────────────────────────────
def list_dirs(path: Path) -> list[str]:
    try:
        return sorted([d.name for d in path.iterdir() if d.is_dir()])
    except PermissionError:
        st.error(f"Sin permisos para leer: {path}")
        return []
    except Exception as e:
        st.error(f"Error leyendo carpeta: {e}")
        return []


def list_images(path: Path) -> list[Path]:
    try:
        return sorted(
            [f for f in path.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTS]
        )
    except Exception:
        return []


def find_merchboards_dir(capsule_path: Path) -> Path | None:
    direct = capsule_path / "_MERCHBOARDS"
    if direct.is_dir():
        return direct
    try:
        for d in capsule_path.iterdir():
            if d.is_dir() and "MERCHBOARD" in d.name.upper():
                return d
    except Exception:
        pass
    return None


def build_merch_map_from_paths(image_paths: list[Path]) -> tuple[dict, str | None, list[str]]:
    merch_map     = {}
    detected_team = None
    unmatched     = []

    for img_path in image_paths:
        parsed = parse_merch_name(img_path.name)
        if parsed and parsed["capsule"]:
            merch_map.setdefault(parsed["key"], []).append({
                "name":  img_path.name,
                "bytes": img_path.read_bytes(),
            })
            if not detected_team:
                detected_team = parsed["team"]
        else:
            unmatched.append(img_path.name)

    for k in merch_map:
        merch_map[k].sort(key=lambda x: x["name"])

    return merch_map, detected_team, unmatched


# ─── PPT OPERATIONS ───────────────────────────────────────────
def scan_pptx(file_bytes: bytes, known_capsules: list | None = None):
    prs        = Presentation(io.BytesIO(file_bytes))
    known_norm = {norm(c) for c in (known_capsules or [])}
    found, missing_gender, unknown_capsule = [], [], []

    for idx, slide in enumerate(prs.slides, start=1):
        if not slide.has_notes_slide:
            continue
        notes = slide.notes_slide.notes_text_frame.text or ""
        for raw in notes.split("\n"):
            p = parse_note_line(raw)
            if not (p and p["capsule"]):
                continue
            found.append({"slide": idx, "key": p["key"], "capsule": p["capsule"], "gender": p["gender"]})
            if not p["gender"]:
                missing_gender.append({"slide": idx, "capsule": p["capsule"]})
            if known_norm and norm(p["capsule"]) not in known_norm:
                unknown_capsule.append({"slide": idx, "capsule": p["capsule"]})
            break

    return found, missing_gender, unknown_capsule


def replace_picture_blob(shape, new_blob: bytes) -> bool:
    blip = shape._element.find(".//" + qn("a:blip"))
    if blip is None:
        return False
    r_id = blip.get(qn("r:embed"))
    if not r_id:
        return False
    img_part       = shape.part.related_parts[r_id]
    img_part._blob = new_blob
    return True


def generate_deck(ppt_bytes: bytes, merch_map: dict):
    prs      = Presentation(io.BytesIO(ppt_bytes))
    used     = defaultdict(int)
    replaced = 0
    log      = []

    for idx, slide in enumerate(prs.slides, start=1):
        if not slide.has_notes_slide:
            continue
        notes       = slide.notes_slide.notes_text_frame.text or ""
        matched_key = find_note_match(notes, merch_map)
        if not matched_key:
            continue

        imgs = merch_map[matched_key]
        if used[matched_key] >= len(imgs):
            log.append(("warn", f"Slide {idx} ({matched_key}): no more images available"))
            continue

        img = imgs[used[matched_key]]
        used[matched_key] += 1

        pics = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]

        if pics:
            if replace_picture_blob(pics[0], img["bytes"]):
                replaced += 1
                log.append(("ok", f"Slide {idx} → {matched_key} → {img['name']}"))
            else:
                log.append(("err", f"Slide {idx} ({matched_key}): could not replace image"))
        else:
            try:
                slide.shapes.add_picture(
                    io.BytesIO(img["bytes"]),
                    left=INSERT_LEFT, top=INSERT_TOP,
                    width=INSERT_WIDTH, height=INSERT_HEIGHT,
                )
                replaced += 1
                log.append(("ok", f"Slide {idx} → {matched_key} → {img['name']} (inserted)"))
            except Exception as e:
                log.append(("err", f"Slide {idx} ({matched_key}): error inserting — {e}"))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out, replaced, log


# ═══════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="main-title">DECK BUILDER</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">PRO STANDARD · MERCHBOARD AUTO-REPLACE</div>', unsafe_allow_html=True)
st.write("")

with st.expander("Cómo preparar tu PPT", expanded=False):
    st.markdown("""
1. En cada slide de **merchboard** agrega en las **Notas** la cápsula + género:
   `FLAGSHIP M`, `SPRING BREAK W`, `TEAM CITY K`  (M = Mens, W = Womens, K = Kids).
2. Las imágenes en el servidor deben seguir el naming:
   `LIGA_EQUIPO_CAPSULA_GENERO-NNN.png`
   Ej: `NBA_CLEVELAND CAVALIERS_TEAM CITY_M-00.png`
    """)

# ─── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### CONFIG")

    quarter = st.selectbox("Cuarto", ["Q1", "Q2", "Q3", "Q4"], index=0)
    known_capsules = CAPSULES_BY_QUARTER.get(quarter, [])
    if known_capsules:
        st.caption(f"Cápsulas {quarter}:")
        for c in known_capsules:
            st.caption(f"· {c}")
    else:
        st.caption(f"Sin cápsulas definidas para {quarter}.")

    st.markdown("---")
    st.markdown("### SERVIDOR")
    server_base = st.text_input(
        "Ruta base",
        value=SERVER_BASE_DEFAULT,
        help="Ruta UNC al servidor, ej: \\\\10.0.1.30\\Sales Toolkits\\...",
    )

# ─── STEP 1: PPT BASE ─────────────────────────────────────────
st.markdown('<div class="section-h">1 · PPT BASE</div>', unsafe_allow_html=True)
ppt_file = st.file_uploader("Sube tu PPT base (.pptx)", type=["pptx"], key="ppt")

ppt_bytes    = None
slides_found = []

if ppt_file:
    ppt_bytes = ppt_file.getvalue()
    try:
        slides_found, missing_gender, unknown_capsule = scan_pptx(ppt_bytes, known_capsules)
        if slides_found:
            st.markdown(f'<span class="pill-ok">✓ {len(slides_found)} SLIDE(S) MARCADAS</span>', unsafe_allow_html=True)
            for s in slides_found:
                st.caption(f"Slide {s['slide']} → {s['capsule']} · género: {s['gender'] or '—'}")
        else:
            st.markdown('<span class="pill-warn">⚠ SIN NOTAS DE CÁPSULA</span>', unsafe_allow_html=True)
            st.caption("Verifica que las slides de merchboard tengan notas tipo `FLAGSHIP M`.")

        if missing_gender:
            st.warning(
                f"⚠ {len(missing_gender)} slide(s) sin género (M/W/K) en la nota:\n\n"
                + "\n".join(f"- Slide {m['slide']} → `{m['capsule']}`" for m in missing_gender)
            )
        if unknown_capsule:
            st.warning(
                f"⚠ {len(unknown_capsule)} slide(s) con cápsulas no listadas en {quarter}:\n\n"
                + "\n".join(f"- Slide {u['slide']} → `{u['capsule']}`" for u in unknown_capsule)
            )
    except Exception as e:
        st.error(f"Error leyendo el PPT: {e}")

# ─── STEP 2: MERCHBOARDS ──────────────────────────────────────
st.markdown('<div class="section-h">2 · MERCHBOARDS</div>', unsafe_allow_html=True)

modo = st.radio(
    "Origen de las imágenes",
    ["🖥️  Desde el servidor", "📁  Subir manualmente"],
    horizontal=True,
)

merch_map       = {}
detected_team   = None
unmatched_names = []

# ── MODO SERVIDOR ─────────────────────────────────────────────
if modo == "🖥️  Desde el servidor":
    base_path = Path(server_base)

    if not base_path.exists():
        st.error(f"No se puede acceder a la ruta: `{server_base}`\n\nVerifica que estés conectado a la red de la empresa.")
    else:
        # Nivel 1: Categoría (MENS, WOMENS, KIDS, etc.) — selección múltiple
        categories = list_dirs(base_path)
        sel_categories = []
        if categories:
            sel_categories = st.multiselect(
                "Categoría (puedes seleccionar varias)",
                categories,
                key="srv_gender",
                help="Ej: selecciona MENS + WOMENS + KIDS para cargar las 3 a la vez",
            )
        else:
            st.warning("No se encontraron carpetas de categoría.")

        # Nivel 2: Año (basado en la primera categoría seleccionada)
        sel_year = None
        if sel_categories:
            year_path = base_path / sel_categories[0]
            years     = list_dirs(year_path)
            if years:
                sel_year = st.selectbox("Año", years, key="srv_year")
            else:
                st.warning(f"No hay subcarpetas en {year_path}")

        # Nivel 3: Cápsulas — selección múltiple (agrega carpetas de todas las categorías)
        sel_capsule_folders = []
        if sel_year and sel_categories:
            all_capsule_folders = set()
            for cat in sel_categories:
                q_path = base_path / cat / sel_year / quarter
                if q_path.is_dir():
                    for folder in list_dirs(q_path):
                        all_capsule_folders.add(folder)

            if all_capsule_folders:
                sel_capsule_folders = st.multiselect(
                    "Cápsula(s) (puedes seleccionar varias)",
                    sorted(all_capsule_folders),
                    key="srv_capsule",
                    help="Selecciona una o más cápsulas",
                )
            else:
                st.warning(f"No hay carpetas de cápsula para {quarter} en las categorías seleccionadas.")

        # Nivel 4: Liga y Equipo (basados en la primera combinación válida)
        sel_league = None
        sel_team   = None

        if sel_capsule_folders and sel_categories and sel_year:
            # Encontrar el primer _MERCHBOARDS válido para listar ligas
            ref_mb_path = None
            for cat in sel_categories:
                for cap in sel_capsule_folders:
                    cap_path = base_path / cat / sel_year / quarter / cap
                    mb = find_merchboards_dir(cap_path)
                    if mb:
                        ref_mb_path = mb
                        break
                if ref_mb_path:
                    break

            if not ref_mb_path:
                st.warning("No se encontró carpeta _MERCHBOARDS en ninguna de las cápsulas seleccionadas.")
            else:
                leagues = list_dirs(ref_mb_path)
                if leagues:
                    sel_league = st.selectbox("Liga", leagues, key="srv_league")
                else:
                    st.warning(f"No hay ligas en {ref_mb_path}")

                if sel_league:
                    # Encontrar equipos en la primera liga válida
                    ref_league_path = None
                    for cat in sel_categories:
                        for cap in sel_capsule_folders:
                            cap_path = base_path / cat / sel_year / quarter / cap
                            mb = find_merchboards_dir(cap_path)
                            if mb:
                                lp = mb / sel_league
                                if lp.is_dir():
                                    ref_league_path = lp
                                    break
                        if ref_league_path:
                            break

                    if ref_league_path:
                        teams = list_dirs(ref_league_path)
                        if teams:
                            sel_team = st.selectbox("Equipo", teams, key="srv_team")
                        else:
                            st.warning(f"No hay equipos en {ref_league_path}")

                    # Cargar imágenes de TODAS las combinaciones categoría × cápsula
                    if sel_team:
                        all_image_paths = []
                        loaded_paths    = []
                        for cat in sel_categories:
                            for cap in sel_capsule_folders:
                                team_dir = base_path / cat / sel_year / quarter / cap
                                mb = find_merchboards_dir(team_dir)
                                if mb:
                                    team_path = mb / sel_league / sel_team
                                    if team_path.is_dir():
                                        imgs = list_images(team_path)
                                        if imgs:
                                            all_image_paths.extend(imgs)
                                            loaded_paths.append(f"{cat} / {cap}")

                        if all_image_paths:
                            merch_map, detected_team, unmatched_names = build_merch_map_from_paths(all_image_paths)

                            st.markdown(
                                f'<div class="team-badge">'
                                f'<span style="color:#888;font-size:0.7rem;letter-spacing:1px;">EQUIPO · </span>'
                                f'<span style="color:#e8c84a;font-family:\'Bebas Neue\',sans-serif;letter-spacing:1px;">'
                                f'{sel_team}</span></div>',
                                unsafe_allow_html=True,
                            )
                            st.caption(f"Rutas cargadas: {', '.join(loaded_paths)}")
                            st.write("")

                            if merch_map:
                                cols = st.columns(min(4, len(merch_map)))
                                for i, (k, v) in enumerate(merch_map.items()):
                                    with cols[i % len(cols)]:
                                        st.markdown(
                                            f'<div class="capsule-card match"><h4>{k}</h4>'
                                            f'<p>{len(v)} image(s)</p></div>',
                                            unsafe_allow_html=True,
                                        )
                            if unmatched_names:
                                with st.expander(f"⚠ {len(unmatched_names)} without correct naming"):
                                    for n in unmatched_names:
                                        st.caption(n)
                        else:
                            st.warning(f"No hay imágenes JPG/PNG para {sel_team} en las rutas seleccionadas.")

# ── MODO MANUAL ───────────────────────────────────────────────
else:
    merch_files = st.file_uploader(
        "Sube las imágenes (JPG/PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="merch_manual",
    )

    if merch_files:
        for f in merch_files:
            parsed = parse_merch_name(f.name)
            if parsed and parsed["capsule"]:
                merch_map.setdefault(parsed["key"], []).append({"name": f.name, "bytes": f.getvalue()})
                if not detected_team:
                    detected_team = parsed["team"]
            else:
                unmatched_names.append(f.name)

        for k in merch_map:
            merch_map[k].sort(key=lambda x: x["name"])

        if detected_team:
            st.markdown(
                f'<div class="team-badge">'
                f'<span style="color:#888;font-size:0.7rem;letter-spacing:1px;">EQUIPO · </span>'
                f'<span style="color:#e8c84a;font-family:\'Bebas Neue\',sans-serif;letter-spacing:1px;">'
                f'{detected_team}</span></div>',
                unsafe_allow_html=True,
            )
            st.write("")

        if merch_map:
            cols = st.columns(min(4, len(merch_map)))
            for i, (k, v) in enumerate(merch_map.items()):
                with cols[i % len(cols)]:
                    st.markdown(
                        f'<div class="capsule-card match"><h4>{k}</h4><p>{len(v)} image(s)</p></div>',
                        unsafe_allow_html=True,
                    )

        if unmatched_names:
            with st.expander(f"⚠ {len(unmatched_names)} without correct naming"):
                for n in unmatched_names:
                    st.caption(n)

# ─── STEP 3: GENERAR DECK ─────────────────────────────────────
st.markdown('<div class="section-h">3 · GENERAR DECK</div>', unsafe_allow_html=True)

if not ppt_bytes:
    st.caption("Sube el PPT base en el paso 1.")
elif not merch_map:
    st.caption("Selecciona o sube los merchboards en el paso 2.")
else:
    mapping_preview = []
    for s in slides_found:
        match = find_note_match(s["key"], merch_map)
        mapping_preview.append({
            "Slide":               s["slide"],
            "Nota":                s["key"],
            "Género":              s["gender"] or "—",
            "Cápsula matcheada":   match or "⚠ sin match",
            "Imágenes disponibles": len(merch_map.get(match, [])) if match else 0,
        })

    if mapping_preview:
        st.caption("Mapping propuesto:")
        st.dataframe(mapping_preview, hide_index=True, use_container_width=True)

    no_match = sum(1 for r in mapping_preview if r["Cápsula matcheada"] == "⚠ sin match")
    if no_match:
        st.warning(f"⚠ {no_match} slide(s) no tienen imágenes que coincidan. Se generará el deck pero esas slides quedarán sin cambios.")

    if st.button("Generar Deck →", type="primary"):
        with st.spinner("Procesando slides..."):
            out, replaced, log = generate_deck(ppt_bytes, merch_map)

        st.markdown(
            f'<span class="pill-ok">✓ {replaced} SLIDE(S) ACTUALIZADAS</span>',
            unsafe_allow_html=True,
        )
        st.write("")

        with st.expander("Detalle", expanded=True):
            for lvl, msg in log:
                icon  = {"ok": "✓", "warn": "⚠", "err": "✗"}.get(lvl, "·")
                color = {"ok": "#4ade80", "warn": "#e8c84a", "err": "#f87171"}.get(lvl, "#888")
                st.markdown(
                    f'<div style="font-size:0.8rem;color:{color};">{icon} {msg}</div>',
                    unsafe_allow_html=True,
                )

        base_name = re.sub(r"\.pptx$", "", ppt_file.name, flags=re.IGNORECASE)
        out_name  = f"{base_name} — {detected_team or 'TEAM'}.pptx"
        st.download_button(
            "⬇ Descargar PPT",
            data=out,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
