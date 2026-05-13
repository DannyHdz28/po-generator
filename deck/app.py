import io
import re
from collections import defaultdict

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
.capsule-card{ background:#1a1a1a; border:1px solid #2a2a2a; border-radius:4px; padding:10px 14px; }
.capsule-card.match { border-color:#4ade80; }
.capsule-card h4 { font-family:'Bebas Neue',sans-serif; font-size:1rem; letter-spacing:1.5px; color:#f5f5f0; margin:0; }
.capsule-card p  { font-size:0.75rem; color:#888; margin:2px 0 0 0; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────
GENDER_CODES = {"M", "W", "K"}

# Cápsulas por cuarto del año. Q1 ya está, Q2-Q4 se llenan cuando salgan.
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

# Tamaño y posición para imágenes insertadas en slides vacías (29.44cm × 19.05cm en (2.21cm, 0)).
INSERT_LEFT   = Emu(795600)
INSERT_TOP    = Emu(0)
INSERT_WIDTH  = Emu(10598400)
INSERT_HEIGHT = Emu(6858000)

# ─── PARSERS ──────────────────────────────────────────────────
def norm(s: str) -> str:
    """Normaliza para comparación: upper, sin &/-, espacios colapsados."""
    return re.sub(r"\s+", " ", re.sub(r"[&\-]", " ", s.upper())).strip()


def parse_merch_name(filename: str):
    """
    Parsea LIGA_EQUIPO_CAPSULA[_GENERO][_NNN].ext
    Ej: NFL_DALLAS COWBOYS_SPRING BREAK_M_001.jpg
    """
    base = re.sub(r"\.[^.]+$", "", filename)
    parts = base.split("_")
    if len(parts) < 3:
        return None

    liga = parts[0].upper().strip()
    end_idx = len(parts) - 1

    # quita número final si existe (001, 002...)
    if parts[end_idx].isdigit():
        end_idx -= 1

    # quita género si existe
    gender = None
    if end_idx >= 0 and parts[end_idx].upper() in GENDER_CODES:
        gender = parts[end_idx].upper()
        end_idx -= 1

    if end_idx < 1:
        return None

    team = parts[1].strip()
    capsule = " ".join(parts[2:end_idx + 1]).upper().strip()
    if not capsule:
        return None

    key = capsule + (f" {gender}" if gender else "")
    return {"liga": liga, "team": team, "capsule": capsule, "gender": gender, "key": key}


def parse_note_line(line: str):
    """Parsea una línea de notas → {key, capsule, gender}"""
    n = norm(line).strip()
    n = re.sub(r"\s+\d+$", "", n).strip()   # quita número de slide que PPT a veces agrega
    if len(n) < 2:
        return None
    parts = n.split(" ")
    last = parts[-1]
    if last in GENDER_CODES:
        return {"key": n, "capsule": " ".join(parts[:-1]), "gender": last}
    return {"key": n, "capsule": n, "gender": None}


def find_note_match(notes_text: str, merch_map: dict):
    """Busca en las notas la primera línea que matchee una key de merch_map."""
    if not notes_text:
        return None
    for raw in notes_text.split("\n"):
        p = parse_note_line(raw)
        if not p:
            continue
        # match exacto con género
        if p["key"] in merch_map:
            return p["key"]
        # match solo cápsula (notas sin género)
        if p["capsule"] in merch_map:
            return p["capsule"]
        # match normalizado (& vs AND, espacios)
        norm_p = norm(p["key"])
        for k in merch_map:
            if norm(k) == norm_p:
                return k
    return None


# ─── PPT OPERATIONS ───────────────────────────────────────────
def scan_pptx(file_bytes: bytes, known_capsules: list | None = None):
    """
    Escanea el PPT. Devuelve:
      - found:           [{slide, key, capsule, gender}]
      - missing_gender:  [{slide, capsule}]  (slides con cápsula pero sin M/W/K)
      - unknown_capsule: [{slide, capsule}]  (cápsula no está en known_capsules; solo si la lista no está vacía)
    """
    prs = Presentation(io.BytesIO(file_bytes))
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
    """Reemplaza los bytes de la imagen embebida de un shape PICTURE."""
    blip = shape._element.find(".//" + qn("a:blip"))
    if blip is None:
        return False
    r_id = blip.get(qn("r:embed"))
    if not r_id:
        return False
    img_part = shape.part.related_parts[r_id]
    img_part._blob = new_blob
    return True


def generate_deck(ppt_bytes: bytes, merch_map: dict):
    """
    Reemplaza imágenes en las slides marcadas. Para slides con varias imágenes
    en una misma cápsula (001, 002...), va consumiendo en orden.
    """
    prs = Presentation(io.BytesIO(ppt_bytes))
    used = defaultdict(int)
    replaced = 0
    log = []

    for idx, slide in enumerate(prs.slides, start=1):
        if not slide.has_notes_slide:
            continue
        notes = slide.notes_slide.notes_text_frame.text or ""
        matched_key = find_note_match(notes, merch_map)
        if not matched_key:
            continue

        imgs = merch_map[matched_key]
        if used[matched_key] >= len(imgs):
            log.append(("warn", f"Slide {idx} ({matched_key}): no hay más imágenes disponibles"))
            continue

        img = imgs[used[matched_key]]
        used[matched_key] += 1

        pics = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]

        if pics:
            # CASO 1: ya hay imagen → reemplaza bytes (conserva tamaño/posición original)
            if replace_picture_blob(pics[0], img["bytes"]):
                replaced += 1
                log.append(("ok", f"Slide {idx} → {matched_key} → {img['name']}"))
            else:
                log.append(("err", f"Slide {idx} ({matched_key}): no se pudo reemplazar"))
        else:
            # CASO 2: slide vacía → inserta imagen nueva con tamaño/posición estándar
            try:
                slide.shapes.add_picture(
                    io.BytesIO(img["bytes"]),
                    left=INSERT_LEFT, top=INSERT_TOP,
                    width=INSERT_WIDTH, height=INSERT_HEIGHT,
                )
                replaced += 1
                log.append(("ok", f"Slide {idx} → {matched_key} → {img['name']} (insertada)"))
            except Exception as e:
                log.append(("err", f"Slide {idx} ({matched_key}): error insertando — {e}"))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out, replaced, log


# ─── UI ───────────────────────────────────────────────────────
st.markdown('<div class="main-title">DECK BUILDER</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">PRO STANDARD · MERCHBOARD AUTO-REPLACE</div>', unsafe_allow_html=True)
st.write("")

with st.expander("Cómo preparar tu PPT", expanded=False):
    st.markdown("""
1. Abre tu PPT base en PowerPoint y ve a cada slide de **merchboard**.
2. En las **Notas** del slide escribe la cápsula + género:
   `FLAGSHIP M`, `SPRING BREAK W`, `TEAM CITY K`  (M = Mens, W = Womens, K = Kids).
3. Guarda el PPT y súbelo abajo.
4. **Nombres de imagen:** `LIGA_EQUIPO_CAPSULA_GENERO[_NNN].ext`
   - 1 imagen: `NFL_LAS VEGAS RAIDERS_FLAGSHIP_M.jpg`
   - Varias:  `NFL_LAS VEGAS RAIDERS_FLAGSHIP_M_001.jpg`, `..._002.jpg`, ...
    """)

# ─── SIDEBAR: cuarto y cápsulas esperadas ─────────────────────
with st.sidebar:
    st.markdown('<div class="section-h" style="font-size:1.1rem; margin-top:0;">CONFIG</div>', unsafe_allow_html=True)
    quarter = st.selectbox("Cuarto del año", ["Q1", "Q2", "Q3", "Q4"], index=0)
    known_capsules = CAPSULES_BY_QUARTER.get(quarter, [])
    if known_capsules:
        st.caption(f"Cápsulas esperadas en {quarter}:")
        st.markdown("<br>".join(f"· {c}" for c in known_capsules), unsafe_allow_html=True)
    else:
        st.caption(f"⚠ Sin cápsulas definidas para {quarter}. La app igual funciona, pero no validará nombres.")

# ─── STEP 1 ───────────────────────────────────────────────────
st.markdown('<div class="section-h">1 · PPT BASE</div>', unsafe_allow_html=True)
ppt_file = st.file_uploader("Sube tu PPT base (.pptx)", type=["pptx"], key="ppt")

ppt_bytes = None
slides_found = []
if ppt_file:
    ppt_bytes = ppt_file.getvalue()
    try:
        slides_found, missing_gender, unknown_capsule = scan_pptx(ppt_bytes, known_capsules)
        if slides_found:
            st.markdown(f'<span class="pill-ok">✓ {len(slides_found)} SLIDE(S) MARCADAS</span>', unsafe_allow_html=True)
            with st.container():
                for s in slides_found:
                    gender_tag = s["gender"] or "—"
                    st.caption(f"Slide {s['slide']} → {s['capsule']} · género: {gender_tag}")
        else:
            st.markdown('<span class="pill-warn">⚠ SIN NOTAS DE CÁPSULA</span>', unsafe_allow_html=True)
            st.caption("Verifica que las slides de merchboard tengan en las Notas algo como `FLAGSHIP M`.")

        if missing_gender:
            st.warning(
                f"⚠ {len(missing_gender)} slide(s) tienen cápsula pero **sin género (M/W/K)**:\n\n"
                + "\n".join(f"- Slide {m['slide']} → `{m['capsule']}`" for m in missing_gender)
                + "\n\nAgrega el género en las notas para evitar matches ambiguos."
            )

        if unknown_capsule:
            st.warning(
                f"⚠ {len(unknown_capsule)} slide(s) con cápsulas no listadas en **{quarter}**:\n\n"
                + "\n".join(f"- Slide {u['slide']} → `{u['capsule']}`" for u in unknown_capsule)
                + f"\n\nSi son correctas, agrégalas a `CAPSULES_BY_QUARTER['{quarter}']` en el código."
            )
    except Exception as e:
        st.error(f"Error leyendo el PPT: {e}")

# ─── STEP 2 ───────────────────────────────────────────────────
st.markdown('<div class="section-h">2 · MERCHBOARDS</div>', unsafe_allow_html=True)
merch_files = st.file_uploader(
    "Sube las imágenes (JPG/PNG). Puedes seleccionar varias.",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="merch",
)

merch_map: dict = {}
detected_team = None
unmatched_names = []

if merch_files:
    for f in merch_files:
        parsed = parse_merch_name(f.name)
        if parsed and parsed["capsule"]:
            merch_map.setdefault(parsed["key"], []).append({"name": f.name, "bytes": f.getvalue()})
            if not detected_team:
                detected_team = parsed["team"]
        else:
            unmatched_names.append(f.name)

    # orden estable por nombre (respeta _001, _002...)
    for k in merch_map:
        merch_map[k].sort(key=lambda x: x["name"])

    if detected_team:
        st.markdown(
            f'<div style="display:inline-block; padding:6px 14px; background:rgba(232,200,74,0.1); '
            f'border:1px solid rgba(232,200,74,0.3); border-radius:3px; font-size:0.85rem;">'
            f'<span style="color:#888; font-size:0.7rem; letter-spacing:1px;">EQUIPO DETECTADO · </span>'
            f'<span style="color:#e8c84a; font-family: Bebas Neue, sans-serif; letter-spacing:1px;">{detected_team}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.write("")

    if merch_map:
        cols = st.columns(min(4, len(merch_map)))
        for i, (k, v) in enumerate(merch_map.items()):
            with cols[i % len(cols)]:
                st.markdown(
                    f'<div class="capsule-card match"><h4>{k}</h4><p>{len(v)} imagen(es)</p></div>',
                    unsafe_allow_html=True,
                )

    if unmatched_names:
        with st.expander(f"⚠ {len(unmatched_names)} archivo(s) sin coincidencia"):
            for n in unmatched_names:
                st.caption(n)

# ─── STEP 3 ───────────────────────────────────────────────────
st.markdown('<div class="section-h">3 · GENERAR DECK</div>', unsafe_allow_html=True)

if not ppt_bytes:
    st.caption("Sube el PPT base.")
elif not merch_map:
    st.caption("Sube los merchboards.")
else:
    # preview del mapping antes de generar
    mapping_preview = []
    for s in slides_found:
        match = find_note_match(s["key"], merch_map)
        mapping_preview.append({
            "Slide": s["slide"],
            "Nota": s["key"],
            "Género": s["gender"] or "—",
            "Cápsula matcheada": match or "—",
            "Imágenes disponibles": len(merch_map.get(match, [])) if match else 0,
        })

    if mapping_preview:
        st.caption("Mapping propuesto:")
        st.dataframe(mapping_preview, hide_index=True, use_container_width=True)

    if st.button("Generar Deck →", type="primary"):
        with st.spinner("Generando..."):
            out, replaced, log = generate_deck(ppt_bytes, merch_map)

        st.markdown(
            f'<span class="pill-ok">✓ {replaced} SLIDE(S) ACTUALIZADAS</span>',
            unsafe_allow_html=True,
        )
        st.write("")

        with st.expander("Detalle", expanded=True):
            for lvl, msg in log:
                icon = {"ok": "✓", "warn": "⚠", "err": "✗"}.get(lvl, "·")
                color = {"ok": "#4ade80", "warn": "#e8c84a", "err": "#f87171"}.get(lvl, "#888")
                st.markdown(
                    f'<div style="font-size:0.8rem; color:{color};">{icon} {msg}</div>',
                    unsafe_allow_html=True,
                )

        base_name = re.sub(r"\.pptx$", "", ppt_file.name, flags=re.IGNORECASE)
        out_name = f"{base_name} — {detected_team or 'EQUIPO'}.pptx"
        st.download_button(
            "⬇ Descargar PPT",
            data=out,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
