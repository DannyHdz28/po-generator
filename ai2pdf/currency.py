"""
Cambio de moneda dentro de PDFs (merchboards exportados de Illustrator).

Los precios que son TEXTO editable se reemplazan directo: se quita solo el
texto del precio original (sin tapar nada) y se escribe el nuevo en la misma
posición, con el mismo tamaño y color.
"""
import re

import pymupdf

# Precios tipo $45 · $45.00 · $ 1,299.99 · USD 45 · USD $45.00
PRICE_RE = re.compile(r"(?:USD\s?)?\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?|USD\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?")

# Tabla PROVISIONAL de monedas. Los tipos de cambio son de ejemplo:
# se reemplazan cuando se conecte la fuente oficial de precios.
DEFAULT_CURRENCIES = [
    {"code": "USD", "prefix": "USD", "symbol": "$",  "rate": 1.0,   "decimals": 2},
    {"code": "CAD", "prefix": "CAD", "symbol": "$",  "rate": 1.37,  "decimals": 2},
    {"code": "MXN", "prefix": "MXN", "symbol": "$",  "rate": 18.50, "decimals": 2},
    {"code": "EUR", "prefix": "EUR", "symbol": "€",  "rate": 0.92,  "decimals": 2},
    {"code": "GBP", "prefix": "GBP", "symbol": "£",  "rate": 0.79,  "decimals": 2},
    {"code": "AED", "prefix": "AED", "symbol": "",   "rate": 3.67,  "decimals": 2},
]

LAYOUT_SINGLE = "single"   # una moneda:            CAD $61.65
LAYOUT_LINES  = "lines"    # varias, una por línea: USD $45.00 / MXN $832.50
LAYOUT_INLINE = "inline"   # varias, misma línea:   USD $45.00 | MXN $832.50

INLINE_SEP      = "  |  "
MAX_WIDTH_RATIO = {LAYOUT_SINGLE: 1.8, LAYOUT_LINES: 1.8, LAYOUT_INLINE: 4.5}  # vs. ancho del precio original
MIN_SIZE_RATIO  = 0.6   # nunca más chico que el 60% del tamaño original
LINE_SPACING    = 1.15


# ─── FORMATO ──────────────────────────────────────────────────
def parse_usd(text: str) -> float | None:
    num = re.sub(r"[^\d.]", "", text.replace("USD", ""))
    try:
        return float(num)
    except ValueError:
        return None


def format_price(usd_value: float, cur: dict) -> str:
    value = usd_value * float(cur["rate"])
    dec   = int(cur.get("decimals", 2))
    body  = f"{cur.get('symbol', '')}{value:,.{dec}f}"
    return f"{cur['prefix']} {body}".strip() if cur.get("prefix") else body


# ─── DETECCIÓN ────────────────────────────────────────────────
def _color_rgb(c: int) -> tuple[float, float, float]:
    return ((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255


def _style(span: dict) -> dict:
    return {
        "size":  span["size"],
        "font":  span["font"],
        "bold":  bool(span["flags"] & 16) or "bold" in span["font"].lower(),
        "color": _color_rgb(span["color"]),
    }


def _bbox(chars) -> pymupdf.Rect | None:
    rect = None
    for ch, _ in chars:
        if ch["c"].strip():
            rect = pymupdf.Rect(ch["bbox"]) if rect is None else rect | pymupdf.Rect(ch["bbox"])
    return rect


def find_price_lines(page: pymupdf.Page) -> list[dict]:
    """
    Líneas de texto que contienen precios. Por cada línea se guarda todo lo que
    va desde el primer precio hasta el final, en segmentos:
      ("plain", texto, estilo)  → texto normal que hay que recorrer
      ("price", texto, estilo, usd, rect)
    """
    lines = []
    raw   = page.get_text("rawdict", flags=pymupdf.TEXTFLAGS_RAWDICT & ~pymupdf.TEXT_PRESERVE_IMAGES)
    for block in raw["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            chars   = [(ch, span) for span in line["spans"] for ch in span["chars"]]
            text    = "".join(ch["c"] for ch, _ in chars)
            matches = list(PRICE_RE.finditer(text))
            if not matches:
                continue
            first    = matches[0].start()
            segments = []
            pos      = first
            for m in matches:
                if m.start() > pos:
                    seg = chars[pos:m.start()]
                    segments.append({"kind": "plain", "text": text[pos:m.start()],
                                     "style": _style(seg[0][1])})
                seg   = chars[m.start():m.end()]
                digit = next((s for ch, s in seg if ch["c"].isdigit()), seg[0][1])
                segments.append({"kind": "price", "text": m.group(0), "usd": parse_usd(m.group(0)),
                                 "style": _style(digit), "rect": _bbox(seg)})
                pos = m.end()
            if pos < len(chars):
                seg = chars[pos:]
                segments.append({"kind": "plain", "text": text[pos:], "style": _style(seg[0][1])})
            lines.append({
                "dir":      tuple(round(v, 3) for v in line["dir"]),
                "origin":   pymupdf.Point(chars[first][0]["origin"]),
                "rect":     _bbox(chars[first:]),
                "segments": segments,
            })
    return lines


def find_prices(page: pymupdf.Page) -> list[dict]:
    """Lista plana de precios de la página (para reportes)."""
    return [s for ln in find_price_lines(page) for s in ln["segments"] if s["kind"] == "price"]


# ─── FUENTES ──────────────────────────────────────────────────
def _strip_subset(name: str) -> str:
    return name.split("+", 1)[1] if "+" in name else name


def _pick_font(doc, page, style: dict, needed: str, cache: dict):
    """
    Usa la fuente original del PDF si tiene todas las letras necesarias
    (Illustrator a veces solo guarda las letras usadas). Si no, una parecida.
    Regresa (Font, es_original: bool)
    """
    key = (page.number, style["font"], style["bold"], needed)
    if key in cache:
        return cache[key]
    wanted = _strip_subset(style["font"])
    for f in page.get_fonts(full=True):
        if _strip_subset(f[3]) != wanted:
            continue
        try:
            _, ext, _, buf = doc.extract_font(f[0])
            if buf and ext not in ("n/a", ""):
                font = pymupdf.Font(fontbuffer=buf)
                if all(font.has_glyph(ord(c)) for c in set(needed) if c.strip()):
                    cache[key] = (font, True)
                    return cache[key]
        except Exception:
            pass
        break
    # Helvetica/Arial sin incrustar → la fuente de respaldo ES la misma
    same = wanted.split("-")[0].lower() in ("helvetica", "arial", "arialmt")
    cache[key] = (pymupdf.Font("hebo" if style["bold"] else "helv"), same)
    return cache[key]


# ─── REEMPLAZO ────────────────────────────────────────────────
def _fit_size(font, lines: list[str], size: float, max_w: float) -> float:
    widest = max(font.text_length(t, fontsize=size) for t in lines)
    if widest <= max_w:
        return size
    return max(size * MIN_SIZE_RATIO, size * max_w / widest)


def _shrink(r: pymupdf.Rect) -> pymupdf.Rect:
    r = pymupdf.Rect(r)
    r.y0 += r.height * 0.2
    r.y1 -= r.height * 0.2
    return r


def apply_currencies(pdf_bytes: bytes, currencies: list[dict], layout: str = LAYOUT_SINGLE,
                     align: str = "left") -> tuple[bytes, list[dict]]:
    """
    Reemplaza cada precio USD por su(s) precio(s) en las monedas dadas.
    Si después del precio hay más texto en la misma línea, se recorre para
    que no se encime. Regresa (pdf_bytes_nuevo, cambios) — cada cambio:
      {page, original, new, status: ok|warn|err, note}
    """
    doc     = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    changes = []
    cache   = {}

    for pno, page in enumerate(doc, start=1):
        lines = find_price_lines(page)
        if not lines:
            continue

        todo = []
        for ln in lines:
            prices = [s for s in ln["segments"] if s["kind"] == "price"]
            bad    = None
            if ln["dir"] != (1.0, 0.0):
                bad = "Texto rotado — revisar a mano"
            elif any(s["usd"] is None for s in prices):
                bad = "No se pudo leer el valor"
            if bad:
                changes += [{"page": pno, "original": s["text"], "new": "", "status": "err", "note": bad}
                            for s in prices]
                continue
            todo.append(ln)
            # quita SOLO el texto (sin relleno, sin tocar gráficos ni imágenes)
            page.add_redact_annot(_shrink(ln["rect"]), fill=False)

        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )
        # zonas que se van a reescribir (para no encimar líneas de precios entre sí)
        pending = [ln["rect"] for ln in todo]

        for ln in todo:
            segs      = ln["segments"]
            n_prices  = sum(1 for s in segs if s["kind"] == "price")
            has_tail  = any(s["kind"] == "plain" and s["text"].strip() for s in segs)
            x, y      = ln["origin"].x, ln["origin"].y
            tw_list   = []   # (texto, punto, font, size, color)

            for seg in segs:
                st_ = seg["style"]
                if seg["kind"] == "plain":
                    font, _ = _pick_font(doc, page, st_, seg["text"], cache)
                    tw_list.append((seg["text"], x, y, font, st_["size"], st_["color"]))
                    x += font.text_length(seg["text"], fontsize=st_["size"])
                    continue

                texts      = [format_price(seg["usd"], c) for c in currencies]
                font, f_ok = _pick_font(doc, page, st_, "".join(texts) + INLINE_SEP, cache)
                max_w      = seg["rect"].width * MAX_WIDTH_RATIO[layout]
                notes      = []

                use_lines = layout == LAYOUT_LINES and len(texts) > 1
                if use_lines and (has_tail or n_prices > 1):
                    use_lines = False
                    notes.append("en una línea: hay más texto junto al precio")
                if use_lines:
                    size = _fit_size(font, texts, st_["size"], max_w)
                    size = min(size, st_["size"] * max(0.6, 1 - 0.1 * (len(texts) - 1)))
                    width = max(font.text_length(t, fontsize=size) for t in texts)
                    below = pymupdf.Rect(x, y + size * 0.3, x + width, y + size * LINE_SPACING * (len(texts) - 1) + size * 0.3)
                    busy  = any(w for w in page.get_text("words", clip=below) if w[4].strip()) or \
                            any(below.intersects(r) for r in pending if r != ln["rect"])
                    if busy:
                        use_lines = False
                        notes.append("en una línea: no hay espacio libre abajo")

                out_lines = texts if use_lines else [INLINE_SEP.join(texts)] if len(texts) > 1 else texts
                ratio     = MAX_WIDTH_RATIO[LAYOUT_INLINE if len(out_lines) == 1 and len(texts) > 1 else layout]
                size      = _fit_size(font, out_lines, st_["size"], seg["rect"].width * ratio)
                if use_lines:
                    size = min(size, st_["size"] * max(0.6, 1 - 0.1 * (len(texts) - 1)))
                width = max(font.text_length(t, fontsize=size) for t in out_lines)

                # alineación solo aplica si el precio está solo en su línea
                bx = x
                if n_prices == 1 and not has_tail:
                    if align == "right":
                        bx = seg["rect"].x1 - width
                    elif align == "center":
                        bx = (seg["rect"].x0 + seg["rect"].x1) / 2 - width / 2

                for k, t in enumerate(out_lines):
                    w  = font.text_length(t, fontsize=size)
                    lx = bx if align == "left" or has_tail or n_prices > 1 else \
                         (bx + width - w if align == "right" else bx + (width - w) / 2)
                    tw_list.append((t, lx, y + k * size * LINE_SPACING, font, size, st_["color"]))
                x = bx + width

                if not f_ok:
                    notes.append("fuente similar (la original no tiene todas las letras)")
                if size < st_["size"] - 0.01:
                    notes.append(f"texto al {int(size / st_['size'] * 100)}% para que quepa")
                changes.append({
                    "page": pno, "original": seg["text"], "new": " / ".join(texts),
                    "status": "warn" if notes else "ok", "note": "; ".join(notes),
                })

            for t, tx, ty, font, size, color in tw_list:
                tw = pymupdf.TextWriter(page.rect, color=color)
                tw.append((tx, ty), t, font=font, fontsize=size)
                tw.write_text(page)

    # Quita los datos privados de Illustrator: si no, al abrir el PDF en
    # Illustrator mostraría los precios ORIGINALES (USD) en vez de los nuevos.
    try:
        doc.xref_set_key(doc.pdf_catalog(), "PieceInfo", "null")
        for page in doc:
            doc.xref_set_key(page.xref, "PieceInfo", "null")
    except Exception:
        pass

    out = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return out, changes


# ─── IMÁGENES ─────────────────────────────────────────────────
def render_images(pdf_bytes: bytes, dpi: int = 200, fmt: str = "jpg") -> list[bytes]:
    """Una imagen por página."""
    doc  = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    for page in doc:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        imgs.append(pix.tobytes("jpeg", jpg_quality=92) if fmt == "jpg" else pix.tobytes("png"))
    doc.close()
    return imgs
