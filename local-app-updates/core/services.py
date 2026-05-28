"""Capa de orquestación: combina catalog + pptx_engine para resolver
los casos de uso (scan multi-team, generate batch con streaming).

Mantiene a api.py tonta (solo traduce HTTP a llamadas Python) y a
catalog.py + pptx_engine.py aislados entre sí.
"""

from __future__ import annotations

import io
import json
import time
import zipfile
from collections.abc import Callable, Iterator
from dataclasses import dataclass

from core.catalog import CatalogSource, TeamScan, VlpsFileRef
from core.pptx_engine import (
    GenerateResult,
    LogEntry,
    MerchImage,
    MerchMap,
    PptInput,
    generate_blank_deck,
    generate_deck,
    parse_merch_name,
)


# ═══════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════

GENDER_CODE: dict[str, str] = {"MENS": "M", "WOMENS": "W", "KIDS": "K"}

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
ZIP_MIME = "application/zip"

DEFAULT_CHUNK_SIZE = 64 * 1024  # 64 KB para streaming HTTP

# Tope duro de selecciones por batch. Defensa contra batches enormes que
# tirarían timeouts / saturación de memoria. Ajustable cuando entren a
# producción y se mida la performance real con SMB.
MAX_SELECTIONS = 25


def _fmt_mb(nbytes: int) -> str:
    """Formatea bytes como MB con 2 decimales."""
    return f"{nbytes / 1_000_000:.2f}MB"


def _fmt_input_size(ppt_input) -> str:
    """Devuelve el tamaño del PPT input como string, sea bytes o un path.
    Evita cargar el path a RAM solo para loggear su tamaño.
    """
    if isinstance(ppt_input, bytes):
        return _fmt_mb(len(ppt_input))
    try:
        from pathlib import Path as _Path
        return _fmt_mb(_Path(ppt_input).stat().st_size)
    except OSError:
        return "?MB"


# ═══════════════════════════════════════════════════════════════
# PROGRESS CALLBACK
# ═══════════════════════════════════════════════════════════════
#
# Tipo del callback que generate_batch/generate_one_deck llaman para
# emitir progress events. El caller (api.py al usar jobs) mapea los
# kwargs a mutaciones del JobStatus correspondiente bajo un lock.
#
# Convención de kwargs aceptados:
#   - current_team: str          → nombre del team siendo procesado
#   - current_team_index: int    → 1-based index dentro de selections
#   - current_phase: str         → "scan" | "fetch" | "pptx" | "zip"
#   - teams_done: int            → cantidad de teams ya completados
#   - append_team_result: dict   → resultado de un team que terminó
#                                  (special: se appendea en vez de set)
#
# El callback puede ser None — los call-sites chequean antes de invocar.
ProgressCallback = Callable[..., None]


# ═══════════════════════════════════════════════════════════════
# DATACLASSES
# ═══════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class TeamSelection:
    """Una entrada del input multi-team del usuario."""
    league: str
    team: str


@dataclass(frozen=True)
class TeamGenerateResult:
    """Outcome de generar un deck para un team puntual."""
    team: str
    league: str
    success: bool
    replaced_count: int
    log: list[LogEntry]
    error: str | None = None


@dataclass(frozen=True)
class BatchOutput:
    """Metadata del output del batch. El stream de bytes va separado en
    el segundo elemento del tuple devuelto por generate_batch.
    """
    filename: str
    media_type: str
    is_zip: bool
    per_team: list[TeamGenerateResult]


# ═══════════════════════════════════════════════════════════════
# SCAN
# ═══════════════════════════════════════════════════════════════

def scan_catalog(
    catalog: CatalogSource,
    year: str,
    quarter: str,
    selections: list[TeamSelection],
) -> list[TeamScan]:
    """Itera cada selection y llama catalog.scan_team. Devuelve la lista
    de TeamScan en el mismo orden que selections.
    """
    if len(selections) > MAX_SELECTIONS:
        raise ValueError(
            f"Máximo {MAX_SELECTIONS} selecciones por batch. Recibido: {len(selections)}"
        )
    return [
        catalog.scan_team(year, quarter, s.league, s.team)
        for s in selections
    ]


# ═══════════════════════════════════════════════════════════════
# BUILD MERCH MAP
# ═══════════════════════════════════════════════════════════════

def build_merch_map(catalog: CatalogSource, team_scan: TeamScan) -> MerchMap:
    """Toma un TeamScan y construye el MerchMap que generate_deck espera.

    - Fetchea los bytes de cada ImageRef vía `catalog.fetch_images` en BATCH,
      lo que permite a la implementación SMB paralelizar el I/O.
    - Las keys se arman como 'CAPSULA CODIGO' (ej: 'FLAGSHIP M').
    - Cápsulas con image_count=0 se saltan (no aportan al map).

    El orden final del map es estable: refleja el orden en que aparecen
    las cápsulas en `team_scan.by_gender` y las imágenes dentro de cada
    cápsula. La paralelización del fetch NO altera el orden — `fetch_images`
    garantiza preservación del orden de entrada.
    """
    t_total = time.monotonic()

    # Fase 1: enumerar todos los (key, ImageRef) pares en orden estable.
    # Esto es O(N) puro Python — el costo I/O viene después.
    pairs = []  # list[tuple[str, ImageRef]]
    refs_to_fetch = []  # list[ImageRef]
    for gender, capsules in team_scan.by_gender.items():
        code = GENDER_CODE.get(gender)
        if not code:
            continue  # género no soportado, skip
        for cap in capsules:
            if cap.image_count == 0:
                continue
            key = f"{cap.capsule} {code}"
            for ref in cap.images:
                pairs.append((key, ref))
                refs_to_fetch.append(ref)

    t_enum = time.monotonic() - t_total

    # Fase 2: fetch en batch. Sobre SmbCatalogSource va en paralelo
    # con el ThreadPoolExecutor compartido; sobre MockCatalogSource es
    # secuencial (default del ABC).
    t_fetch_start = time.monotonic()
    blobs = catalog.fetch_images(refs_to_fetch)
    t_fetch = time.monotonic() - t_fetch_start

    # Fase 3: ensamblar el merch_map preservando el orden original.
    merch_map: MerchMap = {}
    for (key, ref), data in zip(pairs, blobs):
        merch_map.setdefault(key, []).append(MerchImage(
            name=ref.filename,
            bytes=data,
        ))

    # ── Diagnóstico ──
    n_imgs = len(blobs)
    total_bytes = sum(len(b) for b in blobs)
    largest = max((len(b) for b in blobs), default=0)
    avg = (total_bytes // n_imgs) if n_imgs > 0 else 0
    throughput = (total_bytes / t_fetch / 1_000_000) if t_fetch > 0 else 0
    return merch_map


# ═══════════════════════════════════════════════════════════════
# GENERATE ONE DECK
# ═══════════════════════════════════════════════════════════════

def generate_one_deck(
    catalog: CatalogSource,
    ppt_input: PptInput,
    team_scan: TeamScan,
    progress_callback: ProgressCallback | None = None,
) -> tuple[bytes, GenerateResult]:
    """Pipeline completo para un team: build_merch_map → generate_deck.
    Devuelve (deck_bytes, GenerateResult con log).

    `ppt_input` puede ser bytes (modo legacy / tests) o un path en disco
    (modo streaming desde api.py, evita cargar GB a RAM).

    Fast paths para evitar el costo de python-pptx cuando claramente
    no va a haber modificaciones — clave en PPTs grandes (cientos de MB
    a 2GB) donde parse+save sería minutos de CPU desperdiciados.
    """
    # Phase: fetch (build_merch_map fetchea las imágenes — el cuello más caro).
    if progress_callback:
        progress_callback(current_phase="fetch")
    merch_map = build_merch_map(catalog, team_scan)

    # ── Fast path #1: merch_map vacío ──
    # Si el catálogo no devolvió ninguna imagen para este team, no hay
    # nada que reemplazar. Devolvemos el PPT original sin invocar
    # python-pptx (que cargaría/serializaría todo el archivo de gratis).
    if not merch_map:
        # Si el input es un path, leemos el archivo a bytes para cumplir
        # con el contrato (caller espera bytes para meter al ZIP/stream).
        # En este caso "de gratis" pagamos la RAM porque NO hubo procesamiento.
        from core.pptx_engine import _read_ppt_bytes
        ppt_bytes_out = _read_ppt_bytes(ppt_input)
        result = GenerateResult(
            deck_bytes=ppt_bytes_out,
            replaced_count=0,
            log=[LogEntry(
                "warn",
                "Team sin imágenes en el catálogo — deck devuelto sin modificar",
            )],
        )
        return ppt_bytes_out, result

    # ── Path normal ──
    # Timing del parse + manipulación + save del PPT (CPU bound,
    # no se beneficia de paralelización de red).
    if progress_callback:
        progress_callback(current_phase="pptx")
    t_pptx = time.monotonic()
    result = generate_deck(ppt_input, merch_map)
    return result.deck_bytes, result


# ═══════════════════════════════════════════════════════════════
# GENERATE BATCH (streaming)
# ═══════════════════════════════════════════════════════════════

def _scan_team_routed(
    catalog: CatalogSource,
    year: str,
    quarter: str,
    league: str,
    team: str,
    scope: dict | None,
) -> TeamScan:
    """Despacha al método de scan correcto según haya scope o no.

    - `scope is None` → `catalog.scan_team(...)` (path histórico, sin filtrado).
    - `scope is not None` → `catalog.scan_team_scoped(year, quarter, scope, ...)`.

    Centralizamos esta decisión acá para que las dos ramas (single team,
    multi team) usen exactamente la misma lógica de routing. B.1d.2.
    """
    if scope is None:
        return catalog.scan_team(year, quarter, league, team)
    return catalog.scan_team_scoped(year, quarter, scope, league, team)


def generate_batch(
    catalog: CatalogSource,
    ppt_input: PptInput,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections: list[TeamSelection],
    progress_callback: ProgressCallback | None = None,
    scope: dict | None = None,
    multi_team_mode: str = "per_team",
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Multi-team end-to-end con output streameado.

    - 1 selection → output PPTX (bytes streamed directo). Independiente del
      multi_team_mode.
    - N selections → depende del multi_team_mode (D.5a):
        · 'strict': raise (rechaza multi-team explícitamente).
        · 'mixed': fusiona merch_maps de N teams, output 1 PPTX combinado.
        · 'per_team' (default, backward-compat): un deck por team, ZIP.

    Best-effort: si un team falla, se loguea en per_team y los demás
    siguen procesándose.

    `scope` (B.1d.2): si se pasa, las llamadas a `catalog.scan_team(...)`
    se reemplazan por `catalog.scan_team_scoped(year, quarter, scope, ...)`.
    Esto limita el `TeamScan` a las cápsulas/classics presentes en el scope,
    y todo el resto del pipeline (build_merch_map, generate_deck, find_note_match)
    se ajusta automáticamente porque solo ve un merch_map más chico — las
    slides cuya nota no matchea ninguna key se loguean como warning (ver
    `_first_merchboard_note_key` en B.1d.1).

    NOTA sobre streaming: el ZIP se construye completo en BytesIO antes
    de empezar a yieldear bytes (zipfile necesita seek para la central
    directory). El beneficio del stream está al nivel HTTP, no al nivel
    de armado. Si el batch crece a tamaños donde la memoria sea problema
    (decenas de teams × cientos de MB), conviene migrar a una librería
    como stream-zip que no requiera seek.

    Returns:
        (BatchOutput con metadata + per_team, iterator de bytes chunked).
    """
    if not selections:
        raise ValueError("selections no puede estar vacío")
    if len(selections) > MAX_SELECTIONS:
        raise ValueError(
            f"Máximo {MAX_SELECTIONS} selecciones por batch. Recibido: {len(selections)}"
        )
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise ValueError(
            f"multi_team_mode inválido: {multi_team_mode!r} "
            f"(esperaba 'strict', 'mixed' o 'per_team')"
        )

    # ── SINGLE TEAM ──
    # Cualquier modo produce el mismo output con N=1 (sin ZIP, sin merge).
    if len(selections) == 1:
        t_batch = time.monotonic()
        sel = selections[0]

        if progress_callback:
            progress_callback(
                current_team=sel.team,
                current_team_index=1,
                current_phase="scan",
                teams_done=0,
            )
        t_team_start = time.monotonic()
        t_scan = time.monotonic()
        scan = _scan_team_routed(catalog, year, quarter, sel.league, sel.team, scope)

        deck_bytes, result = generate_one_deck(
            catalog, ppt_input, scan, progress_callback=progress_callback
        )

        per_team = [TeamGenerateResult(
            team=sel.team,
            league=sel.league,
            success=True,
            replaced_count=result.replaced_count,
            log=list(result.log),
        )]
        if progress_callback:
            progress_callback(
                teams_done=1,
                append_team_result={
                    "team": sel.team,
                    "league": sel.league,
                    "success": True,
                    "replaced_count": result.replaced_count,
                    "duration_seconds": round(time.monotonic() - t_team_start, 2),
                },
            )
        output = BatchOutput(
            filename=f"{ppt_base_name} — {sel.team}.pptx",
            media_type=PPTX_MIME,
            is_zip=False,
            per_team=per_team,
        )
        return output, _stream_bytes(deck_bytes)

    # ── MULTI TEAM dispatch (D.5a) ──
    if multi_team_mode == "strict":
        teams = [f"{s.league}/{s.team}" for s in selections]
        raise ValueError(
            f"Modo 'strict' activo pero {len(selections)} teams seleccionados: "
            f"{', '.join(teams)}. Cambiá multi_team_mode a 'mixed' o "
            f"'per_team', o seleccioná un solo team."
        )
    if multi_team_mode == "mixed":
        return _generate_batch_mixed(
            catalog, ppt_input, ppt_base_name, year, quarter,
            selections, scope, progress_callback,
        )

    # multi_team_mode == "per_team" — comportamiento histórico (ZIP).
    t_batch = time.monotonic()

    leagues = {s.league for s in selections}
    league_label = next(iter(leagues)) if len(leagues) == 1 else "MULTI"

    buf = io.BytesIO()
    per_team: list[TeamGenerateResult] = []

    # ZIP_STORED en vez de ZIP_DEFLATED: las imágenes embebidas en cada
    # .pptx ya están comprimidas (JPG/PNG), re-comprimirlas a nivel ZIP
    # gasta CPU para ahorrar ~1-5%. Diagnóstico real: pasar a STORED bajó
    # zip.writestr de ~15s a ~1-2s por deck de 600MB.
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_STORED) as zf:
        for idx, sel in enumerate(selections, start=1):
            t_team = time.monotonic()
            if progress_callback:
                progress_callback(
                    current_team=sel.team,
                    current_team_index=idx,
                    current_phase="scan",
                )
            try:
                t_scan = time.monotonic()
                scan = _scan_team_routed(catalog, year, quarter, sel.league, sel.team, scope)

                deck_bytes, result = generate_one_deck(
                    catalog, ppt_input, scan, progress_callback=progress_callback
                )
                filename_in_zip = f"{ppt_base_name} — {sel.team}.pptx"

                if progress_callback:
                    progress_callback(current_phase="zip")
                t_zip = time.monotonic()
                zf.writestr(filename_in_zip, deck_bytes)

                team_duration = round(time.monotonic() - t_team, 2)
                per_team.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=True,
                    replaced_count=result.replaced_count,
                    log=list(result.log),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": sel.team,
                            "league": sel.league,
                            "success": True,
                            "replaced_count": result.replaced_count,
                            "duration_seconds": team_duration,
                        },
                    )
            except Exception as e:
                team_duration = round(time.monotonic() - t_team, 2)
                per_team.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=False,
                    replaced_count=0,
                    log=[],
                    error=str(e),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": sel.team,
                            "league": sel.league,
                            "success": False,
                            "replaced_count": 0,
                            "duration_seconds": team_duration,
                            "error": str(e),
                        },
                    )

        # Embed summary.json dentro del ZIP — el usuario lo encuentra al abrir.
        summary_json = json.dumps(
            [_team_result_to_dict(t) for t in per_team],
            indent=2,
            ensure_ascii=False,
        )
        zf.writestr("summary.json", summary_json)

    output = BatchOutput(
        filename=f"{ppt_base_name} — Batch {league_label} {year} {quarter}.zip",
        media_type=ZIP_MIME,
        is_zip=True,
        per_team=per_team,
    )
    return output, _stream_buffer(buf)


def _generate_batch_mixed(
    catalog: CatalogSource,
    ppt_input: PptInput,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections: list[TeamSelection],
    scope: dict | None,
    progress_callback: ProgressCallback | None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Mixed mode para existing-server (D.5a).

    Fetchea merch_maps de los N teams desde el catálogo, los fusiona por
    capsule key en un único merch_map, y corre `generate_deck` UNA sola
    vez. Output: 1 PPTX (no ZIP).

    Si el PPT base tiene varias slides con la misma nota (ej. 3 slides
    "FLAGSHIP M"), cada una se llena con una imagen distinta — `generate_deck`
    consume las imágenes en orden de lista. El orden es:
    selections[0]'s capsule images → selections[1]'s → ... (ordenado por
    name dentro de cada team).
    """
    t_batch = time.monotonic()

    merged_map: MerchMap = {}
    per_team_results: list[TeamGenerateResult] = []

    for idx, sel in enumerate(selections, start=1):
        t_team = time.monotonic()
        if progress_callback:
            progress_callback(
                current_team=sel.team,
                current_team_index=idx,
                current_phase="scan",
            )
        try:
            t_scan = time.monotonic()
            scan = _scan_team_routed(
                catalog, year, quarter, sel.league, sel.team, scope,
            )
            if progress_callback:
                progress_callback(current_phase="fetch")
            team_map = build_merch_map(catalog, scan)

            team_count = 0
            for cap_key, imgs in team_map.items():
                merged_map.setdefault(cap_key, []).extend(imgs)
                team_count += len(imgs)

            team_duration = round(time.monotonic() - t_team, 2)
            per_team_results.append(TeamGenerateResult(
                team=sel.team,
                league=sel.league,
                success=True,
                replaced_count=team_count,
                log=[LogEntry(
                    "ok",
                    f"{team_count} imágenes contribuidas a la mezcla",
                )],
            ))
            if progress_callback:
                progress_callback(
                    teams_done=idx,
                    append_team_result={
                        "team": sel.team,
                        "league": sel.league,
                        "success": True,
                        "replaced_count": team_count,
                        "duration_seconds": team_duration,
                    },
                )
        except Exception as e:
            team_duration = round(time.monotonic() - t_team, 2)
            per_team_results.append(TeamGenerateResult(
                team=sel.team,
                league=sel.league,
                success=False,
                replaced_count=0,
                log=[],
                error=str(e),
            ))
            if progress_callback:
                progress_callback(
                    teams_done=idx,
                    append_team_result={
                        "team": sel.team,
                        "league": sel.league,
                        "success": False,
                        "replaced_count": 0,
                        "duration_seconds": team_duration,
                        "error": str(e),
                    },
                )

    # Generar UN solo PPT con el merch_map unificado.
    if progress_callback:
        progress_callback(
            current_team="MULTI",
            current_team_index=len(selections),
            current_phase="pptx",
        )
    t_pptx = time.monotonic()
    result = generate_deck(ppt_input, merged_map)

    # Resumen "MULTI" como entry final en per_team_results — combina el
    # log de generate_deck con los teams contribuyentes.
    teams_str = ", ".join(f"{s.league}/{s.team}" for s in selections)
    multi_log = [LogEntry("ok", f"Mezcla de {len(selections)} equipos: {teams_str}")]
    multi_log.extend(result.log)
    per_team_results.append(TeamGenerateResult(
        team="MULTI",
        league="MULTI",
        success=True,
        replaced_count=result.replaced_count,
        log=multi_log,
    ))

    output = BatchOutput(
        filename=f"{ppt_base_name} — MULTI.pptx",
        media_type=PPTX_MIME,
        is_zip=False,
        per_team=per_team_results,
    )
    return output, _stream_bytes(result.deck_bytes)


# ═══════════════════════════════════════════════════════════════
# GENERATE BLANK BATCH (modo "PPT nuevo desde cero" — B.2b)
# ═══════════════════════════════════════════════════════════════
#
# Diferencias estructurales vs generate_batch:
# - No hay PPT input — los slides nacen vacíos y se llenan con imágenes.
# - scope es REQUERIDO (sin él no hay forma de saber qué cápsulas incluir,
#   porque no hay notas en un PPT base que las dicte).
# - Multi-team produce UN SOLO deck con slides interpoladas (decisión del
#   producto: un único entregable ordenado por liga/equipo/cápsula). NO
#   un ZIP de N decks como `generate_batch` multi-team.
# - Best-effort por team: si un team no tiene imágenes en el scope, se
#   saltea silenciosamente (warn en per_team[i].log) y los demás siguen.

def generate_blank_one_deck(
    catalog: CatalogSource,
    team_scan: TeamScan,
    progress_callback: ProgressCallback | None = None,
) -> tuple[bytes, GenerateResult]:
    """Pipeline para UN team en modo blank: fetch + generate_blank_deck.

    Igual semántica que `generate_one_deck` pero sin PPT input. El orden
    de slides preserva el orden natural del merch_map (gender → cápsula
    → filename, lo que ya produce `build_merch_map`).

    Devuelve (deck_bytes, GenerateResult). `replaced_count` representa
    slides creadas (no reemplazadas — el dataclass se reusa por simetría).
    """
    if progress_callback:
        progress_callback(current_phase="fetch")
    merch_map = build_merch_map(catalog, team_scan)

    # Flatten: por cada cápsula en orden, agregar todas sus imágenes en orden.
    # Esto preserva el orden estable que ya garantiza build_merch_map.
    images: list[MerchImage] = []
    for key in merch_map:
        for img in merch_map[key]:
            images.append(img)

    if progress_callback:
        progress_callback(current_phase="pptx")

    t_pptx = time.monotonic()
    result = generate_blank_deck(images)
    return result.deck_bytes, result


def generate_blank_batch(
    catalog: CatalogSource,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections: list[TeamSelection],
    scope: dict,
    progress_callback: ProgressCallback | None = None,
    multi_team_mode: str = "mixed",
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Multi-team end-to-end. Default 'mixed': un solo deck combinado (B.2).

    Modos (D.5a):
    - 'strict' + N>1: raise.
    - 'mixed' (default): un solo deck con todas las imágenes interleaved.
    - 'per_team' + N>1: N decks blank separados, ZIP. Output filename
      "{ppt_base_name} — Batch.zip", cada deck adentro "{ppt_base_name} — {team}.pptx".

    `scope` es obligatorio (en modo "PPT nuevo" no hay otra forma de saber
    qué cápsulas incluir). Si todos los teams están sin imágenes en el
    scope, el deck resultante tiene 0 slides — sigue siendo un .pptx válido.

    Returns:
        (BatchOutput con is_zip=False siempre, iterator de bytes del .pptx).
    """
    if scope is None:
        raise ValueError("scope es requerido en generate_blank_batch")
    if not selections:
        raise ValueError("selections no puede estar vacío")
    if len(selections) > MAX_SELECTIONS:
        raise ValueError(
            f"Máximo {MAX_SELECTIONS} selecciones por batch. Recibido: {len(selections)}"
        )
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise ValueError(
            f"multi_team_mode inválido: {multi_team_mode!r} "
            f"(esperaba 'strict', 'mixed' o 'per_team')"
        )

    # ── Dispatch por modo cuando N>1 (D.5a) ──
    if len(selections) > 1:
        if multi_team_mode == "strict":
            teams = [f"{s.league}/{s.team}" for s in selections]
            raise ValueError(
                f"Modo 'strict' activo pero {len(selections)} teams seleccionados: "
                f"{', '.join(teams)}. Cambiá multi_team_mode a 'mixed' o "
                f"'per_team', o seleccioná un solo team."
            )
        if multi_team_mode == "per_team":
            return _generate_blank_batch_per_team(
                catalog, ppt_base_name, year, quarter,
                selections, scope, progress_callback,
            )
        # multi_team_mode == "mixed" — comportamiento histórico, fallthrough.

    # N==1 (cualquier modo) o N>1+mixed: combinado (comportamiento existente).
    t_batch = time.monotonic()

    # Tuplas (liga, equipo, capsula, filename, MerchImage) para sort estable.
    all_entries: list[tuple[str, str, str, str, MerchImage]] = []
    per_team: list[TeamGenerateResult] = []

    for idx, sel in enumerate(selections, start=1):
        t_team = time.monotonic()
        if progress_callback:
            progress_callback(
                current_team=sel.team,
                current_team_index=idx,
                current_phase="scan",
            )
        try:
            t_scan = time.monotonic()
            scan = catalog.scan_team_scoped(year, quarter, scope, sel.league, sel.team)

            if progress_callback:
                progress_callback(current_phase="fetch")
            merch_map = build_merch_map(catalog, scan)

            team_count = 0
            for capsule_key, imgs in merch_map.items():
                for img in imgs:
                    all_entries.append(
                        (sel.league, sel.team, capsule_key, img.name, img)
                    )
                    team_count += 1

            team_duration = round(time.monotonic() - t_team, 2)

            if team_count == 0:
                # Decisión producto: skip silencioso, warn en per_team.log.
                per_team.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=True,
                    replaced_count=0,
                    log=[LogEntry(
                        "warn",
                        "Team sin imágenes en el scope — sin slides agregadas",
                    )],
                ))
            else:
                per_team.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=True,
                    replaced_count=team_count,
                    log=[LogEntry(
                        "ok",
                        f"{team_count} imágenes contribuidas al deck",
                    )],
                ))

            if progress_callback:
                progress_callback(
                    teams_done=idx,
                    append_team_result={
                        "team": sel.team,
                        "league": sel.league,
                        "success": True,
                        "replaced_count": team_count,
                        "duration_seconds": team_duration,
                    },
                )
        except Exception as e:
            team_duration = round(time.monotonic() - t_team, 2)
            per_team.append(TeamGenerateResult(
                team=sel.team,
                league=sel.league,
                success=False,
                replaced_count=0,
                log=[],
                error=str(e),
            ))
            if progress_callback:
                progress_callback(
                    teams_done=idx,
                    append_team_result={
                        "team": sel.team,
                        "league": sel.league,
                        "success": False,
                        "replaced_count": 0,
                        "duration_seconds": team_duration,
                        "error": str(e),
                    },
                )

    # ── Sort global y generate ──
    # Orden: (liga, equipo, capsule, filename). Espejo del Streamlit reference.
    all_entries.sort(key=lambda e: (e[0], e[1], e[2], e[3]))
    images_ordered = [e[4] for e in all_entries]

    if progress_callback:
        progress_callback(current_phase="pptx")

    t_pptx = time.monotonic()
    result = generate_blank_deck(images_ordered)

    # ── Filename ──
    # Single team → con nombre del equipo. Multi team → sufijo MULTI
    # (decisión de producto: nombre limpio sin enumerar 25 equipos).
    if len(selections) == 1:
        filename = f"{ppt_base_name} — {selections[0].team}.pptx"
    else:
        filename = f"{ppt_base_name} — MULTI.pptx"

    output = BatchOutput(
        filename=filename,
        media_type=PPTX_MIME,
        is_zip=False,
        per_team=per_team,
    )
    return output, _stream_bytes(result.deck_bytes)


def _generate_blank_batch_per_team(
    catalog: CatalogSource,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections: list[TeamSelection],
    scope: dict,
    progress_callback: ProgressCallback | None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Per-team mode para blank server (D.5a).

    Genera N blank decks separados (uno por team) empaquetados en ZIP.
    Cada deck contiene solo las imágenes scopeadas de su team. Espejo
    de `_generate_from_uploads_per_team` pero usando el catálogo en vez
    de uploads manuales.
    """
    t_batch = time.monotonic()
    per_team_results: list[TeamGenerateResult] = []
    buf = io.BytesIO()

    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_STORED) as zf:
        for idx, sel in enumerate(selections, start=1):
            t_team = time.monotonic()
            if progress_callback:
                progress_callback(
                    current_team=sel.team,
                    current_team_index=idx,
                    current_phase="scan",
                )
            try:
                t_scan = time.monotonic()
                scan = catalog.scan_team_scoped(year, quarter, scope, sel.league, sel.team)

                if progress_callback:
                    progress_callback(current_phase="fetch")
                team_map = build_merch_map(catalog, scan)

                # Flatten estable: por capsule key, dentro por filename.
                images: list[MerchImage] = []
                for cap_key in sorted(team_map.keys()):
                    for img in team_map[cap_key]:
                        images.append(img)

                if progress_callback:
                    progress_callback(current_phase="pptx")
                result = generate_blank_deck(images)

                filename_in_zip = f"{ppt_base_name} — {sel.team}.pptx"
                if progress_callback:
                    progress_callback(current_phase="zip")
                zf.writestr(filename_in_zip, result.deck_bytes)

                team_duration = round(time.monotonic() - t_team, 2)
                per_team_results.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=True,
                    replaced_count=result.replaced_count,
                    log=list(result.log),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": sel.team,
                            "league": sel.league,
                            "success": True,
                            "replaced_count": result.replaced_count,
                            "duration_seconds": team_duration,
                        },
                    )
            except Exception as e:
                team_duration = round(time.monotonic() - t_team, 2)
                per_team_results.append(TeamGenerateResult(
                    team=sel.team,
                    league=sel.league,
                    success=False,
                    replaced_count=0,
                    log=[],
                    error=str(e),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": sel.team,
                            "league": sel.league,
                            "success": False,
                            "replaced_count": 0,
                            "duration_seconds": team_duration,
                            "error": str(e),
                        },
                    )

        summary_json = json.dumps(
            [_team_result_to_dict(t) for t in per_team_results],
            indent=2,
            ensure_ascii=False,
        )
        zf.writestr("summary.json", summary_json)

    output = BatchOutput(
        filename=f"{ppt_base_name} — Batch.zip",
        media_type=ZIP_MIME,
        is_zip=True,
        per_team=per_team_results,
    )
    return output, _stream_buffer(buf)


# ═══════════════════════════════════════════════════════════════
# HELPERS (privados)
# ═══════════════════════════════════════════════════════════════

def _stream_bytes(data: bytes, chunk_size: int = DEFAULT_CHUNK_SIZE) -> Iterator[bytes]:
    """Yield bytes en chunks de tamaño fijo. Usado para PPTX en memoria."""
    for i in range(0, len(data), chunk_size):
        yield data[i:i + chunk_size]


def _stream_buffer(
    buf: io.BytesIO,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> Iterator[bytes]:
    """Yield chunks desde un BytesIO. Usado para ZIPs grandes."""
    buf.seek(0)
    while True:
        chunk = buf.read(chunk_size)
        if not chunk:
            return
        yield chunk


def _team_result_to_dict(t: TeamGenerateResult) -> dict:
    """Convierte TeamGenerateResult a dict serializable para summary.json."""
    return {
        "team": t.team,
        "league": t.league,
        "success": t.success,
        "replaced_count": t.replaced_count,
        "log": [{"level": e.level, "message": e.message} for e in t.log],
        "error": t.error,
    }


# ═══════════════════════════════════════════════════════════════
# GENERATE FROM UPLOADS (modo "manual" — D.1)
# ═══════════════════════════════════════════════════════════════
#
# Path para cuando el usuario sube las imágenes directamente (sin pasar
# por el catálogo SMB). Los filenames se parsean con `parse_merch_name`
# para extraer (liga, team, capsule, gender). El resultado se agrupa
# por team (data shape "multi-ready") y se dispatcha:
#
# - output_mode='existing' → generate_deck con el merch_map + PPT base.
# - output_mode='blank'    → generate_blank_deck con la lista de imágenes.
#
# D.1 enforza SINGLE-team (raise si detecta >1). El data shape ya está
# preparado para multi-team — cuando llegue, se loopea sobre los teams
# y se empaqueta el output (ZIP para existing, deck combinado para blank).

def parse_uploads_to_team_merch_maps(
    uploads: list[tuple[str, bytes]],
) -> tuple[dict[tuple[str, str], MerchMap], list[str]]:
    """Parsea filenames y agrupa imágenes por (liga, team).

    Data shape multi-ready: el resultado es un dict indexado por la
    tupla (liga, team), y dentro de cada team hay un MerchMap normal
    (dict[capsule_key, list[MerchImage]]).

    Args:
        uploads: lista de tuplas (filename, bytes) — el caller ya leyó
            las imágenes a memoria.

    Returns:
        - team_merch_maps: {(liga, team): {capsule_key: [MerchImage, ...]}}.
        - unmatched: filenames que `parse_merch_name` rechazó (formato inválido).

    Para naming válido se espera: LIGA_EQUIPO_CAPSULA[_GENERO][_NNN].ext
    Ej: NFL_DALLAS COWBOYS_FLAGSHIP_M_001.jpg
    """
    team_merch_maps: dict[tuple[str, str], MerchMap] = {}
    unmatched: list[str] = []
    for filename, blob in uploads:
        parsed = parse_merch_name(filename)
        if not parsed or not parsed.get("capsule"):
            unmatched.append(filename)
            continue
        team_key = (parsed["liga"], parsed["team"])
        capsule_key = parsed["key"]
        team_merch_maps.setdefault(team_key, {}).setdefault(
            capsule_key, []
        ).append(MerchImage(name=filename, bytes=blob))
    # Orden estable dentro de cada cápsula: por filename.
    for tk in team_merch_maps:
        for ck in team_merch_maps[tk]:
            team_merch_maps[tk][ck].sort(key=lambda img: img.name)
    return team_merch_maps, unmatched


# ─── Helpers internos por modo (D.4a) ────────────────────────────

def _unmatched_warn_entry(unmatched: list[str]) -> LogEntry | None:
    """Genera el LogEntry "warn" que reporta los filenames sin parsing
    válido, o None si la lista está vacía. Truncamos a los primeros 5
    para no inflar logs cuando son muchos."""
    if not unmatched:
        return None
    sample = ", ".join(unmatched[:5])
    suffix = f" (y {len(unmatched) - 5} más)" if len(unmatched) > 5 else ""
    return LogEntry(
        "warn",
        f"{len(unmatched)} archivo(s) sin naming válido: {sample}{suffix}",
    )


def _generate_from_uploads_single(
    team_merch_maps: dict[tuple[str, str], MerchMap],
    unmatched: list[str],
    output_mode: str,
    ppt_input: PptInput | None,
    ppt_base_name: str,
    progress_callback: ProgressCallback | None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Single-team: 1 team → 1 PPTX. Usado por modo 'strict' y por
    'per_team' cuando N=1 (no tiene sentido empaquetar 1 deck en ZIP)."""
    (league, team), merch_map = next(iter(team_merch_maps.items()))
    t_start = time.monotonic()

    if progress_callback:
        progress_callback(
            current_team=team,
            current_team_index=1,
            current_phase="pptx",
        )

    if output_mode == "existing":
        result = generate_deck(ppt_input, merch_map)
    else:  # blank
        images: list[MerchImage] = []
        for key in sorted(merch_map.keys()):
            for img in merch_map[key]:
                images.append(img)
        result = generate_blank_deck(images)

    duration = round(time.monotonic() - t_start, 2)

    log_entries: list[LogEntry] = []
    unmatched_warn = _unmatched_warn_entry(unmatched)
    if unmatched_warn:
        log_entries.append(unmatched_warn)
    log_entries.extend(result.log)

    per_team = [TeamGenerateResult(
        team=team,
        league=league,
        success=True,
        replaced_count=result.replaced_count,
        log=log_entries,
    )]

    if progress_callback:
        progress_callback(
            teams_done=1,
            append_team_result={
                "team": team,
                "league": league,
                "success": True,
                "replaced_count": result.replaced_count,
                "duration_seconds": duration,
            },
        )

    base = ppt_base_name or "Deck"
    filename = f"{base} — {team}.pptx"
    output = BatchOutput(
        filename=filename,
        media_type=PPTX_MIME,
        is_zip=False,
        per_team=per_team,
    )
    return output, _stream_bytes(result.deck_bytes)


def _generate_from_uploads_mixed(
    team_merch_maps: dict[tuple[str, str], MerchMap],
    unmatched: list[str],
    output_mode: str,
    ppt_input: PptInput | None,
    ppt_base_name: str,
    progress_callback: ProgressCallback | None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Mezclado: aplana todos los team_merch_maps en un único merch_map
    indexado por capsule key. Slides con la misma nota se llenan en orden
    con imágenes de distintos equipos. Output: 1 PPTX."""
    # Aplanar: el team de origen se pierde a este nivel (los engines no
    # diferencian por team — solo por capsule key del merch_map).
    merged: MerchMap = {}
    for tk, team_map in team_merch_maps.items():
        for cap_key, imgs in team_map.items():
            merged.setdefault(cap_key, []).extend(imgs)
    # Re-sort estable por filename dentro de cada cápsula (los engines
    # consumen en orden de lista).
    for k in merged:
        merged[k].sort(key=lambda i: i.name)

    n_teams = len(team_merch_maps)
    teams_str = ", ".join(f"{l}/{t}" for l, t in sorted(team_merch_maps.keys()))
    t_start = time.monotonic()

    if progress_callback:
        progress_callback(
            current_team="MULTI",
            current_team_index=1,
            current_phase="pptx",
        )

    if output_mode == "existing":
        result = generate_deck(ppt_input, merged)
    else:  # blank
        images: list[MerchImage] = []
        for key in sorted(merged.keys()):
            for img in merged[key]:
                images.append(img)
        result = generate_blank_deck(images)

    duration = round(time.monotonic() - t_start, 2)

    log_entries: list[LogEntry] = []
    unmatched_warn = _unmatched_warn_entry(unmatched)
    if unmatched_warn:
        log_entries.append(unmatched_warn)
    log_entries.append(LogEntry(
        "ok",
        f"Mezcla de {n_teams} equipos: {teams_str}",
    ))
    log_entries.extend(result.log)

    per_team = [TeamGenerateResult(
        team="MULTI",
        league="MULTI",
        success=True,
        replaced_count=result.replaced_count,
        log=log_entries,
    )]

    if progress_callback:
        progress_callback(
            teams_done=1,
            append_team_result={
                "team": "MULTI",
                "league": "MULTI",
                "success": True,
                "replaced_count": result.replaced_count,
                "duration_seconds": duration,
            },
        )

    base = ppt_base_name or "Deck"
    filename = f"{base} — MULTI.pptx"
    output = BatchOutput(
        filename=filename,
        media_type=PPTX_MIME,
        is_zip=False,
        per_team=per_team,
    )
    return output, _stream_bytes(result.deck_bytes)


def _generate_from_uploads_per_team(
    team_merch_maps: dict[tuple[str, str], MerchMap],
    unmatched: list[str],
    output_mode: str,
    ppt_input: PptInput | None,
    ppt_base_name: str,
    progress_callback: ProgressCallback | None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Per-team: 1 deck por team, todos empaquetados en ZIP con summary.json.
    Solo se llama cuando N > 1 (single-team va por _generate_from_uploads_single).
    Best-effort: si la generación de un team falla, los demás siguen y se
    reporta en per_team[i].error."""
    t_batch = time.monotonic()
    teams_sorted = sorted(team_merch_maps.keys())
    n_teams = len(teams_sorted)

    per_team_results: list[TeamGenerateResult] = []
    buf = io.BytesIO()

    # ZIP_STORED: las imágenes embebidas en cada .pptx ya están comprimidas
    # (JPG/PNG). Comprimir el zip a otro nivel solo gasta CPU sin ganancia.
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_STORED) as zf:
        for idx, team_key in enumerate(teams_sorted, start=1):
            (league, team) = team_key
            merch_map = team_merch_maps[team_key]
            t_team = time.monotonic()
            if progress_callback:
                progress_callback(
                    current_team=team,
                    current_team_index=idx,
                    current_phase="pptx",
                )
            try:
                if output_mode == "existing":
                    result = generate_deck(ppt_input, merch_map)
                else:  # blank
                    images: list[MerchImage] = []
                    for key in sorted(merch_map.keys()):
                        for img in merch_map[key]:
                            images.append(img)
                    result = generate_blank_deck(images)

                deck_bytes = result.deck_bytes
                filename_in_zip = f"{ppt_base_name} — {team}.pptx"
                zf.writestr(filename_in_zip, deck_bytes)

                team_duration = round(time.monotonic() - t_team, 2)
                per_team_results.append(TeamGenerateResult(
                    team=team,
                    league=league,
                    success=True,
                    replaced_count=result.replaced_count,
                    log=list(result.log),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": team,
                            "league": league,
                            "success": True,
                            "replaced_count": result.replaced_count,
                            "duration_seconds": team_duration,
                        },
                    )
            except Exception as e:
                team_duration = round(time.monotonic() - t_team, 2)
                per_team_results.append(TeamGenerateResult(
                    team=team,
                    league=league,
                    success=False,
                    replaced_count=0,
                    log=[],
                    error=str(e),
                ))
                if progress_callback:
                    progress_callback(
                        teams_done=idx,
                        append_team_result={
                            "team": team,
                            "league": league,
                            "success": False,
                            "replaced_count": 0,
                            "duration_seconds": team_duration,
                            "error": str(e),
                        },
                    )

        # Si hubo unmatched, lo anexamos al log del primer team result.
        unmatched_warn = _unmatched_warn_entry(unmatched)
        if unmatched_warn and per_team_results:
            first = per_team_results[0]
            per_team_results[0] = TeamGenerateResult(
                team=first.team,
                league=first.league,
                success=first.success,
                replaced_count=first.replaced_count,
                log=[unmatched_warn] + first.log,
                error=first.error,
            )

        summary_json = json.dumps(
            [_team_result_to_dict(t) for t in per_team_results],
            indent=2,
            ensure_ascii=False,
        )
        zf.writestr("summary.json", summary_json)

    output = BatchOutput(
        filename=f"{ppt_base_name} — Batch.zip",
        media_type=ZIP_MIME,
        is_zip=True,
        per_team=per_team_results,
    )
    return output, _stream_buffer(buf)


def generate_from_uploads(
    uploads: list[tuple[str, bytes]],
    output_mode: str,
    ppt_input: PptInput | None = None,
    ppt_base_name: str = "",
    progress_callback: ProgressCallback | None = None,
    multi_team_mode: str = "strict",
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Genera un deck desde imágenes subidas por el usuario (sin catálogo).

    Args:
        uploads: (filename, bytes) por imagen subida.
        output_mode: 'existing' (reemplaza imágenes en `ppt_input`) o
            'blank' (genera deck nuevo desde cero).
        ppt_input: PPT base (bytes/path). Requerido si output_mode='existing'.
        ppt_base_name: stem del filename del output.
        progress_callback: callback opcional para reportar progreso.
        multi_team_mode (D.4a): cómo manejar uploads de varios teams:
            - 'strict' (default): si detecta >1 team, levanta ValueError.
              Conservador — fuerza al usuario a confirmar que es lo que quiere.
            - 'mixed': aplana las imágenes en un único merch_map indexado
              solo por capsule key. Output: 1 PPTX con slides interleaved.
            - 'per_team': genera 1 deck por team. Output: ZIP con N decks
              + summary.json. Si N=1, output PPTX (no ZIP).

    Returns:
        (BatchOutput, stream de bytes).

    Raises:
        ValueError: validación de input falló (modo inválido, sin imágenes,
            multi-team con mode='strict', etc.).
    """
    # ── Validación ──
    if not uploads:
        raise ValueError("uploads no puede estar vacío")
    if output_mode not in ("existing", "blank"):
        raise ValueError(
            f"output_mode inválido: {output_mode!r} "
            f"(esperaba 'existing' o 'blank')"
        )
    if output_mode == "existing" and ppt_input is None:
        raise ValueError("output_mode='existing' requiere ppt_input")
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise ValueError(
            f"multi_team_mode inválido: {multi_team_mode!r} "
            f"(esperaba 'strict', 'mixed' o 'per_team')"
        )

    # ── Parse + agrupación por team ──
    team_merch_maps, unmatched = parse_uploads_to_team_merch_maps(uploads)

    if not team_merch_maps:
        raise ValueError(
            "Ninguna imagen tiene naming válido. "
            "Formato esperado: LIGA_EQUIPO_CAPSULA[_GENERO][_NNN].ext "
            f"({len(unmatched)} archivos rechazados)"
        )

    n_teams = len(team_merch_maps)

    # ── Dispatch por multi_team_mode ──
    if multi_team_mode == "strict":
        if n_teams > 1:
            teams = [f"{l}/{t}" for l, t in team_merch_maps.keys()]
            raise ValueError(
                f"Modo 'strict' activo pero detectados {n_teams} teams: "
                f"{', '.join(teams)}. Cambiá el modo de procesamiento "
                f"(Mezcla o Per-team) o subí archivos de un solo team."
            )
        return _generate_from_uploads_single(
            team_merch_maps, unmatched, output_mode,
            ppt_input, ppt_base_name, progress_callback,
        )

    if multi_team_mode == "mixed":
        return _generate_from_uploads_mixed(
            team_merch_maps, unmatched, output_mode,
            ppt_input, ppt_base_name, progress_callback,
        )

    # multi_team_mode == "per_team"
    if n_teams == 1:
        # N=1 con per_team: comportamiento idéntico a strict — no tiene
        # sentido empaquetar 1 deck en ZIP. Output PPTX directo.
        return _generate_from_uploads_single(
            team_merch_maps, unmatched, output_mode,
            ppt_input, ppt_base_name, progress_callback,
        )
    return _generate_from_uploads_per_team(
        team_merch_maps, unmatched, output_mode,
        ppt_input, ppt_base_name, progress_callback,
    )


# ═══════════════════════════════════════════════════════════════
# GATHER VLPS FILES (modo "descarga de archivos pre-armados" — G.1)
# ═══════════════════════════════════════════════════════════════
#
# Diferencias estructurales con los generate_* anteriores:
# - No genera nada: solo busca archivos y los empaqueta en ZIP.
# - No tiene PPT base ni merch_map ni reemplazos.
# - Output: SIEMPRE ZIP (incluso si hay un solo archivo).
# - Estructura interna del ZIP: {liga or "_GENERIC"}/{cápsula}/filename.
# - Sin nivel de team — los archivos vienen agrupados por liga.

VLPS_DEFAULT_BASE_NAME = "Pro Standard — VLPS"


def gather_vlps_files(
    catalog: CatalogSource,
    year: str,
    quarter: str,
    scope: dict,
    leagues: list[str],
    file_types: list[str],
    ppt_base_name: str = VLPS_DEFAULT_BASE_NAME,
    progress_callback: ProgressCallback | None = None,
) -> tuple[BatchOutput, Iterator[bytes]]:
    """Empaqueta archivos VLPS encontrados en un ZIP descargable.

    Args:
        catalog: fuente del catálogo.
        year, quarter, scope: igual semántica que los otros endpoints.
        leagues: ligas a incluir. Archivos bajo `{vlps_dir}/{liga}/` solo
            se incluyen si la liga está acá. Archivos FLAT (sin sub-liga)
            siempre se incluyen y van a la "_GENERIC" en el ZIP.
        file_types: subset de `['ppt', 'pdf']`. Determina qué buscar.
        ppt_base_name: prefijo del filename del ZIP de salida.

    Returns:
        (BatchOutput con is_zip=True, iterator de bytes del ZIP).

    Raises:
        ValueError: validación de input falló.
    """
    # Nota: leagues vacío es VÁLIDO — significa "incluir todos los archivos"
    # del scope sin filtrar por liga (ej. HEADWEAR sin estructura _PPTX VLPS/{liga}/).
    if not file_types or not all(ft in ("ppt", "pdf") for ft in file_types):
        raise ValueError(
            f"file_types inválido: {file_types!r} (esperaba subset de ['ppt', 'pdf'])"
        )
    if scope is None:
        raise ValueError("scope es requerido")

    t_batch = time.monotonic()

    # ── FASE 1: SCAN ──
    if progress_callback:
        progress_callback(current_phase="scan")
    refs = catalog.scan_vlps_files(year, quarter, scope, leagues, file_types)

    # ── FASE 2: FETCH ──
    # Si no hay refs, generamos ZIP igual con solo summary.json explicando.
    if progress_callback:
        progress_callback(current_phase="fetch")
    t_fetch = time.monotonic()
    blobs = catalog.fetch_vlps_files(refs)

    # ── FASE 3: ZIP ──
    if progress_callback:
        progress_callback(current_phase="zip")
    t_zip = time.monotonic()
    buf = io.BytesIO()
    files_by_league: dict[str, int] = {}
    bytes_by_league: dict[str, int] = {}
    files_by_type: dict[str, int] = {"ppt": 0, "pdf": 0}

    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_STORED) as zf:
        for ref, blob in zip(refs, blobs):
            # FLAT files (sin liga) van a _GENERIC en el ZIP.
            liga_segment = ref.league or "_GENERIC"
            # Preservamos el filename ORIGINAL (decisión usuario, pregunta 3).
            archive_name = f"{liga_segment}/{ref.capsule}/{ref.filename}"
            zf.writestr(archive_name, blob)
            files_by_league[liga_segment] = files_by_league.get(liga_segment, 0) + 1
            bytes_by_league[liga_segment] = bytes_by_league.get(liga_segment, 0) + len(blob)
            files_by_type[ref.file_type] = files_by_type.get(ref.file_type, 0) + 1

        # Summary embed para que el usuario tenga contexto al abrir el ZIP.
        summary = {
            "year": year,
            "quarter": quarter,
            "leagues_requested": leagues,
            "file_types_requested": file_types,
            "total_files": len(refs),
            "files_by_type": files_by_type,
            "files_by_league": files_by_league,
        }
        zf.writestr(
            "summary.json",
            json.dumps(summary, indent=2, ensure_ascii=False),
        )


    # ── per_team-style results (uno por liga + GENERIC) para la UI ──
    per_team_results: list[TeamGenerateResult] = []
    # Orden estable: ligas alfabéticas + GENERIC al final.
    league_keys = sorted(k for k in files_by_league.keys() if k != "_GENERIC")
    if "_GENERIC" in files_by_league:
        league_keys.append("_GENERIC")

    if not league_keys:
        # Cero archivos encontrados — un solo summary "fila" para que el UI
        # tenga algo que mostrar.
        per_team_results.append(TeamGenerateResult(
            team="(sin archivos)",
            league="(sin archivos)",
            success=True,
            replaced_count=0,
            log=[LogEntry(
                "warn",
                "No se encontraron archivos VLPS que matcheen el scope + ligas",
            )],
        ))
    else:
        for liga in league_keys:
            count = files_by_league[liga]
            size_bytes = bytes_by_league[liga]
            per_team_results.append(TeamGenerateResult(
                team=liga,        # Usamos liga como "team" para reuso de UI.
                league=liga,
                success=True,
                replaced_count=count,
                log=[LogEntry(
                    "ok",
                    f"{count} archivo(s) · {_fmt_mb(size_bytes)}",
                )],
            ))

    if progress_callback:
        # Reporte final agregado para la UI.
        for liga in league_keys:
            progress_callback(
                append_team_result={
                    "team": liga,
                    "league": liga,
                    "success": True,
                    "replaced_count": files_by_league[liga],
                    "duration_seconds": round(time.monotonic() - t_batch, 2),
                },
            )
        progress_callback(teams_done=max(1, len(league_keys)))

    output = BatchOutput(
        filename=f"{ppt_base_name} — {year} {quarter}.zip",
        media_type=ZIP_MIME,
        is_zip=True,
        per_team=per_team_results,
    )
    return output, _stream_buffer(buf)
