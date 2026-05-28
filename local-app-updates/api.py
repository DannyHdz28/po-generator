"""FastAPI app: routes, auth, dependency injection del CatalogSource."""

from __future__ import annotations

import json
import logging
import os
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import (
    Cookie,
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
    status,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, ValidationError

from config import get_settings
from core.catalog import CatalogSource, get_catalog_source
from core.pptx_engine import scan_pptx
from core.services import (
    TeamSelection,
    gather_vlps_files,
    generate_batch,
    generate_blank_batch,
    generate_from_uploads,
    scan_catalog,
)

# D.1 — tope duro de imágenes por upload manual. Defensa contra
# batches enormes; ajustar cuando se mida uso real.
MAX_MANUAL_UPLOADS = 200


# ═══════════════════════════════════════════════════════════════
# Logging y settings (fail-fast)
# ═══════════════════════════════════════════════════════════════

logger = logging.getLogger("deck-builder")

# Si la config es inválida, get_settings() levanta ValidationError aquí
# (al importar el módulo), antes de que uvicorn termine de arrancar.
_settings = get_settings()

if _settings.catalog_source == "mock" and not _settings.debug:
    logger.warning(
        "Running with MOCK catalog source in non-debug mode. "
        "Set CATALOG_SOURCE=smb for production."
    )


def _check_workers_count() -> None:
    """Loggea un warning ruidoso si la app fue lanzada con >1 worker.

    Detección heurística vía env vars `WEB_CONCURRENCY` o `UVICORN_WORKERS`,
    que son las convenciones más comunes (gunicorn, Heroku, deploys con
    uvicorn detrás de gunicorn). Si nada está seteado, asume 1 worker y
    no loggea nada.

    Por qué importa: el registry global `_JOBS` y su lock viven en memoria
    del proceso. Con >1 worker, cada proceso tiene su propio `_JOBS` y los
    clientes ven '404 Job no encontrado' al polear el status de un job
    creado en otro worker. Ver comentario completo en la sección JOBS.

    Soft warning (no fail-fast) porque algunos deploys ya configuran 2
    workers por convención corporativa antes de que migremos `_JOBS` a un
    store compartido (sticky sessions o SQLite). El fix de fondo está
    documentado en PROJECT_STATUS.md.
    """
    raw = os.getenv("WEB_CONCURRENCY") or os.getenv("UVICORN_WORKERS")
    if raw is None:
        return
    try:
        n = int(raw)
    except ValueError:
        return
    if n > 1:
        logger.warning(
            "RIESGO: la app corre con %d workers pero _JOBS es in-memory "
            "por proceso. Jobs creados en un worker no son visibles desde "
            "otro. Sintoma: clientes ven 'Job no encontrado (404)' al "
            "polear status. Fix: sticky sessions en reverse proxy o migrar "
            "_JOBS a SQLite. Ver PROJECT_STATUS.md.",
            n,
        )


_check_workers_count()


# ═══════════════════════════════════════════════════════════════
# App + middlewares
# ═══════════════════════════════════════════════════════════════

app = FastAPI(title="Deck Builder", version="0.1.0")

if _settings.cors_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_settings.cors_origins,
        allow_credentials=True,
        allow_methods=["GET", "POST"],
        allow_headers=["*"],
    )

# Frontend
app.mount("/static", StaticFiles(directory="ui/static"), name="static")
templates = Jinja2Templates(directory="ui/templates")


# ═══════════════════════════════════════════════════════════════
# Dependencies
# ═══════════════════════════════════════════════════════════════

# auto_error=False: no levantar 401 si falta el header, así caemos al
# fallback de cookie. Si ninguno está presente, validamos a mano abajo.
_bearer = HTTPBearer(auto_error=False)

# Nombre del cookie usado para autenticar descargas via <a href>. El JS no
# puede setear `Authorization` en un click nativo, así que el server setea
# este cookie al servir `/` y los browsers lo mandan automáticamente en
# las descargas. Path "/api/" lo restringe a endpoints del API.
COOKIE_AUTH_NAME = "deck_session"


def verify_token(
    creds: HTTPAuthorizationCredentials | None = Depends(_bearer),
    deck_session: str | None = Cookie(default=None, alias=COOKIE_AUTH_NAME),
) -> None:
    """Valida el bearer token desde header `Authorization` o desde cookie.

    Backward compatible: si el cliente manda `Authorization: Bearer X` (caso
    normal del JS para fetches autenticados, scripts curl, tests, etc.),
    se usa eso. Si no manda header pero sí cookie (caso del `<a href>` que
    dispara descarga nativa del browser, sin posibilidad de setear headers
    custom), se usa el cookie.

    Si ninguno coincide o no llega ninguno, 401.
    """
    expected = get_settings().api_token
    if creds is not None and creds.credentials == expected:
        return
    if deck_session is not None and deck_session == expected:
        return
    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Invalid bearer token",
    )


@lru_cache(maxsize=1)
def get_catalog() -> CatalogSource:
    """Singleton del CatalogSource según config. Cacheado: 1 instancia por proceso.

    El caché interno del SmbCatalogSource depende de este singleton — si
    se recreara la instancia por request, el TTLCache se perdería y los
    settings de cache_ttl/cache_maxsize no servirían para nada.
    """
    s = get_settings()
    return get_catalog_source(
        s.catalog_source,
        s.server_base,
        cache_ttl=s.catalog_cache_ttl_seconds,
        cache_maxsize=s.catalog_cache_maxsize,
        walk_workers=s.catalog_walk_workers,
    )


# ═══════════════════════════════════════════════════════════════
# JOBS — Generación asincrónica con polling
# ═══════════════════════════════════════════════════════════════
#
# POST /api/generate ya no devuelve el deck. En su lugar:
# 1. Stream del PPT a un tempfile.
# 2. Crea un JobStatus en _JOBS y spawnea un thread worker.
# 3. Devuelve 202 + {job_id, status_url}.
# 4. El cliente polea /api/jobs/{id}/status hasta status=done.
# 5. El cliente GET /api/jobs/{id}/download para bajar el archivo.
# 6. Tras descarga exitosa, output + JobStatus se borran de inmediato.
# 7. Como fallback, un reaper limpia jobs done/failed que pasan el TTL.
#
# IMPORTANTE: _JOBS es in-memory por proceso. Para escalar a múltiples
# uvicorn workers habría que mover a Redis. Para 2-3 usuarios concurrentes
# con `uvicorn ... --workers 1` (default), esto alcanza.


@dataclass
class JobStatus:
    """Estado mutable de un job de generación."""
    job_id: str
    status: str = "queued"  # "queued" | "running" | "done" | "failed"
    teams_total: int = 0
    teams_done: int = 0
    current_team: str | None = None
    current_team_index: int | None = None  # 1-based
    current_phase: str | None = None        # "scan" | "fetch" | "pptx" | "zip"
    started_at: float = field(default_factory=time.time)
    finished_at: float | None = None
    output_path: Path | None = None
    output_filename: str | None = None
    output_media_type: str | None = None
    per_team_results: list[dict] = field(default_factory=list)
    error: str | None = None

    def to_public_dict(self) -> dict:
        """Serializa para el endpoint /status — incluye fields seguros para
        el cliente, omite paths internos del filesystem.
        """
        elapsed = (
            (self.finished_at or time.time()) - self.started_at
        )
        return {
            "job_id": self.job_id,
            "status": self.status,
            "teams_total": self.teams_total,
            "teams_done": self.teams_done,
            "current_team": self.current_team,
            "current_team_index": self.current_team_index,
            "current_phase": self.current_phase,
            "elapsed_seconds": round(elapsed, 2),
            "per_team_results": list(self.per_team_results),
            "error": self.error,
            "download_url": (
                f"/api/jobs/{self.job_id}/download"
                if self.status == "done"
                else None
            ),
            "output_filename": self.output_filename,
        }


# Registry global de jobs. Lock protege tanto el dict como las mutaciones
# del JobStatus (cualquier read/write debe pasar por _JOBS_LOCK).
_JOBS: dict[str, JobStatus] = {}
_JOBS_LOCK = threading.Lock()


def _update_job(job_id: str, **updates: Any) -> None:
    """Aplica `updates` al JobStatus dado bajo el lock.

    Reglas especiales:
    - `append_team_result=<dict>` appendea a `per_team_results` en vez de
      reemplazarlo (cada team termina y se acumula).
    - Cualquier otra key se setea directo con setattr.
    """
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is None:
            return
        for k, v in updates.items():
            if k == "append_team_result":
                job.per_team_results.append(v)
            else:
                setattr(job, k, v)


def _generate_worker(
    job_id: str,
    ppt_path: Path,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections_list: list[TeamSelection],
    catalog: CatalogSource,
    scope: dict | None = None,
    multi_team_mode: str = "per_team",
) -> None:
    """Corre la generación completa en un thread daemon.

    Lee el PPT desde `ppt_path`, llama generate_batch con progress_callback,
    consume el stream del output a un tempfile en disco, y actualiza el
    JobStatus con cada paso. Maneja errores marcando el job como "failed".
    Siempre limpia el tempfile del upload al terminar.

    `scope` (B.1d.2): si está presente (modo avanzado), se pasa a
    `generate_batch` que lo usa para llamar `catalog.scan_team_scoped(...)`
    en vez de `catalog.scan_team(...)`. `None` preserva el path histórico.

    `multi_team_mode` (D.5a): controla cómo se procesa N>1 teams.
    Default 'per_team' = ZIP histórico. Otras opciones: 'mixed' (1 deck
    combinado) o 'strict' (raise).
    """
    output_path: Path | None = None
    try:
        _update_job(job_id, status="running", started_at=time.time())

        def progress_cb(**updates: Any) -> None:
            _update_job(job_id, **updates)

        output, stream = generate_batch(
            catalog,
            ppt_path,
            ppt_base_name,
            year,
            quarter,
            selections_list,
            progress_callback=progress_cb,
            scope=scope,
            multi_team_mode=multi_team_mode,
        )

        # Consume el stream a un tempfile en disco — evita acumular GBs
        # en RAM mientras la descarga del cliente no haya empezado.
        suffix = ".zip" if output.is_zip else ".pptx"
        fd, output_path_str = tempfile.mkstemp(suffix=suffix)
        output_path = Path(output_path_str)
        with os.fdopen(fd, "wb") as f:
            for chunk in stream:
                f.write(chunk)

        _update_job(
            job_id,
            status="done",
            finished_at=time.time(),
            output_path=output_path,
            output_filename=output.filename,
            output_media_type=output.media_type,
            current_team=None,
            current_phase=None,
        )
    except Exception as e:
        logger.exception("Job %s failed", job_id)
        _update_job(
            job_id,
            status="failed",
            finished_at=time.time(),
            error=str(e),
        )
        # Cleanup del output parcial si quedó algo escrito.
        if output_path is not None and output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass
    finally:
        # Limpiamos el tempfile del upload SIEMPRE, haya éxito o no.
        try:
            ppt_path.unlink(missing_ok=True)
        except OSError:
            pass


def _generate_blank_worker(
    job_id: str,
    ppt_base_name: str,
    year: str,
    quarter: str,
    selections_list: list[TeamSelection],
    catalog: CatalogSource,
    scope: dict,
    multi_team_mode: str = "mixed",
) -> None:
    """Variante del worker para modo "PPT nuevo desde cero" (B.2c + D.5a).

    Diferencias con `_generate_worker`:
    - No recibe `ppt_path: Path` — no hay upload de PPT.
    - No hace cleanup de tempfile del upload (no existe).
    - Llama `generate_blank_batch` en vez de `generate_batch`.
    - `scope` es obligatorio (no Optional) — sin él no hay forma de saber
      qué cápsulas incluir.
    - `multi_team_mode` default 'mixed' = comportamiento histórico (1 deck
      combinado). Otras opciones: 'strict' (raise) o 'per_team' (ZIP).

    Reusa _JOBS, _update_job y el flujo de status/download tal cual, así
    que el cliente no ve diferencia más allá del endpoint inicial.
    """
    output_path: Path | None = None
    try:
        _update_job(job_id, status="running", started_at=time.time())

        def progress_cb(**updates: Any) -> None:
            _update_job(job_id, **updates)

        output, stream = generate_blank_batch(
            catalog,
            ppt_base_name,
            year,
            quarter,
            selections_list,
            scope,
            progress_callback=progress_cb,
            multi_team_mode=multi_team_mode,
        )

        # Consume el stream a un tempfile en disco. El suffix depende del
        # modo: per_team produce ZIP, mixed/strict producen PPTX. Lo
        # decide output.is_zip que ya viene seteado por generate_blank_batch.
        suffix = ".zip" if output.is_zip else ".pptx"
        fd, output_path_str = tempfile.mkstemp(suffix=suffix)
        output_path = Path(output_path_str)
        with os.fdopen(fd, "wb") as f:
            for chunk in stream:
                f.write(chunk)

        _update_job(
            job_id,
            status="done",
            finished_at=time.time(),
            output_path=output_path,
            output_filename=output.filename,
            output_media_type=output.media_type,
            current_team=None,
            current_phase=None,
        )
    except Exception as e:
        logger.exception("Blank job %s failed", job_id)
        _update_job(
            job_id,
            status="failed",
            finished_at=time.time(),
            error=str(e),
        )
        if output_path is not None and output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass


def _generate_from_uploads_worker(
    job_id: str,
    images_data: list[tuple[str, bytes]],
    output_mode: str,
    ppt_path: Path | None,
    ppt_base_name: str,
    multi_team_mode: str = "strict",
) -> None:
    """Worker para modo upload manual (D.1 + D.4a).

    Diferencias con los otros workers:
    - Recibe los bytes de las imágenes en memoria (ya leídos por el endpoint).
    - El PPT base es opcional (solo en output_mode='existing').
    - No usa catálogo — `generate_from_uploads` arma el merch_map del
      filename de cada imagen.
    - `multi_team_mode` (D.4a): cómo manejar uploads de varios teams.
      Default 'strict' = comportamiento conservador, falla si N>1.
      Las otras opciones ('mixed', 'per_team') las elige el usuario en UI.

    Cleanup:
    - `ppt_path` (si existe): el endpoint lo creó como tempfile del upload.
      Acá lo borramos en `finally`.
    - `images_data`: bytes en memoria — Python los recolecta solo cuando
      el worker termina.
    """
    output_path: Path | None = None
    try:
        _update_job(job_id, status="running", started_at=time.time())

        def progress_cb(**updates: Any) -> None:
            _update_job(job_id, **updates)

        output, stream = generate_from_uploads(
            uploads=images_data,
            output_mode=output_mode,
            ppt_input=ppt_path,
            ppt_base_name=ppt_base_name,
            progress_callback=progress_cb,
            multi_team_mode=multi_team_mode,
        )

        # Suffix dinámico: ZIP en per_team con N>1, PPTX en strict/mixed/per_team
        # con N=1. Lo decide el modo de salida via output.is_zip.
        suffix = ".zip" if output.is_zip else ".pptx"
        fd, output_path_str = tempfile.mkstemp(suffix=suffix)
        output_path = Path(output_path_str)
        with os.fdopen(fd, "wb") as f:
            for chunk in stream:
                f.write(chunk)

        _update_job(
            job_id,
            status="done",
            finished_at=time.time(),
            output_path=output_path,
            output_filename=output.filename,
            output_media_type=output.media_type,
            current_team=None,
            current_phase=None,
        )
    except Exception as e:
        logger.exception("Upload job %s failed", job_id)
        _update_job(
            job_id,
            status="failed",
            finished_at=time.time(),
            error=str(e),
        )
        if output_path is not None and output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass
    finally:
        if ppt_path is not None:
            try:
                ppt_path.unlink(missing_ok=True)
            except OSError:
                pass


def _gather_vlps_worker(
    job_id: str,
    year: str,
    quarter: str,
    scope: dict,
    leagues: list[str],
    file_types: list[str],
    catalog: CatalogSource,
) -> None:
    """Worker para el modo VLPS (G.1).

    Diferencias con los otros workers:
    - No usa PPT base (no hay).
    - No tiene multi_team_mode (no aplica — output es siempre ZIP).
    - Output: SIEMPRE .zip con los archivos VLPS encontrados.

    Reusa _JOBS, _update_job y el flujo de status/download tal cual.
    """
    output_path: Path | None = None
    try:
        _update_job(job_id, status="running", started_at=time.time())

        def progress_cb(**updates: Any) -> None:
            _update_job(job_id, **updates)

        output, stream = gather_vlps_files(
            catalog=catalog,
            year=year,
            quarter=quarter,
            scope=scope,
            leagues=leagues,
            file_types=file_types,
            progress_callback=progress_cb,
        )

        # Output siempre .zip (gather_vlps_files lo setea con is_zip=True).
        fd, output_path_str = tempfile.mkstemp(suffix=".zip")
        output_path = Path(output_path_str)
        with os.fdopen(fd, "wb") as f:
            for chunk in stream:
                f.write(chunk)

        _update_job(
            job_id,
            status="done",
            finished_at=time.time(),
            output_path=output_path,
            output_filename=output.filename,
            output_media_type=output.media_type,
            current_team=None,
            current_phase=None,
        )
    except Exception as e:
        logger.exception("VLPS job %s failed", job_id)
        _update_job(
            job_id,
            status="failed",
            finished_at=time.time(),
            error=str(e),
        )
        if output_path is not None and output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass


def _reaper_loop() -> None:
    """Thread daemon que cada N segundos borra jobs done/failed expirados.

    El TTL es `settings.job_retention_seconds`. Si un job tiene
    `finished_at + TTL < now`, se elimina del registry y se borra su
    output_path del disco.

    Jobs todavía en estado "queued" o "running" NUNCA se borran acá
    (solo cuando terminan).
    """
    while True:
        s = get_settings()
        interval = max(s.job_reaper_interval_seconds, 10)
        time.sleep(interval)
        try:
            now = time.time()
            retention = s.job_retention_seconds
            expired_ids: list[str] = []
            expired_paths: list[Path] = []

            with _JOBS_LOCK:
                for jid, job in list(_JOBS.items()):
                    if job.finished_at is None:
                        continue  # aún en curso, no tocar
                    if (now - job.finished_at) > retention:
                        expired_ids.append(jid)
                        if job.output_path is not None:
                            expired_paths.append(job.output_path)
                for jid in expired_ids:
                    _JOBS.pop(jid, None)

            # Borrar los output files FUERA del lock — el unlink puede ser lento.
            for p in expired_paths:
                try:
                    p.unlink(missing_ok=True)
                except OSError as e:
                    logger.warning("Reaper: no pudo borrar %s: %s", p, e)

            if expired_ids:
                logger.info(
                    "Reaper: limpió %d job(s) expirados", len(expired_ids)
                )
        except Exception:
            # Nunca dejar morir al reaper — loggea y sigue.
            logger.exception("Reaper iteration falló — continuando")


# Arrancar el reaper como daemon thread al importar el módulo.
threading.Thread(target=_reaper_loop, daemon=True, name="job-reaper").start()


# ═══════════════════════════════════════════════════════════════
# Pydantic request models
# ═══════════════════════════════════════════════════════════════

class TeamSelectionIn(BaseModel):
    league: str
    team: str


class ScanCatalogRequest(BaseModel):
    year: str
    quarter: str
    selections: list[TeamSelectionIn]


# ─── Schemas del modo avanzado (B.1b) ────────────────────────────

class SeasonalCapsuleScopeIn(BaseModel):
    """Una cápsula seasonal seleccionada en el scope avanzado.

    `category`: string que el cliente recibió de /api/categories
        ("MENS", "KIDS/BOYS", etc.)
    `capsule_folder`: nombre crudo de la carpeta como vino de
        /api/capsules — el backend lo usa para identificar exactamente
        qué carpeta walkear.
    """
    category: str
    capsule_folder: str


class ClassicScopeIn(BaseModel):
    """Un classic_type + sub-productos seleccionados.

    Si `subproducts` está vacío, el classic NO se incluye en el scan
    (el frontend debería evitarlo, pero defensivamente el backend lo skipea).
    El sentinel `_DIRECTO` puede aparecer si hay merchboards al nivel raíz
    del classic_type.
    """
    category: str
    classic_type: str
    subproducts: list[str] = []


class AdvancedScopeIn(BaseModel):
    """Wrapper del scope completo. Una de las dos listas (o ambas) puede
    estar vacía, pero ambas vacías significa "scope vacío" → respuestas
    vacías de los endpoints scoped (no walkear nada).
    """
    seasonal: list[SeasonalCapsuleScopeIn] = []
    classics: list[ClassicScopeIn] = []


class LeaguesScopedRequest(BaseModel):
    year: str
    quarter: str
    scope: AdvancedScopeIn


class TeamsScopedRequest(BaseModel):
    year: str
    quarter: str
    league: str
    scope: AdvancedScopeIn


class ScanCatalogScopedRequest(BaseModel):
    year: str
    quarter: str
    selections: list[TeamSelectionIn]
    scope: AdvancedScopeIn


# ═══════════════════════════════════════════════════════════════
# Frontend
# ═══════════════════════════════════════════════════════════════

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    """Renderiza el index.html del frontend con el bearer token inyectado.

    El JS del frontend lee el token de un placeholder en el HTML y lo manda
    en cada fetch a /api/*. Además, seteamos un cookie httpOnly con el mismo
    token para que las descargas vía `<a href>` (que no pueden setear
    headers custom) también queden autenticadas. Ver `verify_token` para
    el path dual-mode header/cookie.
    """
    response = templates.TemplateResponse(
        request,
        "index.html",
        {"api_token": get_settings().api_token},
    )
    # Cookie de auth para downloads nativos del browser.
    # - httponly: el JS no puede leerlo (anti-XSS robo de token).
    # - samesite='strict': protege CSRF — solo se manda en requests
    #   originadas del mismo sitio. Aceptable para intranet single-domain.
    # - path='/api/': solo viaja a endpoints API, no a `/` ni assets.
    # - secure=False por default (dev/intranet HTTP). En prod HTTPS conviene
    #   pasarlo a True via env var; por ahora lo dejamos hardcoded en False
    #   para no romper localhost. Si se sirve por HTTPS, el browser sigue
    #   mandando el cookie sin problema porque secure=False es permissive.
    response.set_cookie(
        key=COOKIE_AUTH_NAME,
        value=get_settings().api_token,
        httponly=True,
        samesite="strict",
        path="/api/",
        secure=False,
    )
    return response


# ═══════════════════════════════════════════════════════════════
# Catalog listings
# ═══════════════════════════════════════════════════════════════

@app.get("/api/years", dependencies=[Depends(verify_token)])
def api_list_years(catalog: CatalogSource = Depends(get_catalog)):
    return {"years": catalog.list_years()}


@app.get("/api/quarters", dependencies=[Depends(verify_token)])
def api_list_quarters(
    year: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    return {"quarters": catalog.list_quarters(year)}


@app.get("/api/leagues", dependencies=[Depends(verify_token)])
def api_list_leagues(
    year: str,
    quarter: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    return {"leagues": catalog.list_leagues(year, quarter)}


@app.get("/api/teams", dependencies=[Depends(verify_token)])
def api_list_teams(
    year: str,
    quarter: str,
    league: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    return {"teams": catalog.list_teams(year, quarter, league)}


# ═══════════════════════════════════════════════════════════════
# Advanced listings (modo avanzado — B.1a)
# ═══════════════════════════════════════════════════════════════
#
# Exposición granular del árbol del catálogo. El modo simple (los
# endpoints de arriba) auto-descubre todo, mientras que el modo avanzado
# permite al usuario navegar categoría → año → quarter → cápsulas →
# classics → sub-productos → liga → equipos.

@app.get("/api/categories", dependencies=[Depends(verify_token)])
def api_list_categories(catalog: CatalogSource = Depends(get_catalog)):
    """Categorías top-level del catálogo, en formato string
    ("MENS", "KIDS/BOYS", etc.).
    """
    return {"categories": catalog.list_categories()}


@app.get("/api/capsules", dependencies=[Depends(verify_token)])
def api_list_capsules(
    year: str,
    quarter: str,
    category: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Cápsulas seasonal para una (year, quarter, category) específica.

    Devuelve `[{folder_name, capsule_key}, ...]`. El cliente muestra
    `capsule_key` al usuario y guarda el `folder_name` para mandar de
    vuelta en el scope.
    """
    return {"capsules": catalog.list_capsules(year, quarter, category)}


@app.get("/api/classics", dependencies=[Depends(verify_token)])
def api_list_classics(
    category: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Classic-types disponibles para una categoría (atemporal)."""
    return {"classics": catalog.list_classics(category)}


@app.get("/api/classic-subproducts", dependencies=[Depends(verify_token)])
def api_list_classic_subproducts(
    category: str,
    classic_type: str,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Sub-productos válidos dentro de un classic_type.

    Incluye `_DIRECTO` como primer item si el classic tiene merchboards
    directos sin sub-producto.
    """
    return {
        "subproducts": catalog.list_classic_subproducts(category, classic_type)
    }


# ═══════════════════════════════════════════════════════════════
# Scoped listings (modo avanzado — B.1b)
# ═══════════════════════════════════════════════════════════════
#
# POST en vez de GET porque el `scope` puede traer 10+ entries (cápsulas
# + classics + subproductos), demasiado para query string. Body JSON
# es más natural acá.

@app.post("/api/leagues-scoped", dependencies=[Depends(verify_token)])
def api_list_leagues_scoped(
    req: LeaguesScopedRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Devuelve las ligas que tienen contenido DENTRO del scope provisto.

    Walk scopeado — mucho más rápido que `/api/leagues` global cuando el
    scope incluye pocas cápsulas.
    """
    scope_dict = req.scope.model_dump()
    return {"leagues": catalog.list_leagues_scoped(req.year, req.quarter, scope_dict)}


@app.post("/api/teams-scoped", dependencies=[Depends(verify_token)])
def api_list_teams_scoped(
    req: TeamsScopedRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Devuelve los equipos de una liga DENTRO del scope provisto."""
    scope_dict = req.scope.model_dump()
    return {
        "teams": catalog.list_teams_scoped(
            req.year, req.quarter, scope_dict, req.league
        )
    }


class VlpsLeaguesRequest(BaseModel):
    """Payload para /api/vlps-leagues.

    Busca ligas en _PPTX/_PDF VLPS del scope — distinto de /api/leagues-scoped
    que busca en _MERCHBOARDS. Necesario porque en modo VLPS las ligas están
    bajo el folder VLPS, no bajo _MERCHBOARDS.
    """
    year: str
    quarter: str
    scope: AdvancedScopeIn
    file_types: list[str] = ["ppt", "pdf"]


@app.post("/api/vlps-leagues", dependencies=[Depends(verify_token)])
def api_list_vlps_leagues(
    req: VlpsLeaguesRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Ligas disponibles en _PPTX/_PDF VLPS del scope.

    A diferencia de /api/leagues-scoped (que busca en _MERCHBOARDS),
    este endpoint busca en los folders VLPS — que es donde están las
    ligas cuando el usuario está en modo 'Descargar VLPS'.
    """
    scope_dict = req.scope.model_dump()
    return {
        "leagues": catalog.list_vlps_leagues_scoped(
            req.year, req.quarter, scope_dict, req.file_types or ["ppt", "pdf"]
        )
    }


@app.post("/api/scan-catalog-scoped", dependencies=[Depends(verify_token)])
def api_scan_catalog_scoped(
    req: ScanCatalogScopedRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Multi-team scan scopeado: similar a /api/scan-catalog, pero
    solo incluye las cápsulas/classics del scope (no todas las del catálogo).
    """
    if len(req.selections) > 25:  # mismo límite que el de servicios
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Máximo 25 selecciones por batch",
        )
    scope_dict = req.scope.model_dump()
    try:
        return [
            catalog.scan_team_scoped(
                req.year, req.quarter, scope_dict, s.league, s.team
            )
            for s in req.selections
        ]
    except ValueError as e:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(e))


# ═══════════════════════════════════════════════════════════════
# Scan endpoints
# ═══════════════════════════════════════════════════════════════

@app.post("/api/scan-ppt", dependencies=[Depends(verify_token)])
async def api_scan_ppt(file: UploadFile = File(...)):
    """Escanea un PPT y devuelve las notas de cápsula detectadas + warnings.

    El PPT NO se persiste; se procesa en memoria y se descarta. El frontend
    lo vuelve a subir en /api/generate cuando el usuario clickea Generar.
    """
    contents = await file.read()
    try:
        return scan_pptx(contents)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"No se pudo leer el PPT: {e}",
        )


@app.post("/api/scan-catalog", dependencies=[Depends(verify_token)])
def api_scan_catalog(
    req: ScanCatalogRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Multi-team scan: itera cada selection y devuelve la disponibilidad
    por género × cápsula para cada team.
    """
    selections = [TeamSelection(league=s.league, team=s.team) for s in req.selections]
    try:
        return scan_catalog(catalog, req.year, req.quarter, selections)
    except ValueError as e:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(e))


# ═══════════════════════════════════════════════════════════════
# Generate
# ═══════════════════════════════════════════════════════════════

@app.post("/api/generate", dependencies=[Depends(verify_token)], status_code=202)
async def api_generate(
    ppt: UploadFile = File(...),
    year: str = Form(...),
    quarter: str = Form(...),
    selections: str = Form(...),
    scope: str | None = Form(None),
    multi_team_mode: str = Form("per_team"),
    catalog: CatalogSource = Depends(get_catalog),
):
    """Inicia una generación asincrónica y devuelve un job_id.

    Cambio de contrato: este endpoint ya NO devuelve el deck directamente.
    Ahora:
      1. Valida el input + streamea el PPT a disco.
      2. Crea un JobStatus en _JOBS y spawnea un thread worker.
      3. Devuelve 202 con {job_id, status, status_url} para que el cliente
         poolee el progreso.
      4. El cliente descarga vía GET /api/jobs/{id}/download cuando esté listo.

    `scope` (B.1d.2, opcional): JSON string con shape AdvancedScopeIn.
    Cuando viene, la generación usa `scan_team_scoped` para limitar las
    cápsulas/classics procesados al subset elegido por el usuario en el
    modo avanzado del frontend. Si no viene → path histórico (catálogo
    completo).

    Returns:
        202 + JSON con job_id, status="queued", status_url.

    Raises:
        400 si el shape del input es inválido.
    """
    try:
        sel_list = json.loads(selections)
    except json.JSONDecodeError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="selections debe ser un JSON válido",
        )

    try:
        parsed = [
            TeamSelection(league=s["league"], team=s["team"])
            for s in sel_list
        ]
    except (KeyError, TypeError) as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"selections con shape inválido: {e}",
        )

    if not parsed:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="selections no puede estar vacío",
        )

    # ── Validación de multi_team_mode (D.5a) ──
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"multi_team_mode inválido: {multi_team_mode!r} "
                f"(esperaba 'strict', 'mixed' o 'per_team')"
            ),
        )

    # ── Parse y validación opcional del scope (modo avanzado, B.1d.2) ──
    # El frontend en modo simple NO manda este campo → scope_dict = None →
    # generate_batch usa scan_team (path histórico, sin cambios).
    # En modo avanzado, viene un JSON string que validamos vía Pydantic
    # (AdvancedScopeIn) para garantizar el shape antes de mandarlo al worker.
    scope_dict: dict | None = None
    if scope is not None:
        try:
            scope_raw = json.loads(scope)
        except json.JSONDecodeError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="scope debe ser un JSON válido",
            )
        try:
            scope_dict = AdvancedScopeIn(**scope_raw).model_dump()
        except (TypeError, ValidationError) as e:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"scope con shape inválido: {e}",
            )

    # ── Stream del upload a un tempfile en disco ──
    # Para PPTs de 1-2GB, leer todo a RAM con `ppt.read()` rompe el server.
    # Copiamos chunk-by-chunk a disco; el worker thread lo va a leer desde ahí.
    # El cleanup del tempfile lo hace el worker (try/finally) — el endpoint
    # solo devuelve un 202 y se desentiende.
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    ppt_path = Path(tmp.name)
    tmp.close()

    try:
        with ppt_path.open("wb") as out:
            while True:
                chunk = await ppt.read(1024 * 1024)
                if not chunk:
                    break
                out.write(chunk)
    except Exception:
        _safe_unlink(ppt_path)
        raise

    ppt_base_name = Path(ppt.filename or "deck").stem

    # ── Crear el JobStatus y arrancar el worker ──
    job_id = uuid.uuid4().hex
    job = JobStatus(
        job_id=job_id,
        status="queued",
        teams_total=len(parsed),
    )
    with _JOBS_LOCK:
        _JOBS[job_id] = job

    worker = threading.Thread(
        target=_generate_worker,
        args=(job_id, ppt_path, ppt_base_name, year, quarter, parsed, catalog),
        kwargs={"scope": scope_dict, "multi_team_mode": multi_team_mode},
        daemon=True,
        name=f"job-{job_id[:8]}",
    )
    worker.start()

    return {
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/api/jobs/{job_id}/status",
        "teams_total": len(parsed),
    }


# ═══════════════════════════════════════════════════════════════
# Generate BLANK ("PPT nuevo desde cero" — B.2c)
# ═══════════════════════════════════════════════════════════════

@app.post(
    "/api/generate-blank",
    dependencies=[Depends(verify_token)],
    status_code=202,
)
async def api_generate_blank(
    year: str = Form(...),
    quarter: str = Form(...),
    selections: str = Form(...),
    scope: str = Form(...),
    multi_team_mode: str = Form("mixed"),
    catalog: CatalogSource = Depends(get_catalog),
):
    """Inicia una generación "PPT nuevo desde cero" — un deck construido
    de cero usando solo el scope avanzado, sin PPT base.

    Diferencias con `/api/generate`:
    - NO recibe upload de PPT (no hay).
    - `scope` es OBLIGATORIO (sin él no hay forma de saber qué cápsulas
      incluir; no hay notas en un PPT base que las dicten).
    - Multi-team produce UN solo .pptx mixto, no un ZIP — el contrato
      del response (job_id, status_url, descarga via /api/jobs/{id}/download)
      es idéntico al de /api/generate, así que el frontend no necesita
      lógica especial post-POST.

    Returns:
        202 + JSON con job_id, status="queued", status_url.

    Raises:
        400 si selections está vacío/malformado, o si scope falta/es
        vacío (scope.seasonal y scope.classics ambos []) o malformado.
    """
    # ── Validación de selections (igual que /api/generate) ──
    try:
        sel_list = json.loads(selections)
    except json.JSONDecodeError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="selections debe ser un JSON válido",
        )

    try:
        parsed = [
            TeamSelection(league=s["league"], team=s["team"])
            for s in sel_list
        ]
    except (KeyError, TypeError) as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"selections con shape inválido: {e}",
        )

    if not parsed:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="selections no puede estar vacío",
        )

    # ── Validación de multi_team_mode (D.5a) ──
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"multi_team_mode inválido: {multi_team_mode!r} "
                f"(esperaba 'strict', 'mixed' o 'per_team')"
            ),
        )

    # ── Validación del scope (obligatorio + no vacío) ──
    # En modo blank no tiene sentido un scope vacío: el deck saldría con
    # 0 slides. Rechazamos a nivel API para feedback inmediato al cliente.
    try:
        scope_raw = json.loads(scope)
    except json.JSONDecodeError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="scope debe ser un JSON válido",
        )
    try:
        scope_dict = AdvancedScopeIn(**scope_raw).model_dump()
    except (TypeError, ValidationError) as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"scope con shape inválido: {e}",
        )
    if not scope_dict.get("seasonal") and not scope_dict.get("classics"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                "scope debe tener al menos una cápsula seasonal o un "
                "classic — en modo 'PPT nuevo' un scope vacío produciría "
                "un deck sin slides"
            ),
        )

    # ── Naming del output ──
    # Convención: "Pro Standard — Nuevo {year} {quarter}" como ppt_base_name.
    # `generate_blank_batch` agrega "— {team}" para single y "— MULTI" para
    # multi-team. Resultado:
    #   Single: "Pro Standard — Nuevo 2027 Q1 — DALLAS COWBOYS.pptx"
    #   Multi : "Pro Standard — Nuevo 2027 Q1 — MULTI.pptx"
    ppt_base_name = f"Pro Standard — Nuevo {year} {quarter}"

    # ── Crear JobStatus + spawn worker ──
    job_id = uuid.uuid4().hex
    job = JobStatus(
        job_id=job_id,
        status="queued",
        teams_total=len(parsed),
    )
    with _JOBS_LOCK:
        _JOBS[job_id] = job

    worker = threading.Thread(
        target=_generate_blank_worker,
        args=(job_id, ppt_base_name, year, quarter, parsed, catalog),
        kwargs={"scope": scope_dict, "multi_team_mode": multi_team_mode},
        daemon=True,
        name=f"blank-job-{job_id[:8]}",
    )
    worker.start()

    return {
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/api/jobs/{job_id}/status",
        "teams_total": len(parsed),
    }


# ═══════════════════════════════════════════════════════════════
# Generate FROM UPLOADS (modo "manual" — D.1)
# ═══════════════════════════════════════════════════════════════

@app.post(
    "/api/generate-from-uploads",
    dependencies=[Depends(verify_token)],
    status_code=202,
)
async def api_generate_from_uploads(
    images: list[UploadFile] = File(...),
    output_mode: str = Form(...),
    ppt: UploadFile | None = File(None),
    multi_team_mode: str = Form("strict"),
):
    """Inicia una generación usando imágenes subidas manualmente (no del catálogo).

    Espera filenames con formato `LIGA_EQUIPO_CAPSULA[_GENERO][_NNN].ext`
    (ej: `NFL_DALLAS COWBOYS_FLAGSHIP_M_001.jpg`). El backend parsea cada
    filename para armar el merch_map ad-hoc — el catálogo SMB NO se toca.

    Args:
        images: lista de imágenes JPG/PNG (multipart, multiple).
        output_mode: 'existing' (reemplaza imágenes en PPT base) o
            'blank' (genera deck nuevo).
        ppt: PPT base (.pptx). Requerido si `output_mode='existing'`.
        multi_team_mode (D.4a, opcional): cómo manejar uploads de varios
            teams. 'strict' (default) = error si N>1. 'mixed' = un PPTX
            con slides interleaved. 'per_team' = ZIP con N decks.

    Returns:
        202 + {job_id, status, status_url, teams_total}.

    Raises:
        400: validación de input (sin imágenes, output_mode inválido,
            multi_team_mode inválido, existing sin PPT, demasiadas imágenes).
    """
    # ── Validación ──
    if output_mode not in ("existing", "blank"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"output_mode inválido: {output_mode!r} "
                f"(esperaba 'existing' o 'blank')"
            ),
        )
    if multi_team_mode not in ("strict", "mixed", "per_team"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"multi_team_mode inválido: {multi_team_mode!r} "
                f"(esperaba 'strict', 'mixed' o 'per_team')"
            ),
        )
    if not images:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Tenés que subir al menos una imagen",
        )
    if len(images) > MAX_MANUAL_UPLOADS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"Máximo {MAX_MANUAL_UPLOADS} imágenes por upload. "
                f"Recibidas: {len(images)}"
            ),
        )
    if output_mode == "existing" and ppt is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="output_mode='existing' requiere subir un PPT base",
        )

    # ── Leer imágenes a memoria ──
    # Single-team, esperamos < ~50 archivos × ~500KB cada uno = 25MB típico.
    # MAX_MANUAL_UPLOADS=200 × 5MB = ~1GB peor caso — sigue siendo OK en RAM
    # para un único request mientras el worker procesa, pero conviene
    # vigilar uso real en producción.
    images_data: list[tuple[str, bytes]] = []
    for img in images:
        try:
            data = await img.read()
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Error leyendo {img.filename}: {e}",
            )
        images_data.append((img.filename or "unknown", data))

    # ── PPT base (solo si output_mode='existing') → stream a tempfile ──
    ppt_path: Path | None = None
    ppt_base_name = "Pro Standard"
    if ppt is not None:
        tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
        ppt_path = Path(tmp.name)
        tmp.close()
        try:
            with ppt_path.open("wb") as out:
                while True:
                    chunk = await ppt.read(1024 * 1024)
                    if not chunk:
                        break
                    out.write(chunk)
        except Exception:
            _safe_unlink(ppt_path)
            raise
        ppt_base_name = Path(ppt.filename or "deck").stem
    elif output_mode == "blank":
        ppt_base_name = "Pro Standard — Nuevo"

    # ── Crear job + spawn worker ──
    # teams_total=1 ya que D.1 enforza single-team. Si el worker detecta
    # multi-team al parsear, falla y el job termina en "failed" — la UI
    # ve `error` y reporta al usuario.
    job_id = uuid.uuid4().hex
    job = JobStatus(
        job_id=job_id,
        status="queued",
        teams_total=1,
    )
    with _JOBS_LOCK:
        _JOBS[job_id] = job

    worker = threading.Thread(
        target=_generate_from_uploads_worker,
        args=(job_id, images_data, output_mode, ppt_path, ppt_base_name),
        kwargs={"multi_team_mode": multi_team_mode},
        daemon=True,
        name=f"upload-job-{job_id[:8]}",
    )
    worker.start()

    return {
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/api/jobs/{job_id}/status",
        "teams_total": 1,
    }


# ═══════════════════════════════════════════════════════════════
# Gather VLPS — descarga de archivos pre-armados (G.1)
# ═══════════════════════════════════════════════════════════════


class GatherVlpsRequest(BaseModel):
    """Payload del endpoint /api/gather-vlps.

    Reusa `AdvancedScopeIn` para el scope (mismo shape que /api/generate-blank).
    Sin selector de team — los archivos VLPS están agrupados por liga, no team.
    """
    year: str
    quarter: str
    leagues: list[str]                    # ej: ["NFL", "NBA"]
    scope: AdvancedScopeIn
    file_types: list[str] = ["ppt", "pdf"]  # default: ambos


@app.post(
    "/api/gather-vlps",
    dependencies=[Depends(verify_token)],
    status_code=202,
)
async def api_gather_vlps(
    req: GatherVlpsRequest,
    catalog: CatalogSource = Depends(get_catalog),
):
    """Inicia una descarga de archivos VLPS (PPTX/PDF pre-armados) del catálogo.

    A diferencia de los endpoints `/api/generate*`:
    - No procesa nada — solo busca, fetchea y empaqueta en ZIP.
    - No requiere PPT base ni teams — selección a nivel de LIGA.
    - Output siempre es ZIP, descargable vía /api/jobs/{id}/download.

    Args:
        year, quarter: período del catálogo.
        leagues: ligas a incluir. No puede estar vacío.
        scope: scope avanzado (cápsulas seasonal + classics, reusable).
        file_types: subset de ['ppt', 'pdf']. Default ambos.

    Returns:
        202 + {job_id, status, status_url, teams_total}.
        `teams_total` se setea como `len(leagues)` para que la UI muestre
        un progreso coherente (en VLPS no hay teams, leagues hacen ese rol).
    """
    # ── Validación ──
    if not req.leagues:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="leagues no puede estar vacío",
        )
    if not req.file_types or not all(ft in ("ppt", "pdf") for ft in req.file_types):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"file_types inválido: {req.file_types!r} "
                f"(esperaba subset no-vacío de ['ppt', 'pdf'])"
            ),
        )
    scope_dict = req.scope.model_dump()
    if not scope_dict.get("seasonal") and not scope_dict.get("classics"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                "scope debe tener al menos una cápsula seasonal o un "
                "classic — un scope vacío no encontraría archivos"
            ),
        )

    # ── Crear job + spawn worker ──
    job_id = uuid.uuid4().hex
    job = JobStatus(
        job_id=job_id,
        status="queued",
        teams_total=len(req.leagues),
    )
    with _JOBS_LOCK:
        _JOBS[job_id] = job

    worker = threading.Thread(
        target=_gather_vlps_worker,
        args=(
            job_id,
            req.year,
            req.quarter,
            scope_dict,
            list(req.leagues),
            list(req.file_types),
            catalog,
        ),
        daemon=True,
        name=f"vlps-job-{job_id[:8]}",
    )
    worker.start()

    return {
        "job_id": job_id,
        "status": "queued",
        "status_url": f"/api/jobs/{job_id}/status",
        "teams_total": len(req.leagues),
    }


# ═══════════════════════════════════════════════════════════════
# Job status + download
# ═══════════════════════════════════════════════════════════════

@app.get("/api/jobs/{job_id}/status", dependencies=[Depends(verify_token)])
def api_job_status(job_id: str):
    """Devuelve el estado actual de un job. Cliente polea cada N segundos
    hasta que `status` sea "done" o "failed", entonces usa `download_url`.
    """
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Job no encontrado (puede haber expirado o el server se reinició)",
            )
        return job.to_public_dict()


@app.get("/api/jobs/{job_id}/download", dependencies=[Depends(verify_token)])
def api_job_download(job_id: str):
    """Descarga el output de un job terminado.

    El cleanup post-descarga se hace solamente vía el reaper periódico
    (TTL `job_retention_seconds`, default 600s). Antes había un
    BackgroundTask que borraba el archivo apenas el FileResponse terminaba,
    pero eso rompía los retries: si la descarga del cliente fallaba (típico
    con archivos grandes — ver bug del download de 3GB), el archivo se
    perdía y el usuario tenía que regenerar todo. Ahora el archivo vive
    hasta el TTL del reaper, dándole margen al cliente para reintentar.
    """
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Job no encontrado",
            )
        if job.status == "failed":
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"Job falló: {job.error or 'desconocido'}",
            )
        if job.status != "done":
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"Job aún no listo (status={job.status!r}). Reintentá tras polear /status.",
            )
        output_path = job.output_path
        output_filename = job.output_filename or "deck.pptx"
        output_media_type = job.output_media_type or "application/octet-stream"

    if output_path is None or not output_path.exists():
        raise HTTPException(
            status_code=status.HTTP_410_GONE,
            detail="Output del job ya no está disponible (expiró o ya fue descargado)",
        )

    return FileResponse(
        path=str(output_path),
        media_type=output_media_type,
        headers={"Content-Disposition": _content_disposition(output_filename)},
    )


# ═══════════════════════════════════════════════════════════════
# Cache control
# ═══════════════════════════════════════════════════════════════

@app.post("/api/cache/invalidate", dependencies=[Depends(verify_token)])
def api_invalidate_cache(catalog: CatalogSource = Depends(get_catalog)):
    """Vacía el caché interno del CatalogSource.

    Útil cuando se sabe que el árbol del server cambió (cápsulas nuevas)
    y no se quiere esperar al TTL natural. No-op si el source no implementa
    `invalidate_cache` (ej. MockCatalogSource).
    """
    invalidator = getattr(catalog, "invalidate_cache", None)
    if callable(invalidator):
        invalidator()
        return {"status": "ok", "invalidated": True}
    return {"status": "ok", "invalidated": False, "reason": "source has no cache"}


# ═══════════════════════════════════════════════════════════════
# Helpers privados
# ═══════════════════════════════════════════════════════════════

def _content_disposition(filename: str) -> str:
    """Construye el header Content-Disposition con encoding RFC 5987 para
    soportar caracteres no-ASCII en el filename (ej. '—').
    """
    return f"attachment; filename*=UTF-8''{quote(filename)}"


def _safe_unlink(path: Path) -> None:
    """Borra el tempfile del upload. Tolera que ya no exista (race conditions
    o cleanups manuales). No re-raise — corre como BackgroundTask y los
    errores acá quedan loggeados pero no rompen el response al cliente.
    """
    try:
        path.unlink(missing_ok=True)
    except OSError as e:
        logger.warning("No se pudo borrar tempfile %s: %s", path, e)
