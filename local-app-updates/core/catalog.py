"""CatalogSource ABC + dataclasses (ImageRef, CapsuleAvailability, TeamScan)
+ implementaciones (SmbCatalogSource, MockCatalogSource) + factory.

Esta capa abstrae el origen del catálogo. Hoy SMB y Mock; mañana podría
ser S3, HTTP gateway, etc., sin que el resto del backend se entere.
"""

from __future__ import annotations

import io
import re
import threading
import zlib
from abc import ABC, abstractmethod
from collections.abc import Callable, Iterable, Iterator
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from operator import attrgetter
from pathlib import Path
from typing import TypeVar

from cachetools import TTLCache, cachedmethod
from cachetools.keys import hashkey
from PIL import Image, ImageDraw, ImageFont


# ═══════════════════════════════════════════════════════════════
# CACHE DEFAULTS
# ═══════════════════════════════════════════════════════════════

# Defaults usados si no se pasan al constructor de SmbCatalogSource.
# config.py los sobreescribe vía settings (catalog_cache_ttl_seconds,
# catalog_cache_maxsize).
CACHE_TTL_DEFAULT: int = 600       # 10 min
CACHE_MAXSIZE_DEFAULT: int = 2048  # entries por nivel
CACHE_WORKERS_DEFAULT: int = 16    # workers paralelos para walks SMB

# Para type-hinting el helper _parallel_map.
_T = TypeVar("_T")
_U = TypeVar("_U")


def _method_key_factory(method_name: str):
    """Devuelve una key function para `@cachedmethod` que prefija con el
    nombre del método, evitando colisiones cuando varios métodos comparten
    el mismo cache (`_method_cache`).

    Necesario porque el default `cachetools.keys.methodkey` calcula la
    clave SOLO con los argumentos (excluyendo `self`), entonces dos métodos
    con la misma firma (ej. `list_years()` y `list_categories()`, ambos
    sin args) generaban la misma key y se pisaban mutuamente.

    Esta factory excluye `self` igual que `methodkey` pero incluye el
    nombre del método como prefijo. La key resultante es:
        hashkey("<method_name>", *args, **kwargs)
    """
    def key_fn(self, *args, **kwargs):
        return hashkey(method_name, *args, **kwargs)
    return key_fn


# ═══════════════════════════════════════════════════════════════
# CONSTANTES DE FILESYSTEM (estructura del server real)
# ═══════════════════════════════════════════════════════════════

# Extensiones válidas para merchboards.
IMAGE_EXTS: frozenset[str] = frozenset({".jpg", ".jpeg", ".png", ".webp"})

# Códigos de género usados en notas del PPT y en filenames.
GENDER_CODES: frozenset[str] = frozenset({"M", "W", "K"})

# Mapeo carpeta de categoría → código de género del PPT.
# KIDS y sus sub-carpetas (BOYS/GIRLS/KIDS/PLAYER N&N) consolidan en "K"
# porque las notas del PPT solo distinguen M/W/K.
CATEGORY_TO_GENDER: dict[str, str] = {
    "MENS": "M",
    "WOMENS": "W",
    "KIDS": "K",
    "BOYS": "K",
    "GIRLS": "K",
    "PLAYER N&N": "K",
}

# Sub-carpetas dentro de KIDS/. Cada una se trata como una "categoría"
# para efectos de walkear el árbol, pero todas consolidan a gender=KIDS
# en el TeamScan resultante.
KIDS_SUBS: tuple[str, ...] = ("BOYS", "GIRLS", "KIDS", "PLAYER N&N")

# Prefijos que pueden aparecer al inicio del nombre de una carpeta
# de cápsula y deben strippearse para extraer solo el nombre.
# Ej: "01·01 - Q1 2027 MNS FLAGSHIP" → strip prefix "MNS" → "FLAGSHIP".
CAT_FOLDER_PREFIXES: dict[str, list[str]] = {
    "MENS":       ["MENS", "MNS"],
    "WOMENS":     ["WOMENS", "WMNS", "WNS"],
    "KIDS":       ["KIDS", "KDS"],
    "BOYS":       ["BOYS"],
    "GIRLS":      ["GIRLS"],
    "PLAYER N&N": ["PLAYER N&N", "N&N"],
}

# Marcador de "merchboards directo" en _CLASSIC (vs. ruteado por sub-producto).
# Ej: KIDS/BOYS/_CLASSIC/CLASSIC ICON/_MERCHBOARDS    ← directo
#     KIDS/BOYS/_CLASSIC/CLASSIC ICON/JERSEY/_MERCHBOARDS  ← via sub-producto
CLASSIC_DIRECT_KEY: str = "_DIRECTO"


# ─── VLPS: archivos PPTX/PDF pre-armados por cápsula (G.1) ──────────
# Estructura típica: {categoría}/{año}/{quarter}/{cápsula}/_PPTX VLPS/{liga}/file.pptx
# Variantes observadas:
# - Folder name: "_PPTX VLPS" o "PPTX VLPS" (con/sin underscore inicial).
# - "_PDF VLPS"  o "PDF VLPS"  para PDFs.
# - Archivos pueden estar flat directamente bajo el folder VLPS (sin sub-liga).
# - O bajo {liga}/ (con subcarpeta por liga, espejo parcial de _MERCHBOARDS).
VLPS_PPT_PATTERNS: tuple[str, ...] = ("_PPTX VLPS", "PPTX VLPS")
VLPS_PDF_PATTERNS: tuple[str, ...] = ("_PDF VLPS", "PDF VLPS")
VLPS_FILE_TYPES: tuple[str, ...] = ("ppt", "pdf")
# Extensiones aceptadas por tipo. .ppt incluido por compat con archivos legacy.
_VLPS_EXTENSIONS: dict[str, frozenset[str]] = {
    "ppt": frozenset({".pptx", ".ppt"}),
    "pdf": frozenset({".pdf"}),
}


# ═══════════════════════════════════════════════════════════════
# HELPERS — parsing de nombres de carpeta del server
# ═══════════════════════════════════════════════════════════════

def _strip_gender_suffix(key: str) -> str:
    """Si `key` termina en un código de género (M/W/K), devuelve la parte
    sin ese sufijo. Si no, devuelve `key` tal cual.

    "FLAGSHIP M" → "FLAGSHIP"
    "TEAM CITY"  → "TEAM CITY"
    """
    parts = key.split()
    if parts and parts[-1] in GENDER_CODES:
        return " ".join(parts[:-1])
    return key


def extract_capsule_from_folder(folder_name: str, category: str) -> str:
    """Extrae la clave de cápsula (capsule + gender) del nombre de carpeta
    seasonal del server.

    Ejemplos:
        "01·01 - Q1 2027 MNS FLAGSHIP" + "MENS"  →  "FLAGSHIP M"
        "01·02 - Q1 2027 WMNS FLORAL SPORT" + "WOMENS" → "FLORAL SPORT W"
        "01 - Q1 BOYS TEAM CITY" + "BOYS" → "TEAM CITY K"
    """
    # 1. Quitar el prefijo numérico hasta el primer "-" o "–".
    n = re.sub(r"^.+?[-–]\s*", "", folder_name).strip()
    # 2. Quitar el "Q1 2027 " (year opcional, quarter siempre).
    n = re.sub(r"^Q\d\s+\d{4}\s+", "", n, flags=re.IGNORECASE).strip()
    n = re.sub(r"^Q\d\s+", "", n, flags=re.IGNORECASE).strip()
    cat_upper = category.upper()

    # 3. Strippear el prefijo de categoría (versión larga o abreviada).
    for prefix in CAT_FOLDER_PREFIXES.get(cat_upper, [cat_upper]):
        if n.upper().startswith(prefix + " "):
            n = n[len(prefix):].strip()
            break
        if n.upper() == prefix:
            n = ""
            break

    # 4. Resolver el gender desde la categoría.
    gender = CATEGORY_TO_GENDER.get(cat_upper)

    # 5. Si todavía no tenemos gender, intentar leerlo del inicio del string
    #    (caso "COLLEGE MENS TEAM CITY" donde la categoría era "COLLEGE").
    if not gender:
        all_gender_prefixes = [
            (p, g)
            for prefixes, g in (
                (["MENS", "MNS"], "M"),
                (["WOMENS", "WMNS", "WNS"], "W"),
                (["KIDS", "KDS", "BOYS", "GIRLS", "PLAYER N&N", "N&N"], "K"),
            )
            for p in prefixes
        ]
        for prefix, g in all_gender_prefixes:
            if n.upper().startswith(prefix + " "):
                gender = g
                n = n[len(prefix):].strip()
                break
            if n.upper() == prefix:
                gender = g
                n = ""
                break

    key = n.upper()
    if gender and key:
        key = f"{key} {gender}"
    return key or folder_name.upper()


def extract_capsule_from_classic(folder_name: str, category: str) -> str:
    """Extrae la clave de cápsula del nombre de carpeta dentro de _CLASSIC.

    Ejemplos:
        "_CLASSIC ICON" + "MENS" → "CLASSIC ICON M"
        "HERITAGE HUSTLE" + "WOMENS" → "HERITAGE HUSTLE W"
        "PROPERTY OF" + "BOYS" → "PROPERTY OF K"
    """
    n = folder_name.lstrip("_").strip()
    gender = CATEGORY_TO_GENDER.get(category.upper())
    key = n.upper()
    if gender and key:
        key = f"{key} {gender}"
    return key or folder_name.upper()


# ═══════════════════════════════════════════════════════════════
# DATACLASSES
# ═══════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class ImageRef:
    """Identificador opaco de una imagen.

    `key` es el identificador interno del CatalogSource (path SMB, URL,
    S3 key, etc.). Solo el source que lo emitió sabe interpretarlo.
    `filename` es el nombre limpio que usan los parsers (parse_merch_name).
    """
    key: str
    filename: str


@dataclass(frozen=True)
class CapsuleAvailability:
    """Cuántas imágenes hay para una cápsula dentro de un género."""
    capsule: str
    image_count: int
    images: list[ImageRef]


@dataclass(frozen=True)
class VlpsFileRef:
    """Referencia a un archivo VLPS (PPTX o PDF pre-armado por cápsula).

    Análogo a `ImageRef` pero con metadata adicional para preview y
    para el armado del ZIP de salida:
    - `league`: '' si el archivo es FLAT (sin subfolder de liga adentro
      del VLPS dir — convención: "aplica a todas las ligas").
    - `category`: la categoría de origen para display (ej. "MENS", "COLLEGE").
    """
    key: str            # path absoluto en disco (fetchable via fetch_vlps_file)
    filename: str       # nombre del archivo tal cual está en el server
    file_type: str      # 'ppt' o 'pdf'
    capsule: str        # nombre limpio de cápsula (sin gender suffix)
    league: str         # 'NFL', 'NBA', ..., o '' si es flat
    category: str       # categoría del scope que originó este archivo


@dataclass(frozen=True)
class TeamScan:
    """Disponibilidad de imágenes para un team en un período dado.

    by_gender mapea género → lista de cápsulas con su availability.
    Ej: {"MENS": [CapsuleAvailability(...), ...], "WOMENS": [...], "KIDS": [...]}
    """
    team: str
    league: str
    by_gender: dict[str, list[CapsuleAvailability]]


# ═══════════════════════════════════════════════════════════════
# ABSTRACT BASE
# ═══════════════════════════════════════════════════════════════

class CatalogSource(ABC):
    """Interfaz que toda fuente de catálogo debe implementar.

    Las implementaciones concretas viven en este módulo (SmbCatalogSource,
    MockCatalogSource) o se pueden agregar nuevas (S3, HTTP gateway, etc.)
    sin tocar el resto del backend.
    """

    @abstractmethod
    def list_years(self) -> list[str]:
        """Años disponibles en el catálogo."""

    @abstractmethod
    def list_quarters(self, year: str) -> list[str]:
        """Cuartos disponibles para un año."""

    @abstractmethod
    def list_leagues(self, year: str, quarter: str) -> list[str]:
        """Unión de ligas presentes en MENS, WOMENS y KIDS para un período."""

    @abstractmethod
    def list_teams(self, year: str, quarter: str, league: str) -> list[str]:
        """Unión de teams presentes para una liga en cualquier género."""

    @abstractmethod
    def scan_team(self, year: str, quarter: str, league: str, team: str) -> TeamScan:
        """Disponibilidad por género × cápsula para un team."""

    @abstractmethod
    def fetch_image(self, ref: ImageRef) -> bytes:
        """Lee los bytes de una imagen dada su ref."""

    # ─── Listings granulares (modo avanzado) ─────────────────────
    # Estos métodos exponen el árbol del catálogo a un nivel más fino
    # que las APIs principales (list_leagues/list_teams/scan_team), que
    # auto-descubren todo. Acá el caller scopea explícitamente.

    @abstractmethod
    def list_categories(self) -> list[str]:
        """Lista las categorías top-level del catálogo.

        Formato: para categorías simples, el nombre directo ("MENS",
        "WOMENS"). Para KIDS con sub-folders: "KIDS/BOYS", "KIDS/GIRLS",
        etc. Solo devuelve las que EXISTEN en el filesystem.
        """

    @abstractmethod
    def list_capsules(self, year: str, quarter: str, category: str) -> list[dict]:
        """Cápsulas seasonal disponibles para esa (year, quarter, category).

        Cada item es {"folder_name": <str>, "capsule_key": <str>} donde:
        - folder_name: el nombre real de la carpeta en disco
          (ej: "01·01 - Q1 2027 MNS FLAGSHIP").
        - capsule_key: la clave parseada con gender al final
          (ej: "FLAGSHIP M").

        El frontend muestra el capsule_key al usuario, pero pasa de vuelta
        el folder_name en el scope (para que el backend sepa exactamente
        qué carpeta walkear).
        """

    @abstractmethod
    def list_classics(self, category: str) -> list[str]:
        """Lista los classic_types disponibles para una categoría.

        Independiente de year/quarter — los classics no varían con el
        período. Devuelve nombres de carpetas (ej. "_CLASSIC ICON").
        """

    @abstractmethod
    def list_classic_subproducts(self, category: str, classic_type: str) -> list[str]:
        """Sub-carpetas dentro de un classic_type.

        Incluye el sentinel `CLASSIC_DIRECT_KEY` ("_DIRECTO") como primer
        item si el classic tiene merchboards al primer nivel (sin sub-prod).
        Los demás items son nombres reales de sub-producto.
        """

    # ─── Listings scoped (modo avanzado — B.1b) ──────────────────
    # Reciben un `scope` que ya enumera explícitamente qué cápsulas
    # seasonal y qué classic-subproductos se quieren incluir. Esto
    # permite walks mucho más chicos que los métodos globales.
    #
    # Estructura del scope (dict):
    #   {
    #     "seasonal": [
    #       {"category": "MENS", "capsule_folder": "01 - Q1 2027 MNS FLAGSHIP"},
    #       ...
    #     ],
    #     "classics": [
    #       {"category": "MENS", "classic_type": "_CLASSIC ICON",
    #        "subproducts": ["_DIRECTO", "JERSEY"]},
    #       ...
    #     ]
    #   }

    @abstractmethod
    def list_leagues_scoped(self, year: str, quarter: str, scope: dict) -> list[str]:
        """Devuelve las ligas que tienen contenido EXCLUSIVAMENTE dentro
        del scope provisto (capsulas seasonal + classics seleccionados).

        Si el scope está vacío, devuelve `[]` (no walkea nada).
        """

    @abstractmethod
    def list_teams_scoped(
        self, year: str, quarter: str, scope: dict, league: str
    ) -> list[str]:
        """Equipos para una liga, scopeados al subset de cápsulas/classics.

        Si el scope está vacío o la liga no tiene contenido en él, devuelve `[]`.
        """

    @abstractmethod
    def scan_team_scoped(
        self, year: str, quarter: str, scope: dict, league: str, team: str
    ) -> TeamScan:
        """Disponibilidad por género × cápsula para un team, scopeada.

        Solo incluye cápsulas/classics presentes en el scope. La estructura
        de salida es idéntica a `scan_team` para que el resto del pipeline
        (build_merch_map / generate_deck) no necesite cambios.
        """

    def fetch_images(self, refs: list[ImageRef]) -> list[bytes]:
        """Fetch batch de imágenes. Default seriado; subclases pueden
        overridear con paralelismo. Devuelve los bytes en el mismo
        orden que `refs`.

        Provee un punto de extensión sin obligar a las subclases a
        implementar paralelización si no la necesitan (ej. MockCatalogSource
        es ya in-memory, no gana nada paralelizando).
        """
        return [self.fetch_image(ref) for ref in refs]

    # ─── VLPS: archivos pre-armados por cápsula (G.1) ─────────────
    # Path análogo a merchboards pero apunta a otro subfolder de la cápsula
    # (`_PPTX VLPS` o `_PDF VLPS`). El "archivo" es .pptx/.pdf, no .jpg.
    # Sin nivel de team — los archivos están bajo {liga}/ o flat.

    @abstractmethod
    def scan_vlps_files(
        self,
        year: str,
        quarter: str,
        scope: dict,
        leagues: list[str],
        file_types: list[str],
    ) -> list[VlpsFileRef]:
        """Escanea los _PPTX/_PDF VLPS folders del scope para las ligas dadas.

        Args:
            year, quarter, scope: igual semántica que scan_team_scoped.
            leagues: lista de códigos de liga (NFL, NBA, etc.). Archivos
                bajo `{vlps_dir}/{league}/` solo se incluyen si la liga
                está en este set. Archivos FLAT (directamente bajo el
                VLPS folder, sin subliga) SIEMPRE se incluyen.
            file_types: subset de `('ppt', 'pdf')`. Determina qué VLPS
                folders walkear.

        Returns:
            Lista plana de VlpsFileRef. Cada ref carga su metadata
            (capsule, league, file_type, category) para que el caller
            agrupe como prefiera.
        """

    @abstractmethod
    def list_vlps_leagues_scoped(
        self,
        year: str,
        quarter: str,
        scope: dict,
        file_types: list[str],
    ) -> list[str]:
        """Ligas con archivos VLPS en el scope.

        Busca en `_PPTX VLPS`/`_PDF VLPS` — NO en `_MERCHBOARDS`.
        Esto resuelve el bug donde el dropdown de ligas en modo VLPS
        quedaba vacío porque `list_leagues_scoped` solo mira _MERCHBOARDS.
        """

    def fetch_vlps_file(self, ref: VlpsFileRef) -> bytes:
        """Lee los bytes de un archivo VLPS. Default usa Path.read_bytes()."""
        return Path(ref.key).read_bytes()

    def fetch_vlps_files(self, refs: list[VlpsFileRef]) -> list[bytes]:
        """Fetch batch. Default seriado; subclases pueden overridear con
        paralelismo. Devuelve los bytes en el mismo orden que `refs`."""
        return [self.fetch_vlps_file(ref) for ref in refs]


# ═══════════════════════════════════════════════════════════════
# SMB IMPLEMENTATION
# ═══════════════════════════════════════════════════════════════

class SmbCatalogSource(CatalogSource):
    """Catálogo leyendo directamente de un share SMB vía paths UNC
    (o cualquier filesystem montado).

    Estructura esperada bajo `base_path`:

        base/
          MENS/
            2027/
              Q1 2027/               (o "Q1")
                01·01 - Q1 2027 MNS FLAGSHIP/
                  _MERCHBOARDS/
                    NFL/
                      DALLAS COWBOYS/
                        IMG_001.jpg
            _CLASSIC/
              CLASSIC ICON/
                _MERCHBOARDS/         ← directo
                  NFL/...
                JERSEY/                ← o vía sub-producto
                  _MERCHBOARDS/
                    NFL/...
          WOMENS/
            (igual que MENS)
          KIDS/
            BOYS/   (igual que MENS, consolidan a gender=KIDS)
            GIRLS/
            KIDS/
            PLAYER N&N/

    Implementación read-only sobre filesystem local (vía pathlib). Si el
    backend corre en Linux montando el share CIFS, o en Windows accediendo
    al UNC directo, da igual: pathlib lo maneja transparentemente.
    """

    # Géneros expuestos al resto del backend (consumidores de TeamScan).
    GENDERS: tuple[str, ...] = ("MENS", "WOMENS", "KIDS")

    def __init__(
        self,
        base_path: str,
        cache_ttl: int = CACHE_TTL_DEFAULT,
        cache_maxsize: int = CACHE_MAXSIZE_DEFAULT,
        walk_workers: int = CACHE_WORKERS_DEFAULT,
    ):
        """
        Args:
            base_path: Ruta base UNC (o filesystem montado).
            cache_ttl: Segundos antes de que cada entry expire. 0 = caché
                efectivamente desactivado (entries expiran al instante).
            cache_maxsize: Cantidad máxima de entries por nivel de caché.
            walk_workers: Workers paralelos del ThreadPoolExecutor para los
                walks pesados (list_leagues/list_teams/scan_team).
                1 = walks seriales (debug/tests). >=2 paraleliza N tasks.

        Dos cachés instanciadas:
        - `_iter_cache`: para `_safe_iterdir(path)`. Es la primitiva más
          llamada — todos los métodos públicos terminan acá. Cachear
          aquí da el mayor multiplicador.
        - `_method_cache`: para los métodos públicos (list_* y scan_team).
          Cuando el usuario repite la misma combinación de args, devuelve
          el resultado sin walkear.

        Ambas comparten un `RLock` para serializar accesos concurrentes
        desde threads de FastAPI. Lock no contenido: ~50ns de overhead.

        El ThreadPoolExecutor se crea lazy en el primer uso (evita crear
        threads en tests que solo instancian y no walkean).
        """
        self.base = Path(base_path)
        # Lock compartido para ambos cachés. Un solo lock simplifica la
        # razonabilidad (no hay riesgo de deadlock cruzado) y el overhead
        # es despreciable porque los métodos no se llaman bajo contención.
        self._cache_lock = threading.RLock()
        self._iter_cache: TTLCache = TTLCache(maxsize=cache_maxsize, ttl=max(cache_ttl, 0))
        self._method_cache: TTLCache = TTLCache(maxsize=cache_maxsize, ttl=max(cache_ttl, 0))
        # Paralelización de walks. Mínimo 1 (seriado). Executor es lazy.
        self._walk_workers: int = max(walk_workers, 1)
        self._executor: ThreadPoolExecutor | None = None
        self._executor_lock = threading.Lock()  # protege la creación lazy

    # ─── paralelización ──────────────────────────────────────────

    def _get_executor(self) -> ThreadPoolExecutor:
        """Devuelve el ThreadPoolExecutor (creándolo en el primer uso).

        Doble-checked locking: lectura sin lock en el camino caliente,
        creación una sola vez bajo el lock.
        """
        if self._executor is None:
            with self._executor_lock:
                if self._executor is None:
                    self._executor = ThreadPoolExecutor(
                        max_workers=self._walk_workers,
                        thread_name_prefix="smb-walk",
                    )
        return self._executor

    def _parallel_map(
        self,
        fn: "Callable[[_T], _U]",
        items: "Iterable[_T]",
    ) -> list["_U"]:
        """Aplica `fn` a cada `item` en paralelo y devuelve la lista de
        resultados en el mismo orden que `items`.

        - Si `walk_workers == 1`, ejecuta seriado sin tocar el executor
          (sin overhead de threads — útil en tests).
        - Si la lista de items es vacía, devuelve `[]` sin spawn de threads.
        - El orden de salida es el orden de entrada — clave para merges
          determinísticos en `scan_team`.
        """
        items_list = list(items)
        if not items_list:
            return []
        if self._walk_workers == 1:
            return [fn(item) for item in items_list]
        return list(self._get_executor().map(fn, items_list))

    def close(self) -> None:
        """Apaga el ThreadPoolExecutor si fue creado. Útil para tests
        que quieren limpieza explícita; en el singleton de prod no hace
        falta llamarlo (el process exit lo limpia).
        """
        with self._executor_lock:
            if self._executor is not None:
                self._executor.shutdown(wait=False)
                self._executor = None

    def invalidate_cache(self) -> None:
        """Vacía ambos cachés. Llamar tras saber que el árbol cambió
        (ej. el equipo de catálogos subió cápsulas nuevas) sin esperar al TTL.
        """
        with self._cache_lock:
            self._iter_cache.clear()
            self._method_cache.clear()

    # ─── helpers internos ────────────────────────────────────────

    def _category_roots(self) -> Iterator[tuple[Path, str, str]]:
        """Yield (path, label, gender) para cada raíz de categoría existente.

        - `path`: carpeta a iterar (ej. base/MENS, base/KIDS/BOYS).
        - `label`: nombre de la "categoría" para parsing de capsule
          (MENS/WOMENS/BOYS/GIRLS/KIDS/PLAYER N&N).
        - `gender`: el gender consolidado (MENS/WOMENS/KIDS).
        """
        for top in ("MENS", "WOMENS"):
            p = self.base / top
            if p.is_dir():
                yield p, top, top
        kids = self.base / "KIDS"
        if kids.is_dir():
            for sub in KIDS_SUBS:
                sp = kids / sub
                if sp.is_dir():
                    yield sp, sub, "KIDS"

    @cachedmethod(
        attrgetter("_iter_cache"),
        key=_method_key_factory("_safe_iterdir"),
        lock=attrgetter("_cache_lock"),
    )
    def _safe_iterdir(self, path: Path) -> list[Path]:
        """iterdir() pero tolerante a errores de filesystem, y cacheado.

        Devuelve [] si el path no existe o no se puede leer. Útil para
        no fallar el endpoint completo cuando una sub-carpeta del server
        está inaccesible.

        Cacheo: nivel 1, key = (path,). Es la operación SMB más cara y
        más repetida — todos los métodos públicos llegan acá.

        NOTA: el resultado se comparte por referencia entre callers. No
        mutar la lista devuelta; tratarla como inmutable.
        """
        try:
            return list(path.iterdir())
        except (OSError, PermissionError):
            return []

    def _find_quarter_dir(self, year_path: Path, quarter: str) -> Path | None:
        """Encuentra la carpeta del cuarto, manejando 'Q1' y 'Q1 2027'."""
        direct = year_path / quarter
        if direct.is_dir():
            return direct
        for d in self._safe_iterdir(year_path):
            if d.is_dir() and d.name.upper().startswith(quarter.upper()):
                return d
        return None

    def _find_merchboards_dir(self, capsule_path: Path) -> Path | None:
        """Busca la carpeta _MERCHBOARDS dentro de una cápsula.

        Acepta variantes como "MERCHBOARDS", "MERCHBOARD", etc.
        """
        direct = capsule_path / "_MERCHBOARDS"
        if direct.is_dir():
            return direct
        for d in self._safe_iterdir(capsule_path):
            if d.is_dir() and "MERCHBOARD" in d.name.upper():
                return d
        return None

    def _find_classic_merchboards(self, classic_type_path: Path) -> list[Path]:
        """Devuelve TODAS las _MERCHBOARDS dentro de un tipo de classic.

        Cubre los dos caminos:
        - `tipo/_MERCHBOARDS`                  (directo)
        - `tipo/sub_producto/_MERCHBOARDS`     (vía sub-producto)

        Se itera ambos para consolidar; el caller no necesita saber cuál
        es cuál — para "simple mode" todo va al mismo bag.
        """
        results: list[Path] = []
        direct = classic_type_path / "_MERCHBOARDS"
        if direct.is_dir():
            results.append(direct)
        for d in self._safe_iterdir(classic_type_path):
            if not d.is_dir() or "MERCHBOARD" in d.name.upper():
                continue
            sub_mb = d / "_MERCHBOARDS"
            if sub_mb.is_dir():
                results.append(sub_mb)
                continue
            # Fallback: alguna variante del nombre dentro del sub-producto.
            for sd in self._safe_iterdir(d):
                if sd.is_dir() and "MERCHBOARD" in sd.name.upper():
                    results.append(sd)
                    break
        return results

    def _find_classic_base_dir(self, cat_path: Path) -> Path | None:
        """Busca _CLASSIC o _CLASSICS bajo cat_path.
        HEADWEAR usa _CLASSICS (con S); otras categorías usan _CLASSIC.
        """
        for name in ("_CLASSIC", "_CLASSICS"):
            p = cat_path / name
            if p.is_dir():
                return p
        return None

    def _list_images(self, path: Path) -> list[Path]:
        """Lista archivos de imagen en `path` ordenados por nombre."""
        return sorted(
            (f for f in self._safe_iterdir(path)
             if f.is_file() and f.suffix.lower() in IMAGE_EXTS),
            key=lambda p: p.name,
        )

    def _enumerate_capsule_tasks(
        self, year: str, quarter: str,
    ) -> list[tuple[str, Path, str, str, Path]]:
        """Enumera (kind, cat_root, label, gender, path) para cada
        carpeta de cápsula (seasonal) y de tipo classic, en TODAS las
        categorías. Es la parte secuencial — solo ~30-50 ops superficiales.

        El resultado se pasa a `_parallel_map` para hacer el walk pesado
        en paralelo. Orden determinístico: categorías en el orden de
        `_category_roots`, capsulas/classics ordenadas por nombre.

        `kind` es "seasonal" o "classic".
        """
        tasks: list[tuple[str, Path, str, str, Path]] = []
        for cat_root, label, gender in self._category_roots():
            # Seasonal: cada carpeta de cápsula dentro del quarter dir.
            year_path = cat_root / year
            if year_path.is_dir():
                q_dir = self._find_quarter_dir(year_path, quarter)
                if q_dir is not None:
                    for cap in sorted(self._safe_iterdir(q_dir), key=lambda p: p.name):
                        if cap.is_dir():
                            tasks.append(("seasonal", cat_root, label, gender, cap))
            # Classic: cada carpeta de tipo dentro de _CLASSIC o _CLASSICS.
            classic = self._find_classic_base_dir(cat_root)
            if classic is not None:
                for ct in sorted(self._safe_iterdir(classic), key=lambda p: p.name):
                    if ct.is_dir():
                        tasks.append(("classic", cat_root, label, gender, ct))
        return tasks

    # ─── API pública ─────────────────────────────────────────────

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_years"),
        lock=attrgetter("_cache_lock"),
    )
    def list_years(self) -> list[str]:
        """Unión de años en TODAS las categorías top-level (incluyendo HEADWEAR, COLLEGE, etc.)."""
        years: set[str] = set()
        for cat_dir in self._safe_iterdir(self.base):
            if not cat_dir.is_dir() or cat_dir.name.startswith("_"):
                continue
            if cat_dir.name.upper() == "KIDS":
                for sub in self._safe_iterdir(cat_dir):
                    if not sub.is_dir() or sub.name.startswith("_"):
                        continue
                    for y in self._safe_iterdir(sub):
                        if (y.is_dir() and not y.name.startswith("_")
                                and y.name.isdigit() and len(y.name) == 4):
                            years.add(y.name)
            else:
                for y in self._safe_iterdir(cat_dir):
                    if (y.is_dir() and not y.name.startswith("_")
                            and y.name.isdigit() and len(y.name) == 4):
                        years.add(y.name)
        return sorted(years)

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_quarters"),
        lock=attrgetter("_cache_lock"),
    )
    def list_quarters(self, year: str) -> list[str]:
        """Unión de Q1/Q2/Q3/Q4 disponibles para el año en cualquier categoría."""
        quarters: set[str] = set()
        for cat_root, _label, _gender in self._category_roots():
            year_path = cat_root / year
            if not year_path.is_dir():
                continue
            for d in self._safe_iterdir(year_path):
                if not d.is_dir() or d.name.startswith("_"):
                    continue
                m = re.match(r"^Q([1-4])\b", d.name.strip().upper())
                if m:
                    quarters.add(f"Q{m.group(1)}")
        return sorted(quarters)

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_leagues"),
        lock=attrgetter("_cache_lock"),
    )
    def list_leagues(self, year: str, quarter: str) -> list[str]:
        """Unión de ligas encontradas tanto en seasonal como en _CLASSIC.

        Walkea: cat × (year/quarter/* + _CLASSIC/*) × _MERCHBOARDS/* → liga.

        Estrategia: enumera todas las (categoría × cápsula seasonal) +
        (categoría × classic_type) tasks de forma seriada (~30-50 ops),
        después corre el walk pesado de cada task en paralelo con
        `_parallel_map`. Para 100 tasks con 16 workers, ~6x speedup en
        cold start. Las llamadas subsiguientes pegan el caché.
        """
        tasks = self._enumerate_capsule_tasks(year, quarter)

        def scan_one(task: tuple[str, Path, str, str, Path]) -> set[str]:
            kind, _cat_root, _label, _gender, path = task
            found: set[str] = set()
            if kind == "seasonal":
                mb = self._find_merchboards_dir(path)
                if mb is not None:
                    for lg in self._safe_iterdir(mb):
                        if lg.is_dir():
                            found.add(lg.name)
            else:  # classic
                for mb in self._find_classic_merchboards(path):
                    for lg in self._safe_iterdir(mb):
                        if lg.is_dir():
                            found.add(lg.name)
            return found

        partial_sets = self._parallel_map(scan_one, tasks)
        all_leagues: set[str] = set().union(*partial_sets) if partial_sets else set()
        return sorted(all_leagues)

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_teams"),
        lock=attrgetter("_cache_lock"),
    )
    def list_teams(self, year: str, quarter: str, league: str) -> list[str]:
        """Unión de equipos para la liga, en seasonal + _CLASSIC.

        Mismo patrón paralelo que `list_leagues`: enumera tasks
        secuencialmente, walkea cada task en paralelo.
        """
        tasks = self._enumerate_capsule_tasks(year, quarter)

        def scan_one(task: tuple[str, Path, str, str, Path]) -> set[str]:
            kind, _cat_root, _label, _gender, path = task
            found: set[str] = set()
            if kind == "seasonal":
                mb = self._find_merchboards_dir(path)
                if mb is not None:
                    league_dir = mb / league
                    if league_dir.is_dir():
                        for t in self._safe_iterdir(league_dir):
                            if t.is_dir():
                                found.add(t.name)
            else:  # classic
                for mb in self._find_classic_merchboards(path):
                    league_dir = mb / league
                    if league_dir.is_dir():
                        for t in self._safe_iterdir(league_dir):
                            if t.is_dir():
                                found.add(t.name)
            return found

        partial_sets = self._parallel_map(scan_one, tasks)
        all_teams: set[str] = set().union(*partial_sets) if partial_sets else set()
        return sorted(all_teams)

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("scan_team"),
        lock=attrgetter("_cache_lock"),
    )
    def scan_team(self, year: str, quarter: str, league: str, team: str) -> TeamScan:
        """Disponibilidad por género × cápsula para un team.

        Consolida seasonal + _CLASSIC en `by_gender[gender]`. La key de
        `CapsuleAvailability.capsule` es el nombre limpio (sin código de
        gender al final), ya que el gender ya está implícito en la clave
        del dict `by_gender`.

        Estrategia paralela: cada task (capsula seasonal o classic_type)
        devuelve un dict local `{(gender, cap_name): [ImageRef, ...]}`.
        Las tasks se procesan en paralelo y luego se merge-an en orden
        determinístico (el orden de enumeración, que también es el orden
        de retorno garantizado por `_parallel_map`).
        """
        tasks = self._enumerate_capsule_tasks(year, quarter)

        def scan_one(
            task: tuple[str, Path, str, str, Path]
        ) -> dict[tuple[str, str], list[ImageRef]]:
            kind, _cat_root, label, gender, path = task
            local: dict[tuple[str, str], list[ImageRef]] = {}
            if kind == "seasonal":
                mb = self._find_merchboards_dir(path)
                if mb is None:
                    return local
                team_path = mb / league / team
                if not team_path.is_dir():
                    return local
                cap_full = extract_capsule_from_folder(path.name, label)
                cap_name = _strip_gender_suffix(cap_full)
                if not cap_name:
                    return local
                imgs = [
                    ImageRef(key=str(img), filename=img.name)
                    for img in self._list_images(team_path)
                ]
                if imgs:
                    local[(gender, cap_name)] = imgs
            else:  # classic
                cap_full = extract_capsule_from_classic(path.name, label)
                cap_name = _strip_gender_suffix(cap_full)
                if not cap_name:
                    return local
                for mb in self._find_classic_merchboards(path):
                    team_path = mb / league / team
                    if not team_path.is_dir():
                        continue
                    imgs = [
                        ImageRef(key=str(img), filename=img.name)
                        for img in self._list_images(team_path)
                    ]
                    if imgs:
                        local.setdefault((gender, cap_name), []).extend(imgs)
            return local

        partials = self._parallel_map(scan_one, tasks)

        # Merge en orden de tasks (determinístico — preservamos el orden
        # de las imágenes equivalente al modo seriado original).
        acc: dict[str, dict[str, list[ImageRef]]] = {g: {} for g in self.GENDERS}
        for partial in partials:
            for (gender, cap_name), imgs in partial.items():
                acc[gender].setdefault(cap_name, []).extend(imgs)

        # Convertir el acumulador a la estructura final, cápsulas ordenadas.
        by_gender: dict[str, list[CapsuleAvailability]] = {}
        for gender in self.GENDERS:
            caps: list[CapsuleAvailability] = []
            for cap_name in sorted(acc[gender].keys()):
                images = acc[gender][cap_name]
                caps.append(CapsuleAvailability(
                    capsule=cap_name,
                    image_count=len(images),
                    images=images,
                ))
            by_gender[gender] = caps

        return TeamScan(team=team, league=league, by_gender=by_gender)

    def fetch_image(self, ref: ImageRef) -> bytes:
        """Lee los bytes del archivo apuntado por `ref.key`.

        En SMB, `key` es el path absoluto del archivo (str).
        """
        return Path(ref.key).read_bytes()

    def fetch_images(self, refs: list[ImageRef]) -> list[bytes]:
        """Fetch batch en paralelo usando el mismo ThreadPoolExecutor
        de los walks. Cada fetch es ~1 lectura de archivo (I/O bound),
        ideal para paralelizar — el GIL se libera durante read_bytes.

        Devuelve la lista de bytes en el orden de `refs` (garantizado
        por `_parallel_map`).
        """
        return self._parallel_map(self.fetch_image, refs)

    # ─── Listings granulares (modo avanzado) ─────────────────────

    def _category_to_path(self, category: str) -> Path:
        """Traduce el string de categoría a path en disco.

        "MENS" → base/MENS
        "KIDS/BOYS" → base/KIDS/BOYS
        """
        # str(Path) ignora slashes seguros de Windows/Linux, usamos joinpath.
        parts = category.split("/")
        p = self.base
        for part in parts:
            p = p / part
        return p

    def _category_label(self, category: str) -> str:
        """Devuelve el "label" usado para parsing del nombre de cápsula.

        "MENS" → "MENS"
        "KIDS/BOYS" → "BOYS"  (el último segmento — es la sub-folder real
                              de donde sale el prefijo en el folder_name)
        """
        return category.rsplit("/", 1)[-1]

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_categories"),
        lock=attrgetter("_cache_lock"),
    )
    def list_categories(self) -> list[str]:
        """Descubre dinámicamente las categorías reales del filesystem.

        Lógica:
        - Lista todas las carpetas top-level bajo `base/` excepto las
          que empiezan con `_` (que son convenciones internas, ej
          `_CLASSIC` cuando aparece al primer nivel).
        - Caso especial KIDS: en lugar de devolver `"KIDS"` plano,
          expone cada sub-carpeta de KIDS como `"KIDS/<sub>"` (BOYS,
          GIRLS, KIDS, PLAYER N&N, etc.) porque ese es el nivel
          semánticamente "navegable" — bajo KIDS plano no hay años.
        - PLAYER N&N puede coexistir como top-level AND como sub de
          KIDS; ambos aparecen en la lista con paths distintos
          ("PLAYER N&N" vs "KIDS/PLAYER N&N").
        - Orden: alfabético estable. El frontend muestra eso al usuario.

        IMPORTANTE: este método es "liberal" — devuelve categorías
        aunque no tengan estructura `year/quarter/...` adentro (ej.
        `Tool Kits`, `Marketing Events`). Si el usuario elige una
        categoría sin esa estructura, `list_capsules()` devuelve `[]`
        silenciosamente y el frontend muestra el dropdown vacío.

        El modo SIMPLE (list_leagues/list_teams/scan_team) usa otro
        path — `_category_roots()` — que SÍ es conservador y solo
        walkea MENS/WOMENS/KIDS-subs. Esa decisión se mantiene para
        que el modo simple sea rápido y predecible.
        """
        cats: list[str] = []
        for d in sorted(self._safe_iterdir(self.base), key=lambda p: p.name):
            if not d.is_dir() or d.name.startswith("_"):
                continue
            if d.name.upper() == "KIDS":
                # Expandir las sub-carpetas de KIDS al formato "KIDS/<sub>".
                # Cada sub se trata como su propia categoría navegable.
                for sub in sorted(self._safe_iterdir(d), key=lambda p: p.name):
                    if sub.is_dir() and not sub.name.startswith("_"):
                        cats.append(f"KIDS/{sub.name}")
            else:
                # Cualquier otra categoría top-level se agrega tal cual.
                cats.append(d.name)
        return cats

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_capsules"),
        lock=attrgetter("_cache_lock"),
    )
    def list_capsules(self, year: str, quarter: str, category: str) -> list[dict]:
        """Cápsulas seasonal disponibles para esa categoría/year/quarter.

        Walk scopeado a UN solo path — mucho más rápido que el global
        `list_leagues`. Cada item: {"folder_name", "capsule_key"}.
        """
        cat_path = self._category_to_path(category)
        if not cat_path.is_dir():
            return []
        label = self._category_label(category)
        year_path = cat_path / year
        if not year_path.is_dir():
            return []
        q_dir = self._find_quarter_dir(year_path, quarter)
        if q_dir is not None:
            scan_dirs = [q_dir]
        else:
            # Sin Q1-Q4: categorías como HEADWEAR usan temporadas (FALL 2025·).
            # Buscar subdirs no-underscore del año; si tienen sub-carpetas propias
            # son temporadas y las cápsulas reales están dentro de ellas.
            season_dirs = [
                d for d in self._safe_iterdir(year_path)
                if d.is_dir() and not d.name.startswith("_")
            ]
            if season_dirs and any(
                any(s.is_dir() and not s.name.startswith("_")
                    for s in self._safe_iterdir(sd))
                for sd in season_dirs
            ):
                scan_dirs = season_dirs
            else:
                scan_dirs = [year_path]

        items: list[dict] = []
        seen: set[str] = set()
        for scan_dir in scan_dirs:
            for cap_folder in sorted(self._safe_iterdir(scan_dir), key=lambda p: p.name):
                if not cap_folder.is_dir() or cap_folder.name.startswith("_"):
                    continue
                if cap_folder.name in seen:
                    continue
                seen.add(cap_folder.name)
                cap_key = extract_capsule_from_folder(cap_folder.name, label)
                items.append({
                    "folder_name": cap_folder.name,
                    "capsule_key": cap_key,
                })
        return items

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_classics"),
        lock=attrgetter("_cache_lock"),
    )
    def list_classics(self, category: str) -> list[str]:
        """Nombres de classic_types disponibles bajo `<category>/_CLASSIC/`.

        Independiente de year/quarter — los classics son atemporales.
        """
        classic_dir = self._find_classic_base_dir(self._category_to_path(category))
        if classic_dir is None:
            return []
        return sorted(
            d.name for d in self._safe_iterdir(classic_dir) if d.is_dir()
        )

    @cachedmethod(
        attrgetter("_method_cache"),
        key=_method_key_factory("list_classic_subproducts"),
        lock=attrgetter("_cache_lock"),
    )
    def list_classic_subproducts(
        self, category: str, classic_type: str
    ) -> list[str]:
        """Sub-productos válidos dentro de un classic_type.

        Devuelve:
        - `CLASSIC_DIRECT_KEY` ("_DIRECTO") como PRIMER item si el classic
          tiene `_MERCHBOARDS` directo al primer nivel.
        - Los nombres de sub-carpetas que CONTIENEN un `_MERCHBOARDS`
          (filtrando las que no aportan).

        El frontend lo usa para que el usuario elija qué partes del classic
        incluir; el scope final lleva la lista de sub-productos elegidos.
        """
        classic_base = self._find_classic_base_dir(self._category_to_path(category))
        if classic_base is None:
            return []
        ct_path = classic_base / classic_type
        if not ct_path.is_dir():
            return []

        options: list[str] = []
        if (ct_path / "_MERCHBOARDS").is_dir():
            options.append(CLASSIC_DIRECT_KEY)

        for sub in sorted(self._safe_iterdir(ct_path), key=lambda p: p.name):
            if not sub.is_dir() or "MERCHBOARD" in sub.name.upper():
                continue
            if (sub / "_MERCHBOARDS").is_dir():
                options.append(sub.name)
                continue
            # Fallback: cualquier carpeta con un _MERCHBOARDS adentro.
            for sd in self._safe_iterdir(sub):
                if sd.is_dir() and "MERCHBOARD" in sd.name.upper():
                    options.append(sub.name)
                    break

        # Fallback HEADWEAR _CLASSICS: si no hay _MERCHBOARDS en ningún nivel
        # pero sí hay subdirs con archivos de presentación (liga/files), exponer
        # CLASSIC_DIRECT_KEY para que el scope pueda incluir el classic type entero.
        if not options:
            vlps_exts = frozenset({".pptx", ".ppt", ".pdf"})
            has_league_files = any(
                sub.is_dir() and not sub.name.startswith("_") and
                any(f.is_file() and f.suffix.lower() in vlps_exts
                    for f in self._safe_iterdir(sub))
                for sub in self._safe_iterdir(ct_path)
            )
            if has_league_files:
                options.append(CLASSIC_DIRECT_KEY)

        return options

    # ─── Helpers privados para los métodos scoped ─────────────────

    def _category_to_gender(self, category: str) -> str | None:
        """Mapea el string de categoría al gender bucket del TeamScan.

        "MENS" → "MENS", "WOMENS" → "WOMENS", "KIDS/anything" → "KIDS".
        Devuelve None si no reconoce la categoría (defensivo).
        """
        label = self._category_label(category)
        gender_code = CATEGORY_TO_GENDER.get(label.upper())
        if gender_code == "M":
            return "MENS"
        if gender_code == "W":
            return "WOMENS"
        if gender_code == "K":
            return "KIDS"
        return None

    def _resolve_seasonal_cap_path(
        self, year: str, quarter: str, category: str, capsule_folder: str
    ) -> Path | None:
        """Resuelve el directorio de la cápsula seasonal (sin entrar al
        `_MERCHBOARDS` ni al `_PPTX VLPS`). Es el punto desde donde
        seguir buscando subfolders (merchboards o VLPS, según el caller).

        Para categorías sin quarter (ej. HEADWEAR), si no se encuentra la
        carpeta de quarter, busca la cápsula directamente bajo el año.
        """
        cat_path = self._category_to_path(category)
        year_path = cat_path / year
        if not year_path.is_dir():
            return None
        q_dir = self._find_quarter_dir(year_path, quarter)
        if q_dir is not None:
            cap_path = q_dir / capsule_folder
            return cap_path if cap_path.is_dir() else None
        # Sin Q1-Q4: buscar la cápsula dentro de cada temporada (FALL 2025·, etc.).
        for season_dir in self._safe_iterdir(year_path):
            if not season_dir.is_dir() or season_dir.name.startswith("_"):
                continue
            cap_path = season_dir / capsule_folder
            if cap_path.is_dir():
                return cap_path
        return None

    def _resolve_seasonal_mb(
        self, year: str, quarter: str, category: str, capsule_folder: str
    ) -> Path | None:
        """Resuelve el `_MERCHBOARDS` dir para una cápsula seasonal en
        el scope. Devuelve None si algo del path no existe.
        """
        cap_path = self._resolve_seasonal_cap_path(
            year, quarter, category, capsule_folder,
        )
        if cap_path is None:
            return None
        return self._find_merchboards_dir(cap_path)

    def _resolve_classic_mb_for_sub(
        self, category: str, classic_type: str, subproduct: str
    ) -> Path | None:
        """Resuelve el `_MERCHBOARDS` dir para un (classic_type, subproduct).

        Si `subproduct == CLASSIC_DIRECT_KEY`, busca `_MERCHBOARDS` directo
        bajo el classic_type. Si no, busca dentro de la sub-carpeta.
        """
        classic_base = self._find_classic_base_dir(self._category_to_path(category))
        if classic_base is None:
            return None
        ct_path = classic_base / classic_type
        if not ct_path.is_dir():
            return None

        if subproduct == CLASSIC_DIRECT_KEY:
            mb = ct_path / "_MERCHBOARDS"
            return mb if mb.is_dir() else None

        sub_path = ct_path / subproduct
        if not sub_path.is_dir():
            return None
        direct = sub_path / "_MERCHBOARDS"
        if direct.is_dir():
            return direct
        # Fallback: variantes del nombre
        for sd in self._safe_iterdir(sub_path):
            if sd.is_dir() and "MERCHBOARD" in sd.name.upper():
                return sd
        return None

    def _iter_scope_merchboards(
        self, year: str, quarter: str, scope: dict
    ) -> Iterator[tuple[Path, str, str, str]]:
        """Itera todos los `_MERCHBOARDS` dirs incluidos en el scope.

        Yield: (mb_path, category, gender, capsule_name_clean) por cada
        entry del scope que tenga un `_MERCHBOARDS` válido en disco.
        - `gender`: "MENS" | "WOMENS" | "KIDS".
        - `capsule_name_clean`: el nombre de la cápsula SIN sufijo de gender,
          listo para usar como key en `TeamScan.by_gender[g]`.
        """
        # Seasonal
        for sc in scope.get("seasonal", []) or []:
            cat = sc.get("category")
            folder = sc.get("capsule_folder")
            if not cat or not folder:
                continue
            mb = self._resolve_seasonal_mb(year, quarter, cat, folder)
            if mb is None:
                continue
            gender = self._category_to_gender(cat)
            if gender is None:
                continue
            label = self._category_label(cat)
            cap_full = extract_capsule_from_folder(folder, label)
            cap_name = _strip_gender_suffix(cap_full)
            if not cap_name:
                continue
            yield mb, cat, gender, cap_name

        # Classics
        for cs in scope.get("classics", []) or []:
            cat = cs.get("category")
            classic_type = cs.get("classic_type")
            subs = cs.get("subproducts") or []
            if not cat or not classic_type or not subs:
                continue
            gender = self._category_to_gender(cat)
            if gender is None:
                continue
            label = self._category_label(cat)
            cap_full = extract_capsule_from_classic(classic_type, label)
            cap_name = _strip_gender_suffix(cap_full)
            if not cap_name:
                continue
            for sub in subs:
                mb = self._resolve_classic_mb_for_sub(cat, classic_type, sub)
                if mb is None:
                    continue
                yield mb, cat, gender, cap_name

    # ─── API pública scoped ───────────────────────────────────────

    def list_leagues_scoped(self, year: str, quarter: str, scope: dict) -> list[str]:
        leagues: set[str] = set()
        for mb, _cat, _gender, _cap in self._iter_scope_merchboards(year, quarter, scope):
            for lg in self._safe_iterdir(mb):
                if lg.is_dir():
                    leagues.add(lg.name)
        return sorted(leagues)

    def list_teams_scoped(
        self, year: str, quarter: str, scope: dict, league: str
    ) -> list[str]:
        teams: set[str] = set()
        for mb, _cat, _gender, _cap in self._iter_scope_merchboards(year, quarter, scope):
            league_dir = mb / league
            if league_dir.is_dir():
                for t in self._safe_iterdir(league_dir):
                    if t.is_dir():
                        teams.add(t.name)
        return sorted(teams)

    # ─── VLPS implementation (G.1) ────────────────────────────────

    def _find_vlps_dir(self, cap_path: Path, file_type: str) -> Path | None:
        """Encuentra el folder VLPS de un tipo (ppt/pdf) dentro de la
        cápsula. Maneja las variantes "_PPTX VLPS", "PPTX VLPS",
        "_PDF VLPS", "PDF VLPS".

        Estrategia:
        1. Probar los patrones literales en orden.
        2. Si nada matchea, búsqueda flexible: cualquier folder que
           contenga "VLPS" y el discriminador ("PPTX" o "PDF") en su
           nombre. Cubre variaciones tipográficas raras del server.
        """
        patterns = VLPS_PPT_PATTERNS if file_type == "ppt" else VLPS_PDF_PATTERNS
        for pattern in patterns:
            candidate = cap_path / pattern
            if candidate.is_dir():
                return candidate
        # Fallback flexible.
        discriminator = "PPTX" if file_type == "ppt" else "PDF"
        for d in self._safe_iterdir(cap_path):
            if not d.is_dir():
                continue
            name_upper = d.name.upper()
            if "VLPS" in name_upper and discriminator in name_upper:
                return d
        return None

    def _walk_vlps_files_in_folder(
        self,
        vlps_dir: Path,
        leagues: list[str],
        file_type: str,
        capsule: str,
        category: str,
    ) -> list[VlpsFileRef]:
        """Walkea adentro de un VLPS folder. Solo direct children:
        - Subfolder cuyo nombre matchea una liga seleccionada → walkea su contenido.
        - Archivo directo con extensión válida → flat (league='').

        Files más profundos (ej. `_PPTX VLPS/NFL/archive/old.pptx`) se ignoran.
        Si en el futuro aparece esa estructura, agregamos recursión opcional.
        """
        exts = _VLPS_EXTENSIONS.get(file_type, frozenset())
        # Normalize leagues para case-insensitive matching (defensive).
        league_set_upper = {l.upper() for l in leagues}
        out: list[VlpsFileRef] = []

        for entry in self._safe_iterdir(vlps_dir):
            if entry.is_dir():
                if entry.name.upper() in league_set_upper:
                    # Liga seleccionada — listar files adentro.
                    for f in self._safe_iterdir(entry):
                        if f.is_file() and f.suffix.lower() in exts:
                            out.append(VlpsFileRef(
                                key=str(f),
                                filename=f.name,
                                file_type=file_type,
                                capsule=capsule,
                                league=entry.name,
                                category=category,
                            ))
            elif entry.is_file() and entry.suffix.lower() in exts:
                # Flat file — sin subfolder de liga. Aplica a todas.
                out.append(VlpsFileRef(
                    key=str(entry),
                    filename=entry.name,
                    file_type=file_type,
                    capsule=capsule,
                    league="",
                    category=category,
                ))
        return out

    def _walk_vlps_flat_files(
        self,
        cap_path: Path,
        leagues: list[str],
        file_type: str,
        capsule: str,
        category: str,
    ) -> list[VlpsFileRef]:
        """Walkea archivos planos donde la liga aparece al final del nombre.
        Patrón: *_{LIGA}.ext — usado en HEADWEAR (ej. Q3_FA25-ON CAMPUS_MLB.pdf).
        """
        exts = _VLPS_EXTENSIONS.get(file_type, frozenset())
        league_set_upper = {l.upper() for l in leagues}
        out: list[VlpsFileRef] = []
        for f in self._safe_iterdir(cap_path):
            if not f.is_file() or f.suffix.lower() not in exts:
                continue
            parts = f.stem.rsplit("_", 1)
            if len(parts) < 2:
                continue
            league = parts[-1]
            if league.upper() in league_set_upper:
                out.append(VlpsFileRef(
                    key=str(f),
                    filename=f.name,
                    file_type=file_type,
                    capsule=capsule,
                    league=league,
                    category=category,
                ))
        return out

    def scan_vlps_files(
        self,
        year: str,
        quarter: str,
        scope: dict,
        leagues: list[str],
        file_types: list[str],
    ) -> list[VlpsFileRef]:
        # Validación defensiva — el endpoint ya valida, pero por las dudas.
        if not file_types or not all(ft in VLPS_FILE_TYPES for ft in file_types):
            raise ValueError(
                f"file_types inválido: {file_types!r} "
                f"(esperaba subset de {VLPS_FILE_TYPES})"
            )

        out: list[VlpsFileRef] = []

        # ── Seasonal ──
        for sc in scope.get("seasonal", []) or []:
            cat = sc.get("category")
            folder = sc.get("capsule_folder")
            if not cat or not folder:
                continue
            cap_path = self._resolve_seasonal_cap_path(year, quarter, cat, folder)
            if cap_path is None:
                continue
            label = self._category_label(cat)
            cap_full = extract_capsule_from_folder(folder, label)
            cap_name = _strip_gender_suffix(cap_full) or cap_full
            for ft in file_types:
                vlps_dir = self._find_vlps_dir(cap_path, ft)
                if vlps_dir is not None:
                    out.extend(self._walk_vlps_files_in_folder(
                        vlps_dir, leagues, ft, cap_name, cat,
                    ))
                else:
                    # Fallback: archivos planos con liga al final (*_{LIGA}.ext).
                    # Patrón HEADWEAR: Q3_FA25-ON CAMPUS_MLB.pdf
                    out.extend(self._walk_vlps_flat_files(
                        cap_path, leagues, ft, cap_name, cat,
                    ))

        # ── Classics ──
        for cs in scope.get("classics", []) or []:
            cat = cs.get("category")
            classic_type = cs.get("classic_type")
            subs = cs.get("subproducts") or []
            if not cat or not classic_type or not subs:
                continue
            label = self._category_label(cat)
            cap_full = extract_capsule_from_classic(classic_type, label)
            cap_name = _strip_gender_suffix(cap_full) or cap_full
            classic_base = self._find_classic_base_dir(self._category_to_path(cat))
            if classic_base is None:
                continue
            for sub in subs:
                ct_path = classic_base / classic_type
                if not ct_path.is_dir():
                    continue
                if sub == CLASSIC_DIRECT_KEY:
                    target_path = ct_path
                else:
                    target_path = ct_path / sub
                    if not target_path.is_dir():
                        continue
                for ft in file_types:
                    vlps_dir = self._find_vlps_dir(target_path, ft)
                    if vlps_dir is not None:
                        out.extend(self._walk_vlps_files_in_folder(
                            vlps_dir, leagues, ft, cap_name, cat,
                        ))
                    else:
                        # Fallback: subdirs de liga directamente en target_path.
                        # Patrón HEADWEAR _CLASSICS: _CLASSICS/CLASSIC LOGO·/MLB/[files]
                        out.extend(self._walk_vlps_files_in_folder(
                            target_path, leagues, ft, cap_name, cat,
                        ))

        return out

    def list_vlps_leagues_scoped(
        self,
        year: str,
        quarter: str,
        scope: dict,
        file_types: list[str],
    ) -> list[str]:
        """Ligas con archivos VLPS en el scope. Busca en _PPTX/_PDF VLPS."""
        leagues: set[str] = set()
        for sc in scope.get("seasonal", []) or []:
            cat = sc.get("category")
            folder = sc.get("capsule_folder")
            if not cat or not folder:
                continue
            cap_path = self._resolve_seasonal_cap_path(year, quarter, cat, folder)
            if cap_path is None:
                continue
            for ft in file_types:
                vlps_dir = self._find_vlps_dir(cap_path, ft)
                if vlps_dir is not None:
                    for entry in self._safe_iterdir(vlps_dir):
                        if entry.is_dir():
                            leagues.add(entry.name)
                else:
                    # Fallback: extraer liga del nombre del archivo (*_{LIGA}.ext).
                    # Patrón HEADWEAR: Q3_FA25-ON CAMPUS_MLB.pdf
                    exts = _VLPS_EXTENSIONS.get(ft, frozenset())
                    for f in self._safe_iterdir(cap_path):
                        if f.is_file() and f.suffix.lower() in exts:
                            parts = f.stem.rsplit("_", 1)
                            if len(parts) >= 2:
                                leagues.add(parts[-1])
        for cs in scope.get("classics", []) or []:
            cat = cs.get("category")
            classic_type = cs.get("classic_type")
            subs = cs.get("subproducts") or []
            if not cat or not classic_type or not subs:
                continue
            classic_base = self._find_classic_base_dir(self._category_to_path(cat))
            if classic_base is None:
                continue
            for sub in subs:
                ct_path = classic_base / classic_type
                if not ct_path.is_dir():
                    continue
                target_path = ct_path if sub == CLASSIC_DIRECT_KEY else ct_path / sub
                if not target_path.is_dir():
                    continue
                for ft in file_types:
                    vlps_dir = self._find_vlps_dir(target_path, ft)
                    if vlps_dir is not None:
                        for entry in self._safe_iterdir(vlps_dir):
                            if entry.is_dir():
                                leagues.add(entry.name)
                    else:
                        # Fallback: subdirs de liga directamente en target_path.
                        # Patrón HEADWEAR _CLASSICS: _CLASSICS/CLASSIC LOGO·/MLB/[files]
                        exts = _VLPS_EXTENSIONS.get(ft, frozenset())
                        for entry in self._safe_iterdir(target_path):
                            if not entry.is_dir() or entry.name.startswith("_"):
                                continue
                            if any(f.is_file() and f.suffix.lower() in exts
                                   for f in self._safe_iterdir(entry)):
                                leagues.add(entry.name)
        return sorted(leagues)

    def fetch_vlps_files(self, refs: list[VlpsFileRef]) -> list[bytes]:
        """Override paralelo de fetch_vlps_files usando el mismo
        ThreadPoolExecutor compartido con fetch_images. Cada PPTX/PDF
        es I/O-bound (lectura SMB) y los archivos pueden ser grandes
        — vale la pena paralelizar."""
        if not refs:
            return []
        return self._parallel_map(self.fetch_vlps_file, refs)

    def scan_team_scoped(
        self, year: str, quarter: str, scope: dict, league: str, team: str
    ) -> TeamScan:
        # Acumulador por gender × capsule_name → list[ImageRef].
        acc: dict[str, dict[str, list[ImageRef]]] = {g: {} for g in self.GENDERS}

        for mb, _cat, gender, cap_name in self._iter_scope_merchboards(year, quarter, scope):
            team_path = mb / league / team
            if not team_path.is_dir():
                continue
            for img in self._list_images(team_path):
                acc[gender].setdefault(cap_name, []).append(
                    ImageRef(key=str(img), filename=img.name)
                )

        by_gender: dict[str, list[CapsuleAvailability]] = {}
        for gender in self.GENDERS:
            caps: list[CapsuleAvailability] = []
            for cap_name in sorted(acc[gender].keys()):
                images = acc[gender][cap_name]
                caps.append(CapsuleAvailability(
                    capsule=cap_name,
                    image_count=len(images),
                    images=images,
                ))
            by_gender[gender] = caps

        return TeamScan(team=team, league=league, by_gender=by_gender)


# ═══════════════════════════════════════════════════════════════
# MOCK IMPLEMENTATION
# ═══════════════════════════════════════════════════════════════

class MockCatalogSource(CatalogSource):
    """Catálogo mock para desarrollo, testing y demos.

    Devuelve datos hardcodeados y genera PNGs placeholder en memoria.
    Determinístico: mismos inputs siempre producen los mismos resultados
    (clave para tests reproducibles).
    """

    GENDERS: tuple[str, ...] = ("MENS", "WOMENS", "KIDS")
    GENDER_CODE: dict[str, str] = {"MENS": "M", "WOMENS": "W", "KIDS": "K"}

    YEARS: list[str] = ["2024", "2025", "2026", "2027", "2028"]
    QUARTERS: list[str] = ["Q1", "Q2", "Q3", "Q4"]

    LEAGUES_TEAMS: dict[str, list[str]] = {
        "NFL": [
            "DALLAS COWBOYS", "LAS VEGAS RAIDERS", "NEW ENGLAND PATRIOTS",
            "GREEN BAY PACKERS", "KANSAS CITY CHIEFS",
        ],
        "NBA": [
            "LOS ANGELES LAKERS", "BOSTON CELTICS",
            "GOLDEN STATE WARRIORS", "MIAMI HEAT",
        ],
        "MLB": [
            "NEW YORK YANKEES", "LOS ANGELES DODGERS", "BOSTON RED SOX",
        ],
        "NHL": ["MONTREAL CANADIENS", "TORONTO MAPLE LEAFS"],
        "MLS": ["INTER MIAMI", "LA GALAXY"],
    }

    CAPSULES_BY_GENDER: dict[str, list[str]] = {
        "MENS":   ["FLAGSHIP", "SPRING BREAK", "TEAM CITY", "HERITAGE HUSTLE", "PRO FILE"],
        "WOMENS": ["FLAGSHIP", "SPRING BREAK", "PRO FILE"],
        "KIDS":   ["FLAGSHIP", "TEAM CITY"],
    }

    # Cantidad de imágenes por cápsula cuando "tiene contenido".
    IMAGES_PER_CAPSULE = 2

    def list_years(self) -> list[str]:
        return list(self.YEARS)

    def list_quarters(self, year: str) -> list[str]:
        return list(self.QUARTERS)

    def list_leagues(self, year: str, quarter: str) -> list[str]:
        return sorted(self.LEAGUES_TEAMS.keys())

    def list_teams(self, year: str, quarter: str, league: str) -> list[str]:
        return list(self.LEAGUES_TEAMS.get(league, []))

    def scan_team(self, year: str, quarter: str, league: str, team: str) -> TeamScan:
        by_gender: dict[str, list[CapsuleAvailability]] = {}
        for gender in self.GENDERS:
            code = self.GENDER_CODE[gender]
            capsules: list[CapsuleAvailability] = []
            for capsule in self.CAPSULES_BY_GENDER.get(gender, []):
                # Determinismo cross-process: zlib.crc32 sobre un string
                # canónico decide si la carpeta está "vacía" (~10%).
                # Mismos inputs → mismo resultado en cualquier máquina.
                signature = f"{gender}|{capsule}|{team}".encode("utf-8")
                is_empty = (zlib.crc32(signature) % 10) == 0
                count = 0 if is_empty else self.IMAGES_PER_CAPSULE

                images: list[ImageRef] = []
                for i in range(1, count + 1):
                    filename = f"{league}_{team}_{capsule}_{code}_{i:03d}.png"
                    # En el mock, key == filename; fetch_image la usa para
                    # extraer info y dibujar el placeholder.
                    images.append(ImageRef(key=filename, filename=filename))

                capsules.append(CapsuleAvailability(
                    capsule=capsule,
                    image_count=count,
                    images=images,
                ))
            by_gender[gender] = capsules
        return TeamScan(team=team, league=league, by_gender=by_gender)

    def fetch_image(self, ref: ImageRef) -> bytes:
        """Genera un PNG 1200x900 con texto identificable.

        Extrae cápsula + código de género del filename para mostrarlos
        en grande — útil para validar visualmente el matching cuando se
        abre el deck generado.
        """
        # filename format: LIGA_TEAM_CAPSULE_CODE_NNN.png
        base = ref.filename.rsplit(".", 1)[0]
        parts = base.split("_")
        if len(parts) >= 4:
            # parts[-1] = NNN, parts[-2] = code, parts[2..-3] = capsule
            label = f"{' '.join(parts[2:-2])} {parts[-2]}"
        else:
            label = base
        return _make_placeholder_png(label=label, sublabel=ref.filename)

    # ─── Listings granulares (modo avanzado) ─────────────────────
    # Implementaciones simples in-memory para tests. NO replican la
    # estructura real del server — solo proveen datos plausibles para
    # validar el flujo del modo avanzado en tests/dev.

    # Categorías que el mock "tiene": MENS + WOMENS + KIDS/BOYS + KIDS/GIRLS.
    MOCK_CATEGORIES: list[str] = ["MENS", "WOMENS", "KIDS/BOYS", "KIDS/GIRLS"]

    # Classics hardcoded — independientes de la categoría real, los devolvemos
    # para todas las categorías.
    MOCK_CLASSICS: list[str] = ["_CLASSIC ICON", "HERITAGE HUSTLE", "PROPERTY OF"]

    # Sub-productos por classic (también hardcoded).
    MOCK_CLASSIC_SUBPRODUCTS: dict[str, list[str]] = {
        "_CLASSIC ICON":   [CLASSIC_DIRECT_KEY, "JERSEY"],
        "HERITAGE HUSTLE": [CLASSIC_DIRECT_KEY, "HOODIE", "T-SHIRT"],
        "PROPERTY OF":     [CLASSIC_DIRECT_KEY],
    }

    def list_categories(self) -> list[str]:
        return list(self.MOCK_CATEGORIES)

    def list_capsules(self, year: str, quarter: str, category: str) -> list[dict]:
        """Devuelve las cápsulas que el mock asociaría a esa categoría.

        El "folder_name" es sintético — útil solo para que el cliente
        identifique unívocamente esa cápsula al pasarla en un scope.
        """
        # Mapear category → gender bucket del mock.
        cat_label = category.rsplit("/", 1)[-1]
        if category.startswith("KIDS") or cat_label in ("BOYS", "GIRLS", "KIDS", "PLAYER N&N"):
            gender_bucket = "KIDS"
            cat_short = cat_label  # ej. BOYS
        elif category == "MENS":
            gender_bucket = "MENS"
            cat_short = "MNS"
        elif category == "WOMENS":
            gender_bucket = "WOMENS"
            cat_short = "WMNS"
        else:
            return []

        capsules = self.CAPSULES_BY_GENDER.get(gender_bucket, [])
        code = self.GENDER_CODE[gender_bucket]
        items = []
        for i, cap in enumerate(capsules, start=1):
            folder_name = f"{i:02d} - {quarter} {year} {cat_short} {cap}"
            capsule_key = f"{cap} {code}"
            items.append({
                "folder_name": folder_name,
                "capsule_key": capsule_key,
            })
        return items

    def list_classics(self, category: str) -> list[str]:
        if category not in self.MOCK_CATEGORIES:
            return []
        return list(self.MOCK_CLASSICS)

    def list_classic_subproducts(
        self, category: str, classic_type: str
    ) -> list[str]:
        if category not in self.MOCK_CATEGORIES:
            return []
        return list(self.MOCK_CLASSIC_SUBPRODUCTS.get(classic_type, []))

    # ─── Métodos scoped (modo avanzado) ──────────────────────────
    # Mock no tiene filesystem real — solo filtra los resultados del
    # scan_team global por las cápsulas presentes en el scope.

    def _scope_wanted_capsules(self, scope: dict) -> dict[str, set[str]]:
        """Construye el set de capsule_names esperados por gender, a
        partir del scope. Útil para filtrar el TeamScan global del mock.
        """
        wanted: dict[str, set[str]] = {g: set() for g in self.GENDERS}

        for sc in scope.get("seasonal", []) or []:
            cat = sc.get("category") or ""
            folder = sc.get("capsule_folder") or ""
            label = cat.rsplit("/", 1)[-1]
            gender_code = CATEGORY_TO_GENDER.get(label.upper())
            gender = {"M": "MENS", "W": "WOMENS", "K": "KIDS"}.get(gender_code or "")
            if not gender:
                continue
            cap_full = extract_capsule_from_folder(folder, label)
            cap_name = _strip_gender_suffix(cap_full)
            if cap_name:
                wanted[gender].add(cap_name)

        for cs in scope.get("classics", []) or []:
            cat = cs.get("category") or ""
            classic_type = cs.get("classic_type") or ""
            label = cat.rsplit("/", 1)[-1]
            gender_code = CATEGORY_TO_GENDER.get(label.upper())
            gender = {"M": "MENS", "W": "WOMENS", "K": "KIDS"}.get(gender_code or "")
            if not gender:
                continue
            cap_full = extract_capsule_from_classic(classic_type, label)
            cap_name = _strip_gender_suffix(cap_full)
            if cap_name:
                wanted[gender].add(cap_name)

        return wanted

    def list_leagues_scoped(self, year: str, quarter: str, scope: dict) -> list[str]:
        # Si el scope está vacío, devolver vacío (sin nada que walkear).
        wanted = self._scope_wanted_capsules(scope)
        if not any(wanted.values()):
            return []
        # Para el mock, las ligas son las mismas en todas las cápsulas.
        return self.list_leagues(year, quarter)

    def list_teams_scoped(
        self, year: str, quarter: str, scope: dict, league: str
    ) -> list[str]:
        wanted = self._scope_wanted_capsules(scope)
        if not any(wanted.values()):
            return []
        return self.list_teams(year, quarter, league)

    def scan_team_scoped(
        self, year: str, quarter: str, scope: dict, league: str, team: str
    ) -> TeamScan:
        # Tomamos el scan_team global y filtramos por las cápsulas del scope.
        full = self.scan_team(year, quarter, league, team)
        wanted = self._scope_wanted_capsules(scope)

        new_by_gender: dict[str, list[CapsuleAvailability]] = {}
        for gender, caps in full.by_gender.items():
            new_by_gender[gender] = [
                c for c in caps if c.capsule in wanted.get(gender, set())
            ]
        return TeamScan(team=team, league=league, by_gender=new_by_gender)

    # ─── VLPS para Mock (G.1) ──────────────────────────────────────
    # Devolvemos un set ficticio de refs para que tests de integración
    # puedan ejercitar el flujo sin un servidor SMB real.

    def scan_vlps_files(
        self,
        year: str,
        quarter: str,
        scope: dict,
        leagues: list[str],
        file_types: list[str],
    ) -> list[VlpsFileRef]:
        if not file_types or not all(ft in VLPS_FILE_TYPES for ft in file_types):
            raise ValueError(
                f"file_types inválido: {file_types!r} "
                f"(esperaba subset de {VLPS_FILE_TYPES})"
            )
        out: list[VlpsFileRef] = []
        wanted = self._scope_wanted_capsules(scope)
        # Por cada gender que tenga cápsulas pedidas, generamos 1 file
        # por (cápsula × tipo × liga). FLAT lo simulamos como liga="".
        for sc in scope.get("seasonal", []) or []:
            cat = sc.get("category") or ""
            folder = sc.get("capsule_folder") or ""
            label = cat.rsplit("/", 1)[-1]
            cap_full = extract_capsule_from_folder(folder, label)
            cap_name = _strip_gender_suffix(cap_full) or cap_full
            for ft in file_types:
                ext = ".pptx" if ft == "ppt" else ".pdf"
                for liga in leagues:
                    key = f"mock://{cat}/{cap_name}/{ft}/{liga}{ext}"
                    out.append(VlpsFileRef(
                        key=key,
                        filename=f"{liga}_{cap_name.replace(' ', '_')}{ext}",
                        file_type=ft,
                        capsule=cap_name,
                        league=liga,
                        category=cat,
                    ))
        return out

    def list_vlps_leagues_scoped(
        self,
        year: str,
        quarter: str,
        scope: dict,
        file_types: list[str],
    ) -> list[str]:
        wanted = self._scope_wanted_capsules(scope)
        if not any(wanted.values()):
            return []
        return self.list_leagues(year, quarter)

    def fetch_vlps_file(self, ref: VlpsFileRef) -> bytes:
        """Mock: devuelve un payload corto deterministic — útil para tests
        de empaquetado sin generar archivos reales."""
        # Marker mínimo: la key + el filename. Suficiente para que un test
        # verifique que el ZIP contiene el archivo correcto y que los
        # bytes son únicos por ref.
        return f"MOCK_VLPS|{ref.file_type}|{ref.capsule}|{ref.league}|{ref.filename}".encode("utf-8")


def _make_placeholder_png(label: str, sublabel: str) -> bytes:
    """Genera bytes de un PNG 1200x900 con un label grande amarillo y un
    sublabel chico gris sobre fondo oscuro.
    """
    img = Image.new("RGB", (1200, 900), color=(26, 26, 26))
    draw = ImageDraw.Draw(img)

    font_big = ImageFont.load_default(size=56)
    font_small = ImageFont.load_default(size=22)

    draw.text((600, 430), label, fill=(232, 200, 74), anchor="mm", font=font_big)
    draw.text((600, 490), sublabel, fill=(136, 136, 136), anchor="mm", font=font_small)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════
# FACTORY
# ═══════════════════════════════════════════════════════════════

def get_catalog_source(
    kind: str,
    base_path: str | None = None,
    *,
    cache_ttl: int = CACHE_TTL_DEFAULT,
    cache_maxsize: int = CACHE_MAXSIZE_DEFAULT,
    walk_workers: int = CACHE_WORKERS_DEFAULT,
) -> CatalogSource:
    """Devuelve la implementación de CatalogSource adecuada según config.

    Args:
        kind: 'smb' o 'mock'.
        base_path: ruta base UNC, requerida para 'smb'; ignorada para 'mock'.
        cache_ttl: TTL en segundos del caché del Smb (kwarg-only).
            Ignorado para 'mock'. 0 = caché desactivado.
        cache_maxsize: maxsize por nivel de caché del Smb (kwarg-only).
            Ignorado para 'mock'.
        walk_workers: workers paralelos del ThreadPoolExecutor del Smb
            (kwarg-only). Ignorado para 'mock'. 1 = seriado.

    Raises:
        ValueError: si `kind` es desconocido, o si 'smb' sin base_path.
    """
    if kind == "smb":
        if not base_path:
            raise ValueError("base_path es requerido para CatalogSource 'smb'")
        return SmbCatalogSource(
            base_path,
            cache_ttl=cache_ttl,
            cache_maxsize=cache_maxsize,
            walk_workers=walk_workers,
        )
    if kind == "mock":
        return MockCatalogSource()
    raise ValueError(f"CatalogSource desconocido: {kind!r} (usa 'smb' o 'mock')")
