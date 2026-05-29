'use strict';

// ═══════════════════════════════════════════════════════════════
// API_TOKEN (inyectado por Jinja2 en el meta tag)
// ═══════════════════════════════════════════════════════════════
const API_TOKEN = (() => {
  const m = document.querySelector('meta[name="api-token"]');
  return m ? m.content : '';
})();

// ═══════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════
const state = {
  pptxFile: null,
  // B.2d / G.2: modo del OUTPUT (ortogonal al modo de selección de step 2).
  //   'existing' — flujo histórico: sube un PPT base y se reemplazan imágenes.
  //   'blank'    — modo "PPT nuevo desde cero": no hay PPT base, se genera
  //                un deck nuevo con las imágenes del scope (requiere advanced).
  //   'vlps'     — modo "Descargar VLPS": no genera nada, baja archivos
  //                .pptx/.pdf pre-armados del servidor (requiere advanced+server).
  outputMode: 'existing',
  // G.2: filtro de tipos para VLPS. Subset de ['ppt', 'pdf']; vacío = error.
  vlpsFileTypes: ['ppt', 'pdf'],
  // true cuando el servidor devuelve 0 ligas para el scope actual (ej. HEADWEAR
  // sin estructura _PPTX VLPS/). En ese caso la liga no es requerida y se
  // descargan todos los archivos disponibles en las cápsulas.
  vlpsNoLeaguesAvailable: false,
  // D.2: origen de las imágenes (ortogonal a outputMode y selection.mode).
  //   'server' — fetch del catálogo SMB (flujo histórico).
  //   'manual' — usuario sube los merchboards directamente.
  imageOrigin: 'server',
  // D.2: estado de los uploads manuales. Data shape multi-ready (dict por
  // (liga, team)) aunque D.1 enforza single-team. Las claves son
  // "LIGA::TEAM" para poder serializar en logs y rendering.
  manualImagesByTeam: {},  // { "NFL::DALLAS COWBOYS": [{file, parsed}, ...] }
  manualUnmatched: [],     // [{name}, ...] — archivos sin naming válido
  manualFlatFiles: [],     // [File, ...] orden de inserción, para el POST
  // D.4b/D.5b: modo de procesamiento cuando hay archivos de varios teams.
  //   'strict'    — rechaza multi-team (modo seguro).
  //   'mixed'     — mezcla todos los teams en un solo deck.
  //   'per_team'  — un deck por team, ZIP con todos.
  // El default por flow se calcula en getDefaultMultiTeamMode(); el valor
  // efectivo se persiste solo cuando el usuario lo elige explícitamente.
  multiTeamMode: 'strict',
  multiTeamModeExplicit: false,  // true si el user clickeó un radio
  selection: {
    // Modo de selección (simple = flujo histórico; advanced = scoped granular).
    mode: 'simple',

    // ─── Modo simple (estado heredado) ─────────────────────────
    year: null,
    quarter: null,
    leagues: [],   // array de strings (nombres de liga)
    teams: [],     // array de {league, team}

    // ─── Modo avanzado (B.1c) ─────────────────────────────────
    // Independiente del modo simple — se persiste por separado.
    adv: {
      categories: [],   // ["MENS", "KIDS/BOYS", ...]
      year: null,
      quarter: null,
      // Cápsulas seleccionadas: [{category, folder_name, capsule_key}].
      // Conservamos folder_name porque es lo que el backend usa para
      // identificar exactamente qué carpeta walkear.
      capsules: [],
      // Classics: [{category, classic_type}].
      classics: [],
      leagues: [],
      teams: [],
    },
  },
  scanResult: null,
  // outputBlob ya no se usa — la descarga ahora la maneja el browser
  // nativamente via <a href> apuntando a outputDownloadUrl. Se deja en
  // null por compat con cualquier path legacy que pudiera leerlo.
  outputBlob: null,
  outputDownloadUrl: null,
  outputFilename: null,

  // Caché en memoria de sub-productos por (category|classic_type) → list[str].
  // Se llena cuando cargamos los classics; al enviar el scope, incluimos
  // todos los sub-productos de cada classic seleccionado.
  _advSubproductsCache: {},
};

// ═══════════════════════════════════════════════════════════════
// DOM REFS (cache)
// ═══════════════════════════════════════════════════════════════
const el = {};
function cacheEl() {
  const ids = [
    // Step 1
    'inputPPT', 'dropPPT', 'dropEmpty', 'dropLoaded', 'pptName', 'slidesFound', 'btnNext1',
    // Step 2 — single-select containers (Año, Quarter)
    'ssYear', 'ssQuarter', 'btnScan',
    'scanLoading', 'scanError', 'scanErrorBox', 'scanPreview', 'scanSummary', 'scanTree',
    'btnBack2', 'btnNext2',
    // Step 2 — multi-select containers (Ligas, Equipos)
    'msLeagues', 'msTeams',
    // Step 2 — modo avanzado
    'modeSimple', 'modeAdvanced', 'simpleSelectors', 'advancedSelectors',
    'advCategories', 'advYear', 'advQuarter',
    'advCapsules', 'advClassics', 'advLeagues', 'advTeams',
    // Step 3
    'preGenInfo', 'progressFill', 'logBox', 'resultCard', 'resultText',
    'btnDownload', 'btnBack3', 'btnGenerate', 'btnReset',
    // Containers
    'step1', 'step2', 'step3',
    // Stepper
    'stepper1', 'stepper2', 'stepper3',
    // B.2d / G.2 — toggle output mode (PPT existente vs PPT nuevo vs VLPS)
    'outputModeToggle', 'outputModeExisting', 'outputModeBlank', 'outputModeVlps',
    // G.2 — sub-toggle de tipos de archivo VLPS (solo visible en VLPS)
    'vlpsFileTypesRow', 'vlpsFileTypesPpt', 'vlpsFileTypesPdf', 'vlpsFileTypesBoth',
    // D.5b — row del modo multi-team (se oculta en VLPS)
    'processingModeRow',
    // C.2 — toggle theme (claro / oscuro)
    'themeToggle',
    // D.2 — toggle origen de imágenes (servidor / manual) + manual UI
    'imageOriginToggle', 'imageOriginServer', 'imageOriginManual',
    'serverSourceSection', 'manualSourceSection',
    'dropImages', 'inputImages',
    'dropImagesEmpty', 'dropImagesLoaded', 'manualImagesCount',
    'manualMultiTeamWarning', 'manualTeamBadge',
    'manualPreview', 'manualUnmatched',
    'btnClearImages',
    // D.4b / D.5b — 3 radios inline para multi_team_mode (Opción C)
    'multiTeamStrict', 'multiTeamMixed', 'multiTeamPerTeam',
    // G.2-fix — status box del modo VLPS (feedback visual de qué falta /
    // si está listo para descargar)
    'vlpsStatus',
    // G.2-fix2 — textos del Step 3 + stepper3 que cambian por outputMode
    // para que el flujo VLPS no diga "Generar Deck" / "0 equipos" / etc.
    'stepper3Label', 'step3Heading', 'step3Subtitle', 'resultTitle',
    // Modal de confirmación
    'confirmModal', 'confirmTitle', 'confirmMessage', 'confirmOk', 'confirmCancel',
    // Global
    'apiError',
  ];
  for (const id of ids) el[id] = document.getElementById(id);
  el.steps = [null, el.step1, el.step2, el.step3];
  el.steppers = [null, el.stepper1, el.stepper2, el.stepper3];
  // B.2d — stepper lines (entre step 1-2 y 2-3) sin IDs → query directo.
  // [0] = entre 1 y 2 (la ocultamos en modo blank). [1] = entre 2 y 3 (siempre).
  el.stepperLines = document.querySelectorAll('.stepper .stepper-line');
}
cacheEl();

// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════
function esc(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ═══════════════════════════════════════════════════════════════
// SINGLEFLIGHT — anti-race para llamadas que se disparan en cascada
//
// Patrón: una llamada nueva aborta la anterior y le quita validez.
// Triple defensa contra carreras cuando el backend es lento (≥segundos):
//   1. AbortController cancela la request HTTP previa (libera red).
//   2. Generation counter descarta respuestas viejas que igual hayan
//      vuelto antes de procesarse el abort.
//   3. El `work` callback recibe `isCurrent()` y debe chequearlo antes
//      de mutar UI/state — última línea de defensa.
//
// Uso:
//   const loader = createSingleflight();
//   await loader(async (signal, isCurrent) => {
//     const r = await apiFetch(url, { signal });
//     const data = await r.json();
//     if (!isCurrent()) return;
//     msTeams.setOptions(data);
//   });
//
// Si dentro de `work` se quiere catchear errores específicos, hay que
// re-throw los AbortError para que el singleflight los trague — si no,
// dispararían `showApiError` por una cancelación deliberada.
// ═══════════════════════════════════════════════════════════════
function createSingleflight() {
  let currentCtrl = null;
  let currentGen = 0;
  return async function run(work) {
    if (currentCtrl) currentCtrl.abort();
    currentCtrl = new AbortController();
    const myGen = ++currentGen;
    const signal = currentCtrl.signal;
    const isCurrent = () => myGen === currentGen;
    try {
      return await work(signal, isCurrent);
    } catch (e) {
      if (e.name === 'AbortError') return undefined;
      throw e;
    }
  };
}

// Una instancia por endpoint que se dispara en cascada con cambios del
// usuario. `loadYears` no necesita — se llama una sola vez en init.
const pptScanLoader  = createSingleflight();
const quartersLoader = createSingleflight();
const leaguesLoader  = createSingleflight();
const teamsLoader    = createSingleflight();
const scanLoader     = createSingleflight();

// Loaders del modo avanzado (B.1c).
const advCategoriesLoader  = createSingleflight();
const advCapsulesLoader    = createSingleflight();
const advClassicsLoader    = createSingleflight();
const advLeaguesLoader     = createSingleflight();
const advTeamsLoader       = createSingleflight();

// ═══════════════════════════════════════════════════════════════
// MULTI-SELECT COMPONENT
// Dropdown custom con checkboxes, search y "seleccionar todas".
// Setea options con [{value, label, subtitle?}] y notifica cambios.
// ═══════════════════════════════════════════════════════════════
class MultiSelect {
  constructor(rootElement, { placeholder = '— Selecciona —', onChange } = {}) {
    this.root = rootElement;
    this.placeholder = placeholder;
    this.onChange = onChange;
    this.options = [];
    this.selected = new Set();
    this.searchTerm = '';
    this.isOpen = false;
    this._build();
    this._render();
  }

  setOptions(options) {
    this.options = options;
    // Filtrar valores seleccionados que ya no existen en las opciones nuevas
    this.selected = new Set([...this.selected].filter(v => options.some(o => o.value === v)));
    this._render();
  }

  setSelected(values) {
    this.selected = new Set(values);
    this._render();
  }

  getSelected() {
    return [...this.selected];
  }

  setDisabled(disabled) {
    this.trigger.disabled = disabled;
    if (disabled && this.isOpen) this._close();
  }

  _build() {
    this.root.innerHTML = `
      <button type="button" class="multiselect-trigger" aria-haspopup="listbox" aria-expanded="false">
        <span class="multiselect-label placeholder">${esc(this.placeholder)}</span>
        <span class="multiselect-chevron" aria-hidden="true"></span>
      </button>
      <div class="multiselect-panel" hidden>
        <input type="text" class="multiselect-search" placeholder="Buscar...">
        <label class="multiselect-option multiselect-all">
          <input type="checkbox" class="multiselect-all-cb">
          <span>Seleccionar todas</span>
        </label>
        <div class="multiselect-divider"></div>
        <div class="multiselect-options"></div>
      </div>
    `;
    this.trigger = this.root.querySelector('.multiselect-trigger');
    this.panel = this.root.querySelector('.multiselect-panel');
    this.label = this.root.querySelector('.multiselect-label');
    this.searchInput = this.root.querySelector('.multiselect-search');
    this.allCb = this.root.querySelector('.multiselect-all-cb');
    this.optionsContainer = this.root.querySelector('.multiselect-options');

    this.trigger.addEventListener('click', () => this._toggleOpen());
    this.searchInput.addEventListener('input', e => {
      this.searchTerm = e.target.value.toLowerCase();
      this._renderOptions();
    });
    this.allCb.addEventListener('change', e => this._toggleAll(e.target.checked));

    // Click fuera del componente cierra el panel
    document.addEventListener('click', e => {
      if (!this.root.contains(e.target) && this.isOpen) this._close();
    });

    // Escape cierra
    this.root.addEventListener('keydown', e => {
      if (e.key === 'Escape' && this.isOpen) {
        e.preventDefault();
        this._close();
        this.trigger.focus();
      }
    });
  }

  _toggleOpen() {
    if (this.trigger.disabled) return;
    if (this.isOpen) this._close();
    else this._open();
  }

  _open() {
    this.isOpen = true;
    this.panel.hidden = false;
    this.trigger.setAttribute('aria-expanded', 'true');
    this.searchInput.value = '';
    this.searchTerm = '';
    this._renderOptions();
    setTimeout(() => this.searchInput.focus(), 0);
  }

  _close() {
    this.isOpen = false;
    this.panel.hidden = true;
    this.trigger.setAttribute('aria-expanded', 'false');
  }

  _toggleAll(checked) {
    if (checked) this.options.forEach(o => this.selected.add(o.value));
    else this.selected.clear();
    this._render();
    this._notify();
  }

  _render() {
    this._renderLabel();
    this._renderOptions();
    this._renderAllCheckbox();
  }

  _renderLabel() {
    const count = this.selected.size;
    const total = this.options.length;
    if (count === 0) {
      this.label.textContent = this.placeholder;
      this.label.classList.add('placeholder');
    } else if (total > 0 && count === total) {
      this.label.textContent = `Todas (${count})`;
      this.label.classList.remove('placeholder');
    } else if (count === 1) {
      const val = [...this.selected][0];
      const opt = this.options.find(o => o.value === val);
      this.label.textContent = opt?.label || val;
      this.label.classList.remove('placeholder');
    } else {
      this.label.textContent = `${count} seleccionadas`;
      this.label.classList.remove('placeholder');
    }
  }

  _renderOptions() {
    if (this.options.length === 0) {
      this.optionsContainer.innerHTML = '<div class="multiselect-empty">Sin opciones disponibles</div>';
      return;
    }
    const filtered = this.options.filter(o =>
      !this.searchTerm ||
      o.label.toLowerCase().includes(this.searchTerm) ||
      (o.subtitle?.toLowerCase().includes(this.searchTerm))
    );
    if (filtered.length === 0) {
      this.optionsContainer.innerHTML = '<div class="multiselect-empty">Sin resultados</div>';
      return;
    }
    this.optionsContainer.innerHTML = filtered.map(o => `
      <label class="multiselect-option">
        <input type="checkbox" value="${esc(o.value)}"${this.selected.has(o.value) ? ' checked' : ''}>
        <span class="multiselect-opt-label">${esc(o.label)}</span>
        ${o.subtitle ? `<span class="multiselect-opt-sub">${esc(o.subtitle)}</span>` : ''}
      </label>
    `).join('');

    this.optionsContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', e => {
        if (e.target.checked) this.selected.add(e.target.value);
        else this.selected.delete(e.target.value);
        this._renderLabel();
        this._renderAllCheckbox();
        this._notify();
      });
    });
  }

  _renderAllCheckbox() {
    const total = this.options.length;
    const sel = this.selected.size;
    this.allCb.checked = total > 0 && sel === total;
    this.allCb.indeterminate = sel > 0 && sel < total;
    this.allCb.disabled = total === 0;
  }

  _notify() {
    this.onChange?.([...this.selected]);
  }
}

// ═══════════════════════════════════════════════════════════════
// SINGLE-SELECT COMPONENT
// Dropdown custom de valor único — sin checkboxes, sin search.
// Click en opción → setea valor y cierra panel. Item activo en accent.
// Reusa los estilos del MultiSelect (.multiselect-trigger, .multiselect-panel)
// más .ss-option para las filas.
// ═══════════════════════════════════════════════════════════════
class SingleSelect {
  constructor(rootElement, { placeholder = '— Selecciona —', onChange } = {}) {
    this.root = rootElement;
    this.placeholder = placeholder;
    this.onChange = onChange;
    this.options = [];
    this.selected = null;
    this.isOpen = false;
    this._build();
    this._render();
  }

  setOptions(options) {
    this.options = options;
    // Si el valor previo ya no existe en las opciones nuevas, limpiarlo
    if (this.selected !== null && !options.some(o => o.value === this.selected)) {
      this.selected = null;
    }
    this._render();
  }

  setSelected(value) {
    this.selected = value || null;
    this._render();
  }

  getSelected() {
    return this.selected;
  }

  setDisabled(disabled) {
    this.trigger.disabled = disabled;
    if (disabled && this.isOpen) this._close();
  }

  _build() {
    this.root.innerHTML = `
      <button type="button" class="multiselect-trigger" aria-haspopup="listbox" aria-expanded="false">
        <span class="multiselect-label placeholder">${esc(this.placeholder)}</span>
        <span class="multiselect-chevron" aria-hidden="true"></span>
      </button>
      <div class="multiselect-panel" hidden>
        <div class="multiselect-options"></div>
      </div>
    `;
    this.trigger = this.root.querySelector('.multiselect-trigger');
    this.panel = this.root.querySelector('.multiselect-panel');
    this.label = this.root.querySelector('.multiselect-label');
    this.optionsContainer = this.root.querySelector('.multiselect-options');

    this.trigger.addEventListener('click', () => this._toggleOpen());

    document.addEventListener('click', e => {
      if (!this.root.contains(e.target) && this.isOpen) this._close();
    });

    this.root.addEventListener('keydown', e => {
      if (e.key === 'Escape' && this.isOpen) {
        e.preventDefault();
        this._close();
        this.trigger.focus();
      }
    });
  }

  _toggleOpen() {
    if (this.trigger.disabled) return;
    if (this.isOpen) this._close();
    else this._open();
  }

  _open() {
    this.isOpen = true;
    this.panel.hidden = false;
    this.trigger.setAttribute('aria-expanded', 'true');
    this._renderOptions();
  }

  _close() {
    this.isOpen = false;
    this.panel.hidden = true;
    this.trigger.setAttribute('aria-expanded', 'false');
  }

  _select(value) {
    this.selected = value;
    this._renderLabel();
    this._close();
    this.onChange?.(value);
  }

  _render() {
    this._renderLabel();
    this._renderOptions();
  }

  _renderLabel() {
    if (!this.selected) {
      this.label.textContent = this.placeholder;
      this.label.classList.add('placeholder');
    } else {
      const opt = this.options.find(o => o.value === this.selected);
      this.label.textContent = opt?.label || this.selected;
      this.label.classList.remove('placeholder');
    }
  }

  _renderOptions() {
    if (this.options.length === 0) {
      this.optionsContainer.innerHTML = '<div class="multiselect-empty">Sin opciones disponibles</div>';
      return;
    }
    this.optionsContainer.innerHTML = this.options.map(o => `
      <button type="button" class="ss-option${o.value === this.selected ? ' selected' : ''}" data-value="${esc(o.value)}">
        ${esc(o.label)}
      </button>
    `).join('');

    this.optionsContainer.querySelectorAll('.ss-option').forEach(btn => {
      btn.addEventListener('click', () => this._select(btn.dataset.value));
    });
  }
}

const MAX_PPTX_SIZE = 1000 * 1024 * 1024;  // 1000 MB
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function validatePptx(file) {
  if (!file) return 'No se seleccionó archivo';
  if (!file.name.toLowerCase().endsWith('.pptx')) return 'El archivo debe ser .pptx';
  if (file.size === 0) return 'El archivo está vacío';
  if (file.size > MAX_PPTX_SIZE) return `El PPT excede ${MAX_PPTX_SIZE / 1024 / 1024} MB`;
  if (file.type && file.type !== PPTX_MIME && file.type !== 'application/zip') {
    return `Tipo de archivo inesperado: ${file.type}`;
  }
  return null;
}

const GENDERS = ['MENS', 'WOMENS', 'KIDS'];

function parseFilenameFromContentDisposition(header) {
  if (!header) return null;
  // RFC 5987: filename*=UTF-8''...
  let m = header.match(/filename\*=UTF-8''([^;]+)/i);
  if (m) {
    try { return decodeURIComponent(m[1]); } catch (_) {}
  }
  // Fallback: filename="..."
  m = header.match(/filename="?([^";]+)"?/i);
  return m ? m[1] : null;
}

// ═══════════════════════════════════════════════════════════════
// API CLIENT
// ═══════════════════════════════════════════════════════════════
class ApiError extends Error {
  constructor(msg, status) {
    super(msg);
    this.status = status;
  }
}

async function apiFetch(path, opts = {}) {
  const headers = {
    ...(opts.headers || {}),
    'Authorization': `Bearer ${API_TOKEN}`,
  };
  const response = await fetch(path, { ...opts, headers });
  if (!response.ok) {
    let detail = `HTTP ${response.status}`;
    try {
      const body = await response.json();
      if (body.detail) {
        // FastAPI 422 devuelve `detail` como array de objetos de validación
        // (no string). Sin esto se mostraba "[object Object]". Lo formateamos
        // a un mensaje legible: "campo: mensaje; campo2: mensaje2".
        if (typeof body.detail === 'string') {
          detail = body.detail;
        } else if (Array.isArray(body.detail)) {
          detail = body.detail
            .map(d => d && d.msg
              ? `${Array.isArray(d.loc) ? d.loc.join('.') + ': ' : ''}${d.msg}`
              : JSON.stringify(d))
            .join('; ');
        } else {
          detail = JSON.stringify(body.detail);
        }
      }
    } catch (_) { /* respuesta no es JSON, dejamos el HTTP status */ }
    throw new ApiError(detail, response.status);
  }
  return response;
}

// ═══════════════════════════════════════════════════════════════
// Global error banner
// ═══════════════════════════════════════════════════════════════
function showApiError(msg) {
  el.apiError.textContent = `⚠ ${msg}  (clic para cerrar)`;
  el.apiError.style.display = 'block';
}

function hideApiError() {
  el.apiError.style.display = 'none';
}
el.apiError.addEventListener('click', hideApiError);

// ═══════════════════════════════════════════════════════════════
// STEP 1: PPT (upload + scan via backend)
// ═══════════════════════════════════════════════════════════════
el.inputPPT.addEventListener('change', e => handlePPT(e.target.files[0]));

el.dropPPT.addEventListener('dragover', e => {
  e.preventDefault();
  el.dropPPT.classList.add('drag-over');
});
el.dropPPT.addEventListener('dragleave', () => el.dropPPT.classList.remove('drag-over'));
el.dropPPT.addEventListener('drop', e => {
  e.preventDefault();
  el.dropPPT.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file && file.name.endsWith('.pptx')) handlePPT(file);
});
el.dropPPT.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    el.inputPPT.click();
  }
});

function setDropEmpty() {
  el.dropEmpty.hidden = false;
  el.dropLoaded.hidden = true;
}

function setDropLoaded(filename) {
  el.dropEmpty.hidden = true;
  el.dropLoaded.hidden = false;
  el.pptName.textContent = filename;
  el.pptName.title = filename;  // tooltip con el nombre completo en hover
}

async function handlePPT(file) {
  if (!file) return;
  const err = validatePptx(file);
  if (err) {
    // No marcamos el drop-zone como loaded — el archivo es inválido.
    setDropEmpty();
    el.slidesFound.style.display = 'block';
    el.slidesFound.innerHTML = `<div class="slides-found-empty">⚠ ${esc(err)}</div>`;
    el.btnNext1.disabled = true;
    return;
  }
  state.pptxFile = file;
  setDropLoaded(file.name);
  el.slidesFound.style.display = 'block';
  el.slidesFound.innerHTML = `<span style="color:var(--text-muted)">Escaneando notas en el servidor...</span>`;
  el.btnNext1.disabled = true;

  await pptScanLoader(async (signal, isCurrent) => {
    try {
      const form = new FormData();
      form.append('file', file);
      const r = await apiFetch('/api/scan-ppt', { method: 'POST', body: form, signal });
      const data = await r.json();
      if (!isCurrent()) return;
      renderPptScan(data);
      el.btnNext1.disabled = false;
    } catch (e) {
      if (e.name === 'AbortError') throw e;  // dejar que singleflight lo trague
      el.slidesFound.innerHTML = `<div class="slides-found-empty">Error escaneando el PPT: ${esc(e.message)}</div>`;
      el.btnNext1.disabled = true;
    }
  });
}

function renderPptScan(scanResult) {
  const found = scanResult.found || [];
  if (found.length > 0) {
    const items = found
      .map(s => `<li>Slide ${esc(s.slide)} <span class="arrow">→</span> ${esc(s.key)}</li>`)
      .join('');
    el.slidesFound.innerHTML = `
      <details class="slides-found-details">
        <summary>✓ ${found.length} slide(s) de merchboard encontradas</summary>
        <ul>${items}</ul>
      </details>
    `;
  } else {
    el.slidesFound.innerHTML = `
      <div class="slides-found-empty">
        ⚠ No se encontraron slides con notas de cápsulas. Asegúrate de haber agregado el nombre de la cápsula en las notas.
      </div>
    `;
  }
}

// ═══════════════════════════════════════════════════════════════
// STEP 2: SELECTORES + SCAN DEL CATÁLOGO
// ═══════════════════════════════════════════════════════════════

function refreshScanButton() {
  const s = state.selection;
  if (s.mode === 'advanced') {
    const adv = s.adv;
    const hasScope = (adv.capsules.length + adv.classics.length) > 0;
    el.btnScan.disabled = !(adv.year && adv.quarter && hasScope && adv.teams.length > 0);
  } else {
    el.btnScan.disabled = !(s.year && s.quarter && s.leagues.length > 0 && s.teams.length > 0);
  }
}

// Devuelve los teams activos según el modo. Step 3 + autoscan lo usan
// para no tener que saber del modo.
function getActiveTeams() {
  // D.3: en modo manual los teams vienen de los filenames parseados.
  // state.manualImagesByTeam ya filtró los unmatched (sólo válidos).
  if (state.imageOrigin === 'manual') {
    return Object.keys(state.manualImagesByTeam).map(k => {
      const [league, team] = k.split('::');
      return { league, team };
    });
  }
  return state.selection.mode === 'advanced'
    ? state.selection.adv.teams
    : state.selection.teams;
}

// Construye el `scope` para los endpoints scoped a partir de state.adv.
// Para cada classic seleccionado incluye TODOS sus sub-productos
// (cacheados al cargar la lista). Esto coincide con la UX simplificada
// de "elegir classic = incluir todo".
function buildScope() {
  const adv = state.selection.adv;
  return {
    seasonal: adv.capsules.map(c => ({
      category: c.category,
      capsule_folder: c.folder_name,
    })),
    classics: adv.classics.map(c => {
      const key = `${c.category}|${c.classic_type}`;
      return {
        category: c.category,
        classic_type: c.classic_type,
        subproducts: state._advSubproductsCache[key] || [],
      };
    }),
  };
}

function setScanState(view, opts = {}) {
  el.scanLoading.style.display = view === 'loading' ? 'block' : 'none';
  el.scanPreview.style.display = view === 'preview' ? 'block' : 'none';
  el.scanError.style.display   = view === 'error'   ? 'block' : 'none';
  if (view === 'error' && opts.message) el.scanErrorBox.textContent = '⚠ ' + opts.message;
  el.btnScan.disabled = view === 'loading';
  el.btnScan.setAttribute('aria-busy', view === 'loading' ? 'true' : 'false');
}

// ─── Auto-scan (debounced) ────────────────────────────────────
// Cuando los 4 selectores están completos, dispara runScan automáticamente
// tras un breve delay. Si el user cambia algo antes, se cancela y reprograma.
let scanTimer = null;
const SCAN_DEBOUNCE_MS = 400;

function cancelAutoScan() {
  if (scanTimer !== null) {
    clearTimeout(scanTimer);
    scanTimer = null;
  }
}

function maybeScheduleAutoScan() {
  cancelAutoScan();
  setScanState('idle');
  state.scanResult = null;
  el.btnNext2.disabled = true;

  // G.2: en modo VLPS NO se hace scan — refreshNext2Button habilita
  // Continuar basado en scope + ligas + tipos directamente.
  if (state.outputMode === 'vlps') {
    refreshNext2Button();
    return;
  }

  // Sólo auto-scan si los selectores del modo activo están completos.
  const s = state.selection;
  let ready = false;
  if (s.mode === 'advanced') {
    const adv = s.adv;
    const hasScope = (adv.capsules.length + adv.classics.length) > 0;
    ready = !!(adv.year && adv.quarter && hasScope && adv.teams.length > 0);
  } else {
    ready = !!(s.year && s.quarter && s.leagues.length > 0 && s.teams.length > 0);
  }
  if (ready) {
    scanTimer = setTimeout(() => {
      scanTimer = null;
      runScan();
    }, SCAN_DEBOUNCE_MS);
  }
}

// ─── localStorage ─────────────────────────────────────────────
const STORAGE_KEY = 'deckbuilder.selection.v2';  // bumped por cambio de shape
const OUTPUT_MODE_STORAGE_KEY = 'deckbuilder.outputMode';  // B.2d
const IMAGE_ORIGIN_STORAGE_KEY = 'deckbuilder.imageOrigin';  // D.2
const MULTI_TEAM_MODE_STORAGE_KEY = 'deckbuilder.multiTeamMode';  // D.4b
const VLPS_FILE_TYPES_STORAGE_KEY = 'deckbuilder.vlpsFileTypes';  // G.2
function persistSelection() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.selection)); } catch (_) {}
}
function persistOutputMode() {
  try { localStorage.setItem(OUTPUT_MODE_STORAGE_KEY, state.outputMode); } catch (_) {}
}
function persistImageOrigin() {
  try { localStorage.setItem(IMAGE_ORIGIN_STORAGE_KEY, state.imageOrigin); } catch (_) {}
}
function persistMultiTeamMode() {
  try { localStorage.setItem(MULTI_TEAM_MODE_STORAGE_KEY, state.multiTeamMode); } catch (_) {}
}
function persistVlpsFileTypes() {
  // Guardamos el valor del radio ("ppt" | "pdf" | "both") en vez del array
  // — más fácil de restaurar y de chequear.
  const radioValue = state.vlpsFileTypes.length === 2
    ? 'both'
    : state.vlpsFileTypes[0] || 'both';
  try { localStorage.setItem(VLPS_FILE_TYPES_STORAGE_KEY, radioValue); } catch (_) {}
}

// ─── Helpers de codificación team value ───────────────────────
// El multi-select de teams usa value = "LEAGUE|TEAM" para tener una key
// única por par.
function encodeTeamValue(league, team) { return `${league}|${team}`; }
function decodeTeamValue(value) {
  const [league, team] = value.split('|');
  return { league, team };
}

// ─── Cargas en cascada (fetch al backend) ─────────────────────
async function loadYears() {
  try {
    const r = await apiFetch('/api/years');
    const { years } = await r.json();
    ssYear.setOptions(years.map(y => ({ value: String(y), label: String(y) })));
  } catch (e) {
    showApiError(`No se pudieron cargar los años: ${e.message}`);
  }
}

async function loadQuarters(year) {
  await quartersLoader(async (signal, isCurrent) => {
    try {
      const r = await apiFetch(`/api/quarters?year=${encodeURIComponent(year)}`, { signal });
      const { quarters } = await r.json();
      if (!isCurrent()) return;
      ssQuarter.setOptions(quarters.map(q => ({ value: q, label: q })));
      ssQuarter.setDisabled(quarters.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar los quarters: ${e.message}`);
    }
  });
}

async function loadLeaguesIntoMultiSelect(year, quarter) {
  await leaguesLoader(async (signal, isCurrent) => {
    try {
      const params = new URLSearchParams({ year, quarter });
      const r = await apiFetch(`/api/leagues?${params}`, { signal });
      const { leagues } = await r.json();
      if (!isCurrent()) return;
      msLeagues.setOptions(leagues.map(l => ({ value: l, label: l })));
      msLeagues.setDisabled(leagues.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar las ligas: ${e.message}`);
      msLeagues.setOptions([]);
    }
  });
}

async function loadTeamsIntoMultiSelect(year, quarter, leagues) {
  if (leagues.length === 0) {
    // Caso vacío: abortamos cualquier fetch en vuelo (bumpea generation)
    // y limpiamos sincrónicamente. Llamar al singleflight con un work
    // no-op cumple ese rol.
    await teamsLoader(async () => {});
    msTeams.setOptions([]);
    msTeams.setDisabled(true);
    return;
  }
  await teamsLoader(async (signal, isCurrent) => {
    try {
      // Fetch de teams por liga, en paralelo. Todos comparten el mismo
      // signal: si el singleflight aborta, las N fetches caen juntas.
      const lists = await Promise.all(leagues.map(async league => {
        const params = new URLSearchParams({ year, quarter, league });
        const r = await apiFetch(`/api/teams?${params}`, { signal });
        const { teams } = await r.json();
        return teams.map(t => ({
          value: encodeTeamValue(league, t),
          label: t,
          subtitle: leagues.length > 1 ? league : null,  // solo mostrar liga si hay multi
        }));
      }));
      if (!isCurrent()) return;
      const allOptions = lists.flat();
      msTeams.setOptions(allOptions);
      msTeams.setDisabled(allOptions.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar los teams: ${e.message}`);
      msTeams.setOptions([]);
      msTeams.setDisabled(true);
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// LOADERS DEL MODO AVANZADO (B.1c)
// ═══════════════════════════════════════════════════════════════
//
// Cada loader es una llamada singleflight que actualiza un selector del
// modo avanzado. La cascada es:
//   categorías  → capsules + classics (en paralelo)
//   capsules/classics → leagues (scoped)
//   leagues     → teams (scoped)
//   teams       → autoscan

async function loadAdvCategories() {
  await advCategoriesLoader(async (signal, isCurrent) => {
    try {
      const r = await apiFetch('/api/categories', { signal });
      const { categories } = await r.json();
      if (!isCurrent()) return;
      advCategories.setOptions(categories.map(c => ({ value: c, label: c })));
      advCategories.setDisabled(categories.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar las categorías: ${e.message}`);
    }
  });
}

// Cápsulas seasonal de todas las categorías seleccionadas, en una sola
// lista con prefijo "[CAT] capsule_key". El value identifica unívocamente
// `{category, folder_name}`.
async function loadAdvCapsules(categories, year, quarter) {
  if (!categories.length || !year || !quarter) {
    advCapsules.setOptions([]);
    advCapsules.setDisabled(true);
    return;
  }
  await advCapsulesLoader(async (signal, isCurrent) => {
    try {
      const perCat = await Promise.all(categories.map(async cat => {
        const params = new URLSearchParams({ year, quarter, category: cat });
        const r = await apiFetch(`/api/capsules?${params}`, { signal });
        const { capsules } = await r.json();
        return capsules.map(c => ({
          value: JSON.stringify({ category: cat, folder_name: c.folder_name }),
          label: c.capsule_key,
          subtitle: cat,
        }));
      }));
      if (!isCurrent()) return;
      const allOptions = perCat.flat();
      advCapsules.setOptions(allOptions);
      advCapsules.setDisabled(allOptions.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar las cápsulas: ${e.message}`);
      advCapsules.setOptions([]);
    }
  });
}

// Classics + pre-fetch de sub-productos por cada classic, para cachear
// en `state._advSubproductsCache` y poder armar el scope sin más network.
async function loadAdvClassics(categories) {
  if (!categories.length) {
    advClassics.setOptions([]);
    advClassics.setDisabled(true);
    return;
  }
  await advClassicsLoader(async (signal, isCurrent) => {
    try {
      // 1) Fetch lista de classics por cada categoría.
      const perCat = await Promise.all(categories.map(async cat => {
        const params = new URLSearchParams({ category: cat });
        const r = await apiFetch(`/api/classics?${params}`, { signal });
        const { classics } = await r.json();
        return classics.map(ct => ({ category: cat, classic_type: ct }));
      }));
      if (!isCurrent()) return;
      const flatItems = perCat.flat();

      // 2) Pre-fetch de sub-productos por cada (category, classic_type)
      //    en paralelo. Cacheamos en state._advSubproductsCache.
      await Promise.all(flatItems.map(async ({ category, classic_type }) => {
        const key = `${category}|${classic_type}`;
        if (state._advSubproductsCache[key]) return;  // ya cacheado
        const params = new URLSearchParams({ category, classic_type });
        const r = await apiFetch(`/api/classic-subproducts?${params}`, { signal });
        const { subproducts } = await r.json();
        if (!isCurrent()) return;
        state._advSubproductsCache[key] = subproducts;
      }));
      if (!isCurrent()) return;

      // 3) Render del multi-select con prefijo de categoría.
      const options = flatItems.map(({ category, classic_type }) => ({
        value: JSON.stringify({ category, classic_type }),
        // Limpiamos el underscore inicial "_CLASSIC ICON" → "CLASSIC ICON".
        label: classic_type.replace(/^_/, ''),
        subtitle: category,
      }));
      advClassics.setOptions(options);
      advClassics.setDisabled(options.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar los classics: ${e.message}`);
      advClassics.setOptions([]);
    }
  });
}

async function loadAdvLeaguesScoped() {
  const adv = state.selection.adv;
  const isVlps = state.outputMode === 'vlps';
  const hasScope = (adv.capsules.length + adv.classics.length) > 0;
  // En VLPS los classics NO requieren quarter (se resuelven por _CLASSICS/...).
  // El quarter solo es obligatorio si hay cápsulas seasonal en el scope (o en
  // modo no-VLPS). Sin esto, un scope classics-only de HEADWEAR (quarter null)
  // salía antes de tiempo y nunca cargaba las ligas → siempre "0 ligas".
  const quarterRequired = !isVlps || adv.capsules.length > 0;
  if (!adv.year || !hasScope || (quarterRequired && !adv.quarter)) {
    advLeagues.setOptions([]);
    advLeagues.setDisabled(true);
    return;
  }
  await advLeaguesLoader(async (signal, isCurrent) => {
    try {
      // En modo VLPS las ligas están en _PPTX/_PDF VLPS, NO en _MERCHBOARDS.
      // Usamos el endpoint correcto según el modo de output.
      const endpoint = isVlps ? '/api/vlps-leagues' : '/api/leagues-scoped';
      const body = { year: adv.year, quarter: adv.quarter, scope: buildScope() };
      if (isVlps) body.file_types = state.vlpsFileTypes;
      const r = await apiFetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        signal,
      });
      const { leagues } = await r.json();
      if (!isCurrent()) return;
      // Si no hay ligas disponibles en el scope, marcar el flag para que
      // la validación no las exija (ej. HEADWEAR no usa _PPTX VLPS/{liga}/).
      state.vlpsNoLeaguesAvailable = leagues.length === 0;
      advLeagues.setOptions(leagues.map(l => ({ value: l, label: l })));
      advLeagues.setDisabled(leagues.length === 0);
      refreshNext2Button();
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar las ligas (scoped): ${e.message}`);
      advLeagues.setOptions([]);
    }
  });
}

async function loadAdvTeamsScoped(leagues) {
  const adv = state.selection.adv;
  if (!leagues.length || !adv.year || !adv.quarter) {
    advTeams.setOptions([]);
    advTeams.setDisabled(true);
    return;
  }
  await advTeamsLoader(async (signal, isCurrent) => {
    try {
      const scope = buildScope();
      // Una llamada scoped por liga (igual que en modo simple).
      const lists = await Promise.all(leagues.map(async league => {
        const r = await apiFetch('/api/teams-scoped', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            year: adv.year,
            quarter: adv.quarter,
            league,
            scope,
          }),
          signal,
        });
        const { teams } = await r.json();
        return teams.map(t => ({
          value: encodeTeamValue(league, t),
          label: t,
          subtitle: leagues.length > 1 ? league : null,
        }));
      }));
      if (!isCurrent()) return;
      const allOptions = lists.flat();
      advTeams.setOptions(allOptions);
      advTeams.setDisabled(allOptions.length === 0);
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      showApiError(`No se pudieron cargar los teams (scoped): ${e.message}`);
      advTeams.setOptions([]);
    }
  });
}


// Single/Multi-select instances (creadas en initStep2)
let ssYear, ssQuarter, msLeagues, msTeams;
let advCategories, advYear, advQuarter, advCapsules, advClassics, advLeagues, advTeams;

function initStep2() {
  // Year (single-select custom)
  ssYear = new SingleSelect(el.ssYear, {
    placeholder: '— Selecciona año —',
    onChange: async (value) => {
      state.selection.year = value || null;
      state.selection.quarter = null;
      state.selection.leagues = [];
      state.selection.teams = [];
      ssQuarter.setOptions([]);
      ssQuarter.setSelected(null);
      ssQuarter.setDisabled(true);
      msLeagues.setOptions([]);
      msTeams.setOptions([]);
      msTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (value) await loadQuarters(value);
    },
  });

  // Quarter (single-select custom)
  ssQuarter = new SingleSelect(el.ssQuarter, {
    placeholder: '— Selecciona quarter —',
    onChange: async (value) => {
      state.selection.quarter = value || null;
      state.selection.leagues = [];
      state.selection.teams = [];
      msLeagues.setOptions([]);
      msTeams.setOptions([]);
      msTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (value && state.selection.year) {
        await loadLeaguesIntoMultiSelect(state.selection.year, value);
      }
    },
  });
  ssQuarter.setDisabled(true);

  // Leagues (multi-select)
  msLeagues = new MultiSelect(el.msLeagues, {
    placeholder: '— Selecciona ligas —',
    onChange: async (values) => {
      state.selection.leagues = values;
      state.selection.teams = [];
      msTeams.setOptions([]);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (values.length > 0 && state.selection.year && state.selection.quarter) {
        await loadTeamsIntoMultiSelect(state.selection.year, state.selection.quarter, values);
      }
    },
  });
  msLeagues.setDisabled(true);

  // Teams (multi-select)
  msTeams = new MultiSelect(el.msTeams, {
    placeholder: '— Selecciona equipos —',
    onChange: (values) => {
      state.selection.teams = values.map(decodeTeamValue);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
    },
  });
  msTeams.setDisabled(true);

  // Botón scan manual (sin debounce)
  el.btnScan.addEventListener('click', runScan);

  // ─── Modo avanzado ────────────────────────────────────────────
  initAdvancedSelectors();
  initModeToggle();
}

// Constantes Q1-Q4 hardcoded para el quarter del modo avanzado (mismo
// dominio que el simple — el server siempre devuelve estos cuatro).
const ADV_QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

function initAdvancedSelectors() {
  // Categories (multi)
  advCategories = new MultiSelect(el.advCategories, {
    placeholder: '— Selecciona categorías —',
    onChange: async (values) => {
      const adv = state.selection.adv;
      adv.categories = values;
      // Cambiar categorías invalida cápsulas/classics/leagues/teams.
      adv.capsules = []; adv.classics = []; adv.leagues = []; adv.teams = [];
      advCapsules.setOptions([]); advCapsules.setDisabled(true);
      advClassics.setOptions([]); advClassics.setDisabled(true);
      advLeagues.setOptions([]); advLeagues.setDisabled(true);
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      // Si ya hay year/quarter, recargar las dependencias.
      if (values.length > 0) {
        await Promise.all([
          loadAdvClassics(values),
          loadAdvCapsules(values, adv.year, adv.quarter),
        ]);
      }
    },
  });

  // Year (single) — usa el global /api/years (los años no dependen del scope).
  advYear = new SingleSelect(el.advYear, {
    placeholder: '— Selecciona año —',
    onChange: async (value) => {
      const adv = state.selection.adv;
      adv.year = value || null;
      adv.capsules = []; adv.leagues = []; adv.teams = [];
      advCapsules.setOptions([]); advCapsules.setDisabled(true);
      advLeagues.setOptions([]); advLeagues.setDisabled(true);
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (value && adv.categories.length > 0 && adv.quarter) {
        await loadAdvCapsules(adv.categories, value, adv.quarter);
      }
    },
  });

  // Quarter (single) — hardcoded Q1-Q4.
  advQuarter = new SingleSelect(el.advQuarter, {
    placeholder: '— Selecciona quarter —',
    onChange: async (value) => {
      const adv = state.selection.adv;
      adv.quarter = value || null;
      adv.capsules = []; adv.leagues = []; adv.teams = [];
      advCapsules.setOptions([]); advCapsules.setDisabled(true);
      advLeagues.setOptions([]); advLeagues.setDisabled(true);
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (value && adv.categories.length > 0 && adv.year) {
        await loadAdvCapsules(adv.categories, adv.year, value);
      }
    },
  });
  advQuarter.setOptions(ADV_QUARTERS.map(q => ({ value: q, label: q })));

  // Capsules (multi). Value = JSON {category, folder_name}.
  advCapsules = new MultiSelect(el.advCapsules, {
    placeholder: '— Selecciona cápsulas —',
    onChange: async (values) => {
      const adv = state.selection.adv;
      adv.capsules = values.map(v => JSON.parse(v));
      adv.leagues = []; adv.teams = [];
      advLeagues.setOptions([]); advLeagues.setDisabled(true);
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      await loadAdvLeaguesScoped();
    },
  });
  advCapsules.setDisabled(true);

  // Classics (multi). Value = JSON {category, classic_type}.
  advClassics = new MultiSelect(el.advClassics, {
    placeholder: '— Classics (opcional) —',
    onChange: async (values) => {
      const adv = state.selection.adv;
      adv.classics = values.map(v => JSON.parse(v));
      adv.leagues = []; adv.teams = [];
      advLeagues.setOptions([]); advLeagues.setDisabled(true);
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      await loadAdvLeaguesScoped();
    },
  });
  advClassics.setDisabled(true);

  // Leagues (multi). Scopeada.
  advLeagues = new MultiSelect(el.advLeagues, {
    placeholder: '— Selecciona ligas —',
    onChange: async (values) => {
      const adv = state.selection.adv;
      adv.leagues = values;
      adv.teams = [];
      advTeams.setOptions([]); advTeams.setDisabled(true);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
      if (values.length > 0) await loadAdvTeamsScoped(values);
    },
  });
  advLeagues.setDisabled(true);

  // Teams (multi). Scopeada — comparte el encode con el modo simple.
  advTeams = new MultiSelect(el.advTeams, {
    placeholder: '— Selecciona equipos —',
    onChange: (values) => {
      state.selection.adv.teams = values.map(decodeTeamValue);
      persistSelection(); refreshScanButton(); maybeScheduleAutoScan();
    },
  });
  advTeams.setDisabled(true);
}

function initModeToggle() {
  el.modeSimple.addEventListener('change', () => switchMode('simple'));
  el.modeAdvanced.addEventListener('change', () => switchMode('advanced'));
}

async function switchMode(mode) {
  if (mode !== 'simple' && mode !== 'advanced') return;
  state.selection.mode = mode;
  el.simpleSelectors.hidden = (mode === 'advanced');
  el.advancedSelectors.hidden = (mode === 'simple');
  el.modeSimple.checked = (mode === 'simple');
  el.modeAdvanced.checked = (mode === 'advanced');
  persistSelection();
  refreshScanButton();
  // Si entramos a advanced y todavía no cargamos categorías ni años, hacerlo ahora.
  if (mode === 'advanced') {
    if (!advCategories.options || advCategories.options.length === 0) {
      await loadAdvCategories();
    }
    if (!advYear.options || advYear.options.length === 0) {
      // Reusamos /api/years global.
      try {
        const r = await apiFetch('/api/years');
        const { years } = await r.json();
        advYear.setOptions(years.map(y => ({ value: String(y), label: String(y) })));
      } catch (_) {}
    }
  }
}

// ─── B.2d: output mode toggle (PPT existente vs PPT nuevo desde cero) ──

function initOutputModeToggle() {
  el.outputModeExisting.addEventListener('change', () => switchOutputMode('existing'));
  el.outputModeBlank.addEventListener('change', () => switchOutputMode('blank'));
  el.outputModeVlps.addEventListener('change', () => switchOutputMode('vlps'));
}

// G.2: sub-toggle de tipos VLPS.
// Al cambiar el tipo (PPT/PDF/Ambos) en modo VLPS también recarga las ligas
// porque distintos tipos pueden tener ligas distintas en el servidor.
function initVlpsFileTypesToggle() {
  if (!el.vlpsFileTypesPpt) return;
  const handler = (radioValue) => () => {
    state.vlpsFileTypes = (radioValue === 'both' ? ['ppt', 'pdf'] : [radioValue]);
    persistVlpsFileTypes();
    refreshNext2Button();
    if (state.outputMode === 'vlps') loadAdvLeaguesScoped();
  };
  el.vlpsFileTypesPpt.addEventListener('change', handler('ppt'));
  el.vlpsFileTypesPdf.addEventListener('change', handler('pdf'));
  el.vlpsFileTypesBoth.addEventListener('change', handler('both'));
}

// ─── C.2: theme toggle (claro / oscuro) ────────────────────────────
//
// El attr `data-theme` ya fue seteado por el script inline en <head>
// (lee localStorage + prefers-color-scheme). Acá solo wireamos el click
// del botón para flipear el attr y persistir.

const THEME_STORAGE_KEY = 'theme';

function initThemeToggle() {
  if (!el.themeToggle) return;
  el.themeToggle.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme') || 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem(THEME_STORAGE_KEY, next); } catch (_) {}
    // Actualizamos el aria-label para feedback de accesibilidad.
    el.themeToggle.setAttribute(
      'aria-label',
      next === 'dark' ? 'Cambiar a tema claro' : 'Cambiar a tema oscuro'
    );
  });
}

/**
 * Switch entre modo "PPT existente" y "PPT nuevo desde cero".
 *
 * - 'existing' → flujo histórico: Step 1 visible, modo Simple/Avanzado libre.
 * - 'blank'    → Step 1 oculto (no hay upload), modo Avanzado forzado,
 *                Simple deshabilitado con tooltip, botón cambia a "PPT nuevo".
 */
async function switchOutputMode(mode) {
  if (!['existing', 'blank', 'vlps'].includes(mode)) return;
  state.outputMode = mode;
  persistOutputMode();

  // Sincroniza el estado visual de los radios.
  el.outputModeExisting.checked = (mode === 'existing');
  el.outputModeBlank.checked = (mode === 'blank');
  el.outputModeVlps.checked = (mode === 'vlps');

  const isBlank = (mode === 'blank');
  const isVlps = (mode === 'vlps');
  // Ambos blank y VLPS ocultan Step 1 (no requieren PPT base).
  const hidesStep1 = isBlank || isVlps;
  // Ambos blank y VLPS fuerzan advanced (no funcionan sin scope).
  const forcesAdvanced = isBlank || isVlps;

  // ── Stepper visual ──
  el.stepper1.hidden = hidesStep1;
  if (el.stepperLines && el.stepperLines[0]) {
    el.stepperLines[0].hidden = hidesStep1;
  }

  // ── Step 1 content ──
  el.step1.hidden = hidesStep1;

  // ── Modo Simple/Avanzado dentro de Step 2 ──
  // En blank y VLPS: Simple deshabilitado + tooltip, Avanzado forzado.
  el.modeSimple.disabled = forcesAdvanced;
  const simpleOptionLabel = el.modeSimple.closest('.mode-option');
  if (simpleOptionLabel) {
    if (forcesAdvanced) {
      const reason = isVlps
        ? 'El modo "Descargar VLPS" requiere selección avanzada — necesita un scope de cápsulas/classics para buscar archivos.'
        : 'El modo "PPT nuevo desde cero" requiere selección avanzada — no hay notas en un PPT base que dicten las cápsulas.';
      simpleOptionLabel.setAttribute('title', reason);
    } else {
      simpleOptionLabel.removeAttribute('title');
    }
  }

  // ── G.2: Toggle Origen (Servidor/Manual) — VLPS solo soporta Servidor ──
  // El user no debería poder elegir Manual en VLPS (no hay "subir archivos
  // VLPS" — son del catálogo).
  if (el.imageOriginToggle) {
    el.imageOriginToggle.hidden = isVlps;
  }
  if (isVlps) {
    // Forzar imageOrigin=server por si venía de manual.
    if (state.imageOrigin === 'manual') {
      switchImageOrigin('server');
    }
  }

  // ── G.2: Row de "Procesamiento" (multi-team) — VLPS no aplica ──
  // En VLPS no hay teams ni decks por team; siempre sale ZIP con archivos.
  if (el.processingModeRow) {
    el.processingModeRow.hidden = isVlps;
  }

  // ── G.2: Row de "Tipo de archivo" (PPT/PDF/Ambos) — solo en VLPS ──
  if (el.vlpsFileTypesRow) {
    el.vlpsFileTypesRow.hidden = !isVlps;
  }

  // ── G.2: Selector de Teams en advanced — VLPS lo oculta ──
  // VLPS busca archivos por liga, no por team.
  const advTeamsField = el.advTeams ? el.advTeams.closest('.field') : null;
  if (advTeamsField) {
    advTeamsField.hidden = isVlps;
  }
  // En simple también — pero VLPS fuerza advanced, así que el simple no
  // se ve. Aún así dejamos esto consistente por si en el futuro VLPS
  // soporta modo simple.
  const msTeamsField = el.msTeams ? el.msTeams.closest('.field') : null;
  if (msTeamsField) {
    msTeamsField.hidden = isVlps;
  }

  // ── Botón generate y atrás ──
  if (el.btnGenerate) {
    if (isVlps) {
      el.btnGenerate.textContent = 'Descargar VLPS →';
    } else if (isBlank) {
      el.btnGenerate.textContent = 'Generar PPT nuevo →';
    } else {
      el.btnGenerate.textContent = 'Generar Deck →';
    }
    // G.2-fix2: si veníamos de un éxito previo y al cambiar modo el botón
    // estaba oculto, lo volvemos a mostrar. resetAll() ya hace lo mismo
    // pero este path puede ejecutarse sin resetear (ej. cambiar modo
    // después de descargar y antes de "Nuevo Deck").
    el.btnGenerate.style.display = '';
  }

  // G.2-fix2: textos del Step 3 + stepper3 + result card + btnDownload
  // se adaptan al modo. Mantenemos sincrónica la copy para que en VLPS no
  // se vean labels de "Generar Deck" / "0 equipos" / "Descargar PPT".
  if (el.stepper3Label) {
    el.stepper3Label.textContent = isVlps ? 'Descargar' : 'Generar';
  }
  if (el.step3Heading) {
    el.step3Heading.textContent = isVlps ? 'Descargar VLPS' : 'Generar Deck';
  }
  if (el.step3Subtitle) {
    if (isVlps) {
      el.step3Subtitle.textContent = 'Bajá los archivos PPTX/PDF pre-armados del catálogo';
    } else if (isBlank) {
      el.step3Subtitle.textContent = 'Genera un PPT nuevo con los merchboards seleccionados';
    } else {
      el.step3Subtitle.textContent = 'Reemplaza los merchboards y descarga el PPT actualizado';
    }
  }
  if (el.resultTitle) {
    el.resultTitle.textContent = isVlps ? '✓ VLPS LISTO' : '✓ DECK LISTO';
  }
  if (el.btnDownload) {
    el.btnDownload.textContent = isVlps ? '⬇ Descargar VLPS' : '⬇ Descargar PPT';
  }

  // "Atrás" desde Step 2 → Step 1 no aplica si Step 1 está oculto.
  if (el.btnBack2) el.btnBack2.hidden = hidesStep1;

  // ── Modo selección: forzar advanced cuando aplica ──
  if (forcesAdvanced && state.selection.mode !== 'advanced') {
    await switchMode('advanced');
  }

  // D.5b: el modo multi-team default depende del contexto. Si el user no
  // eligió explícitamente, recalculamos. En VLPS este modo no aplica pero
  // dejamos que se recalcule (state interno consistente).
  recomputeMultiTeamModeIfImplicit();

  // ── Navegación al entry step del modo ──
  // existing → step 1; blank/vlps → step 2.
  goStep(hidesStep1 ? 2 : 1);

  // Re-evaluamos validación del botón Continuar después de cambiar modo.
  refreshNext2Button();
}

// ─── D.2: image origin toggle + manual upload ──────────────────
//
// Port del parse_merch_name de Python a JS — debe mantenerse en sync
// con `core/pptx_engine.py::parse_merch_name`. Tests del backend usan
// los mismos filenames de input, así que cualquier divergencia se nota.

const MANUAL_GENDER_CODES = new Set(['M', 'W', 'K']);

function parseMerchName(filename) {
  // Quitar la extensión.
  const base = filename.replace(/\.[^.]+$/, '');
  const parts = base.split('_');
  if (parts.length < 3) return null;

  const liga = parts[0].toUpperCase().trim();
  let endIdx = parts.length - 1;

  // Si el último segmento es un número (NNN), descartarlo.
  if (/^\d+$/.test(parts[endIdx])) endIdx -= 1;

  // Si lo que queda termina en M/W/K, es el género.
  let gender = null;
  if (endIdx >= 0 && MANUAL_GENDER_CODES.has(parts[endIdx].toUpperCase())) {
    gender = parts[endIdx].toUpperCase();
    endIdx -= 1;
  }
  if (endIdx < 1) return null;

  const team = parts[1].trim();
  const capsule = parts.slice(2, endIdx + 1).join(' ').toUpperCase().trim();
  if (!capsule) return null;

  const key = capsule + (gender ? ` ${gender}` : '');
  return { liga, team, capsule, gender, key };
}

function teamKeyOf(parsed) {
  // Clave estable para indexar por team en state.manualImagesByTeam.
  return `${parsed.liga}::${parsed.team}`;
}

function initImageOriginToggle() {
  el.imageOriginServer.addEventListener('change', () => switchImageOrigin('server'));
  el.imageOriginManual.addEventListener('change', () => switchImageOrigin('manual'));

  // Drop-zone wiring (similar al PPT drop-zone).
  // Guard: si el click ya fue sobre el <input> nativo (que está dentro
  // del dropzone), el browser ya abrió el picker — no llamamos `.click()`
  // de nuevo o se abriría dos veces.
  el.dropImages.addEventListener('click', (e) => {
    if (e.target === el.inputImages) return;
    el.inputImages.click();
  });
  el.dropImages.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      el.inputImages.click();
    }
  });
  el.dropImages.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.dropImages.classList.add('dragging');
  });
  el.dropImages.addEventListener('dragleave', () => {
    el.dropImages.classList.remove('dragging');
  });
  el.dropImages.addEventListener('drop', (e) => {
    e.preventDefault();
    el.dropImages.classList.remove('dragging');
    if (e.dataTransfer.files && e.dataTransfer.files.length) {
      handleManualImages(e.dataTransfer.files);
    }
  });
  el.inputImages.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length) {
      handleManualImages(e.target.files);
    }
  });

  el.btnClearImages.addEventListener('click', () => {
    clearManualImages();
  });
}

function switchImageOrigin(origin) {
  if (origin !== 'server' && origin !== 'manual') return;
  state.imageOrigin = origin;
  persistImageOrigin();

  el.imageOriginServer.checked = (origin === 'server');
  el.imageOriginManual.checked = (origin === 'manual');

  const isManual = (origin === 'manual');
  el.serverSourceSection.hidden = isManual;
  el.manualSourceSection.hidden = !isManual;

  // D.5b: el modo multi-team default depende del contexto. Si el usuario
  // no eligió explícitamente, recalculamos.
  recomputeMultiTeamModeIfImplicit();

  // Reglas de habilitación del botón Continuar →
  refreshNext2Button();
}

function handleManualImages(fileList) {
  // Acumulamos files (no reemplazamos) — usuario puede ir agregando.
  const files = Array.from(fileList).filter(f => {
    const lower = f.name.toLowerCase();
    return lower.endsWith('.jpg') || lower.endsWith('.jpeg') || lower.endsWith('.png');
  });

  for (const file of files) {
    const parsed = parseMerchName(file.name);
    if (!parsed) {
      state.manualUnmatched.push({ name: file.name });
      continue;
    }
    const key = teamKeyOf(parsed);
    if (!state.manualImagesByTeam[key]) {
      state.manualImagesByTeam[key] = [];
    }
    state.manualImagesByTeam[key].push({ file, parsed });
    state.manualFlatFiles.push(file);
  }

  // Reset del input para que el mismo file se pueda re-seleccionar si lo
  // quitan y vuelven a agregar.
  el.inputImages.value = '';

  renderManualState();
  refreshNext2Button();
}

function clearManualImages() {
  state.manualImagesByTeam = {};
  state.manualUnmatched = [];
  state.manualFlatFiles = [];
  el.inputImages.value = '';
  renderManualState();
  refreshNext2Button();
}

function renderManualState() {
  const teams = Object.keys(state.manualImagesByTeam);
  const totalValid = state.manualFlatFiles.length - state.manualUnmatched.length;
  const hasAny = state.manualFlatFiles.length > 0;

  // Drop-zone state (empty / loaded)
  el.dropImagesEmpty.hidden = hasAny;
  el.dropImagesLoaded.hidden = !hasAny;
  el.btnClearImages.disabled = !hasAny;
  if (hasAny) {
    const totalCount = state.manualFlatFiles.length;
    el.manualImagesCount.textContent = totalCount === 1
      ? '1 archivo'
      : `${totalCount} archivos`;
  }

  // Warning de multi-team — depende del modo (D.4b).
  // En strict: warning porque el modo rechaza multi-team.
  // En mixed/per_team: info-banner explicando qué va a pasar.
  if (teams.length > 1) {
    const teamLabels = teams
      .map(k => k.replace('::', ' / '))
      .join(', ');
    if (state.multiTeamMode === 'strict') {
      el.manualMultiTeamWarning.innerHTML =
        `<strong>⚠ ${teams.length} teams detectados:</strong> ${esc(teamLabels)}. ` +
        `El modo "Un solo equipo" está activo y los rechaza. ` +
        `Cambiá el modo de Procesamiento (Mezclar o Per-team) o quitá los archivos del team no deseado.`;
    } else if (state.multiTeamMode === 'mixed') {
      el.manualMultiTeamWarning.innerHTML =
        `<strong>ℹ ${teams.length} teams detectados:</strong> ${esc(teamLabels)}. ` +
        `Modo "Mezcla" activo — se va a generar 1 PPTX donde slides con la misma nota ` +
        `se llenan en orden con imágenes de los distintos equipos.`;
    } else {  // per_team
      el.manualMultiTeamWarning.innerHTML =
        `<strong>ℹ ${teams.length} teams detectados:</strong> ${esc(teamLabels)}. ` +
        `Modo "Per-team" activo — se va a generar 1 deck por team empaquetados en un ZIP.`;
    }
    el.manualMultiTeamWarning.style.display = 'block';
  } else {
    el.manualMultiTeamWarning.style.display = 'none';
  }

  // Team badge + capsule cards.
  if (teams.length === 1) {
    // Single team: badge + cards de las cápsulas de ese team.
    const key = teams[0];
    const [liga, team] = key.split('::');
    const items = state.manualImagesByTeam[key];

    el.manualTeamBadge.innerHTML =
      `<div class="manual-team-badge">` +
      `<span class="label">EQUIPO</span>` +
      `<span class="value">${esc(team)}</span>` +
      `<span class="label">· ${esc(liga)}</span>` +
      `</div>`;
    el.manualTeamBadge.style.display = 'block';

    const byCapsule = {};
    for (const it of items) {
      if (!byCapsule[it.parsed.key]) byCapsule[it.parsed.key] = 0;
      byCapsule[it.parsed.key] += 1;
    }
    const cards = Object.entries(byCapsule)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([k, n]) =>
        `<div class="manual-capsule-card">` +
        `<h4>${esc(k)}</h4>` +
        `<p>${n} imagen${n === 1 ? '' : 'es'}</p>` +
        `</div>`
      )
      .join('');
    el.manualPreview.innerHTML = `<div class="manual-capsule-grid">${cards}</div>`;
    el.manualPreview.style.display = 'block';
  } else if (teams.length > 1 && state.multiTeamMode !== 'strict') {
    // Multi-team válido: preview adaptado al modo.
    if (state.multiTeamMode === 'mixed') {
      // Mixed: cards agregadas (totales sumando todos los teams).
      const merged = {};
      for (const tk of teams) {
        for (const it of state.manualImagesByTeam[tk]) {
          if (!merged[it.parsed.key]) merged[it.parsed.key] = 0;
          merged[it.parsed.key] += 1;
        }
      }
      el.manualTeamBadge.innerHTML =
        `<div class="manual-team-badge">` +
        `<span class="label">MEZCLA</span>` +
        `<span class="value">${teams.length} EQUIPOS</span>` +
        `</div>`;
      el.manualTeamBadge.style.display = 'block';
      const cards = Object.entries(merged)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([k, n]) =>
          `<div class="manual-capsule-card">` +
          `<h4>${esc(k)}</h4>` +
          `<p>${n} imagen${n === 1 ? '' : 'es'} (sumadas)</p>` +
          `</div>`
        )
        .join('');
      el.manualPreview.innerHTML = `<div class="manual-capsule-grid">${cards}</div>`;
      el.manualPreview.style.display = 'block';
    } else {
      // Per-team: lista de teams con su conteo de imágenes.
      const rows = teams
        .sort()
        .map(tk => {
          const [liga, team] = tk.split('::');
          const items = state.manualImagesByTeam[tk];
          const byCap = {};
          for (const it of items) {
            byCap[it.parsed.key] = (byCap[it.parsed.key] || 0) + 1;
          }
          const capSummary = Object.entries(byCap)
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([k, n]) => `${esc(k)} (${n})`)
            .join(', ');
          return (
            `<div class="manual-capsule-card">` +
            `<h4>${esc(team)} · ${esc(liga)}</h4>` +
            `<p>${items.length} imagen${items.length === 1 ? '' : 'es'} · ${capSummary}</p>` +
            `</div>`
          );
        })
        .join('');
      el.manualTeamBadge.innerHTML =
        `<div class="manual-team-badge">` +
        `<span class="label">PER-TEAM</span>` +
        `<span class="value">${teams.length} DECKS</span>` +
        `</div>`;
      el.manualTeamBadge.style.display = 'block';
      el.manualPreview.innerHTML = `<div class="manual-capsule-grid">${rows}</div>`;
      el.manualPreview.style.display = 'block';
    }
  } else {
    // strict + multi-team OR 0 teams → no mostrar preview de cards
    el.manualTeamBadge.style.display = 'none';
    el.manualPreview.style.display = 'none';
  }

  // Unmatched expander
  if (state.manualUnmatched.length > 0) {
    const list = state.manualUnmatched
      .map(u => `<li>${esc(u.name)}</li>`)
      .join('');
    el.manualUnmatched.innerHTML =
      `<details>` +
      `<summary>⚠ ${state.manualUnmatched.length} archivo${state.manualUnmatched.length === 1 ? '' : 's'} sin naming válido</summary>` +
      `<ul class="unmatched-list">${list}</ul>` +
      `</details>`;
    el.manualUnmatched.style.display = 'block';
  } else {
    el.manualUnmatched.style.display = 'none';
  }
}

function refreshNext2Button() {
  // "Continuar →" del Step 2 enabled cuando:
  // - imageOrigin=server: scan resultado válido (lógica existente — no toco).
  //   Excepción: en outputMode='vlps' no hace falta scan; basta scope+ligas.
  // - imageOrigin=manual: depende del multiTeamMode (D.4b):
  //     · strict:   exactamente 1 team detectado (≥1 imagen).
  //     · mixed:    ≥1 team detectado (≥1 imagen).
  //     · per_team: ≥1 team detectado (≥1 imagen).

  // G.2-fix: en cualquier transición ocultamos el status VLPS por default.
  // El branch VLPS de abajo lo vuelve a mostrar si corresponde.
  if (el.vlpsStatus) el.vlpsStatus.hidden = true;

  // G.2: VLPS tiene su propia regla — no requiere scan ni teams; basta
  // con scope no vacío + al menos 1 liga.
  if (state.outputMode === 'vlps') {
    const adv = state.selection.adv;
    const hasScope = (adv.capsules.length + adv.classics.length) > 0;
    // Liga requerida solo si hay ligas disponibles. Si el servidor devolvió 0
    // ligas (ej. HEADWEAR sin _PPTX VLPS/), se descarga todo sin filtrar por liga.
    const hasLeagues = adv.leagues.length > 0 || state.vlpsNoLeaguesAvailable;
    const hasFileTypes = state.vlpsFileTypes.length > 0;
    el.btnNext2.disabled = !(hasScope && hasLeagues && hasFileTypes);
    renderVlpsStatus();
    return;
  }

  if (state.imageOrigin === 'manual') {
    const teams = Object.keys(state.manualImagesByTeam);
    let hasValid;
    if (state.multiTeamMode === 'strict') {
      hasValid = teams.length === 1;
    } else {
      hasValid = teams.length >= 1;
    }
    el.btnNext2.disabled = !hasValid;
  }
  // En server mode no-VLPS dejamos que el flujo existente maneje btnNext2
  // (vía maybeScheduleAutoScan + scanResult).
}

// ─── G.2-fix: render del status box exclusivo del modo VLPS ──────
//
// En VLPS no hay scan previo (no se conoce cuántos archivos hay hasta que
// el backend gathereea), así que el usuario no recibe ninguna confirmación
// visual entre seleccionar scope/ligas y presionar Continuar. Esta función
// le da feedback claro: qué tiene seleccionado, qué falta, y si ya puede
// avanzar.
//
// Se invoca desde `refreshNext2Button()` que ya dispara en todos los
// cambios relevantes (categoría, año, quarter, cápsula, classic, liga,
// tipo de archivo, switch de outputMode).
function renderVlpsStatus() {
  if (!el.vlpsStatus) return;
  if (state.outputMode !== 'vlps') {
    el.vlpsStatus.hidden = true;
    return;
  }

  const adv = state.selection.adv;
  const nCapsules = adv.capsules.length;
  const nClassics = adv.classics.length;
  const hasScope = (nCapsules + nClassics) > 0;
  const nLeagues = adv.leagues.length;
  const noLeaguesAvailable = state.vlpsNoLeaguesAvailable;
  const hasLeagues = nLeagues > 0 || noLeaguesAvailable;
  const hasFileTypes = state.vlpsFileTypes.length > 0;

  el.vlpsStatus.hidden = false;
  el.vlpsStatus.classList.remove('ready', 'pending');

  // Caso "todo listo" → mensaje verde con resumen.
  if (hasScope && hasLeagues && hasFileTypes) {
    el.vlpsStatus.classList.add('ready');
    const scopeParts = [];
    if (nCapsules > 0) {
      scopeParts.push(`<strong>${nCapsules}</strong> cápsula${nCapsules === 1 ? '' : 's'}`);
    }
    if (nClassics > 0) {
      scopeParts.push(`<strong>${nClassics}</strong> classic${nClassics === 1 ? '' : 's'}`);
    }
    const typesLabel = state.vlpsFileTypes.length === 2
      ? 'PPT + PDF'
      : (state.vlpsFileTypes[0] === 'ppt' ? 'Solo PPT' : 'Solo PDF');
    const ligasLabel = noLeaguesAvailable
      ? 'Todos los archivos'
      : `<strong>${nLeagues}</strong> liga${nLeagues === 1 ? '' : 's'}`;
    el.vlpsStatus.innerHTML =
      `<div class="vlps-status-title">✓ Listo para descargar</div>` +
      `<div class="vlps-status-detail">` +
      `${scopeParts.join(' · ')} · ${ligasLabel} · ${typesLabel}` +
      `</div>`;
    return;
  }

  // Caso pending → listamos qué falta para que el usuario sepa qué hacer.
  el.vlpsStatus.classList.add('pending');
  const missing = [];
  if (!adv.year) missing.push('año');
  if (!adv.quarter) missing.push('quarter');
  if (adv.categories.length === 0) missing.push('categoría');
  if (!hasScope) missing.push('cápsula o classic');
  if (!hasLeagues) missing.push('liga');
  if (!hasFileTypes) missing.push('tipo de archivo');

  el.vlpsStatus.innerHTML =
    `<div class="vlps-status-title">Falta seleccionar</div>` +
    `<div class="vlps-status-detail">${missing.join(' · ')}</div>`;
}

// ─── D.4b / D.5b: multi-team mode toggle (strict / mixed / per_team) ──

function initMultiTeamModeToggle() {
  // explicitUserAction=true porque viene de un click del usuario.
  el.multiTeamStrict.addEventListener('change', () => switchMultiTeamMode('strict', true));
  el.multiTeamMixed.addEventListener('change', () => switchMultiTeamMode('mixed', true));
  el.multiTeamPerTeam.addEventListener('change', () => switchMultiTeamMode('per_team', true));
}

/**
 * D.5b — Default por flow (Opción B): el modo por default depende de
 * (imageOrigin, outputMode), matching el comportamiento histórico:
 *
 *   - manual            → 'strict'  (D.1 default, conservador)
 *   - server + existing → 'per_team' (lo que `generate_batch` hizo siempre)
 *   - server + blank    → 'mixed'   (lo que `generate_blank_batch` hizo siempre)
 *
 * Esto preserva backward-compat: un usuario que nunca tocó el toggle ve
 * exactamente el comportamiento que veía antes de D.5.
 */
function getDefaultMultiTeamMode() {
  if (state.imageOrigin === 'manual') return 'strict';
  if (state.outputMode === 'blank') return 'mixed';
  return 'per_team';  // server + existing
}

/**
 * Recompute del modo cuando el contexto (imageOrigin, outputMode) cambia
 * Y el usuario NO eligió explícitamente. Si eligió → respetamos.
 */
function recomputeMultiTeamModeIfImplicit() {
  if (state.multiTeamModeExplicit) return;
  const computed = getDefaultMultiTeamMode();
  if (computed !== state.multiTeamMode) {
    switchMultiTeamMode(computed, false);  // no marcar como explícito
  }
}

function switchMultiTeamMode(mode, explicitUserAction = false) {
  if (!['strict', 'mixed', 'per_team'].includes(mode)) return;
  state.multiTeamMode = mode;
  if (explicitUserAction) {
    state.multiTeamModeExplicit = true;
    persistMultiTeamMode();
  }

  // Sincroniza el estado visual de los radios.
  el.multiTeamStrict.checked = (mode === 'strict');
  el.multiTeamMixed.checked = (mode === 'mixed');
  el.multiTeamPerTeam.checked = (mode === 'per_team');

  // Re-renderear preview porque el modo cambia qué warning se muestra,
  // qué texto del badge, etc.
  renderManualState();
  refreshNext2Button();
}

async function loadSelection() {
  // Restaura selección persistida y dispara los fetches en cascada.
  let saved;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    saved = JSON.parse(raw);
  } catch (_) { return; }

  // ─── Modo simple ─────────────────────────────────────────────
  if (saved.year) {
    state.selection.year = String(saved.year);
    ssYear.setSelected(String(saved.year));
    await loadQuarters(saved.year);
  }
  if (saved.quarter) {
    state.selection.quarter = saved.quarter;
    ssQuarter.setSelected(saved.quarter);
    if (state.selection.year) {
      await loadLeaguesIntoMultiSelect(state.selection.year, saved.quarter);
    }
  }
  if (Array.isArray(saved.leagues) && saved.leagues.length > 0) {
    state.selection.leagues = saved.leagues;
    msLeagues.setSelected(saved.leagues);
    if (state.selection.year && state.selection.quarter) {
      await loadTeamsIntoMultiSelect(state.selection.year, state.selection.quarter, saved.leagues);
    }
  }
  if (Array.isArray(saved.teams) && saved.teams.length > 0) {
    state.selection.teams = saved.teams;
    msTeams.setSelected(saved.teams.map(t => encodeTeamValue(t.league, t.team)));
  }

  // ─── Modo + estado avanzado ──────────────────────────────────
  // Si el usuario estaba en modo avanzado, restauramos los selectores
  // en cascada igual que el simple. Si algún path falla, el modo simple
  // ya cargó arriba y queda como fallback usable.
  if (saved.mode === 'advanced' && saved.adv) {
    await switchMode('advanced');  // muestra la sección y carga categorías/años
    const adv = saved.adv;

    if (Array.isArray(adv.categories) && adv.categories.length > 0) {
      state.selection.adv.categories = adv.categories;
      advCategories.setSelected(adv.categories);
    }
    if (adv.year) {
      state.selection.adv.year = String(adv.year);
      advYear.setSelected(String(adv.year));
    }
    if (adv.quarter) {
      state.selection.adv.quarter = adv.quarter;
      advQuarter.setSelected(adv.quarter);
    }
    // Re-load capsules + classics si hay categorías + año + quarter
    if (adv.categories?.length && adv.year && adv.quarter) {
      await Promise.all([
        loadAdvClassics(adv.categories),
        loadAdvCapsules(adv.categories, adv.year, adv.quarter),
      ]);
    }
    if (Array.isArray(adv.capsules) && adv.capsules.length > 0) {
      state.selection.adv.capsules = adv.capsules;
      advCapsules.setSelected(adv.capsules.map(c => JSON.stringify({
        category: c.category, folder_name: c.folder_name,
      })));
    }
    if (Array.isArray(adv.classics) && adv.classics.length > 0) {
      state.selection.adv.classics = adv.classics;
      advClassics.setSelected(adv.classics.map(c => JSON.stringify({
        category: c.category, classic_type: c.classic_type,
      })));
    }
    if ((adv.capsules?.length || 0) + (adv.classics?.length || 0) > 0) {
      await loadAdvLeaguesScoped();
    }
    if (Array.isArray(adv.leagues) && adv.leagues.length > 0) {
      state.selection.adv.leagues = adv.leagues;
      advLeagues.setSelected(adv.leagues);
      await loadAdvTeamsScoped(adv.leagues);
    }
    if (Array.isArray(adv.teams) && adv.teams.length > 0) {
      state.selection.adv.teams = adv.teams;
      advTeams.setSelected(adv.teams.map(t => encodeTeamValue(t.league, t.team)));
    }
  }

  refreshScanButton();
}

async function runScan() {
  cancelAutoScan();
  const s = state.selection;
  setScanState('loading');
  el.btnNext2.disabled = true;

  // Construir el body del scan según el modo.
  let url, body;
  if (s.mode === 'advanced') {
    const adv = s.adv;
    url = '/api/scan-catalog-scoped';
    body = {
      year: adv.year,
      quarter: adv.quarter,
      selections: adv.teams,
      scope: buildScope(),
    };
  } else {
    url = '/api/scan-catalog';
    body = {
      year: s.year,
      quarter: s.quarter,
      selections: s.teams,
    };
  }

  await scanLoader(async (signal, isCurrent) => {
    try {
      const r = await apiFetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        signal,
      });
      const list = await r.json();
      if (!isCurrent()) return;
      if (!list || list.length === 0) throw new Error('Respuesta vacía del servidor');
      state.scanResult = list;
      renderScanPreview(list);
      setScanState('preview');

      const totalImgs = list.reduce((acc, ts) =>
        acc + GENDERS.reduce((a, g) =>
          a + (ts.by_gender[g] || []).reduce((b, c) => b + c.image_count, 0), 0), 0);
      el.btnNext2.disabled = totalImgs === 0;
    } catch (e) {
      if (e.name === 'AbortError') throw e;
      setScanState('error', { message: `Error consultando el servidor: ${e.message}` });
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// renderScanPreview — dispatcher: single → tabs, multi → tabla
// ═══════════════════════════════════════════════════════════════
function renderScanPreview(teamScans) {
  if (teamScans.length === 1) {
    renderScanPreviewSingle(teamScans[0]);
  } else {
    renderScanPreviewMulti(teamScans);
  }
}

function renderScanPreviewSingle(teamScan) {
  const s = state.selection;
  const byG = teamScan.by_gender || {};

  let totalCapsules = 0, totalImgs = 0;
  for (const g of GENDERS) for (const c of (byG[g] || [])) {
    if (c.image_count > 0) totalCapsules++;
    totalImgs += c.image_count;
  }
  el.scanSummary.innerHTML = `
    <h4>${esc(teamScan.league)} · ${esc(teamScan.team)} · ${esc(s.year)} ${esc(s.quarter)}</h4>
    <p>${totalCapsules} cápsulas con contenido · ${totalImgs} imágenes encontradas</p>
  `;

  const allCapsules = new Set();
  for (const g of GENDERS) for (const c of (byG[g] || [])) allCapsules.add(c.capsule);
  const capsuleList = [...allCapsules].sort();

  const totalsByGender = {};
  for (const g of GENDERS) {
    totalsByGender[g] = (byG[g] || []).reduce((a, c) => a + c.image_count, 0);
  }

  const tabsHtml = GENDERS.map(g => {
    const total = totalsByGender[g];
    const badge = total > 0 ? `${total} img` : 'sin contenido';
    return `<button type="button" class="gender-tab" data-gender="${g}"
                    role="tab" aria-selected="false" aria-controls="gender-panel"
                    id="tab-${g}">
              <span class="tab-label">${g}</span>
              <span class="tab-badge">${badge}</span>
            </button>`;
  }).join('');

  el.scanTree.innerHTML = `
    <div class="gender-tabs" role="tablist" aria-label="Disponibilidad por género">${tabsHtml}</div>
    <div class="gender-panel" id="gender-panel" role="tabpanel" tabindex="0"></div>
  `;

  function renderGenderPanel(g) {
    const panel = el.scanTree.querySelector('.gender-panel');
    const caps = byG[g] || [];
    const map = new Map(caps.map(c => [c.capsule, c]));
    const capsCnt = caps.filter(c => c.image_count > 0).length;
    const totalG = totalsByGender[g];
    const rows = capsuleList.map(cap => {
      const c = map.get(cap);
      if (!c) return `<div class="capsule-row miss"><span class="icon">✗</span><span class="name">${esc(cap)}</span><span class="count">no existe en este género</span></div>`;
      if (c.image_count === 0) return `<div class="capsule-row warn"><span class="icon">⚠</span><span class="name">${esc(cap)}</span><span class="count">carpeta vacía</span></div>`;
      return `<div class="capsule-row ok"><span class="icon">✓</span><span class="name">${esc(cap)}</span><span class="count">${c.image_count} img</span></div>`;
    }).join('');
    panel.innerHTML = `
      <div class="gender-panel-meta">${capsCnt} cápsulas con contenido · ${totalG} imagen(es)</div>
      ${rows}
    `;
    panel.setAttribute('aria-labelledby', `tab-${g}`);
  }

  function activateTab(g) {
    el.scanTree.querySelectorAll('.gender-tab').forEach(t => {
      const active = t.dataset.gender === g;
      t.classList.toggle('active', active);
      t.setAttribute('aria-selected', active ? 'true' : 'false');
      t.tabIndex = active ? 0 : -1;
    });
    renderGenderPanel(g);
  }

  el.scanTree.querySelectorAll('.gender-tab').forEach(tab => {
    tab.addEventListener('click', () => activateTab(tab.dataset.gender));
    tab.addEventListener('keydown', e => {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      const idx = GENDERS.indexOf(tab.dataset.gender);
      const next = e.key === 'ArrowRight'
        ? (idx + 1) % GENDERS.length
        : (idx - 1 + GENDERS.length) % GENDERS.length;
      activateTab(GENDERS[next]);
      el.scanTree.querySelector(`#tab-${GENDERS[next]}`).focus();
    });
  });

  const firstWithContent = GENDERS.find(g => totalsByGender[g] > 0) || 'MENS';
  activateTab(firstWithContent);
}

function renderScanPreviewMulti(teamScans) {
  const s = state.selection;

  // Totales globales
  let totalDecks = 0, totalImgs = 0;
  for (const ts of teamScans) {
    const teamTotal = GENDERS.reduce((a, g) =>
      a + (ts.by_gender[g] || []).reduce((b, c) => b + c.image_count, 0), 0);
    if (teamTotal > 0) totalDecks++;
    totalImgs += teamTotal;
  }

  el.scanSummary.innerHTML = `
    <h4>${teamScans.length} EQUIPOS · ${esc(s.year)} ${esc(s.quarter)}</h4>
    <p>${totalDecks} con contenido · ${totalImgs} imágenes totales</p>
  `;

  // Filas: una de overview + una de detalle (hidden) por team
  const rows = teamScans.map((ts, idx) => {
    const byG = ts.by_gender || {};
    const counts = {};
    let teamTotal = 0;
    for (const g of GENDERS) {
      counts[g] = (byG[g] || []).reduce((a, c) => a + c.image_count, 0);
      teamTotal += counts[g];
    }
    let statusClass, statusIcon;
    if (teamTotal === 0) {
      statusClass = 'miss'; statusIcon = '✗';
    } else if (counts.MENS === 0 || counts.WOMENS === 0 || counts.KIDS === 0) {
      statusClass = 'warn'; statusIcon = '⚠';
    } else {
      statusClass = 'ok'; statusIcon = '✓';
    }
    return `
      <tr class="team-row" data-team-idx="${idx}">
        <td class="expand-cell"><span class="expand-icon">▸</span></td>
        <td>
          <div class="team-cell">
            <span>${esc(ts.team)}</span>
            <span class="team-league">${esc(ts.league)}</span>
          </div>
        </td>
        <td class="num">${counts.MENS}</td>
        <td class="num">${counts.WOMENS}</td>
        <td class="num">${counts.KIDS}</td>
        <td class="num"><strong>${teamTotal}</strong></td>
        <td class="status-cell ${statusClass}">${statusIcon}</td>
      </tr>
      <tr class="team-detail-row" data-team-idx="${idx}" hidden>
        <td colspan="7">${renderTeamDetailGrid(ts)}</td>
      </tr>
    `;
  }).join('');

  const scrollClass = teamScans.length >= 10 ? ' scrollable' : '';
  el.scanTree.innerHTML = `
    <div class="scan-table-wrap${scrollClass}">
      <table class="scan-table">
        <thead>
          <tr>
            <th class="expand-cell" aria-hidden="true"></th>
            <th>Equipo</th>
            <th class="num">MENS</th>
            <th class="num">WOMENS</th>
            <th class="num">KIDS</th>
            <th class="num">Total</th>
            <th>Estado</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  // Click handler: toggle de la fila de detalle correspondiente
  el.scanTree.querySelectorAll('.team-row').forEach(row => {
    row.addEventListener('click', () => {
      const idx = row.dataset.teamIdx;
      const detail = el.scanTree.querySelector(`.team-detail-row[data-team-idx="${idx}"]`);
      const isOpen = !detail.hidden;
      detail.hidden = isOpen;
      row.classList.toggle('expanded', !isOpen);
    });
  });
}

// Renderiza la grilla de detalle por cápsula × género para un team.
function renderTeamDetailGrid(teamScan) {
  const byG = teamScan.by_gender || {};

  // Union de cápsulas presentes en cualquier género de este team
  const allCapsules = new Set();
  for (const g of GENDERS) for (const c of (byG[g] || [])) allCapsules.add(c.capsule);
  const capsuleList = [...allCapsules].sort();

  if (capsuleList.length === 0) {
    return `<div class="capsule-detail-empty">Sin cápsulas en este team</div>`;
  }

  const cols = GENDERS.map(g => {
    const map = new Map((byG[g] || []).map(c => [c.capsule, c]));
    const rows = capsuleList.map(cap => {
      const c = map.get(cap);
      if (!c) return `<div class="capsule-row miss"><span class="icon">✗</span><span class="name">${esc(cap)}</span></div>`;
      if (c.image_count === 0) return `<div class="capsule-row warn"><span class="icon">⚠</span><span class="name">${esc(cap)}</span></div>`;
      return `<div class="capsule-row ok"><span class="icon">✓</span><span class="name">${esc(cap)}</span><span class="count">${c.image_count}</span></div>`;
    }).join('');
    return `
      <div class="capsule-col">
        <div class="capsule-col-header">${g}</div>
        ${rows}
      </div>
    `;
  }).join('');

  return `<div class="capsule-detail-grid">${cols}</div>`;
}

// ═══════════════════════════════════════════════════════════════
// STEP 3: GENERATE (todo el trabajo lo hace el backend)
// ═══════════════════════════════════════════════════════════════
function goStep3Setup() {
  const s = state.selection;
  const teams = getActiveTeams();
  const n = teams.length;
  const teamsLabel = n === 1
    ? `${esc(teams[0].league)} · ${esc(teams[0].team)}`
    : `${n} equipos`;
  const isBlank = state.outputMode === 'blank';
  const isVlps = state.outputMode === 'vlps';
  const isManual = state.imageOrigin === 'manual';

  // ── G.2: modo VLPS — resumen propio (no aplica el resto) ──
  if (isVlps) {
    const adv = s.adv;
    const nLeagues = adv.leagues.length;
    const ligasLabel = nLeagues === 1 ? esc(adv.leagues[0]) : `${nLeagues} ligas`;
    const nCapsules = adv.capsules.length;
    const nClassics = adv.classics.length;
    const scopeParts = [];
    if (nCapsules > 0) {
      scopeParts.push(`<span style="color:var(--accent)">${nCapsules}</span> cápsula${nCapsules === 1 ? '' : 's'}`);
    }
    if (nClassics > 0) {
      scopeParts.push(`<span style="color:var(--accent)">${nClassics}</span> classic${nClassics === 1 ? '' : 's'}`);
    }
    const typesLabel = state.vlpsFileTypes.length === 2
      ? 'PPT + PDF'
      : (state.vlpsFileTypes[0] === 'ppt' ? 'Solo PPT' : 'Solo PDF');
    el.preGenInfo.innerHTML = `
      <strong style="color:var(--text)">Resumen (Descargar VLPS):</strong><br>
      Período: <span style="color:var(--accent)">${esc(adv.year)} ${esc(adv.quarter)}</span><br>
      Scope: ${scopeParts.join(' · ') || '<em>vacío</em>'}<br>
      Ligas: <span style="color:var(--accent)">${ligasLabel}</span><br>
      Tipo de archivo: <span style="color:var(--accent)">${typesLabel}</span>
    `;
    return;
  }

  // Etiqueta del modo combinado (output × origen × multi-team mode).
  let modeLabel = '';
  if (isManual && isBlank) modeLabel = ' (PPT nuevo · upload manual)';
  else if (isManual)       modeLabel = ' (PPT existente · upload manual)';
  else if (isBlank)        modeLabel = ' (PPT nuevo desde cero)';
  else if (s.mode === 'advanced') modeLabel = ' (modo avanzado)';
  // D.4b: en manual con multi-team, añadir sub-modo para claridad.
  if (isManual) {
    const nTeams = teams.length;
    if (state.multiTeamMode === 'mixed' && nTeams > 1) {
      modeLabel = modeLabel.replace(')', ' · mezcla)');
    } else if (state.multiTeamMode === 'per_team' && nTeams > 1) {
      modeLabel = modeLabel.replace(')', ' · per-team ZIP)');
    }
  }

  // Construir el cuerpo del resumen según el origen.
  let body;
  if (isManual) {
    // D.3: en modo manual el período + scope no aplican. El team viene
    // del filename. Mostramos el conteo de imágenes + cápsulas detectadas.
    let totalImages = 0;
    const capsules = new Set();
    for (const tk in state.manualImagesByTeam) {
      for (const item of state.manualImagesByTeam[tk]) {
        totalImages += 1;
        capsules.add(item.parsed.key);
      }
    }
    const capsulasLabel = capsules.size === 1 ? 'cápsula' : 'cápsulas';
    const imgsLabel = totalImages === 1 ? 'imagen' : 'imágenes';
    body =
      (isBlank
        ? ''
        : `PPT base: <span style="color:var(--accent)">${esc(state.pptxFile?.name) || '—'}</span><br>`
      ) +
      `Equipo: <span style="color:var(--accent)">${teamsLabel}</span><br>` +
      `Imágenes: <span style="color:var(--accent)">${totalImages}</span> ${imgsLabel} ` +
      `en <span style="color:var(--accent)">${capsules.size}</span> ${capsulasLabel}`;
  } else {
    // Server mode (server): comportamiento histórico — período + scope + equipos.
    const year = s.mode === 'advanced' ? s.adv.year : s.year;
    const quarter = s.mode === 'advanced' ? s.adv.quarter : s.quarter;

    // B.1d.3: en modo avanzado mostramos el scope que se va a aplicar
    // para que el usuario confirme antes de generar.
    let scopeLine = '';
    if (s.mode === 'advanced') {
      const adv = s.adv;
      const parts = [];
      if (adv.categories.length > 0) {
        const cats = adv.categories.length === 1
          ? esc(adv.categories[0])
          : `${adv.categories.length} categorías`;
        parts.push(`<span style="color:var(--accent)">${cats}</span>`);
      }
      if (adv.capsules.length > 0) {
        parts.push(
          `<span style="color:var(--accent)">${adv.capsules.length}</span> ` +
          (adv.capsules.length === 1 ? 'cápsula' : 'cápsulas')
        );
      }
      if (adv.classics.length > 0) {
        parts.push(
          `<span style="color:var(--accent)">${adv.classics.length}</span> ` +
          (adv.classics.length === 1 ? 'classic' : 'classics')
        );
      }
      if (parts.length > 0) {
        scopeLine = `Filtrado por: ${parts.join(' · ')}<br>`;
      }
    }

    body =
      (isBlank
        ? ''
        : `PPT base: <span style="color:var(--accent)">${esc(state.pptxFile?.name) || '—'}</span><br>`
      ) +
      `Período: <span style="color:var(--accent)">${esc(year)} ${esc(quarter)}</span><br>` +
      `${scopeLine}${n === 1 ? 'Equipo' : 'Equipos'}: <span style="color:var(--accent)">${teamsLabel}</span>`;
  }

  el.preGenInfo.innerHTML = `
    <strong style="color:var(--text)">Resumen${modeLabel}:</strong><br>
    ${body}
  `;
}

function goStep(n) {
  // B.2d: en blank no existe Step 1 (no hay upload). Si algo pide ir
  // ahí (back nav, init, reset), redirigimos a Step 2.
  if (n === 1 && state.outputMode === 'blank') {
    n = 2;
  }
  [1, 2, 3].forEach(i => {
    const stepEl = el.steps[i];
    const stepperEl = el.steppers[i];

    // Visibilidad del content area (solo el activo se muestra)
    stepEl.classList.toggle('active', i === n);

    // Estado visual del stepper item
    stepperEl.classList.remove('active', 'done');
    stepperEl.removeAttribute('aria-current');
    if (i === n) {
      stepperEl.classList.add('active');
      stepperEl.setAttribute('aria-current', 'step');
      stepperEl.disabled = false;
    } else if (i < n) {
      stepperEl.classList.add('done');
      stepperEl.disabled = false;
    } else {
      // Pending — bloqueado para forward navigation
      stepperEl.disabled = true;
    }
  });

  if (n === 3) goStep3Setup();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Click en stepper: solo back-navigation (steps con clase 'done').
function initStepper() {
  [1, 2, 3].forEach(i => {
    el.steppers[i].addEventListener('click', () => {
      if (el.steppers[i].classList.contains('done')) {
        goStep(i);
      }
    });
  });
}

function log(msg, type = '') {
  el.logBox.style.display = 'block';
  const line = document.createElement('div');
  line.className = type;
  line.textContent = msg;
  el.logBox.appendChild(line);
  el.logBox.scrollTop = el.logBox.scrollHeight;
}

function setProgress(pct) {
  el.progressFill.style.width = pct + '%';
}

// ─── Modal de confirmación ────────────────────────────────────
const CONFIRM_THRESHOLD = 10;  // ≥ N teams → preguntar antes de generar

function showConfirm({ title, message, onConfirm }) {
  el.confirmTitle.textContent = title;
  el.confirmMessage.textContent = message;
  el.confirmModal.hidden = false;
  const cleanup = () => {
    el.confirmModal.hidden = true;
    el.confirmOk.onclick = null;
    el.confirmCancel.onclick = null;
    document.removeEventListener('keydown', escListener);
  };
  const escListener = (e) => { if (e.key === 'Escape') cleanup(); };
  el.confirmOk.onclick = () => { cleanup(); onConfirm(); };
  el.confirmCancel.onclick = cleanup;
  document.addEventListener('keydown', escListener);
  // Foco inicial en cancelar (opción segura por default)
  setTimeout(() => el.confirmCancel.focus(), 0);
}

async function generateDeck() {
  const isBlank = state.outputMode === 'blank';
  const isVlps = state.outputMode === 'vlps';
  const isManual = state.imageOrigin === 'manual';

  // G.2: VLPS tiene sus propios requisitos — scope no vacío + ≥1 liga.
  // No requiere PPT base ni teams.
  if (isVlps) {
    const adv = state.selection.adv;
    const hasScope = (adv.capsules.length + adv.classics.length) > 0;
    // Liga requerida solo si hay ligas disponibles. Si el servidor devolvió 0
    // ligas (ej. HEADWEAR sin _PPTX VLPS/{liga}/), se descarga todo sin filtro.
    const hasLeagues = adv.leagues.length > 0 || state.vlpsNoLeaguesAvailable;
    if (!hasScope || !hasLeagues) return;
    if (state.vlpsFileTypes.length === 0) return;
    doGenerate();
    return;
  }

  // En modo "PPT existente" se requiere archivo PPT subido (sirve para
  // tanto server como manual). En modo "PPT nuevo desde cero" NO hace falta.
  if (!isBlank && !state.pptxFile) return;
  // En manual mode tiene que haber al menos 1 team detectado (con ≥1 imagen).
  if (isManual) {
    const nManualTeams = Object.keys(state.manualImagesByTeam).length;
    if (state.multiTeamMode === 'strict' && nManualTeams !== 1) return;
    if (state.multiTeamMode !== 'strict' && nManualTeams < 1) return;
  }
  const n = getActiveTeams().length;
  if (n === 0) return;

  // Confirmación para batches grandes
  if (n >= CONFIRM_THRESHOLD) {
    showConfirm({
      title: 'Confirmar generación',
      message: `Vas a generar ${n} decks. Esto puede tardar varios minutos. ¿Continuar?`,
      onConfirm: () => doGenerate(),
    });
  } else {
    doGenerate();
  }
}

// ─── Helper: formatear duración humana ─────────────────────────
function formatElapsed(ms) {
  const total = Math.floor(ms / 1000);
  const m = Math.floor(total / 60);
  const s = total % 60;
  return m > 0 ? `${m}m ${s.toString().padStart(2, '0')}s` : `${s}s`;
}

// ─── Helpers de generación basada en jobs ──────────────────────

const PHASE_LABELS = {
  scan:  'escaneando catálogo',
  fetch: 'descargando imágenes',
  pptx:  'armando deck',
  zip:   'comprimiendo',
};

const PHASE_FRACTIONS = {
  scan:  0.05,
  fetch: 0.40,
  pptx:  0.85,
  zip:   0.97,
};

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function computeProgressPct(job) {
  // Mapea el estado del job a un porcentaje 0-100 para la barra.
  // - Cada team completado vale 100/total.
  // - Dentro del team actual, la fase aporta su fracción.
  const total = Math.max(1, job.teams_total || 1);
  const doneFrac = (job.teams_done || 0) / total;
  const phaseFrac = PHASE_FRACTIONS[job.current_phase] || 0;
  const inProgress = (job.current_team_index && job.current_team_index > (job.teams_done || 0))
    ? phaseFrac / total
    : 0;
  return Math.max(8, Math.min(95, (doneFrac + inProgress) * 100));
}

function buildStatusMsg(job) {
  if (job.status === 'queued') {
    return 'En cola, esperando worker...';
  }
  if (job.status === 'running') {
    const phaseTxt = PHASE_LABELS[job.current_phase] || 'procesando';
    const idx = job.current_team_index || 0;
    const total = job.teams_total || 1;
    if (total === 1) {
      return `${job.current_team || '...'} — ${phaseTxt}`;
    }
    return `${idx}/${total} · ${job.current_team || '...'} · ${phaseTxt}`;
  }
  return job.status;
}

function renderTeamList(listEl, job, selectedTeams) {
  // Reconstruye la lista en cada update — N es chico (≤25) así que el
  // costo es trivial vs la complejidad de diffing.
  const done = job.per_team_results || [];
  const doneByKey = {};
  for (const r of done) doneByKey[`${r.league}/${r.team}`] = r;

  const html = selectedTeams.map((sel, i) => {
    const idx = i + 1;
    const key = `${sel.league}/${sel.team}`;
    const r = doneByKey[key];

    if (r) {
      if (r.success) {
        const dur = formatElapsed(r.duration_seconds * 1000);
        const slides = r.replaced_count != null
          ? ` · ${r.replaced_count} slides`
          : '';
        return `<li class="done"><span class="idx">${idx}</span><span class="icon">✓</span><span class="name">${esc(sel.team)} · ${esc(sel.league)}</span><span class="meta">${dur}${slides}</span></li>`;
      } else {
        return `<li class="failed"><span class="idx">${idx}</span><span class="icon">✗</span><span class="name">${esc(sel.team)} · ${esc(sel.league)}</span><span class="meta">${esc(r.error || 'falló')}</span></li>`;
      }
    }
    if (idx === job.current_team_index) {
      const phase = PHASE_LABELS[job.current_phase] || 'procesando';
      return `<li class="running"><span class="idx">${idx}</span><span class="icon">▶</span><span class="name">${esc(sel.team)} · ${esc(sel.league)}</span><span class="meta">${phase}</span></li>`;
    }
    return `<li class="pending"><span class="idx">${idx}</span><span class="icon">·</span><span class="name">${esc(sel.team)} · ${esc(sel.league)}</span><span class="meta">esperando</span></li>`;
  }).join('');
  listEl.innerHTML = html;
}

const POLL_INTERVAL_MS = 2000;
const POLL_MAX_INTERVAL_MS = 30000;
const POLL_MAX_CONSECUTIVE_ERRORS = 5;

async function doGenerate() {
  el.btnGenerate.disabled = true;
  el.logBox.innerHTML = '';
  el.resultCard.classList.remove('show');

  // ─── Setup del UI: warning + status + lista colapsable ──
  el.logBox.style.display = 'block';

  const warningEl = document.createElement('div');
  warningEl.className = 'log-warning';
  warningEl.textContent = '⚠ No cierres esta pestaña. Para PPTs grandes la generación puede tardar varios minutos.';
  el.logBox.appendChild(warningEl);

  const statusEl = document.createElement('div');
  statusEl.className = 'log-status';
  statusEl.textContent = '⏱ 0s — Subiendo PPT al servidor...';
  el.logBox.appendChild(statusEl);

  const selectedTeams = [...getActiveTeams()];

  const detailsEl = document.createElement('details');
  detailsEl.className = 'log-team-progress';
  detailsEl.innerHTML = `
    <summary class="team-progress-summary">0/${selectedTeams.length} deck${selectedTeams.length === 1 ? '' : 's'} generados</summary>
    <ul class="team-progress-list"></ul>
  `;
  // Multi-team: expandido por default. Single-team: oculto (1 línea no aporta).
  if (selectedTeams.length > 1) {
    detailsEl.open = true;
  } else {
    detailsEl.hidden = true;
  }
  el.logBox.appendChild(detailsEl);
  const summaryEl = detailsEl.querySelector('.team-progress-summary');
  const teamListEl = detailsEl.querySelector('.team-progress-list');

  const t_start = performance.now();
  let currentMsg = 'Subiendo PPT al servidor...';

  // Ticker local: actualiza el timer cada segundo aunque el poll esté
  // esperando (UX sigue "viva" entre polls).
  const localTicker = setInterval(() => {
    const elapsed = performance.now() - t_start;
    statusEl.textContent = `⏱ ${formatElapsed(elapsed)} — ${currentMsg}`;
  }, 1000);

  try {
    // ═══ FASE 1: Upload + crear job ═══
    const s = state.selection;
    const activeTeams = getActiveTeams();
    const isBlank = state.outputMode === 'blank';
    // G.2-fix3: faltaba esta declaración — las ramas VLPS de la Fase 3
    // (resultText + fallback filename + currentMsg de download) hacían
    // referencia a `isVlps` y se rompían con ReferenceError tras el
    // polling, perdiendo el output del job exitoso del backend.
    const isVlps = state.outputMode === 'vlps';
    const isManual = state.imageOrigin === 'manual';

    // year/quarter sólo aplican en server mode. Los declaramos en este
    // scope (no dentro del else) para que el fallback de filename de la
    // fase 3 de abajo los pueda referenciar sin ReferenceError.
    let year = null;
    let quarter = null;
    if (!isManual) {
      year = s.mode === 'advanced' ? s.adv.year : s.year;
      quarter = s.mode === 'advanced' ? s.adv.quarter : s.quarter;
    }

    const form = new FormData();
    let endpoint;
    let useJsonBody = false;
    let jsonBody = null;

    // G.2: modo VLPS — JSON body al endpoint /api/gather-vlps.
    if (state.outputMode === 'vlps') {
      endpoint = '/api/gather-vlps';
      useJsonBody = true;
      const adv = state.selection.adv;
      jsonBody = {
        year: adv.year,
        quarter: adv.quarter,
        leagues: adv.leagues,
        scope: buildScope(),
        file_types: state.vlpsFileTypes,
      };
      currentMsg = 'Buscando archivos VLPS en el servidor...';
      statusEl.textContent = `⏱ 0s — ${currentMsg}`;
    } else if (isManual) {
      // D.3 — modo manual: POST a /api/generate-from-uploads con los
      // archivos válidos parseados. No mandamos selections ni scope ni
      // year/quarter — el endpoint los deriva del filename de cada imagen.
      endpoint = '/api/generate-from-uploads';

      // Adjuntar SOLO los archivos válidos (state.manualImagesByTeam
      // ya filtró los unmatched). Multi-team enforcement viene del backend.
      for (const teamKey in state.manualImagesByTeam) {
        for (const item of state.manualImagesByTeam[teamKey]) {
          form.append('images', item.file, item.file.name);
        }
      }
      form.append('output_mode', isBlank ? 'blank' : 'existing');
      // D.4b: el modo multi-team controla cómo el backend procesa los
      // uploads cuando vienen de varios teams. Default 'strict'.
      form.append('multi_team_mode', state.multiTeamMode);
      if (!isBlank && state.pptxFile) {
        form.append('ppt', state.pptxFile);
      }

      currentMsg = isBlank
        ? 'Subiendo imágenes (PPT nuevo)...'
        : 'Subiendo PPT + imágenes...';
      statusEl.textContent = `⏱ 0s — ${currentMsg}`;
    } else {
      // Server mode (B.1 + B.2 + B.1d + D.5a): año/quarter/selections + scope opcional + multi_team_mode.
      endpoint = isBlank ? '/api/generate-blank' : '/api/generate';

      if (!isBlank) {
        // PPT existente: subimos el archivo + scope opcional (modo avanzado).
        form.append('ppt', state.pptxFile);
      }
      form.append('year', year);
      form.append('quarter', quarter);
      form.append('selections', JSON.stringify(activeTeams));
      // B.2d: en modo blank el scope es OBLIGATORIO; en existing avanzado
      // también lo mandamos para filtrar. Reusamos la misma fuente de verdad
      // (buildScope()) que ya alimenta /api/scan-catalog-scoped.
      if (isBlank || s.mode === 'advanced') {
        form.append('scope', JSON.stringify(buildScope()));
      }
      // D.5a: el modo multi-team también se manda al backend en server.
      // En PPT existente "Combinar en un deck" = concat (secciones por equipo
      // en un solo archivo). En PPT nuevo = mixed (imágenes mezcladas).
      const multiMode = (!isBlank && state.multiTeamMode === 'mixed') ? 'concat' : state.multiTeamMode;
      form.append('multi_team_mode', multiMode);

      // Mensaje inicial del status — refleja qué está pasando.
      if (isBlank) {
        currentMsg = 'Iniciando job (PPT nuevo)...';
        statusEl.textContent = `⏱ 0s — ${currentMsg}`;
      }
    }

    setProgress(5);
    // G.2: VLPS usa JSON body en vez de FormData (sin upload de archivos).
    const fetchOpts = useJsonBody
      ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(jsonBody) }
      : { method: 'POST', body: form };
    const r = await apiFetch(endpoint, fetchOpts);
    const initial = await r.json();
    const { job_id, status_url } = initial;
    if (!job_id || !status_url) {
      throw new Error('Respuesta inesperada del server: falta job_id/status_url');
    }
    currentMsg = 'En cola, esperando worker...';
    setProgress(8);

    // ═══ FASE 2: Polling del estado con backoff exponencial ═══
    let pollInterval = POLL_INTERVAL_MS;
    let consecutiveErrors = 0;
    let job = null;

    while (true) {
      await sleep(pollInterval);
      try {
        const sr = await apiFetch(status_url);
        if (sr.status === 404) {
          throw new Error('El job ya no existe (server reiniciado o expiró).');
        }
        job = await sr.json();
        consecutiveErrors = 0;
        pollInterval = POLL_INTERVAL_MS;

        // ── Actualizar UI con info REAL ──
        currentMsg = buildStatusMsg(job);
        setProgress(computeProgressPct(job));
        summaryEl.textContent = `${job.teams_done || 0}/${job.teams_total || selectedTeams.length} deck${(job.teams_total || 1) === 1 ? '' : 's'} generados`;
        if (selectedTeams.length > 1) {
          renderTeamList(teamListEl, job, selectedTeams);
        }

        if (job.status === 'done') break;
        if (job.status === 'failed') {
          throw new Error(job.error || 'El job falló en el servidor');
        }
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        // Si el error es "job no existe" (404 explícito), abortar sin reintentar.
        if (e.message && e.message.includes('ya no existe')) {
          throw e;
        }
        consecutiveErrors++;
        if (consecutiveErrors >= POLL_MAX_CONSECUTIVE_ERRORS) {
          throw new Error(`Múltiples errores de polling: ${e.message}`);
        }
        pollInterval = Math.min(pollInterval * 2, POLL_MAX_INTERVAL_MS);
        currentMsg = `Reintentando (${consecutiveErrors}/${POLL_MAX_CONSECUTIVE_ERRORS})...`;
      }
    }

    // ═══ FASE 3: Download ═══
    // Cambio importante: NO bajamos el archivo con fetch+blob() porque
    // materializar todo el output en RAM del browser revienta con outputs
    // grandes (un batch de 6 decks puede pesar 3+ GB). En vez de eso,
    // guardamos la URL y dejamos que el click del usuario dispare la
    // descarga nativa del browser via un <a href>. El browser maneja el
    // stream directo a disco con su download manager (progress bar, retry,
    // resume) sin pasar por memoria de la página.
    //
    // Auth: como el <a href> no puede setear el header Authorization, el
    // server setea un cookie httpOnly al cargar `/` (ver api.py::index)
    // y `verify_token` lo acepta como fallback al header. El cookie viaja
    // automáticamente en cualquier request al mismo origin, incluido el
    // click del <a>.
    setProgress(100);
    // Fallback del filename si el server no mandó Content-Disposition o
    // job.output_filename. Conservamos la misma lógica de antes pero ya
    // sin necesidad de fetchear los headers — el browser usa Content-
    // Disposition del server directamente al hacer el download.
    let fallback;
    if (isVlps) {
      const adv = state.selection.adv;
      fallback = `Pro Standard — VLPS ${adv.year || ''} ${adv.quarter || ''}.zip`.replace(/\s+/g, ' ').trim();
    } else {
      const baseName = state.pptxFile
        ? state.pptxFile.name.replace(/\.pptx$/i, '')
        : (isManual
            ? 'Pro Standard — Nuevo'
            : `Pro Standard — Nuevo ${year} ${quarter}`);
      fallback = activeTeams.length === 1
        ? `${baseName} — ${activeTeams[0].team}.pptx`
        : `${baseName} — Batch.${isBlank ? 'pptx' : 'zip'}`;
    }
    const filename = job.output_filename || fallback;

    state.outputDownloadUrl = job.download_url;
    state.outputFilename = filename;

    const elapsed = performance.now() - t_start;
    clearInterval(localTicker);
    setProgress(100);
    currentMsg = `Listo en ${formatElapsed(elapsed)}`;
    statusEl.textContent = currentMsg;
    statusEl.classList.add('done');

    // G.2-fix2: en VLPS el resumen no es "Generado para N equipos" —
    // mostramos el scope + ligas + tipos de archivo, alineado con la copy
    // del pre-gen-info y el status box de Step 2.
    if (isVlps) {
      const adv = state.selection.adv;
      const nCapsules = adv.capsules.length;
      const nClassics = adv.classics.length;
      const nLeagues = adv.leagues.length;
      const scopeParts = [];
      if (nCapsules > 0) {
        scopeParts.push(`${nCapsules} cápsula${nCapsules === 1 ? '' : 's'}`);
      }
      if (nClassics > 0) {
        scopeParts.push(`${nClassics} classic${nClassics === 1 ? '' : 's'}`);
      }
      const typesLabel = state.vlpsFileTypes.length === 2
        ? 'PPT + PDF'
        : (state.vlpsFileTypes[0] === 'ppt' ? 'Solo PPT' : 'Solo PDF');
      el.resultText.textContent =
        `${scopeParts.join(' · ')} · ${nLeagues} liga${nLeagues === 1 ? '' : 's'} · ${typesLabel} ` +
        `· listo en ${formatElapsed(elapsed)}`;
    } else {
      const teamsLabel = activeTeams.length === 1
        ? activeTeams[0].team
        : `${activeTeams.length} equipos`;
      el.resultText.textContent = `Generado para ${teamsLabel} en ${formatElapsed(elapsed)}`;
    }
    el.btnDownload.onclick = () => {
      // Download nativo del browser via <a href>. El browser stream el
      // archivo a disco con su download manager — sin materializar a RAM.
      // Funciona con archivos de cualquier tamaño. La auth viaja en el
      // cookie httpOnly que el server seteó al cargar `/`; verify_token
      // lo acepta como fallback al header Authorization.
      const a = document.createElement('a');
      a.href = state.outputDownloadUrl;
      a.download = state.outputFilename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    };
    el.resultCard.classList.add('show');
    el.btnReset.style.display = 'inline-flex';
    // G.2-fix2: tras el éxito ocultamos el botón "Generar/Descargar VLPS"
    // — ya cumplió su rol y dejarlo greyed-out al lado de "Nuevo Deck"
    // confunde al usuario ("¿por qué sigue ahí si ya descargué?").
    el.btnGenerate.style.display = 'none';
  } catch (e) {
    clearInterval(localTicker);
    const elapsed = performance.now() - t_start;
    currentMsg = `Error tras ${formatElapsed(elapsed)}`;
    statusEl.textContent = currentMsg;
    statusEl.classList.add('err');
    log(e.message, 'err');
    el.btnGenerate.disabled = false;
  }
}

function resetAll() {
  cancelAutoScan();
  state.pptxFile = null;
  // B.2d: reset también el output mode al default (PPT existente).
  state.outputMode = 'existing';
  // G.2: reset vlps file types al default (ambos).
  state.vlpsFileTypes = ['ppt', 'pdf'];
  // D.2: reset image origin + state de uploads manuales.
  state.imageOrigin = 'server';
  state.manualImagesByTeam = {};
  state.manualUnmatched = [];
  state.manualFlatFiles = [];
  // D.4b/D.5b: reset multi-team mode al default por flow (recomputado
  // después de resetear imageOrigin y outputMode), no-explícito.
  state.multiTeamMode = 'strict';
  state.multiTeamModeExplicit = false;
  // Reseteamos AMBAS partes del state (simple + adv) y volvemos al modo simple.
  state.selection = {
    mode: 'simple',
    year: null, quarter: null, leagues: [], teams: [],
    adv: {
      categories: [], year: null, quarter: null,
      capsules: [], classics: [], leagues: [], teams: [],
    },
  };
  state._advSubproductsCache = {};
  state.scanResult = null;
  // Reset del state de descarga. outputBlob ya no se usa (lo reemplazamos
  // por outputDownloadUrl), pero lo dejamos en null por compat por si algún
  // path legacy lo lee.
  state.outputBlob = null;
  state.outputDownloadUrl = null;
  state.outputFilename = null;

  el.inputPPT.value = '';
  setDropEmpty();
  el.slidesFound.style.display = 'none';
  el.slidesFound.innerHTML = '';
  setScanState('idle');
  el.scanTree.innerHTML = '';
  el.scanSummary.innerHTML = '';

  // Reset selectores simples
  ssYear.setSelected(null);
  ssQuarter.setOptions([]); ssQuarter.setSelected(null); ssQuarter.setDisabled(true);
  msLeagues.setOptions([]); msLeagues.setDisabled(true);
  msTeams.setOptions([]); msTeams.setDisabled(true);

  // Reset selectores avanzados
  if (advCategories) {
    advCategories.setOptions([]); advCategories.setDisabled(false);
    advYear.setOptions([]); advYear.setSelected(null);
    advQuarter.setSelected(null);
    advCapsules.setOptions([]); advCapsules.setDisabled(true);
    advClassics.setOptions([]); advClassics.setDisabled(true);
    advLeagues.setOptions([]); advLeagues.setDisabled(true);
    advTeams.setOptions([]); advTeams.setDisabled(true);
  }

  // Restaurar UI al modo simple
  el.modeSimple.checked = true;
  el.modeAdvanced.checked = false;
  el.simpleSelectors.hidden = false;
  el.advancedSelectors.hidden = true;

  el.btnScan.disabled = true;
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(OUTPUT_MODE_STORAGE_KEY);  // B.2d
  localStorage.removeItem(IMAGE_ORIGIN_STORAGE_KEY); // D.2
  localStorage.removeItem(MULTI_TEAM_MODE_STORAGE_KEY); // D.4b
  localStorage.removeItem(VLPS_FILE_TYPES_STORAGE_KEY); // G.2
  el.logBox.innerHTML = '';
  el.logBox.style.display = 'none';
  el.resultCard.classList.remove('show');
  el.btnGenerate.disabled = false;
  el.btnNext1.disabled = true;
  el.btnNext2.disabled = true;
  el.btnReset.style.display = 'none';
  el.progressFill.style.width = '0%';

  // B.2d: aplicar visual reset del output mode toggle (vuelve a 'existing'
  // y deshace los hidden/disabled del path blank).
  el.outputModeExisting.checked = true;
  el.outputModeBlank.checked = false;
  if (el.outputModeVlps) el.outputModeVlps.checked = false;
  el.stepper1.hidden = false;
  if (el.stepperLines && el.stepperLines[0]) el.stepperLines[0].hidden = false;
  el.step1.hidden = false;
  el.modeSimple.disabled = false;
  const simpleOptionLabel = el.modeSimple.closest('.mode-option');
  if (simpleOptionLabel) simpleOptionLabel.removeAttribute('title');
  el.btnGenerate.textContent = 'Generar Deck →';
  // G.2-fix2: tras el éxito ocultamos btnGenerate — al resetear hay que
  // volver a mostrarlo, y también restaurar todos los labels VLPS-specific
  // que setea switchOutputMode() para que no queden "pegados" al modo
  // anterior cuando el user vuelve al default 'existing'.
  el.btnGenerate.style.display = '';
  if (el.stepper3Label) el.stepper3Label.textContent = 'Generar';
  if (el.step3Heading) el.step3Heading.textContent = 'Generar Deck';
  if (el.step3Subtitle) el.step3Subtitle.textContent = 'Reemplaza los merchboards y descarga el PPT actualizado';
  if (el.resultTitle) el.resultTitle.textContent = '✓ DECK LISTO';
  if (el.btnDownload) el.btnDownload.textContent = '⬇ Descargar PPT';
  // G.2: visual reset de VLPS-specific elements.
  if (el.vlpsFileTypesRow) el.vlpsFileTypesRow.hidden = true;
  if (el.vlpsFileTypesBoth) el.vlpsFileTypesBoth.checked = true;
  if (el.vlpsFileTypesPpt) el.vlpsFileTypesPpt.checked = false;
  if (el.vlpsFileTypesPdf) el.vlpsFileTypesPdf.checked = false;
  // G.2-fix: ocultar status VLPS porque al resetear no estamos en VLPS.
  if (el.vlpsStatus) el.vlpsStatus.hidden = true;
  if (el.imageOriginToggle) el.imageOriginToggle.hidden = false;
  if (el.processingModeRow) el.processingModeRow.hidden = false;
  // Re-mostrar el selector de teams en advanced que VLPS oculta.
  const advTeamsField = el.advTeams ? el.advTeams.closest('.field') : null;
  if (advTeamsField) advTeamsField.hidden = false;
  const msTeamsField = el.msTeams ? el.msTeams.closest('.field') : null;
  if (msTeamsField) msTeamsField.hidden = false;
  if (el.btnBack2) el.btnBack2.hidden = false;

  // D.2: visual reset del toggle origen + manual UI
  el.imageOriginServer.checked = true;
  el.imageOriginManual.checked = false;
  el.serverSourceSection.hidden = false;
  el.manualSourceSection.hidden = true;
  // D.4b/D.5b: reset visual de los radios a default 'strict'. Después de
  // resetear outputMode + imageOrigin, el default por flow se recalcula
  // automáticamente vía la lógica de init/recompute.
  el.multiTeamStrict.checked = true;
  el.multiTeamMixed.checked = false;
  el.multiTeamPerTeam.checked = false;
  renderManualState();  // limpia drop-zone, badges, preview

  goStep(1);
}

// ═══════════════════════════════════════════════════════════════
// LISTENERS DE NAVEGACIÓN
// ═══════════════════════════════════════════════════════════════
el.btnNext1.addEventListener('click', () => goStep(2));
el.btnBack2.addEventListener('click', () => goStep(1));
el.btnNext2.addEventListener('click', () => goStep(3));
el.btnBack3.addEventListener('click', () => goStep(2));
el.btnGenerate.addEventListener('click', generateDeck);
el.btnReset.addEventListener('click', resetAll);

// ═══════════════════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════════════════
(async function init() {
  initStepper();
  initStep2();
  initOutputModeToggle();    // B.2d / G.2 — wire del toggle PPT existente / nuevo / VLPS
  initImageOriginToggle();   // D.2 — wire del toggle servidor / manual
  initMultiTeamModeToggle(); // D.4b — wire del toggle multi-team mode
  initVlpsFileTypesToggle(); // G.2 — wire del toggle PPT/PDF/Ambos en VLPS
  initThemeToggle();         // C.2 — wire del toggle light/dark
  await loadYears();
  await loadSelection();  // restaura selección y dispara fetches en cascada

  // B.2d: restaurar output mode persistido al cargar. Hay que hacerlo
  // DESPUÉS de loadSelection porque switchOutputMode('blank') puede
  // necesitar forzar switchMode('advanced') que requiere los selectores
  // ya inicializados con sus listas.
  try {
    const savedOutputMode = localStorage.getItem(OUTPUT_MODE_STORAGE_KEY);
    if (savedOutputMode === 'blank' || savedOutputMode === 'existing' || savedOutputMode === 'vlps') {
      await switchOutputMode(savedOutputMode);
    }
  } catch (_) {
    /* localStorage podría no estar disponible (modo privado en algunos
       browsers) — quedamos con el default 'existing'. */
  }

  // G.2: restaurar VLPS file types persistido.
  try {
    const savedVlpsTypes = localStorage.getItem(VLPS_FILE_TYPES_STORAGE_KEY);
    if (savedVlpsTypes === 'ppt') {
      state.vlpsFileTypes = ['ppt'];
      el.vlpsFileTypesPpt.checked = true;
      el.vlpsFileTypesBoth.checked = false;
    } else if (savedVlpsTypes === 'pdf') {
      state.vlpsFileTypes = ['pdf'];
      el.vlpsFileTypesPdf.checked = true;
      el.vlpsFileTypesBoth.checked = false;
    }
    // 'both' o ausente → default ya inicializado en state.vlpsFileTypes.
  } catch (_) {}

  // D.2: restaurar imageOrigin persistido. Independiente de outputMode —
  // las 4 combinaciones son válidas.
  try {
    const savedOrigin = localStorage.getItem(IMAGE_ORIGIN_STORAGE_KEY);
    if (savedOrigin === 'server' || savedOrigin === 'manual') {
      switchImageOrigin(savedOrigin);
    }
  } catch (_) {}

  // D.4b/D.5b: restaurar multi-team mode persistido.
  // Si hay valor en localStorage → el usuario lo eligió explícitamente
  // en una sesión anterior. Lo restauramos como explicit=true.
  // Si NO hay valor → calculamos el default por flow.
  try {
    const savedMtm = localStorage.getItem(MULTI_TEAM_MODE_STORAGE_KEY);
    if (savedMtm === 'strict' || savedMtm === 'mixed' || savedMtm === 'per_team') {
      switchMultiTeamMode(savedMtm, true);  // marcar explícito
    } else {
      // No hay valor persistido — usar el default per flow.
      switchMultiTeamMode(getDefaultMultiTeamMode(), false);
    }
  } catch (_) {
    switchMultiTeamMode(getDefaultMultiTeamMode(), false);
  }
})();
