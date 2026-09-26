'use strict';

/* ------------------------------------------------------------------ *
 * State
 * ------------------------------------------------------------------ */
const state = {
  mode: 'home',        // 'home' (drop-a-file screen) | 'sheet' (spreadsheet grid) | 'text' (plain-text editor) | 'design' (template editor, design.js)
  text: null,          // the open text document's settings while mode === 'text' (see enterTextMode)
  workbook: null,      // SheetJS workbook object
  hf: null,            // live HyperFormula engine instance, kept in sync with `workbook`
  sheetNames: [],
  activeSheet: null,   // name of active sheet
  originalFileName: '',
  title: '',            // editable display/base name, used when saving
  isCSV: false,
  rangeSel: null,       // { r1, c1, r2, c2, type } zero-based, relative to current !ref range start
                         // type: 'cell' | 'row' | 'col'
  headerAnchor: null,   // { axis: 'row'|'col', index } — for shift-click range extension on headers
  editingOriginal: null, // text content of the cell being edited, captured on focus
  editingTd: null,      // the <td> currently in edit mode (contentEditable), or null when only "selected"
  formulaPick: null,    // { target, surface, anchor, insertStart, insertEnd } while dragging out a formula reference
  formulaHomeSheet: null, // sheet name a formula edit is writing into (may differ from state.activeSheet
                          // while browsing another sheet to pick a cross-sheet reference)
  formulaHomeCell: null,  // { r, c } on formulaHomeSheet, for the same reason
  dirty: false,
  scratchId: null,      // 'sheet' | 'text' | 'design' | null — which home-screen scratchpad (if any) this document is
  clipEnabled: false, // Clip toggle: multi-line cells collapse to one row, ending with "…"
  undoStack: [],
  redoStack: [],
  scratchSheetUndoStack: null, // sheet's undoStack/redoStack, stashed across leave()/reopen of the
  scratchSheetRedoStack: null, // scratch sheet so Undo survives a trip home — mirrors scratch design's S.scratchHistory.
  fileHandle: null,    // FileSystemFileHandle for the current document, when it was opened via the
                        // native file picker, a handle-capable drag/drop, or a bookmark — null
                        // otherwise (classic <input type=file>, blank/scratch/Google-import docs).
                        // Threaded through as an explicit parameter from wherever a handle is
                        // obtained (never a shared/module-level variable, so nothing can leak a
                        // stale handle across an unrelated later action) down to whichever of
                        // activateWorkbook / enterTextMode / startDesign finishes the open, which
                        // copies it in here. Drives the toolbar bookmark star; see updateBookmarkStar().
};

const UNDO_LIMIT = 100;

/* ------------------------------------------------------------------ *
 * DOM refs
 * ------------------------------------------------------------------ */
const el = {
  newBtn: document.getElementById('newBtn'),
  openBtn: document.getElementById('openBtn'),
  createTextBtn: document.getElementById('createTextBtn'),
  createBlankBtn: document.getElementById('createBlankBtn'),
  createDesignBtn: document.getElementById('createDesignBtn'),
  googleImportBtn: document.getElementById('googleImportBtn'),
  googleImportForm: document.getElementById('googleImportForm'),
  googleImportInput: document.getElementById('googleImportInput'),
  googleImportError: document.getElementById('googleImportError'),
  googleImportSubmitBtn: document.getElementById('googleImportSubmitBtn'),
  googleImportCancelBtn: document.getElementById('googleImportCancelBtn'),
  fileInput: document.getElementById('fileInput'),
  fileName: document.getElementById('fileName'),
  renameTitleBtn: document.getElementById('renameTitleBtn'),
  undoBtn: document.getElementById('undoBtn'),
  redoBtn: document.getElementById('redoBtn'),
  deleteBtn: document.getElementById('deleteBtn'),
  addRowBtn: document.getElementById('addRowBtn'),
  addColBtn: document.getElementById('addColBtn'),
  exportFormat: document.getElementById('exportFormat'),
  saveBtn: document.getElementById('saveBtn'),
  clipBtn: document.getElementById('clipBtn'),
  clearSheetBtn: document.getElementById('clearSheetBtn'),
  clearSheetForm: document.getElementById('clearSheetForm'),
  clearSheetTabBtn: document.getElementById('clearSheetTabBtn'),
  clearSheetAllBtn: document.getElementById('clearSheetAllBtn'),
  clearSheetCancelBtn: document.getElementById('clearSheetCancelBtn'),
  formulaBar: document.getElementById('formulaBar'),
  formulaBarRef: document.getElementById('formulaBarRef'),
  formulaBarInput: document.getElementById('formulaBarInput'),
  formulaBarField: document.getElementById('formulaBarField'),
  formulaHint: document.getElementById('formulaHint'),
  formulaHintSyntax: document.getElementById('formulaHintSyntax'),
  formulaHintDesc: document.getElementById('formulaHintDesc'),
  selectionStats: document.getElementById('selectionStats'),
  selectionStatsBtn: document.getElementById('selectionStatsBtn'),
  selectionStatsLabel: document.getElementById('selectionStatsLabel'),
  selectionStatsValue: document.getElementById('selectionStatsValue'),
  dropzone: document.getElementById('dropzone'),
  gridWrapper: document.getElementById('gridWrapper'),
  gridScroll: document.getElementById('gridScroll'),
  grid: document.getElementById('grid'),
  sheetTabs: document.getElementById('sheetTabs'),
  toast: document.getElementById('toast'),
  floatNav: document.getElementById('floatNav'),
  scrollTopBtn: document.getElementById('scrollTopBtn'),
  scrollBottomBtn: document.getElementById('scrollBottomBtn'),
  passwordModal: document.getElementById('passwordModal'),
  passwordInput: document.getElementById('passwordInput'),
  passwordError: document.getElementById('passwordError'),
  passwordCancelBtn: document.getElementById('passwordCancelBtn'),
  passwordSubmitBtn: document.getElementById('passwordSubmitBtn'),
  textWrapper: document.getElementById('textWrapper'),
  textScroll: document.getElementById('textScroll'),
  textGutter: document.getElementById('textGutter'),
  textHlCode: document.getElementById('textHlCode'),
  textInput: document.getElementById('textInput'),
  textPos: document.getElementById('textPos'),
  textJump: document.getElementById('textJump'),
  textJumpForm: document.getElementById('textJumpForm'),
  clearTextBtn: document.getElementById('clearTextBtn'),
  textCount: document.getElementById('textCount'),
  textEnc: document.getElementById('textEnc'),
  textEolBtn: document.getElementById('textEolBtn'),
  textWrapBtn: document.getElementById('textWrapBtn'),
  textLang: document.getElementById('textLang'),
  loadingOverlay: document.getElementById('loadingOverlay'),
  loadingName: document.getElementById('loadingName'),
  loadingTrack: document.getElementById('loadingTrack'),
  loadingFill: document.getElementById('loadingFill'),
  loadingStage: document.getElementById('loadingStage'),
  loadingPct: document.getElementById('loadingPct'),
  scratchTiles: Array.from(document.querySelectorAll('.home-scratch')),
  bookmarkStarBtn: document.getElementById('bookmarkStarBtn'),
  homeCols: Array.from(document.querySelectorAll('.home-col')),
  homeSearch: document.getElementById('homeSearch'),
  homeSummary: document.getElementById('homeSummary'),
};

/* ------------------------------------------------------------------ *
 * Lazy-loaded libraries
 * ------------------------------------------------------------------ *
 * SheetJS + HyperFormula are ~3.1 MB of JavaScript that must be downloaded,
 * parsed and kept in memory. They are only needed once a *spreadsheet* is
 * opened (or a blank one created), so nothing is loaded at startup: the
 * home screen and the text editor never pay for them. The first spreadsheet
 * of a session loads them on demand while the loading bar (below) shows
 * progress; after that they stay loaded. Every code path that touches
 * XLSX / HyperFormula / OfficeCrypto sits behind a loaded workbook, which is
 * only ever created after `await ensureLibs()`.
 * ------------------------------------------------------------------ */
const LIB_SCRIPTS = [
  ['lib/xlsx.full.min.js', 0.31],        // [path, share of the total download, for the progress bar]
  ['lib/hyperformula.full.js', 0.68],
  ['lib/office-crypto.js', 0.01],
];
let libsPromise = null;
let libsLoaded = false;
let libsFraction = 0;
const libsListeners = new Set();

function loadScriptOnce(src) {
  return new Promise((resolve, reject) => {
    const tag = document.createElement('script');
    tag.src = src;
    tag.onload = () => resolve();
    tag.onerror = () => reject(new Error('Failed to load ' + src));
    document.head.appendChild(tag);
  });
}

// `onProgress(0..1)` is called as the scripts arrive (and once immediately).
function ensureLibs(onProgress) {
  if (onProgress) { libsListeners.add(onProgress); onProgress(libsFraction); }
  if (!libsPromise) {
    libsPromise = Promise.all(LIB_SCRIPTS.map(([src, weight]) =>
      loadScriptOnce(src).then(() => {
        libsFraction = Math.min(1, libsFraction + weight);
        libsListeners.forEach((fn) => fn(libsFraction));
      })
    )).then(() => { libsLoaded = true; libsListeners.clear(); }).catch((err) => {
      libsPromise = null; // allow a retry on the next attempt
      libsFraction = 0;
      libsListeners.clear();
      throw err;
    });
  }
  return libsPromise;
}

/* ------------------------------------------------------------------ *
 * Template design editor (Konva + design.js) — loaded the first time it's used
 * ------------------------------------------------------------------ */
let designPromise = null;
function ensureDesign() {
  if (!designPromise) {
    designPromise = loadScriptOnce('lib/konva.min.js')
      .then(() => loadScriptOnce('design.js'))
      .catch((err) => { designPromise = null; throw err; });
  }
  return designPromise;
}

// Called by everything that is about to take the screen over from the design editor.
function leaveDesignMode() {
  if (state.mode === 'design' && window.Design) window.Design.leave();
}

/* ------------------------------------------------------------------ *
 * Loading bar
 * ------------------------------------------------------------------ *
 * Opening a file can take a moment (the first spreadsheet also loads the
 * spreadsheet libraries; a big sheet has to be parsed, calculated and
 * drawn), so a small card with a progress bar shows what's happening until
 * the file is on screen.
 *
 * The heavy steps run on the main thread and can't report progress while
 * they work, so progress is by stage, and each stage yields to the browser
 * first (nextPaint) so the new label really gets drawn before the work
 * starts. The sliding highlight on the bar is a compositor animation: it
 * keeps moving even while the main thread is busy inside a long step.
 * ------------------------------------------------------------------ */
const loading = { active: false, showTimer: 0, shown: false };

// Resolves after the browser has painted whatever was just changed. (A
// background tab never paints, so a timeout stands in rather than waiting forever.)
function nextPaint() {
  return new Promise((resolve) => {
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(); } };
    requestAnimationFrame(() => setTimeout(finish, 0));
    setTimeout(finish, 80);
  });
}

function showLoadingCard() {
  clearTimeout(loading.showTimer);
  if (!loading.active || loading.shown) return;
  loading.shown = true;
  el.loadingOverlay.hidden = false;
}

// immediate: draw the card right away. Otherwise it only appears if the load
// is still going after a short moment, so quick opens don't flash it.
function beginLoading(name, immediate) {
  clearTimeout(loading.showTimer);
  loading.active = true;
  loading.shown = false;
  el.loadingName.textContent = name;
  setLoading(2, 'Starting\u2026');
  if (immediate) showLoadingCard();
  else loading.showTimer = setTimeout(showLoadingCard, 120);
}

function setLoading(pct, label) {
  if (!loading.active) return;
  const p = Math.max(0, Math.min(100, Math.round(pct)));
  el.loadingFill.style.transform = `scaleX(${p / 100})`;
  el.loadingTrack.setAttribute('aria-valuenow', String(p));
  el.loadingPct.textContent = `${p}%`;
  if (label) el.loadingStage.textContent = label;
}

function endLoading() {
  clearTimeout(loading.showTimer);
  loading.active = false;
  loading.shown = false;
  el.loadingOverlay.hidden = true;
}

/* ------------------------------------------------------------------ *
 * Scratchpads (Sheet / Text / Design)
 * ------------------------------------------------------------------ *
 * Three fixed, always-there documents shown on the home screen. Unlike a
 * regular file, a scratchpad is never "opened" from disk and never needs an
 * explicit save: every change is written to a local IndexedDB database a
 * moment after you make it, and reopening the same tile later (even after
 * the tab was closed or crashed) picks up right where you left off. Nothing
 * here ever leaves the browser.
 *
 * state.scratchId names which one (if any) of the three the open document
 * currently is ('sheet' | 'text' | 'design' | null). It's set by
 * openScratchpad() right after the document is loaded, and cleared by
 * leaveScratchpad() — called at the top of every function that tears down
 * the current document (activateWorkbook, enterTextMode, startDesign,
 * goToHome) — which also flushes one last save before the old content is
 * gone. A mutation while a scratchpad is open schedules a save via
 * scheduleScratchAutosave(); see pushUndo/applySnapshot (sheet), the
 * #textInput 'input' listener (text), and commit/restore in design.js.
 * ------------------------------------------------------------------ */
const SCRATCH_DB_NAME = 'sheetEditorScratchpads';
const SCRATCH_STORE = 'pads';
const SCRATCH_AUTOSAVE_DEBOUNCE = 600; // ms of quiet after an edit before it's written
const SCRATCH_LABELS = { sheet: 'Scratch sheet', text: 'Scratch text', design: 'Scratch design' };
const SCRATCH_DEFAULT_NAMES = { sheet: 'Scratch sheet.xlsx', text: 'Scratch text.txt', design: 'Scratch design' };

// Bookmarks share the scratchpad DB (bumped to version 2) rather than a second
// database — same infrastructure, a new object store. onupgradeneeded only
// creates whichever store is missing, so this never touches the existing
// 'pads' store or the scratchpad records already in it.
const BOOKMARK_STORE = 'bookmarks';

let scratchDbPromise = null;
function openScratchDb() {
  if (!scratchDbPromise) {
    scratchDbPromise = new Promise((resolve, reject) => {
      let req;
      try { req = indexedDB.open(SCRATCH_DB_NAME, 2); } catch (err) { reject(err); return; }
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(SCRATCH_STORE)) db.createObjectStore(SCRATCH_STORE, { keyPath: 'id' });
        if (!db.objectStoreNames.contains(BOOKMARK_STORE)) db.createObjectStore(BOOKMARK_STORE, { keyPath: 'id' });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error || new Error('IndexedDB failed to open'));
    }).catch((err) => { scratchDbPromise = null; throw err; });
  }
  return scratchDbPromise;
}
function scratchTxDone(tx) { return new Promise((resolve, reject) => { tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error); tx.onabort = () => reject(tx.error); }); }

async function scratchGet(id) {
  const db = await openScratchDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(SCRATCH_STORE, 'readonly').objectStore(SCRATCH_STORE).get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}
async function scratchGetAll() {
  const db = await openScratchDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(SCRATCH_STORE, 'readonly').objectStore(SCRATCH_STORE).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}
async function scratchPut(record) {
  const db = await openScratchDb();
  const tx = db.transaction(SCRATCH_STORE, 'readwrite');
  tx.objectStore(SCRATCH_STORE).put(record);
  await scratchTxDone(tx);
}
async function scratchDelete(id) {
  const db = await openScratchDb();
  const tx = db.transaction(SCRATCH_STORE, 'readwrite');
  tx.objectStore(SCRATCH_STORE).delete(id);
  await scratchTxDone(tx);
}

/* ------------------------------------------------------------------ *
 * File bookmarks
 * ------------------------------------------------------------------ *
 * A bookmark is { id, type: 'sheet'|'text'|'design', name, handle, addedAt }.
 * `handle` is the FileSystemFileHandle itself (structured-cloneable, so it's
 * stored directly) \u2014 no file content is ever kept, only the handle plus
 * display metadata. Three logical buckets by `type`, matching the homepage's
 * three format sections; each is capped at BOOKMARK_CAP records.
 * Two entry points write/read the same store: the in-editor star toggle
 * (updateBookmarkStar / toggleBookmarkStar below) and the homepage "Add new
 * bookmark" links in the homepage columns (refreshBookmarkPanels and friends) \u2014 no divergence
 * between them.
 * ------------------------------------------------------------------ */
const BOOKMARK_CAP = 50; // per section (text / spreadsheet / design)
const BOOKMARK_TYPE_LABELS = { text: 'text', sheet: 'spreadsheet', design: 'design' };

function bookmarkId() {
  try { return `bm_${crypto.randomUUID()}`; } catch (err) { return `bm_${Date.now()}_${Math.random().toString(36).slice(2)}`; }
}

async function bookmarkGetAllByType(type) {
  const db = await openScratchDb();
  const all = await new Promise((resolve, reject) => {
    const req = db.transaction(BOOKMARK_STORE, 'readonly').objectStore(BOOKMARK_STORE).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
  return all.filter((r) => r.type === type).sort((a, b) => (b.addedAt || '').localeCompare(a.addedAt || ''));
}
async function bookmarkPut(record) {
  const db = await openScratchDb();
  const tx = db.transaction(BOOKMARK_STORE, 'readwrite');
  tx.objectStore(BOOKMARK_STORE).put(record);
  await scratchTxDone(tx);
}
async function bookmarkDelete(id) {
  const db = await openScratchDb();
  const tx = db.transaction(BOOKMARK_STORE, 'readwrite');
  tx.objectStore(BOOKMARK_STORE).delete(id);
  await scratchTxDone(tx);
}

// Finds the stored bookmark (if any) of `type` that points at the same file
// as `handle`, checked via FileSystemFileHandle.isSameEntry \u2014 the only
// reliable way to compare handles, since two handles for the same file
// aren't ===. Runs against up to BOOKMARK_CAP records, so it's cheap.
async function findBookmarkForHandle(type, handle) {
  if (!handle || typeof handle.isSameEntry !== 'function') return null;
  const list = await bookmarkGetAllByType(type);
  for (const rec of list) {
    try { if (rec.handle && await handle.isSameEntry(rec.handle)) return rec; } catch (err) { /* a broken stored handle just never matches */ }
  }
  return null;
}

// "23-Sep-2026 11:27PM" — day-Mon-year, then 12-hour time with no space before AM/PM.
// Month names are hard-coded (not toLocaleDateString) so the format never varies with the browser's locale.
const BOOKMARK_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function bookmarkDateTime(iso) {
  const d = new Date(iso);
  if (!isFinite(d.getTime())) return '';
  const h24 = d.getHours();
  const h12 = h24 % 12 || 12;
  return `${String(d.getDate()).padStart(2, '0')}-${BOOKMARK_MONTHS[d.getMonth()]}-${d.getFullYear()} ${h12}:${String(d.getMinutes()).padStart(2, '0')}${h24 < 12 ? 'AM' : 'PM'}`;
}

// Short uppercase file-type label for a bookmark's name ("CSV", "XLSX", "TXT"…). A saved design
// template is "x.design.json", so it's called DESIGN rather than JSON. No extension → "FILE".
function bookmarkTypeLabel(name) {
  const n = String(name || '');
  if (/\.design\.json$/i.test(n)) return 'DESIGN';
  const m = /\.([^./\\]+)$/.exec(n);
  return m ? m[1].toUpperCase() : 'FILE';
}
// Fixed colours for the common types; anything else gets a stable colour derived from its label.
const BOOKMARK_TYPE_COLORS = {
  CSV: '#12805c', TSV: '#0f8a8a', XLSX: '#1e8e3e', XLSM: '#1e8e3e', XLS: '#5f8a1e', XLSB: '#5f8a1e', ODS: '#2f6fdb',
  DESIGN: '#e0651b',
  TXT: '#5b6b7b', MD: '#7b4fd1', JSON: '#c27a00', JS: '#a8870a', TS: '#2f6fdb', PY: '#2b6cb0', HTML: '#d9480f',
  CSS: '#c2255c', XML: '#8a5a00', YAML: '#a3475f', YML: '#a3475f', LOG: '#6b7280', SQL: '#0b7285',
};
function bookmarkTypeColor(label) {
  if (BOOKMARK_TYPE_COLORS[label]) return BOOKMARK_TYPE_COLORS[label];
  let h = 0;
  for (let i = 0; i < label.length; i++) h = (h * 31 + label.charCodeAt(i)) % 360;
  return `hsl(${h}, 55%, 38%)`;
}

// The type a bookmark star belongs to for whatever's on screen right now,
// or null when the current mode has nothing bookmarkable open.
function currentBookmarkType() {
  if (state.mode === 'sheet') return 'sheet';
  if (state.mode === 'text') return 'text';
  if (state.mode === 'design') return 'design';
  return null;
}

// Shows/hides and fills/outlines the toolbar star for whatever's open now.
// Cheap and safe to call any time \u2014 a no-op when nothing has a handle.
async function updateBookmarkStar() {
  const handle = state.fileHandle;
  const type = currentBookmarkType();
  if (!handle || !type || (type === 'design' && !/\.design\.json$/i.test(handle.name || ''))) {
    el.bookmarkStarBtn.hidden = true;
    return;
  }
  el.bookmarkStarBtn.hidden = false;
  el.bookmarkStarBtn.dataset.type = type;
  let existing = null;
  try { existing = await findBookmarkForHandle(type, handle); } catch (err) { console.error(err); }
  // The mode/handle could have changed while the lookup above was in flight
  // (e.g. rapid file switching) \u2014 don't paint a stale result over it.
  if (state.fileHandle !== handle || currentBookmarkType() !== type) return;
  const on = !!existing;
  el.bookmarkStarBtn.classList.toggle('is-bookmarked', on);
  el.bookmarkStarBtn.setAttribute('aria-pressed', String(on));
  el.bookmarkStarBtn.title = on ? 'Remove bookmark' : 'Bookmark this file';
}

async function toggleBookmarkStar() {
  const handle = state.fileHandle;
  const type = el.bookmarkStarBtn.dataset.type;
  if (!handle || !type) return;
  el.bookmarkStarBtn.disabled = true;
  try {
    const existing = await findBookmarkForHandle(type, handle);
    if (existing) {
      await bookmarkDelete(existing.id);
      showToast(`Removed bookmark for \u201c${existing.name}\u201d`);
    } else {
      const list = await bookmarkGetAllByType(type);
      if (list.length >= BOOKMARK_CAP) {
        showToast(`You can only bookmark up to ${BOOKMARK_CAP} ${BOOKMARK_TYPE_LABELS[type]} files \u2014 remove one first.`, 4000);
        return;
      }
      const name = handle.name || state.title || 'file';
      await bookmarkPut({ id: bookmarkId(), type, name, handle, addedAt: new Date().toISOString() });
      showToast(`Bookmarked \u201c${name}\u201d`);
    }
  } catch (err) {
    console.error(err);
    showToast('Couldn\u2019t update bookmarks.', 3500);
  } finally {
    el.bookmarkStarBtn.disabled = false;
    updateBookmarkStar();
    refreshBookmarkPanels();
  }
}
el.bookmarkStarBtn.addEventListener('click', toggleBookmarkStar);

/* ----- Homepage bookmark tables (one per column: text / sheet / design) ----- */

// Every stored bookmark, newest first, grouped by type. refreshBookmarkPanels()
// reloads it from IndexedDB; renderHomeBookmarks() paints it (through the search
// box's filter) without touching the database, so typing in the search box is instant.
let homeBookmarks = { text: [], sheet: [], design: [] };

function renderHomeBookmarks() {
  const q = el.homeSearch.value.trim().toLowerCase();
  let total = 0;
  for (const col of el.homeCols) {
    const type = col.dataset.type;
    const all = homeBookmarks[type] || [];
    total += all.length;
    const shown = q ? all.filter((r) => `${r.name} ${bookmarkTypeLabel(r.name)}`.toLowerCase().includes(q)) : all;
    col.querySelector('[data-role="count"]').textContent = q ? `${shown.length} of ${all.length}` : `${all.length} / ${BOOKMARK_CAP}`;
    col.querySelector('[data-role="empty"]').textContent = all.length ? 'No matches' : 'No bookmarks yet';
    col.classList.toggle('has-rows', shown.length > 0);
    const listEl = col.querySelector('[data-role="list"]');
    listEl.innerHTML = '';
    for (const [idx, rec] of shown.entries()) {
      const li = document.createElement('li');
      li.className = 'home-bm-row';
      li.tabIndex = 0;
      li.dataset.id = rec.id;
      li.setAttribute('role', 'button');
      li.setAttribute('aria-label', `Open bookmark ${rec.name}`);
      li.innerHTML = `
        <span class="home-bm-num"></span>
        <span class="home-bm-name"></span>
        <span class="home-bm-type"></span>
        <span class="home-bm-date"></span>
        <button type="button" class="home-bm-del" data-role="delete" title="Remove bookmark">
          <svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>
        </button>`;
      const nameEl = li.querySelector('.home-bm-name');
      nameEl.textContent = rec.name;
      nameEl.title = rec.name;
      li.querySelector('[data-role="delete"]').setAttribute('aria-label', `Remove bookmark ${rec.name}`);
      li.querySelector('.home-bm-num').textContent = String(idx + 1);
      const typeLabel = bookmarkTypeLabel(rec.name);
      const typeEl = li.querySelector('.home-bm-type');
      typeEl.textContent = typeLabel;
      typeEl.style.color = bookmarkTypeColor(typeLabel);
      li.querySelector('.home-bm-date').textContent = bookmarkDateTime(rec.addedAt);
      listEl.appendChild(li);
    }
  }
  el.homeSummary.textContent = `${total} bookmark${total === 1 ? '' : 's'} \u00b7 ${el.scratchTiles.length} scratchpads`;
}

// Reloads all three bookmark tables from IndexedDB. Cheap and safe to call any
// time the home screen might be showing (page load, after add/delete/open,
// after the star toggle changes something).
async function refreshBookmarkPanels() {
  const next = { text: [], sheet: [], design: [] };
  for (const type of Object.keys(next)) {
    try { next[type] = await bookmarkGetAllByType(type); } catch (err) { /* IndexedDB unavailable \u2014 the tables just stay empty */ }
  }
  homeBookmarks = next;
  renderHomeBookmarks();
}

// Opens a bookmarked file: re-requests permission (usually silent), reads it,
// and feeds it into the existing loadFile() pipeline \u2014 same path as any
// other opened file. A NotFoundError (moved/deleted/renamed on disk) is
// caught and offered as "remove this bookmark?" instead of a hard failure.
async function openBookmark(rec) {
  try {
    if (typeof rec.handle.requestPermission === 'function') {
      const perm = await rec.handle.requestPermission({ mode: 'read' });
      if (perm !== 'granted') { showToast(`Permission to open \u201c${rec.name}\u201d was denied.`, 4000); return; }
    }
    const file = await rec.handle.getFile();
    loadFile(file, undefined, rec.handle);
  } catch (err) {
    console.error(err);
    if (err && err.name === 'NotFoundError') {
      if (window.confirm(`Can\u2019t find \u201c${rec.name}\u201d \u2014 it may have moved, been renamed, or been deleted. Remove this bookmark?`)) {
        try { await bookmarkDelete(rec.id); } catch (delErr) { console.error(delErr); }
        refreshBookmarkPanels();
      }
    } else {
      showToast(`Couldn\u2019t open \u201c${rec.name}\u201d.`, 4000);
    }
  }
}

// Adds a new bookmark for `type` without loading the file into the editor \u2014
// a second, independent entry point from the in-editor star, reaching the
// same store. Per-type validation mirrors what the star/loadFile path
// already enforces, so a bad pick is rejected before it's ever saved.
async function addBookmarkForType(type) {
  if (!window.showOpenFilePicker) { showToast('Your browser doesn\u2019t support picking a file this way.', 4000); return; }
  const list = await bookmarkGetAllByType(type).catch(() => []);
  if (list.length >= BOOKMARK_CAP) {
    showToast(`You can only bookmark up to ${BOOKMARK_CAP} ${BOOKMARK_TYPE_LABELS[type]} files \u2014 remove one first.`, 4000);
    return;
  }
  const pickerOpts = type === 'sheet'
    ? { types: [{ description: 'Spreadsheets', accept: { 'application/octet-stream': Array.from(SHEET_EXTS).map((e) => `.${e}`) } }] }
    : type === 'design'
    ? { types: [{ description: 'Design templates', accept: { 'application/json': ['.design.json'] } }] }
    : {};
  let handle;
  try {
    [handle] = await window.showOpenFilePicker(pickerOpts);
  } catch (err) {
    if (err && err.name === 'AbortError') return; // user cancelled the picker
    console.error(err);
    showToast('Couldn\u2019t open the file picker.', 3500);
    return;
  }
  const name = handle.name || 'file';
  // Design: enforced by filename, not just the picker's type filter (which is only
  // a UI hint the user can override) \u2014 raw images don't carry the layers/text a
  // saved .design.json does.
  if (type === 'design' && !/\.design\.json$/i.test(name)) {
    showToast(`\u201c${name}\u201d isn\u2019t a saved design template (.design.json).`, 4000);
    return;
  }
  if (type === 'sheet' && !isSpreadsheetName(name)) {
    showToast(`\u201c${name}\u201d doesn\u2019t look like a supported spreadsheet file.`, 4000);
    return;
  }
  if (type === 'text') {
    let file;
    try { file = await handle.getFile(); } catch (err) { console.error(err); showToast('Couldn\u2019t read that file.', 3500); return; }
    if (file.size > TEXT_MAX_BYTES) { showToast(`\u201c${name}\u201d is too large for the text editor.`, 4000); return; }
    let bytes;
    try { bytes = new Uint8Array(await file.arrayBuffer()); } catch (err) { console.error(err); showToast('Couldn\u2019t read that file.', 3500); return; }
    if (looksBinary(bytes)) { showToast(`\u201c${name}\u201d looks like a binary file \u2014 it can\u2019t be bookmarked as text.`, 4500); return; }
  }
  const existing = await findBookmarkForHandle(type, handle).catch(() => null);
  if (existing) { showToast(`\u201c${name}\u201d is already bookmarked.`, 3200); return; }
  try {
    await bookmarkPut({ id: bookmarkId(), type, name, handle, addedAt: new Date().toISOString() });
  } catch (err) {
    console.error(err);
    showToast('Couldn\u2019t save that bookmark.', 3500);
    return;
  }
  showToast(`Bookmarked \u201c${name}\u201d`);
  refreshBookmarkPanels();
  updateBookmarkStar();
}

for (const col of el.homeCols) {
  const type = col.dataset.type;
  const listEl = col.querySelector('[data-role="list"]');
  const openRow = (row) => {
    const rec = (homeBookmarks[type] || []).find((r) => r.id === row.dataset.id);
    if (rec) openBookmark(rec);
  };
  col.querySelector('[data-role="add"]').addEventListener('click', () => addBookmarkForType(type));
  listEl.addEventListener('click', (e) => {
    const row = e.target.closest('.home-bm-row');
    if (!row) return;
    if (e.target.closest('[data-role="delete"]')) {
      e.stopPropagation();
      const name = row.querySelector('.home-bm-name').textContent;
      if (!window.confirm(`Remove the bookmark for \u201c${name}\u201d? This can\u2019t be undone.`)) return;
      bookmarkDelete(row.dataset.id).then(refreshBookmarkPanels).catch((err) => { console.error(err); showToast('Couldn\u2019t remove that bookmark.'); });
      return;
    }
    openRow(row);
  });
  listEl.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    const row = e.target.closest('.home-bm-row');
    if (!row || e.target.closest('[data-role="delete"]')) return;
    e.preventDefault();
    openRow(row);
  });
}

// Search all bookmarks: filters the three tables as you type (name or file type).
// "/" jumps to the box from anywhere on the home screen, Esc clears it, and
// Enter opens the first match.
el.homeSearch.addEventListener('input', renderHomeBookmarks);
el.homeSearch.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    e.preventDefault();
    if (el.homeSearch.value) { el.homeSearch.value = ''; renderHomeBookmarks(); } else el.homeSearch.blur();
  } else if (e.key === 'Enter' && el.homeSearch.value.trim()) {
    e.preventDefault();
    const first = el.homeCols.map((c) => c.querySelector('.home-bm-row')).find(Boolean);
    if (first) first.click();
  }
});
document.addEventListener('keydown', (e) => {
  if (e.key !== '/' || e.ctrlKey || e.metaKey || e.altKey) return;
  if (!document.body.classList.contains('mode-home')) return;
  const t = e.target;
  if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
  e.preventDefault();
  el.homeSearch.focus();
  el.homeSearch.select();
});
refreshBookmarkPanels();

// A short, human line describing what's currently on the sheet/text/canvas,
// shown on the home-screen tile. Never throws — worst case, an empty string.
function sheetScratchSnippet(ws) {
  try {
    if (!ws || !ws['!ref']) return '';
    const range = XLSX.utils.decode_range(ws['!ref']);
    const parts = [];
    for (let r = range.s.r; r <= range.e.r && r <= range.s.r + 6 && parts.length < 6; r++) {
      for (let c = range.s.c; c <= range.e.c && c <= range.s.c + 8 && parts.length < 6; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
          const v = String(cell.v).trim();
          if (v) parts.push(v.length > 22 ? v.slice(0, 22) + '\u2026' : v);
        }
      }
    }
    return parts.join('  \u00b7  ');
  } catch (err) { return ''; }
}
function textScratchSnippet(text) {
  const line = (text || '').split('\n').find((l) => l.trim() !== '');
  if (!line) return '';
  return line.length > 90 ? line.slice(0, 90) + '\u2026' : line;
}

// Builds the record to store for whichever scratchpad is open right now.
// Reads straight from the live document — nothing here is heavy enough to
// need debouncing itself (that's scheduleScratchAutosave's job).
async function buildScratchRecord(id) {
  const now = new Date().toISOString();
  if (id === 'text') {
    if (state.mode !== 'text' || !state.text) return null;
    const text = el.textInput.value;
    return {
      id, updatedAt: now, title: state.title || SCRATCH_DEFAULT_NAMES.text,
      snippet: textScratchSnippet(text),
      payload: { text, lang: state.text.lang, langManual: state.text.langManual, wrap: state.text.wrap },
    };
  }
  if (id === 'sheet') {
    if (state.mode !== 'sheet' || !state.workbook) return null;
    const xlsx = XLSX.write(state.workbook, { type: 'array', bookType: 'xlsx' });
    return {
      id, updatedAt: now, title: state.title || SCRATCH_DEFAULT_NAMES.sheet,
      snippet: sheetScratchSnippet(activeSheetObj()),
      payload: { xlsx, activeSheet: state.activeSheet },
    };
  }
  if (id === 'design') {
    if (state.mode !== 'design' || !window.Design) return null;
    const tpl = await window.Design._serializeTemplate();
    return {
      id, updatedAt: now, title: state.title || SCRATCH_DEFAULT_NAMES.design,
      snippet: `${tpl.layers.length} layer${tpl.layers.length === 1 ? '' : 's'} \u00b7 ${tpl.canvas.w}\u00d7${tpl.canvas.h}`,
      payload: { template: tpl },
    };
  }
  return null;
}

let scratchAutosaveTimer = 0;
// Called on every edit while a scratchpad is open (a no-op otherwise).
function scheduleScratchAutosave() {
  if (!state.scratchId) return;
  clearTimeout(scratchAutosaveTimer);
  scratchAutosaveTimer = setTimeout(() => { flushScratchAutosave(); }, SCRATCH_AUTOSAVE_DEBOUNCE);
}
// Saves right now instead of waiting for the debounce (tab hidden/closing).
function flushScratchAutosaveNow() {
  clearTimeout(scratchAutosaveTimer);
  return flushScratchAutosave();
}
let scratchAutosaveWarned = false; // only nag once per tab, not on every debounce tick
async function flushScratchAutosave(explicitId) {
  const id = explicitId || state.scratchId;
  if (!id) return;
  try {
    const record = await buildScratchRecord(id);
    if (record) await scratchPut(record);
  } catch (err) {
    // Logged with .name/.message explicitly: console.error(err) alone prints an
    // inspectable object in DevTools, but copying/pasting it elsewhere (or a tool
    // that stringifies it) collapses that to the unhelpful "[object DOMException]".
    console.error(`Scratchpad autosave failed: ${err && err.name}: ${err && err.message}`, err);
    if (!scratchAutosaveWarned) {
      scratchAutosaveWarned = true;
      const full = err && (err.name === 'QuotaExceededError' || err.code === 22);
      showToast(
        full
          ? "Autosave failed: the browser's local storage is full. Save this file to disk (Ctrl/Cmd+S) so nothing is lost."
          : 'Autosave failed — see the browser console for details. Your work is safe until you close the tab; save to disk (Ctrl/Cmd+S) to be sure.',
        6000,
      );
    }
  }
}
// Called at the top of every function that tears down the current document
// (see the module comment above). Saves the outgoing scratchpad's very last
// state before anything is cleared, then hands the id back off.
async function leaveScratchpad() {
  if (!state.scratchId) return;
  clearTimeout(scratchAutosaveTimer);
  const id = state.scratchId;
  // Stash the sheet's undo/redo stacks before whatever teardown comes next (activateWorkbook
  // or goToHome) wipes them, so reopening this scratchpad later in the same tab — even after
  // a trip to the homepage — can restore them (see openScratchpad's 'sheet' branch). Mirrors
  // how scratch design stashes S.history/S.hIndex in leave() (design.js).
  if (id === 'sheet') {
    state.scratchSheetUndoStack = state.undoStack;
    state.scratchSheetRedoStack = state.redoStack;
  }
  state.scratchId = null;
  await flushScratchAutosave(id);
}
// Best-effort final save if the tab is hidden or closed before the debounce fires.
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden' && state.scratchId) flushScratchAutosaveNow();
});
window.addEventListener('pagehide', () => { if (state.scratchId) flushScratchAutosaveNow(); });

// Opens one of the three scratchpads: whatever was saved for it, or a blank
// document the first time. `forceBlank` (used by "clear") skips the saved
// copy even if one exists.
async function openScratchpad(id, opts) {
  opts = opts || {};
  if (state.dirty && !state.scratchId && !window.confirm('Changes that you made may not be saved.')) return;
  let rec = null;
  if (!opts.forceBlank) {
    try { rec = await scratchGet(id); } catch (err) { console.error(err); }
  }
  // Design manages its own loading card (via startDesign); Sheet/Text use this one.
  if (id !== 'design') beginLoading(SCRATCH_LABELS[id], true);
  try {
    if (id === 'text') {
      const payload = rec && rec.payload;
      await enterTextMode({
        name: (rec && rec.title) || SCRATCH_DEFAULT_NAMES.text,
        text: (payload && payload.text) || '',
        encoding: 'utf-8', bom: false, eol: '\n', eolMixed: false,
        toast: rec ? `Opened \u201c${rec.title || SCRATCH_DEFAULT_NAMES.text}\u201d` : 'Created a blank text file',
      });
      if (payload && payload.lang && payload.langManual) setTextLanguage(payload.lang, true);
      if (payload && payload.wrap && state.text) { state.text.wrap = true; applyTextViewFlags(); }
    } else if (id === 'sheet') {
      if (!libsLoaded) { setLoading(15, 'Loading spreadsheet engine\u2026'); await ensureLibs((f) => setLoading(15 + 45 * f, 'Loading spreadsheet engine\u2026')); }
      const payload = rec && rec.payload;
      if (payload && payload.xlsx) {
        setLoading(70, 'Opening scratch sheet\u2026');
        const wb = XLSX.read(payload.xlsx, { type: 'array' });
        await activateWorkbook(wb, rec.title || SCRATCH_DEFAULT_NAMES.sheet, false, `Opened \u201c${rec.title || SCRATCH_DEFAULT_NAMES.sheet}\u201d`);
        if (payload.activeSheet && state.sheetNames.includes(payload.activeSheet) && payload.activeSheet !== state.activeSheet) {
          state.activeSheet = payload.activeSheet;
          renderSheetTabs();
          renderSheet();
        }
      } else {
        const ws = { '!ref': XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: BLANK_SHEET_ROWS - 1, c: BLANK_SHEET_COLS - 1 } }) };
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        await activateWorkbook(wb, SCRATCH_DEFAULT_NAMES.sheet, false, 'Created a blank scratch sheet');
        selectCell(0, 0);
      }
      // Resume: activateWorkbook() always wipes undoStack/redoStack, so restore whatever
      // was stashed the last time this scratchpad was left (see leaveScratchpad) — the
      // loaded payload is already whatever was live at that point, since leaving always
      // flushes an autosave first, so the two stay in step.
      if (state.scratchSheetUndoStack) {
        state.undoStack = state.scratchSheetUndoStack;
        state.redoStack = state.scratchSheetRedoStack || [];
        updateUndoRedoButtons();
      }
      state.scratchSheetUndoStack = null;
      state.scratchSheetRedoStack = null;
    } else if (id === 'design') {
      // startDesign() loads the design tools, resets for design mode and calls
      // Design.enter() — all with its own loading card, started and ended here.
      const payload = rec && rec.payload;
      // resume: true lets Design.enter() restore any undo/redo history stashed when this
      // scratchpad was last left, so a Clear + reopen can still be undone (see design.js).
      if (payload && payload.template) {
        await startDesign({ templateText: JSON.stringify(payload.template), title: rec.title || SCRATCH_DEFAULT_NAMES.design, resume: true }, SCRATCH_LABELS.design);
      } else {
        await startDesign({ title: SCRATCH_DEFAULT_NAMES.design, resume: true }, SCRATCH_LABELS.design);
      }
    }
    state.scratchId = id;
    state.dirty = false;
  } catch (err) {
    console.error(err);
    showToast(`Couldn\u2019t open your ${SCRATCH_LABELS[id].toLowerCase()}.`, 4000);
  } finally {
    if (id !== 'design') endLoading();
    refreshScratchTiles();
  }
}

async function clearScratchpad(id) {
  if (!window.confirm(`Clear your ${SCRATCH_LABELS[id].toLowerCase()}? This can\u2019t be undone.`)) return;
  try { await scratchDelete(id); } catch (err) { console.error(err); }
  if (state.scratchId === id) {
    await openScratchpad(id, { forceBlank: true });
  } else {
    refreshScratchTiles();
  }
  showToast(`Cleared your ${SCRATCH_LABELS[id].toLowerCase()}.`, 2200);
}

function scratchRelativeTime(iso) {
  const then = new Date(iso).getTime();
  if (!isFinite(then)) return '';
  const s = Math.max(0, Math.round((Date.now() - then) / 1000));
  if (s < 10) return 'Saved just now';
  if (s < 60) return `Saved ${s}s ago`;
  const m = Math.round(s / 60);
  if (m < 60) return `Saved ${m} minute${m === 1 ? '' : 's'} ago`;
  const h = Math.round(m / 60);
  if (h < 24) return `Saved ${h} hour${h === 1 ? '' : 's'} ago`;
  const d = Math.round(h / 24);
  if (d < 7) return `Saved ${d} day${d === 1 ? '' : 's'} ago`;
  return `Saved ${new Date(iso).toLocaleDateString()}`;
}

// Repaints the three home-screen tiles from IndexedDB. Cheap and safe to
// call any time (page load, returning to the home screen, after a save) —
// the tiles are simply hidden while a different screen is showing.
async function refreshScratchTiles() {
  let records = [];
  try { records = await scratchGetAll(); } catch (err) { /* IndexedDB unavailable (private mode, etc.) — tiles just stay in their empty state */ }
  const byId = new Map(records.map((r) => [r.id, r]));
  for (const tile of el.scratchTiles) {
    const id = tile.dataset.scratch;
    const rec = byId.get(id);
    const snippetEl = tile.querySelector('[data-role="snippet"]');
    const timeEl = tile.querySelector('[data-role="time"]');
    const clearBtn = tile.querySelector('[data-role="clear"]');
    const has = !!rec;
    tile.classList.toggle('is-empty', !has);
    if (snippetEl) snippetEl.textContent = has && rec.snippet ? rec.snippet : (has ? 'Empty' : 'Empty \u2014 click to start');
    if (timeEl) timeEl.textContent = has ? scratchRelativeTime(rec.updatedAt) : 'Not saved yet';
    if (clearBtn) clearBtn.hidden = !has;
  }
}

for (const tile of el.scratchTiles) {
  const id = tile.dataset.scratch;
  tile.addEventListener('click', (e) => {
    if (e.target.closest('[data-role="clear"]')) return; // handled below
    openScratchpad(id);
  });
  tile.addEventListener('keydown', (e) => {
    if ((e.key === 'Enter' || e.key === ' ') && !e.target.closest('[data-role="clear"]')) {
      e.preventDefault();
      openScratchpad(id);
    }
  });
  const clearBtn = tile.querySelector('[data-role="clear"]');
  if (clearBtn) {
    clearBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      clearScratchpad(id);
    });
  }
}
refreshScratchTiles();

/* ------------------------------------------------------------------ *
 * Export format toggle (XLSX | CSV)
 * ------------------------------------------------------------------ */
function getExportFormat() {
  return el.exportFormat.dataset.value === 'csv' ? 'csv' : 'xlsx';
}

function setExportFormat(format) {
  const value = format === 'csv' ? 'csv' : 'xlsx';
  el.exportFormat.dataset.value = value;
  el.exportFormat.querySelectorAll('button[data-value]').forEach((btn) => {
    const on = btn.dataset.value === value;
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-checked', String(on));
    btn.tabIndex = on ? 0 : -1;
  });
}

el.exportFormat.addEventListener('click', (e) => {
  const btn = e.target.closest('button[data-value]');
  if (btn) setExportFormat(btn.dataset.value);
});

// Arrow keys move between the two choices, like a native radio group.
el.exportFormat.addEventListener('keydown', (e) => {
  if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
  e.preventDefault();
  const next = getExportFormat() === 'csv' ? 'xlsx' : 'csv';
  setExportFormat(next);
  const btn = el.exportFormat.querySelector(`button[data-value="${next}"]`);
  if (btn) btn.focus();
});

/* ------------------------------------------------------------------ *
 * Keep this tab from being auto-discarded
 * ------------------------------------------------------------------ *
 * Chrome's Memory Saver discards idle background tabs. Switching back to a
 * discarded tab reloads the entire page from scratch — which is the pause
 * when returning to the editor, and it also throws away any unsaved edits.
 * Opting this tab out of automatic discarding avoids both.
 * ------------------------------------------------------------------ */
try {
  if (typeof chrome !== 'undefined' && chrome.tabs && chrome.tabs.getCurrent) {
    chrome.tabs.getCurrent((tab) => {
      if (chrome.runtime.lastError || !tab) return;
      chrome.tabs.update(tab.id, { autoDiscardable: false }, () => { void chrome.runtime.lastError; });
    });
  }
} catch (err) { /* not running as an extension page; nothing to do */ }

/* ------------------------------------------------------------------ *
 * Utilities
 * ------------------------------------------------------------------ */
function showToast(msg, ms = 2200) {
  el.toast.textContent = msg;
  el.toast.hidden = false;
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => { el.toast.hidden = true; }, ms);
}

function colLetter(n) {
  return XLSX.utils.encode_col(n);
}

function getRange(ws) {
  return XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
}

function activeSheetObj() {
  return state.workbook.Sheets[state.activeSheet];
}

// The worksheet's declared !ref range can extend well past the last row
// that actually has content (e.g. a file formatted/blanked out to row
// 1000). This scans the sheet's real cell entries to find the last
// (absolute) row index that holds an actual value or formula, across
// every column — not just one. Returns -1 if the sheet has no data.
function getLastDataRow(ws) {
  let maxRow = -1;
  for (const addr in ws) {
    if (addr[0] === '!') continue;
    const cell = ws[addr];
    if (!cell) continue;
    const hasContent = cell.v !== undefined || cell.f !== undefined;
    if (!hasContent) continue;
    const pos = XLSX.utils.decode_cell(addr);
    if (pos.r > maxRow) maxRow = pos.r;
  }
  return maxRow;
}

/* ------------------------------------------------------------------ *
 * Formula engine (HyperFormula)
 * ------------------------------------------------------------------ *
 * state.hf is a live HyperFormula instance kept in step with the
 * SheetJS workbook. Two ways of updating it:
 *   - recalcCellEdit(): incremental — pushes a single cell's new
 *     content into the existing engine (fast path for normal typing).
 *   - recalcFormulas(): full rebuild from the current workbook —
 *     used after structural changes (row/col add/delete, undo/redo,
 *     sheet rename, initial load) where it's simplest to just start
 *     the engine over from what's now in the workbook.
 * Either way, computed values are written back into each formula
 * cell's `.v`/`.w`/`.t` so rendering (which already knows how to
 * display a formula cell's cached value) picks them up for free.
 * ------------------------------------------------------------------ */
function destroyHf() {
  if (state.hf) {
    try { state.hf.destroy(); } catch (e) { /* ignore */ }
    state.hf = null;
  }
}

// Converts one HyperFormula CellValue (number | string | boolean | null |
// DetailedCellError) into the { v, w, t } shape SheetJS cells use.
function hfValueToCell(value) {
  if (value === null || value === undefined) {
    return { v: undefined, w: '', t: undefined };
  }
  if (typeof value === 'object' && 'value' in value && 'type' in value) {
    // A HyperFormula DetailedCellError — .value is already an Excel-style
    // error string like "#DIV/0!", "#REF!", "#NAME?".
    return { v: value.value, w: String(value.value), t: 'e' };
  }
  if (typeof value === 'boolean') {
    return { v: value, w: value ? 'TRUE' : 'FALSE', t: 'b' };
  }
  if (typeof value === 'number') {
    return { v: value, w: formatNumberForDisplay(value), t: 'n' };
  }
  return { v: value, w: String(value), t: 's' };
}

// HyperFormula already rounds off most float noise, but this trims any
// that's left so e.g. 0.1 + 0.2 shows "0.3", not "0.30000000000000004".
function formatNumberForDisplay(n) {
  if (Number.isInteger(n)) return String(n);
  return String(parseFloat(n.toPrecision(12)));
}

/* ------------------------------------------------------------------ *
 * Date display
 * ------------------------------------------------------------------ *
 * A spreadsheet stores a date as a plain serial number (2026-09-20 is
 * 46285); only its number format says "show this as a date". Formula
 * results lose that format on their way back from the engine, so dates
 * showed up as bare serials. Every date-valued cell is now shown as
 * 20-Sep-2026 (plus a time, when its format has one).
 *
 * This is display only: the stored value, number format and cached text
 * are left alone, so saved files are unaffected. To change the look, edit
 * DATE_DISPLAY_FORMAT (any SheetJS/Excel date pattern).
 * ------------------------------------------------------------------ */
const DATE_DISPLAY_FORMAT = 'dd-mmm-yyyy';
const MONTH_ABBRS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

// Classifies a number-format string. Returns null when it isn't a date
// format (General, currency, percent, a time of day on its own, a bare month
// or weekday name...), else { kind: 'date' | 'datetime', seconds, ampm }.
function dateInfoOfFormat(z) {
  if (typeof z === 'number') {
    try { z = XLSX.SSF.get_table()[z]; } catch (err) { return null; }
  }
  if (typeof z !== 'string' || z === '' || z.toLowerCase() === 'general') return null;
  try { if (!XLSX.SSF.is_date(z)) return null; } catch (err) { return null; }
  // Drop quoted text, escaped characters and [colour]/[locale] brackets so only
  // real format codes remain; [h], [mm], [ss] (elapsed time) count as time.
  const s = z.replace(/"[^"]*"/g, '').replace(/\\./g, '')
    .replace(/\[(?:h+|m+|s+)\]/gi, 'h').replace(/\[[^\]]*\]/g, '');
  const hasTime = /[hs]|am\/pm|a\/p/i.test(s);
  const hasDate = /y|d|mmm/i.test(s);
  if (!hasDate) return null; // time of day only
  if (/^[^a-z]*d{3,4}[^a-z]*$/i.test(s) || /^[^a-z]*m{3,5}[^a-z]*$/i.test(s)) return null; // weekday-only / month-only
  return { kind: hasTime ? 'datetime' : 'date', seconds: /s/i.test(s), ampm: /am\/pm|a\/p/i.test(s) };
}

// Is this numeric cell a date? Either its own number format says so, or —
// when it has no format of its own — the formula engine reported that the
// formula produced a date (TODAY(), DATE(...), EDATE(...), ...).
function dateInfoOfCell(cell) {
  if (!cell || cell.t !== 'n' || typeof cell.v !== 'number' || !isFinite(cell.v)) return null;
  const z = cell.z;
  const hasOwnFormat = typeof z === 'number' || (typeof z === 'string' && z !== '' && z.toLowerCase() !== 'general');
  if (hasOwnFormat) return dateInfoOfFormat(z); // an explicit non-date format (0.00, %, $...) stays as it is
  if (cell.__dateKind === 'date') return { kind: 'date' };
  if (cell.__dateKind === 'datetime') return { kind: 'datetime', seconds: true };
  return null;
}

function workbookIsDate1904() {
  const wb = state.workbook;
  return !!(wb && wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
}

// The on-screen text for a date cell, or null if it isn't one.
function dateDisplayText(cell) {
  const info = dateInfoOfCell(cell);
  if (!info) return null;
  const v = cell.v;
  if (!(v >= 1 && v < 2958466)) return null; // outside the range of real dates: show it as-is
  let fmt = DATE_DISPLAY_FORMAT;
  if (info.kind === 'datetime') {
    fmt += info.ampm
      ? ' h:mm' + (info.seconds ? ':ss' : '') + ' AM/PM'
      : ' hh:mm' + (info.seconds ? ':ss' : '');
  }
  // The engine is built on the workbook's own date system (see
  // recalcFormulas), so formula results and typed dates share one epoch.
  const date1904 = workbookIsDate1904();
  try { return XLSX.SSF.format(fmt, v, { date1904 }); } catch (err) { return null; }
}

// What the grid shows for a cell.
function cellDisplayText(cell) {
  const d = dateDisplayText(cell);
  if (d !== null) return d;
  if (cell.w !== undefined) return cell.w;
  return cell.v !== undefined ? cell.v : '';
}

// The reverse of dateDisplayText for typing/pasting: turns "20-Sep-2026" or
// "20-Sep-2026 13:30[:45]" (also with AM/PM) back into a serial number, so an
// edited or copy-pasted date stays a real date instead of turning into text.
function parseDisplayDate(raw) {
  const m = /^\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp][Mm]))?)?\s*$/.exec(raw);
  if (!m) return null;
  const mon = MONTH_ABBRS.indexOf(m[2].toLowerCase());
  if (mon < 0) return null;
  const day = Number(m[1]);
  const year = Number(m[3]);
  let hh = m[4] !== undefined ? Number(m[4]) : 0;
  const mi = m[5] !== undefined ? Number(m[5]) : 0;
  const ss = m[6] !== undefined ? Number(m[6]) : 0;
  if (m[7]) {
    if (hh < 1 || hh > 12) return null;
    hh = (hh % 12) + (m[7].toLowerCase() === 'pm' ? 12 : 0);
  }
  if (hh > 23 || mi > 59 || ss > 59) return null;
  const ms = Date.UTC(year, mon, day);
  const check = new Date(ms);
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== mon || check.getUTCDate() !== day) return null; // e.g. 31-Feb
  const epoch = workbookIsDate1904() ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const days = (ms - epoch) / 86400000;
  if (days < (workbookIsDate1904() ? 0 : 61)) return null; // before 1-Mar-1900 (Excel's leap-year quirk): leave as text
  const serial = parseFloat((days + (hh * 3600 + mi * 60 + ss) / 86400).toFixed(10));
  return { serial, hasTime: m[4] !== undefined };
}

// Remembers whether the engine says a formula cell's result is a date, so it
// can be shown as one even though the cell carries no date format.
function setEngineDateKind(cell, sheetId, row, col) {
  if (!cell || !state.hf) return;
  let kind;
  try {
    const t = state.hf.getCellValueDetailedType({ sheet: sheetId, row, col });
    if (t === 'NUMBER_DATE') kind = 'date';
    else if (t === 'NUMBER_DATETIME') kind = 'datetime';
  } catch (err) { /* engine can't say; leave it as a plain number */ }
  if (kind) cell.__dateKind = kind; else delete cell.__dateKind;
}

// HyperFormula doesn't implement every Excel function (e.g. WEBSERVICE,
// FILTERXML, some newer array/statistics functions), and it can't parse
// some formula syntax that only exists in other spreadsheet products —
// notably Google Sheets' internal `__xludf.DUMMYFUNCTION(...)` wrapper,
// which XLSX exports written by Google Sheets leave behind on any formula
// it couldn't translate to native Excel syntax. In both cases HyperFormula
// reports a spreadsheet-style error — "#NAME?" for an unrecognized
// function, "#ERROR!" for a formula it can't parse at all — but that
// doesn't mean the formula is actually wrong; it means *our* engine can't
// evaluate it. The workbook may already be carrying a perfectly good
// cached value for it (either typed by hand or exported by whichever tool
// originally computed it). Overwriting that with HyperFormula's own error
// would silently corrupt data the user never touched. This detects both
// cases (by HyperFormula's own error type/message) so callers can leave
// the existing cached value alone instead of clobbering it, and can flag
// the cell so the UI can visually distinguish it from a normal formula.
function isEngineUnsupportedError(value) {
  if (!value || typeof value !== 'object') return false;
  if (value.type === 'NAME' && typeof value.message === 'string' &&
    value.message.indexOf('Function name') === 0) return true;
  // A generic "ERROR" type is HyperFormula's catch-all for formula text it
  // could not parse/understand at all (as opposed to a real computed
  // result like #REF!, #DIV/0!, #VALUE!, #NUM!, #N/A or #CYCLE!, which
  // HyperFormula *did* evaluate and which should still be shown).
  if (value.type === 'ERROR') return true;
  return false;
}

// Marks (or clears) a cell as one whose cached value came from the
// workbook rather than from HyperFormula, because HyperFormula can't
// evaluate its formula. Kept in sync every time a formula cell's value is
// (or isn't) refreshed from the engine, so a cell that's later fixed —
// or whose formula becomes something the engine does support — loses the
// flag again instead of it going stale.
function markEngineUnsupported(cell, unsupported) {
  if (unsupported) cell.__hfUnsupported = true;
  else if (cell.__hfUnsupported) delete cell.__hfUnsupported;
}

// HyperFormula's built-in FILTER only accepts a single row or single column
// as its source ("a two-dimensional range is not supported") and insists the
// include array has the *same* dimensions as the source, so the standard
// spreadsheet form  =FILTER(C5:E9, B5:B9="Demo")  (2-D source, 1-column
// include) always came back as #N/A. This swaps in an Excel/Google-Sheets
// compatible FILTER:
//   FILTER(array, include [, include2 ...] [, if_empty])
//   - include is a column (same height as array)  -> keeps matching ROWS
//   - include is a row    (same width as array)   -> keeps matching COLUMNS
//   - extra include arrays of the same shape are AND-ed (Sheets style)
//   - a trailing single value is if_empty (Excel style)
//   - no match and no if_empty -> #N/A
// Must run before any HyperFormula instance is built (see recalcFormulas).
let compatFunctionsRegistered = false;
function registerCompatFunctions() {
  if (compatFunctionsRegistered) return;
  compatFunctionsRegistered = true;
  try {
    const { FunctionPlugin, FunctionArgumentType, SimpleRangeValue, CellError, ErrorType, ArraySize } = HyperFormula;

    const truthy = (v) => {
      if (v instanceof CellError) return v;
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (typeof v === 'string') return v.toUpperCase() === 'TRUE';
      return false; // blank
    };
    const isOne = (a) => a.width() === 1 && a.height() === 1;

    class CompatFilterPlugin extends FunctionPlugin {
      filter(ast, state) {
        return this.runFunction(ast.args, state, this.metadata('FILTER'), (source, include, ...rest) => {
          const includes = [include];
          let ifEmpty = null;
          for (let i = 0; i < rest.length; i++) {
            const a = rest[i];
            if (i === rest.length - 1 && isOne(a)) ifEmpty = a;
            else if (a.width() === include.width() && a.height() === include.height()) includes.push(a);
            else return new CellError(ErrorType.VALUE, 'FILTER: include arrays must have the same size.');
          }

          const byRows = include.width() === 1 && include.height() === source.height();
          const byCols = !byRows && include.height() === 1 && include.width() === source.width();
          if (!byRows && !byCols) {
            return new CellError(ErrorType.VALUE, 'FILTER: include must be one column as tall as the array, or one row as wide as the array.');
          }

          const n = byRows ? source.height() : source.width();
          const keep = [];
          for (let k = 0; k < n; k++) {
            let ok = true;
            for (const inc of includes) {
              const t = truthy(byRows ? inc.data[k][0] : inc.data[0][k]);
              if (t instanceof CellError) return t;
              if (!t) ok = false;
            }
            if (ok) keep.push(k);
          }

          if (keep.length === 0) {
            if (ifEmpty) return SimpleRangeValue.onlyValues([[ifEmpty.data[0][0]]]);
            return new CellError(ErrorType.NA, 'FILTER: no matches.');
          }
          const out = byRows
            ? keep.map((k) => source.data[k].slice())
            : source.data.map((row) => keep.map((k) => row[k]));
          return SimpleRangeValue.onlyValues(out);
        });
      }

      // The result can be at most as big as the source, so that is the area
      // the engine reserves for the spill; unused cells stay blank.
      filterArraySize(ast, state) {
        if (ast.args.length < 2) return ArraySize.error();
        const st = Object.assign(Object.create(Object.getPrototypeOf(state)), state);
        st.arraysFlag = true; // evaluate  B5:B9="Demo"  as an array, not a single cell
        const size = this.arraySizeForAst(ast.args[0], st);
        let { width, height } = size;
        // A whole-column / whole-row source (C:C, 5:5) has no fixed extent —
        // the engine reports its size as Infinity, so reserving room for the
        // result failed with #SPILL! ("No space for array result") and the
        // formula showed nothing. Such a range can only ever cover the
        // sheet's used area, so reserve that much instead.
        if (!Number.isFinite(width) || !Number.isFinite(height)) {
          const node = ast.args[0];
          const sheet = node && node.start && node.start.sheet != null ? node.start.sheet : state.formulaAddress.sheet;
          if (!Number.isFinite(height)) height = Math.max(1, this.dependencyGraph.getSheetHeight(sheet));
          if (!Number.isFinite(width)) width = Math.max(1, this.dependencyGraph.getSheetWidth(sheet));
        }
        return new ArraySize(width, height);
      }
    }
    CompatFilterPlugin.implementedFunctions = {
      FILTER: {
        method: 'filter',
        sizeOfResultArrayMethod: 'filterArraySize',
        enableArrayArithmeticForArguments: true,
        // 2+ arguments: array, include, then any number of extras
        parameters: [
          { argumentType: FunctionArgumentType.RANGE },
          { argumentType: FunctionArgumentType.RANGE },
        ],
        repeatLastArgs: 1,
      },
    };

    HyperFormula.unregisterFunction('FILTER');
    HyperFormula.registerFunction('FILTER', CompatFilterPlugin, { enGB: { FILTER: 'FILTER' } });
  } catch (err) {
    console.error('Could not install the compatible FILTER function; using the engine default.', err);
  }
}

// What a worksheet cell contributes to the engine. A cell holding an empty
// string (typical for cells a spreadsheet app "blanked out") counts as blank;
// passed through as '' it would silently block any array formula from
// spilling over it and turn the result into #SPILL!.
function hfCellContent(cell) {
  if (!cell) return null;
  if (cell.f !== undefined) return '=' + cell.f;
  if (cell.v === undefined || cell.v === '') return null;
  return cell.v;
}

// A cell object that exists but carries no data: no formula, and an empty
// or missing value (the "blanked out" cells spreadsheet apps export, or
// style-only stubs). hfCellContent() already hands these to the engine as
// BLANK — which is why an array formula (FILTER, UNIQUE, SEQUENCE, ...) is
// allowed to spill over them — so writing the spilled results back must
// treat them as free space too. Treating them as "a value the user typed"
// made only the first cell of a FILTER result appear and left the rest of
// the spill area blank.
function isBlankPlaceholder(cell) {
  return !!cell && cell.f === undefined && !cell.__spill && (cell.v === undefined || cell.v === '');
}

// Writes one spilled value into the worksheet. If the target was a blank
// placeholder, its original (blank) fields are remembered so they (and any
// style attached to the cell) can be put back when the spill later shrinks.
function writeSpillCell(ws, addr, cell, v, w, t) {
  if (!cell) cell = ws[addr] = {};
  else if (isBlankPlaceholder(cell)) cell.__placeholder = { t: cell.t, v: cell.v, w: cell.w };
  cell.v = v;
  cell.w = w;
  if (t !== undefined) cell.t = t;
  cell.__spill = true;
}

// Removes a spilled value: a plain spill cell disappears; one that sat on a
// blank placeholder goes back to being that placeholder.
function clearSpillCell(ws, addr) {
  const cell = ws[addr];
  if (!cell || !cell.__spill) return;
  const p = cell.__placeholder;
  if (!p) { delete ws[addr]; return; }
  delete cell.__spill;
  delete cell.__placeholder;
  ['t', 'v', 'w'].forEach((k) => { if (p[k] === undefined) delete cell[k]; else cell[k] = p[k]; });
}

// Builds the 0-based rectangular array HyperFormula wants for one sheet.
// Spans from row/col 0 (not ws's !ref start) up through the sheet's used
// range, so that array index === true spreadsheet row/column, matching
// how formula text like "=A1" is addressed.
function buildHFSheetArray(ws) {
  const range = getRange(ws);
  const nRows = range.e.r + 1;
  const nCols = range.e.c + 1;
  const arr = [];
  for (let r = 0; r < nRows; r++) {
    const row = [];
    for (let c = 0; c < nCols; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      // A __spill cell is just a cache of some other formula's last
      // computed output, not independent data — feeding its old value
      // back in as a literal constant would collide with that same
      // formula recomputing its own spill and produce a false #SPILL!
      // error. It'll be freshly recomputed (or cleared) right after by
      // applySheetFromEngine regardless.
      if (!cell || cell.__spill) { row.push(null); continue; }
      row.push(hfCellContent(cell));
    }
    arr.push(row);
  }
  return arr;
}

// Grows a sheet's declared !ref (never shrinks it) so it covers cell (r, c).
// The grid only draws rows/columns inside !ref, and file export only writes
// cells inside it — so a spill that runs past the last row/column has to
// pull !ref out with it, or it is silently cut off.
// This deliberately grows to the cells that actually hold a value, NOT to
// the engine's reported sheet dimensions: those include the whole area an
// array formula *reserves* (which can be far larger than what it fills),
// and feeding that back into !ref made the sheet balloon — and, since the
// next rebuild sizes the reservation from !ref, keep ballooning.
function growRangeToCell(ws, r, c) {
  const range = getRange(ws);
  if (r > range.e.r || c > range.e.c) {
    range.e.r = Math.max(range.e.r, r);
    range.e.c = Math.max(range.e.c, c);
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
}

// Rewrites every cell of `name`'s worksheet from the live engine's current
// computed state. Two kinds of cells get written: a formula's own origin
// cell (keeps its .f, refreshes .v/.w), and — this is what actually makes
// array formulas like SEQUENCE/UNIQUE/FILTER/TRANSPOSE show their full
// result instead of just one cell — every cell inside that formula's
// *spill* range, which has no .f of its own. Those are marked with
// cell.__spill so a later shrink/move/delete of the source formula can
// tell a spilled result apart from a value the user actually typed there,
// and clear it correctly instead of leaving a stale ghost value behind.
function applySheetFromEngine(name) {
  const hf = state.hf;
  if (!hf) return;
  const ws = state.workbook.Sheets[name];
  const sheetId = hf.getSheetId(name);
  if (sheetId === undefined) return;

  const dims = hf.getSheetDimensions(sheetId);
  const serialized = hf.getSheetSerialized(sheetId); // formula text at each origin cell, plain values elsewhere
  const values = hf.getSheetValues(sheetId); // computed value everywhere, aligned to the same [row][col]

  const seen = new Set();
  let maxR = -1;
  let maxC = -1;
  for (let r = 0; r < dims.height; r++) {
    const serRow = serialized[r] || [];
    const valRow = values[r] || [];
    for (let c = 0; c < dims.width; c++) {
      const raw = serRow[c];
      const isFormula = typeof raw === 'string' && raw.charAt(0) === '=';
      const value = valRow[c];
      if (!isFormula && (value === undefined || value === null)) continue; // genuinely empty cell

      const addr = XLSX.utils.encode_cell({ r, c });
      seen.add(addr);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
      let cell = ws[addr];

      if (isFormula) {
        if (!cell) cell = ws[addr] = {};
        cell.f = raw.slice(1);
        delete cell.__spill;
        delete cell.__placeholder;
        if (isEngineUnsupportedError(value)) { markEngineUnsupported(cell, true); continue; }
        markEngineUnsupported(cell, false);
        const { v, w, t } = hfValueToCell(value);
        if (v === undefined) delete cell.v; else cell.v = v;
        cell.w = w;
        if (t !== undefined) cell.t = t;
        setEngineDateKind(cell, sheetId, r, c);
        continue;
      }

      // Not a formula's own cell. If it already holds a plain value the
      // user typed directly (no .f, not a previous spill), leave it alone
      // rather than overwriting it with whatever the engine also has at
      // that same address.
      if (cell && cell.f === undefined && !cell.__spill && !isBlankPlaceholder(cell)) continue;
      if (isEngineUnsupportedError(value)) continue;

      const { v, w, t } = hfValueToCell(value);
      if (v === undefined && w === '') {
        clearSpillCell(ws, addr);
        continue;
      }
      writeSpillCell(ws, addr, cell, v, w, t);
      setEngineDateKind(ws[addr], sheetId, r, c);
    }
  }

  // Array-formula spills can reach past the sheet's declared range —
  // extend it so they have cells to render into.
  if (maxR >= 0) growRangeToCell(ws, maxR, maxC);

  // A spill cell the engine no longer reports anything for at all (the
  // source array formula shrank, moved, or was deleted, and the sheet's
  // dimensions shrank along with it) needs clearing too.
  for (const addr in ws) {
    if (addr[0] === '!') continue;
    const cell = ws[addr];
    if (cell && cell.__spill && !seen.has(addr)) clearSpillCell(ws, addr);
  }
}

// Full rebuild: tears down any existing engine and reconstructs it from
// scratch out of the current workbook, then refreshes every sheet's cells
// (formulas and their spills) from the fresh computation.
function recalcFormulas() {
  if (!state.workbook) return;
  destroyHf();
  registerCompatFunctions();

  const sheetsData = {};
  for (const name of state.sheetNames) {
    sheetsData[name] = buildHFSheetArray(state.workbook.Sheets[name]);
  }

  let hf;
  try {
    // A 1904-system workbook stores its dates as days since 1904-01-01. The
    // engine defaults to the 1900 system, so without this TODAY()/DATE()
    // results are 1462 days off, and YEAR()/date arithmetic on the sheet's
    // own dates is wrong too.
    const hfConfig = { licenseKey: 'gpl-v3' };
    if (workbookIsDate1904()) hfConfig.nullDate = { year: 1904, month: 1, day: 1 };
    hf = HyperFormula.buildFromSheets(sheetsData, hfConfig);
  } catch (err) {
    console.error('Formula engine failed to build', err);
    return;
  }
  state.hf = hf;

  for (const name of state.sheetNames) applySheetFromEngine(name);
  updateSelectionStats();
}

// After a sheet rename, the live engine's internal graph is renamed via
// hf.renameSheet() (which correctly keeps cross-sheet references intact).
// Every sheet then needs its cells refreshed from the engine too — a
// formula like "=OldName!A1" needs its literal text updated (not just its
// cached value) or it'd go stale the next time the engine rebuilds, and
// applySheetFromEngine's own getSheetSerialized() read already reflects
// the rename automatically.
function syncFormulasFromEngine() {
  if (!state.hf) return;
  for (const name of state.sheetNames) applySheetFromEngine(name);
  updateSelectionStats();
}

// Incremental path: pushes one cell's new content into the already-live
// engine and applies just the resulting changes, instead of rebuilding
// everything. This is what normal cell-by-cell editing uses.
function recalcCellEdit(ws, addr) {
  if (!state.hf) { recalcFormulas(); return; }
  const sheetId = state.hf.getSheetId(state.activeSheet);
  if (sheetId === undefined) { recalcFormulas(); return; }

  const pos = XLSX.utils.decode_cell(addr);
  const cell = ws[addr];
  const content = hfCellContent(cell);

  let changes;
  try {
    changes = state.hf.setCellContents({ sheet: sheetId, row: pos.r, col: pos.c }, content);
  } catch (err) {
    console.error('Formula engine update failed; rebuilding', err);
    recalcFormulas();
    return;
  }
  applyHfChanges(changes);
}

// Writes back everything the engine reports as changed: a formula cell's
// own new value, every other formula cell downstream of it, AND — unlike
// before — every cell inside an affected array formula's spill range, so
// SEQUENCE/UNIQUE/FILTER/TRANSPOSE-style results actually show beyond
// their first cell instead of being silently dropped for lacking their own
// .f. HyperFormula's `changes` already only includes cells actually
// affected by this edit (a plain constant elsewhere never "recomputes"),
// including ones that changed *to* blank — e.g. a spill that just shrank —
// so no separate broader cleanup pass is needed here.
// HyperFormula's change list reports cells that got a new value, but NOT
// the cells that just went blank because an array formula's result shrank
// (e.g. a FILTER whose criteria now match fewer rows): editing one cell
// left the old rows behind as ghost values until the next full rebuild.
// So after an edit that touched a formula or spill cell, sweep the sheet's
// spill cells and drop any the engine no longer has a value for.
function pruneStaleSpillCells(sheetName) {
  const hf = state.hf;
  const ws = state.workbook.Sheets[sheetName];
  const sheetId = hfSheetId(sheetName);
  if (!hf || !ws || sheetId === undefined) return;
  for (const addr in ws) {
    if (addr[0] === '!') continue;
    const cell = ws[addr];
    if (!cell || !cell.__spill) continue;
    const p = XLSX.utils.decode_cell(addr);
    let val;
    try { val = hf.getCellValue({ sheet: sheetId, row: p.r, col: p.c }); } catch (err) { continue; }
    if (val === null || val === undefined) clearSpillCell(ws, addr);
  }
}

function applyHfChanges(changes) {
  if (!changes || !state.hf) return;
  const touchedSheets = new Set();
  const sheetsToPrune = new Set();
  for (const change of changes) {
    const sheetName = state.hf.getSheetName(change.sheet);
    const ws = state.workbook.Sheets[sheetName];
    if (!ws) continue;
    touchedSheets.add(sheetName);
    const addr = XLSX.utils.encode_cell({ r: change.row, c: change.col });
    let existing = ws[addr];
    const isSourceCell = existing && existing.f !== undefined;
    if (isSourceCell || (existing && existing.__spill)) sheetsToPrune.add(sheetName); // a formula / spill result changed

    if (isSourceCell) {
      if (isEngineUnsupportedError(change.value)) { markEngineUnsupported(existing, true); continue; }
      markEngineUnsupported(existing, false);
      const { v, w, t } = hfValueToCell(change.value);
      if (v === undefined) delete existing.v; else existing.v = v;
      existing.w = w;
      if (t !== undefined) existing.t = t;
      setEngineDateKind(existing, change.sheet, change.row, change.col);
      continue;
    }

    // Not a formula's own cell — a spilled result cell from some array
    // formula. If it already holds a plain value the user typed directly
    // (no .f, not a previous spill), leave it alone.
    if (existing && !existing.__spill && !isBlankPlaceholder(existing)) continue;
    if (isEngineUnsupportedError(change.value)) continue;

    const { v, w, t } = hfValueToCell(change.value);
    if (v === undefined && w === '') {
      clearSpillCell(ws, addr);
      continue;
    }
    writeSpillCell(ws, addr, existing, v, w, t);
    setEngineDateKind(ws[addr], change.sheet, change.row, change.col);
  }
  // Spills can run past the sheet's current last row/column: make room.
  for (const change of changes) {
    const sheetName = state.hf.getSheetName(change.sheet);
    const ws = state.workbook.Sheets[sheetName];
    if (ws && ws[XLSX.utils.encode_cell({ r: change.row, c: change.col })]) growRangeToCell(ws, change.row, change.col);
  }
  for (const name of sheetsToPrune) pruneStaleSpillCells(name);
  updateSelectionStats(); // the selected cells' values may just have changed
}

/* ------------------------------------------------------------------ *
 * Structural edits (row/column insert & delete)
 * ------------------------------------------------------------------ *
 * Inserting or deleting a row/column physically moves cells, but a
 * formula's *text* (e.g. "=A2") has to be rewritten too, or it silently
 * points at the wrong cell after the move — including formulas on other
 * sheets that reference the affected one. HyperFormula's own addRows/
 * removeRows/addColumns/removeColumns already do this translation
 * correctly across the whole workbook, so structural edits ask the live
 * engine to perform the shift *before* the raw cell data is physically
 * moved in the worksheet, then pull every formula's corrected text and
 * value back out via syncFormulasFromEngine(). If the engine isn't
 * available or the operation fails, callers fall back to recalcFormulas(),
 * which rebuilds from the (now formula-stale) workbook — not perfect, but
 * no worse than before this fix.
 * ------------------------------------------------------------------ */
function hfSheetId(name) {
  if (!state.hf) return undefined;
  try { return state.hf.getSheetId(name); } catch (err) { return undefined; }
}

function hfInsertRows(sheetName, row, count) {
  const sheetId = hfSheetId(sheetName);
  if (sheetId === undefined) return false;
  try { state.hf.addRows(sheetId, [row, count]); return true; }
  catch (err) { console.error('Formula engine addRows failed', err); return false; }
}

function hfRemoveRows(sheetName, row, count) {
  const sheetId = hfSheetId(sheetName);
  if (sheetId === undefined) return false;
  try { state.hf.removeRows(sheetId, [row, count]); return true; }
  catch (err) { console.error('Formula engine removeRows failed', err); return false; }
}

function hfInsertColumns(sheetName, col, count) {
  const sheetId = hfSheetId(sheetName);
  if (sheetId === undefined) return false;
  try { state.hf.addColumns(sheetId, [col, count]); return true; }
  catch (err) { console.error('Formula engine addColumns failed', err); return false; }
}

function hfRemoveColumns(sheetName, col, count) {
  const sheetId = hfSheetId(sheetName);
  if (sheetId === undefined) return false;
  try { state.hf.removeColumns(sheetId, [col, count]); return true; }
  catch (err) { console.error('Formula engine removeColumns failed', err); return false; }
}

// Call once after the matching hf* structural op(s) above and the physical
// cell shift have both been applied. `engineInSync` should be false if any
// of the hf* calls failed (or the engine wasn't available), in which case
// this falls back to a full rebuild instead of trusting a partially-shifted
// engine.
function finishStructuralEdit(engineInSync) {
  if (engineInSync && state.hf) syncFormulasFromEngine();
  else recalcFormulas();
}

/* ------------------------------------------------------------------ *
 * Undo / redo
 * ------------------------------------------------------------------ *
 * Each undo step is a deep snapshot of the workbook plus which sheet
 * was active. Snapshots are taken before a mutation is applied, so
 * undo restores the state as it was just before that change.
 * ------------------------------------------------------------------ */
function snapshotState() {
  return {
    workbook: structuredClone(state.workbook),
    activeSheet: state.activeSheet,
    sheetNames: state.sheetNames.slice(),
  };
}

function pushUndo() {
  if (!state.workbook) return;
  state.undoStack.push(snapshotState());
  if (state.undoStack.length > UNDO_LIMIT) state.undoStack.shift();
  state.redoStack.length = 0;
  updateUndoRedoButtons();
  scheduleScratchAutosave();
}

function applySnapshot(snap) {
  state.workbook = snap.workbook;
  state.activeSheet = snap.activeSheet;
  state.sheetNames = snap.sheetNames;
  state.rangeSel = null;
  state.editingTd = null;
  state.dirty = true;
  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  scheduleScratchAutosave();
}

function undo() {
  if (state.mode === 'design') { if (window.Design) window.Design.undo(); return; }
  if (state.mode === 'text') { textUndoRedo('undo'); return; }
  if (!state.undoStack.length) return;
  const current = snapshotState();
  const prev = state.undoStack.pop();
  state.redoStack.push(current);
  applySnapshot(prev);
  updateUndoRedoButtons();
  showToast('Undo');
}

function redo() {
  if (state.mode === 'design') { if (window.Design) window.Design.redo(); return; }
  if (state.mode === 'text') { textUndoRedo('redo'); return; }
  if (!state.redoStack.length) return;
  const current = snapshotState();
  const next = state.redoStack.pop();
  state.undoStack.push(current);
  applySnapshot(next);
  updateUndoRedoButtons();
  showToast('Redo');
}

function updateUndoRedoButtons() {
  el.undoBtn.disabled = state.undoStack.length === 0;
  el.redoBtn.disabled = state.redoStack.length === 0;
}

/* ------------------------------------------------------------------ *
 * File loading
 * ------------------------------------------------------------------ */
// PNG / JPG / WebP / ... start a template at the picture's own size. (.svg is
// deliberately not here: it's text, and still opens in the text editor.)
const IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'avif']);
function isRasterImageFile(file) {
  const ext = (file.name || '').split('.').pop().toLowerCase();
  return (!!file.type && file.type.startsWith('image/') && file.type !== 'image/svg+xml') || IMAGE_EXTS.has(ext);
}

// Everything that belongs to the grid / text editor is put away before the design editor takes over.
function resetForDesign() {
  leaveTextMode();
  destroyHf();
  state.workbook = null;
  state.sheetNames = [];
  state.activeSheet = null;
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingOriginal = null;
  state.editingTd = null;
  state.formulaPick = null;
  state.formulaHomeSheet = null;
  state.formulaHomeCell = null;
  state.isCSV = false;
  state.clipEnabled = false;
  el.clipBtn.classList.remove('active');
  state.undoStack = [];
  state.redoStack = [];
  closeSheetMenu();
  el.grid.innerHTML = '';
  el.sheetTabs.innerHTML = '';
  updateSelectionStats();
  updateDeleteBtn();
  el.addRowBtn.disabled = true;
  el.addColBtn.disabled = true;
}

async function startDesign(opts, label) {
  const wasScratch = state.scratchId === 'design'; // dropping a file into an open design scratchpad keeps it bound to autosave (see leaveScratchpad's doc comment)
  await leaveScratchpad();
  beginLoading(label, false);
  setLoading(20, 'Loading the design tools\u2026');
  try {
    await ensureDesign();
    setLoading(60, opts.imageFile ? 'Reading image\u2026' : 'Preparing canvas\u2026');
    if (state.mode !== 'design') resetForDesign();
    await window.Design.enter(opts);
    state.fileHandle = opts.fileHandle || null;
    if (wasScratch) {
      state.scratchId = 'design';
      // A file dropped into an already-open scratchpad is tagged scratchpad_<name> so it's
      // clear the scratch slot now holds that file (Clear resets the title back to default).
      state.title = `scratchpad_${state.title}`;
      el.fileName.textContent = state.title;
      scheduleScratchAutosave();
    }
  } catch (err) {
    console.error(err);
    showToast(opts.templateText ? `Couldn\u2019t open \u201c${label}\u201d \u2014 ${err && err.message ? err.message : 'it isn\u2019t a design file'}.`
      : opts.imageFile ? `Couldn\u2019t open \u201c${label}\u201d as an image.` : 'Couldn\u2019t start the design editor.', 5000);
  } finally {
    endLoading();
    updateBookmarkStar();
  }
}

function openImageAsDesign(file) {
  if (state.dirty && !state.scratchId && !window.confirm('Changes that you made may not be saved.')) return;
  startDesign({ imageFile: file }, file.name || 'image');
}

function createBlankDesign() {
  if (state.dirty && !state.scratchId && !window.confirm('Changes that you made may not be saved.')) return;
  startDesign({}, 'New template');
}

// A design saved from here is a .json file whose first key is "type":"sheet-editor-design".
// It's recognised by that (not just its name), so a renamed copy still opens as a design
// while every other .json file keeps opening as text.
async function isDesignTemplateFile(file) {
  const name = file.name || '';
  if (/\.design\.json$/i.test(name)) return true;
  if (!/\.json$/i.test(name) || file.size < 40) return false;
  try { return /"type"\s*:\s*"sheet-editor-design"/.test(await file.slice(0, 400).text()); } catch (err) { return false; }
}

async function openTemplateFile(file, fileHandle) {
  if (state.dirty && !state.scratchId && !window.confirm('Changes that you made may not be saved.')) return;
  let text;
  try { text = await file.text(); } catch (err) { showToast('Failed to read file.'); return; }
  startDesign({ templateText: text, title: (file.name || '').replace(/\.design\.json$/i, '').replace(/\.json$/i, ''), fileHandle }, file.name || 'design');
}

// `fileHandle` is the FileSystemFileHandle `file` came from, when there is
// one (native picker, a handle-capable drop, or a bookmark) — threaded
// through as an explicit parameter, never a shared variable, so it can
// never leak into an unrelated later open. It ends up on state.fileHandle
// once the document actually activates; see activateWorkbook/enterTextMode/
// startDesign. Nothing here needs to clear it — a path with no handle
// simply never passes one, and it defaults to undefined/null throughout.
function loadFile(file, password, fileHandle) {
  const nm = file.name || '';
  if (!password && /\.json$/i.test(nm)) {
    isDesignTemplateFile(file).then((yes) => { if (yes) openTemplateFile(file, fileHandle); else loadFileNow(file, password, fileHandle); });
    return;
  }
  loadFileNow(file, password, fileHandle);
}

function loadFileNow(file, password, fileHandle) {
  const name = file.name || 'workbook.xlsx';
  const ext = name.split('.').pop().toLowerCase();
  if (!password && isRasterImageFile(file)) { openImageAsDesign(file); return; }

  // Spreadsheets open in the grid. Everything else is offered to the
  // plain-text editor, which decides from the file's *content* (not its
  // extension) whether it can be shown — see openTextFile.
  if (!isSpreadsheetName(name)) { openTextFile(file, name, fileHandle); return; }
  // Replacing an edited text document with a spreadsheet: ask first, like New does.
  // (Not on the password retry, which already got its answer.)
  if (!password && (state.mode === 'text' || state.mode === 'design') && state.dirty && !state.scratchId &&
      !window.confirm('Changes that you made may not be saved.')) return;
  // .xlsx/.xls/.ods/.numbers are all ZIP or CFB containers, so they're read as
  // raw bytes and SheetJS sniffs the actual format from the contents. Plain
  // delimited text (.csv/.tsv) is read as a string instead; SheetJS then
  // auto-detects the delimiter (comma vs tab) from the text itself.
  const isCSV = ext === 'csv' || ext === 'tsv';

  // Progress card: shown at once for a first load (the libraries still have to
  // arrive) or a big file, otherwise only if it turns out to take a moment.
  // Not for the password retry, whose dialog already shows its own busy state.
  if (!password) {
    beginLoading(name, !libsLoaded || file.size > 300 * 1024);
    setLoading(6, 'Reading file\u2026');
  }
  const reader = new FileReader();
  reader.onprogress = (e) => {
    if (e.lengthComputable) setLoading(6 + 16 * (e.loaded / e.total), 'Reading file\u2026');
  };
  reader.onload = (e) => {
    onFileBytesLoaded(e.target.result, file, name, isCSV, password, fileHandle);
  };
  reader.onerror = () => { endLoading(); showToast('Failed to read file.'); };

  if (isCSV) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

// Makes `wb` the workbook on screen: resets selection/undo state and swaps
// the drop-a-file screen for the grid. Shared by opening a file and by
// creating a blank spreadsheet, so both end up in exactly the same state.
async function activateWorkbook(wb, name, isCSV, toastMessage, fileHandle) {
  const wasScratch = state.scratchId === 'sheet'; // dropping a file into an open sheet scratchpad keeps it bound to autosave (see leaveScratchpad's doc comment)
  await leaveScratchpad();
  leaveDesignMode();
  leaveTextMode();
  state.mode = 'sheet';
  state.isCSV = isCSV;
  // Default the export format to the kind of file that was opened (a .csv
  // saves back as CSV, everything else as XLSX). Set here, before anything
  // renders, so the toggle is already right the first time it's seen.
  setExportFormat(isCSV ? 'csv' : 'xlsx');
  state.originalFileName = name;
  state.title = name.replace(/\.[^./\\]+$/, '') || 'workbook';
  state.workbook = wb;
  state.sheetNames = wb.SheetNames.slice();
  state.activeSheet = wb.SheetNames[0];
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingTd = null;
  state.dirty = false;
  state.undoStack = [];
  state.redoStack = [];
  state.fileHandle = fileHandle || null;

  el.fileName.textContent = state.title;
  el.renameTitleBtn.hidden = false;
  el.gridWrapper.hidden = false;
  el.gridWrapper.classList.remove('home-empty');
  document.body.classList.remove('mode-home');
  el.dropzone.hidden = true;
  el.gridScroll.hidden = false;
  el.floatNav.hidden = false;
  el.formulaBar.hidden = false;
  el.formulaBarInput.disabled = false;
  el.saveBtn.disabled = false;
  el.clipBtn.disabled = false;
  el.clearSheetBtn.disabled = false;
  hideClearSheetForm();
  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  updateUndoRedoButtons();
  setPasswordModalBusy(false);
  hidePasswordModal();
  showToast(toastMessage);
  updateBookmarkStar(); // the sheet path needs this too (enterTextMode / startDesign already call it)
  if (wasScratch) {
    state.scratchId = 'sheet';
    // A file dropped into an already-open scratchpad is tagged scratchpad_<name> so it's
    // clear the scratch slot now holds that file (Clear resets the title back to default).
    state.title = `scratchpad_${state.title}`;
    el.fileName.textContent = state.title;
    scheduleScratchAutosave();
  }
}

// SheetJS's ODS reader logs a console.error for certain conditional
// number-format expressions ("ODS number format may be incorrect:
// value()...") it isn't fully confident it parsed correctly — purely
// informational on its own end (the file still loads and renders
// correctly either way; this is a defensive log, not a thrown error), but
// each one shows up as a genuine error in Chrome's extension error list,
// which is misleading for something the file opening successfully
// already proves was harmless. Filtered out only for the duration of this
// one read call — everything else still logs normally.
const ODS_NUMFMT_NOISE = /ODS number format may be incorrect/;
function readWorkbookQuietingOdsNoise(data, readOpts) {
  const origError = console.error;
  const origWarn = console.warn;
  const quiet = (orig) => (...args) => {
    if (typeof args[0] === 'string' && ODS_NUMFMT_NOISE.test(args[0])) return;
    orig.apply(console, args);
  };
  console.error = quiet(origError);
  console.warn = quiet(origWarn);
  try {
    return XLSX.read(data, readOpts);
  } finally {
    console.error = origError;
    console.warn = origWarn;
  }
}

async function onFileBytesLoaded(data, file, name, isCSV, password, fileHandle) {
  try {
    await ensureLibs((f) => setLoading(22 + 32 * f, 'Loading spreadsheet engine\u2026'));
  } catch (err) {
    endLoading();
    console.error(err);
    showToast('Could not load the spreadsheet libraries (check the extension\u2019s lib folder).');
    return;
  }
  try {
    // For non-CSV files, check up front whether this is a password-protected
    // OOXML document (a CFB/OLE container) and decrypt it ourselves before
    // handing bytes to SheetJS, which can only detect encryption, not decrypt it.
    if (!isCSV && OfficeCrypto.isCFBFile(data)) {
      if (!password) {
        showPasswordModal(file, fileHandle);
        return;
      }
      setPasswordModalBusy(true);
      try {
        data = await OfficeCrypto.decryptOfficeFile(data, password);
      } catch (err) {
        setPasswordModalBusy(false);
        if (err && err.code === 'WRONG_PASSWORD') {
          showPasswordModalError('Incorrect password. Try again.');
        } else if (err && err.code === 'UNSUPPORTED_ENCRYPTION_VERSION') {
          showPasswordModalError('This file uses an encryption scheme this extension doesn\u2019t support yet.');
        } else {
          console.error(err);
          showPasswordModalError('Could not decrypt this file.');
        }
        return;
      }
    }

    const readOpts = isCSV
      ? { type: 'string', cellFormula: true, cellNF: true }
      : { type: 'array', cellFormula: true, cellNF: true, cellStyles: true };

    setLoading(58, 'Reading spreadsheet\u2026');
    if (loading.active) await nextPaint();
    const wb = readWorkbookQuietingOdsNoise(data, readOpts);
    setLoading(80, 'Building sheet\u2026');
    if (loading.active) await nextPaint();
    await activateWorkbook(wb, name, isCSV, `Loaded "${name}"`, fileHandle);
  } catch (err) {
    setPasswordModalBusy(false);
    console.error(err);
    showToast('Could not read that file. Is it a valid .xlsx, .csv, .tsv, .ods, or .numbers file?');
  } finally {
    endLoading(); // also covers the early returns (password prompt, wrong password, …)
  }
}

/* ------------------------------------------------------------------ *
 * Password-protected file prompt
 * ------------------------------------------------------------------ */
let pendingPasswordFile = null;
let pendingPasswordFileHandle = null;

function showPasswordModal(file, fileHandle) {
  pendingPasswordFile = file;
  pendingPasswordFileHandle = fileHandle || null;
  el.passwordInput.value = '';
  el.passwordError.hidden = true;
  el.passwordModal.hidden = false;
  el.passwordInput.focus();
}

function hidePasswordModal() {
  pendingPasswordFile = null;
  pendingPasswordFileHandle = null;
  el.passwordModal.hidden = true;
  setPasswordModalBusy(false);
}

function showPasswordModalError(msg) {
  el.passwordError.textContent = msg;
  el.passwordError.hidden = false;
}

function setPasswordModalBusy(isBusy) {
  el.passwordSubmitBtn.disabled = isBusy;
  el.passwordCancelBtn.disabled = isBusy;
  el.passwordInput.disabled = isBusy;
  el.passwordSubmitBtn.textContent = isBusy ? 'Decrypting\u2026' : 'Open';
  if (isBusy) el.passwordError.hidden = true;
}

function submitPassword() {
  if (!pendingPasswordFile) return;
  const pwd = el.passwordInput.value;
  if (!pwd) {
    showPasswordModalError('Enter a password.');
    return;
  }
  loadFile(pendingPasswordFile, pwd, pendingPasswordFileHandle);
}

el.passwordCancelBtn.addEventListener('click', hidePasswordModal);
el.passwordSubmitBtn.addEventListener('click', submitPassword);
el.passwordInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); submitPassword(); }
  else if (e.key === 'Escape') { e.preventDefault(); hidePasswordModal(); }
});

/* ------------------------------------------------------------------ *
 * Plain-text / code editor
 * ------------------------------------------------------------------ *
 * Any file that isn't a spreadsheet opens here. The file's extension is only
 * ever used to pick syntax colours — never to decide whether it can be
 * opened. What decides that is the content: a file that looks binary (NUL
 * bytes, lots of control characters, a well-known binary signature) is turned
 * away, and everything else is decoded as text (UTF-8; UTF-16 when it starts
 * with a BOM; Windows-1252 when it isn't valid UTF-8) and edited in a plain
 * <textarea>. Encoding, BOM and line endings are remembered and written back
 * exactly as they were found.
 *
 * Syntax colouring is a nicety layered on top and can never be the reason a
 * file fails to open: highlight.js (its "common" bundle, which already holds
 * the ~35 most used grammars, plus a small file per remaining language) is
 * fetched on demand from lib/hljs/ the first time a text file opens —
 * self-hosted, so the manifest's script-src 'self' is all that's needed. A
 * file with no matching grammar — or with the grammar files missing — is just
 * shown as plain monospace text. The extension → language table is
 * TEXT_LANGUAGES.
 * ------------------------------------------------------------------ */
const SHEET_EXTS = new Set([
  'xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'ods', 'fods', 'numbers', 'csv', 'tsv',
]);

function isSpreadsheetName(name) {
  const dot = name.lastIndexOf('.');
  return dot >= 0 && SHEET_EXTS.has(name.slice(dot + 1).toLowerCase());
}

const TEXT_MAX_BYTES = 20 * 1024 * 1024; // bigger than this isn't opened at all
const TEXT_OVERLAY_MAX = 1000000;        // characters; above this: bare textarea (no colours, no line numbers)
const TEXT_HL_MAX = 400000;              // characters; above this: no syntax colours
const TEXT_HL_SYNC_MAX = 100000;         // characters; up to this, colours are refreshed on every keystroke
                                         // (~50 ms) — above it they catch up once typing pauses
const TEXT_LINE_PX = 20;                 // one text row, in px — keep in step with --code-line in styles.css

// [highlight.js grammar (= file name in lib/hljs/lang/), label, extensions, exact lower-case file names]
const TEXT_LANGUAGES = [
  ['apache',       'Apache config',    [], ['.htaccess', 'httpd.conf', 'apache2.conf']],
  ['bash',         'Shell',            ['sh', 'bash', 'zsh', 'ksh', 'command', 'bashrc', 'zshrc', 'bash_profile', 'bash_aliases', 'zprofile', 'zshenv', 'profile', 'envrc']],
  ['c',            'C',                ['c', 'h']],
  ['clojure',      'Clojure',          ['clj', 'cljs', 'cljc', 'edn']],
  ['cmake',        'CMake',            ['cmake'], ['cmakelists.txt']],
  ['coffeescript', 'CoffeeScript',     ['coffee']],
  ['cpp',          'C++',              ['cpp', 'cc', 'cxx', 'c++', 'hpp', 'hh', 'hxx', 'h++', 'ino', 'ipp', 'tpp']],
  ['csharp',       'C#',               ['cs', 'csx']],
  ['css',          'CSS',              ['css']],
  ['dart',         'Dart',             ['dart']],
  ['diff',         'Diff / patch',     ['diff', 'patch', 'rej']],
  ['dockerfile',   'Dockerfile',       ['dockerfile'], ['dockerfile', 'containerfile']],
  ['dos',          'Batch',            ['bat', 'cmd']],
  ['elixir',       'Elixir',           ['ex', 'exs']],
  ['elm',          'Elm',              ['elm']],
  ['erlang',       'Erlang',           ['erl', 'hrl']],
  ['fsharp',       'F#',               ['fs', 'fsx', 'fsi']],
  ['go',           'Go',               ['go']],
  ['gradle',       'Gradle',           ['gradle']],
  ['groovy',       'Groovy',           ['groovy', 'gvy'], ['jenkinsfile']],
  ['handlebars',   'Handlebars',       ['hbs', 'handlebars', 'mustache']],
  ['haskell',      'Haskell',          ['hs', 'lhs']],
  ['ini',          'INI / TOML',       ['ini', 'toml', 'cfg', 'conf', 'cnf', 'editorconfig', 'gitconfig', 'gitmodules', 'desktop', 'service', 'env']],
  ['java',         'Java',             ['java']],
  ['javascript',   'JavaScript',       ['js', 'mjs', 'cjs', 'jsx']],
  ['json',         'JSON',             ['json', 'jsonc', 'json5', 'jsonl', 'ndjson', 'geojson', 'webmanifest', 'ipynb', 'har', 'babelrc', 'eslintrc', 'prettierrc']],
  ['julia',        'Julia',            ['jl']],
  ['kotlin',       'Kotlin',           ['kt', 'kts']],
  ['latex',        'LaTeX',            ['tex', 'latex', 'sty', 'cls', 'bib']],
  ['less',         'Less',             ['less']],
  ['lisp',         'Lisp',             ['lisp', 'lsp', 'cl', 'el']],
  ['lua',          'Lua',              ['lua']],
  ['makefile',     'Makefile',         ['mk', 'mak'], ['makefile', 'gnumakefile']],
  ['markdown',     'Markdown',         ['md', 'markdown', 'mdx', 'mkd', 'mdown', 'mkdn']],
  ['matlab',       'MATLAB',           []],
  ['nginx',        'Nginx config',     ['nginx'], ['nginx.conf']],
  ['objectivec',   'Objective-C',      ['m', 'mm']],
  ['ocaml',        'OCaml',            ['ml', 'mli']],
  ['perl',         'Perl',             ['pl', 'pm', 't']],
  ['php',          'PHP',              ['php', 'phtml', 'php3', 'php4', 'php5', 'phps']],
  ['powershell',   'PowerShell',       ['ps1', 'psm1', 'psd1']],
  ['properties',   'Properties',       ['properties']],
  ['protobuf',     'Protocol Buffers', ['proto']],
  ['python',       'Python',           ['py', 'pyw', 'pyi', 'wsgi']],
  ['r',            'R',                ['r']],
  ['ruby',         'Ruby',             ['rb', 'rake', 'gemspec', 'ru', 'podspec', 'thor'], ['gemfile', 'rakefile', 'podfile', 'vagrantfile', 'guardfile', 'brewfile', 'capfile', 'fastfile']],
  ['rust',         'Rust',             ['rs']],
  ['scala',        'Scala',            ['scala', 'sc', 'sbt']],
  ['scheme',       'Scheme',           ['scm', 'ss', 'rkt', 'sld']],
  ['scss',         'SCSS',             ['scss']],
  ['sql',          'SQL',              ['sql', 'ddl', 'dml', 'psql', 'mysql']],
  ['swift',        'Swift',            ['swift']],
  ['typescript',   'TypeScript',       ['ts', 'tsx', 'mts', 'cts']],
  ['vbnet',        'VB.NET',           ['vb']],
  ['vbscript',     'VBScript',         ['vbs']],
  ['vim',          'Vim script',       ['vim', 'vimrc']],
  ['xml',          'HTML / XML',       ['xml', 'html', 'htm', 'xhtml', 'svg', 'xsl', 'xslt', 'xsd', 'rss', 'atom', 'plist', 'vue', 'svelte', 'csproj', 'vbproj', 'fsproj', 'props', 'targets', 'resx', 'wsdl', 'kml', 'gpx', 'xaml', 'opml', 'jsp', 'aspx', 'ascx', 'cshtml', 'ejs', 'erb', 'xlf', 'xliff']],
  ['yaml',         'YAML',             ['yaml', 'yml']],
];

// Grammars that embed other languages (an HTML page's <script>/<style>,
// Markdown's inline HTML) — loaded alongside so those parts get coloured too.
const TEXT_LANG_DEPS = { xml: ['javascript', 'css'], markdown: ['xml'] };

const TEXT_LANG_BY_EXT = new Map();
const TEXT_LANG_BY_NAME = new Map();
TEXT_LANGUAGES.forEach(([id, , exts = [], names = []]) => {
  exts.forEach((x) => TEXT_LANG_BY_EXT.set(x, id));
  names.forEach((n) => TEXT_LANG_BY_NAME.set(n, id));
});

// Fill the status bar's language picker: "Plain text" first, the rest A–Z.
(function buildTextLangSelect() {
  const rows = TEXT_LANGUAGES.map(([id, label]) => [id, label]).sort((a, b) => a[1].localeCompare(b[1]));
  [['', 'Plain text']].concat(rows).forEach(([value, label]) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    el.textLang.appendChild(opt);
  });
})();

function fileExt(name) {
  const i = name.lastIndexOf('.');
  return i < 0 ? '' : name.slice(i + 1).toLowerCase();
}

// Which grammar to colour a file with: exact file name first (Dockerfile,
// Makefile, …), then extension, then — only for files with no extension at
// all — a peek at the content (shebang line, <?xml / <html, whole-file JSON).
// null means "plain text".
function detectTextLanguage(name, text) {
  const lower = name.toLowerCase();
  if (TEXT_LANG_BY_NAME.has(lower)) return TEXT_LANG_BY_NAME.get(lower);
  const ext = fileExt(name);
  if (TEXT_LANG_BY_EXT.has(ext)) return TEXT_LANG_BY_EXT.get(ext);
  if (/^\.env(\.|$)/.test(lower)) return 'ini';
  if (/^dockerfile\./.test(lower)) return 'dockerfile';
  return ext ? null : sniffTextLanguage(text);
}

function sniffTextLanguage(text) {
  const head = text.slice(0, 2048);
  const shebang = /^#!\s*(?:\S*\/env\s+(?:-\S+\s+)*)?(\S+)/.exec(head);
  if (shebang) {
    const bin = shebang[1].split('/').pop().toLowerCase();
    if (/^(ba|z|k|da|a)?sh$/.test(bin)) return 'bash';
    if (/^(python|pypy)[\d.]*$/.test(bin)) return 'python';
    if (/^(node|nodejs|deno|bun)$/.test(bin)) return 'javascript';
    if (bin === 'ts-node') return 'typescript';
    if (/^(ruby|perl|php|lua|julia|groovy)$/.test(bin)) return bin;
    if (bin === 'rscript') return 'r';
    if (bin === 'pwsh' || bin === 'powershell') return 'powershell';
    return null;
  }
  const first = head.trimStart();
  if (/^<\?xml\b/i.test(first) || /^<!doctype\s+html/i.test(first) || /^<html[\s>]/i.test(first)) return 'xml';
  if ((first[0] === '{' || first[0] === '[') && text.length <= 262144) {
    try { JSON.parse(text); return 'json'; } catch (err) { /* not JSON after all */ }
  }
  return null;
}

/* ----- highlight.js, loaded lazily (core once; one small grammar per language) ----- */
let hljsCorePromise = null;
const hljsGrammarPromises = new Map();

function ensureHljsCore() {
  if (!hljsCorePromise) {
    hljsCorePromise = loadScriptOnce('lib/hljs/hljs.min.js').catch((err) => {
      hljsCorePromise = null; // allow a retry next time
      throw err;
    });
  }
  return hljsCorePromise;
}

// The core bundle already contains the ~35 most common grammars; the rest
// live one-per-file in lib/hljs/lang/ and are fetched only when needed.
function ensureGrammar(id) {
  if (!hljsGrammarPromises.has(id)) {
    hljsGrammarPromises.set(id,
      ensureHljsCore()
        .then(() => (window.hljs.getLanguage(id) ? undefined : loadScriptOnce(`lib/hljs/lang/${id}.min.js`)))
        .catch((err) => { hljsGrammarPromises.delete(id); throw err; }));
  }
  return hljsGrammarPromises.get(id);
}

// Resolves true when the grammar itself is ready. Companion grammars (see
// TEXT_LANG_DEPS) are best-effort. Never rejects: no colours is not an error.
async function loadTextGrammar(id) {
  const results = await Promise.allSettled([id, ...(TEXT_LANG_DEPS[id] || [])].map(ensureGrammar));
  if (results[0].status !== 'fulfilled') {
    console.warn(`Syntax highlighting unavailable for "${id}" (is lib/hljs/ complete?)`, results[0].reason);
    return false;
  }
  return true;
}

/* ----- Reading bytes: binary check + decoding ----- */
const BINARY_SIGNATURES = [
  [0x25, 0x50, 0x44, 0x46, 0x2D], // %PDF-
  [0x89, 0x50, 0x4E, 0x47],       // PNG
  [0xFF, 0xD8, 0xFF],             // JPEG
  [0x47, 0x49, 0x46, 0x38],       // GIF
  [0x50, 0x4B, 0x03, 0x04],       // ZIP (docx, jar, apk, …)
  [0x1F, 0x8B],                   // gzip
  [0x7F, 0x45, 0x4C, 0x46],       // ELF executable
  [0xD0, 0xCF, 0x11, 0xE0],       // legacy Office (OLE)
];

function looksBinary(bytes) {
  for (const sig of BINARY_SIGNATURES) {
    if (sig.every((b, i) => bytes[i] === b)) return true;
  }
  const n = Math.min(bytes.length, 8192);
  let odd = 0;
  for (let i = 0; i < n; i++) {
    const b = bytes[i];
    if (b === 0) return true; // a NUL byte: text files don't have them
    // control characters other than backspace, tab, LF, VT, FF, CR and ESC (ANSI colours in logs)
    if ((b < 32 && b !== 8 && b !== 9 && b !== 10 && b !== 11 && b !== 12 && b !== 13 && b !== 27) || b === 127) odd++;
  }
  return n > 0 && odd / n > 0.1;
}

// -> { text, encoding, bom }, or null when the bytes aren't text.
function decodeTextBytes(bytes) {
  let start = 0;
  let encoding = 'utf-8';
  let bom = false;
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) { start = 3; bom = true; }
  else if (bytes[0] === 0xFF && bytes[1] === 0xFE) { start = 2; bom = true; encoding = 'utf-16le'; }
  else if (bytes[0] === 0xFE && bytes[1] === 0xFF) { start = 2; bom = true; encoding = 'utf-16be'; }
  const body = bytes.subarray(start);
  try {
    if (encoding !== 'utf-8') {
      return { text: new TextDecoder(encoding, { fatal: true }).decode(body), encoding, bom };
    }
    if (looksBinary(bytes)) return null;
    try {
      return { text: new TextDecoder('utf-8', { fatal: true }).decode(body), encoding, bom };
    } catch (err) {
      // Not valid UTF-8 (and not binary): an old single-byte file. Windows-1252 is what such files nearly always are.
      return { text: new TextDecoder('windows-1252').decode(body), encoding: 'windows-1252', bom: false };
    }
  } catch (err) {
    return null;
  }
}

let win1252Reverse = null;
function windows1252Reverse() {
  if (!win1252Reverse) {
    win1252Reverse = new Map();
    const dec = new TextDecoder('windows-1252');
    for (let b = 0x80; b <= 0x9F; b++) win1252Reverse.set(dec.decode(Uint8Array.of(b)), b);
  }
  return win1252Reverse;
}

// -> Uint8Array, or null when `text` contains something the encoding can't store.
function encodeTextBytes(text, encoding, bom) {
  if (encoding === 'utf-16le' || encoding === 'utf-16be') {
    const little = encoding === 'utf-16le';
    const offset = bom ? 2 : 0;
    const view = new DataView(new ArrayBuffer(offset + text.length * 2));
    if (bom) view.setUint16(0, 0xFEFF, little);
    for (let i = 0; i < text.length; i++) view.setUint16(offset + i * 2, text.charCodeAt(i), little);
    return new Uint8Array(view.buffer);
  }
  if (encoding === 'windows-1252') {
    const reverse = windows1252Reverse();
    const out = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c < 0x80 || (c >= 0xA0 && c <= 0xFF)) { out[i] = c; continue; }
      const b = reverse.get(text[i]);
      if (b === undefined) return null;
      out[i] = b;
    }
    return out;
  }
  const encoded = new TextEncoder().encode(text);
  if (!bom) return encoded;
  const out = new Uint8Array(encoded.length + 3);
  out.set([0xEF, 0xBB, 0xBF]);
  out.set(encoded, 3);
  return out;
}

// The dominant line ending, and whether the file mixed several kinds.
function detectEol(text) {
  let crlf = 0;
  let lf = 0;
  let allCr = 0;
  for (let i = text.indexOf('\n'); i !== -1; i = text.indexOf('\n', i + 1)) {
    if (i > 0 && text.charCodeAt(i - 1) === 13) crlf++; else lf++;
  }
  for (let i = text.indexOf('\r'); i !== -1; i = text.indexOf('\r', i + 1)) allCr++;
  const kinds = [['\r\n', crlf], ['\n', lf], ['\r', allCr - crlf]].sort((a, b) => b[1] - a[1]);
  return { eol: kinds[0][1] > 0 ? kinds[0][0] : '\n', mixed: kinds[1][1] > 0 };
}

// What the Tab key inserts: a tab if the file indents with tabs, otherwise
// the file's usual number of spaces (4 when there's nothing to go on).
function detectIndentUnit(text) {
  let tabs = 0;
  let spaces = 0;
  let prev = 0;
  const steps = new Map();
  for (const line of text.slice(0, 60000).split('\n', 400)) {
    const m = /^([ \t]+)\S/.exec(line);
    if (!m) { if (line.trim() !== '') prev = 0; continue; }
    if (m[1][0] === '\t') { tabs++; continue; }
    spaces++;
    const n = m[1].length;
    const d = Math.abs(n - prev);
    if (d !== 0 && [2, 3, 4, 8].includes(d)) steps.set(d, (steps.get(d) || 0) + 1);
    prev = n;
  }
  if (tabs > spaces) return '\t';
  let best = 4;
  let top = 0;
  steps.forEach((count, d) => { if (count > top) { top = count; best = d; } });
  return ' '.repeat(best);
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(n < 10240 ? 1 : 0)} KB`;
  return `${(n / 1048576).toFixed(1)} MB`;
}

/* ----- Opening ----- */
async function openTextFile(file, name, fileHandle) {
  if (file.size > TEXT_MAX_BYTES) {
    showToast(`"${name}" is ${formatBytes(file.size)} — the text editor opens files up to ${formatBytes(TEXT_MAX_BYTES)}.`, 5000);
    return;
  }
  // Small files open instantly, so the progress card only appears for a big
  // file (or one that turns out to be slow).
  const heavy = file.size > 300 * 1024;
  beginLoading(name, heavy);
  setLoading(10, 'Reading file\u2026');
  try {
    if (heavy) await nextPaint();
    let bytes;
    try {
      bytes = new Uint8Array(await file.arrayBuffer());
    } catch (err) {
      console.error(err);
      showToast('Failed to read file.');
      return;
    }
    setLoading(45, 'Decoding text\u2026');
    if (heavy) await nextPaint();
    const decoded = decodeTextBytes(bytes);
    if (!decoded) {
      showToast(`"${name}" looks like a binary file \u2014 it can\u2019t be opened as text.`, 4500);
      return;
    }
    if (state.dirty && !state.scratchId && !window.confirm('Changes that you made may not be saved.')) return;

    setLoading(70, 'Preparing editor\u2026');
    if (heavy) await nextPaint();
    const { eol, mixed } = detectEol(decoded.text);
    await enterTextMode({
      name,
      text: decoded.text.replace(/\r\n?/g, '\n'), // the textarea works in \n; the original ending is put back on save
      encoding: decoded.encoding,
      bom: decoded.bom,
      eol,
      eolMixed: mixed,
      fileHandle,
    });
  } finally {
    endLoading();
  }
}

async function enterTextMode(doc) {
  const wasScratch = state.scratchId === 'text'; // dropping a file into an open text scratchpad keeps it bound to autosave (see leaveScratchpad's doc comment)
  await leaveScratchpad();
  leaveDesignMode();
  // Tear down whatever spreadsheet is on screen.
  destroyHf();
  state.workbook = null;
  state.sheetNames = [];
  state.activeSheet = null;
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingOriginal = null;
  state.editingTd = null;
  state.formulaPick = null;
  state.formulaHomeSheet = null;
  state.formulaHomeCell = null;
  state.isCSV = false;
  state.clipEnabled = false;
  el.clipBtn.classList.remove('active');
  state.undoStack = [];
  state.redoStack = [];
  state.fileHandle = doc.fileHandle || null;
  closeSheetMenu();
  el.grid.innerHTML = '';
  el.sheetTabs.innerHTML = '';
  updateSelectionStats();
  updateDeleteBtn();
  el.addRowBtn.disabled = true;
  el.addColBtn.disabled = true;

  const lang = detectTextLanguage(doc.name, doc.text);
  const t = {
    encoding: doc.encoding,
    bom: doc.bom,
    eol: doc.eol,
    eolMixed: doc.eolMixed,
    lang: null,
    langManual: false,
    langToken: 0,
    wrap: !lang || lang === 'markdown', // prose wraps; code scrolls sideways with line numbers
    plain: doc.text.length > TEXT_OVERLAY_MAX,
    indent: detectIndentUnit(doc.text),
    tabEscape: false,
    hlTimer: 0,
    caretRaf: 0,
    gutterLines: 0,
  };
  state.text = t;
  state.mode = 'text';
  state.dirty = false;
  state.originalFileName = doc.name;
  state.title = doc.name;
  document.body.classList.add('mode-text');
  document.body.classList.remove('mode-home');

  el.fileName.textContent = doc.name;
  el.renameTitleBtn.hidden = false;
  el.saveBtn.disabled = false;
  el.undoBtn.disabled = false;
  el.redoBtn.disabled = false;

  el.textHlCode.textContent = '';
  el.textGutter.textContent = '';
  el.textInput.value = doc.text;
  el.textScroll.scrollTop = 0;
  el.textScroll.scrollLeft = 0;
  applyTextViewFlags();
  updateTextEncodingLabel();
  updateTextEolLabel();
  renderText();
  setTextLanguage(lang, false); // colours arrive as soon as the grammar file has loaded
  el.textInput.setSelectionRange(0, 0);
  el.textInput.focus({ preventScroll: true });
  showToast(doc.toast || `Opened "${doc.name}"`);
  updateBookmarkStar();
  if (wasScratch) {
    state.scratchId = 'text';
    // A file dropped into an already-open scratchpad is tagged scratchpad_<name> so it's
    // clear the scratch slot now holds that file (Clear resets the title back to default).
    state.title = `scratchpad_${state.title}`;
    el.fileName.textContent = state.title;
    scheduleScratchAutosave();
  }
}

// Back out of the text editor (called when a spreadsheet, or the home
// screen, is about to take over). The caller sets the rest of the UI.
function leaveTextMode() {
  if (state.mode !== 'text') return;
  const t = state.text;
  if (t) { clearTimeout(t.hlTimer); t.langToken++; }
  state.text = null;
  state.mode = 'home';
  document.body.classList.remove('mode-text');
  el.textWrapper.classList.remove('is-wrap', 'is-plain');
  el.textInput.value = '';
  el.textHlCode.textContent = '';
  el.textGutter.textContent = '';
}

/* ----- Painting ----- */
function applyTextViewFlags() {
  const t = state.text;
  el.textWrapper.classList.toggle('is-wrap', !!t && t.wrap);
  el.textWrapper.classList.toggle('is-plain', !!t && t.plain);
  el.textWrapBtn.setAttribute('aria-pressed', String(!!t && t.wrap));
}

function countLines(v) {
  let n = 1;
  for (let i = v.indexOf('\n'); i !== -1; i = v.indexOf('\n', i + 1)) n++;
  return n;
}

// Repaints everything that depends on the text (coloured copy, line numbers,
// status bar). Called after every edit.
function renderText() {
  const t = state.text;
  if (!t) return;
  const v = el.textInput.value;
  const lines = countLines(v);
  const plain = v.length > TEXT_OVERLAY_MAX;
  if (plain !== t.plain) {
    t.plain = plain;
    t.gutterLines = 0;
    applyTextViewFlags();
    if (plain) { el.textHlCode.textContent = ''; el.textGutter.textContent = ''; }
  }
  if (!plain) {
    paintHighlight(v, false);
    if (lines !== t.gutterLines) {
      t.gutterLines = lines;
      const nums = new Array(lines);
      for (let i = 0; i < lines; i++) nums[i] = i + 1;
      el.textGutter.textContent = nums.join('\n');
    }
  }
  el.textCount.textContent = `${lines.toLocaleString()} line${lines === 1 ? '' : 's'} \u00b7 ${v.length.toLocaleString()} chars`;
  updateTextCaret();
}

// Puts the (coloured, when possible) copy of the text in the layer above the
// textarea. `allowSlow` lets it colour a big document right now; otherwise a
// big document shows uncoloured text at once (so nothing typed is ever
// missing) and is coloured after typing pauses.
function paintHighlight(v, allowSlow) {
  const t = state.text;
  clearTimeout(t.hlTimer);
  const canColour = !!t.lang && v.length <= TEXT_HL_MAX && !!window.hljs && !!hljs.getLanguage(t.lang);
  if (canColour && (allowSlow || v.length <= TEXT_HL_SYNC_MAX)) {
    let html = null;
    try { html = hljs.highlight(v, { language: t.lang, ignoreIllegals: true }).value; } catch (err) { /* fall through to plain */ }
    if (html !== null) {
      // A trailing "\n" is what gives the copy the same height as a textarea whose text ends in a line break.
      el.textHlCode.innerHTML = html + '\n';
      return;
    }
  }
  el.textHlCode.textContent = v + '\n';
  if (canColour) {
    t.hlTimer = setTimeout(() => {
      if (state.text === t && !t.plain) paintHighlight(el.textInput.value, true);
    }, 140);
  }
}

async function setTextLanguage(id, manual) {
  const t = state.text;
  if (!t) return;
  t.lang = id || null;
  if (manual) t.langManual = true;
  el.textLang.value = t.lang || '';
  const token = ++t.langToken;
  if (!t.plain) paintHighlight(el.textInput.value, true); // plain, or coloured if the grammar is already loaded
  if (!t.lang) return;
  const ok = await loadTextGrammar(t.lang);
  if (ok && state.text === t && t.langToken === token && !t.plain) paintHighlight(el.textInput.value, true);
}

function updateTextCaret() {
  if (!state.text) return;
  const ta = el.textInput;
  const v = ta.value;
  const a = ta.selectionStart;
  const b = ta.selectionEnd;
  const pos = ta.selectionDirection === 'backward' ? a : b;
  let line = 1;
  let lastBreak = -1;
  for (let i = v.indexOf('\n'); i !== -1 && i < pos; i = v.indexOf('\n', i + 1)) { line++; lastBreak = i; }
  el.textPos.textContent = `Ln ${line.toLocaleString()}, Col ${(pos - lastBreak).toLocaleString()}` +
    (b > a ? ` (${(b - a).toLocaleString()} selected)` : '');
}

function scheduleTextCaret() {
  const t = state.text;
  if (!t || t.caretRaf) return;
  t.caretRaf = requestAnimationFrame(() => { t.caretRaf = 0; updateTextCaret(); });
}

const TEXT_ENCODING_LABELS = { 'utf-8': 'UTF-8', 'utf-16le': 'UTF-16 LE', 'utf-16be': 'UTF-16 BE', 'windows-1252': 'Windows-1252' };

function updateTextEncodingLabel() {
  const t = state.text;
  if (!t) return;
  el.textEnc.textContent = (TEXT_ENCODING_LABELS[t.encoding] || t.encoding) + (t.bom && t.encoding === 'utf-8' ? ' BOM' : '');
}

function updateTextEolLabel() {
  const t = state.text;
  if (!t) return;
  const name = t.eol === '\r\n' ? 'CRLF' : t.eol === '\r' ? 'CR' : 'LF';
  el.textEolBtn.textContent = t.eolMixed ? `Mixed \u2192 ${name}` : name;
  el.textEolBtn.title = t.eolMixed
    ? `This file mixed several kinds of line ending; saving makes them all ${name}. Click to switch.`
    : 'Line endings used when saving (click to switch)';
}

/* ----- Editing ----- */
// Replaces a range through execCommand so the browser's own undo history
// (Ctrl+Z, the Undo button) keeps working across the change.
function replaceTextRange(start, end, str) {
  const ta = el.textInput;
  ta.focus({ preventScroll: true });
  ta.setSelectionRange(start, end);
  if (!document.execCommand('insertText', false, str)) {
    ta.setRangeText(str, start, end, 'end');
    ta.dispatchEvent(new Event('input', { bubbles: true }));
  }
}

function stripIndent(line, unit) {
  if (line[0] === '\t') return line.slice(1);
  const max = unit === '\t' ? 4 : unit.length;
  let n = 0;
  while (n < max && line[n] === ' ') n++;
  return line.slice(n);
}

// Tab / Shift+Tab: indent or outdent every line the selection touches (or
// just insert one indent at the caret when nothing is selected).
function indentSelection(outdent) {
  const t = state.text;
  const ta = el.textInput;
  const v = ta.value;
  const s = ta.selectionStart;
  const e = ta.selectionEnd;
  if (s === e && !outdent) { replaceTextRange(s, e, t.indent); return; }

  const from = s === 0 ? 0 : v.lastIndexOf('\n', s - 1) + 1;
  // A selection that ends right after a line break doesn't reach into the next line.
  const stop = e > s && v[e - 1] === '\n' ? e - 1 : e;
  const nl = v.indexOf('\n', stop);
  const to = nl === -1 ? v.length : nl;
  const block = v.slice(from, to);
  const lines = block.split('\n');
  const out = lines.map(outdent ? (l) => stripIndent(l, t.indent) : (l) => (l === '' ? l : t.indent + l));
  const replaced = out.join('\n');
  if (replaced === block) return;

  const firstDelta = out[0].length - lines[0].length;
  const totalDelta = replaced.length - block.length;
  replaceTextRange(from, to, replaced);
  const newStart = Math.max(from, s + firstDelta);
  ta.setSelectionRange(newStart, Math.max(newStart, e + totalDelta));
}

function textUndoRedo(kind) {
  el.textInput.focus({ preventScroll: true });
  document.execCommand(kind);
}

/* ----- Clear, and Jump to row ----- */
// Clearing goes through the browser's own edit command, so Ctrl+Z / Undo restores everything.
function clearTextDocument() {
  if (!state.text) return;
  const ta = el.textInput;
  if (!ta.value) { showToast('Already empty', 1600); return; }
  ta.focus({ preventScroll: true });
  ta.setSelectionRange(0, ta.value.length);
  if (!document.execCommand('delete')) {
    ta.setRangeText('', 0, ta.value.length, 'end');
    ta.dispatchEvent(new Event('input', { bubbles: true }));
  }
  el.textScroll.scrollTop = 0;
  el.textScroll.scrollLeft = 0;
  // Clearing the scratchpad also drops any "scratchpad_<dropped file>" name it picked up,
  // back to the plain default (not for a regular text file opened outside the scratchpad).
  if (state.scratchId === 'text' && state.title !== SCRATCH_DEFAULT_NAMES.text) {
    state.title = SCRATCH_DEFAULT_NAMES.text;
    state.originalFileName = SCRATCH_DEFAULT_NAMES.text;
    el.fileName.textContent = state.title;
  }
  showToast('Cleared \u2014 Ctrl+Z brings it back', 2600);
}
el.clearTextBtn.addEventListener('click', clearTextDocument);

// Where a line starts, in pixels from the top of the document. Unwrapped text has a fixed row height; wrapped
// text is measured with a hidden copy of the textarea's layout (so long lines that wrap are accounted for).
function textLineMetrics(lineStartIdx, lineEndIdx, lineNo) {
  const ta = el.textInput;
  const cs = getComputedStyle(ta);
  const padTop = parseFloat(cs.paddingTop) || 0;
  const t = state.text;
  if (!t.wrap || lineStartIdx > 2000000) return { top: padTop + (lineNo - 1) * TEXT_LINE_PX, height: TEXT_LINE_PX };
  const m = document.createElement('div');
  for (const p of ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'letterSpacing', 'lineHeight', 'tabSize', 'whiteSpace', 'overflowWrap', 'wordBreak', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft', 'boxSizing']) m.style[p] = cs[p];
  m.style.cssText += ';position:absolute;visibility:hidden;left:-99999px;top:0;';
  m.style.width = `${ta.clientWidth}px`;
  const v = ta.value;
  const measure = (idx) => {
    m.textContent = v.slice(0, idx);
    const mark = document.createElement('span');
    mark.textContent = '\u200b';
    m.appendChild(mark);
    return mark.offsetTop;
  };
  document.body.appendChild(m);
  const top = measure(lineStartIdx);
  const bottom = measure(lineEndIdx);
  m.remove();
  return { top, height: Math.max(TEXT_LINE_PX, bottom - top + TEXT_LINE_PX) };
}

function jumpToTextRow(n) {
  const t = state.text;
  if (!t) return false;
  const ta = el.textInput;
  const v = ta.value;
  const total = countLines(v);
  const wanted = Math.floor(Number(n));
  if (!Number.isFinite(wanted) || wanted < 1) return false;
  const line = Math.min(wanted, total);
  let start = 0;
  for (let i = 1; i < line; i++) start = v.indexOf('\n', start) + 1;
  let end = v.indexOf('\n', start);
  if (end === -1) end = v.length;
  ta.focus({ preventScroll: true });
  ta.setSelectionRange(start, start);
  const { top, height } = textLineMetrics(start, end, line);
  const box = t.plain ? ta : el.textScroll;
  box.scrollTop = Math.max(0, Math.min(box.scrollHeight - box.clientHeight, top + height / 2 - box.clientHeight / 2));
  if (!t.wrap) { box.scrollLeft = 0; }
  // a brief highlight so the eye finds the row (not in the huge-file mode, where the copy layer is off)
  document.querySelectorAll('.text-jump-flash').forEach((x) => x.remove());
  if (!t.plain) {
    const flash = document.createElement('div');
    flash.className = 'text-jump-flash';
    flash.style.top = `${top}px`;
    flash.style.height = `${height}px`;
    document.querySelector('.text-stack').appendChild(flash);
    setTimeout(() => flash.remove(), 1700);
  }
  updateTextCaret();
  if (wanted > total) showToast(`Only ${total.toLocaleString()} row${total === 1 ? '' : 's'} \u2014 jumped to the last one`, 2400);
  return true;
}

el.textJumpForm.addEventListener('submit', (e) => {
  e.preventDefault();
  if (!state.text) return;
  const raw = el.textJump.value.trim().replace(/[,\s_]/g, '');
  if (!/^\d+$/.test(raw) || Number(raw) < 1) {
    el.textJump.classList.remove('is-invalid');
    void el.textJump.offsetWidth; // restart the shake
    el.textJump.classList.add('is-invalid');
    showToast(`Type a row number from 1 to ${countLines(el.textInput.value).toLocaleString()}`, 2400);
    return;
  }
  el.textJump.classList.remove('is-invalid');
  jumpToTextRow(raw);
  el.textJump.value = '';
});
el.textJump.addEventListener('input', () => el.textJump.classList.remove('is-invalid'));
el.textJump.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { e.preventDefault(); el.textJump.value = ''; el.textInput.focus({ preventScroll: true }); }
});
function focusTextJump() { el.textJump.focus(); el.textJump.select(); }
el.textPos.addEventListener('click', focusTextJump);

el.textInput.addEventListener('input', () => {
  if (!state.text) return;
  state.dirty = true;
  renderText();
  scheduleScratchAutosave();
});

el.textInput.addEventListener('keydown', (e) => {
  const t = state.text;
  if (!t) return;
  if (e.key === 'Shift' || e.key === 'Control' || e.key === 'Alt' || e.key === 'Meta') return;
  if (e.key === 'Escape') { t.tabEscape = true; return; } // Esc, then Tab: leave the editor (keyboard-trap escape hatch)
  if ((e.key === 'PageUp' || e.key === 'PageDown') && !e.ctrlKey && !e.metaKey && !e.altKey && !t.plain) {
    // The textarea is as tall as the whole document (the wrapper scrolls, not
    // it), so the native "page" would be the entire file. Move the caret by
    // one window's worth of rows instead, and scroll the same distance.
    e.preventDefault();
    const down = e.key === 'PageDown';
    const rows = Math.max(1, Math.floor(el.textScroll.clientHeight / TEXT_LINE_PX) - 1);
    const sel = window.getSelection();
    for (let i = 0; i < rows; i++) sel.modify(e.shiftKey ? 'extend' : 'move', down ? 'forward' : 'backward', 'line');
    el.textScroll.scrollTop += (down ? rows : -rows) * TEXT_LINE_PX;
    scheduleTextCaret();
    return;
  }
  if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey && !e.isComposing) {
    if (t.tabEscape) { t.tabEscape = false; return; }
    e.preventDefault();
    indentSelection(e.shiftKey);
    return;
  }
  t.tabEscape = false;
});

['keyup', 'mouseup', 'focus', 'select'].forEach((evt) => el.textInput.addEventListener(evt, scheduleTextCaret));

el.textWrapBtn.addEventListener('click', () => {
  const t = state.text;
  if (!t) return;
  t.wrap = !t.wrap;
  applyTextViewFlags();
  el.textInput.focus({ preventScroll: true });
});

el.textEolBtn.addEventListener('click', () => {
  const t = state.text;
  if (!t) return;
  t.eol = t.eol === '\r\n' ? '\n' : '\r\n';
  t.eolMixed = false;
  state.dirty = true;
  updateTextEolLabel();
  el.textInput.focus({ preventScroll: true });
});

el.textLang.addEventListener('change', () => {
  setTextLanguage(el.textLang.value || null, true);
  el.textInput.focus({ preventScroll: true });
});

// Dropping a file on the editor opens it (instead of Chrome navigating away
// to the file and losing the edits). Dragging plain text around is untouched.
const dragHasFiles = (e) => !!e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');

// Best-effort: if the browser exposes a real FileSystemFileHandle for the
// first dropped item, grab it (this is what lets a dropped file show the
// bookmark star). Anything that goes wrong \u2014 unsupported browser, a drag
// source that doesn't expose a handle, permission trouble \u2014 just resolves to
// null and the drop behaves exactly as it did before this feature existed.
async function tryGetDropHandle(item) {
  try {
    if (item && typeof item.getAsFileSystemHandle === 'function') {
      const handle = await item.getAsFileSystemHandle();
      if (handle && handle.kind === 'file') return handle;
    }
  } catch (err) { /* not supported here \u2014 fall through to no handle */ }
  return null;
}

el.textWrapper.addEventListener('dragover', (e) => { if (dragHasFiles(e)) e.preventDefault(); });
el.textWrapper.addEventListener('drop', (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const item = e.dataTransfer.items && e.dataTransfer.items[0];
  tryGetDropHandle(item).then((handle) => loadFile(file, undefined, handle));
});

document.addEventListener('keydown', (e) => {
  if (state.mode !== 'text') return;
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 's') {
    e.preventDefault();
    saveTextFile();
  }
});

/* ----- Saving ----- */
// "notes.txt" -> "notes_edited.txt", the same naming the spreadsheet export uses.
function editedFileName(name) {
  const m = /^(.+)(\.[^./\\]+)$/.exec(name);
  return m ? `${m[1]}_edited${m[2]}` : `${name}_edited`;
}

function saveTextFile() {
  const t = state.text;
  if (!t) return;
  let text = el.textInput.value;
  if (t.eol !== '\n') text = text.replace(/\n/g, t.eol);

  let bytes = encodeTextBytes(text, t.encoding, t.bom);
  let note = '';
  if (!bytes) {
    // Something was typed that Windows-1252 has no character for.
    t.encoding = 'utf-8';
    t.bom = false;
    updateTextEncodingLabel();
    bytes = encodeTextBytes(text, 'utf-8', false);
    note = ' as UTF-8 (it now has characters Windows-1252 can\u2019t store)';
  }
  const outName = editedFileName(state.title || 'untitled.txt');
  // octet-stream, not text/plain: Chrome tacks ".txt" onto extensionless downloads (Dockerfile, Makefile, …) of the latter.
  downloadBlob(new Blob([bytes], { type: 'application/octet-stream' }), outName);
  state.dirty = false;
  showToast(`Saved "${outName}"${note}`, note ? 7000 : 2200);
}

/* ------------------------------------------------------------------ *
 * "New" — back to the empty/home state
 * ------------------------------------------------------------------ */
async function goToHome() {
  if (state.dirty && !state.scratchId) {
    const ok = window.confirm('Changes that you made may not be saved.');
    if (!ok) return;
  }
  await leaveScratchpad();

  leaveDesignMode();
  leaveTextMode();
  state.mode = 'home';
  destroyHf();
  state.workbook = null;
  state.sheetNames = [];
  state.activeSheet = null;
  state.originalFileName = '';
  state.title = '';
  state.isCSV = false;
  setExportFormat('xlsx');
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingOriginal = null;
  state.editingTd = null;
  state.dirty = false;
  state.clipEnabled = false;
  el.clipBtn.classList.remove('active');
  state.undoStack = [];
  state.redoStack = [];
  state.fileHandle = null;

  el.fileName.textContent = 'No file loaded';
  el.renameTitleBtn.hidden = true;
  el.bookmarkStarBtn.hidden = true;
  el.saveBtn.disabled = true;
  el.clipBtn.disabled = true;
  el.clipBtn.classList.remove('active');
  el.clearSheetBtn.disabled = true;
  hideClearSheetForm();
  // Keep the exact open-mode shell visible on the home/new page.
  // Only the sheet area changes its content to the file-drop prompt.
  el.formulaBar.hidden = false;
  el.formulaBarInput.disabled = true;
  el.formulaBarInput.value = '';
  el.formulaBarRef.textContent = 'A1';

  el.gridWrapper.hidden = false;
  el.gridWrapper.classList.add('home-empty');
  document.body.classList.add('mode-home');
  el.gridScroll.hidden = true;
  el.dropzone.hidden = false;
  el.floatNav.hidden = false;

  // The bottom bar only holds sheet tabs, so the homepage doesn't draw it.
  el.sheetTabs.hidden = true;
  el.sheetTabs.innerHTML = '';
  closeSheetMenu();

  updateSelectionStats(); // no workbook/selection anymore — clear the stale sum/avg/etc.
  updateUndoRedoButtons();
  updateDeleteBtn();
  el.addRowBtn.disabled = true; // these follow the selection, and there's none any more
  el.addColBtn.disabled = true;
  showToast('Ready for a new file');
  refreshScratchTiles(); // the tiles were hidden while a file was open — make sure they're current
  el.homeSearch.value = '';
  refreshBookmarkPanels();
}

/* ------------------------------------------------------------------ *
 * Rendering
 * ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ *
 * Sheet order helpers (add-after-current, drag to reorder, delete)
 * ------------------------------------------------------------------ */
// Workbook-level metadata is index-based: wb.Workbook.Sheets[i] holds sheet
// i's visibility, and a defined name's `.Sheet` is the index of the sheet it
// is scoped to. Whenever sheets are added / moved / removed, re-key both by
// sheet *name* so they keep pointing at the right sheets.
function remapWorkbookSheetIndices(wb, oldOrder, newOrder) {
  const meta = wb && wb.Workbook;
  if (!meta) return;
  if (Array.isArray(meta.Sheets)) {
    const byName = new Map();
    oldOrder.forEach((n, i) => { if (meta.Sheets[i]) byName.set(n, meta.Sheets[i]); });
    meta.Sheets = newOrder.map((n) => byName.get(n) || {});
  }
  if (Array.isArray(meta.Names)) {
    meta.Names = meta.Names.filter((nm) => {
      if (nm.Sheet === undefined || nm.Sheet === null) return true;
      const idx = newOrder.indexOf(oldOrder[nm.Sheet]);
      if (idx < 0) return false; // scoped to a sheet that no longer exists
      nm.Sheet = idx;
      return true;
    });
  }
}

// Switches to another sheet (used by the sheet-manager popup). Same rules
// as clicking its tab, including not disturbing a formula that's mid-edit.
function goToSheet(name) {
  if (!state.workbook || !state.sheetNames.includes(name)) return;
  if (state.activeSheet === name) return;
  if (isFormulaPointModeActive()) { switchSheetForFormulaBrowse(name); return; }
  state.activeSheet = name;
  state.rangeSel = null;
  state.headerAnchor = null;
  renderSheetTabs();
  renderSheet();
}

// Commits a full reordering of the sheets in one go (used by the
// sheet-manager popup's drag-to-reorder).
function applySheetOrder(newOrder) {
  if (!state.workbook) return;
  if (newOrder.length !== state.sheetNames.length) return; // sanity guard
  if (newOrder.join('\u0000') === state.sheetNames.join('\u0000')) return; // dropped where it already was

  pushUndo();
  const wb = state.workbook;
  const oldOrder = wb.SheetNames.slice();
  wb.SheetNames = newOrder.slice();
  state.sheetNames = newOrder.slice();
  remapWorkbookSheetIndices(wb, oldOrder, wb.SheetNames);
  state.dirty = true;
  recalcFormulas();
  renderSheetTabs();
}

const PENCIL_SVG = '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20l4.2-1 9.9-9.9a2 2 0 0 0-2.8-2.8L5.4 16.2z"/><path d="M13.8 7.2l3 3"/></svg>';
const PLUS_SVG = '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>';
const LIST_SVG = '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 7h14M5 12h14M5 17h14"/></svg>';
const TRASH_SVG = '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 7h14M9 7V4.5h6V7M7 7l1 13h8l1-13M10 10.5v6M14 10.5v6"/></svg>';
const GRIP_SVG = '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="6" r="1.3"/><circle cx="15" cy="6" r="1.3"/><circle cx="9" cy="12" r="1.3"/><circle cx="15" cy="12" r="1.3"/><circle cx="9" cy="18" r="1.3"/><circle cx="15" cy="18" r="1.3"/></svg>';

/* ------------------------------------------------------------------ *
 * Sheet manager popup (opened from the ≡ button in the tab bar).
 * Renaming, deleting and drag-to-reorder for every sheet all live here;
 * the tab bar strip itself is click-to-switch only, plus "+" to add.
 * ------------------------------------------------------------------ */
let sheetMenu = null;

function closeSheetMenu() {
  if (!sheetMenu) return;
  sheetMenu.remove();
  sheetMenu = null;
  document.removeEventListener('mousedown', onSheetMenuOutside, true);
  document.removeEventListener('keydown', onSheetMenuKey, true);
  window.removeEventListener('resize', closeSheetMenu);
  el.sheetTabs.querySelectorAll('.sheet-list-btn').forEach((b) => {
    b.classList.remove('active');
    b.setAttribute('aria-expanded', 'false');
  });
}

function onSheetMenuOutside(e) {
  if (sheetMenu && !sheetMenu.contains(e.target) && !e.target.closest('.sheet-list-btn')) closeSheetMenu();
}

function onSheetMenuKey(e) {
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeSheetMenu(); }
}

// Rebuilds the popup's rows in place (after a rename / delete / reorder /
// add-sheet elsewhere) without closing it, so managing several sheets in a
// row stays fluid instead of the popup vanishing after each action.
function refreshSheetMenu() {
  if (!sheetMenu) return;
  const list = sheetMenu.querySelector('.sheet-manager-list');
  if (!list) return;
  const prevScroll = list.scrollTop;
  list.innerHTML = '';
  state.sheetNames.forEach((name) => list.appendChild(buildSheetRow(name)));
  list.scrollTop = prevScroll;
}

function toggleSheetMenu(btn) {
  if (sheetMenu) { closeSheetMenu(); return; }
  if (!state.sheetNames.length) return;

  const menu = document.createElement('div');
  menu.className = 'sheet-manager';
  menu.setAttribute('role', 'menu');

  const list = document.createElement('div');
  list.className = 'sheet-manager-list';
  state.sheetNames.forEach((name) => list.appendChild(buildSheetRow(name)));
  menu.appendChild(list);

  document.body.appendChild(menu);

  // The bar sits at the bottom of the window, so open upwards from the button.
  const r = btn.getBoundingClientRect();
  const left = Math.max(8, Math.min(r.left, window.innerWidth - menu.offsetWidth - 8));
  menu.style.left = left + 'px';
  menu.style.bottom = Math.max(8, window.innerHeight - r.top + 6) + 'px';
  const activeRow = list.querySelector('.sheet-row.active');
  if (activeRow && typeof activeRow.scrollIntoView === 'function') activeRow.scrollIntoView({ block: 'nearest' });

  sheetMenu = menu;
  btn.classList.add('active');
  btn.setAttribute('aria-expanded', 'true');
  document.addEventListener('mousedown', onSheetMenuOutside, true);
  document.addEventListener('keydown', onSheetMenuKey, true);
  window.addEventListener('resize', closeSheetMenu);
}

// One row of the sheet manager: drag handle, name (click to switch / or
// rename when the pencil is active), rename button, delete button.
function buildSheetRow(name) {
  const row = document.createElement('div');
  row.className = 'sheet-row' + (name === state.activeSheet ? ' active' : '');
  row.dataset.sheet = name;
  row.setAttribute('role', 'menuitem');

  const handle = document.createElement('button');
  handle.type = 'button';
  handle.className = 'sheet-row-handle';
  handle.title = 'Drag to reorder';
  handle.setAttribute('aria-label', `Reorder ${name}`);
  handle.innerHTML = GRIP_SVG;
  row.appendChild(handle);

  const label = document.createElement('span');
  label.className = 'sheet-row-label';
  label.textContent = name;
  label.title = name;
  row.appendChild(label);

  const editBtn = document.createElement('button');
  editBtn.type = 'button';
  editBtn.className = 'sheet-row-edit';
  editBtn.title = 'Rename sheet';
  editBtn.setAttribute('aria-label', `Rename ${name}`);
  editBtn.innerHTML = PENCIL_SVG;
  editBtn.addEventListener('mousedown', (e) => e.preventDefault());
  editBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    startRowRename(row, label, name);
  });
  row.appendChild(editBtn);

  const delBtn = document.createElement('button');
  delBtn.type = 'button';
  delBtn.className = 'sheet-row-delete';
  delBtn.title = state.sheetNames.length <= 1 ? "Can't delete the only sheet" : 'Delete sheet';
  delBtn.setAttribute('aria-label', `Delete ${name}`);
  delBtn.innerHTML = TRASH_SVG;
  delBtn.disabled = state.sheetNames.length <= 1;
  delBtn.addEventListener('mousedown', (e) => e.preventDefault());
  delBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    deleteSheets([name]);
  });
  row.appendChild(delBtn);

  // Clicking anywhere on the row other than its buttons/handle switches to
  // that sheet and closes the popup — same as clicking a tab.
  row.addEventListener('mousedown', (e) => {
    if (e.target.closest('.sheet-row-handle')) return; // that starts a drag instead
    if (isFormulaPointModeActive()) e.preventDefault();
  });
  row.addEventListener('click', (e) => {
    if (e.target.closest('.sheet-row-edit') || e.target.closest('.sheet-row-delete') || e.target.closest('.sheet-row-handle')) return;
    if (label.isContentEditable) return;
    closeSheetMenu();
    goToSheet(name);
  });

  attachRowDrag(row, handle);
  return row;
}

function startRowRename(rowEl, labelEl, currentName) {
  labelEl.contentEditable = 'true';
  labelEl.spellcheck = false;
  rowEl.classList.add('editing');
  labelEl.focus();
  selectAllText(labelEl);

  function finish(commit) {
    labelEl.removeEventListener('keydown', onKeydown);
    labelEl.removeEventListener('blur', onBlur);
    labelEl.contentEditable = 'false';
    rowEl.classList.remove('editing');
    if (commit) {
      renameSheet(currentName, labelEl.textContent.trim());
    } else {
      labelEl.textContent = currentName;
    }
  }
  function onKeydown(e) {
    if (e.key === 'Enter') { e.preventDefault(); finish(true); }
    else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
  }
  function onBlur() { finish(true); }

  labelEl.addEventListener('keydown', onKeydown);
  labelEl.addEventListener('blur', onBlur);
}

// Smooth pointer-driven drag-to-reorder for a sheet-manager row. The row
// floats with the pointer (position: fixed) while a placeholder holds its
// slot in the list, so the rest of the rows reflow live as it passes over
// them — no native HTML5 drag-and-drop jank.
function attachRowDrag(row, handle) {
  handle.addEventListener('pointerdown', (e) => {
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    e.preventDefault();
    const list = row.parentElement;
    if (!list) return;

    const startRect = row.getBoundingClientRect();
    const listRect = list.getBoundingClientRect();

    const placeholder = document.createElement('div');
    placeholder.className = 'sheet-row-placeholder';
    placeholder.style.height = startRect.height + 'px';
    list.insertBefore(placeholder, row);

    row.classList.add('dragging');
    row.style.width = startRect.width + 'px';
    row.style.left = startRect.left + 'px';
    row.style.top = startRect.top + 'px';
    document.body.appendChild(row); // float above the list, out of its scroll/clip

    const pointerStartY = e.clientY;
    const rowStartTop = startRect.top;
    let lastClientY = e.clientY;
    try { handle.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }

    let scrollRAF = requestAnimationFrame(autoScroll);
    function autoScroll() {
      const r = list.getBoundingClientRect();
      const edge = 26;
      if (lastClientY < r.top + edge && list.scrollTop > 0) list.scrollTop -= 9;
      else if (lastClientY > r.bottom - edge && list.scrollTop < list.scrollHeight - list.clientHeight) list.scrollTop += 9;
      scrollRAF = requestAnimationFrame(autoScroll);
    }

    function onMove(ev) {
      lastClientY = ev.clientY;
      let top = rowStartTop + (ev.clientY - pointerStartY);
      top = Math.max(listRect.top - startRect.height * 0.4, Math.min(top, listRect.bottom - startRect.height * 0.6));
      row.style.top = top + 'px';

      const centerY = top + startRect.height / 2;
      const siblings = Array.from(list.children).filter((c) => c !== placeholder);
      let target = null;
      let after = false;
      for (const sib of siblings) {
        const sr = sib.getBoundingClientRect();
        if (centerY < sr.top + sr.height / 2) { target = sib; after = false; break; }
      }
      if (!target && siblings.length) { target = siblings[siblings.length - 1]; after = true; }
      if (target) {
        const before = after ? target.nextSibling : target;
        if (placeholder.nextSibling !== before) list.insertBefore(placeholder, before);
      }
    }

    function onUp(ev) {
      try { handle.releasePointerCapture(ev.pointerId); } catch (err) { /* ignore */ }
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', onUp);
      document.removeEventListener('pointercancel', onUp);
      cancelAnimationFrame(scrollRAF);

      placeholder.replaceWith(row);
      row.classList.remove('dragging');
      row.style.width = '';
      row.style.left = '';
      row.style.top = '';

      const newOrder = Array.from(list.querySelectorAll('.sheet-row')).map((r) => r.dataset.sheet);
      applySheetOrder(newOrder);
      // applySheetOrder -> renderSheetTabs -> refreshSheetMenu rebuilds the
      // rows from state either way, so this DOM is just a clean fallback.
    }

    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', onUp);
    document.addEventListener('pointercancel', onUp);
  });
}

function buildTabTools() {
  const wrap = document.createElement('div');
  wrap.className = 'sheet-tab-tools';

  const mkTool = (cls, svg, title) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'sheet-tab-tool ' + cls;
    b.title = title;
    b.setAttribute('aria-label', title);
    b.innerHTML = svg;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    return b;
  };

  const addBtn = mkTool('sheet-add-btn', PLUS_SVG, 'Add a new sheet after the current one');
  addBtn.addEventListener('click', addSheet);

  const listBtn = mkTool('sheet-list-btn', LIST_SVG, 'Manage sheets — rename, delete, reorder');
  listBtn.setAttribute('aria-haspopup', 'menu');
  listBtn.setAttribute('aria-expanded', 'false');
  listBtn.addEventListener('click', () => toggleSheetMenu(listBtn));

  wrap.appendChild(addBtn);
  wrap.appendChild(listBtn);
  return wrap;
}

// Brings the active tab fully into view inside the (scrollbar-less) tab strip.
function revealActiveTab(scroller) {
  const active = scroller.querySelector('.sheet-tab.active');
  if (!active) return;
  const s = scroller.getBoundingClientRect();
  const a = active.getBoundingClientRect();
  if (a.left < s.left) scroller.scrollLeft -= (s.left - a.left) + 8;
  else if (a.right > s.right) scroller.scrollLeft += (a.right - s.right) + 8;
}

// ‹ › buttons, shown only while the tabs overflow the bar.
function buildTabScrollArrows(scroller) {
  const wrap = document.createElement('div');
  wrap.className = 'sheet-tab-arrows';
  wrap.hidden = true;
  const mk = (dir, svg, title) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'sheet-tab-arrow';
    b.dataset.dir = String(dir);
    b.title = title;
    b.setAttribute('aria-label', title);
    b.innerHTML = svg;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => {
      const delta = dir * Math.max(120, Math.round(scroller.clientWidth * 0.6));
      if (typeof scroller.scrollBy === 'function') scroller.scrollBy({ left: delta, behavior: 'smooth' });
      else scroller.scrollLeft += delta;
    });
    return b;
  };
  wrap.appendChild(mk(-1, '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M14.5 6l-6 6 6 6"/></svg>', 'Scroll tabs left'));
  wrap.appendChild(mk(1, '<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9.5 6l6 6-6 6"/></svg>', 'Scroll tabs right'));
  return wrap;
}

function updateTabScrollUI() {
  const s = el.sheetTabs.querySelector('.sheet-tabs-scroll');
  const arrows = el.sheetTabs.querySelector('.sheet-tab-arrows');
  if (!s || !arrows) return;
  arrows.hidden = s.scrollWidth <= s.clientWidth + 1;
  const left = arrows.querySelector('[data-dir="-1"]');
  const right = arrows.querySelector('[data-dir="1"]');
  if (left) left.disabled = s.scrollLeft <= 0;
  if (right) right.disabled = s.scrollLeft + s.clientWidth >= s.scrollWidth - 1;
}

// Window / bar resizes change whether the tabs overflow.
if (typeof ResizeObserver === 'function') new ResizeObserver(updateTabScrollUI).observe(el.sheetTabs);
else window.addEventListener('resize', updateTabScrollUI);

// The tab bar strip itself is now click-to-switch only — renaming,
// deleting and reordering all live in the sheet-manager popup (see above).
function renderSheetTabs() {
  const prevScroll = (el.sheetTabs.querySelector('.sheet-tabs-scroll') || {}).scrollLeft || 0;
  el.sheetTabs.innerHTML = '';
  updateDeleteBtn();
  if (!state.sheetNames.length) {
    el.sheetTabs.hidden = true;
    closeSheetMenu();
    return;
  }
  el.sheetTabs.hidden = false;
  el.sheetTabs.appendChild(buildTabTools());
  // Tabs live in their own scroller so the tools (left) and scroll arrows
  // (right) stay put, and so its scrollbar can be hidden Google-Sheets-style.
  const scroller = document.createElement('div');
  scroller.className = 'sheet-tabs-scroll';
  el.sheetTabs.appendChild(scroller);
  state.sheetNames.forEach((name) => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (name === state.activeSheet ? ' active' : '');
    tab.tabIndex = 0;
    tab.dataset.sheet = name;

    const label = document.createElement('span');
    label.className = 'sheet-tab-label';
    label.textContent = name;
    tab.appendChild(label);

    // Mid-formula, this needs to happen on mousedown, before click: the tab
    // itself is focusable (tabIndex=0 below), so a plain click already
    // shifts focus to it as part of its *default* mousedown handling —
    // which blurs the formula bar / in-cell edit first, committing the
    // formula while it's still incomplete (hence "instant error"). Calling
    // preventDefault() here stops that default focus shift from happening
    // at all, so the edit surface stays focused straight through.
    tab.addEventListener('mousedown', (e) => {
      if (!isFormulaPointModeActive()) return;
      e.preventDefault();
      if (state.activeSheet !== name) switchSheetForFormulaBrowse(name);
    });
    tab.addEventListener('click', () => {
      if (state.activeSheet === name) return;
      // Already handled above, before focus could shift and blur/commit
      // the in-progress formula.
      if (isFormulaPointModeActive()) return;

      state.activeSheet = name;
      state.rangeSel = null;
      state.headerAnchor = null;
      renderSheetTabs();
      renderSheet();
    });

    scroller.appendChild(tab);
  });

  el.sheetTabs.appendChild(buildTabScrollArrows(scroller));

  scroller.scrollLeft = prevScroll;
  revealActiveTab(scroller);
  scroller.addEventListener('scroll', updateTabScrollUI, { passive: true });
  // No visible scrollbar: turn the mouse wheel into horizontal scrolling.
  scroller.addEventListener('wheel', (e) => {
    if (scroller.scrollWidth <= scroller.clientWidth) return;
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return; // trackpad sideways swipe already scrolls natively
    e.preventDefault();
    scroller.scrollLeft += e.deltaY;
  }, { passive: false });
  updateTabScrollUI();
  refreshSheetMenu();
}

// Switches the displayed sheet while a formula is being composed, so its
// cells can be clicked/dragged to insert a cross-sheet reference. If the
// formula was being edited directly in a grid cell, that edit is first
// promoted to the formula bar: the grid is about to be torn down and
// rebuilt for the other sheet, and the bar (a plain <input>, entirely
// outside the grid) is the one edit surface that safely survives that — an
// in-cell edit would just get detached/blurred by it, and lost/corrupted.
function switchSheetForFormulaBrowse(name) {
  if (state.editingTd) {
    const td = state.editingTd;
    const text = td.textContent;
    const caret = getCaretOffsetsInTd(td).end;
    td.contentEditable = 'false';
    td.classList.remove('editing');
    state.editingTd = null;
    el.formulaBarInput.value = text;
    el.formulaBarInput.focus();
    el.formulaBarInput.setSelectionRange(caret, caret);
  }

  // formulaHomeSheet/formulaHomeCell (set when the edit began) still point
  // back to the cell actually being written into. Clear the normal
  // selection rather than carrying it over — its r/c would otherwise
  // highlight some unrelated cell on this sheet and enable the row/column
  // buttons for it, neither of which means anything while just browsing
  // for a reference.
  state.rangeSel = null;
  state.activeSheet = name;
  state.headerAnchor = null;
  renderSheetTabs();
  renderSheet();
  refreshFormulaRefHighlights();
  el.formulaBarInput.focus();
}


function selectAllText(node) {
  const range = document.createRange();
  range.selectNodeContents(node);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function renameSheet(oldName, newName) {
  if (newName === oldName) { renderSheetTabs(); return; }
  if (!newName) { showToast('Sheet name cannot be empty.'); renderSheetTabs(); return; }
  if (newName.length > 31) { showToast('Sheet names must be 31 characters or fewer.'); renderSheetTabs(); return; }
  if (/[\\/?*[\]:]/.test(newName)) { showToast("Sheet names can't contain \\ / ? * [ ] :"); renderSheetTabs(); return; }
  if (state.sheetNames.some((n) => n !== oldName && n.toLowerCase() === newName.toLowerCase())) {
    showToast('A sheet with that name already exists.');
    renderSheetTabs();
    return;
  }

  pushUndo();
  const wb = state.workbook;
  const nameIdx = wb.SheetNames.indexOf(oldName);
  wb.Sheets[newName] = wb.Sheets[oldName];
  delete wb.Sheets[oldName];
  wb.SheetNames[nameIdx] = newName;
  state.sheetNames[state.sheetNames.indexOf(oldName)] = newName;
  if (state.activeSheet === oldName) state.activeSheet = newName;
  state.dirty = true;

  // Keep the live engine (and any cross-sheet formula text referencing
  // this sheet by its old name) in sync with the rename.
  let renamedInEngine = false;
  if (state.hf) {
    try {
      const sheetId = state.hf.getSheetId(oldName);
      if (sheetId !== undefined) {
        state.hf.renameSheet(sheetId, newName);
        syncFormulasFromEngine();
        renamedInEngine = true;
      }
    } catch (err) {
      console.error('Formula engine rename failed; rebuilding', err);
    }
  }
  if (!renamedInEngine) recalcFormulas();

  renderSheetTabs();
  renderSheet();
  showToast('Sheet renamed');
}

// "Sheet1", "Sheet2", ... — first name not already used by any existing
// sheet (case-insensitively, matching the same rule renameSheet enforces).
function nextDefaultSheetName() {
  const existing = new Set(state.sheetNames.map((n) => n.toLowerCase()));
  let n = state.sheetNames.length + 1;
  while (existing.has(`sheet${n}`.toLowerCase())) n++;
  return `Sheet${n}`;
}

// Resets a sheet's content to a blank 10×100 grid (A1:J100) — cell content,
// column widths, merges, and row heights dropped. Shared by clearActiveSheet
// and clearWholeWorkbook. Mirrors the Clear in scratch text (clearTextDocument)
// and scratch design (design.js clearAll): no native confirm(), just an undo step.
const clearSheetBlankRef = () => XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: BLANK_SHEET_ROWS - 1, c: BLANK_SHEET_COLS - 1 } });
function sheetIsBlank(ws, blankRef) {
  return !!ws && getLastDataRow(ws) === -1 && !ws['!cols'] && !ws['!merges'] && !ws['!rows'] && ws['!ref'] === blankRef;
}
// Clearing the scratchpad also drops any "scratchpad_<dropped file>" name it picked up,
// back to the plain default (not for a sheet opened outside the scratchpad).
function dropScratchSheetPrefix() {
  if (state.scratchId === 'sheet' && state.title !== SCRATCH_DEFAULT_NAMES.sheet) {
    state.title = SCRATCH_DEFAULT_NAMES.sheet;
    state.originalFileName = SCRATCH_DEFAULT_NAMES.sheet;
    el.fileName.textContent = state.title;
  }
}

// "This sheet only": blanks the active tab, leaving every other tab untouched.
function clearActiveSheet() {
  if (!state.workbook) return;
  const blankRef = clearSheetBlankRef();
  if (sheetIsBlank(activeSheetObj(), blankRef)) { showToast('Already empty', 1600); return; }
  pushUndo();
  state.workbook.Sheets[state.activeSheet] = { '!ref': blankRef };
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingTd = null;
  state.dirty = true;
  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  dropScratchSheetPrefix();
  showToast('Cleared — Ctrl+Z brings it back', 2600);
}

// "Whole workbook": drops every tab, down to a single blank "Sheet1" —
// exactly what "create a blank sheet" starts you with.
function clearWholeWorkbook() {
  if (!state.workbook) return;
  const blankRef = clearSheetBlankRef();
  const alreadyBlank = state.sheetNames.length === 1 && state.activeSheet === 'Sheet1' && sheetIsBlank(activeSheetObj(), blankRef);
  if (alreadyBlank) { showToast('Already empty', 1600); return; }
  pushUndo();
  state.workbook.SheetNames = ['Sheet1'];
  state.workbook.Sheets = { Sheet1: { '!ref': blankRef } };
  state.sheetNames = ['Sheet1'];
  state.activeSheet = 'Sheet1';
  state.rangeSel = null;
  state.headerAnchor = null;
  state.editingTd = null;
  state.dirty = true;
  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  dropScratchSheetPrefix();
  showToast('Cleared — Ctrl+Z brings it back', 2600);
}

// Clicking Clear opens a small popover (mirrors the Google Sheets import
// popover) asking which of the two above to do, rather than guessing or
// showing a native confirm().
function showClearSheetForm() {
  el.clearSheetForm.hidden = false;
  el.clearSheetBtn.classList.add('active');
  el.clearSheetBtn.setAttribute('aria-expanded', 'true');
}
function hideClearSheetForm() {
  el.clearSheetForm.hidden = true;
  el.clearSheetBtn.classList.remove('active');
  el.clearSheetBtn.setAttribute('aria-expanded', 'false');
}
el.clearSheetBtn.addEventListener('click', () => {
  if (el.clearSheetForm.hidden) showClearSheetForm();
  else hideClearSheetForm();
});
el.clearSheetCancelBtn.addEventListener('click', hideClearSheetForm);
el.clearSheetTabBtn.addEventListener('click', () => { hideClearSheetForm(); clearActiveSheet(); });
el.clearSheetAllBtn.addEventListener('click', () => { hideClearSheetForm(); clearWholeWorkbook(); });
el.clearSheetForm.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { e.preventDefault(); hideClearSheetForm(); el.clearSheetBtn.focus(); }
});

function addSheet() {
  if (!state.workbook) return;
  pushUndo();

  const name = nextDefaultSheetName();
  const wb = state.workbook;
  const oldOrder = wb.SheetNames.slice();
  // Goes immediately to the right of the sheet you're on (not at the far end).
  const activeIdx = wb.SheetNames.indexOf(state.activeSheet);
  const insertAt = activeIdx < 0 ? wb.SheetNames.length : activeIdx + 1;
  wb.SheetNames.splice(insertAt, 0, name);
  // Empty, but sized to a comfortable starting grid (10 cols x 100 rows,
  // i.e. A1:J100) rather than shrink-wrapped to a single cell.
  wb.Sheets[name] = { '!ref': XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 99, c: 9 } }) };
  state.sheetNames.splice(insertAt, 0, name);
  remapWorkbookSheetIndices(wb, oldOrder, wb.SheetNames);
  state.activeSheet = name;
  state.rangeSel = null;
  state.headerAnchor = null;
  state.dirty = true;

  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  showToast(`Added "${name}"`);
}

function deleteSheets(names) {
  if (!state.workbook) return;
  const uniq = [...new Set(names)].filter((n) => state.sheetNames.includes(n));
  if (!uniq.length) return;
  if (uniq.length >= state.sheetNames.length) {
    showToast("Can't delete all sheets.");
    return;
  }
  pushUndo();

  const wb = state.workbook;
  const order = state.sheetNames.slice();
  const activeIdx = order.indexOf(state.activeSheet);

  uniq.forEach((name) => {
    wb.SheetNames.splice(wb.SheetNames.indexOf(name), 1);
    delete wb.Sheets[name];
    state.sheetNames.splice(state.sheetNames.indexOf(name), 1);
  });
  remapWorkbookSheetIndices(wb, order, wb.SheetNames);

  if (!state.sheetNames.includes(state.activeSheet)) {
    // Land on the nearest surviving tab to the right, else the nearest to the left.
    const right = order.slice(activeIdx + 1).find((n) => state.sheetNames.includes(n));
    const left = order.slice(0, activeIdx).reverse().find((n) => state.sheetNames.includes(n));
    state.activeSheet = right || left || state.sheetNames[0];
    state.rangeSel = null;
    state.headerAnchor = null;
  }

  state.dirty = true;
  recalcFormulas();
  renderSheetTabs();
  renderSheet();
  showToast(uniq.length > 1 ? `${uniq.length} sheets deleted` : `Deleted "${uniq[0]}"`);
}

/* ------------------------------------------------------------------ *
 * Title (display / base file name) rename
 * ------------------------------------------------------------------ */
function startTitleRename() {
  if (!state.workbook && state.mode !== 'text' && state.mode !== 'design') return;
  el.fileName.contentEditable = 'true';
  el.fileName.spellcheck = false;
  el.fileName.focus();
  selectAllText(el.fileName);
}

function commitTitleRename() {
  el.fileName.contentEditable = 'false';
  const val = el.fileName.textContent.trim();
  state.title = val || state.title;
  el.fileName.textContent = state.title;
  // In the text editor the title is the full file name, so a new extension
  // can mean a new language (unless one was picked by hand).
  if (state.mode === 'text' && state.text && !state.text.langManual) {
    setTextLanguage(detectTextLanguage(state.title, el.textInput.value), false);
  }
  scheduleScratchAutosave(); // a rename is worth remembering too
}

el.fileName.addEventListener('keydown', (e) => {
  if (!el.fileName.isContentEditable) return;
  if (e.key === 'Enter') { e.preventDefault(); el.fileName.blur(); }
  else if (e.key === 'Escape') {
    e.preventDefault();
    el.fileName.textContent = state.title;
    el.fileName.blur();
  }
});
el.fileName.addEventListener('blur', () => {
  if (el.fileName.isContentEditable) commitTitleRename();
});
el.fileName.addEventListener('dblclick', startTitleRename);
el.renameTitleBtn.addEventListener('click', startTitleRename);

const DEFAULT_COL_PX = 90;
const MIN_COL_PX = 28;

// Column sizing model
// ------------------
// Every column sizes itself to its data, unless the user dragged it to a
// width of their own:
//   - a column is exactly as wide as its widest value plus a little
//     breathing room (AUTO_EXTRA_PX: the cell's own padding, its border and
//     1px of sub-pixel slack, which leaves ~7px after the longest value
//     instead of ending flush against the line);
//   - no column is ever wider than AUTO_MAX_COL_PX: anything longer is
//     clipped with an ellipsis ("…"), same as every cell that doesn't fit;
//   - empty column -> DEFAULT_COL_PX.
// A hand-set width (drag, or the same width applied to several selected
// columns) is stored on the sheet as { wpx, __manual: true } and wins over
// all of the above; double-clicking a resize handle drops it again. Only
// these hand-set widths are ever added to the sheet's column info (and so
// to an export) — the automatic widths exist only in the view. Widths that
// were already stored in a loaded file are left exactly as they were (and
// exported as such) but aren't used to size the view.
const AUTO_MAX_COL_PX = 1000;
const AUTO_MIN_COL_PX = 36;
const AUTO_EXTRA_PX = 14;

function getManualColWidthPx(ws, absCol) {
  const info = ws['!cols'] && ws['!cols'][absCol];
  return info && info.__manual && info.wpx ? Math.round(info.wpx) : null;
}

function setColWidthPx(ws, absCol, px) {
  if (!ws['!cols']) ws['!cols'] = [];
  const prev = ws['!cols'][absCol];
  ws['!cols'][absCol] = { wpx: px, __manual: true, ...(prev && prev.hidden ? { hidden: true } : {}) };
}

// Measures every column's data and sets its <col> width. Runs on every
// render, so the widths follow the data — but never while a cell is merely
// being edited, which is what keeps a column steady under the caret.
let measureCtx = null;
function applyColumnLayout(ws, range) {
  const table = el.grid;
  const body = table.tBodies[0];
  const nCols = range.e.c - range.s.c + 1;
  const sampleTd = body && body.querySelector('td');
  if (!sampleTd) return;
  if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
  const cs = getComputedStyle(sampleTd);
  measureCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;

  const widest = new Array(nCols).fill(0);
  const seen = new Map();
  for (const tr of body.rows) {
    for (let c = 0; c < nCols; c++) {
      const text = tr.cells[c + 1].textContent;
      if (!text) continue;
      let w = seen.get(text);
      if (w === undefined) {
        if (text.includes('\n')) {
          const lines = text.split(/\r?\n/);
          // Clip mode shows the first line only, then "…"; otherwise every
          // line is shown, so the widest one decides.
          w = state.clipEnabled
            ? measureCtx.measureText(lines[0]).width + measureCtx.measureText('…').width
            : Math.max(...lines.map((l) => measureCtx.measureText(l).width));
        } else {
          w = measureCtx.measureText(text).width;
        }
        seen.set(text, w);
      }
      if (w > widest[c]) widest[c] = w;
    }
  }

  table.querySelectorAll('colgroup > col[data-c]').forEach((col) => {
    const c = Number(col.dataset.c);
    const manual = getManualColWidthPx(ws, range.s.c + c);
    let px;
    if (manual) px = manual;
    else if (!widest[c]) px = DEFAULT_COL_PX;
    else px = Math.min(AUTO_MAX_COL_PX, Math.max(AUTO_MIN_COL_PX, Math.ceil(widest[c]) + AUTO_EXTRA_PX));
    col.style.width = px + 'px';
  });
  syncGridWidth(table);
}

function renderSheet() {
  clearFormulaRefHighlights(); // drop range boxes that belonged to the previous render
  const ws = activeSheetObj();
  const range = getRange(ws);
  const nRows = range.e.r - range.s.r + 1;
  const nCols = range.e.c - range.s.c + 1;

  const table = document.createElement('table');
  table.id = 'grid';
  if (state.clipEnabled) table.classList.add('clip-rows');

  // Colgroup: one <col> per column so each can be resized independently
  // without re-rendering the whole table (table-layout: fixed).
  const colgroup = document.createElement('colgroup');
  const rowHeadCol = document.createElement('col');
  rowHeadCol.style.width = '44px';
  colgroup.appendChild(rowHeadCol);
  for (let c = 0; c < nCols; c++) {
    const absCol = range.s.c + c;
    const col = document.createElement('col');
    col.dataset.c = c;
    col.style.width = DEFAULT_COL_PX + 'px';
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);
  syncGridWidth(table);

  // Header row (column letters)
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  headRow.appendChild(corner);

  for (let c = 0; c < nCols; c++) {
    const absCol = range.s.c + c;
    const th = document.createElement('th');
    th.dataset.col = c;
    const headerCell = buildHeaderCell(colLetter(absCol), 'col', c);
    const handle = document.createElement('div');
    handle.className = 'col-resize-handle';
    handle.addEventListener('mousedown', (e) => startColResize(e, c));
    handle.addEventListener('dblclick', (e) => {
      e.preventDefault();
      e.stopPropagation();
      resetColumnsToAuto(columnsAffectedByResize(c));
    });
    headerCell.appendChild(handle);
    th.appendChild(headerCell);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  // Body rows
  const tbody = document.createElement('tbody');
  for (let r = 0; r < nRows; r++) {
    const absRow = range.s.r + r;
    const tr = document.createElement('tr');

    const rowTh = document.createElement('th');
    rowTh.dataset.row = r;
    rowTh.appendChild(buildHeaderCell(String(absRow + 1), 'row', r));
    tr.appendChild(rowTh);

    for (let c = 0; c < nCols; c++) {
      const absCol = range.s.c + c;
      const addr = XLSX.utils.encode_cell({ r: absRow, c: absCol });
      const cell = ws[addr];
      const td = document.createElement('td');
      // Cells start out selectable-only (see attachGridListeners): a single
      // click just selects, a double-click switches this on to actually edit.
      td.contentEditable = 'false';
      td.spellcheck = false;
      td.dataset.r = r;
      td.dataset.c = c;

      if (cell) {
        // Formulas aren't parsed/recalculated by this editor, so showing the
        // raw "=SUM(...)" text would be misleading. Show the cached final
        // value instead (w = formatted display value, v = raw value) and
        // rely on the .formula-cell highlight to flag that it's derived.
        td.textContent = cellDisplayText(cell);
        if (cell.f) {
          td.classList.add('formula-cell');
          if (cell.__hfUnsupported) {
            // The formula engine can't evaluate this one (an unrecognized
            // function, or syntax — like Google Sheets' __xludf.DUMMYFUNCTION
            // wrapper — it can't even parse); what's shown is whatever value
            // was already cached in the file, not a live recalculation.
            td.classList.add('engine-unsupported-cell');
            td.title = "This formula can't be recalculated by the built-in engine — showing its last cached value instead.";
          }
        }
        // A value spilled out of an array formula (FILTER, ARRAYFORMULA, ...):
        // it has no formula of its own, so give it a paler tint than the
        // formula cell itself (see .spill-cell in styles.css).
        if (cell.__spill && !cell.f) td.classList.add('spill-cell');
      }
      // A cell with line breaks: keep its text in an inner span so the Clip
      // mode can line-clamp it to one row (td.textContent is unchanged, which
      // is what editing and copying read).
      if (td.textContent.includes('\n')) {
        const text = td.textContent;
        td.textContent = '';
        const span = document.createElement('span');
        span.className = 'ml';
        span.textContent = text;
        td.appendChild(span);
        td.classList.add('multiline');
        tr.classList.add('has-multiline');
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  el.grid.replaceWith(table);
  el.grid = table;
  applyColumnLayout(ws, range);
  attachGridListeners();
  updateSelectionUI();
  updateSelectionStats(); // keep the sum/avg/min/max counter in sync with whatever's selected now
}

// `table-layout: fixed` is only honoured by browsers when the table's own
// width is not `auto` — with an auto width they silently fall back to the
// automatic algorithm, where a column grows to fit its longest (nowrap)
// text and the others get squeezed, and a column's width can even flip
// while a cell is being edited (its text wraps, so the column suddenly
// stops needing the extra room). Pinning the table to the exact sum of its
// <col> widths makes the widths below the single source of truth.
function syncGridWidth(table = el.grid) {
  let total = 0;
  table.querySelectorAll('colgroup > col').forEach((c) => { total += parseFloat(c.style.width) || 0; });
  table.style.width = total + 'px';
}

// Dragging the resize handle of a column that is part of a whole-column
// selection resizes every selected column together (Excel / Google Sheets
// behaviour); otherwise just the column whose handle was grabbed.
function columnsAffectedByResize(relC) {
  const sel = state.rangeSel;
  if (sel && sel.type === 'col' && relC >= sel.c1 && relC <= sel.c2) {
    const cols = [];
    for (let c = sel.c1; c <= sel.c2; c++) cols.push(c);
    return cols;
  }
  return [relC];
}

// Applies { rel -> px } widths as ONE undo step: the snapshot is taken
// before the sheet's column info is touched. `before` holds the widths that
// were on screen when the drag started. Returns whether anything changed.
function commitColumnWidths(widthsByRel, before) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const changed = Object.entries(widthsByRel)
    .map(([rel, px]) => ({ rel: Number(rel), abs: range.s.c + Number(rel), px: Math.round(px) }))
    .filter((x) => x.px !== Math.round(before[x.rel]));
  if (!changed.length) return false;
  pushUndo();
  changed.forEach((x) => setColWidthPx(ws, x.abs, x.px));
  state.dirty = true;
  return true;
}

function startColResize(e, relC) {
  e.preventDefault();
  e.stopPropagation();

  const entries = columnsAffectedByResize(relC)
    .map((rel) => ({ rel, colEl: el.grid.querySelector(`col[data-c="${rel}"]`) }))
    .filter((x) => x.colEl);
  const dragged = entries.find((x) => x.rel === relC);
  if (!dragged) return;

  const startX = e.clientX;
  const startWidth = parseFloat(dragged.colEl.style.width) || DEFAULT_COL_PX;
  const before = {};
  entries.forEach((x) => { before[x.rel] = parseFloat(x.colEl.style.width) || DEFAULT_COL_PX; });
  const handle = e.currentTarget;
  handle.classList.add('resizing');

  const widthNow = () => Math.max(MIN_COL_PX, Math.round(startWidth + (lastX - startX)));
  let lastX = startX;
  function onMove(ev) {
    lastX = ev.clientX;
    // Every column in the selection takes the same width as the one being
    // dragged, so they grow / shrink together.
    const w = widthNow() + 'px';
    entries.forEach((x) => { x.colEl.style.width = w; });
    syncGridWidth();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    handle.classList.remove('resizing');
    const w = widthNow();
    const widths = {};
    entries.forEach((x) => { widths[x.rel] = w; });
    commitColumnWidths(widths, before);
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

// Double-clicking a resize handle hands the column(s) back to automatic
// sizing (fit to the data), like "auto-fit" in Excel / Google Sheets.
function resetColumnsToAuto(relCols) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const targets = relCols
    .map((rel) => range.s.c + rel)
    .filter((abs) => getManualColWidthPx(ws, abs));
  if (!targets.length) return;
  pushUndo();
  targets.forEach((abs) => { delete ws['!cols'][abs]; });
  state.dirty = true;
  renderSheet();
}

function buildHeaderCell(label, kind, index) {
  const wrap = document.createElement('div');
  wrap.className = 'header-cell';

  const text = document.createElement('span');
  text.textContent = label;
  wrap.appendChild(text);

  // Click-and-drag (or shift-click) across row/column headers selects a
  // contiguous range of whole rows/columns, for bulk copy or delete (via the toolbar Delete button) — or,
  // mid-formula, inserts a whole-row/column reference ("C:C", "5:9")
  // instead, same idea as clicking cells does (see startFormulaRangePick).
  wrap.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    if (isFormulaPointModeActive()) {
      startFormulaRangePickFromHeader(kind, index);
      return;
    }
    startHeaderRangeSelect(kind, index, e);
  });

  return wrap;
}

function startHeaderRangeSelect(axis, index, e) {
  if (e.shiftKey && state.headerAnchor && state.headerAnchor.axis === axis) {
    if (axis === 'row') selectRowRange(state.headerAnchor.index, index);
    else selectColRange(state.headerAnchor.index, index);
    return;
  }

  if (axis === 'row') selectRowRange(index, index);
  else selectColRange(index, index);

  function onMove(ev) {
    const target = document.elementFromPoint(ev.clientX, ev.clientY);
    const th = target && target.closest && target.closest(axis === 'row' ? 'th[data-row]' : 'th[data-col]');
    if (!th) return;
    const idx = Number(axis === 'row' ? th.dataset.row : th.dataset.col);
    if (axis === 'row') selectRowRange(index, idx);
    else selectColRange(index, idx);
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

/* ------------------------------------------------------------------ *
 * Selection
 * ------------------------------------------------------------------ */
function attachGridListeners() {
  el.grid.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td');
    if (!td || e.button !== 0) return;

    // Clicking inside the cell that's already being edited just places the
    // caret there normally — let the browser handle it.
    if (td.isContentEditable) return;

    // Mid-formula: clicking (or dragging across) cells inserts/updates a
    // reference at the caret, Google Sheets-style, instead of navigating
    // away and losing (or corrupting) the in-progress formula.
    if (isFormulaPointModeActive()) {
      e.preventDefault();
      startFormulaRangePick(td);
      return;
    }

    // Clicking a different cell while another one is mid-edit commits that
    // edit first (via its blur handler) before we move the selection.
    if (state.editingTd && state.editingTd !== td) state.editingTd.blur();

    e.preventDefault();
    startCellRangeSelect(td);
  });

  // Double-click is what actually opens a cell for editing; a single click
  // only selects it (see mousedown above), so Delete/Backspace can clear a
  // selected cell instead of just deleting one character at the caret.
  el.grid.addEventListener('dblclick', (e) => {
    const td = e.target.closest('td');
    if (!td) return;
    // Mid-formula, a double-click is really two ref-picking single clicks
    // (already handled by mousedown above) — don't also open this cell for
    // its own editing on top of that.
    if (isFormulaPointModeActive()) return;
    enterEditMode(td);
  });

  el.grid.addEventListener('blur', (e) => {
    const td = e.target.closest('td');
    if (!td) return;
    finishEditing(td);
  }, true);

  // Mirrors live keystrokes into the formula bar while editing a cell
  // directly in the grid, so the two stay in sync either way round.
  el.grid.addEventListener('input', (e) => {
    const td = e.target.closest('td');
    if (!td || td !== state.editingTd) return;
    el.formulaBarInput.value = td.textContent;
    refreshFormulaRefHighlights();
  });

  el.grid.addEventListener('keydown', (e) => {
    const td = e.target.closest('td');
    if (!td || !td.isContentEditable) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      // Also stop this from bubbling up to the document-level listener that
      // treats Enter on a *selected* cell as "start editing": td.blur() below
      // runs finishEditing synchronously (committing + re-rendering) before
      // this same event finishes bubbling, so without stopping it here that
      // other listener would see the freshly re-rendered cell as merely
      // "selected" and immediately re-open it for editing.
      e.stopPropagation();
      // Alt+Enter / Shift+Enter inserts a line break within the cell instead
      // of committing, matching Excel/Google Sheets multi-line cell entry.
      if (e.altKey || e.shiftKey) {
        insertTextAtCaret('\n');
        return;
      }
      td.blur();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      td.textContent = state.editingOriginal;
      td.blur();
      clearFormulaRefHighlights();
    }
  });
}

/* ------------------------------------------------------------------ *
 * Formula "point mode": click (or drag) a cell/range while typing a
 * formula to insert its reference at the caret, the way Excel/Sheets do.
 * Works whether the formula is being typed directly in a grid cell or in
 * the formula bar; either way the referenced range(s) get a colored
 * outline on the grid for as long as the formula is being edited.
 * ------------------------------------------------------------------ */
const REF_COLORS = ['#1a73e8', '#188038', '#a142f4', '#e37400', '#d01884', '#12a4af', '#c5221f', '#7a7a00'];
// Matches an A1-style cell ref or range (optionally $-anchored), with an
// optional leading sheet qualifier (Sheet1!A1 or 'Sheet Name'!A1). Careful
// not to match when it's really a function name (followed by "("), or part
// of a longer identifier.
const FORMULA_REF_RE = /(?<![A-Za-z0-9_])(?:(?:'(?<sheetQ>[^']+)'|(?<sheetN>[A-Za-z0-9_]+))!)?(?<c1>\$?[A-Za-z]{1,3}\$?[0-9]{1,7})(?::(?<c2>\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?(?![A-Za-z0-9_(])/g;
// Whole-column ("C:C", "C:E") and whole-row ("5:5", "5:9") refs — same
// guard rails (and optional sheet qualifier) as above, just without a row
// number / column letter on either side of the colon.
const COL_RANGE_RE = /(?<![A-Za-z0-9_])(?:(?:'(?<sheetQ>[^']+)'|(?<sheetN>[A-Za-z0-9_]+))!)?(?<c1>\$?[A-Za-z]{1,3}):(?<c2>\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(])/g;
const ROW_RANGE_RE = /(?<![A-Za-z0-9_])(?:(?:'(?<sheetQ>[^']+)'|(?<sheetN>[A-Za-z0-9_]+))!)?(?<c1>\$?[0-9]{1,7}):(?<c2>\$?[0-9]{1,7})(?![A-Za-z0-9_(])/g;

function isFormulaPointModeActive() {
  let text = null;
  if (state.editingTd) text = state.editingTd.textContent;
  else if (document.activeElement === el.formulaBarInput) text = el.formulaBarInput.value;
  return !!text && text.trim().startsWith('=');
}

// A leading "Sheet2!" (or "'Sheet Two'!" if the name needs quoting) to
// prepend to a reference being picked on state.activeSheet, when that
// differs from the sheet the formula is actually being written into —
// i.e. we've navigated away to browse another sheet's cells for a
// cross-sheet reference. Empty string when picking on the home sheet.
function sheetRefPrefix() {
  const home = state.formulaHomeSheet || state.activeSheet;
  const current = state.activeSheet;
  if (current === home) return '';
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(current) ? current + '!' : `'${current.replace(/'/g, "''")}'!`;
}

function addrForRange(a, b) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const r1 = Math.min(a.r, b.r), r2 = Math.max(a.r, b.r);
  const c1 = Math.min(a.c, b.c), c2 = Math.max(a.c, b.c);
  const start = colLetter(range.s.c + c1) + (range.s.r + r1 + 1);
  const body = (r1 === r2 && c1 === c2) ? start : start + ':' + colLetter(range.s.c + c2) + (range.s.r + r2 + 1);
  return sheetRefPrefix() + body;
}

// Character offsets of the current selection inside a contentEditable td's
// plain text (it's a single text node in practice, since we only ever set
// it via textContent, but this walks generally in case that ever changes).
function getCaretOffsetsInTd(td) {
  const len = td.textContent.length;
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return { start: len, end: len };
  const r = sel.getRangeAt(0);
  if (!td.contains(r.startContainer) || !td.contains(r.endContainer)) return { start: len, end: len };

  const preStart = document.createRange();
  preStart.selectNodeContents(td);
  preStart.setEnd(r.startContainer, r.startOffset);
  const preEnd = document.createRange();
  preEnd.selectNodeContents(td);
  preEnd.setEnd(r.endContainer, r.endOffset);
  return { start: preStart.toString().length, end: preEnd.toString().length };
}

function placeCaretAtOffset(td, offset) {
  const textNode = td.firstChild || td.appendChild(document.createTextNode(''));
  const pos = Math.max(0, Math.min(offset, textNode.textContent.length));
  const range = document.createRange();
  range.setStart(textNode, pos);
  range.collapse(true);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function getFormulaPickSelection(target, surface) {
  if (target === 'bar') return { start: surface.selectionStart, end: surface.selectionEnd };
  return getCaretOffsetsInTd(surface);
}

function setFormulaPickText(target, surface, text, caretPos) {
  if (target === 'bar') {
    surface.value = text;
    surface.setSelectionRange(caretPos, caretPos);
  } else {
    surface.textContent = text;
    placeCaretAtOffset(surface, caretPos);
    // Keep the formula bar mirroring the cell, same as normal typing does.
    el.formulaBarInput.value = text;
  }
  surface.focus();
}

function startFormulaRangePick(startTd) {
  const target = state.editingTd ? 'cell' : 'bar';
  const surface = target === 'cell' ? state.editingTd : el.formulaBarInput;
  const originalText = target === 'cell' ? surface.textContent : surface.value;
  const sel = getFormulaPickSelection(target, surface);

  const anchor = { r: Number(startTd.dataset.r), c: Number(startTd.dataset.c) };
  const addr = addrForRange(anchor, anchor);
  const newText = originalText.slice(0, sel.start) + addr + originalText.slice(sel.end);

  state.formulaPick = { target, surface, anchor, insertStart: sel.start, insertEnd: sel.start + addr.length };
  setFormulaPickText(target, surface, newText, state.formulaPick.insertEnd);
  refreshFormulaRefHighlights();

  function onMove(ev) {
    const p = state.formulaPick;
    if (!p) return;
    const hit = document.elementFromPoint(ev.clientX, ev.clientY);
    const td = hit && hit.closest && hit.closest('#grid td');
    if (!td) return;
    const cur = { r: Number(td.dataset.r), c: Number(td.dataset.c) };
    const newAddr = addrForRange(p.anchor, cur);
    const text = p.target === 'cell' ? p.surface.textContent : p.surface.value;
    const updated = text.slice(0, p.insertStart) + newAddr + text.slice(p.insertEnd);
    p.insertEnd = p.insertStart + newAddr.length;
    setFormulaPickText(p.target, p.surface, updated, p.insertEnd);
    refreshFormulaRefHighlights();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    state.formulaPick = null;
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

// Same idea as startFormulaRangePick, but for clicking (or dragging across)
// row/column headers while typing a formula: inserts a whole-row ("5:5",
// "5:9") or whole-column ("C:C", "C:E") reference instead of a cell range.
function startFormulaRangePickFromHeader(axis, index) {
  const target = state.editingTd ? 'cell' : 'bar';
  const surface = target === 'cell' ? state.editingTd : el.formulaBarInput;
  const originalText = target === 'cell' ? surface.textContent : surface.value;
  const sel = getFormulaPickSelection(target, surface);

  const ws = activeSheetObj();
  const range = getRange(ws);

  function refFor(a, b) {
    const lo = Math.min(a, b), hi = Math.max(a, b);
    let body;
    if (axis === 'row') {
      const r1 = range.s.r + lo + 1, r2 = range.s.r + hi + 1;
      body = `${r1}:${r2}`;
    } else {
      const c1 = colLetter(range.s.c + lo), c2 = colLetter(range.s.c + hi);
      body = `${c1}:${c2}`;
    }
    return sheetRefPrefix() + body;
  }

  const addr = refFor(index, index);
  const newText = originalText.slice(0, sel.start) + addr + originalText.slice(sel.end);

  state.formulaPick = { target, surface, headerAxis: axis, headerAnchor: index, insertStart: sel.start, insertEnd: sel.start + addr.length };
  setFormulaPickText(target, surface, newText, state.formulaPick.insertEnd);
  refreshFormulaRefHighlights();

  function onMove(ev) {
    const p = state.formulaPick;
    if (!p) return;
    const hit = document.elementFromPoint(ev.clientX, ev.clientY);
    const th = hit && hit.closest && hit.closest(axis === 'row' ? 'th[data-row]' : 'th[data-col]');
    if (!th) return;
    const idx = Number(axis === 'row' ? th.dataset.row : th.dataset.col);
    const newAddr = refFor(p.headerAnchor, idx);
    const text = p.target === 'cell' ? p.surface.textContent : p.surface.value;
    const updated = text.slice(0, p.insertStart) + newAddr + text.slice(p.insertEnd);
    p.insertEnd = p.insertStart + newAddr.length;
    setFormulaPickText(p.target, p.surface, updated, p.insertEnd);
    refreshFormulaRefHighlights();
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    state.formulaPick = null;
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function stripQuotedStrings(text) {
  return text.replace(/"(?:[^"\\]|\\.)*"/g, (m) => ' '.repeat(m.length));
}

// Scans formula text for reference tokens and returns their ranges in
// coordinates relative to the currently displayed sheet. A ref with no
// "Sheet!" qualifier is assumed to belong to the formula's home sheet
// (state.formulaHomeSheet, normally the same as the displayed sheet — it
// only differs while browsing another sheet to pick a cross-sheet ref).
// Only refs that resolve to whichever sheet is actually on screen right
// now are returned, since that's all applyRefHighlights can paint.
function extractFormulaRefRanges(text) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const cleaned = stripQuotedStrings(text);
  const results = [];
  const nRows = range.e.r - range.s.r + 1;
  const nCols = range.e.c - range.s.c + 1;
  const homeSheet = state.formulaHomeSheet || state.activeSheet;
  const displayedSheet = state.activeSheet;

  function matchesDisplayedSheet(m) {
    const explicit = m.groups.sheetQ || m.groups.sheetN;
    return (explicit || homeSheet) === displayedSheet;
  }

  FORMULA_REF_RE.lastIndex = 0;
  let m;
  while ((m = FORMULA_REF_RE.exec(cleaned))) {
    if (!matchesDisplayedSheet(m)) continue;
    try {
      const a = XLSX.utils.decode_cell(m.groups.c1.replace(/\$/g, ''));
      const b = XLSX.utils.decode_cell((m.groups.c2 || m.groups.c1).replace(/\$/g, ''));
      results.push({
        r1: Math.min(a.r, b.r) - range.s.r, r2: Math.max(a.r, b.r) - range.s.r,
        c1: Math.min(a.c, b.c) - range.s.c, c2: Math.max(a.c, b.c) - range.s.c,
      });
    } catch (err) { /* not a real reference; skip */ }
  }

  COL_RANGE_RE.lastIndex = 0;
  while ((m = COL_RANGE_RE.exec(cleaned))) {
    if (!matchesDisplayedSheet(m)) continue;
    try {
      const c1 = XLSX.utils.decode_col(m.groups.c1.replace(/\$/g, ''));
      const c2 = XLSX.utils.decode_col(m.groups.c2.replace(/\$/g, ''));
      results.push({
        r1: 0, r2: nRows - 1,
        c1: Math.min(c1, c2) - range.s.c, c2: Math.max(c1, c2) - range.s.c,
      });
    } catch (err) { /* not a real column ref; skip */ }
  }

  ROW_RANGE_RE.lastIndex = 0;
  while ((m = ROW_RANGE_RE.exec(cleaned))) {
    if (!matchesDisplayedSheet(m)) continue;
    const r1 = parseInt(m.groups.c1.replace(/\$/g, ''), 10) - 1;
    const r2 = parseInt(m.groups.c2.replace(/\$/g, ''), 10) - 1;
    if (Number.isNaN(r1) || Number.isNaN(r2)) continue;
    results.push({
      r1: Math.min(r1, r2) - range.s.r, r2: Math.max(r1, r2) - range.s.r,
      c1: 0, c2: nCols - 1,
    });
  }

  return results;
}

function hexToRgba(hex, alpha) {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

function clearFormulaRefHighlights() {
  el.grid.querySelectorAll('td.ref-hl').forEach((td) => {
    td.classList.remove('ref-hl');
    td.style.backgroundColor = '';
    delete td.dataset.refHlColored;
  });
  el.gridScroll.querySelectorAll('.ref-box').forEach((box) => box.remove());
}

function applyRefHighlights(ranges) {
  clearFormulaRefHighlights();
  if (!ranges.length) return;
  const ws = activeSheetObj();
  const range = getRange(ws);
  const nRows = range.e.r - range.s.r + 1;
  const nCols = range.e.c - range.s.c + 1;

  ranges.forEach((rg, i) => {
    const color = REF_COLORS[i % REF_COLORS.length];
    const r1 = Math.max(0, rg.r1), r2 = Math.min(nRows - 1, rg.r2);
    const c1 = Math.max(0, rg.c1), c2 = Math.min(nCols - 1, rg.c2);
    if (r1 > r2 || c1 > c2) return; // ref points entirely outside the sheet as rendered

    // The coloured outline is one box laid over the whole range (see .ref-box).
    const topLeft = el.grid.querySelector(`td[data-r="${r1}"][data-c="${c1}"]`);
    const bottomRight = el.grid.querySelector(`td[data-r="${r2}"][data-c="${c2}"]`);
    if (topLeft && bottomRight) {
      const host = el.gridScroll.getBoundingClientRect();
      const a = topLeft.getBoundingClientRect();
      const b = bottomRight.getBoundingClientRect();
      const box = document.createElement('div');
      box.className = 'ref-box';
      box.style.left = `${a.left - host.left + el.gridScroll.scrollLeft}px`;
      box.style.top = `${a.top - host.top + el.gridScroll.scrollTop}px`;
      box.style.width = `${b.right - a.left}px`;
      box.style.height = `${b.bottom - a.top}px`;
      box.style.borderColor = color;
      el.gridScroll.appendChild(box);
    }

    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const td = el.grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
        if (!td) continue;
        td.classList.add('ref-hl');
        if (!td.dataset.refHlColored) {
          td.style.backgroundColor = hexToRgba(color, 0.12);
          td.dataset.refHlColored = '1';
        }
      }
    }
  });
}

// Re-derives and redraws the colored range highlights from whatever formula
// text is currently live (in the grid cell being edited, or the formula
// bar); clears them once the text is no longer a formula (or editing ends).
function refreshFormulaRefHighlights() {
  let text = null;
  if (state.editingTd) text = state.editingTd.textContent;
  else if (document.activeElement === el.formulaBarInput) text = el.formulaBarInput.value;

  if (!text || !text.trim().startsWith('=')) {
    clearFormulaRefHighlights();
    hideFormulaHint();
    return;
  }
  applyRefHighlights(extractFormulaRefRanges(text));
  updateFormulaHint();
}

/* ------------------------------------------------------------------ *
 * Formula syntax hint
 * ------------------------------------------------------------------ *
 * While a formula is being typed, a compact two-line hint at the right
 * end of the formula bar's input shows the signature of the function the
 * caret is inside — SUM(number1, [number2], …), with the argument being
 * typed in bold — and what that argument expects.
 *
 * It deliberately lives inside the formula bar rather than in a popup
 * near the cell: a popup would sit on top of the very cells the user is
 * about to click or drag to insert a reference, and it would jump around
 * as the caret moves. Here it takes no space from the grid, never moves
 * the layout (the input's box keeps its size; only its inner right
 * padding grows to make room), ignores the mouse (a click on it doesn't
 * steal focus from the edit), and stays blank whenever the caret isn't
 * in a known function.
 *
 * All content comes from HyperFormula itself (hf.getFunctionDetails), so
 * it always matches the functions this engine actually supports.
 * ------------------------------------------------------------------ */
const formulaHintInfoCache = new Map();
let formulaHintKey = null;
let formulaHintRaf = 0;

function getFormulaFunctionInfo(name) {
  const key = name.toUpperCase();
  if (formulaHintInfoCache.has(key)) return formulaHintInfoCache.get(key);
  // Engine not built yet: answer "unknown" for now, but don't remember it.
  if (!state.hf || typeof state.hf.getFunctionDetails !== 'function') return null;
  let info = null;
  try {
    const d = state.hf.getFunctionDetails(key);
    if (d && Array.isArray(d.parameters)) {
      info = {
        name: d.localizedName || key,
        short: d.shortDescription || '',
        repeat: d.repeatLastArgs || 0,
        params: d.parameters.map((p, i) => ({
          name: p.name || 'arg' + (i + 1),
          desc: p.description || '',
          optional: !!p.optional,
        })),
      };
    }
  } catch (_) { info = null; }
  formulaHintInfoCache.set(key, info);
  return info;
}

// Which function call is the caret in, and which argument of it?
// Walks the formula up to the caret, skipping "strings" and 'quoted sheet
// names', ignoring separators inside {array constants}, and tracking
// nested parentheses. Returns { name, arg } for the innermost enclosing
// function call (arg is 0-based), { name, arg: -1 } when the caret sits
// right after a bare function name that hasn't been opened with "(" yet,
// or null.
function locateFormulaCall(text, caret) {
  const s = text.slice(0, caret);
  const stack = [];
  let braces = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"' || ch === "'") {
      i++;
      while (i < s.length) {
        if (s[i] === ch) {
          if (s[i + 1] === ch) { i += 2; continue; } // doubled quote = escaped
          break;
        }
        i++;
      }
    } else if (ch === '{') {
      braces++;
    } else if (ch === '}') {
      braces = Math.max(0, braces - 1);
    } else if (ch === '(') {
      let j = i;
      while (j > 0 && /[A-Za-z0-9_.]/.test(s[j - 1])) j--;
      const name = s.slice(j, i);
      stack.push({ fn: /^[A-Za-z_]/.test(name) ? name : null, arg: 0 });
    } else if (ch === ')') {
      stack.pop();
    } else if (ch === ',' && braces === 0 && stack.length) {
      stack[stack.length - 1].arg++;
    }
  }
  for (let k = stack.length - 1; k >= 0; k--) {
    if (stack[k].fn) return { name: stack[k].fn, arg: stack[k].arg };
  }
  const m = /(?<![A-Za-z0-9_.!$'])([A-Za-z_][A-Za-z0-9_.]*)$/.exec(s);
  return m ? { name: m[1], arg: -1 } : null;
}

// "number1" -> "number2", "criteria_range1" -> "criteria_range2"; names
// without a trailing number are shown unchanged.
function nextRepeatName(name) {
  const m = /^(.*?)(\d+)$/.exec(name);
  return m ? m[1] + (parseInt(m[2], 10) + 1) : name;
}

function renderFormulaHint(info, arg) {
  const key = info ? info.name + '|' + arg : '';
  if (key === formulaHintKey) return;
  formulaHintKey = key;

  const syn = el.formulaHintSyntax;
  const desc = el.formulaHintDesc;
  syn.textContent = '';
  if (!info) {
    desc.textContent = '';
    desc.hidden = true;
    el.formulaHint.title = '';
    el.formulaHint.classList.remove('has-content');
    return;
  }

  const n = info.params.length;
  const rep = Math.min(info.repeat, n);
  // Which displayed parameter is the active one? Arguments past the fixed
  // list map onto the "[next group]" that's shown for repeating functions.
  let act = null;
  if (arg >= 0) {
    if (arg < n) act = { group: false, idx: arg };
    else if (rep > 0) act = { group: true, idx: (arg - n) % rep };
  }

  // One token per displayed parameter (plus the "[next group]" and the
  // trailing "…" for repeating functions), each remembering which of its
  // pieces is the active one so it can be drawn in bold.
  const tokens = [];
  info.params.forEach((p, i) => {
    const on = !!act && !act.group && act.idx === i;
    tokens.push({
      active: on,
      parts: p.optional
        ? [{ text: '[' }, { text: p.name, cls: on ? 'fh-active' : '' }, { text: ']' }]
        : [{ text: p.name, cls: on ? 'fh-active' : '' }],
    });
  });
  if (rep > 0) {
    const parts = [{ text: '[' }];
    info.params.slice(n - rep).forEach((p, j) => {
      if (j) parts.push({ text: ', ' });
      parts.push({ text: nextRepeatName(p.name), cls: act && act.group && act.idx === j ? 'fh-active' : '' });
    });
    parts.push({ text: ']' });
    tokens.push({ active: !!act && act.group, parts });
    tokens.push({ active: false, parts: [{ text: '…' }] });
  }

  const draw = (compact) => {
    syn.textContent = '';
    const put = (text, cls) => {
      const span = document.createElement('span');
      if (cls) span.className = cls;
      span.textContent = text;
      syn.appendChild(span);
    };
    put(info.name, 'fh-fn');
    put('(');
    const ai = tokens.findIndex((t) => t.active);
    let first = true;
    const sep = () => { if (!first) put(', '); first = false; };
    tokens.forEach((t, i) => {
      if (compact && ai >= 0 && i !== ai) {
        // Too long to fit: fold everything before / after the active
        // argument into a single "…" so the active one always stays visible.
        if (i === 0 || i === ai + 1) { sep(); put('…'); }
        return;
      }
      sep();
      t.parts.forEach((part) => put(part.text, part.cls));
    });
    put(')');
  };
  draw(false);
  if (syn.scrollWidth > syn.clientWidth + 1) draw(true);

  // Second line: what the active argument expects, else what the
  // function does.
  const activeParam = act ? info.params[act.group ? (n - rep) + act.idx : act.idx] : null;
  const line = (activeParam && activeParam.desc) || info.short;
  desc.textContent = line;
  desc.hidden = !line;
  el.formulaHint.title = info.name + (info.short ? ': ' + info.short : '') +
    (activeParam && activeParam.desc ? '\n\n' + activeParam.name + ': ' + activeParam.desc : '');
  el.formulaHint.classList.add('has-content');
}

function getLiveFormulaEdit() {
  if (state.editingTd) {
    return { text: state.editingTd.textContent, caret: getCaretOffsetsInTd(state.editingTd).start };
  }
  if (document.activeElement === el.formulaBarInput) {
    const v = el.formulaBarInput.value;
    return { text: v, caret: el.formulaBarInput.selectionStart == null ? v.length : el.formulaBarInput.selectionStart };
  }
  return null;
}

function updateFormulaHint() {
  const live = getLiveFormulaEdit();
  if (!live || !live.text.trim().startsWith('=')) { hideFormulaHint(); return; }
  el.formulaHint.hidden = false;
  el.formulaBarField.classList.add('has-hint');
  const call = locateFormulaCall(live.text, live.caret);
  const info = call ? getFormulaFunctionInfo(call.name) : null;
  renderFormulaHint(info, call ? call.arg : -1);
}

function hideFormulaHint() {
  formulaHintKey = null;
  if (el.formulaHint.hidden) return;
  el.formulaHint.hidden = true;
  el.formulaHint.classList.remove('has-content');
  el.formulaBarField.classList.remove('has-hint');
}

// The caret can move without any text changing (arrow keys, clicking
// inside the formula), so follow the selection too.
function scheduleFormulaHintUpdate() {
  if (formulaHintRaf) return;
  formulaHintRaf = requestAnimationFrame(() => {
    formulaHintRaf = 0;
    if (state.editingTd || document.activeElement === el.formulaBarInput) updateFormulaHint();
  });
}
document.addEventListener('selectionchange', scheduleFormulaHintUpdate);
el.formulaBarInput.addEventListener('keyup', scheduleFormulaHintUpdate);
el.formulaBarInput.addEventListener('mouseup', scheduleFormulaHintUpdate);
// Clicking the hint (e.g. to read its tooltip) must not pull focus out of
// the edit, which would commit the half-typed formula.
el.formulaHint.addEventListener('mousedown', (e) => e.preventDefault());
window.addEventListener('resize', () => { formulaHintKey = null; scheduleFormulaHintUpdate(); });

// Click-and-drag over cells selects a rectangular range of cells (for
// copying, or for clearing several at once with Delete). A plain click with
// no movement just selects that one cell.
function startCellRangeSelect(startTd) {
  const anchor = { r: Number(startTd.dataset.r), c: Number(startTd.dataset.c) };
  selectCell(anchor.r, anchor.c);
  let dragging = false;

  function onMove(ev) {
    const target = document.elementFromPoint(ev.clientX, ev.clientY);
    const td = target && target.closest && target.closest('#grid td');
    if (!td) return;
    const rr = Number(td.dataset.r);
    const cc = Number(td.dataset.c);
    if (!dragging && (rr !== anchor.r || cc !== anchor.c)) dragging = true;
    if (dragging) {
      setRangeSelection(
        Math.min(anchor.r, rr), Math.min(anchor.c, cc),
        Math.max(anchor.r, rr), Math.max(anchor.c, cc),
        'cell'
      );
    }
  }
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
  }
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

// Switches a cell from "selected" into "editing": makes it contentEditable,
// focuses it, and drops the caret at the end of its text. A formula cell
// normally displays its cached value (see renderSheet), but Excel/Sheets
// convention is to show the underlying formula text while actually
// editing it, so that's swapped in here for the duration of the edit.
function enterEditMode(td) {
  if (td.isContentEditable) return;
  const r = Number(td.dataset.r);
  const c = Number(td.dataset.c);
  selectCell(r, c);

  const ws = activeSheetObj();
  const range = getRange(ws);
  const absRow = range.s.r + r;
  const absCol = range.s.c + c;
  const addr = XLSX.utils.encode_cell({ r: absRow, c: absCol });
  const cell = ws[addr];
  const editText = cell && cell.f !== undefined ? '=' + cell.f : td.textContent;

  td.contentEditable = 'true';
  td.classList.add('editing');
  state.editingTd = td;
  state.editingOriginal = editText;
  state.formulaHomeSheet = state.activeSheet;
  state.formulaHomeCell = { r, c };
  td.textContent = editText;
  td.focus();
  placeCaretAtEnd(td);

  // Set directly rather than via updateFormulaBar(), which intentionally
  // no-ops while a cell edit is in progress (see updateFormulaBar).
  el.formulaBarRef.textContent = colLetter(absCol) + (absRow + 1);
  el.formulaBarInput.value = editText;
  refreshFormulaRefHighlights();
}

// Commits whatever was typed and switches a cell back to "selected" (not
// editable) so a lone click on it won't drop a caret inside again. Always
// re-renders: even when nothing changed, a formula cell was showing its
// raw formula text for the edit (see enterEditMode above) and needs to
// revert to its cached display value.
function finishEditing(td) {
  commitCellEdit(td);
  td.contentEditable = 'false';
  td.classList.remove('editing');
  if (state.editingTd === td) state.editingTd = null;
  state.formulaHomeSheet = null;
  state.formulaHomeCell = null;
  clearFormulaRefHighlights();
  hideFormulaHint();
  renderSheet();
}

function placeCaretAtEnd(node) {
  const range = document.createRange();
  range.selectNodeContents(node);
  range.collapse(false);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

// Clears the contents of every cell in the current selection (Delete /
// Backspace), leaving formatting untouched. Skips the undo push entirely
// when the selection was already empty, so hitting Delete on blank cells
// doesn't pollute the undo stack.
function clearSelectedCells() {
  if (!state.rangeSel) return false;
  const { r1, c1, r2, c2 } = state.rangeSel;
  const ws = activeSheetObj();
  const range = getRange(ws);

  const addrs = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      addrs.push(XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c }));
    }
  }
  if (!addrs.some((a) => ws[a] !== undefined)) return false;

  pushUndo();
  addrs.forEach((a) => delete ws[a]);
  state.dirty = true;
  recalcFormulas();
  renderSheet();
  const count = addrs.length;
  showToast(count > 1 ? `${count} cells cleared` : 'Cell cleared');
  return true;
}

function selectCell(r, c) {
  state.rangeSel = { r1: r, c1: c, r2: r, c2: c, type: 'cell' };
  state.headerAnchor = null;
  updateSelectionUI();
}

function setRangeSelection(r1, c1, r2, c2, type) {
  state.rangeSel = { r1, c1, r2, c2, type };
  updateSelectionUI();
}

function selectRowRange(a, b) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const nCols = range.e.c - range.s.c + 1;
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  setRangeSelection(lo, 0, hi, nCols - 1, 'row');
  state.headerAnchor = { axis: 'row', index: a };
}

function selectColRange(a, b) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const nRows = range.e.r - range.s.r + 1;
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  setRangeSelection(0, lo, nRows - 1, hi, 'col');
  state.headerAnchor = { axis: 'col', index: a };
}

function updateSelectionUI() {
  el.grid.querySelectorAll('td.range-selected').forEach((td) => td.classList.remove('range-selected'));
  el.grid.querySelectorAll('td.selected').forEach((td) => td.classList.remove('selected'));
  el.grid.querySelectorAll('.header-selected').forEach((h) => h.classList.remove('header-selected'));

  const has = !!state.rangeSel;
  updateDeleteBtn();
  updateSelectionStats();
  el.addRowBtn.disabled = !has;
  el.addColBtn.disabled = !has;

  if (!has) return;
  const { r1, c1, r2, c2 } = state.rangeSel;

  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const td = el.grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
      if (td) td.classList.add('range-selected');
    }
  }
  const anchorTd = el.grid.querySelector(`td[data-r="${r1}"][data-c="${c1}"]`);
  anchorTd?.classList.add('selected');

  for (let r = r1; r <= r2; r++) {
    el.grid.querySelector(`th[data-row="${r}"] .header-cell`)?.classList.add('header-selected');
  }
  for (let c = c1; c <= c2; c++) {
    el.grid.querySelector(`th[data-col="${c}"] .header-cell`)?.classList.add('header-selected');
  }

  updateFormulaBar();
}

/* ------------------------------------------------------------------ *
 * Formula bar
 * ------------------------------------------------------------------ *
 * Mirrors the anchor cell of the current selection: its A1-style
 * reference on the left, and its raw editable content on the right —
 * the formula text (e.g. "=SUM(A1:A3)") for a formula cell, or the
 * plain value otherwise. Editing here and pressing Enter/blurring
 * commits it exactly as if it had been typed into the cell itself.
 * ------------------------------------------------------------------ */
function getCellEditableText(ws, addr) {
  const cell = ws[addr];
  if (!cell) return '';
  if (cell.f !== undefined) return '=' + cell.f;
  const dateText = dateDisplayText(cell);
  if (dateText !== null) return dateText;
  if (cell.v !== undefined) return String(cell.v);
  return '';
}

function updateFormulaBar() {
  // Don't fight an in-progress edit: the grid's own 'input' listener (see
  // attachGridListeners) keeps the bar in sync while typing in a cell, and
  // while the bar itself is focused it's the source of truth — including
  // while browsing another sheet to pick a cross-sheet reference, when
  // state.rangeSel/state.activeSheet no longer point at the formula's own
  // cell at all.
  if (state.editingTd || document.activeElement === el.formulaBarInput) return;

  if (!state.workbook || !state.rangeSel) {
    el.formulaBarRef.textContent = '';
    el.formulaBarInput.value = '';
    return;
  }

  const ws = activeSheetObj();
  const range = getRange(ws);
  const { r1, c1 } = state.rangeSel;
  const absRow = range.s.r + r1;
  const absCol = range.s.c + c1;
  const addr = XLSX.utils.encode_cell({ r: absRow, c: absCol });

  el.formulaBarRef.textContent = colLetter(absCol) + (absRow + 1);
  el.formulaBarInput.value = getCellEditableText(ws, addr);
}


/* ------------------------------------------------------------------ *
 * Selection summary (right end of the formula bar)
 * ------------------------------------------------------------------ *
 * Like Google Sheets: select a cell or range and a small chip shows one
 * statistic about it (Sum by default). Click it for the full list — Sum,
 * Average, Min, Max, Count (non-empty cells) and Count numbers — and pick
 * which one the chip shows.
 * ------------------------------------------------------------------ */
const SELECTION_STATS = [
  { key: 'sum', label: 'Sum', needsNumbers: true },
  { key: 'avg', label: 'Average', needsNumbers: true },
  { key: 'min', label: 'Min', needsNumbers: true },
  { key: 'max', label: 'Max', needsNumbers: true },
  { key: 'count', label: 'Count', needsNumbers: false },
  { key: 'nums', label: 'Count numbers', needsNumbers: false },
];
let selectionStatKey = 'sum';
try {
  const saved = localStorage.getItem('sheetEditor.selectionStat');
  if (SELECTION_STATS.some((s) => s.key === saved)) selectionStatKey = saved;
} catch (err) { /* storage unavailable: keep the default */ }

let selectionStatsMenu = null;
let lastSelectionStats = null;

function computeSelectionStats() {
  if (!state.workbook || !state.rangeSel) return null;
  const ws = state.workbook.Sheets[state.activeSheet];
  if (!ws) return null;
  const range = getRange(ws);
  const { r1, c1, r2, c2 } = state.rangeSel;
  let count = 0;
  let nums = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })];
      if (!cell) continue;
      const v = cell.v;
      if (v === undefined || v === null || v === '') continue; // blank / blanked-out cell
      count++;
      if (typeof v === 'number' && Number.isFinite(v)) {
        nums++;
        sum += v;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
  }
  return {
    cells: (r2 - r1 + 1) * (c2 - c1 + 1),
    count,
    nums,
    sum: nums ? sum : null,
    avg: nums ? sum / nums : null,
    min: nums ? min : null,
    max: nums ? max : null,
  };
}

function formatStatValue(n) {
  if (n === null || n === undefined) return '\u2014';
  const clean = parseFloat(Number(n).toPrecision(12)); // trim float noise (0.1 + 0.2)
  return clean.toLocaleString(undefined, { maximumFractionDigits: 10 });
}

function selectionStatDisplay(stats, key) {
  return formatStatValue(stats[key]);
}

function closeSelectionStatsMenu() {
  if (!selectionStatsMenu) return;
  selectionStatsMenu.remove();
  selectionStatsMenu = null;
  el.selectionStatsBtn.setAttribute('aria-expanded', 'false');
  document.removeEventListener('mousedown', onSelectionStatsOutside, true);
  document.removeEventListener('keydown', onSelectionStatsKey, true);
  window.removeEventListener('resize', closeSelectionStatsMenu);
}

function onSelectionStatsOutside(e) {
  if (selectionStatsMenu && !selectionStatsMenu.contains(e.target) && !el.selectionStatsBtn.contains(e.target)) closeSelectionStatsMenu();
}

function onSelectionStatsKey(e) {
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeSelectionStatsMenu(); }
}

function renderSelectionStatsMenu() {
  const menu = selectionStatsMenu;
  const stats = lastSelectionStats;
  if (!menu || !stats) return;
  menu.textContent = '';
  SELECTION_STATS.forEach((def) => {
    const row = document.createElement('button');
    row.type = 'button';
    row.className = 'selection-stats-row' + (def.key === selectionStatKey ? ' active' : '');
    row.setAttribute('role', 'menuitemradio');
    row.setAttribute('aria-checked', def.key === selectionStatKey ? 'true' : 'false');
    const label = document.createElement('span');
    label.className = 'selection-stats-row-label';
    label.textContent = def.label;
    const value = document.createElement('span');
    value.className = 'selection-stats-row-value';
    value.textContent = selectionStatDisplay(stats, def.key);
    row.appendChild(label);
    row.appendChild(value);
    row.addEventListener('mousedown', (e) => e.preventDefault()); // don't steal focus from a cell / the bar
    row.addEventListener('click', () => {
      selectionStatKey = def.key;
      try { localStorage.setItem('sheetEditor.selectionStat', def.key); } catch (err) { /* ignore */ }
      closeSelectionStatsMenu();
      updateSelectionStats();
    });
    menu.appendChild(row);
  });
}

function openSelectionStatsMenu() {
  if (selectionStatsMenu || !lastSelectionStats) return;
  const menu = document.createElement('div');
  menu.className = 'selection-stats-menu';
  menu.setAttribute('role', 'menu');
  menu.setAttribute('aria-label', 'Selection summary');
  selectionStatsMenu = menu;
  renderSelectionStatsMenu();
  document.body.appendChild(menu);

  // Drop down from the chip, right edges aligned.
  const r = el.selectionStatsBtn.getBoundingClientRect();
  const left = Math.max(8, Math.min(r.right - menu.offsetWidth, window.innerWidth - menu.offsetWidth - 8));
  menu.style.left = left + 'px';
  menu.style.top = (r.bottom + 6) + 'px';

  el.selectionStatsBtn.setAttribute('aria-expanded', 'true');
  document.addEventListener('mousedown', onSelectionStatsOutside, true);
  document.addEventListener('keydown', onSelectionStatsKey, true);
  window.addEventListener('resize', closeSelectionStatsMenu);
}

function updateSelectionStats() {
  const stats = computeSelectionStats();
  lastSelectionStats = stats;
  // Show it when there's something worth summarising: any numbers, or a
  // multi-cell selection with any content (so Count is useful). A single
  // text cell, or an empty selection, shows nothing — as in Google Sheets.
  const show = !!stats && (stats.nums > 0 || (stats.cells > 1 && stats.count > 0));
  el.selectionStats.hidden = !show;
  if (!show) { closeSelectionStatsMenu(); return; }

  // Fall back to Count when the chosen statistic needs numbers there aren't any of.
  let def = SELECTION_STATS.find((s) => s.key === selectionStatKey) || SELECTION_STATS[0];
  if (def.needsNumbers && stats.nums === 0) def = SELECTION_STATS.find((s) => s.key === 'count');
  el.selectionStatsLabel.textContent = def.label;
  el.selectionStatsValue.textContent = selectionStatDisplay(stats, def.key);
  el.selectionStatsBtn.title = SELECTION_STATS
    .map((s) => `${s.label}: ${selectionStatDisplay(stats, s.key)}`)
    .join('\n');
  if (selectionStatsMenu) renderSelectionStatsMenu();
}

el.selectionStatsBtn.addEventListener('mousedown', (e) => e.preventDefault()); // keep focus where it is
el.selectionStatsBtn.addEventListener('click', () => {
  if (selectionStatsMenu) closeSelectionStatsMenu();
  else openSelectionStatsMenu();
});

// The sheet/cell/original-text a formula-bar edit is actually writing
// into. Normally that's just the selected cell on the displayed sheet, but
// while browsing another sheet to pick a cross-sheet reference (see
// switchSheetForFormulaBrowse), state.activeSheet/state.rangeSel have
// moved on to wherever we're browsing — formulaHomeSheet/formulaHomeCell
// (set when the edit began) are what still points back home.
function getFormulaHomeTarget() {
  const homeSheet = state.formulaHomeSheet || state.activeSheet;
  const homeCell = state.formulaHomeCell || (state.rangeSel ? { r: state.rangeSel.r1, c: state.rangeSel.c1 } : null);
  if (!state.workbook || !homeCell) return null;
  const ws = state.workbook.Sheets[homeSheet];
  if (!ws) return null;
  const range = getRange(ws);
  const addr = XLSX.utils.encode_cell({ r: range.s.r + homeCell.r, c: range.s.c + homeCell.c });
  return { homeSheet, homeCell, addr, original: getCellEditableText(ws, addr) };
}

// Applies the formula bar's current text to whichever cell it's actually
// editing (see getFormulaHomeTarget) — bringing the view back to that
// cell's own sheet first if the edit had wandered off to browse another
// sheet for a reference, the same way typing directly into that cell and
// blurring would.
function commitFormulaBarValue() {
  clearFormulaRefHighlights();
  hideFormulaHint();
  const target = getFormulaHomeTarget();
  state.formulaHomeSheet = null;
  state.formulaHomeCell = null;
  if (!target) return;

  // Capture this now, before anything below (switching back to the home
  // sheet, repainting the selection) can call updateSelectionUI() ->
  // updateFormulaBar(), which would otherwise overwrite the bar's value
  // with the cell's *old* content — since by the time this runs (from the
  // 'blur' listener) the bar has already lost focus, updateFormulaBar's
  // "don't fight an in-progress edit" guard no longer applies, and it was
  // stomping the very value we're about to commit.
  const raw = el.formulaBarInput.value;

  if (state.activeSheet !== target.homeSheet) {
    state.activeSheet = target.homeSheet;
    state.headerAnchor = null;
    renderSheetTabs();
    renderSheet();
  }
  state.rangeSel = { r1: target.homeCell.r, c1: target.homeCell.c, r2: target.homeCell.r, c2: target.homeCell.c, type: 'cell' };
  updateSelectionUI();

  if (raw === target.original) return;

  const td = el.grid.querySelector(`td[data-r="${target.homeCell.r}"][data-c="${target.homeCell.c}"]`);
  if (!td) return;
  state.editingOriginal = target.original;
  td.textContent = raw;
  commitCellEdit(td);
  renderSheet();
}

el.formulaBarInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    // Also stop this from bubbling to the document-level listener that
    // treats Enter on a *selected* cell as "start editing": blur() below
    // runs commitFormulaBarValue() synchronously, which re-renders the
    // grid and leaves state.rangeSel pointed at that same cell — so once
    // focus lands on document.body, that other listener would otherwise
    // see "Enter + a selected cell" and immediately reopen it for editing,
    // right back where we just committed from.
    e.stopPropagation();
    // Blur (not a direct commitFormulaBarValue() call) so this matches the
    // grid cell's own Enter handling: the 'blur' listener below does the
    // actual commit, and blurring is what actually releases the cursor,
    // which is the part that was missing — Enter used to commit the value
    // but leave focus (and the blinking caret) stuck in the input.
    el.formulaBarInput.blur();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    // Restore the original text directly (rather than via updateFormulaBar,
    // which now deliberately no-ops while the bar is focused) so the
    // upcoming blur's commit sees no change and just backs out cleanly —
    // including snapping the view back home if we'd wandered off to browse
    // another sheet.
    const target = getFormulaHomeTarget();
    if (target) el.formulaBarInput.value = target.original;
    clearFormulaRefHighlights();
    el.formulaBarInput.blur();
  }
});
el.formulaBarInput.addEventListener('focus', () => {
  if (state.rangeSel) {
    state.formulaHomeSheet = state.activeSheet;
    state.formulaHomeCell = { r: state.rangeSel.r1, c: state.rangeSel.c1 };
  }
  refreshFormulaRefHighlights();
});
el.formulaBarInput.addEventListener('input', refreshFormulaRefHighlights);
el.formulaBarInput.addEventListener('blur', commitFormulaBarValue);


/* ------------------------------------------------------------------ *
 * Copy (rectangular selection -> tab/newline separated clipboard text)
 * ------------------------------------------------------------------ */
// Wrap a cell's text in quotes (CSV-style, doubling any internal quotes)
// whenever it contains a tab, newline, or quote itself — otherwise a
// multi-line cell's embedded newline gets misread by Excel/Sheets as a
// row break on paste, splitting one cell into several.
function escapeForClipboardCell(text) {
  if (/[\t\n\r"]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

// Custom clipboard flavor that rides alongside the plain-text copy. The
// plain text is what other apps receive, and it holds each cell's displayed
// *value* (a formula cell shows its computed result), the same as Excel/
// Sheets. Pasting back into this editor, though, should keep formulas as
// formulas, so this flavor carries them — plus where the block was copied
// from, which pasting needs in order to shift relative references (a
// formula copied down one row must point one row lower).
const CELLS_MIME = 'application/x-sheet-editor-cells';

// Cheap fingerprint of the plain-text half of a copy, stored in the custom
// flavor so paste can confirm both halves came from the same copy.
function hashText(text) {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16) + ':' + text.length;
}

const MAX_ROW_INDEX = 1048575; // Excel's last row (0-based)
const MAX_COL_INDEX = 16383;   // Excel's last column, XFD (0-based)

// Each helper shifts one reference token by (dRow, dCol), leaving $-anchored
// parts alone. They return the new token, null if the shifted reference
// would fall off the sheet (Excel turns those into #REF!), or undefined if
// the token wasn't a valid reference to begin with (e.g. "ZZZ1" is really
// a name, not a cell) so the caller leaves the text untouched.
function shiftCellToken(tok, dRow, dCol) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/.exec(tok);
  if (!m) return undefined;
  let col = XLSX.utils.decode_col(m[2].toUpperCase());
  let row = parseInt(m[4], 10) - 1;
  if (col > MAX_COL_INDEX || row < 0 || row > MAX_ROW_INDEX) return undefined;
  if (!m[1]) col += dCol;
  if (!m[3]) row += dRow;
  if (col < 0 || col > MAX_COL_INDEX || row < 0 || row > MAX_ROW_INDEX) return null;
  return m[1] + XLSX.utils.encode_col(col) + m[3] + (row + 1);
}

function shiftColToken(tok, dCol) {
  const m = /^(\$?)([A-Za-z]{1,3})$/.exec(tok);
  if (!m) return undefined;
  let col = XLSX.utils.decode_col(m[2].toUpperCase());
  if (col > MAX_COL_INDEX) return undefined;
  if (!m[1]) col += dCol;
  if (col < 0 || col > MAX_COL_INDEX) return null;
  return m[1] + XLSX.utils.encode_col(col);
}

function shiftRowToken(tok, dRow) {
  const m = /^(\$?)([0-9]{1,7})$/.exec(tok);
  if (!m) return undefined;
  let row = parseInt(m[2], 10) - 1;
  if (row < 0 || row > MAX_ROW_INDEX) return undefined;
  if (!m[1]) row += dRow;
  if (row < 0 || row > MAX_ROW_INDEX) return null;
  return m[1] + (row + 1);
}

// Rewrites every relative reference in a formula (text without the leading
// "=") as if the formula had been moved dRow rows down and dCol columns
// right — the same adjustment Excel/Sheets make when a formula cell is
// copied and pasted somewhere else. $-anchored parts stay put, string
// literals are never touched, and a reference pushed off the top/left edge
// of the sheet becomes #REF!. Handles A1 cells, A1:B2 ranges, whole-column
// (A:B) and whole-row (1:2) ranges, and an optional Sheet!/'Sheet Name'!
// qualifier on any of them.
function shiftFormulaRefs(formula, dRow, dCol) {
  if (!dRow && !dCol) return formula;
  const cleaned = stripQuotedStrings(formula); // same length, so match indexes line up
  const edits = [];

  function collect(re, shiftPair) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(cleaned))) {
      const { sheetQ, sheetN, c1, c2 } = m.groups;
      const a = shiftPair(c1);
      const b = c2 === undefined ? '' : shiftPair(c2);
      if (a === undefined || b === undefined) continue; // not a real ref; leave as-is
      const prefix = sheetQ ? `'${sheetQ}'!` : (sheetN ? `${sheetN}!` : '');
      const body = (a === null || b === null) ? '#REF!' : (c2 === undefined ? a : a + ':' + b);
      edits.push({ start: m.index, end: m.index + m[0].length, text: prefix + body });
    }
  }
  collect(FORMULA_REF_RE, (tok) => shiftCellToken(tok, dRow, dCol));
  collect(COL_RANGE_RE, (tok) => shiftColToken(tok, dCol));
  collect(ROW_RANGE_RE, (tok) => shiftRowToken(tok, dRow));

  edits.sort((x, y) => x.start - y.start);
  let out = '';
  let pos = 0;
  for (const ed of edits) {
    if (ed.start < pos) continue; // overlaps an earlier edit
    out += formula.slice(pos, ed.start) + ed.text;
    pos = ed.end;
  }
  return out + formula.slice(pos);
}

document.addEventListener('copy', (e) => {
  if (!state.workbook || el.gridWrapper.hidden) return;

  // If the user has an actual text selection (e.g. mid-edit inside a
  // cell), let the browser's native copy of that text go through.
  const domSel = window.getSelection();
  if (domSel && domSel.toString().length > 0) return;

  if (!state.rangeSel) return;
  const { r1, c1, r2, c2 } = state.rangeSel;

  const ws = activeSheetObj();
  const range = getRange(ws);

  const lines = [];
  const cells = []; // parallel grid: null for plain cells, { f, cv? } for formula cells
  let anyFormula = false;
  for (let r = r1; r <= r2; r++) {
    const rowVals = [];
    const rowCells = [];
    for (let c = c1; c <= c2; c++) {
      // Plain text is always the displayed value, so other apps get values.
      const td = el.grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
      rowVals.push(escapeForClipboardCell(td ? td.textContent : ''));

      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })];
      if (cell && typeof cell.f === 'string' && cell.f !== '') {
        const entry = { f: cell.f };
        // A formula the built-in engine can't evaluate has nothing but its
        // cached value to show, and a pasted copy can't recompute one, so
        // that value travels with it.
        if (cell.__hfUnsupported) entry.cv = { v: cell.v, w: cell.w, t: cell.t };
        rowCells.push(entry);
        anyFormula = true;
      } else {
        rowCells.push(null);
      }
    }
    lines.push(rowVals.join('\t'));
    cells.push(rowCells);
  }
  const text = lines.join('\r\n');
  if (e.clipboardData) {
    e.clipboardData.setData('text/plain', text);
    if (anyFormula) {
      try {
        e.clipboardData.setData(CELLS_MIME, JSON.stringify({
          srcR: range.s.r + r1, // absolute position of the block's top-left cell
          srcC: range.s.c + c1,
          hash: hashText(text),
          cells,
        }));
      } catch (err) { /* plain text alone still works; formulas just paste as values */ }
    }
    e.preventDefault();
  }
  const count = (r2 - r1 + 1) * (c2 - c1 + 1);
  showToast(count > 1 ? `Copied ${count} cells` : 'Cell copied');
});

/* ------------------------------------------------------------------ *
 * Paste (tab/newline separated clipboard text -> cells)
 * ------------------------------------------------------------------ */
// Inverse of escapeForClipboardCell: splits pasted text into a 2D array of
// cell strings on tabs/newlines, honoring "..."-quoted fields (with doubled
// internal quotes) so a cell's own embedded tabs/newlines/quotes round-trip
// correctly. This also happens to read the TSV that Excel/Sheets put on the
// clipboard when copying from them.
function parseClipboardGrid(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const n = text.length;

  while (i < n) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === '"' && field === '') { inQuotes = true; i++; continue; }
    if (ch === '\t') { row.push(field); field = ''; i++; continue; }
    if (ch === '\r') { i++; if (text[i] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
    field += ch; i++;
  }
  row.push(field);
  rows.push(row);

  // A trailing newline produces one bogus empty trailing row; drop it.
  if (rows.length > 1) {
    const last = rows[rows.length - 1];
    if (last.length === 1 && last[0] === '') rows.pop();
  }
  return rows;
}

// Writes one parsed value into the worksheet at addr, the same way typing
// it into a cell and committing would (see computeCellFromRaw).
function writeCellValue(ws, addr, raw) {
  const newCell = computeCellFromRaw(raw, ws[addr] || {});
  if (newCell) ws[addr] = newCell;
  else delete ws[addr];
}

// Pastes clipboard text starting at the current selection's anchor cell.
// A single pasted value dropped onto a multi-cell selection fills the whole
// selection (Excel/Sheets convention); otherwise the pasted block is written
// at its natural size, growing the sheet's declared range if it overflows.
//
// `payload` is the optional CELLS_MIME data written by our own copy handler.
// When present, formulas copied from a cell in this editor have their
// relative references shifted by however far they were moved, instead of
// being pasted as literal text.
function pasteIntoSelection(text, payload) {
  const grid = parseClipboardGrid(text);
  if (!grid.length) return;

  const ws = activeSheetObj();
  const { r1, c1, r2, c2 } = state.rangeSel;

  const fillSingleValue = grid.length === 1 && grid[0].length === 1 && (r2 > r1 || c2 > c1);
  const nRows = fillSingleValue ? (r2 - r1 + 1) : grid.length;
  const nCols = fillSingleValue ? (c2 - c1 + 1) : Math.max(...grid.map((row) => row.length));
  const destR2 = r1 + nRows - 1;
  const destC2 = c1 + nCols - 1;

  pushUndo();

  const range = getRange(ws);
  if (destR2 > range.e.r || destC2 > range.e.c) {
    range.e.r = Math.max(range.e.r, destR2);
    range.e.c = Math.max(range.e.c, destC2);
    ws['!ref'] = XLSX.utils.encode_range(range);
  }

  // Only trust the payload if its fingerprint matches the plain text that
  // arrived with it, so stale or mismatched data can never rewrite content.
  const srcCells = payload && Array.isArray(payload.cells) && payload.hash === hashText(text)
    ? payload.cells : null;

  for (let dr = 0; dr < nRows; dr++) {
    const gr = fillSingleValue ? 0 : dr % grid.length;
    const srcRow = grid[gr];
    for (let dc = 0; dc < nCols; dc++) {
      const gc = fillSingleValue ? 0 : dc;
      let raw = srcRow[gc] ?? '';
      const src = srcCells && srcCells[gr] ? srcCells[gr][gc] : null;
      if (src && typeof src.f === 'string') {
        const dRow = (range.s.r + r1 + dr) - (payload.srcR + gr);
        const dCol = (range.s.c + c1 + dc) - (payload.srcC + gc);
        raw = '=' + shiftFormulaRefs(src.f, dRow, dCol);
      }
      const addr = XLSX.utils.encode_cell({ r: range.s.r + r1 + dr, c: range.s.c + c1 + dc });
      writeCellValue(ws, addr, raw);
      if (src && src.cv && ws[addr] && ws[addr].f !== undefined) {
        // Restore the cached value of a formula the engine can't evaluate;
        // the recalc below leaves it alone and flags the cell.
        if (src.cv.v !== undefined) ws[addr].v = src.cv.v;
        if (src.cv.w !== undefined) ws[addr].w = src.cv.w;
        if (src.cv.t !== undefined) ws[addr].t = src.cv.t;
      }
    }
  }

  state.dirty = true;
  recalcFormulas();
  renderSheet();
  setRangeSelection(r1, c1, destR2, destC2, 'cell');
  const count = nRows * nCols;
  showToast(count > 1 ? `Pasted ${count} cells` : 'Cell pasted');
}

// Inserts text at the caret inside the cell currently being edited,
// replacing any selected text within it — used for Alt+Enter/Shift+Enter
// line breaks and for pasting while mid-edit. execCommand still fires the
// grid's normal 'input' listener (see attachGridListeners), which keeps the
// formula bar mirrored; the manual fallback dispatches that event itself.
function insertTextAtCaret(text) {
  if (document.execCommand && document.execCommand('insertText', false, text)) return;
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  range.deleteContents();
  const node = document.createTextNode(text);
  range.insertNode(node);
  range.setStartAfter(node);
  range.setEndAfter(node);
  sel.removeAllRanges();
  sel.addRange(range);
  document.activeElement?.dispatchEvent(new InputEvent('input', { bubbles: true }));
}

document.addEventListener('paste', (e) => {
  if (!state.workbook || el.gridWrapper.hidden) return;

  // Native inputs (formula bar, filename/tab rename, password prompt) handle
  // their own paste; only step in for the grid itself.
  const ae = document.activeElement;
  if (!state.editingTd && ae && ae !== document.body &&
      (ae.tagName === 'INPUT' || ae.tagName === 'SELECT' || ae.tagName === 'TEXTAREA')) {
    return;
  }

  const clipboardData = e.clipboardData || window.clipboardData;
  const text = clipboardData && clipboardData.getData('text/plain');
  if (!text) return;

  if (state.editingTd) {
    // Mid-edit: paste as plain text at the caret rather than as a grid paste.
    e.preventDefault();
    insertTextAtCaret(text);
    return;
  }

  if (!state.rangeSel) return;
  e.preventDefault();

  let payload = null;
  try {
    const rawPayload = clipboardData.getData(CELLS_MIME);
    if (rawPayload) payload = JSON.parse(rawPayload);
  } catch (err) { payload = null; }

  pasteIntoSelection(text, payload);
});

/* ------------------------------------------------------------------ *
 * Editing a cell's content
 * ------------------------------------------------------------------ */
// Turns a cell's raw typed/pasted text into the SheetJS cell object it
// should become (or null, meaning "delete this cell"), preserving any
// existing style/format refs. Shared by direct editing (commitCellEdit)
// and pasting (writeCellValue) so both interpret text the same way.
function computeCellFromRaw(raw, existing) {
  if (raw === '' || raw === undefined) return null;

  const newCell = {};
  if (existing.s !== undefined) newCell.s = existing.s;
  if (existing.z !== undefined) newCell.z = existing.z;

  if (raw.startsWith('=') && raw.length > 1) {
    newCell.f = raw.slice(1);
    newCell.t = 'n'; // best-effort type; recalculated by Excel/Sheets on open
  } else {
    const num = Number(raw);
    const date = raw.trim() !== '' && Number.isNaN(num) ? parseDisplayDate(raw) : null;
    if (raw.trim() !== '' && !Number.isNaN(num)) {
      newCell.t = 'n';
      newCell.v = num;
    } else if (date) {
      newCell.t = 'n';
      newCell.v = date.serial;
      if (!dateInfoOfFormat(newCell.z)) newCell.z = date.hasTime ? 'dd-mmm-yyyy hh:mm' : 'dd-mmm-yyyy';
    } else {
      newCell.t = 's';
      newCell.v = raw;
    }
  }
  return newCell;
}

function commitCellEdit(td) {
  const raw = td.textContent;
  if (raw === state.editingOriginal) return; // no actual change, skip

  const r = Number(td.dataset.r);
  const c = Number(td.dataset.c);
  const ws = activeSheetObj();
  const range = getRange(ws);
  const absRow = range.s.r + r;
  const absCol = range.s.c + c;
  const addr = XLSX.utils.encode_cell({ r: absRow, c: absCol });

  pushUndo();

  const newCell = computeCellFromRaw(raw, ws[addr] || {});
  if (newCell) {
    ws[addr] = newCell;
    td.classList.toggle('formula-cell', newCell.f !== undefined);
  } else {
    delete ws[addr];
    td.classList.remove('formula-cell');
  }
  state.dirty = true;
  state.editingOriginal = raw;
  recalcCellEdit(ws, addr);
}

/* ------------------------------------------------------------------ *
 * Structural edits: delete / insert rows & columns
 * These operate directly on the worksheet's cell objects (moving them
 * to new addresses) rather than rebuilding the sheet from scratch, so
 * that formulas and any per-cell style refs SheetJS parsed are kept
 * intact on all cells that are not removed.
 *
 * The "Core" variants mutate the worksheet only (no undo push, no
 * re-render) so a batch of several rows/columns can be removed as one
 * undo step; the plain wrappers below handle a single row/column via
 * the row/column "x" buttons.
 * ------------------------------------------------------------------ */
// Keeps merged ranges in step with a row/column being deleted (delta -1) or
// inserted (delta +1) at `index` (absolute, 0-based; for an insert, `index` is
// the position the new line takes, so everything from it onward moves down).
//  - A merge entirely before the change is untouched; one entirely after it
//    just moves.
//  - Deleting a line inside a merge shrinks the merge by one; a merge that
//    would be left as a single cell (or whose only line was deleted) is
//    dropped. Matches Excel.
//  - Inserting strictly inside a merge grows it. Inserting at its first line
//    pushes the whole merge down instead, so its top-left cell stays put.
function shiftMerges(ws, kind, index, delta) {
  if (!ws['!merges']) return;
  const sKey = kind === 'row' ? 'r' : 'c'; // axis being edited
  const oKey = kind === 'row' ? 'c' : 'r'; // the other axis
  const merges = [];
  for (const m of ws['!merges']) {
    const s = m.s[sKey];
    const e = m.e[sKey];
    if (delta < 0) {
      if (index < s) { m.s[sKey] = s - 1; m.e[sKey] = e - 1; }
      else if (index <= e) {
        if (s === e) continue; // the merge sat entirely on the deleted line
        m.e[sKey] = e - 1;
        if (m.e[sKey] === m.s[sKey] && m.e[oKey] === m.s[oKey]) continue; // shrank to one cell: no longer a merge
      }
    } else if (index <= s) {
      m.s[sKey] = s + 1; m.e[sKey] = e + 1;
    } else if (index <= e) {
      m.e[sKey] = e + 1;
    }
    merges.push(m);
  }
  if (merges.length) ws['!merges'] = merges;
  else delete ws['!merges'];
}

function deleteRowCore(relIdx) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const absIdx = range.s.r + relIdx;

  const hfOk = hfRemoveRows(state.activeSheet, absIdx, 1);

  for (let R = absIdx; R < range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const src = XLSX.utils.encode_cell({ r: R + 1, c: C });
      const dst = XLSX.utils.encode_cell({ r: R, c: C });
      if (ws[src] !== undefined) ws[dst] = ws[src];
      else delete ws[dst];
    }
  }
  for (let C = range.s.c; C <= range.e.c; C++) {
    delete ws[XLSX.utils.encode_cell({ r: range.e.r, c: C })];
  }

  range.e.r -= 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
  shiftMerges(ws, 'row', absIdx, -1);
  return hfOk;
}

function deleteColumnCore(relIdx) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const absIdx = range.s.c + relIdx;

  const hfOk = hfRemoveColumns(state.activeSheet, absIdx, 1);

  for (let C = absIdx; C < range.e.c; C++) {
    for (let R = range.s.r; R <= range.e.r; R++) {
      const src = XLSX.utils.encode_cell({ r: R, c: C + 1 });
      const dst = XLSX.utils.encode_cell({ r: R, c: C });
      if (ws[src] !== undefined) ws[dst] = ws[src];
      else delete ws[dst];
    }
  }
  for (let R = range.s.r; R <= range.e.r; R++) {
    delete ws[XLSX.utils.encode_cell({ r: R, c: range.e.c })];
  }

  range.e.c -= 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
  shiftMerges(ws, 'col', absIdx, -1);

  if (ws['!cols']) ws['!cols'].splice(absIdx, 1); // !cols is indexed by absolute column
  return hfOk;
}

function deleteRow(relIdx) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  if (range.e.r === range.s.r) {
    showToast("Can't delete the only remaining row.");
    return;
  }
  pushUndo();
  const hfOk = deleteRowCore(relIdx);
  state.rangeSel = null;
  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast('Row deleted');
}

function deleteRows(relIndices) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const totalRows = range.e.r - range.s.r + 1;
  const uniq = [...new Set(relIndices)];
  if (uniq.length >= totalRows) {
    showToast("Can't delete all rows.");
    return;
  }
  pushUndo();
  let hfOk = true;
  uniq.sort((a, b) => b - a).forEach((idx) => { if (!deleteRowCore(idx)) hfOk = false; });
  state.rangeSel = null;
  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast(uniq.length > 1 ? `${uniq.length} rows deleted` : 'Row deleted');
}

function deleteColumn(relIdx) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  if (range.e.c === range.s.c) {
    showToast("Can't delete the only remaining column.");
    return;
  }
  pushUndo();
  const hfOk = deleteColumnCore(relIdx);
  state.rangeSel = null;
  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast('Column deleted');
}

function deleteColumns(relIndices) {
  const ws = activeSheetObj();
  const range = getRange(ws);
  const totalCols = range.e.c - range.s.c + 1;
  const uniq = [...new Set(relIndices)];
  if (uniq.length >= totalCols) {
    showToast("Can't delete all columns.");
    return;
  }
  pushUndo();
  let hfOk = true;
  uniq.sort((a, b) => b - a).forEach((idx) => { if (!deleteColumnCore(idx)) hfOk = false; });
  state.rangeSel = null;
  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast(uniq.length > 1 ? `${uniq.length} columns deleted` : 'Column deleted');
}

function addRow(relIdx) {
  pushUndo();
  const ws = activeSheetObj();
  const range = getRange(ws);
  const insertAt = range.s.r + relIdx + 1; // insert after selected row
  const hfOk = hfInsertRows(state.activeSheet, insertAt, 1);

  for (let R = range.e.r; R >= insertAt; R--) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const src = XLSX.utils.encode_cell({ r: R, c: C });
      const dst = XLSX.utils.encode_cell({ r: R + 1, c: C });
      if (ws[src] !== undefined) { ws[dst] = ws[src]; delete ws[src]; }
      else delete ws[dst];
    }
  }

  range.e.r += 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
  shiftMerges(ws, 'row', insertAt, 1);

  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast('Row added');
}

function addColumn(relIdx) {
  pushUndo();
  const ws = activeSheetObj();
  const range = getRange(ws);
  const insertAt = range.s.c + relIdx + 1;
  const hfOk = hfInsertColumns(state.activeSheet, insertAt, 1);

  for (let C = range.e.c; C >= insertAt; C--) {
    for (let R = range.s.r; R <= range.e.r; R++) {
      const src = XLSX.utils.encode_cell({ r: R, c: C });
      const dst = XLSX.utils.encode_cell({ r: R, c: C + 1 });
      if (ws[src] !== undefined) { ws[dst] = ws[src]; delete ws[src]; }
      else delete ws[dst];
    }
  }

  range.e.c += 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
  shiftMerges(ws, 'col', insertAt, 1);

  // Column widths live in !cols, indexed by absolute column: open a gap so
  // every width to the right moves along with its column. The new column
  // itself gets no entry (a real hole), i.e. automatic sizing.
  if (ws['!cols'] && ws['!cols'].length > insertAt) {
    ws['!cols'].splice(insertAt, 0, undefined);
    delete ws['!cols'][insertAt];
  }

  state.dirty = true;
  finishStructuralEdit(hfOk);
  renderSheet();
  showToast('Column added');
}

/* ------------------------------------------------------------------ *
 * Saving
 * ------------------------------------------------------------------ */
// How many formula cells (across all sheets) use Google Sheets' ARRAYFORMULA().
// It is not an Excel function, so Excel shows #NAME? for these cells.
function countArrayFormulaCells() {
  if (!state.workbook) return 0;
  const re = /(^|[^A-Za-z0-9_.])ARRAYFORMULA\s*\(/i;
  let n = 0;
  for (const name of state.sheetNames) {
    const ws = state.workbook.Sheets[name];
    if (!ws) continue;
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const cell = ws[addr];
      if (cell && typeof cell.f === 'string' && re.test(cell.f)) n++;
    }
  }
  return n;
}

function saveFile() {
  if (state.mode === 'design') { if (window.Design) window.Design.exportImage(); return; }
  if (state.mode === 'text') { saveTextFile(); return; }
  if (!state.workbook) return;

  const base = state.title || 'workbook';
  const bookType = getExportFormat();
  const outName = `${base}_edited.${bookType}`;

  try {
    if (bookType === 'csv') {
      const ws = activeSheetObj();
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      downloadBlob(blob, outName);
    } else {
      XLSX.writeFile(state.workbook, outName, { bookType: 'xlsx', cellStyles: true });
    }
    state.dirty = false;
    // CSV holds plain values only, so the warning matters for workbook formats.
    const afCount = bookType === 'csv' ? 0 : countArrayFormulaCells();
    if (afCount > 0) {
      showToast(
        `Saved as "${outName}". \u26A0 ${afCount} ARRAYFORMULA cell${afCount === 1 ? '' : 's'} \u2014 ` +
        'a Google Sheets function that Excel doesn\'t support (it will show #NAME?).',
        9000
      );
    } else {
      showToast(`Saved as "${outName}"`);
    }
  } catch (err) {
    console.error(err);
    showToast('Save failed \u2014 see console for details.');
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/* ------------------------------------------------------------------ *
 * Wire up UI
 * ------------------------------------------------------------------ */
// "Open File": tries the native file picker first (so we get a real handle,
// letting the bookmark star show up for whatever's opened), and falls back
// to the classic <input type=file> in browsers that don't support it (or if
// it throws for an unexpected reason) \u2014 that fallback opens the file with
// no handle, exactly as before, so nothing about the non-bookmark path changes.
el.openBtn.addEventListener('click', async () => {
  if (typeof window.showOpenFilePicker !== 'function') { el.fileInput.click(); return; }
  let handle;
  try {
    [handle] = await window.showOpenFilePicker({ multiple: false });
  } catch (err) {
    if (err && err.name === 'AbortError') return; // user cancelled the picker
    console.error(err);
    el.fileInput.click();
    return;
  }
  let file;
  try { file = await handle.getFile(); } catch (err) { console.error(err); showToast('Failed to read file.'); return; }
  loadFile(file, undefined, handle);
});

const BLANK_SHEET_COLS = 10;
const BLANK_SHEET_ROWS = 100;

async function createBlankSpreadsheet() {
  if (!libsLoaded) {
    // First spreadsheet of the session: the libraries still have to load.
    beginLoading('Blank spreadsheet', true);
    await nextPaint();
  }
  try {
    await ensureLibs((f) => setLoading(6 + 80 * f, 'Loading spreadsheet engine\u2026'));
  } catch (err) {
    endLoading();
    console.error(err);
    showToast('Could not load the spreadsheet libraries (check the extension\u2019s lib folder).');
    return;
  }
  if (loading.active) {
    setLoading(92, 'Building sheet\u2026');
    await nextPaint();
  }
  const ws = {
    '!ref': XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: BLANK_SHEET_ROWS - 1, c: BLANK_SHEET_COLS - 1 },
    }),
  };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  await activateWorkbook(wb, 'Untitled.xlsx', false,
    `Created a blank spreadsheet (${BLANK_SHEET_COLS} columns \u00d7 ${BLANK_SHEET_ROWS} rows)`);
  selectCell(0, 0);
  endLoading();
}
el.createBlankBtn.addEventListener('click', createBlankSpreadsheet);

// Home screen "Create blank text": an empty UTF-8 document in the text editor.
// Rename it (the pencil next to the name) to e.g. notes.py to get that
// language's syntax colours.
async function createBlankText() {
  await enterTextMode({
    name: 'Untitled.txt',
    text: '',
    encoding: 'utf-8',
    bom: false,
    eol: '\n',
    eolMixed: false,
    toast: 'Created a blank text file',
  });
}
el.createTextBtn.addEventListener('click', createBlankText);
el.createDesignBtn.addEventListener('click', createBlankDesign);

/* ------------------------------------------------------------------ *
 * "Import from Google Sheet" (home screen)
 * Downloads the sheet's own XLSX export directly from Google and opens
 * it through the same path as a locally-opened file. Only works for
 * sheets shared as "Anyone with the link" — we deliberately don't send
 * the user's Google session cookies, so a private sheet fails instead of
 * silently depending on whoever happens to be signed in.
 * ------------------------------------------------------------------ */
const GOOGLE_SHEET_URL_RE = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)(?:[/?#].*)?$/;

function showGoogleImportForm() {
  el.googleImportForm.hidden = false;
  el.googleImportBtn.classList.add('active');
  el.googleImportBtn.setAttribute('aria-expanded', 'true');
  el.googleImportError.hidden = true;
  el.googleImportInput.value = '';
  el.googleImportInput.focus();
}

function hideGoogleImportForm() {
  el.googleImportForm.hidden = true;
  el.googleImportBtn.classList.remove('active');
  el.googleImportBtn.setAttribute('aria-expanded', 'false');
  el.googleImportError.hidden = true;
}

function setGoogleImportError(msg) {
  el.googleImportError.textContent = msg;
  el.googleImportError.hidden = false;
}

function setGoogleImportBusy(isBusy) {
  el.googleImportSubmitBtn.disabled = isBusy;
  el.googleImportCancelBtn.disabled = isBusy;
  el.googleImportInput.disabled = isBusy;
  el.googleImportSubmitBtn.querySelector('span').textContent = isBusy ? 'Importing\u2026' : 'Import';
}

async function importFromGoogleSheet(rawUrl) {
  const match = GOOGLE_SHEET_URL_RE.exec(rawUrl.trim());
  if (!match) {
    setGoogleImportError('That doesn\u2019t look like a Google Sheets link. It should look like https://docs.google.com/spreadsheets/d/\u2026');
    return;
  }
  const sheetId = match[1];
  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;

  setGoogleImportBusy(true);
  el.googleImportError.hidden = true;
  try {
    // credentials: 'omit' is deliberate — a sheet only comes through if
    // it's genuinely public, regardless of who's signed in to Chrome.
    const res = await fetch(exportUrl, { credentials: 'omit' });
    const contentType = res.headers.get('content-type') || '';
    // A sheet that isn't publicly shared comes back as an HTML sign-in /
    // permission page (sometimes with a 200 status), not a spreadsheet.
    if (!res.ok || !contentType.includes('spreadsheet')) {
      throw new Error('not-accessible');
    }
    const buf = await res.arrayBuffer();
    // Google's export response names the file after the sheet itself
    // (Content-Disposition) — use that so the title bar reads properly
    // instead of showing the raw file ID.
    let fileName = `${sheetId}.xlsx`;
    const disposition = res.headers.get('content-disposition') || '';
    const fileNameMatch = /filename\*?=(?:UTF-8''|")?([^";\n]+)"?/i.exec(disposition);
    if (fileNameMatch) {
      try { fileName = decodeURIComponent(fileNameMatch[1]); } catch (err) { fileName = fileNameMatch[1]; }
      if (!/\.xlsx$/i.test(fileName)) fileName += '.xlsx';
    }
    const file = new File([buf], fileName, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    setGoogleImportBusy(false);
    hideGoogleImportForm();
    loadFile(file); // same path as opening a local .xlsx file
  } catch (err) {
    console.error(err);
    setGoogleImportBusy(false);
    setGoogleImportError('Couldn\u2019t download that sheet. Make sure its sharing is set to "Anyone with the link".');
  }
}

el.googleImportBtn.addEventListener('click', () => {
  if (el.googleImportForm.hidden) showGoogleImportForm();
  else hideGoogleImportForm();
});
el.googleImportCancelBtn.addEventListener('click', hideGoogleImportForm);
el.googleImportInput.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { e.preventDefault(); hideGoogleImportForm(); }
});
el.googleImportForm.addEventListener('submit', (e) => {
  e.preventDefault();
  importFromGoogleSheet(el.googleImportInput.value);
});
el.fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) loadFile(file);
  el.fileInput.value = '';
});

['dragenter', 'dragover'].forEach((evt) =>
  el.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    el.dropzone.classList.add('drag-over');
  })
);
['dragleave', 'drop'].forEach((evt) =>
  el.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    // Moving from the dropzone onto one of its own children (the buttons)
    // fires dragleave too; only clear the highlight when actually leaving.
    if (evt === 'dragleave' && e.relatedTarget && el.dropzone.contains(e.relatedTarget)) return;
    el.dropzone.classList.remove('drag-over');
  })
);
el.dropzone.addEventListener('drop', (e) => {
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const item = e.dataTransfer.items && e.dataTransfer.items[0];
  tryGetDropHandle(item).then((handle) => loadFile(file, undefined, handle));
});

el.newBtn.addEventListener('click', goToHome);
// "New" starts over from a file, so on the homepage (nothing open) it is disabled, just like
// Save / Export. Every screen change flips the body's mode-home class, so follow that one class.
function syncNewBtn() { el.newBtn.disabled = document.body.classList.contains('mode-home'); }
new MutationObserver(syncNewBtn).observe(document.body, { attributes: true, attributeFilter: ['class'] });
syncNewBtn();

// Native `behavior: 'smooth'` scrolling runs at a fixed, fairly slow speed
// the page can't control. This animates gridScroll's own scrollTop/scrollLeft
// instead, so the top/bottom nav buttons can scroll at roughly double that
// speed — tune GRID_SCROLL_DURATION_MS to adjust it further.
const GRID_SCROLL_DURATION_MS = 220;

function easeOutQuad(t) { return t * (2 - t); }

function smoothScrollGridTo(top, left = null, duration = GRID_SCROLL_DURATION_MS) {
  const container = el.gridScroll;
  const startTop = container.scrollTop;
  const startLeft = container.scrollLeft;
  const targetTop = Math.max(0, top);
  const targetLeft = left === null ? startLeft : Math.max(0, left);
  const deltaTop = targetTop - startTop;
  const deltaLeft = targetLeft - startLeft;
  if (Math.abs(deltaTop) < 1 && Math.abs(deltaLeft) < 1) return;

  const startTime = performance.now();
  function step(now) {
    const t = Math.min(1, (now - startTime) / duration);
    const eased = easeOutQuad(t);
    container.scrollTop = startTop + deltaTop * eased;
    container.scrollLeft = startLeft + deltaLeft * eased;
    if (t < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

el.scrollTopBtn.addEventListener('click', () => {
  if (!state.workbook) return;
  smoothScrollGridTo(0, 0);
});
el.scrollBottomBtn.addEventListener('click', () => {
  if (!state.workbook) return;
  const ws = activeSheetObj();
  const range = getRange(ws);
  const maxRow = getLastDataRow(ws);
  if (maxRow < 0) { showToast('No data found.'); return; }
  const relRow = maxRow - range.s.r;
  const td = el.grid.querySelector(`td[data-r="${relRow}"]`);
  if (!td) return;
  const containerRect = el.gridScroll.getBoundingClientRect();
  const tdRect = td.getBoundingClientRect();
  const targetTop = el.gridScroll.scrollTop
    + (tdRect.top - containerRect.top)
    - (containerRect.height / 2)
    + (tdRect.height / 2);
  smoothScrollGridTo(targetTop);
});

el.undoBtn.addEventListener('click', undo);
el.redoBtn.addEventListener('click', redo);

/* ------------------------------------------------------------------ *
 * Unified Delete button
 *   - whole rows/columns selected via the headers  -> deletes them
 *   - plain cell(s) selected                       -> clears their contents
 * (Sheet deletion lives entirely in the sheet-manager popup now — see its
 * per-row trash button — so this button no longer touches sheets at all.)
 * ------------------------------------------------------------------ */
function updateDeleteBtn() {
  if (!el.deleteBtn) return;
  el.deleteBtn.disabled = !(state.workbook && state.rangeSel);
}

function structuralSelection() {
  const s = state.rangeSel;
  if (!s) return null;
  if (s.type === 'row') return { kind: 'row', count: s.r2 - s.r1 + 1 };
  if (s.type === 'col') return { kind: 'col', count: s.c2 - s.c1 + 1 };
  return null;
}

function deleteSelectedRowsOrCols() {
  const s = state.rangeSel;
  const info = structuralSelection();
  if (!info) return;
  const idxs = [];
  if (info.kind === 'row') {
    for (let r = s.r1; r <= s.r2; r++) idxs.push(r);
    deleteRows(idxs);
  } else {
    for (let c = s.c1; c <= s.c2; c++) idxs.push(c);
    deleteColumns(idxs);
  }
}

el.deleteBtn.addEventListener('click', () => {
  if (!state.workbook) return;
  const info = structuralSelection();
  if (info) { deleteSelectedRowsOrCols(); return; }
  if (state.rangeSel && !clearSelectedCells()) showToast('Nothing to clear');
});

el.addRowBtn.addEventListener('click', () => {
  if (state.rangeSel) addRow(state.rangeSel.r2);
});
el.addColBtn.addEventListener('click', () => {
  if (state.rangeSel) addColumn(state.rangeSel.c2);
});
el.saveBtn.addEventListener('click', saveFile);

// Clip: every row is held to a single row height. Cells with line breaks
// show just their first line, ending with "…" (see the .multiline rules in
// styles.css). Off (the default), those cells show all their lines and their
// row grows to fit. Column widths depend on the mode, so re-render.
el.clipBtn.addEventListener('click', () => {
  state.clipEnabled = !state.clipEnabled;
  el.clipBtn.classList.toggle('active', state.clipEnabled);
  renderSheet();
});

document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  if (!mod) return;
  if (state.mode === 'text' && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'g') { e.preventDefault(); focusTextJump(); return; }
  if (state.mode === 'text') return; // the text editor's own (native) undo history handles these
  if (state.mode === 'design') {
    // Typing in one of the design tool bar's fields (or editing text on the canvas) keeps its own native undo.
    const t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
  }
  const key = e.key.toLowerCase();
  if (key === 'z' && !e.shiftKey) {
    e.preventDefault();
    undo();
  } else if (key === 'y' || (key === 'z' && e.shiftKey)) {
    e.preventDefault();
    redo();
  }
});

// Keyboard actions for a cell that's merely *selected* (not being edited).
// Once a cell is actually in edit mode, document.activeElement is that
// contentEditable <td> itself, so this bails out and lets normal text
// editing (including its own Delete/Backspace) happen untouched.
document.addEventListener('keydown', (e) => {
  if (!state.workbook || el.gridWrapper.hidden || !state.rangeSel) return;
  if (sheetMenu) return; // the sheet-manager popup is open
  const ae = document.activeElement;
  if (ae && ae !== document.body && (ae.tagName === 'INPUT' || ae.tagName === 'SELECT' || ae.tagName === 'TEXTAREA' || ae.isContentEditable)) {
    return;
  }

  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    clearSelectedCells();
    return;
  }

  if (state.rangeSel.type !== 'cell') return; // Enter/F2/typing only make sense for an actual cell selection

  if (e.key === 'Enter' || e.key === 'F2') {
    const td = el.grid.querySelector(`td[data-r="${state.rangeSel.r1}"][data-c="${state.rangeSel.c1}"]`);
    if (td) { e.preventDefault(); enterEditMode(td); }
    return;
  }

  // A plain printable character typed over a selected cell starts editing
  // it and replaces its content, same as Excel/Google Sheets.
  if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
    const td = el.grid.querySelector(`td[data-r="${state.rangeSel.r1}"][data-c="${state.rangeSel.c1}"]`);
    if (!td) return;
    e.preventDefault();
    const original = td.textContent;
    selectCell(state.rangeSel.r1, state.rangeSel.c1);
    td.contentEditable = 'true';
    td.classList.add('editing');
    state.editingTd = td;
    state.editingOriginal = original;
    td.textContent = e.key;
    td.focus();
    placeCaretAtEnd(td);
  }
});

window.addEventListener('beforeunload', (e) => {
  if (state.dirty && !state.scratchId) {
    e.preventDefault();
    e.returnValue = '';
  }
});
