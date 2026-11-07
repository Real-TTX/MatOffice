/* MatOffice Sheets – the workbook in memory (pure, no DOM): sheets with sparse cells (a Map by number key), the style table, sizes, merges and
   the other parts of the content (docs/EDITORS.md "Sheets content"); loading and writing the content; undo/redo by transactions that remember
   the cells and parts they touched; the changes of the structure (rows/columns inserted or deleted, cells moved, sorted) with every formula,
   range and name adjusted. The calculation (engine/calc.js) reads the cells through the host the app builds on top of this. */
import { Axis } from './axis.js';
import { key, keyRow, keyCol, parseCell, cellName, colIndex, colName, parseArea, areaName, parseReference, referenceName, MAX_ROWS, MAX_COLS } from './engine/address.js';
import { shiftFormula, moveFormula, renameSheetIn, deleteSheetIn, moveAreaIn, rehostFormula } from './engine/transform.js';

const SHEET_LIST_PROPS = ['merges', 'freeze', 'filter', 'cf', 'validations', 'charts', 'images', 'print'];
const clone = v => (v === undefined ? undefined : structuredClone(v));
const sameName = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();

/** A style object without empty fields, as a stable key. */
export function styleKey(style) {
  const keys = Object.keys(style).filter(k => style[k] !== undefined && style[k] !== null && style[k] !== false && style[k] !== '').sort();
  return JSON.stringify(keys.map(k => [k, style[k]]));
}

function cleanStyle(style) {
  const out = {};
  for (const [k, v] of Object.entries(style)) if (v !== undefined && v !== null && v !== false && v !== '') out[k] = v;
  return out;
}

/** An area moved by inserted (count > 0) or deleted (count < 0) lines; null when it was deleted completely. */
export function shiftArea(a, axis, at, count) {
  const [k1, k2, max] = axis === 'row' ? ['r1', 'r2', MAX_ROWS] : ['c1', 'c2', MAX_COLS];
  let lo = a[k1];
  let hi = a[k2];
  if (count > 0) {
    if (lo >= at) lo += count;
    if (hi >= at && hi < max - 1) hi = Math.min(max - 1, hi + count);
    if (lo > max - 1) return null;
  } else {
    const n = -count;
    const end = at + n - 1;
    if (lo >= at && hi <= end) return null;
    lo = lo < at ? lo : lo > end ? lo - n : at;
    hi = hi < at ? hi : hi > end ? hi - n : at - 1;
  }
  return { ...a, [k1]: lo, [k2]: hi };
}

export class Sheet {
  constructor(id, name) {
    this.id = id;
    this.name = name;
    this.color = null;
    this.hidden = false;
    this.cells = new Map();
    this.cols = new Axis(64, MAX_COLS);
    this.rows = new Axis(20, MAX_ROWS);
    this.colStyles = new Map();
    this.rowStyles = new Map();
    this.merges = [];
    this.freeze = { rows: 0, cols: 0 };
    this.filter = null;
    this.cf = [];
    this.validations = [];
    this.charts = [];
    this.images = [];
    this.grid = true;
    this.headings = true;
    this.zoom = 1;
    this.print = {};
    this.protected = false;
    this.sel = 'A1';
    this.boundsCache = null;
    this.mergeIndex = null;
  }

  cell(row, col) { return this.cells.get(key(row, col)); }

  /** The used area: { rows, cols } = 1 + the last row/column that has a cell (content or format). */
  bounds() {
    if (!this.boundsCache) {
      let rows = 0;
      let cols = 0;
      for (const k of this.cells.keys()) {
        const r = keyRow(k) + 1;
        const c = keyCol(k) + 1;
        if (r > rows) rows = r;
        if (c > cols) cols = c;
      }
      this.boundsCache = { rows, cols };
    }
    return this.boundsCache;
  }

  /** The used area of the values (cells with a value or formula only). */
  dataBounds() {
    let rows = 0;
    let cols = 0;
    for (const [k, c] of this.cells) {
      if (c.v === undefined && !c.f) continue;
      rows = Math.max(rows, keyRow(k) + 1);
      cols = Math.max(cols, keyCol(k) + 1);
    }
    return { rows, cols };
  }

  invalidate() { this.boundsCache = null; }

  /** The style index of a cell: its own, else its row's, else its column's. */
  styleAt(row, col) {
    const c = this.cells.get(key(row, col));
    if (c && c.s !== undefined) return c.s;
    return this.rowStyles.get(row) ?? this.colStyles.get(col) ?? 0;
  }

  /** The merged area a cell belongs to, or null. */
  mergeAt(row, col) {
    if (!this.merges.length) return null;
    if (!this.mergeIndex) {
      this.mergeIndex = new Map();
      for (const m of this.merges) {
        if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) > 100000) continue;
        for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) this.mergeIndex.set(key(r, c), m);
      }
    }
    const hit = this.mergeIndex.get(key(row, col));
    if (hit) return hit;
    for (const m of this.merges) if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) > 100000 && row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) return m;
    return null;
  }

  mergesChanged() { this.mergeIndex = null; }

  /** The filter range as an area (or null). */
  filterArea() { return this.filter ? parseArea(this.filter.range) : null; }
}

export class Workbook {
  constructor() {
    this.sheets = [];
    this.styles = [{}];
    this.styleKeys = new Map([[styleKey({}), 0]]);
    this.names = [];
    this.active = 0;
    this.undoStack = [];
    this.redoStack = [];
    this.tx = null;
    this.changes = [];
    this.nextSheet = 1;
    this.nextObject = 1;
  }

  // ---- content in and out -----------------------------------------------------------------------------------------------------

  /** A workbook from the content JSON (checked by the server; still read defensively). */
  static load(content, defaultName = 'Sheet1') {
    const wb = new Workbook();
    const c = content && typeof content === 'object' ? content : {};
    wb.styles = Array.isArray(c.styles) && c.styles.length ? c.styles.map(s => (s && typeof s === 'object' ? cleanStyle(s) : {})) : [{}];
    wb.styles[0] = {};
    wb.styleKeys = new Map();
    wb.styles.forEach((s, i) => { const k = styleKey(s); if (!wb.styleKeys.has(k)) wb.styleKeys.set(k, i); });
    wb.names = Array.isArray(c.names) ? c.names.filter(n => n && n.name && n.ref).map(n => ({ name: String(n.name), ref: String(n.ref) })) : [];
    const sheets = Array.isArray(c.sheets) && c.sheets.length ? c.sheets : [{ id: 's1', name: defaultName, cells: {} }];
    for (const raw of sheets) wb.sheets.push(wb.readSheet(raw));
    for (const s of wb.sheets) { const n = /^s(\d+)$/.exec(s.id); if (n) wb.nextSheet = Math.max(wb.nextSheet, Number(n[1]) + 1); }
    wb.active = Math.min(Math.max(0, Number(c.active) || 0), wb.sheets.length - 1);
    if (wb.sheets[wb.active].hidden) wb.active = Math.max(0, wb.sheets.findIndex(s => !s.hidden));
    return wb;
  }

  readSheet(raw) {
    const s = new Sheet(String(raw.id || 's' + this.nextSheet++), String(raw.name || 'Sheet'));
    s.color = raw.color || null;
    s.hidden = !!raw.hidden;
    if (raw.colW) s.cols.setDefault(raw.colW);
    if (raw.rowH) s.rows.setDefault(raw.rowH);
    for (const [addr, cell] of Object.entries(raw.cells || {})) {
      const p = parseCell(addr);
      if (!p || !cell || typeof cell !== 'object') continue;
      const c = {};
      if (cell.f !== undefined && cell.f !== '') c.f = String(cell.f);
      if (cell.v !== undefined && cell.v !== null) c.v = cell.v;
      if (cell.t === 'e') c.t = 'e';
      if (Number.isInteger(cell.s) && cell.s >= 0 && cell.s < this.styles.length) c.s = cell.s;
      if (cell.link) c.link = String(cell.link);
      if (cell.note) c.note = String(cell.note);
      if (Object.keys(c).length) s.cells.set(key(p.row, p.col), c);
    }
    for (const [letters, line] of Object.entries(raw.cols || {})) {
      const i = colIndex(letters);
      if (i < 0 || !line) continue;
      if (line.w !== undefined) s.cols.setSize(i, line.w);
      if (line.hidden) s.cols.setHidden(i, true);
      if (Number.isInteger(line.s) && line.s > 0 && line.s < this.styles.length) s.colStyles.set(i, line.s);
    }
    for (const [num, line] of Object.entries(raw.rows || {})) {
      const i = Number(num) - 1;
      if (!(i >= 0 && i < MAX_ROWS) || !line) continue;
      if (line.h !== undefined) s.rows.setSize(i, line.h);
      if (line.hidden) s.rows.setHidden(i, true);
      if (Number.isInteger(line.s) && line.s > 0 && line.s < this.styles.length) s.rowStyles.set(i, line.s);
    }
    s.merges = (raw.merges || []).map(parseArea).filter(Boolean);
    if (raw.freeze) s.freeze = { rows: Math.max(0, raw.freeze.rows | 0), cols: Math.max(0, raw.freeze.cols | 0) };
    if (raw.filter && parseArea(raw.filter.range)) {
      s.filter = clone(raw.filter);
      // rows hidden inside a filter are the filter's (it is applied again when the sheet opens)
      const a = parseArea(raw.filter.range);
      for (let r = a.r1 + 1; r <= a.r2; r++) if (s.rows.hidden.has(r)) s.rows.hidden.delete(r);
      s.rows.dirty = true;
    }
    s.cf = Array.isArray(raw.cf) ? clone(raw.cf) : [];
    s.validations = Array.isArray(raw.validations) ? clone(raw.validations) : [];
    s.charts = Array.isArray(raw.charts) ? clone(raw.charts) : [];
    s.images = Array.isArray(raw.images) ? clone(raw.images) : [];
    for (const o of [...s.charts, ...s.images]) { const n = /^[ci](\d+)$/.exec(o.id || ''); if (n) this.nextObject = Math.max(this.nextObject, Number(n[1]) + 1); }
    s.grid = raw.grid !== false;
    s.headings = raw.headings !== false;
    s.zoom = Number(raw.zoom) > 0 ? Number(raw.zoom) : 1;
    s.print = raw.print && typeof raw.print === 'object' ? clone(raw.print) : {};
    s.protected = !!raw.protected;
    s.sel = raw.sel || 'A1';
    return s;
  }

  /** The content JSON (styles that are not used any more are left out). */
  toContent() {
    const used = new Set([0]);
    for (const s of this.sheets) {
      for (const c of s.cells.values()) if (c.s !== undefined) used.add(c.s);
      for (const v of s.rowStyles.values()) used.add(v);
      for (const v of s.colStyles.values()) used.add(v);
    }
    const order = [...used].filter(i => i < this.styles.length).sort((a, b) => a - b);
    const map = new Map(order.map((old, i) => [old, i]));
    const styles = order.map(i => cleanStyle(this.styles[i]));
    const content = { version: 1, active: this.active, styles };
    if (this.names.length) content.names = this.names.map(n => ({ name: n.name, ref: n.ref }));
    content.sheets = this.sheets.map(s => this.writeSheet(s, map));
    return content;
  }

  writeSheet(s, map) {
    const o = { id: s.id, name: s.name };
    if (s.color) o.color = s.color;
    if (s.hidden) o.hidden = true;
    const cells = {};
    const keys = [...s.cells.keys()].sort((a, b) => a - b);
    for (const k of keys) {
      const c = s.cells.get(k);
      const out = {};
      if (c.f) out.f = c.f;
      if (c.v !== undefined && c.v !== null && !(typeof c.v === 'number' && !Number.isFinite(c.v))) out.v = c.v;
      if (c.t === 'e' && typeof out.v === 'string') out.t = 'e';
      if (c.s !== undefined) {
        const s2 = map.get(c.s) ?? 0;
        const row = keyRow(k);
        const col = keyCol(k);
        if (s2 !== 0 || s.rowStyles.has(row) || s.colStyles.has(col)) out.s = s2;
      }
      if (c.link) out.link = c.link;
      if (c.note) out.note = c.note;
      if (Object.keys(out).length) cells[cellName(keyRow(k), keyCol(k))] = out;
    }
    o.cells = cells;
    const lines = (axis, styles, isRow) => {
      const out = {};
      const idx = new Set([...axis.sizes.keys(), ...axis.hidden, ...axis.filtered, ...styles.keys()]);
      for (const i of [...idx].sort((a, b) => a - b)) {
        const line = {};
        if (axis.sizes.has(i)) line[isRow ? 'h' : 'w'] = axis.sizes.get(i);
        if (axis.hidden.has(i) || axis.filtered.has(i)) line.hidden = true;
        if (styles.has(i)) { const v = map.get(styles.get(i)); if (v) line.s = v; }
        if (Object.keys(line).length) out[isRow ? String(i + 1) : colName(i)] = line;
      }
      return out;
    };
    const cols = lines(s.cols, s.colStyles, false);
    if (Object.keys(cols).length) o.cols = cols;
    const rows = lines(s.rows, s.rowStyles, true);
    if (Object.keys(rows).length) o.rows = rows;
    if (s.cols.def !== 64) o.colW = s.cols.def;
    if (s.rows.def !== 20) o.rowH = s.rows.def;
    if (s.merges.length) o.merges = s.merges.map(areaName);
    if (s.freeze.rows || s.freeze.cols) o.freeze = { rows: s.freeze.rows, cols: s.freeze.cols };
    if (s.filter) o.filter = clone(s.filter);
    if (s.cf.length) o.cf = clone(s.cf);
    if (s.validations.length) o.validations = clone(s.validations);
    if (s.charts.length) o.charts = clone(s.charts);
    if (s.images.length) o.images = clone(s.images);
    if (!s.grid) o.grid = false;
    if (!s.headings) o.headings = false;
    if (s.zoom !== 1) o.zoom = Math.round(s.zoom * 100) / 100;
    if (s.print && Object.keys(s.print).length) o.print = clone(s.print);
    if (s.protected) o.protected = true;
    if (s.sel && s.sel !== 'A1') o.sel = s.sel;
    return o;
  }

  // ---- lookups -----------------------------------------------------------------------------------------------------------------

  get activeSheet() { return this.sheets[this.active]; }
  sheetById(id) { return this.sheets.find(s => s.id === id) || null; }
  sheetByName(name) { return this.sheets.find(s => sameName(s.name, name)) || null; }
  style(i) { return this.styles[i] || this.styles[0]; }

  /** The index of a style (added to the table when it is new). */
  styleIndex(style) {
    const clean = cleanStyle(style);
    const k = styleKey(clean);
    let i = this.styleKeys.get(k);
    if (i === undefined) {
      i = this.styles.length;
      this.styles.push(clean);
      this.styleKeys.set(k, i);
    }
    return i;
  }

  /** The index of the style i with some fields changed (undefined/false/null removes a field). */
  patchStyle(i, patch) {
    return this.styleIndex({ ...this.style(i), ...patch });
  }

  newSheetId() {
    let id;
    do { id = 's' + this.nextSheet++; } while (this.sheets.some(s => s.id === id));
    return id;
  }

  newObjectId(prefix) {
    let id;
    const all = this.sheets.flatMap(s => [...s.charts, ...s.images]);
    do { id = prefix + this.nextObject++; } while (all.some(o => o.id === id));
    return id;
  }

  /** A sheet name that is not taken ("Sheet3", "Sales (2)"). */
  uniqueName(base) {
    const clean = String(base).replace(/[:\\/?*[\]]/g, '').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
    if (!this.sheetByName(clean)) return clean;
    for (let n = 2; ; n++) {
      const suffix = ` (${n})`;
      const name = clean.slice(0, 31 - suffix.length).trimEnd() + suffix;
      if (!this.sheetByName(name)) return name;
    }
  }

  // ---- transactions (undo / redo) ---------------------------------------------------------------------------------------------

  /** Starts a change; nested begin/commit pairs join the outer one. */
  begin(label = '') {
    if (this.tx) { this.tx.depth++; return this.tx; }
    this.tx = { label, cells: new Map(), props: new Map(), book: new Map(), structural: false, depth: 1, selBefore: null, selAfter: null };
    return this.tx;
  }

  /** Ends a change: it goes onto the undo stack (when something changed). */
  commit() {
    const tx = this.tx;
    if (!tx) return null;
    if (--tx.depth > 0) return null;
    this.tx = null;
    const entry = { label: tx.label, cells: [], props: [], book: [], structural: tx.structural, selBefore: tx.selBefore, selAfter: tx.selAfter };
    for (const rec of tx.cells.values()) {
      const now = rec.sheet.cells.get(rec.k);
      const after = now ? { ...now } : undefined;
      if (JSON.stringify(rec.before) !== JSON.stringify(after)) entry.cells.push({ sheet: rec.sheet, k: rec.k, before: rec.before, after });
    }
    for (const rec of tx.props.values()) entry.props.push({ sheet: rec.sheet, prop: rec.prop, before: rec.before, after: this.snapProp(rec.sheet, rec.prop) });
    for (const rec of tx.book.values()) entry.book.push({ prop: rec.prop, before: rec.before, after: this.snapBook(rec.prop) });
    if (!entry.cells.length && !entry.props.length && !entry.book.length) return null;
    this.undoStack.push(entry);
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack = [];
    return entry;
  }

  /** Cancels the open change (restores what it touched). */
  rollback() {
    const tx = this.tx;
    if (!tx) return;
    this.tx = null;
    this.applyEntry({ cells: [...tx.cells.values()].map(r => ({ ...r })), props: [...tx.props.values()], book: [...tx.book.values()] }, 'before');
  }

  touchCell(sheet, k) {
    const tx = this.tx;
    if (!tx) return;
    const id = sheet.id + '|' + k;
    if (!tx.cells.has(id)) { const c = sheet.cells.get(k); tx.cells.set(id, { sheet, k, before: c ? { ...c } : undefined }); }
  }

  touchProp(sheet, prop) {
    const tx = this.tx;
    if (!tx) return;
    const id = sheet.id + '|' + prop;
    if (!tx.props.has(id)) tx.props.set(id, { sheet, prop, before: this.snapProp(sheet, prop) });
  }

  touchBook(prop) {
    const tx = this.tx;
    if (!tx) return;
    if (!tx.book.has(prop)) tx.book.set(prop, { prop, before: this.snapBook(prop) });
  }

  snapProp(sheet, prop) {
    if (prop === 'cols') return { axis: sheet.cols.snapshot(), styles: [...sheet.colStyles] };
    if (prop === 'rows') return { axis: sheet.rows.snapshot(), styles: [...sheet.rowStyles] };
    if (prop === 'merges') return sheet.merges.map(m => ({ ...m }));
    return clone(sheet[prop]);
  }

  restoreProp(sheet, prop, value) {
    if (prop === 'cols') { sheet.cols.restore(value.axis); sheet.colStyles = new Map(value.styles); return; }
    if (prop === 'rows') { sheet.rows.restore(value.axis); sheet.rowStyles = new Map(value.styles); return; }
    if (prop === 'merges') { sheet.merges = value.map(m => ({ ...m })); sheet.mergesChanged(); return; }
    sheet[prop] = clone(value);
  }

  snapBook(prop) {
    if (prop === 'sheets') return [...this.sheets];
    if (prop === 'names') return clone(this.names);
    return this[prop];
  }

  restoreBook(prop, value) {
    if (prop === 'sheets') this.sheets = [...value];
    else if (prop === 'names') this.names = clone(value);
    else this[prop] = value;
  }

  get canUndo() { return this.undoStack.length > 0; }
  get canRedo() { return this.redoStack.length > 0; }

  /** Undo the last change: the entry (with its selection before) or null. */
  undo() {
    const e = this.undoStack.pop();
    if (!e) return null;
    this.applyEntry(e, 'before');
    this.redoStack.push(e);
    return e;
  }

  redo() {
    const e = this.redoStack.pop();
    if (!e) return null;
    this.applyEntry(e, 'after');
    this.undoStack.push(e);
    return e;
  }

  applyEntry(e, which) {
    for (const b of e.book || []) this.restoreBook(b.prop, b[which]);
    for (const p of e.props || []) this.restoreProp(p.sheet, p.prop, p[which]);
    for (const c of e.cells || []) {
      const v = c[which];
      if (v) c.sheet.cells.set(c.k, { ...v }); else c.sheet.cells.delete(c.k);
      c.sheet.invalidate();
      this.changes.push({ sheet: c.sheet.id, row: keyRow(c.k), col: keyCol(c.k) });
    }
    if (this.active >= this.sheets.length) this.active = this.sheets.length - 1;
  }

  /** The cells changed since the last call (for the calculation). */
  takeChanges() {
    const out = this.changes;
    this.changes = [];
    return out;
  }

  // ---- cells ------------------------------------------------------------------------------------------------------------------

  /** Replaces a cell (undefined removes it). */
  setCell(sheet, row, col, cell) {
    const k = key(row, col);
    this.touchCell(sheet, k);
    const has = cell && Object.keys(cell).some(f => cell[f] !== undefined);
    if (has) {
      const clean = {};
      for (const [f, v] of Object.entries(cell)) if (v !== undefined) clean[f] = v;
      sheet.cells.set(k, clean);
      const b = sheet.boundsCache;
      if (b && (row >= b.rows || col >= b.cols)) sheet.boundsCache = { rows: Math.max(b.rows, row + 1), cols: Math.max(b.cols, col + 1) };
    } else {
      sheet.cells.delete(k);
      sheet.invalidate();
    }
    this.changes.push({ sheet: sheet.id, row, col });
  }

  /** Changes some fields of a cell (undefined removes a field). */
  patchCell(sheet, row, col, patch) {
    const before = sheet.cell(row, col) || {};
    const next = { ...before, ...patch };
    for (const f of Object.keys(next)) if (next[f] === undefined) delete next[f];
    this.setCell(sheet, row, col, next);
  }

  /** The result of a formula written into its cell (no undo, no change event: the calculation does it). */
  setResult(sheet, row, col, value) {
    const c = sheet.cells.get(key(row, col));
    if (!c) return;
    if (value && typeof value === 'object' && value.code) { c.v = value.code; c.t = 'e'; } else { c.v = value === null || value === undefined ? 0 : value; delete c.t; }
  }

  // ---- sheet properties ----------------------------------------------------------------------------------------------------------

  /** Sets a property of a sheet with undo (name, color, hidden, freeze, grid, headings, zoom, print, protected, merges, cf …). */
  setSheetProp(sheet, prop, value) {
    this.touchProp(sheet, prop);
    sheet[prop] = value;
    if (prop === 'merges') sheet.mergesChanged();
  }

  // ---- sheets ----------------------------------------------------------------------------------------------------------------

  addSheet(name, at = this.sheets.length) {
    this.touchBook('sheets');
    const s = new Sheet(this.newSheetId(), this.uniqueName(name));
    this.sheets.splice(at, 0, s);
    return s;
  }

  /** Removes a sheet; references to it become #REF!. */
  deleteSheet(sheet) {
    if (this.sheets.filter(s => !s.hidden).length <= 1 && !sheet.hidden) return false;
    this.begin('delete sheet').structural = true;
    this.touchBook('sheets');
    this.touchBook('names');
    this.touchBook('active');
    const index = this.sheets.indexOf(sheet);
    this.sheets.splice(index, 1);
    this.eachFormula((s, row, col, c) => {
      const f = deleteSheetIn(c.f, sheet.name);
      if (f !== c.f) this.patchCell(s, row, col, { f });
    });
    this.names = this.names.map(n => ({ ...n, ref: deleteSheetIn(n.ref, sheet.name) }));
    if (this.active >= this.sheets.length || this.active === index) this.active = Math.max(0, Math.min(index, this.sheets.length - 1));
    if (this.sheets[this.active].hidden) this.active = Math.max(0, this.sheets.findIndex(s => !s.hidden));
    this.commit();
    return true;
  }

  /** Renames a sheet; every formula, name and chart that refers to it follows. Returns false for a name that is taken or not allowed. */
  renameSheet(sheet, name) {
    const clean = String(name).trim();
    if (!clean || clean.length > 31 || /[:\\/?*[\]]/.test(clean) || /^'|'$/.test(clean)) return false;
    if (this.sheets.some(s => s !== sheet && sameName(s.name, clean))) return false;
    if (clean === sheet.name) return true;
    const old = sheet.name;
    this.begin('rename sheet').structural = true;
    this.touchProp(sheet, 'name');
    this.touchBook('names');
    sheet.name = clean;
    this.eachFormula((s, row, col, c) => {
      const f = renameSheetIn(c.f, old, clean);
      if (f !== c.f) this.patchCell(s, row, col, { f });
    });
    this.names = this.names.map(n => ({ ...n, ref: renameSheetIn(n.ref, old, clean) }));
    for (const s of this.sheets) {
      if (!s.charts.some(ch => /!/.test(ch.range))) continue;
      this.touchProp(s, 'charts');
      for (const ch of s.charts) ch.range = renameSheetIn(ch.range, old, clean);
    }
    this.commit();
    return true;
  }

  moveSheet(sheet, to) {
    this.touchBook('sheets');
    this.touchBook('active');
    const activeSheet = this.activeSheet;
    const from = this.sheets.indexOf(sheet);
    this.sheets.splice(from, 1);
    this.sheets.splice(Math.max(0, Math.min(to, this.sheets.length)), 0, sheet);
    this.active = this.sheets.indexOf(activeSheet);
  }

  /** A copy of a sheet after it ("Sales (2)"); formulas keep pointing where they did. */
  duplicateSheet(sheet) {
    this.touchBook('sheets');
    const copy = new Sheet(this.newSheetId(), this.uniqueName(sheet.name));
    for (const [k, c] of sheet.cells) copy.cells.set(k, { ...c });
    copy.cols.restore(sheet.cols.snapshot());
    copy.rows.restore(sheet.rows.snapshot());
    copy.colStyles = new Map(sheet.colStyles);
    copy.rowStyles = new Map(sheet.rowStyles);
    copy.merges = sheet.merges.map(m => ({ ...m }));
    for (const p of ['freeze', 'filter', 'cf', 'validations', 'print']) copy[p] = clone(sheet[p]);
    copy.charts = clone(sheet.charts).map(c => ({ ...c, id: this.newObjectId('c') }));
    copy.images = clone(sheet.images).map(i => ({ ...i, id: this.newObjectId('i') }));
    for (const p of ['grid', 'headings', 'zoom', 'protected', 'color', 'sel']) copy[p] = sheet[p];
    this.sheets.splice(this.sheets.indexOf(sheet) + 1, 0, copy);
    for (const [k] of copy.cells) this.changes.push({ sheet: copy.id, row: keyRow(k), col: keyCol(k) });
    if (this.tx) this.tx.structural = true;
    return copy;
  }

  // ---- structure ------------------------------------------------------------------------------------------------------------

  eachFormula(visit) {
    for (const s of this.sheets) {
      for (const [k, c] of s.cells) if (c.f) visit(s, keyRow(k), keyCol(k), c);
    }
  }

  /**
   * Rows (axis 'row') or columns ('col') inserted (count > 0) or deleted (count < 0) at `at`: cells, sizes, styles, merges, conditional formats,
   * validations, the filter, charts, names and every formula of the workbook follow.
   */
  shiftLines(sheet, axis, at, count) {
    const tx = this.begin('structure');
    tx.structural = true;
    const isRow = axis === 'row';
    const max = isRow ? MAX_ROWS : MAX_COLS;
    // 1. the cells of the sheet
    const moving = [];
    for (const [k, c] of sheet.cells) {
      const i = isRow ? keyRow(k) : keyCol(k);
      if (i >= at) moving.push([k, c]);
    }
    for (const [k] of moving) { this.touchCell(sheet, k); sheet.cells.delete(k); this.changes.push({ sheet: sheet.id, row: keyRow(k), col: keyCol(k) }); }
    for (const [k, c] of moving) {
      const r = keyRow(k);
      const col = keyCol(k);
      const i = isRow ? r : col;
      if (count < 0 && i < at - count) continue; // deleted
      const ni = i + count;
      if (ni < 0 || ni >= max) continue;
      const nk = isRow ? key(ni, col) : key(r, ni);
      this.touchCell(sheet, nk);
      sheet.cells.set(nk, c);
      this.changes.push({ sheet: sheet.id, row: keyRow(nk), col: keyCol(nk) });
    }
    // new lines take the format of the line before them (like Excel)
    if (count > 0 && at > 0) {
      for (const [k, c] of [...sheet.cells]) {
        const i = isRow ? keyRow(k) : keyCol(k);
        if (i !== at - 1 || c.s === undefined) continue;
        for (let n = 0; n < count; n++) {
          const nk = isRow ? key(at + n, keyCol(k)) : key(keyRow(k), at + n);
          if (sheet.cells.has(nk)) continue;
          this.touchCell(sheet, nk);
          sheet.cells.set(nk, { s: c.s });
        }
      }
    }
    sheet.invalidate();
    // 2. sizes and line styles
    this.touchProp(sheet, isRow ? 'rows' : 'cols');
    const ax = isRow ? sheet.rows : sheet.cols;
    const styles = isRow ? sheet.rowStyles : sheet.colStyles;
    if (count > 0) {
      const inherit = at > 0 ? styles.get(at - 1) : undefined;
      ax.insert(at, count);
      const next = new Map();
      for (const [i, s] of styles) { const ni = i >= at ? i + count : i; if (ni < max) next.set(ni, s); }
      if (inherit !== undefined) for (let n = 0; n < count; n++) next.set(at + n, inherit);
      if (isRow) sheet.rowStyles = next; else sheet.colStyles = next;
    } else {
      ax.remove(at, -count);
      const next = new Map();
      for (const [i, s] of styles) { if (i >= at && i < at - count) continue; next.set(i >= at - count ? i + count : i, s); }
      if (isRow) sheet.rowStyles = next; else sheet.colStyles = next;
    }
    // 3. ranges of the sheet
    this.touchProp(sheet, 'merges');
    sheet.merges = sheet.merges.map(m => shiftArea(m, axis, at, count)).filter(m => m && !(m.r1 === m.r2 && m.c1 === m.c2));
    sheet.mergesChanged();
    const shiftText = text => { const a = parseArea(text); const n = a && shiftArea(a, axis, at, count); return n ? areaName(n) : null; };
    if (sheet.cf.length) { this.touchProp(sheet, 'cf'); sheet.cf = sheet.cf.map(r => { const range = shiftText(r.range); return range ? { ...r, range, v1: this.shiftRuleFormula(r.v1, sheet, axis, at, count, r.type === 'formula') } : null; }).filter(Boolean); }
    if (sheet.validations.length) { this.touchProp(sheet, 'validations'); sheet.validations = sheet.validations.map(v => { const range = shiftText(v.range); return range ? { ...v, range } : null; }).filter(Boolean); }
    if (sheet.filter) {
      this.touchProp(sheet, 'filter');
      const range = shiftText(sheet.filter.range);
      if (!range) sheet.filter = null;
      else {
        const cols = {};
        for (const [letters, spec] of Object.entries(sheet.filter.cols || {})) {
          let c = colIndex(letters);
          if (!isRow) { if (c >= at && count < 0 && c < at - count) continue; if (c >= at) c += count; }
          if (c >= 0) cols[colName(c)] = spec;
        }
        sheet.filter = { ...sheet.filter, range, cols };
      }
    }
    if (sheet.print?.titles && isRow) {
      this.touchProp(sheet, 'print');
      const [a, b] = sheet.print.titles.split(':').map(n => Number(n) - 1);
      const n = shiftArea({ r1: a, r2: b, c1: 0, c2: 0 }, 'row', at, count);
      sheet.print = { ...sheet.print, titles: n ? `${n.r1 + 1}:${n.r2 + 1}` : undefined };
      if (!sheet.print.titles) delete sheet.print.titles;
    }
    // 4. formulas, names and charts of the whole workbook
    this.eachFormula((s, row, col, c) => {
      const f = shiftFormula(c.f, { hostName: s.name, sheetName: sheet.name, axis, at, count });
      if (f !== c.f) this.patchCell(s, row, col, { f });
    });
    this.touchBook('names');
    this.names = this.names.map(n => ({ ...n, ref: shiftFormula(n.ref, { hostName: sheet.name, sheetName: sheet.name, axis, at, count }) }));
    for (const s of this.sheets) {
      if (!s.charts.length) continue;
      this.touchProp(s, 'charts');
      for (const ch of s.charts) ch.range = shiftFormula(ch.range, { hostName: s.name, sheetName: sheet.name, axis, at, count });
    }
    this.commit();
  }

  shiftRuleFormula(text, sheet, axis, at, count, isFormula) {
    if (!isFormula || !text) return text;
    return shiftFormula(text, { hostName: sheet.name, sheetName: sheet.name, axis, at, count });
  }

  /**
   * Moves the cells of an area (cut and paste, drag and drop) by (dRow, dCol), possibly to another sheet; what was at the target is replaced.
   * References to the moved cells follow them everywhere in the workbook.
   */
  moveCells(from, area, to, dRow, dCol) {
    const tx = this.begin('move');
    tx.structural = true;
    const taken = [];
    for (let r = area.r1; r <= area.r2; r++) {
      for (let c = area.c1; c <= area.c2; c++) {
        const cell = from.cell(r, c);
        taken.push([r, c, cell ? { ...cell } : undefined]);
        if (cell) this.setCell(from, r, c, undefined);
      }
    }
    // formulas elsewhere follow the cells (before the moved formulas are written)
    this.eachFormula((s, row, col, c) => {
      const f = moveAreaIn(c.f, { hostName: s.name, fromName: from.name, from: area, toName: to.name, dRow, dCol });
      if (f !== c.f) this.patchCell(s, row, col, { f });
    });
    for (const [r, c, cell] of taken) {
      let next = cell;
      if (next?.f) {
        let f = moveAreaIn(next.f, { hostName: from.name, fromName: from.name, from: area, toName: to.name, dRow, dCol });
        f = rehostFormula(f, from.name, to.name);
        next = { ...next, f };
      }
      this.setCell(to, r + dRow, c + dCol, next);
    }
    // merges inside the area move along
    const inside = from.merges.filter(m => m.r1 >= area.r1 && m.r2 <= area.r2 && m.c1 >= area.c1 && m.c2 <= area.c2);
    if (inside.length) {
      this.touchProp(from, 'merges');
      this.touchProp(to, 'merges');
      from.merges = from.merges.filter(m => !inside.includes(m));
      const target = { r1: area.r1 + dRow, c1: area.c1 + dCol, r2: area.r2 + dRow, c2: area.c2 + dCol };
      to.merges = to.merges.filter(m => !(m.r1 <= target.r2 && target.r1 <= m.r2 && m.c1 <= target.c2 && target.c1 <= m.c2));
      to.merges.push(...inside.map(m => ({ r1: m.r1 + dRow, c1: m.c1 + dCol, r2: m.r2 + dRow, c2: m.c2 + dCol })));
      from.mergesChanged();
      to.mergesChanged();
    }
    this.touchBook('names');
    this.names = this.names.map(n => ({ ...n, ref: moveAreaIn(n.ref, { hostName: from.name, fromName: from.name, from: area, toName: to.name, dRow, dCol }) }));
    this.commit();
  }

  /**
   * Sorts the rows of an area by keys [{ col (absolute index), desc }], with valueAt(row, col) giving the values (formulas calculated) and
   * compare(a, b) the order (blanks last). Formulas move with their rows (relative references like a copy).
   */
  sortRows(sheet, area, keys, valueAt, compare) {
    const rows = [];
    for (let r = area.r1; r <= area.r2; r++) rows.push(r);
    const order = rows.map((r, i) => ({ r, i, keys: keys.map(k => valueAt(r, k.col)) }));
    order.sort((a, b) => {
      for (let k = 0; k < keys.length; k++) {
        const x = a.keys[k];
        const y = b.keys[k];
        const blankX = x === null || x === undefined || x === '';
        const blankY = y === null || y === undefined || y === '';
        if (blankX || blankY) { if (blankX !== blankY) return blankX ? 1 : -1; continue; }
        const c = compare(x, y);
        if (c) return keys[k].desc ? -c : c;
      }
      return a.i - b.i;
    });
    if (order.every((o, i) => o.r === rows[i])) return false;
    this.begin('sort');
    const snapshot = new Map();
    for (const r of rows) for (let c = area.c1; c <= area.c2; c++) snapshot.set(key(r, c), sheet.cell(r, c));
    order.forEach((o, i) => {
      const target = rows[i];
      for (let c = area.c1; c <= area.c2; c++) {
        let cell = snapshot.get(key(o.r, c));
        if (cell?.f) cell = { ...cell, f: moveFormula(cell.f, target - o.r, 0) };
        this.setCell(sheet, target, c, cell ? { ...cell } : undefined);
      }
    });
    this.commit();
    return true;
  }

  // ---- names --------------------------------------------------------------------------------------------------------------------

  setNames(names) {
    this.touchBook('names');
    this.names = names.map(n => ({ name: n.name, ref: n.ref }));
    if (this.tx) this.tx.structural = true;
  }

  /** A defined name as an area on a sheet (or null when it is not a simple reference). */
  nameArea(name) {
    const n = this.names.find(x => sameName(x.name, name));
    if (!n) return null;
    const ref = parseReference(n.ref.replace(/\$/g, ''));
    if (!ref) return null;
    const sheet = ref.sheet ? this.sheetByName(ref.sheet) : this.activeSheet;
    return sheet ? { sheet, area: ref.area } : null;
  }
}

/** The text of an area on a sheet for a name or a chart: Sheet1!$A$1:$B$2. */
export function absoluteRef(sheet, a) {
  const abs = t => t.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  return referenceName(sheet.name, a).replace(/![^!]*$/, m => abs(m));
}
