/* MatOffice Sheets – the spreadsheet editor (like Excel): the entry module the frame loads. It holds the workbook (model.js), the
   calculation (engine/calc.js), the grid (grid.js), the editor and formula bar (editor.js) and wires the ribbon, menus, keys, clipboard,
   charts, filters, notes, dialogs, printing and the status bar together. Contract with the frame: docs/EDITORS.md. */
import { startEditor, openPopover } from '../editor/frame.js';
import { el, t, lang } from '../editor/util.js';
import { Workbook } from './model.js';
import { Calc } from './engine/calc.js';
import { setLocale, FormulaError, errorOf, numberToText } from './engine/values.js';
import { parse } from './engine/parser.js';
import { toCanonical, toLocal } from './engine/lexer.js';
import { parseInput, formatForEdit, formatValue, isDateFormat, formatKind } from './engine/numfmt.js';
import { key, keyRow, keyCol, areaName, parseArea, cellName, MAX_ROWS, MAX_COLS } from './engine/address.js';
import { moveFormula } from './engine/transform.js';
import { Selection } from './selection.js';
import { Look, validationAt, validate, listItems, cssFont } from './look.js';
import { Grid } from './grid.js';
import { CellEditor, sizeLabel } from './editor.js';
import { attachPointer, contextMenuAt } from './pointer.js';
import { attachKeys } from './keys.js';
import { Commands } from './commands.js';
import { Clipboard } from './clipboard.js';
import { Objects } from './objects.js';
import { Tabs } from './tabs.js';
import { Filters } from './filter.js';
import { Notes } from './notes.js';
import { Dialogs } from './dialogs.js';
import { ribbonTabs } from './ribbon.js';
import { printSheet, thumbnail } from './print.js';
import { FUNCTIONS } from './engine/functions/index.js';

class SheetsApp {
  constructor() {
    this.readOnly = true;
    this.showFormulas = false;
    this.showBar = localStorage.getItem('mo.sheets.bar') !== '0';
    this.painter = null;
    this.keepEditing = false;
    this.manualRows = new Set();
    this.mode = 'ready';
    this.views = new Map();
  }

  get sheet() { return this.wb.activeSheet; }

  // ---- start ---------------------------------------------------------------------------------------------------------------------

  mount(ctx) {
    this.ctx = ctx;
    this.readOnly = ctx.readOnly;
    this.locale = lang() === 'de' ? 'de' : 'en';
    setLocale(this.locale);
    this.touch = window.matchMedia('(pointer: coarse)').matches;
    this.load(ctx.doc.content);
    this.sel = new Selection();
    this.sel.mergeAt = (r, c) => this.sheet.mergeAt(r, c);
    this.look = new Look(this);
    this.grid = new Grid(this);
    this.commands = new Commands(this);
    this.editor = new CellEditor(this);
    this.clipboard = new Clipboard(this);
    this.objects = new Objects(this);
    this.tabs = new Tabs(this);
    this.filters = new Filters(this);
    this.notes = new Notes(this);
    this.dialogs = new Dialogs(this);
    this.widgets = el('div', { class: 'mo-sheets-widgets' });
    this.listButton = el('button', { type: 'button', class: 'mo-sheets-listbtn', hidden: true, 'aria-label': t('Sheets.PickFromList'), tabindex: '-1' });
    this.listButton.addEventListener('pointerdown', e => { e.preventDefault(); e.stopPropagation(); const rule = validationAt(this.sheet, this.sel.active.row, this.sel.active.col); if (rule) this.openList(rule); });
    this.prompt = el('div', { class: 'mo-sheets-prompt', hidden: true, role: 'tooltip' });
    this.tip = el('div', { class: 'mo-sheets-tip', hidden: true });
    this.grid.overlay.append(this.editor.box, this.listButton, this.prompt, this.tip);
    const main = el('div', { class: 'mo-sheets-main' }, this.grid.root);
    this.root = el('div', { class: 'mo-sheets' + (this.readOnly ? ' is-readonly' : '') + (this.showBar ? '' : ' no-bar') }, this.editor.bar, main, this.tabs.root, this.editor.popup, this.editor.tip);
    ctx.work.replaceChildren(this.root);
    attachKeys(this);
    attachPointer(this);
    this.grid.onRender = m => this.onRender(m);
    ctx.ribbon.set(this.ribbonTabs());
    ctx.status.zoom({ value: this.sheet.zoom, min: 0.1, max: 4, change: f => this.commands.setZoom(f), fit: () => this.commands.zoomToSelection() });
    this.showSheet(this.sheet, { force: true, initial: true });
    // dark mode: charts are drawn again
    new MutationObserver(() => this.objects.refresh()).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', () => this.objects.refresh());
    window.addEventListener('focus', () => { if (!this.editor.active && document.activeElement === document.body) this.focusGrid(); });
    this.focusGrid();
  }

  /** Reads the content into the workbook and calculates it. */
  load(content) {
    const name = t('Sheets.SheetName', 1);
    this.wb = Workbook.load(content, name);
    this.calc = new Calc(this.host());
    this.calc.rebuild();
    this.filtersReady = false;
  }

  /** The host of the calculation: how it reads the workbook. */
  host() {
    const byId = () => {
      if (this.sheetMapOf !== this.wb.sheets || this.sheetMapLen !== this.wb.sheets.length) {
        this.sheetMap = new Map(this.wb.sheets.map((s, i) => [s.id, s]));
        this.sheetMapOf = this.wb.sheets;
        this.sheetMapLen = this.wb.sheets.length;
      }
      return this.sheetMap;
    };
    return {
      cell: (id, row, col) => byId().get(id)?.cells.get(key(row, col)),
      sheetId: name => this.wb.sheetByName(name)?.id ?? null,
      sheetName: id => byId().get(id)?.name ?? null,
      sheetIds: () => this.wb.sheets.map(s => s.id),
      sheetIndex: id => this.wb.sheets.findIndex(s => s.id === id),
      bounds: id => byId().get(id)?.bounds() ?? { rows: 0, cols: 0 },
      formulas: visit => this.wb.eachFormula((s, r, c, cell) => visit(s.id, r, c, cell.f)),
      names: () => this.wb.names,
      rowHidden: (id, row) => !!byId().get(id)?.rows.isHidden(row),
      setResult: (id, row, col, value) => { const s = byId().get(id); if (s) this.wb.setResult(s, row, col, value); },
    };
  }

  ribbonTabs() { return ribbonTabs(this); }

  // ---- sheets ------------------------------------------------------------------------------------------------------------------

  /** Shows a sheet: its selection, scroll position and zoom come back. */
  showSheet(sheet, { keepEditor = false, force = false, initial = false } = {}) {
    if (!sheet) return;
    const before = this.wb.activeSheet;
    if (!force && sheet === before && !initial) return;
    if (before && this.sel && !initial) this.views.set(before.id, { sel: this.sel.toJSON(), top: this.grid.scroller.scrollTop, left: this.grid.scroller.scrollLeft });
    if (before && !initial && this.sel) before.sel = areaName(this.sel.range).replace(/^(\d+):(\d+)$/, 'A$1:A$2');
    this.wb.active = this.wb.sheets.indexOf(sheet);
    this.objects.deselect();
    const view = this.views.get(sheet.id);
    if (view) this.sel.restore(view.sel);
    else {
      const a = parseArea(sheet.sel || 'A1') || { r1: 0, c1: 0, r2: 0, c2: 0 };
      this.sel.setArea(a.r2 - a.r1 > 100000 || a.c2 - a.c1 > 1000 ? { r1: a.r1, c1: a.c1, r2: a.r1, c2: a.c1 } : a);
    }
    this.sel.sheetId = sheet.id;
    if (sheet.filter) this.filters.apply(sheet);
    this.grid.extentRows = 200;
    this.grid.extentCols = 40;
    this.grid.themeCache = null;
    this.grid.updateExtent();
    this.grid.scroller.scrollTop = view?.top || 0;
    this.grid.scroller.scrollLeft = view?.left || 0;
    if (!view) this.grid.scrollIntoView(this.sel.active.row, this.sel.active.col);
    this.ctx.status.setZoom(sheet.zoom);
    this.tabs.render();
    this.objects.refresh();
    this.notes.render();
    this.look.invalidate();
    if (!keepEditor && this.editor.active && this.editor.sheet !== sheet) this.editor.commit(0, 0);
    this.grid.render();
    this.selectionChanged({ scroll: false });
    if (this.editor.active) this.editor.input.focus();
  }

  // ---- changes -------------------------------------------------------------------------------------------------------------------

  /** After every change of the content: calculate, draw, save. */
  changed({ structural = false, layout = false } = {}) {
    const changes = this.wb.takeChanges();
    if (structural) this.calc.rebuild();
    else if (changes.length) this.calc.update(changes);
    this.ctx.changed();
    this.afterCalc();
    if (layout) this.grid.updateExtent();
  }

  afterCalc() {
    this.look.invalidate();
    this.grid.updateExtent();
    this.grid.requestRender();
    this.objects.refresh();
    this.notes.render();
    this.tabs.render();
    this.editor.refresh();
    this.updateStatus();
    this.refreshRibbon();
  }

  zoomChanged() {
    this.ctx.status.setZoom(this.sheet.zoom);
    this.look.invalidate();
    this.grid.updateExtent();
    this.grid.render();
    this.objects.position();
    this.editor.styleBox?.call(this.editor);
  }

  undo() {
    if (this.editor.active) return;
    const e = this.wb.undo();
    if (!e) return;
    this.restoreAfterHistory(e, e.selBefore);
  }

  redo() {
    if (this.editor.active) return;
    const e = this.wb.redo();
    if (!e) return;
    this.restoreAfterHistory(e, e.selAfter || e.selBefore);
  }

  restoreAfterHistory(e, sel) {
    const sheetOf = e.cells[0]?.sheet || e.props[0]?.sheet;
    if (sheetOf && sheetOf !== this.sheet && this.wb.sheets.includes(sheetOf)) this.showSheet(sheetOf);
    if (sel) this.sel.restore(sel);
    this.sheet.mergesChanged();
    this.changed({ structural: e.structural || e.book.length > 0 || e.props.some(p => ['cols', 'rows', 'filter'].includes(p.prop)) });
    if (this.wb.activeSheet !== this.sheet || !this.wb.sheets.includes(this.sheet)) this.showSheet(this.wb.activeSheet, { force: true });
    this.objects.refresh();
    this.selectionChanged();
  }

  refreshRibbon() {
    if (this.ribbonPending) return;
    this.ribbonPending = requestAnimationFrame(() => { this.ribbonPending = 0; this.ctx.ribbon.refresh(); });
  }

  /** The selection changed: draw, scroll, formula bar, status bar, ribbon. */
  selectionChanged({ scroll = true, quick = false, scrollTo = null } = {}) {
    this.sel.sheetId = this.sheet.id;
    if (scroll) {
      const target = scrollTo || this.sel.active;
      this.grid.scrollIntoView(target.row, target.col);
    }
    this.grid.updateExtent();
    this.grid.requestRender();
    if (!this.editor.active) this.editor.refresh();
    this.placeCatcher();
    this.updateCellWidgets();
    if (!quick) { this.updateStatus(); this.refreshRibbon(); }
    else this.updateStatusSoon();
  }

  // ---- typing -------------------------------------------------------------------------------------------------------------------

  /** The text of a cell in the editor and formula bar. */
  editText(sheet, row, col) {
    const cell = sheet.cell(row, col);
    if (!cell) {
      const fc = this.calc.formulaAt(sheet.id, row, col);
      if (fc && fc.row !== row || fc && fc.col !== col) return '';
      return '';
    }
    if (cell.f) return '=' + toLocal(cell.f, this.locale);
    if (cell.v === undefined || cell.v === null) return '';
    if (cell.t === 'e') return String(cell.v);
    return formatForEdit(cell.v, this.wb.style(sheet.styleAt(row, col)).nf, this.locale);
  }

  /** What a typed text becomes in a cell; false keeps the editor open (a formula with an error, a value the validation refuses). */
  commitInput(sheet, row, col, text, options = {}) {
    const style = this.wb.style(sheet.styleAt(row, col));
    let parsed = parseInput(text, this.locale);
    if (style.nf === '@' && parsed.formula === undefined && !parsed.empty) parsed = { value: text };
    let formula;
    if (parsed.formula !== undefined) {
      formula = this.checkFormula(parsed.formula);
      if (formula === null) {
        if (!options.quiet) this.ctx.dialog({ title: t('Sheets.Formula.ProblemTitle'), icon: 'alert', body: el('p', { text: t('Sheets.Formula.Problem') }) }).then(() => this.editor.input.focus());
        return false;
      }
    }
    // data validation
    if (!options.force && formula === undefined) {
      const rule = validationAt(sheet, row, col);
      const value = parsed.empty ? null : parsed.error ? null : parsed.value;
      if (rule && !validate(this, sheet, row, col, value, rule)) {
        if (options.quiet) return false;
        this.validationError(rule, options);
        return false;
      }
    }
    const write = (r, c, f) => {
      const existing = sheet.cell(r, c) || {};
      const next = { ...existing, v: undefined, f: undefined, t: undefined };
      if (f !== undefined) next.f = f;
      else if (parsed.error) { next.v = parsed.error; next.t = 'e'; } else if (!parsed.empty) {
        next.v = parsed.value;
        const s = sheet.styleAt(r, c);
        const nf = this.wb.style(s).nf;
        if (parsed.nf && (!nf || (formatKind(parsed.nf) !== formatKind(nf) && ['date', 'time', 'datetime'].includes(formatKind(parsed.nf)) && !isDateFormat(nf)))) next.s = this.wb.patchStyle(s, { nf: parsed.nf });
      }
      this.wb.setCell(sheet, r, c, next);
    };
    const run = () => {
      if (options.all && !this.sel.isSingle) {
        for (const a0 of this.sel.ranges) {
          const a = this.commands.clip(a0);
          for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) write(r, c, formula !== undefined ? moveFormula(formula, r - row, c - col) : undefined);
        }
      } else write(row, col, formula);
    };
    if (options.quiet && this.wb.tx) { run(); return true; }
    const ok = this.commands.change('typing', run, { editing: true, protect: !!options.all });
    if (ok === false) return false;
    if (style.wrap || style.size > 11 || String(text).includes('\n')) this.autoFitRows([{ r1: row, r2: row, c1: 0, c2: MAX_COLS - 1 }]);
    return true;
  }

  /** A typed formula → canonical text, closing parentheses and quotes that were left open; null when it cannot be read. */
  checkFormula(text) {
    const canon = toCanonical(text, this.locale);
    try { parse(canon); return canon; } catch { /* try to complete it */ }
    let fixed = canon;
    const quotes = (fixed.match(/"/g) || []).length;
    if (quotes % 2) fixed += '"';
    let depth = 0;
    let inStr = false;
    for (const ch of fixed) { if (ch === '"') inStr = !inStr; else if (!inStr && ch === '(') depth++; else if (!inStr && ch === ')') depth--; }
    if (depth > 0) fixed += ')'.repeat(depth);
    try { parse(fixed); return fixed; } catch { return null; }
  }

  async validationError(rule, options) {
    const title = rule.errorTitle || t('Sheets.Valid.DefaultTitle');
    const message = rule.error || t('Sheets.Valid.DefaultText');
    const style = rule.errorStyle || 'stop';
    const buttons = style === 'stop'
      ? [{ label: t('Sheets.Valid.Retry'), value: 'retry', kind: 'primary' }, { label: t('Sheets.Cancel'), value: 'cancel' }]
      : style === 'warning'
        ? [{ label: t('Sheets.Valid.Continue'), value: 'yes', kind: 'primary' }, { label: t('Sheets.Valid.Edit'), value: 'retry' }, { label: t('Sheets.Cancel'), value: 'cancel' }]
        : [{ label: t('Sheets.Ok'), value: 'yes', kind: 'primary' }, { label: t('Sheets.Cancel'), value: 'cancel' }];
    this.keepEditing = true;
    const r = await this.ctx.dialog({ title, icon: style === 'stop' ? 'x-circle' : style === 'warning' ? 'alert' : 'info', body: el('p', { text: message }), buttons });
    this.keepEditing = false;
    if (r === 'yes') { this.editor.commit(options.dr ?? 1, options.dc ?? 0, { ...options, force: true }); return; }
    if (r === 'retry') { this.editor.input.focus(); this.editor.input.select(); return; }
    this.editor.cancel();
  }

  /** After Enter/Tab: the active cell moves (inside a selection of several cells it walks through it; Tab … Enter returns to the first column). */
  afterCommit(dr, dc) {
    const sheet = this.sheet;
    if (!dr && !dc) { this.selectionChanged({ scroll: false }); this.focusGrid(); return; }
    if (this.sel.cycle(dr, dc, r => !sheet.rows.isHidden(r), c => !sheet.cols.isHidden(c))) { this.selectionChanged(); this.focusGrid(); return; }
    const { row, col } = this.sel.active;
    const m = sheet.mergeAt(row, col);
    if (dc) { if (this.tabStart === undefined) this.tabStart = col; } else if (dr > 0 && this.tabStart !== undefined) {
      const c = this.tabStart;
      this.tabStart = undefined;
      this.sel.set(sheet.rows.step(m ? m.r2 : row, 1), c);
      this.selectionChanged();
      this.focusGrid();
      return;
    } else this.tabStart = undefined;
    const r = dr ? sheet.rows.step(dr > 0 && m ? m.r2 : row, dr) : row;
    const c = dc ? sheet.cols.step(dc > 0 && m ? m.c2 : col, dc) : col;
    this.sel.set(Math.max(0, Math.min(MAX_ROWS - 1, r)), Math.max(0, Math.min(MAX_COLS - 1, c)));
    this.selectionChanged();
    this.focusGrid();
  }

  /** A text read like typed input → the fields of a cell (for Text to Columns). */
  parseTyped(text) {
    const p = parseInput(text, this.locale);
    if (p.formula !== undefined) return { v: text, f: undefined, t: undefined };
    if (p.error) return { v: p.error, t: 'e', f: undefined };
    if (p.empty) return { v: undefined, f: undefined, t: undefined };
    return { v: p.value, f: undefined, t: undefined };
  }

  formatText(value, nf) { return formatValue(value, nf, { locale: this.locale }).text; }

  setMode(mode) {
    this.mode = mode;
    this.updateStatus();
  }

  isLocked(sheet, row, col) {
    return !!sheet.protected && !this.wb.style(sheet.styleAt(row, col)).unlocked;
  }

  toastReadOnly() { this.ctx.toast(t('Sheets.ReadOnly'), 'info'); }
  toastLocked() { this.ctx.toast(t('Sheets.Locked.Toast'), 'warning'); }

  /** The last used cell (Ctrl+End). */
  usedArea(sheet) {
    const b = sheet.bounds();
    const e = this.calc.extent(sheet.id);
    return { r2: Math.max(0, Math.max(b.rows, e.rows) - 1), c2: Math.max(0, Math.max(b.cols, e.cols) - 1) };
  }

  nameBoxText() {
    const a = this.sel.range;
    const sheet = this.sheet;
    const named = this.wb.names.find(n => { const x = this.wb.nameArea(n.name); return x && x.sheet === sheet && x.area.r1 === a.r1 && x.area.c1 === a.c1 && x.area.r2 === a.r2 && x.area.c2 === a.c2; });
    if (named) return named.name;
    if (this.sel.isCellArea(a)) return cellName(this.sel.active.row, this.sel.active.col);
    if (this.sel.isWholeCol(a) || this.sel.isWholeRow(a)) return areaName(a);
    return cellName(this.sel.active.row, this.sel.active.col);
  }

  /** The spill area of the formula at the active cell (drawn with a blue border). */
  spillOf() {
    const fc = this.calc.formulaAt(this.sheet.id, this.sel.active.row, this.sel.active.col);
    return fc?.spill ? { sheet: fc.sheet, area: fc.spill } : null;
  }

  /** Rows grow to their content (wrapped text, larger fonts) unless their height was set by hand. */
  autoFitRows(ranges, force = false) {
    const sheet = this.sheet;
    const rows = new Set();
    for (const a of ranges) {
      if (a.r2 - a.r1 > 5000) continue;
      for (let r = a.r1; r <= a.r2; r++) rows.add(r);
    }
    if (!rows.size) return;
    const heights = new Map();
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k);
      if (!rows.has(r)) continue;
      const c = keyCol(k);
      if (sheet.mergeAt(r, c)) continue;
      const style = this.wb.style(sheet.styleAt(r, c));
      const fontPx = (style.size || 11) * 96 / 72;
      let h = Math.ceil(fontPx * 1.36);
      if (style.wrap || (typeof cell.v === 'string' && cell.v.includes('\n'))) {
        const look = this.look.cell(sheet, r, c, 11);
        if (look.text) {
          const width = Math.max(10, sheet.cols.size(c) - 6 - (style.indent || 0) * 9);
          const lines = style.wrap ? this.grid.wrapLines(look.text, cssFont(style, 1), width).length : look.text.split('\n').length;
          h = Math.ceil(lines * fontPx * 1.22 + 6);
        }
      }
      heights.set(r, Math.max(heights.get(r) || 0, h));
    }
    let changed = false;
    this.wb.begin('row height');
    this.wb.touchProp(sheet, 'rows');
    for (const r of rows) {
      if (!force && this.manualRows.has(r)) continue;
      if (!force && sheet.rows.sizes.has(r) && !this.autoRows?.has(r) && !heights.has(r)) continue;
      const need = Math.max(sheet.rows.def, Math.min(409 * 4 / 3, heights.get(r) || 0));
      const now = sheet.rows.rawSize(r);
      if (force ? need !== now : need > now || (this.autoRows?.has(r) && need !== now)) {
        sheet.rows.setSize(r, need);
        (this.autoRows ||= new Set()).add(r);
        changed = true;
      }
    }
    this.wb.commit();
    if (changed) { this.grid.updateExtent(); this.grid.requestRender(); this.ctx.changed(); }
  }

  // ---- merges --------------------------------------------------------------------------------------------------------------------

  isMerged() {
    const a = this.sel.range;
    return this.sheet.merges.some(m => m.r1 <= a.r2 && a.r1 <= m.r2 && m.c1 <= a.c2 && a.c1 <= m.c2);
  }

  /** Merge & center, merge across, merge, unmerge. */
  merge(kind, inTransaction = false) {
    const sheet = this.sheet;
    const run = () => {
      const merges = sheet.merges.filter(m => !this.sel.ranges.some(a => m.r1 <= a.r2 && a.r1 <= m.r2 && m.c1 <= a.c2 && a.c1 <= m.c2));
      const unmerging = kind === 'unmerge' || (kind === 'center' && this.isMerged());
      if (!unmerging) {
        for (const a0 of this.sel.ranges) {
          const a = this.commands.clip(a0);
          if (kind === 'across') { for (let r = a.r1; r <= a.r2; r++) if (a.c2 > a.c1) merges.push({ r1: r, c1: a.c1, r2: r, c2: a.c2 }); } else if (a.r2 > a.r1 || a.c2 > a.c1) merges.push({ ...a });
          // only the value of the first cell stays (like Excel)
          for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
            if (kind === 'across' ? c === a.c1 : r === a.r1 && c === a.c1) continue;
            const cell = sheet.cell(r, c);
            if (cell && (cell.v !== undefined || cell.f)) this.wb.patchCell(sheet, r, c, { v: undefined, f: undefined, t: undefined });
          }
          if (kind === 'center') this.commands.applyStyle(sheet, [{ r1: a.r1, c1: a.c1, r2: a.r1, c2: a.c1 }], s => ({ ...s, ha: 'center' }));
        }
      }
      this.wb.setSheetProp(sheet, 'merges', merges);
    };
    if (inTransaction) run(); else this.commands.change('merge', run);
    this.sel.setArea(this.sel.range, this.sel.active);
    this.selectionChanged({ scroll: false });
  }

  selectedCols() {
    const out = [];
    for (const a of this.sel.ranges) {
      const c2 = Math.min(a.c2, a.c1 + 500);
      for (let c = a.c1; c <= c2; c++) out.push(c);
    }
    return out;
  }

  /** Find & Select: formulas, notes, constants or cells with a validation of the sheet. */
  selectSpecial(kind) {
    const sheet = this.sheet;
    const hits = [];
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      const ok = kind === 'formulas' ? !!cell.f : kind === 'notes' ? !!cell.note : kind === 'constants' ? !cell.f && cell.v !== undefined : !!validationAt(sheet, r, c);
      if (ok) hits.push({ r1: r, c1: c, r2: r, c2: c });
    }
    if (!hits.length) { this.ctx.toast(t('Sheets.Select.None'), 'info'); return; }
    hits.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
    this.sel.ranges = hits.slice(0, 2000);
    this.sel.active = { row: hits[0].r1, col: hits[0].c1 };
    this.sel.anchor = { ...this.sel.active };
    this.selectionChanged();
  }

  jumpNote(d) {
    const sheet = this.sheet;
    const keys = [...sheet.cells.entries()].filter(([, c]) => c.note).map(([k]) => k).sort((a, b) => a - b);
    if (!keys.length) return;
    const cur = key(this.sel.active.row, this.sel.active.col);
    const next = d > 0 ? keys.find(k => k > cur) ?? keys[0] : [...keys].reverse().find(k => k < cur) ?? keys[keys.length - 1];
    this.sel.set(keyRow(next), keyCol(next));
    this.selectionChanged();
  }

  insertFunctionName(name) {
    if (this.readOnly) return;
    const ed = this.editor;
    if (ed.active) { ed.insertText(ed.input, name + '('); ed.input.focus(); return; }
    ed.begin({ text: '=' + name + '()', mode: 'edit' });
    const i = ed.input;
    i.setSelectionRange(i.value.length - 1, i.value.length - 1);
    ed.afterCaret(i);
  }

  insertNameIntoFormula(name) {
    const ed = this.editor;
    if (ed.active) { ed.insertText(ed.input, name); ed.input.focus(); return; }
    if (!this.readOnly) ed.begin({ text: '=' + name, mode: 'enter' });
  }

  toggleBar() {
    this.showBar = !this.showBar;
    try { localStorage.setItem('mo.sheets.bar', this.showBar ? '1' : '0'); } catch { /* private */ }
    this.root.classList.toggle('no-bar', !this.showBar);
    this.grid.resize();
  }

  // ---- widgets on the grid ---------------------------------------------------------------------------------------------------------

  onRender(m) {
    this.notes.position();
    this.updateCellWidgets(m);
  }

  /** The list button of a validation and its input message next to the active cell. */
  updateCellWidgets(m) {
    const { row, col } = this.sel.active;
    const rule = this.editor.active ? null : validationAt(this.sheet, row, col);
    const r = this.grid.cellRect(row, col, m || this.grid.metrics());
    const visible = r.x + r.w > 0 && r.y + r.h > 0 && r.x < this.grid.width && r.y < this.grid.height;
    const showList = !!rule && rule.type === 'list' && rule.dropdown !== false && visible && !this.readOnly;
    this.listButton.hidden = !showList;
    if (showList) {
      const size = Math.max(16, Math.min(22, r.h));
      Object.assign(this.listButton.style, { left: Math.round(r.x + r.w + 1) + 'px', top: Math.round(r.y + r.h - size) + 'px', width: size + 'px', height: size + 'px' });
    }
    const showPrompt = !!rule?.prompt && visible;
    this.prompt.hidden = !showPrompt;
    if (showPrompt) {
      this.prompt.textContent = rule.prompt;
      Object.assign(this.prompt.style, { left: Math.round(r.x + 10) + 'px', top: Math.round(r.y + r.h + 6) + 'px' });
    }
  }

  /** The list of a validation under the active cell. */
  openList(rule) {
    const sheet = this.sheet;
    const { row, col } = this.sel.active;
    const items = listItems(this, sheet, rule);
    if (!items.length || this.readOnly) return;
    const anchor = this.listButton.hidden ? this.grid.root : this.listButton;
    const current = this.calc.valueAt(sheet.id, row, col);
    openPopover(anchor, close => {
      const list = el('div', { class: 'mo-menu mo-sheets-list', role: 'listbox' });
      for (const item of items) {
        list.append(el('button', {
          type: 'button', class: 'mo-menu__item', role: 'option', 'aria-selected': String(String(current) === item),
          onclick: () => { close(); this.commitInput(sheet, row, col, item, { force: true }); this.focusGrid(); },
        }, el('span', { class: 'mo-menu__label', text: item })));
      }
      return list;
    });
  }

  showTip(p, text) {
    this.tip.hidden = false;
    this.tip.textContent = text;
    this.tip.style.left = Math.round(p.x + 12) + 'px';
    this.tip.style.top = Math.round(p.y + 12) + 'px';
  }

  hideTip() { this.tip.hidden = true; }

  widthLabel(px) { return t('Sheets.WidthTip', numberToText(Math.max(0, Math.round(((px - 5) / 7) * 100) / 100)), px); }
  heightLabel(px) { return t('Sheets.HeightTip', numberToText(Math.round(px * 0.75 * 100) / 100), px); }

  contextMenu(x, y, items) { contextMenuAt(this, x, y, items); }
  menuAt(anchor, items) { openPopover(anchor, close => this.ctx.ribbon.menuItems(items, close)); }

  // ---- focus, status -----------------------------------------------------------------------------------------------------------

  /** The keys go to the grid (not on phones: the keyboard would open on every tap). */
  focusGrid(force = false) {
    if (this.editor?.active) return;
    if (this.touch && !force) { if (document.activeElement && document.activeElement !== document.body && this.root.contains(document.activeElement)) document.activeElement.blur(); return; }
    if (document.activeElement !== this.keys) this.keys?.focus({ preventScroll: true });
  }

  placeCatcher() {
    if (!this.keys) return;
    // screen readers hear the active cell (the grid itself is a picture)
    const { row, col } = this.sel.active;
    const shown = this.look.cell(this.sheet, row, col, 11).text;
    this.keys.setAttribute('aria-label', `${t('Sheets.Grid')}: ${cellName(row, col)}${shown ? ' ' + shown : ''}`);
    const r = this.grid.cellRect(this.sel.active.row, this.sel.active.col);
    this.keys.style.left = Math.max(0, Math.min(this.grid.width - 2, Math.round(r.x))) + 'px';
    this.keys.style.top = Math.max(0, Math.min(this.grid.height - 2, Math.round(r.y))) + 'px';
  }

  updateStatusSoon() {
    clearTimeout(this.statusTimer);
    this.statusTimer = setTimeout(() => this.updateStatus(), 120);
  }

  /** Mode (Ready, Enter, Edit, Point) and Average, Count, Sum of the selection. */
  updateStatus() {
    const status = this.ctx.status;
    const parts = [el('span', { class: 'mo-statusbar__item mo-sheets-mode', text: t('Sheets.Mode.' + this.mode) })];
    if (this.readOnly) parts.push(el('span', { class: 'mo-statusbar__item', text: t('Sheets.ReadOnlyMode') }));
    if (this.sheet.protected) parts.push(el('span', { class: 'mo-statusbar__item', text: t('Sheets.ProtectedMode') }));
    const sel = this.sel;
    if (!sel.isSingle) {
      let count = 0;
      let nums = 0;
      let sum = 0;
      const sheet = this.sheet;
      const seen = new Set();
      const visit = (r, c) => {
        const k = key(r, c);
        if (seen.has(k)) return;
        seen.add(k);
        if (sheet.rows.isHidden(r) || sheet.cols.isHidden(c)) return;
        const v = this.calc.valueAt(sheet.id, r, c);
        if (v === null || v === undefined || v === '') return;
        count++;
        if (typeof v === 'number') { nums++; sum += v; }
      };
      let area = 0;
      for (const a0 of sel.ranges) { const a = this.commands.clip(a0); area += (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1); }
      if (area <= 300000) {
        for (const a0 of sel.ranges) { const a = this.commands.clip(a0); for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) visit(r, c); }
      } else {
        for (const k of sheet.cells.keys()) { const r = keyRow(k); const c = keyCol(k); if (sel.contains(r, c)) visit(r, c); }
      }
      const nf = this.wb.style(sheet.styleAt(sel.active.row, sel.active.col)).nf;
      const fmt = v => formatValue(v, nf && !['text'].includes(formatKind(nf)) ? nf : undefined, { locale: this.locale }).text;
      if (nums) parts.push(el('span', { class: 'mo-statusbar__item', text: t('Sheets.Status.Average', fmt(sum / nums)) }));
      if (count) parts.push(el('span', { class: 'mo-statusbar__item', text: t('Sheets.Status.Count', count) }));
      if (nums) parts.push(el('span', { class: 'mo-statusbar__item', text: t('Sheets.Status.Sum', fmt(sum)) }));
    }
    status.left(...parts);
  }

  // ---- frame hooks ---------------------------------------------------------------------------------------------------------------

  getContent() {
    if (this.sel) this.sheet.sel = this.sel.isWholeCol() || this.sel.isWholeRow() ? cellName(this.sel.active.row, this.sel.active.col) : areaName(this.sel.range);
    return this.wb.toContent();
  }

  setContent(content) {
    const activeId = this.sheet?.id;
    const sel = this.sel?.toJSON();
    if (this.editor?.active) this.editor.cancel();
    this.load(content);
    const sheet = this.wb.sheetById(activeId) || this.wb.activeSheet;
    this.wb.active = this.wb.sheets.indexOf(sheet);
    this.views.clear();
    this.showSheet(sheet, { force: true, initial: true });
    if (sel) { this.sel.restore(sel); this.selectionChanged({ scroll: false }); }
    this.ctx.ribbon.set(this.ribbonTabs());
  }

  print() {
    if (this.editor.active) this.editor.commit(0, 0);
    printSheet(this);
  }

  info() {
    let cells = 0;
    let formulas = 0;
    for (const s of this.wb.sheets) for (const c of s.cells.values()) { if (c.v !== undefined || c.f) cells++; if (c.f) formulas++; }
    return el('dl', { class: 'mat-kv' },
      el('dt', { text: t('Sheets.Info.Sheets') }), el('dd', { text: String(this.wb.sheets.length) }),
      el('dt', { text: t('Sheets.Info.Cells') }), el('dd', { text: String(cells) }),
      el('dt', { text: t('Sheets.Info.Formulas') }), el('dd', { text: String(formulas) }),
      el('dt', { text: t('Sheets.Info.Functions') }), el('dd', { text: String(FUNCTIONS.size) }));
  }

  thumbnail() { return thumbnail(this); }
}

const app = new SheetsApp();
// for the browser tests (tests/js and the scripts of the handoff): the editor of the page
Object.defineProperty(window, '__sheets', { value: app });

startEditor({
  mount: ctx => app.mount(ctx),
  getContent: () => app.getContent(),
  setContent: content => app.setContent(content),
  thumbnail: () => app.thumbnail(),
  print: () => app.print(),
  info: () => app.info(),
});

export { app, FormulaError, errorOf, sizeLabel };
