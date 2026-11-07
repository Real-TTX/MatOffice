/* MatOffice Sheets – typing into cells: the editor in the cell and the formula bar (they show the same text), the name box.
   Like Excel: typing replaces the cell ("Enter" mode: arrows leave the cell or point at cells while a formula expects a reference), F2 and a
   double click edit it ("Edit" mode: arrows move the caret), Esc cancels. While a formula is edited its references are coloured in the text and
   on the grid; clicking or dragging over cells (also on other sheets) inserts references; F4 cycles $; function names are suggested and the
   arguments of the function at the caret are shown. */
import { el, icon, t } from '../editor/util.js';
import { FUNCTIONS } from './engine/functions/index.js';
import { referencesOf } from './engine/transform.js';
import { tokenize } from './engine/lexer.js';
import { areaName, sheetPrefix, colName, MAX_ROWS, MAX_COLS } from './engine/address.js';
import { cssFont } from './look.js';

export const REF_COLORS = ['#4472c4', '#c0504d', '#7030a0', '#00a65a', '#d29200', '#0070c0', '#a5509b', '#e46c0a'];
const POINT_AFTER = '=(,;+-*/^&<>:';

export class CellEditor {
  constructor(app) {
    this.app = app;
    this.active = false;
    this.mode = 'enter';
    this.point = null;
    this.suggest = null;
    // the formula bar
    this.nameBox = el('input', { class: 'mo-sheets-namebox', 'aria-label': t('Sheets.NameBox'), 'data-tip': t('Sheets.NameBox'), autocomplete: 'off', spellcheck: 'false' });
    this.nameList = el('button', { type: 'button', class: 'mo-sheets-fbtn mo-sheets-namelist', 'aria-label': t('Sheets.Names'), 'data-tip': t('Sheets.Names') }, icon('chevron-down'));
    this.cancelBtn = el('button', { type: 'button', class: 'mo-sheets-fbtn', 'aria-label': t('Sheets.Cancel'), 'data-tip': t('Sheets.Cancel') }, icon('x'));
    this.enterBtn = el('button', { type: 'button', class: 'mo-sheets-fbtn', 'aria-label': t('Sheets.Enter'), 'data-tip': t('Sheets.Enter') }, icon('check'));
    this.fxBtn = el('button', { type: 'button', class: 'mo-sheets-fbtn mo-sheets-fx', 'aria-label': t('Sheets.InsertFunction'), 'data-tip': t('Sheets.InsertFunction') }, el('span', { text: 'fx' }));
    this.barInput = el('textarea', { class: 'mo-sheets-finput', rows: 1, 'aria-label': t('Sheets.FormulaBar'), spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off' });
    this.barBackdrop = el('div', { class: 'mo-sheets-backdrop mo-sheets-backdrop--bar', 'aria-hidden': 'true' });
    this.expandBtn = el('button', { type: 'button', class: 'mo-sheets-fbtn', 'aria-label': t('Sheets.ExpandBar'), 'data-tip': t('Sheets.ExpandBar') }, icon('chevron-down'));
    this.bar = el('div', { class: 'mo-sheets-fbar' },
      el('div', { class: 'mo-sheets-namewrap' }, this.nameBox, this.nameList),
      el('div', { class: 'mo-sheets-fbtns' }, this.cancelBtn, this.enterBtn, this.fxBtn),
      el('div', { class: 'mo-sheets-fwrap' }, this.barBackdrop, this.barInput),
      this.expandBtn);
    // the editor in the cell
    this.cellInput = el('textarea', { class: 'mo-sheets-cinput', rows: 1, spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off', 'aria-label': t('Sheets.CellEditor') });
    this.cellBackdrop = el('div', { class: 'mo-sheets-backdrop', 'aria-hidden': 'true' });
    this.box = el('div', { class: 'mo-sheets-editor', hidden: true }, this.cellBackdrop, this.cellInput);
    this.popup = el('div', { class: 'mo-sheets-suggest', role: 'listbox', hidden: true });
    this.tip = el('div', { class: 'mo-sheets-argtip', hidden: true });

    for (const input of [this.cellInput, this.barInput]) {
      input.addEventListener('input', () => this.onInput(input));
      input.addEventListener('keydown', e => this.onKey(e, input));
      input.addEventListener('keyup', () => this.afterCaret(input));
      input.addEventListener('click', () => { if (this.active) { this.point = null; this.afterCaret(input); } });
      input.addEventListener('blur', () => setTimeout(() => this.onBlur(), 0));
      input.addEventListener('scroll', () => { (input === this.cellInput ? this.cellBackdrop : this.barBackdrop).scrollTop = input.scrollTop; });
    }
    this.barInput.addEventListener('focus', () => { if (!this.active && !this.app.readOnly) this.begin({ mode: 'edit', where: 'bar' }); });
    this.cancelBtn.addEventListener('click', () => this.cancel());
    this.enterBtn.addEventListener('click', () => this.commit(0, 0));
    this.fxBtn.addEventListener('click', () => this.app.commands.insertFunction());
    this.expandBtn.addEventListener('click', () => { this.bar.classList.toggle('is-expanded'); this.expandBtn.firstChild.replaceWith(icon(this.bar.classList.contains('is-expanded') ? 'chevron-up' : 'chevron-down')); this.app.grid.resize(); });
    this.nameBox.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); this.app.commands.goToText(this.nameBox.value); }
      else if (e.key === 'Escape') { this.refresh(); this.app.focusGrid(); }
    });
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.nameBox.addEventListener('blur', () => this.refresh());
    this.nameList.addEventListener('click', () => this.app.commands.namesMenu(this.nameList));
    this.updateButtons();
  }

  get input() { return this.where === 'bar' ? this.barInput : this.cellInput; }

  // ---- showing the active cell --------------------------------------------------------------------------------------------------

  /** The formula bar and the name box show the active cell (when nothing is edited). */
  refresh() {
    if (document.activeElement !== this.nameBox) this.nameBox.value = this.app.nameBoxText();
    if (this.active) return;
    const text = this.app.editText(this.app.sheet, this.app.sel.active.row, this.app.sel.active.col);
    this.barInput.value = text;
    this.paintBackdrop(this.barBackdrop, text);
    this.barInput.readOnly = this.app.readOnly || this.app.isLocked(this.app.sheet, this.app.sel.active.row, this.app.sel.active.col);
    this.updateButtons();
  }

  updateButtons() {
    this.cancelBtn.disabled = !this.active;
    this.enterBtn.disabled = !this.active;
    this.fxBtn.disabled = this.app.readOnly;
  }

  // ---- editing -----------------------------------------------------------------------------------------------------------------

  /**
   * Starts editing the active cell. options: text (replaces the content; undefined keeps it), mode 'enter' | 'edit', where 'cell' | 'bar'.
   */
  begin({ text, mode = 'edit', where = 'cell' } = {}) {
    const app = this.app;
    if (app.readOnly) return false;
    const { row, col } = app.sel.active;
    if (app.isLocked(app.sheet, row, col)) { app.toastLocked(); return false; }
    this.active = true;
    this.mode = mode;
    this.where = where;
    this.sheet = app.sheet;
    this.row = row;
    this.col = col;
    this.original = app.editText(app.sheet, row, col);
    const value = text === undefined ? this.original : text;
    this.point = null;
    this.cellInput.value = value;
    this.barInput.value = value;
    this.barInput.readOnly = false;
    this.box.hidden = false;
    this.styleBox();
    this.position();
    const input = this.input;
    input.focus({ preventScroll: true });
    const end = value.length;
    input.setSelectionRange(end, end);
    this.onInput(input, true);
    this.updateButtons();
    app.setMode(mode === 'enter' ? 'enter' : 'edit');
    app.grid.requestRender();
    return true;
  }

  /** The look of the cell editor follows the cell (font, size, colours, alignment). */
  styleBox() {
    const app = this.app;
    const style = app.wb.style(this.sheet.styleAt(this.row, this.col));
    const z = this.sheet.zoom;
    const font = cssFont(style, z);
    for (const node of [this.cellInput, this.cellBackdrop]) {
      node.style.font = font;
      node.style.textAlign = style.ha === 'center' || style.ha === 'right' ? style.ha : 'left';
    }
    this.cellInput.style.textDecoration = style.u ? 'underline' : style.strike ? 'line-through' : '';
    this.box.style.background = style.fill || '';
    this.box.style.color = style.color || (style.fill ? '#000000' : '');
    this.wrap = !!style.wrap;
  }

  /** Puts the editor over its cell (called by every render). */
  position(m) {
    if (!this.active) return;
    const app = this.app;
    const visible = app.sheet === this.sheet;
    this.box.hidden = !visible || this.where === 'bar' && !this.cellInput.value;
    if (!visible) return;
    const grid = app.grid;
    const r = grid.cellRect(this.row, this.col, m || grid.metrics());
    const maxW = Math.max(r.w, grid.width - r.x - 4);
    this.box.style.left = Math.round(r.x) + 'px';
    this.box.style.top = Math.round(r.y) + 'px';
    this.box.style.minWidth = Math.round(r.w) + 'px';
    this.box.style.minHeight = Math.round(r.h) + 'px';
    this.box.style.maxWidth = Math.round(this.wrap ? r.w : maxW) + 'px';
    this.box.style.width = this.wrap ? Math.round(r.w) + 'px' : '';
    this.box.style.maxHeight = Math.max(r.h, grid.height - r.y - 4) + 'px';
  }

  onInput(input, initial = false) {
    const other = input === this.cellInput ? this.barInput : this.cellInput;
    if (other.value !== input.value) other.value = input.value;
    if (!initial) this.point = this.point && input.selectionStart === this.point.end ? this.point : null;
    const text = input.value;
    this.paintBackdrop(this.cellBackdrop, text);
    this.paintBackdrop(this.barBackdrop, text);
    this.highlight(text);
    this.afterCaret(input);
    this.position();
  }

  /** The text behind a textarea, with the references of a formula coloured. */
  paintBackdrop(node, text) {
    const parts = [];
    if (text.startsWith('=')) {
      const refs = referencesOf(text.slice(1), this.app.locale);
      const colors = new Map();
      let pos = 0;
      for (const ref of refs) {
        const s = ref.start + 1;
        const e = ref.end + 1;
        const k = text.slice(s, e).replace(/\$/g, '').toUpperCase();
        if (!colors.has(k)) colors.set(k, REF_COLORS[colors.size % REF_COLORS.length]);
        if (s > pos) parts.push(document.createTextNode(text.slice(pos, s)));
        parts.push(el('span', { style: { color: colors.get(k) }, text: text.slice(s, e) }));
        pos = e;
      }
      if (pos < text.length) parts.push(document.createTextNode(text.slice(pos)));
    } else {
      parts.push(document.createTextNode(text));
    }
    parts.push(document.createTextNode('​'));
    node.replaceChildren(...parts);
  }

  /** The references of the formula drawn on the grid in their colours. */
  highlight(text) {
    const grid = this.app.grid;
    grid.highlights = [];
    if (text.startsWith('=')) {
      const colors = new Map();
      for (const ref of referencesOf(text.slice(1), this.app.locale)) {
        const k = text.slice(ref.start + 1, ref.end + 1).replace(/\$/g, '').toUpperCase();
        if (!colors.has(k)) colors.set(k, REF_COLORS[colors.size % REF_COLORS.length]);
        const sheet = ref.sheet ? this.app.wb.sheetByName(ref.sheet) : this.sheet;
        if (!sheet) continue;
        grid.highlights.push({ sheet: sheet.id, area: { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }, color: colors.get(k) });
      }
    }
    grid.requestRender();
  }

  afterCaret(input) {
    if (!this.active) return;
    this.updateSuggest(input);
    this.updateTip(input);
  }

  onBlur() {
    if (!this.active) return;
    const a = document.activeElement;
    if (a === this.cellInput || a === this.barInput || this.popup.contains(a)) return;
    if (this.app.keepEditing) return;
    if (a && (a.closest?.('.mo-sheets-grid, .mo-sheets-tabs, dialog, .mo-pop') || a === document.body)) return;
    this.commit(0, 0);
  }

  close() {
    this.active = false;
    this.point = null;
    this.box.hidden = true;
    this.popup.hidden = true;
    this.tip.hidden = true;
    this.app.grid.highlights = [];
    this.app.setMode('ready');
    this.updateButtons();
    if (this.app.sheet !== this.sheet) this.app.showSheet(this.sheet, { keepEditor: true });
    this.app.grid.requestRender();
  }

  /** Writes the text into the cell and moves (dr, dc). Returns false when the text was refused (the editor stays open). */
  commit(dr = 1, dc = 0, options = {}) {
    if (!this.active) return true;
    const text = this.cellInput.value;
    const ok = this.app.commitInput(this.sheet, this.row, this.col, text, options);
    if (ok === false) { this.input.focus(); return false; }
    this.close();
    this.app.afterCommit(dr, dc);
    this.refresh();
    return true;
  }

  cancel() {
    if (!this.active) return;
    this.close();
    this.refresh();
    this.app.focusGrid();
  }

  // ---- keys --------------------------------------------------------------------------------------------------------------------

  onKey(e, input) {
    if (!this.active) {
      if (e.key === 'Escape') { this.refresh(); this.app.focusGrid(); }
      return;
    }
    if (!this.popup.hidden && this.suggestKey(e)) return;
    const mod = e.ctrlKey || e.metaKey;
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        if (e.altKey) { this.insertText(input, '\n'); return; }
        this.commit(mod ? 0 : e.shiftKey ? -1 : 1, 0, { all: mod });
        return;
      case 'Tab':
        e.preventDefault();
        this.commit(0, e.shiftKey ? -1 : 1);
        return;
      case 'Escape':
        e.preventDefault();
        this.cancel();
        return;
      case 'F2':
        e.preventDefault();
        this.mode = this.mode === 'enter' ? 'edit' : 'enter';
        this.app.setMode(this.mode);
        return;
      case 'F4':
        e.preventDefault();
        this.cycleDollar(input);
        return;
      case 'ArrowUp': case 'ArrowDown': case 'ArrowLeft': case 'ArrowRight': {
        if (this.mode !== 'enter' || e.altKey) return;
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
        if (this.canPoint(input)) {
          e.preventDefault();
          this.pointByKey(d[0], d[1], e.shiftKey, mod);
          return;
        }
        if (input === this.barInput) return;
        e.preventDefault();
        this.commit(d[0], d[1]);
        return;
      }
      default: break;
    }
  }

  insertText(input, text) {
    const s = input.selectionStart;
    const e = input.selectionEnd;
    input.setRangeText(text, s, e, 'end');
    this.onInput(input);
  }

  // ---- point mode ----------------------------------------------------------------------------------------------------------------

  /** Whether a reference can be inserted at the caret (a formula after an operator, "(" or a separator, or replacing the pointed one). */
  canPoint(input = this.input) {
    if (!this.active) return false;
    const text = input.value;
    if (!text.startsWith('=')) return false;
    const caret = input.selectionStart;
    if (this.point && caret === this.point.end && input.selectionEnd === caret) return true;
    if (input.selectionEnd !== caret) return false;
    const before = text.slice(0, caret).replace(/\s+$/, '');
    const after = text.slice(caret).trimStart();
    if (after && !/^[),;+\-*/^&<>=:]/.test(after)) return false;
    return POINT_AFTER.includes(before[before.length - 1] || '');
  }

  get pointing() { return !!this.point; }

  /** Inserts (or replaces) the pointed reference: an area of a sheet. */
  pointAt(sheet, area) {
    const input = this.input;
    const prefix = sheet !== this.sheet ? sheetPrefix(sheet.name) : '';
    const text = prefix + areaName(area);
    const start = this.point ? this.point.start : input.selectionStart;
    const end = this.point ? this.point.end : input.selectionEnd;
    input.setRangeText(text, start, end, 'end');
    this.point = { start, end: start + text.length, sheet, area, anchor: this.point?.anchor || { row: area.r1, col: area.c1 }, cursor: { row: area.r2, col: area.c2 } };
    const keep = this.point;
    this.onInput(input);
    this.point = keep;
    this.app.setMode('point');
  }

  /** Mouse down on a cell while pointing. */
  pointStart(sheet, row, col) {
    this.point = this.point && this.point.end === this.input.selectionStart ? this.point : null;
    if (this.point) this.point.anchor = { row, col };
    else this.point = null;
    const m = sheet.mergeAt(row, col);
    this.pointAt(sheet, m || { r1: row, c1: col, r2: row, c2: col });
    this.point.anchor = { row, col };
  }

  /** Mouse moved while pointing: the reference grows to a range. */
  pointExtend(sheet, row, col) {
    if (!this.point) return;
    const a = this.point.anchor;
    this.pointAt(sheet, { r1: Math.min(a.row, row), c1: Math.min(a.col, col), r2: Math.max(a.row, row), c2: Math.max(a.col, col) });
  }

  pointByKey(dr, dc, extend, jump) {
    const sheet = this.app.sheet;
    let cursor = this.point?.cursor || { row: this.row, col: this.col };
    if (!this.point && sheet !== this.sheet) cursor = { ...this.app.sel.active };
    const next = jump ? this.app.commands.edge(sheet, cursor.row, cursor.col, dr, dc) : {
      row: Math.max(0, Math.min(MAX_ROWS - 1, dr ? sheet.rows.step(cursor.row, dr) : cursor.row)),
      col: Math.max(0, Math.min(MAX_COLS - 1, dc ? sheet.cols.step(cursor.col, dc) : cursor.col)),
    };
    if (extend && this.point) {
      const a = this.point.anchor;
      this.pointAt(sheet, { r1: Math.min(a.row, next.row), c1: Math.min(a.col, next.col), r2: Math.max(a.row, next.row), c2: Math.max(a.col, next.col) });
      this.point.cursor = next;
    } else {
      this.pointAt(sheet, { r1: next.row, c1: next.col, r2: next.row, c2: next.col });
      this.point.anchor = { ...next };
      this.point.cursor = next;
    }
    this.app.grid.scrollIntoView(next.row, next.col);
  }

  /** F4: $A$1 → A$1 → $A1 → A1 for the reference at the caret. */
  cycleDollar(input) {
    const text = input.value;
    if (!text.startsWith('=')) return;
    const caret = input.selectionStart;
    const ref = referencesOf(text.slice(1), this.app.locale).find(r => caret >= r.start + 1 && caret <= r.end + 1);
    if (!ref) return;
    const s = ref.start + 1;
    const e = ref.end + 1;
    const original = text.slice(s, e);
    const bang = original.lastIndexOf('!');
    const prefix = bang >= 0 ? original.slice(0, bang + 1) : '';
    const body = original.slice(prefix.length);
    const next = body.split(':').map(part => {
      const m = /^(\$?)([A-Za-z]{0,3})(\$?)(\d*)$/.exec(part);
      if (!m) return part;
      const [, ac, letters, ar, digits] = m;
      if (!letters || !digits) return (ac || ar ? '' : '$') + letters + digits;
      const state = (ac ? 2 : 0) + (ar ? 1 : 0); // 0 A1, 3 $A$1, 1 A$1, 2 $A1
      const order = { 0: 3, 3: 1, 1: 2, 2: 0 };
      const n = order[state];
      return (n & 2 ? '$' : '') + letters + (n & 1 ? '$' : '') + digits;
    }).join(':');
    input.setRangeText(prefix + next, s, e, 'end');
    this.point = null;
    this.onInput(input);
  }

  // ---- suggestions and argument tips -----------------------------------------------------------------------------------------------

  updateSuggest(input) {
    const text = input.value;
    const caret = input.selectionStart;
    const before = text.slice(0, caret);
    const m = text.startsWith('=') && input.selectionEnd === caret ? /(?:^=|[=(,;+\-*/^&<> ])([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before) : null;
    if (!m || /^[A-Za-z]{1,3}\d+$/.test(m[1])) { this.hideSuggest(); return; }
    const prefix = m[1].toUpperCase();
    const fns = [...FUNCTIONS.keys()].filter(n => n.startsWith(prefix)).sort();
    const names = this.app.wb.names.map(n => n.name).filter(n => n.toUpperCase().startsWith(prefix));
    const items = [...names.map(n => ({ name: n, kind: 'name' })), ...fns.map(n => ({ name: n, kind: 'fn' }))].slice(0, 12);
    if (!items.length || (items.length === 1 && items[0].name.toUpperCase() === prefix && items[0].kind === 'name')) { this.hideSuggest(); return; }
    this.suggest = { items, index: 0, start: caret - m[1].length, input };
    this.popup.replaceChildren(...items.map((it, i) => el('div', {
      class: 'mo-sheets-suggest__item' + (i === 0 ? ' is-active' : ''), role: 'option', 'aria-selected': i === 0 ? 'true' : 'false',
      onmousedown: e => { e.preventDefault(); this.suggest.index = i; this.acceptSuggest(); },
    }, el('span', { class: 'mo-sheets-suggest__icon', text: it.kind === 'fn' ? 'fx' : '#' }), el('span', { text: it.name }))));
    this.popup.hidden = false;
    this.placeBelow(this.popup, input);
  }

  hideSuggest() { this.popup.hidden = true; this.suggest = null; }

  suggestKey(e) {
    const s = this.suggest;
    if (!s) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      s.index = (s.index + (e.key === 'ArrowDown' ? 1 : -1) + s.items.length) % s.items.length;
      [...this.popup.children].forEach((c, i) => { c.classList.toggle('is-active', i === s.index); c.setAttribute('aria-selected', i === s.index ? 'true' : 'false'); });
      return true;
    }
    if (e.key === 'Tab') { e.preventDefault(); this.acceptSuggest(); return true; }
    if (e.key === 'Escape') { e.preventDefault(); this.hideSuggest(); return true; }
    return false;
  }

  acceptSuggest() {
    const s = this.suggest;
    if (!s) return;
    const item = s.items[s.index];
    const input = s.input;
    const insert = item.kind === 'fn' ? item.name + '(' : item.name;
    const after = input.value.slice(input.selectionStart);
    const skip = item.kind === 'fn' && after.startsWith('(') ? 1 : 0;
    input.setRangeText(insert, s.start, input.selectionStart + skip, 'end');
    this.hideSuggest();
    this.onInput(input);
    input.focus();
  }

  updateTip(input) {
    const text = input.value;
    if (!text.startsWith('=')) { this.tip.hidden = true; return; }
    const caret = input.selectionStart;
    let tokens;
    try { tokens = tokenize(text.slice(1, caret), { locale: this.app.locale }); } catch { this.tip.hidden = true; return; }
    const stack = [];
    let pendingFn = null;
    for (const tok of tokens) {
      if (tok.type === 'func') pendingFn = tok.name;
      else if (tok.type === '(') { stack.push({ name: pendingFn, arg: 0 }); pendingFn = null; } else if (tok.type === ')') stack.pop();
      else if (tok.type === 'sep' && stack.length) stack[stack.length - 1].arg++;
      else if (tok.type !== 'ws') pendingFn = null;
    }
    const top = [...stack].reverse().find(x => x.name);
    const spec = top && FUNCTIONS.get(top.name);
    if (!spec) { this.tip.hidden = true; return; }
    const args = spec.args ? spec.args.split(/,\s*/) : [];
    const sep = this.app.locale === 'de' ? '; ' : ', ';
    const parts = [el('span', { class: 'mo-sheets-argtip__fn', text: spec.name + '(' })];
    args.forEach((a, i) => {
      if (i) parts.push(document.createTextNode(sep));
      const current = i === top.arg || (i === args.length - 1 && top.arg >= args.length && /…|\.\.\./.test(a)) || (a === '…' && top.arg >= i);
      parts.push(el(current ? 'b' : 'span', { text: a }));
    });
    parts.push(document.createTextNode(')'));
    this.tip.replaceChildren(...parts);
    this.tip.hidden = false;
    this.placeBelow(this.tip, input, this.popup.hidden ? 0 : this.popup.offsetHeight + 2);
  }

  placeBelow(node, input, extra = 0) {
    const host = this.app.root.getBoundingClientRect();
    const anchor = (input === this.cellInput ? this.box : this.bar.querySelector('.mo-sheets-fwrap')).getBoundingClientRect();
    node.style.left = Math.max(0, Math.round(anchor.left - host.left)) + 'px';
    node.style.top = Math.round(anchor.bottom - host.top + 2 + extra) + 'px';
  }
}

/** The text of an area for the name box while selecting: "3R x 2C" (rows, columns). */
export function sizeLabel(a) {
  const rows = a.r2 - a.r1 + 1;
  const cols = a.c2 - a.c1 + 1;
  if (rows === MAX_ROWS) return cols + 'C';
  if (cols === MAX_COLS) return rows + 'R';
  return `${rows}R x ${cols}C`;
}

export { colName };
