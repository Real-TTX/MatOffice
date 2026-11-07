/* MatOffice Base – the datasheet of Access: record selectors (▶ current, ✎ being edited, * new record), the new record with "(New)" for the
   AutoNumber, saving when the record is left (Esc undoes the field, Esc again the record), sort and filter from the header (a check list of the
   values, text/number/date filters), filter by selection, Find and Replace, column widths, hidden and frozen columns, the totals row and the
   navigation bar "Record: |◀ ◀ 3 of 50 ▶ ▶| ▶*" with a search box. Long tables scroll virtually (only the visible rows exist). */
import { t, el, icon, dialog } from '../editor/util.js';
import { problemText, showProblem, ask, contextMenu, drag, field, textInput, selectInput, checkInput, radioGroup, lang as pageLang, iconButton, openPopover, closePopover } from './ui.js';
import { Parser, compile, exprToText, sortCompare, isNull, compare, BaseError, toText } from './expr.js';
import { targetOf } from './criteria.js';
import { aggregate } from './engine.js';
import { DateVal, serialOfIso, todaySerial, parseDate, parseNumber, formatValue } from './values.js';
import { keyOf, lower } from './model.js';

const RH = 26;
const SEL_W = 26;
const ADD_W = 120;
const DEFAULT_W = { text: 150, memo: 220, number: 96, currency: 110, datetime: 130, yesno: 70, autonumber: 96, hyperlink: 190 };

export class Datasheet {
  /**
   * options: { recordset, settings (persisted object: sort, filter, filterOn, totals, totalsOn, hidden, widths, frozen, order), persist(),
   *            compact (a subform), onStatus(), onCurrent(row), addField(type) (the "Click to Add" column of a table), fieldMenu(col) }
   */
  constructor(app, options) {
    this.app = app;
    this.o = options;
    this.rs = options.recordset;
    this.s = options.settings || {};
    this.cur = { r: 0, c: 0 };
    this.pending = null;
    this.sel = null;
    this.colSel = null;
    this.editing = false;
    this.busy = false;
    this.rowEls = new Map();
    this.el = el('div', { class: 'mo-base-ds' + (options.compact ? ' mo-base-ds--compact' : ''), tabindex: '-1' });
    this.grid = el('div', { class: 'mo-base-ds__grid', role: 'grid', 'aria-label': this.rs.name || t('Base.View.Datasheet') });
    this.head = el('div', { class: 'mo-base-ds__head', role: 'row' });
    this.body = el('div', { class: 'mo-base-ds__body', role: 'rowgroup' });
    this.foot = el('div', { class: 'mo-base-ds__totals', role: 'row' });
    this.grid.append(this.head, this.body, this.foot);
    this.navBar = el('div', { class: 'mo-base-ds__nav' });
    this.el.append(this.grid, this.navBar);
    this.grid.addEventListener('scroll', () => this.renderRows());
    this.el.addEventListener('keydown', e => this.onKey(e));
    this.resizeObserver = new ResizeObserver(() => this.renderRows());
    this.resizeObserver.observe(this.grid);
    this.applyView();
    this.render();
  }

  get readOnly() { return this.app.readOnly || !this.rs.canEdit; }
  get canAdd() { return !this.app.readOnly && this.rs.canAdd; }
  get lang() { return pageLang(); }
  get rowCount() { return this.view.length + (this.canAdd ? 1 : 0); }
  get newIndex() { return this.canAdd ? this.view.length : -1; }
  get currentRow() { return this.cur.r < this.view.length ? this.view[this.cur.r] : null; }
  get currentCol() { return this.cols[this.cur.c] || null; }

  destroy() { this.resizeObserver.disconnect(); this.closeFind(); }

  persist() { this.o.persist?.(); }

  status() { this.o.onStatus?.(); }

  // ---- the view: columns, filter, sort ---------------------------------------------------------------------------------------

  columnsInOrder() {
    const hidden = new Set(this.s.hidden || []);
    let cols = this.rs.columns.filter(c => !hidden.has(c.skey));
    const order = this.s.order || [];
    if (order.length) {
      const pos = new Map(order.map((k, i) => [k, i]));
      cols = [...cols].sort((a, b) => (pos.get(a.skey) ?? 1e6 + this.rs.columns.indexOf(a)) - (pos.get(b.skey) ?? 1e6 + this.rs.columns.indexOf(b)));
    }
    return cols;
  }

  widthOf(c) { return Math.max(30, this.s.widths?.[c.skey] || DEFAULT_W[c.type] || 140); }

  /** Filter and sort the records of the recordset into this.view. */
  applyView() {
    this.cols = this.columnsInOrder();
    let rows = this.rs.rows;
    this.filterError = null;
    if (this.s.filterOn && this.s.filter) {
      try {
        const f = this.compileOnRows(this.s.filter);
        rows = rows.filter(r => { const v = f(r); return v !== null && v !== undefined && v !== false && v !== 0; });
      } catch (e) {
        if (!(e instanceof BaseError)) throw e;
        this.filterError = e;
      }
    }
    const sort = (this.s.sort || []).map(s => ({ col: this.rs.columns.find(c => c.skey === s.field), desc: s.desc })).filter(s => s.col);
    if (sort.length) {
      const lang = this.lang;
      const keyFor = (r, col) => (this.rs.mapOf(col) ? this.rs.display(col, this.rs.get(r, col)) : this.rs.get(r, col));
      const keyed = rows.map((r, i) => ({ r, i, k: sort.map(s => keyFor(r, s.col)) }));
      keyed.sort((a, b) => {
        for (let j = 0; j < sort.length; j++) {
          const c = sortCompare(a.k[j], b.k[j], lang);
          if (c) return sort[j].desc ? -c : c;
        }
        return a.i - b.i;
      });
      rows = keyed.map(x => x.r);
    }
    this.view = rows;
  }

  /** An expression over the columns (by name or caption) as a function of a row of the recordset. */
  compileOnRows(text) {
    const ast = Parser.expression(text);
    const byName = new Map();
    for (const c of this.rs.columns) { byName.set(lower(c.name), c); if (!byName.has(lower(c.caption))) byName.set(lower(c.caption), c); }
    const fn = compile(ast, {
      lang: this.lang,
      resolve: parts => { const c = byName.get(lower(parts[parts.length - 1])); return c ? env => this.rs.get(env, c) : null; },
    });
    return row => fn(row);
  }

  /** Reloads the records (refresh, another view changed the data). */
  reload() {
    const curId = this.currentRow?.id;
    this.rs.load();
    this.applyView();
    const i = this.view.findIndex(r => r.id === curId);
    if (i >= 0) this.cur.r = i;
    this.cur.r = Math.min(this.cur.r, Math.max(0, this.rowCount - 1));
    this.render();
  }

  // ---- rendering -----------------------------------------------------------------------------------------------------------------

  template() {
    const widths = this.cols.map(c => this.widthOf(c) + 'px');
    return [SEL_W + 'px', ...widths, ...(this.o.addField && !this.app.readOnly ? [ADD_W + 'px'] : [])].join(' ');
  }

  frozenLeft(i) {
    let x = SEL_W;
    for (let j = 0; j < i; j++) x += this.widthOf(this.cols[j]);
    return x;
  }

  render() {
    const tpl = this.template();
    const total = SEL_W + this.cols.reduce((s, c) => s + this.widthOf(c), 0) + (this.o.addField && !this.app.readOnly ? ADD_W : 0);
    this.grid.style.setProperty('--mo-ds-cols', tpl);
    this.grid.style.setProperty('--mo-ds-width', total + 'px');
    this.frozen = Math.min(this.s.frozen || 0, this.cols.length);
    this.renderHead();
    this.body.style.height = this.rowCount * RH + 'px';
    for (const elRow of this.rowEls.values()) elRow.remove();
    this.rowEls.clear();
    this.renderRows();
    this.renderTotals();
    this.renderNav();
    this.status();
  }

  stickyCell(cell, i) {
    if (i < this.frozen) {
      cell.classList.add('is-frozen');
      cell.style.left = this.frozenLeft(i) + 'px';
      if (i === this.frozen - 1) cell.classList.add('is-frozen-last');
    }
  }

  renderHead() {
    const cells = [el('div', { class: 'mo-base-ds__corner', role: 'columnheader', 'aria-label': t('Base.Ds.SelectAll'), onclick: () => this.selectAll() })];
    this.cols.forEach((c, i) => {
      const sorted = (this.s.sort || []).find(s => s.field === c.skey);
      const filtered = this.isFiltered(c);
      const cell = el('div', {
        class: 'mo-base-ds__hcell' + (this.colSel === i ? ' is-selected' : '') + (c.align === 'right' ? ' is-right' : ''), role: 'columnheader',
        'aria-sort': sorted ? (sorted.desc ? 'descending' : 'ascending') : null, title: c.field?.description || c.caption,
      },
      el('span', { class: 'mo-base-ds__hname', text: c.caption }),
      sorted ? icon(sorted.desc ? 'arrow-down' : 'arrow-up', 'mo-base-ds__hsort') : null,
      filtered ? icon('filter', 'mo-base-ds__hfilter') : null,
      el('button', { type: 'button', class: 'mo-base-ds__hmenu', 'aria-label': t('Base.Ds.ColumnMenu', c.caption), tabindex: '-1', onclick: e => { e.stopPropagation(); this.columnMenu(i, e.currentTarget); } }, icon('chevron-down')),
      el('span', { class: 'mo-base-ds__resize', onpointerdown: e => this.startResize(e, c), ondblclick: e => { e.stopPropagation(); this.bestFit(c); } }));
      cell.addEventListener('click', () => this.selectColumn(i));
      cell.addEventListener('dblclick', e => { if (!e.target.closest('.mo-base-ds__resize, .mo-base-ds__hmenu')) this.o.renameField?.(c); });
      cell.addEventListener('contextmenu', e => { e.preventDefault(); this.selectColumn(i); this.headerContext(i, { x: e.clientX, y: e.clientY }); });
      this.stickyCell(cell, i);
      cells.push(cell);
    });
    if (this.o.addField && !this.app.readOnly) {
      cells.push(el('button', { type: 'button', class: 'mo-base-ds__hcell mo-base-ds__add', role: 'columnheader', onclick: e => this.o.addField(e.currentTarget) },
        el('span', { text: t('Base.Ds.ClickToAdd') }), icon('chevron-down')));
    }
    this.head.replaceChildren(...cells);
  }

  visibleRange() {
    const top = this.grid.scrollTop;
    const h = this.grid.clientHeight || 400;
    const first = Math.max(0, Math.floor(top / RH) - 4);
    const last = Math.min(this.rowCount - 1, Math.ceil((top + h) / RH) + 4);
    return [first, last];
  }

  renderRows() {
    const [first, last] = this.visibleRange();
    for (const [i, rowEl] of this.rowEls) {
      if (i < first || i > last) {
        if (rowEl.contains(this.editor) && document.activeElement === this.editor) continue;
        rowEl.remove();
        this.rowEls.delete(i);
      }
    }
    for (let i = first; i <= last; i++) {
      if (!this.rowEls.has(i)) {
        const rowEl = this.renderRow(i);
        this.rowEls.set(i, rowEl);
        this.body.append(rowEl);
      }
    }
  }

  rerenderRow(i) {
    const old = this.rowEls.get(i);
    if (!old) return;
    const fresh = this.renderRow(i);
    old.replaceWith(fresh);
    this.rowEls.set(i, fresh);
  }

  /** The values a row shows: the record, with the pending changes of the record being edited (and the defaults of the new record). */
  valuesOf(i) {
    const isNew = i === this.newIndex;
    const row = isNew ? null : this.view[i];
    const editingThis = this.pending && (isNew ? this.pending.id === null : this.pending.id === row?.id);
    const out = {};
    for (const c of this.cols) {
      let v = row ? this.rs.get(row, c) : null;
      if (isNew && !editingThis) v = null;
      if (editingThis && Object.prototype.hasOwnProperty.call(this.pending.values, c.key)) v = this.rs.toEngine(c, this.pending.values[c.key]);
      out[c.key] = v;
    }
    return { out, isNew, editingThis, row };
  }

  renderRow(i) {
    const { out, isNew, editingThis } = this.valuesOf(i);
    const current = i === this.cur.r;
    const selected = this.sel && i >= Math.min(this.sel.from, this.sel.to) && i <= Math.max(this.sel.from, this.sel.to);
    const rowEl = el('div', {
      class: 'mo-base-ds__row' + (current ? ' is-current' : '') + (selected ? ' is-selected' : '') + (isNew ? ' is-new' : '') + (i % 2 ? ' is-alt' : ''),
      role: 'row', 'aria-rowindex': i + 2, 'aria-selected': selected ? 'true' : null,
    });
    rowEl.style.top = i * RH + 'px';
    const marker = editingThis ? 'edit' : current ? 'play' : isNew ? null : null;
    const selCell = el('div', { class: 'mo-base-ds__sel', role: 'rowheader', 'aria-label': isNew ? t('Base.Ds.NewRecord') : t('Base.Ds.RecordN', i + 1) },
      marker ? icon(marker === 'edit' ? 'edit' : 'base-current') : isNew ? el('span', { class: 'mo-base-ds__star', text: '*' }) : null);
    selCell.addEventListener('pointerdown', e => { e.preventDefault(); this.selectRecords(i, e.shiftKey); });
    selCell.addEventListener('contextmenu', e => { e.preventDefault(); if (!this.sel) this.selectRecords(i, false); this.recordContext({ x: e.clientX, y: e.clientY }); });
    rowEl.append(selCell);
    this.cols.forEach((c, ci) => {
      const v = out[c.key];
      const isCurrent = current && ci === this.cur.c && !this.sel;
      const cell = el('div', { class: 'mo-base-ds__cell' + (c.align === 'right' ? ' is-right' : c.align === 'center' ? ' is-center' : '') + (isCurrent ? ' is-current' : '') + (this.colSel === ci ? ' is-colsel' : ''), role: 'gridcell' });
      this.stickyCell(cell, ci);
      if (isCurrent && this.editor && !this.readOnlyCell(c, isNew) && c.type !== 'yesno') {
        cell.append(this.editor);
        if (this.pickerButton) cell.append(this.pickerButton);
      } else this.fillCell(cell, c, v, isNew && !editingThis);
      cell.addEventListener('pointerdown', e => this.onCellDown(e, i, ci));
      cell.addEventListener('contextmenu', e => { e.preventDefault(); this.cellContext(i, ci, { x: e.clientX, y: e.clientY }); });
      rowEl.append(cell);
    });
    if (this.o.addField && !this.app.readOnly) rowEl.append(el('div', { class: 'mo-base-ds__cell mo-base-ds__cell--add' }));
    return rowEl;
  }

  readOnlyCell(c, isNew) {
    if (this.app.readOnly) return true;
    if (isNew) return !this.canAdd || !c.editable;
    return this.readOnly || !c.editable;
  }

  fillCell(cell, c, v, newRow) {
    if (c.type === 'yesno') {
      const on = v === true || (typeof v === 'number' && v !== 0);
      cell.append(el('span', { class: 'mo-base-ds__check' + (on ? ' is-on' : ''), role: 'checkbox', 'aria-checked': on ? 'true' : 'false' }, on ? icon('check') : null));
      return;
    }
    if (newRow && c.type === 'autonumber') { cell.append(el('span', { class: 'mo-base-ds__auto', text: t('Base.Ds.New') })); return; }
    const text = this.rs.display(c, v);
    if (c.type === 'hyperlink' && text) {
      const href = safeHref(String(v));
      if (href) {
        cell.append(el('a', { class: 'mo-base-ds__link', href, target: '_blank', rel: 'noopener noreferrer', text: text.replace(/^mailto:/, ''), onpointerdown: e => e.stopPropagation() }));
        return;
      }
    }
    cell.textContent = text;
    if (c.type === 'memo') cell.title = text.length > 60 ? text.slice(0, 500) : '';
  }

  renderTotals() {
    if (!this.s.totalsOn) { this.foot.hidden = true; this.foot.replaceChildren(); return; }
    this.foot.hidden = false;
    const cells = [el('div', { class: 'mo-base-ds__sel mo-base-ds__tlabel', text: t('Base.Ds.Total') })];
    this.cols.forEach((c, i) => {
      const fn = this.s.totals?.[c.skey] || 'none';
      const cell = el('div', { class: 'mo-base-ds__cell mo-base-ds__tcell' + (c.align === 'right' ? ' is-right' : ''), role: 'gridcell', tabindex: '0', 'aria-label': t('Base.Ds.TotalOf', c.caption) });
      this.stickyCell(cell, i);
      if (fn !== 'none') cell.textContent = this.totalText(c, fn);
      cell.addEventListener('click', () => this.chooseTotal(c, cell));
      cell.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.chooseTotal(c, cell); } });
      cells.push(cell);
    });
    if (this.o.addField && !this.app.readOnly) cells.push(el('div', { class: 'mo-base-ds__cell' }));
    this.foot.replaceChildren(...cells);
  }

  totalText(c, fn) {
    const values = this.view.map(r => this.rs.get(r, c));
    const name = { sum: 'Sum', avg: 'Avg', count: 'Count', min: 'Min', max: 'Max', stdev: 'StDev', var: 'Var' }[fn];
    let v;
    try { v = aggregate(name, values, this.lang); } catch { return '#Error'; }
    if (v === null) return '';
    if (fn === 'count') return String(v);
    if ((fn === 'avg' || fn === 'stdev' || fn === 'var') && c.type !== 'currency') return formatValue(v, { format: 'Standard', lang: this.lang });
    return this.rs.display({ ...c, field: c.field && { ...c.field, lookup: null } }, v);
  }

  chooseTotal(c, cell) {
    const numeric = ['number', 'currency', 'autonumber'].includes(c.type);
    const dated = c.type === 'datetime';
    const list = [['none', t('Base.Total.None')]];
    if (numeric) list.push(['sum', t('Base.Total.Sum')], ['avg', t('Base.Total.Avg')]);
    list.push(['count', t('Base.Total.Count')]);
    if (numeric || dated) list.push(['max', t('Base.Total.Max')], ['min', t('Base.Total.Min')]);
    if (numeric) list.push(['stdev', t('Base.Total.StDev')], ['var', t('Base.Total.Var')]);
    contextMenu(cell, list.map(([v, label]) => ({ label, checked: (this.s.totals?.[c.skey] || 'none') === v, run: () => { this.s.totals = { ...(this.s.totals || {}), [c.skey]: v }; this.persist(); this.renderTotals(); } })));
  }

  // ---- navigation bar ----------------------------------------------------------------------------------------------------------

  renderNav() {
    const total = this.view.length;
    const pos = Math.min(this.cur.r, this.rowCount - 1) + 1;
    const box = el('input', { class: 'mo-base-ds__pos', value: this.cur.r === this.newIndex ? '' : String(pos), 'aria-label': t('Base.Nav.Current'), inputmode: 'numeric' });
    box.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); const n = Number(box.value); if (n >= 1 && n <= total) this.moveTo(n - 1, this.cur.c, { focus: true }); else box.value = String(pos); }
    });
    const navBtn = (name, label, run, disabled) => iconButton(name, label, run, { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: disabled ? true : null });
    const filtered = this.s.filterOn && this.s.filter;
    const hasFilter = !!this.s.filter;
    const search = el('input', { class: 'mo-base-ds__search', type: 'search', placeholder: t('Base.Nav.Search'), 'aria-label': t('Base.Nav.Search') });
    search.addEventListener('input', () => this.searchFor(search.value));
    search.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); this.searchFor(search.value, true); } });
    this.navBar.replaceChildren(
      el('span', { class: 'mo-base-ds__navlabel', text: t('Base.Nav.Record') }),
      navBtn('base-first', t('Base.Nav.First'), () => this.moveTo(0, this.cur.c, { focus: true }), this.cur.r <= 0),
      navBtn('chevron-left', t('Base.Nav.Previous'), () => this.moveTo(this.cur.r - 1, this.cur.c, { focus: true }), this.cur.r <= 0),
      box,
      el('span', { class: 'mo-base-ds__navof', text: t('Base.Nav.Of', total) + (filtered ? ' ' + t('Base.Nav.FilteredShort') : '') }),
      navBtn('chevron-right', t('Base.Nav.Next'), () => this.moveTo(this.cur.r + 1, this.cur.c, { focus: true }), this.cur.r >= this.rowCount - 1),
      navBtn('base-last', t('Base.Nav.Last'), () => this.moveTo(total - 1, this.cur.c, { focus: true }), this.cur.r >= total - 1),
      this.canAdd ? navBtn('base-new-record', t('Base.Nav.New'), () => this.moveTo(this.newIndex, 0, { focus: true }), false) : null,
      el('button', {
        type: 'button', class: 'mo-base-ds__filterflag' + (filtered ? ' is-on' : ''), disabled: hasFilter ? null : true,
        onclick: () => this.toggleFilter(),
      }, icon(filtered ? 'filter' : 'base-filter-off'), el('span', { text: filtered ? t('Base.Nav.Filtered') : hasFilter ? t('Base.Nav.Unfiltered') : t('Base.Nav.NoFilter') })),
      search,
    );
  }

  // ---- current cell and editing ------------------------------------------------------------------------------------------------

  onCellDown(e, i, ci) {
    if (e.button !== 0 || this.busy) return;
    if (e.target.closest('.mo-base-ds__link')) return;
    const col = this.cols[ci];
    if (e.target === this.editor || this.editor?.contains?.(e.target)) return;
    e.preventDefault();
    this.moveTo(i, ci, { focus: true, caret: true }).then(ok => {
      if (ok && col?.type === 'yesno' && !this.readOnlyCell(col, i === this.newIndex)) this.toggleCheck();
    });
  }

  /** Moves the current cell; commits the field and (when the row changes) the record first. Answers false when it has to stay. */
  async moveTo(r, c, { focus = false, caret = false } = {}) {
    if (this.busy) return false;
    r = Math.max(0, Math.min(r, this.rowCount - 1));
    c = Math.max(0, Math.min(c, this.cols.length - 1));
    if (this.rowCount === 0 || !this.cols.length) return false;
    this.busy = true;
    try {
      if (!(await this.commitField())) return false;
      if (r !== this.cur.r && !(await this.commitRecord())) return false;
      const oldR = this.cur.r;
      this.sel = null;
      this.colSel = null;
      this.cur = { r, c };
      this.editing = caret;
      this.makeEditor();
      this.rerenderRow(oldR);
      this.rerenderRow(r);
      this.scrollIntoView(r, c);
      this.renderRows();
      if (oldR !== r) { this.renderNav(); this.o.onCurrent?.(this.currentRow); }
      this.status();
      if (focus) this.focusEditor(caret);
      return true;
    } finally {
      this.busy = false;
    }
  }

  scrollIntoView(r, c) {
    const top = r * RH;
    const head = this.head.offsetHeight || RH;
    const foot = this.s.totalsOn ? this.foot.offsetHeight : 0;
    if (top < this.grid.scrollTop) this.grid.scrollTop = top;
    else if (top + RH > this.grid.scrollTop + this.grid.clientHeight - head - foot) this.grid.scrollTop = top + RH - this.grid.clientHeight + head + foot;
    if (c >= this.frozen) {
      const left = this.frozenLeft(c);
      const w = this.widthOf(this.cols[c]);
      const frozenW = this.frozenLeft(this.frozen);
      if (left - frozenW < this.grid.scrollLeft) this.grid.scrollLeft = left - frozenW;
      else if (left + w > this.grid.scrollLeft + this.grid.clientWidth) this.grid.scrollLeft = left + w - this.grid.clientWidth;
    }
  }

  /** The editor of the current cell: a text box, a list (lookups) or nothing (Yes/No, read-only). */
  makeEditor() {
    this.editor = null;
    this.pickerButton = null;
    const c = this.currentCol;
    if (!c) return;
    const isNew = this.cur.r === this.newIndex;
    const { out } = this.valuesOf(this.cur.r);
    const v = out[c.key];
    if (this.readOnlyCell(c, isNew)) {
      // a read-only cell still takes the focus (copy, find, keys)
      this.editor = el('input', { class: 'mo-base-ds__editor is-readonly', readonly: true, value: c.type === 'yesno' ? '' : (isNew && c.type === 'autonumber' ? t('Base.Ds.New') : this.rs.display(c, v)), 'aria-label': c.caption });
      if (c.type === 'yesno') { this.editor = null; this.focusSink(); return; }
    } else if (c.type === 'yesno') {
      this.editor = null;
      this.focusSink();
      return;
    } else if (c.field?.lookup) {
      const options = this.rs.options(c);
      const lk = c.field.lookup;
      if (lk.kind === 'values' && !lk.limit) {
        const listId = 'mo-base-dl-' + c.key;
        this.editor = el('input', { class: 'mo-base-ds__editor', list: listId, value: this.rs.editText(c, v), 'aria-label': c.caption, autocomplete: 'off' });
        this.datalist = el('datalist', { id: listId }, ...options.map(o => el('option', { value: o.label })));
        this.el.append(this.datalist);
      } else {
        const s = el('select', { class: 'mo-base-ds__editor mo-base-ds__select', 'aria-label': c.caption },
          el('option', { value: '', text: '' }), ...options.map((o, i) => el('option', { value: String(i), text: o.label })));
        const i = options.findIndex(o => keyOf(o.value) === keyOf(v));
        s.value = i >= 0 ? String(i) : '';
        s.lookupOptions = options;
        s.addEventListener('change', () => this.markDirty());
        this.editor = s;
      }
    } else {
      this.editor = el('input', { class: 'mo-base-ds__editor' + (c.align === 'right' ? ' is-right' : ''), value: this.rs.editText(c, v), 'aria-label': c.caption, autocomplete: 'off', spellcheck: c.type === 'text' || c.type === 'memo' ? 'true' : 'false' });
      if (c.type === 'datetime') {
        const picker = el('input', { type: 'date', class: 'mo-base-ds__datepick', tabindex: '-1', 'aria-hidden': 'true' });
        picker.addEventListener('change', () => {
          if (!picker.value) return;
          const serial = serialOfIso(picker.value);
          this.editor.value = formatValue(new DateVal(serial), { type: 'datetime', format: 'Short Date', lang: this.lang });
          this.markDirty();
          this.focusEditor(false);
        });
        this.pickerButton = el('button', { type: 'button', class: 'mo-base-ds__picker', tabindex: '-1', 'aria-label': t('Base.Ds.PickDate'), onpointerdown: e => { e.preventDefault(); e.stopPropagation(); } , onclick: () => { try { picker.showPicker(); } catch { picker.click(); } } }, icon('calendar'), picker);
      }
    }
    if (!this.editor) return;
    this.editorStart = this.editorValue();
    if (this.editor.tagName === 'INPUT' && !this.editor.readOnly) {
      this.editor.addEventListener('input', () => this.markDirty());
      this.editor.addEventListener('paste', e => this.onPaste(e));
    }
    this.editor.addEventListener('blur', () => { if (!this.busy && document.activeElement !== this.editor) this.commitField({ quiet: true }); });
  }

  /** Yes/No cells and read-only Yes/No take the keys on the datasheet itself. */
  focusSink() {
    this.el.setAttribute('tabindex', '-1');
  }

  focusEditor(caret) {
    if (this.editor) {
      this.editor.focus({ preventScroll: true });
      if (this.editor.tagName === 'INPUT' && typeof this.editor.select === 'function') {
        if (caret) { const n = this.editor.value.length; try { this.editor.setSelectionRange(n, n); } catch { /* not a text input */ } }
        else this.editor.select();
      }
    } else this.el.focus({ preventScroll: true });
  }

  editorValue() {
    if (!this.editor) return null;
    if (this.editor.tagName === 'SELECT') return this.editor.value;
    return this.editor.value;
  }

  markDirty() {
    if (this.readOnlyCell(this.currentCol, this.cur.r === this.newIndex)) return;
    this.editing = true;
    this.startPending();
  }

  startPending() {
    if (this.pending) return;
    const isNew = this.cur.r === this.newIndex;
    this.pending = { id: isNew ? null : this.currentRow.id, values: isNew ? { ...this.rs.defaults() } : {} };
    const sel = this.rowEls.get(this.cur.r)?.querySelector('.mo-base-ds__sel');
    sel?.replaceChildren(icon('edit'));
    this.status();
  }

  /** Takes the value of the editor into the pending record (checks the type and the validation rule of the field). */
  async commitField({ quiet = false } = {}) {
    const c = this.currentCol;
    if (!this.editor || !c || this.editor.readOnly || this.editorValue() === this.editorStart) return true;
    let parsed;
    if (this.editor.tagName === 'SELECT') {
      const i = this.editor.value;
      parsed = { ok: true, value: i === '' ? null : this.editor.lookupOptions[Number(i)].value };
    } else {
      let text = this.editor.value;
      if (c.field?.lookup?.kind === 'values') text = text.trim();
      parsed = this.rs.parse(c, text);
    }
    if (!parsed.ok) {
      if (quiet) return false;
      await showProblem({ code: parsed.code === 'TooLong' ? 'TooLong' : parsed.code === 'OutOfRange' ? 'OutOfRange' : 'TypeMismatch', args: [c.caption] });
      this.focusEditor(false);
      return false;
    }
    const rule = this.rs.checkRule(c, parsed.value);
    if (rule) {
      if (quiet) return false;
      await showProblem(rule);
      this.focusEditor(false);
      return false;
    }
    this.startPending();
    this.pending.values[c.key] = parsed.value;
    this.editorStart = this.editorValue();
    return true;
  }

  /** Saves the pending record (when the record is left, Shift+Enter, Save). Answers false (and stays) when it cannot be saved. */
  async commitRecord() {
    if (!this.pending) return true;
    const p = this.pending;
    const r = this.rs.save(p.id, p.values);
    if (!r.ok) {
      await showProblem(r.error);
      const ci = this.cols.findIndex(c => c.key === r.error.field || c.field?.id === r.error.field);
      if (ci >= 0 && ci !== this.cur.c) { this.cur.c = ci; this.makeEditor(); this.rerenderRow(this.cur.r); }
      this.focusEditor(false);
      return false;
    }
    this.pending = null;
    if (p.id === null) {
      // the new record stays where it was typed (at the end), a new "new record" row follows
      this.view.push(r.row);
      this.body.style.height = this.rowCount * RH + 'px';
      this.rerenderRow(this.cur.r);
      this.rowEls.get(this.cur.r + 1)?.remove();
      this.rowEls.delete(this.cur.r + 1);
      this.renderRows();
    } else {
      this.rerenderRow(this.cur.r);
    }
    this.renderTotals();
    this.renderNav();
    this.status();
    return true;
  }

  /** Esc: the field first, then the whole record (like Access). */
  undo() {
    if (this.editor && this.editorValue() !== this.editorStart && !this.editor.readOnly) {
      if (this.editor.tagName === 'SELECT') this.editor.value = this.editorStart; else this.editor.value = this.editorStart;
      if (!this.pending || !Object.keys(this.pending.values).length && this.pending.id !== null) { this.pending = null; this.rerenderRow(this.cur.r); this.focusEditor(false); }
      this.focusEditor(false);
      return true;
    }
    if (this.pending) {
      this.pending = null;
      this.makeEditor();
      this.rerenderRow(this.cur.r);
      this.focusEditor(false);
      this.status();
      return true;
    }
    return false;
  }

  get dirty() { return !!this.pending || (this.editor && !this.editor.readOnly && this.editorValue() !== this.editorStart); }

  /** Commits what is being edited (before another view, saving, closing). */
  async commit() {
    if (!(await this.commitField())) return false;
    return this.commitRecord();
  }

  toggleCheck() {
    const c = this.currentCol;
    if (!c || c.type !== 'yesno' || this.readOnlyCell(c, this.cur.r === this.newIndex)) return;
    const { out } = this.valuesOf(this.cur.r);
    this.startPending();
    this.pending.values[c.key] = !(out[c.key] === true);
    this.rerenderRow(this.cur.r);
  }

  // ---- keys ----------------------------------------------------------------------------------------------------------------------

  onKey(e) {
    if (this.busy) return;
    const ed = this.editor;
    const inEditor = e.target === ed;
    const caretMode = inEditor && this.editing && ed.tagName === 'INPUT' && !ed.readOnly;
    const mod = e.ctrlKey || e.metaKey;
    const go = (r, c) => { e.preventDefault(); this.moveTo(r, c, { focus: true }); };
    if (e.target.closest?.('.mo-base-ds__nav, .mo-base-find')) return;
    switch (e.key) {
      case 'ArrowDown':
        if (e.altKey && ed?.tagName === 'SELECT') { try { ed.showPicker(); } catch { /* old browsers */ } e.preventDefault(); return; }
        if (mod) return go(this.rowCount - 1, this.cur.c);
        return go(this.cur.r + 1, this.cur.c);
      case 'ArrowUp': if (mod) return go(0, this.cur.c); return go(this.cur.r - 1, this.cur.c);
      case 'PageDown': return go(this.cur.r + Math.max(1, Math.floor(this.grid.clientHeight / RH) - 2), this.cur.c);
      case 'PageUp': return go(this.cur.r - Math.max(1, Math.floor(this.grid.clientHeight / RH) - 2), this.cur.c);
      case 'ArrowRight':
        if (caretMode) return;
        if (ed?.tagName === 'SELECT' || !caretMode) return go(this.cur.r, this.cur.c + 1);
        return;
      case 'ArrowLeft':
        if (caretMode) return;
        return go(this.cur.r, this.cur.c - 1);
      case 'Home': if (caretMode && !mod) return; return go(mod ? 0 : this.cur.r, 0);
      case 'End': if (caretMode && !mod) return; return go(mod ? this.rowCount - 1 : this.cur.r, this.cols.length - 1);
      case 'Tab': case 'Enter': {
        if (e.key === 'Enter' && e.shiftKey) { e.preventDefault(); this.commit().then(() => this.focusEditor(false)); return; }
        if (e.key === 'Enter' && e.altKey) return;
        const back = e.shiftKey;
        let r = this.cur.r;
        let c = this.cur.c + (back ? -1 : 1);
        if (c >= this.cols.length) { c = 0; r++; }
        if (c < 0) { c = this.cols.length - 1; r--; }
        if (r < 0 || r >= this.rowCount) { if (e.key === 'Tab' && this.o.compact) return; e.preventDefault(); return; }
        return go(r, c);
      }
      case 'F2':
        e.preventDefault();
        this.editing = !this.editing;
        this.focusEditor(this.editing);
        return;
      case 'Escape':
        if (this.undo()) { e.preventDefault(); e.stopPropagation(); }
        return;
      case ' ':
        if (this.currentCol?.type === 'yesno' && !inEditor) { e.preventDefault(); this.toggleCheck(); }
        return;
      case 'Delete':
        if (this.sel) { e.preventDefault(); this.deleteRecords(); }
        return;
      default: break;
    }
    if (mod && (e.key === 'c' || e.key === 'C') && (this.sel || !caretMode)) { if (this.sel || (ed && ed.selectionStart === ed.selectionEnd)) { e.preventDefault(); this.copy(); } return; }
    if (mod && (e.key === 'a' || e.key === 'A') && !caretMode) { e.preventDefault(); this.selectAll(); return; }
    if (mod && e.key === "'" ) { e.preventDefault(); this.copyFromAbove(); return; }
    if (mod && e.key === ';') { e.preventDefault(); this.insertToday(); return; }
    if (e.shiftKey && e.key === 'F2') { e.preventDefault(); this.zoom(); return; }
    if (mod && e.key === '+') { e.preventDefault(); this.moveTo(this.newIndex, 0, { focus: true }); return; }
    if (mod && e.key === '-') { e.preventDefault(); this.deleteRecords(); return; }
    if (inEditor && e.key.length === 1 && !mod && !this.editing && ed.tagName === 'INPUT' && !ed.readOnly) {
      // typing in navigation mode replaces the value (the whole value is selected) and switches to editing
      this.editing = true;
    }
  }

  onPaste(e) {
    const text = e.clipboardData?.getData('text/plain') || '';
    if (!/[\t\n]/.test(text.replace(/\r?\n$/, ''))) return;
    if (this.cur.r !== this.newIndex) return;
    e.preventDefault();
    this.pasteAppend(text);
  }

  // ---- records: select, delete, copy, paste ---------------------------------------------------------------------------------------

  async selectRecords(i, extend) {
    if (i === this.newIndex) { this.moveTo(i, 0, { focus: true }); return; }
    if (!(await this.commitField()) || (i !== this.cur.r && !(await this.commitRecord()))) return;
    const oldR = this.cur.r;
    this.sel = extend && this.sel ? { from: this.sel.from, to: i } : { from: i, to: i };
    this.cur.r = i;
    this.colSel = null;
    this.editor = null;
    for (const k of [...this.rowEls.keys()]) this.rerenderRow(k);
    if (oldR !== i) { this.renderNav(); this.o.onCurrent?.(this.currentRow); }
    this.el.focus({ preventScroll: true });
    this.status();
  }

  async selectAll() {
    if (!this.view.length) return;
    if (!(await this.commit())) return;
    this.sel = { from: 0, to: this.view.length - 1 };
    this.editor = null;
    for (const k of [...this.rowEls.keys()]) this.rerenderRow(k);
    this.el.focus({ preventScroll: true });
    this.status();
  }

  selectColumn(i) {
    this.colSel = i;
    this.cur.c = i;
    this.renderHead();
    for (const k of [...this.rowEls.keys()]) this.rerenderRow(k);
  }

  selectedRows() {
    if (this.sel) {
      const a = Math.min(this.sel.from, this.sel.to);
      const b = Math.max(this.sel.from, this.sel.to);
      return this.view.slice(a, b + 1);
    }
    return this.currentRow ? [this.currentRow] : [];
  }

  async deleteRecords() {
    if (this.readOnly || !this.rs.canDelete) return;
    if (this.cur.r === this.newIndex && !this.sel) { this.undo(); return; }
    const rows = this.selectedRows();
    if (!rows.length) return;
    if (this.pending) { this.pending = null; }
    const plan = this.rs.planRemove(rows.map(r => r.id));
    if (!plan.ok) { await showProblem(plan.error); return; }
    const message = plan.cascaded
      ? t('Base.Ds.DeleteCascade', rows.length, plan.cascaded)
      : t('Base.Ds.DeleteAsk', rows.length);
    if (!(await ask(message, { title: t('Base.Ds.DeleteTitle'), ok: t('Base.Yes'), cancel: t('Base.No'), danger: true }))) return;
    this.rs.remove(plan);
    this.sel = null;
    this.applyView();
    this.cur.r = Math.min(this.cur.r, Math.max(0, this.rowCount - 1));
    this.makeEditor();
    this.render();
    this.focusEditor(false);
    this.app.toast(t('Base.Ds.Deleted', rows.length + plan.cascaded));
  }

  rowText(row) { return this.cols.map(c => this.rs.display(c, this.rs.get(row, c)).replace(/[\t\n]/g, ' ')).join('\t'); }

  async copy() {
    let text;
    if (this.sel) text = [this.cols.map(c => c.caption).join('\t'), ...this.selectedRows().map(r => this.rowText(r))].join('\n');
    else {
      const c = this.currentCol;
      const { out } = this.valuesOf(this.cur.r);
      text = c ? this.rs.display(c, out[c.key]) : '';
    }
    try { await navigator.clipboard.writeText(text); this.app.toast(t('Base.Copied')); } catch { /* the browser refused */ }
  }

  async cut() {
    await this.copy();
    if (this.sel) { await this.deleteRecords(); return; }
    if (this.editor && !this.editor.readOnly && this.editor.tagName === 'INPUT') { this.editor.value = ''; this.markDirty(); }
  }

  async paste() {
    let text = '';
    try { text = await navigator.clipboard.readText(); } catch { this.app.toast(t('Base.PasteKeys'), 'info'); return; }
    if (/[\t\n]/.test(text.replace(/\r?\n$/, '')) && (this.cur.r === this.newIndex || this.sel)) { await this.pasteAppend(text); return; }
    if (this.editor && !this.editor.readOnly && this.editor.tagName === 'INPUT') {
      this.editor.value = text.replace(/\r?\n$/, '');
      this.markDirty();
      this.focusEditor(true);
    }
  }

  /** Paste Append: the lines of the clipboard become new records (columns in the order of the datasheet; a header line is skipped). */
  async pasteAppend(text) {
    if (!this.canAdd) return;
    const lines = text.replace(/\r/g, '').split('\n').filter(l => l.length);
    if (lines.length && lines[0].split('\t').every((h, i) => lower(h) === lower(this.cols[i]?.caption || ''))) lines.shift();
    let done = 0;
    let failed = 0;
    for (const line of lines) {
      const cells = line.split('\t');
      const values = { ...this.rs.defaults() };
      let bad = false;
      this.cols.forEach((c, i) => {
        if (bad || i >= cells.length || !c.editable) return;
        if (c.field?.lookup?.kind === 'table') {
          const hit = this.rs.options(c).find(o => lower(o.label) === lower(cells[i]));
          if (hit) { values[c.key] = hit.value; return; }
        }
        const p = this.rs.parse(c, cells[i]);
        if (!p.ok) bad = true; else values[c.key] = p.value;
      });
      if (bad) { failed++; continue; }
      const r = this.rs.save(null, values);
      if (r.ok) { done++; this.view.push(r.row); } else failed++;
    }
    this.cur.r = this.newIndex;
    this.render();
    this.focusEditor(false);
    this.app.toast(t('Base.Ds.Pasted', done, failed), failed ? 'warning' : 'success');
  }

  copyFromAbove() {
    const c = this.currentCol;
    if (!c || this.cur.r === 0 || this.readOnlyCell(c, this.cur.r === this.newIndex)) return;
    const above = this.valuesOf(this.cur.r - 1).out[c.key];
    if (this.editor && this.editor.tagName === 'INPUT') { this.editor.value = this.rs.editText(c, above); this.markDirty(); }
    else if (this.editor?.tagName === 'SELECT') { const i = this.editor.lookupOptions.findIndex(o => keyOf(o.value) === keyOf(above)); this.editor.value = i >= 0 ? String(i) : ''; this.markDirty(); }
  }

  insertToday() {
    const c = this.currentCol;
    if (!c || c.type !== 'datetime' || !this.editor || this.editor.readOnly) return;
    this.editor.value = formatValue(new DateVal(todaySerial()), { type: 'datetime', format: 'Short Date', lang: this.lang });
    this.markDirty();
  }

  /** Shift+F2: the Zoom box for long texts. */
  async zoom() {
    const c = this.currentCol;
    if (!c || !this.editor || this.editor.tagName !== 'INPUT') return;
    const area = el('textarea', { class: 'mat-input mat-textarea mo-base-zoom', rows: 12, readonly: this.editor.readOnly ? true : null });
    area.value = this.editor.value;
    const r = await dialog({ title: t('Base.Zoom'), icon: 'search', size: 'md', body: area, onOpen: () => area.focus(), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'ok' && !this.editor.readOnly) { this.editor.value = area.value; this.markDirty(); }
    this.focusEditor(true);
  }

  // ---- sort and filter -----------------------------------------------------------------------------------------------------------

  async sortBy(col, desc) {
    if (!col || !(await this.commit())) return;
    this.s.sort = [{ field: col.skey, desc }];
    this.persist();
    this.refreshView();
  }

  async clearSort() {
    if (!(await this.commit())) return;
    this.s.sort = [];
    this.persist();
    this.refreshView();
  }

  refreshView() {
    const id = this.currentRow?.id;
    this.applyView();
    const i = this.view.findIndex(r => r.id === id);
    this.cur.r = i >= 0 ? i : Math.min(this.cur.r, Math.max(0, this.rowCount - 1));
    this.makeEditor();
    this.render();
    if (this.filterError) showProblem(this.filterError);
  }

  isFiltered(col) {
    if (!this.s.filterOn || !this.s.filter) return false;
    return this.filterTerms().some(term => this.termColumn(term) === col);
  }

  filterTerms() {
    if (!this.s.filter) return [];
    let ast;
    try { ast = Parser.expression(this.s.filter); } catch { return []; }
    const out = [];
    const go = x => { if (x.t === 'bin' && x.op === 'AND') { go(x.l); go(x.r); } else out.push(x); };
    if (ast) go(ast);
    return out;
  }

  termColumn(term) {
    const target = targetOf(term);
    if (!target || target.t !== 'id') return null;
    const name = lower(target.parts[target.parts.length - 1]);
    return this.rs.columns.find(c => lower(c.name) === name) || null;
  }

  /** Adds a condition about a column to the filter (the other conditions about it go) and switches the filter on. */
  async setColumnFilter(col, condition) {
    if (!(await this.commit())) return;
    const keep = this.filterTerms().filter(term => this.termColumn(term) !== col).map(term => exprToText(term));
    if (condition) keep.push(condition);
    this.s.filter = keep.map(x => (keep.length > 1 ? '(' + x + ')' : x)).join(' And ');
    this.s.filterOn = !!this.s.filter;
    this.persist();
    this.refreshView();
  }

  async toggleFilter() {
    if (!this.s.filter || !(await this.commit())) return;
    this.s.filterOn = !this.s.filterOn;
    this.persist();
    this.refreshView();
  }

  async clearFilters() {
    if (!(await this.commit())) return;
    this.s.filter = '';
    this.s.filterOn = false;
    this.persist();
    this.refreshView();
  }

  colRef(col) { return '[' + col.name + ']'; }

  literal(col, v) {
    if (v === null || v === undefined) return 'Null';
    if (v instanceof DateVal) return exprToText({ t: 'date', v: v.s });
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (typeof v === 'number') return exprToText({ t: 'num', v });
    return exprToText({ t: 'str', v: String(v) });
  }

  /** Filter by Selection: equals / does not equal / contains / does not contain the value of the current cell. */
  filterBySelection(kind) {
    const c = this.currentCol;
    if (!c) return;
    const { out } = this.valuesOf(this.cur.r);
    const v = out[c.key];
    const ref = this.colRef(c);
    let cond;
    if (v === null || v === undefined) cond = kind === 'notEqual' ? ref + ' Is Not Null' : ref + ' Is Null';
    else if (kind === 'equal') cond = ref + '=' + this.literal(c, v);
    else if (kind === 'notEqual') cond = '(' + ref + '<>' + this.literal(c, v) + ' Or ' + ref + ' Is Null)';
    else {
      const text = this.rs.display(c, v);
      cond = ref + (kind === 'contains' ? ' Like ' : ' Not Like ') + exprToText({ t: 'str', v: '*' + text.replace(/[*?#[]/g, m => '[' + m + ']') + '*' });
    }
    this.setColumnFilter(c, cond);
  }

  selectionLabel(kind) {
    const c = this.currentCol;
    if (!c) return '';
    const { out } = this.valuesOf(this.cur.r);
    const text = this.rs.display(c, out[c.key]) || t('Base.Filter.Blanks');
    const short = text.length > 24 ? text.slice(0, 24) + '…' : text;
    return { equal: t('Base.Filter.Equals', short), notEqual: t('Base.Filter.NotEquals', short), contains: t('Base.Filter.Contains', short), notContains: t('Base.Filter.NotContains', short) }[kind];
  }

  /** The menu of a column header: sort, clear the filter, the type's filters and the check list of values. */
  columnMenu(i, anchor) {
    const c = this.cols[i];
    if (!c) return;
    this.selectColumn(i);
    const numeric = ['number', 'currency', 'autonumber'].includes(c.type);
    const dated = c.type === 'datetime';
    const asc = numeric ? t('Base.Sort.SmallLarge') : dated ? t('Base.Sort.OldNew') : t('Base.Sort.AZ');
    const desc = numeric ? t('Base.Sort.LargeSmall') : dated ? t('Base.Sort.NewOld') : t('Base.Sort.ZA');
    openPopover(anchor, close => {
      const box = el('div', { class: 'mo-base-colmenu' });
      const item = (ic, label, run, disabled) => el('button', { type: 'button', class: 'mo-menu__item', disabled: disabled ? true : null, onclick: () => { close(); run(); } }, icon(ic), el('span', { class: 'mo-menu__label', text: label }));
      box.append(
        item('base-sort-asc', asc, () => this.sortBy(c, false)),
        item('base-sort-desc', desc, () => this.sortBy(c, true)),
        el('div', { class: 'mo-menu__sep' }),
        item('base-filter-off', t('Base.Filter.ClearFrom', c.caption), () => this.setColumnFilter(c, null), !this.isFiltered(c)),
        item('filter', numeric ? t('Base.Filter.NumberFilters') : dated ? t('Base.Filter.DateFilters') : t('Base.Filter.TextFilters'), () => this.customFilter(c)),
        el('div', { class: 'mo-menu__sep' }),
      );
      // the check list of the values (at most 1000 different ones)
      const seen = new Map();
      let blanks = false;
      for (const r of this.rs.rows) {
        const v = this.rs.get(r, c);
        if (v === null || v === undefined || v === '') { blanks = true; continue; }
        const k = keyOf(v);
        if (!seen.has(k)) seen.set(k, v);
        if (seen.size > 1000) break;
      }
      const values = [...seen.values()].sort((a, b) => sortCompare(a, b, this.lang));
      const active = this.currentChecklist(c);
      const list = el('div', { class: 'mo-base-checklist' });
      const all = checkInput(t('Base.Filter.SelectAll'), !active);
      const boxes = [];
      if (blanks) { const b = checkInput(t('Base.Filter.Blanks'), !active || active.blanks); b.value = null; b.blank = true; boxes.push(b); }
      for (const v of values) {
        const b = checkInput(this.rs.display(c, v) || String(v), !active || active.keys.has(keyOf(v)));
        b.v = v;
        boxes.push(b);
      }
      all.input.addEventListener('change', () => boxes.forEach(b => { b.input.checked = all.input.checked; }));
      list.append(all, ...boxes);
      const ok = el('button', { type: 'button', class: 'mat-btn mat-btn--primary mat-btn--sm', text: t('Editor.Ok') });
      ok.addEventListener('click', () => {
        close();
        const chosen = boxes.filter(b => b.input.checked);
        if (chosen.length === boxes.length) { this.setColumnFilter(c, null); return; }
        const ref = this.colRef(c);
        const vals = chosen.filter(b => !b.blank).map(b => this.literal(c, b.v));
        const parts = [];
        if (vals.length === 1) parts.push(ref + '=' + vals[0]);
        else if (vals.length > 1) parts.push(ref + ' In (' + vals.join(',') + ')');
        if (chosen.some(b => b.blank)) parts.push(ref + ' Is Null');
        this.setColumnFilter(c, parts.length ? parts.join(' Or ') : 'False');
      });
      box.append(list, el('div', { class: 'mo-base-colmenu__foot' }, ok, el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', text: t('Editor.Cancel'), onclick: () => close() })));
      return box;
    });
  }

  /** The values of an active check list filter of a column (to show the boxes as they are). */
  currentChecklist(c) {
    const term = this.filterTerms().find(x => this.termColumn(x) === c);
    if (!term || !this.s.filterOn) return null;
    const keys = new Set();
    let blanks = false;
    const visit = x => {
      if (x.t === 'bin' && x.op === 'OR') { visit(x.l); visit(x.r); return; }
      if (x.t === 'isnull' && !x.not) { blanks = true; return; }
      const lit = n => (n.t === 'date' ? new DateVal(n.v) : n.t === 'str' || n.t === 'num' || n.t === 'bool' ? n.v : undefined);
      if (x.t === 'bin' && x.op === '=') { const v = lit(x.r); if (v !== undefined) keys.add(keyOf(v)); return; }
      if (x.t === 'in' && !x.not) x.list.forEach(n => { const v = lit(n); if (v !== undefined) keys.add(keyOf(v)); });
    };
    visit(term);
    if (!keys.size && !blanks) return null;
    return { keys, blanks };
  }

  /** Text, number and date filters: a condition with an operator and one or two values. */
  async customFilter(c) {
    const numeric = ['number', 'currency', 'autonumber'].includes(c.type);
    const dated = c.type === 'datetime';
    const ops = numeric || dated
      ? [['=', t('Base.Op.Equals')], ['<>', t('Base.Op.NotEquals')], ['<', dated ? t('Base.Op.Before') : t('Base.Op.Less')], ['>', dated ? t('Base.Op.After') : t('Base.Op.Greater')], ['between', t('Base.Op.Between')], ['null', t('Base.Op.Blank')], ['notnull', t('Base.Op.NotBlank')]]
      : [['=', t('Base.Op.Equals')], ['<>', t('Base.Op.NotEquals')], ['begins', t('Base.Op.Begins')], ['notbegins', t('Base.Op.NotBegins')], ['contains', t('Base.Op.Contains')], ['notcontains', t('Base.Op.NotContains')], ['ends', t('Base.Op.Ends')], ['null', t('Base.Op.Blank')], ['notnull', t('Base.Op.NotBlank')]];
    const op = selectInput(ops.map(([value, label]) => ({ value, label })), ops[0][0]);
    const v1 = textInput('');
    const v2 = textInput('');
    const second = field(t('Base.Op.And'), v2);
    const sync = () => { second.hidden = op.value !== 'between'; v1.disabled = op.value === 'null' || op.value === 'notnull'; };
    op.addEventListener('change', sync);
    sync();
    const r = await dialog({
      title: t('Base.Filter.Custom', c.caption), icon: 'filter', body: el('div', { class: 'mo-base-form' }, field(t('Base.Filter.Condition'), op), field(t('Base.Filter.Value'), v1), second),
      onOpen: () => v1.focus(),
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    const ref = this.colRef(c);
    const lit = text => {
      if (numeric) { const n = parseNumber(text, this.lang); return n === null ? exprToText({ t: 'str', v: text }) : String(n); }
      if (dated) { const s = parseDate(text, this.lang); return s === null ? exprToText({ t: 'str', v: text }) : exprToText({ t: 'date', v: s }); }
      return exprToText({ t: 'str', v: text });
    };
    const like = p => exprToText({ t: 'str', v: p });
    const x = v1.value;
    const cond = {
      '=': () => ref + '=' + lit(x), '<>': () => ref + '<>' + lit(x), '<': () => ref + '<' + lit(x), '>': () => ref + '>' + lit(x),
      between: () => ref + ' Between ' + lit(x) + ' And ' + lit(v2.value), null: () => ref + ' Is Null', notnull: () => ref + ' Is Not Null',
      begins: () => ref + ' Like ' + like(x + '*'), notbegins: () => ref + ' Not Like ' + like(x + '*'), contains: () => ref + ' Like ' + like('*' + x + '*'),
      notcontains: () => ref + ' Not Like ' + like('*' + x + '*'), ends: () => ref + ' Like ' + like('*' + x),
    }[op.value]();
    this.setColumnFilter(c, cond);
  }

  /** Advanced → Filter by Form is too much; Advanced Filter: the filter text itself (an expression on the columns). */
  async editFilter() {
    const input = el('textarea', { class: 'mat-input mat-textarea mo-base-code', rows: 4, spellcheck: 'false' });
    input.value = this.s.filter || '';
    const r = await dialog({
      title: t('Base.Filter.Advanced'), icon: 'filter', size: 'md', body: el('div', { class: 'mo-base-form' }, field(t('Base.Filter.Expression'), input, t('Base.Filter.ExpressionHint'))),
      onOpen: () => input.focus(),
      validate: () => { try { if (input.value.trim()) this.compileOnRows(input.value); return true; } catch (e) { showProblem(e); return false; } },
      buttons: [{ label: t('Base.Filter.Apply'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    this.s.filter = input.value.trim();
    this.s.filterOn = !!this.s.filter;
    this.persist();
    this.refreshView();
  }

  // ---- columns: width, hide, freeze -------------------------------------------------------------------------------------------------

  startResize(e, c) {
    e.preventDefault();
    e.stopPropagation();
    const start = this.widthOf(c);
    drag(e, dx => {
      this.s.widths = { ...(this.s.widths || {}), [c.skey]: Math.max(30, Math.round(start + dx)) };
      this.grid.style.setProperty('--mo-ds-cols', this.template());
      this.grid.style.setProperty('--mo-ds-width', SEL_W + this.cols.reduce((s, x) => s + this.widthOf(x), 0) + (this.o.addField && !this.app.readOnly ? ADD_W : 0) + 'px');
    }, () => { this.persist(); this.render(); });
  }

  bestFit(c) {
    const canvas = this._measure ||= document.createElement('canvas').getContext('2d');
    canvas.font = getComputedStyle(this.body).font || '13px sans-serif';
    let w = canvas.measureText(c.caption).width + 44;
    for (const r of this.view.slice(0, 2000)) w = Math.max(w, canvas.measureText(this.rs.display(c, this.rs.get(r, c))).width + 18);
    this.s.widths = { ...(this.s.widths || {}), [c.skey]: Math.min(600, Math.ceil(w)) };
    this.persist();
    this.render();
  }

  async fieldWidth() {
    const c = this.currentCol;
    if (!c) return;
    const input = textInput(String(this.widthOf(c)), { type: 'number', min: 30, max: 2000 });
    const r = await dialog({ title: t('Base.Ds.FieldWidth'), icon: 'sliders', body: field(t('Base.Ds.WidthPx'), input), onOpen: () => input.select(), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Base.Ds.BestFit'), value: 'fit' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'fit') { this.bestFit(c); return; }
    if (r !== 'ok') return;
    const n = Math.max(30, Math.min(2000, Number(input.value) || 120));
    this.s.widths = { ...(this.s.widths || {}), [c.skey]: n };
    this.persist();
    this.render();
  }

  hideColumn(c = this.currentCol) {
    if (!c) return;
    this.s.hidden = [...new Set([...(this.s.hidden || []), c.skey])];
    this.persist();
    this.cur.c = Math.min(this.cur.c, this.cols.length - 2);
    this.refreshView();
  }

  async unhideColumns() {
    const hidden = new Set(this.s.hidden || []);
    const boxes = this.rs.columns.map(c => { const b = checkInput(c.caption, !hidden.has(c.skey)); b.col = c; return b; });
    const r = await dialog({ title: t('Base.Ds.Unhide'), icon: 'eye', body: el('div', { class: 'mo-base-checklist mo-base-checklist--dialog' }, ...boxes), buttons: [{ label: t('Editor.Close'), value: 'ok', kind: 'primary' }] });
    if (r !== 'ok' && r !== null) return;
    this.s.hidden = boxes.filter(b => !b.input.checked).map(b => b.col.skey);
    this.persist();
    this.refreshView();
  }

  freezeColumn(c = this.currentCol) {
    if (!c) return;
    const order = this.columnsInOrder().map(x => x.skey);
    const frozen = Math.min(this.s.frozen || 0, order.length);
    const rest = order.filter(k => k !== c.skey);
    rest.splice(frozen, 0, c.skey);
    this.s.order = rest;
    this.s.frozen = frozen + 1;
    this.persist();
    this.refreshView();
  }

  unfreezeAll() {
    this.s.frozen = 0;
    this.persist();
    this.refreshView();
  }

  async toggleTotals() {
    this.s.totalsOn = !this.s.totalsOn;
    this.persist();
    this.renderTotals();
  }

  // ---- context menus ---------------------------------------------------------------------------------------------------------------

  headerContext(i, at) {
    const c = this.cols[i];
    const numeric = ['number', 'currency', 'autonumber'].includes(c.type);
    contextMenu(at, [
      { label: numeric ? t('Base.Sort.SmallLarge') : t('Base.Sort.AZ'), icon: 'base-sort-asc', run: () => this.sortBy(c, false) },
      { label: numeric ? t('Base.Sort.LargeSmall') : t('Base.Sort.ZA'), icon: 'base-sort-desc', run: () => this.sortBy(c, true) },
      '-',
      { label: t('Base.Copy'), icon: 'copy', run: () => this.copyColumn(c) },
      { label: t('Base.Ds.FieldWidth'), icon: 'sliders', run: () => this.fieldWidth() },
      { label: t('Base.Ds.Hide'), icon: 'eye-off', run: () => this.hideColumn(c) },
      { label: t('Base.Ds.Unhide'), icon: 'eye', run: () => this.unhideColumns() },
      { label: t('Base.Ds.Freeze'), icon: 'base-freeze', run: () => this.freezeColumn(c) },
      { label: t('Base.Ds.Unfreeze'), icon: 'base-unfreeze', run: () => this.unfreezeAll(), disabled: !this.s.frozen },
      '-',
      { label: t('Base.Find.Title'), icon: 'search', run: () => this.openFind(false) },
      ...(this.o.fieldMenu ? ['-', ...this.o.fieldMenu(c)] : []),
    ]);
  }

  copyColumn(c) {
    const text = [c.caption, ...this.view.map(r => this.rs.display(c, this.rs.get(r, c)))].join('\n');
    navigator.clipboard?.writeText(text).then(() => this.app.toast(t('Base.Copied'))).catch(() => {});
  }

  async cellContext(i, ci, at) {
    if (!(await this.moveTo(i, ci, { focus: true }))) return;
    const editable = !this.readOnlyCell(this.currentCol, i === this.newIndex);
    contextMenu(at, [
      { label: t('Base.Cut'), icon: 'base-cut', run: () => this.cut(), disabled: !editable },
      { label: t('Base.Copy'), icon: 'copy', run: () => this.copy() },
      { label: t('Base.Paste'), icon: 'base-paste', run: () => this.paste(), disabled: !editable },
      '-',
      { label: this.selectionLabel('equal'), icon: 'filter', run: () => this.filterBySelection('equal') },
      { label: this.selectionLabel('notEqual'), icon: 'filter', run: () => this.filterBySelection('notEqual') },
      { label: this.selectionLabel('contains'), icon: 'filter', run: () => this.filterBySelection('contains') },
      { label: this.selectionLabel('notContains'), icon: 'filter', run: () => this.filterBySelection('notContains') },
      { label: t('Base.Filter.ClearFrom', this.currentCol.caption), icon: 'base-filter-off', run: () => this.setColumnFilter(this.currentCol, null), disabled: !this.isFiltered(this.currentCol) },
      '-',
      { label: t('Base.Sort.AZ'), icon: 'base-sort-asc', run: () => this.sortBy(this.currentCol, false) },
      { label: t('Base.Sort.ZA'), icon: 'base-sort-desc', run: () => this.sortBy(this.currentCol, true) },
      ...(this.currentCol.type === 'memo' || this.currentCol.type === 'text' ? ['-', { label: t('Base.Zoom'), icon: 'search', run: () => this.zoom() }] : []),
    ]);
  }

  recordContext(at) {
    contextMenu(at, [
      { label: t('Base.Rec.New'), icon: 'base-new-record', run: () => this.moveTo(this.newIndex, 0, { focus: true }), disabled: !this.canAdd },
      { label: t('Base.Rec.Delete'), icon: 'trash', danger: true, run: () => this.deleteRecords(), disabled: this.readOnly || !this.rs.canDelete },
      '-',
      { label: t('Base.Copy'), icon: 'copy', run: () => this.copy() },
      { label: t('Base.Paste'), icon: 'base-paste', run: () => this.paste(), disabled: !this.canAdd },
    ]);
  }

  // ---- find and replace -------------------------------------------------------------------------------------------------------------

  /** Searches all fields from the record after the current one (the search box of the navigation bar). */
  searchFor(text, next = false) {
    if (!text) return;
    const n = this.view.length;
    const needle = text.toLowerCase();
    const start = next ? this.cur.r + 1 : this.cur.r;
    for (let k = 0; k < n; k++) {
      const r = (start + k) % n;
      for (let c = 0; c < this.cols.length; c++) {
        const shown = this.rs.display(this.cols[c], this.rs.get(this.view[r], this.cols[c]));
        if (shown.toLowerCase().includes(needle)) { this.moveTo(r, c); return; }
      }
    }
  }

  openFind(replace = false) {
    this.closeFind();
    const what = textInput(this.lastFind || '', { 'aria-label': t('Base.Find.What') });
    const withText = textInput(this.lastReplace || '', { 'aria-label': t('Base.Find.With') });
    const lookIn = selectInput([{ value: 'field', label: t('Base.Find.CurrentField') }, { value: 'all', label: t('Base.Find.Document') }], 'all');
    const match = selectInput([{ value: 'any', label: t('Base.Find.AnyPart') }, { value: 'whole', label: t('Base.Find.Whole') }, { value: 'start', label: t('Base.Find.Start') }], 'any');
    const dir = selectInput([{ value: 'all', label: t('Base.Find.All') }, { value: 'down', label: t('Base.Find.Down') }, { value: 'up', label: t('Base.Find.Up') }], 'all');
    const matchCase = checkInput(t('Base.Find.Case'), false);
    const tabs = el('div', { class: 'mo-base-find__tabs', role: 'tablist' });
    const replaceRow = field(t('Base.Find.With'), withText);
    const replaceButtons = el('div', { class: 'mo-base-find__more' });
    const setMode = r => {
      replaceRow.hidden = !r;
      replaceButtons.hidden = !r || this.readOnly;
      tabs.querySelectorAll('button').forEach((b, i) => b.setAttribute('aria-selected', String((i === 1) === r)));
    };
    tabs.append(
      el('button', { type: 'button', role: 'tab', text: t('Base.Find.Find'), onclick: () => setMode(false) }),
      el('button', { type: 'button', role: 'tab', text: t('Base.Find.Replace'), onclick: () => setMode(true), disabled: this.readOnly ? true : null }),
    );
    const opts = () => ({ what: what.value, lookIn: lookIn.value, match: match.value, dir: dir.value, matchCase: matchCase.input.checked });
    const findNext = () => { this.lastFind = what.value; if (!this.findNext(opts())) this.app.toast(t('Base.Find.NotFound'), 'info'); };
    replaceButtons.append(
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', text: t('Base.Find.ReplaceOne'), onclick: () => { this.lastReplace = withText.value; this.replaceOne(opts(), withText.value); } }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', text: t('Base.Find.ReplaceAll'), onclick: () => { this.lastReplace = withText.value; this.replaceAll(opts(), withText.value); } }),
    );
    what.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); findNext(); } });
    const panel = el('div', { class: 'mo-base-find', role: 'dialog', 'aria-label': t('Base.Find.Title') },
      el('div', { class: 'mo-base-find__head' }, tabs, iconButton('x', t('Editor.Close'), () => this.closeFind())),
      el('div', { class: 'mo-base-find__body' },
        field(t('Base.Find.What'), what), replaceRow,
        el('div', { class: 'mo-base-find__opts' }, field(t('Base.Find.LookIn'), lookIn), field(t('Base.Find.Match'), match), field(t('Base.Find.Search'), dir)),
        matchCase),
      el('div', { class: 'mo-base-find__foot' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary mat-btn--sm', text: t('Base.Find.Next'), onclick: findNext }), replaceButtons));
    panel.addEventListener('keydown', e => { if (e.key === 'Escape') { e.stopPropagation(); this.closeFind(); this.focusEditor(false); } });
    this.findPanel = panel;
    this.el.append(panel);
    setMode(replace && !this.readOnly);
    what.focus();
    what.select();
  }

  closeFind() { this.findPanel?.remove(); this.findPanel = null; }

  matches(text, o) {
    const a = o.matchCase ? text : text.toLowerCase();
    const b = o.matchCase ? o.what : o.what.toLowerCase();
    if (!b) return false;
    return o.match === 'whole' ? a === b : o.match === 'start' ? a.startsWith(b) : a.includes(b);
  }

  findNext(o) {
    const n = this.view.length;
    if (!n || !o.what) return false;
    const cols = o.lookIn === 'field' ? [this.cur.c] : this.cols.map((_, i) => i);
    const step = o.dir === 'up' ? -1 : 1;
    const cells = n * cols.length;
    let r = this.cur.r === this.newIndex ? 0 : this.cur.r;
    let ci = Math.max(0, cols.indexOf(this.cur.c));
    for (let k = 1; k <= cells; k++) {
      ci += step;
      if (ci >= cols.length) { ci = 0; r++; }
      if (ci < 0) { ci = cols.length - 1; r--; }
      if (r >= n || r < 0) {
        if (o.dir !== 'all') return false;
        r = (r + n) % n;
      }
      const c = this.cols[cols[ci]];
      if (this.matches(this.rs.display(c, this.rs.get(this.view[r], c)), o)) { this.moveTo(r, cols[ci]); return true; }
    }
    return false;
  }

  async replaceOne(o, replacement) {
    const c = this.currentCol;
    const { out } = this.valuesOf(this.cur.r);
    if (c && this.editor && this.editor.tagName === 'INPUT' && !this.editor.readOnly && this.matches(this.rs.display(c, out[c.key]), o)) {
      const text = this.editor.value;
      this.editor.value = o.match === 'any' ? text.replace(new RegExp(escapeRe(o.what), o.matchCase ? '' : 'i'), replacement) : replacement;
      this.markDirty();
      if (!(await this.commit())) return;
    }
    this.findNext(o);
  }

  async replaceAll(o, replacement) {
    if (!(await this.commit())) return;
    const cols = o.lookIn === 'field' ? [this.currentCol] : this.cols.filter(c => c.editable && c.field && !c.field.lookup && c.type !== 'yesno');
    let count = 0;
    let failed = 0;
    for (const row of this.view) {
      const changes = {};
      for (const c of cols) {
        const text = this.rs.display(c, this.rs.get(row, c));
        if (!this.matches(text, o)) continue;
        const next = o.match === 'any' ? text.replace(new RegExp(escapeRe(o.what), o.matchCase ? 'g' : 'gi'), replacement) : replacement;
        const p = this.rs.parse(c, next);
        if (p.ok) changes[c.key] = p.value; else failed++;
      }
      if (!Object.keys(changes).length) continue;
      const r = this.rs.save(row.id, changes);
      if (r.ok) count++; else failed++;
    }
    this.reload();
    this.app.toast(t('Base.Find.Replaced', count, failed), failed ? 'warning' : 'success');
  }
}

function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** A link that may be opened: http(s), mailto, tel, ftp, or a web address without a scheme. */
export function safeHref(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  if (/^(https?|mailto|tel|ftp):/i.test(s)) return s;
  if (/^[a-z][a-z0-9+.-]*:/i.test(s)) return null;
  if (/^www\.|^[\w-]+(\.[\w-]+)+(\/|$)/i.test(s)) return 'https://' + s;
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s)) return 'mailto:' + s;
  return null;
}

export { isNull, compare, toText };
