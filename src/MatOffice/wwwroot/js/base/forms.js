/* MatOffice Base – forms like Access: the Form view shows one record (Single Form), many (Continuous Forms, a row per record), a datasheet
   (Datasheet) or both (Split Form), with the record navigation bar, text boxes, combo boxes for lookups, check boxes for Yes/No, date pickers,
   calculated controls (=[Price]*[Qty], =Sum([Amount]) in the footer), command buttons (next, new, save, delete …) and subforms that show the
   related records of the record (Orders in a Customer form) as a datasheet. Records are saved when they are left, like Access; Esc undoes the
   field, Esc again the record. The Layout and Design views are the designer (designer.js) with the tabs of a form. */
import { t, el, icon, dialog } from '../editor/util.js';
import { BaseView, datasheetCan, datasheetRun, datasheetActive } from './view.js';
import { Datasheet, safeHref } from './datasheet.js';
import { recordsetOf, subformRecordset, EmptyRecordset } from './recordset.js';
import { ControlEval, domainFunction } from './layout.js';
import { Designer, styleControl, actionName } from './designer.js';
import { askParams } from './queryviews.js';
import { parseSql } from './sql.js';
import { showProblem, ask, iconButton, field, textInput, selectInput, checkInput, problemText } from './ui.js';
import { keyOf, lower, queryById, lookupOptions } from './model.js';
import { BaseError, Parser, compile, sortCompare, exprToText } from './expr.js';
import { DateVal, serialOfIso, formatValue, todaySerial } from './values.js';

const SEL_W = 18;

/** The controls of a section in the order people read them (and Tab goes through them). */
const inOrder = list => [...list].sort((a, b) => a.y - b.y || a.x - b.x);

export class FormView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.cur = 0;
    this.pending = null;
    this.s = { sort: [], filter: '', filterOn: false };
    this.binds = [];
    this.subs = [];
    this.selectContextTab = false;
  }

  get form() { return this.obj; }
  get kind() { return this.form?.kind || 'single'; }
  get canEditRecords() { return !this.app.readOnly && this.form.allowEdits !== false && !!this.rs?.canEdit; }
  get canAdd() { return !this.app.readOnly && this.form.allowAdditions !== false && !!this.rs?.canAdd; }
  get canDelete() { return !this.app.readOnly && this.form.allowDeletions !== false && !!this.rs?.canDelete; }
  get newIndex() { return this.canAdd ? this.view.length : -1; }
  get count() { return this.view.length + (this.canAdd ? 1 : 0); }
  get row() { return this.cur < this.view.length ? this.view[this.cur] : null; }
  get isNew() { return this.cur === this.newIndex; }

  // ---- loading ---------------------------------------------------------------------------------------------------------------------

  async mount() {
    const f = this.form;
    this.el.replaceChildren(el('div', { class: 'mo-loading' }, el('span', { class: 'mo-loading__spinner' })));
    this.params = {};
    const q = queryById(this.db, f.source);
    if (q) {
      try {
        const p = await askParams(this.app, parseSql(q.sql));
        if (p === null) { this.app.open('forms', f.id, this.app.readOnly ? 'design' : 'layout'); return; }
        this.params = p;
      } catch (e) {
        if (!(e instanceof BaseError)) throw e;
        this.fail(e);
        return;
      }
    }
    this.build();
    if (this.app.active?.viewObj === this) this.focus();
  }

  fail(e) {
    this.rs = null;
    this.el.replaceChildren(el('div', { class: 'mo-base-empty' }, el('div', { class: 'mo-base-empty__icon' }, icon('alert')), el('h2', { text: t('Base.Form.Failed') }), el('p', { text: problemText(e) }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => this.app.open('forms', this.tab.id, 'design') }, icon('base-v-design'), el('span', { text: t('Base.View.Design') }))));
  }

  build() {
    const f = this.form;
    this.destroyParts();
    let rs = f.source ? recordsetOf(this.app, f.source, this.params) : null;
    if (rs?.error) { this.fail(rs.error); return; }
    if (f.source && !rs) { this.fail({ code: 'NoTable', args: [f.source] }); return; }
    rs ||= new EmptyRecordset(this.app);
    // the properties of the form narrow what the records allow
    const edits = rs.canEdit && f.allowEdits !== false;
    const adds = rs.canAdd && f.allowAdditions !== false;
    const dels = rs.canDelete && f.allowDeletions !== false;
    if (f.allowEdits === false) Object.defineProperty(rs, 'canEdit', { get: () => false });
    Object.defineProperty(rs, 'canAdd', { get: () => adds && !this.app.readOnly });
    Object.defineProperty(rs, 'canDelete', { get: () => dels && !this.app.readOnly });
    void edits;
    this.rs = rs;
    this.ev = new ControlEval(rs, { lang: this.app.lang, currency: this.db.settings.currency, domain: domainFunction(this.db, this.app.lang) });
    this.applyView();
    this.cur = Math.max(0, Math.min(this.cur, this.count - 1));
    this.render();
  }

  /** Filter and sort the records (the commands of the Home tab). */
  applyView() {
    let rows = this.rs.rows;
    this.filterError = null;
    if (this.s.filterOn && this.s.filter) {
      try {
        const fn = this.compileOnRows(this.s.filter);
        rows = rows.filter(r => { const v = fn(r); return v !== null && v !== undefined && v !== false && v !== 0; });
      } catch (e) {
        if (!(e instanceof BaseError)) throw e;
        this.filterError = e;
      }
    }
    const sort = this.s.sort.map(s => ({ col: this.rs.columns.find(c => c.skey === s.field), desc: s.desc })).filter(s => s.col);
    if (sort.length) {
      const lang = this.app.lang;
      const keyed = rows.map((r, i) => ({ r, i, k: sort.map(s => (this.rs.mapOf(s.col) ? this.rs.display(s.col, this.rs.get(r, s.col)) : this.rs.get(r, s.col))) }));
      keyed.sort((a, b) => {
        for (let j = 0; j < sort.length; j++) { const c = sortCompare(a.k[j], b.k[j], lang); if (c) return sort[j].desc ? -c : c; }
        return a.i - b.i;
      });
      rows = keyed.map(x => x.r);
    }
    this.view = rows;
  }

  compileOnRows(text) {
    const byName = new Map();
    for (const c of this.rs.columns) { byName.set(lower(c.name), c); if (!byName.has(lower(c.caption))) byName.set(lower(c.caption), c); }
    const fn = compile(Parser.expression(text), { lang: this.app.lang, resolve: parts => { const c = byName.get(lower(parts[parts.length - 1])); return c ? row => this.rs.get(row, c) : null; } });
    return row => fn(row);
  }

  // ---- the values of the current record --------------------------------------------------------------------------------------------

  /** The record as the controls show it: the saved values, the defaults of a new record, and what is being edited. */
  rowAt(i) {
    const base = i < this.view.length ? this.view[i] : null;
    const values = base ? { ...base.values } : {};
    if (!base && i === this.newIndex) {
      const d = this.rs.defaults();
      for (const c of this.rs.columns) values[c.key] = this.rs.toEngine(c, d[c.key] ?? null);
    }
    if (this.pending && i === this.cur) {
      for (const c of this.rs.columns) if (Object.prototype.hasOwnProperty.call(this.pending.values, c.key)) values[c.key] = this.rs.toEngine(c, this.pending.values[c.key]);
    }
    return { id: base?.id ?? null, values, refs: base?.refs, isNew: !base };
  }

  colOf(c) {
    if (!c.source || c.source.startsWith('=')) return null;
    return this.ev.column(c.source);
  }

  editableFor(col, rowIndex, c) {
    if (!col || !col.editable || c.locked || this.app.readOnly) return false;
    return rowIndex === this.newIndex ? this.canAdd : this.canEditRecords;
  }

  /** The list of a combo box: its Row Source, else the lookup of its field. */
  optionsOf(c, col) {
    const type = col?.field?.type || col?.type || 'text';
    if (c.rowSource?.kind === 'table' || c.rowSource?.kind === 'values') return lookupOptions(this.db, { type, lookup: c.rowSource }, this.app.lang);
    if (col?.field?.lookup) return this.rs.options(col);
    return [];
  }

  // ---- rendering ---------------------------------------------------------------------------------------------------------------------

  destroyParts() {
    for (const s of this.subs) s.ds?.destroy();
    this.subs = [];
    this.ds?.destroy();
    this.ds = null;
    this.binds = [];
    this.resize?.disconnect();
    this.resize = null;
    this.rowsBox = null;
    this.rowEls = null;
  }

  isFlow() { return this.app.isPhone() && this.form.width > (this.el.clientWidth || window.innerWidth) - 8; }

  render() {
    const f = this.form;
    if (!f || !this.rs) return;
    this.destroyParts();
    this.flow = this.isFlow();
    if (this.kind === 'datasheet') { this.renderDatasheet(); return; }
    this.scroll = el('div', { class: 'mo-base-fv__scroll' });
    this.paper = el('div', { class: 'mo-base-fv__paper' + (this.flow ? ' is-flow' : '') });
    if (!this.flow) this.paper.style.width = f.width + SEL_W + 'px';
    this.scroll.append(this.paper);
    this.navBar = el('div', { class: 'mo-base-ds__nav' });
    const parts = [this.scroll];
    if (this.kind === 'split') {
      this.split = el('div', { class: 'mo-base-fv__split' });
      parts.push(el('div', { class: 'mo-base-fv__splitbar', role: 'separator', 'aria-orientation': 'horizontal' }), this.split);
    }
    if (f.source) parts.push(this.navBar);
    this.el.replaceChildren(el('div', { class: 'mo-base-fv mo-base-fv--' + this.kind }, ...parts));
    if (this.kind === 'continuous') this.renderContinuous();
    else this.renderSingle();
    if (this.kind === 'split') this.renderSplitSheet();
    this.renderNav();
    this.app.setStatus();
    this.app.ribbonRefresh();
  }

  sectionEl(key, rowIndex, extraClass = '') {
    const sec = this.form.sections[key];
    if (!sec || sec.visible === false || (!sec.height && !sec.controls.length)) return null;
    const node = el('div', { class: 'mo-base-fv__sec mo-base-fv__sec--' + key + extraClass, 'data-sec': key });
    if (!this.flow) node.style.height = sec.height + 'px';
    if (sec.back) node.style.backgroundColor = sec.back;
    for (const c of this.flow ? inOrder(sec.controls) : sec.controls) {
      if (c.visible === false) continue;
      node.append(this.controlEl(c, key, rowIndex));
    }
    return node;
  }

  renderSingle() {
    const head = this.sectionEl('header', this.cur);
    const record = el('div', { class: 'mo-base-fv__record' });
    this.selector = el('div', { class: 'mo-base-fv__selector', 'aria-hidden': 'true' });
    const detail = this.sectionEl('detail', this.cur) || el('div', { class: 'mo-base-fv__sec' });
    record.append(this.selector, detail);
    const foot = this.sectionEl('footer', this.cur);
    this.paper.replaceChildren(...[head, record, foot].filter(Boolean));
    this.paintSelector();
  }

  renderContinuous() {
    const detail = this.form.sections.detail;
    this.rowH = Math.max(this.flow ? 0 : 20, detail.height || 30);
    const head = this.sectionEl('header', this.cur, ' is-sticky');
    this.rowsBox = el('div', { class: 'mo-base-fv__rows' });
    if (!this.flow) this.rowsBox.style.height = this.count * this.rowH + 'px';
    const foot = this.sectionEl('footer', this.cur, ' is-sticky-bottom');
    this.paper.replaceChildren(...[head, this.rowsBox, foot].filter(Boolean));
    this.rowEls = new Map();
    this.scroll.addEventListener('scroll', () => this.renderRows());
    this.resize = new ResizeObserver(() => this.renderRows());
    this.resize.observe(this.scroll);
    this.renderRows();
  }

  visibleRows() {
    if (this.flow) return [0, Math.min(this.count, 300) - 1];
    const top = this.scroll.scrollTop - this.rowsBox.offsetTop;
    const h = this.scroll.clientHeight || 500;
    return [Math.max(0, Math.floor(top / this.rowH) - 3), Math.min(this.count - 1, Math.ceil((top + h) / this.rowH) + 3)];
  }

  renderRows(force = false) {
    if (!this.rowsBox) return;
    const [first, last] = this.visibleRows();
    for (const [i, node] of this.rowEls) {
      if (force || i < first || i > last) {
        if (!force && node.contains(document.activeElement)) continue;
        node.remove();
        this.rowEls.delete(i);
        this.binds = this.binds.filter(b => b.rowIndex !== i);
      }
    }
    for (let i = first; i <= last; i++) {
      if (this.rowEls.has(i)) continue;
      const node = el('div', { class: 'mo-base-fv__crow' + (i === this.cur ? ' is-current' : '') + (i === this.newIndex ? ' is-new' : '') + (i % 2 ? ' is-alt' : ''), 'data-row': i });
      if (!this.flow) { node.style.top = i * this.rowH + 'px'; node.style.height = this.rowH + 'px'; }
      const sel = el('div', { class: 'mo-base-fv__selector' });
      sel.addEventListener('pointerdown', e => { e.preventDefault(); this.moveTo(i); });
      node.append(sel, this.sectionEl('detail', i) || el('div', { class: 'mo-base-fv__sec' }));
      node.addEventListener('focusin', () => { if (i !== this.cur) this.moveTo(i, { keepFocus: true }); });
      this.rowEls.set(i, node);
      this.rowsBox.append(node);
    }
    this.paintSelector();
  }

  paintSelector() {
    const mark = i => (this.pending && i === this.cur ? icon('edit') : i === this.cur ? icon('base-current') : i === this.newIndex ? el('span', { class: 'mo-base-ds__star', text: '*' }) : null);
    if (this.kind === 'continuous' && this.rowEls) {
      for (const [i, node] of this.rowEls) {
        node.classList.toggle('is-current', i === this.cur);
        const m = mark(i);
        node.firstChild.replaceChildren(...(m ? [m] : []));
      }
    } else if (this.selector) {
      const m = mark(this.cur);
      this.selector.replaceChildren(...(m ? [m] : []));
    }
  }

  renderDatasheet() {
    const cols = inOrder(this.form.sections.detail.controls).map(c => this.colOf(c)).filter(Boolean);
    const order = [...new Set(cols.map(c => c.skey))];
    const settings = { order, hidden: cols.length ? this.rs.columns.filter(c => !order.includes(c.skey)).map(c => c.skey) : [] };
    this.ds = new Datasheet(this.app, { recordset: this.rs, settings, onStatus: () => { this.app.setStatus(); this.app.ribbonRefresh(); } });
    this.el.replaceChildren(el('div', { class: 'mo-base-fv mo-base-fv--datasheet' }, this.ds.el));
  }

  renderSplitSheet() {
    this.ds = new Datasheet(this.app, {
      recordset: this.rs, settings: this.s,
      onStatus: () => { this.app.setStatus(); this.app.ribbonRefresh(); },
      onCurrent: row => {
        if (this.syncing) return;
        this.view = this.ds.view;
        const i = row ? this.view.indexOf(row) : this.newIndex;
        if (i >= 0 && i !== this.cur) this.moveTo(i, { fromSheet: true });
      },
    });
    this.view = this.ds.view;
    this.split.append(this.ds.el);
    this.ds.el.addEventListener('focusin', async () => { if (this.pending && !this.committing) await this.commitRecord(); });
  }

  // ---- controls ----------------------------------------------------------------------------------------------------------------------

  controlEl(c, key, rowIndex) {
    const node = el('div', { class: 'mo-base-fc mo-base-fc--' + c.type, 'data-id': c.id });
    if (!this.flow) styleControl(node, c);
    else { styleControl(node, c, { box: false }); if (c.type === 'subform' || c.h > 40) node.style.minHeight = c.h + 'px'; }
    switch (c.type) {
      case 'label': node.append(el('span', { class: 'mo-base-fc__text', text: c.caption || '' })); break;
      case 'line': node.classList.toggle('is-vertical', c.h > c.w); node.append(el('span', { class: 'mo-base-dc__line' })); break;
      case 'rect': break;
      case 'button': {
        const b = el('button', { type: 'button', class: 'mo-base-fc__button', text: c.caption || actionName(c.action) });
        b.addEventListener('click', () => this.action(c.action));
        node.append(b);
        break;
      }
      case 'subform': this.subformEl(c, node); break;
      default: this.boundEl(c, node, key, rowIndex);
    }
    return node;
  }

  boundEl(c, node, key, rowIndex) {
    const col = this.colOf(c);
    const lookupText = c.type === 'textbox' && col?.field?.lookup;
    const asCombo = c.type === 'combobox' || lookupText;
    const editable = this.editableFor(col, rowIndex, c);
    const bind = { c, col, node, key, rowIndex, editable, calc: !col };
    let input;
    if (c.type === 'checkbox') {
      input = el('input', { type: 'checkbox', class: 'mo-base-fc__check', 'aria-label': c.name || col?.caption || '', disabled: editable ? null : true });
      input.addEventListener('change', () => { this.commitField(bind); });
    } else if (asCombo && editable) {
      const options = this.optionsOf(c, col);
      const free = (c.rowSource?.kind === 'values') || (!c.rowSource && col?.field?.lookup?.kind === 'values' && !col.field.lookup.limit) || !options.length;
      if (free) {
        const listId = 'mo-base-fl-' + this.tab.id + '-' + c.id + '-' + rowIndex;
        input = el('input', { class: 'mo-base-fc__input', list: listId, autocomplete: 'off', 'aria-label': c.name || col?.caption || '' });
        node.append(el('datalist', { id: listId }, ...options.map(o => el('option', { value: o.label }))));
        bind.options = null;
      } else {
        input = el('select', { class: 'mo-base-fc__input mo-base-fc__select', 'aria-label': c.name || col?.caption || '' }, el('option', { value: '', text: '' }), ...options.map((o, i) => el('option', { value: String(i), text: o.label })));
        input.addEventListener('change', () => this.commitField(bind));
        bind.options = options;
      }
    } else {
      const area = (col?.type === 'memo' || c.h >= 44) && !['number', 'currency', 'datetime', 'autonumber'].includes(col?.type);
      input = el(area ? 'textarea' : 'input', { class: 'mo-base-fc__input' + (area ? ' mo-base-fc__area' : ''), autocomplete: 'off', 'aria-label': c.name || col?.caption || '', readonly: editable ? null : true, spellcheck: col?.type === 'text' || col?.type === 'memo' ? 'true' : 'false' });
      if (!editable) input.classList.add('is-readonly');
    }
    bind.input = input;
    if (input.tagName !== 'SELECT' && input.type !== 'checkbox') {
      input.addEventListener('focus', () => this.onFocus(bind));
      input.addEventListener('input', () => this.markDirty(bind));
      input.addEventListener('change', () => this.commitField(bind));
      input.addEventListener('blur', () => { if (this.focusBind === bind && !bind.invalid) this.fill(bind); });
    } else {
      input.addEventListener('focus', () => { this.focusBind = bind; this.lastBind = bind; this.app.ribbonRefresh(); });
    }
    input.addEventListener('keydown', e => this.onKey(e, bind));
    node.append(input);
    if (col?.type === 'datetime' && editable && input.tagName === 'INPUT') {
      const picker = el('input', { type: 'date', class: 'mo-base-ds__datepick', tabindex: '-1', 'aria-hidden': 'true' });
      picker.addEventListener('change', () => {
        if (!picker.value) return;
        input.value = formatValue(new DateVal(serialOfIso(picker.value)), { type: 'datetime', format: 'Short Date', lang: this.app.lang });
        this.markDirty(bind);
        this.commitField(bind);
      });
      node.append(el('button', { type: 'button', class: 'mo-base-fc__pick', tabindex: '-1', 'aria-label': t('Base.Ds.PickDate'), onclick: () => { try { picker.showPicker(); } catch { picker.click(); } } }, icon('calendar'), picker));
    }
    if (col?.type === 'hyperlink' && !editable) {
      const v = this.rowAt(rowIndex).values[col.key];
      const href = v ? safeHref(String(v)) : null;
      if (href) { input.remove(); node.append(el('a', { class: 'mo-base-fc__link', href, target: '_blank', rel: 'noopener noreferrer', text: String(v).replace(/^mailto:/, '') })); }
    }
    this.binds.push(bind);
    this.fill(bind);
  }

  /** Puts the value of the record into a control (formatted; the plain value while it has the focus). */
  fill(bind) {
    const { c, col, input } = bind;
    if (!input) return;
    const row = this.rowAt(bind.rowIndex);
    if (bind.calc) {
      let text = '';
      if (c.source) text = this.ev.text(c, row, this.view);
      if (input.type === 'checkbox') { let v = null; try { v = this.ev.value(c.source, row, this.view); } catch { v = null; } input.checked = v === true; return; }
      input.value = text;
      input.readOnly = true;
      return;
    }
    const v = row.values[col.key];
    if (input.type === 'checkbox') { input.checked = v === true || (typeof v === 'number' && v !== 0); return; }
    if (input.tagName === 'SELECT') {
      const i = bind.options.findIndex(o => keyOf(o.value) === keyOf(v));
      if (i < 0 && v !== null && v !== undefined) { input.append(el('option', { value: 'x', text: this.rs.display(col, v) })); input.value = 'x'; } else input.value = i >= 0 ? String(i) : '';
      return;
    }
    if (row.isNew && col.type === 'autonumber' && (v === null || v === undefined)) { input.value = t('Base.Ds.New'); bind.startText = input.value; return; }
    const focused = document.activeElement === input;
    const fmt = c.format ? { ...col, format: c.format, decimals: c.decimals ?? col.decimals } : col;
    input.value = focused ? this.rs.editText(col, v) : (c.type === 'combobox' || col.field?.lookup) && col.field?.lookup?.kind === 'table' ? this.rs.display(col, v) : this.rs.display(fmt, v);
    if (!focused && col.type === 'hyperlink') input.value = input.value.replace(/^mailto:/i, '');
    if (focused) bind.startText = input.value;
  }

  refreshCalc() {
    for (const b of this.binds) if (b.calc && document.activeElement !== b.input) this.fill(b);
  }

  // ---- editing -------------------------------------------------------------------------------------------------------------------------

  onFocus(bind) {
    this.focusBind = bind;
    this.lastBind = bind;
    if (bind.calc || !bind.editable) { this.app.ribbonRefresh(); return; }
    if (bind.invalid) {
      // back from the message about a wrong value: the value stays for a correction
      bind.invalid = false;
      requestAnimationFrame(() => { try { bind.input.select(); } catch { /* not every input */ } });
      return;
    }
    const row = this.rowAt(bind.rowIndex);
    const v = row.values[bind.col.key];
    if (bind.input.tagName !== 'SELECT' && !(row.isNew && bind.col.type === 'autonumber')) {
      bind.input.value = bind.col.field?.lookup?.kind === 'table' ? this.rs.display(bind.col, v) : this.rs.editText(bind.col, v);
      requestAnimationFrame(() => { try { bind.input.select(); } catch { /* not every input */ } });
    }
    bind.startText = bind.input.value;
    this.app.ribbonRefresh();
  }

  markDirty(bind) {
    if (!bind.editable || bind.calc || bind.rowIndex !== this.cur) return;
    this.startPending();
  }

  startPending() {
    if (this.pending) return;
    this.pending = { id: this.row?.id ?? null, values: this.isNew ? { ...this.rs.defaults() } : {} };
    this.paintSelector();
    this.app.ribbonRefresh();
  }

  /** Reads a control into the record being edited: the type of its field and its validation rule (Access checks them when the field is left). */
  async commitField(bind, { quiet = false } = {}) {
    if (!bind || bind.calc || !bind.editable || !bind.col) return true;
    const { input, col } = bind;
    let parsed;
    if (input.type === 'checkbox') parsed = { ok: true, value: input.checked };
    else if (input.tagName === 'SELECT') {
      const i = input.value;
      if (i === 'x') return true;
      parsed = { ok: true, value: i === '' ? null : bind.options[Number(i)].value };
    } else {
      if (input.value === bind.startText) return true;
      let text = input.value;
      if (col.field?.lookup?.kind === 'table') {
        const hit = this.rs.options(col).find(o => lower(o.label) === lower(text.trim()));
        parsed = hit ? { ok: true, value: hit.value } : text.trim() === '' ? { ok: true, value: null } : this.rs.parse(col, text);
      } else {
        if (col.field?.lookup?.kind === 'values') text = text.trim();
        parsed = this.rs.parse(col, text);
      }
    }
    if (!parsed.ok) {
      if (!quiet) { bind.invalid = true; await showProblem({ code: parsed.code === 'TooLong' || parsed.code === 'OutOfRange' ? parsed.code : 'TypeMismatch', args: [col.caption] }); input.focus(); }
      return false;
    }
    const rule = this.rs.checkRule(col, parsed.value);
    if (rule) {
      if (!quiet) { bind.invalid = true; await showProblem(rule); input.focus(); }
      return false;
    }
    if (bind.rowIndex !== this.cur) return true;
    this.startPending();
    this.pending.values[col.key] = parsed.value;
    bind.startText = input.value;
    for (const b of this.binds) if (b !== bind && b.col === col && b.rowIndex === bind.rowIndex) this.fill(b);
    this.refreshCalc();
    return true;
  }

  /** Saves the record (when it is left, Shift+Enter, Save). False when it cannot be saved: the person stays on it. */
  async commitRecord() {
    if (this.committing) return false;
    this.committing = true;
    try {
      const focus = this.focusBind && document.activeElement === this.focusBind.input ? this.focusBind : null;
      if (focus && !(await this.commitField(focus))) return false;
      if (!this.pending) return true;
      const p = this.pending;
      const r = this.rs.save(p.id, p.values);
      if (!r.ok) {
        await showProblem(r.error);
        const b = this.binds.find(x => x.rowIndex === this.cur && x.col && (x.col.key === r.error.field || x.col.field?.id === r.error.field));
        (b || focus)?.input?.focus();
        return false;
      }
      this.pending = null;
      if (p.id === null) {
        if (!this.view.includes(r.row)) this.view.push(r.row);
        this.cur = this.view.indexOf(r.row);
      } else {
        const vi = this.view.findIndex(x => x.id === r.row.id);
        if (vi >= 0) this.view[vi] = r.row;
      }
      this.afterSave(p.id === null);
      return true;
    } finally {
      this.committing = false;
    }
  }

  afterSave(added) {
    if (this.kind === 'continuous') {
      if (added && !this.flow) this.rowsBox.style.height = this.count * this.rowH + 'px';
      // the saved row shows what was saved, an added one gets a new record row after it (the row with the focus stays as it is)
      const keep = document.activeElement;
      for (const i of [this.cur, this.cur + 1]) {
        const n = this.rowEls.get(i);
        if (n && !n.contains(keep)) { n.remove(); this.rowEls.delete(i); this.binds = this.binds.filter(b => b.rowIndex !== i); }
      }
      this.renderRows();
      for (const b of this.binds) if (b.rowIndex === this.cur && document.activeElement !== b.input) this.fill(b);
    } else {
      for (const b of this.binds) if (document.activeElement !== b.input) this.fill(b); else if (!b.calc && b.col) b.startText = b.input.value;
      if (added) this.rebuildSubforms();
      this.paintSelector();
    }
    if (this.kind === 'split' && this.ds) { this.syncing = true; this.ds.reload(); this.ds.moveTo(this.cur, this.ds.cur.c); this.syncing = false; }
    this.renderNav();
    this.app.setStatus();
    this.app.ribbonRefresh();
  }

  /** Esc: the field first, then the whole record. */
  undo() {
    const b = this.focusBind;
    if (b && document.activeElement === b.input && b.input.tagName !== 'SELECT' && b.input.type !== 'checkbox' && b.input.value !== b.startText) {
      b.input.value = b.startText;
      return true;
    }
    if (this.pending) {
      this.pending = null;
      for (const x of this.binds) this.fill(x);
      if (b && document.activeElement === b.input) this.onFocus(b);
      this.paintSelector();
      this.refreshCalc();
      this.app.ribbonRefresh();
      return true;
    }
    return false;
  }

  /** Goes to a record (the record that was edited is saved first). */
  async moveTo(i, { keepFocus = false, fromSheet = false } = {}) {
    if (!this.rs) return false;
    i = Math.max(0, Math.min(i, this.count - 1));
    if (i === this.cur && !fromSheet) return true;
    const focusId = this.focusBind && this.el.contains(document.activeElement) ? this.focusBind.c.id : null;
    if (!(await this.commitRecord())) return false;
    const old = this.cur;
    this.cur = Math.max(0, Math.min(i, this.count - 1));
    if (this.kind === 'continuous') {
      for (const b of this.binds) if (b.rowIndex === old || b.rowIndex === this.cur) { b.editable = this.editableFor(b.col, b.rowIndex, b.c); this.fill(b); }
      this.renderRows();
      if (!keepFocus) this.scrollToRow(this.cur);
    } else {
      this.binds = [];
      for (const s of this.subs) s.ds?.destroy();
      this.subs = [];
      this.renderSingle();
      if (focusId && !keepFocus) this.binds.find(b => b.c.id === focusId)?.input?.focus();
      if (this.kind === 'split' && this.ds && !fromSheet) { this.syncing = true; await this.ds.moveTo(this.cur < this.view.length ? this.cur : this.ds.newIndex, this.ds.cur.c); this.syncing = false; }
    }
    this.renderNav();
    this.app.setStatus();
    this.app.ribbonRefresh();
    return true;
  }

  scrollToRow(i) {
    if (!this.rowsBox || this.flow) return;
    const top = this.rowsBox.offsetTop + i * this.rowH;
    if (top < this.scroll.scrollTop) this.scroll.scrollTop = top;
    else if (top + this.rowH > this.scroll.scrollTop + this.scroll.clientHeight) this.scroll.scrollTop = top + this.rowH - this.scroll.clientHeight;
  }

  onKey(e, bind) {
    if (e.key === 'Escape') { if (this.undo()) { e.preventDefault(); e.stopPropagation(); } return; }
    if (e.key === 'Enter' && e.shiftKey) { e.preventDefault(); this.commitRecord(); return; }
    if (e.key === 'PageDown' && !e.ctrlKey) { e.preventDefault(); this.moveTo(this.cur + 1); return; }
    if (e.key === 'PageUp' && !e.ctrlKey) { e.preventDefault(); this.moveTo(this.cur - 1); return; }
    if (this.kind === 'continuous' && (e.key === 'ArrowDown' || e.key === 'ArrowUp') && bind.input.tagName === 'INPUT' && !e.altKey) {
      e.preventDefault();
      const next = this.cur + (e.key === 'ArrowDown' ? 1 : -1);
      this.moveTo(next).then(ok => { if (ok) { this.scrollToRow(this.cur); this.renderRows(); this.binds.find(b => b.rowIndex === this.cur && b.c.id === bind.c.id)?.input?.focus(); } });
      return;
    }
    if (e.key === 'Enter' && bind.input.tagName === 'INPUT' && !e.altKey) {
      // Enter goes on to the next control, like Tab (the last one goes to the next record)
      e.preventDefault();
      bind.input.dispatchEvent(new Event('change'));
      const list = this.binds.filter(b => b.rowIndex === bind.rowIndex && b.input && !b.input.disabled && b.input.isConnected);
      const ordered = list.sort((a, b) => (a.key === b.key ? 0 : a.key === 'header' ? -1 : 1) || a.c.y - b.c.y || a.c.x - b.c.x);
      const i = ordered.indexOf(bind);
      if (i >= 0 && i < ordered.length - 1) ordered[i + 1].input.focus();
      else if (this.kind !== 'continuous') this.moveTo(this.cur + 1).then(ok => { if (ok) this.binds.filter(b => b.rowIndex === this.cur && b.editable).sort((a, b) => a.c.y - b.c.y || a.c.x - b.c.x)[0]?.input?.focus(); });
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === ';' && bind.col?.type === 'datetime' && bind.editable) {
      e.preventDefault();
      bind.input.value = formatValue(new DateVal(todaySerial()), { type: 'datetime', format: 'Short Date', lang: this.app.lang });
      this.markDirty(bind);
    }
  }

  // ---- subforms ----------------------------------------------------------------------------------------------------------------------

  subformEl(c, node) {
    const masterCol = c.subform?.linkMaster ? this.ev.column(c.subform.linkMaster) : null;
    const row = this.row;
    const master = row && masterCol ? this.rs.get(row, masterCol) : null;
    let rs = null;
    try { rs = subformRecordset(this.app, c.subform, master); } catch { rs = null; }
    if (!rs) { node.append(el('div', { class: 'mo-base-fc__error', text: '#Name?' })); return; }
    const ds = new Datasheet(this.app, { recordset: rs, settings: {}, compact: true, onStatus: () => this.app.ribbonRefresh() });
    node.append(ds.el);
    node.addEventListener('focusin', async () => {
      if (!this.pending || this.committing) return;
      const wasNew = this.pending.id === null;
      if (await this.commitRecord() && wasNew) this.rebuildSubforms();
    });
    this.subs.push({ c, ds, node, rs });
  }

  rebuildSubforms() {
    for (const s of this.subs) {
      s.ds.destroy();
      s.node.replaceChildren();
    }
    const list = this.subs;
    this.subs = [];
    for (const s of list) this.subformEl(s.c, s.node);
  }

  /** The datasheet the Home commands go to: the focused subform's, the datasheet of a datasheet or split form. */
  activeSheet() {
    const a = document.activeElement;
    const sub = this.subs.find(s => s.node.contains(a));
    if (sub) return sub.ds;
    if (this.ds && (this.kind === 'datasheet' || this.ds.el.contains(a))) return this.ds;
    return null;
  }

  // ---- commands ---------------------------------------------------------------------------------------------------------------------

  async action(a) {
    switch (a) {
      case 'first': return this.moveTo(0);
      case 'previous': return this.moveTo(this.cur - 1);
      case 'next': return this.moveTo(this.cur + 1);
      case 'last': return this.moveTo(Math.max(0, this.view.length - 1));
      case 'new': return this.canAdd ? this.moveTo(this.newIndex) : undefined;
      case 'save': return this.commitRecord();
      case 'delete': return this.deleteRecord();
      case 'undo': return this.undo();
      case 'close': if (await this.commitRecord()) this.app.closeTab(this.tab); return undefined;
      case 'print': return this.app.printObject('forms', this.tab.id);
      default: return undefined;
    }
  }

  async deleteRecord() {
    if (!this.canDelete) return;
    if (this.isNew) { this.undo(); return; }
    const row = this.row;
    if (!row) return;
    const plan = this.rs.planRemove([row.id]);
    if (!plan.ok) { await showProblem(plan.error); return; }
    if (!(await ask(plan.cascaded ? t('Base.Ds.DeleteCascade', 1, plan.cascaded) : t('Base.Ds.DeleteAsk', 1), { title: t('Base.Ds.DeleteTitle'), ok: t('Base.Yes'), cancel: t('Base.No'), danger: true }))) return;
    this.pending = null;
    this.rs.remove(plan);
    this.applyView();
    this.cur = Math.min(this.cur, Math.max(0, this.count - 1));
    this.render();
    this.app.toast(t('Base.Ds.Deleted', 1 + plan.cascaded));
  }

  focusCol() { return this.lastBind?.col || null; }

  can(cmd) {
    if (!this.rs) return false;
    const ds = this.activeSheet();
    if (ds) return datasheetCan(ds, cmd);
    switch (cmd) {
      case 'new': return this.canAdd;
      case 'save': return !!this.pending;
      case 'delete': return this.canDelete && !!this.row;
      case 'undoRecord': return !!this.pending;
      case 'refresh': case 'goto': return !!this.form.source;
      case 'find': return this.view.length > 0;
      case 'sortAsc': case 'sortDesc': case 'filterSel': return !!this.focusCol();
      case 'sortClear': return this.s.sort.length > 0;
      case 'filterToggle': case 'filterClear': return !!this.s.filter;
      case 'filter': case 'filterAdvanced': return this.rs.columns.length > 0;
      case 'copy': return !!this.focusBind;
      case 'cut': case 'paste': return !!this.focusBind?.editable;
      default: return false;
    }
  }

  active(cmd) {
    const ds = this.activeSheet();
    if (ds) return datasheetActive(ds, cmd);
    if (cmd === 'filterToggle') return !!(this.s.filterOn && this.s.filter);
    return false;
  }

  selectionLabel(kind) {
    const ds = this.activeSheet();
    if (ds) return ds.selectionLabel(kind);
    const col = this.focusCol();
    if (!col) return '';
    const v = this.rowAt(this.cur).values[col.key];
    const text = this.rs.display(col, v) || t('Base.Filter.Blanks');
    const short = text.length > 24 ? text.slice(0, 24) + '…' : text;
    return { equal: t('Base.Filter.Equals', short), notEqual: t('Base.Filter.NotEquals', short), contains: t('Base.Filter.Contains', short), notContains: t('Base.Filter.NotContains', short) }[kind];
  }

  async run(cmd, arg) {
    const ds = this.activeSheet();
    if (ds) {
      await datasheetRun(ds, cmd, arg);
      if (ds === this.ds && this.kind === 'split') this.syncFromSheet();
      return;
    }
    switch (cmd) {
      case 'new': await this.moveTo(this.newIndex); this.focusFirst(); return;
      case 'save': await this.commitRecord(); return;
      case 'delete': await this.deleteRecord(); return;
      case 'undoRecord': this.undo(); return;
      case 'refresh': if (await this.commitRecord()) { this.rs.load(); this.applyView(); this.cur = Math.min(this.cur, Math.max(0, this.count - 1)); this.render(); } return;
      case 'goto': {
        const r = { first: 0, previous: this.cur - 1, next: this.cur + 1, last: this.view.length - 1, new: this.newIndex }[arg];
        if (r >= 0) await this.moveTo(r);
        return;
      }
      case 'sortAsc': case 'sortDesc': {
        const col = this.focusCol();
        if (!col || !(await this.commitRecord())) return;
        this.s.sort = [{ field: col.skey, desc: cmd === 'sortDesc' }];
        this.reView();
        return;
      }
      case 'sortClear': if (await this.commitRecord()) { this.s.sort = []; this.reView(); } return;
      case 'filterSel': await this.filterBySelection(arg); return;
      case 'filterToggle': if (await this.commitRecord()) { this.s.filterOn = !this.s.filterOn; this.reView(); } return;
      case 'filterClear': if (await this.commitRecord()) { this.s.filter = ''; this.s.filterOn = false; this.reView(); } return;
      case 'filter': case 'filterAdvanced': await this.editFilter(); return;
      case 'find': await this.find(); return;
      case 'copy': { const b = this.focusBind; if (b?.input) { try { await navigator.clipboard.writeText(b.input.value ?? ''); this.app.toast(t('Base.Copied')); } catch { /* refused */ } } return; }
      case 'cut': { const b = this.focusBind; if (b?.editable && b.input.tagName !== 'SELECT' && b.input.type !== 'checkbox') { try { await navigator.clipboard.writeText(b.input.value); } catch { /* refused */ } b.input.value = ''; this.markDirty(b); b.input.focus(); } return; }
      case 'paste': {
        const b = this.focusBind;
        if (!b?.editable || b.input.tagName === 'SELECT' || b.input.type === 'checkbox') return;
        let text = '';
        try { text = await navigator.clipboard.readText(); } catch { this.app.toast(t('Base.PasteKeys'), 'info'); return; }
        b.input.focus();
        b.input.setRangeText(text.replace(/\r?\n$/, ''), b.input.selectionStart, b.input.selectionEnd, 'end');
        this.markDirty(b);
        return;
      }
      default: return;
    }
  }

  focusFirst() {
    requestAnimationFrame(() => this.binds.filter(b => b.rowIndex === this.cur && b.editable && b.input?.isConnected).sort((a, b) => (a.key === 'detail' ? 0 : 1) - (b.key === 'detail' ? 0 : 1) || a.c.y - b.c.y || a.c.x - b.c.x)[0]?.input?.focus());
  }

  reView() {
    const id = this.row?.id;
    this.applyView();
    const i = this.view.findIndex(r => r.id === id);
    this.cur = i >= 0 ? i : Math.min(this.cur, Math.max(0, this.count - 1));
    this.render();
    if (this.filterError) showProblem(this.filterError);
  }

  syncFromSheet() {
    this.s = this.ds.s;
    this.view = this.ds.view;
    this.cur = Math.max(0, Math.min(this.ds.cur.r, this.count - 1));
    this.binds = [];
    for (const s of this.subs) s.ds?.destroy();
    this.subs = [];
    this.renderSingle();
  }

  async filterBySelection(kind) {
    const col = this.focusCol();
    if (!col || !(await this.commitRecord())) return;
    const v = this.rowAt(this.cur).values[col.key];
    const ref = '[' + col.name + ']';
    const lit = x => (x instanceof DateVal ? exprToText({ t: 'date', v: x.s }) : typeof x === 'boolean' ? (x ? 'True' : 'False') : typeof x === 'number' ? exprToText({ t: 'num', v: x }) : exprToText({ t: 'str', v: String(x) }));
    let cond;
    if (v === null || v === undefined) cond = kind === 'notEqual' ? ref + ' Is Not Null' : ref + ' Is Null';
    else if (kind === 'equal') cond = ref + '=' + lit(v);
    else if (kind === 'notEqual') cond = '(' + ref + '<>' + lit(v) + ' Or ' + ref + ' Is Null)';
    else cond = ref + (kind === 'contains' ? ' Like ' : ' Not Like ') + exprToText({ t: 'str', v: '*' + this.rs.display(col, v).replace(/[*?#[]/g, m => '[' + m + ']') + '*' });
    this.s.filter = this.s.filter && this.s.filterOn ? '(' + this.s.filter + ') And (' + cond + ')' : cond;
    this.s.filterOn = true;
    this.reView();
  }

  async editFilter() {
    if (!(await this.commitRecord())) return;
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
    this.reView();
  }

  /** Find: the next record whose fields (or the field of the focused control) contain the text. */
  async find() {
    const what = textInput(this.lastFind || '');
    const where = selectInput([{ value: 'all', label: t('Base.Find.Document') }, { value: 'field', label: t('Base.Find.CurrentField') }], 'all');
    const matchCase = checkInput(t('Base.Find.Case'), false);
    const r = await dialog({
      title: t('Base.Find.Title'), icon: 'search', body: el('div', { class: 'mo-base-form' }, field(t('Base.Find.What'), what), field(t('Base.Find.LookIn'), where), matchCase),
      onOpen: () => { what.focus(); what.select(); },
      buttons: [{ label: t('Base.Find.Next'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok' || !what.value) return;
    this.lastFind = what.value;
    const needle = matchCase.input.checked ? what.value : what.value.toLowerCase();
    const cols = where.value === 'field' && this.focusCol() ? [this.focusCol()] : this.rs.columns;
    const n = this.view.length;
    for (let k = 1; k <= n; k++) {
      const i = (this.cur + k) % n;
      const hit = cols.find(c => { const s = this.rs.display(c, this.rs.get(this.view[i], c)); return (matchCase.input.checked ? s : s.toLowerCase()).includes(needle); });
      if (hit) {
        if (await this.moveTo(i)) this.binds.find(b => b.rowIndex === this.cur && b.col === hit)?.input?.focus();
        return;
      }
    }
    this.app.toast(t('Base.Find.NotFound'), 'info');
  }

  // ---- navigation bar, status ----------------------------------------------------------------------------------------------------

  renderNav() {
    if (!this.navBar || this.kind === 'split') { if (this.navBar && this.kind === 'split') this.navBar.hidden = true; return; }
    const total = this.view.length;
    const pos = this.isNew ? '' : String(this.cur + 1);
    const box = el('input', { class: 'mo-base-ds__pos', value: pos, 'aria-label': t('Base.Nav.Current'), inputmode: 'numeric' });
    box.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); const n = Number(box.value); if (n >= 1 && n <= total) this.moveTo(n - 1); else box.value = pos; } });
    const btn = (name, label, run, off) => iconButton(name, label, run, { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: off ? true : null });
    const filtered = this.s.filterOn && this.s.filter;
    const search = el('input', { class: 'mo-base-ds__search', type: 'search', placeholder: t('Base.Nav.Search'), 'aria-label': t('Base.Nav.Search') });
    search.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); this.searchFor(search.value); } });
    this.navBar.replaceChildren(
      el('span', { class: 'mo-base-ds__navlabel', text: t('Base.Nav.Record') }),
      btn('base-first', t('Base.Nav.First'), () => this.moveTo(0), this.cur <= 0),
      btn('chevron-left', t('Base.Nav.Previous'), () => this.moveTo(this.cur - 1), this.cur <= 0),
      box,
      el('span', { class: 'mo-base-ds__navof', text: t('Base.Nav.Of', total) + (filtered ? ' ' + t('Base.Nav.FilteredShort') : '') }),
      btn('chevron-right', t('Base.Nav.Next'), () => this.moveTo(this.cur + 1), this.cur >= this.count - 1),
      btn('base-last', t('Base.Nav.Last'), () => this.moveTo(total - 1), this.cur >= total - 1),
      this.canAdd ? btn('base-new-record', t('Base.Nav.New'), () => this.moveTo(this.newIndex).then(() => this.focusFirst()), this.isNew) : null,
      el('button', { type: 'button', class: 'mo-base-ds__filterflag' + (filtered ? ' is-on' : ''), disabled: this.s.filter ? null : true, onclick: () => this.run('filterToggle') },
        icon(filtered ? 'filter' : 'base-filter-off'), el('span', { text: filtered ? t('Base.Nav.Filtered') : this.s.filter ? t('Base.Nav.Unfiltered') : t('Base.Nav.NoFilter') })),
      search);
  }

  searchFor(text) {
    if (!text) return;
    const needle = text.toLowerCase();
    const n = this.view.length;
    for (let k = 1; k <= n; k++) {
      const i = (this.cur + k) % n;
      if (this.rs.columns.some(c => this.rs.display(c, this.rs.get(this.view[i], c)).toLowerCase().includes(needle))) { this.moveTo(i); return; }
    }
  }

  status() {
    if (!this.rs) return [];
    const ds = this.activeSheet();
    const out = [];
    if (this.form.source && !ds) out.push(this.isNew ? t('Base.Ds.NewRecord') : this.view.length ? t('Base.Nav.NOfM', this.cur + 1, this.view.length) : t('Base.Nav.NoRecords'));
    if ((ds?.s || this.s).filterOn && (ds?.s || this.s).filter) out.push(t('Base.Status.Filtered'));
    if (this.app.readOnly || this.form.allowEdits === false) out.push(t('Base.Status.ReadOnlyForm'));
    return out;
  }

  // ---- the view ---------------------------------------------------------------------------------------------------------------------

  refresh() {
    if (!this.form) return;
    if (this.pending) return;
    const id = this.row?.id;
    this.build();
    const i = this.view?.findIndex(r => r.id === id) ?? -1;
    if (i >= 0 && i !== this.cur) { this.cur = i; this.render(); }
  }

  focus() {
    if (this.kind === 'datasheet' && this.ds) { requestAnimationFrame(() => { if (!this.ds.editor) this.ds.moveTo(this.ds.cur.r, this.ds.cur.c, { focus: true }); else this.ds.focusEditor(false); }); return; }
    if (!this.el.contains(document.activeElement)) this.focusFirst();
  }

  async leave() {
    for (const s of this.subs) if (!(await s.ds.commit())) return false;
    if (this.ds && !(await this.ds.commit())) return false;
    return this.commitRecord();
  }

  destroy() { this.destroyParts(); }
}

// ---- Layout and Design view ------------------------------------------------------------------------------------------------------------

export class FormDesignView extends Designer {
  contextTabs() {
    const types = ['label', 'textbox', 'combobox', 'checkbox', 'button', 'subform', 'line', 'rect'];
    return [
      {
        id: 'fdesign', label: this.live ? t('Base.Tab.FormLayout') : t('Base.Tab.FormDesign'), groups: [
          this.controlsGroup(types),
          { label: t('Base.Group.HeaderFooter'), items: [{ type: 'stack', rows: [
            [{ type: 'button', showLabel: true, label: t('Base.Dz.Title'), icon: 'base-title', run: () => this.addTitle() }],
            [{ type: 'button', showLabel: true, label: t('Base.Dz.DateTime'), icon: 'base-date-time', run: () => this.addDate() }],
            [{ type: 'toggle', showLabel: true, label: t('Base.Dz.HeaderFooter'), icon: 'base-header-footer', active: () => (this.obj?.sections.header.height || 0) > 0 || (this.obj?.sections.footer.height || 0) > 0, run: () => this.toggleHeaderFooter() }],
          ] }] },
          this.toolsGroup(),
          { label: t('Base.Group.Views'), items: [{ type: 'button', size: 'large', label: t('Base.View.Form'), icon: 'base-v-form', edit: false, run: () => this.app.open('forms', this.tab.id, 'form') }] },
        ],
      },
      this.formatTab(),
      this.arrangeTab(),
    ];
  }

  /** Form Header/Footer on or off (off keeps the controls only when the person agrees). */
  async toggleHeaderFooter() {
    if (this.ro) return;
    const s = this.obj.sections;
    const on = s.header.height > 0 || s.footer.height > 0;
    if (on && (s.header.controls.length || s.footer.controls.length) && !(await ask(t('Base.Dz.RemoveHeaderFooter'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    this.app.design(() => {
      if (on) { s.header.height = 0; s.footer.height = 0; s.header.controls = []; s.footer.controls = []; } else { s.header.height = 60; s.footer.height = 40; }
    });
    this.sel.clear();
    this.render();
  }
}
