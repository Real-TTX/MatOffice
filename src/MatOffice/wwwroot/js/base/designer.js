/* MatOffice Base – the designer of forms and reports (the Layout and Design views of Access): the sections as bands (Form Header, Detail, Form
   Footer; Report Header, Page Header, the headers and footers of the groups, Page Footer, Report Footer) on a grid, controls to select (a click,
   Shift for more, a frame drawn with the mouse), to move and size with the mouse or the arrow keys (Shift sizes), new controls from the ribbon,
   fields from the Field List, the Property Sheet of the selection, alignment, order, cut/copy/paste. Every change is one undo step.
   The Layout view of a form shows the data of a record in its controls; the Design view shows their sources. */
import { t, el, icon, dialog } from '../editor/util.js';
import { BaseView } from './view.js';
import { newId, lower, tableById, queryById, fieldById } from './model.js';
import { contextMenu, drag, field, textInput, selectInput, iconButton, isTyping, nextId, KIND_ICON } from './ui.js';
import { sourceFieldList, controlWidth, ControlEval, domainFunction } from './layout.js';
import { FORMATS, formatLabel } from './tableviews.js';
import { FUNCTION_NAMES, Parser, BaseError } from './expr.js';
import { recordsetOf, subformRecordset } from './recordset.js';

const SNAP = 4;
const snap = (v, on = true) => (on ? Math.round(v / SNAP) * SNAP : Math.round(v));
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const SIZE = {
  label: { w: 100, h: 24 }, textbox: { w: 160, h: 24 }, combobox: { w: 160, h: 24 }, checkbox: { w: 18, h: 18 }, button: { w: 104, h: 30 },
  subform: { w: 420, h: 200 }, line: { w: 200, h: 1 }, rect: { w: 160, h: 80 },
};
const NAME_PREFIX = { label: 'Label', textbox: 'Text', combobox: 'Combo', checkbox: 'Check', button: 'Command', subform: 'Child', line: 'Line', rect: 'Box' };
export const ACTIONS = ['first', 'previous', 'next', 'last', 'new', 'save', 'delete', 'undo', 'close', 'print'];

/** The names of the control types (the catalog keys written out, so the check of the texts sees them). */
export function controlTypeName(type) {
  return {
    label: t('Base.Ctl.Label'), textbox: t('Base.Ctl.TextBox'), combobox: t('Base.Ctl.ComboBox'), checkbox: t('Base.Ctl.CheckBox'), button: t('Base.Ctl.Button'),
    subform: t('Base.Ctl.Subform'), line: t('Base.Ctl.Line'), rect: t('Base.Ctl.Rectangle'),
  }[type] || type;
}

export const CONTROL_ICON = {
  label: 'base-c-label', textbox: 'base-c-textbox', combobox: 'base-c-combo', checkbox: 'base-c-check', button: 'base-c-button', subform: 'base-c-subform', line: 'base-c-line', rect: 'base-c-rect',
};

/** The words of the actions of command buttons. */
export function actionName(a) {
  return {
    '': t('Base.Action.None'), first: t('Base.Action.First'), previous: t('Base.Action.Previous'), next: t('Base.Action.Next'), last: t('Base.Action.Last'), new: t('Base.Action.New'),
    save: t('Base.Action.Save'), delete: t('Base.Action.Delete'), undo: t('Base.Action.Undo'), close: t('Base.Action.Close'), print: t('Base.Action.Print'),
  }[a ?? ''] || a;
}

export const defaultBorder = type => type === 'textbox' || type === 'combobox' || type === 'subform' || type === 'rect';

/** A font size in points as pixels on the screen. */
export const px = pt => Math.round((Number(pt) || 11) * 4 / 3 * 10) / 10;

/** The box, the font and the colours of a control on an element. */
export function styleControl(node, c, { box = true } = {}) {
  const s = node.style;
  if (box) { s.left = c.x + 'px'; s.top = c.y + 'px'; s.width = c.w + 'px'; s.height = c.h + 'px'; }
  s.fontSize = px(c.fontSize) + 'px';
  if (c.bold) s.fontWeight = '700';
  if (c.italic) s.fontStyle = 'italic';
  s.textAlign = c.align || 'left';
  s.justifyContent = { left: 'flex-start', center: 'center', right: 'flex-end' }[c.align || 'left'];
  if (c.color) { s.color = c.color; s.borderColor = c.color; }
  if (c.back) s.backgroundColor = c.back;
  if (c.visible === false) node.classList.add('is-invisible');
  node.classList.toggle('has-border', !!(c.border ?? defaultBorder(c.type)));
}

/** A new control with every property of the content format. */
export function newControl(type, id, box, more = {}) {
  const c = {
    id, type, x: Math.max(0, Math.round(box.x)), y: Math.max(0, Math.round(box.y)), w: Math.max(1, Math.round(box.w)), h: Math.max(1, Math.round(box.h)),
    name: '', caption: '', source: '', format: '', decimals: null, fontSize: 11, bold: false, italic: false, align: 'left', color: null, back: null,
    border: defaultBorder(type), locked: false, visible: true, hideDuplicates: false, runningSum: 'no', action: '', rowSource: null, ...more,
  };
  if (type === 'subform') c.subform ??= { source: '', linkMaster: '', linkChild: '' };
  return c;
}

/** Every control of a form or report with the key of its section. */
export function controlsOfObject(obj) {
  const s = obj.sections || {};
  const out = [];
  const add = (key, sec) => { if (sec) for (const c of sec.controls || []) out.push({ c, key, sec }); };
  add('header', s.header); add('reportHeader', s.reportHeader); add('pageHeader', s.pageHeader);
  (s.groupHeaders || []).forEach((sec, i) => add('gh' + i, sec));
  add('detail', s.detail);
  (s.groupFooters || []).forEach((sec, i) => add('gf' + i, sec));
  add('footer', s.footer); add('pageFooter', s.pageFooter); add('reportFooter', s.reportFooter);
  return out;
}

let clipboard = null;

function readPanes() {
  try { return JSON.parse(localStorage.getItem('mo.base.designer') || '{}') || {}; } catch { return {}; }
}

function savePanes(state) {
  try { localStorage.setItem('mo.base.designer', JSON.stringify(state)); } catch { /* private mode */ }
}

/** The format names a control can take (with the words of the page). */
function formatOptions(current) {
  const names = [...new Set([...FORMATS.number, ...FORMATS.datetime, ...FORMATS.yesno, '>', '<'])];
  if (current && !names.includes(current)) names.push(current);
  return names.map(x => ({ value: x, label: formatLabel(x) }));
}

export class Designer extends BaseView {
  constructor(app, tab, live = false) {
    super(app, tab);
    this.live = live;
    this.sel = new Set();
    this.target = 'detail';
    this.insert = null;
    this.propTab = 'all';
    this.liveIndex = 0;
    const panes = readPanes();
    this.showProps = panes.props ?? !app.isPhone();
    this.showFields = !!panes.fields && !app.isPhone();
    this.showGroups = panes.groups ?? true;
    this.allTables = false;
  }

  get ro() { return this.app.readOnly; }
  get isReport() { return this.tab.kind === 'reports'; }

  // ---- the sections ------------------------------------------------------------------------------------------------------------------

  /** The sections in the order Access shows them: [{ key, label, sec }]. */
  sections() {
    const o = this.obj;
    const s = o.sections;
    if (!this.isReport) {
      return [
        { key: 'header', label: t('Base.Sec.FormHeader'), sec: s.header },
        { key: 'detail', label: t('Base.Sec.Detail'), sec: s.detail },
        { key: 'footer', label: t('Base.Sec.FormFooter'), sec: s.footer },
      ].filter(x => x.sec);
    }
    const g = o.grouping || [];
    s.groupHeaders ||= [];
    s.groupFooters ||= [];
    const heads = g.map((x, i) => ({ key: 'gh' + i, label: t('Base.Sec.GroupHeader', x.field), sec: s.groupHeaders[i], off: x.header === false }));
    const foots = g.map((x, i) => ({ key: 'gf' + i, label: t('Base.Sec.GroupFooter', x.field), sec: s.groupFooters[i], off: x.footer === false })).reverse();
    return [
      { key: 'reportHeader', label: t('Base.Sec.ReportHeader'), sec: s.reportHeader },
      { key: 'pageHeader', label: t('Base.Sec.PageHeader'), sec: s.pageHeader },
      ...heads,
      { key: 'detail', label: t('Base.Sec.Detail'), sec: s.detail },
      ...foots,
      { key: 'pageFooter', label: t('Base.Sec.PageFooter'), sec: s.pageFooter },
      { key: 'reportFooter', label: t('Base.Sec.ReportFooter'), sec: s.reportFooter },
    ].filter(x => x.sec && !x.off);
  }

  sectionOf(key) { return this.sections().find(x => x.key === key) || null; }

  allControls() { return controlsOfObject(this.obj); }

  find(id) { return this.allControls().find(x => x.c.id === id) || null; }

  selected() { return [...this.sel].map(id => this.find(id)).filter(Boolean); }

  selectedControls() { return this.selected().map(x => x.c); }

  nodeOf(id) { return this.paper.querySelector(`.mo-base-dc[data-id="${CSS.escape(id)}"]`); }

  secNode(key) { return this.paper.querySelector(`[data-sec="${CSS.escape(key)}"]`); }

  /** The section under a point of the screen (null outside the paper). */
  sectionAt(x, y) {
    const list = [...this.paper.querySelectorAll('[data-sec]')];
    const paper = this.paper.getBoundingClientRect();
    if (x < paper.left - 40 || x > paper.right + 40) return null;
    for (const n of list) {
      const r = n.getBoundingClientRect();
      if (y >= r.top && y < Math.max(r.bottom, r.top + 6)) return n.dataset.sec;
    }
    for (const b of this.paper.querySelectorAll('[data-bar]')) {
      const r = b.getBoundingClientRect();
      if (y >= r.top && y < r.bottom) return b.dataset.bar;
    }
    return null;
  }

  /** The sections grow to hold their controls (Access does the same when a control is dragged below the edge). */
  growSections() {
    for (const s of this.sections()) {
      const bottom = Math.max(0, ...s.sec.controls.map(c => c.y + c.h));
      if (bottom > s.sec.height) s.sec.height = bottom;
    }
    const right = Math.max(0, ...this.allControls().map(x => x.c.x + x.c.w));
    if (right > this.obj.width) this.obj.width = Math.min(4000, right);
  }

  // ---- mount and render ----------------------------------------------------------------------------------------------------------

  mount() {
    this.canvas = el('div', { class: 'mo-base-dz__canvas', tabindex: '0', role: 'application', 'aria-roledescription': t('Base.Dz.Designer'), 'aria-label': this.obj.name });
    this.paper = el('div', { class: 'mo-base-dz__paper' });
    this.canvas.append(this.paper);
    this.bottom = el('div', { class: 'mo-base-dz__bottom' });
    this.side = el('aside', { class: 'mo-base-dz__side', 'aria-label': t('Base.Dz.Panes') });
    this.root = el('div', { class: 'mo-base-dz' + (this.live ? ' is-live' : ' is-design') + (this.isReport ? ' is-report' : '') },
      el('div', { class: 'mo-base-dz__main' }, this.canvas, this.bottom), this.side);
    this.el.append(this.root);
    this.paper.addEventListener('pointerdown', e => this.onDown(e));
    this.paper.addEventListener('dblclick', e => this.onDouble(e));
    this.paper.addEventListener('contextmenu', e => this.onContext(e));
    this.canvas.addEventListener('pointerdown', e => { if (e.target === this.canvas) { this.finishInline(true); this.select([], 'object'); } });
    this.canvas.addEventListener('keydown', e => this.onKey(e));
    if (this.options?.created && !this.ro) this.showFields = true;
    if (this.live) this.loadLive();
    this.render();
  }

  destroy() { this.finishInline(false); }

  refresh() {
    if (!this.obj) return;
    for (const id of [...this.sel]) if (!this.find(id)) this.sel.delete(id);
    if (this.target !== 'object' && !this.sectionOf(this.target)) this.target = 'detail';
    this.fieldCache = null;
    if (this.live) this.loadLive();
    this.render();
  }

  focus() { requestAnimationFrame(() => this.canvas?.focus({ preventScroll: true })); }

  render() {
    const o = this.obj;
    if (!o) return;
    this.paper.style.width = o.width + 'px';
    const parts = [];
    for (const s of this.sections()) {
      if (this.live && s.sec.height === 0 && !s.sec.controls.length) continue;
      if (!this.live) parts.push(this.bar(s));
      parts.push(this.band(s));
    }
    if (!this.ro) parts.push(el('div', { class: 'mo-base-dz__wsizer', title: t('Base.Dz.FormWidth') }));
    this.paper.replaceChildren(...parts);
    this.canvas.classList.toggle('is-inserting', !!this.insert);
    this.renderSide();
    this.renderBottom();
    this.app.ribbonRefresh();
    if (this.app.active?.viewObj === this) this.app.setStatus();
  }

  bar(s) {
    return el('div', { class: 'mo-base-dz__bar' + (!this.sel.size && this.target === s.key ? ' is-target' : ''), 'data-bar': s.key }, icon('chevron-down'), el('span', { text: s.label }));
  }

  band(s) {
    const node = el('div', { class: 'mo-base-dz__sec' + (s.sec.visible === false ? ' is-hiddensec' : ''), 'data-sec': s.key });
    node.style.height = s.sec.height + 'px';
    if (s.sec.back) node.style.backgroundColor = s.sec.back;
    for (const c of s.sec.controls) node.append(this.controlNode(c, s.key));
    return el('div', { class: 'mo-base-dz__band' }, node, this.ro ? null : el('div', { class: 'mo-base-dz__sizer', 'data-sizer': s.key, title: t('Base.Dz.SectionHeight') }));
  }

  controlNode(c, key) {
    const n = el('div', { class: 'mo-base-dc mo-base-dc--' + c.type + (this.sel.has(c.id) ? ' is-selected' : ''), 'data-id': c.id, 'data-key': key, title: c.name || controlTypeName(c.type) });
    styleControl(n, c);
    if (c.type === 'line') n.classList.toggle('is-vertical', c.h > c.w);
    n.append(...this.controlContent(c, key));
    if (this.sel.has(c.id) && !this.ro) n.append(...HANDLES.map(h => el('span', { class: 'mo-base-dc__h mo-base-dc__h--' + h, 'data-h': h })));
    return n;
  }

  /** What a control shows: its source in Design view, the data of the record in Layout view. */
  controlContent(c, key) {
    const text = (s, muted = false) => el('span', { class: 'mo-base-dc__text' + (muted ? ' is-muted' : ''), text: s });
    switch (c.type) {
      case 'label': return [text(c.caption || '')];
      case 'button': return [el('span', { class: 'mo-base-dc__button', text: c.caption || actionName(c.action) || controlTypeName('button') })];
      case 'checkbox': {
        const on = this.live ? this.liveValue(c, key) === true : false;
        return [el('span', { class: 'mo-base-dc__box' + (on ? ' is-on' : '') }, on ? icon('check') : null)];
      }
      case 'textbox': case 'combobox': {
        const out = this.live ? [text(this.liveText(c, key))] : [text(c.source || t('Base.Dz.Unbound'), !c.source)];
        if (c.type === 'combobox') out.push(el('span', { class: 'mo-base-dc__arrow' }, icon('chevron-down')));
        return out;
      }
      case 'subform': return [this.subformContent(c)];
      case 'line': return [el('span', { class: 'mo-base-dc__line' })];
      default: return [];
    }
  }

  subformContent(c) {
    const src = c.subform?.source;
    const name = tableById(this.db, src)?.name || queryById(this.db, src)?.name || t('Base.Dz.Unbound');
    const box = el('div', { class: 'mo-base-dc__sub' }, el('div', { class: 'mo-base-dc__subhead' }, icon(tableById(this.db, src) ? KIND_ICON.tables : KIND_ICON.queries), el('span', { text: name })));
    if (this.live && this.liveRs) {
      const row = this.liveRow();
      const masterCol = row && c.subform?.linkMaster ? this.liveEv.column(c.subform.linkMaster) : null;
      const master = row && masterCol ? this.liveRs.get(row, masterCol) : null;
      let rs = null;
      try { rs = subformRecordset(this.app, c.subform, master); } catch { rs = null; }
      if (rs && rs.columns.length) {
        const cols = rs.columns.slice(0, 6);
        const table = el('table', { class: 'mo-base-dc__subtable' },
          el('thead', {}, el('tr', {}, ...cols.map(col => el('th', { text: col.caption })))),
          el('tbody', {}, ...rs.rows.slice(0, 8).map(r => el('tr', {}, ...cols.map(col => el('td', { text: rs.display(col, rs.get(r, col)) }))))));
        box.append(table);
      }
    }
    return box;
  }

  // ---- the data of Layout view -------------------------------------------------------------------------------------------------------

  loadLive() {
    this.liveRs = null;
    this.liveEv = null;
    const o = this.obj;
    if (!o?.source) return;
    try { this.liveRs = recordsetOf(this.app, o.source, {}); } catch { this.liveRs = null; }
    if (!this.liveRs) return;
    this.liveEv = new ControlEval(this.liveRs, { lang: this.app.lang, currency: this.db.settings.currency, domain: domainFunction(this.db, this.app.lang) });
    this.liveIndex = Math.min(this.liveIndex, Math.max(0, this.liveRs.rows.length - 1));
  }

  liveRow() { return this.liveRs?.rows[this.liveIndex] || null; }

  liveText(c, key) {
    if (!this.liveEv) return c.source && c.source.startsWith('=') ? '' : '';
    const row = key === 'detail' ? this.liveRow() : this.liveRow();
    return this.liveEv.text(c, row, this.liveRs.rows);
  }

  liveValue(c) {
    if (!this.liveEv || !c.source) return null;
    try { return this.liveEv.value(c.source, this.liveRow(), this.liveRs.rows); } catch { return null; }
  }

  renderBottom() {
    const parts = [];
    if (this.live && this.liveRs) {
      const n = this.liveRs.rows.length;
      const go = i => { this.liveIndex = Math.max(0, Math.min(n - 1, i)); this.render(); };
      parts.push(el('div', { class: 'mo-base-ds__nav mo-base-dz__nav' },
        el('span', { class: 'mo-base-ds__navlabel', text: t('Base.Nav.Record') }),
        iconButton('base-first', t('Base.Nav.First'), () => go(0), { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: this.liveIndex <= 0 ? true : null }),
        iconButton('chevron-left', t('Base.Nav.Previous'), () => go(this.liveIndex - 1), { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: this.liveIndex <= 0 ? true : null }),
        el('span', { class: 'mo-base-ds__navof', text: n ? t('Base.Nav.NOfM', this.liveIndex + 1, n) : t('Base.Nav.NoRecords') }),
        iconButton('chevron-right', t('Base.Nav.Next'), () => go(this.liveIndex + 1), { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: this.liveIndex >= n - 1 ? true : null }),
        iconButton('base-last', t('Base.Nav.Last'), () => go(n - 1), { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: this.liveIndex >= n - 1 ? true : null })));
    }
    if (this.isReport && this.showGroups) parts.push(this.groupPane());
    this.bottom.replaceChildren(...parts);
    this.bottom.hidden = !parts.length;
  }

  /** Reports draw the Group, Sort and Total pane here (see reports.js). */
  groupPane() { return el('div'); }

  // ---- selection -------------------------------------------------------------------------------------------------------------------

  select(ids, target = this.target) {
    this.sel = new Set(ids);
    if (target) this.target = target;
    this.paintSelection();
  }

  paintSelection() {
    for (const n of this.paper.querySelectorAll('.mo-base-dc')) {
      const on = this.sel.has(n.dataset.id);
      n.classList.toggle('is-selected', on);
      n.querySelectorAll('.mo-base-dc__h').forEach(h => h.remove());
      if (on && !this.ro) n.append(...HANDLES.map(h => el('span', { class: 'mo-base-dc__h mo-base-dc__h--' + h, 'data-h': h })));
    }
    for (const b of this.paper.querySelectorAll('[data-bar]')) b.classList.toggle('is-target', !this.sel.size && this.target === b.dataset.bar);
    this.renderProps();
    this.app.ribbonRefresh();
    this.app.setStatus();
  }

  selectAll() {
    const key = this.target && this.target !== 'object' && !this.sel.size ? this.target : null;
    this.select(this.allControls().filter(x => !key || x.key === key).map(x => x.c.id));
  }

  /** Tab moves the selection to the next control (in the order of the sections, then top to bottom, left to right). */
  selectNext(back) {
    const list = this.sections().flatMap(s => [...s.sec.controls].sort((a, b) => a.y - b.y || a.x - b.x).map(c => ({ c, key: s.key })));
    if (!list.length) return;
    const cur = list.findIndex(x => this.sel.has(x.c.id));
    const next = list[(cur + (back ? -1 : 1) + list.length) % list.length];
    this.select([next.c.id], next.key);
    this.nodeOf(next.c.id)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  // ---- pointer -----------------------------------------------------------------------------------------------------------------------

  onDown(e) {
    if (e.button !== 0) return;
    if (this.inline && e.target === this.inline) return;
    this.finishInline(true);
    const handle = e.target.closest('.mo-base-dc__h');
    const ctl = e.target.closest('.mo-base-dc');
    const sizer = e.target.closest('[data-sizer]');
    const wsizer = e.target.closest('.mo-base-dz__wsizer');
    const bar = e.target.closest('[data-bar]');
    const sec = e.target.closest('[data-sec]');
    this.canvas.focus({ preventScroll: true });
    if (handle && ctl) { e.preventDefault(); this.startResize(e, ctl.dataset.id, handle.dataset.h); return; }
    if (sizer) { e.preventDefault(); this.startSectionResize(e, sizer.dataset.sizer); return; }
    if (wsizer) { e.preventDefault(); this.startWidthResize(e); return; }
    if (bar) { this.select([], bar.dataset.bar); return; }
    if (this.insert && sec) { e.preventDefault(); this.startInsert(e, sec.dataset.sec); return; }
    if (ctl) { e.preventDefault(); this.startMove(e, ctl.dataset.id); return; }
    if (sec) { e.preventDefault(); this.startFrame(e, sec.dataset.sec); }
  }

  startMove(e, id) {
    const hit = this.find(id);
    if (!hit) return;
    if (e.shiftKey || e.ctrlKey || e.metaKey) {
      if (this.sel.has(id)) this.sel.delete(id); else this.sel.add(id);
      this.target = hit.key;
      this.paintSelection();
      return;
    }
    if (!this.sel.has(id)) { this.sel = new Set([id]); this.target = hit.key; this.paintSelection(); }
    if (this.ro) return;
    const items = this.selected();
    const start = items.map(x => ({ x: x.c.x, y: x.c.y }));
    let before = null;
    drag(e, (dx, dy, ev) => {
      if (!before) {
        if (Math.abs(dx) + Math.abs(dy) < 3) return;
        before = this.app.beginDesign();
        this.root.classList.add('is-dragging');
      }
      items.forEach((x, i) => {
        x.c.x = Math.max(0, snap(start[i].x + dx, !ev.altKey));
        x.c.y = snap(start[i].y + dy, !ev.altKey);
        const n = this.nodeOf(x.c.id);
        if (n) { n.style.left = x.c.x + 'px'; n.style.top = x.c.y + 'px'; }
      });
    }, (dx, dy, ev) => {
      this.root.classList.remove('is-dragging');
      if (!before) return;
      this.moveAcross(items, ev);
      this.growSections();
      this.app.commitDesign(before);
      this.render();
    });
  }

  /** Controls dropped on another section move into it (all of the selection must come from one section). */
  moveAcross(items, ev) {
    const keys = new Set(items.map(x => x.key));
    const target = this.sectionAt(ev.clientX, ev.clientY);
    if (target && keys.size === 1 && !keys.has(target)) {
      const from = [...keys][0];
      const dy = this.secNode(from).getBoundingClientRect().top - this.secNode(target).getBoundingClientRect().top;
      const src = this.sectionOf(from).sec;
      const dst = this.sectionOf(target).sec;
      for (const x of items) {
        src.controls.splice(src.controls.indexOf(x.c), 1);
        x.c.y = snap(x.c.y + dy);
        dst.controls.push(x.c);
      }
      this.target = target;
    }
    for (const x of items) { x.c.x = Math.max(0, x.c.x); x.c.y = Math.max(0, x.c.y); }
  }

  startResize(e, id, h) {
    if (this.ro) return;
    if (!this.sel.has(id)) { this.sel = new Set([id]); this.paintSelection(); }
    const items = this.selected();
    const start = items.map(x => ({ x: x.c.x, y: x.c.y, w: x.c.w, h: x.c.h }));
    let before = null;
    drag(e, (dx, dy, ev) => {
      if (!before) before = this.app.beginDesign();
      items.forEach((it, i) => {
        const s = start[i];
        const c = it.c;
        const min = c.type === 'line' ? 1 : 8;
        const on = !ev.altKey;
        let { x, y, w, h: hh } = s;
        if (h.includes('e')) w = Math.max(min, snap(s.w + dx, on));
        if (h.includes('s')) hh = Math.max(min, snap(s.h + dy, on));
        if (h.includes('w')) { x = Math.max(0, Math.min(snap(s.x + dx, on), s.x + s.w - min)); w = s.w + (s.x - x); }
        if (h.includes('n')) { y = Math.max(0, Math.min(snap(s.y + dy, on), s.y + s.h - min)); hh = s.h + (s.y - y); }
        Object.assign(c, { x, y, w, h: hh });
        const n = this.nodeOf(c.id);
        if (n) Object.assign(n.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px' });
      });
    }, () => {
      if (!before) return;
      this.growSections();
      this.app.commitDesign(before);
      this.render();
    });
  }

  startSectionResize(e, key) {
    if (this.ro) return;
    const s = this.sectionOf(key)?.sec;
    if (!s) return;
    const h0 = s.height;
    const min = Math.max(0, ...s.controls.map(c => c.y + c.h));
    const node = this.secNode(key);
    let before = null;
    drag(e, (dx, dy) => {
      if (!before) before = this.app.beginDesign();
      s.height = Math.min(10000, Math.max(min, snap(h0 + dy)));
      node.style.height = s.height + 'px';
    }, () => { if (before) this.app.commitDesign(before); this.render(); });
  }

  startWidthResize(e) {
    if (this.ro) return;
    const o = this.obj;
    const w0 = o.width;
    const min = Math.max(100, ...this.allControls().map(x => x.c.x + x.c.w));
    let before = null;
    drag(e, dx => {
      if (!before) before = this.app.beginDesign();
      o.width = Math.min(4000, Math.max(min, snap(w0 + dx)));
      this.paper.style.width = o.width + 'px';
    }, () => { if (before) this.app.commitDesign(before); this.render(); });
  }

  /** A frame drawn on a section selects the controls it touches; a click selects the section. */
  startFrame(e, key) {
    const node = this.secNode(key);
    const r = node.getBoundingClientRect();
    const x0 = e.clientX - r.left;
    const y0 = e.clientY - r.top;
    let frame = null;
    drag(e, (dx, dy) => {
      if (!frame) {
        if (Math.abs(dx) + Math.abs(dy) < 4) return;
        frame = el('div', { class: 'mo-base-dz__frame' });
        node.append(frame);
      }
      Object.assign(frame.style, { left: Math.min(x0, x0 + dx) + 'px', top: Math.min(y0, y0 + dy) + 'px', width: Math.abs(dx) + 'px', height: Math.abs(dy) + 'px' });
    }, (dx, dy, ev) => {
      if (!frame) { this.select(ev.shiftKey ? [...this.sel] : [], key); return; }
      frame.remove();
      const box = { x: Math.min(x0, x0 + dx), y: Math.min(y0, y0 + dy), w: Math.abs(dx), h: Math.abs(dy) };
      const hit = this.sectionOf(key).sec.controls.filter(c => c.x < box.x + box.w && c.x + c.w > box.x && c.y < box.y + box.h && c.y + c.h > box.y).map(c => c.id);
      this.select(ev.shiftKey ? [...new Set([...this.sel, ...hit])] : hit, key);
    });
  }

  // ---- new controls ------------------------------------------------------------------------------------------------------------------

  /** The ribbon's control buttons: the next click (or frame) on a section places the control. */
  insertMode(type) {
    if (this.ro) return;
    this.insert = this.insert === type ? null : type;
    this.canvas.classList.toggle('is-inserting', !!this.insert);
    this.canvas.focus({ preventScroll: true });
    this.app.ribbonRefresh();
    if (this.insert) this.app.toast(t('Base.Dz.InsertHint', controlTypeName(type)));
  }

  startInsert(e, key) {
    const type = this.insert;
    const node = this.secNode(key);
    const r = node.getBoundingClientRect();
    const x0 = snap(e.clientX - r.left);
    const y0 = snap(e.clientY - r.top);
    let ghost = null;
    drag(e, (dx, dy) => {
      if (!ghost) { ghost = el('div', { class: 'mo-base-dz__frame is-insert' }); node.append(ghost); }
      Object.assign(ghost.style, { left: Math.min(x0, x0 + dx) + 'px', top: Math.min(y0, y0 + dy) + 'px', width: Math.abs(dx) + 'px', height: Math.abs(dy) + 'px' });
    }, (dx, dy) => {
      ghost?.remove();
      const size = SIZE[type];
      const big = Math.abs(dx) > 8 || Math.abs(dy) > 8;
      const box = big
        ? { x: Math.min(x0, x0 + snap(dx)), y: Math.min(y0, y0 + snap(dy)), w: Math.max(type === 'line' ? Math.abs(snap(dx)) || 1 : 12, Math.abs(snap(dx))), h: type === 'line' && Math.abs(dy) < 8 ? 1 : Math.max(type === 'line' ? 1 : 12, Math.abs(snap(dy))) }
        : { x: x0, y: y0, w: size.w, h: size.h };
      this.insert = null;
      this.canvas.classList.remove('is-inserting');
      this.addControl(type, key, box);
    });
  }

  controlName(type) {
    const used = new Set(this.allControls().map(x => lower(x.c.name)));
    for (let i = 1; ; i++) if (!used.has(lower(NAME_PREFIX[type] + i))) return NAME_PREFIX[type] + i;
  }

  async addControl(type, key, box) {
    let more = {};
    if (type === 'button') { more = await this.buttonWizard(); if (!more) return; }
    if (type === 'subform') { more = await this.subformWizard(); if (!more) return; }
    if (type === 'label') more.caption = t('Base.Dz.NewLabel');
    if (this.isReport && type === 'textbox') more.border = false;
    const c = this.app.design(() => {
      const ctl = newControl(type, newId(this.allControls().map(x => x.c), 'c'), box, { name: this.controlName(type), ...more });
      this.sectionOf(key).sec.controls.push(ctl);
      this.growSections();
      return ctl;
    });
    if (!c) return;
    this.sel = new Set([c.id]);
    this.target = key;
    this.render();
    if (type === 'label') this.editInline(c.id);
    else if (type === 'textbox' || type === 'combobox' || type === 'checkbox') { this.showProps = true; this.renderSide(); this.focusProp('source'); }
  }

  /** The Command Button Wizard of Access, short: what the button does and its caption. */
  async buttonWizard(current = null) {
    const action = selectInput(ACTIONS.map(a => ({ value: a, label: actionName(a) })), current?.action || 'next');
    const caption = textInput(current?.caption || actionName(current?.action || 'next'), { maxlength: 255 });
    let touched = !!current;
    caption.addEventListener('input', () => { touched = true; });
    action.addEventListener('change', () => { if (!touched) caption.value = actionName(action.value); });
    const r = await dialog({
      title: t('Base.Dz.ButtonWizard'), icon: 'base-c-button', body: el('div', { class: 'mo-base-form' }, field(t('Base.Dz.ButtonAction'), action), field(t('Base.Prop.Caption'), caption)),
      onOpen: () => action.focus(), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    return r === 'ok' ? { action: action.value, caption: caption.value.trim() || actionName(action.value) } : null;
  }

  /** The Subform Wizard, short: the table or query of the subform and the fields that link it to the main form. */
  async subformWizard(current = null) {
    const db = this.db;
    const own = this.obj.source;
    const options = [...db.tables.filter(x => x.id !== own).map(x => ({ value: x.id, label: x.name })), ...db.queries.filter(x => x.id !== own).map(x => ({ value: x.id, label: x.name }))];
    if (!options.length) { this.app.toast(t('Base.Dz.NoSubformSource'), 'info'); return null; }
    const guess = () => {
      const rel = db.relationships.find(r => r.parent.table === own && r.child.table === source.value);
      if (!rel) return null;
      return { master: fieldById(tableById(db, rel.parent.table), rel.parent.field)?.name || '', child: fieldById(tableById(db, rel.child.table), rel.child.field)?.name || '' };
    };
    const firstRelated = db.relationships.find(r => r.parent.table === own)?.child.table;
    const source = selectInput(options, current?.source || firstRelated || options[0].value);
    const master = textInput(current?.linkMaster || '', { list: 'mo-base-sf-master' });
    const child = textInput(current?.linkChild || '', { list: 'mo-base-sf-child' });
    const masterList = el('datalist', { id: 'mo-base-sf-master' }, ...this.fieldList().map(f => el('option', { value: f.name })));
    const childList = el('datalist', { id: 'mo-base-sf-child' });
    const fill = () => {
      childList.replaceChildren(...sourceFieldList(db, source.value, this.app.lang).map(f => el('option', { value: f.name })));
      const g = guess();
      if (g && !current) { master.value = g.master; child.value = g.child; }
    };
    source.addEventListener('change', fill);
    fill();
    const r = await dialog({
      title: t('Base.Dz.SubformWizard'), icon: 'base-c-subform', size: 'md',
      body: el('div', { class: 'mo-base-form' }, field(t('Base.Prop.SourceObject'), source), field(t('Base.Prop.LinkMaster'), master, t('Base.Dz.LinkHint')), field(t('Base.Prop.LinkChild'), child), masterList, childList),
      onOpen: () => source.focus(), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return null;
    const name = options.find(o => o.value === source.value)?.label || '';
    return { subform: { source: source.value, linkMaster: master.value.trim(), linkChild: child.value.trim() }, name: name.slice(0, 64) };
  }

  // ---- fields of the record source ---------------------------------------------------------------------------------------------------

  fieldList() {
    if (!this.fieldCache || this.fieldCache.source !== this.obj.source) this.fieldCache = { source: this.obj.source, list: sourceFieldList(this.db, this.obj.source, this.app.lang) };
    return this.fieldCache.list;
  }

  /** Puts fields on a section: a label and a bound control for each (a check box with its label right of it), like dragging from the Field List. */
  addFields(fields, key, at = null, tableId = null) {
    if (this.ro || !fields.length) return;
    const ids = [];
    this.app.design(() => {
      const o = this.obj;
      if (!o.source && tableId) o.source = tableId;
      const sec = this.sectionOf(key)?.sec || this.sectionOf('detail').sec;
      let y = at ? snap(at.y) : sec.controls.reduce((m, c) => Math.max(m, c.y + c.h + 8), 16);
      const x = at ? Math.max(0, snap(at.x)) : 24;
      const all = () => this.allControls().map(v => v.c);
      for (const f of fields) {
        const caption = f.caption || f.name;
        if (f.type === 'yesno') {
          const box = newControl('checkbox', newId(all(), 'c'), { x: x + 4, y: y + 3, w: 18, h: 18 }, { source: f.name, name: f.name });
          sec.controls.push(box);
          const label = newControl('label', newId(all(), 'c'), { x: x + 30, y, w: 150, h: 24 }, { caption, name: 'Label_' + f.name.slice(0, 50) });
          sec.controls.push(label);
          ids.push(box.id, label.id);
          y += 32;
          continue;
        }
        const h = f.type === 'memo' ? 72 : 24;
        const label = newControl('label', newId(all(), 'c'), { x, y, w: 150, h: 24 }, { caption, name: 'Label_' + f.name.slice(0, 50) });
        sec.controls.push(label);
        const type = f.lookup ? 'combobox' : 'textbox';
        const ctl = newControl(type, newId(all(), 'c'), { x: x + 160, y, w: Math.min(controlWidth(f), 360), h }, {
          source: f.name, name: f.name, align: ['number', 'currency', 'autonumber'].includes(f.type) && !f.lookup ? 'right' : 'left', locked: f.type === 'autonumber', border: this.isReport ? false : true,
        });
        sec.controls.push(ctl);
        ids.push(label.id, ctl.id);
        y += h + 8;
      }
      this.growSections();
    });
    this.sel = new Set(ids);
    this.target = key;
    this.fieldCache = null;
    if (this.live) this.loadLive();
    this.render();
  }

  // ---- changes of the selection ---------------------------------------------------------------------------------------------------

  /** Sets a property on the selected controls (one undo step). */
  setControls(key, value, list = this.selectedControls()) {
    if (this.ro || !list.length) return;
    this.app.design(() => { for (const c of list) c[key] = value; });
    this.render();
  }

  deleteSelected() {
    if (this.ro || !this.sel.size) return;
    const ids = new Set(this.sel);
    this.app.design(() => { for (const s of this.sections()) s.sec.controls = s.sec.controls.filter(c => !ids.has(c.id)); });
    this.sel.clear();
    this.render();
  }

  copy() {
    const items = this.selected();
    if (!items.length) return false;
    clipboard = items.map(x => structuredClone(x.c));
    return true;
  }

  cut() { if (this.copy()) this.deleteSelected(); }

  paste() {
    if (this.ro || !clipboard?.length) return;
    const key = this.selected()[0]?.key || (this.target !== 'object' && this.sectionOf(this.target) ? this.target : 'detail');
    const ids = [];
    this.app.design(() => {
      const sec = this.sectionOf(key).sec;
      for (const src of clipboard) {
        const c = structuredClone(src);
        c.id = newId(this.allControls().map(x => x.c), 'c');
        if (sec.controls.some(x => x.x === c.x && x.y === c.y)) { c.x += 8; c.y += 8; }
        if (c.name && this.allControls().some(x => lower(x.c.name) === lower(c.name))) c.name = this.controlName(c.type);
        sec.controls.push(c);
        ids.push(c.id);
      }
      this.growSections();
    });
    this.sel = new Set(ids);
    this.target = key;
    this.render();
  }

  duplicate() { if (this.copy()) this.paste(); }

  /** Align Left/Right/Top/Bottom, To Grid. */
  align(which) {
    const cs = this.selectedControls();
    if (this.ro || !cs.length || (which !== 'grid' && cs.length < 2)) return;
    this.app.design(() => {
      const left = Math.min(...cs.map(c => c.x));
      const right = Math.max(...cs.map(c => c.x + c.w));
      const top = Math.min(...cs.map(c => c.y));
      const bottom = Math.max(...cs.map(c => c.y + c.h));
      for (const c of cs) {
        if (which === 'left') c.x = left;
        else if (which === 'right') c.x = right - c.w;
        else if (which === 'top') c.y = top;
        else if (which === 'bottom') c.y = bottom - c.h;
        else { c.x = snap(c.x); c.y = snap(c.y); c.w = Math.max(1, snap(c.w)); c.h = Math.max(1, c.type === 'line' ? c.h : snap(c.h)); }
      }
    });
    this.render();
  }

  /** Size To Tallest/Shortest/Widest/Narrowest. */
  sizeTo(which) {
    const cs = this.selectedControls();
    if (this.ro || cs.length < 2) return;
    this.app.design(() => {
      const hs = cs.map(c => c.h);
      const ws = cs.map(c => c.w);
      for (const c of cs) {
        if (which === 'tallest') c.h = Math.max(...hs);
        else if (which === 'shortest') c.h = Math.min(...hs);
        else if (which === 'widest') c.w = Math.max(...ws);
        else c.w = Math.min(...ws);
      }
      this.growSections();
    });
    this.render();
  }

  /** Bring to Front / Send to Back: the order of the controls in their section. */
  order(front) {
    const items = this.selected();
    if (this.ro || !items.length) return;
    this.app.design(() => {
      for (const x of items) {
        const list = x.sec.controls;
        list.splice(list.indexOf(x.c), 1);
        if (front) list.push(x.c); else list.unshift(x.c);
      }
    });
    this.render();
  }

  // ---- keys ------------------------------------------------------------------------------------------------------------------------

  onKey(e) {
    if (isTyping(e.target)) return;
    const mod = e.ctrlKey || e.metaKey;
    const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    if (arrows[e.key] && this.sel.size) {
      e.preventDefault();
      if (this.ro) return;
      const [dx, dy] = arrows[e.key];
      this.keyBefore ||= this.app.beginDesign();
      for (const c of this.selectedControls()) {
        if (e.shiftKey) { c.w = Math.max(1, c.w + dx); c.h = Math.max(1, c.h + dy); }
        else { c.x = Math.max(0, c.x + dx); c.y = Math.max(0, c.y + dy); }
        const n = this.nodeOf(c.id);
        if (n) Object.assign(n.style, { left: c.x + 'px', top: c.y + 'px', width: c.w + 'px', height: c.h + 'px' });
      }
      clearTimeout(this.keyTimer);
      this.keyTimer = setTimeout(() => { const before = this.keyBefore; this.keyBefore = null; this.growSections(); this.app.commitDesign(before); this.render(); }, 500);
      return;
    }
    switch (e.key) {
      case 'Delete': case 'Backspace': if (this.sel.size) { e.preventDefault(); this.deleteSelected(); } return;
      case 'Escape': e.preventDefault(); if (this.insert) this.insertMode(this.insert); else this.select([]); return;
      case 'Tab': if (!this.sel.size) return; e.preventDefault(); this.selectNext(e.shiftKey); return;
      case 'F6': if (this.propsBox) { e.preventDefault(); this.propsBox.querySelector('select, input')?.focus(); } return;
      case 'F2': case 'Enter': {
        const one = this.selectedControls();
        if (one.length === 1 && (one[0].type === 'label' || one[0].type === 'button')) { e.preventDefault(); this.editInline(one[0].id); }
        else if (e.key === 'Enter' && e.altKey) { e.preventDefault(); this.toggleProps(true); }
        return;
      }
      case 'F4': e.preventDefault(); this.toggleProps(); return;
      default: break;
    }
    if (!mod) return;
    const k = e.key.toLowerCase();
    if (k === 'a') { e.preventDefault(); this.selectAll(); }
    else if (k === 'c') { e.preventDefault(); this.copy(); }
    else if (k === 'x') { e.preventDefault(); this.cut(); }
    else if (k === 'v') { e.preventDefault(); this.paste(); }
    else if (k === 'd') { e.preventDefault(); this.duplicate(); }
  }

  onDouble(e) {
    const ctl = e.target.closest('.mo-base-dc');
    if (!ctl) { const bar = e.target.closest('[data-bar], [data-sec]'); if (bar) this.toggleProps(true); return; }
    const hit = this.find(ctl.dataset.id);
    if (!hit) return;
    if (hit.c.type === 'label' || hit.c.type === 'button') this.editInline(hit.c.id);
    else { this.toggleProps(true); this.focusProp(hit.c.type === 'subform' ? 'subform.source' : 'source'); }
  }

  onContext(e) {
    e.preventDefault();
    const ctl = e.target.closest('.mo-base-dc');
    const sec = e.target.closest('[data-sec], [data-bar]');
    if (ctl && !this.sel.has(ctl.dataset.id)) this.select([ctl.dataset.id], this.find(ctl.dataset.id)?.key);
    else if (!ctl && sec) this.select([], sec.dataset.sec || sec.dataset.bar);
    const ro = this.ro;
    const some = this.sel.size > 0;
    const many = this.sel.size > 1;
    contextMenu({ x: e.clientX, y: e.clientY }, [
      { label: t('Base.Cut'), icon: 'base-cut', run: () => this.cut(), disabled: ro || !some },
      { label: t('Base.Copy'), icon: 'copy', run: () => this.copy(), disabled: !some },
      { label: t('Base.Paste'), icon: 'base-paste', run: () => this.paste(), disabled: ro || !clipboard?.length },
      { label: t('Base.Dz.Delete'), icon: 'trash', danger: true, run: () => this.deleteSelected(), disabled: ro || !some },
      '-',
      { label: t('Base.Dz.Front'), icon: 'base-front', run: () => this.order(true), disabled: ro || !some },
      { label: t('Base.Dz.Back'), icon: 'base-back', run: () => this.order(false), disabled: ro || !some },
      '-',
      { label: t('Base.Dz.AlignLeft'), icon: 'base-align-left', run: () => this.align('left'), disabled: ro || !many },
      { label: t('Base.Dz.AlignTop'), icon: 'base-align-top', run: () => this.align('top'), disabled: ro || !many },
      { label: t('Base.Dz.SizeWidest'), icon: 'base-size-widest', run: () => this.sizeTo('widest'), disabled: ro || !many },
      '-',
      { label: t('Base.Dz.SelectAll'), icon: 'base-select', run: () => this.selectAll() },
      { label: t('Base.Dz.Properties'), icon: 'sliders', run: () => this.toggleProps(true) },
    ]);
  }

  // ---- captions in place -------------------------------------------------------------------------------------------------------------

  editInline(id) {
    const hit = this.find(id);
    if (!hit || this.ro || !(hit.c.type === 'label' || hit.c.type === 'button')) return;
    this.finishInline(true);
    this.select([id], hit.key);
    const n = this.nodeOf(id);
    if (!n) return;
    const input = el('input', { class: 'mo-base-dz__inline', value: hit.c.caption || '', 'aria-label': t('Base.Prop.Caption'), maxlength: 1000 });
    n.append(input);
    this.inline = input;
    this.inlineDone = commit => {
      this.inline = null;
      this.inlineDone = null;
      const value = input.value;
      if (commit && value !== (hit.c.caption || '')) {
        this.app.design(() => {
          hit.c.caption = value;
          // a label grows to its text, like Access does it while the text is typed
          if (hit.c.type === 'label') {
            const g = (this.measure ||= document.createElement('canvas').getContext('2d'));
            g.font = `${hit.c.italic ? 'italic ' : ''}${hit.c.bold ? '700 ' : ''}${px(hit.c.fontSize)}px ${getComputedStyle(this.paper).fontFamily}`;
            hit.c.w = Math.min(4000, Math.max(hit.c.w, Math.ceil(g.measureText(value).width) + 10));
            this.growSections();
          }
        });
      }
      this.render();
      this.canvas.focus({ preventScroll: true });
    };
    input.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); this.finishInline(true); }
      else if (e.key === 'Escape') { e.preventDefault(); this.finishInline(false); }
    });
    input.addEventListener('blur', () => setTimeout(() => this.finishInline(true), 0));
    input.addEventListener('pointerdown', e => e.stopPropagation());
    input.focus();
    input.select();
  }

  finishInline(commit) { if (this.inlineDone) this.inlineDone(commit); }

  // ---- the panes: Field List and Property Sheet ---------------------------------------------------------------------------------------

  toggleProps(force) {
    this.showProps = force ?? !this.showProps;
    savePanes({ ...readPanes(), props: this.showProps });
    this.renderSide();
    this.app.ribbonRefresh();
  }

  toggleFields(force) {
    this.showFields = force ?? !this.showFields;
    savePanes({ ...readPanes(), fields: this.showFields });
    this.renderSide();
    this.app.ribbonRefresh();
  }

  renderSide() {
    const panes = [];
    if (this.showFields) panes.push(this.fieldPane());
    if (this.showProps) { this.propsBox = el('section', { class: 'mo-base-ps', 'aria-label': t('Base.Dz.PropertySheet') }); panes.push(this.propsBox); this.renderProps(); } else this.propsBox = null;
    this.side.replaceChildren(...panes);
    this.side.hidden = !panes.length;
    this.root.classList.toggle('has-side', panes.length > 0);
  }

  fieldPane() {
    const box = el('section', { class: 'mo-base-fl', 'aria-label': t('Base.Dz.FieldList') });
    const head = el('div', { class: 'mo-base-pane__head' }, el('span', { text: t('Base.Dz.FieldList') }), iconButton('x', t('Editor.Close'), () => this.toggleFields(false)));
    const list = el('div', { class: 'mo-base-fl__list' });
    const o = this.obj;
    const row = (f, tableId) => {
      const b = el('button', { type: 'button', class: 'mo-base-fl__field', 'data-name': f.name, title: t('Base.Dz.FieldHint') }, icon(typeIcon(f.type)), el('span', { text: f.name }));
      b.addEventListener('dblclick', () => this.addFields([f], 'detail', null, tableId));
      b.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); this.addFields([f], 'detail', null, tableId); } });
      b.addEventListener('pointerdown', e => this.dragField(e, f, tableId));
      return b;
    };
    if (o.source && !this.allTables) {
      const name = tableById(this.db, o.source)?.name || queryById(this.db, o.source)?.name || '';
      list.append(el('div', { class: 'mo-base-fl__heading', text: t('Base.Dz.FieldsOf', name) }), ...this.fieldList().map(f => row(f, null)));
    } else {
      list.append(el('div', { class: 'mo-base-fl__heading', text: o.source ? t('Base.Dz.AllTables') : t('Base.Dz.NoSource') }));
      for (const tb of this.db.tables) {
        const open = this.openTable === tb.id;
        const head2 = el('button', { type: 'button', class: 'mo-base-fl__table', 'aria-expanded': String(open), onclick: () => { this.openTable = open ? null : tb.id; this.renderSide(); } }, icon(open ? 'chevron-down' : 'chevron-right'), icon(KIND_ICON.tables), el('span', { text: tb.name }));
        list.append(head2);
        if (open) list.append(...tb.fields.map(f => row({ name: f.name, caption: f.caption, type: f.type, lookup: !!f.lookup, field: f }, tb.id)));
      }
    }
    const foot = this.obj.source ? el('button', { type: 'button', class: 'mo-base-fl__toggle', onclick: () => { this.allTables = !this.allTables; this.renderSide(); } }, el('span', { text: this.allTables ? t('Base.Dz.OnlySource') : t('Base.Dz.ShowAllTables') })) : null;
    box.append(head, list, foot);
    return box;
  }

  dragField(e, f, tableId) {
    if (this.ro || e.button !== 0) return;
    let ghost = null;
    drag(e, (dx, dy, ev) => {
      if (!ghost) { if (Math.abs(dx) + Math.abs(dy) < 5) return; ghost = el('div', { class: 'mo-base-dragghost', text: f.name }); document.body.append(ghost); }
      ghost.style.left = ev.clientX + 8 + 'px';
      ghost.style.top = ev.clientY + 8 + 'px';
    }, (dx, dy, ev) => {
      if (!ghost) return;
      ghost.remove();
      if (tableId && this.obj.source && this.obj.source !== tableId) { this.app.toast(t('Base.Dz.OtherTable'), 'info'); return; }
      const key = this.sectionAt(ev.clientX, ev.clientY);
      if (!key) return;
      const r = this.secNode(key).getBoundingClientRect();
      this.addFields([f], key, { x: ev.clientX - r.left, y: Math.max(0, ev.clientY - r.top) }, tableId);
    });
  }

  /** The property rows of the selection: the controls, a section, or the form/report itself. */
  propRows() {
    const cs = this.selectedControls();
    const rows = [];
    const add = (key, label, tab, kind, more = {}) => rows.push({ key, label, tab, kind, ...more });
    if (cs.length) {
      const types = new Set(cs.map(c => c.type));
      const only = (...ts) => [...types].every(x => ts.includes(x));
      const one = cs.length === 1;
      if (one) add('name', t('Base.Prop.Name'), 'other', 'text', { max: 64 });
      if (only('label', 'button')) add('caption', t('Base.Prop.Caption'), 'format', 'text', { max: 1000 });
      if (only('textbox', 'combobox', 'checkbox')) add('source', t('Base.Prop.ControlSource'), 'data', 'source');
      if (one && types.has('combobox')) add('rowSource', t('Base.Prop.RowSource'), 'data', 'rowsource');
      if (one && types.has('subform')) {
        add('subform.source', t('Base.Prop.SourceObject'), 'data', 'enum', { options: () => this.sourceOptions(false), get: c => c.subform?.source || '', set: (c, v) => { c.subform = { ...(c.subform || {}), source: v }; } });
        add('subform.linkMaster', t('Base.Prop.LinkMaster'), 'data', 'text', { max: 64, list: () => this.fieldList().map(f => f.name), get: c => c.subform?.linkMaster || '', set: (c, v) => { c.subform = { ...(c.subform || {}), linkMaster: v }; } });
        add('subform.linkChild', t('Base.Prop.LinkChild'), 'data', 'text', { max: 64, list: () => sourceFieldList(this.db, cs[0].subform?.source, this.app.lang).map(f => f.name), get: c => c.subform?.linkChild || '', set: (c, v) => { c.subform = { ...(c.subform || {}), linkChild: v }; } });
      }
      if (only('textbox', 'combobox')) {
        add('format', t('Base.Prop.Format'), 'format', 'format');
        add('decimals', t('Base.Prop.Decimals'), 'format', 'enum', { options: () => [{ value: '', label: t('Base.Field.Auto') }, ...[...Array(16).keys()].map(n => ({ value: String(n), label: String(n) }))], get: c => (c.decimals === null || c.decimals === undefined ? '' : String(c.decimals)), set: (c, v) => { c.decimals = v === '' ? null : Number(v); } });
      }
      add('visible', t('Base.Prop.Visible'), 'format', 'bool');
      add('x', t('Base.Prop.Left'), 'format', 'int', { min: 0 });
      add('y', t('Base.Prop.Top'), 'format', 'int', { min: 0 });
      add('w', t('Base.Prop.Width'), 'format', 'int', { min: 1 });
      add('h', t('Base.Prop.Height'), 'format', 'int', { min: 1 });
      if (!only('line')) add('back', t('Base.Prop.BackColor'), 'format', 'color');
      add('color', only('line', 'rect') ? t('Base.Prop.BorderColor') : t('Base.Prop.ForeColor'), 'format', 'color');
      if (!only('line')) add('border', t('Base.Prop.Border'), 'format', 'bool', { get: c => !!(c.border ?? defaultBorder(c.type)) });
      if (!only('line', 'rect', 'subform')) {
        add('fontSize', t('Base.Prop.FontSize'), 'format', 'int', { min: 6, max: 96 });
        add('bold', t('Base.Prop.Bold'), 'format', 'bool');
        add('italic', t('Base.Prop.Italic'), 'format', 'bool');
        add('align', t('Base.Prop.TextAlign'), 'format', 'enum', { options: () => [{ value: 'left', label: t('Base.Prop.AlignLeft') }, { value: 'center', label: t('Base.Prop.AlignCenter') }, { value: 'right', label: t('Base.Prop.AlignRight') }] });
      }
      if (!this.isReport && only('textbox', 'combobox', 'checkbox')) add('locked', t('Base.Prop.Locked'), 'data', 'bool');
      if (this.isReport && only('textbox', 'combobox', 'checkbox')) add('hideDuplicates', t('Base.Prop.HideDuplicates'), 'format', 'bool');
      if (this.isReport && only('textbox')) add('runningSum', t('Base.Prop.RunningSum'), 'data', 'enum', { options: () => [{ value: 'no', label: t('Base.Prop.RunNo') }, { value: 'group', label: t('Base.Prop.RunGroup') }, { value: 'all', label: t('Base.Prop.RunAll') }] });
      if (only('button')) add('action', t('Base.Prop.OnClick'), 'other', 'enum', { options: () => ['', ...ACTIONS].map(a => ({ value: a, label: actionName(a) })) });
      return { kind: 'controls', list: cs, rows };
    }
    const s = this.target !== 'object' ? this.sectionOf(this.target) : null;
    if (s) {
      add('height', t('Base.Prop.Height'), 'format', 'int', { min: 0, max: 10000 });
      add('visible', t('Base.Prop.Visible'), 'format', 'bool');
      add('back', t('Base.Prop.BackColor'), 'format', 'color');
      if (/^gh\d+$/.test(s.key)) {
        const i = Number(s.key.slice(2));
        add('keepTogether', t('Base.Prop.KeepTogether'), 'format', 'bool', { get: () => !!this.obj.grouping[i]?.keepTogether, set: (sec, v) => { this.obj.grouping[i].keepTogether = v; } });
      }
      return { kind: 'section', list: [s.sec], rows, label: s.label };
    }
    const o = this.obj;
    add('source', t('Base.Prop.RecordSource'), 'data', 'enum', { options: () => this.sourceOptions(true) });
    add('caption', t('Base.Prop.Caption'), 'format', 'text', { max: 255 });
    if (!this.isReport) {
      add('kind', t('Base.Prop.DefaultView'), 'format', 'enum', { options: () => [{ value: 'single', label: t('Base.Form.Single') }, { value: 'continuous', label: t('Base.Form.Continuous') }, { value: 'datasheet', label: t('Base.Form.Datasheet') }, { value: 'split', label: t('Base.Form.Split') }] });
      add('allowEdits', t('Base.Prop.AllowEdits'), 'data', 'bool');
      add('allowAdditions', t('Base.Prop.AllowAdditions'), 'data', 'bool');
      add('allowDeletions', t('Base.Prop.AllowDeletions'), 'data', 'bool');
    } else {
      add('orientation', t('Base.Prop.Orientation'), 'format', 'enum', { options: () => [{ value: 'portrait', label: t('Base.Page.Portrait') }, { value: 'landscape', label: t('Base.Page.Landscape') }] });
      add('paper', t('Base.Prop.Paper'), 'format', 'enum', { options: () => [{ value: 'a4', label: 'A4' }, { value: 'letter', label: t('Base.Page.Letter') }] });
    }
    add('width', t('Base.Prop.Width'), 'format', 'int', { min: 100, max: 4000 });
    return { kind: 'object', list: [o], rows };
  }

  sourceOptions(withNone) {
    return [
      ...(withNone ? [{ value: '', label: t('Base.Prop.NoSource') }] : []),
      { group: t('Base.Nav.Tables'), options: this.db.tables.map(x => ({ value: x.id, label: x.name })) },
      { group: t('Base.Nav.Queries'), options: this.db.queries.map(x => ({ value: x.id, label: x.name })) },
    ];
  }

  /** The Property Sheet: what is selected, the tabs Format / Data / Other / All, a row per property. */
  renderProps() {
    const box = this.propsBox;
    if (!box) return;
    const target = this.propRows();
    const kindLabel = target.kind === 'controls'
      ? (target.list.length > 1 ? t('Base.Dz.MultipleSelection') : controlTypeName(target.list[0].type))
      : target.kind === 'section' ? t('Base.Dz.Section') : this.isReport ? t('Base.Kind.report') : t('Base.Kind.form');
    const picker = el('select', { class: 'mat-input mat-select mo-base-ps__pick', 'aria-label': t('Base.Dz.SelectObject') });
    picker.append(el('option', { value: 'object', text: this.isReport ? t('Base.Kind.report') : t('Base.Kind.form') }));
    for (const s of this.sections()) picker.append(el('option', { value: 'sec:' + s.key, text: s.label }));
    for (const x of this.allControls()) picker.append(el('option', { value: 'ctl:' + x.c.id, text: (x.c.name || controlTypeName(x.c.type)) + (x.c.type === 'label' && x.c.caption ? ' – ' + x.c.caption.slice(0, 30) : '') }));
    picker.value = target.kind === 'controls' ? (target.list.length === 1 ? 'ctl:' + target.list[0].id : '') : target.kind === 'section' ? 'sec:' + this.target : 'object';
    picker.addEventListener('change', () => {
      const v = picker.value;
      if (v === 'object') this.select([], 'object');
      else if (v.startsWith('sec:')) this.select([], v.slice(4));
      else { const id = v.slice(4); this.select([id], this.find(id)?.key); this.nodeOf(id)?.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
    });
    const tabs = el('div', { class: 'mo-base-ps__tabs', role: 'tablist' }, ...[['format', t('Base.Prop.TabFormat')], ['data', t('Base.Prop.TabData')], ['other', t('Base.Prop.TabOther')], ['all', t('Base.Prop.TabAll')]].map(([id, label]) =>
      el('button', { type: 'button', role: 'tab', 'aria-selected': String(this.propTab === id), onclick: () => { this.propTab = id; this.renderProps(); } }, el('span', { text: label }))));
    const list = el('div', { class: 'mo-base-ps__list' });
    for (const row of target.rows) if (this.propTab === 'all' || row.tab === this.propTab) list.append(this.propRow(row, target.list));
    if (!list.children.length) list.append(el('div', { class: 'mo-base-ps__empty', text: t('Base.Dz.NoProps') }));
    box.replaceChildren(
      el('div', { class: 'mo-base-pane__head' }, el('span', { text: t('Base.Dz.PropertySheet') }), iconButton('x', t('Editor.Close'), () => this.toggleProps(false))),
      el('div', { class: 'mo-base-ps__type', text: t('Base.Dz.SelectionType', kindLabel) }), picker, tabs, list);
  }

  propRow(row, list) {
    const get = row.get || (o => o[row.key]);
    const values = list.map(get);
    const same = values.every(v => JSON.stringify(v) === JSON.stringify(values[0]));
    const value = same ? values[0] : null;
    const id = nextId('mo-base-pp');
    const ro = this.ro;
    const commit = v => this.setProp(row, v, list);
    let input;
    switch (row.kind) {
      case 'bool':
        input = selectInput([{ value: '', label: '' }, { value: 'yes', label: t('Base.Yes') }, { value: 'no', label: t('Base.No') }].filter(o => o.value || !same), same ? (value ? 'yes' : 'no') : '', { disabled: ro ? true : null });
        input.addEventListener('change', () => { if (input.value) commit(input.value === 'yes'); });
        break;
      case 'enum':
        input = selectInput(row.options(), same ? String(value ?? '') : '', { disabled: ro ? true : null });
        if (!same) input.prepend(el('option', { value: '__', text: '' }));
        if (!same) input.value = '__';
        input.addEventListener('change', () => { if (input.value !== '__') commit(input.value); });
        break;
      case 'int':
        input = textInput(same && value !== null && value !== undefined ? String(value) : '', { inputmode: 'numeric', readonly: ro ? true : null });
        input.addEventListener('change', () => {
          const n = Math.round(Number(input.value));
          if (input.value.trim() === '' || !Number.isFinite(n)) { input.value = same ? String(value ?? '') : ''; return; }
          commit(Math.min(row.max ?? 10000, Math.max(row.min ?? 0, n)));
        });
        break;
      case 'color': {
        const pick = el('input', { type: 'color', class: 'mo-base-ps__color', 'aria-label': row.label, disabled: ro ? true : null, value: same && value ? value : '#000000' });
        pick.addEventListener('change', () => commit(pick.value));
        const clear = iconButton('x', t('Base.Prop.Automatic'), () => commit(null), { disabled: ro || !value ? true : null });
        input = el('span', { class: 'mo-base-ps__colorrow' }, pick, el('span', { class: 'mo-base-ps__hex', text: same ? value || t('Base.Prop.Automatic') : '' }), value ? clear : null);
        pick.id = id;
        break;
      }
      case 'rowsource': {
        const c = list[0];
        const text = c.rowSource?.kind === 'table' ? (tableById(this.db, c.rowSource.table)?.name || '') : c.rowSource?.kind === 'values' ? c.rowSource.values.join('; ') : t('Base.Prop.FromField');
        input = el('span', { class: 'mo-base-ps__colorrow' }, el('span', { class: 'mo-base-ps__hex', text }), el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', 'aria-label': row.label, disabled: ro ? true : null, onclick: () => this.editRowSource(c) }, el('span', { text: '…' })));
        break;
      }
      default: {
        const listId = row.kind === 'source' || row.kind === 'format' || row.list ? nextId('mo-base-pl') : null;
        input = textInput(same ? String(value ?? '') : '', { maxlength: row.max || 2000, readonly: ro ? true : null, list: listId, spellcheck: 'false' });
        input.addEventListener('change', () => commit(input.value.trim()));
        input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); input.dispatchEvent(new Event('change')); } });
        if (listId) {
          const options = row.kind === 'source' ? this.fieldList().map(f => f.name) : row.kind === 'format' ? formatOptions(value).map(o => o.value).filter(Boolean) : row.list();
          const dl = el('datalist', { id: listId }, ...options.map(o => el('option', { value: o })));
          input.id = id;
          input = el('span', { class: 'mo-base-ps__colorrow' }, input, dl,
            row.kind === 'source' && !ro ? el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', 'aria-label': t('Base.Dz.Builder'), 'data-tip': t('Base.Dz.Builder'), onclick: () => this.builder(list, row) }, el('span', { text: '…' })) : null);
        }
      }
    }
    const inner = input.matches('input, select') ? input : input.querySelector('input, select');
    if (inner && !inner.id) inner.id = id;
    input.dataset.prop = row.key;
    return el('div', { class: 'mo-base-ps__row' }, inner ? el('label', { for: inner.id, text: row.label }) : el('span', { class: 'mo-base-ps__label', text: row.label }), input);
  }

  focusProp(key) {
    requestAnimationFrame(() => {
      const n = this.propsBox?.querySelector(`[data-prop="${CSS.escape(key)}"]`);
      const target = n?.matches('input, select') ? n : n?.querySelector('input, select');
      target?.focus();
    });
  }

  setProp(row, value, list) {
    if (this.ro) return;
    const set = row.set || ((o, v) => { o[row.key] = v; });
    const get = row.get || (o => o[row.key]);
    if (list.every(o => JSON.stringify(get(o) ?? null) === JSON.stringify(value ?? null))) return;
    const before = JSON.stringify(list.map(get));
    this.app.design(() => {
      for (const o of list) set(o, value);
      if (row.key === 'orientation') {
        // a report as wide as the old page takes the width of the new one
        const wide = { portrait: 718, landscape: 1047 };
        const o = list[0];
        if (o.width === wide[value === 'portrait' ? 'landscape' : 'portrait']) o.width = wide[value];
      }
      this.growSections();
    });
    if (JSON.stringify(list.map(row.get || (o => o[row.key]))) === before) return;
    if (row.key === 'source') { this.fieldCache = null; if (this.live) this.loadLive(); }
    this.render();
  }

  /** The Row Source of a combo box: the lookup of its field, a table (bound and shown field) or a list of values. */
  async editRowSource(c) {
    const db = this.db;
    const kind = selectInput([{ value: 'field', label: t('Base.Prop.FromField') }, { value: 'table', label: t('Base.Lookup.TableQuery') }, { value: 'values', label: t('Base.Lookup.ValueList') }], c.rowSource?.kind || 'field');
    const table = selectInput(db.tables.map(x => ({ value: x.id, label: x.name })), c.rowSource?.table || db.tables[0]?.id);
    const bound = selectInput([], null);
    const shown = selectInput([], null);
    const values = el('textarea', { class: 'mat-input mat-textarea', rows: 5 });
    values.value = c.rowSource?.kind === 'values' ? c.rowSource.values.join('\n') : '';
    const fillFields = () => {
      const tb = tableById(db, table.value);
      const opts = (tb?.fields || []).map(f => el('option', { value: f.id, text: f.name }));
      bound.replaceChildren(...opts.map(o => o.cloneNode(true)));
      shown.replaceChildren(...opts);
      const pk = tb?.primaryKey[0];
      bound.value = c.rowSource?.table === table.value ? c.rowSource.bound : pk || tb?.fields[0]?.id || '';
      shown.value = c.rowSource?.table === table.value ? c.rowSource.display : tb?.fields.find(f => f.id !== bound.value && f.type === 'text')?.id || bound.value;
    };
    const tableBox = el('div', {}, field(t('Base.Lookup.RowSource'), table), field(t('Base.Lookup.Bound'), bound), field(t('Base.Lookup.Shown'), shown));
    const valuesBox = field(t('Base.Lookup.Values'), values, t('Base.Lookup.ValuesHint'));
    const sync = () => { tableBox.hidden = kind.value !== 'table'; valuesBox.hidden = kind.value !== 'values'; };
    table.addEventListener('change', fillFields);
    kind.addEventListener('change', sync);
    fillFields();
    sync();
    const r = await dialog({
      title: t('Base.Prop.RowSource'), icon: 'base-c-combo', size: 'md', body: el('div', { class: 'mo-base-form' }, field(t('Base.Lookup.RowSourceType'), kind), tableBox, valuesBox),
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    const next = kind.value === 'table' ? { kind: 'table', table: table.value, bound: bound.value, display: shown.value }
      : kind.value === 'values' ? { kind: 'values', values: values.value.split(/\r?\n|;/).map(x => x.trim()).filter(Boolean).slice(0, 500) } : null;
    this.app.design(() => { c.rowSource = next; });
    this.render();
  }

  /** The Expression Builder, short: the text with the fields of the record source and the functions to put into it. */
  async builder(list, row) {
    const area = el('textarea', { class: 'mat-input mat-textarea mo-base-code', rows: 4, spellcheck: 'false' });
    area.value = list.length === 1 ? (list[0][row.key] || '') : '';
    const insert = text => { area.setRangeText(text, area.selectionStart, area.selectionEnd, 'end'); area.focus(); };
    const fields = el('div', { class: 'mo-base-builder__list' }, ...this.fieldList().map(f => el('button', { type: 'button', class: 'mo-base-builder__item', text: f.name, onclick: () => insert((area.value.trim() ? '' : '=') + '[' + f.name + ']') })));
    const fns = el('div', { class: 'mo-base-builder__list' }, ...['Sum', 'Avg', 'Count', 'Min', 'Max', ...FUNCTION_NAMES.filter(n => !['Sum', 'Avg', 'Count', 'Min', 'Max'].includes(n))].map(n => el('button', { type: 'button', class: 'mo-base-builder__item', text: n, onclick: () => insert((area.value.trim() ? '' : '=') + n + '()') })));
    const ops = el('div', { class: 'mo-base-builder__ops' }, ...['+', '-', '*', '/', '&', '=', '<', '>', 'And', 'Or', 'Not', '"', '(', ')'].map(o => el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', text: o, onclick: () => insert(o.length > 1 ? ' ' + o + ' ' : o) })));
    const r = await dialog({
      title: t('Base.Dz.Builder'), icon: 'base-builder', size: 'lg',
      body: el('div', { class: 'mo-base-form mo-base-builder' }, area, ops, el('div', { class: 'mo-base-builder__cols' },
        el('div', {}, el('div', { class: 'mo-base-fl__heading', text: t('Base.Dz.Fields') }), fields),
        el('div', {}, el('div', { class: 'mo-base-fl__heading', text: t('Base.Dz.Functions') }), fns)), el('p', { class: 'mat-hint', text: t('Base.Dz.BuilderHint') })),
      onOpen: () => area.focus(),
      validate: () => {
        const v = area.value.trim();
        if (v.startsWith('=')) { try { Parser.expression(v.slice(1)); } catch (e) { if (e instanceof BaseError) { this.app.toast(t('Base.Err.Syntax', v, ''), 'danger'); return false; } } }
        return true;
      },
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r === 'ok') this.setProp(row, area.value.trim(), list);
  }

  // ---- header and footer commands ------------------------------------------------------------------------------------------------

  /** Title: a big label in the header (the form/report header is made visible when it has no room). */
  addTitle() {
    if (this.ro) return;
    const key = this.isReport ? 'reportHeader' : 'header';
    const o = this.obj;
    let id = null;
    this.app.design(() => {
      const sec = this.sectionOf(key).sec;
      const existing = sec.controls.find(c => c.type === 'label' && c.fontSize >= 18);
      if (existing) { id = existing.id; return; }
      if (sec.height < 56) sec.height = 56;
      const c = newControl('label', newId(this.allControls().map(x => x.c), 'c'), { x: this.isReport ? 0 : 24, y: 8, w: Math.min(o.width - 48, 480), h: 40 }, { caption: o.caption || o.name, fontSize: 22, name: 'Title' });
      sec.controls.push(c);
      id = c.id;
    });
    this.select([id], key);
    this.render();
    this.editInline(id);
  }

  /** Date and Time: =Date() right in the header. */
  addDate() {
    if (this.ro) return;
    const key = this.isReport ? 'reportHeader' : 'header';
    const o = this.obj;
    let id = null;
    this.app.design(() => {
      const sec = this.sectionOf(key).sec;
      if (sec.height < 40) sec.height = 40;
      const w = 200;
      const c = newControl('textbox', newId(this.allControls().map(x => x.c), 'c'), { x: Math.max(0, o.width - w - (this.isReport ? 0 : 16)), y: 12, w, h: 22 }, { source: '=Date()', format: 'Long Date', align: 'right', border: false, name: this.controlName('textbox') });
      sec.controls.push(c);
      id = c.id;
    });
    this.select([id], key);
    this.render();
  }

  // ---- ribbon tabs shared by forms and reports -----------------------------------------------------------------------------------

  controlsGroup(types) {
    const b = type => ({ type: 'toggle', showLabel: true, label: controlTypeName(type), icon: CONTROL_ICON[type], active: () => this.insert === type, run: () => this.insertMode(type) });
    const rows = [];
    for (let i = 0; i < types.length; i += 3) rows.push({ type: 'stack', rows: types.slice(i, i + 3).map(x => [b(x)]) });
    return { label: t('Base.Group.Controls'), items: rows };
  }

  toolsGroup() {
    return { label: t('Base.Group.Tools'), items: [
      { type: 'toggle', size: 'large', label: t('Base.Dz.AddFields'), icon: 'base-field-list', edit: false, active: () => this.showFields, run: () => this.toggleFields() },
      { type: 'toggle', size: 'large', label: t('Base.Dz.PropertySheet'), icon: 'sliders', edit: false, key: 'F4', global: false, active: () => this.showProps, run: () => this.toggleProps() },
    ] };
  }

  formatTab() {
    const cs = () => this.selectedControls();
    const textual = () => cs().filter(c => c.type !== 'line' && c.type !== 'rect' && c.type !== 'subform');
    const all = (key, v) => { const l = textual(); return l.length > 0 && l.every(c => c[key] === v); };
    const set = (key, v, list = textual()) => this.setControls(key, v, list);
    const alignBtn = (a, ic, label) => ({ type: 'toggle', label, icon: ic, active: () => all('align', a), disabled: () => !textual().length, run: () => set('align', a) });
    return {
      id: 'dformat', label: t('Base.Tab.Format'), groups: [
        { label: t('Base.Group.Font'), items: [{ type: 'stack', rows: [
          [{ type: 'combo', label: t('Base.Prop.FontSize'), options: ['8', '9', '10', '11', '12', '14', '16', '18', '20', '22', '24', '28', '36', '48'], width: 64, value: () => { const l = textual(); return l.length && l.every(c => c.fontSize === l[0].fontSize) ? String(l[0].fontSize) : ''; }, disabled: () => !textual().length, change: v => { const n = Math.round(Number(String(v).replace(',', '.'))); if (n >= 6 && n <= 96) set('fontSize', n); } }],
          [
            { type: 'toggle', label: t('Base.Prop.Bold'), icon: 'base-bold', active: () => all('bold', true), disabled: () => !textual().length, run: () => set('bold', !all('bold', true)) },
            { type: 'toggle', label: t('Base.Prop.Italic'), icon: 'base-italic', active: () => all('italic', true), disabled: () => !textual().length, run: () => set('italic', !all('italic', true)) },
            alignBtn('left', 'base-text-left', t('Base.Prop.AlignLeft')), alignBtn('center', 'base-text-center', t('Base.Prop.AlignCenter')), alignBtn('right', 'base-text-right', t('Base.Prop.AlignRight')),
          ],
          [
            { type: 'color', label: t('Base.Prop.ForeColor'), icon: 'base-font-color', auto: t('Base.Prop.Automatic'), initial: '#c00000', disabled: () => !cs().length, change: v => set('color', v, cs()) },
            { type: 'color', label: t('Base.Prop.BackColor'), icon: 'base-fill', auto: t('Base.Prop.Automatic'), initial: '#fff2cc', disabled: () => !cs().length, change: v => set('back', v, cs().filter(c => c.type !== 'line')) },
            { type: 'toggle', label: t('Base.Prop.Border'), icon: 'base-border', active: () => cs().length > 0 && cs().every(c => c.border ?? defaultBorder(c.type)), disabled: () => !cs().length, run: () => set('border', !cs().every(c => c.border ?? defaultBorder(c.type)), cs()) },
          ],
        ] }] },
        { label: t('Base.Group.Number'), items: [{ type: 'select', label: t('Base.Prop.Format'), width: 140, options: () => formatOptions(cs()[0]?.format), value: () => { const l = cs().filter(c => c.type === 'textbox' || c.type === 'combobox'); return l.length && l.every(c => c.format === l[0].format) ? l[0].format : ''; }, disabled: () => !cs().some(c => c.type === 'textbox' || c.type === 'combobox'), change: v => set('format', v, cs().filter(c => c.type === 'textbox' || c.type === 'combobox')) }] },
      ],
    };
  }

  arrangeTab() {
    const many = () => this.sel.size > 1;
    const some = () => this.sel.size > 0;
    const sm = (label, ic, run, dis) => ({ type: 'button', showLabel: true, label, icon: ic, run, disabled: dis });
    return {
      id: 'darrange', label: t('Base.Tab.Arrange'), groups: [
        { label: t('Base.Group.Align'), items: [
          { type: 'stack', rows: [[sm(t('Base.Dz.AlignLeft'), 'base-align-left', () => this.align('left'), () => !many())], [sm(t('Base.Dz.AlignRight'), 'base-align-right', () => this.align('right'), () => !many())], [sm(t('Base.Dz.AlignGrid'), 'base-grid', () => this.align('grid'), () => !some())]] },
          { type: 'stack', rows: [[sm(t('Base.Dz.AlignTop'), 'base-align-top', () => this.align('top'), () => !many())], [sm(t('Base.Dz.AlignBottom'), 'base-align-bottom', () => this.align('bottom'), () => !many())]] },
        ] },
        { label: t('Base.Group.Size'), items: [
          { type: 'stack', rows: [[sm(t('Base.Dz.SizeTallest'), 'base-size-tallest', () => this.sizeTo('tallest'), () => !many())], [sm(t('Base.Dz.SizeShortest'), 'base-size-shortest', () => this.sizeTo('shortest'), () => !many())]] },
          { type: 'stack', rows: [[sm(t('Base.Dz.SizeWidest'), 'base-size-widest', () => this.sizeTo('widest'), () => !many())], [sm(t('Base.Dz.SizeNarrowest'), 'base-size-narrowest', () => this.sizeTo('narrowest'), () => !many())]] },
        ] },
        { label: t('Base.Group.Position'), items: [
          { type: 'button', size: 'large', label: t('Base.Dz.Front'), icon: 'base-front', run: () => this.order(true), disabled: () => !some() },
          { type: 'button', size: 'large', label: t('Base.Dz.Back'), icon: 'base-back', run: () => this.order(false), disabled: () => !some() },
        ] },
        { label: t('Base.Group.Selection'), items: [{ type: 'stack', rows: [
          [{ type: 'button', showLabel: true, label: t('Base.Dz.SelectAll'), icon: 'base-select', edit: false, run: () => this.selectAll() }],
          [sm(t('Base.Dz.Delete'), 'trash', () => this.deleteSelected(), () => !some())],
          [sm(t('Base.Dz.Duplicate'), 'copy', () => this.duplicate(), () => !some())],
        ] }] },
      ],
    };
  }

  // ---- Home commands --------------------------------------------------------------------------------------------------------------

  can(cmd) {
    if (cmd === 'copy') return this.sel.size > 0;
    if (cmd === 'cut') return !this.ro && this.sel.size > 0;
    if (cmd === 'paste') return !this.ro && !!clipboard?.length;
    if (cmd === 'refresh') return this.live;
    return false;
  }

  run(cmd) {
    if (cmd === 'copy') this.copy();
    else if (cmd === 'cut') this.cut();
    else if (cmd === 'paste') this.paste();
    else if (cmd === 'refresh') { this.loadLive(); this.render(); }
  }

  status() {
    const cs = this.selectedControls();
    if (cs.length === 1) return [t('Base.Dz.StatusOne', controlTypeName(cs[0].type), cs[0].name || cs[0].caption || '')];
    if (cs.length > 1) return [t('Base.Dz.StatusMany', cs.length)];
    return [];
  }
}

/** The icon of a field type (Field List, wizards). */
export function typeIcon(type) {
  return { text: 'base-t-text', memo: 'base-t-memo', number: 'base-t-number', currency: 'base-t-currency', datetime: 'base-t-date', yesno: 'base-t-yesno', autonumber: 'base-t-auto', hyperlink: 'link' }[type] || 'base-t-text';
}
