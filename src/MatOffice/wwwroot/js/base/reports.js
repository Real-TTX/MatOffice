/* MatOffice Base – reports like Access: the Report view (all records on one long sheet), Print Preview with real pages (page header and footer,
   "Page 1 of 3", the headers and footers of the groups with their totals, the grand totals in the report footer), the Design view (the designer
   with the Group, Sort and Total pane, Totals, Page Numbers), the Database Documenter, and printing: exactly the pages of a report (CSS @page),
   a datasheet as a table, the records of a form, the Relationships window. "PDF" is the browser's print to PDF. */
import { t, el, icon, stamp } from '../editor/util.js';
import { BaseView } from './view.js';
import { recordsetOf, EmptyRecordset, TableRecordset } from './recordset.js';
import { ControlEval, domainFunction, sortRows, buildBands, paginate, pageSize, MARGIN, sourceFieldList, pageExpression, printableWidth } from './layout.js';
import { Designer, styleControl, newControl, controlTypeName } from './designer.js';
import { askParams } from './queryviews.js';
import { parseSql, queryType } from './sql.js';
import { showProblem, ask, iconButton, problemText, selectInput, KIND_ICON, typeName, sizeName } from './ui.js';
import { findObject, queryById, tableById, fieldById, newId, lower, relationshipKind } from './model.js';
import { BaseError, toNumber } from './expr.js';
import { formatValue } from './values.js';

// ---- running a report ------------------------------------------------------------------------------------------------------------------

/** Asks the parameters of the record source once per tab (Report view and Print Preview share them). */
async function paramsFor(app, tab, report) {
  if (tab.params) return tab.params;
  const q = queryById(app.db, report.source);
  if (!q) return {};
  const p = await askParams(app, parseSql(q.sql));
  if (p) tab.params = p;
  return p;
}

/** The records, the bands and the pages of a report: { rs, ev, rows, bands, pages } or { error }. */
export function prepareReport(app, report, params = {}) {
  const db = app.db;
  const rs = report.source ? recordsetOf(app, report.source, params) : new EmptyRecordset(app);
  if (!rs) return { error: { code: 'NoTable', args: [report.source] } };
  if (rs.error) return { error: rs.error };
  const ev = new ControlEval(rs, { lang: app.lang, currency: db.settings.currency, domain: domainFunction(db, app.lang) });
  const rows = sortRows(report, rs, ev);
  const bands = buildBands(report, rows, ev);
  const pages = paginate(report, bands);
  return { rs, ev, rows, bands, pages };
}

const visible = sec => sec && sec.visible !== false && sec.height > 0;

/** What a report keeps while it is drawn in order: the last values of Hide Duplicates, the running sums. */
function newState() { return { prev: new Map(), sums: new Map() }; }

function bandEl(report, band, run, page, state) {
  const sec = band.section;
  const node = el('div', { class: 'mo-base-band mo-base-band--' + band.kind });
  node.style.height = sec.height + 'px';
  node.style.width = report.width + 'px';
  if (sec.back) node.style.backgroundColor = sec.back;
  if (band.kind === 'groupHeader') {
    for (const [id, s] of state.sums) if (s.over === 'group') state.sums.delete(id);
    state.prev.clear();
  }
  for (const c of sec.controls) {
    if (c.visible === false || c.type === 'button') continue;
    node.append(reportControl(c, band, run, page, state));
  }
  return node;
}

function reportControl(c, band, run, page, state) {
  const n = el('div', { class: 'mo-base-rc mo-base-rc--' + c.type });
  styleControl(n, c);
  const { ev, rs } = run;
  switch (c.type) {
    case 'label': n.append(el('span', { class: 'mo-base-rc__text', text: c.caption || '' })); break;
    case 'line': n.classList.toggle('is-vertical', c.h > c.w); n.append(el('span', { class: 'mo-base-dc__line' })); break;
    case 'rect': break;
    case 'checkbox': {
      let v = null;
      try { v = c.source ? ev.value(c.source, band.row, band.rows, page) : null; } catch { v = null; }
      n.append(el('span', { class: 'mo-base-rc__box' + (v === true ? ' is-on' : '') }, v === true ? icon('check') : null));
      break;
    }
    case 'subform': {
      const name = tableById(run.db, c.subform?.source)?.name || queryById(run.db, c.subform?.source)?.name || '';
      n.append(el('div', { class: 'mo-base-rc__sub', text: name }));
      break;
    }
    default: {
      let text;
      if (c.runningSum && c.runningSum !== 'no' && c.source) {
        let v = null;
        try { v = ev.value(c.source, band.row, band.rows, page); } catch { v = null; }
        const s = state.sums.get(c.id) || { over: c.runningSum, total: 0 };
        let add = 0;
        try { add = v === null || v === undefined ? 0 : toNumber(v, run.lang); } catch { add = 0; }
        s.total += add;
        state.sums.set(c.id, s);
        text = rs.display({ key: '__sum', type: 'number', field: null, format: c.format || '', decimals: c.decimals ?? null }, s.total);
      } else {
        text = ev.text(c, band.row, band.rows, page);
      }
      if (c.hideDuplicates && band.kind === 'detail') {
        const prev = state.prev.get(c.id);
        state.prev.set(c.id, text);
        if (prev === text) text = '';
      }
      n.append(el('span', { class: 'mo-base-rc__text', text }));
    }
  }
  return n;
}

/** One page of Print Preview (and of the printout): the page header, the bands, the page footer at the bottom. */
function pageEl(report, p, index, total, run, state) {
  const size = pageSize(report);
  const s = report.sections;
  const page = { page: index + 1, pages: total };
  const node = el('div', { class: 'mo-base-page', 'data-page': index + 1 });
  node.style.width = size.w + 'px';
  node.style.height = size.h + 'px';
  const body = el('div', { class: 'mo-base-page__body' });
  Object.assign(body.style, { left: MARGIN + 'px', top: MARGIN + 'px', width: size.w - 2 * MARGIN + 'px', height: size.h - 2 * MARGIN + 'px' });
  const place = (bandNode, y) => { bandNode.style.top = y + 'px'; body.append(bandNode); };
  const head = visible(s.pageHeader) ? s.pageHeader.height : 0;
  const first = p.bands[0];
  const rh = index === 0 && first?.band.kind === 'reportHeader' ? first.band.section.height : 0;
  for (const { band, y } of p.bands) place(bandEl(report, band, run, page, state), band.kind === 'reportHeader' && index === 0 ? 0 : y + head);
  if (head) place(bandEl(report, { kind: 'pageHeader', section: s.pageHeader, row: p.bands[0]?.band.row || run.rows[0] || null, rows: run.rows }, run, page, state), rh);
  if (visible(s.pageFooter)) {
    const last = p.bands[p.bands.length - 1]?.band;
    place(bandEl(report, { kind: 'pageFooter', section: s.pageFooter, row: last?.row || null, rows: run.rows }, run, page, state), size.h - 2 * MARGIN - s.pageFooter.height);
  }
  node.append(body);
  return node;
}

/** All pages of a report as elements. */
export function reportPages(app, report, run) {
  const state = newState();
  const r = { ...run, db: app.db, lang: app.lang };
  return run.pages.map((p, i) => pageEl(report, p, i, run.pages.length, r, state));
}

// ---- Report view ------------------------------------------------------------------------------------------------------------------------

export class ReportView extends BaseView {
  get report() { return this.obj; }

  async mount() {
    this.el.replaceChildren(el('div', { class: 'mo-loading' }, el('span', { class: 'mo-loading__spinner' })));
    let params;
    try { params = await paramsFor(this.app, this.tab, this.report); } catch (e) { if (!(e instanceof BaseError)) throw e; this.fail(e); return; }
    if (params === null) { this.app.open('reports', this.tab.id, 'design'); return; }
    this.params = params;
    this.draw();
  }

  fail(e) {
    this.data = null;
    this.el.replaceChildren(el('div', { class: 'mo-base-empty' }, el('div', { class: 'mo-base-empty__icon' }, icon('alert')), el('h2', { text: t('Base.Report.Failed') }), el('p', { text: problemText(e) }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => this.app.open('reports', this.tab.id, 'design') }, icon('base-v-design'), el('span', { text: t('Base.View.Design') }))));
  }

  draw() {
    const report = this.report;
    const run = prepareReport(this.app, report, this.params);
    if (run.error) { this.fail(run.error); return; }
    this.data = run;
    const s = report.sections;
    const state = newState();
    const r = { ...run, db: this.db, lang: this.app.lang };
    const page = { page: 1, pages: 1 };
    const sheet = el('div', { class: 'mo-base-rv__sheet' });
    sheet.style.width = report.width + 'px';
    const flow = (band, cls = '') => { const n = bandEl(report, band, r, page, state); n.classList.add('is-flow'); if (cls) n.classList.add(cls); sheet.append(n); };
    const bands = run.bands;
    let i = 0;
    if (bands[0]?.kind === 'reportHeader') flow(bands[i++]);
    if (visible(s.pageHeader)) flow({ kind: 'pageHeader', section: s.pageHeader, row: run.rows[0] || null, rows: run.rows });
    for (; i < bands.length; i++) flow(bands[i]);
    if (visible(s.pageFooter)) flow({ kind: 'pageFooter', section: s.pageFooter, row: run.rows[run.rows.length - 1] || null, rows: run.rows });
    this.el.replaceChildren(el('div', { class: 'mo-base-rv' }, sheet));
    this.app.setStatus();
  }

  refresh() { if (this.data || this.params) this.draw(); }

  can(cmd) { return cmd === 'refresh'; }

  run(cmd) { if (cmd === 'refresh') this.draw(); }

  status() { return this.data ? [t('Base.Report.Records', this.data.rows.length)] : []; }

  contextTabs() { return []; }
}

// ---- Print Preview ----------------------------------------------------------------------------------------------------------------------

export class PreviewView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.z = app.isPhone() ? 0.45 : 1;
    this.two = false;
    this.pageIndex = 0;
  }

  get report() { return this.obj; }

  async mount() {
    this.el.replaceChildren(el('div', { class: 'mo-loading' }, el('span', { class: 'mo-loading__spinner' })));
    let params;
    try { params = await paramsFor(this.app, this.tab, this.report); } catch (e) { if (!(e instanceof BaseError)) throw e; ReportView.prototype.fail.call(this, e); return; }
    if (params === null) { this.app.open('reports', this.tab.id, 'design'); return; }
    this.params = params;
    this.draw();
  }

  draw() {
    const report = this.report;
    const run = prepareReport(this.app, report, this.params);
    if (run.error) { ReportView.prototype.fail.call(this, run.error); return; }
    this.data = run;
    this.pages = el('div', { class: 'mo-base-pv__pages' + (this.two ? ' is-two' : '') });
    this.pages.append(...reportPages(this.app, report, run));
    this.pages.style.zoom = String(this.z);
    this.scroller = el('div', { class: 'mo-base-pv__scroll' }, this.pages);
    this.scroller.addEventListener('scroll', () => this.trackPage());
    this.navBar = el('div', { class: 'mo-base-ds__nav' });
    this.el.replaceChildren(el('div', { class: 'mo-base-pv' }, this.scroller, this.navBar));
    this.pageIndex = Math.min(this.pageIndex, run.pages.length - 1);
    this.renderNav();
  }

  trackPage() {
    const pages = [...this.pages.children];
    const top = this.scroller.scrollTop + 40;
    let i = 0;
    for (let k = 0; k < pages.length; k++) if (pages[k].offsetTop * this.z <= top) i = k;
    if (i !== this.pageIndex) { this.pageIndex = i; this.renderNav(); }
  }

  goPage(i) {
    const n = this.data?.pages.length || 0;
    this.pageIndex = Math.max(0, Math.min(n - 1, i));
    const page = this.pages.children[this.pageIndex];
    if (page) this.scroller.scrollTop = page.offsetTop * this.z - 12;
    this.renderNav();
  }

  renderNav() {
    if (!this.data) return;
    const n = this.data.pages.length;
    const box = el('input', { class: 'mo-base-ds__pos', value: String(this.pageIndex + 1), 'aria-label': t('Base.Preview.Page'), inputmode: 'numeric' });
    box.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); const k = Number(box.value); if (k >= 1 && k <= n) this.goPage(k - 1); } });
    const btn = (name, label, run, off) => iconButton(name, label, run, { class: 'mo-base-ibtn mo-base-ds__navbtn', disabled: off ? true : null });
    this.navBar.replaceChildren(
      el('span', { class: 'mo-base-ds__navlabel', text: t('Base.Preview.Page') }),
      btn('base-first', t('Base.Preview.First'), () => this.goPage(0), this.pageIndex <= 0),
      btn('chevron-left', t('Base.Preview.Previous'), () => this.goPage(this.pageIndex - 1), this.pageIndex <= 0),
      box,
      el('span', { class: 'mo-base-ds__navof', text: t('Base.Nav.Of', n) }),
      btn('chevron-right', t('Base.Preview.Next'), () => this.goPage(this.pageIndex + 1), this.pageIndex >= n - 1),
      btn('base-last', t('Base.Preview.Last'), () => this.goPage(n - 1), this.pageIndex >= n - 1));
  }

  zoom() {
    return {
      value: this.z, min: 0.25, max: 3,
      change: f => { this.z = f; if (this.pages) this.pages.style.zoom = String(f); },
      fit: () => this.fit(),
    };
  }

  fit() {
    if (!this.pages || !this.scroller) return;
    const size = pageSize(this.report);
    const w = this.scroller.clientWidth - 32;
    const h = this.scroller.clientHeight - 32;
    this.z = Math.max(0.25, Math.min(3, Math.floor(Math.min(w / ((this.two ? 2 : 1) * size.w + (this.two ? 24 : 0)), h / size.h) * 100) / 100));
    this.pages.style.zoom = String(this.z);
    this.app.ctx.status.setZoom(this.z);
  }

  setPage(patch) {
    if (this.app.readOnly) { Object.assign(this.report, patch); this.draw(); return; }
    this.app.design(() => {
      const r = this.report;
      const before = r.orientation;
      Object.assign(r, patch);
      if (patch.orientation && before !== patch.orientation) {
        const wide = { portrait: 718, landscape: 1047 };
        if (r.width === wide[before]) r.width = wide[patch.orientation];
      }
    });
    this.draw();
  }

  refresh() { if (this.data) this.draw(); }

  can(cmd) { return cmd === 'refresh'; }

  run(cmd) { if (cmd === 'refresh') this.draw(); }

  status() { return this.data ? [t('Base.Preview.Pages', this.data.pages.length)] : []; }

  contextTabs() {
    const r = () => this.report;
    return [{
      id: 'preview', label: t('Base.Tab.Preview'), groups: [
        { label: t('Base.Group.Print'), items: [{ type: 'button', size: 'large', label: t('Base.Preview.Print'), icon: 'print', edit: false, run: () => this.app.print() }] },
        { label: t('Base.Group.PageSize'), items: [
          { type: 'toggle', size: 'large', label: 'A4', icon: 'base-page-size', edit: false, active: () => r()?.paper !== 'letter', run: () => this.setPage({ paper: 'a4' }) },
          { type: 'toggle', size: 'large', label: t('Base.Page.Letter'), icon: 'base-page-size', edit: false, active: () => r()?.paper === 'letter', run: () => this.setPage({ paper: 'letter' }) },
        ] },
        { label: t('Base.Group.PageLayout'), items: [
          { type: 'toggle', size: 'large', label: t('Base.Page.Portrait'), icon: 'base-portrait', edit: false, active: () => r()?.orientation !== 'landscape', run: () => this.setPage({ orientation: 'portrait' }) },
          { type: 'toggle', size: 'large', label: t('Base.Page.Landscape'), icon: 'base-landscape', edit: false, active: () => r()?.orientation === 'landscape', run: () => this.setPage({ orientation: 'landscape' }) },
        ] },
        { label: t('Base.Group.Zoom'), items: [
          { type: 'button', size: 'large', label: t('Base.Preview.Zoom100'), icon: 'zoom-in', edit: false, run: () => { this.z = 1; this.pages.style.zoom = '1'; this.app.ctx.status.setZoom(1); } },
          { type: 'toggle', size: 'large', label: t('Base.Preview.OnePage'), icon: 'base-one-page', edit: false, active: () => !this.two, run: () => { this.two = false; this.pages?.classList.remove('is-two'); this.fit(); } },
          { type: 'toggle', size: 'large', label: t('Base.Preview.TwoPages'), icon: 'base-two-pages', edit: false, active: () => this.two, run: () => { this.two = true; this.pages?.classList.add('is-two'); this.fit(); } },
        ] },
        { label: t('Base.Group.Data'), items: [
          { type: 'button', size: 'large', label: t('Base.Ext.Pdf'), icon: 'base-pdf', edit: false, tip: t('Base.Ext.PdfTip'), run: () => this.app.print() },
          { type: 'button', size: 'large', label: t('Base.Ext.Excel'), icon: 'base-excel-export', edit: false, disabled: () => !r()?.source, tip: t('Base.Preview.ExcelTip'), run: () => { const src = r().source; this.app.exportExcel(tableById(this.db, src) ? 'tables' : 'queries', src); } },
        ] },
        { label: t('Base.Group.Close'), items: [{ type: 'button', size: 'large', label: t('Base.Preview.Close'), icon: 'x', edit: false, run: () => this.app.open('reports', this.tab.id, 'report') }] },
      ],
    }];
  }
}

// ---- Design view ---------------------------------------------------------------------------------------------------------------------

export class ReportDesignView extends Designer {
  constructor(app, tab) { super(app, tab, false); }

  /** The Group, Sort and Total pane: the grouping levels and the sort order of the report. */
  groupPane() {
    const r = this.obj;
    const ro = this.ro;
    const fields = sourceFieldList(this.db, r.source, this.app.lang).map(f => f.name);
    const fieldSelect = (value, change) => {
      const options = [...new Set([...fields, ...(value ? [value] : [])])].map(f => ({ value: f, label: f }));
      const s = selectInput(options, value, { disabled: ro ? true : null, class: 'mat-input mat-select mo-base-gs__field' });
      s.addEventListener('change', () => change(s.value));
      return s;
    };
    const order = (desc, change) => {
      const s = selectInput([{ value: 'asc', label: t('Base.Group.AtoZ') }, { value: 'desc', label: t('Base.Group.ZtoA') }], desc ? 'desc' : 'asc', { disabled: ro ? true : null, class: 'mat-input mat-select mo-base-gs__sel' });
      s.addEventListener('change', () => change(s.value === 'desc'));
      return s;
    };
    const yesNo = (on, labels, change) => {
      const s = selectInput([{ value: 'yes', label: labels[0] }, { value: 'no', label: labels[1] }], on ? 'yes' : 'no', { disabled: ro ? true : null, class: 'mat-input mat-select mo-base-gs__sel' });
      s.addEventListener('change', () => change(s.value === 'yes'));
      return s;
    };
    const rows = [];
    (r.grouping || []).forEach((g, i) => {
      rows.push(el('div', { class: 'mo-base-gs__row' }, icon('base-group'),
        el('span', { class: 'mo-base-gs__word', text: t('Base.Group.GroupOn') }), fieldSelect(g.field, v => this.changeGroup(i, { field: v })),
        order(g.desc, v => this.changeGroup(i, { desc: v })),
        yesNo(g.header !== false, [t('Base.Group.WithHeader'), t('Base.Group.WithoutHeader')], v => this.changeGroup(i, { header: v })),
        yesNo(!!g.footer, [t('Base.Group.WithFooter'), t('Base.Group.WithoutFooter')], v => this.changeGroup(i, { footer: v })),
        yesNo(!!g.keepTogether, [t('Base.Group.KeepTogether'), t('Base.Group.NoKeep')], v => this.changeGroup(i, { keepTogether: v })),
        ro ? null : iconButton('arrow-up', t('Base.Group.MoveUp'), () => this.moveGroup(i, -1), { disabled: i === 0 ? true : null }),
        ro ? null : iconButton('arrow-down', t('Base.Group.MoveDown'), () => this.moveGroup(i, 1), { disabled: i === r.grouping.length - 1 ? true : null }),
        ro ? null : iconButton('trash', t('Base.Group.Delete'), () => this.removeGroup(i))));
    });
    (r.sorting || []).forEach((s, i) => {
      rows.push(el('div', { class: 'mo-base-gs__row' }, icon('base-sort-asc'),
        el('span', { class: 'mo-base-gs__word', text: i === 0 && !(r.grouping || []).length ? t('Base.Group.SortBy') : t('Base.Group.ThenBy') }), fieldSelect(s.field, v => this.changeSort(i, { field: v })),
        order(s.desc, v => this.changeSort(i, { desc: v })),
        ro ? null : iconButton('trash', t('Base.Group.Delete'), () => this.removeSort(i))));
    });
    const can = !ro && fields.length > 0;
    const buttons = el('div', { class: 'mo-base-gs__add' },
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: can && (r.grouping || []).length < 10 ? null : true, onclick: () => this.addGroup() }, icon('base-group'), el('span', { text: t('Base.Group.AddGroup') })),
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: can && (r.sorting || []).length < 10 ? null : true, onclick: () => this.addSort() }, icon('base-sort-asc'), el('span', { text: t('Base.Group.AddSort') })));
    return el('section', { class: 'mo-base-gs', 'aria-label': t('Base.Group.Pane') },
      el('div', { class: 'mo-base-pane__head' }, el('span', { text: t('Base.Group.Pane') }), iconButton('x', t('Editor.Close'), () => this.toggleGroups(false))),
      el('div', { class: 'mo-base-gs__list' }, ...rows, rows.length ? null : el('div', { class: 'mo-base-gs__empty', text: fields.length ? t('Base.Group.None') : t('Base.Group.NoSource') })),
      buttons);
  }

  toggleGroups(force) {
    this.showGroups = force ?? !this.showGroups;
    try { localStorage.setItem('mo.base.designer', JSON.stringify({ ...JSON.parse(localStorage.getItem('mo.base.designer') || '{}'), groups: this.showGroups })); } catch { /* private mode */ }
    this.renderBottom();
    this.app.ribbonRefresh();
  }

  addGroup() {
    const r = this.obj;
    const fields = sourceFieldList(this.db, r.source, this.app.lang);
    const used = new Set((r.grouping || []).map(g => lower(g.field)));
    const f = fields.find(x => !used.has(lower(x.name))) || fields[0];
    if (!f) return;
    this.app.design(() => {
      r.grouping ||= [];
      r.grouping.push({ field: f.name, header: true, footer: false, desc: false, keepTogether: false });
      const all = () => this.allControls().map(x => x.c);
      r.sections.groupHeaders.push({ height: 30, visible: true, back: null, controls: [newControl('textbox', newId(all(), 'c'), { x: 0, y: 4, w: 300, h: 22 }, { source: f.name, name: this.controlName('textbox'), bold: true, fontSize: 12, border: false })] });
      r.sections.groupFooters.push({ height: 0, visible: true, back: null, controls: [] });
    });
    this.render();
  }

  async removeGroup(i) {
    const r = this.obj;
    const has = r.sections.groupHeaders[i]?.controls.length || r.sections.groupFooters[i]?.controls.length;
    if (has && !(await ask(t('Base.Group.DeleteAsk'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    this.app.design(() => {
      r.grouping.splice(i, 1);
      r.sections.groupHeaders.splice(i, 1);
      r.sections.groupFooters.splice(i, 1);
    });
    this.sel.clear();
    this.target = 'detail';
    this.render();
  }

  moveGroup(i, dir) {
    const r = this.obj;
    const j = i + dir;
    if (j < 0 || j >= r.grouping.length) return;
    this.app.design(() => {
      for (const list of [r.grouping, r.sections.groupHeaders, r.sections.groupFooters]) [list[i], list[j]] = [list[j], list[i]];
    });
    this.render();
  }

  changeGroup(i, patch) {
    const r = this.obj;
    this.app.design(() => {
      Object.assign(r.grouping[i], patch);
      if (patch.header && !r.sections.groupHeaders[i].height) r.sections.groupHeaders[i].height = 30;
      if (patch.footer && !r.sections.groupFooters[i].height) r.sections.groupFooters[i].height = 30;
    });
    this.render();
  }

  addSort() {
    const r = this.obj;
    const fields = sourceFieldList(this.db, r.source, this.app.lang);
    const used = new Set([...(r.grouping || []), ...(r.sorting || [])].map(g => lower(g.field)));
    const f = fields.find(x => !used.has(lower(x.name))) || fields[0];
    if (!f) return;
    this.app.design(() => { r.sorting ||= []; r.sorting.push({ field: f.name, desc: false }); });
    this.render();
  }

  changeSort(i, patch) { this.app.design(() => { Object.assign(this.obj.sorting[i], patch); }); this.render(); }

  removeSort(i) { this.app.design(() => { this.obj.sorting.splice(i, 1); }); this.render(); }

  /** Totals of Access: =Sum([Field]) (or Avg, Count …) under the selected text box in every group footer and in the report footer. */
  totals(fn) {
    const c = this.selectedControls()[0];
    if (this.ro || !c || c.type !== 'textbox' || !c.source || c.source.startsWith('=')) return;
    const r = this.obj;
    const expr = fn === 'Count' ? '=Count(*)' : '=' + fn + '([' + c.source.replace(/^\[|\]$/g, '') + '])';
    this.app.design(() => {
      const put = sec => {
        if (sec.height < 26) sec.height = 26;
        const id = newId(this.allControls().map(x => x.c), 'c');
        sec.controls.push(newControl('textbox', id, { x: c.x, y: 2, w: c.w, h: 22 }, { source: expr, name: this.controlName('textbox'), bold: true, border: false, align: 'right', format: fn === 'Count' ? '' : c.format || '' }));
        if (!sec.controls.some(x => x.type === 'line' && x.y === 0 && x.x === c.x)) sec.controls.push(newControl('line', newId(this.allControls().map(x => x.c), 'c'), { x: c.x, y: 0, w: c.w, h: 1 }, { name: this.controlName('line') }));
      };
      (r.grouping || []).forEach((g, i) => { g.footer = true; put(r.sections.groupFooters[i]); });
      put(r.sections.reportFooter);
    });
    this.render();
  }

  /** Page Numbers: "Page N of M" in the page footer (the old one is replaced). */
  pageNumbers() {
    if (this.ro) return;
    const r = this.obj;
    this.app.design(() => {
      const sec = r.sections.pageFooter;
      sec.controls = sec.controls.filter(c => !/\[Pages?\]/i.test(c.source || ''));
      if (sec.height < 30) sec.height = 30;
      const w = 260;
      sec.controls.push(newControl('textbox', newId(this.allControls().map(x => x.c), 'c'), { x: Math.max(0, r.width - w), y: 6, w, h: 20 }, {
        source: pageExpression({ page: t('Base.Report.PageWord'), of: t('Base.Report.OfWord') }), name: this.controlName('textbox'), align: 'right', border: false, fontSize: 9,
      }));
    });
    this.render();
  }

  contextTabs() {
    const one = () => { const c = this.selectedControls(); return c.length === 1 && c[0].type === 'textbox' && c[0].source && !c[0].source.startsWith('=') ? c[0] : null; };
    const numeric = () => { const c = one(); const f = c && sourceFieldList(this.db, this.obj.source, this.app.lang).find(x => lower(x.name) === lower(c.source)); return f && ['number', 'currency', 'autonumber'].includes(f.type); };
    return [
      {
        id: 'rdesign', label: t('Base.Tab.ReportDesign'), groups: [
          { label: t('Base.Group.GroupingTotals'), items: [
            { type: 'toggle', size: 'large', label: t('Base.Group.GroupSort'), icon: 'base-group', edit: false, active: () => this.showGroups, run: () => this.toggleGroups() },
            { type: 'menu', size: 'large', label: t('Base.Group.Totals'), icon: 'base-totals', disabled: () => !one(), items: () => [
              { label: t('Base.Total.Sum'), run: () => this.totals('Sum'), disabled: () => !numeric() },
              { label: t('Base.Total.Avg'), run: () => this.totals('Avg'), disabled: () => !numeric() },
              { label: t('Base.Total.Count'), run: () => this.totals('Count') },
              { label: t('Base.Total.Max'), run: () => this.totals('Max') },
              { label: t('Base.Total.Min'), run: () => this.totals('Min') },
            ] },
          ] },
          this.controlsGroup(['label', 'textbox', 'checkbox', 'line', 'rect']),
          { label: t('Base.Group.HeaderFooter'), items: [{ type: 'stack', rows: [
            [{ type: 'button', showLabel: true, label: t('Base.Dz.PageNumbers'), icon: 'base-page-numbers', run: () => this.pageNumbers() }],
            [{ type: 'button', showLabel: true, label: t('Base.Dz.Title'), icon: 'base-title', run: () => this.addTitle() }],
            [{ type: 'button', showLabel: true, label: t('Base.Dz.DateTime'), icon: 'base-date-time', run: () => this.addDate() }],
          ] }] },
          this.toolsGroup(),
          { label: t('Base.Group.Views'), items: [
            { type: 'button', size: 'large', label: t('Base.View.Preview'), icon: 'base-v-preview', edit: false, run: () => this.app.open('reports', this.tab.id, 'preview') },
          ] },
        ],
      },
      this.formatTab(),
      this.arrangeTab(),
      {
        id: 'rpage', label: t('Base.Tab.PageSetup'), groups: [
          { label: t('Base.Group.PageSize'), items: [
            { type: 'toggle', size: 'large', label: 'A4', icon: 'base-page-size', active: () => this.obj?.paper !== 'letter', run: () => this.setPage({ paper: 'a4' }) },
            { type: 'toggle', size: 'large', label: t('Base.Page.Letter'), icon: 'base-page-size', active: () => this.obj?.paper === 'letter', run: () => this.setPage({ paper: 'letter' }) },
          ] },
          { label: t('Base.Group.PageLayout'), items: [
            { type: 'toggle', size: 'large', label: t('Base.Page.Portrait'), icon: 'base-portrait', active: () => this.obj?.orientation !== 'landscape', run: () => this.setPage({ orientation: 'portrait' }) },
            { type: 'toggle', size: 'large', label: t('Base.Page.Landscape'), icon: 'base-landscape', active: () => this.obj?.orientation === 'landscape', run: () => this.setPage({ orientation: 'landscape' }) },
          ] },
        ],
      },
    ];
  }

  setPage(patch) {
    PreviewView.prototype.setPage.call({ report: this.obj, app: this.app, draw: () => {} }, patch);
    this.render();
  }

  status() {
    const out = super.status();
    const r = this.obj;
    if (r && r.width > printableWidth(r)) out.push(t('Base.Report.TooWide'));
    return out;
  }
}

// ---- the Database Documenter ---------------------------------------------------------------------------------------------------------------

export class DocumenterView extends BaseView {
  mount() { this.draw(); }

  refresh() { this.draw(); }

  draw() {
    this.doc = documenter(this.app);
    this.el.replaceChildren(el('div', { class: 'mo-base-docv' }, this.doc));
  }

  contextTabs() {
    return [{ id: 'documenter', label: t('Base.Tab.Documenter'), groups: [
      { label: t('Base.Group.Print'), items: [{ type: 'button', size: 'large', label: t('Base.Preview.Print'), icon: 'print', edit: false, run: () => this.app.print() }] },
      { label: t('Base.Group.Close'), items: [{ type: 'button', size: 'large', label: t('Base.Preview.Close'), icon: 'x', edit: false, run: () => this.app.closeTab(this.tab) }] },
    ] }];
  }

  status() { return [t('Base.Doc.Objects', this.db.tables.length + this.db.queries.length + this.db.forms.length + this.db.reports.length)]; }
}

/** A printable description of every object (the Database Documenter of Access). */
function documenter(app) {
  const db = app.db;
  const lang = app.lang;
  const name = app.ctx.doc.mode === 'host' ? app.ctx.doc.fileName : app.ctx.doc.name;
  const table = (head, rows) => el('table', { class: 'mo-base-doc__table' },
    el('thead', {}, el('tr', {}, ...head.map(h => el('th', { text: h })))),
    el('tbody', {}, ...rows.map(r => el('tr', {}, ...r.map((c, i) => el('td', { text: c ?? '', 'data-label': head[i] }))))));
  const out = el('article', { class: 'mo-base-doc' },
    el('h1', { text: t('Base.Doc.Heading', name) }),
    el('p', { class: 'mo-base-doc__meta', text: t('Base.Doc.Made', stamp(new Date())) }));
  const section = (kind, title) => { const s = el('section', { class: 'mo-base-doc__kind' }, el('h2', {}, icon(KIND_ICON[kind]), el('span', { text: title }))); out.append(s); return s; };
  if (db.tables.length) {
    const s = section('tables', t('Base.Nav.Tables'));
    for (const tb of db.tables) {
      const fields = tb.fields.map(f => [
        f.name, typeName(f.type), f.type === 'text' ? String(f.size ?? 255) : f.type === 'number' ? sizeName(f.numberSize || 'long') : '',
        tb.primaryKey.includes(f.id) ? t('Base.Yes') : '', f.required ? t('Base.Yes') : t('Base.No'), { no: t('Base.Index.No'), dup: t('Base.Index.Dup'), unique: t('Base.Index.Unique') }[f.indexed || 'no'], f.description || '',
      ]);
      const rels = db.relationships.filter(r => r.parent.table === tb.id || r.child.table === tb.id).map(r => {
        const p = tableById(db, r.parent.table);
        const c = tableById(db, r.child.table);
        return `${p?.name}.${fieldById(p, r.parent.field)?.name} → ${c?.name}.${fieldById(c, r.child.field)?.name}` + (r.integrity ? ' (' + t('Base.Doc.Integrity') + (r.cascadeUpdate ? ', ' + t('Base.Doc.CascadeUpdate') : '') + (r.cascadeDelete ? ', ' + t('Base.Doc.CascadeDelete') : '') + ', ' + (relationshipKind(db, r) === 'oneToOne' ? '1:1' : '1:n') + ')' : '');
      });
      s.append(el('div', { class: 'mo-base-doc__obj' },
        el('h3', { text: tb.name }),
        tb.description ? el('p', { text: tb.description }) : null,
        el('p', { class: 'mo-base-doc__meta', text: t('Base.Nav.TableInfo', tb.fields.length, tb.rows.length) }),
        table([t('Base.Design.FieldName'), t('Base.Design.DataType'), t('Base.Field.Size'), t('Base.Design.PrimaryKey'), t('Base.Field.Required'), t('Base.Field.Indexed'), t('Base.Design.Description')], fields),
        rels.length ? el('div', {}, el('h4', { text: t('Base.Info.Relationships') }), el('ul', {}, ...rels.map(x => el('li', { text: x })))) : null));
    }
  }
  if (db.queries.length) {
    const s = section('queries', t('Base.Nav.Queries'));
    for (const q of db.queries) {
      let type = 'select';
      try { type = queryType(parseSql(q.sql)); } catch { type = 'select'; }
      s.append(el('div', { class: 'mo-base-doc__obj' }, el('h3', { text: q.name }), el('p', { class: 'mo-base-doc__meta', text: t('Base.Query.Type.' + type) }),
        q.description ? el('p', { text: q.description }) : null, el('pre', { class: 'mo-base-doc__sql', text: q.sql || '' })));
    }
  }
  const sourceName = id => tableById(db, id)?.name || queryById(db, id)?.name || t('Base.Prop.NoSource');
  if (db.forms.length) {
    const s = section('forms', t('Base.Nav.Forms'));
    for (const f of db.forms) {
      const bound = Object.values(f.sections).flatMap(sec => sec.controls).filter(c => c.source && !c.source.startsWith('=')).map(c => c.source);
      s.append(el('div', { class: 'mo-base-doc__obj' }, el('h3', { text: f.name }),
        el('p', { class: 'mo-base-doc__meta', text: t('Base.Doc.FormInfo', sourceName(f.source), { single: t('Base.Form.Single'), continuous: t('Base.Form.Continuous'), datasheet: t('Base.Form.Datasheet'), split: t('Base.Form.Split') }[f.kind] || f.kind, Object.values(f.sections).reduce((n, sec) => n + sec.controls.length, 0)) }),
        bound.length ? el('p', { text: t('Base.Doc.Fields', [...new Set(bound)].join(', ')) }) : null));
    }
  }
  if (db.reports.length) {
    const s = section('reports', t('Base.Nav.Reports'));
    for (const r of db.reports) {
      s.append(el('div', { class: 'mo-base-doc__obj' }, el('h3', { text: r.name }),
        el('p', { class: 'mo-base-doc__meta', text: t('Base.Doc.ReportInfo', sourceName(r.source), r.orientation === 'landscape' ? t('Base.Page.Landscape') : t('Base.Page.Portrait')) }),
        (r.grouping || []).length ? el('p', { text: t('Base.Doc.Grouping', r.grouping.map(g => g.field).join(', ')) }) : null,
        (r.sorting || []).length ? el('p', { text: t('Base.Doc.Sorting', r.sorting.map(g => g.field + (g.desc ? ' ↓' : '')).join(', ')) }) : null));
    }
  }
  void lang;
  return out;
}

// ---- printing ---------------------------------------------------------------------------------------------------------------------------

const cssText = s => '"' + String(s ?? '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n]+/g, ' ') + '"';

/** Prints nodes alone (the rest of the page is hidden for the printout), with the page size and margins given. */
function printNodes(nodes, { size = 'A4', landscape = false, margin = '0', title = null, numbers = false } = {}) {
  document.querySelector('.mo-base-print')?.remove();
  const boxes = [];
  if (title) boxes.push(`@top-left { content: ${cssText(title)}; font: 9pt sans-serif; color: #444; }`, `@top-right { content: ${cssText(stamp(new Date()))}; font: 9pt sans-serif; color: #444; }`);
  if (numbers) boxes.push(`@bottom-right { content: ${cssText(t('Base.Report.PageWord') + ' ')} counter(page) ${cssText(' ' + t('Base.Report.OfWord') + ' ')} counter(pages); font: 9pt sans-serif; color: #444; }`);
  const style = el('style', { text: `@page { size: ${size} ${landscape ? 'landscape' : 'portrait'}; margin: ${margin}; ${boxes.join(' ')} }` });
  const root = el('div', { class: 'mo-base-print' }, style, ...nodes);
  document.body.append(root);
  document.body.classList.add('mo-base-printing');
  const done = () => {
    window.removeEventListener('afterprint', done);
    setTimeout(() => { root.remove(); document.body.classList.remove('mo-base-printing'); }, 100);
  };
  window.addEventListener('afterprint', done);
  setTimeout(() => window.print(), 60);
}

/** Prints the pages of a report exactly as Print Preview shows them. */
export async function printReport(app, report, params = null) {
  let p = params;
  if (!p) {
    const q = queryById(app.db, report.source);
    try { p = q ? await askParams(app, parseSql(q.sql)) : {}; } catch (e) { await showProblem(e); return; }
    if (p === null) return;
  }
  const run = prepareReport(app, report, p);
  if (run.error) { await showProblem(run.error); return; }
  const size = report.paper === 'letter' ? 'letter' : 'A4';
  printNodes(reportPages(app, report, run), { size, landscape: report.orientation === 'landscape' });
}

/** Prints records as a table (a datasheet): the name and the date at the top of every page, the page numbers at the bottom. */
function printSheet(app, title, rs, cols, rows) {
  const head = el('thead', {}, el('tr', {}, ...cols.map(c => el('th', { class: c.align === 'right' ? 'is-right' : null, text: c.caption }))));
  const body = el('tbody', {}, ...rows.map(r => el('tr', {}, ...cols.map(c => {
    const v = rs.get(r, c);
    const text = c.type === 'yesno' ? (v === true ? '☑' : '☐') : rs.display(c, v);
    return el('td', { class: c.align === 'right' ? 'is-right' : c.align === 'center' ? 'is-center' : null, text });
  }))));
  const doc = el('div', { class: 'mo-base-psheet' }, el('h1', { text: title }), el('table', {}, head, body));
  const wide = cols.reduce((s, c) => s + (c.type === 'memo' ? 2 : 1), 0) > 7;
  printNodes([doc], { landscape: wide, margin: '16mm 12mm', title, numbers: true });
}

/** Prints the records of a form one after the other (header first, footer last), like Access prints a form. */
async function printForm(app, form) {
  let params = {};
  const q = queryById(app.db, form.source);
  if (q) {
    try { params = await askParams(app, parseSql(q.sql)); } catch (e) { await showProblem(e); return; }
    if (params === null) return;
  }
  const rs = form.source ? recordsetOf(app, form.source, params) : new EmptyRecordset(app);
  if (!rs || rs.error) { await showProblem(rs?.error || { code: 'NoTable', args: [form.source] }); return; }
  const ev = new ControlEval(rs, { lang: app.lang, currency: app.db.settings.currency, domain: domainFunction(app.db, app.lang) });
  const run = { rs, ev, rows: rs.rows, db: app.db, lang: app.lang };
  const state = newState();
  const page = { page: 1, pages: 1 };
  const s = form.sections;
  const width = form.width;
  const band = (key, row) => {
    const sec = s[key];
    if (!visible(sec)) return null;
    const n = bandEl({ width }, { kind: key === 'detail' ? 'detail' : 'reportHeader', section: sec, row, rows: rs.rows }, run, page, state);
    n.classList.add('is-flow', 'mo-base-pform__band');
    return n;
  };
  const parts = [band('header', rs.rows[0] || null)];
  const rows = form.kind === 'single' || form.kind === 'split' ? rs.rows.slice(0, 500) : rs.rows.slice(0, 2000);
  for (const r of rows) parts.push(band('detail', r));
  if (!rows.length) parts.push(band('detail', null));
  parts.push(band('footer', rs.rows[rs.rows.length - 1] || null));
  const doc = el('div', { class: 'mo-base-pform' + (form.kind === 'single' || form.kind === 'split' ? ' is-single' : '') }, ...parts.filter(Boolean));
  printNodes([doc], { landscape: width > 720, margin: '14mm 12mm', title: form.caption || form.name, numbers: true });
}

/** Prints any object (the navigation pane's "Export as PDF", External Data → PDF). */
export async function printObject(app, kind, id) {
  const obj = findObject(app.db, kind, id);
  if (!obj) return;
  switch (kind) {
    case 'reports': await printReport(app, obj); return;
    case 'forms': await printForm(app, obj); return;
    case 'tables': {
      const rs = new TableRecordset(app, id);
      const hidden = new Set(obj.datasheet?.hidden || []);
      printSheet(app, obj.name, rs, rs.columns.filter(c => !hidden.has(c.skey)), rs.rows);
      return;
    }
    case 'queries': {
      let params;
      try { params = await askParams(app, parseSql(obj.sql)); } catch (e) { await showProblem(e); return; }
      if (params === null) return;
      const rs = recordsetOf(app, id, params);
      if (rs.error) { await showProblem(rs.error); return; }
      printSheet(app, obj.name, rs, rs.columns, rs.rows);
      return;
    }
    default:
  }
}

/** Prints what the active tab shows: a report's pages, the datasheet as it is (sorted, filtered), the records of a form, the Relationships window, the documenter. */
export async function printTab(app, tab) {
  const v = tab.viewObj;
  if (tab.view === 'datasheet' && v?.ds && (tab.kind === 'tables' || tab.kind === 'queries')) {
    if (!(await v.ds.commit())) return;
    printSheet(app, app.tabTitle(tab), v.rs, v.ds.cols, v.ds.view);
    return;
  }
  if (tab.kind === 'relationships' && v?.canvas) {
    const copy = v.canvas.cloneNode(true);
    copy.classList.add('is-print');
    const w = Math.max(v.canvas.scrollWidth, 200);
    const h = Math.max(v.canvas.scrollHeight, 200);
    const frame = el('div', { class: 'mo-base-prel' }, el('h1', { text: t('Base.Rel.Title') }), copy);
    copy.style.width = w + 'px';
    copy.style.height = h + 'px';
    copy.style.zoom = String(Math.min(1, 1000 / w, 640 / h));
    printNodes([frame], { landscape: true, margin: '12mm', title: t('Base.Rel.Title'), numbers: false });
    return;
  }
  if (tab.kind === 'documenter' && v?.doc) {
    printNodes([v.doc.cloneNode(true)], { margin: '16mm 14mm', title: t('Base.Doc.Title'), numbers: true });
    return;
  }
  if (tab.kind === 'reports') {
    const report = findObject(app.db, 'reports', tab.id);
    if (report) await printReport(app, report, tab.params || null);
    return;
  }
  await printObject(app, tab.kind, tab.id);
}

export { controlTypeName, formatValue };
