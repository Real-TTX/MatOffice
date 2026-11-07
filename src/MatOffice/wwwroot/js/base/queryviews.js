/* MatOffice Base – the views of a query: the Datasheet of its result (editable where the query is updatable), the SQL view, and the Query Design
   view of Access: the tables with their join lines above (Show Table, drag a field onto a field to join, Join Properties), the design grid below
   (Field, Table, Total, Sort, Show, Criteria, or; Append To, Update To, Delete for action queries). Grid and SQL stay in step through qbe.js.
   Parameters ([Enter a city]) are asked for in "Enter Parameter Value" dialogs; action queries confirm "You are about to update N row(s)". */
import { t, el, icon, dialog } from '../editor/util.js';
import { BaseView, datasheetCan, datasheetRun } from './view.js';
import { Datasheet } from './datasheet.js';
import { QueryRecordset } from './recordset.js';
import { parseSql, queryType, toSql } from './sql.js';
import { parametersOf, paramValues, planAction, runQueryObject } from './engine.js';
import { sqlToDesign, designToStatement, emptyDesign, emptyColumn, TOTALS, splitAlias, tidyFieldCell } from './qbe.js';
import { tableByName, queryByName, lower, sourceByName, tableById, fieldById, isUniqueField, newObjectId, uniqueName } from './model.js';
import { BaseError } from './expr.js';
import { showProblem, ask, field, textInput, selectInput, checkInput, radioGroup, contextMenu, drag, problemText, queryIcon, iconButton } from './ui.js';

/** Asks the parameters of a statement (one "Enter Parameter Value" dialog each); null when cancelled. */
export async function askParams(app, stmt) {
  const list = parametersOf(app.db, stmt, { lang: app.lang });
  if (!list.length) return {};
  const answers = {};
  for (const p of list) {
    const input = textInput('', { autocomplete: 'off' });
    const r = await dialog({
      title: t('Base.Param.Title'), icon: 'edit', body: field(p.name, input), onOpen: () => input.focus(),
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return null;
    answers[lower(p.name)] = input.value;
  }
  return paramValues(list, answers, app.lang);
}

/** Runs an action query with the questions of Access and its report afterwards. */
export async function runActionQuery(app, query) {
  if (app.readOnly) { app.toast(t('Base.Query.ReadOnly'), 'info'); return; }
  let stmt;
  try { stmt = parseSql(query.sql); } catch (e) { await showProblem(e); return; }
  const type = queryType(stmt);
  if (type === 'select') return;
  if (!(await ask(t('Base.Query.RunAsk.' + type, query.name), { title: t('Base.Query.RunTitle'), danger: type === 'delete' || type === 'update', ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
  const params = await askParams(app, stmt).catch(e => { showProblem(e); return null; });
  if (params === null) return;
  let plan;
  try { plan = planAction(app.db, stmt, { params, lang: app.lang }); } catch (e) { if (e instanceof BaseError) { await showProblem(e); return; } throw e; }
  if (plan.blocked) { await showProblem(plan.blocked); return; }
  if (type === 'maketable' && plan.exists && !(await ask(t('Base.Query.Replace', plan.name), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
  const msg = { update: 'Base.Query.Update', delete: 'Base.Query.Delete', append: 'Base.Query.Append', maketable: 'Base.Query.MakeTable' }[type];
  if (!(await ask(t(msg, plan.count), { title: t('Base.Query.RunTitle'), danger: type === 'delete', ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
  const report = plan.apply();
  const touched = plan.table ? [plan.table.id] : report.table ? [report.table.id] : [];
  app.recordsChanged(touched.length ? touched : app.db.tables.map(x => x.id));
  app.nav.render();
  for (const tab of app.tabs) tab.viewObj.version = -1;
  app.activeView?.activate();
  if (report.error && !report.done) { await showProblem(report.error); return; }
  if (report.typeFailures || report.keyViolations || report.ruleViolations) {
    await dialog({ title: t('Base.Query.Problems'), icon: 'alert', body: el('p', { text: t('Base.Query.ProblemsText', report.typeFailures, report.keyViolations, report.ruleViolations) }) });
  }
  app.toast(t('Base.Query.Done.' + type, report.done) + (report.cascaded ? ' ' + t('Base.Query.Cascaded', report.cascaded) : ''), 'success');
}

// ---- datasheet -----------------------------------------------------------------------------------------------------------------------

export class QueryDatasheetView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.selectContextTab = false;
  }

  async mount() {
    const q = this.obj;
    this.el.replaceChildren(el('div', { class: 'mo-loading' }, el('span', { class: 'mo-loading__spinner' })));
    let params = {};
    try {
      params = await askParams(this.app, parseSql(q.sql));
    } catch (e) {
      if (!(e instanceof BaseError)) throw e;
      this.showError(e);
      return;
    }
    if (params === null) { this.app.open('queries', q.id, 'design'); return; }
    this.params = params;
    this.rs = new QueryRecordset(this.app, q.id, params);
    if (this.rs.error) { this.showError(this.rs.error); return; }
    this.ds = new Datasheet(this.app, {
      recordset: this.rs, settings: q.datasheet, persist: () => this.app.settingsChanged(), onStatus: () => { this.app.setStatus(); this.app.ribbonRefresh(); },
    });
    this.el.replaceChildren(this.ds.el);
    if (this.app.active?.viewObj === this) this.focus();
  }

  showError(e) {
    this.ds = null;
    this.el.replaceChildren(el('div', { class: 'mo-base-empty' }, el('div', { class: 'mo-base-empty__icon' }, icon('alert')), el('h2', { text: t('Base.Query.Failed') }), el('p', { text: problemText(e) }),
      el('div', { class: 'mat-btn-row mat-btn-row--plain' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => this.app.open('queries', this.tab.id, 'design') }, icon('base-v-design'), el('span', { text: t('Base.View.Design') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary', onclick: () => this.app.open('queries', this.tab.id, 'sql') }, icon('base-v-sql'), el('span', { text: t('Base.View.Sql') })))));
  }

  refresh() {
    if (!this.ds || !this.obj) return;
    this.ds.s = this.obj.datasheet;
    this.rs.load();
    if (this.rs.error) { this.showError(this.rs.error); return; }
    this.ds.reload();
  }

  focus() { if (this.ds) requestAnimationFrame(() => { if (!this.ds.editor && !this.ds.sel) this.ds.moveTo(this.ds.cur.r, this.ds.cur.c, { focus: true }); else this.ds.focusEditor(false); }); }

  async leave() { return this.ds ? this.ds.commit() : true; }

  destroy() { this.ds?.destroy(); }

  can(cmd) { return datasheetCan(this.ds, cmd); }

  run(cmd, arg) { return this.ds ? datasheetRun(this.ds, cmd, arg) : undefined; }

  selectionLabel(kind) { return this.ds?.selectionLabel(kind); }

  status() {
    if (!this.ds) return [];
    const out = [];
    if (!this.rs.canEdit) out.push(t('Base.Status.ReadOnlyResult'));
    if (this.ds.s.filterOn && this.ds.s.filter) out.push(t('Base.Status.Filtered'));
    return out;
  }

  contextTabs() { return queryTabs(this.app, this, null); }
}

/** The contextual tab "Query Design" (also in the datasheet and SQL views, for Run and the views). */
function queryTabs(app, view, designView) {
  const q = () => view.obj;
  const type = () => { try { return queryType(parseSql(q().sql)); } catch { return 'select'; } };
  const groups = [{ label: t('Base.Group.Results'), items: [
    { type: 'button', size: 'large', label: t('Base.Query.Run'), icon: 'base-run', edit: false, run: () => view.runQuery() },
  ] }];
  if (designView) {
    const setType = kind => designView.setType(kind);
    const tbtn = (kind, ic) => ({ type: 'toggle', size: 'large', label: t('Base.Query.Type.' + kind), icon: ic, active: () => designView.design?.type === kind, run: () => setType(kind), disabled: () => !designView.design });
    groups.push(
      { label: t('Base.Group.QueryType'), items: [tbtn('select', 'base-query'), tbtn('maketable', 'base-q-maketable'), tbtn('append', 'base-q-append'), tbtn('update', 'base-q-update'), tbtn('delete', 'base-q-delete')] },
      { label: t('Base.Group.QuerySetup'), items: [
        { type: 'button', size: 'large', label: t('Base.Query.ShowTable'), icon: 'base-show-table', run: () => designView.showTable(), disabled: () => !designView.design },
        { type: 'stack', rows: [
          [{ type: 'button', showLabel: true, label: t('Base.Query.InsertColumn'), icon: 'base-insert-col', run: () => designView.insertColumn(), disabled: () => !designView.design }],
          [{ type: 'button', showLabel: true, label: t('Base.Query.DeleteColumn'), icon: 'base-delete-col', run: () => designView.deleteColumn(), disabled: () => !designView.design || designView.colSel === null }],
          [{ type: 'combo', label: t('Base.Query.Return'), width: 80, options: ['All', '5', '25', '100', '5%', '25%'], value: () => designView.design?.top || 'All', change: v => designView.setTop(v), disabled: () => !designView.design || !['select', 'maketable'].includes(designView.design.type) }],
        ] },
      ] },
      { label: t('Base.Group.ShowHide'), items: [
        { type: 'toggle', size: 'large', label: t('Base.Query.Totals'), icon: 'base-totals', active: () => !!designView.design?.totals, run: () => designView.toggleTotals(), disabled: () => !designView.design || !['select', 'maketable', 'append'].includes(designView.design.type) },
        { type: 'button', size: 'large', label: t('Base.Query.Properties'), icon: 'sliders', run: () => designView.properties(), disabled: () => !designView.design },
      ] },
    );
  }
  void type;
  return [{ id: 'qdesign', label: t('Base.Tab.QueryDesign'), groups }];
}

// ---- SQL view --------------------------------------------------------------------------------------------------------------------------

export class QuerySqlView extends BaseView {
  mount() {
    this.area = el('textarea', { class: 'mo-base-sql', spellcheck: 'false', 'aria-label': t('Base.View.Sql'), readonly: this.app.readOnly ? true : null });
    this.area.value = this.obj.sql;
    this.area.addEventListener('keydown', e => {
      if (e.key === 'Tab' && !e.shiftKey && !e.ctrlKey) { e.preventDefault(); this.area.setRangeText('  ', this.area.selectionStart, this.area.selectionEnd, 'end'); }
    });
    this.el.append(el('div', { class: 'mo-base-sqlwrap' }, this.area, el('div', { class: 'mo-base-sqlhint', text: t('Base.Query.SqlHint') })));
  }

  refresh() { if (this.obj && this.area.value !== this.obj.sql && document.activeElement !== this.area) this.area.value = this.obj.sql; }

  focus() { requestAnimationFrame(() => this.area.focus()); }

  /** Leaving the SQL view applies the text (after it was read without errors). */
  async leave() {
    const text = this.area.value.trim();
    if (this.app.readOnly || text === this.obj.sql.trim()) return true;
    if (text) {
      try { parseSql(text); } catch (e) {
        if (!(e instanceof BaseError)) throw e;
        await showProblem(e);
        this.selectError(e);
        return false;
      }
    }
    this.app.design(() => { this.obj.sql = text ? (text.endsWith(';') ? text : text + ';') : ''; });
    return true;
  }

  selectError(e) {
    const pos = Number(e.args?.[1]);
    this.area.focus();
    if (Number.isFinite(pos)) this.area.setSelectionRange(pos, Math.min(this.area.value.length, pos + 1));
  }

  async runQuery() {
    if (!(await this.leave())) return;
    runQuery(this.app, this.obj);
  }

  contextTabs() { return queryTabs(this.app, this, null); }
}

/** Run (!): a select query shows its datasheet, an action query does its work. */
async function runQuery(app, q) {
  let type = 'select';
  try { type = queryType(parseSql(q.sql)); } catch (e) { await showProblem(e); return; }
  if (type === 'select') app.open('queries', q.id, 'datasheet');
  else await runActionQuery(app, q);
}

QueryDatasheetView.prototype.runQuery = function runQueryHere() { if (this.ds) this.ds.commit().then(ok => { if (ok) { this.rs.load(); this.ds.reload(); } }); };

// ---- design view -------------------------------------------------------------------------------------------------------------------

const BOX_W = 170;
const BOX_H = 190;
const FIELD_H = 22;

export class QueryDesignView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.colSel = null;
    this.selectedJoin = null;
  }

  get q() { return this.obj; }

  fieldsOf(name) {
    const t2 = tableByName(this.db, name);
    if (t2) return t2.fields.map(f => f.name);
    const q2 = queryByName(this.db, name);
    if (q2 && lower(q2.name) !== lower(this.q.name)) {
      try { return runQueryObject(this.db, q2, { lang: this.app.lang, dry: true }).columns.map(c => c.name); } catch { return []; }
    }
    return [];
  }

  /** The design of the stored SQL (an empty design for a new query; null when the grid cannot show the SQL). */
  load() {
    const q = this.q;
    this.lastSql = q.sql;
    if (!q.sql.trim()) {
      const d = emptyDesign();
      d.tables = (q.layout?.tables || []).filter(b => sourceByName(this.db, b.name)).map(b => ({ name: b.name, alias: null }));
      d.joins = this.autoJoins(d.tables, []);
      return d;
    }
    try { return sqlToDesign(q.sql, this.db, n => this.fieldsOf(n)); } catch { return null; }
  }

  mount() {
    this.design = this.load();
    if (!this.design) {
      this.el.append(el('div', { class: 'mo-base-empty' }, el('div', { class: 'mo-base-empty__icon' }, icon('base-v-sql')), el('h2', { text: t('Base.Query.NoGridTitle') }), el('p', { text: t('Base.Err.NoGrid') }),
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => this.app.open('queries', this.tab.id, 'sql') }, icon('base-v-sql'), el('span', { text: t('Base.View.Sql') }))));
      return;
    }
    this.canvas = el('div', { class: 'mo-base-qd__tables', 'aria-label': t('Base.Query.TablesPane') });
    this.svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    this.svg.setAttribute('class', 'mo-base-lines');
    this.canvas.append(this.svg);
    this.gridBox = el('div', { class: 'mo-base-qd__grid' });
    const split = el('div', { class: 'mo-base-qd__split', role: 'separator', 'aria-orientation': 'horizontal' });
    this.root = el('div', { class: 'mo-base-qd' }, this.canvas, split, this.gridBox);
    const h = this.q.layout?.grid || 0;
    if (h) this.root.style.setProperty('--mo-qd-top', h + 'px');
    split.addEventListener('pointerdown', e => {
      const start = this.canvas.offsetHeight;
      drag(e, dy => this.root.style.setProperty('--mo-qd-top', Math.max(80, start + dy) + 'px'), dy => {
        if (!this.app.readOnly) { this.q.layout.grid = Math.max(80, Math.round(start + dy)); this.app.settingsChanged(); }
      });
    });
    this.canvas.addEventListener('pointerdown', e => { if (e.target === this.canvas || e.target === this.svg) { this.selectedJoin = null; this.drawLines(); } });
    this.canvas.addEventListener('keydown', e => { if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedJoin !== null) { e.preventDefault(); this.removeJoin(this.selectedJoin); } });
    this.canvas.tabIndex = 0;
    this.canvas.addEventListener('contextmenu', e => { if (e.target === this.canvas || e.target === this.svg) { e.preventDefault(); contextMenu({ x: e.clientX, y: e.clientY }, [{ label: t('Base.Query.ShowTable'), icon: 'base-show-table', run: () => this.showTable(), disabled: this.app.readOnly }, { label: t('Base.Query.Properties'), icon: 'sliders', run: () => this.properties() }]); } });
    this.el.append(this.root);
    this.renderTables();
    this.renderGrid();
    if (!this.design.tables.length && !this.app.readOnly && this.options?.created) setTimeout(() => this.showTable(), 50);
  }

  refresh() {
    if (!this.q || !this.design) return;
    if (this.q.sql === this.lastSql) { this.renderTables(); this.renderGrid(); return; }
    const d = this.load();
    if (!d) { this.el.replaceChildren(); this.mount(); return; }
    this.design = d;
    this.renderTables();
    this.renderGrid();
  }

  get ro() { return this.app.readOnly; }

  // ---- saving the design ------------------------------------------------------------------------------------------------------------

  /** Writes the SQL of the design (when the grid can be read; else the cell with the problem gets the focus). */
  async apply({ quiet = false } = {}) {
    if (this.ro) return true;
    const d = this.design;
    const boxes = this.boxes();
    let sql = '';
    const hasColumns = d.columns.some(c => String(c.field || '').trim()) || (d.type === 'delete' && d.tables.length);
    if (d.tables.length && hasColumns) {
      try {
        sql = toSql(designToStatement(d, this.db, n => this.fieldsOf(n)));
      } catch (e) {
        if (!(e instanceof BaseError)) throw e;
        if (quiet) return false;
        await showProblem(e);
        if (e.column !== undefined) this.focusCell(e.column, e.row !== undefined ? 'crit' + e.row : 'field');
        return false;
      }
    }
    if (sql === this.q.sql && JSON.stringify(boxes) === JSON.stringify(this.q.layout.tables)) return true;
    this.app.design(() => { this.q.sql = sql; this.q.layout.tables = boxes; });
    this.lastSql = sql;
    this.app.paintTab(this.tab);
    return true;
  }

  async leave() {
    if (!this.design) return true;
    const ok = await this.apply();
    if (!ok) return false;
    if (!this.q.sql && this.design.tables.length && !this.ro) {
      return ask(t('Base.Query.NoFieldsLeave'), { ok: t('Base.Query.LeaveAnyway'), cancel: t('Base.Query.Stay') });
    }
    return true;
  }

  boxes() {
    const pos = new Map((this.q.layout?.tables || []).map(b => [lower(b.name), b]));
    return this.design.tables.map((tb, i) => {
      const key = tb.alias || tb.name;
      const b = pos.get(lower(key));
      return { name: key, x: b?.x ?? 20 + i * (BOX_W + 50), y: b?.y ?? 16, w: b?.w ?? BOX_W, h: b?.h ?? BOX_H };
    });
  }

  // ---- tables and joins --------------------------------------------------------------------------------------------------------------

  aliasOf(tb) { return tb.alias || tb.name; }

  /** The join lines Access draws by itself: the relationships between the tables of the query. */
  autoJoins(tables, existing) {
    const joins = [...existing];
    for (const a of tables) for (const b of tables) {
      if (a === b) continue;
      const ta = tableByName(this.db, a.name);
      const tb = tableByName(this.db, b.name);
      if (!ta || !tb) continue;
      for (const r of this.db.relationships) {
        if (r.parent.table !== ta.id || r.child.table !== tb.id) continue;
        const pf = fieldById(ta, r.parent.field)?.name;
        const cf = fieldById(tb, r.child.field)?.name;
        if (!pf || !cf) continue;
        const left = this.aliasOf(a);
        const right = this.aliasOf(b);
        if (joins.some(j => (lower(j.left) === lower(left) && lower(j.right) === lower(right) && lower(j.leftField) === lower(pf) && lower(j.rightField) === lower(cf))
          || (lower(j.left) === lower(right) && lower(j.right) === lower(left) && lower(j.leftField) === lower(cf) && lower(j.rightField) === lower(pf)))) continue;
        joins.push({ left, leftField: pf, right, rightField: cf, kind: r.join === 2 ? 'LEFT' : r.join === 3 ? 'RIGHT' : 'INNER' });
      }
    }
    return joins;
  }

  async showTable() {
    if (this.ro) return;
    const list = el('div', { class: 'mo-base-showtable', role: 'listbox', 'aria-multiselectable': 'true' });
    const fill = which => {
      const items = [
        ...(which !== 'queries' ? this.db.tables.map(x => ({ name: x.name, icon: 'base-table' })) : []),
        ...(which !== 'tables' ? this.db.queries.filter(x => x.id !== this.q.id).map(x => { let ty = 'select'; try { ty = queryType(parseSql(x.sql)); } catch { ty = 'invalid'; } return ty === 'select' ? { name: x.name, icon: queryIcon('select') } : null; }).filter(Boolean) : []),
      ];
      list.replaceChildren(...items.map(i => {
        const b = el('button', { type: 'button', class: 'mo-base-showtable__item', role: 'option', 'aria-selected': 'false', 'data-name': i.name }, icon(i.icon), el('span', { text: i.name }));
        b.addEventListener('click', e => { if (!e.ctrlKey && !e.metaKey && !e.shiftKey) list.querySelectorAll('[aria-selected="true"]').forEach(x => x.setAttribute('aria-selected', 'false')); b.setAttribute('aria-selected', b.getAttribute('aria-selected') === 'true' ? 'false' : 'true'); });
        b.addEventListener('dblclick', () => { this.addTables([i.name]); });
        return b;
      }));
    };
    const tabs = radioGroup('mo-base-st', [{ value: 'tables', label: t('Base.Nav.Tables') }, { value: 'queries', label: t('Base.Nav.Queries') }, { value: 'both', label: t('Base.Query.Both') }], 'tables');
    tabs.classList.add('mo-base-radios--inline');
    tabs.addEventListener('change', () => fill(tabs.value()));
    fill('tables');
    const r = await dialog({
      title: t('Base.Query.ShowTable'), icon: 'base-show-table', body: el('div', { class: 'mo-base-form' }, tabs, list),
      buttons: [{ label: t('Base.Query.Add'), value: 'ok', kind: 'primary' }, { label: t('Editor.Close'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    const names = [...list.querySelectorAll('[aria-selected="true"]')].map(b => b.dataset.name);
    if (names.length) this.addTables(names);
  }

  addTables(names) {
    for (const name of names) {
      let alias = null;
      if (this.design.tables.some(x => lower(this.aliasOf(x)) === lower(name))) {
        let i = 1;
        while (this.design.tables.some(x => lower(this.aliasOf(x)) === lower(name + '_' + i))) i++;
        alias = name + '_' + i;
      }
      const boxes = this.boxes();
      const x = boxes.reduce((m, b) => Math.max(m, b.x + b.w + 50), 20);
      this.design.tables.push({ name, alias });
      this.q.layout.tables = [...boxes, { name: alias || name, x, y: 16, w: BOX_W, h: BOX_H }];
    }
    this.design.joins = this.autoJoins(this.design.tables, this.design.joins);
    this.renderTables();
    this.renderGrid();
    this.apply({ quiet: true }).then(ok => { if (!ok) this.app.settingsChanged(); });
  }

  removeTable(alias) {
    const key = lower(alias);
    this.design.tables = this.design.tables.filter(x => lower(this.aliasOf(x)) !== key);
    this.design.joins = this.design.joins.filter(j => lower(j.left) !== key && lower(j.right) !== key);
    this.design.columns = this.design.columns.filter(c => lower(c.table || '') !== key);
    this.q.layout.tables = this.q.layout.tables.filter(b => lower(b.name) !== key);
    this.renderTables();
    this.renderGrid();
    this.apply({ quiet: true });
  }

  renderTables() {
    for (const n of [...this.canvas.querySelectorAll('.mo-base-box')]) n.remove();
    const boxes = this.boxes();
    this.design.tables.forEach((tb, i) => {
      const b = boxes[i];
      const alias = this.aliasOf(tb);
      const src = sourceByName(this.db, tb.name);
      const fields = this.fieldsOf(tb.name);
      const pk = src?.kind === 'table' ? new Set(src.obj.primaryKey.map(id => fieldById(src.obj, id)?.name)) : new Set();
      const head = el('div', { class: 'mo-base-box__head', title: tb.alias ? tb.name + ' AS ' + tb.alias : tb.name }, icon(src?.kind === 'query' ? 'base-query' : 'base-table'), el('span', { text: alias }),
        this.ro ? null : iconButton('x', t('Base.Query.RemoveTable', alias), () => this.removeTable(alias), { class: 'mo-base-ibtn mo-base-box__x' }));
      const list = el('div', { class: 'mo-base-box__fields' });
      for (const name of ['*', ...fields]) {
        const row = el('div', { class: 'mo-base-box__field' + (pk.has(name) ? ' is-key' : ''), 'data-field': name, title: name }, pk.has(name) ? icon('key') : el('span', { class: 'mo-base-box__nokey' }), el('span', { text: name }));
        row.addEventListener('dblclick', () => this.addColumn(alias, name));
        row.addEventListener('pointerdown', e => this.dragField(e, alias, name, row));
        list.append(row);
      }
      list.addEventListener('scroll', () => this.drawLines());
      const box = el('div', { class: 'mo-base-box', 'data-alias': alias, style: { left: b.x + 'px', top: b.y + 'px', width: b.w + 'px', height: b.h + 'px' } }, head, list,
        el('span', { class: 'mo-base-box__resize', onpointerdown: e => this.resizeBox(e, alias) }));
      head.addEventListener('pointerdown', e => { if (!e.target.closest('button')) this.moveBox(e, alias); });
      head.addEventListener('contextmenu', e => { e.preventDefault(); contextMenu({ x: e.clientX, y: e.clientY }, [{ label: t('Base.Query.RemoveTable', alias), icon: 'x', run: () => this.removeTable(alias), disabled: this.ro }]); });
      this.canvas.append(box);
    });
    this.drawLines();
  }

  layoutOf(alias) { return this.q.layout.tables.find(b => lower(b.name) === lower(alias)); }

  moveBox(e, alias) {
    e.preventDefault();
    const boxes = this.boxes();
    this.q.layout.tables = boxes;
    const b = this.layoutOf(alias);
    const x0 = b.x;
    const y0 = b.y;
    const node = this.canvas.querySelector(`.mo-base-box[data-alias="${CSS.escape(alias)}"]`);
    drag(e, (dx, dy) => { b.x = Math.max(0, Math.round(x0 + dx)); b.y = Math.max(0, Math.round(y0 + dy)); node.style.left = b.x + 'px'; node.style.top = b.y + 'px'; this.drawLines(); },
      () => { if (!this.ro) { this.app.settingsChanged(); } });
  }

  resizeBox(e, alias) {
    e.preventDefault();
    e.stopPropagation();
    this.q.layout.tables = this.boxes();
    const b = this.layoutOf(alias);
    const w0 = b.w;
    const h0 = b.h;
    const node = this.canvas.querySelector(`.mo-base-box[data-alias="${CSS.escape(alias)}"]`);
    drag(e, (dx, dy) => { b.w = Math.max(110, Math.round(w0 + dx)); b.h = Math.max(80, Math.round(h0 + dy)); node.style.width = b.w + 'px'; node.style.height = b.h + 'px'; this.drawLines(); },
      () => { if (!this.ro) this.app.settingsChanged(); });
  }

  /** Dragging a field: onto a field of another table makes a join line, onto the grid a column. */
  dragField(e, alias, name, row) {
    if (this.ro || e.button !== 0) return;
    e.preventDefault();
    let ghost = null;
    drag(e, (dx, dy, ev) => {
      if (!ghost && Math.abs(dx) + Math.abs(dy) > 4) { ghost = el('div', { class: 'mo-base-dragghost', text: alias + '.' + name }); document.body.append(ghost); }
      if (ghost) { ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px'; }
    }, (dx, dy, ev) => {
      if (!ghost) return;
      ghost.remove();
      const target = document.elementFromPoint(ev.clientX, ev.clientY);
      const fieldRow = target?.closest('.mo-base-box__field');
      const box = target?.closest('.mo-base-box');
      if (fieldRow && box && box.dataset.alias !== alias && name !== '*' && fieldRow.dataset.field !== '*') {
        this.design.joins.push({ left: alias, leftField: name, right: box.dataset.alias, rightField: fieldRow.dataset.field, kind: 'INNER' });
        this.drawLines();
        this.apply({ quiet: true });
        return;
      }
      const cell = target?.closest('[data-col]');
      if (cell && this.gridBox.contains(cell)) this.addColumn(alias, name, Number(cell.dataset.col));
    });
    void row;
  }

  fieldPoint(alias, name, side) {
    const box = this.canvas.querySelector(`.mo-base-box[data-alias="${CSS.escape(alias)}"]`);
    if (!box) return null;
    const list = box.querySelector('.mo-base-box__fields');
    const row = [...list.children].find(r => lower(r.dataset.field) === lower(name));
    const b = this.layoutOf(alias) || { x: box.offsetLeft, y: box.offsetTop, w: box.offsetWidth };
    const head = box.querySelector('.mo-base-box__head').offsetHeight;
    let y = b.y + head + (row ? row.offsetTop - list.scrollTop + FIELD_H / 2 : 10);
    y = Math.max(b.y + head, Math.min(b.y + box.offsetHeight - 6, y));
    return { x: side === 'right' ? b.x + box.offsetWidth : b.x, y };
  }

  drawLines() {
    const svg = this.svg;
    const w = Math.max(this.canvas.scrollWidth, this.canvas.clientWidth);
    const h = Math.max(this.canvas.scrollHeight, this.canvas.clientHeight);
    svg.setAttribute('width', w);
    svg.setAttribute('height', h);
    svg.replaceChildren();
    this.design.joins.forEach((j, i) => {
      const lb = this.layoutOf(j.left) || this.boxes().find(b => lower(b.name) === lower(j.left));
      const rb = this.layoutOf(j.right) || this.boxes().find(b => lower(b.name) === lower(j.right));
      if (!lb || !rb) return;
      const leftIsLeft = lb.x + lb.w / 2 <= rb.x + rb.w / 2;
      const a = this.fieldPoint(j.left, j.leftField, leftIsLeft ? 'right' : 'left');
      const b = this.fieldPoint(j.right, j.rightField, leftIsLeft ? 'left' : 'right');
      if (!a || !b) return;
      const stub = 14;
      const ax = a.x + (leftIsLeft ? stub : -stub);
      const bx = b.x + (leftIsLeft ? -stub : stub);
      const d = `M${a.x},${a.y} L${ax},${a.y} L${bx},${b.y} L${b.x},${b.y}`;
      const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      g.setAttribute('class', 'mo-base-line' + (this.selectedJoin === i ? ' is-selected' : ''));
      const hit = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      hit.setAttribute('d', d);
      hit.setAttribute('class', 'mo-base-line__hit');
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', d);
      path.setAttribute('class', 'mo-base-line__path');
      g.append(hit, path);
      if (j.kind !== 'INNER') {
        // the arrow points at the table whose matching records are joined (Access draws it the same way)
        const toRight = j.kind === 'LEFT';
        const tip = toRight ? { x: b.x, y: b.y } : { x: a.x, y: a.y };
        const dir = (toRight ? leftIsLeft : !leftIsLeft) ? 1 : -1;
        const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        arrow.setAttribute('d', `M${tip.x - dir * 8},${tip.y - 4} L${tip.x},${tip.y} L${tip.x - dir * 8},${tip.y + 4}`);
        arrow.setAttribute('class', 'mo-base-line__arrow');
        g.append(arrow);
      }
      g.addEventListener('pointerdown', e => { e.stopPropagation(); this.selectedJoin = i; this.drawLines(); this.canvas.focus(); });
      g.addEventListener('dblclick', () => this.joinProperties(i));
      g.addEventListener('contextmenu', e => {
        e.preventDefault();
        this.selectedJoin = i;
        this.drawLines();
        contextMenu({ x: e.clientX, y: e.clientY }, [
          { label: t('Base.Query.JoinProperties'), icon: 'sliders', run: () => this.joinProperties(i), disabled: this.ro },
          { label: t('Base.Query.DeleteJoin'), icon: 'trash', danger: true, run: () => this.removeJoin(i), disabled: this.ro },
        ]);
      });
      svg.append(g);
    });
  }

  removeJoin(i) {
    if (this.ro) return;
    this.design.joins.splice(i, 1);
    this.selectedJoin = null;
    this.drawLines();
    this.apply({ quiet: true });
  }

  async joinProperties(i) {
    if (this.ro) return;
    const j = this.design.joins[i];
    const choice = radioGroup('mo-base-join', [
      { value: 'INNER', label: t('Base.Join.Inner') },
      { value: 'LEFT', label: t('Base.Join.Left', j.left, j.right) },
      { value: 'RIGHT', label: t('Base.Join.Right', j.right, j.left) },
    ], j.kind);
    const r = await dialog({
      title: t('Base.Query.JoinProperties'), icon: 'base-join', size: 'md',
      body: el('div', { class: 'mo-base-form' }, el('dl', { class: 'mat-kv mo-base-kv' }, el('dt', { text: t('Base.Join.LeftTable') }), el('dd', { text: j.left + '.' + j.leftField }), el('dt', { text: t('Base.Join.RightTable') }), el('dd', { text: j.right + '.' + j.rightField })), choice),
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    j.kind = choice.value();
    this.drawLines();
    this.apply({ quiet: true });
  }

  // ---- the design grid -----------------------------------------------------------------------------------------------------------

  rowsOf() {
    const d = this.design;
    const rows = ['field', 'table'];
    if (d.totals && ['select', 'maketable', 'append'].includes(d.type)) rows.push('total');
    if (d.type === 'update') rows.push('updateTo');
    if (d.type === 'delete') rows.push('del');
    if (d.type === 'select' || d.type === 'maketable' || d.type === 'append') rows.push('sort');
    if (d.type === 'append') rows.push('appendTo');
    if (d.type === 'select' || d.type === 'maketable') rows.push('show');
    const used = Math.max(0, ...d.columns.map(c => (c.criteria || []).length));
    const crit = Math.max(4, used + 2);
    for (let i = 0; i < crit; i++) rows.push('crit' + i);
    return rows;
  }

  rowLabel(r) {
    if (r.startsWith('crit')) return r === 'crit0' ? t('Base.Qbe.Criteria') : t('Base.Qbe.Or');
    return { field: t('Base.Qbe.Field'), table: t('Base.Qbe.Table'), total: t('Base.Qbe.Total'), sort: t('Base.Qbe.Sort'), show: t('Base.Qbe.Show'), appendTo: t('Base.Qbe.AppendTo'), updateTo: t('Base.Qbe.UpdateTo'), del: t('Base.Qbe.Delete') }[r];
  }

  renderGrid() {
    const d = this.design;
    const cols = [...d.columns];
    while (cols.length < d.columns.length + 6) cols.push(null);
    const rows = this.rowsOf();
    const table = el('table', { class: 'mo-base-qbe', role: 'grid', 'aria-label': t('Base.Query.Grid') });
    const bar = el('tr', { class: 'mo-base-qbe__bar' }, el('th', { class: 'mo-base-qbe__corner' }));
    cols.forEach((c, i) => {
      const th = el('th', { class: 'mo-base-qbe__colsel' + (this.colSel === i ? ' is-selected' : ''), 'data-col': i, title: t('Base.Query.SelectColumn') });
      th.addEventListener('click', () => { this.colSel = i; this.renderGrid(); this.app.ribbonRefresh(); });
      th.addEventListener('contextmenu', e => { e.preventDefault(); this.colSel = i; this.renderGrid(); this.columnMenu({ x: e.clientX, y: e.clientY }); });
      bar.append(th);
    });
    const body = el('tbody');
    const fieldList = 'mo-base-qbe-fields-' + this.tab.id;
    const options = [];
    for (const tb of d.tables) { const a = this.aliasOf(tb); options.push(a + '.*'); for (const f of this.fieldsOf(tb.name)) options.push(a + '.' + f); }
    const datalist = el('datalist', { id: fieldList }, ...options.map(o => el('option', { value: o })));
    for (const r of rows) {
      const tr = el('tr', {}, el('th', { scope: 'row', class: 'mo-base-qbe__label', text: this.rowLabel(r) }));
      cols.forEach((c, i) => tr.append(this.cell(r, c, i, fieldList)));
      body.append(tr);
    }
    table.append(el('thead', {}, bar), body);
    this.gridBox.replaceChildren(table, datalist);
  }

  cell(r, c, i, listId) {
    const d = this.design;
    const td = el('td', { 'data-col': i, 'data-row': r, class: this.colSel === i ? 'is-selected' : null });
    const ro = this.ro;
    const ensure = () => {
      if (!this.design.columns[i]) { while (this.design.columns.length < i) this.design.columns.push(emptyColumn()); this.design.columns[i] = emptyColumn(); }
      return this.design.columns[i];
    };
    const onFocus = () => { if (this.colSel !== i) { this.colSel = i; this.gridBox.querySelectorAll('.is-selected').forEach(n => n.classList.remove('is-selected')); this.gridBox.querySelectorAll(`[data-col="${i}"]`).forEach(n => n.classList.add('is-selected')); this.app.ribbonRefresh(); } };
    const keys = input => input.addEventListener('keydown', e => this.gridKey(e, r, i));
    switch (r) {
      case 'field': {
        const value = c ? (c.field === '*' ? (c.table ? c.table + '.*' : '*') : c.field) : '';
        const input = el('input', { class: 'mo-base-qbe__input', value, list: listId, 'aria-label': t('Base.Qbe.Field'), readonly: ro ? true : null, autocomplete: 'off', spellcheck: 'false' });
        input.addEventListener('focus', onFocus);
        input.addEventListener('change', () => this.setField(i, input.value));
        keys(input);
        td.append(input);
        break;
      }
      case 'table': {
        const s = el('select', { class: 'mo-base-qbe__input', 'aria-label': t('Base.Qbe.Table'), disabled: ro || !c ? true : null }, el('option', { value: '', text: '' }), ...d.tables.map(tb => el('option', { value: this.aliasOf(tb), text: this.aliasOf(tb) })));
        s.value = c?.table || '';
        s.addEventListener('focus', onFocus);
        s.addEventListener('change', () => { ensure().table = s.value; this.changed(); });
        keys(s);
        td.append(s);
        break;
      }
      case 'total': {
        const s = el('select', { class: 'mo-base-qbe__input', 'aria-label': t('Base.Qbe.Total'), disabled: ro || !c ? true : null }, ...TOTALS.map(x => el('option', { value: x, text: totalLabel(x) })));
        s.value = c?.total || 'Group By';
        s.addEventListener('focus', onFocus);
        s.addEventListener('change', () => { const col = ensure(); col.total = s.value; if (s.value === 'Where') col.show = false; this.changed(true); });
        keys(s);
        td.append(s);
        break;
      }
      case 'sort': {
        const s = el('select', { class: 'mo-base-qbe__input', 'aria-label': t('Base.Qbe.Sort'), disabled: ro || !c ? true : null },
          el('option', { value: '', text: '' }), el('option', { value: 'asc', text: t('Base.Qbe.Ascending') }), el('option', { value: 'desc', text: t('Base.Qbe.Descending') }));
        s.value = c?.sort || '';
        s.addEventListener('focus', onFocus);
        s.addEventListener('change', () => { ensure().sort = s.value; this.changed(); });
        keys(s);
        td.append(s);
        break;
      }
      case 'show': {
        const box = el('input', { type: 'checkbox', class: 'mo-base-qbe__check', 'aria-label': t('Base.Qbe.Show'), disabled: ro || !c ? true : null });
        box.checked = c ? c.show !== false : false;
        box.addEventListener('focus', onFocus);
        box.addEventListener('change', () => { ensure().show = box.checked; this.changed(); });
        keys(box);
        td.classList.add('is-center');
        td.append(box);
        break;
      }
      case 'appendTo': {
        const target = tableByName(this.db, d.target);
        const s = el('select', { class: 'mo-base-qbe__input', 'aria-label': t('Base.Qbe.AppendTo'), disabled: ro || !c ? true : null }, el('option', { value: '', text: '' }), ...(target?.fields || []).map(f => el('option', { value: f.name, text: f.name })));
        s.value = c?.appendTo || '';
        s.addEventListener('focus', onFocus);
        s.addEventListener('change', () => { ensure().appendTo = s.value; this.changed(); });
        keys(s);
        td.append(s);
        break;
      }
      case 'del': {
        const s = el('select', { class: 'mo-base-qbe__input', 'aria-label': t('Base.Qbe.Delete'), disabled: ro || !c ? true : null }, el('option', { value: 'From', text: t('Base.Qbe.From') }), el('option', { value: 'Where', text: t('Base.Qbe.Where') }));
        s.value = c?.del || (c?.field === '*' ? 'From' : 'Where');
        s.addEventListener('focus', onFocus);
        s.addEventListener('change', () => { ensure().del = s.value; this.changed(); });
        keys(s);
        td.append(s);
        break;
      }
      default: {
        const isUpdate = r === 'updateTo';
        const k = isUpdate ? null : Number(r.slice(4));
        const value = c ? (isUpdate ? c.updateTo || '' : c.criteria?.[k] || '') : '';
        const input = el('input', { class: 'mo-base-qbe__input', value, 'aria-label': this.rowLabel(r), readonly: ro ? true : null, autocomplete: 'off', spellcheck: 'false' });
        input.addEventListener('focus', onFocus);
        input.addEventListener('change', () => {
          const col = ensure();
          if (isUpdate) col.updateTo = input.value.trim();
          else { col.criteria = [...(col.criteria || [])]; while (col.criteria.length <= k) col.criteria.push(''); col.criteria[k] = input.value.trim(); }
          this.changed(r.startsWith('crit') && k >= (this.rowsOf().filter(x => x.startsWith('crit')).length - 2));
        });
        keys(input);
        td.append(input);
      }
    }
    return td;
  }

  gridKey(e, r, i) {
    const rows = this.rowsOf();
    const ri = rows.indexOf(r);
    const go = (nr, ni) => {
      const target = this.gridBox.querySelector(`td[data-col="${ni}"][data-row="${nr}"]`)?.firstElementChild;
      if (target) { e.preventDefault(); target.focus(); target.select?.(); }
    };
    const isSelect = e.target.tagName === 'SELECT';
    if (e.key === 'ArrowDown' && !isSelect) go(rows[ri + 1], i);
    else if (e.key === 'ArrowUp' && !isSelect) go(rows[ri - 1], i);
    else if (e.key === 'Enter' && !isSelect) { e.target.dispatchEvent(new Event('change')); go(rows[ri + 1] || rows[ri], i); }
    else if (e.key === 'F6') { e.preventDefault(); this.canvas.focus(); }
  }

  focusCell(i, r) {
    const target = this.gridBox.querySelector(`td[data-col="${i}"][data-row="${r}"]`)?.firstElementChild;
    target?.focus();
    target?.select?.();
  }

  /** The Field cell: "Table.Field", "Table.*", "Field" or "Alias: expression" (Access adds "Expr1: " to a plain expression). */
  setField(i, text) {
    const d = this.design;
    let value = text.trim();
    if (!value) {
      if (d.columns[i]) { d.columns.splice(i, 1); this.colSel = null; }
      this.changed(true);
      return;
    }
    const col = d.columns[i] || (() => { while (d.columns.length < i) d.columns.push(emptyColumn()); d.columns[i] = emptyColumn(); return d.columns[i]; })();
    const dotted = /^\[?([^[\].]+)\]?\.\[?([^[\]]+)\]?$/.exec(value);
    if (dotted && d.tables.some(tb => lower(this.aliasOf(tb)) === lower(dotted[1])) && !splitAlias(value).alias) {
      const tb = d.tables.find(x => lower(this.aliasOf(x)) === lower(dotted[1]));
      col.table = this.aliasOf(tb);
      col.field = dotted[2] === '*' ? '*' : this.fieldsOf(tb.name).find(f => lower(f) === lower(dotted[2])) || dotted[2];
    } else if (value === '*') {
      col.field = '*';
      col.table = col.table || (d.tables[0] ? this.aliasOf(d.tables[0]) : '');
    } else {
      const plain = /^\[?([^[\]"#()+\-*/&<>=,:]+)\]?$/.exec(value);
      const owner = plain ? d.tables.filter(tb => this.fieldsOf(tb.name).some(f => lower(f) === lower(plain[1].trim()))) : [];
      if (plain && owner.length >= 1) {
        col.field = this.fieldsOf(owner[0].name).find(f => lower(f) === lower(plain[1].trim()));
        col.table = col.table && owner.some(tb => lower(this.aliasOf(tb)) === lower(col.table)) ? col.table : this.aliasOf(owner[0]);
      } else {
        value = tidyFieldCell(value, d.columns.filter((_, k) => k !== i));
        col.field = value;
        col.table = '';
      }
    }
    if (col.show === undefined) col.show = true;
    this.changed(true);
  }

  addColumn(alias, name, at = null) {
    if (this.ro) return;
    const d = this.design;
    const col = { ...emptyColumn(), field: name, table: alias };
    if (at !== null && at < d.columns.length) d.columns.splice(at, 0, col);
    else d.columns.push(col);
    this.changed(true);
  }

  async changed(rerender = false) {
    const ok = await this.apply();
    if (rerender || !ok) this.renderGrid();
    this.app.ribbonRefresh();
  }

  insertColumn() {
    const at = this.colSel ?? this.design.columns.length;
    this.design.columns.splice(Math.min(at, this.design.columns.length), 0, emptyColumn());
    this.renderGrid();
  }

  deleteColumn() {
    if (this.colSel === null || this.colSel >= this.design.columns.length) return;
    this.design.columns.splice(this.colSel, 1);
    this.colSel = null;
    this.changed(true);
  }

  moveColumn(dir) {
    const i = this.colSel;
    const j = i + dir;
    const cols = this.design.columns;
    if (i === null || j < 0 || j >= cols.length || !cols[i]) return;
    [cols[i], cols[j]] = [cols[j], cols[i]];
    this.colSel = j;
    this.changed(true);
  }

  columnMenu(at) {
    contextMenu(at, [
      { label: t('Base.Query.InsertColumn'), icon: 'base-insert-col', run: () => this.insertColumn(), disabled: this.ro },
      { label: t('Base.Query.DeleteColumn'), icon: 'base-delete-col', danger: true, run: () => this.deleteColumn(), disabled: this.ro || !this.design.columns[this.colSel] },
      { label: t('Base.Query.MoveLeft'), icon: 'chevron-left', run: () => this.moveColumn(-1), disabled: this.ro || !this.colSel },
      { label: t('Base.Query.MoveRight'), icon: 'chevron-right', run: () => this.moveColumn(1), disabled: this.ro || this.colSel >= this.design.columns.length - 1 },
    ]);
  }

  async setType(kind) {
    const d = this.design;
    if (d.type === kind) return;
    if (kind === 'maketable' || kind === 'append') {
      const choices = kind === 'append' ? this.db.tables.map(x => ({ value: x.name, label: x.name })) : null;
      const input = choices ? selectInput(choices, d.target || choices[0]?.value) : textInput(d.target || uniqueName(this.db, ['tables', 'queries'], t('Base.Query.NewTableName')), { maxlength: 64 });
      const r = await dialog({
        title: t('Base.Query.Type.' + kind), icon: kind === 'append' ? 'base-q-append' : 'base-q-maketable',
        body: field(kind === 'append' ? t('Base.Query.AppendToTable') : t('Base.Query.NewTable'), input), onOpen: () => input.focus(),
        buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
      });
      if (r !== 'ok' || !input.value.trim()) return;
      d.target = input.value.trim();
      if (kind === 'append') {
        const target = tableByName(this.db, d.target);
        for (const c of d.columns) { const name = c.field === '*' ? '' : splitAlias(c.field).alias || c.field; c.appendTo = target?.fields.find(f => lower(f.name) === lower(name))?.name || ''; }
      }
    }
    if (kind === 'delete') {
      const first = d.tables[0];
      d.columns = d.columns.filter(c => c.field !== '*');
      if (first) d.columns.unshift({ ...emptyColumn(), field: '*', table: this.aliasOf(first), del: 'From' });
      for (const c of d.columns) if (c.field !== '*') c.del = 'Where';
      d.totals = false;
    }
    if (kind === 'update') d.totals = false;
    if (kind !== 'delete') d.columns = d.columns.filter(c => !(c.field === '*' && c.del === 'From' && kind === 'update'));
    d.type = kind;
    this.changed(true);
    this.app.paintTab(this.tab);
    this.app.nav.render();
  }

  setTop(v) {
    this.design.top = /^all$/i.test(v) ? '' : v;
    this.changed();
  }

  toggleTotals() {
    this.design.totals = !this.design.totals;
    if (this.design.totals) for (const c of this.design.columns) c.total ||= 'Group By';
    this.changed(true);
  }

  async properties() {
    const d = this.design;
    const q = this.q;
    const desc = el('textarea', { class: 'mat-input mat-textarea', rows: 2, maxlength: 1000, readonly: this.ro ? true : null });
    desc.value = q.description || '';
    const distinct = checkInput(t('Base.Query.UniqueValues'), d.distinct, { disabled: this.ro ? true : null });
    const top = textInput(d.top || '', { placeholder: t('Base.Query.All'), readonly: this.ro ? true : null });
    const r = await dialog({
      title: t('Base.Query.Properties'), icon: 'sliders', body: el('div', { class: 'mo-base-form' }, field(t('Base.Field.Description'), desc), distinct, field(t('Base.Query.TopValues'), top, t('Base.Query.TopHint'))),
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok' || this.ro) return;
    d.distinct = distinct.input.checked;
    d.top = top.value.trim();
    if (desc.value.trim() !== (q.description || '')) this.app.design(() => { q.description = desc.value.trim(); });
    this.changed(true);
  }

  async runQuery() {
    if (!(await this.apply())) return;
    if (!this.q.sql) { await showProblem({ code: 'NoColumns', args: [] }); return; }
    runQuery(this.app, this.q);
  }

  contextTabs() { return queryTabs(this.app, this, this); }

  status() {
    const d = this.design;
    if (!d) return [];
    return [t('Base.Query.Type.' + d.type)];
  }
}

function totalLabel(x) {
  return {
    'Group By': t('Base.Qbe.GroupBy'), Sum: t('Base.Total.Sum'), Avg: t('Base.Total.Avg'), Min: t('Base.Total.Min'), Max: t('Base.Total.Max'), Count: t('Base.Total.Count'),
    StDev: t('Base.Total.StDev'), Var: t('Base.Total.Var'), First: t('Base.Qbe.First'), Last: t('Base.Qbe.Last'), Expression: t('Base.Qbe.Expression'), Where: t('Base.Qbe.Where'),
  }[x] || x;
}

export { newObjectId, tableById, isUniqueField };
