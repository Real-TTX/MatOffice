/* MatOffice Base – the editor of desktop databases, like Access: the Navigation Pane on the left, the objects as document tabs in the work area,
   each with its views (Datasheet, Design, SQL, Form, Layout, Report, Print Preview), the ribbon (Home, Create, External Data, Database Tools and the
   contextual tabs of the active view), undo/redo of design changes (records are saved one by one, like Access) and the status bar.
   The content is the JSON documented in BaseHandler.cs; the engine (expr, sql, engine, qbe, model) is pure and tested with node. */
import { startEditor } from '../editor/frame.js';
import { t, el, icon, toast as frameToast, dialog } from '../editor/util.js';
import {
  ensureShape, emptyDb, OBJECT_KINDS, findObject, deleteObject, renameObject, duplicateObject, newObjectId, uniqueName, freeName,
  namespaceOf, nameProblem, nameTaken, compactDb, totalRecords, lower, tableById,
} from './model.js';
import { parseSql, queryType } from './sql.js';
import { NavPane } from './nav.js';
import { lang, showProblem, ask, field, textInput, selectInput, radioGroup, KIND_ICON, contextMenu, problemText } from './ui.js';
import { datasheetActive } from './view.js';
import { TableDatasheetView, TableDesignView } from './tableviews.js';
import { QueryDatasheetView, QuerySqlView, QueryDesignView, runActionQuery } from './queryviews.js';
import { RelationshipsView } from './relationships.js';
import { FormView, FormDesignView } from './forms.js';
import { ReportView, PreviewView, ReportDesignView, DocumenterView, printTab, printObject } from './reports.js';
import * as wizards from './wizards.js';
import * as external from './external.js';
import { thumbnail } from './thumb.js';

/** The views of each kind of object, in the order of the View menu. */
const VIEWS = {
  tables: [
    { id: 'datasheet', label: 'Base.View.Datasheet', icon: 'base-v-datasheet', make: (a, tab) => new TableDatasheetView(a, tab) },
    { id: 'design', label: 'Base.View.Design', icon: 'base-v-design', make: (a, tab) => new TableDesignView(a, tab) },
  ],
  queries: [
    { id: 'datasheet', label: 'Base.View.Datasheet', icon: 'base-v-datasheet', make: (a, tab) => new QueryDatasheetView(a, tab) },
    { id: 'sql', label: 'Base.View.Sql', icon: 'base-v-sql', make: (a, tab) => new QuerySqlView(a, tab) },
    { id: 'design', label: 'Base.View.Design', icon: 'base-v-design', make: (a, tab) => new QueryDesignView(a, tab) },
  ],
  forms: [
    { id: 'form', label: 'Base.View.Form', icon: 'base-v-form', make: (a, tab) => new FormView(a, tab) },
    { id: 'layout', label: 'Base.View.Layout', icon: 'base-v-layout', make: (a, tab) => new FormDesignView(a, tab, true) },
    { id: 'design', label: 'Base.View.Design', icon: 'base-v-design', make: (a, tab) => new FormDesignView(a, tab, false) },
  ],
  reports: [
    { id: 'report', label: 'Base.View.Report', icon: 'base-v-report', make: (a, tab) => new ReportView(a, tab) },
    { id: 'preview', label: 'Base.View.Preview', icon: 'base-v-preview', make: (a, tab) => new PreviewView(a, tab) },
    { id: 'design', label: 'Base.View.Design', icon: 'base-v-design', make: (a, tab) => new ReportDesignView(a, tab) },
  ],
  relationships: [{ id: 'relationships', label: 'Base.View.Relationships', icon: 'base-relationships', make: (a, tab) => new RelationshipsView(a, tab) }],
  documenter: [{ id: 'documenter', label: 'Base.View.Documenter', icon: 'base-documenter', make: (a, tab) => new DocumenterView(a, tab) }],
};

/** The undo history of design changes: snapshots of the content without the records (the records of the tables a change touched are kept). */
class DesignHistory {
  constructor(limit = 60) { this.limit = limit; this.past = []; this.future = []; }
  push(entry) { this.past.push(entry); if (this.past.length > this.limit) this.past.shift(); this.future = []; }
  get canUndo() { return this.past.length > 0; }
  get canRedo() { return this.future.length > 0; }
  /** Records of a table changed: snapshots holding its old records cannot come back without losing them. */
  invalidate(tableIds) {
    const hit = e => e.tables.some(id => tableIds.includes(id));
    let cut = -1;
    this.past.forEach((e, i) => { if (hit(e)) cut = i; });
    if (cut >= 0) this.past = this.past.slice(cut + 1);
    if (this.future.some(hit)) this.future = [];
  }
  clear() { this.past = []; this.future = []; }
}

class App {
  constructor() {
    this.db = emptyDb();
    this.tabs = [];
    this.active = null;
    this.version = 0;
    this.history = new DesignHistory();
  }

  get readOnly() { return this.ctx.readOnly; }
  get lang() { return lang(); }
  get activeView() { return this.active?.viewObj || null; }

  // ---- mount -------------------------------------------------------------------------------------------------------------------------

  mount(ctx) {
    this.ctx = ctx;
    this.db = ensureShape(ctx.doc.content || {});
    this.nav = new NavPane(this);
    this.tabsBar = el('div', { class: 'mo-base-tabs', role: 'tablist', 'aria-label': t('Base.Tabs') });
    this.stage = el('div', { class: 'mo-base-stage' });
    this.navToggle = el('button', { type: 'button', class: 'mo-base-navtoggle', 'aria-label': t('Base.Nav.Show'), 'data-tip': t('Base.Nav.Show'), onclick: () => this.toggleNav() }, icon('sidebar'));
    this.main = el('section', { class: 'mo-base-main' }, el('div', { class: 'mo-base-tabrow' }, this.navToggle, this.tabsBar), this.stage);
    this.scrim = el('div', { class: 'mo-base-scrim', onclick: () => this.toggleNav(false) });
    this.root = el('div', { class: 'mo-base' }, this.nav.el, this.scrim, this.main);
    ctx.work.replaceChildren(this.root);
    this.applyNavState();
    this.nav.render();
    this.buildRibbon();
    this.restoreTabs();
    if (!this.tabs.length) this.showEmpty();
    this.setStatus();
    document.addEventListener('keydown', e => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === 's' || e.key === 'S')) this.activeView?.leave?.({ save: true });
      if (e.key === 'F11' && !e.ctrlKey && !e.altKey) { e.preventDefault(); this.toggleNav(); }
    });
    window.addEventListener('beforeunload', () => this.saveTabs());
    ctx.session.addEventListener('saved', () => this.saveTabs());
  }

  isPhone() { return window.matchMedia('(max-width: 719px)').matches; }

  collapsedNav() { return this.isPhone() ? !this.root.classList.contains('is-nav-open') : !!this.db.ui.nav.hidden; }

  applyNavState() {
    this.root.classList.toggle('is-nav-hidden', !!this.db.ui.nav.hidden);
    this.root.style.setProperty('--mo-base-nav-w', (this.db.ui.nav.width || 240) + 'px');
  }

  toggleNav(force) {
    if (this.isPhone()) {
      const open = force ?? !this.root.classList.contains('is-nav-open');
      this.root.classList.toggle('is-nav-open', open);
      return;
    }
    this.db.ui.nav.hidden = force === undefined ? !this.db.ui.nav.hidden : !force;
    this.applyNavState();
    this.settingsChanged();
  }

  showEmpty() {
    this.stage.replaceChildren(el('div', { class: 'mo-base-empty' },
      el('div', { class: 'mo-base-empty__icon' }, icon('kind-base')),
      el('h2', { text: this.db.tables.length ? t('Base.Empty.Title') : t('Base.Empty.NewTitle') }),
      el('p', { text: this.db.tables.length ? t('Base.Empty.Text') : t('Base.Empty.NewText') }),
      this.readOnly ? null : el('div', { class: 'mat-btn-row mat-btn-row--plain' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => wizards.newTable(this) }, icon('base-table'), el('span', { text: t('Base.Create.Table') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary', onclick: () => external.importExcel(this) }, icon('base-excel'), el('span', { text: t('Base.Ext.ImportExcel') })))));
  }

  // ---- changes, history ----------------------------------------------------------------------------------------------------------

  changed() { this.version++; this.ctx.changed(); }

  /** A setting that is no design change (column widths, the collapsed groups): saved, not undoable. */
  settingsChanged() { this.ctx.changed(); }

  /** Records of tables were saved or deleted. */
  recordsChanged(tableIds) {
    this.history.invalidate(tableIds);
    this.changed();
  }

  snapshot(tables) {
    const copy = { ...this.db, tables: this.db.tables.map(tb => ({ ...tb, rows: tables.includes(tb.id) ? tb.rows : null })) };
    return { json: JSON.stringify(copy), tables };
  }

  restore(entry) {
    const s = JSON.parse(entry.json);
    for (const tb of s.tables) {
      const cur = tableById(this.db, tb.id);
      if (tb.rows === null) {
        tb.rows = cur ? cur.rows : [];
        if (cur) {
          tb.nextId = Math.max(tb.nextId || 1, cur.nextId || 1);
          for (const f of tb.fields) { const cf = cur.fields.find(x => x.id === f.id); if (f.type === 'autonumber' && cf?.next) f.next = Math.max(f.next || 1, cf.next); }
        }
      }
    }
    this.db = ensureShape(s);
  }

  /**
   * A design change: fn changes this.db; it can be undone. tables: ids of the tables whose records the change touches (kept in the snapshot).
   * Answers what fn answers.
   */
  design(fn, tables = []) {
    if (this.readOnly) return undefined;
    const before = this.snapshot(tables);
    const result = fn();
    if (result === false) return result;
    this.history.push(before);
    this.changed();
    this.nav.render();
    this.ribbonRefresh();
    return result;
  }

  /** For changes made step by step (dragging a control): the state before, then commitDesign(before) when the drag ends. */
  beginDesign(tables = []) { return this.readOnly ? null : this.snapshot(tables); }

  commitDesign(before) {
    if (!before || this.readOnly) return;
    if (before.json === this.snapshot(before.tables).json) return;
    this.history.push(before);
    this.changed();
    this.ribbonRefresh();
  }

  undo() {
    if (!this.history.canUndo || this.readOnly) return;
    const entry = this.history.past.pop();
    this.history.future.push(this.snapshot(entry.tables));
    this.restore(entry);
    this.afterRestore();
  }

  redo() {
    if (!this.history.canRedo || this.readOnly) return;
    const entry = this.history.future.pop();
    this.history.past.push(this.snapshot(entry.tables));
    this.restore(entry);
    this.afterRestore();
  }

  afterRestore() {
    this.changed();
    this.refreshAll();
  }

  /** Another content (a remote save, a conflict resolved with "theirs", undo): every view reads it again. */
  setContent(content) {
    this.db = ensureShape(content || {});
    this.history.clear();
    this.version++;
    this.refreshAll();
  }

  refreshAll() {
    for (const tab of [...this.tabs]) {
      if (tab.kind !== 'relationships' && tab.kind !== 'documenter' && !findObject(this.db, tab.kind, tab.id)) { this.closeTab(tab, true); continue; }
      tab.viewObj.version = this.version;
      tab.viewObj.refresh();
      this.paintTab(tab);
    }
    this.applyNavState();
    this.nav.render();
    this.rebuildRibbon();
    this.setStatus();
  }

  toast(message, tone = 'info') { frameToast(message, tone); }

  // ---- tabs and views ---------------------------------------------------------------------------------------------------------------

  viewsOf(kind, obj) {
    let list = VIEWS[kind] || [];
    if (kind === 'queries' && obj) {
      let type = 'select';
      try { type = queryType(parseSql(obj.sql)); } catch { type = 'invalid'; }
      if (type === 'invalid' || !obj.sql.trim()) list = [list[2], list[1]];
    }
    if (this.readOnly && (kind === 'forms' || kind === 'reports')) list = list.filter(v => v.id !== 'layout');
    return list;
  }

  /** Opens an object in its tab (or shows its tab), in a view (default: the first one). */
  async open(kind, id, viewId = null, options = {}) {
    const obj = kind === 'relationships' || kind === 'documenter' ? { id: kind, name: t(kind === 'relationships' ? 'Base.Rel.Title' : 'Base.Doc.Title') } : findObject(this.db, kind, id);
    if (!obj) return null;
    if (this.isPhone()) this.toggleNav(false);
    const key = kind + ':' + obj.id;
    let tab = this.tabs.find(x => x.key === key);
    if (!viewId && kind === 'queries' && !tab && !options.noRun) {
      let type = 'select';
      try { type = queryType(parseSql(obj.sql)); } catch { type = 'invalid'; }
      if (type !== 'select' && type !== 'invalid') { await runActionQuery(this, obj); return null; }
    }
    const views = this.viewsOf(kind, obj);
    const view = viewId && views.some(v => v.id === viewId) ? viewId : tab ? tab.view : views[0].id;
    if (tab) {
      if (tab.view !== view) await this.switchView(tab, view);
      this.activate(tab);
      return tab;
    }
    tab = { key, kind, id: obj.id, view, viewObj: null, pane: el('div', { class: 'mo-base-pane' }), btn: null };
    if (!this.tabs.length) this.stage.replaceChildren();
    this.tabs.push(tab);
    this.stage.append(tab.pane);
    this.makeView(tab, view, options);
    this.renderTabs();
    this.activate(tab);
    this.nav.render();
    return tab;
  }

  makeView(tab, viewId, options = {}) {
    const def = (VIEWS[tab.kind] || []).find(v => v.id === viewId) || VIEWS[tab.kind][0];
    tab.view = def.id;
    tab.viewObj = def.make(this, tab);
    tab.viewObj.options = options;
    tab.pane.replaceChildren(tab.viewObj.el);
    try {
      const r = tab.viewObj.mount();
      if (r && typeof r.then === 'function') r.then(() => { if (this.active === tab) { this.rebuildRibbon(); this.setStatus(); } }).catch(e => this.failView(tab, e));
    } catch (e) {
      this.failView(tab, e);
    }
  }

  failView(tab, e) {
    console.error(e);
    tab.pane.replaceChildren(el('div', { class: 'mo-base-empty' }, el('div', { class: 'mo-base-empty__icon' }, icon('alert')), el('p', { text: problemText(e) })));
  }

  async switchView(tab, viewId) {
    if (tab.view === viewId) return true;
    if (tab.viewObj && !(await tab.viewObj.leave())) return false;
    tab.viewObj?.destroy();
    this.makeView(tab, viewId);
    this.paintTab(tab);
    if (this.active === tab) {
      this.rebuildRibbon(true);
      this.setStatus();
      tab.viewObj.activate();
    }
    return true;
  }

  activate(tab) {
    if (this.active && this.active !== tab) this.active.pane.hidden = true;
    const changed = this.active !== tab;
    this.active = tab;
    tab.pane.hidden = false;
    for (const x of this.tabs) x.btn?.setAttribute('aria-selected', x === tab ? 'true' : 'false');
    for (const x of this.tabs) x.btn?.classList.toggle('is-active', x === tab);
    tab.btn?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    tab.viewObj.activate();
    this.rebuildRibbon(changed);
    this.setStatus();
    this.saveTabs();
  }

  async closeTab(tab, force = false) {
    if (!force && !(await tab.viewObj.leave())) return false;
    tab.viewObj.destroy();
    tab.pane.remove();
    const i = this.tabs.indexOf(tab);
    this.tabs.splice(i, 1);
    if (this.active === tab) {
      this.active = null;
      const next = this.tabs[Math.min(i, this.tabs.length - 1)];
      if (next) this.activate(next);
    }
    this.renderTabs();
    if (!this.tabs.length) { this.showEmpty(); this.rebuildRibbon(); this.setStatus(); }
    this.nav.render();
    this.saveTabs();
    return true;
  }

  tabTitle(tab) {
    if (tab.kind === 'relationships') return t('Base.Rel.Title');
    if (tab.kind === 'documenter') return t('Base.Doc.Title');
    return findObject(this.db, tab.kind, tab.id)?.name || '';
  }

  tabIcon(tab) {
    if (tab.kind === 'queries') {
      try { return this.nav.iconOf('queries', findObject(this.db, 'queries', tab.id)); } catch { return 'base-query'; }
    }
    return KIND_ICON[tab.kind];
  }

  renderTabs() {
    this.tabsBar.replaceChildren(...this.tabs.map(tab => {
      const close = el('button', { type: 'button', class: 'mo-base-tab__close', 'aria-label': t('Base.Tab.Close', this.tabTitle(tab)), onclick: e => { e.stopPropagation(); this.closeTab(tab); } }, icon('x'));
      const b = el('div', {
        class: 'mo-base-tab' + (tab === this.active ? ' is-active' : ''), role: 'tab', tabindex: '0', 'aria-selected': tab === this.active ? 'true' : 'false',
      }, icon(this.tabIcon(tab)), el('span', { class: 'mo-base-tab__name', text: this.tabTitle(tab) }), close);
      b.addEventListener('click', () => this.activate(tab));
      b.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.activate(tab); } });
      b.addEventListener('auxclick', e => { if (e.button === 1) this.closeTab(tab); });
      b.addEventListener('contextmenu', e => {
        e.preventDefault();
        contextMenu({ x: e.clientX, y: e.clientY }, [
          ...this.viewsOf(tab.kind, findObject(this.db, tab.kind, tab.id)).map(v => ({ label: t(v.label), icon: v.icon, checked: tab.view === v.id, run: () => this.open(tab.kind, tab.id, v.id) })),
          '-',
          { label: t('Base.Tab.Close', this.tabTitle(tab)), icon: 'x', run: () => this.closeTab(tab) },
          { label: t('Base.Tab.CloseAll'), icon: 'x', run: () => this.closeAll() },
        ]);
      });
      tab.btn = b;
      return b;
    }));
  }

  paintTab(tab) {
    if (!tab.btn) return;
    tab.btn.querySelector('.mo-base-tab__name').textContent = this.tabTitle(tab);
    tab.btn.querySelector('svg use')?.setAttribute('href', '#mat-i-' + this.tabIcon(tab));
  }

  async closeAll() {
    for (const tab of [...this.tabs]) if (!(await this.closeTab(tab))) return;
  }

  storageKey() { return 'mo.base.tabs.' + this.ctx.doc.api.split('?')[0]; }

  saveTabs() {
    try {
      localStorage.setItem(this.storageKey(), JSON.stringify({ tabs: this.tabs.map(x => ({ kind: x.kind, id: x.id, view: x.view })), active: this.active?.key || null }));
    } catch { /* private mode */ }
  }

  restoreTabs() {
    let state = null;
    try { state = JSON.parse(localStorage.getItem(this.storageKey()) || 'null'); } catch { state = null; }
    if (!state?.tabs?.length) return;
    for (const x of state.tabs.slice(0, 12)) {
      const exists = x.kind === 'relationships' || x.kind === 'documenter' || findObject(this.db, x.kind, x.id);
      if (!exists) continue;
      const view = x.view === 'layout' && this.readOnly ? null : x.view;
      if (x.kind === 'queries' && view === 'datasheet') {
        // a query with prompts opens in SQL view after a reload (no prompts at start-up)
        this.openQuiet(x.kind, x.id, 'design');
        continue;
      }
      this.openQuiet(x.kind, x.id, view);
    }
    const active = this.tabs.find(x => x.key === state.active);
    if (active) this.activate(active);
  }

  openQuiet(kind, id, view) {
    this.open(kind, id, view, { noRun: true });
  }

  // ---- objects -----------------------------------------------------------------------------------------------------------------------

  /** Adds an object (from Create) and opens it in a view. */
  addObject(kind, obj, view) {
    this.design(() => { this.db[kind].push(obj); });
    return this.open(kind, obj.id, view, { noRun: true, created: true });
  }

  renameObject(kind, id, name) {
    this.design(() => renameObject(this.db, kind, id, name));
    for (const tab of this.tabs) { if (tab.kind === kind && tab.id === id) this.paintTab(tab); else if (kind === 'tables' || kind === 'queries') tab.viewObj.refresh(); }
    this.setStatus();
  }

  async deleteObject(kind, id) {
    if (this.readOnly) return;
    const obj = findObject(this.db, kind, id);
    if (!obj) return;
    const question = { tables: t('Base.Nav.DeleteTable', obj.name), queries: t('Base.Nav.DeleteQuery', obj.name), forms: t('Base.Nav.DeleteForm', obj.name), reports: t('Base.Nav.DeleteReport', obj.name) }[kind];
    if (!(await ask(question, { title: t('Base.Nav.DeleteTitle'), danger: true, ok: t('Base.Nav.Delete') }))) return;
    if (kind === 'tables' && this.db.relationships.some(r => r.parent.table === id || r.child.table === id)) {
      if (!(await ask(t('Base.Nav.DeleteRelationships', obj.name), { danger: true, ok: t('Base.Yes') }))) return;
    }
    const tab = this.tabs.find(x => x.key === kind + ':' + id);
    if (tab) await this.closeTab(tab, true);
    this.design(() => deleteObject(this.db, kind, id), kind === 'tables' ? [id] : []);
    for (const x of this.tabs) x.viewObj.refresh();
    this.toast(t('Base.Nav.Deleted', obj.name), 'success');
  }

  /** Copy and paste of Access: a copy under a new name (a table with or without its records). */
  async duplicateObject(kind, id) {
    if (this.readOnly) return;
    const obj = findObject(this.db, kind, id);
    if (!obj) return;
    const name = textInput(freeName(this.db, namespaceOf(kind), t('Base.Nav.CopyOf', obj.name)), { maxlength: 64 });
    const what = kind === 'tables' ? radioGroup('mo-base-paste', [{ value: 'structure', label: t('Base.Nav.StructureOnly') }, { value: 'data', label: t('Base.Nav.StructureData') }], 'data') : null;
    const r = await dialog({
      title: t('Base.Nav.PasteAs'), icon: 'copy', body: el('div', { class: 'mo-base-form' }, field(t('Base.Nav.NewName'), name), what ? el('fieldset', { class: 'mo-base-fieldset' }, el('legend', { text: t('Base.Nav.PasteOptions') }), what) : null),
      onOpen: () => name.select(),
      validate: () => {
        const p = nameProblem(name.value.trim()) || (nameTaken(this.db, namespaceOf(kind), name.value.trim()) ? { code: 'NameTaken', args: [name.value.trim()] } : null);
        if (p) { showProblem(p); return false; }
        return true;
      },
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    const copy = this.design(() => {
      const c = duplicateObject(this.db, kind, id, name.value.trim());
      if (kind === 'tables' && what.value() === 'structure') { c.rows = []; c.nextId = 1; for (const f of c.fields) if (f.type === 'autonumber') f.next = 1; }
      return c;
    });
    if (copy) this.nav.select(kind + ':' + copy.id);
  }

  async objectProperties(kind, id) {
    const obj = findObject(this.db, kind, id);
    if (!obj) return;
    const desc = el('textarea', { class: 'mat-input mat-textarea', rows: 3, maxlength: 1000 });
    desc.value = obj.description || '';
    const info = kind === 'tables' ? t('Base.Nav.TableInfo', obj.fields.length, obj.rows.length) : '';
    const r = await dialog({ title: t('Base.Nav.PropertiesOf', obj.name), icon: 'info', body: el('div', { class: 'mo-base-form' }, info ? el('p', { text: info }) : null, field(t('Base.Field.Description'), desc)), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'ok') this.design(() => { obj.description = desc.value.trim(); });
  }

  exportExcel(kind, id) { return external.exportExcel(this, kind, id); }
  exportCsv(kind, id) { return external.exportCsv(this, kind, id); }
  printObject(kind, id) { return printObject(this, kind, id); }

  /** The object a Create or Export command works on: the active tab's, else the one selected in the pane. */
  currentObject(kinds = OBJECT_KINDS) {
    const a = this.active;
    if (a && kinds.includes(a.kind)) { const obj = findObject(this.db, a.kind, a.id); if (obj) return { kind: a.kind, obj }; }
    const sel = this.nav.selection();
    return sel && kinds.includes(sel.kind) ? sel : null;
  }

  async compact() {
    if (this.readOnly) return;
    for (const tab of this.tabs) if (!(await tab.viewObj.leave())) return;
    const fixed = this.design(() => compactDb(this.db), this.db.tables.map(x => x.id));
    this.refreshAll();
    this.toast(fixed ? t('Base.Tools.Compacted', fixed) : t('Base.Tools.CompactedNothing'), 'success');
  }

  async databaseProperties() {
    const currency = selectInput(['EUR', 'USD', 'GBP', 'CHF', 'JPY', 'SEK', 'NOK', 'DKK', 'PLN', 'CZK'].map(c => ({ value: c, label: c })), this.db.settings.currency);
    const stats = el('dl', { class: 'mat-kv mo-base-kv' },
      ...[[t('Base.Nav.Tables'), this.db.tables.length], [t('Base.Info.Records'), totalRecords(this.db)], [t('Base.Nav.Queries'), this.db.queries.length], [t('Base.Nav.Forms'), this.db.forms.length], [t('Base.Nav.Reports'), this.db.reports.length], [t('Base.Info.Relationships'), this.db.relationships.length]]
        .flatMap(([k, v]) => [el('dt', { text: k }), el('dd', { text: String(v) })]));
    const r = await dialog({ title: t('Base.Tools.Properties'), icon: 'settings', body: el('div', { class: 'mo-base-form' }, stats, field(t('Base.Tools.Currency'), currency, t('Base.Tools.CurrencyHint'))), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'ok' && !this.readOnly && currency.value !== this.db.settings.currency) { this.design(() => { this.db.settings.currency = currency.value; }); this.refreshAll(); }
  }

  // ---- commands of the Home tab (they go to the active view) --------------------------------------------------------------------

  can(cmd) { return !!this.activeView?.can(cmd); }

  async command(cmd, arg) {
    const v = this.activeView;
    if (!v) return;
    await v.run(cmd, arg);
    this.ribbonRefresh();
    this.setStatus();
  }

  active_(cmd) { return this.activeView?.active?.(cmd) ?? datasheetActive(this.activeView?.ds, cmd); }

  doUndo() {
    if (this.can('undoRecord')) { this.command('undoRecord'); return; }
    if (this.activeView?.can('undoDesign')) { this.command('undoDesign'); return; }
    this.undo();
  }

  // ---- ribbon --------------------------------------------------------------------------------------------------------------------------

  ribbonRefresh() { this.ctx?.ribbon.refresh(); }

  rebuildRibbon(selectContext = false) {
    this.buildRibbon();
    const ctxTabs = this.activeView?.contextTabs() || [];
    if (selectContext && ctxTabs.length && this.activeView?.selectContextTab !== false) this.ctx.ribbon.select(ctxTabs[0].id);
  }

  buildRibbon() {
    const c = (id, label, ic, more = {}) => ({ type: 'button', label, icon: ic, run: () => this.command(id, more.arg), disabled: () => !this.can(id), edit: false, ...more });
    const views = () => (this.active ? this.viewsOf(this.active.kind, findObject(this.db, this.active.kind, this.active.id)) : []);
    const viewIcon = () => views().find(v => v.id !== this.active?.view)?.icon || 'base-v-design';
    const home = {
      id: 'home', label: t('Base.Tab.Home'), groups: [
        { label: t('Base.Group.Views'), items: [{
          type: 'split', size: 'large', label: t('Base.View.View'), icon: 'base-v-design', edit: false, disabled: () => views().length < 2,
          run: () => { const list = views(); const next = list.find(v => v.id !== this.active.view); if (next) this.open(this.active.kind, this.active.id, next.id); },
          items: () => views().map(v => ({ label: t(v.label), icon: v.icon, checked: () => this.active?.view === v.id, edit: false, run: () => this.open(this.active.kind, this.active.id, v.id) })),
        }] },
        { label: t('Base.Group.Undo'), items: [{ type: 'stack', rows: [
          [{ type: 'button', label: t('Base.Undo'), icon: 'undo', key: 'Mod+Z', global: false, run: () => this.doUndo(), disabled: () => !(this.can('undoRecord') || this.history.canUndo || this.activeView?.can('undoDesign')) }],
          [{ type: 'button', label: t('Base.Redo'), icon: 'redo', key: 'Mod+Y', global: false, run: () => this.redo(), disabled: () => !this.history.canRedo }],
        ] }] },
        { label: t('Base.Group.Clipboard'), items: [
          c('paste', t('Base.Paste'), 'base-paste', { size: 'large', edit: true }),
          { type: 'stack', rows: [[c('cut', t('Base.Cut'), 'base-cut', { showLabel: true, edit: true })], [c('copy', t('Base.Copy'), 'copy', { showLabel: true })]] },
        ] },
        { label: t('Base.Group.SortFilter'), items: [
          c('filter', t('Base.Filter.Filter'), 'filter', { size: 'large' }),
          { type: 'stack', rows: [
            [c('sortAsc', t('Base.Sort.Ascending'), 'base-sort-asc', { showLabel: true })],
            [c('sortDesc', t('Base.Sort.Descending'), 'base-sort-desc', { showLabel: true })],
            [c('sortClear', t('Base.Sort.Remove'), 'base-sort-clear', { showLabel: true })],
          ] },
          { type: 'stack', rows: [
            [{ type: 'menu', label: t('Base.Filter.Selection'), icon: 'base-selection', edit: false, disabled: () => !this.can('filterSel'), items: () => ['equal', 'notEqual', 'contains', 'notContains'].map(k => ({ label: this.activeView?.selectionLabel?.(k) || k, edit: false, run: () => this.command('filterSel', k) })) }],
            [{ type: 'menu', label: t('Base.Filter.Advanced'), icon: 'base-filter-adv', edit: false, disabled: () => !this.can('filterAdvanced') && !this.can('filterClear'), items: [
              { label: t('Base.Filter.ClearAll'), icon: 'base-filter-off', edit: false, disabled: () => !this.can('filterClear'), run: () => this.command('filterClear') },
              { label: t('Base.Filter.AdvancedEdit'), icon: 'filter', edit: false, disabled: () => !this.can('filterAdvanced'), run: () => this.command('filterAdvanced') },
            ] }],
            [{ type: 'toggle', label: t('Base.Filter.Toggle'), icon: 'base-filter-toggle', showLabel: true, edit: false, active: () => this.active_('filterToggle'), disabled: () => !this.can('filterToggle'), run: () => this.command('filterToggle') }],
          ] },
        ] },
        { label: t('Base.Group.Records'), items: [
          c('refresh', t('Base.Rec.Refresh'), 'refresh', { size: 'large' }),
          { type: 'stack', rows: [
            [c('new', t('Base.Rec.New'), 'base-new-record', { showLabel: true, edit: true })],
            [c('save', t('Base.Rec.Save'), 'save', { showLabel: true, edit: true })],
            [c('delete', t('Base.Rec.Delete'), 'trash', { showLabel: true, edit: true })],
          ] },
          { type: 'stack', rows: [
            [{ type: 'toggle', label: t('Base.Rec.Totals'), icon: 'base-totals', showLabel: true, edit: false, active: () => this.active_('totals'), disabled: () => !this.can('totals'), run: () => this.command('totals') }],
            [{ type: 'menu', label: t('Base.Rec.More'), icon: 'more-horizontal', edit: false, disabled: () => !this.can('fieldWidth'), items: [
              { label: t('Base.Ds.FieldWidth'), icon: 'sliders', edit: false, run: () => this.command('fieldWidth') },
              { label: t('Base.Ds.Hide'), icon: 'eye-off', edit: false, run: () => this.command('hide') },
              { label: t('Base.Ds.Unhide'), icon: 'eye', edit: false, run: () => this.command('unhide') },
              '-',
              { label: t('Base.Ds.Freeze'), icon: 'base-freeze', edit: false, run: () => this.command('freeze') },
              { label: t('Base.Ds.Unfreeze'), icon: 'base-unfreeze', edit: false, run: () => this.command('unfreeze') },
            ] }],
          ] },
        ] },
        { label: t('Base.Group.Find'), items: [
          c('find', t('Base.Find.Find'), 'search', { size: 'large', key: 'Mod+F' }),
          { type: 'stack', rows: [
            [c('replace', t('Base.Find.Replace'), 'base-replace', { showLabel: true, key: 'Mod+H', edit: true })],
            [{ type: 'menu', label: t('Base.Find.GoTo'), icon: 'arrow-right', edit: false, disabled: () => !this.can('goto'), items: [
              { label: t('Base.Nav.First'), icon: 'base-first', edit: false, run: () => this.command('goto', 'first') },
              { label: t('Base.Nav.Previous'), icon: 'chevron-left', edit: false, run: () => this.command('goto', 'previous') },
              { label: t('Base.Nav.Next'), icon: 'chevron-right', edit: false, run: () => this.command('goto', 'next') },
              { label: t('Base.Nav.Last'), icon: 'base-last', edit: false, run: () => this.command('goto', 'last') },
              { label: t('Base.Nav.New'), icon: 'base-new-record', run: () => this.command('goto', 'new') },
            ] }],
            [{ type: 'menu', label: t('Base.Find.Select'), icon: 'base-select', edit: false, disabled: () => !this.can('selectAll'), items: [
              { label: t('Base.Find.SelectOne'), edit: false, run: () => this.command('select') },
              { label: t('Base.Find.SelectAll'), edit: false, run: () => this.command('selectAll') },
            ] }],
          ] },
        ] },
      ],
    };
    const big = (label, ic, run, more = {}) => ({ type: 'button', size: 'large', label, icon: ic, run, ...more });
    const small = (label, ic, run, more = {}) => ({ type: 'button', showLabel: true, label, icon: ic, run, ...more });
    const create = {
      id: 'create', label: t('Base.Tab.Create'), groups: [
        { label: t('Base.Group.Tables'), items: [big(t('Base.Create.Table'), 'base-table', () => wizards.newTable(this)), big(t('Base.Create.TableDesign'), 'base-table-design', () => wizards.newTableDesign(this))] },
        { label: t('Base.Group.Queries'), items: [big(t('Base.Create.QueryWizard'), 'base-query-wizard', () => wizards.queryWizard(this)), big(t('Base.Create.QueryDesign'), 'base-query-design', () => wizards.newQueryDesign(this))] },
        { label: t('Base.Group.Forms'), items: [
          big(t('Base.Create.Form'), 'base-form', () => wizards.quickForm(this, 'single'), { disabled: () => !this.currentObject(['tables', 'queries']) }),
          big(t('Base.Create.FormDesign'), 'base-form-design', () => wizards.blankForm(this, 'design')),
          { type: 'stack', rows: [
            [small(t('Base.Create.BlankForm'), 'base-blank-form', () => wizards.blankForm(this, 'layout'))],
            [small(t('Base.Create.FormWizard'), 'base-wizard', () => wizards.formWizard(this))],
            [{ type: 'menu', label: t('Base.Create.MoreForms'), icon: 'base-more-forms', disabled: () => !this.currentObject(['tables', 'queries']), items: [
              { label: t('Base.Create.MultipleItems'), icon: 'base-multiple-items', run: () => wizards.quickForm(this, 'continuous') },
              { label: t('Base.Create.DatasheetForm'), icon: 'base-v-datasheet', run: () => wizards.quickForm(this, 'datasheet') },
              { label: t('Base.Create.SplitForm'), icon: 'base-split-form', run: () => wizards.quickForm(this, 'split') },
            ] }],
          ] },
        ] },
        { label: t('Base.Group.Reports'), items: [
          big(t('Base.Create.Report'), 'base-report', () => wizards.quickReport(this), { disabled: () => !this.currentObject(['tables', 'queries']) }),
          big(t('Base.Create.ReportDesign'), 'base-report-design', () => wizards.blankReport(this)),
          { type: 'stack', rows: [
            [small(t('Base.Create.BlankReport'), 'base-blank-report', () => wizards.blankReport(this))],
            [small(t('Base.Create.ReportWizard'), 'base-wizard', () => wizards.reportWizard(this))],
          ] },
        ] },
      ],
    };
    const ext = {
      id: 'external', label: t('Base.Tab.External'), groups: [
        { label: t('Base.Group.Import'), items: [big(t('Base.Ext.Excel'), 'base-excel', () => external.importExcel(this)), big(t('Base.Ext.Text'), 'base-text-file', () => external.importText(this))] },
        { label: t('Base.Group.Export'), items: [
          big(t('Base.Ext.Excel'), 'base-excel-export', () => { const o = this.currentObject(['tables', 'queries']); if (o) external.exportExcel(this, o.kind, o.obj.id); }, { edit: false, disabled: () => !this.currentObject(['tables', 'queries']), tip: t('Base.Ext.ExportExcelTip') }),
          big(t('Base.Ext.Text'), 'base-text-export', () => { const o = this.currentObject(['tables', 'queries']); if (o) external.exportCsv(this, o.kind, o.obj.id); }, { edit: false, disabled: () => !this.currentObject(['tables', 'queries']), tip: t('Base.Ext.ExportTextTip') }),
          big(t('Base.Ext.Pdf'), 'base-pdf', () => { const o = this.currentObject(); if (o) printObject(this, o.kind, o.obj.id); }, { edit: false, disabled: () => !this.currentObject(), tip: t('Base.Ext.PdfTip') }),
          big(t('Base.Ext.AllExcel'), 'base-excel-all', () => external.exportAll(this), { edit: false, tip: t('Base.Ext.AllExcelTip') }),
        ] },
      ],
    };
    const tools = {
      id: 'tools', label: t('Base.Tab.Tools'), groups: [
        { label: t('Base.Group.Tools'), items: [big(t('Base.Tools.Compact'), 'base-compact', () => this.compact())] },
        { label: t('Base.Group.Relationships'), items: [big(t('Base.Tools.Relationships'), 'base-relationships', () => this.open('relationships', 'relationships'), { edit: false })] },
        { label: t('Base.Group.Analyze'), items: [big(t('Base.Tools.Documenter'), 'base-documenter', () => this.open('documenter', 'documenter'), { edit: false })] },
        { label: t('Base.Group.Database'), items: [big(t('Base.Tools.Properties'), 'settings', () => this.databaseProperties(), { edit: false })] },
      ],
    };
    const ctxTabs = this.activeView?.contextTabs() || [];
    this.ctx.ribbon.set([home, create, ext, tools, ...ctxTabs]);
  }

  // ---- status bar -------------------------------------------------------------------------------------------------------------------

  setStatus() {
    const tab = this.active;
    if (!tab) { this.ctx.status.left(t('Base.Status.Ready')); this.ctx.status.right(); return; }
    const def = (VIEWS[tab.kind] || []).find(v => v.id === tab.view);
    this.ctx.status.left(def ? t(def.label) : '', ...(tab.viewObj.status() || []));
    const views = this.viewsOf(tab.kind, findObject(this.db, tab.kind, tab.id));
    this.ctx.status.right(...(views.length > 1 ? [el('div', { class: 'mo-base-viewbtns', role: 'group', 'aria-label': t('Base.View.View') }, ...views.map(v => el('button', {
      type: 'button', class: 'mo-sbtn' + (v.id === tab.view ? ' is-active' : ''), 'aria-label': t(v.label), 'data-tip': t(v.label), 'aria-pressed': v.id === tab.view ? 'true' : 'false',
      onclick: () => this.open(tab.kind, tab.id, v.id),
    }, icon(v.icon))))] : []));
    if (tab.viewObj.zoom) this.ctx.status.zoom(tab.viewObj.zoom()); else this.ctx.status.zoomOptions = null;
  }

  // ---- the hooks of the frame -------------------------------------------------------------------------------------------------------

  info() {
    return el('dl', { class: 'mat-kv mo-base-kv' }, ...[
      [t('Base.Nav.Tables'), this.db.tables.length], [t('Base.Info.Records'), totalRecords(this.db).toLocaleString(this.lang)], [t('Base.Nav.Queries'), this.db.queries.length],
      [t('Base.Nav.Forms'), this.db.forms.length], [t('Base.Nav.Reports'), this.db.reports.length], [t('Base.Info.Relationships'), this.db.relationships.length],
    ].flatMap(([k, v]) => [el('dt', { text: k }), el('dd', { text: String(v) })]));
  }

  downloads() {
    const o = this.currentObject(['tables', 'queries']);
    if (!o) return [];
    return [
      { ext: 'CSV', label: t('Base.Ext.CsvOf', o.obj.name), run: () => external.exportCsv(this, o.kind, o.obj.id) },
      { ext: 'XLSX', label: t('Base.Ext.ExcelOf', o.obj.name), run: () => external.exportExcel(this, o.kind, o.obj.id) },
    ];
  }

  print() {
    if (this.active) printTab(this, this.active);
    else this.toast(t('Base.Print.Nothing'), 'info');
  }
}

const app = new App();

startEditor({
  mount: ctx => app.mount(ctx),
  getContent: () => app.db,
  setContent: content => app.setContent(content),
  thumbnail: () => thumbnail(app),
  print: () => app.print(),
  info: () => app.info(),
  downloads: () => app.downloads(),
});

export { app, uniqueName, newObjectId, lower };
