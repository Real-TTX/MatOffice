/* MatOffice Base – the Navigation Pane of Access: "All Access Objects" grouped by type (Tables, Queries, Forms, Reports) with a search box,
   collapsible groups, a shutter bar to hide it (a drawer on a phone), and a context menu per object: Open, Design View, Layout View, Rename,
   Delete, Duplicate, Export. */
import { t, el, icon } from '../editor/util.js';
import { contextMenu, KIND_ICON, queryIcon, iconButton, showProblem } from './ui.js';
import { OBJECT_KINDS, lower, nameProblem, nameTaken, namespaceOf } from './model.js';
import { queryType, parseSql } from './sql.js';

const KIND_LABEL = { tables: 'Base.Nav.Tables', queries: 'Base.Nav.Queries', forms: 'Base.Nav.Forms', reports: 'Base.Nav.Reports' };

export class NavPane {
  constructor(app) {
    this.app = app;
    this.filter = '';
    this.only = null;
    this.selected = null;
    this.el = el('aside', { class: 'mo-base-nav', 'aria-label': t('Base.Nav.Title') });
    this.list = el('div', { class: 'mo-base-nav__list', role: 'tree', 'aria-label': t('Base.Nav.Title') });
    this.search = el('input', { type: 'search', class: 'mat-input mo-base-nav__search', placeholder: t('Base.Nav.SearchPlaceholder'), 'aria-label': t('Base.Nav.Search') });
    this.search.addEventListener('input', () => { this.filter = this.search.value.trim().toLowerCase(); this.render(); });
    this.title = el('button', { type: 'button', class: 'mo-base-nav__title', 'aria-haspopup': 'menu', onclick: e => this.categoryMenu(e.currentTarget) });
    this.head = el('div', { class: 'mo-base-nav__head' }, this.title,
      iconButton('chevron-left', t('Base.Nav.Hide'), () => app.toggleNav(), { class: 'mo-base-ibtn mo-base-nav__shutter' }));
    this.el.append(this.head, el('div', { class: 'mo-base-nav__searchbox' }, icon('search'), this.search), this.list);
    this.list.addEventListener('keydown', e => this.onKey(e));
  }

  get ui() { return this.app.db.ui.nav; }

  categoryMenu(anchor) {
    contextMenu(anchor, [
      { heading: t('Base.Nav.Category') },
      { label: t('Base.Nav.All'), checked: !this.only, run: () => { this.only = null; this.render(); } },
      ...OBJECT_KINDS.map(k => ({ label: t(KIND_LABEL[k]), icon: KIND_ICON[k], checked: this.only === k, run: () => { this.only = k; this.render(); } })),
    ]);
  }

  render() {
    const db = this.app.db;
    this.title.replaceChildren(el('span', { text: this.only ? t(KIND_LABEL[this.only]) : t('Base.Nav.All') }), icon('chevron-down'));
    const collapsed = new Set(this.ui.collapsed || []);
    const groups = [];
    for (const kind of OBJECT_KINDS) {
      if (this.only && this.only !== kind) continue;
      const items = db[kind].filter(o => !this.filter || o.name.toLowerCase().includes(this.filter))
        .slice().sort((a, b) => a.name.localeCompare(b.name, this.app.lang, { sensitivity: 'base' }));
      const open = !collapsed.has(kind) || !!this.filter;
      const head = el('button', {
        type: 'button', class: 'mo-base-nav__group', 'aria-expanded': open ? 'true' : 'false',
        onclick: () => this.toggleGroup(kind),
      }, el('span', { text: t(KIND_LABEL[kind]) }), icon(open ? 'chevron-up' : 'chevron-down'));
      const list = el('div', { class: 'mo-base-nav__items', role: 'group', hidden: open ? null : true });
      for (const o of items) list.append(this.item(kind, o));
      if (!items.length && open) list.append(el('div', { class: 'mo-base-nav__empty', text: this.filter ? t('Base.Nav.NoMatch') : t('Base.Nav.None') }));
      groups.push(el('div', { class: 'mo-base-nav__section' }, head, list));
    }
    this.list.replaceChildren(...groups);
  }

  iconOf(kind, o) {
    if (kind !== 'queries') return KIND_ICON[kind];
    try { return queryIcon(queryType(parseSql(o.sql))); } catch { return 'base-query'; }
  }

  item(kind, o) {
    const key = kind + ':' + o.id;
    const open = this.app.tabs.some(x => x.key === key);
    const b = el('button', {
      type: 'button', class: 'mo-base-nav__item' + (this.selected === key ? ' is-selected' : '') + (open ? ' is-open' : ''), role: 'treeitem',
      'data-key': key, title: o.description || o.name,
    }, icon(this.iconOf(kind, o)), el('span', { class: 'mo-base-nav__name', text: o.name }));
    b.addEventListener('click', () => {
      this.select(key);
      if (window.matchMedia('(pointer: coarse)').matches || this.app.isPhone()) this.app.open(kind, o.id);
    });
    b.addEventListener('dblclick', () => this.app.open(kind, o.id));
    b.addEventListener('contextmenu', e => { e.preventDefault(); this.select(key); this.objectMenu(kind, o, { x: e.clientX, y: e.clientY }); });
    return b;
  }

  select(key) {
    this.selected = key;
    for (const b of this.list.querySelectorAll('.mo-base-nav__item')) b.classList.toggle('is-selected', b.dataset.key === key);
    this.app.ribbonRefresh();
  }

  /** The object selected in the pane: { kind, obj } (export and the Create buttons use it). */
  selection() {
    if (!this.selected) return null;
    const [kind, id] = this.selected.split(':');
    const obj = this.app.db[kind]?.find(o => o.id === id);
    return obj ? { kind, obj } : null;
  }

  toggleGroup(kind) {
    const set = new Set(this.ui.collapsed || []);
    if (set.has(kind)) set.delete(kind); else set.add(kind);
    this.ui.collapsed = [...set];
    this.app.settingsChanged();
    this.render();
  }

  onKey(e) {
    const items = [...this.list.querySelectorAll('.mo-base-nav__item')];
    const i = items.indexOf(document.activeElement);
    if (e.key === 'ArrowDown' && i < items.length - 1) { e.preventDefault(); items[i + 1].focus(); this.select(items[i + 1].dataset.key); }
    else if (e.key === 'ArrowUp' && i > 0) { e.preventDefault(); items[i - 1].focus(); this.select(items[i - 1].dataset.key); }
    else if (e.key === 'Enter' && i >= 0) { e.preventDefault(); const [kind, id] = items[i].dataset.key.split(':'); this.app.open(kind, id); }
    else if (e.key === 'F2' && i >= 0) { e.preventDefault(); const [kind, id] = items[i].dataset.key.split(':'); this.rename(kind, id); }
    else if (e.key === 'Delete' && i >= 0) { e.preventDefault(); const [kind, id] = items[i].dataset.key.split(':'); this.app.deleteObject(kind, id); }
  }

  objectMenu(kind, o, at) {
    const app = this.app;
    const ro = app.readOnly;
    const views = app.viewsOf(kind, o);
    const items = [
      { label: t('Base.Nav.Open'), icon: 'folder-open', run: () => app.open(kind, o.id) },
      ...views.slice(1).map(v => ({ label: t(v.label), icon: v.icon, run: () => app.open(kind, o.id, v.id) })),
      '-',
      { label: t('Base.Nav.Rename'), icon: 'edit', run: () => this.rename(kind, o.id), disabled: ro },
      { label: t('Base.Nav.Duplicate'), icon: 'copy', run: () => app.duplicateObject(kind, o.id), disabled: ro },
      { label: t('Base.Nav.Delete'), icon: 'trash', danger: true, run: () => app.deleteObject(kind, o.id), disabled: ro },
      '-',
    ];
    if (kind === 'tables' || kind === 'queries') {
      items.push(
        { label: t('Base.Ext.ExportExcel'), icon: 'base-excel', run: () => app.exportExcel(kind, o.id) },
        { label: t('Base.Ext.ExportText'), icon: 'base-text-file', run: () => app.exportCsv(kind, o.id) },
      );
    }
    items.push({ label: t('Base.Ext.ExportPdf'), icon: 'base-pdf', run: () => app.printObject(kind, o.id) });
    if (kind === 'tables' || kind === 'queries') items.push('-', { label: t('Base.Nav.Properties'), icon: 'info', run: () => app.objectProperties(kind, o.id), disabled: ro });
    contextMenu(at, items);
  }

  /** Renames in the pane itself (like Access: the name becomes a text box). */
  rename(kind, id) {
    if (this.app.readOnly) return;
    const o = this.app.db[kind].find(x => x.id === id);
    if (!o) return;
    if (this.app.collapsedNav()) this.app.toggleNav();
    this.only = null;
    this.filter = '';
    this.search.value = '';
    this.render();
    const b = this.list.querySelector(`[data-key="${kind}:${id}"]`);
    if (!b) return;
    const input = el('input', { class: 'mo-base-nav__rename', value: o.name, 'aria-label': t('Base.Nav.Rename'), maxlength: 64 });
    b.querySelector('.mo-base-nav__name').replaceWith(input);
    input.focus();
    input.select();
    let done = false;
    const finish = async commit => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      if (commit && name !== o.name) {
        const problem = nameProblem(name) || (nameTaken(this.app.db, namespaceOf(kind), name, id) && lower(name) !== lower(o.name) ? { code: 'NameTaken', args: [name] } : null);
        if (problem) { await showProblem(problem); done = false; input.focus(); input.select(); return; }
        this.app.renameObject(kind, id, name);
      }
      this.render();
      this.list.querySelector(`[data-key="${kind}:${id}"]`)?.focus();
    };
    input.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
    input.addEventListener('click', e => e.stopPropagation());
    input.addEventListener('dblclick', e => e.stopPropagation());
  }
}
