/* MatOffice Base – what every view of an object (a document tab) has: the object it shows, refresh after a change elsewhere, leaving (commit
   the record, apply the design), the commands of the Home tab, contextual ribbon tabs and the texts of the status bar. */
import { el } from '../editor/util.js';

export class BaseView {
  constructor(app, tab) {
    this.app = app;
    this.tab = tab;
    this.el = el('div', { class: 'mo-base-view mo-base-view--' + tab.view });
    this.version = app.version;
  }

  get db() { return this.app.db; }
  get obj() { return this.app.db[this.tab.kind]?.find(o => o.id === this.tab.id) || null; }

  /** Builds the view into this.el. */
  mount() {}

  destroy() {}

  /** The content changed elsewhere (another view, undo, a remote save): read it again. */
  refresh() {}

  /** The tab is shown again. */
  activate() {
    if (this.version !== this.app.version) {
      this.version = this.app.version;
      this.refresh();
    }
    this.focus();
  }

  focus() {}

  /** Before another view or closing: false keeps the person here (a record that cannot be saved). */
  async leave() { return true; }

  /** More texts for the status bar ("Record 3 of 12", "Filtered"). */
  status() { return []; }

  can() { return false; }

  run() {}

  /** Contextual tabs of the ribbon (Table Design, Query Design, Form Layout …). */
  contextTabs() { return []; }
}

/** The Home commands of a datasheet (tables, queries, the datasheet of a split form). */
export function datasheetCan(ds, cmd) {
  if (!ds) return false;
  const ro = ds.readOnly;
  switch (cmd) {
    case 'new': return ds.canAdd;
    case 'save': return !!ds.pending;
    case 'delete': return !ro && ds.rs.canDelete && (!!ds.sel || !!ds.currentRow);
    case 'undoRecord': return ds.dirty;
    case 'cut': return !ro;
    case 'paste': return !ro || ds.canAdd;
    case 'copy': case 'refresh': case 'totals': case 'find': case 'selectAll': case 'select': case 'fieldWidth': case 'hide': case 'unhide': case 'freeze': case 'unfreeze':
    case 'sortAsc': case 'sortDesc': case 'filter': case 'filterSel': case 'filterAdvanced': return !!ds.cols.length;
    case 'sortClear': return !!ds.s.sort?.length;
    case 'filterToggle': return !!ds.s.filter;
    case 'filterClear': return !!ds.s.filter;
    case 'replace': return !ro;
    case 'goto': return true;
    default: return false;
  }
}

export function datasheetActive(ds, cmd) {
  if (cmd === 'totals') return !!ds?.s.totalsOn;
  if (cmd === 'filterToggle') return !!(ds?.s.filterOn && ds.s.filter);
  return false;
}

export async function datasheetRun(ds, cmd, arg) {
  switch (cmd) {
    case 'new': return ds.moveTo(ds.newIndex, 0, { focus: true });
    case 'save': await ds.commit(); return ds.focusEditor(false);
    case 'delete': return ds.deleteRecords();
    case 'undoRecord': return ds.undo();
    case 'refresh': if (await ds.commit()) ds.reload(); return undefined;
    case 'totals': return ds.toggleTotals();
    case 'sortAsc': return ds.sortBy(ds.currentCol, false);
    case 'sortDesc': return ds.sortBy(ds.currentCol, true);
    case 'sortClear': return ds.clearSort();
    case 'filter': return ds.columnMenu(ds.cur.c, ds.head.children[ds.cur.c + 1] || ds.head);
    case 'filterSel': return ds.filterBySelection(arg);
    case 'filterToggle': return ds.toggleFilter();
    case 'filterClear': return ds.clearFilters();
    case 'filterAdvanced': return ds.editFilter();
    case 'find': return ds.openFind(false);
    case 'replace': return ds.openFind(true);
    case 'goto': {
      const r = { first: 0, previous: ds.cur.r - 1, next: ds.cur.r + 1, last: ds.view.length - 1, new: ds.newIndex }[arg];
      return r >= 0 ? ds.moveTo(r, ds.cur.c, { focus: true }) : undefined;
    }
    case 'selectAll': return ds.selectAll();
    case 'select': return ds.selectRecords(ds.cur.r, false);
    case 'cut': return ds.cut();
    case 'copy': return ds.copy();
    case 'paste': return ds.paste();
    case 'fieldWidth': return ds.fieldWidth();
    case 'hide': return ds.hideColumn();
    case 'unhide': return ds.unhideColumns();
    case 'freeze': return ds.freezeColumn();
    case 'unfreeze': return ds.unfreezeAll();
    default: return undefined;
  }
}
