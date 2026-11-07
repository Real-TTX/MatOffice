/* MatOffice Sheets – the AutoFilter: buttons in the header row of a range; the menu of a column sorts, searches, ticks the values to show
   (blanks too) or applies a text/number condition. Rows that do not match are hidden (stored as hidden rows, so Excel sees them too). */
import { el, icon, t } from '../editor/util.js';
import { openPopover, closePopover } from '../editor/frame.js';
import { parseArea, areaName, colName } from './engine/address.js';
import { compareText, parseNumber, sortCompare } from './engine/values.js';

export const CONDITIONS = ['eq', 'ne', 'gt', 'ge', 'lt', 'le', 'begins', 'ends', 'contains', 'notContains'];

export class Filters {
  constructor(app) { this.app = app; }

  get sheet() { return this.app.sheet; }

  /** Data → Filter: on (for the data around the active cell, or the selection) or off. */
  toggle() {
    const app = this.app;
    const sheet = this.sheet;
    if (sheet.filter) {
      app.commands.change('filter off', () => {
        app.wb.setSheetProp(sheet, 'filter', null);
        app.wb.touchProp(sheet, 'rows');
        sheet.rows.setFiltered(new Set());
      }, { protect: false });
      return;
    }
    const a = app.sel.isSingle ? app.commands.currentRegion() : app.commands.clip(app.sel.range);
    if (a.r1 === a.r2 && app.calc.valueAt(sheet.id, a.r1, a.c1) === null) { app.ctx.toast(t('Sheets.Filter.NoData'), 'warning'); return; }
    app.commands.change('filter on', () => app.wb.setSheetProp(sheet, 'filter', { range: areaName(a) }), { protect: false });
  }

  /** The rows of the data of the filter (the range grows with data typed right below it). */
  dataArea(sheet = this.sheet) {
    const a = sheet.filterArea();
    if (!a) return null;
    const calc = this.app.calc;
    let r2 = a.r2;
    const has = r => { for (let c = a.c1; c <= a.c2; c++) { const v = calc.valueAt(sheet.id, r, c); if (v !== null && v !== '') return true; } return false; };
    while (r2 + 1 < 1048576 && has(r2 + 1)) r2++;
    return { ...a, r2 };
  }

  /** The text of a cell as the filter compares it (what is shown). */
  textAt(sheet, r, c) { return this.app.look.cell(sheet, r, c, 11).text.trim(); }

  /** Hides the rows that do not match (called after a change of the criteria and when a sheet is opened). */
  apply(sheet = this.sheet) {
    if (!sheet.filter) { if (sheet.rows.filtered.size) sheet.rows.setFiltered(new Set()); return; }
    const a = this.dataArea(sheet);
    if (!a) return;
    const hidden = new Set();
    const cols = Object.entries(sheet.filter.cols || {});
    if (cols.length) {
      for (let r = a.r1 + 1; r <= a.r2; r++) {
        for (const [letters, spec] of cols) {
          const c = parseArea(letters + '1')?.c1;
          if (c === undefined || c < a.c1 || c > a.c2) continue;
          if (!this.matches(sheet, r, c, spec)) { hidden.add(r); break; }
        }
      }
    }
    sheet.rows.setFiltered(hidden);
  }

  matches(sheet, r, c, spec) {
    const text = this.textAt(sheet, r, c);
    if (spec.values) {
      const blank = text === '';
      if (!spec.values.some(v => (v === '' ? blank : compareText(v, text) === 0))) return false;
    }
    if (spec.cond) {
      const value = this.app.calc.valueAt(sheet.id, r, c);
      const a = test(spec.cond.op, value, text, spec.cond.v);
      if (!spec.cond.op2) return a;
      const b = test(spec.cond.op2, value, text, spec.cond.v2);
      return spec.cond.or ? a || b : a && b;
    }
    return true;
  }

  /** Sets (or with null removes) the criterion of a column. */
  setColumn(col, spec) {
    const app = this.app;
    const sheet = this.sheet;
    const cols = { ...(sheet.filter.cols || {}) };
    if (spec) cols[colName(col)] = spec; else delete cols[colName(col)];
    const a = this.dataArea(sheet);
    app.commands.change('filter', () => {
      app.wb.setSheetProp(sheet, 'filter', { ...sheet.filter, range: areaName(a), cols });
      app.wb.touchProp(sheet, 'rows');
      this.apply(sheet);
    }, { protect: false });
  }

  clearAll() {
    const app = this.app;
    const sheet = this.sheet;
    if (!sheet.filter) return;
    app.commands.change('filter', () => {
      app.wb.setSheetProp(sheet, 'filter', { range: sheet.filter.range });
      app.wb.touchProp(sheet, 'rows');
      this.apply(sheet);
    }, { protect: false });
  }

  reapply() {
    const app = this.app;
    const sheet = this.sheet;
    if (!sheet.filter) return;
    app.commands.change('filter', () => { app.wb.touchProp(sheet, 'rows'); this.apply(sheet); }, { protect: false });
  }

  /** Filter by the value of the active cell (context menu). */
  byValue() {
    const app = this.app;
    const sheet = this.sheet;
    const { row, col } = app.sel.active;
    if (!sheet.filter) this.toggle();
    if (!sheet.filter) return;
    const a = sheet.filterArea();
    if (row <= a.r1) return;
    this.setColumn(col, { values: [this.textAt(sheet, row, col)] });
  }

  /** The menu of a filter button. */
  open(col, point) {
    const app = this.app;
    const sheet = this.sheet;
    const a = this.dataArea(sheet);
    if (!a) return;
    const spec = sheet.filter.cols?.[colName(col)] || null;
    // the anchor: the button of the column
    const r = app.grid.cellRect(a.r1, col);
    const host = app.grid.root.getBoundingClientRect();
    const anchor = el('div', { class: 'mo-sheets-anchor' });
    anchor.style.left = Math.round(host.left + r.x + r.w - 18) + 'px';
    anchor.style.top = Math.round(host.top + r.y + r.h) + 'px';
    document.body.append(anchor);
    void point;
    // the values of the column (what is shown), without the rows hidden by the other columns
    const others = Object.entries(sheet.filter.cols || {}).filter(([k]) => k !== colName(col));
    const values = new Map();
    let blanks = false;
    for (let row = a.r1 + 1; row <= a.r2; row++) {
      if (sheet.rows.hidden.has(row)) continue;
      if (others.some(([letters, s]) => !this.matches(sheet, row, parseArea(letters + '1').c1, s))) continue;
      const text = this.textAt(sheet, row, col);
      if (text === '') { blanks = true; continue; }
      const k = text.toLowerCase();
      if (!values.has(k)) values.set(k, { text, value: app.calc.valueAt(sheet.id, row, col) });
    }
    const list = [...values.values()].sort((x, y) => sortCompare(x.value, y.value) || compareText(x.text, y.text));
    const chosen = new Set(spec?.values ? spec.values.map(v => v.toLowerCase()) : [...values.keys(), ...(blanks ? [''] : [])]);
    const readOnly = app.readOnly;

    const pop = openPopover(anchor, close => {
      const box = el('div', { class: 'mo-sheets-filter' });
      const action = (label, iconName, run, disabled = false) => el('button', { type: 'button', class: 'mo-menu__item', disabled: disabled || readOnly ? true : null, onclick: () => { close(); run(); } }, icon(iconName), el('span', { class: 'mo-menu__label', text: label }));
      box.append(
        action(t('Sheets.SortAZ'), 'sheets-sort-az', () => this.sortColumn(col, false)),
        action(t('Sheets.SortZA'), 'sheets-sort-za', () => this.sortColumn(col, true)),
        el('div', { class: 'mo-menu__sep' }),
        action(t('Sheets.Filter.ClearColumn'), 'sheets-filter-clear', () => this.setColumn(col, null), !spec),
        action(t('Sheets.Filter.Condition'), 'filter', () => this.conditionDialog(col, spec?.cond)),
        el('div', { class: 'mo-menu__sep' }),
      );
      const search = el('input', { type: 'search', class: 'mat-input mo-sheets-filter__search', placeholder: t('Sheets.Filter.Search'), 'aria-label': t('Sheets.Filter.Search') });
      const items = el('div', { class: 'mo-sheets-filter__list' });
      const boxes = [];
      const allBox = el('input', { type: 'checkbox' });
      const row = (label, k) => {
        const cb = el('input', { type: 'checkbox' });
        cb.checked = chosen.has(k);
        cb.addEventListener('change', () => { if (cb.checked) chosen.add(k); else chosen.delete(k); syncAll(); });
        boxes.push({ cb, k, label });
        return el('label', { class: 'mat-check mo-sheets-filter__item' }, cb, el('span', { class: 'mat-check__text', text: label }));
      };
      const syncAll = () => {
        const visible = boxes.filter(b => !b.row.hidden);
        allBox.checked = visible.every(b => b.cb.checked);
        allBox.indeterminate = !allBox.checked && visible.some(b => b.cb.checked);
      };
      allBox.addEventListener('change', () => { for (const b of boxes) if (!b.row.hidden) { b.cb.checked = allBox.checked; if (allBox.checked) chosen.add(b.k); else chosen.delete(b.k); } syncAll(); });
      items.append(el('label', { class: 'mat-check mo-sheets-filter__item mo-sheets-filter__all' }, allBox, el('span', { class: 'mat-check__text', text: t('Sheets.Filter.All') })));
      for (const v of list.slice(0, 5000)) { const node = row(v.text, v.text.toLowerCase()); boxes[boxes.length - 1].row = node; items.append(node); }
      if (blanks) { const node = row(t('Sheets.Filter.Blanks'), ''); boxes[boxes.length - 1].row = node; items.append(node); }
      syncAll();
      search.addEventListener('input', () => {
        const q = search.value.trim().toLowerCase();
        for (const b of boxes) b.row.hidden = !!q && !b.label.toLowerCase().includes(q);
        syncAll();
      });
      const ok = el('button', { type: 'button', class: 'mat-btn mat-btn--primary mat-btn--sm', disabled: readOnly ? true : null }, el('span', { text: t('Sheets.Ok') }));
      const cancel = el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm' }, el('span', { text: t('Sheets.Cancel') }));
      ok.addEventListener('click', () => {
        close();
        const q = search.value.trim().toLowerCase();
        let picked = boxes.filter(b => b.cb.checked && (!q || b.label.toLowerCase().includes(q))).map(b => (b.k === '' ? '' : values.get(b.k).text));
        const all = picked.length === boxes.length;
        this.setColumn(col, all && !spec?.cond ? null : { ...(spec?.cond ? { cond: spec.cond } : {}), ...(all ? {} : { values: picked }) });
        picked = null;
      });
      cancel.addEventListener('click', () => close());
      box.append(search, items, el('div', { class: 'mat-btn-row mat-btn-row--plain mo-sheets-filter__buttons' }, ok, cancel));
      return box;
    });
    setTimeout(() => anchor.remove(), 0);
    pop?.querySelector('input[type=search]')?.focus();
  }

  sortColumn(col, desc) {
    const app = this.app;
    const a = this.dataArea();
    app.commands.sortBy(a, true, [{ col, desc }]);
    this.reapply();
  }

  /** A condition (one or two, And/Or) for a column. */
  async conditionDialog(col, cond) {
    const app = this.app;
    const ops = CONDITIONS.map(op => ({ value: op, label: t('Sheets.Cond.' + op) }));
    const select = (value, name) => el('select', { class: 'mat-input mat-select', name }, ...[{ value: '', label: '' }, ...ops].map(o => el('option', { value: o.value, text: o.label, selected: o.value === value ? true : null })));
    const op1 = select(cond?.op || 'eq', 'op1');
    const v1 = el('input', { class: 'mat-input', name: 'v1', value: cond?.v ?? '' });
    const op2 = select(cond?.op2 || '', 'op2');
    const v2 = el('input', { class: 'mat-input', name: 'v2', value: cond?.v2 ?? '' });
    const and = el('input', { type: 'radio', name: 'join', value: 'and', checked: !cond?.or ? true : null });
    const or = el('input', { type: 'radio', name: 'join', value: 'or', checked: cond?.or ? true : null });
    const body = el('div', { class: 'mo-sheets-form' },
      el('p', { class: 'mo-muted', text: t('Sheets.Filter.ConditionText', this.textAt(this.sheet, this.sheet.filterArea().r1, col) || colName(col)) }),
      el('div', { class: 'mo-sheets-form__row' }, op1, v1),
      el('div', { class: 'mat-choices mat-checks--inline mo-sheets-form__row' },
        el('label', { class: 'mat-choice' }, and, el('span', { class: 'mat-choice__text', text: t('Sheets.Filter.And') })),
        el('label', { class: 'mat-choice' }, or, el('span', { class: 'mat-choice__text', text: t('Sheets.Filter.Or') }))),
      el('div', { class: 'mo-sheets-form__row' }, op2, v2));
    const r = await app.ctx.dialog({ title: t('Sheets.Filter.Condition'), icon: 'filter', body, buttons: [{ label: t('Sheets.Ok'), value: 'ok', kind: 'primary' }, { label: t('Sheets.Cancel'), value: 'cancel' }], onOpen: () => v1.focus() });
    if (r !== 'ok' || !op1.value) return;
    const next = { op: op1.value, v: v1.value };
    if (op2.value) { next.op2 = op2.value; next.v2 = v2.value; next.or = or.checked; }
    this.setColumn(col, { cond: next });
    closePopover();
  }
}

/** One condition of a filter on a cell (numbers compare as numbers, texts without case). */
function test(op, value, text, operand) {
  const s = String(operand ?? '');
  const n = parseNumber(s);
  const num = typeof value === 'number' && n !== null;
  const cmp = num ? Math.sign(value - n) : compareText(text, s);
  const lower = text.toLowerCase();
  const o = s.toLowerCase();
  switch (op) {
    case 'eq': return num ? cmp === 0 : wild(s) ? wild(s).test(text) : cmp === 0;
    case 'ne': return num ? cmp !== 0 : wild(s) ? !wild(s).test(text) : cmp !== 0;
    case 'gt': return cmp > 0;
    case 'ge': return cmp >= 0;
    case 'lt': return cmp < 0;
    case 'le': return cmp <= 0;
    case 'begins': return lower.startsWith(o);
    case 'ends': return lower.endsWith(o);
    case 'contains': return lower.includes(o);
    case 'notContains': return !lower.includes(o);
    default: return true;
  }
}

function wild(s) {
  if (!/[*?]/.test(s)) return null;
  return new RegExp('^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
}
