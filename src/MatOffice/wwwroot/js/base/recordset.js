/* MatOffice Base – recordsets: the records a datasheet or a form shows, from a table (editable, with the checks of model.js) or from a query
   (editable where the query is updatable, like Access: one table, no totals). A subform's recordset is a table or query filtered by its link field. */
import {
  tableById, queryById, saveRecord, planDelete, applyDelete, defaultValues, lookupOptions, lookupMap, engineValue, keyOf, parseInput, fieldById, lower,
} from './model.js';
import { runQueryObject } from './engine.js';
import { formatValue, formatDate, DateVal } from './values.js';
import { ruleHolds } from './criteria.js';
import { compile, BaseError } from './expr.js';

class Recordset {
  constructor(app) {
    this.app = app;
    this.columns = [];
    this.rows = [];
    this.maps = new Map();
  }

  get db() { return this.app.db; }
  get lang() { return this.app.lang; }

  /** The value of a column of a row, as expressions see it (dates are DateVal). */
  get(row, col) { return row.values[col.key] ?? null; }

  /** The text of a value on the screen (formats, lookups). */
  display(col, v) {
    if (v === null || v === undefined) return '';
    const map = this.mapOf(col);
    if (map) {
      const shown = map.get(keyOf(v));
      if (shown !== undefined) return shown;
    }
    const f = col.field;
    return formatValue(v, { type: col.type, format: col.format ?? f?.format ?? '', decimals: col.decimals ?? f?.decimals ?? null, lang: this.lang, currency: this.db.settings?.currency || 'EUR' });
  }

  /** The text a person edits (dates without the long formats). */
  editText(col, v) {
    if (v === null || v === undefined) return '';
    if (col.type === 'datetime' || v instanceof DateVal) return formatDate(v, this.lang, '');
    if (col.type === 'number' || col.type === 'currency' || col.type === 'autonumber') {
      if (typeof v === 'number' && /percent/i.test(col.field?.format || '')) return this.display(col, v);
      return typeof v === 'number' ? formatValue(v, { lang: this.lang }) : String(v);
    }
    return this.display(col, v);
  }

  mapOf(col) {
    if (!col.field?.lookup || col.field.lookup.kind !== 'table') return null;
    if (!this.maps.has(col.key)) this.maps.set(col.key, lookupMap(this.db, col.field, this.lang));
    return this.maps.get(col.key);
  }

  options(col) { return col.field?.lookup ? lookupOptions(this.db, col.field, this.lang) : []; }

  /** Reads what a person typed into a column: { ok, value } (a stored value) or { ok: false, code }. */
  parse(col, text) {
    if (!col.field) return { ok: false, code: 'NotUpdatable' };
    return parseInput(col.field, text, this.lang);
  }

  /** The validation rule of the field of a column (checked when the field is left). */
  checkRule(col, stored) {
    const f = col.field;
    if (!f?.validationRule) return null;
    try {
      if (!ruleHolds(f.validationRule, engineValue(f, stored), compile, { lang: this.lang })) return { code: 'Validation', args: [f.validationText || '', f.validationRule, f.name] };
    } catch (e) {
      if (!(e instanceof BaseError)) throw e;
    }
    return null;
  }

  /** The stored value as expressions see it. */
  toEngine(col, stored) { return col.field ? engineValue(col.field, stored) : stored; }

  invalidate() { this.maps.clear(); }
}

/** The records of a table (fixed: the values of a subform link, { fieldId: value }). */
export class TableRecordset extends Recordset {
  constructor(app, tableId, { fixed = null } = {}) {
    super(app);
    this.kind = 'table';
    this.tableId = tableId;
    this.fixed = fixed;
    this.load();
  }

  get table() { return tableById(this.db, this.tableId); }
  get name() { return this.table?.name || ''; }
  get canEdit() { return !this.app.readOnly && !!this.table; }
  get canAdd() { return this.canEdit; }
  get canDelete() { return this.canEdit; }

  load() {
    this.invalidate();
    const t = this.table;
    if (!t) { this.columns = []; this.rows = []; return; }
    this.columns = t.fields.map(f => ({
      key: f.id, skey: f.id, name: f.name, caption: f.caption || f.name, type: f.type, field: f, editable: f.type !== 'autonumber', align: alignOf(f.type),
    }));
    const fixed = this.fixed ? Object.entries(this.fixed) : null;
    const rows = fixed ? t.rows.filter(r => fixed.every(([k, v]) => keyOf(r.v[k]) === keyOf(v) && keyOf(v) !== null)) : t.rows;
    this.rows = rows.map(r => this.wrap(r));
  }

  wrap(r) {
    const values = {};
    for (const c of this.columns) values[c.key] = engineValue(c.field, r.v[c.key] ?? null);
    return { id: r.id, values, base: r };
  }

  defaults() {
    const t = this.table;
    const d = defaultValues(this.db, t, this.lang);
    if (this.fixed) Object.assign(d, this.fixed);
    const out = {};
    for (const c of this.columns) out[c.key] = d[c.key] ?? null;
    return out;
  }

  /** Saves the changes of a record (stored values by column key); a new one when id is null. */
  save(id, changes) {
    const t = this.table;
    const values = { ...changes };
    if (id === null && this.fixed) Object.assign(values, this.fixed);
    const r = saveRecord(this.db, t, id, values, this.lang);
    if (!r.ok) return r;
    this.app.recordsChanged([t.id]);
    const wrapped = this.wrap(r.row);
    const i = this.rows.findIndex(x => x.id === r.row.id);
    if (i >= 0) this.rows[i] = wrapped; else this.rows.push(wrapped);
    return { ok: true, row: wrapped };
  }

  planRemove(ids) { return planDelete(this.db, this.table, ids); }

  remove(plan) {
    applyDelete(this.db, plan.plan);
    this.app.recordsChanged([...plan.plan.keys()]);
    this.load();
  }
}

/** The records of a query (params: the answers of its prompts). Editable columns come from one table. */
export class QueryRecordset extends Recordset {
  constructor(app, queryId, params = {}, { fixed = null } = {}) {
    super(app);
    this.kind = 'query';
    this.queryId = queryId;
    this.params = params;
    this.fixed = fixed;
    this.error = null;
    this.load();
  }

  get query() { return queryById(this.db, this.queryId); }
  get name() { return this.query?.name || ''; }
  get canEdit() { return !this.app.readOnly && !!this.result?.updatable; }
  get canAdd() { return this.canEdit && !!this.result?.base; }
  get canDelete() { return this.canEdit && !!this.result?.base; }

  load() {
    this.invalidate();
    this.error = null;
    try {
      this.result = runQueryObject(this.db, this.query, { params: this.params, lang: this.lang });
    } catch (e) {
      if (!(e instanceof BaseError)) throw e;
      this.error = e;
      this.result = { columns: [], rows: [], updatable: false, base: null, sources: [] };
    }
    const res = this.result;
    this.columns = res.columns.map((c, i) => ({
      key: i, skey: c.name, name: c.name, caption: c.field && lower(c.field.name) === lower(c.name) && c.field.caption ? c.field.caption : c.name,
      type: c.type, field: c.field, editable: c.editable, source: c.source, align: alignOf(c.type),
    }));
    let rows = res.rows.map((r, i) => ({ id: i, values: { ...r.v }, refs: r.refs }));
    if (this.fixed) {
      const fixed = Object.entries(this.fixed).map(([name, v]) => [this.columns.find(c => lower(c.name) === lower(name)), v]);
      rows = rows.filter(r => fixed.every(([c, v]) => c && keyOf(r.values[c.key]) === keyOf(v) && keyOf(v) !== null));
    }
    this.rows = rows;
    this.nextKey = res.rows.length;
  }

  defaults() {
    const base = this.result.base;
    const out = {};
    if (!base) return out;
    const d = defaultValues(this.db, base, this.lang);
    for (const c of this.columns) if (c.field && d[c.field.id] !== undefined) out[c.key] = d[c.field.id];
    if (this.fixed) for (const [name, v] of Object.entries(this.fixed)) { const c = this.columns.find(x => lower(x.name) === lower(name)); if (c) out[c.key] = v; }
    return out;
  }

  save(id, changes) {
    const res = this.result;
    const byTable = new Map();
    for (const [key, v] of Object.entries(changes)) {
      const col = this.columns[Number(key)];
      if (!col?.field || col.source === null || col.source === undefined) continue;
      const tableId = res.sources[col.source]?.table;
      if (!tableId) continue;
      if (!byTable.has(col.source)) byTable.set(col.source, { tableId, values: {} });
      byTable.get(col.source).values[col.field.id] = v;
    }
    const row = id === null ? null : this.rows.find(r => r.id === id);
    if (id === null) {
      const base = res.base;
      if (!base) return { ok: false, error: { code: 'NotUpdatable', args: [] } };
      const values = byTable.get(0)?.values || {};
      if (this.fixed) for (const [name, v] of Object.entries(this.fixed)) { const f = base.fields.find(x => lower(x.name) === lower(name)); if (f) values[f.id] = v; }
      const r = saveRecord(this.db, base, null, values, this.lang);
      if (!r.ok) return r;
      this.app.recordsChanged([base.id]);
      const fresh = { id: this.nextKey++, values: {}, refs: [r.row] };
      for (const c of this.columns) fresh.values[c.key] = c.field && c.source === 0 ? engineValue(c.field, r.row.v[c.field.id] ?? null) : null;
      this.rows.push(fresh);
      return { ok: true, row: fresh };
    }
    for (const [idx, part] of byTable) {
      const t = tableById(this.db, part.tableId);
      const baseRow = row.refs?.[idx];
      if (!t || !baseRow) return { ok: false, error: { code: 'NotUpdatable', args: [] } };
      const r = saveRecord(this.db, t, baseRow.id, part.values, this.lang);
      if (!r.ok) return r;
      this.app.recordsChanged([t.id]);
      for (const c of this.columns) if (c.source === Number(idx) && c.field) row.values[c.key] = engineValue(c.field, r.row.v[c.field.id] ?? null);
    }
    return { ok: true, row };
  }

  planRemove(ids) {
    const base = this.result.base;
    const rowIds = ids.map(id => this.rows.find(r => r.id === id)?.refs?.[0]?.id).filter(x => x !== undefined);
    return planDelete(this.db, base, rowIds);
  }

  remove(plan) {
    applyDelete(this.db, plan.plan);
    this.app.recordsChanged([...plan.plan.keys()]);
    this.load();
  }
}

function alignOf(type) {
  return type === 'number' || type === 'currency' || type === 'autonumber' || type === 'datetime' ? 'right' : type === 'yesno' ? 'center' : 'left';
}

/** The recordset of a record source by id (a table or a query); params for a query with prompts. */
export function recordsetOf(app, sourceId, params, options = {}) {
  if (tableById(app.db, sourceId)) return new TableRecordset(app, sourceId, options);
  if (queryById(app.db, sourceId)) return new QueryRecordset(app, sourceId, params, options);
  return null;
}

/** No records (an unbound form or report): the controls still show labels and expressions. */
export class EmptyRecordset extends Recordset {
  constructor(app) {
    super(app);
    this.kind = 'none';
    this.name = '';
  }

  get canEdit() { return false; }
  get canAdd() { return false; }
  get canDelete() { return false; }
  load() {}
  defaults() { return {}; }
}

/**
 * The records of a subform: its source (a table or a query) filtered by the link field (Link Child Fields = the value of Link Master Fields of the
 * record of the main form). Without a value of the master (a new record that is not saved yet) nothing can be added.
 */
export function subformRecordset(app, sub, masterValue) {
  const db = app.db;
  const child = tableById(db, sub?.source);
  let rs = null;
  if (child) {
    const f = sub.linkChild ? child.fields.find(x => lower(x.name) === lower(sub.linkChild)) : null;
    rs = new TableRecordset(app, child.id, f ? { fixed: { [f.id]: masterValue ?? null } } : {});
  } else if (queryById(db, sub?.source)) {
    rs = new QueryRecordset(app, sub.source, {}, sub.linkChild ? { fixed: { [sub.linkChild]: masterValue ?? null } } : {});
  }
  if (rs && sub.linkChild && (masterValue === null || masterValue === undefined)) {
    for (const prop of ['canAdd', 'canEdit', 'canDelete']) Object.defineProperty(rs, prop, { get: () => false });
  }
  return rs;
}

export { fieldById };
