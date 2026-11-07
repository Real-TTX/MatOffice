/* MatOffice Base – the database model on the content JSON (see the comment of BaseHandler.cs / docs/EDITORS.md "Base content"):
   field types and their values, what a person types into a field, the checks of a record (required, validation rule, primary key, unique index,
   referential integrity), saving and deleting records with cascades, the design changes of tables (type changes that keep the data where they can),
   relationships, names and the Name AutoCorrect of Access. A pure module without the DOM: the editor and the node tests use it.
   Every function that changes the database takes the content object and changes it in place; problems come back as { code, args } (the editor words them). */
import { BaseError, Parser, compile, compare, isNull, toText } from './expr.js';
import { ruleHolds } from './criteria.js';
import { DateVal, serialOfIso, isoOf, parseDate, parseNumber, parseBool, formatValue, formatDate, formatBool, generalNumber, roundHalfEven } from './values.js';
import { renameInSql } from './sql.js';

export const TYPES = ['text', 'memo', 'number', 'currency', 'datetime', 'yesno', 'autonumber', 'hyperlink'];
export const NUMBER_SIZES = ['byte', 'integer', 'long', 'single', 'double', 'decimal'];
export const LIMITS = { records: 200000, fields: 255, objects: 500, name: 64, text: 255, memo: 65536, controls: 1000, sql: 65536 };
export const INDEXED = ['no', 'dup', 'unique'];

const KIND_PREFIX = { tables: 't', queries: 'q', forms: 'm', reports: 'r', relationships: 'rel' };
export const OBJECT_KINDS = ['tables', 'queries', 'forms', 'reports'];
export const KIND_OF = { tables: 'table', queries: 'query', forms: 'form', reports: 'report' };

/** An empty database (the editor fills the missing parts of an older content with it too). */
export function emptyDb(currency = 'EUR') {
  return { version: 1, settings: { currency }, tables: [], relationships: [], queries: [], forms: [], reports: [], ui: { nav: {}, relationships: {} } };
}

/** Fills what an older or a foreign content lacks, so the editor can rely on the shape. */
export function ensureShape(db) {
  const d = db && typeof db === 'object' ? db : {};
  d.version ??= 1;
  d.settings ??= {};
  d.settings.currency ??= 'EUR';
  for (const k of ['tables', 'relationships', 'queries', 'forms', 'reports']) if (!Array.isArray(d[k])) d[k] = [];
  d.ui ??= {};
  d.ui.nav ??= {};
  d.ui.relationships ??= {};
  for (const t of d.tables) {
    t.fields ??= [];
    t.rows ??= [];
    t.primaryKey ??= [];
    t.nextId ??= t.rows.reduce((m, r) => Math.max(m, r.id), 0) + 1;
    t.datasheet ??= {};
    for (const f of t.fields) {
      if (f.type === 'autonumber') f.next ??= t.rows.reduce((m, r) => Math.max(m, Number(r.v[f.id]) || 0), 0) + 1;
    }
  }
  for (const q of d.queries) { q.layout ??= { tables: [] }; q.datasheet ??= {}; }
  return d;
}

// ---- lookups -------------------------------------------------------------------------------------------------------------------

export const lower = s => String(s ?? '').toLowerCase();
export const tableById = (db, id) => db.tables.find(t => t.id === id) || null;
export const queryById = (db, id) => db.queries.find(q => q.id === id) || null;
export const tableByName = (db, name) => db.tables.find(t => lower(t.name) === lower(name)) || null;
export const queryByName = (db, name) => db.queries.find(q => lower(q.name) === lower(name)) || null;
export const fieldById = (table, id) => table?.fields.find(f => f.id === id) || null;
export const fieldByName = (table, name) => table?.fields.find(f => lower(f.name) === lower(name)) || null;
export const findObject = (db, kind, id) => db[kind]?.find(o => o.id === id) || null;

/** A table or a query by id (the record source of a form or report). */
export function sourceById(db, id) {
  const t = tableById(db, id);
  if (t) return { kind: 'table', obj: t };
  const q = queryById(db, id);
  return q ? { kind: 'query', obj: q } : null;
}

/** A table or a query by its name (they share one set of names, like in Access). */
export function sourceByName(db, name) {
  const t = tableByName(db, name);
  if (t) return { kind: 'table', obj: t };
  const q = queryByName(db, name);
  return q ? { kind: 'query', obj: q } : null;
}

export const totalRecords = db => db.tables.reduce((n, t) => n + t.rows.length, 0);

/** A new id of an object of a kind ("t3", "q2" …), unique in the database. */
export function newObjectId(db, kind) {
  const prefix = KIND_PREFIX[kind];
  let max = 0;
  for (const o of db[kind]) {
    const m = new RegExp('^' + prefix + '(\\d+)$').exec(o.id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return prefix + (max + 1);
}

/** A new id inside a list (fields "f3", controls "c12"). */
export function newId(list, prefix) {
  let max = 0;
  for (const o of list) {
    const m = new RegExp('^' + prefix + '(\\d+)$').exec(o.id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return prefix + (max + 1);
}

/** What is wrong with a name of an object or a field (null: nothing). Like Access: up to 64 characters, no . ! ` [ ], no leading space. */
export function nameProblem(name) {
  const s = String(name ?? '');
  if (!s.trim()) return { code: 'NameEmpty' };
  if (s.length > LIMITS.name) return { code: 'NameLong' };
  if (/^\s/.test(s) || /[.!`[\]<>\u0000-\u001f"]/.test(s)) return { code: 'NameChars' };
  return null;
}

/** Whether a name is taken among the given kinds (tables and queries share their names). */
export function nameTaken(db, kinds, name, exceptId = null) {
  return kinds.some(k => db[k].some(o => o.id !== exceptId && lower(o.name) === lower(name)));
}

export const namespaceOf = kind => (kind === 'tables' || kind === 'queries' ? ['tables', 'queries'] : [kind]);

/** The next free name: Table1, Table2 … */
export function uniqueName(db, kinds, base) {
  for (let i = 1; ; i++) {
    const name = base + i;
    if (!nameTaken(db, kinds, name)) return name;
  }
}

/** A free name derived from a wish: "Customers", "Customers1" … */
export function freeName(db, kinds, wish) {
  const clean = String(wish || '').replace(/[.!`[\]<>"\u0000-\u001f]/g, '').trim().slice(0, LIMITS.name - 3) || 'Table';
  if (!nameTaken(db, kinds, clean)) return clean;
  for (let i = 1; ; i++) if (!nameTaken(db, kinds, clean + i)) return clean + i;
}

// ---- fields --------------------------------------------------------------------------------------------------------------------

/** A new field of a type with the defaults of Access. */
export function newField(type, name, id) {
  const f = {
    id, name, type, description: '', format: '', decimals: null, default: '', required: false, indexed: 'no', caption: '',
    validationRule: '', validationText: '', inputMask: '', lookup: null,
  };
  switch (type) {
    case 'text': f.size = 255; break;
    case 'number': f.numberSize = 'long'; break;
    case 'currency': f.format = 'Currency'; break;
    case 'yesno': f.format = 'Yes/No'; f.default = 'No'; break;
    case 'autonumber': f.next = 1; break;
    default: break;
  }
  return f;
}

export const isTextType = type => type === 'text' || type === 'memo' || type === 'hyperlink';
export const isNumericType = type => type === 'number' || type === 'currency' || type === 'autonumber';
export const isIntegerField = f => f.type === 'autonumber' || (f.type === 'number' && ['byte', 'integer', 'long'].includes(f.numberSize || 'long'));

/** The key of a value for primary keys, indexes and relationships (text without case, like Access). */
export function keyOf(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'string') return 's' + v.toLowerCase();
  if (typeof v === 'number') return 'n' + v;
  if (typeof v === 'boolean') return 'n' + (v ? -1 : 0);
  if (v instanceof DateVal) return 's' + isoOf(v.s);
  return 's' + String(v).toLowerCase();
}

/** The key of a stored value of a field (numbers of number fields and autonumbers compare as numbers). */
export const storedKey = v => keyOf(v);

/** A stored value as a value of expressions (dates become DateVal). */
export function engineValue(field, raw) {
  if (raw === null || raw === undefined) return null;
  if (field?.type === 'datetime' && typeof raw === 'string') {
    const s = serialOfIso(raw);
    return s === null ? raw : new DateVal(s);
  }
  return raw;
}

const RANGES = { byte: [0, 255], integer: [-32768, 32767], long: [-2147483648, 2147483647] };

/**
 * Converts a value (of an expression, an import, a type change or typed text) into the stored value of a field.
 * { ok: true, value } or { ok: false, code: 'TypeMismatch'|'TooLong'|'OutOfRange', value: what could be kept }.
 */
export function coerce(field, v, lang = 'en') {
  if (v === undefined) v = null;
  const type = field.type;
  if (type === 'yesno') {
    if (v === null || v === '') return { ok: true, value: false };
    const b = v instanceof DateVal ? v.s !== 0 : parseBool(v);
    return b === null ? { ok: false, code: 'TypeMismatch', value: false } : { ok: true, value: b };
  }
  if (v === null) return { ok: true, value: null };
  if (isTextType(type)) {
    let s;
    if (v instanceof DateVal) s = formatDate(v, lang, '');
    else if (typeof v === 'number') s = generalNumber(v, lang);
    else if (typeof v === 'boolean') s = formatBool(v, lang);
    else s = String(v);
    if (s === '') return { ok: true, value: null };
    const max = type === 'text' ? Number(field.size) || 255 : LIMITS.memo;
    if (s.length > max) return { ok: false, code: 'TooLong', value: s.slice(0, max) };
    return { ok: true, value: s };
  }
  if (type === 'datetime') {
    if (v instanceof DateVal) return { ok: true, value: isoOf(v.s) };
    if (typeof v === 'number') return { ok: true, value: isoOf(v) };
    if (typeof v === 'boolean') return { ok: false, code: 'TypeMismatch', value: null };
    const s = String(v).trim();
    if (!s) return { ok: true, value: null };
    const serial = serialOfIso(s) ?? parseDate(s, lang);
    return serial === null ? { ok: false, code: 'TypeMismatch', value: null } : { ok: true, value: isoOf(serial) };
  }
  // numbers
  let n;
  if (typeof v === 'number') n = v;
  else if (typeof v === 'boolean') n = v ? -1 : 0;
  else if (v instanceof DateVal) n = v.s;
  else {
    const s = String(v).trim();
    if (!s) return { ok: true, value: null };
    n = parseNumber(s, lang);
    if (n === null && lang !== 'en') n = parseNumber(s, 'en');
    if (n === null) return { ok: false, code: 'TypeMismatch', value: null };
  }
  if (!Number.isFinite(n)) return { ok: false, code: 'TypeMismatch', value: null };
  if (type === 'currency') return { ok: true, value: roundHalfEven(n, 4) };
  if (type === 'autonumber') {
    if (!Number.isInteger(n) || n < 1 || n > 2147483647) return { ok: false, code: 'TypeMismatch', value: null };
    return { ok: true, value: n };
  }
  const size = field.numberSize || 'long';
  if (RANGES[size]) {
    const r = roundHalfEven(n);
    const [lo, hi] = RANGES[size];
    if (r < lo || r > hi) return { ok: false, code: 'OutOfRange', value: null };
    return { ok: true, value: r };
  }
  if (size === 'single') return { ok: true, value: Number(n.toPrecision(7)) };
  if (size === 'decimal' && field.decimals !== null && field.decimals !== undefined && field.decimals !== '') return { ok: true, value: roundHalfEven(n, Number(field.decimals)) };
  return { ok: true, value: n };
}

/** What a person typed into a field (empty text = Null). */
export function parseInput(field, text, lang = 'en') {
  if (typeof text === 'string' && text === '' && field.type !== 'yesno') return { ok: true, value: null };
  if (field.type === 'hyperlink' && typeof text === 'string') return coerce(field, text.trim(), lang);
  return coerce(field, text, lang);
}

/** A stored value for the screen, with the format and the decimal places of the field. */
export function displayValue(db, field, raw, lang = 'en', lookupMap = null) {
  if (raw === null || raw === undefined) return '';
  if (lookupMap && field.lookup?.kind === 'table') {
    const shown = lookupMap.get(keyOf(raw));
    if (shown !== undefined) return shown;
  }
  return formatValue(engineValue(field, raw), { type: field.type, format: field.format, decimals: field.decimals, lang, currency: db.settings?.currency || 'EUR' });
}

/** The list of a lookup field: [{ value, label, columns }] (a table's bound field and display field, or a value list). */
export function lookupOptions(db, field, lang = 'en') {
  const lk = field.lookup;
  if (!lk) return [];
  if (lk.kind === 'values') return (lk.values || []).map(v => ({ value: isNumericType(field.type) ? (parseNumber(v, 'en') ?? v) : v, label: String(v) }));
  const t = tableById(db, lk.table);
  if (!t) return [];
  const bound = fieldById(t, lk.bound);
  const shown = fieldById(t, lk.display) || bound;
  if (!bound) return [];
  const list = t.rows.map(r => ({ value: r.v[bound.id] ?? null, label: displayValue(db, shown, r.v[shown.id], lang) }));
  const coll = new Intl.Collator(lang === 'de' ? 'de' : 'en', { sensitivity: 'accent', numeric: true });
  return list.filter(o => o.value !== null).sort((a, b) => coll.compare(a.label, b.label));
}

/** A map key → shown text for a lookup to a table (the datasheet shows the display field instead of the stored key). */
export function lookupMap(db, field, lang = 'en') {
  if (field.lookup?.kind !== 'table') return null;
  const map = new Map();
  for (const o of lookupOptions(db, field, lang)) map.set(keyOf(o.value), o.label);
  return map;
}

// ---- defaults and checks -------------------------------------------------------------------------------------------------------

/** The default values of a new record (Default Value of each field: 0, "Berlin", Date(), No …). AutoNumbers stay empty ("(New)"). */
export function defaultValues(db, table, lang = 'en', clock = null) {
  const values = {};
  for (const f of table.fields) {
    if (f.type === 'autonumber') continue;
    let text = String(f.default ?? '').trim();
    if (!text) { if (f.type === 'yesno') values[f.id] = false; continue; }
    if (text.startsWith('=')) text = text.slice(1);
    try {
      const v = compile(Parser.expression(text), { lang, now: clock ? () => clock : undefined, currency: db.settings?.currency })(null);
      const c = coerce(f, v, lang);
      if (c.ok) values[f.id] = c.value;
    } catch {
      // a default that cannot be read is left out (Access would refuse it in the design already)
    }
  }
  return values;
}

/** Whether a default value or a validation rule can be read (the design view checks it when the cell is left). */
export function expressionProblem(text, kind = 'expr') {
  const s = String(text ?? '').trim();
  if (!s) return null;
  try {
    if (kind === 'rule') ruleHolds(s, null, compile, {});
    else Parser.expression(s.startsWith('=') ? s.slice(1) : s);
    return null;
  } catch (e) {
    if (e instanceof BaseError) return { code: e.code === 'UnknownName' ? 'RuleName' : 'Syntax', args: e.args };
    return { code: 'Syntax', args: [] };
  }
}

/** Whether a field has a unique index (the primary key of one field, or Indexed: Yes (No Duplicates)). */
export function isUniqueField(table, fieldId) {
  const f = fieldById(table, fieldId);
  if (!f) return false;
  if (f.indexed === 'unique') return true;
  return table.primaryKey.length === 1 && table.primaryKey[0] === fieldId;
}

export const childRelationships = (db, tableId) => db.relationships.filter(r => r.parent.table === tableId);
export const parentRelationships = (db, tableId) => db.relationships.filter(r => r.child.table === tableId);

function compositeKey(values, ids) {
  const parts = [];
  for (const id of ids) {
    const k = keyOf(values[id]);
    if (k === null) return null;
    parts.push(k);
  }
  return parts.join('\u0001');
}

/**
 * The first problem of a record (null: none): Required, Validation (rule of a field), PkNull, Duplicate (primary key or unique index), NoParent (integrity).
 * values: the stored values by field id; rowId: the record that is changed (null for a new one).
 */
export function checkRecord(db, table, values, rowId = null, lang = 'en', options = {}) {
  for (const f of table.fields) {
    const v = values[f.id];
    const empty = v === null || v === undefined || v === '';
    if (empty && table.primaryKey.includes(f.id) && f.type !== 'autonumber') return { code: 'PkNull', field: f.id, args: [f.name] };
    if (empty && f.required && f.type !== 'autonumber') return { code: 'Required', field: f.id, args: [table.name + '.' + f.name] };
    if (f.validationRule && !options.skipRules) {
      let holds = true;
      try { holds = ruleHolds(f.validationRule, engineValue(f, v), compile, { lang, currency: db.settings?.currency }); } catch { holds = true; }
      if (!holds) return { code: 'Validation', field: f.id, args: [f.validationText || '', f.validationRule, table.name + '.' + f.name] };
    }
  }
  const unique = [];
  if (table.primaryKey.length) unique.push({ ids: table.primaryKey, pk: true });
  for (const f of table.fields) if (f.indexed === 'unique' && !(table.primaryKey.length === 1 && table.primaryKey[0] === f.id)) unique.push({ ids: [f.id], pk: false });
  for (const u of unique) {
    const key = compositeKey(values, u.ids);
    if (key === null) continue;
    const hit = options.index?.get(u.ids.join(','))?.get(key);
    if (hit !== undefined ? hit !== rowId : table.rows.some(r => r.id !== rowId && compositeKey(r.v, u.ids) === key)) {
      return { code: 'Duplicate', field: u.ids[0], args: [table.name] };
    }
  }
  for (const rel of parentRelationships(db, table.id)) {
    if (!rel.integrity) continue;
    const v = values[rel.child.field];
    const k = keyOf(v);
    if (k === null) continue;
    const parent = tableById(db, rel.parent.table);
    if (!parent) continue;
    const found = options.parentKeys?.get(rel.id)?.has(k) ?? parent.rows.some(r => keyOf(r.v[rel.parent.field]) === k);
    if (!found) return { code: 'NoParent', field: rel.child.field, args: [parent.name] };
  }
  return null;
}

/** The rows of a child table that point at a parent value (through one relationship). */
function childrenOf(db, rel, parentValue) {
  const child = tableById(db, rel.child.table);
  const k = keyOf(parentValue);
  if (!child || k === null) return [];
  return child.rows.filter(r => keyOf(r.v[rel.child.field]) === k);
}

/**
 * Saves a record (new when rowId is null): checks it, gives AutoNumbers, cascades a changed key to related records (Cascade Update) or refuses
 * (the key has related records). Answers { ok, row } or { ok: false, error: { code, field, args } }.
 */
export function saveRecord(db, table, rowId, input, lang = 'en') {
  const existing = rowId === null || rowId === undefined ? null : table.rows.find(r => r.id === rowId);
  if (rowId !== null && rowId !== undefined && !existing) return { ok: false, error: { code: 'RecordGone', args: [] } };
  if (!existing && totalRecords(db) >= LIMITS.records) return { ok: false, error: { code: 'TooManyRecords', args: [LIMITS.records] } };
  const values = {};
  for (const f of table.fields) {
    let v = Object.prototype.hasOwnProperty.call(input, f.id) ? input[f.id] : existing ? existing.v[f.id] ?? null : null;
    if (v === '' ) v = null;
    if (f.type === 'yesno' && (v === null || v === undefined)) v = false;
    values[f.id] = v ?? null;
  }
  const autos = table.fields.filter(f => f.type === 'autonumber');
  if (!existing) for (const f of autos) if (values[f.id] === null) values[f.id] = f.next ?? 1;
  const problem = checkRecord(db, table, values, existing ? existing.id : null, lang);
  if (problem) return { ok: false, error: problem };

  if (existing) {
    // a changed value of a parent field: cascade or refuse
    const cascades = [];
    for (const rel of childRelationships(db, table.id)) {
      const before = existing.v[rel.parent.field];
      const after = values[rel.parent.field];
      if (keyOf(before) === keyOf(after) || keyOf(before) === null) continue;
      const kids = childrenOf(db, rel, before);
      if (!kids.length || !rel.integrity) continue;
      if (!rel.cascadeUpdate) return { ok: false, error: { code: 'HasChildren', field: rel.parent.field, args: [tableById(db, rel.child.table)?.name || ''] } };
      cascades.push([rel, kids, after]);
    }
    for (const [rel, kids, after] of cascades) for (const k of kids) k.v[rel.child.field] = after;
    existing.v = compact(values);
    return { ok: true, row: existing };
  }
  for (const f of autos) f.next = Math.max(f.next ?? 1, Number(values[f.id]) + 1);
  const row = { id: table.nextId ?? 1, v: compact(values) };
  table.nextId = row.id + 1;
  table.rows.push(row);
  return { ok: true, row };
}

/** The stored values without the empty ones (smaller content). Yes/No keeps false. */
function compact(values) {
  const out = {};
  for (const [k, v] of Object.entries(values)) if (v !== null && v !== undefined) out[k] = v;
  return out;
}

/**
 * What deleting rows would do: { ok, plan: Map(tableId → Set(rowIds)), cascaded: n } or { ok: false, error: { code: 'HasChildren', args: [table] } }.
 * Relationships with integrity and Cascade Delete delete the related records too (through every level); without cascade they forbid the delete.
 */
export function planDelete(db, table, rowIds) {
  const plan = new Map([[table.id, new Set(rowIds)]]);
  const queue = [[table, [...rowIds]]];
  while (queue.length) {
    const [t, ids] = queue.shift();
    const rows = t.rows.filter(r => ids.includes(r.id));
    for (const rel of childRelationships(db, t.id)) {
      if (!rel.integrity) continue;
      const child = tableById(db, rel.child.table);
      if (!child) continue;
      const keys = new Set(rows.map(r => keyOf(r.v[rel.parent.field])).filter(k => k !== null));
      if (!keys.size) continue;
      const set = plan.get(child.id) || new Set();
      const kids = child.rows.filter(r => keys.has(keyOf(r.v[rel.child.field])) && !set.has(r.id));
      if (!kids.length) continue;
      if (!rel.cascadeDelete) return { ok: false, error: { code: 'HasChildren', args: [child.name] } };
      for (const k of kids) set.add(k.id);
      plan.set(child.id, set);
      queue.push([child, kids.map(k => k.id)]);
    }
  }
  let cascaded = 0;
  for (const [id, set] of plan) if (id !== table.id) cascaded += set.size;
  return { ok: true, plan, cascaded };
}

/** Deletes rows (after planDelete said yes). */
export function applyDelete(db, plan) {
  for (const [tableId, ids] of plan) {
    const t = tableById(db, tableId);
    if (t) t.rows = t.rows.filter(r => !ids.has(r.id));
  }
}

export function deleteRecords(db, table, rowIds) {
  const p = planDelete(db, table, rowIds);
  if (p.ok) applyDelete(db, p.plan);
  return p;
}

/** Indexes for checks of many records at once (append and update queries): unique keys → row id, parent keys of relationships. */
export function bulkIndexes(db, table) {
  const index = new Map();
  const uniques = [];
  if (table.primaryKey.length) uniques.push(table.primaryKey);
  for (const f of table.fields) if (f.indexed === 'unique') uniques.push([f.id]);
  for (const ids of uniques) {
    const m = new Map();
    for (const r of table.rows) { const k = compositeKey(r.v, ids); if (k !== null) m.set(k, r.id); }
    index.set(ids.join(','), m);
  }
  const parentKeys = new Map();
  for (const rel of parentRelationships(db, table.id)) {
    if (!rel.integrity) continue;
    const parent = tableById(db, rel.parent.table);
    parentKeys.set(rel.id, new Set((parent?.rows || []).map(r => keyOf(r.v[rel.parent.field])).filter(k => k !== null)));
  }
  return { index, parentKeys, add(row) { for (const ids of uniques) { const k = compositeKey(row.v, ids); if (k !== null) index.get(ids.join(',')).set(k, row.id); } } };
}

// ---- relationships -------------------------------------------------------------------------------------------------------------

const typeFamily = f => (f.type === 'autonumber' || (f.type === 'number' && (f.numberSize || 'long') === 'long') ? 'long' : f.type === 'number' ? 'number-' + (f.numberSize || 'long') : isTextType(f.type) ? 'text' : f.type);

/** One-to-many, one-to-one or indeterminate (no unique side). */
export function relationshipKind(db, rel) {
  const p = tableById(db, rel.parent.table);
  const c = tableById(db, rel.child.table);
  if (!p || !c) return 'indeterminate';
  const pu = isUniqueField(p, rel.parent.field);
  const cu = isUniqueField(c, rel.child.field);
  if (pu && cu) return 'oneToOne';
  if (pu) return 'oneToMany';
  return 'indeterminate';
}

/** Which side of a dragged pair is the parent: the unique one (dragging from a foreign key onto a key works too). */
export function orientRelationship(db, from, to) {
  const ft = tableById(db, from.table);
  const tt = tableById(db, to.table);
  if (ft && tt && !isUniqueField(ft, from.field) && isUniqueField(tt, to.field)) return { parent: to, child: from };
  return { parent: from, child: to };
}

/** The problem of a relationship (null: it can be made): RelSame, RelTypes, RelExists, RelNoUnique, RelViolated. */
export function relationshipProblem(db, rel) {
  const p = tableById(db, rel.parent.table);
  const c = tableById(db, rel.child.table);
  const pf = fieldById(p, rel.parent.field);
  const cf = fieldById(c, rel.child.field);
  if (!pf || !cf) return { code: 'RelMissing', args: [] };
  if (p.id === c.id && pf.id === cf.id) return { code: 'RelSame', args: [] };
  if (typeFamily(pf) !== typeFamily(cf)) return { code: 'RelTypes', args: [pf.name, cf.name] };
  const twin = db.relationships.find(r => r.id !== rel.id
    && ((r.parent.table === rel.parent.table && r.parent.field === rel.parent.field && r.child.table === rel.child.table && r.child.field === rel.child.field)
      || (r.parent.table === rel.child.table && r.parent.field === rel.child.field && r.child.table === rel.parent.table && r.child.field === rel.parent.field)));
  if (twin) return { code: 'RelExists', args: [p.name, c.name] };
  if (rel.integrity) {
    if (!isUniqueField(p, pf.id)) return { code: 'RelNoUnique', args: [p.name] };
    const keys = new Set(p.rows.map(r => keyOf(r.v[pf.id])).filter(k => k !== null));
    if (c.rows.some(r => { const k = keyOf(r.v[cf.id]); return k !== null && !keys.has(k); })) return { code: 'RelViolated', args: [c.name, p.name] };
  }
  return null;
}

export function saveRelationship(db, rel) {
  const problem = relationshipProblem(db, rel);
  if (problem) return { ok: false, error: problem };
  if (!rel.integrity) { rel.cascadeUpdate = false; rel.cascadeDelete = false; }
  const i = db.relationships.findIndex(r => r.id === rel.id);
  if (i >= 0) db.relationships[i] = rel;
  else db.relationships.push({ ...rel, id: rel.id || newObjectId(db, 'relationships') });
  return { ok: true };
}

/** Whether a field takes part in a relationship. */
export const fieldInRelationship = (db, tableId, fieldId) => db.relationships.some(r => (r.parent.table === tableId && r.parent.field === fieldId) || (r.child.table === tableId && r.child.field === fieldId));

// ---- design changes of tables --------------------------------------------------------------------------------------------------

/**
 * What a change of a field's properties would do to the data: { blocked: { code, args } } when it cannot be done, else
 * { lost: n (values that cannot be converted and become empty), truncated: n }.
 */
export function previewFieldChange(db, table, fieldId, patch, lang = 'en') {
  const f = fieldById(table, fieldId);
  const next = { ...f, ...patch };
  const typeChanged = next.type !== f.type || (next.type === 'number' && next.numberSize !== f.numberSize) || (next.type === 'text' && Number(next.size) < Number(f.size || 255));
  if (typeChanged && fieldInRelationship(db, table.id, f.id) && typeFamily(next) !== typeFamily(f)) return { blocked: { code: 'FieldInRelationship', args: [f.name] } };
  if (next.type === 'autonumber' && f.type !== 'autonumber' && table.rows.length) return { blocked: { code: 'NoAutoNumberWithData', args: [] } };
  if (next.type === 'autonumber' && table.fields.some(x => x.id !== f.id && x.type === 'autonumber')) return { blocked: { code: 'OneAutoNumber', args: [] } };
  let lost = 0;
  let truncated = 0;
  if (typeChanged) {
    for (const r of table.rows) {
      const v = r.v[f.id];
      if (v === null || v === undefined) continue;
      const c = coerce(next, engineValue(f, v), lang);
      if (!c.ok) { if (c.code === 'TooLong') truncated++; else lost++; }
    }
  }
  const empty = r => { const v = r.v[f.id]; return v === null || v === undefined || v === ''; };
  if (next.required && !f.required && f.type !== 'autonumber' && table.rows.some(empty)) return { blocked: { code: 'RequiredViolated', args: [f.name] } };
  if (next.indexed === 'unique' && f.indexed !== 'unique') {
    const seen = new Set();
    for (const r of table.rows) {
      const k = keyOf(r.v[f.id]);
      if (k === null) continue;
      if (seen.has(k)) return { blocked: { code: 'DuplicatesExist', args: [f.name] } };
      seen.add(k);
    }
  }
  if (next.validationRule && next.validationRule !== f.validationRule) {
    const problem = expressionProblem(next.validationRule, 'rule');
    if (problem) return { blocked: problem };
  }
  if (next.default !== f.default) {
    const problem = expressionProblem(next.default);
    if (problem) return { blocked: problem };
  }
  return { lost, truncated };
}

/** Applies a change of a field's properties (after previewFieldChange), converting the stored values when the type changed. */
export function applyFieldChange(db, table, fieldId, patch, lang = 'en') {
  const f = fieldById(table, fieldId);
  const before = { ...f };
  Object.assign(f, patch);
  if (f.type !== 'text') delete f.size; else f.size = Math.max(1, Math.min(255, Number(f.size) || 255));
  if (f.type !== 'number') delete f.numberSize; else f.numberSize ||= 'long';
  if (f.type === 'autonumber') { f.next ??= 1; f.lookup = null; } else delete f.next;
  if (before.type !== f.type) {
    // the format of the old type means nothing to the new one
    if (patch.format === undefined) f.format = f.type === 'currency' ? 'Currency' : f.type === 'yesno' ? 'Yes/No' : '';
    if (f.type === 'yesno' && patch.default === undefined) f.default = 'No';
    else if (before.type === 'yesno' && patch.default === undefined) f.default = '';
  }
  const typeChanged = before.type !== f.type || before.numberSize !== f.numberSize || before.size !== f.size;
  if (typeChanged) {
    for (const r of table.rows) {
      const v = r.v[f.id];
      if ((v === null || v === undefined) && f.type !== 'yesno') continue;
      const c = coerce(f, engineValue(before, v ?? null), lang);
      const value = c.ok ? c.value : c.code === 'TooLong' ? c.value : f.type === 'yesno' ? false : null;
      if (value === null || value === undefined) delete r.v[f.id]; else r.v[f.id] = value;
    }
  }
  if (f.type === 'autonumber' && !table.primaryKey.length) table.primaryKey = [f.id];
}

/** Adds a field (at an index); answers it. */
export function addField(table, type, name, index = table.fields.length) {
  const f = newField(type, name, newId(table.fields, 'f'));
  table.fields.splice(index, 0, f);
  return f;
}

/** Whether a field can be deleted ({ code } when not). */
export function deleteFieldProblem(db, table, fieldId) {
  if (fieldInRelationship(db, table.id, fieldId)) return { code: 'FieldInRelationship', args: [fieldById(table, fieldId)?.name || ''] };
  return null;
}

export function deleteField(db, table, fieldId) {
  table.fields = table.fields.filter(f => f.id !== fieldId);
  table.primaryKey = table.primaryKey.filter(id => id !== fieldId);
  for (const r of table.rows) delete r.v[fieldId];
  const ds = table.datasheet || {};
  if (ds.widths) delete ds.widths[fieldId];
  if (ds.totals) delete ds.totals[fieldId];
  if (Array.isArray(ds.hidden)) ds.hidden = ds.hidden.filter(id => id !== fieldId);
  if (Array.isArray(ds.sort)) ds.sort = ds.sort.filter(s => s.field !== fieldId);
  for (const t of db.tables) for (const x of t.fields) if (x.lookup?.kind === 'table' && x.lookup.table === table.id && (x.lookup.bound === fieldId || x.lookup.display === fieldId)) x.lookup = null;
}

/** Whether a set of fields can be the primary key (no empty values, no duplicates, not in use by a relationship that needs the old key). */
export function primaryKeyProblem(db, table, ids) {
  for (const id of ids) {
    const f = fieldById(table, id);
    if (!f) return { code: 'Syntax', args: [] };
    if (f.type === 'memo') return { code: 'PkMemo', args: [f.name] };
  }
  if (ids.length && table.rows.some(r => compositeKey(r.v, ids) === null)) return { code: 'PkNulls', args: [] };
  const seen = new Set();
  for (const r of table.rows) {
    const k = compositeKey(r.v, ids);
    if (seen.has(k)) return { code: 'PkDuplicates', args: [] };
    seen.add(k);
  }
  // a relationship whose parent field would no longer be unique
  for (const rel of childRelationships(db, table.id)) {
    if (!rel.integrity) continue;
    const f = fieldById(table, rel.parent.field);
    const stillUnique = f?.indexed === 'unique' || (ids.length === 1 && ids[0] === rel.parent.field);
    if (!stillUnique) return { code: 'PkInRelationship', args: [] };
  }
  return null;
}

// ---- names: rename with AutoCorrect, delete, duplicate ---------------------------------------------------------------------------

/** Renames names in an expression text (control sources, defaults): [Old] → [New]. */
export function renameInExpr(text, from, to) {
  if (!text) return text;
  const s = String(text);
  const eq = s.startsWith('=');
  const body = eq ? s.slice(1) : s;
  if (!eq && lower(body) === lower(from)) return to;
  let tokens;
  try { tokens = new Parser(body).tokens; } catch { return text; }
  let out = body;
  for (const t of [...tokens].reverse()) {
    if ((t.k === 'name' || t.k === 'word') && lower(t.v) === lower(from)) {
      const next = tokens[tokens.indexOf(t) + 1];
      if (t.k === 'word' && next?.k === 'op' && next.v === '(') continue;
      out = out.slice(0, t.pos) + '[' + to + ']' + out.slice(t.end);
    }
  }
  return (eq ? '=' : '') + out;
}

function renameInControls(controls, from, to) {
  for (const c of controls || []) {
    if (c.source) c.source = renameInExpr(c.source, from, to);
  }
}

function sectionsOf(obj) {
  const s = obj.sections || {};
  return [s.header, s.detail, s.footer, s.reportHeader, s.pageHeader, s.pageFooter, s.reportFooter, ...(s.groupHeaders || []), ...(s.groupFooters || [])].filter(Boolean);
}

/** Renames a field and fixes the queries, forms and reports that use it (Name AutoCorrect). */
export function renameField(db, table, fieldId, name) {
  const f = fieldById(table, fieldId);
  const old = f.name;
  f.name = name;
  if (lower(old) === lower(name) && old === name) return;
  for (const q of db.queries) q.sql = renameInSql(q.sql, 'field', old, name, table.name);
  for (const obj of [...db.forms, ...db.reports]) {
    if (obj.source !== table.id) continue;
    for (const sec of sectionsOf(obj)) renameInControls(sec.controls, old, name);
    for (const g of obj.grouping || []) if (lower(g.field) === lower(old)) g.field = name;
    for (const s of obj.sorting || []) if (lower(s.field) === lower(old)) s.field = name;
  }
  for (const obj of db.forms) {
    for (const sec of sectionsOf(obj)) for (const c of sec.controls || []) {
      if (c.type === 'subform' && c.subform?.source === table.id) {
        if (lower(c.subform.linkChild) === lower(old)) c.subform.linkChild = name;
      }
      if (c.type === 'subform' && obj.source === table.id && lower(c.subform?.linkMaster) === lower(old)) c.subform.linkMaster = name;
    }
  }
}

/** Renames a table, query, form or report; queries that use a table or query by name follow. */
export function renameObject(db, kind, id, name) {
  const obj = findObject(db, kind, id);
  const old = obj.name;
  obj.name = name;
  if (kind === 'tables' || kind === 'queries') {
    for (const q of db.queries) q.sql = renameInSql(q.sql, 'object', old, name);
  }
}

/** Deletes an object; a table takes its relationships with it. */
export function deleteObject(db, kind, id) {
  db[kind] = db[kind].filter(o => o.id !== id);
  if (kind === 'tables') {
    db.relationships = db.relationships.filter(r => r.parent.table !== id && r.child.table !== id);
    delete db.ui.relationships?.[id];
    for (const t of db.tables) for (const f of t.fields) if (f.lookup?.kind === 'table' && f.lookup.table === id) f.lookup = null;
  }
}

/** A copy of an object under a new name (a table with its data). */
export function duplicateObject(db, kind, id, name) {
  const obj = findObject(db, kind, id);
  const copy = structuredClone(obj);
  copy.id = newObjectId(db, kind);
  copy.name = name;
  db[kind].push(copy);
  return copy;
}

/** Compact & Repair: AutoNumber seeds and row ids follow the data again, references to things that are gone are removed. */
export function compactDb(db) {
  let fixed = 0;
  for (const t of db.tables) {
    const maxId = t.rows.reduce((m, r) => Math.max(m, r.id), 0);
    if (t.nextId !== maxId + 1) { t.nextId = maxId + 1; fixed++; }
    for (const f of t.fields) {
      if (f.type !== 'autonumber') continue;
      const max = t.rows.reduce((m, r) => Math.max(m, Number(r.v[f.id]) || 0), 0);
      if (f.next !== max + 1) { f.next = max + 1; fixed++; }
    }
    const ids = new Set(t.fields.map(f => f.id));
    for (const r of t.rows) for (const k of Object.keys(r.v)) if (!ids.has(k)) { delete r.v[k]; fixed++; }
    t.primaryKey = t.primaryKey.filter(id => ids.has(id));
    for (const f of t.fields) {
      if (f.lookup?.kind === 'table') {
        const lt = tableById(db, f.lookup.table);
        if (!lt || !fieldById(lt, f.lookup.bound)) { f.lookup = null; fixed++; }
      }
    }
  }
  const before = db.relationships.length;
  db.relationships = db.relationships.filter(r => fieldById(tableById(db, r.parent.table), r.parent.field) && fieldById(tableById(db, r.child.table), r.child.field));
  fixed += before - db.relationships.length;
  for (const id of Object.keys(db.ui.relationships || {})) if (!tableById(db, id)) { delete db.ui.relationships[id]; fixed++; }
  return fixed;
}

/** The value of a record field for expressions by name (forms, reports, defaults). */
export function recordScope(fields, values) {
  const byName = new Map(fields.map(f => [lower(f.name), f]));
  return parts => {
    const f = byName.get(lower(parts[parts.length - 1]));
    if (!f) return null;
    return () => engineValue(f, values()[f.key ?? f.id]);
  };
}

export { toText, compare, isNull };
