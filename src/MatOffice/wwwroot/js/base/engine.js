/* MatOffice Base – the query engine: runs the SQL of sql.js on the tables of the content. FROM with INNER/LEFT/RIGHT joins (hash joins for
   equal fields), saved queries and subqueries as sources, WHERE, GROUP BY with Sum/Avg/Count/Min/Max/First/Last/StDev/Var, HAVING, ORDER BY,
   TOP n [PERCENT] (with ties, like Access), DISTINCT, parameters ([Enter a city]) and the action queries (make table, append, update, delete),
   which go through the checks of model.js (validation, keys, referential integrity, cascades). A pure module without the DOM.

   A result: { columns: [{ name, type, field, source, key, editable }], rows: [{ v: [values], refs: [row of each source] }], updatable, base }. */
import { compile, BaseError, Parser, sortCompare, isNull, AGGREGATES, walk, toNumber, toBool, sameExpr, compare, exprToText, collator } from './expr.js';
import { parseSql, queryType, sqlExpr } from './sql.js';
import {
  tableByName, queryByName, engineValue, coerce, saveRecord, planDelete, applyDelete, bulkIndexes, newObjectId, newField, lower, LIMITS, totalRecords,
  keyOf, fieldByName, isNumericType, childRelationships, parentRelationships, newId,
} from './model.js';
import { DateVal } from './values.js';

const MAX_DEPTH = 12;

/** Reads a query text: { stmt, type } (throws BaseError with the place of a syntax error). */
export function prepare(sql) {
  const stmt = parseSql(sql);
  return { stmt, type: queryType(stmt) };
}

/** The parameters a statement asks for, in the order Access asks them (declared ones first): [{ name, type }]. */
export function parametersOf(db, stmt, options = {}) {
  const collect = new Map();
  for (const p of stmt.params || []) collect.set(lower(p.name), { name: p.name, type: p.type || '' });
  const ctx = { ...context(db, options), collect, dry: true };
  try {
    if (stmt.type === 'select') execSelect(db, stmt, ctx);
    else if (stmt.type === 'insert') { if (stmt.select) execSelect(db, stmt.select, ctx); else stmt.values.forEach(v => compile(v, scopeFor([], ctx, null))); }
    else if (stmt.type === 'update') planUpdate(db, stmt, ctx);
    else if (stmt.type === 'delete') planDeleteQuery(db, stmt, ctx);
  } catch (e) {
    if (!(e instanceof BaseError)) throw e;
    throw e;
  }
  return [...collect.values()];
}

/** Turns the answers of the prompts (texts) into values: numbers, dates and Yes/No where the text reads as one, as Access does for undeclared parameters. */
export function paramValues(list, answers, lang = 'en') {
  const out = {};
  for (const p of list) {
    const text = answers[lower(p.name)] ?? answers[p.name] ?? '';
    const type = (p.type || '').toLowerCase();
    let v = text === '' ? null : text;
    if (v !== null && /date|time/.test(type)) { try { v = new DateVal(Number(compile(Parser.expression('CDate("' + text.replace(/"/g, '""') + '")'), { lang })(null))); } catch { v = text; } }
    else if (v !== null && /long|integer|byte|single|double|currency|decimal|number|short/.test(type)) { const n = Number(String(text).replace(',', '.')); v = Number.isFinite(n) ? n : text; }
    out[lower(p.name)] = v;
  }
  return out;
}

function context(db, options) {
  return {
    db,
    params: options.params || {},
    lang: options.lang || 'en',
    now: options.now,
    currency: db.settings?.currency || 'EUR',
    depth: options.depth || 0,
    collect: options.collect || null,
    dry: !!options.dry,
    stack: options.stack || [],
  };
}

/** Runs a SELECT statement (or the text of one) and answers the result. */
export function runSelect(db, stmtOrSql, options = {}) {
  const stmt = typeof stmtOrSql === 'string' ? parseSql(stmtOrSql) : stmtOrSql;
  if (stmt.type !== 'select') throw new BaseError('NotSelect');
  return execSelect(db, stmt, context(db, options));
}

/** Runs a saved query by its object (for forms, reports, the datasheet of a query): its SELECT part also for a make-table query. */
export function runQueryObject(db, query, options = {}) {
  const stmt = parseSql(query.sql);
  const ctx = context(db, options);
  ctx.stack = [...ctx.stack, lower(query.name)];
  if (stmt.type === 'select') return execSelect(db, { ...stmt, into: null }, ctx);
  if (stmt.type === 'insert' && stmt.select) return execSelect(db, stmt.select, ctx);
  // update/delete: the rows that would change, like the datasheet view of Access shows them
  const from = stmt.from;
  const target = stmt.type === 'delete' ? stmt.target : null;
  const columns = stmt.type === 'update' ? stmt.set.map(s => ({ e: s.target, alias: null })) : [{ e: { t: 'star', table: target }, alias: null }];
  return execSelect(db, { type: 'select', distinct: null, top: null, columns, into: null, from, where: stmt.where, groupBy: [], having: null, orderBy: [], params: stmt.params }, ctx);
}

// ---- sources -------------------------------------------------------------------------------------------------------------------

function openTable(db, name, alias, ctx) {
  const t = tableByName(db, name);
  if (t) {
    return {
      alias: alias || t.name, table: t, rows: ctx.dry ? [] : t.rows,
      fields: t.fields.map(f => ({ name: f.name, key: f.id, type: f.type, field: f })),
    };
  }
  const q = queryByName(db, name);
  if (q) {
    if (ctx.stack.includes(lower(q.name)) || ctx.depth > MAX_DEPTH) throw new BaseError('Circular', q.name);
    const stmt = parseSql(q.sql);
    const inner = { ...ctx, depth: ctx.depth + 1, stack: [...ctx.stack, lower(q.name)] };
    let res;
    if (stmt.type === 'select') res = execSelect(db, { ...stmt, into: null }, inner);
    else throw new BaseError('NotSelect', q.name);
    return subSource(res, alias || q.name);
  }
  throw new BaseError('NoTable', name);
}

function subSource(res, alias) {
  return { alias, table: null, rows: res.rows, fields: res.columns.map((c, i) => ({ name: c.name, key: i, type: c.type, field: c.field })), result: res };
}

const valueOf = (row, f) => (row ? (f.field && typeof f.key === 'string' ? engineValue(f.field, row.v[f.key]) : row.v[f.key]) : null);

/** Flattens the FROM list into sources and builds the rows (tuples: one row per source, null for the missing side of an outer join). */
function buildFrom(db, from, ctx) {
  const sources = [];
  const collectSources = item => {
    if (item.t === 'table') { item._i = sources.length; sources.push(openTable(db, item.name, item.alias, ctx)); }
    else if (item.t === 'sub') {
      item._i = sources.length;
      const res = execSelect(db, item.query, { ...ctx, depth: ctx.depth + 1 });
      sources.push(subSource(res, item.alias));
    } else { collectSources(item.left); collectSources(item.right); }
  };
  from.forEach(collectSources);
  const n = sources.length;
  const scope = scopeFor(sources, ctx, null);

  const evalItem = item => {
    if (item.t === 'table' || item.t === 'sub') {
      const i = item._i;
      return { idx: [i], tuples: sources[i].rows.map(r => { const t = new Array(n).fill(null); t[i] = r; return t; }) };
    }
    const L = evalItem(item.left);
    const R = evalItem(item.right);
    const on = compile(item.on, scope);
    const pairs = equiPairs(item.on, sources, L.idx, R.idx);
    const out = [];
    const merge = (l, r) => { const t = l.slice(); for (const i of R.idx) t[i] = r ? r[i] : null; return t; };
    const matches = (l, r) => toBool(on({ tuple: merge(l, r) })) === true;
    if (pairs) {
      const rightKey = t => pairs.map(p => keyOf(p.right({ tuple: t }))).join('\u0001');
      const leftKey = t => pairs.map(p => keyOf(p.left({ tuple: t }))).join('\u0001');
      if (item.kind === 'RIGHT') {
        const map = new Map();
        for (const l of L.tuples) { if (pairs.some(p => keyOf(p.left({ tuple: l })) === null)) continue; push(map, leftKey(l), l); }
        for (const r of R.tuples) {
          const hits = pairs.some(p => keyOf(p.right({ tuple: r })) === null) ? [] : (map.get(rightKey(r)) || []).filter(l => matches(l, r));
          if (hits.length) for (const l of hits) out.push(merge(l, r));
          else { const t = new Array(n).fill(null); for (const i of R.idx) t[i] = r[i]; out.push(t); }
        }
      } else {
        const map = new Map();
        for (const r of R.tuples) { if (pairs.some(p => keyOf(p.right({ tuple: r })) === null)) continue; push(map, rightKey(r), r); }
        for (const l of L.tuples) {
          const hits = pairs.some(p => keyOf(p.left({ tuple: l })) === null) ? [] : (map.get(leftKey(l)) || []).filter(r => matches(l, r));
          if (hits.length) for (const r of hits) out.push(merge(l, r));
          else if (item.kind === 'LEFT') out.push(merge(l, null));
        }
      }
    } else {
      for (const l of item.kind === 'RIGHT' ? R.tuples : L.tuples) {
        let hit = false;
        for (const r of item.kind === 'RIGHT' ? L.tuples : R.tuples) {
          const [a, b] = item.kind === 'RIGHT' ? [r, l] : [l, r];
          if (matches(a, b)) { out.push(merge(a, b)); hit = true; }
        }
        if (!hit && item.kind === 'LEFT') out.push(merge(l, null));
        if (!hit && item.kind === 'RIGHT') { const t = new Array(n).fill(null); for (const i of R.idx) t[i] = l[i]; out.push(t); }
      }
    }
    return { idx: [...L.idx, ...R.idx], tuples: out };
  };

  let acc = null;
  for (const item of from) {
    const part = evalItem(item);
    if (!acc) { acc = part; continue; }
    const out = [];
    if (acc.tuples.length * part.tuples.length > 5_000_000) throw new BaseError('TooLarge');
    for (const a of acc.tuples) for (const b of part.tuples) { const t = a.slice(); for (const i of part.idx) t[i] = b[i]; out.push(t); }
    acc = { idx: [...acc.idx, ...part.idx], tuples: out };
  }
  return { sources, tuples: acc ? acc.tuples : [] };
}

function push(map, k, v) { const list = map.get(k); if (list) list.push(v); else map.set(k, [v]); }

/** The pairs of an ON condition that compare a field of the left side with one of the right side (A.x = B.y And …), compiled; null when it is something else. */
function equiPairs(on, sources, leftIdx, rightIdx) {
  const terms = [];
  const collect = e => { if (e.t === 'bin' && e.op === 'AND') { collect(e.l); collect(e.r); } else terms.push(e); };
  collect(on);
  const pairs = [];
  for (const t of terms) {
    if (t.t !== 'bin' || t.op !== '=' || t.l.t !== 'id' || t.r.t !== 'id') continue;
    const a = resolveField(sources, t.l.parts);
    const b = resolveField(sources, t.r.parts);
    if (!a || !b) continue;
    const scopeL = { resolve: () => env => valueOf(env.tuple[a.idx], a.f) };
    const scopeR = { resolve: () => env => valueOf(env.tuple[b.idx], b.f) };
    if (leftIdx.includes(a.idx) && rightIdx.includes(b.idx)) pairs.push({ left: compile(t.l, scopeL), right: compile(t.r, scopeR) });
    else if (leftIdx.includes(b.idx) && rightIdx.includes(a.idx)) pairs.push({ left: compile(t.r, scopeR), right: compile(t.l, scopeL) });
  }
  return pairs.length ? pairs : null;
}

/** Finds a field among the sources: Table.Field or Field (an error when several tables have it). */
function resolveField(sources, parts) {
  if (parts.length === 2) {
    const [a, name] = parts.map(lower);
    for (let i = 0; i < sources.length; i++) {
      const s = sources[i];
      if (lower(s.alias) !== a && !(s.table && lower(s.table.name) === a && !s.aliasOnly)) continue;
      const f = s.fields.find(x => lower(x.name) === name);
      if (f) return { idx: i, f, src: s };
    }
    return null;
  }
  const name = lower(parts[0]);
  let hit = null;
  for (let i = 0; i < sources.length; i++) {
    const f = sources[i].fields.find(x => lower(x.name) === name);
    if (f) {
      if (hit) throw new BaseError('Ambiguous', parts[0]);
      hit = { idx: i, f, src: sources[i] };
    }
  }
  return hit;
}

function scopeFor(sources, ctx, aliases) {
  const compiling = new Set();
  const scope = {
    lang: ctx.lang, now: ctx.now, currency: ctx.currency, params: ctx.params,
    resolve(parts) {
      const r = resolveField(sources, parts);
      if (r) return env => valueOf(env.tuple?.[r.idx], r.f);
      if (parts.length === 1 && aliases?.has(lower(parts[0])) && !compiling.has(lower(parts[0]))) {
        const key = lower(parts[0]);
        compiling.add(key);
        try { return compile(aliases.get(key), scope); } finally { compiling.delete(key); }
      }
      const pkey = parts.join('.').toLowerCase();
      if (ctx.params && pkey in ctx.params) return null;
      if (ctx.collect) {
        if (!ctx.collect.has(pkey)) ctx.collect.set(pkey, { name: parts.join('.'), type: '' });
        return () => null;
      }
      return null;
    },
    domain: (fn, exprText, domainName, criteria) => domainAggregate(ctx.db, fn, exprText, domainName, criteria, ctx),
  };
  return scope;
}

/** DLookup, DCount, DSum … : a small query of its own. */
function domainAggregate(db, fn, exprText, domainName, criteria, ctx) {
  if (ctx.dry) return null;
  const agg = { DLookup: 'First', DCount: 'Count', DSum: 'Sum', DAvg: 'Avg', DMin: 'Min', DMax: 'Max', DFirst: 'First', DLast: 'Last' }[fn];
  const e = exprText.trim() === '*' ? { t: 'star', table: null } : Parser.expression(exprText);
  const where = criteria ? Parser.expression(criteria) : null;
  const stmt = {
    type: 'select', distinct: null, top: null, columns: [{ e: { t: 'fn', name: agg, args: [e] }, alias: 'X' }], into: null,
    from: [{ t: 'table', name: domainName, alias: null }], where, groupBy: [], having: null, orderBy: [], params: [],
  };
  const res = execSelect(db, stmt, { ...ctx, depth: ctx.depth + 1 });
  const v = res.rows[0]?.v[0] ?? null;
  return fn === 'DCount' && v === null ? 0 : v;
}

// ---- aggregates ------------------------------------------------------------------------------------------------------------------

function aggregateOf(name, values, lang) {
  const present = values.filter(v => !isNull(v));
  switch (name) {
    case 'Count': return present.length;
    case 'Sum': {
      if (!present.length) return null;
      let s = 0;
      for (const v of present) s += toNumber(v, lang);
      return Math.round(s * 1e10) / 1e10;
    }
    case 'Avg': {
      if (!present.length) return null;
      let s = 0;
      for (const v of present) s += toNumber(v, lang);
      return s / present.length;
    }
    case 'Min': case 'Max': {
      if (!present.length) return null;
      let best = present[0];
      for (const v of present) { const k = compare(v, best, lang); if (name === 'Min' ? k < 0 : k > 0) best = v; }
      return best;
    }
    case 'First': return values.length ? values[0] ?? null : null;
    case 'Last': return values.length ? values[values.length - 1] ?? null : null;
    case 'StDev': case 'StDevP': case 'Var': case 'VarP': {
      const nums = present.map(v => toNumber(v, lang));
      const sample = name === 'StDev' || name === 'Var';
      if (nums.length < (sample ? 2 : 1)) return null;
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - (sample ? 1 : 0));
      return name.startsWith('StDev') ? Math.sqrt(variance) : variance;
    }
    default: return null;
  }
}

/** Sum/Avg/… over a list of rows (also used by reports and the totals row). */
export function aggregate(name, values, lang = 'en') { return aggregateOf(name, values, lang); }

// ---- SELECT ----------------------------------------------------------------------------------------------------------------------

const NAME_OF_AGG = { Sum: 'SumOf', Avg: 'AvgOf', Count: 'CountOf', Min: 'MinOf', Max: 'MaxOf', First: 'FirstOf', Last: 'LastOf', StDev: 'StDevOf', StDevP: 'StDevPOf', Var: 'VarOf', VarP: 'VarPOf' };

/** The name Access gives a column without AS: the field, SumOfQty, or Expr1000. */
export function columnName(e, n) {
  if (e.t === 'id') return e.parts[e.parts.length - 1];
  if (e.t === 'fn' && NAME_OF_AGG[e.name] && e.args[0]?.t === 'id') return NAME_OF_AGG[e.name] + e.args[0].parts[e.args[0].parts.length - 1];
  return 'Expr' + (1000 + n);
}

function containsAggregate(e) {
  let found = false;
  walk(e, x => { if (x.t === 'fn' && AGGREGATES.has(x.name)) found = true; });
  return found;
}

function execSelect(db, s, ctx) {
  const { sources, tuples: all } = buildFrom(db, s.from, ctx);

  // the columns (* stands for every field of every source, Table.* for those of one)
  const cols = [];
  let exprN = 0;
  for (const c of s.columns) {
    if (c.e.t === 'star') {
      const list = c.e.table ? sources.filter(x => lower(x.alias) === lower(c.e.table) || (x.table && lower(x.table.name) === lower(c.e.table))) : sources;
      if (c.e.table && !list.length) throw new BaseError('NoTable', c.e.table);
      for (const src of list) {
        for (const f of src.fields) {
          const dup = cols.some(x => lower(x.name) === lower(f.name)) || (!c.e.table && sources.filter(o => o.fields.some(y => lower(y.name) === lower(f.name))).length > 1);
          cols.push({ e: { t: 'id', parts: [src.alias, f.name] }, name: dup ? src.alias + '.' + f.name : f.name });
        }
      }
    } else {
      const name = c.alias || columnName(c.e, exprN);
      if (!c.alias && c.e.t !== 'id' && !(c.e.t === 'fn' && NAME_OF_AGG[c.e.name] && c.e.args[0]?.t === 'id')) exprN++;
      cols.push({ e: c.e, name, alias: c.alias });
    }
  }
  const aliases = new Map(cols.filter(c => c.alias).map(c => [lower(c.alias), c.e]));
  const scope = scopeFor(sources, ctx, aliases);

  let tuples = all;
  if (s.where) {
    if (containsAggregate(s.where)) throw new BaseError('AggregateInWhere');
    const w = compile(s.where, scope);
    tuples = tuples.filter(t => { const v = w({ tuple: t }); return v !== null && toBool(v) === true; });
  }

  const grouped = s.groupBy.length > 0 || cols.some(c => containsAggregate(c.e)) || (s.having && containsAggregate(s.having));
  let rows;
  let sortKeys = null;

  if (grouped) {
    checkGrouping(s, cols, aliases, sources);
    const aggScope = {
      ...scope,
      aggregate: (name, arg) => {
        if (containsAggregate(arg)) throw new BaseError('NestedAggregate');
        if (arg.t === 'star') return env => env.group.length;
        const f = compile(arg, scope);
        return env => aggregateOf(name, env.group.map(t => f({ tuple: t })), ctx.lang);
      },
    };
    aggScope.resolve = parts => {
      const r = scope.resolve(parts);
      return r ? env => r({ tuple: env.tuple, group: env.group }) : r;
    };
    const keyFns = s.groupBy.map(g => compile(g, scope));
    const groups = new Map();
    for (const t of tuples) {
      const k = keyFns.map(f => keyOf(f({ tuple: t })) ?? '\u0000').join('\u0001');
      const g = groups.get(k);
      if (g) g.push(t); else groups.set(k, [t]);
    }
    let list = [...groups.values()];
    if (!s.groupBy.length && !list.length) list = [[]];
    const colFns = cols.map(c => compile(c.e, aggScope));
    const empty = new Array(sources.length).fill(null);
    let envs = list.map(g => ({ tuple: g[0] || empty, group: g }));
    if (s.having) {
      const h = compile(s.having, aggScope);
      envs = envs.filter(env => { const v = h(env); return v !== null && toBool(v) === true; });
    }
    rows = envs.map(env => ({ v: colFns.map(f => f(env)), refs: null, env }));
    if (s.orderBy.length) {
      const fns = s.orderBy.map(o => orderFn(o.e, cols, aggScope));
      sortKeys = rows.map(r => fns.map(f => f(r.env, r)));
    }
  } else {
    const colFns = cols.map(c => compile(c.e, scope));
    rows = tuples.map(t => { const env = { tuple: t }; return { v: colFns.map(f => f(env)), refs: t, env }; });
    if (s.orderBy.length) {
      const fns = s.orderBy.map(o => orderFn(o.e, cols, scope));
      sortKeys = rows.map(r => fns.map(f => f(r.env, r)));
    }
  }

  if (sortKeys) {
    const desc = s.orderBy.map(o => o.desc);
    const order = rows.map((_, i) => i);
    const lang = ctx.lang;
    order.sort((a, b) => {
      for (let k = 0; k < desc.length; k++) {
        const c = sortCompare(sortKeys[a][k], sortKeys[b][k], lang);
        if (c) return desc[k] ? -c : c;
      }
      return a - b;
    });
    rows = order.map(i => rows[i]);
    sortKeys = order.map(i => sortKeys[i]);
  }

  if (s.distinct === 'DISTINCT') {
    const seen = new Set();
    const keep = [];
    const keepKeys = [];
    rows.forEach((r, i) => {
      const k = r.v.map(v => keyOf(v) ?? '\u0000').join('\u0001');
      if (!seen.has(k)) { seen.add(k); keep.push(r); if (sortKeys) keepKeys.push(sortKeys[i]); }
    });
    rows = keep;
    if (sortKeys) sortKeys = keepKeys;
  }

  if (s.top) {
    let n = s.top.percent ? Math.ceil(rows.length * s.top.n / 100) : Math.floor(s.top.n);
    if (sortKeys && n > 0 && n < rows.length) {
      const last = sortKeys[n - 1];
      while (n < rows.length && sortKeys[n].every((v, k) => sortCompare(v, last[k], ctx.lang) === 0)) n++;
    }
    rows = rows.slice(0, n);
  }

  // describe the columns: the field behind a plain name (format, lookup, editable) or the type of the values
  const single = !grouped && !s.distinct && sources.length >= 1 && sources.every(x => x.table);
  const columns = cols.map((c, i) => {
    let field = null;
    let source = null;
    if (c.e.t === 'id') {
      try {
        const r = resolveField(sources, c.e.parts);
        if (r) { field = r.f.field || null; source = r.idx; }
      } catch { /* ambiguous names were reported already */ }
    }
    const known = field ? null : exprType(c.e, sources, aliases);
    const type = field?.type || (known && known !== 'text' ? known : inferType(rows, i));
    const editable = single && source !== null && !!sources[source].table && field && field.type !== 'autonumber';
    return { name: c.name, type, field, source, key: field?.id ?? null, editable: !!editable };
  });
  for (const r of rows) delete r.env;
  const base = single && sources.length === 1 ? sources[0].table : null;
  return { columns, rows, updatable: single, base, sources: sources.map(x => ({ alias: x.alias, table: x.table?.id || null })) };
}

function orderFn(e, cols, scope) {
  // ORDER BY an alias of the SELECT, or an expression that is a column of it
  if (e.t === 'id' && e.parts.length === 1) {
    const i = cols.findIndex(c => c.alias && lower(c.alias) === lower(e.parts[0]));
    if (i >= 0) return (env, row) => row.v[i];
  }
  const same = cols.findIndex(c => sameExpr(c.e, e));
  if (same >= 0) return (env, row) => row.v[same];
  const f = compile(e, scope);
  return env => f(env);
}

/** Access: every field outside an aggregate must be part of GROUP BY. */
function checkGrouping(s, cols, aliases, sources) {
  const ok = e => {
    if (!e || typeof e !== 'object') return true;
    if (s.groupBy.some(g => sameExpr(g, e))) return true;
    if (e.t === 'fn' && AGGREGATES.has(e.name)) return true;
    if (e.t === 'id') {
      if (e.parts.length === 1 && aliases.has(lower(e.parts[0]))) return true;
      let r = null;
      try { r = resolveField(sources, e.parts); } catch { r = null; }
      if (!r) return true; // a parameter
      return s.groupBy.some(g => g.t === 'id' && sameField(g, e, sources));
    }
    for (const v of Object.values(e)) {
      if (Array.isArray(v)) { if (!v.every(ok)) return false; } else if (v && typeof v === 'object' && v.t && !ok(v)) return false;
    }
    return true;
  };
  for (const c of cols) if (!ok(c.e)) throw new BaseError('NotInGroup', exprToText(c.e));
  for (const o of s.orderBy) if (!ok(o.e)) throw new BaseError('NotInGroup', exprToText(o.e));
}

function sameField(a, b, sources) {
  try {
    const x = resolveField(sources, a.parts);
    const y = resolveField(sources, b.parts);
    return x && y && x.idx === y.idx && x.f === y.f;
  } catch { return false; }
}

const DATE_FNS = new Set(['DATE', 'NOW', 'TIME', 'CDATE', 'DATEADD', 'DATESERIAL', 'TIMESERIAL', 'DATEVALUE', 'TIMEVALUE']);
const TEXT_FNS = new Set(['LEFT', 'RIGHT', 'MID', 'UCASE', 'LCASE', 'TRIM', 'LTRIM', 'RTRIM', 'REPLACE', 'SPACE', 'STRING', 'STRREVERSE', 'CHR', 'CSTR', 'STR', 'FORMAT',
  'FORMATCURRENCY', 'FORMATNUMBER', 'FORMATPERCENT', 'FORMATDATETIME', 'MONTHNAME', 'WEEKDAYNAME']);
const YESNO_FNS = new Set(['ISNULL', 'ISNUMERIC', 'ISDATE', 'CBOOL']);

/**
 * The type of the values of an expression column, as far as it follows from the expression: CCur(…) is Currency, a date function Date/Time, a
 * comparison Yes/No, arithmetic with Currency stays Currency … (null: unknown; the values decide).
 */
function exprType(e, sources, aliases = null, depth = 0) {
  if (!e || depth > 20) return null;
  const of = x => exprType(x, sources, aliases, depth + 1);
  switch (e.t) {
    case 'num': return 'number';
    case 'str': return 'text';
    case 'bool': return 'yesno';
    case 'date': return 'datetime';
    case 'id': {
      let r = null;
      try { r = resolveField(sources, e.parts); } catch { r = null; }
      if (r) { const ty = r.f.field?.type || r.f.type || null; return ty === 'autonumber' ? 'number' : ty; }
      if (e.parts.length === 1 && aliases?.has(lower(e.parts[0]))) return of(aliases.get(lower(e.parts[0])));
      return null;
    }
    case 'un': return e.op === 'NOT' ? 'yesno' : of(e.e);
    case 'bin': {
      if (e.op === '&') return 'text';
      if (['=', '<>', '<', '>', '<=', '>=', 'AND', 'OR', 'XOR', 'EQV', 'IMP'].includes(e.op)) return 'yesno';
      const a = of(e.l);
      const b = of(e.r);
      if (e.op === '-' && a === 'datetime' && b === 'datetime') return 'number';
      if ((e.op === '+' || e.op === '-') && (a === 'datetime' || b === 'datetime')) return 'datetime';
      if (a === 'currency' || b === 'currency') return e.op === '/' && a === 'currency' && b === 'currency' ? 'number' : 'currency';
      return 'number';
    }
    case 'like': case 'between': case 'in': case 'isnull': return 'yesno';
    case 'fn': {
      const n = e.name.toUpperCase();
      if (n === 'CCUR') return 'currency';
      if (n === 'COUNT' || n === 'STDEV' || n === 'STDEVP' || n === 'VAR' || n === 'VARP') return 'number';
      if (['SUM', 'AVG', 'MIN', 'MAX', 'FIRST', 'LAST', 'NZ', 'ROUND', 'ABS', 'INT', 'FIX'].includes(n)) {
        const ty = of(e.args[0]);
        if (n === 'SUM' || n === 'AVG') return ty === 'currency' ? 'currency' : 'number';
        if (n === 'ROUND' || n === 'ABS' || n === 'INT' || n === 'FIX') return ty === 'currency' ? 'currency' : 'number';
        return ty;
      }
      if (n === 'IIF') { const a = of(e.args[1]); const b = of(e.args[2]); return a === b ? a : a && b && a !== 'text' && b !== 'text' && (a === 'currency' || b === 'currency') ? 'currency' : null; }
      if (DATE_FNS.has(n)) return 'datetime';
      if (TEXT_FNS.has(n)) return 'text';
      if (YESNO_FNS.has(n)) return 'yesno';
      if (FUNCTIONS_NUMERIC.has(n)) return 'number';
      return null;
    }
    default: return null;
  }
}

const FUNCTIONS_NUMERIC = new Set(['LEN', 'INSTR', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'DATEDIFF', 'DATEPART', 'VAL', 'CINT', 'CLNG', 'CDBL',
  'CSNG', 'CDEC', 'SGN', 'SQR', 'EXP', 'LOG', 'ASC', 'STRCOMP']);

function inferType(rows, i) {
  for (const r of rows) {
    const v = r.v[i];
    if (isNull(v)) continue;
    if (v instanceof DateVal) return 'datetime';
    if (typeof v === 'boolean') return 'yesno';
    if (typeof v === 'number') return 'number';
    return String(v).length > 255 ? 'memo' : 'text';
  }
  return 'text';
}

// ---- action queries --------------------------------------------------------------------------------------------------------------

/**
 * Plans an action query: { type, count, table, apply() → report } where report = { done, typeFailures, keyViolations, ruleViolations, lockFailures }.
 * The editor asks "You are about to update N row(s)" with count, then calls apply(). Make-table: { exists, blocked } too.
 */
export function planAction(db, stmt, options = {}) {
  const ctx = context(db, options);
  switch (stmt.type) {
    case 'update': return planUpdate(db, stmt, ctx);
    case 'delete': return planDeleteQuery(db, stmt, ctx);
    case 'insert': return planAppend(db, stmt, ctx);
    case 'select': if (stmt.into) return planMakeTable(db, stmt, ctx); break;
    default: break;
  }
  throw new BaseError('NotAction');
}

function targetTable(db, stmt, sources, name) {
  if (name) {
    const s = sources.find(x => lower(x.alias) === lower(name) || (x.table && lower(x.table.name) === lower(name)));
    if (!s?.table) throw new BaseError('NoTable', name);
    return sources.indexOf(s);
  }
  const tables = sources.map((s, i) => [s, i]).filter(([s]) => s.table);
  if (tables.length !== 1) throw new BaseError('WhichTable');
  return tables[0][1];
}

function planUpdate(db, stmt, ctx) {
  const { sources, tuples: all } = buildFrom(db, stmt.from, ctx);
  const scope = scopeFor(sources, ctx, null);
  const sets = stmt.set.map(s => {
    const r = resolveField(sources, s.target.parts);
    if (!r || !r.src.table) throw new BaseError('UnknownName', s.target.parts.join('.'));
    return { idx: r.idx, field: r.f.field, fn: compile(s.e, scope) };
  });
  const where = stmt.where ? compile(stmt.where, scope) : null;
  const tuples = where ? all.filter(t => toBool(where({ tuple: t }) ?? false) === true) : all;
  const tableIdx = sets[0].idx;
  if (sets.some(x => x.idx !== tableIdx)) {
    // several tables: each change goes to its own table
  }
  const changes = new Map(); // "idx:rowId" → { idx, row, values }
  for (const t of tuples) {
    const env = { tuple: t };
    for (const set of sets) {
      const row = t[set.idx];
      if (!row) continue;
      const key = set.idx + ':' + row.id;
      let ch = changes.get(key);
      if (!ch) { ch = { idx: set.idx, row, values: {} }; changes.set(key, ch); }
      ch.values[set.field.id] = set.fn(env);
    }
  }
  const table = sources[tableIdx].table;
  return {
    type: 'update', table, count: changes.size,
    apply() {
      const report = newReport();
      for (const ch of changes.values()) {
        const t = sources[ch.idx].table;
        const values = {};
        let failed = false;
        for (const [fid, v] of Object.entries(ch.values)) {
          const f = t.fields.find(x => x.id === fid);
          const c = coerce(f, v, ctx.lang);
          if (!c.ok) { report.typeFailures++; failed = true; break; }
          values[fid] = c.value;
        }
        if (failed) continue;
        const r = saveRecord(db, t, ch.row.id, values, ctx.lang);
        if (r.ok) report.done++; else countFailure(report, r.error);
      }
      return report;
    },
  };
}

function planDeleteQuery(db, stmt, ctx) {
  const { sources, tuples: all } = buildFrom(db, stmt.from, ctx);
  const scope = scopeFor(sources, ctx, null);
  const idx = targetTable(db, stmt, sources, stmt.target);
  const where = stmt.where ? compile(stmt.where, scope) : null;
  const tuples = where ? all.filter(t => toBool(where({ tuple: t }) ?? false) === true) : all;
  const ids = [...new Set(tuples.map(t => t[idx]?.id).filter(id => id !== undefined && id !== null))];
  const table = sources[idx].table;
  return {
    type: 'delete', table, count: ids.length,
    apply() {
      const report = newReport();
      const plan = planDelete(db, table, ids);
      if (!plan.ok) { report.keyViolations = ids.length; report.error = plan.error; return report; }
      applyDelete(db, plan.plan);
      report.done = ids.length;
      report.cascaded = plan.cascaded;
      return report;
    },
  };
}

function planAppend(db, stmt, ctx) {
  const table = tableByName(db, stmt.table);
  if (!table) throw new BaseError('NoTable', stmt.table);
  let names;
  let rows;
  if (stmt.values) {
    const scope = scopeFor([], ctx, null);
    rows = [stmt.values.map(v => compile(v, scope)({ tuple: [] }))];
    names = stmt.columns || table.fields.slice(0, stmt.values.length).map(f => f.name);
  } else {
    const res = execSelect(db, stmt.select, ctx);
    rows = res.rows.map(r => r.v);
    names = stmt.columns || res.columns.map(c => c.name);
  }
  const fields = names.map(n => {
    const f = fieldByName(table, n);
    if (!f) throw new BaseError('UnknownName', stmt.table + '.' + n);
    return f;
  });
  return {
    type: 'append', table, count: rows.length,
    apply() {
      const report = newReport();
      for (const v of rows) {
        if (totalRecords(db) >= LIMITS.records) { report.keyViolations += 1; continue; }
        const values = {};
        let failed = false;
        fields.forEach((f, i) => {
          if (failed) return;
          const c = coerce(f, v[i], ctx.lang);
          if (!c.ok) { failed = true; return; }
          values[f.id] = c.value;
        });
        if (failed) { report.typeFailures++; continue; }
        const r = saveRecord(db, table, null, values, ctx.lang);
        if (r.ok) report.done++; else countFailure(report, r.error);
      }
      return report;
    },
  };
}

function planMakeTable(db, stmt, ctx) {
  const res = execSelect(db, { ...stmt, into: null }, ctx);
  const existing = tableByName(db, stmt.into);
  const blocked = existing && db.relationships.some(r => r.parent.table === existing.id || r.child.table === existing.id) ? { code: 'TableInRelationship', args: [existing.name] } : null;
  if (queryByName(db, stmt.into)) throw new BaseError('NameTaken', stmt.into);
  return {
    type: 'maketable', name: stmt.into, count: res.rows.length, exists: !!existing, blocked,
    apply() {
      const report = newReport();
      if (blocked) { report.error = blocked; return report; }
      if (totalRecords(db) - (existing?.rows.length || 0) + res.rows.length > LIMITS.records) { report.error = { code: 'TooManyRecords', args: [LIMITS.records] }; return report; }
      const table = makeTableFrom(db, stmt.into, res, existing);
      report.done = table.rows.length;
      report.table = table;
      return report;
    },
  };
}

/** A new table from a result (make-table query, an exported query): the fields take the types of the columns. */
export function makeTableFrom(db, name, res, replace = null) {
  const fields = [];
  const used = new Set();
  res.columns.forEach((c, i) => {
    let fname = String(c.name).replace(/[.!`[\]]/g, '_').slice(0, 64) || 'Field' + (i + 1);
    while (used.has(lower(fname))) fname = fname.slice(0, 60) + (i + 1);
    used.add(lower(fname));
    let type = c.type || 'text';
    if (type === 'autonumber') type = 'number';
    const f = newField(type, fname, 'f' + (i + 1));
    if (c.field) {
      if (c.field.type === 'text') f.size = c.field.size;
      if (c.field.type === 'number') f.numberSize = c.field.numberSize;
      f.format = c.field.type === 'autonumber' ? '' : c.field.format || f.format;
      f.decimals = c.field.decimals ?? null;
      f.caption = c.field.caption || '';
    } else if (type === 'number') f.numberSize = 'double';
    if (type === 'text') {
      const longest = res.rows.reduce((m, r) => Math.max(m, isNull(r.v[i]) ? 0 : String(r.v[i]).length), 0);
      if (longest > 255) { f.type = 'memo'; delete f.size; }
    }
    fields.push(f);
  });
  const table = {
    id: replace ? replace.id : newObjectId(db, 'tables'), name, description: '', fields, primaryKey: [], rows: [], nextId: 1, datasheet: {},
  };
  for (const r of res.rows) {
    const v = {};
    fields.forEach((f, i) => {
      const c = coerce(f, r.v[i], 'en');
      const value = c.ok ? c.value : c.value ?? null;
      if (value !== null && value !== undefined) v[f.id] = value;
    });
    table.rows.push({ id: table.nextId++, v });
  }
  if (replace) {
    const i = db.tables.indexOf(replace);
    db.tables[i] = table;
  } else db.tables.push(table);
  return table;
}

function newReport() { return { done: 0, typeFailures: 0, keyViolations: 0, ruleViolations: 0, cascaded: 0, error: null }; }

function countFailure(report, error) {
  if (['Duplicate', 'NoParent', 'HasChildren', 'PkNull', 'TooManyRecords'].includes(error.code)) report.keyViolations++;
  else report.ruleViolations++;
  report.error ??= error;
}

export { sqlExpr, newId, childRelationships, parentRelationships, isNumericType, collator };
