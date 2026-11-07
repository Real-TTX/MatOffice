/* MatOffice Base – the design grid of a query (QBE: Field, Table, Total, Sort, Show, Criteria, or) and the SQL it stands for, both ways, so the
   Design view and the SQL view stay in step for every query the grid can show (like Access). A pure module without the DOM.

   design = {
     type: 'select'|'maketable'|'append'|'update'|'delete', target: '' (new table / table to append to), distinct, top: '' | '5' | '25%', totals,
     tables: [{ name, alias }], joins: [{ left, leftField, right, rightField, kind: 'INNER'|'LEFT'|'RIGHT' }] (left/right are aliases or names),
     columns: [{ field: 'City' | '*' | 'Total: [Price]*[Qty]', table: 'Customers'|'', total: 'Group By'|'Sum'|…|'Expression'|'Where', sort: ''|'asc'|'desc',
                 show: true, criteria: ['"Berlin"', …], appendTo: '', updateTo: '', del: 'From'|'Where' }],
     params: [{ name, type }]
   } */
import { Parser, BaseError, exprToText, sameExpr, mapExpr, AGGREGATES, quoteName, walk } from './expr.js';
import { parseCriteria, criteriaText, targetOf } from './criteria.js';
import { toSql, parseSql } from './sql.js';
import { tableByName, queryByName, lower } from './model.js';

export const TOTALS = ['Group By', 'Sum', 'Avg', 'Min', 'Max', 'Count', 'StDev', 'Var', 'First', 'Last', 'Expression', 'Where'];

export function emptyDesign() {
  return { type: 'select', target: '', distinct: false, top: '', totals: false, tables: [], joins: [], columns: [], params: [] };
}

export function emptyColumn() {
  return { field: '', table: '', total: 'Group By', sort: '', show: true, criteria: [], appendTo: '', updateTo: '', del: 'Where' };
}

/** The fields of a table or a query by name (a query's are the names of its result columns). */
export function sourceFields(db, name, run) {
  const t = tableByName(db, name);
  if (t) return t.fields.map(f => f.name);
  const q = queryByName(db, name);
  if (q && run) {
    try { return run(q).columns.map(c => c.name); } catch { return []; }
  }
  return [];
}

/** Splits "Alias: expression" (the colon outside of brackets, quotes and #dates#). */
export function splitAlias(text) {
  const s = String(text ?? '');
  let quote = null;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quote) { if (c === quote) quote = null; continue; }
    if (c === '[') quote = ']';
    else if (c === '"' || c === '#' || c === "'") quote = c;
    else if (c === ':') {
      const alias = s.slice(0, i).trim();
      if (alias && !/[[\]"#()]/.test(alias)) return { alias, expr: s.slice(i + 1).trim() };
      return { alias: null, expr: s.trim() };
    }
  }
  return { alias: null, expr: s.trim() };
}

const aliasOf = t => t.alias || t.name;

// ---- design → SQL --------------------------------------------------------------------------------------------------------------

/** The statement of a design (throws BaseError for a cell that cannot be read; err.column tells which column). */
export function designToStatement(design, db, fieldsOf) {
  const tables = design.tables;
  const fieldsByAlias = new Map(tables.map(t => [lower(aliasOf(t)), fieldsOf(t.name)]));
  const ownerOf = name => {
    const hits = tables.filter(t => (fieldsByAlias.get(lower(aliasOf(t))) || []).some(f => lower(f) === lower(name)));
    return hits.length === 1 ? aliasOf(hits[0]) : null;
  };
  const qualify = e => mapExpr(e, x => {
    if (x.t === 'id' && x.parts.length === 1) { const o = ownerOf(x.parts[0]); return o ? { t: 'id', parts: [o, x.parts[0]] } : x; }
    return undefined;
  });

  const from = buildFromTree(design);
  const select = [];
  const groupBy = [];
  const orderBy = [];
  const whereRows = [];
  const havingRows = [];
  const appendCols = [];
  const sets = [];
  let deleteTarget = null;
  let exprN = 0;
  const usedAliases = new Set(design.columns.map(c => splitAlias(c.field).alias).filter(Boolean).map(lower));

  design.columns.forEach((col, ci) => {
    if (!String(col.field || '').trim()) return;
    let e;
    let alias;
    try {
      ({ e, alias } = columnExpr(col, qualify, ownerOf));
    } catch (err) {
      if (err instanceof BaseError) err.column = ci;
      throw err;
    }
    const total = design.totals && design.type !== 'update' && design.type !== 'delete' ? col.total || 'Group By' : null;
    let out = e;
    if (total && AGGREGATES.has(total)) {
      out = { t: 'fn', name: total, args: [e] };
      if (!alias && e.t !== 'id' && e.t !== 'star') { do { exprN++; } while (usedAliases.has('expr' + exprN)); alias = 'Expr' + exprN; }
      if (!alias && e.t === 'id') alias = total + 'Of' + e.parts[e.parts.length - 1];
    } else if (!alias && e.t !== 'id' && e.t !== 'star' && (!total || total !== 'Where')) {
      do { exprN++; } while (usedAliases.has('expr' + exprN));
      alias = 'Expr' + exprN;
    }
    if (total === 'Group By') groupBy.push(e);

    const shown = design.type === 'select' || design.type === 'maketable' ? col.show !== false && total !== 'Where'
      : design.type === 'append' ? !!String(col.appendTo || '').trim() : false;
    if (shown) {
      select.push({ e: out, alias: alias && !(e.t === 'id' && lower(alias) === lower(e.parts[e.parts.length - 1]) && out === e) ? alias : null });
      if (design.type === 'append') appendCols.push(col.appendTo.trim());
    }
    if (col.sort && design.type !== 'update' && design.type !== 'delete') orderBy.push({ e: out.t === 'star' ? null : out, desc: col.sort === 'desc' });
    if (design.type === 'update' && String(col.updateTo || '').trim()) {
      if (e.t !== 'id') throw Object.assign(new BaseError('UpdateField'), { column: ci });
      let value;
      try { value = Parser.expression(col.updateTo); } catch (err) { if (err instanceof BaseError) err.column = ci; throw err; }
      sets.push({ target: e, e: value });
    }
    if (design.type === 'delete' && col.del === 'From' && e.t === 'star') deleteTarget = e.table;

    const critTarget = e.t === 'star' ? null : (total && AGGREGATES.has(total) ? out : e);
    (col.criteria || []).forEach((text, ri) => {
      if (!String(text || '').trim() || !critTarget) return;
      let cond;
      try { cond = parseCriteria(text, critTarget); } catch (err) { if (err instanceof BaseError) { err.column = ci; err.row = ri; } throw err; }
      const rows = total && total !== 'Where' ? havingRows : whereRows;
      (rows[ri] ||= []).push(cond);
    });
  });

  const orOf = rows => {
    const branches = rows.filter(r => r && r.length).map(terms => terms.reduce((a, b) => ({ t: 'bin', op: 'AND', l: a, r: b })));
    return branches.length ? branches.reduce((a, b) => ({ t: 'bin', op: 'OR', l: a, r: b })) : null;
  };
  const where = orOf(whereRows);
  const having = orOf(havingRows);
  const params = design.params || [];

  if (design.type === 'update') {
    if (!sets.length) throw new BaseError('NoUpdate');
    return { type: 'update', from, set: sets, where, params };
  }
  if (design.type === 'delete') {
    if (!deleteTarget && tables.length > 1) deleteTarget = aliasOf(tables[0]);
    return { type: 'delete', target: deleteTarget, from, where, params };
  }
  if (!select.length) throw new BaseError('NoColumns');
  const top = parseTop(design.top);
  const stmt = {
    type: 'select', distinct: design.distinct ? 'DISTINCT' : null, top, columns: select, into: design.type === 'maketable' ? design.target || null : null,
    from, where, groupBy, having, orderBy: orderBy.filter(o => o.e), params: design.type === 'append' ? [] : params,
  };
  if (design.type === 'maketable' && !stmt.into) throw new BaseError('NoTarget');
  if (design.type === 'append') {
    if (!design.target) throw new BaseError('NoTarget');
    return { type: 'insert', table: design.target, columns: appendCols, select: stmt, values: null, params };
  }
  return stmt;
}

function parseTop(text) {
  const s = String(text ?? '').trim();
  if (!s || /^all$/i.test(s)) return null;
  const m = /^(\d+(?:\.\d+)?)\s*(%)?$/.exec(s);
  if (!m || Number(m[1]) <= 0) return null;
  return { n: Number(m[1]), percent: !!m[2] };
}

/** The expression of a grid column: Field (+ Table) or "Alias: expression". */
function columnExpr(col, qualify, ownerOf, options = {}) {
  const { alias, expr } = splitAlias(col.field);
  const table = String(col.table || '').trim();
  if (expr === '*') return { e: { t: 'star', table: table || null }, alias: null };
  const plain = /^\[?([^[\]]+)\]?$/.exec(expr);
  if (plain && !alias && /^[^"#()+\-*/&<>=,]+$/.test(plain[1]) && !/^\d/.test(plain[1]) && (table || ownerOf(plain[1]))) {
    // a field: Field and Table rows of the grid
    if (/^\s*\[?[^[\]]+\]?\s*$/.test(expr) && !/\s(and|or|like|mod)\s/i.test(expr)) {
      return { e: { t: 'id', parts: [table || ownerOf(plain[1]), plain[1].trim()] }, alias: null };
    }
  }
  const e = Parser.expression(expr);
  if (!e) throw new BaseError('Syntax', expr);
  if (e.t === 'id' && e.parts.length === 1 && table) return { e: { t: 'id', parts: [table, e.parts[0]] }, alias };
  // an expression stays as it was typed ([Price]*[Qty]); only the comparisons of the grid qualify it
  return { e: options.qualifyAll ? qualify(e) : e, alias };
}

/** FROM of the tables and the join lines: A INNER JOIN B ON … nested like Access writes it, tables without lines after a comma. */
function buildFromTree(design) {
  const tables = design.tables.map(t => ({ t: 'table', name: t.name, alias: t.alias || null, key: lower(aliasOf(t)) }));
  if (!tables.length) throw new BaseError('NoTables');
  const lines = design.joins.map(j => ({ ...j, l: lower(j.left), r: lower(j.right) })).filter(j => tables.some(t => t.key === j.l) && tables.some(t => t.key === j.r));
  const used = new Set();
  const items = [];
  const eq = j => ({ t: 'bin', op: '=', l: { t: 'id', parts: [j.left, j.leftField] }, r: { t: 'id', parts: [j.right, j.rightField] } });
  for (const start of tables) {
    if (used.has(start.key)) continue;
    used.add(start.key);
    let tree = { t: 'table', name: start.name, alias: start.alias };
    const inTree = new Set([start.key]);
    for (;;) {
      const line = lines.find(j => !j.done && (inTree.has(j.l) !== inTree.has(j.r)));
      if (!line) break;
      const outerKey = inTree.has(line.l) ? line.r : line.l;
      const other = tables.find(t => t.key === outerKey);
      // every line between the tree and this table goes into the ON
      const all = lines.filter(j => !j.done && ((inTree.has(j.l) && j.r === outerKey) || (inTree.has(j.r) && j.l === outerKey)));
      all.forEach(j => { j.done = true; });
      const on = all.map(eq).reduce((a, b) => ({ t: 'bin', op: 'AND', l: a, r: b }));
      let kind = line.kind || 'INNER';
      if (kind !== 'INNER' && inTree.has(line.r)) kind = kind === 'LEFT' ? 'RIGHT' : 'LEFT';
      tree = { t: 'join', kind, left: tree, right: { t: 'table', name: other.name, alias: other.alias }, on };
      inTree.add(outerKey);
      used.add(outerKey);
    }
    items.push(tree);
  }
  return items;
}

/** The SQL text of a design. */
export function designToSql(design, db, fieldsOf) {
  return toSql(designToStatement(design, db, fieldsOf));
}

// ---- SQL → design --------------------------------------------------------------------------------------------------------------

/** The design of a statement, or null when the grid cannot show it (then Access stays in SQL view). */
export function statementToDesign(stmt, db, fieldsOf) {
  const design = emptyDesign();
  design.params = stmt.params || [];
  let from;
  let s = stmt;
  if (stmt.type === 'insert') {
    if (!stmt.select) return null;
    design.type = 'append';
    design.target = stmt.table;
    s = stmt.select;
    from = s.from;
  } else if (stmt.type === 'update') { design.type = 'update'; from = stmt.from; }
  else if (stmt.type === 'delete') { design.type = 'delete'; from = stmt.from; }
  else { design.type = stmt.into ? 'maketable' : 'select'; design.target = stmt.into || ''; from = stmt.from; }

  // tables and join lines
  const ok = from.every(item => collectFrom(item, design));
  if (!ok) return null;
  const tables = design.tables;
  const fieldsByAlias = new Map(tables.map(t => [lower(aliasOf(t)), fieldsOf(t.name)]));
  const aliasNames = new Map(tables.map(t => [lower(t.name), aliasOf(t)]));
  for (const t of tables) aliasNames.set(lower(aliasOf(t)), aliasOf(t));
  const ownerOf = name => {
    const hits = tables.filter(t => (fieldsByAlias.get(lower(aliasOf(t))) || []).some(f => lower(f) === lower(name)));
    return hits.length === 1 ? aliasOf(hits[0]) : null;
  };
  const qualify = e => mapExpr(e, x => {
    if (x.t === 'id' && x.parts.length === 1) { const o = ownerOf(x.parts[0]); return o ? { t: 'id', parts: [o, x.parts[0]] } : x; }
    if (x.t === 'id' && x.parts.length === 2) { const a = aliasNames.get(lower(x.parts[0])); return a ? { t: 'id', parts: [a, x.parts[1]] } : x; }
    return undefined;
  });
  const shortName = parts => {
    if (parts.length === 2 && ownerOf(parts[1])) return '[' + parts[1] + ']';
    return parts.length === 1 ? '[' + parts[0] + ']' : parts.map(p => '[' + p + ']').join('.');
  };
  const text = e => exprToText(e, { name: shortName });
  for (const j of design.joins) { j.left = aliasNames.get(lower(j.left)) || j.left; j.right = aliasNames.get(lower(j.right)) || j.right; }

  const columns = [];
  const add = col => { columns.push({ ...emptyColumn(), ...col }); return columns.length - 1; };
  const fieldCell = (e, alias) => {
    if (e.t === 'star') return { field: '*', table: e.table ? aliasNames.get(lower(e.table)) || e.table : '' };
    if (e.t === 'id' && e.parts.length === 2 && (!alias || lower(alias) === lower(e.parts[1]))) return { field: e.parts[1], table: e.parts[0] };
    return { field: (alias ? alias + ': ' : '') + text(e), table: e.t === 'id' && e.parts.length === 2 ? e.parts[0] : '' };
  };

  if (design.type === 'update') {
    for (const set of stmt.set) {
      const target = qualify(set.target);
      add({ ...fieldCell(target, null), updateTo: text(qualify(set.e)), show: false });
    }
  } else if (design.type === 'delete') {
    const target = stmt.target ? aliasNames.get(lower(stmt.target)) || stmt.target : tables.length === 1 ? aliasOf(tables[0]) : null;
    if (target) add({ field: '*', table: target, del: 'From', show: false });
  } else {
    design.distinct = s.distinct === 'DISTINCT';
    design.top = s.top ? String(s.top.n) + (s.top.percent ? '%' : '') : '';
    const groupBy = s.groupBy.map(qualify);
    const grouped = groupBy.length > 0 || s.columns.some(c => hasAgg(c.e)) || (s.having && hasAgg(s.having));
    design.totals = grouped;
    s.columns.forEach((c, i) => {
      const e = qualify(c.e);
      const appendTo = design.type === 'append' ? (stmt.columns?.[i] || c.alias || (e.t === 'id' ? e.parts[e.parts.length - 1] : '')) : '';
      if (!grouped) { add({ ...fieldCell(e, c.alias), appendTo }); return; }
      if (e.t === 'fn' && AGGREGATES.has(e.name) && e.args.length === 1 && e.args[0].t !== 'star' && !hasAgg(e.args[0]) && TOTALS.includes(e.name)) {
        const arg = e.args[0];
        const defaultAlias = arg.t === 'id' ? e.name + 'Of' + arg.parts[arg.parts.length - 1] : null;
        const alias = c.alias && lower(c.alias) !== lower(defaultAlias || '') ? c.alias : null;
        const cell = arg.t === 'id' && !alias ? fieldCell(arg, null) : { field: (alias || c.alias || 'Expr' + (i + 1)) + ': ' + text(arg), table: arg.t === 'id' ? arg.parts[0] : '' };
        if (arg.t === 'id' && alias) { cell.field = alias + ': ' + arg.parts[1]; cell.table = arg.parts[0]; }
        add({ ...cell, total: e.name, appendTo });
        return;
      }
      if (groupBy.some(g => sameExpr(g, e))) { add({ ...fieldCell(e, c.alias), total: 'Group By', appendTo }); return; }
      add({ ...fieldCell(e, c.alias), total: 'Expression', appendTo });
    });
    for (const g of groupBy) {
      if (!columns.some(col => col.total === 'Group By' && sameExpr(colExpr(col, qualify, ownerOf), g))) add({ ...fieldCell(g, null), total: 'Group By', show: false });
    }
    if (design.type === 'append') design.target = stmt.table;
  }

  // criteria: OR branches are the rows, AND terms go to their column
  const where = design.type === 'update' || design.type === 'delete' ? stmt.where : s.where;
  const placeCriteria = (cond, having) => {
    const rows = splitOr(cond);
    rows.forEach((branch, ri) => {
      for (const termRaw of splitAnd(branch)) {
        const term = qualify(termRaw);
        let target = targetOf(term);
        let ci = -1;
        if (target) {
          if (design.totals && !having) {
            ci = columns.findIndex(col => col.total === 'Where' && sameExpr(colExpr(col, qualify, ownerOf), target));
            if (ci < 0) ci = add({ ...fieldCell(target, null), total: 'Where', show: false });
          } else {
            ci = columns.findIndex(col => matchesTarget(col, target, design.totals, qualify, ownerOf));
            if (ci < 0) {
              if (design.totals) {
                const agg = target.t === 'fn' && AGGREGATES.has(target.name) && target.args[0]?.t !== 'star';
                ci = agg ? add({ ...fieldCell(target.args[0], null), total: target.name, show: false }) : add({ ...fieldCell(target, null), total: hasAgg(target) ? 'Expression' : 'Group By', show: false });
              } else if (design.type === 'delete') ci = add({ ...fieldCell(target, null), del: 'Where', show: false });
              else ci = add({ ...fieldCell(target, null), show: false });
            }
          }
        } else {
          // a condition that is not about one column: an expression column with the criteria True
          target = term;
          ci = add({ field: 'Expr' + (columns.length + 1) + ': ' + text(term), table: '', show: false, total: design.totals ? (having ? 'Expression' : 'Where') : 'Group By', del: 'Where' });
          const col = columns[ci];
          col.criteria[ri] = 'True';
          continue;
        }
        const col = columns[ci];
        const crit = criteriaText(term, target, { name: shortName });
        col.criteria[ri] = col.criteria[ri] ? col.criteria[ri] + ' And ' + crit : crit;
      }
    });
  };
  if (where) placeCriteria(where, false);
  if (s.having && design.type !== 'update' && design.type !== 'delete') placeCriteria(s.having, true);
  const rowCount = Math.max(0, ...columns.map(c => c.criteria.length));
  for (const col of columns) for (let i = 0; i < rowCount; i++) col.criteria[i] ??= '';

  // sorting: on the columns from left to right when the order fits, else hidden sort columns
  if (s.orderBy?.length && design.type !== 'update' && design.type !== 'delete') {
    const hits = s.orderBy.map(o => {
      const key = qualify(o.e);
      if (key.t === 'id' && key.parts.length === 1) {
        const i = columns.findIndex(col => lower(splitAlias(col.field).alias || '') === lower(key.parts[0]) && col.show !== false);
        if (i >= 0) return i;
      }
      return columns.findIndex(col => col.show !== false && sameExpr(outputExpr(col, design.totals, qualify, ownerOf), key));
    });
    const increasing = hits.every((h, i) => h >= 0 && (i === 0 || h > hits[i - 1]));
    if (increasing) hits.forEach((h, i) => { columns[h].sort = s.orderBy[i].desc ? 'desc' : 'asc'; });
    else {
      s.orderBy.forEach(o => {
        const key = qualify(o.e);
        const agg = key.t === 'fn' && AGGREGATES.has(key.name) && key.args[0]?.t !== 'star';
        const cell = agg ? fieldCell(key.args[0], null) : fieldCell(key, null);
        add({ ...cell, sort: o.desc ? 'desc' : 'asc', show: false, total: design.totals ? (agg ? key.name : hasAgg(key) ? 'Expression' : 'Group By') : 'Group By' });
      });
    }
  }

  design.columns = columns;
  return design;
}

function collectFrom(item, design) {
  if (item.t === 'table') {
    design.tables.push({ name: item.name, alias: item.alias || null });
    return true;
  }
  if (item.t !== 'join') return false;
  if (!collectFrom(item.left, design) || !collectFrom(item.right, design)) return false;
  const terms = splitAnd(item.on);
  for (const t of terms) {
    if (t.t !== 'bin' || t.op !== '=' || t.l.t !== 'id' || t.r.t !== 'id' || t.l.parts.length !== 2 || t.r.parts.length !== 2) return false;
    const rightNames = new Set(namesOf(item.right).map(lower));
    let [a, b] = [t.l, t.r];
    if (rightNames.has(lower(a.parts[0])) && !rightNames.has(lower(b.parts[0]))) [a, b] = [b, a];
    design.joins.push({ left: a.parts[0], leftField: a.parts[1], right: b.parts[0], rightField: b.parts[1], kind: item.kind });
  }
  return true;
}

function namesOf(item) {
  if (item.t === 'table') return [item.alias || item.name];
  if (item.t === 'join') return [...namesOf(item.left), ...namesOf(item.right)];
  return [];
}

function splitOr(e) { const out = []; const go = x => { if (x.t === 'bin' && x.op === 'OR') { go(x.l); go(x.r); } else out.push(x); }; go(e); return out; }
function splitAnd(e) { const out = []; const go = x => { if (x.t === 'bin' && x.op === 'AND') { go(x.l); go(x.r); } else out.push(x); }; go(e); return out; }
function hasAgg(e) { let f = false; walk(e, x => { if (x.t === 'fn' && AGGREGATES.has(x.name)) f = true; }); return f; }

function colExpr(col, qualify, ownerOf) {
  try { return columnExpr(col, qualify, ownerOf, { qualifyAll: true }).e; } catch { return null; }
}

function outputExpr(col, totals, qualify, ownerOf) {
  const e = colExpr(col, qualify, ownerOf);
  if (!e) return null;
  if (totals && AGGREGATES.has(col.total)) return { t: 'fn', name: col.total, args: [e] };
  return e;
}

function matchesTarget(col, target, totals, qualify, ownerOf) {
  if (totals && col.total === 'Where') return false;
  const out = outputExpr(col, totals, qualify, ownerOf);
  return !!out && sameExpr(out, target);
}

/** The design of a SQL text (null when the grid cannot show it). */
export function sqlToDesign(sql, db, fieldsOf) {
  return statementToDesign(parseSql(sql), db, fieldsOf);
}

/** A name of the grid as Access shows it after the cell is left: plain expressions get "Expr1: ". */
export function tidyFieldCell(text, columns) {
  const s = String(text ?? '').trim();
  if (!s) return '';
  const { alias, expr } = splitAlias(s);
  if (alias || expr === '*' || /^\[?[^[\]"#()+\-*/&<>=,]+\]?$/.test(expr)) return s;
  let n = 1;
  const taken = new Set(columns.map(c => lower(splitAlias(c.field).alias || '')));
  while (taken.has('expr' + n)) n++;
  return 'Expr' + n + ': ' + expr;
}

export { quoteName };
