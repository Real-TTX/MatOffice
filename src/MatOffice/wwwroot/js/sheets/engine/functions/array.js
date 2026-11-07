/* MatOffice Sheets – dynamic array functions (pure, no DOM): UNIQUE, SORT, SORTBY, FILTER, SEQUENCE, RANDARRAY and the functions that stack,
   cut and reshape arrays. An empty result is #CALC!; an optional argument given empty counts as omitted; missing cells are filled with #N/A. */
import { define, deref, matrix, makeMatrix, flat, Errors } from '../lib.js';
import { FormulaError, toNumber, toBool, sortCompare } from '../values.js';
import { optNumber } from './stat.js';

const isErr = v => v instanceof FormulaError;
const { VALUE, CALC, NA, NUM } = Errors;
const first = v => (Array.isArray(v) ? v[0][0] : v);
const omitted = v => v === undefined || v === null;
const transpose = m => m[0].map((_, c) => m.map(row => row[c]));
const padded = (row, width, filler) => (row.length < width ? [...row, ...new Array(width - row.length).fill(filler)] : row);
const MAX_CELLS = 2097152;

function optBool(v, ctx, def) {
  if (omitted(v)) return def;
  return toBool(first(deref(v, ctx)));
}

/** The size of a new array (rows × cols): #CALC! when empty, #VALUE! when negative, #NUM! when too large; else null. */
function badSize(rows, cols) {
  if (rows < 0 || cols < 0) return VALUE;
  if (!rows || !cols) return CALC;
  return rows > 1048576 || cols > 16384 || rows * cols > MAX_CELLS ? NUM : null;
}

// ---- unique, sort, filter ---------------------------------------------------------------------------------------------------------

/** The key of a value for UNIQUE: text without case, a blank apart from 0. */
const keyOf = v => (v === null ? 'e' : typeof v === 'number' ? 'n' + v : typeof v === 'string' ? 's' + v.toLowerCase() : typeof v === 'boolean' ? 'b' + v : 'x' + v.code);

define('UNIQUE', {
  min: 1, max: 3, category: 'lookup', args: 'array, [by_col], [exactly_once]',
  fn: ([array, byCol, once], ctx) => {
    if (isErr(array)) return array;
    const bc = optBool(byCol, ctx, false);
    if (isErr(bc)) return bc;
    const eo = optBool(once, ctx, false);
    if (isErr(eo)) return eo;
    const m = bc ? transpose(matrix(array, ctx)) : matrix(array, ctx);
    const seen = new Map();
    for (const row of m) {
      const k = row.map(keyOf).join('\u0000');
      const entry = seen.get(k);
      if (entry) entry.n++;
      else seen.set(k, { row, n: 1 });
    }
    const out = [...seen.values()].filter(e => !eo || e.n === 1).map(e => e.row);
    return !out.length ? CALC : bc ? transpose(out) : out;
  },
});

/** The order of SORT/SORTBY: Excel's sort order (numbers, text without case, logicals, errors), reversed when descending; blanks stay last. */
function order(a, b, descending) {
  const blankA = a === null || a === undefined || a === '';
  const blankB = b === null || b === undefined || b === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  const c = sortCompare(a, b);
  return descending ? -c : c;
}

define('SORT', {
  min: 1, max: 4, category: 'lookup', args: 'array, [sort_index], [sort_order], [by_col]',
  fn: ([array, index, sortOrder, byCol], ctx) => {
    if (isErr(array)) return array;
    const bc = optBool(byCol, ctx, false);
    if (isErr(bc)) return bc;
    const m = bc ? transpose(matrix(array, ctx)) : matrix(array, ctx);
    const indexes = omitted(index) ? [1] : flat(index, ctx).map(toNumber);
    const orders = omitted(sortOrder) ? [1] : flat(sortOrder, ctx).map(toNumber);
    const keys = [];
    for (let i = 0; i < indexes.length; i++) {
      const k = indexes[i];
      const o = orders[Math.min(i, orders.length - 1)];
      if (isErr(k)) return k;
      if (isErr(o)) return o;
      const col = Math.trunc(k) - 1;
      if (col < 0 || col >= m[0].length || (o !== 1 && o !== -1)) return VALUE;
      keys.push({ col, descending: o === -1 });
    }
    const out = m.map((row, i) => ({ row, i }));
    out.sort((p, q) => {
      for (const k of keys) {
        const c = order(p.row[k.col], q.row[k.col], k.descending);
        if (c) return c;
      }
      return p.i - q.i;
    });
    const rows = out.map(x => x.row);
    return bc ? transpose(rows) : rows;
  },
});

/** SORTBY(array, by_array1, [sort_order1], …): the by arrays are columns as tall as the array (sorting rows) or rows as wide (sorting columns). */
define('SORTBY', {
  min: 2, max: 255, category: 'lookup', args: 'array, by_array1, [sort_order1], [by_array2, sort_order2], …',
  fn: ([array, ...rest], ctx) => {
    if (isErr(array)) return array;
    const m = matrix(array, ctx);
    const keys = [];
    let across = null;
    for (let i = 0; i < rest.length; i += 2) {
      if (isErr(rest[i])) return rest[i];
      const b = matrix(rest[i], ctx);
      const o = optNumber(rest[i + 1], ctx, 1);
      if (isErr(o)) return o;
      if (o !== 1 && o !== -1) return VALUE;
      const isColumn = b[0].length === 1 && b.length === m.length;
      const isRow = b.length === 1 && b[0].length === m[0].length;
      if (!isColumn && !isRow) return VALUE;
      const horizontal = isColumn && isRow ? across ?? false : isRow;
      if (across !== null && across !== horizontal) return VALUE;
      across = horizontal;
      keys.push({ values: horizontal ? b[0] : b.map(r => r[0]), descending: o === -1 });
    }
    const idx = [...Array(across ? m[0].length : m.length).keys()];
    idx.sort((p, q) => {
      for (const k of keys) {
        const c = order(k.values[p], k.values[q], k.descending);
        if (c) return c;
      }
      return p - q;
    });
    return across ? m.map(row => idx.map(i => row[i])) : idx.map(i => m[i]);
  },
});

/** FILTER(array, include, [if_empty]): the rows (include is a column) or columns (include is a row) where include is TRUE or a non-zero number. */
define('FILTER', {
  min: 2, max: 3, category: 'lookup', args: 'array, include, [if_empty]',
  fn: ([array, include, ifEmpty], ctx) => {
    if (isErr(array)) return array;
    if (isErr(include)) return include;
    const m = matrix(array, ctx);
    const f = matrix(include, ctx);
    let byRows;
    if (f[0].length === 1 && f.length === m.length) byRows = true;
    else if (f.length === 1 && f[0].length === m[0].length) byRows = false;
    else return VALUE;
    const keep = [];
    for (const v of byRows ? f.map(r => r[0]) : f[0]) {
      if (isErr(v)) return v;
      if (typeof v === 'string' && v !== '') return VALUE;
      keep.push(v === true || (typeof v === 'number' && v !== 0));
    }
    const out = byRows ? m.filter((_, i) => keep[i]) : m.map(row => row.filter((_, i) => keep[i]));
    if (byRows ? out.length : out[0].length) return out;
    return omitted(ifEmpty) ? CALC : deref(ifEmpty, ctx);
  },
});

// ---- new arrays ------------------------------------------------------------------------------------------------------------------

define('SEQUENCE', {
  min: 1, max: 4, category: 'math', args: 'rows, [columns], [start], [step]',
  fn: ([rows, cols, start, step], ctx) => {
    const a = [optNumber(rows, ctx, 1), optNumber(cols, ctx, 1), optNumber(start, ctx, 1), optNumber(step, ctx, 1)];
    const e = a.find(isErr);
    if (e) return e;
    const r = Math.trunc(a[0]);
    const c = Math.trunc(a[1]);
    return badSize(r, c) || makeMatrix(r, c, (i, j) => a[2] + (i * c + j) * a[3]);
  },
});

define('RANDARRAY', {
  min: 0, max: 5, category: 'math', volatile: true, args: '[rows], [columns], [min], [max], [whole_number]',
  fn: ([rows, cols, low, high, whole], ctx) => {
    const a = [optNumber(rows, ctx, 1), optNumber(cols, ctx, 1), optNumber(low, ctx, 0), optNumber(high, ctx, 1)];
    const e = a.find(isErr);
    if (e) return e;
    const w = optBool(whole, ctx, false);
    if (isErr(w)) return w;
    const r = Math.trunc(a[0]);
    const c = Math.trunc(a[1]);
    const bad = badSize(r, c);
    if (bad) return bad;
    if (a[2] > a[3]) return VALUE;
    if (!w) return makeMatrix(r, c, () => a[2] + ctx.random() * (a[3] - a[2]));
    const lo = Math.ceil(a[2]);
    const hi = Math.floor(a[3]);
    return lo > hi ? VALUE : makeMatrix(r, c, () => lo + Math.floor(ctx.random() * (hi - lo + 1)));
  },
});

// ---- stacking and shaping ---------------------------------------------------------------------------------------------------------

define('VSTACK', {
  min: 1, max: 255, category: 'lookup', args: 'array1, [array2], …',
  fn: (args, ctx) => {
    const ms = args.map(a => matrix(a, ctx));
    const width = ms.reduce((w, m) => Math.max(w, m[0].length), 0);
    return ms.flatMap(m => m.map(row => padded(row, width, NA)));
  },
});

define('HSTACK', {
  min: 1, max: 255, category: 'lookup', args: 'array1, [array2], …',
  fn: (args, ctx) => {
    const ms = args.map(a => matrix(a, ctx));
    const height = ms.reduce((h, m) => Math.max(h, m.length), 0);
    return Array.from({ length: height }, (_, r) => ms.flatMap(m => (r < m.length ? m[r] : new Array(m[0].length).fill(NA))));
  },
});

/** n > 0: the first n, n < 0: the last |n| (take); or everything but those (drop). */
const part = (list, n, take) => (take ? (n >= 0 ? list.slice(0, n) : list.slice(Math.max(0, list.length + n))) : n >= 0 ? list.slice(n) : list.slice(0, Math.max(0, list.length + n)));

/** TAKE/DROP(array, rows, [columns]). */
function cut(take) {
  return ([array, rows, cols], ctx) => {
    if (isErr(array)) return array;
    const r = omitted(rows) ? null : optNumber(rows, ctx, 0);
    if (isErr(r)) return r;
    const c = omitted(cols) ? null : optNumber(cols, ctx, 0);
    if (isErr(c)) return c;
    if (take && (r === 0 || c === 0)) return CALC;
    let out = matrix(array, ctx);
    if (r !== null) out = part(out, Math.trunc(r), take);
    if (c !== null) out = out.map(row => part(row, Math.trunc(c), take));
    return out.length && out[0].length ? out : CALC;
  };
}

define('TAKE', { min: 2, max: 3, category: 'lookup', args: 'array, rows, [columns]', fn: cut(true) });
define('DROP', { min: 2, max: 3, category: 'lookup', args: 'array, rows, [columns]', fn: cut(false) });

/** CHOOSECOLS/CHOOSEROWS(array, n1, …): the columns (rows) at the positions, negative ones counted from the end. */
function choose(columns) {
  return ([array, ...nums], ctx) => {
    if (isErr(array)) return array;
    const m = matrix(array, ctx);
    const n = columns ? m[0].length : m.length;
    const picks = [];
    for (const arg of nums) {
      for (const v of flat(arg, ctx)) {
        const x = toNumber(v);
        if (isErr(x)) return x;
        const k = Math.trunc(x);
        if (k === 0 || Math.abs(k) > n) return VALUE;
        picks.push(k > 0 ? k - 1 : n + k);
      }
    }
    return columns ? m.map(row => picks.map(i => row[i])) : picks.map(i => m[i]);
  };
}

define('CHOOSECOLS', { min: 2, max: 255, category: 'lookup', args: 'array, col_num1, [col_num2], …', fn: choose(true) });
define('CHOOSEROWS', { min: 2, max: 255, category: 'lookup', args: 'array, row_num1, [row_num2], …', fn: choose(false) });

/** TOCOL/TOROW(array, [ignore], [scan_by_column]): ignore 1 blanks, 2 errors, 3 both. */
function line(column) {
  return ([array, ignore, byCol], ctx) => {
    if (isErr(array)) return array;
    const ig = optNumber(ignore, ctx, 0);
    if (isErr(ig)) return ig;
    const i = Math.trunc(ig);
    if (i < 0 || i > 3) return VALUE;
    const bc = optBool(byCol, ctx, false);
    if (isErr(bc)) return bc;
    const m = matrix(array, ctx);
    const list = (bc ? transpose(m) : m).flat().filter(v => !((i & 1) && v === null) && !((i & 2) && isErr(v)));
    if (!list.length) return CALC;
    return column ? list.map(v => [v]) : [list];
  };
}

define('TOCOL', { min: 1, max: 3, category: 'lookup', args: 'array, [ignore], [scan_by_column]', fn: line(true) });
define('TOROW', { min: 1, max: 3, category: 'lookup', args: 'array, [ignore], [scan_by_column]', fn: line(false) });

define('TRANSPOSE', { min: 1, max: 1, category: 'lookup', args: 'array', fn: ([a], ctx) => (isErr(a) ? a : transpose(matrix(a, ctx))) });

/** WRAPROWS/WRAPCOLS(vector, wrap_count, [pad_with]): a row or column cut into rows (columns) of wrap_count values. */
function wrap(rows) {
  return ([vector, count, pad], ctx) => {
    if (isErr(vector)) return vector;
    const m = matrix(vector, ctx);
    if (m.length > 1 && m[0].length > 1) return VALUE;
    const n = optNumber(count, ctx, 0);
    if (isErr(n)) return n;
    const k = Math.trunc(n);
    if (k < 1) return NUM;
    const list = m.flat();
    const filler = omitted(pad) ? NA : first(deref(pad, ctx));
    const out = makeMatrix(Math.ceil(list.length / k), k, (r, c) => (r * k + c < list.length ? list[r * k + c] : filler));
    return rows ? out : transpose(out);
  };
}

define('WRAPROWS', { min: 2, max: 3, category: 'lookup', args: 'vector, wrap_count, [pad_with]', fn: wrap(true) });
define('WRAPCOLS', { min: 2, max: 3, category: 'lookup', args: 'vector, wrap_count, [pad_with]', fn: wrap(false) });

define('EXPAND', {
  min: 2, max: 4, category: 'lookup', args: 'array, rows, [columns], [pad_with]',
  fn: ([array, rows, cols, pad], ctx) => {
    if (isErr(array)) return array;
    const m = matrix(array, ctx);
    const r = optNumber(rows, ctx, m.length);
    if (isErr(r)) return r;
    const c = optNumber(cols, ctx, m[0].length);
    if (isErr(c)) return c;
    const height = Math.trunc(r);
    const width = Math.trunc(c);
    if (height < m.length || width < m[0].length) return VALUE;
    const bad = badSize(height, width);
    if (bad) return bad;
    const filler = omitted(pad) ? NA : first(deref(pad, ctx));
    return makeMatrix(height, width, (i, j) => (i < m.length && j < m[0].length ? m[i][j] : filler));
  },
});
