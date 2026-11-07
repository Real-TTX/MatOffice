/* MatOffice Sheets – lookup & reference functions (pure, no DOM). INDEX, OFFSET, INDIRECT, CHOOSE and XLOOKUP give references where they get
   one, so =SUM(A1:INDEX(A:A, 3)) and =ROW(INDEX(A1:A9, 3)) work. Exact matches compare text without case and know the wildcards * ? ~
   (MATCH 0, VLOOKUP FALSE, XLOOKUP 2); approximate matches search sorted data like Excel's binary search. */
import { define, scalar, deref, matrix, makeMatrix, pick, Errors, Ref } from '../lib.js';
import { FormulaError, toNumber, toText, toBool, compare, compareText, wildcard } from '../values.js';
import { MAX_ROWS, MAX_COLS, colName, sheetPrefix } from '../address.js';
import { lift, optNumber } from './stat.js';

const isErr = v => v instanceof FormulaError;
const { NA, REF, VALUE } = Errors;
const first = v => (Array.isArray(v) ? v[0][0] : v);
const kind = v => (typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0);
const round15 = n => (n === 0 ? 0 : Number(n.toPrecision(15)));

// ---- matching ----------------------------------------------------------------------------------------------------------------------

/** The test of an exact match: the same type, text without case (and with wildcards when wild). A blank lookup value is 0. */
function exactly(x, wild) {
  const v = x ?? 0;
  if (typeof v === 'string') {
    const re = wild && /[*?~]/.test(v) ? wildcard(v) : null;
    return y => typeof y === 'string' && (re ? re.test(y) : compareText(y, v) === 0);
  }
  if (typeof v === 'number') {
    const r = round15(v);
    return y => typeof y === 'number' && (y === v || round15(y) === r);
  }
  if (typeof v === 'boolean') return y => y === v;
  return () => false;
}

/**
 * Excel's binary search over sorted values: among the values of x's type (others are skipped), the last one ≤ x (descending: ≥ x).
 * Returns { idx (their positions), k (the index into idx, -1 when there is none) }.
 */
function binary(list, x, descending) {
  const v = x ?? 0;
  const t = kind(v);
  const idx = [];
  for (let i = 0; i < list.length; i++) if (kind(list[i]) === t) idx.push(i);
  let lo = 0;
  let hi = idx.length - 1;
  let k = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const c = compare(list[idx[mid]], v);
    if (descending ? c >= 0 : c <= 0) { k = mid; lo = mid + 1; } else hi = mid - 1;
  }
  return { idx, k };
}

/** The position of the approximate match (MATCH 1/-1, VLOOKUP TRUE, LOOKUP), or -1. */
const approximate = (list, x, descending = false) => { const { idx, k } = binary(list, x, descending); return k < 0 ? -1 : idx[k]; };

/**
 * The position of x for XLOOKUP/XMATCH, or -1. match_mode: 0 exact, -1 exact or the next smaller, 1 exact or the next larger, 2 wildcards;
 * search_mode: 1 first to last, -1 last to first, 2 binary on ascending data, -2 binary on descending data.
 */
function xfind(list, x, mode, search) {
  if (search === 2 || search === -2) {
    const { idx, k } = binary(list, x, search === -2);
    if (k >= 0 && exactly(x, false)(list[idx[k]])) return idx[k];
    if (mode === 0 || mode === 2) return -1;
    const next = (mode === 1) === (search === 2) ? k + 1 : k;
    return next >= 0 && next < idx.length ? idx[next] : -1;
  }
  const eq = exactly(x, mode === 2);
  const v = x ?? 0;
  const n = list.length;
  let best = -1;
  for (let j = 0; j < n; j++) {
    const i = search === -1 ? n - 1 - j : j;
    const y = list[i];
    if (eq(y)) return i;
    if ((mode === 1 || mode === -1) && kind(y) === kind(v)) {
      const c = compare(y, v);
      if (c === mode && (best < 0 || compare(y, list[best]) === -mode)) best = i;
    }
  }
  return best;
}

/** A lookup vector (one row or one column): { list (its values), horizontal, length (of the whole reference) }; null for a 2-D range. */
function vectorOf(v, ctx) {
  const m = matrix(v, ctx);
  const rows = v instanceof Ref ? v.height : m.length;
  const cols = v instanceof Ref ? v.width : m[0].length;
  if (rows > 1 && cols > 1) return null;
  const horizontal = rows === 1 && cols > 1;
  return { list: horizontal ? m[0] : m.map(r => r[0]), horizontal, length: horizontal ? cols : rows };
}

/** match_mode and search_mode of XLOOKUP/XMATCH, or #VALUE!. */
function modes(mode, search, ctx) {
  const m = optNumber(mode, ctx, 0);
  if (isErr(m)) return m;
  const s = optNumber(search, ctx, 1);
  if (isErr(s)) return s;
  return [0, -1, 1, 2].includes(m) && [1, -1, 2, -2].includes(s) ? [m, s] : VALUE;
}

// ---- the lookups ---------------------------------------------------------------------------------------------------------------

define('MATCH', {
  min: 2, max: 3, category: 'lookup', args: 'lookup_value, lookup_array, [match_type]',
  fn: ([x, array, type], ctx) => {
    if (isErr(array)) return array;
    const vec = vectorOf(array, ctx);
    if (!vec) return NA;
    const t = type === undefined ? 1 : toNumber(first(deref(type, ctx)));
    if (isErr(t)) return t;
    return lift([x], ctx, ([v]) => {
      if (isErr(v)) return v;
      const i = t === 0 ? vec.list.findIndex(exactly(v, true)) : approximate(vec.list, v, t < 0);
      return i < 0 ? NA : i + 1;
    });
  },
});

define('XMATCH', {
  min: 2, max: 4, category: 'lookup', args: 'lookup_value, lookup_array, [match_mode], [search_mode]',
  fn: ([x, array, mode, search], ctx) => {
    if (isErr(array)) return array;
    const vec = vectorOf(array, ctx);
    if (!vec) return VALUE;
    const ms = modes(mode, search, ctx);
    if (isErr(ms)) return ms;
    return lift([x], ctx, ([v]) => {
      if (isErr(v)) return v;
      const i = xfind(vec.list, v, ms[0], ms[1]);
      return i < 0 ? NA : i + 1;
    });
  },
});

/** VLOOKUP/HLOOKUP: the first column (row) is searched, the value comes from column (row) index of the match. */
function tableLookup(horizontal) {
  return ([x, table, index, approx], ctx) => {
    if (isErr(table)) return table;
    const m = matrix(table, ctx);
    const size = table instanceof Ref ? (horizontal ? table.height : table.width) : horizontal ? m.length : m[0].length;
    let exact = false;
    if (approx !== undefined) {
      const b = toBool(first(deref(approx, ctx)));
      if (isErr(b)) return b;
      exact = !b;
    }
    const keys = horizontal ? m[0] : m.map(r => r[0]);
    return lift([x, index], ctx, ([v, n]) => {
      if (isErr(v)) return v;
      const k = toNumber(n);
      if (isErr(k)) return k;
      const i = Math.trunc(k);
      if (i < 1) return VALUE;
      if (i > size) return REF;
      const pos = exact ? keys.findIndex(exactly(v, true)) : approximate(keys, v);
      if (pos < 0) return NA;
      return (horizontal ? m[i - 1]?.[pos] : m[pos][i - 1]) ?? null;
    });
  };
}

define('VLOOKUP', { min: 3, max: 4, category: 'lookup', args: 'lookup_value, table_array, col_index_num, [range_lookup]', fn: tableLookup(false) });
define('HLOOKUP', { min: 3, max: 4, category: 'lookup', args: 'lookup_value, table_array, row_index_num, [range_lookup]', fn: tableLookup(true) });

/**
 * XLOOKUP: the row (column) of return_array where lookup_array matches – a reference when return_array is one. An array of lookup values gives
 * one result each; a column of them with a several columns wide return array gives the matching rows.
 */
define('XLOOKUP', {
  min: 3, max: 6, category: 'lookup', args: 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]',
  fn: ([x, lookupArray, returnArray, notFound, mode, search], ctx) => {
    if (isErr(lookupArray)) return lookupArray;
    if (isErr(returnArray)) return returnArray;
    const vec = vectorOf(lookupArray, ctx);
    if (!vec) return VALUE;
    const ms = modes(mode, search, ctx);
    if (isErr(ms)) return ms;
    const ref = returnArray instanceof Ref ? returnArray : null;
    const rm = ref ? null : matrix(returnArray, ctx);
    const rows = ref ? ref.height : rm.length;
    const cols = ref ? ref.width : rm[0].length;
    if ((vec.horizontal ? cols : rows) !== vec.length) return VALUE;
    const result = i => {
      if (ref) return vec.horizontal ? new Ref(ref.sheet, ref.r1, ref.c1 + i, ref.r2, ref.c1 + i) : new Ref(ref.sheet, ref.r1 + i, ref.c1, ref.r1 + i, ref.c2);
      return vec.horizontal ? rm.map(r => [r[i]]) : [rm[i]];
    };
    const one = v => {
      if (isErr(v)) return v;
      const i = xfind(vec.list, v, ms[0], ms[1]);
      return i >= 0 ? result(i) : notFound === undefined || notFound === null ? NA : notFound;
    };
    const xv = deref(x, ctx);
    if (!Array.isArray(xv)) return one(xv);
    const results = xv.map(row => row.map(v => { const r = one(v); return r instanceof Ref || Array.isArray(r) ? matrix(r, ctx) : r; }));
    const cell = (r, i, j) => (Array.isArray(r) ? r[i]?.[j] ?? null : r);
    if (xv[0].length === 1 && !vec.horizontal) return makeMatrix(xv.length, cols, (i, j) => cell(results[i][0], 0, j));
    if (xv.length === 1 && vec.horizontal) return makeMatrix(rows, xv[0].length, (i, j) => cell(results[0][j], i, 0));
    return results.map(row => row.map(r => cell(r, 0, 0)));
  },
});

/** LOOKUP: the vector form (lookup_vector, result_vector) or the array form (searches the first row/column, returns from the last). */
define('LOOKUP', {
  min: 2, max: 3, category: 'lookup', args: 'lookup_value, lookup_vector, [result_vector]',
  fn: ([x, vector, resultVector], ctx) => {
    if (isErr(vector)) return vector;
    if (isErr(resultVector)) return resultVector;
    const m = matrix(vector, ctx);
    const rows = vector instanceof Ref ? vector.height : m.length;
    const cols = vector instanceof Ref ? vector.width : m[0].length;
    const wide = cols > rows;
    const keys = wide ? m[0] : m.map(r => r[0]);
    let value;
    if (resultVector === undefined) {
      value = wide ? i => m[rows - 1]?.[i] : i => m[i][cols - 1];
    } else if (resultVector instanceof Ref) {
      // a result vector of another size goes on in its direction (like Excel)
      const across = resultVector.height === 1 && resultVector.width > 1;
      value = i => (across ? ctx.value(resultVector.sheet, resultVector.r1, resultVector.c1 + i) : ctx.value(resultVector.sheet, resultVector.r1 + i, resultVector.c1));
    } else {
      const r = matrix(resultVector, ctx);
      const list = r.length === 1 ? r[0] : r.map(row => row[0]);
      value = i => (i < list.length ? list[i] : NA);
    }
    return lift([x], ctx, ([v]) => {
      if (isErr(v)) return v;
      const i = approximate(keys, v);
      return i < 0 ? NA : value(i) ?? null;
    });
  },
});

// ---- references ------------------------------------------------------------------------------------------------------------------

/**
 * INDEX(array, row_num, [column_num], [area_num]): a reference into a reference (row 0 / column 0 = the whole column / row), else the value or
 * the part of the array. With one row, the single number is the column. Arrays of positions give the values at each position.
 */
define('INDEX', {
  min: 2, max: 4, category: 'lookup', args: 'array, row_num, [column_num], [area_num]',
  fn: ([array, rowNum, colNum, areaNum], ctx) => {
    if (isErr(array)) return array;
    if (areaNum !== undefined && areaNum !== null) {
      const a = toNumber(first(deref(areaNum, ctx)));
      if (isErr(a)) return a;
      if (Math.trunc(a) !== 1) return REF;
    }
    const ref = array instanceof Ref ? array : null;
    const m = ref ? null : matrix(array, ctx);
    const rows = ref ? ref.height : m.length;
    const cols = ref ? ref.width : m[0].length;
    const at = (rv, cv) => {
      const rn = toNumber(rv);
      if (isErr(rn)) return rn;
      const cn = cv === undefined ? 0 : toNumber(cv);
      if (isErr(cn)) return cn;
      let r = Math.trunc(rn);
      let c = Math.trunc(cn);
      if (cv === undefined && rows === 1) [r, c] = [0, r];
      if (r < 0 || c < 0) return VALUE;
      if (r > rows || c > cols) return REF;
      if (ref) {
        return new Ref(ref.sheet, r ? ref.r1 + r - 1 : ref.r1, c ? ref.c1 + c - 1 : ref.c1, r ? ref.r1 + r - 1 : ref.r2, c ? ref.c1 + c - 1 : ref.c2);
      }
      const part = (r ? [m[r - 1]] : m).map(row => (c ? [row[c - 1]] : row));
      return part.length === 1 && part[0].length === 1 ? part[0][0] : part;
    };
    const rv = deref(rowNum, ctx);
    const cv = colNum === undefined ? undefined : deref(colNum, ctx);
    if (!Array.isArray(rv) && !Array.isArray(cv)) return at(rv, cv);
    const size = [rv, cv].filter(Array.isArray);
    const height = Math.max(...size.map(a => a.length));
    const width = Math.max(...size.map(a => a[0].length));
    return makeMatrix(height, width, (i, j) => {
      const v = deref(at(Array.isArray(rv) ? pick(rv, i, j) : rv, Array.isArray(cv) ? pick(cv, i, j) : cv), ctx);
      return first(v);
    });
  },
});

/** CHOOSE(index_num, value1, …): lazy; an array of indexes chooses element by element (=CHOOSE({1,2}, B1:B9, A1:A9) is a two-column array). */
define('CHOOSE', {
  min: 2, max: 255, category: 'lookup', lazy: true, args: 'index_num, value1, [value2], …',
  fn: ([index, ...values], ctx) => {
    const i = deref(index(), ctx);
    const which = v => {
      const n = isErr(v) ? v : toNumber(v);
      if (isErr(n)) return n;
      const k = Math.trunc(n);
      return k >= 1 && k <= values.length ? k : VALUE;
    };
    if (!Array.isArray(i)) {
      const k = which(i);
      if (isErr(k)) return k;
      const v = values[k - 1]();
      return v === null ? 0 : v;
    }
    const ks = i.map(row => row.map(which));
    const chosen = new Map();
    for (const k of ks.flat()) if (!isErr(k) && !chosen.has(k)) chosen.set(k, deref(values[k - 1](), ctx));
    let height = i.length;
    let width = i[0].length;
    for (const v of chosen.values()) if (Array.isArray(v)) { height = Math.max(height, v.length); width = Math.max(width, v[0].length); }
    return makeMatrix(height, width, (r, c) => {
      const k = pick(ks, r, c);
      if (isErr(k)) return k;
      const v = chosen.get(k);
      return Array.isArray(v) ? pick(v, r, c) : v;
    });
  },
});

/** How many rows (columns) of a reference ROW/COLUMN list: all of them, a huge one (A:A) only up to the used area. */
function lines(ref, ctx, rows) {
  const n = rows ? ref.height : ref.width;
  if (n <= 65536) return n;
  const m = ctx.values(rows ? new Ref(ref.sheet, ref.r1, ref.c1, ref.r2, ref.c1) : new Ref(ref.sheet, ref.r1, ref.c1, ref.r1, ref.c2));
  return rows ? m.length : m[0].length;
}

define('ROW', {
  min: 0, max: 1, category: 'lookup', args: '[reference]',
  fn: ([ref], ctx) => {
    if (ref === undefined || ref === null) return ctx.row + 1;
    if (isErr(ref)) return ref;
    if (!(ref instanceof Ref)) return VALUE;
    return ref.height === 1 ? ref.r1 + 1 : makeMatrix(lines(ref, ctx, true), 1, r => ref.r1 + r + 1);
  },
});

define('COLUMN', {
  min: 0, max: 1, category: 'lookup', args: '[reference]',
  fn: ([ref], ctx) => {
    if (ref === undefined || ref === null) return ctx.col + 1;
    if (isErr(ref)) return ref;
    if (!(ref instanceof Ref)) return VALUE;
    return ref.width === 1 ? ref.c1 + 1 : makeMatrix(1, lines(ref, ctx, false), (r, c) => ref.c1 + c + 1);
  },
});

const sizeOf = (name, rows) => define(name, {
  min: 1, max: 1, category: 'lookup', args: 'array',
  fn: ([a]) => (isErr(a) ? a : a instanceof Ref ? (rows ? a.height : a.width) : Array.isArray(a) ? (rows ? a.length : a[0].length) : 1),
});
sizeOf('ROWS', true);
sizeOf('COLUMNS', false);

define('AREAS', { min: 1, max: 1, category: 'lookup', args: 'reference', fn: ([a]) => (isErr(a) ? a : a instanceof Ref ? 1 : VALUE) });

/** OFFSET(reference, rows, cols, [height], [width]): a reference moved and resized; a negative height/width reaches up/left. */
define('OFFSET', {
  min: 3, max: 5, category: 'lookup', volatile: true, args: 'reference, rows, cols, [height], [width]',
  fn: ([ref, rows, cols, height, width], ctx) => {
    if (isErr(ref)) return ref;
    if (!(ref instanceof Ref)) return VALUE;
    const nums = [optNumber(rows, ctx, 0), optNumber(cols, ctx, 0), optNumber(height, ctx, ref.height), optNumber(width, ctx, ref.width)];
    const e = nums.find(isErr);
    if (e) return e;
    const [dr, dc, h, w] = nums.map(Math.trunc);
    if (h === 0 || w === 0) return REF;
    const r = ref.r1 + dr;
    const c = ref.c1 + dc;
    const [r1, r2] = h > 0 ? [r, r + h - 1] : [r + h + 1, r];
    const [c1, c2] = w > 0 ? [c, c + w - 1] : [c + w + 1, c];
    return r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS ? REF : new Ref(ref.sheet, r1, c1, r2, c2);
  },
});

const R1C1 = /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i;

/** A reference in R1C1 notation ("R2C3", "R[-1]C", "Sheet2!R1C1:R2C2"); relative parts count from the formula's cell. */
function r1c1(text, ctx) {
  let sheet = ctx.sheet;
  let rest = text.trim();
  const bang = rest.lastIndexOf('!');
  if (bang >= 0) {
    let name = rest.slice(0, bang);
    if (/^'.*'$/.test(name)) name = name.slice(1, -1).replace(/''/g, "'");
    sheet = ctx.sheetId(name);
    if (sheet === null || sheet === undefined) return REF;
    rest = rest.slice(bang + 1);
  }
  const cells = rest.split(':').map(p => {
    const m = R1C1.exec(p.trim());
    if (!m) return null;
    const at = (s, base) => (s === undefined ? base : s[0] === '[' ? base + Number(s.slice(1, -1)) : Number(s) - 1);
    const r = at(m[1], ctx.row);
    const c = at(m[2], ctx.col);
    return r >= 0 && r < MAX_ROWS && c >= 0 && c < MAX_COLS ? { r, c } : null;
  });
  if (cells.length > 2 || cells.some(x => !x)) return REF;
  const [a, b = a] = cells;
  return new Ref(sheet, a.r, a.c, b.r, b.c);
}

define('INDIRECT', {
  min: 1, max: 2, category: 'lookup', volatile: true, args: 'ref_text, [a1]',
  fn: ([text, a1], ctx) => {
    const t = toText(first(deref(text, ctx)));
    if (isErr(t)) return t;
    let style = true;
    if (a1 !== undefined) {
      const b = toBool(first(deref(a1, ctx)));
      if (isErr(b)) return b;
      style = b;
    }
    return style ? ctx.parseRef(t) : r1c1(t, ctx);
  },
});

/** ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]): abs_num 1 $A$1, 2 A$1, 3 $A1, 4 A1; a1 FALSE writes R1C1. */
define('ADDRESS', {
  min: 2, max: 5, category: 'lookup', args: 'row_num, column_num, [abs_num], [a1], [sheet_text]',
  fn: scalar(['number', 'number', 'any?', 'any?', 'any?'], (r, c, abs, a1, sheet) => {
    const mode = abs === undefined || abs === null ? 1 : toNumber(abs);
    if (isErr(mode)) return mode;
    const row = Math.trunc(r);
    const col = Math.trunc(c);
    const m = Math.trunc(mode);
    if (row < 1 || row > MAX_ROWS || col < 1 || col > MAX_COLS || m < 1 || m > 4) return VALUE;
    const style = a1 === undefined || a1 === null ? true : toBool(a1);
    if (isErr(style)) return style;
    const absRow = m === 1 || m === 2;
    const absCol = m === 1 || m === 3;
    const text = style
      ? (absCol ? '$' : '') + colName(col - 1) + (absRow ? '$' : '') + row
      : 'R' + (absRow ? row : `[${row}]`) + 'C' + (absCol ? col : `[${col}]`);
    if (sheet === undefined || sheet === null) return text;
    const s = toText(sheet);
    if (isErr(s)) return s;
    return (s === '' ? '!' : sheetPrefix(s)) + text;
  }),
});

define('HYPERLINK', {
  min: 1, max: 2, category: 'lookup', args: 'link_location, [friendly_name]',
  fn: scalar(['any', 'any?'], (link, name) => (name === undefined || name === null ? link : name)),
});

define('FORMULATEXT', {
  min: 1, max: 1, category: 'lookup', args: 'reference',
  fn: ([ref], ctx) => {
    if (isErr(ref)) return ref;
    if (!(ref instanceof Ref)) return NA;
    const f = ctx.formulaAt(ref.sheet, ref.r1, ref.c1);
    return f ? '=' + f : NA;
  },
});
