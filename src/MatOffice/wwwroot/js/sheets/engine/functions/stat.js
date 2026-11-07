/* MatOffice Sheets – statistical functions (pure, no DOM). Also what the other categories share: arguments as grids, lifting over arrays,
   the criteria of the *IF/*IFS functions (SUMIF/SUMIFS live in math.js) and the statistics of a list of numbers (SUBTOTAL/AGGREGATE in math.js). */
import { define, scalar, numbers, deref, matrix, each, makeMatrix, pick, Errors, Ref } from '../lib.js';
import { FormulaError, criteria, toNumber, toBool, parseNumber } from '../values.js';
import { MAX_ROWS, MAX_COLS } from '../address.js';

const isErr = v => v instanceof FormulaError;
const { NUM, DIV0, NA, VALUE } = Errors;

// ---- arguments ----------------------------------------------------------------------------------------------------------------

/**
 * An argument as a grid of its full size: { rows, cols, m, at(r, c) }. A huge reference (A:A) keeps its size although its values (m) are cut
 * at the used area; at() is blank beyond.
 */
export function grid(v, ctx) {
  const m = matrix(v, ctx);
  const ref = v instanceof Ref;
  return { rows: ref ? v.height : m.length, cols: ref ? v.width : m[0].length, m, at: (r, c) => m[r]?.[c] ?? null };
}

/** compute(values) for scalar arguments, or element by element when one of them is an array or a range (=RANK(A1:A3, A1:A9) spills). */
export function lift(args, ctx, compute) {
  const vals = args.map(a => deref(a, ctx));
  if (!vals.some(Array.isArray)) return compute(vals);
  let rows = 1;
  let cols = 1;
  for (const v of vals) if (Array.isArray(v)) { rows = Math.max(rows, v.length); cols = Math.max(cols, v[0].length); }
  return makeMatrix(rows, cols, (r, c) => compute(vals.map(v => (Array.isArray(v) ? pick(v, r, c) : v))));
}

/** A number argument that may be omitted (undefined or an empty argument): its default, else the number (or an error). */
export function optNumber(v, ctx, def) {
  if (v === undefined || v === null) return def;
  const x = deref(v, ctx);
  return toNumber(Array.isArray(x) ? x[0][0] : x);
}

// ---- statistics of a list of numbers ------------------------------------------------------------------------------------------------

export const sorted = list => [...list].sort((a, b) => a - b);
export const sum = list => { let s = 0; for (const x of list) s += x; return s; };
export const mean = list => (list.length ? sum(list) / list.length : DIV0);
export const maxOf = list => { let m = -Infinity; for (const x of list) if (x > m) m = x; return list.length ? m : 0; };
export const minOf = list => { let m = Infinity; for (const x of list) if (x < m) m = x; return list.length ? m : 0; };
export const product = list => { if (!list.length) return 0; let p = 1; for (const x of list) p *= x; return p; };
const devsq = list => { const m = sum(list) / list.length; let s = 0; for (const x of list) s += (x - m) * (x - m); return s; };
export const variance = (list, sample) => (list.length < (sample ? 2 : 1) ? DIV0 : devsq(list) / (list.length - (sample ? 1 : 0)));
export const stdev = (list, sample) => { const v = variance(list, sample); return isErr(v) ? v : Math.sqrt(v); };

export function median(list) {
  if (!list.length) return NUM;
  const s = sorted(list);
  const h = s.length >> 1;
  return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
}

/** The most frequent numbers in the order they first appear; #N/A when no number repeats. */
export function modes(list) {
  const counts = new Map();
  for (const x of list) counts.set(x, (counts.get(x) || 0) + 1);
  let best = 1;
  for (const c of counts.values()) if (c > best) best = c;
  return best < 2 ? NA : [...counts].filter(([, c]) => c === best).map(([x]) => x);
}
export const mode = list => { const m = modes(list); return isErr(m) ? m : m[0]; };

/** The k-th largest (or smallest) of sorted numbers; k is truncated, #NUM! outside 1 … n. */
export function kth(s, k, largest) {
  const i = Math.trunc(k);
  if (i < 1 || i > s.length) return NUM;
  return largest ? s[s.length - i] : s[i - 1];
}

/** The percentile p of sorted numbers, interpolated like PERCENTILE.INC (exclusive: PERCENTILE.EXC). */
export function percentile(s, p, exclusive) {
  const n = s.length;
  if (!n || p < 0 || p > 1) return NUM;
  const h = Number(((exclusive ? n + 1 : n - 1) * p).toPrecision(15)) - (exclusive ? 1 : 0);
  if (h < 0 || h > n - 1) return NUM;
  const i = Math.floor(h);
  return i + 1 < n ? s[i] + (h - i) * (s[i + 1] - s[i]) : s[i];
}

/** QUARTILE(.INC/.EXC): quart 0 … 4 (exclusive 1 … 3) as a percentile. */
export function quartile(s, q, exclusive) {
  const i = Math.trunc(q);
  return (exclusive ? i < 1 || i > 3 : i < 0 || i > 4) ? NUM : percentile(s, i / 4, exclusive);
}

// ---- the criteria of COUNTIF, SUMIFS & co --------------------------------------------------------------------------------------

/**
 * Where all criteria hold: ranges (grids of one size), one criterion (a scalar) each – a blank criterion cell counts as 0, like Excel. hit(r, c)
 * is called for every match in the used area; returns how many blank cells beyond it match as well (only COUNTIFS counts them).
 */
function where(ranges, crits, hit) {
  const tests = crits.map(c => criteria(c ?? 0));
  const { rows, cols } = ranges[0];
  let dr = 0;
  let dc = 0;
  for (const g of ranges) { dr = Math.max(dr, g.m.length); dc = Math.max(dc, g.m[0].length); }
  dr = Math.min(dr, rows);
  dc = Math.min(dc, cols);
  for (let r = 0; r < dr; r++) {
    for (let c = 0; c < dc; c++) {
      let ok = true;
      for (let i = 0; ok && i < ranges.length; i++) ok = tests[i](ranges[i].at(r, c));
      if (ok) hit(r, c);
    }
  }
  return rows * cols > dr * dc && tests.every(t => t(null)) ? rows * cols - dr * dc : 0;
}

/**
 * COUNTIFS (target null) and the aggregates of the target cells where the criteria hold (SUMIFS, AVERAGEIFS, MAXIFS …): pairs is
 * [range1, criteria1, range2, criteria2 …]; every range must have the size of the target (else #VALUE!). reduce(numbers) gives the result.
 */
export function ifs(target, pairs, ctx, reduce) {
  if (pairs.length % 2) return VALUE;
  const ranges = [];
  const crits = [];
  for (let i = 0; i < pairs.length; i += 2) {
    if (isErr(pairs[i])) return pairs[i];
    const g = grid(pairs[i], ctx);
    const size = target || ranges[0];
    if (size && (g.rows !== size.rows || g.cols !== size.cols)) return VALUE;
    ranges.push(g);
    crits.push(pairs[i + 1]);
  }
  return lift(crits, ctx, values => {
    if (!target) {
      let n = 0;
      const rest = where(ranges, values, () => { n++; });
      return n + rest;
    }
    const list = [];
    let error = null;
    where(ranges, values, (r, c) => {
      const v = target.at(r, c);
      if (isErr(v)) error ??= v;
      else if (typeof v === 'number') list.push(v);
    });
    return error || reduce(list);
  });
}

/** SUMIF/AVERAGEIF(range, criteria, [values]): values of another size start at their top-left cell and take the size of range (like Excel). */
export function ifOne(reduce) {
  return ([range, crit, values], ctx) => {
    if (isErr(range)) return range;
    if (isErr(values)) return values;
    const g = grid(range, ctx);
    let target = g;
    if (values !== undefined && values !== null) {
      let v = values;
      if (v instanceof Ref && (v.height !== g.rows || v.width !== g.cols)) {
        v = new Ref(v.sheet, v.r1, v.c1, Math.min(v.r1 + g.rows, MAX_ROWS) - 1, Math.min(v.c1 + g.cols, MAX_COLS) - 1);
      }
      target = grid(v, ctx);
    }
    return ifs(target, [range, crit], ctx, reduce);
  };
}

const ifsWith = reduce => ([values, ...pairs], ctx) => (isErr(values) ? values : ifs(grid(values, ctx), pairs, ctx, reduce));

define('COUNTIF', { min: 2, max: 2, category: 'stat', args: 'range, criteria', fn: (args, ctx) => ifs(null, args, ctx) });
define('COUNTIFS', {
  min: 2, max: 254, category: 'stat', args: 'criteria_range1, criteria1, [criteria_range2, criteria2], …',
  fn: (args, ctx) => ifs(null, args, ctx),
});
define('AVERAGEIF', { min: 2, max: 3, category: 'stat', args: 'range, criteria, [average_range]', fn: ifOne(mean) });
define('AVERAGEIFS', {
  min: 3, max: 255, category: 'stat', args: 'average_range, criteria_range1, criteria1, [criteria_range2, criteria2], …', fn: ifsWith(mean),
});
define('MAXIFS', {
  min: 3, max: 255, category: 'stat', args: 'max_range, criteria_range1, criteria1, [criteria_range2, criteria2], …', fn: ifsWith(maxOf),
});
define('MINIFS', {
  min: 3, max: 255, category: 'stat', args: 'min_range, criteria_range1, criteria1, [criteria_range2, criteria2], …', fn: ifsWith(minOf),
});

// ---- counting ---------------------------------------------------------------------------------------------------------------------

define('COUNT', {
  min: 1, max: 255, category: 'stat', args: 'value1, [value2], …',
  fn: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      each(a, ctx, (v, fromRange) => {
        if (typeof v === 'number') n++;
        else if (!fromRange && (v === null || typeof v === 'boolean' || (typeof v === 'string' && parseNumber(v) !== null))) n++;
      });
    }
    return n;
  },
});

define('COUNTA', {
  min: 1, max: 255, category: 'stat', args: 'value1, [value2], …',
  fn: (args, ctx) => {
    let n = 0;
    for (const a of args) each(a, ctx, (v, fromRange) => { if (!fromRange || v !== null) n++; });
    return n;
  },
});

/** COUNTBLANK counts every cell of the range (also beyond the used area of A:A) that is empty or "". */
define('COUNTBLANK', {
  min: 1, max: 1, category: 'stat', args: 'range',
  fn: ([range], ctx) => {
    if (isErr(range)) return range;
    const g = grid(range, ctx);
    let filled = 0;
    for (const row of g.m) for (const v of row) if (v !== null && v !== '') filled++;
    return g.rows * g.cols - filled;
  },
});

// ---- aggregates of numbers ----------------------------------------------------------------------------------------------------

/** A function of all numbers of its arguments (mode 'a': text in ranges counts 0, logicals 1/0). */
function aggregate(names, reduce, mode = '') {
  const args = mode ? 'value1, [value2], …' : 'number1, [number2], …';
  for (const name of [].concat(names)) {
    define(name, {
      min: 1, max: 255, category: 'stat', args,
      fn: (a, ctx) => { const list = numbers(a, ctx, mode); return isErr(list) ? list : reduce(list); },
    });
  }
}

aggregate('AVERAGE', mean);
aggregate('AVERAGEA', mean, 'a');
aggregate('MAX', maxOf);
aggregate('MAXA', maxOf, 'a');
aggregate('MIN', minOf);
aggregate('MINA', minOf, 'a');
aggregate('MEDIAN', median);
aggregate(['MODE', 'MODE.SNGL'], mode);
aggregate('MODE.MULT', list => { const m = modes(list); return isErr(m) ? m : m.map(x => [x]); });
aggregate(['STDEV', 'STDEV.S'], l => stdev(l, true));
aggregate(['STDEVP', 'STDEV.P'], l => stdev(l, false));
aggregate('STDEVA', l => stdev(l, true), 'a');
aggregate('STDEVPA', l => stdev(l, false), 'a');
aggregate(['VAR', 'VAR.S'], l => variance(l, true));
aggregate(['VARP', 'VAR.P'], l => variance(l, false));
aggregate('VARA', l => variance(l, true), 'a');
aggregate('VARPA', l => variance(l, false), 'a');
aggregate('DEVSQ', l => (l.length ? devsq(l) : NUM));
aggregate('AVEDEV', l => {
  if (!l.length) return NUM;
  const m = sum(l) / l.length;
  return sum(l.map(x => Math.abs(x - m))) / l.length;
});
aggregate('GEOMEAN', l => (!l.length || l.some(x => x <= 0) ? NUM : Math.exp(sum(l.map(Math.log)) / l.length)));
aggregate('HARMEAN', l => (!l.length || l.some(x => x <= 0) ? NUM : l.length / sum(l.map(x => 1 / x))));
aggregate('SKEW', l => {
  const n = l.length;
  const s = n > 2 ? stdev(l, true) : 0;
  if (!s || isErr(s)) return DIV0;
  const m = sum(l) / n;
  return (n / ((n - 1) * (n - 2))) * sum(l.map(x => ((x - m) / s) ** 3));
});
aggregate('SKEW.P', l => {
  const s = l.length ? stdev(l, false) : 0;
  if (!s || isErr(s)) return DIV0;
  const m = sum(l) / l.length;
  return sum(l.map(x => ((x - m) / s) ** 3)) / l.length;
});
aggregate('KURT', l => {
  const n = l.length;
  const s = n > 3 ? stdev(l, true) : 0;
  if (!s || isErr(s)) return DIV0;
  const m = sum(l) / n;
  const q = sum(l.map(x => ((x - m) / s) ** 4));
  return (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))) * q - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
});

/** A function of the numbers of one range and one more argument k that may be an array (=LARGE(A1:A9, {1,2,3}) spills). */
function withK(name, args, f) {
  define(name, {
    min: 2, max: 2, category: 'stat', args,
    fn: ([data, k], ctx) => {
      const list = numbers([data], ctx);
      if (isErr(list)) return list;
      const s = sorted(list);
      return lift([k], ctx, ([v]) => { const x = toNumber(v); return isErr(x) ? x : f(s, x); });
    },
  });
}

withK('LARGE', 'array, k', (s, k) => kth(s, k, true));
withK('SMALL', 'array, k', (s, k) => kth(s, k, false));
for (const name of ['PERCENTILE', 'PERCENTILE.INC']) withK(name, 'array, k', (s, p) => percentile(s, p, false));
withK('PERCENTILE.EXC', 'array, k', (s, p) => (p <= 0 || p >= 1 ? NUM : percentile(s, p, true)));
for (const name of ['QUARTILE', 'QUARTILE.INC']) withK(name, 'array, quart', (s, q) => quartile(s, q, false));
withK('QUARTILE.EXC', 'array, quart', (s, q) => quartile(s, q, true));

define('TRIMMEAN', {
  min: 2, max: 2, category: 'stat', args: 'array, percent',
  fn: ([data, pct], ctx) => {
    const list = numbers([data], ctx);
    if (isErr(list)) return list;
    const p = optNumber(pct, ctx, 0);
    if (isErr(p)) return p;
    if (p < 0 || p >= 1 || !list.length) return NUM;
    const k = Math.floor(Number((list.length * p / 2).toPrecision(15)));
    return mean(sorted(list).slice(k, list.length - k));
  },
});

/** RANK(.EQ/.AVG)(number, ref, [order]): order 0 (default) ranks descending. A number that is not in ref gives #N/A. */
function rank(average) {
  return ([x, ref, order], ctx) => {
    const list = numbers([ref], ctx);
    if (isErr(list)) return list;
    const o = optNumber(order, ctx, 0);
    if (isErr(o)) return o;
    return lift([x], ctx, ([v]) => {
      const n = toNumber(v);
      if (isErr(n)) return n;
      let before = 0;
      let same = 0;
      for (const y of list) {
        if (y === n) same++;
        else if (o ? y < n : y > n) before++;
      }
      if (!same) return NA;
      return average ? before + (same + 1) / 2 : before + 1;
    });
  };
}

for (const name of ['RANK', 'RANK.EQ']) define(name, { min: 2, max: 3, category: 'stat', args: 'number, ref, [order]', fn: rank(false) });
define('RANK.AVG', { min: 2, max: 3, category: 'stat', args: 'number, ref, [order]', fn: rank(true) });

/** PERCENTRANK(.INC/.EXC)(array, x, [significance]): the rank of x as a fraction, interpolated between the values, truncated to significance digits. */
function percentRank(exclusive) {
  return ([data, x, sig], ctx) => {
    const list = numbers([data], ctx);
    if (isErr(list)) return list;
    const digits = optNumber(sig, ctx, 3);
    if (isErr(digits)) return digits;
    const d = Math.trunc(digits);
    if (d < 1 || !list.length) return NUM;
    const s = sorted(list);
    const n = s.length;
    return lift([x], ctx, ([v]) => {
      const t = toNumber(v);
      if (isErr(t)) return t;
      if (t < s[0] || t > s[n - 1]) return NA;
      let less = 0;
      while (s[less] < t) less++;
      let r;
      if (s[less] === t) r = exclusive ? (less + 1) / (n + 1) : less ? less / (n - 1) : 0;
      else {
        const frac = (t - s[less - 1]) / (s[less] - s[less - 1]);
        r = exclusive ? (less + frac) / (n + 1) : (less - 1 + frac) / (n - 1);
      }
      const f = 10 ** d;
      return Math.floor(Number((r * f).toPrecision(15))) / f;
    });
  };
}

for (const name of ['PERCENTRANK', 'PERCENTRANK.INC']) {
  define(name, { min: 2, max: 3, category: 'stat', args: 'array, x, [significance]', fn: percentRank(false) });
}
define('PERCENTRANK.EXC', { min: 2, max: 3, category: 'stat', args: 'array, x, [significance]', fn: percentRank(true) });

/** FREQUENCY(data, bins): how many values fall into each bin (≤ bin, above the bin before) and above the last, as a vertical array. */
define('FREQUENCY', {
  min: 2, max: 2, category: 'stat', args: 'data_array, bins_array',
  fn: ([data, bins], ctx) => {
    const d = numbers([data], ctx);
    if (isErr(d)) return d;
    const b = numbers([bins], ctx);
    if (isErr(b)) return b;
    const order = b.map((v, i) => [v, i]).sort((x, y) => x[0] - y[0]);
    const counts = new Array(b.length + 1).fill(0);
    for (const v of d) {
      let lo = 0;
      let hi = order.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (v <= order[mid][0]) hi = mid; else lo = mid + 1; }
      counts[lo === order.length ? b.length : order[lo][1]]++;
    }
    return counts.map(c => [c]);
  },
});

// ---- two variables ---------------------------------------------------------------------------------------------------------------

/** The pairs of numbers of two ranges with the same count of values (CORREL, SLOPE, SUMX2MY2 …): [xs, ys]; #N/A for different counts. */
export function pairs(a, b, ctx) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const ga = grid(a, ctx);
  const gb = grid(b, ctx);
  if (ga.rows * ga.cols !== gb.rows * gb.cols) return NA;
  let x = [];
  let y = [];
  if (ga.rows === gb.rows && ga.cols === gb.cols) {
    const dr = Math.max(ga.m.length, gb.m.length);
    const dc = Math.max(ga.m[0].length, gb.m[0].length);
    for (let r = 0; r < dr; r++) for (let c = 0; c < dc; c++) { x.push(ga.at(r, c)); y.push(gb.at(r, c)); }
  } else {
    x = ga.m.flat();
    y = gb.m.flat();
  }
  const xs = [];
  const ys = [];
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    const u = x[i] ?? null;
    const v = y[i] ?? null;
    if (isErr(u)) return u;
    if (isErr(v)) return v;
    if (typeof u === 'number' && typeof v === 'number') { xs.push(u); ys.push(v); }
  }
  return [xs, ys];
}

/** The moments of paired data: count, means and the sums of squares and products of the deviations. */
function moments([xs, ys]) {
  const n = xs.length;
  const mx = sum(xs) / n;
  const my = sum(ys) / n;
  let sxx = 0;
  let syy = 0;
  let sxy = 0;
  for (let i = 0; i < n; i++) {
    const dx = xs[i] - mx;
    const dy = ys[i] - my;
    sxx += dx * dx;
    syy += dy * dy;
    sxy += dx * dy;
  }
  return { n, mx, my, sxx, syy, sxy };
}

/** A function of two ranges (x first); regression functions take (known_y's, known_x's). */
function twoVariables(names, args, f, yFirst = false) {
  for (const name of [].concat(names)) {
    define(name, {
      min: 2, max: 2, category: 'stat', args,
      fn: ([a, b], ctx) => { const p = yFirst ? pairs(b, a, ctx) : pairs(a, b, ctx); return isErr(p) ? p : p[0].length ? f(moments(p)) : DIV0; },
    });
  }
}

const correl = m => (m.sxx && m.syy ? m.sxy / Math.sqrt(m.sxx * m.syy) : DIV0);
const slope = m => (m.sxx ? m.sxy / m.sxx : DIV0);
const intercept = m => (m.sxx ? m.my - (m.sxy / m.sxx) * m.mx : DIV0);

twoVariables(['CORREL', 'PEARSON'], 'array1, array2', correl);
twoVariables('RSQ', "known_y's, known_x's", m => { const r = correl(m); return isErr(r) ? r : r * r; }, true);
twoVariables(['COVAR', 'COVARIANCE.P'], 'array1, array2', m => m.sxy / m.n);
twoVariables('COVARIANCE.S', 'array1, array2', m => (m.n > 1 ? m.sxy / (m.n - 1) : DIV0));
twoVariables('SLOPE', "known_y's, known_x's", slope, true);
twoVariables('INTERCEPT', "known_y's, known_x's", intercept, true);
twoVariables('STEYX', "known_y's, known_x's", m => (m.n > 2 && m.sxx ? Math.sqrt((m.syy - (m.sxy * m.sxy) / m.sxx) / (m.n - 2)) : DIV0), true);

for (const name of ['FORECAST', 'FORECAST.LINEAR']) {
  define(name, {
    min: 3, max: 3, category: 'stat', args: "x, known_y's, known_x's",
    fn: ([x, ky, kx], ctx) => {
      const p = pairs(kx, ky, ctx);
      if (isErr(p)) return p;
      const m = p[0].length ? moments(p) : null;
      return lift([x], ctx, ([v]) => {
        const t = toNumber(v);
        if (isErr(t)) return t;
        if (!m || !m.sxx) return DIV0;
        return intercept(m) + slope(m) * t;
      });
    },
  });
}

// ---- distributions ---------------------------------------------------------------------------------------------------------------

/** erfc(x) after W. J. Cody's rational Chebyshev approximations (double precision). */
const ERF_A = [3.16112374387056560e0, 1.13864154151050156e2, 3.77485237685302021e2, 3.20937758913846947e3, 1.85777706184603153e-1];
const ERF_B = [2.36012909523441209e1, 2.44024637934444173e2, 1.28261652607737228e3, 2.84423683343917062e3];
const ERF_C = [5.64188496988670089e-1, 8.88314979438837594e0, 6.61191906371416295e1, 2.98635138197400131e2, 8.81952221241769090e2,
  1.71204761263407058e3, 2.05107837782607147e3, 1.23033935479799725e3, 2.15311535474403846e-8];
const ERF_D = [1.57449261107098347e1, 1.17693950891312499e2, 5.37181101862009858e2, 1.62138957456669019e3, 3.29079923573345963e3,
  4.36261909014324716e3, 3.43936767414372164e3, 1.23033935480374942e3];
const ERF_P = [3.05326634961232344e-1, 3.60344899949804439e-1, 1.25781726111229246e-1, 1.60837851487422766e-2, 6.58749161529837803e-4,
  1.63153871373020978e-2];
const ERF_Q = [2.56852019228982242e0, 1.87295284992346725e0, 5.27905102951428412e-1, 6.05183413124413191e-2, 2.33520497626869185e-3];

export function erfc(x) {
  const y = Math.abs(x);
  if (y <= 0.46875) {
    const ysq = y > 1.11e-16 ? y * y : 0;
    let num = ERF_A[4] * ysq;
    let den = ysq;
    for (let i = 0; i < 3; i++) { num = (num + ERF_A[i]) * ysq; den = (den + ERF_B[i]) * ysq; }
    return 1 - (x * (num + ERF_A[3])) / (den + ERF_B[3]);
  }
  let r = 0;
  if (y <= 4) {
    let num = ERF_C[8] * y;
    let den = y;
    for (let i = 0; i < 7; i++) { num = (num + ERF_C[i]) * y; den = (den + ERF_D[i]) * y; }
    r = (num + ERF_C[7]) / (den + ERF_D[7]);
  } else if (y < 26.543) {
    const ysq = 1 / (y * y);
    let num = ERF_P[5] * ysq;
    let den = ysq;
    for (let i = 0; i < 4; i++) { num = (num + ERF_P[i]) * ysq; den = (den + ERF_Q[i]) * ysq; }
    r = (0.56418958354775628695 - (ysq * (num + ERF_P[4])) / (den + ERF_Q[4])) / y;
  }
  if (r) {
    const t = Math.trunc(y * 16) / 16;
    r *= Math.exp(-t * t) * Math.exp(-(y - t) * (y + t));
  }
  return x < 0 ? 2 - r : r;
}

export const normCdf = z => 0.5 * erfc(-z / Math.SQRT2);
export const normPdf = z => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);

/** The inverse of the standard normal distribution (Wichura, AS 241, about 16 digits). */
export function normInv(p) {
  const q = p - 0.5;
  if (Math.abs(q) <= 0.425) {
    const r = 0.180625 - q * q;
    return (q * (((((((2509.0809287301226727 * r + 33430.575583588128105) * r + 67265.770927008700853) * r + 45921.953931549871457) * r
      + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r + 3.387132872796366608))
      / (((((((5226.495278852545925 * r + 28729.085735721942674) * r + 39307.89580009271061) * r + 21213.794301586595867) * r
      + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1);
  }
  let r = Math.sqrt(-Math.log(q < 0 ? p : 1 - p));
  let v;
  if (r <= 5) {
    r -= 1.6;
    v = (((((((7.7454501427834140764e-4 * r + 0.0227238449892691845833) * r + 0.24178072517745061177) * r + 1.27045825245236838258) * r
      + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r + 1.42343711074968357734)
      / (((((((1.05075007164441684324e-9 * r + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r + 0.14810397642748007459) * r
      + 0.68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1);
  } else {
    r -= 5;
    v = (((((((2.01033439929228813265e-7 * r + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r + 0.026532189526576123093) * r
      + 0.29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r + 6.6579046435011037772)
      / (((((((2.04426310338993978564e-15 * r + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r
      + 0.0148753612908506148525) * r + 0.13692988092273580531) * r + 0.59983220655588793769) * r + 1);
  }
  return q < 0 ? -v : v;
}

const LANCZOS = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
  -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];

/** ln Γ(x) for x > 0 (Lanczos, g = 7). */
export function gammaLn(x) {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - gammaLn(1 - x);
  const z = x - 1;
  let a = LANCZOS[0];
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  const t = z + 7.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

const FACTORIALS = [1];
for (let i = 1; i <= 170; i++) FACTORIALS.push(FACTORIALS[i - 1] * i);
/** n! for integers 0 … 170 (exact where a double can hold it), Infinity beyond. */
export const factorial = n => (n <= 170 ? FACTORIALS[n] : Infinity);

/** Γ(x); NaN at 0 and the negative integers. */
export function gamma(x) {
  if (Number.isInteger(x)) return x <= 0 ? NaN : factorial(x - 1);
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  if (x > 171.7) return Infinity;
  const z = x - 1;
  let a = LANCZOS[0];
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  const t = z + 7.5;
  const h = Math.pow(t, (z + 0.5) / 2);
  return Math.sqrt(2 * Math.PI) * h * (h * Math.exp(-t)) * a;
}

/** The binomial coefficient, exact while it fits into 2^53. */
export function combin(n, k) {
  const m = Math.min(k, n - k);
  let c = 1;
  for (let i = 1; i <= m; i++) c = (c * (n - m + i)) / i;
  return c;
}

function binomPmf(k, n, p) {
  if (p === 0 || p === 1) return (p === 0 ? k === 0 : k === n) ? 1 : 0;
  const v = combin(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k);
  if (Number.isFinite(v) && v > 1e-280) return v;
  return Math.exp(gammaLn(n + 1) - gammaLn(k + 1) - gammaLn(n - k + 1) + k * Math.log(p) + (n - k) * Math.log1p(-p));
}

function poissonPmf(k, m) {
  const v = (Math.exp(-m) * Math.pow(m, k)) / factorial(k);
  return Number.isFinite(v) && v > 1e-280 ? v : Math.exp(k * Math.log(m) - m - gammaLn(k + 1));
}

const dist = (names, args, types, f) => {
  for (const name of [].concat(names)) define(name, { min: types.length, max: types.length, category: 'stat', args, fn: scalar(types, f) });
};

dist(['NORM.DIST', 'NORMDIST'], 'x, mean, standard_dev, cumulative', ['number', 'number', 'number', 'bool'],
  (x, m, s, cum) => (s <= 0 ? NUM : cum ? normCdf((x - m) / s) : normPdf((x - m) / s) / s));
dist('NORM.S.DIST', 'z, cumulative', ['number', 'bool'], (z, cum) => (cum ? normCdf(z) : normPdf(z)));
dist('NORMSDIST', 'z', ['number'], normCdf);
dist(['NORM.INV', 'NORMINV'], 'probability, mean, standard_dev', ['number', 'number', 'number'],
  (p, m, s) => (p <= 0 || p >= 1 || s <= 0 ? NUM : m + s * normInv(p)));
dist(['NORM.S.INV', 'NORMSINV'], 'probability', ['number'], p => (p <= 0 || p >= 1 ? NUM : normInv(p)));
dist('LOGNORM.DIST', 'x, mean, standard_dev, cumulative', ['number', 'number', 'number', 'bool'], (x, m, s, cum) => {
  if (x <= 0 || s <= 0) return NUM;
  const z = (Math.log(x) - m) / s;
  return cum ? normCdf(z) : normPdf(z) / (x * s);
});
dist('LOGNORMDIST', 'x, mean, standard_dev', ['number', 'number', 'number'], (x, m, s) => (x <= 0 || s <= 0 ? NUM : normCdf((Math.log(x) - m) / s)));
dist(['LOGNORM.INV', 'LOGINV'], 'probability, mean, standard_dev', ['number', 'number', 'number'],
  (p, m, s) => (p <= 0 || p >= 1 || s <= 0 ? NUM : Math.exp(m + s * normInv(p))));
dist('STANDARDIZE', 'x, mean, standard_dev', ['number', 'number', 'number'], (x, m, s) => (s <= 0 ? NUM : (x - m) / s));
dist('PHI', 'x', ['number'], normPdf);
dist('GAUSS', 'z', ['number'], z => normCdf(z) - 0.5);
dist(['CONFIDENCE.NORM', 'CONFIDENCE'], 'alpha, standard_dev, size', ['number', 'number', 'number'], (a, s, n) => {
  const size = Math.trunc(n);
  return a <= 0 || a >= 1 || s <= 0 || size < 1 ? NUM : (normInv(1 - a / 2) * s) / Math.sqrt(size);
});
dist(['BINOM.DIST', 'BINOMDIST'], 'number_s, trials, probability_s, cumulative', ['number', 'number', 'number', 'bool'], (k, n, p, cum) => {
  k = Math.trunc(k);
  n = Math.trunc(n);
  if (k < 0 || n < 0 || k > n || p < 0 || p > 1) return NUM;
  if (!cum) return binomPmf(k, n, p);
  let s = 0;
  for (let i = 0; i <= k; i++) s += binomPmf(i, n, p);
  return Math.min(1, s);
});
dist(['POISSON.DIST', 'POISSON'], 'x, mean, cumulative', ['number', 'number', 'bool'], (x, m, cum) => {
  x = Math.trunc(x);
  if (x < 0 || m < 0) return NUM;
  if (!cum) return poissonPmf(x, m);
  let s = 0;
  for (let i = 0; i <= x; i++) s += poissonPmf(i, m);
  return Math.min(1, s);
});
dist(['EXPON.DIST', 'EXPONDIST'], 'x, lambda, cumulative', ['number', 'number', 'bool'],
  (x, l, cum) => (x < 0 || l <= 0 ? NUM : cum ? -Math.expm1(-l * x) : l * Math.exp(-l * x)));
dist(['WEIBULL.DIST', 'WEIBULL'], 'x, alpha, beta, cumulative', ['number', 'number', 'number', 'bool'], (x, a, b, cum) => {
  if (x < 0 || a <= 0 || b <= 0) return NUM;
  const e = Math.exp(-Math.pow(x / b, a));
  return cum ? 1 - e : (a / Math.pow(b, a)) * Math.pow(x, a - 1) * e;
});
dist('GAMMA', 'number', ['number'], x => (x <= 0 && Number.isInteger(x) ? NUM : gamma(x)));
dist(['GAMMALN', 'GAMMALN.PRECISE'], 'x', ['number'], x => (x <= 0 ? NUM : gammaLn(x)));
dist('FISHER', 'x', ['number'], x => (x <= -1 || x >= 1 ? NUM : 0.5 * Math.log((1 + x) / (1 - x))));
dist('FISHERINV', 'y', ['number'], Math.tanh);

export { toBool };
