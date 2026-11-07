/* MatOffice Sheets – math & trigonometry functions (pure, no DOM). */
import { define, scalar, numbers, matrix, each, makeMatrix, flat, Errors, Ref } from '../lib.js';
import { FormulaError, toNumber } from '../values.js';
import {
  grid, lift, optNumber, sorted, sum, mean, maxOf, minOf, product, variance, stdev, median, mode, kth, percentile, quartile, ifs, ifOne, pairs,
  factorial, combin,
} from './stat.js';

const isErr = v => v instanceof FormulaError;
const { NUM, DIV0, VALUE } = Errors;

/** A number kept to 15 significant digits, the precision Excel works with (2.9999999999999996 → 3). */
const snap = x => (x === 0 || !Number.isFinite(x) ? x : Number(x.toPrecision(15)));

const fn1 = (name, f, args = 'number') => define(name, { min: 1, max: 1, category: 'math', args, fn: scalar(['number'], f) });
const fn2 = (name, args, f, types = ['number', 'number']) => define(name, { min: types.filter(t => !t.endsWith('?')).length, max: types.length, category: 'math', args, fn: scalar(types, f) });
const ofNumbers = (name, f) => define(name, {
  min: 1, max: 255, category: 'math', args: 'number1, [number2], …',
  fn: (args, ctx) => { const list = numbers(args, ctx); return isErr(list) ? list : f(list); },
});

// ---- sums and products ------------------------------------------------------------------------------------------------------------

ofNumbers('SUM', sum);
ofNumbers('PRODUCT', product);
ofNumbers('SUMSQ', l => sum(l.map(x => x * x)));

define('SUMIF', { min: 2, max: 3, category: 'math', args: 'range, criteria, [sum_range]', fn: ifOne(sum) });
define('SUMIFS', {
  min: 3, max: 255, category: 'math', args: 'sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …',
  fn: ([values, ...rest], ctx) => (isErr(values) ? values : ifs(grid(values, ctx), rest, ctx, sum)),
});

/** SUMPRODUCT: arrays of one size multiplied element by element (anything but a number counts 0) and summed. */
define('SUMPRODUCT', {
  min: 1, max: 255, category: 'math', args: 'array1, [array2], [array3], …',
  fn: (args, ctx) => {
    const grids = [];
    for (const a of args) {
      if (isErr(a)) return a;
      const g = grid(a, ctx);
      if (grids.length && (g.rows !== grids[0].rows || g.cols !== grids[0].cols)) return VALUE;
      grids.push(g);
    }
    let dr = 0;
    let dc = 0;
    for (const g of grids) { dr = Math.max(dr, g.m.length); dc = Math.max(dc, g.m[0].length); }
    let s = 0;
    for (let r = 0; r < dr; r++) {
      for (let c = 0; c < dc; c++) {
        let p = 1;
        for (const g of grids) {
          const v = g.at(r, c);
          if (isErr(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        s += p;
      }
    }
    return s;
  },
});

const sumOfPairs = (name, f) => define(name, {
  min: 2, max: 2, category: 'math', args: 'array_x, array_y',
  fn: ([a, b], ctx) => { const p = pairs(a, b, ctx); return isErr(p) ? p : sum(p[0].map((x, i) => f(x, p[1][i]))); },
});
sumOfPairs('SUMX2MY2', (x, y) => x * x - y * y);
sumOfPairs('SUMX2PY2', (x, y) => x * x + y * y);
sumOfPairs('SUMXMY2', (x, y) => (x - y) * (x - y));

define('SERIESSUM', {
  min: 4, max: 4, category: 'math', args: 'x, n, m, coefficients',
  fn: ([x, n, m, coefficients], ctx) => {
    const a = [x, n, m].map(v => optNumber(v, ctx, 0));
    const e = a.find(isErr);
    if (e) return e;
    let s = 0;
    const list = flat(coefficients, ctx);
    for (let i = 0; i < list.length; i++) {
      const c = list[i] ?? 0;
      if (isErr(c)) return c;
      if (typeof c !== 'number') return VALUE;
      s += c * a[0] ** (a[1] + i * a[2]);
    }
    return s;
  },
});

// ---- rounding ---------------------------------------------------------------------------------------------------------------------

/** x rounded to digits (negative: left of the decimal point) by f on |x|: Math.round (halves away from zero), Math.ceil (away from zero), Math.trunc. */
function roundWith(x, digits, f) {
  const d = Math.trunc(digits);
  if (x === 0 || !Number.isFinite(x)) return x;
  const a = Math.abs(x);
  let r;
  if (d >= 0) {
    const p = 10 ** d;
    if (a * p >= 1e15) return x; // no digit to round within the 15 significant digits
    r = f(Number((a * p).toPrecision(15))) / p;
  } else {
    if (d < -308) return f === Math.ceil ? Infinity : 0;
    const p = 10 ** -d;
    r = f(Number((a / p).toPrecision(15))) * p;
  }
  return x < 0 ? -r : r;
}

/** ROUND like Excel: halves away from zero, decimal digits may be negative (ROUND(1234, -2) = 1200). */
export const roundTo = (x, digits) => roundWith(x, digits, Math.round);

fn2('ROUND', 'number, num_digits', (x, d) => roundWith(x, d, Math.round));
fn2('ROUNDUP', 'number, num_digits', (x, d) => roundWith(x, d, Math.ceil));
fn2('ROUNDDOWN', 'number, num_digits', (x, d) => roundWith(x, d, Math.trunc));
fn2('TRUNC', 'number, [num_digits]', (x, d = 0) => roundWith(x, d, Math.trunc), ['number', 'number?']);
fn1('INT', Math.floor);

/** n as a multiple of s, f (Math.ceil / Math.floor / Math.round) applied to the quotient; both keep 15 digits (CEILING(1.1, 0.1) = 1.1). */
const toMultiple = (n, s, f) => snap(f(snap(n / s)) * s);

fn2('MROUND', 'number, multiple', (n, m) => {
  if (n === 0 || m === 0) return 0;
  if (n * m < 0) return NUM;
  const v = toMultiple(Math.abs(n), Math.abs(m), Math.round);
  return n < 0 ? -v : v;
});
fn2('CEILING', 'number, significance', (n, s) => {
  if (n === 0 || s === 0) return 0;
  if (n > 0 && s < 0) return NUM;
  return s < 0 ? -toMultiple(-n, -s, Math.ceil) : toMultiple(n, s, Math.ceil);
});
fn2('FLOOR', 'number, significance', (n, s) => {
  if (n === 0) return 0;
  if (s === 0) return DIV0;
  if (n > 0 && s < 0) return NUM;
  return s < 0 ? -toMultiple(-n, -s, Math.floor) : toMultiple(n, s, Math.floor);
});

/** CEILING.MATH/FLOOR.MATH: the sign of significance is ignored; mode ≠ 0 rounds negative numbers the other way (away from / toward zero). */
const mathRounding = (up, withMode) => (n, s = 1, m = 0) => {
  if (n === 0 || s === 0) return 0;
  const a = Math.abs(s);
  const f = up ? Math.ceil : Math.floor;
  return n < 0 && withMode && m ? -toMultiple(-n, a, f) : toMultiple(n, a, f);
};
fn2('CEILING.MATH', 'number, [significance], [mode]', mathRounding(true, true), ['number', 'number?', 'number?']);
fn2('FLOOR.MATH', 'number, [significance], [mode]', mathRounding(false, true), ['number', 'number?', 'number?']);
fn2('CEILING.PRECISE', 'number, [significance]', mathRounding(true, false), ['number', 'number?']);
fn2('ISO.CEILING', 'number, [significance]', mathRounding(true, false), ['number', 'number?']);
fn2('FLOOR.PRECISE', 'number, [significance]', mathRounding(false, false), ['number', 'number?']);

fn1('EVEN', x => { const v = toMultiple(Math.abs(x), 2, Math.ceil); return x < 0 ? -v : v; });
fn1('ODD', x => {
  let v = Math.ceil(snap(Math.abs(x)));
  if (v % 2 === 0) v++;
  return x < 0 ? -v : v;
});

// ---- arithmetic -------------------------------------------------------------------------------------------------------------------

fn1('ABS', Math.abs);
fn1('SIGN', Math.sign);
fn1('SQRT', x => (x < 0 ? NUM : Math.sqrt(x)));
fn1('SQRTPI', x => (x < 0 ? NUM : Math.sqrt(x * Math.PI)));
fn1('EXP', Math.exp);
fn1('LN', x => (x <= 0 ? NUM : Math.log(x)));
fn1('LOG10', x => (x <= 0 ? NUM : Math.log10(x)));
fn2('LOG', 'number, [base]', (x, b = 10) => {
  if (x <= 0 || b <= 0) return NUM;
  if (b === 1) return DIV0;
  return b === 10 ? Math.log10(x) : b === 2 ? Math.log2(x) : snap(Math.log(x) / Math.log(b));
}, ['number', 'number?']);
fn2('POWER', 'number, power', (a, b) => (a === 0 && b === 0 ? NUM : a === 0 && b < 0 ? DIV0 : a ** b));
fn2('MOD', 'number, divisor', (n, d) => (d === 0 ? DIV0 : n - d * Math.floor(n / d)));
fn2('QUOTIENT', 'numerator, denominator', (n, d) => (d === 0 ? DIV0 : Math.trunc(n / d)));

define('PI', { min: 0, max: 0, category: 'math', args: '', fn: () => Math.PI });
define('RAND', { min: 0, max: 0, category: 'math', volatile: true, args: '', fn: (args, ctx) => ctx.random() });
define('RANDBETWEEN', {
  min: 2, max: 2, category: 'math', volatile: true, args: 'bottom, top',
  fn: (args, ctx) => scalar(['number', 'number'], (a, b) => {
    const lo = Math.ceil(a);
    const hi = Math.floor(b);
    return lo > hi ? NUM : lo + Math.floor(ctx.random() * (hi - lo + 1));
  })(args, ctx),
});

fn1('FACT', x => { const n = Math.trunc(x); return n < 0 ? NUM : factorial(n); });
fn1('FACTDOUBLE', x => {
  const n = Math.trunc(x);
  if (n < -1) return NUM;
  if (n > 400) return Infinity;
  let p = 1;
  for (let i = n; i > 1; i -= 2) p *= i;
  return p;
});
const exact = c => (c < 2 ** 53 ? Math.round(c) : c);
fn2('COMBIN', 'number, number_chosen', (a, b) => {
  const n = Math.trunc(a);
  const k = Math.trunc(b);
  return n < 0 || k < 0 || n < k ? NUM : exact(combin(n, k));
});
fn2('COMBINA', 'number, number_chosen', (a, b) => {
  const n = Math.trunc(a);
  const k = Math.trunc(b);
  if (n < 0 || k < 0) return NUM;
  return n === 0 ? (k === 0 ? 1 : 0) : exact(combin(n + k - 1, k));
});
fn2('PERMUT', 'number, number_chosen', (a, b) => {
  const n = Math.trunc(a);
  const k = Math.trunc(b);
  if (n < 0 || k < 0 || n < k) return NUM;
  let p = 1;
  for (let i = n - k + 1; i <= n && Number.isFinite(p); i++) p *= i;
  return p;
});
fn2('PERMUTATIONA', 'number, number_chosen', (a, b) => {
  const n = Math.trunc(a);
  const k = Math.trunc(b);
  return n < 0 || k < 0 ? NUM : n ** k;
});

/** The whole numbers of the arguments for GCD/LCM/MULTINOMIAL (truncated; negative or ≥ 2^53 → #NUM!). */
function wholeNumbers(args, ctx) {
  const list = numbers(args, ctx);
  if (isErr(list)) return list;
  const out = [];
  for (const x of list) {
    if (x < 0 || x >= 2 ** 53) return NUM;
    out.push(Math.trunc(x));
  }
  return out;
}
const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };
const ofWholeNumbers = (name, f) => define(name, {
  min: 1, max: 255, category: 'math', args: 'number1, [number2], …',
  fn: (args, ctx) => { const list = wholeNumbers(args, ctx); return isErr(list) ? list : f(list); },
});
ofWholeNumbers('GCD', l => l.reduce(gcd2, 0));
ofWholeNumbers('LCM', l => {
  if (l.some(x => x === 0)) return 0;
  const r = l.reduce((a, b) => (a / gcd2(a, b)) * b, 1);
  return r >= 2 ** 53 ? NUM : r;
});
ofWholeNumbers('MULTINOMIAL', l => {
  let total = 0;
  let r = 1;
  for (const x of l) { total += x; r *= combin(total, x); }
  return exact(r);
});

// ---- trigonometry ----------------------------------------------------------------------------------------------------------------

const trig = f => x => (Math.abs(x) >= 134217728 ? NUM : f(x));
fn1('SIN', trig(Math.sin));
fn1('COS', trig(Math.cos));
fn1('TAN', trig(Math.tan));
fn1('ASIN', x => (x < -1 || x > 1 ? NUM : Math.asin(x)));
fn1('ACOS', x => (x < -1 || x > 1 ? NUM : Math.acos(x)));
fn1('ATAN', Math.atan);
fn2('ATAN2', 'x_num, y_num', (x, y) => (x === 0 && y === 0 ? DIV0 : Math.atan2(y, x)));
fn1('SINH', Math.sinh);
fn1('COSH', Math.cosh);
fn1('TANH', Math.tanh);
fn1('ASINH', Math.asinh);
fn1('ACOSH', x => (x < 1 ? NUM : Math.acosh(x)));
fn1('ATANH', x => (x <= -1 || x >= 1 ? NUM : Math.atanh(x)));
fn1('COT', trig(x => (x === 0 ? DIV0 : 1 / Math.tan(x))));
fn1('CSC', trig(x => (x === 0 ? DIV0 : 1 / Math.sin(x))));
fn1('SEC', trig(x => 1 / Math.cos(x)));
fn1('COTH', x => (x === 0 ? DIV0 : 1 / Math.tanh(x)));
fn1('CSCH', x => (x === 0 ? DIV0 : 1 / Math.sinh(x)));
fn1('SECH', x => 1 / Math.cosh(x));
fn1('ACOT', x => Math.PI / 2 - Math.atan(x));
fn1('ACOTH', x => (Math.abs(x) <= 1 ? NUM : 0.5 * Math.log((x + 1) / (x - 1))));
fn1('DEGREES', x => (x * 180) / Math.PI, 'angle');
fn1('RADIANS', x => (x * Math.PI) / 180, 'angle');

// ---- SUBTOTAL and AGGREGATE --------------------------------------------------------------------------------------------------------

const NESTED = /\b(SUBTOTAL|AGGREGATE)\s*\(/i;

/** The values of SUBTOTAL/AGGREGATE arguments: references skip hidden rows (when asked) and the cells of other SUBTOTAL/AGGREGATE formulas. */
function subtotalValues(args, ctx, hidden, nested) {
  const out = [];
  for (const a of args) {
    if (!(a instanceof Ref)) { each(a, ctx, v => { out.push(v); }); continue; }
    const m = ctx.values(a);
    for (let r = 0; r < m.length; r++) {
      const row = a.r1 + r;
      if (hidden === true ? ctx.rowHidden(a.sheet, row) : hidden === 'filtered' && ctx.rowFiltered?.(a.sheet, row)) continue;
      for (let c = 0; c < m[r].length; c++) {
        if (nested && m[r][c] !== null && NESTED.test(ctx.formulaAt(a.sheet, row, a.c1 + c) || '')) continue;
        out.push(m[r][c]);
      }
    }
  }
  return out;
}

/** The functions 1 … 13 of SUBTOTAL (1 … 11) and AGGREGATE: AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, STDEV.S, STDEV.P, SUM, VAR.S, VAR.P, MEDIAN, MODE.SNGL. */
const STATS = [null, mean, null, null, maxOf, minOf, product, l => stdev(l, true), l => stdev(l, false), sum, l => variance(l, true),
  l => variance(l, false), median, mode];

function statistic(f, values, ignoreErrors) {
  if (f === 3) {
    let n = 0;
    for (const v of values) if (v !== null && !(ignoreErrors && isErr(v))) n++;
    return n;
  }
  const list = [];
  for (const v of values) {
    if (typeof v === 'number') list.push(v);
    else if (isErr(v) && !ignoreErrors && f !== 2) return v;
  }
  return f === 2 ? list.length : STATS[f](list);
}

/** SUBTOTAL(function_num, ref1, …): 1 … 11 (101 … 111 also skip hidden rows); cells with SUBTOTAL/AGGREGATE formulas are left out. */
define('SUBTOTAL', {
  min: 2, max: 255, category: 'math', volatile: true, args: 'function_num, ref1, [ref2], …',
  fn: ([num, ...refs], ctx) => lift([num], ctx, ([v]) => {
    const n = toNumber(v);
    if (isErr(n)) return n;
    const f = Math.trunc(n);
    const kind = f > 100 ? f - 100 : f;
    if (kind < 1 || kind > 11 || (f > 11 && f < 101)) return VALUE;
    // 1 … 11 leave out the rows the AutoFilter hides, 101 … 111 every hidden row (like Excel)
    return statistic(kind, subtotalValues(refs, ctx, f > 100 ? true : 'filtered', true), false);
  }),
});

/**
 * AGGREGATE(function_num, options, ref1, …) for 1 … 13, AGGREGATE(function_num, options, array, k) for 14 … 19 (LARGE, SMALL, PERCENTILE.INC,
 * QUARTILE.INC, PERCENTILE.EXC, QUARTILE.EXC). Options 0 … 7 leave out nested SUBTOTAL/AGGREGATE (0–3), hidden rows (odd) and errors (2, 3, 6, 7).
 */
define('AGGREGATE', {
  min: 3, max: 255, category: 'math', volatile: true, args: 'function_num, options, ref1, [ref2], …',
  fn: ([num, opt, ...rest], ctx) => {
    const f = optNumber(num, ctx, NaN);
    if (isErr(f)) return f;
    const o = optNumber(opt, ctx, 0);
    if (isErr(o)) return o;
    const fi = Math.trunc(f);
    const oi = Math.trunc(o);
    if (!(fi >= 1 && fi <= 19) || oi < 0 || oi > 7) return VALUE;
    const hidden = oi % 2 === 1;
    const errors = [2, 3, 6, 7].includes(oi);
    const nested = oi < 4;
    if (fi <= 13) return statistic(fi, subtotalValues(rest, ctx, hidden, nested), errors);
    if (rest.length < 2) return VALUE;
    const list = [];
    for (const v of subtotalValues([rest[0]], ctx, hidden, nested)) {
      if (typeof v === 'number') list.push(v);
      else if (isErr(v) && !errors) return v;
    }
    const s = sorted(list);
    return lift([rest[1]], ctx, ([k]) => {
      const x = toNumber(k);
      if (isErr(x)) return x;
      switch (fi) {
        case 14: return kth(s, x, true);
        case 15: return kth(s, x, false);
        case 16: return percentile(s, x, false);
        case 17: return quartile(s, x, false);
        case 18: return x <= 0 || x >= 1 ? NUM : percentile(s, x, true);
        default: return quartile(s, x, true);
      }
    });
  },
});

// ---- number systems ---------------------------------------------------------------------------------------------------------------

const ROMAN_CHARS = 'MDCLXVI';
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** A number 0 … 3999 as Roman numerals; form 0 (classic) … 4 (simplified) allows more subtractions (499: CDXCIX, LDVLIV, XDIX, VDIV, ID). */
function roman(value, form) {
  let v = value;
  let out = '';
  for (let i = 0; i <= 3; i++) {
    let idx = 2 * i;
    const digit = Math.floor(v / ROMAN_VALUES[idx]);
    if (digit % 5 === 4) {
      const idx2 = digit === 4 ? idx - 1 : idx - 2;
      for (let steps = 0; steps < form && idx < 6; steps++) {
        if (ROMAN_VALUES[idx2] - ROMAN_VALUES[idx + 1] <= v) idx++;
        else break;
      }
      out += ROMAN_CHARS[idx] + ROMAN_CHARS[idx2];
      v += ROMAN_VALUES[idx] - ROMAN_VALUES[idx2];
    } else {
      if (digit > 4) out += ROMAN_CHARS[idx - 1];
      out += ROMAN_CHARS[idx].repeat(digit % 5);
      v %= ROMAN_VALUES[idx];
    }
  }
  return out;
}

fn2('ROMAN', 'number, [form]', (x, form) => {
  const n = Math.trunc(x);
  let f = 0;
  if (typeof form === 'boolean') f = form ? 0 : 4;
  else if (form !== undefined) {
    const v = toNumber(form);
    if (isErr(v)) return v;
    f = Math.trunc(v);
  }
  return n < 0 || n > 3999 || f < 0 || f > 4 ? VALUE : roman(n, f);
}, ['number', 'any?']);

fn2('ARABIC', 'text', t => {
  let s = t.trim().toUpperCase();
  if (s.length > 255) return VALUE;
  let sign = 1;
  if (s.startsWith('-')) { sign = -1; s = s.slice(1); }
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const v = ROMAN_VALUES[ROMAN_CHARS.indexOf(s[i])];
    if (!v) return VALUE;
    const next = ROMAN_VALUES[ROMAN_CHARS.indexOf(s[i + 1])] || 0;
    total += v < next ? -v : v;
  }
  return sign * total;
}, ['text']);

fn2('BASE', 'number, radix, [min_length]', (x, radix, len = 0) => {
  const n = Math.trunc(x);
  const r = Math.trunc(radix);
  const l = Math.trunc(len);
  if (n < 0 || n >= 2 ** 53 || r < 2 || r > 36 || l < 0 || l > 255) return NUM;
  return n.toString(r).toUpperCase().padStart(l, '0');
}, ['number', 'number', 'number?']);

fn2('DECIMAL', 'text, radix', (t, radix) => {
  const r = Math.trunc(radix);
  if (r < 2 || r > 36 || t.length > 255) return NUM;
  let v = 0;
  for (const ch of t.toUpperCase()) {
    const d = parseInt(ch, 36);
    if (Number.isNaN(d) || d >= r) return NUM;
    v = v * r + d;
  }
  return v;
}, ['text', 'number']);

// ---- matrices --------------------------------------------------------------------------------------------------------------------

/** An argument as a matrix of numbers: anything else in it is #VALUE!, an error stays. */
function numericMatrix(v, ctx) {
  if (isErr(v)) return v;
  const m = matrix(v, ctx);
  for (const row of m) {
    for (const x of row) {
      if (isErr(x)) return x;
      if (typeof x !== 'number') return VALUE;
    }
  }
  return m;
}

/** Gauss–Jordan elimination with partial pivoting: the determinant and (when it is not 0) the inverse. */
function eliminate(m) {
  const n = m.length;
  const a = m.map((row, i) => [...row, ...row.map((_, j) => (i === j ? 1 : 0))]);
  let det = 1;
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r;
    if (a[p][i] === 0) return { det: 0, inverse: null };
    if (p !== i) { [a[p], a[i]] = [a[i], a[p]]; det = -det; }
    const pivot = a[i][i];
    det *= pivot;
    for (let c = 0; c < 2 * n; c++) a[i][c] /= pivot;
    for (let r = 0; r < n; r++) {
      const f = a[r][i];
      if (r !== i && f) for (let c = 0; c < 2 * n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return { det, inverse: a.map(row => row.slice(n)) };
}

define('MMULT', {
  min: 2, max: 2, category: 'math', args: 'array1, array2',
  fn: ([a, b], ctx) => {
    const x = numericMatrix(a, ctx);
    if (isErr(x)) return x;
    const y = numericMatrix(b, ctx);
    if (isErr(y)) return y;
    if (x[0].length !== y.length) return VALUE;
    return makeMatrix(x.length, y[0].length, (r, c) => {
      let s = 0;
      for (let k = 0; k < y.length; k++) s += x[r][k] * y[k][c];
      return s;
    });
  },
});

const square = (name, f) => define(name, {
  min: 1, max: 1, category: 'math', args: 'array',
  fn: ([a], ctx) => {
    const m = numericMatrix(a, ctx);
    if (isErr(m)) return m;
    return m.length === m[0].length ? f(eliminate(m)) : VALUE;
  },
});
square('MDETERM', e => snap(e.det));
square('MINVERSE', e => e.inverse || NUM);

fn2('MUNIT', 'dimension', x => {
  const n = Math.trunc(x);
  return n < 1 || n > 16384 ? VALUE : makeMatrix(n, n, (r, c) => (r === c ? 1 : 0));
}, ['number']);
