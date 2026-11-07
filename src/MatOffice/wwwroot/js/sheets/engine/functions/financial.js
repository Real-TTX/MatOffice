/* MatOffice Sheets – financial functions (pure, no DOM). Money paid is negative, money received positive; type 1 means payments at the start
   of each period. The time value of money: pv·(1+r)^n + pmt·(1+r·type)·((1+r)^n − 1)/r + fv = 0. */
import { define, scalar, numbers, flat, Errors } from '../lib.js';
import { FormulaError, toNumber } from '../values.js';
import { optNumber } from './stat.js';

const isErr = v => v instanceof FormulaError;
const { NUM, DIV0, VALUE } = Errors;
/** A number kept to 15 significant digits (DOLLARDE(1.02, 16) = 1.125, not 1.1250000000000002). */
const snap = x => (x === 0 || !Number.isFinite(x) ? x : Number(x.toPrecision(15)));

const fin = (name, args, types, f) => define(name, {
  min: types.filter(t => !t.endsWith('?')).length, max: types.length, category: 'financial', args, fn: scalar(types, f),
});
const N = 'number';
const O = 'number?';

/** (1 + r)^n − 1, also precise for tiny rates. */
const growth = (r, n) => (Math.abs(r) < 1e-4 ? Math.expm1(n * Math.log1p(r)) : (1 + r) ** n - 1);

export function pmt(r, n, pv, fv = 0, type = 0) {
  if (r === 0) return -(pv + fv) / n;
  const g = growth(r, n);
  return -(pv * (1 + g) + fv) / (((1 + r * type) * g) / r);
}

export function fv(r, n, p, pv = 0, type = 0) {
  if (r === 0) return -(pv + p * n);
  const g = growth(r, n);
  return -(pv * (1 + g) + (p * (1 + r * type) * g) / r);
}

export function pv(r, n, p, fv = 0, type = 0) {
  if (r === 0) return -(fv + p * n);
  const g = growth(r, n);
  return -(fv + (p * (1 + r * type) * g) / r) / (1 + g);
}

/** The interest part of payment per (1 … n). */
function ipmt(r, per, n, pvalue, fvalue, type) {
  const p = pmt(r, n, pvalue, fvalue, type);
  const balance = per === 1 ? (type ? 0 : -pvalue) : type ? fv(r, per - 2, p, pvalue, 1) - p : fv(r, per - 1, p, pvalue, 0);
  return balance * r;
}

const flag = t => (t ? 1 : 0);

fin('PMT', 'rate, nper, pv, [fv], [type]', [N, N, N, O, O], (r, n, p, f = 0, t = 0) => (n === 0 ? NUM : pmt(r, n, p, f, flag(t))));
fin('FV', 'rate, nper, pmt, [pv], [type]', [N, N, N, O, O], (r, n, p, v = 0, t = 0) => fv(r, n, p, v, flag(t)));
fin('PV', 'rate, nper, pmt, [fv], [type]', [N, N, N, O, O], (r, n, p, f = 0, t = 0) => pv(r, n, p, f, flag(t)));
fin('IPMT', 'rate, per, nper, pv, [fv], [type]', [N, N, N, N, O, O],
  (r, per, n, p, f = 0, t = 0) => (per < 1 || per > n ? NUM : ipmt(r, per, n, p, f, flag(t))));
fin('PPMT', 'rate, per, nper, pv, [fv], [type]', [N, N, N, N, O, O],
  (r, per, n, p, f = 0, t = 0) => (per < 1 || per > n ? NUM : pmt(r, n, p, f, flag(t)) - ipmt(r, per, n, p, f, flag(t))));
fin('NPER', 'rate, pmt, pv, [fv], [type]', [N, N, N, O, O], (r, p, v, f = 0, t = 0) => {
  if (r === 0) return p === 0 ? NUM : -(v + f) / p;
  const a = (p * (1 + r * flag(t))) / r;
  const x = (a - f) / (a + v);
  return x > 0 ? Math.log(x) / Math.log1p(r) : NUM;
});
fin('ISPMT', 'rate, per, nper, pv', [N, N, N, N], (r, per, n, v) => (n === 0 ? DIV0 : v * r * (per / n - 1)));

/**
 * Newton's method from a guess: f(x) → [value, derivative]; the root, or #NUM! when it does not converge or what it converges to is no root
 * (|value| above a millionth of scale, the size of the terms).
 */
function solve(f, guess, scale) {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const [y, dy] = f(x);
    if (!Number.isFinite(y) || !Number.isFinite(dy) || dy === 0) return NUM;
    let next = x - y / dy;
    if (next <= -1) next = (x - 1) / 2; // rates stay above -100 %
    if (Math.abs(next - x) <= 1e-13 * Math.max(1, Math.abs(next))) return Math.abs(f(next)[0]) <= 1e-7 * scale ? next : NUM;
    x = next;
  }
  return NUM;
}

/** RATE: the rate of the annuity equation by Newton's method (near 0 by its series). */
fin('RATE', 'nper, pmt, pv, [fv], [type], [guess]', [N, N, N, O, O, O], (n, p, v, f = 0, t = 0, guess = 0.1) => {
  const type = flag(t);
  return solve(r => {
    if (Math.abs(r) < 1e-10) return [v + p * n + f, v * n + p * (type * n + (n * (n - 1)) / 2)];
    const g = growth(r, n);
    const dq = (n * (1 + g)) / (1 + r);
    const y = v * (1 + g) + (p * (1 + r * type) * g) / r + f;
    const dy = v * dq + p * ((type * g) / r + ((1 + r * type) * (dq * r - g)) / (r * r));
    return [y, dy];
  }, guess, Math.abs(v) + Math.abs(p * n) + Math.abs(f));
});

define('NPV', {
  min: 2, max: 255, category: 'financial', args: 'rate, value1, [value2], …',
  fn: ([rate, ...values], ctx) => {
    const r = optNumber(rate, ctx, 0);
    if (isErr(r)) return r;
    const list = numbers(values, ctx);
    if (isErr(list)) return list;
    if (r === -1) return DIV0;
    let s = 0;
    for (let i = 0; i < list.length; i++) s += list[i] / (1 + r) ** (i + 1);
    return s;
  },
});

define('IRR', {
  min: 1, max: 2, category: 'financial', args: 'values, [guess]',
  fn: ([values, guess], ctx) => {
    const list = numbers([values], ctx);
    if (isErr(list)) return list;
    const g = optNumber(guess, ctx, 0.1);
    if (isErr(g)) return g;
    if (!list.some(x => x > 0) || !list.some(x => x < 0)) return NUM;
    return solve(r => {
      let y = 0;
      let dy = 0;
      for (let i = 0; i < list.length; i++) {
        const d = (1 + r) ** i;
        y += list[i] / d;
        dy -= (i * list[i]) / (d * (1 + r));
      }
      return [y, dy];
    }, g, list.reduce((t, x) => t + Math.abs(x), 0));
  },
});

define('MIRR', {
  min: 3, max: 3, category: 'financial', args: 'values, finance_rate, reinvest_rate',
  fn: ([values, financeRate, reinvestRate], ctx) => {
    const list = numbers([values], ctx);
    if (isErr(list)) return list;
    const f = optNumber(financeRate, ctx, 0);
    if (isErr(f)) return f;
    const r = optNumber(reinvestRate, ctx, 0);
    if (isErr(r)) return r;
    const n = list.length;
    let gains = 0;
    let costs = 0;
    list.forEach((v, i) => {
      if (v > 0) gains += v * (1 + r) ** (n - 1 - i);
      else if (v < 0) costs += v / (1 + f) ** i;
    });
    return gains && costs ? (-gains / costs) ** (1 / (n - 1)) - 1 : DIV0;
  },
});

/** The cash flows of XNPV/XIRR: [values, whole days]; the counts must match, values be numbers, no date before the first. */
function cashflows(values, dates, ctx) {
  const v = flat(values, ctx);
  const d = flat(dates, ctx);
  if (v.length !== d.length || !v.length) return NUM;
  const days = [];
  for (let i = 0; i < v.length; i++) {
    if (isErr(v[i])) return v[i];
    if (isErr(d[i])) return d[i];
    if (typeof v[i] !== 'number') return VALUE;
    const day = toNumber(d[i]);
    if (isErr(day)) return day;
    days.push(Math.trunc(day));
  }
  return days.some(x => x < days[0] || x < 0) ? NUM : [v, days];
}

/** The net present value of dated cash flows and its derivative by the rate. */
function xnpv(r, [values, days]) {
  let y = 0;
  let dy = 0;
  for (let i = 0; i < values.length; i++) {
    const t = (days[i] - days[0]) / 365;
    const d = (1 + r) ** t;
    y += values[i] / d;
    dy -= (t * values[i]) / (d * (1 + r));
  }
  return [y, dy];
}

define('XNPV', {
  min: 3, max: 3, category: 'financial', args: 'rate, values, dates',
  fn: ([rate, values, dates], ctx) => {
    const r = optNumber(rate, ctx, 0);
    if (isErr(r)) return r;
    const flows = cashflows(values, dates, ctx);
    if (isErr(flows)) return flows;
    return r <= -1 ? NUM : xnpv(r, flows)[0];
  },
});

define('XIRR', {
  min: 2, max: 3, category: 'financial', args: 'values, dates, [guess]',
  fn: ([values, dates, guess], ctx) => {
    const flows = cashflows(values, dates, ctx);
    if (isErr(flows)) return flows;
    const g = optNumber(guess, ctx, 0.1);
    if (isErr(g)) return g;
    if (!flows[0].some(x => x > 0) || !flows[0].some(x => x < 0)) return NUM;
    return solve(r => xnpv(r, flows), g, flows[0].reduce((t, x) => t + Math.abs(x), 0));
  },
});

/** CUMIPMT/CUMPRINC: the interest (principal) paid from period start to end (whole periods, type 0 or 1, positive rate, nper and pv). */
function cumulative(principal) {
  return (r, n, v, start, end, type) => {
    const s = Math.trunc(start);
    const e = Math.trunc(end);
    if (s < 1 || e < s || e > n || r <= 0 || n <= 0 || v <= 0 || (type !== 0 && type !== 1)) return NUM;
    const p = pmt(r, n, v, 0, type);
    let sum = 0;
    let from = s;
    if (s === 1) {
      sum = principal ? (type ? p : p + v * r) : type ? 0 : -v * r;
      from = 2;
    }
    for (let i = from; i <= e; i++) {
      const interest = (type ? fv(r, i - 2, p, v, 1) - p : fv(r, i - 1, p, v, 0)) * r;
      sum += principal ? p - interest : interest;
    }
    return sum;
  };
}

const cumArgs = 'rate, nper, pv, start_period, end_period, type';
fin('CUMIPMT', cumArgs, [N, N, N, N, N, N], cumulative(false));
fin('CUMPRINC', cumArgs, [N, N, N, N, N, N], cumulative(true));

// ---- depreciation ---------------------------------------------------------------------------------------------------------------

fin('SLN', 'cost, salvage, life', [N, N, N], (c, s, l) => (l === 0 ? DIV0 : (c - s) / l));
fin('SYD', 'cost, salvage, life, per', [N, N, N, N], (c, s, l, p) => (l <= 0 || p <= 0 || p > l ? NUM : ((c - s) * (l - p + 1) * 2) / (l * (l + 1))));

/** DB: fixed-declining balance with the rate rounded to three decimals; month = the months of the first year. */
fin('DB', 'cost, salvage, life, period, [month]', [N, N, N, N, O], (cost, salvage, life, period, month = 12) => {
  const months = Math.trunc(month);
  if (months < 1 || months > 12 || cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life + 1) return NUM;
  if (cost === 0) return 0;
  const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
  const firstYear = (cost * rate * months) / 12;
  const p = Math.trunc(period);
  if (p === 1) return firstYear;
  let total = firstYear;
  let dep = 0;
  for (let i = 2; i <= Math.min(p, Math.trunc(life)); i++) {
    dep = (cost - total) * rate;
    total += dep;
  }
  return p > life ? ((cost - total) * rate * (12 - months)) / 12 : dep;
});

/** DDB: double-declining (or factor) balance, never below the salvage value. */
fin('DDB', 'cost, salvage, life, period, [factor]', [N, N, N, N, O], (cost, salvage, life, period, factor = 2) => {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0 || period > life) return NUM;
  let rate = factor / life;
  let before;
  if (rate >= 1) {
    rate = 1;
    before = period === 1 ? cost : 0;
  } else {
    before = cost * (1 - rate) ** (period - 1);
  }
  const after = cost * (1 - rate) ** period;
  const d = after < salvage ? before - salvage : before - after;
  return d < 0 ? 0 : d;
});

// ---- rates ---------------------------------------------------------------------------------------------------------------------

fin('EFFECT', 'nominal_rate, npery', [N, N], (r, n) => {
  const k = Math.trunc(n);
  return r <= 0 || k < 1 ? NUM : (1 + r / k) ** k - 1;
});
fin('NOMINAL', 'effect_rate, npery', [N, N], (e, n) => {
  const k = Math.trunc(n);
  return e <= 0 || k < 1 ? NUM : k * ((1 + e) ** (1 / k) - 1);
});
fin('PDURATION', 'rate, pv, fv', [N, N, N], (r, p, f) => (r <= 0 || p <= 0 || f <= 0 ? NUM : (Math.log(f) - Math.log(p)) / Math.log1p(r)));
fin('RRI', 'nper, pv, fv', [N, N, N], (n, p, f) => (n <= 0 || p === 0 ? NUM : (f / p) ** (1 / n) - 1));

define('FVSCHEDULE', {
  min: 2, max: 2, category: 'financial', args: 'principal, schedule',
  fn: ([principal, schedule], ctx) => {
    let v = optNumber(principal, ctx, 0);
    if (isErr(v)) return v;
    for (const x of flat(schedule, ctx)) {
      if (x === null) continue;
      const r = typeof x === 'number' || isErr(x) ? x : toNumber(x);
      if (isErr(r)) return r === x ? r : VALUE;
      v *= 1 + r;
    }
    return v;
  },
});

/** DOLLARDE/DOLLARFR: dollar prices as fractions (1.02 with fraction 16 = 1 2/16 = 1.125) and back. */
function dollar(toDecimal) {
  return (x, f) => {
    const k = Math.trunc(f);
    if (k < 0) return NUM;
    if (k === 0) return DIV0;
    const whole = Math.trunc(x);
    const scale = 10 ** Math.ceil(Math.log10(k));
    return snap(whole + (toDecimal ? ((x - whole) * scale) / k : ((x - whole) * k) / scale));
  };
}
fin('DOLLARDE', 'fractional_dollar, fraction', [N, N], dollar(true));
fin('DOLLARFR', 'decimal_dollar, fraction', [N, N], dollar(false));
