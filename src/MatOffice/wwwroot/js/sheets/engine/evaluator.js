/* MatOffice Sheets – the evaluation of a syntax tree (pure, no DOM).
   Operators work element by element on arrays and ranges (a smaller array is broadcast, like Excel 365); the result of a formula may be an array
   that spills into the cells below and to the right. */
import { Errors, FormulaError, Ref, errorOf, compare, toNumber, toText } from './values.js';
import { FUNCTIONS } from './lib.js';

/** Evaluates a node: a scalar, a 2-D array or a Ref. */
export function evaluate(node, ctx) {
  switch (node.type) {
    case 'num': case 'str': case 'bool': return node.value;
    case 'err': return errorOf(node.code) || Errors.NAME;
    case 'empty': return null;
    case 'ref': {
      const sheet = node.sheet === null ? ctx.sheet : ctx.sheetId(node.sheet);
      return sheet === null || sheet === undefined ? Errors.REF : new Ref(sheet, node.r1, node.c1, node.r2, node.c2);
    }
    case 'name': return ctx.name ? ctx.name(node.name) : Errors.NAME;
    case 'array': return node.rows.map(r => r.map(n => (n.type === 'err' ? errorOf(n.code) || Errors.NA : n.value)));
    case 'at': return implicit(evaluate(node.arg, ctx), ctx);
    case 'unary': {
      const v = evaluate(node.arg, ctx);
      if (node.op === '+') return v;
      return map1(v, ctx, x => { const n = toNumber(x); return n instanceof FormulaError ? n : n === 0 ? 0 : -n; });
    }
    case 'percent': return map1(evaluate(node.arg, ctx), ctx, x => { const n = toNumber(x); return n instanceof FormulaError ? n : n / 100; });
    case 'binary': {
      if (node.op === ':') return rangeOf(evaluate(node.left, ctx), evaluate(node.right, ctx));
      const a = evaluate(node.left, ctx);
      const b = evaluate(node.right, ctx);
      return map2(a, b, ctx, (x, y) => operate(node.op, x, y));
    }
    case 'fn': return call(node, ctx);
    default: return Errors.VALUE;
  }
}

function call(node, ctx) {
  const spec = FUNCTIONS.get(node.name);
  if (!spec) return Errors.NAME;
  const n = node.args.length;
  if (n < spec.min || n > spec.max) return Errors.VALUE;
  let result;
  try {
    result = spec.lazy ? spec.fn(node.args.map(a => () => evaluate(a, ctx)), ctx) : spec.fn(node.args.map(a => evaluate(a, ctx)), ctx);
  } catch (e) {
    if (e instanceof RangeError) throw e; // too deep: the engine handles it
    return Errors.VALUE;
  }
  if (result === undefined) return Errors.VALUE;
  if (typeof result === 'number' && !Number.isFinite(result)) return Errors.NUM;
  return result;
}

/** A value (or a reference) as a scalar or a 2-D array. */
function plain(v, ctx) {
  if (v instanceof Ref) return v.isCell ? ctx.value(v.sheet, v.r1, v.c1) : ctx.values(v);
  return v;
}

function map1(v, ctx, f) {
  const p = plain(v, ctx);
  if (!Array.isArray(p)) return f(p);
  return p.map(row => row.map(f));
}

function map2(a, b, ctx, f) {
  const x = plain(a, ctx);
  const y = plain(b, ctx);
  const ax = Array.isArray(x);
  const ay = Array.isArray(y);
  if (!ax && !ay) return f(x, y);
  const rows = Math.max(ax ? x.length : 1, ay ? y.length : 1);
  const cols = Math.max(ax ? x[0].length : 1, ay ? y[0].length : 1);
  const at = (m, isArr, r, c) => {
    if (!isArr) return m;
    const rr = m.length === 1 ? 0 : r;
    const cc = m[0].length === 1 ? 0 : c;
    return rr < m.length && cc < m[0].length ? m[rr][cc] : Errors.NA;
  };
  const out = [];
  for (let r = 0; r < rows; r++) {
    const line = [];
    for (let c = 0; c < cols; c++) line.push(f(at(x, ax, r, c), at(y, ay, r, c)));
    out.push(line);
  }
  return out;
}

const round15 = n => (n === 0 ? 0 : Number(n.toPrecision(15)));

/** One operator on two scalars. */
export function operate(op, x, y) {
  if (x instanceof FormulaError) return x;
  if (y instanceof FormulaError) return y;
  switch (op) {
    case '&': {
      const a = toText(x);
      const b = toText(y);
      return a instanceof FormulaError ? a : b instanceof FormulaError ? b : a + b;
    }
    case '=': case '<>': case '<': case '>': case '<=': case '>=': {
      const c = typeof x === 'number' && typeof y === 'number' ? Math.sign(round15(x) - round15(y)) : compare(x, y);
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        default: return c >= 0;
      }
    }
    default: {
      const a = toNumber(x);
      if (a instanceof FormulaError) return a;
      const b = toNumber(y);
      if (b instanceof FormulaError) return b;
      let r;
      switch (op) {
        case '+': r = a + b; break;
        case '-': r = a - b; break;
        case '*': r = a * b; break;
        case '/': if (b === 0) return Errors.DIV0; r = a / b; break;
        case '^':
          if (a === 0 && b === 0) return Errors.NUM;
          if (a === 0 && b < 0) return Errors.DIV0;
          r = Math.pow(a, b);
          break;
        default: return Errors.VALUE;
      }
      // the sum of numbers that cancel out (0.1 + 0.2 - 0.3) is 0 for Excel
      if ((op === '+' || op === '-') && r !== 0 && Math.abs(r) < Math.max(Math.abs(a), Math.abs(b)) * 1e-15) r = 0;
      return Number.isFinite(r) ? r : Errors.NUM;
    }
  }
}

/** A:B with references on both sides: the area around both (on one sheet). */
function rangeOf(a, b) {
  if (a instanceof FormulaError) return a;
  if (b instanceof FormulaError) return b;
  if (!(a instanceof Ref) || !(b instanceof Ref) || a.sheet !== b.sheet) return Errors.VALUE;
  return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
}

/** @A1:A9 – the cell of the range in the formula's row (or column); the first element of an array. */
export function implicit(v, ctx) {
  if (v instanceof Ref) {
    if (v.isCell) return ctx.value(v.sheet, v.r1, v.c1);
    if (v.c1 === v.c2 && ctx.row >= v.r1 && ctx.row <= v.r2) return ctx.value(v.sheet, ctx.row, v.c1);
    if (v.r1 === v.r2 && ctx.col >= v.c1 && ctx.col <= v.c2) return ctx.value(v.sheet, v.r1, ctx.col);
    return Errors.VALUE;
  }
  if (Array.isArray(v)) return v[0][0];
  return v;
}

/**
 * The value a formula leaves in its cell: a scalar (a blank becomes 0, like =A1 of an empty A1) or a 2-D array to spill.
 */
export function result(v, ctx) {
  let p = plain(v, ctx);
  if (Array.isArray(p)) {
    if (p.length === 1 && p[0].length === 1) p = p[0][0];
    else return p.map(row => row.map(x => (x === null || x === undefined ? 0 : Array.isArray(x) ? Errors.VALUE : x)));
  }
  if (p === null || p === undefined) return 0;
  if (typeof p === 'number') return p === 0 ? 0 : p;
  return p;
}
