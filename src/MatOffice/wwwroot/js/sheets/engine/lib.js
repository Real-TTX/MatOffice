/* MatOffice Sheets – the registry of worksheet functions and the helpers to write them (pure, no DOM).

   A function gets its arguments evaluated: a scalar (number, string, boolean, null = blank/omitted, FormulaError), a 2-D array (array of rows) or a
   Ref (a reference; ctx.values(ref) gives its values). Lazy functions (IF, IFERROR, CHOOSE …) get thunks () => value instead.
   It returns a scalar, a 2-D array (spills into the cells around) or a Ref. A number that is not finite becomes #NUM!.

   ctx (the evaluation context): sheet/row/col (the formula's cell), values(ref) → 2-D array (huge whole-column references are cut at the used range),
   value(sheet, row, col), sheetId(name), sheetName(id), sheetIndex(id), sheetCount(), now() → Date, random(), parseRef(text) → Ref | error,
   formulaAt(sheet, row, col) → text | null, rowHidden(sheet, row). */
import { Errors, FormulaError, Ref, isError, toNumber, toText, toBool } from './values.js';

/** name → { name, min, max, category, args, fn, lazy, volatile, refs } */
export const FUNCTIONS = new Map();

/** The categories of the function library (Formulas tab, Insert Function). */
export const CATEGORIES = ['financial', 'logical', 'text', 'date', 'lookup', 'math', 'stat', 'info'];

/**
 * Adds a function. spec: { min, max (Infinity for variadic, ≤ 255 arguments), category, args: 'number, [num_digits]' (the syntax for the help),
 * fn(args, ctx), lazy (thunks instead of values), volatile (recalculated on every change: NOW, RAND, OFFSET, INDIRECT) }.
 */
export function define(name, spec) {
  FUNCTIONS.set(name, { name, min: 0, max: 255, lazy: false, volatile: false, ...spec, max: Math.min(255, spec.max ?? 255) });
}

/** A Ref of one cell → its value, a larger Ref → its 2-D array; anything else stays. */
export function deref(v, ctx) {
  if (v instanceof Ref) return v.isCell ? ctx.value(v.sheet, v.r1, v.c1) : ctx.values(v);
  return v;
}

/** Any argument as a 2-D array (a scalar is [[v]]). */
export function matrix(v, ctx) {
  if (v instanceof Ref) return v.isCell ? [[ctx.value(v.sheet, v.r1, v.c1)]] : ctx.values(v);
  if (Array.isArray(v)) return v;
  return [[v]];
}

/** Every scalar of an argument: visit(value, fromRange) – fromRange is true for values of a reference or an array. */
export function each(v, ctx, visit) {
  if (v instanceof Ref || Array.isArray(v)) {
    const m = matrix(v, ctx);
    for (const row of m) for (const x of row) if (visit(x, true) === false) return;
  } else {
    visit(v, false);
  }
}

/** The first error among values (scalars), or null. */
export function firstError(...values) {
  for (const v of values) if (v instanceof FormulaError) return v;
  return null;
}

/**
 * The numbers of the arguments the way SUM/AVERAGE/MIN/MAX count them: from references and arrays only numbers (text, logicals and blanks are
 * skipped); typed arguments: numbers, logicals (1/0), text that is a number (else #VALUE!), an omitted argument is 0. Errors are returned.
 * mode 'a' (AVERAGEA, MAXA …): in references text counts 0 and logicals 1/0.
 * Returns an array of numbers or a FormulaError.
 */
export function numbers(args, ctx, mode = '') {
  const out = [];
  let error = null;
  for (const a of args) {
    each(a, ctx, (v, fromRange) => {
      if (v instanceof FormulaError) { error = v; return false; }
      if (typeof v === 'number') { out.push(v); return true; }
      if (fromRange) {
        if (mode === 'a') {
          if (typeof v === 'boolean') out.push(v ? 1 : 0);
          else if (typeof v === 'string') out.push(0);
        }
        return true;
      }
      if (v === null) { out.push(0); return true; }
      const n = toNumber(v);
      if (n instanceof FormulaError) { error = n; return false; }
      out.push(n);
      return true;
    });
    if (error) return error;
  }
  return out;
}

const parseType = t => ({ base: t.replace(/^\.\.\.|\?$/g, ''), optional: t.endsWith('?'), rest: t.startsWith('...') });

/**
 * A function of scalars: impl(...values) gets its arguments converted by types ('number', 'int' (truncated), 'text', 'bool', 'any' (as it is);
 * 'number?' = optional (undefined when it is not given), '...number' = all further arguments). An error among converted arguments is the result.
 * An array or a range for a scalar argument gives an array of results (like Excel 365: =LEN(A1:A3) spills).
 * Type 'raw' passes the argument without lifting (a range for a lookup, an array).
 */
export function scalar(types, impl) {
  const specs = types.map(parseType);
  const typeAt = i => (i < specs.length ? specs[i] : specs[specs.length - 1]?.rest ? specs[specs.length - 1] : { base: 'any' });
  return (args, ctx) => {
    const vals = args.map((a, i) => (typeAt(i).base === 'raw' ? a : deref(a, ctx)));
    let rows = 0;
    let cols = 0;
    vals.forEach((v, i) => {
      if (Array.isArray(v) && typeAt(i).base !== 'raw') { rows = Math.max(rows, v.length); cols = Math.max(cols, v[0]?.length || 0); }
    });
    if (!rows) return call(vals);
    const out = [];
    for (let r = 0; r < rows; r++) {
      const line = [];
      for (let c = 0; c < cols; c++) line.push(call(vals.map((v, i) => (Array.isArray(v) && typeAt(i).base !== 'raw' ? pick(v, r, c) : v))));
      out.push(line);
    }
    return out;
  };
  function call(vals) {
    const converted = [];
    for (let i = 0; i < vals.length; i++) {
      const v = convert(vals[i], typeAt(i).base);
      if (v instanceof FormulaError && typeAt(i).base !== 'any' && typeAt(i).base !== 'raw') return v;
      converted.push(v);
    }
    const r = impl(...converted);
    return typeof r === 'number' && !Number.isFinite(r) ? Errors.NUM : r;
  }
}

/** The element (r, c) of an array broadcast to a larger one (a single row/column repeats; beyond its size #N/A). */
export function pick(m, r, c) {
  const rr = m.length === 1 ? 0 : r;
  const cc = m[0].length === 1 ? 0 : c;
  return rr < m.length && cc < m[0].length ? m[rr][cc] : Errors.NA;
}

function convert(v, type) {
  if (Array.isArray(v) && type !== 'raw') v = v[0]?.[0] ?? null;
  switch (type) {
    case 'number': return toNumber(v);
    case 'int': { const n = toNumber(v); return n instanceof FormulaError ? n : Math.trunc(n); }
    case 'text': return toText(v);
    case 'bool': return toBool(v);
    default: return v;
  }
}

/** A 2-D array of the given size filled by fn(r, c). */
export function makeMatrix(rows, cols, fn) {
  const out = [];
  for (let r = 0; r < rows; r++) {
    const line = [];
    for (let c = 0; c < cols; c++) line.push(fn(r, c));
    out.push(line);
  }
  return out;
}

/** The values of an argument as one flat list (row by row). */
export function flat(v, ctx) {
  const out = [];
  each(v, ctx, x => { out.push(x); });
  return out;
}

/** Whether a value is an error (re-exported for the function modules). */
export { isError, Errors, Ref };
