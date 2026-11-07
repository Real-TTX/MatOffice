/* MatOffice Sheets – logical functions (pure, no DOM). IF, IFS, IFERROR, IFNA and SWITCH are lazy: only the arguments that are needed are
   calculated. An array as test (or as expression of SWITCH) gives an array of results: =IF(A1:A3>1, "x", "y") spills. */
import { define, scalar, deref, each, makeMatrix, pick, Errors } from '../lib.js';
import { FormulaError, toBool, compare } from '../values.js';

const isErr = v => v instanceof FormulaError;
const { VALUE, NA } = Errors;

/** A test value as a logical: an error stays, text other than TRUE/FALSE is #VALUE!. */
const truth = v => (isErr(v) ? v : toBool(v));
/** The value of a lazy argument that was given empty is 0 in the cell (=IF(TRUE,,1)). */
const valueOf = thunk => { const v = thunk(); return v === null ? 0 : v; };
/** The element (r, c) of an array (broadcast), or the scalar itself. */
const at = (v, r, c) => (Array.isArray(v) ? pick(v, r, c) : v);

/** The size of an array result: the largest of each dimension among the arrays. */
function sizeOf(values) {
  let rows = 1;
  let cols = 1;
  for (const v of values) if (Array.isArray(v)) { rows = Math.max(rows, v.length); cols = Math.max(cols, v[0].length); }
  return [rows, cols];
}

define('IF', {
  min: 1, max: 3, category: 'logical', lazy: true, args: 'logical_test, [value_if_true], [value_if_false]',
  fn: ([test, yes, no], ctx) => {
    const branch = b => (b ? (yes ? yes() : true) : no ? no() : false);
    const t = deref(test(), ctx);
    if (!Array.isArray(t)) {
      const b = truth(t);
      if (isErr(b)) return b;
      const r = branch(b);
      return r === null ? 0 : r;
    }
    // element by element: each branch is calculated once, and only when an element needs it
    const tests = t.map(row => row.map(truth));
    const need = new Set(tests.flat().filter(b => typeof b === 'boolean'));
    const values = { true: need.has(true) ? deref(branch(true), ctx) : null, false: need.has(false) ? deref(branch(false), ctx) : null };
    const [rows, cols] = sizeOf([t, values.true, values.false]);
    return makeMatrix(rows, cols, (r, c) => {
      const b = pick(tests, r, c);
      return isErr(b) ? b : at(values[b], r, c);
    });
  },
});

/** IFS(test1, value1, …): the value of the first test that is TRUE; #N/A when none is. */
define('IFS', {
  min: 2, max: 254, category: 'logical', lazy: true, args: 'logical_test1, value_if_true1, [logical_test2, value_if_true2], …',
  fn: (args, ctx) => {
    if (args.length % 2) return NA;
    for (let i = 0; i < args.length; i += 2) {
      const t = deref(args[i](), ctx);
      if (Array.isArray(t)) {
        // an array test: the remaining tests and values element by element
        const tests = [t];
        const vals = [];
        for (let k = i; k < args.length; k += 2) {
          if (k > i) tests.push(deref(args[k](), ctx));
          vals.push(deref(args[k + 1](), ctx));
        }
        const [rows, cols] = sizeOf([...tests, ...vals]);
        return makeMatrix(rows, cols, (r, c) => {
          for (let k = 0; k < tests.length; k++) {
            const b = truth(at(tests[k], r, c));
            if (isErr(b)) return b;
            if (b) return at(vals[k], r, c);
          }
          return NA;
        });
      }
      const b = truth(t);
      if (isErr(b)) return b;
      if (b) return valueOf(args[i + 1]);
    }
    return NA;
  },
});

/** IFERROR/IFNA: the value, or the fallback where it is an error (#N/A for IFNA); element by element for arrays. */
function fallback(hit) {
  return ([value, alt], ctx) => {
    const v = deref(value(), ctx);
    if (!Array.isArray(v)) return hit(v) ? valueOf(alt) : v ?? 0;
    if (!v.some(row => row.some(hit))) return v;
    const a = deref(alt(), ctx);
    return makeMatrix(v.length, v[0].length, (r, c) => (hit(v[r][c]) ? at(a, r, c) ?? 0 : v[r][c]));
  };
}

define('IFERROR', { min: 2, max: 2, category: 'logical', lazy: true, args: 'value, value_if_error', fn: fallback(isErr) });
define('IFNA', { min: 2, max: 2, category: 'logical', lazy: true, args: 'value, value_if_na', fn: fallback(v => isErr(v) && v.code === NA.code) });

/**
 * The logicals of AND/OR/XOR: from ranges and arrays the logicals and numbers (text and blank cells are skipped), typed arguments converted
 * ("TRUE" is TRUE, other text #VALUE!). An error is the result; nothing to test is #VALUE!.
 */
function logicals(args, ctx) {
  const out = [];
  let error = null;
  for (const a of args) {
    each(a, ctx, (v, fromRange) => {
      if (isErr(v)) { error = v; return false; }
      if (typeof v === 'boolean') out.push(v);
      else if (typeof v === 'number') out.push(v !== 0);
      else if (!fromRange) {
        const b = toBool(v);
        if (isErr(b)) { error = b; return false; }
        out.push(b);
      }
      return true;
    });
    if (error) return error;
  }
  return out.length ? out : VALUE;
}

const logical = (name, f) => define(name, {
  min: 1, max: 255, category: 'logical', args: 'logical1, [logical2], …',
  fn: (args, ctx) => { const list = logicals(args, ctx); return isErr(list) ? list : f(list); },
});
logical('AND', l => l.every(Boolean));
logical('OR', l => l.some(Boolean));
logical('XOR', l => l.filter(Boolean).length % 2 === 1);

define('NOT', { min: 1, max: 1, category: 'logical', args: 'logical', fn: scalar(['bool'], b => !b) });

const same = (a, b) => compare(a, b) === 0;

/** SWITCH(expression, value1, result1, …, [default]): the result of the first value equal to the expression; the default or #N/A. */
define('SWITCH', {
  min: 3, max: 254, category: 'logical', lazy: true, args: 'expression, value1, result1, [default_or_value2, result2], …',
  fn: ([expr, ...rest], ctx) => {
    const x = deref(expr(), ctx);
    const count = rest.length >> 1;
    const hasDefault = rest.length % 2 === 1;
    if (Array.isArray(x)) {
      const vals = [];
      const results = [];
      for (let i = 0; i < count; i++) { vals.push(deref(rest[2 * i](), ctx)); results.push(deref(rest[2 * i + 1](), ctx)); }
      const def = hasDefault ? deref(rest[rest.length - 1](), ctx) : NA;
      const [rows, cols] = sizeOf([x, ...vals, ...results, def]);
      return makeMatrix(rows, cols, (r, c) => {
        const e = at(x, r, c);
        if (isErr(e)) return e;
        for (let k = 0; k < count; k++) {
          const v = at(vals[k], r, c);
          if (isErr(v)) return v;
          if (same(e, v)) return at(results[k], r, c);
        }
        return at(def, r, c);
      });
    }
    if (isErr(x)) return x;
    for (let i = 0; i < count; i++) {
      const v = at(deref(rest[2 * i](), ctx), 0, 0);
      if (isErr(v)) return v;
      if (same(x, v)) return valueOf(rest[2 * i + 1]);
    }
    return hasDefault ? valueOf(rest[rest.length - 1]) : NA;
  },
});

define('TRUE', { min: 0, max: 0, category: 'logical', args: '', fn: () => true });
define('FALSE', { min: 0, max: 0, category: 'logical', args: '', fn: () => false });
