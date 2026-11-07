/* MatOffice Base – criteria the way Access reads them in the design grid of a query and in the validation rule of a field: the column is implied.
   ">10" → [col] > 10, "Between 1 And 5", "Like "A*"", "Is Null", "Not "Berlin"", "In("a","b")", ">10 And <20", ""Berlin" Or "Paris"".
   A bare word becomes text (Berlin → "Berlin"), a bare pattern becomes Like (A* → Like "A*"), a bare date gets its # (1.1.2026 → #1.1.2026#),
   like Access does when the cell is left. The writer turns such conditions back into the text of a cell. A pure module without the DOM. */
import { Parser, BaseError, exprToText, sameExpr } from './expr.js';

const COMPARE = ['=', '<>', '<', '>', '<=', '>='];

/** Puts quotes and # where Access would: bare words become text, bare dates become date literals, a bare pattern becomes Like. */
export function tidyCriteria(text) {
  let s = String(text ?? '').trim();
  if (!s) return '';
  const pattern = /^(not\s+)?([^\s"'[\]#()<>=,]+)$/i.exec(s);
  if (pattern && /[*?]/.test(pattern[2]) && !/^[\d.,]+\*[\d.,]+$/.test(pattern[2])) {
    return (pattern[1] ? 'Not ' : '') + 'Like "' + pattern[2] + '"';
  }
  const like = /^(not\s+)?like\s+([^\s"'[\]#()]+)$/i.exec(s);
  if (like) return (like[1] ? 'Not ' : '') + 'Like "' + like[2] + '"';
  // bare dates outside of quotes and #…#
  s = outsideQuotes(s, part => part.replace(/(^|[^#"\d./\w])(\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{4}-\d{2}-\d{2})(?=$|[^#\d./\w])/g, (m, pre, d) => pre + '#' + d + '#'));
  return s;
}

function outsideQuotes(s, fn) {
  let out = '';
  let buf = '';
  let quote = null;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quote) {
      buf += c;
      if (c === quote) { out += buf; buf = ''; quote = null; }
    } else if (c === '"' || c === '#' || c === '[') {
      out += fn(buf);
      buf = c;
      quote = c === '[' ? ']' : c;
    } else buf += c;
  }
  return out + (quote ? buf : fn(buf));
}

class CriteriaParser extends Parser {
  constructor(text, target) {
    super(text);
    this.target = target;
  }

  criteria() {
    const e = this.orC();
    if (!this.atEnd()) this.fail();
    return e;
  }

  orC() { let l = this.andC(); while (this.acceptKw('OR')) l = { t: 'bin', op: 'OR', l, r: this.andC() }; return l; }
  andC() { let l = this.notC(); while (this.acceptKw('AND')) l = { t: 'bin', op: 'AND', l, r: this.notC() }; return l; }

  notC() {
    if (this.isKw('NOT') && !this.isKw('NOT', 'LIKE') && !this.isKw('NOT', 'BETWEEN') && !this.isKw('NOT', 'IN')) {
      this.next();
      return { t: 'un', op: 'NOT', e: this.notC() };
    }
    return this.atom();
  }

  atom() {
    const col = this.target;
    const t = this.tok;
    if (t.k === 'op' && COMPARE.includes(t.v)) { this.next(); return { t: 'bin', op: t.v, l: col, r: this.value() }; }
    const not = this.acceptKw('NOT');
    if (this.acceptKw('LIKE')) return { t: 'like', e: col, p: this.value(), not };
    if (this.acceptKw('BETWEEN')) {
      const lo = this.value();
      this.expectKw('AND');
      return { t: 'between', e: col, lo, hi: this.value(), not };
    }
    if (this.acceptKw('IN')) return { t: 'in', e: col, list: this.list(), not };
    if (not) this.fail();
    if (this.acceptKw('IS')) {
      const isNot = this.acceptKw('NOT');
      this.expectKw('NULL');
      return { t: 'isnull', e: col, not: isNot };
    }
    if (this.isOp('(')) {
      // a group of criteria: (>1 And <5) Or 7
      const save = this.i;
      this.next();
      try {
        const inner = this.orC();
        this.expectOp(')');
        return inner;
      } catch {
        this.i = save;
      }
    }
    const v = this.value();
    const op = this.tok;
    if (op.k === 'op' && COMPARE.includes(op.v)) {
      // a whole condition of its own ([a] > [b])
      this.next();
      return { t: 'bin', op: op.v, l: v, r: this.value() };
    }
    return { t: 'bin', op: '=', l: col, r: v };
  }

  /** A value: an expression without And/Or at its top; a bare word is text. */
  value() {
    const t = this.tok;
    const n = this.peek();
    if (t.k === 'word' && !['NULL', 'TRUE', 'FALSE', 'YES', 'NO', 'NOT', 'DATE', 'NOW'].includes(t.v.toUpperCase())
      && !(n.k === 'op' && (n.v === '(' || n.v === '.' || n.v === '!'))
      && !['AND', 'OR', 'LIKE', 'BETWEEN', 'IN', 'IS'].includes(t.v.toUpperCase())) {
      // Berlin → "Berlin" (several bare words: New York → "New York")
      const words = [this.next().v];
      while (this.tok.k === 'word' && !['AND', 'OR', 'NOT', 'LIKE', 'BETWEEN', 'IN', 'IS'].includes(this.tok.v.toUpperCase())) words.push(this.next().v);
      return { t: 'str', v: words.join(' ') };
    }
    return this.concat();
  }
}

/**
 * The condition of a criteria cell for a column (target = the expression of the column). Answers null for an empty cell.
 * Names in [brackets] that are no field of the query are parameters (asked when the query runs).
 */
export function parseCriteria(text, target) {
  const s = tidyCriteria(text);
  if (!s) return null;
  return new CriteriaParser(s, target).criteria();
}

/** The column a condition is about (its left side), or null when it is not a condition on one column. */
export function targetOf(e) {
  switch (e?.t) {
    case 'bin':
      if (COMPARE.includes(e.op)) return isConstant(e.l) ? null : (isConstant(e.r) || !refersTo(e.r, e.l) ? e.l : null);
      if (e.op === 'AND' || e.op === 'OR') {
        const a = targetOf(e.l);
        const b = targetOf(e.r);
        return a && b && sameExpr(a, b) ? a : null;
      }
      return null;
    case 'like': case 'between': case 'in': case 'isnull': return isConstant(e.e) ? null : e.e;
    case 'un': return e.op === 'NOT' ? targetOf(e.e) : null;
    default: return null;
  }
}

function isConstant(e) {
  let constant = true;
  const visit = x => {
    if (!x || typeof x !== 'object') return;
    if (x.t === 'id' || x.t === 'fn' && ['Sum', 'Avg', 'Count', 'Min', 'Max', 'First', 'Last', 'StDev', 'Var'].includes(x.name)) constant = false;
    for (const v of Object.values(x)) {
      if (Array.isArray(v)) v.forEach(visit); else if (v && typeof v === 'object') visit(v);
    }
  };
  visit(e);
  return constant;
}

function refersTo(e, target) {
  let found = false;
  const visit = x => {
    if (!x || typeof x !== 'object' || found) return;
    if (sameExpr(x, target)) { found = true; return; }
    for (const v of Object.values(x)) {
      if (Array.isArray(v)) v.forEach(visit); else if (v && typeof v === 'object') visit(v);
    }
  };
  visit(e);
  return found;
}

/** The text of a criteria cell for a condition on a column (the reverse of parseCriteria). */
export function criteriaText(e, target, options = {}) {
  const val = x => exprToText(x, options);
  const w = x => {
    switch (x.t) {
      case 'bin':
        if (x.op === 'AND' || x.op === 'OR') {
          const word = x.op === 'AND' ? ' And ' : ' Or ';
          const side = y => (y.t === 'bin' && (y.op === 'AND' || y.op === 'OR') && y.op !== x.op ? '(' + w(y) + ')' : w(y));
          return side(x.l) + word + side(x.r);
        }
        if (sameExpr(x.l, target)) return (x.op === '=' ? '' : x.op) + val(x.r);
        return val(x);
      case 'like': return (x.not ? 'Not Like ' : 'Like ') + val(x.p);
      case 'between': return (x.not ? 'Not Between ' : 'Between ') + val(x.lo) + ' And ' + val(x.hi);
      case 'in': return (x.not ? 'Not In (' : 'In (') + x.list.map(val).join(',') + ')';
      case 'isnull': return x.not ? 'Is Not Null' : 'Is Null';
      case 'un': if (x.op === 'NOT') return 'Not ' + w(x.e); return val(x);
      default: return val(x);
    }
  };
  return w(e);
}

/** Checks a validation rule (text) against a value: true, false, or throws BaseError for a rule that cannot be read. */
export function ruleHolds(rule, value, compileFn, scope = {}) {
  const target = { t: 'id', parts: ['__value'] };
  const e = parseCriteria(rule, target);
  if (!e) return true;
  const fn = compileFn(e, { ...scope, resolve: parts => (parts.length === 1 && parts[0] === '__value' ? () => value : scope.resolve?.(parts) || null) });
  const r = fn(null);
  // Access: a rule that gives Null (e.g. >0 for an empty field) is not broken
  return r === null || r === undefined ? true : !!r;
}

export { BaseError };
