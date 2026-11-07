/* MatOffice Base – expressions of Access: the tokenizer and parser (shared with the SQL parser), the formatter back to text, and the compiler that
   turns an expression into a function of a record. Operators: ^ - * / \ Mod + - & = <> < > <= >= Like Between In Is [Not] Null Not And Or Xor Eqv Imp;
   literals: 12.5, "text", 'text', #2026-11-07#, #11/7/2026#, True/False/Yes/No, Null; names: Field, [Field name], Table.Field, [Table]![Field].
   Text comparisons ignore case (like Access). Null spreads through arithmetic and comparisons; & treats Null as "".
   A pure module without the DOM (node tests: tests/js/base). Errors are BaseError with a code the editor turns into a text. */
import {
  DateVal, serialOf, partsOf, serialOfIso, parseDate, parseNumber, nowSerial, todaySerial, formatValue, formatNumber, formatDate,
  daysInMonth, monthName, dayName, roundHalfEven, generalNumber,
} from './values.js';

export class BaseError extends Error {
  constructor(code, ...args) {
    super(code + (args.length ? ': ' + args.join(' | ') : ''));
    this.code = code;
    this.args = args;
  }
}

// ---- tokens --------------------------------------------------------------------------------------------------------------------

const WORD_START = /[A-Za-z_À-ɏͰ-Ͽ]/;
const WORD = /[A-Za-z0-9_À-ɏͰ-Ͽ]/;

/** Splits text into tokens: { k: num|str|date|word|name|op|end, v, pos, end }. A "word" is unbracketed (may be a keyword), a "name" was in [brackets]. */
export function tokenize(text) {
  const src = String(text ?? '');
  const tokens = [];
  let i = 0;
  const push = (k, v, start) => tokens.push({ k, v, pos: start, end: i });
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    const start = i;
    if (c === '[') {
      const close = src.indexOf(']', i + 1);
      if (close < 0) throw new BaseError('Syntax', src.slice(i, i + 20), i);
      i = close + 1;
      push('name', src.slice(start + 1, close), start);
      continue;
    }
    if (c === '"' || c === "'") {
      let s = '';
      i++;
      for (;;) {
        if (i >= src.length) throw new BaseError('Syntax', src.slice(start, start + 20), start);
        if (src[i] === c) {
          if (src[i + 1] === c) { s += c; i += 2; continue; }
          i++;
          break;
        }
        s += src[i++];
      }
      push('str', s, start);
      continue;
    }
    if (c === '#') {
      const close = src.indexOf('#', i + 1);
      const inner = close > 0 ? src.slice(i + 1, close) : '';
      if (close > 0 && /^[\d\s/.:\-apmAPM]+$/.test(inner) && /\d/.test(inner)) {
        const serial = parseLiteralDate(inner);
        if (serial === null) throw new BaseError('BadDate', inner);
        i = close + 1;
        push('date', serial, start);
        continue;
      }
      i++;
      push('op', '#', start);
      continue;
    }
    const prev = tokens[tokens.length - 1];
    const afterValue = prev && (prev.k === 'name' || prev.k === 'word' || prev.k === 'num' || prev.k === 'str' || (prev.k === 'op' && prev.v === ')'));
    if (/\d/.test(c) || (c === '.' && /\d/.test(src[i + 1] || '') && !afterValue)) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
      i += m[0].length;
      push('num', Number(m[0]), start);
      continue;
    }
    if (WORD_START.test(c)) {
      while (i < src.length && WORD.test(src[i])) i++;
      push('word', src.slice(start, i), start);
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>' || two === '!=') { i += 2; push('op', two === '!=' ? '<>' : two, start); continue; }
    if ('=<>+-*/\\^&(),.!;:?%'.includes(c)) { i++; push('op', c, start); continue; }
    throw new BaseError('Syntax', src.slice(i, i + 20), i);
  }
  tokens.push({ k: 'end', v: '', pos: src.length, end: src.length });
  return tokens;
}

/** A date literal between # (SQL is always US: m/d/yyyy; ISO and the German d.m.yyyy are read too). */
export function parseLiteralDate(inner) {
  const s = inner.trim();
  if (/^\d{1,2}\.\d{1,2}\.\d{2,4}/.test(s)) return parseDate(s, 'de');
  return parseDate(s, 'en');
}

// ---- keywords ------------------------------------------------------------------------------------------------------------------

export const RESERVED = new Set([
  'AND', 'OR', 'NOT', 'XOR', 'EQV', 'IMP', 'MOD', 'LIKE', 'BETWEEN', 'IN', 'IS', 'NULL', 'TRUE', 'FALSE', 'YES', 'NO',
  'SELECT', 'FROM', 'WHERE', 'GROUP', 'BY', 'HAVING', 'ORDER', 'ASC', 'DESC', 'AS', 'INNER', 'LEFT', 'RIGHT', 'OUTER', 'JOIN', 'ON',
  'DISTINCT', 'DISTINCTROW', 'TOP', 'PERCENT', 'INTO', 'INSERT', 'VALUES', 'UPDATE', 'SET', 'DELETE', 'PARAMETERS', 'UNION', 'ALL',
  'EXISTS', 'CROSS', 'FULL', 'TRANSFORM', 'PIVOT', 'CREATE', 'DROP', 'ALTER', 'TABLE',
]);

/** A name as Access writes it: plain when it is a simple word, else in [brackets]. */
export function quoteName(name, always = false) {
  if (!always && /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !RESERVED.has(name.toUpperCase()) && !FUNCTIONS[name.toUpperCase()]) return name;
  return '[' + name + ']';
}

// ---- parser --------------------------------------------------------------------------------------------------------------------

/**
 * Expression nodes:
 *   { t: 'num'|'str'|'bool', v }  { t: 'date', v: serial }  { t: 'null' }
 *   { t: 'id', parts: ['Table', 'Field'] }  { t: 'star', table: 'Table'|null }  { t: 'param', name } (only from the criteria of the grid)
 *   { t: 'un', op: '-'|'+'|'NOT', e }  { t: 'bin', op, l, r }  (op: ^ * / \ MOD + - & = <> < > <= >= AND OR XOR EQV IMP)
 *   { t: 'like', e, p, not }  { t: 'between', e, lo, hi, not }  { t: 'in', e, list, not }  { t: 'isnull', e, not }
 *   { t: 'fn', name: 'IIf', args: [...] }  (name in its canonical spelling)
 */
export class Parser {
  constructor(text) {
    this.text = String(text ?? '');
    this.tokens = tokenize(this.text);
    this.i = 0;
  }

  get tok() { return this.tokens[this.i]; }
  peek(n = 1) { return this.tokens[Math.min(this.i + n, this.tokens.length - 1)]; }
  next() { const t = this.tokens[this.i]; if (this.i < this.tokens.length - 1) this.i++; return t; }

  /** Whether the current token is the unbracketed keyword (one or several words). */
  isKw(...words) {
    for (let n = 0; n < words.length; n++) {
      const t = this.peek(n);
      if (t.k !== 'word' || t.v.toUpperCase() !== words[n]) return false;
    }
    return true;
  }

  acceptKw(...words) {
    if (!this.isKw(...words)) return false;
    this.i += words.length;
    return true;
  }

  expectKw(word) {
    if (!this.acceptKw(word)) this.fail();
  }

  isOp(v) { return this.tok.k === 'op' && this.tok.v === v; }
  acceptOp(v) { if (this.isOp(v)) { this.next(); return true; } return false; }
  expectOp(v) { if (!this.acceptOp(v)) this.fail(); }

  fail(code = 'Syntax') {
    const t = this.tok;
    throw new BaseError(code, t.k === 'end' ? this.text.slice(-20) : this.text.slice(t.pos, t.pos + 24), t.pos);
  }

  atEnd() { return this.tok.k === 'end'; }

  /** A whole expression (nothing may follow). */
  static expression(text) {
    const p = new Parser(text);
    if (p.atEnd()) return null;
    const e = p.expr();
    if (!p.atEnd()) p.fail();
    return e;
  }

  expr() { return this.imp(); }

  imp() { let l = this.eqv(); while (this.acceptKw('IMP')) l = { t: 'bin', op: 'IMP', l, r: this.eqv() }; return l; }
  eqv() { let l = this.xor(); while (this.acceptKw('EQV')) l = { t: 'bin', op: 'EQV', l, r: this.xor() }; return l; }
  xor() { let l = this.or(); while (this.acceptKw('XOR')) l = { t: 'bin', op: 'XOR', l, r: this.or() }; return l; }
  or() { let l = this.and(); while (this.acceptKw('OR')) l = { t: 'bin', op: 'OR', l, r: this.and() }; return l; }
  and() { let l = this.not(); while (this.acceptKw('AND')) l = { t: 'bin', op: 'AND', l, r: this.not() }; return l; }
  not() { if (this.acceptKw('NOT')) return { t: 'un', op: 'NOT', e: this.not() }; return this.compare(); }

  compare() {
    let l = this.concat();
    for (;;) {
      const t = this.tok;
      if (t.k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) {
        this.next();
        l = { t: 'bin', op: t.v, l, r: this.concat() };
        continue;
      }
      const not = this.isKw('NOT') && (this.isKw('NOT', 'LIKE') || this.isKw('NOT', 'BETWEEN') || this.isKw('NOT', 'IN'));
      if (not) this.next();
      if (this.acceptKw('LIKE')) { l = { t: 'like', e: l, p: this.concat(), not }; continue; }
      if (this.acceptKw('BETWEEN')) {
        const lo = this.concat();
        this.expectKw('AND');
        l = { t: 'between', e: l, lo, hi: this.concat(), not };
        continue;
      }
      if (this.acceptKw('IN')) { l = { t: 'in', e: l, list: this.list(), not }; continue; }
      if (not) this.fail();
      if (this.acceptKw('IS')) {
        const isNot = this.acceptKw('NOT');
        this.expectKw('NULL');
        l = { t: 'isnull', e: l, not: isNot };
        continue;
      }
      return l;
    }
  }

  list() {
    this.expectOp('(');
    const items = [];
    if (this.isKw('SELECT')) this.fail('Unsupported');
    if (!this.isOp(')')) {
      do { items.push(this.expr()); } while (this.acceptOp(','));
    }
    this.expectOp(')');
    return items;
  }

  concat() { let l = this.add(); while (this.isOp('&')) { this.next(); l = { t: 'bin', op: '&', l, r: this.add() }; } return l; }

  add() {
    let l = this.mod();
    while (this.isOp('+') || this.isOp('-')) { const op = this.next().v; l = { t: 'bin', op, l, r: this.mod() }; }
    return l;
  }

  mod() { let l = this.idiv(); while (this.acceptKw('MOD')) l = { t: 'bin', op: 'MOD', l, r: this.idiv() }; return l; }
  idiv() { let l = this.mul(); while (this.isOp('\\')) { this.next(); l = { t: 'bin', op: '\\', l, r: this.mul() }; } return l; }

  mul() {
    let l = this.unary();
    while (this.isOp('*') || this.isOp('/')) { const op = this.next().v; l = { t: 'bin', op, l, r: this.unary() }; }
    return l;
  }

  unary() {
    if (this.isOp('-') || this.isOp('+')) {
      const op = this.next().v;
      const e = this.unary();
      if (op === '-' && e.t === 'num') return { t: 'num', v: -e.v };
      return op === '-' ? { t: 'un', op: '-', e } : e;
    }
    return this.power();
  }

  power() {
    let l = this.primary();
    while (this.isOp('^')) { this.next(); l = { t: 'bin', op: '^', l, r: this.unaryPow() }; }
    return l;
  }

  unaryPow() {
    if (this.isOp('-')) { this.next(); return { t: 'un', op: '-', e: this.primary() }; }
    return this.primary();
  }

  primary() {
    const t = this.tok;
    if (t.k === 'num') { this.next(); return { t: 'num', v: t.v }; }
    if (t.k === 'str') { this.next(); return { t: 'str', v: t.v }; }
    if (t.k === 'date') { this.next(); return { t: 'date', v: t.v }; }
    if (t.k === 'op' && t.v === '(') {
      this.next();
      if (this.isKw('SELECT')) this.fail('Unsupported');
      const e = this.expr();
      this.expectOp(')');
      return e;
    }
    if (t.k === 'word') {
      const up = t.v.toUpperCase();
      if (up === 'NULL') { this.next(); return { t: 'null' }; }
      if (up === 'TRUE' || up === 'YES') { this.next(); return { t: 'bool', v: true }; }
      if (up === 'FALSE' || up === 'NO') { this.next(); return { t: 'bool', v: false }; }
      if (this.peek().k === 'op' && this.peek().v === '(' && (FUNCTIONS[up] || !RESERVED.has(up))) return this.call();
      if (RESERVED.has(up) && !['ON'].includes(up)) this.fail();
      return this.ident();
    }
    if (t.k === 'name') return this.ident();
    this.fail();
    return null;
  }

  ident() {
    const parts = [this.next().v];
    while ((this.isOp('.') || this.isOp('!')) && (this.peek().k === 'name' || this.peek().k === 'word' || (this.peek().k === 'op' && this.peek().v === '*'))) {
      this.next();
      if (this.isOp('*')) { this.next(); return { t: 'star', table: parts.join('.') }; }
      parts.push(this.next().v);
    }
    if (parts.length > 2) {
      // Forms!Customers!Name and the like: keep the last two parts
      return { t: 'id', parts: parts.slice(-2) };
    }
    return { t: 'id', parts };
  }

  call() {
    const raw = this.next().v;
    const name = canonicalFunction(raw);
    this.expectOp('(');
    const args = [];
    if (!this.isOp(')')) {
      do {
        if (this.isOp('*') && (this.peek().k === 'op' && this.peek().v === ')')) { this.next(); args.push({ t: 'star', table: null }); }
        else if (this.isKw('DISTINCT')) { this.fail('Unsupported'); }
        else args.push(this.expr());
      } while (this.acceptOp(','));
    }
    this.expectOp(')');
    return { t: 'fn', name, args };
  }
}

// ---- formatting back to text ---------------------------------------------------------------------------------------------------

const PREC = { IMP: 1, EQV: 2, XOR: 3, OR: 4, AND: 5, NOT: 6, '=': 7, '<>': 7, '<': 7, '>': 7, '<=': 7, '>=': 7, LIKE: 7, BETWEEN: 7, IN: 7, IS: 7, '&': 8, '+': 9, '-': 9, MOD: 10, '\\': 11, '*': 12, '/': 12, NEG: 13, '^': 14 };

function precOf(e) {
  switch (e.t) {
    case 'bin': return PREC[e.op];
    case 'un': return e.op === 'NOT' ? PREC.NOT : PREC.NEG;
    case 'like': case 'between': case 'in': case 'isnull': return 7;
    case 'num': return e.v < 0 ? PREC.NEG : 99;
    default: return 99;
  }
}

/** The date of a literal as SQL writes it: #m/d/yyyy# (with the time when there is one). */
export function dateLiteral(serial) {
  const [y, m, d, h, n, s] = partsOf(serial);
  const date = `${m}/${d}/${y}`;
  const time = h || n || s ? `${h}:${String(n).padStart(2, '0')}:${String(s).padStart(2, '0')}` : '';
  if (Math.floor(serial) === 0 && time) return '#' + time + '#';
  return '#' + date + (time ? ' ' + time : '') + '#';
}

export function numberLiteral(n) {
  if (Number.isInteger(n)) return String(n);
  return String(Number(n.toPrecision(15)));
}

/**
 * An expression as text. options.qualify: 'keep' (as parsed) | function(parts) → text for names (the SQL writer brackets them its way).
 * Unqualified names are written in [brackets] like Access does in expressions.
 */
export function exprToText(e, options = {}) {
  const name = options.name || (parts => (parts.length === 1 ? '[' + parts[0] + ']' : parts.map(p => quoteName(p)).join('.')));
  const go = (x, parentPrec = 0, right = false) => {
    const s = write(x);
    const p = precOf(x);
    return p < parentPrec || (right && p === parentPrec && p < 99) ? '(' + s + ')' : s;
  };
  const sp = options.compact ? '' : ' ';
  const write = x => {
    switch (x.t) {
      case 'num': return numberLiteral(x.v);
      case 'str': return '"' + x.v.replace(/"/g, '""') + '"';
      case 'date': return dateLiteral(x.v);
      case 'bool': return x.v ? 'True' : 'False';
      case 'null': return 'Null';
      case 'id': return name(x.parts);
      case 'param': return '[' + x.name + ']';
      case 'star': return x.table ? quoteName(x.table) + '.*' : '*';
      case 'un': return x.op === 'NOT' ? 'Not ' + go(x.e, PREC.NOT) : '-' + go(x.e, PREC.NEG);
      case 'bin': {
        const p = PREC[x.op];
        const word = /^[A-Z]+$/.test(x.op);
        const op = x.op === 'MOD' ? ' Mod ' : word ? ' ' + (options.compact ? x.op : x.op[0] + x.op.slice(1).toLowerCase()) + ' ' : x.op === '&' ? ' & ' : sp + x.op + sp;
        return go(x.l, p) + op + go(x.r, p, true);
      }
      case 'like': return go(x.e, 8) + (x.not ? ' Not Like ' : ' Like ') + go(x.p, 8);
      case 'between': return go(x.e, 8) + (x.not ? ' Not Between ' : ' Between ') + go(x.lo, 8) + ' And ' + go(x.hi, 8);
      case 'in': return go(x.e, 8) + (x.not ? ' Not In (' : ' In (') + x.list.map(i => go(i)).join(',') + ')';
      case 'isnull': return go(x.e, 8) + (x.not ? ' Is Not Null' : ' Is Null');
      case 'fn': return x.name + '(' + x.args.map(a => go(a)).join(',') + ')';
      default: return '?';
    }
  };
  return write(e);
}

/** Whether two expressions are the same (names compared without case). */
export function sameExpr(a, b) {
  if (!a || !b) return a === b;
  return JSON.stringify(normalizeForCompare(a)) === JSON.stringify(normalizeForCompare(b));
}

function normalizeForCompare(e) {
  if (Array.isArray(e)) return e.map(normalizeForCompare);
  if (e && typeof e === 'object') {
    const out = {};
    for (const [k, v] of Object.entries(e)) out[k] = k === 'parts' ? v.map(p => p.toLowerCase()) : k === 'table' && typeof v === 'string' ? v.toLowerCase() : normalizeForCompare(v);
    return out;
  }
  return e;
}

/** Walks an expression (visit(node) for every node). */
export function walk(e, visit) {
  if (!e || typeof e !== 'object') return;
  visit(e);
  switch (e.t) {
    case 'un': walk(e.e, visit); break;
    case 'bin': walk(e.l, visit); walk(e.r, visit); break;
    case 'like': walk(e.e, visit); walk(e.p, visit); break;
    case 'between': walk(e.e, visit); walk(e.lo, visit); walk(e.hi, visit); break;
    case 'in': walk(e.e, visit); e.list.forEach(x => walk(x, visit)); break;
    case 'isnull': walk(e.e, visit); break;
    case 'fn': e.args.forEach(x => walk(x, visit)); break;
    default: break;
  }
}

/** A copy of an expression with nodes replaced (map(node) returns a node or undefined to keep going). */
export function mapExpr(e, map) {
  if (!e || typeof e !== 'object') return e;
  const r = map(e);
  if (r !== undefined) return r;
  switch (e.t) {
    case 'un': return { ...e, e: mapExpr(e.e, map) };
    case 'bin': return { ...e, l: mapExpr(e.l, map), r: mapExpr(e.r, map) };
    case 'like': return { ...e, e: mapExpr(e.e, map), p: mapExpr(e.p, map) };
    case 'between': return { ...e, e: mapExpr(e.e, map), lo: mapExpr(e.lo, map), hi: mapExpr(e.hi, map) };
    case 'in': return { ...e, e: mapExpr(e.e, map), list: e.list.map(x => mapExpr(x, map)) };
    case 'isnull': return { ...e, e: mapExpr(e.e, map) };
    case 'fn': return { ...e, args: e.args.map(x => mapExpr(x, map)) };
    default: return { ...e };
  }
}

export const AGGREGATES = new Set(['Sum', 'Avg', 'Count', 'Min', 'Max', 'First', 'Last', 'StDev', 'StDevP', 'Var', 'VarP']);

export function hasAggregate(e) {
  let found = false;
  walk(e, n => { if (n.t === 'fn' && AGGREGATES.has(n.name)) found = true; });
  return found;
}

// ---- values and operators ------------------------------------------------------------------------------------------------------

const collators = new Map();
export function collator(lang = 'en') {
  let c = collators.get(lang);
  if (!c) {
    c = new Intl.Collator(lang === 'de' ? 'de' : 'en', { sensitivity: 'accent', numeric: false });
    collators.set(lang, c);
  }
  return c;
}

export const isNull = v => v === null || v === undefined;

/** A value as a number for arithmetic (Yes = -1, a date is its serial, "" is 0); throws on text that is no number. */
export function toNumber(v, lang = 'en') {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? -1 : 0;
  if (v instanceof DateVal) return v.s;
  if (typeof v === 'string') {
    if (v.trim() === '') return 0;
    const n = parseNumber(v, 'en');
    if (n === null) {
      const n2 = parseNumber(v, lang);
      if (n2 === null) throw new BaseError('TypeMismatch', v);
      return n2;
    }
    return n;
  }
  throw new BaseError('TypeMismatch', String(v));
}

/** A value as text (& and the string functions). */
export function toText(v, lang = 'en') {
  if (isNull(v)) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return generalNumber(v, lang);
  if (typeof v === 'boolean') return v ? (lang === 'de' ? 'Wahr' : 'True') : lang === 'de' ? 'Falsch' : 'False';
  if (v instanceof DateVal) return formatDate(v, lang, '');
  return String(v);
}

export function toBool(v) {
  if (isNull(v)) return null;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof DateVal) return v.s !== 0;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    if (s === 'true' || s === 'yes' || s === 'wahr' || s === 'ja') return true;
    if (s === 'false' || s === 'no' || s === 'falsch' || s === 'nein' || s === '') return false;
    const n = parseNumber(v);
    if (n !== null) return n !== 0;
    throw new BaseError('TypeMismatch', v);
  }
  return !!v;
}

export function toDate(v, lang = 'en') {
  if (isNull(v)) return null;
  if (v instanceof DateVal) return v;
  if (typeof v === 'number') return new DateVal(v);
  if (typeof v === 'string') {
    const s = serialOfIso(v) ?? parseDate(v, lang) ?? parseDate(v, lang === 'de' ? 'en' : 'de');
    if (s === null) throw new BaseError('TypeMismatch', v);
    return new DateVal(s);
  }
  if (typeof v === 'boolean') return new DateVal(v ? -1 : 0);
  throw new BaseError('TypeMismatch', String(v));
}

/** Compares two values like Access: -1, 0, 1, or null when one is Null. Text without case. */
export function compare(a, b, lang = 'en') {
  if (isNull(a) || isNull(b)) return null;
  if (typeof a === 'string' && typeof b === 'string') {
    const r = collator(lang).compare(a, b);
    return r < 0 ? -1 : r > 0 ? 1 : 0;
  }
  if (a instanceof DateVal || b instanceof DateVal) {
    let x;
    let y;
    try { x = toDate(a, lang).s; y = toDate(b, lang).s; } catch { return compare(toText(a, lang), toText(b, lang), lang); }
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === 'string' || typeof b === 'string') {
    const s = typeof a === 'string' ? a : b;
    const n = parseNumber(s, 'en') ?? parseNumber(s, lang);
    if (n === null) return compare(toText(a, lang), toText(b, lang), lang);
    const x = typeof a === 'string' ? n : toNumber(a);
    const y = typeof b === 'string' ? n : toNumber(b);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = toNumber(a);
  const y = toNumber(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

/** Sorting: Null first (like Access), then by compare. */
export function sortCompare(a, b, lang = 'en') {
  const an = isNull(a);
  const bn = isNull(b);
  if (an || bn) return an && bn ? 0 : an ? -1 : 1;
  return compare(a, b, lang) ?? 0;
}

const likeCache = new Map();
/** The Like pattern of Access: * any text, ? one character, # one digit, [abc], [!abc], [a-z]; without case. */
export function likeRegex(pattern) {
  let re = likeCache.get(pattern);
  if (re) return re;
  let out = '^';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '*') out += '[\\s\\S]*';
    else if (c === '?') out += '[\\s\\S]';
    else if (c === '#') out += '\\d';
    else if (c === '[') {
      const close = pattern.indexOf(']', i + 1);
      if (close < 0) { out += '\\['; continue; }
      let set = pattern.slice(i + 1, close);
      i = close;
      if (set === '') { out += ''; continue; }
      let neg = false;
      if (set.startsWith('!')) { neg = true; set = set.slice(1); }
      out += '[' + (neg ? '^' : '') + set.replace(/[\\\]^]/g, '\\$&') + ']';
    } else out += c.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  re = new RegExp(out + '$', 'iu');
  if (likeCache.size > 500) likeCache.clear();
  likeCache.set(pattern, re);
  return re;
}

function arith(op, a, b, lang) {
  if (isNull(a) || isNull(b)) return null;
  if (op === '+' && typeof a === 'string' && typeof b === 'string') return a + b;
  const da = a instanceof DateVal;
  const db = b instanceof DateVal;
  const x = toNumber(a, lang);
  const y = toNumber(b, lang);
  switch (op) {
    case '+': return da || db ? new DateVal(x + y) : x + y;
    case '-': return da && !db ? new DateVal(x - y) : x - y;
    case '*': return x * y;
    case '/': if (y === 0) throw new BaseError('DivZero'); return x / y;
    case '\\': { const yy = roundHalfEven(y); if (yy === 0) throw new BaseError('DivZero'); return Math.trunc(roundHalfEven(x) / yy); }
    case 'MOD': { const yy = roundHalfEven(y); if (yy === 0) throw new BaseError('DivZero'); return roundHalfEven(x) % yy; }
    case '^': return x ** y;
    default: return null;
  }
}

function logic(op, a, b) {
  const x = isNull(a) ? null : toBool(a);
  const y = isNull(b) ? null : toBool(b);
  switch (op) {
    case 'AND': if (x === false || y === false) return false; if (x === null || y === null) return null; return true;
    case 'OR': if (x === true || y === true) return true; if (x === null || y === null) return null; return false;
    case 'XOR': if (x === null || y === null) return null; return x !== y;
    case 'EQV': if (x === null || y === null) return null; return x === y;
    case 'IMP': if (x === false || y === true) return true; if (x === null || y === null) return null; return false;
    default: return null;
  }
}

// ---- compiling -----------------------------------------------------------------------------------------------------------------

/**
 * Turns an expression into a function of an environment: fn(env) → value.
 * scope: {
 *   resolve(parts) → (env) => value | null      names of fields (null: unknown → a parameter, when scope.params knows it)
 *   params: { lower-case name: value }          answers of [Enter a city] prompts
 *   aggregate(name, argAst) → (env) => value     Sum/Count/... over the rows of a group (not given: an aggregate is an error)
 *   domain(name, exprText, domainName, criteria) → value   DLookup, DSum, DCount ...
 *   lang, now (Date), currency
 * }
 */
export function compile(e, scope = {}) {
  const lang = scope.lang || 'en';
  const c = x => compile(x, scope);
  switch (e?.t) {
    case undefined: return () => null;
    case 'num': case 'str': case 'bool': { const v = e.v; return () => v; }
    case 'date': { const v = new DateVal(e.v); return () => v; }
    case 'null': return () => null;
    case 'param': {
      const key = e.name.toLowerCase();
      return () => { if (scope.params && key in scope.params) return scope.params[key]; throw new BaseError('ParamMissing', e.name); };
    }
    case 'id': {
      const get = scope.resolve?.(e.parts);
      if (get) return get;
      const key = e.parts.join('.').toLowerCase();
      if (scope.params && key in scope.params) { const v = scope.params[key]; return () => v; }
      if (e.parts.length === 1 && /^(page|pages)$/i.test(e.parts[0]) && scope.page) { const which = e.parts[0].toLowerCase(); return () => scope.page[which]; }
      throw new BaseError('UnknownName', e.parts.join('.'));
    }
    case 'star': throw new BaseError('Syntax', '*');
    case 'un': {
      const f = c(e.e);
      if (e.op === 'NOT') return env => { const v = f(env); return isNull(v) ? null : !toBool(v); };
      return env => { const v = f(env); if (isNull(v)) return null; return v instanceof DateVal ? new DateVal(-v.s) : -toNumber(v, lang); };
    }
    case 'bin': {
      const l = c(e.l);
      const r = c(e.r);
      const op = e.op;
      if (op === '&') return env => toText(l(env), lang) + toText(r(env), lang);
      if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
        return env => {
          const k = compare(l(env), r(env), lang);
          if (k === null) return null;
          switch (op) { case '=': return k === 0; case '<>': return k !== 0; case '<': return k < 0; case '>': return k > 0; case '<=': return k <= 0; default: return k >= 0; }
        };
      }
      if (op === 'AND') return env => { const a = l(env); if (a === false || (a !== null && toBool(a) === false)) return false; return logic(op, a, r(env)); };
      if (op === 'OR') return env => { const a = l(env); if (a === true || (a !== null && !isNull(a) && toBool(a) === true)) return true; return logic(op, a, r(env)); };
      if (['XOR', 'EQV', 'IMP'].includes(op)) return env => logic(op, l(env), r(env));
      return env => arith(op, l(env), r(env), lang);
    }
    case 'like': {
      const v = c(e.e);
      const p = c(e.p);
      return env => {
        const a = v(env);
        const b = p(env);
        if (isNull(a) || isNull(b)) return null;
        const hit = likeRegex(toText(b, lang)).test(toText(a, lang));
        return e.not ? !hit : hit;
      };
    }
    case 'between': {
      const v = c(e.e);
      const lo = c(e.lo);
      const hi = c(e.hi);
      return env => {
        const a = v(env);
        const k1 = compare(a, lo(env), lang);
        const k2 = compare(a, hi(env), lang);
        if (k1 === null || k2 === null) return null;
        const hit = k1 >= 0 && k2 <= 0;
        return e.not ? !hit : hit;
      };
    }
    case 'in': {
      const v = c(e.e);
      const items = e.list.map(c);
      return env => {
        const a = v(env);
        if (isNull(a)) return null;
        const hit = items.some(f => compare(a, f(env), lang) === 0);
        return e.not ? !hit : hit;
      };
    }
    case 'isnull': {
      const v = c(e.e);
      return env => { const n = isNull(v(env)); return e.not ? !n : n; };
    }
    case 'fn': return compileCall(e, scope);
    default: throw new BaseError('Syntax', e.t);
  }
}

function compileCall(e, scope) {
  const lang = scope.lang || 'en';
  if (AGGREGATES.has(e.name)) {
    if (!scope.aggregate) throw new BaseError('AggregateHere', e.name);
    if (e.args.length !== 1) throw new BaseError('ArgCount', e.name);
    return scope.aggregate(e.name, e.args[0]);
  }
  const spec = FUNCTIONS[e.name.toUpperCase()];
  if (!spec) throw new BaseError('UnknownFunction', e.name);
  if (e.args.length < spec.min || e.args.length > spec.max) throw new BaseError('ArgCount', e.name);
  if (spec.domain) {
    const args = e.args.map(a => compile(a, scope));
    return env => {
      if (!scope.domain) throw new BaseError('UnknownFunction', e.name);
      const vals = args.map(f => f(env));
      return scope.domain(spec.domain, toText(vals[0], lang), toText(vals[1], lang), isNull(vals[2]) ? '' : toText(vals[2], lang));
    };
  }
  if (spec.lazy) return spec.lazy(e.args.map(a => compile(a, scope)), scope);
  const args = e.args.map(a => compile(a, scope));
  const fn = spec.fn;
  const ctx = { lang, now: scope.now, currency: scope.currency || 'EUR' };
  return env => fn(ctx, ...args.map(f => f(env)));
}

/** Evaluates an expression text once (default values, validation rules, calculated controls). */
export function evaluate(text, scope = {}, env = null) {
  const ast = typeof text === 'string' ? Parser.expression(text) : text;
  if (!ast) return null;
  return compile(ast, scope)(env);
}

// ---- functions -----------------------------------------------------------------------------------------------------------------

const nullIn = (...xs) => xs.some(isNull);
const num = (v, ctx) => toNumber(v, ctx?.lang);
const txt = (v, ctx) => toText(v, ctx?.lang);
const intArg = v => Math.trunc(toNumber(v));

const now = ctx => nowSerial(ctx.now ? ctx.now() : new Date());

function dateParts(v, ctx) { return partsOf(toDate(v, ctx.lang).s); }

function addInterval(interval, n, serial) {
  const [y, m, d, h, mi, s] = partsOf(serial);
  const time = serial - Math.floor(serial);
  const k = Math.trunc(n);
  switch (String(interval).toLowerCase()) {
    case 'yyyy': return addMonths(y, m, d, k * 12) + time;
    case 'q': return addMonths(y, m, d, k * 3) + time;
    case 'm': return addMonths(y, m, d, k) + time;
    case 'y': case 'd': case 'w': return serial + k;
    case 'ww': return serial + 7 * k;
    case 'h': return serialOf(y, m, d, h, mi, s) + k / 24;
    case 'n': return serial + k / 1440;
    case 's': return serial + k / 86400;
    default: throw new BaseError('BadArgument', interval);
  }
}

function addMonths(y, m, d, k) {
  const total = y * 12 + (m - 1) + k;
  const ny = Math.floor(total / 12);
  const nm = (total % 12) + 1;
  return serialOf(ny, nm, Math.min(d, daysInMonth(ny, nm)));
}

function diffInterval(interval, a, b) {
  const [y1, m1] = partsOf(a);
  const [y2, m2] = partsOf(b);
  switch (String(interval).toLowerCase()) {
    case 'yyyy': return y2 - y1;
    case 'q': return (y2 * 4 + Math.floor((m2 - 1) / 3)) - (y1 * 4 + Math.floor((m1 - 1) / 3));
    case 'm': return (y2 * 12 + m2) - (y1 * 12 + m1);
    case 'y': case 'd': return Math.floor(b) - Math.floor(a);
    case 'w': return Math.trunc((Math.floor(b) - Math.floor(a)) / 7);
    case 'ww': { const sa = Math.floor(a) - (partsOf(a)[6] - 1); const sb = Math.floor(b) - (partsOf(b)[6] - 1); return Math.round((sb - sa) / 7); }
    case 'h': return Math.floor(Math.round(b * 24 * 3600) / 3600) - Math.floor(Math.round(a * 24 * 3600) / 3600);
    case 'n': return Math.floor(Math.round(b * 86400) / 60) - Math.floor(Math.round(a * 86400) / 60);
    case 's': return Math.round(b * 86400) - Math.round(a * 86400);
    default: throw new BaseError('BadArgument', interval);
  }
}

function datePart(interval, serial) {
  const [y, m, d, h, n, s, wd] = partsOf(serial);
  switch (String(interval).toLowerCase()) {
    case 'yyyy': return y;
    case 'q': return Math.ceil(m / 3);
    case 'm': return m;
    case 'y': return Math.floor(serial) - serialOf(y, 1, 1) + 1;
    case 'd': return d;
    case 'w': return wd;
    case 'ww': return Math.ceil((Math.floor(serial) - serialOf(y, 1, 1) + partsOf(serialOf(y, 1, 1))[6]) / 7);
    case 'h': return h;
    case 'n': return n;
    case 's': return s;
    default: throw new BaseError('BadArgument', interval);
  }
}

function formatFn(ctx, v, fmt) {
  if (isNull(v)) return null;
  const f = isNull(fmt) ? '' : txt(fmt, ctx);
  if (!f) return toText(v, ctx.lang);
  const lower = f.toLowerCase();
  const numberNamed = /^(currency|euro|fixed|standard|percent|scientific|general number)$/.test(lower);
  const dateNamed = /^(general date|long date|medium date|short date|long time|medium time|short time)$/.test(lower);
  if (!numberNamed && (v instanceof DateVal || dateNamed || (/[dmyhnsq]/i.test(f.replace(/"[^"]*"/g, '')) && !/[0#]/.test(f)))) {
    let d;
    try { d = toDate(v, ctx.lang); } catch { return toText(v, ctx.lang); }
    return formatDate(d, ctx.lang, f);
  }
  if (/^(yes\/no|true\/false|on\/off)$/.test(lower)) return formatValue(toBool(v), { format: f, lang: ctx.lang });
  if (f === '>' || f === '<') return f === '>' ? txt(v, ctx).toUpperCase() : txt(v, ctx).toLowerCase();
  let n;
  try { n = toNumber(v, ctx.lang); } catch { return toText(v, ctx.lang); }
  return formatNumber(n, ctx.lang, f, null, ctx.currency);
}

function lazyIIf([c, a, b]) { return env => { const v = c(env); return !isNull(v) && toBool(v) ? a(env) : b(env); }; }
function lazySwitch(args) {
  return env => {
    for (let i = 0; i + 1 < args.length; i += 2) { const v = args[i](env); if (!isNull(v) && toBool(v)) return args[i + 1](env); }
    return null;
  };
}
function lazyChoose([i, ...items]) { return env => { const k = Math.trunc(toNumber(i(env))); return k >= 1 && k <= items.length ? items[k - 1](env) : null; }; }
function lazyNz([v, d]) { return env => { const x = v(env); return isNull(x) ? (d ? d(env) : '') : x; }; }

function round(n, digits = 0) { return roundHalfEven(n, digits); }

const F = (min, max, fn) => ({ min, max, fn });

/** The functions of expressions (Access/VBA names, upper case keys). */
export const FUNCTIONS = {
  IIF: { min: 3, max: 3, lazy: lazyIIf },
  SWITCH: { min: 2, max: 40, lazy: lazySwitch },
  CHOOSE: { min: 2, max: 40, lazy: lazyChoose },
  NZ: { min: 1, max: 2, lazy: lazyNz },
  ISNULL: F(1, 1, (ctx, v) => isNull(v)),
  ISNUMERIC: F(1, 1, (ctx, v) => typeof v === 'number' || (typeof v === 'string' && parseNumber(v, ctx.lang) !== null) || typeof v === 'boolean'),
  ISDATE: F(1, 1, (ctx, v) => { if (v instanceof DateVal) return true; if (typeof v !== 'string') return false; try { toDate(v, ctx.lang); return true; } catch { return false; } }),
  LEFT: F(2, 2, (ctx, s, n) => (isNull(s) ? null : txt(s, ctx).slice(0, Math.max(0, intArg(n))))),
  RIGHT: F(2, 2, (ctx, s, n) => { if (isNull(s)) return null; const t = txt(s, ctx); const k = Math.max(0, intArg(n)); return k === 0 ? '' : t.slice(-k); }),
  MID: F(2, 3, (ctx, s, start, len) => { if (isNull(s)) return null; const t = txt(s, ctx); const a = Math.max(1, intArg(start)) - 1; return isNull(len) || len === undefined ? t.slice(a) : t.substr(a, Math.max(0, intArg(len))); }),
  LEN: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).length)),
  UCASE: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).toUpperCase())),
  LCASE: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).toLowerCase())),
  TRIM: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).trim())),
  LTRIM: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).replace(/^\s+/, ''))),
  RTRIM: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).replace(/\s+$/, ''))),
  INSTR: F(2, 3, (ctx, a, b, c) => {
    let start = 1; let s1 = a; let s2 = b;
    if (c !== undefined) { start = intArg(a); s1 = b; s2 = c; }
    if (nullIn(s1, s2)) return null;
    return txt(s1, ctx).toLowerCase().indexOf(txt(s2, ctx).toLowerCase(), Math.max(0, start - 1)) + 1;
  }),
  REPLACE: F(3, 3, (ctx, s, find, repl) => {
    if (isNull(s)) return null;
    const f = txt(find, ctx);
    if (!f) return txt(s, ctx);
    return txt(s, ctx).replace(new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), () => txt(repl, ctx));
  }),
  SPACE: F(1, 1, (ctx, n) => ' '.repeat(Math.max(0, intArg(n)))),
  STRING: F(2, 2, (ctx, n, ch) => (isNull(ch) ? null : txt(ch, ctx).slice(0, 1).repeat(Math.max(0, intArg(n))))),
  STRREVERSE: F(1, 1, (ctx, s) => (isNull(s) ? null : [...txt(s, ctx)].reverse().join(''))),
  STRCOMP: F(2, 3, (ctx, a, b) => compare(txt(a, ctx), txt(b, ctx), ctx.lang)),
  ASC: F(1, 1, (ctx, s) => (isNull(s) ? null : txt(s, ctx).charCodeAt(0) || 0)),
  CHR: F(1, 1, (ctx, n) => String.fromCharCode(intArg(n))),
  FORMAT: F(1, 2, formatFn),
  FORMATCURRENCY: F(1, 2, (ctx, v, d) => (isNull(v) ? null : formatNumber(num(v, ctx), ctx.lang, 'Currency', d === undefined ? null : intArg(d), ctx.currency))),
  FORMATNUMBER: F(1, 2, (ctx, v, d) => (isNull(v) ? null : formatNumber(num(v, ctx), ctx.lang, 'Standard', d === undefined ? null : intArg(d)))),
  FORMATPERCENT: F(1, 2, (ctx, v, d) => (isNull(v) ? null : formatNumber(num(v, ctx), ctx.lang, 'Percent', d === undefined ? null : intArg(d)))),
  FORMATDATETIME: F(1, 2, (ctx, v, k) => (isNull(v) ? null : formatDate(toDate(v, ctx.lang), ctx.lang, ['General Date', 'Long Date', 'Short Date', 'Long Time', 'Short Time'][k === undefined ? 0 : intArg(k)] || ''))),
  DATE: F(0, 0, ctx => new DateVal(Math.floor(now(ctx)))),
  NOW: F(0, 0, ctx => new DateVal(now(ctx))),
  TIME: F(0, 0, ctx => { const s = now(ctx); return new DateVal(s - Math.floor(s)); }),
  YEAR: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[0])),
  MONTH: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[1])),
  DAY: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[2])),
  HOUR: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[3])),
  MINUTE: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[4])),
  SECOND: F(1, 1, (ctx, v) => (isNull(v) ? null : dateParts(v, ctx)[5])),
  WEEKDAY: F(1, 2, (ctx, v, first) => { if (isNull(v)) return null; const wd = dateParts(v, ctx)[6]; const f = first === undefined || isNull(first) ? 1 : intArg(first) || 1; return ((wd - f + 7) % 7) + 1; }),
  MONTHNAME: F(1, 2, (ctx, m, short) => (isNull(m) ? null : monthName(intArg(m), ctx.lang, !!(short && toBool(short))))),
  WEEKDAYNAME: F(1, 2, (ctx, d, short) => (isNull(d) ? null : dayName(intArg(d), ctx.lang, !!(short && toBool(short))))),
  DATEADD: F(3, 3, (ctx, i, n, d) => (nullIn(i, n, d) ? null : new DateVal(addInterval(txt(i, ctx), num(n, ctx), toDate(d, ctx.lang).s)))),
  DATEDIFF: F(3, 5, (ctx, i, a, b) => (nullIn(i, a, b) ? null : diffInterval(txt(i, ctx), toDate(a, ctx.lang).s, toDate(b, ctx.lang).s))),
  DATEPART: F(2, 4, (ctx, i, d) => (nullIn(i, d) ? null : datePart(txt(i, ctx), toDate(d, ctx.lang).s))),
  DATESERIAL: F(3, 3, (ctx, y, m, d) => { if (nullIn(y, m, d)) return null; const base = addMonths(intArg(y), 1, 1, intArg(m) - 1); return new DateVal(base + intArg(d) - 1); }),
  TIMESERIAL: F(3, 3, (ctx, h, m, s) => (nullIn(h, m, s) ? null : new DateVal((intArg(h) * 3600 + intArg(m) * 60 + intArg(s)) / 86400))),
  DATEVALUE: F(1, 1, (ctx, v) => (isNull(v) ? null : new DateVal(Math.floor(toDate(v, ctx.lang).s)))),
  TIMEVALUE: F(1, 1, (ctx, v) => { if (isNull(v)) return null; const s = toDate(v, ctx.lang).s; return new DateVal(s - Math.floor(s)); }),
  CDATE: F(1, 1, (ctx, v) => (isNull(v) ? null : toDate(v, ctx.lang))),
  CSTR: F(1, 1, (ctx, v) => (isNull(v) ? null : txt(v, ctx))),
  STR: F(1, 1, (ctx, v) => { if (isNull(v)) return null; const n = num(v, ctx); return (n >= 0 ? ' ' : '') + generalNumber(n, 'en'); }),
  VAL: F(1, 1, (ctx, v) => { if (isNull(v)) return 0; const m = /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(txt(v, ctx).replace(/\s+/g, '')); return m ? Number(m[0]) : 0; }),
  CINT: F(1, 1, (ctx, v) => (isNull(v) ? null : round(num(v, ctx)))),
  CLNG: F(1, 1, (ctx, v) => (isNull(v) ? null : round(num(v, ctx)))),
  CDBL: F(1, 1, (ctx, v) => (isNull(v) ? null : num(v, ctx))),
  CSNG: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.fround(num(v, ctx)))),
  CDEC: F(1, 1, (ctx, v) => (isNull(v) ? null : num(v, ctx))),
  CCUR: F(1, 1, (ctx, v) => (isNull(v) ? null : round(num(v, ctx), 4))),
  CBOOL: F(1, 1, (ctx, v) => (isNull(v) ? null : toBool(v))),
  ROUND: F(1, 2, (ctx, v, d) => (isNull(v) ? null : round(num(v, ctx), d === undefined || isNull(d) ? 0 : intArg(d)))),
  INT: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.floor(num(v, ctx)))),
  FIX: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.trunc(num(v, ctx)))),
  ABS: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.abs(num(v, ctx)))),
  SGN: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.sign(num(v, ctx)))),
  SQR: F(1, 1, (ctx, v) => { if (isNull(v)) return null; const n = num(v, ctx); if (n < 0) throw new BaseError('BadArgument', 'Sqr'); return Math.sqrt(n); }),
  EXP: F(1, 1, (ctx, v) => (isNull(v) ? null : Math.exp(num(v, ctx)))),
  LOG: F(1, 1, (ctx, v) => { if (isNull(v)) return null; const n = num(v, ctx); if (n <= 0) throw new BaseError('BadArgument', 'Log'); return Math.log(n); }),
  DLOOKUP: { min: 2, max: 3, domain: 'DLookup' },
  DCOUNT: { min: 2, max: 3, domain: 'DCount' },
  DSUM: { min: 2, max: 3, domain: 'DSum' },
  DAVG: { min: 2, max: 3, domain: 'DAvg' },
  DMIN: { min: 2, max: 3, domain: 'DMin' },
  DMAX: { min: 2, max: 3, domain: 'DMax' },
  DFIRST: { min: 2, max: 3, domain: 'DFirst' },
  DLAST: { min: 2, max: 3, domain: 'DLast' },
};

const CANONICAL = {};
for (const name of [
  'IIf', 'Switch', 'Choose', 'Nz', 'IsNull', 'IsNumeric', 'IsDate', 'Left', 'Right', 'Mid', 'Len', 'UCase', 'LCase', 'Trim', 'LTrim', 'RTrim', 'InStr', 'Replace', 'Space',
  'String', 'StrReverse', 'StrComp', 'Asc', 'Chr', 'Format', 'FormatCurrency', 'FormatNumber', 'FormatPercent', 'FormatDateTime', 'Date', 'Now', 'Time', 'Year', 'Month', 'Day',
  'Hour', 'Minute', 'Second', 'Weekday', 'MonthName', 'WeekdayName', 'DateAdd', 'DateDiff', 'DatePart', 'DateSerial', 'TimeSerial', 'DateValue', 'TimeValue', 'CDate', 'CStr',
  'Str', 'Val', 'CInt', 'CLng', 'CDbl', 'CSng', 'CDec', 'CCur', 'CBool', 'Round', 'Int', 'Fix', 'Abs', 'Sgn', 'Sqr', 'Exp', 'Log', 'DLookup', 'DCount', 'DSum', 'DAvg', 'DMin',
  'DMax', 'DFirst', 'DLast', ...AGGREGATES,
]) CANONICAL[name.toUpperCase()] = name;

export function canonicalFunction(name) { return CANONICAL[name.toUpperCase()] || name; }

/** The names of the functions (for the expression help and the property sheet). */
export const FUNCTION_NAMES = Object.values(CANONICAL).sort();

export { DateVal, formatValue };
