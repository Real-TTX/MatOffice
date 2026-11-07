/* MatOffice Sheets – the values of the formula engine (pure, no DOM).
   A value is a number, a string, a boolean, null (an empty cell) or a FormulaError; a 2-D array (array of rows) is an array value;
   a Ref is a reference to an area of a sheet (functions such as ROW, OFFSET, INDEX need the reference, the others its values). */
import { parseDate } from './dates.js';

export class FormulaError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}

const make = code => Object.freeze(new FormulaError(code));

/** The error values (one shared instance each). #CIRC! is the editor's own: a circular reference. */
export const Errors = Object.freeze({
  NULL: make('#NULL!'), DIV0: make('#DIV/0!'), VALUE: make('#VALUE!'), REF: make('#REF!'), NAME: make('#NAME?'),
  NUM: make('#NUM!'), NA: make('#N/A'), SPILL: make('#SPILL!'), CALC: make('#CALC!'), CIRC: make('#CIRC!'),
});

const BY_CODE = new Map(Object.values(Errors).map(e => [e.code, e]));

/** The error of a code ("#DIV/0!"), or null. */
export const errorOf = code => BY_CODE.get(String(code).toUpperCase()) || null;
export const isError = v => v instanceof FormulaError;

/** A reference to an area of a sheet (0-based, inclusive). sheet is the id of the sheet. */
export class Ref {
  constructor(sheet, r1, c1, r2 = r1, c2 = c1) {
    this.sheet = sheet;
    this.r1 = Math.min(r1, r2); this.r2 = Math.max(r1, r2);
    this.c1 = Math.min(c1, c2); this.c2 = Math.max(c1, c2);
  }
  get height() { return this.r2 - this.r1 + 1; }
  get width() { return this.c2 - this.c1 + 1; }
  get isCell() { return this.r1 === this.r2 && this.c1 === this.c2; }
}

export const isRef = v => v instanceof Ref;
export const isArray = v => Array.isArray(v);

// ---- the locale of the values (decimal comma in German) -----------------------------------------------------------------------

let locale = 'en';
/** "en" or "de": how numbers are written as text (TEXT, &, VALUE) and how typed text is read. */
export function setLocale(value) { locale = value === 'de' ? 'de' : 'en'; }
export const getLocale = () => locale;
export const decimalSep = () => (locale === 'de' ? ',' : '.');
export const groupSep = () => (locale === 'de' ? '.' : ',');

// ---- numbers as text -------------------------------------------------------------------------------------------------------------

/** A number as Excel writes it into text (up to 15 significant digits, exponent for very large or small numbers). */
export function numberToText(n) {
  if (!Number.isFinite(n)) return Errors.NUM.code;
  if (n === 0) return '0';
  const abs = Math.abs(n);
  let text;
  if (abs >= 1e15 || abs < 1e-9) {
    const [mantissa, exp] = n.toExponential(14).split('e');
    const m = mantissa.replace(/\.?0+$/, '');
    const e = Number(exp);
    text = m + 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(2, '0');
  } else {
    text = String(Number(n.toPrecision(15)));
    if (/e/i.test(text)) text = Number(n.toPrecision(15)).toFixed(20).replace(/\.?0+$/, '');
  }
  return locale === 'de' ? text.replace('.', ',') : text;
}

/**
 * Text that is a number for Excel ("1,234.5", "12%", "$5", "1e3", "-(4)", a date or a time) → the number; null otherwise.
 * The decimal and group separators follow the locale; "." is always taken as decimal point when it cannot be a group separator.
 */
export function parseNumber(text) {
  if (typeof text !== 'string') return null;
  let t = text.trim();
  if (!t) return null;
  let negative = false;
  if (/^\(.*\)$/.test(t)) { negative = true; t = t.slice(1, -1).trim(); }
  let percent = false;
  if (t.endsWith('%')) { percent = true; t = t.slice(0, -1).trim(); }
  t = t.replace(/^[$€£¥]\s*|\s*[$€£¥]$|\s*(EUR|USD)$/g, '').trim();
  if (/^[+-]/.test(t)) { if (t[0] === '-') negative = !negative; t = t.slice(1).trim(); }
  t = t.replace(/^[$€£¥]\s*/, '');
  let n = plainNumber(t);
  if (n === null && !percent) {
    const d = parseDate(t, locale);
    if (d) n = d.value;
  }
  if (n === null) return null;
  if (percent) n /= 100;
  return negative ? -n : n;
}

function plainNumber(t) {
  const dec = decimalSep();
  const grp = groupSep();
  if (/^\d+([eE][+-]?\d+)?$/.test(t)) return Number(t);
  // with the locale's separators: 1,234,567.89 (en) / 1.234.567,89 (de)
  const g = grp === ',' ? ',' : '\\.';
  const d = dec === '.' ? '\\.' : ',';
  const local = new RegExp(`^(\\d{1,3}(${g}\\d{3})+|\\d+)(${d}\\d*)?([eE][+-]?\\d+)?$|^${d}\\d+([eE][+-]?\\d+)?$`);
  if (local.test(t)) return Number(t.split(grp).join('').replace(dec, '.'));
  // a plain decimal point is understood in German too (12.5), unless it looks like a group (1.234)
  if (dec === ',' && /^\d*\.\d+([eE][+-]?\d+)?$/.test(t) && !/^\d{1,3}\.\d{3}$/.test(t)) return Number(t);
  return null;
}

// ---- coercion ----------------------------------------------------------------------------------------------------------------------

/** A value as a number for arithmetic: blank 0, TRUE 1, text that is a number; #VALUE! otherwise. Errors stay errors. */
export function toNumber(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    if (v === '') return Errors.VALUE;
    const n = parseNumber(v);
    return n === null ? Errors.VALUE : n;
  }
  if (v instanceof FormulaError) return v;
  return Errors.VALUE;
}

/** A value as text: blank "", numbers like Excel writes them, TRUE/FALSE. Errors stay errors. */
export function toText(v) {
  if (typeof v === 'string') return v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return numberToText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof FormulaError) return v;
  return Errors.VALUE;
}

/** A value as a logical: blank FALSE, numbers ≠ 0, the texts TRUE/FALSE (also WAHR/FALSCH); #VALUE! otherwise. */
export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (v === null || v === undefined) return false;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const u = v.trim().toUpperCase();
    if (u === 'TRUE' || u === 'WAHR') return true;
    if (u === 'FALSE' || u === 'FALSCH') return false;
    return Errors.VALUE;
  }
  if (v instanceof FormulaError) return v;
  return Errors.VALUE;
}

const collator = new Intl.Collator('en', { sensitivity: 'accent', numeric: false });

/** Text comparison without case (like Excel): -1, 0, 1. */
export const compareText = (a, b) => Math.sign(collator.compare(a, b));

const rank = v => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);

/**
 * Excel's comparison of two values (=, <, > …): numbers < text < logicals; text without case; a blank cell is 0, "" or FALSE
 * depending on the other side. Errors must be handled before.
 */
export function compare(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 0) return a < b ? -1 : a > b ? 1 : 0;
  if (ra === 1) return compareText(a, b);
  if (ra === 2) return a === b ? 0 : a ? 1 : -1;
  return 0;
}

/** The order of Data → Sort (ascending): numbers, text, logicals, errors, blanks last. */
export function sortCompare(a, b) {
  const blankA = a === null || a === undefined || a === '';
  const blankB = b === null || b === undefined || b === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  const ra = a instanceof FormulaError ? 3 : rank(a);
  const rb = b instanceof FormulaError ? 3 : rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b;
  if (ra === 1) return collator.compare(a, b);
  if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return 0;
}

// ---- criteria (COUNTIF, SUMIF, …) ------------------------------------------------------------------------------------------------

/** A wildcard pattern of Excel (* any text, ? one character, ~ escapes) as a RegExp (whole text, no case). */
export function wildcard(pattern) {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) { re += '\\' + pattern[++i]; continue; }
    if (c === '*') re += '[\\s\\S]*';
    else if (c === '?') re += '[\\s\\S]';
    else re += c.replace(/[.+^${}()|[\]\\/-]/g, '\\$&');
  }
  return new RegExp('^' + re + '$', 'i');
}

/**
 * The test of a criterion of COUNTIF & co: 5, "5", ">=10", "<>x", "a*", "=" (blank), "<>" (not blank), TRUE, "#N/A".
 * Returns a function value → boolean.
 */
export function criteria(crit) {
  if (crit === null || crit === undefined) return v => v === null || v === undefined || v === '';
  if (crit instanceof FormulaError) return v => v instanceof FormulaError && v.code === crit.code;
  if (typeof crit === 'number') return v => (typeof v === 'number' ? v === crit : typeof v === 'string' && v !== '' && parseNumber(v) === crit);
  if (typeof crit === 'boolean') return v => v === crit;
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(String(crit));
  const op = m[1] || '=';
  const operand = m[2];
  if (operand === '') {
    if (op === '=') return v => v === null || v === undefined || v === '';
    if (op === '<>') return v => !(v === null || v === undefined || v === '');
    return () => false;
  }
  const err = errorOf(operand);
  if (err) return op === '<>' ? v => !(v instanceof FormulaError && v.code === err.code) : op === '=' ? v => v instanceof FormulaError && v.code === err.code : () => false;
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    return test(op, v => (typeof v === 'boolean' ? (v === b ? 0 : v ? 1 : -1) : null), op === '<>');
  }
  const num = parseNumber(operand);
  if (num !== null) {
    return test(op, v => {
      if (typeof v === 'number') return v === num ? 0 : v < num ? -1 : 1;
      if (typeof v === 'string' && (op === '=' || op === '<>')) { const n = parseNumber(v); return n === null ? null : n === num ? 0 : 1; }
      return null;
    }, op === '<>');
  }
  if (op === '=' || op === '<>') {
    const re = /[*?~]/.test(operand) ? wildcard(operand) : null;
    const eq = v => {
      if (v === null || v === undefined) return false;
      const s = typeof v === 'string' ? v : typeof v === 'number' ? numberToText(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : null;
      if (s === null) return false;
      return re ? re.test(s) : compareText(s, operand) === 0;
    };
    return op === '=' ? eq : v => !eq(v);
  }
  return test(op, v => (typeof v === 'string' ? compareText(v, operand) : null), false);
}

function test(op, cmp, matchOthers) {
  return v => {
    const c = cmp(v);
    if (c === null) return matchOthers;
    switch (op) {
      case '=': return c === 0;
      case '<>': return c !== 0;
      case '<': return c < 0;
      case '<=': return c <= 0;
      case '>': return c > 0;
      default: return c >= 0;
    }
  };
}
