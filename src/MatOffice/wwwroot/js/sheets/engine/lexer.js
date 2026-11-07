/* MatOffice Sheets – the tokens of a formula (pure, no DOM).
   The tokens keep their text, so a formula can be written again with only some tokens changed (references after an inserted row, a renamed sheet,
   the separators of another language) and the spaces of the person stay. Formulas are stored canonical: English function names, "," between
   arguments, "." as decimal point; with locale "de" the person types and reads ";" and the decimal comma (both forms are accepted). */
import { colIndex, colName, sheetPrefix, MAX_ROWS, MAX_COLS } from './address.js';

const ERRORS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!', '#CIRC!', '#GETTING_DATA'];
const IDENT_START = /[A-Za-z_\\À-￿]/;
const IDENT = /[A-Za-z0-9_.\\À-￿]/;
const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})/;
const COLS_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
const ROWS_RE = /^(\$?)([1-9]\d{0,6}):(\$?)([1-9]\d{0,6})/;

/**
 * The tokens of a formula (without the leading "="). Types: ws, num (value, canon), str (value), bool (value), err (code),
 * ref (sheet, kind cell|area|cols|rows|bad, r1 c1 r2 c2, ar1 ac1 ar2 ac2 = absolute flags), func (name), name, op, ( ) sep { } arow acol, bad.
 */
export function tokenize(src, { locale = 'en' } = {}) {
  const de = locale === 'de';
  const tokens = [];
  let i = 0;
  let braces = 0;
  const n = src.length;
  const add = (type, start, extra) => { const t = { type, text: src.slice(start, i), pos: start }; if (extra) Object.assign(t, extra); tokens.push(t); return t; };

  while (i < n) {
    const c = src[i];
    const start = i;
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === ' ') {
      while (i < n && /[ \t\n\r ]/.test(src[i])) i++;
      add('ws', start);
      continue;
    }
    if (c === '"') {
      let value = '';
      i++;
      let closed = false;
      while (i < n) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') { value += '"'; i += 2; continue; }
          i++;
          closed = true;
          break;
        }
        value += src[i++];
      }
      add('str', start, { value, closed });
      continue;
    }
    if (c === '#') {
      const upper = src.slice(i).toUpperCase();
      const code = ERRORS.find(e => upper.startsWith(e));
      if (code) { i += code.length; add(code === '#REF!' ? 'ref' : 'err', start, code === '#REF!' ? { sheet: null, kind: 'bad' } : { code }); continue; }
      i++;
      add('bad', start);
      continue;
    }
    if (c === "'") {
      // a quoted sheet name: 'My sheet'!A1
      let name = '';
      let j = i + 1;
      let closed = false;
      while (j < n) {
        if (src[j] === "'") {
          if (src[j + 1] === "'") { name += "'"; j += 2; continue; }
          closed = true;
          j++;
          break;
        }
        name += src[j++];
      }
      if (closed && src[j] === '!') {
        i = j + 1;
        if (refAfterSheet(name, start)) continue;
      }
      i = closed ? j : n;
      add('bad', start);
      continue;
    }
    if (c === '{') { braces++; i++; add('{', start); continue; }
    if (c === '}') { braces = Math.max(0, braces - 1); i++; add('}', start); continue; }
    if (c === '(' || c === ')') { i++; add(c, start); continue; }
    if (c === ';') { i++; add(braces ? 'arow' : 'sep', start); continue; }
    if (c === ',') {
      // German ",5" where an operand starts is a number; after an operand the comma separates arguments (people paste English formulas)
      if (de && !braces && /\d/.test(src[i + 1] || '') && operandStarts() && number()) continue;
      i++;
      add(braces ? 'acol' : 'sep', start);
      continue;
    }
    if (c === '\\' && braces) { i++; add('acol', start); continue; }
    if (c === '.' && braces && de) { i++; add('acol', start); continue; }
    if (/[0-9.]/.test(c)) {
      if (rowsRef(null, start)) continue;
      if (number()) continue;
      i++;
      add('bad', start);
      continue;
    }
    if (c === '$') {
      if (plainRef(null, start)) continue;
      i++;
      add('bad', start);
      continue;
    }
    if (IDENT_START.test(c)) {
      if (plainRef(null, start)) continue;
      let j = i;
      while (j < n && IDENT.test(src[j])) j++;
      const word = src.slice(i, j);
      if (src[j] === '!') {
        i = j + 1;
        if (refAfterSheet(word, start)) continue;
        add('bad', start);
        continue;
      }
      i = j;
      if (src[j] === '(') {
        add('func', start, { name: canonicalFunction(word) });
        continue;
      }
      const upper = word.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') { add('bool', start, { value: upper === 'TRUE' }); continue; }
      if (de && (upper === 'WAHR' || upper === 'FALSCH')) { add('bool', start, { value: upper === 'WAHR' }); continue; }
      add('name', start, { name: word });
      continue;
    }
    // operators
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { i += 2; add('op', start, { op: two }); continue; }
    if ('+-*/^&=<>%:@'.includes(c)) { i++; add('op', start, { op: c }); continue; }
    i++;
    add('bad', start);
  }
  return tokens;

  function operandStarts() {
    for (let k = tokens.length - 1; k >= 0; k--) {
      if (tokens[k].type === 'ws') continue;
      return ['op', '(', 'sep', '{', 'arow', 'acol'].includes(tokens[k].type);
    }
    return true;
  }

  function number() {
    const rest = src.slice(i);
    const re = de ? (braces ? /^(\d+(?:,\d*)?|,\d+)([eE][+-]?\d+)?/ : /^(\d+(?:[,.]\d*)?|[,.]\d+)([eE][+-]?\d+)?/) : /^(\d+(?:\.\d*)?|\.\d+)([eE][+-]?\d+)?/;
    const m = re.exec(rest);
    if (!m || IDENT_START.test(rest[m[0].length] || '')) return false;
    const start = i;
    i += m[0].length;
    const canon = m[0].replace(',', '.');
    add('num', start, { value: Number(canon), canon });
    return true;
  }

  function after(j) { return src[j] || ''; }

  /** A reference without a sheet at i (A1, $A$1, A1:B2, A:C, 1:3). */
  function plainRef(sheet, start) {
    const rest = src.slice(i);
    let m = CELL_RE.exec(rest);
    if (m) {
      const col = colIndex(m[2]);
      const next = after(i + m[0].length);
      if (col >= 0 && Number(m[4]) <= MAX_ROWS && !IDENT.test(next) && next !== '(' && next !== '!') {
        let j = i + m[0].length;
        const a = { r: Number(m[4]) - 1, c: col, ar: m[3] === '$', ac: m[1] === '$' };
        if (src[j] === ':') {
          const m2 = CELL_RE.exec(src.slice(j + 1));
          const col2 = m2 ? colIndex(m2[2]) : -1;
          const next2 = m2 ? after(j + 1 + m2[0].length) : '';
          if (m2 && col2 >= 0 && Number(m2[4]) <= MAX_ROWS && !IDENT.test(next2) && next2 !== '(' && next2 !== '!') {
            i = j + 1 + m2[0].length;
            add('ref', start, { sheet, kind: 'area', r1: a.r, c1: a.c, r2: Number(m2[4]) - 1, c2: col2, ar1: a.ar, ac1: a.ac, ar2: m2[3] === '$', ac2: m2[1] === '$' });
            return true;
          }
        }
        i = j;
        add('ref', start, { sheet, kind: 'cell', r1: a.r, c1: a.c, r2: a.r, c2: a.c, ar1: a.ar, ac1: a.ac, ar2: a.ar, ac2: a.ac });
        return true;
      }
    }
    m = COLS_RE.exec(rest);
    if (m) {
      const c1 = colIndex(m[2]);
      const c2 = colIndex(m[4]);
      const next = after(i + m[0].length);
      if (c1 >= 0 && c2 >= 0 && !IDENT.test(next) && next !== '(') {
        i += m[0].length;
        add('ref', start, { sheet, kind: 'cols', r1: 0, c1, r2: MAX_ROWS - 1, c2, ar1: false, ac1: m[1] === '$', ar2: false, ac2: m[3] === '$' });
        return true;
      }
    }
    return rowsRef(sheet, start);
  }

  function rowsRef(sheet, start) {
    const m = ROWS_RE.exec(src.slice(i));
    if (!m) return false;
    const r1 = Number(m[2]);
    const r2 = Number(m[4]);
    const next = after(i + m[0].length);
    if (r1 > MAX_ROWS || r2 > MAX_ROWS || IDENT.test(next) || next === '(' || next === '.') return false;
    i += m[0].length;
    add('ref', start, { sheet, kind: 'rows', r1: r1 - 1, c1: 0, r2: r2 - 1, c2: MAX_COLS - 1, ar1: m[1] === '$', ac1: false, ar2: m[3] === '$', ac2: false });
    return true;
  }

  function refAfterSheet(sheet, start) {
    if (src.slice(i, i + 5).toUpperCase() === '#REF!') {
      i += 5;
      add('ref', start, { sheet, kind: 'bad' });
      return true;
    }
    return plainRef(sheet, start);
  }
}

/** The canonical name of a function: upper case, without the prefixes of newer Excel functions in files (_xlfn., _xlws.). */
export function canonicalFunction(word) {
  return word.replace(/^(_xlfn\.|_xlws\.|_xludf\.)+/i, '').toUpperCase();
}

const cellText = (r, c, ar, ac) => (ac ? '$' : '') + colName(c) + (ar ? '$' : '') + (r + 1);

/** The text of a reference token (after its fields were changed). */
export function refText(t) {
  const prefix = t.sheet !== null && t.sheet !== undefined ? sheetPrefix(t.sheet) : '';
  switch (t.kind) {
    case 'cell': return prefix + cellText(t.r1, t.c1, t.ar1, t.ac1);
    case 'area': return prefix + cellText(t.r1, t.c1, t.ar1, t.ac1) + ':' + cellText(t.r2, t.c2, t.ar2, t.ac2);
    case 'cols': return prefix + (t.ac1 ? '$' : '') + colName(t.c1) + ':' + (t.ac2 ? '$' : '') + colName(t.c2);
    case 'rows': return prefix + (t.ar1 ? '$' : '') + (t.r1 + 1) + ':' + (t.ar2 ? '$' : '') + (t.r2 + 1);
    default: return prefix + '#REF!';
  }
}

/**
 * The text of tokens in a locale: "en" is the canonical form (stored), "de" uses ";" between arguments, the decimal comma and "." between the
 * columns of an array constant. Changed reference tokens (dirty) are written from their fields.
 */
export function printTokens(tokens, locale = 'en') {
  const de = locale === 'de';
  let out = '';
  for (const t of tokens) {
    switch (t.type) {
      case 'num': out += de ? t.canon.replace('.', ',') : t.canon; break;
      case 'sep': out += de ? ';' : ','; break;
      case 'acol': out += de ? '.' : ','; break;
      case 'arow': out += ';'; break;
      case 'ref': out += t.dirty ? refText(t) : t.text; break;
      case 'func': out += t.text.replace(/^(_xlfn\.|_xlws\.|_xludf\.)+/i, ''); break;
      case 'bool': out += de ? t.text : (t.value ? 'TRUE' : 'FALSE'); break;
      default: out += t.text;
    }
  }
  return out;
}

/** A formula typed in a locale → the canonical text that is stored (without "="). */
export const toCanonical = (text, locale) => (locale === 'en' ? printTokens(tokenize(text, { locale: 'en' })) : printTokens(tokenize(text, { locale }), 'en'));

/** A stored (canonical) formula → the text shown in a locale (without "="). */
export const toLocal = (text, locale) => (locale === 'en' ? text : printTokens(tokenize(text, { locale: 'en' }), locale));
