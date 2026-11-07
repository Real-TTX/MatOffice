/* MatOffice Sheets – text functions (pure, no DOM). TEXT and DOLLAR live with the number formats. Texts are UTF-16 like in Excel
   (LEN counts code units); a result longer than 32767 characters is #VALUE!. */
import { define, scalar, deref, each, flat, Errors } from '../lib.js';
import { FormulaError, toText, toBool, wildcard, parseNumber, decimalSep, groupSep } from '../values.js';
import { optNumber } from './stat.js';
import { roundTo } from './math.js';

const isErr = v => v instanceof FormulaError;
const { VALUE, NA, CALC } = Errors;
const fitting = s => (s.length > 32767 ? VALUE : s);
const first = v => (Array.isArray(v) ? v[0][0] : v);

/** A text function of scalars (lifted over arrays): types as for scalar(); optional ones end with '?', '...text' takes the rest. */
const text = (name, args, types, f) => define(name, {
  min: types.filter(t => !t.endsWith('?') && !t.startsWith('...')).length || (types[0]?.startsWith('...') ? 1 : 0),
  max: types.some(t => t.startsWith('...')) ? 255 : types.length,
  category: 'text', args, fn: scalar(types, f),
});

/** The texts of an argument, all of them (ranges row by row); an error is returned instead. */
function texts(v, ctx) {
  const out = [];
  for (const x of flat(v, ctx)) {
    const t = toText(x);
    if (isErr(t)) return t;
    out.push(t);
  }
  return out;
}

// ---- joining -------------------------------------------------------------------------------------------------------------------

define('CONCAT', {
  min: 1, max: 255, category: 'text', args: 'text1, [text2], …',
  fn: (args, ctx) => {
    let out = '';
    for (const a of args) {
      const list = texts(a, ctx);
      if (isErr(list)) return list;
      out += list.join('');
    }
    return fitting(out);
  },
});

text('CONCATENATE', 'text1, [text2], …', ['...text'], (...parts) => fitting(parts.join('')));

define('TEXTJOIN', {
  min: 3, max: 255, category: 'text', args: 'delimiter, ignore_empty, text1, [text2], …',
  fn: ([delimiter, ignore, ...rest], ctx) => {
    const delims = texts(delimiter, ctx);
    if (isErr(delims)) return delims;
    const skip = toBool(first(deref(ignore, ctx)));
    if (isErr(skip)) return skip;
    const parts = [];
    for (const a of rest) {
      const list = texts(a, ctx);
      if (isErr(list)) return list;
      for (const t of list) if (!(skip && t === '')) parts.push(t);
    }
    let out = parts[0] ?? '';
    for (let i = 1; i < parts.length; i++) out += delims[(i - 1) % delims.length] + parts[i];
    return fitting(out);
  },
});

// ---- parts of a text --------------------------------------------------------------------------------------------------------------

text('LEFT', 'text, [num_chars]', ['text', 'number?'], (t, n = 1) => (n < 0 ? VALUE : t.slice(0, Math.trunc(n))));
text('RIGHT', 'text, [num_chars]', ['text', 'number?'], (t, n = 1) => (n < 0 ? VALUE : t.slice(Math.max(0, t.length - Math.trunc(n)))));
text('MID', 'text, start_num, num_chars', ['text', 'number', 'number'], (t, s, n) => {
  const start = Math.trunc(s);
  return start < 1 || n < 0 ? VALUE : t.slice(start - 1, start - 1 + Math.trunc(n));
});
text('LEN', 'text', ['text'], t => t.length);

text('FIND', 'find_text, within_text, [start_num]', ['text', 'text', 'number?'], (f, w, s = 1) => {
  const start = Math.trunc(s);
  if (start < 1 || start > w.length + 1) return VALUE;
  const i = w.indexOf(f, start - 1);
  return i < 0 ? VALUE : i + 1;
});

/** SEARCH: without case, with the wildcards * ? and ~. */
text('SEARCH', 'find_text, within_text, [start_num]', ['text', 'text', 'number?'], (f, w, s = 1) => {
  const start = Math.trunc(s);
  if (start < 1 || start > w.length + 1) return VALUE;
  if (f === '') return start;
  const m = new RegExp(wildcard(f).source.slice(1, -1), 'i').exec(w.slice(start - 1));
  return m ? m.index + start : VALUE;
});

text('SUBSTITUTE', 'text, old_text, new_text, [instance_num]', ['text', 'text', 'text', 'number?'], (t, o, n, i) => {
  if (i !== undefined && Math.trunc(i) < 1) return VALUE;
  if (o === '') return t;
  if (i === undefined) return fitting(t.split(o).join(n));
  let pos = -1;
  let from = 0;
  for (let k = 0; k < Math.trunc(i); k++) {
    pos = t.indexOf(o, from);
    if (pos < 0) return t;
    from = pos + o.length;
  }
  return fitting(t.slice(0, pos) + n + t.slice(pos + o.length));
});

text('REPLACE', 'old_text, start_num, num_chars, new_text', ['text', 'number', 'number', 'text'], (t, s, n, r) => {
  const start = Math.trunc(s);
  return start < 1 || n < 0 ? VALUE : fitting(t.slice(0, start - 1) + r + t.slice(start - 1 + Math.trunc(n)));
});

text('REPT', 'text, number_times', ['text', 'number'], (t, x) => {
  const n = Math.trunc(x);
  return n < 0 || t.length * n > 32767 ? VALUE : t.repeat(n);
});

// ---- case and cleaning ----------------------------------------------------------------------------------------------------------

/** Upper case like Excel: ß stays ß (JavaScript would write SS). */
const upper = t => t.split('ß').map(p => p.toUpperCase()).join('ß');

text('UPPER', 'text', ['text'], upper);
text('LOWER', 'text', ['text'], t => t.toLowerCase());
/** PROPER: every letter after a character that is not a letter in upper case, the others in lower case ("it's" → "It'S", like Excel). */
text('PROPER', 'text', ['text'], t => t.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + upper(b)));
/** TRIM removes the spaces (only char 32) at both ends and leaves one between words. */
text('TRIM', 'text', ['text'], t => t.replace(/ {2,}/g, ' ').replace(/^ | $/g, ''));
text('CLEAN', 'text', ['text'], t => t.replace(/[\x00-\x1f]/g, ''));
text('EXACT', 'text1, text2', ['text', 'text'], (a, b) => a === b);
text('T', 'value', ['any'], v => (isErr(v) ? v : typeof v === 'string' ? v : ''));

// ---- characters ----------------------------------------------------------------------------------------------------------------

/** The characters 128 … 159 of Windows-1252, the code page of CHAR and CODE in Excel for Windows. */
const CP1252 = '€\x81‚ƒ„…†‡ˆ‰Š‹Œ\x8DŽ\x8F\x90‘’“”•–—˜™š›œ\x9DžŸ';

text('CHAR', 'number', ['number'], x => {
  const n = Math.trunc(x);
  return n < 1 || n > 255 ? VALUE : n >= 128 && n < 160 ? CP1252[n - 128] : String.fromCharCode(n);
});
text('CODE', 'text', ['text'], t => {
  if (!t) return VALUE;
  const c = t.charCodeAt(0);
  if (c < 128 || (c >= 160 && c <= 255)) return c;
  const i = CP1252.indexOf(t[0]);
  return i >= 0 ? 128 + i : 63;
});
text('UNICHAR', 'number', ['number'], x => {
  const n = Math.trunc(x);
  return n < 1 || n > 0x10ffff ? VALUE : n >= 0xd800 && n <= 0xdfff ? NA : String.fromCodePoint(n);
});
text('UNICODE', 'text', ['text'], t => (t ? t.codePointAt(0) : VALUE));

// ---- numbers and text -----------------------------------------------------------------------------------------------------------

/** VALUE: text as a number the way it is typed in the locale (also dates, times, percent and currency). */
text('VALUE', 'text', ['any'], v => {
  if (isErr(v) || typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v !== 'string') return VALUE;
  const n = parseNumber(v);
  return n === null ? VALUE : n;
});

/** NUMBERVALUE: text with given separators (default: the locale's); spaces are ignored, every trailing % divides by 100. */
text('NUMBERVALUE', 'text, [decimal_separator], [group_separator]', ['text', 'text?', 'text?'], (t, d, g) => {
  const dec = d === undefined ? decimalSep() : d[0];
  const grp = g === undefined ? groupSep() : g[0] ?? '';
  if (!dec || dec === grp) return VALUE;
  let s = t.replace(/\s/g, '');
  if (!s) return 0;
  let percent = 0;
  while (s.endsWith('%')) { percent++; s = s.slice(0, -1); }
  const di = s.indexOf(dec);
  if (di >= 0 && s.indexOf(dec, di + 1) >= 0) return VALUE;
  let whole = di >= 0 ? s.slice(0, di) : s;
  const frac = di >= 0 ? s.slice(di + 1) : '';
  if (grp) {
    if (frac.includes(grp)) return VALUE;
    whole = whole.split(grp).join('');
  }
  const plain = whole + (di >= 0 ? '.' + frac : '');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(plain)) return VALUE;
  return Number(plain) / 100 ** percent;
});

/** FIXED: a number rounded to decimals as text with the separators of the locale (no_commas leaves out the group separator). */
text('FIXED', 'number, [decimals], [no_commas]', ['number', 'number?', 'bool?'], (x, d = 2, plain = false) => {
  const digits = Math.trunc(d);
  if (digits > 127) return VALUE;
  const r = roundTo(x, digits);
  let [whole, frac] = Math.abs(r).toFixed(Math.min(Math.max(digits, 0), 100)).split('.');
  if (!plain) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, groupSep());
  return (r < 0 ? '-' : '') + whole + (frac ? decimalSep() + frac : '');
});

// ---- TEXTBEFORE, TEXTAFTER, TEXTSPLIT ----------------------------------------------------------------------------------------------

/** Where the delimiters are in a text: [start, end) pairs, left to right without overlaps (the longest delimiter wins at a position). */
function delimitersIn(t, delims, caseless) {
  const s = caseless ? t.toLowerCase() : t;
  const ds = delims.map(d => (caseless ? d.toLowerCase() : d)).sort((a, b) => b.length - a.length);
  const out = [];
  for (let i = 0; i <= s.length;) {
    const d = ds.find(x => s.startsWith(x, i));
    if (d === undefined) { i++; continue; }
    out.push([i, i + d.length]);
    i += d.length || 1;
  }
  return out;
}

/** TEXTBEFORE/TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]); text may be an array. */
function around(after) {
  return ([value, delimiter, instance, mode, end, notFound], ctx) => {
    const delims = texts(delimiter, ctx);
    if (isErr(delims)) return delims;
    const opts = [optNumber(instance, ctx, 1), optNumber(mode, ctx, 0), optNumber(end, ctx, 0)];
    const e = opts.find(isErr);
    if (e) return e;
    const n = Math.trunc(opts[0]);
    const caseless = opts[1] !== 0;
    const matchEnd = opts[2] !== 0;
    const missing = notFound === undefined || notFound === null ? NA : first(deref(notFound, ctx));
    const one = v => {
      const t = toText(v);
      if (isErr(t)) return t;
      if (n === 0 || Math.abs(n) > t.length) return VALUE;
      const found = delimitersIn(t, delims, caseless);
      let hit = null;
      if (n > 0) hit = found.length >= n ? found[n - 1] : matchEnd && found.length === n - 1 ? [t.length, t.length] : null;
      else hit = found.length >= -n ? found[found.length + n] : matchEnd && found.length === -n - 1 ? [0, 0] : null;
      if (!hit) return missing;
      return after ? t.slice(hit[1]) : t.slice(0, hit[0]);
    };
    const v = deref(value, ctx);
    return Array.isArray(v) ? v.map(row => row.map(one)) : one(v);
  };
}

const aroundArgs = 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]';
define('TEXTBEFORE', { min: 2, max: 6, category: 'text', args: aroundArgs, fn: around(false) });
define('TEXTAFTER', { min: 2, max: 6, category: 'text', args: aroundArgs, fn: around(true) });

define('TEXTSPLIT', {
  min: 2, max: 6, category: 'text', args: 'text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]',
  fn: ([value, colDelimiter, rowDelimiter, ignore, mode, pad], ctx) => {
    const t = toText(first(deref(value, ctx)));
    if (isErr(t)) return t;
    const delimiters = d => (d === undefined || d === null ? [] : texts(d, ctx));
    const cols = delimiters(colDelimiter);
    if (isErr(cols)) return cols;
    const rows = delimiters(rowDelimiter);
    if (isErr(rows)) return rows;
    const skip = optNumber(ignore, ctx, 0);
    if (isErr(skip)) return skip;
    const caseless = optNumber(mode, ctx, 0);
    if (isErr(caseless)) return caseless;
    const filler = pad === undefined || pad === null ? NA : first(deref(pad, ctx));
    const split = (s, delims) => {
      const parts = [];
      let from = 0;
      for (const [a, b] of delimitersIn(s, delims.filter(Boolean), caseless !== 0)) { parts.push(s.slice(from, a)); from = b; }
      parts.push(s.slice(from));
      return skip ? parts.filter(p => p !== '') : parts;
    };
    const table = (rows.some(Boolean) ? split(t, rows) : [t]).map(line => (cols.some(Boolean) ? split(line, cols) : [line])).filter(r => r.length);
    if (!table.length) return CALC;
    const width = table.reduce((w, r) => Math.max(w, r.length), 0);
    return table.map(r => (r.length < width ? [...r, ...new Array(width - r.length).fill(filler)] : r));
  },
});
