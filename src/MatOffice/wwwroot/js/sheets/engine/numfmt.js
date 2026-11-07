/* MatOffice Sheets – Excel number formats (pure, no DOM): the text of a value under a format code, what a typed text becomes, and the
   format codes of the ribbon. Codes are stored the way Excel writes them in a file ("." decimal point, "," thousands); the text follows the
   locale (German: 1.234,56). Sections: positive;negative;zero;text, conditions [>=100], colours [Red], currencies [$€-407], elapsed [h]:mm. */
import { ymd, hms, weekday, parseDate, MONTHS, DAYS } from './dates.js';

const COLOURS = { black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00a000', magenta: '#ff00ff', red: '#ff0000', white: '#ffffff', yellow: '#ffff00' };
const PALETTE56 = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff', '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99', '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];

const SHORT_MONTHS = {
  en: MONTHS.en.map(m => m.slice(0, 3)),
  de: ['Jan', 'Feb', 'Mrz', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
};
const SHORT_DAYS = { en: DAYS.en.map(d => d.slice(0, 3)), de: ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'] };

// ---- parsing a code ---------------------------------------------------------------------------------------------------------------

const cache = new Map();

/** The sections of a code, parsed (cached). */
export function compile(code) {
  const c = code === undefined || code === null || code === '' ? 'General' : String(code);
  let compiled = cache.get(c);
  if (!compiled) {
    compiled = splitSections(c).map(parseSection);
    if (cache.size > 2000) cache.clear();
    cache.set(c, compiled);
  }
  return compiled;
}

function splitSections(code) {
  const out = [];
  let cur = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const j = end < 0 ? code.length - 1 : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += code.slice(i, i + 2);
      i++;
    } else if (ch === '[') {
      const end = code.indexOf(']', i);
      const j = end < 0 ? code.length - 1 : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === ';') {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.slice(0, 4);
}

const DATE_CHARS = /[yYmMdDhHsS]|AM\/PM|A\/P/;

function parseSection(src) {
  const s = { tokens: [], color: null, cond: null, date: false, elapsed: false, general: false, text: false, percent: 0, scale: 0, group: false, exp: null, frac: null };
  const t = s.tokens;
  const lit = text => { const last = t[t.length - 1]; if (last && last.type === 'lit') last.text += text; else t.push({ type: 'lit', text }); };
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    const rest = src.slice(i);
    if (ch === '"') {
      const end = src.indexOf('"', i + 1);
      const j = end < 0 ? src.length : end;
      lit(src.slice(i + 1, j));
      i = j;
    } else if (ch === '\\') {
      if (i + 1 < src.length) lit(src[++i]);
    } else if (ch === '_') {
      if (i + 1 < src.length) { i++; lit(' '); }
    } else if (ch === '*') {
      if (i + 1 < src.length) { i++; t.push({ type: 'fill', char: src[i] }); }
    } else if (ch === '[') {
      const end = src.indexOf(']', i);
      const inner = src.slice(i + 1, end < 0 ? src.length : end);
      i = end < 0 ? src.length : end;
      const lower = inner.toLowerCase();
      if (COLOURS[lower]) s.color = COLOURS[lower];
      else if (/^color\s*\d+$/i.test(inner)) s.color = PALETTE56[(Number(inner.replace(/\D/g, '')) - 1) % 56] || null;
      else if (/^(<=|>=|<>|<|>|=)\s*-?[\d.]+$/.test(inner)) { const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner); s.cond = { op: m[1], v: Number(m[2]) }; }
      else if (inner.startsWith('$')) { const sym = inner.slice(1).split('-')[0]; if (sym) lit(sym); }
      else if (/^(h+|m+|s+)$/i.test(inner)) { s.date = true; s.elapsed = true; t.push({ type: 'elapsed', unit: lower[0], len: inner.length }); }
      // anything else ([DBNum1], [$-409] …) is ignored
    } else if (/^general/i.test(rest)) {
      s.general = true;
      t.push({ type: 'general' });
      i += 6;
    } else if (ch === '@') {
      s.text = true;
      t.push({ type: 'at' });
    } else if (/^(AM\/PM|am\/pm)/.test(rest)) {
      s.date = true;
      t.push({ type: 'ampm', full: true });
      i += 4;
    } else if (/^(A\/P|a\/p)/.test(rest)) {
      s.date = true;
      t.push({ type: 'ampm', full: false, lower: ch === 'a' });
      i += 2;
    } else if (/[yYdD]/.test(ch) || /[mMhHsS]/.test(ch) && !/[eE]/.test(ch)) {
      let j = i;
      while (j < src.length && src[j].toLowerCase() === ch.toLowerCase()) j++;
      s.date = true;
      t.push({ type: 'date', unit: ch.toLowerCase(), len: j - i });
      i = j - 1;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      t.push({ type: 'digit', char: ch });
    } else if (ch === '.') {
      t.push({ type: 'point' });
    } else if (ch === ',') {
      t.push({ type: 'comma' });
    } else if (ch === '%') {
      s.percent++;
      lit('%');
    } else if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) {
      t.push({ type: 'exp', sign: src[i + 1] });
      i++;
    } else if (ch === '/' && t.some(x => x.type === 'digit') && /^\/\s*[0-9#?]/.test(rest)) {
      t.push({ type: 'slash' });
    } else {
      lit(ch);
    }
  }
  if (s.date) resolveMinutes(t);
  else analyseNumber(s);
  return s;
}

/** In a date section a point, a comma and a digit are text (except the fraction of seconds, see resolveMinutes). */
function dateLiterals(t) {
  for (let i = 0; i < t.length; i++) {
    if (t[i].type === 'point') t[i] = { type: 'lit', text: '.' };
    else if (t[i].type === 'comma') t[i] = { type: 'lit', text: ',' };
    else if (t[i].type === 'digit') t[i] = { type: 'lit', text: t[i].char === '0' ? '0' : '' };
  }
  for (let i = t.length - 1; i > 0; i--) {
    if (t[i].type === 'lit' && t[i - 1].type === 'lit') { t[i - 1].text += t[i].text; t.splice(i, 1); }
  }
}

/** "m" after an hour or before a second is minutes; otherwise it is the month. */
function resolveMinutes(t) {
  const units = t.filter(x => x.type === 'date' || x.type === 'elapsed');
  units.forEach((u, i) => {
    if (u.type !== 'date' || u.unit !== 'm' || u.len > 2) return;
    const prev = units[i - 1];
    const next = units[i + 1];
    if ((prev && (prev.unit === 'h')) || (next && next.unit === 's')) u.unit = 'n';
  });
  // fractional seconds: "ss.00"
  for (let i = 0; i < t.length - 1; i++) {
    if (t[i].type === 'date' && t[i].unit === 's' && t[i + 1].type === 'point') {
      let j = i + 2;
      let n = 0;
      while (j < t.length && t[j].type === 'digit' && t[j].char === '0') { n++; j++; }
      if (n) { t[i].frac = n; t.splice(i + 1, j - i - 1); }
    }
  }
  dateLiterals(t);
}

function analyseNumber(s) {
  const t = s.tokens;
  const expAt = t.findIndex(x => x.type === 'exp');
  const slashAt = t.findIndex(x => x.type === 'slash');
  const pointAt = t.findIndex(x => x.type === 'point');
  // commas: between integer digits = grouping; after the last digit of the number = scaling by 1000 each; anything else is text
  const intEnd = pointAt >= 0 ? pointAt : expAt >= 0 ? expAt : t.length;
  const numEnd = expAt >= 0 ? expAt : t.length;
  for (let i = 0; i < t.length; i++) {
    if (t[i].type !== 'comma') continue;
    const before = t.slice(0, i).some(x => x.type === 'digit');
    const afterInt = i < intEnd && t.slice(i + 1, intEnd).some(x => x.type === 'digit');
    const afterAny = t.slice(i + 1, numEnd).some(x => x.type === 'digit');
    if (before && afterInt) s.group = true;
    else if (before && !afterAny) s.scale++;
    else t[i] = { type: 'lit', text: ',' };
  }
  if (expAt >= 0) s.exp = { at: expAt, sign: t[expAt].sign };
  if (slashAt >= 0) s.frac = { at: slashAt };
}

// ---- formatting ---------------------------------------------------------------------------------------------------------------------

/**
 * The text of a value under a format code: { text, color (hex or null), fill (index where a "*" fill pads, or -1) }.
 * options: locale ('en'|'de'), width (characters for General; numbers are shortened to fit).
 */
export function formatValue(value, code, options = {}) {
  const locale = options.locale === 'de' ? 'de' : 'en';
  const sections = compile(code);
  if (value === null || value === undefined) return { text: '', color: null, fill: -1 };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', color: null, fill: -1 };
  if (typeof value === 'object' && value && value.code) return { text: value.code, color: null, fill: -1 };
  if (typeof value === 'string') {
    const textSection = sections.length >= 4 ? sections[3] : sections.find(x => x.text) && sections.length === 1 ? sections[0] : null;
    if (!textSection) return { text: value, color: null, fill: -1 };
    return render(textSection, value, locale, options, true);
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) return { text: '#NUM!', color: null, fill: -1 };
  const { section, abs } = pickSection(sections, value);
  if (!section) return { text: '', color: null, fill: -1 };
  return render(section, abs ? Math.abs(value) : value, locale, options, false);
}

function pickSection(sections, v) {
  const numeric = sections.slice(0, 3);
  if (numeric.some(s => s.cond)) {
    for (let i = 0; i < numeric.length; i++) {
      const s = numeric[i];
      if (s.cond && test(s.cond, v)) return { section: s, abs: i === 1 && v < 0 };
    }
    const rest = numeric.filter(s => !s.cond);
    return { section: rest[0] || numeric[numeric.length - 1], abs: false };
  }
  if (numeric.length === 1) return { section: numeric[0], abs: false };
  if (v > 0 || (v === 0 && numeric.length === 2)) return { section: numeric[0], abs: false };
  if (v < 0) return { section: numeric[1], abs: true };
  return { section: numeric[2], abs: false };
}

function test(c, v) {
  switch (c.op) {
    case '<': return v < c.v;
    case '<=': return v <= c.v;
    case '>': return v > c.v;
    case '>=': return v >= c.v;
    case '=': return v === c.v;
    default: return v !== c.v;
  }
}

function render(s, value, locale, options, isText) {
  let text;
  if (isText) text = s.tokens.map(x => (x.type === 'at' ? value : x.type === 'lit' ? x.text : x.type === 'fill' ? '\u0001' : '')).join('');
  else if (s.general || (!s.date && !s.tokens.some(x => x.type === 'digit'))) text = generalWithLiterals(s, value, locale, options);
  else if (s.date) text = formatDate(s, value, locale);
  else text = formatNumber(s, value, locale);
  const fill = text.indexOf('\u0001');
  return { text: text.replace(/\u0001/g, ''), color: s.color, fill };
}

function generalWithLiterals(s, value, locale, options) {
  let out = '';
  let used = false;
  for (const x of s.tokens) {
    if (x.type === 'general' && !used) { out += general(value, locale, options.width); used = true; } else if (x.type === 'lit') out += x.text;
    else if (x.type === 'fill') out += '\u0001';
    else if (x.type === 'at') out += general(value, locale, options.width);
  }
  return used || s.tokens.some(x => x.type === 'at') ? out : out || '';
}

/** Excel's General: up to 11 characters (fewer when the column is narrow), exponent notation for very large and small numbers. */
export function general(v, locale = 'en', width = 11) {
  const max = Math.max(1, Math.min(11, width || 11));
  if (v === 0) return '0';
  const dec = locale === 'de' ? ',' : '.';
  const neg = v < 0;
  const abs = Math.abs(v);
  const room = max - (neg ? 1 : 0);
  const intLen = abs >= 1 ? Math.floor(Math.log10(abs)) + 1 : 1;
  if (abs >= 1e-5 && intLen <= room) {
    const decimals = Math.max(0, Math.min(abs >= 1 ? room - intLen - 1 : room - 2, 15 - (abs >= 1 ? intLen : 0)));
    let txt = roundFixed(abs, decimals);
    if (txt.includes('.')) txt = txt.replace(/\.?0+$/, '');
    if (txt.length <= room && txt !== '0') return (neg ? '-' : '') + txt.replace('.', dec);
  }
  // exponent notation
  for (let digits = Math.max(0, room - 6); digits >= 0; digits--) {
    const [m, e] = abs.toExponential(digits).split('e');
    const mant = m.includes('.') ? m.replace(/\.?0+$/, '') : m;
    const exp = Number(e);
    const txt = mant.replace('.', dec) + 'E' + (exp < 0 ? '-' : '+') + String(Math.abs(exp)).padStart(2, '0');
    if (txt.length <= room || digits === 0) return (neg ? '-' : '') + txt;
  }
  return '#';
}

/** A non-negative number rounded half up to a number of decimals, from its 15 significant digits (so 1.005 → "1.01"). */
export function roundFixed(x, decimals) {
  if (x === 0) return decimals > 0 ? '0.' + '0'.repeat(decimals) : '0';
  const [m, e] = x.toExponential(14).split('e');
  let digits = m.replace('.', '').replace(/0+$/, '') || '0';
  let point = Number(e) + 1; // digits before the decimal point
  const cut = point + decimals;
  if (cut < 0) return decimals > 0 ? '0.' + '0'.repeat(decimals) : '0';
  if (digits.length > cut) {
    const up = digits[cut] >= '5';
    digits = digits.slice(0, cut);
    if (up) {
      let i = digits.length - 1;
      const arr = digits.split('');
      while (i >= 0 && arr[i] === '9') { arr[i] = '0'; i--; }
      if (i >= 0) arr[i] = String(Number(arr[i]) + 1);
      else { arr.unshift('1'); point++; }
      digits = arr.join('');
    }
  }
  digits = digits.padEnd(point + decimals, '0');
  let int;
  let frac;
  if (point <= 0) { int = '0'; frac = ('0'.repeat(-point) + digits).slice(0, decimals).padEnd(decimals, '0'); } else { int = digits.slice(0, point) || '0'; frac = digits.slice(point, point + decimals); }
  int = int.replace(/^0+(?=\d)/, '');
  return decimals > 0 ? int + '.' + frac : int;
}

function groupDigits(int, sep) {
  let out = '';
  for (let i = 0; i < int.length; i++) {
    if (i && (int.length - i) % 3 === 0) out += sep;
    out += int[i];
  }
  return out;
}

function formatNumber(s, value, locale) {
  const dec = locale === 'de' ? ',' : '.';
  const grp = locale === 'de' ? '.' : ',';
  let v = value * Math.pow(100, s.percent) / Math.pow(1000, s.scale);
  const negative = v < 0;
  v = Math.abs(v);
  const t = s.tokens;
  if (s.exp) return (negative ? '-' : '') + scientific(s, v, dec);
  if (s.frac) return fraction(s, v, negative, dec);
  const pointAt = t.findIndex(x => x.type === 'point');
  const intTokens = pointAt >= 0 ? t.slice(0, pointAt) : t;
  const fracTokens = pointAt >= 0 ? t.slice(pointAt + 1) : [];
  const fracDigits = fracTokens.filter(x => x.type === 'digit');
  const fixed = roundFixed(v, fracDigits.length);
  let [intPart, fracPart = ''] = fixed.split('.');
  if (intPart === '0') intPart = '';
  // integer part: fill the placeholders from the right
  const intDigits = intTokens.filter(x => x.type === 'digit');
  let intText = '';
  if (s.group) {
    const zeros = intDigits.filter(x => x.char === '0').length;
    const padded = intPart.padStart(zeros, '0');
    const firstDigit = intTokens.findIndex(x => x.type === 'digit');
    const prefix = intTokens.slice(0, Math.max(0, firstDigit)).map(tokenText).join('');
    const lastDigit = intTokens.map(x => x.type).lastIndexOf('digit');
    const suffix = intTokens.slice(lastDigit + 1).map(tokenText).join('');
    intText = prefix + groupDigits(padded, grp) + suffix;
    if (!padded && intDigits.some(x => x.char === '?')) intText = prefix + ' ' + suffix;
  } else {
    let digits = intPart;
    const out = [];
    let remaining = intDigits.length;
    for (let i = intTokens.length - 1; i >= 0; i--) {
      const x = intTokens[i];
      if (x.type !== 'digit') { out.push(tokenText(x)); continue; }
      remaining--;
      if (remaining === 0) {
        if (digits) { out.push(digits); digits = ''; } else out.push(x.char === '0' ? '0' : x.char === '?' ? ' ' : '');
      } else if (digits) {
        out.push(digits[digits.length - 1]);
        digits = digits.slice(0, -1);
      } else {
        out.push(x.char === '0' ? '0' : x.char === '?' ? ' ' : '');
      }
    }
    intText = out.reverse().join('');
  }
  // fraction part: fill from the left
  let fracText = '';
  if (pointAt >= 0) {
    let i = 0;
    const chars = [];
    for (const x of fracTokens) {
      if (x.type === 'digit') { chars.push({ d: fracPart[i] || '0', char: x.char }); i++; } else chars.push({ lit: tokenText(x) });
    }
    // trailing zeros of # and ? placeholders
    for (let k = chars.length - 1; k >= 0; k--) {
      const c = chars[k];
      if (c.lit !== undefined) continue;
      if (c.d !== '0' || c.char === '0') break;
      c.d = c.char === '?' ? ' ' : '';
    }
    fracText = dec + chars.map(c => (c.lit !== undefined ? c.lit : c.d)).join('');
  }
  return (negative ? '-' : '') + intText + fracText;
}

function tokenText(x) {
  switch (x.type) {
    case 'lit': return x.text;
    case 'fill': return '\u0001';
    case 'comma': return '';
    case 'general': return '';
    default: return '';
  }
}

function scientific(s, v, dec) {
  const t = s.tokens;
  const at = s.exp.at;
  const mantTokens = t.slice(0, at);
  const expTokens = t.slice(at + 1);
  const pointAt = mantTokens.findIndex(x => x.type === 'point');
  const intCount = (pointAt >= 0 ? mantTokens.slice(0, pointAt) : mantTokens).filter(x => x.type === 'digit').length || 1;
  const fracCount = pointAt >= 0 ? mantTokens.slice(pointAt + 1).filter(x => x.type === 'digit').length : 0;
  let exp = v === 0 ? 0 : Math.floor(Math.log10(v));
  if (intCount > 1 && mantTokens.some(x => x.char === '#')) exp = Math.floor(exp / intCount) * intCount; // engineering notation
  else exp -= intCount - 1;
  let mant = v / Math.pow(10, exp);
  let txt = roundFixed(mant, fracCount);
  if (Number(txt) >= Math.pow(10, intCount) && !(intCount > 1 && mantTokens.some(x => x.char === '#'))) { exp += 1; mant = v / Math.pow(10, exp); txt = roundFixed(mant, fracCount); }
  const prefix = mantTokens.filter(x => x.type === 'lit').map(x => x.text).join('');
  const expDigits = expTokens.filter(x => x.type === 'digit').length || 1;
  const sign = exp < 0 ? '-' : s.exp.sign === '+' ? '+' : '';
  return prefix + txt.replace('.', dec) + 'E' + sign + String(Math.abs(exp)).padStart(expDigits, '0') + expTokens.filter(x => x.type === 'lit').map(x => x.text).join('');
}

function fraction(s, v, negative, dec) {
  const t = s.tokens;
  const slash = s.frac.at;
  // the numerator is the digit group right before the slash, an integer part may stand before it (separated by a literal)
  let k = slash - 1;
  while (k >= 0 && t[k].type === 'digit') k--;
  const numTokens = t.slice(k + 1, slash);
  const head = t.slice(0, k + 1);
  const intTokens = head.filter(x => x.type === 'digit');
  const denTokens = [];
  let j = slash + 1;
  while (j < t.length && (t[j].type === 'digit' || (t[j].type === 'lit' && /^\d+$/.test(t[j].text)))) denTokens.push(t[j++]);
  const tail = t.slice(j).filter(x => x.type === 'lit').map(x => x.text).join('');
  const fixedDen = denTokens.length && denTokens.every(x => x.type === 'lit' || x.char === '0') && denTokens.some(x => x.type === 'lit') ? Number(denTokens.map(x => x.text || '').join('')) : null;
  const literalDen = denTokens.map(x => (x.type === 'lit' ? x.text : '')).join('');
  const den = fixedDen || (literalDen && /^\d+$/.test(literalDen) ? Number(literalDen) : null);
  const maxDen = den ? null : Math.pow(10, Math.max(1, denTokens.length)) - 1;
  let whole = intTokens.length ? Math.floor(v) : 0;
  let rest = v - whole;
  let num;
  let d;
  if (den) { d = den; num = Math.round(rest * den); } else [num, d] = approximate(rest, maxDen);
  if (num === d && intTokens.length) { whole += 1; num = 0; }
  const width = Math.max(numTokens.length, 1);
  const headText = head.filter(x => x.type === 'lit').map(x => x.text).join('');
  const sign = negative ? '-' : '';
  if (num === 0 && intTokens.length) {
    const blank = ' '.repeat(width + 1 + Math.max(denTokens.length, 1));
    return sign + (whole || intTokens.some(x => x.char === '0') ? String(whole) : '0') + (headText ? ' '.repeat(headText.length) : '') + blank + tail;
  }
  const wholeText = intTokens.length && (whole || intTokens.some(x => x.char === '0')) ? String(whole) + headText : intTokens.length ? ' '.repeat(String(whole).length) + headText.replace(/./g, ' ') : '';
  const numText = String(num).padStart(width, numTokens.some(x => x.char === '0') ? '0' : ' ');
  const denText = den ? String(d) : String(d).padEnd(Math.max(denTokens.length, 1), ' ');
  return sign + wholeText + numText + '/' + denText + tail + (dec ? '' : '');
}

/** The closest fraction n/d with d ≤ maxDen (continued fractions). */
export function approximate(x, maxDen) {
  if (x === 0) return [0, 1];
  let [h0, h1, k0, k1] = [0, 1, 1, 0];
  let f = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(f);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDen) break;
    [h0, h1, k0, k1] = [h1, h2, k1, k2];
    if (Math.abs(f - a) < 1e-12) break;
    f = 1 / (f - a);
  }
  // the best semi-convergent below the limit
  const bound = Math.floor((maxDen - k0) / k1);
  const h = bound * h1 + h0;
  const k = bound * k1 + k0;
  if (k > 0 && k <= maxDen && Math.abs(x - h / k) < Math.abs(x - h1 / k1)) return [h, k];
  return [h1, k1];
}

function formatDate(s, value, locale) {
  if (value < 0) return '#'.repeat(8);
  // rounding: to the precision that is shown
  const fracDigits = Math.max(0, ...s.tokens.filter(x => x.frac).map(x => x.frac));
  const hasSeconds = s.tokens.some(x => x.unit === 's');
  const unit = hasSeconds ? 86400 * Math.pow(10, fracDigits) : s.tokens.some(x => x.unit === 'n' || x.unit === 'h') ? 1440 : 86400;
  const v = Math.round(value * unit) / unit;
  const { y, m, d } = ymd(v);
  const time = hms(v + 1e-9);
  const ampm = s.tokens.some(x => x.type === 'ampm');
  const names = { months: MONTHS[locale], short: SHORT_MONTHS[locale], days: DAYS[locale], shortDays: SHORT_DAYS[locale] };
  let out = '';
  for (const x of s.tokens) {
    switch (x.type) {
      case 'lit': out += x.text; break;
      case 'fill': out += '\u0001'; break;
      case 'ampm': {
        const pm = time.h >= 12;
        out += x.full ? (pm ? 'PM' : 'AM') : x.lower ? (pm ? 'p' : 'a') : (pm ? 'P' : 'A');
        break;
      }
      case 'elapsed': {
        const total = x.unit === 'h' ? Math.floor(v * 24) : x.unit === 'm' ? Math.floor(v * 1440) : Math.floor(v * 86400 + 1e-6);
        out += String(total).padStart(x.len, '0');
        break;
      }
      case 'date':
        switch (x.unit) {
          case 'y': out += x.len <= 2 ? String(y % 100).padStart(2, '0') : String(y); break;
          case 'm':
            if (x.len === 1) out += m;
            else if (x.len === 2) out += String(m).padStart(2, '0');
            else if (x.len === 3) out += names.short[m - 1];
            else if (x.len === 5) out += names.months[m - 1][0];
            else out += names.months[m - 1];
            break;
          case 'd':
            if (x.len === 1) out += d;
            else if (x.len === 2) out += String(d).padStart(2, '0');
            else if (x.len === 3) out += names.shortDays[weekday(v)];
            else out += names.days[weekday(v)];
            break;
          case 'h': {
            let h = time.h;
            if (ampm) h = h % 12 || 12;
            out += x.len >= 2 ? String(h).padStart(2, '0') : String(h);
            break;
          }
          case 'n': out += x.len >= 2 ? String(time.mi).padStart(2, '0') : String(time.mi); break;
          case 's': {
            out += x.len >= 2 ? String(time.s).padStart(2, '0') : String(time.s);
            if (x.frac) out += (locale === 'de' ? ',' : '.') + String(Math.floor(time.ms / Math.pow(10, 3 - x.frac))).padStart(x.frac, '0');
            break;
          }
          default: break;
        }
        break;
      case 'digit': out += x.char === '0' ? '0' : ''; break;
      default: break;
    }
  }
  return out;
}

// ---- what a format is -------------------------------------------------------------------------------------------------------------

/** 'general' | 'number' | 'currency' | 'accounting' | 'percent' | 'scientific' | 'fraction' | 'date' | 'time' | 'datetime' | 'text'. */
export function formatKind(code) {
  if (!code || /^general$/i.test(code)) return 'general';
  const s = compile(code)[0];
  if (s.date) {
    const units = new Set(s.tokens.filter(x => x.type === 'date' || x.type === 'elapsed').map(x => (x.type === 'elapsed' ? 'h' : x.unit)));
    const hasDate = units.has('y') || units.has('m') || units.has('d');
    const hasTime = units.has('h') || units.has('n') || units.has('s');
    return hasDate && hasTime ? 'datetime' : hasDate ? 'date' : 'time';
  }
  if (s.text && !s.tokens.some(x => x.type === 'digit')) return 'text';
  if (s.percent) return 'percent';
  if (s.exp) return 'scientific';
  if (s.frac) return 'fraction';
  if (/[$€£¥]|\[\$/.test(code)) return s.tokens.some(x => x.type === 'fill') ? 'accounting' : 'currency';
  return 'number';
}

export const isDateFormat = code => ['date', 'time', 'datetime'].includes(formatKind(code));

/** The number of decimals a code shows (first section). */
export function decimalsOf(code) {
  const s = compile(code)[0];
  const t = s.tokens;
  const p = t.findIndex(x => x.type === 'point');
  if (p < 0) return 0;
  let n = 0;
  for (let i = p + 1; i < t.length && (t[i].type === 'digit'); i++) n++;
  return n;
}

/** Increase (+1) or decrease (-1) the decimals of a code (General takes the decimals the value shows). */
export function changeDecimals(code, delta, value) {
  if (!code || /^general$/i.test(code)) {
    let shown = 0;
    if (typeof value === 'number') { const g = general(value); const m = /[.,](\d+)/.exec(g); shown = m && !/E/.test(g) ? m[1].length : 0; }
    const n = Math.max(0, shown + delta);
    return n ? '0.' + '0'.repeat(n) : '0';
  }
  if (isDateFormat(code) || formatKind(code) === 'text') return code;
  return splitSections(code).map(section => changeSection(section, delta)).join(';');
}

function changeSection(src, delta) {
  // find the last digit placeholder of the number part (outside quotes and brackets)
  const parts = [];
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') { const e = src.indexOf('"', i + 1); const j = e < 0 ? src.length - 1 : e; parts.push({ lit: true, text: src.slice(i, j + 1) }); i = j; } else if (ch === '[') { const e = src.indexOf(']', i); const j = e < 0 ? src.length - 1 : e; parts.push({ lit: true, text: src.slice(i, j + 1) }); i = j; } else if (ch === '\\' || ch === '_' || ch === '*') { parts.push({ lit: true, text: src.slice(i, i + 2) }); i++; } else parts.push({ lit: false, text: ch });
  }
  const isDigit = p => !p.lit && /[0#?]/.test(p.text);
  const expAt = parts.findIndex(p => !p.lit && /[eE]/.test(p.text));
  const limit = expAt >= 0 ? expAt : parts.length;
  const pointAt = parts.findIndex((p, i) => i < limit && !p.lit && p.text === '.');
  let last = -1;
  for (let i = 0; i < limit; i++) if (isDigit(parts[i])) last = i;
  if (last < 0) return src;
  if (delta > 0) {
    if (pointAt >= 0) parts.splice(last + 1, 0, { lit: false, text: '0' });
    else parts.splice(last + 1, 0, { lit: false, text: '.' }, { lit: false, text: '0' });
  } else if (pointAt >= 0) {
    if (last > pointAt) parts.splice(last, 1);
    const stillDigits = parts.some((p, i) => i > pointAt && i < (expAt >= 0 ? expAt : parts.length) && isDigit(p));
    if (!stillDigits) parts.splice(pointAt, 1);
  }
  return parts.map(p => p.text).join('');
}

// ---- formats of the ribbon ------------------------------------------------------------------------------------------------------

/** The format codes of the ribbon's number format list, by locale. */
export function presets(locale) {
  const de = locale === 'de';
  return {
    general: null,
    number: '0.00',
    currency: de ? '#,##0.00 "€"' : '"$"#,##0.00',
    currencyEuro: '#,##0.00 "€"',
    currencyDollar: '"$"#,##0.00',
    currencyPound: '"£"#,##0.00',
    accounting: de ? '_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-' : '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    shortDate: de ? 'dd.mm.yyyy' : 'm/d/yyyy',
    longDate: de ? 'dddd, d. mmmm yyyy' : 'dddd, mmmm d, yyyy',
    time: de ? 'hh:mm:ss' : 'h:mm:ss AM/PM',
    percent: '0.00%',
    fraction: '# ?/?',
    scientific: '0.00E+00',
    text: '@',
    comma: '#,##0.00',
    percentStyle: '0%',
    dateTime: de ? 'dd.mm.yyyy hh:mm' : 'm/d/yyyy h:mm',
    shortTime: de ? 'hh:mm' : 'h:mm',
  };
}

// ---- typed input --------------------------------------------------------------------------------------------------------------

/**
 * What a typed text becomes: { formula } ("=…" without "="), { value, nf? } (a number with the format it suggests, a date, a logical, a text),
 * { error } (#N/A typed), or { empty: true }. A leading ' keeps the rest as text.
 */
export function parseInput(text, locale = 'en', currentYear) {
  const de = locale === 'de';
  if (text === '' || text === null || text === undefined) return { empty: true };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (text.length > 1 && text[0] === '=') return { formula: text.slice(1) };
  const t = text.trim();
  if (!t) return { value: text };
  if (/^[+-]/.test(t) && t.length > 1 && /[A-Za-z(]/.test(t[1]) && !/^[+-]?\d/.test(t)) return { formula: t };
  const up = t.toUpperCase();
  if (up === 'TRUE' || (de && up === 'WAHR')) return { value: true };
  if (up === 'FALSE' || (de && up === 'FALSCH')) return { value: false };
  if (/^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)$/.test(up)) return { error: up };
  const p = presets(locale);
  const dec = de ? ',' : '.';
  const grp = de ? '.' : ',';
  // percent
  let m = /^([+-]?)\s*([\d.,]+)\s*%$/.exec(t);
  if (m) {
    const n = number(m[2], dec, grp);
    if (n !== null) {
      const decimals = (m[2].split(dec)[1] || '').length;
      return { value: (m[1] === '-' ? -n : n) / 100, nf: decimals ? '0.' + '0'.repeat(Math.min(decimals, 10)) + '%' : '0%' };
    }
  }
  // currency
  m = /^([+-]?)\s*([$€£])\s*([\d.,]+)$|^([+-]?)\s*([\d.,]+)\s*([$€£]|EUR)$/.exec(t);
  if (m) {
    const sym = m[2] || m[6];
    const n = number(m[3] || m[5], dec, grp);
    if (n !== null) {
      const negative = (m[1] || m[4]) === '-';
      const code = sym === '€' || sym === 'EUR' ? p.currencyEuro : sym === '£' ? p.currencyPound : p.currencyDollar;
      const decimals = ((m[3] || m[5]).split(dec)[1] || '').length;
      return { value: negative ? -n : n, nf: decimals ? code : code.replace('.00', '') };
    }
  }
  // fraction "1 1/2", "0 3/4"
  m = /^([+-]?)(\d+)\s+(\d+)\/(\d+)$/.exec(t);
  if (m && Number(m[4]) > 0 && Number(m[3]) < Number(m[4])) {
    const v = Number(m[2]) + Number(m[3]) / Number(m[4]);
    return { value: m[1] === '-' ? -v : v, nf: Number(m[4]) > 9 ? '# ??/??' : '# ?/?' };
  }
  // plain numbers, with thousands separators, exponent
  m = /^([+-]?)([\d.,]+)$/.exec(t);
  if (m) {
    const n = number(m[2], dec, grp);
    if (n !== null) {
      const grouped = m[2].includes(grp) && new RegExp(`^\\d{1,3}(\\${grp}\\d{3})+(\\${dec}\\d*)?$`).test(m[2]);
      const decimals = (m[2].split(dec)[1] || '').length;
      return grouped ? { value: m[1] === '-' ? -n : n, nf: decimals ? '#,##0.00' : '#,##0' } : { value: m[1] === '-' ? -n : n };
    }
  }
  m = /^([+-]?)(\d+(?:[.,]\d+)?)[eE]([+-]?\d+)$/.exec(t);
  if (m) {
    const v = Number(m[2].replace(',', '.') + 'e' + m[3]);
    if (Number.isFinite(v)) return { value: m[1] === '-' ? -v : v, nf: '0.00E+00' };
  }
  // dates and times
  const d = parseDate(t, locale, currentYear);
  if (d) {
    if (d.kind === 'time') return { value: d.value, nf: /:\d{1,2}:\d/.test(t) ? (de ? 'hh:mm:ss' : 'h:mm:ss') : /[ap]m?$/i.test(t) ? 'h:mm AM/PM' : p.shortTime };
    if (d.kind === 'datetime') return { value: d.value, nf: p.dateTime };
    return { value: d.value, nf: /[A-Za-zÄÖÜäöü]{3}/.test(t) ? (de ? 'dd. mmm yyyy' : 'd-mmm-yyyy') : p.shortDate };
  }
  return { value: text };
}

function number(s, dec, grp) {
  if (/^\d+$/.test(s)) return Number(s);
  const re = new RegExp(`^(\\d{1,3}(\\${grp}\\d{3})+|\\d+)?(\\${dec}\\d+)?$`);
  if (s && re.test(s) && /\d/.test(s)) return Number(s.split(grp).join('').replace(dec, '.'));
  // a German "." used as decimal point (12.5) and an English "," that is not a group
  if (dec === ',' && /^\d+\.\d+$/.test(s) && !/^\d{1,3}\.\d{3}$/.test(s)) return Number(s);
  return null;
}

/** The text of a constant in the editor and the formula bar (a date as a date, a percentage with %, full precision; text that would be read as something else gets a leading '). */
export function formatForEdit(value, code, locale = 'en') {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') {
    const parsed = parseInput(value, locale);
    return parsed.value === value ? value : "'" + value;
  }
  if (typeof value !== 'number') return String(value);
  const kind = formatKind(code);
  const p = presets(locale);
  if (kind === 'date' || kind === 'datetime' || kind === 'time') {
    const hasTime = value % 1 !== 0;
    if (kind === 'time' && value < 1) return formatValue(value, locale === 'de' ? 'hh:mm:ss' : 'h:mm:ss AM/PM', { locale }).text;
    const date = formatValue(value, p.shortDate, { locale }).text;
    return hasTime ? date + ' ' + formatValue(value, locale === 'de' ? 'hh:mm:ss' : 'h:mm:ss AM/PM', { locale }).text : date;
  }
  const dec = locale === 'de' ? ',' : '.';
  const plain = n => {
    if (n === 0) return '0';
    const abs = Math.abs(n);
    let txt = abs >= 1e15 || abs < 1e-9 ? n.toExponential(14).replace(/\.?0+e/, 'E').replace(/E\+?/, 'E+').replace('E+-', 'E-') : String(Number(n.toPrecision(15)));
    if (/e/.test(txt)) txt = Number(n.toPrecision(15)).toFixed(20).replace(/\.?0+$/, '');
    return txt.replace('.', dec);
  };
  if (kind === 'percent') return plain(Number((value * 100).toPrecision(15))) + '%';
  return plain(value);
}
