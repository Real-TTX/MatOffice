/* MatOffice Base – values like Access keeps them: dates as serial numbers (days since 30 Dec 1899, the time as the fraction), numbers, text, Yes/No.
   Parsing what a person types (in the language of the page) and formatting for the screen (Access formats: Currency, Short Date, Yes/No …).
   A pure module without the DOM: the browser and the node tests use it. Content keeps dates as ISO text ("2026-11-07", "2026-11-07T10:30:00"). */

/** A date/time value of an expression (Access keeps the same: a double). */
export class DateVal {
  constructor(serial) { this.s = serial; }
  valueOf() { return this.s; }
}

const EPOCH = 25569; // serial of 1970-01-01

/** Days since 1970-01-01 for a civil date (proleptic Gregorian). */
export function daysFromCivil(y, m, d) {
  y -= m <= 2 ? 1 : 0;
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const doy = Math.floor((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5) + d - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

export function civilFromDays(z) {
  z += 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365);
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const m = mp + (mp < 10 ? 3 : -9);
  return [yoe + era * 400 + (m <= 2 ? 1 : 0), m, d];
}

export const serialOf = (y, m, d, h = 0, n = 0, s = 0) => daysFromCivil(y, m, d) + EPOCH + (h * 3600 + n * 60 + s) / 86400;

/** The parts of a serial: [year, month, day, hour, minute, second, weekday (1 = Sunday)]. */
export function partsOf(serial) {
  let day = Math.floor(serial);
  let secs = Math.round((serial - day) * 86400);
  if (secs >= 86400) { day += 1; secs -= 86400; }
  const [y, m, d] = civilFromDays(day - EPOCH);
  const wd = ((((day - EPOCH) % 7) + 7 + 4) % 7) + 1; // 1970-01-01 was a Thursday (5)
  return [y, m, d, Math.floor(secs / 3600), Math.floor((secs % 3600) / 60), secs % 60, wd];
}

export const isLeap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
export const daysInMonth = (y, m) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

const pad = (n, w = 2) => String(n).padStart(w, '0');

/** ISO text of the content ("2026-11-07" or "2026-11-07T10:30:00"). */
export function isoOf(serial) {
  const [y, m, d, h, n, s] = partsOf(serial);
  const date = `${pad(y, 4)}-${pad(m)}-${pad(d)}`;
  return h || n || s ? `${date}T${pad(h)}:${pad(n)}:${pad(s)}` : date;
}

const ISO = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/;

/** The serial of an ISO text, or null. */
export function serialOfIso(text) {
  if (typeof text !== 'string') return null;
  const m = ISO.exec(text);
  if (!m) return null;
  const [y, mo, d, h, n, s] = m.slice(1).map(x => (x === undefined ? 0 : Number(x)));
  if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(y, mo) || h > 23 || n > 59 || s > 59) return null;
  return serialOf(y, mo, d, h, n, s);
}

export const isIsoDate = text => serialOfIso(text) !== null;

/** Today and now as serials (local time of the browser). */
export function nowSerial(clock = new Date()) {
  return serialOf(clock.getFullYear(), clock.getMonth() + 1, clock.getDate(), clock.getHours(), clock.getMinutes(), clock.getSeconds());
}
export const todaySerial = clock => Math.floor(nowSerial(clock));

const twoDigitYear = y => (y < 30 ? 2000 + y : 1900 + y);

/** Reads a date a person typed: ISO, 7.11.2026 (German), 11/7/2026 (English), with an optional time; "10:30" alone is a time. */
export function parseDate(text, lang = 'en', clock) {
  const s = String(text ?? '').trim();
  if (!s) return null;
  const iso = serialOfIso(s.replace(' ', 'T'));
  if (iso !== null) return iso;
  const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(s);
  if (timeOnly) return parseTime(timeOnly);
  const m = /^(\d{1,4})([./-])(\d{1,2})(?:\2(\d{1,4}))?\.?(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?)?$/i.exec(s);
  if (m) {
    let a = Number(m[1]);
    let b = Number(m[3]);
    let y = m[4] === undefined ? (clock ? clock.getFullYear() : new Date().getFullYear()) : Number(m[4]);
    let mo;
    let d;
    if (m[1].length === 4) { y = a; mo = b; d = Number(m[4] ?? 1); }
    else if (m[2] === '.' || (m[2] === '-' && lang === 'de')) { d = a; mo = b; }
    else if (lang === 'de') { d = a; mo = b; }
    else { mo = a; d = b; }
    if (m[4] !== undefined && m[4].length <= 2 && m[1].length !== 4) y = twoDigitYear(y);
    if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(y, mo)) return null;
    let time = 0;
    if (m[5] !== undefined) {
      const t = parseTime([null, m[5], m[6], m[7], m[8]]);
      if (t === null) return null;
      time = t;
    }
    return serialOf(y, mo, d) + time;
  }
  const named = parseNamedDate(s, clock);
  return named;
}

function parseTime(m) {
  let h = Number(m[1]);
  const n = Number(m[2]);
  const sec = Number(m[3] || 0);
  const ampm = (m[4] || '').toLowerCase();
  if (ampm === 'pm' && h < 12) h += 12;
  if (ampm === 'am' && h === 12) h = 0;
  if (h > 23 || n > 59 || sec > 59) return null;
  return (h * 3600 + n * 60 + sec) / 86400; // a time alone is the fraction of day 0 (30 Dec 1899), as in Access
}

const MONTHS = {
  en: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  de: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
};
const DAYS = {
  en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  de: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'],
};
export const monthName = (m, lang = 'en', short = false) => { const n = (MONTHS[lang] || MONTHS.en)[m - 1] || ''; return short ? n.slice(0, 3) : n; };
export const dayName = (wd, lang = 'en', short = false) => { const n = (DAYS[lang] || DAYS.en)[wd - 1] || ''; return short ? n.slice(0, lang === 'de' ? 2 : 3) : n; };

function parseNamedDate(s, clock) {
  // "7 Nov 2026", "Nov 7, 2026", "7. November 2026"
  const m = /^(\d{1,2})\.?\s+([A-Za-zäÄ]+)\.?,?\s+(\d{2,4})$/.exec(s) || null;
  const m2 = m ? null : /^([A-Za-zäÄ]+)\.?\s+(\d{1,2}),?\s+(\d{2,4})$/.exec(s);
  const parts = m ? [m[1], m[2], m[3]] : m2 ? [m2[2], m2[1], m2[3]] : null;
  if (!parts) return null;
  const word = parts[1].toLowerCase();
  let mo = 0;
  for (const list of Object.values(MONTHS)) {
    const i = list.findIndex(n => n.toLowerCase().startsWith(word.slice(0, 3)) && n.toLowerCase().startsWith(word));
    if (i >= 0) { mo = i + 1; break; }
  }
  if (!mo) return null;
  let y = Number(parts[2]);
  if (parts[2].length <= 2) y = twoDigitYear(y);
  const d = Number(parts[0]);
  if (d < 1 || d > daysInMonth(y, mo)) return null;
  void clock;
  return serialOf(y, mo, d);
}

/** Reads a number a person typed (1.234,56 in German, 1,234.56 in English; currency signs, spaces and % are allowed). */
export function parseNumber(text, lang = 'en') {
  if (typeof text === 'number') return Number.isFinite(text) ? text : null;
  let s = String(text ?? '').trim().replace(/[\s\u00a0\u202f€$£¥]|EUR|USD|CHF|GBP/g, '');
  if (!s) return null;
  let percent = false;
  if (s.endsWith('%')) { percent = true; s = s.slice(0, -1); }
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (lang === 'de') {
    if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (/^[+-]?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  } else {
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d*)?$/.test(s)) s = s.replace(/,/g, '');
    else if (s.includes(',') && !s.includes('.')) s = s.replace(',', '.');
  }
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) return null;
  let n = Number(s);
  if (!Number.isFinite(n)) return null;
  if (neg) n = -n;
  return percent ? n / 100 : n;
}

const YES = new Set(['yes', 'true', 'on', 'ja', 'wahr', 'ein', '-1', '1', 'x', '✓']);
const NO = new Set(['no', 'false', 'off', 'nein', 'falsch', 'aus', '0', '']);

export function parseBool(text) {
  if (typeof text === 'boolean') return text;
  if (typeof text === 'number') return text !== 0;
  const s = String(text ?? '').trim().toLowerCase();
  if (YES.has(s)) return true;
  if (NO.has(s)) return false;
  return null;
}

// ---- formatting ------------------------------------------------------------------------------------------------------------------

const numberFormats = new Map();
function nf(lang, options) {
  const key = lang + JSON.stringify(options);
  let f = numberFormats.get(key);
  if (!f) {
    try { f = new Intl.NumberFormat(lang === 'de' ? 'de-DE' : 'en-US', options); } catch { f = new Intl.NumberFormat('en-US', { maximumFractionDigits: 15 }); }
    numberFormats.set(key, f);
  }
  return f;
}

/** A number in "General Number": up to 15 significant digits, the decimal sign of the language, no grouping. */
export function generalNumber(n, lang = 'en') {
  if (!Number.isFinite(n)) return '';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-10)) return n.toExponential().replace('.', lang === 'de' ? ',' : '.').toUpperCase();
  const s = String(Number(n.toPrecision(15)));
  return lang === 'de' ? s.replace('.', ',') : s;
}

export function formatNumber(n, lang = 'en', format = '', decimals = null, currency = 'EUR') {
  if (n === null || n === undefined || n === '') return '';
  if (typeof n === 'boolean') n = n ? -1 : 0;
  if (n instanceof DateVal) n = n.s;
  if (typeof n !== 'number') return String(n);
  const f = String(format || '').toLowerCase();
  const dec = decimals === null || decimals === undefined || decimals === '' ? null : Number(decimals);
  switch (f) {
    case 'currency': case 'euro': {
      const d = dec ?? 2;
      return nf(lang, { style: 'currency', currency: f === 'euro' ? 'EUR' : currency || 'EUR', minimumFractionDigits: d, maximumFractionDigits: d }).format(n);
    }
    case 'fixed': { const d = dec ?? 2; return nf(lang, { useGrouping: false, minimumFractionDigits: d, maximumFractionDigits: d }).format(n); }
    case 'standard': { const d = dec ?? 2; return nf(lang, { minimumFractionDigits: d, maximumFractionDigits: d }).format(n); }
    case 'percent': { const d = dec ?? 2; return nf(lang, { style: 'percent', minimumFractionDigits: d, maximumFractionDigits: d }).format(n); }
    case 'scientific': return n.toExponential(dec ?? 2).replace('.', lang === 'de' ? ',' : '.').toUpperCase();
    case 'general number': case '':
      if (dec !== null) return nf(lang, { useGrouping: false, minimumFractionDigits: dec, maximumFractionDigits: dec }).format(n);
      return generalNumber(n, lang);
    default:
      if (/[0#]/.test(format)) return customNumber(n, format, lang);
      return generalNumber(n, lang);
  }
}

/** A custom number format of Access/VBA: "0.00", "#,##0", "0%", "00000" (sections for negative numbers with ";"). */
function customNumber(n, format, lang) {
  const sections = format.split(';');
  let fmt = sections[0];
  let value = n;
  if (n < 0 && sections[1]) { fmt = sections[1]; value = -n; }
  else if (n === 0 && sections[2]) return sections[2].replace(/"/g, '');
  const percent = fmt.includes('%');
  if (percent) value *= 100;
  const m = /([#0,]*)(\.([0#]*))?/.exec(fmt.replace(/"[^"]*"/g, ''));
  const intPart = m[1] || '0';
  const fracPart = m[3] || '';
  const minFrac = (fracPart.match(/0/g) || []).length;
  const maxFrac = fracPart.length;
  const minInt = (intPart.replace(/,/g, '').match(/0/g) || []).length;
  const grouping = intPart.includes(',');
  let body = nf(lang, { useGrouping: grouping, minimumFractionDigits: minFrac, maximumFractionDigits: maxFrac, minimumIntegerDigits: Math.max(1, minInt) }).format(Math.abs(value));
  if (minInt === 0 && Math.abs(value) < 1 && body.startsWith('0')) body = body.slice(1);
  const sign = value < 0 ? '-' : '';
  // keep literal text around the number ("0.00 kg", "\$0")
  const pre = fmt.slice(0, fmt.search(/[#0]/)).replace(/["\\]/g, '');
  const post = fmt.slice(fmt.search(/[#0][^#0,.]*$/) + 1).replace(/["\\]/g, '').replace(/^[#0,.]*/, '');
  return sign + pre + body + post;
}

export function shortDate(serial, lang) {
  const [y, m, d] = partsOf(serial);
  return lang === 'de' ? `${pad(d)}.${pad(m)}.${y}` : `${m}/${d}/${y}`;
}

export function longTime(serial, lang) {
  const [, , , h, n, s] = partsOf(serial);
  if (lang === 'de') return `${pad(h)}:${pad(n)}:${pad(s)}`;
  return `${((h + 11) % 12) + 1}:${pad(n)}:${pad(s)} ${h < 12 ? 'AM' : 'PM'}`;
}

export function shortTime(serial) {
  const [, , , h, n] = partsOf(serial);
  return `${pad(h)}:${pad(n)}`;
}

/** A date in a named Access format (General Date, Long Date, Medium Date, Short Date, Long Time, Medium Time, Short Time) or a custom one (dd.mm.yyyy …). */
export function formatDate(value, lang = 'en', format = '') {
  if (value === null || value === undefined || value === '') return '';
  const serial = value instanceof DateVal ? value.s : typeof value === 'number' ? value : serialOfIso(value);
  if (serial === null || !Number.isFinite(serial)) return String(value);
  const f = String(format || '').toLowerCase();
  const [y, m, d, h, n, s, wd] = partsOf(serial);
  const hasTime = h || n || s;
  const datePart = Math.floor(serial) !== serialOf(1899, 12, 30);
  switch (f) {
    case '': case 'general date':
      if (!hasTime) return shortDate(serial, lang);
      return datePart ? `${shortDate(serial, lang)} ${longTime(serial, lang)}` : longTime(serial, lang);
    case 'short date': return shortDate(serial, lang);
    case 'medium date': return lang === 'de' ? `${pad(d)}. ${monthName(m, lang, true)}. ${String(y).slice(-2)}` : `${pad(d)}-${monthName(m, lang, true)}-${String(y).slice(-2)}`;
    case 'long date': return lang === 'de' ? `${dayName(wd, lang)}, ${d}. ${monthName(m, lang)} ${y}` : `${dayName(wd, lang)}, ${monthName(m, lang)} ${d}, ${y}`;
    case 'long time': return longTime(serial, lang);
    case 'medium time': return lang === 'de' ? shortTime(serial) : `${((h + 11) % 12) + 1}:${pad(n)} ${h < 12 ? 'AM' : 'PM'}`;
    case 'short time': return shortTime(serial);
    default: return customDate(serial, format, lang);
  }
}

/** Custom date formats of VBA: yyyy yy q m mm mmm mmmm d dd ddd dddd h hh n nn s ss ww y AM/PM, "literal". */
export function customDate(serial, format, lang = 'en') {
  const [y, m, d, h, n, s, wd] = partsOf(serial);
  const ampm = /am\/pm/i.test(format);
  const hour = ampm ? ((h + 11) % 12) + 1 : h;
  const tokens = /("[^"]*"|\\.|yyyy|yy|y|q|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|ww|w|am\/pm|ttttt|c)/gi;
  let prevWasHour = false;
  return String(format).replace(tokens, tok => {
    const lower = tok.toLowerCase();
    let out;
    switch (lower) {
      case 'yyyy': out = String(y); break;
      case 'yy': out = String(y).slice(-2); break;
      case 'y': out = String(Math.floor(serial) - serialOf(y, 1, 1) + 1); break;
      case 'q': out = String(Math.ceil(m / 3)); break;
      case 'mmmm': out = monthName(m, lang); break;
      case 'mmm': out = monthName(m, lang, true); break;
      case 'mm': out = prevWasHour ? pad(n) : pad(m); break;
      case 'm': out = prevWasHour ? String(n) : String(m); break;
      case 'dddd': out = dayName(wd, lang); break;
      case 'ddd': out = dayName(wd, lang, true); break;
      case 'dd': out = pad(d); break;
      case 'd': out = String(d); break;
      case 'hh': out = pad(hour); break;
      case 'h': out = String(hour); break;
      case 'nn': out = pad(n); break;
      case 'n': out = String(n); break;
      case 'ss': out = pad(s); break;
      case 's': out = String(s); break;
      case 'ww': out = String(Math.ceil((Math.floor(serial) - serialOf(y, 1, 1) + partsOf(serialOf(y, 1, 1))[6]) / 7)); break;
      case 'w': out = String(wd); break;
      case 'am/pm': out = h < 12 ? (tok === 'am/pm' ? 'am' : 'AM') : (tok === 'am/pm' ? 'pm' : 'PM'); break;
      case 'ttttt': out = longTime(serial, lang); break;
      case 'c': out = formatDate(serial, lang, ''); break;
      default: out = tok.startsWith('"') ? tok.slice(1, -1) : tok.startsWith('\\') ? tok.slice(1) : tok;
    }
    prevWasHour = lower === 'h' || lower === 'hh';
    return out;
  });
}

/** Yes/No in a format: Yes/No, True/False, On/Off (in the language of the page). */
export function formatBool(value, lang = 'en', format = '') {
  if (value === null || value === undefined) return '';
  const v = typeof value === 'boolean' ? value : Number(value) !== 0;
  const f = String(format || 'yes/no').toLowerCase();
  const words = {
    en: { 'yes/no': ['Yes', 'No'], 'true/false': ['True', 'False'], 'on/off': ['On', 'Off'] },
    de: { 'yes/no': ['Ja', 'Nein'], 'true/false': ['Wahr', 'Falsch'], 'on/off': ['Ein', 'Aus'] },
  }[lang === 'de' ? 'de' : 'en'];
  const pair = words[f] || words['yes/no'];
  return v ? pair[0] : pair[1];
}

/** Text formats of Access: ">" upper case, "<" lower case. */
export function formatText(value, format = '') {
  if (value === null || value === undefined) return '';
  const s = String(value);
  if (format === '>') return s.toUpperCase();
  if (format === '<') return s.toLowerCase();
  return s;
}

/** A value of an expression or a field for the screen. type: the field type when known (datetime, currency, yesno …). */
export function formatValue(value, { type = null, format = '', decimals = null, lang = 'en', currency = 'EUR' } = {}) {
  if (value === null || value === undefined) return '';
  if (value instanceof DateVal) return formatDate(value, lang, format);
  if (type === 'datetime') return formatDate(value, lang, format);
  if (typeof value === 'boolean') {
    if (/^(yes\/no|true\/false|on\/off)$/i.test(format) || !format) return formatBool(value, lang, format);
    return formatNumber(value ? -1 : 0, lang, format, decimals, currency);
  }
  if (typeof value === 'number') {
    if (type === 'currency' && !format) return formatNumber(value, lang, 'Currency', decimals, currency);
    if (/^(short date|long date|medium date|general date|short time|long time|medium time)$/i.test(format)) return formatDate(value, lang, format);
    return formatNumber(value, lang, format, decimals, currency);
  }
  if (type === 'yesno') return formatBool(parseBool(value), lang, format);
  return formatText(value, format);
}

/** Banker's rounding of VBA (Round(2.5) = 2). */
export function roundHalfEven(n, digits = 0) {
  if (!Number.isFinite(n)) return n;
  const f = 10 ** digits;
  const x = n * f;
  const r = Math.round(x);
  const isHalf = Math.abs(x % 1) === 0.5 || Math.abs(Math.abs(x - Math.trunc(x)) - 0.5) < 1e-9;
  const v = isHalf ? 2 * Math.round(x / 2) : r;
  return v / f;
}
