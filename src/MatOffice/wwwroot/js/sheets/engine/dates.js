/* MatOffice Sheets – dates and times as Excel serial numbers (the 1900 date system; pure, no DOM).
   Serial 1 = 1900-01-01, 60 = the 29th of February 1900 that Excel keeps for Lotus, 61 = 1900-03-01; the fraction is the time of day. */

const DAY = 86400000;
const BASE = Date.UTC(1899, 11, 30);

export const isLeap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
export const daysInMonth = (y, m) => (m === 2 ? (isLeap(y) ? 29 : 28) : [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]);

/** The serial number of a date; months and days may overflow like DATE() (month 13 = January of the next year, day 0 = the last day of the month before). */
export function serial(y, m, d) {
  if (y === 1900 && m === 2 && d === 29) return 60;
  const first = new Date(0);
  first.setUTCFullYear(y, m - 1, 1); // not Date.UTC: it reads the years 0-99 as 1900-1999
  const days = Math.round((first.getTime() + (d - 1) * DAY - BASE) / DAY);
  return days < 61 ? days - 1 : days;
}

/** { y, m, d } of a serial number (its whole part); serial 0 is the "0th of January 1900" of Excel. */
export function ymd(value) {
  const s = Math.floor(value);
  if (s === 60) return { y: 1900, m: 2, d: 29 };
  if (s === 0) return { y: 1900, m: 1, d: 0 };
  const date = new Date(BASE + (s < 60 ? s + 1 : s) * DAY);
  return { y: date.getUTCFullYear(), m: date.getUTCMonth() + 1, d: date.getUTCDate() };
}

/** { h, mi, s, ms } of the time part of a serial number (rounded to the millisecond). */
export function hms(value) {
  let ms = Math.round((value - Math.floor(value)) * DAY);
  if (ms >= DAY) ms = DAY - 1;
  const h = Math.floor(ms / 3600000);
  ms -= h * 3600000;
  const mi = Math.floor(ms / 60000);
  ms -= mi * 60000;
  const s = Math.floor(ms / 1000);
  return { h, mi, s, ms: ms - s * 1000 };
}

/** The day of the week of a serial number: 0 = Sunday … 6 = Saturday (Excel: serial 1 was a Sunday). */
export function weekday(value) {
  return (Math.floor(value) + 6) % 7; // with the fictional 1900-02-29 this is Excel's own count
}

/** The serial number of a JS Date in local time (TODAY, NOW). */
export function fromJsDate(date) {
  const day = serial(date.getFullYear(), date.getMonth() + 1, date.getDate());
  return day + (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds() + date.getMilliseconds() / 1000) / 86400;
}

/** A JS Date (UTC) of a serial number. */
export function toJsDate(value) {
  const { y, m, d } = ymd(value);
  const t = hms(value);
  return new Date(Date.UTC(y, m - 1, d, t.h, t.mi, t.s, t.ms));
}

/** Month names in English and German (full and short) for parsing typed dates. */
export const MONTHS = {
  en: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  de: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
};
export const DAYS = {
  en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  de: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'],
};

function monthOf(word) {
  const w = word.toLowerCase().replace(/\.$/, '');
  if (w.length < 3) return 0;
  for (const list of [MONTHS.en, MONTHS.de]) {
    const i = list.findIndex(m => m.toLowerCase().startsWith(w) || (w === 'mrz' && m === 'März') || (w === 'mae' && m === 'März'));
    if (i >= 0) return i + 1;
  }
  return 0;
}

const valid = (y, m, d) => m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m) && y >= 1900 && y <= 9999;
const fullYear = y => (y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y);

/** A time "13:45", "13:45:30", "1:45 PM", "13:45:30,5" → fraction of a day; null otherwise. Hours may pass 24 ("36:00" = 1.5). */
export function parseTime(text) {
  const m = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2})(?:[.,](\d+))?)?\s*(am|pm|a|p)?$/i.exec(text.trim());
  if (!m) return null;
  let h = Number(m[1]);
  const mi = Number(m[2]);
  const s = m[3] ? Number(m[3]) + (m[4] ? Number('0.' + m[4]) : 0) : 0;
  if (mi > 59 || s >= 60) return null;
  if (m[5]) {
    if (h < 1 || h > 12) return null;
    const pm = m[5][0].toLowerCase() === 'p';
    h = (h % 12) + (pm ? 12 : 0);
  }
  return (h * 3600 + mi * 60 + s) / 86400;
}

/**
 * A typed date (with an optional time) → { value, kind: 'date'|'datetime'|'time', order } or null.
 * ISO 2026-01-31 always; with locale "de" 31.01.2026 / 31.1.26 / 31.1.; with "en" 1/31/2026 / 1/31; month names "31 Jan 2026", "Jan 31, 2026", "31. Januar 2026".
 */
export function parseDate(text, locale = 'en', currentYear = new Date().getFullYear()) {
  const t = text.trim();
  if (!t) return null;
  const time = parseTime(t);
  if (time !== null) return { value: time, kind: 'time' };
  // date and time separated by a space (or T for ISO)
  const split = /^(.+?)(?:[ T]+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:[.,]\d+)?)?\s*(?:am|pm|a|p)?))$/i.exec(t);
  const datePart = split ? split[1] : t;
  const timePart = split ? parseTime(split[2]) : null;
  if (split && timePart === null) return null;
  const day = parseDay(datePart.trim(), locale, currentYear);
  if (day === null) return null;
  return timePart === null ? { value: day.value, kind: 'date', order: day.order } : { value: day.value + timePart, kind: 'datetime', order: day.order };
}

function parseDay(t, locale, currentYear) {
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(t);
  if (m) return make(+m[1], +m[2], +m[3], 'iso');
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})?$/.exec(t);
  if (m && locale === 'de') return make(m[3] ? fullYear(+m[3]) : currentYear, +m[2], +m[1], 'dmy');
  m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2}|\d{4}))?$/.exec(t);
  if (m) {
    const y = m[3] ? fullYear(+m[3]) : currentYear;
    return locale === 'de' ? make(y, +m[2], +m[1], 'dmy') : make(y, +m[1], +m[2], 'mdy');
  }
  m = /^(\d{1,2})\.?\s*([A-Za-zÄÖÜäöü]{3,}\.?)\s*(\d{2}|\d{4})?$/.exec(t);
  if (m && monthOf(m[2])) return make(m[3] ? fullYear(+m[3]) : currentYear, monthOf(m[2]), +m[1], 'dmy');
  m = /^(\d{1,2})[-\s]([A-Za-zÄÖÜäöü]{3,})(?:[-\s](\d{2}|\d{4}))?$/.exec(t);
  if (m && monthOf(m[2])) return make(m[3] ? fullYear(+m[3]) : currentYear, monthOf(m[2]), +m[1], 'dmy');
  m = /^([A-Za-zÄÖÜäöü]{3,}\.?)\s+(\d{1,2}),?\s*(\d{4})?$/.exec(t);
  if (m && monthOf(m[1])) return make(m[3] ? +m[3] : currentYear, monthOf(m[1]), +m[2], 'mdy');
  return null;
}

function make(y, mo, d, order) {
  return valid(y, mo, d) ? { value: serial(y, mo, d), order } : null;
}
