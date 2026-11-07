/* MatOffice Sheets – date and time functions (pure, no DOM). Dates are Excel serial numbers (1 = 1900-01-01, with Excel's 29 February 1900,
   2958465 = 9999-12-31); the fraction is the time of day. TODAY and NOW read ctx.now(). */
import { define, scalar, flat, deref, Errors } from '../lib.js';
import { FormulaError, toNumber, getLocale } from '../values.js';
import { serial, ymd, weekday, fromJsDate, parseDate, daysInMonth, isLeap } from '../dates.js';
import { lift } from './stat.js';

const isErr = v => v instanceof FormulaError;
const { NUM, VALUE } = Errors;
const MAX_DATE = 2958465;
/** Whether a serial number is a date Excel knows (0 … 9999-12-31 23:59:59). */
const valid = v => v >= 0 && v < MAX_DATE + 1;
const first = v => (Array.isArray(v) ? v[0][0] : v);

const date = (name, args, types, f) => define(name, {
  min: types.filter(t => !t.endsWith('?')).length, max: types.length, category: 'date', args, fn: scalar(types, f),
});

/** The serial number of day d of month m in year y; months and days overflow into the next ones (no range check). */
const dayOf = (y, m, d) => {
  const months = y * 12 + m - 1;
  const yy = Math.floor(months / 12);
  return serial(yy, months - yy * 12 + 1, 1) + d - 1;
};

/** DATE like Excel: years 0 … 1899 mean 1900 … 3799; months and days overflow; a result outside 1900-01-00 … 9999-12-31 is #NUM!. */
export function makeDate(y, m, d) {
  let year = Math.trunc(y);
  if (year < 0 || year >= 10000) return NUM;
  if (year < 1900) year += 1900;
  const s = dayOf(year, Math.trunc(m), Math.trunc(d));
  return s >= 0 && s <= MAX_DATE ? s : NUM;
}

date('DATE', 'year, month, day', ['number', 'number', 'number'], makeDate);
date('TIME', 'hour, minute, second', ['number', 'number', 'number'], (h, m, s) => {
  const [a, b, c] = [h, m, s].map(Math.trunc);
  if (a > 32767 || b > 32767 || c > 32767) return NUM;
  const t = a * 3600 + b * 60 + c;
  return t < 0 ? NUM : (t % 86400) / 86400;
});

define('TODAY', { min: 0, max: 0, category: 'date', volatile: true, args: '', fn: (args, ctx) => Math.floor(fromJsDate(ctx.now())) });
define('NOW', { min: 0, max: 0, category: 'date', volatile: true, args: '', fn: (args, ctx) => fromJsDate(ctx.now()) });

// ---- parts of a date and a time --------------------------------------------------------------------------------------------------

const part = (name, f) => date(name, 'serial_number', ['number'], v => (valid(v) ? f(v) : NUM));
/** The seconds of the day, rounded to the second like Excel's HOUR/MINUTE/SECOND. */
const seconds = v => Math.round((v - Math.floor(v)) * 86400) % 86400;

part('YEAR', v => ymd(v).y);
part('MONTH', v => ymd(v).m);
part('DAY', v => ymd(v).d);
part('HOUR', v => Math.floor(seconds(v) / 3600));
part('MINUTE', v => Math.floor(seconds(v) / 60) % 60);
part('SECOND', v => seconds(v) % 60);

/** The first day of the week (0 = Sunday) of the return types of WEEKDAY and WEEKNUM. */
const WEEK_START = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };

date('WEEKDAY', 'serial_number, [return_type]', ['number', 'number?'], (v, t = 1) => {
  if (!valid(v)) return NUM;
  const type = Math.trunc(t);
  const wd = weekday(v);
  if (type === 3) return (wd + 6) % 7;
  const start = WEEK_START[type];
  return start === undefined ? NUM : ((wd - start + 7) % 7) + 1;
});

/** The ISO week: weeks start on Monday, week 1 holds the first Thursday of the year. */
function isoWeek(v) {
  const d = Math.floor(v);
  const thursday = d - ((weekday(d) + 6) % 7) + 3;
  return Math.floor((thursday - serial(ymd(thursday).y, 1, 1)) / 7) + 1;
}

/** WEEKNUM: week 1 holds the 1st of January; return_type gives the first day of the week (21: ISO weeks). */
date('WEEKNUM', 'serial_number, [return_type]', ['number', 'number?'], (v, t = 1) => {
  if (!valid(v)) return NUM;
  const type = Math.trunc(t);
  if (type === 21) return isoWeek(v);
  const start = WEEK_START[type];
  if (start === undefined) return NUM;
  const jan1 = serial(ymd(v).y, 1, 1);
  return Math.floor((Math.floor(v) - jan1 + ((weekday(jan1) - start + 7) % 7)) / 7) + 1;
});
date('ISOWEEKNUM', 'date', ['number'], v => (valid(v) ? isoWeek(v) : NUM));

// ---- months and differences ------------------------------------------------------------------------------------------------------

/** The date months after v (EDATE: the same day, at most the last of the month; EOMONTH: the last day of that month). */
function addMonths(v, months, endOfMonth) {
  if (!valid(v)) return NUM;
  const { y, m, d } = ymd(v);
  const total = y * 12 + m - 1 + Math.trunc(months);
  const yy = Math.floor(total / 12);
  const mm = total - yy * 12 + 1;
  if (yy < 1900 || yy > 9999) return NUM;
  const last = yy === 1900 && mm === 2 ? 29 : daysInMonth(yy, mm); // Excel's 1900 has a 29 February
  return serial(yy, mm, 1) + (endOfMonth ? last : Math.min(d, last)) - 1;
}

date('EDATE', 'start_date, months', ['number', 'number'], (v, m) => addMonths(v, m, false));
date('EOMONTH', 'start_date, months', ['number', 'number'], (v, m) => addMonths(v, m, true));

/** DATEDIF(start, end, unit): complete years "Y", months "M", days "D", and "MD", "YM", "YD" (days/months without the larger units). */
date('DATEDIF', 'start_date, end_date, unit', ['number', 'number', 'text'], (s, e, unit) => {
  if (!valid(s) || !valid(e)) return NUM;
  const a0 = Math.floor(s);
  const b0 = Math.floor(e);
  if (a0 > b0) return NUM;
  const a = ymd(a0);
  const b = ymd(b0);
  const beforeDay = b.d < a.d;
  switch (unit.toUpperCase()) {
    case 'Y': return b.y - a.y - (b.m < a.m || (b.m === a.m && beforeDay) ? 1 : 0);
    case 'M': return (b.y - a.y) * 12 + b.m - a.m - (beforeDay ? 1 : 0);
    case 'D': return b0 - a0;
    case 'MD': return beforeDay ? b0 - dayOf(b.y, b.m - 1, a.d) : b.d - a.d;
    case 'YM': return (((b.m - a.m - (beforeDay ? 1 : 0)) % 12) + 12) % 12;
    case 'YD': return b0 - dayOf(b.m > a.m || (b.m === a.m && !beforeDay) ? b.y : b.y - 1, a.m, a.d);
    default: return NUM;
  }
});

date('DAYS', 'end_date, start_date', ['number', 'number'], (e, s) => (valid(e) && valid(s) ? Math.floor(e) - Math.floor(s) : NUM));

/** DAYS360: the US (NASD) method like Excel, or the European one (method TRUE: the 31st is the 30th). */
date('DAYS360', 'start_date, end_date, [method]', ['number', 'number', 'bool?'], (s, e, european = false) => {
  if (!valid(s) || !valid(e)) return NUM;
  let a = Math.floor(s);
  let b = Math.floor(e);
  let sign = 1;
  if (european && b < a) { [a, b] = [b, a]; sign = -1; }
  const d1 = ymd(a);
  const d2 = ymd(b);
  let day1 = d1.d;
  let day2 = d2.d;
  if (day1 === 31) day1 = 30;
  else if (!european && d1.m === 2 && (day1 === 29 || (day1 === 28 && !isLeap(d1.y)))) day1 = 30;
  if (day2 === 31 && (european || day1 === 30)) day2 = 30;
  return sign * ((d2.y - d1.y) * 360 + (d2.m - d1.m) * 30 + day2 - day1);
});

/** YEARFRAC(start, end, [basis]): 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360 (Excel's own rules). */
date('YEARFRAC', 'start_date, end_date, [basis]', ['number', 'number', 'number?'], (s, e, b = 0) => {
  const basis = Math.trunc(b);
  if (!valid(s) || !valid(e) || basis < 0 || basis > 4) return NUM;
  let a0 = Math.floor(s);
  let b0 = Math.floor(e);
  if (a0 > b0) [a0, b0] = [b0, a0];
  if (a0 === b0) return 0;
  const a = ymd(a0);
  const z = ymd(b0);
  const days360 = (d1, d2) => ((z.y - a.y) * 360 + (z.m - a.m) * 30 + d2 - d1) / 360;
  switch (basis) {
    case 0: {
      const lastFeb = d => d.m === 2 && d.d === daysInMonth(d.y, 2);
      let d1 = a.d;
      let d2 = z.d;
      if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
      else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      else if (lastFeb(a) && lastFeb(z)) { d1 = 30; d2 = 30; }
      else if (lastFeb(a)) d1 = 30;
      return days360(d1, d2);
    }
    case 1: {
      const withinYear = a.y === z.y || (z.y === a.y + 1 && (a.m > z.m || (a.m === z.m && a.d >= z.d)));
      if (!withinYear) return (b0 - a0) / ((serial(z.y + 1, 1, 1) - serial(a.y, 1, 1)) / (z.y - a.y + 1));
      const leapDay = (isLeap(a.y) && a0 < serial(a.y, 3, 1)) || (isLeap(z.y) && b0 >= serial(z.y, 3, 1));
      const length = (a.y === z.y && isLeap(a.y)) || leapDay || (z.m === 2 && z.d === 29) ? 366 : 365;
      return (b0 - a0) / length;
    }
    case 2: return (b0 - a0) / 360;
    case 3: return (b0 - a0) / 365;
    default: return days360(Math.min(a.d, 30), Math.min(z.d, 30));
  }
});

// ---- working days ------------------------------------------------------------------------------------------------------------------

/** The weekend of the .INTL functions as 7 flags by weekday (0 = Sunday): 1 … 7 (two days), 11 … 17 (one day) or "0000011" (Monday first). */
function weekendOf(w) {
  if (w === undefined || w === null) return [true, false, false, false, false, false, true];
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w) || w === '1111111') return VALUE;
    return [0, 1, 2, 3, 4, 5, 6].map(wd => w[(wd + 6) % 7] === '1');
  }
  const n = toNumber(w);
  if (isErr(n)) return n;
  const k = Math.trunc(n);
  const flags = new Array(7).fill(false);
  if (k >= 1 && k <= 7) { flags[(k + 5) % 7] = true; flags[(k + 6) % 7] = true; }
  else if (k >= 11 && k <= 17) flags[k - 11] = true;
  else return NUM;
  return flags;
}

/** The holidays (whole serial numbers) of an argument; blank cells are skipped. */
function holidaysOf(h, ctx) {
  const set = new Set();
  if (h === undefined || h === null) return set;
  for (const v of flat(h, ctx)) {
    if (v === null) continue;
    const n = toNumber(v);
    if (isErr(n)) return n;
    if (!valid(n)) return NUM;
    set.add(Math.floor(n));
  }
  return set;
}

/** The working days from a to b, both counted (negative when b is before a). */
function workdays(a, b, weekend, holidays) {
  const [lo, hi] = a <= b ? [a, b] : [b, a];
  const weeks = Math.floor((hi - lo + 1) / 7);
  let n = weeks * weekend.filter(x => !x).length;
  for (let d = lo + weeks * 7; d <= hi; d++) if (!weekend[weekday(d)]) n++;
  for (const h of holidays) if (h >= lo && h <= hi && !weekend[weekday(h)]) n--;
  return a <= b ? n : -n;
}

/** The working day that is days working days after (before) start. */
function addWorkdays(start, days, weekend, holidays) {
  let d = start;
  let left = Math.trunc(days);
  const step = left < 0 ? -1 : 1;
  while (left !== 0) {
    d += step;
    if (d < 1 || d > MAX_DATE) return NUM;
    if (!weekend[weekday(d)] && !holidays.has(d)) left -= step;
  }
  return d;
}

/** NETWORKDAYS(.INTL) and WORKDAY(.INTL): f(start, number, weekend, holidays); start and number may be arrays. */
function working(f) {
  return ([start, n, w, h], ctx) => {
    const weekend = weekendOf(w === undefined || w === null ? w : first(deref(w, ctx)));
    if (isErr(weekend)) return weekend;
    const holidays = holidaysOf(h, ctx);
    if (isErr(holidays)) return holidays;
    return lift([start, n], ctx, ([a, b]) => {
      const x = toNumber(a);
      if (isErr(x)) return x;
      const y = toNumber(b);
      if (isErr(y)) return y;
      return valid(x) ? f(Math.floor(x), y, weekend, holidays) : NUM;
    });
  };
}

const networkdays = working((a, b, weekend, holidays) => (valid(b) ? workdays(a, Math.floor(b), weekend, holidays) : NUM));
const workday = working(addWorkdays);

define('NETWORKDAYS', { min: 2, max: 3, category: 'date', args: 'start_date, end_date, [holidays]', fn: ([s, e, h], ctx) => networkdays([s, e, undefined, h], ctx) });
define('NETWORKDAYS.INTL', { min: 2, max: 4, category: 'date', args: 'start_date, end_date, [weekend], [holidays]', fn: networkdays });
define('WORKDAY', { min: 2, max: 3, category: 'date', args: 'start_date, days, [holidays]', fn: ([s, d, h], ctx) => workday([s, d, undefined, h], ctx) });
define('WORKDAY.INTL', { min: 2, max: 4, category: 'date', args: 'start_date, days, [weekend], [holidays]', fn: workday });

// ---- text as dates -------------------------------------------------------------------------------------------------------------

/** DATEVALUE/TIMEVALUE: a date (time) typed as text in the locale; dates without a year are in the current year. */
function fromText(f) {
  return (args, ctx) => scalar(['any'], v => {
    if (isErr(v)) return v;
    if (typeof v !== 'string') return VALUE;
    const p = parseDate(v, getLocale(), ctx.now().getFullYear());
    return p ? f(p) : VALUE;
  })(args, ctx);
}

define('DATEVALUE', { min: 1, max: 1, category: 'date', args: 'date_text', fn: fromText(p => (p.kind === 'time' ? VALUE : Math.floor(p.value))) });
define('TIMEVALUE', { min: 1, max: 1, category: 'date', args: 'time_text', fn: fromText(p => p.value - Math.floor(p.value)) });
