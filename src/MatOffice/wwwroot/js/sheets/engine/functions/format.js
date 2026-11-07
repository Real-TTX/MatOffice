/* MatOffice Sheets – TEXT and DOLLAR: numbers as text with a number format code (pure, no DOM). */
import { define, scalar, Errors } from '../lib.js';
import { formatValue } from '../numfmt.js';
import { getLocale, parseNumber } from '../values.js';
import { roundTo } from './math.js';

/** German format codes (TT.MM.JJJJ) as the English ones the formatter reads; quoted text stays. */
function englishCode(code) {
  if (getLocale() !== 'de' || !/[JT]/.test(code)) return code;
  return code.replace(/("[^"]*")|J/g, (m, q) => q || 'y').replace(/("[^"]*")|T/g, (m, q) => q || 'd');
}

define('TEXT', {
  min: 2, max: 2, category: 'text', args: 'value, format_text',
  fn: scalar(['any', 'text'], (value, code) => {
    if (value instanceof Object && value.code) return value;
    let v = value;
    if (typeof v === 'string') { const n = parseNumber(v); if (n !== null) v = n; }
    if (v === null) v = 0;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return formatValue(v, englishCode(code), { locale: getLocale() }).text;
  }),
});

define('DOLLAR', {
  min: 1, max: 2, category: 'text', args: 'number, [decimals]',
  fn: scalar(['number', 'number?'], (n, d = 2) => {
    const digits = Math.trunc(d);
    if (digits > 127) return Errors.VALUE;
    const decimals = Math.max(0, digits);
    const v = digits >= 0 ? roundTo(n, digits) : Math.round(n / 10 ** -digits) * 10 ** -digits;
    const body = '#,##0' + (decimals ? '.' + '0'.repeat(decimals) : '');
    const code = getLocale() === 'de' ? `${body} "€";-${body} "€"` : `"$"${body};("$"${body})`;
    return formatValue(v, code, { locale: getLocale() }).text;
  }),
});
