// MatOffice Sheets – number formats: the text Excel shows for a value under a format code, typed input, decimals.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { formatValue, formatKind, changeDecimals, parseInput, formatForEdit, general, roundFixed, presets } from '../../../src/MatOffice/wwwroot/js/sheets/engine/numfmt.js';

const f = (v, code, locale = 'en') => formatValue(v, code, { locale }).text;

test('General', () => {
  assert.equal(f(0, undefined), '0');
  assert.equal(f(1234.5, 'General'), '1234.5');
  assert.equal(f(1 / 3), '0.333333333');
  assert.equal(f(-1 / 3), '-0.33333333');
  assert.equal(f(0.1 + 0.2), '0.3');
  assert.equal(f(123456789012), '1.23457E+11');
  assert.equal(f(12345678901), '12345678901');
  assert.equal(f(0.0001234), '0.0001234');
  assert.equal(f(1e-10), '1E-10');
  assert.equal(f(2.5, undefined, 'de'), '2,5');
  assert.equal(general(1234.5678, 'en', 6), '1234.6');
  assert.equal(general(123456789, 'en', 5), '1E+08');
  assert.equal(f('text', undefined), 'text');
  assert.equal(f(true, '0.00'), 'TRUE');
});

test('rounding of fixed decimals', () => {
  assert.equal(roundFixed(1.005, 2), '1.01');
  assert.equal(roundFixed(999.999, 2), '1000.00');
  assert.equal(roundFixed(0.5, 0), '1');
  assert.equal(roundFixed(0.006, 2), '0.01');
  assert.equal(roundFixed(0.0004, 2), '0.00');
  assert.equal(roundFixed(2.675, 2), '2.68');
});

test('number codes', () => {
  assert.equal(f(1.005, '0.00'), '1.01');
  assert.equal(f(-2.5, '0.00'), '-2.50');
  assert.equal(f(1234567, '#,##0'), '1,234,567');
  assert.equal(f(1234567, '#,##0', 'de'), '1.234.567');
  assert.equal(f(0, '#,##0'), '0');
  assert.equal(f(1234.5, '#,##0.00'), '1,234.50');
  assert.equal(f(-1234.5, '#,##0.00'), '-1,234.50');
  assert.equal(f(0.5, '#,##0.00'), '0.50');
  assert.equal(f(1234.5, '#,##0.00 "€"', 'de'), '1.234,50 €');
  assert.equal(f(5, '"$"#,##0.00'), '$5.00');
  assert.equal(f(-5, '"$"#,##0.00'), '-$5.00');
  assert.equal(f(5551234, '000-0000'), '555-1234');
  assert.equal(f(123, '00000'), '00123');
  assert.equal(f(1234567, '0.0,,"M"'), '1.2M');
  assert.equal(f(1234567, '#,##0,'), '1,235');
  assert.equal(f(2.5, '#.##'), '2.5');
  assert.equal(f(2, '0.0#'), '2.0');
  assert.equal(f(1.2, '?.??'), '1.2 ');
});

test('sections, colours and conditions', () => {
  const r = formatValue(-5, '#,##0;[Red]-#,##0');
  assert.equal(r.text, '-5');
  assert.equal(r.color, '#ff0000');
  assert.equal(f(-5, '0;(0)'), '(5)');
  assert.equal(f(0, '0;-0;"zero"'), 'zero');
  assert.equal(f('x', '0;-0;0;"Text: "@'), 'Text: x');
  assert.equal(f(150, '[>=100]"big";"small"'), 'big');
  assert.equal(f(50, '[>=100]"big";"small"'), 'small');
  assert.equal(f(1234.5, '[$€-407] #,##0.00'), '€ 1,234.50');
});

test('percent, scientific, fractions', () => {
  assert.equal(f(0.125, '0%'), '13%');
  assert.equal(f(0.125, '0.00%'), '12.50%');
  assert.equal(f(12345, '0.00E+00'), '1.23E+04');
  assert.equal(f(0.000123, '0.00E+00'), '1.23E-04');
  assert.equal(f(1.5, '# ?/?'), '1 1/2');
  assert.equal(f(0.25, '# ?/?').trim(), '1/4');
  assert.equal(f(3.14159, '# ??/??'), '3 14/99');
  assert.equal(f(0.5, '?/4').trim(), '2/4');
});

test('dates and times', () => {
  const d = 46053; // 2026-01-31, a Saturday
  assert.equal(f(d, 'dd.mm.yyyy'), '31.01.2026');
  assert.equal(f(d, 'm/d/yyyy'), '1/31/2026');
  assert.equal(f(d, 'dddd, mmmm d, yyyy'), 'Saturday, January 31, 2026');
  assert.equal(f(d, 'dddd, d. mmmm yyyy', 'de'), 'Samstag, 31. Januar 2026');
  assert.equal(f(d, 'mmm-yy'), 'Jan-26');
  assert.equal(f(d, 'ddd d mmm', 'de'), 'Sa 31 Jan');
  assert.equal(f(0.5, 'h:mm AM/PM'), '12:00 PM');
  assert.equal(f(0.75, 'hh:mm'), '18:00');
  assert.equal(f(1.5, '[h]:mm'), '36:00');
  assert.equal(f(1 / 1440, 'h:mm:ss'), '0:01:00');
  assert.equal(f(1 / 1440, 'mm:ss'), '01:00');
  assert.equal(f(d + 0.5, 'yyyy-mm-dd hh:mm'), '2026-01-31 12:00');
  assert.equal(f(1.5 / 86400, 'ss.0'), '01.5');
  assert.equal(f(0.99999999, 'hh:mm'), '00:00');
});

test('kinds of formats and the decimal buttons', () => {
  assert.equal(formatKind(undefined), 'general');
  assert.equal(formatKind('0.00'), 'number');
  assert.equal(formatKind('0%'), 'percent');
  assert.equal(formatKind('dd.mm.yyyy'), 'date');
  assert.equal(formatKind('h:mm'), 'time');
  assert.equal(formatKind('m/d/yyyy h:mm'), 'datetime');
  assert.equal(formatKind('"$"#,##0.00'), 'currency');
  assert.equal(formatKind(presets('en').accounting), 'accounting');
  assert.equal(formatKind('@'), 'text');
  assert.equal(formatKind('0.00E+00'), 'scientific');
  assert.equal(changeDecimals('0.00', 1), '0.000');
  assert.equal(changeDecimals('0.0', -1), '0');
  assert.equal(changeDecimals('#,##0', 1), '#,##0.0');
  assert.equal(changeDecimals('0%', 1), '0.0%');
  assert.equal(changeDecimals('#,##0.00 "€"', -1), '#,##0.0 "€"');
  assert.equal(changeDecimals('0.00E+00', -1), '0.0E+00');
  assert.equal(changeDecimals(undefined, 1, 2.5), '0.00');
  assert.equal(changeDecimals(undefined, -1, 2.25), '0.0');
  assert.equal(changeDecimals('dd.mm.yyyy', 1), 'dd.mm.yyyy');
});

test('typed input', () => {
  assert.deepEqual(parseInput('=SUM(A1:A3)'), { formula: 'SUM(A1:A3)' });
  assert.deepEqual(parseInput('42'), { value: 42 });
  assert.deepEqual(parseInput('-1.5'), { value: -1.5 });
  assert.deepEqual(parseInput('1,234'), { value: 1234, nf: '#,##0' });
  assert.deepEqual(parseInput('1.234,5', 'de'), { value: 1234.5, nf: '#,##0.00' });
  assert.deepEqual(parseInput('12,5', 'de'), { value: 12.5 });
  assert.deepEqual(parseInput('12%'), { value: 0.12, nf: '0%' });
  assert.deepEqual(parseInput('12,5 %', 'de'), { value: 0.125, nf: '0.0%' });
  assert.deepEqual(parseInput('$12.50'), { value: 12.5, nf: '"$"#,##0.00' });
  assert.deepEqual(parseInput('12 €', 'de'), { value: 12, nf: '#,##0 "€"' });
  assert.deepEqual(parseInput('TRUE'), { value: true });
  assert.deepEqual(parseInput('wahr', 'de'), { value: true });
  assert.deepEqual(parseInput('#N/A'), { error: '#N/A' });
  assert.deepEqual(parseInput("'007"), { value: '007' });
  assert.deepEqual(parseInput('1/31/2026'), { value: 46053, nf: 'm/d/yyyy' });
  assert.deepEqual(parseInput('31.1.2026', 'de'), { value: 46053, nf: 'dd.mm.yyyy' });
  assert.deepEqual(parseInput('13:30'), { value: 13.5 / 24, nf: 'h:mm' });
  assert.deepEqual(parseInput('1 1/2'), { value: 1.5, nf: '# ?/?' });
  assert.deepEqual(parseInput('hello'), { value: 'hello' });
  assert.deepEqual(parseInput(''), { empty: true });
  assert.deepEqual(parseInput('-A1'), { formula: '-A1' });
});

test('the text in the editor', () => {
  assert.equal(formatForEdit(46053, 'dd.mm.yyyy', 'de'), '31.01.2026');
  assert.equal(formatForEdit(46053, 'mmm-yy', 'en'), '1/31/2026');
  assert.equal(formatForEdit(46053.5, 'm/d/yyyy h:mm', 'en'), '1/31/2026 12:00:00 PM');
  assert.equal(formatForEdit(0.125, '0%'), '12.5%');
  assert.equal(formatForEdit(1234.5, '#,##0.00', 'de'), '1234,5');
  assert.equal(formatForEdit(0.1 + 0.2, undefined), '0.3');
  assert.equal(formatForEdit('007', undefined), "'007");
  assert.equal(formatForEdit('hello', undefined), 'hello');
  assert.equal(formatForEdit(true, undefined), 'TRUE');
});
