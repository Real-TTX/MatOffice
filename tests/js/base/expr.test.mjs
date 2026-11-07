// node --test tests/js/base — the expressions, values and criteria of MatOffice Base (pure modules of wwwroot/js/base).
import test from 'node:test';
import assert from 'node:assert/strict';
import { Parser, compile, evaluate, exprToText, BaseError, likeRegex } from '../../../src/MatOffice/wwwroot/js/base/expr.js';
import { parseCriteria, criteriaText, tidyCriteria, ruleHolds } from '../../../src/MatOffice/wwwroot/js/base/criteria.js';
import { serialOf, isoOf, parseDate, parseNumber, formatValue, formatDate, DateVal, partsOf, roundHalfEven } from '../../../src/MatOffice/wwwroot/js/base/values.js';

const clock = new Date(2026, 10, 7, 14, 30, 0); // 7 Nov 2026 14:30
const ev = (text, vars = {}, lang = 'en') => evaluate(text, {
  lang, now: () => clock,
  resolve: parts => { const k = parts[parts.length - 1].toLowerCase(); return k in vars ? () => vars[k] : null; },
});

test('dates are serial numbers like Access', () => {
  assert.equal(serialOf(1899, 12, 30), 0);
  assert.equal(serialOf(2026, 11, 7), 46333);
  assert.equal(isoOf(serialOf(2026, 11, 7, 10, 30)), '2026-11-07T10:30:00');
  assert.deepEqual(partsOf(serialOf(2024, 2, 29)).slice(0, 3), [2024, 2, 29]);
  assert.equal(parseDate('7.11.2026', 'de'), 46333);
  assert.equal(parseDate('11/7/2026', 'en'), 46333);
  assert.equal(parseDate('2026-11-07', 'de'), 46333);
  assert.equal(parseDate('31.2.2026', 'de'), null);
  assert.equal(formatDate(46333, 'de', 'Short Date'), '07.11.2026');
  assert.equal(formatDate(46333, 'en', 'Long Date'), 'Saturday, November 7, 2026');
  assert.equal(formatDate(46333, 'en', 'dd/mm/yyyy'), '07/11/2026');
});

test('numbers in German and English', () => {
  assert.equal(parseNumber('1.234,56', 'de'), 1234.56);
  assert.equal(parseNumber('1,234.56', 'en'), 1234.56);
  assert.equal(parseNumber('12,5', 'de'), 12.5);
  assert.equal(parseNumber('€ 3,00', 'de'), 3);
  assert.equal(parseNumber('abc', 'en'), null);
  assert.equal(formatValue(1234.5, { type: 'currency', lang: 'de', currency: 'EUR' }).replace(/\s/g, ' '), '1.234,50 €');
  assert.equal(formatValue(0.25, { format: 'Percent', lang: 'en' }), '25.00%');
  assert.equal(formatValue(true, { type: 'yesno', lang: 'de' }), 'Ja');
  assert.equal(roundHalfEven(2.5), 2);
  assert.equal(roundHalfEven(3.5), 4);
});

test('operators and precedence', () => {
  assert.equal(ev('1 + 2 * 3'), 7);
  assert.equal(ev('-2 ^ 2'), -4);
  assert.equal(ev('7 \\ 2'), 3);
  assert.equal(ev('7 Mod 3'), 1);
  assert.equal(ev('"a" & 1 & Null'), 'a1');
  assert.equal(ev('1 + Null'), null);
  assert.equal(ev('Null = Null'), null);
  assert.equal(ev('"Berlin" = "berlin"'), true);
  assert.equal(ev('3 Between 1 And 5'), true);
  assert.equal(ev('3 Not Between 1 And 2'), true);
  assert.equal(ev('"Anna" Like "A*"'), true);
  assert.equal(ev('"B12" Like "[A-C]##"'), true);
  assert.equal(ev('"x" In ("a","X")'), true);
  assert.equal(ev('Null Is Null'), true);
  assert.equal(ev('Not True Or False'), false);
  assert.equal(ev('False And Null'), false);
  assert.equal(ev('True Or Null'), true);
  assert.equal(ev('[Price] * [Qty]', { price: 2.5, qty: 4 }), 10);
});

test('functions of Access', () => {
  assert.equal(ev('IIf(1 > 2, "a", "b")'), 'b');
  assert.equal(ev('Nz(Null, 5)'), 5);
  assert.equal(ev('Left("Berlin", 3) & Right("Berlin", 2) & Mid("Berlin", 2, 2)'), 'Beriner');
  assert.equal(ev('Len("abc")'), 3);
  assert.equal(ev('UCase("abc") & LCase("DEF") & Trim("  x  ")'), 'ABCdefx');
  assert.equal(ev('InStr("Hello", "l")'), 3);
  assert.equal(ev('Replace("a-b-c", "-", "+")'), 'a+b+c');
  assert.equal(ev('Year(#11/7/2026#)'), 2026);
  assert.equal(ev('Month(#2026-11-07#) & "/" & Day(#7.11.2026#)'), '11/7');
  assert.equal(ev('Date()').s, 46333);
  assert.equal(isoOf(ev('DateAdd("m", 1, #1/31/2026#)').s), '2026-02-28');
  assert.equal(ev('DateDiff("d", #1/1/2026#, #3/1/2026#)'), 59);
  assert.equal(ev('DateDiff("yyyy", #12/31/2025#, #1/1/2026#)'), 1);
  assert.equal(ev('Round(2.345, 2)'), 2.34);
  assert.equal(ev('Int(-2.5) & Fix(-2.5) & Abs(-3)'), '-3-23');
  assert.equal(ev('Format(1234.5, "Currency")', {}, 'en'), '€1,234.50');
  assert.equal(ev('Format(#11/7/2026#, "yyyy-mm-dd")'), '2026-11-07');
  assert.equal(ev('Format(0.5, "Percent")'), '50.00%');
  assert.equal(ev('Switch(1 > 2, "a", 2 > 1, "b")'), 'b');
  assert.equal(ev('Choose(2, "x", "y", "z")'), 'y');
  assert.equal(ev('Val("12abc")'), 12);
  assert.equal(ev('CInt("3,5")', {}, 'de'), 4);
  assert.equal(ev('#11/7/2026# + 1').s, 46334);
  assert.equal(ev('#11/8/2026# - #11/7/2026#'), 1);
});

test('errors are coded', () => {
  assert.throws(() => Parser.expression('1 +'), e => e instanceof BaseError && e.code === 'Syntax');
  assert.throws(() => ev('Foo(1)'), e => e.code === 'UnknownFunction');
  assert.throws(() => ev('[Nope] + 1'), e => e.code === 'UnknownName');
  assert.throws(() => ev('1 / 0'), e => e.code === 'DivZero');
  assert.throws(() => ev('"abc" * 2'), e => e.code === 'TypeMismatch');
});

test('expressions are written back the Access way', () => {
  const t = s => exprToText(Parser.expression(s));
  assert.equal(t('[Price]*[Qty]'), '[Price] * [Qty]');
  assert.equal(t('(1+2)*3'), '(1 + 2) * 3');
  assert.equal(t('a - (b - c)'), '[a] - ([b] - [c])');
  assert.equal(t('#2026-11-07#'), '#11/7/2026#');
  assert.equal(t('x Not Like "A*"'), '[x] Not Like "A*"');
  assert.equal(t('iif(x>1,"a","b")'), 'IIf([x] > 1,"a","b")');
});

test('Like patterns', () => {
  assert.ok(likeRegex('A*').test('anna'));
  assert.ok(!likeRegex('A?').test('Abc'));
  assert.ok(likeRegex('[!a]*').test('Bob'));
  assert.ok(likeRegex('*.txt').test('notes.txt'));
});

test('criteria of the grid', () => {
  const col = { t: 'id', parts: ['Customers', 'City'] };
  const round = text => criteriaText(parseCriteria(text, col), col);
  assert.equal(round('"Berlin"'), '"Berlin"');
  assert.equal(round('Berlin'), '"Berlin"');
  assert.equal(round('>10 And <20'), '>10 And <20');
  assert.equal(round('"Berlin" Or "Paris"'), '"Berlin" Or "Paris"');
  assert.equal(round('Between 1 And 5'), 'Between 1 And 5');
  assert.equal(round('Like "B*"'), 'Like "B*"');
  assert.equal(round('B*'), 'Like "B*"');
  assert.equal(round('Is Null'), 'Is Null');
  assert.equal(round('Is Not Null'), 'Is Not Null');
  assert.equal(round('Not "Berlin"'), 'Not "Berlin"');
  assert.equal(round('In("a","b")'), 'In ("a","b")');
  assert.equal(round('<Date()'), '<Date()');
  assert.equal(tidyCriteria('1.1.2026'), '#1.1.2026#');
  assert.equal(round('[Enter a city]'), '[Enter a city]');
});

test('validation rules', () => {
  const holds = (rule, v) => ruleHolds(rule, v, compile, { lang: 'en' });
  assert.equal(holds('>0', 5), true);
  assert.equal(holds('>0', -1), false);
  assert.equal(holds('>0', null), true);
  assert.equal(holds('Between 1 And 10', 11), false);
  assert.equal(holds('Like "A*" Or Is Null', 'Bob'), false);
  assert.equal(holds('<=Date()', new DateVal(serialOf(2000, 1, 1))), true);
  assert.equal(holds('"M" Or "F"', 'f'), true);
});
