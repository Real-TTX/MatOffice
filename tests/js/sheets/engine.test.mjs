// MatOffice Sheets – the formula engine: tokens, parser, operators, recalculation, circles, spills.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { workbook, evaluate } from './harness.mjs';
import { tokenize, printTokens, toCanonical, toLocal } from '../../../src/MatOffice/wwwroot/js/sheets/engine/lexer.js';
import { parse, FormulaSyntaxError } from '../../../src/MatOffice/wwwroot/js/sheets/engine/parser.js';
import { colName, colIndex, parseArea, areaName, parseReference, sheetPrefix } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';
import { serial, ymd, weekday, parseDate, parseTime } from '../../../src/MatOffice/wwwroot/js/sheets/engine/dates.js';
import { parseNumber, numberToText, criteria, setLocale, compare } from '../../../src/MatOffice/wwwroot/js/sheets/engine/values.js';

test('addresses', () => {
  assert.equal(colName(0), 'A');
  assert.equal(colName(25), 'Z');
  assert.equal(colName(26), 'AA');
  assert.equal(colName(16383), 'XFD');
  assert.equal(colIndex('xfd'), 16383);
  assert.equal(colIndex('XFE'), -1);
  assert.deepEqual(parseArea('$B$2:a1'), { r1: 0, c1: 0, r2: 1, c2: 1 });
  assert.equal(areaName(parseArea('C:A')), 'A:C');
  assert.equal(areaName(parseArea('3:2')), '2:3');
  assert.deepEqual(parseReference("'It''s'!A1"), { sheet: "It's", area: { r1: 0, c1: 0, r2: 0, c2: 0 } });
  assert.equal(sheetPrefix('Sheet1'), 'Sheet1!');
  assert.equal(sheetPrefix('My sheet'), "'My sheet'!");
  assert.equal(sheetPrefix('A1'), "'A1'!");
});

test('dates are Excel serial numbers', () => {
  assert.equal(serial(1900, 1, 1), 1);
  assert.equal(serial(1900, 2, 28), 59);
  assert.equal(serial(1900, 2, 29), 60);
  assert.equal(serial(1900, 3, 1), 61);
  assert.equal(serial(2026, 1, 31), 46053);
  assert.equal(serial(2026, 13, 1), serial(2027, 1, 1));
  assert.equal(serial(2026, 3, 0), serial(2026, 2, 28));
  assert.deepEqual(ymd(46053), { y: 2026, m: 1, d: 31 });
  assert.deepEqual(ymd(60), { y: 1900, m: 2, d: 29 });
  assert.equal(weekday(46053), 6); // a Saturday
  assert.equal(weekday(1), 0); // Excel's Sunday
  assert.equal(parseDate('2026-01-31').value, 46053);
  assert.equal(parseDate('31.01.2026', 'de').value, 46053);
  assert.equal(parseDate('1/31/2026', 'en').value, 46053);
  assert.equal(parseDate('31 Jan 2026').value, 46053);
  assert.equal(parseDate('31. Januar 2026', 'de').value, 46053);
  assert.equal(parseDate('2026-01-31 12:00').value, 46053.5);
  assert.equal(parseDate('31.02.2026', 'de'), null);
  assert.equal(parseTime('12:00'), 0.5);
  assert.equal(parseTime('6:00 PM'), 0.75);
});

test('numbers as text and text as numbers', () => {
  setLocale('en');
  assert.equal(numberToText(0.1 + 0.2), '0.3');
  assert.equal(numberToText(1 / 3), '0.333333333333333');
  assert.equal(numberToText(1e21), '1E+21');
  assert.equal(numberToText(-12.5), '-12.5');
  assert.equal(parseNumber('1,234.5'), 1234.5);
  assert.equal(parseNumber('12%'), 0.12);
  assert.equal(parseNumber('$5'), 5);
  assert.equal(parseNumber('(4)'), -4);
  assert.equal(parseNumber('abc'), null);
  setLocale('de');
  assert.equal(parseNumber('1.234,5'), 1234.5);
  assert.equal(parseNumber('12,5 €'), 12.5);
  assert.equal(numberToText(2.5), '2,5');
  setLocale('en');
});

test('comparisons and criteria', () => {
  assert.equal(compare(1, 'a'), -1);
  assert.equal(compare('a', true), -1);
  assert.equal(compare('abc', 'ABC'), 0);
  assert.equal(compare(null, 0), 0);
  assert.equal(compare(null, ''), 0);
  assert.ok(criteria('>5')(6));
  assert.ok(!criteria('>5')('6'));
  assert.ok(criteria(5)('5'));
  assert.ok(criteria('a*')('Apple'));
  assert.ok(criteria('?b')('ab'));
  assert.ok(!criteria('?b')('abc'));
  assert.ok(criteria('<>x')('y'));
  assert.ok(criteria('')(null));
  assert.ok(criteria('<>')('a'));
  assert.ok(!criteria('<>')(null));
  assert.ok(criteria('~*')('*'));
  assert.ok(criteria('TRUE')(true));
});

test('tokens keep the text and the locale changes the separators', () => {
  const t = tokenize('SUM( A1:B2 , 1.5 )');
  assert.equal(printTokens(t), 'SUM( A1:B2 , 1.5 )');
  assert.equal(printTokens(t, 'de'), 'SUM( A1:B2 ; 1,5 )');
  assert.equal(toCanonical('ROUND(A1;2)', 'de'), 'ROUND(A1,2)');
  assert.equal(toCanonical('ROUND(1,25;1)', 'de'), 'ROUND(1.25,1)');
  assert.equal(toCanonical('SUM(A1,B1)', 'de'), 'SUM(A1,B1)');
  assert.equal(toCanonical('IF(A1>0,5;"x";"y")', 'de'), 'IF(A1>0.5,"x","y")');
  assert.equal(toLocal('IF(A1>0.5,"a,b",{1,2;3,4})', 'de'), 'IF(A1>0,5;"a,b";{1.2;3.4})');
  assert.equal(toCanonical('{1,5.2;3.4}', 'de'), '{1.5,2;3,4}');
  const refs = tokenize("'My sheet'!$A$1:B2+Sheet2!C:C+3:4+#REF!+LOG10(100)").filter(x => x.type === 'ref');
  assert.deepEqual(refs.map(r => [r.sheet, r.kind]), [['My sheet', 'area'], ['Sheet2', 'cols'], [null, 'rows'], [null, 'bad']]);
  assert.equal(tokenize('_xlfn.XLOOKUP(1,A:A,B:B)')[0].name, 'XLOOKUP');
});

test('the parser follows the precedence of Excel', () => {
  assert.equal(evaluate('-2^2'), 4);
  assert.equal(evaluate('2^3^2'), 64);
  assert.equal(evaluate('1+2*3'), 7);
  assert.equal(evaluate('(1+2)*3'), 9);
  assert.equal(evaluate('50%*2'), 1);
  assert.equal(evaluate('1+2&"x"'), '3x');
  assert.equal(evaluate('1+1=2'), true);
  assert.equal(evaluate('"a"<"B"'), true);
  assert.equal(evaluate('0.1+0.2=0.3'), true);
  assert.equal(evaluate('0.1+0.2-0.3'), 0);
  assert.equal(evaluate('"3"+4'), 7);
  assert.equal(evaluate('"x"+4'), '#VALUE!');
  assert.equal(evaluate('1/0'), '#DIV/0!');
  assert.equal(evaluate('0^0'), '#NUM!');
  assert.equal(evaluate('TRUE+TRUE'), 2);
  assert.equal(evaluate('NOPE(1)'), '#NAME?');
  assert.equal(evaluate('nope'), '#NAME?');
  assert.throws(() => parse('1+'), FormulaSyntaxError);
  assert.throws(() => parse('SUM(1'), FormulaSyntaxError);
  assert.throws(() => parse('"abc'), FormulaSyntaxError);
  assert.deepEqual(evaluate('{1,2;3,4}*10'), [[10, 20], [30, 40]]);
});

test('references, ranges and other sheets', () => {
  const wb = workbook({
    Sheet1: { A1: 1, A2: 2, A3: 3, B1: '=SUM(A1:A3)', B2: '=A1+A2', B3: "='Data sheet'!A1*2", B4: '=SUM(A:A)', B5: '=C9' },
    'Data sheet': { A1: 21 },
  });
  assert.equal(wb.get('B1'), 6);
  assert.equal(wb.get('B2'), 3);
  assert.equal(wb.get('B3'), 42);
  assert.equal(wb.get('B4'), 6);
  assert.equal(wb.get('B5'), 0); // an empty cell reads as 0
  wb.set({ A2: 10 });
  assert.equal(wb.get('B1'), 14);
  assert.equal(wb.get('B2'), 11);
  assert.equal(wb.get('B4'), 14);
  wb.set({ A1: 1 }, 'Data sheet');
  assert.equal(wb.get('B3'), 2);
});

test('recalculation follows chains in both directions', () => {
  const cells = {};
  for (let r = 1; r <= 3000; r++) cells['A' + r] = r === 1 ? 1 : `=A${r - 1}+1`;
  for (let r = 1; r <= 3000; r++) cells['B' + r] = r === 3000 ? 1 : `=B${r + 1}+1`;
  const wb = workbook({ Sheet1: cells });
  assert.equal(wb.get('A3000'), 3000);
  assert.equal(wb.get('B1'), 3000);
  wb.set({ A1: 10, B3000: 10 });
  assert.equal(wb.get('A3000'), 3009);
  assert.equal(wb.get('B1'), 3009);
});

test('a circular reference gives #CIRC!', () => {
  const wb = workbook({ Sheet1: { A1: '=B1+1', B1: '=A1+1', C1: '=C1', D1: '=A1*2', E1: 5 } });
  assert.equal(wb.get('A1'), '#CIRC!');
  assert.equal(wb.get('B1'), '#CIRC!');
  assert.equal(wb.get('C1'), '#CIRC!');
  assert.equal(wb.get('D1'), '#CIRC!');
  wb.set({ B1: '=E1' });
  assert.equal(wb.get('A1'), 6);
  assert.equal(wb.get('D1'), 12);
});

test('arrays spill and are blocked by values', () => {
  const wb = workbook({ Sheet1: { A1: 1, A2: 2, A3: 3, B1: '=A1:A3*10', C1: '=SUM(B1:B3)', D1: '=B2' } });
  assert.equal(wb.get('B1'), 10);
  assert.equal(wb.get('B2'), 20);
  assert.equal(wb.get('B3'), 30);
  assert.equal(wb.get('C1'), 60);
  assert.equal(wb.get('D1'), 20);
  wb.set({ B3: 'x' });
  assert.equal(wb.get('B1'), '#SPILL!');
  assert.equal(wb.get('B2'), null);
  assert.equal(wb.get('D1'), 0);
  assert.equal(wb.get('C1'), '#SPILL!');
  wb.set({ B3: null });
  assert.equal(wb.get('B2'), 20);
  assert.equal(wb.get('C1'), 60);
  wb.set({ A2: 5 });
  assert.equal(wb.get('D1'), 50);
});

test('defined names', () => {
  const wb = workbook({ Sheet1: { A1: 0.25, A2: 100, B1: '=Rate*A2', B2: '=SUM(Amounts)' } }, [{ name: 'Rate', ref: 'Sheet1!$A$1' }, { name: 'Amounts', ref: 'Sheet1!$A$1:$A$2' }]);
  assert.equal(wb.get('B1'), 25);
  assert.equal(wb.get('B2'), 100.25);
  wb.set({ A1: 0.5 });
  assert.equal(wb.get('B1'), 50);
});

test('German formulas are read with semicolons and decimal commas', () => {
  assert.equal(evaluate(toCanonical('ROUND(2,345;2)', 'de')), 2.35);
});
