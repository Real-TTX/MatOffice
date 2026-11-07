// MatOffice Sheets – information functions: they see blanks, "" and errors as they are.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate, workbook } from './harness.mjs';

const check = (cases, cells = {}) => { for (const [f, want] of cases) assert.deepEqual(evaluate(f, cells), want, f); };

const DATA = { A1: 1, A2: 2, A3: 3, A4: 'x', A5: true, D1: '=""', E1: '#N/A' };

test('the IS functions', () => {
  check([
    ['ISBLANK(A6)', true],
    ['ISBLANK(D1)', false],
    ['ISBLANK(A1)', false],
    ['ISBLANK(A5:A6)', [[false], [true]]],
    ['ISNUMBER(A1)', true],
    ['ISNUMBER("1")', false],
    ['ISNUMBER(A4)', false],
    ['ISTEXT(A4)', true],
    ['ISTEXT(A6)', false],
    ['ISTEXT(D1)', true],
    ['ISNONTEXT(A6)', true],
    ['ISNONTEXT(A4)', false],
    ['ISNONTEXT(E1)', true],
    ['ISLOGICAL(A5)', true],
    ['ISLOGICAL(1)', false],
    ['ISERROR(E1)', true],
    ['ISERROR(1/0)', true],
    ['ISERROR(1)', false],
    ['ISERR(E1)', false],
    ['ISERR(1/0)', true],
    ['ISNA(E1)', true],
    ['ISNA(1/0)', false],
    ['ISEVEN(-1)', false],
    ['ISEVEN(2.5)', true],
    ['ISEVEN(A6)', true],
    ['ISEVEN("x")', '#VALUE!'],
    ['ISEVEN(TRUE)', '#VALUE!'],
    ['ISODD(5)', true],
    ['ISODD(-3)', true],
    ['ISFORMULA(D1)', true],
    ['ISFORMULA(A1)', false],
    ['ISFORMULA(1)', '#VALUE!'],
    ['ISREF(A1)', true],
    ['ISREF(A1:B2)', true],
    ['ISREF("A1")', false],
    ['ISREF(INDEX(A1:A3,2))', true],
    ['ISREF(1)', false],
  ], DATA);
});

test('NA, ERROR.TYPE, TYPE and N', () => {
  check([
    ['NA()', '#N/A'],
    ['ERROR.TYPE(#NULL!)', 1],
    ['ERROR.TYPE(1/0)', 2],
    ['ERROR.TYPE(#VALUE!)', 3],
    ['ERROR.TYPE(#REF!)', 4],
    ['ERROR.TYPE(#NAME?)', 5],
    ['ERROR.TYPE(#NUM!)', 6],
    ['ERROR.TYPE(NA())', 7],
    ['ERROR.TYPE(SEQUENCE(0))', 14],
    ['ERROR.TYPE(1)', '#N/A'],
    ['TYPE(2)', 1],
    ['TYPE("Smith")', 2],
    ['TYPE(TRUE)', 4],
    ['TYPE(E1)', 16],
    ['TYPE({1,2})', 64],
    ['TYPE(A1:A3)', 64],
    ['TYPE(A6)', 1],
    ['N(7)', 7],
    ['N("Even")', 0],
    ['N(TRUE)', 1],
    ['N(DATE(2011,4,17))', 40650],
    ['N(E1)', '#N/A'],
    ['N(A6)', 0],
  ], DATA);
});

test('SHEET, SHEETS and CELL', () => {
  const wb = workbook({
    Sheet1: {
      A1: 1, A4: 'x', B3: 5,
      C1: '=SHEET()', C2: '=SHEET("Sheet2")', C3: '=SHEET(Sheet2!A1)', C4: '=SHEET("nope")', C5: '=SHEETS()', C6: '=SHEETS(A1)',
      D1: '=CELL("address",B3)', D2: '=CELL("row",B3)', D3: '=CELL("col",B3)', D4: '=CELL("contents",B3)', D5: '=CELL("type",A4)',
      D6: '=CELL("type",A9)', D7: '=CELL("type",A1)', D8: '=CELL("width",A1)', D9: '=CELL("address",Sheet2!C5)',
    },
    Sheet2: { A1: 2 },
  });
  assert.equal(wb.get('C1'), 1);
  assert.equal(wb.get('C2'), 2);
  assert.equal(wb.get('C3'), 2);
  assert.equal(wb.get('C4'), '#N/A');
  assert.equal(wb.get('C5'), 2);
  assert.equal(wb.get('C6'), 1);
  assert.equal(wb.get('D1'), '$B$3');
  assert.equal(wb.get('D2'), 3);
  assert.equal(wb.get('D3'), 2);
  assert.equal(wb.get('D4'), 5);
  assert.equal(wb.get('D5'), 'l');
  assert.equal(wb.get('D6'), 'b');
  assert.equal(wb.get('D7'), 'v');
  assert.equal(wb.get('D8'), '#VALUE!');
  assert.equal(wb.get('D9'), 'Sheet2!$C$5');
});
