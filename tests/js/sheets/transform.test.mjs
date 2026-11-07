// MatOffice Sheets – formulas after inserted/deleted rows and columns, copies, moves and renamed sheets.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shiftFormula, moveFormula, renameSheetIn, deleteSheetIn, moveAreaIn, referencesOf } from '../../../src/MatOffice/wwwroot/js/sheets/engine/transform.js';

const rows = (f, at, count, sheetName = 'Sheet1', hostName = 'Sheet1') => shiftFormula(f, { hostName, sheetName, axis: 'row', at, count });
const cols = (f, at, count) => shiftFormula(f, { hostName: 'Sheet1', sheetName: 'Sheet1', axis: 'col', at, count });

test('inserting rows moves and grows references', () => {
  assert.equal(rows('SUM(A1:A10)', 0, 1), 'SUM(A2:A11)');
  assert.equal(rows('SUM(A1:A10)', 4, 2), 'SUM(A1:A12)');
  assert.equal(rows('SUM(A1:A10)', 10, 1), 'SUM(A1:A10)');
  assert.equal(rows('$B$5 + B5', 2, 1), '$B$6 + B6');
  assert.equal(rows('SUM(C:C)', 0, 5), 'SUM(C:C)');
  assert.equal(rows('SUM(2:3)', 0, 1), 'SUM(3:4)');
  assert.equal(rows('Other!A5', 0, 1), 'Other!A5');
  assert.equal(rows('Other!A5', 0, 1, 'Other'), 'Other!A6');
  assert.equal(rows('A5', 0, 1, 'Other', 'Other'), 'A6');
});

test('deleting rows shrinks references and breaks deleted ones', () => {
  assert.equal(rows('SUM(A1:A10)', 4, -2), 'SUM(A1:A8)');
  assert.equal(rows('A5*2', 4, -1), '#REF!*2');
  assert.equal(rows('A6*2', 4, -1), 'A5*2');
  assert.equal(rows('SUM(A3:A5)', 2, -3), 'SUM(#REF!)');
  assert.equal(rows('SUM(A3:A8)', 0, -4), 'SUM(A1:A4)');
  assert.equal(rows("'My sheet'!B9", 0, -2, 'My sheet'), "'My sheet'!B7");
});

test('columns', () => {
  assert.equal(cols('B1+C1', 1, 1), 'C1+D1');
  assert.equal(cols('SUM(A1:D1)', 2, -1), 'SUM(A1:C1)');
  assert.equal(cols('SUM(B:D)', 0, 1), 'SUM(C:E)');
  assert.equal(cols('SUM(1:1)', 0, 1), 'SUM(1:1)');
});

test('copied formulas move their relative references', () => {
  assert.equal(moveFormula('A1+$B$1+C$1+$D1', 2, 1), 'B3+$B$1+D$1+$D3');
  assert.equal(moveFormula('SUM(A1:A3)', 1, 0), 'SUM(A2:A4)');
  assert.equal(moveFormula('A1', -1, 0), '#REF!');
  assert.equal(moveFormula('SUM(A:A)', 5, 1), 'SUM(B:B)');
  assert.equal(moveFormula('"A1"&A1', 0, 1), '"A1"&B1');
});

test('sheets renamed and deleted', () => {
  assert.equal(renameSheetIn('Sheet1!A1+sheet1!B2+A3', 'Sheet1', 'Data 2026'), "'Data 2026'!A1+'Data 2026'!B2+A3");
  assert.equal(deleteSheetIn("SUM('Old one'!A1:B2)+1", 'Old one'), "SUM('Old one'!#REF!)+1");
});

test('references follow cells that were cut and pasted', () => {
  const move = (f, opts) => moveAreaIn(f, { hostName: 'Sheet1', fromName: 'Sheet1', toName: 'Sheet1', ...opts });
  const from = { r1: 0, c1: 0, r2: 9, c2: 0 };
  assert.equal(move('A1+A5+B1', { from, dRow: 0, dCol: 2 }), 'C1+C5+B1');
  assert.equal(move('SUM(A1:A10)', { from, dRow: 1, dCol: 0 }), 'SUM(A2:A11)');
  assert.equal(move('SUM(A1:A20)', { from, dRow: 1, dCol: 0 }), 'SUM(A1:A20)');
  assert.equal(move('A1', { from, dRow: 0, dCol: 0, toName: 'Other' }), 'Other!A1');
});

test('references with their places for the colours of the editor', () => {
  const refs = referencesOf('SUM(A1:B2)+Sheet2!C3');
  assert.deepEqual(refs.map(r => [r.start, r.end, r.sheet]), [[4, 9, null], [11, 20, 'Sheet2']]);
});
