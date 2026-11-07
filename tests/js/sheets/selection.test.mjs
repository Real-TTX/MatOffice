// MatOffice Sheets – the selection: merged cells are selected whole, Enter/Tab walk through a range, whole rows and columns.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Selection } from '../../../src/MatOffice/wwwroot/js/sheets/selection.js';

const merges = [{ r1: 1, c1: 1, r2: 2, c2: 3 }];
const mergeAt = (r, c) => merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;

test('a merged cell is selected whole and ranges grow around merges', () => {
  const sel = new Selection();
  sel.mergeAt = mergeAt;
  sel.set(2, 2);
  assert.deepEqual(sel.active, { row: 1, col: 1 });
  assert.deepEqual(sel.range, { r1: 1, c1: 1, r2: 2, c2: 3 });
  assert.ok(sel.isSingle);
  sel.set(0, 0);
  sel.extendTo(1, 1);
  assert.deepEqual(sel.range, { r1: 0, c1: 0, r2: 2, c2: 3 });
});

test('Enter and Tab walk through a range and wrap', () => {
  const sel = new Selection();
  sel.setArea({ r1: 0, c1: 0, r2: 1, c2: 1 });
  const seen = [];
  for (let i = 0; i < 5; i++) { sel.cycle(1, 0); seen.push([sel.active.row, sel.active.col]); }
  assert.deepEqual(seen, [[1, 0], [0, 1], [1, 1], [0, 0], [1, 0]]);
  sel.setArea({ r1: 0, c1: 0, r2: 1, c2: 1 });
  sel.cycle(0, 1);
  assert.deepEqual(sel.active, { row: 0, col: 1 });
  sel.cycle(0, 1);
  assert.deepEqual(sel.active, { row: 1, col: 0 });
  sel.set(5, 5);
  assert.equal(sel.cycle(1, 0), false);
});

test('rows, columns, everything and Ctrl+click ranges', () => {
  const sel = new Selection();
  sel.cols(2, 4);
  assert.ok(sel.isWholeCol());
  assert.equal(sel.range.r2, 1048575);
  sel.rows(3, 3, true);
  assert.equal(sel.ranges.length, 2);
  assert.ok(sel.isWholeRow());
  assert.ok(sel.contains(3, 100));
  assert.ok(sel.contains(500, 3));
  sel.all();
  assert.ok(sel.contains(1048575, 16383));
  const saved = sel.toJSON();
  sel.set(1, 1);
  sel.restore(saved);
  assert.equal(sel.range.c2, 16383);
});
