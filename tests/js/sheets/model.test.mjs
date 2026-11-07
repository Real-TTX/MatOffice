// MatOffice Sheets – the workbook model: content in and out, undo/redo, inserted/deleted rows, moved and sorted cells, sizes.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { Workbook } from '../../../src/MatOffice/wwwroot/js/sheets/model.js';
import { Axis } from '../../../src/MatOffice/wwwroot/js/sheets/axis.js';
import { key } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';
import { sortCompare } from '../../../src/MatOffice/wwwroot/js/sheets/engine/values.js';

const templates = new URL('../../../src/MatOffice/Resources/Templates/sheets/', import.meta.url);
const localize = (node, lang) => {
  if (Array.isArray(node)) return node.map(n => localize(n, lang));
  if (node && typeof node === 'object') {
    const keys = Object.keys(node);
    if (keys.length === 2 && keys.includes('en') && keys.includes('de')) return node[lang];
    return Object.fromEntries(Object.entries(node).map(([k, v]) => [k, localize(v, lang)]));
  }
  return node;
};

const sample = () => ({
  version: 1, active: 0, styles: [{}, { b: true }, { nf: '0.00' }],
  sheets: [
    { id: 's1', name: 'Sheet1', cells: { A1: { v: 'Name', s: 1 }, A2: { v: 3 }, A3: { v: 1 }, A4: { v: 2 }, B2: { f: 'A2*2', v: 6 }, B3: { f: 'A3*2', v: 2 }, B4: { f: 'A4*2', v: 4 }, C1: { f: 'SUM(A2:A4)', v: 6, s: 2 } }, cols: { A: { w: 100 } }, rows: { 3: { h: 30 } }, merges: ['D1:E1'] },
    { id: 's2', name: 'Other', cells: { A1: { f: 'Sheet1!A2+1', v: 4 } } },
  ],
});

test('every template loads and is written back the same', () => {
  for (const file of readdirSync(templates)) {
    for (const lang of ['en', 'de']) {
      const content = localize(JSON.parse(readFileSync(new URL(file, templates), 'utf8')), lang);
      const wb = Workbook.load(content);
      const again = Workbook.load(wb.toContent()).toContent();
      assert.deepEqual(again, wb.toContent(), file);
      assert.ok(wb.sheets[0].cells.size > 5, file);
    }
  }
});

test('content round trip keeps cells, sizes, merges and styles', () => {
  const wb = Workbook.load(sample());
  const out = wb.toContent();
  assert.deepEqual(out.sheets[0].cells.A1, { v: 'Name', s: 1 });
  assert.deepEqual(out.sheets[0].cols, { A: { w: 100 } });
  assert.deepEqual(out.sheets[0].rows, { 3: { h: 30 } });
  assert.deepEqual(out.sheets[0].merges, ['D1:E1']);
  assert.deepEqual(out.styles, [{}, { b: true }, { nf: '0.00' }]);
});

test('unused styles are left out and indexes renumbered', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  wb.begin();
  wb.patchCell(s, 0, 0, { s: undefined });
  wb.commit();
  const out = wb.toContent();
  assert.deepEqual(out.styles, [{}, { nf: '0.00' }]);
  assert.equal(out.sheets[0].cells.C1.s, 1);
});

test('undo and redo restore cells and properties', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  wb.begin('type');
  wb.setCell(s, 9, 9, { v: 'x' });
  wb.setSheetProp(s, 'freeze', { rows: 1, cols: 0 });
  wb.commit();
  assert.equal(s.cell(9, 9).v, 'x');
  wb.undo();
  assert.equal(s.cell(9, 9), undefined);
  assert.deepEqual(s.freeze, { rows: 0, cols: 0 });
  wb.redo();
  assert.equal(s.cell(9, 9).v, 'x');
  assert.deepEqual(s.freeze, { rows: 1, cols: 0 });
});

test('inserting and deleting rows adjusts formulas on every sheet', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  wb.shiftLines(s, 'row', 2, 1); // a new row 3
  assert.equal(s.cell(4, 1).f, 'A5*2');
  assert.equal(s.cell(0, 2).f, 'SUM(A2:A5)');
  assert.equal(wb.sheets[1].cell(0, 0).f, 'Sheet1!A2+1');
  assert.equal(s.rows.rawSize(3), 30); // the old row 3 is row 4 now
  wb.shiftLines(s, 'row', 1, -1); // delete row 2
  assert.equal(wb.sheets[1].cell(0, 0).f, 'Sheet1!#REF!+1');
  assert.equal(s.cell(0, 2).f, 'SUM(A2:A4)');
  wb.undo();
  wb.undo();
  assert.deepEqual(wb.toContent(), Workbook.load(sample()).toContent());
});

test('inserting columns moves merges and widths', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  wb.shiftLines(s, 'col', 0, 2);
  assert.equal(s.cols.rawSize(2), 100);
  assert.deepEqual(s.merges, [{ r1: 0, c1: 5, r2: 0, c2: 6 }]);
  assert.equal(s.cell(1, 3).f, 'C2*2');
});

test('renaming and deleting sheets rewrites references', () => {
  const wb = Workbook.load(sample());
  assert.ok(wb.renameSheet(wb.sheets[0], 'Data 2026'));
  assert.equal(wb.sheets[1].cell(0, 0).f, "'Data 2026'!A2+1");
  assert.ok(!wb.renameSheet(wb.sheets[1], 'data 2026'));
  assert.ok(wb.deleteSheet(wb.sheets[0]));
  assert.equal(wb.sheets[0].cell(0, 0).f, "'Data 2026'!#REF!+1");
  wb.undo();
  assert.equal(wb.sheets.length, 2);
  assert.equal(wb.sheets[1].cell(0, 0).f, "'Data 2026'!A2+1");
});

test('moved cells take their references along', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  wb.moveCells(s, { r1: 1, c1: 0, r2: 3, c2: 0 }, s, 0, 5); // A2:A4 → F2:F4
  assert.equal(s.cell(1, 5).v, 3);
  assert.equal(s.cell(1, 0), undefined);
  assert.equal(s.cell(1, 1).f, 'F2*2');
  assert.equal(s.cell(0, 2).f, 'SUM(F2:F4)');
  assert.equal(wb.sheets[1].cell(0, 0).f, 'Sheet1!F2+1');
});

test('sorting rows moves formulas with their row', () => {
  const wb = Workbook.load(sample());
  const s = wb.sheets[0];
  const valueAt = (r, c) => s.cell(r, c)?.v ?? null;
  wb.sortRows(s, { r1: 1, c1: 0, r2: 3, c2: 1 }, [{ col: 0, desc: false }], valueAt, sortCompare);
  assert.deepEqual([1, 2, 3].map(r => s.cell(r, 0).v), [1, 2, 3]);
  assert.deepEqual([1, 2, 3].map(r => s.cell(r, 1).f), ['A2*2', 'A3*2', 'A4*2']);
});

test('the axis finds positions with custom and hidden sizes', () => {
  const ax = new Axis(20, 1000);
  ax.setSize(2, 50);
  ax.setHidden(4, true);
  assert.equal(ax.offset(0), 0);
  assert.equal(ax.offset(3), 90);
  assert.equal(ax.offset(5), 110);
  assert.equal(ax.indexAt(0), 0);
  assert.equal(ax.indexAt(45), 2);
  assert.equal(ax.indexAt(89), 2);
  assert.equal(ax.indexAt(90), 3);
  assert.equal(ax.indexAt(110), 5);
  assert.equal(ax.indexAt(1e9), 999);
  assert.equal(ax.step(3, 1), 5);
  ax.insert(0, 2);
  assert.equal(ax.rawSize(4), 50);
  assert.ok(ax.isHidden(6));
  ax.remove(0, 3);
  assert.equal(ax.rawSize(1), 50);
  assert.equal(key(1, 1), 16385);
});
