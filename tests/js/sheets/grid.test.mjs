// MatOffice Sheets – the grid drawn into a context that only records texts: every pane (frozen rows and columns), merges, borders,
// conditional formats, notes, filter buttons, overflowing text, the selection; hit testing maps points back to cells.
import { fakeContext } from './fake-browser.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../../../src/MatOffice/wwwroot/js/sheets/model.js';
import { Calc } from '../../../src/MatOffice/wwwroot/js/sheets/engine/calc.js';
import '../../../src/MatOffice/wwwroot/js/sheets/engine/functions/index.js';
import { Selection } from '../../../src/MatOffice/wwwroot/js/sheets/selection.js';
import { Look } from '../../../src/MatOffice/wwwroot/js/sheets/look.js';
import { Commands } from '../../../src/MatOffice/wwwroot/js/sheets/commands.js';
import { Grid } from '../../../src/MatOffice/wwwroot/js/sheets/grid.js';
import { key } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';

function makeApp() {
  const content = {
    version: 1, active: 0,
    styles: [{}, { b: true, fill: '#107c41', color: '#ffffff', bb: 'medium #000000' }, { nf: '#,##0.00', bt: 'double #ff0000' }, { wrap: true, ha: 'center', va: 'middle' }],
    sheets: [{
      id: 's1', name: 'Sheet1',
      cells: {
        A1: { v: 'A long title that runs into the empty cells', s: 1 }, A3: { v: 'Item', s: 1 }, B3: { v: 'Amount', s: 1 },
        A4: { v: 'Tea' }, B4: { v: 1234.5, s: 2, note: 'Bought in May' }, A5: { v: 'Coffee' }, B5: { v: -7, s: 2 }, B6: { f: 'SUM(B4:B5)', s: 2 },
        C4: { v: 'wrapped text in a cell', s: 3 }, D2: { f: 'SEQUENCE(3)' }, E4: { v: true }, F4: { f: '1/0' },
      },
      merges: ['A1:C1'], freeze: { rows: 3, cols: 1 }, filter: { range: 'A3:B6' },
      cf: [{ range: 'B4:B6', type: 'cell', op: 'lt', v1: '0', style: { color: '#9c0006', fill: '#ffc7ce' } }, { range: 'B4:B6', type: 'bar', color: '#638ec6' }],
      cols: { C: { w: 90 } }, rows: { 4: { h: 34 } },
    }],
  };
  const wb = Workbook.load(content);
  const app = { wb, locale: 'en', readOnly: false, showFormulas: false, get sheet() { return wb.activeSheet; } };
  const s = () => wb.sheets[0];
  app.calc = new Calc({
    cell: (id, r, c) => s().cells.get(key(r, c)), sheetId: () => 's1', sheetName: () => 'Sheet1', sheetIds: () => ['s1'], sheetIndex: () => 0,
    bounds: () => s().bounds(), formulas: visit => wb.eachFormula((sh, r, c, cell) => visit(sh.id, r, c, cell.f)), names: () => [], rowHidden: () => false,
    setResult: (id, r, c, v) => wb.setResult(s(), r, c, v),
  });
  app.calc.rebuild();
  app.sel = new Selection();
  app.sel.mergeAt = (r, c) => app.sheet.mergeAt(r, c);
  app.sel.sheetId = 's1';
  app.look = new Look(app);
  app.commands = new Commands(app);
  app.grid = new Grid(app);
  Object.assign(app.grid, { width: 900, height: 500, dpr: 1 });
  return app;
}

test('the grid draws every pane, the headers and the formatted values', () => {
  const app = makeApp();
  const grid = app.grid;
  app.sel.setArea({ r1: 3, c1: 1, r2: 5, c2: 1 });
  grid.highlights = [{ sheet: 's1', area: { r1: 3, c1: 0, r2: 4, c2: 0 }, color: '#4472c4' }];
  grid.setMarquee({ sheet: 's1', area: { r1: 3, c1: 1, r2: 3, c2: 1 } });
  grid.updateExtent();
  grid.render();
  grid.setMarquee(null);
  const texts = grid.g.texts;
  for (const t of ['A long title that runs into the empty cells', 'Item', 'Tea', '1,234.50', '-7.00', '1,227.50', 'TRUE', '#DIV/0!', 'A', 'B', '1', '6']) assert.ok(texts.includes(t), 'drawn: ' + t);
  assert.ok(texts.includes('2') && texts.includes('3'), 'the spilled SEQUENCE');
});

test('points map to cells, headers, resize handles and the fill handle', () => {
  const app = makeApp();
  const grid = app.grid;
  const m = grid.metrics();
  assert.equal(m.fr, 3);
  assert.equal(m.fc, 1);
  const r = grid.cellRect(4, 2, m);
  const hit = grid.hitTest(r.x + 5, r.y + 5, m);
  assert.deepEqual([hit.zone, hit.row, hit.col], ['cell', 4, 2]);
  const frozen = grid.cellRect(0, 0, m);
  const merged = grid.hitTest(frozen.x + 3, frozen.y + 3, m);
  assert.deepEqual([merged.row, merged.col], [0, 0]);
  assert.ok(frozen.w > 150, 'A1:C1 is one cell');
  assert.equal(grid.hitTest(r.x + 5, 3, m).zone, 'col');
  assert.equal(grid.hitTest(3, r.y + 5, m).zone, 'row');
  assert.equal(grid.hitTest(2, 2, m).zone, 'corner');
  const colEdge = grid.hitTest(r.x + r.w - 1, 5, m);
  assert.deepEqual(colEdge.resize, { axis: 'col', index: 2 });
  app.sel.set(4, 2);
  const fill = grid.hitTest(r.x + r.w - 1, r.y + r.h - 1, m);
  assert.ok(fill.fill);
  const header = grid.cellRect(2, 1, m);
  const button = grid.filterButtonRect(header, 1);
  assert.equal(grid.hitTest(button.x + 2, button.y + 2, m).filter, 1);
});

test('scrolling, zoom and the picture of the gallery', () => {
  const app = makeApp();
  const grid = app.grid;
  grid.scrollIntoView(200, 30);
  assert.ok(grid.scroller.scrollTop > 0);
  assert.ok(grid.scroller.scrollLeft > 0);
  const m = grid.metrics();
  const r = grid.cellRect(200, 30, m);
  assert.ok(r.y >= m.mainY && r.y + r.h <= grid.height + 1, 'the row is in the window');
  app.sheet.zoom = 1.5;
  grid.render();
  app.sheet.headings = false;
  app.sheet.grid = false;
  grid.render();
  const ctx = fakeContext();
  grid.renderTo(ctx, 480, 360, app.sheet, 1);
  assert.ok(ctx.texts.includes('Item'));
  assert.equal(grid.width, 900, 'the screen size comes back');
  assert.equal(typeof grid.theme, 'function');
  grid.render();
});
