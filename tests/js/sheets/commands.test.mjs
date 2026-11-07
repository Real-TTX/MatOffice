// MatOffice Sheets – the commands of the ribbon and the keys on a workbook without a screen: fill series, AutoSum, sort, clear, borders,
// rows and columns, merges, number formats. The UI parts of the app are stubs.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../../../src/MatOffice/wwwroot/js/sheets/model.js';
import { Calc } from '../../../src/MatOffice/wwwroot/js/sheets/engine/calc.js';
import '../../../src/MatOffice/wwwroot/js/sheets/engine/functions/index.js';
import { Selection } from '../../../src/MatOffice/wwwroot/js/sheets/selection.js';
import { Look } from '../../../src/MatOffice/wwwroot/js/sheets/look.js';
import { Commands } from '../../../src/MatOffice/wwwroot/js/sheets/commands.js';
import { key } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';
import { FormulaError, setLocale } from '../../../src/MatOffice/wwwroot/js/sheets/engine/values.js';
import { splitText } from '../../../src/MatOffice/wwwroot/js/sheets/clipboard.js';

function makeApp(cells = {}) {
  setLocale('en');
  const sheetCells = {};
  for (const [a, v] of Object.entries(cells)) sheetCells[a] = typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v };
  const wb = Workbook.load({ version: 1, active: 0, styles: [{}], sheets: [{ id: 's1', name: 'Sheet1', cells: sheetCells }] });
  const app = {
    wb, locale: 'en', readOnly: false, showFormulas: false,
    get sheet() { return wb.activeSheet; },
    editor: { active: false },
    grid: { measure: (font, text) => text.length * 7, requestRender() {}, updateExtent() {} },
    changed() {
      const changes = wb.takeChanges();
      app.calc.update(changes);
      app.look.invalidate();
    },
    selectionChanged() {}, toastReadOnly() {}, toastLocked() {}, autoFitRows() {}, focusGrid() {},
    manualRows: new Set(),
    ctx: { toast() {}, ribbon: { refresh() {} } },
    root: { classList: { add() {}, remove() {} } },
  };
  const s = id => wb.sheets.find(x => x.id === id);
  app.calc = new Calc({
    cell: (id, r, c) => s(id)?.cells.get(key(r, c)),
    sheetId: name => wb.sheetByName(name)?.id ?? null,
    sheetName: id => s(id)?.name,
    sheetIds: () => wb.sheets.map(x => x.id),
    sheetIndex: id => wb.sheets.findIndex(x => x.id === id),
    bounds: id => s(id).bounds(),
    formulas: visit => wb.eachFormula((sh, r, c, cell) => visit(sh.id, r, c, cell.f)),
    names: () => wb.names,
    rowHidden: () => false,
    setResult: (id, r, c, v) => wb.setResult(s(id), r, c, v),
  });
  app.calc.rebuild();
  app.sel = new Selection();
  app.sel.mergeAt = (r, c) => app.sheet.mergeAt(r, c);
  app.look = new Look(app);
  app.commands = new Commands(app);
  app.value = (r, c) => { const v = app.calc.valueAt('s1', r, c); return v instanceof FormulaError ? v.code : v; };
  return app;
}

test('the fill handle continues numbers, dates, numbered texts, weekdays and formulas', () => {
  const app = makeApp({ A1: 1, A2: 3, B1: 'Item 1', C1: 'Monday', D1: '=A1*10', E1: 5 });
  const cmd = app.commands;
  cmd.fill({ r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 0, c1: 0, r2: 4, c2: 0 });
  assert.deepEqual([0, 1, 2, 3, 4].map(r => app.value(r, 0)), [1, 3, 5, 7, 9]);
  cmd.fill({ r1: 0, c1: 1, r2: 0, c2: 1 }, { r1: 0, c1: 1, r2: 2, c2: 1 });
  assert.deepEqual([0, 1, 2].map(r => app.value(r, 1)), ['Item 1', 'Item 2', 'Item 3']);
  cmd.fill({ r1: 0, c1: 2, r2: 0, c2: 2 }, { r1: 0, c1: 2, r2: 7, c2: 2 });
  assert.equal(app.value(1, 2), 'Tuesday');
  assert.equal(app.value(6, 2), 'Sunday');
  assert.equal(app.value(7, 2), 'Monday');
  cmd.fill({ r1: 0, c1: 3, r2: 0, c2: 3 }, { r1: 0, c1: 3, r2: 2, c2: 3 });
  assert.equal(app.sheet.cell(2, 3).f, 'A3*10');
  assert.equal(app.value(2, 3), 50);
  cmd.fill({ r1: 0, c1: 4, r2: 0, c2: 4 }, { r1: 0, c1: 4, r2: 0, c2: 6 });
  assert.deepEqual([4, 5, 6].map(c => app.value(0, c)), [5, 5, 5]); // one number is copied
  app.wb.undo();
  assert.equal(app.sheet.cell(0, 5), undefined);
});

test('dates fill by days, or by months when the days match', () => {
  const app = makeApp({ A1: 46053, A2: 46054, B1: 46023, B2: 46054 }); // 2026-01-31, 02-01; 01-01, 02-01
  app.wb.begin();
  app.wb.setCell(app.sheet, 0, 1, { v: 46023, s: app.wb.styleIndex({ nf: 'dd.mm.yyyy' }) });
  app.wb.setCell(app.sheet, 1, 1, { v: 46054, s: app.wb.styleIndex({ nf: 'dd.mm.yyyy' }) });
  app.wb.commit();
  app.changed();
  app.commands.fill({ r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 0, c1: 0, r2: 3, c2: 0 });
  assert.equal(app.value(3, 0), 46056);
  app.commands.fill({ r1: 0, c1: 1, r2: 1, c2: 1 }, { r1: 0, c1: 1, r2: 3, c2: 1 });
  assert.equal(app.value(2, 1), 46082); // 2026-03-01
  assert.equal(app.value(3, 1), 46113); // 2026-04-01
});

test('AutoSum below a column of numbers', () => {
  const app = makeApp({ A1: 2, A2: 3, A3: 4, B1: 10, B2: 20, B3: 30 });
  app.sel.setArea({ r1: 0, c1: 0, r2: 2, c2: 1 });
  app.commands.autoSum('SUM');
  assert.equal(app.sheet.cell(3, 0).f, 'SUM(A1:A3)');
  assert.equal(app.value(3, 1), 60);
});

test('sorting the data around the active cell keeps the header', () => {
  const app = makeApp({ A1: 'Name', B1: 'Score', A2: 'Cleo', B2: 7, A3: 'Ann', B3: 9, A4: 'Bob', B4: 3 });
  app.sel.set(1, 1);
  app.commands.sort(true);
  assert.deepEqual([0, 1, 2, 3].map(r => app.value(r, 0)), ['Name', 'Ann', 'Cleo', 'Bob']);
  app.sel.set(2, 0);
  app.commands.sort(false);
  assert.deepEqual([1, 2, 3].map(r => app.value(r, 0)), ['Ann', 'Bob', 'Cleo']);
});

test('clear, borders, number formats and merges', () => {
  const app = makeApp({ A1: 1, B1: 2, A2: 3, B2: 4 });
  const cmd = app.commands;
  app.sel.setArea({ r1: 0, c1: 0, r2: 1, c2: 1 });
  cmd.borders('outside');
  const st = (r, c) => app.wb.style(app.sheet.styleAt(r, c));
  assert.equal(st(0, 0).bt, 'thin #000000');
  assert.equal(st(0, 0).bl, 'thin #000000');
  assert.equal(st(0, 0).bb, undefined);
  assert.equal(st(1, 1).br, 'thin #000000');
  cmd.toggle('b');
  assert.equal(st(1, 0).b, true);
  cmd.numberFormat('0.00%');
  cmd.decimals(-1);
  assert.equal(st(0, 1).nf, '0.0%');
  cmd.clear('formats');
  assert.deepEqual(st(0, 0), {});
  assert.equal(app.value(0, 0), 1);
  cmd.clear('contents');
  assert.equal(app.value(0, 0), null);
});

test('inserting and deleting rows from the selection keeps formulas right', () => {
  const app = makeApp({ A1: 1, A2: 2, A3: '=A1+A2' });
  app.sel.rows(1, 1);
  app.commands.insertLines('row');
  assert.equal(app.sheet.cell(3, 0).f, 'A1+A3');
  assert.equal(app.value(3, 0), 3);
  app.sel.rows(0, 0);
  app.commands.deleteLines('row');
  assert.equal(app.sheet.cell(2, 0).f, '#REF!+A2');
  app.wb.undo();
  app.wb.undo();
  app.calc.rebuild();
  assert.equal(app.sheet.cell(2, 0).f, 'A1+A2');
});

test('Ctrl+arrow jumps to the edges of the data', () => {
  const app = makeApp({ A1: 1, A2: 2, A3: 3, A6: 6, A7: 7 });
  const cmd = app.commands;
  assert.deepEqual(cmd.edge(app.sheet, 0, 0, 1, 0), { row: 2, col: 0 });
  assert.deepEqual(cmd.edge(app.sheet, 2, 0, 1, 0), { row: 5, col: 0 });
  assert.deepEqual(cmd.edge(app.sheet, 5, 0, 1, 0), { row: 6, col: 0 });
  assert.deepEqual(cmd.edge(app.sheet, 6, 0, 1, 0), { row: 1048575, col: 0 });
  assert.deepEqual(cmd.edge(app.sheet, 6, 0, -1, 0), { row: 5, col: 0 });
  const region = cmd.currentRegion(app.sheet, 1, 0);
  assert.deepEqual(region, { r1: 0, c1: 0, r2: 2, c2: 0 });
});

test('conditional formats colour the cells that match', () => {
  const app = makeApp({ A1: 5, A2: 50, A3: 500, B1: 'late', B2: 'ok' });
  app.wb.begin();
  app.wb.setSheetProp(app.sheet, 'cf', [
    { range: 'A1:A3', type: 'cell', op: 'gt', v1: '10', style: { fill: '#ffc7ce' } },
    { range: 'A1:A3', type: 'top', n: 1, style: { b: true } },
    { range: 'B1:B2', type: 'text', op: 'contains', v1: 'LATE', style: { color: '#9c0006' } },
    { range: 'A1:A3', type: 'bar', color: '#638ec6' },
  ]);
  app.wb.commit();
  app.look.invalidate();
  const look = (r, c) => app.look.cell(app.sheet, r, c);
  assert.equal(look(0, 0).fill, null);
  assert.equal(look(1, 0).fill, '#ffc7ce');
  assert.equal(look(2, 0).b, true);
  assert.equal(look(1, 0).b, false);
  assert.equal(look(0, 1).color, '#9c0006');
  assert.ok(Math.abs(look(2, 0).bar.ratio - 1) < 1e-9);
  assert.equal(look(0, 0).text, '5');
  assert.equal(look(0, 0).ha, 'right');
  assert.equal(look(0, 1).ha, 'left');
});

test('pasted text with tabs, lines and quotes', () => {
  assert.deepEqual(splitText('a\tb\r\n1\t2\r\n'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(splitText('"x\ny"\t"say ""hi"""'), [['x\ny', 'say "hi"']]);
});
