// MatOffice Sheets – charts: the data read like Excel (series names, categories) and every chart type drawn without an error
// (a tiny stand-in for the SVG elements of the browser).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../../../src/MatOffice/wwwroot/js/sheets/model.js';
import { Calc } from '../../../src/MatOffice/wwwroot/js/sheets/engine/calc.js';
import '../../../src/MatOffice/wwwroot/js/sheets/engine/functions/index.js';
import { Look } from '../../../src/MatOffice/wwwroot/js/sheets/look.js';
import { Commands } from '../../../src/MatOffice/wwwroot/js/sheets/commands.js';
import { chartData, renderChart, niceScale, CHART_TYPES } from '../../../src/MatOffice/wwwroot/js/sheets/charts.js';
import { key } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';

class FakeNode {
  constructor(tag) { this.tag = tag; this.attrs = {}; this.children = []; this.textContent = ''; }
  setAttribute(k, v) { this.attrs[k] = v; }
  append(...nodes) { this.children.push(...nodes); }
  insertBefore(node) { this.children.unshift(node); }
  count(tag) { return (this.tag === tag ? 1 : 0) + this.children.reduce((n, c) => n + c.count(tag), 0); }
}
globalThis.document = { createElementNS: (ns, tag) => new FakeNode(tag) };

function makeApp(cells) {
  const sheetCells = {};
  for (const [a, v] of Object.entries(cells)) sheetCells[a] = typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v };
  const wb = Workbook.load({ version: 1, active: 0, styles: [{}], sheets: [{ id: 's1', name: 'Sheet1', cells: sheetCells }] });
  const app = { wb, locale: 'en', showFormulas: false, get sheet() { return wb.activeSheet; } };
  const s = () => wb.sheets[0];
  app.calc = new Calc({
    cell: (id, r, c) => s().cells.get(key(r, c)), sheetId: () => 's1', sheetName: () => 'Sheet1', sheetIds: () => ['s1'], sheetIndex: () => 0,
    bounds: () => s().bounds(), formulas: visit => wb.eachFormula((sh, r, c, cell) => visit(sh.id, r, c, cell.f)), names: () => [], rowHidden: () => false,
    setResult: (id, r, c, v) => wb.setResult(s(), r, c, v),
  });
  app.calc.rebuild();
  app.look = new Look(app);
  app.commands = new Commands(app);
  return app;
}

const sales = { A1: 'Region', B1: 'Q1', C1: 'Q2', A2: 'North', B2: 10, C2: 12, A3: 'South', B3: 7, C3: '=B3*2', A4: 'West', B4: -3, C4: 5 };

test('the data of a chart: series from the header row, categories from the first column', () => {
  const app = makeApp(sales);
  const d = chartData(app, app.sheet, { type: 'column', range: 'A1:C4' });
  assert.deepEqual(d.categories, ['North', 'South', 'West']);
  assert.deepEqual(d.series.map(s => s.name), ['Q1', 'Q2']);
  assert.deepEqual(d.series[1].values, [12, 14, 5]);
  const rows = chartData(app, app.sheet, { type: 'column', range: 'A1:C4', byRows: true });
  assert.deepEqual(rows.series.map(s => s.name), ['North', 'South', 'West']);
  assert.deepEqual(rows.categories, ['Q1', 'Q2']);
});

test('every type of chart is drawn', () => {
  const app = makeApp(sales);
  for (const type of CHART_TYPES) {
    for (const extra of [{}, { stacked: true, labels: true, title: 'Sales', xTitle: 'Region', yTitle: 'EUR', legend: 'bottom' }]) {
      const chart = { type, range: 'A1:C4', w: 480, h: 300, ...extra };
      const svg = renderChart(chart, chartData(app, app.sheet, chart), { width: 480, height: 300 });
      assert.equal(svg.tag, 'svg', type);
      const marks = svg.count('rect') + svg.count('path') + svg.count('circle');
      assert.ok(marks > 3, type + ' draws its marks');
    }
  }
  const empty = renderChart({ type: 'pie', range: 'Z1:Z2' }, { categories: [], series: [] }, { width: 200, height: 100 });
  assert.equal(empty.tag, 'svg');
});

test('nice scales of the value axis', () => {
  assert.deepEqual(niceScale(0, 14), { min: 0, max: 15, step: 5 });
  assert.deepEqual(niceScale(-3, 14), { min: -5, max: 15, step: 5 });
  assert.deepEqual(niceScale(0, 0), { min: 0, max: 1, step: 0.2 });
});
