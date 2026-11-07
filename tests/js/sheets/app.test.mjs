// MatOffice Sheets – the whole editor mounted in node with a stand-in for the frame: typing, formulas, the keys, undo, the clipboard,
// the content saved and loaded, the picture, and every dialog built (a dialog that answers "cancel" or "ok").
import './fake-browser.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { SheetsApp } from '../../../src/MatOffice/wwwroot/js/sheets/app.js';

const template = name => {
  const localize = (node, lang) => {
    if (Array.isArray(node)) return node.map(n => localize(n, lang));
    if (node && typeof node === 'object') {
      const keys = Object.keys(node);
      if (keys.length === 2 && keys.includes('en') && keys.includes('de')) return node[lang];
      return Object.fromEntries(Object.entries(node).map(([k, v]) => [k, localize(v, lang)]));
    }
    return node;
  };
  return localize(JSON.parse(readFileSync(new URL(`../../../src/MatOffice/Resources/Templates/sheets/${name}.json`, import.meta.url), 'utf8')), 'en');
};

function mount(content, { readOnly = false, answer = 'cancel' } = {}) {
  const app = new SheetsApp();
  const log = { changed: 0, toasts: [], dialogs: [] };
  const ctx = {
    doc: { content, canEdit: !readOnly, mode: 'document', user: { name: 'Admin' } },
    readOnly, work: document.createElement('main'),
    ribbon: { set(tabs) { this.tabs = tabs; }, refresh() {}, select() {}, menuItems: () => document.createElement('div') },
    status: { left(...parts) { this.parts = parts; }, zoom() {}, setZoom() {} },
    changed: () => { log.changed++; },
    toast: (text, tone) => log.toasts.push([text, tone]),
    dialog: async options => { log.dialogs.push(options.title); return typeof answer === 'function' ? answer(options) : answer; },
  };
  app.mount(ctx);
  Object.assign(app.grid, { width: 900, height: 500, dpr: 1 });
  app.grid.updateExtent();
  app.grid.render();
  return { app, ctx, log };
}

const blank = () => ({ version: 1, active: 0, styles: [{}], sheets: [{ id: 's1', name: 'Sheet1', cells: {} }] });
const keydown = (target, key, extra = {}) => target.dispatch('keydown', { key, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, ...extra });
const typeCell = (app, text) => {
  app.keys.value = text[0];
  app.keys.dispatch('input');
  const input = app.editor.input;
  input.setRangeText(text.slice(1), input.value.length, input.value.length, 'end');
  app.editor.onInput(input);
};

test('typing values and a formula, Enter and Tab, the keys move', () => {
  const { app, log } = mount(blank());
  typeCell(app, 'Item');
  keydown(app.editor.input, 'Tab');
  typeCell(app, '12.5');
  keydown(app.editor.input, 'Enter');
  assert.deepEqual(app.sel.active, { row: 1, col: 0 }, 'Tab … Enter goes back to the first column');
  typeCell(app, 'x');
  keydown(app.editor.input, 'Tab');
  typeCell(app, '=B1*2');
  keydown(app.editor.input, 'Enter');
  assert.equal(app.sheet.cell(0, 0).v, 'Item');
  assert.equal(app.sheet.cell(0, 1).v, 12.5);
  assert.equal(app.sheet.cell(1, 1).f, 'B1*2');
  assert.equal(app.sheet.cell(1, 1).v, 25);
  assert.ok(log.changed >= 4);
  keydown(app.keys, 'ArrowUp');
  keydown(app.keys, 'ArrowRight');
  assert.deepEqual(app.sel.active, { row: 1, col: 1 });
  keydown(app.keys, 'ArrowUp', { shiftKey: true });
  assert.deepEqual(app.sel.range, { r1: 0, c1: 1, r2: 1, c2: 1 });
  keydown(app.keys, 'Delete');
  assert.equal(app.sheet.cell(0, 1)?.v, undefined);
  assert.equal(app.calc.valueAt('s1', 1, 1), null);
  keydown(app.keys, 'z', { ctrlKey: true });
  assert.equal(app.sheet.cell(0, 1).v, 12.5);
  assert.equal(app.sheet.cell(1, 1).v, 25);
  assert.equal(app.calc.valueAt('s1', 1, 1), 25);
  keydown(app.keys, 'y', { ctrlKey: true });
  assert.equal(app.sheet.cell(0, 1)?.v, undefined);
  keydown(app.keys, 'z', { ctrlKey: true });
  keydown(app.keys, 'Home', { ctrlKey: true });
  assert.deepEqual(app.sel.active, { row: 0, col: 0 });
  const content = app.getContent();
  assert.equal(content.sheets[0].cells.A1.v, 'Item');
  assert.equal(content.sheets[0].cells.B2.f, 'B1*2');
  app.setContent(content);
  assert.equal(app.sheet.cell(0, 0).v, 'Item');
});

test('a formula that is not complete is closed; a wrong one keeps the editor open', () => {
  const { app, log } = mount(blank());
  typeCell(app, '=SUM(1,2');
  keydown(app.editor.input, 'Enter');
  assert.equal(app.sheet.cell(0, 0).f, 'SUM(1,2)');
  assert.equal(app.sheet.cell(0, 0).v, 3);
  typeCell(app, '=1+*');
  keydown(app.editor.input, 'Enter');
  assert.ok(app.editor.active);
  assert.ok(log.dialogs.length >= 1);
  keydown(app.editor.input, 'Escape');
  assert.ok(!app.editor.active);
});

test('copy and paste through the events of the browser', () => {
  const { app } = mount(blank());
  typeCell(app, '7');
  keydown(app.editor.input, 'Enter');
  app.sel.set(0, 0);
  const data = {};
  app.keys.dispatch('copy', { clipboardData: { setData: (type, value) => { data[type] = value; } } });
  assert.equal(data['text/plain'], '7');
  assert.match(data['text/html'], /<table/);
  app.sel.set(3, 3);
  app.keys.dispatch('paste', { clipboardData: { files: [], getData: type => (type === 'text/plain' ? '7' : '') } });
  assert.equal(app.sheet.cell(3, 3).v, 7);
  app.sel.set(5, 0);
  app.keys.dispatch('paste', { clipboardData: { files: [], getData: type => (type === 'text/plain' ? 'a\tb\r\n1\t2\r\n' : '') } });
  assert.equal(app.sheet.cell(6, 1).v, 2);
  app.clipboard.clearMarquee(); // the marching ants stop (their timer would keep node running)
});

test('a template opens, its formulas are calculated and the picture is drawn', async () => {
  for (const name of ['budget', 'invoice', 'tasks', 'schedule', 'loan', 'sales']) {
    const { app } = mount(template(name));
    let errors = 0;
    app.wb.eachFormula((s, r, c, cell) => { if (cell.t === 'e') errors++; });
    assert.equal(errors, 0, name + ' calculates without errors');
    await app.thumbnail();
    assert.ok(app.info());
  }
});

test('every dialog can be opened (and cancelled), and the main ones applied', async () => {
  const { app, log } = mount(template('sales'));
  const d = app.dialogs;
  app.sel.setArea({ r1: 3, c1: 1, r2: 8, c2: 2 });
  await d.formatCells();
  await d.size('col');
  await d.size('row');
  await d.shiftCells('insert');
  await d.goTo();
  d.find(true);
  await d.sort();
  await d.removeDuplicates();
  await d.validation();
  await d.cfRule('gt');
  await d.cfRule('between');
  await d.cfRule('top');
  await d.cfFormula();
  await d.manageRules();
  await d.defineName();
  await d.nameManager();
  await d.insertFunction();
  await d.link();
  await d.chart(app.sheet.charts[0]);
  await d.printTitles();
  await d.zoom();
  app.sel.setArea({ r1: 3, c1: 0, r2: 8, c2: 0 });
  await d.textToColumns();
  assert.ok(log.dialogs.length >= 20);
  // with OK: the format cells dialog applies what it shows, a quick rule is added
  const ok = mount(template('sales'), { answer: 'ok' });
  ok.app.sel.setArea({ r1: 3, c1: 1, r2: 8, c2: 1 });
  await ok.app.dialogs.formatCells();
  await ok.app.dialogs.cfRule('gt');
  assert.equal(ok.app.sheet.cf.length > 0, true);
  await ok.app.dialogs.sort();
});

/** Every item of the ribbon: its state callbacks, then its command (menus and split buttons included). Returns the errors. */
function runRibbon(app, { run = true } = {}) {
  const errors = [];
  const items = [];
  const collect = list => {
    for (const item of (typeof list === 'function' ? list() : list) || []) {
      if (!item || item === '-' || item.heading) continue;
      items.push(item);
      if (item.type === 'stack') item.rows.forEach(collect);
      if (item.items) collect(item.items);
    }
  };
  for (const tab of app.ctx.ribbon.tabs) for (const group of tab.groups) collect(group.items);
  for (const item of items) {
    for (const f of ['value', 'active', 'disabled', 'hidden', 'checked', 'options']) {
      try { if (typeof item[f] === 'function') item[f](); } catch (e) { errors.push(`${item.label} ${f}: ${e.message}`); }
    }
    if (item.render) { try { item.render(); } catch (e) { errors.push(`${item.label} render: ${e.message}`); } }
  }
  if (run) {
    for (const item of items) {
      if (typeof item.run !== 'function') continue;
      if (/Print/.test(item.label || '')) continue;
      try {
        const r = item.run();
        if (r && typeof r.catch === 'function') r.catch(e => errors.push(`${item.label} run (async): ${e.message}`));
      } catch (e) { errors.push(`${item.label} run: ${e.message}\n${(e.stack || '').split('\n').slice(1, 3).join('\n')}`); }
      if (app.editor.active) app.editor.cancel();
    }
    for (const item of items) if (item.change) {
      try { item.change(item.type === 'select' && item.options ? (typeof item.options === 'function' ? item.options() : item.options)[1]?.value : item.type === 'color' ? '#ff0000' : '12'); } catch (e) { errors.push(`${item.label} change: ${e.message}`); }
    }
  }
  return { errors, count: items.length };
}

test('every command of the ribbon runs on a real workbook', async () => {
  const { app } = mount(template('sales'));
  app.sel.setArea({ r1: 3, c1: 1, r2: 6, c2: 2 });
  const first = runRibbon(app, { run: false });
  assert.deepEqual(first.errors, []);
  assert.ok(first.count > 150, 'ribbon items: ' + first.count);
  const second = runRibbon(app);
  await new Promise(r => setTimeout(r, 20));
  assert.deepEqual(second.errors, []);
  // the chart tab (a chart selected) and the picture tab
  app.objects.select(app.sheet.charts[0].id);
  app.ctx.ribbon.tabs = app.ribbonTabs();
  const chart = runRibbon(app);
  assert.deepEqual(chart.errors, []);
  for (let i = 0; i < 30 && app.wb.canUndo; i++) app.undo();
  assert.ok(app.getContent().sheets.length >= 1);
});

test('sheets: add, rename, duplicate, move, colour, hide, show, delete – references follow', async () => {
  const { app } = mount(blank());
  const cmd = app.commands;
  typeCell(app, '5');
  keydown(app.editor.input, 'Enter');
  cmd.addSheet();
  assert.equal(app.wb.sheets.length, 2);
  assert.equal(app.sheet, app.wb.sheets[1]);
  typeCell(app, '=Sheet1!A1*2');
  keydown(app.editor.input, 'Enter');
  assert.equal(app.sheet.cell(0, 0).v, 10);
  assert.ok(cmd.renameSheet(app.wb.sheets[0], 'Data 2026'));
  assert.equal(app.wb.sheets[1].cell(0, 0).f, "'Data 2026'!A1*2");
  cmd.duplicateSheet(app.wb.sheets[0]);
  assert.equal(app.wb.sheets[1].name, 'Data 2026 (2)');
  cmd.moveSheet(app.wb.sheets[1], 2);
  assert.equal(app.wb.sheets[2].name, 'Data 2026 (2)');
  cmd.tabColor(app.wb.sheets[0], '#ff0000');
  assert.equal(app.wb.sheets[0].color, '#ff0000');
  cmd.hideSheet(app.wb.sheets[2]);
  assert.ok(app.wb.sheets[2].hidden);
  cmd.unhideSheet(app.wb.sheets[2]);
  globalThis.MatUi.confirm = async () => true;
  await cmd.deleteSheet(app.wb.sheets[0]);
  globalThis.MatUi.confirm = async () => false;
  assert.equal(app.wb.sheets.length, 2);
  assert.equal(app.wb.sheets[0].cell(0, 0).f, "'Data 2026'!#REF!*2");
  app.undo();
  assert.equal(app.wb.sheets.length, 3);
  assert.equal(app.wb.sheets[1].cell(0, 0).v, 10);
});

test('charts, the AutoFilter, freezing and merging through the app', () => {
  const { app } = mount(template('sales'));
  app.sel.set(4, 1);
  app.objects.insertChart('line');
  assert.equal(app.sheet.charts.length, 2);
  const chart = app.sheet.charts[1];
  app.objects.updateChart({ title: 'Trend', legend: 'bottom' }, chart);
  assert.equal(app.sheet.charts[1].title, 'Trend');
  app.objects.select(chart.id);
  app.objects.removeSelected();
  assert.equal(app.sheet.charts.length, 1);
  app.sel.set(4, 0);
  app.filters.toggle();
  assert.ok(app.sheet.filter);
  const a = app.sheet.filterArea();
  app.filters.setColumn(a.c1, { values: [app.look.cell(app.sheet, a.r1 + 1, a.c1).text] });
  assert.ok(app.sheet.rows.filtered.size > 0);
  app.filters.open(a.c1);
  app.filters.toggle();
  assert.equal(app.sheet.rows.filtered.size, 0);
  app.commands.freeze('none');
  app.sel.set(5, 2);
  app.commands.freeze('panes');
  assert.deepEqual(app.sheet.freeze, { rows: 5, cols: 2 });
  app.commands.freeze('none');
  app.sel.setArea({ r1: 20, c1: 0, r2: 21, c2: 2 });
  app.merge('center');
  assert.ok(app.sheet.merges.some(m => m.r1 === 20 && m.c2 === 2));
  app.merge('unmerge');
  assert.ok(!app.sheet.merges.some(m => m.r1 === 20 && m.c2 === 2));
  app.notes.toggleAll();
  app.notes.toggleAll();
  app.print();
});

/** A point in the middle of a cell (or at an offset from its top left) in coordinates of the grid. */
const at = (app, row, col, dx, dy) => { const r = app.grid.cellRect(row, col); return { clientX: r.x + (dx ?? r.w / 2), clientY: r.y + (dy ?? r.h / 2) }; };
const pointer = (app, type, p, extra = {}) => app.grid.scroller.dispatch(type, { pointerId: 1, pointerType: 'mouse', button: 0, isPrimary: true, ctrlKey: false, metaKey: false, shiftKey: false, ...p, ...extra });

test('the mouse: click, drag a range, Shift+click, headers, the fill handle, moving by the border, resizing', () => {
  const { app } = mount(blank());
  for (const [r, v] of [[0, '1'], [1, '2']]) { app.sel.set(r, 0); typeCell(app, v); keydown(app.editor.input, 'Enter'); }
  pointer(app, 'pointerdown', at(app, 0, 0));
  pointer(app, 'pointermove', at(app, 3, 2));
  pointer(app, 'pointerup', at(app, 3, 2));
  assert.deepEqual(app.sel.range, { r1: 0, c1: 0, r2: 3, c2: 2 });
  pointer(app, 'pointerdown', at(app, 5, 1));
  pointer(app, 'pointerup', at(app, 5, 1));
  pointer(app, 'pointerdown', at(app, 7, 3), { shiftKey: true });
  pointer(app, 'pointerup', at(app, 7, 3));
  assert.deepEqual(app.sel.range, { r1: 5, c1: 1, r2: 7, c2: 3 });
  // a column header selects the column, a row header the row
  const col = app.grid.cellRect(0, 2);
  pointer(app, 'pointerdown', { clientX: col.x + col.w / 2, clientY: 5 });
  pointer(app, 'pointerup', { clientX: col.x + col.w / 2, clientY: 5 });
  assert.ok(app.sel.isWholeCol());
  assert.equal(app.sel.range.c1, 2);
  // the fill handle: A1:A2 (1, 2) dragged to A5 makes 1 … 5
  app.sel.setArea({ r1: 0, c1: 0, r2: 1, c2: 0 });
  app.selectionChanged();
  const a2 = app.grid.cellRect(1, 0);
  pointer(app, 'pointerdown', { clientX: a2.x + a2.w - 1, clientY: a2.y + a2.h - 1 });
  pointer(app, 'pointermove', at(app, 4, 0));
  pointer(app, 'pointerup', at(app, 4, 0));
  assert.deepEqual([0, 1, 2, 3, 4].map(r => app.sheet.cell(r, 0)?.v), [1, 2, 3, 4, 5]);
  // dragging the border of the selection moves the cells
  app.sel.setArea({ r1: 0, c1: 0, r2: 4, c2: 0 });
  app.selectionChanged();
  const top = app.grid.cellRect(0, 0);
  pointer(app, 'pointerdown', { clientX: top.x + top.w / 2, clientY: top.y + 1 });
  pointer(app, 'pointermove', at(app, 0, 3, top.w / 2, 1));
  pointer(app, 'pointerup', at(app, 0, 3, top.w / 2, 1));
  assert.equal(app.sheet.cell(0, 3)?.v, 1);
  assert.equal(app.sheet.cell(0, 0), undefined);
  // the border of a column header resizes it
  const d = app.grid.cellRect(0, 3);
  const before = app.sheet.cols.rawSize(3);
  pointer(app, 'pointerdown', { clientX: d.x + d.w - 1, clientY: 5 });
  pointer(app, 'pointermove', { clientX: d.x + d.w + 39, clientY: 5 });
  pointer(app, 'pointerup', { clientX: d.x + d.w + 39, clientY: 5 });
  assert.equal(app.sheet.cols.rawSize(3), before + 40);
  app.undo();
  assert.equal(app.sheet.cols.rawSize(3), before);
  // a double click edits the cell
  app.grid.scroller.dispatch('dblclick', at(app, 0, 3));
  assert.ok(app.editor.active);
  app.editor.cancel();
  // the context menu of a cell
  app.grid.scroller.dispatch('contextmenu', at(app, 2, 2));
  assert.deepEqual(app.sel.active, { row: 2, col: 2 });
});

test('the keys: Ctrl+arrows, PageDown, Tab inside a selection, F2, Backspace, Ctrl+A, Ctrl+D, Ctrl+;, sheets, hiding', () => {
  const { app } = mount(blank());
  for (let r = 0; r < 5; r++) { app.sel.set(r, 0); typeCell(app, String(r + 1)); keydown(app.editor.input, 'Enter'); }
  app.sel.set(0, 0);
  keydown(app.keys, 'ArrowDown', { ctrlKey: true });
  assert.deepEqual(app.sel.active, { row: 4, col: 0 });
  keydown(app.keys, 'ArrowUp', { ctrlKey: true, shiftKey: true });
  assert.deepEqual(app.sel.range, { r1: 0, c1: 0, r2: 4, c2: 0 });
  app.sel.set(2, 0);
  keydown(app.keys, 'a', { ctrlKey: true });
  assert.deepEqual(app.sel.range, { r1: 0, c1: 0, r2: 4, c2: 0 });
  keydown(app.keys, 'a', { ctrlKey: true });
  assert.ok(app.sel.isWholeRow() && app.sel.isWholeCol());
  app.sel.setArea({ r1: 0, c1: 1, r2: 1, c2: 2 });
  keydown(app.keys, 'Tab');
  assert.deepEqual(app.sel.active, { row: 0, col: 2 });
  keydown(app.keys, 'Tab');
  assert.deepEqual(app.sel.active, { row: 1, col: 1 });
  app.sel.set(0, 0);
  keydown(app.keys, 'PageDown');
  assert.ok(app.sel.active.row > 5);
  keydown(app.keys, 'Home', { ctrlKey: true });
  keydown(app.keys, 'F2');
  assert.ok(app.editor.active);
  assert.equal(app.editor.input.value, '1');
  keydown(app.editor.input, 'Escape');
  keydown(app.keys, 'Backspace');
  assert.ok(app.editor.active);
  assert.equal(app.editor.input.value, '');
  keydown(app.editor.input, 'Escape');
  app.sel.setArea({ r1: 0, c1: 0, r2: 2, c2: 0 });
  keydown(app.keys, 'd', { ctrlKey: true });
  assert.deepEqual([0, 1, 2].map(r => app.sheet.cell(r, 0).v), [1, 1, 1]);
  app.undo();
  app.sel.set(7, 0);
  keydown(app.keys, ';', { ctrlKey: true });
  assert.equal(typeof app.sheet.cell(7, 0).v, 'number');
  app.commands.addSheet();
  keydown(app.keys, 'PageUp', { ctrlKey: true });
  assert.equal(app.sheet, app.wb.sheets[0]);
  app.sel.set(2, 0);
  keydown(app.keys, '9', { ctrlKey: true });
  assert.ok(app.sheet.rows.isHidden(2));
  keydown(app.keys, '9', { ctrlKey: true, shiftKey: true });
  app.sel.setArea({ r1: 1, c1: 0, r2: 3, c2: 0 });
  keydown(app.keys, '(', { ctrlKey: true, shiftKey: true });
  assert.ok(!app.sheet.rows.isHidden(2));
});

test('read-only: typing is refused, selection and copy still work', () => {
  const { app, log } = mount(template('budget'), { readOnly: true });
  app.sel.set(0, 0);
  const plain = () => { const c = app.getContent(); for (const s of c.sheets) delete s.sel; return JSON.stringify(c); };
  const before = plain();
  app.keys.value = 'x';
  app.keys.dispatch('input');
  assert.ok(!app.editor.active);
  keydown(app.keys, 'Delete');
  keydown(app.keys, 'ArrowDown');
  assert.equal(app.sel.active.row, 1);
  assert.equal(plain(), before);
  assert.ok(log.toasts.length >= 1);
});
