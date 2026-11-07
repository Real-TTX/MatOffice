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
