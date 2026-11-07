// MatOffice Sheets – the editor of a cell and the formula bar, driven like a person types: point mode, F4, function suggestions,
// argument help, Enter/Tab/Esc, the text of the formula bar.
import './fake-browser.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../../../src/MatOffice/wwwroot/js/sheets/model.js';
import { Selection } from '../../../src/MatOffice/wwwroot/js/sheets/selection.js';
import { CellEditor } from '../../../src/MatOffice/wwwroot/js/sheets/editor.js';

function makeEditor() {
  const wb = Workbook.load({ version: 1, active: 0, styles: [{}], sheets: [{ id: 's1', name: 'Sheet1', cells: { A1: { v: 5 } } }, { id: 's2', name: 'Other sheet', cells: {} }] });
  const log = [];
  const sel = new Selection();
  const app = {
    wb, sel, locale: 'en', readOnly: false, keepEditing: false,
    get sheet() { return wb.activeSheet; },
    isLocked: () => false, toastLocked() {}, editText: () => '', nameBoxText: () => 'A1', setMode: m => log.push('mode:' + m),
    grid: { highlights: [], width: 800, height: 600, cellRect: () => ({ x: 10, y: 10, w: 64, h: 20 }), metrics: () => ({}), requestRender() {}, scrollIntoView() {} },
    root: { getBoundingClientRect: () => ({ left: 0, top: 0 }) },
    commands: { edge: (s, r, c, dr, dc) => ({ row: r + dr * 5, col: c + dc * 5 }) },
    commitInput: (sheet, row, col, text) => { log.push(`commit ${sheet.name}!${row},${col}=${text}`); return true; },
    afterCommit: (dr, dc) => log.push(`move ${dr},${dc}`),
    focusGrid() {}, showSheet() {},
  };
  const editor = new CellEditor(app);
  app.editor = editor;
  return { app, editor, log };
}

const key = (editor, input, k, extra = {}) => editor.onKey({ key: k, preventDefault() {}, ...extra }, input);
const typeInto = (editor, input, text) => { input.setRangeText(text, input.selectionStart, input.selectionEnd, 'end'); editor.onInput(input); };

test('typing a formula, pointing at cells and on another sheet, Enter commits', () => {
  const { app, editor, log } = makeEditor();
  app.sel.set(4, 2); // C5
  editor.begin({ text: '=', mode: 'enter' });
  const input = editor.input;
  assert.ok(editor.canPoint());
  editor.pointStart(app.sheet, 0, 0);
  assert.equal(input.value, '=A1');
  editor.pointExtend(app.sheet, 2, 1);
  assert.equal(input.value, '=A1:B3');
  typeInto(editor, input, '+');
  assert.ok(editor.canPoint());
  editor.pointStart(app.wb.sheets[1], 1, 1);
  assert.equal(input.value, "=A1:B3+'Other sheet'!B2");
  assert.equal(editor.barInput.value, input.value);
  assert.deepEqual(app.grid.highlights.map(h => h.sheet), ['s1', 's2']);
  key(editor, input, 'Enter');
  assert.equal(log.filter(l => l.startsWith('commit'))[0], "commit Sheet1!4,2==A1:B3+'Other sheet'!B2");
  assert.ok(log.includes('move 1,0'));
  assert.equal(editor.active, false);
});

test('arrow keys point in Enter mode and leave the cell otherwise', () => {
  const { app, editor, log } = makeEditor();
  app.sel.set(9, 3);
  editor.begin({ text: '=', mode: 'enter' });
  key(editor, editor.input, 'ArrowUp');
  assert.equal(editor.input.value, '=D9');
  key(editor, editor.input, 'ArrowLeft', { shiftKey: true });
  assert.equal(editor.input.value, '=C9:D9');
  editor.cancel();
  editor.begin({ text: '42', mode: 'enter' });
  key(editor, editor.input, 'ArrowRight');
  assert.ok(log.includes('commit Sheet1!9,3=42'));
  assert.ok(log.includes('move 0,1'));
});

test('F4 cycles the dollars of the reference at the caret', () => {
  const { editor } = makeEditor();
  editor.begin({ text: '=A1+B2', mode: 'edit' });
  const input = editor.input;
  input.setSelectionRange(2, 2);
  const seen = [];
  for (let i = 0; i < 4; i++) { key(editor, input, 'F4'); seen.push(input.value); input.setSelectionRange(2, 2); }
  assert.deepEqual(seen, ['=$A$1+B2', '=A$1+B2', '=$A1+B2', '=A1+B2']);
});

test('function names are suggested and the arguments are shown', () => {
  const { editor } = makeEditor();
  editor.begin({ text: '=ROU', mode: 'enter' });
  assert.equal(editor.popup.hidden, false);
  assert.ok(editor.suggest.items.some(i => i.name === 'ROUND'));
  editor.suggest.index = editor.suggest.items.findIndex(i => i.name === 'ROUND');
  key(editor, editor.input, 'Tab');
  assert.equal(editor.input.value, '=ROUND(');
  assert.equal(editor.tip.hidden, false);
  typeInto(editor, editor.input, 'A1, ');
  const bold = editor.tip.children.find(c => c.tagName === 'B');
  assert.equal(bold.textContent, 'num_digits');
  key(editor, editor.input, 'Escape');
  assert.equal(editor.active, false);
});

test('Alt+Enter is a line break, Tab commits to the right, the formula bar shows the cell', () => {
  const { app, editor, log } = makeEditor();
  app.editText = () => '=SUM(A1:A3)';
  editor.refresh();
  assert.equal(editor.barInput.value, '=SUM(A1:A3)');
  editor.begin({ text: 'a', mode: 'enter' });
  key(editor, editor.input, 'Enter', { altKey: true });
  typeInto(editor, editor.input, 'b');
  assert.equal(editor.input.value, 'a\nb');
  key(editor, editor.input, 'Tab');
  assert.ok(log.includes('commit Sheet1!0,0=a\nb'));
  assert.ok(log.includes('move 0,1'));
});
