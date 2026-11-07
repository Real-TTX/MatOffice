// MatOffice Sheets – the templates of the app calculate without errors (AVERAGE COUNTA COUNTIF COUNTIFS EDATE EOMONTH IF IFERROR INDEX IPMT
// MATCH MAX OR PMT PPMT ROUND ROW SUM TODAY WEEKDAY YEAR). "Now" is the harness's 2026-01-15 12:00.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { workbook } from './harness.mjs';
import { serial } from '../../../src/MatOffice/wwwroot/js/sheets/engine/dates.js';

const folder = new URL('../../../src/MatOffice/Resources/Templates/sheets/', import.meta.url);
/** A text of a template ({ en, de } or one for both) in English. */
const english = v => (v && typeof v === 'object' ? v.en : v);

/** A template as a workbook of the harness (the English texts). */
function open(file) {
  const json = JSON.parse(readFileSync(new URL(file, folder), 'utf8'));
  const spec = {};
  for (const sheet of json.sheets) {
    const cells = {};
    for (const [addr, cell] of Object.entries(sheet.cells || {})) {
      if (cell.f) cells[addr] = '=' + english(cell.f);
      else if (cell.v !== undefined && cell.v !== null) cells[addr] = english(cell.v);
    }
    spec[english(sheet.name)] = cells;
  }
  return { wb: workbook(spec), json };
}

test('no formula of any template gives an error', () => {
  const files = readdirSync(folder).filter(f => f.endsWith('.json'));
  assert.ok(files.length >= 6);
  for (const file of files) {
    const { wb, json } = open(file);
    for (const sheet of json.sheets) {
      for (const [addr, cell] of Object.entries(sheet.cells || {})) {
        if (!cell.f) continue;
        const v = wb.get(addr, english(sheet.name));
        assert.ok(!(typeof v === 'string' && /^#[A-Z/0-9]+[!?]?$/.test(v) && v !== '#'), `${file} ${addr} =${english(cell.f)} gives ${v}`);
      }
    }
  }
});

test('the loan template: payment, schedule and the balance at the end', () => {
  const { wb } = open('loan.json');
  const r = 0.049 / 12;
  const payment = (25000 * r) / (1 - (1 + r) ** -60);
  assert.equal(wb.get('B8'), serial(2026, 2, 1)); // EOMONTH(TODAY(),0)+1
  assert.ok(Math.abs(wb.get('B10') - payment) < 1e-9);
  assert.equal(wb.get('B11'), 60);
  assert.equal(wb.get('B16'), serial(2026, 2, 1)); // EDATE of the first payment
  assert.equal(wb.get('B17'), serial(2026, 3, 1));
  assert.ok(Math.abs(wb.get('D16') - 25000 * r) < 1e-9); // the first interest (IPMT)
  assert.ok(Math.abs(wb.get('D16') + wb.get('E16') - payment) < 1e-9); // IPMT + PPMT = PMT
  assert.ok(Math.abs(wb.get('F75')) < 1e-6); // paid off after 60 payments
  assert.equal(wb.get('A76'), ''); // no 61st payment
});

test('the task, sales, invoice and schedule templates', () => {
  const tasks = open('tasks.json').wb;
  assert.equal(tasks.get('J4'), 5); // COUNTIF open
  assert.equal(tasks.get('J5'), 2); // in progress
  assert.equal(tasks.get('J6'), 1); // done
  assert.equal(tasks.get('J7'), 1); // COUNTIFS overdue and not done
  assert.equal(tasks.get('J8'), 1 / 8); // done / COUNTA
  assert.equal(tasks.get('F5'), -2); // days left
  assert.equal(tasks.get('F4'), ''); // done: no days left
  const sales = open('sales.json').wb;
  assert.equal(sales.get('B2'), 2026); // YEAR(TODAY())
  assert.equal(sales.get('B13'), 'South'); // INDEX/MATCH of the largest total
  assert.equal(sales.get('B14'), 'Online'); // and of the fastest growth
  const invoice = open('invoice.json').wb;
  assert.equal(invoice.get('A14'), 1); // ROW()-13
  assert.equal(invoice.get('E14'), 760); // ROUND(8*95,2)
  assert.equal(invoice.get('E16'), 118.8);
  assert.equal(invoice.get('A17'), '');
  const schedule = open('schedule.json').wb;
  assert.equal(schedule.get('B2'), serial(2026, 1, 12)); // the Monday of the week (WEEKDAY …, 2)
});
