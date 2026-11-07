// MatOffice Sheets – a small workbook in memory for the tests of the formula engine (node --test "tests/js/sheets/*.test.mjs").
import { Calc } from '../../../src/MatOffice/wwwroot/js/sheets/engine/calc.js';
import '../../../src/MatOffice/wwwroot/js/sheets/engine/functions/index.js';
import { parseCell, key, keyRow, keyCol } from '../../../src/MatOffice/wwwroot/js/sheets/engine/address.js';
import { errorOf, FormulaError, setLocale } from '../../../src/MatOffice/wwwroot/js/sheets/engine/values.js';

/** The fixed "now" of the tests: 2026-01-15 12:00 local time (serial 46037.5). */
export const NOW = new Date(2026, 0, 15, 12, 0, 0);

function toCell(value) {
  if (typeof value === 'string' && value.startsWith('=')) return { f: value.slice(1) };
  if (typeof value === 'string' && errorOf(value)) return { v: errorOf(value).code, t: 'e' };
  return { v: value };
}

/**
 * A workbook: workbook({ Sheet1: { A1: 1, B1: '=A1*2' }, 'Other sheet': { … } }, [{ name: 'Rate', ref: 'Sheet1!$A$1' }]).
 * Values: numbers, strings, booleans, '#N/A' (an error), '=…' (a formula).
 */
export function workbook(spec, names = [], { locale = 'en' } = {}) {
  setLocale(locale);
  const sheets = Object.keys(spec);
  const data = new Map();
  for (const [name, cells] of Object.entries(spec)) {
    const m = new Map();
    for (const [addr, val] of Object.entries(cells)) {
      const { row, col } = parseCell(addr);
      m.set(key(row, col), toCell(val));
    }
    data.set(name, m);
  }
  const host = {
    cell: (s, r, c) => data.get(s)?.get(key(r, c)),
    sheetId: name => sheets.find(n => n.toLowerCase() === String(name).toLowerCase()) ?? null,
    sheetName: id => id,
    sheetIds: () => sheets,
    sheetIndex: id => sheets.indexOf(id),
    bounds: s => {
      let rows = 0;
      let cols = 0;
      for (const k of data.get(s)?.keys() || []) { rows = Math.max(rows, keyRow(k) + 1); cols = Math.max(cols, keyCol(k) + 1); }
      return { rows, cols };
    },
    formulas: visit => { for (const [s, m] of data) for (const [k, c] of m) if (c.f) visit(s, keyRow(k), keyCol(k), c.f); },
    names: () => names,
    rowHidden: () => false,
    setResult: (s, r, c, v) => {
      const cell = data.get(s).get(key(r, c));
      if (v instanceof FormulaError) { cell.v = v.code; cell.t = 'e'; } else { cell.v = v; delete cell.t; }
    },
  };
  const calc = new Calc(host);
  calc.now = () => NOW;
  let seed = 1;
  calc.random = () => { seed = (seed * 16807) % 2147483647; return (seed - 1) / 2147483646; };
  calc.rebuild();
  const place = (addr, sheet) => ({ sheet: sheet ?? sheets[0], ...parseCell(addr) });
  return {
    calc, host, data,
    /** The value of a cell as formulas see it (errors as their code). */
    get(addr, sheet) {
      const p = place(addr, sheet);
      const v = calc.valueAt(p.sheet, p.row, p.col);
      return v instanceof FormulaError ? v.code : v;
    },
    /** Changes cells: set({ A1: 5, B2: '=A1+1', C3: null }) and recalculates. */
    set(cells, sheet) {
      const positions = [];
      for (const [addr, val] of Object.entries(cells)) {
        const p = place(addr, sheet);
        const m = data.get(p.sheet);
        if (val === null || val === undefined) m.delete(key(p.row, p.col));
        else m.set(key(p.row, p.col), toCell(val));
        positions.push(p);
      }
      calc.update(positions);
    },
  };
}

/** The result of one formula ("=…" or without "=") over some cells of Sheet1: an error as its code, an array as a 2-D array. */
export function evaluate(formula, cells = {}, options = {}) {
  const f = formula.startsWith('=') ? formula : '=' + formula;
  const wb = workbook({ Sheet1: { ...cells, ZZ1000: f } }, options.names || [], options);
  const fc = wb.calc.formulaAt('Sheet1', 999, 701);
  if (fc && fc.array) return fc.array.map(row => row.map(v => (v instanceof FormulaError ? v.code : v)));
  return wb.get('ZZ1000');
}
