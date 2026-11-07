/* MatOffice Sheets – what the ribbon, the menus and the keys do: formats, borders, number formats, rows and columns, fill series, AutoSum,
   sorting, clearing, sheets, freezing, zoom … Every change runs in one transaction of the model (one step of undo). */
import { t, confirm } from '../editor/util.js';
import { MAX_ROWS, MAX_COLS, key, keyRow, keyCol, areaName, parseArea, parseReference, colName, area as mkArea } from './engine/address.js';
import { moveFormula } from './engine/transform.js';
import { changeDecimals, presets, formatKind, isDateFormat } from './engine/numfmt.js';
import { sortCompare, numberToText } from './engine/values.js';
import { ymd, serial, fromJsDate, MONTHS, DAYS, daysInMonth } from './engine/dates.js';
import { cssFont, listItems, validationAt } from './look.js';

export class Commands {
  constructor(app) {
    this.app = app;
    this.borderStyle = 'thin';
    this.borderColor = '#000000';
  }

  get wb() { return this.app.wb; }
  get sheet() { return this.app.sheet; }
  get sel() { return this.app.sel; }

  /** Runs a change of the content: commits the editor, checks read-only and protection, wraps it into one undo step. */
  change(label, fn, { structural = false, protect = true } = {}) {
    const app = this.app;
    if (app.readOnly) { app.toastReadOnly(); return false; }
    if (app.editor.active && !app.editor.commit(0, 0)) return false;
    if (protect && this.sheet.protected && !this.selectionUnlocked()) { app.toastLocked(); return false; }
    const tx = this.wb.begin(label);
    tx.selBefore = tx.selBefore || this.sel.toJSON();
    if (structural) tx.structural = true;
    let result;
    try {
      result = fn();
    } catch (e) {
      this.wb.rollback();
      throw e;
    }
    tx.selAfter = this.sel.toJSON();
    const entry = this.wb.commit();
    app.changed({ structural: structural || !!entry?.structural });
    return result;
  }

  selectionUnlocked() {
    const s = this.sheet;
    for (const a of this.sel.ranges) {
      if (a.r2 - a.r1 > 2000 || a.c2 - a.c1 > 200) return false;
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) if (!this.wb.style(s.styleAt(r, c)).unlocked) return false;
    }
    return true;
  }

  /** The used part of an area (whole columns/rows are cut at the used range). */
  clip(a, sheet = this.sheet) {
    const b = sheet.bounds();
    const e = this.app.calc.extent(sheet.id);
    const rows = Math.max(b.rows, e.rows);
    const cols = Math.max(b.cols, e.cols);
    return { r1: a.r1, c1: a.c1, r2: Math.min(a.r2, Math.max(a.r1, rows - 1)), c2: Math.min(a.c2, Math.max(a.c1, cols - 1)) };
  }

  // ---- formats -----------------------------------------------------------------------------------------------------------------

  /** The style of the active cell. */
  get style() { return this.wb.style(this.sheet.styleAt(this.sel.active.row, this.sel.active.col)); }

  /** Changes the style of every selected cell: patch is an object of fields or a function style → style. */
  format(patch, label = 'format') {
    return this.change(label, () => this.applyStyle(this.sheet, this.sel.ranges, patch));
  }

  applyStyle(sheet, ranges, patch) {
    const wb = this.wb;
    const fn = typeof patch === 'function' ? patch : s => ({ ...s, ...patch });
    const memo = new Map();
    const next = i => { if (!memo.has(i)) memo.set(i, wb.styleIndex(fn({ ...wb.style(i) }))); return memo.get(i); };
    const setCell = (r, c) => {
      const before = sheet.styleAt(r, c);
      const after = next(before);
      const own = sheet.cell(r, c);
      if (after !== before || (own && own.s === undefined && after !== 0)) wb.patchCell(sheet, r, c, { s: after });
    };
    for (const a of ranges) {
      const wholeCols = a.r1 === 0 && a.r2 === MAX_ROWS - 1;
      const wholeRows = a.c1 === 0 && a.c2 === MAX_COLS - 1;
      if (wholeCols || wholeRows) {
        const b = sheet.bounds();
        const lines = wholeCols && wholeRows ? null : wholeCols ? [a.c1, Math.min(a.c2, a.c1 + 16383)] : [a.r1, Math.min(a.r2, a.r1 + 100000)];
        if (wholeCols) {
          wb.touchProp(sheet, 'cols');
          const [c1, c2] = lines || [0, Math.max(b.cols, 26) - 1];
          for (let c = c1; c <= c2; c++) { const s = next(sheet.colStyles.get(c) ?? 0); if (s) sheet.colStyles.set(c, s); else sheet.colStyles.delete(c); }
        }
        if (wholeRows && !wholeCols) {
          wb.touchProp(sheet, 'rows');
          const [r1, r2] = lines;
          for (let r = r1; r <= r2; r++) { const s = next(sheet.rowStyles.get(r) ?? 0); if (s) sheet.rowStyles.set(r, s); else sheet.rowStyles.delete(r); }
        }
        // the cells that already have something, and the crossings with styled rows/columns
        for (const k of [...sheet.cells.keys()]) {
          const r = keyRow(k);
          const c = keyCol(k);
          if (r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2) setCell(r, c);
        }
        if (wholeCols) for (const [r] of sheet.rowStyles) for (let c = a.c1; c <= Math.min(a.c2, a.c1 + 500); c++) if (!sheet.cell(r, c)) setCell(r, c);
        if (wholeRows && !wholeCols) for (const [c] of sheet.colStyles) for (let r = a.r1; r <= Math.min(a.r2, a.r1 + 5000); r++) if (!sheet.cell(r, c)) setCell(r, c);
        continue;
      }
      const cells = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
      if (cells > 2_000_000) continue;
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) setCell(r, c);
    }
  }

  toggle(prop) {
    const on = !this.style[prop];
    const patch = { [prop]: on || undefined };
    if (prop === 'u' && on) patch.u = true;
    this.format(patch);
  }

  setFont(font) { this.format({ font: font === 'Calibri' ? undefined : font }); }
  setSize(size) {
    const n = Number(String(size).replace(',', '.'));
    if (!(n >= 1 && n <= 409)) return;
    this.format({ size: n === 11 ? undefined : n });
    this.app.autoFitRows(this.sel.ranges);
  }
  growFont(d) {
    const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
    this.format(s => {
      const cur = s.size || 11;
      const next = d > 0 ? sizes.find(x => x > cur) ?? cur + 10 : [...sizes].reverse().find(x => x < cur) ?? Math.max(1, cur - 1);
      return { ...s, size: next === 11 ? undefined : next };
    });
    this.app.autoFitRows(this.sel.ranges);
  }
  setColor(c) { this.format({ color: c || undefined }); }
  setFill(c) { this.format({ fill: c || undefined }); }
  align(ha) { this.format(s => ({ ...s, ha: s.ha === ha ? undefined : ha })); }
  valign(va) { this.format({ va: va === 'bottom' ? undefined : va }); }
  wrapText() { this.format({ wrap: !this.style.wrap || undefined }); this.app.autoFitRows(this.sel.ranges); }
  indent(d) { this.format(s => { const n = Math.max(0, Math.min(15, (s.indent || 0) + d)); return { ...s, indent: n || undefined, ha: n && !s.ha ? 'left' : s.ha }; }); }
  numberFormat(code) { this.format({ nf: code || undefined }); }

  decimals(delta) {
    const value = this.app.calc.valueAt(this.sheet.id, this.sel.active.row, this.sel.active.col);
    const code = changeDecimals(this.style.nf, delta, value);
    this.format(s => ({ ...s, nf: s.nf === this.style.nf ? code : changeDecimals(s.nf, delta, value) }));
  }

  /** The number format list of the ribbon: the kind of the active cell's format. */
  numberKind() {
    const nf = this.style.nf;
    if (!nf) return 'general';
    const p = presets(this.app.locale);
    for (const [k, code] of Object.entries(p)) if (code === nf && ['number', 'currency', 'accounting', 'shortDate', 'longDate', 'time', 'percent', 'fraction', 'scientific', 'text'].includes(k)) return k;
    const kind = formatKind(nf);
    return { date: 'shortDate', datetime: 'shortDate', time: 'time', percent: 'percent', scientific: 'scientific', fraction: 'fraction', text: 'text', currency: 'currency', accounting: 'accounting', number: 'number' }[kind] || 'custom';
  }

  applyNumberKind(kind) {
    if (kind === 'custom') { this.app.dialogs.formatCells('number'); return; }
    const p = presets(this.app.locale);
    this.numberFormat(kind === 'general' ? null : p[kind]);
  }

  /** Borders of the selection: bottom, top, left, right, all, outside, thickOutside, thickBottom, doubleBottom, topBottom, inside, none. */
  borders(kind) {
    const line = `${this.borderStyle} ${this.borderColor}`;
    const thick = `thick ${this.borderColor}`;
    this.change('borders', () => {
      const wb = this.wb;
      const sheet = this.sheet;
      for (const a0 of this.sel.ranges) {
        const a = this.clip(a0);
        if ((a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) > 500000) continue;
        for (let r = a.r1; r <= a.r2; r++) {
          for (let c = a.c1; c <= a.c2; c++) {
            const s = { ...wb.style(sheet.styleAt(r, c)) };
            const top = r === a.r1;
            const bottom = r === a.r2;
            const left = c === a.c1;
            const right = c === a.c2;
            switch (kind) {
              case 'none': delete s.bt; delete s.bb; delete s.bl; delete s.br; break;
              case 'all': s.bt = s.bb = s.bl = s.br = line; break;
              case 'inside': if (!top) s.bt = line; if (!bottom) s.bb = line; if (!left) s.bl = line; if (!right) s.br = line; break;
              case 'outside': case 'thickOutside': { const v = kind === 'thickOutside' ? thick : line; if (top) s.bt = v; if (bottom) s.bb = v; if (left) s.bl = v; if (right) s.br = v; break; }
              case 'bottom': if (bottom) s.bb = line; break;
              case 'top': if (top) s.bt = line; break;
              case 'left': if (left) s.bl = line; break;
              case 'right': if (right) s.br = line; break;
              case 'thickBottom': if (bottom) s.bb = thick; break;
              case 'doubleBottom': if (bottom) s.bb = `double ${this.borderColor}`; break;
              case 'topBottom': if (top) s.bt = line; if (bottom) s.bb = line; break;
              case 'topThickBottom': if (top) s.bt = line; if (bottom) s.bb = thick; break;
              default: break;
            }
            const i = wb.styleIndex(s);
            if (i !== sheet.styleAt(r, c) || sheet.cell(r, c)?.s === undefined) wb.patchCell(sheet, r, c, { s: i });
          }
        }
      }
    });
  }

  /** A cell style of the gallery (Good, Bad, Heading 1 …). */
  cellStyle(style) {
    this.format(s => {
      const keep = { nf: s.nf, ha: s.ha, va: s.va, wrap: s.wrap, indent: s.indent, unlocked: s.unlocked };
      return { ...keep, ...style };
    }, 'cell style');
    this.app.autoFitRows(this.sel.ranges);
  }

  /** Format painter: remembers the styles of the selection; the next selection gets them (sticky = double click: until Esc). */
  startPainter(sticky = false) {
    const a = this.sel.range;
    const sheet = this.sheet;
    const rows = Math.min(a.r2 - a.r1 + 1, 50);
    const cols = Math.min(a.c2 - a.c1 + 1, 50);
    const styles = [];
    for (let r = 0; r < rows; r++) { const line = []; for (let c = 0; c < cols; c++) line.push(sheet.styleAt(a.r1 + r, a.c1 + c)); styles.push(line); }
    this.app.painter = { styles, sticky };
    this.app.root.classList.add('is-painting');
  }

  applyPainter() {
    const p = this.app.painter;
    if (!p) return;
    const a = this.sel.range;
    const sheet = this.sheet;
    this.change('paint format', () => {
      const area = this.clip(a);
      const rows = Math.max(area.r2 - area.r1 + 1, p.styles.length);
      const cols = Math.max(area.c2 - area.c1 + 1, p.styles[0].length);
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) this.wb.patchCell(sheet, area.r1 + r, area.c1 + c, { s: p.styles[r % p.styles.length][c % p.styles[0].length] });
    });
    if (!p.sticky) this.stopPainter();
  }

  stopPainter() {
    this.app.painter = null;
    this.app.root.classList.remove('is-painting');
    this.app.ctx.ribbon.refresh();
  }

  /** Clear: all, formats, contents, notes, links. */
  clear(kind) {
    this.change('clear', () => {
      const wb = this.wb;
      const sheet = this.sheet;
      for (const a0 of this.sel.ranges) {
        const a = this.clip(a0);
        const whole = (a0.r1 === 0 && a0.r2 === MAX_ROWS - 1) || (a0.c1 === 0 && a0.c2 === MAX_COLS - 1);
        const keys = whole || (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) > sheet.cells.size
          ? [...sheet.cells.keys()].filter(k => { const r = keyRow(k); const c = keyCol(k); return r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2; })
          : null;
        const visit = fn => { if (keys) for (const k of keys) fn(keyRow(k), keyCol(k)); else for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) fn(r, c); };
        visit((r, c) => {
          const cell = sheet.cell(r, c);
          if (!cell) return;
          if (kind === 'all') wb.setCell(sheet, r, c, undefined);
          else if (kind === 'contents') { if (cell.v !== undefined || cell.f || cell.t) wb.patchCell(sheet, r, c, { v: undefined, f: undefined, t: undefined }); } else if (kind === 'formats') { if (cell.s !== undefined) wb.patchCell(sheet, r, c, { s: undefined }); } else if (kind === 'notes') { if (cell.note) wb.patchCell(sheet, r, c, { note: undefined }); } else if (kind === 'links') { if (cell.link) wb.patchCell(sheet, r, c, { link: undefined }); }
        });
        if ((kind === 'all' || kind === 'formats') && whole) {
          if (a0.r1 === 0 && a0.r2 === MAX_ROWS - 1) { wb.touchProp(sheet, 'cols'); for (let c = a0.c1; c <= Math.min(a0.c2, a0.c1 + 16383); c++) sheet.colStyles.delete(c); }
          if (a0.c1 === 0 && a0.c2 === MAX_COLS - 1) { wb.touchProp(sheet, 'rows'); for (const r of [...sheet.rowStyles.keys()]) if (r >= a0.r1 && r <= a0.r2) sheet.rowStyles.delete(r); }
        }
        if (kind === 'all' || kind === 'formats') {
          const before = sheet.cf.length;
          const inside = rule => { const x = parseArea(rule.range); return x && x.r1 >= a0.r1 && x.r2 <= a0.r2 && x.c1 >= a0.c1 && x.c2 <= a0.c2; };
          if (sheet.cf.some(inside)) { wb.setSheetProp(sheet, 'cf', sheet.cf.filter(rule => !inside(rule))); }
          if (before !== sheet.cf.length) this.app.look.invalidate();
        }
      }
    });
  }

  // ---- cell contents ----------------------------------------------------------------------------------------------------------

  insertNow(kind) {
    const now = fromJsDate(new Date());
    const p = presets(this.app.locale);
    const value = kind === 'date' ? Math.floor(now) : now - Math.floor(now);
    const nf = kind === 'date' ? p.shortDate : p.shortTime;
    if (this.app.editor.active) {
      const input = this.app.editor.input;
      this.app.editor.insertText(input, this.app.formatText(value, nf));
      return;
    }
    this.change('insert date', () => {
      const { row, col } = this.sel.active;
      const s = this.sheet.styleAt(row, col);
      const cell = { ...(this.sheet.cell(row, col) || {}), v: value, f: undefined, t: undefined };
      if (!isDateFormat(this.wb.style(s).nf)) cell.s = this.wb.patchStyle(s, { nf });
      this.wb.setCell(this.sheet, row, col, cell);
    });
  }

  /** Ctrl+D / Ctrl+R: the first row (column) of the selection is copied into the others; a single cell copies the one above (left). */
  fillDirection(dir) {
    const a = this.sel.range;
    const single = this.sel.isCellArea(a);
    if (dir === 'down') {
      if (single && a.r1 === 0) return;
      const src = single ? { ...a, r1: a.r1 - 1, r2: a.r1 - 1 } : { ...a, r2: a.r1 };
      const dst = single ? a : { ...a, r1: a.r1 + 1 };
      this.change('fill', () => this.copyBlock(src, dst));
    } else {
      if (single && a.c1 === 0) return;
      const src = single ? { ...a, c1: a.c1 - 1, c2: a.c1 - 1 } : { ...a, c2: a.c1 };
      const dst = single ? a : { ...a, c1: a.c1 + 1 };
      this.change('fill', () => this.copyBlock(src, dst));
    }
  }

  /** Copies a block repeatedly into an area (formulas follow relative). */
  copyBlock(src, dst, sheet = this.sheet) {
    const h = src.r2 - src.r1 + 1;
    const w = src.c2 - src.c1 + 1;
    const clipDst = this.clip(dst, sheet);
    const area = (dst.r2 - dst.r1 > 100000 || dst.c2 - dst.c1 > 1000) ? clipDst : dst;
    for (let r = area.r1; r <= area.r2; r++) {
      for (let c = area.c1; c <= area.c2; c++) {
        const sr = src.r1 + (((r - src.r1) % h) + h) % h;
        const sc = src.c1 + (((c - src.c1) % w) + w) % w;
        let cell = sheet.cell(sr, sc);
        if (cell) { cell = { ...cell }; delete cell.note; if (cell.f) cell.f = moveFormula(cell.f, r - sr, c - sc); }
        this.wb.setCell(sheet, r, c, cell);
      }
    }
  }

  /** The fill handle was dragged from src to dst: series of numbers, dates, weekdays, months, numbered texts; formulas follow. */
  fill(src, dst) {
    if (dst.r2 < src.r2 && dst.r1 === src.r1 && dst.c1 === src.c1 && dst.c2 === src.c2 || dst.c2 < src.c2 && dst.c1 === src.c1 && dst.r1 === src.r1 && dst.r2 === src.r2) {
      // dragged back into the selection: clear what is left out
      this.change('clear', () => {
        const cut = dst.r2 < src.r2 ? { ...src, r1: dst.r2 + 1 } : { ...src, c1: dst.c2 + 1 };
        for (let r = cut.r1; r <= cut.r2; r++) for (let c = cut.c1; c <= cut.c2; c++) { const cell = this.sheet.cell(r, c); if (cell) this.wb.patchCell(this.sheet, r, c, { v: undefined, f: undefined, t: undefined }); }
      });
      this.sel.setArea(dst);
      this.app.selectionChanged();
      return;
    }
    this.change('fill', () => {
      const sheet = this.sheet;
      if (dst.r2 > src.r2 || dst.r1 < src.r1) {
        const down = dst.r2 > src.r2;
        for (let c = src.c1; c <= src.c2; c++) {
          const cells = [];
          for (let r = src.r1; r <= src.r2; r++) cells.push({ r, c, cell: sheet.cell(r, c) });
          const targets = [];
          if (down) for (let r = src.r2 + 1; r <= dst.r2; r++) targets.push({ r, c });
          else for (let r = src.r1 - 1; r >= dst.r1; r--) targets.push({ r, c });
          this.series(cells, targets, down ? 1 : -1);
        }
      } else {
        const right = dst.c2 > src.c2;
        for (let r = src.r1; r <= src.r2; r++) {
          const cells = [];
          for (let c = src.c1; c <= src.c2; c++) cells.push({ r, c, cell: sheet.cell(r, c) });
          const targets = [];
          if (right) for (let c = src.c2 + 1; c <= dst.c2; c++) targets.push({ r, c });
          else for (let c = src.c1 - 1; c >= dst.c1; c--) targets.push({ r, c });
          this.series(cells, targets, right ? 1 : -1);
        }
      }
    });
    this.sel.setArea(dst, this.sel.active);
    this.app.selectionChanged();
  }

  /** One line of a fill: the source cells (in order) continue into the targets (direction +1/-1). */
  series(cells, targets, dir) {
    const wb = this.wb;
    const sheet = this.sheet;
    const src = dir > 0 ? cells : [...cells].reverse();
    const n = src.length;
    const at = i => src[i % n];
    const nfOf = c => wb.style(sheet.styleAt(c.r, c.c)).nf;
    const plain = src.every(x => x.cell && !x.cell.f && typeof x.cell.v === 'number');
    const write = (target, source, value) => {
      const base = source.cell ? { ...source.cell } : undefined;
      if (base) delete base.note;
      if (value === undefined) {
        if (base?.f) base.f = moveFormula(base.f, target.r - source.r, target.c - source.c);
        wb.setCell(sheet, target.r, target.c, base);
      } else {
        wb.setCell(sheet, target.r, target.c, { ...(base || {}), v: value, f: undefined, t: undefined });
      }
    };
    // numbers and dates
    if (plain) {
      const values = src.map(x => x.cell.v);
      const date = isDateFormat(nfOf(src[0]));
      if (n === 1) {
        targets.forEach((tg, i) => write(tg, src[0], date ? values[0] + (i + 1) * dir : undefined));
        return;
      }
      if (date) {
        const parts = values.map(v => ymd(v));
        const sameDay = parts.every(p => p.d === parts[0].d);
        const months = parts.slice(1).map((p, i) => (p.y - parts[i].y) * 12 + p.m - parts[i].m);
        if (sameDay && months.every(m => m === months[0]) && months[0] !== 0 && values.every(v => v % 1 === 0)) {
          const last = parts[n - 1];
          targets.forEach((tg, i) => {
            const total = last.y * 12 + (last.m - 1) + months[0] * (i + 1);
            const y = Math.floor(total / 12);
            const m = (total % 12) + 1;
            write(tg, at(i), serial(y, m, Math.min(last.d, daysInMonth(y, m))));
          });
          return;
        }
      }
      // a linear trend (least squares, like Excel)
      const xs = values.map((_, i) => i);
      const mx = (n - 1) / 2;
      const my = values.reduce((a, b) => a + b, 0) / n;
      let num = 0;
      let den = 0;
      xs.forEach((x, i) => { num += (x - mx) * (values[i] - my); den += (x - mx) * (x - mx); });
      const slope = den ? num / den : 0;
      targets.forEach((tg, i) => write(tg, at(i), Number((my + slope * (n - 1 + i + 1 - mx)).toPrecision(15))));
      return;
    }
    // texts with a number at the end ("Item 1", "Q1"), weekday and month names
    const lists = nameLists();
    const texts = src.every(x => x.cell && !x.cell.f && typeof x.cell.v === 'string');
    if (texts) {
      const found = lists.map(list => ({ list, idx: src.map(x => list.findIndex(w => w.toLowerCase() === x.cell.v.toLowerCase())) })).find(x => x.idx.every(i => i >= 0));
      if (found) {
        const step = n > 1 ? found.idx[n - 1] - found.idx[n - 2] || 1 : dir;
        const sample = src[n - 1].cell.v;
        const caseOf = w => (sample === sample.toUpperCase() ? w.toUpperCase() : sample === sample.toLowerCase() ? w.toLowerCase() : w);
        targets.forEach((tg, i) => {
          const len = found.list.length;
          const idx = (((found.idx[n - 1] + step * (i + 1)) % len) + len) % len;
          write(tg, at(i), caseOf(found.list[idx]));
        });
        return;
      }
      const m = src.map(x => /^(.*?)(\d+)(\D*)$/.exec(x.cell.v));
      if (m.every(x => x) && m.every(x => x[1] === m[0][1] && x[3] === m[0][3])) {
        const nums = m.map(x => Number(x[2]));
        const step = n > 1 ? nums[n - 1] - nums[n - 2] : dir;
        const width = m[n - 1][2].length;
        const q = /^Q$/i.test(m[0][1]) && !m[0][3] ? 4 : 0;
        targets.forEach((tg, i) => {
          let v = nums[n - 1] + step * (i + 1);
          if (q) v = ((v - 1) % 4 + 4) % 4 + 1;
          write(tg, at(i), m[0][1] + String(Math.max(0, v)).padStart(m[n - 1][2].startsWith('0') ? width : 0, '0') + m[0][3]);
        });
        return;
      }
    }
    // everything else is repeated (formulas follow)
    targets.forEach((tg, i) => write(tg, at(i)));
  }

  /** Double click on the fill handle: fill down as far as the data next to it goes. */
  fillDownToData() {
    const a = this.sel.range;
    const sheet = this.sheet;
    const has = (r, c) => { const cell = sheet.cell(r, c); return !!cell && (cell.v !== undefined && cell.v !== '' || !!cell.f); };
    const neighbour = a.c1 > 0 && has(a.r2 + 1, a.c1 - 1) ? a.c1 - 1 : has(a.r2 + 1, a.c2 + 1) ? a.c2 + 1 : -1;
    if (neighbour < 0) return;
    let last = a.r2 + 1;
    while (last + 1 < MAX_ROWS && has(last + 1, neighbour)) last++;
    this.fill({ ...a }, { ...a, r2: last });
  }

  /** Drag-moving the selection (Ctrl copies). */
  moveSelection(dr, dc, copy) {
    const a = this.sel.range;
    const sheet = this.sheet;
    const target = { r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 + dr, c2: a.c2 + dc };
    if (copy) {
      this.change('copy', () => {
        const snap = [];
        for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) snap.push([r, c, sheet.cell(r, c)]);
        for (const [r, c, cell] of snap) {
          const next = cell ? { ...cell } : undefined;
          if (next?.f) next.f = moveFormula(next.f, dr, dc);
          this.wb.setCell(sheet, r + dr, c + dc, next);
        }
      });
    } else {
      this.change('move', () => this.wb.moveCells(sheet, a, sheet, dr, dc), { structural: true });
    }
    this.sel.setArea(target, { row: this.sel.active.row + dr, col: this.sel.active.col + dc });
    this.app.selectionChanged();
  }

  /** AutoSum (and Average, Count, Max, Min): below a column of numbers or right of a row; a single cell gets a formula to confirm. */
  autoSum(fn = 'SUM') {
    const app = this.app;
    if (app.readOnly) return;
    if (app.editor.active) { app.editor.insertText(app.editor.input, fn + '('); return; }
    const sheet = this.sheet;
    const a = this.sel.range;
    const num = (r, c) => typeof app.calc.valueAt(sheet.id, r, c) === 'number' && !sheet.cell(r, c)?.f?.match(/^SUBTOTAL|^SUM\(/);
    const isNum = (r, c) => typeof app.calc.valueAt(sheet.id, r, c) === 'number';
    if (this.sel.isSingle) {
      const { row, col } = this.sel.active;
      let range = null;
      if (row > 0 && isNum(row - 1, col)) {
        let r = row - 1;
        while (r > 0 && isNum(r - 1, col) && num(r - 1, col)) r--;
        range = { r1: r, c1: col, r2: row - 1, c2: col };
      } else if (col > 0 && isNum(row, col - 1)) {
        let c = col - 1;
        while (c > 0 && isNum(row, c - 1)) c--;
        range = { r1: row, c1: c, r2: row, c2: col - 1 };
      }
      const text = '=' + fn + '(' + (range ? areaName(range) : '') + ')';
      app.editor.begin({ text, mode: 'enter' });
      if (!range) { const input = app.editor.input; input.setSelectionRange(text.length - 1, text.length - 1); app.editor.afterCaret(input); }
      return;
    }
    this.change('autosum', () => {
      const area = this.clip(a);
      const sep = ',';
      for (let c = area.c1; c <= area.c2; c++) {
        const lastEmpty = !sheet.cell(area.r2, c)?.v && !sheet.cell(area.r2, c)?.f;
        const target = lastEmpty && area.r2 > area.r1 ? area.r2 : area.r2 + 1;
        const end = target - 1;
        if (end < area.r1) continue;
        this.wb.patchCell(sheet, target, c, { f: `${fn}(${areaName({ r1: area.r1, c1: c, r2: end, c2: c })})`, v: undefined, t: undefined });
      }
      void sep;
    });
  }

  // ---- navigation ------------------------------------------------------------------------------------------------------------

  /** Ctrl+arrow: the edge of the data in a direction. */
  edge(sheet, row, col, dr, dc) {
    const calc = this.app.calc;
    const filled = (r, c) => { const v = calc.valueAt(sheet.id, r, c); return v !== null && v !== undefined && v !== ''; };
    const b = sheet.bounds();
    const e = calc.extent(sheet.id);
    const lastRow = Math.max(b.rows, e.rows) - 1;
    const lastCol = Math.max(b.cols, e.cols) - 1;
    const step = (r, c) => ({ r: dr ? sheet.rows.step(r, dr) : r, c: dc ? sheet.cols.step(c, dc) : c });
    let r = row;
    let c = col;
    let n = step(r, c);
    if (n.r === r && n.c === c) return { row, col };
    if (filled(r, c) && filled(n.r, n.c)) {
      for (;;) {
        const m = step(r, c);
        if ((m.r === r && m.c === c) || !filled(m.r, m.c)) break;
        r = m.r;
        c = m.c;
      }
      return { row: r, col: c };
    }
    for (;;) {
      n = step(r, c);
      if (n.r === r && n.c === c) return { row: r, col: c };
      r = n.r;
      c = n.c;
      if (filled(r, c)) return { row: r, col: c };
      if (dr > 0 && r > lastRow) return { row: MAX_ROWS - 1, col: c };
      if (dc > 0 && c > lastCol) return { row: r, col: MAX_COLS - 1 };
      if (dr < 0 && r === 0) return { row: 0, col: c };
      if (dc < 0 && c === 0) return { row: r, col: 0 };
    }
  }

  /** The block of data around a cell (Ctrl+A, sort, filter, a new chart). */
  currentRegion(sheet = this.sheet, row = this.sel.active.row, col = this.sel.active.col) {
    const calc = this.app.calc;
    const filled = (r, c) => { if (r < 0 || c < 0) return false; const v = calc.valueAt(sheet.id, r, c); return v !== null && v !== undefined && v !== ''; };
    let a = { r1: row, c1: col, r2: row, c2: col };
    for (let guard = 0; guard < 10000; guard++) {
      let grown = false;
      const rowHas = r => { for (let c = Math.max(0, a.c1 - 1); c <= a.c2 + 1; c++) if (filled(r, c)) return true; return false; };
      const colHas = c => { for (let r = Math.max(0, a.r1 - 1); r <= a.r2 + 1; r++) if (filled(r, c)) return true; return false; };
      if (a.r1 > 0 && rowHas(a.r1 - 1)) { a.r1--; grown = true; }
      if (a.r2 < MAX_ROWS - 1 && rowHas(a.r2 + 1)) { a.r2++; grown = true; }
      if (a.c1 > 0 && colHas(a.c1 - 1)) { a.c1--; grown = true; }
      if (a.c2 < MAX_COLS - 1 && colHas(a.c2 + 1)) { a.c2++; grown = true; }
      if (!grown) break;
    }
    return a;
  }

  selectAll() {
    const region = this.currentRegion();
    const a = this.sel.range;
    if (!this.sel.isSingle && a.r1 === region.r1 && a.r2 === region.r2 && a.c1 === region.c1 && a.c2 === region.c2 || (region.r1 === region.r2 && region.c1 === region.c2)) this.sel.all();
    else this.sel.setArea(region, this.sel.active);
    this.app.selectionChanged({ scroll: false });
  }

  /** The name box: an address, a range, a name or a sheet reference; a new name defines the selection. */
  goToText(text) {
    const app = this.app;
    const v = String(text || '').trim();
    if (!v) { app.focusGrid(); return; }
    const ref = parseReference(v);
    if (ref) {
      const sheet = ref.sheet ? app.wb.sheetByName(ref.sheet) : app.sheet;
      if (sheet) { this.goTo(sheet, ref.area); return; }
    }
    const named = app.wb.nameArea(v);
    if (named) { this.goTo(named.sheet, named.area); return; }
    if (!app.readOnly && /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(v)) {
      this.change('define name', () => app.wb.setNames([...app.wb.names, { name: v, ref: absoluteRefOf(app.sheet, this.sel.range) }]), { protect: false });
      app.focusGrid();
      return;
    }
    app.ctx.toast(t('Sheets.Goto.Invalid'), 'danger');
    app.editor.refresh();
  }

  goTo(sheet, a) {
    if (sheet === undefined) { this.app.dialogs.goTo(); return; }
    const app = this.app;
    if (sheet !== app.sheet) app.showSheet(sheet);
    app.sel.setArea(a);
    app.selectionChanged({ scrollTo: { row: a.r1, col: a.c1 } });
    app.focusGrid();
  }

  followLink(link) {
    if (link.startsWith('#')) {
      this.goToText(link.slice(1));
      return;
    }
    if (/^(https?:|mailto:)/i.test(link)) window.open(link, '_blank', 'noopener,noreferrer');
  }

  jumpFromBorder() { /* a double click on the border of the selection: keep it simple, no jump */ }

  nextSheet(d) {
    const app = this.app;
    const visible = app.wb.sheets.filter(s => !s.hidden);
    const i = visible.indexOf(app.sheet);
    const next = visible[(i + d + visible.length) % visible.length];
    if (next && next !== app.sheet) app.showSheet(next);
  }

  // ---- rows and columns ----------------------------------------------------------------------------------------------------

  /** The rows (or columns) of the selection as [from, to] spans. */
  spans(axis) {
    const out = [];
    for (const a of this.sel.ranges) out.push(axis === 'row' ? [a.r1, a.r2] : [a.c1, a.c2]);
    out.sort((x, y) => y[0] - x[0]);
    return out;
  }

  insertLines(axis, after = false) {
    const a = this.sel.range;
    const count = axis === 'row' ? a.r2 - a.r1 + 1 : a.c2 - a.c1 + 1;
    if (count >= (axis === 'row' ? MAX_ROWS : MAX_COLS)) return;
    const at = axis === 'row' ? (after ? a.r2 + 1 : a.r1) : (after ? a.c2 + 1 : a.c1);
    this.change(axis === 'row' ? 'insert rows' : 'insert columns', () => this.wb.shiftLines(this.sheet, axis, at, count), { structural: true });
  }

  deleteLines(axis) {
    const spans = this.spans(axis);
    if (spans.some(([lo, hi]) => hi - lo + 1 >= (axis === 'row' ? MAX_ROWS : MAX_COLS))) return;
    this.change(axis === 'row' ? 'delete rows' : 'delete columns', () => {
      for (const [lo, hi] of spans) this.wb.shiftLines(this.sheet, axis, lo, -(hi - lo + 1));
    }, { structural: true });
    const a = this.sel.range;
    this.sel.set(axis === 'row' ? Math.min(a.r1, MAX_ROWS - 1) : this.sel.active.row, axis === 'col' ? Math.min(a.c1, MAX_COLS - 1) : this.sel.active.col);
    this.app.selectionChanged();
  }

  /** Insert cells: shift down / right, or whole rows / columns. */
  insertCells(mode) {
    if (mode === 'rows') return this.insertLines('row');
    if (mode === 'cols') return this.insertLines('col');
    const a = this.sel.range;
    const sheet = this.sheet;
    const b = sheet.bounds();
    this.change('insert cells', () => {
      if (mode === 'down') {
        const block = { r1: a.r1, c1: a.c1, r2: Math.max(a.r1, b.rows - 1), c2: a.c2 };
        if (b.rows > a.r1) this.wb.moveCells(sheet, block, sheet, a.r2 - a.r1 + 1, 0);
      } else {
        const block = { r1: a.r1, c1: a.c1, r2: a.r2, c2: Math.max(a.c1, b.cols - 1) };
        if (b.cols > a.c1) this.wb.moveCells(sheet, block, sheet, 0, a.c2 - a.c1 + 1);
      }
    }, { structural: true });
  }

  deleteCells(mode) {
    if (mode === 'rows') return this.deleteLines('row');
    if (mode === 'cols') return this.deleteLines('col');
    const a = this.sel.range;
    const sheet = this.sheet;
    const b = sheet.bounds();
    this.change('delete cells', () => {
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) this.wb.setCell(sheet, r, c, undefined);
      if (mode === 'up' && b.rows > a.r2 + 1) this.wb.moveCells(sheet, { r1: a.r2 + 1, c1: a.c1, r2: b.rows - 1, c2: a.c2 }, sheet, -(a.r2 - a.r1 + 1), 0);
      if (mode === 'left' && b.cols > a.c2 + 1) this.wb.moveCells(sheet, { r1: a.r1, c1: a.c2 + 1, r2: a.r2, c2: b.cols - 1 }, sheet, 0, -(a.c2 - a.c1 + 1));
    }, { structural: true });
  }

  insertDialog() {
    if (this.sel.isWholeRow()) return this.insertLines('row');
    if (this.sel.isWholeCol()) return this.insertLines('col');
    this.app.dialogs.shiftCells('insert');
  }

  deleteDialog() {
    if (this.sel.isWholeRow()) return this.deleteLines('row');
    if (this.sel.isWholeCol()) return this.deleteLines('col');
    this.app.dialogs.shiftCells('delete');
  }

  /** Hide (or unhide) the rows / columns of the selection. */
  hideLines(axis, hide) {
    this.change(hide ? 'hide' : 'unhide', () => {
      const sheet = this.sheet;
      this.wb.touchProp(sheet, axis === 'row' ? 'rows' : 'cols');
      const ax = axis === 'row' ? sheet.rows : sheet.cols;
      for (const [lo, hi] of this.spans(axis)) {
        const from = hide ? lo : Math.max(0, lo - 1);
        const to = hide ? hi : Math.min((axis === 'row' ? MAX_ROWS : MAX_COLS) - 1, hi + 1);
        if (to - from > 100000) continue;
        for (let i = from; i <= to; i++) {
          ax.setHidden(i, hide);
          if (!hide && ax.rawSize(i) === 0) ax.setSize(i, ax.def);
        }
      }
    }, { protect: false });
    if (hide) {
      const sheet = this.sheet;
      const { row, col } = this.sel.active;
      this.sel.set(sheet.rows.isHidden(row) ? sheet.rows.step(row, 1) : row, sheet.cols.isHidden(col) ? sheet.cols.step(col, 1) : col);
      this.app.selectionChanged();
    }
  }

  /** Width of columns / height of rows (px). */
  setSize(axis, px) {
    this.change(axis === 'row' ? 'row height' : 'column width', () => {
      const sheet = this.sheet;
      this.wb.touchProp(sheet, axis === 'row' ? 'rows' : 'cols');
      const ax = axis === 'row' ? sheet.rows : sheet.cols;
      for (const [lo, hi] of this.spans(axis)) for (let i = lo; i <= Math.min(hi, lo + 100000); i++) { ax.setSize(i, px); if (axis === 'row') this.app.manualRows.add(i); }
    }, { protect: false });
  }

  setDefaultWidth(px) {
    this.change('default width', () => { this.wb.touchProp(this.sheet, 'cols'); this.sheet.cols.setDefault(px); }, { protect: false });
  }

  /** Fit the width of a column (or every selected column) to its content; rows fit their height. */
  autofit(axis, index) {
    const sel = this.sel;
    const indexes = [];
    const whole = axis === 'col' ? sel.isWholeCol() : sel.isWholeRow();
    const a = sel.range;
    if (whole && (axis === 'col' ? index >= a.c1 && index <= a.c2 : index >= a.r1 && index <= a.r2)) {
      for (let i = axis === 'col' ? a.c1 : a.r1; i <= (axis === 'col' ? a.c2 : a.r2) && indexes.length < 2000; i++) indexes.push(i);
    } else indexes.push(index);
    if (axis === 'col') this.fitColumns(indexes); else this.app.autoFitRows(indexes.map(r => ({ r1: r, r2: r, c1: 0, c2: MAX_COLS - 1 })), true);
  }

  fitColumns(cols) {
    const app = this.app;
    const sheet = this.sheet;
    this.change('autofit', () => {
      this.wb.touchProp(sheet, 'cols');
      const widths = new Map();
      for (const [k] of sheet.cells) {
        const c = keyCol(k);
        if (!cols.includes(c)) continue;
        const r = keyRow(k);
        if (sheet.mergeAt(r, c)) continue;
        const look = app.look.cell(sheet, r, c, 11);
        if (!look.text || look.wrap) continue;
        const w = app.grid.measure(cssFont({ ...look.style, b: look.b, i: look.i }, 1), look.text) + 8 + look.indent * 9;
        widths.set(c, Math.max(widths.get(c) || 0, w));
      }
      for (const c of cols) {
        const w = widths.get(c);
        if (w) sheet.cols.setSize(c, Math.min(1000, Math.ceil(w)));
      }
    }, { protect: false });
  }

  // ---- sorting ------------------------------------------------------------------------------------------------------------------

  /** The area a sort applies to: the selection, or the data around the active cell; and whether its first row is a header. */
  sortArea() {
    const a = this.sel.range;
    const single = this.sel.isCellArea(a);
    const area = single ? this.currentRegion() : this.clip(a);
    const sheet = this.sheet;
    const calc = this.app.calc;
    // a header when the first row is text over a column with numbers or a different kind
    let header = false;
    if (single && area.r2 > area.r1) {
      for (let c = area.c1; c <= area.c2; c++) {
        const top = calc.valueAt(sheet.id, area.r1, c);
        const below = calc.valueAt(sheet.id, area.r1 + 1, c);
        if (typeof top === 'string' && (typeof below === 'number' || typeof below === 'boolean')) header = true;
      }
      if (!header && this.wb.style(sheet.styleAt(area.r1, area.c1)).b && !this.wb.style(sheet.styleAt(area.r1 + 1, area.c1)).b) header = true;
    }
    if (sheet.filter) { const f = sheet.filterArea(); if (f && f.c1 <= this.sel.active.col && f.c2 >= this.sel.active.col) return { area: f, header: true }; }
    return { area, header };
  }

  /** Sorts by the active column (A→Z or Z→A). */
  sort(desc) {
    const { area, header } = this.sortArea();
    this.sortBy(area, header, [{ col: this.sel.active.col, desc }]);
  }

  sortBy(area, header, keys) {
    const sheet = this.sheet;
    const calc = this.app.calc;
    const body = { ...area, r1: area.r1 + (header ? 1 : 0) };
    if (body.r2 <= body.r1) return;
    this.change('sort', () => {
      this.wb.sortRows(sheet, body, keys, (r, c) => calc.valueAt(sheet.id, r, c), sortCompare);
    });
  }

  // ---- sheets ------------------------------------------------------------------------------------------------------------------

  addSheet() {
    const app = this.app;
    let created;
    this.change('add sheet', () => {
      const n = app.wb.sheets.length + 1;
      created = app.wb.addSheet(t('Sheets.SheetName', n), app.wb.active + 1);
    }, { protect: false, structural: true });
    if (created) app.showSheet(created);
  }

  async deleteSheet(sheet = this.sheet) {
    const app = this.app;
    if (app.readOnly) return;
    if (app.wb.sheets.filter(s => !s.hidden).length <= 1) { app.ctx.toast(t('Sheets.Sheet.Last'), 'warning'); return; }
    const hasData = sheet.cells.size > 0 || sheet.charts.length || sheet.images.length;
    if (hasData && !(await confirm(t('Sheets.Sheet.DeleteConfirm', sheet.name), { danger: true, ok: t('Sheets.Delete') }))) return;
    this.change('delete sheet', () => app.wb.deleteSheet(sheet), { protect: false, structural: true });
    app.showSheet(app.wb.activeSheet, { force: true });
  }

  renameSheet(sheet, name) {
    const app = this.app;
    let ok = false;
    this.change('rename sheet', () => { ok = app.wb.renameSheet(sheet, name); }, { protect: false, structural: true });
    if (!ok && String(name).trim() !== sheet.name) app.ctx.toast(t('Sheets.Sheet.BadName'), 'danger');
    app.tabs.render();
    return ok;
  }

  duplicateSheet(sheet = this.sheet) {
    let copy;
    this.change('duplicate sheet', () => { copy = this.app.wb.duplicateSheet(sheet); }, { protect: false, structural: true });
    if (copy) this.app.showSheet(copy);
  }

  moveSheet(sheet, to) {
    this.change('move sheet', () => this.app.wb.moveSheet(sheet, to), { protect: false });
    this.app.tabs.render();
  }

  tabColor(sheet, color) {
    this.change('tab colour', () => this.wb.setSheetProp(sheet, 'color', color || null), { protect: false });
    this.app.tabs.render();
  }

  hideSheet(sheet = this.sheet) {
    const app = this.app;
    if (app.wb.sheets.filter(s => !s.hidden).length <= 1) { app.ctx.toast(t('Sheets.Sheet.Last'), 'warning'); return; }
    this.change('hide sheet', () => this.wb.setSheetProp(sheet, 'hidden', true), { protect: false });
    if (sheet === app.sheet) this.nextSheet(1);
    app.tabs.render();
  }

  unhideSheet(sheet) {
    this.change('unhide sheet', () => this.wb.setSheetProp(sheet, 'hidden', false), { protect: false });
    this.app.showSheet(sheet);
  }

  protect() {
    const sheet = this.sheet;
    this.change('protect', () => this.wb.setSheetProp(sheet, 'protected', !sheet.protected), { protect: false });
    this.app.ctx.toast(t(sheet.protected ? 'Sheets.Protect.On' : 'Sheets.Protect.Off'), 'info');
  }

  /** Locks or unlocks the selected cells (they stay editable when the sheet is protected). */
  toggleLocked() {
    const unlocked = !this.style.unlocked;
    this.change('lock', () => this.applyStyle(this.sheet, this.sel.ranges, s => ({ ...s, unlocked: unlocked || undefined })), { protect: false });
  }

  // ---- view ---------------------------------------------------------------------------------------------------------------------

  freeze(kind) {
    const sheet = this.sheet;
    const { row, col } = this.sel.active;
    const value = kind === 'none' ? { rows: 0, cols: 0 } : kind === 'row' ? { rows: 1, cols: 0 } : kind === 'col' ? { rows: 0, cols: 1 } : { rows: row, cols: col };
    if (kind === 'panes' && sheet.freeze.rows + sheet.freeze.cols > 0) { value.rows = 0; value.cols = 0; }
    this.change('freeze', () => this.wb.setSheetProp(sheet, 'freeze', value), { protect: false });
    this.app.grid.scroller.scrollTop = 0;
    this.app.grid.scroller.scrollLeft = 0;
  }

  toggleView(prop) {
    const sheet = this.sheet;
    this.change('view', () => this.wb.setSheetProp(sheet, prop, !sheet[prop]), { protect: false });
  }

  setZoom(f) {
    const sheet = this.sheet;
    const z = Math.max(0.1, Math.min(4, Math.round(f * 100) / 100));
    if (z === sheet.zoom) return;
    if (this.app.readOnly) sheet.zoom = z;
    else { this.wb.begin('zoom'); this.wb.setSheetProp(sheet, 'zoom', z); this.wb.commit(); this.app.ctx.changed(); }
    this.app.zoomChanged();
  }

  zoomBy(d) { this.setZoom(this.sheet.zoom + d); }

  /** Zoom so that the selection fills the window. */
  zoomToSelection() {
    const a = this.clip(this.sel.range);
    const sheet = this.sheet;
    const grid = this.app.grid;
    const w = sheet.cols.offset(a.c2 + 1) - sheet.cols.offset(a.c1);
    const h = sheet.rows.offset(a.r2 + 1) - sheet.rows.offset(a.r1);
    const z = Math.min((grid.width - 60) / Math.max(1, w), (grid.height - 30) / Math.max(1, h));
    this.setZoom(Math.max(0.1, Math.min(4, z)));
    grid.scrollIntoView(a.r1, a.c1);
  }

  toggleFormulas() {
    this.app.showFormulas = !this.app.showFormulas;
    this.app.look.invalidate();
    this.app.grid.requestRender();
  }

  recalc() {
    this.app.calc.rebuild();
    this.app.look.invalidate();
    this.app.afterCalc();
  }

  // ---- page layout --------------------------------------------------------------------------------------------------------------

  setPrint(patch) {
    const sheet = this.sheet;
    const next = { ...(sheet.print || {}), ...patch };
    for (const k of Object.keys(next)) if (next[k] === undefined || next[k] === false || next[k] === 'none') delete next[k];
    this.change('page setup', () => this.wb.setSheetProp(sheet, 'print', next), { protect: false });
  }

  // ---- dropdowns of cells -------------------------------------------------------------------------------------------------------

  /** Alt+Down: the list of a validation or the filter menu of a header cell. */
  openDropdown() {
    const app = this.app;
    const { row, col } = this.sel.active;
    const fa = this.sheet.filterArea();
    if (fa && row === fa.r1 && col >= fa.c1 && col <= fa.c2) { app.filters.open(col); return; }
    const rule = validationAt(this.sheet, row, col);
    if (rule?.type === 'list') app.openList(rule);
  }

  listItemsAt(row, col) {
    const rule = validationAt(this.sheet, row, col);
    return rule?.type === 'list' ? listItems(this.app, this.sheet, rule) : [];
  }

  // ---- menus --------------------------------------------------------------------------------------------------------------------

  namesMenu(anchor) {
    const app = this.app;
    const items = app.wb.names.length
      ? app.wb.names.map(n => ({ label: n.name, icon: 'tag', edit: false, run: () => this.goToText(n.name) }))
      : [{ label: t('Sheets.Names.None'), disabled: () => true, edit: false }];
    items.push('-', { label: t('Sheets.NameManager'), icon: 'sheets-names', run: () => app.dialogs.nameManager() });
    app.menuAt(anchor, items);
  }

  insertFunction() {
    this.app.dialogs.insertFunction();
  }

  editNote() { this.app.notes.edit(); }

  /** The context menu of cells, column headers or row headers. */
  contextMenu(kind) {
    const app = this.app;
    const clip = app.clipboard;
    const items = [
      { label: t('Sheets.Cut'), icon: 'sheets-cut', key: 'Mod+X', run: () => clip.cutCommand() },
      { label: t('Sheets.Copy'), icon: 'copy', key: 'Mod+C', edit: false, run: () => clip.copyCommand() },
      { label: t('Sheets.Paste'), icon: 'sheets-paste', key: 'Mod+V', run: () => clip.pasteCommand() },
      { label: t('Sheets.PasteSpecial'), icon: 'sheets-paste-special', run: () => app.dialogs.pasteSpecial() },
      '-',
    ];
    if (kind === 'cols') {
      items.push(
        { label: t('Sheets.InsertCols'), icon: 'sheets-insert-col', run: () => this.insertLines('col') },
        { label: t('Sheets.DeleteCols'), icon: 'sheets-delete-col', run: () => this.deleteLines('col') },
        { label: t('Sheets.ClearContents'), icon: 'broom', run: () => this.clear('contents') },
        '-',
        { label: t('Sheets.FormatCells'), icon: 'sheets-format-cells', run: () => app.dialogs.formatCells() },
        { label: t('Sheets.ColWidth'), icon: 'sheets-col-width', run: () => app.dialogs.size('col') },
        { label: t('Sheets.Hide'), icon: 'eye-off', run: () => this.hideLines('col', true) },
        { label: t('Sheets.Unhide'), icon: 'eye', run: () => this.hideLines('col', false) },
      );
    } else if (kind === 'rows') {
      items.push(
        { label: t('Sheets.InsertRows'), icon: 'sheets-insert-row', run: () => this.insertLines('row') },
        { label: t('Sheets.DeleteRows'), icon: 'sheets-delete-row', run: () => this.deleteLines('row') },
        { label: t('Sheets.ClearContents'), icon: 'broom', run: () => this.clear('contents') },
        '-',
        { label: t('Sheets.FormatCells'), icon: 'sheets-format-cells', run: () => app.dialogs.formatCells() },
        { label: t('Sheets.RowHeight'), icon: 'sheets-row-height', run: () => app.dialogs.size('row') },
        { label: t('Sheets.Hide'), icon: 'eye-off', run: () => this.hideLines('row', true) },
        { label: t('Sheets.Unhide'), icon: 'eye', run: () => this.hideLines('row', false) },
      );
    } else {
      const { row, col } = this.sel.active;
      const cell = this.sheet.cell(row, col);
      items.push(
        { label: t('Sheets.InsertCells'), icon: 'sheets-insert-cells', run: () => this.insertDialog() },
        { label: t('Sheets.DeleteCells'), icon: 'sheets-delete-cells', run: () => this.deleteDialog() },
        { label: t('Sheets.ClearContents'), icon: 'broom', run: () => this.clear('contents') },
        '-',
        { label: t('Sheets.SortAZ'), icon: 'sheets-sort-az', run: () => this.sort(false) },
        { label: t('Sheets.SortZA'), icon: 'sheets-sort-za', run: () => this.sort(true) },
        { label: t('Sheets.FilterBySelection'), icon: 'filter', run: () => app.filters.byValue() },
        '-',
        { label: cell?.note ? t('Sheets.EditNote') : t('Sheets.NewNote'), icon: 'comment', run: () => app.notes.edit() },
        cell?.note ? { label: t('Sheets.DeleteNote'), icon: 'trash', run: () => this.clear('notes') } : null,
        { label: t('Sheets.FormatCells'), icon: 'sheets-format-cells', key: 'Mod+1', run: () => app.dialogs.formatCells() },
        { label: t('Sheets.PickFromList'), icon: 'list', hidden: () => !this.listItemsAt(row, col).length, run: () => this.openDropdown() },
        { label: t('Sheets.DefineName'), icon: 'tag', run: () => app.dialogs.defineName() },
        { label: cell?.link ? t('Sheets.EditLink') : t('Sheets.Link'), icon: 'link', key: 'Mod+K', run: () => app.dialogs.link() },
        cell?.link ? { label: t('Sheets.OpenLink'), icon: 'external', edit: false, run: () => this.followLink(cell.link) } : null,
      );
    }
    return items.filter(Boolean);
  }
}

/** The absolute reference of an area on a sheet: Sheet1!$A$1:$B$2. */
export function absoluteRefOf(sheet, a) {
  const cell = (r, c) => '$' + colName(c) + '$' + (r + 1);
  const body = a.r1 === 0 && a.r2 === MAX_ROWS - 1 ? `$${colName(a.c1)}:$${colName(a.c2)}`
    : a.c1 === 0 && a.c2 === MAX_COLS - 1 ? `$${a.r1 + 1}:$${a.r2 + 1}`
      : a.r1 === a.r2 && a.c1 === a.c2 ? cell(a.r1, a.c1) : cell(a.r1, a.c1) + ':' + cell(a.r2, a.c2);
  return (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) && !/^[A-Za-z]{1,3}\d+$/.test(sheet.name) ? sheet.name : "'" + sheet.name.replace(/'/g, "''") + "'") + '!' + body;
}

/** The lists of names a fill continues (weekdays and months, English and German, long and short). */
function nameLists() {
  return [MONTHS.en, MONTHS.de, MONTHS.en.map(m => m.slice(0, 3)), ['Jan', 'Feb', 'Mrz', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
    DAYS.en, DAYS.de, DAYS.en.map(d => d.slice(0, 3)), ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa']];
}

export { mkArea, key, numberToText };
