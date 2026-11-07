/* MatOffice Sheets – the selection: one or more ranges (Ctrl+click adds one), the active cell inside the last range and the anchor that
   Shift+click / Shift+arrows extend from. Merged cells are always selected whole (pure, no DOM). */
import { MAX_ROWS, MAX_COLS, area, areaContains } from './engine/address.js';

export class Selection {
  constructor() {
    this.ranges = [area(0, 0)];
    this.active = { row: 0, col: 0 };
    this.anchor = { row: 0, col: 0 };
    this.mergeAt = () => null;
  }

  get range() { return this.ranges[this.ranges.length - 1]; }
  get isSingle() { return this.ranges.length === 1 && this.isCellArea(this.range); }

  /** Whether an area is one cell or exactly one merged area. */
  isCellArea(a) {
    if (a.r1 === a.r2 && a.c1 === a.c2) return true;
    const m = this.mergeAt(a.r1, a.c1);
    return !!m && m.r1 === a.r1 && m.c1 === a.c1 && m.r2 === a.r2 && m.c2 === a.c2;
  }

  /** The area grown until no merged area sticks out of it. */
  grow(a) {
    let r = { ...a };
    for (let guard = 0; guard < 50; guard++) {
      let changed = false;
      for (let row = r.r1; row <= r.r2 && row <= r.r1 + 2000; row++) {
        for (const col of row === r.r1 || row === r.r2 ? range(r.c1, r.c2) : [r.c1, r.c2]) {
          const m = this.mergeAt(row, col);
          if (m && (m.r1 < r.r1 || m.r2 > r.r2 || m.c1 < r.c1 || m.c2 > r.c2)) {
            r = { r1: Math.min(r.r1, m.r1), c1: Math.min(r.c1, m.c1), r2: Math.max(r.r2, m.r2), c2: Math.max(r.c2, m.c2) };
            changed = true;
          }
        }
      }
      if (!changed) break;
    }
    return r;
  }

  /** One cell (or its merged area). */
  set(row, col) {
    const m = this.mergeAt(row, col);
    const r = m ? m.r1 : row;
    const c = m ? m.c1 : col;
    this.active = { row: r, col: c };
    this.anchor = { row: r, col: c };
    this.ranges = [m ? { ...m } : area(r, c)];
  }

  /** The last range from the anchor to a cell (Shift+click, Shift+arrows); the active cell stays. */
  extendTo(row, col) {
    const a = this.grow(area(this.anchor.row, this.anchor.col, row, col));
    this.ranges[this.ranges.length - 1] = a;
  }

  /** Adds a range (Ctrl+click). */
  add(row, col) {
    const m = this.mergeAt(row, col);
    this.active = { row: m ? m.r1 : row, col: m ? m.c1 : col };
    this.anchor = { ...this.active };
    this.ranges.push(m ? { ...m } : area(this.active.row, this.active.col));
  }

  /** Selects an area; the active cell is its first cell (or keeps a given one). */
  setArea(a, active) {
    const g = this.grow(a);
    this.ranges = [g];
    this.active = active && areaContains(g, active.row, active.col) ? { ...active } : { row: g.r1, col: g.c1 };
    this.anchor = { ...this.active };
  }

  /** Whole rows r1..r2 (the active cell goes to the first visible column of the first row). */
  rows(r1, r2, add = false, col = 0) {
    const a = area(r1, 0, r2, MAX_COLS - 1);
    if (add) this.ranges.push(a); else this.ranges = [a];
    this.active = { row: Math.min(r1, r2), col };
    this.anchor = { row: r1, col };
  }

  cols(c1, c2, add = false, row = 0) {
    const a = area(0, c1, MAX_ROWS - 1, c2);
    if (add) this.ranges.push(a); else this.ranges = [a];
    this.active = { row, col: Math.min(c1, c2) };
    this.anchor = { row, col: c1 };
  }

  all() {
    this.ranges = [area(0, 0, MAX_ROWS - 1, MAX_COLS - 1)];
    this.active = { row: 0, col: 0 };
    this.anchor = { row: 0, col: 0 };
  }

  contains(row, col) { return this.ranges.some(a => areaContains(a, row, col)); }

  /** Whether the whole column / row is in a selected range. */
  isWholeCol(a = this.range) { return a.r1 === 0 && a.r2 === MAX_ROWS - 1; }
  isWholeRow(a = this.range) { return a.c1 === 0 && a.c2 === MAX_COLS - 1; }

  /**
   * Enter/Tab inside a selection of several cells: the active cell walks through it (rows first for Tab, columns first for Enter) and wraps.
   * Returns true when it moved inside the selection.
   */
  cycle(dr, dc, visibleRow = () => true, visibleCol = () => true) {
    const a = this.range;
    if (this.ranges.length === 1 && this.isCellArea(a)) return false;
    let { row, col } = this.active;
    for (let guard = 0; guard < 100000; guard++) {
      if (dc) {
        col += dc;
        if (col > a.c2) { col = a.c1; row = row + 1 > a.r2 ? a.r1 : row + 1; } else if (col < a.c1) { col = a.c2; row = row - 1 < a.r1 ? a.r2 : row - 1; }
      } else {
        row += dr;
        if (row > a.r2) { row = a.r1; col = col + 1 > a.c2 ? a.c1 : col + 1; } else if (row < a.r1) { row = a.r2; col = col - 1 < a.c1 ? a.c2 : col - 1; }
      }
      const m = this.mergeAt(row, col);
      if (m && (m.r1 !== row || m.c1 !== col)) continue;
      if (visibleRow(row) && visibleCol(col)) break;
    }
    this.active = { row, col };
    return true;
  }

  toJSON() { return { ranges: this.ranges.map(r => ({ ...r })), active: { ...this.active }, anchor: { ...this.anchor } }; }

  restore(s) {
    if (!s) return;
    this.ranges = s.ranges.map(r => ({ ...r }));
    this.active = { ...s.active };
    this.anchor = { ...s.anchor };
  }
}

function range(a, b) {
  const out = [];
  for (let i = a; i <= b && out.length < 2000; i++) out.push(i);
  return out;
}
