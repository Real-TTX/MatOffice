/* MatOffice Sheets – formulas written again after the sheet changed (pure, no DOM): inserted and deleted rows/columns, a formula copied or
   filled to another cell (relative references move), cells cut and pasted elsewhere (references to them follow), a sheet renamed or deleted.
   Only the reference tokens change; the rest of the text (spaces, names, strings) stays as it was. */
import { tokenize, printTokens } from './lexer.js';
import { MAX_ROWS, MAX_COLS } from './address.js';

const same = (a, b) => a !== null && b !== null && a !== undefined && b !== undefined && a.toLowerCase() === b.toLowerCase();

/** Applies change(token) to every reference token; change returns true when it changed the token. Returns the new text, or the old one. */
export function rewrite(formula, change) {
  let tokens;
  try { tokens = tokenize(formula); } catch { return formula; }
  let changed = false;
  for (const t of tokens) {
    if (t.type === 'ref' && change(t)) { t.dirty = true; changed = true; }
  }
  return changed ? printTokens(tokens) : formula;
}

const bad = t => { t.kind = 'bad'; return true; };

/**
 * Rows or columns inserted (count > 0) or deleted (count < 0) at index `at` (0-based) of sheet `sheetName`; the formula is on sheet `hostName`.
 * Inserted inside a range, the range grows; deleted cells become #REF!; ranges shrink.
 */
export function shiftFormula(formula, { hostName, sheetName, axis, at, count }) {
  return rewrite(formula, t => {
    if (t.kind === 'bad') return false;
    if (!same(t.sheet ?? hostName, sheetName)) return false;
    if (axis === 'row' && t.kind === 'cols') return false;
    if (axis === 'col' && t.kind === 'rows') return false;
    const [k1, k2, max] = axis === 'row' ? ['r1', 'r2', MAX_ROWS] : ['c1', 'c2', MAX_COLS];
    const lo = Math.min(t[k1], t[k2]);
    const hi = Math.max(t[k1], t[k2]);
    let nlo = lo;
    let nhi = hi;
    if (count > 0) {
      if (lo >= at) nlo = lo + count;
      if (hi >= at) nhi = hi + count;
      if (nhi >= max) {
        if (nlo >= max) return bad(t);
        nhi = max - 1;
      }
    } else {
      const n = -count;
      const end = at + n - 1;
      if (lo >= at && hi <= end) return bad(t);
      nlo = lo < at ? lo : lo > end ? lo - n : at;
      nhi = hi < at ? hi : hi > end ? hi - n : at - 1;
      if (nhi < nlo) return bad(t);
    }
    if (nlo === lo && nhi === hi) return false;
    const swapped = t[k1] > t[k2];
    t[k1] = swapped ? nhi : nlo;
    t[k2] = swapped ? nlo : nhi;
    return true;
  });
}

/** A formula copied from one cell to another: its relative references move by (dRow, dCol); a reference that leaves the sheet becomes #REF!. */
export function moveFormula(formula, dRow, dCol) {
  if (!dRow && !dCol) return formula;
  return rewrite(formula, t => {
    if (t.kind === 'bad') return false;
    let changed = false;
    const shift = (k, abs, d, max) => {
      if (t[abs] || !d) return true;
      const v = t[k] + d;
      if (v < 0 || v >= max) return false;
      t[k] = v;
      changed = true;
      return true;
    };
    let ok = true;
    if (t.kind !== 'cols') ok = shift('r1', 'ar1', dRow, MAX_ROWS) && shift('r2', 'ar2', dRow, MAX_ROWS) && ok;
    if (t.kind !== 'rows') ok = shift('c1', 'ac1', dCol, MAX_COLS) && shift('c2', 'ac2', dCol, MAX_COLS) && ok;
    if (!ok) return bad(t);
    return changed;
  });
}

/** A sheet renamed: references to it get the new name. */
export function renameSheetIn(formula, oldName, newName) {
  return rewrite(formula, t => {
    if (!same(t.sheet, oldName)) return false;
    t.sheet = newName;
    return true;
  });
}

/** A sheet deleted: references to it become #REF!. */
export function deleteSheetIn(formula, name) {
  return rewrite(formula, t => (same(t.sheet, name) ? bad(t) : false));
}

/**
 * Cells moved (cut and pasted, dragged): references into the area `from` of sheet `fromName` that lie completely inside it now point to the new place
 * (dRow, dCol, on sheet `toName`). The formula is on sheet `hostName`.
 */
export function moveAreaIn(formula, { hostName, fromName, from, toName, dRow, dCol }) {
  return rewrite(formula, t => {
    if (t.kind === 'bad' || t.kind === 'cols' || t.kind === 'rows') return false;
    if (!same(t.sheet ?? hostName, fromName)) return false;
    const r1 = Math.min(t.r1, t.r2);
    const r2 = Math.max(t.r1, t.r2);
    const c1 = Math.min(t.c1, t.c2);
    const c2 = Math.max(t.c1, t.c2);
    if (r1 < from.r1 || r2 > from.r2 || c1 < from.c1 || c2 > from.c2) return false;
    if (r1 + dRow < 0 || r2 + dRow >= MAX_ROWS || c1 + dCol < 0 || c2 + dCol >= MAX_COLS) return bad(t);
    t.r1 += dRow; t.r2 += dRow; t.c1 += dCol; t.c2 += dCol;
    if (!same(toName, fromName)) t.sheet = same(toName, hostName) ? null : toName;
    return true;
  });
}

/**
 * A formula that moves to another sheet with its cell (cut and paste between sheets): references without a sheet keep pointing to the old sheet,
 * references to the new sheet lose their prefix.
 */
export function rehostFormula(formula, fromName, toName) {
  if (same(fromName, toName)) return formula;
  return rewrite(formula, t => {
    if (t.sheet === null || t.sheet === undefined) { t.sheet = fromName; return true; }
    if (same(t.sheet, toName)) { t.sheet = null; return true; }
    return false;
  });
}

/** The references of a formula (for the colours while editing): [{ start, end, sheet, r1, c1, r2, c2 }] with positions in the text. */
export function referencesOf(formula, locale = 'en') {
  let tokens;
  try { tokens = tokenize(formula, { locale }); } catch { return []; }
  return tokens.filter(t => t.type === 'ref' && t.kind !== 'bad').map(t => ({
    start: t.pos, end: t.pos + t.text.length, sheet: t.sheet ?? null,
    r1: Math.min(t.r1, t.r2), c1: Math.min(t.c1, t.c2), r2: Math.max(t.r1, t.r2), c2: Math.max(t.c1, t.c2),
  }));
}
