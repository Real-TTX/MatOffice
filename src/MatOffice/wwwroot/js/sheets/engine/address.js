/* MatOffice Sheets – cell addresses in A1 notation (pure, no DOM; tested in tests/js/sheets).
   Rows and columns are 0-based inside the editor (row 0 = "1", column 0 = "A"); the content and the formulas use A1 text. */

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** A number key of a cell inside one sheet (row * 16384 + col): fast keys for Maps. */
export const key = (row, col) => row * MAX_COLS + col;
export const keyRow = k => Math.floor(k / MAX_COLS);
export const keyCol = k => k % MAX_COLS;

/** 0 → "A", 25 → "Z", 26 → "AA". */
export function colName(col) {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" → 0, "xfd" → 16383; -1 for anything else. */
export function colIndex(name) {
  if (!name || name.length > 3) return -1;
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const k = (name.charCodeAt(i) | 32) - 96;
    if (k < 1 || k > 26) return -1;
    n = n * 26 + k;
  }
  return n <= MAX_COLS ? n - 1 : -1;
}

export const cellName = (row, col) => colName(col) + (row + 1);

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

/** "B3" or "$B$3" → { row: 2, col: 1 }; null when it is not a cell of the sheet. */
export function parseCell(text) {
  const m = CELL.exec(String(text).trim());
  if (!m) return null;
  const col = colIndex(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || row < 0 || row >= MAX_ROWS || m[2][0] === '0') return null;
  return { row, col };
}

/** An area { r1, c1, r2, c2 } (inclusive, ordered). */
export function area(r1, c1, r2 = r1, c2 = c1) {
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/** "A1", "A1:C3", "A:C" (whole columns), "2:5" (whole rows), with or without $ → an area; null otherwise. */
export function parseArea(text) {
  const t = String(text ?? '').trim().replace(/\$/g, '');
  if (!t) return null;
  const parts = t.split(':');
  if (parts.length === 1) {
    const c = parseCell(parts[0]);
    return c ? area(c.row, c.col) : null;
  }
  if (parts.length !== 2) return null;
  const [a, b] = parts;
  const ca = parseCell(a);
  const cb = parseCell(b);
  if (ca && cb) return area(ca.row, ca.col, cb.row, cb.col);
  if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
    const x = colIndex(a);
    const y = colIndex(b);
    return x < 0 || y < 0 ? null : area(0, x, MAX_ROWS - 1, y);
  }
  if (/^\d{1,7}$/.test(a) && /^\d{1,7}$/.test(b)) {
    const x = Number(a) - 1;
    const y = Number(b) - 1;
    return x < 0 || y < 0 || x >= MAX_ROWS || y >= MAX_ROWS ? null : area(x, 0, y, MAX_COLS - 1);
  }
  return null;
}

/** An area as text: "B3", "A1:C3", "A:C" (whole columns), "2:5" (whole rows). */
export function areaName(a) {
  if (a.r1 === 0 && a.r2 === MAX_ROWS - 1) return colName(a.c1) + ':' + colName(a.c2);
  if (a.c1 === 0 && a.c2 === MAX_COLS - 1) return (a.r1 + 1) + ':' + (a.r2 + 1);
  const first = cellName(a.r1, a.c1);
  return a.r1 === a.r2 && a.c1 === a.c2 ? first : first + ':' + cellName(a.r2, a.c2);
}

export const areaContains = (a, row, col) => row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2;
export const areasOverlap = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
export const areaCells = a => (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
export const sameArea = (a, b) => !!a && !!b && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;

/** The smallest area around both. */
export const areaUnion = (a, b) => area(Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));

/** The overlap of two areas, or null. */
export function areaIntersect(a, b) {
  if (!areasOverlap(a, b)) return null;
  return { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
}

/** Whether a sheet name must be quoted in a reference ('My sheet'!A1). */
export function needsQuotes(name) {
  if (!/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name)) return true;
  if (parseCell(name)) return true; // a name like "A1" or "AB12"
  if (/^(R\d*C\d*|R\d*|C\d*)$/i.test(name)) return true; // R1C1-like
  return /^(TRUE|FALSE)$/i.test(name);
}

/** The sheet part of a reference: Sheet1! or 'My sheet'!. */
export const sheetPrefix = name => (needsQuotes(name) ? "'" + name.replace(/'/g, "''") + "'" : name) + '!';

/** "Sheet1!A1:B2" / "'My sheet'!$A$1" → { sheet, area }; a reference without sheet gives sheet null. */
export function parseReference(text) {
  const t = String(text ?? '').trim();
  let sheet = null;
  let rest = t;
  if (t.startsWith("'")) {
    let i = 1;
    let name = '';
    while (i < t.length) {
      if (t[i] === "'") {
        if (t[i + 1] === "'") { name += "'"; i += 2; continue; }
        break;
      }
      name += t[i++];
    }
    if (t[i] !== "'" || t[i + 1] !== '!') return null;
    sheet = name;
    rest = t.slice(i + 2);
  } else {
    const bang = t.lastIndexOf('!');
    if (bang >= 0) { sheet = t.slice(0, bang); rest = t.slice(bang + 1); }
  }
  const a = parseArea(rest);
  return a ? { sheet, area: a } : null;
}

/** A reference as text, with the sheet when one is given. */
export const referenceName = (sheet, a) => (sheet ? sheetPrefix(sheet) : '') + areaName(a);

/** Absolute reference text ($A$1:$B$2) of an area. */
export function absoluteName(a) {
  const cell = (r, c) => '$' + colName(c) + '$' + (r + 1);
  if (a.r1 === 0 && a.r2 === MAX_ROWS - 1) return '$' + colName(a.c1) + ':$' + colName(a.c2);
  if (a.c1 === 0 && a.c2 === MAX_COLS - 1) return '$' + (a.r1 + 1) + ':$' + (a.r2 + 1);
  return a.r1 === a.r2 && a.c1 === a.c2 ? cell(a.r1, a.c1) : cell(a.r1, a.c1) + ':' + cell(a.r2, a.c2);
}
