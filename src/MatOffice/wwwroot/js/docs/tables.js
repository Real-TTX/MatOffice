/* MatOffice Docs – tables: insert (grid picker, dialog), rows and columns, merge and split, table styles, header row, shading, borders,
   vertical alignment, distribute columns, selecting, Tab from cell to cell and column widths by dragging a border. Every change reads the table
   into the model (htmlio.readTable), changes it there and renders it again; the caret goes back into the same cell. */
import { readTable, renderBlock, domLength } from './htmlio.js';
import { newTable, tableGrid, numStr, TABLE_STYLES, textBlock } from './model.js';

export class Tables {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    this.setupResize();
  }

  /** The table, row, cell (elements) of the selection, or null. */
  current() {
    const sel = this.ed.selection();
    if (!sel) return null;
    const cell = sel.start.block.closest('td,th');
    if (!cell || !sel.host.contains(cell)) return null;
    const table = cell.closest('table');
    const endCell = sel.end.block.closest('td,th');
    return { table, row: cell.parentNode, cell, endCell: endCell && endCell.closest('table') === table ? endCell : cell, sel };
  }

  /** The position of a cell element in the model: { r, i } (row index, index in its row) and its grid column. */
  locate(table, cellEl) {
    const rows = rowsOf(table);
    const r = rows.indexOf(cellEl.parentNode);
    const i = [...cellEl.parentNode.children].filter(c => /^(TD|TH)$/.test(c.tagName)).indexOf(cellEl);
    return { r, i };
  }

  /** Changes the table through the model: fn(model, info) → { r, i } of the cell for the caret (or null), then renders it. */
  change(fn, kind = 'table') {
    const cur = this.current();
    if (!cur) return;
    this.ed.checkpoint(kind);
    const model = readTable(cur.table);
    const at = this.locate(cur.table, cur.cell);
    const end = this.locate(cur.table, cur.endCell);
    const grid = tableGrid(model);
    const colOf = (r, i) => grid[r]?.findIndex(slot => slot && slot.cell === model.rows[r].cells[i]) ?? 0;
    const info = { at: { ...at, c: colOf(at.r, at.i) }, end: { ...end, c: colOf(end.r, end.i) }, grid };
    const result = fn(model, info);
    if (result === false) return;
    if (!model.rows.length) { this.removeTable(cur.table); return; }
    normalizeCols(model);
    const fresh = renderBlock(model, this.ed.renderOptions());
    cur.table.replaceWith(fresh);
    const target = result && typeof result === 'object' ? result : { r: Math.min(at.r, model.rows.length - 1), i: 0 };
    this.focusCell(fresh, target.r, target.i, target.end);
  }

  focusCell(table, r, i, atEnd = false) {
    const rows = rowsOf(table);
    const row = rows[Math.max(0, Math.min(r, rows.length - 1))];
    if (!row) return;
    const cells = [...row.children].filter(c => /^(TD|TH)$/.test(c.tagName));
    const cell = cells[Math.max(0, Math.min(i, cells.length - 1))];
    const blocks = cell ? [...cell.querySelectorAll('p,h1,h2,h3,h4,h5,h6,blockquote,pre')] : [];
    const b = atEnd ? blocks[blocks.length - 1] : blocks[0];
    if (b) this.ed.select(b, atEnd ? domLength(b) : 0);
  }

  removeTable(table) {
    const next = table.nextElementSibling;
    const container = table.parentNode;
    table.remove();
    this.ed.ensureTail(container);
    const target = next && next.matches('p,h1,h2,h3,h4,h5,h6,blockquote,pre') ? next : this.ed.textBlocks(this.app.body)[0];
    if (target) this.ed.select(target, 0);
  }

  // ---- insert -----------------------------------------------------------------------------------------------------------------------

  insert(rows, cols, options = {}) {
    this.ed.focus();
    if (!this.ed.selection()) {
      const first = this.ed.textBlocks(this.app.body)[0];
      if (first) this.ed.select(first, 0);
    }
    const table = newTable(Math.max(1, Math.min(500, rows)), Math.max(1, Math.min(63, cols)), options);
    this.ed.insertBlocks([table, textBlock()]);
    // the caret into the first cell
    const sel = this.ed.selection();
    const prev = sel?.start.block.previousElementSibling;
    const t = prev?.matches('table') ? prev : sel?.start.block.closest('table');
    if (t) this.focusCell(t, 0, 0);
  }

  // ---- rows and columns -------------------------------------------------------------------------------------------------------------

  insertRow(below) {
    this.change((m, { at, end, grid }) => {
      const r = below ? Math.max(at.r, end.r) : Math.min(at.r, end.r);
      const ref = grid[r];
      const width = ref.length;
      const row = { header: false, cells: [] };
      const target = below ? r + 1 : r;
      const seen = new Set();
      for (let c = 0; c < width; c++) {
        const slot = ref[c];
        if (!slot) { row.cells.push(emptyCell()); continue; }
        const spansAcross = below ? slot.row + slot.cell.rowspan - 1 > r : slot.row < r;
        if (spansAcross) {
          if (!seen.has(slot.cell)) { slot.cell.rowspan += 1; seen.add(slot.cell); }
          c += slot.cell.colspan - 1 - (c - slot.col);
          continue;
        }
        const cell = emptyCell();
        cell.colspan = slot.cell.colspan - (c - slot.col);
        cell.css = { ...slot.cell.css };
        delete cell.css['background-color'];
        row.cells.push(cell);
        c += cell.colspan - 1;
      }
      m.rows.splice(target, 0, row);
      if (m.rows[target + 1]?.header && !below) row.header = true;
      return { r: target, i: 0 };
    }, 'row');
  }

  insertColumn(right) {
    this.change((m, { at, end, grid }) => {
      const col = right ? Math.max(at.c + (m.rows[at.r].cells[at.i].colspan - 1), end.c) + 1 : Math.min(at.c, end.c);
      const n = grid[0]?.length || 1;
      const seen = new Set();
      grid.forEach((slots, r) => {
        const inside = slots[col] && slots[col].col < col ? slots[col] : null;
        if (inside) { if (!seen.has(inside.cell)) { inside.cell.colspan += 1; seen.add(inside.cell); } return; }
        if (slots[col] && slots[col].row !== r) return; // covered by a cell from a row above
        const before = slots.slice(0, col).filter((s, c) => s && s.row === r && s.col === c).length;
        const cell = emptyCell(m.rows[r].header);
        m.rows[r].cells.splice(before, 0, cell);
        if (slots[col]?.cell.rowspan > 1 && slots[col].row === r) cell.rowspan = slots[col].cell.rowspan;
      });
      const widths = colWidths(m, n);
      const share = 100 / (n + 1);
      const scaled = widths.map(w => w * (100 - share) / 100);
      scaled.splice(col, 0, share);
      m.cols = scaled.map(w => numStr(w) + '%');
      return { r: at.r, i: right ? at.i + 1 : at.i };
    }, 'column');
  }

  deleteRow() {
    this.change((m, { at, end, grid }) => {
      const from = Math.min(at.r, end.r), to = Math.max(at.r, end.r);
      for (let r = to; r >= from; r--) {
        const slots = grid[r];
        const done = new Set();
        for (let c = 0; c < slots.length; c++) {
          const slot = slots[c];
          if (!slot || done.has(slot.cell)) continue;
          done.add(slot.cell);
          if (slot.row < r) slot.cell.rowspan = Math.max(1, slot.cell.rowspan - 1);
          else if (slot.cell.rowspan > 1 && m.rows[r + 1]) {
            // the merged cell moves down into the next row
            const moved = { ...slot.cell, rowspan: slot.cell.rowspan - 1 };
            const below = grid[r + 1];
            const index = below.slice(0, c).filter((s, x) => s && s.row === r + 1 && s.col === x).length;
            m.rows[r + 1].cells.splice(index, 0, moved);
          }
        }
        m.rows.splice(r, 1);
      }
      return { r: Math.max(0, from - (from >= m.rows.length ? 1 : 0)), i: 0 };
    }, 'row');
  }

  deleteColumn() {
    this.change((m, { at, end, grid }) => {
      const from = Math.min(at.c, end.c), to = Math.max(at.c + m.rows[at.r].cells[at.i].colspan - 1, end.c);
      const n = grid[0]?.length || 1;
      if (from === 0 && to >= n - 1) { m.rows = []; return null; }
      const done = new Set();
      grid.forEach(slots => {
        for (let c = from; c <= to; c++) {
          const slot = slots[c];
          if (!slot || done.has(slot.cell)) continue;
          done.add(slot.cell);
          const covered = Math.min(to, slot.col + slot.cell.colspan - 1) - Math.max(from, slot.col) + 1;
          if (slot.cell.colspan > covered) slot.cell.colspan -= covered;
          else m.rows[slot.row].cells.splice(m.rows[slot.row].cells.indexOf(slot.cell), 1);
        }
      });
      m.rows = m.rows.filter(r => r.cells.length);
      const widths = colWidths(m, n);
      widths.splice(from, to - from + 1);
      const sum = widths.reduce((a, b) => a + b, 0) || 1;
      m.cols = widths.map(w => numStr(w / sum * 100) + '%');
      return { r: at.r, i: Math.max(0, at.i - 1) };
    }, 'column');
  }

  deleteTable() {
    const cur = this.current();
    if (!cur) return;
    this.ed.checkpoint('table');
    this.removeTable(cur.table);
  }

  // ---- merge and split ----------------------------------------------------------------------------------------------------------------

  canMerge() {
    const cur = this.current();
    return !!cur && cur.cell !== cur.endCell;
  }

  merge() {
    this.change((m, { at, end, grid }) => {
      const r0 = Math.min(at.r, end.r), c0 = Math.min(at.c, end.c);
      const r1 = Math.max(at.r, end.r), c1 = Math.max(at.c + m.rows[at.r].cells[at.i].colspan - 1, end.c + m.rows[end.r].cells[end.i].colspan - 1);
      const cells = new Set();
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (grid[r]?.[c]) cells.add(grid[r][c]);
      // the area must be a rectangle of whole cells
      let rr1 = r1, cc1 = c1;
      for (const s of cells) { rr1 = Math.max(rr1, s.row + s.cell.rowspan - 1); cc1 = Math.max(cc1, s.col + s.cell.colspan - 1); }
      const origin = grid[r0][c0];
      if (!origin || origin.row !== r0 || origin.col !== c0) return false;
      const blocks = [];
      for (const s of [...cells].sort((a, b) => a.row - b.row || a.col - b.col)) {
        const nonEmpty = s.cell.blocks.filter(b => !(b.runs && !b.runs.length));
        blocks.push(...nonEmpty);
        if (s.cell !== origin.cell) m.rows[s.row].cells.splice(m.rows[s.row].cells.indexOf(s.cell), 1);
      }
      origin.cell.blocks = blocks.length ? blocks : [textBlock()];
      origin.cell.colspan = cc1 - c0 + 1;
      origin.cell.rowspan = rr1 - r0 + 1;
      return { r: r0, i: m.rows[r0].cells.indexOf(origin.cell) };
    }, 'merge');
  }

  /** A merged cell falls apart into single cells; a single cell is split into two columns. */
  split() {
    this.change((m, { at, grid }) => {
      const cell = m.rows[at.r].cells[at.i];
      if (cell.colspan > 1 || cell.rowspan > 1) {
        const { colspan, rowspan } = cell;
        cell.colspan = 1;
        cell.rowspan = 1;
        for (let r = at.r; r < at.r + rowspan && r < m.rows.length; r++) {
          const index = r === at.r ? at.i + 1 : grid[r].slice(0, at.c).filter((s, x) => s && s.row === r && s.col === x).length;
          const count = r === at.r ? colspan - 1 : colspan;
          for (let k = 0; k < count; k++) m.rows[r].cells.splice(index, 0, emptyCell(m.rows[r].header));
        }
        return { r: at.r, i: at.i };
      }
      const n = grid[0]?.length || 1;
      const done = new Set();
      grid.forEach((slots, r) => {
        const slot = slots[at.c];
        if (!slot || done.has(slot.cell)) return;
        done.add(slot.cell);
        if (slot.cell === cell) m.rows[r].cells.splice(at.i + 1, 0, { ...emptyCell(m.rows[r].header), rowspan: cell.rowspan });
        else slot.cell.colspan += 1;
      });
      const widths = colWidths(m, n);
      const half = widths[at.c] / 2;
      widths.splice(at.c, 1, half, half);
      m.cols = widths.map(w => numStr(w) + '%');
      return { r: at.r, i: at.i + 1 };
    }, 'split');
  }

  // ---- look ---------------------------------------------------------------------------------------------------------------------------

  setStyle(style) {
    if (!TABLE_STYLES.includes(style)) return;
    this.change(m => { m.tableStyle = style; return null; }, 'style');
  }

  hasHeader() { return !!this.current()?.table.querySelector(':scope > thead'); }

  toggleHeader() {
    this.change(m => {
      const on = !m.rows[0]?.header;
      m.rows.forEach((r, i) => { r.header = on && i === 0; r.cells.forEach(c => { c.header = r.header; }); });
      return null;
    });
  }

  /** The cells of the selection (a rectangle when it spans cells). */
  cellsOf(m, { at, end, grid }) {
    const r0 = Math.min(at.r, end.r), r1 = Math.max(at.r, end.r), c0 = Math.min(at.c, end.c), c1 = Math.max(at.c, end.c);
    const out = new Set();
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (grid[r]?.[c]) out.add(grid[r][c].cell);
    return [...out];
  }

  shading(hex) {
    this.change((m, info) => { for (const c of this.cellsOf(m, info)) { if (hex) c.css['background-color'] = hex; else delete c.css['background-color']; } return { r: info.at.r, i: info.at.i }; });
  }

  borders(kind) {
    const line = '0.5pt solid #000000';
    this.change((m, info) => {
      const cells = kind === 'all' || kind === 'none' ? m.rows.flatMap(r => r.cells) : this.cellsOf(m, info);
      for (const c of cells) {
        for (const side of ['top', 'right', 'bottom', 'left']) {
          if (kind === 'all') c.css['border-' + side] = line;
          else if (kind === 'none') delete c.css['border-' + side];
          else if (side === kind) { if (c.css['border-' + side]) delete c.css['border-' + side]; else c.css['border-' + side] = line; }
        }
      }
      if (kind === 'none') m.tableStyle = 'plain';
      return { r: info.at.r, i: info.at.i };
    });
  }

  valign() { const cur = this.current(); return cur ? cur.cell.style.verticalAlign || 'top' : 'top'; }

  setValign(value) {
    this.change((m, info) => { for (const c of this.cellsOf(m, info)) { if (value === 'top') delete c.css['vertical-align']; else c.css['vertical-align'] = value; } return { r: info.at.r, i: info.at.i }; });
  }

  distribute() {
    this.change((m, { grid, at }) => { const n = grid[0]?.length || 1; m.cols = new Array(n).fill(numStr(100 / n) + '%'); return { r: at.r, i: at.i }; });
  }

  select(what) {
    const cur = this.current();
    if (!cur) return;
    const blocksOf = cells => cells.flatMap(c => [...c.querySelectorAll('p,h1,h2,h3,h4,h5,h6,blockquote,pre')].filter(b => b.closest('td,th') === c));
    let cells;
    if (what === 'cell') cells = [cur.cell];
    else if (what === 'row') cells = [...cur.row.children].filter(c => /^(TD|TH)$/.test(c.tagName));
    else if (what === 'column') {
      const m = readTable(cur.table);
      const grid = tableGrid(m);
      const at = this.locate(cur.table, cur.cell);
      const col = grid[at.r].findIndex(s => s && s.cell === m.rows[at.r].cells[at.i]);
      const els = rowsOf(cur.table).map(tr => [...tr.children].filter(c => /^(TD|TH)$/.test(c.tagName)));
      cells = [];
      grid.forEach((slots, r) => { const s = slots[col]; if (s && s.row === r) cells.push(els[r][m.rows[r].cells.indexOf(s.cell)]); });
      cells = cells.filter(Boolean);
    } else cells = [...cur.table.querySelectorAll('td,th')].filter(c => c.closest('table') === cur.table);
    const blocks = blocksOf(cells);
    if (blocks.length) this.ed.select(blocks[0], 0, blocks[blocks.length - 1], domLength(blocks[blocks.length - 1]), false);
  }

  // ---- Tab ---------------------------------------------------------------------------------------------------------------------------

  /** Tab / Shift+Tab inside a table: the next / previous cell (its text selected); Tab in the last cell adds a row. */
  tab(back) {
    const cur = this.current();
    if (!cur) return false;
    const cells = [...cur.table.querySelectorAll('td,th')].filter(c => c.closest('table') === cur.table);
    const i = cells.indexOf(cur.cell);
    const next = cells[i + (back ? -1 : 1)];
    if (!next) {
      if (back || this.app.ctx.readOnly) return true;
      this.insertRow(true);
      return true;
    }
    const blocks = [...next.querySelectorAll('p,h1,h2,h3,h4,h5,h6,blockquote,pre')].filter(b => b.closest('td,th') === next);
    if (blocks.length) this.ed.select(blocks[0], 0, blocks[blocks.length - 1], domLength(blocks[blocks.length - 1]));
    return true;
  }

  // ---- column widths by dragging -------------------------------------------------------------------------------------------------------

  setupResize() {
    const body = this.app.body;
    const handle = document.createElement('div');
    handle.className = 'mo-docs-colresize';
    handle.hidden = true;
    handle.setAttribute('data-mo-deco', '');
    document.body.append(handle);
    let target = null;
    body.addEventListener('pointermove', e => {
      if (this.app.ctx.readOnly || drag) return;
      const cell = e.target.closest?.('td,th');
      if (!cell || !body.contains(cell)) { if (!handle.matches(':hover')) handle.hidden = true; return; }
      const r = cell.getBoundingClientRect();
      const nearRight = Math.abs(e.clientX - r.right) < 5, nearLeft = Math.abs(e.clientX - r.left) < 5;
      const table = cell.closest('table');
      const cells = [...cell.parentNode.children];
      const index = cells.indexOf(cell);
      let border = null;
      if (nearRight && index < cells.length - 1) border = { table, cell, x: r.right };
      else if (nearLeft && index > 0) border = { table, cell: cells[index - 1], x: r.left };
      if (!border) { handle.hidden = true; return; }
      const tr = table.getBoundingClientRect();
      target = border;
      handle.hidden = false;
      handle.style.left = border.x + 'px';
      handle.style.top = tr.top + 'px';
      handle.style.height = tr.height + 'px';
    });
    let drag = null;
    handle.addEventListener('pointerdown', e => {
      if (!target) return;
      e.preventDefault();
      handle.setPointerCapture(e.pointerId);
      handle.classList.add('is-active');
      const table = target.table;
      const m = readTable(table);
      const grid = tableGrid(m);
      const rows = rowsOf(table);
      const r = rows.indexOf(target.cell.parentNode);
      const i = [...target.cell.parentNode.children].indexOf(target.cell);
      const slot = grid[r]?.find(s => s && s.cell === m.rows[r].cells[i]);
      if (!slot) return;
      const col = slot.col + slot.cell.colspan - 1;
      const n = grid[0].length;
      if (col >= n - 1) return;
      const widths = colWidths(m, n);
      drag = { table, col, widths, startX: e.clientX, width: table.getBoundingClientRect().width, checkpointed: false };
    });
    handle.addEventListener('pointermove', e => {
      if (!drag) return;
      const delta = (e.clientX - drag.startX) / drag.width * 100;
      const a = drag.widths[drag.col], b = drag.widths[drag.col + 1];
      const d = Math.max(-a + 3, Math.min(b - 3, delta));
      if (!drag.checkpointed && Math.abs(d) > 0.2) { this.ed.checkpoint('table'); drag.checkpointed = true; }
      const widths = drag.widths.slice();
      widths[drag.col] = a + d;
      widths[drag.col + 1] = b - d;
      applyWidths(drag.table, widths);
      handle.style.left = (drag.startX + d / 100 * drag.width) + 'px';
    });
    const end = () => { if (drag) { drag = null; handle.classList.remove('is-active'); handle.hidden = true; this.app.changed(); } };
    handle.addEventListener('pointerup', end);
    handle.addEventListener('pointercancel', end);
    this.app.view.scroller.addEventListener('scroll', () => { if (!drag) handle.hidden = true; }, { passive: true });
  }
}

// ---- helpers --------------------------------------------------------------------------------------------------------------------------

function rowsOf(table) {
  return [...table.querySelectorAll(':scope > thead > tr, :scope > tbody > tr, :scope > tr')];
}

function emptyCell(header = false) {
  return { header, colspan: 1, rowspan: 1, css: {}, blocks: [textBlock()] };
}

/** The widths of the grid columns in percent (missing ones shared equally). */
function colWidths(m, n) {
  const given = (m.cols || []).slice(0, n).map(c => (c ? parseFloat(c) : NaN));
  while (given.length < n) given.push(NaN);
  const known = given.filter(Number.isFinite);
  const rest = n - known.length;
  const left = Math.max(0, 100 - known.reduce((a, b) => a + b, 0));
  const each = rest ? Math.max(3, left / rest) : 0;
  const widths = given.map(w => (Number.isFinite(w) ? w : each));
  const sum = widths.reduce((a, b) => a + b, 0) || 1;
  return widths.map(w => w / sum * 100);
}

function normalizeCols(m) {
  const n = Math.max(1, ...tableGrid(m).map(r => r.length));
  m.cols = colWidths(m, n).map(w => numStr(w) + '%');
}

function applyWidths(table, widths) {
  let group = table.querySelector(':scope > colgroup');
  if (!group) { group = document.createElement('colgroup'); table.prepend(group); }
  while (group.children.length < widths.length) group.append(document.createElement('col'));
  [...group.children].forEach((col, i) => { if (widths[i] !== undefined) col.style.width = numStr(widths[i]) + '%'; });
}
