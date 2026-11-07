/* MatOffice Sheets – the grid: a canvas that draws only what is visible (any number of rows and columns), under a native scroll area
   (scroll bars, wheel and touch scrolling work like on every page). Four panes when rows/columns are frozen; headers; cell fills, gridlines,
   text with overflow into empty neighbours, wrapping, merged cells, borders, conditional formats, notes, filter buttons; the selection with its
   fill handle, the coloured ranges of a formula being edited, the moving border of a copy. Hit testing maps a point back to a cell. */
import { el } from '../editor/util.js';
import { colName, MAX_ROWS, MAX_COLS, areaContains } from './engine/address.js';
import { cssFont, parseBorder, isDark, DEFAULT_SIZE } from './look.js';

const HEADER_H = 20;
const PAD = 3;

export class Grid {
  constructor(app) {
    this.app = app;
    this.root = el('div', { class: 'mo-sheets-grid' });
    this.canvas = el('canvas', { class: 'mo-sheets-canvas', 'aria-hidden': 'true' });
    this.scroller = el('div', { class: 'mo-sheets-scroller', 'data-scroller': '' });
    this.spacer = el('div', { class: 'mo-sheets-spacer' });
    this.scroller.append(this.spacer);
    this.objects = el('div', { class: 'mo-sheets-objects' });
    this.objectsMain = el('div', { class: 'mo-sheets-objects__layer' });
    this.objects.append(this.objectsMain);
    this.overlay = el('div', { class: 'mo-sheets-overlay' });
    this.root.append(this.canvas, this.scroller, this.objects, this.overlay);
    this.g = this.canvas.getContext('2d');
    this.width = 0;
    this.height = 0;
    this.extentRows = 200;
    this.extentCols = 40;
    this.widths = new Map();
    this.highlights = [];
    this.marquee = null;
    this.fillPreview = null;
    this.movePreview = null;
    this.dashOffset = 0;
    this.pending = false;
    this.themeCache = null;
    this.scroller.addEventListener('scroll', () => { this.grow(); this.requestRender(); }, { passive: true });
    this.resizeObserver = new ResizeObserver(() => this.resize());
    this.resizeObserver.observe(this.root);
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    mq.addEventListener?.('change', () => { this.themeCache = null; this.requestRender(); });
    new MutationObserver(() => { this.themeCache = null; this.requestRender(); }).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme', 'data-scheme'] });
  }

  get sheet() { return this.override?.sheet || this.app.sheet; }

  resize() {
    const r = this.root.getBoundingClientRect();
    this.width = Math.max(0, Math.floor(r.width));
    this.height = Math.max(0, Math.floor(r.height));
    const dpr = window.devicePixelRatio || 1;
    this.dpr = dpr;
    this.canvas.width = Math.round(this.width * dpr);
    this.canvas.height = Math.round(this.height * dpr);
    this.canvas.style.width = this.width + 'px';
    this.canvas.style.height = this.height + 'px';
    this.updateExtent();
    if (this.pendingScroll && this.width && this.height) {
      const { row, col } = this.pendingScroll;
      this.pendingScroll = null;
      this.scrollIntoView(row, col);
    }
    this.render();
  }

  requestRender() {
    if (this.pending) return;
    this.pending = true;
    requestAnimationFrame(() => this.render());
  }

  // ---- theme -----------------------------------------------------------------------------------------------------------------

  theme() {
    if (this.themeCache) return this.themeCache;
    const cs = getComputedStyle(this.root);
    const v = name => cs.getPropertyValue(name).trim();
    const bg = v('--mat-surface') || '#ffffff';
    const dark = isDarkCss(bg);
    this.themeCache = dark ? {
      dark, cell: '#1e1e1e', text: '#e8e8e8', grid: '#3a3a3a', header: '#2a2a2a', headerText: '#c8c8c8', headerLine: '#454545',
      headerSel: '#1f4d34', headerSelText: '#ffffff', headerAll: '#2f6b48', accent: '#21a366', selFill: 'rgba(33, 163, 102, 0.18)', freeze: '#777777', outside: '#1e1e1e',
    } : {
      dark, cell: '#ffffff', text: '#000000', grid: '#e1e1e1', header: '#f5f5f5', headerText: '#444444', headerLine: '#d4d4d4',
      headerSel: '#d3eadb', headerSelText: '#0b5a2f', headerAll: '#9fd5b7', accent: '#107c41', selFill: 'rgba(16, 124, 65, 0.11)', freeze: '#9a9a9a', outside: '#ffffff',
    };
    return this.themeCache;
  }

  // ---- geometry ----------------------------------------------------------------------------------------------------------------

  rowHeaderWidth(sheet) {
    const top = this.override ? 0 : sheet.rows.indexAt(this.scroller.scrollTop / sheet.zoom + sheet.rows.offset(sheet.freeze.rows));
    const visible = Math.ceil(this.height / Math.max(1, sheet.rows.def * sheet.zoom));
    const digits = String(Math.max(top + visible + 1, sheet.freeze.rows + 1)).length;
    return Math.max(28, 9 + digits * 8);
  }

  /** The measures of the screen: zoom, header sizes, frozen sizes, scroll offsets (content pixels) and the panes. */
  metrics() {
    const s = this.sheet;
    const z = this.override ? this.override.zoom : s.zoom;
    const hw = s.headings ? Math.round(this.rowHeaderWidth(s) * z) : 0;
    const hh = s.headings ? Math.round(HEADER_H * z) : 0;
    const fr = Math.min(s.freeze.rows, MAX_ROWS - 1);
    const fc = Math.min(s.freeze.cols, MAX_COLS - 1);
    const fH = s.rows.offset(fr);
    const fW = s.cols.offset(fc);
    const sx = this.override ? 0 : this.scroller.scrollLeft / z;
    const sy = this.override ? 0 : this.scroller.scrollTop / z;
    const mainX = hw + fW * z;
    const mainY = hh + fH * z;
    const panes = [];
    const mk = (x0, y0, w, h, cx0, cy0, colMin, colMax, rowMin, rowMax) => ({ x0, y0, w, h, cx0, cy0, colMin, colMax, rowMin, rowMax });
    if (fr && fc) panes.push(mk(hw, hh, fW * z, fH * z, 0, 0, 0, fc - 1, 0, fr - 1));
    if (fr) panes.push(mk(mainX, hh, Math.max(0, this.width - mainX), fH * z, fW + sx, 0, fc, MAX_COLS - 1, 0, fr - 1));
    if (fc) panes.push(mk(hw, mainY, fW * z, Math.max(0, this.height - mainY), 0, fH + sy, 0, fc - 1, fr, MAX_ROWS - 1));
    panes.push(mk(mainX, mainY, Math.max(0, this.width - mainX), Math.max(0, this.height - mainY), fW + sx, fH + sy, fc, MAX_COLS - 1, fr, MAX_ROWS - 1));
    for (const p of panes) {
      p.c1 = Math.max(p.colMin, s.cols.indexAt(p.cx0));
      p.c2 = Math.min(p.colMax, s.cols.indexAt(p.cx0 + p.w / z));
      p.r1 = Math.max(p.rowMin, s.rows.indexAt(p.cy0));
      p.r2 = Math.min(p.rowMax, s.rows.indexAt(p.cy0 + p.h / z));
    }
    return { s, z, hw, hh, fr, fc, fW, fH, sx, sy, mainX, mainY, panes };
  }

  /** The pane that shows a cell. */
  paneOf(m, row, col) {
    const frozenRow = row < m.fr;
    const frozenCol = col < m.fc;
    return m.panes.find(p => (frozenRow ? p.rowMax < m.fr : p.rowMin >= m.fr) && (frozenCol ? p.colMax < m.fc : p.colMin >= m.fc)) || m.panes[m.panes.length - 1];
  }

  /** The rectangle of a cell (or of the merged area it starts) on the screen: { x, y, w, h } in CSS pixels relative to the grid. */
  cellRect(row, col, m = this.metrics()) {
    const s = m.s;
    const p = this.paneOf(m, row, col);
    const merge = s.mergeAt(row, col);
    const r1 = merge ? merge.r1 : row;
    const c1 = merge ? merge.c1 : col;
    const r2 = merge ? merge.r2 : row;
    const c2 = merge ? merge.c2 : col;
    const x = p.x0 + (s.cols.offset(c1) - p.cx0) * m.z;
    const y = p.y0 + (s.rows.offset(r1) - p.cy0) * m.z;
    const w = (s.cols.offset(c2 + 1) - s.cols.offset(c1)) * m.z;
    const h = (s.rows.offset(r2 + 1) - s.rows.offset(r1)) * m.z;
    return { x, y, w, h, pane: p };
  }

  /** The screen rectangle of an area (in the pane of its first cell). */
  areaRect(a, m = this.metrics()) {
    const s = m.s;
    const p = this.paneOf(m, a.r1, a.c1);
    const x = p.x0 + (s.cols.offset(a.c1) - p.cx0) * m.z;
    const y = p.y0 + (s.rows.offset(a.r1) - p.cy0) * m.z;
    const x2 = this.paneOf(m, a.r1, a.c2).x0 + (s.cols.offset(Math.min(a.c2, MAX_COLS - 1) + 1) - this.paneOf(m, a.r1, a.c2).cx0) * m.z;
    const y2 = this.paneOf(m, a.r2, a.c1).y0 + (s.rows.offset(Math.min(a.r2, MAX_ROWS - 1) + 1) - this.paneOf(m, a.r2, a.c1).cy0) * m.z;
    return { x, y, w: x2 - x, h: y2 - y };
  }

  /**
   * What is at a point: { zone: 'corner'|'col'|'row'|'cell'|'none', row, col, resize: { axis, index } | null, fill: bool, filter: col | null }.
   */
  hitTest(x, y, m = this.metrics()) {
    const s = m.s;
    const res = { zone: 'cell', row: 0, col: 0, resize: null, fill: false, filter: null };
    const colAt = px => (px < m.mainX ? s.cols.indexAt((px - m.hw) / m.z) : s.cols.indexAt((px - m.hw) / m.z + m.sx));
    const rowAt = py => (py < m.mainY ? s.rows.indexAt((py - m.hh) / m.z) : s.rows.indexAt((py - m.hh) / m.z + m.sy));
    const colX = c => (c < m.fc ? m.hw + s.cols.offset(c) * m.z : m.hw + (s.cols.offset(c) - m.sx) * m.z);
    const rowY = r => (r < m.fr ? m.hh + s.rows.offset(r) * m.z : m.hh + (s.rows.offset(r) - m.sy) * m.z);
    if (x < m.hw && y < m.hh) { res.zone = 'corner'; return res; }
    if (y < m.hh) {
      res.zone = 'col';
      res.col = colAt(Math.max(m.hw, x));
      const left = colX(res.col);
      const right = left + s.cols.size(res.col) * m.z;
      if (x > right - 4) res.resize = { axis: 'col', index: res.col };
      else if (x < left + 4 && res.col > 0) res.resize = { axis: 'col', index: this.prevVisible(s.cols, res.col) };
      return res;
    }
    if (x < m.hw) {
      res.zone = 'row';
      res.row = rowAt(Math.max(m.hh, y));
      const top = rowY(res.row);
      const bottom = top + s.rows.size(res.row) * m.z;
      if (y > bottom - 3) res.resize = { axis: 'row', index: res.row };
      else if (y < top + 3 && res.row > 0) res.resize = { axis: 'row', index: this.prevVisible(s.rows, res.row) };
      return res;
    }
    res.row = rowAt(y);
    res.col = colAt(x);
    // the fill handle
    const a = this.app.sel.range;
    if (!this.app.readOnly && !(a.r2 === MAX_ROWS - 1 && a.c2 === MAX_COLS - 1)) {
      const r = this.areaRect(a, m);
      if (Math.abs(x - (r.x + r.w - 1)) <= 5 && Math.abs(y - (r.y + r.h - 1)) <= 5) res.fill = true;
      else if (this.onBorder(x, y, r)) res.border = true;
    }
    // a filter button
    const fa = s.filterArea();
    if (fa && res.row === fa.r1 && res.col >= fa.c1 && res.col <= fa.c2) {
      const cr = this.cellRect(res.row, res.col, m);
      const b = this.filterButtonRect(cr, m.z);
      if (x >= b.x && x <= b.x + b.w && y >= b.y && y <= b.y + b.h) res.filter = res.col;
    }
    return res;
  }

  onBorder(x, y, r) {
    const near = (v, edge) => Math.abs(v - edge) <= 2;
    const insideX = x >= r.x - 3 && x <= r.x + r.w + 3;
    const insideY = y >= r.y - 3 && y <= r.y + r.h + 3;
    return (insideY && (near(x, r.x) || near(x, r.x + r.w - 1))) || (insideX && (near(y, r.y) || near(y, r.y + r.h - 1)));
  }

  prevVisible(axis, i) {
    let j = i - 1;
    while (j > 0 && axis.isHidden(j)) j--;
    return Math.max(0, j);
  }

  filterButtonRect(cr, z) {
    const size = Math.min(cr.h - 2, Math.round(16 * z));
    return { x: cr.x + cr.w - size - 2, y: cr.y + cr.h - size - 2, w: size, h: size };
  }

  // ---- scrolling ------------------------------------------------------------------------------------------------------------------

  /** The size of the scroll area: the used part of the sheet and a margin that grows while scrolling (like Excel's scroll bars). */
  updateExtent() {
    const s = this.sheet;
    if (!s) return;
    const b = s.bounds();
    const e = this.app.calc.extent(s.id);
    const sel = this.app.sel;
    const z = s.zoom;
    const visRows = Math.ceil(this.height / Math.max(4, s.rows.def * z)) + 2;
    const visCols = Math.ceil(this.width / Math.max(4, s.cols.def * z)) + 2;
    const objects = [...s.charts, ...s.images];
    let objRow = 0;
    let objCol = 0;
    for (const o of objects) { objRow = Math.max(objRow, s.rows.indexAt(o.y + o.h)); objCol = Math.max(objCol, s.cols.indexAt(o.x + o.w)); }
    this.extentRows = Math.min(MAX_ROWS, Math.max(this.extentRows, b.rows + visRows, e.rows + visRows, sel.active.row + visRows, objRow + visRows, 100));
    this.extentCols = Math.min(MAX_COLS, Math.max(this.extentCols, b.cols + visCols, e.cols + visCols, sel.active.col + visCols, objCol + visCols, 26));
    const m = this.metrics();
    const w = m.hw + s.cols.offset(this.extentCols) * z;
    const h = m.hh + s.rows.offset(this.extentRows) * z;
    this.spacer.style.width = Math.min(w, 15e6) + 'px';
    this.spacer.style.height = Math.min(h, 15e6) + 'px';
  }

  /** Near the end of the scroll area it grows. */
  grow() {
    const sc = this.scroller;
    let changed = false;
    if (sc.scrollTop + sc.clientHeight > sc.scrollHeight - 300 && this.extentRows < MAX_ROWS) { this.extentRows = Math.min(MAX_ROWS, this.extentRows + 200); changed = true; }
    if (sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 300 && this.extentCols < MAX_COLS) { this.extentCols = Math.min(MAX_COLS, this.extentCols + 20); changed = true; }
    if (changed) this.updateExtent();
  }

  /** Scrolls so that a cell is visible (cells in frozen panes are always visible in their direction). */
  scrollIntoView(row, col) {
    const s = this.sheet;
    if (!this.width || !this.height) { this.pendingScroll = { row, col }; return; }
    const m = this.metrics();
    const z = m.z;
    const merge = s.mergeAt(row, col);
    const r2 = merge ? merge.r2 : row;
    const c2 = merge ? merge.c2 : col;
    if (row >= m.fr && row >= this.extentRows - 5) { this.extentRows = Math.min(MAX_ROWS, row + 100); this.updateExtent(); }
    if (col >= m.fc && col >= this.extentCols - 3) { this.extentCols = Math.min(MAX_COLS, col + 20); this.updateExtent(); }
    const sc = this.scroller;
    if (row >= m.fr) {
      const top = s.rows.offset(row) - m.fH;
      const bottom = s.rows.offset(r2 + 1) - m.fH;
      const viewH = (this.height - m.mainY) / z;
      if (top < m.sy) sc.scrollTop = top * z;
      else if (bottom > m.sy + viewH) sc.scrollTop = Math.max(0, (bottom - viewH) * z + 2);
    }
    if (col >= m.fc) {
      const left = s.cols.offset(col) - m.fW;
      const right = s.cols.offset(c2 + 1) - m.fW;
      const viewW = (this.width - m.mainX) / z;
      if (left < m.sx) sc.scrollLeft = left * z;
      else if (right > m.sx + viewW) sc.scrollLeft = Math.max(0, Math.min(left, right - viewW) * z + 2);
    }
    this.requestRender();
  }

  /** The number of rows that fit into the main pane (for PageUp/PageDown). */
  pageRows() {
    const m = this.metrics();
    return Math.max(1, Math.floor((this.height - m.mainY) / (m.s.rows.def * m.z)) - 1);
  }

  // ---- drawing ---------------------------------------------------------------------------------------------------------------------

  render() {
    if (!this.override) this.pending = false;
    if (!this.width || !this.height || !this.sheet) return;
    const g = this.g;
    const t = this.theme();
    const m = this.metrics();
    const dpr = this.dpr || 1;
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.fillStyle = t.cell;
    g.fillRect(0, 0, this.width, this.height);
    for (const p of m.panes) if (p.w > 0 && p.h > 0) this.drawPane(g, m, p, t);
    // the lines of the frozen panes
    if (m.fr || m.fc) {
      g.fillStyle = t.freeze;
      if (m.fr) g.fillRect(m.hw, Math.round(m.mainY) - 1, this.width - m.hw, 1);
      if (m.fc) g.fillRect(Math.round(m.mainX) - 1, m.hh, 1, this.height - m.hh);
    }
    if (m.s.headings) this.drawHeaders(g, m, t);
    if (this.override) return;
    this.app.objects?.position(m);
    this.app.editor?.position(m);
    this.app.onRender?.(m);
  }

  /** Draws a sheet from A1 into another canvas context (the picture of the gallery), without selection. */
  renderTo(g, width, height, sheet, zoom = 1) {
    const saved = { g: this.g, width: this.width, height: this.height, dpr: this.dpr, theme: this.themeCache };
    this.override = { sheet, zoom };
    this.themeCache = null;
    Object.assign(this, { g, width, height, dpr: 1 });
    const dark = this.theme;
    this.theme = () => ({ ...LIGHT });
    try { this.render(); } finally {
      this.theme = dark;
      this.override = null;
      Object.assign(this, saved);
      this.themeCache = saved.theme;
    }
  }

  drawPane(g, m, p, t) {
    const s = m.s;
    const z = m.z;
    const app = this.app;
    g.save();
    g.beginPath();
    g.rect(p.x0, p.y0, p.w, p.h);
    g.clip();
    const cols = [];
    for (let c = p.c1; c <= p.c2; c++) {
      const w = s.cols.size(c) * z;
      if (w > 0) cols.push({ c, x: p.x0 + (s.cols.offset(c) - p.cx0) * z, w });
    }
    const rows = [];
    for (let r = p.r1; r <= p.r2; r++) {
      const h = s.rows.size(r) * z;
      if (h > 0) rows.push({ r, y: p.y0 + (s.rows.offset(r) - p.cy0) * z, h });
    }
    const merges = s.merges.filter(a => a.r2 >= p.r1 && a.r1 <= p.r2 && a.c2 >= p.c1 && a.c1 <= p.c2);
    const inMerge = (r, c) => { for (const a of merges) if (r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2) return a; return null; };
    const showGrid = s.grid;
    const px = 1 / (this.dpr || 1);

    // the looks of the visible cells, and the text that overflows into empty neighbours
    const lines = [];
    for (const row of rows) {
      const cells = [];
      for (const col of cols) {
        const merge = merges.length ? inMerge(row.r, col.c) : null;
        if (merge) { cells.push({ col, merge, look: null }); continue; }
        const look = s.cells.has(row.r * MAX_COLS + col.c) || s.rowStyles.size || s.colStyles.size || s.cf.length || app.calc.formulaAt(s.id, row.r, col.c)
          ? app.look.cell(s, row.r, col.c, this.chars(col.w / z)) : null;
        cells.push({ col, merge: null, look });
      }
      lines.push({ row, cells, spans: this.overflow(s, m, row, cols, cells) });
    }

    // gridlines
    if (showGrid) {
      g.fillStyle = t.grid;
      for (const col of cols) g.fillRect(snap(col.x + col.w, this.dpr) - px, p.y0, px, p.h);
      for (const row of rows) g.fillRect(p.x0, snap(row.y + row.h, this.dpr) - px, p.w, px);
    }

    // fills, text overflow areas, data bars
    for (const line of lines) {
      const { row } = line;
      for (const cell of line.cells) {
        const look = cell.look;
        if (!look) continue;
        if (look.fill) { g.fillStyle = look.fill; g.fillRect(cell.col.x, row.y, cell.col.w, row.h); }
        if (look.bar) this.drawBar(g, cell.col.x, row.y, cell.col.w, row.h, look.bar, z);
      }
      for (const span of line.spans) {
        // hide the gridlines under text that runs into the neighbours
        for (const cell of line.cells) {
          if (cell.col.x >= span.x1 - 0.5 && cell.col.x + cell.col.w <= span.x2 + 0.5 && cell.col.c !== span.from && !(cell.look?.fill)) {
            g.fillStyle = t.cell;
            g.fillRect(cell.col.x - (cell.col.c > span.firstCol ? px : 0), row.y, cell.col.w + (cell.col.c > span.firstCol ? px : 0) - (cell.col.c === span.lastCol ? px : 0), row.h - px);
          }
        }
        const own = line.cells.find(x => x.col.c === span.from);
        if (own && !own.look?.fill && span.x2 > own.col.x + own.col.w) { g.fillStyle = t.cell; g.fillRect(own.col.x + own.col.w - px, row.y, px, row.h - px); }
      }
    }

    // merged areas: one cell
    for (const a of merges) {
      const x = p.x0 + (s.cols.offset(a.c1) - p.cx0) * z;
      const y = p.y0 + (s.rows.offset(a.r1) - p.cy0) * z;
      const w = (s.cols.offset(a.c2 + 1) - s.cols.offset(a.c1)) * z;
      const h = (s.rows.offset(a.r2 + 1) - s.rows.offset(a.r1)) * z;
      if (w <= 0 || h <= 0) continue;
      const look = app.look.cell(s, a.r1, a.c1, this.chars(w / z));
      g.fillStyle = look.fill || t.cell;
      g.fillRect(x, y, w - px, h - px);
      if (look.bar) this.drawBar(g, x, y, w, h, look.bar, z);
      this.drawText(g, look, x, y, w, h, z, t, null);
      a.drawn = { x, y, w, h, look };
    }

    // text
    for (const line of lines) {
      for (const cell of line.cells) {
        if (!cell.look || !cell.look.text) continue;
        const span = line.spans.find(sp => sp.from === cell.col.c);
        this.drawText(g, cell.look, cell.col.x, line.row.y, cell.col.w, line.row.h, z, t, span);
      }
    }
    // text of cells left of the pane that runs into it
    for (const line of lines) {
      for (const span of line.spans) {
        if (span.outside) this.drawText(g, span.look, span.cellX, line.row.y, span.cellW, line.row.h, z, t, span);
      }
    }

    // borders
    for (const line of lines) {
      for (const cell of line.cells) {
        const style = cell.look?.style;
        if (!style || !(style.bt || style.bb || style.bl || style.br)) continue;
        this.drawBorders(g, style, cell.col.x, line.row.y, cell.col.w, line.row.h);
      }
    }
    for (const a of merges) {
      if (!a.drawn) continue;
      const { x, y, w, h, look } = a.drawn;
      const bottomRight = app.wb.style(s.styleAt(a.r2, a.c2));
      this.drawBorders(g, { bt: look.style.bt, bl: look.style.bl, bb: bottomRight.bb || look.style.bb, br: bottomRight.br || look.style.br }, x, y, w, h);
      delete a.drawn;
    }

    // notes, filter buttons
    const fa = s.filterArea();
    for (const line of lines) {
      for (const cell of line.cells) {
        const note = s.cell(line.row.r, cell.col.c)?.note;
        if (note) {
          const merge = cell.merge;
          const right = merge ? p.x0 + (s.cols.offset(merge.c2 + 1) - p.cx0) * z : cell.col.x + cell.col.w;
          const top = merge ? p.y0 + (s.rows.offset(merge.r1) - p.cy0) * z : line.row.y;
          if (!merge || (merge.r1 === line.row.r && merge.c2 === cell.col.c)) {
            const k = Math.max(5, 6 * z);
            g.fillStyle = '#d13438';
            g.beginPath();
            g.moveTo(right - 1 - k, top);
            g.lineTo(right - 1, top);
            g.lineTo(right - 1, top + k);
            g.closePath();
            g.fill();
          }
        }
        if (fa && line.row.r === fa.r1 && cell.col.c >= fa.c1 && cell.col.c <= fa.c2) {
          const b = this.filterButtonRect({ x: cell.col.x, y: line.row.y, w: cell.col.w, h: line.row.h }, z);
          const active = !!s.filter.cols?.[colName(cell.col.c)];
          this.drawFilterButton(g, b, active, t);
        }
      }
    }

    this.drawSelection(g, m, p, t);
    g.restore();
  }

  chars(widthPx) { return Math.max(1, Math.min(11, Math.floor((widthPx - 5) / 7))); }

  measure(font, text) {
    const k = font + '\u0000' + text;
    let w = this.widths.get(k);
    if (w === undefined) {
      this.g.font = font;
      w = this.g.measureText(text).width;
      if (this.widths.size > 20000) this.widths.clear();
      this.widths.set(k, w);
    }
    return w;
  }

  /** Text that does not fit and may run into empty cells next to it: spans { from, x1, x2 }. */
  overflow(s, m, row, cols, cells) {
    const spans = [];
    const z = m.z;
    const empty = c => {
      const cell = s.cell(row.r, c);
      if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== '' || cell.f)) return false;
      if (this.app.calc.formulaAt(s.id, row.r, c)) return false;
      return !s.mergeAt(row.r, c);
    };
    const candidate = look => look && look.text && !look.wrap && !look.isNumber && look.ha !== 'justify' && look.ha !== 'fill' && look.fillAt < 0;
    const fits = (look, w) => this.measure(cssFont({ ...look.style, b: look.b, i: look.i }, z), look.text) + (PAD * 2 + look.indent * 9) * z <= w;
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      const look = cell.look;
      if (!candidate(look) || fits(look, cell.col.w)) continue;
      const need = this.measure(cssFont({ ...look.style, b: look.b, i: look.i }, z), look.text) + (PAD * 2 + look.indent * 9) * z;
      let x1 = cell.col.x;
      let x2 = cell.col.x + cell.col.w;
      let firstCol = cell.col.c;
      let lastCol = cell.col.c;
      if (look.ha === 'left' || look.ha === 'center') {
        for (let j = i + 1; j < cells.length && x2 - x1 < need; j++) {
          if (!empty(cells[j].col.c)) break;
          x2 = cells[j].col.x + cells[j].col.w;
          lastCol = cells[j].col.c;
        }
      }
      if (look.ha === 'right' || look.ha === 'center') {
        for (let j = i - 1; j >= 0 && x2 - x1 < need; j--) {
          if (!empty(cells[j].col.c)) break;
          x1 = cells[j].col.x;
          firstCol = cells[j].col.c;
        }
      }
      if (x1 !== cell.col.x || x2 !== cell.col.x + cell.col.w) spans.push({ from: cell.col.c, x1, x2, firstCol, lastCol });
    }
    // text of a cell left of the first visible column
    if (cols.length && cells.length && cells[0].look === null && empty(cols[0].c)) {
      for (let c = cols[0].c - 1, n = 0; c >= 0 && n < 12; c--, n++) {
        if (s.cols.isHidden(c)) continue;
        if (!empty(c)) {
          const look = this.app.look.cell(s, row.r, c, this.chars(s.cols.size(c)));
          if (!candidate(look) || look.ha !== 'left') break;
          const cellX = cols[0].x - (s.cols.offset(cols[0].c) - s.cols.offset(c)) * z;
          const cellW = s.cols.size(c) * z;
          const need = this.measure(cssFont({ ...look.style, b: look.b, i: look.i }, z), look.text) + (PAD * 2 + look.indent * 9) * z;
          if (cellW >= need) break;
          let x2 = cellX + cellW;
          let lastCol = c;
          for (const cell of cells) {
            if (x2 - cellX >= need) break;
            if (!empty(cell.col.c)) break;
            x2 = cell.col.x + cell.col.w;
            lastCol = cell.col.c;
          }
          if (x2 > cols[0].x) spans.push({ from: c, x1: cellX, x2, outside: true, look, cellX, cellW, firstCol: c, lastCol });
          break;
        }
      }
    }
    return spans;
  }

  drawText(g, look, x, y, w, h, z, t, span) {
    const text = look.text;
    if (!text) return;
    const style = look.style;
    const font = cssFont({ ...style, b: look.b, i: look.i }, z);
    const fontPx = (style.size || DEFAULT_SIZE) * 96 / 72 * z;
    let color = look.color;
    if (!color) color = look.fill ? (isDark(look.fill) ? '#ffffff' : '#000000') : t.text;
    g.font = font;
    g.fillStyle = color;
    const pad = PAD * z;
    const indent = look.indent * 9 * z;
    const clipX1 = span ? span.x1 : x;
    const clipX2 = span ? span.x2 : x + w;
    g.save();
    g.beginPath();
    g.rect(clipX1, y, clipX2 - clipX1 - 1, h - 1);
    g.clip();
    if (look.wrap) {
      const lines = this.wrapLines(text, font, Math.max(4, w - pad * 2 - indent));
      const lineH = fontPx * 1.22;
      const total = lines.length * lineH;
      let top = look.va === 'top' ? y + 2 * z : look.va === 'middle' ? y + (h - total) / 2 : y + h - total - 2 * z;
      g.textBaseline = 'top';
      for (const line of lines) {
        const lw = this.measure(font, line);
        const lx = look.ha === 'right' ? x + w - pad - lw - indent : look.ha === 'center' ? x + (w - lw) / 2 : x + pad + indent;
        g.fillText(line, lx, top + fontPx * 0.08);
        this.decorate(g, look, lx, top + fontPx * 0.08, lw, fontPx, 'top');
        top += lineH;
      }
      g.restore();
      return;
    }
    let display = text;
    let tw = this.measure(font, display);
    if (look.isNumber && tw + pad * 2 > w && !span) {
      // a number that does not fit: ### like Excel
      const hash = this.measure(font, '#');
      display = '#'.repeat(Math.max(1, Math.floor((w - pad * 2) / Math.max(1, hash))));
      tw = this.measure(font, display);
    }
    let baseY;
    if (look.va === 'top') { g.textBaseline = 'top'; baseY = y + 2 * z; } else if (look.va === 'middle') { g.textBaseline = 'middle'; baseY = y + h / 2 + 0.5; } else { g.textBaseline = 'alphabetic'; baseY = y + h - Math.max(3, fontPx * 0.24 + 1); }
    if (look.fillAt >= 0 && display === text) {
      // accounting: the part before the fill on the left, the rest on the right
      const left = text.slice(0, look.fillAt);
      const right = text.slice(look.fillAt);
      g.fillText(left, x + pad, baseY);
      const rw = this.measure(font, right);
      g.fillText(right, x + w - pad - rw, baseY);
      g.restore();
      return;
    }
    let tx;
    const ha = look.ha === 'justify' || look.ha === 'fill' ? 'left' : look.ha;
    if (span) {
      tx = ha === 'right' ? x + w - pad - tw - indent : ha === 'center' ? x + (w - tw) / 2 : x + pad + indent;
    } else {
      tx = ha === 'right' ? x + w - pad - tw - indent : ha === 'center' ? x + (w - tw) / 2 : x + pad + indent;
    }
    g.fillText(display, tx, baseY);
    this.decorate(g, look, tx, baseY, tw, fontPx, g.textBaseline);
    g.restore();
  }

  decorate(g, look, x, y, w, fontPx, baseline) {
    if (!look.u && !look.strike) return;
    const thick = Math.max(1, fontPx / 14);
    const base = baseline === 'top' ? y + fontPx * 0.92 : baseline === 'middle' ? y + fontPx * 0.36 : y;
    if (look.u) g.fillRect(x, Math.round(base + thick + 0.5), w, thick);
    if (look.strike) g.fillRect(x, Math.round(base - fontPx * 0.3), w, thick);
  }

  wrapLines(text, font, width) {
    const out = [];
    for (const para of String(text).split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const word of words) {
        const next = line + word;
        if (line && this.measure(font, next.trimEnd()) > width) {
          out.push(line.trimEnd());
          line = word.trimStart();
          // a single word longer than the line is cut into pieces
          while (line && this.measure(font, line) > width && line.length > 1) {
            let cut = line.length - 1;
            while (cut > 1 && this.measure(font, line.slice(0, cut)) > width) cut--;
            out.push(line.slice(0, cut));
            line = line.slice(cut);
          }
        } else {
          line = next;
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  drawBar(g, x, y, w, h, bar, z) {
    const pad = 2 * z;
    const bw = Math.max(0, (w - pad * 2) * bar.ratio);
    const grad = g.createLinearGradient(x + pad, 0, x + pad + Math.max(1, bw), 0);
    grad.addColorStop(0, bar.color);
    grad.addColorStop(1, bar.color + '33');
    g.fillStyle = grad;
    g.fillRect(x + pad, y + pad, bw, h - pad * 2 - 1);
    g.strokeStyle = bar.color;
    g.lineWidth = 1;
    g.strokeRect(x + pad + 0.5, y + pad + 0.5, Math.max(0, bw - 1), h - pad * 2 - 2);
  }

  drawBorders(g, style, x, y, w, h) {
    const edge = (spec, x1, y1, x2, y2) => {
      const b = parseBorder(spec);
      if (!b) return;
      g.strokeStyle = b.color;
      g.lineWidth = b.style === 'double' ? 1 : b.width;
      g.setLineDash(b.style === 'dashed' ? [4, 2] : b.style === 'dotted' ? [1, 2] : b.style === 'hair' ? [1, 1] : []);
      const off = b.lineWidth % 2 ? 0.5 : 0;
      g.beginPath();
      if (b.style === 'double') {
        const vertical = x1 === x2;
        for (const d of [-1, 1]) {
          g.moveTo(x1 + (vertical ? d : 0) + 0.5, y1 + (vertical ? 0 : d) + 0.5);
          g.lineTo(x2 + (vertical ? d : 0) + 0.5, y2 + (vertical ? 0 : d) + 0.5);
        }
      } else {
        g.moveTo(Math.round(x1) + (b.width % 2 ? 0.5 : 0) + off * 0, Math.round(y1) + (b.width % 2 ? 0.5 : 0));
        g.lineTo(Math.round(x2) + (b.width % 2 ? 0.5 : 0), Math.round(y2) + (b.width % 2 ? 0.5 : 0));
      }
      g.stroke();
      g.setLineDash([]);
    };
    const right = x + w - 1;
    const bottom = y + h - 1;
    if (style.bt) edge(style.bt, x - 1, y - 1, right, y - 1);
    if (style.bb) edge(style.bb, x - 1, bottom, right, bottom);
    if (style.bl) edge(style.bl, x - 1, y - 1, x - 1, bottom);
    if (style.br) edge(style.br, right, y - 1, right, bottom);
  }

  drawFilterButton(g, b, active, t) {
    g.fillStyle = t.dark ? '#3a3a3a' : '#ffffff';
    g.strokeStyle = t.dark ? '#666666' : '#b0b0b0';
    g.lineWidth = 1;
    g.fillRect(b.x, b.y, b.w, b.h);
    g.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1);
    g.fillStyle = t.dark ? '#dddddd' : '#444444';
    const cx = b.x + b.w / 2;
    const cy = b.y + b.h / 2;
    const k = b.w / 5;
    g.beginPath();
    if (active) {
      // a funnel
      g.moveTo(cx - k * 1.4, cy - k * 1.2);
      g.lineTo(cx + k * 1.4, cy - k * 1.2);
      g.lineTo(cx + k * 0.3, cy);
      g.lineTo(cx + k * 0.3, cy + k * 1.4);
      g.lineTo(cx - k * 0.3, cy + k * 1.1);
      g.lineTo(cx - k * 0.3, cy);
    } else {
      g.moveTo(cx - k, cy - k * 0.5);
      g.lineTo(cx + k, cy - k * 0.5);
      g.lineTo(cx, cy + k * 0.6);
    }
    g.closePath();
    g.fill();
  }

  drawSelection(g, m, p, t) {
    if (this.override) return;
    const app = this.app;
    const s = m.s;
    const z = m.z;
    const rectOf = a => {
      const r1 = Math.max(a.r1, p.rowMin);
      const r2 = Math.min(a.r2, p.rowMax);
      const c1 = Math.max(a.c1, p.colMin);
      const c2 = Math.min(a.c2, p.colMax);
      if (r1 > r2 || c1 > c2) return null;
      const x = p.x0 + (s.cols.offset(c1) - p.cx0) * z;
      const y = p.y0 + (s.rows.offset(r1) - p.cy0) * z;
      const x2 = p.x0 + (s.cols.offset(Math.min(c2, MAX_COLS - 1) + 1) - p.cx0) * z;
      const y2 = p.y0 + (s.rows.offset(Math.min(r2, MAX_ROWS - 1) + 1) - p.cy0) * z;
      return { x, y, w: x2 - x, h: y2 - y, cutTop: r1 !== a.r1, cutLeft: c1 !== a.c1, cutBottom: r2 !== a.r2, cutRight: c2 !== a.c2 };
    };
    // the coloured ranges of a formula that is edited
    for (const hl of this.highlights) {
      if (hl.sheet !== s.id) continue;
      const r = rectOf(hl.area);
      if (!r) continue;
      g.fillStyle = hl.color + '1f';
      g.fillRect(r.x, r.y, r.w, r.h);
      g.strokeStyle = hl.color;
      g.lineWidth = 2;
      g.setLineDash(hl.dashed ? [5, 3] : []);
      g.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
      g.setLineDash([]);
      g.fillStyle = hl.color;
      for (const [cx, cy] of [[r.x, r.y], [r.x + r.w - 2, r.y], [r.x, r.y + r.h - 2], [r.x + r.w - 2, r.y + r.h - 2]]) g.fillRect(cx - 2, cy - 2, 5, 5);
    }
    if (app.editor?.active && app.editor.pointing) { /* the selection is drawn as a pointed range already */ }
    const sel = app.sel;
    const showSel = !app.editor?.active || app.editor.sheet === s;
    if (showSel && sel.sheetId === s.id) {
      const act = sel.active;
      const activeArea = s.mergeAt(act.row, act.col) || { r1: act.row, c1: act.col, r2: act.row, c2: act.col };
      g.fillStyle = t.selFill;
      const ar = rectOf(activeArea);
      for (const a of sel.ranges) {
        const r = rectOf(a);
        if (!r) continue;
        if (sel.ranges.length === 1 && a.r1 === activeArea.r1 && a.r2 === activeArea.r2 && a.c1 === activeArea.c1 && a.c2 === activeArea.c2) continue;
        if (ar && areaContains(a, act.row, act.col)) {
          // the active cell keeps its own look: shade around it
          const top = ar.y - r.y;
          const bottom = r.y + r.h - (ar.y + ar.h);
          if (top > 0) g.fillRect(r.x, r.y, r.w, top);
          if (bottom > 0) g.fillRect(r.x, ar.y + ar.h, r.w, bottom);
          if (ar.x > r.x) g.fillRect(r.x, ar.y, ar.x - r.x, ar.h);
          if (r.x + r.w > ar.x + ar.w) g.fillRect(ar.x + ar.w, ar.y, r.x + r.w - (ar.x + ar.w), ar.h);
        } else {
          g.fillRect(r.x, r.y, r.w, r.h);
        }
      }
      // the border of the last range (and the active cell)
      const last = rectOf(sel.range);
      g.strokeStyle = t.accent;
      if (last) {
        g.lineWidth = 2;
        g.strokeRect(last.x - 0.5, last.y - 0.5, last.w, last.h);
        if (!app.readOnly && !(sel.range.r2 === MAX_ROWS - 1 && sel.range.c2 === MAX_COLS - 1) && !last.cutBottom && !last.cutRight && !app.editor?.active) {
          g.fillStyle = t.cell;
          g.fillRect(last.x + last.w - 4, last.y + last.h - 4, 7, 7);
          g.fillStyle = t.accent;
          g.fillRect(last.x + last.w - 3, last.y + last.h - 3, 5, 5);
        }
      }
      if (ar && sel.ranges.length > 1) { g.lineWidth = 1; g.strokeRect(ar.x + 0.5, ar.y + 0.5, ar.w - 2, ar.h - 2); }
    }
    // spill range of the active cell's formula
    const spill = app.spillOf?.();
    if (spill && spill.sheet === s.id) {
      const r = rectOf(spill.area);
      if (r) { g.strokeStyle = '#4472c4'; g.lineWidth = 1; g.setLineDash([3, 2]); g.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1); g.setLineDash([]); }
    }
    // the fill handle being dragged, a move being dragged
    for (const preview of [this.fillPreview, this.movePreview]) {
      if (!preview || preview.sheet !== s.id) continue;
      const r = rectOf(preview.area);
      if (!r) continue;
      g.strokeStyle = t.dark ? '#bbbbbb' : '#555555';
      g.lineWidth = 1;
      g.setLineDash([3, 3]);
      g.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1);
      g.setLineDash([]);
    }
    // the marching ants of a copy
    if (this.marquee && this.marquee.sheet === s.id) {
      const r = rectOf(this.marquee.area);
      if (r) {
        g.lineWidth = 2;
        g.strokeStyle = t.accent;
        g.setLineDash([5, 4]);
        g.lineDashOffset = -this.dashOffset;
        g.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
        g.setLineDash([]);
        g.lineDashOffset = 0;
      }
    }
  }

  /** The marching ants around copied cells (null stops them). */
  setMarquee(m) {
    this.marquee = m;
    clearInterval(this.antTimer);
    if (m) this.antTimer = setInterval(() => { this.dashOffset = (this.dashOffset + 1) % 9; this.requestRender(); }, 90);
    this.requestRender();
  }

  drawHeaders(g, m, t) {
    const s = m.s;
    const z = m.z;
    const sel = this.app.sel;
    const showSel = !this.override && sel.sheetId === s.id;
    const fontPx = Math.max(8, Math.round(11 * z * 10) / 10);
    g.font = `${fontPx}px "Segoe UI", system-ui, sans-serif`;
    g.textBaseline = 'middle';
    g.textAlign = 'center';
    const px = 1 / (this.dpr || 1);
    const selCols = (c) => showSel && sel.ranges.some(a => c >= a.c1 && c <= a.c2);
    const fullCols = (c) => showSel && sel.ranges.some(a => c >= a.c1 && c <= a.c2 && a.r1 === 0 && a.r2 === MAX_ROWS - 1);
    const selRows = (r) => showSel && sel.ranges.some(a => r >= a.r1 && r <= a.r2);
    const fullRows = (r) => showSel && sel.ranges.some(a => r >= a.r1 && r <= a.r2 && a.c1 === 0 && a.c2 === MAX_COLS - 1);
    // column headers
    g.fillStyle = t.header;
    g.fillRect(m.hw, 0, this.width - m.hw, m.hh);
    for (const p of m.panes.filter(x => x.rowMin === m.fr || (m.fr === 0 && x.rowMin === 0))) {
      if (p.rowMin !== (m.fr ? m.fr : 0)) continue;
      g.save();
      g.beginPath();
      g.rect(p.x0, 0, p.w, m.hh);
      g.clip();
      for (let c = p.c1; c <= p.c2; c++) {
        const w = s.cols.size(c) * z;
        if (w <= 0) continue;
        const x = p.x0 + (s.cols.offset(c) - p.cx0) * z;
        if (selCols(c)) {
          g.fillStyle = fullCols(c) ? t.headerAll : t.headerSel;
          g.fillRect(x, 0, w, m.hh);
          g.fillStyle = t.accent;
          g.fillRect(x, m.hh - 2, w, 2);
        }
        g.fillStyle = t.headerLine;
        g.fillRect(snap(x + w, this.dpr) - px, 0, px, m.hh);
        g.fillStyle = selCols(c) ? t.headerSelText : t.headerText;
        if (w > 10) g.fillText(colName(c), x + w / 2, m.hh / 2 + 0.5);
      }
      g.restore();
    }
    // row headers
    g.fillStyle = t.header;
    g.fillRect(0, m.hh, m.hw, this.height - m.hh);
    for (const p of m.panes) {
      if (p.colMin !== (m.fc ? m.fc : 0)) continue;
      g.save();
      g.beginPath();
      g.rect(0, p.y0, m.hw, p.h);
      g.clip();
      for (let r = p.r1; r <= p.r2; r++) {
        const h = s.rows.size(r) * z;
        if (h <= 0) continue;
        const y = p.y0 + (s.rows.offset(r) - p.cy0) * z;
        if (selRows(r)) {
          g.fillStyle = fullRows(r) ? t.headerAll : t.headerSel;
          g.fillRect(0, y, m.hw, h);
          g.fillStyle = t.accent;
          g.fillRect(m.hw - 2, y, 2, h);
        }
        g.fillStyle = t.headerLine;
        g.fillRect(0, snap(y + h, this.dpr) - px, m.hw, px);
        g.fillStyle = selRows(r) ? t.headerSelText : t.headerText;
        if (h > 8) g.fillText(String(r + 1), m.hw / 2, y + h / 2 + 0.5);
      }
      g.restore();
    }
    // the lines between headers and cells, the corner
    g.fillStyle = t.headerLine;
    g.fillRect(0, m.hh - px, this.width, px);
    g.fillRect(m.hw - px, 0, px, this.height);
    g.fillStyle = t.header;
    g.fillRect(0, 0, m.hw - px, m.hh - px);
    g.fillStyle = t.dark ? '#777777' : '#b4b4b4';
    g.beginPath();
    g.moveTo(m.hw - 4, m.hh - 4 - Math.min(10, m.hh * 0.5));
    g.lineTo(m.hw - 4, m.hh - 4);
    g.lineTo(m.hw - 4 - Math.min(10, m.hh * 0.5), m.hh - 4);
    g.closePath();
    g.fill();
    g.textAlign = 'start';
  }
}

const LIGHT = {
  dark: false, cell: '#ffffff', text: '#000000', grid: '#e1e1e1', header: '#f5f5f5', headerText: '#444444', headerLine: '#d4d4d4',
  headerSel: '#d3eadb', headerSelText: '#0b5a2f', headerAll: '#9fd5b7', accent: '#107c41', selFill: 'rgba(16, 124, 65, 0.11)', freeze: '#9a9a9a', outside: '#ffffff',
};

function snap(v, dpr = 1) { return Math.round(v * dpr) / dpr; }

function isDarkCss(color) {
  const m = /rgba?\((\d+)[ ,]+(\d+)[ ,]+(\d+)/.exec(color);
  if (m) return 0.299 * m[1] + 0.587 * m[2] + 0.114 * m[3] < 128;
  if (/^#[0-9a-f]{6}$/i.test(color)) return isDark(color);
  return false;
}
