/* MatOffice Docs – the page view (Print Layout like Word): white sheets on a grey desk and the text flowing over them.
   The body is one editing host inside a "flow" that the browser paginates itself: CSS multi-column with column-height (the text height of a page)
   and column-wrap (every row of columns is one page; row-gap = the margins and the gap between two sheets). Lines of a paragraph that do not fit
   continue on the next sheet – nothing is split in the DOM, caret and selection stay native. The sheets (with their header and footer) are drawn
   behind the flow at the same distances. Browsers without column-wrap get a continuous page with page-break lines.
   Zoom is a transform (the layout and so the page breaks never depend on it). Printing uses the browser's own pages (@page with the size and the
   margins, the header and footer as margin boxes). */
import { pageGeometry, PX_PER_MM, fieldText, numStr } from './model.js';
import { parseHtml, renderBlock } from './htmlio.js';

export const SUPPORTS_PAGES = typeof CSS !== 'undefined' && CSS.supports?.('column-wrap', 'wrap') && CSS.supports?.('column-height', '100px');
const GAP = 20;          // between two sheets (px)
const PAD = 24;          // around the sheets

export class PageView {
  constructor({ el, t, lang, name }) {
    this.el = el;
    this.t = t;
    this.lang = lang;
    this.name = name;
    this.zoom = 1;
    this.view = 'print';
    this.pages = 1;
    this.page = null;
    this.headers = { header: '', footer: '', firstHeader: '', firstFooter: '' };
    this.railWidth = 0;
    this.scroller = el('div', { class: 'mo-docs-scroller', tabindex: '-1' });
    this.sizer = el('div', { class: 'mo-docs-sizer' });
    this.canvas = el('div', { class: 'mo-docs-canvas' });
    this.stack = el('div', { class: 'mo-docs-stack' });
    this.sheets = el('div', { class: 'mo-docs-sheets', 'aria-hidden': 'true' });
    this.flow = el('div', { class: 'mo-docs-flow' });
    this.body = el('div', { class: 'mo-docs-body mo-docs-content', 'aria-label': t('Docs.Body') });
    this.notes = el('div', { class: 'mo-docs-notes', hidden: true });
    this.rail = el('div', { class: 'mo-docs-rail', hidden: true });
    this.flow.append(this.body, this.notes);
    this.stack.append(this.sheets, this.flow);
    this.canvas.append(this.stack, this.rail);
    this.sizer.append(this.canvas);
    this.scroller.append(this.sizer);
    this.paged = SUPPORTS_PAGES;
    this.stack.classList.toggle('is-continuous', !this.paged);
    new ResizeObserver(() => this.schedule()).observe(this.flow);
    new ResizeObserver(() => { if (this.fitting) this.fit(); else this.size(); }).observe(this.scroller);
  }

  /** The page setup (content.page) – the CSS variables of the geometry. */
  setPage(page) {
    this.page = page;
    const g = this.geometry = pageGeometry(page);
    const s = this.stack.style;
    const px = v => numStr(v) + 'px';
    s.setProperty('--pw', px(g.paperWidth));
    s.setProperty('--ph', px(g.paperHeight));
    s.setProperty('--mt', px(g.top));
    s.setProperty('--mr', px(g.right));
    s.setProperty('--mb', px(g.bottom));
    s.setProperty('--ml', px(g.left));
    s.setProperty('--tw', px(g.textWidth));
    s.setProperty('--th', px(g.textHeight));
    s.setProperty('--hd', px(g.header));
    s.setProperty('--fd', px(g.footer));
    s.setProperty('--gap', px(GAP));
    s.setProperty('--cols', String(g.columns));
    s.setProperty('--cgap', px(g.columnGap));
    this.stack.dataset.columns = String(g.columns);
    this.sheetCache = new Map();
    this.update();
  }

  setHeaders(headers) {
    this.headers = { ...this.headers, ...headers };
    this.sheetCache = new Map();
    this.update();
  }

  setView(view) {
    this.view = view === 'web' ? 'web' : 'print';
    this.stack.classList.toggle('is-web', this.view === 'web');
    this.update();
    if (this.fitting) this.fit();
  }

  // ---- zoom ---------------------------------------------------------------------------------------------------------------------

  setZoom(f, keepFit = false) {
    this.zoom = Math.min(5, Math.max(0.1, f));
    if (!keepFit) this.fitting = false;
    this.canvas.style.transform = this.zoom === 1 ? '' : `scale(${this.zoom})`;
    this.size();
  }

  /** Page width: the sheets fill the width of the window. */
  fit() {
    this.fitting = true;
    const width = this.scroller.clientWidth - 8;
    const natural = this.naturalWidth();
    if (width > 0 && natural > 0) this.setZoom(Math.min(3, width / natural), true);
  }

  /** One page: the whole sheet can be seen. */
  fitPage() {
    const h = this.scroller.clientHeight - 8, w = this.scroller.clientWidth - 8;
    const g = this.geometry;
    if (!g || h <= 0) return;
    this.setZoom(Math.min(w / this.naturalWidth(), h / (g.paperHeight + 2 * PAD)));
  }

  naturalWidth() {
    if (this.view === 'web') return this.scroller.clientWidth;
    return (this.geometry?.paperWidth || 800) + 2 * PAD + this.railWidth;
  }

  /** The size of the scrolled area follows the scaled canvas. */
  size() {
    const w = this.view === 'web' ? this.scroller.clientWidth / this.zoom : this.naturalWidth();
    this.canvas.style.width = numStr(w) + 'px';
    const h = this.canvas.offsetHeight;
    this.sizer.style.width = numStr(w * this.zoom) + 'px';
    this.sizer.style.height = numStr(h * this.zoom) + 'px';
  }

  setRail(width) {
    this.railWidth = width;
    this.rail.hidden = !width;
    this.canvas.style.setProperty('--rail', width + 'px');
    if (this.fitting) this.fit(); else this.size();
  }

  // ---- pages --------------------------------------------------------------------------------------------------------------------

  schedule() {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => { this.frame = 0; this.update(); });
  }

  /** Counts the pages, draws the sheets with headers and footers, and the things that depend on the pages (TOC numbers, breaks). */
  update() {
    const g = this.geometry;
    if (!g) return;
    if (this.view === 'web') {
      this.pages = Math.max(1, Math.ceil(this.flow.offsetHeight / (g.textHeight || 1)));
      this.sheets.replaceChildren();
      this.stack.style.height = '';
      this.size();
      this.onPages?.(this.pages);
      return;
    }
    if (this.paged) this.fixBreaks();
    const h = this.flow.offsetHeight;
    const pitch = g.paperHeight + GAP;
    const pages = this.paged
      ? Math.max(1, Math.round((h - g.top + (g.top + g.bottom + GAP)) / pitch))
      : Math.max(1, Math.ceil((h - g.top - g.bottom) / g.textHeight));
    this.pages = pages;
    if (this.paged) this.stack.style.height = numStr(pages * pitch - GAP) + 'px';
    else this.stack.style.height = numStr(Math.max(g.paperHeight, h + g.bottom)) + 'px';
    this.drawSheets(pages);
    this.size();
    this.onPages?.(pages);
  }

  drawSheets(n) {
    const g = this.geometry;
    const pitch = g.paperHeight + GAP;
    const sheets = this.sheets;
    if (!this.paged) {
      // one long sheet with the lines where a page would end
      while (sheets.children.length > 1) sheets.lastChild.remove();
      let sheet = sheets.firstChild;
      if (!sheet) { sheet = this.makeSheet(); sheets.append(sheet); }
      sheet.style.top = '0px';
      sheet.style.height = this.stack.style.height;
      sheet.querySelectorAll('.mo-docs-pageline').forEach(x => x.remove());
      for (let i = 1; i < n; i++) {
        const line = document.createElement('div');
        line.className = 'mo-docs-pageline';
        line.style.top = numStr(g.top + i * g.textHeight) + 'px';
        sheet.append(line);
      }
      this.fillSheet(sheet, 0, n);
      return;
    }
    while (sheets.children.length < n) sheets.append(this.makeSheet());
    while (sheets.children.length > n) sheets.lastChild.remove();
    [...sheets.children].forEach((sheet, i) => {
      sheet.style.top = numStr(i * pitch) + 'px';
      sheet.style.height = '';
      this.fillSheet(sheet, i, n);
    });
  }

  makeSheet() {
    const sheet = document.createElement('div');
    sheet.className = 'mo-docs-sheet';
    const hdr = document.createElement('div');
    hdr.className = 'mo-docs-hdr mo-docs-content';
    const ftr = document.createElement('div');
    ftr.className = 'mo-docs-ftr mo-docs-content';
    sheet.append(hdr, ftr);
    return sheet;
  }

  /** The header and footer of a sheet (the first page has its own when the document says so). */
  fillSheet(sheet, index, pages) {
    const first = index === 0 && this.page?.differentFirst;
    const header = first ? this.headers.firstHeader : this.headers.header;
    const footer = first ? this.headers.firstFooter : this.headers.footer;
    const key = `${header}\u0001${footer}\u0001${index}\u0001${pages}\u0001${this.name()}`;
    if (sheet._key === key) return;
    sheet._key = key;
    const hdr = sheet.querySelector('.mo-docs-hdr');
    const ftr = sheet.querySelector('.mo-docs-ftr');
    const options = { fieldText: kind => fieldText(kind, { lang: this.lang, name: this.name(), page: index + 1, pages }) };
    const draw = (box, html) => {
      if (box.classList.contains('is-editing')) return;
      box.replaceChildren(...(html ? parseHtml(html) : []).map(b => renderBlock(b, options)));
      box.classList.toggle('is-empty', !html);
    };
    draw(hdr, header);
    draw(ftr, footer);
  }

  /** In a page of several columns a page break must jump over the columns left on that page: the break grows to fill them. */
  fixBreaks() {
    const g = this.geometry;
    const breaks = [...this.body.querySelectorAll(':scope > div[data-break="page"]')];
    this.quiet(() => {
      for (const b of breaks) b.style.height = '';
      if (g.columns <= 1 || !breaks.length) return;
      for (let pass = 0; pass < 3; pass++) {
        let changed = false;
        const stackTop = this.stack.getBoundingClientRect().top;
        const flowLeft = this.flow.getBoundingClientRect().left;
        const colPitch = (g.textWidth - (g.columns - 1) * g.columnGap) / g.columns + g.columnGap;
        for (const b of breaks) {
          const r = b.getBoundingClientRect();
          const y = (r.top - stackTop) / this.zoom;
          const x = (r.left - flowLeft) / this.zoom;
          const row = Math.floor(y / (g.paperHeight + GAP));
          const within = y - (row * (g.paperHeight + GAP) + g.top);
          const column = Math.min(g.columns - 1, Math.max(0, Math.floor((x + 1) / colPitch)));
          const skip = g.columns - 1 - column;
          const height = skip > 0 || within > 0 ? Math.max(0, g.textHeight - within - 1 + skip * g.textHeight) : 0;
          const value = height > 0 ? numStr(height) + 'px' : '';
          if (b.style.height !== value) { b.style.height = value; changed = true; }
        }
        if (!changed) break;
      }
    });
  }

  quiet(fn) { if (this.quietly) this.quietly(fn); else fn(); }

  // ---- positions ---------------------------------------------------------------------------------------------------------------

  /** The page (0-based) of a client rectangle. */
  pageOfRect(rect) {
    const g = this.geometry;
    if (!g || !rect) return 0;
    const top = (rect.top - this.stack.getBoundingClientRect().top) / this.zoom;
    if (this.view === 'web') return Math.max(0, Math.min(this.pages - 1, Math.floor(top / g.textHeight)));
    if (!this.paged) return Math.max(0, Math.min(this.pages - 1, Math.floor((top - g.top) / g.textHeight)));
    return Math.max(0, Math.min(this.pages - 1, Math.floor(top / (g.paperHeight + GAP))));
  }

  /** The page the reader looks at (the middle of the window). */
  visiblePage() {
    const r = this.scroller.getBoundingClientRect();
    return this.pageOfRect({ top: r.top + r.height / 3 });
  }

  scrollToPage(index) {
    const g = this.geometry;
    if (!g) return;
    const top = this.view === 'print' && this.paged ? index * (g.paperHeight + GAP) : index * g.textHeight;
    this.scroller.scrollTo({ top: Math.max(0, (top + PAD - 8) * this.zoom), behavior: 'smooth' });
  }

  /** Local coordinates (unscaled, relative to the stack) of a client rectangle. */
  local(rect) {
    const s = this.stack.getBoundingClientRect();
    return { top: (rect.top - s.top) / this.zoom, left: (rect.left - s.left) / this.zoom, bottom: (rect.bottom - s.top) / this.zoom, right: (rect.right - s.left) / this.zoom };
  }

  /** The sheet and the area (header/footer/body) under a client point. */
  hitSheet(x, y) {
    const g = this.geometry;
    if (!g || this.view === 'web') return null;
    const s = this.stack.getBoundingClientRect();
    const ly = (y - s.top) / this.zoom, lx = (x - s.left) / this.zoom;
    if (lx < 0 || lx > g.paperWidth) return null;
    const pitch = this.paged ? g.paperHeight + GAP : Infinity;
    const index = this.paged ? Math.floor(ly / pitch) : 0;
    if (index < 0 || index >= this.pages) return null;
    const within = this.paged ? ly - index * pitch : ly;
    if (within > g.paperHeight && this.paged) return null;
    const area = within < g.top ? 'header' : this.paged && within > g.paperHeight - g.bottom ? 'footer' : 'body';
    return { index, area };
  }

  // ---- print ---------------------------------------------------------------------------------------------------------------------

  /** The @page rules for printing: the paper, the margins, the header and footer as margin boxes (text, fields as counters). */
  printCss({ header, footer, firstHeader, firstFooter }) {
    const page = this.page;
    const paperW = page.orientation === 'landscape' ? page.height : page.width;
    const paperH = page.orientation === 'landscape' ? page.width : page.height;
    const m = page.margins;
    const mm = v => numStr(v) + 'mm';
    let css = `@page { size: ${mm(paperW)} ${mm(paperH)}; margin: ${mm(m.top)} ${mm(m.right)} ${mm(m.bottom)} ${mm(m.left)};`;
    css += boxes(header, 'top', page.header, this) + boxes(footer, 'bottom', page.footer, this) + ' }';
    if (page.differentFirst) css += ` @page :first { ${boxes(firstHeader, 'top', page.header, this, true)}${boxes(firstFooter, 'bottom', page.footer, this, true)} }`;
    return css;
  }
}

/** The margin boxes of a header or footer: each paragraph goes to the box of its alignment, the text as a CSS string with counters for fields. */
function boxes(html, edge, distanceMm, view, reset = false) {
  const groups = { left: [], center: [], right: [] };
  const blocks = html ? parseHtml(html) : [];
  let font = null;
  for (const b of blocks) {
    if (!b.runs) continue;
    const align = b.css?.['text-align'] === 'center' ? 'center' : b.css?.['text-align'] === 'right' ? 'right' : 'left';
    const parts = [];
    for (const r of b.runs) {
      if (!font && r.f && (r.f.font || r.f.size || r.f.color || r.f.b || r.f.i)) font = r.f;
      if (!r.obj) parts.push(cssString(r.text));
      else if (r.obj === 'br') parts.push('"\\A "');
      else if (r.obj === 'field') {
        if (r.kind === 'page') parts.push('counter(page)');
        else if (r.kind === 'pages') parts.push('counter(pages)');
        else parts.push(cssString(fieldText(r.kind, { lang: view.lang, name: view.name() })));
      }
    }
    if (parts.length) groups[align].push(parts.join(' '));
  }
  const pad = `${numStr(Math.max(0, distanceMm))}mm`;
  let css = '';
  for (const [align, lines] of Object.entries(groups)) {
    const name = `@${edge}-${align}`;
    if (!lines.length) { if (reset) css += ` ${name} { content: none; }`; continue; }
    const content = lines.join(' "\\A " ');
    const style = [
      `content: ${content}`, 'white-space: pre-wrap', `vertical-align: ${edge === 'top' ? 'top' : 'bottom'}`,
      `padding-${edge}: ${pad}`, `text-align: ${align}`,
      `font-family: ${font?.font || "Calibri, Carlito, sans-serif"}`, `font-size: ${font?.size || '11pt'}`, `color: ${font?.color || '#000'}`,
      font?.b ? 'font-weight: bold' : '', font?.i ? 'font-style: italic' : '',
    ].filter(Boolean).join('; ');
    css += ` ${name} { ${style}; }`;
  }
  return css;
}

/** Text as a CSS string literal (quotes, backslashes and line ends escaped – the text of a document cannot leave the string). */
export function cssString(text) {
  return '"' + String(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\A ').replace(/[\u0000-\u001f\u007f]/g, ' ') + '"';
}

export { PX_PER_MM };
