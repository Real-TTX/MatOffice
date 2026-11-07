/* MatOffice Docs – between HTML, the DOM of the editor and the model (model.js):
   - parseHtml(html): the client sanitizer. Any HTML (the content of the server, the clipboard of Word, Google Docs or a web page) is parsed in an inert
     document (DOMParser runs no script and loads nothing) and read through an allowlist into blocks; nothing of it ever goes into innerHTML.
   - readBlock(el) / readBlocks(container): the live DOM of the editor back into blocks (whatever the browser made of it while typing).
   - renderBlock(block): a block as DOM for the editor (empty blocks get a placeholder <br>, the last <br> of a block is never content).
   - offsetOf / pointAt: a DOM position <-> a character offset in a text block (pictures, line breaks, fields and note references count as one). */
import {
  TEXT_TAGS, P_STYLES, isListKind, cleanFormat, normalizeRuns, textRun, blockCss, inlineFormat, parseStyle, fontStack, colour, safeHref, safeImage,
  isId, isItemId, toPt, numStr, TABLE_STYLES, cleanText, runsLength,
} from './model.js';

const DROP = new Set(['script', 'style', 'template', 'iframe', 'object', 'embed', 'noscript', 'head', 'title', 'meta', 'link', 'svg', 'math', 'canvas',
  'video', 'audio', 'source', 'track', 'input', 'button', 'select', 'textarea', 'option', 'datalist', 'map', 'area', 'frame', 'frameset', 'applet', 'base',
  'xml', 'colgroup', 'col', 'caption', 'picture', 'dialog', 'portal']);
const BLOCKISH = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'div', 'section', 'article', 'header', 'footer', 'main', 'aside',
  'nav', 'figure', 'figcaption', 'address', 'center', 'li', 'ul', 'ol', 'dl', 'dt', 'dd', 'table', 'hr', 'form', 'fieldset', 'details', 'summary', 'body', 'html']);
const FONT_SIZES = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 };
const MONO = fontStack('Consolas');

/** The tag name of an element in lower case; Word's o:p, v:shape … keep their prefix (unknown inline elements, never a paragraph). */
const tagOf = el => (el.localName || el.nodeName || '').toLowerCase();

const isPlaceholderBr = (block, br) => {
  // the last <br> of a block (nothing visible after it) is not content
  for (let n = nextInFlow(br, block); n; n = nextInFlow(n, block)) {
    if (n.nodeType === 3 && n.data.length) return false;
    if (n.nodeType === 1 && (isAtomic(n) || tagOf(n) === 'br')) return false;
  }
  return true;
};

function nextInFlow(node, root) {
  if (node.nodeType === 1 && node.firstChild && !isAtomic(node)) return node.firstChild;
  while (node && node !== root) {
    if (node.nextSibling) return node.nextSibling;
    node = node.parentNode;
  }
  return null;
}

/** Elements that are one character of the text: pictures, line breaks, fields, note references. */
export function isAtomic(n) {
  if (n.nodeType !== 1) return false;
  const tag = tagOf(n);
  return tag === 'img' || tag === 'br' || (tag === 'span' && n.hasAttribute('data-field')) || (tag === 'sup' && n.hasAttribute('data-fn'));
}

/** Decorations of the editor inside the content (resize handles, labels): not content, not text. */
const isDecoration = n => n.nodeType === 1 && n.hasAttribute && n.hasAttribute('data-mo-deco');

// ---- reading HTML (content, clipboard) -------------------------------------------------------------------------------------------

/**
 * Parses HTML into blocks through the allowlist. options: collapse (collapse source whitespace like a browser: true for foreign HTML),
 * images (a map of picture keys of the undo history back to data: addresses).
 */
export function parseHtml(html, options = {}) {
  const doc = new DOMParser().parseFromString('<!DOCTYPE html><html><head></head><body>' + (html || '') + '</body></html>', 'text/html');
  const reader = new Reader(options);
  reader.container(doc.body, reader.blocks, {});
  return reader.finish();
}

class Reader {
  constructor(options) {
    this.options = options;
    this.blocks = [];
    this.open = null;
    this.depth = 0;
  }

  finish() {
    this.close();
    const out = this.blocks.filter(b => b);
    return out;
  }

  close() {
    if (this.open) {
      const b = this.open.block;
      b.runs = finishRuns(b.runs, b.type === 'pre', this.options.collapse);
      this.open = null;
    }
  }

  /** The text block inline content goes into (opened with the attributes of the context). */
  target(out, ctx) {
    if (this.open && this.open.out === out) return this.open.block;
    this.close();
    const b = { type: ctx.type || 'p', css: { ...(ctx.css || {}) }, runs: [] };
    if (ctx.style) b.style = ctx.style;
    if (ctx.list) { b.list = ctx.list; b.level = ctx.level || 0; if (ctx.start) b.start = ctx.start; ctx.start = null; }
    if (ctx.id && isId(ctx.id)) { b.id = ctx.id; ctx.id = null; }
    out.push(b);
    this.open = { out, block: b };
    return b;
  }

  /** Reads the children of a block container (body, div, cell, list) into out. */
  container(el, out, ctx) {
    if (++this.depth > 60) { this.depth--; return; }
    for (const child of [...el.childNodes]) this.node(child, out, ctx, {});
    this.depth--;
  }

  node(n, out, ctx, f) {
    if (n.nodeType === 3) {
      const text = n.data;
      if (this.options.collapse && !ctx.pre && !(this.open && this.open.out === out) && !/[^\s]/.test(text)) return; // whitespace between blocks
      if (!text) return;
      this.target(out, ctx).runs.push({ text, f: cleanFormat(f) });
      return;
    }
    if (n.nodeType !== 1) return;
    const tag = tagOf(n);
    if (DROP.has(tag) || isDecoration(n)) return;
    const style = n.getAttribute('style') || '';
    if (/mso-list\s*:\s*ignore/i.test(style) || /display\s*:\s*none/i.test(style)) return;
    if (BLOCKISH.has(tag)) { this.block(n, tag, out, ctx); return; }
    this.inline(n, tag, out, ctx, f);
  }

  inline(n, tag, out, ctx, f) {
    if (tag === 'br') { this.target(out, ctx).runs.push({ obj: 'br', f: cleanFormat(f) }); return; }
    if (tag === 'img') {
      const src = this.options.images ? this.options.images(n.getAttribute('src')) : n.getAttribute('src');
      const safe = safeImage(src);
      if (!safe) return;
      const css = parseStyle(n.getAttribute('style'));
      const w = sizePx(css.find(d => d[0] === 'width')?.[1]) ?? numAttr(n.getAttribute('width'));
      const h = sizePx(css.find(d => d[0] === 'height')?.[1]) ?? numAttr(n.getAttribute('height'));
      const align = n.getAttribute('data-align');
      const run = { obj: 'img', src: safe, alt: cleanText(n.getAttribute('alt') || '').replace(/\n/g, ' ').slice(0, 1000), f: cleanFormat(linkOnly(f)) };
      if (w) run.width = w;
      if (h) run.height = h;
      if (align === 'left' || align === 'right' || align === 'center') run.align = align;
      else { const fl = css.find(d => d[0] === 'float')?.[1]; if (fl === 'left' || fl === 'right') run.align = fl; }
      this.target(out, ctx).runs.push(run);
      return;
    }
    if (tag === 'span' && n.hasAttribute('data-field')) {
      const kind = n.getAttribute('data-field');
      if (['page', 'pages', 'date', 'time', 'title'].includes(kind)) this.target(out, ctx).runs.push({ obj: 'field', kind, f: cleanFormat(f) });
      return;
    }
    if (tag === 'sup' && n.hasAttribute('data-fn')) {
      const id = n.getAttribute('data-fn');
      if (isItemId(id)) this.target(out, ctx).runs.push({ obj: 'note', id, f: cleanFormat(linkOnly(f)) });
      return;
    }
    const g = { ...f };
    switch (tag) {
      case 'strong': case 'b': g.b = true; break;
      case 'em': case 'i': case 'cite': case 'dfn': case 'var': g.i = true; break;
      case 'u': case 'ins': g.u = true; break;
      case 's': case 'strike': case 'del': g.s = true; break;
      case 'sub': g.sub = true; delete g.sup; break;
      case 'sup': g.sup = true; delete g.sub; break;
      case 'code': case 'kbd': case 'samp': case 'tt': g.font = MONO; break;
      case 'mark': g.bg = '#ffff00'; break;
      case 'a': { const href = safeHref(n.getAttribute('href')); if (href) g.href = href; break; }
      case 'font': {
        const face = n.getAttribute('face');
        if (face) g.font = fontStack(face.split(',')[0]);
        const c = colour(n.getAttribute('color'));
        if (c) g.color = c;
        const size = FONT_SIZES[n.getAttribute('size')];
        if (size) g.size = size + 'pt';
        break;
      }
      default: break;
    }
    if (tag === 'span' && n.hasAttribute('data-comment')) {
      const ids = (n.getAttribute('data-comment') || '').split(/\s+/).filter(isItemId);
      if (ids.length) g.cmt = [...new Set([...(g.cmt ? g.cmt.split(' ') : []), ...ids])].join(' ');
    }
    if (n.hasAttribute('style')) Object.assign(g, inlineFormat(parseStyle(n.getAttribute('style'))));
    for (const child of [...n.childNodes]) this.node(child, out, ctx, g);
  }

  block(n, tag, out, ctx) {
    const style = parseStyle(n.getAttribute('style'));
    switch (tag) {
      case 'hr': this.close(); out.push({ type: 'hr' }); return;
      case 'table': this.close(); { const t = this.table(n); if (t) out.push(t); } return;
      case 'ul': case 'ol': {
        this.close();
        const level = ctx.list ? Math.min(8, (ctx.level ?? 0) + 1) : 0;
        const kind = listKind(n, tag, style, level);
        const start = tag === 'ol' && Number(n.getAttribute('start')) > 1 ? Math.min(99999, Number(n.getAttribute('start'))) : null;
        const inner = { ...ctx, list: kind, level, start, type: 'p', style: null, inList: true };
        for (const child of [...n.childNodes]) {
          if (child.nodeType === 1 && tagOf(child) === 'li') this.listItem(child, out, inner);
          else this.node(child, out, inner, {});
          inner.start = null;
        }
        this.close();
        return;
      }
      case 'li': this.listItem(n, out, { ...ctx, list: ctx.list || 'bullet', level: ctx.level || 0 }); return;
      case 'div':
        if (n.getAttribute('data-break') === 'page' || n.getAttribute('data-break') === 'column') { this.close(); out.push({ type: 'break', kind: n.getAttribute('data-break') }); return; }
        if (/page-break-(before|after)\s*:\s*always|break-(before|after)\s*:\s*page/i.test(n.getAttribute('style') || '') && !n.textContent.trim()) { this.close(); out.push({ type: 'break', kind: 'page' }); return; }
        break;
      case 'br': return;
      default: break;
    }
    this.close();
    const inner = { ...ctx };
    const css = blockCss(style);
    const align = (n.getAttribute('align') || '').toLowerCase();
    if (align && !css['text-align'] && /^(left|right|center|justify)$/.test(align)) css['text-align'] = align;
    if (tag === 'center') css['text-align'] = 'center';
    if (TEXT_TAGS.has(tag)) {
      inner.type = tag;
      inner.style = null;
      const ds = n.getAttribute('data-style');
      if (tag === 'p' && P_STYLES.has(ds)) inner.style = ds;
      if (tag === 'blockquote' && ds === 'IntenseQuote') inner.style = ds;
      if (/^h[1-6]$/.test(tag) || tag === 'p') {
        const list = n.getAttribute('data-list');
        if (isListKind(list)) {
          inner.list = list;
          inner.level = Math.min(8, Math.max(0, parseInt(n.getAttribute('data-level') || '0', 10) || 0));
          const start = parseInt(n.getAttribute('data-start') || '', 10);
          inner.start = start >= 1 && start <= 99999 ? start : null;
        } else if (!ctx.inList) { delete inner.list; delete inner.level; delete inner.start; }
        const mso = /mso-list\s*:\s*l(\d+)\s+level(\d+)/i.exec(n.getAttribute('style') || '');
        if (mso && !isListKind(list)) {
          inner.list = wordListKind(n);
          inner.level = Math.min(8, Math.max(0, Number(mso[2]) - 1));
          inner.start = null;
        }
      } else { delete inner.list; delete inner.level; delete inner.start; }
      if (tag === 'pre') inner.pre = true;
      inner.css = css;
      inner.id = n.getAttribute('id');
      if (inner.list && inner.style === 'ListParagraph') inner.style = null;
      // a Word list paragraph keeps its own indents out (the list draws them)
      if (inner.list) { delete inner.css['margin-left']; delete inner.css['text-indent']; }
      this.target(out, inner);
      for (const child of [...n.childNodes]) this.node(child, out, inner, inheritedFormat(style));
      this.close();
      return;
    }
    // a wrapper (div, section, figure …): its alignment goes to the paragraphs inside
    if (css['text-align']) inner.css = { ...(inner.css || {}), 'text-align': css['text-align'] };
    inner.type = 'p';
    inner.style = tag === 'figcaption' ? 'Caption' : null;
    if (tag === 'dt') inner.style = null;
    if (!ctx.inList) { delete inner.list; delete inner.level; }
    this.container(n, out, inner);
    this.close();
  }

  listItem(li, out, ctx) {
    this.close();
    const inner = { ...ctx, type: 'p', style: null, inList: true };
    const css = blockCss(parseStyle(li.getAttribute('style')));
    delete css['margin-left'];
    delete css['text-indent'];
    inner.css = css;
    const aria = parseInt(li.getAttribute('aria-level') || '', 10);
    if (aria >= 1 && aria <= 9 && !ctx.nested) inner.level = Math.max(inner.level || 0, aria - 1);
    const value = parseInt(li.getAttribute('value') || '', 10);
    if (value >= 1) inner.start = value;
    for (const child of [...li.childNodes]) {
      if (child.nodeType === 1 && (tagOf(child) === 'ul' || tagOf(child) === 'ol')) { this.close(); this.block(child, tagOf(child), out, { ...inner, nested: true }); continue; }
      if (child.nodeType === 1 && tagOf(child) === 'p' && !this.open) {
        // <li><p>text</p></li> (Google Docs): the paragraph is the item
        this.target(out, inner);
        for (const c of [...child.childNodes]) this.node(c, out, inner, inheritedFormat(parseStyle(child.getAttribute('style'))));
        continue;
      }
      this.node(child, out, inner, {});
    }
    if (!this.open || this.open.out !== out) this.target(out, inner);
    this.close();
    ctx.start = null;
  }

  table(el) {
    if (++this.depth > 20) { this.depth--; return null; }
    const rows = [];
    const collect = node => {
      for (const child of [...node.children]) {
        const t = tagOf(child);
        if (t === 'tr') rows.push({ el: child, head: tagOf(node) === 'thead' });
        else if (t === 'thead' || t === 'tbody' || t === 'tfoot') collect(child);
      }
    };
    collect(el);
    const t = { type: 'table', tableStyle: 'grid', width: '100%', cols: [], rows: [] };
    const ds = el.getAttribute('data-table-style');
    if (TABLE_STYLES.includes(ds)) t.tableStyle = ds;
    else if (el.getAttribute('border') === '0' && this.options.collapse) t.tableStyle = 'grid';
    const w = parseStyle(el.getAttribute('style')).find(d => d[0] === 'width')?.[1] || el.getAttribute('width');
    if (w && /%$/.test(w)) t.width = Math.min(100, Math.max(5, parseFloat(w))) + '%';
    const cols = [...el.querySelectorAll(':scope > colgroup > col, :scope > col')].map(c => {
      const cw = parseStyle(c.getAttribute('style')).find(d => d[0] === 'width')?.[1] || c.getAttribute('width');
      return cw && /%$/.test(cw) ? numStr(Math.min(100, Math.max(1, parseFloat(cw)))) + '%' : null;
    });
    for (const { el: tr, head } of rows.slice(0, 5000)) {
      const cells = [...tr.children].filter(c => tagOf(c) === 'td' || tagOf(c) === 'th');
      if (!cells.length) continue;
      const row = { header: head || cells.every(c => tagOf(c) === 'th'), cells: [] };
      for (const td of cells) {
        const cell = { header: tagOf(td) === 'th', colspan: span(td.getAttribute('colspan')), rowspan: span(td.getAttribute('rowspan')), css: cellCss(td), blocks: [] };
        const saved = this.open;
        this.open = null;
        this.container(td, cell.blocks, {});
        this.close();
        this.open = saved;
        cell.blocks = cell.blocks.filter(b => b.type !== 'break');
        if (!cell.blocks.length) cell.blocks.push({ type: 'p', css: {}, runs: [] });
        row.cells.push(cell);
      }
      t.rows.push(row);
    }
    // only a leading run of header rows is a header
    let inHead = true;
    for (const r of t.rows) { if (!r.header) inHead = false; r.header = inHead && r.header; }
    if (!t.rows.length) { this.depth--; return null; }
    const n = Math.max(...t.rows.map(r => r.cells.reduce((s, c) => s + c.colspan, 0)));
    t.cols = cols.length === n && cols.every(Boolean) ? cols : widthsFromFirstRow(rows[0]?.el, n);
    this.depth--;
    return t;
  }
}

const span = v => Math.min(1000, Math.max(1, parseInt(v || '1', 10) || 1));
const numAttr = v => { const n = parseFloat(v); return Number.isFinite(n) && n > 0 && n < 20000 ? Math.round(n * 100) / 100 : null; };
const sizePx = v => { if (!v) return null; const p = toPt(v); return p && p > 0 ? Math.round(p / 0.75 * 100) / 100 : null; };
const linkOnly = f => ({ href: f.href, cmt: f.cmt });

/** Formatting a block element gives its text (Word puts fonts on the paragraph). */
function inheritedFormat(decls) {
  const f = inlineFormat(decls.filter(d => d[0] !== 'background-color' && d[0] !== 'background'));
  return f;
}

function cellCss(td) {
  const css = {};
  const decls = parseStyle(td.getAttribute('style'));
  const bg = colour(decls.find(d => d[0] === 'background-color' || d[0] === 'background')?.[1] || td.getAttribute('bgcolor'));
  if (bg) css['background-color'] = bg;
  const va = (decls.find(d => d[0] === 'vertical-align')?.[1] || td.getAttribute('valign') || '').toLowerCase();
  if (va === 'middle' || va === 'center') css['vertical-align'] = 'middle';
  else if (va === 'bottom') css['vertical-align'] = 'bottom';
  const block = blockCss(decls);
  for (const side of ['top', 'right', 'bottom', 'left']) if (block['border-' + side]) css['border-' + side] = block['border-' + side];
  return css;
}

function widthsFromFirstRow(tr, n) {
  if (!tr || !n) return [];
  const widths = [];
  for (const td of [...tr.children]) {
    const w = parseStyle(td.getAttribute('style')).find(d => d[0] === 'width')?.[1] || td.getAttribute('width');
    const colspan = span(td.getAttribute('colspan'));
    const px = w && /%$/.test(w) ? null : sizePx(w) ?? (Number(w) > 0 ? Number(w) : null);
    const pct = w && /%$/.test(w) ? parseFloat(w) : null;
    for (let i = 0; i < colspan; i++) widths.push(pct !== null ? { pct: pct / colspan } : px !== null ? { px: px / colspan } : null);
  }
  if (widths.length !== n || widths.some(w => !w)) return new Array(n).fill(numStr(100 / n) + '%');
  if (widths.every(w => w.pct !== undefined)) { const sum = widths.reduce((s, w) => s + w.pct, 0) || 100; return widths.map(w => numStr(w.pct / sum * 100) + '%'); }
  if (widths.every(w => w.px !== undefined)) { const sum = widths.reduce((s, w) => s + w.px, 0) || 1; return widths.map(w => numStr(w.px / sum * 100) + '%'); }
  return new Array(n).fill(numStr(100 / n) + '%');
}

function listKind(el, tag, style, level) {
  const type = (style.find(d => d[0] === 'list-style-type')?.[1] || el.getAttribute('type') || '').toLowerCase();
  if (tag === 'ul') {
    if (type === 'circle') return level % 3 === 1 ? 'bullet' : 'circle';
    if (type === 'square') return level % 3 === 2 ? 'bullet' : 'square';
    return 'bullet';
  }
  const attr = el.getAttribute('type');
  if (attr === 'a') return 'lower-alpha';
  if (attr === 'A') return 'upper-alpha';
  if (attr === 'i') return 'lower-roman';
  if (attr === 'I') return 'upper-roman';
  if (type === 'lower-alpha' || type === 'lower-latin') return level % 3 === 1 ? 'decimal' : 'lower-alpha';
  if (type === 'upper-alpha' || type === 'upper-latin') return 'upper-alpha';
  if (type === 'lower-roman') return level % 3 === 2 ? 'decimal' : 'lower-roman';
  if (type === 'upper-roman') return 'upper-roman';
  return 'decimal';
}

/** The kind of a Word list paragraph from its marker (the mso-list:Ignore span Word writes before the text). */
function wordListKind(p) {
  const ignore = [...p.querySelectorAll('span')].find(s => /mso-list\s*:\s*ignore/i.test(s.getAttribute('style') || ''));
  const marker = (ignore?.textContent || '').replace(/\s+/g, '').trim();
  if (!marker) return 'bullet';
  if (/^\d+[.)]?$/.test(marker)) return 'decimal';
  if (/^[ivxlc]+[.)]$/.test(marker)) return 'lower-roman';
  if (/^[IVXLC]+[.)]$/.test(marker)) return 'upper-roman';
  if (/^[a-z][.)]$/.test(marker)) return 'lower-alpha';
  if (/^[A-Z][.)]$/.test(marker)) return 'upper-alpha';
  if (/^\d+(\.\d+)+\.?$/.test(marker)) return 'outline';
  if (marker === 'o') return 'circle';
  if (marker === '§' || marker === '▪' || marker === '■') return 'square';
  if (marker === '-' || marker === '–') return 'dash';
  if (marker === 'ü' || marker === '✓') return 'check';
  if (marker === 'Ø' || marker === '➢') return 'arrow';
  return 'bullet';
}

/** The runs of a block when it is complete: line ends become breaks (not in code), the last break is the placeholder, whitespace as a browser shows it. */
function finishRuns(runs, pre, collapse) {
  let out = [];
  for (const r of runs) {
    if (r.obj) { out.push(r); continue; }
    let text = cleanText(r.text);
    if (collapse && !pre) text = text.replace(/[ \t\n\r\f]+/g, ' ');
    if (!pre && text.includes('\n')) {
      text.split('\n').forEach((part, i) => { if (i) out.push({ obj: 'br', f: r.f }); if (part) out.push({ text: part, f: r.f }); });
    } else if (text) out.push({ text, f: r.f });
  }
  if (collapse && !pre) {
    // trim like a browser: no space at the start, before a break, at the end
    while (out.length && !out[0].obj && !out[0].text.replace(/^ +/, '')) out.shift();
    if (out.length && !out[0].obj) out[0] = { text: out[0].text.replace(/^ +/, ''), f: out[0].f };
    for (let i = 0; i < out.length; i++) {
      if (out[i].obj === 'br' && i > 0 && !out[i - 1].obj) out[i - 1] = { text: out[i - 1].text.replace(/ +$/, ''), f: out[i - 1].f };
      if (out[i].obj === 'br' && i + 1 < out.length && !out[i + 1].obj) out[i + 1] = { text: out[i + 1].text.replace(/^ +/, ''), f: out[i + 1].f };
    }
    const last = out[out.length - 1];
    if (last && !last.obj) out[out.length - 1] = { text: last.text.replace(/ +$/, ''), f: last.f };
    if (out.every(r => !r.obj && !r.text.replace(/ /g, '').trim())) out = [];
  }
  if (out.length && out[out.length - 1].obj === 'br') out.pop();
  if (pre && out.length && !out[0].obj && out[0].text.startsWith('\n') && collapse) out[0] = { text: out[0].text.slice(1), f: out[0].f };
  return normalizeRuns(out);
}

// ---- reading the live DOM of the editor -----------------------------------------------------------------------------------------

/** A text block element of the editor as a block of the model. */
export function readBlock(el) {
  const type = tagOf(el);
  const b = { type: TEXT_TAGS.has(type) ? type : 'p', css: blockCss(parseStyle(el.getAttribute('style'))), runs: [] };
  const id = el.getAttribute('id');
  if (id && isId(id) && /^h[1-6]$/.test(b.type)) b.id = id;
  const ds = el.getAttribute('data-style');
  if (b.type === 'p' && P_STYLES.has(ds)) b.style = ds;
  if (b.type === 'blockquote' && ds === 'IntenseQuote') b.style = ds;
  const list = el.getAttribute('data-list');
  if ((b.type === 'p' || /^h[1-6]$/.test(b.type)) && isListKind(list)) {
    b.list = list;
    b.level = Math.min(8, Math.max(0, parseInt(el.getAttribute('data-level') || '0', 10) || 0));
    const start = parseInt(el.getAttribute('data-start') || '', 10);
    if (start >= 1 && start <= 99999) b.start = start;
  }
  const runs = [];
  walkInline(el, el, {}, runs, b.type === 'pre');
  b.runs = normalizeRuns(runs);
  return b;
}

function walkInline(block, node, f, runs, pre) {
  for (let n = node.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 3) {
      const text = cleanText(n.data);
      if (!text) continue;
      if (!pre && text.includes('\n')) text.split('\n').forEach((part, i) => { if (i) runs.push({ obj: 'br', f: cleanFormat(f) }); if (part) runs.push({ text: part, f: cleanFormat(f) }); });
      else runs.push({ text, f: cleanFormat(f) });
      continue;
    }
    if (n.nodeType !== 1 || isDecoration(n)) continue;
    const tag = tagOf(n);
    if (DROP.has(tag)) continue;
    if (tag === 'br') { if (!isPlaceholderBr(block, n)) runs.push({ obj: 'br', f: cleanFormat(f) }); continue; }
    if (tag === 'img') {
      const src = safeImage(n.getAttribute('src'));
      if (!src) continue;
      const r = { obj: 'img', src, alt: n.getAttribute('alt') || '', f: cleanFormat(linkOnly(f)) };
      const w = sizePx(n.style.width), h = sizePx(n.style.height);
      if (w) r.width = w;
      if (h) r.height = h;
      const align = n.getAttribute('data-align');
      if (align === 'left' || align === 'right' || align === 'center') r.align = align;
      runs.push(r);
      continue;
    }
    if (tag === 'span' && n.hasAttribute('data-field')) { runs.push({ obj: 'field', kind: n.getAttribute('data-field'), f: cleanFormat(f) }); continue; }
    if (tag === 'sup' && n.hasAttribute('data-fn')) { const id = n.getAttribute('data-fn'); if (isItemId(id)) runs.push({ obj: 'note', id, f: cleanFormat(linkOnly(f)) }); continue; }
    const g = { ...f };
    switch (tag) {
      case 'strong': case 'b': g.b = true; break;
      case 'em': case 'i': g.i = true; break;
      case 'u': case 'ins': g.u = true; break;
      case 's': case 'strike': case 'del': g.s = true; break;
      case 'sub': g.sub = true; delete g.sup; break;
      case 'sup': g.sup = true; delete g.sub; break;
      case 'a': { const href = safeHref(n.getAttribute('href')); if (href) g.href = href; break; }
      case 'font': {
        const face = n.getAttribute('face');
        if (face) g.font = fontStack(face.split(',')[0]);
        const c = colour(n.getAttribute('color'));
        if (c) g.color = c;
        break;
      }
      default: break;
    }
    if (tag === 'span' && n.hasAttribute('data-comment')) {
      const ids = (n.getAttribute('data-comment') || '').split(/\s+/).filter(isItemId);
      if (ids.length) g.cmt = [...new Set([...(g.cmt ? g.cmt.split(' ') : []), ...ids])].join(' ');
    }
    if (n.hasAttribute('style')) Object.assign(g, inlineFormat(parseStyle(n.getAttribute('style'))));
    walkInline(block, n, g, runs, pre);
  }
}

/** The blocks of a container of the editor (body, cell, note, header): text blocks, tables, lines, breaks; stray content becomes paragraphs. */
export function readBlocks(container) {
  const out = [];
  let loose = null;
  for (const n of [...container.childNodes]) {
    if (n.nodeType === 1 && isDecoration(n)) continue;
    const tag = n.nodeType === 1 ? tagOf(n) : '';
    if (n.nodeType === 1 && (TEXT_TAGS.has(tag) || tag === 'table' || tag === 'hr' || (tag === 'div' && n.hasAttribute('data-break')))) {
      loose = null;
      if (TEXT_TAGS.has(tag)) out.push(readBlock(n));
      else if (tag === 'table') out.push(readTable(n));
      else if (tag === 'hr') out.push({ type: 'hr' });
      else out.push({ type: 'break', kind: n.getAttribute('data-break') === 'column' ? 'column' : 'page' });
      continue;
    }
    if (n.nodeType === 1 && tagOf(n) === 'div') { loose = null; out.push(...readBlocks(n)); continue; }
    if (n.nodeType === 3 && !n.data.trim() && !loose) continue;
    if (n.nodeType !== 1 && n.nodeType !== 3) continue;
    if (!loose) { loose = { type: 'p', css: {}, runs: [] }; out.push(loose); }
    const wrap = document.createElement('p');
    wrap.append(n.cloneNode(true));
    loose.runs = normalizeRuns([...loose.runs, ...readBlock(wrap).runs]);
  }
  return out;
}

/** A table element of the editor as a table block. */
export function readTable(el) {
  const t = { type: 'table', tableStyle: TABLE_STYLES.includes(el.getAttribute('data-table-style')) ? el.getAttribute('data-table-style') : 'grid', width: null, cols: [], rows: [] };
  const w = el.style.width;
  if (w && /%$/.test(w)) t.width = numStr(Math.min(100, Math.max(5, parseFloat(w)))) + '%';
  t.cols = [...el.querySelectorAll(':scope > colgroup > col')].map(c => (c.style.width && /%$/.test(c.style.width) ? numStr(parseFloat(c.style.width)) + '%' : null));
  for (const tr of el.querySelectorAll(':scope > thead > tr, :scope > tbody > tr, :scope > tr')) {
    if (tr.hasAttribute('data-mo-deco')) continue;
    const head = tagOf(tr.parentNode) === 'thead';
    const row = { header: head, cells: [] };
    for (const td of tr.children) {
      const tag = tagOf(td);
      if (tag !== 'td' && tag !== 'th') continue;
      const blocks = readBlocks(td);
      row.cells.push({ header: tag === 'th' || head, colspan: span(td.getAttribute('colspan')), rowspan: span(td.getAttribute('rowspan')), css: cellCss(td), blocks: blocks.length ? blocks.filter(b => b.type !== 'break') : [{ type: 'p', css: {}, runs: [] }] });
    }
    if (row.cells.length) t.rows.push(row);
  }
  return t;
}

// ---- rendering --------------------------------------------------------------------------------------------------------------------

/** A block as DOM of the editor. options.fieldText(kind) gives fields a visible value, options.noteNumber(id) the number of a note. */
export function renderBlock(b, options = {}) {
  if (b.type === 'hr') return el('hr', { contenteditable: 'false' });
  if (b.type === 'break') {
    const d = el('div', { 'data-break': b.kind === 'column' ? 'column' : 'page', contenteditable: 'false', class: 'mo-docs-break' });
    if (options.breakLabel) d.setAttribute('data-mo-label', options.breakLabel(b.kind));
    return d;
  }
  if (b.type === 'table') return renderTable(b, options);
  const node = el(TEXT_TAGS.has(b.type) ? b.type : 'p', {});
  if (b.id) node.id = b.id;
  if (b.style) node.setAttribute('data-style', b.style);
  if (b.list) {
    node.setAttribute('data-list', b.list);
    if (b.level > 0) node.setAttribute('data-level', String(b.level));
    if (b.start) node.setAttribute('data-start', String(b.start));
  }
  applyBlockCss(node, b.css);
  if (b.list && b.start) node.style.counterSet = counterSet(b.level || 0, b.start);
  renderRuns(node, b.runs, options);
  if (!b.runs.length || b.runs[b.runs.length - 1].obj === 'br' || (b.type === 'pre' && /\n$/.test(b.runs[b.runs.length - 1].text || ''))) node.append(el('br', {}));
  return node;
}

/** The counters of a list item that starts its level at n (the deeper levels start anew). */
export function counterSet(level, n) {
  let s = 'l' + level + ' ' + n;
  for (let k = level + 1; k < 9; k++) s += ' l' + k + ' 0';
  return s;
}

/** The paragraph CSS of the model on an element: space before (and its padding) go through custom properties, so they add up like in Word. */
export function applyBlockCss(node, css) {
  node.removeAttribute('style');
  for (const [k, v] of Object.entries(css || {})) {
    if (k === 'margin-top') node.style.setProperty('--mo-sb', v);
    else if (k === 'padding-top') node.style.setProperty('--mo-pt', v);
    else node.style.setProperty(k, v);
  }
  if (!node.getAttribute('style')) node.removeAttribute('style');
}

function el(tag, attrs) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, v);
  return n;
}

/** Appends runs to a parent in the canonical nesting (neighbours share their wrappers). */
export function renderRuns(parent, runs, options = {}) {
  const stack = [{ key: null, node: parent }];
  for (const r of runs) {
    const chain = wrappers(r.f || {}, !!r.obj && r.obj !== 'br');
    let common = 0;
    while (common < stack.length - 1 && common < chain.length && stack[common + 1].key === chain[common].key) common++;
    stack.length = common + 1;
    for (let i = common; i < chain.length; i++) {
      const w = chain[i].make();
      stack[stack.length - 1].node.append(w);
      stack.push({ key: chain[i].key, node: w });
    }
    const into = stack[stack.length - 1].node;
    if (!r.obj) into.append(document.createTextNode(r.text));
    else into.append(renderObject(r, options));
  }
}

function renderObject(r, options) {
  switch (r.obj) {
    case 'br': return el('br', {});
    case 'img': {
      const img = el('img', { src: r.src, alt: r.alt || '', draggable: 'false' });
      if (r.width) img.style.width = numStr(r.width) + 'px';
      if (r.height) img.style.height = numStr(r.height) + 'px';
      if (r.align) img.setAttribute('data-align', r.align);
      return img;
    }
    case 'field': {
      const s = el('span', { 'data-field': r.kind, contenteditable: 'false' });
      s.setAttribute('data-mo-text', options.fieldText ? options.fieldText(r.kind) : '#');
      return s;
    }
    case 'note': {
      const s = el('sup', { 'data-fn': r.id, contenteditable: 'false' });
      s.setAttribute('data-mo-n', options.noteNumber ? String(options.noteNumber(r.id) ?? '') : '');
      return s;
    }
    default: return document.createTextNode('');
  }
}

function wrappers(f, object) {
  const w = [];
  if (f.href) w.push({ key: 'a:' + f.href, make: () => el('a', { href: f.href }) });
  if (f.cmt) w.push({ key: 'c:' + f.cmt, make: () => el('span', { 'data-comment': f.cmt }) });
  const style = [];
  if (f.font) style.push(['font-family', f.font]);
  if (f.size) style.push(['font-size', f.size]);
  if (f.color) style.push(['color', f.color]);
  if (f.bg) style.push(['background-color', f.bg]);
  if (f.ls) style.push(['letter-spacing', f.ls]);
  if (style.length) {
    const key = 's:' + style.map(s => s.join(':')).join(';');
    w.push({ key, make: () => { const s = el('span', {}); for (const [k, v] of style) s.style.setProperty(k, v); return s; } });
  }
  if (!object) {
    if (f.b) w.push({ key: 'b', make: () => el('strong', {}) });
    if (f.i) w.push({ key: 'i', make: () => el('em', {}) });
    if (f.u) w.push({ key: 'u', make: () => el('u', {}) });
    if (f.s) w.push({ key: 's', make: () => el('s', {}) });
    if (f.sup) w.push({ key: 'sup', make: () => el('sup', {}) });
    else if (f.sub) w.push({ key: 'sub', make: () => el('sub', {}) });
  }
  return w;
}

function renderTable(t, options) {
  const table = el('table', { 'data-table-style': t.tableStyle || 'grid' });
  if (t.width) table.style.width = t.width;
  if ((t.cols || []).length) {
    const cg = el('colgroup', {});
    for (const c of t.cols) { const col = el('col', {}); if (c) col.style.width = c; cg.append(col); }
    table.append(cg);
  }
  const head = [];
  for (const r of t.rows) { if (r.header) head.push(r); else break; }
  const section = (rows, tag, header) => {
    const s = el(tag, {});
    for (const r of rows) {
      const tr = el('tr', {});
      for (const c of r.cells) {
        const td = el(header || c.header ? 'th' : 'td', {});
        if (c.colspan > 1) td.setAttribute('colspan', String(c.colspan));
        if (c.rowspan > 1) td.setAttribute('rowspan', String(c.rowspan));
        for (const [k, v] of Object.entries(c.css || {})) td.style.setProperty(k, v);
        const blocks = c.blocks.length ? c.blocks : [{ type: 'p', css: {}, runs: [] }];
        for (const b of blocks) td.append(renderBlock(b, options));
        tr.append(td);
      }
      s.append(tr);
    }
    return s;
  };
  if (head.length) table.append(section(head, 'thead', true));
  if (t.rows.length > head.length) table.append(section(t.rows.slice(head.length), 'tbody', false));
  return table;
}

export function renderBlocks(blocks, options = {}) {
  const frag = document.createDocumentFragment();
  for (const b of blocks) frag.append(renderBlock(b, options));
  return frag;
}

// ---- positions --------------------------------------------------------------------------------------------------------------------

export const isTextBlockEl = n => n && n.nodeType === 1 && TEXT_TAGS.has(tagOf(n));

/** The character offset of a DOM position inside a text block. */
export function offsetOf(block, node, offset) {
  if (!block.contains(node) && node !== block) return node.compareDocumentPosition(block) & Node.DOCUMENT_POSITION_FOLLOWING ? 0 : domLength(block);
  // a position inside an atomic element counts as before (or after) it
  for (let a = node.nodeType === 1 ? node : node.parentNode; a && a !== block; a = a.parentNode) {
    if (isAtomic(a) || isDecoration(a)) { const idx = indexOf(a); return offsetOf(block, a.parentNode, idx + (offset > 0 ? 1 : 0)); }
  }
  let count = 0;
  let done = false;
  const target = node.nodeType === 3 ? node : node.childNodes[offset] || null;
  const visit = n => {
    for (let c = n.firstChild; c && !done; c = c.nextSibling) {
      if (c === target) { if (node.nodeType === 3) count += Math.min(offset, c.data.length); done = true; return; }
      if (c.nodeType === 3) count += c.data.length;
      else if (c.nodeType === 1) {
        if (isDecoration(c)) continue;
        if (tagOf(c) === 'br') { if (!isPlaceholderBr(block, c)) count += 1; }
        else if (isAtomic(c)) count += 1;
        else visit(c);
      }
    }
    if (!done && n === node && node.nodeType === 1 && offset >= node.childNodes.length) done = true;
  };
  visit(block);
  return count;
}

const indexOf = n => Array.prototype.indexOf.call(n.parentNode.childNodes, n);

/** The length of a text block in the DOM (the same as the model's). */
export function domLength(block) {
  let count = 0;
  const visit = n => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) count += c.data.length;
      else if (c.nodeType === 1 && !isDecoration(c)) {
        if (tagOf(c) === 'br') { if (!isPlaceholderBr(block, c)) count += 1; }
        else if (isAtomic(c)) count += 1;
        else visit(c);
      }
    }
  };
  visit(block);
  return count;
}

/**
 * The DOM position of a character offset in a text block. At the border between two pieces the end of the piece before wins
 * (typing continues its format) unless after is asked for.
 */
export function pointAt(block, offset, after = false) {
  let remaining = Math.max(0, offset);
  let result = null;
  let lastText = null;
  const visit = n => {
    for (let c = n.firstChild; c && !result; c = c.nextSibling) {
      if (c.nodeType === 3) {
        const len = c.data.length;
        if (remaining < len || (remaining === len && !after)) { result = { node: c, offset: remaining }; return; }
        remaining -= len;
        lastText = c;
      } else if (c.nodeType === 1 && !isDecoration(c)) {
        const atomic = isAtomic(c);
        if (atomic) {
          if (tagOf(c) === 'br' && isPlaceholderBr(block, c)) { if (remaining === 0) { result = { node: c.parentNode, offset: indexOf(c) }; return; } continue; }
          if (remaining === 0) { result = { node: c.parentNode, offset: indexOf(c) }; return; }
          remaining -= 1;
          if (remaining === 0 && !after) {
            // just after the object: in the text that follows it, or after it
            const next = c.nextSibling;
            result = next && next.nodeType === 3 ? { node: next, offset: 0 } : { node: c.parentNode, offset: indexOf(c) + 1 };
            return;
          }
        } else visit(c);
      }
    }
  };
  visit(block);
  if (result) return result;
  if (lastText) return { node: lastText, offset: lastText.data.length };
  // past the end (or an empty block): before the placeholder <br>
  const br = block.lastChild && block.lastChild.nodeType === 1 && tagOf(block.lastChild) === 'br' ? block.lastChild : null;
  return br ? { node: block, offset: indexOf(br) } : { node: block, offset: block.childNodes.length };
}

export { runsLength };
