/* MatOffice Slides – draws a slide as DOM (HTML + SVG) in slide pixels: background, the decoration of the design, the elements, the footers.
   One renderer for everything: the canvas of the editor, the thumbnails, the slide show and the presenter view, printing and the PNG pictures
   (the same DOM drawn into a canvas through an SVG foreignObject). Text is written with textContent only; colours and fonts are checked values.
   The few rules that cannot be inline (bullets) are RENDER_CSS, put into the page once and into every SVG picture. */
import { t, lang } from '../editor/util.js';
import {
  PT, INSET, resolveColor, alpha, backgroundOf, decorOf, textDefaults, indentOf, numberText, isNumbering, isEmptyBody, slotOf,
  textColor, tableLook, FOOTER,
} from './model.js';
import { shapePath, textRect, arrowHead, arrowInset, dashArray, adjOf } from './geometry.js';

export const RENDER_CSS = `
.mo-sr{position:relative;overflow:hidden;box-sizing:border-box;color:#000;text-align:left;line-height:normal;font-kerning:normal;-webkit-font-smoothing:antialiased;letter-spacing:normal;word-spacing:normal;text-transform:none;font-family:Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:16px;}
.mo-sr *,.mo-sr *::before{box-sizing:border-box;}
.mo-sr-el{position:absolute;transform-origin:50% 50%;}
.mo-sr-geom{position:absolute;left:0;top:0;overflow:visible;}
.mo-sr-text{position:absolute;display:flex;flex-direction:column;}
.mo-sr-body{flex:none;outline:none;min-height:1px;}
.mo-sr-p{position:relative;margin:0;white-space:pre-wrap;overflow-wrap:break-word;tab-size:4;}
.mo-sr-nowrap .mo-sr-p{white-space:pre;}
.mo-sr-p[data-bullet]::before{content:attr(data-bullet);position:absolute;left:var(--bx,0px);top:0;font-size:var(--bs);color:var(--bc);font-family:var(--bf);font-weight:var(--bw,400);font-style:normal;text-decoration:none;white-space:pre;}
.mo-sr-table{border-collapse:collapse;table-layout:fixed;}
.mo-sr-table td{padding:0;}
.mo-sr-img{position:absolute;max-width:none;max-height:none;user-select:none;-webkit-user-drag:none;}
.mo-sr-footer{position:absolute;display:flex;align-items:center;white-space:nowrap;overflow:hidden;padding:0 9.6px;}
`;

/** Puts RENDER_CSS into a document once. */
export function ensureStyles(doc = document) {
  if (doc.getElementById('mo-slides-render-css')) return;
  const style = doc.createElement('style');
  style.id = 'mo-slides-render-css';
  style.textContent = RENDER_CSS;
  doc.head.append(style);
}

const SERIF = ['Cambria', 'Georgia', 'Constantia', 'Palatino Linotype', 'Times New Roman', 'Garamond', 'Book Antiqua', 'Century', 'Rockwell', 'Baskerville', 'Bodoni', 'Didot'];
const FALLBACK = {
  'Calibri': 'Carlito, "Segoe UI", Arial, sans-serif',
  'Calibri Light': '"Calibri", Carlito, "Segoe UI Light", "Segoe UI", Arial, sans-serif',
  'Segoe UI Light': '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
  'Segoe UI Semibold': '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
  'Cambria': 'Caladea, Georgia, serif',
  'Corbel': '"Segoe UI", "Trebuchet MS", Arial, sans-serif',
  'Candara': '"Segoe UI", "Trebuchet MS", Arial, sans-serif',
  'Constantia': 'Georgia, "Times New Roman", serif',
  'Consolas': '"Cascadia Mono", Menlo, "Courier New", monospace',
  'Courier New': 'Courier, monospace',
  'Arial Black': '"Arial", "Helvetica Neue", sans-serif',
  'Franklin Gothic Medium': '"Arial Narrow", Arial, sans-serif',
};

/** A CSS font stack for a font name (the name is a checked value; quotes are taken out anyway). */
export function fontStack(name) {
  const clean = String(name || 'Calibri').replace(/["\\;{}<>]/g, '');
  const fallback = FALLBACK[clean] || (SERIF.some(s => clean.startsWith(s)) ? 'Georgia, "Times New Roman", serif' : '"Segoe UI", "Helvetica Neue", Arial, sans-serif');
  return `"${clean}", ${fallback}`;
}

const SVG = 'http://www.w3.org/2000/svg';

function mk(doc, tag, cls, style) {
  const n = doc.createElement(tag);
  if (cls) n.className = cls;
  if (style) Object.assign(n.style, style);
  return n;
}

function svg(doc, tag, attrs) {
  const n = doc.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs || {})) if (v !== null && v !== undefined) n.setAttribute(k, String(v));
  return n;
}

const px = v => Math.round(v * 100) / 100 + 'px';

/** A fill as a CSS background value. */
export function cssBackground(fill, theme) {
  if (!fill || fill.type === 'none') return 'transparent';
  if (fill.type === 'solid') return resolveColor(fill.color, theme) || 'transparent';
  if (fill.type === 'gradient') return `linear-gradient(${fill.angle ?? 90}deg, ${resolveColor(fill.from, theme)}, ${resolveColor(fill.to, theme)})`;
  if (fill.type === 'image' && /^data:image\/(png|jpeg|gif|webp);base64,/.test(fill.src)) return `url("${fill.src}") center / 100% 100% no-repeat`;
  return 'transparent';
}

/**
 * A slide as DOM, w × h slide pixels. opts: { doc, mode: 'edit'|'view', number (1-based), editing (id of the element whose text is edited:
 * no prompt), hide (Set of ids drawn invisible: animations not played yet) }.
 */
export function renderSlide(content, slide, opts = {}) {
  const doc = opts.doc || document;
  const { w, h } = content.size;
  const root = mk(doc, 'div', 'mo-sr', { width: px(w), height: px(h), background: cssBackground(backgroundOf(content, slide), content.theme) });
  root.dataset.slide = slide.id;
  for (const d of decorOf(content, slide)) root.append(renderDecor(doc, content, d));
  for (const e of slide.elements) {
    const node = renderElement(content, slide, e, opts);
    if (node) root.append(node);
  }
  renderFooters(doc, root, content, slide, opts);
  return root;
}

function renderDecor(doc, content, d) {
  const { w, h } = content.size;
  const x = d.x * w, y = d.y * h, dw = d.w * w, dh = d.h * h;
  const node = mk(doc, 'div', 'mo-sr-el mo-sr-decor', { left: px(x), top: px(y), width: px(dw), height: px(dh), pointerEvents: 'none' });
  if (d.rot) node.style.transform = `rotate(${d.rot}deg)`;
  if (d.src) {
    const img = mk(doc, 'img', 'mo-sr-img', { left: 0, top: 0, width: '100%', height: '100%', opacity: d.opacity ?? 1 });
    img.src = d.src;
    img.alt = '';
    node.append(img);
    return node;
  }
  const g = svg(doc, 'svg', { class: 'mo-sr-geom', width: dw, height: dh, viewBox: `0 0 ${dw} ${dh}`, 'aria-hidden': 'true' });
  if (d.flipH || d.flipV) g.style.transform = `scale(${d.flipH ? -1 : 1}, ${d.flipV ? -1 : 1})`;
  g.append(svg(doc, 'path', { d: shapePath(d.shape || 'rect', dw, dh), fill: resolveColor(d.fill || 'accent1', content.theme), 'fill-opacity': d.opacity ?? 1 }));
  node.append(g);
  return node;
}

/** One element as DOM (absolutely placed in the slide). */
export function renderElement(content, slide, e, optsIn = {}) {
  const opts = optsIn.editing === e.id ? { ...optsIn, editable: true } : optsIn;
  const doc = opts.doc || document;
  const node = mk(doc, 'div', 'mo-sr-el', { left: px(e.x), top: px(e.y), width: px(Math.max(0, e.w)), height: px(Math.max(0, e.h)) });
  node.dataset.id = e.id;
  node.dataset.type = e.type;
  if (e.rot) node.style.transform = `rotate(${e.rot}deg)`;
  if (opts.hide?.has(e.id)) node.style.visibility = 'hidden';
  switch (e.type) {
    case 'image': renderImage(doc, node, content, e); break;
    case 'line': renderLine(doc, node, content, e, opts); break;
    case 'table': renderTable(doc, node, content, slide, e, opts); break;
    default: renderShape(doc, node, content, slide, e, opts); break;
  }
  return node;
}

function shadowFilter() { return 'drop-shadow(3.8px 3.8px 4px rgba(0, 0, 0, 0.4))'; }

function fillPaint(doc, defs, content, fill, key, w, h) {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'solid') return resolveColor(fill.color, content.theme);
  if (fill.type === 'gradient') {
    const id = 'mo-g-' + key;
    const a = (((fill.angle ?? 90) - 90) * Math.PI) / 180;
    const dx = Math.cos(a) / 2, dy = Math.sin(a) / 2;
    const g = svg(doc, 'linearGradient', { id, x1: 0.5 - dx, y1: 0.5 - dy, x2: 0.5 + dx, y2: 0.5 + dy });
    g.append(svg(doc, 'stop', { offset: 0, 'stop-color': resolveColor(fill.from, content.theme) }), svg(doc, 'stop', { offset: 1, 'stop-color': resolveColor(fill.to, content.theme) }));
    defs.append(g);
    return `url(#${id})`;
  }
  if (fill.type === 'image') {
    const id = 'mo-p-' + key;
    const p = svg(doc, 'pattern', { id, patternUnits: 'userSpaceOnUse', width: w, height: h });
    p.append(svg(doc, 'image', { href: fill.src, width: w, height: h, preserveAspectRatio: 'xMidYMid slice' }));
    defs.append(p);
    return `url(#${id})`;
  }
  return 'none';
}

function renderShape(doc, node, content, slide, e, opts) {
  const theme = content.theme;
  const shape = e.type === 'shape' ? e.shape || 'rect' : 'rect';
  const slot = slotOf(slide, e.ph);
  const hasLook = (e.fill && e.fill.type !== 'none') || e.line;
  if (hasLook) {
    const g = svg(doc, 'svg', { class: 'mo-sr-geom', width: e.w, height: e.h, viewBox: `0 0 ${Math.max(e.w, 0.01)} ${Math.max(e.h, 0.01)}`, 'aria-hidden': 'true' });
    const defs = svg(doc, 'defs');
    g.append(defs);
    const key = (slide.id + '-' + e.id).replace(/[^A-Za-z0-9_-]/g, '');
    const path = svg(doc, 'path', { d: shapePath(shape, e.w, e.h, e.adj), fill: fillPaint(doc, defs, content, e.fill, key, e.w, e.h), 'fill-rule': 'evenodd' });
    if (e.opacity !== undefined && e.opacity < 1) path.setAttribute('fill-opacity', e.opacity);
    if (e.line) {
      const lw = e.line.width * PT;
      path.setAttribute('stroke', resolveColor(e.line.color, theme));
      path.setAttribute('stroke-width', lw);
      path.setAttribute('stroke-linejoin', 'round');
      const dash = dashArray(e.line.dash, lw);
      if (dash) path.setAttribute('stroke-dasharray', dash);
    }
    g.append(path);
    if (e.flipH || e.flipV) g.style.transform = `scale(${e.flipH ? -1 : 1}, ${e.flipV ? -1 : 1})`;
    if (e.shadow) g.style.filter = shadowFilter();
    node.append(g);
  } else if (e.shadow && e.type === 'text') {
    node.style.filter = shadowFilter();
  }

  // a picture placeholder that is still empty
  if (slot?.role === 'picture' && opts.mode === 'edit' && isEmptyBody(e.text)) {
    node.classList.add('mo-sr-ph-empty');
    node.append(promptBox(doc, t('Slides.Prompt.Picture'), ['picture'], content, slide, e));
    return;
  }

  if (!e.text) return;
  const rect = textRect(shape, e.w, e.h);
  const prompt = opts.mode === 'edit' && e.ph && opts.editing !== e.id && isEmptyBody(e.text);
  if (prompt) {
    node.classList.add('mo-sr-ph-empty');
    const role = slot?.role || (e.ph === 'title' ? 'title' : 'body');
    const text = role === 'title' ? t('Slides.Prompt.Title') : e.ph === 'subtitle' ? t('Slides.Prompt.Subtitle') : t('Slides.Prompt.Text');
    const body = { ...e.text, paras: [{ ...(e.text.paras[0] || {}), runs: [{ ...(e.text.paras[0]?.runs[0] || {}), t: text }] }] };
    const box = textBox(doc, content, slide, e, body, rect, opts);
    box.classList.add('mo-sr-prompt');
    node.append(box);
    if (slot?.content) node.append(insertIcons(doc));
    return;
  }
  node.append(textBox(doc, content, slide, e, e.text, rect, opts));
}

function promptBox(doc, text, icons) {
  const box = mk(doc, 'div', 'mo-sr-pic-prompt');
  box.append(insertIcons(doc, icons), mk(doc, 'span', 'mo-sr-pic-prompt__text'));
  box.lastChild.textContent = text;
  return box;
}

/** The buttons of an empty content placeholder (insert a table, a picture), like the office suites. */
function insertIcons(doc, which = ['table', 'picture']) {
  const box = mk(doc, 'div', 'mo-sr-insert');
  for (const kind of which) {
    const b = mk(doc, 'button', 'mo-sr-insert__btn');
    b.type = 'button';
    b.dataset.insert = kind;
    b.setAttribute('aria-label', t(kind === 'table' ? 'Slides.InsertTable' : 'Slides.InsertPicture'));
    b.dataset.tip = b.getAttribute('aria-label');
    const ic = doc.createElementNS(SVG, 'svg');
    ic.setAttribute('class', 'mat-icon');
    ic.setAttribute('aria-hidden', 'true');
    const use = doc.createElementNS(SVG, 'use');
    use.setAttribute('href', kind === 'table' ? '#mat-i-slides-table' : '#mat-i-slides-picture');
    ic.append(use);
    b.append(ic);
    box.append(b);
  }
  return box;
}

/** The text area of a shape or text box: insets, vertical anchor, the paragraphs. */
export function textBox(doc, content, slide, e, body, rect, opts = {}, cell = null) {
  const d = textDefaults(content, slide, e, 0, cell);
  const inset = body.inset || INSET;
  const anchor = body.anchor || d.anchor;
  const box = mk(doc, 'div', 'mo-sr-text', {
    left: px(rect.x), top: px(rect.y), width: px(Math.max(rect.w, 0)), height: px(Math.max(rect.h, 0)),
    padding: `${px(inset[1])} ${px(inset[2])} ${px(inset[3])} ${px(inset[0])}`,
    justifyContent: anchor === 'middle' ? 'center' : anchor === 'bottom' ? 'flex-end' : 'flex-start',
  });
  if (body.wrap === false) box.classList.add('mo-sr-nowrap');
  box.dataset.fit = body.autofit || d.autofit;
  box.append(renderBody(doc, content, slide, e, body, opts, cell));
  return box;
}

/** The paragraphs of a text body (used again by the text editor, with data-r on every run so it can read them back). */
export function renderBody(doc, content, slide, e, body, opts = {}, cell = null) {
  const theme = content.theme;
  const wrap = mk(doc, 'div', 'mo-sr-body');
  const counters = [];
  body.paras.forEach((p, i) => {
    const level = p.level || 0;
    const d = textDefaults(content, slide, e, level, cell);
    const bullet = p.bullet ?? d.bullet;
    const first = p.runs.find(r => r.t) || p.runs[0] || {};
    const size = (first.size || d.size) * PT;
    const para = mk(doc, 'div', 'mo-sr-p');
    if (opts.editable) { const { runs: _, ...pp } = p; para.dataset.p = JSON.stringify(pp); }
    para.style.textAlign = p.align || d.align;
    para.style.lineHeight = String(1.2 * (p.lineSpacing ?? d.lineSpacing));
    para.style.fontSize = `calc(var(--fs, 1) * ${px(size)})`;
    para.style.fontFamily = fontStack(first.font || d.font);
    const before = i === 0 ? 0 : p.spaceBefore ?? d.spaceBefore;
    if (before) para.style.marginTop = `calc(var(--fs, 1) * ${px(before * PT)})`;
    if (p.spaceAfter) para.style.marginBottom = `calc(var(--fs, 1) * ${px(p.spaceAfter * PT)})`;
    for (let l = level + 1; l < counters.length; l++) counters[l] = 0;
    let mark = null;
    if (bullet && bullet !== 'none') {
      if (isNumbering(bullet)) { counters[level] = (counters[level] || 0) + 1; mark = numberText(bullet, counters[level]); }
      else { counters[level] = 0; mark = bullet; }
    } else counters[level] = 0;
    const ind = indentOf(bullet, level);
    para.style.paddingLeft = px(ind.left);
    if (mark && !isEmptyLine(p, opts)) {
      para.dataset.bullet = mark;
      para.style.setProperty('--bx', px(ind.left - ind.hang));
      para.style.setProperty('--bs', `calc(var(--fs, 1) * ${px(size)})`);
      para.style.setProperty('--bc', resolveColor(first.color || d.color, theme));
      para.style.setProperty('--bf', isNumbering(bullet) ? fontStack(first.font || d.font) : 'Arial, "Segoe UI Symbol", sans-serif');
      if ((first.b ?? d.bold) && isNumbering(bullet)) para.style.setProperty('--bw', '700');
    }
    for (const r of p.runs) {
      if (!r.t && p.runs.length > 1) continue;
      const span = mk(doc, 'span');
      if (opts.editable) {
        const { t: _, ...props } = r;
        span.dataset.r = JSON.stringify(props);
      }
      styleRun(span, r, d, theme);
      const parts = r.t.split('\n');
      parts.forEach((part, k) => {
        if (k > 0) span.append(doc.createElement('br'));
        if (part) span.append(doc.createTextNode(part));
      });
      para.append(span);
    }
    const text = p.runs.map(r => r.t).join('');
    if (!text || text.endsWith('\n')) para.append(doc.createElement('br'));
    wrap.append(para);
  });
  return wrap;
}

/** An empty paragraph shows its bullet only while it is edited (like the office suites). */
function isEmptyLine(p, opts) { return !opts.editable && p.runs.every(r => !r.t); }

/** The look of a run as inline styles. */
export function styleRun(span, r, d, theme) {
  span.style.fontFamily = fontStack(r.font || d.font);
  span.style.fontSize = `calc(var(--fs, 1) * ${px((r.size || d.size) * PT)})`;
  span.style.fontWeight = (r.b ?? d.bold) ? '700' : '400';
  span.style.fontStyle = r.i ? 'italic' : 'normal';
  const deco = [r.u ? 'underline' : '', r.s ? 'line-through' : ''].filter(Boolean).join(' ');
  span.style.textDecoration = deco || 'none';
  span.style.color = resolveColor(r.color || d.color, theme);
  if (r.hl) span.style.backgroundColor = resolveColor(r.hl, theme);
}

function renderImage(doc, node, content, e) {
  node.style.overflow = 'hidden';
  const c = e.crop || { l: 0, t: 0, r: 0, b: 0 };
  const iw = e.w / Math.max(0.02, 1 - c.l - c.r), ih = e.h / Math.max(0.02, 1 - c.t - c.b);
  const img = mk(doc, 'img', 'mo-sr-img', { left: px(-c.l * iw), top: px(-c.t * ih), width: px(iw), height: px(ih) });
  img.src = e.src;
  img.alt = e.alt || '';
  img.draggable = false;
  if (e.opacity !== undefined && e.opacity < 1) img.style.opacity = e.opacity;
  if (e.flipH || e.flipV) img.style.transform = `scale(${e.flipH ? -1 : 1}, ${e.flipV ? -1 : 1})`;
  node.append(img);
  if (e.line) {
    const lw = e.line.width * PT;
    node.style.outline = `${px(lw)} ${e.line.dash === 'dot' ? 'dotted' : e.line.dash && e.line.dash !== 'solid' ? 'dashed' : 'solid'} ${resolveColor(e.line.color, content.theme)}`;
    node.style.outlineOffset = px(-lw / 2);
  }
  if (e.shadow) node.style.filter = shadowFilter();
}

function renderLine(doc, node, content, e, opts) {
  const lw = (e.line?.width || 1) * PT;
  const color = resolveColor(e.line?.color || 'dk1', content.theme);
  const g = svg(doc, 'svg', { class: 'mo-sr-geom', width: Math.max(e.w, 1), height: Math.max(e.h, 1), 'aria-hidden': 'true' });
  g.style.overflow = 'visible';
  let x1 = e.flipH ? e.w : 0, y1 = e.flipV ? e.h : 0, x2 = e.flipH ? 0 : e.w, y2 = e.flipV ? 0 : e.h;
  const angle = Math.atan2(y2 - y1, x2 - x1);
  const len = Math.hypot(x2 - x1, y2 - y1);
  const cut = (kind, fromEnd) => {
    const d = Math.min(arrowInset(kind, lw), len / 2);
    if (!d) return;
    if (fromEnd) { x2 -= Math.cos(angle) * d; y2 -= Math.sin(angle) * d; } else { x1 += Math.cos(angle) * d; y1 += Math.sin(angle) * d; }
  };
  const ends = [[e.arrowStart, e.flipH ? e.w : 0, e.flipV ? e.h : 0, angle + Math.PI], [e.arrowEnd, e.flipH ? 0 : e.w, e.flipV ? 0 : e.h, angle]];
  cut(e.arrowStart, false);
  cut(e.arrowEnd, true);
  if (opts.mode === 'edit') g.append(svg(doc, 'line', { x1: 0 + (e.flipH ? e.w : 0), y1: e.flipV ? e.h : 0, x2: e.flipH ? 0 : e.w, y2: e.flipV ? 0 : e.h, stroke: 'transparent', 'stroke-width': Math.max(12, lw + 8), class: 'mo-sr-hit' }));
  const line = svg(doc, 'line', { x1, y1, x2, y2, stroke: color, 'stroke-width': lw, 'stroke-linecap': e.line?.dash === 'dot' ? 'round' : 'butt' });
  const dash = dashArray(e.line?.dash, lw);
  if (dash) line.setAttribute('stroke-dasharray', dash);
  g.append(line);
  for (const [kind, x, y, a] of ends) {
    const head = arrowHead(kind, x, y, a, lw);
    if (!head) continue;
    if (head.type === 'circle') g.append(svg(doc, 'circle', { cx: head.cx, cy: head.cy, r: head.r, fill: color }));
    else if (head.type === 'polyline') g.append(svg(doc, 'polyline', { points: head.points.map(p => p.join(',')).join(' '), fill: 'none', stroke: color, 'stroke-width': lw, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
    else g.append(svg(doc, 'polygon', { points: head.points.map(p => p.join(',')).join(' '), fill: color }));
  }
  if (e.shadow) g.style.filter = shadowFilter();
  node.append(g);
}

function renderTable(doc, node, content, slide, e, opts) {
  const theme = content.theme;
  const table = mk(doc, 'table', 'mo-sr-table', { width: px(e.cols.reduce((a, b) => a + b, 0)) });
  const cols = mk(doc, 'colgroup');
  for (const w of e.cols) cols.append(mk(doc, 'col', null, { width: px(w) }));
  table.append(cols);
  const tbody = mk(doc, 'tbody');
  e.rows.forEach((row, r) => {
    const tr = mk(doc, 'tr', null, { height: px(row.h) });
    row.cells.forEach((cell, c) => {
      const look = tableLook(e, r, c);
      const td = mk(doc, 'td');
      td.dataset.cell = r + ',' + c;
      const fill = cell.fill ? cssBackground(cell.fill, theme) : look.fill ? resolveColor(look.fill, theme) : 'transparent';
      td.style.background = fill;
      td.style.border = look.border ? `1px solid ${resolveColor(look.border, theme)}` : '1px solid transparent';
      const body = cell.text || { paras: [{ runs: [{ t: '' }] }] };
      const inset = body.inset || INSET;
      td.style.padding = `${px(inset[1])} ${px(inset[2])} ${px(inset[3])} ${px(inset[0])}`;
      const anchor = body.anchor || 'middle';
      td.style.verticalAlign = anchor === 'top' ? 'top' : anchor === 'bottom' ? 'bottom' : 'middle';
      td.append(renderBody(doc, content, slide, e, body, opts, { text: look.text, bold: look.bold }));
      tr.append(td);
    });
    tbody.append(tr);
  });
  table.append(tbody);
  node.append(table);
}

function renderFooters(doc, root, content, slide, opts) {
  const f = content.footer;
  if (!f || (slide.layout === 'title' && f.skipTitle !== false)) return;
  const items = [];
  if (f.date) items.push(['date', f.dateText || new Date().toLocaleDateString(lang())]);
  if (f.text) items.push(['text', f.text]);
  if (f.number && opts.number) items.push(['number', String(opts.number)]);
  if (!items.length) return;
  const { w, h } = content.size;
  const color = resolveColor(textColor(content, slide) + ':+40', content.theme);
  for (const [kind, text] of items) {
    const s = FOOTER[kind];
    const n = mk(doc, 'div', 'mo-sr-footer', {
      left: px(s.x * w), top: px(s.y * h), width: px(s.w * w), height: px(s.h * h), color, fontSize: px(FOOTER.size * PT),
      fontFamily: fontStack(content.theme.fonts.minor), justifyContent: s.align === 'center' ? 'center' : s.align === 'right' ? 'flex-end' : 'flex-start',
    });
    n.textContent = text;
    root.append(n);
  }
}

/** Shrinks the text of boxes that are set to "shrink on overflow" until it fits (call when the slide is in the document). */
export function fitText(root) {
  for (const box of root.querySelectorAll('.mo-sr-text[data-fit="shrink"]')) fitBox(box);
}

export function fitBox(box) {
  const body = box.querySelector('.mo-sr-body');
  if (!body) return;
  body.style.removeProperty('--fs');
  const room = box.clientHeight - parseFloat(box.style.paddingTop || 0) - parseFloat(box.style.paddingBottom || 0);
  if (room <= 0 || body.scrollHeight <= room + 0.5) return;
  let lo = 0.25, hi = 1;
  for (let i = 0; i < 7; i++) {
    const mid = (lo + hi) / 2;
    body.style.setProperty('--fs', String(mid));
    if (body.scrollHeight <= room + 0.5) lo = mid; else hi = mid;
  }
  body.style.setProperty('--fs', String(Math.round(lo * 1000) / 1000));
}

/** The resolved colour (for swatches of the palette). */
export const colorOf = (c, theme) => resolveColor(c, theme);
export { alpha, adjOf };
