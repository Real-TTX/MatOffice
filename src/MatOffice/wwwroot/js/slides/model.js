/* MatOffice Slides – the model: layouts, designs, colours, the defaults of text, new slides and elements. No DOM here.
   The content format is documented in Services/Formats/Slides/SlidesHandler.cs and docs/EDITORS.md ("Slides content"). */
import { LAYOUT_DATA } from './layouts-data.js';
import { THEME_DATA } from './themes-data.js';

export const LAYOUTS = LAYOUT_DATA.layouts;
export const FOOTER = LAYOUT_DATA.footer;
export const BODY = LAYOUT_DATA.body;
export const THEMES = THEME_DATA.themes;
export const VARIANTS = THEME_DATA.variants;
export const PT = 4 / 3; // px per pt at 96 dpi
export const INSET = [9.6, 4.8, 9.6, 4.8];

export const layoutOf = id => LAYOUTS.find(l => l.id === id) || LAYOUTS.find(l => l.id === 'blank');
export const slotOf = (slide, ph) => (ph ? layoutOf(slide.layout).slots.find(s => s.ph === ph) : null);
export const groupOf = slide => layoutOf(slide.layout).group;

let counter = 0;
/** A short id, unique in a document. */
export function newId(prefix = 'e') {
  counter = (counter + 1) % 1296;
  return prefix + Date.now().toString(36).slice(-5) + counter.toString(36).padStart(2, '0') + Math.random().toString(36).slice(2, 4);
}

export const clone = v => (v === undefined ? undefined : structuredClone(v));
export const round = v => Math.round(v * 100) / 100;

// ---- colours ------------------------------------------------------------------------------------------------------------------

function toHsl(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255, g = parseInt(hex.slice(3, 5), 16) / 255, b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function fromHsl(h, s, l) {
  const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  let r, g, b;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
  return '#' + [r, g, b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}

/** Lighter (+n %) or darker (-n %), the way lumMod/lumOff work. */
export function lum(hex, mod) {
  const [h, s, l] = toHsl(hex);
  const nl = mod > 0 ? l * (1 - mod / 100) + mod / 100 : l * (1 + mod / 100);
  return fromHsl(h, s, Math.min(1, Math.max(0, nl)));
}

/** A colour of the content (#rrggbb or "accent1:+40") as #rrggbb in a design. */
export function resolveColor(c, theme) {
  if (!c) return null;
  if (c[0] === '#') return c;
  const [name, mod] = c.split(':');
  const base = theme.colors[name] || '#000000';
  return mod ? lum(base, Number(mod)) : base;
}

/** The theme colours of the palette: the ten slots and their shades, like the office suites. */
export const PALETTE_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
export function paletteColumn(slot) {
  // light slots get darker shades, the others lighter and darker ones
  if (slot === 'lt1') return ['lt1', 'lt1:-5', 'lt1:-15', 'lt1:-25', 'lt1:-35', 'lt1:-50'];
  if (slot === 'dk1') return ['dk1', 'dk1:+50', 'dk1:+35', 'dk1:+25', 'dk1:+15', 'dk1:+5'];
  if (slot === 'lt2') return ['lt2', 'lt2:-10', 'lt2:-25', 'lt2:-50', 'lt2:-75', 'lt2:-90'];
  return [slot, slot + ':+80', slot + ':+60', slot + ':+40', slot + ':-25', slot + ':-50'];
}
export const STANDARD_COLORS = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];

// ---- designs ------------------------------------------------------------------------------------------------------------------

export function themeById(id) { return THEMES.find(t => t.id === id); }

/** A design as it is kept in the content (a copy, with a colour variant). */
export function makeTheme(id, variant = 0) {
  const t = clone(themeById(id) || THEMES[0]);
  if (variant > 0 && VARIANTS[variant - 1]) {
    for (const [k, v] of Object.entries(VARIANTS[variant - 1])) if (k !== 'id') t.colors[k] = v;
    t.variant = variant;
  }
  if (!t.decor) t.decor = [];
  return t;
}

/** The background of a slide: its own, else the group's of the design, else the design's. */
export function backgroundOf(content, slide) {
  return slide.background || content.theme.groups?.[groupOf(slide)]?.background || content.theme.background || { type: 'solid', color: 'lt1' };
}

export function headingColor(content, slide) { return content.theme.groups?.[groupOf(slide)]?.heading || content.theme.heading || 'dk1'; }
export function textColor(content, slide) { return content.theme.groups?.[groupOf(slide)]?.text || content.theme.text || 'dk1'; }

/** The decoration of the design that a slide shows. */
export function decorOf(content, slide) {
  const layout = layoutOf(slide.layout);
  return (content.theme.decor || []).filter(d => (d.on || ['all']).some(o => o === 'all' || o === layout.id || o === layout.group));
}

// ---- text defaults -----------------------------------------------------------------------------------------------------------

/** The role of an element's text: title, body (placeholders), text (text boxes), shape, cell. */
export function roleOf(slide, el) {
  const slot = slotOf(slide, el.ph);
  if (slot) return slot.role === 'picture' ? 'body' : slot.role;
  if (el.ph === 'title') return 'title';
  if (el.ph) return 'body';
  return el.type === 'text' ? 'text' : 'shape';
}

/** What a paragraph/run of an element shows when it does not say it itself. */
export function textDefaults(content, slide, el, level = 0, cell = null) {
  const role = cell ? 'cell' : roleOf(slide, el);
  const slot = slotOf(slide, el.ph);
  const theme = content.theme;
  const d = { role, font: theme.fonts.minor, size: 18, color: textColor(content, slide), bold: false, align: 'left', anchor: 'top', bullet: 'none', spaceBefore: 0, lineSpacing: 1, autofit: 'none' };
  if (role === 'title') {
    Object.assign(d, { font: theme.fonts.major, size: slot?.size || 44, color: headingColor(content, slide), align: slot?.align || 'left', anchor: slot?.anchor || 'middle', lineSpacing: LAYOUT_DATA.titleLineSpacing, autofit: 'shrink' });
  } else if (role === 'body') {
    const size = slot?.size || BODY.sizes[Math.min(level, BODY.sizes.length - 1)];
    const bulleted = slot ? slot.bullet !== false : true;
    Object.assign(d, {
      size, align: slot?.align || 'left', anchor: slot?.anchor || 'top', bold: !!slot?.bold, autofit: 'shrink',
      bullet: bulleted ? (level % 2 === 0 ? '•' : '–') : 'none', spaceBefore: bulleted ? BODY.spaceBefore[Math.min(level, 8)] : 0,
    });
    if (slot?.muted) d.color = d.color + (d.color.includes(':') ? '' : ':+25');
  } else if (role === 'shape') {
    Object.assign(d, { color: 'lt1', align: 'center', anchor: 'middle' });
  } else if (role === 'cell') {
    Object.assign(d, { color: cell.text, bold: cell.bold, anchor: 'middle' });
  } else {
    d.autofit = 'grow';
  }
  return d;
}

/** The indent of a paragraph (px): left margin of the text and the hanging space of the bullet. */
export function indentOf(bullet, level) {
  if (bullet && bullet !== 'none') return { left: BODY.indent + BODY.step * level, hang: BODY.indent };
  return { left: BODY.step * level, hang: 0 };
}

export const NUMBERINGS = ['1.', '1)', 'a.', 'a)', 'A.', 'A)', 'I.', 'i.'];
export const BULLETS = ['•', '◦', '■', '□', '◆', '➢', '✓', '–'];

/** The text of a numbering at position n (1-based). */
export function numberText(format, n) {
  const suffix = format.slice(-1);
  const kind = format.slice(0, -1);
  const roman = v => { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [k, r] of m) while (v >= k) { s += r; v -= k; } return s; };
  const alpha26 = v => { let s = ''; while (v > 0) { v--; s = String.fromCharCode(97 + (v % 26)) + s; v = Math.floor(v / 26); } return s; };
  let text;
  if (kind === '1') text = String(n);
  else if (kind === 'a') text = alpha26(n);
  else if (kind === 'A') text = alpha26(n).toUpperCase();
  else if (kind === 'i') text = roman(n);
  else if (kind === 'I') text = roman(n).toUpperCase();
  else text = String(n);
  return text + suffix;
}

export const isNumbering = b => NUMBERINGS.includes(b);

// ---- text bodies ----------------------------------------------------------------------------------------------------------------

export const emptyBody = (extra = {}) => ({ paras: [{ runs: [{ t: '' }] }], ...extra });

export function bodyText(body) {
  return (body?.paras || []).map(p => p.runs.map(r => r.t).join('')).join('\n');
}

export const isEmptyBody = body => !body || body.paras.every(p => p.runs.every(r => !r.t));

/** A text body from plain text (one paragraph per line; leading tabs = level). */
export function bodyFromText(text, extra = {}) {
  const paras = String(text).split(/\r?\n/).map(line => {
    const level = (line.match(/^\t*/) || [''])[0].length;
    const p = { runs: [{ t: line.slice(level) }] };
    if (level) p.level = Math.min(8, level);
    return p;
  });
  return { paras: paras.length ? paras : [{ runs: [{ t: '' }] }], ...extra };
}

// ---- slides and elements ------------------------------------------------------------------------------------------------------

/** A placeholder element of a layout slot on a slide of the given size. */
export function placeholder(slot, size) {
  const e = { id: newId(), type: slot.role === 'picture' ? 'shape' : 'text', ph: slot.ph, x: round(slot.x * size.w), y: round(slot.y * size.h), w: round(slot.w * size.w), h: round(slot.h * size.h) };
  if (slot.role === 'picture') e.shape = 'rect';
  e.text = emptyBody({ autofit: 'shrink' });
  return e;
}

/** A new slide of a layout with its empty placeholders. */
export function newSlide(content, layoutId) {
  const layout = layoutOf(layoutId);
  return { id: newId('s'), layout: layout.id, elements: layout.slots.map(s => placeholder(s, content.size)) };
}

/** The layout a new slide after this one gets (like the office suites: after the title slide "Title and Content", else the same). */
export function nextLayout(slide) {
  if (!slide || slide.layout === 'title') return 'titleContent';
  if (slide.layout === 'section') return 'titleContent';
  return slide.layout;
}

/** Moves the placeholders of a slide to the slots of another layout; empty ones that have no place go, missing ones come. */
export function applyLayout(content, slide, layoutId) {
  const layout = layoutOf(layoutId);
  const old = layoutOf(slide.layout);
  const free = [...layout.slots];
  const take = pred => { const i = free.findIndex(pred); return i < 0 ? null : free.splice(i, 1)[0]; };
  const kept = [];
  for (const e of slide.elements) {
    if (!e.ph) { kept.push(e); continue; }
    const oldSlot = old.slots.find(s => s.ph === e.ph);
    const role = oldSlot?.role || (e.ph === 'title' ? 'title' : 'body');
    let slot = take(s => s.ph === e.ph) || (role === 'title' ? take(s => s.role === 'title') : null)
      || (e.type === 'image' ? take(s => s.role === 'picture') || take(s => s.content) : null)
      || (role !== 'title' ? take(s => s.role === 'body') : null);
    const empty = e.type !== 'image' && e.type !== 'table' && isEmptyBody(e.text);
    if (!slot) {
      if (!empty) { delete e.ph; kept.push(e); }
      continue;
    }
    e.ph = slot.ph;
    if (e.type !== 'image') Object.assign(e, { x: round(slot.x * content.size.w), y: round(slot.y * content.size.h), w: round(slot.w * content.size.w), h: round(slot.h * content.size.h) });
    else fitInto(e, slot, content.size);
    if (slot.role === 'picture' && e.type === 'text') { e.type = 'shape'; e.shape = 'rect'; }
    if (slot.role !== 'picture' && e.type === 'shape' && e.ph && isEmptyBody(e.text) && !e.fill) e.type = 'text';
    kept.push(e);
  }
  for (const slot of free) kept.push(placeholder(slot, content.size));
  slide.elements = kept;
  slide.layout = layout.id;
}

/** A picture placed into a slot (keeps its proportions, centred). */
export function fitInto(e, slot, size) {
  const bw = slot.w * size.w, bh = slot.h * size.h, bx = slot.x * size.w, by = slot.y * size.h;
  const ratio = e.w && e.h ? e.w / e.h : 1;
  let w = bw, h = bw / ratio;
  if (h > bh) { h = bh; w = bh * ratio; }
  Object.assign(e, { x: round(bx + (bw - w) / 2), y: round(by + (bh - h) / 2), w: round(w), h: round(h) });
}

/** Puts the placeholders of a slide back to their layout (position, size, formatting); missing ones come back. */
export function resetSlide(content, slide) {
  const layout = layoutOf(slide.layout);
  for (const e of slide.elements) {
    const slot = layout.slots.find(s => s.ph === e.ph);
    if (!slot) continue;
    if (e.type === 'image') { fitInto(e, slot, content.size); continue; }
    Object.assign(e, { x: round(slot.x * content.size.w), y: round(slot.y * content.size.h), w: round(slot.w * content.size.w), h: round(slot.h * content.size.h) });
    delete e.rot; delete e.fill; delete e.line; delete e.shadow; delete e.opacity; delete e.flipH; delete e.flipV;
    if (e.text) {
      e.text = { paras: e.text.paras.map(p => ({ ...(p.level ? { level: p.level } : {}), runs: [{ t: p.runs.map(r => r.t).join('') }] })), autofit: 'shrink' };
    }
  }
  for (const slot of layout.slots) if (!slide.elements.some(e => e.ph === slot.ph)) slide.elements.push(placeholder(slot, content.size));
}

/** Every element id of a slide that belongs to the same group(s) as the given ones. */
export function withGroups(slide, ids) {
  const set = new Set(ids);
  const groups = new Set(slide.elements.filter(e => set.has(e.id) && e.group).map(e => e.group));
  if (groups.size) for (const e of slide.elements) if (groups.has(e.group)) set.add(e.id);
  return [...set];
}

/** The box around elements (rotation ignored). */
export function bounds(els) {
  if (!els.length) return null;
  const x = Math.min(...els.map(e => e.x)), y = Math.min(...els.map(e => e.y));
  const r = Math.max(...els.map(e => e.x + e.w)), b = Math.max(...els.map(e => e.y + e.h));
  return { x, y, w: r - x, h: b - y };
}

/** The box around a rotated element. */
export function rotatedBounds(e) {
  const rot = ((e.rot || 0) * Math.PI) / 180;
  if (!rot) return { x: e.x, y: e.y, w: e.w, h: e.h };
  const cx = e.x + e.w / 2, cy = e.y + e.h / 2;
  const c = Math.abs(Math.cos(rot)), s = Math.abs(Math.sin(rot));
  const w = e.w * c + e.h * s, h = e.w * s + e.h * c;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

/** The look of a table cell by its style (the same rules as PptxXml.TableLook on the server). */
export function tableLook(t, row, col) {
  const style = t.tableStyle || 'medium';
  const accent = (t.accent || 'accent1').split(':')[0];
  const rows = t.rows.length;
  const header = t.header !== false, banded = t.banded !== false;
  const isHeader = header && row === 0;
  const isLast = !!t.lastRow && row === rows - 1 && !isHeader;
  const isFirstCol = !!t.firstCol && col === 0 && !isHeader && !isLast;
  const bodyIndex = row - (header ? 1 : 0);
  const odd = banded && bodyIndex % 2 === 0;
  switch (style) {
    case 'none': return { fill: null, text: 'dk1', bold: isHeader || isLast, border: 'dk1:+65' };
    case 'light':
      if (isHeader) return { fill: null, text: accent + ':-25', bold: true, border: accent };
      return { fill: odd ? accent + ':+80' : null, text: 'dk1', bold: isLast || isFirstCol, border: accent + ':+40' };
    case 'dark':
      if (isHeader) return { fill: 'dk1', text: 'lt1', bold: true, border: 'lt1' };
      if (isLast || isFirstCol) return { fill: accent + ':-50', text: 'lt1', bold: true, border: 'lt1' };
      return { fill: odd ? accent + ':-25' : accent, text: 'lt1', bold: false, border: 'lt1' };
    default:
      if (isHeader || isLast || isFirstCol) return { fill: accent, text: 'lt1', bold: true, border: 'lt1' };
      return { fill: odd ? accent + ':+60' : accent + ':+80', text: 'dk1', bold: false, border: 'lt1' };
  }
}

/** A new table of rows × cols in a box. */
export function newTable(rows, cols, x, y, w, rowH = 40) {
  const colW = round(w / cols);
  return {
    id: newId(), type: 'table', x: round(x), y: round(y), w: round(colW * cols), h: round(rowH * rows),
    cols: Array(cols).fill(colW), rows: Array.from({ length: rows }, () => ({ h: rowH, cells: Array.from({ length: cols }, () => ({ text: emptyBody() })) })),
    tableStyle: 'medium', accent: 'accent1', header: true, banded: true, firstCol: false, lastRow: false,
  };
}

/** The shapes of the gallery, in groups like the office suites. */
export const SHAPE_GROUPS = [
  { id: 'lines', items: ['line', 'arrow', 'doubleArrow'] },
  { id: 'rect', items: ['rect', 'roundRect', 'snip1Rect'] },
  { id: 'basic', items: ['ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'octagon', 'plus', 'donut', 'can', 'heart', 'cloud', 'star5', 'star6'] },
  { id: 'arrows', items: ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'chevron', 'homePlate'] },
  { id: 'callouts', items: ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout'] },
];

/** A new shape (or line) of a kind in a box, styled like the office suites' default shape. */
export function newShape(kind, x, y, w, h) {
  if (kind === 'line' || kind === 'arrow' || kind === 'doubleArrow') {
    return { id: newId(), type: 'line', x: round(x), y: round(y), w: round(w), h: round(h), line: { color: 'accent1', width: 1.5, dash: 'solid' }, arrowStart: kind === 'doubleArrow' ? 'triangle' : 'none', arrowEnd: kind === 'line' ? 'none' : 'triangle' };
  }
  const e = { id: newId(), type: 'shape', shape: kind, x: round(x), y: round(y), w: round(w), h: round(h), fill: { type: 'solid', color: 'accent1' }, line: { color: 'accent1:-25', width: 1, dash: 'solid' } };
  e.text = emptyBody({ anchor: 'middle' });
  return e;
}

/** The default size of a new shape (some are not square). */
export function defaultShapeSize(kind) {
  if (['rightArrow', 'leftArrow', 'leftRightArrow', 'homePlate', 'chevron', 'rect', 'roundRect', 'snip1Rect', 'parallelogram', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout'].includes(kind)) return [200, 120];
  if (['upArrow', 'downArrow', 'upDownArrow', 'can'].includes(kind)) return [110, 180];
  if (['line', 'arrow', 'doubleArrow'].includes(kind)) return [220, 0];
  return [150, 150];
}

const HEX = /^#[0-9a-fA-F]{6}$/;
const SCHEME_RE = /^(dk1|lt1|dk2|lt2|accent[1-6]|hlink)(:[+-][0-9]{1,2})?$/;
const PICTURE = /^data:image\/(png|jpeg|gif|webp);base64,[A-Za-z0-9+/=]+$/;
const FONT = /^[\p{L}\p{N} _.-]{1,64}$/u;
export const okColor = c => typeof c === 'string' && (HEX.test(c) || SCHEME_RE.test(c));
export const okPicture = s => typeof s === 'string' && PICTURE.test(s);

/** Elements from outside (the clipboard of another page): only what the server would take, so that saving never fails because of them. */
export function sanitizeElements(list) {
  const out = [];
  const fill = f => !f || typeof f !== 'object' ? undefined : f.type === 'none' ? f : f.type === 'solid' && okColor(f.color) ? f : f.type === 'gradient' && okColor(f.from) && okColor(f.to) ? f : f.type === 'image' && okPicture(f.src) ? f : undefined;
  const line = l => (l && typeof l === 'object' && okColor(l.color) ? { color: l.color, width: Number(l.width) || 1, dash: l.dash || 'solid' } : undefined);
  const body = b => {
    if (!b || !Array.isArray(b.paras)) return emptyBody();
    for (const p of b.paras) {
      p.runs = Array.isArray(p.runs) ? p.runs.filter(r => r && typeof r.t === 'string') : [];
      if (!p.runs.length) p.runs = [{ t: '' }];
      for (const r of p.runs) { if (r.color && !okColor(r.color)) delete r.color; if (r.hl && !okColor(r.hl)) delete r.hl; if (r.font && !FONT.test(r.font)) delete r.font; }
    }
    return b;
  };
  for (const e of Array.isArray(list) ? list : []) {
    if (!e || typeof e !== 'object' || !['text', 'shape', 'line', 'image', 'table'].includes(e.type)) continue;
    for (const k of ['x', 'y', 'w', 'h']) e[k] = Number.isFinite(Number(e[k])) ? Number(e[k]) : 0;
    if (e.type === 'image' && !okPicture(e.src)) continue;
    if ('fill' in e) { e.fill = fill(e.fill); if (!e.fill) delete e.fill; }
    if ('line' in e) { e.line = line(e.line); if (!e.line) delete e.line; }
    if (e.type === 'line' && !e.line) e.line = { color: 'dk1', width: 1, dash: 'solid' };
    if (e.text) e.text = body(e.text);
    if (e.type === 'table') {
      if (!Array.isArray(e.cols) || !Array.isArray(e.rows) || !e.cols.length || !e.rows.length) continue;
      for (const r of e.rows) { r.cells = Array.isArray(r.cells) ? r.cells : []; for (const c of r.cells) { c.text = body(c.text); if (c.fill) { c.fill = fill(c.fill); if (!c.fill) delete c.fill; } } }
      if (e.accent && !okColor(e.accent)) e.accent = 'accent1';
    }
    out.push(e);
  }
  return out;
}

/** Quick styles of shapes: for each colour a filled, a subtle, an outlined and an intense look. */
export const QUICK_STYLES = ['dk1', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].flatMap(c => [
  { id: c + '-fill', fill: { type: 'solid', color: c }, line: { color: c + ':-25', width: 1, dash: 'solid' }, text: c === 'dk1' ? 'lt1' : 'lt1', shadow: false },
  { id: c + '-subtle', fill: { type: 'solid', color: c + ':+80' }, line: { color: c, width: 1, dash: 'solid' }, text: c + ':-25', shadow: false },
  { id: c + '-outline', fill: { type: 'solid', color: 'lt1' }, line: { color: c, width: 2, dash: 'solid' }, text: c, shadow: false },
  { id: c + '-intense', fill: { type: 'gradient', from: c + ':+40', to: c + ':-25', angle: 180 }, line: null, text: 'lt1', shadow: true },
]);
