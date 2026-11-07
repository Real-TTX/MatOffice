/* MatOffice Docs – the document model without the DOM (tested in Node: tests/js/docs-model.test.mjs).
   A document is a list of blocks (Docs HTML, see docs/EDITORS.md "Docs content"):
     text block  { type: 'p'|'h1'…'h6'|'blockquote'|'pre', style?, list?, level?, start?, id?, css: { 'text-align': 'center', … }, runs: [...] }
     table       { type: 'table', tableStyle, width, cols: ['30%', …], rows: [{ header, cells: [{ header, colspan, rowspan, css, blocks }] }] }
     rule        { type: 'hr' }        break { type: 'break', kind: 'page'|'column' }
   A run is text with a format, or one object: { text, f } | { obj: 'br'|'img'|'field'|'note', …, f }.
   A format f holds only what is set: b, i, u, s, sub, sup (true), font (a font-family list), size ('12pt'), color, bg ('#rrggbb'), ls ('1pt'),
   href, cmt (comment ids separated by spaces). The writer produces the canonical nesting of the server:
   a › span[data-comment] › span[style] › strong › em › u › s › sub/sup › text. */

// ---- units and the page ---------------------------------------------------------------------------------------------------------

export const PX_PER_MM = 96 / 25.4;
export const PX_PER_PT = 96 / 72;
export const PT_PER_MM = 72 / 25.4;

export const PAGE_SIZES = { A4: [210, 297], A5: [148, 210], A3: [297, 420], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };

export const DEFAULT_PAGE = Object.freeze({
  size: 'A4', width: 210, height: 297, orientation: 'portrait',
  margins: Object.freeze({ top: 25, right: 25, bottom: 20, left: 25 }), header: 12.5, footer: 12.5, columns: 1, differentFirst: false,
});

/** Margin presets of Layout → Margins (mm). */
export const MARGIN_PRESETS = {
  normal: { top: 25, right: 25, bottom: 20, left: 25 },
  word: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  narrow: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  moderate: { top: 25.4, right: 19.05, bottom: 25.4, left: 19.05 },
  wide: { top: 25.4, right: 50.8, bottom: 25.4, left: 50.8 },
};

const num = (v, fallback) => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v)) ? Number(v) : fallback);
const round2 = v => Math.round(v * 100) / 100;

/** The page with the rules of the server: named sizes have their measures, the rest is clamped, at least 20 mm of text stay. */
export function normalizePage(input) {
  const p = input && typeof input === 'object' ? input : {};
  const m = p.margins && typeof p.margins === 'object' ? p.margins : {};
  let size = typeof p.size === 'string' && (PAGE_SIZES[p.size] || p.size === 'Custom') ? p.size : 'A4';
  let width = num(p.width, 210);
  let height = num(p.height, 297);
  if (PAGE_SIZES[size]) [width, height] = PAGE_SIZES[size];
  else { width = round2(Math.min(1000, Math.max(50, width))); height = round2(Math.min(1000, Math.max(50, height))); }
  const landscape = p.orientation === 'landscape';
  const clampM = v => round2(Math.min(150, Math.max(0, num(v, 0))));
  let top = clampM(num(m.top, DEFAULT_PAGE.margins.top));
  let right = clampM(num(m.right, DEFAULT_PAGE.margins.right));
  let bottom = clampM(num(m.bottom, DEFAULT_PAGE.margins.bottom));
  let left = clampM(num(m.left, DEFAULT_PAGE.margins.left));
  const paperW = landscape ? height : width;
  const paperH = landscape ? width : height;
  [left, right] = fit(left, right, paperW);
  [top, bottom] = fit(top, bottom, paperH);
  return {
    size, width, height, orientation: landscape ? 'landscape' : 'portrait', margins: { top, right, bottom, left },
    header: clampM(num(p.header, 12.5)), footer: clampM(num(p.footer, 12.5)),
    columns: Math.min(3, Math.max(1, Math.round(num(p.columns, 1)))), differentFirst: p.differentFirst === true,
  };
}

function fit(a, b, paper) {
  const available = paper - 20;
  if (a + b <= available || a + b <= 0) return [a, b];
  const factor = Math.max(0, available) / (a + b);
  return [Math.floor(a * factor * 100) / 100, Math.floor(b * factor * 100) / 100];
}

/** The size of the paper as it lies (mm). */
export function paperSize(page) {
  return page.orientation === 'landscape' ? { width: page.height, height: page.width } : { width: page.width, height: page.height };
}

/** The geometry of a page in CSS pixels: paper, margins, the text area, the header/footer distances, the gap between columns. */
export function pageGeometry(page) {
  const paper = paperSize(page);
  const g = {
    paperWidth: paper.width * PX_PER_MM, paperHeight: paper.height * PX_PER_MM,
    top: page.margins.top * PX_PER_MM, right: page.margins.right * PX_PER_MM, bottom: page.margins.bottom * PX_PER_MM, left: page.margins.left * PX_PER_MM,
    header: page.header * PX_PER_MM, footer: page.footer * PX_PER_MM, columns: page.columns, columnGap: 12.5 * PX_PER_MM,
  };
  g.textWidth = Math.max(20, g.paperWidth - g.left - g.right);
  g.textHeight = Math.max(20, g.paperHeight - g.top - g.bottom);
  return g;
}

/** The size name of a paper in mm (portrait), or Custom. */
export function sizeName(width, height) {
  for (const [name, [w, h]] of Object.entries(PAGE_SIZES)) if (Math.abs(w - width) <= 2 && Math.abs(h - height) <= 2) return name;
  return 'Custom';
}

// ---- styles ---------------------------------------------------------------------------------------------------------------------

/** The paragraph styles (the same values as docs.css and the server's DocsStyles); next = the style of the paragraph after Enter. */
export const STYLES = [
  { key: 'Normal', tag: 'p', font: 'Calibri', size: 11, before: 0, after: 8, line: 1.3, next: 'Normal' },
  { key: 'NoSpacing', tag: 'p', style: 'NoSpacing', font: 'Calibri', size: 11, before: 0, after: 0, line: 1.2, next: 'NoSpacing' },
  { key: 'Title', tag: 'p', style: 'Title', font: 'Calibri Light', size: 28, before: 0, after: 4, line: 1.2, next: 'Normal' },
  { key: 'Subtitle', tag: 'p', style: 'Subtitle', font: 'Calibri', size: 11, color: '#5a5a5a', before: 0, after: 8, line: 1.3, next: 'Normal' },
  { key: 'Heading1', tag: 'h1', font: 'Calibri Light', size: 16, color: '#2f5496', before: 12, after: 0, line: 1.3, next: 'Normal', heading: 1 },
  { key: 'Heading2', tag: 'h2', font: 'Calibri Light', size: 13, color: '#2f5496', before: 2, after: 0, line: 1.3, next: 'Normal', heading: 2 },
  { key: 'Heading3', tag: 'h3', font: 'Calibri Light', size: 12, color: '#1f3763', before: 2, after: 0, line: 1.3, next: 'Normal', heading: 3 },
  { key: 'Heading4', tag: 'h4', font: 'Calibri Light', size: 11, color: '#2f5496', italic: true, before: 2, after: 0, line: 1.3, next: 'Normal', heading: 4 },
  { key: 'Heading5', tag: 'h5', font: 'Calibri Light', size: 11, color: '#2f5496', before: 2, after: 0, line: 1.3, next: 'Normal', heading: 5 },
  { key: 'Heading6', tag: 'h6', font: 'Calibri Light', size: 11, color: '#1f3763', before: 2, after: 0, line: 1.3, next: 'Normal', heading: 6 },
  { key: 'Quote', tag: 'blockquote', font: 'Calibri', size: 11, color: '#404040', italic: true, before: 10, after: 8, line: 1.3, next: 'Normal' },
  { key: 'IntenseQuote', tag: 'blockquote', style: 'IntenseQuote', font: 'Calibri', size: 11, color: '#2f5496', italic: true, before: 18, after: 18, line: 1.3, next: 'Normal' },
  { key: 'ListParagraph', tag: 'p', style: 'ListParagraph', font: 'Calibri', size: 11, before: 0, after: 8, line: 1.3, next: 'ListParagraph' },
  { key: 'Caption', tag: 'p', style: 'Caption', font: 'Calibri', size: 9, color: '#44546a', italic: true, before: 0, after: 10, line: 1.3, next: 'Normal' },
  { key: 'Code', tag: 'pre', font: 'Consolas', size: 10, before: 0, after: 8, line: 1.2, next: 'Normal' },
  { key: 'TOCHeading', tag: 'p', style: 'TOCHeading', font: 'Calibri Light', size: 16, color: '#2f5496', before: 12, after: 0, line: 1.3, next: 'Normal' },
  { key: 'TOC1', tag: 'p', style: 'TOC1', font: 'Calibri', size: 11, before: 0, after: 5, line: 1.3, next: 'Normal' },
  { key: 'TOC2', tag: 'p', style: 'TOC2', font: 'Calibri', size: 11, before: 0, after: 5, line: 1.3, next: 'Normal' },
  { key: 'TOC3', tag: 'p', style: 'TOC3', font: 'Calibri', size: 11, before: 0, after: 5, line: 1.3, next: 'Normal' },
];
const STYLE_BY_KEY = new Map(STYLES.map(s => [s.key, s]));

export const P_STYLES = new Set(['NoSpacing', 'Title', 'Subtitle', 'ListParagraph', 'Caption', 'TOCHeading', 'TOC1', 'TOC2', 'TOC3']);
export const TEXT_TAGS = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre']);

export const styleOf = key => STYLE_BY_KEY.get(key) || STYLE_BY_KEY.get('Normal');

/** The style key of a text block. */
export function styleKey(block) {
  switch (block.type) {
    case 'blockquote': return block.style === 'IntenseQuote' ? 'IntenseQuote' : 'Quote';
    case 'pre': return 'Code';
    case 'p': return block.style && P_STYLES.has(block.style) ? block.style : 'Normal';
    default: return /^h[1-6]$/.test(block.type) ? 'Heading' + block.type[1] : 'Normal';
  }
}

/** A text block with another style (lists stay on p and headings; a quote or code is no list). */
export function withStyle(block, key) {
  const s = styleOf(key);
  const b = { ...block, type: s.tag };
  delete b.style;
  if (s.style) b.style = s.style;
  if (s.tag === 'blockquote' || s.tag === 'pre' || key === 'Title' || key === 'Subtitle' || key.startsWith('TOC')) { delete b.list; delete b.level; delete b.start; }
  if (!/^h[1-6]$/.test(s.tag)) delete b.id;
  return b;
}

// ---- lists ----------------------------------------------------------------------------------------------------------------------

export const BULLETS = ['bullet', 'circle', 'square', 'dash', 'check', 'arrow'];
export const NUMBERS = ['decimal', 'lower-alpha', 'upper-alpha', 'lower-roman', 'upper-roman', 'outline'];
export const isListKind = kind => BULLETS.includes(kind) || NUMBERS.includes(kind);
export const isBullet = kind => BULLETS.includes(kind);

/** The marker of a level: bullet cycles • ◦ ▪ and decimal 1. a. i.; the other kinds keep theirs. */
export function levelFormat(kind, level) {
  if (kind === 'bullet') return ['bullet', 'circle', 'square'][level % 3];
  if (kind === 'decimal') return ['decimal', 'lower-alpha', 'lower-roman'][level % 3];
  if (kind === 'outline') return 'decimal';
  return kind;
}

const BULLET_CHARS = { bullet: '•', circle: '◦', square: '▪', dash: '–', check: '✓', arrow: '➢' };

/** The text of a marker (path = the numbers of the levels above, for outline lists). */
export function marker(kind, level, n, path) {
  const format = levelFormat(kind, level);
  if (BULLET_CHARS[format]) return BULLET_CHARS[format];
  if (kind === 'outline' && path) return path.join('.') + '.';
  switch (format) {
    case 'lower-alpha': return alpha(n) + '.';
    case 'upper-alpha': return alpha(n).toUpperCase() + '.';
    case 'lower-roman': return roman(n).toLowerCase() + '.';
    case 'upper-roman': return roman(n) + '.';
    default: return n + '.';
  }
}

export function alpha(n) {
  if (n <= 0) return '0';
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

export function roman(n) {
  if (n <= 0 || n >= 4000) return String(n);
  const table = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, t] of table) while (n >= v) { s += t; n -= v; }
  return s;
}

/** The numbers of the list items of a block list the way Docs counts them (a list ends at any other block; deeper levels restart). */
export function listNumbers(blocks) {
  const result = new Map();
  let counters = null, kinds = null;
  for (const b of blocks) {
    if (!b || !b.list) { counters = null; kinds = null; continue; }
    if (!counters) { counters = new Array(9).fill(0); kinds = new Array(9).fill(null); }
    const level = Math.min(8, Math.max(0, b.level || 0));
    // a level whose kind changes starts again (like the server and Word)
    if (kinds[level] !== b.list) { counters[level] = 0; kinds[level] = b.list; }
    counters[level] = b.start ? b.start : counters[level] + 1;
    for (let i = level + 1; i < 9; i++) { counters[i] = 0; kinds[i] = null; }
    result.set(b, { n: counters[level], path: counters.slice(0, level + 1).map(c => Math.max(1, c)) });
  }
  return result;
}

// ---- fonts and sizes ------------------------------------------------------------------------------------------------------------

/** Font families that exist on Windows, macOS or Linux (with metric-compatible substitutes in the stacks). */
export const FONTS = [
  'Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Arial Narrow', 'Book Antiqua', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas',
  'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Helvetica', 'Impact', 'Lucida Console', 'Lucida Sans Unicode',
  'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana',
];

const SUBSTITUTES = {
  'calibri': 'Carlito', 'calibri light': 'Carlito', 'cambria': 'Caladea', 'arial': "'Liberation Sans', Helvetica", 'helvetica': 'Arial',
  'times new roman': "'Liberation Serif', Times", 'courier new': "'Liberation Mono', Courier", 'consolas': "'Courier New'", 'georgia': "'Times New Roman'",
  'segoe ui': "'Helvetica Neue', Arial", 'aptos': 'Calibri, Carlito', 'aptos display': 'Calibri, Carlito', 'arial narrow': "'Liberation Sans Narrow'",
  'book antiqua': "Palatino, 'Palatino Linotype'", 'palatino linotype': "Palatino, 'Book Antiqua'", 'garamond': "'EB Garamond'", 'tahoma': 'Verdana',
  'lucida console': 'Monaco', 'lucida sans unicode': "'Lucida Grande'", 'century gothic': "'URW Gothic'", 'trebuchet ms': 'Tahoma',
};
const SERIF = new Set(['times new roman', 'times', 'cambria', 'georgia', 'garamond', 'book antiqua', 'palatino', 'palatino linotype', 'constantia', 'baskerville', 'liberation serif', 'caladea', 'noto serif', 'century schoolbook']);
const MONO = new Set(['consolas', 'courier new', 'courier', 'lucida console', 'cascadia code', 'cascadia mono', 'liberation mono', 'menlo', 'monaco', 'source code pro', 'fira code']);

export function genericOf(name) {
  const n = name.trim().toLowerCase();
  return SERIF.has(n) ? 'serif' : MONO.has(n) || n.includes('mono') ? 'monospace' : 'sans-serif';
}

/** A font-family value with fallbacks: "Calibri, Carlito, sans-serif". */
export function fontStack(name) {
  const clean = String(name || '').replace(/["';\\<>{}]/g, '').trim();
  if (!clean) return 'sans-serif';
  const lower = clean.toLowerCase();
  if (lower === 'serif' || lower === 'sans-serif' || lower === 'monospace') return lower;
  const quoted = /[^A-Za-z0-9-]/.test(clean) ? `'${clean}'` : clean;
  const sub = SUBSTITUTES[lower];
  return sub ? `${quoted}, ${sub}, ${genericOf(clean)}` : `${quoted}, ${genericOf(clean)}`;
}

/** The first family of a font-family list, without quotes. */
export function firstFamily(value) {
  if (!value) return '';
  let quote = '';
  for (let i = 0; i < value.length; i++) {
    const c = value[i];
    if (quote) { if (c === quote) quote = ''; }
    else if (c === '"' || c === "'") quote = c;
    else if (c === ',') return value.slice(0, i).trim().replace(/^["']|["']$/g, '').trim();
  }
  return value.trim().replace(/^["']|["']$/g, '').trim();
}

export const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** The next font size up (dir 1) or down (-1) in the list of sizes, like Word's Grow/Shrink Font. */
export function stepSize(pt, dir) {
  if (dir > 0) return SIZES.find(s => s > pt + 0.01) ?? Math.min(1638, Math.round(pt + 10));
  for (let i = SIZES.length - 1; i >= 0; i--) if (SIZES[i] < pt - 0.01) return SIZES[i];
  return Math.max(1, Math.round(pt - 1));
}

// ---- colours --------------------------------------------------------------------------------------------------------------------

const NAMED = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', lime: '#00ff00', blue: '#0000ff', yellow: '#ffff00', cyan: '#00ffff',
  aqua: '#00ffff', magenta: '#ff00ff', fuchsia: '#ff00ff', gray: '#808080', grey: '#808080', silver: '#c0c0c0', maroon: '#800000', olive: '#808000',
  navy: '#000080', purple: '#800080', teal: '#008080', orange: '#ffa500', darkblue: '#00008b', darkred: '#8b0000', darkgreen: '#006400',
  lightgray: '#d3d3d3', lightgrey: '#d3d3d3', darkgray: '#a9a9a9', darkgrey: '#a9a9a9', pink: '#ffc0cb', brown: '#a52a2a', gold: '#ffd700',
  windowtext: '#000000', window: '#ffffff',
};

const hex2 = v => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0');

/** A colour as #rrggbb (lower case), or null (transparent, inherit, unknown). */
export function colour(value) {
  if (value === null || value === undefined) return null;
  const v = String(value).trim().toLowerCase();
  if (!v) return null;
  if (v[0] === '#') {
    const h = v.slice(1);
    if (!/^[0-9a-f]+$/.test(h)) return null;
    if (h.length === 3 || h.length === 4) return '#' + h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (h.length === 6 || h.length === 8) return '#' + h.slice(0, 6);
    return null;
  }
  const m = /^rgba?\(([^)]*)\)$/.exec(v);
  if (m) {
    const parts = m[1].replace(/\//g, ' ').split(/[\s,]+/).filter(Boolean);
    if (parts.length < 3) return null;
    const ch = p => (p.endsWith('%') ? parseFloat(p) * 2.55 : parseFloat(p));
    if (parts.length > 3) { const a = parts[3].endsWith('%') ? parseFloat(parts[3]) / 100 : parseFloat(parts[3]); if (a <= 0.01) return null; }
    const [r, g, b] = parts.slice(0, 3).map(ch);
    if ([r, g, b].some(x => !Number.isFinite(x))) return null;
    return '#' + hex2(r) + hex2(g) + hex2(b);
  }
  return NAMED[v] || null;
}

/** The highlight colours of the ribbon (Word's highlighter colours). */
export const HIGHLIGHTS = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0', '#000000'];

// ---- CSS of the style attribute -------------------------------------------------------------------------------------------------

const FORBIDDEN_CSS = /url\(|expression|javascript|vbscript|@import|var\(|behavior|binding|image\(|image-set\(/i;

/** The declarations of a style attribute: [[name, value], …] with lower-case names; unsafe values are dropped. */
export function parseStyle(text) {
  const out = [];
  if (!text) return out;
  let depth = 0, quote = '', start = 0;
  const parts = [];
  for (let i = 0; i <= text.length; i++) {
    const c = text[i];
    if (i === text.length || (c === ';' && !quote && depth === 0)) { parts.push(text.slice(start, i)); start = i + 1; continue; }
    if (quote) { if (c === quote) quote = ''; }
    else if (c === '"' || c === "'") quote = c;
    else if (c === '(') depth++;
    else if (c === ')') depth = Math.max(0, depth - 1);
  }
  for (const part of parts) {
    const colon = part.indexOf(':');
    if (colon <= 0) continue;
    const name = part.slice(0, colon).trim().toLowerCase();
    const value = part.slice(colon + 1).replace(/!important/gi, '').trim();
    if (!name || !value || FORBIDDEN_CSS.test(value) || /[<>{}\\]/.test(value)) continue;
    if (name.startsWith('--') && !name.startsWith('--mo-')) continue;
    out.push([name, value]);
  }
  return out;
}

/** A CSS length in points (px, pt, pc, in, cm, mm, em/rem relative to emPt), or null. */
export function toPt(value, emPt = 11) {
  if (value === null || value === undefined) return null;
  const v = String(value).trim().toLowerCase();
  if (v === '0') return 0;
  const m = /^(-?\d*\.?\d+)(px|pt|pc|in|cm|mm|em|rem|q)?$/.exec(v);
  if (!m) {
    if (v === 'thin') return 0.75;
    if (v === 'medium') return 2.25;
    if (v === 'thick') return 3.75;
    return null;
  }
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case 'pt': return n;
    case 'pc': return n * 12;
    case 'in': return n * 72;
    case 'cm': return n * 72 / 2.54;
    case 'mm': return n * 72 / 25.4;
    case 'q': return n * 72 / 101.6;
    case 'em': case 'rem': return n * emPt;
    case 'px': return n * 0.75;
    default: return n === 0 ? 0 : null;
  }
}

/** A number written with at most two decimals. */
export const numStr = n => String(Math.round(n * 100) / 100);
export const pt = n => numStr(n) + 'pt';

const lengthPt = (v, min, max) => { const p = toPt(v); return p !== null && p >= min && p <= max ? pt(p) : null; };

/** A border in canonical form (0.5pt solid #000000), or null. */
export function borderCss(value) {
  let width = null, style = null, col = null;
  for (const part of String(value).trim().split(/\s+(?![^(]*\))/)) {
    const p = part.toLowerCase();
    if (p === 'none' || p === 'hidden') return null;
    if (/^(solid|dashed|dotted|double|groove|ridge|inset|outset)$/.test(p)) style = p;
    else if (toPt(p) !== null && /^[\d.]|^thin$|^medium$|^thick$/.test(p)) width = toPt(p);
    else if (colour(p)) col = colour(p);
  }
  if (!style || (width !== null && width <= 0)) return null;
  return `${pt(Math.min(6, Math.max(0.25, width ?? 0.75)))} ${style} ${col || '#000000'}`;
}

const BLOCK_PROPS = ['text-align', 'line-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right', 'text-indent', 'background-color',
  'border-top', 'border-right', 'border-bottom', 'border-left', 'padding-top', 'padding-right', 'padding-bottom', 'padding-left'];

/** Paragraph formatting from declarations, in canonical values (the order of the server: the order they come in). */
export function blockCss(decls) {
  const css = {};
  const set = (name, value) => { if (value !== null && value !== undefined) css[name] = value; else delete css[name]; };
  const sides = ['top', 'right', 'bottom', 'left'];
  for (const [name, value] of decls) {
    switch (name) {
      case 'text-align': {
        const v = value.toLowerCase();
        set(name, v === 'start' ? 'left' : v === 'end' ? 'right' : v === '-webkit-center' ? 'center' : /^(left|right|center|justify)$/.test(v) ? v : null);
        break;
      }
      case 'line-height': {
        const v = value.toLowerCase();
        if (/^\d*\.?\d+%$/.test(v)) { const n = parseFloat(v) / 100; set(name, n >= 0.5 && n <= 10 ? numStr(n) : null); }
        else if (/^\d*\.?\d+$/.test(v)) { const n = parseFloat(v); set(name, n >= 0.5 && n <= 10 ? numStr(n) : null); }
        else set(name, lengthPt(v, 1, 1000));
        break;
      }
      case '--mo-sb': set('margin-top', lengthPt(value, -1000, 1000)); break;
      case '--mo-pt': set('padding-top', lengthPt(value, 0, 200)); break;
      case 'margin': case 'padding': {
        const p = value.split(/\s+/);
        if (p.length < 1 || p.length > 4) break;
        const v = [p[0], p[1] ?? p[0], p[2] ?? p[0], p[3] ?? p[1] ?? p[0]];
        sides.forEach((s, i) => set(`${name}-${s}`, name === 'margin' ? lengthPt(v[i], -1000, 1000) : lengthPt(v[i], 0, 200)));
        break;
      }
      case 'margin-top': case 'margin-bottom': case 'margin-left': case 'margin-right': case 'text-indent': set(name, lengthPt(value, -1000, 1000)); break;
      case 'padding-top': case 'padding-bottom': case 'padding-left': case 'padding-right': set(name, lengthPt(value, 0, 200)); break;
      case 'background-color': case 'background': set('background-color', colour(value)); break;
      case 'border': sides.forEach(s => set('border-' + s, borderCss(value))); break;
      case 'border-top': case 'border-right': case 'border-bottom': case 'border-left': set(name, borderCss(value)); break;
      default: break;
    }
  }
  const ordered = {};
  for (const k of Object.keys(css)) if (BLOCK_PROPS.includes(k)) ordered[k] = css[k];
  return ordered;
}

/** The inline formatting of declarations (a span's style, Word's and Google's spans): a partial format. */
export function inlineFormat(decls) {
  const f = {};
  for (const [name, value] of decls) {
    const v = value.toLowerCase();
    switch (name) {
      case 'font-family': { const family = firstFamily(value); if (family && !/^(inherit|initial|unset)$/i.test(family)) f.font = fontStack(family); break; }
      case 'font-size': { const p = /%$/.test(v) ? parseFloat(v) / 100 * 11 : toPt(v); if (p && p > 0 && p <= 1638) f.size = pt(p); break; }
      case 'color': { const c = colour(v); if (c) f.color = c; break; }
      case 'background-color': case 'background': { const c = colour(v); if (c) f.bg = c; break; }
      case 'letter-spacing': { const p = toPt(v); if (p !== null && Math.abs(p) <= 100 && p !== 0) f.ls = pt(p); break; }
      case 'font-weight': if (v === 'bold' || v === 'bolder' || parseInt(v, 10) >= 600) f.b = true; else if (v === 'normal' || parseInt(v, 10) < 600) f.b = false; break;
      case 'font-style': if (v === 'italic' || v === 'oblique') f.i = true; else if (v === 'normal') f.i = false; break;
      case 'text-decoration': case 'text-decoration-line':
        if (v.includes('underline')) f.u = true;
        if (v.includes('line-through')) f.s = true;
        if (v === 'none') { f.u = false; f.s = false; }
        break;
      case 'vertical-align': if (v === 'super') f.sup = true; else if (v === 'sub') f.sub = true; break;
      default: break;
    }
  }
  return f;
}

// ---- escaping -------------------------------------------------------------------------------------------------------------------

export const escText = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const escAttr = s => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;');

/** Text of a document: no control characters except tab and newline (and those only where they belong). */
export function cleanText(s) {
  // eslint-disable-next-line no-control-regex
  return String(s).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f﻿]/g, '');
}

// ---- formats and runs -----------------------------------------------------------------------------------------------------------

const FORMAT_KEYS = ['href', 'cmt', 'font', 'size', 'color', 'bg', 'ls', 'b', 'i', 'u', 's', 'sub', 'sup'];

/** A format with only the keys that are set (false/empty dropped; sub and sup exclude each other). */
export function cleanFormat(f) {
  const out = {};
  if (!f) return out;
  for (const k of FORMAT_KEYS) {
    const v = f[k];
    if (v === undefined || v === null || v === false || v === '') continue;
    out[k] = v === true || typeof v === 'string' ? v : String(v);
  }
  if (out.sub && out.sup) delete out.sub;
  return out;
}

export function sameFormat(a, b) {
  for (const k of FORMAT_KEYS) if ((a[k] || undefined) !== (b[k] || undefined)) return false;
  return true;
}

export const runLength = r => (r.obj ? 1 : r.text.length);
export const runsLength = runs => runs.reduce((n, r) => n + runLength(r), 0);
export const textRun = (text, f = {}) => ({ text, f: cleanFormat(f) });

/** Adjacent text runs with the same format are merged, empty ones dropped. */
export function normalizeRuns(runs) {
  const out = [];
  for (const r of runs) {
    if (!r.obj && !r.text) continue;
    const last = out[out.length - 1];
    if (last && !last.obj && !r.obj && sameFormat(last.f, r.f)) out[out.length - 1] = { text: last.text + r.text, f: last.f };
    else out.push(r);
  }
  return out;
}

/** The runs split at an offset: [before, after]. */
export function splitRuns(runs, at) {
  const left = [], right = [];
  let pos = 0;
  for (const r of runs) {
    const len = runLength(r);
    if (pos + len <= at) left.push(r);
    else if (pos >= at) right.push(r);
    else { left.push({ text: r.text.slice(0, at - pos), f: r.f }); right.push({ text: r.text.slice(at - pos), f: r.f }); }
    pos += len;
  }
  return [left, right];
}

export function sliceRuns(runs, start, end) {
  return splitRuns(splitRuns(runs, end)[0], start)[1];
}

export function deleteRuns(runs, start, end) {
  if (end <= start) return runs;
  const [a] = splitRuns(runs, start);
  const [, b] = splitRuns(runs, end);
  return normalizeRuns([...a, ...b]);
}

export function insertRuns(runs, at, inserted) {
  const [a, b] = splitRuns(runs, at);
  return normalizeRuns([...a, ...inserted, ...b]);
}

/** Applies a format change to [start, end): patch is { b: true, size: '12pt', color: null … } or a function f => newF. */
export function formatRuns(runs, start, end, patch) {
  if (end <= start) return runs;
  const [a, rest] = splitRuns(runs, start);
  const [mid, b] = splitRuns(rest, end - start);
  const apply = typeof patch === 'function' ? patch : f => ({ ...f, ...patch });
  return normalizeRuns([...a, ...mid.map(r => ({ ...r, f: cleanFormat(apply({ ...r.f })) })), ...b]);
}

/** The format typing at an offset continues: the one of the character before (or after at the start); links and comments do not continue at their end. */
export function formatAt(runs, offset) {
  let pos = 0;
  let before = null, after = null;
  for (const r of runs) {
    const len = runLength(r);
    if (pos < offset && offset <= pos + len) before = r;
    if (pos <= offset && offset < pos + len) { after = r; break; }
    pos += len;
  }
  const src = before && before.obj !== 'br' ? before : after || before;
  if (!src) return {};
  const f = { ...src.f };
  if (before !== after) {
    if (src === before) {
      if (f.href && after?.f.href !== f.href) delete f.href;
      if (f.cmt && after?.f.cmt !== f.cmt) delete f.cmt;
    } else {
      delete f.href;
      delete f.cmt;
    }
  }
  return cleanFormat(f);
}

/** The text of runs; objects count as one character (U+FFFC), line breaks as \n. */
export function runsText(runs, objects = '￼') {
  let s = '';
  for (const r of runs) s += r.obj ? (r.obj === 'br' ? '\n' : objects) : r.text;
  return s;
}

/** Whether every character of [start, end) has the format test (or, for an empty range, the character before). */
export function rangeHas(runs, start, end, test) {
  if (end <= start) return !!test(formatAt(runs, start));
  let pos = 0, any = false;
  for (const r of runs) {
    const len = runLength(r);
    if (pos < end && pos + len > start && !(r.obj === 'br')) { any = true; if (!test(r.f)) return false; }
    pos += len;
  }
  return any;
}

/** The value of a format key over [start, end): the common value or '' when it differs. */
export function rangeValue(runs, start, end, key) {
  if (end <= start) return formatAt(runs, start)[key] ?? null;
  let pos = 0, value;
  for (const r of runs) {
    const len = runLength(r);
    if (pos < end && pos + len > start && r.obj !== 'br') {
      const v = r.f[key] ?? null;
      if (value === undefined) value = v;
      else if (value !== v) return '';
    }
    pos += len;
  }
  return value === undefined ? null : value;
}

// ---- blocks ---------------------------------------------------------------------------------------------------------------------

export const isTextBlock = b => !!b && TEXT_TAGS.has(b.type);
export const blockLength = b => runsLength(b.runs || []);

export function textBlock(type = 'p', props = {}) {
  return { type, css: {}, runs: [], ...props };
}

/** A copy of a text block's attributes without its text (Enter: the new paragraph). */
export function shellOf(b) {
  const c = { type: b.type, css: { ...b.css }, runs: [] };
  if (b.style) c.style = b.style;
  if (b.list) { c.list = b.list; c.level = b.level || 0; }
  return c;
}

/** Splits a text block at an offset: the first keeps id and list start, the second gets the attributes without them. */
export function splitBlock(b, at) {
  const [left, right] = splitRuns(b.runs, at);
  const first = { ...b, css: { ...b.css }, runs: normalizeRuns(left) };
  const second = shellOf(b);
  second.runs = normalizeRuns(right);
  return [first, second];
}

/** Two text blocks become one (the attributes of the first). */
export function mergeBlocks(a, b) {
  return { ...a, css: { ...a.css }, runs: normalizeRuns([...a.runs, ...b.runs]) };
}

export const blockText = b => runsText(b.runs || []);

/** Every text block of a block list, also those in table cells, in document order. */
export function* textBlocks(blocks) {
  for (const b of blocks) {
    if (isTextBlock(b)) yield b;
    else if (b.type === 'table') for (const row of b.rows) for (const cell of row.cells) yield* textBlocks(cell.blocks);
  }
}

// ---- the writer -----------------------------------------------------------------------------------------------------------------

function spanStyle(f) {
  const parts = [];
  if (f.font) parts.push('font-family: ' + f.font);
  if (f.size) parts.push('font-size: ' + f.size);
  if (f.color) parts.push('color: ' + f.color);
  if (f.bg) parts.push('background-color: ' + f.bg);
  if (f.ls) parts.push('letter-spacing: ' + f.ls);
  return parts.join('; ');
}

function chainOf(f, object) {
  const chain = [];
  if (f.href) chain.push(['<a href="' + escAttr(f.href) + '">', '</a>']);
  if (f.cmt) chain.push(['<span data-comment="' + escAttr(f.cmt) + '">', '</span>']);
  const style = spanStyle(f);
  if (style) chain.push(['<span style="' + escAttr(style) + '">', '</span>']);
  if (!object) {
    if (f.b) chain.push(['<strong>', '</strong>']);
    if (f.i) chain.push(['<em>', '</em>']);
    if (f.u) chain.push(['<u>', '</u>']);
    if (f.s) chain.push(['<s>', '</s>']);
    if (f.sup) chain.push(['<sup>', '</sup>']);
    else if (f.sub) chain.push(['<sub>', '</sub>']);
  }
  return chain;
}

/** Writes runs as canonical inline HTML; options.image(src) can replace picture data (the undo history keeps pictures once). */
export function writeRuns(runs, options = {}) {
  let html = '';
  let open = [];
  for (const r of runs) {
    const chain = chainOf(r.f || {}, !!r.obj && r.obj !== 'br');
    let common = 0;
    while (common < open.length && common < chain.length && open[common][0] === chain[common][0]) common++;
    for (let i = open.length - 1; i >= common; i--) html += open[i][1];
    open = open.slice(0, common);
    for (let i = common; i < chain.length; i++) { html += chain[i][0]; open.push(chain[i]); }
    if (!r.obj) html += escText(r.text);
    else if (r.obj === 'br') html += '<br>';
    else if (r.obj === 'img') {
      const src = options.image ? options.image(r.src) : r.src;
      const size = [r.width ? 'width: ' + numStr(r.width) + 'px' : '', r.height ? 'height: ' + numStr(r.height) + 'px' : ''].filter(Boolean).join('; ');
      html += '<img src="' + escAttr(src) + '" alt="' + escAttr(r.alt || '') + '"' + (size ? ' style="' + size + '"' : '') + (r.align ? ' data-align="' + r.align + '"' : '') + '>';
    } else if (r.obj === 'field') html += '<span data-field="' + escAttr(r.kind) + '"></span>';
    else if (r.obj === 'note') html += '<sup data-fn="' + escAttr(r.id) + '"></sup>';
  }
  for (let i = open.length - 1; i >= 0; i--) html += open[i][1];
  return html;
}

const cssText = css => Object.entries(css || {}).map(([k, v]) => `${k}: ${v}`).join('; ');

function attr(name, value) { return value === undefined || value === null || value === '' ? '' : ` ${name}="${escAttr(String(value))}"`; }

/** Writes blocks as canonical Docs HTML. */
export function writeBlocks(blocks, options = {}) {
  let html = '';
  for (const b of blocks) html += writeBlock(b, options);
  return html;
}

export function writeBlock(b, options = {}) {
  if (isTextBlock(b)) {
    let open = '<' + b.type + attr('id', b.id) + attr('data-style', b.style);
    if (b.list) open += attr('data-list', b.list) + attr('data-level', b.level > 0 ? b.level : null) + attr('data-start', b.start || null);
    const css = cssText(b.css);
    open += attr('style', css) + '>';
    const inner = writeRuns(b.runs, options);
    return open + (b.type === 'pre' && /^\n/.test(b.runs[0]?.text || '') ? '\n' : '') + inner + '</' + b.type + '>';
  }
  if (b.type === 'hr') return '<hr>';
  if (b.type === 'break') return '<div data-break="' + (b.kind === 'column' ? 'column' : 'page') + '"></div>';
  if (b.type === 'table') return writeTable(b, options);
  return '';
}

function writeTable(t, options) {
  let html = '<table' + attr('data-table-style', t.tableStyle || 'grid') + attr('style', t.width ? 'width: ' + t.width : null) + '>';
  if ((t.cols || []).some(Boolean)) html += '<colgroup>' + t.cols.map(c => '<col' + attr('style', c ? 'width: ' + c : null) + '>').join('') + '</colgroup>';
  const head = [];
  for (const r of t.rows) { if (r.header) head.push(r); else break; }
  const row = (r, header) => '<tr>' + r.cells.map(c => {
    const tag = header || c.header ? 'th' : 'td';
    const inner = c.blocks.length ? writeBlocks(c.blocks, options) : '<p></p>';
    return '<' + tag + attr('colspan', c.colspan > 1 ? c.colspan : null) + attr('rowspan', c.rowspan > 1 ? c.rowspan : null) + attr('style', cssText(c.css)) + '>' + inner + '</' + tag + '>';
  }).join('') + '</tr>';
  if (head.length) html += '<thead>' + head.map(r => row(r, true)).join('') + '</thead>';
  if (t.rows.length > head.length) html += '<tbody>' + t.rows.slice(head.length).map(r => row(r, false)).join('') + '</tbody>';
  return html + '</table>';
}

// ---- tables ---------------------------------------------------------------------------------------------------------------------

export const TABLE_STYLES = ['grid', 'plain', 'light', 'striped', 'accent1', 'accent2', 'accent6'];

/** A new table: rows × columns of empty paragraphs, equal column widths, the first row a header row when asked. */
export function newTable(rows, cols, { header = false, style = 'grid' } = {}) {
  const width = numStr(100 / cols) + '%';
  return {
    type: 'table', tableStyle: style, width: '100%', cols: new Array(cols).fill(width),
    rows: Array.from({ length: rows }, (_, r) => ({ header: header && r === 0, cells: Array.from({ length: cols }, () => ({ header: header && r === 0, colspan: 1, rowspan: 1, css: {}, blocks: [textBlock()] })) })),
  };
}

/** The grid of a table: grid[r][c] = { cell, row, col (of the cell's origin) } – merged cells fill all their slots. */
export function tableGrid(t) {
  const grid = [];
  t.rows.forEach((row, r) => {
    grid[r] ??= [];
    let c = 0;
    for (const cell of row.cells) {
      while (grid[r][c]) c++;
      for (let dr = 0; dr < (cell.rowspan || 1) && r + dr < t.rows.length; dr++) {
        grid[r + dr] ??= [];
        for (let dc = 0; dc < (cell.colspan || 1); dc++) grid[r + dr][c + dc] = { cell, row: r, col: c };
      }
      c += cell.colspan || 1;
    }
  });
  return grid;
}

export function columnCount(t) {
  return Math.max(t.cols?.length || 0, ...tableGrid(t).map(r => r.length), 1);
}

// ---- text tools -----------------------------------------------------------------------------------------------------------------

/** Change case like Word (Shift+F3): sentence, lower, upper, title (capitalize each word), toggle. */
export function changeCase(text, mode, locale) {
  switch (mode) {
    case 'lower': return text.toLocaleLowerCase(locale);
    case 'upper': return text.toLocaleUpperCase(locale);
    case 'title': return text.toLocaleLowerCase(locale).replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase(locale));
    case 'toggle': return [...text].map(c => (c === c.toLocaleUpperCase(locale) ? c.toLocaleLowerCase(locale) : c.toLocaleUpperCase(locale))).join('');
    case 'sentence': {
      const lower = text.toLocaleLowerCase(locale);
      return lower.replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase(locale));
    }
    default: return text;
  }
}

/** The next mode of Shift+F3: lower → UPPER → Title → lower … starting from what the text is. */
export function nextCase(text) {
  if (text === text.toUpperCase() && text !== text.toLowerCase()) return 'lower';
  if (text === text.toLowerCase()) return 'upper';
  return 'upper';
}

const WORD = /[\p{L}\p{N}][\p{L}\p{N}'’\-_.@]*/gu;

export function countWords(text) {
  const m = text.match(WORD);
  return m ? m.length : 0;
}

/** Counts for the Word Count dialog. */
export function statistics(texts) {
  let words = 0, chars = 0, charsNoSpaces = 0, paragraphs = 0;
  for (const t of texts) {
    const clean = t.replace(/￼/g, '');
    words += countWords(clean);
    chars += clean.replace(/\n/g, '').length;
    charsNoSpaces += clean.replace(/\s/g, '').length;
    if (clean.trim()) paragraphs++;
  }
  return { words, chars, charsNoSpaces, paragraphs };
}

/** The matches of a search in a text: [[start, end], …]; options: matchCase, wholeWord. */
export function findAll(text, query, { matchCase = false, wholeWord = false } = {}) {
  if (!query) return [];
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const source = wholeWord ? `(?<![\\p{L}\\p{N}_])${escaped}(?![\\p{L}\\p{N}_])` : escaped;
  let re;
  try { re = new RegExp(source, matchCase ? 'gu' : 'giu'); } catch { return []; }
  const out = [];
  for (const m of text.matchAll(re)) { if (m[0].length) out.push([m.index, m.index + m[0].length]); if (out.length > 10000) break; }
  return out;
}

/** An id for a heading, a note or a comment. */
export function newId(prefix = '', random = Math.random) {
  let s = '';
  for (let i = 0; i < 6; i++) s += 'abcdefghijkmnopqrstuvwxyz23456789'[Math.floor(random() * 33)];
  return prefix + s;
}

/** The blocks of a table of contents for headings [{ level, text, id }] (levels 1–3). */
export function tocBlocks(headings, title) {
  const blocks = [{ type: 'p', style: 'TOCHeading', css: {}, runs: [textRun(title)] }];
  for (const h of headings) {
    if (h.level > 3 || !h.text.trim()) continue;
    blocks.push({ type: 'p', style: 'TOC' + h.level, css: {}, runs: [textRun(h.text.replace(/\s+/g, ' ').trim(), { href: '#' + h.id })] });
  }
  if (blocks.length === 1) blocks.push({ type: 'p', style: 'TOC1', css: {}, runs: [] });
  return blocks;
}

/** The text of a field outside of a page (date, time, title). */
export function fieldText(kind, { lang = 'en', name = '', page = 1, pages = 1, now = new Date() } = {}) {
  switch (kind) {
    case 'page': return String(page);
    case 'pages': return String(pages);
    case 'date': return now.toLocaleDateString(lang, { year: 'numeric', month: '2-digit', day: '2-digit' });
    case 'time': return now.toLocaleTimeString(lang, { hour: '2-digit', minute: '2-digit' });
    case 'title': return name;
    default: return '';
  }
}

/** Safe link targets (http, https, mailto, #id); anything else is null. */
export function safeHref(value) {
  if (!value) return null;
  const v = String(value).trim();
  if (/^#[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(v)) return v;
  if (/^mailto:[^\s<>"]+$/i.test(v)) return v;
  try {
    const u = new URL(v);
    if (u.protocol === 'http:' || u.protocol === 'https:') return u.href;
  } catch { /* not an address */ }
  return null;
}

/** What a person typed into the link dialog becomes an address: example.com → https://example.com/, a@b.c → mailto:. */
export function linkFromInput(value) {
  const v = String(value || '').trim();
  if (!v) return null;
  if (/^#/.test(v)) return safeHref(v);
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return 'mailto:' + v;
  if (/^(https?:|mailto:)/i.test(v)) return safeHref(v);
  if (/^[\w-]+(\.[\w-]+)+(:\d+)?(\/.*)?$/.test(v)) return safeHref('https://' + v);
  return null;
}

/** Safe pictures: data:image/(png|jpeg|gif|webp);base64 only. */
export const safeImage = src => (typeof src === 'string' && /^data:image\/(png|jpeg|gif|webp);base64,[A-Za-z0-9+/=\s]+$/.test(src) ? src.replace(/\s+/g, '') : null);

/** Ids of headings, notes and comments: letters, digits, _ and -. */
export const isId = id => typeof id === 'string' && /^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(id);
export const isItemId = id => typeof id === 'string' && /^[A-Za-z0-9_-]{1,40}$/.test(id);

/** The name of a language in the language of the page (de → Deutsch / German). */
export function languageName(code, ui = 'en') {
  try { return new Intl.DisplayNames([ui], { type: 'language' }).of(code) || code; } catch { return code; }
}
