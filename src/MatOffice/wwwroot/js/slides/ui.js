/* MatOffice Slides – small pieces of UI: context menus, the colour picker with the colours of the design, the outline menu, the galleries of
   shapes, layouts, designs and quick styles, the table grid, symbols. All built with el() (never HTML) in the popovers of the frame. */
import { openPopover } from '../editor/frame.js';
import { t, el, icon } from '../editor/util.js';
import {
  PALETTE_SLOTS, paletteColumn, STANDARD_COLORS, resolveColor, SHAPE_GROUPS, LAYOUTS, QUICK_STYLES, makeTheme, newSlide, decorOf,
  backgroundOf, layoutOf, THEMES,
} from './model.js';
import { shapePath } from './geometry.js';
import { renderSlide, cssBackground, fontStack } from './render.js';

const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (tag, attrs) => { const n = document.createElementNS(SVG, tag); for (const [k, v] of Object.entries(attrs || {})) n.setAttribute(k, String(v)); return n; };

/** A popover at a point of the screen (a context menu). */
export function popAt(x, y, content) {
  const anchor = el('div', { class: 'mo-slides-anchor', style: { left: x + 'px', top: y + 'px' } });
  document.body.append(anchor);
  const pop = openPopover(anchor, content);
  const cleanup = new MutationObserver(() => { if (!pop.isConnected) { anchor.remove(); cleanup.disconnect(); } });
  cleanup.observe(document.body, { childList: true });
  return pop;
}

/** A menu: [{ label, icon, key, run, checked, disabled, danger, sub: () => items }, '-', { heading }]. */
export function menuList(items, close) {
  const list = el('div', { class: 'mo-menu', role: 'menu' });
  for (const entry of items) {
    if (!entry) continue;
    if (entry === '-') { list.append(el('div', { class: 'mo-menu__sep', role: 'separator' })); continue; }
    if (entry.heading) { list.append(el('div', { class: 'mo-menu__heading', text: entry.heading })); continue; }
    if (entry.hidden) continue;
    const checked = typeof entry.checked === 'function' ? entry.checked() : entry.checked;
    const b = el('button', {
      type: 'button', role: checked === undefined ? 'menuitem' : 'menuitemcheckbox', 'aria-checked': checked === undefined ? null : String(!!checked),
      class: 'mo-menu__item' + (entry.danger ? ' mo-menu__item--danger' : ''), disabled: (typeof entry.disabled === 'function' ? entry.disabled() : entry.disabled) ? true : null,
      onclick: () => {
        if (entry.sub) { const r = b.getBoundingClientRect(); close(); setTimeout(() => popAt(r.right - 8, r.top, c => menuList(entry.sub(), c)), 0); return; }
        close();
        entry.run?.();
      },
    }, entry.render ? entry.render() : entry.icon ? icon(entry.icon) : el('span', { class: 'mo-menu__noicon' }), el('span', { class: 'mo-menu__label', text: entry.label }),
    entry.key ? el('span', { class: 'mo-menu__key', text: entry.key }) : null, entry.sub ? icon('chevron-right', 'mo-menu__sub') : checked ? icon('check', 'mo-menu__check') : null);
    list.append(b);
  }
  return list;
}

export function contextMenu(x, y, items) { return popAt(x, y, close => menuList(items, close)); }

// ---- colours ------------------------------------------------------------------------------------------------------------------

/**
 * The colour picker: the colours of the design in shades, the standard colours, "no fill"/"automatic" and a custom colour.
 * options: { theme, none: label?, auto: label?, current, pick(color|null), extra: Node? }
 */
export function colorPicker(anchor, options) {
  return openPopover(anchor, close => colorPanel(options, close));
}

export function colorPanel({ theme, none, auto, pick, extra, current }, close) {
  const box = el('div', { class: 'mo-palette mo-slides-palette' });
  const choose = c => { close(); pick(c); };
  if (auto) box.append(el('button', { type: 'button', class: 'mo-palette__auto', onclick: () => choose(null) }, el('span', { class: 'mo-palette__chip mo-palette__chip--auto' }), el('span', { text: auto })));
  box.append(el('div', { class: 'mo-menu__heading', text: t('Slides.Color.Theme') }));
  const grid = el('div', { class: 'mo-slides-palette__grid' });
  const columns = PALETTE_SLOTS.map(paletteColumn);
  for (let row = 0; row < 6; row++) {
    const line = el('div', { class: 'mo-palette__row' + (row === 0 ? ' mo-palette__row--gap' : '') });
    for (const col of columns) {
      const c = col[row];
      const hex = resolveColor(c, theme);
      const chip = el('button', { type: 'button', class: 'mo-palette__chip' + (c === current ? ' is-current' : ''), 'aria-label': colorName(c), 'data-tip': colorName(c), onclick: () => choose(c) });
      chip.style.background = hex;
      line.append(chip);
    }
    grid.append(line);
  }
  box.append(grid, el('div', { class: 'mo-menu__heading', text: t('Slides.Color.Standard') }));
  const std = el('div', { class: 'mo-palette__row mo-palette__row--gap' });
  for (const c of STANDARD_COLORS) {
    const chip = el('button', { type: 'button', class: 'mo-palette__chip', 'aria-label': c, 'data-tip': c, onclick: () => choose(c) });
    chip.style.background = c;
    std.append(chip);
  }
  box.append(std);
  if (none) box.append(el('button', { type: 'button', class: 'mo-palette__auto', onclick: () => choose('none') }, icon('slides-no-fill'), el('span', { text: none })));
  const custom = el('input', { type: 'color', class: 'mo-palette__custom', 'aria-label': t('Slides.Color.More'), value: current && current[0] === '#' ? current : '#3a6ea5' });
  custom.addEventListener('change', () => choose(custom.value));
  box.append(el('label', { class: 'mo-palette__more' }, icon('palette'), el('span', { text: t('Slides.Color.More') }), custom));
  if (extra) box.append(extra(close));
  return box;
}

export function colorName(c) {
  if (!c) return '';
  if (c[0] === '#') return c;
  const [name, mod] = c.split(':');
  const base = t('Slides.Color.' + name);
  if (!mod) return base;
  const n = Number(mod);
  return n > 0 ? t('Slides.Color.Lighter', base, n) : t('Slides.Color.Darker', base, -n);
}

/** A split colour button for the ribbon (custom item): the main part applies the last colour, the arrow opens the picker. */
export function colorButton({ label, iconName, theme, none, auto, value, change, canRun }) {
  let last = null;
  const bar = el('span', { class: 'mo-rcolor__bar' });
  const main = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--icon mo-rcolor__main', 'aria-label': label, 'data-tip': label }, icon(iconName), bar);
  const arrow = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--arrow', 'aria-label': t('Editor.MoreColours', label), 'aria-haspopup': 'true' }, icon('chevron-down'));
  const paint = () => { const c = last ?? value?.(); bar.style.background = c && c !== 'none' ? resolveColor(c, theme()) : 'transparent'; };
  main.addEventListener('mousedown', e => e.preventDefault());
  arrow.addEventListener('mousedown', e => e.preventDefault());
  main.addEventListener('click', () => { if (!canRun()) return; const c = last ?? value?.() ?? null; change(c); });
  arrow.addEventListener('click', () => {
    if (!canRun()) return;
    colorPicker(arrow, { theme: theme(), none, auto, current: value?.(), pick: c => { last = c; paint(); change(c); } });
  });
  const node = el('span', { class: 'mo-rsplit mo-rcolor' }, main, arrow);
  node.refresh = () => { main.disabled = arrow.disabled = !canRun(); paint(); };
  return node;
}

// ---- galleries ------------------------------------------------------------------------------------------------------------------

/** A small picture of a shape for menus and galleries. */
export function shapeIcon(kind, size = 22) {
  const s = svgEl('svg', { class: 'mo-slides-shapeicon', width: size, height: size, viewBox: '-2 -2 28 28', 'aria-hidden': 'true' });
  if (kind === 'line' || kind === 'arrow' || kind === 'doubleArrow') {
    s.append(svgEl('line', { x1: 2, y1: 22, x2: 22, y2: 2, stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round' }));
    if (kind !== 'line') s.append(svgEl('polygon', { points: '22,2 13,5 19,11', fill: 'currentColor' }));
    if (kind === 'doubleArrow') s.append(svgEl('polygon', { points: '2,22 5,13 11,19', fill: 'currentColor' }));
    return s;
  }
  const [w, h] = ['rightArrow', 'leftArrow', 'leftRightArrow', 'homePlate', 'chevron'].includes(kind) ? [24, 14] : ['upArrow', 'downArrow', 'upDownArrow'].includes(kind) ? [14, 24] : [24, 24 * (kind.includes('Callout') ? 0.7 : 1)];
  const g = svgEl('g', { transform: `translate(${(24 - w) / 2} ${(24 - h) / 2})` });
  g.append(svgEl('path', { d: shapePath(kind, w, h), fill: 'none', stroke: 'currentColor', 'stroke-width': 1.6, 'stroke-linejoin': 'round', 'fill-rule': 'evenodd' }));
  s.append(g);
  return s;
}

export function shapeGallery(anchor, pick) {
  return openPopover(anchor, close => {
    const box = el('div', { class: 'mo-slides-gallery mo-slides-gallery--shapes' });
    for (const group of SHAPE_GROUPS) {
      box.append(el('div', { class: 'mo-menu__heading', text: t('Slides.Shapes.' + group.id) }));
      const grid = el('div', { class: 'mo-slides-gallery__grid' });
      for (const kind of group.items) {
        grid.append(el('button', { type: 'button', class: 'mo-slides-gallery__item', 'aria-label': t('Slides.Shape.' + kind), 'data-tip': t('Slides.Shape.' + kind), onclick: () => { close(); pick(kind); } }, shapeIcon(kind)));
      }
      box.append(grid);
    }
    return box;
  });
}

/** A picture of a layout in the colours of the design (placeholders as outlines). */
export function layoutPreview(content, layoutId, width = 112) {
  const layout = layoutOf(layoutId);
  const ratio = content.size.h / content.size.w;
  const h = Math.round(width * ratio);
  const fake = { id: 'preview', layout: layout.id, elements: [] };
  const box = el('span', { class: 'mo-slides-layoutpic', style: { width: width + 'px', height: h + 'px', background: cssBackground(backgroundOf(content, fake), content.theme) } });
  const s = svgEl('svg', { width, height: h, viewBox: `0 0 ${width} ${h}`, 'aria-hidden': 'true' });
  for (const d of decorOf(content, fake)) {
    if (d.src) continue;
    const g = svgEl('g', { transform: `translate(${d.x * width} ${d.y * h})` });
    g.append(svgEl('path', { d: shapePath(d.shape || 'rect', d.w * width, d.h * h), fill: resolveColor(d.fill || 'accent1', content.theme), 'fill-opacity': d.opacity ?? 1 }));
    s.append(g);
  }
  const ink = resolveColor(content.theme.groups?.[layout.group]?.text || content.theme.text || 'dk1', content.theme);
  for (const slot of layout.slots) {
    s.append(svgEl('rect', { x: slot.x * width, y: slot.y * h, width: slot.w * width, height: slot.h * h, fill: 'none', stroke: ink, 'stroke-opacity': 0.55, 'stroke-dasharray': '2 1.5', 'stroke-width': 0.8 }));
    if (slot.role === 'title') s.append(svgEl('rect', { x: slot.x * width + 4, y: (slot.y + slot.h / 2) * h - 1.5, width: Math.min(slot.w * width - 8, 40), height: 3, fill: ink, 'fill-opacity': 0.75 }));
    else if (slot.role === 'body') for (let i = 0; i < Math.min(3, Math.floor(slot.h * h / 6)); i++) s.append(svgEl('rect', { x: slot.x * width + 5, y: slot.y * h + 4 + i * 5, width: Math.max(4, slot.w * width * (0.7 - i * 0.12)), height: 1.6, fill: ink, 'fill-opacity': 0.5 }));
    else if (slot.role === 'picture') s.append(svgEl('rect', { x: slot.x * width + slot.w * width / 2 - 5, y: slot.y * h + slot.h * h / 2 - 4, width: 10, height: 8, fill: 'none', stroke: ink, 'stroke-width': 0.8 }));
  }
  box.append(s);
  return box;
}

export function layoutGallery(anchor, content, current, pick) {
  return openPopover(anchor, close => {
    const box = el('div', { class: 'mo-slides-gallery mo-slides-gallery--layouts' });
    const own = THEMES.some(x => x.id === content.theme.id);
    box.append(el('div', { class: 'mo-menu__heading', text: own ? t('Slides.Theme.' + content.theme.id) : content.theme.name || t('Slides.Design') }));
    const grid = el('div', { class: 'mo-slides-gallery__layouts' });
    for (const l of LAYOUTS) {
      grid.append(el('button', { type: 'button', class: 'mo-slides-layoutbtn' + (l.id === current ? ' is-current' : ''), 'aria-pressed': l.id === current ? 'true' : 'false', onclick: () => { close(); pick(l.id); } },
        layoutPreview(content, l.id), el('span', { class: 'mo-slides-layoutbtn__label', text: t('Slides.Layout.' + l.id) })));
    }
    box.append(grid);
    return box;
  });
}

/** A picture of a design ("Aa" in its fonts and colours on its background) for the Design gallery. */
export function themePreview(content, themeId, variant = 0, width = 116) {
  const theme = makeTheme(themeId, variant);
  const fake = { ...content, theme, footer: null, size: content.size };
  const slide = newSlide(fake, 'title');
  slide.elements[0].text = { paras: [{ runs: [{ t: 'Aa', size: 140 }] }], autofit: 'none' };
  slide.elements.length = 1;
  const node = renderSlide(fake, slide, { mode: 'view' });
  const scale = width / content.size.w;
  const h = Math.round(content.size.h * scale);
  node.style.transform = `scale(${scale})`;
  node.style.transformOrigin = '0 0';
  const bars = el('span', { class: 'mo-slides-themepic__bars' });
  for (const k of ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']) bars.append(el('span', { style: { background: theme.colors[k] } }));
  return el('span', { class: 'mo-slides-themepic', style: { width: width + 'px', height: h + 'px' } }, node, bars);
}

export function variantPreview(content, variant, width = 64) {
  const theme = makeTheme(content.theme.id && THEMES.some(x => x.id === content.theme.id) ? content.theme.id : 'paper', variant);
  const bg = cssBackground(theme.background, theme);
  const box = el('span', { class: 'mo-slides-variantpic', style: { width: width + 'px', height: Math.round(width * content.size.h / content.size.w) + 'px', background: bg } });
  const bars = el('span', { class: 'mo-slides-themepic__bars mo-slides-themepic__bars--big' });
  for (const k of ['accent1', 'accent2', 'accent3', 'accent4']) bars.append(el('span', { style: { background: theme.colors[k] } }));
  box.append(el('span', { class: 'mo-slides-variantpic__aa', text: 'Aa', style: { color: resolveColor(theme.heading, theme), fontFamily: fontStack(theme.fonts.major) } }), bars);
  return box;
}


/** The quick styles of shapes as a grid. */
export function quickStyles(anchor, theme, pick) {
  return openPopover(anchor, close => {
    const box = el('div', { class: 'mo-slides-gallery mo-slides-gallery--styles' });
    box.append(el('div', { class: 'mo-menu__heading', text: t('Slides.QuickStyles') }));
    const grid = el('div', { class: 'mo-slides-gallery__styles' });
    for (const style of QUICK_STYLES) {
      const fill = style.fill.type === 'gradient' ? `linear-gradient(180deg, ${resolveColor(style.fill.from, theme)}, ${resolveColor(style.fill.to, theme)})` : resolveColor(style.fill.color, theme);
      const [base, look] = style.id.split('-');
      const label = t('Slides.QuickStyle', colorName(base), t('Slides.QuickStyle.' + look));
      const b = el('button', { type: 'button', class: 'mo-slides-stylebtn', 'aria-label': label, 'data-tip': label, onclick: () => { close(); pick(style); } },
        el('span', { class: 'mo-slides-stylebtn__swatch', text: 'Abc', style: { background: fill, color: resolveColor(style.text, theme), borderColor: style.line ? resolveColor(style.line.color, theme) : 'transparent', boxShadow: style.shadow ? '2px 2px 3px rgba(0,0,0,.35)' : 'none' } }));
      grid.append(b);
    }
    box.append(grid);
    return box;
  });
}

/** The grid to pick the size of a new table, and "Insert table …" for larger ones. */
export function tableGrid(anchor, pick, more) {
  return openPopover(anchor, close => {
    const box = el('div', { class: 'mo-slides-tablegrid' });
    const label = el('div', { class: 'mo-menu__heading', text: t('Slides.Table.Insert') });
    const grid = el('div', { class: 'mo-slides-tablegrid__cells', role: 'grid' });
    const cells = [];
    const mark = (r, c) => { cells.forEach(x => x.classList.toggle('is-on', x._r <= r && x._c <= c)); label.textContent = r >= 0 ? t('Slides.Table.Size', c + 1, r + 1) : t('Slides.Table.Insert'); };
    for (let r = 0; r < 8; r++) {
      for (let c = 0; c < 10; c++) {
        const cell = el('button', { type: 'button', class: 'mo-slides-tablegrid__cell', 'aria-label': t('Slides.Table.Size', c + 1, r + 1), onclick: () => { close(); pick(r + 1, c + 1); } });
        cell._r = r; cell._c = c;
        cell.addEventListener('mouseenter', () => mark(r, c));
        cell.addEventListener('focus', () => mark(r, c));
        cells.push(cell);
        grid.append(cell);
      }
    }
    box.append(label, grid, el('div', { class: 'mo-menu__sep' }), el('button', { type: 'button', class: 'mo-menu__item', onclick: () => { close(); more(); } }, icon('slides-table'), el('span', { class: 'mo-menu__label', text: t('Slides.Table.InsertMore') })));
    return box;
  });
}

const SYMBOLS = '© ® ™ € £ ¥ ¢ § ¶ † ‡ • … – — ‰ ° ± × ÷ ≈ ≠ ≤ ≥ ∞ √ ∑ ∏ ∫ ∂ ∆ π Ω µ α β γ δ ε λ σ φ ← → ↑ ↓ ↔ ⇒ ⇔ ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠ ☺ ☼ ♪ ✉ ☎ ⌘ ½ ¼ ¾ ¹ ² ³ « » „ “ ” ‚ ‘ ’'.split(' ');

export function symbolPicker(anchor, pick) {
  return openPopover(anchor, close => {
    const box = el('div', { class: 'mo-slides-symbols' });
    for (const s of SYMBOLS) box.append(el('button', { type: 'button', class: 'mo-slides-symbols__item', text: s, 'aria-label': s, onclick: () => { close(); pick(s); } }));
    return box;
  });
}

