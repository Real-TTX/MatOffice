/* MatOffice Docs – the commands of the ribbon and the keyboard (font, paragraph, styles, lists, case, painter, breaks, symbols …), built on the
   primitives of the engine (editor.js): formatText for characters, formatBlocks for paragraphs, insertBlocks / typeText for new content. */
import { readBlock, domLength } from './htmlio.js';
import {
  fontStack, stepSize, styleOf, styleKey, withStyle, isTextBlock, changeCase, nextCase, runsText, splitRuns, normalizeRuns, cleanFormat,
  blockLength, numStr, pt, toPt, textRun, sliceRuns, formatAt,
} from './model.js';

/** Line spacing values of the menu: Word's "multiple" → the CSS line height (× 1.2). */
export const LINE_SPACINGS = [[1, '1.2'], [1.15, '1.38'], [1.5, '1.8'], [2, '2.4'], [2.5, '3'], [3, '3.6']];

export function createCommands(app) {
  const ed = app.editor;
  const state = () => ed.state();

  const blockCss = (b, changes) => {
    const css = { ...b.css };
    for (const [k, v] of Object.entries(changes)) { if (v === null || v === undefined || v === '') delete css[k]; else css[k] = v; }
    return { ...b, css };
  };

  const c = {
    // ---- font ------------------------------------------------------------------------------------------------------------------
    setFont(name) { if (!name) return; ed.focus(); ed.formatText({ font: fontStack(name) }); },
    setSize(value) {
      const n = parseFloat(String(value).replace(',', '.'));
      if (!Number.isFinite(n) || n < 1 || n > 1638) return;
      ed.focus();
      ed.formatText({ size: pt(Math.round(n * 2) / 2) });
    },
    grow(dir) {
      ed.focus();
      const base = state().sizePt || 11;
      ed.formatText(f => {
        const now = f.size ? parseFloat(f.size) : base;
        return { ...f, size: pt(stepSize(now, dir)) };
      });
    },
    toggle(key) { ed.focus(); ed.toggle(key); },
    setColor(hex) { ed.focus(); ed.formatText({ color: hex || null }); },
    setHighlight(hex) { ed.focus(); ed.formatText({ bg: hex || null }, { word: false }); },
    changeCase(mode) {
      ed.focus();
      const sel = ed.selection();
      if (!sel || sel.collapsed) return;
      const ranges = ed.ranges(sel);
      if (!mode) {
        const text = ranges.map(({ el, start, end }) => runsText(sliceRuns(readBlock(el).runs, start, end))).join(' ');
        mode = text === text.toUpperCase() && text !== text.toLowerCase() ? 'lower' : text === text.toLowerCase() ? 'upper' : nextCase(text);
        if (mode === 'upper' && /\p{Lu}/u.test(text) && /\p{Ll}/u.test(text)) mode = 'upper';
      }
      ed.checkpoint('format');
      const mark = ed.bookmark(sel);
      const locale = app.language();
      for (const { el, start, end } of ranges) {
        if (end <= start) continue;
        const b = readBlock(el);
        const [head, rest] = splitRuns(b.runs, start);
        const [mid, tail] = splitRuns(rest, end - start);
        const text = mid.map(r => (r.obj ? '￼' : r.text)).join('');
        const changed = changeCase(text, mode, locale);
        let out;
        if (changed.length === text.length) {
          let pos = 0;
          out = mid.map(r => { if (r.obj) { pos += 1; return r; } const t = changed.slice(pos, pos + r.text.length); pos += r.text.length; return { ...r, text: t }; });
        } else out = mid.map(r => (r.obj ? r : { ...r, text: changeCase(r.text, mode, locale) }));
        b.runs = normalizeRuns([...head, ...out, ...tail]);
        ed.morph(el, b);
      }
      ed.restoreBookmark(mark, false);
    },
    /** Clear All Formatting: the paragraphs become Normal, the characters plain (links and comments stay). */
    clearFormatting() {
      ed.focus();
      const sel = ed.selection();
      if (!sel) return;
      ed.checkpoint('format');
      const mark = ed.bookmark(sel);
      for (const { el, start, end } of ed.ranges(sel)) {
        let b = readBlock(el);
        const whole = sel.collapsed || (start === 0 && end >= blockLength(b));
        if (whole || el !== sel.start.block || el !== sel.end.block) {
          b = withStyle(b, 'Normal');
          b.css = {};
          delete b.list; delete b.level; delete b.start;
        }
        const s = sel.collapsed ? 0 : start, e = sel.collapsed ? blockLength(b) : end;
        const [head, rest] = splitRuns(b.runs, s);
        const [mid, tail] = splitRuns(rest, e - s);
        b.runs = normalizeRuns([...head, ...mid.map(r => ({ ...r, f: cleanFormat({ href: r.f.href, cmt: r.f.cmt }) })), ...tail]);
        ed.morph(el, b);
      }
      ed.restoreBookmark(mark, false);
    },

    // ---- paragraph -------------------------------------------------------------------------------------------------------------------
    align(value) { ed.focus(); ed.formatBlocks(b => blockCss(b, { 'text-align': value === 'left' ? null : value })); },
    lineSpacing(value) { ed.focus(); ed.formatBlocks(b => blockCss(b, { 'line-height': value })); },
    spacing(which, value) {
      ed.focus();
      const key = which === 'before' ? 'margin-top' : 'margin-bottom';
      ed.formatBlocks(b => {
        const style = styleOf(styleKey(b));
        const def = which === 'before' ? style.before : style.after;
        const v = typeof value === 'function' ? value(currentSpacing(b, which)) : value;
        return blockCss(b, { [key]: v === null || Math.abs(v - def) < 0.01 ? null : pt(v) });
      });
    },
    indent(dir) {
      ed.focus();
      ed.formatBlocks(b => {
        if (b.list) return { ...b, level: Math.min(8, Math.max(0, (b.level || 0) + dir)) };
        const left = toPt(b.css['margin-left']) ?? (styleKey(b) === 'ListParagraph' ? 36 : styleKey(b) === 'Quote' || styleKey(b) === 'IntenseQuote' ? 43.2 : 0);
        const next = dir > 0 ? (Math.floor(left / 36 + 0.001) + 1) * 36 : Math.max(0, (Math.ceil(left / 36 - 0.001) - 1) * 36);
        return blockCss(b, { 'margin-left': next ? pt(next) : null });
      });
    },
    setIndents({ left, right, first }) {
      ed.focus();
      ed.formatBlocks(b => blockCss(b, {
        'margin-left': left === undefined ? b.css['margin-left'] : left ? pt(left) : null,
        'margin-right': right === undefined ? b.css['margin-right'] : right ? pt(right) : null,
        'text-indent': first === undefined ? b.css['text-indent'] : first ? pt(first) : null,
      }));
    },
    shading(hex) { ed.focus(); ed.formatBlocks(b => blockCss(b, { 'background-color': hex || null })); },
    borders(kind) {
      ed.focus();
      const line = '0.5pt solid #000000';
      ed.formatBlocks(b => {
        const has = side => !!b.css['border-' + side];
        const all = ['top', 'right', 'bottom', 'left'];
        const css = {};
        if (kind === 'none') for (const s of all) { css['border-' + s] = null; css['padding-' + s] = null; }
        else if (kind === 'box') { for (const s of all) css['border-' + s] = line; css['padding-top'] = '1pt'; css['padding-bottom'] = '1pt'; css['padding-left'] = '4pt'; css['padding-right'] = '4pt'; }
        else {
          css['border-' + kind] = has(kind) ? null : line;
          css['padding-' + kind] = has(kind) ? null : kind === 'left' || kind === 'right' ? '4pt' : '1pt';
        }
        return blockCss(b, css);
      });
    },

    // ---- lists ----------------------------------------------------------------------------------------------------------------------
    /** Bullets or numbering on the selected paragraphs; the same kind again takes it away. */
    list(kind) {
      ed.focus();
      const sel = ed.selection();
      if (!sel) return;
      const blocks = ed.ranges(sel).map(r => readBlock(r.el));
      const off = blocks.every(b => b.list === kind);
      ed.formatBlocks(b => {
        if (off) { const n = { ...b }; delete n.list; delete n.level; delete n.start; return n; }
        let n = b;
        if (b.type === 'blockquote' || b.type === 'pre' || (b.style && b.style !== 'ListParagraph' && b.type === 'p')) n = withStyle(b, 'Normal');
        n = { ...n, list: kind, level: b.list ? b.level || 0 : 0 };
        if (n.style === 'ListParagraph') delete n.style;
        const css = { ...n.css };
        delete css['margin-left'];
        delete css['text-indent'];
        return { ...n, css };
      }, 'list');
    },
    /** Starts the numbering of the list item at the caret anew (at n) – or continues it (n = null). */
    numberFrom(n) {
      ed.focus();
      const s = state();
      if (!s.blockEl || !s.block?.list) return;
      ed.checkpoint('list');
      const b = { ...s.block };
      if (n) b.start = n; else delete b.start;
      ed.morph(s.blockEl, b);
      ed.restoreBookmark(ed.bookmark(), false);
    },
    level(level) { ed.focus(); ed.formatBlocks(b => (b.list ? { ...b, level: Math.min(8, Math.max(0, level)) } : b)); },

    // ---- styles ---------------------------------------------------------------------------------------------------------------------
    style(key) {
      ed.focus();
      ed.formatBlocks(b => {
        const n = withStyle(b, key);
        if (key !== 'Normal' && key !== 'ListParagraph') { delete n.list; delete n.level; delete n.start; }
        return n;
      }, 'style');
    },

    // ---- breaks and objects ------------------------------------------------------------------------------------------------------------
    pageBreak(kind = 'page') {
      ed.focus();
      const sel = ed.selection();
      if (!sel || sel.start.block.closest('td,th') || sel.host !== app.body) return;
      ed.checkpoint('insert');
      ed.insertBlocks([{ type: 'break', kind }, { type: 'p', css: {}, runs: [] }]);
    },
    blankPage() {
      ed.focus();
      const sel = ed.selection();
      if (!sel || sel.start.block.closest('td,th') || sel.host !== app.body) return;
      ed.insertBlocks([{ type: 'break', kind: 'page' }, { type: 'p', css: {}, runs: [] }, { type: 'break', kind: 'page' }, { type: 'p', css: {}, runs: [] }]);
    },
    horizontalLine() {
      ed.focus();
      if (!ed.selection()) return;
      ed.insertBlocks([{ type: 'hr' }, { type: 'p', css: {}, runs: [] }]);
    },
    text(text) { ed.focus(); ed.checkpoint('insert'); ed.typeText(text); },
    field(kind) {
      ed.focus();
      const sel = ed.selection();
      if (!sel) return;
      const b = readBlock(sel.start.block);
      ed.insertBlocks([{ type: 'p', css: {}, runs: [{ obj: 'field', kind, f: cleanFormat(formatAt(b.runs, sel.start.offset)) }] }]);
    },
    selectAll() {
      const host = ed.selection()?.host || app.body;
      const blocks = ed.textBlocks(host);
      if (!blocks.length) return;
      ed.select(blocks[0], 0, blocks[blocks.length - 1], domLength(blocks[blocks.length - 1]), false);
    },
  };

  // ---- format painter: copy the format of the selection, apply it to the next selection (once, or until Escape after a double click) ----
  const painter = { active: false, sticky: false, inline: null, block: null };
  c.painter = painter;
  c.painterStart = (sticky = false) => {
    const s = state();
    if (!s.sel) return;
    if (painter.active) { c.painterStop(); return; }
    const runs = s.block?.runs || [];
    painter.inline = { ...formatAt(runs, s.sel.collapsed ? s.sel.start.offset : Math.min(blockLength(s.block), s.sel.start.offset + 1)) };
    delete painter.inline.href;
    delete painter.inline.cmt;
    painter.block = s.block ? { type: s.block.type, style: s.block.style, css: { ...s.block.css }, list: s.block.list, level: s.block.level } : null;
    painter.active = true;
    painter.sticky = sticky;
    app.body.classList.add('mo-docs-painting');
    app.toastInfo?.('Docs.Painter.Hint');
  };
  c.painterStop = () => {
    painter.active = false;
    app.body.classList.remove('mo-docs-painting');
    app.ribbonRefresh();
  };
  /** Called after a mouse selection while the painter is on. */
  c.painterApply = () => {
    if (!painter.active) return;
    const sel = ed.selection();
    if (!sel) return;
    const inline = painter.inline;
    if (!sel.collapsed) ed.formatText(() => ({ ...inline }), { word: false });
    else {
      const range = (() => { const b = readBlock(sel.start.block); const text = runsText(b.runs, ' '); let s = sel.start.offset, e = s; while (s > 0 && /[\p{L}\p{N}_]/u.test(text[s - 1])) s--; while (e < text.length && /[\p{L}\p{N}_]/u.test(text[e])) e++; return s < e ? [s, e] : null; })();
      if (range) { ed.select(sel.start.block, range[0], sel.start.block, range[1], false); ed.formatText(() => ({ ...inline }), { word: false }); }
      else if (painter.block) {
        const pb = painter.block;
        ed.formatBlocks(b => {
          const n = { ...b, type: isTextBlock({ type: pb.type }) ? pb.type : b.type, css: { ...pb.css } };
          if (pb.style) n.style = pb.style; else delete n.style;
          if (pb.list) { n.list = pb.list; n.level = pb.level || 0; } else { delete n.list; delete n.level; delete n.start; }
          return n;
        });
      }
    }
    if (!painter.sticky) c.painterStop();
  };

  return c;
}

/** The space before/after a paragraph has (its own or its style's), in pt. */
export function currentSpacing(b, which) {
  const key = which === 'before' ? 'margin-top' : 'margin-bottom';
  const own = toPt(b.css?.[key]);
  if (own !== null) return own;
  const s = styleOf(styleKey(b));
  return which === 'before' ? s.before : s.after;
}

/** The line height of a paragraph (CSS L) – its own or its style's. */
export function currentLineHeight(b) {
  return b?.css?.['line-height'] || String(styleOf(styleKey(b || { type: 'p' })).line);
}

export { numStr, textRun };
