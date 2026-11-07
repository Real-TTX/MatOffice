/* MatOffice Sheets – copy, cut and paste. Inside the editor the cells travel with formulas (references follow), styles and notes; a cut moves
   them (references to them follow). Outside, the selection goes as text with tabs and as an HTML table with its formats (Word, other sheets,
   e-mail); pasted HTML tables (Excel, Google Sheets, web pages) and text with tabs are read like typed values. Pictures paste as pictures. */
import { t } from '../editor/util.js';
import { MAX_ROWS, MAX_COLS } from './engine/address.js';
import { moveFormula } from './engine/transform.js';
import { parseInput } from './engine/numfmt.js';
import { toCanonical } from './engine/lexer.js';
import { FormulaError } from './engine/values.js';
import { fontFamily, parseBorder } from './look.js';

export class Clipboard {
  constructor(app) {
    this.app = app;
    this.data = null;
  }

  // ---- copy ---------------------------------------------------------------------------------------------------------------------

  /** The selection as data: cells (with their effective style), texts, HTML. */
  capture(cut) {
    const app = this.app;
    const sheet = app.sheet;
    const a = app.commands.clip(app.sel.range);
    const rows = a.r2 - a.r1 + 1;
    const cols = a.c2 - a.c1 + 1;
    if (rows * cols > 1_000_000) { app.ctx.toast(t('Sheets.Clipboard.TooLarge'), 'warning'); return null; }
    const cells = [];
    const lines = [];
    for (let r = a.r1; r <= a.r2; r++) {
      const line = [];
      const texts = [];
      const hiddenRow = sheet.rows.isHidden(r);
      for (let c = a.c1; c <= a.c2; c++) {
        const cell = sheet.cell(r, c);
        const s = sheet.styleAt(r, c);
        line.push(cell || s ? { ...(cell || {}), s } : undefined);
        if (!hiddenRow && !sheet.cols.isHidden(c)) texts.push(this.textOf(sheet, r, c));
      }
      cells.push(line);
      if (!hiddenRow) lines.push(texts.join('\t'));
    }
    const text = lines.join('\r\n') + (rows * cols > 1 ? '\r\n' : '');
    const merges = sheet.merges.filter(m => m.r1 >= a.r1 && m.r2 <= a.r2 && m.c1 >= a.c1 && m.c2 <= a.c2).map(m => ({ r1: m.r1 - a.r1, c1: m.c1 - a.c1, r2: m.r2 - a.r1, c2: m.c2 - a.c1 }));
    const widths = [];
    for (let c = a.c1; c <= a.c2; c++) widths.push(sheet.cols.rawSize(c));
    return { sheetId: sheet.id, area: a, cells, text, html: this.html(sheet, a), cut, merges, widths };
  }

  /** What a cell shows, for the text of the clipboard (formatted like on the screen). */
  textOf(sheet, r, c) {
    const look = this.app.look.cell(sheet, r, c, 11);
    let text = this.app.showFormulas ? look.text : look.text;
    if (look.isNumber && /#{3,}/.test(text)) text = String(look.value);
    return /[\t\r\n"]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }

  html(sheet, a) {
    const app = this.app;
    const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    let out = '<table style="border-collapse:collapse">';
    const covered = new Set();
    for (let r = a.r1; r <= a.r2; r++) {
      if (sheet.rows.isHidden(r)) continue;
      out += `<tr style="height:${Math.round(sheet.rows.size(r) * 0.75)}pt">`;
      for (let c = a.c1; c <= a.c2; c++) {
        if (sheet.cols.isHidden(c) || covered.has(r + ':' + c)) continue;
        const look = app.look.cell(sheet, r, c, 11);
        const st = look.style;
        const css = [`font-family:${esc(fontFamily(st.font))}`, `font-size:${st.size || 11}pt`];
        if (look.b) css.push('font-weight:bold');
        if (look.i) css.push('font-style:italic');
        if (look.u || look.strike) css.push(`text-decoration:${[look.u ? 'underline' : '', look.strike ? 'line-through' : ''].join(' ').trim()}`);
        if (look.color) css.push(`color:${look.color}`);
        if (look.fill) css.push(`background:${look.fill}`);
        css.push(`text-align:${look.ha === 'justify' ? 'justify' : look.ha}`, `vertical-align:${look.va === 'middle' ? 'middle' : look.va}`);
        if (look.wrap) css.push('white-space:normal'); else css.push('white-space:nowrap');
        for (const [edge, prop] of [['bt', 'top'], ['bb', 'bottom'], ['bl', 'left'], ['br', 'right']]) {
          const b = parseBorder(st[edge]);
          if (b) css.push(`border-${prop}:${b.width}px ${b.style === 'double' ? 'double' : b.style === 'dashed' ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid'} ${b.color}`);
        }
        const m = sheet.mergeAt(r, c);
        let span = '';
        if (m && m.r1 === r && m.c1 === c) {
          span = ` rowspan="${Math.min(m.r2, a.r2) - r + 1}" colspan="${Math.min(m.c2, a.c2) - c + 1}"`;
          for (let rr = r; rr <= Math.min(m.r2, a.r2); rr++) for (let cc = c; cc <= Math.min(m.c2, a.c2); cc++) if (rr !== r || cc !== c) covered.add(rr + ':' + cc);
        }
        out += `<td${span} style="${css.join(';')}">${esc(look.text).replace(/\n/g, '<br>')}</td>`;
      }
      out += '</tr>';
    }
    return out + '</table>';
  }

  /** The copy and cut events of the browser (Ctrl+C, Ctrl+X on the grid). */
  onCopy(e, cut) {
    if (this.app.editor.active) return;
    if (cut && this.app.readOnly) { e.preventDefault(); this.app.toastReadOnly(); return; }
    if (this.app.sel.ranges.length > 1) { e.preventDefault(); this.app.ctx.toast(t('Sheets.Clipboard.Multiple'), 'warning'); return; }
    const data = this.capture(cut);
    if (!data) { e.preventDefault(); return; }
    e.preventDefault();
    e.clipboardData.setData('text/plain', data.text);
    e.clipboardData.setData('text/html', data.html);
    this.data = data;
    this.app.grid.setMarquee({ sheet: data.sheetId, area: data.area });
  }

  copyCommand() { this.command(false); }
  cutCommand() { if (this.app.readOnly) return; this.command(true); }

  command(cut) {
    const app = this.app;
    if (app.editor.active) { document.execCommand(cut ? 'cut' : 'copy'); return; }
    app.focusGrid(true);
    let ok = false;
    try { ok = document.execCommand(cut ? 'cut' : 'copy'); } catch { ok = false; }
    if (!ok) {
      const data = this.capture(cut);
      if (!data) return;
      this.data = data;
      app.grid.setMarquee({ sheet: data.sheetId, area: data.area });
      navigator.clipboard?.writeText(data.text).catch(() => {});
    }
  }

  clearMarquee() {
    this.app.grid.setMarquee(null);
    if (this.data?.cut) this.data = null;
  }

  // ---- paste ----------------------------------------------------------------------------------------------------------------------

  /** The paste event of the browser (Ctrl+V on the grid). */
  onPaste(e) {
    const app = this.app;
    if (app.editor.active) return;
    e.preventDefault();
    if (app.readOnly) { app.toastReadOnly(); return; }
    const files = [...(e.clipboardData?.files || [])].filter(f => /^image\/(png|jpeg|gif|webp)$/.test(f.type));
    const text = e.clipboardData?.getData('text/plain') || '';
    const html = e.clipboardData?.getData('text/html') || '';
    if (files.length && !text) { app.objects.insertPicture(files[0]); return; }
    this.paste({ text, html });
  }

  /** The Paste button: the clipboard of the system if the browser allows reading it, else the cells copied in the editor. */
  async pasteCommand(mode) {
    const app = this.app;
    if (app.readOnly) return;
    if (mode && this.data) { this.pasteInternal(this.data, mode); return; }
    let text = '';
    let html = '';
    try {
      if (navigator.clipboard?.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          const image = item.types.find(tp => /^image\/(png|jpeg|gif|webp)$/.test(tp));
          if (image && !item.types.includes('text/plain')) { const blob = await item.getType(image); app.objects.insertPicture(new File([blob], 'image', { type: image })); return; }
          if (item.types.includes('text/html')) html = await (await item.getType('text/html')).text();
          if (item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
        }
      } else if (navigator.clipboard?.readText) {
        text = await navigator.clipboard.readText();
      }
    } catch {
      if (this.data) { this.pasteInternal(this.data, mode); return; }
      app.ctx.toast(t('Sheets.Clipboard.UseKeys'), 'info');
      return;
    }
    if (!text && !html && this.data) { this.pasteInternal(this.data, mode); return; }
    this.paste({ text, html });
  }

  paste({ text, html }) {
    const norm = s => String(s || '').replace(/\r\n/g, '\n').replace(/\n$/, '');
    if (this.data && norm(text) === norm(this.data.text)) { this.pasteInternal(this.data); return; }
    const table = html ? readHtmlTable(html) : null;
    if (table && table.length) { this.pasteValues(table); return; }
    if (!text) return;
    this.pasteValues(splitText(text).map(line => line.map(v => ({ text: v }))));
  }

  /** The target area of a paste: the selection when it is a multiple of the source, else the source size at the active cell. */
  target(rows, cols) {
    const a = this.app.sel.range;
    const h = a.r2 - a.r1 + 1;
    const w = a.c2 - a.c1 + 1;
    if (!(this.app.sel.isWholeCol(a) || this.app.sel.isWholeRow(a)) && h % rows === 0 && w % cols === 0 && h * w <= 1_000_000) return { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
    return { r1: a.r1, c1: a.c1, r2: Math.min(MAX_ROWS - 1, a.r1 + rows - 1), c2: Math.min(MAX_COLS - 1, a.c1 + cols - 1) };
  }

  /** Cells copied in the editor: mode undefined (all), 'values', 'formulas', 'formats', 'transpose', 'widths'. */
  pasteInternal(data, mode, options = {}) {
    const app = this.app;
    const sheet = app.sheet;
    const from = app.wb.sheetById(data.sheetId);
    const transpose = mode === 'transpose' || options.transpose;
    const rows = transpose ? data.cells[0].length : data.cells.length;
    const cols = transpose ? data.cells.length : data.cells[0].length;
    const target = this.target(rows, cols);
    if (data.cut && from && !mode) {
      const dr = target.r1 - data.area.r1;
      const dc = target.c1 - data.area.c1;
      app.commands.change('move', () => app.wb.moveCells(from, data.area, sheet, dr, dc), { structural: true });
      this.data = null;
      app.grid.setMarquee(null);
      app.sel.setArea({ r1: target.r1, c1: target.c1, r2: target.r1 + rows - 1, c2: target.c1 + cols - 1 });
      app.selectionChanged();
      return;
    }
    app.commands.change('paste', () => {
      const wb = app.wb;
      for (let r = target.r1; r <= target.r2; r++) {
        for (let c = target.c1; c <= target.c2; c++) {
          const i = (r - target.r1) % rows;
          const j = (c - target.c1) % cols;
          const src = transpose ? data.cells[j][i] : data.cells[i][j];
          const srcRow = data.area.r1 + (transpose ? j : i);
          const srcCol = data.area.c1 + (transpose ? i : j);
          const existing = sheet.cell(r, c);
          if (mode === 'widths') continue;
          if (mode === 'formats') { wb.patchCell(sheet, r, c, { s: src?.s ?? 0 }); continue; }
          if (mode === 'values') {
            const v = from ? app.calc.valueAt(from.id, srcRow, srcCol) : src?.v;
            const isErr = v instanceof FormulaError;
            wb.setCell(sheet, r, c, { ...(existing || {}), v: v === null || v === undefined ? undefined : isErr ? v.code : v, t: isErr ? 'e' : undefined, f: undefined });
            continue;
          }
          if (mode === 'formulas') {
            const f = src?.f ? moveFormula(src.f, r - srcRow, c - srcCol) : undefined;
            wb.setCell(sheet, r, c, { ...(existing || {}), v: src?.v, t: src?.t, f });
            continue;
          }
          if (!src) { wb.setCell(sheet, r, c, undefined); continue; }
          const next = { ...src };
          if (next.f) next.f = moveFormula(next.f, r - srcRow, c - srcCol);
          if (options.skipBlanks && next.v === undefined && !next.f) continue;
          wb.setCell(sheet, r, c, next);
        }
      }
      if (mode === 'widths' || (!mode && data.widths && app.sel.isWholeCol())) {
        wb.touchProp(sheet, 'cols');
        data.widths.forEach((w, j) => sheet.cols.setSize(target.c1 + j, w));
      }
      if (!mode && data.merges?.length && !transpose) {
        const merges = sheet.merges.filter(m => !(m.r1 <= target.r2 && target.r1 <= m.r2 && m.c1 <= target.c2 && target.c1 <= m.c2));
        for (const m of data.merges) merges.push({ r1: m.r1 + target.r1, c1: m.c1 + target.c1, r2: m.r2 + target.r1, c2: m.c2 + target.c1 });
        wb.setSheetProp(sheet, 'merges', merges);
      }
    });
    app.sel.setArea(target);
    app.selectionChanged();
  }

  /** Values from outside (a table of { text, style }) are read like typed input. */
  pasteValues(table) {
    const app = this.app;
    const sheet = app.sheet;
    const rows = table.length;
    const cols = Math.max(...table.map(r => r.length));
    if (!rows || !cols) return;
    const single = rows === 1 && cols === 1;
    const target = single ? app.commands.clip(app.sel.range) : this.target(rows, cols);
    if ((target.r2 - target.r1 + 1) * (target.c2 - target.c1 + 1) > 2_000_000) { app.ctx.toast(t('Sheets.Clipboard.TooLarge'), 'warning'); return; }
    app.commands.change('paste', () => {
      const wb = app.wb;
      for (let r = target.r1; r <= target.r2; r++) {
        for (let c = target.c1; c <= target.c2; c++) {
          const src = table[(r - target.r1) % rows]?.[(c - target.c1) % cols];
          const existing = sheet.cell(r, c) || {};
          const next = { ...existing, v: undefined, f: undefined, t: undefined };
          const parsed = parseInput(src?.text ?? '', app.locale);
          let s = existing.s ?? sheet.styleAt(r, c);
          if (parsed.formula !== undefined) next.f = toCanonical(parsed.formula, app.locale);
          else if (parsed.error) { next.v = parsed.error; next.t = 'e'; } else if (!parsed.empty) {
            next.v = parsed.value;
            if (parsed.nf && !wb.style(s).nf) s = wb.patchStyle(s, { nf: parsed.nf });
          }
          if (src?.style) s = wb.patchStyle(s, src.style);
          if (s) next.s = s;
          wb.setCell(sheet, r, c, next);
        }
      }
    });
    app.sel.setArea(target);
    app.selectionChanged();
    app.autoFitRows([target]);
  }
}

/** Text with tabs and lines (and quoted fields, like Excel writes them) → rows of strings. */
export function splitText(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const s = String(text).replace(/\r\n?/g, '\n');
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (quoted) {
      if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else quoted = false; } else field += ch;
      continue;
    }
    if (ch === '"' && field === '') { quoted = true; continue; }
    if (ch === '\t') { row.push(field); field = ''; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; continue; }
    field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** The first table of pasted HTML → rows of { text, style } (bold, italic, underline, colours, alignment of the cells). */
export function readHtmlTable(html) {
  let doc;
  try { doc = new DOMParser().parseFromString(html, 'text/html'); } catch { return null; }
  const table = doc.querySelector('table');
  if (!table) return null;
  // the class rules of Excel's HTML (<style> .xl65 { … })
  const classes = new Map();
  for (const st of doc.querySelectorAll('style')) {
    for (const m of st.textContent.matchAll(/\.([\w-]+)\s*\{([^}]*)\}/g)) classes.set(m[1], m[2]);
  }
  const rows = [];
  const spans = [];
  for (const tr of table.querySelectorAll('tr')) {
    const row = rows.length;
    rows[row] = rows[row] || [];
    let col = 0;
    for (const td of tr.querySelectorAll('td, th')) {
      while (spans[row]?.[col]) col++;
      const text = cellText(td);
      const css = [...td.classList].map(c => classes.get(c) || '').join(';') + ';' + (td.getAttribute('style') || '');
      rows[row][col] = { text, style: styleOf(css, td) };
      const rs = Number(td.getAttribute('rowspan')) || 1;
      const cs = Number(td.getAttribute('colspan')) || 1;
      for (let r = 0; r < rs; r++) for (let c = 0; c < cs; c++) if (r || c) { spans[row + r] = spans[row + r] || []; spans[row + r][col + c] = true; }
      col += cs;
    }
  }
  return rows.map(r => Array.from({ length: r.length }, (_, i) => r[i] || { text: '' }));
}

function cellText(td) {
  let out = '';
  const walk = node => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) out += n.nodeValue.replace(/\s+/g, ' ');
      else if (n.nodeName === 'BR') out += '\n';
      else if (n.nodeType === 1 && !/^(STYLE|SCRIPT)$/.test(n.nodeName)) { walk(n); if (/^(P|DIV)$/.test(n.nodeName) && n.nextSibling) out += '\n'; }
    }
  };
  walk(td);
  return out.replace(/ /g, ' ').trim();
}

function styleOf(css, td) {
  const s = {};
  const get = prop => { const m = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([^;]+)', 'i').exec(css); return m ? m[1].trim() : null; };
  const weight = get('font-weight');
  if (weight && (/bold/i.test(weight) || Number(weight) >= 600) || td.querySelector('b, strong')) s.b = true;
  if (/italic/i.test(get('font-style') || '') || td.querySelector('i, em')) s.i = true;
  const deco = get('text-decoration') || get('text-decoration-line') || '';
  if (/underline/i.test(deco) || td.querySelector('u')) s.u = true;
  if (/line-through/i.test(deco) || td.querySelector('s, strike, del')) s.strike = true;
  const color = hex(get('color'));
  if (color && color !== '#000000') s.color = color;
  const bg = hex(get('background-color') || get('background'));
  if (bg && bg !== '#ffffff') s.fill = bg;
  const align = (get('text-align') || '').toLowerCase();
  if (['center', 'right', 'justify'].includes(align)) s.ha = align;
  return Object.keys(s).length ? s : null;
}

function hex(value) {
  if (!value) return null;
  const v = value.trim().toLowerCase();
  let m = /^#([0-9a-f]{6})\b/.exec(v);
  if (m) return '#' + m[1];
  m = /^#([0-9a-f]{3})\b/.exec(v);
  if (m) return '#' + m[1].split('').map(x => x + x).join('');
  m = /rgba?\(\s*(\d+)[ ,]+(\d+)[ ,]+(\d+)(?:[ ,/]+([\d.]+))?/.exec(v);
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return null;
    return '#' + [m[1], m[2], m[3]].map(n => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('');
  }
  const named = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00' };
  return named[v] || null;
}
