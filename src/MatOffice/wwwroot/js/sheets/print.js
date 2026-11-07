/* MatOffice Sheets – printing and the picture of the gallery. Print builds an HTML table of the used range of the active sheet (formats,
   merged cells, sizes, optional gridlines and headings, title rows repeated on every page, landscape, fit to the page width) with the charts
   and pictures at their places, prints it and removes it again. */
import { el } from '../editor/util.js';
import { colName } from './engine/address.js';
import { fontFamily, parseBorder, isDark } from './look.js';
import { chartData, renderChart } from './charts.js';

export function printSheet(app) {
  const sheet = app.sheet;
  const p = sheet.print || {};
  const b = sheet.bounds();
  const e = app.calc.extent(sheet.id);
  let rows = Math.max(b.rows, e.rows);
  let cols = Math.max(b.cols, e.cols);
  for (const o of [...sheet.charts, ...sheet.images]) { rows = Math.max(rows, sheet.rows.indexAt(o.y + o.h) + 1); cols = Math.max(cols, sheet.cols.indexAt(o.x + o.w) + 1); }
  rows = Math.min(rows, 20000);
  cols = Math.min(cols, 500);
  const titles = p.titles ? p.titles.split(':').map(n => Number(n) - 1) : null;

  const table = el('table', { class: 'mo-sheets-print__table' + (p.gridlines ? ' has-grid' : '') });
  const colgroup = el('colgroup');
  if (p.headings) colgroup.append(el('col', { style: { width: '36px' } }));
  for (let c = 0; c < cols; c++) if (!sheet.cols.isHidden(c)) colgroup.append(el('col', { style: { width: sheet.cols.size(c) + 'px' } }));
  table.append(colgroup);
  const thead = el('thead');
  const tbody = el('tbody');
  if (p.headings) {
    const head = el('tr', { class: 'mo-sheets-print__heads' }, el('th'));
    for (let c = 0; c < cols; c++) if (!sheet.cols.isHidden(c)) head.append(el('th', { text: colName(c) }));
    thead.append(head);
  }
  const covered = new Set();
  for (let r = 0; r < rows; r++) {
    if (sheet.rows.isHidden(r)) continue;
    const tr = el('tr', { style: { height: sheet.rows.size(r) + 'px' } });
    if (p.headings) tr.append(el('th', { text: String(r + 1) }));
    for (let c = 0; c < cols; c++) {
      if (sheet.cols.isHidden(c) || covered.has(r * 100000 + c)) continue;
      const look = app.look.cell(sheet, r, c, Math.max(1, Math.floor((sheet.cols.size(c) - 5) / 7)));
      const td = el('td');
      const st = look.style;
      const css = td.style;
      css.fontFamily = fontFamily(st.font);
      css.fontSize = (st.size || 11) + 'pt';
      if (look.b) css.fontWeight = '700';
      if (look.i) css.fontStyle = 'italic';
      if (look.u || look.strike) css.textDecoration = [look.u ? 'underline' : '', look.strike ? 'line-through' : ''].join(' ').trim();
      if (look.fill) css.background = look.fill;
      css.color = look.color || (look.fill && isDark(look.fill) ? '#ffffff' : '#000000');
      css.textAlign = look.ha === 'justify' ? 'justify' : look.ha;
      css.verticalAlign = look.va === 'middle' ? 'middle' : look.va;
      css.whiteSpace = look.wrap ? 'pre-wrap' : 'pre';
      if (look.indent) css.paddingLeft = (3 + look.indent * 9) + 'px';
      for (const [edge, prop] of [['bt', 'borderTop'], ['bb', 'borderBottom'], ['bl', 'borderLeft'], ['br', 'borderRight']]) {
        const bd = parseBorder(st[edge]);
        if (bd) css[prop] = `${bd.width}px ${bd.style === 'double' ? 'double' : bd.style === 'dashed' ? 'dashed' : bd.style === 'dotted' || bd.style === 'hair' ? 'dotted' : 'solid'} ${bd.color}`;
      }
      const m = sheet.mergeAt(r, c);
      if (m && m.r1 === r && m.c1 === c) {
        let rs = 0;
        let cs = 0;
        for (let rr = m.r1; rr <= m.r2; rr++) if (!sheet.rows.isHidden(rr)) rs++;
        for (let cc = m.c1; cc <= m.c2; cc++) if (!sheet.cols.isHidden(cc)) cs++;
        td.rowSpan = Math.max(1, rs);
        td.colSpan = Math.max(1, cs);
        for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) covered.add(rr * 100000 + cc);
      }
      if (look.bar) {
        const bar = el('span', { class: 'mo-sheets-print__bar' });
        bar.style.width = Math.round(look.bar.ratio * 100) + '%';
        bar.style.background = look.bar.color;
        td.append(bar);
      }
      td.append(document.createTextNode(look.text));
      tr.append(td);
    }
    if (titles && r >= titles[0] && r <= titles[1]) thead.append(tr); else tbody.append(tr);
  }
  table.append(thead, tbody);

  const wrap = el('div', { class: 'mo-sheets-print__sheet' }, table);
  const left = p.headings ? 36 : 0;
  const top = p.headings ? 20 : 0;
  for (const chart of sheet.charts) {
    const box = el('div', { class: 'mo-sheets-print__object' }, renderChart(chart, chartData(app, sheet, chart), { width: chart.w, height: chart.h, dark: false, locale: app.locale }));
    Object.assign(box.style, { left: left + chart.x + 'px', top: top + chart.y + 'px', width: chart.w + 'px', height: chart.h + 'px' });
    wrap.append(box);
  }
  for (const image of sheet.images) {
    const img = el('img', { class: 'mo-sheets-print__object', src: image.src, alt: image.alt || '' });
    Object.assign(img.style, { left: left + image.x + 'px', top: top + image.y + 'px', width: image.w + 'px', height: image.h + 'px' });
    wrap.append(img);
  }

  document.getElementById('mo-sheets-print')?.remove();
  const page = el('style', { id: 'mo-sheets-page' });
  page.textContent = `@page { size: ${p.landscape ? 'landscape' : 'portrait'}; margin: 12mm; }`;
  const host = el('div', { id: 'mo-sheets-print', class: 'mo-sheets-print' }, page, wrap);
  document.body.append(host);
  document.body.classList.add('mo-sheets-printing');
  // fit: the width (or the whole page) of the paper
  const paperW = (p.landscape ? 297 : 210) - 24;
  const paperH = (p.landscape ? 210 : 297) - 24;
  const pxPerMm = 96 / 25.4;
  const width = wrap.scrollWidth;
  const height = wrap.scrollHeight;
  let scale = 1;
  if (p.fit === 'width' || p.fit === 'page') scale = Math.min(1, (paperW * pxPerMm) / Math.max(1, width));
  if (p.fit === 'page') scale = Math.min(scale, (paperH * pxPerMm) / Math.max(1, height));
  if (scale < 1) wrap.style.zoom = String(Math.max(0.1, scale));
  const done = () => {
    document.body.classList.remove('mo-sheets-printing');
    host.remove();
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  setTimeout(() => {
    window.print();
    // some browsers do not send afterprint for a print that was cancelled
    setTimeout(() => { if (document.body.contains(host) && !window.matchMedia('print').matches) done(); }, 1000);
  }, 50);
}

/** The picture of the gallery: the top left of the first visible sheet, ~480 × 360, PNG. */
export function thumbnail(app) {
  const sheet = app.wb.sheets.find(s => !s.hidden) || app.sheet;
  const canvas = el('canvas', { width: 480, height: 360 });
  const g = canvas.getContext('2d');
  app.grid.renderTo(g, 480, 360, sheet, 1);
  // charts and pictures are drawn as boxes
  for (const o of [...sheet.images].slice(0, 10)) {
    const x = 34 + o.x;
    const y = 20 + o.y;
    if (x > 480 || y > 360) continue;
    const img = new Image();
    img.src = o.src;
    if (img.complete && img.naturalWidth) try { g.drawImage(img, x, y, o.w, o.h); } catch { /* not ready */ }
  }
  for (const c of sheet.charts.slice(0, 5)) {
    const x = 34 + c.x;
    const y = 20 + c.y;
    if (x > 480 || y > 360) continue;
    g.fillStyle = '#ffffff';
    g.strokeStyle = '#c8c8c8';
    g.fillRect(x, y, c.w, c.h);
    g.strokeRect(x + 0.5, y + 0.5, c.w - 1, c.h - 1);
    const colors = c.colors?.length ? c.colors : ['#4472c4', '#ed7d31', '#a5a5a5'];
    for (let i = 0; i < 6; i++) {
      g.fillStyle = colors[i % colors.length];
      const h = (c.h - 60) * (0.3 + 0.6 * ((i * 37) % 10) / 10);
      g.fillRect(x + 30 + i * ((c.w - 60) / 6), y + c.h - 20 - h, (c.w - 60) / 9, h);
    }
  }
  return new Promise(resolve => canvas.toBlob(b => resolve(b), 'image/png'));
}
