/* MatOffice Docs – the picture of the document in the lists: the first page (header, text, footer) drawn through an SVG foreignObject into a
   PNG of 360 × 480 (480 × 360 for landscape). If the browser does not allow that, a simplified page with grey lines where the text is. */
import { numStr } from './model.js';

let cssCache = null;

async function stylesheet() {
  if (cssCache !== null) return cssCache;
  const link = [...document.querySelectorAll('link[rel="stylesheet"]')].find(l => /\/css\/docs\.css/.test(l.href));
  try { cssCache = link ? await (await fetch(link.href, { credentials: 'same-origin' })).text() : ''; } catch { cssCache = ''; }
  return cssCache;
}

export async function thumbnail(app) {
  const g = app.view.geometry;
  if (!g) return null;
  const W = g.paperWidth, H = g.paperHeight;
  const landscape = W > H;
  const outW = landscape ? 480 : 360, outH = landscape ? 360 : 480;
  const scale = Math.min(outW / W, outH / H);
  const blocks = [];
  for (const n of app.body.children) {
    const r = n.getClientRects()[0];
    if (!r) continue;
    if (app.viewMode === 'print' && app.view.pageOfRect(r) > 0) break;
    blocks.push(n);
    if (blocks.length > 80) break;
  }
  try {
    return await rendered(app, g, blocks, outW, outH, scale);
  } catch {
    return simple(app, g, blocks, outW, outH, scale);
  }
}

async function rendered(app, g, blocks, outW, outH, scale) {
  const css = await stylesheet();
  const xml = new XMLSerializer();
  const sheet = app.view.sheets.firstElementChild;
  const part = sel => (sheet ? [...(sheet.querySelector(sel)?.childNodes || [])].map(n => xml.serializeToString(n)).join('') : '');
  const px = v => numStr(v) + 'px';
  const vars = `--pw:${px(g.paperWidth)};--ph:${px(g.paperHeight)};--mt:${px(g.top)};--mb:${px(g.bottom)};--ml:${px(g.left)};--mr:${px(g.right)};--tw:${px(g.textWidth)};--th:${px(g.textHeight)};--hd:${px(g.header)};--fd:${px(g.footer)};--cgap:${px(g.columnGap)}`;
  const columns = g.columns > 1 ? `column-count:${g.columns};column-gap:${px(g.columnGap)};column-fill:auto;` : '';
  const body = blocks.map(n => xml.serializeToString(n)).join('');
  const page = `<div xmlns="http://www.w3.org/1999/xhtml" style="${vars};position:relative;width:${px(g.paperWidth)};height:${px(g.paperHeight)};overflow:hidden;background:#fff">`
    + `<style>${css.replace(/<\/?style/gi, '')}</style>`
    + `<div class="mo-docs-hdr mo-docs-content">${part('.mo-docs-hdr')}</div>`
    + `<div class="mo-docs-content" style="position:absolute;top:${px(g.top)};left:${px(g.left)};width:${px(g.textWidth)};height:${px(g.textHeight)};overflow:hidden;${columns}">${body}</div>`
    + `<div class="mo-docs-ftr mo-docs-content">${part('.mo-docs-ftr')}</div></div>`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(g.paperWidth)}" height="${Math.ceil(g.paperHeight)}"><foreignObject x="0" y="0" width="100%" height="100%">${page}</foreignObject></svg>`;
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
  try {
    const img = await new Promise((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = url; });
    const canvas = document.createElement('canvas');
    canvas.width = outW;
    canvas.height = outH;
    const c = canvas.getContext('2d');
    c.fillStyle = '#ffffff';
    c.fillRect(0, 0, outW, outH);
    const w = g.paperWidth * scale, h = g.paperHeight * scale;
    c.drawImage(img, (outW - w) / 2, (outH - h) / 2, w, h);
    return await new Promise((resolve, reject) => canvas.toBlob(b => (b ? resolve(b) : reject(new Error('blob'))), 'image/png'));
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** A drawn page: grey bars for the lines of text, blue for headings, boxes for pictures and tables. */
function simple(app, g, blocks, outW, outH, scale) {
  const canvas = document.createElement('canvas');
  canvas.width = outW;
  canvas.height = outH;
  const c = canvas.getContext('2d');
  c.fillStyle = '#ffffff';
  c.fillRect(0, 0, outW, outH);
  const ox = (outW - g.paperWidth * scale) / 2, oy = (outH - g.paperHeight * scale) / 2;
  for (const n of blocks) {
    for (const r of n.getClientRects()) {
      const l = app.view.local(r);
      const y = l.top, x = l.left;
      if (y > g.paperHeight) break;
      c.fillStyle = /^H[1-6]$/.test(n.tagName) ? '#2f5496' : n.tagName === 'TABLE' ? '#bfbfbf' : '#9a9a9a';
      const lines = Math.max(1, Math.round((l.bottom - l.top) / 16));
      for (let i = 0; i < lines; i++) c.fillRect(ox + x * scale, oy + (y + i * 16 + 4) * scale, Math.max(2, (l.right - l.left) * scale * (i === lines - 1 ? 0.7 : 1)), Math.max(1, 8 * scale));
    }
  }
  return new Promise(resolve => canvas.toBlob(b => resolve(b), 'image/png'));
}
