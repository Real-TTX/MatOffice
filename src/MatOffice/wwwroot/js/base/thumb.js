/* MatOffice Base – the picture of a database for the lists (~480 × 360 PNG): the Navigation Pane with the objects and the datasheet of the first
   table, drawn on a canvas in the colours of Access. */
import { t } from '../editor/util.js';
import { displayValue } from './model.js';

const W = 480;
const H = 360;
const RED = '#a4373a';

export function thumbnail(app) {
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const g = canvas.getContext('2d');
  const db = app.db;
  const font = (size, weight = 400) => `${weight} ${size}px system-ui, "Segoe UI", sans-serif`;
  const clip = (text, width) => {
    let s = String(text ?? '');
    if (g.measureText(s).width <= width) return s;
    while (s.length && g.measureText(s + '…').width > width) s = s.slice(0, -1);
    return s + '…';
  };

  g.fillStyle = '#ffffff';
  g.fillRect(0, 0, W, H);

  // the navigation pane
  const navW = 150;
  g.fillStyle = '#f3f2f1';
  g.fillRect(0, 0, navW, H);
  g.fillStyle = '#e1dfdd';
  g.fillRect(navW - 1, 0, 1, H);
  g.fillStyle = RED;
  g.fillRect(0, 0, navW, 26);
  g.fillStyle = '#ffffff';
  g.font = font(12, 600);
  g.fillText(clip(t('Base.Nav.All'), navW - 16), 8, 17);
  let y = 42;
  const kinds = [['tables', t('Base.Nav.Tables')], ['queries', t('Base.Nav.Queries')], ['forms', t('Base.Nav.Forms')], ['reports', t('Base.Nav.Reports')]];
  for (const [kind, label] of kinds) {
    if (y > H - 20) break;
    g.fillStyle = '#605e5c';
    g.font = font(11, 600);
    g.fillText(label, 8, y);
    y += 6;
    for (const o of db[kind].slice(0, 6)) {
      y += 16;
      if (y > H - 8) break;
      g.fillStyle = kind === 'tables' ? RED : kind === 'queries' ? '#6b4c9a' : kind === 'forms' ? '#2b6cb0' : '#2f855a';
      g.fillRect(12, y - 9, 10, 10);
      g.fillStyle = '#323130';
      g.font = font(11);
      g.fillText(clip(o.name, navW - 36), 28, y);
    }
    y += 22;
  }

  // the datasheet of the first table
  const x0 = navW + 10;
  const table = db.tables[0];
  g.fillStyle = '#ffffff';
  g.fillRect(x0 - 10, 0, W - navW, H);
  if (!table) {
    g.fillStyle = '#a19f9d';
    g.font = font(14);
    g.fillText(clip(t('Base.Empty.NewTitle'), W - x0 - 10), x0, 60);
    return toBlob(canvas);
  }
  g.fillStyle = '#edebe9';
  g.fillRect(x0 - 10, 0, W - navW, 26);
  g.fillStyle = RED;
  g.fillRect(x0 - 2, 22, Math.min(140, g.measureText(table.name).width + 40), 4);
  g.fillStyle = '#323130';
  g.font = font(12, 600);
  g.fillText(clip(table.name, W - x0 - 20), x0 + 4, 17);
  const fields = table.fields.slice(0, 4);
  const selW = 14;
  const colW = Math.floor((W - x0 - selW - 6) / Math.max(1, fields.length));
  const rowH = 20;
  let top = 34;
  g.fillStyle = '#f3f2f1';
  g.fillRect(x0, top, W - x0, rowH);
  g.font = font(11, 600);
  g.fillStyle = '#323130';
  fields.forEach((f, i) => g.fillText(clip(f.caption || f.name, colW - 8), x0 + selW + 4 + i * colW, top + 14));
  g.strokeStyle = '#d2d0ce';
  g.lineWidth = 1;
  const lang = app.lang;
  const rows = table.rows.slice(0, Math.floor((H - top - rowH) / rowH));
  rows.forEach((r, k) => {
    top += rowH;
    if (k % 2) { g.fillStyle = '#faf9f8'; g.fillRect(x0, top, W - x0, rowH); }
    g.fillStyle = '#f3f2f1';
    g.fillRect(x0, top, selW, rowH);
    g.fillStyle = '#323130';
    g.font = font(11);
    fields.forEach((f, i) => {
      let text = '';
      try { text = f.type === 'yesno' ? (r.v[f.id] ? '☑' : '☐') : displayValue(db, f, r.v[f.id] ?? null, lang); } catch { text = ''; }
      const right = ['number', 'currency', 'autonumber'].includes(f.type);
      const cell = clip(text, colW - 8);
      const x = x0 + selW + i * colW;
      g.fillText(cell, right ? x + colW - 4 - g.measureText(cell).width : x + 4, top + 14);
    });
    g.beginPath();
    g.moveTo(x0, top + rowH + 0.5);
    g.lineTo(W, top + rowH + 0.5);
    g.stroke();
  });
  for (let i = 0; i <= fields.length; i++) {
    const x = Math.round(x0 + selW + i * colW) + 0.5;
    g.beginPath();
    g.moveTo(x, 34);
    g.lineTo(x, Math.min(H, top + rowH));
    g.stroke();
  }
  return toBlob(canvas);
}

function toBlob(canvas) {
  return new Promise(resolve => canvas.toBlob(b => resolve(b), 'image/png'));
}
