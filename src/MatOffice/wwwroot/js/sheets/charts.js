/* MatOffice Sheets – charts drawn as SVG from a range of the sheet: column, bar (clustered or stacked), line, area, pie, doughnut and
   scatter; title, legend, axis titles, data labels, the colours of the Office theme. The data is read like Excel reads it: a first row of
   texts names the series, a first column of texts (or dates) labels the categories; byRows turns it around. */
import { parseReference } from './engine/address.js';
import { formatValue, isDateFormat } from './engine/numfmt.js';
import { FormulaError } from './engine/values.js';

export const CHART_COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e', '#636363', '#997300', '#255e91', '#43682b'];
export const CHART_TYPES = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];
const NS = 'http://www.w3.org/2000/svg';

/** The series of a chart: { categories: [text], series: [{ name, values: [number|null], xs?: [number] }], format }. */
export function chartData(app, sheet, chart) {
  const ref = parseReference(chart.range || '');
  if (!ref) return { categories: [], series: [], format: null };
  const src = ref.sheet ? app.wb.sheetByName(ref.sheet) : sheet;
  if (!src) return { categories: [], series: [], format: null };
  const a = app.commands.clip(ref.area, src);
  const value = (r, c) => { const v = app.calc.valueAt(src.id, r, c); return v instanceof FormulaError ? null : v; };
  const text = (r, c) => app.look.cell(src, r, c, 11).text;
  let rows = [];
  for (let r = a.r1; r <= a.r2; r++) { if (src.rows.isHidden(r)) continue; const line = []; for (let c = a.c1; c <= a.c2; c++) { if (!src.cols.isHidden(c)) line.push({ r, c }); } rows.push(line); }
  if (!rows.length || !rows[0].length) return { categories: [], series: [], format: null };
  if (chart.byRows) rows = rows[0].map((_, j) => rows.map(line => line[j]));
  const isText = p => { const v = value(p.r, p.c); return typeof v === 'string' || v === null; };
  const first = rows[0];
  const headerRow = rows.length > 1 && first.slice(1).some(p => typeof value(p.r, p.c) === 'string') || (rows.length > 1 && value(first[0].r, first[0].c) === null && first.length > 1 && first.slice(1).every(p => isText(p)));
  const body = headerRow ? rows.slice(1) : rows;
  const dateFirst = body.length && body.every(line => isDateFormat(app.wb.style(src.styleAt(line[0].r, line[0].c)).nf));
  const labelCol = first.length > 1 && (body.every(line => isText(line[0])) || (headerRow && value(first[0].r, first[0].c) === null) || dateFirst) && chart.type !== 'scatter';
  const scatter = chart.type === 'scatter';
  const startCol = labelCol || scatter && first.length > 1 ? 1 : 0;
  const categories = body.map((line, i) => (labelCol ? text(line[0].r, line[0].c) : String(i + 1)));
  const series = [];
  for (let j = startCol; j < first.length; j++) {
    const name = headerRow ? text(first[j].r, first[j].c) : `Series ${series.length + 1}`;
    const values = body.map(line => { const v = value(line[j].r, line[j].c); return typeof v === 'number' ? v : null; });
    const s = { name, values };
    if (scatter && first.length > 1) s.xs = body.map(line => { const v = value(line[0].r, line[0].c); return typeof v === 'number' ? v : null; });
    series.push(s);
  }
  const sample = body[0]?.[startCol];
  const format = sample ? app.wb.style(src.styleAt(sample.r, sample.c)).nf || null : null;
  return { categories, series, format, xFormat: body[0]?.[0] ? app.wb.style(src.styleAt(body[0][0].r, body[0][0].c)).nf : null };
}

function svg(tag, attrs = {}, text) {
  const node = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined && v !== null) node.setAttribute(k, String(v));
  if (text !== undefined) node.textContent = text;
  return node;
}

/** Ticks of a value axis: { min, max, step }. */
export function niceScale(min, max, count = 5) {
  if (!(Number.isFinite(min) && Number.isFinite(max))) return { min: 0, max: 1, step: 0.2 };
  if (min === max) { if (min === 0) return { min: 0, max: 1, step: 0.2 }; min = Math.min(0, min); max = Math.max(0, max); }
  const span = max - min;
  const raw = span / count;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

/**
 * The SVG of a chart. options: { width, height, dark, locale }.
 */
export function renderChart(chart, data, { width, height, dark = false, locale = 'en' }) {
  const fg = dark ? '#e6e6e6' : '#404040';
  const muted = dark ? '#a0a0a0' : '#767676';
  const line = dark ? '#4a4a4a' : '#d9d9d9';
  const root = svg('svg', { viewBox: `0 0 ${width} ${height}`, width, height, class: 'mo-sheets-chart', role: 'img', 'aria-label': chart.title || chart.type, 'font-family': 'Segoe UI, system-ui, sans-serif' });
  root.append(svg('rect', { x: 0, y: 0, width, height, fill: dark ? '#262626' : '#ffffff' }));
  const colors = chart.colors?.length ? [...chart.colors, ...CHART_COLORS] : CHART_COLORS;
  const fmt = (v, code) => formatValue(v, code || undefined, { locale, width: 8 }).text;
  let top = 10;
  if (chart.title) {
    root.append(svg('text', { x: width / 2, y: 26, 'text-anchor': 'middle', 'font-size': 15, fill: fg }, chart.title));
    top = 40;
  }
  const pie = chart.type === 'pie' || chart.type === 'doughnut';
  const legendItems = pie ? data.categories.map((c, i) => ({ name: c, color: colors[i % colors.length] })) : data.series.map((s, i) => ({ name: s.name, color: colors[i % colors.length] }));
  const legend = chart.legend || (legendItems.length > 1 || pie ? 'right' : 'none');
  const plot = { x: 10, y: top, w: width - 20, h: height - top - 10 };
  if (legend !== 'none' && legendItems.length) {
    const g = svg('g', { class: 'mo-sheets-chart__legend', 'font-size': 11, fill: fg });
    if (legend === 'right' || legend === 'left') {
      const lw = Math.min(width * 0.32, 20 + Math.max(...legendItems.map(it => String(it.name).length)) * 6.2);
      const lx = legend === 'right' ? width - lw - 4 : 8;
      let ly = plot.y + Math.max(0, (plot.h - legendItems.length * 18) / 2);
      for (const it of legendItems.slice(0, 30)) {
        g.append(svg('rect', { x: lx, y: ly + 2, width: 10, height: 10, fill: it.color }));
        g.append(svg('text', { x: lx + 15, y: ly + 11 }, clip(it.name, Math.floor((lw - 20) / 6))));
        ly += 18;
      }
      if (legend === 'right') plot.w -= lw + 6; else { plot.x += lw + 6; plot.w -= lw + 6; }
    } else {
      const ly = legend === 'top' ? plot.y : height - 18;
      let total = 0;
      const widths = legendItems.map(it => 18 + Math.min(20, String(it.name).length) * 6.2 + 12);
      widths.forEach(w => { total += w; });
      let lx = Math.max(8, (width - total) / 2);
      legendItems.forEach((it, i) => {
        g.append(svg('rect', { x: lx, y: ly + 2, width: 10, height: 10, fill: it.color }));
        g.append(svg('text', { x: lx + 15, y: ly + 11 }, clip(it.name, 20)));
        lx += widths[i];
      });
      if (legend === 'top') { plot.y += 22; plot.h -= 22; } else plot.h -= 24;
    }
    root.append(g);
  }
  if (!data.series.length) {
    root.append(svg('text', { x: width / 2, y: height / 2, 'text-anchor': 'middle', 'font-size': 12, fill: muted }, '—'));
    return root;
  }
  if (pie) { drawPie(root, chart, data, plot, colors, fg, dark, fmt); return root; }

  // axes
  const horizontal = chart.type === 'bar';
  const stacked = !!chart.stacked && chart.type !== 'scatter';
  const n = data.categories.length;
  let min = Infinity;
  let max = -Infinity;
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of data.series) { const v = s.values[i] || 0; if (v >= 0) pos += v; else neg += v; }
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    }
  } else {
    for (const s of data.series) for (const v of s.values) if (v !== null) { min = Math.min(min, v); max = Math.max(max, v); }
  }
  if (chart.type !== 'line' && chart.type !== 'scatter') { min = Math.min(0, min); max = Math.max(0, max); }
  if (!Number.isFinite(min)) { min = 0; max = 1; }
  const scale = niceScale(min, max);
  let xScale = null;
  if (chart.type === 'scatter') {
    let xmin = Infinity;
    let xmax = -Infinity;
    for (const s of data.series) for (const x of s.xs || s.values.map((_, i) => i + 1)) if (x !== null) { xmin = Math.min(xmin, x); xmax = Math.max(xmax, x); }
    xScale = niceScale(xmin, xmax);
  }
  // room for the labels and axis titles
  const labelW = Math.max(...[scale.min, scale.max, scale.min + scale.step].map(v => fmt(v, data.format).length)) * 6.5 + 8;
  const left = (horizontal ? Math.min(plot.w * 0.35, Math.max(...data.categories.map(c => String(c).length)) * 6.2 + 10) : labelW) + (chart.yTitle ? 18 : 0);
  const bottom = 22 + (chart.xTitle ? 18 : 0);
  const area = { x: plot.x + left, y: plot.y + 6, w: Math.max(10, plot.w - left - 10), h: Math.max(10, plot.h - bottom - 6) };
  const ticks = [];
  for (let v = scale.min; v <= scale.max + scale.step / 2; v += scale.step) ticks.push(Number(v.toPrecision(12)));
  const valuePos = v => (horizontal ? area.x + ((v - scale.min) / (scale.max - scale.min)) * area.w : area.y + area.h - ((v - scale.min) / (scale.max - scale.min)) * area.h);
  const axis = svg('g', { 'font-size': 10, fill: muted });
  for (const v of ticks) {
    const p = valuePos(v);
    if (horizontal) {
      root.append(svg('line', { x1: p, y1: area.y, x2: p, y2: area.y + area.h, stroke: line, 'stroke-width': 1 }));
      axis.append(svg('text', { x: p, y: area.y + area.h + 14, 'text-anchor': 'middle' }, fmt(v, data.format)));
    } else {
      root.append(svg('line', { x1: area.x, y1: p, x2: area.x + area.w, y2: p, stroke: line, 'stroke-width': 1 }));
      axis.append(svg('text', { x: area.x - 6, y: p + 3.5, 'text-anchor': 'end' }, fmt(v, data.format)));
    }
  }
  root.append(axis);
  if (chart.xTitle) root.append(svg('text', { x: area.x + area.w / 2, y: plot.y + plot.h - 2, 'text-anchor': 'middle', 'font-size': 11, fill: fg }, chart.xTitle));
  if (chart.yTitle) {
    const ty = svg('text', { x: 0, y: 0, 'text-anchor': 'middle', 'font-size': 11, fill: fg, transform: `translate(${plot.x + 10} ${area.y + area.h / 2}) rotate(-90)` }, chart.yTitle);
    root.append(ty);
  }
  const g = svg('g');
  root.append(g);
  if (chart.type === 'scatter') {
    const xPos = x => area.x + ((x - xScale.min) / (xScale.max - xScale.min)) * area.w;
    for (let v = xScale.min; v <= xScale.max + xScale.step / 2; v += xScale.step) {
      const p = xPos(v);
      root.insertBefore(svg('line', { x1: p, y1: area.y, x2: p, y2: area.y + area.h, stroke: line, 'stroke-width': 1 }), axis);
      axis.append(svg('text', { x: p, y: area.y + area.h + 14, 'text-anchor': 'middle' }, fmt(Number(v.toPrecision(12)), data.xFormat)));
    }
    data.series.forEach((s, si) => {
      const color = colors[si % colors.length];
      s.values.forEach((y, i) => {
        const x = s.xs ? s.xs[i] : i + 1;
        if (y === null || x === null) return;
        g.append(svg('circle', { cx: xPos(x), cy: valuePos(y), r: 4, fill: color, stroke: dark ? '#262626' : '#ffffff', 'stroke-width': 1 }, undefined));
        if (chart.labels) g.append(svg('text', { x: xPos(x) + 6, y: valuePos(y) - 6, 'font-size': 10, fill: fg }, fmt(y, data.format)));
      });
    });
    return root;
  }
  // the category axis
  const band = (horizontal ? area.h : area.w) / Math.max(1, n);
  const catPos = i => (horizontal ? area.y + band * i + band / 2 : area.x + band * i + band / 2);
  const every = Math.max(1, Math.ceil(n / Math.max(1, (horizontal ? area.h / 14 : area.w / 50))));
  data.categories.forEach((c, i) => {
    if (i % every) return;
    if (horizontal) axis.append(svg('text', { x: area.x - 6, y: catPos(i) + 3.5, 'text-anchor': 'end' }, clip(c, Math.floor(left / 6.2))));
    else axis.append(svg('text', { x: catPos(i), y: area.y + area.h + 14, 'text-anchor': 'middle' }, clip(c, Math.max(3, Math.floor(band * every / 6.2)))));
  });
  const zero = valuePos(Math.max(scale.min, Math.min(0, scale.max)));
  root.append(svg('line', horizontal ? { x1: zero, y1: area.y, x2: zero, y2: area.y + area.h, stroke: muted, 'stroke-width': 1 } : { x1: area.x, y1: zero, x2: area.x + area.w, y2: zero, stroke: muted, 'stroke-width': 1 }));
  const label = (x, y, v, anchor = 'middle') => { if (chart.labels && v !== null) g.append(svg('text', { x, y, 'text-anchor': anchor, 'font-size': 10, fill: fg }, fmt(v, data.format))); };
  if (chart.type === 'column' || chart.type === 'bar') {
    const k = data.series.length;
    const inner = band * 0.72;
    const barW = stacked ? inner : inner / k;
    const pos = new Array(n).fill(0);
    const neg = new Array(n).fill(0);
    data.series.forEach((s, si) => {
      const color = colors[si % colors.length];
      s.values.forEach((v, i) => {
        if (v === null) return;
        let from;
        let to;
        if (stacked) { const base = v >= 0 ? pos[i] : neg[i]; from = base; to = base + v; if (v >= 0) pos[i] = to; else neg[i] = to; } else { from = 0; to = v; }
        const offset = stacked ? -inner / 2 : -inner / 2 + barW * si;
        const p1 = valuePos(from);
        const p2 = valuePos(to);
        if (horizontal) {
          const y = catPos(i) + offset;
          g.append(svg('rect', { x: Math.min(p1, p2), y, width: Math.max(1, Math.abs(p2 - p1)), height: Math.max(1, barW - 1), fill: color }));
          label(Math.max(p1, p2) + 4, y + barW / 2 + 3.5, v, 'start');
        } else {
          const x = catPos(i) + offset;
          g.append(svg('rect', { x, y: Math.min(p1, p2), width: Math.max(1, barW - 1), height: Math.max(1, Math.abs(p2 - p1)), fill: color }));
          label(x + barW / 2, Math.min(p1, p2) - 4, v);
        }
      });
    });
  } else {
    const base = new Array(n).fill(0);
    data.series.forEach((s, si) => {
      const color = colors[si % colors.length];
      const pts = s.values.map((v, i) => {
        if (v === null) return null;
        const y = stacked ? base[i] + v : v;
        return { x: catPos(i), y: valuePos(y), v, b: valuePos(stacked ? base[i] : Math.max(scale.min, Math.min(0, scale.max))) };
      });
      if (chart.type === 'area') {
        const valid = pts.filter(Boolean);
        if (valid.length) {
          const d = 'M' + valid.map(p => `${p.x},${p.y}`).join(' L') + ' L' + [...valid].reverse().map(p => `${p.x},${p.b}`).join(' L') + ' Z';
          g.append(svg('path', { d, fill: color, 'fill-opacity': stacked ? 0.9 : 0.75, stroke: color, 'stroke-width': 1 }));
        }
      } else {
        let d = '';
        pts.forEach((p, i) => { if (!p) return; d += (d && pts[i - 1] ? ' L' : ' M') + `${p.x},${p.y}`; });
        g.append(svg('path', { d: d.trim(), fill: 'none', stroke: color, 'stroke-width': 2.25, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
        for (const p of pts) if (p && n <= 60) g.append(svg('circle', { cx: p.x, cy: p.y, r: 3, fill: color }));
      }
      for (const p of pts) if (p) label(p.x, p.y - 6, p.v);
      if (stacked) s.values.forEach((v, i) => { base[i] += v || 0; });
    });
  }
  return root;
}

function drawPie(root, chart, data, plot, colors, fg, dark, fmt) {
  const s = data.series[0];
  const values = s.values.map(v => (v && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const r = Math.max(10, Math.min(plot.w, plot.h) / 2 - 8);
  const inner = chart.type === 'doughnut' ? r * 0.55 : 0;
  const g = svg('g');
  root.append(g);
  if (!total) return;
  let angle = -Math.PI / 2;
  values.forEach((v, i) => {
    if (!v) return;
    const a = (v / total) * Math.PI * 2;
    const end = angle + a;
    const large = a > Math.PI ? 1 : 0;
    const p = (rad, ang) => `${cx + rad * Math.cos(ang)},${cy + rad * Math.sin(ang)}`;
    let d;
    if (a >= Math.PI * 2 - 1e-9) {
      d = `M${p(r, 0)} A${r},${r} 0 1 1 ${p(r, Math.PI)} A${r},${r} 0 1 1 ${p(r, 0)}` + (inner ? ` M${p(inner, 0)} A${inner},${inner} 0 1 0 ${p(inner, Math.PI)} A${inner},${inner} 0 1 0 ${p(inner, 0)}` : '');
    } else if (inner) {
      d = `M${p(r, angle)} A${r},${r} 0 ${large} 1 ${p(r, end)} L${p(inner, end)} A${inner},${inner} 0 ${large} 0 ${p(inner, angle)} Z`;
    } else {
      d = `M${cx},${cy} L${p(r, angle)} A${r},${r} 0 ${large} 1 ${p(r, end)} Z`;
    }
    g.append(svg('path', { d, fill: colors[i % colors.length], stroke: dark ? '#262626' : '#ffffff', 'stroke-width': 1.5, 'fill-rule': 'evenodd' }));
    if (chart.labels || values.filter(Boolean).length <= 8) {
      const mid = angle + a / 2;
      const lr = inner ? (r + inner) / 2 : r * 0.65;
      if (a > 0.25) g.append(svg('text', { x: cx + lr * Math.cos(mid), y: cy + lr * Math.sin(mid) + 4, 'text-anchor': 'middle', 'font-size': 11, fill: '#ffffff', 'font-weight': 600 }, chart.labels ? fmt(v, data.format) : Math.round((v / total) * 100) + '%'));
    }
    angle = end;
  });
  void fg;
}

function clip(text, n) {
  const s = String(text ?? '');
  return s.length > n ? s.slice(0, Math.max(1, n - 1)) + '…' : s;
}
