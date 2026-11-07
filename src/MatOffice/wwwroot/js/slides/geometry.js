/* MatOffice Slides – the geometry of the shapes: an SVG path for each preset in a w × h box (with its one adjustment), the rectangle the text
   of a shape sits in, arrow heads of lines. The names are the preset names of PowerPoint (prstGeom), so a file keeps its shapes. */

const f = n => Math.round(n * 100) / 100;

/** The default adjustment of a shape (a fraction). */
export const DEFAULT_ADJ = {
  roundRect: 0.1667, snip1Rect: 0.1667, triangle: 0.5, parallelogram: 0.25, trapezoid: 0.25, hexagon: 0.25, octagon: 0.29, plus: 0.25, donut: 0.25, can: 0.25,
  rightArrow: 0.5, leftArrow: 0.5, upArrow: 0.5, downArrow: 0.5, leftRightArrow: 0.5, upDownArrow: 0.5, chevron: 0.5, homePlate: 0.5, star5: 0.38, star6: 0.5,
  wedgeRoundRectCallout: 0.1667,
};

export const adjOf = (shape, adj) => (adj === undefined || adj === null ? DEFAULT_ADJ[shape] ?? 0 : adj);

/** The path of a preset shape in a w × h box. */
export function shapePath(shape, w, h, adjIn) {
  const a = adjOf(shape, adjIn);
  const ss = Math.min(w, h);
  const P = pts => 'M' + pts.map(p => f(p[0]) + ' ' + f(p[1])).join(' L') + ' Z';
  switch (shape) {
    case 'roundRect': {
      const r = Math.min(Math.max(0, a), 0.5) * ss;
      return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`;
    }
    case 'snip1Rect': { const d = Math.min(a, 0.5) * ss; return P([[0, 0], [w - d, 0], [w, d], [w, h], [0, h]]); }
    case 'ellipse': return `M0 ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w)} ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 0 ${f(h / 2)} Z`;
    case 'triangle': return P([[w * a, 0], [w, h], [0, h]]);
    case 'rtTriangle': return P([[0, 0], [w, h], [0, h]]);
    case 'diamond': return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': { const d = Math.min(a, 1) * ss; return P([[d, 0], [w, 0], [w - d, h], [0, h]]); }
    case 'trapezoid': { const d = Math.min(a, 0.5) * ss; return P([[d, 0], [w - d, 0], [w, h], [0, h]]); }
    case 'pentagon': return P([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'hexagon': { const d = Math.min(a, 0.5) * w; return P([[d, 0], [w - d, 0], [w, h / 2], [w - d, h], [d, h], [0, h / 2]]); }
    case 'octagon': { const d = Math.min(a, 0.5) * ss; return P([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]); }
    case 'plus': {
      const d = Math.min(a, 0.5) * ss;
      return P([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]);
    }
    case 'donut': {
      const d = Math.min(a, 0.5) * ss;
      const outer = `M0 ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w)} ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 0 ${f(h / 2)} Z`;
      const rx = w / 2 - d, ry = h / 2 - d;
      return rx > 0 && ry > 0 ? outer + ` M${f(d)} ${f(h / 2)} A${f(rx)} ${f(ry)} 0 1 0 ${f(w - d)} ${f(h / 2)} A${f(rx)} ${f(ry)} 0 1 0 ${f(d)} ${f(h / 2)} Z` : outer;
    }
    case 'can': {
      const ry = Math.min(a, 0.5) * ss / 2;
      return `M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(w)} ${f(ry)} V${f(h - ry)} A${f(w / 2)} ${f(ry)} 0 0 1 0 ${f(h - ry)} Z M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 0 ${f(w)} ${f(ry)}`;
    }
    case 'heart': {
      return `M${f(w / 2)} ${f(h * 0.25)} C${f(w / 2)} ${f(-h * 0.1)} ${f(-w * 0.15)} ${f(h * 0.02)} ${f(w * 0.02)} ${f(h * 0.4)} C${f(w * 0.12)} ${f(h * 0.62)} ${f(w * 0.4)} ${f(h * 0.8)} ${f(w / 2)} ${f(h)} `
        + `C${f(w * 0.6)} ${f(h * 0.8)} ${f(w * 0.88)} ${f(h * 0.62)} ${f(w * 0.98)} ${f(h * 0.4)} C${f(w * 1.15)} ${f(h * 0.02)} ${f(w / 2)} ${f(-h * 0.1)} ${f(w / 2)} ${f(h * 0.25)} Z`;
    }
    case 'cloud': {
      const c = (x, y) => `${f(x * w)} ${f(y * h)}`;
      return `M${c(0.14, 0.58)} C${c(0.02, 0.56)} ${c(0.02, 0.33)} ${c(0.17, 0.32)} C${c(0.15, 0.12)} ${c(0.38, 0.05)} ${c(0.45, 0.17)} C${c(0.52, 0.02)} ${c(0.72, 0.03)} ${c(0.75, 0.2)} `
        + `C${c(0.92, 0.13)} ${c(1.02, 0.35)} ${c(0.92, 0.47)} C${c(1.04, 0.6)} ${c(0.95, 0.85)} ${c(0.78, 0.8)} C${c(0.72, 0.98)} ${c(0.5, 0.99)} ${c(0.44, 0.86)} `
        + `C${c(0.33, 0.98)} ${c(0.12, 0.92)} ${c(0.14, 0.58)} Z`;
    }
    case 'star5': case 'star6': {
      const n = shape === 'star5' ? 5 : 6;
      const inner = shape === 'star5' ? Math.min(Math.max(a, 0.1), 0.9) : Math.min(Math.max(a, 0.1), 0.9) * 0.58 / 0.5;
      const pts = [];
      for (let i = 0; i < n * 2; i++) {
        const r = i % 2 === 0 ? 1 : inner * (shape === 'star5' ? 1 : 0.5 / 0.58);
        const ang = -Math.PI / 2 + (i * Math.PI) / n;
        pts.push([Math.cos(ang) * r, Math.sin(ang) * r]);
      }
      const minX = Math.min(...pts.map(p => p[0])), maxX = Math.max(...pts.map(p => p[0]));
      const minY = Math.min(...pts.map(p => p[1])), maxY = Math.max(...pts.map(p => p[1]));
      return P(pts.map(([x, y]) => [((x - minX) / (maxX - minX)) * w, ((y - minY) / (maxY - minY)) * h]));
    }
    case 'rightArrow': { const hl = Math.min(a * ss, w); return P([[0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [0, h * 0.75]]); }
    case 'leftArrow': { const hl = Math.min(a * ss, w); return P([[w, h * 0.25], [hl, h * 0.25], [hl, 0], [0, h / 2], [hl, h], [hl, h * 0.75], [w, h * 0.75]]); }
    case 'upArrow': { const hl = Math.min(a * ss, h); return P([[w * 0.25, h], [w * 0.25, hl], [0, hl], [w / 2, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h]]); }
    case 'downArrow': { const hl = Math.min(a * ss, h); return P([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [w * 0.25, h - hl]]); }
    case 'leftRightArrow': { const hl = Math.min(a * ss, w / 2); return P([[0, h / 2], [hl, 0], [hl, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [hl, h * 0.75], [hl, h]]); }
    case 'upDownArrow': { const hl = Math.min(a * ss, h / 2); return P([[w / 2, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [w * 0.25, h - hl], [w * 0.25, hl], [0, hl]]); }
    case 'chevron': { const d = Math.min(Math.min(a, 1) * ss, w); return P([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case 'homePlate': { const d = Math.min(Math.min(a, 1) * ss, w); return P([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    case 'wedgeRectCallout': return P([[0, 0], [w, 0], [w, h], [w * 0.36, h], [w * 0.2, h * 1.25], [w * 0.22, h], [0, h]]);
    case 'wedgeRoundRectCallout': {
      const r = Math.min(Math.max(0, a), 0.5) * ss;
      return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(w * 0.36)} L${f(w * 0.2)} ${f(h * 1.25)} L${f(w * 0.22)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`;
    }
    case 'wedgeEllipseCallout': {
      // an ellipse with a wedge to the lower left
      const cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2;
      const p = t => [cx + rx * Math.cos(t), cy + ry * Math.sin(t)];
      const s = p(Math.PI * 0.62), e = p(Math.PI * 0.72);
      return `M${f(s[0])} ${f(s[1])} A${f(rx)} ${f(ry)} 0 1 1 ${f(e[0])} ${f(e[1])} L${f(w * 0.12)} ${f(h * 1.2)} Z`;
    }
    case 'rect':
    default:
      return `M0 0 H${f(w)} V${f(h)} H0 Z`;
  }
}

/** The rectangle the text of a shape sits in (fractions of the box), like the text rectangles of the presets. */
export function textRect(shape, w, h) {
  const r = (x, y, rw, rh) => ({ x: x * w, y: y * h, w: rw * w, h: rh * h });
  switch (shape) {
    case 'ellipse': case 'donut': case 'wedgeEllipseCallout': return r(0.146, 0.146, 0.708, 0.708);
    case 'triangle': return r(0.25, 0.5, 0.5, 0.5);
    case 'rtTriangle': return r(0.083, 0.5, 0.5, 0.417);
    case 'diamond': return r(0.25, 0.25, 0.5, 0.5);
    case 'pentagon': return r(0.191, 0.25, 0.618, 0.65);
    case 'hexagon': case 'octagon': return r(0.15, 0.15, 0.7, 0.7);
    case 'star5': case 'star6': return r(0.3, 0.38, 0.4, 0.4);
    case 'heart': return r(0.2, 0.2, 0.6, 0.5);
    case 'cloud': return r(0.15, 0.2, 0.7, 0.6);
    case 'can': return r(0, 0.2, 1, 0.8);
    case 'rightArrow': case 'leftArrow': case 'leftRightArrow': return r(0.1, 0.25, 0.8, 0.5);
    case 'upArrow': case 'downArrow': case 'upDownArrow': return r(0.25, 0.1, 0.5, 0.8);
    case 'chevron': return r(0.2, 0, 0.6, 1);
    case 'homePlate': return r(0, 0, 0.85, 1);
    case 'parallelogram': case 'trapezoid': return r(0.12, 0, 0.76, 1);
    default: return r(0, 0, 1, 1);
  }
}

/** The points of an arrow head at the tip (x, y) pointing along the angle (radians), for a line of width lw (px). */
export function arrowHead(kind, x, y, angle, lw) {
  const len = Math.max(6, lw * 3.2), half = Math.max(4, lw * 1.8);
  const back = (d, side) => [x - Math.cos(angle) * d + Math.cos(angle + Math.PI / 2) * side, y - Math.sin(angle) * d + Math.sin(angle + Math.PI / 2) * side];
  switch (kind) {
    case 'triangle': return { type: 'polygon', points: [[x, y], back(len, half), back(len, -half)] };
    case 'stealth': return { type: 'polygon', points: [[x, y], back(len, half), back(len * 0.6, 0), back(len, -half)] };
    case 'arrow': return { type: 'polyline', points: [back(len, half), [x, y], back(len, -half)] };
    case 'diamond': return { type: 'polygon', points: [[x, y], back(len / 2, half), back(len, 0), back(len / 2, -half)] };
    case 'oval': return { type: 'circle', cx: x - Math.cos(angle) * half, cy: y - Math.sin(angle) * half, r: half };
    default: return null;
  }
}

/** How far a line end has to be pulled back so that it ends inside its arrow head. */
export const arrowInset = (kind, lw) => (kind === 'triangle' || kind === 'stealth' ? Math.max(6, lw * 3.2) * 0.7 : kind === 'oval' || kind === 'diamond' ? Math.max(4, lw * 1.8) : 0);

/** The dash pattern of a line (px) for its width. */
export function dashArray(dash, lw) {
  const w = Math.max(1, lw);
  switch (dash) {
    case 'dash': return `${w * 4} ${w * 3}`;
    case 'dot': return `${w} ${w}`;
    case 'dashDot': return `${w * 4} ${w * 3} ${w} ${w * 3}`;
    case 'longDash': return `${w * 8} ${w * 3}`;
    default: return null;
  }
}
