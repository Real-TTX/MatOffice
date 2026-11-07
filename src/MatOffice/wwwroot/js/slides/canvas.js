/* MatOffice Slides – the slide in the middle of the normal view: drawn by render.js and scaled to the zoom, with an overlay in screen pixels for
   the selection (handles, rotate handle, line ends), smart guides, the marquee, gridlines and guides. Pointer and keyboard work like the office suites:
   click/Shift-click/drag a marquee to select, drag to move (Ctrl copies, Shift keeps the direction, Alt does not snap), handles resize
   (Shift keeps the proportions, Ctrl from the centre), the round handle rotates (Shift in 15° steps), arrows nudge (Shift more, Ctrl less). */
import { t, el, icon, mod } from '../editor/util.js';
import { renderSlide, fitText, fitBox } from './render.js';
import { rotatedBounds, bounds, withGroups, tableLook, round, roleOf } from './model.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const SNAP = 6;

export class Canvas {
  constructor(app) {
    this.app = app;
    this.zoom = 1;
    this.fit = true;
    this.stage = el('div', { class: 'mo-slides-stage', tabindex: '0', role: 'region', 'aria-label': t('Slides.Canvas') });
    this.sheet = el('div', { class: 'mo-slides-sheet' });
    this.host = el('div', { class: 'mo-slides-slide' });
    this.overlay = el('div', { class: 'mo-slides-overlay', 'aria-hidden': 'true' });
    this.sheet.append(this.host, this.overlay);
    this.stage.append(this.sheet);
    this.guides = [];
    this.stage.addEventListener('pointerdown', e => this.down(e));
    this.stage.addEventListener('dblclick', e => this.dblclick(e));
    this.stage.addEventListener('keydown', e => this.key(e));
    this.stage.addEventListener('contextmenu', e => this.context(e));
    this.stage.addEventListener('dragover', e => { if ([...e.dataTransfer.types].includes('Files')) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; } });
    this.stage.addEventListener('drop', e => this.drop(e));
    this.stage.addEventListener('focus', () => { this.app.area = 'stage'; });
    this.stage.addEventListener('scroll', () => this.drawSelection(), { passive: true });
    new ResizeObserver(() => { if (this.fit) this.fitZoom(); else this.layout(); }).observe(this.stage);
  }

  focus() { this.stage.focus({ preventScroll: true }); }

  get slide() { return this.app.slide; }

  // ---- drawing ----------------------------------------------------------------------------------------------------------------

  /** Draws the current slide again (the element whose text is edited keeps its DOM). */
  render() {
    const { content } = this.app;
    const slide = this.slide;
    if (!slide) { this.host.replaceChildren(el('div', { class: 'mo-slides-noslide', text: t('Slides.NoSlides') })); this.overlay.replaceChildren(); return; }
    const editing = this.app.text.active;
    const keep = editing ? this.host.querySelector(`.mo-sr-el[data-id="${CSS.escape(editing.id)}"]`) : null;
    const node = renderSlide(content, slide, { mode: this.app.readOnly ? 'view' : 'edit', number: this.app.index + 1, editing: editing?.id });
    if (keep) {
      const fresh = node.querySelector(`.mo-sr-el[data-id="${CSS.escape(editing.id)}"]`);
      if (fresh) { this.syncGeometry(keep, slide.elements.find(e => e.id === editing.id)); fresh.replaceWith(keep); }
    }
    this.host.replaceChildren(node);
    this.slideNode = node;
    this.layout();
    fitText(node);
    this.drawSelection();
  }

  /** Draws the slide with the element whose text is edited as editable; answers the body element to edit. */
  renderEditing(active) {
    const { content } = this.app;
    const node = renderSlide(content, this.slide, { mode: 'edit', number: this.app.index + 1, editing: active.id });
    this.host.replaceChildren(node);
    this.slideNode = node;
    this.layout();
    fitText(node);
    this.drawSelection();
    const elNode = node.querySelector(`.mo-sr-el[data-id="${CSS.escape(active.id)}"]`);
    if (!elNode) return null;
    if (active.cell) return elNode.querySelector(`td[data-cell="${active.cell[0]},${active.cell[1]}"] .mo-sr-body`);
    return elNode.querySelector('.mo-sr-text .mo-sr-body');
  }

  cellLook(e, [r, c]) { const look = tableLook(e, r, c); return { text: look.text, bold: look.bold }; }

  syncGeometry(node, e) {
    if (!node || !e) return;
    Object.assign(node.style, { left: e.x + 'px', top: e.y + 'px', width: Math.max(0, e.w) + 'px', height: Math.max(0, e.h) + 'px', transform: e.rot ? `rotate(${e.rot}deg)` : '' });
    const geom = node.querySelector(':scope > .mo-sr-geom');
    if (geom && e.type !== 'line') { geom.setAttribute('width', e.w); geom.setAttribute('height', e.h); }
  }

  nodeOf(id) { return this.host.querySelector(`.mo-sr-el[data-id="${CSS.escape(id)}"]`); }

  /** After the text of the edited element changed: grow the box, shrink the text, let table rows grow. */
  afterText() {
    const active = this.app.text.active;
    if (!active) return;
    const e = this.slide.elements.find(x => x.id === active.id);
    const node = this.nodeOf(active.id);
    if (!e || !node) return;
    if (e.type === 'table') {
      const rows = node.querySelectorAll('tr');
      rows.forEach((tr, i) => { if (e.rows[i]) e.rows[i].h = round(Math.max(e.rows[i].h, tr.offsetHeight)); });
      e.h = round(e.rows.reduce((a, r) => a + r.h, 0));
      this.syncGeometry(node, e);
      this.drawSelection();
      return;
    }
    const box = node.querySelector('.mo-sr-text');
    if (!box) return;
    const fit = box.dataset.fit;
    if (fit === 'shrink') fitBox(box);
    if (fit === 'grow' && !e.rot) {
      const body = box.querySelector('.mo-sr-body');
      const st = box.style;
      const padY = parseFloat(st.paddingTop) + parseFloat(st.paddingBottom);
      const padX = parseFloat(st.paddingLeft) + parseFloat(st.paddingRight);
      const h = round(body.offsetHeight + padY);
      if (Math.abs(h - e.h) > 0.5) { e.h = h; }
      if (e.text?.wrap === false) {
        const w = round(Math.max(body.scrollWidth, 8) + padX);
        if (Math.abs(w - e.w) > 0.5) e.w = w;
      }
      this.syncGeometry(node, e);
      box.style.width = e.w + 'px';
      box.style.height = e.h + 'px';
    }
    this.drawSelection();
  }

  /** The size of the stage and where the slide sits in it. */
  layout() {
    const { w, h } = this.app.content.size;
    const sw = this.stage.clientWidth, sh = this.stage.clientHeight;
    const margin = this.app.phone ? 10 : 28;
    const width = w * this.zoom, height = h * this.zoom;
    const outerW = Math.max(sw, width + margin * 2), outerH = Math.max(sh, height + margin * 2);
    this.sheet.style.width = outerW + 'px';
    this.sheet.style.height = outerH + 'px';
    this.left = Math.round((outerW - width) / 2);
    this.top = Math.round((outerH - height) / 2);
    Object.assign(this.host.style, { left: this.left + 'px', top: this.top + 'px', width: w + 'px', height: h + 'px', transform: `scale(${this.zoom})` });
    Object.assign(this.overlay.style, { left: this.left + 'px', top: this.top + 'px', width: width + 'px', height: height + 'px' });
    this.overlay.classList.toggle('has-grid', !!this.app.view.grid);
    this.overlay.style.setProperty('--grid', (48 * this.zoom) + 'px');
    this.overlay.classList.toggle('has-guides', !!this.app.view.guides);
  }

  fitZoom() {
    const { w, h } = this.app.content.size;
    const margin = this.app.phone ? 20 : 56;
    const sw = this.stage.clientWidth - margin, sh = this.stage.clientHeight - margin;
    if (sw <= 0 || sh <= 0) return;
    this.zoom = Math.max(0.05, Math.min(4, Math.min(sw / w, sh / h)));
    this.fit = true;
    this.layout();
    this.drawSelection();
    this.app.zoomChanged();
  }

  setZoom(z) {
    this.fit = false;
    this.zoom = Math.max(0.1, Math.min(4, z));
    this.layout();
    this.drawSelection();
    this.app.zoomChanged();
  }

  /** A point of the screen in slide pixels. */
  toSlide(clientX, clientY) {
    const r = this.overlay.getBoundingClientRect();
    return { x: (clientX - r.left) / this.zoom, y: (clientY - r.top) / this.zoom };
  }

  // ---- the selection --------------------------------------------------------------------------------------------------------------

  selected() { return this.slide ? this.slide.elements.filter(e => this.app.selection.includes(e.id)) : []; }

  /** The selection as units: a group is one unit. */
  units() {
    const units = new Map();
    for (const e of this.selected()) {
      const key = e.group || e.id;
      if (!units.has(key)) units.set(key, []);
      units.get(key).push(e);
    }
    return [...units.values()];
  }

  drawSelection() {
    const o = this.overlay;
    o.replaceChildren();
    if (!this.slide) return;
    const z = this.zoom;
    for (const g of this.guides) {
      o.append(el('div', { class: 'mo-slides-guide ' + (g.axis === 'x' ? 'is-v' : 'is-h'), style: g.axis === 'x' ? { left: g.at * z + 'px' } : { top: g.at * z + 'px' } }));
    }
    if (this.marquee) {
      const m = this.marquee;
      o.append(el('div', { class: 'mo-slides-marquee', style: { left: Math.min(m.x0, m.x1) * z + 'px', top: Math.min(m.y0, m.y1) * z + 'px', width: Math.abs(m.x1 - m.x0) * z + 'px', height: Math.abs(m.y1 - m.y0) * z + 'px' } }));
    }
    const sel = this.selected();
    if (!sel.length) return;
    const units = this.units();
    const editing = this.app.text.active;
    const ro = this.app.readOnly;
    if (units.length === 1 && units[0].length === 1) {
      const e = units[0][0];
      if (e.type === 'line') { this.lineHandles(e, ro); return; }
      const box = this.frame(e.x, e.y, e.w, e.h, e.rot || 0, editing ? 'is-editing' : '');
      if (!ro) this.handles(box, e.type !== 'table');
      o.append(box);
      return;
    }
    for (const e of sel) {
      if (e.type === 'line') { this.lineHandles(e, true); continue; }
      o.append(this.frame(e.x, e.y, e.w, e.h, e.rot || 0, 'is-member'));
    }
    const b = bounds(sel.map(rotatedBounds));
    const box = this.frame(b.x, b.y, b.w, b.h, 0, units.length === 1 ? 'is-group' : 'is-multi');
    if (!ro) this.handles(box, true);
    o.append(box);
  }

  frame(x, y, w, h, rot, cls) {
    const z = this.zoom;
    return el('div', { class: 'mo-slides-frame ' + (cls || ''), style: { left: x * z + 'px', top: y * z + 'px', width: Math.max(1, w * z) + 'px', height: Math.max(1, h * z) + 'px', transform: rot ? `rotate(${rot}deg)` : '' } });
  }

  handles(box, rotate) {
    for (const h of HANDLES) box.append(el('div', { class: 'mo-slides-handle is-' + h, 'data-handle': h }));
    if (rotate) box.append(el('div', { class: 'mo-slides-rotate', 'data-handle': 'rotate', 'data-tip': t('Slides.Rotate') }, icon('refresh')));
  }

  lineHandles(e, ro) {
    const z = this.zoom;
    const [x1, y1, x2, y2] = lineEnds(e);
    const line = el('div', { class: 'mo-slides-lineframe' });
    const len = Math.hypot(x2 - x1, y2 - y1) * z;
    Object.assign(line.style, { left: x1 * z + 'px', top: y1 * z + 'px', width: len + 'px', transform: `rotate(${Math.atan2(y2 - y1, x2 - x1)}rad)` });
    this.overlay.append(line);
    if (ro) return;
    this.overlay.append(el('div', { class: 'mo-slides-handle is-end', 'data-handle': 'start', style: { left: x1 * z + 'px', top: y1 * z + 'px' } }));
    this.overlay.append(el('div', { class: 'mo-slides-handle is-end', 'data-handle': 'end', style: { left: x2 * z + 'px', top: y2 * z + 'px' } }));
  }

  // ---- pointer ----------------------------------------------------------------------------------------------------------------------

  hit(target) {
    const node = target.closest?.('.mo-sr-el[data-id]');
    if (!node || !this.host.contains(node)) return null;
    return this.slide.elements.find(e => e.id === node.dataset.id) || null;
  }

  down(e) {
    if (e.button === 2) return;
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    this.app.area = 'stage';
    const app = this.app;
    const insert = e.target.closest?.('[data-insert]');
    if (insert && !app.readOnly) {
      e.preventDefault();
      const target = this.hit(insert);
      if (target) app.fillPlaceholder(target.id, insert.dataset.insert);
      return;
    }
    const editing = app.text.active;
    if (editing && app.text.bodyEl?.contains(e.target)) return; // the caret moves in the text
    if (app.drawing && !app.readOnly) { e.preventDefault(); if (editing) app.text.stop(); this.drawShape(e); return; }
    const handle = e.target.closest?.('[data-handle]');
    if (handle && !app.readOnly) { e.preventDefault(); this.dragHandle(e, handle.dataset.handle); return; }
    const hitEl = this.hit(e.target);
    if (editing) app.text.stop();
    if (app.painter && hitEl) { e.preventDefault(); app.paintFormat(hitEl.id); return; }
    if (!hitEl) {
      if (!e.shiftKey && !mod(e)) app.select([]);
      if (!app.readOnly) this.dragMarquee(e);
      this.focus();
      return;
    }
    e.preventDefault();
    this.focus();
    if (app.readOnly) { app.select(withGroups(this.slide, [hitEl.id])); return; }
    const unit = withGroups(this.slide, [hitEl.id]);
    const wasSelected = app.selection.includes(hitEl.id);
    const wasOnly = wasSelected && app.selection.length === unit.length;
    if (e.shiftKey || mod(e)) {
      if (wasSelected) app.select(app.selection.filter(id => !unit.includes(id)));
      else app.select([...app.selection, ...unit]);
      if (!wasSelected) this.dragMove(e, hitEl, false);
      return;
    }
    // a member of a selected group: the second click selects the member alone
    if (wasOnly && unit.length > 1 && hitEl.group) { app.select([hitEl.id]); }
    else if (!wasSelected) app.select(unit);
    this.dragMove(e, hitEl, wasOnly);
  }

  /** Insert mode: a drag draws the box of the new shape (Shift keeps it square), a click puts it there in its default size. */
  drawShape(e) {
    const p0 = this.toSlide(e.clientX, e.clientY);
    this.stage.setPointerCapture?.(e.pointerId);
    const line = ['line', 'arrow', 'doubleArrow'].includes(this.app.drawing);
    let box = { x: p0.x, y: p0.y, w: 0, h: 0 };
    const move = ev => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      let w = p.x - p0.x, h = p.y - p0.y;
      if (ev.shiftKey && !line) { const s = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * s; h = Math.sign(h || 1) * s; }
      if (ev.shiftKey && line) { const a = Math.round(Math.atan2(h, w) / (Math.PI / 4)) * (Math.PI / 4); const len = Math.hypot(w, h); w = Math.cos(a) * len; h = Math.sin(a) * len; }
      box = { x: Math.min(p0.x, p0.x + w), y: Math.min(p0.y, p0.y + h), w: Math.abs(w), h: Math.abs(h), flipH: line && w < 0, flipV: line && h < 0 };
      this.marquee = { x0: p0.x, y0: p0.y, x1: p0.x + w, y1: p0.y + h };
      this.drawSelection();
    };
    const up = () => {
      this.stage.removeEventListener('pointermove', move);
      this.stage.removeEventListener('pointerup', up);
      this.stage.removeEventListener('pointercancel', up);
      this.marquee = null;
      const click = Math.max(box.w, box.h) * this.zoom < 6;
      this.app.drawn(click ? { x: p0.x, y: p0.y, w: 0, h: 0 } : box, click);
    };
    this.stage.addEventListener('pointermove', move);
    this.stage.addEventListener('pointerup', up);
    this.stage.addEventListener('pointercancel', up);
  }

  dblclick(e) {
    if (this.app.readOnly) return;
    const hitEl = this.hit(e.target);
    if (!hitEl || this.app.text.active) return;
    if (hitEl.type === 'line' || hitEl.type === 'image') { if (hitEl.type === 'image') this.app.showPane('format'); return; }
    this.startText(hitEl, e);
  }

  /** Starts editing the text of an element at a point (a table: the cell there). */
  startText(target, e) {
    const app = this.app;
    if (target.type === 'image' || target.type === 'line') return false;
    app.select([target.id]);
    let cell = null;
    if (target.type === 'table') {
      const td = e ? document.elementFromPoint(e.clientX, e.clientY)?.closest('td[data-cell]') : null;
      cell = td ? td.dataset.cell.split(',').map(Number) : [0, 0];
    }
    return app.text.start({ id: target.id, cell }, e ? { point: [e.clientX, e.clientY] } : { end: true });
  }

  dragMove(e, hitEl, wasOnly) {
    const app = this.app;
    const start = { x: e.clientX, y: e.clientY };
    const ids = [...app.selection];
    const orig = new Map(this.slide.elements.filter(x => ids.includes(x.id)).map(x => [x.id, { x: x.x, y: x.y }]));
    let moving = false;
    const pointerId = e.pointerId;
    this.stage.setPointerCapture?.(pointerId);
    const move = ev => {
      const dx = (ev.clientX - start.x) / this.zoom, dy = (ev.clientY - start.y) / this.zoom;
      if (!moving && Math.hypot(ev.clientX - start.x, ev.clientY - start.y) < 4) return;
      if (!moving) { moving = true; app.checkpoint(); this.stage.classList.add('is-moving'); }
      let mx = dx, my = dy;
      if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) my = 0; else mx = 0; }
      const els = this.slide.elements.filter(x => orig.has(x.id));
      const box = bounds(els.map(x => rotatedBounds({ ...x, x: orig.get(x.id).x, y: orig.get(x.id).y })));
      const snapped = ev.altKey || !app.view.snap ? { dx: mx, dy: my, guides: [] } : this.snapMove(box, mx, my, orig);
      for (const x of els) {
        const o = orig.get(x.id);
        x.x = round(o.x + snapped.dx);
        x.y = round(o.y + snapped.dy);
        this.syncGeometry(this.nodeOf(x.id), x);
      }
      this.guides = snapped.guides;
      this.drawSelection();
    };
    const up = ev => {
      this.stage.removeEventListener('pointermove', move);
      this.stage.removeEventListener('pointerup', up);
      this.stage.removeEventListener('pointercancel', up);
      this.stage.classList.remove('is-moving');
      this.guides = [];
      if (moving) {
        if (mod(ev) && ev.type === 'pointerup') {
          // Ctrl: a copy where it was dropped, the originals stay
          const moved = this.slide.elements.filter(x => orig.has(x.id)).map(x => ({ id: x.id, x: x.x, y: x.y }));
          for (const m of moved) { const x = this.slide.elements.find(q => q.id === m.id); x.x = orig.get(m.id).x; x.y = orig.get(m.id).y; }
          app.duplicateAt(moved);
        } else app.changed();
        return;
      }
      this.drawSelection();
      if (ev.type !== 'pointerup' || app.readOnly) return;
      // a click without moving: into the text
      const textual = hitEl.type === 'text' || (hitEl.type === 'shape' && hitEl.text && roleOf(this.slide, hitEl) !== 'shape');
      if ((wasOnly && hitEl.type !== 'image' && hitEl.type !== 'line') || (textual && app.selection.length === 1 && !hitEl.group)) {
        this.startText(hitEl, ev);
      }
    };
    this.stage.addEventListener('pointermove', move);
    this.stage.addEventListener('pointerup', up);
    this.stage.addEventListener('pointercancel', up);
  }

  /** The lines a moving box snaps to: the slide's edges and centre, the edges and centres of the other elements. */
  targets(exclude) {
    const { w, h } = this.app.content.size;
    const xs = [0, w / 2, w], ys = [0, h / 2, h];
    for (const e of this.slide.elements) {
      if (exclude.has(e.id)) continue;
      const b = rotatedBounds(e);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w);
      ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    return { xs, ys };
  }

  snapMove(box, dx, dy, exclude) {
    const { xs, ys } = this.targets(exclude);
    const tol = SNAP / this.zoom;
    const best = (values, candidates) => {
      let found = null;
      for (const v of values) for (const c of candidates) { const d = c - v; if (Math.abs(d) <= tol && (!found || Math.abs(d) < Math.abs(found.d))) found = { d, at: c }; }
      return found;
    };
    const sx = best([box.x + dx, box.x + box.w / 2 + dx, box.x + box.w + dx], xs);
    const sy = best([box.y + dy, box.y + box.h / 2 + dy, box.y + box.h + dy], ys);
    const guides = [];
    if (sx) guides.push({ axis: 'x', at: sx.at });
    if (sy) guides.push({ axis: 'y', at: sy.at });
    let ndx = dx + (sx?.d || 0), ndy = dy + (sy?.d || 0);
    if (this.app.view.gridSnap && !sx) ndx = Math.round((box.x + dx) / 12) * 12 - box.x;
    if (this.app.view.gridSnap && !sy) ndy = Math.round((box.y + dy) / 12) * 12 - box.y;
    return { dx: ndx, dy: ndy, guides };
  }

  dragMarquee(e) {
    const p0 = this.toSlide(e.clientX, e.clientY);
    const base = e.shiftKey || mod(e) ? [...this.app.selection] : [];
    this.marquee = { x0: p0.x, y0: p0.y, x1: p0.x, y1: p0.y };
    this.stage.setPointerCapture?.(e.pointerId);
    let moved = false;
    const move = ev => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      this.marquee.x1 = p.x;
      this.marquee.y1 = p.y;
      moved = moved || Math.hypot(p.x - p0.x, p.y - p0.y) * this.zoom > 3;
      this.drawSelection();
    };
    const up = () => {
      this.stage.removeEventListener('pointermove', move);
      this.stage.removeEventListener('pointerup', up);
      this.stage.removeEventListener('pointercancel', up);
      const m = this.marquee;
      this.marquee = null;
      if (moved) {
        const x = Math.min(m.x0, m.x1), y = Math.min(m.y0, m.y1), r = Math.max(m.x0, m.x1), b = Math.max(m.y0, m.y1);
        const inside = this.slide.elements.filter(el2 => { const bb = rotatedBounds(el2); return bb.x >= x && bb.y >= y && bb.x + bb.w <= r && bb.y + bb.h <= b; }).map(x2 => x2.id);
        this.app.select(withGroups(this.slide, [...new Set([...base, ...inside])]));
      } else this.drawSelection();
    };
    this.stage.addEventListener('pointermove', move);
    this.stage.addEventListener('pointerup', up);
    this.stage.addEventListener('pointercancel', up);
  }

  dragHandle(e, handle) {
    const app = this.app;
    const sel = this.selected();
    if (!sel.length) return;
    const units = this.units();
    const single = units.length === 1 && units[0].length === 1 ? units[0][0] : null;
    const orig = new Map(sel.map(x => [x.id, structuredClone(x)]));
    const start = this.toSlide(e.clientX, e.clientY);
    let changed = false;
    this.stage.setPointerCapture?.(e.pointerId);
    const begin = () => { if (!changed) { changed = true; app.checkpoint(); } };
    const move = ev => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      if (!changed && Math.hypot(p.x - start.x, p.y - start.y) * this.zoom < 2) return;
      begin();
      if (handle === 'start' || handle === 'end') this.moveLineEnd(single, orig.get(single.id), handle, p, ev);
      else if (handle === 'rotate') this.rotate(sel, orig, p, ev);
      else if (single && single.type !== 'table') this.resizeOne(single, orig.get(single.id), handle, p, ev);
      else this.resizeMany(sel, orig, handle, p, ev);
      for (const x of sel) this.syncGeometry(this.nodeOf(x.id), x);
      this.renderLive(sel);
      this.drawSelection();
    };
    const up = () => {
      this.stage.removeEventListener('pointermove', move);
      this.stage.removeEventListener('pointerup', up);
      this.stage.removeEventListener('pointercancel', up);
      this.guides = [];
      if (changed) app.changed(); else this.drawSelection();
    };
    this.stage.addEventListener('pointermove', move);
    this.stage.addEventListener('pointerup', up);
    this.stage.addEventListener('pointercancel', up);
  }

  /** While resizing, the elements are drawn again (their shape depends on the size); cheap enough for a few. */
  renderLive(els) {
    if (els.length > 12 || !this.slideNode) return;
    const { content } = this.app;
    for (const x of els) {
      const old = this.nodeOf(x.id);
      if (!old) continue;
      const tmp = renderSlide(content, { ...this.slide, elements: [x] }, { mode: 'edit', number: this.app.index + 1 });
      const fresh = tmp.querySelector('.mo-sr-el[data-id]');
      if (fresh) { old.replaceWith(fresh); fresh.querySelectorAll('.mo-sr-text[data-fit="shrink"]').forEach(fitBox); }
    }
  }

  resizeOne(e, o, handle, p, ev) {
    const rot = ((o.rot || 0) * Math.PI) / 180;
    const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
    const cos = Math.cos(-rot), sin = Math.sin(-rot);
    const lx = (p.x - cx) * cos - (p.y - cy) * sin, ly = (p.x - cx) * sin + (p.y - cy) * cos;
    const sx = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
    const sy = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
    const fromCentre = mod(ev);
    const ax = fromCentre ? 0 : -sx * o.w / 2, ay = fromCentre ? 0 : -sy * o.h / 2;
    let w = sx ? Math.max(2, (lx - ax) * sx * (fromCentre ? 2 : 1)) : o.w;
    let h = sy ? Math.max(2, (ly - ay) * sy * (fromCentre ? 2 : 1)) : o.h;
    const keep = (e.type === 'image') !== ev.shiftKey;
    if (keep && sx && sy) { const k = Math.max(w / o.w, h / o.h); w = o.w * k; h = o.h * k; }
    else if (keep && (sx || sy) && e.type === 'image') { if (sx) h = o.h * (w / o.w); else w = o.w * (h / o.h); }
    // snap the moving edges of an element that is not rotated
    this.guides = [];
    if (!o.rot && !ev.altKey && this.app.view.snap) {
      const { xs, ys } = this.targets(new Set([e.id]));
      const tol = SNAP / this.zoom;
      if (sx && !fromCentre) {
        const edge = sx > 0 ? o.x + w : o.x + o.w - w;
        const c = xs.find(v => Math.abs(v - edge) <= tol);
        if (c !== undefined) { w = sx > 0 ? c - o.x : o.x + o.w - c; this.guides.push({ axis: 'x', at: c }); }
      }
      if (sy && !fromCentre) {
        const edge = sy > 0 ? o.y + h : o.y + o.h - h;
        const c = ys.find(v => Math.abs(v - edge) <= tol);
        if (c !== undefined) { h = sy > 0 ? c - o.y : o.y + o.h - c; this.guides.push({ axis: 'y', at: c }); }
      }
    }
    const ncx = fromCentre ? 0 : ax + sx * w / 2, ncy = fromCentre ? 0 : ay + sy * h / 2;
    const gx = cx + ncx * Math.cos(rot) - ncy * Math.sin(rot), gy = cy + ncx * Math.sin(rot) + ncy * Math.cos(rot);
    e.w = round(w);
    e.h = round(h);
    e.x = round(gx - w / 2);
    e.y = round(gy - h / 2);
    if (e.text?.autofit === 'grow' && sx && !sy) { /* the height follows the text */ }
  }

  resizeMany(sel, orig, handle, p) {
    const b = bounds([...orig.values()].map(rotatedBounds));
    const sx = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
    const sy = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
    const ax = sx > 0 ? b.x : b.x + b.w, ay = sy > 0 ? b.y : b.y + b.h;
    const kx = sx ? Math.max(0.02, (sx > 0 ? p.x - b.x : b.x + b.w - p.x) / b.w) : 1;
    const ky = sy ? Math.max(0.02, (sy > 0 ? p.y - b.y : b.y + b.h - p.y) / b.h) : 1;
    for (const e of sel) {
      const o = orig.get(e.id);
      e.x = round(sx ? ax + (o.x - ax) * kx : o.x);
      e.y = round(sy ? ay + (o.y - ay) * ky : o.y);
      e.w = round(o.w * kx);
      e.h = round(o.h * ky);
      if (e.type === 'table') { e.cols = o.cols.map(c => round(c * kx)); e.rows = o.rows.map(r => ({ ...r, h: round(r.h * ky) })); e.w = round(e.cols.reduce((a, c) => a + c, 0)); e.h = round(e.rows.reduce((a, r) => a + r.h, 0)); }
    }
  }

  rotate(sel, orig, p, ev) {
    const b = bounds([...orig.values()].map(rotatedBounds));
    const single = sel.length === 1 ? orig.get(sel[0].id) : null;
    const cx = single ? single.x + single.w / 2 : b.x + b.w / 2, cy = single ? single.y + single.h / 2 : b.y + b.h / 2;
    let angle = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
    const base = single ? single.rot || 0 : 0;
    if (ev.shiftKey) angle = Math.round(angle / 15) * 15;
    else for (const s of [0, 90, 180, 270, 360]) if (Math.abs(((angle % 360) + 360) % 360 - s) < 3) angle = s;
    const delta = single ? angle - base : angle;
    for (const e of sel) {
      const o = orig.get(e.id);
      if (!single) {
        const ox = o.x + o.w / 2 - cx, oy = o.y + o.h / 2 - cy;
        const r = (delta * Math.PI) / 180;
        const nx = cx + ox * Math.cos(r) - oy * Math.sin(r), ny = cy + ox * Math.sin(r) + oy * Math.cos(r);
        e.x = round(nx - o.w / 2);
        e.y = round(ny - o.h / 2);
      }
      const rot = ((((o.rot || 0) + delta) % 360) + 360) % 360;
      if (Math.round(rot * 100) / 100 === 0) delete e.rot; else e.rot = round(rot);
    }
  }

  moveLineEnd(e, o, which, p, ev) {
    const [x1, y1, x2, y2] = lineEnds(o);
    let px = p.x, py = p.y;
    const [fx, fy] = which === 'start' ? [x2, y2] : [x1, y1];
    if (ev.shiftKey) {
      const a = Math.round(Math.atan2(py - fy, px - fx) / (Math.PI / 4)) * (Math.PI / 4);
      const len = Math.hypot(px - fx, py - fy);
      px = fx + Math.cos(a) * len;
      py = fy + Math.sin(a) * len;
    } else if (!ev.altKey && this.app.view.snap) {
      const { xs, ys } = this.targets(new Set([e.id]));
      const tol = SNAP / this.zoom;
      const cx = xs.find(v => Math.abs(v - px) <= tol), cy = ys.find(v => Math.abs(v - py) <= tol);
      this.guides = [];
      if (cx !== undefined) { px = cx; this.guides.push({ axis: 'x', at: cx }); }
      if (cy !== undefined) { py = cy; this.guides.push({ axis: 'y', at: cy }); }
    }
    const [ax, ay, bx, by] = which === 'start' ? [px, py, x2, y2] : [x1, y1, px, py];
    e.x = round(Math.min(ax, bx));
    e.y = round(Math.min(ay, by));
    e.w = round(Math.abs(bx - ax));
    e.h = round(Math.abs(by - ay));
    if (ax > bx) e.flipH = true; else delete e.flipH;
    if (ay > by) e.flipV = true; else delete e.flipV;
  }

  // ---- keyboard, menus, drop ---------------------------------------------------------------------------------------------------

  /** Keys of the slide; also the keys pressed while nothing has the focus (the page was just opened, a menu closed). */
  key(e) {
    const app = this.app;
    if (app.text.active || (e.target !== this.stage && e.target !== document.body) || e.defaultPrevented) return;
    if (e.target === document.body && (app.view.mode !== 'normal' || app.show.running || document.querySelector('dialog[open]') || !document.getElementById('mo-backstage').hidden)) return;
    const sel = this.selected();
    const k = e.key;
    if (e.target === document.body && (k === 'Tab' || k === 'Escape')) return;
    if (k === 'Escape') { if (app.painter) app.cancelPainter(); else app.select([]); e.preventDefault(); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); app.goTo(app.index + (k === 'PageDown' ? 1 : -1)); return; }
    if (k === 'Tab') {
      e.preventDefault();
      const list = this.slide?.elements || [];
      if (!list.length) return;
      const i = list.findIndex(x => x.id === app.selection[app.selection.length - 1]);
      const next = list[(i + (e.shiftKey ? -1 : 1) + list.length) % list.length];
      app.select([next.id]);
      return;
    }
    if (app.readOnly) return;
    if ((k === 'Delete' || k === 'Backspace') && sel.length) { e.preventDefault(); app.deleteSelection(); return; }
    if (k.startsWith('Arrow') && sel.length) {
      e.preventDefault();
      const step = e.shiftKey ? 25 : mod(e) ? 1 : 5;
      const dx = k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, dy = k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0;
      app.checkpoint('nudge');
      for (const x of sel) { x.x = round(x.x + dx); x.y = round(x.y + dy); }
      app.changed();
      return;
    }
    if ((k === 'Enter' || k === 'F2') && sel.length === 1) {
      e.preventDefault();
      if (sel[0].type !== 'image' && sel[0].type !== 'line') app.text.start({ id: sel[0].id, cell: sel[0].type === 'table' ? [0, 0] : null }, { all: k === 'Enter' && sel[0].type !== 'table', end: true });
      return;
    }
    if (k.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && sel.length === 1 && sel[0].type !== 'image' && sel[0].type !== 'line' && sel[0].type !== 'table') {
      e.preventDefault();
      if (app.text.start({ id: sel[0].id }, { end: true })) app.text.insert(app.text.range(), k);
    }
  }

  context(e) {
    e.preventDefault();
    const app = this.app;
    if (app.text.active && app.text.bodyEl?.contains(e.target)) { app.textMenu(e.clientX, e.clientY); return; }
    if (app.text.active) app.text.stop();
    // a handle of the selection belongs to the selected objects
    if (e.target.closest?.('[data-handle]') && app.selection.length) { app.objectMenu(e.clientX, e.clientY, true); return; }
    const hitEl = this.hit(e.target);
    if (hitEl && !app.selection.includes(hitEl.id)) app.select(withGroups(this.slide, [hitEl.id]));
    if (!hitEl) app.select([]);
    app.objectMenu(e.clientX, e.clientY, !!hitEl);
  }

  drop(e) {
    const files = [...e.dataTransfer.files].filter(f => f.type.startsWith('image/'));
    if (!files.length || this.app.readOnly) return;
    e.preventDefault();
    const p = this.toSlide(e.clientX, e.clientY);
    this.app.insertPictures(files, p);
  }
}

/** The two ends of a line element. */
export function lineEnds(e) {
  return [e.flipH ? e.x + e.w : e.x, e.flipV ? e.y + e.h : e.y, e.flipH ? e.x : e.x + e.w, e.flipV ? e.y : e.y + e.h];
}
