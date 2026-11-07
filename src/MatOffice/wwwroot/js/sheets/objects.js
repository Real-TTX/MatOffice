/* MatOffice Sheets – charts and pictures floating over the grid: drawn in a layer that scrolls with the cells, selected by a click (eight
   handles), moved by dragging, resized at the handles (Shift keeps the proportions), removed with Delete; a chart updates with its data. */
import { el, t, shrinkImage, pickFile } from '../editor/util.js';
import { chartData, renderChart } from './charts.js';
import { areaName, referenceName } from './engine/address.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

export class Objects {
  constructor(app) {
    this.app = app;
    this.layer = app.grid.objectsMain;
    this.host = app.grid.objects;
    this.nodes = new Map();
    this.selected = null;
    this.frame = el('div', { class: 'mo-sheets-frame', hidden: true }, ...HANDLES.map(h => el('span', { class: 'mo-sheets-frame__h mo-sheets-frame__h--' + h, 'data-handle': h })));
    this.layer.append(this.frame);
    this.host.addEventListener('wheel', e => {
      if (e.ctrlKey || e.metaKey) return;
      app.grid.scroller.scrollTop += e.deltaY;
      app.grid.scroller.scrollLeft += e.deltaX;
    }, { passive: true });
    this.layer.addEventListener('pointerdown', e => this.onDown(e));
    this.layer.addEventListener('dblclick', e => {
      const node = e.target.closest('[data-object]');
      if (node && node.dataset.kind === 'chart' && !app.readOnly) app.dialogs.chart(this.find(node.dataset.object));
    });
    this.layer.addEventListener('contextmenu', e => {
      const node = e.target.closest('[data-object]');
      if (!node) return;
      e.preventDefault();
      this.select(node.dataset.object);
      app.contextMenu(e.clientX, e.clientY, this.menu());
    });
  }

  get sheet() { return this.app.sheet; }

  find(id) { return [...this.sheet.charts, ...this.sheet.images].find(o => o.id === id) || null; }
  kindOf(o) { return this.sheet.charts.includes(o) ? 'chart' : 'image'; }

  /** Builds the elements of the active sheet again (after a change of the content or of the sheet). */
  refresh() {
    const sheet = this.sheet;
    const keep = new Set();
    const dark = this.app.grid.theme().dark;
    for (const chart of sheet.charts) {
      keep.add(chart.id);
      let node = this.nodes.get(chart.id);
      if (!node || node.dataset.kind !== 'chart') {
        node?.remove();
        node = el('div', { class: 'mo-sheets-object mo-sheets-object--chart', 'data-object': chart.id, 'data-kind': 'chart', tabindex: '-1' });
        this.layer.insertBefore(node, this.frame);
        this.nodes.set(chart.id, node);
      }
      let data;
      try { data = chartData(this.app, sheet, chart); } catch { data = { categories: [], series: [], format: null }; }
      const sig = JSON.stringify([chart, data, dark, this.app.locale]);
      if (node.dataset.sig !== sig) {
        try {
          node.replaceChildren(renderChart(chart, data, { width: chart.w, height: chart.h, dark, locale: this.app.locale }));
        } catch {
          node.replaceChildren(); // a chart that cannot be drawn stays empty instead of stopping the editor
        }
        node.dataset.sig = sig;
      }
    }
    for (const image of sheet.images) {
      keep.add(image.id);
      let node = this.nodes.get(image.id);
      if (!node || node.dataset.kind !== 'image') {
        node?.remove();
        node = el('div', { class: 'mo-sheets-object mo-sheets-object--image', 'data-object': image.id, 'data-kind': 'image', tabindex: '-1' }, el('img', { alt: image.alt || '', draggable: 'false' }));
        this.layer.insertBefore(node, this.frame);
        this.nodes.set(image.id, node);
      }
      const img = node.firstChild;
      if (img.getAttribute('src') !== image.src) img.setAttribute('src', image.src);
      img.alt = image.alt || '';
    }
    for (const [id, node] of this.nodes) if (!keep.has(id)) { node.remove(); this.nodes.delete(id); }
    if (this.selected && !keep.has(this.selected)) this.deselect();
    this.position();
  }

  /** Places the objects for the scroll position and zoom (every render of the grid). */
  position(m = this.app.grid.metrics()) {
    const z = m.z;
    this.host.style.left = m.hw + 'px';
    this.host.style.top = m.hh + 'px';
    const place = (node, o) => {
      const x = (o.x >= m.fW ? o.x - m.sx : o.x) * z;
      const y = (o.y >= m.fH ? o.y - m.sy : o.y) * z;
      node.style.transform = `translate(${Math.round(x)}px, ${Math.round(y)}px) scale(${z})`;
      node.style.width = o.w + 'px';
      node.style.height = o.h + 'px';
    };
    for (const o of [...this.sheet.charts, ...this.sheet.images]) {
      const node = this.nodes.get(o.id);
      if (node) place(node, o);
    }
    const o = this.selected && this.find(this.selected);
    this.frame.hidden = !o;
    if (o) {
      const x = (o.x >= m.fW ? o.x - m.sx : o.x) * z;
      const y = (o.y >= m.fH ? o.y - m.sy : o.y) * z;
      this.frame.style.transform = `translate(${Math.round(x)}px, ${Math.round(y)}px)`;
      this.frame.style.width = Math.round(o.w * z) + 'px';
      this.frame.style.height = Math.round(o.h * z) + 'px';
    }
  }

  select(id) {
    this.selected = id;
    for (const [oid, node] of this.nodes) node.classList.toggle('is-selected', oid === id);
    this.position();
    this.app.focusGrid();
    this.app.ctx.ribbon.set(this.app.ribbonTabs());
    if (this.find(id) && this.kindOf(this.find(id)) === 'chart') this.app.ctx.ribbon.select('chart');
    this.app.ctx.ribbon.refresh();
  }

  deselect() {
    if (!this.selected) return;
    this.selected = null;
    for (const node of this.nodes.values()) node.classList.remove('is-selected');
    this.frame.hidden = true;
    this.app.ctx.ribbon.set(this.app.ribbonTabs());
    this.app.ctx.ribbon.refresh();
  }

  get selectedObject() { return this.selected ? this.find(this.selected) : null; }

  onDown(e) {
    if (e.button !== 0) return;
    const app = this.app;
    const handle = e.target.closest('[data-handle]')?.dataset.handle;
    const node = e.target.closest('[data-object]');
    const o = handle ? this.selectedObject : node ? this.find(node.dataset.object) : null;
    if (!o) return;
    e.preventDefault();
    e.stopPropagation();
    if (app.editor.active && !app.editor.commit(0, 0)) return;
    if (!handle) this.select(o.id);
    if (app.readOnly) return;
    const start = { x: e.clientX, y: e.clientY, o: { ...o } };
    const z = this.sheet.zoom;
    const sheet = this.sheet;
    app.wb.begin(handle ? 'resize object' : 'move object');
    app.wb.touchProp(sheet, this.kindOf(o) === 'chart' ? 'charts' : 'images');
    let moved = false;
    const target = e.target;
    target.setPointerCapture?.(e.pointerId);
    const move = ev => {
      const dx = (ev.clientX - start.x) / z;
      const dy = (ev.clientY - start.y) / z;
      if (Math.abs(dx) + Math.abs(dy) > 1) moved = true;
      const live = this.find(o.id);
      if (!live) return;
      if (!handle) {
        live.x = Math.max(0, Math.round(start.o.x + dx));
        live.y = Math.max(0, Math.round(start.o.y + dy));
      } else {
        let { x, y, w, h } = start.o;
        if (handle.includes('e')) w = Math.max(20, start.o.w + dx);
        if (handle.includes('s')) h = Math.max(20, start.o.h + dy);
        if (handle.includes('w')) { w = Math.max(20, start.o.w - dx); x = start.o.x + start.o.w - w; }
        if (handle.includes('n')) { h = Math.max(20, start.o.h - dy); y = start.o.y + start.o.h - h; }
        if (ev.shiftKey && handle.length === 2) { const ratio = start.o.w / start.o.h; if (w / h > ratio) w = h * ratio; else h = w / ratio; if (handle.includes('w')) x = start.o.x + start.o.w - w; if (handle.includes('n')) y = start.o.y + start.o.h - h; }
        Object.assign(live, { x: Math.max(0, Math.round(x)), y: Math.max(0, Math.round(y)), w: Math.round(w), h: Math.round(h) });
      }
      this.position();
      if (handle && this.kindOf(live) === 'chart') this.refreshSoon();
    };
    const up = () => {
      target.removeEventListener('pointermove', move);
      target.removeEventListener('pointerup', up);
      target.removeEventListener('pointercancel', up);
      if (moved) { app.wb.commit(); app.changed({ layout: true }); } else app.wb.rollback();
      this.refresh();
    };
    target.addEventListener('pointermove', move);
    target.addEventListener('pointerup', up);
    target.addEventListener('pointercancel', up);
  }

  refreshSoon() {
    if (this.soon) return;
    this.soon = requestAnimationFrame(() => { this.soon = 0; this.refresh(); });
  }

  /** Arrow keys move the selected object (Shift: by 10 pixels). */
  nudge(dr, dc, big) {
    const o = this.selectedObject;
    if (!o || this.app.readOnly) return;
    const step = big ? 10 : 1;
    this.app.commands.change('move object', () => {
      this.app.wb.touchProp(this.sheet, this.kindOf(o) === 'chart' ? 'charts' : 'images');
      const live = this.find(o.id);
      live.x = Math.max(0, live.x + dc * step);
      live.y = Math.max(0, live.y + dr * step);
    }, { protect: false });
    this.refresh();
  }

  removeSelected() {
    const o = this.selectedObject;
    if (!o || this.app.readOnly) return;
    const kind = this.kindOf(o);
    this.app.commands.change('delete object', () => {
      const prop = kind === 'chart' ? 'charts' : 'images';
      this.app.wb.setSheetProp(this.sheet, prop, this.sheet[prop].filter(x => x.id !== o.id));
    }, { protect: false });
    this.deselect();
    this.refresh();
  }

  /** Changes the selected chart (type, title, legend …). */
  updateChart(patch, chart = this.selectedObject) {
    if (!chart) return;
    this.app.commands.change('chart', () => {
      this.app.wb.touchProp(this.sheet, 'charts');
      const live = this.find(chart.id);
      for (const [k, v] of Object.entries(patch)) { if (v === undefined || v === '' || v === false) delete live[k]; else live[k] = v; }
    }, { protect: false });
    this.refresh();
  }

  /** A new chart of the selection (or of the data around the active cell). */
  insertChart(type) {
    const app = this.app;
    if (app.readOnly) return;
    const sheet = this.sheet;
    const sel = app.sel;
    const a = sel.isSingle ? app.commands.currentRegion() : app.commands.clip(sel.range);
    if (a.r1 === a.r2 && a.c1 === a.c2 && app.calc.valueAt(sheet.id, a.r1, a.c1) === null) { app.ctx.toast(t('Sheets.Chart.NoData'), 'warning'); return; }
    const x = sheet.cols.offset(a.c2 + 1) + 16;
    const y = sheet.rows.offset(a.r1);
    const chart = { id: app.wb.newObjectId('c'), type, range: areaName(a), x, y, w: 480, h: 288, title: '' };
    if (type === 'pie' || type === 'doughnut') chart.legend = 'right';
    const header = app.calc.valueAt(sheet.id, a.r1, a.c2);
    if (typeof header === 'string' && a.c2 - a.c1 === 1) chart.title = header;
    app.commands.change('insert chart', () => app.wb.setSheetProp(sheet, 'charts', [...sheet.charts, chart]), { protect: false });
    this.refresh();
    this.select(chart.id);
    app.grid.scrollIntoView(a.r1, a.c2 + 1);
  }

  /** A picture from a file (the picker, a paste): kept inside the document, made smaller when it is large. */
  async insertPicture(file) {
    const app = this.app;
    if (app.readOnly) return;
    const f = file || await pickFile('image/png,image/jpeg,image/gif,image/webp');
    if (!f) return;
    if (!/^image\/(png|jpeg|gif|webp)$/.test(f.type)) { app.ctx.toast(t('Sheets.Picture.Type'), 'danger'); return; }
    let src;
    try { src = await shrinkImage(f, 1600); } catch { app.ctx.toast(t('Sheets.Picture.Failed'), 'danger'); return; }
    if (!/^data:image\/(png|jpeg|gif|webp);base64,/.test(src)) { app.ctx.toast(t('Sheets.Picture.Type'), 'danger'); return; }
    const size = await new Promise(resolve => { const img = new Image(); img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight }); img.onerror = () => resolve({ w: 320, h: 240 }); img.src = src; });
    const scale = Math.min(1, 400 / size.w, 300 / size.h);
    const sheet = this.sheet;
    const { row, col } = app.sel.active;
    const image = { id: app.wb.newObjectId('i'), src, x: sheet.cols.offset(col), y: sheet.rows.offset(row), w: Math.round(size.w * scale), h: Math.round(size.h * scale), alt: f.name && f.name !== 'image' ? f.name.replace(/\.[^.]+$/, '') : '' };
    app.commands.change('insert picture', () => app.wb.setSheetProp(sheet, 'images', [...sheet.images, image]), { protect: false });
    this.refresh();
    this.select(image.id);
  }

  menu() {
    const app = this.app;
    const o = this.selectedObject;
    const items = [];
    if (o && this.kindOf(o) === 'chart') items.push({ label: t('Sheets.Chart.Edit'), icon: 'sheets-chart-column', run: () => app.dialogs.chart(o) }, { label: t('Sheets.Chart.SelectData'), icon: 'table', edit: false, run: () => this.showData(o) });
    if (o && this.kindOf(o) === 'image') items.push({ label: t('Sheets.Picture.Alt'), icon: 'type', run: () => app.dialogs.altText(o) });
    items.push({ label: t('Sheets.BringFront'), icon: 'sheets-bring-front', run: () => this.order(o, 1) }, { label: t('Sheets.SendBack'), icon: 'sheets-send-back', run: () => this.order(o, -1) }, '-', { label: t('Sheets.Delete'), icon: 'trash', danger: true, run: () => this.removeSelected() });
    return items;
  }

  order(o, d) {
    if (!o) return;
    const prop = this.kindOf(o) === 'chart' ? 'charts' : 'images';
    this.app.commands.change('order', () => {
      const list = this.sheet[prop].filter(x => x.id !== o.id);
      const live = this.find(o.id);
      if (d > 0) list.push(live); else list.unshift(live);
      this.app.wb.setSheetProp(this.sheet, prop, list);
    }, { protect: false });
    for (const x of this.sheet[prop]) { const n = this.nodes.get(x.id); if (n) this.layer.insertBefore(n, this.frame); }
    this.refresh();
  }

  /** Selects the data range of a chart on the grid. */
  showData(o) {
    const app = this.app;
    const ref = String(o.range || '');
    app.commands.goToText(ref.includes('!') ? ref : referenceName(this.sheet.name, { r1: 0, c1: 0, r2: 0, c2: 0 }).split('!')[0] + '!' + ref);
  }
}
