/* MatOffice Slides – the presentation editor (like PowerPoint), the entry module of the page. It keeps the state (the content, the current
   slide, the selection, the view), the history of undo/redo and every action of the ribbon, the menus and the keyboard; the parts are
   canvas.js (the slide), thumbs.js (the slide pane and the sorter), text.js (editing text), pane.js (the format pane), show.js (the show and the
   presenter view), output.js (PNG, print), ribbon.js (the tabs). The content format: Services/Formats/Slides/SlidesHandler.cs. */
import { startEditor } from '../editor/frame.js';
import { t, el, icon, debounce, shrinkImage, readAsDataUrl, pickFile, download } from '../editor/util.js';
import {
  newSlide, nextLayout, applyLayout, resetSlide, newId, clone, round, withGroups, bounds, rotatedBounds, newShape, defaultShapeSize,
  newTable, emptyBody, bodyFromText, textDefaults, isEmptyBody, slotOf, fitInto, makeTheme, themeById, layoutOf, isNumbering,
  backgroundOf, bodyText, sanitizeElements,
} from './model.js';
import { ensureStyles } from './render.js';
import { Canvas } from './canvas.js';
import { Thumbs } from './thumbs.js';
import { TextEditor, deleteRange, insertText, runAt } from './text.js';
import { Pane } from './pane.js';
import { Show } from './show.js';
import { slidePng, printDeck, askPrint } from './output.js';
import { buildRibbon, FONT_SIZES } from './ribbon.js';
import { contextMenu, layoutGallery } from './ui.js';

const CLIP_TYPE = 'application/x-matoffice-slides+json';

/** Makes the font sizes a text says itself smaller or larger (a slide size change). */
function scaleText(body, k) {
  for (const p of body?.paras || []) for (const r of p.runs) if (r.size) r.size = Math.max(1, Math.round(r.size * k * 2) / 2);
}

/** Undo by snapshots of the content; pictures (long data: addresses) are kept once in a pool, not in every snapshot. */
class Snapshots {
  constructor(limit = 120) { this.limit = limit; this.past = []; this.future = []; this.pool = new Map(); this.list = []; }
  encode(state) { return JSON.stringify(state, (k, v) => (typeof v === 'string' && v.length > 400 && v.startsWith('data:') ? this.intern(v) : v)); }
  intern(v) { let id = this.pool.get(v); if (id === undefined) { id = this.list.length; this.list.push(v); this.pool.set(v, id); } return '\u0001img:' + id; }
  decode(s) { return JSON.parse(s, (k, v) => (typeof v === 'string' && v.startsWith('\u0001img:') ? this.list[Number(v.slice(5))] : v)); }
  push(state) { this.past.push(this.encode(state)); if (this.past.length > this.limit) this.past.shift(); this.future = []; }
  undo(current) { if (!this.past.length) return null; this.future.push(this.encode(current)); return this.decode(this.past.pop()); }
  redo(current) { if (!this.future.length) return null; this.past.push(this.encode(current)); return this.decode(this.future.pop()); }
}

class SlidesApp {
  constructor() {
    this.content = null;
    this.index = 0;
    this.selection = [];
    this.slideSel = new Set();
    this.view = { mode: 'normal', notes: true, grid: false, guides: false, snap: true, gridSnap: false, pane: null };
    this.history = new Snapshots();
    this.lastKey = null;
    this.lastAt = 0;
    this.area = 'stage';
    this.painter = null;
    this.drawing = null;
    this.clip = null;
    this.phone = matchMedia('(max-width: 719px)').matches;
  }

  get slide() { return this.content?.slides[this.index] || null; }

  // ---- the frame ------------------------------------------------------------------------------------------------------------------

  mount(ctx) {
    this.ctx = ctx;
    this.readOnly = ctx.readOnly;
    this.work = ctx.work;
    this.content = this.prepare(ctx.doc.content);
    ensureStyles();
    try { Object.assign(this.view, JSON.parse(localStorage.getItem('mo.slides.view') || '{}'), { mode: 'normal', pane: null }); } catch { /* defaults */ }
    if (this.phone) this.view.notes = false;
    this.canvas = new Canvas(this);
    this.thumbs = new Thumbs(this, 'pane');
    this.sorter = new Thumbs(this, 'sorter');
    this.text = new TextEditor(this);
    this.pane = new Pane(this);
    this.show = new Show(this);
    this.buildLayout();
    this.slideSel = new Set(this.slide ? [this.slide.id] : []);
    this.ribbonKey = '';
    this.updateRibbon(true);
    this.statusBar();
    this.render();
    requestAnimationFrame(() => { this.canvas.fitZoom(); if (document.activeElement === document.body) this.canvas.focus(); });
    this.bindDocument();
    matchMedia('(max-width: 719px)').addEventListener('change', e => { this.phone = e.matches; this.layoutChanged(); this.thumbs.invalidate(); this.render(); });
  }

  /** Makes a content usable when parts are missing (an empty document, an old one). */
  prepare(c) {
    const content = c && typeof c === 'object' ? c : {};
    content.version = 1;
    content.size ||= { w: 1280, h: 720 };
    if (!content.theme?.colors) content.theme = makeTheme('paper');
    content.theme.decor ||= [];
    content.footer ||= { number: false, date: false, dateText: '', text: '', skipTitle: true };
    content.slides = Array.isArray(content.slides) ? content.slides : [];
    for (const s of content.slides) { s.elements ||= []; s.layout ||= 'blank'; }
    return content;
  }

  getContent() { if (this.text?.active) this.text.sync(); return this.content; }

  setContent(content, info = {}) {
    const slideId = this.slide?.id;
    if (this.text.active) this.text.stop();
    this.content = this.prepare(content);
    const i = this.content.slides.findIndex(s => s.id === slideId);
    this.index = Math.max(0, Math.min(i >= 0 ? i : this.index, this.content.slides.length - 1));
    this.selection = this.selection.filter(id => this.slide?.elements.some(e => e.id === id));
    if (!info.remote) this.history = new Snapshots();
    this.thumbs.invalidate();
    this.sorter.invalidate();
    this.render();
  }

  // ---- the page ---------------------------------------------------------------------------------------------------------------

  buildLayout() {
    const notes = el('textarea', { class: 'mo-slides-notes__text', 'aria-label': t('Slides.Notes'), placeholder: t('Slides.Notes.Placeholder'), readonly: this.readOnly ? true : null });
    notes.addEventListener('input', () => {
      if (!this.slide) return;
      this.checkpoint('notes:' + this.slide.id);
      if (notes.value) this.slide.notes = notes.value; else delete this.slide.notes;
      this.ctx.changed();
    });
    notes.addEventListener('focus', () => { this.area = 'notes'; if (this.text.active) this.text.stop(); });
    this.notesText = notes;
    const splitter = el('div', { class: 'mo-slides-splitter', role: 'separator', 'aria-orientation': 'horizontal', 'aria-label': t('Slides.Notes.Resize'), tabindex: '0' });
    splitter.addEventListener('pointerdown', e => this.dragNotes(e));
    splitter.addEventListener('keydown', e => { if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); this.setNotesHeight((this.view.notesHeight || 110) + (e.key === 'ArrowUp' ? 20 : -20)); } });
    this.notesBox = el('div', { class: 'mo-slides-notes' }, notes);
    this.main = el('div', { class: 'mo-slides-main' }, this.canvas.stage, splitter, this.notesBox);
    this.root = el('div', { class: 'mo-slides' }, el('div', { class: 'mo-slides-side' }, this.thumbs.list), this.main, this.pane.node);
    this.sorterBox = el('div', { class: 'mo-slides-sorter', hidden: true }, this.sorter.list);
    this.ctx.work.replaceChildren(this.root, this.sorterBox);
    this.ctx.work.addEventListener('focus', () => { if (this.text.active) this.text.refocus(); });
    this.layoutChanged();
  }

  layoutChanged() {
    const v = this.view;
    this.root.classList.toggle('has-pane', !!this.pane.kind);
    this.root.classList.toggle('no-notes', !v.notes);
    this.root.classList.toggle('is-phone', this.phone);
    this.main.style.setProperty('--notes-h', (v.notes ? v.notesHeight || 110 : 0) + 'px');
    this.root.hidden = v.mode === 'sorter';
    this.sorterBox.hidden = v.mode !== 'sorter';
    try { localStorage.setItem('mo.slides.view', JSON.stringify({ notes: v.notes, notesHeight: v.notesHeight, grid: v.grid, guides: v.guides, snap: v.snap, gridSnap: v.gridSnap })); } catch { /* private mode */ }
    if (this.canvas) { if (this.canvas.fit) requestAnimationFrame(() => this.canvas.fitZoom()); else this.canvas.layout(); }
  }

  dragNotes(e) {
    e.preventDefault();
    const start = e.clientY, h0 = this.view.notesHeight || 110;
    const move = ev => this.setNotesHeight(h0 + (start - ev.clientY));
    const up = () => { document.removeEventListener('pointermove', move); document.removeEventListener('pointerup', up); };
    document.addEventListener('pointermove', move);
    document.addEventListener('pointerup', up);
  }

  setNotesHeight(h) {
    this.view.notesHeight = Math.max(40, Math.min(window.innerHeight * 0.6, h));
    this.view.notes = true;
    this.layoutChanged();
  }

  statusBar() {
    const btn = (name, label, run, active) => {
      const b = el('button', { type: 'button', class: 'mo-sbtn mo-slides-viewbtn', 'aria-label': label, 'data-tip': label, 'aria-pressed': active() ? 'true' : 'false', onclick: () => { run(); this.statusBar(); } }, icon(name));
      return b;
    };
    this.ctx.status.right(
      btn('slides-notes', t('Slides.View.Notes'), () => this.toggleNotes(), () => this.view.notes),
      btn('slides-normal', t('Slides.View.Normal'), () => this.setView('normal'), () => this.view.mode === 'normal'),
      btn('slides-sorter', t('Slides.View.Sorter'), () => this.setView('sorter'), () => this.view.mode === 'sorter'),
      btn('slides-reading', t('Slides.View.Reading'), () => this.startShow(this.index, { reading: true }), () => false),
      btn('slides-from-current', t('Slides.Show.FromCurrent'), () => this.startShow(this.index), () => false));
    this.ctx.status.zoom({ value: this.canvas.zoom, min: 0.1, max: 4, change: z => this.canvas.setZoom(z), fit: () => this.canvas.fitZoom() });
  }

  zoomChanged() { this.ctx.status.setZoom(this.canvas.zoom); }

  statusLeft() {
    const n = this.content.slides.length;
    this.ctx.status.left(n ? t('Slides.SlideNofM', this.index + 1, n) : t('Slides.NoSlides'), this.readOnly ? t('Editor.Badge.ReadOnly') : null);
  }

  // ---- drawing everything --------------------------------------------------------------------------------------------------------

  render() {
    if (this.view.mode === 'sorter') this.sorter.render();
    else {
      this.canvas.render();
      this.thumbs.render();
    }
    if (this.notesText && document.activeElement !== this.notesText) this.notesText.value = this.slide?.notes || '';
    this.notesText.disabled = !this.slide;
    this.statusLeft();
    this.pane.render();
    this.updateRibbon();
  }

  /** After a change of the model: saved soon, drawn again. */
  changed({ all = false } = {}) {
    if (all) { this.thumbs.invalidate(); this.sorter.invalidate(); }
    else if (this.slide) { this.thumbs.invalidate(this.slide.id); this.sorter.invalidate(this.slide.id); }
    this.ctx.changed();
    this.render();
  }

  /** The text being edited changed (typing): saved soon, the box grows or shrinks its text, the thumbnail follows a moment later. */
  textChanged() {
    this.ctx.changed();
    this.canvas.afterText();
    this.thumbLater();
    this.ribbonLater();
  }

  thumbLater = debounce(() => { if (this.slide) { this.thumbs.invalidate(this.slide.id); this.sorter.invalidate(this.slide.id); this.thumbs.render(); } }, 500);

  ribbonLater = debounce(() => this.updateRibbon(), 80);

  editingStopped(id, cell) {
    if (cell) this.lastCell = { id, cell };
    this.canvas.render();
    if (this.slide) this.thumbs.invalidate(this.slide.id);
    this.thumbs.render();
    this.updateRibbon();
    this.pane.render();
  }

  /** The ribbon: set again when the contextual tabs change (Shape Format, Picture Format, Table), else only refreshed. */
  updateRibbon(force = false) {
    const sel = this.canvas?.selected() || [];
    // the galleries show the design and the slide size: a new design or size builds the ribbon again
    const look = this.content ? [this.content.theme.id, this.content.theme.variant || 0, this.content.theme.colors?.accent1, this.content.size.w, this.content.size.h] : [];
    const key = [sel.some(e => e.type === 'shape' || e.type === 'text' || e.type === 'line') ? 'shape' : '', sel.some(e => e.type === 'image') ? 'picture' : '', sel.length === 1 && sel[0].type === 'table' ? 'table' : '', ...look].join('|');
    if (force || key !== this.ribbonKey) {
      this.ribbonKey = key;
      this.ctx.ribbon.set(buildRibbon(this));
    } else this.ctx.ribbon.refresh();
  }

  selectionChanged(textOnly = false) {
    if (textOnly) { this.ribbonLater(); return; }
    this.canvas.drawSelection();
    this.pane.render();
    this.updateRibbon();
  }

  // ---- history -------------------------------------------------------------------------------------------------------------------

  state() { return { content: this.content, index: this.index, selection: this.selection }; }

  /** Takes a snapshot before a change; changes with the same key close together (typing, nudging) make one step. */
  checkpoint(key = null) {
    if (this.readOnly) return;
    const now = Date.now();
    if (key && key === this.lastKey && now - this.lastAt < 1500) { this.lastAt = now; return; }
    this.lastKey = key;
    this.lastAt = now;
    this.history.push(this.state());
  }

  restore(s) {
    if (!s) return;
    if (this.text.active) this.text.stop();
    this.content = s.content;
    this.index = Math.max(0, Math.min(s.index, this.content.slides.length - 1));
    this.selection = s.selection.filter(id => this.slide?.elements.some(e => e.id === id));
    this.slideSel = new Set(this.slide ? [this.slide.id] : []);
    this.lastKey = null;
    this.changed({ all: true });
  }

  undo() {
    if (this.readOnly) return;
    if (this.text.active) this.text.sync();
    const s = this.history.undo(this.state());
    if (s) this.restore(s);
  }

  redo() {
    if (this.readOnly) return;
    const s = this.history.redo(this.state());
    if (s) this.restore(s);
  }

  /** A change of the model: snapshot, fn, saved and drawn. */
  edit(fn, opts = {}) {
    if (this.readOnly) return;
    this.checkpoint(opts.key || null);
    const r = fn();
    this.changed(opts);
    return r;
  }

  // ---- selection and navigation --------------------------------------------------------------------------------------------------

  select(ids) {
    if (this.text.active && !ids.includes(this.text.active.id)) this.text.stop();
    this.selection = [...new Set(ids)];
    this.selectionChanged();
  }

  selectAll() {
    if (this.area === 'thumbs' || this.view.mode === 'sorter') { this.slideSel = new Set(this.content.slides.map(s => s.id)); this.render(); return; }
    if (this.slide) this.select(this.slide.elements.map(e => e.id));
  }

  goTo(i, opts = {}) {
    if (!this.content.slides.length) return;
    const index = Math.max(0, Math.min(this.content.slides.length - 1, i));
    if (this.text.active) this.text.stop();
    const changedSlide = index !== this.index;
    this.index = index;
    if (changedSlide) this.selection = [];
    if (!opts.keepSlideSel) this.slideSel = new Set([this.slide.id]);
    this.render();
  }

  selectedSlideIds() {
    const ids = this.content.slides.filter(s => this.slideSel.has(s.id)).map(s => s.id);
    return ids.length ? ids : this.slide ? [this.slide.id] : [];
  }

  setView(mode) {
    if (this.text.active) this.text.stop();
    this.view.mode = mode;
    this.layoutChanged();
    this.render();
    this.statusBar();
    if (mode === 'normal') requestAnimationFrame(() => { this.canvas.fitZoom(); this.thumbs.scrollToCurrent(); });
  }

  toggleNotes() { this.view.notes = !this.view.notes; this.layoutChanged(); }

  showPane(kind) {
    if (this.pane.kind === kind) this.pane.close();
    else this.pane.open(kind);
    this.layoutChanged();
  }

  // ---- slides ------------------------------------------------------------------------------------------------------------------------

  newSlide(layoutId) {
    this.edit(() => {
      const slide = newSlide(this.content, layoutId || nextLayout(this.slide));
      const at = this.content.slides.length ? this.index + 1 : 0;
      this.content.slides.splice(at, 0, slide);
      this.index = at;
      this.selection = [];
      this.slideSel = new Set([slide.id]);
    });
  }

  duplicateSlides() {
    const ids = this.selectedSlideIds();
    if (!ids.length) return;
    this.edit(() => {
      const last = Math.max(...ids.map(id => this.content.slides.findIndex(s => s.id === id)));
      const copies = ids.map(id => this.copySlide(this.content.slides.find(s => s.id === id)));
      this.content.slides.splice(last + 1, 0, ...copies);
      this.index = last + 1;
      this.slideSel = new Set(copies.map(s => s.id));
      this.selection = [];
    });
  }

  /** A copy of a slide with new ids (elements, groups, animations follow). */
  copySlide(slide) {
    const s = clone(slide);
    s.id = newId('s');
    const map = new Map(), groups = new Map();
    for (const e of s.elements) {
      const id = newId();
      map.set(e.id, id);
      e.id = id;
      if (e.group) { if (!groups.has(e.group)) groups.set(e.group, newId('g')); e.group = groups.get(e.group); }
    }
    if (s.animations) s.animations = s.animations.filter(a => map.has(a.el)).map(a => ({ ...a, el: map.get(a.el) }));
    return s;
  }

  deleteSlides() {
    const ids = new Set(this.selectedSlideIds());
    if (!ids.size) return;
    this.edit(() => {
      const first = this.content.slides.findIndex(s => ids.has(s.id));
      this.content.slides = this.content.slides.filter(s => !ids.has(s.id));
      this.index = Math.max(0, Math.min(first, this.content.slides.length - 1));
      this.slideSel = new Set(this.slide ? [this.slide.id] : []);
      this.selection = [];
    }, { all: true });
  }

  moveSlides(ids, target) {
    const set = new Set(ids);
    const moving = this.content.slides.filter(s => set.has(s.id));
    const before = this.content.slides.slice(0, target).filter(s => !set.has(s.id)).length;
    const rest = this.content.slides.filter(s => !set.has(s.id));
    const order = [...rest.slice(0, before), ...moving, ...rest.slice(before)];
    if (order.every((s, i) => s === this.content.slides[i])) return;
    const current = this.slide?.id;
    this.edit(() => {
      this.content.slides = order;
      this.index = Math.max(0, order.findIndex(s => s.id === current));
    }, { all: true });
  }

  setLayout(layoutId) {
    const ids = new Set(this.selectedSlideIds());
    this.edit(() => { for (const s of this.content.slides) if (ids.has(s.id)) applyLayout(this.content, s, layoutId); this.selection = []; }, { all: true });
  }

  resetSlides() {
    const ids = new Set(this.selectedSlideIds());
    this.edit(() => { for (const s of this.content.slides) if (ids.has(s.id)) resetSlide(this.content, s); }, { all: true });
  }

  toggleHidden() {
    const ids = new Set(this.selectedSlideIds());
    const hide = !this.content.slides.filter(s => ids.has(s.id)).every(s => s.hidden);
    this.edit(() => { for (const s of this.content.slides) if (ids.has(s.id)) { if (hide) s.hidden = true; else delete s.hidden; } }, { all: true });
  }

  // ---- elements ---------------------------------------------------------------------------------------------------------------

  /** Changes elements of the current slide (stops editing text first, the box is drawn again). */
  editElements(ids, fn) {
    if (this.readOnly || !this.slide) return;
    if (this.text.active) this.text.stop();
    const set = new Set(ids);
    this.edit(() => { for (const e of this.slide.elements) if (set.has(e.id)) fn(e); });
  }

  setGeometry(id, props) { this.editElements([id], e => Object.assign(e, Object.fromEntries(Object.entries(props).map(([k, v]) => [k, round(v)])))); }

  elementName(e) {
    if (e.name) return e.name;
    const text = e.text ? bodyText(e.text).trim().split('\n')[0] : '';
    if (text) return text.length > 24 ? text.slice(0, 24) + '…' : text;
    return t('Slides.Type.' + (e.type === 'shape' ? 'shape' : e.type));
  }

  textDefaultsOf(e) { return textDefaults(this.content, this.slide, e, 0); }

  /** Adds an element to the current slide, selected. */
  addElement(e, select = true) {
    if (!this.slide) this.newSlide('blank');
    this.edit(() => { this.slide.elements.push(e); if (select) this.selection = [e.id]; });
  }

  center(w, h) { return { x: round((this.content.size.w - w) / 2), y: round((this.content.size.h - h) / 2) }; }

  /** Insert → a shape: the next drag on the slide draws it, a click puts it there in its default size. */
  startDrawing(kind) {
    if (this.readOnly) return;
    if (this.text.active) this.text.stop();
    this.drawing = kind;
    this.canvas.stage.classList.add('is-drawing');
    // the hint once; afterwards the cross-hair says it
    if (!this.drawHinted) { this.drawHinted = true; this.ctx.toast(t('Slides.DrawHint'), 'info'); }
    this.canvas.focus();
  }

  cancelDrawing() { this.drawing = null; this.canvas.stage.classList.remove('is-drawing'); }

  /** Draws the shape of the drawing mode from a box (slide px). */
  drawn(box, click) {
    const kind = this.drawing;
    this.cancelDrawing();
    if (!kind) return;
    let e;
    if (kind === 'textbox') {
      e = { id: newId(), type: 'text', x: round(box.x), y: round(box.y), w: round(click ? 40 : Math.max(40, box.w)), h: 40, text: emptyBody({ autofit: 'grow', ...(click ? { wrap: false } : {}) }) };
    } else {
      const [dw, dh] = defaultShapeSize(kind);
      e = click ? newShape(kind, box.x, box.y, dw, dh) : newShape(kind, box.x, box.y, Math.max(2, box.w), kind === 'line' || kind === 'arrow' || kind === 'doubleArrow' ? box.h : Math.max(2, box.h));
      if (!click && box.flipH) e.flipH = true;
      if (!click && box.flipV) e.flipV = true;
    }
    this.addElement(e);
    if (e.type === 'text') requestAnimationFrame(() => this.text.start({ id: e.id }, { end: true }));
  }

  insertTextBox() { this.startDrawing('textbox'); }

  /** A picture file as a data: address the server takes: PNG/JPEG/GIF/WebP made smaller; any other picture (SVG, BMP …) drawn as PNG. */
  async pictureData(file) {
    if (/^image\/(png|jpeg|gif|webp)$/.test(file.type)) return shrinkImage(file, 1600);
    const url = await readAsDataUrl(file);
    const img = await new Promise((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = url; });
    const w = img.naturalWidth || 800, h = img.naturalHeight || 600;
    const k = Math.min(1, 1600 / Math.max(w, h));
    const canvas = el('canvas', { width: Math.round(w * k), height: Math.round(h * k) });
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/png');
  }

  async pickPictures() {
    const files = await pickFile('image/*', true);
    if (files?.length) this.insertPictures(files);
  }

  /** Puts pictures on the slide: into the selected (or the first) empty picture placeholder, else in the middle (or at a point). */
  async insertPictures(files, point) {
    if (this.readOnly) return;
    const images = [];
    for (const f of files.slice(0, 20)) {
      try {
        const src = await this.pictureData(f);
        const size = await new Promise(resolve => { const i = new Image(); i.onload = () => resolve([i.naturalWidth, i.naturalHeight]); i.onerror = () => resolve([400, 300]); i.src = src; });
        images.push({ src, size, alt: f.name.replace(/\.[^.]+$/, '') });
      } catch { this.ctx.toast(t('Slides.Picture.Failed'), 'danger'); }
    }
    if (!images.length) return;
    if (!this.slide) this.newSlide('blank');
    this.edit(() => {
      const added = [];
      for (const img of images) {
        const target = this.emptyPlaceholder('picture');
        const e = { id: newId(), type: 'image', src: img.src, x: 0, y: 0, w: img.size[0], h: img.size[1], alt: img.alt };
        if (target && !point) {
          const slot = slotOf(this.slide, target.ph) || { x: target.x / this.content.size.w, y: target.y / this.content.size.h, w: target.w / this.content.size.w, h: target.h / this.content.size.h };
          fitInto(e, slot, this.content.size);
          e.ph = target.ph;
          this.slide.elements.splice(this.slide.elements.indexOf(target), 1, e);
        } else {
          const maxW = this.content.size.w * 0.7, maxH = this.content.size.h * 0.7;
          const k = Math.min(1, maxW / e.w, maxH / e.h);
          e.w = round(e.w * k); e.h = round(e.h * k);
          const c = point ? { x: round(point.x - e.w / 2), y: round(point.y - e.h / 2) } : this.center(e.w, e.h);
          e.x = c.x + added.length * 16; e.y = c.y + added.length * 16;
          this.slide.elements.push(e);
        }
        added.push(e.id);
      }
      this.selection = added;
    });
  }

  /** The empty placeholder a picture or table goes into (like the office suites): the selected one, else a picture placeholder, else the first content placeholder. */
  emptyPlaceholder(kind) {
    const slide = this.slide;
    const takes = e => {
      const slot = slotOf(slide, e.ph);
      if (!slot || e.type === 'image' || e.type === 'table' || !isEmptyBody(e.text)) return false;
      return kind === 'picture' ? slot.role === 'picture' || slot.content : slot.content;
    };
    const sel = slide.elements.find(e => this.selection.includes(e.id) && takes(e));
    if (sel) return sel;
    if (kind === 'picture') { const pic = slide.elements.find(e => takes(e) && slotOf(slide, e.ph)?.role === 'picture'); if (pic) return pic; }
    return slide.elements.find(takes) || null;
  }

  /** The buttons in an empty placeholder. */
  async fillPlaceholder(id, kind) {
    this.select([id]);
    if (kind === 'picture') { const f = await pickFile('image/*'); if (f) this.insertPictures([f]); return; }
    this.insertTableDialog();
  }

  insertTable(rows, cols) {
    if (this.readOnly) return;
    if (!this.slide) this.newSlide('blank');
    const target = this.emptyPlaceholder('table');
    const { w } = this.content.size;
    const width = target ? target.w : w * 0.75;
    const rowH = 40;
    const x = target ? target.x : (w - width) / 2;
    const y = target ? target.y : (this.content.size.h - rowH * rows) / 2;
    const table = newTable(rows, cols, x, y, width, rowH);
    this.edit(() => {
      if (target) { table.ph = target.ph; this.slide.elements.splice(this.slide.elements.indexOf(target), 1, table); }
      else this.slide.elements.push(table);
      this.selection = [table.id];
    });
  }

  async insertTableDialog() {
    const cols = el('input', { type: 'number', class: 'mat-input', name: 'cols', min: 1, max: 30, value: 5 });
    const rows = el('input', { type: 'number', class: 'mat-input', name: 'rows', min: 1, max: 60, value: 2 });
    const body = el('div', { class: 'mo-slides-form' },
      el('label', { class: 'mat-field' }, el('span', { class: 'mat-field__label', text: t('Slides.Table.Cols') }), cols),
      el('label', { class: 'mat-field' }, el('span', { class: 'mat-field__label', text: t('Slides.Table.Rows') }), rows));
    const r = await this.ctx.dialog({ title: t('Slides.Table.InsertTitle'), icon: 'slides-table', body, buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }], onOpen: () => cols.select() });
    if (r !== 'ok') return;
    this.insertTable(Math.max(1, Math.min(60, Number(rows.value) || 2)), Math.max(1, Math.min(30, Number(cols.value) || 5)));
  }

  insertSymbol(ch) {
    if (this.text.active) { this.text.insert(this.text.range(), ch); return; }
    const sel = this.canvas.selected().filter(e => e.type === 'text' || e.type === 'shape');
    if (sel.length === 1 && this.text.start({ id: sel[0].id }, { end: true })) { this.text.insert(this.text.range(), ch); return; }
    const e = { id: newId(), type: 'text', ...this.center(80, 60), w: 80, h: 60, text: { paras: [{ runs: [{ t: ch, size: 32 }] }], autofit: 'grow', wrap: false } };
    this.addElement(e);
  }

  deleteSelection() {
    if (!this.slide || !this.selection.length) return;
    const ids = new Set(this.selection);
    this.edit(() => {
      const out = [];
      for (const e of this.slide.elements) {
        if (!ids.has(e.id)) { out.push(e); continue; }
        // a filled placeholder becomes empty again (its prompt comes back), like the office suites
        const slot = slotOf(this.slide, e.ph);
        if (slot && (e.type === 'image' || e.type === 'table' || !isEmptyBody(e.text))) {
          out.push({ id: newId(), type: slot.role === 'picture' ? 'shape' : 'text', ...(slot.role === 'picture' ? { shape: 'rect' } : {}), ph: slot.ph, x: round(slot.x * this.content.size.w), y: round(slot.y * this.content.size.h), w: round(slot.w * this.content.size.w), h: round(slot.h * this.content.size.h), text: emptyBody({ autofit: 'shrink' }) });
        }
      }
      this.slide.elements = out;
      if (this.slide.animations) { this.slide.animations = this.slide.animations.filter(a => out.some(e => e.id === a.el)); if (!this.slide.animations.length) delete this.slide.animations; }
      this.selection = [];
    });
  }

  /** Ctrl+D: the selected objects, or the selected slides when the slide pane (or the sorter) has the focus. */
  duplicate() {
    if (this.area === 'thumbs' || this.view.mode === 'sorter') { this.duplicateSlides(); return; }
    if (!this.selection.length) return;
    if (this.text.active) this.text.stop();
    const els = this.slide.elements.filter(e => this.selection.includes(e.id));
    this.pasteElements(clone(els), 24);
  }

  /** Ctrl+drag: copies at the places where they were dropped. */
  duplicateAt(moved) {
    const els = moved.map(m => ({ ...clone(this.slide.elements.find(e => e.id === m.id)), x: m.x, y: m.y }));
    this.pasteElements(els, 0);
  }

  /** Puts copies of elements on the current slide (new ids; offset when they would lie exactly on the originals). */
  pasteElements(els, offset = null) {
    if (this.readOnly || !els.length) return;
    if (!this.slide) this.newSlide('blank');
    const clean = sanitizeElements(els);
    if (!clean.length) return;
    this.edit(() => {
      const groups = new Map();
      const ids = [];
      const dx = offset ?? (clean.some(c => this.slide.elements.some(e => e.x === c.x && e.y === c.y && e.type === c.type)) ? 24 : 0);
      for (const e of clean) {
        e.id = newId();
        if (e.group) { if (!groups.has(e.group)) groups.set(e.group, newId('g')); e.group = groups.get(e.group); }
        if (e.ph && this.slide.elements.some(x => x.ph === e.ph) || !slotOf(this.slide, e.ph)) delete e.ph;
        e.x = round(e.x + dx);
        e.y = round(e.y + dx);
        this.slide.elements.push(e);
        ids.push(e.id);
      }
      this.selection = ids;
    });
  }

  // ---- clipboard ------------------------------------------------------------------------------------------------------------------

  clipboardPayload() {
    if (this.area === 'thumbs' || this.view.mode === 'sorter') {
      const slides = this.selectedSlideIds().map(id => clone(this.content.slides.find(s => s.id === id)));
      return { kind: 'slides', slides, text: slides.map(s => s.elements.map(e => e.text ? bodyText(e.text) : '').filter(Boolean).join('\n')).join('\n\n') };
    }
    const els = this.slide?.elements.filter(e => this.selection.includes(e.id)) || [];
    if (!els.length) return null;
    return { kind: 'elements', elements: clone(els), text: els.map(e => (e.text ? bodyText(e.text) : e.type === 'table' ? e.rows.map(r => r.cells.map(c => bodyText(c.text)).join('\t')).join('\n') : '')).filter(Boolean).join('\n') };
  }

  /** Whether the clipboard event is the editor's (not a field's, not the text being edited). */
  ownsClipboard() {
    const a = document.activeElement;
    if (this.show.running || this.text.active) return false;
    if (a && (a.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName))) return false;
    return !document.querySelector('dialog[open]');
  }

  onCopy(e, cut) {
    if (!this.ownsClipboard()) return;
    const payload = this.clipboardPayload();
    if (!payload) return;
    e.preventDefault();
    this.clip = payload;
    e.clipboardData.setData('text/plain', payload.text || ' ');
    try { e.clipboardData.setData(CLIP_TYPE, JSON.stringify(payload)); } catch { /* only the inner clipboard */ }
    if (cut && !this.readOnly) { if (payload.kind === 'slides') this.deleteSlides(); else this.deleteSelection(); }
  }

  onPaste(e) {
    if (!this.ownsClipboard() || this.readOnly) return;
    e.preventDefault();
    let payload = null;
    try { const raw = e.clipboardData.getData(CLIP_TYPE); if (raw) payload = JSON.parse(raw); } catch { payload = null; }
    const files = [...(e.clipboardData.files || [])].filter(f => f.type.startsWith('image/'));
    if (!payload && files.length) { this.insertPictures(files); return; }
    const text = e.clipboardData.getData('text/plain');
    if (!payload && this.clip && text === (this.clip.text || ' ')) payload = this.clip;
    if (payload) { this.pastePayload(payload); return; }
    if (text) this.pasteText(text);
  }

  pastePayload(p) {
    if (p.kind === 'slides' && Array.isArray(p.slides)) {
      this.edit(() => {
        const copies = p.slides.filter(s => s && Array.isArray(s.elements)).map(s => { const c = this.copySlide(s); c.elements = sanitizeElements(c.elements); return c; });
        this.content.slides.splice(this.index + 1, 0, ...copies);
        if (copies.length) { this.index += 1; this.slideSel = new Set(copies.map(s => s.id)); }
      }, { all: true });
    } else if (p.kind === 'elements' && Array.isArray(p.elements)) this.pasteElements(clone(p.elements));
  }

  pasteText(text) {
    const sel = this.canvas.selected().filter(e => e.type === 'text' || e.type === 'shape');
    if (sel.length === 1 && this.text.start({ id: sel[0].id }, { end: true })) { this.text.insert(this.text.range(), text.replace(/\r\n?/g, '\n'), true); return; }
    const e = { id: newId(), type: 'text', ...this.center(600, 60), w: 600, h: 60, text: bodyFromText(text.replace(/\r\n?/g, '\n'), { autofit: 'grow' }) };
    this.addElement(e);
  }

  copy() { this.area = this.area === 'notes' ? 'stage' : this.area; if (!document.execCommand('copy')) this.clip = this.clipboardPayload(); }

  cut() { if (!document.execCommand('cut') && !this.readOnly) { this.clip = this.clipboardPayload(); this.deleteSelection(); } }

  async paste() {
    if (this.text.active) {
      try { const text = await navigator.clipboard.readText(); if (text) this.text.insert(this.text.range(), text.replace(/\r\n?/g, '\n'), true); } catch { this.ctx.toast(t('Slides.PasteKeys'), 'info'); }
      return;
    }
    if (this.clip) { this.pastePayload(this.clip); return; }
    try { const text = await navigator.clipboard.readText(); if (text) this.pasteText(text); } catch { this.ctx.toast(t('Slides.PasteKeys'), 'info'); }
  }

  bindDocument() {
    document.addEventListener('copy', e => this.onCopy(e, false));
    document.addEventListener('cut', e => this.onCopy(e, true));
    document.addEventListener('paste', e => this.onPaste(e));
    // keyboard or pointer: the focus ring of the slide area shows only for the keyboard
    document.addEventListener('keydown', e => { if (e.key === 'Tab' || e.key === 'F6') document.body.classList.add('mo-slides-kbd'); }, true);
    document.addEventListener('pointerdown', () => document.body.classList.remove('mo-slides-kbd'), true);
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && this.drawing) { this.cancelDrawing(); e.preventDefault(); }
      if (e.key === 'Escape' && this.painter) { this.cancelPainter(); }
      if (e.target === document.body) this.canvas.key(e);
    });
  }

  // ---- arrange ---------------------------------------------------------------------------------------------------------------------

  /** Bring to front / forward, send backward / to back; a group moves as a block. */
  arrange(op) {
    if (!this.selection.length) return;
    this.editElements([], () => {});
    this.edit(() => {
      const list = this.slide.elements;
      const ids = new Set(withGroups(this.slide, this.selection));
      const sel = list.filter(e => ids.has(e.id)), rest = list.filter(e => !ids.has(e.id));
      if (op === 'front') this.slide.elements = [...rest, ...sel];
      else if (op === 'back') this.slide.elements = [...sel, ...rest];
      else {
        const first = list.findIndex(e => ids.has(e.id)), last = list.length - 1 - [...list].reverse().findIndex(e => ids.has(e.id));
        if (op === 'forward') {
          const next = list.slice(last + 1).find(e => !ids.has(e.id));
          if (!next) return;
          const block = list.filter(e => next.group ? e.group === next.group : e.id === next.id);
          const at = rest.indexOf(block[block.length - 1]) + 1;
          this.slide.elements = [...rest.slice(0, at), ...sel, ...rest.slice(at)];
        } else {
          const prev = [...list.slice(0, first)].reverse().find(e => !ids.has(e.id));
          if (!prev) return;
          const block = list.filter(e => prev.group ? e.group === prev.group : e.id === prev.id);
          const at = rest.indexOf(block[0]);
          this.slide.elements = [...rest.slice(0, at), ...sel, ...rest.slice(at)];
        }
      }
    });
  }

  group() {
    const sel = this.canvas.selected();
    if (sel.length < 2) return;
    this.edit(() => {
      const g = newId('g');
      const ids = new Set(sel.map(e => e.id));
      for (const e of sel) e.group = g;
      // the members stand together where the topmost was
      const list = this.slide.elements;
      const top = list.map((e, i) => (ids.has(e.id) ? i : -1)).filter(i => i >= 0).pop();
      const rest = list.filter(e => !ids.has(e.id));
      const at = rest.indexOf(list.slice(top + 1).find(e => !ids.has(e.id)));
      const members = list.filter(e => ids.has(e.id));
      this.slide.elements = at < 0 ? [...rest, ...members] : [...rest.slice(0, at), ...members, ...rest.slice(at)];
    });
  }

  ungroup() {
    const sel = this.canvas.selected().filter(e => e.group);
    if (!sel.length) return;
    this.edit(() => { for (const e of sel) delete e.group; });
  }

  /** Align: to the slide when one unit is selected (or toSlide), else to the selection. */
  align(kind, toSlide = null) {
    const units = this.canvas.units();
    if (!units.length) return;
    const all = units.flat();
    const slideBox = { x: 0, y: 0, w: this.content.size.w, h: this.content.size.h };
    const ref = (toSlide ?? (units.length === 1)) ? slideBox : bounds(all.map(rotatedBounds));
    this.edit(() => {
      for (const u of units) {
        const b = bounds(u.map(rotatedBounds));
        let dx = 0, dy = 0;
        if (kind === 'left') dx = ref.x - b.x;
        if (kind === 'center') dx = ref.x + ref.w / 2 - (b.x + b.w / 2);
        if (kind === 'right') dx = ref.x + ref.w - (b.x + b.w);
        if (kind === 'top') dy = ref.y - b.y;
        if (kind === 'middle') dy = ref.y + ref.h / 2 - (b.y + b.h / 2);
        if (kind === 'bottom') dy = ref.y + ref.h - (b.y + b.h);
        for (const e of u) { e.x = round(e.x + dx); e.y = round(e.y + dy); }
      }
    });
  }

  distribute(axis) {
    const units = this.canvas.units();
    if (units.length < 3) { if (units.length) this.align(axis === 'h' ? 'center' : 'middle', true); return; }
    const boxes = units.map(u => ({ u, b: bounds(u.map(rotatedBounds)) })).sort((a, b) => (axis === 'h' ? a.b.x - b.b.x : a.b.y - b.b.y));
    const first = boxes[0].b, last = boxes[boxes.length - 1].b;
    const total = boxes.reduce((s, x) => s + (axis === 'h' ? x.b.w : x.b.h), 0);
    const span = axis === 'h' ? last.x + last.w - first.x : last.y + last.h - first.y;
    const gap = (span - total) / (boxes.length - 1);
    this.edit(() => {
      let pos = axis === 'h' ? first.x : first.y;
      for (const { u, b } of boxes) {
        const d = pos - (axis === 'h' ? b.x : b.y);
        for (const e of u) { if (axis === 'h') e.x = round(e.x + d); else e.y = round(e.y + d); }
        pos += (axis === 'h' ? b.w : b.h) + gap;
      }
    });
  }

  flip(axis) { this.editElements(this.selection, e => { const k = axis === 'h' ? 'flipH' : 'flipV'; if (e[k]) delete e[k]; else e[k] = true; }); }

  rotateBy(deg) { this.editElements(this.selection, e => { if (e.type === 'table') return; const r = ((((e.rot || 0) + deg) % 360) + 360) % 360; if (r) e.rot = r; else delete e.rot; }); }

  // ---- formatting of shapes ------------------------------------------------------------------------------------------------------

  shapeTargets() { return this.canvas.selected().filter(e => e.type === 'shape' || e.type === 'text'); }

  setFill(color) {
    const ids = this.shapeTargets().map(e => e.id);
    if (!ids.length) return;
    this.editElements(ids, e => { e.fill = color === 'none' || color === null ? { type: 'none' } : { type: 'solid', color }; });
  }

  setLine(props) {
    const ids = this.canvas.selected().filter(e => e.type !== 'table').map(e => e.id);
    this.editElements(ids, e => {
      if (props === null) { if (e.type !== 'line') delete e.line; return; }
      e.line = { color: 'dk1', width: 1, dash: 'solid', ...(e.line || {}), ...props };
    });
  }

  toggleShadow() {
    const sel = this.canvas.selected().filter(e => e.type !== 'table');
    const on = !sel.every(e => e.shadow);
    this.editElements(sel.map(e => e.id), e => { if (on) e.shadow = true; else delete e.shadow; });
  }

  quickStyle(style) {
    this.editElements(this.shapeTargets().map(e => e.id), e => {
      e.fill = clone(style.fill);
      if (style.line) e.line = clone(style.line); else delete e.line;
      if (style.shadow) e.shadow = true; else delete e.shadow;
      if (e.text) for (const p of e.text.paras) for (const r of p.runs) r.color = style.text;
    });
  }

  changeShape(kind) { this.editElements(this.canvas.selected().filter(e => e.type === 'shape' || e.type === 'text').map(e => e.id), e => { e.type = 'shape'; e.shape = kind; delete e.adj; }); }

  setRadius(v) { this.editElements(this.canvas.selected().filter(e => e.type === 'shape').map(e => e.id), e => { if (e.shape === 'rect' && v > 0) e.shape = 'roundRect'; e.adj = v; }); }

  // ---- formatting of text --------------------------------------------------------------------------------------------------------

  /** The text bodies of the selection (all cells of a table). */
  bodies() {
    const out = [];
    for (const e of this.canvas.selected()) {
      if (e.text) out.push({ e, body: e.text });
      if (e.type === 'table') e.rows.forEach((r, ri) => r.cells.forEach((c, ci) => out.push({ e, body: c.text, cell: [ri, ci] })));
    }
    return out;
  }

  /** The look of the text at the caret, or of the first selected text. */
  fmt() {
    const slide = this.slide;
    if (!slide) return {};
    let e, para, run, cell = null, body;
    if (this.text.active) {
      e = this.text.el;
      if (!e) return {};
      body = this.text.body;
      run = this.text.runFormat();
      para = this.text.selectedParagraphs()[0] || body.paras[0];
      cell = this.text.active.cell;
    } else {
      const b = this.bodies()[0];
      if (!b) return {};
      ({ e, body } = b);
      cell = b.cell || null;
      para = body.paras[0];
      run = para.runs.find(r => r.t) || para.runs[0] || {};
    }
    const look = cell ? this.canvas.cellLook(e, cell) : null;
    const d = textDefaults(this.content, slide, e, para?.level || 0, look);
    return {
      font: run.font || d.font, size: run.size || d.size, b: run.b ?? d.bold, i: !!run.i, u: !!run.u, s: !!run.s, color: run.color || null, hl: run.hl || null,
      align: para?.align || d.align, bullet: para?.bullet ?? d.bullet, level: para?.level || 0, lineSpacing: para?.lineSpacing ?? d.lineSpacing,
      anchor: body?.anchor || d.anchor, autofit: body?.autofit || d.autofit, hasText: true,
    };
  }

  hasText() { return this.text.active || this.bodies().length > 0; }

  setRun(props) {
    if (this.readOnly) return;
    if (this.text.active) { this.text.applyRun(props); return; }
    const bodies = this.bodies();
    if (!bodies.length) return;
    this.edit(() => {
      for (const { body } of bodies) for (const p of body.paras) for (const r of p.runs) for (const [k, v] of Object.entries(props)) { if (v === null || v === undefined) delete r[k]; else r[k] = v; }
    });
  }

  toggleRun(key) { const f = this.fmt(); this.setRun({ [key]: !f[key] }); }

  growFont(dir) {
    const f = this.fmt();
    const cur = f.size || 18;
    const next = dir > 0 ? FONT_SIZES.find(s => s > cur) ?? cur + 8 : [...FONT_SIZES].reverse().find(s => s < cur) ?? Math.max(1, cur - 1);
    this.setRun({ size: next });
  }

  clearFormat() {
    if (this.text.active) { this.text.applyRun({ b: null, i: null, u: null, s: null, font: null, size: null, color: null, hl: null }); return; }
    this.setRun({ b: null, i: null, u: null, s: null, font: null, size: null, color: null, hl: null });
  }

  /** Changes the selected paragraphs (or all of the selected boxes). */
  setPara(fn) {
    if (this.readOnly) return;
    if (this.text.active) { this.text.applyPara(fn); return; }
    const bodies = this.bodies();
    if (!bodies.length) return;
    this.edit(() => { for (const { body } of bodies) body.paras.forEach(fn); });
  }

  setAlign(align) { this.setPara(p => { p.align = align; }); }

  setBullet(b) { this.setPara(p => { p.bullet = b; }); }

  toggleBullets(kind = 'bullet') {
    const f = this.fmt();
    const on = kind === 'number' ? isNumbering(f.bullet) : f.bullet && f.bullet !== 'none' && !isNumbering(f.bullet);
    this.setBullet(on ? 'none' : kind === 'number' ? '1.' : '•');
  }

  changeLevel(d) {
    if (this.text.active) { this.text.changeLevel(d); return; }
    this.setPara(p => { const l = Math.min(8, Math.max(0, (p.level || 0) + d)); if (l) p.level = l; else delete p.level; });
  }

  setLineSpacing(v) { this.setPara(p => { p.lineSpacing = v; }); }

  setBodyProp(key, value) {
    if (this.readOnly) return;
    if (this.text.active) this.text.stop();
    const bodies = this.bodies();
    if (!bodies.length) return;
    this.edit(() => { for (const { body } of bodies) body[key] = value; });
  }

  // ---- tables -------------------------------------------------------------------------------------------------------------------------

  table() { const sel = this.canvas.selected(); return sel.length === 1 && sel[0].type === 'table' ? sel[0] : null; }

  /** The cell a table command works on: the one being edited, else the last one edited in this table, else the fallback (e.g. the last row). */
  cell(tb, fallback) {
    if (this.text.active?.cell) return this.text.active.cell;
    if (tb && this.lastCell?.id === tb.id) return [Math.min(this.lastCell.cell[0], tb.rows.length - 1), Math.min(this.lastCell.cell[1], tb.cols.length - 1)];
    return fallback ? fallback(tb) : [0, 0];
  }

  tableEdit(fn, fallback) {
    const tb = this.table();
    if (!tb || this.readOnly) return;
    const cell = this.cell(tb, fallback);
    if (this.text.active) this.text.stop();
    this.edit(() => {
      fn(tb, cell);
      if (!tb.rows.length || !tb.cols.length) { this.slide.elements = this.slide.elements.filter(e => e !== tb); this.selection = []; return; }
      tb.w = round(tb.cols.reduce((a, b) => a + b, 0));
      tb.h = round(tb.rows.reduce((a, r) => a + r.h, 0));
    });
  }

  insertRow(below) {
    this.tableEdit((tb, [r]) => { const h = tb.rows[r]?.h || 40; tb.rows.splice(below ? r + 1 : r, 0, { h, cells: tb.cols.map(() => ({ text: emptyBody() })) }); }, tb => [below ? tb.rows.length - 1 : 0, 0]);
  }

  insertCol(right) {
    this.tableEdit((tb, [, c]) => {
      const w = tb.cols[c] || 100;
      const at = right ? c + 1 : c;
      tb.cols.splice(at, 0, w);
      for (const row of tb.rows) row.cells.splice(at, 0, { text: emptyBody() });
    }, tb => [0, right ? tb.cols.length - 1 : 0]);
  }

  deleteRow() { this.tableEdit((tb, [r]) => { tb.rows.splice(r, 1); }, tb => [tb.rows.length - 1, 0]); }

  deleteCol() { this.tableEdit((tb, [, c]) => { tb.cols.splice(c, 1); for (const row of tb.rows) row.cells.splice(c, 1); }, tb => [0, tb.cols.length - 1]); }

  deleteTable() { if (this.table()) { if (this.text.active) this.text.stop(); this.deleteSelection(); } }

  distributeRows() { this.tableEdit(tb => { const h = round(tb.rows.reduce((a, r) => a + r.h, 0) / tb.rows.length); for (const r of tb.rows) r.h = h; }); }

  distributeCols() { this.tableEdit(tb => { const w = round(tb.cols.reduce((a, b) => a + b, 0) / tb.cols.length); tb.cols = tb.cols.map(() => w); }); }

  setTableProp(props) { this.tableEdit(tb => Object.assign(tb, props)); }

  cellFill(color) {
    const tb = this.table();
    if (!tb) return;
    const editingCell = this.text.active?.cell;
    this.tableEdit(t2 => {
      const apply = c => { if (color === null) delete c.fill; else c.fill = color === 'none' ? { type: 'none' } : { type: 'solid', color }; };
      if (editingCell) apply(t2.rows[editingCell[0]].cells[editingCell[1]]);
      else for (const r of t2.rows) for (const c of r.cells) apply(c);
    });
  }

  /** Tab in a table: the next cell (at the end a new row), Shift+Tab the previous one. */
  moveCell(dir) {
    const tb = this.table();
    const a = this.text.active;
    if (!tb || !a?.cell) return;
    let [r, c] = a.cell;
    c += dir;
    if (c >= tb.cols.length) { c = 0; r += 1; }
    if (c < 0) { c = tb.cols.length - 1; r -= 1; }
    if (r < 0) return;
    if (r >= tb.rows.length) {
      this.text.stop();
      this.edit(() => { tb.rows.push({ h: tb.rows[tb.rows.length - 1].h, cells: tb.cols.map(() => ({ text: emptyBody() })) }); tb.h = round(tb.rows.reduce((s, x) => s + x.h, 0)); });
    } else this.text.stop();
    this.text.start({ id: tb.id, cell: [r, c] }, { all: true });
  }

  // ---- design --------------------------------------------------------------------------------------------------------------------

  setTheme(id) {
    this.edit(() => { this.content.theme = makeTheme(id, 0); }, { all: true });
  }

  setVariant(v) {
    const base = themeById(this.content.theme.id) ? this.content.theme.id : null;
    this.edit(() => {
      if (base) {
        const keep = { decor: this.content.theme.decor };
        this.content.theme = makeTheme(base, v);
        if (!themeById(base)) Object.assign(this.content.theme, keep);
      }
    }, { all: true });
  }

  /** 16:9 or 4:3, like "Ensure fit" of the office suites: placeholders take the places of their layout, everything else is scaled evenly and centred. */
  setSize(kind) {
    const w = kind === '4:3' ? 960 : 1280, h = 720;
    const old = this.content.size;
    if (old.w === w && old.h === h) return;
    const k = Math.min(w / old.w, h / old.h);
    const dx = (w - old.w * k) / 2, dy = (h - old.h * k) / 2;
    const near = (a, b) => Math.abs(a - b) < 1.5;
    this.edit(() => {
      for (const s of this.content.slides) {
        const layout = layoutOf(s.layout);
        for (const e of s.elements) {
          const slot = e.ph ? layout.slots.find(x => x.ph === e.ph) : null;
          if (slot && e.type !== 'image' && e.type !== 'table' && !e.rot
            && near(e.x, slot.x * old.w) && near(e.y, slot.y * old.h) && near(e.w, slot.w * old.w) && near(e.h, slot.h * old.h)) {
            Object.assign(e, { x: round(slot.x * w), y: round(slot.y * h), w: round(slot.w * w), h: round(slot.h * h) });
            continue;
          }
          Object.assign(e, { x: round(e.x * k + dx), y: round(e.y * k + dy), w: round(e.w * k), h: round(e.h * k) });
          if (k === 1) continue;
          if (e.type === 'table') {
            e.cols = e.cols.map(c => round(c * k));
            for (const r of e.rows) { r.h = round(r.h * k); for (const c of r.cells) scaleText(c.text, k); }
            e.w = round(e.cols.reduce((a, b) => a + b, 0));
            e.h = round(e.rows.reduce((a, r) => a + r.h, 0));
          } else if (e.text && !slot) scaleText(e.text, k);
        }
      }
      this.content.size = { w, h };
    }, { all: true });
    requestAnimationFrame(() => this.canvas.fitZoom());
  }

  backgroundOfSlide() { return clone(backgroundOf(this.content, this.slide)); }

  /** A background for the current slide(s) or every slide; null takes the design's again. */
  setBackground(fill, all) {
    const ids = new Set(this.selectedSlideIds());
    this.edit(() => {
      for (const s of this.content.slides) {
        if (!all && !ids.has(s.id)) continue;
        if (fill) s.background = clone(fill); else delete s.background;
      }
      if (all && fill) { this.content.theme.background = clone(fill); for (const s of this.content.slides) delete s.background; if (this.content.theme.groups) for (const g of Object.values(this.content.theme.groups)) delete g.background; }
    }, { all: true });
  }

  // ---- transitions and animations --------------------------------------------------------------------------------------------------

  transitionOf() { return this.slide?.transition || { type: 'none', duration: 0.7 }; }

  setTransition(props) {
    const ids = new Set(this.selectedSlideIds());
    this.edit(() => {
      for (const s of this.content.slides) {
        if (!ids.has(s.id)) continue;
        const tr = { type: 'none', duration: 0.7, ...(s.transition || {}), ...props };
        if (props.type && props.type !== s.transition?.type) delete tr.dir;
        if (tr.auto === null) delete tr.auto;
        if (tr.type === 'none' && tr.auto === undefined) delete s.transition; else s.transition = tr;
      }
    }, { all: true });
    if (props.type && props.type !== 'none') this.previewTransition();
  }

  transitionToAll() {
    const tr = this.slide?.transition;
    this.edit(() => { for (const s of this.content.slides) { if (tr) s.transition = clone(tr); else delete s.transition; } }, { all: true });
  }

  previewTransition() {
    const layer = this.canvas.host.firstChild;
    const tr = this.slide?.transition;
    if (!layer || !tr || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const ms = (tr.duration || 0.7) * 1000;
    const frames = { fade: [{ opacity: 0 }, { opacity: 1 }], zoom: [{ scale: 0.3, opacity: 0 }, { scale: 1, opacity: 1 }], wipe: [{ clipPath: 'inset(0 0 0 100%)' }, { clipPath: 'inset(0 0 0 0)' }], split: [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }], push: [{ translate: '100% 0' }, { translate: '0 0' }], cover: [{ translate: '100% 0' }, { translate: '0 0' }] }[tr.type];
    if (frames) layer.animate(frames, { duration: ms, easing: 'ease-out' });
  }

  animationOf(id) { return this.slide?.animations?.find(a => a.el === id) || null; }

  /** Gives the selected elements an entrance effect ('none' takes it away). */
  setAnimation(effect) {
    const ids = withGroups(this.slide, this.selection);
    if (!ids.length) return;
    this.edit(() => {
      const list = this.slide.animations || [];
      for (const id of ids) {
        const i = list.findIndex(a => a.el === id);
        if (effect === 'none') { if (i >= 0) list.splice(i, 1); continue; }
        const a = { el: id, effect, trigger: 'click', duration: effect === 'appear' ? 0.5 : 0.5, ...(effect === 'fly' || effect === 'wipe' ? { dir: 'bottom' } : {}) };
        if (i >= 0) list[i] = { ...list[i], effect, ...(effect === 'fly' || effect === 'wipe' ? { dir: list[i].dir || 'bottom' } : {}) };
        else list.push(a);
      }
      if (list.length) this.slide.animations = list; else delete this.slide.animations;
    });
    if (effect !== 'none') this.previewAnimations(ids);
  }

  editAnimations(fn) {
    const ids = new Set(withGroups(this.slide, this.selection));
    this.edit(() => { for (const a of this.slide.animations || []) if (ids.has(a.el)) fn(a); });
  }

  editAnimation(i, fn) { this.edit(() => { const a = this.slide.animations?.[i]; if (a) fn(a); }); }

  moveAnimation(i, d) {
    this.edit(() => { const l = this.slide.animations; const j = i + d; if (!l || j < 0 || j >= l.length) return; [l[i], l[j]] = [l[j], l[i]]; });
  }

  removeAnimation(i) { this.edit(() => { this.slide.animations?.splice(i, 1); if (!this.slide.animations?.length) delete this.slide.animations; }); }

  /** Plays the effects of the slide (or of some elements) on the canvas. */
  previewAnimations(only) {
    const list = (this.slide?.animations || []).filter(a => !only || only.includes(a.el));
    if (!list.length || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    let delay = 0;
    for (const a of list) {
      const node = this.canvas.nodeOf(a.el);
      if (!node) continue;
      const frames = a.effect === 'fade' ? [{ opacity: 0 }, { opacity: 1 }] : a.effect === 'zoom' ? [{ opacity: 0, scale: 0.2 }, { opacity: 1, scale: 1 }]
        : a.effect === 'fly' ? [{ translate: { bottom: '0 300px', top: '0 -300px', left: '-300px 0', right: '300px 0' }[a.dir || 'bottom'], opacity: 0 }, { translate: '0 0', opacity: 1 }]
          : a.effect === 'wipe' ? [{ clipPath: { bottom: 'inset(100% 0 0 0)', top: 'inset(0 0 100% 0)', left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)' }[a.dir || 'bottom'] }, { clipPath: 'inset(0 0 0 0)' }]
            : [{ opacity: 0 }, { opacity: 0, offset: 0.99 }, { opacity: 1 }];
      const ms = (a.duration || 0.5) * 1000;
      node.animate(frames, { duration: ms, delay, fill: 'backwards', easing: 'ease-out' });
      delay += a.trigger === 'with' ? 0 : ms * 0.8;
    }
  }

  // ---- the show ------------------------------------------------------------------------------------------------------------------

  startShow(from = 0, opts = {}) {
    if (!this.content.slides.length) return;
    if (this.text.active) this.text.stop();
    this.show.start(from, opts);
  }

  showEnded(lastIndex) {
    if (this.view.mode === 'normal' && lastIndex !== undefined && lastIndex !== this.index) this.goTo(lastIndex);
    this.canvas.focus();
  }

  // ---- format painter -------------------------------------------------------------------------------------------------------------

  startPainter() {
    const sel = this.canvas.selected()[0];
    if (!sel) return;
    if (this.text.active) {
      this.painter = { run: this.text.runFormat() };
      this.text.stop();
    } else {
      const run = sel.text?.paras[0]?.runs.find(r => r.t) || sel.text?.paras[0]?.runs[0];
      this.painter = { fill: clone(sel.fill), line: clone(sel.line), shadow: sel.shadow, opacity: sel.opacity, run: run ? (({ t: _, ...r }) => r)(run) : null, para: sel.text?.paras[0] ? (({ runs: _, ...p }) => p)(sel.text.paras[0]) : null, type: sel.type };
    }
    this.canvas.stage.classList.add('is-painting');
    this.ctx.ribbon.refresh();
  }

  cancelPainter() { this.painter = null; this.canvas.stage.classList.remove('is-painting'); this.ctx.ribbon.refresh(); }

  paintFormat(id) {
    const p = this.painter;
    this.cancelPainter();
    if (!p) return;
    this.select(withGroups(this.slide, [id]));
    this.editElements(withGroups(this.slide, [id]), e => {
      if (p.type && e.type !== 'line' && e.type !== 'image' && e.type !== 'table' && (p.type === 'shape' || p.type === 'text')) {
        if (p.fill) e.fill = clone(p.fill); else delete e.fill;
        if (p.line) e.line = clone(p.line); else delete e.line;
        if (p.shadow) e.shadow = true; else delete e.shadow;
        if (p.opacity !== undefined) e.opacity = p.opacity; else delete e.opacity;
      }
      if (e.type === 'line' && p.line) e.line = clone(p.line);
      const bodies = e.text ? [e.text] : e.type === 'table' ? e.rows.flatMap(r => r.cells.map(c => c.text)) : [];
      for (const body of bodies) for (const para of body.paras) {
        if (p.para) for (const k of ['align', 'lineSpacing']) { if (p.para[k] !== undefined) para[k] = p.para[k]; }
        if (p.run) for (const r of para.runs) { for (const k of ['b', 'i', 'u', 's', 'font', 'size', 'color', 'hl']) { if (p.run[k] !== undefined) r[k] = p.run[k]; else delete r[k]; } }
      }
    });
  }

  // ---- find and replace -------------------------------------------------------------------------------------------------------------

  /** The find (and replace) box: modeless, at the top right of the slide like the office suites. */
  find(replace = false) {
    if (this.findBox) { this.findBox.querySelector('input').select(); this.findBox.classList.toggle('has-replace', replace || this.findBox.classList.contains('has-replace')); return; }
    const what = el('input', { type: 'search', class: 'mat-input', 'aria-label': t('Slides.Find.What'), placeholder: t('Slides.Find.What') });
    const by = el('input', { type: 'text', class: 'mat-input mo-slides-find__replace', 'aria-label': t('Slides.Find.With'), placeholder: t('Slides.Find.With') });
    const caseBox = el('input', { type: 'checkbox' });
    const info = el('span', { class: 'mo-slides-find__info', 'aria-live': 'polite' });
    const btn = (label, run, cls = 'mat-btn--secondary') => el('button', { type: 'button', class: 'mat-btn mat-btn--sm ' + cls, onclick: run }, el('span', { text: label }));
    const opts = () => ({ text: what.value, matchCase: caseBox.checked });
    const box = el('div', { class: 'mo-slides-find' + (replace ? ' has-replace' : ''), role: 'dialog', 'aria-label': t('Slides.Find') },
      el('div', { class: 'mo-slides-find__row' }, what, btn(t('Slides.Find.Next'), () => { info.textContent = this.findNext(opts()) ? '' : t('Slides.Find.NotFound'); }, 'mat-btn--primary'),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Editor.Close'), onclick: () => close() }, icon('x'))),
      el('div', { class: 'mo-slides-find__row mo-slides-find__replace-row' }, by,
        btn(t('Slides.Find.Replace'), () => { info.textContent = this.replaceOne(opts(), by.value) ? '' : t('Slides.Find.NotFound'); }),
        btn(t('Slides.Find.ReplaceAll'), () => { const n = this.replaceAll(opts(), by.value); info.textContent = t('Slides.Find.Replaced', n); })),
      el('div', { class: 'mo-slides-find__row' }, el('label', { class: 'mat-check' }, caseBox, el('span', { text: t('Slides.Find.MatchCase') })), info,
        replace ? null : btn(t('Slides.Find.Replace') + ' …', () => box.classList.add('has-replace'), 'mat-btn--ghost')));
    const close = () => { box.remove(); this.findBox = null; this.lastFind = null; };
    what.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); info.textContent = this.findNext(opts()) ? '' : t('Slides.Find.NotFound'); } if (e.key === 'Escape') close(); });
    by.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
    if (this.readOnly) by.disabled = true;
    const sel = this.text.active ? this.text.selectedText() : '';
    if (sel && !sel.includes('\n')) what.value = sel;
    this.findBox = box;
    this.work.append(box);
    what.focus();
  }

  /** Every place of the text: { slide, id, cell, p, o } in reading order. */
  matches({ text, matchCase }) {
    if (!text) return [];
    const needle = matchCase ? text : text.toLowerCase();
    const out = [];
    this.content.slides.forEach((s, si) => {
      for (const e of s.elements) {
        const bodies = e.text ? [[null, e.text]] : e.type === 'table' ? e.rows.flatMap((r, ri) => r.cells.map((c, ci) => [[ri, ci], c.text])) : [];
        for (const [cell, body] of bodies) {
          body?.paras.forEach((p, pi) => {
            const hay = p.runs.map(r => r.t).join('');
            const h = matchCase ? hay : hay.toLowerCase();
            let i = h.indexOf(needle);
            while (i >= 0) { out.push({ si, id: e.id, cell, p: pi, o: i, len: text.length }); i = h.indexOf(needle, i + Math.max(1, needle.length)); }
          });
        }
      }
    });
    return out;
  }

  findNext(opts) {
    const all = this.matches(opts);
    if (!all.length) return false;
    const cur = this.lastFind;
    const key = m => [m.si, this.content.slides[m.si].elements.findIndex(e => e.id === m.id), m.cell ? m.cell[0] * 1000 + m.cell[1] : -1, m.p, m.o];
    const after = m => { if (!cur) return m.si >= this.index; const a = key(m), b = key(cur); for (let i = 0; i < a.length; i++) { if (a[i] !== b[i]) return a[i] > b[i]; } return false; };
    const m = all.find(after) || all[0];
    this.lastFind = m;
    this.showMatch(m);
    return true;
  }

  showMatch(m) {
    if (this.index !== m.si) this.goTo(m.si);
    this.selection = [m.id];
    this.text.start({ id: m.id, cell: m.cell }, {});
    this.text.setRange({ start: { p: m.p, o: m.o }, end: { p: m.p, o: m.o + m.len } });
    this.findBox?.querySelector('input')?.focus();
  }

  replaceOne(opts, by) {
    const m = this.lastFind;
    if (m && this.text.active && this.text.active.id === m.id && String(this.text.active.cell) === String(m.cell)) {
      const sel = this.text.body.paras[m.p]?.runs.map(r => r.t).join('').slice(m.o, m.o + m.len);
      if (sel !== undefined && (opts.matchCase ? sel === opts.text : sel.toLowerCase() === opts.text.toLowerCase())) {
        this.text.insert({ start: { p: m.p, o: m.o }, end: { p: m.p, o: m.o + m.len } }, by);
        this.lastFind = { ...m, o: m.o + by.length - 1 };
      }
    }
    return this.findNext(opts);
  }

  replaceAll(opts, by) {
    const all = this.matches(opts);
    if (!all.length || this.readOnly) return 0;
    if (this.text.active) this.text.stop();
    this.checkpoint();
    for (const m of [...all].reverse()) {
      const e = this.content.slides[m.si].elements.find(x => x.id === m.id);
      const body = m.cell ? e.rows[m.cell[0]].cells[m.cell[1]].text : e.text;
      const props = runAt(body.paras[m.p], m.o + 1);
      const at = deleteRange(body.paras, { p: m.p, o: m.o }, { p: m.p, o: m.o + m.len });
      if (by) insertText(body.paras, at, by, props);
    }
    this.lastFind = null;
    this.changed({ all: true });
    return all.length;
  }

  // ---- header and footer -------------------------------------------------------------------------------------------------------

  async headerFooter() {
    const f = { ...this.content.footer };
    const date = el('input', { type: 'checkbox', checked: f.date ? true : null });
    const fixed = el('input', { type: 'text', class: 'mat-input', value: f.dateText || '', placeholder: t('Slides.Footer.DateAuto') });
    const number = el('input', { type: 'checkbox', checked: f.number ? true : null });
    const footer = el('input', { type: 'checkbox', checked: f.text ? true : null });
    const text = el('input', { type: 'text', class: 'mat-input', value: f.text || '', maxlength: 300 });
    const skip = el('input', { type: 'checkbox', checked: f.skipTitle !== false ? true : null });
    const body = el('div', { class: 'mo-slides-form' },
      el('label', { class: 'mat-check' }, date, el('span', { text: t('Slides.Footer.Date') })), el('label', { class: 'mat-field' }, el('span', { class: 'mat-field__label', text: t('Slides.Footer.DateFixed') }), fixed),
      el('label', { class: 'mat-check' }, number, el('span', { text: t('Slides.Footer.Number') })),
      el('label', { class: 'mat-check' }, footer, el('span', { text: t('Slides.Footer.Text') })), text,
      el('label', { class: 'mat-check' }, skip, el('span', { text: t('Slides.Footer.SkipTitle') })));
    text.addEventListener('input', () => { footer.checked = !!text.value; });
    const r = await this.ctx.dialog({ title: t('Slides.Footer.Title'), icon: 'slides-header-footer', body, size: 'md', buttons: [{ label: t('Slides.Footer.ApplyAll'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r !== 'ok') return;
    this.edit(() => { this.content.footer = { number: number.checked, date: date.checked, dateText: fixed.value.trim().slice(0, 100), text: footer.checked ? text.value.trim().slice(0, 300) : '', skipTitle: skip.checked }; }, { all: true });
  }

  toggleSlideNumbers() { this.edit(() => { this.content.footer = { ...this.content.footer, number: !this.content.footer.number }; }, { all: true }); }

  // ---- menus -----------------------------------------------------------------------------------------------------------------------

  objectMenu(x, y, onElement) {
    const ro = this.readOnly;
    const sel = this.canvas.selected();
    const items = onElement ? [
      { label: t('Slides.Cut'), icon: 'slides-cut', key: 'Ctrl+X', run: () => this.cut(), disabled: ro },
      { label: t('Slides.Copy'), icon: 'copy', key: 'Ctrl+C', run: () => this.copy() },
      { label: t('Slides.Paste'), icon: 'slides-paste', key: 'Ctrl+V', run: () => this.paste(), disabled: ro },
      '-',
      sel.length === 1 && sel[0].type !== 'image' && sel[0].type !== 'line' ? { label: t('Slides.EditText'), icon: 'type', run: () => this.canvas.startText(sel[0], null), disabled: ro } : null,
      { label: t('Slides.Duplicate'), icon: 'slides-duplicate', key: 'Ctrl+D', run: () => this.duplicate(), disabled: ro },
      sel.length > 1 && !sel.every(e => e.group && e.group === sel[0].group) ? { label: t('Slides.Group'), icon: 'slides-group', run: () => this.group(), disabled: ro } : null,
      sel.some(e => e.group) ? { label: t('Slides.Ungroup'), icon: 'slides-ungroup', run: () => this.ungroup(), disabled: ro } : null,
      { label: t('Slides.BringFront'), icon: 'slides-bring-front', run: () => this.arrange('front'), disabled: ro },
      { label: t('Slides.SendBack'), icon: 'slides-send-back', run: () => this.arrange('back'), disabled: ro },
      '-',
      { label: sel.every(e => e.type === 'image') ? t('Slides.Pane.Picture') + ' …' : t('Slides.Pane.Shape') + ' …', icon: 'slides-format', run: () => this.showPane('format') },
      { label: t('Slides.Animation') + ' …', icon: 'slides-animation', run: () => this.showPane('animations') },
      '-',
      { label: t('Slides.Delete'), icon: 'trash', key: 'Del', run: () => this.deleteSelection(), danger: true, disabled: ro },
    ] : [
      { label: t('Slides.Paste'), icon: 'slides-paste', key: 'Ctrl+V', run: () => this.paste(), disabled: ro },
      '-',
      { label: t('Slides.Layout'), icon: 'slides-layout', sub: () => this.layoutItems(), disabled: ro },
      { label: t('Slides.ResetSlide'), icon: 'slides-reset', run: () => this.resetSlides(), disabled: ro },
      { label: t('Slides.FormatBackground') + ' …', icon: 'slides-background', run: () => this.showPane('format') },
      '-',
      { label: t('Slides.View.Gridlines'), icon: 'slides-gridlines', checked: !!this.view.grid, run: () => { this.view.grid = !this.view.grid; this.layoutChanged(); } },
      { label: t('Slides.View.Guides'), icon: 'slides-guides', checked: !!this.view.guides, run: () => { this.view.guides = !this.view.guides; this.layoutChanged(); } },
    ];
    contextMenu(x, y, items);
  }

  layoutItems() { return ['title', 'titleContent', 'section', 'twoContent', 'comparison', 'titleOnly', 'blank', 'pictureCaption'].map(id => ({ layoutId: id, label: t('Slides.Layout.' + id), icon: 'slides-layout', checked: this.slide?.layout === id, run: () => this.setLayout(id) })); }

  slideMenu(x, y, onSlide) {
    const ro = this.readOnly;
    const hidden = this.content.slides.filter(s => this.slideSel.has(s.id)).every(s => s.hidden);
    const items = onSlide ? [
      { label: t('Slides.Cut'), icon: 'slides-cut', run: () => { this.area = 'thumbs'; this.cut(); }, disabled: ro },
      { label: t('Slides.Copy'), icon: 'copy', run: () => { this.area = 'thumbs'; this.copy(); } },
      { label: t('Slides.Paste'), icon: 'slides-paste', run: () => { this.area = 'thumbs'; this.paste(); }, disabled: ro },
      '-',
      { label: t('Slides.NewSlide'), icon: 'slides-new-slide', run: () => this.newSlide(), disabled: ro },
      { label: t('Slides.DuplicateSlide'), icon: 'slides-duplicate', run: () => this.duplicateSlides(), disabled: ro },
      { label: t('Slides.DeleteSlide'), icon: 'trash', run: () => this.deleteSlides(), disabled: ro, danger: true },
      '-',
      { label: t('Slides.Layout'), icon: 'slides-layout', sub: () => this.layoutItems(), disabled: ro },
      { label: t('Slides.ResetSlide'), icon: 'slides-reset', run: () => this.resetSlides(), disabled: ro },
      { label: t('Slides.FormatBackground') + ' …', icon: 'slides-background', run: () => { this.select([]); this.showPane('format'); } },
      { label: t('Slides.HideSlide'), icon: 'slides-hide', checked: hidden, run: () => this.toggleHidden(), disabled: ro },
    ] : [
      { label: t('Slides.NewSlide'), icon: 'slides-new-slide', run: () => this.newSlide(), disabled: ro },
      { label: t('Slides.Paste'), icon: 'slides-paste', run: () => { this.area = 'thumbs'; this.paste(); }, disabled: ro },
    ];
    contextMenu(x, y, items);
  }

  textMenu(x, y) {
    const ro = this.readOnly;
    contextMenu(x, y, [
      { label: t('Slides.Cut'), icon: 'slides-cut', key: 'Ctrl+X', run: () => document.execCommand('cut'), disabled: ro },
      { label: t('Slides.Copy'), icon: 'copy', key: 'Ctrl+C', run: () => document.execCommand('copy') },
      { label: t('Slides.Paste'), icon: 'slides-paste', key: 'Ctrl+V', run: () => this.paste(), disabled: ro },
      '-',
      { label: t('Slides.Bold'), icon: 'slides-bold', run: () => this.toggleRun('b'), disabled: ro },
      { label: t('Slides.Italic'), icon: 'slides-italic', run: () => this.toggleRun('i'), disabled: ro },
      { label: t('Slides.Bullets'), icon: 'slides-bullets', run: () => this.toggleBullets(), disabled: ro },
    ]);
  }

  layoutPicker(anchor, pick) { layoutGallery(anchor, this.content, this.slide?.layout, pick); }

  // ---- outputs ---------------------------------------------------------------------------------------------------------------------

  async thumbnail() {
    const first = this.content.slides[0];
    if (!first) return null;
    return slidePng(this.content, first, 1, 480);
  }

  async print() {
    if (this.text.active) this.text.stop();
    const kind = await askPrint(this.ctx.dialog);
    if (kind) printDeck(this.content, kind);
  }

  async exportPng() {
    if (!this.slide) return;
    try {
      const png = await slidePng(this.content, this.slide, this.index + 1, 1920);
      download(`${this.ctx.doc.name || 'slide'} ${this.index + 1}.png`, png, 'image/png');
    } catch { this.ctx.toast(t('Slides.Png.Failed'), 'danger'); }
  }

  downloads() {
    return [
      { ext: 'PNG', label: t('Slides.Download.Png'), run: () => this.exportPng() },
      { ext: 'PDF', label: t('Slides.Download.Pdf'), run: () => { document.getElementById('mo-backstage').hidden = true; document.body.classList.remove('mo-backstage-open'); this.print(); } },
    ];
  }

  info() {
    const slides = this.content.slides;
    const rows = [
      [t('Slides.Info.Slides'), String(slides.length)],
      [t('Slides.Info.Hidden'), String(slides.filter(s => s.hidden).length)],
      [t('Slides.Info.Notes'), String(slides.filter(s => s.notes).length)],
      [t('Slides.Info.Size'), this.content.size.w / this.content.size.h > 1.5 ? '16:9' : '4:3'],
      [t('Slides.Info.Design'), themeById(this.content.theme.id) ? t('Slides.Theme.' + this.content.theme.id) : this.content.theme.name || ''],
    ];
    return el('dl', { class: 'mat-kv mo-slides-info' }, ...rows.flatMap(([k, v]) => [el('dt', { text: k }), el('dd', { text: v })]));
  }
}

const app = new SlidesApp();

startEditor({
  mount: ctx => app.mount(ctx),
  getContent: () => app.getContent(),
  setContent: (content, info) => app.setContent(content, info),
  thumbnail: () => app.thumbnail(),
  print: () => app.print(),
  info: () => app.info(),
  downloads: () => app.downloads(),
});

/** The editor of the page (for tests and the console). */
export { app };
