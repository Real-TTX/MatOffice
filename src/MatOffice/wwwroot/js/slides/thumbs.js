/* MatOffice Slides – the slides as small pictures: the pane on the left of the normal view (a strip on a phone) and the slide sorter.
   Click selects (Ctrl adds, Shift a range), drag reorders, the context menu has new / duplicate / delete / hide / layout / reset,
   the keyboard moves (arrows, Home, End), deletes (Delete), adds (Enter) and duplicates (Ctrl+D). */
import { t, el, icon, mod } from '../editor/util.js';
import { renderSlide, fitText } from './render.js';

export class Thumbs {
  /** mode: 'pane' | 'sorter' */
  constructor(app, mode = 'pane') {
    this.app = app;
    this.mode = mode;
    this.cache = new Map();
    this.list = el('div', { class: 'mo-slides-thumbs mo-slides-thumbs--' + mode, role: 'listbox', tabindex: '0', 'aria-multiselectable': 'true', 'aria-label': t('Slides.Thumbnails') });
    this.list.addEventListener('pointerdown', e => this.down(e));
    this.list.addEventListener('keydown', e => this.key(e));
    this.list.addEventListener('contextmenu', e => this.context(e));
    this.list.addEventListener('focus', () => { this.app.area = 'thumbs'; });
    this.list.addEventListener('dblclick', e => { const item = e.target.closest('.mo-slides-thumb'); if (item && this.mode === 'sorter') this.app.setView('normal'); });
  }

  get width() { return this.mode === 'sorter' ? 236 : this.app.phone ? 104 : 176; }

  /** Marks slides to be drawn again (all when no id). */
  invalidate(id) { if (id) this.cache.delete(id); else this.cache.clear(); }

  render() {
    const { content } = this.app;
    const keep = new Set(content.slides.map(s => s.id));
    for (const id of [...this.cache.keys()]) if (!keep.has(id)) this.cache.delete(id);
    const scale = this.width / content.size.w;
    const height = Math.round(content.size.h * scale);
    const items = content.slides.map((slide, i) => {
      let entry = this.cache.get(slide.id);
      if (!entry || entry.width !== this.width) {
        const node = renderSlide(content, slide, { mode: 'view', number: i + 1 });
        node.style.transform = `scale(${scale})`;
        node.style.transformOrigin = '0 0';
        entry = { node, width: this.width, fitted: false, number: i + 1 };
        this.cache.set(slide.id, entry);
      }
      const current = i === this.app.index;
      const selected = this.app.slideSel.has(slide.id) || current;
      const frame = el('div', { class: 'mo-slides-thumb__frame', style: { width: this.width + 'px', height: height + 'px' } }, entry.node);
      const marks = el('div', { class: 'mo-slides-thumb__marks' },
        slide.transition && slide.transition.type !== 'none' ? icon('slides-transition', 'mo-slides-thumb__mark') : null,
        slide.animations?.length ? icon('slides-animation', 'mo-slides-thumb__mark') : null);
      return el('div', {
        class: 'mo-slides-thumb' + (current ? ' is-current' : '') + (selected ? ' is-selected' : '') + (slide.hidden ? ' is-hidden' : ''),
        role: 'option', 'aria-selected': selected ? 'true' : 'false', 'aria-label': t('Slides.SlideN', i + 1) + (slide.hidden ? ' · ' + t('Slides.Hidden') : ''), 'data-id': slide.id,
      }, el('div', { class: 'mo-slides-thumb__num' }, el('span', { text: String(i + 1) }), marks), frame);
    });
    this.list.replaceChildren(...items);
    if (!content.slides.length) this.list.append(el('button', { type: 'button', class: 'mo-slides-thumbs__empty', onclick: () => this.app.newSlide() }, icon('plus'), el('span', { text: t('Slides.AddFirst') })));
    for (const entry of this.cache.values()) if (!entry.fitted && entry.node.isConnected) { fitText(entry.node); entry.fitted = true; }
    this.scrollToCurrent();
  }

  scrollToCurrent() {
    const cur = this.list.querySelector('.mo-slides-thumb.is-current');
    if (!cur) return;
    const l = this.list.getBoundingClientRect(), c = cur.getBoundingClientRect();
    if (c.top < l.top || c.bottom > l.bottom || c.left < l.left || c.right > l.right) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  indexOf(node) { return this.app.content.slides.findIndex(s => s.id === node.dataset.id); }

  down(e) {
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    this.app.area = 'thumbs';
    const item = e.target.closest('.mo-slides-thumb');
    if (!item) return;
    const app = this.app;
    const i = this.indexOf(item);
    const id = item.dataset.id;
    if (e.shiftKey) {
      const a = Math.min(app.index, i), b = Math.max(app.index, i);
      app.slideSel = new Set(app.content.slides.slice(a, b + 1).map(s => s.id));
      app.goTo(i, { keepSlideSel: true });
      return;
    }
    if (mod(e)) {
      const set = new Set(app.slideSel.size ? app.slideSel : [app.slide?.id]);
      if (set.has(id) && set.size > 1) set.delete(id); else set.add(id);
      app.slideSel = set;
      app.goTo(i, { keepSlideSel: true });
      return;
    }
    if (!app.slideSel.has(id)) app.slideSel = new Set([id]);
    app.goTo(i, { keepSlideSel: true });
    if (app.readOnly) return;
    this.dragReorder(e);
  }

  dragReorder(e) {
    const app = this.app;
    const start = { x: e.clientX, y: e.clientY };
    let dragging = false, target = -1, marker = null;
    const grid = this.mode === 'sorter' || this.app.phone;
    this.list.setPointerCapture?.(e.pointerId);
    const move = ev => {
      if (!dragging && Math.hypot(ev.clientX - start.x, ev.clientY - start.y) < 6) return;
      if (!dragging) { dragging = true; this.list.classList.add('is-dragging'); marker = el('div', { class: 'mo-slides-thumbs__marker' + (grid ? ' is-vertical' : '') }); }
      const items = [...this.list.querySelectorAll('.mo-slides-thumb')];
      target = items.length;
      for (let k = 0; k < items.length; k++) {
        const r = items[k].getBoundingClientRect();
        const before = grid ? (ev.clientY < r.bottom && ev.clientX < r.left + r.width / 2) || ev.clientY < r.top : ev.clientY < r.top + r.height / 2;
        if (before) { target = k; break; }
      }
      const ref = items[target];
      if (ref) ref.before(marker); else this.list.append(marker);
      const l = this.list.getBoundingClientRect();
      if (ev.clientY < l.top + 30) this.list.scrollTop -= 12; else if (ev.clientY > l.bottom - 30) this.list.scrollTop += 12;
    };
    const up = () => {
      this.list.removeEventListener('pointermove', move);
      this.list.removeEventListener('pointerup', up);
      this.list.removeEventListener('pointercancel', up);
      this.list.classList.remove('is-dragging');
      marker?.remove();
      if (dragging && target >= 0) app.moveSlides([...app.selectedSlideIds()], target);
    };
    this.list.addEventListener('pointermove', move);
    this.list.addEventListener('pointerup', up);
    this.list.addEventListener('pointercancel', up);
  }

  key(e) {
    const app = this.app;
    const n = app.content.slides.length;
    const back = ['ArrowUp', 'ArrowLeft'].includes(e.key), fwd = ['ArrowDown', 'ArrowRight'].includes(e.key);
    if (back || fwd || e.key === 'Home' || e.key === 'End') {
      e.preventDefault();
      const i = e.key === 'Home' ? 0 : e.key === 'End' ? n - 1 : app.index + (fwd ? 1 : -1);
      if (i < 0 || i >= n) return;
      if (e.shiftKey) { const set = new Set(app.slideSel.size ? app.slideSel : [app.slide.id]); set.add(app.content.slides[i].id); app.slideSel = set; app.goTo(i, { keepSlideSel: true }); }
      else app.goTo(i);
      return;
    }
    if (app.readOnly) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); app.deleteSlides(); return; }
    if (e.key === 'Enter') { e.preventDefault(); app.newSlide(); return; }
    if (mod(e) && (e.key === 'd' || e.key === 'D')) { e.preventDefault(); e.stopPropagation(); app.duplicateSlides(); }
  }

  context(e) {
    e.preventDefault();
    const item = e.target.closest('.mo-slides-thumb');
    if (item) {
      const id = item.dataset.id;
      if (!this.app.slideSel.has(id)) { this.app.slideSel = new Set([id]); this.app.goTo(this.indexOf(item), { keepSlideSel: true }); }
    }
    this.app.slideMenu(e.clientX, e.clientY, !!item);
  }
}
