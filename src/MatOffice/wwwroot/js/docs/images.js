/* MatOffice Docs – pictures: insert from a file, the clipboard or by dropping (kept inside the document as data: addresses, made smaller in
   the browser), select one with a click (a frame with handles: drag a corner to resize, the proportions stay), wrap it (in line with the text,
   left, right, centred), size in centimetres, the original size, alternative text, delete. */
import { pickFile, shrinkImage, readAsDataUrl } from '../editor/util.js';
import { numStr, safeImage } from './model.js';

const PX_PER_CM = 96 / 2.54;

export class Images {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    this.selected = null;
    this.frame = document.createElement('div');
    this.frame.className = 'mo-docs-imgsel';
    this.frame.setAttribute('data-mo-deco', '');
    this.frame.hidden = true;
    for (const h of ['nw', 'ne', 'sw', 'se']) {
      const handle = document.createElement('span');
      handle.className = 'mo-docs-imgsel__h';
      handle.dataset.h = h;
      handle.addEventListener('pointerdown', e => this.startResize(e, h));
      this.frame.append(handle);
    }
    app.view.stack.append(this.frame);
    const onClick = e => {
      const img = e.target.closest?.('img');
      const host = this.ed.hostOf(e.target);
      if (img && host) { this.select(img); return; }
      if (this.selected && !this.frame.contains(e.target)) this.clear();
    };
    app.view.scroller.addEventListener('click', onClick);
    this.ed.addEventListener('selection', () => {
      if (!this.selected) return;
      const sel = document.getSelection();
      const r = sel.rangeCount ? sel.getRangeAt(0) : null;
      if (!this.selected.isConnected || !r || !(r.intersectsNode(this.selected))) this.clear();
      else this.position();
    });
    new ResizeObserver(() => this.position()).observe(app.view.stack);
  }

  // ---- inserting ----------------------------------------------------------------------------------------------------------------------

  async pick() {
    if (this.app.ctx.readOnly) return;
    this.ed.focus();
    const mark = this.ed.bookmark();
    const files = await pickFile('image/png,image/jpeg,image/gif,image/webp,image/bmp', true);
    if (!files?.length) return;
    if (mark) this.ed.restoreBookmark(mark, false);
    await this.insertFiles(files);
  }

  async insertFiles(files) {
    const runs = [];
    for (const file of files.slice(0, 20)) {
      try {
        const src = await toDataUrl(file);
        if (!src) { this.app.ctx.toast(this.app.ctx.t('Docs.Picture.Unsupported'), 'warning'); continue; }
        const size = await naturalSize(src);
        const max = this.maxWidth();
        const scale = Math.min(1, max / size.width);
        runs.push({ obj: 'img', src, alt: '', width: Math.round(size.width * scale), height: Math.round(size.height * scale), f: {} });
      } catch {
        this.app.ctx.toast(this.app.ctx.t('Docs.Picture.Failed'), 'danger');
      }
    }
    if (!runs.length) return;
    this.ed.focus();
    this.ed.insertBlocks([{ type: 'p', css: {}, runs }]);
  }

  /** The widest a picture may be where the caret is: the text width, a cell, a column. */
  maxWidth() {
    const g = this.app.view.geometry;
    let width = g ? (g.textWidth - (g.columns - 1) * g.columnGap) / g.columns : 600;
    const sel = this.ed.selection();
    const cell = sel?.start.block.closest('td,th');
    if (cell) width = Math.max(40, cell.getBoundingClientRect().width / this.app.view.zoom - 16);
    return Math.max(40, Math.floor(width));
  }

  // ---- the selected picture ---------------------------------------------------------------------------------------------------------

  select(img) {
    if (this.selected && this.selected !== img) this.selected.classList.remove('is-selected');
    this.selected = img;
    img.classList.add('is-selected');
    const r = document.createRange();
    r.selectNode(img);
    const sel = document.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    this.position();
    this.ed.dispatchEvent(new CustomEvent('selection'));
  }

  clear() {
    if (!this.selected) return;
    this.selected.classList.remove('is-selected');
    this.selected = null;
    this.frame.hidden = true;
    this.ed.dispatchEvent(new CustomEvent('selection'));
  }

  position() {
    const img = this.selected;
    if (!img || !img.isConnected) { this.frame.hidden = true; return; }
    const r = this.app.view.local(img.getBoundingClientRect());
    Object.assign(this.frame.style, { left: numStr(r.left) + 'px', top: numStr(r.top) + 'px', width: numStr(r.right - r.left) + 'px', height: numStr(r.bottom - r.top) + 'px' });
    const editable = !this.app.ctx.readOnly && !!this.ed.hostOf(img);
    this.frame.hidden = false;
    this.frame.querySelectorAll('.mo-docs-imgsel__h').forEach(h => {
      h.hidden = !editable;
      const [v, hz] = h.dataset.h;
      h.style.top = v === 'n' ? '0' : '100%';
      h.style.left = hz === 'w' ? '0' : '100%';
    });
  }

  align() { return this.selected?.getAttribute('data-align') || ''; }

  setAlign(value) {
    const img = this.selected;
    if (!img) return;
    this.ed.checkpoint('picture');
    if (value) img.setAttribute('data-align', value); else img.removeAttribute('data-align');
    requestAnimationFrame(() => this.position());
  }

  sizeCm(dim) {
    const img = this.selected;
    if (!img) return '';
    const px = parseFloat(img.style[dim]) || (dim === 'width' ? img.getBoundingClientRect().width / this.app.view.zoom : img.getBoundingClientRect().height / this.app.view.zoom);
    return numStr(px / PX_PER_CM);
  }

  setSizeCm(dim, cm) {
    const img = this.selected;
    if (!img || !(cm > 0)) return;
    const w = parseFloat(img.style.width) || img.naturalWidth, h = parseFloat(img.style.height) || img.naturalHeight;
    const ratio = w && h ? w / h : 1;
    const px = cm * PX_PER_CM;
    this.ed.checkpoint('picture');
    if (dim === 'width') this.setSize(px, px / ratio);
    else this.setSize(px * ratio, px);
  }

  setSize(w, h) {
    const img = this.selected;
    img.style.width = numStr(Math.max(8, w)) + 'px';
    img.style.height = numStr(Math.max(8, h)) + 'px';
    requestAnimationFrame(() => this.position());
  }

  resetSize() {
    const img = this.selected;
    if (!img) return;
    this.ed.checkpoint('picture');
    const max = this.maxWidth();
    const scale = Math.min(1, max / (img.naturalWidth || max));
    this.setSize((img.naturalWidth || 100) * scale, (img.naturalHeight || 100) * scale);
  }

  async altText() {
    const img = this.selected;
    if (!img) return;
    const { t, el } = this.app.ctx;
    const area = el('textarea', { class: 'mat-input', rows: 4, maxlength: 1000, 'aria-label': t('Docs.Picture.Alt') });
    area.value = img.getAttribute('alt') || '';
    const answer = await this.app.ctx.dialog({
      title: t('Docs.Picture.Alt'), icon: 'image', body: el('div', { class: 'mo-docs-form' }, el('p', { class: 'mo-muted', text: t('Docs.Picture.AltText') }), area),
      buttons: [{ label: t('Docs.Ok'), value: 'ok', kind: 'primary' }, { label: t('Docs.Cancel'), value: 'cancel' }], onOpen: () => area.focus(),
    });
    if (answer !== 'ok' || !img.isConnected) return;
    this.ed.checkpoint('picture');
    img.setAttribute('alt', area.value.replace(/\s+/g, ' ').trim());
  }

  remove() {
    const img = this.selected;
    if (!img) return;
    this.ed.checkpoint('picture');
    const block = img.closest('p,h1,h2,h3,h4,h5,h6,blockquote,pre');
    this.clear();
    img.remove();
    if (block && !block.textContent && !block.querySelector('img,br')) block.append(document.createElement('br'));
  }

  // ---- resizing with the handles ---------------------------------------------------------------------------------------------------------

  startResize(e, corner) {
    const img = this.selected;
    if (!img || this.app.ctx.readOnly) return;
    e.preventDefault();
    e.stopPropagation();
    const handle = e.currentTarget;
    handle.setPointerCapture(e.pointerId);
    const zoom = this.app.view.zoom;
    const rect = img.getBoundingClientRect();
    const start = { x: e.clientX, w: rect.width / zoom, h: rect.height / zoom };
    const ratio = start.w / (start.h || 1);
    const max = this.maxWidth();
    let moved = false;
    const move = ev => {
      const dx = (ev.clientX - start.x) / zoom;
      let w = corner.endsWith('e') ? start.w + dx : start.w - dx;
      w = Math.max(16, Math.min(max, w));
      if (!moved) { this.ed.checkpoint('picture'); moved = true; }
      this.setSize(w, w / ratio);
    };
    const up = () => {
      handle.removeEventListener('pointermove', move);
      handle.removeEventListener('pointerup', up);
      handle.removeEventListener('pointercancel', up);
      this.position();
    };
    handle.addEventListener('pointermove', move);
    handle.addEventListener('pointerup', up);
    handle.addEventListener('pointercancel', up);
  }
}

/** A picture file as a data: address of an allowed type (png, jpeg, gif, webp; anything else is drawn into a PNG/JPEG). */
async function toDataUrl(file) {
  let src = /^image\/(png|jpeg|gif|webp)$/.test(file.type) ? await shrinkImage(file, 1600) : null;
  if (!src || !safeImage(src)) {
    const raw = await readAsDataUrl(file);
    const img = await load(raw);
    const scale = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    src = canvas.toDataURL(file.type === 'image/jpeg' ? 'image/jpeg' : 'image/png', 0.88);
  }
  return safeImage(src);
}

function load(src) {
  return new Promise((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = src; });
}

async function naturalSize(src) {
  const img = await load(src);
  return { width: img.naturalWidth || 100, height: img.naturalHeight || 100 };
}
