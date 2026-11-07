/* MatOffice Slides – the slide show: full screen (or the reading view inside the window), transitions between slides and entrance animations
   on click / with / after the previous one, the keys of the office suites (→ Space PgDn next, ← PgUp previous, Home, End, a number + Enter,
   B/W black or white screen, L laser pointer, Esc end), click to go on, hidden slides skipped. The presenter view is a second window of the
   same page (window.open, driven from here: no server, no other site): the current and the next slide, the notes, a timer and the counter. */
import { t, el, icon } from '../editor/util.js';
import { renderSlide, fitText, ensureStyles } from './render.js';

const KEYS_NEXT = ['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'];
const KEYS_PREV = ['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'];

/** The click steps of a slide: groups of blocks ("after previous" starts a block) of effects ("with previous" joins one). */
export function stepsOf(slide) {
  const ids = new Set(slide.elements.map(e => e.id));
  const groups = [];
  for (const a of slide.animations || []) {
    if (!ids.has(a.el)) continue;
    if (a.trigger === 'click' || !groups.length) groups.push({ auto: !groups.length && a.trigger !== 'click', blocks: [[a]] });
    else if (a.trigger === 'after') groups[groups.length - 1].blocks.push([a]);
    else groups[groups.length - 1].blocks[groups[groups.length - 1].blocks.length - 1].push(a);
  }
  return groups;
}

export class Show {
  constructor(app) {
    this.app = app;
    this.running = false;
    this.onKey = e => this.key(e);
  }

  /** Starts at a slide index; options: { reading: true } shows it inside the work area. */
  start(from = 0, { reading = false, presenter = false } = {}) {
    if (this.running) this.end();
    const { content } = this.app;
    if (!content.slides.length) return;
    this.reading = reading;
    this.order = content.slides.map((s, i) => i).filter(i => !content.slides[i].hidden || i === from);
    if (!this.order.length) this.order = [from];
    this.pos = Math.max(0, this.order.indexOf(from));
    if (this.pos < 0) this.pos = 0;
    this.running = true;
    this.black = null;
    this.digits = '';
    this.started = Date.now();
    this.paused = 0;
    this.pauseAt = null;
    this.root = el('div', { class: 'mo-slides-show' + (reading ? ' is-reading' : ''), tabindex: '0', role: 'application', 'aria-roledescription': t('Slides.Show'), 'aria-label': t('Slides.Show') });
    this.screen = el('div', { class: 'mo-slides-show__screen' });
    this.cover = el('div', { class: 'mo-slides-show__cover', hidden: true });
    this.laser = el('div', { class: 'mo-slides-show__laser', hidden: true });
    this.live = el('div', { class: 'mat-sr-only', 'aria-live': 'polite' });
    this.bar = this.toolbar();
    this.root.append(this.screen, this.cover, this.laser, this.bar, this.live);
    (reading ? this.app.work : document.body).append(this.root);
    this.root.addEventListener('click', e => { if (e.target.closest('.mo-slides-show__bar')) return; if (this.laserOn) return; this.next(); });
    this.root.addEventListener('contextmenu', e => { e.preventDefault(); this.prev(); });
    this.root.addEventListener('wheel', e => { if (Math.abs(e.deltaY) < 10) return; e.preventDefault(); if (e.deltaY > 0) this.next(); else this.prev(); }, { passive: false });
    this.root.addEventListener('pointermove', e => this.pointer(e));
    document.addEventListener('keydown', this.onKey, true);
    this.resize = () => this.layout();
    window.addEventListener('resize', this.resize);
    if (!reading && document.documentElement.requestFullscreen && !document.fullscreenElement) {
      document.documentElement.requestFullscreen({ navigationUI: 'hide' }).then(() => { this.fullscreen = true; }).catch(() => { this.fullscreen = false; });
      this.onFs = () => { if (!document.fullscreenElement && this.fullscreen && this.running) this.end(); };
      document.addEventListener('fullscreenchange', this.onFs);
    }
    this.root.focus({ preventScroll: true });
    this.enter(false);
    if (presenter) this.openPresenter();
  }

  get slideIndex() { return this.order[this.pos]; }

  get slide() { return this.app.content.slides[this.slideIndex]; }

  toolbar() {
    const b = (name, label, run) => el('button', { type: 'button', class: 'mo-slides-show__btn', 'aria-label': label, 'data-tip': label, onclick: e => { e.stopPropagation(); run(); } }, icon(name));
    this.counter = el('span', { class: 'mo-slides-show__count' });
    return el('div', { class: 'mo-slides-show__bar' },
      b('chevron-left', t('Slides.Show.Previous'), () => this.prev()),
      this.counter,
      b('chevron-right', t('Slides.Show.Next'), () => this.next()),
      b('slides-laser', t('Slides.Show.Laser'), () => this.toggleLaser()),
      b('slides-black', t('Slides.Show.Black'), () => this.blank('black')),
      this.reading ? null : b('slides-presenter', t('Slides.Show.Presenter'), () => this.openPresenter()),
      b('x', t('Slides.Show.End'), () => this.end()));
  }

  layout() {
    if (!this.running) return;
    const { w, h } = this.app.content.size;
    const r = this.screen.getBoundingClientRect();
    const scale = Math.min(r.width / w, r.height / h);
    this.scale = scale;
    for (const layer of this.screen.children) this.place(layer, scale);
    this.updatePresenter();
  }

  place(layer, scale = this.scale) {
    const { w, h } = this.app.content.size;
    const r = this.screen.getBoundingClientRect();
    const node = layer.firstChild;
    Object.assign(layer.style, { width: w * scale + 'px', height: h * scale + 'px', left: (r.width - w * scale) / 2 + 'px', top: (r.height - h * scale) / 2 + 'px' });
    if (node) { node.style.transform = `scale(${scale})`; node.style.transformOrigin = '0 0'; }
  }

  /** Draws the current slide with the effects up to the current step done. */
  layer(step) {
    const slide = this.slide;
    const groups = stepsOf(slide);
    const animated = new Set(groups.flatMap(g => g.blocks.flat().map(a => a.el)));
    for (const g of groups.slice(0, step)) for (const a of g.blocks.flat()) animated.delete(a.el);
    const node = renderSlide(this.app.content, slide, { mode: 'view', number: this.slideIndex + 1, hide: animated });
    const layer = el('div', { class: 'mo-slides-show__layer' }, node);
    return layer;
  }

  /** Shows the current slide; with a transition when going forward. */
  enter(animate, atEnd = false) {
    clearTimeout(this.autoTimer);
    this.finishRunning();
    const slide = this.slide;
    const groups = stepsOf(slide);
    this.groups = groups;
    this.step = atEnd ? groups.length : 0;
    const layer = this.layer(this.step);
    const old = [...this.screen.children];
    this.screen.append(layer);
    this.place(layer);
    fitText(layer.firstChild);
    const tr = slide.transition;
    if (animate && tr && tr.type !== 'none' && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
      this.transition(layer, old, tr).finally(() => old.forEach(n => n.remove()));
    } else old.forEach(n => n.remove());
    this.live.textContent = t('Slides.SlideNofM', this.slideIndex + 1, this.app.content.slides.length);
    this.counter.textContent = `${this.slideIndex + 1} / ${this.app.content.slides.length}`;
    if (!atEnd && groups[0]?.auto) { this.step = 1; this.play(groups[0], layer); }
    this.scheduleAuto();
    this.updatePresenter();
    this.app.showMoved?.(this.slideIndex);
  }

  scheduleAuto() {
    clearTimeout(this.autoTimer);
    const auto = this.slide.transition?.auto;
    if (auto === undefined || auto === null || this.reading) return;
    this.autoTimer = setTimeout(() => { if (this.running && !this.black) this.next(); }, Math.max(0, auto) * 1000 + this.pendingTime());
  }

  pendingTime() { return this.runningUntil ? Math.max(0, this.runningUntil - Date.now()) : 0; }

  transition(layer, old, tr) {
    const ms = Math.max(100, (tr.duration || 0.7) * 1000);
    const r = this.screen.getBoundingClientRect();
    const dir = tr.dir || (tr.type === 'push' ? 'down' : 'right');
    const off = { right: [r.width, 0], left: [-r.width, 0], down: [0, r.height], up: [0, -r.height] }[dir] || [r.width, 0];
    const opts = { duration: ms, easing: 'cubic-bezier(.4, 0, .2, 1)' };
    const clip = { right: 'inset(0 0 0 100%)', left: 'inset(0 100% 0 0)', down: 'inset(100% 0 0 0)', up: 'inset(0 0 100% 0)' }[dir];
    let a;
    switch (tr.type) {
      case 'fade': a = layer.animate([{ opacity: 0 }, { opacity: 1 }], opts); break;
      case 'push':
        a = layer.animate([{ translate: `${off[0]}px ${off[1]}px` }, { translate: '0 0' }], opts);
        old.forEach(o => o.animate([{ translate: '0 0' }, { translate: `${-off[0]}px ${-off[1]}px` }], { ...opts, fill: 'forwards' }));
        break;
      case 'cover': a = layer.animate([{ translate: `${off[0]}px ${off[1]}px` }, { translate: '0 0' }], opts); break;
      case 'wipe': a = layer.animate([{ clipPath: clip }, { clipPath: 'inset(0 0 0 0)' }], opts); break;
      case 'split': a = layer.animate([{ clipPath: dir === 'up' || dir === 'down' ? 'inset(50% 0 50% 0)' : 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }], opts); break;
      case 'zoom': a = layer.animate([{ opacity: 0, scale: 0.3 }, { opacity: 1, scale: 1 }], opts); break;
      default: return Promise.resolve();
    }
    this.transitionAnim = a;
    return a.finished.catch(() => {}).then(() => { this.transitionAnim = null; });
  }

  /** Plays a click group: its blocks one after the other, the effects of a block together. */
  play(group, layer = this.screen.lastChild) {
    let start = 0;
    this.running_ = this.running_ || [];
    for (const block of group.blocks) {
      let longest = 0;
      for (const a of block) {
        const node = layer.querySelector(`.mo-sr-el[data-id="${CSS.escape(a.el)}"]`);
        if (!node) continue;
        const delay = start + (a.delay || 0) * 1000;
        const duration = a.effect === 'appear' ? 1 : Math.max(50, (a.duration || 0.5) * 1000);
        longest = Math.max(longest, (a.delay || 0) * 1000 + duration);
        const frames = this.effect(a, node);
        setTimeout(() => { node.style.visibility = 'visible'; }, delay);
        if (frames && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
          this.running_.push(node.animate(frames, { duration, delay, easing: a.effect === 'fly' ? 'cubic-bezier(.2, .8, .2, 1)' : 'ease-out', fill: 'backwards' }));
        }
      }
      start += longest;
    }
    this.runningUntil = Date.now() + start;
    this.playTimer = setTimeout(() => { this.runningUntil = 0; }, start);
  }

  effect(a, node) {
    const { w, h } = this.app.content.size;
    const box = { x: parseFloat(node.style.left), y: parseFloat(node.style.top), w: parseFloat(node.style.width), h: parseFloat(node.style.height) };
    switch (a.effect) {
      case 'fade': return [{ opacity: 0 }, { opacity: 1 }];
      case 'zoom': return [{ opacity: 0, scale: 0.2 }, { opacity: 1, scale: 1 }];
      case 'wipe': {
        const from = { bottom: 'inset(100% 0 0 0)', top: 'inset(0 0 100% 0)', left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)' }[a.dir || 'bottom'];
        return [{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }];
      }
      case 'fly': {
        const d = a.dir || 'bottom';
        const from = d === 'bottom' ? `0 ${h - box.y + 10}px` : d === 'top' ? `0 ${-(box.y + box.h + 10)}px` : d === 'left' ? `${-(box.x + box.w + 10)}px 0` : `${w - box.x + 10}px 0`;
        return [{ translate: from }, { translate: '0 0' }];
      }
      default: return null;
    }
  }

  finishRunning() {
    clearTimeout(this.playTimer);
    this.runningUntil = 0;
    for (const a of this.running_ || []) { try { a.finish(); } catch { /* gone */ } }
    this.running_ = [];
    this.screen?.querySelectorAll('.mo-sr-el').forEach(n => { if (n.style.visibility === 'hidden' && n.dataset.shown === '1') n.style.visibility = 'visible'; });
    if (this.transitionAnim) { try { this.transitionAnim.finish(); } catch { /* gone */ } }
  }

  next() {
    if (!this.running) return;
    if (this.black) { this.blank(null); return; }
    if (this.ended) { this.end(); return; }
    if (this.runningUntil && Date.now() < this.runningUntil) { this.finishRunning(); this.showStep(this.step); return; }
    if (this.step < this.groups.length) {
      const g = this.groups[this.step];
      this.step++;
      this.play(g);
      this.scheduleAuto();
      this.updatePresenter();
      return;
    }
    if (this.pos < this.order.length - 1) { this.pos++; this.enter(true); return; }
    this.endScreen();
  }

  /** Draws the current slide with all effects up to a step done (after a skip or going back). */
  showStep(step) {
    this.step = step;
    const layer = this.layer(step);
    this.screen.replaceChildren(layer);
    this.place(layer);
    fitText(layer.firstChild);
    this.updatePresenter();
  }

  prev() {
    if (!this.running) return;
    if (this.black) { this.blank(null); return; }
    if (this.ended) { this.ended = false; this.cover.hidden = true; this.cover.className = 'mo-slides-show__cover'; return; }
    this.finishRunning();
    const auto = this.groups[0]?.auto ? 1 : 0;
    if (this.step > auto) { this.showStep(this.step - 1); return; }
    if (this.pos > 0) { this.pos--; this.enter(false, true); }
  }

  goTo(slideIndex) {
    const i = this.order.indexOf(slideIndex);
    if (i >= 0) this.pos = i;
    else { this.order.push(slideIndex); this.order.sort((a, b) => a - b); this.pos = this.order.indexOf(slideIndex); }
    this.ended = false;
    this.cover.hidden = true;
    this.enter(false);
  }

  endScreen() {
    this.ended = true;
    this.cover.hidden = false;
    this.cover.className = 'mo-slides-show__cover is-end';
    this.cover.replaceChildren(el('span', { text: t('Slides.Show.EndScreen') }));
    this.updatePresenter();
  }

  /** A black or white screen (B, W); any key or click returns. */
  blank(color) {
    this.black = color && this.black !== color ? color : null;
    this.cover.hidden = !this.black;
    this.cover.className = 'mo-slides-show__cover' + (this.black ? ' is-' + this.black : '');
    this.cover.replaceChildren();
    this.updatePresenter();
  }

  toggleLaser() {
    this.laserOn = !this.laserOn;
    this.laser.hidden = !this.laserOn;
    this.root.classList.toggle('has-laser', this.laserOn);
  }

  pointer(e) {
    if (this.laserOn) { this.laser.style.left = e.clientX - (this.reading ? this.root.getBoundingClientRect().left : 0) + 'px'; this.laser.style.top = e.clientY - (this.reading ? this.root.getBoundingClientRect().top : 0) + 'px'; }
    this.root.classList.add('is-awake');
    clearTimeout(this.sleep);
    this.sleep = setTimeout(() => this.root?.classList.remove('is-awake'), 2500);
  }

  key(e) {
    if (!this.running) return;
    if (e.target.closest?.('.mo-dialog')) return;
    const k = e.key;
    if (this.black && !['b', 'B', 'w', 'W', '.', ','].includes(k)) { e.preventDefault(); e.stopPropagation(); this.blank(null); return; }
    if (/^[0-9]$/.test(k)) { this.digits += k; e.preventDefault(); e.stopPropagation(); return; }
    if (k === 'Enter' && this.digits) {
      e.preventDefault(); e.stopPropagation();
      const n = Number(this.digits) - 1;
      this.digits = '';
      if (n >= 0 && n < this.app.content.slides.length) this.goTo(n);
      return;
    }
    this.digits = '';
    let handled = true;
    if (KEYS_NEXT.includes(k)) this.next();
    else if (KEYS_PREV.includes(k)) this.prev();
    else if (k === 'Home') { this.pos = 0; this.ended = false; this.cover.hidden = true; this.enter(false); }
    else if (k === 'End') { this.pos = this.order.length - 1; this.ended = false; this.cover.hidden = true; this.enter(false, true); }
    else if (k === 'b' || k === 'B' || k === '.') this.blank('black');
    else if (k === 'w' || k === 'W' || k === ',') this.blank('white');
    else if (k === 'l' || k === 'L') this.toggleLaser();
    else if (k === 'Escape' || k === '-') this.end();
    else if (k === 'F5' || (e.ctrlKey && (k === 'p' || k === 'P'))) handled = true;
    else handled = false;
    if (handled) { e.preventDefault(); e.stopPropagation(); }
  }

  end() {
    if (!this.running) return;
    this.running = false;
    clearTimeout(this.autoTimer);
    clearTimeout(this.playTimer);
    clearTimeout(this.sleep);
    clearInterval(this.clock);
    document.removeEventListener('keydown', this.onKey, true);
    window.removeEventListener('resize', this.resize);
    if (this.onFs) document.removeEventListener('fullscreenchange', this.onFs);
    if (document.fullscreenElement && this.fullscreen) document.exitFullscreen().catch(() => {});
    this.root.remove();
    if (this.pv && !this.pv.closed) this.pv.close();
    this.pv = null;
    const last = this.slideIndex;
    this.app.showEnded(last);
  }

  // ---- presenter view ----------------------------------------------------------------------------------------------------------------

  openPresenter() {
    if (this.pv && !this.pv.closed) { this.pv.focus(); return; }
    const w = window.open('', 'mo-slides-presenter', 'popup=yes,width=1180,height=740');
    if (!w) { this.app.ctx.toast(t('Slides.Show.PopupBlocked'), 'warning'); return; }
    this.pv = w;
    const d = w.document;
    d.title = t('Slides.Show.Presenter') + ' · ' + (this.app.ctx.doc.name || '');
    d.documentElement.lang = document.documentElement.lang;
    d.documentElement.dataset.theme = 'dark';
    d.documentElement.dataset.kind = 'slides';
    d.body.replaceChildren();
    d.body.className = 'mo-pv-body';
    for (const href of ['/matui/matui.css', '/css/editor.css', '/css/slides.css']) {
      const link = d.createElement('link');
      link.rel = 'stylesheet';
      link.href = new URL(href, location.href).href;
      d.head.append(link);
    }
    ensureStyles(d);
    const sprite = document.querySelector('svg.mat-sprite');
    if (sprite) for (const s of document.querySelectorAll('body > svg')) d.body.append(d.importNode(s, true));
    const btn = (name, label, run) => { const b = d.createElement('button'); b.type = 'button'; b.className = 'mo-pv__btn'; b.setAttribute('aria-label', label); b.title = label; b.append(d.importNode(icon(name), true)); const s = d.createElement('span'); s.textContent = label; b.append(s); b.addEventListener('click', run); return b; };
    const part = (cls, ...kids) => { const n = d.createElement('div'); n.className = cls; n.append(...kids); return n; };
    this.pvTimer = d.createElement('span'); this.pvTimer.className = 'mo-pv__timer';
    this.pvClock = d.createElement('span'); this.pvClock.className = 'mo-pv__clock';
    this.pvCount = d.createElement('span'); this.pvCount.className = 'mo-pv__count';
    this.pvCurrent = part('mo-pv__current');
    this.pvNext = part('mo-pv__next');
    this.pvNotes = part('mo-pv__notes');
    const nextLabel = d.createElement('div'); nextLabel.className = 'mo-pv__label'; nextLabel.textContent = t('Slides.Show.NextSlide');
    const notesLabel = d.createElement('div'); notesLabel.className = 'mo-pv__label'; notesLabel.textContent = t('Slides.Notes');
    const pause = btn('pause', t('Slides.Show.Pause'), () => { if (this.pauseAt) { this.paused += Date.now() - this.pauseAt; this.pauseAt = null; } else this.pauseAt = Date.now(); this.tick(); });
    const reset = btn('refresh', t('Slides.Show.Reset'), () => { this.started = Date.now(); this.paused = 0; this.pauseAt = null; this.tick(); });
    const larger = btn('plus', t('Slides.Show.Larger'), () => { this.notesSize = Math.min(48, (this.notesSize || 20) + 2); this.pvNotes.style.fontSize = this.notesSize + 'px'; });
    const smaller = btn('minus', t('Slides.Show.Smaller'), () => { this.notesSize = Math.max(10, (this.notesSize || 20) - 2); this.pvNotes.style.fontSize = this.notesSize + 'px'; });
    d.body.append(part('mo-pv',
      part('mo-pv__top', this.pvTimer, pause, reset, part('mo-pv__spacer'), this.pvClock, btn('x', t('Slides.Show.End'), () => this.end())),
      part('mo-pv__main', part('mo-pv__left', this.pvCurrent, part('mo-pv__nav', btn('chevron-left', t('Slides.Show.Previous'), () => this.prev()), this.pvCount, btn('chevron-right', t('Slides.Show.Next'), () => this.next()), btn('slides-black', t('Slides.Show.Black'), () => this.blank('black')))),
        part('mo-pv__right', nextLabel, this.pvNext, part('mo-pv__notesbar', notesLabel, smaller, larger), this.pvNotes))));
    d.addEventListener('keydown', e => { this.key(e); });
    w.addEventListener('resize', () => this.updatePresenter());
    clearInterval(this.clock);
    this.clock = setInterval(() => this.tick(), 1000);
    this.tick();
    setTimeout(() => this.updatePresenter(), 50);
  }

  tick() {
    if (!this.pv || this.pv.closed) return;
    const now = this.pauseAt || Date.now();
    const s = Math.max(0, Math.floor((now - this.started - this.paused) / 1000));
    const two = n => String(n).padStart(2, '0');
    this.pvTimer.textContent = `${two(Math.floor(s / 3600))}:${two(Math.floor(s / 60) % 60)}:${two(s % 60)}`;
    this.pvTimer.classList.toggle('is-paused', !!this.pauseAt);
    this.pvClock.textContent = new Date().toLocaleTimeString(document.documentElement.lang || 'en', { hour: '2-digit', minute: '2-digit' });
  }

  updatePresenter() {
    if (!this.pv || this.pv.closed || !this.running) return;
    const d = this.pv.document;
    const { content } = this.app;
    const draw = (box, slide, index, step) => {
      box.replaceChildren();
      if (!slide) { const end = d.createElement('div'); end.className = 'mo-pv__end'; end.textContent = t('Slides.Show.EndScreen'); box.append(end); return; }
      const groups = stepsOf(slide);
      const hide = new Set(groups.slice(step).flatMap(g => g.blocks.flat().map(a => a.el)));
      const node = renderSlide(content, slide, { doc: d, mode: 'view', number: index + 1, hide });
      const r = box.getBoundingClientRect();
      const scale = Math.min(r.width / content.size.w, r.height / content.size.h) || 0.3;
      node.style.transform = `scale(${scale})`;
      node.style.transformOrigin = '0 0';
      const frame = d.createElement('div');
      frame.className = 'mo-pv__frame';
      frame.style.width = content.size.w * scale + 'px';
      frame.style.height = content.size.h * scale + 'px';
      frame.append(node);
      box.append(frame);
      fitText(node);
    };
    draw(this.pvCurrent, this.ended ? null : this.slide, this.slideIndex, this.step);
    if (this.black) { const c = d.createElement('div'); c.className = 'mo-pv__blank is-' + this.black; this.pvCurrent.firstChild?.append(c); }
    const nextIndex = this.order[this.pos + 1];
    const more = this.step < (this.groups?.length || 0);
    draw(this.pvNext, more ? this.slide : nextIndex !== undefined ? content.slides[nextIndex] : null, more ? this.slideIndex : nextIndex, more ? this.step + 1 : 0);
    this.pvNotes.textContent = this.ended ? '' : this.slide.notes || t('Slides.Show.NoNotes');
    this.pvNotes.classList.toggle('is-empty', !this.slide.notes);
    this.pvCount.textContent = t('Slides.SlideNofM', this.slideIndex + 1, content.slides.length);
  }
}
