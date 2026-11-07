/* MatOffice Slides – the pane on the right of the normal view, like the task panes of the office suites:
   "Format Shape" / "Format Picture" (fill, line, shadow, transparency, size and position, text box, crop, alt text), "Format Background"
   when nothing is selected (solid, gradient, picture; apply to all), and the "Animation Pane" (the order of the effects, trigger, duration). */
import { t, el, icon, pickFile, shrinkImage } from '../editor/util.js';
import { resolveColor, round, PT } from './model.js';
import { colorPicker, colorName } from './ui.js';

export class Pane {
  constructor(app) {
    this.app = app;
    this.kind = null;
    this.node = el('aside', { class: 'mo-slides-pane', 'aria-label': t('Slides.Pane.Format'), hidden: true });
  }

  open(kind) { this.kind = kind; this.node.hidden = !kind; this.render(true); }

  close() { this.kind = null; this.node.hidden = true; this.app.layoutChanged(); }

  /** Draws the pane again (not while a field of it has the focus, unless forced). */
  render(force = false) {
    if (!this.kind) return;
    if (!force && this.node.contains(document.activeElement) && document.activeElement.tagName === 'INPUT' && document.activeElement.type !== 'checkbox' && document.activeElement.type !== 'radio') return;
    const app = this.app;
    const sel = app.canvas.selected();
    const title = this.kind === 'animations' ? t('Slides.Pane.Animations') : !sel.length ? t('Slides.Pane.Background') : sel.every(e => e.type === 'image') ? t('Slides.Pane.Picture') : t('Slides.Pane.Shape');
    const head = el('div', { class: 'mo-slides-pane__head' }, el('h2', { text: title }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Editor.Close'), onclick: () => this.close() }, icon('x')));
    const body = el('div', { class: 'mo-slides-pane__body' });
    if (this.kind === 'animations') this.animations(body);
    else if (!sel.length) this.background(body);
    else this.shape(body, sel);
    this.node.replaceChildren(head, body);
  }

  // ---- controls ----------------------------------------------------------------------------------------------------------------

  section(title, ...kids) {
    return el('details', { class: 'mo-slides-pane__section', open: true }, el('summary', { text: title }), el('div', { class: 'mo-slides-pane__rows' }, ...kids));
  }

  row(label, control) { return el('label', { class: 'mo-slides-pane__row' }, el('span', { text: label }), control); }

  number(value, change, { min = -99999, max = 99999, step = 1, unit = '' } = {}) {
    const input = el('input', { type: 'number', class: 'mat-input mo-slides-pane__num', value: value ?? '', min, max, step, disabled: this.app.readOnly ? true : null });
    input.addEventListener('change', () => { const n = Number(input.value); if (Number.isFinite(n)) change(Math.min(max, Math.max(min, n))); });
    return unit ? el('span', { class: 'mo-slides-pane__unit' }, input, el('small', { text: unit })) : input;
  }

  radios(name, options, value, change) {
    const box = el('div', { class: 'mo-slides-pane__radios', role: 'radiogroup' });
    for (const [v, label] of options) {
      const input = el('input', { type: 'radio', name: 'mo-pane-' + name, value: v, checked: v === value ? true : null, disabled: this.app.readOnly ? true : null });
      input.addEventListener('change', () => change(v));
      box.append(el('label', { class: 'mat-check' }, input, el('span', { text: label })));
    }
    return box;
  }

  check(label, value, change) {
    const input = el('input', { type: 'checkbox', checked: value ? true : null, disabled: this.app.readOnly ? true : null });
    input.addEventListener('change', () => change(input.checked));
    return el('label', { class: 'mat-check mo-slides-pane__check' }, input, el('span', { text: label }));
  }

  swatch(value, change, opts = {}) {
    const theme = this.app.content.theme;
    const chip = el('span', { class: 'mo-slides-pane__chip' });
    chip.style.background = value && value !== 'none' ? resolveColor(value, theme) : 'transparent';
    const b = el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm mo-slides-pane__color', disabled: this.app.readOnly ? true : null, 'aria-label': opts.label || t('Slides.Color'), 'data-tip': value ? colorName(value) : '' },
      chip, el('span', { text: value ? colorName(value) : '–' }), icon('chevron-down'));
    b.addEventListener('click', () => colorPicker(b, { theme, current: value, pick: c => change(c), none: opts.none }));
    return b;
  }

  slider(value, change, { min = 0, max = 100, unit = '%' } = {}) {
    const range = el('input', { type: 'range', min, max, value, class: 'mo-slides-pane__range', disabled: this.app.readOnly ? true : null });
    const out = el('output', { text: value + unit });
    range.addEventListener('input', () => { out.textContent = range.value + unit; });
    range.addEventListener('change', () => change(Number(range.value)));
    return el('span', { class: 'mo-slides-pane__slider' }, range, out);
  }

  // ---- shape and picture -------------------------------------------------------------------------------------------------------

  shape(body, sel) {
    const app = this.app;
    const e = sel[0];
    const all = fn => app.editElements(sel.map(x => x.id), fn);
    const shapes = sel.filter(x => x.type === 'shape' || x.type === 'text');
    if (shapes.length) {
      const fill = e.fill || { type: 'none' };
      const kids = [this.radios('fill', [['none', t('Slides.Fill.None')], ['solid', t('Slides.Fill.Solid')], ['gradient', t('Slides.Fill.Gradient')], ['image', t('Slides.Fill.Picture')]], fill.type, async v => {
        if (v === 'image') { const src = await this.pickPicture(); if (src) all(x => { if (x.type !== 'image' && x.type !== 'line' && x.type !== 'table') x.fill = { type: 'image', src }; }); else this.render(true); return; }
        all(x => {
          if (x.type === 'image' || x.type === 'line' || x.type === 'table') return;
          const base = x.fill?.color || x.fill?.from || 'accent1';
          x.fill = v === 'none' ? { type: 'none' } : v === 'solid' ? { type: 'solid', color: base } : { type: 'gradient', from: base.split(':')[0] + ':+40', to: base.split(':')[0] + ':-25', angle: 180 };
        });
      })];
      if (fill.type === 'solid') kids.push(this.row(t('Slides.Color'), this.swatch(fill.color, c => all(x => { if (x.fill?.type === 'solid' || c) x.fill = c === 'none' ? { type: 'none' } : { type: 'solid', color: c || 'accent1' }; }))));
      if (fill.type === 'gradient') {
        kids.push(this.row(t('Slides.Gradient.From'), this.swatch(fill.from, c => all(x => { if (x.fill?.type === 'gradient' && c && c !== 'none') x.fill.from = c; }))));
        kids.push(this.row(t('Slides.Gradient.To'), this.swatch(fill.to, c => all(x => { if (x.fill?.type === 'gradient' && c && c !== 'none') x.fill.to = c; }))));
        kids.push(this.row(t('Slides.Gradient.Angle'), this.number(fill.angle ?? 90, v => all(x => { if (x.fill?.type === 'gradient') x.fill.angle = ((v % 360) + 360) % 360; }), { min: 0, max: 359, unit: '°' })));
      }
      if (fill.type !== 'none') kids.push(this.row(t('Slides.Transparency'), this.slider(Math.round((1 - (e.opacity ?? 1)) * 100), v => all(x => { if (v > 0) x.opacity = round(1 - v / 100); else delete x.opacity; }))));
      body.append(this.section(t('Slides.Pane.Fill'), ...kids));
    }

    const line = e.line;
    const lineKids = [this.radios('line', [['none', t('Slides.Line.None')], ['solid', t('Slides.Line.Solid')]], line ? 'solid' : 'none', v => all(x => {
      if (x.type === 'table') return;
      if (v === 'none') { if (x.type === 'line') return; delete x.line; } else x.line = x.line || { color: 'accent1', width: 1, dash: 'solid' };
    }))];
    if (line) {
      lineKids.push(this.row(t('Slides.Color'), this.swatch(line.color, c => all(x => { if (x.line && c && c !== 'none') x.line.color = c; }))));
      lineKids.push(this.row(t('Slides.Line.Width'), this.number(line.width, v => all(x => { if (x.line) x.line.width = Math.max(0.25, v); }), { min: 0.25, max: 100, step: 0.25, unit: 'pt' })));
      const dash = el('select', { class: 'mat-input mo-slides-pane__select', disabled: app.readOnly ? true : null }, ...['solid', 'dash', 'dot', 'dashDot', 'longDash'].map(d => el('option', { value: d, text: t('Slides.Dash.' + d), selected: line.dash === d ? true : null })));
      dash.addEventListener('change', () => all(x => { if (x.line) x.line.dash = dash.value; }));
      lineKids.push(this.row(t('Slides.Line.Dash'), dash));
      if (e.type === 'line') {
        for (const end of ['arrowStart', 'arrowEnd']) {
          const s = el('select', { class: 'mat-input mo-slides-pane__select', disabled: app.readOnly ? true : null }, ...['none', 'triangle', 'arrow', 'stealth', 'oval', 'diamond'].map(a => el('option', { value: a, text: t('Slides.Arrow.' + a), selected: (e[end] || 'none') === a ? true : null })));
          s.addEventListener('change', () => all(x => { if (x.type === 'line') x[end] = s.value; }));
          lineKids.push(this.row(t('Slides.Line.' + end), s));
        }
      }
    }
    if (!sel.every(x => x.type === 'table')) body.append(this.section(e.type === 'image' ? t('Slides.Pane.Border') : t('Slides.Pane.Line'), ...lineKids));

    if (!sel.every(x => x.type === 'table')) {
      body.append(this.section(t('Slides.Pane.Effects'), this.check(t('Slides.Shadow'), !!e.shadow, v => all(x => { if (x.type !== 'table') { if (v) x.shadow = true; else delete x.shadow; } }))));
    }

    if (e.type === 'image' && sel.length === 1) {
      const crop = e.crop || { l: 0, t: 0, r: 0, b: 0 };
      const c = k => this.number(Math.round(crop[k] * 100), v => app.editElements([e.id], x => { const cr = x.crop || { l: 0, t: 0, r: 0, b: 0 }; cr[k] = Math.max(0, Math.min(0.9, v / 100)); if (cr.l + cr.r < 0.95 && cr.t + cr.b < 0.95) x.crop = cr; if (!cr.l && !cr.t && !cr.r && !cr.b) delete x.crop; }), { min: 0, max: 90, unit: '%' });
      body.append(this.section(t('Slides.Pane.Picture'),
        this.row(t('Slides.Transparency'), this.slider(Math.round((1 - (e.opacity ?? 1)) * 100), v => app.editElements([e.id], x => { if (v > 0) x.opacity = round(1 - v / 100); else delete x.opacity; }))),
        el('div', { class: 'mo-slides-pane__heading', text: t('Slides.Crop') }),
        this.row(t('Slides.Crop.Left'), c('l')), this.row(t('Slides.Crop.Top'), c('t')), this.row(t('Slides.Crop.Right'), c('r')), this.row(t('Slides.Crop.Bottom'), c('b')),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: app.readOnly ? true : null, onclick: () => app.editElements([e.id], x => { delete x.crop; }) }, el('span', { text: t('Slides.Crop.Reset') }))));
      const alt = el('textarea', { class: 'mat-input mo-slides-pane__alt', rows: 3, disabled: app.readOnly ? true : null, placeholder: t('Slides.AltText.Hint') });
      alt.value = e.alt || '';
      alt.addEventListener('change', () => app.editElements([e.id], x => { if (alt.value.trim()) x.alt = alt.value.trim().slice(0, 2000); else delete x.alt; }));
      body.append(this.section(t('Slides.AltText'), alt));
    }

    if (sel.length === 1) {
      const kids = [
        this.row(t('Slides.Size.Width'), this.number(round(e.w), v => app.setGeometry(e.id, { w: Math.max(1, v) }), { min: 1, max: 20000, unit: 'px' })),
        this.row(t('Slides.Size.Height'), this.number(round(e.h), v => app.setGeometry(e.id, { h: Math.max(e.type === 'line' ? 0 : 1, v) }), { min: 0, max: 20000, unit: 'px' })),
        this.row(t('Slides.Size.Left'), this.number(round(e.x), v => app.setGeometry(e.id, { x: v }), { unit: 'px' })),
        this.row(t('Slides.Size.Top'), this.number(round(e.y), v => app.setGeometry(e.id, { y: v }), { unit: 'px' })),
      ];
      if (e.type !== 'table') kids.push(this.row(t('Slides.Size.Rotation'), this.number(round(e.rot || 0), v => app.setGeometry(e.id, { rot: ((v % 360) + 360) % 360 }), { min: -360, max: 360, unit: '°' })));
      body.append(this.section(t('Slides.Pane.Size'), ...kids));
    }

    const texts = sel.filter(x => x.text && (x.type === 'text' || x.type === 'shape'));
    if (texts.length) {
      const tb = texts[0].text;
      const d = app.textDefaultsOf(texts[0]);
      const inset = tb.inset || [9.6, 4.8, 9.6, 4.8];
      const setInset = (i, v) => app.editElements(texts.map(x => x.id), x => { const ins = [...(x.text.inset || [9.6, 4.8, 9.6, 4.8])]; ins[i] = Math.max(0, v); x.text.inset = ins; });
      body.append(this.section(t('Slides.Pane.TextBox'),
        this.row(t('Slides.Anchor'), this.radios('anchor', [['top', t('Slides.Anchor.top')], ['middle', t('Slides.Anchor.middle')], ['bottom', t('Slides.Anchor.bottom')]], tb.anchor || d.anchor, v => app.editElements(texts.map(x => x.id), x => { x.text.anchor = v; }))),
        el('div', { class: 'mo-slides-pane__heading', text: t('Slides.Autofit') }),
        this.radios('autofit', [['none', t('Slides.Autofit.none')], ['shrink', t('Slides.Autofit.shrink')], ['grow', t('Slides.Autofit.grow')]], tb.autofit || d.autofit, v => app.editElements(texts.map(x => x.id), x => { x.text.autofit = v; })),
        this.row(t('Slides.Inset.Left'), this.number(round(inset[0]), v => setInset(0, v), { min: 0, max: 500, unit: 'px' })),
        this.row(t('Slides.Inset.Top'), this.number(round(inset[1]), v => setInset(1, v), { min: 0, max: 500, unit: 'px' })),
        this.row(t('Slides.Inset.Right'), this.number(round(inset[2]), v => setInset(2, v), { min: 0, max: 500, unit: 'px' })),
        this.row(t('Slides.Inset.Bottom'), this.number(round(inset[3]), v => setInset(3, v), { min: 0, max: 500, unit: 'px' })),
        this.check(t('Slides.Wrap'), tb.wrap !== false, v => app.editElements(texts.map(x => x.id), x => { if (v) delete x.text.wrap; else x.text.wrap = false; }))));
    }
    void PT;
  }

  async pickPicture() {
    const file = await pickFile('image/png,image/jpeg,image/gif,image/webp');
    if (!file) return null;
    return this.app.pictureData(file);
  }

  // ---- background --------------------------------------------------------------------------------------------------------------

  background(body) {
    const app = this.app;
    const slide = app.slide;
    if (!slide) return;
    const fill = app.backgroundOfSlide();
    const set = f => app.setBackground(f, false);
    const kids = [this.radios('bg', [['solid', t('Slides.Fill.Solid')], ['gradient', t('Slides.Fill.Gradient')], ['image', t('Slides.Fill.Picture')]], fill.type, async v => {
      if (v === 'image') { const src = await this.pickPicture(); if (src) set({ type: 'image', src }); else this.render(true); return; }
      const base = fill.color || fill.from || 'lt1';
      set(v === 'solid' ? { type: 'solid', color: base } : { type: 'gradient', from: base, to: (base.split(':')[0]) + ':-25', angle: 180 });
    })];
    if (fill.type === 'solid') kids.push(this.row(t('Slides.Color'), this.swatch(fill.color, c => c && c !== 'none' && set({ type: 'solid', color: c }))));
    if (fill.type === 'gradient') {
      kids.push(this.row(t('Slides.Gradient.From'), this.swatch(fill.from, c => c && c !== 'none' && set({ ...fill, from: c }))));
      kids.push(this.row(t('Slides.Gradient.To'), this.swatch(fill.to, c => c && c !== 'none' && set({ ...fill, to: c }))));
      kids.push(this.row(t('Slides.Gradient.Angle'), this.number(fill.angle ?? 90, v => set({ ...fill, angle: ((v % 360) + 360) % 360 }), { min: 0, max: 359, unit: '°' })));
    }
    if (fill.type === 'image') kids.push(el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: app.readOnly ? true : null, onclick: async () => { const src = await this.pickPicture(); if (src) set({ type: 'image', src }); } }, icon('image'), el('span', { text: t('Slides.Background.Picture') })));
    body.append(this.section(t('Slides.Pane.Fill'), ...kids),
      el('div', { class: 'mo-slides-pane__buttons' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: app.readOnly ? true : null, onclick: () => app.setBackground(app.backgroundOfSlide(), true) }, el('span', { text: t('Slides.Background.ApplyAll') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: app.readOnly ? true : null, onclick: () => app.setBackground(null, false) }, el('span', { text: t('Slides.Background.Reset') }))));
  }

  // ---- animations ---------------------------------------------------------------------------------------------------------------

  animations(body) {
    const app = this.app;
    const slide = app.slide;
    const list = slide?.animations || [];
    if (!list.length) { body.append(el('p', { class: 'mo-muted', text: t('Slides.Anim.None') })); return; }
    const ol = el('ol', { class: 'mo-slides-animlist' });
    let step = 0;
    list.forEach((a, i) => {
      if (a.trigger === 'click' || i === 0) step += a.trigger === 'click' ? 1 : 0;
      const target = slide.elements.find(e => e.id === a.el);
      const name = target ? app.elementName(target) : '?';
      const selected = app.selection.includes(a.el);
      const trig = el('select', { class: 'mat-input mo-slides-pane__select', 'aria-label': t('Slides.Anim.Trigger'), disabled: app.readOnly ? true : null },
        ...['click', 'with', 'after'].map(v => el('option', { value: v, text: t('Slides.Anim.' + v), selected: a.trigger === v ? true : null })));
      trig.addEventListener('change', () => app.editAnimation(i, x => { x.trigger = trig.value; }));
      const dur = this.number(a.duration ?? 0.5, v => app.editAnimation(i, x => { x.duration = Math.max(0.05, v); }), { min: 0.05, max: 20, step: 0.05, unit: 's' });
      const li = el('li', { class: 'mo-slides-animlist__item' + (selected ? ' is-selected' : '') },
        el('button', { type: 'button', class: 'mo-slides-animlist__name', onclick: () => app.select([a.el]) },
          el('span', { class: 'mo-slides-animlist__step', text: a.trigger === 'click' ? String(step) : '' }), icon('slides-anim-' + a.effect), el('span', { text: `${t('Slides.Effect.' + a.effect)} · ${name}` })),
        el('div', { class: 'mo-slides-animlist__ctrl' }, trig, dur,
          el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Slides.Anim.Earlier'), 'data-tip': t('Slides.Anim.Earlier'), disabled: i === 0 || app.readOnly ? true : null, onclick: () => app.moveAnimation(i, -1) }, icon('arrow-up')),
          el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Slides.Anim.Later'), 'data-tip': t('Slides.Anim.Later'), disabled: i === list.length - 1 || app.readOnly ? true : null, onclick: () => app.moveAnimation(i, 1) }, icon('arrow-down')),
          el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Slides.Anim.Remove'), 'data-tip': t('Slides.Anim.Remove'), disabled: app.readOnly ? true : null, onclick: () => app.removeAnimation(i) }, icon('trash'))));
      ol.append(li);
    });
    body.append(ol, el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => app.previewAnimations() }, icon('play'), el('span', { text: t('Slides.Anim.Play') })));
  }
}
