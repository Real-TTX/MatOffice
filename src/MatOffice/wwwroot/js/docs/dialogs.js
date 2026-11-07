/* MatOffice Docs – the dialogs (built with el(), never HTML): Paragraph, Page Setup, Font, Link, Symbol, Date and Time, Word Count,
   Insert Table, Go To, Language, Set Numbering Value, Caption. */
import { readBlock } from './htmlio.js';
import {
  FONTS, SIZES, PAGE_SIZES, numStr, pt, toPt, styleOf, styleKey, statistics, runsText, linkFromInput, textRun, fontStack, sliceRuns, insertRuns, formatRuns,
  sizeName, languageName,
} from './model.js';
import { LINE_SPACINGS, currentSpacing, currentLineHeight } from './commands.js';

const PT_PER_CM = 72 / 2.54;

const LANGUAGES = ['en', 'en-US', 'en-GB', 'de', 'de-DE', 'de-AT', 'de-CH', 'fr', 'fr-FR', 'es', 'es-ES', 'it', 'it-IT', 'nl', 'nl-NL', 'pt', 'pt-BR', 'pl', 'cs', 'da', 'sv', 'nb', 'fi', 'tr', 'ru', 'uk', 'el', 'ja', 'zh', 'ko'];

const SYMBOLS = {
  common: '©®™§¶†‡°±×÷≈≠≤≥∞√∑∏∫∂∆µπΩ€£¥¢$₹₽¼½¾⅓⅔‰№',
  punctuation: '–—…•·‹›«»‚„“”‘’¡¿⁂※',
  arrows: '←↑→↓↔↕⇐⇑⇒⇓⇔➔➢➤▶◀▲▼',
  shapes: '■□▪▫●○◆◇★☆✓✔✗✘☐☑☒♠♣♥♦',
  greek: 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ',
  letters: 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿŒœŠšŽžŁłČčŘř',
};

export class Dialogs {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
  }

  get t() { return this.app.ctx.t; }
  get el() { return this.app.ctx.el; }

  field(label, input) { return this.el('label', {}, this.el('span', { text: label }), input); }

  number(name, value, { min = 0, max = 1000, step = 1 } = {}) {
    return this.el('input', { class: 'mat-input', type: 'number', name, value: value === '' || value === null || value === undefined ? '' : String(value), min, max, step, inputmode: 'decimal' });
  }

  select(name, options, value) {
    const s = this.el('select', { class: 'mat-input', name });
    for (const [v, label] of options) s.append(this.el('option', { value: v, text: label }));
    s.value = value;
    return s;
  }

  buttons() { return [{ label: this.t('Docs.Ok'), value: 'ok', kind: 'primary' }, { label: this.t('Docs.Cancel'), value: 'cancel' }]; }

  /** Runs a dialog with the selection kept (the focus goes into the dialog and back). */
  async ask(options) {
    const mark = this.ed.bookmark();
    const answer = await this.app.ctx.dialog({ buttons: this.buttons(), ...options });
    if (mark) this.ed.restoreBookmark(mark, false);
    return answer;
  }

  // ---- Paragraph --------------------------------------------------------------------------------------------------------------------

  async paragraph() {
    const t = this.t, el = this.el;
    const s = this.ed.state();
    const b = s.block;
    if (!b) return;
    const css = b.css || {};
    const align = this.select('align', [['left', t('Docs.AlignLeft')], ['center', t('Docs.AlignCenter')], ['right', t('Docs.AlignRight')], ['justify', t('Docs.Justify')]], css['text-align'] || (b.type === 'blockquote' ? 'center' : 'left'));
    const left = this.number('left', numStr((toPt(css['margin-left']) ?? 0) / PT_PER_CM), { max: 50, step: 0.1 });
    const right = this.number('right', numStr((toPt(css['margin-right']) ?? 0) / PT_PER_CM), { max: 50, step: 0.1 });
    const indent = toPt(css['text-indent']) ?? 0;
    const special = this.select('special', [['none', t('Docs.Para.None')], ['first', t('Docs.Para.FirstLine')], ['hanging', t('Docs.Para.Hanging')]], indent > 0 ? 'first' : indent < 0 ? 'hanging' : 'none');
    const by = this.number('by', numStr(Math.abs(indent) / PT_PER_CM), { max: 20, step: 0.1 });
    const before = this.number('before', numStr(currentSpacing(b, 'before')), { max: 1584, step: 6 });
    const after = this.number('after', numStr(currentSpacing(b, 'after')), { max: 1584, step: 6 });
    const lh = currentLineHeight(b);
    const exact = /pt$/.test(lh);
    const known = LINE_SPACINGS.find(([, v]) => Math.abs(parseFloat(v) - parseFloat(lh)) < 0.02);
    const rule = this.select('rule', [['1.2', t('Docs.Para.Single')], ['1.38', '1,15'.replace(',', t('Docs.DecimalPoint'))], ['1.8', t('Docs.Para.OneHalf')], ['2.4', t('Docs.Para.Double')], ['multiple', t('Docs.Para.Multiple')], ['exact', t('Docs.Para.Exactly')]],
      exact ? 'exact' : known && ['1.2', '1.38', '1.8', '2.4'].includes(known[1]) ? known[1] : 'multiple');
    const at = this.number('at', exact ? numStr(parseFloat(lh)) : numStr(parseFloat(lh) / 1.2), { min: 0.5, max: 1000, step: 0.1 });
    const syncAt = () => { at.disabled = !(rule.value === 'multiple' || rule.value === 'exact'); };
    rule.addEventListener('change', () => { if (rule.value === 'exact') at.value = '12'; else if (rule.value === 'multiple') at.value = '1.08'; syncAt(); });
    syncAt();
    const body = el('div', { class: 'mo-docs-form' },
      el('fieldset', {}, el('legend', { text: t('Docs.Para.General') }), this.field(t('Docs.Para.Alignment'), align)),
      el('fieldset', {}, el('legend', { text: t('Docs.Para.Indentation') }),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Para.LeftCm'), left), this.field(t('Docs.Para.RightCm'), right)),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Para.Special'), special), this.field(t('Docs.Para.ByCm'), by))),
      el('fieldset', {}, el('legend', { text: t('Docs.Para.Spacing') }),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Para.BeforePt'), before), this.field(t('Docs.Para.AfterPt'), after)),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.LineSpacing'), rule), this.field(t('Docs.Para.At'), at))));
    const answer = await this.ask({ title: t('Docs.Paragraph.Dialog'), icon: 'docs-paragraph', body, size: 'md' });
    if (answer !== 'ok') return;
    const n = v => Math.max(0, Number(String(v).replace(',', '.')) || 0);
    const byPt = n(by.value) * PT_PER_CM;
    const line = rule.value === 'exact' ? pt(Math.max(1, n(at.value))) : rule.value === 'multiple' ? numStr(Math.max(0.5, Math.min(10, n(at.value) * 1.2))) : rule.value;
    this.ed.focus();
    this.ed.formatBlocks(blk => {
      const style = styleOf(styleKey(blk));
      const c = { ...blk.css };
      const set = (k, v) => { if (v === null) delete c[k]; else c[k] = v; };
      set('text-align', align.value === 'left' && blk.type !== 'blockquote' ? null : align.value);
      set('margin-left', n(left.value) ? pt(n(left.value) * PT_PER_CM) : null);
      set('margin-right', n(right.value) ? pt(n(right.value) * PT_PER_CM) : null);
      set('text-indent', special.value === 'none' || !byPt ? null : pt(special.value === 'first' ? byPt : -byPt));
      set('margin-top', Math.abs(n(before.value) - style.before) < 0.01 ? null : pt(n(before.value)));
      set('margin-bottom', Math.abs(n(after.value) - style.after) < 0.01 ? null : pt(n(after.value)));
      set('line-height', Math.abs(parseFloat(line) - style.line) < 0.01 && !/pt$/.test(line) ? null : line);
      return { ...blk, css: c };
    });
  }

  // ---- Page Setup ----------------------------------------------------------------------------------------------------------------------

  async pageSetup(focus = 'margins') {
    const t = this.t, el = this.el;
    const p = this.app.page;
    const cm = v => numStr(v / 10);
    const top = this.number('top', cm(p.margins.top), { max: 15, step: 0.1 });
    const bottom = this.number('bottom', cm(p.margins.bottom), { max: 15, step: 0.1 });
    const left = this.number('left', cm(p.margins.left), { max: 15, step: 0.1 });
    const right = this.number('right', cm(p.margins.right), { max: 15, step: 0.1 });
    const orientation = this.select('orientation', [['portrait', t('Docs.Portrait')], ['landscape', t('Docs.Landscape')]], p.orientation);
    const size = this.select('size', [...Object.keys(PAGE_SIZES).map(k => [k, k]), ['Custom', t('Docs.Size.CustomName')]], p.size);
    const width = this.number('width', cm(p.width), { min: 5, max: 100, step: 0.1 });
    const height = this.number('height', cm(p.height), { min: 5, max: 100, step: 0.1 });
    const header = this.number('header', cm(p.header), { max: 15, step: 0.1 });
    const footer = this.number('footer', cm(p.footer), { max: 15, step: 0.1 });
    const columns = this.select('columns', [['1', t('Docs.Columns.1')], ['2', t('Docs.Columns.2')], ['3', t('Docs.Columns.3')]], String(p.columns));
    const first = el('input', { type: 'checkbox', name: 'first' });
    first.checked = p.differentFirst;
    const sync = () => {
      const custom = size.value === 'Custom';
      width.disabled = height.disabled = !custom;
      if (!custom) { const [w, h] = PAGE_SIZES[size.value]; width.value = cm(w); height.value = cm(h); }
    };
    size.addEventListener('change', sync);
    sync();
    const body = el('div', { class: 'mo-docs-form' },
      el('fieldset', {}, el('legend', { text: t('Docs.Margins') }),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Margin.TopCm'), top), this.field(t('Docs.Margin.BottomCm'), bottom)),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Margin.LeftCm'), left), this.field(t('Docs.Margin.RightCm'), right))),
      el('fieldset', {}, el('legend', { text: t('Docs.Paper') }),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Size'), size), this.field(t('Docs.Orientation'), orientation)),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Size.WidthCm'), width), this.field(t('Docs.Size.HeightCm'), height))),
      el('fieldset', {}, el('legend', { text: t('Docs.Layout') }),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.HeaderFrom'), header), this.field(t('Docs.FooterFrom'), footer)),
        el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Columns'), columns), el('label', { class: 'mo-docs-check' }, first, el('span', { text: t('Docs.DifferentFirst') })))));
    const answer = await this.ask({ title: t('Docs.PageSetup'), icon: 'docs-margins', body, size: 'md', onOpen: () => (focus === 'paper' ? size : top).focus() });
    if (answer !== 'ok') return;
    const mm = input => Math.max(0, Number(String(input.value).replace(',', '.')) || 0) * 10;
    const w = mm(width), h = mm(height);
    const named = size.value === 'Custom' ? sizeName(Math.min(w, h), Math.max(w, h)) : size.value;
    this.app.layoutCommands.setPage({
      margins: { top: mm(top), right: mm(right), bottom: mm(bottom), left: mm(left) },
      size: named, width: Math.min(w, h), height: Math.max(w, h), orientation: orientation.value,
      header: mm(header), footer: mm(footer), columns: Number(columns.value), differentFirst: first.checked,
    });
  }

  // ---- Font ------------------------------------------------------------------------------------------------------------------------------

  async font() {
    const t = this.t, el = this.el;
    const s = this.ed.state();
    const f = s.format || {};
    const family = this.select('font', FONTS.map(x => [x, x]), s.fontName && FONTS.includes(s.fontName) ? s.fontName : 'Calibri');
    const size = this.el('input', { class: 'mat-input', name: 'size', list: 'mo-docs-sizes', value: s.sizePt ? numStr(s.sizePt) : '', inputmode: 'decimal' });
    const list = el('datalist', { id: 'mo-docs-sizes' }, ...SIZES.map(x => el('option', { value: String(x) })));
    const colour = el('input', { type: 'color', name: 'color', value: f.color || '#000000', class: 'mo-palette__custom' });
    const checks = ['b', 'i', 'u', 's', 'sup', 'sub'].map(k => {
      const box = el('input', { type: 'checkbox', name: k });
      box.checked = !!f[k];
      return [k, box];
    });
    const preview = el('div', { class: 'mo-docs-preview', text: t('Docs.Font.Preview') });
    const paint = () => {
      preview.style.fontFamily = fontStack(family.value);
      preview.style.fontSize = (parseFloat(size.value) || 11) + 'pt';
      preview.style.color = colour.value;
      const on = k => checks.find(c => c[0] === k)[1].checked;
      preview.style.fontWeight = on('b') ? 'bold' : 'normal';
      preview.style.fontStyle = on('i') ? 'italic' : 'normal';
      preview.style.textDecoration = [on('u') ? 'underline' : '', on('s') ? 'line-through' : ''].join(' ').trim() || 'none';
    };
    for (const x of [family, size, colour, ...checks.map(c => c[1])]) x.addEventListener('input', paint);
    paint();
    const body = el('div', { class: 'mo-docs-form' },
      el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Font'), family), this.field(t('Docs.Size'), el('span', {}, size, list)), this.field(t('Docs.FontColour'), colour)),
      el('fieldset', {}, el('legend', { text: t('Docs.Font.Effects') }), el('div', { class: 'mo-docs-form__row' },
        ...checks.map(([k, box]) => el('label', { class: 'mo-docs-check' }, box, el('span', { text: t('Docs.Font.' + k) }))))),
      preview);
    const answer = await this.ask({ title: t('Docs.Font.Dialog'), icon: 'type', body, size: 'md' });
    if (answer !== 'ok') return;
    const patch = { font: fontStack(family.value), color: colour.value === '#000000' && !f.color ? null : colour.value };
    const sz = parseFloat(String(size.value).replace(',', '.'));
    if (sz > 0 && sz <= 1638) patch.size = pt(sz);
    for (const [k, box] of checks) patch[k] = box.checked;
    if (patch.sub && patch.sup) patch.sub = false;
    this.ed.focus();
    this.ed.formatText(x => ({ ...x, ...patch }));
  }

  // ---- Link ------------------------------------------------------------------------------------------------------------------------------

  async link() {
    const t = this.t, el = this.el;
    if (this.app.ctx.readOnly) return;
    this.ed.focus();
    const sel = this.ed.selection();
    if (!sel) return;
    const b = readBlock(sel.start.block);
    // inside a link: its whole extent
    let start = sel.start.offset, end = sel.start.block === sel.end.block ? sel.end.offset : runsText(b.runs).length;
    const href = this.ed.state().format.href || '';
    if (href && sel.collapsed) {
      let pos = 0; const spans = [];
      for (const r of b.runs) { const len = r.obj ? 1 : r.text.length; if (r.f.href === href) spans.push([pos, pos + len]); pos += len; }
      const hit = spans.find(([a, z]) => a <= start && start <= z);
      if (hit) { start = hit[0]; end = hit[1]; while (spans.some(([a, z]) => a === end && z > end)) end = spans.find(([a]) => a === end)[1]; }
    }
    const shown = runsText(sliceRuns(b.runs, start, end), '');
    const text = el('input', { class: 'mat-input', name: 'text', value: shown, maxlength: 2000 });
    const address = el('input', { class: 'mat-input', name: 'address', value: href, placeholder: 'https://', inputmode: 'url', autocomplete: 'off' });
    const headings = [...this.app.body.querySelectorAll(':scope > h1, :scope > h2, :scope > h3')];
    const place = this.select('place', [['', t('Docs.Link.Web')], ...headings.map((h, i) => [String(i), (h.textContent || '').trim().slice(0, 80) || '…'])], '');
    place.addEventListener('change', () => { if (place.value !== '') address.value = '#'; });
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.Link.Text'), text), this.field(t('Docs.Link.Address'), address), headings.length ? this.field(t('Docs.Link.Place'), place) : null);
    const buttons = [{ label: t('Docs.Ok'), value: 'ok', kind: 'primary' }, { label: t('Docs.Cancel'), value: 'cancel' }];
    if (href) buttons.push({ label: t('Docs.Link.Remove'), value: 'remove', kind: 'danger' });
    const answer = await this.ask({
      title: t('Docs.Link'), icon: 'link', body, buttons, size: 'md', onOpen: () => (shown ? address : text).focus(),
      validate: (form, value) => {
        if (value !== 'ok') return true;
        if (place.value !== '') return true;
        if (!linkFromInput(address.value)) { address.setCustomValidity(t('Docs.Link.Invalid')); address.reportValidity(); return false; }
        return true;
      },
    });
    if (!answer || answer === 'cancel') return;
    const el2 = sel.start.block;
    if (!el2.isConnected) return;
    this.ed.checkpoint('link');
    const block = readBlock(el2);
    if (answer === 'remove') {
      block.runs = formatRuns(block.runs, start, end, f => { const g = { ...f }; delete g.href; return g; });
      this.ed.morph(el2, block);
      this.ed.select(el2, end);
      return;
    }
    let target;
    if (place.value !== '') {
      const h = headings[Number(place.value)];
      if (!h.id) { let id; do id = 'h-' + Math.random().toString(36).slice(2, 8); while (document.getElementById(id)); h.id = id; }
      target = '#' + h.id;
    } else target = linkFromInput(address.value);
    if (!target) return;
    const label = text.value || address.value;
    if (label !== shown || start === end) {
      const f = { ...(block.runs.length ? (this.ed.currentFormat()) : {}), href: target };
      delete f.cmt;
      block.runs = insertRuns(runsDelete(block.runs, start, end), start, [textRun(label, f)]);
      this.ed.morph(el2, block);
      this.ed.select(el2, start + label.length);
    } else {
      block.runs = formatRuns(block.runs, start, end, { href: target });
      this.ed.morph(el2, block);
      this.ed.select(el2, end);
    }
  }

  // ---- Symbols, date, word count, table, go to, language, numbering value, caption ------------------------------------------------------

  async symbols() {
    const t = this.t, el = this.el;
    const groups = Object.entries(SYMBOLS);
    const grid = el('div', { class: 'mo-docs-symbols', role: 'listbox', 'aria-label': t('Docs.Symbol') });
    const group = this.select('group', groups.map(([k]) => [k, t('Docs.Symbol.' + k)]), 'common');
    let chosen = null;
    let resolveDialog;
    const fill = () => {
      grid.replaceChildren(...[...SYMBOLS[group.value]].map(ch => el('button', {
        type: 'button', class: 'mo-docs-symbol', text: ch, 'aria-label': ch + ' U+' + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'),
        'data-tip': 'U+' + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'),
        onclick: () => { chosen = ch; resolveDialog?.(); },
      })));
    };
    group.addEventListener('change', fill);
    fill();
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.Symbol.Group'), group), grid);
    const mark = this.ed.bookmark();
    const promise = this.app.ctx.dialog({ title: t('Docs.Symbol'), icon: 'docs-symbol', body, size: 'md', buttons: [{ label: t('Docs.Close'), value: 'cancel' }] });
    resolveDialog = () => { const d = grid.closest('dialog'); if (d) window.MatUi.dialog.close(d, 'ok'); };
    await promise;
    if (mark) this.ed.restoreBookmark(mark, false);
    if (chosen) this.app.commands.text(chosen);
  }

  async dateTime() {
    const t = this.t, el = this.el;
    const now = new Date();
    const lang = this.app.language();
    const formats = [
      { dateStyle: 'short' }, { dateStyle: 'medium' }, { dateStyle: 'long' }, { dateStyle: 'full' },
      { year: 'numeric', month: '2-digit', day: '2-digit' }, { timeStyle: 'short' }, { dateStyle: 'short', timeStyle: 'short' }, { dateStyle: 'long', timeStyle: 'short' },
    ].map(o => now.toLocaleString(lang, o));
    const iso = now.toISOString().slice(0, 10);
    const all = [...new Set([...formats, iso])];
    const list = this.select('format', all.map(x => [x, x]), all[0]);
    list.size = Math.min(9, all.length);
    const auto = el('input', { type: 'checkbox', name: 'auto' });
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.DateTime.Formats'), list), el('label', { class: 'mo-docs-check' }, auto, el('span', { text: t('Docs.DateTime.Update') })));
    const answer = await this.ask({ title: t('Docs.DateTime'), icon: 'calendar', body, onOpen: () => list.focus() });
    if (answer !== 'ok') return;
    if (auto.checked) this.app.commands.field('date');
    else this.app.commands.text(list.value);
  }

  async wordCount() {
    const t = this.t, el = this.el;
    const sel = this.ed.selection();
    let texts;
    let scope = t('Docs.WordCount.Document');
    if (sel && !sel.collapsed && sel.host === this.app.body) {
      texts = this.ed.ranges(sel).map(({ el: b, start, end }) => runsText(sliceRuns(readBlock(b).runs, start, end)));
      scope = t('Docs.WordCount.Selection');
    } else texts = this.ed.textBlocks(this.app.body).map(b => runsText(readBlock(b).runs));
    const notes = el('input', { type: 'checkbox' });
    const values = el('dl', { class: 'mo-docs-stats' });
    const paint = () => {
      const all = notes.checked ? [...texts, ...this.app.notes.hosts().flatMap(h => this.ed.textBlocks(h).map(b => runsText(readBlock(b).runs)))] : texts;
      const s = statistics(all);
      const n = v => v.toLocaleString(this.app.language());
      values.replaceChildren(
        el('dt', { text: t('Docs.WordCount.Pages') }), el('dd', { text: n(this.app.view.pages) }),
        el('dt', { text: t('Docs.WordCount.Words') }), el('dd', { text: n(s.words) }),
        el('dt', { text: t('Docs.WordCount.CharsNoSpaces') }), el('dd', { text: n(s.charsNoSpaces) }),
        el('dt', { text: t('Docs.WordCount.Chars') }), el('dd', { text: n(s.chars) }),
        el('dt', { text: t('Docs.WordCount.Paragraphs') }), el('dd', { text: n(s.paragraphs) }));
    };
    notes.addEventListener('change', paint);
    paint();
    const body = el('div', { class: 'mo-docs-form' }, el('p', { class: 'mo-muted', text: scope }), values, el('label', { class: 'mo-docs-check' }, notes, el('span', { text: t('Docs.WordCount.Notes') })));
    await this.ask({ title: t('Docs.WordCount'), icon: 'docs-word-count', body, buttons: [{ label: t('Docs.Close'), value: 'ok', kind: 'primary' }] });
  }

  async insertTable() {
    const t = this.t, el = this.el;
    const cols = this.number('cols', 3, { min: 1, max: 63 });
    const rows = this.number('rows', 3, { min: 1, max: 500 });
    const header = el('input', { type: 'checkbox', name: 'header' });
    const style = this.select('style', ['grid', 'plain', 'light', 'striped', 'accent1', 'accent2', 'accent6'].map(s => [s, t('Docs.TableStyle.' + s)]), 'grid');
    const body = el('div', { class: 'mo-docs-form' },
      el('div', { class: 'mo-docs-form__row' }, this.field(t('Docs.Table.Columns'), cols), this.field(t('Docs.Table.Rows'), rows)),
      this.field(t('Docs.Table.Styles'), style), el('label', { class: 'mo-docs-check' }, header, el('span', { text: t('Docs.Table.HeaderRow') })));
    const answer = await this.ask({ title: t('Docs.Table.Dialog'), icon: 'table', body, onOpen: () => cols.select() });
    if (answer !== 'ok') return;
    this.app.tables.insert(Number(rows.value) || 1, Number(cols.value) || 1, { header: header.checked, style: style.value });
  }

  async goTo() {
    const t = this.t, el = this.el;
    const page = this.number('page', this.app.view.visiblePage() + 1, { min: 1, max: this.app.view.pages });
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.GoTo.Page', this.app.view.pages), page));
    const answer = await this.app.ctx.dialog({ title: t('Docs.GoTo'), icon: 'docs-goto', body, buttons: this.buttons(), onOpen: () => page.select() });
    if (answer !== 'ok') return;
    const n = Math.min(this.app.view.pages, Math.max(1, Number(page.value) || 1));
    if (this.app.viewMode !== 'print') this.app.setView('print');
    this.app.view.scrollToPage(n - 1);
    // the caret to the first line of that page
    setTimeout(() => {
      const target = this.ed.textBlocks(this.app.body).find(b => { const r = b.getClientRects(); return r.length && this.app.view.pageOfRect(r[r.length - 1]) >= n - 1; });
      if (target) this.ed.select(target, 0, target, 0, false);
    }, 400);
  }

  async language() {
    const t = this.t, el = this.el;
    const current = this.app.settings.language;
    const list = this.select('language', [...new Set([current, ...LANGUAGES])].map(code => [code, `${languageName(code, this.app.uiLanguage())} (${code})`]), current);
    list.size = 10;
    const spell = el('input', { type: 'checkbox', name: 'spell' });
    spell.checked = this.app.settings.spellcheck;
    const body = el('div', { class: 'mo-docs-form' }, el('p', { class: 'mo-muted', text: t('Docs.Language.Text') }), list, el('label', { class: 'mo-docs-check' }, spell, el('span', { text: t('Docs.Spelling.Check') })));
    const answer = await this.ask({ title: t('Docs.Language'), icon: 'globe', body, onOpen: () => list.focus() });
    if (answer !== 'ok') return;
    this.app.layoutCommands.setSettings({ language: list.value, spellcheck: spell.checked });
  }

  async numberValue() {
    const t = this.t, el = this.el;
    const b = this.ed.state().block;
    if (!b?.list) return;
    const value = this.number('value', b.start || 1, { min: 1, max: 99999 });
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.List.SetValueTo'), value));
    const answer = await this.ask({ title: t('Docs.List.SetValue'), icon: 'docs-numbering', body, onOpen: () => value.select() });
    if (answer !== 'ok') return;
    this.app.commands.numberFrom(Math.min(99999, Math.max(1, Number(value.value) || 1)));
  }

  async caption() {
    const t = this.t, el = this.el;
    const label = this.select('label', [['figure', t('Docs.Caption.Figure')], ['table', t('Docs.Caption.Table')], ['equation', t('Docs.Caption.Equation')]], 'figure');
    const text = el('input', { class: 'mat-input', name: 'text', maxlength: 500 });
    const body = el('div', { class: 'mo-docs-form' }, this.field(t('Docs.Caption.Label'), label), this.field(t('Docs.Caption.Text'), text));
    const answer = await this.ask({ title: t('Docs.Caption.Insert'), icon: 'docs-caption', body, onOpen: () => text.focus() });
    if (answer !== 'ok') return;
    const word = t('Docs.Caption.' + (label.value === 'figure' ? 'Figure' : label.value === 'table' ? 'Table' : 'Equation'));
    const count = [...this.app.body.querySelectorAll('p[data-style="Caption"]')].filter(p => (p.textContent || '').startsWith(word + ' ')).length + 1;
    const line = `${word} ${count}${text.value.trim() ? ': ' + text.value.trim() : ''}`;
    this.ed.focus();
    const sel = this.ed.selection();
    if (!sel) return;
    this.ed.insertBlocks([{ type: 'p', style: 'Caption', css: {}, runs: [textRun(line)] }, { type: 'p', css: {}, runs: [] }]);
  }
}

function runsDelete(runs, start, end) {
  return end > start ? [...sliceRuns(runs, 0, start), ...sliceRuns(runs, end, Infinity)] : runs;
}

