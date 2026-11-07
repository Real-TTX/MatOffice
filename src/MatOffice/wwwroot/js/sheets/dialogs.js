/* MatOffice Sheets – the dialogs: Format Cells, Insert Function, Find and Replace, Go To, Sort, Data Validation, conditional formats,
   Remove Duplicates, Text to Columns, Name Manager, Link, Chart, Paste Special, sizes, sheets, print titles, zoom. Built with el() and the
   toolkit's dialog (never HTML from a document). */
import { el, icon, t, uid } from '../editor/util.js';
import { PALETTE } from '../editor/frame.js';
import { FUNCTIONS, CATEGORIES } from './engine/functions/index.js';
import { formatValue, presets, formatKind, decimalsOf } from './engine/numfmt.js';
import { areaName, parseArea, parseReference, colName, MAX_ROWS, MAX_COLS, cellName } from './engine/address.js';
import { compareText, numberToText, FormulaError, parseNumber } from './engine/values.js';
import { parse } from './engine/parser.js';
import { toCanonical, toLocal } from './engine/lexer.js';
import { FONTS, fontFamily, validationAt } from './look.js';
import { CHART_TYPES, CHART_COLORS } from './charts.js';
import { absoluteRefOf } from './commands.js';

// ---- small form helpers ----------------------------------------------------------------------------------------------------------

const input = (attrs = {}) => el('input', { class: 'mat-input', autocomplete: 'off', id: uid('f'), ...attrs });
const select = (options, value, attrs = {}) => el('select', { class: 'mat-input mat-select', id: uid('f'), ...attrs },
  ...options.map(o => el('option', { value: o.value, text: o.label, selected: String(o.value) === String(value ?? '') ? true : null })));
const field = (label, control, hint) => el('div', { class: 'mat-field' }, el('label', { class: 'mat-label', for: control.id || null, text: label }), control, hint ? el('div', { class: 'mat-hint', text: hint }) : null);
const check = (label, checked, attrs = {}) => {
  const box = el('input', { type: 'checkbox', checked: checked ? true : null, ...attrs });
  const node = el('label', { class: 'mat-check' }, box, el('span', { class: 'mat-check__text', text: label }));
  node.box = box;
  return node;
};
const radio = (name, value, label, checked) => {
  const r = el('input', { type: 'radio', name, value, checked: checked ? true : null });
  const node = el('label', { class: 'mat-choice' }, r, el('span', { class: 'mat-choice__text', text: label }));
  node.box = r;
  return node;
};
const row = (...items) => el('div', { class: 'mo-sheets-form__row' }, ...items);
const OK = () => ({ label: t('Sheets.Ok'), value: 'ok', kind: 'primary' });
const CANCEL = () => ({ label: t('Sheets.Cancel'), value: 'cancel' });

/** The style presets of the quick conditional formats (like Excel's list). */
export const CF_PRESETS = [
  { id: 'lightRedDark', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { id: 'yellowDark', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { id: 'greenDark', style: { fill: '#c6efce', color: '#006100' } },
  { id: 'lightRed', style: { fill: '#ffc7ce' } },
  { id: 'redText', style: { color: '#9c0006' } },
  { id: 'bold', style: { b: true } },
];

export class Dialogs {
  constructor(app) { this.app = app; }

  get sheet() { return this.app.sheet; }

  open(options) {
    this.app.keepEditing = true;
    const p = this.app.ctx.dialog(options);
    p.finally(() => { this.app.keepEditing = false; });
    return p;
  }

  /** A dialog that stays open while its buttons work (Find and Replace). buttons: [{ label, kind, run(close) }]. */
  panel({ title, iconName, body, buttons, size = 'sm' }) {
    const d = el('dialog', { class: `mat-dialog mat-dialog--${size} mo-dialog mo-sheets-panel`, role: 'dialog', 'data-static': true });
    const titleId = uid('dlg');
    d.setAttribute('aria-labelledby', titleId);
    const close = () => window.MatUi.dialog.close(d, 'close');
    const head = el('div', { class: 'mat-dialog__head' }, iconName ? icon(iconName) : null, el('h2', { class: 'mat-dialog__title', id: titleId, text: title }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-dialog__close', 'aria-label': t('Sheets.Close'), onclick: close }, icon('x')));
    const foot = el('div', { class: 'mat-btn-row mat-btn-row--plain' }, ...buttons.map(b => el('button', { type: 'button', class: `mat-btn mat-btn--${b.kind || 'secondary'}`, onclick: () => b.run(close) }, el('span', { text: b.label }))));
    const form = el('form', { class: 'mat-dialog__form', novalidate: true, onsubmit: e => { e.preventDefault(); buttons[0].run(close); } }, head, el('div', { class: 'mat-dialog__body' }, body), el('div', { class: 'mat-dialog__foot' }, foot));
    d.append(form);
    document.body.append(d);
    d.addEventListener('close', () => setTimeout(() => d.remove(), 0));
    window.MatUi.dialog.open(d);
    return { dialog: d, close };
  }

  // ---- Format Cells -----------------------------------------------------------------------------------------------------------

  async formatCells(tab = 'number') {
    const app = this.app;
    const cmd = app.commands;
    const style = cmd.style;
    const locale = app.locale;
    const p = presets(locale);
    const value = app.calc.valueAt(this.sheet.id, app.sel.active.row, app.sel.active.col);
    const sample = typeof value === 'number' ? value : 1234.5678;
    const tabs = ['number', 'alignment', 'font', 'border', 'fill', 'protection'];
    const ids = Object.fromEntries(tabs.map(k => [k, uid('tab')]));
    const tabBar = el('div', { class: 'mat-tabs', role: 'tablist', 'data-tabs': '' }, ...tabs.map(k => el('button', {
      type: 'button', class: 'mat-tab', role: 'tab', 'aria-selected': k === tab ? 'true' : 'false', 'aria-controls': ids[k], tabindex: k === tab ? '0' : '-1', text: t('Sheets.Format.Tab.' + k),
    })));
    const panel = (k, ...children) => el('div', { class: 'mat-tabpanel mo-sheets-fmt__panel', id: ids[k], role: 'tabpanel', hidden: k !== tab ? true : null }, ...children);

    // number
    const code = input({ value: style.nf || 'General', 'aria-label': t('Sheets.Format.Code') });
    const preview = el('div', { class: 'mo-sheets-fmt__preview' });
    const categories = ['general', 'number', 'currency', 'accounting', 'date', 'time', 'percent', 'fraction', 'scientific', 'text', 'custom'];
    const kindNow = formatKind(style.nf);
    const catSelect = select(categories.map(c => ({ value: c, label: t('Sheets.Nf.' + c) })), kindNow === 'datetime' ? 'date' : categories.includes(kindNow) ? kindNow : 'custom', { size: 11, class: 'mat-input mo-sheets-fmt__cats' });
    const decimals = input({ type: 'number', min: 0, max: 15, value: decimalsOf(style.nf) || 2 });
    const grouping = check(t('Sheets.Format.Grouping'), /#,##/.test(style.nf || ''));
    const symbol = select([{ value: '€', label: '€' }, { value: '$', label: '$' }, { value: '£', label: '£' }, { value: '', label: t('Sheets.Format.NoSymbol') }], /\$/.test(style.nf || '') ? '$' : /£/.test(style.nf || '') ? '£' : locale === 'de' ? '€' : '$');
    const negRed = check(t('Sheets.Format.NegativeRed'), /\[Red\]/i.test(style.nf || ''));
    const dateCodes = locale === 'de' ? ['dd.mm.yyyy', 'dd.mm.yy', 'd. mmmm yyyy', 'dddd, d. mmmm yyyy', 'mmm yy', 'mmmm yyyy', 'yyyy-mm-dd', 'dd.mm.yyyy hh:mm'] : ['m/d/yyyy', 'm/d/yy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'mmm-yy', 'mmmm yyyy', 'yyyy-mm-dd', 'm/d/yyyy h:mm'];
    const timeCodes = locale === 'de' ? ['hh:mm', 'hh:mm:ss', 'h:mm AM/PM', '[h]:mm:ss', 'mm:ss'] : ['h:mm', 'h:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', '[h]:mm:ss', 'mm:ss'];
    const fracCodes = ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'];
    const customCodes = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0;[Red]-#,##0', '#,##0.00;[Red]-#,##0.00', '0%', '0.00%', '0.00E+00', '# ?/?', '@', p.shortDate, p.longDate, p.time, '[h]:mm:ss', p.currency, p.accounting];
    const typeList = select([], '', { size: 7, class: 'mat-input mo-sheets-fmt__types' });
    const options = el('div', { class: 'mo-sheets-fmt__options' });
    const build = () => {
      const c = catSelect.value;
      const d = Math.max(0, Math.min(15, Number(decimals.value) || 0));
      const dec = d ? '.' + '0'.repeat(d) : '';
      const sym = symbol.value;
      const red = negRed.box.checked;
      switch (c) {
        case 'general': return 'General';
        case 'number': { const b = (grouping.box.checked ? '#,##0' : '0') + dec; return red ? `${b};[Red]-${b}` : b; }
        case 'currency': { const b = sym ? (sym === '€' && locale === 'de' ? `#,##0${dec} "€"` : `"${sym}"#,##0${dec}`) : `#,##0${dec}`; return red ? `${b};[Red]-${b}` : b; }
        case 'accounting': return sym === '€' && locale === 'de' ? `_-* #,##0${dec} "€"_-;-* #,##0${dec} "€"_-;_-* "-"?? "€"_-;_-@_-` : `_("${sym || ''}"* #,##0${dec}_);_("${sym || ''}"* (#,##0${dec});_("${sym || ''}"* "-"??_);_(@_)`;
        case 'percent': return '0' + dec + '%';
        case 'scientific': return '0' + (d ? '.' + '0'.repeat(d) : '') + 'E+00';
        case 'text': return '@';
        default: return typeList.value || code.value;
      }
    };
    const fillTypes = list => typeList.replaceChildren(...list.map(x => el('option', { value: x, text: x, selected: x === style.nf ? true : null })));
    const showOptions = () => {
      const c = catSelect.value;
      const parts = [];
      if (['number', 'currency', 'accounting', 'percent', 'scientific'].includes(c)) parts.push(field(t('Sheets.Format.Decimals'), decimals));
      if (c === 'number') parts.push(grouping);
      if (c === 'currency' || c === 'accounting') parts.push(field(t('Sheets.Format.Symbol'), symbol));
      if (c === 'number' || c === 'currency') parts.push(negRed);
      if (c === 'date') { fillTypes(dateCodes); parts.push(field(t('Sheets.Format.Type'), typeList)); }
      if (c === 'time') { fillTypes(timeCodes); parts.push(field(t('Sheets.Format.Type'), typeList)); }
      if (c === 'fraction') { fillTypes(fracCodes); parts.push(field(t('Sheets.Format.Type'), typeList)); }
      if (c === 'custom') { fillTypes(customCodes); parts.push(field(t('Sheets.Format.Type'), typeList)); }
      if (c === 'general') parts.push(el('p', { class: 'mo-muted', text: t('Sheets.Format.GeneralText') }));
      if (c === 'text') parts.push(el('p', { class: 'mo-muted', text: t('Sheets.Format.TextText') }));
      options.replaceChildren(...parts);
    };
    const update = fromCode => {
      if (!fromCode) code.value = build();
      const v = code.value.trim() || 'General';
      let shown;
      try { shown = formatValue(catSelect.value === 'text' && typeof value === 'string' ? value : sample, v === 'General' ? undefined : v, { locale }); } catch { shown = { text: '?' }; }
      preview.textContent = shown.text;
      preview.style.color = shown.color || '';
    };
    catSelect.addEventListener('change', () => { showOptions(); if (['date', 'time', 'fraction'].includes(catSelect.value)) typeList.selectedIndex = 0; update(); });
    for (const x of [decimals, grouping.box, symbol, negRed.box, typeList]) x.addEventListener('change', () => update());
    decimals.addEventListener('input', () => update());
    code.addEventListener('input', () => update(true));
    showOptions();
    update(true);
    const numberPanel = panel('number', el('div', { class: 'mo-sheets-fmt__number' }, field(t('Sheets.Format.Category'), catSelect),
      el('div', { class: 'mo-sheets-fmt__right' }, field(t('Sheets.Format.Sample'), preview), options, field(t('Sheets.Format.Code'), code))));

    // alignment
    const ha = select(['', 'left', 'center', 'right', 'justify'].map(v => ({ value: v, label: t('Sheets.Align.' + (v || 'general')) })), style.ha || '');
    const va = select(['top', 'middle', 'bottom'].map(v => ({ value: v, label: t('Sheets.Align.' + v) })), style.va || 'bottom');
    const indent = input({ type: 'number', min: 0, max: 15, value: style.indent || 0 });
    const wrap = check(t('Sheets.WrapText'), style.wrap);
    const merged = !!this.sheet.mergeAt(app.sel.active.row, app.sel.active.col);
    const merge = check(t('Sheets.MergeCells'), merged);
    const alignPanel = panel('alignment', row(field(t('Sheets.Align.Horizontal'), ha), field(t('Sheets.Align.Vertical'), va)), field(t('Sheets.Indent'), indent), wrap, merge);

    // font
    const font = select(FONTS.map(f => ({ value: f, label: f })), style.font || 'Calibri');
    const size = input({ type: 'number', min: 1, max: 409, step: 0.5, value: style.size || 11 });
    const bold = check(t('Sheets.Bold'), style.b);
    const italic = check(t('Sheets.Italic'), style.i);
    const underline = check(t('Sheets.Underline'), style.u);
    const strike = check(t('Sheets.Strike'), style.strike);
    const autoColor = check(t('Sheets.Automatic'), !style.color);
    const color = input({ type: 'color', value: style.color || '#000000', class: 'mo-sheets-fmt__color' });
    color.addEventListener('input', () => { autoColor.box.checked = false; fontPreview(); });
    const fontSample = el('div', { class: 'mo-sheets-fmt__fontsample', text: 'AaBbCcYyZz 123' });
    const fontPreview = () => {
      fontSample.style.fontFamily = fontFamily(font.value);
      fontSample.style.fontSize = Math.min(36, Number(size.value) || 11) + 'pt';
      fontSample.style.fontWeight = bold.box.checked ? '700' : '400';
      fontSample.style.fontStyle = italic.box.checked ? 'italic' : 'normal';
      fontSample.style.textDecoration = [underline.box.checked ? 'underline' : '', strike.box.checked ? 'line-through' : ''].join(' ').trim() || 'none';
      fontSample.style.color = autoColor.box.checked ? '' : color.value;
    };
    for (const x of [font, size, bold.box, italic.box, underline.box, strike.box, autoColor.box]) x.addEventListener('input', fontPreview);
    fontPreview();
    const fontPanel = panel('font', row(field(t('Sheets.Font'), font), field(t('Sheets.Size'), size)), el('div', { class: 'mat-checks mat-checks--inline' }, bold, italic, underline, strike),
      row(field(t('Sheets.FontColor'), color), autoColor), field(t('Sheets.Format.Sample'), fontSample));

    // border
    const edges = { bt: check(t('Sheets.Border.Top'), !!style.bt), bb: check(t('Sheets.Border.Bottom'), !!style.bb), bl: check(t('Sheets.Border.Left'), !!style.bl), br: check(t('Sheets.Border.Right'), !!style.br) };
    const inside = check(t('Sheets.Border.Inside'), false);
    const lineStyle = select(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'].map(v => ({ value: v, label: t('Sheets.Line.' + v) })), (style.bb || style.bt || 'thin').split(' ')[0]);
    const lineColor = input({ type: 'color', value: (style.bb || style.bt || 'thin #000000').split(' ')[1] || '#000000', class: 'mo-sheets-fmt__color' });
    const presetBtn = (label, ic, fn) => el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: fn }, icon(ic), el('span', { text: label }));
    const borderPanel = panel('border',
      el('div', { class: 'mat-btn-row mat-btn-row--plain' },
        presetBtn(t('Sheets.Border.None'), 'sheets-border-none', () => { for (const e of Object.values(edges)) e.box.checked = false; inside.box.checked = false; }),
        presetBtn(t('Sheets.Border.Outline'), 'sheets-border-outside', () => { for (const e of Object.values(edges)) e.box.checked = true; }),
        presetBtn(t('Sheets.Border.InsideButton'), 'sheets-border-inside', () => { inside.box.checked = true; })),
      el('div', { class: 'mat-checks mat-checks--inline' }, edges.bt, edges.bb, edges.bl, edges.br, inside),
      row(field(t('Sheets.Border.Style'), lineStyle), field(t('Sheets.Border.Color'), lineColor)));

    // fill
    let fill = style.fill || null;
    const swatch = el('div', { class: 'mo-sheets-fmt__swatch' });
    const paintSwatch = () => { swatch.style.background = fill || ''; swatch.classList.toggle('is-none', !fill); swatch.textContent = fill ? '' : t('Sheets.NoFill'); };
    paintSwatch();
    const chips = el('div', { class: 'mo-palette' }, ...PALETTE.map(line => el('div', { class: 'mo-palette__row' }, ...line.map(c => {
      const chip = el('button', { type: 'button', class: 'mo-palette__chip', 'aria-label': c, onclick: () => { fill = c; paintSwatch(); } });
      chip.style.background = c;
      return chip;
    }))));
    const more = input({ type: 'color', value: fill || '#ffffff', class: 'mo-sheets-fmt__color' });
    more.addEventListener('input', () => { fill = more.value; paintSwatch(); });
    const fillPanel = panel('fill', el('div', { class: 'mo-sheets-fmt__fill' }, chips, el('div', {}, swatch,
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => { fill = null; paintSwatch(); } }, el('span', { text: t('Sheets.NoFill') })),
      field(t('Sheets.MoreColors'), more))));

    // protection
    const locked = check(t('Sheets.Locked'), !style.unlocked);
    const protPanel = panel('protection', locked, el('p', { class: 'mo-muted', text: t('Sheets.LockedText') }));

    const body = el('div', { class: 'mo-sheets-fmt' }, tabBar, numberPanel, alignPanel, fontPanel, borderPanel, fillPanel, protPanel);
    const r = await this.open({ title: t('Sheets.FormatCells'), icon: 'sheets-format-cells', body, size: 'lg', buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const nf = code.value.trim();
    const line = `${lineStyle.value} ${lineColor.value}`;
    const fontPatch = {
      font: font.value === 'Calibri' ? undefined : font.value, size: Number(size.value) && Number(size.value) !== 11 ? Number(size.value) : undefined,
      b: bold.box.checked || undefined, i: italic.box.checked || undefined, u: underline.box.checked || undefined, strike: strike.box.checked || undefined,
      color: autoColor.box.checked ? undefined : color.value,
    };
    const ranges = app.sel.ranges.map(a => cmd.clip(a));
    cmd.change('format cells', () => {
      cmd.applyStyle(this.sheet, app.sel.ranges, s => {
        const next = {
          ...s, ...fontPatch, nf: !nf || /^general$/i.test(nf) ? undefined : nf, ha: ha.value || undefined, va: va.value === 'bottom' ? undefined : va.value,
          indent: Number(indent.value) || undefined, wrap: wrap.box.checked || undefined, fill: fill || undefined, unlocked: locked.box.checked ? undefined : true,
        };
        return next;
      });
      // borders: the edges of each range (inside lines between the cells)
      for (const a of ranges) {
        if ((a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) > 200000) continue;
        for (let r = a.r1; r <= a.r2; r++) {
          for (let c = a.c1; c <= a.c2; c++) {
            const s = { ...app.wb.style(this.sheet.styleAt(r, c)) };
            const set = (k, on) => { if (on) s[k] = line; else delete s[k]; };
            set('bt', r === a.r1 ? edges.bt.box.checked : inside.box.checked);
            set('bb', r === a.r2 ? edges.bb.box.checked : inside.box.checked);
            set('bl', c === a.c1 ? edges.bl.box.checked : inside.box.checked);
            set('br', c === a.c2 ? edges.br.box.checked : inside.box.checked);
            app.wb.patchCell(this.sheet, r, c, { s: app.wb.styleIndex(s) });
          }
        }
      }
      if (merge.box.checked !== merged) app.merge(merge.box.checked ? 'merge' : 'unmerge', true);
    });
    app.autoFitRows(app.sel.ranges);
  }

  // ---- sizes ---------------------------------------------------------------------------------------------------------------------

  async size(axis, isDefault = false) {
    const app = this.app;
    const sheet = this.sheet;
    const current = isDefault ? (axis === 'col' ? sheet.cols.def : sheet.rows.def) : axis === 'col' ? sheet.cols.rawSize(app.sel.active.col) : sheet.rows.rawSize(app.sel.active.row);
    // Excel's units: column width in characters, row height in points
    const shown = axis === 'col' ? Math.max(0, Math.round(((current - 5) / 7) * 100) / 100) : Math.round(current * 0.75 * 100) / 100;
    const box = input({ type: 'number', min: 0, max: axis === 'col' ? 255 : 409, step: axis === 'col' ? 0.01 : 0.25, value: shown });
    const label = axis === 'col' ? t(isDefault ? 'Sheets.DefaultWidth' : 'Sheets.ColWidth') : t('Sheets.RowHeight');
    const r = await this.open({ title: label, icon: axis === 'col' ? 'sheets-col-width' : 'sheets-row-height', body: field(label, box, t(axis === 'col' ? 'Sheets.ColWidth.Unit' : 'Sheets.RowHeight.Unit')), buttons: [OK(), CANCEL()], onOpen: () => { box.focus(); box.select(); } });
    if (r !== 'ok') return;
    const v = Number(String(box.value).replace(',', '.'));
    if (!Number.isFinite(v) || v < 0) return;
    const px = axis === 'col' ? (v === 0 ? 0 : Math.round(v * 7 + 5)) : Math.round(v / 0.75);
    if (isDefault) app.commands.setDefaultWidth(px); else app.commands.setSize(axis, px);
  }

  // ---- insert / delete cells -------------------------------------------------------------------------------------------------------

  async shiftCells(kind) {
    const options = kind === 'insert' ? ['right', 'down', 'rows', 'cols'] : ['left', 'up', 'rows', 'cols'];
    const radios = options.map((o, i) => radio('shift', o, t('Sheets.Shift.' + o), i === 1));
    const r = await this.open({ title: t(kind === 'insert' ? 'Sheets.InsertCells' : 'Sheets.DeleteCells'), icon: kind === 'insert' ? 'sheets-insert-cells' : 'sheets-delete-cells', body: el('div', { class: 'mat-choices' }, ...radios), buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const mode = radios.find(x => x.box.checked)?.box.value;
    if (kind === 'insert') this.app.commands.insertCells(mode); else this.app.commands.deleteCells(mode);
  }

  // ---- go to, find and replace -------------------------------------------------------------------------------------------------

  async goTo() {
    const app = this.app;
    const names = select(app.wb.names.map(n => ({ value: n.name, label: `${n.name}  (${n.ref})` })), '', { size: Math.min(8, Math.max(3, app.wb.names.length)) });
    const ref = input({ placeholder: 'A1, Sheet2!B5, A1:C10' });
    names.addEventListener('change', () => { ref.value = names.value; });
    names.addEventListener('dblclick', () => ref.form?.requestSubmit());
    const r = await this.open({ title: t('Sheets.GoTo'), icon: 'arrow-right', body: el('div', {}, app.wb.names.length ? field(t('Sheets.Names'), names) : null, field(t('Sheets.Reference'), ref)), buttons: [OK(), CANCEL()], onOpen: () => ref.focus() });
    if (r === 'ok' && ref.value.trim()) app.commands.goToText(ref.value.trim());
  }

  find(replace = false) {
    const app = this.app;
    const what = input({ value: this.lastFind || '', 'aria-label': t('Sheets.Find.What') });
    const withText = input({ value: this.lastReplace || '', 'aria-label': t('Sheets.Find.With') });
    const matchCase = check(t('Sheets.Find.MatchCase'), false);
    const whole = check(t('Sheets.Find.Whole'), false);
    const within = select([{ value: 'sheet', label: t('Sheets.Find.Sheet') }, { value: 'book', label: t('Sheets.Find.Book') }], 'sheet');
    const lookIn = select([{ value: 'formulas', label: t('Sheets.Find.Formulas') }, { value: 'values', label: t('Sheets.Find.Values') }], 'formulas');
    const results = el('div', { class: 'mo-sheets-find__results', role: 'list' });
    const status = el('div', { class: 'mo-muted mo-sheets-find__status', 'aria-live': 'polite' });
    const replaceRow = field(t('Sheets.Find.With'), withText);
    replaceRow.hidden = !replace;
    const body = el('div', { class: 'mo-sheets-find' }, field(t('Sheets.Find.What'), what), replaceRow,
      row(field(t('Sheets.Find.Within'), within), field(t('Sheets.Find.LookIn'), lookIn)), el('div', { class: 'mat-checks mat-checks--inline' }, matchCase, whole), status, results);
    const matcher = () => {
      const q = what.value;
      if (!q) return null;
      const esc = q.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]');
      return new RegExp(whole.box.checked ? '^' + esc + '$' : esc, matchCase.box.checked ? '' : 'i');
    };
    const textOf = (sheet, r, c, cell) => (lookIn.value === 'formulas' && cell.f ? '=' + toLocal(cell.f, app.locale) : app.look.cell(sheet, r, c, 11).text);
    const all = () => {
      const re = matcher();
      if (!re) return [];
      const out = [];
      const sheets = within.value === 'book' ? app.wb.sheets : [app.sheet];
      for (const sheet of sheets) {
        const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
        for (const k of keys) {
          const cell = sheet.cells.get(k);
          if (cell.v === undefined && !cell.f) continue;
          const r = Math.floor(k / 16384);
          const c = k % 16384;
          if (re.test(textOf(sheet, r, c, cell))) out.push({ sheet, r, c });
        }
      }
      return out;
    };
    const go = hit => { if (hit.sheet !== app.sheet) app.showSheet(hit.sheet); app.sel.set(hit.r, hit.c); app.selectionChanged({ scrollTo: { row: hit.r, col: hit.c } }); };
    const findNext = () => {
      this.lastFind = what.value;
      const hits = all();
      if (!hits.length) { status.textContent = t('Sheets.Find.None'); return null; }
      const cur = { sheet: app.wb.sheets.indexOf(app.sheet), k: app.sel.active.row * 16384 + app.sel.active.col };
      const order = h => app.wb.sheets.indexOf(h.sheet) * 1e12 + h.r * 16384 + h.c;
      const next = hits.find(h => order(h) > cur.sheet * 1e12 + cur.k) || hits[0];
      go(next);
      status.textContent = t('Sheets.Find.Count', hits.length);
      return next;
    };
    const findAll = () => {
      const hits = all();
      status.textContent = hits.length ? t('Sheets.Find.Count', hits.length) : t('Sheets.Find.None');
      results.replaceChildren(...hits.slice(0, 500).map(h => el('button', { type: 'button', class: 'mo-sheets-find__hit', role: 'listitem', onclick: () => go(h) },
        el('span', { class: 'mo-sheets-find__where', text: `${h.sheet.name}!${cellName(h.r, h.c)}` }), el('span', { text: app.look.cell(h.sheet, h.r, h.c, 11).text }))));
    };
    const replaceIn = (sheet, r, c, re) => {
      const cell = sheet.cell(r, c);
      if (!cell) return false;
      const flags = matchCase.box.checked ? 'g' : 'gi';
      const global = new RegExp(re.source, flags);
      if (cell.f) {
        const local = '=' + toLocal(cell.f, app.locale);
        const nextText = whole.box.checked ? withText.value : local.replace(global, () => withText.value);
        if (nextText === local) return false;
        return app.commitInput(sheet, r, c, nextText, { quiet: true });
      }
      if (typeof cell.v !== 'string' && typeof cell.v !== 'number') return false;
      const text = typeof cell.v === 'string' ? cell.v : numberToText(cell.v);
      const nextText = whole.box.checked ? withText.value : text.replace(global, () => withText.value);
      if (nextText === text) return false;
      return app.commitInput(sheet, r, c, nextText, { quiet: true });
    };
    const buttons = [
      { label: t('Sheets.Find.Next'), kind: 'primary', run: () => findNext() },
      { label: t('Sheets.Find.All'), run: () => findAll() },
    ];
    if (!app.readOnly) {
      buttons.push({ label: t('Sheets.Find.Replace'), run: () => {
        if (replaceRow.hidden) { replaceRow.hidden = false; withText.focus(); return; }
        this.lastReplace = withText.value;
        const re = matcher();
        if (!re) return;
        const { row: r, col: c } = app.sel.active;
        const cell = app.sheet.cell(r, c);
        if (cell && re.test(textOf(app.sheet, r, c, cell))) replaceIn(app.sheet, r, c, re);
        findNext();
      } }, { label: t('Sheets.Find.ReplaceAll'), run: () => {
        if (replaceRow.hidden) { replaceRow.hidden = false; withText.focus(); return; }
        this.lastReplace = withText.value;
        const re = matcher();
        if (!re) return;
        const hits = all();
        let n = 0;
        app.commands.change('replace', () => { for (const h of hits) if (replaceIn(h.sheet, h.r, h.c, re) !== false) n++; }, { protect: false });
        status.textContent = t('Sheets.Find.Replaced', n);
      } });
    }
    buttons.push({ label: t('Sheets.Close'), run: close => close() });
    this.panel({ title: t(replace ? 'Sheets.Find.ReplaceTitle' : 'Sheets.Find.Title'), iconName: 'search', body, buttons, size: 'sm' });
    setTimeout(() => { (replace ? (what.value ? withText : what) : what).focus(); what.select(); }, 30);
  }

  // ---- sort, duplicates, text to columns -------------------------------------------------------------------------------------

  async sort() {
    const app = this.app;
    const cmd = app.commands;
    const { area, header } = cmd.sortArea();
    const hasHeader = check(t('Sheets.Sort.Headers'), header);
    const columns = () => {
      const out = [];
      for (let c = area.c1; c <= area.c2; c++) {
        const name = hasHeader.box.checked ? app.look.cell(this.sheet, area.r1, c, 11).text : '';
        out.push({ value: c, label: name ? `${name}` : t('Sheets.Sort.Column', colName(c)) });
      }
      return out;
    };
    const levels = el('div', { class: 'mo-sheets-sort__levels' });
    const addLevel = (col = app.sel.active.col, desc = false) => {
      const colSel = select(columns(), col, { 'aria-label': t('Sheets.Sort.By') });
      const order = select([{ value: 'asc', label: t('Sheets.Sort.Asc') }, { value: 'desc', label: t('Sheets.Sort.Desc') }], desc ? 'desc' : 'asc', { 'aria-label': t('Sheets.Sort.Order') });
      const remove = el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Sheets.Sort.Remove'), onclick: () => { line.remove(); } }, icon('trash'));
      const line = el('div', { class: 'mo-sheets-sort__level' }, el('span', { class: 'mo-sheets-sort__label', text: levels.children.length ? t('Sheets.Sort.Then') : t('Sheets.Sort.By') }), colSel, order, remove);
      line.read = () => ({ col: Number(colSel.value), desc: order.value === 'desc' });
      levels.append(line);
    };
    addLevel();
    hasHeader.box.addEventListener('change', () => { for (const line of levels.children) { const s = line.querySelector('select'); const v = s.value; s.replaceChildren(...columns().map(o => el('option', { value: o.value, text: o.label }))); s.value = v; } });
    const body = el('div', { class: 'mo-sheets-sort' }, el('div', { class: 'mat-btn-row mat-btn-row--plain' },
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => addLevel(area.c1, false) }, icon('plus'), el('span', { text: t('Sheets.Sort.Add') })), hasHeader),
    levels, el('p', { class: 'mo-muted', text: t('Sheets.Sort.Range', areaName(area)) }));
    const r = await this.open({ title: t('Sheets.Sort.Title'), icon: 'sheets-sort', body, size: 'lg', buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const keys = [...levels.children].map(l => l.read());
    if (keys.length) cmd.sortBy(area, hasHeader.box.checked, keys);
  }

  async removeDuplicates() {
    const app = this.app;
    const cmd = app.commands;
    const { area, header } = cmd.sortArea();
    const hasHeader = check(t('Sheets.Sort.Headers'), header);
    const boxes = [];
    for (let c = area.c1; c <= area.c2; c++) {
      const name = app.look.cell(this.sheet, area.r1, c, 11).text;
      const box = check(header && name ? name : t('Sheets.Sort.Column', colName(c)), true);
      box.col = c;
      boxes.push(box);
    }
    const body = el('div', {}, el('p', { text: t('Sheets.Dup.Text', areaName(area)) }), hasHeader, el('div', { class: 'mat-checks mo-sheets-dup' }, ...boxes));
    const r = await this.open({ title: t('Sheets.Dup.Title'), icon: 'sheets-remove-dup', body, buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const cols = boxes.filter(b => b.box.checked).map(b => b.col);
    if (!cols.length) return;
    const sheet = this.sheet;
    const start = area.r1 + (hasHeader.box.checked ? 1 : 0);
    const seen = new Set();
    const keep = [];
    let removed = 0;
    for (let r2 = start; r2 <= area.r2; r2++) {
      const k = JSON.stringify(cols.map(c => { const v = app.calc.valueAt(sheet.id, r2, c); return typeof v === 'string' ? v.toLowerCase() : v; }));
      if (seen.has(k)) { removed++; continue; }
      seen.add(k);
      keep.push(r2);
    }
    if (removed) {
      cmd.change('remove duplicates', () => {
        const snap = keep.map(r2 => { const line = []; for (let c = area.c1; c <= area.c2; c++) line.push(sheet.cell(r2, c)); return { r: r2, line }; });
        for (let r2 = start; r2 <= area.r2; r2++) for (let c = area.c1; c <= area.c2; c++) app.wb.setCell(sheet, r2, c, undefined);
        snap.forEach(({ line }, i) => line.forEach((cell, j) => app.wb.setCell(sheet, start + i, area.c1 + j, cell ? { ...cell } : undefined)));
      });
    }
    app.ctx.toast(t('Sheets.Dup.Done', removed, keep.length), 'info');
  }

  async textToColumns() {
    const app = this.app;
    const a = app.commands.clip(app.sel.range);
    if (a.c1 !== a.c2) { app.ctx.toast(t('Sheets.TextCols.OneColumn'), 'warning'); return; }
    const tab = check(t('Sheets.TextCols.Tab'), true);
    const semi = check(t('Sheets.TextCols.Semicolon'), app.locale === 'de');
    const comma = check(t('Sheets.TextCols.Comma'), app.locale !== 'de');
    const space = check(t('Sheets.TextCols.Space'), false);
    const other = input({ maxlength: 1, 'aria-label': t('Sheets.TextCols.Other'), class: 'mat-input mo-sheets-short' });
    const consecutive = check(t('Sheets.TextCols.Consecutive'), false);
    const body = el('div', {}, el('p', { text: t('Sheets.TextCols.Text') }), el('div', { class: 'mat-checks' }, tab, semi, comma, space), field(t('Sheets.TextCols.Other'), other), consecutive);
    const r = await this.open({ title: t('Sheets.TextCols.Title'), icon: 'sheets-text-columns', body, buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const seps = [tab.box.checked && '\t', semi.box.checked && ';', comma.box.checked && ',', space.box.checked && ' ', other.value].filter(Boolean);
    if (!seps.length) return;
    const re = new RegExp('[' + seps.map(s => s.replace(/[\]\\^-]/g, '\\$&')).join('') + ']' + (consecutive.box.checked ? '+' : ''));
    const sheet = this.sheet;
    app.commands.change('text to columns', () => {
      for (let r2 = a.r1; r2 <= a.r2; r2++) {
        const cell = sheet.cell(r2, a.c1);
        if (!cell || typeof cell.v !== 'string' || cell.f) continue;
        cell.v.split(re).forEach((part, j) => {
          const parsed = app.parseTyped(part);
          app.wb.patchCell(sheet, r2, a.c1 + j, parsed);
        });
      }
    });
  }

  // ---- data validation ---------------------------------------------------------------------------------------------------------

  async validation() {
    const app = this.app;
    const sheet = this.sheet;
    const { row, col } = app.sel.active;
    const current = validationAt(sheet, row, col) || {};
    const types = ['any', 'whole', 'decimal', 'list', 'date', 'time', 'length', 'custom'];
    const type = select(types.map(v => ({ value: v, label: t('Sheets.Valid.' + v) })), current.type || 'any');
    const ops = ['between', 'notBetween', 'eq', 'ne', 'gt', 'lt', 'ge', 'le'];
    const op = select(ops.map(v => ({ value: v, label: t('Sheets.Op.' + v) })), current.op || 'between');
    const v1 = input({ value: current.v1 ?? '' });
    const v2 = input({ value: current.v2 ?? '' });
    const blank = check(t('Sheets.Valid.Blank'), current.blank !== false);
    const dropdown = check(t('Sheets.Valid.Dropdown'), current.dropdown !== false);
    const prompt = el('textarea', { class: 'mat-input', rows: 2, id: uid('f') });
    prompt.value = current.prompt || '';
    const errorStyle = select(['stop', 'warning', 'info'].map(v => ({ value: v, label: t('Sheets.Valid.Style.' + v) })), current.errorStyle || 'stop');
    const errorTitle = input({ value: current.errorTitle || '', maxlength: 32 });
    const error = el('textarea', { class: 'mat-input', rows: 2, id: uid('f'), maxlength: 225 });
    error.value = current.error || '';
    const f1 = field(t('Sheets.Valid.Minimum'), v1);
    const f2 = field(t('Sheets.Valid.Maximum'), v2);
    const fop = field(t('Sheets.Valid.Data'), op);
    const update = () => {
      const tp = type.value;
      const range = op.value === 'between' || op.value === 'notBetween';
      fop.hidden = ['any', 'list', 'custom'].includes(tp);
      f1.hidden = tp === 'any';
      f2.hidden = fop.hidden || !range;
      f1.querySelector('label').textContent = tp === 'list' ? t('Sheets.Valid.Source') : tp === 'custom' ? t('Sheets.Valid.Formula') : range ? t('Sheets.Valid.Minimum') : t('Sheets.Valid.Value');
      dropdown.hidden = tp !== 'list';
    };
    type.addEventListener('change', update);
    op.addEventListener('change', update);
    update();
    const body = el('div', { class: 'mo-sheets-form' }, row(field(t('Sheets.Valid.Allow'), type), fop), row(f1, f2), el('div', { class: 'mat-checks' }, blank, dropdown),
      el('h3', { class: 'mo-sheets-form__head', text: t('Sheets.Valid.Input') }), field(t('Sheets.Valid.Prompt'), prompt),
      el('h3', { class: 'mo-sheets-form__head', text: t('Sheets.Valid.Error') }), row(field(t('Sheets.Valid.ErrorStyle'), errorStyle), field(t('Sheets.Valid.ErrorTitle'), errorTitle)), field(t('Sheets.Valid.ErrorText'), error));
    const r = await this.open({ title: t('Sheets.Validation'), icon: 'sheets-validation', body, size: 'lg', buttons: [OK(), { label: t('Sheets.Valid.Clear'), value: 'clear' }, CANCEL()] });
    if (r !== 'ok' && r !== 'clear') return;
    const area = app.commands.clip(app.sel.range);
    const range = areaName(app.sel.range.r2 === MAX_ROWS - 1 ? { ...app.sel.range } : area);
    const others = sheet.validations.filter(v => { const x = parseArea(v.range); return !(x && x.r1 >= area.r1 && x.r2 <= area.r2 && x.c1 >= area.c1 && x.c2 <= area.c2); });
    let rule = null;
    if (r === 'ok' && type.value !== 'any') {
      const canon = text => (String(text).startsWith('=') ? '=' + toCanonical(String(text).slice(1), app.locale) : String(text).trim());
      rule = { range, type: type.value, v1: canon(v1.value) };
      if (!['list', 'custom'].includes(type.value)) { rule.op = op.value; if (op.value === 'between' || op.value === 'notBetween') rule.v2 = canon(v2.value); }
      if (!rule.v1) { app.ctx.toast(t('Sheets.Valid.Missing'), 'danger'); return; }
      rule.blank = blank.box.checked;
      if (type.value === 'list') rule.dropdown = dropdown.box.checked;
      if (prompt.value.trim()) rule.prompt = prompt.value.trim();
      if (error.value.trim()) rule.error = error.value.trim();
      if (errorTitle.value.trim()) rule.errorTitle = errorTitle.value.trim();
      rule.errorStyle = errorStyle.value;
    }
    app.commands.change('validation', () => app.wb.setSheetProp(sheet, 'validations', rule ? [...others, rule] : others), { protect: false });
  }

  // ---- conditional formats -----------------------------------------------------------------------------------------------------

  /** A quick rule: kind gt lt between eq text dup unique top bottom topPercent bottomPercent above below. */
  async cfRule(kind) {
    const app = this.app;
    const presetSel = select(CF_PRESETS.map(p => ({ value: p.id, label: t('Sheets.Cf.Preset.' + p.id) })), CF_PRESETS[0].id);
    const v1 = input({ value: '' });
    const v2 = input({ value: '' });
    const n = input({ type: 'number', min: 1, max: 1000, value: 10 });
    let body;
    const prefix = el('p', { text: t('Sheets.Cf.Prompt.' + kind) });
    if (['gt', 'lt', 'eq', 'text'].includes(kind)) body = el('div', {}, prefix, row(v1, field(t('Sheets.Cf.With'), presetSel)));
    else if (kind === 'between') body = el('div', {}, prefix, row(v1, el('span', { text: t('Sheets.Cf.And') }), v2), field(t('Sheets.Cf.With'), presetSel));
    else if (['top', 'bottom', 'topPercent', 'bottomPercent'].includes(kind)) body = el('div', {}, prefix, row(n, field(t('Sheets.Cf.With'), presetSel)));
    else body = el('div', {}, prefix, field(t('Sheets.Cf.With'), presetSel));
    // a sensible proposal: the average of the selection
    const nums = app.look.numbersIn(this.sheet, app.commands.clip(app.sel.range));
    if (nums.length) { const avg = nums.reduce((a, b) => a + b, 0) / nums.length; v1.value = numberToText(Math.round(avg * 100) / 100); v2.value = numberToText(Math.max(...nums)); }
    const r = await this.open({ title: t('Sheets.Cf.Title.' + kind), icon: 'sheets-cf', body, buttons: [OK(), CANCEL()], onOpen: () => (v1.isConnected ? v1 : presetSel).focus() });
    if (r !== 'ok') return;
    const style = CF_PRESETS.find(p => p.id === presetSel.value).style;
    const range = areaName(app.commands.clip(app.sel.range));
    const value = v => { const s = String(v).trim(); return s.startsWith('=') ? '=' + toCanonical(s.slice(1), app.locale) : (parseNumber(s) !== null ? numberToText(parseNumber(s)).replace(',', '.') : s); };
    let rule;
    switch (kind) {
      case 'gt': case 'lt': case 'eq': rule = { range, type: 'cell', op: kind, v1: value(v1.value), style }; break;
      case 'between': rule = { range, type: 'cell', op: 'between', v1: value(v1.value), v2: value(v2.value), style }; break;
      case 'text': rule = { range, type: 'text', op: 'contains', v1: v1.value, style }; break;
      case 'dup': case 'unique': rule = { range, type: kind, style }; break;
      case 'top': case 'bottom': rule = { range, type: kind, n: Number(n.value) || 10, style }; break;
      case 'topPercent': case 'bottomPercent': rule = { range, type: kind === 'topPercent' ? 'top' : 'bottom', n: Number(n.value) || 10, percent: true, style }; break;
      case 'above': case 'below': rule = { range, type: kind, style }; break;
      default: return;
    }
    if (rule.type === 'cell' && !rule.v1) return;
    this.addRule(rule);
  }

  addRule(rule) {
    const app = this.app;
    const sheet = this.sheet;
    app.commands.change('conditional format', () => app.wb.setSheetProp(sheet, 'cf', [...sheet.cf, rule]), { protect: false });
    app.look.invalidate();
    app.grid.requestRender();
  }

  /** New rule with a formula (and a format of its own). */
  async cfFormula(existing) {
    const app = this.app;
    const formula = input({ value: existing ? '=' + toLocal(String(existing.v1).replace(/^=/, ''), app.locale) : '=' });
    const fill = input({ type: 'color', value: existing?.style?.fill || '#ffc7ce', class: 'mo-sheets-fmt__color' });
    const useFill = check(t('Sheets.Fill'), existing ? !!existing.style?.fill : true);
    const color = input({ type: 'color', value: existing?.style?.color || '#9c0006', class: 'mo-sheets-fmt__color' });
    const useColor = check(t('Sheets.FontColor'), existing ? !!existing.style?.color : true);
    const bold = check(t('Sheets.Bold'), !!existing?.style?.b);
    const body = el('div', { class: 'mo-sheets-form' }, el('p', { text: t('Sheets.Cf.FormulaText') }), field(t('Sheets.Cf.Formula'), formula),
      row(useFill, fill), row(useColor, color), bold);
    const r = await this.open({ title: t('Sheets.Cf.NewRule'), icon: 'sheets-cf', body, buttons: [OK(), CANCEL()], onOpen: () => formula.focus() });
    if (r !== 'ok') return;
    const text = formula.value.trim().replace(/^=/, '');
    if (!text) return;
    let canon;
    try { canon = toCanonical(text, app.locale); parse(canon); } catch { app.ctx.toast(t('Sheets.Formula.Invalid'), 'danger'); return; }
    const style = {};
    if (useFill.box.checked) style.fill = fill.value;
    if (useColor.box.checked) style.color = color.value;
    if (bold.box.checked) style.b = true;
    const rule = { range: existing?.range || areaName(app.commands.clip(app.sel.range)), type: 'formula', v1: canon, style };
    if (existing) {
      const sheet = this.sheet;
      app.commands.change('conditional format', () => app.wb.setSheetProp(sheet, 'cf', sheet.cf.map(x => (x === existing ? rule : x))), { protect: false });
      app.look.invalidate();
    } else this.addRule(rule);
  }

  bar(color) { this.addRule({ range: areaName(this.app.commands.clip(this.app.sel.range)), type: 'bar', color }); }
  scale(colors) { this.addRule({ range: areaName(this.app.commands.clip(this.app.sel.range)), type: 'scale', colors }); }

  clearRules(all) {
    const app = this.app;
    const sheet = this.sheet;
    const a = app.commands.clip(app.sel.range);
    const keep = all ? [] : sheet.cf.filter(r => { const x = parseArea(r.range); return !(x && x.r1 <= a.r2 && a.r1 <= x.r2 && x.c1 <= a.c2 && a.c1 <= x.c2); });
    app.commands.change('clear rules', () => app.wb.setSheetProp(sheet, 'cf', keep), { protect: false });
    app.look.invalidate();
  }

  async manageRules() {
    const app = this.app;
    const sheet = this.sheet;
    let rules = sheet.cf.map(r => ({ ...r }));
    const list = el('div', { class: 'mo-sheets-rules' });
    const describe = r => {
      switch (r.type) {
        case 'cell': return t('Sheets.Cf.Desc.cell', t('Sheets.Op.' + r.op), r.v1 + (r.v2 !== undefined ? ' / ' + r.v2 : ''));
        case 'text': return t('Sheets.Cf.Desc.text', r.v1);
        case 'top': case 'bottom': return t('Sheets.Cf.Desc.' + r.type, r.n + (r.percent ? ' %' : ''));
        case 'formula': return '=' + toLocal(String(r.v1).replace(/^=/, ''), app.locale);
        default: return t('Sheets.Cf.Desc.' + r.type);
      }
    };
    const paint = () => {
      list.replaceChildren(...(rules.length ? rules.map((r, i) => {
        const range = input({ value: r.range, class: 'mat-input mo-sheets-rules__range', 'aria-label': t('Sheets.Cf.AppliesTo') });
        range.addEventListener('change', () => { const a = parseArea(range.value); if (a) r.range = areaName(a); else range.value = r.range; });
        const sample = el('span', { class: 'mo-sheets-rules__sample', text: 'AaBbCc' });
        if (r.style?.fill) sample.style.background = r.style.fill;
        if (r.style?.color) sample.style.color = r.style.color;
        if (r.style?.b) sample.style.fontWeight = '700';
        if (r.type === 'bar') sample.style.background = `linear-gradient(90deg, ${r.color} 60%, transparent 60%)`;
        if (r.type === 'scale') sample.style.background = `linear-gradient(90deg, ${r.colors.join(', ')})`;
        const btn = (ic, label, fn, disabled) => el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': label, 'data-tip': label, disabled: disabled ? true : null, onclick: fn }, icon(ic));
        return el('div', { class: 'mo-sheets-rules__row' }, el('span', { class: 'mo-sheets-rules__desc', text: describe(r) }), sample, range,
          btn('arrow-up', t('Sheets.Cf.Up'), () => { [rules[i - 1], rules[i]] = [rules[i], rules[i - 1]]; paint(); }, i === 0),
          btn('arrow-down', t('Sheets.Cf.Down'), () => { [rules[i + 1], rules[i]] = [rules[i], rules[i + 1]]; paint(); }, i === rules.length - 1),
          btn('trash', t('Sheets.Delete'), () => { rules.splice(i, 1); paint(); }));
      }) : [el('p', { class: 'mo-muted', text: t('Sheets.Cf.None') })]));
    };
    paint();
    const r = await this.open({ title: t('Sheets.Cf.Manage'), icon: 'sheets-cf', body: list, size: 'lg', buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    app.commands.change('conditional formats', () => app.wb.setSheetProp(sheet, 'cf', rules), { protect: false });
    app.look.invalidate();
    rules = null;
  }

  // ---- names --------------------------------------------------------------------------------------------------------------------

  async defineName(existing) {
    const app = this.app;
    const name = input({ value: existing?.name || '', maxlength: 255 });
    const ref = input({ value: existing ? '=' + toLocal(existing.ref, app.locale) : '=' + absoluteRefOf(app.sheet, app.commands.clip(app.sel.range)) });
    const r = await this.open({ title: t(existing ? 'Sheets.Name.Edit' : 'Sheets.DefineName'), icon: 'tag', body: el('div', {}, field(t('Sheets.Name.Name'), name), field(t('Sheets.Name.RefersTo'), ref)), buttons: [OK(), CANCEL()], onOpen: () => name.focus() });
    if (r !== 'ok') return false;
    const n = name.value.trim();
    if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^(R|C|R\d+C\d+|TRUE|FALSE)$/i.test(n)) { app.ctx.toast(t('Sheets.Name.Invalid'), 'danger'); return false; }
    if (app.wb.names.some(x => x !== existing && x.name.toLowerCase() === n.toLowerCase())) { app.ctx.toast(t('Sheets.Name.Taken'), 'danger'); return false; }
    let canon;
    try { canon = toCanonical(ref.value.trim().replace(/^=/, ''), app.locale); parse(canon); } catch { app.ctx.toast(t('Sheets.Formula.Invalid'), 'danger'); return false; }
    const names = app.wb.names.filter(x => x !== existing);
    names.push({ name: n, ref: canon });
    names.sort((a, b) => compareText(a.name, b.name));
    app.commands.change('name', () => app.wb.setNames(names), { protect: false, structural: true });
    return true;
  }

  async nameManager() {
    const app = this.app;
    const list = el('div', { class: 'mo-sheets-names' });
    const paint = () => {
      list.replaceChildren(...(app.wb.names.length ? app.wb.names.map(n => el('div', { class: 'mo-sheets-names__row' },
        el('span', { class: 'mo-sheets-names__name', text: n.name }), el('span', { class: 'mo-sheets-names__ref', text: '=' + toLocal(n.ref, app.locale) }),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Sheets.Name.Edit'), disabled: app.readOnly ? true : null, onclick: async () => { await this.defineName(n); paint(); } }, icon('edit')),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Sheets.Delete'), disabled: app.readOnly ? true : null, onclick: () => { app.commands.change('name', () => app.wb.setNames(app.wb.names.filter(x => x !== n)), { protect: false, structural: true }); paint(); } }, icon('trash'))))
        : [el('p', { class: 'mo-muted', text: t('Sheets.Names.None') })]));
    };
    paint();
    const addBtn = el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', disabled: app.readOnly ? true : null, onclick: async () => { await this.defineName(); paint(); } }, icon('plus'), el('span', { text: t('Sheets.Name.New') }));
    await this.open({ title: t('Sheets.NameManager'), icon: 'sheets-names', body: el('div', {}, el('div', { class: 'mat-btn-row mat-btn-row--plain' }, addBtn), list), size: 'lg', buttons: [{ label: t('Sheets.Close'), value: 'ok', kind: 'primary' }] });
  }

  // ---- insert function -----------------------------------------------------------------------------------------------------------

  async insertFunction() {
    const app = this.app;
    const search = input({ type: 'search', placeholder: t('Sheets.FnDlg.Search'), 'aria-label': t('Sheets.FnDlg.Search') });
    const cats = select([{ value: 'all', label: t('Sheets.Cat.all') }, { value: 'recent', label: t('Sheets.Cat.recent') }, ...CATEGORIES.map(c => ({ value: c, label: t('Sheets.Cat.' + c) }))], this.lastCategory || 'recent');
    const list = select([], '', { size: 10, class: 'mat-input mo-sheets-fn__list', 'aria-label': t('Sheets.FnDlg.List') });
    const syntax = el('div', { class: 'mo-sheets-fn__syntax' });
    const desc = el('div', { class: 'mo-sheets-fn__desc' });
    const recent = () => (this.recent?.length ? this.recent : ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'XLOOKUP', 'ROUND', 'TODAY']).filter(n => FUNCTIONS.has(n));
    const fill = () => {
      const q = search.value.trim().toUpperCase();
      let names = [...FUNCTIONS.values()].filter(f => !q || f.name.includes(q) || t('Sheets.Fn.' + f.name).toUpperCase().includes(q)).map(f => f.name);
      if (!q && cats.value === 'recent') names = recent();
      else if (!q && cats.value !== 'all') names = names.filter(n => FUNCTIONS.get(n).category === cats.value);
      names.sort();
      list.replaceChildren(...names.map(n => el('option', { value: n, text: n })));
      if (names.length) list.selectedIndex = 0;
      show();
    };
    const show = () => {
      const f = FUNCTIONS.get(list.value);
      syntax.textContent = f ? `${f.name}(${(f.args || '').split(/,\s*/).join(app.locale === 'de' ? '; ' : ', ')})` : '';
      desc.textContent = f ? t('Sheets.Fn.' + f.name) : '';
    };
    search.addEventListener('input', fill);
    cats.addEventListener('change', () => { search.value = ''; this.lastCategory = cats.value; fill(); });
    list.addEventListener('change', show);
    list.addEventListener('dblclick', () => list.form?.requestSubmit());
    fill();
    const body = el('div', { class: 'mo-sheets-fn' }, field(t('Sheets.FnDlg.Search'), search), field(t('Sheets.FnDlg.Category'), cats), list, syntax, desc);
    const r = await this.open({ title: t('Sheets.InsertFunction'), icon: 'sheets-function', body, size: 'lg', buttons: [OK(), CANCEL()], onOpen: () => search.focus() });
    if (r !== 'ok' || !list.value) return;
    const name = list.value;
    this.recent = [name, ...(this.recent || []).filter(n => n !== name)].slice(0, 10);
    const editor = app.editor;
    if (editor.active) {
      const input2 = editor.input;
      editor.insertText(input2, (input2.value ? '' : '=') + name + '()');
      input2.setSelectionRange(input2.selectionStart - 1, input2.selectionStart - 1);
      editor.afterCaret(input2);
      input2.focus();
    } else if (!app.readOnly) {
      editor.begin({ text: '=' + name + '()', mode: 'edit' });
      const i2 = editor.input;
      i2.setSelectionRange(i2.value.length - 1, i2.value.length - 1);
      editor.afterCaret(i2);
    }
  }

  // ---- links, pictures, charts -------------------------------------------------------------------------------------------------

  async link() {
    const app = this.app;
    const sheet = this.sheet;
    const { row, col } = app.sel.active;
    const cell = sheet.cell(row, col) || {};
    const current = cell.link || '';
    const inside = current.startsWith('#');
    const text = input({ value: typeof cell.v === 'string' ? cell.v : cell.v !== undefined ? app.look.cell(sheet, row, col, 11).text : '' });
    const address = input({ value: inside ? '' : current, placeholder: 'https://', type: 'url' });
    const target = select(app.wb.sheets.map(s => ({ value: s.name, label: s.name })), inside ? (parseReference(current.slice(1))?.sheet || sheet.name) : sheet.name);
    const cellRef = input({ value: inside ? areaName(parseReference(current.slice(1))?.area || { r1: 0, c1: 0, r2: 0, c2: 0 }) : 'A1' });
    const web = radio('kind', 'web', t('Sheets.Link.Web'), !inside);
    const place = radio('kind', 'place', t('Sheets.Link.Place'), inside);
    const webBox = el('div', {}, field(t('Sheets.Link.Address'), address));
    const placeBox = el('div', {}, row(field(t('Sheets.Link.Sheet'), target), field(t('Sheets.Link.Cell'), cellRef)));
    const sync = () => { webBox.hidden = !web.box.checked; placeBox.hidden = web.box.checked; };
    web.box.addEventListener('change', sync);
    place.box.addEventListener('change', sync);
    sync();
    const body = el('div', { class: 'mo-sheets-form' }, field(t('Sheets.Link.Text'), text), el('div', { class: 'mat-choices mat-checks--inline' }, web, place), webBox, placeBox);
    const r = await this.open({ title: t(current ? 'Sheets.EditLink' : 'Sheets.Link'), icon: 'link', body, buttons: [OK(), current ? { label: t('Sheets.Link.Remove'), value: 'remove' } : null, CANCEL()].filter(Boolean), onOpen: () => (current || !text.value ? text : address).focus() });
    if (r === 'remove') { app.commands.change('link', () => app.wb.patchCell(sheet, row, col, { link: undefined })); return; }
    if (r !== 'ok') return;
    let link;
    if (web.box.checked) {
      link = address.value.trim();
      if (link && !/^(https?:\/\/|mailto:)/i.test(link)) link = (/@/.test(link) && !/\//.test(link) ? 'mailto:' : 'https://') + link;
      if (link && !/^(https?:\/\/[^\s]+|mailto:[^\s]+)$/i.test(link)) { app.ctx.toast(t('Sheets.Link.Invalid'), 'danger'); return; }
    } else {
      const a = parseArea(cellRef.value);
      if (!a) { app.ctx.toast(t('Sheets.Goto.Invalid'), 'danger'); return; }
      link = '#' + absoluteRefOf(app.wb.sheetByName(target.value) || sheet, a).replace(/\$/g, '');
    }
    if (!link) return;
    app.commands.change('link', () => {
      const patch = { link };
      if (text.value.trim() && !cell.f) { patch.v = text.value; patch.t = undefined; } else if (cell.v === undefined && !cell.f) patch.v = text.value || link.replace(/^mailto:/, '');
      app.wb.patchCell(sheet, row, col, patch);
    });
  }

  async altText(image) {
    const app = this.app;
    const text = el('textarea', { class: 'mat-input', rows: 3, id: uid('f') });
    text.value = image.alt || '';
    const r = await this.open({ title: t('Sheets.Picture.Alt'), icon: 'image', body: field(t('Sheets.Picture.AltText'), text), buttons: [OK(), CANCEL()], onOpen: () => text.focus() });
    if (r !== 'ok') return;
    app.commands.change('alt text', () => { app.wb.touchProp(this.sheet, 'images'); const live = this.sheet.images.find(i => i.id === image.id); if (live) { if (text.value.trim()) live.alt = text.value.trim(); else delete live.alt; } }, { protect: false });
    app.objects.refresh();
  }

  async chart(chart) {
    const app = this.app;
    if (!chart) return;
    const type = select(CHART_TYPES.map(v => ({ value: v, label: t('Sheets.Chart.' + v) })), chart.type);
    const title = input({ value: chart.title || '' });
    const range = input({ value: chart.range });
    const legend = select(['', 'right', 'bottom', 'top', 'left', 'none'].map(v => ({ value: v, label: t('Sheets.Legend.' + (v || 'auto')) })), chart.legend || '');
    const xTitle = input({ value: chart.xTitle || '' });
    const yTitle = input({ value: chart.yTitle || '' });
    const stacked = check(t('Sheets.Chart.Stacked'), chart.stacked);
    const labels = check(t('Sheets.Chart.Labels'), chart.labels);
    const byRows = check(t('Sheets.Chart.ByRows'), chart.byRows);
    const colors = (chart.colors?.length ? chart.colors : CHART_COLORS).slice(0, 6).map(c => input({ type: 'color', value: c, class: 'mo-sheets-fmt__color' }));
    const body = el('div', { class: 'mo-sheets-form' }, row(field(t('Sheets.Chart.Type'), type), field(t('Sheets.Chart.Range'), range)), field(t('Sheets.Chart.Title'), title),
      row(field(t('Sheets.Chart.Legend'), legend), field(t('Sheets.Chart.XTitle'), xTitle), field(t('Sheets.Chart.YTitle'), yTitle)),
      el('div', { class: 'mat-checks mat-checks--inline' }, stacked, labels, byRows), field(t('Sheets.Chart.Colors'), el('div', { class: 'mo-sheets-colors' }, ...colors)));
    const r = await this.open({ title: t('Sheets.Chart.Edit'), icon: 'sheets-chart-column', body, size: 'lg', buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const ref = parseReference(range.value.trim());
    if (!ref) { app.ctx.toast(t('Sheets.Goto.Invalid'), 'danger'); return; }
    app.objects.updateChart({
      type: type.value, title: title.value.trim() || undefined, range: range.value.trim(), legend: legend.value || undefined, xTitle: xTitle.value.trim() || undefined,
      yTitle: yTitle.value.trim() || undefined, stacked: stacked.box.checked || undefined, labels: labels.box.checked || undefined, byRows: byRows.box.checked || undefined,
      colors: colors.map(c => c.value),
    }, chart);
  }

  // ---- paste special, sheets, page -----------------------------------------------------------------------------------------

  async pasteSpecial() {
    const app = this.app;
    const data = app.clipboard.data;
    if (!data) { app.ctx.toast(t('Sheets.Clipboard.Empty'), 'info'); return; }
    const modes = ['all', 'formulas', 'values', 'formats', 'widths'];
    const radios = modes.map((m, i) => radio('paste', m, t('Sheets.PasteMode.' + m), i === 0));
    const skip = check(t('Sheets.PasteMode.SkipBlanks'), false);
    const transpose = check(t('Sheets.PasteMode.Transpose'), false);
    const r = await this.open({ title: t('Sheets.PasteSpecial'), icon: 'sheets-paste-special', body: el('div', {}, el('div', { class: 'mat-choices' }, ...radios), el('div', { class: 'mat-checks mat-checks--inline' }, skip, transpose)), buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const mode = radios.find(x => x.box.checked).box.value;
    app.clipboard.pasteInternal({ ...data, cut: false }, mode === 'all' ? (transpose.box.checked ? 'transpose' : undefined) : mode, { transpose: transpose.box.checked, skipBlanks: skip.box.checked });
  }

  async unhideSheet() {
    const app = this.app;
    const hidden = app.wb.sheets.filter(s => s.hidden);
    if (!hidden.length) return;
    const list = select(hidden.map(s => ({ value: s.id, label: s.name })), hidden[0].id, { size: Math.min(8, hidden.length + 1) });
    const r = await this.open({ title: t('Sheets.Sheet.Unhide'), icon: 'eye', body: field(t('Sheets.Sheet.UnhideText'), list), buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const sheet = app.wb.sheetById(list.value);
    if (sheet) app.commands.unhideSheet(sheet);
  }

  async printTitles() {
    const app = this.app;
    const rows = input({ value: this.sheet.print?.titles || '', placeholder: '1:1' });
    const r = await this.open({ title: t('Sheets.PrintTitles'), icon: 'sheets-print-titles', body: field(t('Sheets.PrintTitles.Rows'), rows, t('Sheets.PrintTitles.Hint')), buttons: [OK(), CANCEL()], onOpen: () => rows.focus() });
    if (r !== 'ok') return;
    const v = rows.value.trim().replace(/\$/g, '');
    if (!v) { app.commands.setPrint({ titles: undefined }); return; }
    const m = /^(\d+)(?::(\d+))?$/.exec(v);
    if (!m) { app.ctx.toast(t('Sheets.Goto.Invalid'), 'danger'); return; }
    const a = Number(m[1]);
    const b = Number(m[2] || m[1]);
    app.commands.setPrint({ titles: `${Math.min(a, b)}:${Math.max(a, b)}` });
  }

  async zoom() {
    const app = this.app;
    const z = Math.round(this.sheet.zoom * 100);
    const values = [200, 100, 75, 50, 25];
    const radios = values.map(v => radio('zoom', String(v), v + ' %', v === z));
    const fit = radio('zoom', 'fit', t('Sheets.ZoomSelection'), false);
    const custom = radio('zoom', 'custom', t('Sheets.ZoomCustom'), !values.includes(z));
    const box = input({ type: 'number', min: 10, max: 400, value: z, class: 'mat-input mo-sheets-short' });
    box.addEventListener('focus', () => { custom.box.checked = true; });
    const r = await this.open({ title: t('Sheets.Zoom'), icon: 'zoom-in', body: el('div', { class: 'mat-choices' }, ...radios, fit, el('div', { class: 'mo-sheets-form__row' }, custom, box)), buttons: [OK(), CANCEL()] });
    if (r !== 'ok') return;
    const v = [...radios, fit, custom].find(x => x.box.checked)?.box.value;
    if (v === 'fit') app.commands.zoomToSelection();
    else if (v === 'custom') app.commands.setZoom((Number(box.value) || 100) / 100);
    else if (v) app.commands.setZoom(Number(v) / 100);
  }
}

export { MAX_COLS, FormulaError };
