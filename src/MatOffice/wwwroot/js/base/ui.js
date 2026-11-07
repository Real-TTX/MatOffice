/* MatOffice Base – small helpers of the editor: the texts of problems (codes of the pure modules → words of the catalog), form fields for dialogs,
   context menus, dragging, the language of the page. */
import { t, el, icon, dialog, lang as pageLang } from '../editor/util.js';
import { openPopover, closePopover } from '../editor/frame.js';
import { BaseError } from './expr.js';

export const lang = () => (pageLang() === 'de' ? 'de' : 'en');

/** The catalog keys of the problem codes (written out so the catalog test sees every key). */
const PROBLEMS = {
  Syntax: 'Base.Err.Syntax', Unsupported: 'Base.Err.Unsupported', Empty: 'Base.Err.Empty', BadDate: 'Base.Err.BadDate', TypeMismatch: 'Base.Err.TypeMismatch',
  DivZero: 'Base.Err.DivZero', UnknownName: 'Base.Err.UnknownName', UnknownFunction: 'Base.Err.UnknownFunction', ArgCount: 'Base.Err.ArgCount',
  BadArgument: 'Base.Err.BadArgument', ParamMissing: 'Base.Err.ParamMissing', AggregateHere: 'Base.Err.AggregateHere', NestedAggregate: 'Base.Err.NestedAggregate',
  Ambiguous: 'Base.Err.Ambiguous', NoTable: 'Base.Err.NoTable', Circular: 'Base.Err.Circular', NotSelect: 'Base.Err.NotSelect', TooLarge: 'Base.Err.TooLarge',
  NotInGroup: 'Base.Err.NotInGroup', AggregateInWhere: 'Base.Err.AggregateInWhere', WhichTable: 'Base.Err.WhichTable', NameTaken: 'Base.Err.NameTaken',
  NotAction: 'Base.Err.NotAction', NoColumns: 'Base.Err.NoColumns', NoTables: 'Base.Err.NoTables', NoTarget: 'Base.Err.NoTarget', NoUpdate: 'Base.Err.NoUpdate',
  UpdateField: 'Base.Err.UpdateField', NameEmpty: 'Base.Err.NameEmpty', NameLong: 'Base.Err.NameLong', NameChars: 'Base.Err.NameChars', PkNull: 'Base.Err.PkNull',
  Required: 'Base.Err.Required', Validation: 'Base.Err.Validation', Duplicate: 'Base.Err.Duplicate', NoParent: 'Base.Err.NoParent', HasChildren: 'Base.Err.HasChildren',
  RecordGone: 'Base.Err.RecordGone', TooManyRecords: 'Base.Err.TooManyRecords', TooLong: 'Base.Err.TooLong', OutOfRange: 'Base.Err.OutOfRange',
  RelSame: 'Base.Err.RelSame', RelTypes: 'Base.Err.RelTypes', RelExists: 'Base.Err.RelExists', RelNoUnique: 'Base.Err.RelNoUnique', RelViolated: 'Base.Err.RelViolated',
  RelMissing: 'Base.Err.RelMissing', FieldInRelationship: 'Base.Err.FieldInRelationship', NoAutoNumberWithData: 'Base.Err.NoAutoNumberWithData',
  OneAutoNumber: 'Base.Err.OneAutoNumber', RequiredViolated: 'Base.Err.RequiredViolated', DuplicatesExist: 'Base.Err.DuplicatesExist', RuleName: 'Base.Err.RuleName',
  PkMemo: 'Base.Err.PkMemo', PkNulls: 'Base.Err.PkNulls', PkDuplicates: 'Base.Err.PkDuplicates', PkInRelationship: 'Base.Err.PkInRelationship',
  TableInRelationship: 'Base.Err.TableInRelationship', NotUpdatable: 'Base.Err.NotUpdatable', NoGrid: 'Base.Err.NoGrid',
};

/** The words for a problem: a BaseError, { code, args }, or anything else (a general failure). */
export function problemText(p) {
  if (!p) return t('Base.Err.General');
  const code = p.code;
  const args = p.args || [];
  if (code === 'Validation' && args[0]) return String(args[0]);
  const key = PROBLEMS[code];
  if (!key) return p instanceof Error && !(p instanceof BaseError) ? t('Base.Err.General') : t('Base.Err.General');
  return t(key, ...args.map(a => String(a ?? '')));
}

/** Shows a problem in the Access way: a dialog with a warning sign and OK. */
export function showProblem(p, title) {
  return dialog({ title: title || t('Base.AppName'), icon: 'alert', body: el('p', { class: 'mo-base-problem', text: problemText(p) }), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }] });
}

/** Asks yes/no in a dialog (danger: the positive button is red). */
export async function ask(message, { title, ok, cancel, danger = false, icon: ic } = {}) {
  const r = await dialog({
    title: title || t('Base.AppName'), icon: ic || (danger ? 'alert' : 'help-circle'), body: el('p', { text: message }),
    buttons: danger
      ? [{ label: cancel || t('Editor.Cancel'), value: 'cancel', kind: 'secondary' }, { label: ok || t('Base.Yes'), value: 'ok', kind: 'danger' }]
      : [{ label: ok || t('Base.Yes'), value: 'ok', kind: 'primary' }, { label: cancel || t('Base.No'), value: 'cancel', kind: 'secondary' }],
  });
  return r === 'ok';
}

let uidN = 0;
export const nextId = (p = 'mo-b') => p + (++uidN);

/** A labelled field for a dialog: { label, input } → <div class="mat-field">. */
export function field(label, input, hint) {
  const id = input.id || nextId();
  input.id = id;
  return el('div', { class: 'mat-field' }, el('label', { class: 'mat-label', for: id, text: label }), input, hint ? el('div', { class: 'mat-hint', text: hint }) : null);
}

export function textInput(value = '', attrs = {}) {
  return el('input', { class: 'mat-input', type: 'text', value: value ?? '', autocomplete: 'off', spellcheck: 'false', ...attrs });
}

export function selectInput(options, value, attrs = {}) {
  const s = el('select', { class: 'mat-input mat-select', ...attrs });
  for (const o of options) {
    if (o.group) { const g = el('optgroup', { label: o.group }); for (const x of o.options) g.append(el('option', { value: x.value, text: x.label })); s.append(g); }
    else s.append(el('option', { value: o.value, text: o.label }));
  }
  if (value !== undefined && value !== null) s.value = String(value);
  return s;
}

export function checkInput(label, checked, attrs = {}) {
  const box = el('input', { type: 'checkbox', ...attrs });
  box.checked = !!checked;
  const wrap = el('label', { class: 'mat-check mo-base-check' }, box, el('span', { class: 'mat-check__text', text: label }));
  wrap.input = box;
  return wrap;
}

export function radioGroup(name, options, value) {
  const box = el('div', { class: 'mo-base-radios', role: 'radiogroup' });
  for (const o of options) {
    const r = el('input', { type: 'radio', name, value: o.value });
    r.checked = o.value === value;
    box.append(el('label', { class: 'mat-check mo-base-check' }, r, el('span', { class: 'mat-check__text', text: o.label })));
  }
  box.value = () => box.querySelector('input:checked')?.value ?? null;
  return box;
}

/** A menu at a point (right click) or under an anchor: items like the ribbon's menus [{ label, icon, run, disabled, checked, danger } | '-' | { heading }]. */
export function contextMenu(at, items) {
  const anchor = at instanceof Element ? at : (() => {
    const a = el('span', { class: 'mo-base-anchor' });
    a.style.left = at.x + 'px';
    a.style.top = at.y + 'px';
    document.body.append(a);
    setTimeout(() => a.remove(), 0);
    return a;
  })();
  const build = close => {
    const list = el('div', { class: 'mo-menu', role: 'menu' });
    for (const entry of items) {
      if (!entry) continue;
      if (entry === '-') { list.append(el('div', { class: 'mo-menu__sep', role: 'separator' })); continue; }
      if (entry.heading) { list.append(el('div', { class: 'mo-menu__heading', text: entry.heading })); continue; }
      const checked = entry.checked;
      list.append(el('button', {
        type: 'button', role: checked === undefined ? 'menuitem' : 'menuitemcheckbox', 'aria-checked': checked === undefined ? null : String(!!checked),
        class: 'mo-menu__item' + (entry.danger ? ' mo-menu__item--danger' : ''), disabled: entry.disabled ? true : null,
        onclick: () => { close(); entry.run?.(); },
      }, entry.icon ? icon(entry.icon) : el('span', { class: 'mo-menu__noicon' }), el('span', { class: 'mo-menu__label', text: entry.label }),
      entry.key ? el('span', { class: 'mo-menu__key', text: entry.key }) : null, checked ? icon('check', 'mo-menu__check') : null));
    }
    return list;
  };
  const pop = openPopover(anchor, build);
  if (!(at instanceof Element)) {
    pop.style.left = Math.max(8, Math.min(at.x, window.innerWidth - pop.offsetWidth - 8)) + 'px';
    pop.style.top = Math.max(8, Math.min(at.y, window.innerHeight - pop.offsetHeight - 8)) + 'px';
  }
  return pop;
}

export { openPopover, closePopover };

/** Follows the pointer from a pointerdown until it is released: move(dx, dy, e), up(dx, dy, e). */
export function drag(e, move, up) {
  const x0 = e.clientX;
  const y0 = e.clientY;
  const target = e.currentTarget || e.target;
  try { target.setPointerCapture?.(e.pointerId); } catch { /* not every element */ }
  const onMove = ev => move(ev.clientX - x0, ev.clientY - y0, ev);
  const onUp = ev => {
    document.removeEventListener('pointermove', onMove);
    document.removeEventListener('pointerup', onUp);
    document.removeEventListener('pointercancel', onUp);
    up?.(ev.clientX - x0, ev.clientY - y0, ev);
  };
  document.addEventListener('pointermove', onMove);
  document.addEventListener('pointerup', onUp);
  document.addEventListener('pointercancel', onUp);
}

/** The names of the field types in the language of the page. */
export function typeName(type, numberSize) {
  const names = {
    text: t('Base.Type.Text'), memo: t('Base.Type.Memo'), number: t('Base.Type.Number'), currency: t('Base.Type.Currency'), datetime: t('Base.Type.DateTime'),
    yesno: t('Base.Type.YesNo'), autonumber: t('Base.Type.AutoNumber'), hyperlink: t('Base.Type.Hyperlink'), lookup: t('Base.Type.Lookup'),
  };
  void numberSize;
  return names[type] || type;
}

export const TYPE_OPTIONS = () => ['text', 'memo', 'number', 'datetime', 'currency', 'autonumber', 'yesno', 'hyperlink'].map(v => ({ value: v, label: typeName(v) }));

export function sizeName(size) {
  return { byte: t('Base.Size.Byte'), integer: t('Base.Size.Integer'), long: t('Base.Size.Long'), single: t('Base.Size.Single'), double: t('Base.Size.Double'), decimal: t('Base.Size.Decimal') }[size] || size;
}

/** The icon of an object kind (navigation pane, tabs). */
export const KIND_ICON = { tables: 'base-table', queries: 'base-query', forms: 'base-form', reports: 'base-report', relationships: 'base-relationships', documenter: 'base-documenter' };

export function queryIcon(type) {
  return { select: 'base-query', maketable: 'base-q-maketable', append: 'base-q-append', update: 'base-q-update', delete: 'base-q-delete' }[type] || 'base-query';
}

/** A ribbon-like small icon button for panes inside the work area. */
export function iconButton(name, label, run, attrs = {}) {
  return el('button', { type: 'button', class: 'mo-base-ibtn', 'aria-label': label, 'data-tip': label, onclick: run, ...attrs }, icon(name));
}

/** Whether the target of a key event is a field where people type. */
export const isTyping = target => target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));

export { t, el, icon };
