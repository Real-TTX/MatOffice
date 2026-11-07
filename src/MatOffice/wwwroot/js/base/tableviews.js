/* MatOffice Base – the views of a table: the Datasheet (with "Click to Add" and the Table Fields tab) and the Design view of Access: the grid of
   field name, data type and description, the key icon, and the field properties below (General: Field Size, Format, Decimal Places, Input Mask,
   Caption, Default Value, Validation Rule and Text, Required, Indexed; Lookup: Display Control, Row Source, Bound Column, Limit To List). */
import { t, el, icon, dialog } from '../editor/util.js';
import { BaseView, datasheetCan, datasheetRun } from './view.js';
import { Datasheet } from './datasheet.js';
import { TableRecordset } from './recordset.js';
import {
  addField, previewFieldChange, applyFieldChange, deleteField, deleteFieldProblem, renameField, primaryKeyProblem, nameProblem, uniqueName, lower,
  fieldById, tableById, NUMBER_SIZES, engineValue, fieldInRelationship, isUniqueField,
} from './model.js';
import { ruleHolds } from './criteria.js';
import { compile } from './expr.js';
import { showProblem, ask, field, textInput, selectInput, typeName, sizeName, TYPE_OPTIONS, contextMenu, problemText, iconButton } from './ui.js';
import { lookupWizard } from './wizards.js';

export const FORMATS = {
  number: ['', 'General Number', 'Currency', 'Euro', 'Fixed', 'Standard', 'Percent', 'Scientific'],
  currency: ['Currency', 'Euro', 'General Number', 'Fixed', 'Standard', 'Percent', 'Scientific'],
  autonumber: ['', 'General Number', 'Fixed', 'Standard'],
  datetime: ['', 'General Date', 'Long Date', 'Medium Date', 'Short Date', 'Long Time', 'Medium Time', 'Short Time'],
  yesno: ['Yes/No', 'True/False', 'On/Off'],
  text: ['', '>', '<'],
  memo: ['', '>', '<'],
  hyperlink: [''],
};

/** The name of a format in the language of the page (the stored name stays the Access one). */
export function formatLabel(f) {
  if (!f) return t('Base.Format.None');
  const keys = {
    'General Number': 'Base.Format.GeneralNumber', Currency: 'Base.Format.Currency', Euro: 'Base.Format.Euro', Fixed: 'Base.Format.Fixed', Standard: 'Base.Format.Standard',
    Percent: 'Base.Format.Percent', Scientific: 'Base.Format.Scientific', 'General Date': 'Base.Format.GeneralDate', 'Long Date': 'Base.Format.LongDate', 'Medium Date': 'Base.Format.MediumDate',
    'Short Date': 'Base.Format.ShortDate', 'Long Time': 'Base.Format.LongTime', 'Medium Time': 'Base.Format.MediumTime', 'Short Time': 'Base.Format.ShortTime',
    'Yes/No': 'Base.Format.YesNo', 'True/False': 'Base.Format.TrueFalse', 'On/Off': 'Base.Format.OnOff', '>': 'Base.Format.Upper', '<': 'Base.Format.Lower',
  };
  return keys[f] ? t(keys[f]) : f;
}

/** Asks for a name (a field, an object): checks the rules of Access names and that it is free. */
export async function askName(title, value, taken) {
  const input = textInput(value, { maxlength: 64 });
  const r = await dialog({
    title, icon: 'edit', body: field(t('Base.Name'), input), onOpen: () => { input.focus(); input.select(); },
    validate: () => {
      const name = input.value.trim();
      const p = nameProblem(name) || (taken(name) ? { code: 'NameTaken', args: [name] } : null);
      if (p) { showProblem(p); return false; }
      return true;
    },
    buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
  });
  return r === 'ok' ? input.value.trim() : null;
}

/** Changes a field's properties with the checks and warnings of Access (data that would be lost, rules that existing data breaks). */
export async function changeField(app, table, fieldId, patch) {
  const preview = previewFieldChange(app.db, table, fieldId, patch, app.lang);
  if (preview.blocked) { await showProblem(preview.blocked); return false; }
  if (preview.lost || preview.truncated) {
    const message = preview.truncated && !preview.lost ? t('Base.Design.Truncate', preview.truncated) : t('Base.Design.Lost', preview.lost + preview.truncated);
    if (!(await ask(message, { title: t('Base.Design.DataLossTitle'), danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return false;
  }
  app.design(() => applyFieldChange(app.db, table, fieldId, patch, app.lang), [table.id]);
  return true;
}

// ---- datasheet -------------------------------------------------------------------------------------------------------------------------

export class TableDatasheetView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.selectContextTab = false;
  }

  mount() {
    const table = this.obj;
    this.rs = new TableRecordset(this.app, table.id);
    this.ds = new Datasheet(this.app, {
      recordset: this.rs,
      settings: table.datasheet,
      persist: () => this.app.settingsChanged(),
      onStatus: () => { this.app.setStatus(); this.app.ribbonRefresh(); },
      addField: this.app.readOnly ? null : anchor => this.addFieldMenu(anchor),
      renameField: c => this.renameColumn(c),
      fieldMenu: c => (this.app.readOnly ? [] : [
        { label: t('Base.Fields.Insert'), icon: 'plus', run: () => this.addFieldMenu({ x: 200, y: 200 }, this.rs.columns.indexOf(c)) },
        { label: t('Base.Fields.Rename'), icon: 'edit', run: () => this.renameColumn(c) },
        { label: t('Base.Fields.Delete'), icon: 'trash', danger: true, run: () => this.deleteColumn(c) },
      ]),
    });
    this.el.append(this.ds.el);
  }

  get table() { return this.obj; }

  refresh() {
    if (!this.table) return;
    this.ds.s = this.table.datasheet;
    this.ds.reload();
  }

  focus() { requestAnimationFrame(() => { if (!this.ds.editor && !this.ds.sel) this.ds.moveTo(this.ds.cur.r, this.ds.cur.c, { focus: true }); else this.ds.focusEditor(false); }); }

  async leave() { return this.ds.commit(); }

  destroy() { this.ds.destroy(); }

  can(cmd) { return datasheetCan(this.ds, cmd); }

  run(cmd, arg) { return datasheetRun(this.ds, cmd, arg); }

  selectionLabel(kind) { return this.ds.selectionLabel(kind); }

  status() {
    const ds = this.ds;
    const out = [];
    if (ds.s.filterOn && ds.s.filter) out.push(t('Base.Status.Filtered'));
    return out;
  }

  addFieldMenu(anchor, index) {
    const types = [['text', 'base-t-text'], ['number', 'base-t-number'], ['currency', 'base-t-currency'], ['datetime', 'base-t-date'], ['yesno', 'base-t-yesno'], ['memo', 'base-t-memo'], ['hyperlink', 'link']];
    contextMenu(anchor, [
      { heading: t('Base.Fields.AddHeading') },
      ...types.map(([type, ic]) => ({ label: typeName(type), icon: ic, run: () => this.addColumn(type, index) })),
      '-',
      { label: t('Base.Type.Lookup'), icon: 'base-lookup', run: () => this.addLookup(index) },
    ]);
  }

  async addColumn(type, index) {
    if (!(await this.ds.commit())) return;
    const table = this.table;
    const name = fieldFreeName(table, t('Base.Fields.DefaultName'));
    const f = this.app.design(() => addField(table, type, name, index ?? table.fields.length));
    this.refresh();
    const c = this.rs.columns.find(x => x.key === f.id);
    if (c) await this.renameColumn(c, true);
  }

  async addLookup(index) {
    if (!(await this.ds.commit())) return;
    const table = this.table;
    const name = fieldFreeName(table, t('Base.Fields.DefaultName'));
    await lookupWizard(this.app, table, null, { name, index: index ?? table.fields.length });
    this.refresh();
  }

  async renameColumn(c, fresh = false) {
    if (this.app.readOnly || !c.field) return;
    const table = this.table;
    const name = await askName(fresh ? t('Base.Fields.NameNew') : t('Base.Fields.Rename'), c.field.name, n => table.fields.some(f => f.id !== c.field.id && lower(f.name) === lower(n)));
    if (!name || name === c.field.name) return;
    this.app.design(() => renameField(this.app.db, table, c.field.id, name));
    this.refresh();
  }

  async deleteColumn(c) {
    const table = this.table;
    const problem = deleteFieldProblem(this.app.db, table, c.field.id);
    if (problem) { await showProblem(problem); return; }
    if (!(await ask(t('Base.Design.DeleteFields'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    this.app.design(() => deleteField(this.app.db, table, c.field.id), [table.id]);
    this.refresh();
  }

  currentField() { return this.ds.currentCol?.field || null; }

  contextTabs() {
    if (this.app.readOnly) return [];
    const f = () => this.currentField();
    const set = async patch => { const fld = f(); if (fld && await changeField(this.app, this.table, fld.id, patch)) this.refresh(); };
    const add = (type, ic) => ({ type: 'button', label: typeName(type), icon: ic, showLabel: true, run: () => this.addColumn(type) });
    return [{
      id: 'fields', label: t('Base.Tab.Fields'), groups: [
        { label: t('Base.Group.AddDelete'), items: [
          { type: 'stack', rows: [[add('text', 'base-t-text')], [add('number', 'base-t-number')], [add('currency', 'base-t-currency')]] },
          { type: 'stack', rows: [[add('datetime', 'base-t-date')], [add('yesno', 'base-t-yesno')], [{ type: 'menu', label: t('Base.Fields.More'), icon: 'more-horizontal', items: () => [
            { label: typeName('memo'), icon: 'base-t-memo', run: () => this.addColumn('memo') },
            { label: typeName('hyperlink'), icon: 'link', run: () => this.addColumn('hyperlink') },
            { label: t('Base.Type.Lookup'), icon: 'base-lookup', run: () => this.addLookup() },
          ] }]] },
          { type: 'button', size: 'large', label: t('Base.Fields.Delete'), icon: 'trash', disabled: () => !f(), run: () => { const c = this.ds.currentCol; if (c) this.deleteColumn(c); } },
        ] },
        { label: t('Base.Group.Properties'), items: [
          { type: 'button', size: 'large', label: t('Base.Fields.NameCaption'), icon: 'base-name-caption', disabled: () => !f(), run: () => this.nameAndCaption() },
          { type: 'stack', rows: [
            [{ type: 'button', showLabel: true, label: t('Base.Field.Default'), icon: 'base-default', disabled: () => !f() || f().type === 'autonumber', run: () => this.editProperty('default') }],
            [{ type: 'number', label: t('Base.Field.Size'), width: 64, min: 1, max: 255, disabled: () => f()?.type !== 'text', value: () => (f()?.type === 'text' ? f().size : ''), change: n => set({ size: n }) }],
          ] },
        ] },
        { label: t('Base.Group.Formatting'), items: [{ type: 'stack', rows: [
          [{ type: 'select', label: t('Base.Field.DataType'), width: 150, options: () => TYPE_OPTIONS(), value: () => f()?.type ?? '', disabled: () => !f(), change: v => set({ type: v }) }],
          [{ type: 'select', label: t('Base.Field.Format'), width: 150, options: () => (FORMATS[f()?.type] || ['']).map(x => ({ value: x, label: formatLabel(x) })), value: () => f()?.format ?? '', disabled: () => !f(), change: v => set({ format: v }) }],
        ] }] },
        { label: t('Base.Group.Validation'), items: [{ type: 'stack', rows: [
          [{ type: 'toggle', showLabel: true, label: t('Base.Field.Required'), icon: 'base-required', active: () => !!f()?.required, disabled: () => !f() || f().type === 'autonumber' || f().type === 'yesno', run: () => set({ required: !f().required }) }],
          [{ type: 'toggle', showLabel: true, label: t('Base.Field.Unique'), icon: 'base-unique', active: () => f()?.indexed === 'unique', disabled: () => !f(), run: () => set({ indexed: f().indexed === 'unique' ? 'no' : 'unique' }) }],
          [{ type: 'toggle', showLabel: true, label: t('Base.Field.Indexed'), icon: 'base-index', active: () => f()?.indexed !== 'no' && !!f(), disabled: () => !f(), run: () => set({ indexed: f().indexed === 'no' ? 'dup' : 'no' }) }],
        ] }, { type: 'menu', size: 'large', label: t('Base.Field.Validation'), icon: 'base-validation', disabled: () => !f(), items: () => [
          { label: t('Base.Field.ValidationRule'), run: () => this.editProperty('validationRule') },
          { label: t('Base.Field.ValidationText'), run: () => this.editProperty('validationText') },
        ] }] },
      ],
    }];
  }

  async nameAndCaption() {
    const fld = this.currentField();
    if (!fld) return;
    const table = this.table;
    const name = textInput(fld.name, { maxlength: 64 });
    const caption = textInput(fld.caption || '', { maxlength: 255 });
    const desc = textInput(fld.description || '', { maxlength: 255 });
    const r = await dialog({
      title: t('Base.Fields.NameCaption'), icon: 'edit', body: el('div', { class: 'mo-base-form' }, field(t('Base.Name'), name), field(t('Base.Field.Caption'), caption), field(t('Base.Field.Description'), desc)),
      validate: () => {
        const n = name.value.trim();
        const p = nameProblem(n) || (table.fields.some(x => x.id !== fld.id && lower(x.name) === lower(n)) ? { code: 'NameTaken', args: [n] } : null);
        if (p) { showProblem(p); return false; }
        return true;
      },
      buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    });
    if (r !== 'ok') return;
    this.app.design(() => {
      if (name.value.trim() !== fld.name) renameField(this.app.db, table, fld.id, name.value.trim());
      fld.caption = caption.value.trim();
      fld.description = desc.value.trim();
    });
    this.refresh();
  }

  async editProperty(prop) {
    const fld = this.currentField();
    if (!fld) return;
    const label = { default: t('Base.Field.Default'), validationRule: t('Base.Field.ValidationRule'), validationText: t('Base.Field.ValidationText') }[prop];
    const input = textInput(fld[prop] || '', { maxlength: prop === 'validationRule' ? 2048 : 255 });
    const r = await dialog({ title: label, icon: 'edit', body: field(label, input, t('Base.Help.' + prop)), onOpen: () => input.focus(), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'ok' && await changeField(this.app, this.table, fld.id, { [prop]: input.value.trim() })) this.refresh();
  }
}

function fieldFreeName(table, base) {
  for (let i = 1; ; i++) if (!table.fields.some(f => lower(f.name) === lower(base + i))) return base + i;
}

// ---- design view ---------------------------------------------------------------------------------------------------------------------

const PROPS = {
  text: ['size', 'format', 'inputMask', 'caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
  memo: ['format', 'caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
  number: ['numberSize', 'format', 'decimals', 'inputMask', 'caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
  currency: ['format', 'decimals', 'inputMask', 'caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
  datetime: ['format', 'inputMask', 'caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
  yesno: ['format', 'caption', 'default', 'validationRule', 'validationText', 'indexed'],
  autonumber: ['numberSizeAuto', 'format', 'caption', 'indexed'],
  hyperlink: ['caption', 'default', 'validationRule', 'validationText', 'required', 'indexed'],
};

export class TableDesignView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.selected = [0];
    this.propTab = 'general';
  }

  get table() { return this.obj; }

  mount() {
    this.grid = el('div', { class: 'mo-base-td__grid', role: 'grid', 'aria-label': t('Base.Design.Fields') });
    this.props = el('div', { class: 'mo-base-td__props' });
    this.el.append(el('div', { class: 'mo-base-td' }, this.grid, this.props));
    this.render();
  }

  refresh() { if (this.table) { this.selected = this.selected.filter(i => i <= this.table.fields.length); if (!this.selected.length) this.selected = [0]; this.render(); } }

  focus() { requestAnimationFrame(() => this.grid.querySelector(`[data-row="${this.selected[0]}"] .mo-base-td__name`)?.focus()); }

  get ro() { return this.app.readOnly; }

  render() {
    const table = this.table;
    const head = el('div', { class: 'mo-base-td__row mo-base-td__row--head', role: 'row' },
      el('div', { class: 'mo-base-td__sel', role: 'columnheader' }),
      el('div', { role: 'columnheader', text: t('Base.Design.FieldName') }),
      el('div', { role: 'columnheader', text: t('Base.Design.DataType') }),
      el('div', { role: 'columnheader', text: t('Base.Design.Description') }));
    const rows = [head];
    const count = table.fields.length + (this.ro ? 0 : Math.max(6, 18 - table.fields.length));
    for (let i = 0; i < count; i++) rows.push(this.row(i, table.fields[i] || null));
    this.grid.replaceChildren(...rows);
    this.renderProps();
  }

  row(i, f) {
    const table = this.table;
    const pk = f && table.primaryKey.includes(f.id);
    const selected = this.selected.includes(i);
    const sel = el('div', { class: 'mo-base-td__sel', role: 'rowheader', 'aria-label': pk ? t('Base.Design.PrimaryKey') : t('Base.Design.Row', i + 1) },
      pk ? icon('key') : selected ? icon('base-current') : null);
    sel.addEventListener('pointerdown', e => { e.preventDefault(); this.select(i, e.shiftKey || e.ctrlKey || e.metaKey ? (e.shiftKey ? 'range' : 'toggle') : null); });
    sel.addEventListener('contextmenu', e => { e.preventDefault(); if (!this.selected.includes(i)) this.select(i); this.rowMenu({ x: e.clientX, y: e.clientY }); });
    const name = el('input', { class: 'mo-base-td__name', value: f?.name || '', 'aria-label': t('Base.Design.FieldName'), maxlength: 64, readonly: this.ro ? true : null, autocomplete: 'off', spellcheck: 'false' });
    name.addEventListener('focus', () => this.select(i, null, false));
    name.addEventListener('change', () => this.commitName(i, f, name));
    name.addEventListener('keydown', e => this.gridKey(e, i, 0));
    const type = el('select', { class: 'mo-base-td__type', 'aria-label': t('Base.Design.DataType'), disabled: this.ro ? true : null },
      ...(f ? [] : [el('option', { value: '', text: '' })]),
      ...TYPE_OPTIONS().map(o => el('option', { value: o.value, text: o.label })),
      el('option', { value: 'lookup', text: t('Base.Design.LookupWizard') }));
    type.value = f ? f.type : '';
    type.addEventListener('focus', () => this.select(i, null, false));
    type.addEventListener('change', () => this.commitType(i, f, type));
    type.addEventListener('keydown', e => this.gridKey(e, i, 1));
    const desc = el('input', { class: 'mo-base-td__desc', value: f?.description || '', 'aria-label': t('Base.Design.Description'), maxlength: 255, readonly: this.ro || !f ? true : null, autocomplete: 'off' });
    desc.addEventListener('focus', () => this.select(i, null, false));
    desc.addEventListener('change', () => { if (f) this.app.design(() => { f.description = desc.value.trim(); }); });
    desc.addEventListener('keydown', e => this.gridKey(e, i, 2));
    return el('div', { class: 'mo-base-td__row' + (selected ? ' is-selected' : '') + (f ? '' : ' is-empty'), role: 'row', 'data-row': i }, sel, name, type, desc);
  }

  gridKey(e, i, col) {
    const move = r => { e.preventDefault(); const target = this.grid.querySelector(`[data-row="${r}"]`)?.children[col + 1]; target?.focus(); };
    if (e.key === 'ArrowDown' && !(e.target.tagName === 'SELECT')) move(i + 1);
    else if (e.key === 'ArrowUp' && !(e.target.tagName === 'SELECT')) move(i - 1);
    else if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.target.dispatchEvent(new Event('change')); move(col === 2 ? i + 1 : i); if (col < 2) this.grid.querySelector(`[data-row="${i}"]`)?.children[col + 2]?.focus(); }
    else if (e.key === 'F6') { e.preventDefault(); this.props.querySelector('input, select, textarea')?.focus(); }
  }

  select(i, mode = null, rerender = true) {
    if (mode === 'range' && this.selected.length) {
      const a = this.selected[0];
      this.selected = [];
      for (let k = Math.min(a, i); k <= Math.max(a, i); k++) this.selected.push(k);
    } else if (mode === 'toggle') {
      this.selected = this.selected.includes(i) ? this.selected.filter(x => x !== i) : [...this.selected, i];
    } else {
      if (this.selected.length === 1 && this.selected[0] === i) { return; }
      this.selected = [i];
    }
    for (const r of this.grid.querySelectorAll('.mo-base-td__row[data-row]')) {
      const k = Number(r.dataset.row);
      const on = this.selected.includes(k);
      r.classList.toggle('is-selected', on);
      const f = this.table.fields[k];
      const pk = f && this.table.primaryKey.includes(f.id);
      r.firstChild.replaceChildren(...(pk ? [icon('key')] : on ? [icon('base-current')] : []));
    }
    this.renderProps();
    this.app.ribbonRefresh();
    void rerender;
  }

  selectedFields() { return this.selected.map(i => this.table.fields[i]).filter(Boolean); }

  async commitName(i, f, input) {
    const table = this.table;
    const name = input.value.trim();
    if (!f) {
      if (!name) return;
      const p = nameProblem(name) || (table.fields.some(x => lower(x.name) === lower(name)) ? { code: 'NameTaken', args: [name] } : null);
      if (p) { await showProblem(p); input.focus(); input.select(); return; }
      this.app.design(() => addField(table, 'text', name, table.fields.length));
      this.selected = [table.fields.length - 1];
      this.render();
      this.grid.querySelector(`[data-row="${table.fields.length - 1}"] .mo-base-td__type`)?.focus();
      return;
    }
    if (name === f.name) return;
    const p = nameProblem(name) || (table.fields.some(x => x.id !== f.id && lower(x.name) === lower(name)) ? { code: 'NameTaken', args: [name] } : null);
    if (p) { await showProblem(p); input.value = f.name; input.focus(); input.select(); return; }
    this.app.design(() => renameField(this.app.db, table, f.id, name));
  }

  async commitType(i, f, select) {
    const table = this.table;
    let fld = f;
    if (select.value === 'lookup') {
      if (!fld) fld = this.app.design(() => addField(table, 'text', fieldFreeName(table, t('Base.Fields.DefaultName')), table.fields.length));
      await lookupWizard(this.app, table, fld.id);
      this.selected = [table.fields.indexOf(fld)];
      this.render();
      return;
    }
    if (!fld) {
      fld = this.app.design(() => addField(table, select.value, fieldFreeName(table, t('Base.Fields.DefaultName')), table.fields.length));
      this.selected = [table.fields.length - 1];
      this.render();
      this.grid.querySelector(`[data-row="${table.fields.length - 1}"] .mo-base-td__name`)?.select();
      return;
    }
    if (select.value === fld.type) return;
    const ok = await changeField(this.app, table, fld.id, { type: select.value });
    if (!ok) select.value = fld.type;
    this.render();
  }

  rowMenu(at) {
    contextMenu(at, [
      { label: t('Base.Design.PrimaryKey'), icon: 'key', run: () => this.togglePrimaryKey(), disabled: this.ro || !this.selectedFields().length },
      '-',
      { label: t('Base.Design.InsertRows'), icon: 'base-insert-rows', run: () => this.insertRows(), disabled: this.ro },
      { label: t('Base.Design.DeleteRows'), icon: 'base-delete-rows', danger: true, run: () => this.deleteRows(), disabled: this.ro || !this.selectedFields().length },
      { label: t('Base.Design.MoveUp'), icon: 'arrow-up', run: () => this.moveRow(-1), disabled: this.ro || this.selected[0] === 0 || !this.selectedFields().length },
      { label: t('Base.Design.MoveDown'), icon: 'arrow-down', run: () => this.moveRow(1), disabled: this.ro || this.selected[0] >= this.table.fields.length - 1 },
      '-',
      { label: t('Base.Design.ModifyLookups'), icon: 'base-lookup', run: () => this.modifyLookup(), disabled: this.ro || this.selectedFields().length !== 1 },
    ]);
  }

  async togglePrimaryKey() {
    const table = this.table;
    const ids = this.selectedFields().map(f => f.id);
    if (!ids.length) return;
    const same = ids.length === table.primaryKey.length && ids.every(id => table.primaryKey.includes(id));
    const next = same ? [] : ids;
    if (next.length) {
      const p = primaryKeyProblem(this.app.db, table, next);
      if (p) { await showProblem(p); return; }
    } else if (table.primaryKey.length && this.app.db.relationships.some(r => r.integrity && r.parent.table === table.id && table.primaryKey.includes(r.parent.field) && !isUniqueField({ ...table, primaryKey: [] }, r.parent.field))) {
      await showProblem({ code: 'PkInRelationship', args: [] });
      return;
    }
    this.app.design(() => {
      table.primaryKey = next;
      for (const id of next) { const f = fieldById(table, id); if (f) f.required = f.type === 'autonumber' ? f.required : f.required; }
    });
    this.render();
  }

  insertRows() {
    const table = this.table;
    const at = Math.min(this.selected[0] ?? table.fields.length, table.fields.length);
    this.app.design(() => addField(table, 'text', fieldFreeName(table, t('Base.Fields.DefaultName')), at));
    this.selected = [at];
    this.render();
    this.grid.querySelector(`[data-row="${at}"] .mo-base-td__name`)?.select();
  }

  async deleteRows() {
    const table = this.table;
    const fields = this.selectedFields();
    if (!fields.length) return;
    for (const f of fields) {
      const p = deleteFieldProblem(this.app.db, table, f.id);
      if (p) { await showProblem(p); return; }
    }
    const hasData = table.rows.some(r => fields.some(f => r.v[f.id] !== undefined && r.v[f.id] !== null && f.type !== 'yesno'));
    if (hasData && !(await ask(t('Base.Design.DeleteFields'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    if (fields.some(f => table.primaryKey.includes(f.id)) && !(await ask(t('Base.Design.DeleteKey'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    this.app.design(() => { for (const f of fields) deleteField(this.app.db, table, f.id); }, [table.id]);
    this.selected = [Math.min(this.selected[0], Math.max(0, table.fields.length - 1))];
    this.render();
  }

  moveRow(dir) {
    const table = this.table;
    const i = this.selected[0];
    const j = i + dir;
    if (j < 0 || j >= table.fields.length || !table.fields[i]) return;
    this.app.design(() => { const [f] = table.fields.splice(i, 1); table.fields.splice(j, 0, f); });
    this.selected = [j];
    this.render();
  }

  async modifyLookup() {
    const f = this.selectedFields()[0];
    if (!f) return;
    await lookupWizard(this.app, this.table, f.id);
    this.render();
  }

  /** Checks every record against the validation rules and required settings (Test Validation Rules). */
  async testRules() {
    const table = this.table;
    const lines = [];
    for (const f of table.fields) {
      let broken = 0;
      let empty = 0;
      for (const r of table.rows) {
        const v = r.v[f.id];
        if (f.required && (v === null || v === undefined || v === '')) empty++;
        if (f.validationRule) {
          try { if (!ruleHolds(f.validationRule, engineValue(f, v ?? null), compile, { lang: this.app.lang })) broken++; } catch { broken++; }
        }
      }
      if (broken) lines.push(t('Base.Design.RuleBroken', f.name, broken));
      if (empty) lines.push(t('Base.Design.RequiredEmpty', f.name, empty));
    }
    await dialog({ title: t('Base.Design.TestRules'), icon: lines.length ? 'alert' : 'check-circle', body: el('div', {}, el('p', { text: lines.length ? t('Base.Design.TestFailed') : t('Base.Design.TestPassed') }), lines.length ? el('ul', {}, ...lines.map(x => el('li', { text: x }))) : null) });
  }

  async showIndexes() {
    const table = this.table;
    const rows = [];
    if (table.primaryKey.length) rows.push([t('Base.Design.PrimaryKeyIndex'), table.primaryKey.map(id => fieldById(table, id)?.name).join(', '), t('Base.Yes')]);
    for (const f of table.fields) if (f.indexed !== 'no' && !(table.primaryKey.length === 1 && table.primaryKey[0] === f.id)) rows.push([f.name, f.name, f.indexed === 'unique' ? t('Base.Yes') : t('Base.No')]);
    const list = el('table', { class: 'mo-base-simpletable' }, el('thead', {}, el('tr', {}, ...[t('Base.Design.IndexName'), t('Base.Design.IndexFields'), t('Base.Design.IndexUnique')].map(h => el('th', { text: h })))),
      el('tbody', {}, ...rows.map(r => el('tr', {}, ...r.map((c, i) => el('td', { text: c, 'data-label': [t('Base.Design.IndexName'), t('Base.Design.IndexFields'), t('Base.Design.IndexUnique')][i] }))))));
    await dialog({ title: t('Base.Design.Indexes', table.name), icon: 'base-index', size: 'md', body: rows.length ? list : el('p', { text: t('Base.Design.NoIndexes') }) });
  }

  async tableProperties() {
    const table = this.table;
    const desc = el('textarea', { class: 'mat-input mat-textarea', rows: 3, maxlength: 1000, readonly: this.ro ? true : null });
    desc.value = table.description || '';
    const r = await dialog({ title: t('Base.Design.PropertySheet'), icon: 'sliders', body: el('div', { class: 'mo-base-form' }, field(t('Base.Field.Description'), desc)), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    if (r === 'ok' && !this.ro) this.app.design(() => { table.description = desc.value.trim(); });
  }

  // ---- field properties ------------------------------------------------------------------------------------------------------------

  renderProps() {
    const table = this.table;
    const f = table.fields[this.selected[0]] || null;
    const tabs = el('div', { class: 'mo-base-td__tabs', role: 'tablist' },
      el('button', { type: 'button', role: 'tab', 'aria-selected': String(this.propTab === 'general'), onclick: () => { this.propTab = 'general'; this.renderProps(); } }, el('span', { text: t('Base.Design.General') })),
      el('button', { type: 'button', role: 'tab', 'aria-selected': String(this.propTab === 'lookup'), onclick: () => { this.propTab = 'lookup'; this.renderProps(); } }, el('span', { text: t('Base.Design.Lookup') })));
    this.help = el('div', { class: 'mo-base-td__help', text: t('Base.Help.Grid') });
    const list = el('div', { class: 'mo-base-td__plist' });
    if (f) {
      if (this.propTab === 'general') for (const p of PROPS[f.type] || []) list.append(this.propRow(f, p));
      else this.lookupProps(f, list);
    }
    this.props.replaceChildren(el('div', { class: 'mo-base-td__pleft' }, el('div', { class: 'mo-base-td__ptitle', text: t('Base.Design.FieldProperties') }), tabs, list), this.help);
  }

  propRow(f, p) {
    const table = this.table;
    const label = {
      size: t('Base.Field.Size'), numberSize: t('Base.Field.Size'), numberSizeAuto: t('Base.Field.Size'), format: t('Base.Field.Format'), decimals: t('Base.Field.Decimals'), inputMask: t('Base.Field.InputMask'),
      caption: t('Base.Field.Caption'), default: t('Base.Field.Default'), validationRule: t('Base.Field.ValidationRule'), validationText: t('Base.Field.ValidationText'),
      required: t('Base.Field.Required'), indexed: t('Base.Field.Indexed'),
    }[p];
    let input;
    const commit = async value => {
      if (this.ro) return;
      if (String(value ?? '') === String(f[p] ?? '')) return;
      const ok = await changeField(this.app, table, f.id, { [p]: value });
      this.render();
      if (!ok) return;
      this.props.querySelector(`[data-prop="${p}"]`)?.focus();
    };
    switch (p) {
      case 'size': input = el('input', { type: 'number', min: 1, max: 255, value: f.size ?? 255 }); input.addEventListener('change', () => commit(Math.max(1, Math.min(255, Number(input.value) || 255)))); break;
      case 'numberSize': input = selectInput(NUMBER_SIZES.map(s => ({ value: s, label: sizeName(s) })), f.numberSize || 'long'); input.addEventListener('change', () => commit(input.value)); break;
      case 'numberSizeAuto': input = el('input', { value: sizeName('long'), readonly: true }); break;
      case 'format': {
        const options = FORMATS[f.type] || [''];
        input = selectInput([...options, ...(options.includes(f.format) ? [] : [f.format])].map(x => ({ value: x, label: formatLabel(x) })), f.format || '');
        input.addEventListener('change', () => commit(input.value));
        break;
      }
      case 'decimals': input = selectInput([{ value: '', label: t('Base.Field.Auto') }, ...[...Array(16).keys()].map(n => ({ value: String(n), label: String(n) }))], f.decimals === null || f.decimals === undefined ? '' : String(f.decimals)); input.addEventListener('change', () => commit(input.value === '' ? null : Number(input.value))); break;
      case 'required': input = selectInput([{ value: 'no', label: t('Base.No') }, { value: 'yes', label: t('Base.Yes') }], f.required ? 'yes' : 'no'); input.addEventListener('change', () => commit(input.value === 'yes')); break;
      case 'indexed': {
        const pk = table.primaryKey.length === 1 && table.primaryKey[0] === f.id;
        input = selectInput([{ value: 'no', label: t('Base.Index.No') }, { value: 'dup', label: t('Base.Index.Dup') }, { value: 'unique', label: t('Base.Index.Unique') }], pk ? 'unique' : f.indexed || 'no', { disabled: pk ? true : null });
        input.addEventListener('change', () => commit(input.value));
        break;
      }
      default:
        input = el('input', { value: f[p] ?? '', maxlength: p === 'validationRule' ? 2048 : 255, autocomplete: 'off', spellcheck: p === 'caption' || p === 'validationText' ? 'true' : 'false' });
        input.addEventListener('change', () => commit(input.value.trim()));
        input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); input.dispatchEvent(new Event('change')); } });
    }
    input.classList.add('mo-base-td__pinput');
    input.dataset.prop = p;
    if (this.ro) { input.setAttribute('readonly', ''); if (input.tagName === 'SELECT') input.disabled = true; }
    input.addEventListener('focus', () => { this.help.textContent = t('Base.Help.' + (p === 'numberSize' || p === 'numberSizeAuto' ? 'size' : p)); });
    const id = 'mo-base-p-' + p;
    input.id = id;
    return el('div', { class: 'mo-base-td__prow' }, el('label', { for: id, text: label }), input);
  }

  lookupProps(f, list) {
    const lk = f.lookup;
    const row = (label, value) => el('div', { class: 'mo-base-td__prow' }, el('span', { class: 'mo-base-td__plabel', text: label }), el('span', { class: 'mo-base-td__pvalue', text: value }));
    list.append(row(t('Base.Lookup.Display'), lk ? t('Base.Lookup.Combo') : t('Base.Lookup.TextBox')));
    if (lk?.kind === 'table') {
      const target = tableById(this.app.db, lk.table);
      list.append(row(t('Base.Lookup.RowSourceType'), t('Base.Lookup.TableQuery')), row(t('Base.Lookup.RowSource'), target?.name || ''),
        row(t('Base.Lookup.Bound'), fieldById(target, lk.bound)?.name || ''), row(t('Base.Lookup.Shown'), fieldById(target, lk.display)?.name || ''));
    } else if (lk?.kind === 'values') {
      list.append(row(t('Base.Lookup.RowSourceType'), t('Base.Lookup.ValueList')), row(t('Base.Lookup.RowSource'), lk.values.join('; ')), row(t('Base.Lookup.Limit'), lk.limit ? t('Base.Yes') : t('Base.No')));
    }
    if (!this.ro && f.type !== 'autonumber') {
      list.append(el('div', { class: 'mat-btn-row mat-btn-row--plain mo-base-td__pbuttons' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => this.modifyLookup() }, icon('base-lookup'), el('span', { text: lk ? t('Base.Design.ModifyLookups') : t('Base.Design.LookupWizard') })),
        lk ? el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => { this.app.design(() => { f.lookup = null; }); this.render(); } }, icon('x'), el('span', { text: t('Base.Lookup.Remove') })) : null));
    }
  }

  // ---- ribbon ----------------------------------------------------------------------------------------------------------------------

  contextTabs() {
    const one = () => this.selectedFields().length > 0;
    return [{
      id: 'tdesign', label: t('Base.Tab.TableDesign'), groups: [
        { label: t('Base.Group.Tools'), items: [
          { type: 'toggle', size: 'large', label: t('Base.Design.PrimaryKey'), icon: 'key', active: () => this.selectedFields().some(f => this.table.primaryKey.includes(f.id)), disabled: () => !one(), run: () => this.togglePrimaryKey() },
          { type: 'stack', rows: [
            [{ type: 'button', showLabel: true, label: t('Base.Design.TestRules'), icon: 'base-validation', edit: false, run: () => this.testRules() }],
            [{ type: 'button', showLabel: true, label: t('Base.Design.ModifyLookups'), icon: 'base-lookup', disabled: () => this.selectedFields().length !== 1, run: () => this.modifyLookup() }],
          ] },
        ] },
        { label: t('Base.Group.Rows'), items: [{ type: 'stack', rows: [
          [{ type: 'button', showLabel: true, label: t('Base.Design.InsertRows'), icon: 'base-insert-rows', run: () => this.insertRows() }],
          [{ type: 'button', showLabel: true, label: t('Base.Design.DeleteRows'), icon: 'base-delete-rows', disabled: () => !one(), run: () => this.deleteRows() }],
          [{ type: 'button', showLabel: true, label: t('Base.Design.MoveUp'), icon: 'arrow-up', disabled: () => !one() || this.selected[0] === 0, run: () => this.moveRow(-1) }],
        ] }] },
        { label: t('Base.Group.ShowHide'), items: [
          { type: 'button', size: 'large', label: t('Base.Design.PropertySheet'), icon: 'sliders', edit: false, run: () => this.tableProperties() },
          { type: 'button', size: 'large', label: t('Base.Design.IndexesButton'), icon: 'base-index', edit: false, run: () => this.showIndexes() },
        ] },
      ],
    }];
  }

  status() {
    const f = this.selectedFields()[0];
    return f ? [t('Base.Status.Field', f.name, typeName(f.type))] : [];
  }

  can(cmd) { return cmd === 'find' ? false : false; }

  run() {}
}

export { problemText, iconButton, fieldInRelationship };
