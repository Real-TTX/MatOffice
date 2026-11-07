/* MatOffice Base – the commands of the Create tab and the wizards of Access: Table, Table Design, Query Design, the Simple Query Wizard, the Form
   button (a columnar form with a subform of the related records), Multiple Items, Datasheet and Split Form, Blank Form, the Form Wizard, the Report
   button, Blank Report, the Report Wizard (fields, grouping levels, sort order, summary options, layout, orientation, title), and the Lookup Wizard
   of the table design (values from another table with a bound and a shown field, or a list of values typed in). */
import { t, el, icon, dialog } from '../editor/util.js';
import {
  newObjectId, uniqueName, freeName, newField, addField, tableById, queryById, fieldById, fieldByName, lower, nameProblem, nameTaken, sourceById,
  previewFieldChange, applyFieldChange, saveRelationship, relationshipProblem, renameField, LIMITS,
} from './model.js';
import { columnarForm, tabularForm, blankForm as emptyForm, blankReport as emptyReport, wizardReport, sourceFieldList } from './layout.js';
import { emptyDesign, designToStatement } from './qbe.js';
import { toSql } from './sql.js';
import { showProblem, ask, field, textInput, selectInput, checkInput, radioGroup, KIND_ICON, typeName } from './ui.js';

// ---- small building blocks ---------------------------------------------------------------------------------------------------------------

/**
 * A wizard in one dialog: steps [{ title, render(box), check?() → text of a problem | null, skip?() }]; Next, Back, Finish, Cancel.
 * Resolves true when the person finished it.
 */
async function wizard({ title, icon: ic, steps, size = 'md', finish = t('Base.Wiz.Finish') }) {
  let i = 0;
  let form = null;
  const stepTitle = el('div', { class: 'mo-base-wiz__title' });
  const counter = el('div', { class: 'mo-base-wiz__count' });
  const box = el('div', { class: 'mo-base-wiz__body' });
  const error = el('p', { class: 'mo-base-wiz__error', role: 'alert', hidden: true });
  const active = () => steps.filter(s => !s.skip?.());
  const show = () => {
    const list = active();
    const s = list[i];
    stepTitle.textContent = s.title;
    counter.textContent = t('Base.Wiz.Step', i + 1, list.length);
    error.hidden = true;
    box.replaceChildren();
    s.render(box);
    if (form) {
      const next = form.querySelector('button[value="next"]');
      const back = form.querySelector('button[value="back"]');
      const ok = form.querySelector('button[value="ok"]');
      back.disabled = i === 0;
      next.disabled = i === list.length - 1;
      next.classList.toggle('mat-btn--primary', i < list.length - 1);
      next.classList.toggle('mat-btn--secondary', i === list.length - 1);
      ok.classList.toggle('mat-btn--primary', i === list.length - 1);
      ok.classList.toggle('mat-btn--secondary', i < list.length - 1);
      requestAnimationFrame(() => box.querySelector('input:not([type="hidden"]), select, textarea, button')?.focus());
    }
  };
  const fail = text => { error.textContent = text; error.hidden = false; };
  show();
  const r = await dialog({
    title, icon: ic, size, body: el('div', { class: 'mo-base-wiz' }, el('div', { class: 'mo-base-wiz__head' }, stepTitle, counter), box, error),
    buttons: [{ label: t('Base.Wiz.Next'), value: 'next', kind: 'primary' }, { label: t('Base.Wiz.Back'), value: 'back' }, { label: finish, value: 'ok' }, { label: t('Editor.Cancel'), value: 'cancel' }],
    onOpen: f => { form = f; show(); },
    validate: (f, value) => {
      const list = active();
      if (value === 'back') { if (i > 0) { i--; show(); } return false; }
      const problem = list[i].check?.();
      if (problem) { fail(problem); return false; }
      if (value === 'next') { if (i < list.length - 1) { i++; show(); } return false; }
      for (let k = i + 1; k < list.length; k++) {
        const p = list[k].check?.();
        if (p) { i = k; show(); fail(p); return false; }
      }
      return true;
    },
  });
  return r === 'ok';
}

/** The tables and queries as options of a select ("Table: Customers"). */
function sourceOptions(db, { queries = true } = {}) {
  return [
    ...db.tables.map(x => ({ value: x.id, label: t('Base.Wiz.TableOf', x.name) })),
    ...(queries ? db.queries.filter(q => { try { return !/^\s*(update|delete|insert)\b/i.test(q.sql); } catch { return false; } }).map(x => ({ value: x.id, label: t('Base.Wiz.QueryOf', x.name) })) : []),
  ];
}

const sourceName = (db, id) => tableById(db, id)?.name || queryById(db, id)?.name || '';

/**
 * Two lists like the wizards of Access: the available fields and the chosen ones, with > >> < << (and a double click). chosen: [{ key, label }] in order.
 */
function fieldPicker(available, chosen, onChange = () => {}) {
  const left = el('select', { class: 'mat-input mo-base-picker__list', multiple: true, size: 10, 'aria-label': t('Base.Wiz.Available') });
  const right = el('select', { class: 'mat-input mo-base-picker__list', multiple: true, size: 10, 'aria-label': t('Base.Wiz.Selected') });
  const fill = () => {
    const used = new Set(chosen.map(c => c.key));
    left.replaceChildren(...available().filter(a => !used.has(a.key)).map(a => el('option', { value: a.key, text: a.label })));
    right.replaceChildren(...chosen.map(c => el('option', { value: c.key, text: c.label })));
    onChange();
  };
  const add = keys => { const map = new Map(available().map(a => [a.key, a])); for (const k of keys) if (map.has(k) && !chosen.some(c => c.key === k)) chosen.push(map.get(k)); fill(); };
  const remove = keys => { for (const k of keys) { const i = chosen.findIndex(c => c.key === k); if (i >= 0) chosen.splice(i, 1); } fill(); };
  const sel = s => [...s.selectedOptions].map(o => o.value);
  const b = (label, aria, run) => el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', 'aria-label': aria, 'data-tip': aria, text: label, onclick: run });
  left.addEventListener('dblclick', () => add(sel(left)));
  right.addEventListener('dblclick', () => remove(sel(right)));
  left.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); add(sel(left)); } });
  right.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === 'Delete') { e.preventDefault(); e.stopPropagation(); remove(sel(right)); } });
  fill();
  const node = el('div', { class: 'mo-base-picker' },
    el('div', {}, el('div', { class: 'mat-label', text: t('Base.Wiz.Available') }), left),
    el('div', { class: 'mo-base-picker__buttons' },
      b('>', t('Base.Wiz.AddOne'), () => add(sel(left))), b('>>', t('Base.Wiz.AddAll'), () => add(available().map(a => a.key))),
      b('<', t('Base.Wiz.RemoveOne'), () => remove(sel(right))), b('<<', t('Base.Wiz.RemoveAll'), () => remove(chosen.map(c => c.key)))),
    el('div', {}, el('div', { class: 'mat-label', text: t('Base.Wiz.Selected') }), right));
  node.refill = fill;
  return node;
}

const isNumeric = type => type === 'number' || type === 'currency' || type === 'autonumber';

/** The name of a new object: free, valid; the person may change it in the last step. */
function nameStep(app, kinds, nameInput, extra = null) {
  return {
    title: t('Base.Wiz.TitleStep'),
    render: box => box.append(field(t('Base.Wiz.Name'), nameInput), extra),
    check: () => {
      const v = nameInput.value.trim();
      const p = nameProblem(v) || (nameTaken(app.db, kinds, v) ? { code: 'NameTaken', args: [v] } : null);
      return p ? (p.code === 'NameTaken' ? t('Base.Err.NameTaken', v) : p.code === 'NameLong' ? t('Base.Err.NameLong') : p.code === 'NameChars' ? t('Base.Err.NameChars') : t('Base.Err.NameEmpty')) : null;
    },
  };
}

const layoutTexts = () => ({
  summaryFor: t('Base.Report.SummaryForWord'), grandTotal: t('Base.Report.GrandTotalWord'), records: t('Base.Report.RecordsWord'), page: t('Base.Report.PageWord'), of: t('Base.Report.OfWord'),
  sum: t('Base.Total.Sum'), avg: t('Base.Total.Avg'), min: t('Base.Total.Min'), max: t('Base.Total.Max'),
});

/** The first child table of a table (a relationship from it): the subform the Form button adds. */
function childOf(db, tableId) {
  const rel = db.relationships.find(r => r.parent.table === tableId && r.child.table !== tableId);
  if (!rel) return null;
  const parent = tableById(db, rel.parent.table);
  const child = tableById(db, rel.child.table);
  if (!parent || !child) return null;
  return { sourceId: child.id, caption: child.name, linkMaster: fieldById(parent, rel.parent.field)?.name || '', linkChild: fieldById(child, rel.child.field)?.name || '' };
}

function limitObjects(app, kind) {
  if (app.db[kind].length >= LIMITS.objects) { app.toast(t('Base.Err.TooManyObjects', LIMITS.objects), 'warning'); return false; }
  return !app.readOnly;
}

// ---- tables -----------------------------------------------------------------------------------------------------------------------------

/** Create → Table: a new table with an ID (AutoNumber, the primary key), open in Datasheet view to type the records. */
export function newTable(app) {
  if (!limitObjects(app, 'tables')) return null;
  const db = app.db;
  const id = newObjectId(db, 'tables');
  const f = newField('autonumber', 'ID', 'f1');
  const table = { id, name: uniqueName(db, ['tables', 'queries'], t('Base.Create.TableBase')), description: '', fields: [f], primaryKey: ['f1'], rows: [], nextId: 1, datasheet: {} };
  return app.addObject('tables', table, 'datasheet');
}

/** Create → Table Design: a new table without fields, open in Design view. */
export function newTableDesign(app) {
  if (!limitObjects(app, 'tables')) return null;
  const db = app.db;
  const table = { id: newObjectId(db, 'tables'), name: uniqueName(db, ['tables', 'queries'], t('Base.Create.TableBase')), description: '', fields: [], primaryKey: [], rows: [], nextId: 1, datasheet: {} };
  return app.addObject('tables', table, 'design');
}

// ---- queries ----------------------------------------------------------------------------------------------------------------------------

/** Create → Query Design: an empty query in Design view (Show Table opens by itself). */
export function newQueryDesign(app) {
  if (!limitObjects(app, 'queries')) return null;
  const db = app.db;
  const q = { id: newObjectId(db, 'queries'), name: uniqueName(db, ['tables', 'queries'], t('Base.Create.QueryBase')), description: '', sql: '', layout: { tables: [] }, datasheet: {} };
  return app.addObject('queries', q, 'design');
}

/** The names of the fields of a table or query by name (for the design grid). */
function fieldsOfName(db, name, lang) {
  const s = db.tables.find(x => lower(x.name) === lower(name)) || db.queries.find(x => lower(x.name) === lower(name));
  return s ? sourceFieldList(db, s.id, lang).map(f => f.name) : [];
}

/** The joins between the chosen tables from their relationships; null when they cannot all be connected. */
function joinsFor(db, sources) {
  const tables = sources.map(id => tableById(db, id));
  if (sources.length < 2) return [];
  if (tables.some(x => !x)) return null;
  const joins = [];
  const parent = new Map(sources.map(id => [id, id]));
  const root = x => { while (parent.get(x) !== x) x = parent.get(x); return x; };
  for (const r of db.relationships) {
    if (!sources.includes(r.parent.table) || !sources.includes(r.child.table) || r.parent.table === r.child.table) continue;
    const p = tableById(db, r.parent.table);
    const c = tableById(db, r.child.table);
    joins.push({ left: p.name, leftField: fieldById(p, r.parent.field)?.name, right: c.name, rightField: fieldById(c, r.child.field)?.name, kind: r.join === 2 ? 'LEFT' : r.join === 3 ? 'RIGHT' : 'INNER' });
    parent.set(root(r.parent.table), root(r.child.table));
  }
  return new Set(sources.map(root)).size === 1 ? joins : null;
}

/** Create → Query Wizard (the Simple Query Wizard): fields of one or more related tables, detail or summary, the name. */
export async function queryWizard(app) {
  if (!limitObjects(app, 'queries')) return;
  const db = app.db;
  const options = sourceOptions(db);
  if (!options.length) { await showProblem({ code: 'NoTables', args: [] }); return; }
  const start = app.currentObject(['tables', 'queries'])?.obj.id || options[0].value;
  const source = selectInput(options, start);
  const chosen = [];
  const available = () => sourceFieldList(db, source.value, app.lang).map(f => ({ key: source.value + '|' + f.name, label: sourceName(db, source.value) + '.' + f.name, source: source.value, f }));
  const picker = fieldPicker(available, chosen);
  source.addEventListener('change', () => picker.refill());
  const mode = radioGroup('mo-base-qw-mode', [{ value: 'detail', label: t('Base.Wiz.Detail') }, { value: 'summary', label: t('Base.Wiz.Summary') }], 'detail');
  const summary = new Map();
  const countBox = checkInput(t('Base.Wiz.CountRecords'), false);
  const name = textInput('', { maxlength: 64 });
  const open = radioGroup('mo-base-qw-open', [{ value: 'open', label: t('Base.Wiz.OpenQuery') }, { value: 'design', label: t('Base.Wiz.ModifyQuery') }], 'open');
  const numericChosen = () => chosen.filter(c => isNumeric(c.f.type));
  const ok = await wizard({
    title: t('Base.Wiz.QueryTitle'), icon: 'base-query-wizard', size: 'lg',
    steps: [
      {
        title: t('Base.Wiz.QueryFields'),
        render: box => box.append(field(t('Base.Wiz.Source'), source), picker),
        check: () => {
          if (!chosen.length) return t('Base.Wiz.NoFields');
          const sources = [...new Set(chosen.map(c => c.source))];
          if (!joinsFor(db, sources)) return t('Base.Wiz.NotRelated');
          if (!name.value) name.value = freeName(db, ['tables', 'queries'], t('Base.Wiz.QueryName', sourceName(db, chosen[0].source)));
          return null;
        },
      },
      {
        title: t('Base.Wiz.DetailOrSummary'),
        skip: () => !numericChosen().length,
        render: box => {
          const opts = el('div', { class: 'mo-base-wiz__grid' }, el('span'), ...['Sum', 'Avg', 'Min', 'Max'].map(fn => el('span', { class: 'mat-label', text: { Sum: t('Base.Total.Sum'), Avg: t('Base.Total.Avg'), Min: t('Base.Total.Min'), Max: t('Base.Total.Max') }[fn] })));
          for (const c of numericChosen()) {
            const set = summary.get(c.key) || new Set();
            summary.set(c.key, set);
            opts.append(el('span', { text: c.label }), ...['Sum', 'Avg', 'Min', 'Max'].map(fn => {
              const b = el('input', { type: 'checkbox', 'aria-label': fn + ' ' + c.label });
              b.checked = set.has(fn);
              b.addEventListener('change', () => { if (b.checked) set.add(fn); else set.delete(fn); });
              return el('span', { class: 'mo-base-wiz__cell' }, b);
            }));
          }
          const sync = () => { opts.hidden = countBox.hidden = mode.value() !== 'summary'; };
          mode.addEventListener('change', sync);
          sync();
          box.append(mode, opts, countBox);
        },
      },
      nameStep(app, ['tables', 'queries'], name, open),
    ],
  });
  if (!ok) return;
  const sources = [...new Set(chosen.map(c => c.source))];
  const design = emptyDesign();
  design.tables = sources.map(id => ({ name: sourceName(db, id), alias: null }));
  design.joins = joinsFor(db, sources) || [];
  const summaryMode = mode.value() === 'summary' && numericChosen().length > 0;
  design.totals = summaryMode;
  for (const c of chosen) {
    const table = sourceName(db, c.source);
    const fns = summaryMode && isNumeric(c.f.type) ? [...(summary.get(c.key) || [])] : [];
    if (fns.length) for (const fn of fns) design.columns.push({ field: t('Base.Wiz.FnOf', { Sum: t('Base.Total.Sum'), Avg: t('Base.Total.Avg'), Min: t('Base.Total.Min'), Max: t('Base.Total.Max') }[fn], c.f.name).replace(/[.!`[\]:]/g, ' ') + ': ' + c.f.name, table, total: fn, sort: '', show: true, criteria: [], appendTo: '', updateTo: '', del: 'Where' });
    else design.columns.push({ field: c.f.name, table, total: 'Group By', sort: '', show: true, criteria: [], appendTo: '', updateTo: '', del: 'Where' });
  }
  if (summaryMode && countBox.input.checked) {
    const first = chosen[0];
    design.columns.push({ field: t('Base.Wiz.CountOf', sourceName(db, first.source)).replace(/[.!`[\]:]/g, ' ') + ': ' + first.f.name, table: sourceName(db, first.source), total: 'Count', sort: '', show: true, criteria: [], appendTo: '', updateTo: '', del: 'Where' });
  }
  let sql;
  try { sql = toSql(designToStatement(design, db, n => fieldsOfName(db, n, app.lang))); } catch (e) { await showProblem(e); return; }
  const q = {
    id: newObjectId(db, 'queries'), name: name.value.trim(), description: '', sql,
    layout: { tables: design.tables.map((tb, i) => ({ name: tb.name, x: 20 + i * 220, y: 16, w: 170, h: 190 })) }, datasheet: {},
  };
  await app.addObject('queries', q, open.value() === 'design' ? 'design' : null);
}

// ---- forms ------------------------------------------------------------------------------------------------------------------------------

/** The Form button (and More Forms): a form of the table or query that is open or selected, in Layout view. */
export function quickForm(app, kind = 'single') {
  if (!limitObjects(app, 'forms')) return null;
  const db = app.db;
  const cur = app.currentObject(['tables', 'queries']);
  if (!cur) { app.toast(t('Base.Wiz.PickSource'), 'info'); return null; }
  const fields = sourceFieldList(db, cur.obj.id, app.lang);
  const id = newObjectId(db, 'forms');
  const name = freeName(db, ['forms'], cur.obj.name);
  const subform = kind === 'single' && cur.kind === 'tables' ? childOf(db, cur.obj.id) : null;
  const form = kind === 'continuous'
    ? tabularForm({ id, name, sourceId: cur.obj.id, title: cur.obj.name, fields })
    : columnarForm({ id, name, sourceId: cur.obj.id, title: cur.obj.name, fields, subform, kind });
  return app.addObject('forms', form, kind === 'datasheet' ? 'form' : 'layout');
}

/** Blank Form (Layout view) and Form Design (Design view): an unbound form; the Field List is open. */
export function blankForm(app, view = 'layout') {
  if (!limitObjects(app, 'forms')) return null;
  const db = app.db;
  const form = emptyForm({ id: newObjectId(db, 'forms'), name: uniqueName(db, ['forms'], t('Base.Create.FormBase')) });
  return app.addObject('forms', form, view);
}

/** The Form Wizard: a source and its fields, the layout (columnar, tabular, datasheet, split), a subform of related records, the title. */
export async function formWizard(app) {
  if (!limitObjects(app, 'forms')) return;
  const db = app.db;
  const options = sourceOptions(db);
  if (!options.length) { await showProblem({ code: 'NoTables', args: [] }); return; }
  const source = selectInput(options, app.currentObject(['tables', 'queries'])?.obj.id || options[0].value);
  const chosen = [];
  const available = () => sourceFieldList(db, source.value, app.lang).map(f => ({ key: f.name, label: f.name, f }));
  const picker = fieldPicker(available, chosen);
  source.addEventListener('change', () => { chosen.length = 0; picker.refill(); });
  const layout = radioGroup('mo-base-fw-layout', [
    { value: 'single', label: t('Base.Wiz.Columnar') }, { value: 'continuous', label: t('Base.Wiz.Tabular') }, { value: 'datasheet', label: t('Base.Wiz.Datasheet') }, { value: 'split', label: t('Base.Form.Split') },
  ], 'single');
  const sub = checkInput('', true);
  const title = textInput('', { maxlength: 64 });
  const open = radioGroup('mo-base-fw-open', [{ value: 'form', label: t('Base.Wiz.OpenForm') }, { value: 'design', label: t('Base.Wiz.ModifyForm') }], 'form');
  const ok = await wizard({
    title: t('Base.Wiz.FormTitle'), icon: 'base-wizard', size: 'lg',
    steps: [
      {
        title: t('Base.Wiz.FormFields'),
        render: box => box.append(field(t('Base.Wiz.Source'), source), picker),
        check: () => {
          if (!chosen.length) return t('Base.Wiz.NoFields');
          if (!title.value) title.value = freeName(db, ['forms'], sourceName(db, source.value));
          return null;
        },
      },
      {
        title: t('Base.Wiz.FormLayout'),
        render: box => {
          const child = tableById(db, source.value) ? childOf(db, source.value) : null;
          box.append(layout);
          if (child) {
            sub.querySelector('.mat-check__text').textContent = t('Base.Wiz.WithSubform', child.caption);
            const sync = () => { sub.hidden = layout.value() !== 'single'; };
            layout.addEventListener('change', sync);
            sync();
            box.append(sub);
          }
        },
      },
      nameStep(app, ['forms'], title, open),
    ],
  });
  if (!ok) return;
  const all = sourceFieldList(db, source.value, app.lang);
  const fields = chosen.map(c => all.find(f => f.name === c.key)).filter(Boolean);
  const kind = layout.value();
  const id = newObjectId(db, 'forms');
  const name = title.value.trim();
  const child = kind === 'single' && sub.input.checked && tableById(db, source.value) ? childOf(db, source.value) : null;
  const form = kind === 'continuous' ? tabularForm({ id, name, sourceId: source.value, title: name, fields }) : columnarForm({ id, name, sourceId: source.value, title: name, fields, subform: child, kind });
  await app.addObject('forms', form, open.value() === 'design' ? 'design' : 'form');
}

// ---- reports ----------------------------------------------------------------------------------------------------------------------------

/** The Report button: a tabular report of the table or query that is open or selected. */
export function quickReport(app) {
  if (!limitObjects(app, 'reports')) return null;
  const db = app.db;
  const cur = app.currentObject(['tables', 'queries']);
  if (!cur) { app.toast(t('Base.Wiz.PickSource'), 'info'); return null; }
  const fields = sourceFieldList(db, cur.obj.id, app.lang);
  const report = wizardReport({
    id: newObjectId(db, 'reports'), name: freeName(db, ['reports'], cur.obj.name), sourceId: cur.obj.id, title: cur.obj.name, fields,
    landscape: fields.length > 6, layout: 'stepped', texts: layoutTexts(),
  });
  return app.addObject('reports', report, 'report');
}

/** Blank Report / Report Design: an unbound report in Design view with the Field List. */
export function blankReport(app) {
  if (!limitObjects(app, 'reports')) return null;
  const db = app.db;
  const report = emptyReport({ id: newObjectId(db, 'reports'), name: uniqueName(db, ['reports'], t('Base.Create.ReportBase')), texts: layoutTexts() });
  return app.addObject('reports', report, 'design');
}

/** The Report Wizard: fields, grouping levels, sort order and summary options, layout and orientation, title. */
export async function reportWizard(app) {
  if (!limitObjects(app, 'reports')) return;
  const db = app.db;
  const options = sourceOptions(db);
  if (!options.length) { await showProblem({ code: 'NoTables', args: [] }); return; }
  const source = selectInput(options, app.currentObject(['tables', 'queries'])?.obj.id || options[0].value);
  const chosen = [];
  const available = () => sourceFieldList(db, source.value, app.lang).map(f => ({ key: f.name, label: f.name, f }));
  const picker = fieldPicker(available, chosen);
  source.addEventListener('change', () => { chosen.length = 0; picker.refill(); });
  const groups = [];
  const groupPicker = fieldPicker(() => chosen.filter(c => c.f.type !== 'memo'), groups);
  const sorts = [0, 1, 2, 3].map(() => ({ field: selectInput([], ''), desc: selectInput([{ value: 'asc', label: t('Base.Group.AtoZ') }, { value: 'desc', label: t('Base.Group.ZtoA') }], 'asc') }));
  const summary = new Map();
  const layout = radioGroup('mo-base-rw-layout', [{ value: 'stepped', label: t('Base.Wiz.Stepped') }, { value: 'block', label: t('Base.Wiz.Block') }, { value: 'outline', label: t('Base.Wiz.Outline') }], 'stepped');
  const orientation = radioGroup('mo-base-rw-orient', [{ value: 'portrait', label: t('Base.Page.Portrait') }, { value: 'landscape', label: t('Base.Page.Landscape') }], 'portrait');
  const title = textInput('', { maxlength: 64 });
  const open = radioGroup('mo-base-rw-open', [{ value: 'preview', label: t('Base.Wiz.Preview') }, { value: 'design', label: t('Base.Wiz.ModifyReport') }], 'preview');
  const detailFields = () => chosen.filter(c => !groups.some(g => g.key === c.key));
  const ok = await wizard({
    title: t('Base.Wiz.ReportTitle'), icon: 'base-wizard', size: 'lg',
    steps: [
      {
        title: t('Base.Wiz.ReportFields'),
        render: box => box.append(field(t('Base.Wiz.Source'), source), picker),
        check: () => {
          if (!chosen.length) return t('Base.Wiz.NoFields');
          for (let k = groups.length - 1; k >= 0; k--) if (!chosen.some(c => c.key === groups[k].key)) groups.splice(k, 1);
          if (!title.value) title.value = freeName(db, ['reports'], sourceName(db, source.value));
          if (chosen.length > 6) orientation.querySelector('input[value="landscape"]').checked = true;
          return null;
        },
      },
      {
        title: t('Base.Wiz.Grouping'),
        render: box => { groupPicker.refill(); box.append(el('p', { class: 'mat-hint', text: t('Base.Wiz.GroupingHint') }), groupPicker); },
        check: () => (groups.length > 4 ? t('Base.Wiz.TooManyGroups') : null),
      },
      {
        title: t('Base.Wiz.SortSummary'),
        render: box => {
          const names = detailFields().map(c => ({ value: c.key, label: c.label }));
          const rows = sorts.map((s, i) => {
            const v = s.field.value;
            s.field.replaceChildren(el('option', { value: '', text: '' }), ...names.map(o => el('option', { value: o.value, text: o.label })));
            s.field.value = names.some(o => o.value === v) ? v : '';
            s.field.setAttribute('aria-label', t('Base.Wiz.SortN', i + 1));
            s.desc.setAttribute('aria-label', t('Base.Wiz.SortOrder'));
            return el('div', { class: 'mo-base-wiz__sort' }, el('span', { text: String(i + 1) }), s.field, s.desc);
          });
          box.append(el('div', { class: 'mo-base-wiz__sorts' }, ...rows));
          const nums = detailFields().filter(c => isNumeric(c.f.type));
          if (nums.length) {
            const grid = el('div', { class: 'mo-base-wiz__grid' }, el('span', { class: 'mat-label', text: t('Base.Wiz.SummaryOptions') }), ...['Sum', 'Avg', 'Min', 'Max'].map(fn => el('span', { class: 'mat-label', text: { Sum: t('Base.Total.Sum'), Avg: t('Base.Total.Avg'), Min: t('Base.Total.Min'), Max: t('Base.Total.Max') }[fn] })));
            for (const c of nums) {
              const set = summary.get(c.key) || new Set();
              summary.set(c.key, set);
              grid.append(el('span', { text: c.label }), ...['Sum', 'Avg', 'Min', 'Max'].map(fn => {
                const b = el('input', { type: 'checkbox', 'aria-label': fn + ' ' + c.label });
                b.checked = set.has(fn);
                b.addEventListener('change', () => { if (b.checked) set.add(fn); else set.delete(fn); });
                return el('span', { class: 'mo-base-wiz__cell' }, b);
              }));
            }
            box.append(grid);
          }
        },
      },
      {
        title: t('Base.Wiz.ReportLayout'),
        render: box => box.append(el('div', { class: 'mo-base-wiz__two' }, el('fieldset', { class: 'mo-base-fieldset' }, el('legend', { text: t('Base.Wiz.Layout') }), layout), el('fieldset', { class: 'mo-base-fieldset' }, el('legend', { text: t('Base.Prop.Orientation') }), orientation))),
      },
      nameStep(app, ['reports'], title, open),
    ],
  });
  if (!ok) return;
  const all = sourceFieldList(db, source.value, app.lang);
  const fields = chosen.map(c => all.find(f => f.name === c.key)).filter(Boolean);
  const sum = {};
  for (const [key, set] of summary) if (set.size && detailFields().some(c => c.key === key)) sum[key] = ['Sum', 'Avg', 'Min', 'Max'].filter(fn => set.has(fn));
  const name = title.value.trim();
  const report = wizardReport({
    id: newObjectId(db, 'reports'), name, sourceId: source.value, title: name, fields,
    groups: groups.map(g => g.key), sort: sorts.filter(s => s.field.value).map(s => ({ field: s.field.value, desc: s.desc.value === 'desc' })),
    summary: sum, landscape: orientation.value() === 'landscape', layout: layout.value(), texts: layoutTexts(),
  });
  await app.addObject('reports', report, open.value() === 'design' ? 'design' : 'preview');
}

// ---- the Lookup Wizard ---------------------------------------------------------------------------------------------------------------------

/**
 * Makes a field a lookup (a new one when fieldId is null: { name, index } say where): the values of another table (a bound field, usually its key, and
 * a field that is shown) or a list of values. Optionally a relationship with referential integrity (Enable Data Integrity).
 */
export async function lookupWizard(app, table, fieldId, { name: wish = null, index = null } = {}) {
  if (app.readOnly) return false;
  const db = app.db;
  const existing = fieldId ? fieldById(table, fieldId) : null;
  const lk = existing?.lookup;
  const kind = radioGroup('mo-base-lw-kind', [{ value: 'table', label: t('Base.Lookup.FromTable') }, { value: 'values', label: t('Base.Lookup.TypeValues') }], lk?.kind || (db.tables.some(x => x.id !== table.id) ? 'table' : 'values'));
  const others = db.tables.filter(x => x.id !== table.id || db.tables.length === 1);
  const target = selectInput(others.map(x => ({ value: x.id, label: x.name })), lk?.kind === 'table' ? lk.table : others[0]?.id);
  const bound = selectInput([], null);
  const shown = selectInput([], null);
  const integrity = checkInput(t('Base.Lookup.Integrity'), false);
  const fill = () => {
    const tb = tableById(db, target.value);
    const opts = (tb?.fields || []).map(f => ({ value: f.id, label: f.name + ' (' + typeName(f.type) + ')' }));
    bound.replaceChildren(...opts.map(o => el('option', { value: o.value, text: o.label })));
    shown.replaceChildren(...opts.map(o => el('option', { value: o.value, text: o.label })));
    const same = lk?.kind === 'table' && lk.table === target.value;
    bound.value = same ? lk.bound : tb?.primaryKey[0] || tb?.fields[0]?.id || '';
    shown.value = same ? lk.display : tb?.fields.find(f => f.id !== bound.value && f.type === 'text')?.id || bound.value;
  };
  target.addEventListener('change', fill);
  fill();
  const values = el('textarea', { class: 'mat-input mat-textarea', rows: 7, spellcheck: 'true' });
  values.value = lk?.kind === 'values' ? lk.values.join('\n') : '';
  const limit = checkInput(t('Base.Lookup.Limit'), lk?.kind === 'values' ? !!lk.limit : true);
  const name = textInput(existing?.name || wish || '', { maxlength: 64 });
  const isTable = () => kind.value() === 'table';
  const ok = await wizard({
    title: t('Base.Lookup.Wizard'), icon: 'base-lookup', size: 'md',
    steps: [
      { title: t('Base.Lookup.How'), render: box => box.append(kind), check: () => (isTable() && !others.length ? t('Base.Lookup.NoTables') : null) },
      { title: t('Base.Lookup.WhichTable'), skip: () => !isTable(), render: box => box.append(field(t('Base.Lookup.RowSource'), target)), check: () => (target.value ? null : t('Base.Lookup.NoTables')) },
      {
        title: t('Base.Lookup.WhichFields'), skip: () => !isTable(),
        render: box => box.append(field(t('Base.Lookup.Bound'), bound, t('Base.Lookup.BoundHint')), field(t('Base.Lookup.Shown'), shown), integrity),
        check: () => (bound.value ? null : t('Base.Wiz.NoFields')),
      },
      {
        title: t('Base.Lookup.WhichValues'), skip: () => isTable(),
        render: box => box.append(field(t('Base.Lookup.Values'), values, t('Base.Lookup.ValuesHint')), limit),
        check: () => (values.value.split(/\r?\n/).some(x => x.trim()) ? null : t('Base.Lookup.NoValues')),
      },
      {
        title: t('Base.Lookup.Label'),
        render: box => box.append(field(t('Base.Wiz.Name'), name)),
        check: () => {
          const v = name.value.trim();
          const p = nameProblem(v) || (table.fields.some(f => f.id !== fieldId && lower(f.name) === lower(v)) ? { code: 'NameTaken', args: [v] } : null);
          return p ? (p.code === 'NameTaken' ? t('Base.Err.NameTaken', v) : t('Base.Err.NameChars')) : null;
        },
      },
    ],
  });
  if (!ok) return false;
  let lookup;
  let type = { type: 'text' };
  if (isTable()) {
    const tb = tableById(db, target.value);
    const bf = fieldById(tb, bound.value);
    lookup = { kind: 'table', table: tb.id, bound: bf.id, display: shown.value || bf.id };
    type = bf.type === 'autonumber' ? { type: 'number', numberSize: 'long' } : bf.type === 'number' ? { type: 'number', numberSize: bf.numberSize || 'long' } : { type: bf.type === 'memo' ? 'text' : bf.type };
  } else {
    const list = [...new Set(values.value.split(/\r?\n/).map(x => x.trim()).filter(Boolean))].slice(0, 500);
    lookup = { kind: 'values', values: list, limit: limit.input.checked };
    type = existing && existing.type !== 'autonumber' ? { type: existing.type, numberSize: existing.numberSize } : { type: 'text' };
  }
  const fname = name.value.trim();
  let f = existing;
  if (!existing) {
    f = app.design(() => {
      const nf = addField(table, type.type, fname, index ?? table.fields.length);
      if (type.numberSize) nf.numberSize = type.numberSize;
      nf.lookup = lookup;
      return nf;
    }, [table.id]);
  } else {
    const patch = { lookup };
    if (existing.type !== type.type || (type.type === 'number' && existing.numberSize !== type.numberSize)) Object.assign(patch, type);
    const preview = previewFieldChange(db, table, existing.id, patch, app.lang);
    if (preview.blocked) { await showProblem(preview.blocked); return false; }
    if ((preview.lost || preview.truncated) && !(await ask(t('Base.Design.Lost', preview.lost + preview.truncated), { title: t('Base.Design.DataLossTitle'), danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return false;
    app.design(() => {
      applyFieldChange(db, table, existing.id, patch, app.lang);
      if (fname !== existing.name) renameField(db, table, existing.id, fname);
    }, [table.id]);
  }
  if (f && isTable() && integrity.input.checked) {
    const rel = { id: newObjectId(db, 'relationships'), parent: { table: lookup.table, field: lookup.bound }, child: { table: table.id, field: f.id }, integrity: true, cascadeUpdate: false, cascadeDelete: false, join: 1 };
    const problem = relationshipProblem(db, rel);
    if (problem) await showProblem(problem);
    else app.design(() => saveRelationship(db, rel));
  }
  return true;
}

export { KIND_ICON, icon, fieldByName, sourceById };
