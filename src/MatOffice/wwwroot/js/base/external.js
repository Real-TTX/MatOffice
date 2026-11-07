/* MatOffice Base – the External Data tab: import from Excel (the server reads the workbook: /live/convert/base/read, every sheet a table with guessed
   types) and from text files (CSV, read here), as new tables or appended to a table; export of a table or a query result to Excel (/live/convert/base/write
   with a content of only that table) and to CSV (made here), and of the whole database to Excel. In a host session the secret of the session goes along. */
import { t, el, dialog, download, request, token, pickFile } from '../editor/util.js';
import {
  newObjectId, freeName, newField, tableById, queryById, fieldByName, lower, coerce, engineValue, saveRecord, totalRecords, LIMITS, nameProblem, nameTaken,
} from './model.js';
import { TableRecordset, QueryRecordset } from './recordset.js';
import { askParams } from './queryviews.js';
import { parseSql } from './sql.js';
import { parseNumber, parseDate, parseBool } from './values.js';
import { showProblem, field, textInput, selectInput, checkInput, problemText } from './ui.js';

/** "&host=<secret>" in a host session (the last part of the address of the file). */
function hostParam(app) {
  const doc = app.ctx.doc;
  if (doc.mode !== 'host') return '';
  const secret = doc.api.split('?')[0].split('/').filter(Boolean).pop();
  return '&host=' + encodeURIComponent(secret);
}

/** A message of the server (or of a failed conversion) in the Access way. */
function say(message) {
  return dialog({ title: t('Base.AppName'), icon: 'alert', body: el('p', { class: 'mo-base-problem', text: message }), buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }] });
}

function dbName(app) {
  const d = app.ctx.doc;
  const name = d.mode === 'host' ? String(d.fileName || '').replace(/\.[^.]+$/, '') : d.name;
  return name || t('Base.AppName');
}

// ---- import ------------------------------------------------------------------------------------------------------------------------------

/** External Data → Excel: the sheets of a workbook as new tables (or appended to a table). */
export async function importExcel(app) {
  if (app.readOnly) { app.toast(t('Base.Ext.ReadOnly'), 'info'); return; }
  const file = await pickFile('.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  if (!file) return;
  if (file.size > 30 * 1024 * 1024) { app.toast(t('Base.Ext.TooLarge'), 'warning'); return; }
  app.toast(t('Base.Ext.Reading', file.name));
  let content;
  try {
    content = await request('POST', '/live/convert/base/read?ext=.xlsx' + hostParam(app), file);
  } catch (e) {
    await say(e.body?.error || (e.status === 0 ? t('Base.Ext.Offline') : t('Base.Ext.ImportFailed')));
    return;
  }
  const tables = (content?.tables || []).filter(tb => tb.fields?.length);
  if (!tables.length) { app.toast(t('Base.Ext.NothingToImport'), 'info'); return; }
  await importTables(app, tables, file.name);
}

/** The Import dialog: per table of the file whether to import it, as which new table or into which table. */
async function importTables(app, tables, fileName) {
  const db = app.db;
  const rows = tables.map(tb => {
    const use = checkInput(t('Base.Ext.ImportThis', tb.name, tb.rows.length), true);
    const name = textInput(freeName(db, ['tables', 'queries'], tb.name), { maxlength: 64, 'aria-label': t('Base.Ext.NewName') });
    const mode = selectInput([{ value: '', label: t('Base.Ext.AsNewTable') }, ...db.tables.map(x => ({ value: x.id, label: t('Base.Ext.AppendTo', x.name) }))], '', { 'aria-label': t('Base.Ext.Target') });
    const sync = () => { name.disabled = !use.input.checked || mode.value !== ''; mode.disabled = !use.input.checked; };
    use.input.addEventListener('change', sync);
    mode.addEventListener('change', sync);
    sync();
    return { tb, use, name, mode, node: el('div', { class: 'mo-base-imp__row' }, use, el('div', { class: 'mo-base-imp__opts' }, mode, name)) };
  });
  const r = await dialog({
    title: t('Base.Ext.ImportTitle', fileName), icon: 'base-excel', size: 'md',
    body: el('div', { class: 'mo-base-form mo-base-imp' }, el('p', { text: t('Base.Ext.ImportText') }), ...rows.map(x => x.node)),
    validate: () => {
      const names = new Set();
      for (const x of rows) {
        if (!x.use.input.checked || x.mode.value) continue;
        const v = x.name.value.trim();
        const p = nameProblem(v) || (nameTaken(db, ['tables', 'queries'], v) || names.has(lower(v)) ? { code: 'NameTaken', args: [v] } : null);
        if (p) { showProblem(p); x.name.focus(); return false; }
        names.add(lower(v));
      }
      const adding = rows.filter(x => x.use.input.checked).reduce((n, x) => n + x.tb.rows.length, 0);
      if (totalRecords(db) + adding > LIMITS.records) { showProblem({ code: 'TooManyRecords', args: [LIMITS.records] }); return false; }
      return true;
    },
    buttons: [{ label: t('Base.Ext.Import'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
  });
  if (r !== 'ok') return;
  const fresh = rows.filter(x => x.use.input.checked && !x.mode.value);
  const appends = rows.filter(x => x.use.input.checked && x.mode.value);
  let added = [];
  if (fresh.length) {
    added = app.design(() => fresh.map(x => {
      const tb = structuredClone(x.tb);
      tb.id = newObjectId(db, 'tables');
      tb.name = x.name.value.trim();
      tb.datasheet ||= {};
      db.tables.push(tb);
      return tb;
    })) || [];
  }
  let done = 0;
  let failed = 0;
  const touched = [];
  for (const x of appends) {
    const target = tableById(db, x.mode.value);
    if (!target) continue;
    const res = appendRows(app, target, x.tb.fields, x.tb.rows.map(row => x.tb.fields.map(f => engineValue(f, row.v[f.id] ?? null))));
    done += res.done;
    failed += res.failed;
    touched.push(target.id);
  }
  if (touched.length) app.recordsChanged(touched);
  for (const tab of app.tabs) tab.viewObj.version = -1;
  app.activeView?.activate();
  app.nav.render();
  const parts = [];
  if (added.length) parts.push(t('Base.Ext.Imported', added.length, added.reduce((n, tb) => n + tb.rows.length, 0)));
  if (appends.length) parts.push(t('Base.Ext.Appended', done, failed));
  app.toast(parts.join(' '), failed ? 'warning' : 'success');
  if (added[0]) app.open('tables', added[0].id, 'datasheet');
}

/** Appends records to a table: the columns go to the fields of the same name; every record goes through the checks of the table. */
function appendRows(app, target, fields, rows) {
  const db = app.db;
  const map = fields.map(f => fieldByName(target, f.name));
  let done = 0;
  let failed = 0;
  for (const values of rows) {
    if (totalRecords(db) >= LIMITS.records) { failed++; continue; }
    const input = {};
    let bad = false;
    map.forEach((tf, i) => {
      if (!tf || tf.type === 'autonumber' || bad) return;
      const c = coerce(tf, values[i], app.lang);
      if (!c.ok) bad = true; else if (c.value !== null) input[tf.id] = c.value;
    });
    if (bad) { failed++; continue; }
    const r = saveRecord(db, target, null, input, app.lang);
    if (r.ok) done++; else failed++;
  }
  return { done, failed };
}

/** External Data → Text File: a CSV or tab separated file, read in the browser (separator, first row = field names, the types guessed). */
export async function importText(app) {
  if (app.readOnly) { app.toast(t('Base.Ext.ReadOnly'), 'info'); return; }
  const file = await pickFile('.csv,.txt,.tsv,text/csv,text/plain,text/tab-separated-values');
  if (!file) return;
  if (file.size > 30 * 1024 * 1024) { app.toast(t('Base.Ext.TooLarge'), 'warning'); return; }
  const bytes = new Uint8Array(await file.arrayBuffer());
  let text = new TextDecoder('utf-8').decode(bytes);
  if (text.includes('�')) text = new TextDecoder('windows-1252').decode(bytes);
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/, 1)[0] || '';
  const counts = { ';': (first.match(/;/g) || []).length, ',': (first.match(/,/g) || []).length, '\t': (first.match(/\t/g) || []).length, '|': (first.match(/\|/g) || []).length };
  const guess = Object.entries(counts).sort((a, b) => b[1] - a[1])[0][1] ? Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0] : ',';
  const sep = selectInput([{ value: ';', label: t('Base.Ext.Semicolon') }, { value: ',', label: t('Base.Ext.Comma') }, { value: '\t', label: t('Base.Ext.Tab') }, { value: '|', label: '|' }], guess);
  const header = checkInput(t('Base.Ext.FirstRowNames'), true);
  const name = textInput(freeName(app.db, ['tables', 'queries'], file.name.replace(/\.[^.]+$/, '')), { maxlength: 64 });
  const mode = selectInput([{ value: '', label: t('Base.Ext.AsNewTable') }, ...app.db.tables.map(x => ({ value: x.id, label: t('Base.Ext.AppendTo', x.name) }))], '');
  const preview = el('div', { class: 'mo-base-imp__preview' });
  let parsed = null;
  const update = () => {
    const rows = parseCsv(text, sep.value);
    parsed = shapeRows(rows, header.input.checked, app.lang);
    const head = el('tr', {}, ...parsed.fields.map(f => el('th', {}, el('span', { text: f.name }), el('small', { text: ' ' + typeLabel(f) }))));
    preview.replaceChildren(el('table', { class: 'mo-base-simpletable' }, el('thead', {}, head),
      el('tbody', {}, ...parsed.rows.slice(0, 6).map(r => el('tr', {}, ...r.map(v => el('td', { text: v ?? '' })))))),
    el('p', { class: 'mat-hint', text: t('Base.Ext.RowsFound', parsed.rows.length) }));
    name.disabled = mode.value !== '';
  };
  sep.addEventListener('change', update);
  header.input.addEventListener('change', update);
  mode.addEventListener('change', update);
  update();
  const r = await dialog({
    title: t('Base.Ext.ImportTitle', file.name), icon: 'base-text-file', size: 'lg',
    body: el('div', { class: 'mo-base-form' }, el('div', { class: 'mo-base-imp__opts' }, field(t('Base.Ext.Separator'), sep), field(t('Base.Ext.Target'), mode), field(t('Base.Ext.NewName'), name)), header, preview),
    validate: () => {
      if (!parsed.fields.length) { showProblem({ code: 'NoColumns', args: [] }); return false; }
      if (!mode.value) {
        const v = name.value.trim();
        const p = nameProblem(v) || (nameTaken(app.db, ['tables', 'queries'], v) ? { code: 'NameTaken', args: [v] } : null);
        if (p) { showProblem(p); return false; }
      }
      if (totalRecords(app.db) + parsed.rows.length > LIMITS.records) { showProblem({ code: 'TooManyRecords', args: [LIMITS.records] }); return false; }
      return true;
    },
    buttons: [{ label: t('Base.Ext.Import'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }],
  });
  if (r !== 'ok') return;
  const db = app.db;
  if (mode.value) {
    const target = tableById(db, mode.value);
    const fields = parsed.fields.map(f => ({ name: f.name }));
    const res = appendRows(app, target, fields, parsed.rows.map(row => row.map(v => (v === '' ? null : v))));
    app.recordsChanged([target.id]);
    for (const tab of app.tabs) tab.viewObj.version = -1;
    app.activeView?.activate();
    app.toast(t('Base.Ext.Appended', res.done, res.failed), res.failed ? 'warning' : 'success');
    return;
  }
  const table = app.design(() => {
    const fields = parsed.fields.map((f, i) => { const nf = newField(f.type, f.name, 'f' + (i + 1)); if (f.type === 'number') nf.numberSize = f.numberSize; return nf; });
    const tb = { id: newObjectId(db, 'tables'), name: name.value.trim(), description: '', fields, primaryKey: [], rows: [], nextId: 1, datasheet: {} };
    let failed = 0;
    for (const row of parsed.rows) {
      const v = {};
      fields.forEach((f, i) => {
        const raw = row[i];
        if (raw === undefined || raw === null || raw === '') { if (f.type === 'yesno') v[f.id] = false; return; }
        const c = coerce(f, raw, app.lang);
        if (c.ok || c.value !== null) { if (c.value !== null && c.value !== undefined) v[f.id] = c.value; if (!c.ok) failed++; } else failed++;
      });
      tb.rows.push({ id: tb.nextId++, v });
    }
    db.tables.push(tb);
    tb.failed = failed;
    return tb;
  });
  if (!table) return;
  const failed = table.failed;
  delete table.failed;
  app.nav.render();
  app.toast(t('Base.Ext.Imported', 1, table.rows.length) + (failed ? ' ' + t('Base.Ext.ValuesFailed', failed) : ''), failed ? 'warning' : 'success');
  app.open('tables', table.id, 'datasheet');
}

/** RFC 4180 CSV with any separator: quotes, doubled quotes, line breaks inside quotes. */
export function parseCsv(text, sep = ',') {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false; } else cell += c;
      continue;
    }
    if (c === '"' && cell === '') quoted = true;
    else if (c === sep) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      cell = '';
      if (row.length > 1 || row[0] !== '') rows.push(row);
      row = [];
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); if (row.length > 1 || row[0] !== '') rows.push(row); }
  return rows.slice(0, LIMITS.records + 1);
}

/** Field names (unique, valid) and guessed types of the columns of a text file. */
function shapeRows(rows, header, lang) {
  const width = Math.min(255, rows.reduce((m, r) => Math.max(m, r.length), 0));
  const names = [];
  const used = new Set();
  for (let i = 0; i < width; i++) {
    let n = header ? String(rows[0]?.[i] ?? '').replace(/[.!`[\]<>"\u0000-\u001f]/g, ' ').trim().slice(0, 64) : '';
    if (!n) n = t('Base.Ext.FieldN', i + 1);
    let k = n;
    for (let j = 2; used.has(lower(k)); j++) k = n.slice(0, 60) + j;
    used.add(lower(k));
    names.push(k);
  }
  const body = (header ? rows.slice(1) : rows).map(r => { const out = r.slice(0, width); while (out.length < width) out.push(''); return out.map(x => x.trim()); });
  const fields = names.map((name, i) => ({ name, ...guessType(body.map(r => r[i]), lang) }));
  return { fields, rows: body };
}

/** The type Access would guess for a column of texts: Yes/No, a number (whole or with decimals), currency, a date, short or long text. */
function guessType(values, lang) {
  const present = values.filter(v => v !== '' && v !== null && v !== undefined).slice(0, 5000);
  if (!present.length) return { type: 'text' };
  if (present.every(v => parseBool(v) !== null && !/^-?\d+$/.test(v))) return { type: 'yesno' };
  const num = v => parseNumber(v, lang) ?? parseNumber(v, 'en');
  if (present.every(v => /^[-+]?[\d.,\s]+\s*[€$£]?$|^[€$£]\s*[-+]?[\d.,]+$/.test(v) && num(v.replace(/[€$£\s]/g, '')) !== null)) {
    if (present.some(v => /[€$£]/.test(v))) return { type: 'currency' };
    const ns = present.map(v => num(v));
    if (ns.every(n => Number.isInteger(n) && Math.abs(n) <= 2147483647)) return { type: 'number', numberSize: 'long' };
    return { type: 'number', numberSize: 'double' };
  }
  if (present.every(v => /\d/.test(v) && /[./-]/.test(v) && parseDate(v, lang) !== null)) return { type: 'datetime' };
  if (present.some(v => v.length > 255)) return { type: 'memo' };
  return { type: 'text' };
}

function typeLabel(f) {
  return { text: t('Base.Type.Text'), memo: t('Base.Type.Memo'), number: t('Base.Type.Number'), currency: t('Base.Type.Currency'), datetime: t('Base.Type.DateTime'), yesno: t('Base.Type.YesNo') }[f.type] || f.type;
}

// ---- export ------------------------------------------------------------------------------------------------------------------------------

/** The records of a table or a query (with its prompts); null when cancelled. */
async function recordsOf(app, kind, id) {
  if (kind === 'tables') return tableById(app.db, id) ? new TableRecordset(app, id) : null;
  const q = queryById(app.db, id);
  if (!q) return null;
  let params;
  try { params = await askParams(app, parseSql(q.sql)); } catch (e) { await showProblem(e); return null; }
  if (params === null) return null;
  const rs = new QueryRecordset(app, id, params);
  if (rs.error) { await showProblem(rs.error); return null; }
  return rs;
}

/** A table of a content made of a recordset (the shown values of lookups, dates as dates, numbers as numbers). */
function tableOf(rs, name) {
  const used = new Set();
  const cols = rs.columns.map((c, i) => {
    let fname = String(c.name).replace(/[.!`[\]<>"\u0000-\u001f]/g, '_').trim().slice(0, 64) || 'Field' + (i + 1);
    for (let j = 2; used.has(lower(fname)); j++) fname = String(c.name).slice(0, 60) + j;
    used.add(lower(fname));
    const lookup = !!c.field?.lookup;
    let type = lookup ? 'text' : c.type === 'autonumber' ? 'number' : c.type || 'text';
    if (!['text', 'memo', 'number', 'currency', 'datetime', 'yesno', 'hyperlink'].includes(type)) type = 'text';
    return { c, lookup, f: newField(type, fname, 'f' + (i + 1)) };
  });
  for (const x of cols) {
    if (x.f.type === 'number') x.f.numberSize = x.c.field?.numberSize || (x.c.type === 'autonumber' ? 'long' : 'double');
    if (x.f.type === 'text' && rs.rows.some(r => { const v = x.lookup ? rs.display(x.c, rs.get(r, x.c)) : rs.get(r, x.c); return typeof v === 'string' && v.length > 255; })) { x.f.type = 'memo'; delete x.f.size; }
  }
  const rows = rs.rows.map((r, k) => {
    const v = {};
    for (const x of cols) {
      let value = rs.get(r, x.c);
      if (x.lookup) value = rs.display(x.c, value) || null;
      const co = coerce(x.f, value, 'en');
      const keep = co.ok ? co.value : co.value ?? null;
      if (keep !== null && keep !== undefined) v[x.f.id] = keep;
    }
    return { id: k + 1, v };
  });
  return { id: 't1', name: String(name).replace(/[.!`[\]<>"]/g, '_').slice(0, 64) || 'Table', description: '', fields: cols.map(x => x.f), primaryKey: [], rows, nextId: rows.length + 1, datasheet: {} };
}

/** Sends a content to the server and gets a file back. */
async function convert(app, content, name) {
  const url = '/live/convert/base/write?ext=.xlsx&name=' + encodeURIComponent(name) + hostParam(app);
  let response;
  try {
    response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json', RequestVerificationToken: token(), Accept: 'application/json, application/octet-stream' }, body: JSON.stringify(content), credentials: 'same-origin', cache: 'no-store' });
  } catch {
    throw Object.assign(new Error('network'), { status: 0 });
  }
  if (!response.ok) {
    const body = (response.headers.get('Content-Type') || '').includes('json') ? await response.json().catch(() => null) : null;
    throw Object.assign(new Error('HTTP ' + response.status), { status: response.status, body });
  }
  return response.blob();
}

const emptyContent = db => ({ version: 1, settings: { currency: db.settings.currency }, tables: [], relationships: [], queries: [], forms: [], reports: [], ui: { nav: {}, relationships: {} } });

/** Export → Excel: one table or the result of one query as a workbook with one sheet. */
export async function exportExcel(app, kind, id) {
  const obj = (kind === 'tables' ? tableById : queryById)(app.db, id);
  if (!obj) return;
  const rs = await recordsOf(app, kind, id);
  if (!rs) return;
  const content = emptyContent(app.db);
  content.tables.push(tableOf(rs, obj.name));
  try {
    const blob = await convert(app, content, obj.name);
    download(obj.name + '.xlsx', blob);
    app.toast(t('Base.Ext.Exported', obj.name, rs.rows.length), 'success');
  } catch (e) {
    await say(e.body?.error || (e.status === 0 ? t('Base.Ext.Offline') : t('Base.Ext.ExportFailed')));
  }
}

/** Export → Excel (all): every table of the database as a sheet. */
export async function exportAll(app) {
  const db = app.db;
  if (!db.tables.length) { app.toast(t('Base.Ext.NoTables'), 'info'); return; }
  const content = emptyContent(db);
  for (const tb of db.tables) content.tables.push(tableOf(new TableRecordset(app, tb.id), tb.name));
  content.tables.forEach((tb, i) => { tb.id = 't' + (i + 1); });
  const name = dbName(app);
  try {
    const blob = await convert(app, content, name);
    download(name + '.xlsx', blob);
  } catch (e) {
    await say(e.body?.error || (e.status === 0 ? t('Base.Ext.Offline') : t('Base.Ext.ExportFailed')));
  }
}

/** Export → Text File: CSV with the separator of the language (";" in German, "," in English), UTF-8 with BOM so Excel reads the umlauts. */
export async function exportCsv(app, kind, id) {
  const obj = (kind === 'tables' ? tableById : queryById)(app.db, id);
  if (!obj) return;
  const rs = await recordsOf(app, kind, id);
  if (!rs) return;
  const sep = app.lang === 'de' ? ';' : ',';
  const quote = s => (/["\r\n]/.test(s) || s.includes(sep) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
  const lines = [rs.columns.map(c => quote(c.name)).join(sep)];
  for (const r of rs.rows) {
    lines.push(rs.columns.map(c => {
      const v = rs.get(r, c);
      if (v === null || v === undefined) return '';
      const s = c.type === 'datetime' || c.type === 'number' || c.type === 'currency' || c.type === 'autonumber' ? rs.editText(c, v) : rs.display(c, v);
      return quote(String(s));
    }).join(sep));
  }
  download(obj.name + '.csv', '﻿' + lines.join('\r\n') + '\r\n', 'text/csv;charset=utf-8');
  app.toast(t('Base.Ext.Exported', obj.name, rs.rows.length), 'success');
}

export { problemText };
