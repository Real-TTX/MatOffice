/* MatOffice Base – layouts of forms and reports without the DOM: what the Form and Report buttons and the wizards build (the same layouts the
   server makes for the templates, BaseLayout.cs), the bands of a report (headers and footers of the groups, details, totals) and their pages,
   and the values of controls (=[Price]*[Qty], =Sum([Amount]), ="Page " & [Page] & " of " & [Pages]). */
import { Parser, compile, BaseError, sortCompare, isNull } from './expr.js';
import { aggregate, runQueryObject, runSelect } from './engine.js';
import { keyOf, lower } from './model.js';

export const PAPER = { a4: { w: 794, h: 1123 }, letter: { w: 816, h: 1056 } };
export const MARGIN = 38;

export function pageSize(report) {
  const p = PAPER[report.paper] || PAPER.a4;
  return report.orientation === 'landscape' ? { w: p.h, h: p.w } : { w: p.w, h: p.h };
}

export const printableWidth = report => pageSize(report).w - 2 * MARGIN;

let n = 0;
const control = (type, x, y, w, h, more = {}) => ({ id: 'c' + (++n), type, x, y, w, h, ...more });
const section = (height, controls = []) => ({ height, visible: true, controls });
const numeric = f => f.type === 'number' || f.type === 'currency' || f.type === 'autonumber';
const caption = f => f.caption || f.name;
const controlType = f => (f.type === 'yesno' ? 'checkbox' : f.lookup ? 'combobox' : 'textbox');

export function controlWidth(f) {
  if (f.type === 'memo') return 400;
  if (f.type === 'number' || f.type === 'autonumber') return 120;
  if (f.type === 'currency' || f.type === 'datetime') return 140;
  if (f.type === 'yesno') return 24;
  if (f.type === 'hyperlink') return 300;
  return f.lookup ? 240 : 260;
}

function renumber(obj) {
  let k = 0;
  const visit = s => { if (s) for (const c of s.controls) c.id = 'c' + (++k); };
  const s = obj.sections;
  [s.header, s.footer, s.reportHeader, s.pageHeader, ...(s.groupHeaders || []), s.detail, ...(s.groupFooters || []), s.pageFooter, s.reportFooter].forEach(visit);
  return obj;
}

/**
 * The fields of a record source (a table or a query by id) as the layouts, the designer and the wizards need them:
 * [{ name, caption, type, lookup, field, format }] (a query's columns come from a run without its parameters).
 */
export function sourceFieldList(db, sourceId, lang = 'en') {
  const table = db.tables.find(x => x.id === sourceId);
  if (table) return table.fields.map(f => ({ name: f.name, caption: f.caption || '', type: f.type, lookup: !!f.lookup, field: f, format: f.format || '' }));
  const query = db.queries.find(x => x.id === sourceId);
  if (!query) return [];
  try {
    return runQueryObject(db, query, { lang, collect: new Map() }).columns.map(c => ({
      name: c.name, caption: c.field && lower(c.field.name) === lower(c.name) ? c.field.caption || '' : '', type: c.type || 'text', lookup: !!c.field?.lookup, field: c.field || null, format: c.field?.format || '',
    }));
  } catch {
    return [];
  }
}

/**
 * A columnar form (the Form button): labels left, controls right, the title in the header; optional navigation buttons (texts.buttons) and a subform
 * { sourceId, caption, linkMaster, linkChild }.
 */
export function columnarForm({ id, name, sourceId, title, fields, subform = null, buttons = null, kind = 'single' }) {
  const width = subform ? 720 : 640;
  const header = [control('label', 24, 18, width - 48, 38, { caption: title, fontSize: 22 })];
  const detail = [];
  let y = 16;
  for (const f of fields) {
    const h = f.type === 'memo' ? 72 : 24;
    detail.push(control('label', 24, y, 150, 24, { caption: caption(f) }));
    detail.push(control(controlType(f), 184, y, Math.min(controlWidth(f), width - 208), h, { source: f.name, name: f.name, align: numeric(f) ? 'right' : 'left', locked: f.type === 'autonumber' }));
    y += h + 10;
  }
  if (subform) {
    y += 6;
    detail.push(control('label', 24, y, 300, 24, { caption: subform.caption, bold: true }));
    y += 28;
    detail.push(control('subform', 24, y, width - 48, 240, { name: subform.caption, subform: { source: subform.sourceId, linkMaster: subform.linkMaster, linkChild: subform.linkChild } }));
    y += 250;
  }
  const footer = [];
  if (buttons) {
    let x = 24;
    for (const [action, label] of buttons) { footer.push(control('button', x, 12, 96, 30, { caption: label, action })); x += 104; }
  }
  return renumber({
    id, name, source: sourceId, caption: title, kind, width, allowEdits: true, allowAdditions: true, allowDeletions: true,
    sections: { header: section(72, header), detail: section(y + 8, detail), footer: section(buttons ? 54 : 0, footer) },
  });
}

/** A Multiple Items form: the labels in the header, one row of controls per record. */
export function tabularForm({ id, name, sourceId, title, fields }) {
  const labels = [];
  const detail = [];
  let x = 24;
  for (const f of fields) {
    const w = f.type === 'memo' ? 220 : numeric(f) ? (f.type === 'currency' ? 110 : 90) : f.type === 'datetime' ? 110 : f.type === 'yesno' ? 70 : 160;
    labels.push(control('label', x, 80, w, 24, { caption: caption(f), bold: true }));
    detail.push(control(controlType(f), f.type === 'yesno' ? x + w / 2 - 12 : x, 4, f.type === 'yesno' ? 24 : w, 24, { source: f.name, name: f.name, align: numeric(f) ? 'right' : 'left' }));
    x += w + 8;
  }
  const width = Math.max(640, x + 16);
  return renumber({
    id, name, source: sourceId, caption: title, kind: 'continuous', width, allowEdits: true, allowAdditions: true, allowDeletions: true,
    sections: { header: section(110, [control('label', 24, 18, width - 48, 38, { caption: title, fontSize: 22 }), ...labels]), detail: section(32, detail), footer: section(0) },
  });
}

/** An empty form or report to design from scratch. */
export function blankForm({ id, name, sourceId = '' }) {
  return { id, name, source: sourceId, caption: name, kind: 'single', width: 640, allowEdits: true, allowAdditions: true, allowDeletions: true, sections: { header: section(0), detail: section(320), footer: section(0) } };
}

export function blankReport({ id, name, sourceId = '', texts }) {
  const width = 718;
  return renumber({
    id, name, source: sourceId, caption: name, orientation: 'portrait', paper: 'a4', width, grouping: [], sorting: [],
    sections: {
      reportHeader: section(0), pageHeader: section(30), groupHeaders: [], detail: section(200), groupFooters: [],
      pageFooter: section(32, [control('textbox', width - 260, 8, 260, 20, { source: pageExpression(texts), align: 'right', border: false, fontSize: 9 })]), reportFooter: section(0),
    },
  });
}

export const pageExpression = texts => '="' + texts.page.replace(/"/g, '""') + ' " & [Page] & " ' + texts.of.replace(/"/g, '""') + ' " & [Pages]';

const weight = f => (f.type === 'memo' ? 5 : f.type === 'number' || f.type === 'autonumber' ? 1.4 : f.type === 'currency' || f.type === 'datetime' ? 1.8 : f.type === 'yesno' ? 1 : 3);

/**
 * A report of the wizard: layout 'stepped' (the group values in their own columns), 'block' (the group values in the detail line) or 'outline'
 * (the group value in a header line with its label); sums: names of numeric fields with Sum (summary: { [name]: ['Sum','Avg','Min','Max'] }).
 * texts: { summaryFor, grandTotal, records, page, of }.
 */
export function wizardReport({ id, name, sourceId, title, fields, groups = [], sort = [], summary = {}, landscape = false, layout = 'stepped', texts }) {
  const width = landscape ? 1047 : 718;
  const groupFields = groups.map(g => fields.find(f => f.name === g)).filter(Boolean);
  const detailFields = fields.filter(f => !groups.includes(f.name));
  const block = layout === 'block';
  const outline = layout === 'outline';
  const columnsOf = block ? [...groupFields, ...detailFields] : outline ? detailFields : [...groupFields, ...detailFields];
  const total = columnsOf.reduce((s, f) => s + weight(f), 0) || 1;
  const gap = 8;
  const usable = width - gap * Math.max(0, columnsOf.length - 1);
  const cols = new Map();
  let x = 0;
  for (const f of columnsOf) { const w = Math.floor(usable * weight(f) / total); cols.set(f.name, { x, w }); x += w + gap; }
  const align = f => (numeric(f) ? 'right' : 'left');
  const fns = Object.entries(summary).flatMap(([field, list]) => list.map(fn => ({ field, fn })));
  const fnRows = [...new Set(fns.map(x => x.fn))];

  const reportHeader = [
    control('label', 0, 6, Math.round(width * 0.65), 40, { caption: title, fontSize: 22 }),
    control('textbox', width - 220, 14, 220, 22, { source: '=Date()', format: 'Long Date', align: 'right', border: false }),
  ];
  const pageHeader = outline ? [] : columnsOf.map(f => control('label', cols.get(f.name).x, 6, cols.get(f.name).w, 22, { caption: caption(f), bold: true, align: align(f) }));
  if (!outline) pageHeader.push(control('line', 0, 30, width, 1));
  const summaryControls = (yStart, label) => {
    const out = [];
    let y = yStart;
    for (const fn of fnRows) {
      out.push(control('label', 0, y, 160, 20, { caption: label(fn), italic: true }));
      for (const s of fns.filter(x => x.fn === fn)) {
        const c = cols.get(s.field) || { x: width - 160, w: 160 };
        const f = fields.find(x => x.name === s.field);
        out.push(control('textbox', c.x, y, c.w, 20, { source: '=' + s.fn + '([' + s.field + '])', bold: true, border: false, align: 'right', format: f?.type === 'currency' && s.fn !== 'Count' ? 'Currency' : s.fn === 'Avg' ? 'Standard' : '' }));
      }
      y += 22;
    }
    return { controls: out, height: y };
  };
  const fnLabel = fn => ({ Sum: texts.sum, Avg: texts.avg, Min: texts.min, Max: texts.max }[fn] || fn);
  const groupHeaders = [];
  const groupFooters = [];
  groupFields.forEach((g, i) => {
    const c = cols.get(g.name) || { x: 0, w: 300 };
    if (outline) {
      groupHeaders.push(section(54, [
        control('label', i * 24, 6, 140, 22, { caption: caption(g), bold: true }),
        control('textbox', i * 24 + 150, 6, 300, 22, { source: g.name, bold: true, fontSize: 12, border: false }),
        ...(i === groupFields.length - 1 ? detailFields.map(f => control('label', cols.get(f.name).x, 32, cols.get(f.name).w, 20, { caption: caption(f), bold: true, align: align(f) })) : []),
        ...(i === groupFields.length - 1 ? [control('line', 0, 52, width, 1)] : []),
      ]));
    } else if (block) {
      groupHeaders.push(section(0));
    } else {
      groupHeaders.push(section(30, [control('textbox', c.x, 4, c.w, 22, { source: g.name, bold: true, fontSize: 12, border: false })]));
    }
    if (fns.length) {
      const head = control('textbox', 0, 4, Math.max(300, c.w), 20, { source: '="' + texts.summaryFor.replace(/"/g, '""') + ' " & [' + g.name + ']', italic: true, border: false });
      const lines = fns.map(s => { const sc = cols.get(s.field) || { x: width - 160, w: 160 }; return control('line', sc.x, 2, sc.w, 1); });
      const sum = summaryControls(26, fnLabel);
      groupFooters.push(section(sum.height + 6, [head, ...lines, ...sum.controls]));
    } else groupFooters.push(section(0));
  });
  const detail = [];
  for (const f of block ? columnsOf : detailFields) {
    const c = cols.get(f.name);
    detail.push(control(f.type === 'yesno' ? 'checkbox' : 'textbox', c.x, 2, f.type === 'yesno' ? 18 : c.w, f.type === 'yesno' ? 18 : 20, {
      source: f.name, border: false, align: align(f), hideDuplicates: block && groups.includes(f.name),
    }));
  }
  let reportFooter;
  if (fns.length) {
    const sum = summaryControls(10, fn => texts.grandTotal + (fnRows.length > 1 ? ' (' + fnLabel(fn) + ')' : ''));
    reportFooter = section(sum.height + 10, [...fns.map(s => { const sc = cols.get(s.field) || { x: 0, w: 100 }; return control('line', sc.x, 4, sc.w, 1); }), ...sum.controls]);
  } else reportFooter = section(40, [control('textbox', 0, 10, 300, 22, { source: '=Count(*) & " ' + texts.records.replace(/"/g, '""') + '"', bold: true, border: false })]);
  const pageFooter = section(32, [
    control('textbox', 0, 8, 260, 20, { source: '=Now()', border: false, fontSize: 9 }),
    control('textbox', width - 260, 8, 260, 20, { source: pageExpression(texts), align: 'right', border: false, fontSize: 9 }),
  ]);
  return renumber({
    id, name, source: sourceId, caption: title, orientation: landscape ? 'landscape' : 'portrait', paper: 'a4', width,
    grouping: groups.map(g => ({ field: g, header: true, footer: fns.length > 0, desc: false, keepTogether: false })),
    sorting: sort.map(s => (typeof s === 'string' ? { field: s, desc: false } : s)),
    sections: { reportHeader: section(56, reportHeader), pageHeader: section(outline ? 0 : 34, pageHeader), groupHeaders, detail: section(24, detail), groupFooters, pageFooter, reportFooter },
  });
}

// ---- values of controls -------------------------------------------------------------------------------------------------------------

/**
 * Evaluates control sources against a recordset: a field name gives the value of the row, "=…" an expression with the row's fields, aggregates over
 * the rows of the band (a group, the report), [Page] and [Pages]. Compiled once per source.
 */
export class ControlEval {
  constructor(rs, { lang = 'en', currency = 'EUR', domain = null } = {}) {
    this.rs = rs;
    this.lang = lang;
    this.currency = currency;
    this.domain = domain;
    this.cache = new Map();
    this.byName = new Map();
    for (const c of rs.columns) { this.byName.set(lower(c.name), c); if (!this.byName.has(lower(c.caption))) this.byName.set(lower(c.caption), c); }
    this.env = { row: null, rows: [], page: { page: 1, pages: 1 } };
  }

  column(name) { return this.byName.get(lower(String(name || '').replace(/^\[|\]$/g, ''))) || null; }

  compile(source) {
    if (this.cache.has(source)) return this.cache.get(source);
    let fn;
    try {
      const text = source.startsWith('=') ? source.slice(1) : '[' + source.replace(/^\[|\]$/g, '') + ']';
      const ast = Parser.expression(text);
      const env = this.env;
      const rowScope = {
        lang: this.lang, currency: this.currency, domain: this.domain, page: env.page,
        resolve: parts => { const c = this.column(parts[parts.length - 1]); return c ? x => (x?.row ? this.rs.get(x.row, c) : null) : null; },
      };
      const scope = {
        ...rowScope,
        aggregate: (name, arg) => {
          if (arg.t === 'star') return x => (x?.rows || []).length;
          const f = compile(arg, rowScope);
          return x => aggregate(name, (x?.rows || []).map(r => f({ row: r })), this.lang);
        },
      };
      const compiled = compile(ast, scope);
      fn = x => compiled(x);
      fn.ok = true;
    } catch (e) {
      const err = e instanceof BaseError ? e : new BaseError('Syntax', source);
      fn = () => { throw err; };
      fn.ok = false;
    }
    this.cache.set(source, fn);
    return fn;
  }

  /** The value of a control source for a row (and the rows of its band). Throws BaseError for #Name? and #Error. */
  value(source, row, rows = null, page = null) {
    if (!source) return null;
    if (page) { this.env.page.page = page.page; this.env.page.pages = page.pages; }
    return this.compile(source)({ row, rows: rows || (row ? [row] : []) });
  }

  /** The text a control shows: its format, the display value of a lookup, #Name? / #Error for a broken source. */
  text(ctl, row, rows = null, page = null) {
    if (!ctl.source) return '';
    let v;
    try { v = this.value(ctl.source, row, rows, page); } catch (e) { return e instanceof BaseError && (e.code === 'UnknownName' || e.code === 'Syntax') ? '#Name?' : '#Error'; }
    const col = ctl.source.startsWith('=') ? null : this.column(ctl.source);
    if (col) return this.rs.display(ctl.format ? { ...col, format: ctl.format, decimals: ctl.decimals ?? col.decimals } : col, v);
    if (v === null || v === undefined) return '';
    return this.rs.display({ key: '__expr', type: inferType(v), field: null, format: ctl.format || '', decimals: ctl.decimals ?? null }, v);
  }
}

/** DLookup, DSum, DCount … of control sources: a small query of their own on the database. */
export function domainFunction(db, lang = 'en') {
  return (fn, exprText, domainName, criteria) => {
    const agg = { DLookup: 'First', DCount: 'Count', DSum: 'Sum', DAvg: 'Avg', DMin: 'Min', DMax: 'Max', DFirst: 'First', DLast: 'Last' }[fn];
    const e = String(exprText).trim() === '*' ? { t: 'star', table: null } : Parser.expression(exprText);
    const where = criteria ? Parser.expression(criteria) : null;
    const res = runSelect(db, {
      type: 'select', distinct: null, top: null, columns: [{ e: { t: 'fn', name: agg, args: [e] }, alias: 'X' }], into: null,
      from: [{ t: 'table', name: domainName, alias: null }], where, groupBy: [], having: null, orderBy: [], params: [],
    }, { lang });
    const v = res.rows[0]?.v[0] ?? null;
    return fn === 'DCount' && v === null ? 0 : v;
  };
}

function inferType(v) {
  if (typeof v === 'boolean') return 'yesno';
  if (typeof v === 'number') return 'number';
  if (v && typeof v === 'object' && 's' in v) return 'datetime';
  return 'text';
}

// ---- bands and pages of a report ------------------------------------------------------------------------------------------------------

/** Sorts the rows like the grouping and sorting of the report say. */
export function sortRows(report, rs, ev) {
  const keys = [...(report.grouping || []).map(g => ({ field: g.field, desc: g.desc })), ...(report.sorting || [])].filter(k => k.field);
  if (!keys.length) return rs.rows.slice();
  const getters = keys.map(k => row => { try { return ev.value(k.field, row); } catch { return null; } });
  const keyed = rs.rows.map((r, i) => ({ r, i, k: getters.map(g => g(r)) }));
  keyed.sort((a, b) => {
    for (let j = 0; j < keys.length; j++) {
      const c = sortCompare(a.k[j], b.k[j], ev.lang);
      if (c) return keys[j].desc ? -c : c;
    }
    return a.i - b.i;
  });
  return keyed.map(x => x.r);
}

/**
 * The bands of a report in printing order: { kind: reportHeader|groupHeader|detail|groupFooter|reportFooter, level, section, row, rows }.
 * rows: the records of the band's scope (the group, the whole report) for aggregates.
 */
export function buildBands(report, rows, ev) {
  const s = report.sections;
  const groups = report.grouping || [];
  const bands = [];
  const keyAt = (row, level) => { try { return keyOf(ev.value(groups[level].field, row)) ?? '\u0000'; } catch { return '\u0000'; } };
  const visible = sec => sec && sec.visible !== false && sec.height > 0;
  if (visible(s.reportHeader)) bands.push({ kind: 'reportHeader', section: s.reportHeader, row: rows[0] || null, rows });
  const open = [];
  const closeFrom = level => {
    for (let l = open.length - 1; l >= level; l--) {
      const g = open[l];
      if (groups[l].footer !== false && visible(s.groupFooters?.[l])) bands.push({ kind: 'groupFooter', level: l, section: s.groupFooters[l], row: g.rows[g.rows.length - 1], rows: g.rows });
      open.pop();
    }
  };
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    let changed = -1;
    for (let l = 0; l < groups.length; l++) {
      if (!open[l] || open[l].key !== keyAt(row, l)) { changed = l; break; }
    }
    if (changed >= 0) {
      closeFrom(changed);
      for (let l = changed; l < groups.length; l++) {
        const key = keyAt(row, l);
        let j = i;
        while (j < rows.length && keyAt(rows[j], l) === key && (l === 0 || open.slice(0, l).every((o, k) => keyAt(rows[j], k) === o.key))) j++;
        const g = { key, rows: rows.slice(i, j) };
        open.push(g);
        if (groups[l].header !== false && visible(s.groupHeaders?.[l])) bands.push({ kind: 'groupHeader', level: l, section: s.groupHeaders[l], row, rows: g.rows, keep: true });
      }
    }
    if (visible(s.detail)) bands.push({ kind: 'detail', section: s.detail, row, rows: open.length ? open[open.length - 1].rows : rows, index: i });
  }
  closeFrom(0);
  if (visible(s.reportFooter)) bands.push({ kind: 'reportFooter', section: s.reportFooter, row: rows[rows.length - 1] || null, rows });
  return bands;
}

/** Lays the bands out on pages: [{ bands: [{ band, y }] }] (y below the top margin and the page header). */
export function paginate(report, bands) {
  const { h } = pageSize(report);
  const s = report.sections;
  const head = s.pageHeader?.visible !== false ? s.pageHeader?.height || 0 : 0;
  const foot = s.pageFooter?.visible !== false ? s.pageFooter?.height || 0 : 0;
  const room = Math.max(60, h - 2 * MARGIN - head - foot);
  const pages = [];
  let page = { bands: [] };
  let y = 0;
  bands.forEach((band, i) => {
    let need = band.section.height;
    if (band.keep && bands[i + 1]) need += bands[i + 1].section.height;
    const groupKeep = band.kind === 'groupHeader' && (report.grouping?.[band.level]?.keepTogether);
    if (groupKeep) need = Math.min(room, band.rows.length * (s.detail?.height || 0) + band.section.height);
    if (y > 0 && y + Math.min(need, room) > room) { pages.push(page); page = { bands: [] }; y = 0; }
    page.bands.push({ band, y });
    y += band.section.height;
  });
  pages.push(page);
  return pages;
}

export { isNull };
