/* MatOffice Sheets – how a cell looks (pure, no DOM): the formatted text, its colour and alignment, the font, fill and borders of its style,
   and what conditional formats add (colours, data bars, colour scales). The grid, the print view, the clipboard (HTML) and the picture of the
   gallery all draw from here. Also the checks of data validation. */
import { formatValue } from './engine/numfmt.js';
import { FormulaError, compare, compareText, parseNumber, toText, numberToText } from './engine/values.js';
import { parseDate } from './engine/dates.js';
import { parseArea, areaContains, key } from './engine/address.js';
import { toLocal } from './engine/lexer.js';
import { moveFormula } from './engine/transform.js';

export const DEFAULT_FONT = 'Calibri';
export const DEFAULT_SIZE = 11;

/** The font families of the font list (the first ones are the office fonts with free look-alikes as fallback). */
export const FONTS = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Candara', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

const FALLBACK = {
  Calibri: 'Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  Aptos: 'Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  Cambria: 'Cambria, Caladea, Georgia, "Times New Roman", serif',
  Consolas: 'Consolas, "Courier New", monospace',
  'Courier New': '"Courier New", Courier, monospace',
  'Times New Roman': '"Times New Roman", Times, serif',
  Georgia: 'Georgia, "Times New Roman", serif',
};

/** The CSS font-family of a font name (with fallbacks). */
export const fontFamily = name => FALLBACK[name] || `"${name || DEFAULT_FONT}", ${FALLBACK.Calibri}`;

/** A CSS/canvas font: "italic bold 14.67px Calibri, …". */
export function cssFont(style, zoom = 1) {
  const px = (style.size || DEFAULT_SIZE) * 96 / 72 * zoom;
  return `${style.i ? 'italic ' : ''}${style.b ? 'bold ' : ''}${Math.round(px * 100) / 100}px ${fontFamily(style.font)}`;
}

/** Whether a colour is dark (for the text on a fill when the text colour is automatic). */
export function isDark(hex) {
  if (!hex || hex.length !== 7) return false;
  const n = parseInt(hex.slice(1), 16);
  const r = n >> 16;
  const g = (n >> 8) & 255;
  const b = n & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}

/** A border "thin #000000" → { width, style, color }. */
export function parseBorder(text) {
  if (!text) return null;
  const [style, color = '#000000'] = text.split(' ');
  const width = style === 'medium' ? 2 : style === 'thick' ? 3 : style === 'double' ? 3 : 1;
  return { style, width, color };
}

export class Look {
  constructor(app) {
    this.app = app;
    this.cache = new Map();
    this.rules = new Map();
  }

  invalidate() {
    this.cache.clear();
    this.rules.clear();
  }

  /**
   * The look of a cell: { value, text, color, fill, style, ha, va, wrap, indent, isNumber, fillAt, bar, link, note, formula }.
   * chars: the width of the column in characters (General shortens numbers to it).
   */
  cell(sheet, row, col, chars = 11) {
    const k = sheet.id + ':' + key(row, col) + ':' + chars;
    let look = this.cache.get(k);
    if (look) return look;
    const app = this.app;
    const raw = sheet.cell(row, col);
    const styleIndex = sheet.styleAt(row, col);
    const style = app.wb.style(styleIndex);
    const value = app.calc.valueAt(sheet.id, row, col);
    let text;
    let color = null;
    let fillAt = -1;
    const showFormula = app.showFormulas && raw?.f;
    if (showFormula) {
      text = '=' + toLocal(raw.f, app.locale);
    } else {
      const r = formatValue(value instanceof FormulaError ? value : value, style.nf, { locale: app.locale, width: chars });
      text = r.text;
      color = r.color;
      fillAt = r.fill;
    }
    const isNumber = typeof value === 'number' && !showFormula;
    const ha = style.ha || (showFormula ? 'left' : isNumber ? 'right' : typeof value === 'boolean' || value instanceof FormulaError ? 'center' : 'left');
    look = {
      value, text, style, styleIndex, color: color || style.color || null, fill: style.fill || null,
      ha, va: style.va || 'bottom', wrap: !!style.wrap, indent: style.indent || 0, isNumber,
      fillAt, link: raw?.link || null, note: raw?.note || null, formula: raw?.f || null, bar: null, scale: null,
      b: !!style.b, i: !!style.i, u: !!style.u, strike: !!style.strike,
    };
    if (raw?.link && !style.color) { look.color = '#0563c1'; look.u = true; }
    this.applyRules(sheet, row, col, look);
    if (this.cache.size > 50000) this.cache.clear();
    this.cache.set(k, look);
    return look;
  }

  // ---- conditional formats -------------------------------------------------------------------------------------------------------

  prepared(sheet) {
    let rules = this.rules.get(sheet.id);
    if (rules) return rules;
    rules = [];
    for (const rule of sheet.cf || []) {
      const area = parseArea(rule.range);
      if (!area) continue;
      try { rules.push(this.prepare(sheet, rule, area)); } catch { /* a broken rule draws nothing */ }
    }
    this.rules.set(sheet.id, rules);
    return rules;
  }

  /** The numbers of an area (for top/bottom, averages, bars and scales). */
  numbersIn(sheet, a) {
    const out = [];
    const b = sheet.dataBounds();
    const calc = this.app.calc;
    const ext = calc.extent(sheet.id);
    const r2 = Math.min(a.r2, Math.max(b.rows, ext.rows) - 1);
    const c2 = Math.min(a.c2, Math.max(b.cols, ext.cols) - 1);
    for (let r = a.r1; r <= r2; r++) for (let c = a.c1; c <= c2; c++) { const v = calc.valueAt(sheet.id, r, c); if (typeof v === 'number') out.push(v); }
    return out;
  }

  prepare(sheet, rule, area) {
    const app = this.app;
    const operand = text => {
      if (text === undefined || text === null) return null;
      const t = String(text);
      if (t.startsWith('=')) return app.calc.evaluateText(t.slice(1), sheet.id, area.r1, area.c1);
      const n = parseNumber(t);
      return n !== null ? n : t;
    };
    const p = { area, rule, test: null };
    switch (rule.type) {
      case 'cell': {
        const v1 = operand(rule.v1);
        const v2 = operand(rule.v2);
        const lo = typeof v1 === 'number' && typeof v2 === 'number' ? Math.min(v1, v2) : v1;
        const hi = typeof v1 === 'number' && typeof v2 === 'number' ? Math.max(v1, v2) : v2;
        p.test = v => {
          if (v === null || v === undefined || v === '' || v instanceof FormulaError) return false;
          const c1 = compare(v, v1);
          switch (rule.op) {
            case 'gt': return c1 > 0;
            case 'ge': return c1 >= 0;
            case 'lt': return c1 < 0;
            case 'le': return c1 <= 0;
            case 'eq': return c1 === 0;
            case 'ne': return c1 !== 0;
            case 'between': return compare(v, lo) >= 0 && compare(v, hi) <= 0;
            case 'notBetween': return compare(v, lo) < 0 || compare(v, hi) > 0;
            default: return false;
          }
        };
        break;
      }
      case 'text': {
        const needle = String(rule.v1 ?? '').toLowerCase();
        p.test = v => {
          if (v === null || v === undefined || v instanceof FormulaError) return false;
          const s = (typeof v === 'string' ? v : toText(v)).toLowerCase();
          switch (rule.op) {
            case 'contains': return s.includes(needle);
            case 'notContains': return !s.includes(needle);
            case 'begins': return s.startsWith(needle);
            case 'ends': return s.endsWith(needle);
            default: return false;
          }
        };
        break;
      }
      case 'dup': case 'unique': {
        const counts = new Map();
        const keyOf = v => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v === 'number' ? 'n' + v : typeof v === 'boolean' ? 'b' + v : null);
        const b = sheet.dataBounds();
        for (let r = area.r1; r <= Math.min(area.r2, b.rows - 1); r++) {
          for (let c = area.c1; c <= Math.min(area.c2, b.cols - 1); c++) {
            const k = keyOf(app.calc.valueAt(sheet.id, r, c));
            if (k) counts.set(k, (counts.get(k) || 0) + 1);
          }
        }
        p.test = v => { const k = keyOf(v); return !!k && (rule.type === 'dup' ? counts.get(k) > 1 : counts.get(k) === 1); };
        break;
      }
      case 'top': case 'bottom': {
        const nums = this.numbersIn(sheet, area).sort((a, b) => (rule.type === 'top' ? b - a : a - b));
        const n = rule.percent ? Math.max(1, Math.floor(nums.length * (rule.n || 10) / 100)) : rule.n || 10;
        const threshold = nums.length ? nums[Math.min(n, nums.length) - 1] : null;
        p.test = v => typeof v === 'number' && threshold !== null && (rule.type === 'top' ? v >= threshold : v <= threshold);
        break;
      }
      case 'above': case 'below': {
        const nums = this.numbersIn(sheet, area);
        const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        p.test = v => typeof v === 'number' && (rule.type === 'above' ? v > avg : v < avg);
        break;
      }
      case 'formula': {
        const text = String(rule.v1 || '').replace(/^=/, '');
        p.test = (v, row, col) => {
          const f = moveFormula(text, row - area.r1, col - area.c1);
          const r = app.calc.evaluateText(f, sheet.id, row, col);
          return r === true || (typeof r === 'number' && r !== 0);
        };
        break;
      }
      case 'bar': {
        const nums = this.numbersIn(sheet, area);
        const min = Math.min(0, ...nums);
        const max = Math.max(...nums, 0);
        p.bar = v => (typeof v === 'number' && max > min ? { ratio: (v - min) / (max - min), color: rule.color || '#638ec6', negative: v < 0 } : null);
        break;
      }
      case 'scale': {
        const nums = this.numbersIn(sheet, area).sort((a, b) => a - b);
        const min = nums[0];
        const max = nums[nums.length - 1];
        const mid = nums.length ? nums[Math.floor((nums.length - 1) / 2)] : 0;
        const colors = rule.colors || ['#f8696b', '#63be7b'];
        p.scale = v => {
          if (typeof v !== 'number' || !nums.length) return null;
          if (max === min) return colors[colors.length - 1];
          if (colors.length === 2) return mix(colors[0], colors[1], (v - min) / (max - min));
          return v <= mid ? mix(colors[0], colors[1], mid === min ? 1 : (v - min) / (mid - min)) : mix(colors[1], colors[2], max === mid ? 1 : (v - mid) / (max - mid));
        };
        break;
      }
      default: break;
    }
    return p;
  }

  applyRules(sheet, row, col, look) {
    if (!sheet.cf?.length) return;
    const taken = new Set();
    for (const p of this.prepared(sheet)) {
      if (!areaContains(p.area, row, col)) continue;
      if (p.bar) { if (!look.bar) look.bar = p.bar(look.value); continue; }
      if (p.scale) { if (!taken.has('fill')) { const c = p.scale(look.value); if (c) { look.fill = c; taken.add('fill'); } } continue; }
      if (!p.test) continue;
      let hit = false;
      try { hit = p.test(look.value, row, col); } catch { hit = false; }
      if (!hit) continue;
      const s = p.rule.style || {};
      for (const prop of ['color', 'fill', 'b', 'i', 'u', 'strike']) {
        if (s[prop] === undefined || taken.has(prop)) continue;
        look[prop] = s[prop];
        taken.add(prop);
      }
    }
  }
}

/** The colour between two colours (t from 0 to 1). */
export function mix(a, b, t) {
  const x = parseInt(a.slice(1), 16);
  const y = parseInt(b.slice(1), 16);
  const k = Math.max(0, Math.min(1, t));
  const ch = s => Math.round(((x >> s) & 255) * (1 - k) + ((y >> s) & 255) * k);
  return '#' + [16, 8, 0].map(s => ch(s).toString(16).padStart(2, '0')).join('');
}

// ---- data validation -----------------------------------------------------------------------------------------------------------

/** The validation rule of a cell (the last one that covers it), or null. */
export function validationAt(sheet, row, col) {
  const list = sheet.validations || [];
  for (let i = list.length - 1; i >= 0; i--) {
    const a = parseArea(list[i].range);
    if (a && areaContains(a, row, col)) return list[i];
  }
  return null;
}

/** The items of a list validation ("a,b,c" or a range "=$E$1:$E$5"). */
export function listItems(app, sheet, rule) {
  const v1 = String(rule.v1 || '');
  if (v1.startsWith('=')) {
    const r = app.calc.evaluateText(v1.slice(1), sheet.id, 0, 0);
    const flat = Array.isArray(r) ? r.flat() : [r];
    return flat.filter(v => v !== null && v !== undefined && v !== '' && !(v instanceof FormulaError)).map(v => (typeof v === 'number' ? numberToText(v) : String(v)));
  }
  return v1.split(app.locale === 'de' && v1.includes(';') ? ';' : ',').map(s => s.trim()).filter(Boolean);
}

/** Whether a value passes the validation of a cell. */
export function validate(app, sheet, row, col, value, rule = validationAt(sheet, row, col)) {
  if (!rule) return true;
  const blank = value === null || value === undefined || value === '';
  if (blank) return rule.blank !== false;
  const num = t => {
    if (t === undefined || t === null) return null;
    const s = String(t);
    if (s.startsWith('=')) { const r = app.calc.evaluateText(s.slice(1), sheet.id, row, col); return typeof r === 'number' ? r : null; }
    const n = parseNumber(s);
    if (n !== null) return n;
    const d = parseDate(s, app.locale);
    return d ? d.value : null;
  };
  switch (rule.type) {
    case 'list': {
      const items = listItems(app, sheet, rule);
      const text = typeof value === 'number' ? numberToText(value) : String(value);
      return items.some(i => compareText(i, text) === 0);
    }
    case 'custom': {
      const a = parseArea(rule.range);
      const f = moveFormula(String(rule.v1).replace(/^=/, ''), row - a.r1, col - a.c1);
      const r = app.calc.evaluateText(f, sheet.id, row, col);
      return r === true || (typeof r === 'number' && r !== 0);
    }
    default: {
      let v;
      if (rule.type === 'length') v = String(typeof value === 'number' ? numberToText(value) : value).length;
      else if (typeof value === 'number') v = value;
      else return false;
      if (rule.type === 'whole' && !Number.isInteger(v)) return false;
      if (rule.type === 'time') v = v % 1;
      const a = num(rule.v1);
      const b = num(rule.v2);
      if (a === null) return true;
      switch (rule.op || 'between') {
        case 'between': return v >= Math.min(a, b ?? a) && v <= Math.max(a, b ?? a);
        case 'notBetween': return v < Math.min(a, b ?? a) || v > Math.max(a, b ?? a);
        case 'eq': return v === a;
        case 'ne': return v !== a;
        case 'gt': return v > a;
        case 'ge': return v >= a;
        case 'lt': return v < a;
        case 'le': return v <= a;
        default: return true;
      }
    }
  }
}
