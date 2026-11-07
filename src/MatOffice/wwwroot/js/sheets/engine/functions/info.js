/* MatOffice Sheets – information functions (pure, no DOM). The IS functions and TYPE see errors as values. */
import { define, scalar, deref, Errors, Ref } from '../lib.js';
import { FormulaError, toNumber, toText } from '../values.js';
import { colName, sheetPrefix } from '../address.js';

const isErr = v => v instanceof FormulaError;
const { NA, VALUE } = Errors;
const first = v => (Array.isArray(v) ? v[0][0] : v);

/** A function of one value (errors included), lifted over arrays: =ISBLANK(A1:A3) spills. */
const is = (name, f, args = 'value') => define(name, { min: 1, max: 1, category: 'info', args, fn: scalar(['any'], f) });

is('ISBLANK', v => v === null);
is('ISNUMBER', v => typeof v === 'number');
is('ISTEXT', v => typeof v === 'string');
is('ISNONTEXT', v => typeof v !== 'string');
is('ISLOGICAL', v => typeof v === 'boolean');
is('ISERROR', isErr);
is('ISERR', v => isErr(v) && v.code !== NA.code);
is('ISNA', v => isErr(v) && v.code === NA.code);

/** ISEVEN/ISODD: the number truncated; logicals and text that is no number are #VALUE!. */
const parity = odd => v => {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return VALUE;
  const n = toNumber(v);
  return isErr(n) ? n : Math.abs(Math.trunc(n)) % 2 === (odd ? 1 : 0);
};
is('ISEVEN', parity(false), 'number');
is('ISODD', parity(true), 'number');

is('N', v => (isErr(v) ? v : typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0));

/** The numbers of ERROR.TYPE (#SPILL! 9 and #CALC! 14 like Excel 365). */
const ERROR_TYPES = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 };
is('ERROR.TYPE', v => (isErr(v) && ERROR_TYPES[v.code] ? ERROR_TYPES[v.code] : NA), 'error_val');

define('NA', { min: 0, max: 0, category: 'info', args: '', fn: () => NA });
define('ISREF', { min: 1, max: 1, category: 'info', args: 'value', fn: ([v]) => v instanceof Ref });
define('ISFORMULA', {
  min: 1, max: 1, category: 'info', args: 'reference',
  fn: ([v], ctx) => (isErr(v) ? v : v instanceof Ref ? !!ctx.formulaAt(v.sheet, v.r1, v.c1) : VALUE),
});

/** TYPE: 1 number (and blank), 2 text, 4 logical, 16 error, 64 array (also a range of several cells). */
define('TYPE', {
  min: 1, max: 1, category: 'info', args: 'value',
  fn: ([v], ctx) => {
    if ((v instanceof Ref && !v.isCell) || Array.isArray(v)) return 64;
    const x = deref(v, ctx);
    return isErr(x) ? 16 : typeof x === 'string' ? 2 : typeof x === 'boolean' ? 4 : 1;
  },
});

/** SHEET([value]): the number of the formula's sheet, of a reference's sheet or of a sheet named by text (#N/A when there is none). */
define('SHEET', {
  min: 0, max: 1, category: 'info', args: '[value]',
  fn: ([v], ctx) => {
    if (v === undefined || v === null) return ctx.sheetIndex(ctx.sheet) + 1;
    if (v instanceof Ref) return ctx.sheetIndex(v.sheet) + 1;
    if (isErr(v)) return v;
    const name = first(v);
    if (typeof name !== 'string') return NA;
    const id = ctx.sheetId(name);
    return id === null || id === undefined ? NA : ctx.sheetIndex(id) + 1;
  },
});

define('SHEETS', {
  min: 0, max: 1, category: 'info', args: '[reference]',
  fn: ([v], ctx) => (v === undefined || v === null ? ctx.sheetCount() : isErr(v) ? v : v instanceof Ref ? 1 : NA),
});

/** CELL(info_type, [reference]) for "address", "col", "row", "contents" and "type" (b blank, l text, v anything else). */
define('CELL', {
  min: 1, max: 2, category: 'info', args: 'info_type, [reference]',
  fn: ([info, ref], ctx) => {
    const type = toText(first(deref(info, ctx)));
    if (isErr(type)) return type;
    if (isErr(ref)) return ref;
    if (ref !== undefined && ref !== null && !(ref instanceof Ref)) return VALUE;
    const r = ref instanceof Ref ? ref : { sheet: ctx.sheet, r1: ctx.row, c1: ctx.col };
    const value = () => ctx.value(r.sheet, r.r1, r.c1);
    switch (type.toLowerCase()) {
      case 'address': return (r.sheet === ctx.sheet ? '' : sheetPrefix(ctx.sheetName(r.sheet))) + '$' + colName(r.c1) + '$' + (r.r1 + 1);
      case 'col': return r.c1 + 1;
      case 'row': return r.r1 + 1;
      case 'contents': return value();
      case 'type': { const v = value(); return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v'; }
      default: return VALUE;
    }
  },
});
