/* MatOffice Sheets – the parser of formulas (pure, no DOM): tokens → syntax tree.
   Precedence like Excel (low → high): comparisons, &, + -, * /, ^ (left to right), %, unary minus, the range operator ":".
   Nodes: num, str, bool, err, empty (an omitted argument), ref, name, fn, unary, percent, binary, array, at. */
import { tokenize } from './lexer.js';

export class FormulaSyntaxError extends Error {
  constructor(message, pos) { super(message); this.pos = pos; }
}

const INFIX = {
  '=': 10, '<>': 10, '<': 10, '>': 10, '<=': 10, '>=': 10,
  '&': 20,
  '+': 30, '-': 30,
  '*': 40, '/': 40,
  '^': 50,
  '%': 60,
  ':': 80,
};
const UNARY = 70;

/** Parses a formula (without "="): { ast, tokens }; throws FormulaSyntaxError for a formula that is not complete. */
export function parse(text, options = {}) {
  const tokens = tokenize(text, options);
  const list = tokens.filter(t => t.type !== 'ws');
  let i = 0;
  const peek = () => list[i];
  const next = () => list[i++];
  const fail = (t, what) => { throw new FormulaSyntaxError(what || 'syntax', t ? t.pos : text.length); };

  const expr = bp => {
    let left = prefix();
    for (;;) {
      const t = peek();
      if (!t || t.type !== 'op') break;
      const p = INFIX[t.op];
      if (p === undefined || p <= bp) break;
      i++;
      if (t.op === '%') { left = { type: 'percent', arg: left }; continue; }
      const right = expr(p);
      left = { type: 'binary', op: t.op, left, right };
    }
    return left;
  };

  const prefix = () => {
    const t = next();
    if (!t) fail(null, 'end');
    switch (t.type) {
      case 'num': return { type: 'num', value: t.value };
      case 'str':
        if (t.closed === false) fail(t, 'string');
        return { type: 'str', value: t.value };
      case 'bool': return { type: 'bool', value: t.value };
      case 'err': return { type: 'err', code: t.code };
      case 'ref': return refNode(t);
      case 'name': return { type: 'name', name: t.name };
      case 'func': return call(t);
      case '(': {
        const e = expr(0);
        if (peek()?.type !== ')') fail(peek(), 'paren');
        i++;
        return e;
      }
      case '{': return array(t);
      case 'op':
        if (t.op === '-' || t.op === '+') return { type: 'unary', op: t.op, arg: expr(UNARY) };
        if (t.op === '@') return { type: 'at', arg: expr(UNARY) };
        return fail(t);
      default: return fail(t);
    }
  };

  const call = t => {
    if (next()?.type !== '(') fail(t, 'paren');
    const args = [];
    if (peek()?.type === ')') { i++; return { type: 'fn', name: t.name, args }; }
    for (;;) {
      const p = peek();
      if (!p) fail(null, 'paren');
      if (p.type === 'sep' || p.type === ')') args.push({ type: 'empty' });
      else args.push(expr(0));
      const q = next();
      if (!q) fail(null, 'paren');
      if (q.type === ')') break;
      if (q.type !== 'sep') fail(q);
    }
    if (args.length > 255) fail(t, 'args');
    return { type: 'fn', name: t.name, args };
  };

  const array = t => {
    const rows = [[]];
    for (;;) {
      let v = next();
      if (!v) fail(t, 'brace');
      let negative = false;
      if (v.type === 'op' && (v.op === '-' || v.op === '+')) { negative = v.op === '-'; v = next(); }
      if (!v) fail(t, 'brace');
      let node;
      if (v.type === 'num') node = { type: 'num', value: negative ? -v.value : v.value };
      else if (!negative && v.type === 'str') node = { type: 'str', value: v.value };
      else if (!negative && v.type === 'bool') node = { type: 'bool', value: v.value };
      else if (!negative && v.type === 'err') node = { type: 'err', code: v.code };
      else fail(v);
      rows[rows.length - 1].push(node);
      const s = next();
      if (!s) fail(t, 'brace');
      if (s.type === '}') break;
      if (s.type === 'arow') rows.push([]);
      else if (s.type !== 'acol' && s.type !== 'sep') fail(s);
    }
    if (rows.some(r => r.length !== rows[0].length)) fail(t, 'array');
    return { type: 'array', rows };
  };

  if (!list.length) fail(null, 'empty');
  const ast = expr(0);
  if (i < list.length) fail(list[i]);
  return { ast, tokens };
}

export function refNode(t) {
  if (t.kind === 'bad') return { type: 'err', code: '#REF!', sheet: t.sheet };
  return {
    type: 'ref', sheet: t.sheet ?? null, kind: t.kind,
    r1: Math.min(t.r1, t.r2), c1: Math.min(t.c1, t.c2), r2: Math.max(t.r1, t.r2), c2: Math.max(t.c1, t.c2),
  };
}

/** Walks a syntax tree: visit(node) for every node. */
export function walk(node, visit) {
  if (!node) return;
  visit(node);
  switch (node.type) {
    case 'fn': node.args.forEach(a => walk(a, visit)); break;
    case 'unary': case 'percent': case 'at': walk(node.arg, visit); break;
    case 'binary': walk(node.left, visit); walk(node.right, visit); break;
    default: break;
  }
}
