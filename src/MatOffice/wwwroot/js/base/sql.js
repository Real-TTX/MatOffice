/* MatOffice Base – the SQL of Access queries: the parser (SELECT … INTO … FROM … [INNER|LEFT|RIGHT JOIN … ON …] WHERE … GROUP BY … HAVING … ORDER BY …,
   TOP n [PERCENT], DISTINCT; INSERT INTO … SELECT|VALUES; UPDATE … SET … WHERE; DELETE [t.*] FROM … WHERE; PARAMETERS …;) and the writer that
   produces Access-style SQL (Table.Field, nested joins in parentheses, the closing semicolon). A pure module without the DOM.

   Statement nodes:
     { type: 'select', distinct: null|'DISTINCT'|'DISTINCTROW', top: { n, percent }|null, columns: [{ e, alias }], into: name|null,
       from: [FromItem], where, groupBy: [e], having, orderBy: [{ e, desc }], params: [{ name, type }] }
     { type: 'insert', table, columns: [name]|null, select|null, values: [e]|null, params }
     { type: 'update', from: [FromItem], set: [{ target: idNode, e }], where, params }
     { type: 'delete', target: name|null, from: [FromItem], where, params }
   FromItem: { t: 'table', name, alias } | { t: 'join', kind: 'INNER'|'LEFT'|'RIGHT', left, right, on } | { t: 'sub', query, alias } */
import { Parser, BaseError, exprToText, quoteName } from './expr.js';

class SqlParser extends Parser {
  statement() {
    const params = this.parameters();
    let stmt;
    if (this.isKw('SELECT')) stmt = this.select();
    else if (this.isKw('INSERT')) stmt = this.insert();
    else if (this.isKw('UPDATE')) stmt = this.update();
    else if (this.isKw('DELETE')) stmt = this.delete_();
    else if (this.isKw('TRANSFORM')) this.fail('Unsupported');
    else this.fail();
    stmt.params = params;
    this.acceptOp(';');
    if (!this.atEnd()) {
      if (this.isKw('UNION')) this.fail('Unsupported');
      this.fail();
    }
    return stmt;
  }

  parameters() {
    const list = [];
    if (!this.acceptKw('PARAMETERS')) return list;
    do {
      const t = this.next();
      if (t.k !== 'name' && t.k !== 'word') this.fail();
      let type = '';
      if (this.tok.k === 'word') {
        type = this.next().v;
        if (this.acceptOp('(')) { this.next(); this.expectOp(')'); }
      }
      list.push({ name: t.v, type });
    } while (this.acceptOp(','));
    this.expectOp(';');
    return list;
  }

  select() {
    this.expectKw('SELECT');
    let distinct = null;
    if (this.acceptKw('DISTINCTROW')) distinct = 'DISTINCTROW';
    else if (this.acceptKw('DISTINCT')) distinct = 'DISTINCT';
    else this.acceptKw('ALL');
    let top = null;
    if (this.acceptKw('TOP')) {
      const t = this.next();
      if (t.k !== 'num' || t.v <= 0) this.fail();
      top = { n: t.v, percent: this.acceptKw('PERCENT') };
    }
    const columns = [];
    do { columns.push(this.column()); } while (this.acceptOp(','));
    let into = null;
    if (this.acceptKw('INTO')) into = this.objectName();
    this.expectKw('FROM');
    const from = this.fromList();
    const where = this.acceptKw('WHERE') ? this.expr() : null;
    let groupBy = [];
    if (this.acceptKw('GROUP', 'BY')) { do { groupBy.push(this.expr()); } while (this.acceptOp(',')); }
    const having = this.acceptKw('HAVING') ? this.expr() : null;
    const orderBy = [];
    if (this.acceptKw('ORDER', 'BY')) {
      do {
        const e = this.expr();
        let desc = false;
        if (this.acceptKw('DESC')) desc = true; else this.acceptKw('ASC');
        orderBy.push({ e, desc });
      } while (this.acceptOp(','));
    }
    return { type: 'select', distinct, top, columns, into, from, where, groupBy, having, orderBy, params: [] };
  }

  column() {
    if (this.isOp('*')) { this.next(); return { e: { t: 'star', table: null }, alias: null }; }
    const e = this.expr();
    let alias = null;
    if (this.acceptKw('AS')) alias = this.aliasName();
    return { e, alias };
  }

  aliasName() {
    const t = this.next();
    if (t.k === 'name' || t.k === 'str') return t.v;
    if (t.k === 'word' && !['FROM', 'WHERE', 'INTO'].includes(t.v.toUpperCase())) return t.v;
    this.i--;
    this.fail();
    return null;
  }

  objectName() {
    const t = this.next();
    if (t.k === 'name' || t.k === 'word') {
      if (t.k === 'word' && ['SELECT', 'FROM', 'WHERE', 'SET', 'VALUES'].includes(t.v.toUpperCase())) { this.i--; this.fail(); }
      return t.v;
    }
    this.i--;
    this.fail();
    return null;
  }

  fromList() {
    const items = [this.fromItem()];
    while (this.acceptOp(',')) items.push(this.fromItem());
    return items;
  }

  fromItem() {
    let left = this.fromPrimary();
    for (;;) {
      let kind = null;
      if (this.acceptKw('INNER', 'JOIN')) kind = 'INNER';
      else if (this.acceptKw('LEFT', 'OUTER', 'JOIN') || this.acceptKw('LEFT', 'JOIN')) kind = 'LEFT';
      else if (this.acceptKw('RIGHT', 'OUTER', 'JOIN') || this.acceptKw('RIGHT', 'JOIN')) kind = 'RIGHT';
      else if (this.isKw('JOIN') || this.isKw('FULL') || this.isKw('CROSS')) this.fail('Unsupported');
      if (!kind) return left;
      const right = this.fromPrimary();
      this.expectKw('ON');
      const on = this.expr();
      left = { t: 'join', kind, left, right, on };
    }
  }

  fromPrimary() {
    if (this.isOp('(')) {
      this.next();
      if (this.isKw('SELECT')) {
        const query = this.select();
        this.expectOp(')');
        this.acceptKw('AS');
        const alias = this.objectName();
        return { t: 'sub', query, alias };
      }
      const inner = this.fromItem();
      this.expectOp(')');
      return inner;
    }
    const name = this.objectName();
    let alias = null;
    if (this.acceptKw('AS')) alias = this.objectName();
    else if ((this.tok.k === 'name' || this.tok.k === 'word') && !this.isReservedHere()) alias = this.objectName();
    return { t: 'table', name, alias };
  }

  isReservedHere() {
    const t = this.tok;
    return t.k === 'word' && ['WHERE', 'INNER', 'LEFT', 'RIGHT', 'JOIN', 'ON', 'GROUP', 'ORDER', 'HAVING', 'SET', 'UNION', 'FULL', 'CROSS', 'OUTER'].includes(t.v.toUpperCase());
  }

  insert() {
    this.expectKw('INSERT');
    this.expectKw('INTO');
    const table = this.objectName();
    let columns = null;
    if (this.isOp('(') && !(this.peek().k === 'word' && this.peek().v.toUpperCase() === 'SELECT')) {
      this.next();
      columns = [];
      do {
        const t = this.next();
        if (t.k !== 'name' && t.k !== 'word') this.fail();
        let name = t.v;
        if (this.acceptOp('.')) name = this.next().v;
        columns.push(name);
      } while (this.acceptOp(','));
      this.expectOp(')');
    }
    if (this.acceptKw('VALUES')) {
      this.expectOp('(');
      const values = [];
      do { values.push(this.expr()); } while (this.acceptOp(','));
      this.expectOp(')');
      return { type: 'insert', table, columns, select: null, values, params: [] };
    }
    const paren = this.acceptOp('(');
    const select = this.select();
    if (paren) this.expectOp(')');
    return { type: 'insert', table, columns, select, values: null, params: [] };
  }

  update() {
    this.expectKw('UPDATE');
    const from = this.fromList();
    this.expectKw('SET');
    const set = [];
    do {
      const target = this.ident();
      this.expectOp('=');
      set.push({ target, e: this.expr() });
    } while (this.acceptOp(','));
    const where = this.acceptKw('WHERE') ? this.expr() : null;
    return { type: 'update', from, set, where, params: [] };
  }

  delete_() {
    this.expectKw('DELETE');
    let target = null;
    if (!this.isKw('FROM')) {
      if (this.isOp('*')) this.next();
      else {
        const t = this.next();
        if (t.k !== 'name' && t.k !== 'word') this.fail();
        target = t.v;
        this.expectOp('.');
        this.expectOp('*');
      }
    }
    this.expectKw('FROM');
    const from = this.fromList();
    const where = this.acceptKw('WHERE') ? this.expr() : null;
    return { type: 'delete', target, from, where, params: [] };
  }
}

/** Parses one SQL statement; throws BaseError('Syntax', near, position) or BaseError('Unsupported', …). */
export function parseSql(text) {
  const p = new SqlParser(text);
  if (p.atEnd()) throw new BaseError('Empty');
  return p.statement();
}

/** The kind of a statement: select, maketable, append, update, delete. */
export function queryType(stmt) {
  if (stmt.type === 'select') return stmt.into ? 'maketable' : 'select';
  if (stmt.type === 'insert') return 'append';
  return stmt.type;
}

// ---- writing -------------------------------------------------------------------------------------------------------------------

/** A name of a field in SQL: Table.Field (brackets where needed). */
export const sqlName = parts => parts.map(p => quoteName(p)).join('.');

export const sqlExpr = e => exprToText(e, { name: parts => (parts.length === 1 ? '[' + parts[0] + ']' : sqlName(parts)), compact: true });

function fromText(item, nested = false) {
  switch (item.t) {
    case 'table': return quoteName(item.name) + (item.alias ? ' AS ' + quoteName(item.alias) : '');
    case 'sub': return '(' + selectText(item.query, true).replace(/;$/, '') + ') AS ' + quoteName(item.alias);
    case 'join': {
      const left = fromText(item.left, true);
      const s = `${left} ${item.kind} JOIN ${fromText(item.right, true)} ON ${sqlExpr(item.on)}`;
      return nested ? '(' + s + ')' : s;
    }
    default: return '';
  }
}

function paramsText(params) {
  if (!params?.length) return '';
  return 'PARAMETERS ' + params.map(p => '[' + p.name + ']' + (p.type ? ' ' + p.type : '')).join(', ') + ';\n';
}

function selectText(s, inline = false) {
  const lines = [];
  let head = 'SELECT ';
  if (s.distinct) head += s.distinct + ' ';
  if (s.top) head += 'TOP ' + s.top.n + (s.top.percent ? ' PERCENT' : '') + ' ';
  head += s.columns.map(c => sqlExpr(c.e) + (c.alias ? ' AS ' + quoteName(c.alias) : '')).join(', ');
  lines.push(head);
  if (s.into) lines.push('INTO ' + quoteName(s.into));
  lines.push('FROM ' + s.from.map(f => fromText(f)).join(', '));
  if (s.where) lines.push('WHERE ' + wrapCriteria(s.where));
  if (s.groupBy?.length) lines.push('GROUP BY ' + s.groupBy.map(sqlExpr).join(', '));
  if (s.having) lines.push('HAVING ' + wrapCriteria(s.having));
  if (s.orderBy?.length) lines.push('ORDER BY ' + s.orderBy.map(o => sqlExpr(o.e) + (o.desc ? ' DESC' : '')).join(', '));
  return lines.join(inline ? ' ' : '\n') + ';';
}

/** Access puts the criteria in parentheses: WHERE ((Customers.City="Berlin") AND (Orders.Total>10)). */
function wrapCriteria(e) {
  const branches = splitBy(e, 'OR');
  const branch = (b, alone) => {
    const terms = splitBy(b, 'AND');
    if (alone && terms.length === 1) return '(' + sqlExpr(terms[0]) + ')';
    return '(' + terms.map(t => '(' + sqlExpr(t) + ')').join(' AND ') + ')';
  };
  if (branches.length === 1) return branch(e, true);
  return branches.map(b => branch(b, false)).join(' OR ');
}

function splitBy(e, op) {
  const out = [];
  const go = x => { if (x.t === 'bin' && x.op === op) { go(x.l); go(x.r); } else out.push(x); };
  go(e);
  return out;
}

/** The SQL text of a statement, the way Access writes it. */
export function toSql(stmt) {
  const pre = paramsText(stmt.params);
  switch (stmt.type) {
    case 'select': return pre + selectText(stmt);
    case 'insert': {
      const cols = stmt.columns ? ' ( ' + stmt.columns.map(c => quoteName(c)).join(', ') + ' )' : '';
      if (stmt.values) return pre + `INSERT INTO ${quoteName(stmt.table)}${cols}\nVALUES (${stmt.values.map(sqlExpr).join(', ')});`;
      return pre + `INSERT INTO ${quoteName(stmt.table)}${cols}\n` + selectText(stmt.select);
    }
    case 'update': {
      const lines = ['UPDATE ' + stmt.from.map(f => fromText(f)).join(', '), 'SET ' + stmt.set.map(s => sqlExpr(s.target) + ' = ' + sqlExpr(s.e)).join(', ')];
      if (stmt.where) lines.push('WHERE ' + wrapCriteria(stmt.where));
      return pre + lines.join('\n') + ';';
    }
    case 'delete': {
      const lines = ['DELETE ' + (stmt.target ? quoteName(stmt.target) + '.*' : '*'), 'FROM ' + stmt.from.map(f => fromText(f)).join(', ')];
      if (stmt.where) lines.push('WHERE ' + wrapCriteria(stmt.where));
      return pre + lines.join('\n') + ';';
    }
    default: return '';
  }
}

/** The tables and queries a statement reads (names as written). */
export function sourcesOf(stmt) {
  const out = [];
  const visit = item => {
    if (!item) return;
    if (item.t === 'table') out.push(item.name);
    else if (item.t === 'join') { visit(item.left); visit(item.right); }
    else if (item.t === 'sub') item.query.from.forEach(visit);
  };
  if (stmt.type === 'insert') { if (stmt.select) stmt.select.from.forEach(visit); }
  else stmt.from.forEach(visit);
  return out;
}

/**
 * Renames a table/query or a field in SQL text, token by token, so the rest of the text stays as the person wrote it (Name AutoCorrect of Access).
 * kind 'object': every name equal to `from` (alone or before a dot) becomes `to`. kind 'field': `from` after `table.` and, when `table` is a source
 * of the statement, alone.
 */
export function renameInSql(sql, kind, from, to, table = null) {
  let tokens;
  try { tokens = new Parser(sql).tokens; } catch { return sql; }
  const lower = from.toLowerCase();
  let usesTable = false;
  if (kind === 'field' && table) {
    try { usesTable = sourcesOf(parseSql(sql)).some(n => n.toLowerCase() === table.toLowerCase()); } catch { usesTable = false; }
  }
  const edits = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if ((t.k !== 'name' && t.k !== 'word') || t.v.toLowerCase() !== lower) continue;
    const prev = tokens[i - 1];
    const next = tokens[i + 1];
    const afterDot = prev && prev.k === 'op' && (prev.v === '.' || prev.v === '!');
    const beforeParen = next && next.k === 'op' && next.v === '(';
    if (t.k === 'word' && beforeParen) continue;
    if (kind === 'object') {
      if (afterDot) continue;
    } else {
      if (afterDot) {
        const owner = tokens[i - 2];
        if (!owner || owner.v.toLowerCase() !== (table || '').toLowerCase()) continue;
      } else {
        const beforeDot = next && next.k === 'op' && (next.v === '.' || next.v === '!');
        if (beforeDot || !usesTable) continue;
      }
    }
    edits.push([t.pos, t.end, t.k === 'name' ? '[' + to + ']' : quoteName(to)]);
  }
  let out = sql;
  for (const [a, b, text] of edits.reverse()) out = out.slice(0, a) + text + out.slice(b);
  return out;
}
