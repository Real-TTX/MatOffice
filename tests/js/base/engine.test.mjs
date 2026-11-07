// node --test tests/js/base — the SQL parser, the query engine, the design grid and the database model of MatOffice Base.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseSql, toSql, renameInSql, queryType } from '../../../src/MatOffice/wwwroot/js/base/sql.js';
import { runSelect, planAction, parametersOf, paramValues, runQueryObject } from '../../../src/MatOffice/wwwroot/js/base/engine.js';
import { designToStatement, statementToDesign, designToSql, sqlToDesign, emptyDesign, emptyColumn, splitAlias } from '../../../src/MatOffice/wwwroot/js/base/qbe.js';
import {
  saveRecord, deleteRecords, planDelete, previewFieldChange, applyFieldChange, relationshipProblem, coerce, parseInput, renameField, renameObject,
  defaultValues, checkRecord, compactDb, primaryKeyProblem, orientRelationship, relationshipKind, lookupOptions, tableByName,
} from '../../../src/MatOffice/wwwroot/js/base/model.js';
import { isoOf } from '../../../src/MatOffice/wwwroot/js/base/values.js';
import { sampleDb } from './fixture.mjs';

const values = res => res.rows.map(r => r.v.map(v => (v && typeof v === 'object' && 's' in v ? isoOf(v.s) : v)));

test('SQL is parsed and written back the Access way', () => {
  const sql = 'SELECT Customers.Name, Orders.OrderDate, [Amount]*2 AS Twice\nFROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID\nWHERE ((Customers.City="Berlin"))\nORDER BY Orders.OrderDate DESC;';
  const stmt = parseSql(sql);
  assert.equal(stmt.type, 'select');
  assert.equal(stmt.from[0].kind, 'INNER');
  const again = toSql(stmt);
  assert.deepEqual(parseSql(again), stmt);
  assert.match(again, /^SELECT Customers\.Name, Orders\.OrderDate, \[Amount\]\*2 AS Twice\nFROM Customers INNER JOIN Orders ON Customers\.CustomerID=Orders\.CustomerID\nWHERE \(Customers\.City="Berlin"\)\nORDER BY Orders\.OrderDate DESC;$/);
  assert.equal(queryType(parseSql('SELECT * INTO Copy FROM Customers')), 'maketable');
  assert.equal(queryType(parseSql('INSERT INTO Customers (Name) VALUES ("X")')), 'append');
  assert.equal(queryType(parseSql('UPDATE Orders SET Amount = Amount * 2')), 'update');
  assert.equal(queryType(parseSql('DELETE * FROM Orders WHERE Amount < 50')), 'delete');
  assert.throws(() => parseSql('SELECT FROM'), e => e.code === 'Syntax');
  const nested = parseSql('SELECT * FROM (Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID');
  assert.equal(nested.from[0].left.t, 'join');
  assert.match(toSql(nested), /FROM \(Customers INNER JOIN Orders ON .*\) INNER JOIN \[Order Details\] ON/);
});

test('select with where, order, top, distinct', () => {
  const db = sampleDb();
  assert.deepEqual(values(runSelect(db, 'SELECT Name FROM Customers WHERE City = "berlin" ORDER BY Name DESC')), [['Clara Weber'], ['Anna Schmidt']]);
  assert.deepEqual(values(runSelect(db, 'SELECT DISTINCT City FROM Customers WHERE City Is Not Null ORDER BY City')), [['Berlin'], ['Hamburg']]);
  assert.deepEqual(values(runSelect(db, 'SELECT TOP 2 Amount FROM Orders ORDER BY Amount DESC')), [[300], [120]]);
  assert.equal(runSelect(db, 'SELECT TOP 50 PERCENT * FROM Orders ORDER BY OrderID').rows.length, 3);
  assert.deepEqual(values(runSelect(db, 'SELECT Name FROM Customers WHERE Name Like "*er*" AND Since >= #1/1/2021#')), [['Ben Miller'], ['Clara Weber']]);
  assert.deepEqual(values(runSelect(db, 'SELECT Name, Year(Since) AS Y FROM Customers WHERE Vip ORDER BY 2 + 0, Name')).length, 2);
  const res = runSelect(db, 'SELECT * FROM Customers');
  assert.equal(res.columns.length, 5);
  assert.ok(res.columns[1].editable);
  assert.ok(!res.columns[0].editable, 'AutoNumber is not editable');
});

test('joins: inner, left, right', () => {
  const db = sampleDb();
  const inner = runSelect(db, 'SELECT Customers.Name, Orders.Amount FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID');
  assert.equal(inner.rows.length, 6);
  const left = runSelect(db, 'SELECT Customers.Name, Orders.Amount FROM Customers LEFT JOIN Orders ON Customers.CustomerID = Orders.CustomerID WHERE Orders.OrderID Is Null');
  assert.deepEqual(values(left), [['David Fox', null]]);
  const right = runSelect(db, 'SELECT Customers.Name, Orders.OrderID FROM Orders RIGHT JOIN Customers ON Orders.CustomerID = Customers.CustomerID');
  assert.equal(right.rows.length, 7);
  const three = runSelect(db, 'SELECT Customers.Name, [Order Details].Qty FROM (Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID');
  assert.deepEqual(values(three), [['Anna Schmidt', 2], ['Anna Schmidt', 1]]);
  const comma = runSelect(db, 'SELECT c.Name FROM Customers AS c, Orders AS o WHERE c.CustomerID = o.CustomerID AND o.Amount > 100');
  assert.deepEqual(values(comma), [['Anna Schmidt'], ['Ben Miller']]);
});

test('group by with aggregates and having', () => {
  const db = sampleDb();
  const res = runSelect(db, 'SELECT Customers.City, Count(*) AS N, Sum(Orders.Amount) AS Total, Avg(Orders.Amount) FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID GROUP BY Customers.City HAVING Sum(Orders.Amount) > 200 ORDER BY Sum(Orders.Amount) DESC');
  assert.deepEqual(res.columns.map(c => c.name), ['City', 'N', 'Total', 'AvgOfAmount']);
  assert.deepEqual(values(res), [['Berlin', 5, 320.5, 64.1], ['Hamburg', 1, 300, 300]]);
  assert.deepEqual(values(runSelect(db, 'SELECT Count(*), Max(Amount), Min(OrderDate) FROM Orders')), [[6, 300, '2026-01-05']]);
  assert.deepEqual(values(runSelect(db, 'SELECT Count(*) FROM Orders WHERE Amount > 1000')), [[0]]);
  assert.throws(() => runSelect(db, 'SELECT City, Name, Count(*) FROM Customers GROUP BY City'), e => e.code === 'NotInGroup');
});

test('saved queries, subqueries and domain functions', () => {
  const db = sampleDb();
  assert.equal(runSelect(db, 'SELECT * FROM [Berlin Customers]').rows.length, 2);
  assert.deepEqual(values(runSelect(db, 'SELECT X.Name FROM (SELECT Name, City FROM Customers WHERE Vip) AS X ORDER BY X.Name')), [['Anna Schmidt'], ['David Fox']]);
  assert.deepEqual(values(runSelect(db, 'SELECT Name, DCount("*", "Orders", "CustomerID=" & [CustomerID]) AS N FROM Customers WHERE CustomerID = 3')), [['Clara Weber', 3]]);
  assert.deepEqual(values(runSelect(db, 'SELECT DLookup("Name", "Customers", "CustomerID=2") FROM Orders WHERE OrderID = 1')), [['Ben Miller']]);
  db.queries.push({ id: 'q2', name: 'Loop', sql: 'SELECT * FROM Loop', layout: { tables: [] } });
  assert.throws(() => runSelect(db, 'SELECT * FROM Loop'), e => e.code === 'Circular');
});

test('parameters are found and asked for', () => {
  const db = sampleDb();
  const stmt = parseSql('SELECT Name FROM Customers WHERE City = [Enter a city] AND Since > [From date]');
  const params = parametersOf(db, stmt);
  assert.deepEqual(params.map(p => p.name), ['Enter a city', 'From date']);
  const res = runSelect(db, stmt, { params: paramValues(params, { 'enter a city': 'Berlin', 'from date': '1/1/2021' }) });
  assert.deepEqual(values(res), [['Clara Weber']]);
  const declared = parseSql('PARAMETERS [Min] Long; SELECT Amount FROM Orders WHERE Amount >= [Min]');
  const p2 = parametersOf(db, declared);
  assert.deepEqual(p2, [{ name: 'Min', type: 'Long' }]);
  assert.equal(runSelect(db, declared, { params: paramValues(p2, { min: '100' }) }).rows.length, 2);
});

test('action queries go through the checks', () => {
  const db = sampleDb();
  const up = planAction(db, parseSql('UPDATE Orders SET Amount = Amount * 2 WHERE CustomerID = 3'));
  assert.equal(up.count, 3);
  const report = up.apply();
  assert.equal(report.done, 3);
  assert.deepEqual(values(runSelect(db, 'SELECT Amount FROM Orders WHERE CustomerID = 3')), [[90], [120], [30]]);
  const bad = planAction(db, parseSql('UPDATE Orders SET Amount = -1 WHERE OrderID = 1')).apply();
  assert.equal(bad.ruleViolations, 1);
  const append = planAction(db, parseSql('INSERT INTO Customers ( Name, City ) SELECT Name & " (copy)", City FROM Customers WHERE City = "Hamburg"'));
  assert.equal(append.count, 1);
  assert.equal(append.apply().done, 1);
  assert.equal(tableByName(db, 'Customers').rows.at(-1).v.f1, 5, 'AutoNumber continues');
  const orphan = planAction(db, parseSql('INSERT INTO Orders (CustomerID, Amount) VALUES (99, 5)')).apply();
  assert.equal(orphan.keyViolations, 1);
  const make = planAction(db, parseSql('SELECT Name, City INTO BerlinCopy FROM Customers WHERE City = "Berlin"'));
  assert.equal(make.apply().done, 2);
  assert.equal(tableByName(db, 'BerlinCopy').fields[0].type, 'text');
  const del = planAction(db, parseSql('DELETE * FROM Customers WHERE CustomerID = 2'));
  assert.equal(del.count, 1);
  const delReport = del.apply();
  assert.equal(delReport.done, 1);
  assert.equal(delReport.cascaded, 1, 'the order of Ben goes with him (cascade delete)');
  const refused = planAction(db, parseSql('DELETE * FROM Orders WHERE OrderID = 1')).apply();
  assert.equal(refused.error.code, 'HasChildren');
});

test('records: required, rules, keys, integrity, cascades', () => {
  const db = sampleDb();
  const customers = db.tables[0];
  const orders = db.tables[1];
  assert.equal(saveRecord(db, customers, null, { f3: 'Köln' }).error.code, 'Required');
  const ok = saveRecord(db, customers, null, { f2: 'Eva Stein', f3: 'Köln' });
  assert.ok(ok.ok);
  assert.equal(ok.row.v.f1, 5);
  assert.equal(customers.fields[0].next, 6);
  assert.equal(saveRecord(db, orders, null, { f2: 42, f4: 1 }).error.code, 'NoParent');
  assert.equal(saveRecord(db, orders, null, { f2: 1, f4: -5 }).error.code, 'Validation');
  assert.equal(saveRecord(db, orders, 1, { f1: 2 }).error.code, 'Duplicate');
  // cascade update: a new key of Anna moves her orders along
  assert.ok(saveRecord(db, customers, 1, { f1: 100 }).ok);
  assert.equal(orders.rows.filter(r => r.v.f2 === 100).length, 2);
  // no cascade update between Orders and Order Details
  assert.equal(saveRecord(db, orders, 1, { f1: 77 }).error.code, 'HasChildren');
  // cascade delete through one level, refused through the next
  assert.equal(planDelete(db, customers, [1]).ok, false);
  db.relationships[1].cascadeDelete = true;
  const plan = deleteRecords(db, customers, [1]);
  assert.ok(plan.ok);
  assert.equal(plan.cascaded, 4);
  assert.equal(db.tables[2].rows.length, 0);
});

test('values typed into fields', () => {
  const db = sampleDb();
  const [id, name, , since, vip] = db.tables[0].fields;
  const amount = db.tables[1].fields[3];
  assert.deepEqual(parseInput(since, '7.11.2026', 'de'), { ok: true, value: '2026-11-07' });
  assert.equal(parseInput(since, 'gestern', 'de').ok, false);
  assert.deepEqual(parseInput(amount, '1.234,5', 'de'), { ok: true, value: 1234.5 });
  assert.deepEqual(parseInput(vip, 'ja', 'de'), { ok: true, value: true });
  assert.deepEqual(parseInput(name, '', 'en'), { ok: true, value: null });
  assert.equal(coerce({ ...name, size: 3 }, 'abcdef').code, 'TooLong');
  assert.equal(coerce({ type: 'number', numberSize: 'byte' }, 300).code, 'OutOfRange');
  assert.equal(coerce({ type: 'number', numberSize: 'integer' }, 2.5).value, 2);
  assert.equal(id.type, 'autonumber');
  const defaults = defaultValues(db, { fields: [{ id: 'a', type: 'datetime', default: 'Date()' }, { id: 'b', type: 'text', default: '"Berlin"' }, { id: 'c', type: 'yesno', default: 'Yes' }] }, 'en', new Date(2026, 10, 7));
  assert.deepEqual(defaults, { a: '2026-11-07', b: 'Berlin', c: true });
});

test('design changes keep the data where they can', () => {
  const db = sampleDb();
  const customers = db.tables[0];
  const city = customers.fields[2];
  assert.deepEqual(previewFieldChange(db, customers, city.id, { type: 'number' }), { lost: 3, truncated: 0 });
  assert.deepEqual(previewFieldChange(db, customers, city.id, { size: 4 }), { lost: 0, truncated: 3 });
  assert.equal(previewFieldChange(db, customers, city.id, { required: true }).blocked.code, 'RequiredViolated');
  assert.equal(previewFieldChange(db, customers, city.id, { indexed: 'unique' }).blocked.code, 'DuplicatesExist');
  assert.equal(previewFieldChange(db, customers, customers.fields[0].id, { type: 'text' }).blocked.code, 'FieldInRelationship');
  applyFieldChange(db, customers, city.id, { size: 4 });
  assert.equal(customers.rows[0].v.f3, 'Berl');
  const since = customers.fields[3];
  applyFieldChange(db, customers, since.id, { type: 'text' });
  assert.equal(customers.rows[0].v.f4, '3/15/2020');
  applyFieldChange(db, customers, since.id, { type: 'datetime' });
  assert.equal(customers.rows[0].v.f4, '2020-03-15');
  assert.equal(primaryKeyProblem(db, customers, [city.id]).code, 'PkNulls');
});

test('relationships: types, unique side, existing data', () => {
  const db = sampleDb();
  const rel = { id: 'x', parent: { table: 't1', field: 'f3' }, child: { table: 't2', field: 'f2' }, integrity: false };
  assert.equal(relationshipProblem(db, rel).code, 'RelTypes');
  const dup = { id: 'y', parent: { table: 't1', field: 'f1' }, child: { table: 't2', field: 'f2' }, integrity: true };
  assert.equal(relationshipProblem(db, dup).code, 'RelExists');
  db.relationships = [];
  db.tables[1].rows.push({ id: 99, v: { f1: 99, f2: 55 } });
  assert.equal(relationshipProblem(db, dup).code, 'RelViolated');
  const o = orientRelationship(db, { table: 't2', field: 'f2' }, { table: 't1', field: 'f1' });
  assert.equal(o.parent.table, 't1');
  assert.equal(relationshipKind(db, { parent: { table: 't1', field: 'f1' }, child: { table: 't2', field: 'f2' } }), 'oneToMany');
});

test('rename keeps queries and lookups working (Name AutoCorrect)', () => {
  const db = sampleDb();
  renameField(db, db.tables[0], 'f3', 'Town');
  assert.equal(db.queries[0].sql, 'SELECT Customers.Name, Customers.Town FROM Customers WHERE (Customers.Town="Berlin");');
  renameObject(db, 'tables', 't1', 'Clients');
  assert.equal(db.queries[0].sql, 'SELECT Clients.Name, Clients.Town FROM Clients WHERE (Clients.Town="Berlin");');
  assert.equal(runSelect(db, db.queries[0].sql).rows.length, 2);
  assert.equal(renameInSql('SELECT [Name] FROM [Order Details]', 'object', 'Order Details', 'Lines'), 'SELECT [Name] FROM [Lines]');
});

test('lookups and compact', () => {
  const db = sampleDb();
  db.tables[1].fields[1].lookup = { kind: 'table', table: 't1', bound: 'f1', display: 'f2' };
  const options = lookupOptions(db, db.tables[1].fields[1], 'en');
  assert.deepEqual(options.map(o => o.label), ['Anna Schmidt', 'Ben Miller', 'Clara Weber', 'David Fox']);
  db.tables[0].fields[0].next = 999;
  db.tables[0].rows[0].v.zz = 1;
  assert.ok(compactDb(db) >= 2);
  assert.equal(db.tables[0].fields[0].next, 5);
});

// ---- the design grid -------------------------------------------------------------------------------------------------------------

const fieldsOf = db => name => {
  const t = tableByName(db, name);
  if (t) return t.fields.map(f => f.name);
  const q = db.queries.find(x => x.name.toLowerCase() === name.toLowerCase());
  return q ? runQueryObject(db, q).columns.map(c => c.name) : [];
};

function design(over) {
  return { ...emptyDesign(), ...over, columns: over.columns.map(c => ({ ...emptyColumn(), ...c })) };
}

test('the grid becomes SQL and the SQL becomes the same grid', () => {
  const db = sampleDb();
  const d = design({
    tables: [{ name: 'Customers', alias: null }, { name: 'Orders', alias: null }],
    joins: [{ left: 'Customers', leftField: 'CustomerID', right: 'Orders', rightField: 'CustomerID', kind: 'INNER' }],
    columns: [
      { field: 'Name', table: 'Customers', sort: 'asc' },
      { field: 'OrderDate', table: 'Orders' },
      { field: 'Twice: [Amount]*2', table: '' },
      { field: 'City', table: 'Customers', show: false, criteria: ['"Berlin"', 'Hamburg'] },
      { field: 'Amount', table: 'Orders', show: false, criteria: ['>50'] },
    ],
  });
  const sql = designToSql(d, db, fieldsOf(db));
  assert.equal(sql, 'SELECT Customers.Name, Orders.OrderDate, [Amount]*2 AS Twice\nFROM Customers INNER JOIN Orders ON Customers.CustomerID=Orders.CustomerID\nWHERE ((Customers.City="Berlin") AND (Orders.Amount>50)) OR ((Customers.City="Hamburg"))\nORDER BY Customers.Name;');
  const back = sqlToDesign(sql, db, fieldsOf(db));
  assert.deepEqual(back.columns.map(c => [c.field, c.table, c.sort, c.show, c.criteria.join('|')]), [
    ['Name', 'Customers', 'asc', true, '|'],
    ['OrderDate', 'Orders', '', true, '|'],
    ['Twice: [Amount] * 2', '', '', true, '|'],
    ['City', 'Customers', '', false, '"Berlin"|"Hamburg"'],
    ['Amount', 'Orders', '', false, '>50|'],
  ]);
  assert.equal(designToSql(back, db, fieldsOf(db)).replace('[Amount] * 2', '[Amount]*2'), sql);
  const res = runSelect(db, sql);
  assert.deepEqual(values(res).map(r => r[0]), ['Anna Schmidt', 'Anna Schmidt', 'Ben Miller', 'Clara Weber']);
});

test('totals in the grid', () => {
  const db = sampleDb();
  const d = design({
    totals: true,
    tables: [{ name: 'Customers', alias: null }, { name: 'Orders', alias: null }],
    joins: [{ left: 'Customers', leftField: 'CustomerID', right: 'Orders', rightField: 'CustomerID', kind: 'LEFT' }],
    columns: [
      { field: 'City', table: 'Customers', total: 'Group By' },
      { field: 'Amount', table: 'Orders', total: 'Sum', sort: 'desc', criteria: ['>100'] },
      { field: 'OrderDate', table: 'Orders', total: 'Where', show: false, criteria: ['>=#1/1/2026#'] },
    ],
  });
  const sql = designToSql(d, db, fieldsOf(db));
  assert.equal(sql, 'SELECT Customers.City, Sum(Orders.Amount) AS SumOfAmount\nFROM Customers LEFT JOIN Orders ON Customers.CustomerID=Orders.CustomerID\nWHERE (Orders.OrderDate>=#1/1/2026#)\nGROUP BY Customers.City\nHAVING (Sum(Orders.Amount)>100)\nORDER BY Sum(Orders.Amount) DESC;');
  const back = sqlToDesign(sql, db, fieldsOf(db));
  assert.deepEqual(back.columns.map(c => [c.field, c.total, c.sort, c.show, c.criteria.join('|')]), [
    ['City', 'Group By', '', true, ''],
    ['Amount', 'Sum', 'desc', true, '>100'],
    ['OrderDate', 'Where', '', false, '>=#1/1/2026#'],
  ]);
  assert.deepEqual(values(runSelect(db, sql)), [['Berlin', 320.5], ['Hamburg', 300]]);
});

test('action queries in the grid', () => {
  const db = sampleDb();
  const update = design({ type: 'update', tables: [{ name: 'Orders' }], columns: [{ field: 'Amount', table: 'Orders', updateTo: '[Amount]*1.1', criteria: ['<50'] }] });
  const sql = designToSql(update, db, fieldsOf(db));
  assert.equal(sql, 'UPDATE Orders\nSET Orders.Amount = [Amount]*1.1\nWHERE (Orders.Amount<50);');
  const back = sqlToDesign(sql, db, fieldsOf(db));
  assert.equal(back.type, 'update');
  assert.equal(back.columns[0].updateTo, '[Amount] * 1.1');
  const append = design({ type: 'append', target: 'Customers', tables: [{ name: 'Customers' }], columns: [{ field: 'Name', table: 'Customers', appendTo: 'Name' }, { field: 'City', table: 'Customers', criteria: ['"Hamburg"'] }] });
  assert.equal(designToSql(append, db, fieldsOf(db)), 'INSERT INTO Customers ( Name )\nSELECT Customers.Name\nFROM Customers\nWHERE (Customers.City="Hamburg");');
  const del = design({ type: 'delete', tables: [{ name: 'Orders' }], columns: [{ field: '*', table: 'Orders', del: 'From' }, { field: 'Amount', table: 'Orders', del: 'Where', criteria: ['<20'] }] });
  assert.equal(designToSql(del, db, fieldsOf(db)), 'DELETE Orders.*\nFROM Orders\nWHERE (Orders.Amount<20);');
  const make = design({ type: 'maketable', target: 'Archive', tables: [{ name: 'Orders' }], columns: [{ field: '*', table: 'Orders' }] });
  assert.equal(designToSql(make, db, fieldsOf(db)), 'SELECT Orders.*\nINTO Archive\nFROM Orders;');
  assert.equal(sqlToDesign('INSERT INTO Customers (Name) VALUES ("x")', db, fieldsOf(db)), null);
  assert.deepEqual(splitAlias('Total: [Price]*[Qty]'), { alias: 'Total', expr: '[Price]*[Qty]' });
});

test('SQL written by hand shows in the grid', () => {
  const db = sampleDb();
  const d = sqlToDesign('SELECT Name, City FROM Customers WHERE City = "Berlin" OR Vip = True ORDER BY Name', db, fieldsOf(db));
  assert.deepEqual(d.columns.map(c => [c.field, c.table, c.show, c.criteria.join('|')]), [
    ['Name', 'Customers', true, '|'],
    ['City', 'Customers', true, '"Berlin"|'],
    ['Vip', 'Customers', false, '|True'],
  ]);
  assert.equal(d.columns[0].sort, 'asc');
  const p = sqlToDesign('SELECT Name FROM Customers WHERE City = [Which city?]', db, fieldsOf(db));
  assert.equal(p.columns[1].criteria[0], '[Which city?]');
});
