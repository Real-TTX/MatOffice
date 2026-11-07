// A small database for the tests: Customers 1—n Orders (integrity, cascade delete), Products, Order Details.
export function sampleDb() {
  const f = (id, name, type, extra = {}) => ({
    id, name, type, description: '', format: '', decimals: null, default: '', required: false, indexed: 'no', caption: '', validationRule: '', validationText: '',
    inputMask: '', lookup: null, ...(type === 'text' ? { size: 255 } : {}), ...(type === 'number' ? { numberSize: 'long' } : {}), ...extra,
  });
  const rows = list => list.map((v, i) => ({ id: i + 1, v }));
  return {
    version: 1,
    settings: { currency: 'EUR' },
    tables: [
      {
        id: 't1', name: 'Customers', description: '', primaryKey: ['f1'], nextId: 5, datasheet: {},
        fields: [f('f1', 'CustomerID', 'autonumber', { next: 5 }), f('f2', 'Name', 'text', { required: true }), f('f3', 'City', 'text'), f('f4', 'Since', 'datetime'), f('f5', 'Vip', 'yesno')],
        rows: rows([
          { f1: 1, f2: 'Anna Schmidt', f3: 'Berlin', f4: '2020-03-15', f5: true },
          { f1: 2, f2: 'Ben Miller', f3: 'Hamburg', f4: '2021-07-01', f5: false },
          { f1: 3, f2: 'Clara Weber', f3: 'Berlin', f4: '2022-01-10', f5: false },
          { f1: 4, f2: 'David Fox', f4: '2023-11-30', f5: true },
        ]),
      },
      {
        id: 't2', name: 'Orders', description: '', primaryKey: ['f1'], nextId: 7, datasheet: {},
        fields: [f('f1', 'OrderID', 'autonumber', { next: 7 }), f('f2', 'CustomerID', 'number'), f('f3', 'OrderDate', 'datetime'), f('f4', 'Amount', 'currency', { format: 'Currency', validationRule: '>=0', validationText: 'No negative amounts' })],
        rows: rows([
          { f1: 1, f2: 1, f3: '2026-01-05', f4: 120 },
          { f1: 2, f2: 1, f3: '2026-02-11', f4: 80.5 },
          { f1: 3, f2: 2, f3: '2026-02-20', f4: 300 },
          { f1: 4, f2: 3, f3: '2026-03-01', f4: 45 },
          { f1: 5, f2: 3, f3: '2026-03-15', f4: 60 },
          { f1: 6, f2: 3, f3: '2026-04-02', f4: 15 },
        ]),
      },
      {
        id: 't3', name: 'Order Details', description: '', primaryKey: ['f1'], nextId: 3, datasheet: {},
        fields: [f('f1', 'DetailID', 'autonumber', { next: 3 }), f('f2', 'OrderID', 'number'), f('f3', 'Qty', 'number'), f('f4', 'Price', 'currency')],
        rows: rows([{ f1: 1, f2: 1, f3: 2, f4: 50 }, { f1: 2, f2: 1, f3: 1, f4: 20 }]),
      },
    ],
    relationships: [
      { id: 'rel1', parent: { table: 't1', field: 'f1' }, child: { table: 't2', field: 'f2' }, integrity: true, cascadeUpdate: true, cascadeDelete: true, join: 1 },
      { id: 'rel2', parent: { table: 't2', field: 'f1' }, child: { table: 't3', field: 'f2' }, integrity: true, cascadeUpdate: false, cascadeDelete: false, join: 1 },
    ],
    queries: [
      { id: 'q1', name: 'Berlin Customers', sql: 'SELECT Customers.Name, Customers.City FROM Customers WHERE (Customers.City="Berlin");', layout: { tables: [] }, datasheet: {} },
    ],
    forms: [],
    reports: [],
    ui: { nav: {}, relationships: {} },
  };
}
