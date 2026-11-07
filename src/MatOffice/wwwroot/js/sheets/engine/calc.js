/* MatOffice Sheets – the calculation of a workbook (pure, no DOM).
   Every formula cell knows the areas it reads (from its syntax tree); an index finds the formulas that read a cell. After a change only the formulas
   that depend on it are calculated again, in the order of their dependencies (Kahn); a dependency that is only known while calculating (INDIRECT,
   OFFSET, names) is calculated on demand. A circular reference gives #CIRC!. An array result spills into the empty cells below and to the right
   (#SPILL! when something is in the way).

   The host (the workbook model) gives: cell(sheet, row, col) → { v, f, t } | undefined, sheetId(name), sheetName(id), sheetIds(), sheetIndex(id),
   bounds(sheet) → { rows, cols } (the used area), formulas(visit(sheet, row, col, text)), names() → [{ name, ref }], rowHidden(sheet, row),
   setResult(sheet, row, col, value) (the last value of a formula, written into the content). */
import { parse, walk } from './parser.js';
import { evaluate, result } from './evaluator.js';
import { Errors, FormulaError, Ref, errorOf } from './values.js';
import { FUNCTIONS } from './lib.js';
import { key, MAX_ROWS, MAX_COLS } from './address.js';

const COMPUTING = 1;
const MAX_BUCKETS = 256;
const MAX_DEPTH = 400;
const CLIP_CELLS = 20000;
const bucketKey = (rb, cb) => rb * 2048 + cb;

class FormulaCell {
  constructor(sheet, row, col) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.text = '';
    this.ast = null;
    this.entries = [];
    this.volatile = false;
    this.value = null;
    this.array = null;   // the spilled array
    this.spill = null;   // { r1, c1, r2, c2 } while spilled
    this.blocked = null; // the area it wants while #SPILL!
    this.dirty = false;
    this.state = 0;
    this.indeg = 0;
  }
}

export class Calc {
  constructor(host) {
    this.host = host;
    this.cells = new Map();   // sheet → Map<key, FormulaCell>
    this.index = new Map();   // sheet → { cell: Map<key, Set<FormulaCell>>, buckets: Map<bucket, Set<entry>>, big: Set<entry> }
    this.spills = new Map();  // sheet → Map<key, FormulaCell> (the cells a spill covers, not its anchor)
    this.anchors = new Set();
    this.volatiles = new Set();
    this.names = new Map();
    this.asts = new Map();    // text → syntax tree (rules, validations)
    this.late = [];
    this.depth = 0;
    this.now = () => new Date();
    this.random = Math.random;
    const calc = this;
    this.base = {
      value(sheet, row, col) { return calc.valueAt(sheet, row, col); },
      values(ref) { return calc.values(ref); },
      sheetId(name) { return calc.host.sheetId(name); },
      sheetName(id) { return calc.host.sheetName(id); },
      sheetIndex(id) { return calc.host.sheetIndex(id); },
      sheetCount() { return calc.host.sheetIds().length; },
      now() { return calc.now(); },
      random() { return calc.random(); },
      name(name) { return calc.nameValue(name, this); },
      parseRef(text) { return calc.parseRef(text, this); },
      formulaAt(sheet, row, col) { return calc.host.cell(sheet, row, col)?.f ?? null; },
      rowHidden(sheet, row) { return !!calc.host.rowHidden?.(sheet, row); },
    };
  }

  // ---- building ------------------------------------------------------------------------------------------------------------------

  /** Reads every formula of the workbook again and calculates all of them. */
  rebuild() {
    this.cells.clear();
    this.index.clear();
    this.spills.clear();
    this.anchors.clear();
    this.volatiles.clear();
    this.readNames();
    const all = [];
    this.host.formulas((sheet, row, col, text) => all.push(this.register(sheet, row, col, text)));
    this.evaluateSet(new Set(all));
    this.drainLate();
  }

  readNames() {
    this.names.clear();
    for (const n of this.host.names?.() || []) {
      let ast = null;
      try { ast = parse(n.ref).ast; } catch { ast = { type: 'err', code: '#NAME?' }; }
      this.names.set(n.name.toUpperCase(), { name: n.name, ast, computing: false });
    }
  }

  register(sheet, row, col, text) {
    let sheetCells = this.cells.get(sheet);
    if (!sheetCells) this.cells.set(sheet, (sheetCells = new Map()));
    const k = key(row, col);
    let fc = sheetCells.get(k);
    if (fc) this.unindex(fc);
    else sheetCells.set(k, (fc = new FormulaCell(sheet, row, col)));
    fc.text = text;
    try {
      fc.ast = parse(text).ast;
    } catch {
      fc.ast = { type: 'err', code: '#NAME?' };
    }
    this.indexDeps(fc);
    return fc;
  }

  /** Removes a formula cell; returns the area its spill covered (or null). */
  unregister(sheet, row, col) {
    const sheetCells = this.cells.get(sheet);
    const fc = sheetCells?.get(key(row, col));
    if (!fc) return null;
    const spill = fc.spill || fc.blocked;
    this.unindex(fc);
    this.clearSpill(fc);
    this.anchors.delete(fc);
    this.volatiles.delete(fc);
    sheetCells.delete(key(row, col));
    return spill || { r1: row, c1: col, r2: row, c2: col };
  }

  /** The areas a formula reads, put into the index (and whether it is volatile). */
  indexDeps(fc) {
    const deps = [];
    let volatile = false;
    const seenNames = new Set();
    const visit = (ast, sheet) => walk(ast, n => {
      if (n.type === 'ref') {
        const id = n.sheet === null ? sheet : this.host.sheetId(n.sheet);
        if (id !== null && id !== undefined) deps.push({ sheet: id, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 });
      } else if (n.type === 'fn') {
        if (FUNCTIONS.get(n.name)?.volatile) volatile = true;
      } else if (n.type === 'binary' && n.op === ':' && (n.left.type !== 'ref' || n.right.type !== 'ref')) {
        volatile = true;
      } else if (n.type === 'name') {
        const up = n.name.toUpperCase();
        const def = this.names.get(up);
        if (def && !seenNames.has(up)) { seenNames.add(up); visit(def.ast, sheet); }
        if (!def) volatile = true; // a name that may be defined later
      }
    });
    visit(fc.ast, fc.sheet);
    fc.volatile = volatile;
    if (volatile) this.volatiles.add(fc); else this.volatiles.delete(fc);
    for (const d of deps) this.addEntry(fc, d);
  }

  addEntry(fc, d) {
    let idx = this.index.get(d.sheet);
    if (!idx) this.index.set(d.sheet, (idx = { cell: new Map(), buckets: new Map(), big: new Set() }));
    const entry = { fc, ...d };
    if (d.r1 === d.r2 && d.c1 === d.c2) {
      const k = key(d.r1, d.c1);
      let set = idx.cell.get(k);
      if (!set) idx.cell.set(k, (set = new Set()));
      set.add(fc);
      fc.entries.push({ idx, k, single: true });
      return;
    }
    const rb1 = d.r1 >> 6;
    const rb2 = d.r2 >> 6;
    const cb1 = d.c1 >> 4;
    const cb2 = d.c2 >> 4;
    if ((rb2 - rb1 + 1) * (cb2 - cb1 + 1) > MAX_BUCKETS) {
      idx.big.add(entry);
      fc.entries.push({ idx, entry, big: true });
      return;
    }
    for (let rb = rb1; rb <= rb2; rb++) {
      for (let cb = cb1; cb <= cb2; cb++) {
        const bk = bucketKey(rb, cb);
        let set = idx.buckets.get(bk);
        if (!set) idx.buckets.set(bk, (set = new Set()));
        set.add(entry);
      }
    }
    fc.entries.push({ idx, entry, rb1, rb2, cb1, cb2 });
  }

  unindex(fc) {
    for (const e of fc.entries) {
      if (e.single) {
        const set = e.idx.cell.get(e.k);
        set?.delete(fc);
        if (set && !set.size) e.idx.cell.delete(e.k);
      } else if (e.big) {
        e.idx.big.delete(e.entry);
      } else {
        for (let rb = e.rb1; rb <= e.rb2; rb++) {
          for (let cb = e.cb1; cb <= e.cb2; cb++) {
            const bk = bucketKey(rb, cb);
            const set = e.idx.buckets.get(bk);
            set?.delete(e.entry);
            if (set && !set.size) e.idx.buckets.delete(bk);
          }
        }
      }
    }
    fc.entries = [];
  }

  /** The formulas that read the cell (sheet, row, col). */
  dependents(sheet, row, col, visit) {
    const idx = this.index.get(sheet);
    if (!idx) return;
    const single = idx.cell.get(key(row, col));
    if (single) for (const fc of single) visit(fc);
    const bucket = idx.buckets.get(bucketKey(row >> 6, col >> 4));
    if (bucket) for (const e of bucket) if (row >= e.r1 && row <= e.r2 && col >= e.c1 && col <= e.c2) visit(e.fc);
    for (const e of idx.big) if (row >= e.r1 && row <= e.r2 && col >= e.c1 && col <= e.c2) visit(e.fc);
  }

  /** The cells a formula writes: its own and its spill. */
  outputs(fc, visit) {
    visit(fc.row, fc.col);
    const s = fc.spill;
    if (!s) return;
    for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) if (r !== fc.row || c !== fc.col) visit(r, c);
  }

  // ---- changes ------------------------------------------------------------------------------------------------------------------

  /**
   * Cells were changed (typed, cleared, pasted …): positions [{ sheet, row, col }]. Formulas are read again from the host, then everything that
   * depends on the changes is calculated.
   */
  update(positions) {
    const dirty = new Set();
    const stack = [];
    const mark = fc => { if (!dirty.has(fc)) { dirty.add(fc); stack.push(fc); } };
    const freed = [];
    for (const p of positions) {
      const cell = this.host.cell(p.sheet, p.row, p.col);
      if (cell && typeof cell.f === 'string' && cell.f) {
        const existing = this.cells.get(p.sheet)?.get(key(p.row, p.col));
        if (!existing || existing.text !== cell.f) mark(this.register(p.sheet, p.row, p.col, cell.f));
        else mark(existing);
      } else {
        const gone = this.unregister(p.sheet, p.row, p.col);
        if (gone) freed.push({ sheet: p.sheet, ...gone });
      }
    }
    // the cells a removed formula spilled into are empty now
    for (const a of freed) {
      for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) this.dependents(a.sheet, r, c, mark);
    }
    for (const p of positions) {
      this.dependents(p.sheet, p.row, p.col, mark);
      for (const a of this.anchors) {
        const area = a.spill || a.blocked;
        if (a.sheet === p.sheet && area && p.row >= area.r1 && p.row <= area.r2 && p.col >= area.c1 && p.col <= area.c2 && !(p.row === a.row && p.col === a.col)) mark(a);
      }
    }
    for (const fc of this.volatiles) mark(fc);
    while (stack.length) {
      const fc = stack.pop();
      this.outputs(fc, (r, c) => this.dependents(fc.sheet, r, c, mark));
    }
    this.evaluateSet(dirty);
    this.drainLate();
  }

  /** Calculates the volatile formulas (NOW, RAND, INDIRECT …) and what depends on them: "Calculate now". */
  recalcVolatile() { this.update([]); }

  evaluateSet(set) {
    for (const fc of set) { fc.dirty = true; fc.indeg = 0; fc.state = 0; }
    const adjacency = new Map();
    for (const fc of set) {
      const next = new Set();
      this.outputs(fc, (r, c) => this.dependents(fc.sheet, r, c, g => { if (g !== fc && g.dirty) next.add(g); }));
      adjacency.set(fc, next);
      for (const g of next) g.indeg++;
    }
    const ready = [];
    for (const fc of set) if (fc.indeg === 0) ready.push(fc);
    ready.sort((a, b) => b.row - a.row || b.col - a.col);
    while (ready.length) {
      const fc = ready.pop();
      if (fc.dirty) this.compute(fc);
      for (const g of adjacency.get(fc)) if (--g.indeg === 0) ready.push(g);
    }
    // what is left waits for a circle (or for a cell behind one)
    for (const fc of set) if (fc.dirty) this.compute(fc);
  }

  drainLate() {
    let guard = 0;
    while (this.late.length && guard++ < 100000) {
      const fc = this.late.pop();
      if (fc.dirty && this.cells.get(fc.sheet)?.get(key(fc.row, fc.col)) === fc) this.compute(fc);
    }
    this.late = [];
  }

  compute(fc) {
    if (fc.state === COMPUTING) return;
    fc.state = COMPUTING;
    this.depth++;
    let v;
    try {
      if (this.depth > MAX_DEPTH) {
        v = Errors.CALC;
      } else {
        const ctx = Object.create(this.base);
        ctx.sheet = fc.sheet;
        ctx.row = fc.row;
        ctx.col = fc.col;
        v = result(evaluate(fc.ast, ctx), ctx);
      }
    } catch (e) {
      v = e instanceof RangeError ? Errors.CALC : Errors.VALUE;
    } finally {
      this.depth--;
      fc.state = 0;
      fc.dirty = false;
    }
    this.apply(fc, v);
  }

  apply(fc, v) {
    const before = fc.spill;
    if (Array.isArray(v)) {
      const rows = v.length;
      const cols = v[0].length;
      const area = { r1: fc.row, c1: fc.col, r2: fc.row + rows - 1, c2: fc.col + cols - 1 };
      if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS || this.blockedArea(fc, area)) {
        this.clearSpill(fc);
        fc.array = null;
        fc.blocked = area;
        fc.value = Errors.SPILL;
        this.anchors.add(fc);
      } else {
        this.clearSpill(fc);
        let covered = this.spills.get(fc.sheet);
        if (!covered) this.spills.set(fc.sheet, (covered = new Map()));
        for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) if (r !== fc.row || c !== fc.col) covered.set(key(r, c), fc);
        fc.array = v;
        fc.spill = area;
        fc.blocked = null;
        fc.value = v[0][0];
        this.anchors.add(fc);
      }
    } else {
      this.clearSpill(fc);
      fc.array = null;
      fc.blocked = null;
      fc.value = v;
      this.anchors.delete(fc);
    }
    this.host.setResult(fc.sheet, fc.row, fc.col, fc.value);
    const after = fc.spill;
    if (before || after) this.spillMoved(fc, before, after);
  }

  /** A spill grew, shrank or went: the formulas that read those cells and the other spills there are calculated again. */
  spillMoved(fc, before, after) {
    const visitArea = a => {
      if (!a) return;
      for (let r = a.r1; r <= a.r2; r++) {
        for (let c = a.c1; c <= a.c2; c++) {
          if (r === fc.row && c === fc.col) continue;
          this.dependents(fc.sheet, r, c, g => { if (g !== fc && !g.dirty && g.state !== COMPUTING) { g.dirty = true; this.late.push(g); } });
        }
      }
      for (const other of this.anchors) {
        const o = other.spill || other.blocked;
        if (other !== fc && other.sheet === fc.sheet && o && o.r1 <= a.r2 && a.r1 <= o.r2 && o.c1 <= a.c2 && a.c1 <= o.c2 && !other.dirty) {
          other.dirty = true;
          this.late.push(other);
        }
      }
    };
    if (before && after && before.r2 === after.r2 && before.c2 === after.c2) {
      visitArea(after); // same size, new values
      return;
    }
    visitArea(before);
    visitArea(after);
  }

  blockedArea(fc, area) {
    const covered = this.spills.get(fc.sheet);
    for (let r = area.r1; r <= area.r2; r++) {
      for (let c = area.c1; c <= area.c2; c++) {
        if (r === fc.row && c === fc.col) continue;
        const cell = this.host.cell(fc.sheet, r, c);
        if (cell && ((cell.v !== undefined && cell.v !== null && cell.v !== '') || cell.f)) return true;
        const owner = covered?.get(key(r, c));
        if (owner && owner !== fc) return true;
      }
    }
    return false;
  }

  clearSpill(fc) {
    const s = fc.spill;
    if (!s) return;
    const covered = this.spills.get(fc.sheet);
    for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) if (covered?.get(key(r, c)) === fc) covered.delete(key(r, c));
    fc.spill = null;
  }

  // ---- values ------------------------------------------------------------------------------------------------------------------

  /** The value of a cell as formulas see it (a formula's result, a spilled value, a constant, or null). */
  valueAt(sheet, row, col) {
    const k = key(row, col);
    const fc = this.cells.get(sheet)?.get(k);
    if (fc) {
      if (fc.state === COMPUTING) return Errors.CIRC;
      if (fc.dirty) this.compute(fc);
      return fc.value;
    }
    const owner = this.spills.get(sheet)?.get(k);
    if (owner) {
      if (owner.state === COMPUTING) return Errors.CIRC;
      if (owner.dirty) this.compute(owner);
      const s = owner.spill;
      if (s && row >= s.r1 && row <= s.r2 && col >= s.c1 && col <= s.c2) return owner.array[row - owner.row][col - owner.col];
    }
    const cell = this.host.cell(sheet, row, col);
    if (!cell || cell.v === undefined || cell.v === null) return null;
    if (cell.t === 'e') return errorOf(cell.v) || Errors.VALUE;
    return cell.v;
  }

  /** The used area of a sheet, spills included. */
  extent(sheet) {
    const b = this.host.bounds(sheet);
    let rows = b.rows;
    let cols = b.cols;
    for (const a of this.anchors) if (a.sheet === sheet && a.spill) { rows = Math.max(rows, a.spill.r2 + 1); cols = Math.max(cols, a.spill.c2 + 1); }
    return { rows, cols };
  }

  /** The values of a reference as a 2-D array (a huge one, e.g. A:A, is cut at the used area). */
  values(ref) {
    let { r2, c2 } = ref;
    if (ref.height * ref.width > CLIP_CELLS) {
      const e = this.extent(ref.sheet);
      r2 = Math.max(ref.r1, Math.min(r2, e.rows - 1));
      c2 = Math.max(ref.c1, Math.min(c2, e.cols - 1));
    }
    const out = new Array(r2 - ref.r1 + 1);
    for (let r = ref.r1; r <= r2; r++) {
      const line = new Array(c2 - ref.c1 + 1);
      for (let c = ref.c1; c <= c2; c++) line[c - ref.c1] = this.valueAt(ref.sheet, r, c);
      out[r - ref.r1] = line;
    }
    return out;
  }

  nameValue(name, ctx) {
    const def = this.names.get(String(name).toUpperCase());
    if (!def) return Errors.NAME;
    if (def.computing) return Errors.CIRC;
    def.computing = true;
    try {
      return evaluate(def.ast, ctx);
    } finally {
      def.computing = false;
    }
  }

  /** INDIRECT: a reference written as text ("B3", "Sheet2!A1:B4", a defined name). */
  parseRef(text, ctx) {
    let ast;
    try { ast = parse(String(text).trim()).ast; } catch { return Errors.REF; }
    if (ast.type === 'ref' || (ast.type === 'binary' && ast.op === ':') || ast.type === 'name') {
      const v = evaluate(ast, ctx);
      return v instanceof Ref ? v : Errors.REF;
    }
    return Errors.REF;
  }

  // ---- for the editor ------------------------------------------------------------------------------------------------------------

  /** The formula cell at a place (or the anchor whose spill covers it): { fc, anchor: bool }. */
  formulaAt(sheet, row, col) {
    const fc = this.cells.get(sheet)?.get(key(row, col));
    if (fc) return fc;
    return this.spills.get(sheet)?.get(key(row, col)) || null;
  }

  /** Evaluates a formula that is not in a cell (conditional formats, validations, a preview), as if it were in (sheet, row, col). */
  evaluateText(text, sheet, row, col) {
    let ast = this.asts.get(text);
    if (!ast) {
      try { ast = parse(text).ast; } catch { ast = { type: 'err', code: '#NAME?' }; }
      if (this.asts.size > 5000) this.asts.clear();
      this.asts.set(text, ast);
    }
    const ctx = Object.create(this.base);
    ctx.sheet = sheet;
    ctx.row = row;
    ctx.col = col;
    try {
      return result(evaluate(ast, ctx), ctx);
    } catch {
      return Errors.VALUE;
    }
  }

  /** The values of every cell of a spill area (for copy and the display): the array of the anchor. */
  spillArray(fc) { return fc.array; }
}

export { FormulaError };
