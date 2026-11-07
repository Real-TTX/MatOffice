/* MatOffice Sheets – the sizes along one axis of a sheet: column widths or row heights (pure, no DOM).
   Most lines have the default size; the others are kept in a map. offset(i) and indexAt(px) use a sorted list with running sums (binary search),
   so a sheet with a million rows scrolls as fast as a small one. A hidden line (by hand or by the AutoFilter) has the size 0. */

export class Axis {
  constructor(def, count) {
    this.def = def;
    this.count = count;
    this.sizes = new Map();
    this.hidden = new Set();
    this.filtered = new Set();
    this.dirty = true;
    this.keys = [];
    this.deltas = [];
  }

  /** The size of line i on the screen (0 when hidden). */
  size(i) { return this.hidden.has(i) || this.filtered.has(i) ? 0 : this.sizes.get(i) ?? this.def; }

  /** The size of line i when it is shown. */
  rawSize(i) { return this.sizes.get(i) ?? this.def; }

  isHidden(i) { return this.hidden.has(i) || this.filtered.has(i); }

  setSize(i, px) {
    const v = Math.max(0, Math.round(px));
    if (v === this.def) this.sizes.delete(i); else this.sizes.set(i, v);
    this.dirty = true;
  }

  setHidden(i, on) {
    if (on) this.hidden.add(i); else this.hidden.delete(i);
    this.dirty = true;
  }

  setFiltered(set) {
    this.filtered = set;
    this.dirty = true;
  }

  setDefault(px) {
    this.def = Math.max(1, Math.round(px));
    this.dirty = true;
  }

  build() {
    if (!this.dirty) return;
    const special = new Set([...this.sizes.keys(), ...this.hidden, ...this.filtered]);
    this.keys = [...special].sort((a, b) => a - b);
    this.deltas = new Array(this.keys.length + 1);
    this.deltas[0] = 0;
    for (let k = 0; k < this.keys.length; k++) this.deltas[k + 1] = this.deltas[k] + this.size(this.keys[k]) - this.def;
    this.dirty = false;
  }

  /** The number of special lines before i. */
  rank(i) {
    let lo = 0;
    let hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1; else hi = mid;
    }
    return lo;
  }

  /** The position where line i starts. */
  offset(i) {
    this.build();
    return i * this.def + this.deltas[this.rank(i)];
  }

  /** The total size of lines [0, n). */
  total(n = this.count) { return this.offset(n); }

  /** The line at a position (the last line for positions past the end). Hidden lines are never returned. */
  indexAt(pos) {
    this.build();
    if (pos <= 0) return this.firstVisible(0);
    // the last special line that starts at or before pos
    let lo = 0;
    let hi = this.keys.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const start = this.keys[mid] * this.def + this.deltas[mid];
      if (start <= pos) { found = mid; lo = mid + 1; } else hi = mid - 1;
    }
    let i;
    if (found < 0) {
      i = Math.floor(pos / this.def);
    } else {
      const k = this.keys[found];
      const start = k * this.def + this.deltas[found];
      const size = this.size(k);
      if (pos < start + size) return k;
      i = k + 1 + Math.floor((pos - start - size) / this.def);
    }
    return Math.min(this.firstVisible(i), this.count - 1);
  }

  /** The first visible line at or after i. */
  firstVisible(i) {
    let j = i;
    while (j < this.count - 1 && this.isHidden(j)) j++;
    return j;
  }

  /** The next visible line from i in a direction (-1, +1); i itself when there is none. */
  step(i, dir) {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.isHidden(j)) j += dir;
    return j < 0 || j >= this.count ? i : j;
  }

  /** Lines inserted at `at`: everything from there moves by `count`. */
  insert(at, count) {
    const move = (map, isSet) => {
      const next = isSet ? new Set() : new Map();
      for (const entry of map) {
        const [k, v] = isSet ? [entry, null] : entry;
        const nk = k >= at ? k + count : k;
        if (nk < this.count) isSet ? next.add(nk) : next.set(nk, v);
      }
      return next;
    };
    this.sizes = move(this.sizes, false);
    this.hidden = move(this.hidden, true);
    this.filtered = move(this.filtered, true);
    this.dirty = true;
  }

  /** Lines [at, at + count) removed: what follows moves up. */
  remove(at, count) {
    const end = at + count;
    const move = (map, isSet) => {
      const next = isSet ? new Set() : new Map();
      for (const entry of map) {
        const [k, v] = isSet ? [entry, null] : entry;
        if (k >= at && k < end) continue;
        const nk = k >= end ? k - count : k;
        isSet ? next.add(nk) : next.set(nk, v);
      }
      return next;
    };
    this.sizes = move(this.sizes, false);
    this.hidden = move(this.hidden, true);
    this.filtered = move(this.filtered, true);
    this.dirty = true;
  }

  snapshot() { return { def: this.def, sizes: [...this.sizes], hidden: [...this.hidden], filtered: [...this.filtered] }; }

  restore(s) {
    this.def = s.def;
    this.sizes = new Map(s.sizes);
    this.hidden = new Set(s.hidden);
    this.filtered = new Set(s.filtered);
    this.dirty = true;
  }
}
