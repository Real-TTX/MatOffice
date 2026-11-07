/* MatOffice Base – the Relationships window of Access: the tables as boxes (drag them by the title, the key fields with a key icon), a field dragged
   onto a field of another table opens Edit Relationships (Enforce Referential Integrity, Cascade Update, Cascade Delete, the relationship type,
   Join Type), the lines carry 1 and ∞ where integrity is enforced; a line is edited by a double click and deleted with Delete. */
import { t, el, icon, dialog } from '../editor/util.js';
import { BaseView } from './view.js';
import {
  tableById, fieldById, relationshipKind, orientRelationship, relationshipProblem, saveRelationship, newObjectId, isUniqueField,
} from './model.js';
import { showProblem, ask, field, selectInput, checkInput, radioGroup, contextMenu, drag, iconButton } from './ui.js';

const BOX_W = 180;
const ROW_H = 22;
const HEAD_H = 28;
const MAX_ROWS = 12;

export class RelationshipsView extends BaseView {
  constructor(app, tab) {
    super(app, tab);
    this.selected = null;
    this.selectedTable = null;
  }

  get ro() { return this.app.readOnly; }
  get positions() { return this.db.ui.relationships; }

  mount() {
    this.canvas = el('div', { class: 'mo-base-rel', tabindex: '0', 'aria-label': t('Base.Rel.Title') });
    this.svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    this.svg.setAttribute('class', 'mo-base-lines');
    this.canvas.append(this.svg);
    this.el.append(this.canvas);
    this.canvas.addEventListener('pointerdown', e => { if (e.target === this.canvas || e.target === this.svg) { this.selected = null; this.selectedTable = null; this.render(); this.app.ribbonRefresh(); } });
    this.canvas.addEventListener('keydown', e => {
      if ((e.key === 'Delete' || e.key === 'Backspace') && this.selected) { e.preventDefault(); this.deleteRelationship(this.selected); }
      if ((e.key === 'Delete' || e.key === 'Backspace') && !this.selected && this.selectedTable) { e.preventDefault(); this.hideTable(this.selectedTable); }
    });
    this.canvas.addEventListener('contextmenu', e => {
      if (e.target !== this.canvas && e.target !== this.svg) return;
      e.preventDefault();
      contextMenu({ x: e.clientX, y: e.clientY }, [
        { label: t('Base.Rel.AddTables'), icon: 'base-show-table', run: () => this.addTables(), disabled: this.ro },
        { label: t('Base.Rel.All'), icon: 'base-relationships', run: () => this.showAll() },
        { label: t('Base.Rel.Clear'), icon: 'base-clear-layout', run: () => this.clearLayout(), disabled: this.ro },
      ]);
    });
    if (!Object.keys(this.positions).length) this.autoLayout(this.involvedTables(), false);
    this.render();
  }

  refresh() { this.render(); }

  focus() { this.canvas?.focus({ preventScroll: true }); }

  involvedTables() {
    const ids = new Set();
    for (const r of this.db.relationships) { ids.add(r.parent.table); ids.add(r.child.table); }
    return this.db.tables.filter(x => ids.has(x.id)).map(x => x.id);
  }

  /** Places the tables in rows (parents left of their children, like the layout people make by hand). */
  autoLayout(ids, save = true) {
    const placed = new Set(Object.keys(this.positions));
    const order = [...ids].sort((a, b) => this.depth(a) - this.depth(b));
    const perColumn = new Map();
    for (const id of order) {
      if (placed.has(id)) continue;
      const d = this.depth(id);
      const n = perColumn.get(d) || 0;
      perColumn.set(d, n + 1);
      this.positions[id] = { x: 30 + d * (BOX_W + 90), y: 24 + n * 260 };
    }
    if (save && !this.ro) this.app.settingsChanged();
  }

  depth(id, seen = new Set()) {
    if (seen.has(id)) return 0;
    seen.add(id);
    const parents = this.db.relationships.filter(r => r.child.table === id && r.parent.table !== id).map(r => r.parent.table);
    return parents.length ? 1 + Math.max(...parents.map(p => this.depth(p, seen))) : 0;
  }

  boxHeight(table) { return HEAD_H + Math.min(MAX_ROWS, table.fields.length) * ROW_H + 6; }

  render() {
    for (const n of [...this.canvas.querySelectorAll('.mo-base-box')]) n.remove();
    for (const [id, pos] of Object.entries(this.positions)) {
      const table = tableById(this.db, id);
      if (!table) continue;
      const head = el('div', { class: 'mo-base-box__head' }, icon('base-table'), el('span', { text: table.name }),
        this.ro ? null : iconButton('x', t('Base.Rel.HideTable', table.name), () => this.hideTable(id), { class: 'mo-base-ibtn mo-base-box__x' }));
      const list = el('div', { class: 'mo-base-box__fields' });
      for (const f of table.fields) {
        const pk = table.primaryKey.includes(f.id);
        const row = el('div', { class: 'mo-base-box__field' + (pk ? ' is-key' : ''), 'data-field': f.id, title: f.name }, pk ? icon('key') : el('span', { class: 'mo-base-box__nokey' }), el('span', { text: f.name }));
        row.addEventListener('pointerdown', e => this.dragField(e, id, f.id));
        list.append(row);
      }
      list.addEventListener('scroll', () => this.drawLines());
      const box = el('div', {
        class: 'mo-base-box' + (this.selectedTable === id ? ' is-selected' : ''), 'data-table': id,
        style: { left: pos.x + 'px', top: pos.y + 'px', width: BOX_W + 'px', height: this.boxHeight(table) + 'px' },
      }, head, list);
      head.addEventListener('pointerdown', e => { if (!e.target.closest('button')) { this.selectedTable = id; this.selected = null; box.classList.add('is-selected'); this.moveBox(e, id, box); } });
      head.addEventListener('contextmenu', e => {
        e.preventDefault();
        contextMenu({ x: e.clientX, y: e.clientY }, [
          { label: t('Base.Rel.Direct'), icon: 'base-relationships', run: () => this.showDirect(id) },
          { label: t('Base.Rel.HideTable', table.name), icon: 'eye-off', run: () => this.hideTable(id), disabled: this.ro },
          '-',
          { label: t('Base.View.Design'), icon: 'base-v-design', run: () => this.app.open('tables', id, 'design') },
        ]);
      });
      this.canvas.append(box);
    }
    this.drawLines();
  }

  moveBox(e, id, node) {
    e.preventDefault();
    const pos = this.positions[id];
    const x0 = pos.x;
    const y0 = pos.y;
    drag(e, (dx, dy) => { pos.x = Math.max(0, Math.round(x0 + dx)); pos.y = Math.max(0, Math.round(y0 + dy)); node.style.left = pos.x + 'px'; node.style.top = pos.y + 'px'; this.drawLines(); },
      (dx, dy) => { if ((dx || dy) && !this.ro) this.app.settingsChanged(); this.app.ribbonRefresh(); });
  }

  fieldPoint(tableId, fieldId, side) {
    const box = this.canvas.querySelector(`.mo-base-box[data-table="${tableId}"]`);
    const pos = this.positions[tableId];
    if (!box || !pos) return null;
    const list = box.querySelector('.mo-base-box__fields');
    const row = list.querySelector(`[data-field="${fieldId}"]`);
    // the row's offsetTop counts from the box (its positioned parent), the head included; 1 px of border above it
    let y = pos.y + 1 + (row ? row.offsetTop - list.scrollTop + ROW_H / 2 : HEAD_H + 8);
    y = Math.max(pos.y + HEAD_H, Math.min(pos.y + box.offsetHeight - 6, y));
    return { x: side === 'right' ? pos.x + BOX_W : pos.x, y };
  }

  drawLines() {
    const svg = this.svg;
    svg.setAttribute('width', Math.max(this.canvas.scrollWidth, this.canvas.clientWidth));
    svg.setAttribute('height', Math.max(this.canvas.scrollHeight, this.canvas.clientHeight));
    svg.replaceChildren();
    const ns = 'http://www.w3.org/2000/svg';
    for (const r of this.db.relationships) {
      const pp = this.positions[r.parent.table];
      const cp = this.positions[r.child.table];
      if (!pp || !cp) continue;
      const parentLeft = pp.x + BOX_W / 2 <= cp.x + BOX_W / 2;
      const a = this.fieldPoint(r.parent.table, r.parent.field, parentLeft ? 'right' : 'left');
      const b = this.fieldPoint(r.child.table, r.child.field, parentLeft ? 'left' : 'right');
      if (!a || !b) continue;
      const stub = 18;
      const ax = a.x + (parentLeft ? stub : -stub);
      const bx = b.x + (parentLeft ? -stub : stub);
      const d = `M${a.x},${a.y} L${ax},${a.y} L${bx},${b.y} L${b.x},${b.y}`;
      const g = document.createElementNS(ns, 'g');
      g.setAttribute('class', 'mo-base-line' + (this.selected === r.id ? ' is-selected' : '') + (r.integrity ? ' is-integrity' : ''));
      const hit = document.createElementNS(ns, 'path');
      hit.setAttribute('d', d);
      hit.setAttribute('class', 'mo-base-line__hit');
      const path = document.createElementNS(ns, 'path');
      path.setAttribute('d', d);
      path.setAttribute('class', 'mo-base-line__path');
      g.append(hit, path);
      if (r.integrity) {
        const kind = relationshipKind(this.db, r);
        const label = (x, y, text) => {
          const tx = document.createElementNS(ns, 'text');
          tx.setAttribute('x', x);
          tx.setAttribute('y', y);
          tx.setAttribute('class', 'mo-base-line__label');
          tx.textContent = text;
          return tx;
        };
        g.append(label(a.x + (parentLeft ? 4 : -12), a.y - 4, '1'), label(b.x + (parentLeft ? -14 : 4), b.y - 4, kind === 'oneToOne' ? '1' : '∞'));
      }
      g.addEventListener('pointerdown', e => { e.stopPropagation(); this.selected = r.id; this.selectedTable = null; this.drawLines(); this.canvas.focus({ preventScroll: true }); this.app.ribbonRefresh(); });
      g.addEventListener('dblclick', () => this.edit(r));
      g.addEventListener('contextmenu', e => {
        e.preventDefault();
        this.selected = r.id;
        this.drawLines();
        contextMenu({ x: e.clientX, y: e.clientY }, [
          { label: t('Base.Rel.Edit'), icon: 'edit', run: () => this.edit(r) },
          { label: t('Base.Rel.Delete'), icon: 'trash', danger: true, run: () => this.deleteRelationship(r.id), disabled: this.ro },
        ]);
      });
      svg.append(g);
    }
  }

  dragField(e, tableId, fieldId) {
    if (e.button !== 0) return;
    e.preventDefault();
    let ghost = null;
    const table = tableById(this.db, tableId);
    drag(e, (dx, dy, ev) => {
      if (this.ro) return;
      if (!ghost && Math.abs(dx) + Math.abs(dy) > 4) { ghost = el('div', { class: 'mo-base-dragghost', text: table.name + '.' + fieldById(table, fieldId)?.name }); document.body.append(ghost); }
      if (ghost) { ghost.style.left = ev.clientX + 8 + 'px'; ghost.style.top = ev.clientY + 8 + 'px'; }
    }, (dx, dy, ev) => {
      if (!ghost) return;
      ghost.remove();
      const target = document.elementFromPoint(ev.clientX, ev.clientY);
      const row = target?.closest('.mo-base-box__field');
      const box = target?.closest('.mo-base-box');
      if (!row || !box || (box.dataset.table === tableId && row.dataset.field === fieldId)) return;
      const pair = orientRelationship(this.db, { table: tableId, field: fieldId }, { table: box.dataset.table, field: row.dataset.field });
      this.edit({ id: null, ...pair, integrity: false, cascadeUpdate: false, cascadeDelete: false, join: 1 });
    });
  }

  /** The dialog Edit Relationships (a new relationship when rel.id is null). */
  async edit(rel) {
    const db = this.db;
    const parent = tableById(db, rel.parent.table);
    const child = tableById(db, rel.child.table);
    if (!parent || !child) return;
    const pf = selectInput(parent.fields.map(f => ({ value: f.id, label: f.name })), rel.parent.field, { disabled: this.ro ? true : null });
    const cf = selectInput(child.fields.map(f => ({ value: f.id, label: f.name })), rel.child.field, { disabled: this.ro ? true : null });
    const integrity = checkInput(t('Base.Rel.Integrity'), rel.integrity, { disabled: this.ro ? true : null });
    const cu = checkInput(t('Base.Rel.CascadeUpdate'), rel.cascadeUpdate, { disabled: this.ro ? true : null });
    const cd = checkInput(t('Base.Rel.CascadeDelete'), rel.cascadeDelete, { disabled: this.ro ? true : null });
    const kind = el('dd');
    let join = rel.join || 1;
    const draft = () => ({ ...rel, parent: { table: parent.id, field: pf.value }, child: { table: child.id, field: cf.value }, integrity: integrity.input.checked, cascadeUpdate: cu.input.checked, cascadeDelete: cd.input.checked, join });
    const sync = () => {
      cu.input.disabled = cd.input.disabled = this.ro || !integrity.input.checked;
      if (!integrity.input.checked) { cu.input.checked = false; cd.input.checked = false; }
      const k = relationshipKind(db, draft());
      kind.textContent = { oneToMany: t('Base.Rel.OneToMany'), oneToOne: t('Base.Rel.OneToOne'), indeterminate: t('Base.Rel.Indeterminate') }[k];
    };
    [pf, cf, integrity.input].forEach(x => x.addEventListener('change', sync));
    sync();
    const joinButton = el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: async () => { const j = await this.joinType(parent, child, join); if (j) join = j; } }, icon('base-join'), el('span', { text: t('Base.Rel.JoinType') }));
    const body = el('div', { class: 'mo-base-form mo-base-reledit' },
      el('div', { class: 'mo-base-reledit__pair' },
        el('div', {}, el('div', { class: 'mo-base-reledit__head', text: t('Base.Rel.Table') }), el('div', { class: 'mo-base-reledit__name', text: parent.name }), pf),
        el('div', {}, el('div', { class: 'mo-base-reledit__head', text: t('Base.Rel.RelatedTable') }), el('div', { class: 'mo-base-reledit__name', text: child.name }), cf)),
      integrity, el('div', { class: 'mo-base-reledit__cascade' }, cu, cd),
      el('dl', { class: 'mat-kv mo-base-kv' }, el('dt', { text: t('Base.Rel.Type') }), kind),
      joinButton);
    const buttons = this.ro ? [{ label: t('Editor.Close'), value: 'cancel', kind: 'primary' }]
      : [{ label: rel.id ? t('Editor.Ok') : t('Base.Rel.Create'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }];
    const r = await dialog({
      title: t('Base.Rel.EditTitle'), icon: 'base-relationships', size: 'md', body, buttons,
      validate: () => {
        const p = relationshipProblem(db, { ...draft(), id: rel.id || '__new' });
        if (p) { showProblem(p); return false; }
        return true;
      },
    });
    if (r !== 'ok' || this.ro) return;
    const next = draft();
    if (!rel.id) next.id = newObjectId(db, 'relationships');
    const result = this.app.design(() => saveRelationship(db, next));
    if (result && !result.ok) { await showProblem(result.error); return; }
    this.selected = next.id;
    this.render();
  }

  async joinType(parent, child, current) {
    const choice = radioGroup('mo-base-reljoin', [
      { value: '1', label: t('Base.Join.Inner') },
      { value: '2', label: t('Base.Join.Left', parent.name, child.name) },
      { value: '3', label: t('Base.Join.Right', child.name, parent.name) },
    ], String(current));
    const r = await dialog({ title: t('Base.Rel.JoinType'), icon: 'base-join', size: 'md', body: choice, buttons: [{ label: t('Editor.Ok'), value: 'ok', kind: 'primary' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
    return r === 'ok' ? Number(choice.value()) : null;
  }

  async deleteRelationship(id) {
    if (this.ro) return;
    if (!(await ask(t('Base.Rel.DeleteAsk'), { danger: true, ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    this.app.design(() => { this.db.relationships = this.db.relationships.filter(r => r.id !== id); });
    this.selected = null;
    this.render();
  }

  hideTable(id) {
    if (this.ro) return;
    delete this.positions[id];
    this.selectedTable = null;
    this.app.settingsChanged();
    this.render();
  }

  async addTables() {
    if (this.ro) return;
    const list = el('div', { class: 'mo-base-showtable', role: 'listbox', 'aria-multiselectable': 'true' },
      ...this.db.tables.map(x => {
        const b = el('button', { type: 'button', class: 'mo-base-showtable__item', role: 'option', 'aria-selected': 'false', 'data-id': x.id }, icon('base-table'), el('span', { text: x.name }));
        b.addEventListener('click', () => b.setAttribute('aria-selected', b.getAttribute('aria-selected') === 'true' ? 'false' : 'true'));
        b.addEventListener('dblclick', () => { this.place([x.id]); });
        return b;
      }));
    const r = await dialog({ title: t('Base.Rel.AddTables'), icon: 'base-show-table', body: list, buttons: [{ label: t('Base.Query.Add'), value: 'ok', kind: 'primary' }, { label: t('Editor.Close'), value: 'cancel' }] });
    if (r === 'ok') this.place([...list.querySelectorAll('[aria-selected="true"]')].map(b => b.dataset.id));
  }

  place(ids) {
    let x = Object.values(this.positions).reduce((m, p) => Math.max(m, p.x + BOX_W + 60), 30);
    for (const id of ids) {
      if (this.positions[id]) continue;
      this.positions[id] = { x, y: 30 };
      x += BOX_W + 60;
    }
    this.app.settingsChanged();
    this.render();
  }

  showAll() {
    this.autoLayout(this.involvedTables());
    this.render();
  }

  showDirect(id) {
    const ids = new Set([id]);
    for (const r of this.db.relationships) { if (r.parent.table === id) ids.add(r.child.table); if (r.child.table === id) ids.add(r.parent.table); }
    this.autoLayout([...ids]);
    this.render();
  }

  async clearLayout() {
    if (!(await ask(t('Base.Rel.ClearAsk'), { ok: t('Base.Yes'), cancel: t('Base.No') }))) return;
    for (const k of Object.keys(this.positions)) delete this.positions[k];
    this.app.settingsChanged();
    this.render();
  }

  contextTabs() {
    return [{
      id: 'reldesign', label: t('Base.Tab.RelDesign'), groups: [
        { label: t('Base.Group.Tools'), items: [
          { type: 'button', size: 'large', label: t('Base.Rel.Edit'), icon: 'edit', edit: false, disabled: () => !this.selected, run: () => { const r = this.db.relationships.find(x => x.id === this.selected); if (r) this.edit(r); } },
          { type: 'button', size: 'large', label: t('Base.Rel.Clear'), icon: 'base-clear-layout', run: () => this.clearLayout() },
          { type: 'button', size: 'large', label: t('Base.Rel.Report'), icon: 'print', edit: false, run: () => this.app.print() },
        ] },
        { label: t('Base.Group.Relationships'), items: [
          { type: 'button', size: 'large', label: t('Base.Rel.AddTables'), icon: 'base-show-table', run: () => this.addTables() },
          { type: 'stack', rows: [
            [{ type: 'button', showLabel: true, label: t('Base.Rel.Hide'), icon: 'eye-off', disabled: () => !this.selectedTable, run: () => this.hideTable(this.selectedTable) }],
            [{ type: 'button', showLabel: true, label: t('Base.Rel.Direct'), icon: 'base-relationships', edit: false, disabled: () => !this.selectedTable, run: () => this.showDirect(this.selectedTable) }],
            [{ type: 'button', showLabel: true, label: t('Base.Rel.All'), icon: 'base-relationships', edit: false, run: () => this.showAll() }],
          ] },
          { type: 'button', size: 'large', label: t('Base.Rel.Close'), icon: 'x', edit: false, run: () => this.app.closeTab(this.tab) },
        ] },
      ],
    }];
  }

  status() { return [t('Base.Rel.Count', this.db.relationships.length)]; }
}

export { isUniqueField };
