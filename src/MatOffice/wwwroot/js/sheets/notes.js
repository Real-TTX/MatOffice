/* MatOffice Sheets – notes on cells (plain text): a red corner marks them, they show while the pointer rests on the cell; "New note" opens
   a yellow box next to the cell to type into; "Show all notes" keeps every box of the sheet visible. */
import { el, t } from '../editor/util.js';

export class Notes {
  constructor(app) {
    this.app = app;
    this.showAll = false;
    this.boxes = [];
    this.editing = null;
  }

  /** Opens the note of the active cell for editing (a new one when it has none). */
  edit() {
    const app = this.app;
    if (app.readOnly) return;
    if (app.editor.active && !app.editor.commit(0, 0)) return;
    this.close(true);
    const sheet = app.sheet;
    const { row, col } = app.sel.active;
    const before = sheet.cell(row, col)?.note || '';
    const area = el('textarea', { class: 'mo-sheets-notebox__input', 'aria-label': t('Sheets.Note'), spellcheck: 'true' });
    area.value = before;
    const box = el('div', { class: 'mo-sheets-notebox is-editing' }, el('div', { class: 'mo-sheets-notebox__author', text: app.ctx.doc.user?.name || '' }), area);
    app.grid.overlay.append(box);
    this.editing = { box, area, sheet, row, col, before };
    this.place(box, row, col);
    area.focus();
    area.setSelectionRange(area.value.length, area.value.length);
    area.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); this.close(false); app.focusGrid(); }
    });
    area.addEventListener('blur', () => setTimeout(() => this.close(true), 0));
  }

  close(save) {
    const ed = this.editing;
    if (!ed) return;
    this.editing = null;
    ed.box.remove();
    const text = ed.area.value.replace(/\s+$/, '');
    if (save && text !== ed.before) {
      this.app.commands.change(text ? 'note' : 'delete note', () => this.app.wb.patchCell(ed.sheet, ed.row, ed.col, { note: text || undefined }), { protect: false });
    }
    this.render();
  }

  place(box, row, col) {
    const r = this.app.grid.cellRect(row, col);
    box.style.left = Math.round(r.x + r.w + 10) + 'px';
    box.style.top = Math.max(0, Math.round(r.y - 4)) + 'px';
  }

  toggleAll() {
    this.showAll = !this.showAll;
    this.render();
  }

  /** The boxes of "Show all notes" (placed again on every render of the grid). */
  render() {
    for (const b of this.boxes) b.remove();
    this.boxes = [];
    if (!this.showAll) return;
    const app = this.app;
    const sheet = app.sheet;
    const m = app.grid.metrics();
    for (const [k, cell] of sheet.cells) {
      if (!cell.note) continue;
      const row = Math.floor(k / 16384);
      const col = k % 16384;
      if (sheet.rows.isHidden(row) || sheet.cols.isHidden(col)) continue;
      const r = app.grid.cellRect(row, col, m);
      if (r.x > app.grid.width || r.y > app.grid.height || r.x + r.w < 0 || r.y + r.h < 0) continue;
      const box = el('div', { class: 'mo-sheets-notebox' });
      box.textContent = cell.note;
      box.style.left = Math.round(r.x + r.w + 10) + 'px';
      box.style.top = Math.max(0, Math.round(r.y - 4)) + 'px';
      app.grid.overlay.append(box);
      this.boxes.push(box);
      if (this.boxes.length > 200) break;
    }
  }

  /** Called by every render: the boxes follow the cells. */
  position() {
    if (this.editing) this.place(this.editing.box, this.editing.row, this.editing.col);
    if (this.showAll) this.render();
  }
}
