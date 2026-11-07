/* MatOffice Sheets – the sheet tabs at the bottom, like Excel: click to show a sheet (also while a formula is typed: references to that
   sheet), + adds one, a double click renames it in place, dragging moves it, the context menu inserts, deletes, renames, duplicates, colours,
   hides and shows sheets. */
import { el, icon, t } from '../editor/util.js';
import { openPopover, colourPalette } from '../editor/frame.js';

export class Tabs {
  constructor(app) {
    this.app = app;
    this.list = el('div', { class: 'mo-sheets-tabs__list', role: 'tablist', 'aria-label': t('Sheets.Sheets') });
    this.add = el('button', { type: 'button', class: 'mo-sheets-tabs__add', 'aria-label': t('Sheets.Sheet.Add'), 'data-tip': t('Sheets.Sheet.Add') + ' (Shift+F11)' }, icon('plus'));
    this.prev = el('button', { type: 'button', class: 'mo-sheets-tabs__nav', 'aria-label': t('Sheets.Sheet.Prev') }, icon('chevron-left'));
    this.next = el('button', { type: 'button', class: 'mo-sheets-tabs__nav', 'aria-label': t('Sheets.Sheet.Next') }, icon('chevron-right'));
    this.root = el('div', { class: 'mo-sheets-tabs' }, this.prev, this.next, this.list, this.add);
    this.add.addEventListener('click', () => app.commands.addSheet());
    this.prev.addEventListener('click', () => { this.list.scrollLeft -= 120; });
    this.next.addEventListener('click', () => { this.list.scrollLeft += 120; });
    this.list.addEventListener('wheel', e => { if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) { this.list.scrollLeft += e.deltaY; e.preventDefault(); } }, { passive: false });
  }

  render() {
    const app = this.app;
    this.add.hidden = app.readOnly;
    const tabs = [];
    app.wb.sheets.forEach((sheet, index) => {
      if (sheet.hidden) return;
      const active = sheet === app.sheet;
      const tab = el('button', {
        type: 'button', role: 'tab', class: 'mo-sheets-tab' + (active ? ' is-active' : '') + (sheet.color ? ' has-color' : ''),
        'aria-selected': active ? 'true' : 'false', 'data-index': index, title: sheet.name,
      }, el('span', { class: 'mo-sheets-tab__name', text: sheet.name }), sheet.protected ? icon('lock', 'mo-sheets-tab__lock') : null);
      if (sheet.color) tab.style.setProperty('--tab-color', sheet.color);
      tab.addEventListener('click', () => this.activate(sheet));
      tab.addEventListener('dblclick', () => { if (!app.readOnly) this.rename(sheet, tab); });
      tab.addEventListener('contextmenu', e => { e.preventDefault(); this.activate(sheet); app.contextMenu(e.clientX, e.clientY, this.menu(sheet)); });
      tab.addEventListener('pointerdown', e => this.drag(e, sheet, tab));
      tab.addEventListener('keydown', e => {
        if (e.key === 'F2' && !app.readOnly) { e.preventDefault(); this.rename(sheet, tab); }
        if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') { e.preventDefault(); app.commands.nextSheet(e.key === 'ArrowRight' ? 1 : -1); this.list.querySelector('.is-active')?.focus(); }
      });
      tabs.push(tab);
    });
    this.list.replaceChildren(...tabs);
    const active = this.list.querySelector('.is-active');
    active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  activate(sheet) {
    const app = this.app;
    if (sheet === app.sheet) return;
    // while a formula is typed, another sheet can be shown to point at its cells
    if (app.editor.active && !(app.editor.canPoint() || app.editor.pointing)) { if (!app.editor.commit(0, 0)) return; }
    app.showSheet(sheet, { keepEditor: app.editor.active });
  }

  rename(sheet, tab) {
    const app = this.app;
    const input = el('input', { class: 'mo-sheets-tab__input', value: sheet.name, maxlength: 31, 'aria-label': t('Sheets.Sheet.Rename') });
    tab.replaceChildren(input);
    input.focus();
    input.select();
    let done = false;
    const finish = save => {
      if (done) return;
      done = true;
      if (save && input.value.trim() && input.value.trim() !== sheet.name) app.commands.renameSheet(sheet, input.value.trim());
      this.render();
      app.focusGrid();
    };
    input.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); finish(true); } else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
    input.addEventListener('pointerdown', e => e.stopPropagation());
    input.addEventListener('click', e => e.stopPropagation());
  }

  /** Dragging a tab moves the sheet. */
  drag(e, sheet, tab) {
    const app = this.app;
    if (e.button !== 0 || app.readOnly || e.target.tagName === 'INPUT') return;
    const startX = e.clientX;
    let moving = false;
    let marker = null;
    let to = -1;
    const move = ev => {
      if (!moving && Math.abs(ev.clientX - startX) < 6) return;
      moving = true;
      tab.classList.add('is-dragging');
      const tabs = [...this.list.children];
      to = tabs.length;
      for (let i = 0; i < tabs.length; i++) {
        const r = tabs[i].getBoundingClientRect();
        if (ev.clientX < r.left + r.width / 2) { to = i; break; }
      }
      if (!marker) { marker = el('span', { class: 'mo-sheets-tabs__marker' }); this.root.append(marker); }
      const ref = tabs[Math.min(to, tabs.length - 1)].getBoundingClientRect();
      const host = this.root.getBoundingClientRect();
      marker.style.left = Math.round((to < tabs.length ? ref.left : ref.right) - host.left - 1) + 'px';
    };
    const up = () => {
      document.removeEventListener('pointermove', move);
      document.removeEventListener('pointerup', up);
      marker?.remove();
      tab.classList.remove('is-dragging');
      if (!moving || to < 0) return;
      const visible = app.wb.sheets.filter(s => !s.hidden);
      const target = visible[to];
      const from = app.wb.sheets.indexOf(sheet);
      let index = target ? app.wb.sheets.indexOf(target) : app.wb.sheets.length;
      if (index > from) index--;
      if (index !== from) app.commands.moveSheet(sheet, index);
    };
    document.addEventListener('pointermove', move);
    document.addEventListener('pointerup', up);
  }

  menu(sheet) {
    const app = this.app;
    const cmd = app.commands;
    const hidden = app.wb.sheets.filter(s => s.hidden);
    const index = app.wb.sheets.indexOf(sheet);
    return [
      { label: t('Sheets.Sheet.Insert'), icon: 'plus', run: () => cmd.addSheet() },
      { label: t('Sheets.Sheet.Delete'), icon: 'trash', run: () => cmd.deleteSheet(sheet) },
      { label: t('Sheets.Sheet.Rename'), icon: 'edit', run: () => { const tab = this.list.querySelector('.is-active'); if (tab) this.rename(sheet, tab); } },
      { label: t('Sheets.Sheet.Duplicate'), icon: 'copy', run: () => cmd.duplicateSheet(sheet) },
      { label: t('Sheets.Sheet.MoveLeft'), icon: 'chevron-left', disabled: () => index === 0, run: () => cmd.moveSheet(sheet, index - 1) },
      { label: t('Sheets.Sheet.MoveRight'), icon: 'chevron-right', disabled: () => index === app.wb.sheets.length - 1, run: () => cmd.moveSheet(sheet, index + 1) },
      '-',
      { label: t('Sheets.Sheet.Color'), icon: 'palette', run: () => this.colorMenu(sheet) },
      { label: t('Sheets.Sheet.Hide'), icon: 'eye-off', run: () => cmd.hideSheet(sheet) },
      { label: t('Sheets.Sheet.Unhide'), icon: 'eye', disabled: () => !hidden.length, run: () => app.dialogs.unhideSheet() },
      '-',
      { label: sheet.protected ? t('Sheets.Unprotect') : t('Sheets.Protect'), icon: 'lock', run: () => { if (sheet !== app.sheet) app.showSheet(sheet); cmd.protect(); } },
    ];
  }

  colorMenu(sheet) {
    const tab = this.list.querySelector('.is-active') || this.root;
    openPopover(tab, colourPalette(t('Sheets.NoColor'), c => this.app.commands.tabColor(sheet, c)));
  }
}
