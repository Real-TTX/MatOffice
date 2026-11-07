/* MatOffice Sheets – the keyboard on the grid. A hidden text field near the active cell has the focus: what is typed into it starts editing
   the cell (also with an input method), the clipboard events of the browser arrive there, and the keys move and extend the selection like
   in Excel (arrows, Ctrl+arrows to the edges of the data, Home, Ctrl+Home/End, PageUp/PageDown, Tab/Enter inside a selection …). */
import { el, t } from '../editor/util.js';
import { MAX_ROWS, MAX_COLS } from './engine/address.js';

export function attachKeys(app) {
  const catcher = el('textarea', {
    class: 'mo-sheets-keys', 'aria-label': t('Sheets.Grid'), autocomplete: 'off', autocapitalize: 'off', autocorrect: 'off', spellcheck: 'false', tabindex: '0',
  });
  app.grid.overlay.append(catcher);
  app.keys = catcher;
  let composing = false;

  catcher.addEventListener('compositionstart', () => { composing = true; });
  catcher.addEventListener('compositionend', () => {
    composing = false;
    const text = catcher.value;
    catcher.value = '';
    if (text) startTyping(text);
  });
  catcher.addEventListener('input', () => {
    if (composing) return;
    const text = catcher.value;
    catcher.value = '';
    if (text && !text.includes('\n')) startTyping(text);
  });

  function startTyping(text) {
    if (app.readOnly) { app.toastReadOnly(); return; }
    app.editor.begin({ text, mode: 'enter' });
  }

  catcher.addEventListener('keydown', e => {
    if (composing || e.isComposing) return;
    if (app.editor.active) return;
    const handled = onKey(e);
    if (handled) { e.preventDefault(); e.stopPropagation(); }
  });

  catcher.addEventListener('copy', e => app.clipboard.onCopy(e, false));
  catcher.addEventListener('cut', e => app.clipboard.onCopy(e, true));
  catcher.addEventListener('paste', e => app.clipboard.onPaste(e));

  function onKey(e) {
    const mod = e.ctrlKey || e.metaKey;
    const sel = app.sel;
    const sheet = app.sheet;
    const cmd = app.commands;
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[e.key] && !e.altKey) {
      const [dr, dc] = arrows[e.key];
      if (app.objects?.selected) { app.objects.nudge(dr, dc, e.shiftKey); return true; }
      move(dr, dc, e.shiftKey, mod);
      return true;
    }
    if (e.key === 'ArrowDown' && e.altKey) { cmd.openDropdown(); return true; }
    switch (e.key) {
      case 'Tab':
        if (mod) return false;
        if (!sel.cycle(0, e.shiftKey ? -1 : 1, r => !sheet.rows.isHidden(r), c => !sheet.cols.isHidden(c))) move(0, e.shiftKey ? -1 : 1, false, false);
        app.selectionChanged();
        return true;
      case 'Enter':
        if (mod || e.altKey) return false;
        if (!sel.cycle(e.shiftKey ? -1 : 1, 0, r => !sheet.rows.isHidden(r), c => !sheet.cols.isHidden(c))) move(e.shiftKey ? -1 : 1, 0, false, false);
        app.selectionChanged();
        return true;
      case 'Home': {
        if (mod) {
          const r = sheet.rows.firstVisible(sheet.freeze.rows);
          const c = sheet.cols.firstVisible(sheet.freeze.cols);
          if (e.shiftKey) sel.extendTo(r, c); else sel.set(r, c);
        } else {
          const c = sheet.cols.firstVisible(0);
          if (e.shiftKey) sel.extendTo(sel.active.row, c); else sel.set(sel.active.row, c);
        }
        app.selectionChanged();
        return true;
      }
      case 'End':
        if (!mod) return false;
        {
          const b = app.usedArea(sheet);
          if (e.shiftKey) sel.extendTo(b.r2, b.c2); else sel.set(b.r2, b.c2);
          app.selectionChanged();
        }
        return true;
      case 'PageDown': case 'PageUp': {
        if (mod) { cmd.nextSheet(e.key === 'PageDown' ? 1 : -1); return true; }
        const n = app.grid.pageRows();
        const d = e.key === 'PageDown' ? 1 : -1;
        if (e.altKey) {
          const cols = Math.max(1, Math.floor(app.grid.width / (sheet.cols.def * sheet.zoom)) - 1);
          const c = clampCol(sel.active.col + d * cols);
          if (e.shiftKey) sel.extendTo(sel.active.row, c); else sel.set(sel.active.row, c);
          app.grid.scroller.scrollLeft += d * cols * sheet.cols.def * sheet.zoom;
        } else {
          let r = sel.active.row;
          for (let i = 0; i < n; i++) r = sheet.rows.step(r, d);
          if (e.shiftKey) sel.extendTo(r, sel.active.col); else sel.set(r, sel.active.col);
          app.grid.scroller.scrollTop += d * (sheet.rows.offset(Math.max(r, sel.active.row)) - sheet.rows.offset(Math.min(r, sel.active.row))) * sheet.zoom || d * n * sheet.rows.def * sheet.zoom;
        }
        app.selectionChanged();
        return true;
      }
      case 'F2':
        if (e.shiftKey) { cmd.editNote(); return true; }
        if (app.readOnly) return true;
        app.editor.begin({ mode: 'edit' });
        return true;
      case 'Delete':
        if (app.objects?.selected) { app.objects.removeSelected(); return true; }
        cmd.clear('contents');
        return true;
      case 'Backspace':
        if (app.objects?.selected) { app.objects.removeSelected(); return true; }
        if (app.readOnly) return true;
        app.editor.begin({ text: '', mode: 'enter' });
        return true;
      case 'Escape':
        if (app.grid.marquee) { app.clipboard.clearMarquee(); return true; }
        if (app.painter) { cmd.stopPainter(); return true; }
        if (app.objects?.selected) { app.objects.deselect(); return true; }
        return false;
      case ' ':
        if (mod && e.shiftKey) { sel.all(); app.selectionChanged({ scroll: false }); return true; }
        if (mod) { sel.cols(sel.range.c1, sel.range.c2, false, sel.active.row); app.selectionChanged({ scroll: false }); return true; }
        if (e.shiftKey) { sel.rows(sel.range.r1, sel.range.r2, false, sel.active.col); app.selectionChanged({ scroll: false }); return true; }
        return false;
      case 'F5':
        cmd.goTo();
        return true;
      case 'F11':
        if (e.shiftKey) { cmd.addSheet(); return true; }
        return false;
      default: break;
    }
    if (!mod) return false;
    const k = e.key.length === 1 ? e.key.toLowerCase() : e.key;
    switch (k) {
      case 'z': if (e.shiftKey) app.redo(); else app.undo(); return true;
      case 'y': app.redo(); return true;
      case 'a': cmd.selectAll(); return true;
      case 'd': cmd.fillDirection('down'); return true;
      case 'r': cmd.fillDirection('right'); return true;
      case ';': cmd.insertNow(e.shiftKey ? 'time' : 'date'); return true;
      case ':': cmd.insertNow('time'); return true;
      case '-': cmd.deleteDialog(); return true;
      case '+': cmd.insertDialog(); return true;
      case '=': if (e.shiftKey) { cmd.insertDialog(); return true; } return false;
      case '9': cmd.hideLines('row', !e.shiftKey); return true;
      case '0': cmd.hideLines('col', !e.shiftKey); return true;
      case '(': cmd.hideLines('row', false); return true;
      case ')': cmd.hideLines('col', false); return true;
      case 'g': cmd.goTo(); return true;
      case '1': app.dialogs.formatCells(); return true;
      case 'f': app.dialogs.find(false); return true;
      case 'h': app.dialogs.find(true); return true;
      case 'k': if (!app.readOnly) app.dialogs.link(); return true;
      case 'l': if (e.shiftKey) { app.filters.toggle(); return true; } return false;
      case 'enter': return false;
      default: return false;
    }
  }

  const clampRow = r => Math.max(0, Math.min(MAX_ROWS - 1, r));
  const clampCol = c => Math.max(0, Math.min(MAX_COLS - 1, c));

  /** Moves (or with Shift extends) the selection by one cell or (Ctrl) to the edge of the data. */
  function move(dr, dc, extend, jump) {
    const sel = app.sel;
    const sheet = app.sheet;
    // extending moves the far corner of the range, not the active cell
    let from;
    if (extend) {
      const a = sel.range;
      from = { row: sel.anchor.row === a.r1 ? a.r2 : a.r1, col: sel.anchor.col === a.c1 ? a.c2 : a.c1 };
    } else {
      from = { ...sel.active };
      const m = sheet.mergeAt(from.row, from.col);
      if (m) from = { row: dr > 0 ? m.r2 : m.r1, col: dc > 0 ? m.c2 : m.c1 };
    }
    let next;
    if (jump) next = app.commands.edge(sheet, from.row, from.col, dr, dc);
    else next = { row: clampRow(dr ? sheet.rows.step(from.row, dr) : from.row), col: clampCol(dc ? sheet.cols.step(from.col, dc) : from.col) };
    if (extend) {
      sel.extendTo(next.row, next.col);
    } else {
      sel.set(next.row, next.col);
    }
    app.selectionChanged({ scrollTo: next });
  }
}
