/* MatOffice Sheets – the mouse and the finger on the grid: select cells, ranges, rows and columns (Shift, Ctrl), drag the fill handle,
   move a selection by its border (Ctrl copies), resize rows and columns at the headers (double click fits them), filter buttons, notes on
   hover, the context menu, Ctrl+wheel zoom, and pointing at cells while a formula is typed. Touch: a tap selects, the grid scrolls natively,
   a second tap on the active cell (or a double tap) edits it. */
import { el } from '../editor/util.js';
import { openPopover } from '../editor/frame.js';
import { MAX_ROWS, MAX_COLS } from './engine/address.js';
import { sizeLabel } from './editor.js';

export function attachPointer(app) {
  const grid = app.grid;
  const sc = grid.scroller;
  let drag = null;
  let lastTap = { time: 0, row: -1, col: -1 };
  let hoverTimer = 0;
  let noteBox = null;

  const local = e => {
    const r = grid.root.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };
  const onScrollbar = (e, p) => p.x >= sc.clientWidth || p.y >= sc.clientHeight;

  sc.addEventListener('pointerdown', e => {
    const p = local(e);
    if (onScrollbar(e, p)) return;
    hideNote();
    if (e.pointerType === 'touch' || e.pointerType === 'pen' && !e.isPrimary) {
      drag = { touch: true, x: p.x, y: p.y, id: e.pointerId };
      return;
    }
    if (e.button === 2) return; // the context menu decides
    if (e.button !== 0) return;
    const hit = grid.hitTest(p.x, p.y);
    const editor = app.editor;
    // pointing at cells while a formula is typed
    if (editor.active && hit.zone === 'cell' && (editor.canPoint() || editor.pointing)) {
      e.preventDefault();
      editor.pointStart(app.sheet, hit.row, hit.col);
      drag = { kind: 'point', id: e.pointerId };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (editor.active && !editor.commit(0, 0)) { e.preventDefault(); return; }
    app.objects?.deselect();
    e.preventDefault();
    app.focusGrid();
    const sel = app.sel;
    if (hit.zone === 'corner') { sel.all(); app.selectionChanged(); return; }
    if (hit.zone === 'col' || hit.zone === 'row') {
      if (hit.resize) { startResize(e, hit.resize, p); return; }
      const isCol = hit.zone === 'col';
      const index = isCol ? hit.col : hit.row;
      if (e.shiftKey) {
        const from = isCol ? sel.anchor.col : sel.anchor.row;
        if (isCol) sel.ranges[sel.ranges.length - 1] = { r1: 0, c1: Math.min(from, index), r2: MAX_ROWS - 1, c2: Math.max(from, index) };
        else sel.ranges[sel.ranges.length - 1] = { r1: Math.min(from, index), c1: 0, r2: Math.max(from, index), c2: MAX_COLS - 1 };
      } else if (isCol) sel.cols(index, index, e.ctrlKey || e.metaKey, grid.sheet.rows.firstVisible(0));
      else sel.rows(index, index, e.ctrlKey || e.metaKey, grid.sheet.cols.firstVisible(0));
      drag = { kind: isCol ? 'cols' : 'rows', from: isCol ? sel.anchor.col : sel.anchor.row, id: e.pointerId };
      sc.setPointerCapture(e.pointerId);
      app.selectionChanged({ scroll: false });
      return;
    }
    if (hit.filter !== null) { app.filters.open(hit.filter, p); return; }
    if (hit.fill && !app.readOnly) {
      drag = { kind: 'fill', source: { ...sel.range }, id: e.pointerId, target: { ...sel.range } };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    if (hit.border && !app.readOnly && sel.ranges.length === 1) {
      drag = { kind: 'move', source: { ...sel.range }, start: { row: hit.row, col: hit.col }, id: e.pointerId, dr: 0, dc: 0 };
      sc.setPointerCapture(e.pointerId);
      return;
    }
    const link = app.sheet.cell(hit.row, hit.col)?.link;
    if (link && (e.ctrlKey || e.metaKey)) { app.commands.followLink(link); return; }
    if (e.shiftKey) sel.extendTo(hit.row, hit.col);
    else if (e.ctrlKey || e.metaKey) sel.add(hit.row, hit.col);
    else sel.set(hit.row, hit.col);
    drag = { kind: 'cells', id: e.pointerId };
    sc.setPointerCapture(e.pointerId);
    app.selectionChanged({ scroll: false });
  });

  sc.addEventListener('pointermove', e => {
    const p = local(e);
    if (!drag) { hover(e, p); return; }
    if (drag.touch) return;
    drag.last = p;
    dragTo(p);
    autoScroll();
  });

  const finish = e => {
    clearInterval(drag?.scrollTimer);
    const d = drag;
    drag = null;
    if (!d) return;
    try { sc.releasePointerCapture(e.pointerId); } catch { /* not captured */ }
    if (d.touch) {
      const p = local(e);
      if (e.type === 'pointerup' && Math.hypot(p.x - d.x, p.y - d.y) < 10) tap(p);
      return;
    }
    if (d.kind === 'resize') endResize(d);
    else if (d.kind === 'fill') { grid.fillPreview = null; if (!sameArea(d.target, d.source)) app.commands.fill(d.source, d.target); grid.requestRender(); }
    else if (d.kind === 'move') {
      grid.movePreview = null;
      if (d.dr || d.dc) app.commands.moveSelection(d.dr, d.dc, e.ctrlKey || e.metaKey);
      else { app.sel.set(d.start.row, d.start.col); app.selectionChanged({ scroll: false }); } // a click on the border selects the cell
      grid.requestRender();
    }
    else if (d.kind === 'point') { app.editor.input.focus(); }
    else if (app.painter && (d.kind === 'cells' || d.kind === 'cols' || d.kind === 'rows')) app.commands.applyPainter();
    app.editor.refresh();
    app.updateStatus();
  };
  sc.addEventListener('pointerup', finish);
  sc.addEventListener('pointercancel', finish);

  function dragTo(p) {
    const hit = grid.hitTest(Math.max(0, Math.min(grid.width - 1, p.x)), Math.max(0, Math.min(grid.height - 1, p.y)));
    const sel = app.sel;
    const sheet = app.sheet;
    switch (drag.kind) {
      case 'cells':
        if (hit.zone === 'cell' || hit.zone === 'col' || hit.zone === 'row') {
          sel.extendTo(hit.zone === 'col' ? sheet.rows.firstVisible(0) : hit.row, hit.zone === 'row' ? 0 : hit.col);
          app.selectionChanged({ scroll: false, quick: true });
          app.editor.nameBox.value = sizeLabel(sel.range);
        }
        break;
      case 'cols': {
        const c = hit.col;
        sel.ranges[sel.ranges.length - 1] = { r1: 0, c1: Math.min(drag.from, c), r2: MAX_ROWS - 1, c2: Math.max(drag.from, c) };
        app.selectionChanged({ scroll: false, quick: true });
        break;
      }
      case 'rows': {
        const r = hit.row;
        sel.ranges[sel.ranges.length - 1] = { r1: Math.min(drag.from, r), c1: 0, r2: Math.max(drag.from, r), c2: MAX_COLS - 1 };
        app.selectionChanged({ scroll: false, quick: true });
        break;
      }
      case 'point':
        if (hit.zone === 'cell') app.editor.pointExtend(sheet, hit.row, hit.col);
        break;
      case 'fill': {
        const s = drag.source;
        const r = hit.row;
        const c = hit.col;
        // the direction that leaves the source the most
        const down = r > s.r2 ? r - s.r2 : 0;
        const up = r < s.r1 ? s.r1 - r : 0;
        const right = c > s.c2 ? c - s.c2 : 0;
        const left = c < s.c1 ? s.c1 - c : 0;
        let target = { ...s };
        if (Math.max(down, up) >= Math.max(right, left)) {
          if (down) target.r2 = r; else if (up) target.r1 = r;
        } else if (right) target.c2 = c; else if (left) target.c1 = c;
        drag.target = target;
        grid.fillPreview = { sheet: sheet.id, area: target };
        grid.requestRender();
        break;
      }
      case 'move': {
        const s = drag.source;
        const dr = Math.max(-s.r1, Math.min(MAX_ROWS - 1 - s.r2, hit.row - drag.start.row));
        const dc = Math.max(-s.c1, Math.min(MAX_COLS - 1 - s.c2, hit.col - drag.start.col));
        drag.dr = dr;
        drag.dc = dc;
        grid.movePreview = { sheet: sheet.id, area: { r1: s.r1 + dr, c1: s.c1 + dc, r2: s.r2 + dr, c2: s.c2 + dc } };
        grid.requestRender();
        break;
      }
      case 'resize': {
        const z = sheet.zoom;
        const delta = (drag.axis === 'col' ? p.x - drag.x : p.y - drag.y) / z;
        const size = Math.max(0, Math.round(drag.size + delta));
        (drag.axis === 'col' ? sheet.cols : sheet.rows).setSize(drag.index, size);
        drag.changed = true;
        app.grid.updateExtent();
        grid.requestRender();
        app.showTip?.(p, drag.axis === 'col' ? app.widthLabel(size) : app.heightLabel(size));
        break;
      }
      default: break;
    }
  }

  function autoScroll() {
    if (!drag || drag.kind === 'resize') return;
    clearInterval(drag.scrollTimer);
    const step = () => {
      if (!drag?.last) return;
      const p = drag.last;
      let dx = 0;
      let dy = 0;
      if (p.y > grid.height - 8) dy = 24; else if (p.y < 8 + (app.sheet.headings ? 20 : 0) && drag.kind !== 'cols') dy = -24;
      if (p.x > grid.width - 8) dx = 24; else if (p.x < 8 + (app.sheet.headings ? 30 : 0) && drag.kind !== 'rows') dx = -24;
      if (!dx && !dy) { clearInterval(drag.scrollTimer); return; }
      sc.scrollTop += dy;
      sc.scrollLeft += dx;
      dragTo(p);
    };
    drag.scrollTimer = setInterval(step, 40);
  }

  function startResize(e, resize, p) {
    if (app.readOnly) return;
    const sheet = app.sheet;
    const axis = resize.axis === 'col' ? sheet.cols : sheet.rows;
    app.wb.begin('resize');
    app.wb.touchProp(sheet, resize.axis === 'col' ? 'cols' : 'rows');
    drag = { kind: 'resize', axis: resize.axis, index: resize.index, size: axis.rawSize(resize.index), x: p.x, y: p.y, id: e.pointerId };
    sc.setPointerCapture(e.pointerId);
  }

  function endResize(d) {
    const sheet = app.sheet;
    const axis = d.axis === 'col' ? sheet.cols : sheet.rows;
    app.hideTip?.();
    if (d.changed) {
      const size = axis.rawSize(d.index);
      // every selected whole column (row) gets the same size
      const sel = app.sel;
      for (const a of sel.ranges) {
        const whole = d.axis === 'col' ? a.r1 === 0 && a.r2 === MAX_ROWS - 1 : a.c1 === 0 && a.c2 === MAX_COLS - 1;
        const lo = d.axis === 'col' ? a.c1 : a.r1;
        const hi = d.axis === 'col' ? a.c2 : a.r2;
        if (whole && d.index >= lo && d.index <= hi && hi - lo < 5000) for (let i = lo; i <= hi; i++) axis.setSize(i, size);
      }
      if (size === 0) axis.setHidden(d.index, true);
      if (d.axis === 'row') app.manualRows?.add(d.index);
    }
    app.wb.commit();
    app.changed({ layout: true });
  }

  sc.addEventListener('dblclick', e => {
    const p = local(e);
    if (onScrollbar(e, p)) return;
    const hit = grid.hitTest(p.x, p.y);
    if ((hit.zone === 'col' || hit.zone === 'row') && hit.resize) { app.commands.autofit(hit.resize.axis, hit.resize.index); return; }
    if (hit.zone === 'cell' && hit.fill) { app.commands.fillDownToData(); return; }
    if (hit.zone === 'cell' && !app.editor.active) {
      if (hit.border) { app.commands.jumpFromBorder(p); return; }
      app.sel.set(hit.row, hit.col);
      app.selectionChanged({ scroll: false });
      app.editor.begin({ mode: 'edit' });
    }
  });

  sc.addEventListener('contextmenu', e => {
    const p = local(e);
    if (onScrollbar(e, p)) return;
    e.preventDefault();
    if (app.editor.active && !app.editor.commit(0, 0)) return;
    const hit = grid.hitTest(p.x, p.y);
    const sel = app.sel;
    let kind = 'cells';
    if (hit.zone === 'col') {
      kind = 'cols';
      if (!sel.ranges.some(a => a.r1 === 0 && a.r2 === MAX_ROWS - 1 && hit.col >= a.c1 && hit.col <= a.c2)) sel.cols(hit.col, hit.col);
    } else if (hit.zone === 'row') {
      kind = 'rows';
      if (!sel.ranges.some(a => a.c1 === 0 && a.c2 === MAX_COLS - 1 && hit.row >= a.r1 && hit.row <= a.r2)) sel.rows(hit.row, hit.row);
    } else if (hit.zone === 'cell') {
      if (!sel.contains(hit.row, hit.col)) sel.set(hit.row, hit.col);
    } else {
      sel.all();
    }
    app.selectionChanged({ scroll: false });
    contextMenuAt(app, e.clientX, e.clientY, app.commands.contextMenu(kind));
  });

  sc.addEventListener('wheel', e => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    app.commands.zoomBy(e.deltaY < 0 ? 0.1 : -0.1);
  }, { passive: false });

  sc.addEventListener('pointerleave', () => { clearTimeout(hoverTimer); scheduleHide(); });

  function hover(e, p) {
    if (e.pointerType === 'touch') return;
    const hit = grid.hitTest(p.x, p.y);
    let cursor = 'cell';
    if (hit.zone === 'col') cursor = hit.resize ? 'col-resize' : 's-resize';
    else if (hit.zone === 'row') cursor = hit.resize ? 'row-resize' : 'e-resize';
    else if (hit.zone === 'corner') cursor = 'default';
    else if (hit.fill) cursor = 'crosshair';
    else if (hit.border) cursor = 'move';
    else if (hit.filter !== null) cursor = 'pointer';
    else if (app.editor.active && app.editor.canPoint()) cursor = 'cell';
    if (sc.style.cursor !== cursor) sc.style.cursor = cursor;
    // a note shows while the pointer rests on its cell
    clearTimeout(hoverTimer);
    if (hit.zone !== 'cell') { scheduleHide(); return; }
    const note = app.sheet.cell(hit.row, hit.col)?.note;
    if (!note) { scheduleHide(); return; }
    if (noteBox?.dataset.cell === hit.row + ':' + hit.col) return;
    hoverTimer = setTimeout(() => showNote(hit.row, hit.col, note), 250);
  }

  function showNote(row, col, note) {
    hideNote();
    const r = grid.cellRect(row, col);
    noteBox = el('div', { class: 'mo-sheets-note', role: 'note', 'data-cell': row + ':' + col });
    noteBox.textContent = note;
    noteBox.style.left = Math.round(r.x + r.w + 8) + 'px';
    noteBox.style.top = Math.round(r.y) + 'px';
    grid.overlay.append(noteBox);
    const box = noteBox.getBoundingClientRect();
    const host = grid.root.getBoundingClientRect();
    if (box.right > host.right - 4) noteBox.style.left = Math.max(4, Math.round(r.x - box.width - 8)) + 'px';
  }

  let hideTimer = 0;
  function scheduleHide() { clearTimeout(hideTimer); hideTimer = setTimeout(hideNote, 150); }
  function hideNote() { clearTimeout(hideTimer); noteBox?.remove(); noteBox = null; }
  app.hideNote = hideNote;

  function tap(p) {
    const hit = grid.hitTest(p.x, p.y);
    const sel = app.sel;
    if (app.editor.active) {
      if (hit.zone === 'cell' && app.editor.canPoint()) { app.editor.pointStart(app.sheet, hit.row, hit.col); return; }
      if (!app.editor.commit(0, 0)) return;
    }
    if (hit.zone === 'corner') { sel.all(); app.selectionChanged({ scroll: false }); return; }
    if (hit.zone === 'col') { sel.cols(hit.col, hit.col); app.selectionChanged({ scroll: false }); return; }
    if (hit.zone === 'row') { sel.rows(hit.row, hit.row); app.selectionChanged({ scroll: false }); return; }
    if (hit.filter !== null) { app.filters.open(hit.filter, p); return; }
    const now = Date.now();
    const again = sel.active.row === hit.row && sel.active.col === hit.col && sel.isSingle;
    const double = now - lastTap.time < 400 && lastTap.row === hit.row && lastTap.col === hit.col;
    lastTap = { time: now, row: hit.row, col: hit.col };
    if ((again || double) && !app.readOnly) { sel.set(hit.row, hit.col); app.selectionChanged({ scroll: false }); app.editor.begin({ mode: 'edit' }); return; }
    sel.set(hit.row, hit.col);
    app.selectionChanged({ scroll: false });
    if (app.painter) app.commands.applyPainter();
  }
}

/** Opens a menu at a point of the screen (a context menu). */
export function contextMenuAt(app, x, y, items) {
  const anchor = el('div', { class: 'mo-sheets-anchor' });
  anchor.style.left = x + 'px';
  anchor.style.top = y + 'px';
  document.body.append(anchor);
  openPopover(anchor, close => app.ctx.ribbon.menuItems ? menuOf(app, items, () => { close(); }) : menuOf(app, items, close));
  setTimeout(() => anchor.remove(), 0);
}

/** A menu like the ribbon's (items: { label, icon, key, run, disabled, checked, danger, edit } | '-'). */
export function menuOf(app, items, close) {
  const ribbon = app.ctx.ribbon;
  return ribbon.menuItems(items, close);
}

const sameArea = (a, b) => a.r1 === b.r1 && a.r2 === b.r2 && a.c1 === b.c1 && a.c2 === b.c2;
