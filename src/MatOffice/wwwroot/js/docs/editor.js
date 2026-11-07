/* MatOffice Docs – the editing engine. The DOM of the editing hosts (the body, a header or footer while it is edited, the notes) IS the document:
   typing inside a paragraph is left to the browser (IME, spelling, autocorrect stay native), everything that changes the structure – Enter,
   Backspace/Delete at the border of a paragraph, deleting a selection over several paragraphs, paste, drop, formatting, styles, lists – goes
   through the model (model.js): the affected blocks are read, changed and rendered again, and the selection is put back by character offsets.
   Undo/redo is our own: snapshots of the content (the app's snapshot()), cheap because unchanged blocks keep their serialized string. */
import {
  readBlock, readBlocks, renderBlock, offsetOf, pointAt, domLength, isTextBlockEl, parseHtml, isAtomic,
} from './htmlio.js';
import {
  isTextBlock, blockLength, splitBlock, mergeBlocks, insertRuns, deleteRuns, formatRuns, formatAt, sliceRuns, textRun, runsLength, styleOf, styleKey,
  withStyle, writeBlocks, runsText, cleanFormat, shellOf, normalizeRuns, cleanText,
} from './model.js';

const TEXT_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,blockquote,pre';
const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,blockquote,pre,table,hr,div[data-break]';
const HISTORY_LIMIT = 200;

export class DocEditor extends EventTarget {
  /**
   * options: { readOnly, snapshot(): state, restore(state), changed(host), renderOptions(): { fieldText, noteNumber }, onKey(e, host) → handled,
   *            onPasteFiles(files, at), clipboardText(blocks) }
   */
  constructor(options) {
    super();
    this.options = options;
    this.readOnly = !!options.readOnly;
    this.hosts = new Set();
    this.past = [];
    this.future = [];
    this.run = null;          // the typing run: { kind, at } – keystrokes of one run are one undo step
    this.runTimer = 0;
    this.pending = null;      // a format for the next typed text (Bold with a collapsed selection)
    this.composing = false;
    this.lastRange = null;    // the selection when the focus went to the ribbon
    this.selectionVersion = 0;
    this.restoring = false;
    document.addEventListener('selectionchange', () => this.onSelectionChange());
  }

  // ---- hosts ------------------------------------------------------------------------------------------------------------------

  /** Makes an element an editing host (kind: body, header, footer, note …). */
  addHost(host, kind) {
    host.dataset.moHost = kind;
    host.setAttribute('contenteditable', this.readOnly ? 'false' : 'true');
    host.setAttribute('role', 'textbox');
    host.setAttribute('aria-multiline', 'true');
    this.hosts.add(host);
    const on = (type, fn) => host.addEventListener(type, e => fn.call(this, e, host));
    on('beforeinput', this.onBeforeInput);
    on('input', this.onInput);
    on('keydown', this.onKeyDown);
    on('paste', this.onPaste);
    on('copy', this.onCopy);
    on('cut', this.onCut);
    on('drop', this.onDrop);
    on('dragstart', this.onDragStart);
    on('compositionstart', () => { this.composing = true; });
    on('compositionend', (e, h) => { this.composing = false; this.afterInput(h); });
    on('focus', () => { this.lastHost = host; });
    on('pointerdown', () => this.endRun());
    if (!this.readOnly) {
      const observer = new MutationObserver(records => this.onMutations(records, host));
      observer.observe(host, { subtree: true, childList: true, characterData: true, attributes: true, attributeOldValue: false });
      host._moObserver = observer;
    }
    return host;
  }

  removeHost(host) {
    host._moObserver?.disconnect();
    this.hosts.delete(host);
  }

  hostOf(node) {
    for (let n = node; n; n = n.parentNode) if (n.nodeType === 1 && this.hosts.has(n)) return n;
    return null;
  }

  setReadOnly(readOnly) {
    this.readOnly = readOnly;
    for (const h of this.hosts) h.setAttribute('contenteditable', readOnly ? 'false' : 'true');
  }

  /** The text blocks of a host in document order (also the ones in table cells). */
  textBlocks(host) {
    return [...host.querySelectorAll(TEXT_SELECTOR)].filter(b => this.hostOf(b) === host && !b.closest('[data-mo-deco]'));
  }

  // ---- changes ----------------------------------------------------------------------------------------------------------------

  onMutations(records, host) {
    if (this.silent) return;
    let content = false;
    this._state = null;
    for (const r of records) {
      if (r.type === 'attributes' && (r.attributeName.startsWith('data-mo-') || r.attributeName === 'contenteditable' || r.attributeName === 'class' || r.attributeName === 'spellcheck' || r.attributeName === 'lang')) continue;
      content = true;
      this.options.dirty?.(r.target, host);
    }
    if (content) this.options.changed?.(host);
  }

  /** Runs fn without telling anybody about the DOM changes it makes (rendering decorations). */
  quietly(fn) {
    this.silent = true;
    try { return fn(); } finally {
      // the observer delivers the records later: take them and drop them
      for (const h of this.hosts) h._moObserver?.takeRecords();
      this.silent = false;
    }
  }

  // ---- selection ----------------------------------------------------------------------------------------------------------------

  onSelectionChange() {
    const sel = document.getSelection();
    if (!sel || !sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    const host = this.hostOf(r.startContainer);
    if (!host || !host.contains(r.endContainer)) return;
    this.lastRange = r.cloneRange();
    this.lastHost = host;
    if (this.pending && !(r.collapsed && r.startContainer === this.pending.node && r.startOffset === this.pending.offset)) this.pending = null;
    if (this.run && !this.run.expect(r)) this.endRun();
    this.selectionVersion++;
    this.dispatchEvent(new CustomEvent('selection'));
  }

  /** Puts the focus and the last selection back into the document (after a click in the ribbon). */
  focus() {
    const sel = document.getSelection();
    const r = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    if (r && this.hostOf(r.startContainer)) return true;
    if (this.lastRange && this.lastHost?.isConnected && this.lastHost.contains(this.lastRange.startContainer)) {
      this.lastHost.focus({ preventScroll: true });
      sel.removeAllRanges();
      sel.addRange(this.lastRange.cloneRange());
      return true;
    }
    return false;
  }

  /**
   * The selection as text-block positions: { host, start: { block, offset }, end, collapsed, range }. Positions between blocks (on a table, a line,
   * the host itself) are moved into the nearest text block.
   */
  selection() {
    const sel = document.getSelection();
    let r = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    if (!r || !this.hostOf(r.startContainer)) r = this.lastRange && this.lastHost?.isConnected ? this.lastRange : null;
    if (!r) return null;
    const host = this.hostOf(r.startContainer);
    if (!host || !host.contains(r.endContainer)) return null;
    const start = this.position(host, r.startContainer, r.startOffset, false);
    const end = r.collapsed ? start : this.position(host, r.endContainer, r.endOffset, true);
    if (!start || !end) return null;
    return { host, start, end, collapsed: start.block === end.block && start.offset === end.offset, range: r };
  }

  position(host, node, offset, isEnd) {
    const block = closestTextBlock(node, host);
    if (block) return { block, offset: offsetOf(block, node, offset) };
    const blocks = this.textBlocks(host);
    if (!blocks.length) return null;
    const probe = document.createRange();
    try { probe.setStart(node, offset); } catch { return { block: blocks[0], offset: 0 }; }
    if (!isEnd) {
      for (const b of blocks) if (probe.comparePoint(b, 0) >= 0 || b.contains(node)) return { block: b, offset: 0 };
      const last = blocks[blocks.length - 1];
      return { block: last, offset: domLength(last) };
    }
    for (let i = blocks.length - 1; i >= 0; i--) {
      const b = blocks[i];
      if (probe.comparePoint(b, b.childNodes.length) <= 0) return { block: b, offset: domLength(b) };
    }
    return { block: blocks[0], offset: 0 };
  }

  /** Selects from (block a, offset) to (block b, offset); without b a caret. */
  select(a, aOffset, b = a, bOffset = aOffset, scroll = true) {
    if (!a || !a.isConnected) return;
    const sel = document.getSelection();
    const p1 = pointAt(a, aOffset);
    const p2 = b === a && bOffset === aOffset ? p1 : pointAt(b, bOffset);
    const r = document.createRange();
    try {
      r.setStart(p1.node, p1.offset);
      r.setEnd(p2.node, p2.offset);
    } catch { return; }
    const host = this.hostOf(a);
    if (host && document.activeElement !== host && !host.contains(document.activeElement)) host.focus({ preventScroll: true });
    sel.removeAllRanges();
    sel.addRange(r);
    this.lastRange = r.cloneRange();
    this.lastHost = host;
    this._state = null;
    if (scroll) scrollToCaret(r);
  }

  /** The selection as indexes (to put it back after the blocks were rendered again, or after undo). */
  bookmark(sel = this.selection()) {
    if (!sel) return null;
    const blocks = this.textBlocks(sel.host);
    return { host: sel.host.dataset.moHost, hostEl: sel.host, start: [blocks.indexOf(sel.start.block), sel.start.offset], end: [blocks.indexOf(sel.end.block), sel.end.offset] };
  }

  restoreBookmark(mark, scroll = true) {
    if (!mark) return;
    let host = mark.hostEl?.isConnected ? mark.hostEl : [...this.hosts].find(h => h.dataset.moHost === mark.host && h.isConnected);
    if (!host) return;
    const blocks = this.textBlocks(host);
    if (!blocks.length) return;
    const at = ([i, o]) => { const b = blocks[Math.min(Math.max(0, i), blocks.length - 1)]; return [b, Math.min(o, domLength(b))]; };
    const [a, ao] = at(mark.start);
    const [b, bo] = at(mark.end);
    this.select(a, ao, b, bo, scroll);
  }

  // ---- history --------------------------------------------------------------------------------------------------------------------

  /** Remembers the state before a change (one undo step). */
  checkpoint(kind = 'edit') {
    if (this.readOnly || this.restoring) return;
    if (kind !== 'type' && kind !== 'delete') this.endRun();
    const state = this.options.snapshot();
    const mark = this.bookmark();
    const last = this.past[this.past.length - 1];
    if (last && sameState(last.state, state)) { last.mark = last.mark || mark; this.future = []; return; }
    this.past.push({ state, mark });
    if (this.past.length > HISTORY_LIMIT) this.past.shift();
    this.future = [];
  }

  get canUndo() { return this.past.length > 0; }
  get canRedo() { return this.future.length > 0; }

  undo() {
    this.endRun();
    const current = { state: this.options.snapshot(), mark: this.bookmark() };
    while (this.past.length) {
      const step = this.past.pop();
      if (sameState(step.state, current.state)) continue;
      this.future.push(current);
      this.apply(step);
      return true;
    }
    return false;
  }

  redo() {
    this.endRun();
    const step = this.future.pop();
    if (!step) return false;
    this.past.push({ state: this.options.snapshot(), mark: this.bookmark() });
    this.apply(step);
    return true;
  }

  apply(step) {
    this.restoring = true;
    try {
      this.options.restore(step.state);
      this.restoreBookmark(step.mark);
    } finally { this.restoring = false; }
    this.options.changed?.(null);
    this.dispatchEvent(new CustomEvent('selection'));
  }

  clearHistory() { this.past = []; this.future = []; this.endRun(); }

  /** Typing keeps one undo step until the caret goes somewhere else, an other action comes or a pause of 2 s. */
  beginRun(kind) {
    clearTimeout(this.runTimer);
    this.runTimer = setTimeout(() => this.endRun(), 2000);
    if (this.run && this.run.kind === kind) return;
    this.endRun();
    this.checkpoint(kind);
    this.run = { kind, expect: () => true };
  }

  endRun() { this.run = null; clearTimeout(this.runTimer); }

  // ---- input --------------------------------------------------------------------------------------------------------------------

  onBeforeInput(e, host) {
    if (this.readOnly) { e.preventDefault(); return; }
    const type = e.inputType;
    if (type === 'insertCompositionText') { this.beginRun('type'); return; }
    const sel = this.selection();
    switch (type) {
      case 'insertText':
      case 'insertReplacementText': {
        const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
        if (type === 'insertReplacementText' && sel && sel.start.block === sel.end.block) { this.beginRun('type'); return; }
        if (!sel || this.needsManualTyping(sel, host)) { e.preventDefault(); this.typeText(text); return; }
        this.beginRun('type');
        return;
      }
      case 'insertParagraph': e.preventDefault(); this.enter(); return;
      case 'insertLineBreak': e.preventDefault(); this.lineBreak(); return;
      case 'deleteContentBackward': case 'deleteWordBackward': case 'deleteSoftLineBackward': case 'deleteHardLineBackward':
        this.deleteKey(e, sel, true); return;
      case 'deleteContentForward': case 'deleteWordForward': case 'deleteSoftLineForward': case 'deleteHardLineForward':
        this.deleteKey(e, sel, false); return;
      case 'deleteContent':
        e.preventDefault(); this.checkpoint(); this.deleteSelection(); return;
      case 'historyUndo': e.preventDefault(); this.undo(); return;
      case 'historyRedo': e.preventDefault(); this.redo(); return;
      case 'formatBold': e.preventDefault(); this.toggle('b'); return;
      case 'formatItalic': e.preventDefault(); this.toggle('i'); return;
      case 'formatUnderline': e.preventDefault(); this.toggle('u'); return;
      case 'insertFromPaste': case 'insertFromPasteAsQuotation': case 'insertFromDrop': case 'deleteByDrag': case 'deleteByCut': case 'insertFromYank':
        // handled in the paste / drop / cut events
        if (type === 'insertFromPaste' && !e.dataTransfer) return;
        e.preventDefault(); return;
      default:
        if (type.startsWith('format') || type.startsWith('insert')) e.preventDefault();
    }
  }

  /** Typing the browser would do wrong: over several paragraphs, outside a paragraph, at the border of a link, with a waiting format. */
  needsManualTyping(sel, host) {
    if (this.pending) return true;
    if (sel.start.block !== sel.end.block) return true;
    const r = sel.range;
    if (!closestTextBlock(r.startContainer, host)) return true;
    if (!sel.collapsed) {
      // a selection that holds objects other than text: the browser keeps odd wrappers
      return false;
    }
    // at the start or the end of a link or a comment the new text does not belong to it (like Word)
    const node = r.startContainer;
    const wrap = (node.nodeType === 1 ? node : node.parentNode).closest?.('a,span[data-comment]');
    if (wrap && wrap !== sel.start.block && sel.start.block.contains(wrap)) {
      const before = document.createRange();
      before.setStart(wrap, 0);
      before.setEnd(r.startContainer, r.startOffset);
      const textBefore = before.toString().length;
      if (textBefore === 0 || textBefore === wrap.textContent.length) return true;
    }
    // next to a non-editable object the browser may put the text into it
    if (node.nodeType === 1) {
      const prev = node.childNodes[r.startOffset - 1];
      if (prev && prev.nodeType === 1 && prev.getAttribute('contenteditable') === 'false') return true;
    }
    return false;
  }

  onInput(e, host) {
    if (this.composing) return;
    this.afterInput(host);
  }

  /** After the browser typed: an empty paragraph keeps its placeholder, stray wrappers the browser made inside a paragraph are cleaned. */
  afterInput(host) {
    const sel = document.getSelection();
    if (!sel.rangeCount) return;
    const block = closestTextBlock(sel.getRangeAt(0).startContainer, host);
    if (!block) {
      // the browser typed outside of a paragraph: wrap it
      this.repairHost(host);
      return;
    }
    if (domLength(block) === 0 && !block.querySelector('br')) block.append(document.createElement('br'));
    if (block.querySelector('div,p,h1,h2,h3,h4,h5,h6,blockquote,font')) this.rerender([block], true);
  }

  /** Puts loose text and odd elements of a host back into paragraphs (after the browser did something unexpected). */
  repairHost(host) {
    const mark = this.bookmark();
    const blocks = readBlocks(host);
    if (!blocks.length) blocks.push({ type: 'p', css: {}, runs: [] });
    const loose = [...host.childNodes].some(n => n.nodeType === 3 ? n.data.trim() : n.nodeType === 1 && !n.matches(BLOCK_SELECTOR) && !n.hasAttribute('data-mo-deco'));
    if (!loose) return;
    host.replaceChildren(...blocks.map(b => renderBlock(b, this.renderOptions())));
    this.ensureTail(host);
    this.restoreBookmark(mark, false);
  }

  renderOptions() { return this.options.renderOptions?.() || {}; }

  // ---- keys -----------------------------------------------------------------------------------------------------------------------

  onKeyDown(e, host) {
    if (this.composing || e.isComposing) return;
    if (this.options.onKey?.(e, host)) { e.preventDefault(); return; }
    if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown'].includes(e.key)) this.endRun();
    if (this.readOnly) return;
    if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      this.dispatchEvent(new CustomEvent('tab', { detail: { shift: e.shiftKey } }));
    }
  }

  // ---- typing -------------------------------------------------------------------------------------------------------------------

  /** Inserts text at the selection through the model (with the waiting format, outside of a link end …). */
  typeText(text) {
    text = cleanText(text);
    if (!text) return;
    this.beginRun('type');
    let sel = this.selection();
    if (!sel) return;
    if (!sel.collapsed) { this.deleteSelection(sel, false); sel = this.selection(); if (!sel) return; }
    const { block, offset } = sel.start;
    const b = readBlock(block);
    const f = this.pending ? { ...formatAt(b.runs, offset), ...this.pending.format } : formatAt(b.runs, offset);
    this.pending = null;
    const isPre = b.type === 'pre';
    const pieces = isPre ? [textRun(text, f)] : text.split('\n').flatMap((part, i) => (i ? [{ obj: 'br', f: cleanFormat(f) }] : []).concat(part ? [textRun(part, f)] : []));
    b.runs = insertRuns(b.runs, offset, pieces);
    this.morph(block, b);
    this.select(block, offset + runsLength(pieces));
  }

  /** A format for the next typed text (Bold, a font … with a collapsed selection). */
  setPending(format) {
    const sel = document.getSelection();
    if (!sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    this.pending = { format: { ...(this.pending?.format || {}), ...format }, node: r.startContainer, offset: r.startOffset };
    this._state = null;
    this.dispatchEvent(new CustomEvent('selection'));
  }

  // ---- structure: Enter, line breaks, deleting ------------------------------------------------------------------------------------

  enter() {
    this.checkpoint('enter');
    let sel = this.selection();
    if (!sel) return;
    if (!sel.collapsed) { this.deleteSelection(sel, false); sel = this.selection(); if (!sel) return; }
    const el = sel.start.block;
    const offset = sel.start.offset;
    const b = readBlock(el);
    const len = blockLength(b);
    if (b.type === 'pre') {
      const text = runsText(b.runs);
      if (offset === len && len > 0 && text.endsWith('\n')) {
        // Enter on the empty last line of code leaves the code
        b.runs = deleteRuns(b.runs, len - 1, len);
        this.morph(el, b);
        const next = renderBlock({ type: 'p', css: {}, runs: [] }, this.renderOptions());
        el.after(next);
        this.select(next, 0);
        return;
      }
      b.runs = insertRuns(b.runs, offset, [textRun('\n', formatAt(b.runs, offset))]);
      this.morph(el, b);
      this.select(el, offset + 1);
      return;
    }
    if (b.list && len === 0) {
      // an empty list item ends the list (or goes one level up)
      if (b.level > 0) b.level -= 1;
      else { delete b.list; delete b.level; delete b.start; }
      this.morph(el, b);
      this.select(el, 0);
      return;
    }
    let [first, second] = splitBlock(b, offset);
    const typing = formatAt(b.runs, offset);
    if (offset === 0 && len > 0) {
      // Enter at the start: an empty paragraph before, the text keeps its id and number
      first = shellOf(b);
      second = { ...b, css: { ...b.css } };
    } else if (offset === len) {
      const next = styleOf(styleKey(b)).next;
      if (next !== styleKey(b)) second = withStyle(second, next);
      if (/^h[1-6]$/.test(b.type)) { delete second.list; delete second.level; }
      delete second.start;
    }
    delete second.start;
    if (offset !== 0 || len === 0) delete second.id;
    const nodes = this.replace(el, [first, second]);
    const target = nodes[1];
    this.select(target, 0);
    if (offset === len && len > 0 && styleKey(second) === styleKey(b)) {
      const carry = { ...typing };
      delete carry.href; delete carry.cmt;
      if (Object.keys(carry).length) this.setPending(carry);
    }
  }

  lineBreak() {
    this.checkpoint('enter');
    let sel = this.selection();
    if (!sel) return;
    if (!sel.collapsed) { this.deleteSelection(sel, false); sel = this.selection(); if (!sel) return; }
    const { block, offset } = sel.start;
    const b = readBlock(block);
    const piece = b.type === 'pre' ? textRun('\n', formatAt(b.runs, offset)) : { obj: 'br', f: cleanFormat(formatAt(b.runs, offset)) };
    b.runs = insertRuns(b.runs, offset, [piece]);
    this.morph(block, b);
    this.select(block, offset + 1);
  }

  deleteKey(e, sel, backward) {
    if (!sel) { e.preventDefault(); return; }
    if (!sel.collapsed) {
      if (sel.start.block !== sel.end.block || this.selectionHasStructure(sel)) { e.preventDefault(); this.checkpoint(); this.deleteSelection(sel); return; }
      this.beginRun('delete');
      return;
    }
    const { block, offset } = sel.start;
    if (backward && offset === 0) { e.preventDefault(); this.checkpoint('join'); this.backspaceAtStart(block); return; }
    if (!backward && offset >= domLength(block)) { e.preventDefault(); this.checkpoint('join'); this.deleteAtEnd(block); return; }
    this.beginRun('delete');
  }

  selectionHasStructure(sel) {
    const r = sel.range;
    const frag = r.cloneContents();
    return !!frag.querySelector?.('table,hr,div[data-break],p,h1,h2,h3,h4,h5,h6,blockquote,pre');
  }

  backspaceAtStart(el) {
    const b = readBlock(el);
    if (b.list) {
      // like Word: the first Backspace takes the number or bullet away
      delete b.list; delete b.level; delete b.start;
      this.morph(el, b);
      this.select(el, 0);
      return;
    }
    const prev = previousBlock(el);
    if (!prev) {
      if (styleKey(b) !== 'Normal' && blockLength(b) === 0) { this.morph(el, withStyle(b, 'Normal')); this.select(el, 0); }
      return;
    }
    if (isTextBlockEl(prev)) {
      const pb = readBlock(prev);
      const at = blockLength(pb);
      if (at === 0 && !pb.list) { prev.remove(); this.select(el, 0); return; }
      this.morph(prev, mergeBlocks(pb, b));
      el.remove();
      this.select(prev, at);
      return;
    }
    if (prev.matches('hr,div[data-break]')) { prev.remove(); this.select(el, 0); return; }
    if (prev.matches('table')) {
      const cells = this.textBlocksIn(prev);
      const last = cells[cells.length - 1];
      if (blockLength(b) === 0 && el.nextElementSibling) el.remove();
      if (last) this.select(last, domLength(last));
    }
  }

  deleteAtEnd(el) {
    const next = nextBlock(el);
    if (!next) return;
    if (isTextBlockEl(next)) {
      const b = readBlock(el);
      const at = blockLength(b);
      if (at === 0 && !b.list) { el.remove(); this.select(next, 0); return; }
      this.morph(el, mergeBlocks(b, readBlock(next)));
      next.remove();
      this.select(el, at);
      return;
    }
    if (next.matches('hr,div[data-break]')) { next.remove(); this.select(el, domLength(el)); return; }
    if (next.matches('table')) {
      const first = this.textBlocksIn(next)[0];
      if (domLength(el) === 0 && el.previousElementSibling) el.remove();
      if (first) this.select(first, 0);
    }
  }

  textBlocksIn(node) { return [...node.querySelectorAll(TEXT_SELECTOR)]; }

  /** Deletes the selection (over several paragraphs, tables and cells): Word-like, cells of one table are emptied, not removed. */
  deleteSelection(sel = this.selection(), select = true) {
    if (!sel || sel.collapsed) return;
    const { start, end } = sel;
    const a = start.block, z = end.block;
    if (a === z) {
      const b = readBlock(a);
      b.runs = deleteRuns(b.runs, start.offset, end.offset);
      this.morph(a, b);
      if (select) this.select(a, start.offset);
      return;
    }
    const cellA = a.closest('td,th'), cellZ = z.closest('td,th');
    if (cellA && cellZ && cellA !== cellZ && cellA.closest('table') === cellZ.closest('table') && sel.host.contains(cellA)) {
      this.clearCells(sel, cellA, cellZ);
      return;
    }
    const container = commonContainer(a, z, sel.host);
    const topA = topIn(a, container), topZ = topIn(z, container);
    for (let n = topA.nextSibling; n && n !== topZ;) { const next = n.nextSibling; if (!(n.nodeType === 1 && n.hasAttribute('data-mo-deco'))) n.remove(); n = next; }
    const textA = topA === a, textZ = topZ === z;
    if (textA && textZ) {
      const ba = readBlock(a), bz = readBlock(z);
      ba.runs = normalizeRuns([...sliceRuns(ba.runs, 0, start.offset), ...sliceRuns(bz.runs, end.offset, blockLength(bz))]);
      this.morph(a, ba);
      z.remove();
      if (select) this.select(a, start.offset);
    } else if (textA) {
      const ba = readBlock(a);
      ba.runs = sliceRuns(ba.runs, 0, start.offset);
      this.morph(a, ba);
      topZ.remove();
      if (select) this.select(a, start.offset);
    } else if (textZ) {
      const bz = readBlock(z);
      bz.runs = sliceRuns(bz.runs, end.offset, blockLength(bz));
      this.morph(z, bz);
      topA.remove();
      if (select) this.select(z, 0);
    } else {
      const p = renderBlock({ type: 'p', css: {}, runs: [] }, this.renderOptions());
      topA.before(p);
      topA.remove();
      topZ.remove();
      if (select) this.select(p, 0);
    }
    this.ensureTail(container);
  }

  clearCells(sel, cellA, cellZ) {
    const table = cellA.closest('table');
    const cells = [...table.querySelectorAll('td,th')].filter(c => c.closest('table') === table);
    const i = cells.indexOf(cellA), j = cells.indexOf(cellZ);
    const [from, to] = i <= j ? [i, j] : [j, i];
    for (let k = from; k <= to; k++) {
      const cell = cells[k];
      if (cell === cellA && i <= j) {
        const blocks = this.textBlocksIn(cell);
        const first = sel.start.block;
        const b = readBlock(first);
        b.runs = sliceRuns(b.runs, 0, sel.start.offset);
        this.morph(first, b);
        for (const other of blocks) if (other !== first && other.compareDocumentPosition(first) & Node.DOCUMENT_POSITION_PRECEDING) other.remove();
        continue;
      }
      if (cell === cellZ && i <= j) {
        const last = sel.end.block;
        const b = readBlock(last);
        b.runs = sliceRuns(b.runs, sel.end.offset, blockLength(b));
        this.morph(last, b);
        for (const other of this.textBlocksIn(cell)) if (other !== last && other.compareDocumentPosition(last) & Node.DOCUMENT_POSITION_FOLLOWING) other.remove();
        continue;
      }
      cell.replaceChildren(renderBlock({ type: 'p', css: {}, runs: [] }, this.renderOptions()));
    }
    this.select(sel.start.block.isConnected ? sel.start.block : this.textBlocksIn(cellA)[0], sel.start.block.isConnected ? sel.start.offset : 0);
  }

  /** A container never ends with a table, a line or a break (there must be a paragraph to type in), and it is never empty. */
  ensureTail(container) {
    if (!container || !container.isConnected) return;
    const blocks = [...container.children].filter(n => n.matches(BLOCK_SELECTOR));
    const last = blocks[blocks.length - 1];
    if (!last || !isTextBlockEl(last)) {
      const p = renderBlock({ type: 'p', css: {}, runs: [] }, this.renderOptions());
      if (last) last.after(p); else container.prepend(p);
    }
  }

  // ---- replacing blocks -----------------------------------------------------------------------------------------------------------

  /** Shows a block again in the same element (attributes and children replaced; the element stays, references to it keep working). */
  morph(el, block) {
    this._state = null;
    if (!isTextBlock(block)) { return this.replace(el, [block])[0]; }
    const fresh = renderBlock(block, this.renderOptions());
    if (fresh.tagName !== el.tagName) { el.replaceWith(fresh); return fresh; }
    for (const a of [...el.attributes]) if (!fresh.hasAttribute(a.name) && !a.name.startsWith('data-mo-')) el.removeAttribute(a.name);
    for (const a of [...fresh.attributes]) if (el.getAttribute(a.name) !== a.value) el.setAttribute(a.name, a.value);
    el.replaceChildren(...fresh.childNodes);
    return el;
  }

  /** Replaces a block element by the rendering of blocks; returns the new elements. */
  replace(el, blocks) {
    this._state = null;
    const nodes = blocks.map(b => renderBlock(b, this.renderOptions()));
    el.replaceWith(...nodes);
    return nodes;
  }

  /** Renders text blocks again from their own DOM (after the browser left odd wrappers), keeping the selection. */
  rerender(blocks, keepSelection = true) {
    const mark = keepSelection ? this.bookmark() : null;
    for (const el of blocks) if (el.isConnected) this.morph(el, readBlock(el));
    if (mark) this.restoreBookmark(mark, false);
  }

  // ---- formatting primitives ---------------------------------------------------------------------------------------------------------

  /** The text blocks the selection touches with the range in each: [{ el, start, end }]. */
  ranges(sel = this.selection()) {
    if (!sel) return [];
    const blocks = this.textBlocks(sel.host);
    const i = blocks.indexOf(sel.start.block), j = blocks.indexOf(sel.end.block);
    if (i < 0 || j < 0) return [];
    const out = [];
    for (let k = i; k <= j; k++) {
      const el = blocks[k];
      out.push({ el, start: k === i ? sel.start.offset : 0, end: k === j ? sel.end.offset : domLength(el) });
    }
    return out;
  }

  /** Changes the format of the selected text: patch { b: true } or a function f => f. A caret: the word under it (like Word), else the next typed text. */
  formatText(patch, { word = true } = {}) {
    let sel = this.selection();
    if (!sel) return;
    if (sel.collapsed) {
      const range = word ? wordAt(sel.start.block, sel.start.offset) : null;
      if (!range) {
        const f = typeof patch === 'function' ? patch({ ...this.currentFormat() }) : { ...this.currentFormat(), ...patch };
        this.setPending(cleanFormat(f));
        return;
      }
      this.checkpoint('format');
      const b = readBlock(sel.start.block);
      b.runs = formatRuns(b.runs, range[0], range[1], patch);
      this.morph(sel.start.block, b);
      this.select(sel.start.block, sel.start.offset, sel.start.block, sel.start.offset, false);
      return;
    }
    this.checkpoint('format');
    const mark = this.bookmark(sel);
    for (const { el, start, end } of this.ranges(sel)) {
      if (end <= start) continue;
      const b = readBlock(el);
      b.runs = formatRuns(b.runs, start, end, patch);
      this.morph(el, b);
    }
    this.restoreBookmark(mark, false);
  }

  /** Toggles b/i/u/s/sub/sup over the selection. */
  toggle(key) {
    const on = this.state().format[key];
    this.formatText(f => {
      const g = { ...f, [key]: !on };
      if (key === 'sub' && !on) delete g.sup;
      if (key === 'sup' && !on) delete g.sub;
      return g;
    });
  }

  /** The format at the caret (with the waiting format). */
  currentFormat() {
    const sel = this.selection();
    if (!sel) return {};
    const b = readBlock(sel.start.block);
    return { ...formatAt(b.runs, sel.collapsed ? sel.start.offset : Math.min(sel.start.offset + 1, blockLength(b))), ...(this.pending?.format || {}) };
  }

  /** Changes the selected paragraphs: fn(block) returns the new block. */
  formatBlocks(fn, kind = 'paragraph') {
    const sel = this.selection();
    if (!sel) return;
    this.checkpoint(kind);
    const mark = this.bookmark(sel);
    const list = this.ranges(sel);
    list.forEach(({ el }, index) => {
      const b = fn(readBlock(el), index, list.length);
      if (b) this.morph(el, b);
    });
    this.restoreBookmark(mark, false);
  }

  /**
   * The state of the selection for the ribbon: format (common values; '' when mixed), the block of the start (style, alignment, list …),
   * the computed font and size at the caret.
   */
  state() {
    if (this._state && this._state.version === this.selectionVersion && this._state.pending === this.pending) return this._state;
    const sel = this.selection();
    const s = { version: this.selectionVersion, pending: this.pending, format: {}, block: null, sel };
    if (sel) {
      const list = this.ranges(sel);
      const keys = ['b', 'i', 'u', 's', 'sub', 'sup', 'href', 'cmt', 'font', 'size', 'color', 'bg'];
      const firstBlock = readBlock(sel.start.block);
      s.block = firstBlock;
      s.blockEl = sel.start.block;
      if (sel.collapsed) s.format = { ...formatAt(firstBlock.runs, sel.start.offset), ...(this.pending?.format || {}) };
      else {
        const values = {};
        for (const { el, start, end } of list) {
          if (end <= start) continue;
          const runs = el === sel.start.block ? firstBlock.runs : readBlock(el).runs;
          let pos = 0;
          for (const r of runs) {
            const len = r.obj ? 1 : r.text.length;
            if (pos < end && pos + len > start && r.obj !== 'br') {
              for (const k of keys) {
                const v = r.f[k] ?? null;
                if (!(k in values)) values[k] = v;
                else if (values[k] !== v) values[k] = '';
              }
            }
            pos += len;
          }
        }
        for (const k of keys) if (values[k]) s.format[k] = values[k];
        s.mixed = values;
      }
      // the font and size people see (the style's when the text has none)
      const style = styleOf(styleKey(firstBlock));
      s.styleKey = styleKey(firstBlock);
      s.fontName = s.format.font ? firstFamilyOf(s.format.font) : (sel.collapsed || s.mixed?.font === null ? style.font : '');
      s.sizePt = s.format.size ? parseFloat(s.format.size) : (sel.collapsed || s.mixed?.size === null ? style.size : null);
    }
    this._state = s;
    return s;
  }

  // ---- clipboard ----------------------------------------------------------------------------------------------------------------

  /** The selection as blocks of the model (for copy, drag, format painter). */
  selectedBlocks(sel = this.selection()) {
    if (!sel || sel.collapsed) return [];
    if (sel.start.block === sel.end.block) {
      const b = readBlock(sel.start.block);
      return [{ ...b, runs: sliceRuns(b.runs, sel.start.offset, sel.end.offset), id: undefined, start: undefined }];
    }
    const container = commonContainer(sel.start.block, sel.end.block, sel.host);
    const topA = topIn(sel.start.block, container), topZ = topIn(sel.end.block, container);
    const out = [];
    for (let n = topA; n; n = n.nextSibling) {
      if (n.nodeType !== 1 || !n.matches(BLOCK_SELECTOR)) { if (n === topZ) break; continue; }
      const blocks = readBlocks(wrapOne(n));
      for (let b of blocks) {
        if (n === topA && isTextBlock(b) && n === sel.start.block) b = { ...b, runs: sliceRuns(b.runs, sel.start.offset, blockLength(b)) };
        if (n === topZ && isTextBlock(b) && n === sel.end.block) b = { ...b, runs: sliceRuns(b.runs, 0, sel.end.offset) };
        out.push(b);
      }
      if (n === topZ) break;
    }
    return out;
  }

  onCopy(e) {
    const blocks = this.selectedBlocks();
    if (!blocks.length) return;
    e.preventDefault();
    e.clipboardData.setData('text/html', '<meta charset="utf-8"><div data-mo-docs="1">' + writeBlocks(blocks) + '</div>');
    e.clipboardData.setData('text/plain', this.options.clipboardText ? this.options.clipboardText(blocks) : blocks.map(b => runsText(b.runs || [], '')).join('\n'));
  }

  onCut(e) {
    if (this.readOnly) return;
    const sel = this.selection();
    if (!sel || sel.collapsed) return;
    this.onCopy(e);
    this.checkpoint('cut');
    this.deleteSelection(sel);
  }

  onPaste(e) {
    e.preventDefault();
    if (this.readOnly) return;
    const dt = e.clipboardData;
    const plain = this.plainNext;
    this.plainNext = false;
    const html = dt.getData('text/html');
    const text = dt.getData('text/plain');
    const files = [...(dt.files || [])].filter(f => /^image\/(png|jpeg|gif|webp|bmp)$/.test(f.type));
    if (!plain && html) {
      const blocks = parseHtml(html, { collapse: !html.includes('data-mo-docs="1"') });
      if (blocks.length && blocks.some(b => !isTextBlock(b) || b.runs.length)) { this.insertBlocks(blocks); return; }
    }
    if (!plain && files.length) { this.options.onPasteFiles?.(files); return; }
    if (text) this.insertPlainText(text);
  }

  /** Plain text: one paragraph per line, in the style of the paragraph it goes into. */
  insertPlainText(text) {
    const lines = cleanText(text).split('\n');
    const sel = this.selection();
    if (!sel) return;
    if (lines.length === 1) { this.checkpoint('paste'); this.typeText(lines[0]); return; }
    const base = readBlock(sel.start.block);
    const f = formatAt(base.runs, sel.start.offset);
    const shell = shellOf(base);
    delete shell.start;
    this.insertBlocks(lines.map(line => ({ ...shell, css: { ...shell.css }, runs: line ? [textRun(line.replace(/\r/g, ''), f)] : [] })));
  }

  /** Inserts blocks at the selection (paste, drop, a template part): one paragraph goes into the text, several split it. */
  insertBlocks(blocks) {
    if (!blocks.length) return;
    this.checkpoint('paste');
    let sel = this.selection();
    if (!sel) return;
    if (!sel.collapsed) { this.deleteSelection(sel, false); sel = this.selection(); if (!sel) return; }
    const el = sel.start.block;
    const offset = sel.start.offset;
    const inCell = !!el.closest('td,th');
    blocks = blocks.filter(b => !(inCell && b.type === 'break'));
    if (!blocks.length) return;
    if (blocks.length === 1 && isTextBlock(blocks[0])) {
      const b = readBlock(el);
      const runs = blocks[0].runs;
      b.runs = insertRuns(b.runs, offset, runs);
      if (blockLength(readBlock(el)) === 0 && styleKey(blocks[0]) !== 'Normal') {
        Object.assign(b, { type: blocks[0].type, css: { ...blocks[0].css } });
        if (blocks[0].style) b.style = blocks[0].style; else delete b.style;
      }
      this.morph(el, b);
      this.select(el, offset + runsLength(runs));
      return;
    }
    const b = readBlock(el);
    const [head, tail] = splitBlock(b, offset);
    const out = [];
    const first = blocks[0];
    if (blockLength(head) === 0) out.push(...blocks);
    else if (isTextBlock(first)) { out.push({ ...head, runs: normalizeRuns([...head.runs, ...first.runs]) }, ...blocks.slice(1)); }
    else out.push(head, ...blocks);
    let caretIndex = out.length - 1;
    let caretOffset = isTextBlock(out[caretIndex]) ? blockLength(out[caretIndex]) : 0;
    const last = out[out.length - 1];
    if (blockLength(tail) > 0 || !isTextBlock(last)) {
      if (isTextBlock(last) && blockLength(tail) > 0) {
        // an empty paragraph that came with the insertion (after a break or a line) takes the rest with its attributes
        out[out.length - 1] = blockLength(last) === 0 && last.type === 'p' && !last.style && !last.list ? tail : { ...last, runs: normalizeRuns([...last.runs, ...tail.runs]) };
      }
      else { out.push(tail); caretIndex = out.length - 1; caretOffset = 0; }
    }
    const nodes = this.replace(el, out);
    const target = nodes[caretIndex];
    if (target && isTextBlockEl(target)) this.select(target, caretOffset);
    else {
      const blocksIn = this.textBlocksIn(target || nodes[nodes.length - 1]);
      if (blocksIn.length) this.select(blocksIn[blocksIn.length - 1], domLength(blocksIn[blocksIn.length - 1]));
    }
  }

  onDragStart(e) {
    // moving text inside the document by dragging is not supported (it would bypass the model); pictures are moved with cut and paste
    e.preventDefault();
  }

  onDrop(e, host) {
    if (this.readOnly) return;
    e.preventDefault();
    const point = caretFromPoint(e.clientX, e.clientY);
    if (point && host.contains(point.node)) {
      const r = document.createRange();
      r.setStart(point.node, point.offset);
      const sel = document.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    }
    const dt = e.dataTransfer;
    const files = [...(dt.files || [])].filter(f => /^image\//.test(f.type));
    if (files.length) { this.options.onPasteFiles?.(files); return; }
    const html = dt.getData('text/html');
    if (html) { const blocks = parseHtml(html, { collapse: true }); if (blocks.length) { this.insertBlocks(blocks); return; } }
    const text = dt.getData('text/plain');
    if (text) this.insertPlainText(text);
  }
}

// ---- helpers ------------------------------------------------------------------------------------------------------------------------

export function closestTextBlock(node, host) {
  for (let n = node; n && n !== host; n = n.parentNode) {
    if (n.nodeType === 1 && isTextBlockEl(n)) return n;
    if (n.nodeType === 1 && n.getAttribute?.('contenteditable') === 'false' && !isAtomic(n)) return null;
  }
  return null;
}

/** The block before an element in its container (decorations skipped). */
export function previousBlock(el) {
  let n = el.previousElementSibling;
  while (n && (n.hasAttribute('data-mo-deco') || !n.matches(BLOCK_SELECTOR))) n = n.previousElementSibling;
  return n;
}

export function nextBlock(el) {
  let n = el.nextElementSibling;
  while (n && (n.hasAttribute('data-mo-deco') || !n.matches(BLOCK_SELECTOR))) n = n.nextElementSibling;
  return n;
}

/** The innermost container (host or cell) that holds both elements. */
function commonContainer(a, z, host) {
  for (let n = a.parentNode; n; n = n.parentNode) {
    if ((n === host || n.matches?.('td,th')) && n.contains(z)) return n;
    if (n === host) return host;
  }
  return host;
}

/** The child of a container an element lies in. */
function topIn(el, container) {
  let n = el;
  while (n.parentNode && n.parentNode !== container) n = n.parentNode;
  return n;
}

function wrapOne(n) {
  const d = document.createElement('div');
  d.append(n.cloneNode(true));
  return d;
}

/** The word at an offset of a block ([start, end]) or null when the caret is not inside a word. */
export function wordAt(block, offset) {
  const text = runsText(readBlock(block).runs, ' ');
  const isWord = c => !!c && /[\p{L}\p{N}_'’-]/u.test(c);
  if (!isWord(text[offset - 1]) || !isWord(text[offset])) return null;
  let s = offset, e = offset;
  while (s > 0 && isWord(text[s - 1])) s--;
  while (e < text.length && isWord(text[e])) e++;
  return s < e ? [s, e] : null;
}

function firstFamilyOf(stack) {
  const m = /^\s*(['"]?)([^,'"]+)\1/.exec(stack || '');
  return m ? m[2].trim() : stack;
}

function sameState(a, b) {
  if (a === b) return true;
  if (!a || !b) return false;
  const ka = Object.keys(a), kb = Object.keys(b);
  if (ka.length !== kb.length) return false;
  for (const k of ka) {
    const x = a[k], y = b[k];
    if (Array.isArray(x) && Array.isArray(y)) { if (x.length !== y.length || x.some((v, i) => v !== y[i])) return false; }
    else if (x !== y) return false;
  }
  return true;
}

export function caretFromPoint(x, y) {
  if (document.caretPositionFromPoint) {
    const p = document.caretPositionFromPoint(x, y);
    return p ? { node: p.offsetNode, offset: p.offset } : null;
  }
  if (document.caretRangeFromPoint) {
    const r = document.caretRangeFromPoint(x, y);
    return r ? { node: r.startContainer, offset: r.startOffset } : null;
  }
  return null;
}

/** Scrolls the nearest scrolling box so that the caret (or the start of the selection) can be seen. */
export function scrollToCaret(range) {
  let rect = range.getBoundingClientRect();
  if (!rect || (!rect.height && !rect.width)) {
    const node = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentNode;
    rect = node?.getBoundingClientRect?.();
  }
  if (!rect) return;
  let box = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentNode);
  while (box && box !== document.body && !(box.scrollHeight > box.clientHeight + 2 && /auto|scroll/.test(getComputedStyle(box).overflowY))) box = box.parentNode;
  if (!box || box === document.body) return;
  const b = box.getBoundingClientRect();
  const margin = 40;
  if (rect.top < b.top + margin) box.scrollTop -= b.top + margin - rect.top;
  else if (rect.bottom > b.bottom - margin) box.scrollTop += rect.bottom - (b.bottom - margin);
}
