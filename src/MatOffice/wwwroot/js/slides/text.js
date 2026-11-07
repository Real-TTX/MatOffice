/* MatOffice Slides – editing the text of a box or a table cell. The box becomes contenteditable; the model (paragraphs and runs) stays the truth:
   - typing and deleting inside one paragraph is left to the browser and read back after each input (every run carries its formatting in data-r),
   - everything that changes the paragraphs (Enter, Shift+Enter, joining paragraphs, a selection over several paragraphs, paste, cut) and every
     formatting command is done on the model, then the paragraphs are drawn again and the selection is put back (as paragraph/offset pairs). */
import { renderBody } from './render.js';
import { emptyBody } from './model.js';

const ZW = /​/g;

export const stable = o => JSON.stringify(o, Object.keys(o || {}).sort());
const propsOf = r => { const { t, ...rest } = r; return rest; };
const paraText = p => p.runs.map(r => r.t).join('');
const paraProps = p => { const { runs, ...rest } = p; return rest; };

/** Merges neighbouring runs with the same look; keeps one (maybe empty) run. */
export function tidy(p) {
  const out = [];
  for (const r of p.runs) {
    const last = out[out.length - 1];
    if (last && stable(propsOf(last)) === stable(propsOf(r))) last.t += r.t;
    else out.push({ ...r });
  }
  const kept = out.filter(r => r.t);
  p.runs = kept.length ? kept : [out[0] || { t: '' }];
  p.runs[0].t ??= '';
  return p;
}

/** Makes a run boundary at offset o of a paragraph; answers the index of the run that starts there. */
function splitAt(p, o) {
  let pos = 0;
  for (let i = 0; i < p.runs.length; i++) {
    const r = p.runs[i];
    if (o === pos) return i;
    if (o < pos + r.t.length) {
      p.runs.splice(i, 1, { ...r, t: r.t.slice(0, o - pos) }, { ...r, t: r.t.slice(o - pos) });
      return i + 1;
    }
    pos += r.t.length;
  }
  return p.runs.length;
}

/** The look of the run before an offset (what typed text takes over). */
export function runAt(p, o) {
  let pos = 0;
  for (const r of p.runs) {
    if (o <= pos + r.t.length && (o > pos || pos === 0)) return propsOf(r);
    pos += r.t.length;
  }
  return propsOf(p.runs[p.runs.length - 1] || {});
}

const before = (a, b) => a.p < b.p || (a.p === b.p && a.o <= b.o);

/** Sets (or with null removes) run properties on a range of the paragraphs. */
export function applyRunProps(paras, start, end, props) {
  for (let i = start.p; i <= end.p; i++) {
    const p = paras[i];
    const len = paraText(p).length;
    const a = i === start.p ? start.o : 0;
    const b = i === end.p ? end.o : len;
    let from, to;
    if (len === 0) { from = 0; to = p.runs.length; }
    else { if (a === b) continue; from = splitAt(p, a); to = splitAt(p, b); }
    for (let k = from; k < to; k++) {
      for (const [key, value] of Object.entries(props)) {
        if (value === null || value === undefined) delete p.runs[k][key];
        else p.runs[k][key] = value;
      }
    }
    tidy(p);
  }
}

/** Deletes a range; answers the position where it was. */
export function deleteRange(paras, start, end) {
  if (start.p === end.p) {
    const p = paras[start.p];
    const a = splitAt(p, start.o), b = splitAt(p, end.o);
    const keep = p.runs[a] ? propsOf(p.runs[a]) : runAt(p, start.o);
    p.runs.splice(a, b - a);
    if (!p.runs.length) p.runs.push({ ...keep, t: '' });
    tidy(p);
    return { ...start };
  }
  const first = paras[start.p], last = paras[end.p];
  splitAt(first, start.o);
  const headLen = start.o;
  let pos = 0;
  first.runs = first.runs.filter(r => { const keep = pos < headLen; pos += r.t.length; return keep; });
  const b = splitAt(last, end.o);
  const tail = last.runs.slice(b);
  first.runs.push(...tail);
  if (!first.runs.length) first.runs.push({ t: '' });
  tidy(first);
  paras.splice(start.p + 1, end.p - start.p);
  return { ...start };
}

/** Inserts text (with "\n" as soft line breaks, or as new paragraphs when asParagraphs) at a position; answers the position after it. */
export function insertText(paras, pos, text, props, asParagraphs = false) {
  const lines = asParagraphs ? text.split('\n') : [text];
  let at = { ...pos };
  lines.forEach((line, i) => {
    if (i > 0) at = splitParagraph(paras, at);
    if (!line) return;
    const p = paras[at.p];
    const k = splitAt(p, at.o);
    p.runs.splice(k, 0, { ...(props || runAt(p, at.o)), t: line });
    tidy(p);
    at = { p: at.p, o: at.o + line.length };
  });
  return at;
}

/** Splits a paragraph at a position (Enter): the new one keeps the paragraph's look and the look of the text at the caret. */
export function splitParagraph(paras, pos) {
  const p = paras[pos.p];
  const look = runAt(p, pos.o);
  const k = splitAt(p, pos.o);
  const tail = p.runs.splice(k);
  if (!p.runs.length) p.runs.push({ ...look, t: '' });
  const next = { ...paraProps(p), runs: tail.length ? tail : [{ ...look, t: '' }] };
  tidy(p);
  tidy(next);
  paras.splice(pos.p + 1, 0, next);
  return { p: pos.p + 1, o: 0 };
}

/** The word around an offset (formatting with a caret inside a word formats the word, like the office suites). */
function wordAt(p, o) {
  const text = paraText(p);
  const isWord = c => !!c && /[\p{L}\p{N}_'’-]/u.test(c);
  if (!isWord(text[o - 1]) || !isWord(text[o])) return null;
  let a = o, b = o;
  while (a > 0 && isWord(text[a - 1])) a--;
  while (b < text.length && isWord(text[b])) b++;
  return [a, b];
}

/** The text of a range ("\n" between paragraphs). */
export function textIn(paras, start, end) {
  const out = [];
  for (let i = start.p; i <= end.p; i++) {
    const text = paraText(paras[i]);
    out.push(text.slice(i === start.p ? start.o : 0, i === end.p ? end.o : text.length));
  }
  return out.join('\n');
}

export class TextEditor {
  constructor(app) {
    this.app = app;
    this.active = null;
    this.onSel = () => this.selectionChanged();
  }

  get editing() { return !!this.active; }

  /** Starts editing: target { id, cell: [r, c] | null }; where: { point: [clientX, clientY], all, end }. */
  start(target, where = {}) {
    if (this.active) this.stop();
    const el = this.app.slide?.elements.find(e => e.id === target.id);
    if (!el || this.app.readOnly) return false;
    if (!target.cell && !el.text) el.text = emptyBody(el.type === 'shape' ? { anchor: 'middle' } : {});
    this.active = { id: el.id, cell: target.cell || null, burst: 0, pending: null, pendingAt: null, last: null };
    const bodyEl = this.app.canvas.renderEditing(this.active);
    if (!bodyEl) { this.active = null; return false; }
    this.mount(bodyEl);
    this.place(where);
    this.app.selectionChanged();
    return true;
  }

  get el() { return this.app.slide?.elements.find(e => e.id === this.active?.id); }

  get body() {
    const el = this.el;
    if (!el) return null;
    if (this.active.cell) { const [r, c] = this.active.cell; return el.rows[r]?.cells[c]?.text || emptyBody(); }
    return el.text;
  }

  set body(b) {
    const el = this.el;
    if (!el) return;
    if (this.active.cell) { const [r, c] = this.active.cell; el.rows[r].cells[c].text = b; }
    else el.text = b;
  }

  mount(bodyEl) {
    this.bodyEl = bodyEl;
    bodyEl.contentEditable = 'true';
    bodyEl.spellcheck = true;
    bodyEl.setAttribute('role', 'textbox');
    bodyEl.setAttribute('aria-multiline', 'true');
    bodyEl.classList.add('mo-slides-editing');
    this.handlers = {
      keydown: e => this.onKey(e),
      beforeinput: e => this.onBeforeInput(e),
      input: () => this.sync(),
      paste: e => this.onPaste(e),
      cut: e => this.onCut(e),
      drop: e => e.preventDefault(),
      compositionend: () => this.sync(),
    };
    for (const [k, fn] of Object.entries(this.handlers)) bodyEl.addEventListener(k, fn);
    document.addEventListener('selectionchange', this.onSel);
    bodyEl.focus({ preventScroll: true });
  }

  /** Puts the caret where the person clicked, at the end, or selects everything. */
  place(where) {
    const paras = this.body.paras;
    if (where.point && document.caretRangeFromPoint) {
      const r = document.caretRangeFromPoint(where.point[0], where.point[1]);
      if (r && this.bodyEl.contains(r.startContainer)) {
        const sel = getSelection();
        sel.removeAllRanges();
        sel.addRange(r);
        this.active.last = this.getRange();
        return;
      }
    }
    if (where.all) {
      const last = paras.length - 1;
      this.setRange({ start: { p: 0, o: 0 }, end: { p: last, o: paraText(paras[last]).length } });
      return;
    }
    const last = paras.length - 1;
    const end = { p: last, o: paraText(paras[last]).length };
    this.setRange({ start: end, end });
  }

  /** Ends editing (the text is already in the model). */
  stop() {
    if (!this.active) return;
    this.sync(true);
    const bodyEl = this.bodyEl;
    if (bodyEl) {
      for (const [k, fn] of Object.entries(this.handlers)) bodyEl.removeEventListener(k, fn);
      bodyEl.contentEditable = 'false';
      bodyEl.classList.remove('mo-slides-editing');
    }
    document.removeEventListener('selectionchange', this.onSel);
    const { id, cell } = this.active;
    this.active = null;
    this.bodyEl = null;
    if (getSelection().rangeCount && bodyEl?.contains(getSelection().anchorNode)) getSelection().removeAllRanges();
    this.app.editingStopped(id, cell);
  }

  // ---- reading the DOM back -------------------------------------------------------------------------------------------------------

  paraNodes() { return [...this.bodyEl.children].filter(n => n.classList?.contains('mo-sr-p')); }

  /** The paragraphs as they are on the screen now. */
  read() {
    const old = this.body.paras;
    return this.paraNodes().map((pd, i) => {
      let pp = {};
      try { pp = pd.dataset.p ? JSON.parse(pd.dataset.p) : paraProps(old[i] || old[old.length - 1] || {}); } catch { pp = {}; }
      const items = [];
      const walk = (node, props) => {
        for (const ch of node.childNodes) {
          if (ch.nodeType === 3) { const t = ch.data.replace(ZW, ''); if (t) items.push({ t, props }); }
          else if (ch.nodeName === 'BR') items.push({ t: '\n', props, br: true });
          else if (ch.nodeType === 1) {
            let p = props;
            if (ch.dataset?.r) { try { p = JSON.parse(ch.dataset.r); } catch { /* keep */ } }
            walk(ch, p);
          }
        }
      };
      walk(pd, null);
      if (items.length && items[items.length - 1].br) items.pop();
      const firstProps = (() => { const s = pd.querySelector('[data-r]'); try { return s ? JSON.parse(s.dataset.r) : propsOf(old[i]?.runs[0] || {}); } catch { return {}; } })();
      let prev = null;
      for (const it of items) { if (it.props) prev = it.props; else it.props = prev; }
      let next = null;
      for (let k = items.length - 1; k >= 0; k--) { if (items[k].props) next = items[k].props; else items[k].props = next || firstProps; }
      const p = { ...pp, runs: items.length ? items.map(it => ({ ...(it.props || {}), t: it.t })) : [{ ...firstProps, t: '' }] };
      return tidy(p);
    });
  }

  /** Reads the screen into the model (after typing). */
  sync(final = false) {
    if (!this.active || !this.bodyEl) return;
    const paras = this.read();
    if (!paras.length) paras.push({ runs: [{ t: '' }] });
    const before = JSON.stringify(this.body.paras);
    if (JSON.stringify(paras) === before) return;
    this.app.checkpoint('typing:' + this.active.id);
    this.body = { ...this.body, paras };
    this.app.textChanged(final);
  }

  // ---- selection as paragraph/offset pairs -------------------------------------------------------------------------------------

  pointToOffset(node, offset) {
    const paras = this.paraNodes();
    if (node === this.bodyEl) {
      const k = Math.min(offset, paras.length);
      if (k >= paras.length) { const p = paras.length - 1; return { p, o: this.lengthOf(paras[p]) }; }
      return { p: Math.max(0, k), o: 0 };
    }
    const el = node.nodeType === 1 ? node : node.parentElement;
    const pd = el?.closest('.mo-sr-p');
    const p = Math.max(0, paras.indexOf(pd));
    if (!pd) return { p: 0, o: 0 };
    const r = document.createRange();
    r.setStart(pd, 0);
    try { r.setEnd(node, offset); } catch { return { p, o: 0 }; }
    return { p, o: Math.min(this.count(r.cloneContents()), this.lengthOf(pd)) };
  }

  count(fragment) {
    let n = 0;
    const walk = node => {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) n += ch.data.replace(ZW, '').length;
        else if (ch.nodeName === 'BR') n += 1;
        else walk(ch);
      }
    };
    walk(fragment);
    return n;
  }

  lengthOf(pd) {
    const r = document.createRange();
    r.selectNodeContents(pd);
    let n = this.count(r.cloneContents());
    const last = lastLeaf(pd);
    if (last?.nodeName === 'BR') n -= 1;
    return Math.max(0, n);
  }

  offsetToPoint({ p, o }) {
    const pd = this.paraNodes()[Math.min(p, this.paraNodes().length - 1)];
    if (!pd) return { node: this.bodyEl, offset: 0 };
    let rest = o;
    const walker = document.createTreeWalker(pd, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    let firstSpan = null;
    while (walker.nextNode()) {
      const n = walker.currentNode;
      if (n.nodeType === 1 && n.dataset?.r && !firstSpan) firstSpan = n;
      if (n.nodeType === 3) {
        const len = n.data.replace(ZW, '').length;
        if (rest <= len) return { node: n, offset: rest };
        rest -= len;
      } else if (n.nodeName === 'BR') {
        if (rest === 0) return { node: n.parentNode, offset: [...n.parentNode.childNodes].indexOf(n) };
        rest -= 1;
      }
    }
    if (firstSpan && !firstSpan.textContent) return { node: firstSpan, offset: 0 };
    const last = pd.lastChild;
    if (last?.nodeName === 'BR') return { node: pd, offset: pd.childNodes.length - 1 };
    return { node: pd, offset: pd.childNodes.length };
  }

  getRange() {
    const sel = getSelection();
    if (!sel.rangeCount || !this.bodyEl) return null;
    const r = sel.getRangeAt(0);
    if (!this.bodyEl.contains(r.startContainer) || !this.bodyEl.contains(r.endContainer)) return null;
    const start = this.pointToOffset(r.startContainer, r.startOffset);
    const end = this.pointToOffset(r.endContainer, r.endOffset);
    return before(start, end) ? { start, end } : { start: end, end: start };
  }

  setRange(range) {
    if (!range || !this.bodyEl) return;
    const a = this.offsetToPoint(range.start), b = this.offsetToPoint(range.end);
    const r = document.createRange();
    try { r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); } catch { return; }
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    this.active.last = { start: { ...range.start }, end: { ...range.end } };
  }

  /** The range to act on: the selection, else the last one known (a ribbon control may have the focus). */
  range() {
    const live = this.getRange();
    if (live && this.active) this.active.last = live;
    return live || this.active?.last || { start: { p: 0, o: 0 }, end: { p: 0, o: 0 } };
  }

  selectionChanged() {
    if (!this.active) return;
    const r = this.getRange();
    if (!r) return;
    const moved = !this.active.pendingAt || r.start.p !== this.active.pendingAt.p || r.start.o !== this.active.pendingAt.o;
    if (this.active.pending && moved) { this.active.pending = null; this.active.pendingAt = null; }
    this.active.last = r;
    this.app.selectionChanged(true);
  }

  /** Gives the focus back to the text (after a ribbon control had it); the caret stays where it is, or comes back where it was. */
  refocus() {
    if (!this.bodyEl) return;
    const live = this.getRange();
    this.bodyEl.focus({ preventScroll: true });
    if (live) this.setRange(live);
    else if (this.active?.last) this.setRange(this.active.last);
  }

  // ---- changes on the model --------------------------------------------------------------------------------------------------------

  /** Changes the paragraphs on the model, draws them again and puts the selection (fn answers it) back. */
  change(fn) {
    if (!this.active) return;
    this.sync();
    this.app.checkpoint();
    const paras = structuredClone(this.body.paras);
    const range = fn(paras) || this.range();
    this.body = { ...this.body, paras };
    this.redraw(range);
    this.app.textChanged();
  }

  redraw(range) {
    const { content } = this.app;
    const look = this.active.cell ? this.app.canvas.cellLook(this.el, this.active.cell) : null;
    const fresh = renderBody(document, content, this.app.slide, this.el, this.body, { editable: true }, look);
    this.bodyEl.replaceChildren(...fresh.childNodes);
    this.bodyEl.focus({ preventScroll: true });
    if (range) this.setRange(range);
  }

  insert(range, text, asParagraphs = false) {
    this.change(paras => {
      const at = deleteRange(paras, range.start, range.end);
      const pending = this.active.pending;
      const props = pending ? { ...runAt(paras[at.p], at.o), ...pending } : null;
      const end = insertText(paras, at, text, props, asParagraphs);
      if (pending) { this.active.pending = null; this.active.pendingAt = null; }
      return { start: end, end };
    });
  }

  remove(range) { this.change(paras => { const at = deleteRange(paras, range.start, range.end); return { start: at, end: at }; }); }

  splitParagraph(range) {
    this.change(paras => {
      const at = deleteRange(paras, range.start, range.end);
      const next = splitParagraph(paras, at);
      return { start: next, end: next };
    });
  }

  joinBackward(pos) {
    if (pos.p === 0) {
      const p = this.body.paras[0];
      if (p.level) this.change(paras => { paras[0].level = Math.max(0, (paras[0].level || 0) - 1) || undefined; if (!paras[0].level) delete paras[0].level; return { start: pos, end: pos }; });
      return;
    }
    this.change(paras => {
      const prevLen = paraText(paras[pos.p - 1]).length;
      return { start: deleteRange(paras, { p: pos.p - 1, o: prevLen }, { p: pos.p, o: 0 }), end: { p: pos.p - 1, o: prevLen } };
    });
  }

  joinForward(pos) {
    if (pos.p >= this.body.paras.length - 1) return;
    this.change(paras => { deleteRange(paras, pos, { p: pos.p + 1, o: 0 }); return { start: pos, end: pos }; });
  }

  /** The look of the text at the selection (for the ribbon). */
  runFormat() {
    if (!this.active) return {};
    const r = this.range();
    const paras = this.body.paras;
    const p = paras[Math.min(r.start.p, paras.length - 1)];
    const atSel = r.start.p === r.end.p && r.start.o === r.end.o ? runAt(p, r.start.o) : runAt(p, r.start.o + 1);
    return { ...atSel, ...(this.active.pending || {}) };
  }

  /** The paragraphs of the selection. */
  selectedParagraphs() {
    if (!this.active) return [];
    const r = this.range();
    return this.body.paras.slice(r.start.p, r.end.p + 1);
  }

  /** Formats the selected text (a caret inside a word formats the word; elsewhere the next typed text). */
  applyRun(props) {
    if (!this.active) return;
    const range = this.range();
    let target = range;
    if (range.start.p === range.end.p && range.start.o === range.end.o) {
      const word = wordAt(this.body.paras[range.start.p], range.start.o);
      if (!word) {
        this.active.pending = { ...(this.active.pending || {}), ...props };
        this.active.pendingAt = { ...range.start };
        this.refocus();
        this.app.selectionChanged(true);
        return;
      }
      target = { start: { p: range.start.p, o: word[0] }, end: { p: range.start.p, o: word[1] } };
    }
    this.change(paras => { applyRunProps(paras, target.start, target.end, props); return range; });
  }

  /** Changes the selected paragraphs: fn(paragraph, index) changes it. */
  applyPara(fn) {
    if (!this.active) return;
    const range = this.range();
    this.change(paras => { for (let i = range.start.p; i <= range.end.p; i++) fn(paras[i], i); return range; });
  }

  changeLevel(delta) {
    this.applyPara(p => {
      const level = Math.min(8, Math.max(0, (p.level || 0) + delta));
      if (level) p.level = level; else delete p.level;
    });
  }

  selectedText() {
    if (!this.active) return '';
    const r = this.range();
    return textIn(this.body.paras, r.start, r.end);
  }

  // ---- events -----------------------------------------------------------------------------------------------------------------------

  onBeforeInput(e) {
    const type = e.inputType;
    const range = this.getRange();
    if (!range) return;
    const collapsed = range.start.p === range.end.p && range.start.o === range.end.o;
    const multi = range.start.p !== range.end.p;
    if (type === 'insertParagraph') { e.preventDefault(); this.splitParagraph(range); return; }
    if (type === 'insertLineBreak') { e.preventDefault(); this.insert(range, '\n'); return; }
    if (type === 'insertText' || type === 'insertReplacementText') {
      if (multi || this.active.pending) { e.preventDefault(); this.insert(range, e.data ?? e.dataTransfer?.getData('text/plain') ?? ''); }
      return;
    }
    if (type.startsWith('delete') && type !== 'deleteByCut' && type !== 'deleteByDrag') {
      if (multi) { e.preventDefault(); this.remove(range); return; }
      const backward = type.includes('Backward');
      if (collapsed && backward && range.start.o === 0) { e.preventDefault(); this.joinBackward(range.start); return; }
      if (collapsed && !backward && range.start.o >= paraText(this.body.paras[range.start.p]).length) { e.preventDefault(); this.joinForward(range.start); return; }
      return;
    }
    if (type === 'formatBold' || type === 'formatItalic' || type === 'formatUnderline' || type === 'formatStrikeThrough') {
      e.preventDefault();
      const key = { formatBold: 'b', formatItalic: 'i', formatUnderline: 'u', formatStrikeThrough: 's' }[type];
      this.app.toggleRun(key);
      return;
    }
    if (type === 'historyUndo') { e.preventDefault(); this.app.undo(); return; }
    if (type === 'historyRedo') { e.preventDefault(); this.app.redo(); return; }
    if (type.startsWith('insertFrom') || type === 'deleteByCut' || type === 'deleteByDrag' || type.startsWith('format')) e.preventDefault();
  }

  onKey(e) {
    if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === 'z' || e.key === 'Z' || e.key === 'y' || e.key === 'Y')) {
      e.preventDefault();
      if (e.key === 'y' || e.key === 'Y' || e.shiftKey) this.app.redo(); else this.app.undo();
      return;
    }
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.stop(); this.app.canvas.focus(); return; }
    if (e.key === 'F2') { e.preventDefault(); this.stop(); this.app.canvas.focus(); return; }
    if (e.key === 'Tab') {
      e.preventDefault();
      if (this.active.cell) { this.app.moveCell(e.shiftKey ? -1 : 1); return; }
      const r = this.range();
      if (r.start.o === 0 || r.start.p !== r.end.p) this.changeLevel(e.shiftKey ? -1 : 1);
      else if (!e.shiftKey) this.insert(r, '\t');
      return;
    }
    if (e.altKey && e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) { e.preventDefault(); this.changeLevel(e.key === 'ArrowRight' ? 1 : -1); }
  }

  onPaste(e) {
    e.preventDefault();
    const files = [...(e.clipboardData?.files || [])].filter(f => f.type.startsWith('image/'));
    if (files.length && !e.clipboardData.getData('text/plain')) { this.stop(); this.app.insertPictures(files); return; }
    const text = (e.clipboardData?.getData('text/plain') || '').replace(/\r\n?/g, '\n');
    if (!text) return;
    this.insert(this.range(), text, true);
  }

  onCut(e) {
    const r = this.range();
    if (r.start.p === r.end.p && r.start.o === r.end.o) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', textIn(this.body.paras, r.start, r.end));
    this.remove(r);
  }
}

function lastLeaf(node) {
  let n = node.lastChild;
  while (n && n.nodeType === 1 && n.nodeName !== 'BR' && n.lastChild) n = n.lastChild;
  return n;
}
