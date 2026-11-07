/* MatOffice Docs – find and replace (Ctrl+F, Ctrl+H): a bar over the pages, every match marked (CSS Custom Highlight API, nothing changes in
   the DOM), previous/next with "3 of 12", match case, whole words, replace and replace all (one undo step; the replacement takes the format of
   the first character of each match). It searches the body and the notes. */
import { readBlock, pointAt } from './htmlio.js';
import { findAll, runsText, splitRuns, normalizeRuns, textRun } from './model.js';

export class Find {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    const { t, el, icon } = { t: k => app.ctx?.t(k) ?? k, el: (...a) => app.ctx.el(...a), icon: n => app.ctx.icon(n) };
    this.matches = [];
    this.index = -1;
    this.query = el('input', { class: 'mat-input', type: 'search', 'aria-label': t('Docs.Find.What'), placeholder: t('Docs.Find.What'), autocomplete: 'off', spellcheck: 'false' });
    this.with = el('input', { class: 'mat-input', type: 'text', 'aria-label': t('Docs.Replace.With'), placeholder: t('Docs.Replace.With'), autocomplete: 'off', spellcheck: 'false' });
    this.count = el('span', { class: 'mo-docs-find__count', 'aria-live': 'polite' });
    this.matchCase = el('input', { type: 'checkbox' });
    this.wholeWord = el('input', { type: 'checkbox' });
    const btn = (ic, label, run, cls = 'mat-btn--ghost mat-btn--icon') => el('button', { type: 'button', class: 'mat-btn mat-btn--sm ' + cls, 'aria-label': label, 'data-tip': label, onclick: run }, icon(ic));
    this.replaceRow = el('div', { class: 'mo-docs-find__row' }, this.with,
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => this.replaceOne() }, el('span', { text: t('Docs.Replace.One') })),
      el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => this.replaceAll() }, el('span', { text: t('Docs.Replace.All') })));
    this.el = el('div', { class: 'mo-docs-find', hidden: true, role: 'search', 'aria-label': t('Docs.Find') },
      el('div', { class: 'mo-docs-find__row' }, this.query, this.count,
        btn('chevron-up', t('Docs.Find.Previous'), () => this.step(-1)), btn('chevron-down', t('Docs.Find.Next'), () => this.step(1)),
        btn('docs-replace', t('Docs.Replace'), () => this.showReplace(this.replaceRow.hidden)), btn('x', t('Docs.Close'), () => this.close())),
      this.replaceRow,
      el('div', { class: 'mo-docs-find__opts' }, el('label', {}, this.matchCase, el('span', { text: t('Docs.Find.MatchCase') })), el('label', {}, this.wholeWord, el('span', { text: t('Docs.Find.WholeWord') }))));
    this.replaceRow.hidden = true;
    let timer = 0;
    this.query.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => this.search(true), 120); });
    this.query.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); this.step(e.shiftKey ? -1 : 1); }
      else if (e.key === 'Escape') { e.preventDefault(); this.close(); }
    });
    this.with.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); this.replaceOne(); }
      else if (e.key === 'Escape') { e.preventDefault(); this.close(); }
    });
    this.matchCase.addEventListener('change', () => this.search(true));
    this.wholeWord.addEventListener('change', () => this.search(true));
    // a change of the text makes the matches stale
    this.ed.addEventListener('selection', () => { if (!this.el.hidden && this.stale) this.search(false); });
  }

  open(replace) {
    const sel = this.ed.selection();
    if (sel && !sel.collapsed && sel.start.block === sel.end.block) {
      const text = runsText(readBlock(sel.start.block).runs).slice(sel.start.offset, sel.end.offset);
      if (text && text.length < 200 && !text.includes('\n')) this.query.value = text;
    }
    this.el.hidden = false;
    this.showReplace(replace && !this.app.ctx.readOnly);
    this.query.focus();
    this.query.select();
    this.search(true);
  }

  showReplace(on) { this.replaceRow.hidden = !on || this.app.ctx.readOnly; }

  close() {
    this.el.hidden = true;
    this.matches = [];
    this.highlight();
    this.ed.focus();
  }

  /** The text blocks that are searched: the body, then the notes. */
  blocks() {
    return [...this.ed.textBlocks(this.app.body), ...this.app.notes.hosts().flatMap(h => this.ed.textBlocks(h))];
  }

  search(fromCaret) {
    this.stale = false;
    const query = this.query.value;
    const options = { matchCase: this.matchCase.checked, wholeWord: this.wholeWord.checked };
    this.matches = [];
    if (query) {
      for (const block of this.blocks()) {
        const text = runsText(readBlock(block).runs);
        for (const [start, end] of findAll(text, query, options)) this.matches.push({ block, start, end });
      }
    }
    this.index = -1;
    if (this.matches.length) {
      const sel = this.ed.selection();
      let i = 0;
      if (fromCaret && sel) {
        const blocks = this.blocks();
        const caretBlock = blocks.indexOf(sel.start.block);
        i = this.matches.findIndex(m => { const b = blocks.indexOf(m.block); return b > caretBlock || (b === caretBlock && m.start >= sel.start.offset); });
        if (i < 0) i = 0;
      }
      this.index = i;
    }
    this.highlight();
    this.paintCount();
  }

  range(m) {
    const a = pointAt(m.block, m.start, true), b = pointAt(m.block, m.end);
    const r = new Range();
    try { r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); } catch { return null; }
    return r;
  }

  highlight() {
    if (typeof CSS === 'undefined' || !CSS.highlights || typeof Highlight === 'undefined') return;
    CSS.highlights.delete('mo-docs-find');
    CSS.highlights.delete('mo-docs-find-current');
    if (!this.matches.length || this.el.hidden) return;
    const all = this.matches.slice(0, 3000).map(m => this.range(m)).filter(Boolean);
    CSS.highlights.set('mo-docs-find', new Highlight(...all));
    const current = this.matches[this.index] && this.range(this.matches[this.index]);
    if (current) CSS.highlights.set('mo-docs-find-current', new Highlight(current));
  }

  paintCount() {
    const t = k => this.app.ctx.t(k);
    this.count.textContent = !this.query.value ? '' : this.matches.length ? this.app.ctx.t('Docs.Find.Count', this.index + 1, this.matches.length) : t('Docs.Find.None');
  }

  /** Goes to the next/previous match and selects it in the document. */
  step(dir) {
    if (this.stale || !this.matches.length) this.search(true);
    if (!this.matches.length) return;
    if (this.index < 0) this.index = 0;
    else this.index = (this.index + dir + this.matches.length) % this.matches.length;
    const m = this.matches[this.index];
    if (!m.block.isConnected) { this.search(true); return; }
    this.ed.select(m.block, m.start, m.block, m.end);
    const r = this.range(m);
    r?.startContainer.parentElement?.scrollIntoView({ block: 'center' });
    this.query.focus();
    this.highlight();
    this.paintCount();
  }

  replaceOne() {
    if (this.app.ctx.readOnly) return;
    const m = this.matches[this.index];
    const sel = this.ed.selection();
    if (!m || !sel || sel.start.block !== m.block || sel.start.offset !== m.start || sel.end.offset !== m.end) { this.step(1); return; }
    this.ed.checkpoint('replace');
    this.replaceIn(m.block, [[m.start, m.end]]);
    this.app.changed();
    this.search(false);
    const blocks = this.blocks();
    const at = blocks.indexOf(m.block);
    this.index = this.matches.findIndex(x => { const b = blocks.indexOf(x.block); return b > at || (b === at && x.start >= m.start + this.with.value.length); });
    if (this.index < 0 && this.matches.length) this.index = 0;
    if (this.index >= 0) { this.index -= 1; this.step(1); } else this.paintCount();
  }

  replaceAll() {
    if (this.app.ctx.readOnly) return;
    this.search(false);
    if (!this.matches.length) { this.app.ctx.toast(this.app.ctx.t('Docs.Find.None'), 'info'); return; }
    this.ed.checkpoint('replace');
    const byBlock = new Map();
    for (const m of this.matches) { if (!byBlock.has(m.block)) byBlock.set(m.block, []); byBlock.get(m.block).push([m.start, m.end]); }
    const n = this.matches.length;
    for (const [block, ranges] of byBlock) this.replaceIn(block, ranges);
    this.app.changed();
    this.search(false);
    this.app.ctx.toast(this.app.ctx.t('Docs.Replace.Done', n), 'success');
  }

  /** Replaces ranges of one block (from the back, so the offsets stay right). */
  replaceIn(block, ranges) {
    const b = readBlock(block);
    let runs = b.runs;
    for (const [start, end] of [...ranges].sort((x, y) => y[0] - x[0])) {
      const [head, rest] = splitRuns(runs, start);
      const [mid, tail] = splitRuns(rest, end - start);
      const f = (mid.find(r => !r.obj) || mid[0])?.f || {};
      runs = normalizeRuns([...head, ...(this.with.value ? [textRun(this.with.value, f)] : []), ...tail]);
    }
    b.runs = runs;
    this.ed.morph(block, b);
  }
}
