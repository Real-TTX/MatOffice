/* MatOffice Docs – notes (footnotes and endnotes of Word, shown as endnotes at the end of the document): a reference <sup data-fn="n1"></sup>
   in the text, numbered by the position of its first reference; the text of each note is a small editing host in the notes area after the body
   (it flows over the pages with the text). Click a reference to go to its note, click the number of a note to go back. */
import { parseHtml, readBlocks, renderBlock, domLength } from './htmlio.js';
import { writeBlocks, isItemId, newId, cleanFormat } from './model.js';

export class Notes {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    this.map = new Map();       // id → { html, host }
    this.numbers = new Map();   // id → number
    this.area = app.view.notes;
    app.body.addEventListener('click', e => {
      const ref = e.target.closest?.('sup[data-fn]');
      if (ref) { e.preventDefault(); this.goToNote(ref.getAttribute('data-fn')); }
    });
  }

  load(notes) {
    for (const entry of this.map.values()) if (entry.host) this.ed.removeHost(entry.host);
    this.map = new Map();
    for (const n of notes) if (n && isItemId(n.id) && !this.map.has(n.id)) this.map.set(n.id, { html: typeof n.html === 'string' ? writeBlocks(parseHtml(n.html)) : '' });
    this.renumber(true);
  }

  hosts() { return [...this.map.values()].map(e => e.host).filter(Boolean); }

  /** The html of a note (from its host while it is shown). */
  htmlOf(id) {
    const e = this.map.get(id);
    if (!e) return '';
    if (e.host && e.dirty !== false) { e.html = writeBlocks(readBlocks(e.host)); e.dirty = false; }
    return e.html;
  }

  dirty(host) { for (const e of this.map.values()) if (e.host === host) e.dirty = true; }

  list() { return [...this.map.keys()].map(id => ({ id, html: this.htmlOf(id) || '<p></p>' })); }

  snapshot(poolKey) {
    return JSON.stringify([...this.map.keys()].map(id => {
      const e = this.map.get(id);
      const html = e.host ? writeBlocks(readBlocks(e.host), { image: poolKey }) : writeBlocks(parseHtml(e.html), { image: poolKey });
      return [id, html];
    }));
  }

  restore(json, poolUrl) {
    const list = JSON.parse(json || '[]');
    const keep = new Set(list.map(([id]) => id));
    for (const [id, e] of this.map) if (!keep.has(id)) { if (e.host) { this.ed.removeHost(e.host); e.host.closest('.mo-docs-note')?.remove(); } this.map.delete(id); }
    for (const [id, html] of list) {
      const blocks = parseHtml(html, { images: poolUrl });
      const e = this.map.get(id) || {};
      e.html = writeBlocks(blocks);
      e.dirty = false;
      if (e.host) this.ed.quietly(() => e.host.replaceChildren(...(blocks.length ? blocks : [{ type: 'p', css: {}, runs: [] }]).map(b => renderBlock(b, this.ed.renderOptions()))));
      this.map.set(id, e);
    }
    this.renumber(true);
  }

  number(id) { return this.numbers.get(id); }

  /** The numbers of the references (the order of first appearance) and the notes area in that order. */
  renumber(force = false) {
    const refs = [...this.app.body.querySelectorAll('sup[data-fn]')];
    const numbers = new Map();
    for (const r of refs) { const id = r.getAttribute('data-fn'); if (this.map.has(id) && !numbers.has(id)) numbers.set(id, numbers.size + 1); }
    const changed = force || numbers.size !== this.numbers.size || [...numbers].some(([id, n]) => this.numbers.get(id) !== n);
    this.numbers = numbers;
    this.ed.quietly(() => {
      for (const r of refs) { const n = String(numbers.get(r.getAttribute('data-fn')) ?? ''); if (r.getAttribute('data-mo-n') !== n) r.setAttribute('data-mo-n', n); }
    });
    if (changed) this.render();
  }

  render() {
    const { t, el } = this.app.ctx;
    const ids = [...this.numbers.keys()];
    this.ed.quietly(() => {
      if (!ids.length) { this.area.hidden = true; this.area.replaceChildren(); return; }
      this.area.hidden = false;
      const rows = ids.map(id => {
        const e = this.map.get(id);
        if (!e.host) {
          e.host = el('div', { class: 'mo-docs-note__body mo-docs-content', 'aria-label': t('Docs.Note.Text', this.numbers.get(id)) });
          const blocks = parseHtml(e.html);
          e.host.append(...(blocks.length ? blocks : [{ type: 'p', css: {}, runs: [] }]).map(b => renderBlock(b, this.ed.renderOptions())));
          this.ed.addHost(e.host, 'note');
          e.host.setAttribute('lang', this.app.settings.language);
          e.host.setAttribute('spellcheck', this.app.settings.spellcheck ? 'true' : 'false');
          e.dirty = false;
        }
        const row = e.host.closest('.mo-docs-note') || el('div', { class: 'mo-docs-note', 'data-note': id },
          el('span', { class: 'mo-docs-note__num', role: 'button', tabindex: '0', 'data-tip': t('Docs.Note.Back'), onclick: () => this.goToRef(id) }), e.host);
        row.querySelector('.mo-docs-note__num').textContent = String(this.numbers.get(id));
        return row;
      });
      this.area.replaceChildren(el('div', { class: 'mo-docs-notes__rule' }), el('p', { class: 'mo-docs-notes__title', text: t('Docs.Notes') }), ...rows);
    });
  }

  /** Inserts a note at the caret and puts the caret into its text. */
  insert() {
    if (this.app.ctx.readOnly) return;
    this.ed.focus();
    const sel = this.ed.selection();
    if (!sel || sel.host !== this.app.body) return;
    let id;
    do id = newId('n'); while (this.map.has(id));
    this.ed.checkpoint('note');
    this.map.set(id, { html: '<p></p>' });
    const at = sel.collapsed ? sel.start : sel.end;
    this.ed.select(at.block, at.offset);
    this.ed.insertBlocks([{ type: 'p', css: {}, runs: [{ obj: 'note', id, f: cleanFormat({}) }] }]);
    this.renumber(true);
    this.app.view.update();
    requestAnimationFrame(() => this.goToNote(id));
  }

  goToNote(id) {
    const e = this.map.get(id);
    if (!e?.host) return;
    const row = e.host.closest('.mo-docs-note');
    row?.classList.add('is-target');
    setTimeout(() => row?.classList.remove('is-target'), 1500);
    const blocks = this.ed.textBlocks(e.host);
    const last = blocks[blocks.length - 1];
    if (last) this.ed.select(last, domLength(last));
    else e.host.scrollIntoView({ block: 'center' });
  }

  goToRef(id) {
    const ref = this.app.body.querySelector(`sup[data-fn="${CSS.escape(id)}"]`);
    if (!ref) return;
    const block = ref.closest('p,h1,h2,h3,h4,h5,h6,blockquote,pre');
    ref.scrollIntoView({ block: 'center' });
    if (block) {
      const r = document.createRange();
      r.setStartAfter(ref);
      r.collapse(true);
      this.app.body.focus({ preventScroll: true });
      const sel = document.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    }
  }

  /** The next note reference after the caret (Next Footnote). */
  next() {
    const refs = [...this.app.body.querySelectorAll('sup[data-fn]')];
    if (!refs.length) return;
    const sel = document.getSelection();
    const r = sel.rangeCount ? sel.getRangeAt(0) : null;
    const after = r ? refs.find(x => r.comparePoint(x, 0) > 0) : null;
    const target = after || refs[0];
    target.scrollIntoView({ block: 'center' });
    const range = document.createRange();
    range.setStartAfter(target);
    range.collapse(true);
    this.app.body.focus({ preventScroll: true });
    sel.removeAllRanges();
    sel.addRange(range);
  }

  show() {
    if (this.area.hidden) { this.app.toastInfo('Docs.Note.None'); return; }
    this.area.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }
}
