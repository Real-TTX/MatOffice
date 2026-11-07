/* MatOffice Docs – the navigation pane (View → Navigation Pane: the headings of the document, click to go there, the current section marked,
   a search box) and the table of contents (References: insert, update, remove – TOC1–TOC3 paragraphs with links to the ids of the headings;
   the page numbers next to them are only shown on the screen and in print, they are not content). */
import { readBlock, renderBlock, domLength } from './htmlio.js';
import { tocBlocks, newId, runsText } from './model.js';

export class Nav {
  constructor(app) {
    this.app = app;
    const { t, el, icon } = app.ctx;
    this.visible = false;
    this.list = el('div', { class: 'mo-docs-nav__list', role: 'navigation', 'aria-label': t('Docs.Nav') });
    const search = el('input', { class: 'mat-input', type: 'search', placeholder: t('Docs.Nav.Search'), 'aria-label': t('Docs.Nav.Search') });
    search.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); app.findUi.query.value = search.value; app.findUi.open(false); }
    });
    this.el = el('aside', { class: 'mo-docs-nav', hidden: true },
      el('div', { class: 'mo-docs-nav__head' }, el('span', { text: t('Docs.Nav') }),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Docs.Close'), onclick: () => this.toggle() }, icon('x'))),
      el('div', { class: 'mo-docs-nav__search' }, search), this.list);
  }

  toggle() {
    this.visible = !this.visible;
    this.el.hidden = !this.visible;
    if (this.visible) this.refresh();
    const view = this.app.view;
    if (view.fitting) requestAnimationFrame(() => view.fit());
    this.app.ribbonRefresh();
  }

  headings() {
    return [...this.app.body.querySelectorAll(':scope > h1, :scope > h2, :scope > h3, :scope > h4, :scope > h5, :scope > h6')];
  }

  refresh() {
    if (!this.visible) return;
    const { t, el } = this.app.ctx;
    const items = this.headings().map(h => {
      const text = (h.textContent || '').replace(/\s+/g, ' ').trim() || t('Docs.Nav.Empty');
      const b = el('button', { type: 'button', class: 'mo-docs-nav__item', 'data-level': h.tagName[1], text, title: text });
      b._heading = h;
      b.addEventListener('click', () => {
        h.scrollIntoView({ block: 'start', behavior: 'smooth' });
        this.app.editor.select(h, 0, h, 0, false);
        if (window.innerWidth < 720) this.toggle();
      });
      return b;
    });
    this.list.replaceChildren(...(items.length ? items : [el('p', { class: 'mo-docs-nav__empty', text: t('Docs.Nav.None') })]));
    this.highlight();
  }

  /** Marks the heading of the section the caret is in. */
  highlight() {
    if (!this.visible) return;
    const sel = document.getSelection();
    const node = sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
    let current = null;
    if (node && this.app.body.contains(node)) {
      for (const b of this.list.querySelectorAll('.mo-docs-nav__item')) {
        if (b._heading.compareDocumentPosition(node) & (Node.DOCUMENT_POSITION_FOLLOWING | Node.DOCUMENT_POSITION_CONTAINED_BY) || b._heading.contains(node)) current = b;
      }
    }
    for (const b of this.list.querySelectorAll('.mo-docs-nav__item')) b.classList.toggle('is-active', b === current);
  }

  // ---- table of contents --------------------------------------------------------------------------------------------------------

  tocRange() {
    const kids = [...this.app.body.children];
    const start = kids.findIndex(n => /^TOC/.test(n.getAttribute('data-style') || ''));
    if (start < 0) return null;
    let end = start;
    while (end + 1 < kids.length && /^TOC[123]$/.test(kids[end + 1].getAttribute('data-style') || '')) end++;
    return kids.slice(start, end + 1);
  }

  hasToc() { return !!this.tocRange(); }

  /** The headings 1–3 with ids (new ids for the ones without). */
  entries() {
    return this.headings().filter(h => Number(h.tagName[1]) <= 3).map(h => {
      if (!h.id) {
        let id;
        do id = newId('h-'); while (document.getElementById(id));
        h.id = id;
      }
      return { level: Number(h.tagName[1]), text: runsText(readBlock(h).runs, ''), id: h.id };
    });
  }

  tocNodes() {
    return tocBlocks(this.entries(), this.app.ctx.t('Docs.Toc.Title')).map(b => renderBlock(b, this.app.editor.renderOptions()));
  }

  insertToc() {
    const app = this.app;
    if (app.ctx.readOnly) return;
    if (this.hasToc()) { this.updateToc(); return; }
    app.editor.focus();
    const sel = app.editor.selection();
    app.editor.checkpoint('toc');
    const nodes = this.tocNodes();
    let at = sel && sel.host === app.body ? sel.start.block : app.body.firstElementChild;
    while (at && at.parentNode !== app.body) at = at.parentNode;
    if (at && sel && sel.start.offset > 0 && domLength(sel.start.block) > 0 && at === sel.start.block) at = at.nextElementSibling;
    if (at) for (const n of nodes) app.body.insertBefore(n, at);
    else app.body.append(...nodes);
    // an empty paragraph the table was inserted into goes (unless it is the last one)
    if (at && at.matches('p') && !at.textContent && !at.querySelector('img,[data-field],[data-fn]') && at.nextElementSibling) at.remove();
    app.editor.ensureTail(app.body);
    app.changed();
    setTimeout(() => app.tocPages(), 200);
  }

  updateToc() {
    const app = this.app;
    const range = this.tocRange();
    if (!range || app.ctx.readOnly) return;
    app.editor.checkpoint('toc');
    const nodes = this.tocNodes();
    const first = range[0];
    const keepTitle = first.getAttribute('data-style') === 'TOCHeading';
    if (keepTitle) nodes.shift();
    const anchor = keepTitle ? first.nextSibling : first;
    for (const n of nodes) app.body.insertBefore(n, anchor);
    for (const n of range) if (!(keepTitle && n === first)) n.remove();
    app.changed();
    setTimeout(() => app.tocPages(), 200);
    app.toastInfo('Docs.Toc.Updated');
  }

  removeToc() {
    const app = this.app;
    const range = this.tocRange();
    if (!range || app.ctx.readOnly) return;
    app.editor.checkpoint('toc');
    for (const n of range) n.remove();
    app.editor.ensureTail(app.body);
    app.changed();
  }
}
