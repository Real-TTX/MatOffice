/* MatOffice Docs – comments like Word's: a comment on the selected text (<span data-comment="c1"> around it), the cards in the margin next to
   their text with author and date, replies, resolve/reopen, edit, delete; previous/next; on narrow screens (and in Web Layout) a panel at the
   bottom. The comments themselves live in content.comments; a comment whose text was deleted is not saved any more. */
import { readBlock, domLength } from './htmlio.js';
import { formatRuns, newId, cleanText } from './model.js';
import { stamp } from '../editor/util.js';

const RAIL = 280;

export class Comments {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    this.visible = true;
    this.active = null;
    this.editing = null;
    this.sheet = null;
    app.body.addEventListener('click', e => {
      const span = e.target.closest?.('span[data-comment]');
      if (span) this.activate(span.getAttribute('data-comment').split(/\s+/).pop(), false);
    });
    new ResizeObserver(() => this.render()).observe(app.view.scroller);
  }

  count() { return this.anchored().length; }

  /** The comments that still have text in the body, in the order of the document: [{ comment, span }]. */
  anchored() {
    const first = new Map();
    for (const span of this.app.body.querySelectorAll('span[data-comment]')) {
      for (const id of span.getAttribute('data-comment').split(/\s+/)) if (!first.has(id)) first.set(id, span);
    }
    return this.app.comments.filter(c => first.has(c.id)).map(c => ({ comment: c, span: first.get(c.id) }))
      .sort((a, b) => (a.span === b.span ? 0 : a.span.compareDocumentPosition(b.span) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1));
  }

  wide() { return this.app.viewMode === 'print' && this.app.view.scroller.clientWidth >= (this.app.view.geometry?.paperWidth || 800) + RAIL + 80; }

  toggle() { this.visible = !this.visible; this.render(); }

  /** Draws the cards (in the margin or in the panel) and marks resolved and active anchors. */
  render() {
    const app = this.app;
    const list = this.anchored();
    app.shell.classList.toggle('mo-docs-hide-comments', !this.visible);
    this.ed.quietly(() => {
      for (const span of app.body.querySelectorAll('span[data-comment]')) {
        const ids = span.getAttribute('data-comment').split(/\s+/);
        const comments = ids.map(id => app.comments.find(c => c.id === id)).filter(Boolean);
        const done = comments.length > 0 && comments.every(c => c.done);
        if (done !== span.hasAttribute('data-mo-done')) span.toggleAttribute('data-mo-done', done);
        const active = this.active && ids.includes(this.active);
        if (!!active !== span.hasAttribute('data-mo-active')) span.toggleAttribute('data-mo-active', !!active);
      }
    });
    const inRail = this.visible && list.length > 0 && this.wide();
    app.view.setRail(inRail ? RAIL : 0);
    if (inRail) {
      this.closeSheet();
      this.drawRail(list);
    } else {
      app.view.rail.replaceChildren();
      if (this.sheet) this.drawSheet(list);
    }
  }

  drawRail(list) {
    const rail = this.app.view.rail;
    const cards = list.map(({ comment, span }) => {
      const card = this.card(comment);
      card.dataset.top = String(this.app.view.local(span.getBoundingClientRect()).top);
      return card;
    });
    rail.replaceChildren(...cards);
    // stacked without overlapping, the active one where its text is
    let bottom = 0;
    for (const card of cards) {
      const want = Math.max(0, Number(card.dataset.top));
      const top = Math.max(want, bottom + 8);
      card.style.top = top + 'px';
      bottom = top + card.offsetHeight;
    }
  }

  drawSheet(list) {
    const { t, el, icon } = this.app.ctx;
    const head = el('div', { class: 'mo-docs-commentsheet__head' }, el('span', { text: t('Docs.Comments') + ` (${list.length})` }),
      el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm', 'aria-label': t('Docs.Close'), onclick: () => this.closeSheet() }, icon('x')));
    this.sheet.replaceChildren(head, ...(list.length ? list.map(({ comment }) => this.card(comment)) : [el('p', { class: 'mo-muted', text: t('Docs.Comment.None') })]));
  }

  openSheet() {
    if (!this.sheet) {
      this.sheet = this.app.ctx.el('div', { class: 'mo-docs-commentsheet', role: 'region', 'aria-label': this.app.ctx.t('Docs.Comments') });
      this.app.shell.append(this.sheet);
    }
    this.drawSheet(this.anchored());
  }

  closeSheet() { this.sheet?.remove(); this.sheet = null; }

  card(c) {
    const { t, el, icon } = this.app.ctx;
    const ro = this.app.ctx.readOnly;
    const card = el('div', { class: 'mo-docs-card' + (c.id === this.active ? ' is-active' : '') + (c.done ? ' is-done' : ''), 'data-id': c.id, tabindex: '-1' });
    card.addEventListener('click', e => { if (!e.target.closest('button,textarea')) this.activate(c.id, true); });
    const person = (who, when, text) => el('div', {},
      el('div', { class: 'mo-docs-card__head' }, el('span', { class: 'mo-docs-card__avatar', 'aria-hidden': 'true', text: initials(who) }),
        el('span', { class: 'mo-docs-card__who' }, el('span', { class: 'mo-docs-card__name', text: who || t('Docs.Comment.Someone') }), el('span', { class: 'mo-docs-card__date', text: when ? stamp(when) : '' }))),
      text !== null ? el('p', { class: 'mo-docs-card__text', text }) : null);
    if (this.editing === c.id && !ro) {
      const area = el('textarea', { class: 'mo-docs-card__edit', 'aria-label': t('Docs.Comment.Text'), maxlength: 10000 });
      area.value = c.text;
      const save = () => this.saveText(c.id, area.value);
      area.addEventListener('keydown', e => {
        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); save(); }
        else if (e.key === 'Escape') { e.preventDefault(); this.editing = null; if (!c.text) this.remove(c.id, false); else this.render(); }
      });
      card.append(person(c.author, c.date, null), area, el('div', { class: 'mo-docs-card__actions' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary mat-btn--sm', onclick: save }, el('span', { text: t('Docs.Comment.Post') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => { this.editing = null; if (!c.text) this.remove(c.id, false); else this.render(); } }, el('span', { text: t('Docs.Cancel') }))));
      requestAnimationFrame(() => { area.focus(); area.setSelectionRange(area.value.length, area.value.length); });
    } else {
      card.append(person(c.author, c.date, c.text));
      if (c.done) card.querySelector('.mo-docs-card__name')?.append(el('span', { class: 'mo-docs-card__badge', text: t('Docs.Comment.Resolved') }));
    }
    for (const r of c.replies) card.append(el('div', { class: 'mo-docs-card__reply' }, person(r.author, r.date, r.text)));
    if (!ro && c.id === this.active && this.editing !== c.id) {
      const reply = el('textarea', { class: 'mo-docs-card__edit', placeholder: t('Docs.Comment.Reply'), 'aria-label': t('Docs.Comment.Reply'), rows: 2, maxlength: 10000 });
      const send = () => { if (reply.value.trim()) this.reply(c.id, reply.value); };
      reply.addEventListener('keydown', e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); send(); } });
      card.append(reply, el('div', { class: 'mo-docs-card__actions' },
        el('button', { type: 'button', class: 'mat-btn mat-btn--primary mat-btn--sm', onclick: send }, icon('message-square'), el('span', { text: t('Docs.Comment.ReplyButton') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--secondary mat-btn--sm', onclick: () => this.resolve(c.id) }, icon(c.done ? 'refresh' : 'check'), el('span', { text: t(c.done ? 'Docs.Comment.Reopen' : 'Docs.Comment.Resolve') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--sm', onclick: () => { this.editing = c.id; this.render(); } }, icon('edit'), el('span', { text: t('Docs.Comment.Edit') })),
        el('button', { type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--sm', 'aria-label': t('Docs.Comment.Delete'), 'data-tip': t('Docs.Comment.Delete'), onclick: () => this.remove(c.id) }, icon('trash'))));
    }
    return card;
  }

  /** Makes a comment the active one (its card raised, its text marked); scroll = bring its text into view. */
  activate(id, scroll) {
    if (this.active === id && !scroll) return;
    this.active = id;
    this.render();
    if (scroll) {
      const span = this.app.body.querySelector(`span[data-comment~="${CSS.escape(id)}"]`);
      span?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
    if (!this.wide() && this.visible) this.openSheet();
    this.app.ribbonRefresh();
  }

  /** The caret moved: the comment under it becomes active. */
  follow() {
    const sel = document.getSelection();
    if (!sel.rangeCount) return;
    const node = sel.getRangeAt(0).startContainer;
    const span = (node.nodeType === 1 ? node : node.parentNode)?.closest?.('span[data-comment]');
    const id = span && this.app.body.contains(span) ? span.getAttribute('data-comment').split(/\s+/).pop() : null;
    if (id && id !== this.active) this.activate(id, false);
  }

  /** New comment on the selection (or the word at the caret). */
  add() {
    const app = this.app;
    if (app.ctx.readOnly) return;
    this.ed.focus();
    let sel = this.ed.selection();
    if (!sel || sel.host !== app.body) return;
    if (sel.collapsed) {
      const text = readBlock(sel.start.block).runs.map(r => (r.obj ? ' ' : r.text)).join('');
      let s = sel.start.offset, e = s;
      while (s > 0 && /[\p{L}\p{N}_]/u.test(text[s - 1])) s--;
      while (e < text.length && /[\p{L}\p{N}_]/u.test(text[e])) e++;
      if (s === e) { if (!text.length) { app.toastInfo('Docs.Comment.SelectText'); return; } s = Math.max(0, s - 1); e = Math.min(text.length, Math.max(e, s + 1)); }
      this.ed.select(sel.start.block, s, sel.start.block, e, false);
      sel = this.ed.selection();
    }
    let id;
    do id = newId('c'); while (app.comments.some(c => c.id === id));
    this.ed.checkpoint('comment');
    const mark = this.ed.bookmark(sel);
    for (const { el, start, end } of this.ed.ranges(sel)) {
      if (end <= start) continue;
      const b = readBlock(el);
      b.runs = formatRuns(b.runs, start, end, f => ({ ...f, cmt: [...new Set([...(f.cmt ? f.cmt.split(' ') : []), id])].join(' ') }));
      this.ed.morph(el, b);
    }
    this.ed.restoreBookmark(mark, false);
    app.comments.push({ id, author: app.ctx.doc.user?.name || app.ctx.t('Docs.Comment.Someone'), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), text: '', done: false, replies: [] });
    this.visible = true;
    this.editing = id;
    this.active = id;
    app.changed();
    this.render();
    if (!this.wide()) this.openSheet();
  }

  saveText(id, text) {
    const c = this.app.comments.find(x => x.id === id);
    if (!c) return;
    const clean = cleanText(text).trim().slice(0, 10000);
    this.editing = null;
    if (!clean && !c.text) { this.remove(id, false); return; }
    this.ed.checkpoint('comment');
    c.text = clean || c.text;
    this.app.changed();
    this.render();
  }

  reply(id, text) {
    const c = this.app.comments.find(x => x.id === id);
    if (!c) return;
    const clean = cleanText(text).trim().slice(0, 10000);
    if (!clean || c.replies.length >= 500) return;
    this.ed.checkpoint('comment');
    let rid;
    do rid = newId('r'); while (this.app.comments.some(x => x.id === rid || x.replies.some(r => r.id === rid)));
    c.replies.push({ id: rid, author: this.app.ctx.doc.user?.name || this.app.ctx.t('Docs.Comment.Someone'), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), text: clean });
    this.app.changed();
    this.render();
  }

  resolve(id) {
    const c = this.app.comments.find(x => x.id === id);
    if (!c) return;
    this.ed.checkpoint('comment');
    c.done = !c.done;
    this.app.changed();
    this.render();
  }

  /** Deletes a comment and its marks in the text. */
  remove(id, checkpoint = true) {
    if (!id) return;
    const app = this.app;
    if (checkpoint) this.ed.checkpoint('comment');
    const mark = this.ed.bookmark();
    const blocks = new Set([...app.body.querySelectorAll(`span[data-comment~="${CSS.escape(id)}"]`)].map(s => s.closest('p,h1,h2,h3,h4,h5,h6,blockquote,pre')).filter(Boolean));
    for (const el of blocks) {
      const b = readBlock(el);
      b.runs = formatRuns(b.runs, 0, domLength(el), f => {
        if (!f.cmt) return f;
        const ids = f.cmt.split(' ').filter(x => x !== id);
        const g = { ...f };
        if (ids.length) g.cmt = ids.join(' '); else delete g.cmt;
        return g;
      });
      this.ed.morph(el, b);
    }
    app.comments = app.comments.filter(c => c.id !== id);
    if (this.active === id) this.active = null;
    if (this.editing === id) this.editing = null;
    if (mark) this.ed.restoreBookmark(mark, false);
    app.changed();
    this.render();
  }

  /** Previous / next comment. */
  step(dir) {
    const list = this.anchored();
    if (!list.length) return;
    const i = list.findIndex(x => x.comment.id === this.active);
    const next = list[(i + dir + list.length) % list.length] || list[0];
    this.visible = true;
    this.activate(next.comment.id, true);
  }
}

function initials(name) {
  const parts = String(name || '?').trim().split(/\s+/).filter(Boolean);
  return ((parts[0]?.[0] || '?') + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}
