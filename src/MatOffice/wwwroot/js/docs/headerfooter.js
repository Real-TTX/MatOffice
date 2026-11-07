/* MatOffice Docs – headers and footers, edited in place like in Word: a double click into the top or bottom margin of a page (or Insert →
   Header/Footer) makes that page's header an editing host, the body steps back; Escape, Close or a click into the body ends it. Every sheet
   shows the header with its own page number; the first page can have its own (Different First Page). */
import { parseHtml, readBlocks, renderBlock, domLength } from './htmlio.js';
import { writeBlocks, textRun } from './model.js';
import { caretFromPoint } from './editor.js';

export class HeaderFooter {
  constructor(app) {
    this.app = app;
    this.ed = app.editor;
    this.editing = null;
    const scroller = app.view.scroller;
    scroller.addEventListener('dblclick', e => {
      if (app.ctx.readOnly || this.editing) return;
      const hit = app.view.hitSheet(e.clientX, e.clientY);
      if (hit && (hit.area === 'header' || hit.area === 'footer')) { e.preventDefault(); this.open(hit.area, hit.index); }
    });
    scroller.addEventListener('pointerdown', e => {
      if (!this.editing || this.editing.host.contains(e.target) || e.target.closest?.('.mo-docs-imgsel')) return;
      const hit = app.view.hitSheet(e.clientX, e.clientY);
      if (hit && hit.area === this.editing.kind && hit.index === this.editing.index) return;
      const { clientX, clientY } = e;
      this.close();
      if (hit?.area === 'body') {
        e.preventDefault();
        requestAnimationFrame(() => {
          const p = caretFromPoint(clientX, clientY);
          if (p && app.body.contains(p.node)) {
            const r = document.createRange();
            r.setStart(p.node, p.offset);
            app.body.focus({ preventScroll: true });
            const sel = document.getSelection();
            sel.removeAllRanges();
            sel.addRange(r);
          }
        });
      }
    });
  }

  keyOf(kind, index) {
    const first = index === 0 && this.app.page.differentFirst;
    return kind === 'header' ? (first ? 'firstHeader' : 'header') : first ? 'firstFooter' : 'footer';
  }

  hosts() { return this.editing ? [this.editing.host] : []; }

  dirty(host) {
    if (this.editing?.host !== host) return;
    clearTimeout(this.timer);
    this.timer = setTimeout(() => { if (this.editing) { this.commit(); this.app.view.setHeaders(this.app.headers); } }, 400);
  }

  /** Edits the header or footer of a page (default: the page in view). */
  open(kind, index) {
    const app = this.app;
    if (app.ctx.readOnly) return;
    if (app.viewMode !== 'print') app.setView('print');
    this.close();
    if (index === undefined) index = Math.min(app.view.pages - 1, app.view.visiblePage());
    const sheet = app.view.sheets.children[index] || app.view.sheets.firstChild;
    if (!sheet) return;
    const box = sheet.querySelector(kind === 'header' ? '.mo-docs-hdr' : '.mo-docs-ftr');
    const key = this.keyOf(kind, index);
    const blocks = parseHtml(app.headers[key]);
    box.classList.add('is-editing');
    box.replaceChildren(...(blocks.length ? blocks : [{ type: 'p', css: {}, runs: [] }]).map(b => renderBlock(b, this.renderOptions(index))));
    this.ed.addHost(box, kind);
    box.setAttribute('lang', app.settings.language);
    box.setAttribute('spellcheck', app.settings.spellcheck ? 'true' : 'false');
    const tag = app.ctx.el('div', { class: 'mo-docs-hf-tag', 'data-mo-deco': '', text: app.ctx.t(kind === 'header' ? (key === 'firstHeader' ? 'Docs.Header.First' : 'Docs.Header') : (key === 'firstFooter' ? 'Docs.Footer.First' : 'Docs.Footer')) });
    tag.style.left = 'var(--ml)';
    tag.style.top = kind === 'header' ? `calc(var(--mt) + 2px)` : `calc(var(--ph) - var(--mb) - 22px)`;
    sheet.append(tag);
    app.view.stack.classList.add('is-hf');
    app.view.flow.style.pointerEvents = 'none';
    app.body.setAttribute('contenteditable', 'false');
    this.editing = { kind, index, key, host: box, tag, sheet };
    const blocksIn = this.ed.textBlocks(box);
    const last = blocksIn[blocksIn.length - 1];
    if (last) this.ed.select(last, domLength(last));
    app.ctx.ribbon.set(app.buildRibbon());
    app.ctx.ribbon.select('hf');
  }

  renderOptions(index) {
    const base = this.ed.renderOptions();
    return { ...base, fieldText: kind => (kind === 'page' ? String(index + 1) : base.fieldText(kind)) };
  }

  /** Takes the edited header into the content (pooled: only for the undo history, nothing is closed). */
  commit() {
    const e = this.editing;
    if (!e) return;
    const blocks = readBlocks(e.host);
    const empty = blocks.every(b => b.runs && !b.runs.length && !Object.keys(b.css || {}).length && b.type === 'p');
    const html = empty ? '' : writeBlocks(blocks);
    this.app.headers[e.key] = html;
  }

  /** After undo: the header being edited shows the restored text. */
  reload() {
    const e = this.editing;
    if (!e) return;
    const blocks = parseHtml(this.app.headers[e.key]);
    this.ed.quietly(() => e.host.replaceChildren(...(blocks.length ? blocks : [{ type: 'p', css: {}, runs: [] }]).map(b => renderBlock(b, this.renderOptions(e.index)))));
  }

  close(commit = true) {
    const e = this.editing;
    if (!e) return;
    if (commit) this.commit();
    this.editing = null;
    this.ed.removeHost(e.host);
    e.host.removeAttribute('contenteditable');
    e.host.removeAttribute('role');
    e.host.removeAttribute('aria-multiline');
    e.host.classList.remove('is-editing');
    e.tag.remove();
    const app = this.app;
    app.view.stack.classList.remove('is-hf');
    app.view.flow.style.pointerEvents = '';
    if (!app.ctx.readOnly) app.body.setAttribute('contenteditable', 'true');
    e.sheet._key = null;
    app.view.setHeaders(app.headers);
    if (commit) app.changed();
    app.ctx.ribbon.set(app.buildRibbon());
    app.ctx.ribbon.select('insert');
  }

  remove(kind) {
    const app = this.app;
    app.editor.checkpoint('header');
    if (this.editing?.kind === kind) this.close(false);
    if (kind === 'header') { app.headers.header = ''; app.headers.firstHeader = ''; } else { app.headers.footer = ''; app.headers.firstFooter = ''; }
    app.view.setHeaders(app.headers);
    app.changed();
  }

  /** Page numbers: a centred (or right-aligned) number in the header or footer, like Word's "Plain Number" presets. */
  pageNumbers(kind, right = false) {
    const app = this.app;
    app.editor.checkpoint('header');
    this.close();
    const block = { type: 'p', css: right ? { 'text-align': 'right' } : { 'text-align': 'center' }, runs: [{ obj: 'field', kind: 'page', f: {} }] };
    const html = writeBlocks([block]);
    if (kind === 'header') { app.headers.header = html; if (app.page.differentFirst) app.headers.firstHeader = ''; }
    else { app.headers.footer = html; if (app.page.differentFirst) app.headers.firstFooter = ''; }
    app.view.setHeaders(app.headers);
    app.changed();
  }

  removePageNumbers() {
    const app = this.app;
    app.editor.checkpoint('header');
    this.close();
    for (const key of Object.keys(app.headers)) {
      if (!app.headers[key]) continue;
      const blocks = parseHtml(app.headers[key]).map(b => (b.runs ? { ...b, runs: b.runs.filter(r => !(r.obj === 'field' && (r.kind === 'page' || r.kind === 'pages'))) } : b));
      const html = writeBlocks(blocks);
      app.headers[key] = blocks.every(b => b.runs && !b.runs.some(r => r.obj || r.text.trim())) ? '' : html;
    }
    app.view.setHeaders(app.headers);
    app.changed();
  }

  /** "Page 1 of 3" at the caret of the header being edited. */
  insertPageOf() {
    if (!this.editing) return;
    const text = this.app.ctx.t('Docs.PageOf', '\u0001', '\u0002');
    const runs = [];
    for (const part of text.split(/(\u0001|\u0002)/)) {
      if (part === '\u0001') runs.push({ obj: 'field', kind: 'page', f: {} });
      else if (part === '\u0002') runs.push({ obj: 'field', kind: 'pages', f: {} });
      else if (part) runs.push(textRun(part));
    }
    this.ed.insertBlocks([{ type: 'p', css: {}, runs }]);
  }

  toggleDifferentFirst() {
    const app = this.app;
    const wasEditing = this.editing ? { kind: this.editing.kind, index: this.editing.index } : null;
    this.close();
    app.layoutCommands.setPage({ differentFirst: !app.page.differentFirst });
    if (wasEditing) this.open(wasEditing.kind, wasEditing.index);
  }
}
