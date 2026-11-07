/* MatOffice Docs – the word processor (like Word). The entry of the editor: builds the work area (navigation pane, the desk with the pages,
   find and replace, the comments), connects the editing engine (editor.js), the page view (layout.js) and the content of the server
   (docs/EDITORS.md "Docs content"), the ribbon (ribbon.js), the keyboard and the status bar. */
import { startEditor } from '../editor/frame.js';
import { t, el, icon, mod, debounce, lang as pageLang } from '../editor/util.js';
import { DocEditor } from './editor.js';
import { PageView } from './layout.js';
import { parseHtml, readBlock, readBlocks, readTable, renderBlock, isTextBlockEl, counterSet } from './htmlio.js';
import { normalizePage, writeBlock, writeBlocks, countWords, isItemId, runsText, cleanText, languageName as nameOf } from './model.js';
import { createCommands, LINE_SPACINGS } from './commands.js';
import { buildRibbon } from './ribbon.js';
import { Tables } from './tables.js';
import { Images } from './images.js';
import { Comments } from './comments.js';
import { Notes } from './notes.js';
import { Find } from './find.js';
import { Nav } from './nav.js';
import { Dialogs } from './dialogs.js';
import { HeaderFooter } from './headerfooter.js';
import { thumbnail } from './thumb.js';

const app = {
  ctx: null, editor: null, view: null, body: null, commands: null,
  page: normalizePage({}), settings: { language: 'en', spellcheck: true },
  headers: { header: '', footer: '', firstHeader: '', firstFooter: '' },
  comments: [], marks: false, viewMode: 'print',
};

// ---- the cache of serialized blocks (saving and the undo history read every block; only changed ones are written again) ----------
const cache = new WeakMap();
const pool = { byUrl: new Map(), byKey: new Map(), next: 1 };
const poolKey = url => {
  let key = pool.byUrl.get(url);
  if (!key) { key = 'mo-img:' + pool.next++; pool.byUrl.set(url, key); pool.byKey.set(key, url); }
  return key;
};
const poolUrl = src => (src && src.startsWith('mo-img:') ? pool.byKey.get(src) || null : src);

function blockOfElement(n) {
  if (isTextBlockEl(n)) return readBlock(n);
  if (n.matches('table')) return readTable(n);
  if (n.matches('hr')) return { type: 'hr' };
  if (n.matches('div[data-break]')) return { type: 'break', kind: n.getAttribute('data-break') === 'column' ? 'column' : 'page' };
  return null;
}

function entry(n) {
  let e = cache.get(n);
  if (!e) {
    const block = blockOfElement(n);
    e = { full: block ? writeBlock(block) : '', pooled: block ? writeBlock(block, { image: poolKey }) : '' };
    cache.set(n, e);
  }
  return e;
}

const contentChildren = host => [...host.children].filter(n => !n.hasAttribute('data-mo-deco'));

/** The body as Docs HTML (pooled: pictures as keys of the pool, for the undo history). */
app.serializeBody = (pooled = false) => {
  const parts = contentChildren(app.body).map(n => (pooled ? entry(n).pooled : entry(n).full)).filter(Boolean);
  return parts.join('');
};

app.dirty = (target, host) => {
  if (host !== app.body) { app.notes?.dirty(host); app.hf?.dirty(host); return; }
  let n = target;
  while (n && n.parentNode !== app.body) n = n.parentNode;
  if (n && n.nodeType === 1) cache.delete(n);
};

app.language = () => app.settings.language || pageLang();
app.name = () => app.ctx?.doc?.mode === 'host' ? app.ctx.doc.fileName || '' : app.ctx?.doc?.name || '';
app.ribbonRefresh = () => app.ctx?.ribbon.refresh();
app.toastInfo = key => app.ctx?.toast(t(key), 'info');
app.renderOptions = () => ({
  fieldText: kind => (kind === 'page' ? String((app.view?.page ? 1 : 1)) : kind === 'pages' ? String(app.view?.pages || 1) : kind === 'title' ? app.name() : kind === 'date' ? new Date().toLocaleDateString(app.language()) : kind === 'time' ? new Date().toLocaleTimeString(app.language(), { hour: '2-digit', minute: '2-digit' }) : ''),
  noteNumber: id => app.notes?.number(id) ?? '',
  breakLabel: kind => t(kind === 'column' ? 'Docs.Break.column' : 'Docs.Break.page'),
});

// ---- content in and out ----------------------------------------------------------------------------------------------------------

function load(content) {
  const c = content && typeof content === 'object' ? content : {};
  app.page = normalizePage(c.page);
  app.settings = {
    language: typeof c.settings?.language === 'string' && /^[a-z]{2}(-[A-Z]{2})?$/.test(c.settings.language) ? c.settings.language : pageLang(),
    spellcheck: c.settings?.spellcheck !== false,
  };
  app.headers = { header: str(c.header), footer: str(c.footer), firstHeader: str(c.firstHeader), firstFooter: str(c.firstFooter) };
  for (const k of Object.keys(app.headers)) app.headers[k] = app.headers[k] ? writeBlocks(parseHtml(app.headers[k])) : '';
  app.comments = Array.isArray(c.comments) ? c.comments.filter(x => x && isItemId(x.id)).map(x => ({
    id: x.id, author: str(x.author), date: str(x.date), text: str(x.text), done: x.done === true,
    replies: Array.isArray(x.replies) ? x.replies.filter(r => r && isItemId(r.id)).map(r => ({ id: r.id, author: str(r.author), date: str(r.date), text: str(r.text) })) : [],
  })) : [];
  const blocks = parseHtml(str(c.html));
  app.editor.quietly(() => {
    app.body.replaceChildren(...(blocks.length ? blocks : [{ type: 'p', css: {}, runs: [] }]).map(b => renderBlock(b, app.renderOptions())));
    app.editor.ensureTail(app.body);
  });
  app.notes.load(Array.isArray(c.notes) ? c.notes : []);
  app.listPass?.();
  applySettings();
  app.view.setPage(app.page);
  app.view.setHeaders(app.headers);
  app.commentsUi?.render();
  app.navUi?.refresh();
}

const str = v => (typeof v === 'string' ? v : '');

function applySettings() {
  for (const host of [app.body, ...app.notes.hosts(), ...(app.hf?.hosts() || [])]) {
    host.setAttribute('lang', app.settings.language);
    host.setAttribute('spellcheck', app.settings.spellcheck ? 'true' : 'false');
  }
}
app.applySettings = applySettings;

/** The whole content as JSON for the server (the contract of docs/EDITORS.md). */
function getContent() {
  app.hf?.commit();
  const html = app.serializeBody(false);
  const referenced = new Set([...app.body.querySelectorAll('sup[data-fn]')].map(s => s.getAttribute('data-fn')));
  const anchored = new Set([...app.body.querySelectorAll('span[data-comment]')].flatMap(s => s.getAttribute('data-comment').split(/\s+/)));
  return {
    version: 1,
    page: app.page,
    html: html || '<p></p>',
    header: app.headers.header, footer: app.headers.footer, firstHeader: app.headers.firstHeader, firstFooter: app.headers.firstFooter,
    notes: app.notes.list().filter(n => referenced.has(n.id)),
    comments: app.comments.filter(c => anchored.has(c.id)).map(c => ({ ...c, replies: c.replies.map(r => ({ ...r })) })),
    settings: { language: app.settings.language, spellcheck: app.settings.spellcheck },
  };
}

/** The state for the undo history: unchanged blocks share their strings, pictures are keys of the pool. */
function snapshot() {
  app.hf?.commit(true);
  return {
    body: contentChildren(app.body).map(n => entry(n).pooled),
    headers: JSON.stringify(app.headers),
    notes: app.notes.snapshot(poolKey),
    comments: JSON.stringify(app.comments),
    page: JSON.stringify(app.page),
    settings: JSON.stringify(app.settings),
  };
}

function restore(s) {
  // the body: only the blocks between the common start and end are rendered again
  const children = contentChildren(app.body);
  const now = children.map(n => entry(n).pooled);
  let start = 0;
  while (start < now.length && start < s.body.length && now[start] === s.body[start]) start++;
  let endNow = now.length, endNew = s.body.length;
  while (endNow > start && endNew > start && now[endNow - 1] === s.body[endNew - 1]) { endNow--; endNew--; }
  const fresh = s.body.slice(start, endNew).map(html => {
    const block = parseHtml(html, { images: poolUrl })[0] || { type: 'p', css: {}, runs: [] };
    const node = renderBlock(block, app.renderOptions());
    cache.set(node, { full: writeBlock(block), pooled: html });
    return node;
  });
  const anchor = children[endNow] || null;
  for (const n of children.slice(start, endNow)) n.remove();
  for (const n of fresh) app.body.insertBefore(n, anchor);
  app.editor.ensureTail(app.body);
  app.headers = JSON.parse(s.headers);
  app.hf?.reload();
  app.view.setHeaders(app.headers);
  app.notes.restore(s.notes, poolUrl);
  app.comments = JSON.parse(s.comments);
  const page = JSON.parse(s.page);
  if (JSON.stringify(page) !== JSON.stringify(app.page)) { app.page = page; app.view.setPage(page); }
  app.settings = JSON.parse(s.settings);
  applySettings();
  app.commentsUi?.render();
  app.navUi?.refresh();
}

// ---- changes -------------------------------------------------------------------------------------------------------------------

/**
 * The numbers of lists that CSS alone cannot know: data-start, and a level whose kind changes (1. 2. then bullets then 1. again) starts anew.
 * An inline counter-set (only shown, never content: the serializer reads no counter-set).
 */
app.listPass = (root = app.body) => {
  app.editor.quietly(() => {
    for (const el of root.querySelectorAll('[data-list]')) {
      const level = Number(el.getAttribute('data-level')) || 0;
      const start = Number(el.getAttribute('data-start')) || 0;
      let want = '';
      if (start) want = counterSet(level, start);
      else {
        let p = el.previousElementSibling, kind = null;
        while (p && p.hasAttribute('data-list')) {
          const l = Number(p.getAttribute('data-level')) || 0;
          if (l === level) { kind = p.getAttribute('data-list'); break; }
          if (l < level) break;
          p = p.previousElementSibling;
        }
        if (kind && kind !== el.getAttribute('data-list')) want = counterSet(level, 1);
      }
      if (el.style.counterSet !== want) el.style.counterSet = want;
      if (!el.getAttribute('style')) el.removeAttribute('style');
    }
  });
};

const later = debounce(() => {
  app.listPass();
  if (app.findUi && !app.findUi.el.hidden) app.findUi.search(false);
  app.notes.renumber();
  app.view.update();
  app.navUi?.refresh();
  app.commentsUi?.render();
  updateStatus();
  app.tocPages();
}, 350);

function changed() {
  app.ctx.changed();
  app.view.schedule();
  later();
}
app.changed = changed;

/** The page numbers next to the entries of a table of contents (only shown, not content). */
app.tocPages = () => {
  if (app.viewMode !== 'print') return;
  app.editor.quietly(() => {
    for (const p of app.body.querySelectorAll('p[data-style^="TOC"]')) {
      if (p.getAttribute('data-style') === 'TOCHeading') continue;
      const href = p.querySelector('a[href^="#"]')?.getAttribute('href');
      const target = href && href.length > 1 ? app.body.querySelector('#' + CSS.escape(href.slice(1))) : null;
      const page = target ? String(app.view.pageOfRect(target.getBoundingClientRect()) + 1) : '';
      if (p.getAttribute('data-mo-page') !== page) { if (page) p.setAttribute('data-mo-page', page); else p.removeAttribute('data-mo-page'); }
    }
  });
};

// ---- the status bar --------------------------------------------------------------------------------------------------------------

const status = { page: null, words: null, lang: null };

function setupStatus(ctx) {
  status.page = el('button', { type: 'button', class: 'mo-docs-sbtn', 'data-tip': t('Docs.GoTo'), onclick: () => app.dialogs.goTo() });
  status.words = el('button', { type: 'button', class: 'mo-docs-sbtn', 'data-tip': t('Docs.WordCount'), onclick: () => app.dialogs.wordCount() });
  status.lang = el('button', { type: 'button', class: 'mo-docs-sbtn', 'data-tip': t('Docs.Language'), onclick: () => app.dialogs.language() });
  ctx.status.left(status.page, status.words, status.lang);
  const viewButton = (mode, ic, label) => el('button', {
    type: 'button', class: 'mo-docs-sbtn', 'aria-label': label, 'data-tip': label, 'aria-pressed': String(app.viewMode === mode), 'data-view': mode,
    onclick: () => app.setView(mode),
  }, icon(ic));
  status.views = el('div', { class: 'mo-docs-views' }, viewButton('print', 'docs-print-layout', t('Docs.View.Print')), viewButton('web', 'docs-web-layout', t('Docs.View.Web')));
  ctx.status.right(status.views);
  ctx.status.zoom({
    value: app.view.zoom, min: 0.25, max: 3,
    change: f => { app.view.setZoom(f); app.imagesUi?.position(); },
    fit: () => { app.view.fit(); ctx.status.setZoom(app.view.zoom); },
  });
}

function updateStatus() {
  if (!status.page) return;
  const sel = app.editor.selection();
  let page = app.view.visiblePage();
  if (sel && sel.host === app.body) {
    const r = sel.range.getClientRects()[0] || sel.start.block.getClientRects()[0];
    if (r) page = app.view.pageOfRect(r);
  }
  status.page.textContent = t('Docs.Status.Page', page + 1, app.view.pages);
  const words = countWords(app.body.textContent || '');
  status.words.textContent = t(words === 1 ? 'Docs.Status.Word' : 'Docs.Status.Words', words.toLocaleString(app.language()));
  status.lang.textContent = languageName(app.settings.language);
  for (const b of status.views.querySelectorAll('[data-view]')) b.setAttribute('aria-pressed', String(b.dataset.view === app.viewMode));
}
app.updateStatus = updateStatus;

function languageName(code) { return nameOf(code, pageLang()); }
app.uiLanguage = () => pageLang();

// ---- views ---------------------------------------------------------------------------------------------------------------------

app.setView = mode => {
  app.viewMode = mode;
  app.shell.classList.toggle('is-web', mode === 'web');
  app.view.setView(mode);
  if (mode === 'web') app.view.setZoom(1);
  else if (window.innerWidth < 720) app.view.fit();
  app.ctx.status.setZoom(app.view.zoom);
  app.commentsUi?.render();
  updateStatus();
  app.ribbonRefresh();
};

app.toggleMarks = () => {
  app.marks = !app.marks;
  app.shell.classList.toggle('mo-marks', app.marks);
  app.view.schedule();
};

// ---- keys -------------------------------------------------------------------------------------------------------------------------

function onKey(e) {
  const c = app.commands;
  const m = mod(e);
  const key = e.key.length === 1 ? e.key.toLowerCase() : e.key;
  if (e.key === 'Escape') {
    if (c.painter.active) { c.painterStop(); return true; }
    if (app.hf?.editing) { app.hf.close(); return true; }
    return false;
  }
  if (e.key === 'F3' && e.shiftKey && !m) { c.changeCase(null); return true; }
  if (!m) return false;
  const ro = app.ctx.readOnly;
  const edit = fn => { if (!ro) fn(); return true; };
  if (e.altKey && !e.shiftKey) {
    const digit = /^Digit([1-6])$/.exec(e.code);
    if (digit) return edit(() => c.style('Heading' + digit[1]));
    if (key === 'm' || e.code === 'KeyM') return edit(() => app.commentsUi.add());
    if (key === 'f' || e.code === 'KeyF' || key === 'd' || e.code === 'KeyD') return edit(() => app.notes.insert());
    return false;
  }
  if (e.shiftKey) {
    switch (e.code) {
      case 'KeyL': return edit(() => c.list('bullet'));
      case 'KeyN': return edit(() => c.style('Normal'));
      case 'KeyC': c.painterStart(false); return true;
      case 'KeyV': app.editor.plainNext = true; return false;
      case 'KeyM': return edit(() => c.indent(-1));
      case 'KeyZ': return edit(() => app.editor.redo());
      case 'Period': return edit(() => c.grow(1));
      case 'Comma': return edit(() => c.grow(-1));
      case 'Enter': case 'NumpadEnter': return edit(() => c.pageBreak('column'));
      default: break;
    }
    if (key === '+' || key === '=' || key === '*') return edit(() => c.toggle('sup'));
    if (key === '>') return edit(() => c.grow(1));
    if (key === '<') return edit(() => c.grow(-1));
    return false;
  }
  switch (key) {
    case 'b': return edit(() => c.toggle('b'));
    case 'i': return edit(() => c.toggle('i'));
    case 'u': return edit(() => c.toggle('u'));
    case 'e': return edit(() => c.align('center'));
    case 'l': return edit(() => c.align('left'));
    case 'r': return edit(() => c.align('right'));
    case 'j': return edit(() => c.align('justify'));
    case 'z': return edit(() => app.editor.undo());
    case 'y': return edit(() => app.editor.redo());
    case 'k': return edit(() => app.dialogs.link());
    case 'f': app.findUi.open(false); return true;
    case 'h': if (!ro) app.findUi.open(true); else app.findUi.open(false); return true;
    case 'g': app.dialogs.goTo(); return true;
    case 'd': return edit(() => app.dialogs.font());
    case 'm': return edit(() => c.indent(1));
    case '=': return edit(() => c.toggle('sub'));
    case ']': return edit(() => c.grow(1));
    case '[': return edit(() => c.grow(-1));
    case ' ': return edit(() => app.editor.formatText(f => ({ href: f.href, cmt: f.cmt }), { word: false }));
    case '1': return edit(() => c.lineSpacing(LINE_SPACINGS[0][1]));
    case '2': return edit(() => c.lineSpacing(LINE_SPACINGS[3][1]));
    case '5': return edit(() => c.lineSpacing(LINE_SPACINGS[2][1]));
    case '0': return edit(() => c.spacing('before', v => (v >= 12 ? 0 : 12)));
    case 'Enter': return edit(() => c.pageBreak('page'));
    default: return false;
  }
}

// ---- mount ------------------------------------------------------------------------------------------------------------------------

async function mount(ctx) {
  app.ctx = ctx;
  app.view = new PageView({ el, t, lang: pageLang(), name: () => app.name() });
  app.body = app.view.body;
  app.editor = new DocEditor({
    readOnly: ctx.readOnly,
    snapshot, restore, changed: () => changed(), dirty: (target, host) => app.dirty(target, host),
    renderOptions: () => app.renderOptions(), onKey,
    onPasteFiles: files => app.imagesUi.insertFiles(files),
    clipboardText: blocks => blocks.map(b => (b.type === 'table' ? b.rows.map(r => r.cells.map(c => c.blocks.map(x => runsText(x.runs || [], '')).join(' ')).join('\t')).join('\n') : runsText(b.runs || [], ''))).join('\n'),
  });
  app.view.quietly = fn => app.editor.quietly(fn);
  app.view.onPages = () => { updateStatus(); };
  app.editor.addHost(app.body, 'body');

  app.shell = el('div', { class: 'mo-docs-shell' });
  app.navUi = new Nav(app);
  app.findUi = new Find(app);
  app.shell.append(app.navUi.el, app.view.scroller, app.findUi.el);
  ctx.work.replaceChildren(app.shell);

  app.commands = createCommands(app);
  app.dialogs = new Dialogs(app);
  app.tables = new Tables(app);
  app.imagesUi = new Images(app);
  app.notes = new Notes(app);
  app.commentsUi = new Comments(app);
  app.hf = new HeaderFooter(app);

  load(ctx.doc.content);
  setupStatus(ctx);
  ctx.ribbon.set(buildRibbon(app));
  if (ctx.readOnly) app.editor.setReadOnly(true);

  // the ribbon follows the selection (contextual tabs for tables and pictures)
  let frame = 0;
  let context = '';
  app.editor.addEventListener('selection', () => {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      const now = app.contextKey();
      if (now !== context) { context = now; ctx.ribbon.set(buildRibbon(app)); }
      else ctx.ribbon.refresh();
      updateStatus();
      app.navUi.highlight();
      app.commentsUi.follow();
    });
  });
  app.editor.addEventListener('tab', e => onTab(e.detail.shift));
  app.body.addEventListener('mouseup', () => { if (app.commands.painter.active) setTimeout(() => app.commands.painterApply(), 0); });
  app.view.scroller.addEventListener('scroll', debounce(() => updateStatus(), 120), { passive: true });

  // first view: page width on narrow screens (web layout on a phone)
  if (window.innerWidth < 720) app.setView('web');
  else { app.view.setZoom(1); ctx.status.setZoom(1); }
  if (!ctx.readOnly) {
    const first = app.editor.textBlocks(app.body)[0];
    if (first) app.editor.select(first, 0, first, 0, false);
  }
  app.view.update();
  updateStatus();
  setTimeout(() => app.tocPages(), 300);
}

/** Tab: in a table to the next cell, in a list one level deeper, else a tab character. */
function onTab(shift) {
  const s = app.editor.state();
  if (!s.sel) return;
  if (s.sel.start.block.closest('td,th') && app.tables.tab(shift)) return;
  if (s.block?.list && (s.sel.start.offset === 0 || !s.sel.collapsed)) { app.commands.indent(shift ? -1 : 1); return; }
  if (shift) { if (s.block?.list) app.commands.indent(-1); return; }
  app.commands.text('\t');
}

/** Which contextual tabs the ribbon shows now: a table, a picture, a header or footer. */
app.contextKey = () => [app.tables?.current() ? 'table' : '', app.imagesUi?.selected ? 'image' : '', app.hf?.editing ? 'hf' : ''].join('|');

// ---- page setup, settings, zoom, the clipboard of the ribbon ----------------------------------------------------------------------

app.buildRibbon = () => buildRibbon(app);

app.layoutCommands = {
  /** Changes the page setup (one undo step). */
  setPage(changes) {
    if (app.ctx.readOnly) return;
    app.editor.checkpoint('page');
    const next = { ...app.page, ...changes, margins: { ...app.page.margins, ...(changes.margins || {}) } };
    if (changes.size && changes.size !== 'Custom' && !changes.width) { delete next.width; delete next.height; }
    app.page = normalizePage(next);
    app.view.setPage(app.page);
    app.changed();
    app.ribbonRefresh();
  },
  setSettings(changes) {
    if (app.ctx.readOnly) return;
    app.editor.checkpoint('settings');
    app.settings = { ...app.settings, ...changes };
    applySettings();
    app.changed();
    updateStatus();
    app.ribbonRefresh();
  },
};

app.zoomTo = f => { app.view.setZoom(f); app.ctx.status.setZoom(app.view.zoom); app.imagesUi?.position(); };

/** Paste from the ribbon: the clipboard of the system through the browser (it may ask for permission); else the hint to press Ctrl+V. */
app.clipboardPaste = async plain => {
  if (app.ctx.readOnly) return;
  app.editor.focus();
  const mark = app.editor.bookmark();
  try {
    if (!navigator.clipboard?.read) throw new Error('no clipboard');
    const items = await navigator.clipboard.read();
    if (mark) app.editor.restoreBookmark(mark, false);
    for (const item of items) {
      if (!plain && item.types.includes('text/html')) {
        const html = await (await item.getType('text/html')).text();
        const blocks = parseHtml(html, { collapse: !html.includes('data-mo-docs="1"') });
        if (blocks.length) { app.editor.insertBlocks(blocks); return; }
      }
      const image = item.types.find(x => /^image\/(png|jpeg|gif|webp)$/.test(x));
      if (!plain && image) { const blob = await item.getType(image); await app.imagesUi.insertFiles([new File([blob], 'picture', { type: image })]); return; }
      if (item.types.includes('text/plain')) { app.editor.insertPlainText(await (await item.getType('text/plain')).text()); return; }
    }
  } catch {
    app.ctx.toast(t(plain ? 'Docs.Paste.HintPlain' : 'Docs.Paste.Hint'), 'info');
  }
};

// ---- the rest of the contract ------------------------------------------------------------------------------------------------------

function setContent(content, info) {
  app.hf?.close(false);
  load(content);
  app.editor.clearHistory();
  if (info?.remote) app.ctx.ribbon.refresh();
  app.view.update();
  updateStatus();
}

function print() {
  app.hf?.close();
  app.imagesUi?.clear();
  let style = document.getElementById('mo-docs-print');
  if (!style) { style = document.createElement('style'); style.id = 'mo-docs-print'; document.head.append(style); }
  style.textContent = app.view.printCss(app.headers);
  setTimeout(() => window.print(), 30);
}

function info() {
  const texts = app.editor.textBlocks(app.body).map(b => runsText(readBlock(b).runs, ''));
  const words = countWords(texts.join('\n'));
  const chars = texts.join('').replace(/\n/g, '').length;
  return el('dl', { class: 'mat-kv' },
    el('dt', { text: t('Docs.Info.Pages') }), el('dd', { text: String(app.view.pages) }),
    el('dt', { text: t('Docs.Info.Words') }), el('dd', { text: words.toLocaleString(app.language()) }),
    el('dt', { text: t('Docs.Info.Characters') }), el('dd', { text: chars.toLocaleString(app.language()) }),
    el('dt', { text: t('Docs.Info.Language') }), el('dd', { text: languageName(app.settings.language) }));
}

startEditor({
  mount,
  getContent,
  setContent,
  thumbnail: () => thumbnail(app),
  print,
  info,
  downloads: () => [{ ext: 'PDF', label: t('Docs.Download.Pdf'), run: () => { app.ctx.backstage?.('info'); document.getElementById('mo-backstage').hidden = true; document.body.classList.remove('mo-backstage-open'); print(); } }],
});

export { app, readBlocks, cleanText };
