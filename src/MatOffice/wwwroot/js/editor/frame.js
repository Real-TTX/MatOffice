/* MatOffice – the frame every editor shares, so the four apps look and behave like one suite (and like the office suites people know):
   - the title bar: back, the icon of the app, the name (rename), star, the state of saving, read-only / version / host badges, Share, the person
   - the ribbon: tabs (File first) and groups of controls, built from a declarative list the app gives (ctx.ribbon.set), collapsible
   - "File" opens the backstage: info, new, open, save a copy, download as, print, share, version history, close
   - the status bar: the app's texts on the left, zoom on the right
   - saving by itself (session.js), Ctrl+S / Ctrl+P, the picture of the list after a save, a banner for a version that is only looked at

   An app (wwwroot/js/<slug>/app.js) calls startEditor({ mount(ctx), getContent(), setContent(content, info), thumbnail?(), print?(), setZoom?(f) }).
   ctx: { doc, readOnly, work, ribbon, status, session, changed(), t, ... } – see the README of the editors (docs/EDITORS.md). */
import { t, el, icon, request, toast, download, dialog, keyLabel, mod, stamp, token } from './util.js';
import { EditorSession } from './session.js';

const root = document.getElementById('mo-editor');
const config = root ? { ...root.dataset } : {};

/** Starts the editor of the page. */
export async function startEditor(app) {
  const session = new EditorSession({ api: config.api }, {
    getContent: () => app.getContent(),
    setContent: (content, info) => { app.setContent(content, info); ctx.ribbon.refresh(); },
  });

  let doc;
  try {
    doc = await session.load();
  } catch (e) {
    showFailure(e);
    return null;
  }

  const readOnly = !doc.canEdit;
  document.body.classList.toggle('mo-readonly', readOnly);
  const work = document.getElementById('mo-work');
  const ribbon = new Ribbon(document.getElementById('mo-ribbon'), () => ctx);
  const status = new StatusBar(document.getElementById('mo-statusbar'));
  const ctx = {
    doc, readOnly, work, ribbon, status, session, config,
    /** The editor changed the content: it is saved a moment later. */
    changed: () => { if (!readOnly) session.changed(); },
    /** Opens the backstage (File) at a section: info, new, open, download, print, share. */
    backstage: section => backstage.open(section),
    download: (name, data, type) => download(name, data, type),
    t, el, icon, request, toast, dialog,
  };

  titleBar(doc, session, readOnly);
  const backstage = new Backstage(doc, app, session);
  ribbon.fileTab = () => backstage.open('info');
  versionBanner(doc);
  hostBanner(doc);

  // keys of every editor
  document.addEventListener('keydown', e => {
    if (mod(e) && !e.altKey && (e.key === 's' || e.key === 'S')) { e.preventDefault(); session.flush(); }
    else if (mod(e) && !e.altKey && (e.key === 'p' || e.key === 'P')) { e.preventDefault(); print(app); }
    else if (e.key === 'F10' && !e.shiftKey) { e.preventDefault(); ribbon.focus(); }
  });

  // the picture of the list: after a save, at most every 30 s
  let lastThumb = 0;
  session.addEventListener('saved', async () => {
    if (doc.mode !== 'document' || !app.thumbnail || Date.now() - lastThumb < 30000) return;
    lastThumb = Date.now();
    try {
      const png = await app.thumbnail();
      if (png) await request('PUT', `/live/documents/${doc.id}/thumbnail`, png);
    } catch { /* the list keeps the old picture */ }
  });
  session.addEventListener('remote', e => toast(e.detail.by ? t('Editor.Remote.By', e.detail.by) : t('Editor.Remote'), 'info'));
  session.addEventListener('download-copy', () => downloadNative(doc, app));

  await app.mount(ctx);
  ribbon.refresh();
  document.body.classList.add('mo-ready');
  if (doc.mode === 'document' && doc.canEdit && app.thumbnail) {
    // a document without a picture yet gets one soon after it was opened
    setTimeout(async () => { try { const png = await app.thumbnail(); if (png) await request('PUT', `/live/documents/${doc.id}/thumbnail`, png); } catch { /* later */ } }, 4000);
  }
  return ctx;
}

function print(app) {
  if (app.print) app.print();
  else window.print();
}

function downloadNative(doc, app) {
  const native = doc.formats[doc.formats.length - 1];
  const file = { format: 'matoffice', kind: doc.kind, version: 1, name: doc.name, content: app.getContent() };
  download(doc.name + native.ext, JSON.stringify(file), 'application/json');
}

function showFailure(e) {
  const work = document.getElementById('mo-work');
  const gone = e.status === 404 || e.status === 401 || e.status === 403;
  document.body.classList.add('mo-failed');
  work.replaceChildren(el('div', { class: 'mat-empty mo-failure' },
    el('span', { class: 'mat-empty__icon' }, icon(gone ? 'lock' : 'cloud-off')),
    el('div', { class: 'mat-empty__title', text: t(gone ? 'Editor.Gone.Title' : 'Editor.Fail.Title') }),
    el('p', { text: e.body?.error || t(gone ? 'Editor.Gone.Text' : 'Editor.Fail.Text') }),
    el('div', { class: 'mat-btn-row mat-btn-row--plain' },
      el('a', { class: 'mat-btn mat-btn--primary', href: location.href }, icon('refresh'), el('span', { text: t('Editor.Retry') })),
      el('a', { class: 'mat-btn mat-btn--secondary', href: config.back || '/home' }, icon('back'), el('span', { text: t('Editor.Back') })))));
}

// ---- title bar ---------------------------------------------------------------------------------------------------------------

function titleBar(doc, session, readOnly) {
  const name = document.getElementById('mo-name');
  const state = document.getElementById('mo-state');
  const star = document.getElementById('mo-star');
  const badge = document.getElementById('mo-badge');
  const back = document.getElementById('mo-back');

  if (doc.mode === 'host' && doc.returnUrl) back.href = doc.returnUrl;
  name.value = doc.mode === 'host' ? doc.fileName : doc.name;
  document.title = (doc.mode === 'host' ? doc.fileName : doc.name) + ' · ' + t('Kinds.' + doc.kind);
  const renamable = doc.mode === 'document' && doc.role === 'owner' || (doc.mode === 'document' && doc.canEdit);
  name.readOnly = !renamable;
  if (renamable) {
    let before = name.value;
    const commit = async () => {
      const value = name.value.trim();
      if (!value || value === before) { name.value = before; return; }
      try {
        const r = await request('PATCH', doc.api.split('?')[0], { name: value });
        before = name.value = r.name;
        doc.name = r.name;
        document.title = r.name + ' · ' + t('Kinds.' + doc.kind);
      } catch (e) {
        name.value = before;
        toast(e.body?.error || t('Editor.Status.Failed'), 'danger');
      }
    };
    name.addEventListener('change', commit);
    name.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); name.blur(); } else if (e.key === 'Escape') { name.value = before; name.blur(); } });
  }

  if (star) {
    star.hidden = !(doc.mode === 'document' && doc.signedIn && !doc.versionId);
    const paint = () => {
      star.setAttribute('aria-pressed', doc.starred ? 'true' : 'false');
      star.replaceChildren(icon(doc.starred ? 'star-filled' : 'star'));
      star.setAttribute('aria-label', t(doc.starred ? 'Editor.Unstar' : 'Editor.Star'));
      star.dataset.tip = t(doc.starred ? 'Editor.Unstar' : 'Editor.Star');
    };
    paint();
    star.addEventListener('click', async () => {
      doc.starred = !doc.starred;
      paint();
      try { await request('POST', `/live/documents/${doc.id}/star`, { starred: doc.starred }); } catch { doc.starred = !doc.starred; paint(); }
    });
  }

  const labels = { versionId: doc.versionId ? t('Editor.Badge.Version', doc.versionName || '') : null };
  const text = doc.mode === 'version' ? labels.versionId : doc.mode === 'host' ? (doc.host || t('Editor.Badge.Host')) + (readOnly ? ' · ' + t('Editor.Badge.ReadOnly') : '') : readOnly ? t('Editor.Badge.ReadOnly') : null;
  if (text) { badge.hidden = false; badge.textContent = text; }

  const paintState = ({ state: s, message }) => {
    const map = {
      idle: ['cloud-check', t('Editor.Status.Saved')], saved: ['cloud-check', t('Editor.Status.Saved')], dirty: ['edit', t('Editor.Status.Dirty')],
      saving: ['refresh', t('Editor.Status.Saving')], offline: ['cloud-off', message || t('Editor.Status.Offline')], error: ['alert', message || t('Editor.Status.Failed')],
    };
    const [ic, label] = map[s] || map.idle;
    state.replaceChildren(icon(ic), el('span', { class: 'mo-state__text', text: label }));
    state.dataset.state = s;
    state.title = label;
  };
  if (readOnly) state.hidden = true;
  else paintState({ state: doc.lastError ? 'error' : 'idle', message: doc.lastError });
  session.addEventListener('state', e => paintState(e.detail));
}

function versionBanner(doc) {
  if (doc.mode !== 'version') return;
  const bar = document.getElementById('mo-banner');
  bar.hidden = false;
  bar.replaceChildren(icon('history'), el('span', { text: t('Editor.VersionBanner', doc.versionName || '') }),
    el('a', { class: 'mat-btn mat-btn--sm mat-btn--secondary', href: `/Documents/Versions?id=${doc.id}` }, el('span', { text: t('Editor.Versions') })),
    el('a', { class: 'mat-btn mat-btn--sm mat-btn--primary', href: location.pathname }, el('span', { text: t('Editor.BackToDocument') })));
}

function hostBanner(doc) {
  if (doc.mode !== 'host' || !doc.lastError) return;
  const bar = document.getElementById('mo-banner');
  bar.hidden = false;
  bar.replaceChildren(icon('alert'), el('span', { text: doc.lastError }));
}

// ---- ribbon --------------------------------------------------------------------------------------------------------------------

/**
 * The ribbon: ctx.ribbon.set([{ id, label, groups: [{ label, items: [...] }] }]) – the File tab comes first by itself.
 * Items (all take id, label, icon, tip, key (e.g. 'Mod+B'), edit (default true: disabled when read-only), disabled(), hidden()):
 *   { type: 'button', size: 'large'|'small', showLabel, run() }      { type: 'toggle', active(), run() }
 *   { type: 'select', options: [{ value, label }] | () => [...], value(), change(v), width }
 *   { type: 'combo', options: [...], value(), change(v), width }  (free text with suggestions, e.g. a font size)
 *   { type: 'number', value(), change(n), min, max, step, width }
 *   { type: 'color', value(), change(hex|null), auto: label of "Automatic" }   (a split button: applies the last colour, the arrow opens the palette)
 *   { type: 'menu', size, items: [{ label, icon, key, run(), checked(), disabled(), danger } | '-' | { heading }] | () => [...] }
 *   { type: 'split', size, run(), items: [...] }                       { type: 'gallery', items: [{ label, render(): Node, run(), active() }] }
 *   { type: 'stack', rows: [[items], [items]] }  (small controls in rows, the way the office ribbons do it)
 *   { type: 'custom', render(): Node, refresh?(node) }
 */
class Ribbon {
  constructor(host, getCtx) {
    this.host = host;
    this.getCtx = getCtx;
    this.tabs = [];
    this.active = null;
    this.updaters = [];
    this.collapsed = localStorage.getItem('mo.ribbon.collapsed') === '1';
    this.fileTab = () => {};
    this.keys = new Map();
    document.addEventListener('keydown', e => this.onKey(e));
  }

  set(tabs) {
    this.tabs = tabs;
    if (!this.active || !tabs.some(x => x.id === this.active)) this.active = tabs[0]?.id;
    this.keys.clear();
    for (const tab of tabs) for (const g of tab.groups) this.collectKeys(g.items);
    this.render();
  }

  select(id) {
    this.active = id;
    if (this.collapsed) this.host.classList.add('is-peek');
    this.render();
  }

  focus() { this.host.querySelector('.mo-ribbon__tab.is-active, .mo-ribbon__tab')?.focus(); }

  get readOnly() { return this.getCtx()?.readOnly; }

  collectKeys(items) {
    for (const item of items || []) {
      if (item.type === 'stack') item.rows.forEach(r => this.collectKeys(r));
      else if (item.key && (item.run || item.type === 'toggle')) this.keys.set(normalizeKey(item.key), item);
    }
  }

  onKey(e) {
    if (!this.keys.size || e.defaultPrevented) return;
    const combo = comboOf(e);
    const item = this.keys.get(combo);
    if (!item || (item.edit !== false && this.readOnly) || item.disabled?.()) return;
    if (item.global === false && isTyping(e.target)) return;
    e.preventDefault();
    item.run();
    this.refresh();
  }

  render() {
    this.updaters = [];
    const tabsBar = el('div', { class: 'mo-ribbon__tabs', role: 'tablist', 'aria-label': t('Editor.Ribbon') });
    tabsBar.append(el('button', { type: 'button', class: 'mo-ribbon__tab mo-ribbon__tab--file', role: 'tab', 'aria-selected': 'false', onclick: () => this.fileTab() }, el('span', { text: t('Editor.Tab.File') })));
    for (const tab of this.tabs) {
      if (tab.hidden?.()) continue;
      const active = tab.id === this.active;
      tabsBar.append(el('button', {
        type: 'button', class: 'mo-ribbon__tab' + (active ? ' is-active' : ''), role: 'tab', 'aria-selected': active ? 'true' : 'false', id: 'mo-tab-' + tab.id,
        'aria-controls': 'mo-panel',
        onclick: () => this.select(tab.id),
        ondblclick: () => this.toggleCollapsed(),
      }, el('span', { text: tab.label })));
    }
    const collapse = el('button', {
      type: 'button', class: 'mat-btn mat-btn--ghost mat-btn--icon mat-btn--sm mo-ribbon__collapse', 'aria-label': t(this.collapsed ? 'Editor.Ribbon.Expand' : 'Editor.Ribbon.Collapse'),
      'data-tip': t(this.collapsed ? 'Editor.Ribbon.Expand' : 'Editor.Ribbon.Collapse'), onclick: () => this.toggleCollapsed(),
    }, icon(this.collapsed ? 'ribbon-expand' : 'ribbon-collapse'));
    tabsBar.append(el('span', { class: 'mo-ribbon__spacer' }), collapse);

    const panel = el('div', { class: 'mo-ribbon__panel', id: 'mo-panel', role: 'tabpanel', 'aria-labelledby': 'mo-tab-' + this.active });
    const tab = this.tabs.find(x => x.id === this.active);
    for (const group of tab?.groups || []) {
      if (group.hidden?.()) continue;
      const items = el('div', { class: 'mo-rgroup__items' });
      for (const item of group.items) this.renderItem(item, items);
      panel.append(el('div', { class: 'mo-rgroup' + (group.wide ? ' mo-rgroup--wide' : '') }, items, el('div', { class: 'mo-rgroup__label', text: group.label || '' })));
    }

    this.host.classList.toggle('is-collapsed', this.collapsed);
    this.host.replaceChildren(tabsBar, panel);
    this.refresh();
  }

  toggleCollapsed() {
    this.collapsed = !this.collapsed;
    this.host.classList.remove('is-peek');
    try { localStorage.setItem('mo.ribbon.collapsed', this.collapsed ? '1' : '0'); } catch { /* private mode */ }
    this.render();
  }

  /** Re-reads every state (active, disabled, value) – call it after the selection changed. */
  refresh() {
    for (const update of this.updaters) {
      try { update(); } catch { /* one control must not stop the others */ }
    }
  }

  canRun(item) { return !(item.edit !== false && this.readOnly) && !item.disabled?.(); }

  tipOf(item) { return (item.tip || item.label) + (item.key ? ` (${keyLabel(item.key)})` : ''); }

  renderItem(item, into) {
    if (item.hidden?.()) return;
    const make = {
      button: () => this.button(item), toggle: () => this.button(item, true), select: () => this.select_(item), combo: () => this.combo(item),
      number: () => this.number(item), color: () => this.color(item), menu: () => this.menu(item), split: () => this.split(item),
      gallery: () => this.gallery(item), stack: () => this.stack(item), custom: () => this.custom(item), separator: () => el('span', { class: 'mo-rsep' }),
    }[item.type || 'button'];
    const node = make();
    if (node) into.append(node);
  }

  button(item, toggle = false) {
    const large = item.size === 'large';
    const showLabel = item.showLabel ?? large;
    const b = el('button', {
      type: 'button', class: 'mo-rbtn' + (large ? ' mo-rbtn--lg' : '') + (showLabel ? '' : ' mo-rbtn--icon'),
      'aria-label': showLabel ? null : item.label, 'data-tip': showLabel && !item.key ? null : this.tipOf(item), 'data-ribbon-id': item.id || null,
      onclick: () => { if (this.canRun(item)) { item.run?.(); this.refresh(); } },
    }, item.icon ? icon(item.icon, large ? 'mo-rbtn__icon--lg' : '') : null, showLabel ? el('span', { class: 'mo-rbtn__label', text: item.label }) : null);
    this.updaters.push(() => {
      b.disabled = !this.canRun(item);
      if (toggle) b.setAttribute('aria-pressed', item.active?.() ? 'true' : 'false');
    });
    return b;
  }

  select_(item) {
    const s = el('select', { class: 'mat-input mo-rselect', 'aria-label': item.label, 'data-tip': item.tip || item.label, 'data-ribbon-id': item.id || null });
    if (item.width) s.style.width = item.width + 'px';
    const fill = () => {
      const options = typeof item.options === 'function' ? item.options() : item.options;
      s.replaceChildren(...options.map(o => el('option', { value: o.value, text: o.label })));
    };
    fill();
    s.addEventListener('change', () => { if (this.canRun(item)) { item.change(s.value); this.refresh(); } });
    s.addEventListener('mousedown', () => { if (typeof item.options === 'function') { const v = s.value; fill(); s.value = v; } });
    this.updaters.push(() => {
      s.disabled = !this.canRun(item);
      const v = item.value?.();
      if (v !== undefined && document.activeElement !== s) {
        if (![...s.options].some(o => o.value === String(v))) s.append(el('option', { value: String(v), text: String(v) }));
        s.value = String(v ?? '');
      }
    });
    return s;
  }

  combo(item) {
    const listId = 'mo-dl-' + Math.random().toString(36).slice(2, 8);
    const options = typeof item.options === 'function' ? item.options() : item.options;
    const list = el('datalist', { id: listId }, ...options.map(o => el('option', { value: o.value ?? o })));
    const input = el('input', { class: 'mat-input mo-rcombo', list: listId, 'aria-label': item.label, 'data-tip': item.tip || item.label, autocomplete: 'off', inputmode: item.inputmode || null });
    if (item.width) input.style.width = item.width + 'px';
    const apply = () => { if (this.canRun(item) && input.value.trim() !== '') { item.change(input.value.trim()); this.refresh(); } };
    input.addEventListener('change', apply);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); apply(); this.getCtx()?.work?.focus?.(); } });
    this.updaters.push(() => {
      input.disabled = !this.canRun(item);
      if (document.activeElement !== input) input.value = item.value?.() ?? '';
    });
    return el('span', { class: 'mo-rcombo-wrap' }, input, list);
  }

  number(item) {
    const input = el('input', { type: 'number', class: 'mat-input mo-rnumber', 'aria-label': item.label, 'data-tip': item.tip || item.label, min: item.min, max: item.max, step: item.step || 1 });
    if (item.width) input.style.width = item.width + 'px';
    input.addEventListener('change', () => { const n = Number(input.value); if (this.canRun(item) && Number.isFinite(n)) { item.change(n); this.refresh(); } });
    this.updaters.push(() => {
      input.disabled = !this.canRun(item);
      if (document.activeElement !== input) input.value = item.value?.() ?? '';
    });
    return input;
  }

  color(item) {
    let last = item.initial || '#c00000';
    const swatch = el('span', { class: 'mo-rcolor__bar' });
    const main = el('button', {
      type: 'button', class: 'mo-rbtn mo-rbtn--icon mo-rcolor__main', 'aria-label': item.label, 'data-tip': this.tipOf(item),
      onclick: () => { if (this.canRun(item)) { item.change(last); this.refresh(); } },
    }, item.icon ? icon(item.icon) : null, swatch);
    const arrow = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--arrow', 'aria-label': t('Editor.MoreColours', item.label), 'aria-haspopup': 'true' }, icon('chevron-down'));
    arrow.addEventListener('click', () => {
      if (!this.canRun(item)) return;
      openPopover(arrow, colourPalette(item.auto, c => { if (c) last = c; swatch.style.background = last; item.change(c); this.refresh(); }));
    });
    swatch.style.background = last;
    this.updaters.push(() => { main.disabled = arrow.disabled = !this.canRun(item); });
    return el('span', { class: 'mo-rsplit mo-rcolor' }, main, arrow);
  }

  menuItems(items, close) {
    const list = el('div', { class: 'mo-menu', role: 'menu' });
    const entries = typeof items === 'function' ? items() : items;
    for (const entry of entries) {
      if (entry === '-') { list.append(el('div', { class: 'mo-menu__sep', role: 'separator' })); continue; }
      if (entry.heading) { list.append(el('div', { class: 'mo-menu__heading', text: entry.heading })); continue; }
      if (entry.hidden?.()) continue;
      const checked = entry.checked?.();
      const disabled = (entry.edit !== false && this.readOnly) || entry.disabled?.();
      list.append(el('button', {
        type: 'button', role: checked === undefined ? 'menuitem' : 'menuitemcheckbox', 'aria-checked': checked === undefined ? null : String(!!checked),
        class: 'mo-menu__item' + (entry.danger ? ' mo-menu__item--danger' : ''), disabled: disabled ? true : null,
        onclick: () => { close(); entry.run?.(); this.refresh(); },
      }, entry.render ? entry.render() : entry.icon ? icon(entry.icon) : el('span', { class: 'mo-menu__noicon' }), el('span', { class: 'mo-menu__label', text: entry.label }),
      entry.key ? el('span', { class: 'mo-menu__key', text: keyLabel(entry.key) }) : null, checked ? icon('check', 'mo-menu__check') : null));
    }
    return list;
  }

  menu(item) {
    const large = item.size === 'large';
    const b = el('button', {
      type: 'button', class: 'mo-rbtn mo-rbtn--menu' + (large ? ' mo-rbtn--lg' : ''), 'aria-haspopup': 'menu', 'data-tip': large ? null : this.tipOf(item),
      'aria-label': item.showLabel === false ? item.label : null,
    }, item.icon ? icon(item.icon, large ? 'mo-rbtn__icon--lg' : '') : null,
    item.showLabel === false ? null : el('span', { class: 'mo-rbtn__label', text: item.label }), icon('chevron-down', 'mo-rbtn__chev'));
    b.addEventListener('click', () => openPopover(b, close => this.menuItems(item.items, close)));
    this.updaters.push(() => { b.disabled = item.edit !== false && this.readOnly ? true : !!item.disabled?.(); });
    return b;
  }

  split(item) {
    const large = item.size === 'large';
    const main = el('button', {
      type: 'button', class: 'mo-rbtn' + (large ? ' mo-rbtn--lg' : '') + (item.showLabel === false ? ' mo-rbtn--icon' : ''), 'data-tip': this.tipOf(item),
      'aria-label': item.showLabel === false ? item.label : null,
      onclick: () => { if (this.canRun(item)) { item.run(); this.refresh(); } },
    }, item.icon ? icon(item.icon, large ? 'mo-rbtn__icon--lg' : '') : null, item.showLabel === false ? null : el('span', { class: 'mo-rbtn__label', text: item.label }));
    const arrow = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--arrow', 'aria-haspopup': 'menu', 'aria-label': t('Editor.MoreOptions', item.label) }, icon('chevron-down'));
    arrow.addEventListener('click', () => openPopover(arrow, close => this.menuItems(item.items, close)));
    this.updaters.push(() => { main.disabled = arrow.disabled = !this.canRun(item); if (item.active) main.setAttribute('aria-pressed', item.active() ? 'true' : 'false'); });
    return el('span', { class: 'mo-rsplit' + (large ? ' mo-rsplit--lg' : '') }, main, arrow);
  }

  gallery(item) {
    const strip = el('div', { class: 'mo-rgallery', role: 'listbox', 'aria-label': item.label });
    const buttons = [];
    for (const entry of item.items) {
      const b = el('button', {
        type: 'button', class: 'mo-rgallery__item', role: 'option', 'data-tip': entry.label, 'aria-label': entry.label,
        onclick: () => { if (this.canRun(item)) { entry.run(); this.refresh(); } },
      }, entry.render ? entry.render() : el('span', { text: entry.label }));
      buttons.push([b, entry]);
      strip.append(b);
    }
    this.updaters.push(() => buttons.forEach(([b, entry]) => { b.disabled = !this.canRun(item); b.setAttribute('aria-selected', entry.active?.() ? 'true' : 'false'); }));
    return strip;
  }

  stack(item) {
    const box = el('div', { class: 'mo-rstack' });
    for (const row of item.rows) {
      const line = el('div', { class: 'mo-rstack__row' });
      for (const child of row) this.renderItem(child, line);
      box.append(line);
    }
    return box;
  }

  custom(item) {
    const node = item.render();
    if (item.refresh) this.updaters.push(() => item.refresh(node, this.readOnly));
    return node;
  }
}

function normalizeKey(combo) {
  return combo.split('+').map(p => p.trim()).map(p => (p.length === 1 ? p.toLowerCase() : p)).sort((a, b) => (a.length === 1) - (b.length === 1) || a.localeCompare(b)).join('+');
}

function comboOf(e) {
  const parts = [];
  if (mod(e)) parts.push('Mod');
  if (e.shiftKey) parts.push('Shift');
  if (e.altKey) parts.push('Alt');
  const key = e.key.length === 1 ? e.key.toLowerCase() : e.key;
  parts.push(key);
  return normalizeKey(parts.join('+'));
}

function isTyping(target) {
  return target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));
}

// ---- popovers ----------------------------------------------------------------------------------------------------------------

let openPop = null;

/** Opens a popover under (or above) an anchor: content is a Node or (close) => Node. Closes on a click outside and on Escape. */
export function openPopover(anchor, content) {
  closePopover();
  const pop = el('div', { class: 'mo-pop' });
  const close = () => closePopover();
  pop.append(typeof content === 'function' ? content(close) : content);
  document.body.append(pop);
  const r = anchor.getBoundingClientRect();
  const w = pop.offsetWidth;
  const h = pop.offsetHeight;
  let left = Math.min(r.left, window.innerWidth - w - 8);
  let top = r.bottom + 4;
  if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 4);
  pop.style.left = Math.max(8, left) + 'px';
  pop.style.top = top + 'px';
  const outside = e => { if (!pop.contains(e.target) && !anchor.contains(e.target)) closePopover(); };
  const esc = e => { if (e.key === 'Escape') { closePopover(); anchor.focus(); } };
  setTimeout(() => { document.addEventListener('pointerdown', outside, true); document.addEventListener('keydown', esc, true); }, 0);
  openPop = { pop, outside, esc };
  pop.querySelector('button, input, select')?.focus();
  return pop;
}

export function closePopover() {
  if (!openPop) return;
  document.removeEventListener('pointerdown', openPop.outside, true);
  document.removeEventListener('keydown', openPop.esc, true);
  openPop.pop.remove();
  openPop = null;
}

/** The colours of the palette: the theme colours of the office suites in shades, and the standard colours. */
export const PALETTE = [
  ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'],
  ['#f2f2f2', '#7f7f7f', '#d0cece', '#d6dce4', '#d9e2f3', '#fbe5d5', '#ededed', '#fff2cc', '#deebf6', '#e2efd9'],
  ['#d8d8d8', '#595959', '#aeabab', '#adb9ca', '#b4c6e7', '#f7cbac', '#dbdbdb', '#fee599', '#bdd7ee', '#c5e0b3'],
  ['#bfbfbf', '#3f3f3f', '#757070', '#8496b0', '#8eaadb', '#f4b183', '#c9c9c9', '#ffd965', '#9cc3e5', '#a8d08d'],
  ['#a5a5a5', '#262626', '#3a3838', '#323f4f', '#2f5496', '#c55a11', '#7b7b7b', '#bf9000', '#2e75b5', '#538135'],
  ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'],
];

/** A colour palette for a popover: pick(hex) or pick(null) for "Automatic". */
export function colourPalette(autoLabel, pick) {
  return close => {
    const box = el('div', { class: 'mo-palette' });
    if (autoLabel) box.append(el('button', { type: 'button', class: 'mo-palette__auto', onclick: () => { close(); pick(null); } }, el('span', { class: 'mo-palette__chip mo-palette__chip--auto' }), el('span', { text: autoLabel })));
    PALETTE.forEach((row, i) => {
      const line = el('div', { class: 'mo-palette__row' + (i === 0 || i === PALETTE.length - 1 ? ' mo-palette__row--gap' : '') });
      for (const c of row) {
        const chip = el('button', { type: 'button', class: 'mo-palette__chip', 'aria-label': c, 'data-tip': c, onclick: () => { close(); pick(c); } });
        chip.style.background = c;
        line.append(chip);
      }
      box.append(line);
    });
    const custom = el('input', { type: 'color', class: 'mo-palette__custom', 'aria-label': t('Editor.CustomColour') });
    custom.addEventListener('change', () => { close(); pick(custom.value); });
    box.append(el('label', { class: 'mo-palette__more' }, icon('palette'), el('span', { text: t('Editor.CustomColour') }), custom));
    return box;
  };
}

// ---- backstage (File) -----------------------------------------------------------------------------------------------------------

class Backstage {
  constructor(doc, app, session) {
    this.doc = doc;
    this.app = app;
    this.session = session;
    this.node = document.getElementById('mo-backstage');
  }

  sections() {
    const d = this.doc;
    const docMode = d.mode === 'document' || d.mode === 'version';
    return [
      { id: 'info', label: t('Editor.File.Info'), icon: 'info' },
      { id: 'new', label: t('Editor.File.New'), icon: 'file-plus', hidden: !docMode || !d.signedIn },
      { id: 'open', label: t('Editor.File.Open'), icon: 'folder-open', hidden: !d.signedIn },
      { id: 'copy', label: t('Editor.File.Copy'), icon: 'copy', hidden: !docMode || !d.signedIn },
      { id: 'download', label: t('Editor.File.Download'), icon: 'download' },
      { id: 'print', label: t('Editor.File.Print'), icon: 'print' },
      { id: 'share', label: t('Editor.File.Share'), icon: 'share', hidden: !docMode || d.role !== 'owner' },
      { id: 'versions', label: t('Editor.File.Versions'), icon: 'history', hidden: !docMode || !d.signedIn },
      { id: 'close', label: t('Editor.File.Close'), icon: 'x' },
    ].filter(s => !s.hidden);
  }

  open(section = 'info') {
    const nav = el('nav', { class: 'mo-backstage__nav', 'aria-label': t('Editor.Tab.File') },
      el('button', { type: 'button', class: 'mo-backstage__back', onclick: () => this.close(), 'aria-label': t('Editor.Back') }, icon('back'), el('span', { text: t('Editor.Back') })));
    const body = el('div', { class: 'mo-backstage__body' });
    for (const s of this.sections()) {
      nav.append(el('button', {
        type: 'button', class: 'mo-backstage__item' + (s.id === section ? ' is-active' : ''), 'aria-current': s.id === section ? 'page' : null,
        onclick: () => this.open(s.id),
      }, icon(s.icon), el('span', { text: s.label })));
    }
    this.fill(section, body);
    this.node.replaceChildren(nav, body);
    this.node.hidden = false;
    document.body.classList.add('mo-backstage-open');
    this.escape ??= e => { if (e.key === 'Escape' && !this.node.hidden) this.close(); };
    document.addEventListener('keydown', this.escape);
    nav.querySelector('.is-active')?.focus();
  }

  close() {
    this.node.hidden = true;
    document.body.classList.remove('mo-backstage-open');
  }

  fill(section, body) {
    const d = this.doc;
    const title = text => el('h1', { class: 'mo-backstage__title', text });
    switch (section) {
      case 'info': {
        const rows = [
          [t('Editor.Info.Name'), d.mode === 'host' ? d.fileName : d.name],
          [t('Editor.Info.Kind'), t('Kinds.' + d.kind)],
          d.owner ? [t('Editor.Info.Owner'), d.owner] : null,
          d.mode === 'host' ? [t('Editor.Info.Host'), d.host] : null,
          [t('Editor.Info.Access'), t('Editor.Role.' + d.role)],
          [t('Editor.Info.Revision'), String(this.session.revision)],
        ].filter(Boolean);
        body.replaceChildren(title(t('Editor.File.Info')), el('dl', { class: 'mat-kv' }, ...rows.flatMap(([k, v]) => [el('dt', { text: k }), el('dd', { text: v })])));
        if (this.app.info) body.append(this.app.info());
        break;
      }
      case 'new': {
        const form = el('form', { method: 'post', action: '/home?handler=Create', class: 'mo-backstage__new' },
          el('input', { type: 'hidden', name: '__RequestVerificationToken', value: token() }),
          el('input', { type: 'hidden', name: 'kind', value: d.kind }),
          el('input', { type: 'hidden', name: 'template', value: 'blank' }),
          el('button', { type: 'submit', class: 'mo-starter__button mo-kind--' + d.kind },
            el('span', { class: 'mo-starter__paper mo-starter__paper--' + d.kind }, icon('plus', 'mo-starter__icon')),
            el('span', { class: 'mo-starter__name', text: t('Editor.New.Blank') })));
        body.replaceChildren(title(t('Editor.File.New')), form, el('p', {}, el('a', { href: '/' + d.kind, text: t('Editor.New.Templates') })));
        break;
      }
      case 'open':
        body.replaceChildren(title(t('Editor.File.Open')), el('p', { text: t('Editor.Open.Text') }),
          el('div', { class: 'mat-btn-row mat-btn-row--plain' },
            el('a', { class: 'mat-btn mat-btn--primary', href: '/' + d.kind }, icon('kind-' + d.kind), el('span', { text: t('Kinds.' + d.kind) })),
            el('a', { class: 'mat-btn mat-btn--secondary', href: '/home' }, icon('home'), el('span', { text: t('Editor.Open.All') }))));
        break;
      case 'copy': {
        const button = el('button', { type: 'button', class: 'mat-btn mat-btn--primary' }, icon('copy'), el('span', { text: t('Editor.Copy.Button') }));
        button.addEventListener('click', async () => {
          await this.session.flush();
          try {
            const r = await request('POST', `/live/documents/${d.id}/duplicate`);
            location.assign(r.url);
          } catch { toast(t('Editor.Status.Failed'), 'danger'); }
        });
        body.replaceChildren(title(t('Editor.File.Copy')), el('p', { text: t('Editor.Copy.Text') }), el('div', { class: 'mat-btn-row mat-btn-row--plain' }, button));
        break;
      }
      case 'download': {
        const list = el('div', { class: 'mo-formats' });
        const base = d.api.split('?')[0];
        const query = d.api.includes('?') ? '&' + d.api.split('?')[1] : '';
        for (const f of d.formats) {
          const a = el('a', { class: 'mo-format', href: `${base}/export?format=${encodeURIComponent(f.ext)}${query}`, download: '' },
            el('span', { class: 'mo-format__ext', text: f.ext.replace('.', '').toUpperCase() }), el('span', { class: 'mo-format__label', text: f.label }));
          a.addEventListener('click', async e => {
            // the server converts what is SAVED: save first
            if (this.session.dirty || this.session.saving) { e.preventDefault(); await this.session.flush(); location.assign(a.href); }
          });
          list.append(a);
        }
        const extra = this.app.downloads?.() || [];
        for (const x of extra) {
          const b = el('button', { type: 'button', class: 'mo-format', onclick: () => x.run() }, el('span', { class: 'mo-format__ext', text: x.ext }), el('span', { class: 'mo-format__label', text: x.label }));
          list.append(b);
        }
        body.replaceChildren(title(t('Editor.File.Download')), el('p', { text: t('Editor.Download.Text') }), list);
        break;
      }
      case 'print':
        body.replaceChildren(title(t('Editor.File.Print')), el('p', { text: t('Editor.Print.Text') }),
          el('div', { class: 'mat-btn-row mat-btn-row--plain' }, el('button', { type: 'button', class: 'mat-btn mat-btn--primary', onclick: () => { this.close(); setTimeout(() => print(this.app), 50); } }, icon('print'), el('span', { text: t('Editor.Print.Button') }))));
        break;
      case 'share':
        body.replaceChildren(title(t('Editor.File.Share')), el('p', { text: t('Editor.Share.Text') }),
          el('div', { class: 'mat-btn-row mat-btn-row--plain' },
            el('a', { class: 'mat-btn mat-btn--primary', href: `/Documents/Members?id=${d.id}` }, icon('users'), el('span', { text: t('Editor.Share.People') })),
            el('a', { class: 'mat-btn mat-btn--secondary', href: `/Documents/Share?id=${d.id}` }, icon('link'), el('span', { text: t('Editor.Share.Link') }))));
        break;
      case 'versions': {
        const list = el('ul', { class: 'mo-versions' }, el('li', { class: 'mo-muted', text: t('Editor.Loading') }));
        body.replaceChildren(title(t('Editor.File.Versions')), list, el('div', { class: 'mat-btn-row mat-btn-row--plain' },
          el('a', { class: 'mat-btn mat-btn--secondary', href: `/Documents/Versions?id=${d.id}` }, icon('history'), el('span', { text: t('Editor.Versions.All') })),
          d.canEdit ? el('a', { class: 'mat-btn mat-btn--secondary', href: `/Documents/Version?documentId=${d.id}` }, icon('plus'), el('span', { text: t('Editor.Versions.Name') })) : null));
        request('GET', `/live/documents/${d.id}/versions`).then(rows => {
          list.replaceChildren(...(rows.length ? rows.slice(0, 15).map(v => el('li', {}, el('a', { href: `${location.pathname}?v=${v.id}` },
            icon(v.isAuto ? 'clock' : 'version'), el('span', { text: v.name || t('Editor.Versions.Auto') }), el('small', { text: `${v.createdBy} · ${stamp(v.createDate)}` })))) : [el('li', { class: 'mo-muted', text: t('Editor.Versions.None') })]));
        }).catch(() => list.replaceChildren(el('li', { text: t('Editor.Status.Failed') })));
        break;
      }
      case 'close':
        this.session.flush().finally(async () => {
          if (d.mode === 'host') {
            try { await request('POST', d.api + '/close'); } catch { /* it expires by itself */ }
            location.assign(d.returnUrl || '/');
          } else {
            location.assign(config.back || '/' + d.kind);
          }
        });
        body.replaceChildren(title(t('Editor.File.Close')), el('p', { text: t('Editor.Closing') }));
        break;
    }
  }
}

// ---- status bar ----------------------------------------------------------------------------------------------------------------

class StatusBar {
  constructor(host) {
    this.host = host;
    this.leftBox = el('div', { class: 'mo-statusbar__left', 'aria-live': 'polite' });
    this.rightBox = el('div', { class: 'mo-statusbar__right' });
    host.replaceChildren(this.leftBox, this.rightBox);
  }

  /** The texts on the left ("Page 2 of 5", "Words: 120", "Sum: 42"). */
  left(...items) { this.leftBox.replaceChildren(...items.filter(x => x !== null && x !== undefined).map(x => (x instanceof Node ? x : el('span', { class: 'mo-statusbar__item', text: String(x) })))); }

  /** Something on the right before the zoom (view buttons). */
  right(...nodes) { this.rightExtra = nodes; this.paintRight(); }

  /** The zoom: { value (1 = 100 %), min, max, change(f), fit?() }. */
  zoom(options) { this.zoomOptions = options; this.paintRight(); }

  setZoom(value) { if (this.zoomOptions) { this.zoomOptions.value = value; this.paintRight(); } }

  paintRight() {
    const parts = [...(this.rightExtra || [])];
    const z = this.zoomOptions;
    if (z) {
      const set = v => { const f = Math.min(z.max ?? 4, Math.max(z.min ?? 0.25, Math.round(v * 100) / 100)); z.value = f; z.change(f); this.paintRight(); };
      const range = el('input', { type: 'range', class: 'mo-zoom__range', min: Math.round((z.min ?? 0.25) * 100), max: Math.round((z.max ?? 4) * 100), step: 5, value: Math.round(z.value * 100), 'aria-label': t('Editor.Zoom') });
      range.addEventListener('input', () => set(Number(range.value) / 100));
      parts.push(el('div', { class: 'mo-zoom' },
        z.fit ? el('button', { type: 'button', class: 'mo-sbtn', 'aria-label': t('Editor.ZoomFit'), 'data-tip': t('Editor.ZoomFit'), onclick: () => z.fit() }, icon('zoom-fit')) : null,
        el('button', { type: 'button', class: 'mo-sbtn', 'aria-label': t('Editor.ZoomOut'), onclick: () => set(z.value - 0.1) }, icon('minus')),
        range,
        el('button', { type: 'button', class: 'mo-sbtn', 'aria-label': t('Editor.ZoomIn'), onclick: () => set(z.value + 0.1) }, icon('plus')),
        el('button', { type: 'button', class: 'mo-zoom__value', 'data-tip': t('Editor.Zoom100'), onclick: () => set(1), text: Math.round(z.value * 100) + ' %' })));
    }
    this.rightBox.replaceChildren(...parts);
  }
}
