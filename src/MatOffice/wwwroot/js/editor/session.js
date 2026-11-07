/* MatOffice – the file that is open in an editor: loading, saving by itself, the work of others and a missing network.
   - The editor tells the session that something changed (changed()); the session saves the whole content a moment later (PUT {api} with the revision it knows).
   - 409: somebody else saved in between -> the person chooses: take theirs (the editor shows it) or keep mine (saved over it).
   - 412 (a file of a host): the file was changed at the host -> reload it from there, or download the own copy.
   - No network: the content waits on this device (localStorage mo.pending.<api>) and is sent as soon as the network is back.
   - Every 20 s the session asks for the revision; a newer one of somebody else is loaded while nothing waits to be saved. */
import { t, request, debounce, confirm, dialog, el, download } from './util.js';

const SAVE_DELAY = 1200;
const POLL_EVERY = 20000;

export class EditorSession extends EventTarget {
  /** config: { api, mode, canEdit } from the page; app: { getContent(), setContent(content, info) }. */
  constructor(config, app) {
    super();
    this.api = config.api;
    this.app = app;
    this.doc = null;
    this.revision = 0;
    this.dirty = false;
    this.saving = false;
    this.state = 'idle';
    this.message = '';
    this.pendingKey = 'mo.pending.' + this.api.split('?')[0];
    this.scheduleSave = debounce(() => this.save(), SAVE_DELAY);
    this.pollTimer = 0;
  }

  get canEdit() { return !!this.doc?.canEdit; }

  /** Loads the file; answers the JSON of the editor (or throws with .status). */
  async load({ reload = false } = {}) {
    const url = reload && this.doc?.mode === 'host' ? this.api + '?reload=true' : this.api;
    this.doc = await request('GET', url);
    this.revision = this.doc.revision;
    this.restorePending();
    if (this.canEdit) this.startPolling();
    window.addEventListener('online', () => { if (this.dirty) this.save(); });
    window.addEventListener('beforeunload', e => { if (this.dirty || this.saving) { this.saveNowBeacon(); e.preventDefault(); e.returnValue = ''; } });
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && this.dirty) this.save(); });
    return this.doc;
  }

  /** The editor changed something: save soon. */
  changed() {
    if (!this.canEdit) return;
    this.dirty = true;
    this.set('dirty');
    this.scheduleSave();
  }

  /** Saves now (Ctrl+S). */
  async flush() {
    this.scheduleSave.cancel();
    await this.save();
  }

  set(state, message = '') {
    this.state = state;
    this.message = message;
    this.dispatchEvent(new CustomEvent('state', { detail: { state, message } }));
  }

  async save() {
    if (!this.canEdit || !this.dirty) return;
    if (this.saving) { this.again = true; return; }
    this.saving = true;
    this.again = false;
    const content = this.app.getContent();
    this.dirty = false;
    this.set('saving');
    try {
      const answer = await request('PUT', this.api, { revision: this.revision, content });
      this.revision = answer.revision;
      this.clearPending();
      this.set('saved');
      this.dispatchEvent(new CustomEvent('saved', { detail: answer }));
    } catch (e) {
      await this.failed(e, content);
    } finally {
      this.saving = false;
      if (this.again || (this.dirty && this.state !== 'error' && this.state !== 'offline')) this.scheduleSave();
    }
  }

  async failed(e, content) {
    if (e.status === 0) {
      this.dirty = true;
      this.keepPending(content);
      this.set('offline', t('Editor.Status.Offline'));
      return;
    }

    if (e.status === 409 && e.body) {
      // somebody else saved meanwhile
      const choice = await dialog({
        title: t('Editor.Conflict.Title'), icon: 'users',
        body: el('p', { text: t('Editor.Conflict.Text') }),
        buttons: [{ label: t('Editor.Conflict.Theirs'), value: 'theirs', kind: 'primary' }, { label: t('Editor.Conflict.Mine'), value: 'mine', kind: 'secondary' }],
      });
      if (choice === 'mine') {
        this.revision = e.body.revision;
        this.dirty = true;
        this.scheduleSave.flush();
      } else {
        this.revision = e.body.revision;
        this.dirty = false;
        this.clearPending();
        this.app.setContent(e.body.content, { remote: true });
        this.set('saved');
      }
      return;
    }

    if (e.status === 412) {
      // the file was changed at the host (MatDrive) since it was opened
      this.revision = e.body?.revision ?? this.revision;
      const choice = await dialog({
        title: t('Editor.HostChanged.Title'), icon: 'alert',
        body: el('p', { text: e.body?.error || t('Editor.HostChanged.Text') }),
        buttons: [{ label: t('Editor.HostChanged.Reload'), value: 'reload', kind: 'primary' }, { label: t('Editor.HostChanged.Keep'), value: 'keep', kind: 'secondary' }],
      });
      if (choice === 'reload') {
        await this.reloadFromHost();
      } else {
        this.dispatchEvent(new CustomEvent('download-copy'));
        this.set('error', e.body?.error || t('Editor.HostChanged.Text'));
      }
      return;
    }

    if (e.status === 502) {
      // the working copy is safe in MatOffice; the host did not take the file
      this.revision = e.body?.revision ?? this.revision;
      this.set('error', e.body?.error || t('Editor.Status.HostFailed'));
      this.dirty = true;
      setTimeout(() => { if (this.dirty) this.save(); }, 15000);
      return;
    }

    if (e.status === 422) {
      this.set('error', e.body?.error || t('Editor.Status.Invalid'));
      return;
    }

    if (e.status === 401 || e.status === 403 || e.status === 404) {
      this.set('error', t('Editor.Status.Gone'));
      this.stopPolling();
      return;
    }

    this.dirty = true;
    this.set('error', t('Editor.Status.Failed'));
    setTimeout(() => { if (this.dirty) this.save(); }, 10000);
  }

  async reloadFromHost() {
    const doc = await request('GET', this.api + '?reload=true');
    this.doc = doc;
    this.revision = doc.revision;
    this.dirty = false;
    this.clearPending();
    this.app.setContent(doc.content, { remote: true });
    this.set('saved');
  }

  startPolling() {
    this.stopPolling();
    this.pollTimer = setInterval(() => this.poll(), POLL_EVERY);
  }

  stopPolling() { clearInterval(this.pollTimer); }

  async poll() {
    if (document.visibilityState === 'hidden' || this.saving || this.dirty || navigator.onLine === false) return;
    try {
      const r = await request('GET', this.api.replace(/(\?|$)/, '/revision$1'));
      if (r.revision > this.revision) {
        const doc = await request('GET', this.api);
        if (this.dirty || this.saving) return;
        this.revision = doc.revision;
        this.doc.name = doc.name;
        this.app.setContent(doc.content, { remote: true, by: r.by });
        this.dispatchEvent(new CustomEvent('remote', { detail: { by: r.by } }));
      }
    } catch { /* the next poll tries again */ }
  }

  keepPending(content) {
    try {
      localStorage.setItem(this.pendingKey, JSON.stringify({ revision: this.revision, at: Date.now(), content }));
    } catch {
      this.set('error', t('Editor.Status.TooLargeOffline'));
    }
  }

  clearPending() {
    try { localStorage.removeItem(this.pendingKey); } catch { /* private mode */ }
  }

  /** Changes that waited on this device (the network was gone) are taken over when they were made on the same revision. */
  restorePending() {
    let pending = null;
    try { pending = JSON.parse(localStorage.getItem(this.pendingKey) || 'null'); } catch { pending = null; }
    if (!pending || !this.canEdit) return;
    if (pending.revision === this.revision) {
      this.doc.content = pending.content;
      this.dirty = true;
      setTimeout(() => this.save(), 500);
    } else {
      // made on an older revision: offer the copy as a download instead of overwriting the newer work
      this.clearPending();
      confirm(t('Editor.Pending.Old'), { ok: t('Editor.Pending.Download') }).then(ok => {
        if (ok) download((this.doc.name || 'document') + '.json', JSON.stringify(pending.content), 'application/json');
      });
    }
  }

  /** The last chance when the page is closed: the browser sends the content even when the page is gone (only when it is small enough for a beacon). */
  saveNowBeacon() {
    try { this.keepPending(this.app.getContent()); } catch { /* nothing */ }
  }
}
