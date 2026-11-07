// MatOffice Sheets – a small stand-in for the browser's DOM: enough to build the editor's elements in node and drive its text fields
// (value, caret, setRangeText). Import it before the modules under test.
class FakeNode {
  constructor(tag = '#text') {
    this.tagName = String(tag).toUpperCase();
    this.attrs = {};
    this.children = [];
    this.listeners = {};
    this.style = {};
    this.dataset = {};
    this.textContent = '';
    this.hidden = false;
    this.disabled = false;
    this.readOnly = false;
    this.value = '';
    this.selectionStart = 0;
    this.selectionEnd = 0;
    const classes = new Set();
    this.classList = { add: c => classes.add(c), remove: c => classes.delete(c), toggle: (c, on) => { if (on ?? !classes.has(c)) classes.add(c); else classes.delete(c); }, contains: c => classes.has(c) };
  }
  setAttribute(k, v) { this.attrs[k] = String(v); }
  getAttribute(k) { return this.attrs[k] ?? null; }
  removeAttribute(k) { delete this.attrs[k]; }
  addEventListener(type, fn) { (this.listeners[type] ||= []).push(fn); }
  removeEventListener() {}
  dispatch(type, event = {}) { for (const fn of this.listeners[type] || []) fn({ preventDefault() {}, stopPropagation() {}, target: this, ...event }); }
  append(...nodes) { this.children.push(...nodes); }
  appendChild(node) { this.children.push(node); return node; }
  replaceChildren(...nodes) { this.children = nodes; }
  replaceWith() {}
  remove() {}
  contains(node) { return node === this || this.children.some(c => c.contains?.(node)); }
  closest() { return null; }
  querySelector() { return null; }
  querySelectorAll() { return []; }
  get firstChild() { return this.children[0] || null; }
  get offsetHeight() { return 20; }
  getBoundingClientRect() { return { left: 0, top: 0, right: 100, bottom: 20, width: 100, height: 20 }; }
  focus() { globalThis.document.activeElement = this; }
  blur() {}
  select() { this.selectionStart = 0; this.selectionEnd = this.value.length; }
  setSelectionRange(s, e) { this.selectionStart = s; this.selectionEnd = e; }
  setRangeText(text, start, end, mode) {
    this.value = this.value.slice(0, start) + text + this.value.slice(end);
    const caret = mode === 'end' ? start + text.length : start;
    this.selectionStart = this.selectionEnd = caret;
  }
}

globalThis.Node = FakeNode;
globalThis.document = {
  activeElement: null,
  createElement: tag => new FakeNode(tag),
  createElementNS: (ns, tag) => new FakeNode(tag),
  createTextNode: text => { const n = new FakeNode('#text'); n.textContent = text; return n; },
  getElementById: () => null,
  addEventListener: () => {},
  documentElement: new FakeNode('html'),
};
globalThis.document.body = new FakeNode('body');

export { FakeNode };
