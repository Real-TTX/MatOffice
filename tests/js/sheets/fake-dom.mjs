// MatOffice Sheets – a tiny stand-in for the document of the browser, enough for modules that build SVG or import the frame (import it first).
export class FakeNode {
  constructor(tag) { this.tag = tag; this.attrs = {}; this.children = []; this.textContent = ''; }
  setAttribute(k, v) { this.attrs[k] = v; }
  append(...nodes) { this.children.push(...nodes); }
  insertBefore(node) { this.children.unshift(node); }
  count(tag) { return (this.tag === tag ? 1 : 0) + this.children.reduce((n, c) => n + c.count(tag), 0); }
}

globalThis.document = {
  createElementNS: (ns, tag) => new FakeNode(tag),
  getElementById: () => null,
  addEventListener: () => {},
};
