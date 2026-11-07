/* MatOffice – the editor of this app (placeholder of the skeleton: shows the content as text until the real editor is built). */
import { startEditor } from '../editor/frame.js';
import { el, t } from '../editor/util.js';

let content = {};
let view;

startEditor({
  mount(ctx) {
    content = ctx.doc.content;
    view = el('pre', { class: 'mo-placeholder' });
    view.textContent = JSON.stringify(content, null, 2);
    ctx.work.replaceChildren(view);
    ctx.ribbon.set([{ id: 'home', label: t('Kinds.docs'), groups: [] }]);
  },
  getContent: () => content,
  setContent(next) { content = next; view.textContent = JSON.stringify(content, null, 2); },
});
