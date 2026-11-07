/* MatOffice Docs – the ribbon (like Word's): Home, Insert, Layout, References, Review, View, and the contextual tabs Table, Picture and
   Header & Footer while a table, a picture or a header is being worked on. Built with the frame's declarative items (docs/EDITORS.md). */
import { t, el, icon } from '../editor/util.js';
import { openPopover } from '../editor/frame.js';
import { FONTS, SIZES, STYLES, HIGHLIGHTS, MARGIN_PRESETS, PAGE_SIZES, TABLE_STYLES, numStr, toPt, fontStack } from './model.js';
import { LINE_SPACINGS, currentLineHeight, currentSpacing } from './commands.js';

const PT_PER_CM = 72 / 2.54;

export function buildRibbon(app) {
  const c = app.commands;
  const ed = app.editor;
  const s = () => ed.state();
  const f = key => !!s().format[key];
  const block = () => s().block;
  const tables = app.tables;
  const images = app.imagesUi;
  const hf = app.hf;
  const inBody = () => { const sel = s().sel; return !!sel && sel.host === app.body; };
  const inCell = () => !!tables.current();

  const styleItems = ['Normal', 'NoSpacing', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Quote', 'IntenseQuote', 'ListParagraph', 'Code'].map(key => ({
    label: t('Docs.Style.' + key), render: () => styleChip(key), run: () => c.style(key), active: () => s().styleKey === key,
  }));

  const bulletItems = [['bullet', '•'], ['circle', '◦'], ['square', '▪'], ['dash', '–'], ['check', '✓'], ['arrow', '➢']].map(([kind, ch]) => ({
    kind, label: t('Docs.List.' + kind), render: () => chip(ch), run: () => c.list(kind), checked: () => block()?.list === kind,
  }));
  const numberItems = [['decimal', '1.'], ['lower-alpha', 'a.'], ['upper-alpha', 'A.'], ['lower-roman', 'i.'], ['upper-roman', 'I.']].map(([kind, ch]) => ({
    kind, label: t('Docs.List.' + kind), render: () => chip(ch), run: () => c.list(kind), checked: () => block()?.list === kind,
  }));

  const home = {
    id: 'home', label: t('Docs.Tab.Home'), groups: [
      { label: t('Docs.Group.Undo'), items: [
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Undo'), icon: 'undo', key: 'Mod+Z', global: false, disabled: () => !ed.canUndo, run: () => { ed.focus(); ed.undo(); } }],
          [{ type: 'button', label: t('Docs.Redo'), icon: 'redo', key: 'Mod+Y', global: false, disabled: () => !ed.canRedo, run: () => { ed.focus(); ed.redo(); } }],
        ] },
      ] },
      { label: t('Docs.Group.Clipboard'), items: [
        { type: 'split', id: 'paste', label: t('Docs.Paste'), icon: 'docs-paste', size: 'large', run: () => app.clipboardPaste(false), items: [
          { label: t('Docs.Paste'), icon: 'docs-paste', key: 'Mod+V', run: () => app.clipboardPaste(false) },
          { label: t('Docs.PastePlain'), icon: 'docs-paste-text', key: 'Mod+Shift+V', run: () => app.clipboardPaste(true) },
        ] },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Cut'), icon: 'docs-cut', showLabel: true, run: () => { ed.focus(); document.execCommand('cut'); } }],
          [{ type: 'button', label: t('Docs.Copy'), icon: 'copy', showLabel: true, edit: false, run: () => { ed.focus(); document.execCommand('copy'); } }],
          [{ type: 'toggle', label: t('Docs.Painter'), icon: 'docs-painter', showLabel: true, tip: t('Docs.Painter.Tip'), active: () => c.painter.active, run: () => c.painterStart(false) }],
        ] },
      ] },
      { label: t('Docs.Group.Font'), items: [
        { type: 'stack', rows: [
          [
            { type: 'select', label: t('Docs.Font'), width: 148, options: () => fontOptions(s().fontName), value: () => s().fontName || '', change: v => c.setFont(v) },
            { type: 'combo', label: t('Docs.Size'), width: 52, inputmode: 'decimal', options: SIZES.map(String), value: () => (s().sizePt ? numStr(s().sizePt) : ''), change: v => c.setSize(v) },
            { type: 'button', label: t('Docs.Grow'), icon: 'docs-grow', key: 'Mod+]', run: () => c.grow(1), global: false },
            { type: 'button', label: t('Docs.Shrink'), icon: 'docs-shrink', key: 'Mod+[', run: () => c.grow(-1), global: false },
            { type: 'menu', label: t('Docs.Case'), icon: 'docs-case', showLabel: false, items: [
              { label: t('Docs.Case.sentence'), run: () => c.changeCase('sentence') },
              { label: t('Docs.Case.lower'), run: () => c.changeCase('lower') },
              { label: t('Docs.Case.upper'), run: () => c.changeCase('upper') },
              { label: t('Docs.Case.title'), run: () => c.changeCase('title') },
              { label: t('Docs.Case.toggle'), run: () => c.changeCase('toggle') },
            ] },
            { type: 'button', label: t('Docs.ClearFormatting'), icon: 'docs-clear-format', run: () => c.clearFormatting() },
          ],
          [
            { type: 'toggle', label: t('Docs.Bold'), icon: 'docs-bold', key: 'Mod+B', global: false, active: () => f('b'), run: () => c.toggle('b') },
            { type: 'toggle', label: t('Docs.Italic'), icon: 'docs-italic', key: 'Mod+I', global: false, active: () => f('i'), run: () => c.toggle('i') },
            { type: 'toggle', label: t('Docs.Underline'), icon: 'docs-underline', key: 'Mod+U', global: false, active: () => f('u'), run: () => c.toggle('u') },
            { type: 'toggle', label: t('Docs.Strike'), icon: 'docs-strike', active: () => f('s'), run: () => c.toggle('s') },
            { type: 'toggle', label: t('Docs.Subscript'), icon: 'docs-subscript', key: 'Mod+=', global: false, active: () => f('sub'), run: () => c.toggle('sub') },
            { type: 'toggle', label: t('Docs.Superscript'), icon: 'docs-superscript', key: 'Mod+Shift++', global: false, active: () => f('sup'), run: () => c.toggle('sup') },
            { type: 'custom', render: () => highlightButton(app), refresh: (node, ro) => node.querySelectorAll('button').forEach(b => { b.disabled = ro; }) },
            { type: 'color', label: t('Docs.FontColour'), icon: 'docs-font-colour', auto: t('Docs.Automatic'), initial: '#c00000', change: v => c.setColor(v) },
          ],
        ] },
      ] },
      { label: t('Docs.Group.Paragraph'), items: [
        { type: 'stack', rows: [
          [
            { type: 'split', label: t('Docs.Bullets'), icon: 'docs-bullets', showLabel: false, key: 'Mod+Shift+L', global: false, active: () => isBullet(block()?.list), run: () => c.list(lastBullet.kind), items: () => bulletItems.map(i => ({ ...i, run: () => { lastBullet.kind = kindOf(i); i.run(); } })) },
            { type: 'split', label: t('Docs.Numbering'), icon: 'docs-numbering', showLabel: false, active: () => isNumber(block()?.list), run: () => c.list(lastNumber.kind), items: () => [
              ...numberItems.map(i => ({ ...i, run: () => { lastNumber.kind = kindOf(i); i.run(); } })), '-',
              { label: t('Docs.List.Restart'), icon: 'refresh', disabled: () => !block()?.list, run: () => c.numberFrom(1) },
              { label: t('Docs.List.Continue'), disabled: () => !block()?.start, run: () => c.numberFrom(null) },
              { label: t('Docs.List.SetValue'), disabled: () => !block()?.list, run: () => app.dialogs.numberValue() },
            ] },
            { type: 'menu', label: t('Docs.Multilevel'), icon: 'docs-multilevel', showLabel: false, items: () => [
              { label: t('Docs.List.outline'), render: () => chip('1.1'), run: () => c.list('outline'), checked: () => block()?.list === 'outline' },
              { label: t('Docs.List.decimalLevels'), render: () => chip('a.'), run: () => c.list('decimal'), checked: () => block()?.list === 'decimal' },
              { label: t('Docs.List.bulletLevels'), render: () => chip('◦'), run: () => c.list('bullet'), checked: () => block()?.list === 'bullet' },
              '-', { heading: t('Docs.List.Level') },
              ...[0, 1, 2, 3, 4, 5, 6, 7, 8].map(l => ({ label: t('Docs.List.LevelN', l + 1), disabled: () => !block()?.list, checked: () => !!block()?.list && (block().level || 0) === l, run: () => c.level(l) })),
            ] },
            { type: 'button', label: t('Docs.Outdent'), icon: 'docs-outdent', key: 'Mod+Shift+M', global: false, run: () => c.indent(-1) },
            { type: 'button', label: t('Docs.Indent'), icon: 'docs-indent', key: 'Mod+M', global: false, run: () => c.indent(1) },
            { type: 'toggle', label: t('Docs.Marks'), icon: 'docs-pilcrow', edit: false, active: () => app.marks, run: () => app.toggleMarks() },
          ],
          [
            { type: 'toggle', label: t('Docs.AlignLeft'), icon: 'docs-align-left', key: 'Mod+L', global: false, active: () => align() === 'left', run: () => c.align('left') },
            { type: 'toggle', label: t('Docs.AlignCenter'), icon: 'docs-align-center', key: 'Mod+E', global: false, active: () => align() === 'center', run: () => c.align('center') },
            { type: 'toggle', label: t('Docs.AlignRight'), icon: 'docs-align-right', key: 'Mod+R', global: false, active: () => align() === 'right', run: () => c.align('right') },
            { type: 'toggle', label: t('Docs.Justify'), icon: 'docs-align-justify', key: 'Mod+J', global: false, active: () => align() === 'justify', run: () => c.align('justify') },
            { type: 'menu', label: t('Docs.LineSpacing'), icon: 'docs-line-spacing', showLabel: false, items: () => [
              ...LINE_SPACINGS.map(([word, css]) => ({ label: numStr(word).replace('.', t('Docs.DecimalPoint')), checked: () => !!block() && Math.abs(parseFloat(currentLineHeight(block())) - parseFloat(css)) < 0.02, run: () => c.lineSpacing(css) })),
              '-', { label: t('Docs.LineSpacing.Options'), run: () => app.dialogs.paragraph() },
              '-', { label: currentSpacing(block() || { type: 'p', css: {} }, 'before') > 0 ? t('Docs.Spacing.RemoveBefore') : t('Docs.Spacing.AddBefore'), run: () => c.spacing('before', v => (v > 0 ? 0 : 12)) },
              { label: currentSpacing(block() || { type: 'p', css: {} }, 'after') > 0 ? t('Docs.Spacing.RemoveAfter') : t('Docs.Spacing.AddAfter'), run: () => c.spacing('after', v => (v > 0 ? 0 : 8)) },
            ] },
            { type: 'color', label: t('Docs.Shading'), icon: 'docs-shading', auto: t('Docs.NoColour'), initial: '#d9e2f3', change: v => c.shading(v) },
            { type: 'menu', label: t('Docs.Borders'), icon: 'docs-borders', showLabel: false, items: [
              { label: t('Docs.Border.bottom'), icon: 'docs-border-bottom', run: () => c.borders('bottom') },
              { label: t('Docs.Border.top'), icon: 'docs-border-top', run: () => c.borders('top') },
              { label: t('Docs.Border.left'), icon: 'docs-border-left', run: () => c.borders('left') },
              { label: t('Docs.Border.right'), icon: 'docs-border-right', run: () => c.borders('right') },
              { label: t('Docs.Border.box'), icon: 'docs-borders', run: () => c.borders('box') },
              { label: t('Docs.Border.none'), icon: 'docs-border-none', run: () => c.borders('none') },
              '-', { label: t('Docs.HorizontalLine'), icon: 'docs-hline', run: () => c.horizontalLine() },
            ] },
          ],
        ] },
      ] },
      { label: t('Docs.Group.Styles'), wide: true, items: [{ type: 'gallery', label: t('Docs.Styles'), items: styleItems }] },
      { label: t('Docs.Group.Editing'), items: [
        { type: 'stack', rows: [
          [{ type: 'split', label: t('Docs.Find'), icon: 'search', showLabel: true, edit: false, key: 'Mod+F', global: false, run: () => app.findUi.open(false), items: [
            { label: t('Docs.Find'), icon: 'search', key: 'Mod+F', edit: false, run: () => app.findUi.open(false) },
            { label: t('Docs.GoTo'), icon: 'docs-goto', key: 'Mod+G', edit: false, run: () => app.dialogs.goTo() },
          ] }],
          [{ type: 'button', label: t('Docs.Replace'), icon: 'docs-replace', showLabel: true, key: 'Mod+H', global: false, run: () => app.findUi.open(true) }],
          [{ type: 'menu', label: t('Docs.Select'), icon: 'docs-select', edit: false, items: [
            { label: t('Docs.SelectAll'), key: 'Mod+A', edit: false, run: () => c.selectAll() },
          ] }],
        ] },
      ] },
    ],
  };

  const insert = {
    id: 'insert', label: t('Docs.Tab.Insert'), groups: [
      { label: t('Docs.Group.Pages'), items: [
        { type: 'button', label: t('Docs.PageBreak'), icon: 'docs-page-break', size: 'large', key: 'Mod+Enter', global: false, disabled: () => !inBody() || inCell(), run: () => c.pageBreak('page') },
        { type: 'button', label: t('Docs.BlankPage'), icon: 'docs-blank-page', size: 'large', disabled: () => !inBody() || inCell(), run: () => c.blankPage() },
      ] },
      { label: t('Docs.Group.Tables'), items: [
        { type: 'custom', render: () => tablePicker(app), refresh: (node, ro) => { node.disabled = ro; } },
      ] },
      { label: t('Docs.Group.Illustrations'), items: [
        { type: 'button', label: t('Docs.Pictures'), icon: 'image', size: 'large', run: () => images.pick() },
        { type: 'button', label: t('Docs.HorizontalLine'), icon: 'docs-hline', size: 'large', run: () => c.horizontalLine() },
      ] },
      { label: t('Docs.Group.Links'), items: [
        { type: 'button', label: t('Docs.Link'), icon: 'link', size: 'large', key: 'Mod+K', global: false, run: () => app.dialogs.link() },
      ] },
      { label: t('Docs.Group.Comments'), items: [
        { type: 'button', label: t('Docs.Comment'), icon: 'comment', size: 'large', disabled: () => !inBody(), run: () => app.commentsUi.add() },
      ] },
      { label: t('Docs.Group.HeaderFooter'), items: [
        { type: 'menu', label: t('Docs.Header'), icon: 'docs-header', size: 'large', items: [
          { label: t('Docs.Header.Edit'), icon: 'edit', run: () => hf.open('header') },
          { label: t('Docs.Header.Remove'), icon: 'trash', disabled: () => !app.headers.header && !app.headers.firstHeader, run: () => hf.remove('header') },
        ] },
        { type: 'menu', label: t('Docs.Footer'), icon: 'docs-footer', size: 'large', items: [
          { label: t('Docs.Footer.Edit'), icon: 'edit', run: () => hf.open('footer') },
          { label: t('Docs.Footer.Remove'), icon: 'trash', disabled: () => !app.headers.footer && !app.headers.firstFooter, run: () => hf.remove('footer') },
        ] },
        { type: 'menu', label: t('Docs.PageNumber'), icon: 'docs-page-number', size: 'large', items: [
          { label: t('Docs.PageNumber.Top'), run: () => hf.pageNumbers('header') },
          { label: t('Docs.PageNumber.Bottom'), run: () => hf.pageNumbers('footer') },
          { label: t('Docs.PageNumber.BottomOf'), run: () => hf.pageNumbers('footer', true) },
          { label: t('Docs.PageNumber.Here'), run: () => c.field('page') },
          '-', { label: t('Docs.PageNumber.Remove'), icon: 'trash', run: () => hf.removePageNumbers() },
        ] },
      ] },
      { label: t('Docs.Group.Text'), items: [
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.DateTime'), icon: 'calendar', showLabel: true, run: () => app.dialogs.dateTime() },
            { type: 'button', label: t('Docs.InsertFile'), icon: 'file-text', showLabel: true, disabled: () => !inBody(), run: () => app.insertFile() }],
          [{ type: 'menu', label: t('Docs.Symbol'), icon: 'docs-symbol', items: () => [
            ...'©®™§¶°±×÷€£¥–—…•→←≤≥≠∞½¼¾✓'.split('').map(ch => ({ label: ch, render: () => chip(ch), run: () => c.text(ch) })),
            '-', { label: t('Docs.Symbol.More'), run: () => app.dialogs.symbols() },
          ] }],
          [{ type: 'menu', label: t('Docs.Field'), icon: 'docs-field', items: [
            { label: t('Docs.Field.page'), run: () => c.field('page') },
            { label: t('Docs.Field.pages'), run: () => c.field('pages') },
            { label: t('Docs.Field.date'), run: () => c.field('date') },
            { label: t('Docs.Field.time'), run: () => c.field('time') },
            { label: t('Docs.Field.title'), run: () => c.field('title') },
          ] }],
        ] },
      ] },
    ],
  };

  const page = () => app.page;
  const setPage = changes => app.layoutCommands.setPage(changes);
  const layout = {
    id: 'layout', label: t('Docs.Tab.Layout'), groups: [
      { label: t('Docs.Group.PageSetup'), items: [
        { type: 'menu', label: t('Docs.Margins'), icon: 'docs-margins', size: 'large', items: () => [
          ...Object.entries(MARGIN_PRESETS).map(([key, m]) => ({ label: t('Docs.Margins.' + key), checked: () => sameMargins(page().margins, m), run: () => setPage({ margins: { ...m } }) })),
          '-', { label: t('Docs.Margins.Custom'), run: () => app.dialogs.pageSetup() },
        ] },
        { type: 'menu', label: t('Docs.Orientation'), icon: 'docs-orientation', size: 'large', items: [
          { label: t('Docs.Portrait'), icon: 'docs-portrait', checked: () => page().orientation === 'portrait', run: () => setPage({ orientation: 'portrait' }) },
          { label: t('Docs.Landscape'), icon: 'docs-landscape', checked: () => page().orientation === 'landscape', run: () => setPage({ orientation: 'landscape' }) },
        ] },
        { type: 'menu', label: t('Docs.Size'), icon: 'docs-size', size: 'large', items: () => [
          ...Object.entries(PAGE_SIZES).map(([name, [w, h]]) => ({ label: `${name}  ${numStr(w / 10)} × ${numStr(h / 10)} cm`, checked: () => page().size === name, run: () => setPage({ size: name }) })),
          '-', { label: t('Docs.Size.Custom'), run: () => app.dialogs.pageSetup('paper') },
        ] },
        { type: 'menu', label: t('Docs.Columns'), icon: 'docs-columns', size: 'large', items: [1, 2, 3].map(n => ({ label: t('Docs.Columns.' + n), checked: () => page().columns === n, run: () => setPage({ columns: n }) })) },
        { type: 'menu', label: t('Docs.Breaks'), icon: 'docs-breaks', size: 'large', disabled: () => !inBody() || inCell(), items: [
          { label: t('Docs.Break.page'), icon: 'docs-page-break', key: 'Mod+Enter', run: () => c.pageBreak('page') },
          { label: t('Docs.Break.column'), icon: 'docs-columns', key: 'Mod+Shift+Enter', run: () => c.pageBreak('column') },
        ] },
      ] },
      { label: t('Docs.Group.Paragraph'), items: [
        { type: 'stack', rows: [
          [{ type: 'custom', render: () => tinyLabel(t('Docs.Indent.Title')) }],
          [
            { type: 'custom', render: () => tinyLabel(t('Docs.Indent.Left')) },
            { type: 'number', label: t('Docs.Indent.Left'), width: 64, min: 0, max: 50, step: 0.25, value: () => cm(block()?.css?.['margin-left']), change: v => c.setIndents({ left: v * PT_PER_CM }) },
          ],
          [
            { type: 'custom', render: () => tinyLabel(t('Docs.Indent.Right')) },
            { type: 'number', label: t('Docs.Indent.Right'), width: 64, min: 0, max: 50, step: 0.25, value: () => cm(block()?.css?.['margin-right']), change: v => c.setIndents({ right: v * PT_PER_CM }) },
          ],
        ] },
        { type: 'stack', rows: [
          [{ type: 'custom', render: () => tinyLabel(t('Docs.Spacing.Title')) }],
          [
            { type: 'custom', render: () => tinyLabel(t('Docs.Spacing.Before')) },
            { type: 'number', label: t('Docs.Spacing.Before'), width: 64, min: 0, max: 600, step: 6, value: () => (block() ? numStr(currentSpacing(block(), 'before')) : ''), change: v => c.spacing('before', v) },
          ],
          [
            { type: 'custom', render: () => tinyLabel(t('Docs.Spacing.After')) },
            { type: 'number', label: t('Docs.Spacing.After'), width: 64, min: 0, max: 600, step: 6, value: () => (block() ? numStr(currentSpacing(block(), 'after')) : ''), change: v => c.spacing('after', v) },
          ],
        ] },
        { type: 'button', label: t('Docs.Paragraph.Dialog'), icon: 'docs-paragraph', size: 'large', run: () => app.dialogs.paragraph() },
      ] },
    ],
  };

  const references = {
    id: 'references', label: t('Docs.Tab.References'), groups: [
      { label: t('Docs.Group.Toc'), items: [
        { type: 'menu', label: t('Docs.Toc'), icon: 'docs-toc', size: 'large', items: [
          { label: t('Docs.Toc.Insert'), icon: 'docs-toc', run: () => app.navUi.insertToc() },
          { label: t('Docs.Toc.Update'), icon: 'refresh', disabled: () => !app.navUi.hasToc(), run: () => app.navUi.updateToc() },
          { label: t('Docs.Toc.Remove'), icon: 'trash', disabled: () => !app.navUi.hasToc(), run: () => app.navUi.removeToc() },
        ] },
        { type: 'button', label: t('Docs.Toc.Update'), icon: 'refresh', showLabel: true, disabled: () => !app.navUi.hasToc(), run: () => app.navUi.updateToc() },
      ] },
      { label: t('Docs.Group.Notes'), items: [
        { type: 'button', label: t('Docs.Note.Insert'), icon: 'docs-footnote', size: 'large', key: 'Mod+Alt+F', global: false, disabled: () => !inBody(), run: () => app.notes.insert() },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Note.Endnote'), icon: 'docs-endnote', showLabel: true, disabled: () => !inBody(), run: () => app.notes.insert() }],
          [{ type: 'button', label: t('Docs.Note.Next'), icon: 'chevron-down', showLabel: true, edit: false, run: () => app.notes.next() }],
          [{ type: 'button', label: t('Docs.Note.Show'), icon: 'eye', showLabel: true, edit: false, run: () => app.notes.show() }],
        ] },
      ] },
      { label: t('Docs.Group.Captions'), items: [
        { type: 'button', label: t('Docs.Caption.Insert'), icon: 'docs-caption', size: 'large', disabled: () => !inBody(), run: () => app.dialogs.caption() },
      ] },
    ],
  };

  const review = {
    id: 'review', label: t('Docs.Tab.Review'), groups: [
      { label: t('Docs.Group.Proofing'), items: [
        { type: 'toggle', label: t('Docs.Spelling'), icon: 'docs-spelling', size: 'large', active: () => app.settings.spellcheck, run: () => app.layoutCommands.setSettings({ spellcheck: !app.settings.spellcheck }) },
        { type: 'button', label: t('Docs.WordCount'), icon: 'docs-word-count', size: 'large', edit: false, run: () => app.dialogs.wordCount() },
        { type: 'button', label: t('Docs.Language'), icon: 'globe', size: 'large', run: () => app.dialogs.language() },
      ] },
      { label: t('Docs.Group.Comments'), items: [
        { type: 'button', label: t('Docs.Comment.New'), icon: 'comment', size: 'large', key: 'Mod+Alt+M', global: false, disabled: () => !inBody(), run: () => app.commentsUi.add() },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Comment.Delete'), icon: 'trash', showLabel: true, disabled: () => !app.commentsUi.active, run: () => app.commentsUi.remove(app.commentsUi.active) }],
          [{ type: 'button', label: t('Docs.Comment.Previous'), icon: 'chevron-up', showLabel: true, edit: false, disabled: () => !app.commentsUi.count(), run: () => app.commentsUi.step(-1) },
            { type: 'button', label: t('Docs.Comment.Next'), icon: 'chevron-down', showLabel: true, edit: false, disabled: () => !app.commentsUi.count(), run: () => app.commentsUi.step(1) }],
          [{ type: 'toggle', label: t('Docs.Comment.Show'), icon: 'message-square', showLabel: true, edit: false, active: () => app.commentsUi.visible, run: () => app.commentsUi.toggle() }],
        ] },
      ] },
    ],
  };

  const view = {
    id: 'view', label: t('Docs.Tab.View'), groups: [
      { label: t('Docs.Group.Views'), items: [
        { type: 'toggle', label: t('Docs.View.Print'), icon: 'docs-print-layout', size: 'large', edit: false, active: () => app.viewMode === 'print', run: () => app.setView('print') },
        { type: 'toggle', label: t('Docs.View.Web'), icon: 'docs-web-layout', size: 'large', edit: false, active: () => app.viewMode === 'web', run: () => app.setView('web') },
      ] },
      { label: t('Docs.Group.Show'), items: [
        { type: 'stack', rows: [
          [{ type: 'toggle', label: t('Docs.Nav'), icon: 'docs-nav', showLabel: true, edit: false, active: () => app.navUi.visible, run: () => app.navUi.toggle() }],
          [{ type: 'toggle', label: t('Docs.Marks'), icon: 'docs-pilcrow', showLabel: true, edit: false, active: () => app.marks, run: () => app.toggleMarks() }],
          [{ type: 'toggle', label: t('Docs.Comment.Show'), icon: 'message-square', showLabel: true, edit: false, active: () => app.commentsUi.visible, run: () => app.commentsUi.toggle() }],
        ] },
      ] },
      { label: t('Docs.Group.Zoom'), items: [
        { type: 'menu', label: t('Docs.Zoom'), icon: 'zoom-in', size: 'large', edit: false, items: [50, 75, 100, 125, 150, 200].map(p => ({ label: p + ' %', edit: false, checked: () => Math.round(app.view.zoom * 100) === p, run: () => app.zoomTo(p / 100) })) },
        { type: 'button', label: '100 %', icon: 'docs-zoom-100', size: 'large', edit: false, run: () => app.zoomTo(1) },
        { type: 'button', label: t('Docs.Zoom.OnePage'), icon: 'docs-one-page', size: 'large', edit: false, disabled: () => app.viewMode !== 'print', run: () => { app.view.fitPage(); app.ctx.status.setZoom(app.view.zoom); } },
        { type: 'button', label: t('Docs.Zoom.PageWidth'), icon: 'zoom-fit', size: 'large', edit: false, disabled: () => app.viewMode !== 'print', run: () => { app.view.fit(); app.ctx.status.setZoom(app.view.zoom); } },
      ] },
    ],
  };

  const tableTab = {
    id: 'table', label: t('Docs.Tab.Table'), hidden: () => !tables.current(), groups: [
      { label: t('Docs.Group.TableStyles'), wide: true, items: [
        { type: 'gallery', label: t('Docs.Table.Styles'), items: TABLE_STYLES.map(style => ({ label: t('Docs.TableStyle.' + style), render: () => tableChip(style), run: () => tables.setStyle(style), active: () => tables.current()?.table.getAttribute('data-table-style') === style })) },
        { type: 'stack', rows: [
          [{ type: 'toggle', label: t('Docs.Table.HeaderRow'), icon: 'docs-table-header', showLabel: true, active: () => tables.hasHeader(), run: () => tables.toggleHeader() }],
          [{ type: 'color', label: t('Docs.Table.Shading'), icon: 'docs-shading', auto: t('Docs.NoColour'), initial: '#d9e2f3', change: v => tables.shading(v) }],
          [{ type: 'menu', label: t('Docs.Table.Borders'), icon: 'docs-borders', items: [
            { label: t('Docs.Border.all'), icon: 'docs-borders', run: () => tables.borders('all') },
            { label: t('Docs.Border.bottom'), icon: 'docs-border-bottom', run: () => tables.borders('bottom') },
            { label: t('Docs.Border.top'), icon: 'docs-border-top', run: () => tables.borders('top') },
            { label: t('Docs.Border.none'), icon: 'docs-border-none', run: () => tables.borders('none') },
          ] }],
        ] },
      ] },
      { label: t('Docs.Group.RowsColumns'), items: [
        { type: 'button', label: t('Docs.Table.InsertAbove'), icon: 'docs-row-above', size: 'large', run: () => tables.insertRow(false) },
        { type: 'button', label: t('Docs.Table.InsertBelow'), icon: 'docs-row-below', size: 'large', run: () => tables.insertRow(true) },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Table.InsertLeft'), icon: 'docs-col-left', showLabel: true, run: () => tables.insertColumn(false) }],
          [{ type: 'button', label: t('Docs.Table.InsertRight'), icon: 'docs-col-right', showLabel: true, run: () => tables.insertColumn(true) }],
          [{ type: 'menu', label: t('Docs.Table.Delete'), icon: 'docs-table-delete', items: [
            { label: t('Docs.Table.DeleteRow'), icon: 'docs-row-delete', run: () => tables.deleteRow() },
            { label: t('Docs.Table.DeleteColumn'), icon: 'docs-col-delete', run: () => tables.deleteColumn() },
            { label: t('Docs.Table.DeleteTable'), icon: 'trash', danger: true, run: () => tables.deleteTable() },
          ] }],
        ] },
      ] },
      { label: t('Docs.Group.Merge'), items: [
        { type: 'button', label: t('Docs.Table.Merge'), icon: 'docs-merge', size: 'large', disabled: () => !tables.canMerge(), run: () => tables.merge() },
        { type: 'button', label: t('Docs.Table.Split'), icon: 'docs-split', size: 'large', run: () => tables.split() },
      ] },
      { label: t('Docs.Group.CellSize'), items: [
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Docs.Table.Distribute'), icon: 'docs-distribute', showLabel: true, run: () => tables.distribute() }],
          [{ type: 'toggle', label: t('Docs.Table.AlignTop'), icon: 'docs-valign-top', active: () => tables.valign() === 'top', run: () => tables.setValign('top') },
            { type: 'toggle', label: t('Docs.Table.AlignMiddle'), icon: 'docs-valign-middle', active: () => tables.valign() === 'middle', run: () => tables.setValign('middle') },
            { type: 'toggle', label: t('Docs.Table.AlignBottom'), icon: 'docs-valign-bottom', active: () => tables.valign() === 'bottom', run: () => tables.setValign('bottom') }],
          [{ type: 'menu', label: t('Docs.Select'), icon: 'docs-select', edit: false, items: [
            { label: t('Docs.Table.SelectCell'), edit: false, run: () => tables.select('cell') },
            { label: t('Docs.Table.SelectRow'), edit: false, run: () => tables.select('row') },
            { label: t('Docs.Table.SelectColumn'), edit: false, run: () => tables.select('column') },
            { label: t('Docs.Table.SelectTable'), edit: false, run: () => tables.select('table') },
          ] }],
        ] },
      ] },
    ],
  };

  const pictureTab = {
    id: 'picture', label: t('Docs.Tab.Picture'), hidden: () => !images.selected, groups: [
      { label: t('Docs.Group.Arrange'), items: [
        { type: 'toggle', label: t('Docs.Wrap.inline'), icon: 'docs-wrap-inline', size: 'large', active: () => images.align() === '', run: () => images.setAlign('') },
        { type: 'toggle', label: t('Docs.Wrap.left'), icon: 'docs-wrap-left', size: 'large', active: () => images.align() === 'left', run: () => images.setAlign('left') },
        { type: 'toggle', label: t('Docs.Wrap.right'), icon: 'docs-wrap-right', size: 'large', active: () => images.align() === 'right', run: () => images.setAlign('right') },
        { type: 'toggle', label: t('Docs.Wrap.center'), icon: 'docs-wrap-center', size: 'large', active: () => images.align() === 'center', run: () => images.setAlign('center') },
      ] },
      { label: t('Docs.Group.Size'), items: [
        { type: 'stack', rows: [
          [{ type: 'custom', render: () => tinyLabel(t('Docs.Picture.Width')) }, { type: 'number', label: t('Docs.Picture.Width'), width: 72, min: 0.1, max: 100, step: 0.1, value: () => images.sizeCm('width'), change: v => images.setSizeCm('width', v) }],
          [{ type: 'custom', render: () => tinyLabel(t('Docs.Picture.Height')) }, { type: 'number', label: t('Docs.Picture.Height'), width: 72, min: 0.1, max: 100, step: 0.1, value: () => images.sizeCm('height'), change: v => images.setSizeCm('height', v) }],
        ] },
        { type: 'button', label: t('Docs.Picture.Reset'), icon: 'refresh', size: 'large', run: () => images.resetSize() },
      ] },
      { label: t('Docs.Group.Accessibility'), items: [
        { type: 'button', label: t('Docs.Picture.Alt'), icon: 'docs-alt-text', size: 'large', run: () => images.altText() },
        { type: 'button', label: t('Docs.Picture.Delete'), icon: 'trash', size: 'large', run: () => images.remove() },
      ] },
    ],
  };

  const hfTab = {
    id: 'hf', label: t('Docs.Tab.HeaderFooter'), hidden: () => !hf.editing, groups: [
      { label: t('Docs.Group.Insert'), items: [
        { type: 'menu', label: t('Docs.Field'), icon: 'docs-field', size: 'large', items: [
          { label: t('Docs.Field.page'), run: () => c.field('page') },
          { label: t('Docs.Field.pages'), run: () => c.field('pages') },
          { label: t('Docs.Field.pageOf'), run: () => hf.insertPageOf() },
          { label: t('Docs.Field.date'), run: () => c.field('date') },
          { label: t('Docs.Field.title'), run: () => c.field('title') },
        ] },
        { type: 'button', label: t('Docs.Picture'), icon: 'image', size: 'large', run: () => images.pick() },
      ] },
      { label: t('Docs.Group.Options'), items: [
        { type: 'toggle', label: t('Docs.DifferentFirst'), icon: 'docs-first-page', size: 'large', active: () => page().differentFirst, run: () => hf.toggleDifferentFirst() },
        { type: 'stack', rows: [
          [{ type: 'custom', render: () => tinyLabel(t('Docs.HeaderFrom')) }, { type: 'number', label: t('Docs.HeaderFrom'), width: 64, min: 0, max: 15, step: 0.1, value: () => numStr(page().header / 10), change: v => setPage({ header: v * 10 }) }],
          [{ type: 'custom', render: () => tinyLabel(t('Docs.FooterFrom')) }, { type: 'number', label: t('Docs.FooterFrom'), width: 64, min: 0, max: 15, step: 0.1, value: () => numStr(page().footer / 10), change: v => setPage({ footer: v * 10 }) }],
        ] },
      ] },
      { label: t('Docs.Group.Close'), items: [
        { type: 'button', label: t('Docs.HeaderFooter.Close'), icon: 'x', size: 'large', edit: false, run: () => hf.close() },
      ] },
    ],
  };

  return [home, insert, layout, references, review, view, tableTab, pictureTab, hfTab];

  function align() {
    const b = block();
    if (!b) return 'left';
    return b.css?.['text-align'] || (b.type === 'blockquote' ? 'center' : 'left');
  }

}

// ---- small pieces ----------------------------------------------------------------------------------------------------------------

const lastBullet = { kind: 'bullet' };
const lastNumber = { kind: 'decimal' };
const BULLET_KINDS = ['bullet', 'circle', 'square', 'dash', 'check', 'arrow'];
const isBullet = k => BULLET_KINDS.includes(k);
const isNumber = k => !!k && !isBullet(k);
const kindOf = item => item.kind;

function chip(text) { return el('span', { class: 'mo-docs-menuchip', text }); }

function tinyLabel(text) { return el('span', { class: 'mo-docs-tiny', text }); }

function cm(css) {
  const p = toPt(css || '0') || 0;
  return numStr(p / PT_PER_CM);
}

function sameMargins(a, b) { return ['top', 'right', 'bottom', 'left'].every(k => Math.abs(a[k] - b[k]) < 0.05); }

function fontOptions(current) {
  const list = [...FONTS];
  if (current && !list.includes(current)) list.unshift(current);
  return list.map(f => ({ value: f, label: f }));
}

function styleChip(key) {
  const s = STYLES.find(x => x.key === key);
  const sample = el('span', { class: 'mo-docs-stylechip__sample', text: key === 'Title' ? 'AaBb' : 'AaBbCc' });
  sample.style.fontFamily = fontStack(s.font);
  sample.style.fontSize = Math.min(18, Math.max(10, s.size * (key === 'Title' ? 0.62 : 1))) + 'pt';
  if (s.color) sample.style.color = s.color;
  if (s.italic) sample.style.fontStyle = 'italic';
  if (key === 'Code') sample.style.background = '#f2f2f2';
  return el('span', { class: 'mo-docs-stylechip' }, sample, el('span', { class: 'mo-docs-stylechip__name', text: t('Docs.Style.' + key) }));
}

function tableChip(style) {
  const box = el('span', { class: 'mo-docs-tablechip', 'data-table-style': style });
  const look = {
    grid: ['#000', null, null], plain: ['transparent', null, null], light: ['#bfbfbf', null, null], striped: ['transparent', null, '#f2f2f2'],
    accent1: ['#8eaadb', '#4472c4', '#d9e2f3'], accent2: ['#f4b083', '#ed7d31', '#fbe4d5'], accent6: ['#a8d08d', '#70ad47', '#e2efd9'],
  }[style];
  for (let i = 0; i < 12; i++) {
    const cell = el('span');
    cell.style.borderColor = look[0];
    const row = Math.floor(i / 4);
    if (row === 0 && look[1]) cell.style.background = look[1];
    else if (row % 2 === 1 && look[2]) cell.style.background = look[2];
    box.append(cell);
  }
  return el('span', { class: 'mo-docs-stylechip', style: 'height: 54px' }, box);
}

/** The highlight colour: a split button with Word's highlighter colours. */
function highlightButton(app) {
  let last = '#ffff00';
  const bar = el('span', { class: 'mo-rcolor__bar' });
  bar.style.background = last;
  const main = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--icon mo-rcolor__main', 'aria-label': t('Docs.Highlight'), 'data-tip': t('Docs.Highlight'), onclick: () => { app.commands.setHighlight(last); app.ribbonRefresh(); } }, icon('docs-highlight'), bar);
  const arrow = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--arrow', 'aria-label': t('Docs.Highlight.More'), 'aria-haspopup': 'true' }, icon('chevron-down'));
  arrow.addEventListener('click', () => openPopover(arrow, close => {
    const box = el('div', { class: 'mo-palette' });
    const row = el('div', { class: 'mo-palette__row' });
    for (const c of HIGHLIGHTS) {
      const chipEl = el('button', { type: 'button', class: 'mo-palette__chip', 'aria-label': c, 'data-tip': c, onclick: () => { close(); last = c; bar.style.background = c; app.commands.setHighlight(c); app.ribbonRefresh(); } });
      chipEl.style.background = c;
      row.append(chipEl);
    }
    box.append(row, el('button', { type: 'button', class: 'mo-palette__auto', onclick: () => { close(); app.commands.setHighlight(null); app.ribbonRefresh(); } }, el('span', { class: 'mo-palette__chip mo-palette__chip--auto' }), el('span', { text: t('Docs.NoColour') })));
    return box;
  }));
  return el('span', { class: 'mo-rsplit mo-rcolor' }, main, arrow);
}

/** Insert → Table: a grid to pick rows × columns (10 × 8, like Word) and the dialog for more. */
function tablePicker(app) {
  const button = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--menu mo-rbtn--lg', 'aria-haspopup': 'true' }, icon('table', 'mo-rbtn__icon--lg'), el('span', { class: 'mo-rbtn__label', text: t('Docs.Table') }), icon('chevron-down', 'mo-rbtn__chev'));
  button.addEventListener('click', () => {
    if (app.ctx.readOnly) return;
    openPopover(button, close => {
      const label = el('div', { class: 'mo-docs-grid__label', text: t('Docs.Table.Insert') });
      const grid = el('div', { class: 'mo-docs-grid', role: 'grid', 'aria-label': t('Docs.Table.Insert') });
      const cells = [];
      const mark = (r, c) => {
        cells.forEach(x => x.classList.toggle('is-on', Number(x.dataset.r) <= r && Number(x.dataset.c) <= c));
        label.textContent = r >= 0 ? t('Docs.Table.Size', c + 1, r + 1) : t('Docs.Table.Insert');
      };
      for (let r = 0; r < 8; r++) {
        for (let c = 0; c < 10; c++) {
          const cell = el('button', { type: 'button', class: 'mo-docs-grid__cell', 'aria-label': t('Docs.Table.Size', c + 1, r + 1), 'data-r': r, 'data-c': c });
          cell.addEventListener('pointerenter', () => mark(r, c));
          cell.addEventListener('focus', () => mark(r, c));
          cell.addEventListener('click', () => { close(); app.tables.insert(r + 1, c + 1); });
          cells.push(cell);
          grid.append(cell);
        }
      }
      return el('div', {}, label, grid, el('div', { class: 'mo-menu' },
        el('button', { type: 'button', class: 'mo-menu__item', onclick: () => { close(); app.dialogs.insertTable(); } }, icon('table'), el('span', { class: 'mo-menu__label', text: t('Docs.Table.Dialog') }))));
    });
  });
  return button;
}

