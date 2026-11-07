/* MatOffice Slides – the ribbon: Home, Insert, Design, Transitions, Animations, Slide Show, View and the contextual tabs Shape Format,
   Picture Format and Table (shown while such an object is selected), laid out like the office suites. Items: see frame.js (Ribbon). */
import { t, el, icon } from '../editor/util.js';
import { openPopover } from '../editor/frame.js';
import { THEMES, VARIANTS, isNumbering, NUMBERINGS, BULLETS } from './model.js';
import { colorButton, shapeGallery, quickStyles, tableGrid, symbolPicker, themePreview, variantPreview, layoutGallery } from './ui.js';
import { tableLook, resolveColor } from './model.js';

export const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Arial Black', 'Bahnschrift', 'Cambria', 'Candara', 'Consolas', 'Constantia', 'Corbel', 'Courier New',
  'Franklin Gothic Medium', 'Georgia', 'Palatino Linotype', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const TRANSITIONS = ['none', 'fade', 'push', 'wipe', 'split', 'zoom', 'cover'];
const EFFECTS = ['none', 'appear', 'fade', 'fly', 'zoom', 'wipe'];

export function buildRibbon(app) {
  const ro = () => app.readOnly;
  const sel = () => app.canvas.selected();
  const none = () => !sel().length;
  const noText = () => !app.hasText();
  const f = () => app.fmt();
  const canRun = () => !ro();
  const theme = () => app.content.theme;

  /** A custom ribbon item from a node with refresh(). */
  const custom = make => ({ type: 'custom', render: () => make(), refresh: node => node.refresh?.() });

  const fill = (label, iconName, opts) => custom(() => colorButton({ label, iconName, theme, canRun: () => canRun() && opts.enabled(), ...opts }));

  // ---- Home -----------------------------------------------------------------------------------------------------------------------------

  const clipboard = {
    label: t('Slides.Group.Clipboard'), items: [
      { type: 'split', size: 'large', id: 'paste', label: t('Slides.Paste'), icon: 'slides-paste', run: () => app.paste(), items: [
        { label: t('Slides.Paste'), icon: 'slides-paste', key: 'Ctrl+V', run: () => app.paste() },
        { label: t('Slides.PasteText'), icon: 'type', run: async () => { try { const text = await navigator.clipboard.readText(); if (text) app.pasteText(text); } catch { app.ctx.toast(t('Slides.PasteKeys'), 'info'); } } },
      ] },
      { type: 'stack', rows: [
        [{ type: 'button', label: t('Slides.Cut'), icon: 'slides-cut', showLabel: true, tip: t('Slides.Cut') + ' (Ctrl+X)', run: () => app.cut(), disabled: () => none() && app.area !== 'thumbs' }],
        [{ type: 'button', label: t('Slides.Copy'), icon: 'copy', showLabel: true, tip: t('Slides.Copy') + ' (Ctrl+C)', edit: false, run: () => app.copy() }],
        [{ type: 'toggle', label: t('Slides.FormatPainter'), icon: 'slides-format-painter', showLabel: true, active: () => !!app.painter, run: () => (app.painter ? app.cancelPainter() : app.startPainter()), disabled: () => none() }],
      ] },
    ],
  };

  const slidesGroup = {
    label: t('Slides.Group.Slides'), items: [
      { type: 'split', size: 'large', id: 'new-slide', label: t('Slides.NewSlide'), icon: 'slides-new-slide', key: 'Mod+M', run: () => app.newSlide(), items: () => [
        { heading: t('Slides.Layouts') },
        ...app.layoutItems().map(x => ({ label: x.label, icon: x.icon, run: () => app.newSlide(x.layoutId) })),
        '-',
        { label: t('Slides.DuplicateSlide'), icon: 'slides-duplicate', run: () => app.duplicateSlides(), disabled: () => !app.slide },
      ] },
      { type: 'stack', rows: [
        [custom(() => anchorBtnLabel(t('Slides.Layout'), 'slides-layout', b => layoutGallery(b, app.content, app.slide?.layout, id => app.setLayout(id)), () => ro() || !app.slide))],
        [{ type: 'button', label: t('Slides.Reset'), icon: 'slides-reset', showLabel: true, tip: t('Slides.ResetTip'), run: () => app.resetSlides(), disabled: () => !app.slide }],
        [{ type: 'button', label: t('Slides.DeleteSlide'), icon: 'slides-delete-slide', showLabel: true, run: () => app.deleteSlides(), disabled: () => !app.slide }],
      ] },
    ],
  };

  const fontGroup = {
    label: t('Slides.Group.Font'), items: [
      { type: 'stack', rows: [
        [
          { type: 'select', id: 'font', label: t('Slides.Font'), width: 150, options: () => fontOptions(app), value: () => f().font ?? '', change: v => app.setRun({ font: v }), disabled: noText },
          { type: 'combo', id: 'size', label: t('Slides.FontSize'), width: 56, options: FONT_SIZES.map(String), inputmode: 'decimal', value: () => (f().size ? String(f().size) : ''), change: v => { const n = parseFloat(String(v).replace(',', '.')); if (n >= 1 && n <= 400) app.setRun({ size: Math.round(n * 2) / 2 }); }, disabled: noText },
          { type: 'button', label: t('Slides.Grow'), icon: 'slides-grow-font', key: 'Mod+Shift+>', run: () => app.growFont(1), disabled: noText },
          { type: 'button', label: t('Slides.Shrink'), icon: 'slides-shrink-font', key: 'Mod+Shift+<', run: () => app.growFont(-1), disabled: noText },
          { type: 'button', label: t('Slides.ClearFormat'), icon: 'slides-clear-format', run: () => app.clearFormat(), disabled: noText },
        ],
        [
          { type: 'toggle', id: 'bold', label: t('Slides.Bold'), icon: 'slides-bold', key: 'Mod+B', active: () => !!f().b, run: () => app.toggleRun('b'), disabled: noText },
          { type: 'toggle', id: 'italic', label: t('Slides.Italic'), icon: 'slides-italic', key: 'Mod+I', active: () => !!f().i, run: () => app.toggleRun('i'), disabled: noText },
          { type: 'toggle', id: 'underline', label: t('Slides.Underline'), icon: 'slides-underline', key: 'Mod+U', active: () => !!f().u, run: () => app.toggleRun('u'), disabled: noText },
          { type: 'toggle', label: t('Slides.Strike'), icon: 'slides-strike', active: () => !!f().s, run: () => app.toggleRun('s'), disabled: noText },
          { type: 'separator' },
          fill(t('Slides.Highlight'), 'slides-highlight', { none: t('Slides.NoHighlight'), value: () => f().hl || '#ffff00', change: c => app.setRun({ hl: c === 'none' ? null : c || '#ffff00' }), enabled: () => !noText() }),
          fill(t('Slides.FontColor'), 'slides-font-color', { auto: t('Slides.Automatic'), value: () => f().color || '#c00000', change: c => app.setRun({ color: c }), enabled: () => !noText() }),
        ],
      ] },
    ],
  };

  const paragraph = {
    label: t('Slides.Group.Paragraph'), items: [
      { type: 'stack', rows: [
        [
          { type: 'split', label: t('Slides.Bullets'), icon: 'slides-bullets', showLabel: false, active: () => { const b = f().bullet; return !!b && b !== 'none' && !isNumbering(b); }, run: () => app.toggleBullets('bullet'), disabled: noText, items: () => [
            { label: t('Slides.NoBullets'), icon: 'slides-no-fill', run: () => app.setBullet('none') }, '-',
            ...BULLETS.map(b => ({ label: b, render: () => el('span', { class: 'mo-slides-bulletpic', text: b }), checked: f().bullet === b, run: () => app.setBullet(b) })),
          ] },
          { type: 'split', label: t('Slides.Numbering'), icon: 'slides-numbering', showLabel: false, active: () => isNumbering(f().bullet), run: () => app.toggleBullets('number'), disabled: noText, items: () => [
            { label: t('Slides.NoNumbering'), icon: 'slides-no-fill', run: () => app.setBullet('none') }, '-',
            ...NUMBERINGS.map(n => ({ label: numberingLabel(n), render: () => el('span', { class: 'mo-slides-bulletpic', text: n }), checked: f().bullet === n, run: () => app.setBullet(n) })),
          ] },
          { type: 'button', label: t('Slides.LevelDown'), icon: 'slides-indent-less', key: 'Alt+Shift+ArrowLeft', run: () => app.changeLevel(-1), disabled: noText },
          { type: 'button', label: t('Slides.LevelUp'), icon: 'slides-indent-more', key: 'Alt+Shift+ArrowRight', run: () => app.changeLevel(1), disabled: noText },
          { type: 'menu', label: t('Slides.LineSpacing'), icon: 'slides-line-spacing', showLabel: false, disabled: noText, items: () => [1, 1.15, 1.5, 2, 2.5, 3].map(v => ({ label: String(v).replace('.', t('Slides.DecimalSep')), checked: Math.abs((f().lineSpacing || 1) - v) < 0.01, run: () => app.setLineSpacing(v) })) },
        ],
        [
          { type: 'toggle', label: t('Slides.AlignLeft'), icon: 'slides-align-left', key: 'Mod+L', active: () => f().align === 'left', run: () => app.setAlign('left'), disabled: noText },
          { type: 'toggle', label: t('Slides.AlignCenter'), icon: 'slides-align-center', key: 'Mod+E', active: () => f().align === 'center', run: () => app.setAlign('center'), disabled: noText },
          { type: 'toggle', label: t('Slides.AlignRight'), icon: 'slides-align-right', key: 'Mod+R', active: () => f().align === 'right', run: () => app.setAlign('right'), disabled: noText },
          { type: 'toggle', label: t('Slides.Justify'), icon: 'slides-align-justify', key: 'Mod+J', active: () => f().align === 'justify', run: () => app.setAlign('justify'), disabled: noText },
          { type: 'menu', label: t('Slides.AlignText'), icon: 'slides-anchor-' + 'middle', showLabel: false, disabled: noText, items: () => ['top', 'middle', 'bottom'].map(a => ({ label: t('Slides.Anchor.' + a), icon: 'slides-anchor-' + a, checked: f().anchor === a, run: () => app.setBodyProp('anchor', a) })) },
        ],
      ] },
    ],
  };

  const drawing = {
    label: t('Slides.Group.Drawing'), items: [
      custom(() => anchorBtnLarge(t('Slides.Shapes'), 'slides-shapes', b => shapeGallery(b, kind => app.startDrawing(kind)), ro)),
      { type: 'menu', size: 'large', label: t('Slides.Arrange'), icon: 'slides-arrange', disabled: none, items: () => arrangeItems(app) },
      custom(() => anchorBtnLarge(t('Slides.QuickStyles'), 'slides-quick-styles', b => quickStyles(b, theme(), s => app.quickStyle(s)), () => ro() || !app.shapeTargets().length)),
      { type: 'stack', rows: [
        [fill(t('Slides.ShapeFill'), 'slides-fill', { none: t('Slides.NoFill'), value: () => sel()[0]?.fill?.color || 'accent1', change: c => app.setFill(c || 'accent1'), enabled: () => app.shapeTargets().length > 0 })],
        [fill(t('Slides.ShapeOutline'), 'slides-outline', { none: t('Slides.NoOutline'), value: () => sel()[0]?.line?.color || 'dk1', change: c => app.setLine(c === 'none' ? null : { color: c || 'dk1' }), enabled: () => sel().some(e => e.type !== 'table') })],
        [{ type: 'menu', label: t('Slides.ShapeEffects'), icon: 'slides-effects', showLabel: false, disabled: () => !sel().some(e => e.type !== 'table'), items: () => effectItems(app) }],
      ] },
    ],
  };

  const editing = {
    label: t('Slides.Group.Editing'), items: [{ type: 'stack', rows: [
      [{ type: 'button', label: t('Slides.Find'), icon: 'search', showLabel: true, edit: false, key: 'Mod+F', run: () => app.find(false) }],
      [{ type: 'button', label: t('Slides.Replace'), icon: 'slides-replace', showLabel: true, key: 'Mod+H', run: () => app.find(true) }],
      [{ type: 'menu', label: t('Slides.Select'), icon: 'cursor', edit: false, items: () => [
        { label: t('Slides.SelectAll'), icon: 'slides-select-all', key: 'Ctrl+A', run: () => app.selectAll() },
        { label: t('Slides.SelectNone'), icon: 'x', run: () => app.select([]) },
      ] }],
    ] }],
  };

  // ---- Insert ---------------------------------------------------------------------------------------------------------------------------

  const insert = [
    { label: t('Slides.Group.Slides'), items: [slidesGroup.items[0]] },
    { label: t('Slides.Group.Tables'), items: [custom(() => anchorBtnLarge(t('Slides.Table'), 'slides-table', b => tableGrid(b, (r, c) => app.insertTable(r, c), () => app.insertTableDialog()), ro))] },
    { label: t('Slides.Group.Images'), items: [{ type: 'button', size: 'large', label: t('Slides.Pictures'), icon: 'slides-picture', run: () => app.pickPictures() }] },
    { label: t('Slides.Group.Illustrations'), items: [custom(() => anchorBtnLarge(t('Slides.Shapes'), 'slides-shapes', b => shapeGallery(b, kind => app.startDrawing(kind)), ro))] },
    { label: t('Slides.Group.Text'), items: [
      { type: 'button', size: 'large', label: t('Slides.TextBox'), icon: 'slides-textbox', run: () => app.insertTextBox() },
      { type: 'button', size: 'large', label: t('Slides.HeaderFooter'), icon: 'slides-header-footer', run: () => app.headerFooter() },
      { type: 'stack', rows: [
        [{ type: 'toggle', label: t('Slides.SlideNumber'), icon: 'slides-slide-number', showLabel: true, active: () => !!app.content.footer?.number, run: () => app.toggleSlideNumbers() }],
        [{ type: 'button', label: t('Slides.DateTime'), icon: 'calendar', showLabel: true, run: () => app.headerFooter() }],
      ] },
    ] },
    { label: t('Slides.Group.Symbols'), items: [custom(() => anchorBtnLarge(t('Slides.Symbol'), 'slides-symbol', b => symbolPicker(b, ch => app.insertSymbol(ch)), ro))] },
  ];

  // ---- Design ---------------------------------------------------------------------------------------------------------------------------

  const design = [
    { label: t('Slides.Group.Themes'), items: [themeGallery(app)] },
    { label: t('Slides.Group.Variants'), items: [{ type: 'gallery', label: t('Slides.Variants'), items: [0, ...VARIANTS.map((v, i) => i + 1)].map(v => ({ label: v === 0 ? t('Slides.Variant.Own') : t('Slides.Variant.' + VARIANTS[v - 1].id), render: () => variantPreview(app.content, v, 58), active: () => (app.content.theme.variant || 0) === v, run: () => app.setVariant(v) })) }] },
    { label: t('Slides.Group.Customize'), items: [
      { type: 'menu', size: 'large', label: t('Slides.SlideSize'), icon: 'slides-slide-size', items: () => [
        { label: t('Slides.Size.Standard'), icon: 'slides-size-4-3', checked: app.content.size.w / app.content.size.h < 1.5, run: () => app.setSize('4:3') },
        { label: t('Slides.Size.Wide'), icon: 'slides-size-16-9', checked: app.content.size.w / app.content.size.h >= 1.5, run: () => app.setSize('16:9') },
      ] },
      { type: 'toggle', size: 'large', label: t('Slides.FormatBackground'), icon: 'slides-background', edit: false, active: () => app.pane.kind === 'format' && none(), run: () => { app.select([]); app.showPane('format'); } },
    ] },
  ];

  // ---- Transitions ----------------------------------------------------------------------------------------------------------------------

  const tr = () => app.transitionOf();
  const transitions = [
    { label: t('Slides.Group.Preview'), items: [{ type: 'button', size: 'large', label: t('Slides.Preview'), icon: 'slides-preview', edit: false, run: () => app.previewTransition(), disabled: () => tr().type === 'none' }] },
    { label: t('Slides.Group.Transition'), items: [
      { type: 'gallery', label: t('Slides.Transitions'), items: TRANSITIONS.map(ty => ({ label: t('Slides.Transition.' + ty), render: () => galleryTile('slides-tr-' + ty, t('Slides.Transition.' + ty)), active: () => tr().type === ty, run: () => app.setTransition({ type: ty }) })) },
      { type: 'menu', size: 'large', label: t('Slides.EffectOptions'), icon: 'slides-effect-options', disabled: () => !['push', 'wipe', 'cover', 'split'].includes(tr().type), items: () => ['right', 'left', 'down', 'up'].map(d => ({ label: t('Slides.From.' + d), icon: 'slides-dir-' + ({ right: 'left', left: 'right', down: 'up', up: 'down' }[d]), checked: (tr().dir || (tr().type === 'push' ? 'down' : 'right')) === d, run: () => app.setTransition({ dir: d }) })) },
    ] },
    { label: t('Slides.Group.Timing'), items: [
      { type: 'stack', rows: [
        [{ type: 'number', label: t('Slides.Duration'), tip: t('Slides.DurationTip'), min: 0.1, max: 10, step: 0.1, width: 70, value: () => tr().duration ?? 0.7, change: v => app.setTransition({ duration: Math.max(0.1, Math.min(10, v)) }), disabled: () => !app.slide }, labelNode(t('Slides.Duration'))],
        [{ type: 'toggle', label: t('Slides.AdvanceAfter'), icon: 'clock', showLabel: true, active: () => tr().auto !== undefined, run: () => app.setTransition({ auto: tr().auto !== undefined ? null : 5 }), disabled: () => !app.slide },
          { type: 'number', label: t('Slides.AdvanceSeconds'), min: 0, max: 3600, step: 1, width: 64, value: () => tr().auto ?? '', change: v => app.setTransition({ auto: Math.max(0, v) }), disabled: () => tr().auto === undefined }],
        [{ type: 'button', label: t('Slides.ApplyAll'), icon: 'slides-apply-all', showLabel: true, run: () => app.transitionToAll(), disabled: () => !app.slide }],
      ] },
    ] },
  ];

  // ---- Animations ---------------------------------------------------------------------------------------------------------------------

  const anim = () => (sel()[0] ? app.animationOf(sel()[0].id) : null);
  const animations = [
    { label: t('Slides.Group.Preview'), items: [{ type: 'button', size: 'large', label: t('Slides.Preview'), icon: 'slides-preview', edit: false, run: () => app.previewAnimations(), disabled: () => !app.slide?.animations?.length }] },
    { label: t('Slides.Group.Animation'), items: [
      { type: 'gallery', label: t('Slides.Animations'), items: EFFECTS.map(ef => ({ label: t('Slides.Effect.' + ef), render: () => galleryTile(ef === 'none' ? 'slides-tr-none' : 'slides-anim-' + ef, t('Slides.Effect.' + ef)), active: () => (anim()?.effect || 'none') === ef && !none(), run: () => app.setAnimation(ef) })), disabled: none },
      { type: 'menu', size: 'large', label: t('Slides.EffectOptions'), icon: 'slides-effect-options', disabled: () => !['fly', 'wipe'].includes(anim()?.effect), items: () => ['bottom', 'top', 'left', 'right'].map(d => ({ label: t('Slides.FromSide.' + d), icon: 'slides-dir-' + ({ bottom: 'up', top: 'down', left: 'right', right: 'left' }[d]), checked: (anim()?.dir || 'bottom') === d, run: () => app.editAnimations(a => { a.dir = d; }) })) },
    ] },
    { label: t('Slides.Group.Advanced'), items: [{ type: 'toggle', size: 'large', label: t('Slides.Pane.Animations'), icon: 'slides-animation-pane', edit: false, active: () => app.pane.kind === 'animations', run: () => app.showPane('animations') }] },
    { label: t('Slides.Group.Timing'), items: [{ type: 'stack', rows: [
      [labelNode(t('Slides.Anim.Start')), { type: 'select', label: t('Slides.Anim.Start'), width: 130, options: ['click', 'with', 'after'].map(v => ({ value: v, label: t('Slides.Anim.' + v) })), value: () => anim()?.trigger || 'click', change: v => app.editAnimations(a => { a.trigger = v; }), disabled: () => !anim() }],
      [labelNode(t('Slides.Duration')), { type: 'number', label: t('Slides.Duration'), min: 0.05, max: 20, step: 0.05, width: 70, value: () => anim()?.duration ?? '', change: v => app.editAnimations(a => { a.duration = Math.max(0.05, v); }), disabled: () => !anim() }],
      [labelNode(t('Slides.Delay')), { type: 'number', label: t('Slides.Delay'), min: 0, max: 60, step: 0.25, width: 70, value: () => anim()?.delay ?? (anim() ? 0 : ''), change: v => app.editAnimations(a => { if (v > 0) a.delay = v; else delete a.delay; }), disabled: () => !anim() }],
    ] }] },
  ];

  // ---- Slide Show ---------------------------------------------------------------------------------------------------------------------

  const show = [
    { label: t('Slides.Group.StartShow'), items: [
      { type: 'button', size: 'large', label: t('Slides.Show.FromStart'), icon: 'slides-from-start', key: 'F5', edit: false, run: () => app.startShow(0), disabled: () => !app.content.slides.length },
      { type: 'button', size: 'large', label: t('Slides.Show.FromCurrent'), icon: 'slides-from-current', key: 'Shift+F5', edit: false, run: () => app.startShow(app.index), disabled: () => !app.content.slides.length },
      { type: 'button', size: 'large', label: t('Slides.Show.Presenter'), icon: 'slides-presenter', key: 'Alt+F5', edit: false, run: () => app.startShow(app.index, { presenter: true }), disabled: () => !app.content.slides.length },
    ] },
    { label: t('Slides.Group.SetUp'), items: [
      { type: 'toggle', size: 'large', label: t('Slides.HideSlide'), icon: 'slides-hide', active: () => !!app.slide?.hidden, run: () => app.toggleHidden(), disabled: () => !app.slide },
    ] },
  ];

  // ---- View -------------------------------------------------------------------------------------------------------------------------

  const view = [
    { label: t('Slides.Group.Views'), items: [
      { type: 'toggle', size: 'large', label: t('Slides.View.Normal'), icon: 'slides-normal', edit: false, active: () => app.view.mode === 'normal', run: () => { app.setView('normal'); } },
      { type: 'toggle', size: 'large', label: t('Slides.View.Sorter'), icon: 'slides-sorter', edit: false, active: () => app.view.mode === 'sorter', run: () => app.setView('sorter') },
      { type: 'button', size: 'large', label: t('Slides.View.Reading'), icon: 'slides-reading', edit: false, run: () => app.startShow(app.index, { reading: true }), disabled: () => !app.content.slides.length },
    ] },
    { label: t('Slides.Group.Show'), items: [{ type: 'stack', rows: [
      [{ type: 'toggle', label: t('Slides.View.Gridlines'), icon: 'slides-gridlines', showLabel: true, edit: false, active: () => !!app.view.grid, run: () => { app.view.grid = !app.view.grid; app.layoutChanged(); } }],
      [{ type: 'toggle', label: t('Slides.View.Guides'), icon: 'slides-guides', showLabel: true, edit: false, active: () => !!app.view.guides, run: () => { app.view.guides = !app.view.guides; app.layoutChanged(); } }],
      [{ type: 'toggle', label: t('Slides.View.Notes'), icon: 'slides-notes', showLabel: true, edit: false, active: () => !!app.view.notes, run: () => { app.toggleNotes(); app.statusBar(); } }],
    ] }] },
    { label: t('Slides.Group.Snap'), items: [{ type: 'stack', rows: [
      [{ type: 'toggle', label: t('Slides.View.SmartGuides'), icon: 'slides-smart-guides', showLabel: true, edit: false, active: () => !!app.view.snap, run: () => { app.view.snap = !app.view.snap; app.layoutChanged(); } }],
      [{ type: 'toggle', label: t('Slides.View.SnapGrid'), icon: 'grid', showLabel: true, edit: false, active: () => !!app.view.gridSnap, run: () => { app.view.gridSnap = !app.view.gridSnap; app.layoutChanged(); } }],
    ] }] },
    { label: t('Slides.Group.Zoom'), items: [
      { type: 'menu', size: 'large', label: t('Slides.Zoom'), icon: 'zoom-in', edit: false, items: () => [0.33, 0.5, 0.66, 1, 1.5, 2].map(z => ({ label: Math.round(z * 100) + ' %', checked: Math.abs(app.canvas.zoom - z) < 0.01, run: () => app.canvas.setZoom(z) })) },
      { type: 'button', size: 'large', label: t('Slides.FitWindow'), icon: 'zoom-fit', edit: false, run: () => app.canvas.fitZoom() },
    ] },
  ];

  // ---- contextual tabs ------------------------------------------------------------------------------------------------------------

  const one = () => sel().length === 1 ? sel()[0] : null;
  const sizeGroup = { label: t('Slides.Group.Size'), items: [{ type: 'stack', rows: [
    [labelNode(t('Slides.Size.Height')), { type: 'number', label: t('Slides.Size.Height'), min: 0, max: 20000, step: 1, width: 76, value: () => (one() ? Math.round(one().h * 10) / 10 : ''), change: v => one() && app.setGeometry(one().id, { h: Math.max(0, v) }), disabled: () => !one() }],
    [labelNode(t('Slides.Size.Width')), { type: 'number', label: t('Slides.Size.Width'), min: 1, max: 20000, step: 1, width: 76, value: () => (one() ? Math.round(one().w * 10) / 10 : ''), change: v => one() && app.setGeometry(one().id, { w: Math.max(1, v) }), disabled: () => !one() }],
  ] }] };
  const arrangeGroup = { label: t('Slides.Group.Arrange'), items: [{ type: 'stack', rows: [
    [{ type: 'button', label: t('Slides.BringForward'), icon: 'slides-bring-forward', showLabel: true, run: () => app.arrange('forward'), disabled: none }],
    [{ type: 'button', label: t('Slides.SendBackward'), icon: 'slides-send-backward', showLabel: true, run: () => app.arrange('backward'), disabled: none }],
    [{ type: 'menu', label: t('Slides.Align'), icon: 'slides-align-objects', disabled: none, items: () => alignItems(app) },
      { type: 'menu', label: t('Slides.Rotate'), icon: 'slides-rotate', disabled: none, items: () => rotateItems(app) },
      { type: 'button', label: t('Slides.Group'), icon: 'slides-group', run: () => app.group(), disabled: () => sel().length < 2 },
      { type: 'button', label: t('Slides.Ungroup'), icon: 'slides-ungroup', run: () => app.ungroup(), disabled: () => !sel().some(e => e.group) }],
  ] }] };

  const shapeFormat = { id: 'shape', label: t('Slides.Tab.ShapeFormat'), contextual: true, hidden: () => !sel().some(e => e.type === 'shape' || e.type === 'text' || e.type === 'line'), groups: [
    { label: t('Slides.Group.InsertShapes'), items: [
      custom(() => anchorBtnLarge(t('Slides.Shapes'), 'slides-shapes', b => shapeGallery(b, kind => app.startDrawing(kind)), ro)),
      custom(() => anchorBtnLarge(t('Slides.ChangeShape'), 'slides-change-shape', b => shapeGallery(b, kind => { if (!['line', 'arrow', 'doubleArrow'].includes(kind)) app.changeShape(kind); }), () => ro() || !app.shapeTargets().length)),
    ] },
    { label: t('Slides.Group.ShapeStyles'), items: [
      custom(() => anchorBtnLarge(t('Slides.QuickStyles'), 'slides-quick-styles', b => quickStyles(b, theme(), s => app.quickStyle(s)), () => ro() || !app.shapeTargets().length)),
      { type: 'stack', rows: [
        [fill(t('Slides.ShapeFill'), 'slides-fill', { none: t('Slides.NoFill'), value: () => sel()[0]?.fill?.color || 'accent1', change: c => app.setFill(c || 'accent1'), enabled: () => app.shapeTargets().length > 0 }), labelNode(t('Slides.ShapeFill'))],
        [fill(t('Slides.ShapeOutline'), 'slides-outline', { none: t('Slides.NoOutline'), value: () => sel()[0]?.line?.color || 'dk1', change: c => app.setLine(c === 'none' ? null : { color: c || 'dk1' }), enabled: () => sel().some(e => e.type !== 'table') }), labelNode(t('Slides.ShapeOutline'))],
        [{ type: 'menu', label: t('Slides.Weight'), icon: 'slides-line-weight', disabled: () => !sel().some(e => e.type !== 'table'), items: () => [0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map(w => ({ label: String(w).replace('.', t('Slides.DecimalSep')) + ' pt', checked: sel()[0]?.line?.width === w, run: () => app.setLine({ width: w }) })) },
          { type: 'menu', label: t('Slides.Line.Dash'), icon: 'slides-dash', disabled: () => !sel().some(e => e.type !== 'table'), items: () => ['solid', 'dash', 'dot', 'dashDot', 'longDash'].map(d => ({ label: t('Slides.Dash.' + d), checked: sel()[0]?.line?.dash === d, run: () => app.setLine({ dash: d }) })) },
          { type: 'menu', label: t('Slides.ShapeEffects'), icon: 'slides-effects', disabled: () => !sel().some(e => e.type !== 'table'), items: () => effectItems(app) }],
      ] },
      { type: 'button', size: 'large', label: t('Slides.Pane.Shape'), icon: 'slides-format', edit: false, run: () => app.showPane('format') },
    ] },
    arrangeGroup, sizeGroup,
  ] };

  const pictureFormat = { id: 'picture', label: t('Slides.Tab.PictureFormat'), contextual: true, hidden: () => !sel().some(e => e.type === 'image'), groups: [
    { label: t('Slides.Group.Adjust'), items: [
      { type: 'menu', size: 'large', label: t('Slides.Transparency'), icon: 'slides-transparency', items: () => [0, 15, 30, 50, 65, 80].map(v => ({ label: v + ' %', checked: Math.round((1 - (sel()[0]?.opacity ?? 1)) * 100) === v, run: () => app.editElements(sel().filter(e => e.type === 'image').map(e => e.id), e => { if (v) e.opacity = 1 - v / 100; else delete e.opacity; }) })) },
      { type: 'button', size: 'large', label: t('Slides.ChangePicture'), icon: 'slides-change-picture', run: async () => { const p = sel().find(e => e.type === 'image'); if (!p) return; const src = await app.pane.pickPicture(); if (src) app.editElements([p.id], e => { e.src = src; delete e.crop; }); } },
      { type: 'button', size: 'large', label: t('Slides.ResetPicture'), icon: 'slides-reset-picture', run: () => app.editElements(sel().filter(e => e.type === 'image').map(e => e.id), e => { delete e.crop; delete e.opacity; delete e.line; delete e.shadow; delete e.rot; delete e.flipH; delete e.flipV; }) },
    ] },
    { label: t('Slides.Group.PictureStyles'), items: [{ type: 'stack', rows: [
      [fill(t('Slides.PictureBorder'), 'slides-outline', { none: t('Slides.NoOutline'), value: () => sel()[0]?.line?.color || 'dk1', change: c => app.setLine(c === 'none' ? null : { color: c || 'dk1' }), enabled: () => true }), labelNode(t('Slides.PictureBorder'))],
      [{ type: 'toggle', label: t('Slides.Shadow'), icon: 'slides-effects', showLabel: true, active: () => !!sel()[0]?.shadow, run: () => app.toggleShadow() }],
      [{ type: 'button', label: t('Slides.AltText'), icon: 'slides-alt-text', showLabel: true, edit: false, run: () => app.showPane('format') }],
    ] }] },
    { label: t('Slides.Group.Crop'), items: [{ type: 'button', size: 'large', label: t('Slides.Crop'), icon: 'slides-crop', run: () => app.showPane('format') }] },
    arrangeGroup, sizeGroup,
  ] };

  const tb = () => app.table();
  const table = { id: 'table', label: t('Slides.Tab.Table'), contextual: true, hidden: () => !(sel().length === 1 && sel()[0].type === 'table'), groups: [
    { label: t('Slides.Group.TableOptions'), items: [{ type: 'stack', rows: [
      [{ type: 'toggle', label: t('Slides.Table.Header'), icon: 'slides-table-header', showLabel: true, active: () => tb()?.header !== false, run: () => app.setTableProp({ header: !(tb()?.header !== false) }) }],
      [{ type: 'toggle', label: t('Slides.Table.Total'), icon: 'slides-table-total', showLabel: true, active: () => !!tb()?.lastRow, run: () => app.setTableProp({ lastRow: !tb()?.lastRow }) }],
      [{ type: 'toggle', label: t('Slides.Table.Banded'), icon: 'slides-table-banded', showLabel: true, active: () => tb()?.banded !== false, run: () => app.setTableProp({ banded: !(tb()?.banded !== false) }) }],
    ] }, { type: 'stack', rows: [
      [{ type: 'toggle', label: t('Slides.Table.FirstCol'), icon: 'slides-table-first-col', showLabel: true, active: () => !!tb()?.firstCol, run: () => app.setTableProp({ firstCol: !tb()?.firstCol }) }],
    ] }] },
    { label: t('Slides.Group.TableStyles'), items: [
      { type: 'gallery', label: t('Slides.Table.Styles'), items: ['medium', 'light', 'dark', 'none'].flatMap(kind => ['accent1', 'accent2', 'accent4', 'accent6'].map(acc => ({ label: t('Slides.Table.Style.' + kind) + ' · ' + t('Slides.Color.' + acc), render: () => tablePic(app, kind, acc), active: () => tb()?.tableStyle === kind && (tb()?.accent || 'accent1') === acc, run: () => app.setTableProp({ tableStyle: kind, accent: acc }) }))) },
      fill(t('Slides.Table.Shading'), 'slides-fill', { none: t('Slides.NoFill'), auto: t('Slides.Table.StyleFill'), value: () => 'accent2:+60', change: c => app.cellFill(c), enabled: () => !!tb() }),
    ] },
    { label: t('Slides.Group.RowsCols'), items: [
      { type: 'stack', rows: [
        [{ type: 'button', label: t('Slides.Table.RowAbove'), icon: 'slides-row-above', showLabel: true, run: () => app.insertRow(false) }],
        [{ type: 'button', label: t('Slides.Table.RowBelow'), icon: 'slides-row-below', showLabel: true, run: () => app.insertRow(true) }],
        [{ type: 'menu', label: t('Slides.Table.Delete'), icon: 'slides-table-delete', items: () => [
          { label: t('Slides.Table.DeleteRow'), icon: 'slides-delete-row', run: () => app.deleteRow() },
          { label: t('Slides.Table.DeleteCol'), icon: 'slides-delete-col', run: () => app.deleteCol() },
          { label: t('Slides.Table.DeleteTable'), icon: 'trash', danger: true, run: () => app.deleteTable() },
        ] }],
      ] },
      { type: 'stack', rows: [
        [{ type: 'button', label: t('Slides.Table.ColLeft'), icon: 'slides-col-left', showLabel: true, run: () => app.insertCol(false) }],
        [{ type: 'button', label: t('Slides.Table.ColRight'), icon: 'slides-col-right', showLabel: true, run: () => app.insertCol(true) }],
      ] },
      { type: 'stack', rows: [
        [{ type: 'button', label: t('Slides.Table.DistRows'), icon: 'slides-distribute-rows', showLabel: true, run: () => app.distributeRows() }],
        [{ type: 'button', label: t('Slides.Table.DistCols'), icon: 'slides-distribute-cols', showLabel: true, run: () => app.distributeCols() }],
      ] },
    ] },
    { label: t('Slides.Group.Alignment'), items: [{ type: 'stack', rows: [
      [{ type: 'toggle', label: t('Slides.AlignLeft'), icon: 'slides-align-left', active: () => f().align === 'left', run: () => app.setAlign('left') },
        { type: 'toggle', label: t('Slides.AlignCenter'), icon: 'slides-align-center', active: () => f().align === 'center', run: () => app.setAlign('center') },
        { type: 'toggle', label: t('Slides.AlignRight'), icon: 'slides-align-right', active: () => f().align === 'right', run: () => app.setAlign('right') }],
      [...['top', 'middle', 'bottom'].map(a => ({ type: 'toggle', label: t('Slides.Anchor.' + a), icon: 'slides-anchor-' + a, active: () => f().anchor === a, run: () => app.setBodyProp('anchor', a) }))],
    ] }] },
    sizeGroup,
  ] };

  return wrapMenus([
    { id: 'home', label: t('Slides.Tab.Home'), groups: [clipboard, slidesGroup, fontGroup, paragraph, drawing, editing] },
    { id: 'insert', label: t('Slides.Tab.Insert'), groups: insert },
    { id: 'design', label: t('Slides.Tab.Design'), groups: design },
    { id: 'transitions', label: t('Slides.Tab.Transitions'), groups: transitions },
    { id: 'animations', label: t('Slides.Tab.Animations'), groups: animations },
    { id: 'show', label: t('Slides.Tab.SlideShow'), groups: show },
    { id: 'view', label: t('Slides.Tab.View'), groups: view },
    shapeFormat, pictureFormat, table,
    { id: 'undo', label: '', hidden: () => true, groups: [{ label: '', items: [
      { type: 'button', label: t('Slides.Undo'), icon: 'undo', key: 'Mod+Z', global: false, run: () => app.undo() },
      { type: 'button', label: t('Slides.Redo'), icon: 'redo', key: 'Mod+Y', global: false, run: () => app.redo() },
      { type: 'button', label: t('Slides.Redo'), icon: 'redo', key: 'Mod+Shift+Z', global: false, run: () => app.redo() },
      { type: 'button', label: t('Slides.SelectAll'), icon: 'cursor', key: 'Mod+A', global: false, edit: false, run: () => app.selectAll() },
      { type: 'button', label: t('Slides.Duplicate'), icon: 'copy', key: 'Mod+D', run: () => app.duplicate() },
      { type: 'button', label: t('Slides.Group'), icon: 'copy', key: 'Mod+G', global: false, run: () => app.group() },
      { type: 'button', label: t('Slides.Ungroup'), icon: 'copy', key: 'Mod+Shift+G', global: false, run: () => app.ungroup() },
    ] }] },
  ]);
}

/** The menus of the frame call checked() and disabled(): plain values become functions. */
function wrapMenus(tabs) {
  const fn = v => (typeof v === 'function' || v === undefined ? v : () => v);
  const fix = list => list.map(x => (x && typeof x === 'object' ? { ...x, checked: fn(x.checked), disabled: fn(x.disabled) } : x));
  const walk = items => {
    for (const item of items || []) {
      if (item.type === 'stack') item.rows.forEach(walk);
      if ((item.type === 'menu' || item.type === 'split') && item.items) {
        const src = item.items;
        item.items = typeof src === 'function' ? () => fix(src()) : fix(src);
      }
    }
  };
  for (const tab of tabs) for (const g of tab.groups) walk(g.items);
  return tabs;
}

function fontOptions(app) {
  const theme = app.content.theme;
  const list = [...new Set([theme.fonts.major, theme.fonts.minor, ...FONTS])];
  const cur = app.fmt().font;
  if (cur && !list.includes(cur)) list.unshift(cur);
  return list.map(v => ({ value: v, label: v === theme.fonts.major ? `${v} (${t('Slides.Font.Headings')})` : v === theme.fonts.minor ? `${v} (${t('Slides.Font.Body')})` : v }));
}

function numberingLabel(n) { return n; }

function labelNode(text) { return { type: 'custom', render: () => el('span', { class: 'mo-slides-rlabel', text }) }; }

function galleryTile(iconName, label) { return el('span', { class: 'mo-slides-tile' }, icon(iconName), el('span', { class: 'mo-slides-tile__label', text: label })); }

function tablePic(app, kind, accent) {
  const theme = app.content.theme;
  const fake = { tableStyle: kind, accent, header: true, banded: true, rows: Array(4).fill(0) };
  const box = el('span', { class: 'mo-slides-tablepic' });
  for (let r = 0; r < 4; r++) {
    const row = el('span', { class: 'mo-slides-tablepic__row' });
    for (let c = 0; c < 3; c++) {
      const look = tableLook(fake, r, c);
      const cell = el('span', { class: 'mo-slides-tablepic__cell' });
      cell.style.background = look.fill ? resolveColor(look.fill, theme) : 'transparent';
      cell.style.borderColor = look.border ? resolveColor(look.border, theme) : 'transparent';
      row.append(cell);
    }
    box.append(row);
  }
  return box;
}

/** The designs: a strip in the ribbon and, like the office suites, an arrow that opens all of them as a grid. */
function themeGallery(app) {
  const tile = (th, width, close) => {
    const label = t('Slides.Theme.' + th.id);
    const b = el('button', { type: 'button', class: 'mo-rgallery__item', role: 'option', 'aria-label': label, 'data-tip': label, 'data-theme': th.id },
      themePreview(app.content, th.id, 0, width));
    b.addEventListener('click', () => { if (app.readOnly) return; close?.(); app.setTheme(th.id); });
    return b;
  };
  return {
    type: 'custom',
    render: () => {
      const strip = el('div', { class: 'mo-rgallery mo-slides-themestrip', role: 'listbox', 'aria-label': t('Slides.Themes') }, ...THEMES.map(th => tile(th, 96)));
      const more = el('button', { type: 'button', class: 'mo-rbtn mo-slides-gallerymore', 'aria-label': t('Slides.Themes.More'), 'data-tip': t('Slides.Themes.More'), 'aria-haspopup': 'true' }, icon('chevron-down'));
      more.addEventListener('click', () => openPopover(more, close => el('div', { class: 'mo-slides-gallery' },
        el('div', { class: 'mo-menu__heading', text: t('Slides.Themes') }),
        el('div', { class: 'mo-slides-gallery__themes', role: 'listbox', 'aria-label': t('Slides.Themes') }, ...THEMES.map(th => {
          const b = tile(th, 116, close);
          b.setAttribute('aria-selected', app.content.theme.id === th.id ? 'true' : 'false');
          return b;
        })))));
      const node = el('div', { class: 'mo-slides-galwrap' }, strip, more);
      node.refresh = () => {
        for (const b of strip.children) {
          const on = app.content.theme.id === b.dataset.theme;
          b.disabled = app.readOnly;
          b.setAttribute('aria-selected', on ? 'true' : 'false');
          if (on && (b.offsetLeft < strip.scrollLeft || b.offsetLeft + b.offsetWidth > strip.scrollLeft + strip.clientWidth)) strip.scrollLeft = b.offsetLeft - 4;
        }
        more.disabled = app.readOnly;
      };
      return node;
    },
    refresh: node => node.refresh(),
  };
}

/** A large button that opens a popover at itself (shapes, table, styles). */
function anchorBtnLarge(label, iconName, open, disabled) {
  const b = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--lg mo-rbtn--menu', 'aria-haspopup': 'true' }, icon(iconName, 'mo-rbtn__icon--lg'), el('span', { class: 'mo-rbtn__label', text: label }), icon('chevron-down', 'mo-rbtn__chev'));
  b.addEventListener('click', () => { if (!disabled()) open(b); });
  b.refresh = () => { b.disabled = disabled(); };
  return b;
}

/** A small button with its label that opens a popover. */
function anchorBtnLabel(label, iconName, open, disabled) {
  const b = el('button', { type: 'button', class: 'mo-rbtn mo-rbtn--menu', 'aria-haspopup': 'true' }, icon(iconName), el('span', { class: 'mo-rbtn__label', text: label }), icon('chevron-down', 'mo-rbtn__chev'));
  b.addEventListener('click', () => { if (!disabled()) open(b); });
  b.refresh = () => { b.disabled = disabled(); };
  return b;
}

function arrangeItems(app) {
  return [
    { heading: t('Slides.OrderObjects') },
    { label: t('Slides.BringFront'), icon: 'slides-bring-front', run: () => app.arrange('front') },
    { label: t('Slides.SendBack'), icon: 'slides-send-back', run: () => app.arrange('back') },
    { label: t('Slides.BringForward'), icon: 'slides-bring-forward', run: () => app.arrange('forward') },
    { label: t('Slides.SendBackward'), icon: 'slides-send-backward', run: () => app.arrange('backward') },
    { heading: t('Slides.GroupObjects') },
    { label: t('Slides.Group'), icon: 'slides-group', key: 'Ctrl+G', run: () => app.group(), disabled: app.canvas.selected().length < 2 },
    { label: t('Slides.Ungroup'), icon: 'slides-ungroup', key: 'Ctrl+Shift+G', run: () => app.ungroup(), disabled: !app.canvas.selected().some(e => e.group) },
    { heading: t('Slides.PositionObjects') },
    ...alignItems(app).filter(x => x !== '-' && !x.checked),
    '-',
    ...rotateItems(app),
  ];
}

function alignItems(app) {
  const one = app.canvas.units().length === 1;
  return [
    { label: t('Slides.AlignObj.left'), icon: 'slides-obj-left', run: () => app.align('left') },
    { label: t('Slides.AlignObj.center'), icon: 'slides-obj-center', run: () => app.align('center') },
    { label: t('Slides.AlignObj.right'), icon: 'slides-obj-right', run: () => app.align('right') },
    { label: t('Slides.AlignObj.top'), icon: 'slides-obj-top', run: () => app.align('top') },
    { label: t('Slides.AlignObj.middle'), icon: 'slides-obj-middle', run: () => app.align('middle') },
    { label: t('Slides.AlignObj.bottom'), icon: 'slides-obj-bottom', run: () => app.align('bottom') },
    '-',
    { label: t('Slides.DistributeH'), icon: 'slides-distribute-h', run: () => app.distribute('h') },
    { label: t('Slides.DistributeV'), icon: 'slides-distribute-v', run: () => app.distribute('v') },
    '-',
    { label: one ? t('Slides.AlignToSlide') : t('Slides.AlignToSelection'), icon: 'info', checked: true, disabled: true },
  ];
}

function rotateItems(app) {
  return [
    { label: t('Slides.RotateRight'), icon: 'slides-rotate', run: () => app.rotateBy(90) },
    { label: t('Slides.RotateLeft'), icon: 'slides-rotate-left', run: () => app.rotateBy(-90) },
    { label: t('Slides.FlipV'), icon: 'slides-flip-v', run: () => app.flip('v') },
    { label: t('Slides.FlipH'), icon: 'slides-flip-h', run: () => app.flip('h') },
  ];
}

function effectItems(app) {
  const sel = app.canvas.selected();
  return [
    { label: t('Slides.Shadow'), icon: 'slides-effects', checked: sel.length > 0 && sel.every(e => e.shadow), run: () => app.toggleShadow() },
    '-',
    { heading: t('Slides.Corners') },
    ...[0, 0.08, 0.17, 0.3, 0.5].map(v => ({ label: v === 0 ? t('Slides.Corners.Square') : Math.round(v * 100) + ' %', checked: sel[0]?.shape === 'roundRect' && Math.abs((sel[0].adj ?? 0.1667) - v) < 0.02, disabled: !sel.some(e => e.shape === 'rect' || e.shape === 'roundRect'), run: () => app.setRadius(v) })),
  ];
}

