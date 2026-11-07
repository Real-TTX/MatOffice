/* MatOffice Sheets – the ribbon, like Excel's: Home, Insert, Page Layout, Formulas, Data, Review, View, and the tabs of a selected chart or
   picture. Built from the declarative items of the shared frame (docs/EDITORS.md). */
import { el, t } from '../editor/util.js';
import { FUNCTIONS, CATEGORIES } from './engine/functions/index.js';
import { presets } from './engine/numfmt.js';
import { FONTS, fontFamily } from './look.js';
import { CHART_TYPES, CHART_COLORS } from './charts.js';

/** The cell styles of the gallery (like Excel's). */
export const CELL_STYLES = [
  { id: 'normal', style: {} },
  { id: 'good', style: { fill: '#c6efce', color: '#006100' } },
  { id: 'bad', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { id: 'neutral', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { id: 'calculation', style: { fill: '#f2f2f2', color: '#fa7d00', b: true, bt: 'thin #7f7f7f', bb: 'thin #7f7f7f', bl: 'thin #7f7f7f', br: 'thin #7f7f7f' } },
  { id: 'check', style: { fill: '#a5a5a5', color: '#ffffff', b: true, bt: 'double #3f3f3f', bb: 'double #3f3f3f', bl: 'double #3f3f3f', br: 'double #3f3f3f' } },
  { id: 'explanatory', style: { color: '#7f7f7f', i: true } },
  { id: 'input', style: { fill: '#ffcc99', color: '#3f3f76', bt: 'thin #7f7f7f', bb: 'thin #7f7f7f', bl: 'thin #7f7f7f', br: 'thin #7f7f7f' } },
  { id: 'linked', style: { color: '#fa7d00', bb: 'double #ff8001' } },
  { id: 'note', style: { fill: '#ffffcc', bt: 'thin #b2b2b2', bb: 'thin #b2b2b2', bl: 'thin #b2b2b2', br: 'thin #b2b2b2' } },
  { id: 'output', style: { fill: '#f2f2f2', color: '#3f3f3f', b: true, bt: 'thin #3f3f3f', bb: 'thin #3f3f3f', bl: 'thin #3f3f3f', br: 'thin #3f3f3f' } },
  { id: 'warning', style: { color: '#ff0000' } },
  { id: 'title', style: { size: 18, color: '#44546a', font: 'Calibri Light' } },
  { id: 'heading1', style: { size: 15, b: true, color: '#44546a', bb: 'thick #4472c4' } },
  { id: 'heading2', style: { size: 13, b: true, color: '#44546a', bb: 'thick #a2b8e1' } },
  { id: 'heading3', style: { b: true, color: '#44546a', bb: 'medium #8ea9db' } },
  { id: 'heading4', style: { b: true, color: '#44546a' } },
  { id: 'total', style: { b: true, bt: 'thin #4472c4', bb: 'double #4472c4' } },
  { id: 'accent1', style: { fill: '#4472c4', color: '#ffffff' } },
  { id: 'accent2', style: { fill: '#ed7d31', color: '#ffffff' } },
  { id: 'accent3', style: { fill: '#a5a5a5', color: '#ffffff' } },
  { id: 'accent4', style: { fill: '#ffc000', color: '#000000' } },
  { id: 'accent5', style: { fill: '#5b9bd5', color: '#ffffff' } },
  { id: 'accent6', style: { fill: '#70ad47', color: '#ffffff' } },
  { id: 'accent1Light', style: { fill: '#d9e1f2' } },
  { id: 'accent6Light', style: { fill: '#e2efda' } },
];

const SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72'];

export function ribbonTabs(app) {
  const cmd = app.commands;
  const dlg = app.dialogs;
  const clip = app.clipboard;
  const style = () => cmd.style;
  const p = presets(app.locale);
  const chart = () => app.objects.selectedObject && app.objects.kindOf(app.objects.selectedObject) === 'chart' ? app.objects.selectedObject : null;
  const image = () => app.objects.selectedObject && app.objects.kindOf(app.objects.selectedObject) === 'image' ? app.objects.selectedObject : null;
  const insertFn = name => app.insertFunctionName(name);
  const fnMenu = cat => () => [...FUNCTIONS.values()].filter(f => f.category === cat).map(f => f.name).sort().map(name => ({ label: name, run: () => insertFn(name) }));
  const borderItems = [
    { label: t('Sheets.Border.Bottom'), icon: 'sheets-border-bottom', run: () => cmd.borders('bottom') },
    { label: t('Sheets.Border.Top'), icon: 'sheets-border-top', run: () => cmd.borders('top') },
    { label: t('Sheets.Border.Left'), icon: 'sheets-border-left', run: () => cmd.borders('left') },
    { label: t('Sheets.Border.Right'), icon: 'sheets-border-right', run: () => cmd.borders('right') },
    '-',
    { label: t('Sheets.Border.None'), icon: 'sheets-border-none', run: () => cmd.borders('none') },
    { label: t('Sheets.Border.All'), icon: 'sheets-border-all', run: () => cmd.borders('all') },
    { label: t('Sheets.Border.Outside'), icon: 'sheets-border-outside', key: 'Mod+Shift+&', run: () => cmd.borders('outside') },
    { label: t('Sheets.Border.ThickOutside'), icon: 'sheets-border-thick', run: () => cmd.borders('thickOutside') },
    { label: t('Sheets.Border.Inside'), icon: 'sheets-border-inside', run: () => cmd.borders('inside') },
    '-',
    { label: t('Sheets.Border.DoubleBottom'), icon: 'sheets-border-double', run: () => cmd.borders('doubleBottom') },
    { label: t('Sheets.Border.ThickBottom'), icon: 'sheets-border-bottom', run: () => cmd.borders('thickBottom') },
    { label: t('Sheets.Border.TopBottom'), icon: 'sheets-border-topbottom', run: () => cmd.borders('topBottom') },
    '-',
    { heading: t('Sheets.Border.LineStyle') },
    ...['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].map(s => ({ label: t('Sheets.Line.' + s), checked: () => cmd.borderStyle === s, run: () => { cmd.borderStyle = s; } })),
    { heading: t('Sheets.Border.Color') },
    ...[['#000000', 'black'], ['#7f7f7f', 'gray'], ['#4472c4', 'blue'], ['#c00000', 'red'], ['#107c41', 'green']].map(([c, n]) => ({
      label: t('Sheets.Color.' + n), render: () => { const s = el('span', { class: 'mo-sheets-swatch' }); s.style.background = c; return s; }, checked: () => cmd.borderColor === c, run: () => { cmd.borderColor = c; },
    })),
  ];
  const styleItem = s => ({
    label: t('Sheets.Style.' + s.id),
    render: () => { const n = el('span', { class: 'mo-sheets-stylechip', text: 'Aa' }); applyCss(n, s.style); return n; },
    run: () => cmd.cellStyle(s.style),
  });
  const cfMenu = [
    { heading: t('Sheets.Cf.Highlight') },
    { label: t('Sheets.Cf.Title.gt'), icon: 'sheets-cf-gt', run: () => dlg.cfRule('gt') },
    { label: t('Sheets.Cf.Title.lt'), icon: 'sheets-cf-lt', run: () => dlg.cfRule('lt') },
    { label: t('Sheets.Cf.Title.between'), icon: 'sheets-cf-between', run: () => dlg.cfRule('between') },
    { label: t('Sheets.Cf.Title.eq'), icon: 'sheets-cf-eq', run: () => dlg.cfRule('eq') },
    { label: t('Sheets.Cf.Title.text'), icon: 'sheets-cf-text', run: () => dlg.cfRule('text') },
    { label: t('Sheets.Cf.Title.dup'), icon: 'sheets-cf-dup', run: () => dlg.cfRule('dup') },
    { heading: t('Sheets.Cf.TopBottom') },
    { label: t('Sheets.Cf.Title.top'), icon: 'sheets-cf-top', run: () => dlg.cfRule('top') },
    { label: t('Sheets.Cf.Title.topPercent'), icon: 'sheets-cf-top', run: () => dlg.cfRule('topPercent') },
    { label: t('Sheets.Cf.Title.bottom'), icon: 'sheets-cf-bottom', run: () => dlg.cfRule('bottom') },
    { label: t('Sheets.Cf.Title.bottomPercent'), icon: 'sheets-cf-bottom', run: () => dlg.cfRule('bottomPercent') },
    { label: t('Sheets.Cf.Title.above'), icon: 'sheets-cf-above', run: () => dlg.cfRule('above') },
    { label: t('Sheets.Cf.Title.below'), icon: 'sheets-cf-below', run: () => dlg.cfRule('below') },
    { heading: t('Sheets.Cf.Bars') },
    ...[['#638ec6', 'blue'], ['#63be7b', 'green'], ['#ff555a', 'red'], ['#ffb628', 'orange'], ['#008aef', 'lightBlue'], ['#d6007b', 'purple']].map(([c, n]) => ({
      label: t('Sheets.Cf.Bar', t('Sheets.Color.' + n)), render: () => { const s = el('span', { class: 'mo-sheets-swatch mo-sheets-swatch--bar' }); s.style.background = `linear-gradient(90deg, ${c}, ${c}33)`; return s; }, run: () => dlg.bar(c),
    })),
    { heading: t('Sheets.Cf.Scales') },
    ...[['#f8696b', '#ffeb84', '#63be7b'], ['#63be7b', '#ffeb84', '#f8696b'], ['#f8696b', '#ffffff', '#63be7b'], ['#5a8ac6', '#ffffff', '#f8696b'], ['#ffffff', '#63be7b'], ['#ffffff', '#f8696b']].map((cs, i) => ({
      label: t('Sheets.Cf.Scale', i + 1), render: () => { const s = el('span', { class: 'mo-sheets-swatch mo-sheets-swatch--bar' }); s.style.background = `linear-gradient(90deg, ${cs.join(', ')})`; return s; }, run: () => dlg.scale(cs),
    })),
    '-',
    { label: t('Sheets.Cf.NewRule'), icon: 'plus', run: () => dlg.cfFormula() },
    { label: t('Sheets.Cf.ClearSelection'), icon: 'broom', run: () => dlg.clearRules(false) },
    { label: t('Sheets.Cf.ClearSheet'), icon: 'broom', run: () => dlg.clearRules(true) },
    { label: t('Sheets.Cf.Manage'), icon: 'sheets-cf', run: () => dlg.manageRules() },
  ];
  const numberOptions = () => [
    ['general', 'Sheets.Nf.general'], ['number', 'Sheets.Nf.number'], ['currency', 'Sheets.Nf.currency'], ['accounting', 'Sheets.Nf.accounting'],
    ['shortDate', 'Sheets.Nf.shortDate'], ['longDate', 'Sheets.Nf.longDate'], ['time', 'Sheets.Nf.time'], ['percent', 'Sheets.Nf.percent'],
    ['fraction', 'Sheets.Nf.fraction'], ['scientific', 'Sheets.Nf.scientific'], ['text', 'Sheets.Nf.text'], ['custom', 'Sheets.Nf.custom'],
  ].map(([value, key]) => ({ value, label: t(key) }));

  const tabs = [
    { id: 'home', label: t('Sheets.Tab.Home'), groups: [
      { label: t('Sheets.Group.Undo'), items: [{ type: 'stack', rows: [
        [{ type: 'button', label: t('Sheets.Undo'), icon: 'undo', key: 'Mod+Z', global: false, disabled: () => !app.wb.canUndo, run: () => app.undo() }],
        [{ type: 'button', label: t('Sheets.Redo'), icon: 'redo', key: 'Mod+Y', global: false, disabled: () => !app.wb.canRedo, run: () => app.redo() }],
      ] }] },
      { label: t('Sheets.Group.Clipboard'), items: [
        { type: 'split', size: 'large', label: t('Sheets.Paste'), icon: 'sheets-paste', run: () => clip.pasteCommand(), items: [
          { label: t('Sheets.Paste'), icon: 'sheets-paste', run: () => clip.pasteCommand() },
          { label: t('Sheets.PasteMode.values'), icon: 'sheets-paste-values', run: () => clip.pasteCommand('values') },
          { label: t('Sheets.PasteMode.formulas'), icon: 'sheets-function', run: () => clip.pasteCommand('formulas') },
          { label: t('Sheets.PasteMode.formats'), icon: 'sheets-painter', run: () => clip.pasteCommand('formats') },
          { label: t('Sheets.PasteMode.Transpose'), icon: 'sheets-transpose', run: () => clip.pasteCommand('transpose') },
          '-',
          { label: t('Sheets.PasteSpecial'), icon: 'sheets-paste-special', run: () => dlg.pasteSpecial() },
        ] },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.Cut'), icon: 'sheets-cut', showLabel: true, run: () => clip.cutCommand() }],
          [{ type: 'button', label: t('Sheets.Copy'), icon: 'copy', showLabel: true, edit: false, run: () => clip.copyCommand() }],
          [{ type: 'toggle', label: t('Sheets.Painter'), icon: 'sheets-painter', showLabel: true, active: () => !!app.painter, run: () => (app.painter ? cmd.stopPainter() : cmd.startPainter(false)) }],
        ] },
      ] },
      { label: t('Sheets.Group.Font'), items: [{ type: 'stack', rows: [
        [
          { type: 'select', label: t('Sheets.Font'), width: 132, options: () => FONTS.map(f => ({ value: f, label: f })), value: () => style().font || 'Calibri', change: v => cmd.setFont(v) },
          { type: 'combo', label: t('Sheets.Size'), width: 52, options: SIZES, inputmode: 'decimal', value: () => String(style().size || 11), change: v => cmd.setSize(v) },
          { type: 'button', label: t('Sheets.GrowFont'), icon: 'sheets-font-grow', run: () => cmd.growFont(1) },
          { type: 'button', label: t('Sheets.ShrinkFont'), icon: 'sheets-font-shrink', run: () => cmd.growFont(-1) },
        ],
        [
          { type: 'toggle', label: t('Sheets.Bold'), icon: 'sheets-bold', key: 'Mod+B', active: () => !!style().b, run: () => cmd.toggle('b') },
          { type: 'toggle', label: t('Sheets.Italic'), icon: 'sheets-italic', key: 'Mod+I', active: () => !!style().i, run: () => cmd.toggle('i') },
          { type: 'toggle', label: t('Sheets.Underline'), icon: 'sheets-underline', key: 'Mod+U', active: () => !!style().u, run: () => cmd.toggle('u') },
          { type: 'toggle', label: t('Sheets.Strike'), icon: 'sheets-strike', key: 'Mod+5', active: () => !!style().strike, run: () => cmd.toggle('strike') },
          { type: 'menu', label: t('Sheets.Borders'), icon: 'sheets-border-bottom', showLabel: false, items: borderItems },
          { type: 'color', label: t('Sheets.FillColor'), icon: 'sheets-fill', auto: t('Sheets.NoFill'), initial: '#ffff00', change: c => cmd.setFill(c) },
          { type: 'color', label: t('Sheets.FontColor'), icon: 'sheets-font-color', auto: t('Sheets.Automatic'), initial: '#c00000', change: c => cmd.setColor(c) },
        ],
      ] }] },
      { label: t('Sheets.Group.Alignment'), items: [{ type: 'stack', rows: [
        [
          { type: 'toggle', label: t('Sheets.Align.top'), icon: 'sheets-align-top', active: () => style().va === 'top', run: () => cmd.valign('top') },
          { type: 'toggle', label: t('Sheets.Align.middle'), icon: 'sheets-align-middle', active: () => style().va === 'middle', run: () => cmd.valign('middle') },
          { type: 'toggle', label: t('Sheets.Align.bottom'), icon: 'sheets-align-bottom', active: () => !style().va || style().va === 'bottom', run: () => cmd.valign('bottom') },
          { type: 'toggle', label: t('Sheets.WrapText'), icon: 'sheets-wrap', showLabel: true, active: () => !!style().wrap, run: () => cmd.wrapText() },
        ],
        [
          { type: 'toggle', label: t('Sheets.Align.left'), icon: 'sheets-align-left', active: () => style().ha === 'left', run: () => cmd.align('left') },
          { type: 'toggle', label: t('Sheets.Align.center'), icon: 'sheets-align-center', active: () => style().ha === 'center', run: () => cmd.align('center') },
          { type: 'toggle', label: t('Sheets.Align.right'), icon: 'sheets-align-right', active: () => style().ha === 'right', run: () => cmd.align('right') },
          { type: 'button', label: t('Sheets.IndentLess'), icon: 'sheets-indent-less', run: () => cmd.indent(-1) },
          { type: 'button', label: t('Sheets.IndentMore'), icon: 'sheets-indent-more', run: () => cmd.indent(1) },
          { type: 'split', label: t('Sheets.MergeCenter'), icon: 'sheets-merge', showLabel: true, active: () => app.isMerged(), run: () => app.merge('center'), items: [
            { label: t('Sheets.MergeCenter'), icon: 'sheets-merge', run: () => app.merge('center') },
            { label: t('Sheets.MergeAcross'), icon: 'sheets-merge-across', run: () => app.merge('across') },
            { label: t('Sheets.MergeCells'), icon: 'sheets-merge', run: () => app.merge('merge') },
            { label: t('Sheets.Unmerge'), icon: 'sheets-unmerge', run: () => app.merge('unmerge') },
          ] },
        ],
      ] }] },
      { label: t('Sheets.Group.Number'), items: [{ type: 'stack', rows: [
        [{ type: 'select', label: t('Sheets.NumberFormat'), width: 150, options: numberOptions, value: () => cmd.numberKind(), change: v => cmd.applyNumberKind(v) }],
        [
          { type: 'split', label: t('Sheets.Currency'), icon: 'sheets-currency', run: () => cmd.numberFormat(p.currency), items: [
            { label: '€ ' + t('Sheets.Currency.Euro'), run: () => cmd.numberFormat(p.currencyEuro) },
            { label: '$ ' + t('Sheets.Currency.Dollar'), run: () => cmd.numberFormat(p.currencyDollar) },
            { label: '£ ' + t('Sheets.Currency.Pound'), run: () => cmd.numberFormat(p.currencyPound) },
            '-',
            { label: t('Sheets.Nf.accounting'), run: () => cmd.numberFormat(p.accounting) },
            { label: t('Sheets.MoreFormats'), run: () => dlg.formatCells('number') },
          ] },
          { type: 'button', label: t('Sheets.Percent'), icon: 'sheets-percent', key: 'Mod+Shift+%', run: () => cmd.numberFormat(p.percentStyle) },
          { type: 'button', label: t('Sheets.Comma'), icon: 'sheets-comma', run: () => cmd.numberFormat(p.comma) },
          { type: 'button', label: t('Sheets.DecimalsMore'), icon: 'sheets-dec-more', run: () => cmd.decimals(1) },
          { type: 'button', label: t('Sheets.DecimalsLess'), icon: 'sheets-dec-less', run: () => cmd.decimals(-1) },
        ],
      ] }] },
      { label: t('Sheets.Group.Styles'), items: [
        { type: 'menu', size: 'large', label: t('Sheets.CondFormat'), icon: 'sheets-cf', items: cfMenu },
        { type: 'menu', size: 'large', label: t('Sheets.CellStyles'), icon: 'sheets-cell-styles', items: () => [
          { heading: t('Sheets.Style.GoodBad') }, ...CELL_STYLES.slice(0, 4).map(styleItem),
          { heading: t('Sheets.Style.Data') }, ...CELL_STYLES.slice(4, 12).map(styleItem),
          { heading: t('Sheets.Style.Titles') }, ...CELL_STYLES.slice(12, 18).map(styleItem),
          { heading: t('Sheets.Style.Themed') }, ...CELL_STYLES.slice(18).map(styleItem),
        ] },
      ] },
      { label: t('Sheets.Group.Cells'), items: [
        { type: 'menu', size: 'large', label: t('Sheets.Insert'), icon: 'sheets-insert-cells', items: [
          { label: t('Sheets.InsertCells'), icon: 'sheets-insert-cells', run: () => cmd.insertDialog() },
          { label: t('Sheets.InsertRows'), icon: 'sheets-insert-row', run: () => cmd.insertLines('row') },
          { label: t('Sheets.InsertCols'), icon: 'sheets-insert-col', run: () => cmd.insertLines('col') },
          { label: t('Sheets.Sheet.Insert'), icon: 'sheets-sheet-add', key: 'Shift+F11', run: () => cmd.addSheet() },
        ] },
        { type: 'menu', size: 'large', label: t('Sheets.Delete'), icon: 'sheets-delete-cells', items: [
          { label: t('Sheets.DeleteCells'), icon: 'sheets-delete-cells', key: 'Mod+-', run: () => cmd.deleteDialog() },
          { label: t('Sheets.DeleteRows'), icon: 'sheets-delete-row', run: () => cmd.deleteLines('row') },
          { label: t('Sheets.DeleteCols'), icon: 'sheets-delete-col', run: () => cmd.deleteLines('col') },
          { label: t('Sheets.Sheet.Delete'), icon: 'trash', run: () => cmd.deleteSheet() },
        ] },
        { type: 'menu', size: 'large', label: t('Sheets.FormatMenu'), icon: 'sheets-format-cells', items: [
          { heading: t('Sheets.CellSize') },
          { label: t('Sheets.RowHeight'), icon: 'sheets-row-height', run: () => dlg.size('row') },
          { label: t('Sheets.AutofitRow'), icon: 'sheets-row-height', run: () => cmd.autofit('row', app.sel.active.row) },
          { label: t('Sheets.ColWidth'), icon: 'sheets-col-width', run: () => dlg.size('col') },
          { label: t('Sheets.AutofitCol'), icon: 'sheets-col-width', run: () => cmd.fitColumns(app.selectedCols()) },
          { label: t('Sheets.DefaultWidth'), icon: 'sheets-col-width', run: () => dlg.size('col', true) },
          { heading: t('Sheets.Visibility') },
          { label: t('Sheets.HideRows'), icon: 'eye-off', key: 'Mod+9', run: () => cmd.hideLines('row', true) },
          { label: t('Sheets.HideCols'), icon: 'eye-off', key: 'Mod+0', run: () => cmd.hideLines('col', true) },
          { label: t('Sheets.UnhideRows'), icon: 'eye', run: () => cmd.hideLines('row', false) },
          { label: t('Sheets.UnhideCols'), icon: 'eye', run: () => cmd.hideLines('col', false) },
          { label: t('Sheets.Sheet.Hide'), icon: 'eye-off', run: () => cmd.hideSheet() },
          { label: t('Sheets.Sheet.Unhide'), icon: 'eye', disabled: () => !app.wb.sheets.some(s => s.hidden), run: () => dlg.unhideSheet() },
          { heading: t('Sheets.OrganizeSheets') },
          { label: t('Sheets.Sheet.Rename'), icon: 'edit', run: () => app.tabs.rename(app.sheet, app.tabs.list.querySelector('.is-active')) },
          { label: t('Sheets.Sheet.Duplicate'), icon: 'copy', run: () => cmd.duplicateSheet() },
          { label: t('Sheets.Sheet.Color'), icon: 'palette', run: () => app.tabs.colorMenu(app.sheet) },
          { heading: t('Sheets.Protection') },
          { label: t('Sheets.Protect'), icon: 'lock', checked: () => !!app.sheet.protected, run: () => cmd.protect() },
          { label: t('Sheets.LockCell'), icon: 'lock', checked: () => !style().unlocked, run: () => cmd.toggleLocked() },
          { label: t('Sheets.FormatCells') + '…', icon: 'sheets-format-cells', key: 'Mod+1', run: () => dlg.formatCells() },
        ] },
      ] },
      { label: t('Sheets.Group.Editing'), items: [{ type: 'stack', rows: [
        [{ type: 'split', label: t('Sheets.AutoSum'), icon: 'sheets-autosum', showLabel: true, key: 'Alt+=', run: () => cmd.autoSum('SUM'), items: [
          { label: t('Sheets.Sum.Sum'), run: () => cmd.autoSum('SUM') },
          { label: t('Sheets.Sum.Average'), run: () => cmd.autoSum('AVERAGE') },
          { label: t('Sheets.Sum.Count'), run: () => cmd.autoSum('COUNT') },
          { label: t('Sheets.Sum.Max'), run: () => cmd.autoSum('MAX') },
          { label: t('Sheets.Sum.Min'), run: () => cmd.autoSum('MIN') },
          '-',
          { label: t('Sheets.MoreFunctions'), icon: 'sheets-function', run: () => dlg.insertFunction() },
        ] }],
        [{ type: 'menu', label: t('Sheets.Fill'), icon: 'sheets-fill-down', items: [
          { label: t('Sheets.FillDown'), icon: 'arrow-down', key: 'Mod+D', run: () => cmd.fillDirection('down') },
          { label: t('Sheets.FillRight'), icon: 'arrow-right', key: 'Mod+R', run: () => cmd.fillDirection('right') },
        ] }],
        [{ type: 'menu', label: t('Sheets.Clear'), icon: 'sheets-clear', items: [
          { label: t('Sheets.ClearAll'), icon: 'sheets-clear', run: () => cmd.clear('all') },
          { label: t('Sheets.ClearFormats'), icon: 'sheets-clear-format', run: () => cmd.clear('formats') },
          { label: t('Sheets.ClearContents'), icon: 'broom', key: 'Delete', run: () => cmd.clear('contents') },
          { label: t('Sheets.ClearNotes'), icon: 'comment', run: () => cmd.clear('notes') },
          { label: t('Sheets.ClearLinks'), icon: 'link', run: () => cmd.clear('links') },
        ] }],
      ] }, { type: 'stack', rows: [
        [{ type: 'menu', label: t('Sheets.SortFilter'), icon: 'sheets-sort', items: [
          { label: t('Sheets.SortAZ'), icon: 'sheets-sort-az', run: () => cmd.sort(false) },
          { label: t('Sheets.SortZA'), icon: 'sheets-sort-za', run: () => cmd.sort(true) },
          { label: t('Sheets.CustomSort'), icon: 'sheets-sort', run: () => dlg.sort() },
          '-',
          { label: t('Sheets.Filter'), icon: 'filter', key: 'Mod+Shift+L', checked: () => !!app.sheet.filter, run: () => app.filters.toggle() },
          { label: t('Sheets.Filter.Clear'), icon: 'sheets-filter-clear', disabled: () => !app.sheet.filter?.cols || !Object.keys(app.sheet.filter.cols).length, run: () => app.filters.clearAll() },
          { label: t('Sheets.Filter.Reapply'), icon: 'refresh', disabled: () => !app.sheet.filter, run: () => app.filters.reapply() },
        ] }],
        [{ type: 'menu', label: t('Sheets.FindSelect'), icon: 'search', edit: false, items: [
          { label: t('Sheets.Find.Title'), icon: 'search', key: 'Mod+F', edit: false, run: () => dlg.find(false) },
          { label: t('Sheets.Find.ReplaceTitle'), icon: 'sheets-replace', key: 'Mod+H', run: () => dlg.find(true) },
          { label: t('Sheets.GoTo'), icon: 'arrow-right', key: 'Mod+G', edit: false, run: () => dlg.goTo() },
          '-',
          { label: t('Sheets.Select.Formulas'), icon: 'sheets-function', edit: false, run: () => app.selectSpecial('formulas') },
          { label: t('Sheets.Select.Notes'), icon: 'comment', edit: false, run: () => app.selectSpecial('notes') },
          { label: t('Sheets.Select.Constants'), icon: 'type', edit: false, run: () => app.selectSpecial('constants') },
          { label: t('Sheets.Select.Validation'), icon: 'sheets-validation', edit: false, run: () => app.selectSpecial('validation') },
        ] }],
      ] }] },
    ] },
    { id: 'insert', label: t('Sheets.Tab.Insert'), groups: [
      { label: t('Sheets.Group.Illustrations'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Pictures'), icon: 'image', run: () => app.objects.insertPicture() },
      ] },
      { label: t('Sheets.Group.Charts'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Chart.column'), icon: 'sheets-chart-column', run: () => app.objects.insertChart('column') },
        { type: 'button', size: 'large', label: t('Sheets.Chart.line'), icon: 'sheets-chart-line', run: () => app.objects.insertChart('line') },
        { type: 'button', size: 'large', label: t('Sheets.Chart.pie'), icon: 'sheets-chart-pie', run: () => app.objects.insertChart('pie') },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.Chart.bar'), icon: 'sheets-chart-bar', showLabel: true, run: () => app.objects.insertChart('bar') }],
          [{ type: 'button', label: t('Sheets.Chart.area'), icon: 'sheets-chart-area', showLabel: true, run: () => app.objects.insertChart('area') }],
          [{ type: 'button', label: t('Sheets.Chart.doughnut'), icon: 'sheets-chart-doughnut', showLabel: true, run: () => app.objects.insertChart('doughnut') }],
        ] },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.Chart.scatter'), icon: 'sheets-chart-scatter', showLabel: true, run: () => app.objects.insertChart('scatter') }],
        ] },
      ] },
      { label: t('Sheets.Group.Links'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Link'), icon: 'link', key: 'Mod+K', run: () => dlg.link() },
      ] },
      { label: t('Sheets.Group.Notes'), items: [
        { type: 'button', size: 'large', label: t('Sheets.NewNote'), icon: 'comment', key: 'Shift+F2', run: () => app.notes.edit() },
      ] },
      { label: t('Sheets.Group.Functions'), items: [
        { type: 'button', size: 'large', label: t('Sheets.InsertFunction'), icon: 'sheets-function', key: 'Shift+F3', run: () => dlg.insertFunction() },
      ] },
    ] },
    { id: 'layout', label: t('Sheets.Tab.Layout'), groups: [
      { label: t('Sheets.Group.PageSetup'), items: [
        { type: 'menu', size: 'large', label: t('Sheets.Orientation'), icon: 'sheets-orientation', items: [
          { label: t('Sheets.Portrait'), icon: 'file', checked: () => !app.sheet.print?.landscape, run: () => cmd.setPrint({ landscape: undefined }) },
          { label: t('Sheets.Landscape'), icon: 'sheets-landscape', checked: () => !!app.sheet.print?.landscape, run: () => cmd.setPrint({ landscape: true }) },
        ] },
        { type: 'menu', size: 'large', label: t('Sheets.Scale'), icon: 'sheets-fit', items: [
          { label: t('Sheets.Fit.none'), checked: () => !app.sheet.print?.fit, run: () => cmd.setPrint({ fit: undefined }) },
          { label: t('Sheets.Fit.width'), checked: () => app.sheet.print?.fit === 'width', run: () => cmd.setPrint({ fit: 'width' }) },
          { label: t('Sheets.Fit.page'), checked: () => app.sheet.print?.fit === 'page', run: () => cmd.setPrint({ fit: 'page' }) },
        ] },
        { type: 'button', size: 'large', label: t('Sheets.PrintTitles'), icon: 'sheets-print-titles', run: () => dlg.printTitles() },
      ] },
      { label: t('Sheets.Group.SheetOptions'), items: [{ type: 'stack', rows: [
        [{ type: 'toggle', label: t('Sheets.Gridlines.View'), icon: 'sheets-gridlines', showLabel: true, edit: false, active: () => app.sheet.grid, run: () => cmd.toggleView('grid') }],
        [{ type: 'toggle', label: t('Sheets.Gridlines.Print'), icon: 'print', showLabel: true, active: () => !!app.sheet.print?.gridlines, run: () => cmd.setPrint({ gridlines: !app.sheet.print?.gridlines || undefined }) }],
      ] }, { type: 'stack', rows: [
        [{ type: 'toggle', label: t('Sheets.Headings.View'), icon: 'sheets-headings', showLabel: true, edit: false, active: () => app.sheet.headings, run: () => cmd.toggleView('headings') }],
        [{ type: 'toggle', label: t('Sheets.Headings.Print'), icon: 'print', showLabel: true, active: () => !!app.sheet.print?.headings, run: () => cmd.setPrint({ headings: !app.sheet.print?.headings || undefined }) }],
      ] }] },
      { label: t('Sheets.Group.Print'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Print'), icon: 'print', edit: false, run: () => app.print() },
      ] },
    ] },
    { id: 'formulas', label: t('Sheets.Tab.Formulas'), groups: [
      { label: t('Sheets.Group.Library'), items: [
        { type: 'button', size: 'large', label: t('Sheets.InsertFunction'), icon: 'sheets-function', run: () => dlg.insertFunction() },
        { type: 'split', size: 'large', label: t('Sheets.AutoSum'), icon: 'sheets-autosum', run: () => cmd.autoSum('SUM'), items: [
          { label: t('Sheets.Sum.Sum'), run: () => cmd.autoSum('SUM') }, { label: t('Sheets.Sum.Average'), run: () => cmd.autoSum('AVERAGE') },
          { label: t('Sheets.Sum.Count'), run: () => cmd.autoSum('COUNT') }, { label: t('Sheets.Sum.Max'), run: () => cmd.autoSum('MAX') }, { label: t('Sheets.Sum.Min'), run: () => cmd.autoSum('MIN') },
        ] },
        ...['financial', 'logical', 'text', 'date', 'lookup', 'math'].map(cat => ({ type: 'menu', size: 'large', label: t('Sheets.Cat.' + cat), icon: 'sheets-cat-' + cat, items: fnMenu(cat) })),
        { type: 'menu', size: 'large', label: t('Sheets.MoreFunctions'), icon: 'sheets-function', items: () => [
          { heading: t('Sheets.Cat.stat') }, ...fnMenu('stat')(),
          { heading: t('Sheets.Cat.info') }, ...fnMenu('info')(),
        ] },
      ] },
      { label: t('Sheets.Group.Names'), items: [
        { type: 'button', size: 'large', label: t('Sheets.NameManager'), icon: 'sheets-names', run: () => dlg.nameManager() },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.DefineName'), icon: 'tag', showLabel: true, run: () => dlg.defineName() }],
          [{ type: 'menu', label: t('Sheets.UseInFormula'), icon: 'sheets-function', items: () => (app.wb.names.length ? app.wb.names.map(n => ({ label: n.name, run: () => app.insertNameIntoFormula(n.name) })) : [{ label: t('Sheets.Names.None'), disabled: () => true }]) }],
        ] },
      ] },
      { label: t('Sheets.Group.Auditing'), items: [
        { type: 'toggle', size: 'large', label: t('Sheets.ShowFormulas'), icon: 'sheets-show-formulas', key: 'Mod+`', edit: false, active: () => app.showFormulas, run: () => cmd.toggleFormulas() },
      ] },
      { label: t('Sheets.Group.Calculation'), items: [
        { type: 'button', size: 'large', label: t('Sheets.CalcNow'), icon: 'sheets-calc', key: 'F9', edit: false, run: () => cmd.recalc() },
      ] },
    ] },
    { id: 'data', label: t('Sheets.Tab.Data'), groups: [
      { label: t('Sheets.Group.SortFilter'), items: [
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.SortAZ'), icon: 'sheets-sort-az', run: () => cmd.sort(false) }],
          [{ type: 'button', label: t('Sheets.SortZA'), icon: 'sheets-sort-za', run: () => cmd.sort(true) }],
        ] },
        { type: 'button', size: 'large', label: t('Sheets.Sort'), icon: 'sheets-sort', run: () => dlg.sort() },
        { type: 'toggle', size: 'large', label: t('Sheets.Filter'), icon: 'filter', active: () => !!app.sheet.filter, run: () => app.filters.toggle() },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.Filter.Clear'), icon: 'sheets-filter-clear', showLabel: true, disabled: () => !app.sheet.filter?.cols || !Object.keys(app.sheet.filter.cols).length, run: () => app.filters.clearAll() }],
          [{ type: 'button', label: t('Sheets.Filter.Reapply'), icon: 'refresh', showLabel: true, disabled: () => !app.sheet.filter, run: () => app.filters.reapply() }],
        ] },
      ] },
      { label: t('Sheets.Group.DataTools'), items: [
        { type: 'button', size: 'large', label: t('Sheets.TextCols.Title'), icon: 'sheets-text-columns', run: () => dlg.textToColumns() },
        { type: 'button', size: 'large', label: t('Sheets.Dup.Title'), icon: 'sheets-remove-dup', run: () => dlg.removeDuplicates() },
        { type: 'button', size: 'large', label: t('Sheets.Validation'), icon: 'sheets-validation', run: () => dlg.validation() },
      ] },
      { label: t('Sheets.Group.Names'), items: [
        { type: 'button', size: 'large', label: t('Sheets.NameManager'), icon: 'sheets-names', run: () => dlg.nameManager() },
      ] },
    ] },
    { id: 'review', label: t('Sheets.Tab.Review'), groups: [
      { label: t('Sheets.Group.Notes'), items: [
        { type: 'button', size: 'large', label: t('Sheets.NewNote'), icon: 'comment', run: () => app.notes.edit() },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.DeleteNote'), icon: 'trash', showLabel: true, disabled: () => !app.sheet.cell(app.sel.active.row, app.sel.active.col)?.note, run: () => cmd.clear('notes') }],
          [{ type: 'button', label: t('Sheets.PrevNote'), icon: 'arrow-up', showLabel: true, edit: false, run: () => app.jumpNote(-1) }],
          [{ type: 'button', label: t('Sheets.NextNote'), icon: 'arrow-down', showLabel: true, edit: false, run: () => app.jumpNote(1) }],
        ] },
        { type: 'toggle', size: 'large', label: t('Sheets.ShowNotes'), icon: 'sheets-notes', edit: false, active: () => app.notes.showAll, run: () => app.notes.toggleAll() },
      ] },
      { label: t('Sheets.Group.Protect'), items: [
        { type: 'toggle', size: 'large', label: t('Sheets.Protect'), icon: 'lock', active: () => !!app.sheet.protected, run: () => cmd.protect() },
        { type: 'toggle', size: 'large', label: t('Sheets.LockCell'), icon: 'key', active: () => !style().unlocked, run: () => cmd.toggleLocked() },
      ] },
    ] },
    { id: 'view', label: t('Sheets.Tab.View'), groups: [
      { label: t('Sheets.Group.Show'), items: [{ type: 'stack', rows: [
        [{ type: 'toggle', label: t('Sheets.FormulaBar'), icon: 'sheets-formula-bar', showLabel: true, edit: false, active: () => app.showBar, run: () => app.toggleBar() }],
        [{ type: 'toggle', label: t('Sheets.Gridlines'), icon: 'sheets-gridlines', showLabel: true, edit: false, active: () => app.sheet.grid, run: () => cmd.toggleView('grid') }],
        [{ type: 'toggle', label: t('Sheets.Headings'), icon: 'sheets-headings', showLabel: true, edit: false, active: () => app.sheet.headings, run: () => cmd.toggleView('headings') }],
      ] }] },
      { label: t('Sheets.Group.Zoom'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Zoom'), icon: 'zoom-in', edit: false, run: () => dlg.zoom() },
        { type: 'button', size: 'large', label: '100 %', icon: 'sheets-zoom-100', edit: false, run: () => cmd.setZoom(1) },
        { type: 'button', size: 'large', label: t('Sheets.ZoomSelection'), icon: 'zoom-fit', edit: false, run: () => cmd.zoomToSelection() },
      ] },
      { label: t('Sheets.Group.Window'), items: [
        { type: 'menu', size: 'large', label: t('Sheets.Freeze'), icon: 'sheets-freeze', edit: false, items: [
          { label: app.sheet.freeze.rows || app.sheet.freeze.cols ? t('Sheets.Unfreeze') : t('Sheets.FreezePanes'), icon: 'sheets-freeze', edit: false, run: () => cmd.freeze(app.sheet.freeze.rows || app.sheet.freeze.cols ? 'none' : 'panes') },
          { label: t('Sheets.FreezeRow'), icon: 'sheets-freeze-row', edit: false, run: () => cmd.freeze('row') },
          { label: t('Sheets.FreezeCol'), icon: 'sheets-freeze-col', edit: false, run: () => cmd.freeze('col') },
        ] },
      ] },
    ] },
    { id: 'chart', label: t('Sheets.Tab.Chart'), hidden: () => !chart(), groups: [
      { label: t('Sheets.Group.ChartType'), items: CHART_TYPES.map(type => ({ type: 'toggle', size: 'large', label: t('Sheets.Chart.' + type), icon: 'sheets-chart-' + type, active: () => chart()?.type === type, run: () => app.objects.updateChart({ type }) })) },
      { label: t('Sheets.Group.ChartLayout'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Chart.Edit'), icon: 'sheets-chart-edit', run: () => dlg.chart(chart()) },
        { type: 'stack', rows: [
          [{ type: 'menu', label: t('Sheets.Chart.Legend'), icon: 'sheets-legend', items: ['right', 'bottom', 'top', 'left', 'none'].map(v => ({ label: t('Sheets.Legend.' + v), checked: () => (chart()?.legend || '') === v, run: () => app.objects.updateChart({ legend: v }) })) }],
          [{ type: 'toggle', label: t('Sheets.Chart.Labels'), icon: 'sheets-labels', showLabel: true, active: () => !!chart()?.labels, run: () => app.objects.updateChart({ labels: !chart()?.labels }) }],
          [{ type: 'toggle', label: t('Sheets.Chart.Stacked'), icon: 'sheets-stacked', showLabel: true, active: () => !!chart()?.stacked, run: () => app.objects.updateChart({ stacked: !chart()?.stacked }) }],
        ] },
        { type: 'stack', rows: [
          [{ type: 'button', label: t('Sheets.Chart.SwitchRows'), icon: 'sheets-transpose', showLabel: true, run: () => app.objects.updateChart({ byRows: !chart()?.byRows }) }],
          [{ type: 'menu', label: t('Sheets.Chart.Colors'), icon: 'palette', items: [
            ['office', CHART_COLORS], ['green', ['#107c41', '#33a167', '#6cc191', '#a8dcbf', '#0b5a2f', '#d3eadb']], ['blue', ['#1f4e79', '#2e75b6', '#5b9bd5', '#9dc3e6', '#bdd7ee', '#deebf7']],
            ['warm', ['#c00000', '#ed7d31', '#ffc000', '#f4b183', '#ff6f61', '#a5402d']], ['gray', ['#262626', '#595959', '#7f7f7f', '#a5a5a5', '#c9c9c9', '#e7e6e6']],
          ].map(([name, colors]) => ({ label: t('Sheets.Palette.' + name), render: () => { const s = el('span', { class: 'mo-sheets-swatch mo-sheets-swatch--bar' }); s.style.background = `linear-gradient(90deg, ${colors.slice(0, 4).map((c, i) => `${c} ${i * 25}% ${(i + 1) * 25}%`).join(', ')})`; return s; }, run: () => app.objects.updateChart({ colors }) })) }],
          [{ type: 'button', label: t('Sheets.Chart.SelectData'), icon: 'table', showLabel: true, edit: false, run: () => app.objects.showData(chart()) }],
        ] },
      ] },
      { label: t('Sheets.Group.Arrange'), items: [
        { type: 'button', size: 'large', label: t('Sheets.Delete'), icon: 'trash', run: () => app.objects.removeSelected() },
      ] },
    ] },
    { id: 'picture', label: t('Sheets.Tab.Picture'), hidden: () => !image(), groups: [
      { label: t('Sheets.Group.Accessibility'), items: [{ type: 'button', size: 'large', label: t('Sheets.Picture.Alt'), icon: 'type', run: () => dlg.altText(image()) }] },
      { label: t('Sheets.Group.Arrange'), items: [
        { type: 'button', size: 'large', label: t('Sheets.BringFront'), icon: 'sheets-bring-front', run: () => app.objects.order(image(), 1) },
        { type: 'button', size: 'large', label: t('Sheets.SendBack'), icon: 'sheets-send-back', run: () => app.objects.order(image(), -1) },
        { type: 'button', size: 'large', label: t('Sheets.Delete'), icon: 'trash', run: () => app.objects.removeSelected() },
      ] },
    ] },
  ];
  return tabs;
}

function applyCss(node, s) {
  if (s.fill) node.style.background = s.fill;
  if (s.color) node.style.color = s.color;
  if (s.b) node.style.fontWeight = '700';
  if (s.i) node.style.fontStyle = 'italic';
  if (s.font) node.style.fontFamily = fontFamily(s.font);
  if (s.bb) node.style.borderBottom = (s.bb.startsWith('double') ? '3px double ' : s.bb.startsWith('thick') ? '3px solid ' : s.bb.startsWith('medium') ? '2px solid ' : '1px solid ') + s.bb.split(' ')[1];
  if (s.bt) node.style.borderTop = '1px solid ' + s.bt.split(' ')[1];
}

export { CATEGORIES };
