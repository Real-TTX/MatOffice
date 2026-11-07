# The four editors – the contract

MatOffice is one app with four editors: **Docs** (like Word), **Sheets** (like Excel), **Slides** (like PowerPoint) and **Base** (like Access).
They share one frame (title bar, ribbon, backstage "File", status bar, saving) so the suite feels like one product, and each one owns its work area,
its content format and its file conversions. This file is the contract between the frame and an editor. Keep it in sync with the code.

## Who owns what

| Part | Files | Owner |
|---|---|---|
| Frame (title bar, ribbon builder, backstage, status bar, popovers, palette) | `wwwroot/js/editor/frame.js`, `wwwroot/css/editor.css`, `Pages/Shared/_LayoutEditor.cshtml` | shared – change only with care, all editors use it |
| Session (load, autosave, conflicts, offline queue, polling) | `wwwroot/js/editor/session.js` | shared |
| Helpers (`t`, `el`, `icon`, `request`, `dialog`, `History`, `shrinkImage` …) | `wwwroot/js/editor/util.js` | shared |
| The editor of an app | `wwwroot/js/<slug>/*.js` (entry `app.js`), `wwwroot/css/<slug>.css`, `wwwroot/img/icons-<slug>.svg` | the app |
| Content check, templates, file formats | `Services/Formats/<Kind>/*.cs` (`<Kind>Handler : IKindHandler`), `Resources/Templates/<slug>/*` | the app |
| Texts | `Resources/Texts/<Kind>.json` (server keys `<Kind>.*`), `Resources/Texts/Js<Kind>.json` (script keys `Js.<Kind>.*`) | the app |
| Tests | `tests/MatOffice.Tests/<Kind>*Tests.cs` | the app |

`<slug>` is `docs`, `sheets`, `slides`, `base`; `<Kind>` is `Docs`, `Sheets`, `Slides`, `Base`. CSS classes of an editor use the prefix `mo-<slug>-`
(`mo-sheets-grid`), the icons of its ribbon the ids `mat-i-<slug>-<name>` in `img/icons-<slug>.svg` (24 × 24, stroke 2, round ends, colour =
`currentColor`, drawn like the toolkit's; the layout puts the sprite into the page of that editor only). The toolkit's icons (`matui/matui-icons.svg`) and the
app's (`img/app-icons.svg`: undo, redo, zoom-in/out/fit, print, comment, text, share, history …) can be used as they are.

## The page

`/docs/12`, `/sheets/12`, `/slides/12`, `/base/12` (a document; `?v=<id>` a version, read-only), `/s/<secret>` (a link), `/h/<secret>` (a file of a
connected app such as MatDrive). All of them render `_LayoutEditor`: the title bar, `#mo-ribbon`, `#mo-banner`, `#mo-work` (the work area, `overflow: hidden`:
the editor scrolls inside it), `#mo-statusbar`, `#mo-backstage`, and load `/js/<slug>/app.js` as an ES module. The CSP of these pages allows inline
**styles** (`style-src 'unsafe-inline'`, the content of a document carries styles) but never inline script, `eval` or anything from another site.

## The entry module

```js
import { startEditor } from '../editor/frame.js';

startEditor({
  async mount(ctx) { /* build the UI into ctx.work, the ribbon with ctx.ribbon.set([...]), the status bar with ctx.status */ },
  getContent() { return content; },            // the whole content as a JSON-able object (saved as it is)
  setContent(content, info) { /* show another content: a conflict resolved with "theirs", a change of somebody else (info.remote) */ },
  async thumbnail() { return pngBlob; },       // optional: ~ 480 × 360 PNG of the first page/sheet/slide for the gallery (≤ 512 KB)
  print() { /* optional: prepare a print view, then window.print() */ },
  info() { return node; },                     // optional: more lines for File → Info (words, cells, slides, records …)
  downloads() { return [{ ext: 'PDF', label: '…', run() {} }]; }, // optional: downloads made in the browser (File → Download as)
});
```

`ctx`:

| Field | What it is |
|---|---|
| `ctx.doc` | the JSON of the server: `{ mode: document|version|host, id, kind, name, fileName?, revision, role, canEdit, signedIn, owner, api, formats: [{ ext, label, mime }], user: { id, name }, content }` |
| `ctx.readOnly` | `!doc.canEdit` (a viewer, a version, a link without edit right, a read-only file of a host): the editor must not change anything then |
| `ctx.changed()` | call after EVERY change of the content: the session saves the whole content ~1.2 s later (PUT `{ revision, content }`), Ctrl+S saves at once |
| `ctx.work` | the element of the work area |
| `ctx.ribbon` | `set(tabs)`, `select(id)`, `refresh()` (call after the selection changed: buttons show their state) – see below |
| `ctx.status` | `left(...texts or nodes)`, `right(...nodes)`, `zoom({ value, min, max, change(f), fit() })`, `setZoom(f)` |
| `ctx.backstage(section)` | opens File at `info`, `download`, `print`, `share` … |
| `ctx.session` | events `saved`, `remote`, `state`; `flush()` |
| `ctx.t, ctx.el, ctx.icon, ctx.request, ctx.toast, ctx.dialog, ctx.download` | the helpers of `util.js` |

Undo/redo, the clipboard, keyboard handling inside the work area and the selection are the editor's own. `util.js` has a snapshot `History` that is enough
for most cases. Never use `window.alert/confirm/prompt` (DIALOG-01): `ctx.dialog({ title, icon, body: Node, buttons })`, `MatUi.confirm`, `MatUi.prompt`.
Put text into the page with `textContent` or through the sanitizer of the editor; never put content of a document into `innerHTML` unsanitized.

### The ribbon

```js
ctx.ribbon.set([
  { id: 'home', label: t('Docs.Tab.Home'), groups: [
    { label: t('Docs.Group.Clipboard'), items: [
      { type: 'button', id: 'paste', label: t('Docs.Paste'), icon: 'clipboard', size: 'large', run: paste },
      { type: 'stack', rows: [
        [{ type: 'button', label: t('Docs.Cut'), icon: 'docs-cut', run: cut, key: 'Mod+X', global: false }],
        [{ type: 'button', label: t('Docs.Copy'), icon: 'copy', run: copy, edit: false }],
      ] },
    ] },
    { label: t('Docs.Group.Font'), items: [
      { type: 'select', label: t('Docs.Font'), options: FONTS, value: () => sel.font, change: setFont, width: 150 },
      { type: 'combo', label: t('Docs.Size'), options: ['8','9','10','11','12','14','18','24','36'], value: () => sel.size, change: setSize, width: 56 },
      { type: 'toggle', label: t('Docs.Bold'), icon: 'docs-bold', key: 'Mod+B', active: () => sel.bold, run: toggleBold },
      { type: 'color', label: t('Docs.FontColour'), icon: 'docs-font-colour', auto: t('Docs.Automatic'), change: setColour },
      { type: 'menu', label: t('Docs.Lists'), icon: 'list', items: [{ label: '…', icon: '…', run, checked: () => … }, '-', { heading: '…' }] },
      { type: 'split', label: …, icon: …, run, items: [...] },
      { type: 'gallery', label: t('Docs.Styles'), items: [{ label, render: () => node, run, active: () => bool }] },
      { type: 'number', label, value, change, min, max, step, width },
      { type: 'custom', render: () => node, refresh: (node, readOnly) => {} },
    ] },
  ] },
]);
```

Every item takes `label` (the tooltip and the accessible name), `icon`, `tip`, `key` (`'Mod+B'`, `'Mod+Shift+Z'`, `'F5'`: the frame handles the key
everywhere on the page; `global: false` leaves it to inputs while one has the focus), `edit` (default `true`: disabled when read-only; `false` for
actions that change nothing – copy, zoom, print, find), `disabled()`, `hidden()`. `size: 'large'` is the big button with its label below the icon.
The tab **File** comes first by itself (the backstage). The tab labels and group labels are words of the app (`Home`, `Insert`, `Layout` …).

## The content and the server

The content is one JSON object; its shape is the app's (documented in the handler and here below when it is built). The server never trusts it:
`IKindHandler.Normalize(json)` checks the shape and the limits (`Kinds.MaxContentBytes` = 40 MB, pictures are inside as `data:` addresses) and **removes
everything that could run script** (Docs: an HTML allowlist through `HtmlSanitizer`; no `javascript:` links, no event attributes, no `<script>`, `<iframe>`, `<object>`, no external images);
it returns the canonical JSON that is stored. The same check runs for imports and for the files of hosts.

| Method | Purpose |
|---|---|
| `Create(templateId, language)` | the content of a new document; `Templates` lists the templates of the start page (`blank` first; `Name`/`Description` are text keys) |
| `Normalize(json)` | the check above |
| `Formats` | the foreign formats, `FileFormat(".xlsx", mime, label, CanImport, CanExport)`; the own format (`.matdoc`, `.matsheet`, `.matslides`, `.matbase`) is handled by `NativeFile` |
| `Import(stream, ext)` / `Export(content, ext, name)` | conversions (throw `Formats.FormatException` with a worded message); used by File → Download as, by the import on the start page and for the files of hosts (MatDrive opens `.docx`, saves `.docx`) |
| `PlainText(content, max)` | the text (search, description of a file) |

Libraries in the project for the conversions: **DocumentFormat.OpenXml** (docx, pptx, also xlsx), **ClosedXML** (xlsx), **HtmlSanitizer** (Docs).
Formats that cannot be written well are import-only (`CanExport: false`); a host then opens such a file read-only.

## Saving, versions, conflicts (the frame does it)

- Autosave: `ctx.changed()` → PUT `{api}` `{ revision, content }` → `{ revision }`. 409 → the person chooses theirs/mine. 412 (host file changed at the host) → reload or download. 502 (host down) → kept, retried.
- Every 20 s the session asks `{api}/revision`; a newer revision of somebody else is loaded with `setContent(content, { remote: true })` while nothing waits to be saved.
- Offline: the content waits in `localStorage` (`mo.pending.<api>`) and is sent when the network is back.
- Snapshots for the version history are made by the server (at most one per 10 minutes).

## Checks

- `dotnet build MatOffice.slnx` (0 warnings), `dotnet test` (TextCatalogTests checks every `t('Docs.X')` against `Js.Docs.X` in English and German, and every server key).
- Pure logic (formula engine, layout, converters) gets unit tests: C# in `tests/MatOffice.Tests`, JavaScript modules that do not touch the DOM can be tested in Node (`node --test`).
- In the browser: `./scripts/redeploy.ps1`, sign in with `admin@example.com` / `Start12345`, create a document of the app on `/home`, use it with mouse and keyboard, reload (saved?), a phone width (390 px), dark mode, read-only (a version), print preview. Look at the screenshots.

## Conversions while an editor is open

Two endpoints convert without reading or writing a document (a signed-in person, or `?host=<secret>` of an open host session; anti-forgery header):

- `POST /live/convert/<slug>/read?ext=.xlsx` – body: the bytes of a file → the content of that app as JSON (`IKindHandler.Import` + `Normalize`). Base uses it to take
  the sheets of an Excel file as tables, Slides to take the slides of a PowerPoint file, Docs to insert a Word file.
- `POST /live/convert/<slug>/write?ext=.xlsx&name=…` – body: content JSON → the file (`Export`), e.g. the state on the screen, or one table of a database.

The secret of a host session is the last part of `ctx.doc.api` (`/live/host/<secret>`).

## Slides content

The content of a presentation (version 1; the full comment is in `Services/Formats/Slides/SlidesHandler.cs`). Coordinates are px of a slide at 96 dpi
(16:9 = 1280 × 720, 4:3 = 960 × 720; 1 px = 9525 EMU), font sizes and line widths pt. Text is structured – paragraphs of runs, never HTML – so the server
checks it field by field and writes it as DrawingML.

```js
{
  version: 1,
  size: { w: 1280, h: 720 },
  theme: {                                  // the design, kept whole in the file (an imported pptx brings its own)
    id, name, variant?, fonts: { major, minor }, colors: { dk1, lt1, dk2, lt2, accent1 … accent6, hlink },   // #rrggbb
    background: Fill, heading: Color, text: Color,
    groups?: { cover|content|other: { background?, heading?, text? } },               // per layout group
    decor: [{ on: ['all'|group|layout], shape, x, y, w, h /* fractions */, fill: Color | src: picture, opacity?, flipH?, flipV?, rot? }],
  },
  footer: { number, date, dateText /* '' = today */, text, skipTitle },
  slides: [{
    id, layout: 'title|titleContent|section|twoContent|comparison|titleOnly|blank|pictureCaption',
    hidden?, background?: Fill, notes?: 'plain text',
    transition?: { type: 'none|fade|push|wipe|split|zoom|cover', dir?: 'left|right|up|down', duration /* s */, auto? /* s */ },
    elements: [Element],                    // back to front
    animations?: [{ el, effect: 'appear|fade|fly|zoom|wipe', trigger: 'click|with|after', dir?: 'bottom|top|left|right', duration, delay? }],
  }],
}
// Element: { id, type, x, y, w, h, rot?, flipH?, flipV?, ph? (placeholder slot: title, subtitle, body, left, right, leftHead, rightHead, picture, caption),
//            group? (id shared by the members), name? } and by type:
//   text   { text: TextBody, fill?, line?, shadow?, opacity? }
//   shape  { shape: 'rect|roundRect|ellipse|triangle|…' (PowerPoint preset names), adj?, fill?, line?, shadow?, opacity?, text? }
//   line   { line: Line, arrowStart, arrowEnd: 'none|triangle|arrow|stealth|oval|diamond' }   from (x, y) to (x + w, y + h), mirrored by flipH/flipV
//   image  { src: 'data:image/(png|jpeg|gif|webp);base64,…' (never SVG), crop?: { l, t, r, b }, line?, shadow?, opacity?, alt? }
//   table  { cols: [px], rows: [{ h, cells: [{ text: TextBody, fill? }] }], tableStyle: 'none|light|medium|dark', accent, header, banded, firstCol, lastRow }
// TextBody: { paras: [{ runs: [{ t, b?, i?, u?, s?, font?, size?, color?, hl? }], align?, level? (0–8), bullet? ('none', '•', '1.', 'a)' …),
//             lineSpacing?, spaceBefore?, spaceAfter? }], anchor?: 'top|middle|bottom', autofit?: 'none|shrink|grow', inset?: [l, t, r, b], wrap?: false }
// Fill: { type: 'none' } | { type: 'solid', color } | { type: 'gradient', from, to, angle } | { type: 'image', src }
// Line: { color, width /* pt */, dash: 'solid|dash|dot|dashDot|longDash' }
// Color: '#rrggbb' or a colour of the design 'dk1|lt1|dk2|lt2|accent1..6|hlink' with an optional ':+40' (lighter) / ':-25' (darker)
```

What a run or paragraph does not say comes from the layout slot and the design (title 44 pt in the major font, body 28/24/20/18 pt with bullets …).
The layouts and the designs are data shared by both sides: `Resources/Templates/slides/layouts.json` + `themes.json` on the server,
`wwwroot/js/slides/layouts-data.js` + `themes-data.js` in the editor (`SlidesFormatTests` checks that they agree). `SlidesNormalizer` copies only what it
knows (unknown fields are dropped, numbers clamped, colours and fonts checked, control characters removed, pictures only as PNG/JPEG/GIF/WebP data
addresses); `PptxWriter`/`PptxReader` convert to and from PowerPoint, `.pdf` is printing in the browser.

## Base content

The content of a database (`.matbase`), checked by `BaseNormalizer` (`Services/Formats/Base`, documented in `BaseHandler`), edited by `wwwroot/js/base` (entry
`app.js`). The engine – `values.js`, `expr.js`, `criteria.js`, `sql.js`, `engine.js`, `qbe.js`, `model.js`, `layout.js` – has no DOM and is tested with
`node --test "tests/js/base/*.test.mjs"`.

```
{ version: 1, settings: { currency: "EUR" },
  tables: [{ id: "t1", name, description,
             fields: [{ id: "f1", name, type: text|memo|number|currency|datetime|yesno|autonumber|hyperlink, size (text), numberSize (number: byte|integer|long|
                        single|double|decimal), next (autonumber), description, format (Access names: Currency, Short Date, Yes/No …), decimals, default (an
                        expression), required, indexed: no|dup|unique, caption, validationRule, validationText, inputMask,
                        lookup: null | { kind: "table", table, bound, display } | { kind: "values", values: [], limit } }],
             primaryKey: [field ids], rows: [{ id, v: { fieldId: value } }], nextId,
             datasheet: { sort, filter, filterOn, frozen, totals, totalsOn, hidden, widths, order } }],
  relationships: [{ id, parent: { table, field }, child: { table, field }, integrity, cascadeUpdate, cascadeDelete, join: 1|2|3 }],
  queries: [{ id, name, description, sql (Access SQL; the design grid is read from it), layout: { tables: [{ name, x, y, w, h }], grid }, datasheet }],
  forms: [{ id, name, source (a table or query id, "" = unbound), caption, kind: single|continuous|datasheet|split, width, allowEdits, allowAdditions,
            allowDeletions, sections: { header, detail, footer } }],
  reports: [{ id, name, source, caption, orientation: portrait|landscape, paper: a4|letter, width, grouping: [{ field, header, footer, desc, keepTogether }],
              sorting: [{ field, desc }], sections: { reportHeader, pageHeader, groupHeaders: [], detail, groupFooters: [], pageFooter, reportFooter } }],
  ui: { nav: { collapsed: [], width, hidden }, relationships: { tableId: { x, y } } } }
Section = { height, visible, back, controls: [Control] }
Control = { id, type: label|textbox|combobox|checkbox|button|subform|line|rect, x, y, w, h, name, caption, source (a field, or "=expression"), format,
            decimals, fontSize (pt), bold, italic, align, color, back, border, locked, visible, hideDuplicates, runningSum: no|group|all,
            action: first|previous|next|last|new|save|delete|undo|close|print, rowSource (combo box), subform: { source, linkMaster, linkChild } }
```

Dates are ISO text, Yes/No is true/false, numbers and currency are JSON numbers; empty values are left out of `v`. The server checks the shape, the values of the
records against their types and sizes, unique primary keys and indexes, required fields, enforced referential integrity and at most 200,000 records, and it
removes markup. Records are saved one by one (like Access); design changes can be undone (snapshots without the records). Excel: `read` makes every sheet a
table (the first row gives the field names, the types are guessed), `write` every table a sheet; to export one table or one query result the editor sends a
content with only that table. The tabs that are open are kept per device (`localStorage`, `mo.base.tabs.<api>`), not in the content.
