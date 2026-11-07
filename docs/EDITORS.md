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

## Sheets content

The content of a spreadsheet (`Services/Formats/Sheets/SheetsHandler.cs` checks it, `wwwroot/js/sheets/model.js` reads and writes it). Everything that is
not needed is left out (a sparse format): an empty cell is not there, a default is not written. Sizes are pixels at 100 % zoom, font sizes points.

```jsonc
{
  "version": 1,
  "active": 0,                                  // index of the sheet that is shown
  "styles": [ {}, { "b": true, "fill": "#fff2cc", "nf": "#,##0.00" } ],   // the style table; cells, rows and columns refer to it by index ("s"); 0 is the default ({})
  "names": [ { "name": "Rate", "ref": "Sheet1!$B$2" } ],                   // defined names: a reference or formula without "="
  "sheets": [ {
    "id": "s1", "name": "Sheet1",               // stable id [A-Za-z0-9_-]{1,32}; name 1-31 characters, unique, none of : \ / ? * [ ]
    "color": "#c00000", "hidden": true,         // tab colour, hidden sheet (optional)
    "cells": {
      "A1": { "v": "Text" },                    // v: number | string | boolean
      "B2": { "v": 46000, "s": 3 },             // a date is a number (Excel serial: 1 = 1900-01-01) with a date format in its style
      "C2": { "f": "SUM(B2:B9)", "v": 42 },     // f: the formula, canonical (English names, "," between arguments, "." decimals), without "="; v: its last value
      "D2": { "f": "1/0", "v": "#DIV/0!", "t": "e" },   // t "e": v is an error (#DIV/0! #N/A #NAME? #NULL! #NUM! #REF! #VALUE! #SPILL! #CALC! #CIRC!)
      "E2": { "s": 2 },                         // an empty cell with a format
      "F2": { "v": "Site", "link": "https://example.com", "note": "A note" }  // link: http(s), mailto or "#Sheet1!A1"; note: plain text
    },
    "cols": { "A": { "w": 120 }, "C": { "hidden": true }, "D": { "s": 2 } },  // width, hidden, style of the whole column
    "rows": { "1": { "h": 30 }, "5": { "hidden": true }, "7": { "s": 4 } },   // height, hidden, style of the whole row (rows count from 1)
    "colW": 64, "rowH": 20,                     // default width / height (optional, 64 / 20)
    "merges": [ "A1:C1" ],
    "freeze": { "rows": 1, "cols": 0 },          // frozen rows at the top / columns at the left
    "filter": { "range": "A1:D20", "cols": {    // AutoFilter (the header row is the first row of the range); key = column letter
      "B": { "values": ["North", "South", ""] }, // only these values are shown ("" = blanks)
      "C": { "cond": { "op": "gt", "v": "100", "op2": "lt", "v2": "500", "or": false } } } },  // ops: eq ne gt ge lt le begins ends contains notContains
    "cf": [                                     // conditional formats, in order of priority
      { "range": "B2:B20", "type": "cell", "op": "gt", "v1": "100", "style": { "color": "#9c0006", "fill": "#ffc7ce" } },  // ops: gt ge lt le eq ne between notBetween (v2)
      { "range": "A2:A20", "type": "text", "op": "contains", "v1": "late", "style": { "b": true } },   // contains notContains begins ends
      { "range": "A2:A20", "type": "dup", "style": { … } },        // also "unique"
      { "range": "C2:C20", "type": "top", "n": 3, "percent": false, "style": { … } },   // also "bottom"
      { "range": "C2:C20", "type": "above", "style": { … } },      // above / below the average
      { "range": "A2:D20", "type": "formula", "v1": "$D2>100", "style": { … } },
      { "range": "C2:C20", "type": "bar", "color": "#638ec6" },    // data bars
      { "range": "C2:C20", "type": "scale", "colors": ["#f8696b", "#ffeb84", "#63be7b"] }   // 2 or 3 colours, minimum → (midpoint) → maximum
    ],                                          // style of a rule: color, fill, b, i, u, strike
    "validations": [ { "range": "C2:C50", "type": "list", "v1": "Open,Done", "blank": true, "dropdown": true,
                       "error": "…", "errorTitle": "…", "errorStyle": "stop", "prompt": "…" } ],
                                                // type: list (v1 = items separated by "," or "=$E$1:$E$5"), whole, decimal, date, time, length, custom (v1 = formula);
                                                // op (not for list/custom): between notBetween eq ne gt ge lt le with v1 / v2; errorStyle: stop warning info
    "charts": [ { "id": "c1", "type": "column", "range": "A1:C6", "x": 400, "y": 40, "w": 480, "h": 300,   // type: column bar line area pie doughnut scatter
                  "title": "Sales", "legend": "right", "xTitle": "", "yTitle": "", "colors": ["#4472c4"], "byRows": false, "stacked": false, "labels": false } ],
    "images": [ { "id": "i1", "src": "data:image/png;base64,…", "x": 10, "y": 10, "w": 200, "h": 150, "alt": "" } ],   // png, jpeg, gif, webp only
    "grid": false, "headings": false,           // gridlines / row and column headings hidden (shown when left out)
    "zoom": 1.25,
    "print": { "landscape": true, "fit": "width", "gridlines": true, "headings": false, "titles": "1:1" },  // fit: width page none; titles: rows repeated on every page
    "protected": true,                          // locked cells cannot be changed (style "unlocked": true opens a cell)
    "sel": "B3"                                 // the active cell when it was saved
  } ]
}
```

A style (in `styles` and, as a subset, in conditional formats): `font` (family), `size` (pt), `b`, `i`, `u`, `strike`, `color`, `fill` (`#rrggbb`), `ha`
(`left center right justify`), `va` (`top middle bottom`), `wrap`, `indent` (0-15), `nf` (an Excel number format code: `0.00`, `#,##0 "€"`, `0%`, `dd.mm.yyyy`,
`[Red]-#,##0`, `@` …; General when left out), `bt` `bb` `bl` `br` (borders top, bottom, left, right: `"thin|medium|thick|dashed|dotted|double|hair #rrggbb"`),
`unlocked`. The style of a cell is its own `s`, else the one of its row, else the one of its column.

The server (`Normalize`) keeps only these fields, drops what does not fit (colours, fonts, links other than http/https/mailto/#, pictures other than
`data:image/png|jpeg|gif|webp`), and limits sizes: 200 sheets, 2,000,000 cells, 1,048,576 rows × 16,384 columns (A1 … XFD1048576), 32,767 characters
of text, 8,192 of a formula. Formulas are text for the server: it never runs them, the editor calculates (`wwwroot/js/sheets/engine/`).
Files: `.xlsx` (ClosedXML: values, formulas with their cached values, formats, sizes, merges, frozen panes, hidden rows/columns, tab colours, notes, links;
charts are not written), `.csv` and `.tsv` (UTF-8 with BOM; the import detects `,` `;` or tab and a decimal comma).
