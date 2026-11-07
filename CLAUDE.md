# CLAUDE.md

Guidance for working in this repository. Keep it in sync with the code: update it whenever architecture, interfaces or conventions change.

## What MatOffice is

An office suite in the browser, four apps in one: **Docs** (text documents, like Word), **Sheets** (spreadsheets, like Excel), **Slides** (presentations,
like PowerPoint) and **Base** (desktop databases, like Access). Documents live in MatOffice (start page `/home` like the office suites, per app `/docs`,
`/sheets`, `/slides`, `/base`), are shared with people (viewer/editor) or through a link, keep a version history (snapshots + named versions), can be starred,
copied and lie 30 days in the trash. Word/Excel/PowerPoint/CSV files are imported and exported. **Connected apps** (hosts, MatDrive first) open their own
files in the editors with an API key (WOPI-like, `Services/Hosts`). Part of the Mat family: sign-in through **MatIdentity**, mail through **MatSMTP**, the
platform contract for **Mattiverse** (all optional, switched on by environment variables).

Stack: ASP.NET Core on **.NET 10**, C#, **Razor Pages**, EF Core + **PostgreSQL**, the **MatUi** toolkit of the sister project MatDesign (`wwwroot/matui`,
release 0.9.3, copied: never edit by hand), plain JavaScript ES modules without a build step. File formats: **DocumentFormat.OpenXml**, **ClosedXML**,
**HtmlSanitizer**. Docker-first. Web port **4820** (pgAdmin 4821, local Postgres 4822). UI languages: English and German (catalog in `Resources/Texts/*.json`).
Accent: Office orange `#d83b01` (white text 4.6:1; accent text `#b23000`, dark mode `#ff8a5c`). The apps: Docs `#185abd`, Sheets `#107c41`, Slides `#c43e1c`, Base `#a4373a`.

## Build / run / test

- **Build:** `dotnet build MatOffice.slnx` (SDK 10, `global.json`). Expected: 0 warnings, 0 errors.
- **Tests:** `dotnet test`. Database tests run when `MATOFFICE_TEST_DB` is set (an administrator connection string, e.g.
  `Host=localhost;Port=4822;Username=matoffice;Password=matoffice;Database=postgres` while the local stack runs); each test creates, migrates and drops a database of its own.
  `TextCatalogTests` checks every text key of the server code and of the editors (`t('Editor.X')` → `Js.Editor.X`) in English and German.
  Pure JavaScript modules of the editors (no DOM) have Node tests in `tests/js/` (`node --test tests/js/`).
- **Run / test the app: always rebuild the container** (`./scripts/redeploy.ps1`, `-Clean` for empty volumes). CSS, JS and cshtml are baked into the image.
- **Test account (always):** `admin@example.com` / `Start12345` – the first account of a fresh stack (it becomes the administrator).
- **Real tests are end to end in a browser** (Playwright with the installed Edge, `channel: 'msedge'`): create a document of each app, edit it with mouse and keyboard, reload,
  offline, a phone (390 px), dark mode, a version (read-only), a host session. Look at the screenshots.
- **Migrations:** `dotnet ef migrations add <Name> -p src/MatOffice -o Data/Migrations` (local tool in `dotnet-tools.json`; a design-time factory exists). Applied at start-up.
- **Updating MatUi:** take the release of MatDesign (`origin/main`): the CSS is the parts of `src/MatUi/wwwroot/css` joined in the order of the first line of `matui.css`;
  `matui.js`, `matui.boot.js`, `matui-icons.svg` as they are.
- Write files with the editor tools, not with shell heredocs that contain apostrophes or backslashes.

## Map

```
src/MatOffice
  Program.cs                 composition root: services, cookie auth (+ MatIdentity), antiforgery header, razor conventions (/docs … = Documents/Index), rate limit, static files, endpoints
  Infrastructure.cs          HealthCheck (--healthcheck), Database.MigrateWithRetry, SecurityHeaders (CSP; editor pages /docs/ /sheets/ /slides/ /base/ /h/ /s/ also allow inline STYLES),
                             MaintenanceService (InstanceService.CleanUpAsync every 6 h), EmptyStringBinderProvider
  Data/                      Entities.cs (User, UserSession, Document, DocumentMember, DocumentVisit, DocumentVersion, HostClient, HostSession, SystemSetting), AppDbContext.cs, Migrations/
  Services/Documents/        THE business layer, static classes (permission check in every function):
                             DocumentAccess (Actor = person and/or link token -> DocumentRole; null = "not found"; administrators have NO access to the documents of others),
                             DocumentService (lists with scope/kind/search/sort, create from template, import, open, save with revision -> Saved/Conflict/Invalid/Denied,
                             rename, copy, star, visits, trash/restore/delete forever, export, thumbnail), VersionService (snapshots ≤ 1 per 10 min, named versions, restore keeps
                             the state before, thinning after 30 days), ShareService (members, link Off/View/Edit), DocumentEndpoints (/live/documents: JSON for the editors)
  Services/Formats/          IKindHandler + Kinds (registry by kind/slug/extension, NativeFile = .matdoc/.matsheet/.matslides/.matbase JSON) and one folder per app
                             (Docs/, Sheets/, Slides/, Base/): Normalize (the server never trusts the browser), templates, Import/Export (docx, xlsx, csv, pptx …)
  Services/Hosts/            HostAdmin (connected apps + API keys, hash only, admins), HostSessionService (open a file of a host: read via its contentUrl with its bearer token,
                             convert, write back on every save with If-Match; 412 = changed at the host), HostEndpoints (/api/v1/host/formats|new|sessions with
                             Authorization: Bearer mo_…; /live/host/{token} for the editor)
  Services/Accounts/         static too: Profiles (own account, paper size), UserAdmin (administrators only; the last administrator stays)
  Services/Mail/MailService  MatSMTP through MailKit, queued in the background; off unless MATSMTP_HOST is set
  Services/Auth/             CurrentUser, PasswordService, SessionService (sessions in the DB), AccountService, ExternalSignInService (MatIdentity), AccountDataService (GDPR export/erase)
  Services/                  Ui (icons, logo, bytes), MenuService, InstanceService (numbers, settings, housekeeping), ListModels (toolbar, pager), Dates, Assets (sprites, per editor icons-<slug>.svg)
  MatPlatform/               copy of the shared package (MatDesign): platform contract, MatIdentity sign-in, back-channel logout, MatSmtpSettings. Never edit here.
  Platform/                  MatOfficePlatform (tiles, settings, actions for Mattiverse), MatOfficeIdentityUsers
  Controls/                  ONE code base per control: <mo-field>, <mo-buttons>, <mo-table>, <mo-tabs>/<mo-tab>
  Pages/                     Documents/ (Index = /home and /<app>: start row + gallery/table, New = template or upload, Edit/Members/Share/Versions = the tabs of a document,
                             Version), Editor (/<app>/{id}), Host (/h/{token}), Public/Link (/s/{token}), Admin/ (Users, Hosts, System), Account/, About,
                             Shared/ (_Layout, _LayoutEditor = the frame of the editors, _DocumentTabs)
  Resources/Texts/           the text catalog (en + de), one file per area; Js*.json = texts of the scripts (key Js.*)
  wwwroot/js/editor/         the shared frame of the editors: frame.js (title bar, ribbon builder, backstage, status bar, popovers, palette), session.js (load, autosave,
                             conflicts, offline queue, polling), util.js (t, el, icon, request, dialog, History …) – contract in docs/EDITORS.md
  wwwroot/js/<app>/          the editors (app.js = entry): docs, sheets, slides, base
  wwwroot/css/               app.css (accent, start row, gallery; prefix mo-), editor.css (the frame), <app>.css (the work area of each editor)
  wwwroot/sw.js              service worker: static cache, network-first for editor pages and /live/documents/{id} (copies for offline), offline.html
tests/MatOffice.Tests         unit + database tests (AccountTests, TextCatalogTests, per app <Kind>*Tests)
docs/EDITORS.md              the contract between the frame and the four editors (ribbon items, ctx, content check)
```

## Conventions

**Database (PostgreSQL)**: tables PascalCase and singular, key always `Id` BIGINT, every record has `CreateDate/CreateUserId/UpdateDate/UpdateUserId` (filled by
`AppDbContext.SaveChanges`), `UpdateState` (0 deleted, 1 created, 2 updated) for soft delete (`Document` = the trash with `DeletedDate`, `User`, `HostClient`), a `Token`
(uuid) for a security key (`UserSession.Token`, `Document.ShareToken`, `HostSession.Token`). Content of documents and versions is `jsonb`.

**Security**: every document query goes through `DocumentAccess`; a document you may not see answers 404. Every change of `/live` needs the antiforgery header
`RequestVerificationToken`; documents need a signed-in person with the editor right (checked in the business layer), a host file the secret of its session.
Saving compares the revision in the UPDATE itself (no lost update). The server normalizes every content (`IKindHandler.Normalize`): no script can come in through a document.
The access token of a host is never sent to a browser; API keys are stored as SHA-256.

**UI (MatDesign rules)**: CRUD = list page (toolbar above, table/gallery, pager under it, actions left below) + separate edit page (Save, Back, space, Delete); a document has
**tabs** (General, People, Link, Versions) that are pages of their own. The editors are the special UI of the app (like the office suites): their own page without the menu,
built on the tokens of the toolkit, with the colour of their app. Icons via `Ui.Icon` / `icon()` (toolkit sprite + `img/app-icons.svg` + `img/icons-<app>.svg`). Every `td` has `data-label`.
The version at the foot of the menu links to `/About`.

**Code**: Microsoft C# guidelines, file-scoped namespaces, `.editorconfig`. The business layer is static classes taking the `AppDbContext` and the actor. No database queries in page
models or views beyond calling the business layer. English identifiers and comments, UI texts only in the catalog.

**Compose files**: no environment variables to set, no placeholders; a stack (`name:`) per file; `docker-compose.yml` (release), `.dev.yml` (nightly + pgAdmin), `.local.yml`
(build + pgAdmin, subnet 10.253.49.0/24).

**Git**: commit under the user's name; never add a Claude/Anthropic co-author trailer or a "generated with" line (global rule of the user; it overrides any default). Branches:
`main` = release, `dev` = development. Versions: release `<major>.<minor>.<build>-<date>`, development `nightly-<build>-<date>`, local `local-<date>` (`Directory.Build.props`;
`.github/workflows/build.yml` builds, tests and pushes the image).
