using System.Text.Json;
using MatOffice.Services;
using MatPlatform;

namespace MatOffice.Platform;

/// <summary>
/// MatOffice as an app of the platform contract (MatDesign docs/PLATFORM.md, contract 2) on the shared package MatPlatform (the folder <c>MatPlatform/</c>, a copy):
/// how a consumer such as Mattiverse shows MatOffice (<see cref="Describe"/>: tiles, the few settings, buttons), what it provides (<c>office.edit</c>: the hosts, see
/// <see cref="MatOfficeHostBindings"/>) and consumes (MatIdentity, MatSMTP: set by their variables), what it adds to the home screen of a shell (the four apps
/// with their "New …" and two widgets for administrators), its values and what its buttons do.
/// The settings are the two of the admin page System (registration, public address), kept in the same place (<see cref="InstanceService"/>).
/// The values come from the services (no DbContext here: this is the web layer).
/// </summary>
public sealed class MatOfficePlatform(InstanceService instance, ILogger<MatOfficePlatform> log)
    : IMatPlatformState, IMatPlatformConfiguration, IMatPlatformActions, IMatPlatformHealth
{
    /// <summary>The id of MatOffice in the platform (and the default client id at MatIdentity), as in MatDesign <c>AppRegistry.cs</c>.</summary>
    public const string AppId = "matoffice";

    /// <summary>The accent of MatOffice (a calm office blue, white text 7.2:1; see <c>wwwroot/css/app.css</c>).</summary>
    public const string Accent = "#205a8e";

    /// <summary>The MatApiDefinition of MatOffice (<c>GET /api</c>); <c>auth</c> and <c>endpoints</c> are set by MatPlatform.</summary>
    public static void Describe(MatPlatformBuilder o) => Describe(o, null);

    /// <summary>The MatApiDefinition; <paramref name="publicUrl"/> = <c>MAT_PUBLIC_URL</c>: the field of the public address is fixed while it is set.</summary>
    public static void Describe(MatPlatformBuilder o, string? publicUrl)
    {
        var T = o.T;

        o.App(AppId, "MatOffice", T("Office im Browser: Texte, Tabellen, Präsentationen und Datenbanken", "Office in the browser: documents, spreadsheets, presentations and databases"), AppInfo.Version, "/img/logo/matoffice-mark.svg", Accent, "/home");

        o.Page("overview", T("Übersicht", "Overview"), "dashboard",
            MatSection.Tiles("users", "documents", "docs", "sheets", "slides", "bases", "sharedDocuments", "linkedDocuments", "versions", "hosts", "openHostFiles", "changedLast7Days", "registration", "lastActivity"),
            MatSection.Actions("openApp", "openUsers", "openConnections", "openSystem", "cleanUp"));
        o.Page("settings", T("Einstellungen", "Settings"), "settings",
            MatSection.Form(T("Allgemein", "General"), "registrationOpen", "publicUrl"));

        o.Tile("users", T("Aktive Benutzer", "Active users"), icon: "users");
        o.Tile("documents", T("Dokumente", "Documents"), icon: "file");
        o.Tile("docs", T("Texte", "Documents (text)"), icon: "file-text");
        o.Tile("sheets", T("Tabellen", "Spreadsheets"), icon: "table");
        o.Tile("slides", T("Präsentationen", "Presentations"), icon: "image");
        o.Tile("bases", T("Datenbanken", "Databases"), icon: "database");
        o.Tile("sharedDocuments", T("Geteilte Dokumente", "Shared documents"), icon: "users");
        o.Tile("linkedDocuments", T("Dokumente mit Link", "Documents with a link"), icon: "link");
        o.Tile("versions", T("Versionen", "Versions"), icon: "git");
        o.Tile("hosts", T("Verbundene Apps", "Connected apps"), icon: "plug");
        o.Tile("openHostFiles", T("Offene Dateien von Apps", "Open files of apps"), icon: "external");
        o.Tile("changedLast7Days", T("Geändert in 7 Tagen", "Changed in 7 days"), icon: "activity");
        o.Tile("registration", T("Registrierung", "Sign-up"), "status", icon: "user");
        o.Tile("lastActivity", T("Letzte Änderung an einem Dokument", "Last change to a document"), "time", icon: "activity");

        o.Field(new MatField("registrationOpen", T("Registrierung offen", "Open sign-up"), "bool")
        {
            Hint = T("Aus: nur Administratoren legen Konten an.", "Off: only administrators create accounts."),
        });
        var address = new MatField("publicUrl", T("Öffentliche Adresse", "Public address"), "url")
        {
            Max = 300,
            Placeholder = "https://office.example.com",
            Hint = T("Für die Adressen in E-Mails und für angebundene Apps; leer: die Adresse des Aufrufs.", "For the addresses in mails and for connected apps; empty: the address of the request."),
        };
        o.Field(publicUrl is null ? address : address with
        {
            Fixed = true,
            Hint = T("Festgelegt durch MAT_PUBLIC_URL (Mattiverse); hier lässt sie sich nicht ändern.", "Set by MAT_PUBLIC_URL (Mattiverse); it cannot be changed here."),
        });

        o.LinkAction("openApp", T("Öffnen", "Open"), "external", "/home");
        o.LinkAction("openUsers", T("Benutzer", "Users"), "users", "/Admin/Users");
        o.LinkAction("openConnections", T("Verbindungen", "Connections"), "plug", "/Admin/Connections");
        o.LinkAction("openSystem", T("System", "System"), "settings", "/Admin/System");
        o.PostAction(
            "cleanUp",
            T("Aufräumen", "Clean up"),
            "broom",
            T("Beendete Sitzungen, abgelaufene App-Dateien und Dokumente, die seit mehr als 30 Tagen im Papierkorb liegen, jetzt endgültig entfernen?",
              "Remove ended sessions, expired files of apps and the documents that lie in the trash for more than 30 days for good now?"));

        o.Role("admin", T("Administrator", "Administrator"));
        o.Role("user", T("Benutzer", "User"));

        // What MatOffice offers the other apps: its editors for their files (a binding is a host, MatOfficeHostBindings), and what it uses of them. Sign-in and
        // mail stay set by their variables (MATIDENTITY_*, MATSMTP_*); naming them here shows them in the Connections area.
        o.Provides(MatCapabilities.OfficeEdit, label: T("Bearbeitet Office-Dateien für", "Edits office files for"));
        o.Consumes(MatCapabilities.IdentityOidc, optional: true);
        o.Consumes(MatCapabilities.MailSmtp, optional: true);

        // The home screen of a shell: an icon per app (its list of documents, the picture of its tab), "New …" in the menu of each icon (the page New with the app
        // chosen), and two widgets for administrators from the values of /api/state.
        Launcher(o, "docs", T("Texte", "Docs"), T("Neues Dokument", "New document"), "#185abd");
        Launcher(o, "sheets", T("Tabellen", "Sheets"), T("Neue Tabelle", "New spreadsheet"), "#107c41");
        Launcher(o, "slides", T("Präsentationen", "Slides"), T("Neue Präsentation", "New presentation"), "#c43e1c");
        Launcher(o, "base", T("Datenbank", "Base"), T("Neue Datenbank", "New database"), "#a4373a");
        o.Widget(new MatWidget("documents", T("Dokumente", "Documents"), MatWidgetKinds.Tiles, "medium")
        {
            Items = ["docs", "sheets", "slides", "bases"],
            Icon = "file",
            Url = "/Admin/System",
            Refresh = 300,
            Audience = MatAudience.Admin,
        });
        o.Widget(new MatWidget("users", T("Aktive Benutzer", "Active users"), MatWidgetKinds.Stat)
        {
            Items = ["users"],
            Icon = "users",
            Url = "/Admin/Users",
            Refresh = 300,
            Audience = MatAudience.Admin,
        });
    }

    /// <summary>The launcher of one app (<c>/docs</c> …: its documents and templates) with its "New …" (<c>/Documents/New?kind=docs</c>: the app chosen).</summary>
    private static void Launcher(MatPlatformBuilder o, string slug, string label, string newLabel, string accent)
    {
        o.Launcher(slug, label, $"/img/logo/{slug}-favicon.svg", "/" + slug, accent: accent);
        o.QuickAction(slug, newLabel, "plus", url: "/Documents/New?kind=" + slug);
    }

    public async Task<IReadOnlyList<MatHealthCheck>> CheckHealthAsync(MatPlatformCall call, CancellationToken ct)
    {
        if (!await instance.CanConnectAsync(ct))
        {
            return [new MatHealthCheck("database", "down", call.T("Die Datenbank ist nicht erreichbar.", "The database cannot be reached."))];
        }

        var checks = new List<MatHealthCheck> { new("database", "ok") };
        if (await instance.AnyUserAsync() && (await instance.SnapshotAsync(ct)).Admins == 0)
        {
            checks.Add(new MatHealthCheck("administrators", "degraded", call.T("Kein aktiver Administrator.", "No active administrator.")));
        }

        return checks;
    }

    public async Task<IReadOnlyDictionary<string, MatStateValue>> GetStateAsync(MatPlatformCall call, CancellationToken ct)
    {
        var s = await instance.SnapshotAsync(ct);
        var (open, _) = await instance.SettingsAsync();
        return new Dictionary<string, MatStateValue>
        {
            ["users"] = new(s.Users),
            ["documents"] = new(s.Documents),
            ["docs"] = new(s.Docs),
            ["sheets"] = new(s.Sheets),
            ["slides"] = new(s.Slides),
            ["bases"] = new(s.Bases),
            ["sharedDocuments"] = new(s.SharedDocuments),
            ["linkedDocuments"] = new(s.LinkedDocuments),
            ["versions"] = new(s.Versions),
            ["hosts"] = new(s.Hosts),
            ["openHostFiles"] = new(s.OpenHostFiles),
            ["changedLast7Days"] = new(s.ChangedLast7Days),
            ["registration"] = open
                ? new MatStateValue("info", Text: call.T("Offen", "Open"))
                : new MatStateValue("ok", Text: call.T("Geschlossen", "Closed")),
            ["lastActivity"] = new(s.LastActivityUtc is { } last ? new DateTimeOffset(DateTime.SpecifyKind(last, DateTimeKind.Utc)) : null),
        };
    }

    public async Task<IReadOnlyDictionary<string, object?>> GetConfigurationAsync(MatPlatformCall call, CancellationToken ct)
    {
        var (open, url) = await instance.SettingsAsync();
        return new Dictionary<string, object?>
        {
            ["registrationOpen"] = open,
            ["publicUrl"] = url ?? string.Empty,
        };
    }

    /// <summary>The values are checked against the definition already (MatPlatform, <c>CheckValues</c>); the address is checked once more like on the admin page.</summary>
    public async Task<IReadOnlyDictionary<string, string>> ConfigureAsync(IReadOnlyDictionary<string, JsonElement> values, MatPlatformCall call, CancellationToken ct)
    {
        var errors = new Dictionary<string, string>();
        bool? open = null;
        string? url = null;
        foreach (var (key, value) in values)
        {
            switch (key)
            {
                case "registrationOpen" when value.ValueKind is JsonValueKind.True or JsonValueKind.False:
                    open = value.GetBoolean();
                    break;
                case "publicUrl" when instance.PublicUrlFixed:
                    errors[key] = call.T("Festgelegt durch MAT_PUBLIC_URL.", "Set by MAT_PUBLIC_URL.");
                    break;
                case "publicUrl" when value.ValueKind is JsonValueKind.String:
                    url = value.GetString()!;
                    break;
                case "publicUrl" when value.ValueKind is JsonValueKind.Null:
                    url = string.Empty;
                    break;
                default:
                    errors[key] = call.T("Dieser Wert passt nicht.", "This value does not fit.");
                    break;
            }
        }

        if (url is not null && InstanceService.NormalizePublicUrl(url) is { } normalized && !InstanceService.IsPublicUrl(normalized))
        {
            errors["publicUrl"] = call.T("Eine vollständige Adresse mit http:// oder https://.", "A complete address with http:// or https://.");
        }

        if (errors.Count > 0)
        {
            return errors;
        }

        await instance.SaveSettingsAsync(open, url);

        // MatOffice has no activity log of its own: the change (the keys, no values) goes to the log of the app.
        log.LogInformation("Settings changed through the platform API by {Caller}: {Keys}", call.Caller, string.Join(", ", values.Keys));
        return errors;
    }

    public async Task<MatActionResult?> RunActionAsync(string id, MatPlatformCall call, CancellationToken ct)
    {
        if (id != "cleanUp")
        {
            return null;
        }

        var result = await instance.CleanUpAsync(ct);
        log.LogInformation("Clean-up through the platform API by {Caller}: {Sessions} sessions, {Documents} documents", call.Caller, result.Sessions, result.Documents);
        return new MatActionResult(true, call.T(
            $"Aufgeräumt: {result.Sessions} Sitzungen, {result.Documents} Dokumente, {result.Versions} Zwischenstände.",
            $"Cleaned up: {result.Sessions} sessions, {result.Documents} documents, {result.Versions} snapshots."));
    }
}
