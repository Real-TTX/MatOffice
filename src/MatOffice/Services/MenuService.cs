using MatOffice.Data;
using MatOffice.Services.Documents;

namespace MatOffice.Services;

/// <summary>What the menu shows: the documents the person opened last and how many there are.</summary>
public sealed class MenuModel
{
    public IReadOnlyList<(long Id, string Name, DocumentKind Kind)> Recent { get; init; } = [];

    public int Mine { get; init; }

    public int Shared { get; init; }

    public int Starred { get; init; }

    public int Trash { get; init; }
}

/// <summary>One menu of the switcher at the foot of the menu.</summary>
public sealed record MenuArea(string Id, string Title, string Icon);

public sealed class MenuService(AppDbContext db, CurrentUser me)
{
    private MenuModel? _model;

    /// <summary>The menus of the switcher: the documents, the settings of the person and, for an administrator, the administration.</summary>
    public static IReadOnlyList<MenuArea> Areas(bool admin)
    {
        var list = new List<MenuArea>
        {
            new("documents", Localization.Text.T("Shell.Area.Documents"), "file"),
            new("settings", Localization.Text.T("Settings.Title"), "settings"),
        };
        if (admin)
        {
            list.Add(new MenuArea("admin", Localization.Text.T("Admin.Title"), "shield"));
        }

        return list;
    }

    /// <summary>The menu a page belongs to, by its address.</summary>
    public static string AreaOf(string? path)
    {
        path ??= "/";
        bool Under(string prefix) => path.Equals(prefix, StringComparison.OrdinalIgnoreCase) || path.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase);
        if (Under("/Admin"))
        {
            return "admin";
        }

        return Under("/Account") ? "settings" : "documents";
    }

    public async Task<MenuModel> GetAsync()
    {
        if (_model is not null)
        {
            return _model;
        }

        var userId = me.RequiredId;
        var (mine, shared, starred, trash) = await DocumentService.CountsAsync(db, userId);
        _model = new MenuModel
        {
            Recent = await DocumentService.RecentAsync(db, userId, 6),
            Mine = mine,
            Shared = shared,
            Starred = starred,
            Trash = trash,
        };
        return _model;
    }
}
