using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Documents;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>The version history of a document: the named versions and the snapshots the app keeps, each can be opened (read-only), named or restored.</summary>
public sealed class VersionsModel(AppDbContext db) : AppPageModel
{
    private const int PageSize = 25;

    [BindProperty(SupportsGet = true)] public long Id { get; set; }

    public string DocumentName { get; private set; } = string.Empty;

    public string EditorUrl { get; private set; } = "/home";

    public bool IsOwner { get; private set; }

    public bool CanEdit { get; private set; }

    public IReadOnlyList<VersionRow> Items { get; private set; } = [];

    public ToolbarModel Toolbar { get; private set; } = null!;

    public PagerModel Pager { get; private set; } = null!;

    public string Zone => Me.TimeZone;

    public async Task<IActionResult> OnGetAsync()
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(UserId), Id) is not { } access || await VersionService.ListAsync(db, Actor.Of(UserId), Id) is not { } all)
        {
            return NotFound();
        }

        DocumentName = access.Document.Name;
        EditorUrl = DocumentEndpoints.EditorUrl(access.Document.Kind, Id);
        IsOwner = access.IsOwner;
        CanEdit = access.CanEdit;

        var search = ListQuery.Text(Request, "q");
        var kind = ListQuery.Text(Request, "kind") is "named" or "auto" ? ListQuery.Text(Request, "kind") : null;
        var desc = ListQuery.Descending(Request, byDefault: true);
        var page = ListQuery.Page(Request);
        IEnumerable<VersionRow> rows = all;
        if (search is not null)
        {
            rows = rows.Where(v => v.Name.Contains(search, StringComparison.CurrentCultureIgnoreCase) || (v.Description ?? string.Empty).Contains(search, StringComparison.CurrentCultureIgnoreCase));
        }

        if (kind is not null)
        {
            rows = rows.Where(v => v.IsAuto == (kind == "auto"));
        }

        rows = desc ? rows.OrderByDescending(v => v.CreateDate) : rows.OrderBy(v => v.CreateDate);
        var matching = rows.ToList();
        Items = matching.Skip((page - 1) * PageSize).Take(PageSize).ToList();
        Toolbar = new ToolbarModel
        {
            Id = "versions",
            Action = "/Documents/Versions",
            Label = Text.T("Versions.Title"),
            Search = search,
            SearchField = true,
            Filters = [new FilterField("kind", Text.T("Versions.Col.Version"), [new("named", Text.T("Versions.Named")), new("auto", Text.T("Versions.Autos"))], kind)],
            SortOptions = [new("date", Text.T("Versions.Col.Date"))],
            Sort = "date",
            Descending = desc,
            DefaultSort = "date",
            DefaultDescending = true,
            Total = all.Count,
            Matching = matching.Count,
            Noun = Text.T("Versions.Many"),
            NounOne = Text.T("Versions.One"),
            Keep = new Dictionary<string, string?> { ["id"] = Id.ToString(System.Globalization.CultureInfo.InvariantCulture) },
        };
        Pager = new PagerModel { Page = page, PageSize = PageSize, Total = matching.Count, Url = Toolbar.UrlWithPage };
        return Page();
    }

    public async Task<IActionResult> OnPostRestoreAsync(long versionId)
    {
        if (!await VersionService.RestoreAsync(db, UserId, Id, versionId))
        {
            return NotFound();
        }

        Flash(Text.T("Versions.Restored"));
        return LocalRedirect($"/Documents/Versions?id={Id}");
    }
}
