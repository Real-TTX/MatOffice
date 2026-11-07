using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Documents;
using MatOffice.Services.Formats;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>
/// The start page (<c>/home</c>) and the start page of each app (<c>/docs</c>, <c>/sheets</c>, <c>/slides</c>, <c>/base</c>): a row to start something new (the blank
/// documents of the four apps, or the templates of one app), and the documents as pictures or as a table with search, filter (whose, starred, recent, trash) and sort.
/// </summary>
public sealed class IndexModel(AppDbContext db) : AppPageModel
{
    private const int PageSize = 24;
    private static readonly string[] Sorts = ["modified", "opened", "name", "created"];
    private static readonly string[] Scopes = ["mine", "shared", "starred", "recent", "trash"];

    [BindProperty(SupportsGet = true)] public string? Kind { get; set; }

    public IKindHandler? Handler { get; private set; }

    public DocumentScope Scope { get; private set; }

    public ToolbarModel Toolbar { get; private set; } = null!;

    public PagerModel Pager { get; private set; } = null!;

    public DocumentPage Listing { get; private set; } = new([], 0);

    public string Zone => Me.TimeZone;

    /// <summary>The address of this list without the query.</summary>
    public string Action => Handler is null ? "/home" : "/" + Handler.Slug;

    public async Task<IActionResult> OnGetAsync()
    {
        Handler = Kinds.BySlug(Kind);
        var search = ListQuery.Text(Request, "q");
        var scopeText = ListQuery.Text(Request, "scope") is { } st && Scopes.Contains(st) ? st : null;
        Scope = scopeText is null ? DocumentScope.All : Enum.Parse<DocumentScope>(scopeText, ignoreCase: true);
        var kindFilter = Handler is null && ListQuery.Text(Request, "kind") is { } kf && Kinds.BySlug(kf) is { } kh ? kh : null;
        var sort = ListQuery.Text(Request, "sort") is { } s && Sorts.Contains(s) ? s : "modified";
        var desc = ListQuery.Descending(Request, byDefault: sort != "name");
        var page = ListQuery.Page(Request);

        var query = new DocumentQuery(Scope, (Handler ?? kindFilter)?.Kind, search, Enum.Parse<DocumentSort>(sort, ignoreCase: true), desc, page, PageSize);
        Listing = await DocumentService.ListAsync(db, UserId, query);

        var all = Scope == DocumentScope.All && search is null && kindFilter is null ? Listing.Total : (await DocumentService.ListAsync(db, UserId, query with { Scope = DocumentScope.All, Kind = Handler?.Kind, Search = null, Page = 1, PageSize = 1 })).Total;
        var filters = new List<FilterField>
        {
            new("scope", Text.T("Documents.Filter.Show"), Scopes.Select(x => new FilterOption(x, Text.T("Documents.Scope." + x))).ToList(), scopeText),
        };
        if (Handler is null)
        {
            filters.Add(new FilterField("kind", Text.T("Documents.Col.Kind"), Kinds.All.Select(h => new FilterOption(h.Slug, Text.T("Kinds." + h.Kind))).ToList(), kindFilter?.Slug));
        }

        Toolbar = new ToolbarModel
        {
            Id = "documents",
            Action = Action,
            Label = Text.T("Documents.Title"),
            Search = search,
            SearchField = true,
            SearchPlaceholder = Text.T("Documents.Search"),
            Filters = filters,
            SortOptions = [new("modified", Text.T("Documents.Col.Changed")), new("opened", Text.T("Documents.Sort.Opened")), new("name", Text.T("Common.Name")), new("created", Text.T("Documents.Created"))],
            Sort = sort,
            Descending = desc,
            DefaultSort = "modified",
            DefaultDescending = true,
            Views = [new ViewOption("gallery", Text.T("Lists.View.Gallery"), "grid"), new ViewOption("list", Text.T("Lists.View.Table"), "list")],
            Total = all,
            Matching = Listing.Total,
            Noun = Text.T("Documents.Many"),
            NounOne = Text.T("Documents.One"),
        };
        Pager = new PagerModel { Page = page, PageSize = PageSize, Total = Listing.Total, Url = Toolbar.UrlWithPage };
        return Page();
    }

    /// <summary>The templates of the row "Start something new": the blank document of each app on /home, the templates of one app on its own page.</summary>
    public IEnumerable<(IKindHandler Handler, KindTemplate Template)> Starters()
        => Handler is { } h ? h.Templates.Select(t => (h, t)) : Kinds.All.Select(k => (k, k.Templates[0]));

    public async Task<IActionResult> OnPostCreateAsync(string kind, string? template)
    {
        if (Kinds.BySlug(kind) is not { } handler)
        {
            return NotFound();
        }

        var created = await DocumentService.CreateAsync(db, UserId, handler.Kind, template, null, Languages.Current);
        if (!created.Ok)
        {
            FlashError(created.Error!);
            return LocalRedirect("/" + handler.Slug);
        }

        return LocalRedirect(DocumentEndpoints.EditorUrl(handler.Kind, created.Value!.Id));
    }

    public async Task<IActionResult> OnPostStarAsync(long id, bool starred, string? back)
    {
        if (!await DocumentService.StarAsync(db, UserId, id, starred))
        {
            return NotFound();
        }

        return Back(back, "/home");
    }

    public async Task<IActionResult> OnPostDuplicateAsync(long id)
    {
        var copy = await DocumentService.DuplicateAsync(db, UserId, id);
        if (!copy.Ok)
        {
            return NotFound();
        }

        Flash(Text.T("Documents.Duplicated", copy.Value!.Name));
        return LocalRedirect(DocumentEndpoints.EditorUrl(copy.Value.Kind, copy.Value.Id));
    }

    public async Task<IActionResult> OnPostTrashAsync(long id, string? back)
    {
        if (!await DocumentService.TrashAsync(db, UserId, id))
        {
            return NotFound();
        }

        Flash(Text.T("Documents.TrashedShort"));
        return Back(back, "/home");
    }

    public async Task<IActionResult> OnPostLeaveAsync(long id)
    {
        if (!await ShareService.RemoveMemberAsync(db, UserId, id, UserId))
        {
            return NotFound();
        }

        Flash(Text.T("Documents.Left"));
        return LocalRedirect("/home");
    }

    public async Task<IActionResult> OnPostRestoreAsync(long id, string? back)
    {
        if (!await DocumentService.RestoreAsync(db, UserId, id))
        {
            return NotFound();
        }

        Flash(Text.T("Documents.Restored"));
        return Back(back, "/home?scope=trash");
    }

    public async Task<IActionResult> OnPostDeleteForeverAsync(long id, string? back)
    {
        if (!await DocumentService.DeleteForeverAsync(db, UserId, id))
        {
            return NotFound();
        }

        Flash(Text.T("Documents.DeletedForever"));
        return Back(back, "/home?scope=trash");
    }

    public async Task<IActionResult> OnPostEmptyTrashAsync()
    {
        var count = await DocumentService.EmptyTrashAsync(db, UserId);
        Flash(Text.T("Documents.TrashEmptied", count));
        return LocalRedirect("/home?scope=trash");
    }
}
