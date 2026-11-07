using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Hosts;

namespace MatOffice.Pages.Admin.Hosts;

/// <summary>The apps that open their files in MatOffice (MatDrive and others) with their API keys. Administrators only.</summary>
public sealed class IndexModel(AppDbContext db, InstanceService instance) : AppPageModel
{
    private const int PageSize = 25;

    public ToolbarModel Toolbar { get; private set; } = null!;

    public PagerModel Pager { get; private set; } = null!;

    public List<HostClientRow> Items { get; private set; } = [];

    public string ApiAddress { get; private set; } = string.Empty;

    public string Zone => Me.TimeZone;

    public async Task OnGetAsync()
    {
        var search = ListQuery.Text(Request, "q");
        var state = ListQuery.Text(Request, "state") is "active" or "off" ? ListQuery.Text(Request, "state") : null;
        var desc = ListQuery.Descending(Request);
        var page = ListQuery.Page(Request);
        var all = await HostAdmin.ListAsync(db, UserId, HttpContext.RequestAborted) ?? [];
        IEnumerable<HostClientRow> rows = all;
        if (search is not null)
        {
            rows = rows.Where(h => h.Name.Contains(search, StringComparison.CurrentCultureIgnoreCase) || h.KeyPrefix.Contains(search, StringComparison.Ordinal));
        }

        if (state is not null)
        {
            rows = rows.Where(h => h.IsActive == (state == "active"));
        }

        rows = desc ? rows.OrderByDescending(h => h.Name, StringComparer.CurrentCultureIgnoreCase) : rows.OrderBy(h => h.Name, StringComparer.CurrentCultureIgnoreCase);
        var matching = rows.ToList();
        Items = matching.Skip((page - 1) * PageSize).Take(PageSize).ToList();

        var (_, publicUrl) = await instance.SettingsAsync();
        ApiAddress = (publicUrl ?? $"{Request.Scheme}://{Request.Host}") + "/api/v1/host";
        Toolbar = new ToolbarModel
        {
            Id = "hosts",
            Action = "/Admin/Hosts",
            Label = Text.T("Hosts.Title"),
            Search = search,
            SearchField = true,
            SearchPlaceholder = Text.T("Hosts.Search"),
            Filters = [new FilterField("state", Text.T("Common.Status"), [new("active", Text.T("Common.Active")), new("off", Text.T("Hosts.Off"))], state)],
            SortOptions = [new("name", Text.T("Common.Name"))],
            Sort = "name",
            Descending = desc,
            DefaultSort = "name",
            Total = all.Count,
            Matching = matching.Count,
            Noun = Text.T("Hosts.Many"),
            NounOne = Text.T("Hosts.One"),
        };
        Pager = new PagerModel { Page = page, PageSize = PageSize, Total = matching.Count, Url = Toolbar.UrlWithPage };
    }
}
