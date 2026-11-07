using MatOffice.Data;
using MatOffice.Services.Documents;
using MatOffice.Services.Formats;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages;

/// <summary>
/// The editor of a document (<c>/docs/12</c>, <c>/sheets/12</c>, <c>/slides/12</c>, <c>/base/12</c>; <c>?v=</c> a version, read-only). Signed-in people only (a link
/// opens <c>/s/&lt;secret&gt;</c>). A document the person may not see answers 404; a wrong app in the address leads to the right one.
/// </summary>
public sealed class EditorModel(AppDbContext db) : AppPageModel
{
    [BindProperty(SupportsGet = true)] public string Kind { get; set; } = string.Empty;

    [BindProperty(SupportsGet = true)] public long Id { get; set; }

    [BindProperty(SupportsGet = true, Name = "v")] public long? VersionId { get; set; }

    public async Task<IActionResult> OnGetAsync()
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(UserId), Id) is not { } access)
        {
            return NotFound();
        }

        var d = access.Document;
        var slug = Kinds.SlugOf(d.Kind);
        if (!string.Equals(slug, Kind, StringComparison.Ordinal))
        {
            return LocalRedirect($"/{slug}/{Id}{Request.QueryString}");
        }

        var api = $"/live/documents/{Id}" + (VersionId is { } v ? $"?v={v}" : string.Empty);
        ViewData["Editor"] = new EditorPage(d.Kind, d.Name, api, VersionId is null ? "document" : "version", "/" + slug, access.IsOwner ? $"/Documents/Members?id={Id}" : null);
        Response.Headers.CacheControl = "no-store";
        return Page();
    }
}
