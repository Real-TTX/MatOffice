using MatOffice.Data;
using MatOffice.Services;
using MatOffice.Services.Documents;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MatOffice.Pages.Public;

/// <summary>
/// A document opened through its link (<c>/s/&lt;secret&gt;</c>), by anybody: read-only for a guest, editable for a signed-in person when the link allows it.
/// A link that is off, or wrong, answers 404 exactly like a document that is not there.
/// </summary>
public sealed class LinkModel(AppDbContext db, CurrentUser me) : PageModel
{
    public async Task<IActionResult> OnGetAsync(Guid token)
    {
        if (await DocumentAccess.FindByLinkAsync(db, new Actor(me.Id), token) is not { } viaLink)
        {
            return NotFound();
        }

        var d = viaLink.Document;
        ViewData["Editor"] = new EditorPage(d.Kind, d.Name, $"/live/documents/{d.Id}?t={token}", "document", me.IsAuthenticated ? "/home" : "/", GuestName: me.IsAuthenticated ? null : string.Empty);
        Response.Headers.CacheControl = "no-store";
        return Page();
    }
}
