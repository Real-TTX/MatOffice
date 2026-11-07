using MatOffice.Data;
using MatOffice.Services.Hosts;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MatOffice.Pages;

/// <summary>
/// A file of a host (MatDrive) in the editor of its app: <c>/h/&lt;secret of the session&gt;</c>, opened by the browser after the host started the session. No account
/// of MatOffice is needed: the host decided who may open and change the file; the secret of the address is the right.
/// </summary>
public sealed class HostModel(AppDbContext db) : PageModel
{
    public async Task<IActionResult> OnGetAsync(Guid token)
    {
        if (await HostSessionService.FindAsync(db, token) is not { } session)
        {
            return NotFound();
        }

        ViewData["Editor"] = new EditorPage(session.Kind, Path.GetFileNameWithoutExtension(session.FileName), "/live/host/" + token.ToString("N"), "host", session.ReturnUrl ?? "/",
            GuestName: session.UserName);
        Response.Headers.CacheControl = "no-store";
        Response.Headers["Referrer-Policy"] = "no-referrer";
        return Page();
    }
}
