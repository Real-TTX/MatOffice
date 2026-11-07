using MatOffice.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MatOffice.Pages;

/// <summary>The start page: for a visitor what MatOffice is, for a signed-in person straight to the documents.</summary>
public sealed class IndexModel(AccountService accounts, CurrentUser me) : PageModel
{
    public bool RegistrationOpen { get; private set; }

    public async Task<IActionResult> OnGetAsync()
    {
        if (me.IsAuthenticated)
        {
            return LocalRedirect("/home");
        }

        RegistrationOpen = !await accounts.AnyUserAsync() || (await accounts.SettingsAsync()).RegistrationOpen;
        return Page();
    }
}
