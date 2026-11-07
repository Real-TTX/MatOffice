using MatOffice.Controls;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Documents;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>The link of a model: whether anybody who has it may look at the model (read-only), and a new secret when the old address got out. Only the owner.</summary>
public sealed class ShareModel(AppDbContext db, InstanceService instance) : AppPageModel
{
    [BindProperty(SupportsGet = true)] public long Id { get; set; }

    [BindProperty] public string Mode { get; set; } = nameof(LinkMode.Off);

    public string DocumentName { get; private set; } = string.Empty;

    public string EditorUrl { get; private set; } = "/home";

    public string LinkUrl { get; private set; } = string.Empty;

    public List<Choice> Modes { get; } =
    [
        new(nameof(LinkMode.Off), Text.T("Share.Mode.Off")),
        new(nameof(LinkMode.View), Text.T("Share.Mode.View")),
        new(nameof(LinkMode.Edit), Text.T("Share.Mode.Edit")),
    ];

    public async Task<IActionResult> OnGetAsync() => await LoadAsync() ? Page() : NotFound();

    public Task<IActionResult> OnPostAsync() => ChangeAsync(renew: false);

    public Task<IActionResult> OnPostRenewAsync() => ChangeAsync(renew: true);

    private async Task<IActionResult> ChangeAsync(bool renew)
    {
        var mode = Enum.TryParse<LinkMode>(Mode, out var parsed) && Enum.IsDefined(parsed) ? parsed : LinkMode.Off;
        if (await ShareService.SetLinkAsync(db, UserId, Id, mode, renew) is null)
        {
            return NotFound();
        }

        Flash(Text.T(renew ? "Share.Renewed" : "Common.Saved"));
        return LocalRedirect($"/Documents/Share?id={Id}");
    }

    private async Task<bool> LoadAsync()
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(UserId), Id, DocumentRole.Owner) is not { } access)
        {
            return false;
        }

        DocumentName = access.Document.Name;
        EditorUrl = DocumentEndpoints.EditorUrl(access.Document.Kind, Id);
        Mode = access.Document.LinkMode.ToString();
        var (_, publicUrl) = await instance.SettingsAsync();
        LinkUrl = $"{publicUrl ?? $"{Request.Scheme}://{Request.Host}"}/s/{access.Document.ShareToken}";
        return true;
    }
}
