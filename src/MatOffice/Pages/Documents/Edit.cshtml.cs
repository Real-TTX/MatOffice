using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Documents;
using MatOffice.Services.Formats;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>The tab "General" of a document: its name and what it is; moving it to the trash. Only the owner (the business layer checks it again).</summary>
public sealed class EditModel(AppDbContext db) : AppPageModel
{
    [BindProperty(SupportsGet = true)] public long Id { get; set; }

    [BindProperty] public string Name { get; set; } = string.Empty;

    public DocumentKind Kind { get; private set; }

    public string Slug => Kinds.SlugOf(Kind);

    public string EditorUrl => DocumentEndpoints.EditorUrl(Kind, Id);

    public DateTime Created { get; private set; }

    public DateTime Changed { get; private set; }

    public long Size { get; private set; }

    public string Zone => Me.TimeZone;

    public async Task<IActionResult> OnGetAsync() => await LoadAsync() ? Page() : NotFound();

    public async Task<IActionResult> OnPostAsync()
    {
        var saved = await DocumentService.RenameAsync(db, UserId, Id, Name);
        if (saved.NotAllowed)
        {
            return NotFound();
        }

        if (!saved.Ok)
        {
            ModelState.AddModelError(nameof(Name), saved.Error!);
            await LoadAsync(keepName: true);
            return Page();
        }

        Flash(Text.T("Common.Saved"));
        return LocalRedirect("/home");
    }

    public async Task<IActionResult> OnPostDeleteAsync()
    {
        if (!await LoadAsync() || !await DocumentService.TrashAsync(db, UserId, Id))
        {
            return NotFound();
        }

        Flash(Text.T("Documents.Trashed", Name));
        return LocalRedirect("/home");
    }

    private async Task<bool> LoadAsync(bool keepName = false)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(UserId), Id, DocumentRole.Owner) is not { } access)
        {
            return false;
        }

        var d = access.Document;
        if (!keepName)
        {
            Name = d.Name;
        }

        Kind = d.Kind;
        Created = d.CreateDate;
        Changed = d.ContentDate;
        Size = d.ContentLength;
        return true;
    }
}
