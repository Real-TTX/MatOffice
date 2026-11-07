using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Documents;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>Create a named version of the content now, or name / describe / delete a version. Editors of the document (the business layer checks it again).</summary>
public sealed class VersionModel(AppDbContext db) : AppPageModel
{
    [BindProperty(SupportsGet = true)] public long DocumentId { get; set; }

    [BindProperty(SupportsGet = true)] public long? Id { get; set; }

    [BindProperty] public string Name { get; set; } = string.Empty;

    [BindProperty] public string Description { get; set; } = string.Empty;

    public string DocumentName { get; private set; } = string.Empty;

    public bool IsNew => Id is null;

    public async Task<IActionResult> OnGetAsync()
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(UserId), DocumentId, DocumentRole.Editor) is not { } access)
        {
            return NotFound();
        }

        DocumentName = access.Document.Name;
        if (Id is { } id)
        {
            if (await VersionService.FindAsync(db, Actor.Of(UserId), DocumentId, id) is not { } version)
            {
                return NotFound();
            }

            Name = version.Name;
            Description = version.Description ?? string.Empty;
        }

        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        Outcome<bool> result;
        if (Id is { } id)
        {
            result = await VersionService.UpdateAsync(db, UserId, DocumentId, id, Name, Description);
        }
        else
        {
            var created = await VersionService.CreateAsync(db, UserId, DocumentId, Name, Description);
            result = created.NotAllowed ? Outcome<bool>.Denied() : created.Ok ? Outcome<bool>.Success(true) : Outcome<bool>.Fail(created.Error!);
        }

        if (result.NotAllowed)
        {
            return NotFound();
        }

        if (!result.Ok)
        {
            ModelState.AddModelError(nameof(Name), result.Error!);
            DocumentName = (await DocumentAccess.FindAsync(db, Actor.Of(UserId), DocumentId))?.Document.Name ?? string.Empty;
            return Page();
        }

        Flash(Text.T(IsNew ? "Versions.Created" : "Common.Saved"));
        return LocalRedirect($"/Documents/Versions?id={DocumentId}");
    }

    public async Task<IActionResult> OnPostDeleteAsync()
    {
        if (Id is not { } id || !await VersionService.DeleteAsync(db, UserId, DocumentId, id))
        {
            return NotFound();
        }

        Flash(Text.T("Versions.Deleted"));
        return LocalRedirect($"/Documents/Versions?id={DocumentId}");
    }
}
