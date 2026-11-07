using MatOffice.Controls;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services.Documents;
using MatOffice.Services.Formats;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Documents;

/// <summary>A new document: the app, a template and a name, or a file that is read into a new document (Word, Excel, PowerPoint, CSV ... and the own formats).</summary>
public sealed class NewModel(AppDbContext db) : AppPageModel
{
    public const long MaxUpload = 100L * 1024 * 1024;

    [BindProperty(SupportsGet = true)] public string Kind { get; set; } = "docs";

    [BindProperty] public string Name { get; set; } = string.Empty;

    /// <summary>The template per app (one field per app, only the one of the chosen app is shown and read).</summary>
    [BindProperty] public Dictionary<string, string> Template { get; set; } = [];

    [BindProperty] public IFormFile? Upload { get; set; }

    public List<Choice> KindChoices { get; } = Kinds.All.Select(h => new Choice(h.Slug, Text.T("Kinds." + h.Kind))).ToList();

    /// <summary>The extensions the upload takes, for the <c>accept</c> of the field.</summary>
    public string Accept { get; } = string.Join(',', Kinds.All.SelectMany(h => h.Formats.Where(f => f.CanImport).Select(f => f.Extension).Append(h.NativeExtension)).Distinct());

    public IActionResult OnGet()
    {
        Kind = Kinds.BySlug(Kind)?.Slug ?? "docs";
        return Page();
    }

    public List<Choice> TemplatesOf(IKindHandler h) => h.Templates.Select(t => new Choice(t.Id, Text.T(t.Name))).ToList();

    public async Task<IActionResult> OnPostAsync()
    {
        if (Kinds.BySlug(Kind) is not { } handler)
        {
            ModelState.AddModelError(nameof(Kind), Text.T("Documents.Kind.Required"));
            return Page();
        }

        var created = await DocumentService.CreateAsync(db, UserId, handler.Kind, Template.GetValueOrDefault(handler.Slug), Name, Languages.Current);
        if (!created.Ok)
        {
            ModelState.AddModelError(nameof(Name), created.Error!);
            return Page();
        }

        return LocalRedirect(DocumentEndpoints.EditorUrl(handler.Kind, created.Value!.Id));
    }

    public async Task<IActionResult> OnPostUploadAsync()
    {
        if (Upload is null || Upload.Length == 0)
        {
            ModelState.AddModelError(nameof(Upload), Text.T("Documents.Import.Required"));
            return Page();
        }

        if (Upload.Length > MaxUpload)
        {
            ModelState.AddModelError(nameof(Upload), Text.T("Documents.TooLarge"));
            return Page();
        }

        await using var stream = Upload.OpenReadStream();
        var imported = await DocumentService.ImportAsync(db, UserId, Upload.FileName, stream);
        if (!imported.Ok)
        {
            ModelState.AddModelError(nameof(Upload), imported.Error!);
            return Page();
        }

        Flash(Text.T("Documents.Imported", imported.Value!.Name));
        return LocalRedirect(DocumentEndpoints.EditorUrl(imported.Value.Kind, imported.Value.Id));
    }
}
