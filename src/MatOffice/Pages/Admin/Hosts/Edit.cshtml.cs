using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services;
using MatOffice.Services.Hosts;
using Microsoft.AspNetCore.Mvc;

namespace MatOffice.Pages.Admin.Hosts;

/// <summary>A connected app: its name, whether it may open files, the addresses it may hand over, a new key, delete. The key is shown once, right after it was made.</summary>
public sealed class EditModel(AppDbContext db, InstanceService instance) : AppPageModel
{
    [BindProperty(SupportsGet = true)] public long? Id { get; set; }

    [BindProperty] public string Name { get; set; } = "MatDrive";

    [BindProperty] public bool IsActive { get; set; } = true;

    [BindProperty] public string AllowedUrlPrefix { get; set; } = string.Empty;

    [BindProperty] public string Note { get; set; } = string.Empty;

    public bool IsNew => Id is null;

    /// <summary>The key, only right after it was made (TempData survives the one redirect).</summary>
    public string? NewKey { get; private set; }

    public string KeyPrefix { get; private set; } = string.Empty;

    public string Address { get; private set; } = string.Empty;

    public async Task<IActionResult> OnGetAsync()
    {
        NewKey = TempData["HostKey"] as string;
        if (Id is not { } id)
        {
            return Page();
        }

        if (await HostAdmin.FindAsync(db, UserId, id) is not { } client)
        {
            return NotFound();
        }

        Name = client.Name;
        IsActive = client.IsActive;
        AllowedUrlPrefix = client.AllowedUrlPrefix ?? string.Empty;
        Note = client.Note ?? string.Empty;
        KeyPrefix = client.KeyPrefix;
        var (_, publicUrl) = await instance.SettingsAsync();
        Address = publicUrl ?? $"{Request.Scheme}://{Request.Host}";
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        var input = new HostClientInput(Name, IsActive, AllowedUrlPrefix, Note);
        if (Id is { } id)
        {
            var saved = await HostAdmin.UpdateAsync(db, UserId, id, input);
            if (saved.NotAllowed)
            {
                return NotFound();
            }

            if (!saved.Ok)
            {
                ModelState.AddModelError(nameof(Name), saved.Error!);
                return Page();
            }

            Flash(Text.T("Common.Saved"));
            return LocalRedirect("/Admin/Hosts");
        }

        var created = await HostAdmin.CreateAsync(db, UserId, input);
        if (created.NotAllowed)
        {
            return NotFound();
        }

        if (!created.Ok)
        {
            ModelState.AddModelError(nameof(Name), created.Error!);
            return Page();
        }

        TempData["HostKey"] = created.Value.Key;
        Flash(Text.T("Hosts.Created"));
        return LocalRedirect($"/Admin/Hosts/Edit?id={created.Value.Client.Id}");
    }

    public async Task<IActionResult> OnPostRenewAsync()
    {
        if (Id is not { } id || await HostAdmin.RenewKeyAsync(db, UserId, id) is not { } key)
        {
            return NotFound();
        }

        TempData["HostKey"] = key;
        Flash(Text.T("Hosts.Renewed"));
        return LocalRedirect($"/Admin/Hosts/Edit?id={id}");
    }

    public async Task<IActionResult> OnPostDeleteAsync()
    {
        if (Id is not { } id || !await HostAdmin.DeleteAsync(db, UserId, id))
        {
            return NotFound();
        }

        Flash(Text.T("Common.DeletedNamed", Name));
        return LocalRedirect("/Admin/Hosts");
    }
}
