namespace MatOffice.Pages.Admin;

/// <summary>
/// The Connections area (MatPlatform 2.0): what MatOffice uses of other apps (sign-in, mail), what it offers them (<c>office.edit</c>: the hosts such as MatDrive, made
/// by Mattiverse or here by hand, with a new key and remove) and what it adds to a home screen. The content is the shared <c>&lt;mat-connections /&gt;</c> of MatPlatform,
/// the same in every Mat app; its forms post to MatPlatform. Administrators only (the folder, and <c>ConnectionsAccess</c> in Program.cs).
/// </summary>
public sealed class ConnectionsModel : AppPageModel
{
    public void OnGet()
    {
    }
}
