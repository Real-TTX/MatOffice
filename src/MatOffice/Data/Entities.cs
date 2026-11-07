namespace MatOffice.Data;

/// <summary>What every record has: an Id (BIGINT) and who created and changed it, and when (the columns are filled by <see cref="AppDbContext"/>).</summary>
public abstract class Entity
{
    public long Id { get; set; }

    public DateTime CreateDate { get; set; }

    public long? CreateUserId { get; set; }

    public DateTime UpdateDate { get; set; }

    public long? UpdateUserId { get; set; }
}

/// <summary>A record that supports soft delete: 0 = deleted, 1 = created, 2 = updated. A deleted record is hidden by a query filter.</summary>
public abstract class SoftEntity : Entity
{
    public const short Deleted = 0;
    public const short Created = 1;
    public const short Updated = 2;

    public short UpdateState { get; set; } = Created;
}

public enum UserRole
{
    Admin = 1,
    User = 2,
}

/// <summary>The four apps of the suite. The number is stored; the name is the address (<c>/docs</c>, <c>/sheets</c> ...).</summary>
public enum DocumentKind
{
    /// <summary>Text documents (like Word): MatOffice Docs.</summary>
    Docs = 1,

    /// <summary>Spreadsheets (like Excel): MatOffice Sheets.</summary>
    Sheets = 2,

    /// <summary>Presentations (like PowerPoint): MatOffice Slides.</summary>
    Slides = 3,

    /// <summary>Desktop databases (like Access): MatOffice Base.</summary>
    Base = 4,
}

/// <summary>What a person may do with a document they do not own. The owner may do everything.</summary>
public enum DocumentRole
{
    /// <summary>No access.</summary>
    None = 0,

    /// <summary>May open the document, read it, print and download it.</summary>
    Viewer = 1,

    /// <summary>May change the content and create versions.</summary>
    Editor = 2,

    /// <summary>May also rename, share and delete the document.</summary>
    Owner = 3,
}

/// <summary>What a link to a document allows for anybody who has it.</summary>
public enum LinkMode
{
    Off = 0,

    /// <summary>Anybody with the link may read the document (without an account).</summary>
    View = 1,

    /// <summary>Anybody who is signed in and has the link may edit it (it then shows up under "Shared with me").</summary>
    Edit = 2,
}

public sealed class User : SoftEntity
{
    /// <summary>The address to sign in with; unique, stored in lower case.</summary>
    public string Email { get; set; } = string.Empty;

    public string DisplayName { get; set; } = string.Empty;

    public string PasswordHash { get; set; } = string.Empty;

    public UserRole Role { get; set; } = UserRole.User;

    public bool IsActive { get; set; } = true;

    /// <summary>The IANA time zone dates are shown in.</summary>
    public string TimeZone { get; set; } = "Europe/Berlin";

    /// <summary>The language of the interface: en or de; empty = the language of the browser.</summary>
    public string Language { get; set; } = string.Empty;

    public int FailedLogins { get; set; }

    public DateTime? LockoutUntil { get; set; }

    public DateTime? LastLoginDate { get; set; }

    /// <summary>system, light or dark.</summary>
    public string Theme { get; set; } = "system";

    /// <summary>standard, paper, contrast or oled.</summary>
    public string Scheme { get; set; } = "standard";

    /// <summary>The paper size new documents start with: A4 or Letter.</summary>
    public string PaperSize { get; set; } = "A4";

    /// <summary>Mail about documents that are shared with the person (only when the installation can send mail).</summary>
    public bool MailOnShare { get; set; } = true;

    /// <summary>The person of MatIdentity (the claim <c>sub</c>) this account is linked with; null = a local account only.</summary>
    public string? ExternalSubject { get; set; }
}

/// <summary>A signed-in browser. The cookie carries the <see cref="Token"/>; the session lives in the database.</summary>
public sealed class UserSession : Entity
{
    public Guid Token { get; set; } = Guid.NewGuid();

    public long UserId { get; set; }

    public User? User { get; set; }

    public DateTime LastSeenDate { get; set; }

    public DateTime ExpiresDate { get; set; }

    public bool IsRevoked { get; set; }

    public string? UserAgent { get; set; }

    public string? IpAddress { get; set; }

    /// <summary>The ID token of a sign-in through MatIdentity (for the sign-out there); null for a local sign-in. A secret: never exported.</summary>
    public string? IdentityToken { get; set; }

    /// <summary>The session of MatIdentity (the claim <c>sid</c>) this session started with: the back-channel logout ends it by this.</summary>
    public string? IdentitySessionId { get; set; }
}

/// <summary>
/// A document of one of the four apps. <see cref="Content"/> is JSON in the shape of its kind (see <c>Services/Formats</c>); the editor in the browser reads and
/// writes it as a whole, the <see cref="Revision"/> keeps two browsers from overwriting each other. A deleted document lies in the trash for 30 days
/// (<see cref="SoftEntity.UpdateState"/> 0 and <see cref="DeletedDate"/>).
/// </summary>
public sealed class Document : SoftEntity
{
    public string Name { get; set; } = string.Empty;

    public DocumentKind Kind { get; set; }

    public long OwnerId { get; set; }

    public User? Owner { get; set; }

    public string Content { get; set; } = "{}";

    /// <summary>Counts every change: a browser that holds an older number knows its copy is old.</summary>
    public long Revision { get; set; }

    /// <summary>The size of <see cref="Content"/> in bytes (UTF-8), for the list and the limits.</summary>
    public long ContentLength { get; set; }

    /// <summary>A small picture of the first page / sheet / slide (PNG), made by the editor; shown in the gallery.</summary>
    public byte[]? Thumbnail { get; set; }

    /// <summary>The secret of the link (<c>/s/&lt;token&gt;</c>); changing it ends every link that was given out.</summary>
    public Guid ShareToken { get; set; } = Guid.NewGuid();

    public LinkMode LinkMode { get; set; } = LinkMode.Off;

    /// <summary>When it was moved to the trash; null while it is not.</summary>
    public DateTime? DeletedDate { get; set; }

    /// <summary>The last time the content changed (the name or the sharing do not count).</summary>
    public DateTime ContentDate { get; set; }

    public ICollection<DocumentMember> Members { get; set; } = [];
}

/// <summary>A person a document is shared with (the owner is not a member, they are <see cref="Document.OwnerId"/>).</summary>
public sealed class DocumentMember : Entity
{
    public long DocumentId { get; set; }

    public Document? Document { get; set; }

    public long UserId { get; set; }

    public User? User { get; set; }

    public DocumentRole Role { get; set; } = DocumentRole.Editor;
}

/// <summary>What one person did with a document: when they opened it last (the list "Recent") and whether they starred it.</summary>
public sealed class DocumentVisit : Entity
{
    public long DocumentId { get; set; }

    public Document? Document { get; set; }

    public long UserId { get; set; }

    public DateTime VisitDate { get; set; }

    public bool IsStarred { get; set; }
}

/// <summary>
/// A state of a document: a named version that somebody saved on purpose, or a snapshot the app keeps by itself while people work (at most one every
/// 10 minutes of changes). A version can be opened read-only and restored.
/// </summary>
public sealed class DocumentVersion : Entity
{
    public long DocumentId { get; set; }

    public Document? Document { get; set; }

    /// <summary>The name a person gave it; empty for a snapshot of the app.</summary>
    public string Name { get; set; } = string.Empty;

    public string? Description { get; set; }

    /// <summary>True for a snapshot the app made by itself (they are thinned out after 30 days).</summary>
    public bool IsAuto { get; set; }

    public long Revision { get; set; }

    public string Content { get; set; } = "{}";

    public long ContentLength { get; set; }
}

/// <summary>
/// An app that may open its files in MatOffice (a <em>host</em>, e.g. MatDrive): it signs its calls with an API key. Only a hash of the key is stored; the key is
/// shown once when it is made.
/// </summary>
public sealed class HostClient : SoftEntity
{
    public string Name { get; set; } = string.Empty;

    /// <summary>SHA-256 of the key (hex).</summary>
    public string KeyHash { get; set; } = string.Empty;

    /// <summary>The first characters of the key, to recognise it in the list.</summary>
    public string KeyPrefix { get; set; } = string.Empty;

    public bool IsActive { get; set; } = true;

    /// <summary>Optional: the start of the addresses the host may hand over for its files (<c>https://drive.example.com</c>); empty = any http(s) address.</summary>
    public string? AllowedUrlPrefix { get; set; }

    public DateTime? LastUsedDate { get; set; }

    public string? Note { get; set; }
}

/// <summary>
/// A file of a host that is open in MatOffice. The host hands over where to read and write the file and a token for it (WOPI-like); the browser only ever
/// sees <see cref="Token"/>. The working copy (<see cref="Content"/>) is the file converted into the shape of its kind; every save converts it back and writes
/// the file at the host.
/// </summary>
public sealed class HostSession : Entity
{
    /// <summary>The secret of the editor address <c>/h/&lt;token&gt;</c>.</summary>
    public Guid Token { get; set; } = Guid.NewGuid();

    public long HostClientId { get; set; }

    public HostClient? HostClient { get; set; }

    /// <summary>The id of the file at the host (for the log and to find a session again).</summary>
    public string FileId { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public DocumentKind Kind { get; set; }

    /// <summary>The extension the file is written back in (<c>.docx</c>, <c>.xlsx</c>, <c>.matsheet</c> ...).</summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>Where the file is read (GET) and written (PUT) at the host.</summary>
    public string ContentUrl { get; set; } = string.Empty;

    /// <summary>The token the host gave for the file (sent as <c>Authorization: Bearer</c>). A secret: never sent to a browser.</summary>
    public string AccessToken { get; set; } = string.Empty;

    public bool CanWrite { get; set; }

    /// <summary>The person at the host (their id, name and address), for the version history of the host and the presence in the editor.</summary>
    public string UserRef { get; set; } = string.Empty;

    public string UserName { get; set; } = string.Empty;

    public string? UserEmail { get; set; }

    /// <summary>Where "Back" leads (the folder of the file at the host).</summary>
    public string? ReturnUrl { get; set; }

    public DateTime ExpiresDate { get; set; }

    /// <summary>The working copy in the shape of its kind; empty until the file was read.</summary>
    public string? Content { get; set; }

    public long Revision { get; set; }

    /// <summary>The version of the file at the host when it was read or written last (its ETag): a save sends it as <c>If-Match</c>, so a file that was changed at the host meanwhile is not overwritten.</summary>
    public string? HostVersion { get; set; }

    public DateTime? LastSavedDate { get; set; }

    /// <summary>The last error writing the file back; null when the last save worked.</summary>
    public string? LastError { get; set; }
}

/// <summary>The few settings of the whole installation (one row, Id 1).</summary>
public sealed class SystemSetting : Entity
{
    /// <summary>Whether anyone may create an account. Closed until an administrator opens it (Administration → System): the first account of an installation can always
    /// sign up (it becomes the administrator), after it the administrators create the accounts.</summary>
    public bool RegistrationOpen { get; set; }

    /// <summary>The address the app is reached at from outside (for the links in mails); empty = the address of the request.</summary>
    public string? PublicUrl { get; set; }
}
