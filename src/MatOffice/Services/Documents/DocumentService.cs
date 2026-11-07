using System.Text;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Documents;

/// <summary>Which documents a list shows.</summary>
public enum DocumentScope
{
    /// <summary>Everything the person may see.</summary>
    All,

    Mine,

    Shared,

    Starred,

    /// <summary>The ones the person opened, the last first.</summary>
    Recent,

    /// <summary>The documents the person deleted (owner only), the last first.</summary>
    Trash,
}

public enum DocumentSort
{
    Modified,
    Opened,
    Name,
    Created,
}

/// <summary>What a list asks for.</summary>
public sealed record DocumentQuery(DocumentScope Scope, DocumentKind? Kind, string? Search, DocumentSort Sort, bool Descending, int Page, int PageSize);

/// <summary>One row of a list.</summary>
public sealed record DocumentRow(
    long Id,
    string Name,
    DocumentKind Kind,
    string OwnerName,
    bool IsOwner,
    DocumentRole Role,
    DateTime ContentDate,
    DateTime CreateDate,
    DateTime? VisitDate,
    bool IsStarred,
    bool HasThumbnail,
    bool IsShared,
    long ContentLength,
    DateTime? DeletedDate);

/// <summary>A page of a list and how many there are.</summary>
public sealed record DocumentPage(IReadOnlyList<DocumentRow> Items, int Total);

/// <summary>What the editor reads.</summary>
public sealed record EditorData(
    long Id,
    string Name,
    DocumentKind Kind,
    long Revision,
    string Content,
    DocumentRole Role,
    string OwnerName,
    long? VersionId,
    string? VersionName,
    LinkMode LinkMode,
    bool IsStarred)
{
    public bool CanEdit => Role >= DocumentRole.Editor && VersionId is null;
}

public enum SaveStatus
{
    Saved,
    Conflict,
    Invalid,
    Denied,
}

/// <summary>The answer of a save: the new revision, or the revision and content of the newer copy (conflict), or the problem.</summary>
public sealed record SaveResult(SaveStatus Status, long Revision = 0, string? Content = null, string? Problem = null);

/// <summary>A file to hand to the browser.</summary>
public sealed record FileDownload(byte[] Bytes, string FileName, string MimeType);

/// <summary>
/// The documents: lists, create (blank, template, import), open, save, rename, copy, star, trash, export, the picture of the list. Every function checks the
/// right of the person itself (<see cref="DocumentAccess"/>).
/// </summary>
public static class DocumentService
{
    public const int NameMax = 200;
    public const int ThumbnailMax = 512 * 1024;

    /// <summary>The time a deleted document stays in the trash.</summary>
    public static readonly TimeSpan TrashKeep = TimeSpan.FromDays(30);

    // ---- lists -------------------------------------------------------------------------------------------------------------------

    public static async Task<DocumentPage> ListAsync(AppDbContext db, long userId, DocumentQuery q, CancellationToken ct = default)
    {
        IQueryable<Document> docs = q.Scope == DocumentScope.Trash
            ? db.Documents.IgnoreQueryFilters().Where(d => d.OwnerId == userId && d.UpdateState == SoftEntity.Deleted && d.DeletedDate != null)
            : DocumentAccess.Visible(db, userId);

        if (q.Kind is { } kind)
        {
            docs = docs.Where(d => d.Kind == kind);
        }

        if (!string.IsNullOrWhiteSpace(q.Search))
        {
            var pattern = "%" + q.Search.Trim().Replace("\\", "\\\\").Replace("%", "\\%").Replace("_", "\\_") + "%";
            docs = docs.Where(d => EF.Functions.ILike(d.Name, pattern));
        }

        var visits = db.DocumentVisits.Where(v => v.UserId == userId);
        switch (q.Scope)
        {
            case DocumentScope.Mine:
                docs = docs.Where(d => d.OwnerId == userId);
                break;
            case DocumentScope.Shared:
                docs = docs.Where(d => d.OwnerId != userId);
                break;
            case DocumentScope.Starred:
                docs = docs.Where(d => visits.Any(v => v.DocumentId == d.Id && v.IsStarred));
                break;
            case DocumentScope.Recent:
                docs = docs.Where(d => visits.Any(v => v.DocumentId == d.Id));
                break;
        }

        var rows = docs.Select(d => new
        {
            d.Id,
            d.Name,
            d.Kind,
            OwnerName = d.Owner!.DisplayName,
            d.OwnerId,
            MemberRole = d.Members.Where(m => m.UserId == userId).Select(m => (DocumentRole?)m.Role).FirstOrDefault(),
            d.ContentDate,
            d.CreateDate,
            Visit = visits.Where(v => v.DocumentId == d.Id).Select(v => (DateTime?)v.VisitDate).FirstOrDefault(),
            Starred = visits.Any(v => v.DocumentId == d.Id && v.IsStarred),
            HasThumbnail = d.Thumbnail != null,
            IsShared = d.Members.Any() || d.LinkMode != LinkMode.Off,
            d.ContentLength,
            d.DeletedDate,
        });

        var total = await rows.CountAsync(ct);
        var scope = q.Scope;
        rows = (q.Sort, q.Descending) switch
        {
            _ when scope == DocumentScope.Trash => rows.OrderByDescending(r => r.DeletedDate).ThenBy(r => r.Id),
            _ when scope == DocumentScope.Recent && q.Sort == DocumentSort.Modified => rows.OrderByDescending(r => r.Visit).ThenBy(r => r.Id),
            (DocumentSort.Name, false) => rows.OrderBy(r => r.Name).ThenBy(r => r.Id),
            (DocumentSort.Name, true) => rows.OrderByDescending(r => r.Name).ThenBy(r => r.Id),
            (DocumentSort.Opened, false) => rows.OrderBy(r => r.Visit).ThenBy(r => r.Id),
            (DocumentSort.Opened, true) => rows.OrderByDescending(r => r.Visit).ThenBy(r => r.Id),
            (DocumentSort.Created, false) => rows.OrderBy(r => r.CreateDate).ThenBy(r => r.Id),
            (DocumentSort.Created, true) => rows.OrderByDescending(r => r.CreateDate).ThenBy(r => r.Id),
            (_, false) => rows.OrderBy(r => r.ContentDate).ThenBy(r => r.Id),
            _ => rows.OrderByDescending(r => r.ContentDate).ThenBy(r => r.Id),
        };

        var page = await rows.Skip(Math.Max(0, q.Page - 1) * q.PageSize).Take(q.PageSize).ToListAsync(ct);
        var items = page.Select(r => new DocumentRow(
            r.Id, r.Name, r.Kind, r.OwnerName, r.OwnerId == userId,
            r.OwnerId == userId ? DocumentRole.Owner : r.MemberRole ?? DocumentRole.Viewer,
            r.ContentDate, r.CreateDate, r.Visit, r.Starred, r.HasThumbnail, r.IsShared, r.ContentLength, r.DeletedDate)).ToList();
        return new DocumentPage(items, total);
    }

    /// <summary>The numbers of the menu.</summary>
    public static async Task<(int Mine, int Shared, int Starred, int Trash)> CountsAsync(AppDbContext db, long userId, CancellationToken ct = default)
    {
        var mine = await db.Documents.CountAsync(d => d.OwnerId == userId, ct);
        var shared = await db.Documents.CountAsync(d => d.OwnerId != userId && d.Members.Any(m => m.UserId == userId), ct);
        var starred = await DocumentAccess.Visible(db, userId).CountAsync(d => db.DocumentVisits.Any(v => v.UserId == userId && v.DocumentId == d.Id && v.IsStarred), ct);
        var trash = await db.Documents.IgnoreQueryFilters().CountAsync(d => d.OwnerId == userId && d.UpdateState == SoftEntity.Deleted && d.DeletedDate != null, ct);
        return (mine, shared, starred, trash);
    }

    /// <summary>The documents the person opened last (the menu).</summary>
    public static async Task<IReadOnlyList<(long Id, string Name, DocumentKind Kind)>> RecentAsync(AppDbContext db, long userId, int count, CancellationToken ct = default)
    {
        var rows = await DocumentAccess.Visible(db, userId)
            .Join(db.DocumentVisits.Where(v => v.UserId == userId), d => d.Id, v => v.DocumentId, (d, v) => new { d.Id, d.Name, d.Kind, v.VisitDate })
            .OrderByDescending(r => r.VisitDate)
            .Take(count)
            .ToListAsync(ct);
        return rows.Select(r => (r.Id, r.Name, r.Kind)).ToList();
    }

    // ---- create ------------------------------------------------------------------------------------------------------------------

    /// <summary>A new document of a kind from a template (the blank one when the id is unknown).</summary>
    public static async Task<Outcome<Document>> CreateAsync(AppDbContext db, long userId, DocumentKind kind, string? templateId, string? name, string language, CancellationToken ct = default)
    {
        var handler = Kinds.Of(kind);
        var template = handler.Templates.FirstOrDefault(t => t.Id == templateId) ?? handler.Templates[0];
        var checkedContent = handler.Normalize(handler.Create(template.Id, language));
        if (!checkedContent.Ok)
        {
            return Outcome<Document>.Fail(checkedContent.Error!);
        }

        var title = CleanName(name);
        if (title.Length == 0)
        {
            title = template.Id == "blank" ? Text.T("Documents.Untitled." + kind) : Text.T(template.Name);
        }

        return Outcome<Document>.Success(await AddAsync(db, userId, kind, title, checkedContent.Json!, ct));
    }

    /// <summary>A document from a file (native or foreign); the kind follows from the extension.</summary>
    public static async Task<Outcome<Document>> ImportAsync(AppDbContext db, long userId, string fileName, Stream input, CancellationToken ct = default)
    {
        var ext = Path.GetExtension(fileName);
        if (Kinds.ForExtension(ext) is not { } found)
        {
            return Outcome<Document>.Fail(Text.T("Documents.Import.UnknownType", ext));
        }

        var result = Kinds.Read(found.Handler, input, ext);
        if (!result.Ok)
        {
            return Outcome<Document>.Fail(result.Error!);
        }

        var name = CleanName(Path.GetFileNameWithoutExtension(fileName));
        return Outcome<Document>.Success(await AddAsync(db, userId, found.Handler.Kind, name.Length == 0 ? Text.T("Documents.Untitled." + found.Handler.Kind) : name, result.Json!, ct));
    }

    private static async Task<Document> AddAsync(AppDbContext db, long userId, DocumentKind kind, string name, string content, CancellationToken ct)
    {
        var now = DateTime.UtcNow;
        var document = new Document
        {
            Name = name,
            Kind = kind,
            OwnerId = userId,
            Content = content,
            ContentLength = Encoding.UTF8.GetByteCount(content),
            Revision = 1,
            ContentDate = now,
        };
        db.Documents.Add(document);
        await db.SaveChangesAsync(ct);
        await TouchAsync(db, userId, document.Id, ct);
        return document;
    }

    public static string CleanName(string? name)
    {
        var text = new string((name ?? string.Empty).Where(c => !char.IsControl(c)).ToArray()).Trim();
        return text.Length > NameMax ? text[..NameMax] : text;
    }

    // ---- open and save -----------------------------------------------------------------------------------------------------------

    /// <summary>The document as the editor reads it (or a version of it, read-only). Records the visit of a signed-in person.</summary>
    public static async Task<EditorData?> OpenAsync(AppDbContext db, Actor actor, long id, long? versionId = null, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, actor, id, ct: ct);
        if (access is null)
        {
            return null;
        }

        var d = access.Document;
        var owner = await db.Users.IgnoreQueryFilters().Where(u => u.Id == d.OwnerId).Select(u => u.DisplayName).FirstOrDefaultAsync(ct) ?? string.Empty;
        var starred = false;
        if (actor.UserId is { } userId)
        {
            starred = await db.DocumentVisits.AnyAsync(v => v.UserId == userId && v.DocumentId == id && v.IsStarred, ct);
            if (versionId is null && (access.IsOwner || await db.DocumentMembers.AnyAsync(m => m.DocumentId == id && m.UserId == userId, ct)))
            {
                await TouchAsync(db, userId, id, ct);
            }
        }

        if (versionId is { } vid)
        {
            var version = await db.DocumentVersions.AsNoTracking().FirstOrDefaultAsync(v => v.Id == vid && v.DocumentId == id, ct);
            if (version is null)
            {
                return null;
            }

            var label = version.Name.Length > 0 ? version.Name : Text.T("Versions.Auto");
            return new EditorData(d.Id, d.Name, d.Kind, version.Revision, version.Content, DocumentRole.Viewer, owner, version.Id, label, d.LinkMode, starred);
        }

        return new EditorData(d.Id, d.Name, d.Kind, d.Revision, d.Content, access.Role, owner, null, null, d.LinkMode, starred);
    }

    /// <summary>The revision of the document now (the editor asks this to notice the changes of others) and who made it.</summary>
    public static async Task<(long Revision, string? By)?> RevisionAsync(AppDbContext db, Actor actor, long id, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, actor, id, ct: ct);
        if (access is null)
        {
            return null;
        }

        var by = access.Document.UpdateUserId is { } u ? await db.Users.IgnoreQueryFilters().Where(x => x.Id == u).Select(x => x.DisplayName).FirstOrDefaultAsync(ct) : null;
        return (access.Document.Revision, by);
    }

    /// <summary>
    /// Saves the content when <paramref name="revision"/> is the revision of the document now: the UPDATE itself compares it, so two browsers never overwrite each other.
    /// Keeps a snapshot of the app at most every 10 minutes (the version history).
    /// </summary>
    public static async Task<SaveResult> SaveAsync(AppDbContext db, long userId, long id, long revision, string content, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), id, DocumentRole.Editor, ct: ct);
        if (access is null)
        {
            return new SaveResult(SaveStatus.Denied);
        }

        var check = Kinds.Of(access.Document.Kind).Normalize(content);
        if (!check.Ok)
        {
            return new SaveResult(SaveStatus.Invalid, Problem: check.Error);
        }

        var json = check.Json!;
        var now = DateTime.UtcNow;
        var length = Encoding.UTF8.GetByteCount(json);
        var changed = await db.Documents
            .Where(d => d.Id == id && d.Revision == revision)
            .ExecuteUpdateAsync(s => s
                .SetProperty(d => d.Content, json)
                .SetProperty(d => d.ContentLength, length)
                .SetProperty(d => d.Revision, revision + 1)
                .SetProperty(d => d.ContentDate, now)
                .SetProperty(d => d.UpdateDate, now)
                .SetProperty(d => d.UpdateUserId, userId)
                .SetProperty(d => d.UpdateState, SoftEntity.Updated), ct);
        if (changed == 0)
        {
            var current = await db.Documents.AsNoTracking().Where(d => d.Id == id).Select(d => new { d.Revision, d.Content }).FirstOrDefaultAsync(ct);
            return current is null ? new SaveResult(SaveStatus.Denied) : new SaveResult(SaveStatus.Conflict, current.Revision, current.Content);
        }

        await VersionService.SnapshotIfDueAsync(db, userId, id, revision + 1, json, length, ct);
        return new SaveResult(SaveStatus.Saved, revision + 1);
    }

    // ---- change ------------------------------------------------------------------------------------------------------------------

    public static async Task<Outcome<string>> RenameAsync(AppDbContext db, long userId, long id, string name, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), id, DocumentRole.Editor, track: true, ct: ct);
        if (access is null)
        {
            return Outcome<string>.Denied();
        }

        var clean = CleanName(name);
        if (clean.Length == 0)
        {
            return Outcome<string>.Fail(Text.T("Documents.Name.Required"));
        }

        access.Document.Name = clean;
        await db.SaveChangesAsync(ct);
        return Outcome<string>.Success(clean);
    }

    /// <summary>A copy that belongs to the person (anybody who may read a document may copy it).</summary>
    public static async Task<Outcome<Document>> DuplicateAsync(AppDbContext db, long userId, long id, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), id, ct: ct);
        if (access is null)
        {
            return Outcome<Document>.Denied();
        }

        var d = access.Document;
        var name = CleanName(Text.T("Documents.CopyOf", d.Name));
        var copy = await AddAsync(db, userId, d.Kind, name, d.Content, ct);
        if (d.Thumbnail is { } thumb)
        {
            await db.Documents.Where(x => x.Id == copy.Id).ExecuteUpdateAsync(s => s.SetProperty(x => x.Thumbnail, thumb), ct);
        }

        return Outcome<Document>.Success(copy);
    }

    /// <summary>Puts a document into the trash (owner only); it stays there for 30 days.</summary>
    public static async Task<bool> TrashAsync(AppDbContext db, long userId, long id, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), id, DocumentRole.Owner, track: true, ct: ct);
        if (access is null)
        {
            return false;
        }

        access.Document.DeletedDate = DateTime.UtcNow;
        db.Documents.Remove(access.Document);
        await db.SaveChangesAsync(ct);
        return true;
    }

    public static async Task<bool> RestoreAsync(AppDbContext db, long userId, long id, CancellationToken ct = default)
    {
        var document = await DocumentAccess.FindTrashedAsync(db, userId, id, ct);
        if (document is null)
        {
            return false;
        }

        document.DeletedDate = null;
        document.UpdateState = SoftEntity.Updated;
        await db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>Removes a document of the trash for good.</summary>
    public static async Task<bool> DeleteForeverAsync(AppDbContext db, long userId, long id, CancellationToken ct = default)
    {
        var document = await DocumentAccess.FindTrashedAsync(db, userId, id, ct);
        if (document is null)
        {
            return false;
        }

        await db.Documents.IgnoreQueryFilters().Where(d => d.Id == document.Id).ExecuteDeleteAsync(ct);
        return true;
    }

    public static Task<int> EmptyTrashAsync(AppDbContext db, long userId, CancellationToken ct = default)
        => db.Documents.IgnoreQueryFilters().Where(d => d.OwnerId == userId && d.UpdateState == SoftEntity.Deleted && d.DeletedDate != null).ExecuteDeleteAsync(ct);

    /// <summary>Stars or unstars a document for the person (anybody who may see it).</summary>
    public static async Task<bool> StarAsync(AppDbContext db, long userId, long id, bool starred, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), id, ct: ct) is null)
        {
            return false;
        }

        var visit = await db.DocumentVisits.FirstOrDefaultAsync(v => v.UserId == userId && v.DocumentId == id, ct);
        if (visit is null)
        {
            db.DocumentVisits.Add(new DocumentVisit { UserId = userId, DocumentId = id, VisitDate = DateTime.UtcNow, IsStarred = starred });
        }
        else
        {
            visit.IsStarred = starred;
        }

        await db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>Notes that the person opened the document now (the list "Recent").</summary>
    public static async Task TouchAsync(AppDbContext db, long userId, long id, CancellationToken ct = default)
    {
        var now = DateTime.UtcNow;
        var updated = await db.DocumentVisits.Where(v => v.UserId == userId && v.DocumentId == id)
            .ExecuteUpdateAsync(s => s.SetProperty(v => v.VisitDate, now), ct);
        if (updated == 0)
        {
            db.DocumentVisits.Add(new DocumentVisit { UserId = userId, DocumentId = id, VisitDate = now });
            try
            {
                await db.SaveChangesAsync(ct);
            }
            catch (DbUpdateException)
            {
                // two tabs opened it at the same moment: the other one wrote the visit
            }
        }
    }

    // ---- files -------------------------------------------------------------------------------------------------------------------

    /// <summary>The document (or a version of it) as a file of one of the formats of its kind.</summary>
    public static async Task<Outcome<FileDownload>> ExportAsync(AppDbContext db, Actor actor, long id, string extension, long? versionId = null, CancellationToken ct = default)
    {
        var data = await OpenAsync(db, actor, id, versionId, ct);
        if (data is null)
        {
            return Outcome<FileDownload>.Denied();
        }

        return Export(data.Kind, data.Content, data.Name, extension);
    }

    /// <summary>Content as a file (also used for the files of a host).</summary>
    public static Outcome<FileDownload> Export(DocumentKind kind, string content, string name, string extension)
    {
        var handler = Kinds.Of(kind);
        var ext = Kinds.NormalizeExtension(extension);
        if (!Kinds.ExportFormats(handler).Any(f => f.Extension == ext))
        {
            return Outcome<FileDownload>.Fail(Text.T("Documents.Export.NotSupported"));
        }

        try
        {
            var bytes = Kinds.Write(handler, content, ext, name);
            return Outcome<FileDownload>.Success(new FileDownload(bytes, SafeFileName(name) + ext, Kinds.MimeOf(handler, ext)));
        }
        catch (Formats.FormatException ex)
        {
            return Outcome<FileDownload>.Fail(ex.Message);
        }
    }

    /// <summary>A file name without characters that file systems refuse.</summary>
    public static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars().Concat(['/', '\\', ':', '*', '?', '"', '<', '>', '|']).ToHashSet();
        var clean = new string(name.Select(c => invalid.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
        return clean.Length == 0 ? "document" : clean;
    }

    public static async Task<byte[]?> ThumbnailAsync(AppDbContext db, Actor actor, long id, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, actor, id, ct: ct);
        return access?.Document.Thumbnail;
    }

    /// <summary>Stores the picture of the list (a PNG made by the editor; editors only).</summary>
    public static async Task<bool> SetThumbnailAsync(AppDbContext db, long userId, long id, byte[] png, CancellationToken ct = default)
    {
        if (png.Length is 0 or > ThumbnailMax || !IsPng(png))
        {
            return false;
        }

        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), id, DocumentRole.Editor, ct: ct) is null)
        {
            return false;
        }

        await db.Documents.Where(d => d.Id == id).ExecuteUpdateAsync(s => s.SetProperty(d => d.Thumbnail, png), ct);
        return true;
    }

    private static bool IsPng(byte[] data) => data.Length > 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47;
}
