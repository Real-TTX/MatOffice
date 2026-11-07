using System.Text;
using MatOffice.Data;
using MatOffice.Localization;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Documents;

/// <summary>One row of the version history.</summary>
public sealed record VersionRow(long Id, string Name, string? Description, bool IsAuto, long Revision, string CreatedBy, DateTime CreateDate, long ContentLength);

/// <summary>
/// The version history of a document: snapshots the app keeps while people work (at most one per 10 minutes, thinned out after 30 days) and versions a
/// person names on purpose. Restoring keeps the state before it as a snapshot, so nothing is ever lost by a restore.
/// </summary>
public static class VersionService
{
    public const int NameMax = 120;

    /// <summary>The time between two snapshots of the app.</summary>
    public static readonly TimeSpan SnapshotEvery = TimeSpan.FromMinutes(10);

    public static async Task<IReadOnlyList<VersionRow>?> ListAsync(AppDbContext db, Actor actor, long documentId, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, actor, documentId, ct: ct) is null)
        {
            return null;
        }

        return await db.DocumentVersions.AsNoTracking()
            .Where(v => v.DocumentId == documentId)
            .OrderByDescending(v => v.CreateDate)
            .Select(v => new VersionRow(
                v.Id, v.Name, v.Description, v.IsAuto, v.Revision,
                db.Users.IgnoreQueryFilters().Where(u => u.Id == v.CreateUserId).Select(u => u.DisplayName).FirstOrDefault() ?? string.Empty,
                v.CreateDate, v.ContentLength))
            .ToListAsync(ct);
    }

    public static async Task<VersionRow?> FindAsync(AppDbContext db, Actor actor, long documentId, long versionId, CancellationToken ct = default)
        => (await ListAsync(db, actor, documentId, ct))?.FirstOrDefault(v => v.Id == versionId);

    /// <summary>Called after every save: a snapshot when the last one is older than <see cref="SnapshotEvery"/>.</summary>
    public static async Task SnapshotIfDueAsync(AppDbContext db, long userId, long documentId, long revision, string content, long length, CancellationToken ct = default)
    {
        var since = DateTime.UtcNow - SnapshotEvery;
        if (await db.DocumentVersions.AnyAsync(v => v.DocumentId == documentId && v.CreateDate > since, ct))
        {
            return;
        }

        db.DocumentVersions.Add(new DocumentVersion
        {
            DocumentId = documentId,
            IsAuto = true,
            Revision = revision,
            Content = content,
            ContentLength = length,
            CreateUserId = userId,
        });
        await db.SaveChangesAsync(ct);
    }

    /// <summary>A named version of the content now (editors).</summary>
    public static async Task<Outcome<DocumentVersion>> CreateAsync(AppDbContext db, long userId, long documentId, string name, string? description, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Editor, ct: ct);
        if (access is null)
        {
            return Outcome<DocumentVersion>.Denied();
        }

        var clean = DocumentService.CleanName(name);
        if (clean.Length == 0)
        {
            return Outcome<DocumentVersion>.Fail(Text.T("Versions.Name.Required"));
        }

        var version = new DocumentVersion
        {
            DocumentId = documentId,
            Name = clean.Length > NameMax ? clean[..NameMax] : clean,
            Description = Trim(description, 1000),
            Revision = access.Document.Revision,
            Content = access.Document.Content,
            ContentLength = access.Document.ContentLength,
        };
        db.DocumentVersions.Add(version);
        await db.SaveChangesAsync(ct);
        return Outcome<DocumentVersion>.Success(version);
    }

    /// <summary>Names a version (a snapshot of the app becomes a named version that is kept).</summary>
    public static async Task<Outcome<bool>> UpdateAsync(AppDbContext db, long userId, long documentId, long versionId, string name, string? description, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Editor, ct: ct) is null)
        {
            return Outcome<bool>.Denied();
        }

        var version = await db.DocumentVersions.FirstOrDefaultAsync(v => v.Id == versionId && v.DocumentId == documentId, ct);
        if (version is null)
        {
            return Outcome<bool>.Denied();
        }

        var clean = DocumentService.CleanName(name);
        if (clean.Length == 0)
        {
            return Outcome<bool>.Fail(Text.T("Versions.Name.Required"));
        }

        version.Name = clean.Length > NameMax ? clean[..NameMax] : clean;
        version.Description = Trim(description, 1000);
        version.IsAuto = false;
        await db.SaveChangesAsync(ct);
        return Outcome<bool>.Success(true);
    }

    public static async Task<bool> DeleteAsync(AppDbContext db, long userId, long documentId, long versionId, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Editor, ct: ct) is null)
        {
            return false;
        }

        return await db.DocumentVersions.Where(v => v.Id == versionId && v.DocumentId == documentId).ExecuteDeleteAsync(ct) > 0;
    }

    /// <summary>Makes a version the content of the document again; the content before it is kept as a snapshot.</summary>
    public static async Task<bool> RestoreAsync(AppDbContext db, long userId, long documentId, long versionId, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Editor, track: true, ct: ct);
        if (access is null)
        {
            return false;
        }

        var version = await db.DocumentVersions.AsNoTracking().FirstOrDefaultAsync(v => v.Id == versionId && v.DocumentId == documentId, ct);
        if (version is null)
        {
            return false;
        }

        var d = access.Document;
        db.DocumentVersions.Add(new DocumentVersion
        {
            DocumentId = d.Id,
            IsAuto = true,
            Name = string.Empty,
            Description = Text.T("Versions.BeforeRestore"),
            Revision = d.Revision,
            Content = d.Content,
            ContentLength = d.ContentLength,
        });
        d.Content = version.Content;
        d.ContentLength = Encoding.UTF8.GetByteCount(version.Content);
        d.Revision++;
        d.ContentDate = DateTime.UtcNow;
        await db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>The housekeeping: snapshots of the app older than 30 days go, except the last one of each day.</summary>
    public static async Task<int> ThinOutAsync(AppDbContext db, CancellationToken ct = default)
    {
        var limit = DateTime.UtcNow.AddDays(-30);
        var old = await db.DocumentVersions.AsNoTracking()
            .Where(v => v.IsAuto && v.CreateDate < limit)
            .Select(v => new { v.Id, v.DocumentId, v.CreateDate })
            .ToListAsync(ct);
        var keep = old.GroupBy(v => (v.DocumentId, v.CreateDate.Date)).Select(g => g.MaxBy(v => v.CreateDate)!.Id).ToHashSet();
        var remove = old.Where(v => !keep.Contains(v.Id)).Select(v => v.Id).ToList();
        return remove.Count == 0 ? 0 : await db.DocumentVersions.Where(v => remove.Contains(v.Id)).ExecuteDeleteAsync(ct);
    }

    private static string? Trim(string? text, int max)
    {
        var t = text?.Trim();
        return string.IsNullOrEmpty(t) ? null : t.Length > max ? t[..max] : t;
    }
}
