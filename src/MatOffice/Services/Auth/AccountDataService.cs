using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using MatOffice.Data;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Auth;

/// <summary>
/// The data of a person under the GDPR: the export of everything that is about them (Art. 15 and 20, a JSON file), and the erasure of the account (Art. 17).
/// <para>
/// The export holds the account and its settings, the documents the person owns with everything in them (the content and every version, with the people the document is shared with),
/// the documents shared with them (name and role) and the sessions (and whether the account is linked with a person of MatIdentity). No secret: no password hash, no token of a session or
/// of a link, no ID token. The pictures of the list are left out (they are made from the content).
/// </para>
/// <para>
/// The erasure deletes for good (the account is not just hidden): the documents the person owns with everything in them, their memberships, visits and sessions. What they changed in a
/// document of somebody else stays there, without their name (the audit columns are cleared).
/// </para>
/// </summary>
public sealed class AccountDataService(AppDbContext db)
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>The export as JSON (UTF-8).</summary>
    public async Task<byte[]> ExportAsync(long userId, CancellationToken ct = default)
    {
        var user = await db.Users.AsNoTracking().FirstAsync(u => u.Id == userId, ct);
        var documents = await db.Documents.IgnoreQueryFilters().AsNoTracking().Where(d => d.OwnerId == userId).OrderBy(d => d.Id).ToListAsync(ct);
        var ids = documents.Select(d => d.Id).ToList();
        var versions = await db.DocumentVersions.AsNoTracking().Where(v => ids.Contains(v.DocumentId)).OrderBy(v => v.Id).ToListAsync(ct);
        var members = await db.DocumentMembers.AsNoTracking().Where(m => ids.Contains(m.DocumentId)).Select(m => new { m.DocumentId, m.User!.DisplayName, m.Role }).ToListAsync(ct);
        var shared = await db.DocumentMembers.AsNoTracking().Where(m => m.UserId == userId).Select(m => new { m.Document!.Name, m.Document.Kind, m.Role, m.CreateDate }).ToListAsync(ct);

        var export = new
        {
            Format = "MatOffice export 1",
            ExportedAt = DateTime.UtcNow,
            Account = new
            {
                user.Email,
                user.DisplayName,
                user.Role,
                user.TimeZone,
                user.Language,
                user.Theme,
                user.Scheme,
                user.PaperSize,
                user.MailOnShare,
                user.CreateDate,
                user.LastLoginDate,
                LinkedWithMatIdentity = user.ExternalSubject != null,
            },
            Documents = documents.Select(d => new
            {
                d.Id,
                d.Name,
                d.Kind,
                d.LinkMode,
                d.CreateDate,
                d.ContentDate,
                InTrash = d.UpdateState == SoftEntity.Deleted ? true : (bool?)null,
                Members = members.Where(m => m.DocumentId == d.Id).Select(m => new { Name = m.DisplayName, m.Role }).ToList(),
                Content = JsonNode.Parse(d.Content),
                Versions = versions.Where(v => v.DocumentId == d.Id).Select(v => new { v.Name, v.Description, v.IsAuto, v.CreateDate, Content = JsonNode.Parse(v.Content) }).ToList(),
            }).ToList(),
            SharedWithYou = shared.Select(s => new { s.Name, s.Kind, s.Role, Since = s.CreateDate }).ToList(),
            Starred = await db.DocumentVisits.AsNoTracking().Where(v => v.UserId == userId && v.IsStarred).Select(v => v.Document!.Name).ToListAsync(ct),
            Sessions = await db.Sessions.AsNoTracking().Where(s => s.UserId == userId)
                .Select(s => new { s.CreateDate, s.LastSeenDate, s.ExpiresDate, s.IsRevoked, s.UserAgent, s.IpAddress }).ToListAsync(ct),
        };

        return JsonSerializer.SerializeToUtf8Bytes(export, Json);
    }

    /// <summary>Erases the account for good (see the class). The caller checks the password and that it is not the last administrator.</summary>
    public async Task EraseAsync(long userId, CancellationToken ct = default)
    {
        await using var tx = await db.Database.BeginTransactionAsync(ct);

        // what the person changed in other people's documents stays, without their name
        await db.DocumentVersions.Where(v => v.CreateUserId == userId).ExecuteUpdateAsync(s => s.SetProperty(x => x.CreateUserId, (long?)null), ct);
        await db.DocumentVersions.Where(v => v.UpdateUserId == userId).ExecuteUpdateAsync(s => s.SetProperty(x => x.UpdateUserId, (long?)null), ct);
        await db.Documents.IgnoreQueryFilters().Where(d => d.UpdateUserId == userId && d.OwnerId != userId).ExecuteUpdateAsync(s => s.SetProperty(x => x.UpdateUserId, (long?)null), ct);
        await db.Documents.IgnoreQueryFilters().Where(d => d.CreateUserId == userId && d.OwnerId != userId).ExecuteUpdateAsync(s => s.SetProperty(x => x.CreateUserId, (long?)null), ct);
        await db.DocumentVisits.Where(v => v.UserId == userId).ExecuteDeleteAsync(ct);

        // the account: the database deletes its sessions, its documents (with their versions, members and visits) and its memberships with it
        await db.Users.IgnoreQueryFilters().Where(u => u.Id == userId).ExecuteDeleteAsync(ct);
        await tx.CommitAsync(ct);
    }

    /// <summary>Is the person the last active administrator (who must not go, or nobody can run the installation)?</summary>
    public async Task<bool> IsLastAdminAsync(long userId)
        => await db.Users.AnyAsync(u => u.Id == userId && u.Role == UserRole.Admin && u.IsActive)
           && !await db.Users.AnyAsync(u => u.Id != userId && u.Role == UserRole.Admin && u.IsActive);
}
