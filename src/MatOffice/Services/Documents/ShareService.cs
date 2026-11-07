using MatOffice.Data;
using MatOffice.Localization;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Documents;

public sealed record MemberRow(long UserId, string Name, string Email, DocumentRole Role);

/// <summary>Sharing: the people a document is shared with and its link. Only the owner changes them; members may see who else has access.</summary>
public static class ShareService
{
    public static async Task<IReadOnlyList<MemberRow>?> MembersAsync(AppDbContext db, long userId, long documentId, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, ct: ct) is null)
        {
            return null;
        }

        var rows = await db.DocumentMembers.AsNoTracking().Where(m => m.DocumentId == documentId)
            .Select(m => new { m.UserId, m.User!.DisplayName, m.User.Email, m.Role }).ToListAsync(ct);
        return rows.OrderBy(r => r.DisplayName, StringComparer.CurrentCultureIgnoreCase).Select(r => new MemberRow(r.UserId, r.DisplayName, r.Email, r.Role)).ToList();
    }

    /// <summary>Shares a document with a person who has an account (by their address).</summary>
    public static async Task<Outcome<MemberRow>> AddMemberAsync(AppDbContext db, long userId, long documentId, string email, DocumentRole role, CancellationToken ct = default)
    {
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Owner, ct: ct) is not { } access)
        {
            return Outcome<MemberRow>.Denied();
        }

        if (role is not (DocumentRole.Viewer or DocumentRole.Editor))
        {
            return Outcome<MemberRow>.Fail(Text.T("Members.RoleInvalid"));
        }

        var address = AccountService.Normalize(email);
        var person = await db.Users.AsNoTracking().Where(u => u.Email == address && u.IsActive).Select(u => new { u.Id, u.DisplayName, u.Email }).FirstOrDefaultAsync(ct);
        if (person is null)
        {
            return Outcome<MemberRow>.Fail(Text.T("Members.NoAccount"));
        }

        if (person.Id == access.Document.OwnerId)
        {
            return Outcome<MemberRow>.Fail(Text.T("Members.IsOwner"));
        }

        if (await db.DocumentMembers.AnyAsync(m => m.DocumentId == documentId && m.UserId == person.Id, ct))
        {
            return Outcome<MemberRow>.Fail(Text.T("Members.AlreadyMember"));
        }

        db.DocumentMembers.Add(new DocumentMember { DocumentId = documentId, UserId = person.Id, Role = role, CreateUserId = userId });
        await db.SaveChangesAsync(ct);
        return Outcome<MemberRow>.Success(new MemberRow(person.Id, person.DisplayName, person.Email, role));
    }

    public static async Task<bool> SetMemberRoleAsync(AppDbContext db, long userId, long documentId, long memberUserId, DocumentRole role, CancellationToken ct = default)
    {
        if (role is not (DocumentRole.Viewer or DocumentRole.Editor) || await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Owner, ct: ct) is null)
        {
            return false;
        }

        var member = await db.DocumentMembers.FirstOrDefaultAsync(m => m.DocumentId == documentId && m.UserId == memberUserId, ct);
        if (member is null)
        {
            return false;
        }

        member.Role = role;
        await db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>Ends the sharing with a person (the owner), or a member leaves the document themselves.</summary>
    public static async Task<bool> RemoveMemberAsync(AppDbContext db, long userId, long documentId, long memberUserId, CancellationToken ct = default)
    {
        var minimum = memberUserId == userId ? DocumentRole.Viewer : DocumentRole.Owner;
        if (await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, minimum, ct: ct) is null)
        {
            return false;
        }

        return await db.DocumentMembers.Where(m => m.DocumentId == documentId && m.UserId == memberUserId).ExecuteDeleteAsync(ct) > 0;
    }

    /// <summary>Sets what the link allows; <paramref name="renew"/> makes a new secret, so every link given out before stops working.</summary>
    public static async Task<Document?> SetLinkAsync(AppDbContext db, long userId, long documentId, LinkMode mode, bool renew, CancellationToken ct = default)
    {
        var access = await DocumentAccess.FindAsync(db, Actor.Of(userId), documentId, DocumentRole.Owner, track: true, ct: ct);
        if (access is null)
        {
            return null;
        }

        access.Document.LinkMode = mode;
        if (renew)
        {
            access.Document.ShareToken = Guid.NewGuid();
        }

        await db.SaveChangesAsync(ct);
        return access.Document;
    }

    /// <summary>Whether a person wants a mail when something is shared with them.</summary>
    public static async Task<bool> WantsShareMailAsync(AppDbContext db, long userId, CancellationToken ct = default)
        => await db.Users.AsNoTracking().Where(u => u.Id == userId).Select(u => u.MailOnShare).FirstOrDefaultAsync(ct);

    /// <summary>The language of a person (for a mail in their words); empty = the default.</summary>
    public static async Task<string> LanguageOfAsync(AppDbContext db, long userId, CancellationToken ct = default)
        => await db.Users.AsNoTracking().Where(u => u.Id == userId).Select(u => u.Language).FirstOrDefaultAsync(ct) ?? string.Empty;
}
