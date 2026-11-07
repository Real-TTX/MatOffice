using System.Security.Cryptography;
using System.Text;
using MatOffice.Data;
using MatOffice.Localization;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Hosts;

public sealed record HostClientRow(long Id, string Name, string KeyPrefix, bool IsActive, string? AllowedUrlPrefix, DateTime? LastUsedDate, int OpenSessions, string? Note);

/// <summary>What the edit page of a host sends.</summary>
public sealed record HostClientInput(string Name, bool IsActive, string? AllowedUrlPrefix, string? Note);

/// <summary>
/// The hosts (apps such as MatDrive that open their files in MatOffice) and their API keys. Administrators only: every function checks it. The key is made here,
/// shown once and only its hash is kept.
/// </summary>
public static class HostAdmin
{
    public const string KeyPrefix = "mo_";

    public static async Task<IReadOnlyList<HostClientRow>?> ListAsync(AppDbContext db, long adminId, CancellationToken ct = default)
    {
        if (!await IsAdminAsync(db, adminId, ct))
        {
            return null;
        }

        var now = DateTime.UtcNow;
        return await db.HostClients.AsNoTracking().OrderBy(h => h.Name)
            .Select(h => new HostClientRow(h.Id, h.Name, h.KeyPrefix, h.IsActive, h.AllowedUrlPrefix, h.LastUsedDate,
                db.HostSessions.Count(s => s.HostClientId == h.Id && s.ExpiresDate > now), h.Note))
            .ToListAsync(ct);
    }

    public static async Task<HostClient?> FindAsync(AppDbContext db, long adminId, long id, CancellationToken ct = default)
        => await IsAdminAsync(db, adminId, ct) ? await db.HostClients.AsNoTracking().FirstOrDefaultAsync(h => h.Id == id, ct) : null;

    /// <summary>A new host with a new key; the key is in the answer (and nowhere else afterwards).</summary>
    public static async Task<Outcome<(HostClient Client, string Key)>> CreateAsync(AppDbContext db, long adminId, HostClientInput input, CancellationToken ct = default)
    {
        if (!await IsAdminAsync(db, adminId, ct))
        {
            return Outcome<(HostClient, string)>.Denied();
        }

        var error = Check(input);
        if (error is not null)
        {
            return Outcome<(HostClient, string)>.Fail(error);
        }

        var key = NewKey();
        var client = new HostClient
        {
            Name = input.Name.Trim(),
            IsActive = input.IsActive,
            AllowedUrlPrefix = Clean(input.AllowedUrlPrefix),
            Note = Clean(input.Note),
            KeyHash = Hash(key),
            KeyPrefix = key[..10],
        };
        db.HostClients.Add(client);
        await db.SaveChangesAsync(ct);
        return Outcome<(HostClient, string)>.Success((client, key));
    }

    public static async Task<Outcome<bool>> UpdateAsync(AppDbContext db, long adminId, long id, HostClientInput input, CancellationToken ct = default)
    {
        if (!await IsAdminAsync(db, adminId, ct))
        {
            return Outcome<bool>.Denied();
        }

        var client = await db.HostClients.FirstOrDefaultAsync(h => h.Id == id, ct);
        if (client is null)
        {
            return Outcome<bool>.Denied();
        }

        var error = Check(input);
        if (error is not null)
        {
            return Outcome<bool>.Fail(error);
        }

        client.Name = input.Name.Trim();
        client.IsActive = input.IsActive;
        client.AllowedUrlPrefix = Clean(input.AllowedUrlPrefix);
        client.Note = Clean(input.Note);
        await db.SaveChangesAsync(ct);
        return Outcome<bool>.Success(true);
    }

    /// <summary>A new key for a host: the old one stops working at once, and so do the files it opened.</summary>
    public static async Task<string?> RenewKeyAsync(AppDbContext db, long adminId, long id, CancellationToken ct = default)
    {
        if (!await IsAdminAsync(db, adminId, ct))
        {
            return null;
        }

        var client = await db.HostClients.FirstOrDefaultAsync(h => h.Id == id, ct);
        if (client is null)
        {
            return null;
        }

        var key = NewKey();
        client.KeyHash = Hash(key);
        client.KeyPrefix = key[..10];
        await db.SaveChangesAsync(ct);
        await db.HostSessions.Where(s => s.HostClientId == id).ExecuteDeleteAsync(ct);
        return key;
    }

    public static async Task<bool> DeleteAsync(AppDbContext db, long adminId, long id, CancellationToken ct = default)
    {
        if (!await IsAdminAsync(db, adminId, ct))
        {
            return false;
        }

        var client = await db.HostClients.FirstOrDefaultAsync(h => h.Id == id, ct);
        if (client is null)
        {
            return false;
        }

        await db.HostSessions.Where(s => s.HostClientId == id).ExecuteDeleteAsync(ct);
        db.HostClients.Remove(client);
        await db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>The active host a key belongs to (the API of the hosts), and notes when it was used.</summary>
    public static async Task<HostClient?> AuthenticateAsync(AppDbContext db, string? key, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(key) || !key.StartsWith(KeyPrefix, StringComparison.Ordinal) || key.Length > 200)
        {
            return null;
        }

        var hash = Hash(key);
        var client = await db.HostClients.FirstOrDefaultAsync(h => h.KeyHash == hash && h.IsActive, ct);
        if (client is not null && (client.LastUsedDate is null || client.LastUsedDate < DateTime.UtcNow.AddMinutes(-5)))
        {
            client.LastUsedDate = DateTime.UtcNow;
            await db.SaveChangesAsync(ct);
        }

        return client;
    }

    public static string Hash(string key) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(key)));

    private static string NewKey() => KeyPrefix + Convert.ToBase64String(RandomNumberGenerator.GetBytes(30)).Replace('+', 'x').Replace('/', 'y').TrimEnd('=');

    private static string? Check(HostClientInput input)
    {
        if (string.IsNullOrWhiteSpace(input.Name))
        {
            return Text.T("Hosts.Name.Required");
        }

        if (input.Name.Trim().Length > 120)
        {
            return Text.T("Hosts.Name.TooLong");
        }

        var prefix = Clean(input.AllowedUrlPrefix);
        if (prefix is not null && !(Uri.TryCreate(prefix, UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https"))
        {
            return Text.T("Hosts.Prefix.Invalid");
        }

        return null;
    }

    private static string? Clean(string? text)
    {
        var t = text?.Trim();
        return string.IsNullOrEmpty(t) ? null : t;
    }

    private static Task<bool> IsAdminAsync(AppDbContext db, long userId, CancellationToken ct)
        => db.Users.AnyAsync(u => u.Id == userId && u.IsActive && u.Role == UserRole.Admin, ct);
}
