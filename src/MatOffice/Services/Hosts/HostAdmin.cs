using System.Security.Cryptography;
using System.Text;
using MatOffice.Data;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Hosts;

/// <summary>
/// The API keys of the hosts (apps such as MatDrive that open their files in MatOffice): the format (<c>mo_</c> + 40 random characters), the check of a key on
/// every call of the host API and the removal of a host with its open files. Only the hash of a key is kept; the key is shown once when it is made. Hosts are
/// made, renewed and removed as the bindings of <c>office.edit</c> (<see cref="HostBindings"/>).
/// </summary>
public static class HostAdmin
{
    public const string KeyPrefix = "mo_";

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

    /// <summary>Gives <paramref name="client"/> a new key (its hash and its first characters) and returns the key; the caller saves.</summary>
    internal static string IssueKey(HostClient client)
    {
        var key = NewKey();
        client.KeyHash = Hash(key);
        client.KeyPrefix = key[..10];
        return key;
    }

    /// <summary>Removes a host: its open files close (their sessions go) and the host is deleted (softly, its key stops working at once).</summary>
    internal static async Task RemoveAsync(AppDbContext db, HostClient client, CancellationToken ct)
    {
        await EndSessionsAsync(db, client.Id, ct);
        db.HostClients.Remove(client);
        await db.SaveChangesAsync(ct);
    }

    /// <summary>Ends the open files of a host (after a new key: the files the old key opened close too).</summary>
    internal static async Task EndSessionsAsync(AppDbContext db, long clientId, CancellationToken ct)
    {
        await db.HostSessions.Where(s => s.HostClientId == clientId).ExecuteDeleteAsync(ct);

        // the rows are gone: a session the context still tracks must not be deleted a second time (the cascade of a removed host)
        foreach (var entry in db.ChangeTracker.Entries<HostSession>().Where(e => e.Entity.HostClientId == clientId).ToList())
        {
            entry.State = EntityState.Detached;
        }
    }

    private static string NewKey() => KeyPrefix + Convert.ToBase64String(RandomNumberGenerator.GetBytes(30)).Replace('+', 'x').Replace('/', 'y').TrimEnd('=');
}
