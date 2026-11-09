using System.Globalization;
using MatOffice.Data;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Hosts;

/// <summary>A host as a binding of <c>office.edit</c>: the consumer (its id), its name, who manages it, since when and when its key was last used.</summary>
public sealed record HostBinding(long Id, string Consumer, string Name, string ManagedBy, DateTime Since, DateTime? LastUsed);

/// <summary>The host of a binding and its key when one was made now; null = the host has its key already, and it is not handed out again.</summary>
public sealed record HostBound(HostClient Client, string? Key, bool Created);

/// <summary>
/// The hosts as the bindings of the capability <c>office.edit</c> (MatPlatform 2.0, docs/PLATFORM.md "Bindings"): a broker such as Mattiverse, or an administrator in the
/// Connections area, binds a consumer (MatDrive) and gets the key it calls the host API with. No check of a person here: the broker is not a person. Who may call this
/// is decided by MatPlatform (the key <c>MAT_API_KEY</c> for the broker, <c>ConnectionsAccess</c> = the administrators for the Connections area).
/// A host made by hand before the bindings has no consumer id: it is listed as <c>host-&lt;Id&gt;</c>, so it stays visible and can be renewed and removed.
/// </summary>
public static class HostBindings
{
    /// <summary>Who made a binding by hand (the Connections area, or a host from before the bindings).</summary>
    public const string ManagedByHand = "manual";

    /// <summary>The longest consumer id and name of a manager (the columns).</summary>
    public const int ConsumerMax = 64;
    public const int ManagedByMax = 40;

    private const string LegacyPrefix = "host-";
    private const int NameMax = 120;

    /// <summary>Every host (not deleted), by name.</summary>
    public static async Task<IReadOnlyList<HostBinding>> ListAsync(AppDbContext db, CancellationToken ct = default)
    {
        var rows = await db.HostClients.AsNoTracking().OrderBy(h => h.Name).ThenBy(h => h.Id)
            .Select(h => new { h.Id, h.ConsumerId, h.Name, h.ManagedBy, h.CreateDate, h.LastUsedDate })
            .ToListAsync(ct);
        return [.. rows.Select(h => new HostBinding(h.Id, h.ConsumerId ?? LegacyPrefix + h.Id.ToString(CultureInfo.InvariantCulture), h.Name, h.ManagedBy ?? ManagedByHand, h.CreateDate, h.LastUsedDate))];
    }

    /// <summary>
    /// Makes or updates the binding of <paramref name="consumer"/> (the same call twice is one host): a new host when there is none, with a new key; an existing one keeps
    /// its key unless <paramref name="renew"/> (then the files the old key opened close). A bound host is active.
    /// </summary>
    public static async Task<HostBound> BindAsync(AppDbContext db, string consumer, string? name, string managedBy, bool renew, CancellationToken ct = default)
    {
        Check(consumer, managedBy);
        var client = await FindAsync(db, consumer, ct);
        var created = client is null;
        if (client is null)
        {
            client = new HostClient { ConsumerId = consumer, AllowedUrlPrefix = null };
            db.HostClients.Add(client);
        }

        var given = name?.Trim();
        client.Name = string.IsNullOrEmpty(given) ? (created ? consumer : client.Name) : given.Length > NameMax ? given[..NameMax] : given;
        client.ManagedBy = managedBy;
        client.IsActive = true;
        var key = created || renew ? HostAdmin.IssueKey(client) : null;
        await db.SaveChangesAsync(ct);
        if (renew && !created)
        {
            await HostAdmin.EndSessionsAsync(db, client.Id, ct);
        }

        return new HostBound(client, key, created);
    }

    /// <summary>Removes the binding of <paramref name="consumer"/> with its open files: only one made by <paramref name="managedBy"/>, any when it is null. False when there was none.</summary>
    public static async Task<bool> UnbindAsync(AppDbContext db, string consumer, string? managedBy, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(consumer))
        {
            return false;
        }

        var client = await FindAsync(db, consumer, ct);
        if (client is null || (managedBy is not null && (client.ManagedBy ?? ManagedByHand) != managedBy))
        {
            return false;
        }

        await HostAdmin.RemoveAsync(db, client, ct);
        return true;
    }

    /// <summary>The host of a consumer id, or of <c>host-&lt;Id&gt;</c> for a host from before the bindings.</summary>
    private static async Task<HostClient?> FindAsync(AppDbContext db, string consumer, CancellationToken ct)
    {
        var client = await db.HostClients.FirstOrDefaultAsync(h => h.ConsumerId == consumer, ct);
        if (client is null && consumer.StartsWith(LegacyPrefix, StringComparison.Ordinal)
            && long.TryParse(consumer.AsSpan(LegacyPrefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out var id)
            && consumer == LegacyPrefix + id.ToString(CultureInfo.InvariantCulture))
        {
            client = await db.HostClients.FirstOrDefaultAsync(h => h.Id == id && h.ConsumerId == null, ct);
        }

        return client;
    }

    private static void Check(string consumer, string managedBy)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(consumer);
        ArgumentException.ThrowIfNullOrWhiteSpace(managedBy);
        if (consumer.Length > ConsumerMax)
        {
            throw new ArgumentException($"A consumer id has at most {ConsumerMax} characters.", nameof(consumer));
        }

        if (managedBy.Length > ManagedByMax)
        {
            throw new ArgumentException($"The name of a manager has at most {ManagedByMax} characters.", nameof(managedBy));
        }
    }
}
