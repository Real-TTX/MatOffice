using MatOffice.Data;
using MatOffice.Services.Hosts;
using MatPlatform;

namespace MatOffice.Platform;

/// <summary>
/// The bindings of <c>office.edit</c> for MatPlatform (<c>/api/bindings</c> of a broker and the Connections area): a binding is a host of MatOffice, a row of
/// <see cref="HostClient"/> with its key and its open files, not the generic key store of MatPlatform. The work is done by <see cref="HostBindings"/>; this class
/// only translates. The values are the addresses of MatOffice as the broker names them (or as the request reached it) and the key, once.
/// </summary>
public sealed class MatOfficeHostBindings(AppDbContext db) : IMatBindingHandler
{
    public string Capability => MatCapabilities.OfficeEdit;

    public async Task<IReadOnlyList<MatBindingInfo>> ListAsync(CancellationToken ct)
        => [.. (await HostBindings.ListAsync(db, ct)).Select(h => new MatBindingInfo(Capability, h.Consumer, h.Name, h.ManagedBy, Utc(h.Since), h.LastUsed is { } used ? Utc(used) : null))];

    public async Task<MatBindResult> BindAsync(MatBindRequest request, CancellationToken ct)
    {
        ArgumentNullException.ThrowIfNull(request);
        var bound = await HostBindings.BindAsync(db, request.Consumer.Id, request.Consumer.Name, request.ManagedBy, request.Renew, ct);
        return new MatBindResult(new Dictionary<string, string?>
        {
            [MatCapabilities.Url] = request.Provider.Url,
            [MatCapabilities.InternalUrl] = request.Provider.InternalUrl,
            [MatCapabilities.Key] = bound.Key,
        });
    }

    public Task<bool> UnbindAsync(string consumer, string? managedBy, CancellationToken ct)
        => HostBindings.UnbindAsync(db, consumer, managedBy, ct);

    private static DateTimeOffset Utc(DateTime time) => new(DateTime.SpecifyKind(time, DateTimeKind.Utc));
}
