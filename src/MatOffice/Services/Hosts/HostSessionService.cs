using System.Net;
using System.Net.Http.Headers;
using System.Text;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Services.Hosts;

/// <summary>What a host sends to open one of its files (<c>POST /api/v1/host/sessions</c>).</summary>
public sealed record HostSessionRequest(
    string? FileId,
    string? FileName,
    string? ContentUrl,
    string? AccessToken,
    bool CanWrite,
    HostUser? User,
    string? ReturnUrl);

public sealed record HostUser(string? Id, string? Name, string? Email);

/// <summary>The answer to the host: the address the browser opens.</summary>
public sealed record HostSessionStarted(Guid Token, string Path, string Kind, bool ReadOnly, DateTime Expires);

/// <summary>What the editor reads in a host session.</summary>
public sealed record HostEditorData(Guid Token, DocumentKind Kind, string Name, string FileName, string Format, long Revision, string Content, bool CanWrite, string? ReturnUrl, string UserName, string HostName, string? LastError);

public enum HostSaveStatus
{
    Saved,
    Conflict,
    HostChanged,
    HostFailed,
    Invalid,
    Denied,
}

public sealed record HostSaveResult(HostSaveStatus Status, long Revision = 0, string? Content = null, string? Problem = null);

/// <summary>
/// The files of hosts that are open in MatOffice (like WOPI): the host names where the file is read and written and gives a token for it, MatOffice reads the
/// file, converts it into the content of its kind for the editor, and writes it back in its own format on every save (with <c>If-Match</c>, so a file that was changed at
/// the host is not overwritten). The browser only knows the secret of the session.
/// </summary>
public static class HostSessionService
{
    /// <summary>How long the address of a session works.</summary>
    public static readonly TimeSpan Lifetime = TimeSpan.FromHours(12);

    public const long MaxFileBytes = 100L * 1024 * 1024;

    public static async Task<Outcome<HostSessionStarted>> StartAsync(AppDbContext db, HostClient client, HostSessionRequest request, CancellationToken ct = default)
    {
        var fileName = (request.FileName ?? string.Empty).Trim();
        if (fileName.Length is 0 or > 255)
        {
            return Outcome<HostSessionStarted>.Fail("fileName is required (at most 255 characters).");
        }

        var ext = Kinds.NormalizeExtension(Path.GetExtension(fileName));
        if (Kinds.ForExtension(ext) is not { } found)
        {
            return Outcome<HostSessionStarted>.Fail($"MatOffice cannot open files of the type '{ext}'.");
        }

        if (!Uri.TryCreate(request.ContentUrl, UriKind.Absolute, out var url) || url.Scheme is not ("http" or "https"))
        {
            return Outcome<HostSessionStarted>.Fail("contentUrl must be an absolute http(s) address.");
        }

        if (client.AllowedUrlPrefix is { } prefix && !url.AbsoluteUri.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return Outcome<HostSessionStarted>.Fail("contentUrl is not an address this host may use.");
        }

        if (string.IsNullOrEmpty(request.AccessToken) || request.AccessToken.Length > 500)
        {
            return Outcome<HostSessionStarted>.Fail("accessToken is required (at most 500 characters).");
        }

        string? back = null;
        if (!string.IsNullOrEmpty(request.ReturnUrl))
        {
            if (!Uri.TryCreate(request.ReturnUrl, UriKind.Absolute, out var r) || r.Scheme is not ("http" or "https") || request.ReturnUrl.Length > 1000)
            {
                return Outcome<HostSessionStarted>.Fail("returnUrl must be an absolute http(s) address.");
            }

            back = r.AbsoluteUri;
        }

        var canWrite = request.CanWrite && found.Format.CanExport;
        var session = new HostSession
        {
            HostClientId = client.Id,
            FileId = Cut(request.FileId ?? string.Empty, 200),
            FileName = fileName,
            Kind = found.Handler.Kind,
            Format = ext,
            ContentUrl = url.AbsoluteUri,
            AccessToken = request.AccessToken,
            CanWrite = canWrite,
            UserRef = Cut(request.User?.Id ?? string.Empty, 200),
            UserName = Cut(string.IsNullOrWhiteSpace(request.User?.Name) ? "?" : request.User!.Name!.Trim(), 120),
            UserEmail = string.IsNullOrWhiteSpace(request.User?.Email) ? null : Cut(request.User!.Email!.Trim(), 254),
            ReturnUrl = back,
            ExpiresDate = DateTime.UtcNow + Lifetime,
        };
        db.HostSessions.Add(session);
        await db.SaveChangesAsync(ct);
        return Outcome<HostSessionStarted>.Success(new HostSessionStarted(session.Token, "/h/" + session.Token.ToString("N"), found.Handler.Slug, !canWrite, session.ExpiresDate));
    }

    /// <summary>The session of an editor address, while it is valid (and its host is active).</summary>
    public static Task<HostSession?> FindAsync(AppDbContext db, Guid token, CancellationToken ct = default)
    {
        var now = DateTime.UtcNow;
        return db.HostSessions.Include(s => s.HostClient)
            .FirstOrDefaultAsync(s => s.Token == token && s.ExpiresDate > now && s.HostClient!.IsActive && s.HostClient.UpdateState != SoftEntity.Deleted, ct);
    }

    /// <summary>The file for the editor: read from the host the first time (or after <paramref name="reload"/>), then the working copy.</summary>
    public static async Task<Outcome<HostEditorData>> OpenAsync(AppDbContext db, HttpClient http, Guid token, bool reload = false, CancellationToken ct = default)
    {
        var session = await FindAsync(db, token, ct);
        if (session is null)
        {
            return Outcome<HostEditorData>.Denied();
        }

        if (session.Content is null || reload)
        {
            var handler = Kinds.Of(session.Kind);
            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, session.ContentUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", session.AccessToken);
                using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
                if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden or HttpStatusCode.NotFound)
                {
                    return Outcome<HostEditorData>.Fail(Text.T("Host.FileGone"));
                }

                if (!response.IsSuccessStatusCode)
                {
                    return Outcome<HostEditorData>.Fail(Text.T("Host.ReadFailed", (int)response.StatusCode));
                }

                if (response.Content.Headers.ContentLength > MaxFileBytes)
                {
                    return Outcome<HostEditorData>.Fail(Text.T("Host.TooLarge"));
                }

                await using var body = await response.Content.ReadAsStreamAsync(ct);
                using var buffer = new MemoryStream();
                await CopyLimitedAsync(body, buffer, MaxFileBytes, ct);
                buffer.Position = 0;

                // an empty file (a new file of the host) starts as a blank document
                var read = buffer.Length == 0 ? handler.Normalize(handler.Create("blank", Languages.Current)) : Kinds.Read(handler, buffer, session.Format);
                if (!read.Ok)
                {
                    return Outcome<HostEditorData>.Fail(read.Error!);
                }

                session.Content = read.Json;
                session.Revision++;
                session.HostVersion = response.Headers.ETag?.Tag;
                session.LastError = null;
                await db.SaveChangesAsync(ct);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidDataException)
            {
                return Outcome<HostEditorData>.Fail(Text.T("Host.Unreachable"));
            }
        }

        return Outcome<HostEditorData>.Success(Data(session));
    }

    /// <summary>
    /// Saves the working copy and writes the file back to the host. The working copy is kept even when the host cannot be reached (the next save tries again),
    /// so nothing typed is lost.
    /// </summary>
    public static async Task<HostSaveResult> SaveAsync(AppDbContext db, HttpClient http, Guid token, long revision, string content, CancellationToken ct = default)
    {
        var session = await FindAsync(db, token, ct);
        if (session is null || !session.CanWrite)
        {
            return new HostSaveResult(HostSaveStatus.Denied);
        }

        if (session.Revision != revision)
        {
            return new HostSaveResult(HostSaveStatus.Conflict, session.Revision, session.Content);
        }

        var handler = Kinds.Of(session.Kind);
        var check = handler.Normalize(content);
        if (!check.Ok)
        {
            return new HostSaveResult(HostSaveStatus.Invalid, Problem: check.Error);
        }

        byte[] bytes;
        try
        {
            bytes = Kinds.Write(handler, check.Json!, session.Format, Path.GetFileNameWithoutExtension(session.FileName));
        }
        catch (Formats.FormatException ex)
        {
            return new HostSaveResult(HostSaveStatus.Invalid, Problem: ex.Message);
        }

        // the working copy first: what was typed is safe in MatOffice whatever the host answers
        var changed = await db.HostSessions.Where(s => s.Id == session.Id && s.Revision == revision)
            .ExecuteUpdateAsync(s => s.SetProperty(x => x.Content, check.Json).SetProperty(x => x.Revision, revision + 1).SetProperty(x => x.UpdateDate, DateTime.UtcNow), ct);
        if (changed == 0)
        {
            await db.Entry(session).ReloadAsync(ct);
            return new HostSaveResult(HostSaveStatus.Conflict, session.Revision, session.Content);
        }

        string? problem;
        HostSaveStatus status;
        string? newVersion = session.HostVersion;
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Put, session.ContentUrl) { Content = new ByteArrayContent(bytes) };
            request.Content.Headers.ContentType = new MediaTypeHeaderValue(Kinds.MimeOf(handler, session.Format));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", session.AccessToken);
            if (!string.IsNullOrEmpty(session.HostVersion) && EntityTagHeaderValue.TryParse(session.HostVersion, out var tag))
            {
                request.Headers.IfMatch.Add(tag);
            }

            request.Headers.TryAddWithoutValidation("X-MatOffice-User", Uri.EscapeDataString(session.UserName));
            using var response = await http.SendAsync(request, ct);
            if (response.IsSuccessStatusCode)
            {
                (status, problem) = (HostSaveStatus.Saved, null);
                newVersion = response.Headers.ETag?.Tag ?? newVersion;
            }
            else if (response.StatusCode is HttpStatusCode.PreconditionFailed or HttpStatusCode.Conflict)
            {
                (status, problem) = (HostSaveStatus.HostChanged, Text.T("Host.ChangedThere"));
            }
            else if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden or HttpStatusCode.NotFound)
            {
                (status, problem) = (HostSaveStatus.HostFailed, Text.T("Host.FileGone"));
            }
            else if (response.StatusCode == HttpStatusCode.RequestEntityTooLarge || (int)response.StatusCode == 507)
            {
                (status, problem) = (HostSaveStatus.HostFailed, Text.T("Host.NoSpace"));
            }
            else
            {
                (status, problem) = (HostSaveStatus.HostFailed, Text.T("Host.WriteFailed", (int)response.StatusCode));
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            (status, problem) = (HostSaveStatus.HostFailed, Text.T("Host.Unreachable"));
        }

        var now = DateTime.UtcNow;
        await db.HostSessions.Where(s => s.Id == session.Id).ExecuteUpdateAsync(s => s
            .SetProperty(x => x.LastError, problem)
            .SetProperty(x => x.HostVersion, newVersion)
            .SetProperty(x => x.LastSavedDate, status == HostSaveStatus.Saved ? now : session.LastSavedDate)
            .SetProperty(x => x.ExpiresDate, now + Lifetime), ct);
        return new HostSaveResult(status, revision + 1, Problem: problem);
    }

    /// <summary>The working copy as a file of another format of its kind (Download as ...).</summary>
    public static async Task<Outcome<Documents.FileDownload>> ExportAsync(AppDbContext db, Guid token, string extension, CancellationToken ct = default)
    {
        var session = await FindAsync(db, token, ct);
        if (session?.Content is null)
        {
            return Outcome<Documents.FileDownload>.Denied();
        }

        return Documents.DocumentService.Export(session.Kind, session.Content, Path.GetFileNameWithoutExtension(session.FileName), extension);
    }

    /// <summary>The revision of the working copy (the editor polls it like for a document).</summary>
    public static async Task<long?> RevisionAsync(AppDbContext db, Guid token, CancellationToken ct = default)
        => (await FindAsync(db, token, ct))?.Revision;

    /// <summary>Ends a session (the editor closes); the working copy goes.</summary>
    public static Task<int> CloseAsync(AppDbContext db, Guid token, CancellationToken ct = default)
        => db.HostSessions.Where(s => s.Token == token).ExecuteDeleteAsync(ct);

    /// <summary>The housekeeping: expired sessions go.</summary>
    public static Task<int> CleanUpAsync(AppDbContext db, CancellationToken ct = default)
    {
        var now = DateTime.UtcNow;
        return db.HostSessions.Where(s => s.ExpiresDate < now).ExecuteDeleteAsync(ct);
    }

    /// <summary>A blank file of a kind in one of its formats (for "New document" at the host).</summary>
    public static Outcome<byte[]> Blank(string kindSlug, string? extension, string name, string language)
    {
        var handler = Kinds.BySlug(kindSlug);
        if (handler is null)
        {
            return Outcome<byte[]>.Fail("Unknown kind.");
        }

        var ext = string.IsNullOrEmpty(extension) ? handler.NativeExtension : Kinds.NormalizeExtension(extension);
        if (!Kinds.ExportFormats(handler).Any(f => f.Extension == ext))
        {
            return Outcome<byte[]>.Fail($"MatOffice cannot write files of the type '{ext}'.");
        }

        var content = handler.Normalize(handler.Create("blank", language));
        try
        {
            return content.Ok ? Outcome<byte[]>.Success(Kinds.Write(handler, content.Json!, ext, name)) : Outcome<byte[]>.Fail(content.Error!);
        }
        catch (Formats.FormatException ex)
        {
            return Outcome<byte[]>.Fail(ex.Message);
        }
    }

    private static HostEditorData Data(HostSession s) => new(
        s.Token, s.Kind, Path.GetFileNameWithoutExtension(s.FileName), s.FileName, s.Format, s.Revision, s.Content ?? "{}", s.CanWrite, s.ReturnUrl, s.UserName,
        s.HostClient?.Name ?? string.Empty, s.LastError);

    private static async Task CopyLimitedAsync(Stream from, Stream to, long max, CancellationToken ct)
    {
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await from.ReadAsync(buffer, ct)) > 0)
        {
            total += read;
            if (total > max)
            {
                throw new InvalidDataException("The file is too large.");
            }

            await to.WriteAsync(buffer.AsMemory(0, read), ct);
        }
    }

    private static string Cut(string text, int max) => text.Length > max ? text[..max] : text;
}
