using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using MatOffice.Services.Hosts;
using Microsoft.AspNetCore.Antiforgery;

namespace MatOffice.Services.Documents;

/// <summary>What the editor sends to save.</summary>
public sealed record SaveRequest(long Revision, JsonElement Content);

public sealed record NameRequest(string? Name);

public sealed record VersionRequest(string? Name, string? Description);

public sealed record StarRequest(bool Starred);

/// <summary>
/// The addresses the editors talk to (JSON under <c>/live</c>). Every editor reads the same shape (<see cref="EditorJson"/>), whether the file is a document of
/// MatOffice (<c>/live/documents/{id}</c>) or a file of a host such as MatDrive (<c>/live/host/{token}</c>): the field <c>api</c> says where it saves.
/// Reading is open to the anonymous request on purpose, because a link gives guests a look: what a request may see is decided by the business layer (the cookie
/// of a person and/or the secret of the link in <c>?t=</c>). Every change needs the anti-forgery token of the page (header <c>RequestVerificationToken</c>).
/// </summary>
public static class DocumentEndpoints
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public static void MapDocumentEndpoints(this IEndpointRouteBuilder app)
    {
        var live = app.MapGroup("/live/documents");

        live.MapGet("/{id:long}", async (long id, long? v, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var data = await DocumentService.OpenAsync(db, ActorOf(http, me), id, v, http.RequestAborted);
            http.Response.Headers.CacheControl = "no-store";
            if (data is null)
            {
                return Results.NotFound();
            }

            var link = http.Request.Query["t"].ToString();
            var api = $"/live/documents/{id}" + (Guid.TryParse(link, out _) ? "?t=" + link : string.Empty);
            return Results.Text(EditorJson(data, api, me).ToJsonString(), "application/json");
        }).AllowAnonymous();

        live.MapPut("/{id:long}", async (long id, SaveRequest body, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var content = body.Content.ValueKind == JsonValueKind.Undefined ? string.Empty : body.Content.GetRawText();
            var result = await DocumentService.SaveAsync(db, me.RequiredId, id, body.Revision, content, http.RequestAborted);
            return result.Status switch
            {
                SaveStatus.Saved => Results.Json(new { revision = result.Revision, savedAt = DateTime.UtcNow }, Json),
                SaveStatus.Conflict => Results.Text(new JsonObject { ["revision"] = result.Revision, ["content"] = JsonNode.Parse(result.Content!) }.ToJsonString(), "application/json", statusCode: StatusCodes.Status409Conflict),
                SaveStatus.Invalid => Results.Json(new { error = result.Problem }, Json, statusCode: StatusCodes.Status422UnprocessableEntity),
                _ => Results.NotFound(),
            };
        }).AddEndpointFilter(Antiforgery).WithRequestSizeLimit();

        live.MapGet("/{id:long}/revision", async (long id, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var revision = await DocumentService.RevisionAsync(db, ActorOf(http, me), id, http.RequestAborted);
            http.Response.Headers.CacheControl = "no-store";
            return revision is { } r ? Results.Json(new { revision = r.Revision, by = r.By }, Json) : Results.NotFound();
        }).AllowAnonymous();

        live.MapPatch("/{id:long}", async (long id, NameRequest body, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var renamed = await DocumentService.RenameAsync(db, me.RequiredId, id, body.Name ?? string.Empty, http.RequestAborted);
            return renamed.NotAllowed ? Results.NotFound()
                : renamed.Ok ? Results.Json(new { name = renamed.Value }, Json)
                : Results.Json(new { error = renamed.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        }).AddEndpointFilter(Antiforgery);

        live.MapPost("/{id:long}/star", async (long id, StarRequest body, HttpContext http, AppDbContext db, CurrentUser me) =>
            await DocumentService.StarAsync(db, me.RequiredId, id, body.Starred, http.RequestAborted) ? Results.NoContent() : Results.NotFound()).AddEndpointFilter(Antiforgery);

        live.MapPost("/{id:long}/duplicate", async (long id, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var copy = await DocumentService.DuplicateAsync(db, me.RequiredId, id, http.RequestAborted);
            return copy.Ok ? Results.Json(new { id = copy.Value!.Id, url = EditorUrl(copy.Value.Kind, copy.Value.Id) }, Json) : Results.NotFound();
        }).AddEndpointFilter(Antiforgery);

        live.MapGet("/{id:long}/export", async (long id, string? format, long? v, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var file = await DocumentService.ExportAsync(db, ActorOf(http, me), id, format ?? string.Empty, v, http.RequestAborted);
            return file.NotAllowed ? Results.NotFound()
                : file.Ok ? Results.File(file.Value!.Bytes, file.Value.MimeType, file.Value.FileName)
                : Results.Json(new { error = file.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        }).AllowAnonymous();

        live.MapGet("/{id:long}/thumbnail", async (long id, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var png = await DocumentService.ThumbnailAsync(db, ActorOf(http, me), id, http.RequestAborted);
            if (png is null)
            {
                return Results.NotFound();
            }

            var tag = new Microsoft.Net.Http.Headers.EntityTagHeaderValue($"\"{Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(png))[..16]}\"");
            http.Response.Headers.CacheControl = "private, no-cache";
            return Results.File(png, "image/png", entityTag: tag);
        }).AllowAnonymous();

        live.MapPut("/{id:long}/thumbnail", async (long id, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            using var buffer = new MemoryStream();
            await http.Request.Body.CopyToAsync(buffer, http.RequestAborted);
            if (buffer.Length > DocumentService.ThumbnailMax)
            {
                return Results.StatusCode(StatusCodes.Status413PayloadTooLarge);
            }

            return await DocumentService.SetThumbnailAsync(db, me.RequiredId, id, buffer.ToArray(), http.RequestAborted) ? Results.NoContent() : Results.NotFound();
        }).AddEndpointFilter(Antiforgery);

        live.MapGet("/{id:long}/versions", async (long id, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var rows = await VersionService.ListAsync(db, ActorOf(http, me), id, http.RequestAborted);
            http.Response.Headers.CacheControl = "no-store";
            return rows is null ? Results.NotFound() : Results.Json(rows.Take(100), Json);
        }).AllowAnonymous();

        live.MapPost("/{id:long}/versions", async (long id, VersionRequest body, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            var created = await VersionService.CreateAsync(db, me.RequiredId, id, body.Name ?? string.Empty, body.Description, http.RequestAborted);
            return created.NotAllowed ? Results.NotFound()
                : created.Ok ? Results.Json(new { created.Value!.Id, created.Value.Name }, Json)
                : Results.Json(new { error = created.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        }).AddEndpointFilter(Antiforgery);

        // pictures that are put into a document are kept inside it (data: addresses); this only checks and shrinks nothing, it tells the editor the limit
        live.MapGet("/limits", () => Results.Json(new { maxContentBytes = Kinds.MaxContentBytes, maxImageBytes = 5 * 1024 * 1024 }, Json)).AllowAnonymous();

        // conversions for an editor that is open (no document is read or written): a file into the content of an app (Base takes an Excel file as tables, Slides
        // the slides of a PowerPoint file …) and content into a file (the state on the screen, not the saved one). A signed-in person or the secret of an open host session.
        var convert = app.MapGroup("/live/convert");
        convert.MapPost("/{slug}/read", async (string slug, string ext, Guid? host, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            if (!await MayConvertAsync(db, me, host, http.RequestAborted))
            {
                return Results.NotFound();
            }

            if (Kinds.BySlug(slug) is not { } handler)
            {
                return Results.NotFound();
            }

            using var buffer = new MemoryStream();
            await http.Request.Body.CopyToAsync(buffer, http.RequestAborted);
            buffer.Position = 0;
            var read = Kinds.Read(handler, buffer, ext);
            return read.Ok ? Results.Text(read.Json!, "application/json") : Results.Json(new { error = read.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        }).AllowAnonymous().AddEndpointFilter(Antiforgery).WithRequestSizeLimit();

        convert.MapPost("/{slug}/write", async (string slug, string ext, string? name, Guid? host, HttpContext http, AppDbContext db, CurrentUser me) =>
        {
            if (!await MayConvertAsync(db, me, host, http.RequestAborted) || Kinds.BySlug(slug) is not { } handler)
            {
                return Results.NotFound();
            }

            using var reader = new StreamReader(http.Request.Body);
            var content = await reader.ReadToEndAsync(http.RequestAborted);
            var check = handler.Normalize(content);
            if (!check.Ok)
            {
                return Results.Json(new { error = check.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
            }

            var file = DocumentService.Export(handler.Kind, check.Json!, string.IsNullOrWhiteSpace(name) ? "document" : name, ext);
            return file.Ok ? Results.File(file.Value!.Bytes, file.Value.MimeType, file.Value.FileName) : Results.Json(new { error = file.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        }).AllowAnonymous().AddEndpointFilter(Antiforgery).WithRequestSizeLimit();
    }

    /// <summary>A conversion is for a signed-in person, or for the editor of an open host session (whose browser knows its secret).</summary>
    private static async Task<bool> MayConvertAsync(AppDbContext db, CurrentUser me, Guid? host, CancellationToken ct)
        => me.Id is not null || (host is { } token && await HostSessionService.FindAsync(db, token, ct) is not null);

    /// <summary>The person of the cookie (if any) and the secret of the link (if one is given).</summary>
    public static Actor ActorOf(HttpContext http, CurrentUser me)
    {
        var actor = new Actor(me.Id);
        return Guid.TryParse(http.Request.Query["t"], out var token) ? actor.WithLink(token) : actor;
    }

    /// <summary>The address of the editor of a document.</summary>
    public static string EditorUrl(DocumentKind kind, long id) => $"/{Kinds.SlugOf(kind)}/{id}";

    /// <summary>The formats of a kind as the editor lists them in "Download as".</summary>
    public static JsonArray FormatsJson(IKindHandler handler)
    {
        var list = new JsonArray();
        foreach (var f in Kinds.ExportFormats(handler))
        {
            list.Add(new JsonObject { ["ext"] = f.Extension, ["label"] = f.Label, ["mime"] = f.MimeType });
        }

        return list;
    }

    /// <summary>A document as the editor reads it.</summary>
    public static JsonObject EditorJson(EditorData data, string api, CurrentUser me)
    {
        var handler = Kinds.Of(data.Kind);
        return new JsonObject
        {
            ["mode"] = data.VersionId is not null ? "version" : "document",
            ["id"] = data.Id,
            ["kind"] = handler.Slug,
            ["name"] = data.Name,
            ["revision"] = data.Revision,
            ["role"] = data.Role.ToString().ToLowerInvariant(),
            ["canEdit"] = data.CanEdit && me.Id is not null,
            ["signedIn"] = me.Id is not null,
            ["starred"] = data.IsStarred,
            ["owner"] = data.OwnerName,
            ["versionId"] = data.VersionId,
            ["versionName"] = data.VersionName,
            ["api"] = api,
            ["formats"] = FormatsJson(handler),
            ["user"] = me.Snapshot is { } u ? new JsonObject { ["id"] = u.Id, ["name"] = u.DisplayName } : null,
            ["content"] = JsonNode.Parse(data.Content),
        };
    }

    /// <summary>A file of a host as the editor reads it.</summary>
    public static JsonObject EditorJson(HostEditorData data)
    {
        var handler = Kinds.Of(data.Kind);
        var token = data.Token.ToString("N");
        return new JsonObject
        {
            ["mode"] = "host",
            ["id"] = null,
            ["kind"] = handler.Slug,
            ["name"] = data.Name,
            ["fileName"] = data.FileName,
            ["format"] = data.Format,
            ["revision"] = data.Revision,
            ["role"] = data.CanWrite ? "editor" : "viewer",
            ["canEdit"] = data.CanWrite,
            ["signedIn"] = false,
            ["host"] = data.HostName,
            ["returnUrl"] = data.ReturnUrl,
            ["lastError"] = data.LastError,
            ["api"] = "/live/host/" + token,
            ["formats"] = FormatsJson(handler),
            ["user"] = new JsonObject { ["id"] = null, ["name"] = data.UserName },
            ["content"] = JsonNode.Parse(data.Content),
        };
    }

    /// <summary>The anti-forgery token of the page in the header <c>RequestVerificationToken</c>; without it a change answers 400.</summary>
    public static async ValueTask<object?> Antiforgery(EndpointFilterInvocationContext context, EndpointFilterDelegate next)
    {
        var anti = context.HttpContext.RequestServices.GetRequiredService<IAntiforgery>();
        try
        {
            await anti.ValidateRequestAsync(context.HttpContext);
        }
        catch (AntiforgeryValidationException)
        {
            return Results.BadRequest();
        }

        return await next(context);
    }

    /// <summary>A document with its pictures may be large: the body of a save may be as large as the largest content.</summary>
    public static RouteHandlerBuilder WithRequestSizeLimit(this RouteHandlerBuilder builder)
        => builder.WithMetadata(new Microsoft.AspNetCore.Mvc.RequestSizeLimitAttribute(Kinds.MaxContentBytes + 1024 * 1024));
}
