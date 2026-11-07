using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Data;
using MatOffice.Localization;
using MatOffice.Services.Documents;
using MatOffice.Services.Formats;

namespace MatOffice.Services.Hosts;

/// <summary>
/// The integration of hosts (MatDrive and others): the API the host calls with its key (<c>Authorization: Bearer mo_...</c>, under <c>/api/v1/host</c>) and the
/// addresses the editor of a host file talks to (<c>/live/host/{token}</c>, the secret of the session is the right; the anti-forgery token of the page for changes).
/// </summary>
public static class HostEndpoints
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public static void MapHostEndpoints(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/v1/host").AllowAnonymous();

        // what MatOffice can open and write: the host shows "Open with MatOffice" and "New" from it
        api.MapGet("/formats", async (HttpContext http, AppDbContext db) =>
        {
            if (await ClientAsync(http, db) is null)
            {
                return Results.Unauthorized();
            }

            var kinds = new JsonArray();
            foreach (var h in Kinds.All)
            {
                var formats = new JsonArray();
                foreach (var f in h.Formats.Append(Kinds.NativeFormat(h)))
                {
                    formats.Add(new JsonObject { ["ext"] = f.Extension, ["label"] = f.Label, ["mime"] = f.MimeType, ["open"] = f.CanImport, ["write"] = f.CanExport });
                }

                kinds.Add(new JsonObject
                {
                    ["kind"] = h.Slug,
                    ["name"] = Text.T("Kinds." + h.Kind),
                    ["colour"] = h.Colour,
                    ["native"] = h.NativeExtension,
                    ["newFormat"] = h.Formats.FirstOrDefault(f => f.CanImport && f.CanExport)?.Extension ?? h.NativeExtension,
                    ["formats"] = formats,
                });
            }

            return Results.Text(new JsonObject { ["app"] = "matoffice", ["version"] = AppInfo.Version, ["kinds"] = kinds }.ToJsonString(), "application/json");
        });

        // a blank file of a kind (the host stores it and opens it like any other file)
        api.MapGet("/new", async (string kind, string? format, string? name, string? language, HttpContext http, AppDbContext db) =>
        {
            if (await ClientAsync(http, db) is null)
            {
                return Results.Unauthorized();
            }

            var blank = HostSessionService.Blank(kind, format, string.IsNullOrWhiteSpace(name) ? "document" : name, Languages.Normalize(language));
            return blank.Ok ? Results.File(blank.Value!, "application/octet-stream") : Results.Json(new { error = blank.Error }, Json, statusCode: StatusCodes.Status400BadRequest);
        });

        api.MapPost("/sessions", async (HostSessionRequest body, HttpContext http, AppDbContext db) =>
        {
            var client = await ClientAsync(http, db);
            if (client is null)
            {
                return Results.Unauthorized();
            }

            var started = await HostSessionService.StartAsync(db, client, body, http.RequestAborted);
            if (!started.Ok)
            {
                return Results.Json(new { error = started.Error }, Json, statusCode: StatusCodes.Status400BadRequest);
            }

            var s = started.Value!;
            var settings = await http.RequestServices.GetRequiredService<AccountService>().SettingsAsync();
            var origin = InstanceService.NormalizePublicUrl(settings.PublicUrl) ?? $"{http.Request.Scheme}://{http.Request.Host}";
            return Results.Json(new { token = s.Token.ToString("N"), path = s.Path, url = origin + s.Path, kind = s.Kind, readOnly = s.ReadOnly, expires = s.Expires }, Json);
        });

        var live = app.MapGroup("/live/host/{token:guid}").AllowAnonymous();

        live.MapGet("", async (Guid token, bool? reload, HttpContext http, AppDbContext db, IHttpClientFactory clients) =>
        {
            http.Response.Headers.CacheControl = "no-store";
            var data = await HostSessionService.OpenAsync(db, clients.CreateClient("host"), token, reload == true, http.RequestAborted);
            return data.NotAllowed ? Results.NotFound()
                : data.Ok ? Results.Text(DocumentEndpoints.EditorJson(data.Value!).ToJsonString(), "application/json")
                : Results.Json(new { error = data.Error }, Json, statusCode: StatusCodes.Status502BadGateway);
        });

        live.MapPut("", async (Guid token, SaveRequest body, HttpContext http, AppDbContext db, IHttpClientFactory clients) =>
        {
            var content = body.Content.ValueKind == JsonValueKind.Undefined ? string.Empty : body.Content.GetRawText();
            var result = await HostSessionService.SaveAsync(db, clients.CreateClient("host"), token, body.Revision, content, http.RequestAborted);
            return result.Status switch
            {
                HostSaveStatus.Saved => Results.Json(new { revision = result.Revision, savedAt = DateTime.UtcNow }, Json),
                HostSaveStatus.Conflict => Results.Text(new JsonObject { ["revision"] = result.Revision, ["content"] = JsonNode.Parse(result.Content ?? "{}") }.ToJsonString(), "application/json", statusCode: StatusCodes.Status409Conflict),
                HostSaveStatus.HostChanged => Results.Json(new { revision = result.Revision, error = result.Problem, hostChanged = true }, Json, statusCode: StatusCodes.Status412PreconditionFailed),
                HostSaveStatus.HostFailed => Results.Json(new { revision = result.Revision, error = result.Problem }, Json, statusCode: StatusCodes.Status502BadGateway),
                HostSaveStatus.Invalid => Results.Json(new { error = result.Problem }, Json, statusCode: StatusCodes.Status422UnprocessableEntity),
                _ => Results.NotFound(),
            };
        }).AddEndpointFilter(DocumentEndpoints.Antiforgery).WithRequestSizeLimit();

        live.MapGet("/revision", async (Guid token, HttpContext http, AppDbContext db) =>
        {
            http.Response.Headers.CacheControl = "no-store";
            return await HostSessionService.RevisionAsync(db, token, http.RequestAborted) is { } r ? Results.Json(new { revision = r, by = (string?)null }, Json) : Results.NotFound();
        });

        live.MapGet("/export", async (Guid token, string? format, HttpContext http, AppDbContext db) =>
        {
            var file = await HostSessionService.ExportAsync(db, token, format ?? string.Empty, http.RequestAborted);
            return file.NotAllowed ? Results.NotFound()
                : file.Ok ? Results.File(file.Value!.Bytes, file.Value.MimeType, file.Value.FileName)
                : Results.Json(new { error = file.Error }, Json, statusCode: StatusCodes.Status422UnprocessableEntity);
        });

        live.MapPost("/close", async (Guid token, HttpContext http, AppDbContext db) =>
        {
            await HostSessionService.CloseAsync(db, token, http.RequestAborted);
            return Results.NoContent();
        }).AddEndpointFilter(DocumentEndpoints.Antiforgery);
    }

    /// <summary>The host of the key in <c>Authorization: Bearer</c> (or <c>X-MatOffice-Key</c>).</summary>
    private static Task<HostClient?> ClientAsync(HttpContext http, AppDbContext db)
    {
        var header = http.Request.Headers.Authorization.ToString();
        var key = header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? header[7..].Trim() : http.Request.Headers["X-MatOffice-Key"].ToString();
        return HostAdmin.AuthenticateAsync(db, key, http.RequestAborted);
    }
}
