using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using MatOffice.Data;
using MatOffice.Platform;
using MatOffice.Services.Hosts;
using MatPlatform;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.EntityFrameworkCore;

namespace MatOffice.Tests;

/// <summary>
/// MatOffice in the platform contract 2 (MatPlatform 2.0): the definition is valid, the bindings of <c>office.edit</c> are the hosts (bind, renew, list, unbind),
/// and a host key opens <c>/api/provides/office.edit/...</c> as it opened <c>/api/v1/host/...</c>.
/// </summary>
public sealed class PlatformTests
{
    private static readonly MatBindingParty Office = new("matoffice", "MatOffice", "https://office.example.com", "http://mattiverse-matoffice-web:4820");

    [Theory]
    [InlineData("de")]
    [InlineData("en")]
    public void The_definition_is_valid(string language)
    {
        var definition = MatPlatformBuilder.Build(MatOfficePlatform.Describe, language);
        var sprite = MatPlatformValidator.ReadSprite(Path.Combine(Root(), "src", "MatOffice", "wwwroot", "matui", "matui-icons.svg"));
        Assert.Empty(MatPlatformValidator.Validate(definition, sprite));
        Assert.Equal(MatApiDefinition.ContractVersion, definition.Contract);
    }

    [Fact]
    public void MAT_PUBLIC_URL_fixes_the_public_address()
    {
        var free = MatPlatformBuilder.Build(o => MatOfficePlatform.Describe(o, null), "en").Configure.Single(f => f.Key == "publicUrl");
        Assert.Null(free.Fixed);

        var fixedField = MatPlatformBuilder.Build(o => MatOfficePlatform.Describe(o, "https://office.example.com"), "en").Configure.Single(f => f.Key == "publicUrl");
        Assert.True(fixedField.Fixed);
        Assert.Contains("MAT_PUBLIC_URL", fixedField.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void It_provides_office_edit_contributes_the_four_apps_and_names_what_it_consumes()
    {
        var definition = MatPlatformBuilder.Build(MatOfficePlatform.Describe, "en");
        Assert.Equal(MatCapabilities.OfficeEdit, Assert.Single(definition.Provides).Capability);
        Assert.Equal(new[] { MatCapabilities.IdentityOidc, MatCapabilities.MailSmtp }, definition.Consumes.Select(c => c.Capability));
        Assert.All(definition.Consumes, c => Assert.True(c.Optional));

        var contributes = Assert.IsType<MatContributions>(definition.Contributes);
        Assert.Equal(new[] { "docs", "sheets", "slides", "base" }, contributes.Launchers.Select(l => l.Id));
        foreach (var launcher in contributes.Launchers)
        {
            // the picture is a file of the app, the address the list of its app, "New …" the page New with the app chosen
            Assert.True(File.Exists(Path.Combine(Root(), "src", "MatOffice", "wwwroot", launcher.Icon.TrimStart('/').Replace('/', Path.DirectorySeparatorChar))), launcher.Icon);
            Assert.Equal("/" + launcher.Id, launcher.Url);
            Assert.Equal("/Documents/New?kind=" + launcher.Id, Assert.Single(contributes.QuickActions, q => q.Launcher == launcher.Id).Url);
        }

        Assert.NotEmpty(contributes.Widgets);
        Assert.All(contributes.Widgets, w => Assert.Equal(MatAudience.Admin, w.Audience));
        Assert.Contains(definition.Actions, a => a.Id == "openConnections" && a.Url == "/Admin/Connections");
    }

    [DbFact]
    public async Task The_handler_binds_renews_lists_and_unbinds_hosts()
    {
        await using var t = await TestDatabase.CreateAsync();

        // a host made by hand before the bindings: no consumer id
        var legacy = new HostClient { Name = "Old drive" };
        var legacyKey = HostAdmin.IssueKey(legacy);
        t.Db.HostClients.Add(legacy);
        await t.Db.SaveChangesAsync();

        var handler = new MatOfficeHostBindings(t.Db);
        Assert.Equal(MatCapabilities.OfficeEdit, handler.Capability);

        // a broker binds MatDrive: the values are the addresses of MatOffice and a new key in the format of the hosts
        var made = await handler.BindAsync(Request("mattiverse", renew: false), default);
        Assert.Equal("https://office.example.com", made.Values[MatCapabilities.Url]);
        Assert.Equal("http://mattiverse-matoffice-web:4820", made.Values[MatCapabilities.InternalUrl]);
        var key = made.Values[MatCapabilities.Key]!;
        Assert.StartsWith(HostAdmin.KeyPrefix, key, StringComparison.Ordinal);
        Assert.Equal("matdrive", (await HostAdmin.AuthenticateAsync(t.Db, key))!.ConsumerId);

        // the same request again is the same host, and the key is not handed out again
        var again = await handler.BindAsync(Request("mattiverse", renew: false), default);
        Assert.Null(again.Values[MatCapabilities.Key]);
        Assert.Equal(2, await t.Db.HostClients.CountAsync());

        // both hosts are listed: the old one as host-<id>, by hand
        var list = await handler.ListAsync(default);
        Assert.Equal(2, list.Count);
        var drive = Assert.Single(list, b => b.Consumer == "matdrive");
        Assert.Equal(("MatDrive", "mattiverse"), (drive.Name, drive.ManagedBy));
        Assert.NotNull(drive.Since);
        var old = Assert.Single(list, b => b.Consumer == $"host-{legacy.Id}");
        Assert.Equal(("Old drive", HostBindings.ManagedByHand), (old.Name, old.ManagedBy));

        // a new key: the old one is refused and the files it opened close
        var driveId = (await HostAdmin.AuthenticateAsync(t.Db, key))!.Id;
        t.Db.HostSessions.Add(new HostSession { HostClientId = driveId, FileId = "1", FileName = "a.docx", Kind = DocumentKind.Docs, Format = ".docx", ContentUrl = "https://drive.example.com/f/1", ExpiresDate = DateTime.UtcNow.AddHours(1) });
        await t.Db.SaveChangesAsync();
        var renewed = (await handler.BindAsync(Request("mattiverse", renew: true), default)).Values[MatCapabilities.Key]!;
        Assert.NotEqual(key, renewed);
        Assert.Null(await HostAdmin.AuthenticateAsync(t.Db, key));
        Assert.NotNull(await HostAdmin.AuthenticateAsync(t.Db, renewed));
        Assert.Equal(0, await t.Db.HostSessions.CountAsync(s => s.HostClientId == driveId));

        // the old host gets a new key through its listed id, and keeps its name
        var oldRenewed = await handler.BindAsync(Request(HostBindings.ManagedByHand, renew: true, consumer: old.Consumer, name: old.Name), default);
        Assert.Null(await HostAdmin.AuthenticateAsync(t.Db, legacyKey));
        Assert.Equal(legacy.Id, (await HostAdmin.AuthenticateAsync(t.Db, oldRenewed.Values[MatCapabilities.Key]))!.Id);
        Assert.Equal(2, (await handler.ListAsync(default)).Count);

        // only who made it removes it (any, without a name); then the key is refused
        Assert.False(await handler.UnbindAsync("matdrive", "someone", default));
        Assert.True(await handler.UnbindAsync("matdrive", "mattiverse", default));
        Assert.Null(await HostAdmin.AuthenticateAsync(t.Db, renewed));
        Assert.False(await handler.UnbindAsync("matdrive", null, default));
        Assert.True(await handler.UnbindAsync(old.Consumer, null, default));
        Assert.Empty(await handler.ListAsync(default));

        // bound again after it was removed: a new host with a new key
        var rebound = (await handler.BindAsync(Request("mattiverse", renew: false), default)).Values[MatCapabilities.Key];
        Assert.NotNull(await HostAdmin.AuthenticateAsync(t.Db, rebound));
    }

    [DbFact]
    public async Task A_host_key_opens_the_endpoints_of_office_edit_and_the_old_ones()
    {
        await using var t = await TestDatabase.CreateAsync();
        var data = Directory.CreateTempSubdirectory("matoffice-tests-");
        try
        {
            // not Development: without the settings below the app would not find a database, never the one of the local stack
            await using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(b => b
                .UseEnvironment("Testing")
                .UseSetting("ConnectionStrings:Default", t.Db.Database.GetConnectionString())
                .UseSetting("DataDirectory", data.FullName)
                .UseSetting(MatPlatformSettings.ApiKeyVariable, "test-api-key"));
            var client = factory.CreateClient(new WebApplicationFactoryClientOptions { AllowAutoRedirect = false });

            // the broker binds MatDrive through MatPlatform (/api/bindings with MAT_API_KEY) ...
            using var put = new HttpRequestMessage(HttpMethod.Put, "/api/bindings/office.edit/matdrive")
            {
                Content = new StringContent(JsonSerializer.Serialize(new { version = 1, managedBy = "mattiverse", consumer = new { name = "MatDrive" }, provider = new { url = "https://office.example.com" } }), Encoding.UTF8, "application/json"),
            };
            put.Headers.Add(MatApiAccess.HeaderName, "test-api-key");
            var bound = await client.SendAsync(put);
            Assert.Equal(HttpStatusCode.OK, bound.StatusCode);
            var variables = (await bound.Content.ReadFromJsonAsync<JsonElement>()).GetProperty("variables");
            Assert.Equal("https://office.example.com", variables.GetProperty("MAT_OFFICE_EDIT_URL").GetString());
            var key = variables.GetProperty("MAT_OFFICE_EDIT_KEY").GetString()!;
            Assert.StartsWith(HostAdmin.KeyPrefix, key, StringComparison.Ordinal);

            // ... and the key opens the endpoints of the capability, under the new and the old address (the /api group of MatPlatform does not stand in the way)
            foreach (var path in new[] { "/api/provides/office.edit/formats", "/api/v1/host/formats" })
            {
                using var formats = new HttpRequestMessage(HttpMethod.Get, path);
                formats.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);
                var answer = await client.SendAsync(formats);
                Assert.Equal(HttpStatusCode.OK, answer.StatusCode);
                var body = await answer.Content.ReadFromJsonAsync<JsonElement>();
                Assert.Equal("matoffice", body.GetProperty("app").GetString());
                Assert.Equal(4, body.GetProperty("kinds").GetArrayLength());

                using var wrong = new HttpRequestMessage(HttpMethod.Get, path);
                wrong.Headers.Authorization = new AuthenticationHeaderValue("Bearer", "mo_not-a-key");
                Assert.Equal(HttpStatusCode.Unauthorized, (await client.SendAsync(wrong)).StatusCode);
            }

            // the key of the app is not a host key
            using var apiKey = new HttpRequestMessage(HttpMethod.Get, "/api/provides/office.edit/formats");
            apiKey.Headers.Add(MatApiAccess.HeaderName, "test-api-key");
            Assert.Equal(HttpStatusCode.Unauthorized, (await client.SendAsync(apiKey)).StatusCode);

            // the broker sees the binding, and it is listed in /api/connections
            using var list = new HttpRequestMessage(HttpMethod.Get, "/api/bindings");
            list.Headers.Add(MatApiAccess.HeaderName, "test-api-key");
            var bindings = (await (await client.SendAsync(list)).Content.ReadFromJsonAsync<JsonElement>()).GetProperty("bindings");
            Assert.Equal("matdrive", Assert.Single(bindings.EnumerateArray()).GetProperty("consumer").GetString());

            // the old admin pages lead to the Connections area
            foreach (var path in new[] { "/Admin/Hosts", "/Admin/Hosts/Edit" })
            {
                var hosts = await client.GetAsync(path);
                Assert.Equal(HttpStatusCode.MovedPermanently, hosts.StatusCode);
                Assert.Equal("/Admin/Connections", hosts.Headers.Location?.OriginalString);
            }
        }
        finally
        {
            data.Delete(recursive: true);
        }
    }

    private static MatBindRequest Request(string managedBy, bool renew, string consumer = "matdrive", string name = "MatDrive")
        => new(MatCapabilities.OfficeEdit, 1, new MatBindingParty(consumer, name), Office, new Dictionary<string, string>(), managedBy, renew);

    private static string Root()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "MatOffice.slnx")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName ?? throw new InvalidOperationException("MatOffice.slnx not found.");
    }
}
