using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Data;

namespace MatOffice.Services.Formats;

/// <summary>A file format a kind can read (import, open from a host) and/or write (export, save back to a host).</summary>
public sealed record FileFormat(string Extension, string MimeType, string Label, bool CanImport, bool CanExport);

/// <summary>A template a new document can start from. <see cref="Name"/> and <see cref="Description"/> are keys of the text catalog.</summary>
public sealed record KindTemplate(string Id, string Name, string Description, string Icon = "file");

/// <summary>The answer of <see cref="IKindHandler.Normalize"/>: the checked content (canonical JSON) or what is wrong with it (worded for the person).</summary>
public sealed record ContentCheck(string? Json, string? Error)
{
    public bool Ok => Error is null && Json is not null;

    public static ContentCheck Valid(string json) => new(json, null);

    public static ContentCheck Invalid(string error) => new(null, error);
}

/// <summary>A file that could not be read or written; the message is worded for the person.</summary>
public sealed class FormatException(string message) : Exception(message);

/// <summary>
/// One app of the suite on the server: what a new document holds, the check of what the editor saves (the server never trusts the browser), and the
/// conversion from and to files. The editor in the browser (<c>wwwroot/js/&lt;slug&gt;/</c>) works on exactly the JSON this handler accepts.
/// </summary>
public interface IKindHandler
{
    DocumentKind Kind { get; }

    /// <summary>The address and the folder of the scripts: <c>docs</c>, <c>sheets</c>, <c>slides</c>, <c>base</c>.</summary>
    string Slug { get; }

    /// <summary>The colour of the app (title bar, icon), like the colours of the office suites people know.</summary>
    string Colour { get; }

    /// <summary>The icon of the sprite (<c>wwwroot/img/app-icons.svg</c>).</summary>
    string Icon { get; }

    /// <summary>The extension of the own file format (<c>.matdoc</c> ...): a JSON file with the content, see <see cref="NativeFile"/>.</summary>
    string NativeExtension { get; }

    /// <summary>The other formats (not the native one), in the order of the menu "Download as".</summary>
    IReadOnlyList<FileFormat> Formats { get; }

    /// <summary>The templates of the start page; the first one is the blank document.</summary>
    IReadOnlyList<KindTemplate> Templates { get; }

    /// <summary>The content of a new document from a template (the texts in the template in <paramref name="language"/>, en or de).</summary>
    string Create(string templateId, string language);

    /// <summary>Checks and cleans what the editor (or an import) delivers: the shape of the kind, limits, no markup that could run script.</summary>
    ContentCheck Normalize(string json);

    /// <summary>Reads a file of one of <see cref="Formats"/> into content (not normalized yet). Throws <see cref="FormatException"/>.</summary>
    string Import(Stream input, string extension);

    /// <summary>Writes the content as a file of one of <see cref="Formats"/>. Throws <see cref="FormatException"/>.</summary>
    byte[] Export(string content, string extension, string name);

    /// <summary>The text of the document (for the search and the description of a file), at most <paramref name="max"/> characters.</summary>
    string PlainText(string content, int max);
}

/// <summary>The four handlers and how to find them by kind, address and file extension.</summary>
public static class Kinds
{
    /// <summary>The largest content a document may have (pictures are inside it).</summary>
    public const int MaxContentBytes = 40 * 1024 * 1024;

    public static IReadOnlyList<IKindHandler> All { get; } =
    [
        new Docs.DocsHandler(),
        new Sheets.SheetsHandler(),
        new Slides.SlidesHandler(),
        new Base.BaseHandler(),
    ];

    public static IKindHandler Of(DocumentKind kind) => All.First(h => h.Kind == kind);

    public static IKindHandler? BySlug(string? slug) => All.FirstOrDefault(h => string.Equals(h.Slug, slug, StringComparison.OrdinalIgnoreCase));

    public static string SlugOf(DocumentKind kind) => Of(kind).Slug;

    /// <summary>The handler that reads files with this extension (native first), or null.</summary>
    public static (IKindHandler Handler, FileFormat Format)? ForExtension(string? extension)
    {
        var ext = NormalizeExtension(extension);
        foreach (var h in All)
        {
            if (h.NativeExtension == ext)
            {
                return (h, NativeFormat(h));
            }
        }

        foreach (var h in All)
        {
            var f = h.Formats.FirstOrDefault(f => f.Extension == ext && f.CanImport);
            if (f is not null)
            {
                return (h, f);
            }
        }

        return null;
    }

    /// <summary>The native format of a kind as a <see cref="FileFormat"/>.</summary>
    public static FileFormat NativeFormat(IKindHandler h) => new(h.NativeExtension, "application/json", "MatOffice", true, true);

    /// <summary>Every format a kind can write, the native one last.</summary>
    public static IEnumerable<FileFormat> ExportFormats(IKindHandler h) => h.Formats.Where(f => f.CanExport).Append(NativeFormat(h));

    public static string NormalizeExtension(string? extension)
    {
        var ext = (extension ?? string.Empty).Trim().ToLowerInvariant();
        if (ext.Length > 0 && ext[0] != '.')
        {
            ext = "." + ext;
        }

        return ext;
    }

    /// <summary>Reads a file (native or foreign) into normalized content.</summary>
    public static ContentCheck Read(IKindHandler h, Stream input, string extension)
    {
        var ext = NormalizeExtension(extension);
        try
        {
            var raw = ext == h.NativeExtension ? NativeFile.Read(input, h.Kind) : h.Import(input, ext);
            return h.Normalize(raw);
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (Exception ex) when (ex is JsonException or InvalidDataException or IOException or InvalidOperationException or ArgumentException or NotSupportedException)
        {
            return ContentCheck.Invalid(Localization.Text.T("Documents.Import.Unreadable"));
        }
    }

    /// <summary>Writes content as a file (native or foreign).</summary>
    public static byte[] Write(IKindHandler h, string content, string extension, string name)
    {
        var ext = NormalizeExtension(extension);
        return ext == h.NativeExtension ? NativeFile.Write(h.Kind, content, name) : h.Export(content, ext, name);
    }

    /// <summary>The MIME type of a format of a kind.</summary>
    public static string MimeOf(IKindHandler h, string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext == h.NativeExtension ? "application/json" : h.Formats.FirstOrDefault(f => f.Extension == ext)?.MimeType ?? "application/octet-stream";
    }

    /// <summary>A plain check every handler can start with: the content is one JSON object and not too large.</summary>
    public static JsonObject ParseObject(string json)
    {
        if (Encoding.UTF8.GetByteCount(json) > MaxContentBytes)
        {
            throw new FormatException(Localization.Text.T("Documents.TooLarge"));
        }

        return JsonNode.Parse(json, documentOptions: new JsonDocumentOptions { MaxDepth = 128 }) as JsonObject
               ?? throw new FormatException(Localization.Text.T("Documents.Invalid"));
    }
}

/// <summary>
/// The own file format of MatOffice: UTF-8 JSON <c>{ "format": "matoffice", "kind": "sheets", "version": 1, "name": "...", "content": { ... } }</c>.
/// It keeps everything the editor knows; the foreign formats keep what they can.
/// </summary>
public static class NativeFile
{
    public const string Marker = "matoffice";

    public static string Read(Stream input, DocumentKind kind)
    {
        var node = JsonNode.Parse(input, documentOptions: new JsonDocumentOptions { MaxDepth = 128 }) as JsonObject
                   ?? throw new FormatException(Localization.Text.T("Documents.Import.Unreadable"));
        if (node["format"]?.GetValue<string>() != Marker || node["content"] is not JsonObject content)
        {
            throw new FormatException(Localization.Text.T("Documents.Import.Unreadable"));
        }

        var fileKind = node["kind"]?.GetValue<string>();
        if (!string.Equals(fileKind, Kinds.SlugOf(kind), StringComparison.OrdinalIgnoreCase))
        {
            throw new FormatException(Localization.Text.T("Documents.Import.WrongKind"));
        }

        return content.ToJsonString();
    }

    public static byte[] Write(DocumentKind kind, string content, string name)
    {
        var node = new JsonObject
        {
            ["format"] = Marker,
            ["kind"] = Kinds.SlugOf(kind),
            ["version"] = 1,
            ["name"] = name,
            ["content"] = JsonNode.Parse(content),
        };
        return Encoding.UTF8.GetBytes(node.ToJsonString());
    }
}
