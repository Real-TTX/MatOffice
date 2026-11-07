using System.Text;
using System.Text.Json;
using MatOffice.Data;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// MatOffice Docs: text documents (like Word). The content (the full contract is <c>docs/EDITORS.md</c>, section "Docs content"):
/// <code>
/// {
///   "version": 1,
///   "page": { "size": "A4|A5|A3|Letter|Legal|Custom", "width": 210, "height": 297, "orientation": "portrait|landscape",
///             "margins": { "top": 25, "right": 25, "bottom": 20, "left": 25 }, "header": 12.5, "footer": 12.5, "columns": 1, "differentFirst": false },
///   "html": "&lt;h1 id=\"h-intro\"&gt;Introduction&lt;/h1&gt;&lt;p&gt;Text …&lt;/p&gt;",
///   "header": "", "footer": "", "firstHeader": "", "firstFooter": "",
///   "notes": [ { "id": "n1", "html": "&lt;p&gt;…&lt;/p&gt;" } ],
///   "comments": [ { "id": "c1", "author": "Ada", "date": "2026-11-07T10:00:00Z", "text": "…", "done": false, "replies": [ { "id", "author", "date", "text" } ] } ],
///   "settings": { "language": "en", "spellcheck": true }
/// }
/// </code>
/// Lengths of the page are millimetres (width/height = the portrait size; named sizes have their own; Custom 50–1000 mm; margins 0–150 mm with at least
/// 20 mm of text left; columns 1–3). Every HTML is <b>Docs HTML</b>: a flat list of blocks (<c>p</c> with <c>data-style</c>, <c>h1</c>–<c>h6</c>,
/// <c>blockquote</c>, <c>pre</c>, list items as <c>p data-list data-level data-start</c>, <c>table</c> with <c>data-table-style</c>, <c>hr</c>,
/// <c>div data-break</c>) with inline content (<c>strong em u s sub sup span[style] a img br</c>, <c>sup data-fn</c>, <c>span data-comment</c>,
/// <c>span data-field</c>), sanitized and canonicalized by <see cref="DocsHtml"/>. Notes need a <c>sup data-fn</c> in the body; comments may lose their anchor.
/// </summary>
public sealed class DocsHandler : IKindHandler
{
    public DocumentKind Kind => DocumentKind.Docs;

    public string Slug => "docs";

    public string Colour => "#185abd";

    public string Icon => "kind-docs";

    public string NativeExtension => ".matdoc";

    /// <summary>The foreign formats; the labels follow the language of the request.</summary>
    public IReadOnlyList<FileFormat> Formats =>
    [
        new(".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", Text.T("Docs.Format.Docx"), true, true),
        new(".html", "text/html", Text.T("Docs.Format.Html"), true, true),
        new(".md", "text/markdown", Text.T("Docs.Format.Markdown"), true, true),
        new(".txt", "text/plain", Text.T("Docs.Format.Text"), true, true),
    ];

    public IReadOnlyList<KindTemplate> Templates => DocsTemplates.All;

    public string Create(string templateId, string language) => DocsTemplates.Create(templateId, language);

    public ContentCheck Normalize(string json)
    {
        try
        {
            var result = DocsDocument.FromJson(json).ToJson();
            return Encoding.UTF8.GetByteCount(result) > Kinds.MaxContentBytes
                ? ContentCheck.Invalid(Text.T("Documents.TooLarge"))
                : ContentCheck.Valid(result);
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or ArgumentException)
        {
            return ContentCheck.Invalid(Text.T("Documents.Invalid"));
        }
    }

    public string Import(Stream input, string extension)
    {
        var ext = Kinds.NormalizeExtension(extension);
        var document = ext switch
        {
            ".docx" => DocxReader.Read(input),
            ".html" or ".htm" => HtmlFormat.Read(input),
            ".md" or ".markdown" => MarkdownFormat.Read(input),
            ".txt" => TextFormat.Read(input),
            _ => throw new FormatException(Text.T("Documents.Import.NotSupported")),
        };
        return document.ToJson();
    }

    public byte[] Export(string content, string extension, string name)
    {
        var ext = Kinds.NormalizeExtension(extension);
        var document = DocsDocument.FromJson(content);
        try
        {
            return ext switch
            {
                ".docx" => DocxWriter.Write(document, name),
                ".html" => HtmlFormat.Write(document, name),
                ".md" => MarkdownFormat.Write(document),
                ".txt" => TextFormat.Write(document),
                _ => throw new FormatException(Text.T("Documents.Export.NotSupported")),
            };
        }
        catch (Exception ex) when (ex is not FormatException and not OutOfMemoryException)
        {
            throw new FormatException(Text.T("Docs.Export.Failed"));
        }
    }

    public string PlainText(string content, int max)
    {
        if (max <= 0)
        {
            return string.Empty;
        }

        try
        {
            var document = DocsDocument.FromJson(content);
            var sb = new StringBuilder();
            DocsHtml.PlainText(sb, document.Body, max);
            var text = sb.ToString().TrimEnd();
            if (text.Length <= max)
            {
                return text;
            }

            // Never half a character (a surrogate pair) at the end.
            var cut = char.IsHighSurrogate(text[max - 1]) ? max - 1 : max;
            return text[..cut].TrimEnd();
        }
        catch (Exception ex) when (ex is FormatException or JsonException or InvalidOperationException or ArgumentException)
        {
            return string.Empty;
        }
    }
}
