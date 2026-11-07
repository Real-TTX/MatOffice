using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>The content check, the templates and the file formats of Docs.</summary>
public sealed class DocsFormatTests
{
    internal const string Png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static readonly DocsHandler Handler = new();

    /// <summary>A document with a bit of everything Docs knows.</summary>
    internal static string Rich()
    {
        var body = "<h1 id=\"h-intro\">Introduction</h1>"
            + "<p data-style=\"Title\">The title</p><p data-style=\"Subtitle\">A subtitle</p>"
            + "<p style=\"text-align: justify; line-height: 1.8; margin-top: 6pt; margin-left: 18pt; text-indent: -9pt\">Normal <strong>bold</strong> <em>italic</em> <u>under</u> <s>strike</s> x<sup>2</sup> H<sub>2</sub>O "
            + "<span style=\"font-family: Arial, sans-serif; font-size: 14pt; color: #ff0000; background-color: #ffff00\">styled</span> "
            + "<span style=\"background-color: #ddeeff\">shaded</span> <a href=\"https://example.com/a?b=1&amp;c=2\">a link</a> <a href=\"#h-intro\">back</a>"
            + "<span data-comment=\"c1\">commented</span> and a note<sup data-fn=\"n1\"></sup>.\tTab<br>Line two</p>"
            + "<p data-list=\"bullet\">One</p><p data-list=\"bullet\" data-level=\"1\">One.one</p><p data-list=\"bullet\">Two</p>"
            + "<p>between</p>"
            + "<p data-list=\"decimal\" data-start=\"3\">Three</p><p data-list=\"decimal\" data-level=\"1\">Three.a</p><p data-list=\"decimal\">Four</p>"
            + "<p data-list=\"outline\">O1</p><p data-list=\"outline\" data-level=\"1\">O1.1</p><p data-list=\"check\">done</p><p data-list=\"upper-roman\">R</p>"
            + "<blockquote>A quote</blockquote><blockquote data-style=\"IntenseQuote\">Intense</blockquote><pre>code\n  indented</pre><hr>"
            + "<h2 id=\"h-table\">Table</h2>"
            + "<table data-table-style=\"accent1\" style=\"width: 100%\"><colgroup><col style=\"width: 30%\"><col style=\"width: 40%\"><col style=\"width: 30%\"></colgroup>"
            + "<thead><tr><th><p>A</p></th><th><p>B</p></th><th><p>C</p></th></tr></thead>"
            + "<tbody><tr><td rowspan=\"2\"><p>merged down</p></td><td colspan=\"2\" style=\"background-color: #ff0000; vertical-align: middle\"><p>merged right</p></td></tr>"
            + "<tr><td><p>b2</p></td><td><p>c2</p></td></tr>"
            + "<tr><td><p>a3</p></td><td><p>b3</p></td><td style=\"border-bottom: 1.5pt double #00ff00\"><p>c3</p></td></tr></tbody></table>"
            + "<p><img src=\"" + Png + "\" alt=\"A dot\" style=\"width: 40px; height: 20px\"> inline and <img src=\"" + Png + "\" alt=\"floating\" style=\"width: 30px; height: 30px\" data-align=\"right\"></p>"
            + "<p><img src=\"" + Png + "\" alt=\"centred\" style=\"width: 50px; height: 50px\" data-align=\"center\"></p>"
            + "<div data-break=\"page\"></div><p data-style=\"TOC1\"><a href=\"#h-table\">Table</a></p><p data-style=\"Caption\">Figure 1</p>";
        var content = new JsonObject
        {
            ["page"] = new JsonObject
            {
                ["size"] = "Letter", ["orientation"] = "landscape", ["margins"] = new JsonObject { ["top"] = 20, ["right"] = 15, ["bottom"] = 18, ["left"] = 30 },
                ["header"] = 10, ["footer"] = 11, ["columns"] = 2, ["differentFirst"] = true,
            },
            ["html"] = body,
            ["header"] = "<p style=\"text-align: right\">Header <span data-field=\"title\"></span></p>",
            ["footer"] = "<p style=\"text-align: center\">Page <span data-field=\"page\"></span> of <span data-field=\"pages\"></span></p>",
            ["firstHeader"] = "",
            ["firstFooter"] = "<p><span data-field=\"date\"></span></p>",
            ["notes"] = new JsonArray(new JsonObject { ["id"] = "n1", ["html"] = "<p>The <em>note</em> text</p>" }),
            ["comments"] = new JsonArray(new JsonObject
            {
                ["id"] = "c1", ["author"] = "Ada Lovelace", ["date"] = "2026-11-07T10:00:00Z", ["text"] = "Check this\nplease", ["done"] = false,
                ["replies"] = new JsonArray(new JsonObject { ["id"] = "r1", ["author"] = "Bob", ["date"] = "2026-11-07T11:00:00Z", ["text"] = "Done" }),
            }),
            ["settings"] = new JsonObject { ["language"] = "de-DE", ["spellcheck"] = false },
        };
        var check = Handler.Normalize(content.ToJsonString());
        Assert.True(check.Ok, check.Error);
        return check.Json!;
    }

    internal static JsonObject Parse(string json) => (JsonObject)JsonNode.Parse(json)!;

    internal static string Body(string json) => Parse(json)["html"]!.GetValue<string>();

    private static string RoundTrip(string content, string extension)
    {
        var bytes = Handler.Export(content, extension, "Round trip");
        var result = Kinds.Read(Handler, new MemoryStream(bytes), extension);
        Assert.True(result.Ok, result.Error);
        return result.Json!;
    }

    // ------------------------------------------------------------------ docx export

    [Fact]
    public void The_exported_docx_is_valid()
    {
        var bytes = Handler.Export(Rich(), ".docx", "Valid");
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    [Fact]
    public void An_empty_document_exports_to_a_valid_docx()
    {
        var empty = Handler.Normalize("{}");
        var bytes = Handler.Export(empty.Json!, ".docx", "Empty");
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
    }
}
