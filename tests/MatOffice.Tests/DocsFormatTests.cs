using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>The file formats of Docs: what each one keeps of a document on the way out and back in, and how foreign files are read.</summary>
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

    internal static string Field(string json, string name) => Parse(json)[name]!.GetValue<string>();

    /// <summary>Content as a file of a format and back (normalized like an import).</summary>
    internal static string RoundTrip(string content, string extension)
    {
        var bytes = Handler.Export(content, extension, "Round trip");
        var result = Kinds.Read(Handler, new MemoryStream(bytes), extension);
        Assert.True(result.Ok, result.Error);
        return result.Json!;
    }

    /// <summary>Normalized content of a file.</summary>
    internal static string Import(byte[] bytes, string extension)
    {
        var result = Kinds.Read(Handler, new MemoryStream(bytes), extension);
        Assert.True(result.Ok, result.Error);
        return result.Json!;
    }

    internal static string Import(string text, string extension) => Import(Encoding.UTF8.GetBytes(text), extension);

    /// <summary>Equal HTML, or a message that shows where the two differ.</summary>
    internal static void AssertSameHtml(string expected, string actual)
    {
        if (expected == actual)
        {
            return;
        }

        var i = 0;
        while (i < expected.Length && i < actual.Length && expected[i] == actual[i])
        {
            i++;
        }

        var from = Math.Max(0, i - 80);
        Assert.Fail($"The HTML differs at {i}:\n expected: …{expected[from..Math.Min(expected.Length, i + 120)]}\n actual:   …{actual[from..Math.Min(actual.Length, i + 120)]}");
    }

    private static void AssertValid(byte[] docx)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    // ------------------------------------------------------------------ docx

    [Fact]
    public void The_exported_docx_is_valid() => AssertValid(Handler.Export(Rich(), ".docx", "Valid"));

    [Fact]
    public void An_empty_document_exports_to_a_valid_docx() => AssertValid(Handler.Export(Handler.Normalize("{}").Json!, ".docx", "Empty"));

    [Fact]
    public void Docx_keeps_the_whole_document()
    {
        var rich = Rich();
        var back = RoundTrip(rich, ".docx");

        // Word knows comments by numbers, a font by its name, and a centred picture as a centred paragraph.
        var expected = Body(rich)
            .Replace("data-comment=\"c1\"", "data-comment=\"c0\"", StringComparison.Ordinal)
            .Replace("font-family: Arial, sans-serif", "font-family: Arial, Liberation Sans, sans-serif", StringComparison.Ordinal)
            .Replace($"<p><img src=\"{Png}\" alt=\"centred\" style=\"width: 50px; height: 50px\" data-align=\"center\"></p>", $"<p style=\"text-align: center\"><img src=\"{Png}\" alt=\"centred\" style=\"width: 50px; height: 50px\"></p>", StringComparison.Ordinal);
        AssertSameHtml(expected, Body(back));

        var original = Parse(rich);
        var result = Parse(back);
        foreach (var field in new[] { "page", "header", "footer", "firstHeader", "firstFooter", "notes", "settings" })
        {
            Assert.True(JsonNode.DeepEquals(original[field], result[field]), $"{field}: {original[field]?.ToJsonString()} became {result[field]?.ToJsonString()}");
        }

        Assert.Equal(
            "[{\"id\":\"c0\",\"author\":\"Ada Lovelace\",\"date\":\"2026-11-07T10:00:00Z\",\"text\":\"Check this\\nplease\",\"done\":false,"
            + "\"replies\":[{\"id\":\"c1\",\"author\":\"Bob\",\"date\":\"2026-11-07T11:00:00Z\",\"text\":\"Done\"}]}]",
            result["comments"]!.ToJsonString());
    }

    [Fact]
    public void Docx_keeps_every_table_style_and_list_kind()
    {
        var styles = string.Concat(DocsStyles.TableStyles.Order(StringComparer.Ordinal).Select(style =>
            $"<table data-table-style=\"{style}\" style=\"width: 100%\"><colgroup><col style=\"width: 35%\"><col style=\"width: 65%\"></colgroup><thead><tr><th><p>{style}</p></th><th><p>b</p></th></tr></thead>"
            + "<tbody><tr><td><p>1</p></td><td><p>2</p></td></tr><tr><td><p>3</p></td><td><p>4</p></td></tr><tr><td><p>5</p></td><td style=\"background-color: #123456\"><p>6</p></td></tr></tbody></table><p>-</p>"));
        var kinds = string.Concat(DocsStyles.BulletKinds.Concat(DocsStyles.NumberKinds).Order(StringComparer.Ordinal).Select(kind =>
            $"<p data-list=\"{kind}\">{kind}</p><p data-list=\"{kind}\" data-level=\"1\">{kind} 2</p><p data-list=\"{kind}\" data-level=\"2\">{kind} 3</p>"));
        var html = styles + kinds + "<p data-list=\"decimal\">a</p><p data-list=\"bullet\" data-level=\"1\">b</p><p data-list=\"decimal\" data-start=\"7\">c</p>";
        var content = Handler.Normalize(new JsonObject { ["html"] = html }.ToJsonString()).Json!;
        var exported = Handler.Export(content, ".docx", "Kinds");
        AssertValid(exported);
        AssertSameHtml(Body(content), Body(Import(exported, ".docx")));
    }

    [Fact]
    public void Docx_numbers_lists_the_way_the_editor_counts_them()
    {
        const string html = "<p data-list=\"decimal\">one</p><p data-list=\"decimal\">two</p><p data-list=\"bullet\">dot</p><p data-list=\"decimal\">one again</p>"
            + "<p data-list=\"decimal\" data-level=\"1\">a</p><p data-list=\"decimal\" data-level=\"1\">b</p><p data-list=\"decimal\">two again</p><p data-list=\"decimal\" data-level=\"1\">a again</p>"
            + "<p>other</p><p data-list=\"decimal\">one after a paragraph</p><p data-list=\"decimal\" data-start=\"10\">ten</p><p data-list=\"decimal\">eleven</p>";
        var content = Handler.Normalize(new JsonObject { ["html"] = html }.ToJsonString()).Json!;
        using var package = WordprocessingDocument.Open(new MemoryStream(Handler.Export(content, ".docx", "Numbers")), false);
        var numbering = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var paragraphs = package.MainDocumentPart.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();

        // Replays Word's counting: a numbering instance counts per level, its start values come from its overrides.
        var counters = new Dictionary<int, int[]>();
        var shown = new List<string>();
        foreach (var p in paragraphs)
        {
            var numPr = p.ParagraphProperties?.NumberingProperties;
            if (numPr is null)
            {
                continue;
            }

            var numId = numPr.NumberingId!.Val!.Value;
            var level = numPr.NumberingLevelReference!.Val!.Value;
            var num = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.NumberingInstance>().Single(n => n.NumberID!.Value == numId);
            var start = num.Elements<DocumentFormat.OpenXml.Wordprocessing.LevelOverride>().Single(o => o.LevelIndex!.Value == level).StartOverrideNumberingValue!.Val!.Value;
            var abstractNum = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.AbstractNum>().Single(a => a.AbstractNumberId!.Value == num.AbstractNumId!.Val!.Value);
            var format = abstractNum.Elements<DocumentFormat.OpenXml.Wordprocessing.Level>().Single(l => l.LevelIndex!.Value == level).NumberingFormat!.Val!.InnerText;
            if (!counters.TryGetValue(numId, out var c))
            {
                counters[numId] = c = new int[9];
            }

            c[level] = c[level] == 0 ? start : c[level] + 1;
            for (var deeper = level + 1; deeper < 9; deeper++)
            {
                c[deeper] = 0;
            }

            shown.Add(format == "bullet" ? "•" : format == "lowerLetter" ? DocsStyles.Alpha(c[level]) : c[level].ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        Assert.Equal(["1", "2", "•", "1", "a", "b", "2", "a", "1", "10", "11"], shown);
    }

    [Fact]
    public void Docx_lists_continued_after_a_paragraph_keep_their_numbers()
    {
        var html = "<p data-list=\"decimal\">one</p><p>break</p><p data-list=\"decimal\" data-start=\"2\">two</p><p data-list=\"upper-alpha\" data-start=\"4\">D</p>";
        var content = Handler.Normalize(new JsonObject { ["html"] = html }.ToJsonString()).Json!;
        AssertSameHtml(Body(content), Body(RoundTrip(content, ".docx")));
    }

    [Fact]
    public void Docx_keeps_a_custom_page_and_the_first_page_footer()
    {
        var content = Handler.Normalize(new JsonObject
        {
            ["page"] = new JsonObject { ["size"] = "Custom", ["width"] = 120, ["height"] = 180, ["margins"] = new JsonObject { ["top"] = 10, ["right"] = 12.5, ["bottom"] = 14, ["left"] = 16 }, ["differentFirst"] = true },
            ["html"] = "<p>a</p>",
            ["firstFooter"] = "<p>first <span data-field=\"page\"></span></p>",
        }.ToJsonString()).Json!;
        var back = Parse(RoundTrip(content, ".docx"));
        Assert.Equal("{\"size\":\"Custom\",\"width\":120,\"height\":180,\"orientation\":\"portrait\",\"margins\":{\"top\":10,\"right\":12.5,\"bottom\":14,\"left\":16},\"header\":12.5,\"footer\":12.5,\"columns\":1,\"differentFirst\":true}", back["page"]!.ToJsonString());
        Assert.Equal("<p>first <span data-field=\"page\"></span></p>", back["firstFooter"]!.GetValue<string>());
        Assert.Equal("", back["footer"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("")]
    [InlineData("not a zip at all")]
    [InlineData("PK\u0003\u0004 broken")]
    public void A_file_that_is_not_a_word_document_is_refused(string text)
    {
        var result = Kinds.Read(Handler, new MemoryStream(Encoding.Latin1.GetBytes(text)), ".docx");
        Assert.False(result.Ok);
        Assert.Equal(Text.T("Docs.Import.Damaged"), result.Error);
    }

    [Fact]
    public void A_zip_without_a_document_is_refused()
    {
        using var stream = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("hello.txt").Open());
            writer.Write("hello");
        }

        var result = Kinds.Read(Handler, new MemoryStream(stream.ToArray()), ".docx");
        Assert.False(result.Ok);
        Assert.Equal(Text.T("Docs.Import.Damaged"), result.Error);
    }

    [Fact]
    public void Odd_html_never_breaks_the_writer()
    {
        var html = "<p><img src=\"data:image/png;base64,AAAA\" alt=\"not a picture\"><img src=\"data:image/webp;base64,UklGRhYAAABXRUJQVlA4TAoAAAAvAAAAAEX/I/of\" alt=\"webp\"></p>"
            + "<table><tr><td colspan=\"5\" rowspan=\"9\"><p>wide</p></td></tr><tr><td><p>x</p></td><td rowspan=\"3\"><p>y</p></td></tr></table>"
            + "<p data-list=\"decimal\" data-level=\"8\" data-start=\"99999\">deep</p><p data-list=\"outline\" data-level=\"5\">outline</p>"
            + "<p>\u0001 control ￾ and <a href=\"#nowhere\">a link to nothing</a><sup data-fn=\"missing\"></sup><span data-comment=\"missing\">x</span></p>"
            + "<h6 id=\"x-1\" data-list=\"bullet\">heading in a list</h6><pre></pre><blockquote></blockquote>";
        var content = Handler.Normalize(new JsonObject { ["html"] = html, ["header"] = "<table><tr><td><p><span data-field=\"pages\"></span></p></td></tr></table>" }.ToJsonString()).Json!;
        var docx = Handler.Export(content, ".docx", "Odd \u0007 name");
        AssertValid(docx);
        Assert.Contains("wide", Handler.PlainText(Import(docx, ".docx"), 10_000), StringComparison.Ordinal);
        foreach (var extension in new[] { ".html", ".md", ".txt" })
        {
            Assert.NotEmpty(Handler.Export(content, extension, "Odd"));
        }
    }

    [Fact]
    public void An_unknown_format_is_refused()
    {
        Assert.Throws<MatOffice.Services.Formats.FormatException>(() => Handler.Export(Rich(), ".pdf", "x"));
        Assert.Throws<MatOffice.Services.Formats.FormatException>(() => Handler.Import(new MemoryStream([1, 2, 3]), ".rtf"));
        Assert.Equal([".docx", ".html", ".md", ".txt"], Handler.Formats.Select(f => f.Extension));
        Assert.All(Handler.Formats, f => Assert.True(f.CanImport && f.CanExport));
    }

    // ------------------------------------------------------------------ html

    [Fact]
    public void Html_keeps_the_body_the_page_and_the_notes()
    {
        var rich = Rich();
        var back = RoundTrip(rich, ".html");

        // An HTML page has no comments; headers and footers stay in Docs.
        AssertSameHtml(Body(rich).Replace("<span data-comment=\"c1\">commented</span>", "commented", StringComparison.Ordinal), Body(back));
        var original = Parse(rich);
        var result = Parse(back);
        Assert.True(JsonNode.DeepEquals(original["notes"], result["notes"]), result["notes"]!.ToJsonString());
        Assert.True(JsonNode.DeepEquals(original["settings"], result["settings"]), result["settings"]!.ToJsonString());
        Assert.True(JsonNode.DeepEquals(original["page"]!["margins"], result["page"]!["margins"]));
        Assert.Equal("Letter", result["page"]!["size"]!.GetValue<string>());
        Assert.Equal("landscape", result["page"]!["orientation"]!.GetValue<string>());
        Assert.Empty(result["comments"]!.AsArray());
    }

    [Fact]
    public void The_html_export_is_a_standalone_page_with_inline_styles()
    {
        var html = Encoding.UTF8.GetString(Handler.Export(Rich(), ".html", "A <title> & more"));
        Assert.StartsWith("<!DOCTYPE html>\n<html lang=\"de-DE\" spellcheck=\"false\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta charset=\"utf-8\">", html, StringComparison.Ordinal);
        Assert.Contains("<title>A &lt;title&gt; &amp; more</title>", html, StringComparison.Ordinal);
        Assert.Contains("@page { size: 279.4mm 215.9mm; margin: 20mm 15mm 18mm 30mm; }", html, StringComparison.Ordinal);
        Assert.Contains("<ol start=\"3\" style=\"list-style-type: decimal;", html, StringComparison.Ordinal);
        Assert.Contains("<ul style=\"list-style-type: circle;", html, StringComparison.Ordinal);
        Assert.Contains("background-color: #4472c4", html, StringComparison.Ordinal);
        Assert.Contains("<li id=\"fn-n1\">", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-style", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-comment", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Html_of_word_becomes_styles_lists_and_notes()
    {
        const string html = """
            <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
            <head><meta name=Generator content="Microsoft Word 15">
            <style>@page WordSection1 {size:595.3pt 841.9pt; margin:70.85pt 70.85pt 2.0cm 70.85pt;}</style></head>
            <body lang=DE><div class=WordSection1>
            <p class=MsoTitle>Titel</p>
            <p class=MsoListParagraphCxSpFirst style='text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span
            style='font-family:Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;
            </span></span></span><![endif]>Eins</p>
            <p class=MsoListParagraphCxSpLast style='margin-left:72.0pt;text-indent:-18.0pt;mso-list:l0 level2 lfo1'><![if !supportLists]><span
            style='font-family:"Courier New"'><span style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;
            </span></span></span><![endif]>Zwei</p>
            <p class=MsoListParagraph style='text-indent:-18.0pt;mso-list:l1 level1 lfo2'><![if !supportLists]><span
            style='mso-list:Ignore'>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span><![endif]>Drei</p>
            <p class=MsoNormal style='line-height:150%'>Text mit <b>fett</b> und <i>kursiv</i>.<a style='mso-footnote-id:ftn1' href="#_ftn1"
            name="_ftnref1" title=""><span class=MsoFootnoteReference><span style='mso-special-character:footnote'><![if !supportFootnotes]>[1]<![endif]></span></span></a></p>
            <p class=MsoQuote>Zitat</p>
            <span style='font-size:11.0pt;line-height:107%;font-family:"Calibri",sans-serif'><br clear=all style='page-break-before:always'></span>
            <p class=MsoNormal align=center style='text-align:center'><span style='font-size:14.0pt;color:red'>Mitte</span></p>
            </div>
            <div style='mso-element:footnote-list'><![if !supportFootnotes]><br clear=all><hr align=left size=1 width="33%"><![endif]>
            <div style='mso-element:footnote' id=ftn1><p class=MsoFootnoteText><a style='mso-footnote-id:ftn1' href="#_ftnref1" name="_ftn1"
            title=""><span class=MsoFootnoteReference>[1]</span></a> Die Fußnote.</p></div></div>
            </body></html>
            """;
        var content = Parse(Import(html, ".html"));
        AssertSameHtml(
            "<p data-style=\"Title\">Titel</p><p data-list=\"bullet\">Eins</p><p data-list=\"bullet\" data-level=\"1\">Zwei</p><p data-list=\"decimal\" data-start=\"3\">Drei</p>"
            + "<p style=\"line-height: 1.8\">Text mit <strong>fett</strong> und <em>kursiv</em>.<sup data-fn=\"ftn1\"></sup></p><blockquote>Zitat</blockquote>"
            + "<div data-break=\"page\"></div><p style=\"text-align: center\"><span style=\"font-size: 14pt; color: #ff0000\">Mitte</span></p>",
            content["html"]!.GetValue<string>());
        Assert.Equal("[{\"id\":\"ftn1\",\"html\":\"<p>Die Fußnote.</p>\"}]", content["notes"]!.ToJsonString(new System.Text.Json.JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
        Assert.Equal("A4", content["page"]!["size"]!.GetValue<string>());
        Assert.Equal(20, content["page"]!["margins"]!["bottom"]!.GetValue<double>());
        Assert.Equal(25, content["page"]!["margins"]!["left"]!.GetValue<double>());
    }

    [Fact]
    public void Html_of_google_docs_keeps_bold_spans_and_lists()
    {
        const string html = "<meta charset=\"utf-8\"><b style=\"font-weight:normal;\" id=\"docs-internal-guid-1234\"><p dir=\"ltr\" style=\"line-height:1.38;margin-top:0pt;margin-bottom:0pt;\">"
            + "<span style=\"font-size:11pt;font-family:Arial,sans-serif;color:#000000;background-color:transparent;font-weight:700;font-style:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;\">Bold</span>"
            + "<span style=\"font-size:16px;font-family:Arial,sans-serif;color:#000000;background-color:transparent;font-weight:400;font-style:italic;white-space:pre-wrap;\"> and italic</span></p>"
            + "<ul style=\"margin-top:0;margin-bottom:0;padding-inline-start:48px;\"><li dir=\"ltr\" style=\"list-style-type:disc;font-size:11pt;\" aria-level=\"1\">"
            + "<p dir=\"ltr\" style=\"line-height:1.38;margin-top:0pt;margin-bottom:0pt;\" role=\"presentation\"><span style=\"font-size:11pt;\">Item</span></p></li>"
            + "<ul style=\"margin-top:0;margin-bottom:0;padding-inline-start:48px;\"><li dir=\"ltr\" style=\"list-style-type:circle;font-size:11pt;\" aria-level=\"2\">"
            + "<p dir=\"ltr\" style=\"line-height:1.38;margin-top:0pt;margin-bottom:0pt;\" role=\"presentation\"><span style=\"font-size:11pt;\">Nested</span></p></li></ul></ul></b>";
        AssertSameHtml(
            "<p style=\"line-height: 1.38; margin-bottom: 0pt\"><span style=\"font-family: Arial,sans-serif\"><strong>Bold</strong></span><span style=\"font-family: Arial,sans-serif; font-size: 12pt\"><em> and italic</em></span></p>"
            + "<p data-list=\"bullet\" style=\"line-height: 1.38\">Item</p><p data-list=\"bullet\" data-level=\"1\" style=\"line-height: 1.38\">Nested</p>",
            Body(Import(html, ".html")));
    }

    [Fact]
    public void Html_links_to_anchors_keep_working_and_other_ids_go()
    {
        const string html = "<p class=MsoToc1><a href=\"#_Toc123\">Intro</a></p><h1><a name=\"_Toc123\"></a>Intro</h1><p id=\"unused\">x</p><div id=\"wrapper\"><p>y</p></div>"
            + "<h2 id=\"h-two\">Two</h2><p><a href=\"#h-two\">to two</a> <a href=\"#missing\">nowhere</a></p>";
        Assert.Equal(
            "<p data-style=\"TOC1\"><a href=\"#x_Toc123\">Intro</a></p><h1 id=\"x_Toc123\">Intro</h1><p>x</p><p>y</p><h2 id=\"h-two\">Two</h2>"
            + "<p><a href=\"#h-two\">to two</a> <a href=\"#missing\">nowhere</a></p>",
            Body(Import(html, ".html")));
    }

    [Fact]
    public void Html_import_reads_the_charset_and_drops_what_is_not_allowed()
    {
        var bytes = Encoding.Latin1.GetBytes("<html><head><meta charset=\"windows-1252\"><title>t</title><script>alert(1)</script></head><body onload=\"x()\">"
            + "<p>Grüße <img src=\"https://example.com/x.png\"><a href=\"page.html\">relative</a> <a href=\"https://example.com/\">absolute</a></p>"
            + "<iframe src=\"https://example.com\"></iframe><form><input value=\"x\"></form><p style=\"color: expression(alert(1)); font-size: 16px\">big</p></body></html>");
        AssertSameHtml("<p>Grüße relative <a href=\"https://example.com/\">absolute</a></p><p><span style=\"font-size: 12pt\">big</span></p>", Body(Import(bytes, ".html")));

        var utf16 = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("<p>€ 5</p>")).ToArray();
        Assert.Equal("<p>€ 5</p>", Body(Import(utf16, ".html")));
        Assert.Equal("<p>€ 5</p>", Body(Import(Encoding.Latin1.GetBytes("<p>\u0080 5</p>"), ".htm")));
    }

    [Fact]
    public void Html_import_collapses_the_whitespace_of_the_source_but_not_in_pre()
    {
        const string html = "<body>\n  <p>\n    one\n    two  three\n  </p>\n  <pre>  keep\n    this</pre>\n  <ul>\n    <li>\n      item\n    </li>\n  </ul>\n</body>";
        Assert.Equal("<p>one two three</p><pre>  keep\n    this</pre><p data-list=\"bullet\">item</p>", Body(Import(html, ".html")));
    }

    // ------------------------------------------------------------------ markdown

    private const string MarkdownBody = "<h1>Heading</h1><h3>Sub</h3>"
        + "<p>Text <strong>bold</strong> <em>italic</em> <strong><em>both</em></strong> <s>strike</s> <u>under</u> x<sup>2</sup> H<sub>2</sub>O <a href=\"https://example.com/\">link</a> "
        + "<a href=\"mailto:a@example.com\">mail</a> and a note<sup data-fn=\"n1\"></sup>.<br>Next line with * star, _under_, [brackets], a | pipe, a # hash, `ticks` and 1. a number</p>"
        + "<p>#not a heading</p><p>- not a list</p><p>2. not a list either</p>"
        + "<p data-list=\"bullet\">a</p><p data-list=\"bullet\" data-level=\"1\">b</p><p data-list=\"bullet\" data-level=\"2\">c</p><p data-list=\"bullet\">d</p>"
        + "<p>between</p><p data-list=\"decimal\" data-start=\"3\">three</p><p data-list=\"decimal\">four</p><p data-list=\"decimal\" data-level=\"1\">four.one</p>"
        + "<blockquote>quote</blockquote><pre>code\n  more ``` fences</pre><hr>"
        + "<table data-table-style=\"grid\" style=\"width: 100%\"><thead><tr><th><p>A</p></th><th><p>B | C</p></th></tr></thead><tbody><tr><td><p>1</p></td><td><p><strong>2</strong></p></td></tr></tbody></table>"
        + "<p><img src=\"" + Png + "\" alt=\"dot\"></p>";

    [Fact]
    public void Markdown_keeps_what_markdown_can_say()
    {
        var content = Handler.Normalize(new JsonObject
        {
            ["html"] = MarkdownBody,
            ["notes"] = new JsonArray(new JsonObject { ["id"] = "n1", ["html"] = "<p>The <em>note</em></p>" }),
        }.ToJsonString()).Json!;
        var back = Parse(RoundTrip(content, ".md"));

        // Notes are numbered in Markdown ([^1]).
        AssertSameHtml(MarkdownBody.Replace("data-fn=\"n1\"", "data-fn=\"1\"", StringComparison.Ordinal), back["html"]!.GetValue<string>());
        Assert.Equal("[{\"id\":\"1\",\"html\":\"<p>The <em>note</em></p>\"}]", back["notes"]!.ToJsonString(new System.Text.Json.JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
    }

    [Fact]
    public void Markdown_export_writes_common_markdown()
    {
        var markdown = Encoding.UTF8.GetString(Handler.Export(Rich(), ".md", "x"));
        Assert.StartsWith("# Introduction\n\n# The title\n\n*A subtitle*\n\n", markdown, StringComparison.Ordinal);
        Assert.Contains("[a link](https://example.com/a?b=1&c=2)", markdown, StringComparison.Ordinal);
        Assert.Contains("- One\n  - One.one\n- Two\n", markdown, StringComparison.Ordinal);
        Assert.Contains("3. Three\n   1. Three.a\n4. Four\n", markdown, StringComparison.Ordinal);
        Assert.Contains("```\ncode\n  indented\n```", markdown, StringComparison.Ordinal);
        Assert.Contains("| A | B | C |\n| --- | --- | --- |\n", markdown, StringComparison.Ordinal);
        Assert.Contains("[^1]: The *note* text", markdown, StringComparison.Ordinal);
        Assert.EndsWith("\n", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("\r", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_import_reads_the_usual_forms()
    {
        const string markdown = "Title\n=====\n\nSub\n---\n\n    indented code\n\n* star\n+ plus\n\n1) one\n2) two\n\n   more of two\n\n<https://example.com> and www.example.org and <a@example.com>\n\n"
            + "[ref link][r] and ![](data:image/gif;base64,R0lGODlhAQABAAAAACw=) and ![external](https://example.com/x.png)\n\n[r]: https://example.com/ref \"Title\"\n\n"
            + "Hard  \nbreak\\\nand &amp; &copy; entities, <mark>marked</mark>, ~one tilde~ and ~~two~~\n\n> quoted\nlazy line\n\n- [link *in* list](#x)\n\n***\n";
        AssertSameHtml(
            "<h1>Title</h1><h2>Sub</h2><pre>indented code</pre><p data-list=\"bullet\">star</p><p data-list=\"bullet\">plus</p><p data-list=\"decimal\">one</p><p data-list=\"decimal\">two</p>"
            + "<p data-style=\"ListParagraph\">more of two</p>"
            + "<p><a href=\"https://example.com\">https://example.com</a> and <a href=\"http://www.example.org\">www.example.org</a> and <a href=\"mailto:a@example.com\">a@example.com</a></p>"
            + "<p><a href=\"https://example.com/ref\">ref link</a> and <img src=\"data:image/gif;base64,R0lGODlhAQABAAAAACw=\" alt=\"\"> and</p>"
            + "<p>Hard<br>break<br>and &amp; © entities, <span style=\"background-color: #ffff00\">marked</span>, ~one tilde~ and <s>two</s></p>"
            + "<blockquote>quoted lazy line</blockquote><p data-list=\"bullet\">link <em>in</em> list</p><hr>",
            Body(Import(markdown, ".md")));
    }

    [Fact]
    public void Markdown_links_inside_the_document_point_to_the_anchors_of_the_headings()
    {
        var content = Handler.Normalize(new JsonObject
        {
            ["html"] = "<h1 id=\"h-intro\">Introduction</h1><p><a href=\"#h-intro\">back</a> and <a href=\"#h-x\">to the other</a></p><h2 id=\"h-x\">Ünïcode Tëxt!</h2>"
                + "<h2>Introduction</h2><p><a href=\"#h-nowhere\">nowhere</a></p>",
        }.ToJsonString()).Json!;
        var markdown = Encoding.UTF8.GetString(Handler.Export(content, ".md", "x"));
        Assert.Contains("[back](#introduction) and [to the other](#ünïcode-tëxt)", markdown, StringComparison.Ordinal);
        Assert.Equal(
            "<h1 id=\"introduction\">Introduction</h1><p><a href=\"#introduction\">back</a> and <a href=\"#x_n_code-t_xt\">to the other</a></p>"
            + "<h2 id=\"x_n_code-t_xt\">Ünïcode Tëxt!</h2><h2>Introduction</h2><p>nowhere</p>",
            Body(RoundTrip(content, ".md")));
        Assert.Equal("<h2>A</h2><h2 id=\"a-1\">A</h2><p><a href=\"#a-1\">second</a></p>", Body(Import("## A\n\n## A\n\n[second](#a-1)", ".md")));
    }

    // ------------------------------------------------------------------ text

    [Fact]
    public void Text_keeps_lines_and_tabs()
    {
        var content = Handler.Normalize("{\"html\":\"<p>a\\tb &lt;c&gt; &amp; d</p><p></p><p>last</p>\"}").Json!;
        var text = Encoding.UTF8.GetString(Handler.Export(content, ".txt", "x"));
        Assert.Equal("a\tb <c> & d\r\n\r\nlast\r\n", text);
        AssertSameHtml(Body(content), Body(RoundTrip(content, ".txt")));
    }

    [Fact]
    public void Text_export_writes_list_markers_cells_and_notes()
    {
        var text = Encoding.UTF8.GetString(Handler.Export(Rich(), ".txt", "x"));
        Assert.Contains("\r\n• One\r\n  ◦ One.one\r\n• Two\r\nbetween\r\n3. Three\r\n  a. Three.a\r\n4. Four\r\n1. O1\r\n  1.1. O1.1\r\n✓ done\r\nI. R\r\n", text, StringComparison.Ordinal);
        Assert.Contains("A\tB\tC\r\n", text, StringComparison.Ordinal);
        Assert.Contains("a note[1].\tTab\r\nLine two\r\n", text, StringComparison.Ordinal);
        Assert.EndsWith("[1] The note text\r\n", text, StringComparison.Ordinal);
        Assert.False(text.StartsWith('﻿'));
    }

    [Theory]
    [InlineData(new byte[] { 0xEF, 0xBB, 0xBF, 0x47, 0x72, 0xC3, 0xBC, 0xC3, 0x9F, 0x65 }, "Grüße")]
    [InlineData(new byte[] { 0x47, 0x72, 0xC3, 0xBC, 0xC3, 0x9F, 0x65 }, "Grüße")]
    [InlineData(new byte[] { 0x47, 0x72, 0xFC, 0xDF, 0x65, 0x20, 0x80 }, "Grüße €")]
    [InlineData(new byte[] { 0xFF, 0xFE, 0x47, 0x00, 0xFC, 0x00 }, "Gü")]
    [InlineData(new byte[] { 0xFE, 0xFF, 0x00, 0x47, 0x00, 0xFC }, "Gü")]
    public void Text_import_finds_the_encoding(byte[] bytes, string expected) => Assert.Equal($"<p>{expected}</p>", Body(Import(bytes, ".txt")));

    [Fact]
    public void Text_import_makes_a_paragraph_of_every_line()
    {
        Assert.Equal("<p>one</p><p></p><p>two</p><p>three</p>", Body(Import("one\r\n\r\ntwo\rthree\n", ".txt")));
        Assert.Equal("<p>page one</p><div data-break=\"page\"></div><p>page two</p>", Body(Import("page one\fpage two", ".txt")));
        Assert.Equal("<p></p>", Body(Import(string.Empty, ".txt")));
        Assert.Equal("<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>", Body(Import("<script>alert(1)</script>", ".txt")));
    }
}
