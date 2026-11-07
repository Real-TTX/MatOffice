using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>The templates of Docs: canonical content in English and German, with the structure each one promises.</summary>
public sealed class DocsTemplateTests
{
    private static readonly DocsHandler Handler = new();

    private static readonly string[] HtmlFields = ["html", "header", "footer", "firstHeader", "firstFooter"];

    public static TheoryData<string, string> Templates()
    {
        var data = new TheoryData<string, string>();
        foreach (var template in Handler.Templates.Skip(1))
        {
            data.Add(template.Id, "en");
            data.Add(template.Id, "de");
        }

        return data;
    }

    [Fact]
    public void The_list_starts_with_the_blank_document_and_has_six_templates()
    {
        Assert.Equal(["blank", "letter", "report", "meeting", "resume", "proposal", "invoice"], Handler.Templates.Select(t => t.Id));
        Assert.Equal("kind-docs", Handler.Templates[0].Icon);
        Assert.All(Handler.Templates.Skip(1), t => Assert.StartsWith("Docs.Template.", t.Name, StringComparison.Ordinal));
        Assert.All(Handler.Templates.Skip(1), t => Assert.Equal(t.Name + ".Text", t.Description));
    }

    [Theory]
    [MemberData(nameof(Templates))]
    public void A_template_is_canonical_content(string id, string language)
    {
        var raw = DocsFormatTests.Parse(Handler.Create(id, language));
        var check = Handler.Normalize(raw.ToJsonString());
        Assert.True(check.Ok, check.Error);
        var normal = DocsFormatTests.Parse(check.Json!);
        foreach (var field in HtmlFields)
        {
            AssertSame(raw[field]!.GetValue<string>(), normal[field]!.GetValue<string>(), $"{id}/{language} {field}");
        }

        foreach (var field in new[] { "version", "page", "notes", "comments", "settings" })
        {
            Assert.True(JsonNode.DeepEquals(raw[field], normal[field]), $"{id}/{language} {field}: {raw[field]?.ToJsonString()} became {normal[field]?.ToJsonString()}");
        }

        Assert.Equal(normal.Select(p => p.Key).Order(StringComparer.Ordinal), raw.Select(p => p.Key).Order(StringComparer.Ordinal));
        Assert.Equal(check.Json, Handler.Normalize(check.Json!).Json);
    }

    [Theory]
    [MemberData(nameof(Templates))]
    public void A_template_is_in_its_language_and_its_links_work(string id, string language)
    {
        var document = Document(id, language);
        var other = Document(id, language == "de" ? "en" : "de");
        Assert.Equal(language == "de" ? "de-DE" : "en-GB", document.Language);
        Assert.NotEqual(DocsHtml.Write(document.Body), DocsHtml.Write(other.Body));
        Assert.True(DocsHtml.Write(document.Body).Length > 1000, "a template has real content");

        var ids = DocsHtml.Paragraphs(document.Body).Where(p => p.Id is not null).Select(p => p.Id!).ToHashSet(StringComparer.Ordinal);
        var links = DocsHtml.Paragraphs(document.Body).SelectMany(p => p.Inlines).Select(i => i.Format.Href).OfType<string>().ToList();
        Assert.All(links.Where(h => h.StartsWith('#')), h => Assert.Contains(h[1..], ids));
        Assert.All(document.Notes, n => Assert.Contains(DocsHtml.Paragraphs(document.Body).SelectMany(p => p.Inlines), i => i is DocNoteRef r && r.Id == n.Id));
    }

    [Theory]
    [MemberData(nameof(Templates))]
    public void A_template_exports_to_every_format(string id, string language)
    {
        var content = Handler.Normalize(Handler.Create(id, language)).Json!;
        var docx = Handler.Export(content, ".docx", id);
        using (var package = WordprocessingDocument.Open(new MemoryStream(docx), false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package).Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
            Assert.True(errors.Count == 0, string.Join("\n", errors));
        }

        foreach (var extension in new[] { ".docx", ".html", ".md", ".txt" })
        {
            var bytes = extension == ".docx" ? docx : Handler.Export(content, extension, id);
            var back = Kinds.Read(Handler, new MemoryStream(bytes), extension);
            Assert.True(back.Ok, $"{extension}: {back.Error}");
            Assert.Contains(FirstWords(content), Handler.PlainText(back.Json!, 100_000), StringComparison.Ordinal);
        }
    }

    [Fact]
    public void The_letter_has_a_letterhead_on_the_first_page_and_page_numbers_after_it()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var letter = Document("letter", language);
            Assert.True(letter.Page.DifferentFirst);
            Assert.Contains(language == "de" ? "Muster GmbH" : "Northwind Consulting", Text(letter.FirstHeader), StringComparison.Ordinal);
            Assert.Equal(["page", "pages"], Fields(letter.Footer));
            Assert.Contains(DocsHtml.Paragraphs(letter.Body), p => p.Inlines.Count == 1 && p.Inlines[0].Format.Bold);
            Assert.Contains(language == "de" ? "Mit freundlichen Grüßen" : "Yours sincerely,", Text(letter.Body), StringComparison.Ordinal);
        }

        // DIN 5008: the address field and the information block side by side.
        var din = Document("letter", "de");
        var table = Assert.IsType<DocTable>(din.Body[0]);
        Assert.Equal("plain", table.TableStyle);
        Assert.Equal(["61%", "39%"], table.Columns);
        Assert.Equal(25, din.Page.Left);
        Assert.Equal(20, din.Page.Right);
        Assert.Equal(45, din.Page.Top);
    }

    [Fact]
    public void The_report_has_a_title_page_a_table_of_contents_chapters_a_table_and_page_numbers()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var report = Document("report", language);
            var paragraphs = report.Body.OfType<DocParagraph>().ToList();
            Assert.Equal("Title", paragraphs[0].Style);
            Assert.Equal("Subtitle", paragraphs[1].Style);
            Assert.Contains(paragraphs, p => p.Style == "TOCHeading");
            var toc = paragraphs.Where(p => p.Style is "TOC1" or "TOC2").ToList();
            var headings = paragraphs.Where(p => p.HeadingLevel is 1 or 2).ToList();
            Assert.Equal(headings.Count, toc.Count);
            for (var i = 0; i < toc.Count; i++)
            {
                Assert.Equal("TOC" + headings[i].HeadingLevel, toc[i].Style);
                Assert.Equal("#" + headings[i].Id, toc[i].Inlines[0].Format.Href);
                Assert.Equal(headings[i].PlainText(), toc[i].PlainText());
            }

            Assert.Contains(report.Body, b => b is DocBreak { Kind: "page" });
            var table = Assert.Single(report.Body.OfType<DocTable>());
            Assert.Equal("accent1", table.TableStyle);
            Assert.True(table.Rows[0].Header);
            Assert.Contains(paragraphs, p => p.Style == "Caption");
            Assert.Contains(paragraphs, p => p.Tag == "blockquote");
            Assert.Single(report.Notes);
            Assert.True(report.Page.DifferentFirst);
            Assert.Empty(report.FirstFooter);
            Assert.Equal(["page", "pages"], Fields(report.Footer));
            Assert.Equal(["title"], Fields(report.Header));
        }
    }

    [Fact]
    public void Meeting_notes_have_details_an_agenda_decisions_and_action_items()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var meeting = Document("meeting", language);
            var tables = meeting.Body.OfType<DocTable>().ToList();
            Assert.Equal(["light", "accent1"], tables.Select(t => t.TableStyle));
            Assert.Equal(4, tables[1].ColumnCount());
            Assert.True(tables[1].Rows[0].Header);
            var lists = meeting.Body.OfType<DocParagraph>().Where(p => p.List is not null).GroupBy(p => p.List).ToDictionary(g => g.Key!, g => g.Count());
            Assert.Equal(5, lists["decimal"]);
            Assert.Equal(3, lists["check"]);
            Assert.Equal(["page", "pages"], Fields(meeting.Footer));
        }
    }

    [Fact]
    public void The_resume_has_a_name_contact_links_sections_and_bullets()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var resume = Document("resume", language);
            var paragraphs = resume.Body.OfType<DocParagraph>().ToList();
            Assert.Equal("Title", paragraphs[0].Style);
            Assert.Contains(paragraphs[2].Inlines, i => i.Format.Href?.StartsWith("mailto:", StringComparison.Ordinal) == true);
            Assert.Equal(5, paragraphs.Count(p => p.HeadingLevel == 1));
            Assert.True(paragraphs.Count(p => p.List == "bullet") >= 8);
        }
    }

    [Fact]
    public void The_proposal_has_numbered_and_outline_lists_and_three_tables()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var proposal = Document("proposal", language);
            var paragraphs = proposal.Body.OfType<DocParagraph>().ToList();
            Assert.Contains(paragraphs, p => p is { List: "outline", Level: 1 });
            Assert.Equal(3, paragraphs.Count(p => p.List == "decimal"));
            Assert.Equal(["light", "striped", "accent1", "grid"], proposal.Body.OfType<DocTable>().Select(t => t.TableStyle));
        }
    }

    [Fact]
    public void The_invoice_has_a_letterhead_items_and_totals_that_add_up()
    {
        foreach (var language in new[] { "en", "de" })
        {
            var invoice = Document("invoice", language);
            Assert.Contains(language == "de" ? "Muster GmbH" : "Northwind Consulting", Text(invoice.Header), StringComparison.Ordinal);
            Assert.Equal(["page", "pages"], Fields(invoice.Footer));
            var items = invoice.Body.OfType<DocTable>().Single(t => t.TableStyle == "light");
            Assert.Equal(5, items.ColumnCount());
            var totals = items.Rows.Where(r => r.Cells.Count == 2).ToList();
            Assert.Equal(3, totals.Count);
            Assert.All(totals, r => Assert.Equal(4, r.Cells[0].ColSpan));

            var german = language == "de";
            decimal Amount(DocCell cell)
            {
                var text = Text(cell.Blocks).Trim().Replace("€", string.Empty, StringComparison.Ordinal).Replace("£", string.Empty, StringComparison.Ordinal).Trim();
                return decimal.Parse(german ? text.Replace(".", string.Empty, StringComparison.Ordinal).Replace(',', '.') : text.Replace(",", string.Empty, StringComparison.Ordinal), System.Globalization.CultureInfo.InvariantCulture);
            }

            var lines = items.Rows.Skip(1).Where(r => r.Cells.Count == 5).ToList();
            Assert.All(lines, r => Assert.Equal(Amount(r.Cells[2]) * Amount(r.Cells[3]), Amount(r.Cells[4])));
            var subtotal = lines.Sum(r => Amount(r.Cells[4]));
            Assert.Equal(subtotal, Amount(totals[0].Cells[1]));
            Assert.Equal(Math.Round(subtotal * (german ? 0.19m : 0.20m), 2), Amount(totals[1].Cells[1]));
            Assert.Equal(Amount(totals[0].Cells[1]) + Amount(totals[1].Cells[1]), Amount(totals[2].Cells[1]));
        }
    }

    [Theory]
    [InlineData("de-AT", "de-DE")]
    [InlineData("DE", "de-DE")]
    [InlineData("fr", "en-GB")]
    [InlineData("", "en-GB")]
    public void The_language_chooses_the_text(string language, string expected)
        => Assert.Equal(expected, DocsFormatTests.Parse(Handler.Create("letter", language))["settings"]!["language"]!.GetValue<string>());

    [Theory]
    [InlineData("blank", "en", "en")]
    [InlineData("blank", "de", "de")]
    [InlineData("no-such-template", "de", "de")]
    [InlineData("../report", "en", "en")]
    public void Blank_and_unknown_templates_give_an_empty_document(string id, string language, string expected)
    {
        var content = DocsFormatTests.Parse(Handler.Create(id, language));
        Assert.Equal("<p></p>", content["html"]!.GetValue<string>());
        Assert.Equal(expected, content["settings"]!["language"]!.GetValue<string>());
        Assert.True(Handler.Normalize(content.ToJsonString()).Ok);
    }

    private static DocsDocument Document(string id, string language) => DocsDocument.FromJson(Handler.Normalize(Handler.Create(id, language)).Json!);

    private static string Text(IEnumerable<DocBlock> blocks) => string.Join('\n', DocsHtml.Paragraphs(blocks).Select(p => p.PlainText()));

    private static List<string> Fields(IEnumerable<DocBlock> blocks) => DocsHtml.Paragraphs(blocks).SelectMany(p => p.Inlines).OfType<DocField>().Select(f => f.Kind).ToList();

    private static string FirstWords(string content)
    {
        var text = Handler.PlainText(content, 200).Split('\n', StringSplitOptions.RemoveEmptyEntries)[0].Trim();
        return text.Length > 12 ? text[..12] : text;
    }

    /// <summary>Equal HTML, or a message that shows where the two differ.</summary>
    private static void AssertSame(string expected, string actual, string what)
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

        var from = Math.Max(0, i - 60);
        Assert.Fail($"{what} is not canonical at {i}:\n template: …{expected[from..Math.Min(expected.Length, i + 80)]}\n canonical: …{actual[from..Math.Min(actual.Length, i + 80)]}");
    }
}
