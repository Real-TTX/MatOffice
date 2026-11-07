using System.Text.Json.Nodes;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>The content check of Docs (<see cref="DocsHandler.Normalize"/>) and the plain text of a document.</summary>
public sealed class DocsContentTests
{
    private const string Defaults = "{\"version\":1,\"page\":{\"size\":\"A4\",\"width\":210,\"height\":297,\"orientation\":\"portrait\","
        + "\"margins\":{\"top\":25,\"right\":25,\"bottom\":20,\"left\":25},\"header\":12.5,\"footer\":12.5,\"columns\":1,\"differentFirst\":false},"
        + "\"html\":\"<p></p>\",\"header\":\"\",\"footer\":\"\",\"firstHeader\":\"\",\"firstFooter\":\"\",\"notes\":[],\"comments\":[],"
        + "\"settings\":{\"language\":\"en\",\"spellcheck\":true}}";

    private static readonly DocsHandler Handler = new();

    private static JsonObject Normal(string json)
    {
        var check = Handler.Normalize(json);
        Assert.True(check.Ok, check.Error);
        return DocsFormatTests.Parse(check.Json!);
    }

    private static JsonNode Page(string page) => Normal("{\"page\":" + page + "}")["page"]!;

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"page\":{\"size\":\"A4\",\"orientation\":\"portrait\"},\"html\":\"<p></p>\"}")]
    [InlineData("{\"html\":\"\",\"header\":null,\"notes\":null,\"comments\":\"x\",\"settings\":[]}")]
    [InlineData("{\"html\":\"   \\n \",\"page\":\"A5\",\"version\":7,\"unknown\":{\"a\":1}}")]
    public void Missing_parts_get_their_defaults(string json) => Assert.Equal(Defaults, Handler.Normalize(json).Json);

    [Theory]
    [InlineData("[]")]
    [InlineData("\"text\"")]
    [InlineData("42")]
    [InlineData("null")]
    [InlineData("not json")]
    [InlineData("{\"html\":5}")]
    [InlineData("{\"html\":{\"p\":1}}")]
    [InlineData("{\"footer\":[\"<p>a</p>\"]}")]
    [InlineData("{\"html\":\"<p>a</p>\"")]
    public void Content_that_is_not_a_document_is_refused(string json)
    {
        var check = Handler.Normalize(json);
        Assert.False(check.Ok);
        Assert.Equal(Text.T("Documents.Invalid"), check.Error);
    }

    [Fact]
    public void Content_larger_than_the_limit_is_refused()
    {
        // The size is checked before the JSON is read.
        var check = Handler.Normalize(new string(' ', Kinds.MaxContentBytes + 1));
        Assert.False(check.Ok);
        Assert.Equal(Text.T("Documents.TooLarge"), check.Error);
    }

    [Theory]
    [InlineData("{\"size\":\"Custom\",\"width\":10,\"height\":5000}", "Custom", 50, 1000)]
    [InlineData("{\"size\":\"Custom\",\"width\":\"123.456\",\"height\":180}", "Custom", 123.46, 180)]
    [InlineData("{\"size\":\"A5\",\"width\":999,\"height\":1}", "A5", 148, 210)]
    [InlineData("{\"size\":\"A3\"}", "A3", 297, 420)]
    [InlineData("{\"size\":\"Letter\"}", "Letter", 215.9, 279.4)]
    [InlineData("{\"size\":\"Legal\"}", "Legal", 215.9, 355.6)]
    [InlineData("{\"size\":\"B5\",\"width\":176,\"height\":250}", "A4", 210, 297)]
    [InlineData("{\"size\":\"a4\"}", "A4", 210, 297)]
    [InlineData("{\"width\":100}", "A4", 210, 297)]
    public void The_paper_size_is_a_known_one_or_a_clamped_custom_one(string page, string size, double width, double height)
    {
        var result = Page(page);
        Assert.Equal(size, result["size"]!.GetValue<string>());
        Assert.Equal(width, result["width"]!.GetValue<double>());
        Assert.Equal(height, result["height"]!.GetValue<double>());
    }

    [Fact]
    public void Landscape_keeps_the_portrait_measures()
    {
        var page = Page("{\"size\":\"Letter\",\"orientation\":\"landscape\"}");
        Assert.Equal("landscape", page["orientation"]!.GetValue<string>());
        Assert.Equal(215.9, page["width"]!.GetValue<double>());
        Assert.Equal(279.4, page["height"]!.GetValue<double>());
        Assert.Equal("portrait", Page("{\"orientation\":\"sideways\"}")["orientation"]!.GetValue<string>());
    }

    [Fact]
    public void Margins_are_clamped_and_leave_room_for_the_text()
    {
        var page = Page("{\"margins\":{\"top\":-5,\"right\":500,\"bottom\":\"12\",\"left\":null},\"header\":200,\"footer\":-1,\"columns\":7,\"differentFirst\":true}");
        Assert.Equal(0, page["margins"]!["top"]!.GetValue<double>());
        Assert.Equal(150, page["margins"]!["right"]!.GetValue<double>());
        Assert.Equal(12, page["margins"]!["bottom"]!.GetValue<double>());
        Assert.Equal(25, page["margins"]!["left"]!.GetValue<double>());
        Assert.Equal(150, page["header"]!.GetValue<double>());
        Assert.Equal(0, page["footer"]!.GetValue<double>());
        Assert.Equal(3, page["columns"]!.GetValue<int>());
        Assert.True(page["differentFirst"]!.GetValue<bool>());

        // A narrow custom paper: the margins shrink so that 20 mm of text remain.
        var narrow = Page("{\"size\":\"Custom\",\"width\":60,\"height\":100,\"margins\":{\"top\":60,\"right\":50,\"bottom\":60,\"left\":50}}");
        var left = narrow["margins"]!["left"]!.GetValue<double>();
        var right = narrow["margins"]!["right"]!.GetValue<double>();
        var top = narrow["margins"]!["top"]!.GetValue<double>();
        var bottom = narrow["margins"]!["bottom"]!.GetValue<double>();
        Assert.True(60 - left - right >= 20, $"{left} + {right}");
        Assert.True(100 - top - bottom >= 20, $"{top} + {bottom}");
        Assert.Equal(left, right);
    }

    [Theory]
    [InlineData("0", 1)]
    [InlineData("2", 2)]
    [InlineData("\"3\"", 3)]
    [InlineData("2.6", 3)]
    [InlineData("-4", 1)]
    [InlineData("\"many\"", 1)]
    public void Columns_are_one_to_three(string columns, int expected)
        => Assert.Equal(expected, Page("{\"columns\":" + columns + "}")["columns"]!.GetValue<int>());

    [Fact]
    public void Every_html_field_is_sanitized()
    {
        var content = Normal("{\"html\":\"<p onclick=\\\"x()\\\">body<script>alert(1)</script></p>\",\"header\":\"<div>head<iframe src=\\\"https://x\\\"></iframe></div>\","
            + "\"footer\":\"<p><a href=\\\"javascript:alert(1)\\\">foot</a></p>\",\"firstHeader\":\"plain\",\"firstFooter\":\"<p style=\\\"color: red; behavior: url(x.htc)\\\">ff</p>\"}");
        Assert.Equal("<p>body</p>", content["html"]!.GetValue<string>());
        Assert.Equal("<p>head</p>", content["header"]!.GetValue<string>());
        Assert.Equal("<p>foot</p>", content["footer"]!.GetValue<string>());
        Assert.Equal("<p>plain</p>", content["firstHeader"]!.GetValue<string>());
        Assert.Equal("<p><span style=\"color: #ff0000\">ff</span></p>", content["firstFooter"]!.GetValue<string>());
    }

    [Fact]
    public void Notes_need_a_valid_id_and_a_reference_in_the_body()
    {
        var content = Normal("{\"html\":\"<p>a<sup data-fn=\\\"n1\\\"></sup> b<sup data-fn=\\\"n3\\\"></sup><sup data-fn=\\\"bad id\\\"></sup></p>\","
            + "\"notes\":[{\"id\":\"n1\",\"html\":\"<p>one<script>x</script></p>\"},{\"id\":\"n2\",\"html\":\"<p>two</p>\"},{\"id\":\"bad id\",\"html\":\"x\"},"
            + "{\"id\":\"n1\",\"html\":\"<p>again</p>\"},{\"id\":\"n3\"},\"n4\",{\"html\":\"<p>no id</p>\"}]}");
        Assert.Equal("<p>a<sup data-fn=\"n1\"></sup> b<sup data-fn=\"n3\"></sup></p>", content["html"]!.GetValue<string>());
        var notes = content["notes"]!.AsArray().Select(n => (n!["id"]!.GetValue<string>(), n["html"]!.GetValue<string>())).ToList();
        Assert.Equal([("n1", "<p>one</p>"), ("n3", "<p></p>")], notes);
    }

    [Fact]
    public void A_note_referenced_from_a_table_cell_stays()
    {
        var content = Normal("{\"html\":\"<table><tr><td>cell<sup data-fn=\\\"t1\\\"></sup></td></tr></table>\",\"notes\":[{\"id\":\"t1\",\"html\":\"<p>in a table</p>\"}]}");
        Assert.Single(content["notes"]!.AsArray());
    }

    [Fact]
    public void Comments_are_checked_and_kept_without_an_anchor()
    {
        var longAuthor = new string('a', 300);
        var longText = new string('t', 12_000);
        var content = Normal("{\"html\":\"<p><span data-comment=\\\"c1\\\">x</span></p>\",\"comments\":["
            + "{\"id\":\"c1\",\"author\":\"Ada\\nLovelace\",\"date\":\"2026-11-07T10:00:00Z\",\"text\":\"Check\\r\\nthis\\u0007\",\"done\":true,\"extra\":1,"
            + "\"replies\":[{\"id\":\"r1\",\"author\":\"Bob\",\"date\":\"yesterday\",\"text\":\"Done\",\"done\":true},{\"id\":\"c1\",\"text\":\"same id\"},{\"text\":\"no id\"}]},"
            + "{\"id\":\"c2\",\"author\":\"" + longAuthor + "\",\"date\":\"2026-11-07\",\"text\":\"" + longText + "\",\"done\":\"yes\"},"
            + "{\"id\":\"c1\",\"text\":\"duplicate\"},{\"id\":\"bad id\"},42]}");
        var comments = content["comments"]!.AsArray();
        Assert.Equal(2, comments.Count);
        var first = comments[0]!;
        Assert.Equal("c1", first["id"]!.GetValue<string>());
        Assert.Equal("Ada Lovelace", first["author"]!.GetValue<string>());
        Assert.Equal("2026-11-07T10:00:00Z", first["date"]!.GetValue<string>());
        Assert.Equal("Check\nthis", first["text"]!.GetValue<string>());
        Assert.True(first["done"]!.GetValue<bool>());
        Assert.Null(first["extra"]);
        var reply = Assert.Single(first["replies"]!.AsArray())!;
        Assert.Equal("{\"id\":\"r1\",\"author\":\"Bob\",\"date\":\"\",\"text\":\"Done\"}", reply.ToJsonString());

        var second = comments[1]!;
        Assert.Equal(200, second["author"]!.GetValue<string>().Length);
        Assert.Equal(10_000, second["text"]!.GetValue<string>().Length);
        Assert.Equal("2026-11-07", second["date"]!.GetValue<string>());
        Assert.False(second["done"]!.GetValue<bool>());
        Assert.Empty(second["replies"]!.AsArray());
    }

    [Theory]
    [InlineData("{\"language\":\"de-DE\",\"spellcheck\":false}", "de-DE", false)]
    [InlineData("{\"language\":\"de\"}", "de", true)]
    [InlineData("{\"language\":\"german\",\"spellcheck\":\"no\"}", "en", true)]
    [InlineData("{\"language\":\"DE\"}", "en", true)]
    [InlineData("{\"language\":\"en-us\"}", "en", true)]
    [InlineData("{\"language\":5}", "en", true)]
    public void Settings_have_a_language_and_the_spell_check(string settings, string language, bool spellcheck)
    {
        var result = Normal("{\"settings\":" + settings + "}")["settings"]!;
        Assert.Equal(language, result["language"]!.GetValue<string>());
        Assert.Equal(spellcheck, result["spellcheck"]!.GetValue<bool>());
    }

    [Fact]
    public void Normalizing_twice_changes_nothing()
    {
        var once = DocsFormatTests.Rich();
        Assert.Equal(once, Handler.Normalize(once).Json);
        foreach (var template in Handler.Templates)
        {
            var content = Handler.Normalize(Handler.Create(template.Id, "de")).Json!;
            Assert.Equal(content, Handler.Normalize(content).Json);
        }
    }

    [Fact]
    public void The_rich_document_keeps_what_the_editor_wrote()
    {
        var content = DocsFormatTests.Parse(DocsFormatTests.Rich());
        Assert.Equal("Letter", content["page"]!["size"]!.GetValue<string>());
        Assert.Equal(2, content["page"]!["columns"]!.GetValue<int>());
        Assert.Equal("<p style=\"text-align: right\">Header <span data-field=\"title\"></span></p>", content["header"]!.GetValue<string>());
        Assert.Equal("", content["firstHeader"]!.GetValue<string>());
        Assert.Equal("de-DE", content["settings"]!["language"]!.GetValue<string>());
        Assert.False(content["settings"]!["spellcheck"]!.GetValue<bool>());
        Assert.Single(content["notes"]!.AsArray());
        Assert.Single(content["comments"]!.AsArray()[0]!["replies"]!.AsArray());
        var html = content["html"]!.GetValue<string>();
        Assert.StartsWith("<h1 id=\"h-intro\">Introduction</h1><p data-style=\"Title\">The title</p>", html, StringComparison.Ordinal);
        Assert.Contains("<p data-list=\"decimal\" data-start=\"3\">Three</p>", html, StringComparison.Ordinal);
        Assert.Contains("<td rowspan=\"2\"><p>merged down</p></td><td colspan=\"2\" style=\"background-color: #ff0000; vertical-align: middle\">", html, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ plain text

    [Fact]
    public void The_plain_text_has_a_line_per_paragraph_and_tabs_between_cells()
    {
        var text = Handler.PlainText(DocsFormatTests.Rich(), 100_000);
        var lines = text.Split('\n');
        Assert.Equal("Introduction", lines[0]);
        Assert.Equal("The title", lines[1]);
        Assert.Contains("One.one", lines);
        Assert.Contains("A\tB\tC", lines);
        Assert.Contains("merged down\tmerged right", lines);
        Assert.Contains("code", lines);
        Assert.Contains("  indented", lines);
        Assert.Contains(lines, l => l.StartsWith("Normal bold italic under strike x2 H2O styled shaded a link back", StringComparison.Ordinal));
        Assert.Contains(" and a note.\tTab", text, StringComparison.Ordinal);
        Assert.DoesNotContain("<", text, StringComparison.Ordinal);
        Assert.False(text.EndsWith('\n'));
    }

    [Theory]
    [InlineData(5)]
    [InlineData(1)]
    [InlineData(34)]
    [InlineData(37)]
    public void The_plain_text_is_cut_at_the_limit(int max)
    {
        var text = Handler.PlainText(DocsFormatTests.Rich(), max);
        Assert.True(text.Length <= max);
        Assert.Equal(text.TrimEnd(), text);
        Assert.StartsWith(text, Handler.PlainText(DocsFormatTests.Rich(), 100_000), StringComparison.Ordinal);
    }

    [Fact]
    public void The_plain_text_does_not_cut_a_character_in_half()
    {
        var text = Handler.PlainText("{\"html\":\"<p>ab😀c</p>\"}", 3);
        Assert.Equal("ab", text);
    }

    [Theory]
    [InlineData("{\"html\":\"<p>a</p>\"}", 0, "")]
    [InlineData("{\"html\":\"<p>a</p>\"}", -3, "")]
    [InlineData("not json", 100, "")]
    [InlineData("[1,2]", 100, "")]
    [InlineData("{}", 100, "")]
    [InlineData("{\"html\":\"<p>a &amp; b</p><p></p><p>c<br>d</p>\"}", 100, "a & b\n\nc\nd")]
    public void The_plain_text_of_odd_content(string json, int max, string expected) => Assert.Equal(expected, Handler.PlainText(json, max));
}
