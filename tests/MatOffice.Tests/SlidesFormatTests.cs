using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Slides;

namespace MatOffice.Tests;

/// <summary>Slides: the content check, the templates, PowerPoint export (valid for the SDK's validator) and import (a round trip), the data shared with the editor.</summary>
public sealed class SlidesFormatTests
{
    private const string Png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static readonly SlidesHandler Handler = new();

    public static TheoryData<string, string> Templates()
    {
        var data = new TheoryData<string, string>();
        foreach (var t in SlidesTemplates.All)
        {
            data.Add(t.Id, "en");
            data.Add(t.Id, "de");
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(Templates))]
    public void Templates_are_valid_and_complete(string id, string language)
    {
        var check = Handler.Normalize(Handler.Create(id, language));
        Assert.True(check.Ok, check.Error);
        var root = JsonNode.Parse(check.Json!)!.AsObject();
        var slides = root["slides"]!.AsArray();
        Assert.True(id == "blank" ? slides.Count == 1 : slides.Count >= 5, $"{id}: {slides.Count} slides");
        Assert.All(slides, s => Assert.True(SlidesData.IsLayout(s!["layout"]!.GetValue<string>())));
        if (id != "blank")
        {
            Assert.True(slides.Select(s => s!["layout"]!.GetValue<string>()).Distinct().Count() >= 3, id + " uses several layouts");
            Assert.NotEqual("paper", root["theme"]!["id"]!.GetValue<string>());
            Assert.False(string.IsNullOrWhiteSpace(Handler.PlainText(check.Json!, 2000)));
        }

        Assert.All(SlidesTemplates.All, t => Assert.False(MatOffice.Localization.Text.T(t.Name) == t.Name, t.Name));
    }

    [Fact]
    public void Templates_speak_the_language()
    {
        var en = Handler.PlainText(Handler.Create("status", "en"), 4000);
        var de = Handler.PlainText(Handler.Create("status", "de"), 4000);
        Assert.Contains("Milestones", en, StringComparison.Ordinal);
        Assert.Contains("Meilensteine", de, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("{ \"size\": { \"w\": 1280, \"h\": 720 } }")]
    [InlineData("{ \"slides\": 5 }")]
    [InlineData("{ \"slides\": [ 1 ] }")]
    [InlineData("{ \"slides\": [ { \"elements\": [ { \"type\": \"script\" } ] } ] }")]
    [InlineData("{ \"slides\": [ { \"elements\": [ { \"type\": \"image\", \"src\": \"https://example.com/a.png\" } ] } ] }")]
    [InlineData("{ \"slides\": [ { \"elements\": [ { \"type\": \"image\", \"src\": \"data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=\" } ] } ] }")]
    [InlineData("{ \"slides\": [ { \"elements\": [ { \"type\": \"image\", \"src\": \"data:image/png;base64,<script>\" } ] } ] }")]
    [InlineData("{ \"slides\": [ { \"background\": { \"type\": \"image\", \"src\": \"javascript:alert(1)\" }, \"elements\": [] } ] }")]
    [InlineData("{ \"slides\": [ { \"elements\": [ { \"type\": \"table\", \"cols\": [], \"rows\": [] } ] } ] }")]
    public void Normalize_rejects_bad_content(string json) => Assert.False(Handler.Normalize(json).Ok, json);

    [Fact]
    public void Normalize_rejects_too_many_slides()
    {
        var slides = new JsonArray();
        for (var i = 0; i <= SlidesNormalizer.MaxSlides; i++)
        {
            slides.Add(new JsonObject { ["elements"] = new JsonArray() });
        }

        Assert.False(Handler.Normalize(new JsonObject { ["slides"] = slides }.ToJsonString()).Ok);
    }

    [Fact]
    public void Normalize_keeps_only_what_it_knows()
    {
        var json = """
        {
          "evil": "<script>alert(1)</script>",
          "size": { "w": 99999, "h": "720" },
          "theme": { "id": "slate", "colors": { "accent1": "red; background: url(x)" }, "fonts": { "major": "Arial\"; } body { x: y", "minor": "Segoe UI" } },
          "slides": [ {
            "id": "a b", "layout": "nope", "onload": "x", "notes": "Hello\u0000 notes\r\nline 2",
            "transition": { "type": "explode", "duration": 99 },
            "elements": [
              { "id": "e1", "type": "text", "x": "12", "y": 1e9, "w": -5, "h": 50, "rot": -90, "html": "<b>x</b>", "ph": "title<script>",
                "text": { "paras": [ { "runs": [ { "t": "Hi\u0007 there", "b": "yes", "i": true, "size": 9999, "color": "url(javascript:x)", "font": "Comic Sans MS" } ], "align": "middle", "level": 42, "bullet": "1." } ] } },
              { "id": "e1", "type": "shape", "shape": "pentagram", "fill": { "type": "solid", "color": "accent3:+40" }, "line": { "color": "#ABCDEF", "width": 500, "dash": "wavy" } },
              { "id": "e3", "type": "line", "group": "g1" }
            ],
            "animations": [ { "el": "missing" }, { "el": "e1", "effect": "spin", "trigger": "with" } ]
          } ]
        }
        """;
        var check = Handler.Normalize(json);
        Assert.True(check.Ok, check.Error);
        var root = JsonNode.Parse(check.Json!)!.AsObject();
        Assert.Null(root["evil"]);
        Assert.Equal(6000, root["size"]!["w"]!.GetValue<double>());
        Assert.Equal(720, root["size"]!["h"]!.GetValue<double>());
        Assert.Equal("#2ec4b6", root["theme"]!["colors"]!["accent1"]!.GetValue<string>());
        Assert.Equal("Segoe UI Light", root["theme"]!["fonts"]!["major"]!.GetValue<string>());
        var slide = root["slides"]![0]!.AsObject();
        Assert.Equal("s1", slide["id"]!.GetValue<string>());
        Assert.Equal("blank", slide["layout"]!.GetValue<string>());
        Assert.Null(slide["onload"]);
        Assert.Null(slide["transition"]);
        Assert.Equal("Hello notes\nline 2", slide["notes"]!.GetValue<string>());
        var elements = slide["elements"]!.AsArray();
        var text = elements[0]!.AsObject();
        Assert.Null(text["html"]);
        Assert.Null(text["ph"]);
        Assert.Equal(12, text["x"]!.GetValue<double>());
        Assert.Equal(20000, text["y"]!.GetValue<double>());
        Assert.Equal(0, text["w"]!.GetValue<double>());
        Assert.Equal(270, text["rot"]!.GetValue<double>());
        var para = text["text"]!["paras"]![0]!.AsObject();
        Assert.Null(para["align"]);
        Assert.Equal(8, para["level"]!.GetValue<int>());
        var run = para["runs"]![0]!.AsObject();
        Assert.Equal("Hi there", run["t"]!.GetValue<string>());
        Assert.Null(run["b"]);
        Assert.True(run["i"]!.GetValue<bool>());
        Assert.Equal(400, run["size"]!.GetValue<double>());
        Assert.Null(run["color"]);
        Assert.Equal("Comic Sans MS", run["font"]!.GetValue<string>());
        var shape = elements[1]!.AsObject();
        Assert.Equal("e1_2", shape["id"]!.GetValue<string>());
        Assert.Equal("rect", shape["shape"]!.GetValue<string>());
        Assert.Equal("accent3:+40", shape["fill"]!["color"]!.GetValue<string>());
        Assert.Equal("#abcdef", shape["line"]!["color"]!.GetValue<string>());
        Assert.Equal(100, shape["line"]!["width"]!.GetValue<double>());
        Assert.Equal("solid", shape["line"]!["dash"]!.GetValue<string>());
        Assert.Null(elements[2]!["group"]);
        var anims = slide["animations"]!.AsArray();
        Assert.Single(anims);
        Assert.Equal("fade", anims[0]!["effect"]!.GetValue<string>());
        Assert.Equal("with", anims[0]!["trigger"]!.GetValue<string>());
    }

    [Fact]
    public void Plain_text_has_slides_tables_and_notes()
    {
        var text = Handler.PlainText(Sample().ToJsonString(), 10000);
        Assert.Contains("Quarterly figures", text, StringComparison.Ordinal);
        Assert.Contains("Revenue", text, StringComparison.Ordinal);
        Assert.Contains("Say hello first", text, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Templates))]
    public void Export_of_every_template_passes_the_validator(string id, string language)
    {
        var content = Handler.Normalize(Handler.Create(id, language)).Json!;
        var bytes = Handler.Export(content, ".pptx", id);
        AssertValid(bytes);
    }

    [Fact]
    public void Export_of_everything_passes_the_validator_and_reads_back()
    {
        var content = Handler.Normalize(Sample().ToJsonString());
        Assert.True(content.Ok, content.Error);
        var bytes = Handler.Export(content.Json!, ".pptx", "Sample");
        AssertValid(bytes);

        var read = Kinds.Read(Handler, new MemoryStream(bytes), ".pptx");
        Assert.True(read.Ok, read.Error);
        var back = JsonNode.Parse(read.Json!)!.AsObject();
        var original = JsonNode.Parse(content.Json!)!.AsObject();

        Assert.Equal(1280, back["size"]!["w"]!.GetValue<double>());
        Assert.Equal(720, back["size"]!["h"]!.GetValue<double>());
        Assert.Equal(original["theme"]!["colors"]!["accent1"]!.GetValue<string>(), back["theme"]!["colors"]!["accent1"]!.GetValue<string>());
        Assert.Equal(original["theme"]!["fonts"]!["major"]!.GetValue<string>(), back["theme"]!["fonts"]!["major"]!.GetValue<string>());
        var slides = back["slides"]!.AsArray();
        Assert.Equal(3, slides.Count);

        // slide 1: title layout, placeholders with their text
        var s1 = slides[0]!.AsObject();
        Assert.Equal("title", s1["layout"]!.GetValue<string>());
        Assert.Equal("Say hello first", s1["notes"]!.GetValue<string>());
        Assert.Equal("fade", s1["transition"]!["type"]!.GetValue<string>());
        var title = s1["elements"]!.AsArray().First(e => e!["ph"]?.GetValue<string>() == "title")!;
        Assert.Equal("Quarterly figures", Text(title));
        Assert.Equal(160, title["x"]!.GetValue<double>(), 0);

        // slide 2: shapes, a line, a picture, a table, a group, an animation, a hidden slide with a background
        var s2 = slides[1]!.AsObject();
        Assert.Equal("titleOnly", s2["layout"]!.GetValue<string>());
        Assert.True(s2["hidden"]!.GetValue<bool>());
        Assert.Equal("#102030", s2["background"]!["color"]!.GetValue<string>());
        Assert.Equal("push", s2["transition"]!["type"]!.GetValue<string>());
        Assert.Equal("left", s2["transition"]!["dir"]!.GetValue<string>());
        var elements = s2["elements"]!.AsArray().Select(e => e!.AsObject()).ToList();
        var star = elements.First(e => e["shape"]?.GetValue<string>() == "star5");
        Assert.Equal(300, star["x"]!.GetValue<double>(), 0);
        Assert.Equal(45, star["rot"]!.GetValue<double>(), 0);
        Assert.Equal("accent2:+40", star["fill"]!["color"]!.GetValue<string>());
        Assert.Equal("dash", star["line"]!["dash"]!.GetValue<string>());
        Assert.True(star["shadow"]!.GetValue<bool>());
        Assert.Equal("Star", Text(star));
        var run = star["text"]!["paras"]![0]!["runs"]![0]!;
        Assert.True(run["b"]!.GetValue<bool>());
        Assert.Equal(28, run["size"]!.GetValue<double>());
        Assert.Equal("#ff0000", run["color"]!.GetValue<string>());
        var line = elements.First(e => e["type"]!.GetValue<string>() == "line");
        Assert.Equal("triangle", line["arrowEnd"]!.GetValue<string>());
        Assert.True(line["flipV"]!.GetValue<bool>());
        var image = elements.First(e => e["type"]!.GetValue<string>() == "image");
        Assert.StartsWith("data:image/png;base64,", image["src"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal(0.1, image["crop"]!["l"]!.GetValue<double>(), 3);
        var table = elements.First(e => e["type"]!.GetValue<string>() == "table");
        Assert.Equal(3, table["cols"]!.AsArray().Count);
        Assert.Equal(2, table["rows"]!.AsArray().Count);
        Assert.Equal("Revenue", Text(table["rows"]![0]!["cells"]![1]!));
        var grouped = elements.Where(e => e["group"] is not null).ToList();
        Assert.Equal(2, grouped.Count);
        Assert.Single(grouped.Select(g => g["group"]!.GetValue<string>()).Distinct());
        Assert.Equal(700, grouped[0]["x"]!.GetValue<double>(), 0);
        var anims = s2["animations"]!.AsArray();
        Assert.Equal(2, anims.Count);
        Assert.Equal("fly", anims[0]!["effect"]!.GetValue<string>());
        Assert.Equal(star["id"]!.GetValue<string>(), anims[0]!["el"]!.GetValue<string>());
        Assert.Equal("after", anims[1]!["trigger"]!.GetValue<string>());

        // slide 3: bullets and levels in the body placeholder
        var s3 = slides[2]!.AsObject();
        Assert.Equal("twoContent", s3["layout"]!.GetValue<string>());
        var left = s3["elements"]!.AsArray().First(e => e!["ph"]?.GetValue<string>() == "left")!;
        var paras = left["text"]!["paras"]!.AsArray();
        Assert.Equal(3, paras.Count);
        Assert.Equal(1, paras[1]!["level"]!.GetValue<int>());
        Assert.Equal("1.", paras[2]!["bullet"]!.GetValue<string>());
        Assert.Equal("center", paras[2]!["align"]!.GetValue<string>());
        Assert.True(back["footer"]!["number"]!.GetValue<bool>());
        Assert.Equal("Company", back["footer"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Import_of_garbage_fails_with_a_message()
    {
        var read = Kinds.Read(Handler, new MemoryStream([1, 2, 3, 4]), ".pptx");
        Assert.False(read.Ok);
        Assert.False(string.IsNullOrWhiteSpace(read.Error));
    }

    [Fact]
    public void The_editor_has_the_same_layouts_and_designs()
    {
        var root = FindRoot();
        Assert.True(JsonNode.DeepEquals(Strip(SlidesData.RawLayouts), FromScript(Path.Combine(root, "src", "MatOffice", "wwwroot", "js", "slides", "layouts-data.js"))), "layouts-data.js differs from layouts.json");
        Assert.True(JsonNode.DeepEquals(Strip(SlidesData.RawThemes), FromScript(Path.Combine(root, "src", "MatOffice", "wwwroot", "js", "slides", "themes-data.js"))), "themes-data.js differs from themes.json");
        Assert.True(SlidesData.ThemeIds.Count() >= 8);
    }

    [Fact]
    public void Four_to_three_is_written_as_such()
    {
        var content = SlidesTemplates.Build("blank", "en");
        content["size"] = new JsonObject { ["w"] = 960, ["h"] = 720 };
        var bytes = Handler.Export(Handler.Normalize(content.ToJsonString()).Json!, ".pptx", "x");
        AssertValid(bytes);
        var back = JsonNode.Parse(Kinds.Read(Handler, new MemoryStream(bytes), ".pptx").Json!)!;
        Assert.Equal(960, back["size"]!["w"]!.GetValue<double>());
    }

    private static string Text(JsonNode element)
        => string.Join('\n', element["text"]!["paras"]!.AsArray().Select(p => string.Concat(p!["runs"]!.AsArray().Select(r => r!["t"]!.GetValue<string>()))));

    private static void AssertValid(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = PresentationDocument.Open(ms, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).Take(20)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static JsonObject Sample()
    {
        var deck = new Deck("meridian");
        deck.Root["footer"] = new JsonObject { ["number"] = true, ["date"] = true, ["dateText"] = string.Empty, ["text"] = "Company", ["skipTitle"] = true };
        deck.Slide("title").Ph("title", "Quarterly figures").Ph("subtitle", "Q3 2026").Notes("Say hello first").Slide["transition"] = new JsonObject { ["type"] = "fade", ["duration"] = 0.7 };
        var s2 = deck.Slide("titleOnly").Ph("title", "Everything");
        s2.Slide["hidden"] = true;
        s2.Slide["background"] = new JsonObject { ["type"] = "solid", ["color"] = "#102030" };
        s2.Slide["transition"] = new JsonObject { ["type"] = "push", ["dir"] = "left", ["duration"] = 0.5 };
        var star = s2.Shape("star5", 300, 200, 200, 200, "accent2:+40", "Star", 28, "#ff0000", bold: true);
        var el = s2.Slide["elements"]!.AsArray().Last()!.AsObject();
        el["rot"] = 45;
        el["shadow"] = true;
        el["line"] = new JsonObject { ["color"] = "accent1", ["width"] = 2, ["dash"] = "dash" };
        s2.Slide["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "ln", ["type"] = "line", ["x"] = 100, ["y"] = 500, ["w"] = 300, ["h"] = 100, ["flipV"] = true,
            ["line"] = new JsonObject { ["color"] = "dk1", ["width"] = 3, ["dash"] = "solid" }, ["arrowStart"] = "none", ["arrowEnd"] = "triangle",
        });
        s2.Slide["elements"]!.AsArray().Add(new JsonObject
        {
            ["id"] = "img", ["type"] = "image", ["x"] = 900, ["y"] = 400, ["w"] = 200, ["h"] = 150, ["src"] = Png, ["alt"] = "A dot",
            ["crop"] = new JsonObject { ["l"] = 0.1, ["t"] = 0, ["r"] = 0, ["b"] = 0.05 },
        });
        s2.Table(100, 560, 600, 40, [["Region", "Revenue", "Change"], ["North", "€3.6 m", "+16 %"]], "accent1");
        var g1 = s2.Shape("rect", 700, 100, 100, 50, "accent3");
        var g2 = s2.Shape("ellipse", 820, 100, 100, 50, "accent4");
        foreach (var e in s2.Slide["elements"]!.AsArray().OfType<JsonObject>().Where(e => e["id"]!.GetValue<string>() == g1 || e["id"]!.GetValue<string>() == g2))
        {
            e["group"] = "grp";
        }

        s2.Animate(star, "fly", "click", "left").Animate(g1, "fade", "after");
        var s3 = deck.Slide("twoContent").Ph("title", "Two").Ph("left", "One", "\tSub", "Three").Ph("right", "Right");
        var third = s3.Slide["elements"]!.AsArray().OfType<JsonObject>().First(e => e["ph"]?.GetValue<string>() == "left")["text"]!["paras"]![2]!.AsObject();
        third["bullet"] = "1.";
        third["align"] = "center";
        return deck.Root;
    }

    private static JsonObject Strip(JsonObject o)
    {
        var copy = o.DeepClone().AsObject();
        copy.Remove("//");
        return copy;
    }

    private static JsonNode FromScript(string path)
    {
        var text = File.ReadAllText(path);
        var start = text.IndexOf("/*DATA*/", StringComparison.Ordinal) + "/*DATA*/".Length;
        var end = text.LastIndexOf("/*DATA*/", StringComparison.Ordinal);
        return JsonNode.Parse(text[start..end])!;
    }

    private static string FindRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "MatOffice.slnx")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName ?? throw new InvalidOperationException("MatOffice.slnx not found.");
    }
}
