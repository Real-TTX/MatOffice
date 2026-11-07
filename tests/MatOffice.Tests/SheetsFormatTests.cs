using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Sheets;

namespace MatOffice.Tests;

/// <summary>Sheets on the server: the check of the content, the templates, and the conversions from and to .xlsx, .csv and .tsv.</summary>
public sealed class SheetsFormatTests
{
    private static readonly SheetsHandler Handler = new();

    // ---- templates and the blank workbook --------------------------------------------------------------------------------------------

    public static TheoryData<string, string> TemplatesAndLanguages()
    {
        var data = new TheoryData<string, string>();
        foreach (var template in Handler.Templates.Where(t => t.Id != "blank"))
        {
            data.Add(template.Id, "en");
            data.Add(template.Id, "de");
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(TemplatesAndLanguages))]
    public void Every_template_is_a_valid_workbook_with_formulas(string id, string language)
    {
        var raw = Handler.Create(id, language);
        Assert.DoesNotContain("\"en\":", raw, StringComparison.Ordinal);
        var check = Handler.Normalize(raw);
        Assert.True(check.Ok, check.Error);
        var content = Parse(check.Json!);
        var cells = content["sheets"]![0]!["cells"]!.AsObject();
        Assert.True(cells.Count > 10);
        Assert.Contains(cells, c => c.Value!["f"] is not null);

        // nothing was dropped by the check: the template is canonical already
        var original = Parse(raw)["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal(original.Count, cells.Count);
        Assert.Equal(Parse(raw)["styles"]!.AsArray().Count, content["styles"]!.AsArray().Count);
    }

    [Fact]
    public void Templates_speak_the_language_of_the_document()
    {
        var en = Parse(Handler.Normalize(Handler.Create("tasks", "en")).Json!);
        var de = Parse(Handler.Normalize(Handler.Create("tasks", "de")).Json!);
        Assert.Equal("Tasks", (string?)en["sheets"]![0]!["name"]);
        Assert.Equal("Aufgaben", (string?)de["sheets"]![0]!["name"]);
        Assert.Equal("Offen,In Arbeit,Erledigt", (string?)de["sheets"]![0]!["validations"]![0]!["v1"]);
        var styles = de["styles"]!.AsArray().Select(s => (string?)s!["nf"]).ToList();
        Assert.Contains("dd.mm.yyyy", styles);
        Assert.Contains("TODAY()", (string?)en["sheets"]![0]!["cells"]!["E4"]!["f"], StringComparison.Ordinal);
        Assert.Equal("Sheets.Template.Tasks", Handler.Templates.Single(t => t.Id == "tasks").Name);
        Assert.Equal("blank", Handler.Templates[0].Id);
    }

    [Fact]
    public void The_sales_template_has_a_chart()
    {
        var content = Parse(Handler.Normalize(Handler.Create("sales", "en")).Json!);
        var chart = content["sheets"]![0]!["charts"]![0]!;
        Assert.Equal("column", (string?)chart["type"]);
        Assert.Equal("A4:E9", (string?)chart["range"]);
    }

    [Theory]
    [InlineData("en", "Sheet1")]
    [InlineData("de", "Tabelle1")]
    [InlineData("fr", "Sheet1")]
    public void A_blank_workbook_has_one_empty_sheet(string language, string name)
    {
        var check = Handler.Normalize(Handler.Create("blank", language));
        Assert.True(check.Ok, check.Error);
        var content = Parse(check.Json!);
        var sheet = Assert.Single(content["sheets"]!.AsArray())!;
        Assert.Equal(name, (string?)sheet["name"]);
        Assert.Empty(sheet["cells"]!.AsObject());
        Assert.Equal(0, (int)content["active"]!);
        Assert.Single(content["styles"]!.AsArray());
        Assert.Equal(Handler.Create("unknown", language), Handler.Create("blank", language));
    }

    [Fact]
    public void The_formats_are_found_by_extension()
    {
        foreach (var ext in (string[])[".xlsx", ".csv", ".tsv", "XLSX"])
        {
            var found = Kinds.ForExtension(ext);
            Assert.NotNull(found);
            Assert.Same(typeof(SheetsHandler), found.Value.Handler.GetType());
        }

        Assert.Equal("Excel-Arbeitsmappe (.xlsx)", Languages.In("de", () => Handler.Formats[0].Label));
        Assert.Equal("Excel workbook (.xlsx)", Languages.In("en", () => Handler.Formats[0].Label));
    }

    // ---- the check ----------------------------------------------------------------------------------------------------------------------

    [Theory]
    [InlineData("[]")]
    [InlineData("{}")]
    [InlineData("""{ "sheets": [] }""")]
    [InlineData("""{ "sheets": {} }""")]
    [InlineData("""{ "sheets": [ 1, "x" ] }""")]
    [InlineData("not json")]
    public void Normalize_refuses_what_is_not_a_workbook(string json)
        => Assert.False(Handler.Normalize(json).Ok);

    [Fact]
    public void Normalize_refuses_too_many_sheets_and_cells()
    {
        var sheets = new JsonArray();
        for (var i = 0; i < 201; i++)
        {
            sheets.Add(new JsonObject { ["name"] = "S" + i, ["cells"] = new JsonObject() });
        }

        Assert.False(Handler.Normalize(new JsonObject { ["sheets"] = sheets }.ToJsonString()).Ok);

        var sb = new StringBuilder("""{ "sheets": [ { "cells": {""");
        for (var i = 1; i <= 2_000_001; i++)
        {
            sb.Append(i == 1 ? "\"K" : ",\"K").Append(i).Append("\":{}");
        }

        sb.Append("} } ] }");
        var check = Languages.In("en", () => Handler.Normalize(sb.ToString()));
        Assert.False(check.Ok);
        Assert.Contains("2,000,000", check.Error, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalize_drops_what_does_not_fit()
    {
        var json = """
        {
          "version": 7, "active": 9, "evil": "<script>",
          "styles": [ { "b": true }, { "b": true, "fill": "#FFC7CE", "color": "red", "font": "Comic<script>", "size": 900, "nf": "#,##0.00", "bt": "thin", "bb": "thick #ABCDEF", "bl": "wavy #000000", "onclick": "x" } ],
          "names": [ { "name": "Rate", "ref": "=Sheet1!$B$2" }, { "name": "A1", "ref": "Sheet1!$A$1" }, { "name": "rate", "ref": "1" } ],
          "sheets": [
            { "id": "s1", "name": "Data", "x": 1, "cells": {
                "a1": { "v": "Text\u0001with\r\ncontrol", "s": 1, "link": "javascript:alert(1)", "style": "x" },
                "B2": { "v": "#DIV/0!", "t": "e", "f": "=1/0" },
                "C3": { "v": 5, "s": 7, "link": "https://example.com/a?b=1" },
                "D4": { "v": "x", "link": "data:text/html,<b>" },
                "E5": { "v": "mail", "link": "mailto:a@example.com", "note": "hello\u0007" },
                "F6": { "v": "#FOO!", "t": "e" },
                "ZZZZ1": { "v": 1 }, "A0": { "v": 1 }, "XFD1048577": { "v": 1 },
                "G7": { "f": "LONG", "v": 3 },
                "H8": { "v": { "nested": true } },
                "I9": {}
              },
              "images": [
                { "src": "data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=", "x": 1, "y": 1, "w": 10, "h": 10 },
                { "src": "data:text/html;base64,PGI+", "x": 1, "y": 1, "w": 10, "h": 10 },
                { "src": "data:image/png;base64,iVBORw0KGgo=", "x": -5, "y": 1, "w": 10, "h": 10, "onload": "x" }
              ],
              "merges": [ "A1:B2", "B2:C3", "D1", "E1:F1" ],
              "color": "#GGGGGG", "zoom": 9, "grid": false, "freeze": { "rows": 1, "cols": "x" } },
            { "id": "s1", "name": "data", "cells": {} },
            { "id": "bad id!", "name": "Bad:/\\?*[]Name that is far too long for Excel", "cells": {}, "hidden": true }
          ]
        }
        """.Replace("LONG", new string('1', 8193), StringComparison.Ordinal);
        var check = Languages.In("en", () => Handler.Normalize(json));
        Assert.True(check.Ok, check.Error);
        var content = Parse(check.Json!);
        Assert.Null(content["evil"]);
        Assert.Equal(1, (int)content["version"]!);
        Assert.Equal(0, (int)content["active"]!);

        var styles = content["styles"]!.AsArray();
        Assert.Empty(styles[0]!.AsObject());
        var style = styles[1]!.AsObject();
        Assert.Equal("#ffc7ce", (string?)style["fill"]);
        Assert.Null(style["color"]);
        Assert.Null(style["font"]);
        Assert.Null(style["size"]);
        Assert.Null(style["onclick"]);
        Assert.Null(style["bl"]);
        Assert.Equal("thin #000000", (string?)style["bt"]);
        Assert.Equal("thick #abcdef", (string?)style["bb"]);

        var names = content["names"]!.AsArray();
        Assert.Equal("Sheet1!$B$2", (string?)Assert.Single(names)!["ref"]);

        var sheets = content["sheets"]!.AsArray();
        Assert.Equal(["Data", "data (2)"], sheets.Take(2).Select(s => (string?)s!["name"]));
        Assert.Equal("BadName that is far too long fo", (string?)sheets[2]!["name"]);
        Assert.Equal(3, sheets.Select(s => (string?)s!["id"]).Distinct().Count());
        Assert.Equal("s1", (string?)sheets[0]!["id"]);

        var sheet = sheets[0]!.AsObject();
        Assert.Null(sheet["x"]);
        Assert.Null(sheet["color"]);
        Assert.Null(sheet["zoom"]);
        Assert.False((bool)sheet["grid"]!);
        Assert.Equal(0, (int)sheet["freeze"]!["cols"]!);
        Assert.Equal(["A1:B2", "E1:F1"], sheet["merges"]!.AsArray().Select(m => (string?)m));

        var cells = sheet["cells"]!.AsObject();
        Assert.Equal(["A1", "B2", "C3", "D4", "E5", "F6", "G7"], cells.Select(c => c.Key));
        Assert.Equal("Textwith\ncontrol", (string?)cells["A1"]!["v"]);
        Assert.Null(cells["A1"]!["link"]);
        Assert.Null(cells["A1"]!["style"]);
        Assert.Equal(1, (int)cells["A1"]!["s"]!);
        Assert.Equal("1/0", (string?)cells["B2"]!["f"]);
        Assert.Equal("e", (string?)cells["B2"]!["t"]);
        Assert.Null(cells["C3"]!["s"]);
        Assert.Equal("https://example.com/a?b=1", (string?)cells["C3"]!["link"]);
        Assert.Null(cells["D4"]!["link"]);
        Assert.Equal("mailto:a@example.com", (string?)cells["E5"]!["link"]);
        Assert.Equal("hello", (string?)cells["E5"]!["note"]);
        Assert.Null(cells["F6"]!["t"]);
        Assert.Null(cells["G7"]!["f"]);
        Assert.Equal(3, (double)cells["G7"]!["v"]!);

        var image = Assert.Single(sheet["images"]!.AsArray())!;
        Assert.StartsWith("data:image/png;base64,", (string?)image["src"], StringComparison.Ordinal);
        Assert.Null(image["onload"]);
        Assert.Equal(0, (double)image["x"]!);

        Assert.Null(sheets[1]!["hidden"]);
        Assert.True((bool)sheets[2]!["hidden"]!);
    }

    [Fact]
    public void Normalize_keeps_a_visible_active_sheet_and_cuts_long_text()
    {
        var json = new JsonObject
        {
            ["active"] = 0,
            ["sheets"] = new JsonArray(
                new JsonObject { ["name"] = "A", ["hidden"] = true, ["cells"] = new JsonObject { ["A1"] = new JsonObject { ["v"] = new string('x', 40_000) } } },
                new JsonObject { ["name"] = "B", ["hidden"] = true }),
        }.ToJsonString();
        var content = Parse(Handler.Normalize(json).Json!);
        Assert.Equal(0, (int)content["active"]!);
        Assert.Null(content["sheets"]![0]!["hidden"]);
        Assert.Equal(32_767, ((string?)content["sheets"]![0]!["cells"]!["A1"]!["v"])!.Length);
        Assert.NotNull(content["sheets"]![1]!["cells"]);
    }

    [Fact]
    public void Normalize_is_stable()
    {
        foreach (var template in Handler.Templates)
        {
            var once = Handler.Normalize(Handler.Create(template.Id, "de")).Json!;
            Assert.Equal(once, Handler.Normalize(once).Json);
        }
    }

    // ---- xlsx ---------------------------------------------------------------------------------------------------------------------------

    [Fact]
    public void Xlsx_round_trip_keeps_the_workbook()
    {
        var content = """
        {
          "version": 1, "active": 1,
          "styles": [ {}, { "b": true, "fill": "#107c41", "color": "#ffffff", "bt": "thin #000000", "bb": "double #ff0000", "ha": "center", "va": "middle", "wrap": true },
                      { "nf": "dd.mm.yyyy" }, { "nf": "#,##0.00", "i": true, "u": true, "strike": true, "font": "Arial", "size": 14 }, { "fill": "#fff2cc", "indent": 2 } ],
          "names": [ { "name": "Rate", "ref": "Data!$B$2" } ],
          "sheets": [
            { "id": "s1", "name": "Data", "color": "#c00000",
              "cells": {
                "A1": { "v": "Header", "s": 1 },
                "B1": { "v": "Ünïcødé €", "note": "A note\nwith two lines" },
                "A2": { "v": 46000, "s": 2 },
                "B2": { "v": 0.25 },
                "C2": { "v": 1234.5, "s": 3 },
                "D2": { "v": true },
                "A3": { "f": "SUM(B2:C2)", "v": 1234.75 },
                "B3": { "f": "1/0", "v": "#DIV/0!", "t": "e" },
                "C3": { "f": "IF(D2,\"yes\",\"no\")", "v": "yes" },
                "D3": { "f": "XLOOKUP(1,B2:C2,B2:C2)", "v": 1 },
                "E3": { "f": "SUM(", "v": 5 },
                "F3": { "f": "B2>0", "v": true },
                "A4": { "v": "Site", "link": "https://example.com/" },
                "B4": { "v": "Jump", "link": "#'Second sheet'!B2" },
                "C5": { "s": 4 }
              },
              "cols": { "A": { "w": 120 }, "C": { "hidden": true }, "D": { "s": 4 } },
              "rows": { "1": { "h": 30 }, "5": { "hidden": true } },
              "merges": [ "A6:C7" ],
              "freeze": { "rows": 1, "cols": 1 },
              "filter": { "range": "A1:D3" },
              "cf": [ { "range": "B2:C2", "type": "cell", "op": "gt", "v1": "100", "style": { "color": "#9c0006", "fill": "#ffc7ce" } },
                      { "range": "A2:A9", "type": "scale", "colors": [ "#f8696b", "#ffeb84", "#63be7b" ] } ],
              "validations": [ { "range": "E2:E9", "type": "list", "v1": "Open,Done", "blank": true, "dropdown": true, "error": "Pick one", "errorStyle": "stop" } ],
              "zoom": 1.25, "grid": false, "print": { "landscape": true, "fit": "width", "titles": "1:1" } },
            { "id": "s2", "name": "Second sheet", "hidden": false, "cells": { "B2": { "v": "second" } }, "colW": 80, "rowH": 24 },
            { "id": "s3", "name": "Hidden", "hidden": true, "cells": {} }
          ]
        }
        """;
        var bytes = Handler.Export(content, ".xlsx", "Book");
        var back = Parse(Handler.Normalize(Handler.Import(new MemoryStream(bytes), ".xlsx")).Json!);

        Assert.Equal(1, (int)back["active"]!);
        var sheets = back["sheets"]!.AsArray();
        Assert.Equal(["Data", "Second sheet", "Hidden"], sheets.Select(s => (string?)s!["name"]));
        Assert.True((bool)sheets[2]!["hidden"]!);
        Assert.Equal("#c00000", (string?)sheets[0]!["color"]);
        Assert.Equal(80, (double)sheets[1]!["colW"]!);
        Assert.Equal(24, (double)sheets[1]!["rowH"]!);
        Assert.Equal("Data!$B$2", (string?)Assert.Single(back["names"]!.AsArray())!["ref"]);

        var styles = back["styles"]!.AsArray();
        var data = sheets[0]!.AsObject();
        var cells = data["cells"]!.AsObject();
        JsonObject StyleOf(string address) => styles[(int)cells[address]!["s"]!]!.AsObject();

        Assert.Equal("Header", (string?)cells["A1"]!["v"]);
        var header = StyleOf("A1");
        Assert.True((bool)header["b"]!);
        Assert.Equal("#107c41", (string?)header["fill"]);
        Assert.Equal("#ffffff", (string?)header["color"]);
        Assert.Equal("thin #000000", (string?)header["bt"]);
        Assert.Equal("double #ff0000", (string?)header["bb"]);
        Assert.Equal("center", (string?)header["ha"]);
        Assert.Equal("middle", (string?)header["va"]);
        Assert.True((bool)header["wrap"]!);

        Assert.Equal("Ünïcødé €", (string?)cells["B1"]!["v"]);
        Assert.Contains("A note\nwith two lines", (string?)cells["B1"]!["note"], StringComparison.Ordinal);
        Assert.Equal(46000, (double)cells["A2"]!["v"]!);
        Assert.Equal("dd.mm.yyyy", (string?)StyleOf("A2")["nf"]);
        var number = StyleOf("C2");
        Assert.Equal("#,##0.00", (string?)number["nf"]);
        Assert.True((bool)number["i"]! && (bool)number["u"]! && (bool)number["strike"]!);
        Assert.Equal("Arial", (string?)number["font"]);
        Assert.Equal(14, (double)number["size"]!);
        Assert.True((bool)cells["D2"]!["v"]!);

        Assert.Equal("SUM(B2:C2)", (string?)cells["A3"]!["f"]);
        Assert.Equal(1234.75, (double)cells["A3"]!["v"]!);
        Assert.Equal("1/0", (string?)cells["B3"]!["f"]);
        Assert.Equal("#DIV/0!", (string?)cells["B3"]!["v"]);
        Assert.Equal("e", (string?)cells["B3"]!["t"]);
        Assert.Equal("yes", (string?)cells["C3"]!["v"]);
        Assert.Equal("XLOOKUP(1,B2:C2,B2:C2)", (string?)cells["D3"]!["f"]);
        Assert.Null(cells["E3"]!["f"]);
        Assert.Equal(5, (double)cells["E3"]!["v"]!);
        Assert.True((bool)cells["F3"]!["v"]!);

        Assert.Equal("https://example.com/", (string?)cells["A4"]!["link"]);
        Assert.Equal("#'Second sheet'!B2", (string?)cells["B4"]!["link"]);
        Assert.Equal(2, (int)StyleOf("C5")["indent"]!);
        Assert.Equal("#fff2cc", (string?)StyleOf("C5")["fill"]);

        Assert.Equal(120, (double)data["cols"]!["A"]!["w"]!);
        Assert.True((bool)data["cols"]!["C"]!["hidden"]!);
        Assert.Equal("#fff2cc", (string?)styles[(int)data["cols"]!["D"]!["s"]!]!["fill"]);
        Assert.Equal(30, (double)data["rows"]!["1"]!["h"]!);
        Assert.True((bool)data["rows"]!["5"]!["hidden"]!);
        Assert.Equal("A6:C7", (string?)Assert.Single(data["merges"]!.AsArray()));
        Assert.Equal(1, (int)data["freeze"]!["rows"]!);
        Assert.Equal(1, (int)data["freeze"]!["cols"]!);
        Assert.Equal("A1:D3", (string?)data["filter"]!["range"]);
        Assert.Equal(1.25, (double)data["zoom"]!);
        Assert.False((bool)data["grid"]!);
        Assert.True((bool)data["print"]!["landscape"]!);
        Assert.Equal("width", (string?)data["print"]!["fit"]);
        Assert.Equal("1:1", (string?)data["print"]!["titles"]);

        var rules = data["cf"]!.AsArray();
        Assert.Equal("cell", (string?)rules[0]!["type"]);
        Assert.Equal("gt", (string?)rules[0]!["op"]);
        Assert.Equal("100", (string?)rules[0]!["v1"]);
        Assert.Equal("#ffc7ce", (string?)rules[0]!["style"]!["fill"]);
        Assert.Equal("scale", (string?)rules[1]!["type"]);
        Assert.Equal(3, rules[1]!["colors"]!.AsArray().Count);
        var validation = Assert.Single(data["validations"]!.AsArray())!;
        Assert.Equal("list", (string?)validation["type"]);
        Assert.Equal("Open,Done", (string?)validation["v1"]);
        Assert.Equal("Pick one", (string?)validation["error"]);

        Assert.Equal("second", (string?)sheets[1]!["cells"]!["B2"]!["v"]);
    }

    [Fact]
    public void Xlsx_export_writes_the_cached_values_of_formulas()
    {
        var content = """{ "sheets": [ { "name": "S", "cells": { "A1": { "v": 2 }, "A2": { "f": "A1*21", "v": 42 }, "A3": { "f": "\"a\"&\"b\"", "v": "ab" } } } ] }""";
        var bytes = Handler.Export(content, ".xlsx", "x");
        using var wb = new XLWorkbook(new MemoryStream(bytes));
        var ws = wb.Worksheet(1);
        Assert.Equal("A1*21", ws.Cell("A2").FormulaA1);
        Assert.Equal(42, ws.Cell("A2").CachedValue.GetNumber());
        Assert.Equal("ab", ws.Cell("A3").CachedValue.GetText());
    }

    [Fact]
    public void Xlsx_import_reads_a_real_workbook_quickly()
    {
        using var source = new XLWorkbook();
        var ws = source.AddWorksheet("Report");
        ws.Cell("A1").Value = "Repeated";
        ws.Cell("A2").Value = "Repeated";
        ws.Cell("A3").Value = "Repeated";
        ws.Cell("B1").Value = new DateTime(2026, 1, 31);
        ws.Cell("B2").Value = XLError.DivisionByZero;
        ws.Cell("B3").Value = 10;
        ws.Cell("B4").FormulaA1 = "B3*2";
        ws.Cell("B5").FormulaA1 = "_xlfn.XLOOKUP(10,B3:B4,B3:B4)";
        ws.Cell("C1").Value = "Bold";
        ws.Cell("C1").Style.Font.Bold = true;
        ws.Cell("C2").Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
        ws.Style.Font.Italic = true; // the whole sheet (A1:XFD1048576): <col min="1" max="16384" style=…>
        ws.Range("H1:H50000").Style.Fill.BackgroundColor = XLColor.LightGray; // 50,000 cells with a format and nothing else
        ws.Column("F").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Row(8).Style.Font.Bold = true;
        ws.Row(9).Height = 30;
        ws.Column("E").Width = 20;
        ws.SheetView.FreezeRows(1);
        source.AddWorksheet("Second").Cell("A1").Value = "x";
        using var file = new MemoryStream();
        source.SaveAs(file);

        var watch = Stopwatch.StartNew();
        var raw = Handler.Import(new MemoryStream(file.ToArray()), ".xlsx");
        watch.Stop();
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(10), $"import took {watch.Elapsed}");

        var check = Handler.Normalize(raw);
        Assert.True(check.Ok, check.Error);
        var content = Parse(check.Json!);
        var sheet = content["sheets"]![0]!.AsObject();
        var cells = sheet["cells"]!.AsObject();
        Assert.InRange(cells.Count, 50_000, 50_100);
        Assert.Equal("#d3d3d3", (string?)content["styles"]![(int)cells["H50000"]!["s"]!]!["fill"]);
        Assert.Equal("Repeated", (string?)cells["A3"]!["v"]);
        Assert.Equal(46053, (double)cells["B1"]!["v"]!);
        var styles = content["styles"]!.AsArray();
        Assert.Equal(DateKind.Date, NumberFormats.Classify((string?)styles[(int)cells["B1"]!["s"]!]!["nf"]));
        Assert.Equal("#DIV/0!", (string?)cells["B2"]!["v"]);
        Assert.Equal("e", (string?)cells["B2"]!["t"]);
        Assert.Equal("B3*2", (string?)cells["B4"]!["f"]);
        Assert.Null(cells["B4"]!["v"]); // ClosedXML wrote no cached value
        Assert.Equal("XLOOKUP(10,B3:B4,B3:B4)", (string?)cells["B5"]!["f"]);
        Assert.True((bool)styles[(int)cells["C1"]!["s"]!]!["b"]!);
        Assert.Matches("^#[0-9a-f]{6}$", (string?)styles[(int)cells["C2"]!["s"]!]!["fill"]);
        Assert.Equal("#ffff00", (string?)styles[(int)sheet["cols"]!["F"]!["s"]!]!["fill"]);
        Assert.True((bool)styles[(int)sheet["rows"]!["8"]!["s"]!]!["b"]!);
        Assert.Equal(40, (double)sheet["rows"]!["9"]!["h"]!);
        Assert.Equal(145, (double)sheet["cols"]!["E"]!["w"]!);
        Assert.Equal(1, (int)sheet["freeze"]!["rows"]!);
        Assert.Equal(2, content["sheets"]!.AsArray().Count);
    }

    [Fact]
    public void Xlsx_import_refuses_a_broken_file()
    {
        var error = Assert.Throws<Services.Formats.FormatException>(() => Languages.In("en", () => Handler.Import(new MemoryStream([1, 2, 3, 4]), ".xlsx")));
        Assert.Equal(Languages.In("en", () => Text.T("Documents.Import.Unreadable")), error.Message);
        Assert.False(Kinds.Read(Handler, new MemoryStream("PK\u0003\u0004 not a zip"u8.ToArray()), ".xlsx").Ok);
    }

    [Fact]
    public void Every_template_can_be_written_as_xlsx_and_read_again()
    {
        foreach (var template in Handler.Templates)
        {
            var content = Handler.Normalize(Handler.Create(template.Id, "de")).Json!;
            var bytes = Handler.Export(content, ".xlsx", template.Id);
            var check = Handler.Normalize(Handler.Import(new MemoryStream(bytes), ".xlsx"));
            Assert.True(check.Ok, check.Error);
            var before = Parse(content)["sheets"]![0]!["cells"]!.AsObject();
            var after = Parse(check.Json!)["sheets"]![0]!["cells"]!.AsObject();
            foreach (var (address, cell) in before)
            {
                Assert.Equal((string?)cell!["f"], (string?)after[address]?["f"]);
            }
        }
    }

    // ---- csv and tsv --------------------------------------------------------------------------------------------------------------------

    [Fact]
    public void Csv_round_trip_keeps_quotes_line_breaks_and_unicode()
    {
        var content = """
        { "styles": [ {}, { "nf": "yyyy-mm-dd" }, { "nf": "dd.mm.yyyy hh:mm" } ],
          "sheets": [ { "name": "S", "cells": {
            "A1": { "v": "Name" }, "B1": { "v": "Note" }, "C1": { "v": "Amount" }, "D1": { "v": "Date" },
            "A2": { "v": "Müller, Anna" }, "B2": { "v": "Says \"hi\"\nand leaves" }, "C2": { "v": 1234.5 }, "D2": { "v": 46053, "s": 1 },
            "A3": { "v": "日本語 🙂" }, "C3": { "f": "C2*2", "v": 2469 }, "D3": { "v": 46053.5, "s": 2 },
            "A5": { "v": true }, "B5": { "v": "#N/A", "t": "e" }, "C5": { "v": 0.30000000000000004 }
          } } ] }
        """;
        var bytes = Handler.Export(content, ".csv", "x");
        Assert.Equal([0xEF, 0xBB, 0xBF], bytes[..3]);
        var text = Encoding.UTF8.GetString(bytes[3..]);
        Assert.Contains("\"Müller, Anna\",\"Says \"\"hi\"\"\nand leaves\",1234.5,2026-01-31\r\n", text, StringComparison.Ordinal);
        Assert.Contains("2026-01-31 12:00:00", text, StringComparison.Ordinal);
        Assert.Contains("TRUE,#N/A,0.3,\r\n", text, StringComparison.Ordinal);
        Assert.Contains("\r\n,,,\r\n", text, StringComparison.Ordinal);

        var back = Parse(Languages.In("en", () => Handler.Normalize(Handler.Import(new MemoryStream(bytes), ".csv")).Json!));
        var sheet = back["sheets"]![0]!;
        Assert.Equal("Sheet1", (string?)sheet["name"]);
        var cells = sheet["cells"]!.AsObject();
        Assert.Equal("Müller, Anna", (string?)cells["A2"]!["v"]);
        Assert.Equal("Says \"hi\"\nand leaves", (string?)cells["B2"]!["v"]);
        Assert.Equal(1234.5, (double)cells["C2"]!["v"]!);
        Assert.Equal(46053, (double)cells["D2"]!["v"]!);
        Assert.Equal("yyyy-mm-dd", (string?)back["styles"]![(int)cells["D2"]!["s"]!]!["nf"]);
        Assert.Equal(46053.5, (double)cells["D3"]!["v"]!, 6);
        Assert.Equal("日本語 🙂", (string?)cells["A3"]!["v"]);
        Assert.Equal(2469, (double)cells["C3"]!["v"]!);
        Assert.True((bool)cells["A5"]!["v"]!);
        Assert.Equal(0.3, (double)cells["C5"]!["v"]!);
    }

    [Fact]
    public void Csv_import_reads_semicolons_and_decimal_commas()
    {
        var csv = "Artikel;Preis;Menge;Datum;Anteil;Nummer\r\nÄpfel;1.234,56;3;31.01.2026;12,5 %;00123\r\nBirnen;0,99;-2;01.02.2026;WAHR;4930123456789012\r\n";
        var bytes = Encoding.Latin1.GetBytes(csv);
        var content = Parse(Handler.Normalize(Handler.Import(new MemoryStream(bytes), ".csv")).Json!);
        var cells = content["sheets"]![0]!["cells"]!.AsObject();
        var styles = content["styles"]!.AsArray();
        Assert.Equal("Äpfel", (string?)cells["A2"]!["v"]);
        Assert.Equal(1234.56, (double)cells["B2"]!["v"]!);
        Assert.Equal(0.99, (double)cells["B3"]!["v"]!);
        Assert.Equal(-2, (double)cells["C3"]!["v"]!);
        Assert.Equal(46053, (double)cells["D2"]!["v"]!);
        Assert.Equal("dd.mm.yyyy", (string?)styles[(int)cells["D2"]!["s"]!]!["nf"]);
        Assert.Equal(0.125, (double)cells["E2"]!["v"]!);
        Assert.Equal("0.0%", (string?)styles[(int)cells["E2"]!["s"]!]!["nf"]);
        Assert.True((bool)cells["E3"]!["v"]!);
        Assert.Equal("00123", (string?)cells["F2"]!["v"]);
        Assert.Equal("4930123456789012", (string?)cells["F3"]!["v"]);
    }

    [Fact]
    public void Csv_import_detects_commas_and_dots()
    {
        var csv = "a,b,c\n\"1,234.5\",2026-01-31,3/4/2026\n12.5,7:30,x\n";
        var content = Parse(Handler.Normalize(Handler.Import(new MemoryStream(Encoding.UTF8.GetBytes(csv)), ".csv")).Json!);
        var cells = content["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal(1234.5, (double)cells["A2"]!["v"]!);
        Assert.Equal(46053, (double)cells["B2"]!["v"]!);
        Assert.Equal(46085, (double)cells["C2"]!["v"]!);
        Assert.Equal(12.5, (double)cells["A3"]!["v"]!);
        Assert.Equal(7.5 / 24, (double)cells["B3"]!["v"]!, 9);
    }

    [Fact]
    public void Tsv_import_and_export()
    {
        var tsv = "Name\tValue\nA, B\t1.5\n\"x\ty\"\t2\n";
        var content = Handler.Normalize(Handler.Import(new MemoryStream(Encoding.UTF8.GetBytes(tsv)), ".tsv")).Json!;
        var cells = Parse(content)["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal("A, B", (string?)cells["A2"]!["v"]);
        Assert.Equal(1.5, (double)cells["B2"]!["v"]!);
        Assert.Equal("x\ty", (string?)cells["A3"]!["v"]);

        var text = Encoding.UTF8.GetString(Handler.Export(content, ".tsv", "x")[3..]);
        Assert.Equal("Name\tValue\r\nA, B\t1.5\r\n\"x\ty\"\t2\r\n", text);
    }

    [Fact]
    public void Csv_export_protects_formulas_and_the_import_removes_the_protection()
    {
        var content = """{ "active": 1, "sheets": [ { "name": "A", "cells": { "A1": { "v": "not this" } } }, { "name": "B", "cells": { "A1": { "v": "=HYPERLINK(\"http://evil\")" }, "B1": { "v": "+1" }, "C1": { "v": "-x" }, "D1": { "v": "@SUM(A1)" }, "E1": { "v": "plain" } } } ] }""";
        var bytes = Handler.Export(content, ".csv", "x");
        var text = Encoding.UTF8.GetString(bytes[3..]);
        Assert.Equal("\"'=HYPERLINK(\"\"http://evil\"\")\",'+1,'-x,'@SUM(A1),plain\r\n", text);

        var cells = Parse(Handler.Normalize(Handler.Import(new MemoryStream(bytes), ".csv")).Json!)["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal("=HYPERLINK(\"http://evil\")", (string?)cells["A1"]!["v"]);
        Assert.Equal("+1", (string?)cells["B1"]!["v"]);
        Assert.Equal("-x", (string?)cells["C1"]!["v"]);
        Assert.Null(cells["A1"]!["f"]);
    }

    [Fact]
    public void Csv_reads_windows_1252_and_utf16()
    {
        var latin = Encoding.Latin1.GetBytes("Größe;Maß\nü;ß\n");
        var cells = Parse(Handler.Import(new MemoryStream(latin), ".csv"))["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal("Größe", (string?)cells["A1"]!["v"]);
        Assert.Equal("ß", (string?)cells["B2"]!["v"]);

        var utf16 = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("a\tb\n1\t2\n")).ToArray();
        cells = Parse(Handler.Import(new MemoryStream(utf16), ".csv"))["sheets"]![0]!["cells"]!.AsObject();
        Assert.Equal("b", (string?)cells["B1"]!["v"]);
        Assert.Equal(2, (double)cells["B2"]!["v"]!);
    }

    // ---- plain text --------------------------------------------------------------------------------------------------------------------

    [Fact]
    public void PlainText_lists_the_values_of_every_sheet()
    {
        var content = """{ "sheets": [ { "name": "First", "cells": { "B1": { "v": "b" }, "A1": { "v": "a" }, "A2": { "v": 1.5 }, "C3": { "f": "1=1", "v": true } } }, { "name": "Second", "cells": { "A1": { "v": "z" } } } ] }""";
        Assert.Equal("First\na\tb\n1.5\nTRUE\n\nSecond\nz", Handler.PlainText(content, 1000));
        Assert.Equal("First\na", Handler.PlainText(content, 7));
        Assert.Equal(string.Empty, Handler.PlainText("not json", 100));
        Assert.Equal(string.Empty, Handler.PlainText("""{ "sheets": 5 }""", 100));
        Assert.Contains("Monatsbudget", Handler.PlainText(Handler.Create("budget", "de"), 10_000), StringComparison.Ordinal);
    }

    // ---- helpers ------------------------------------------------------------------------------------------------------------------------

    [Theory]
    [InlineData("dd.mm.yyyy", "Date")]
    [InlineData("m/d/yyyy h:mm", "DateTime")]
    [InlineData("h:mm:ss", "Time")]
    [InlineData("[h]:mm", "Time")]
    [InlineData("mmm-yy", "Date")]
    [InlineData("#,##0.00 \"days\"", "None")]
    [InlineData("[Red]-#,##0", "None")]
    [InlineData("0.00E+00", "None")]
    [InlineData("General", "None")]
    public void Number_formats_are_classified(string code, string kind) => Assert.Equal(kind, NumberFormats.Classify(code).ToString());

    [Fact]
    public void Excel_serial_numbers_match_the_dates()
    {
        Assert.Equal(1, NumberFormats.ToSerial(new DateTime(1900, 1, 1)));
        Assert.Equal(61, NumberFormats.ToSerial(new DateTime(1900, 3, 1)));
        Assert.Equal(46053, NumberFormats.ToSerial(new DateTime(2026, 1, 31)));
        Assert.Equal(new DateTime(2026, 1, 31, 12, 0, 0), NumberFormats.FromSerial(46053.5));
        Assert.Null(NumberFormats.ToSerial(new DateTime(1899, 12, 31)));
    }

    [Fact]
    public void Cell_addresses_are_read_and_written()
    {
        Assert.Equal("XFD", A1.Column(16_384));
        Assert.Equal(28, A1.ColumnIndex("ab"));
        Assert.True(A1.TryParseCell("XFD1048576", out var r, out var c) && r == 1_048_576 && c == 16_384);
        Assert.False(A1.TryParseCell("XFE1", out _, out _));
        Assert.False(A1.TryParseCell("A01", out _, out _));
        Assert.Equal("A1:C3", A1.NormalizeArea("c3:$A$1"));
        Assert.Equal("'My sheet'!A1:B2", A1.NormalizeReference("'My sheet'!$A$1:B2"));
        Assert.Null(A1.NormalizeReference("Bad sheet!A1"));
    }

    private static JsonObject Parse(string json) => JsonNode.Parse(json)!.AsObject();
}
