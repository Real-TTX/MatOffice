using System.Text.Json.Nodes;
using ClosedXML.Excel;
using MatOffice.Localization;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Base;

namespace MatOffice.Tests;

/// <summary>The content check, the templates and the Excel conversion of MatOffice Base.</summary>
public sealed class BaseFormatTests
{
    private static readonly BaseHandler Handler = new();

    private static JsonObject Db(Action<JsonObject>? change = null)
    {
        var db = JsonNode.Parse("""
        {
          "version": 1, "settings": { "currency": "EUR" },
          "tables": [
            { "id": "t1", "name": "Customers", "primaryKey": ["f1"], "nextId": 3,
              "fields": [
                { "id": "f1", "name": "CustomerID", "type": "autonumber", "next": 3 },
                { "id": "f2", "name": "Name", "type": "text", "size": 50, "required": true },
                { "id": "f3", "name": "Since", "type": "datetime" },
                { "id": "f4", "name": "Vip", "type": "yesno" },
                { "id": "f5", "name": "Web", "type": "hyperlink" }
              ],
              "rows": [ { "id": 1, "v": { "f1": 1, "f2": "Anna", "f3": "2020-03-15", "f4": true } }, { "id": 2, "v": { "f1": 2, "f2": "Ben" } } ] },
            { "id": "t2", "name": "Orders", "primaryKey": ["f1"],
              "fields": [
                { "id": "f1", "name": "OrderID", "type": "autonumber" },
                { "id": "f2", "name": "CustomerID", "type": "number", "numberSize": "long" },
                { "id": "f3", "name": "Amount", "type": "currency" }
              ],
              "rows": [ { "id": 1, "v": { "f1": 1, "f2": 1, "f3": 12.5 } } ] }
          ],
          "relationships": [ { "id": "rel1", "parent": { "table": "t1", "field": "f1" }, "child": { "table": "t2", "field": "f2" }, "integrity": true, "cascadeDelete": true } ],
          "queries": [ { "id": "q1", "name": "Big orders", "sql": "SELECT * FROM Orders WHERE Amount>10;" } ],
          "forms": [], "reports": []
        }
        """)!.AsObject();
        change?.Invoke(db);
        return db;
    }

    private static ContentCheck Check(JsonObject db) => Languages.In("en", () => Handler.Normalize(db.ToJsonString()));

    private static JsonObject Rows(JsonObject db, int table) => db["tables"]![table]!["rows"]![0]!["v"]!.AsObject();

    [Fact]
    public void A_good_database_passes_and_stays_the_same_when_checked_again()
    {
        var first = Check(Db());
        Assert.True(first.Ok, first.Error);
        var second = Check(JsonNode.Parse(first.Json!)!.AsObject());
        Assert.Equal(first.Json, second.Json);
        var root = JsonNode.Parse(first.Json!)!;
        Assert.False((bool)root["tables"]![0]!["rows"]![1]!["v"]!["f4"]!);
        Assert.Equal(3, (int)root["tables"]![0]!["fields"]![0]!["next"]!);
    }

    [Fact]
    public void A_value_of_the_wrong_type_is_refused()
    {
        var check = Check(Db(db => Rows(db, 1)["f3"] = "twelve"));
        Assert.False(check.Ok);
        Assert.Contains("Amount", check.Error);
        Assert.False(Check(Db(db => Rows(db, 0)["f3"] = "15.03.2020")).Ok);
        Assert.False(Check(Db(db => Rows(db, 0)["f4"] = 1)).Ok);
        Assert.False(Check(Db(db => Rows(db, 1)["f2"] = 1.5)).Ok);
        Assert.False(Check(Db(db => Rows(db, 0)["f2"] = new string('x', 51))).Ok);
    }

    [Fact]
    public void Duplicate_primary_keys_are_refused()
    {
        var check = Check(Db(db => db["tables"]![0]!["rows"]![1]!["v"]!["f1"] = 1));
        Assert.False(check.Ok);
        Assert.Contains("Customers", check.Error);
    }

    [Fact]
    public void Unique_indexes_are_checked_without_case()
    {
        var check = Check(Db(db =>
        {
            db["tables"]![0]!["fields"]![1]!["indexed"] = "unique";
            db["tables"]![0]!["rows"]![1]!["v"]!["f2"] = "ANNA";
        }));
        Assert.False(check.Ok);
    }

    [Fact]
    public void Orphans_are_refused_where_integrity_is_enforced()
    {
        Assert.False(Check(Db(db => Rows(db, 1)["f2"] = 99)).Ok);
        Assert.True(Check(Db(db =>
        {
            Rows(db, 1)["f2"] = 99;
            db["relationships"]![0]!["integrity"] = false;
        })).Ok);
    }

    [Fact]
    public void Required_fields_need_a_value()
    {
        var check = Check(Db(db => db["tables"]![0]!["rows"]![1]!["v"]!.AsObject().Remove("f2")));
        Assert.False(check.Ok);
        Assert.Contains("Name", check.Error);
    }

    [Fact]
    public void Integrity_needs_a_unique_parent_field()
    {
        var check = Check(Db(db => db["relationships"]![0]!["parent"]!["field"] = "f2"));
        Assert.False(check.Ok);
    }

    [Fact]
    public void Markup_and_unknown_properties_are_removed()
    {
        var check = Check(Db(db =>
        {
            Rows(db, 0)["f2"] = "Anna <script>alert(1)</script><b>Schmidt</b>";
            Rows(db, 0)["f5"] = "javascript:alert(1)";
            db["tables"]![0]!["description"] = "<img src=x onerror=alert(1)>Main";
            db["tables"]![0]!["onclick"] = "alert(1)";
            db["queries"]![0]!["sql"] = "SELECT * FROM Orders WHERE Amount<10 <script>x</script>;";
        }));
        Assert.True(check.Ok, check.Error);
        var root = JsonNode.Parse(check.Json!)!;
        Assert.Equal("Anna alert(1)Schmidt", (string)root["tables"]![0]!["rows"]![0]!["v"]!["f2"]!);
        Assert.Null(root["tables"]![0]!["rows"]![0]!["v"]!["f5"]);
        Assert.Equal("Main", (string)root["tables"]![0]!["description"]!);
        Assert.Null(root["tables"]![0]!["onclick"]);
        Assert.Equal("SELECT * FROM Orders WHERE Amount<10 x;", (string)root["queries"]![0]!["sql"]!);
        Assert.DoesNotContain("<", check.Json!.Replace("Amount<10", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);
    }

    [Fact]
    public void Names_with_forbidden_characters_are_refused()
    {
        Assert.False(Check(Db(db => db["tables"]![0]!["name"] = "Cust.omers")).Ok);
        Assert.False(Check(Db(db => db["tables"]![0]!["name"] = "Orders")).Ok);
        Assert.False(Check(Db(db => db["queries"]![0]!["name"] = "Customers")).Ok);
        Assert.False(Check(Db(db => db["tables"]![0]!["fields"]![1]!["name"] = "Since")).Ok);
    }

    [Fact]
    public void Controls_of_forms_keep_only_safe_values()
    {
        var check = Check(Db(db => db["forms"] = JsonNode.Parse("""
            [{ "id": "m1", "name": "Customers", "source": "t1", "sections": { "detail": { "height": 200, "controls": [
              { "id": "c1", "type": "label", "x": 10, "y": 10, "w": 100, "h": 20, "caption": "<i>Name</i>", "color": "red", "back": "#FFEEDD", "style": "x" },
              { "id": "c2", "type": "subform", "x": 10, "y": 40, "w": 300, "h": 200, "subform": { "source": "t2", "linkMaster": "CustomerID", "linkChild": "CustomerID" } }
            ] } } }]
            """)));
        Assert.True(check.Ok, check.Error);
        var c1 = JsonNode.Parse(check.Json!)!["forms"]![0]!["sections"]!["detail"]!["controls"]![0]!;
        Assert.Equal("Name", (string)c1["caption"]!);
        Assert.Null(c1["color"]);
        Assert.Equal("#ffeedd", (string)c1["back"]!);
        Assert.Null(c1["style"]);
        Assert.False(Check(Db(db => db["forms"] = JsonNode.Parse("""[{ "id": "m1", "name": "F", "sections": { "detail": { "controls": [{ "id": "c1", "type": "iframe" }] } } }]"""))).Ok);
    }

    [Fact]
    public void A_database_holds_at_most_200000_records()
    {
        var check = Check(Db(db =>
        {
            var rows = new JsonArray();
            for (var i = 1; i <= BaseNormalizer.MaxRecords; i++)
            {
                rows.Add(new JsonObject { ["id"] = i, ["v"] = new JsonObject { ["f1"] = i, ["f3"] = 1 } });
            }

            db["tables"]![1]!["rows"] = rows;
            db["relationships"] = new JsonArray();
        }));
        Assert.False(check.Ok);
    }

    [Theory]
    [InlineData("en")]
    [InlineData("de")]
    public void Every_template_is_a_valid_database(string language)
    {
        foreach (var template in Handler.Templates)
        {
            var raw = Handler.Create(template.Id, language);
            var check = Languages.In(language, () => Handler.Normalize(raw));
            Assert.True(check.Ok, $"{template.Id} ({language}): {check.Error}");
            var root = JsonNode.Parse(check.Json!)!;
            if (template.Id == "blank")
            {
                Assert.Empty(root["tables"]!.AsArray());
                continue;
            }

            Assert.True(root["tables"]!.AsArray().Count >= 3, template.Id);
            Assert.NotEmpty(root["relationships"]!.AsArray());
            Assert.NotEmpty(root["queries"]!.AsArray());
            Assert.NotEmpty(root["forms"]!.AsArray());
            Assert.NotEmpty(root["reports"]!.AsArray());
            Assert.All(root["tables"]!.AsArray(), t => Assert.NotEmpty(t!["rows"]!.AsArray()));
            Assert.Equal(check.Json, Handler.Normalize(check.Json!).Json);
            Assert.True(Text.T(template.Name) != template.Name, template.Name);
        }
    }

    [Fact]
    public void Templates_speak_the_language_of_the_person()
    {
        var de = JsonNode.Parse(Handler.Create("orders", "de"))!;
        var en = JsonNode.Parse(Handler.Create("orders", "en"))!;
        Assert.Equal("Kunden", (string)de["tables"]![2]!["name"]!);
        Assert.Equal("Customers", (string)en["tables"]![2]!["name"]!);
        Assert.Equal("EUR", (string)de["settings"]!["currency"]!);
        var sql = (string)de["queries"]![0]!["sql"]!;
        Assert.Contains("FROM Artikel INNER JOIN Bestelldetails ON Artikel.ArtikelNr=Bestelldetails.ArtikelNr", sql);
        Assert.Contains("[Menge]", sql);
        var enSql = (string)en["queries"]![4]!["sql"]!;
        Assert.Contains("Between [Start date] And [End date]", enSql);
    }

    [Fact]
    public void Excel_round_trip_keeps_tables_types_and_values()
    {
        var content = Check(Db()).Json!;
        var bytes = Handler.Export(content, ".xlsx", "Test");
        using var book = new XLWorkbook(new MemoryStream(bytes));
        Assert.Equal(["Customers", "Orders"], book.Worksheets.Select(w => w.Name));
        Assert.Equal("Name", book.Worksheet("Customers").Cell(1, 2).GetString());

        var imported = Languages.In("en", () => Kinds.Read(Handler, new MemoryStream(bytes), ".xlsx"));
        Assert.True(imported.Ok, imported.Error);
        var root = JsonNode.Parse(imported.Json!)!;
        var customers = root["tables"]![0]!;
        Assert.Equal("Customers", (string)customers["name"]!);
        var fields = customers["fields"]!.AsArray().Select(f => ((string)f!["name"]!, (string)f["type"]!)).ToList();
        Assert.Contains(("CustomerID", "number"), fields);
        Assert.Contains(("Name", "text"), fields);
        Assert.Contains(("Since", "datetime"), fields);
        Assert.Contains(("Vip", "yesno"), fields);
        Assert.Equal(2, customers["rows"]!.AsArray().Count);
        var orders = root["tables"]![1]!;
        Assert.Contains(orders["fields"]!.AsArray(), f => (string)f!["name"]! == "Amount" && (string)f["type"]! == "currency");
        var amount = orders["fields"]!.AsArray().First(f => (string)f!["name"]! == "Amount")!["id"]!.GetValue<string>();
        Assert.Equal(12.5, (double)orders["rows"]![0]!["v"]![amount]!);
    }

    [Fact]
    public void Sheets_without_an_id_column_get_an_autonumber_key()
    {
        using var book = new XLWorkbook();
        var sheet = book.Worksheets.Add("Price list");
        sheet.Cell(1, 1).Value = "Product";
        sheet.Cell(1, 2).Value = "Price";
        sheet.Cell(1, 3).Value = "Since";
        sheet.Cell(2, 1).Value = "Coffee";
        sheet.Cell(2, 2).Value = 4.5;
        sheet.Cell(2, 3).Value = new DateTime(2026, 1, 2);
        sheet.Cell(3, 1).Value = "Tea";
        sheet.Cell(3, 2).Value = 3;
        sheet.Cell(4, 1).Value = 42;
        book.Worksheets.Add("Empty");
        using var stream = new MemoryStream();
        book.SaveAs(stream);
        stream.Position = 0;
        var check = Languages.In("en", () => Kinds.Read(Handler, stream, ".xlsx"));
        Assert.True(check.Ok, check.Error);
        var table = JsonNode.Parse(check.Json!)!["tables"]!.AsArray().Single()!;
        Assert.Equal("Price list", (string)table["name"]!);
        var types = table["fields"]!.AsArray().Select(f => (string)f!["type"]!).ToList();
        Assert.Equal(["autonumber", "text", "number", "datetime"], types);
        Assert.Equal(3, table["rows"]!.AsArray().Count);
        Assert.Equal("42", (string)table["rows"]![2]!["v"]!["f2"]!);
        Assert.Equal("2026-01-02", (string)table["rows"]![0]!["v"]!["f4"]!);
    }

    [Fact]
    public void The_plain_text_has_names_and_record_texts()
    {
        var text = Handler.PlainText(Check(Db()).Json!, 1000);
        Assert.Contains("Customers", text);
        Assert.Contains("Anna", text);
        Assert.Contains("Big orders", text);
        Assert.DoesNotContain("2020-03-15", text);
    }
}
