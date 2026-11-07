using System.Text;
using System.Text.Json.Nodes;
using MatOffice.Data;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Base;

/// <summary>
/// MatOffice Base: desktop databases like Access – tables, relationships, queries, forms and reports in one JSON content.
/// <para>
/// The content (version 1; the editor in <c>wwwroot/js/base</c> works on exactly this, <see cref="BaseNormalizer"/> checks it):
/// <code>
/// {
///   "version": 1,
///   "settings": { "currency": "EUR" },                       // ISO code of the Currency fields
///   "tables": [{
///     "id": "t1", "name": "Customers", "description": "",
///     "fields": [{
///       "id": "f1", "name": "CustomerID", "type": "autonumber",  // text | memo | number | currency | datetime | yesno | autonumber | hyperlink
///       "size": 255,                                         // text only (1–255)
///       "numberSize": "long",                                // number only: byte | integer | long | single | double | decimal
///       "next": 5,                                           // autonumber only: the next value
///       "description": "", "format": "", "decimals": null,  // format: Currency, Short Date, Yes/No, Percent, &gt; … (Access names)
///       "default": "", "required": false, "indexed": "no",   // indexed: no | dup | unique; default: an expression (0, "Berlin", Date())
///       "caption": "", "validationRule": "", "validationText": "", "inputMask": "",
///       "lookup": null                                       // or { "kind": "table", "table": "t2", "bound": "f1", "display": "f2" }
///     }],                                                    //    or { "kind": "values", "values": ["A", "B"], "limit": true }
///     "primaryKey": ["f1"],                                  // field ids (several = a composite key)
///     "rows": [{ "id": 1, "v": { "f1": 1, "f2": "Anna" } }],  // stable record ids; values by field id, empty values left out
///     "nextId": 2,
///     "datasheet": { "sort": [{ "field": "f2", "desc": false }], "filter": "", "filterOn": false, "frozen": 0,
///                    "totals": { "f3": "sum" }, "hidden": ["f4"], "widths": { "f2": 160 }, "totalsOn": false }
///   }],
///   "relationships": [{ "id": "rel1", "parent": { "table": "t1", "field": "f1" }, "child": { "table": "t2", "field": "f2" },
///                       "integrity": true, "cascadeUpdate": false, "cascadeDelete": true, "join": 1 }],   // join: 1 inner, 2 all of parent, 3 all of child
///   "queries": [{ "id": "q1", "name": "Berlin customers", "description": "", "sql": "SELECT … ;",          // Access SQL (the design grid follows from it)
///                 "layout": { "tables": [{ "name": "Customers", "x": 20, "y": 20, "w": 170, "h": 190 }], "grid": 120 },
///                 "datasheet": { "widths": { "Name": 160 } } }],
///   "forms": [{ "id": "m1", "name": "Customers", "source": "t1" (a table or query id, "" = unbound), "caption": "Customers",
///               "kind": "single" | "continuous" | "datasheet" | "split", "width": 640,
///               "allowEdits": true, "allowAdditions": true, "allowDeletions": true,
///               "sections": { "header": Section, "detail": Section, "footer": Section } }],
///   "reports": [{ "id": "r1", "name": "Sales", "source": "q2", "caption": "Sales", "orientation": "portrait" | "landscape", "paper": "a4" | "letter",
///                 "width": 718, "grouping": [{ "field": "Category", "header": true, "footer": true, "desc": false, "keepTogether": false }],
///                 "sorting": [{ "field": "Name", "desc": false }],
///                 "sections": { "reportHeader", "pageHeader", "groupHeaders": [Section …], "detail", "groupFooters": [Section …], "pageFooter", "reportFooter" } }],
///   "ui": { "nav": { "collapsed": ["forms"], "width": 240, "hidden": false }, "relationships": { "t1": { "x": 40, "y": 30 } } }
/// }
/// Section = { "height": 300, "visible": true, "back": null, "controls": [Control] }
/// Control = { "id": "c1", "type": "label" | "textbox" | "combobox" | "checkbox" | "button" | "subform" | "line" | "rect",
///             "x", "y", "w", "h" (px), "name": "", "caption": "Name", "source": "Name" | "=[Price]*[Qty]" | "=Sum([Amount])",
///             "format": "", "decimals": null, "fontSize": 11, "bold": false, "italic": false, "align": "left" | "center" | "right",
///             "color": "#rrggbb" | null, "back": "#rrggbb" | null, "border": true, "locked": false, "visible": true, "hideDuplicates": false,
///             "action": "first" | "previous" | "next" | "last" | "new" | "save" | "delete" | "undo" | "close" | "print",
///             "rowSource": null (the lookup of the bound field) | { "kind": "table", "table": "t2", "bound": "f1", "display": "f2" } | { "kind": "values", "values": [] },
///             "subform": { "source": "t2", "linkMaster": "CustomerID", "linkChild": "CustomerID" }, "runningSum": "no" | "group" | "all" }
/// </code>
/// Dates are ISO text ("2026-11-07", "2026-11-07T10:30:00"), Yes/No is true/false, numbers and currency are JSON numbers.
/// The checks: the shape and types above, record values against their field types and sizes, unique primary keys and unique indexes,
/// required fields, referential integrity of enforced relationships, at most 200,000 records, no markup in any text.
/// </para>
/// </summary>
public sealed class BaseHandler : IKindHandler
{
    public DocumentKind Kind => DocumentKind.Base;

    public string Slug => "base";

    public string Colour => "#a4373a";

    public string Icon => "kind-base";

    public string NativeExtension => ".matbase";

    /// <summary>Excel holds only the tables of a database (no queries, forms, reports, relationships): a new database is the own format.</summary>
    public string NewFileExtension => NativeExtension;

    public IReadOnlyList<FileFormat> Formats { get; } =
    [
        new(".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Excel", CanImport: true, CanExport: true),
    ];

    public IReadOnlyList<KindTemplate> Templates => BaseTemplates.List;

    public string Create(string templateId, string language) => BaseTemplates.Build(templateId, language);

    public ContentCheck Normalize(string json)
    {
        try
        {
            return ContentCheck.Valid(BaseNormalizer.Normalize(Kinds.ParseObject(json)).ToJsonString());
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (System.Text.Json.JsonException)
        {
            return ContentCheck.Invalid(Text.T("Documents.Invalid"));
        }
        catch (InvalidOperationException)
        {
            return ContentCheck.Invalid(Text.T("Documents.Invalid"));
        }
    }

    public string Import(Stream input, string extension) => Kinds.NormalizeExtension(extension) switch
    {
        ".xlsx" => BaseXlsx.Import(input),
        _ => throw new FormatException(Text.T("Documents.Import.NotSupported")),
    };

    public byte[] Export(string content, string extension, string name) => Kinds.NormalizeExtension(extension) switch
    {
        ".xlsx" => BaseXlsx.Export(content),
        _ => throw new FormatException(Text.T("Documents.Export.NotSupported")),
    };

    /// <summary>The names of the objects and fields and the texts of the records.</summary>
    public string PlainText(string content, int max)
    {
        JsonObject root;
        try
        {
            root = Kinds.ParseObject(content);
        }
        catch (Exception ex) when (ex is FormatException or System.Text.Json.JsonException)
        {
            return string.Empty;
        }

        var text = new StringBuilder();
        void Add(string? s)
        {
            if (!string.IsNullOrWhiteSpace(s) && text.Length < max)
            {
                text.Append(s).Append(' ');
            }
        }

        foreach (var kind in new[] { "tables", "queries", "forms", "reports" })
        {
            foreach (var o in root[kind] as JsonArray ?? [])
            {
                Add(o?["name"]?.GetValue<string>());
            }
        }

        foreach (var t in root["tables"] as JsonArray ?? [])
        {
            foreach (var f in t?["fields"] as JsonArray ?? [])
            {
                Add(f?["name"]?.GetValue<string>());
            }

            foreach (var r in t?["rows"] as JsonArray ?? [])
            {
                if (text.Length >= max)
                {
                    break;
                }

                foreach (var (_, v) in r?["v"] as JsonObject ?? [])
                {
                    if (v is JsonValue value && value.TryGetValue<string>(out var s) && !BaseNormalizer.IsIsoDate(s))
                    {
                        Add(s);
                    }
                }
            }
        }

        var all = text.ToString().Trim();
        return all.Length > max ? all[..max] : all;
    }
}
