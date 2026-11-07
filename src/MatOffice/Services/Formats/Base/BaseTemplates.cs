using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Base;

/// <summary>
/// The templates of Base: the blank database and ready databases with tables, sample records, relationships, queries, forms and reports.
/// A template is written once in English (<c>Resources/Templates/base/&lt;id&gt;.json</c>) with a dictionary <c>"de"</c> that turns names, captions,
/// texts of records and the names in SQL (<c>{Customers}.{City}</c>) and its strings (<c>"Berlin"</c>) into German.
/// Spec: { de, tables: [{ name, description, fields: [[name, type, { pk, size, required, unique, indexed, default, rule, ruleText, format, decimals,
/// caption, values: [..], lookup: "Table.Bound.Display", numberSize, description }]], rows: [[…]] }], relationships: [["A.f", "B.g", "cascade"|"integrity"|"none"]],
/// queries: [{ name, sql }], forms: [{ name, source, kind, fields, subform: { source, link }, buttons }], reports: [{ name, source, fields, group, sort, sums,
/// landscape }], layout: { Table: [x, y] } }. A date "@+7" is today + 7 days.
/// </summary>
public static partial class BaseTemplates
{
    public static IReadOnlyList<KindTemplate> List { get; } =
    [
        new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-base"),
        new("contacts", "Base.Template.Contacts", "Base.Template.Contacts.Text", "users"),
        new("orders", "Base.Template.Orders", "Base.Template.Orders.Text", "store"),
        new("projects", "Base.Template.Projects", "Base.Template.Projects.Text", "list-checks"),
        new("events", "Base.Template.Events", "Base.Template.Events.Text", "calendar"),
        new("library", "Base.Template.Library", "Base.Template.Library.Text", "book"),
    ];

    private static readonly HashSet<string> Reserved = new(StringComparer.OrdinalIgnoreCase)
    {
        "AND", "OR", "NOT", "XOR", "EQV", "IMP", "MOD", "LIKE", "BETWEEN", "IN", "IS", "NULL", "TRUE", "FALSE", "YES", "NO", "SELECT", "FROM", "WHERE", "GROUP", "BY",
        "HAVING", "ORDER", "ASC", "DESC", "AS", "INNER", "LEFT", "RIGHT", "OUTER", "JOIN", "ON", "DISTINCT", "TOP", "PERCENT", "INTO", "INSERT", "VALUES", "UPDATE",
        "SET", "DELETE", "DATE", "TIME", "NOW", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "FORMAT", "SUM", "AVG", "COUNT", "MIN", "MAX", "FIRST", "LAST",
        "TABLE", "ALL", "UNION", "PARAMETERS", "LEN", "MID", "INT", "VAL", "STR",
    };

    [GeneratedRegex(@"\{([^{}""]+)\}|""((?:[^""]|"""")*)""")]
    private static partial Regex SqlTokenRegex();

    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_]*$")]
    private static partial Regex PlainNameRegex();

    /// <summary>A name as Access writes it in SQL: plain when it is a simple word, else in [brackets].</summary>
    public static string QuoteName(string name) => PlainNameRegex().IsMatch(name) && !Reserved.Contains(name) ? name : "[" + name + "]";

    public static string Blank(string currency)
        => new JsonObject
        {
            ["version"] = 1,
            ["settings"] = new JsonObject { ["currency"] = currency },
            ["tables"] = new JsonArray(),
            ["relationships"] = new JsonArray(),
            ["queries"] = new JsonArray(),
            ["forms"] = new JsonArray(),
            ["reports"] = new JsonArray(),
            ["ui"] = new JsonObject { ["nav"] = new JsonObject(), ["relationships"] = new JsonObject() },
        }.ToJsonString();

    /// <summary>The content of a new database from a template in a language (en, de).</summary>
    public static string Build(string templateId, string language)
    {
        var lang = Languages.Normalize(language);
        var currency = lang == "de" ? "EUR" : "USD";
        var spec = Load(templateId);
        if (spec is null)
        {
            return Blank(currency);
        }

        return Languages.In(lang, () => new Builder(spec, lang, currency, DateTime.Today).Build().ToJsonString());
    }

    private static JsonObject? Load(string id)
    {
        if (id == "blank" || !List.Any(t => t.Id == id))
        {
            return null;
        }

        var assembly = typeof(BaseTemplates).Assembly;
        var resource = assembly.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith(".Templates.base." + id + ".json", StringComparison.Ordinal));
        if (resource is null)
        {
            return null;
        }

        using var stream = assembly.GetManifestResourceStream(resource)!;
        return JsonNode.Parse(stream, documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true }) as JsonObject;
    }

    private sealed class Builder(JsonObject spec, string lang, string currency, DateTime today)
    {
        private readonly Dictionary<string, string> _de = (spec["de"] as JsonObject ?? []).ToDictionary(p => p.Key, p => (string)p.Value!, StringComparer.Ordinal);
        private readonly Dictionary<string, TableSpec> _tables = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _queryIds = new(StringComparer.Ordinal);

        private sealed record FieldSpec(string Id, string Name, string Type, JsonObject Node, JsonObject Options);

        private sealed record TableSpec(string Id, string Name, List<FieldSpec> Fields, JsonObject Node);

        private string Tr(string s) => lang == "de" && _de.TryGetValue(s, out var t) ? t : s;

        /// <summary>SQL or an expression of the template in the language: {Name} becomes a (translated, quoted) name, "text" a translated string.</summary>
        private string TrSql(string sql) => SqlTokenRegex().Replace(sql, m => m.Groups[1].Success
            ? QuoteName(Tr(m.Groups[1].Value))
            : "\"" + Tr(m.Groups[2].Value.Replace("\"\"", "\"", StringComparison.Ordinal)).Replace("\"", "\"\"", StringComparison.Ordinal) + "\"");

        public JsonObject Build()
        {
            var tables = new JsonArray();
            var t = 0;
            foreach (var ts in spec["tables"] as JsonArray ?? [])
            {
                var table = BuildTable((JsonObject)ts!, "t" + ++t);
                tables.Add(table.Node);
            }

            // lookups to other tables (after all tables exist)
            foreach (var table in _tables.Values)
            {
                foreach (var f in table.Fields)
                {
                    if (f.Options["lookup"] is JsonValue lv && ((string)lv!).Split('.') is [var lt, var lb, var ld])
                    {
                        var target = _tables[lt];
                        f.Node["lookup"] = new JsonObject
                        {
                            ["kind"] = "table",
                            ["table"] = target.Id,
                            ["bound"] = target.Fields.First(x => x.Name == lb).Id,
                            ["display"] = target.Fields.First(x => x.Name == ld).Id,
                        };
                    }
                }
            }

            var relationships = new JsonArray();
            var r = 0;
            foreach (var rs in spec["relationships"] as JsonArray ?? [])
            {
                var parts = rs!.AsArray().Select(x => (string)x!).ToArray();
                var (pt, pf) = Split(parts[0]);
                var (ct, cf) = Split(parts[1]);
                var integrity = parts[2] != "none";
                relationships.Add(new JsonObject
                {
                    ["id"] = "rel" + ++r,
                    ["parent"] = new JsonObject { ["table"] = pt.Id, ["field"] = pf.Id },
                    ["child"] = new JsonObject { ["table"] = ct.Id, ["field"] = cf.Id },
                    ["integrity"] = integrity,
                    ["cascadeUpdate"] = parts[2] == "cascade",
                    ["cascadeDelete"] = parts[2] == "cascade",
                    ["join"] = 1,
                });
            }

            var queries = new JsonArray();
            var q = 0;
            foreach (var qs in spec["queries"] as JsonArray ?? [])
            {
                var id = "q" + ++q;
                var name = (string)qs!["name"]!;
                _queryIds[name] = id;
                queries.Add(new JsonObject
                {
                    ["id"] = id,
                    ["name"] = Tr(name),
                    ["description"] = Tr((string?)qs["description"] ?? string.Empty),
                    ["sql"] = TrSql((string)qs["sql"]!),
                    ["layout"] = new JsonObject { ["tables"] = new JsonArray() },
                    ["datasheet"] = new JsonObject(),
                });
            }

            var forms = new JsonArray();
            var m = 0;
            foreach (var fs in spec["forms"] as JsonArray ?? [])
            {
                forms.Add(BuildForm((JsonObject)fs!, "m" + ++m));
            }

            var reports = new JsonArray();
            var p = 0;
            foreach (var rs in spec["reports"] as JsonArray ?? [])
            {
                reports.Add(BuildReport((JsonObject)rs!, "r" + ++p));
            }

            var positions = new JsonObject();
            foreach (var (name, pos) in spec["layout"] as JsonObject ?? [])
            {
                if (_tables.TryGetValue(name, out var ts) && pos is JsonArray xy)
                {
                    positions[ts.Id] = new JsonObject { ["x"] = (int)xy[0]!, ["y"] = (int)xy[1]! };
                }
            }

            return new JsonObject
            {
                ["version"] = 1,
                ["settings"] = new JsonObject { ["currency"] = currency },
                ["tables"] = tables,
                ["relationships"] = relationships,
                ["queries"] = queries,
                ["forms"] = forms,
                ["reports"] = reports,
                ["ui"] = new JsonObject { ["nav"] = new JsonObject(), ["relationships"] = positions },
            };
        }

        private (TableSpec Table, FieldSpec Field) Split(string dotted)
        {
            var i = dotted.IndexOf('.', StringComparison.Ordinal);
            var table = _tables[dotted[..i]];
            return (table, table.Fields.First(f => f.Name == dotted[(i + 1)..]));
        }

        private TableSpec BuildTable(JsonObject ts, string id)
        {
            var name = (string)ts["name"]!;
            var fields = new JsonArray();
            var specs = new List<FieldSpec>();
            var pk = new JsonArray();
            var n = 0;
            foreach (var fsNode in ts["fields"] as JsonArray ?? [])
            {
                var arr = fsNode!.AsArray();
                var fname = (string)arr[0]!;
                var type = (string)arr[1]!;
                var opt = arr.Count > 2 ? (JsonObject)arr[2]! : [];
                var fid = "f" + ++n;
                var f = new JsonObject { ["id"] = fid, ["name"] = Tr(fname), ["type"] = type };
                if (type == "text")
                {
                    f["size"] = (int?)opt["size"] ?? 255;
                }

                if (type == "number")
                {
                    f["numberSize"] = (string?)opt["numberSize"] ?? "long";
                }

                if (type == "autonumber")
                {
                    f["next"] = 1;
                }

                f["description"] = Tr((string?)opt["description"] ?? string.Empty);
                f["format"] = (string?)opt["format"] ?? (type == "currency" ? "Currency" : type == "yesno" ? "Yes/No" : string.Empty);
                f["decimals"] = (int?)opt["decimals"];
                f["default"] = opt["default"] is null ? (type == "yesno" ? "No" : string.Empty) : TrSql((string)opt["default"]!);
                f["required"] = (bool?)opt["required"] ?? false;
                f["indexed"] = (bool?)opt["unique"] == true ? "unique" : (bool?)opt["indexed"] == true ? "dup" : "no";
                f["caption"] = Tr((string?)opt["caption"] ?? string.Empty);
                f["validationRule"] = opt["rule"] is null ? string.Empty : TrSql((string)opt["rule"]!);
                f["validationText"] = Tr((string?)opt["ruleText"] ?? string.Empty);
                f["inputMask"] = string.Empty;
                f["lookup"] = null;
                if (opt["values"] is JsonArray values)
                {
                    f["lookup"] = new JsonObject { ["kind"] = "values", ["values"] = new JsonArray(values.Select(v => (JsonNode?)Tr((string)v!)).ToArray()), ["limit"] = false };
                }

                if ((bool?)opt["pk"] == true)
                {
                    pk.Add(fid);
                }

                fields.Add(f);
                specs.Add(new FieldSpec(fid, fname, type, f, opt));
            }

            var rows = new JsonArray();
            long rowId = 0;
            foreach (var rowNode in ts["rows"] as JsonArray ?? [])
            {
                var values = new JsonObject();
                var arr = rowNode!.AsArray();
                for (var i = 0; i < specs.Count && i < arr.Count; i++)
                {
                    var v = Value(specs[i], arr[i]);
                    if (v is not null)
                    {
                        values[specs[i].Id] = v;
                    }
                }

                rows.Add(new JsonObject { ["id"] = ++rowId, ["v"] = values });
            }

            foreach (var f in specs.Where(f => f.Type == "autonumber"))
            {
                var max = rows.Select(r => (double?)r!["v"]![f.Id] ?? 0).DefaultIfEmpty(0).Max();
                f.Node["next"] = (long)max + 1;
            }

            var node = new JsonObject
            {
                ["id"] = id,
                ["name"] = Tr(name),
                ["description"] = Tr((string?)ts["description"] ?? string.Empty),
                ["fields"] = fields,
                ["primaryKey"] = pk,
                ["rows"] = rows,
                ["nextId"] = rowId + 1,
                ["datasheet"] = new JsonObject(),
            };
            var table = new TableSpec(id, name, specs, node);
            _tables[name] = table;
            return table;
        }

        private JsonNode? Value(FieldSpec f, JsonNode? raw)
        {
            if (raw is null)
            {
                return f.Type == "yesno" ? false : null;
            }

            switch (f.Type)
            {
                case "datetime":
                {
                    var s = (string)raw!;
                    if (s.StartsWith('@'))
                    {
                        var parts = s[1..].Split('T');
                        var date = today.AddDays(int.Parse(parts[0], CultureInfo.InvariantCulture));
                        return parts.Length > 1 ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "T" + parts[1] : date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    }

                    return s;
                }

                case "yesno":
                    return (bool)raw!;
                case "text":
                case "memo":
                case "hyperlink":
                    return Tr((string)raw!);
                default:
                    return (double)raw!;
            }
        }

        /// <summary>The fields of a record source: its own fields, or for a query the names given (Name or Name:type).</summary>
        private List<BaseLayout.SourceField> SourceFields(string source, JsonArray? wanted)
        {
            _tables.TryGetValue(source, out var table);
            var list = new List<BaseLayout.SourceField>();
            if (wanted is null && table is not null)
            {
                foreach (var f in table.Fields)
                {
                    list.Add(Field(f));
                }

                return list;
            }

            foreach (var w in wanted ?? [])
            {
                var text = (string)w!;
                var parts = text.Split(':');
                var name = parts[0];
                var own = table?.Fields.FirstOrDefault(f => f.Name == name) ?? _tables.Values.SelectMany(x => x.Fields).FirstOrDefault(f => f.Name == name);
                if (parts.Length > 1)
                {
                    list.Add(new BaseLayout.SourceField(Tr(name), parts[1], Format: parts[1] == "currency" ? "Currency" : string.Empty));
                }
                else if (own is not null)
                {
                    list.Add(Field(own) with { Lookup = table is not null && own.Node["lookup"] is not null });
                }
                else
                {
                    list.Add(new BaseLayout.SourceField(Tr(name), "text"));
                }
            }

            return list;
        }

        private BaseLayout.SourceField Field(FieldSpec f)
            => new(Tr(f.Name), f.Type, (string?)f.Node["caption"] ?? string.Empty, f.Node["lookup"] is not null || f.Options["lookup"] is not null, (string?)f.Node["format"] ?? string.Empty);

        private string SourceId(string source) => _tables.TryGetValue(source, out var t) ? t.Id : _queryIds[source];

        private JsonObject BuildForm(JsonObject fs, string id)
        {
            var name = Tr((string)fs["name"]!);
            var source = (string)fs["source"]!;
            var fields = SourceFields(source, fs["fields"] as JsonArray);
            var title = Tr((string?)fs["title"] ?? (string)fs["name"]!);
            if ((string?)fs["kind"] == "continuous")
            {
                return BaseLayout.TabularForm(id, name, SourceId(source), title, fields);
            }

            BaseLayout.Subform? sub = null;
            if (fs["subform"] is JsonObject so)
            {
                var child = (string)so["source"]!;
                var link = (string)so["link"]!;
                var master = (string?)so["master"] ?? link;
                sub = new BaseLayout.Subform(SourceId(child), Tr((string?)so["caption"] ?? child), Tr(master), Tr(link));
            }

            var form = BaseLayout.ColumnarForm(id, name, SourceId(source), title, fields, sub, (bool?)fs["buttons"] ?? true);
            if ((string?)fs["kind"] is "split" or "datasheet")
            {
                form["kind"] = (string)fs["kind"]!;
            }

            return form;
        }

        private JsonObject BuildReport(JsonObject rs, string id)
        {
            var source = (string)rs["source"]!;
            var fields = SourceFields(source, rs["fields"] as JsonArray);
            List<string> Names(string key) => (rs[key] as JsonArray ?? []).Select(x => Tr(((string)x!).Split(':')[0])).ToList();
            return BaseLayout.SteppedReport(id, Tr((string)rs["name"]!), SourceId(source), Tr((string?)rs["title"] ?? (string)rs["name"]!), fields, Names("group"), Names("sort"),
                Names("sums"), (bool?)rs["landscape"] ?? false);
        }
    }
}
