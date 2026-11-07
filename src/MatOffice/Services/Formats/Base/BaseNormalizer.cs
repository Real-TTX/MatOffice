using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Base;

/// <summary>
/// The check of a database content (see <see cref="BaseHandler"/> for the shape): only the known properties are kept (in canonical form), the values of the
/// records must fit the types of their fields, primary keys and unique indexes hold, enforced relationships have no orphans, and no text carries markup.
/// A problem of the data throws <see cref="FormatException"/> with a worded message; small things (a missing default, a stale counter) are repaired.
/// </summary>
public static partial class BaseNormalizer
{
    public const int MaxRecords = 200_000;
    public const int MaxObjects = 500;
    public const int MaxFields = 255;
    public const int MaxControls = 1000;
    public const int MaxName = 64;
    public const int MaxMemo = 65_536;
    public const int MaxSql = 65_536;

    public static readonly string[] FieldTypes = ["text", "memo", "number", "currency", "datetime", "yesno", "autonumber", "hyperlink"];
    public static readonly string[] NumberSizes = ["byte", "integer", "long", "single", "double", "decimal"];
    private static readonly string[] Indexed = ["no", "dup", "unique"];
    private static readonly string[] Totals = ["none", "sum", "avg", "count", "min", "max", "stdev", "var"];
    private static readonly string[] ControlTypes = ["label", "textbox", "combobox", "checkbox", "button", "subform", "line", "rect"];
    private static readonly string[] Actions = ["", "first", "previous", "next", "last", "new", "save", "delete", "undo", "close", "print"];
    private static readonly string[] Aligns = ["left", "center", "right"];
    private static readonly string[] FormKinds = ["single", "continuous", "datasheet", "split"];
    private static readonly string[] NavKinds = ["tables", "queries", "forms", "reports"];

    [GeneratedRegex("^[A-Za-z][A-Za-z0-9_-]{0,31}$")]
    private static partial Regex IdRegex();

    [GeneratedRegex(@"<\s*/?\s*[A-Za-z!][^<>]*>")]
    private static partial Regex TagRegex();

    [GeneratedRegex(@"<\s*/?\s*(script|iframe|object|embed|style|link|meta|svg|img|frame|base|form|math)\b[^>]*>?|(java|vb)script\s*:|data\s*:\s*text/html", RegexOptions.IgnoreCase)]
    private static partial Regex ScriptRegex();

    [GeneratedRegex("^#[0-9a-fA-F]{6}$")]
    private static partial Regex ColourRegex();

    [GeneratedRegex(@"^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?$")]
    private static partial Regex IsoRegex();

    [GeneratedRegex(@"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]")]
    private static partial Regex ControlCharRegex();

    [GeneratedRegex(@"^\s*([A-Za-z][A-Za-z0-9+.-]*):")]
    private static partial Regex SchemeRegex();

    public static JsonObject Normalize(JsonObject root)
    {
        var version = root["version"] is null ? 1 : Int(root["version"], "version", 1, 0, int.MaxValue);
        if (version > 1)
        {
            throw Fail("Base.Invalid.Version");
        }

        var settings = root["settings"] as JsonObject;
        var currency = settings?["currency"] is JsonValue c && c.TryGetValue<string>(out var cur) && Regex.IsMatch(cur, "^[A-Z]{3}$") ? cur : "EUR";

        var tableIds = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tables = new JsonArray();
        var info = new Dictionary<string, TableInfo>(StringComparer.Ordinal);
        var records = 0;
        foreach (var (node, i) in Items(root["tables"], "tables", MaxObjects))
        {
            var t = NormalizeTable(Obj(node, $"tables[{i}]"), $"tables[{i}]", out var ti);
            if (!tableIds.Add(ti.Id))
            {
                throw Shape($"tables[{i}].id");
            }

            if (!names.Add(ti.Name))
            {
                throw Fail("Base.Invalid.NameTaken", ti.Name);
            }

            records += ti.Rows.Count;
            if (records > MaxRecords)
            {
                throw Fail("Base.Invalid.TooMany", MaxRecords);
            }

            info[ti.Id] = ti;
            tables.Add(t);
        }

        // lookups point at tables and fields that exist
        foreach (var t in info.Values)
        {
            foreach (var f in t.FieldNodes)
            {
                if (f["lookup"] is JsonObject lk && (string?)lk["kind"] == "table")
                {
                    if (!info.TryGetValue((string)lk["table"]!, out var target) || !target.Fields.ContainsKey((string)lk["bound"]!))
                    {
                        f["lookup"] = null;
                    }
                    else if (!target.Fields.ContainsKey((string)lk["display"]!))
                    {
                        lk["display"] = (string)lk["bound"]!;
                    }
                }
            }
        }

        var relationships = new JsonArray();
        var relIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (node, i) in Items(root["relationships"], "relationships", 2000))
        {
            var r = NormalizeRelationship(Obj(node, $"relationships[{i}]"), $"relationships[{i}]", info);
            if (r is null)
            {
                continue;
            }

            if (!relIds.Add((string)r["id"]!))
            {
                throw Shape($"relationships[{i}].id");
            }

            relationships.Add(r);
        }

        var queries = new JsonArray();
        var queryIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (node, i) in Items(root["queries"], "queries", MaxObjects))
        {
            var q = NormalizeQuery(Obj(node, $"queries[{i}]"), $"queries[{i}]");
            if (!queryIds.Add((string)q["id"]!) || tableIds.Contains((string)q["id"]!))
            {
                throw Shape($"queries[{i}].id");
            }

            if (!names.Add((string)q["name"]!))
            {
                throw Fail("Base.Invalid.NameTaken", (string)q["name"]!);
            }

            queries.Add(q);
        }

        var forms = NormalizeObjects(root["forms"], "forms", NormalizeForm);
        var reports = NormalizeObjects(root["reports"], "reports", NormalizeReport);

        var ui = new JsonObject();
        var uiIn = root["ui"] as JsonObject;
        var nav = new JsonObject();
        var navIn = uiIn?["nav"] as JsonObject ?? [];
        {
            var collapsed = new JsonArray();
            foreach (var (k, _) in Items(navIn["collapsed"], "ui.nav.collapsed", 10))
            {
                if (k is JsonValue kv && kv.TryGetValue<string>(out var ks) && NavKinds.Contains(ks) && !collapsed.Any(x => (string?)x == ks))
                {
                    collapsed.Add(ks);
                }
            }

            nav["collapsed"] = collapsed;
            nav["width"] = Int(navIn["width"], "ui.nav.width", 240, 160, 600);
            nav["hidden"] = Bool(navIn["hidden"], "ui.nav.hidden", false);
        }

        ui["nav"] = nav;
        var positions = new JsonObject();
        if (uiIn?["relationships"] is JsonObject relIn)
        {
            foreach (var (id, pos) in relIn)
            {
                if (tableIds.Contains(id) && pos is JsonObject p)
                {
                    positions[id] = new JsonObject
                    {
                        ["x"] = Int(p["x"], "ui", 20, 0, 20000),
                        ["y"] = Int(p["y"], "ui", 20, 0, 20000),
                    };
                }
            }
        }

        ui["relationships"] = positions;

        return new JsonObject
        {
            ["version"] = 1,
            ["settings"] = new JsonObject { ["currency"] = currency },
            ["tables"] = tables,
            ["relationships"] = relationships,
            ["queries"] = queries,
            ["forms"] = forms,
            ["reports"] = reports,
            ["ui"] = ui,
        };
    }

    // ---- tables -------------------------------------------------------------------------------------------------------------

    private sealed class TableInfo
    {
        public string Id { get; init; } = string.Empty;

        public string Name { get; init; } = string.Empty;

        public Dictionary<string, JsonObject> Fields { get; } = new(StringComparer.Ordinal);

        public List<JsonObject> FieldNodes { get; } = [];

        public List<JsonObject> Rows { get; } = [];

        public List<string> PrimaryKey { get; } = [];

        public bool IsUnique(string fieldId)
            => (PrimaryKey.Count == 1 && PrimaryKey[0] == fieldId) || (Fields.TryGetValue(fieldId, out var f) && (string?)f["indexed"] == "unique");
    }

    private static JsonObject NormalizeTable(JsonObject t, string path, out TableInfo info)
    {
        var id = Id(t["id"], path + ".id");
        var name = Name(t["name"], path + ".name");
        info = new TableInfo { Id = id, Name = name };
        var fields = new JsonArray();
        var fieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var autoNumbers = 0;
        foreach (var (node, i) in Items(t["fields"], path + ".fields", MaxFields))
        {
            var f = NormalizeField(Obj(node, $"{path}.fields[{i}]"), $"{path}.fields[{i}]");
            var fid = (string)f["id"]!;
            if (info.Fields.ContainsKey(fid))
            {
                throw Shape($"{path}.fields[{i}].id");
            }

            if (!fieldNames.Add((string)f["name"]!))
            {
                throw Fail("Base.Invalid.NameTaken", name + "." + (string)f["name"]!);
            }

            if ((string?)f["type"] == "autonumber" && ++autoNumbers > 1)
            {
                throw Fail("Base.Invalid.AutoNumbers", name);
            }

            info.Fields[fid] = f;
            info.FieldNodes.Add(f);
            fields.Add(f);
        }

        var pk = new JsonArray();
        foreach (var (node, i) in Items(t["primaryKey"], path + ".primaryKey", 10))
        {
            var fid = Id(node, $"{path}.primaryKey[{i}]");
            if (!info.Fields.TryGetValue(fid, out var pf) || info.PrimaryKey.Contains(fid) || (string?)pf["type"] == "memo")
            {
                throw Shape($"{path}.primaryKey[{i}]");
            }

            info.PrimaryKey.Add(fid);
            pk.Add(fid);
        }

        // the records
        var rows = new JsonArray();
        var rowIds = new HashSet<long>();
        long maxId = 0;
        var fieldsById = info.Fields;
        var uniqueSets = new List<(string[] Ids, HashSet<string> Seen, string Label)>();
        if (info.PrimaryKey.Count > 0)
        {
            uniqueSets.Add((info.PrimaryKey.ToArray(), new HashSet<string>(StringComparer.Ordinal), string.Join(", ", info.PrimaryKey.Select(x => (string)fieldsById[x]["name"]!))));
        }

        foreach (var f in info.FieldNodes)
        {
            if ((string?)f["indexed"] == "unique" && !(info.PrimaryKey.Count == 1 && info.PrimaryKey[0] == (string)f["id"]!))
            {
                uniqueSets.Add(([(string)f["id"]!], new HashSet<string>(StringComparer.Ordinal), (string)f["name"]!));
            }
        }

        var autoMax = new Dictionary<string, long>(StringComparer.Ordinal);
        foreach (var (node, i) in Items(t["rows"], path + ".rows", MaxRecords))
        {
            var r = Obj(node, $"{path}.rows[{i}]");
            var rowId = Long(r["id"], $"{path}.rows[{i}].id", 1, long.MaxValue / 2);
            if (!rowIds.Add(rowId))
            {
                throw Shape($"{path}.rows[{i}].id");
            }

            maxId = Math.Max(maxId, rowId);
            var values = new JsonObject();
            var vin = r["v"] as JsonObject ?? (r["v"] is null ? [] : throw Shape($"{path}.rows[{i}].v"));
            foreach (var f in info.FieldNodes)
            {
                var fid = (string)f["id"]!;
                var value = NormalizeValue(f, vin[fid], name, rowId);
                if (value is not null)
                {
                    values[fid] = value;
                }

                var required = (bool)f["required"]! || info.PrimaryKey.Contains(fid);
                if (required && value is null)
                {
                    throw Fail("Base.Invalid.Required", name, (string)f["name"]!, rowId);
                }

                if ((string?)f["type"] == "autonumber" && value is not null)
                {
                    autoMax[fid] = Math.Max(autoMax.GetValueOrDefault(fid), (long)value.GetValue<double>());
                }
            }

            foreach (var (ids, seen, label) in uniqueSets)
            {
                var key = KeyOf(values, ids);
                if (key is not null && !seen.Add(key))
                {
                    throw Fail("Base.Invalid.Duplicate", name, label);
                }
            }

            var row = new JsonObject { ["id"] = rowId, ["v"] = values };
            info.Rows.Add(row);
            rows.Add(row);
        }

        foreach (var f in info.FieldNodes.Where(f => (string?)f["type"] == "autonumber"))
        {
            var next = f["next"]!.GetValue<long>();
            f["next"] = Math.Max(next, autoMax.GetValueOrDefault((string)f["id"]!) + 1);
        }

        var nextId = t["nextId"] is null ? maxId + 1 : Math.Max(Long(t["nextId"], path + ".nextId", 1, long.MaxValue / 2), maxId + 1);
        return new JsonObject
        {
            ["id"] = id,
            ["name"] = name,
            ["description"] = Text(t["description"], path + ".description", 1000),
            ["fields"] = fields,
            ["primaryKey"] = pk,
            ["rows"] = rows,
            ["nextId"] = nextId,
            ["datasheet"] = Datasheet(t["datasheet"] as JsonObject, path + ".datasheet", info.Fields.Keys.ToHashSet(StringComparer.Ordinal)),
        };
    }

    private static JsonObject NormalizeField(JsonObject f, string path)
    {
        var type = Enum(f["type"], path + ".type", FieldTypes, null);
        var o = new JsonObject
        {
            ["id"] = Id(f["id"], path + ".id"),
            ["name"] = Name(f["name"], path + ".name"),
            ["type"] = type,
        };
        if (type == "text")
        {
            o["size"] = Int(f["size"], path + ".size", 255, 1, 255);
        }

        if (type == "number")
        {
            o["numberSize"] = Enum(f["numberSize"], path + ".numberSize", NumberSizes, "long");
        }

        if (type == "autonumber")
        {
            o["next"] = Long(f["next"], path + ".next", 1, int.MaxValue, 1);
        }

        o["description"] = Text(f["description"], path + ".description", 255);
        o["format"] = Text(f["format"], path + ".format", 64);
        o["decimals"] = f["decimals"] is null ? null : Int(f["decimals"], path + ".decimals", 0, 0, 15);
        o["default"] = Expr(f["default"], path + ".default", 255);
        o["required"] = Bool(f["required"], path + ".required", false);
        o["indexed"] = Enum(f["indexed"], path + ".indexed", Indexed, "no");
        o["caption"] = Text(f["caption"], path + ".caption", 255);
        o["validationRule"] = Expr(f["validationRule"], path + ".validationRule", 2048);
        o["validationText"] = Text(f["validationText"], path + ".validationText", 255);
        o["inputMask"] = Text(f["inputMask"], path + ".inputMask", 64);
        o["lookup"] = type == "autonumber" ? null : Lookup(f["lookup"] as JsonObject, path + ".lookup");
        return o;
    }

    private static JsonObject? Lookup(JsonObject? lk, string path)
    {
        if (lk is null)
        {
            return null;
        }

        var kind = Enum(lk["kind"], path + ".kind", ["table", "values"], null);
        if (kind == "table")
        {
            var bound = Id(lk["bound"], path + ".bound");
            return new JsonObject
            {
                ["kind"] = "table",
                ["table"] = Id(lk["table"], path + ".table"),
                ["bound"] = bound,
                ["display"] = lk["display"] is null ? bound : Id(lk["display"], path + ".display"),
            };
        }

        var values = new JsonArray();
        foreach (var (v, i) in Items(lk["values"], path + ".values", 500))
        {
            values.Add(Text(v, $"{path}.values[{i}]", 255));
        }

        return new JsonObject { ["kind"] = "values", ["values"] = values, ["limit"] = Bool(lk["limit"], path + ".limit", false) };
    }

    /// <summary>A record value against the type of its field (null: empty). Throws for a value of the wrong type.</summary>
    private static JsonNode? NormalizeValue(JsonObject field, JsonNode? v, string table, long rowId)
    {
        var type = (string)field["type"]!;
        if (v is null)
        {
            return type == "yesno" ? JsonValue.Create(false) : null;
        }

        FormatException Wrong() => Fail("Base.Invalid.Value", table, (string)field["name"]!, rowId);
        if (v is not JsonValue value)
        {
            throw Wrong();
        }

        var kind = value.GetValueKind();
        switch (type)
        {
            case "text":
            case "memo":
            case "hyperlink":
            {
                if (kind != JsonValueKind.String)
                {
                    throw Wrong();
                }

                var raw = value.GetValue<string>();
                if (type == "hyperlink" && SchemeRegex().Match(raw) is { Success: true } m
                    && !new[] { "http", "https", "mailto", "ftp", "tel" }.Contains(m.Groups[1].Value.ToLowerInvariant()))
                {
                    // a link that could run script (javascript:, data:) is not kept
                    return null;
                }

                var s = CleanText(raw, type == "text" ? int.MaxValue : MaxMemo, type == "memo");
                if (s.Length == 0)
                {
                    return null;
                }

                var max = type == "text" ? (int)field["size"]! : type == "hyperlink" ? 2048 : MaxMemo;
                if (s.Length > max)
                {
                    throw Wrong();
                }

                return s;
            }

            case "yesno":
                return kind switch
                {
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    _ => throw Wrong(),
                };
            case "datetime":
            {
                if (kind != JsonValueKind.String || !IsIsoDate(value.GetValue<string>()))
                {
                    throw Wrong();
                }

                return value.GetValue<string>();
            }

            default:
            {
                if (kind != JsonValueKind.Number)
                {
                    throw Wrong();
                }

                var n = value.GetValue<double>();
                if (!double.IsFinite(n))
                {
                    throw Wrong();
                }

                if (type == "autonumber" && (n < 1 || n > int.MaxValue || n != Math.Floor(n)))
                {
                    throw Wrong();
                }

                if (type == "number")
                {
                    var (lo, hi) = (string?)field["numberSize"] switch
                    {
                        "byte" => (0d, 255d),
                        "integer" => (-32768d, 32767d),
                        "long" => (-2147483648d, 2147483647d),
                        _ => (double.MinValue, double.MaxValue),
                    };
                    var integral = (string?)field["numberSize"] is "byte" or "integer" or "long";
                    if (n < lo || n > hi || (integral && n != Math.Floor(n)))
                    {
                        throw Wrong();
                    }
                }

                return JsonValue.Create(n);
            }
        }
    }

    /// <summary>Whether a text is an ISO date or date-time of the content ("2026-11-07", "2026-11-07T10:30:00").</summary>
    public static bool IsIsoDate(string? s)
    {
        if (s is null || IsoRegex().Match(s) is not { Success: true } m)
        {
            return false;
        }

        int P(int g) => m.Groups[g].Success ? int.Parse(m.Groups[g].Value, CultureInfo.InvariantCulture) : 0;
        var (y, mo, d) = (P(1), P(2), P(3));
        return y is >= 1 and <= 9999 && mo is >= 1 and <= 12 && d >= 1 && d <= DateTime.DaysInMonth(y, mo) && P(4) <= 23 && P(5) <= 59 && P(6) <= 59;
    }

    private static string? KeyOf(JsonObject values, string[] ids)
    {
        var parts = new List<string>(ids.Length);
        foreach (var id in ids)
        {
            var v = values[id];
            if (v is null)
            {
                return null;
            }

            var kind = v.GetValueKind();
            parts.Add(kind switch
            {
                JsonValueKind.String => "s" + v.GetValue<string>().ToLowerInvariant(),
                JsonValueKind.Number => "n" + v.GetValue<double>().ToString("R", CultureInfo.InvariantCulture),
                JsonValueKind.True => "n-1",
                JsonValueKind.False => "n0",
                _ => "?",
            });
        }

        return string.Join('\u0001', parts);
    }

    private static JsonObject Datasheet(JsonObject? d, string path, HashSet<string>? fieldIds)
    {
        var o = new JsonObject();
        d ??= [];

        bool Known(string key) => fieldIds is null ? key.Length is > 0 and <= MaxName : fieldIds.Contains(key);
        var sort = new JsonArray();
        foreach (var (s, i) in Items(d["sort"], path + ".sort", 10))
        {
            var so = Obj(s, $"{path}.sort[{i}]");
            var field = Text(so["field"], $"{path}.sort[{i}].field", MaxName);
            if (Known(field))
            {
                sort.Add(new JsonObject { ["field"] = field, ["desc"] = Bool(so["desc"], path, false) });
            }
        }

        o["sort"] = sort;
        o["filter"] = Expr(d["filter"], path + ".filter", 4000);
        o["filterOn"] = Bool(d["filterOn"], path + ".filterOn", false);
        o["frozen"] = Int(d["frozen"], path + ".frozen", 0, 0, MaxFields);
        o["totalsOn"] = Bool(d["totalsOn"], path + ".totalsOn", false);
        var totals = new JsonObject();
        if (d["totals"] is JsonObject tin)
        {
            foreach (var (k, v) in tin)
            {
                if (Known(k) && v is JsonValue tv && tv.TryGetValue<string>(out var fn) && Totals.Contains(fn))
                {
                    totals[CleanText(k, MaxName)] = fn;
                }
            }
        }

        o["totals"] = totals;
        var hidden = new JsonArray();
        foreach (var (h, _) in Items(d["hidden"], path + ".hidden", MaxFields))
        {
            if (h is JsonValue hv && hv.TryGetValue<string>(out var hs) && Known(hs))
            {
                hidden.Add(CleanText(hs, MaxName));
            }
        }

        o["hidden"] = hidden;
        var widths = new JsonObject();
        if (d["widths"] is JsonObject win)
        {
            foreach (var (k, v) in win)
            {
                if (Known(k) && v is JsonValue)
                {
                    widths[CleanText(k, MaxName)] = Int(v, path + ".widths", 120, 20, 2000);
                }
            }
        }

        o["widths"] = widths;
        return o;
    }

    // ---- relationships, queries -------------------------------------------------------------------------------------------

    private static JsonObject? NormalizeRelationship(JsonObject r, string path, Dictionary<string, TableInfo> tables)
    {
        var parent = Obj(r["parent"], path + ".parent");
        var child = Obj(r["child"], path + ".child");
        var pt = Id(parent["table"], path + ".parent.table");
        var pf = Id(parent["field"], path + ".parent.field");
        var ct = Id(child["table"], path + ".child.table");
        var cf = Id(child["field"], path + ".child.field");
        if (!tables.TryGetValue(pt, out var p) || !tables.TryGetValue(ct, out var c) || !p.Fields.ContainsKey(pf) || !c.Fields.ContainsKey(cf))
        {
            // a relationship of a table or field that is gone
            return null;
        }

        var integrity = Bool(r["integrity"], path + ".integrity", false);
        if (integrity)
        {
            if (!p.IsUnique(pf))
            {
                throw Fail("Base.Invalid.NoUnique", p.Name, (string)p.Fields[pf]["name"]!);
            }

            var keys = new HashSet<string>(p.Rows.Select(row => KeyOf((JsonObject)row["v"]!, [pf])).OfType<string>(), StringComparer.Ordinal);
            foreach (var row in c.Rows)
            {
                var k = KeyOf((JsonObject)row["v"]!, [cf]);
                if (k is not null && !keys.Contains(k))
                {
                    throw Fail("Base.Invalid.Orphan", c.Name, p.Name);
                }
            }
        }

        return new JsonObject
        {
            ["id"] = Id(r["id"], path + ".id"),
            ["parent"] = new JsonObject { ["table"] = pt, ["field"] = pf },
            ["child"] = new JsonObject { ["table"] = ct, ["field"] = cf },
            ["integrity"] = integrity,
            ["cascadeUpdate"] = integrity && Bool(r["cascadeUpdate"], path + ".cascadeUpdate", false),
            ["cascadeDelete"] = integrity && Bool(r["cascadeDelete"], path + ".cascadeDelete", false),
            ["join"] = Int(r["join"], path + ".join", 1, 1, 3),
        };
    }

    private static JsonObject NormalizeQuery(JsonObject q, string path)
    {
        var layout = new JsonObject();
        var lin = q["layout"] as JsonObject;
        var boxes = new JsonArray();
        foreach (var (b, i) in Items(lin?["tables"], path + ".layout.tables", 50))
        {
            var bo = Obj(b, $"{path}.layout.tables[{i}]");
            boxes.Add(new JsonObject
            {
                ["name"] = Text(bo["name"], path + ".layout", MaxName),
                ["x"] = Int(bo["x"], path + ".layout", 20, 0, 20000),
                ["y"] = Int(bo["y"], path + ".layout", 20, 0, 20000),
                ["w"] = Int(bo["w"], path + ".layout", 170, 80, 1000),
                ["h"] = Int(bo["h"], path + ".layout", 190, 60, 1000),
            });
        }

        layout["tables"] = boxes;
        layout["grid"] = Int(lin?["grid"], path + ".layout.grid", 0, 0, 4000);
        return new JsonObject
        {
            ["id"] = Id(q["id"], path + ".id"),
            ["name"] = Name(q["name"], path + ".name"),
            ["description"] = Text(q["description"], path + ".description", 1000),
            ["sql"] = Expr(q["sql"], path + ".sql", MaxSql),
            ["layout"] = layout,
            ["datasheet"] = Datasheet(q["datasheet"] as JsonObject, path + ".datasheet", null),
        };
    }

    // ---- forms and reports ------------------------------------------------------------------------------------------------

    private static JsonArray NormalizeObjects(JsonNode? list, string path, Func<JsonObject, string, JsonObject> one)
    {
        var result = new JsonArray();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (node, i) in Items(list, path, MaxObjects))
        {
            var o = one(Obj(node, $"{path}[{i}]"), $"{path}[{i}]");
            if (!ids.Add((string)o["id"]!))
            {
                throw Shape($"{path}[{i}].id");
            }

            if (!names.Add((string)o["name"]!))
            {
                throw Fail("Base.Invalid.NameTaken", (string)o["name"]!);
            }

            result.Add(o);
        }

        return result;
    }

    private static string SourceId(JsonNode? n, string path)
    {
        var s = Text(n, path, 32);
        return s.Length == 0 || IdRegex().IsMatch(s) ? s : throw Shape(path);
    }

    private static JsonObject NormalizeForm(JsonObject f, string path)
    {
        var count = 0;
        var sin = f["sections"] as JsonObject;
        return new JsonObject
        {
            ["id"] = Id(f["id"], path + ".id"),
            ["name"] = Name(f["name"], path + ".name"),
            ["source"] = SourceId(f["source"], path + ".source"),
            ["caption"] = Text(f["caption"], path + ".caption", 255),
            ["kind"] = Enum(f["kind"], path + ".kind", FormKinds, "single"),
            ["width"] = Int(f["width"], path + ".width", 640, 100, 4000),
            ["allowEdits"] = Bool(f["allowEdits"], path + ".allowEdits", true),
            ["allowAdditions"] = Bool(f["allowAdditions"], path + ".allowAdditions", true),
            ["allowDeletions"] = Bool(f["allowDeletions"], path + ".allowDeletions", true),
            ["sections"] = new JsonObject
            {
                ["header"] = Section(sin?["header"], path + ".sections.header", ref count),
                ["detail"] = Section(sin?["detail"], path + ".sections.detail", ref count),
                ["footer"] = Section(sin?["footer"], path + ".sections.footer", ref count),
            },
        };
    }

    private static JsonObject NormalizeReport(JsonObject r, string path)
    {
        var count = 0;
        var sin = r["sections"] as JsonObject;
        var grouping = new JsonArray();
        foreach (var (g, i) in Items(r["grouping"], path + ".grouping", 10))
        {
            var go = Obj(g, $"{path}.grouping[{i}]");
            grouping.Add(new JsonObject
            {
                ["field"] = Text(go["field"], path + ".grouping", 2000, expression: true),
                ["header"] = Bool(go["header"], path, true),
                ["footer"] = Bool(go["footer"], path, false),
                ["desc"] = Bool(go["desc"], path, false),
                ["keepTogether"] = Bool(go["keepTogether"], path, false),
            });
        }

        var sorting = new JsonArray();
        foreach (var (s, i) in Items(r["sorting"], path + ".sorting", 10))
        {
            var so = Obj(s, $"{path}.sorting[{i}]");
            sorting.Add(new JsonObject { ["field"] = Text(so["field"], path + ".sorting", 2000, expression: true), ["desc"] = Bool(so["desc"], path, false) });
        }

        JsonArray Bands(JsonNode? list, string p, ref int n)
        {
            var arr = list as JsonArray;
            var outArr = new JsonArray();
            for (var i = 0; i < grouping.Count; i++)
            {
                outArr.Add(Section(arr is not null && i < arr.Count ? arr[i] : null, $"{p}[{i}]", ref n));
            }

            return outArr;
        }

        return new JsonObject
        {
            ["id"] = Id(r["id"], path + ".id"),
            ["name"] = Name(r["name"], path + ".name"),
            ["source"] = SourceId(r["source"], path + ".source"),
            ["caption"] = Text(r["caption"], path + ".caption", 255),
            ["orientation"] = Enum(r["orientation"], path + ".orientation", ["portrait", "landscape"], "portrait"),
            ["paper"] = Enum(r["paper"], path + ".paper", ["a4", "letter"], "a4"),
            ["width"] = Int(r["width"], path + ".width", 718, 100, 4000),
            ["grouping"] = grouping,
            ["sorting"] = sorting,
            ["sections"] = new JsonObject
            {
                ["reportHeader"] = Section(sin?["reportHeader"], path + ".sections.reportHeader", ref count),
                ["pageHeader"] = Section(sin?["pageHeader"], path + ".sections.pageHeader", ref count),
                ["groupHeaders"] = Bands(sin?["groupHeaders"], path + ".sections.groupHeaders", ref count),
                ["detail"] = Section(sin?["detail"], path + ".sections.detail", ref count),
                ["groupFooters"] = Bands(sin?["groupFooters"], path + ".sections.groupFooters", ref count),
                ["pageFooter"] = Section(sin?["pageFooter"], path + ".sections.pageFooter", ref count),
                ["reportFooter"] = Section(sin?["reportFooter"], path + ".sections.reportFooter", ref count),
            },
        };
    }

    private static JsonObject Section(JsonNode? node, string path, ref int count)
    {
        var s = node as JsonObject;
        var controls = new JsonArray();
        foreach (var (c, i) in Items(s?["controls"], path + ".controls", MaxControls))
        {
            if (++count > MaxControls)
            {
                throw Fail("Base.Invalid.TooManyControls", MaxControls);
            }

            controls.Add(Control(Obj(c, $"{path}.controls[{i}]"), $"{path}.controls[{i}]"));
        }

        return new JsonObject
        {
            ["height"] = Int(s?["height"], path + ".height", 0, 0, 10000),
            ["visible"] = Bool(s?["visible"], path + ".visible", true),
            ["back"] = Colour(s?["back"]),
            ["controls"] = controls,
        };
    }

    private static JsonObject Control(JsonObject c, string path)
    {
        var type = Enum(c["type"], path + ".type", ControlTypes, null);
        var o = new JsonObject
        {
            ["id"] = Id(c["id"], path + ".id"),
            ["type"] = type,
            ["x"] = Int(c["x"], path + ".x", 0, 0, 10000),
            ["y"] = Int(c["y"], path + ".y", 0, 0, 10000),
            ["w"] = Int(c["w"], path + ".w", 100, 1, 10000),
            ["h"] = Int(c["h"], path + ".h", 24, 1, 10000),
            ["name"] = Text(c["name"], path + ".name", MaxName),
            ["caption"] = Text(c["caption"], path + ".caption", 1000),
            ["source"] = Expr(c["source"], path + ".source", 2000),
            ["format"] = Text(c["format"], path + ".format", 64),
            ["decimals"] = c["decimals"] is null ? null : Int(c["decimals"], path + ".decimals", 0, 0, 15),
            ["fontSize"] = Int(c["fontSize"], path + ".fontSize", 11, 6, 96),
            ["bold"] = Bool(c["bold"], path + ".bold", false),
            ["italic"] = Bool(c["italic"], path + ".italic", false),
            ["align"] = Enum(c["align"], path + ".align", Aligns, "left"),
            ["color"] = Colour(c["color"]),
            ["back"] = Colour(c["back"]),
            ["border"] = Bool(c["border"], path + ".border", type is "textbox" or "combobox" or "subform" or "rect"),
            ["locked"] = Bool(c["locked"], path + ".locked", false),
            ["visible"] = Bool(c["visible"], path + ".visible", true),
            ["hideDuplicates"] = Bool(c["hideDuplicates"], path + ".hideDuplicates", false),
            ["runningSum"] = Enum(c["runningSum"], path + ".runningSum", ["no", "group", "all"], "no"),
            ["action"] = Enum(c["action"], path + ".action", Actions, string.Empty),
        };
        if (type == "combobox" && c["rowSource"] is JsonObject rs)
        {
            var kind = Enum(rs["kind"], path + ".rowSource.kind", ["table", "values"], null);
            if (kind == "table")
            {
                o["rowSource"] = new JsonObject
                {
                    ["kind"] = "table",
                    ["table"] = Id(rs["table"], path + ".rowSource.table"),
                    ["bound"] = Id(rs["bound"], path + ".rowSource.bound"),
                    ["display"] = rs["display"] is null ? (string)rs["bound"]! : Id(rs["display"], path + ".rowSource.display"),
                };
            }
            else
            {
                var values = new JsonArray();
                foreach (var (v, i) in Items(rs["values"], path + ".rowSource.values", 500))
                {
                    values.Add(Text(v, $"{path}.rowSource.values[{i}]", 255));
                }

                o["rowSource"] = new JsonObject { ["kind"] = "values", ["values"] = values };
            }
        }
        else
        {
            o["rowSource"] = null;
        }

        if (type == "subform")
        {
            var sf = c["subform"] as JsonObject;
            o["subform"] = new JsonObject
            {
                ["source"] = SourceId(sf?["source"], path + ".subform.source"),
                ["linkMaster"] = Text(sf?["linkMaster"], path + ".subform.linkMaster", MaxName),
                ["linkChild"] = Text(sf?["linkChild"], path + ".subform.linkChild", MaxName),
            };
        }

        return o;
    }

    // ---- small readers -----------------------------------------------------------------------------------------------------

    private static FormatException Fail(string key, params object[] args) => new(Localization.Text.T(key, args));

    private static FormatException Shape(string path) => Fail("Base.Invalid.Shape", path);

    private static JsonObject Obj(JsonNode? n, string path) => n as JsonObject ?? throw Shape(path);

    private static IEnumerable<(JsonNode? Node, int Index)> Items(JsonNode? n, string path, int max)
    {
        if (n is null)
        {
            yield break;
        }

        if (n is not JsonArray a)
        {
            throw Shape(path);
        }

        if (a.Count > max)
        {
            throw path.EndsWith(".rows", StringComparison.Ordinal) ? Fail("Base.Invalid.TooMany", MaxRecords) : Fail("Base.Invalid.TooManyItems", path, max);
        }

        for (var i = 0; i < a.Count; i++)
        {
            yield return (a[i], i);
        }
    }

    private static string Id(JsonNode? n, string path)
        => n is JsonValue v && v.TryGetValue<string>(out var s) && IdRegex().IsMatch(s) ? s : throw Shape(path);

    /// <summary>A name of an object or a field like Access allows it: 1–64 characters, no . ! ` [ ] and no markup.</summary>
    private static string Name(JsonNode? n, string path)
    {
        if (n is not JsonValue v || !v.TryGetValue<string>(out var raw))
        {
            throw Shape(path);
        }

        var s = CleanText(raw, 1000).Trim();
        if (s.Length is 0 or > MaxName || s.IndexOfAny(['.', '!', '`', '[', ']', '<', '>', '"']) >= 0)
        {
            throw Fail("Base.Invalid.Name", s.Length > 40 ? s[..40] + "…" : s);
        }

        return s;
    }

    private static string Text(JsonNode? n, string path, int max, bool expression = false)
    {
        if (n is null)
        {
            return string.Empty;
        }

        if (n is not JsonValue v || !v.TryGetValue<string>(out var s))
        {
            throw Shape(path);
        }

        var clean = expression ? CleanExpr(s) : CleanText(s, max);
        return clean.Length > max ? clean[..max] : clean;
    }

    /// <summary>An expression or SQL: never shown as HTML, so comparisons with &lt; stay; only what could run script is taken out.</summary>
    private static string Expr(JsonNode? n, string path, int max)
    {
        if (n is null)
        {
            return string.Empty;
        }

        if (n is not JsonValue v || !v.TryGetValue<string>(out var s))
        {
            throw Shape(path);
        }

        if (s.Length > max)
        {
            throw Fail("Base.Invalid.TooLong", path);
        }

        return CleanExpr(s);
    }

    private static string CleanExpr(string s) => ScriptRegex().Replace(ControlCharRegex().Replace(s, string.Empty), string.Empty);

    /// <summary>A text without markup and control characters (line breaks are kept in long texts).</summary>
    public static string CleanText(string s, int max, bool multiline = false)
    {
        var t = ControlCharRegex().Replace(s, string.Empty);
        if (!multiline)
        {
            t = t.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ').Replace('\t', ' ');
        }

        t = ScriptRegex().Replace(TagRegex().Replace(t, string.Empty), string.Empty);
        return t.Length > max ? t[..max] : t;
    }

    private static string Enum(JsonNode? n, string path, string[] allowed, string? fallback)
    {
        if (n is null && fallback is not null)
        {
            return fallback;
        }

        if (n is JsonValue v && v.TryGetValue<string>(out var s) && allowed.Contains(s))
        {
            return s;
        }

        return fallback ?? throw Shape(path);
    }

    private static bool Bool(JsonNode? n, string path, bool fallback) => n switch
    {
        null => fallback,
        JsonValue v when v.GetValueKind() == JsonValueKind.True => true,
        JsonValue v when v.GetValueKind() == JsonValueKind.False => false,
        _ => throw Shape(path),
    };

    private static int Int(JsonNode? n, string path, int fallback, int min, int max)
    {
        if (n is null)
        {
            return fallback;
        }

        if (n is not JsonValue v || v.GetValueKind() != JsonValueKind.Number)
        {
            throw Shape(path);
        }

        var d = v.GetValue<double>();
        return double.IsFinite(d) ? (int)Math.Clamp(Math.Round(d), min, max) : throw Shape(path);
    }

    private static long Long(JsonNode? n, string path, long min, long max, long? fallback = null)
    {
        if (n is null && fallback is not null)
        {
            return fallback.Value;
        }

        if (n is not JsonValue v || v.GetValueKind() != JsonValueKind.Number)
        {
            throw Shape(path);
        }

        var d = v.GetValue<double>();
        if (!double.IsFinite(d) || d != Math.Floor(d) || d < min || d > max)
        {
            throw Shape(path);
        }

        return (long)d;
    }

    private static string? Colour(JsonNode? n) => n is JsonValue v && v.TryGetValue<string>(out var s) && ColourRegex().IsMatch(s) ? s.ToLowerInvariant() : null;
}
