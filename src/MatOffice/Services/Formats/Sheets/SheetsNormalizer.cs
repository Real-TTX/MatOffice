using System.Text.Json.Nodes;
using MatOffice.Localization;
using static MatOffice.Services.Formats.Sheets.SheetsJson;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// The check of a workbook (docs/EDITORS.md, "Sheets content"): it builds a new object with only the documented fields of the documented types and drops
/// what does not fit. Only content that is not a workbook at all (no sheets, too many sheets or cells) is refused with a <see cref="FormatException"/>.
/// </summary>
internal static class SheetsNormalizer
{
    private static readonly HashSet<string> BorderStyles = ["thin", "medium", "thick", "dashed", "dotted", "double", "hair"];
    private static readonly HashSet<string> CellOps = ["gt", "ge", "lt", "le", "eq", "ne", "between", "notBetween"];
    private static readonly HashSet<string> TextOps = ["contains", "notContains", "begins", "ends"];
    private static readonly HashSet<string> FilterOps = ["eq", "ne", "gt", "ge", "lt", "le", "begins", "ends", "contains", "notContains"];
    private static readonly HashSet<string> ValidationTypes = ["list", "whole", "decimal", "date", "time", "length", "custom"];
    private static readonly HashSet<string> ChartTypes = ["column", "bar", "line", "area", "pie", "doughnut", "scatter"];
    private static readonly HashSet<string> Legends = ["right", "left", "top", "bottom", "none"];

    public static JsonObject Normalize(JsonObject root)
    {
        if (root["sheets"] is not JsonArray rawSheets || rawSheets.Count == 0)
        {
            throw new FormatException(Text.T("Sheets.Invalid.NoSheets"));
        }

        if (rawSheets.Count > SheetsLimits.MaxSheets)
        {
            throw new FormatException(Text.T("Sheets.Invalid.TooManySheets"));
        }

        var cellCount = rawSheets.OfType<JsonObject>().Sum(s => (s["cells"] as JsonObject)?.Count ?? 0);
        if (cellCount > SheetsLimits.MaxCells)
        {
            throw new FormatException(Text.T("Sheets.Invalid.TooManyCells"));
        }

        var styles = Styles(root["styles"]);
        var sheetObjects = rawSheets.OfType<JsonObject>().ToList();
        if (sheetObjects.Count == 0)
        {
            throw new FormatException(Text.T("Sheets.Invalid.NoSheets"));
        }

        var ids = new UniqueIds();
        foreach (var s in sheetObjects)
        {
            ids.Reserve(Id(Str(s["id"])));
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var objectIds = new UniqueIds();
        var sheets = new JsonArray();
        for (var i = 0; i < sheetObjects.Count; i++)
        {
            sheets.Add(Sheet(sheetObjects[i], i, styles.Count, ids, names, objectIds));
        }

        var active = Int(root["active"], 0, sheets.Count - 1) ?? 0;
        if (sheets.All(s => IsTrue(s!["hidden"])))
        {
            sheets[active]!.AsObject().Remove("hidden");
        }

        if (IsTrue(sheets[active]!["hidden"]))
        {
            active = sheets.Select((s, i) => (s, i)).First(p => !IsTrue(p.s!["hidden"])).i;
        }

        var result = new JsonObject
        {
            ["version"] = 1,
            ["active"] = active,
            ["styles"] = new JsonArray([.. styles]),
        };
        if (Names(root["names"]) is { Count: > 0 } definedNames)
        {
            result["names"] = definedNames;
        }

        result["sheets"] = sheets;
        return result;
    }

    // ---- styles ---------------------------------------------------------------------------------------------------------------------

    private static List<JsonNode?> Styles(JsonNode? node)
    {
        var list = new List<JsonNode?> { new JsonObject() };
        if (node is JsonArray array)
        {
            foreach (var item in array.Skip(1).Take(SheetsLimits.MaxStyles - 1))
            {
                list.Add(item is JsonObject o ? Style(o, subset: false) : new JsonObject());
            }
        }

        return list;
    }

    /// <summary>A style of the style table, or (<paramref name="subset"/>) the style of a conditional format: color, fill, b, i, u, strike.</summary>
    public static JsonObject Style(JsonObject s, bool subset)
    {
        var o = new JsonObject();
        if (!subset && Font(Str(s["font"])) is { } font)
        {
            o["font"] = font;
        }

        if (!subset && Range(s["size"], 1, 409) is { } size)
        {
            o["size"] = size;
        }

        foreach (var flag in (string[])["b", "i", "u", "strike"])
        {
            if (IsTrue(s[flag]))
            {
                o[flag] = true;
            }
        }

        if (Colour(Str(s["color"])) is { } color)
        {
            o["color"] = color;
        }

        if (Colour(Str(s["fill"])) is { } fill)
        {
            o["fill"] = fill;
        }

        if (subset)
        {
            return o;
        }

        if (Str(s["ha"]) is "left" or "center" or "right" or "justify")
        {
            o["ha"] = Str(s["ha"]);
        }

        if (Str(s["va"]) is "top" or "middle" or "bottom")
        {
            o["va"] = Str(s["va"]);
        }

        if (IsTrue(s["wrap"]))
        {
            o["wrap"] = true;
        }

        if (Int(s["indent"], 1, 15) is { } indent)
        {
            o["indent"] = indent;
        }

        if (Str(s["nf"]) is { } nf && Clean(nf, 255) is { Length: > 0 } cleanNf && !cleanNf.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            o["nf"] = cleanNf;
        }

        foreach (var edge in (string[])["bt", "bb", "bl", "br"])
        {
            if (Border(Str(s[edge])) is { } border)
            {
                o[edge] = border;
            }
        }

        if (IsTrue(s["unlocked"]))
        {
            o["unlocked"] = true;
        }

        return o;
    }

    /// <summary>A border <c>"thin #rrggbb"</c>; without a (valid) colour it is black.</summary>
    public static string? Border(string? text)
    {
        if (text is null)
        {
            return null;
        }

        var parts = text.Trim().Split(' ', 2, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || !BorderStyles.Contains(parts[0]))
        {
            return null;
        }

        return parts[0] + " " + (parts.Length > 1 ? Colour(parts[1]) ?? "#000000" : "#000000");
    }

    // ---- names ----------------------------------------------------------------------------------------------------------------------

    private static JsonArray Names(JsonNode? node)
    {
        var result = new JsonArray();
        if (node is not JsonArray array)
        {
            return result;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in array.OfType<JsonObject>())
        {
            if (result.Count >= SheetsLimits.MaxNames)
            {
                break;
            }

            var name = Str(item["name"])?.Trim();
            var reference = Formula(Str(item["ref"]));
            if (name is not null && reference is not null && IsName(name) && seen.Add(name))
            {
                result.Add(new JsonObject { ["name"] = name, ["ref"] = reference });
            }
        }

        return result;
    }

    /// <summary>A defined name as Excel takes it: a letter, _ or \ first, then letters, digits, _ . \ ?; not a cell address (A1, R1C1).</summary>
    public static bool IsName(string name)
    {
        if (name.Length is 0 or > 255 || !(char.IsLetter(name[0]) || name[0] is '_' or '\\') || !name.All(c => char.IsLetterOrDigit(c) || c is '_' or '.' or '\\' or '?'))
        {
            return false;
        }

        if (A1.TryParseCell(name, out _, out _))
        {
            return false;
        }

        var upper = name.ToUpperInvariant();
        if (upper is "R" or "C")
        {
            return false;
        }

        // R1C1-like: R, R12, C3, R1C1
        var i = 0;
        if (i < upper.Length && upper[i] == 'R')
        {
            i++;
            while (i < upper.Length && char.IsAsciiDigit(upper[i]))
            {
                i++;
            }
        }

        if (i < upper.Length && upper[i] == 'C')
        {
            i++;
            while (i < upper.Length && char.IsAsciiDigit(upper[i]))
            {
                i++;
            }
        }

        return i < upper.Length;
    }

    // ---- sheets ---------------------------------------------------------------------------------------------------------------------

    private static JsonObject Sheet(JsonObject s, int index, int styleCount, UniqueIds ids, HashSet<string> names, UniqueIds objectIds)
    {
        var o = new JsonObject
        {
            ["id"] = ids.Take(Id(Str(s["id"])), "s"),
            ["name"] = UniqueName(Str(s["name"]), index, names),
        };
        if (Colour(Str(s["color"])) is { } color)
        {
            o["color"] = color;
        }

        if (IsTrue(s["hidden"]))
        {
            o["hidden"] = true;
        }

        o["cells"] = Cells(s["cells"] as JsonObject, styleCount);
        if (Lines(s["cols"] as JsonObject, styleCount, isRow: false) is { Count: > 0 } cols)
        {
            o["cols"] = cols;
        }

        if (Lines(s["rows"] as JsonObject, styleCount, isRow: true) is { Count: > 0 } rows)
        {
            o["rows"] = rows;
        }

        if (Range(s["colW"], 1, 2000) is { } colW && Math.Round(colW) != SheetsLimits.DefaultColWidth)
        {
            o["colW"] = Math.Round(colW);
        }

        if (Range(s["rowH"], 1, 2000) is { } rowH && Math.Round(rowH) != SheetsLimits.DefaultRowHeight)
        {
            o["rowH"] = Math.Round(rowH);
        }

        if (Merges(s["merges"]) is { Count: > 0 } merges)
        {
            o["merges"] = merges;
        }

        if (s["freeze"] is JsonObject freeze)
        {
            var fr = Int(freeze["rows"], 0, A1.MaxRows - 1) ?? 0;
            var fc = Int(freeze["cols"], 0, A1.MaxCols - 1) ?? 0;
            if (fr > 0 || fc > 0)
            {
                o["freeze"] = new JsonObject { ["rows"] = fr, ["cols"] = fc };
            }
        }

        if (s["filter"] is JsonObject filter && Filter(filter) is { } f)
        {
            o["filter"] = f;
        }

        if (List(s["cf"], SheetsLimits.MaxRules, Rule) is { Count: > 0 } cf)
        {
            o["cf"] = cf;
        }

        if (List(s["validations"], SheetsLimits.MaxValidations, Validation) is { Count: > 0 } validations)
        {
            o["validations"] = validations;
        }

        if (List(s["charts"], SheetsLimits.MaxCharts, c => Chart(c, objectIds)) is { Count: > 0 } charts)
        {
            o["charts"] = charts;
        }

        if (List(s["images"], SheetsLimits.MaxImages, i => Picture(i, objectIds)) is { Count: > 0 } images)
        {
            o["images"] = images;
        }

        if (Bool(s["grid"]) == false)
        {
            o["grid"] = false;
        }

        if (Bool(s["headings"]) == false)
        {
            o["headings"] = false;
        }

        if (Range(s["zoom"], 0.1, 4) is { } zoom && zoom != 1)
        {
            o["zoom"] = Math.Round(zoom, 2);
        }

        if (s["print"] is JsonObject print && Print(print) is { Count: > 0 } p)
        {
            o["print"] = p;
        }

        if (IsTrue(s["protected"]))
        {
            o["protected"] = true;
        }

        if (A1.NormalizeArea(Str(s["sel"])) is { } sel)
        {
            o["sel"] = sel;
        }

        return o;
    }

    /// <summary>A sheet name as Excel takes it (1-31 characters, none of : \ / ? * [ ], no ' at the ends), unique in the workbook ("Sheet1 (2)").</summary>
    public static string UniqueName(string? raw, int index, HashSet<string> taken)
    {
        var name = Clean(raw ?? string.Empty, 1000).Replace('\t', ' ').Replace('\n', ' ');
        name = new string(name.Where(c => c is not (':' or '\\' or '/' or '?' or '*' or '[' or ']')).ToArray()).Trim().Trim('\'').Trim();
        name = Cut(name, 31).Trim();
        if (name.Length == 0)
        {
            name = Text.T("Sheets.SheetName", index + 1);
        }

        var candidate = name;
        for (var n = 2; !taken.Add(candidate); n++)
        {
            var suffix = $" ({n})";
            candidate = Cut(name, 31 - suffix.Length).TrimEnd() + suffix;
        }

        return candidate;
    }

    private static JsonObject Cells(JsonObject? cells, int styleCount)
    {
        var result = new JsonObject();
        if (cells is null)
        {
            return result;
        }

        foreach (var (key, value) in cells)
        {
            if (value is not JsonObject c || !A1.TryParseCell(key, out var row, out var col))
            {
                continue;
            }

            var address = A1.Cell(row, col);
            if (!result.ContainsKey(address) && Cell(c, styleCount) is { } cell)
            {
                result[address] = cell;
            }
        }

        return result;
    }

    private static JsonObject? Cell(JsonObject c, int styleCount)
    {
        var o = new JsonObject();
        if (Formula(Str(c["f"])) is { } formula)
        {
            o["f"] = formula;
        }

        var v = c["v"];
        if (Str(v) is { } text)
        {
            var clean = Clean(text, SheetsLimits.MaxText);
            o["v"] = clean;
            if (Str(c["t"]) == "e" && Errors.Contains(clean))
            {
                o["t"] = "e";
            }
        }
        else if (Num(v) is { } number)
        {
            o["v"] = number;
        }
        else if (Bool(v) is { } flag)
        {
            o["v"] = flag;
        }

        if (Int(c["s"], 0, styleCount - 1) is { } style)
        {
            o["s"] = style;
        }

        if (Link(Str(c["link"])) is { } link)
        {
            o["link"] = link;
        }

        if (Str(c["note"]) is { } note && Clean(note, SheetsLimits.MaxText) is { Length: > 0 } cleanNote)
        {
            o["note"] = cleanNote;
        }

        return o.Count > 0 ? o : null;
    }

    /// <summary>The entries of <c>cols</c> (key A … XFD) or <c>rows</c> (key 1 … 1048576): w / h, hidden, s.</summary>
    private static JsonObject Lines(JsonObject? lines, int styleCount, bool isRow)
    {
        var result = new JsonObject();
        if (lines is null)
        {
            return result;
        }

        foreach (var (key, value) in lines)
        {
            string? canonical = null;
            if (isRow && A1.TryParseRow(key, out var r))
            {
                canonical = r.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else if (!isRow && A1.ColumnIndex(key) is > 0 and var c)
            {
                canonical = A1.Column(c);
            }

            if (canonical is null || value is not JsonObject line || result.ContainsKey(canonical))
            {
                continue;
            }

            var o = new JsonObject();
            var size = isRow ? "h" : "w";
            if (Range(line[size], 0, isRow ? 2000 : 5000) is { } px)
            {
                o[size] = Math.Round(px);
            }

            if (IsTrue(line["hidden"]))
            {
                o["hidden"] = true;
            }

            if (Int(line["s"], 1, styleCount - 1) is { } style)
            {
                o["s"] = style;
            }

            if (o.Count > 0)
            {
                result[canonical] = o;
            }
        }

        return result;
    }

    private static JsonArray Merges(JsonNode? node)
    {
        var result = new JsonArray();
        if (node is not JsonArray array)
        {
            return result;
        }

        var kept = new List<Area>();
        foreach (var item in array)
        {
            if (kept.Count >= SheetsLimits.MaxMerges)
            {
                break;
            }

            if (Str(item) is { } text && A1.TryParseArea(text.Trim(), out var area) && !area.IsCell && !kept.Any(k => k.Overlaps(area)))
            {
                kept.Add(area);
                result.Add(area.ToString());
            }
        }

        return result;
    }

    private static JsonObject? Filter(JsonObject f)
    {
        if (!A1.TryParseArea(Str(f["range"]) ?? string.Empty, out var area, allowDollar: true))
        {
            return null;
        }

        var o = new JsonObject { ["range"] = area.ToString() };
        if (f["cols"] is not JsonObject cols)
        {
            return o;
        }

        var result = new JsonObject();
        foreach (var (key, value) in cols)
        {
            var col = A1.ColumnIndex(key);
            if (col < area.Col1 || col > area.Col2 || value is not JsonObject spec || result.ContainsKey(A1.Column(col)))
            {
                continue;
            }

            var entry = new JsonObject();
            if (spec["values"] is JsonArray values)
            {
                var list = new JsonArray();
                foreach (var v in values.Take(10_000))
                {
                    if (StrOrNum(v) is { } text)
                    {
                        list.Add(Clean(text, 255));
                    }
                }

                entry["values"] = list;
            }

            if (spec["cond"] is JsonObject cond && Str(cond["op"]) is { } op && FilterOps.Contains(op))
            {
                var c = new JsonObject { ["op"] = op, ["v"] = Clean(StrOrNum(cond["v"]) ?? string.Empty, 255) };
                if (Str(cond["op2"]) is { } op2 && FilterOps.Contains(op2))
                {
                    c["op2"] = op2;
                    c["v2"] = Clean(StrOrNum(cond["v2"]) ?? string.Empty, 255);
                    c["or"] = IsTrue(cond["or"]);
                }

                entry["cond"] = c;
            }

            if (entry.Count > 0)
            {
                result[A1.Column(col)] = entry;
            }
        }

        if (result.Count > 0)
        {
            o["cols"] = result;
        }

        return o;
    }

    private static JsonArray List(JsonNode? node, int max, Func<JsonObject, JsonObject?> item)
    {
        var result = new JsonArray();
        if (node is JsonArray array)
        {
            foreach (var o in array.OfType<JsonObject>())
            {
                if (result.Count >= max)
                {
                    break;
                }

                if (item(o) is { } clean)
                {
                    result.Add(clean);
                }
            }
        }

        return result;
    }

    /// <summary>A conditional format: cell, text, dup, unique, top, bottom, above, below, formula (with a style); bar, scale (colours).</summary>
    private static JsonObject? Rule(JsonObject r)
    {
        var range = A1.NormalizeArea(Str(r["range"]));
        var type = Str(r["type"]);
        if (range is null || type is null)
        {
            return null;
        }

        var o = new JsonObject { ["range"] = range, ["type"] = type };
        switch (type)
        {
            case "cell":
                var op = Str(r["op"]);
                var v1 = Value(r["v1"]);
                if (op is null || !CellOps.Contains(op) || v1 is null)
                {
                    return null;
                }

                o["op"] = op;
                o["v1"] = v1;
                if (op is "between" or "notBetween")
                {
                    o["v2"] = Value(r["v2"]) ?? v1;
                }

                break;
            case "text":
                var textOp = Str(r["op"]);
                if (textOp is null || !TextOps.Contains(textOp))
                {
                    return null;
                }

                o["op"] = textOp;
                o["v1"] = Clean(StrOrNum(r["v1"]) ?? string.Empty, 255);
                break;
            case "top" or "bottom":
                o["n"] = Int(r["n"], 1, 1000) ?? 10;
                if (IsTrue(r["percent"]))
                {
                    o["percent"] = true;
                }

                break;
            case "formula":
                if (Value(r["v1"]) is not { } condition)
                {
                    return null;
                }

                o["v1"] = condition;
                break;
            case "bar":
                o["color"] = Colour(Str(r["color"])) ?? "#638ec6";
                return o;
            case "scale":
                var colors = (r["colors"] as JsonArray)?.Select(c => Colour(Str(c))).ToList();
                if (colors is not { Count: 2 or 3 } || colors.Any(c => c is null))
                {
                    return null;
                }

                o["colors"] = new JsonArray([.. colors.Select(c => (JsonNode)c!)]);
                return o;
            case "dup" or "unique" or "above" or "below":
                break;
            default:
                return null;
        }

        o["style"] = r["style"] is JsonObject style ? Style(style, subset: true) : new JsonObject();
        return o;
    }

    /// <summary>A value or formula of a rule or a validation (text; numbers are written as text).</summary>
    private static string? Value(JsonNode? node)
        => StrOrNum(node) is { } text && Clean(text, SheetsLimits.MaxFormula).Trim() is { Length: > 0 } clean ? clean : null;

    private static JsonObject? Validation(JsonObject v)
    {
        var range = A1.NormalizeArea(Str(v["range"]));
        var type = Str(v["type"]);
        var v1 = Value(v["v1"]);
        if (range is null || type is null || !ValidationTypes.Contains(type) || v1 is null)
        {
            return null;
        }

        var o = new JsonObject { ["range"] = range, ["type"] = type };
        if (type is not ("list" or "custom"))
        {
            var op = Str(v["op"]) is { } given && CellOps.Contains(given) ? given : "between";
            o["op"] = op;
            o["v1"] = v1;
            if (op is "between" or "notBetween")
            {
                o["v2"] = Value(v["v2"]) ?? v1;
            }
        }
        else
        {
            o["v1"] = type == "list" && !v1.StartsWith('=') ? Cut(v1, 255) : v1;
        }

        foreach (var flag in (string[])["blank", "dropdown"])
        {
            if (Bool(v[flag]) is { } b)
            {
                o[flag] = b;
            }
        }

        AddText(o, "error", Str(v["error"]), 225);
        AddText(o, "errorTitle", Str(v["errorTitle"]), 32);
        if (Str(v["errorStyle"]) is "stop" or "warning" or "info")
        {
            o["errorStyle"] = Str(v["errorStyle"]);
        }

        AddText(o, "prompt", Str(v["prompt"]), 255);
        return o;
    }

    private static void AddText(JsonObject o, string key, string? text, int max)
    {
        if (text is not null && Clean(text, max) is { Length: > 0 } clean)
        {
            o[key] = clean;
        }
    }

    private static JsonObject? Chart(JsonObject c, UniqueIds ids)
    {
        var type = Str(c["type"]);
        var range = A1.NormalizeReference(Str(c["range"]));
        if (type is null || !ChartTypes.Contains(type) || range is null)
        {
            return null;
        }

        var o = new JsonObject { ["id"] = ids.Take(Id(Str(c["id"])), "c"), ["type"] = type, ["range"] = range };
        Box(c, o);
        AddText(o, "title", Str(c["title"]), 255);
        if (Str(c["legend"]) is { } legend && Legends.Contains(legend))
        {
            o["legend"] = legend;
        }

        AddText(o, "xTitle", Str(c["xTitle"]), 255);
        AddText(o, "yTitle", Str(c["yTitle"]), 255);
        if (c["colors"] is JsonArray colors)
        {
            var list = colors.Take(64).Select(x => Colour(Str(x))).Where(x => x is not null).Select(x => (JsonNode)x!).ToArray();
            if (list.Length > 0)
            {
                o["colors"] = new JsonArray(list);
            }
        }

        foreach (var flag in (string[])["byRows", "stacked", "labels"])
        {
            if (IsTrue(c[flag]))
            {
                o[flag] = true;
            }
        }

        return o;
    }

    private static JsonObject? Picture(JsonObject i, UniqueIds ids)
    {
        if (Image(Str(i["src"])) is not { } src)
        {
            return null;
        }

        var o = new JsonObject { ["id"] = ids.Take(Id(Str(i["id"])), "i"), ["src"] = src };
        Box(i, o);
        AddText(o, "alt", Str(i["alt"]), 1000);
        return o;
    }

    /// <summary>The place of a chart or picture in pixels: x, y ≥ 0, w, h 1 … 10,000.</summary>
    private static void Box(JsonObject from, JsonObject to)
    {
        to["x"] = Math.Round(Range(from["x"], 0, 10_000_000) ?? 0);
        to["y"] = Math.Round(Range(from["y"], 0, 100_000_000) ?? 0);
        to["w"] = Math.Round(Range(from["w"], 1, 10_000) ?? 320);
        to["h"] = Math.Round(Range(from["h"], 1, 10_000) ?? 240);
    }

    private static JsonObject Print(JsonObject p)
    {
        var o = new JsonObject();
        if (IsTrue(p["landscape"]))
        {
            o["landscape"] = true;
        }

        if (Str(p["fit"]) is "width" or "page")
        {
            o["fit"] = Str(p["fit"]);
        }

        if (IsTrue(p["gridlines"]))
        {
            o["gridlines"] = true;
        }

        if (IsTrue(p["headings"]))
        {
            o["headings"] = true;
        }

        if (Str(p["titles"]) is { } titles && RowSpan(titles) is { } span)
        {
            o["titles"] = span;
        }

        return o;
    }

    /// <summary>A span of rows <c>1:3</c> (also <c>$1:$3</c>), or null.</summary>
    public static string? RowSpan(string text)
    {
        var parts = text.Replace("$", string.Empty, StringComparison.Ordinal).Split(':');
        if (parts.Length != 2 || !A1.TryParseRow(parts[0].Trim(), out var r1) || !A1.TryParseRow(parts[1].Trim(), out var r2))
        {
            return null;
        }

        return $"{Math.Min(r1, r2)}:{Math.Max(r1, r2)}";
    }

    /// <summary>Ids that are unique in the workbook: valid ones are kept (first come), the others get a new one (<c>s3</c>, <c>c7</c> …).</summary>
    private sealed class UniqueIds
    {
        private readonly HashSet<string> _reserved = new(StringComparer.Ordinal);
        private readonly HashSet<string> _taken = new(StringComparer.Ordinal);
        private int _next = 1;

        public void Reserve(string? id)
        {
            if (id is not null)
            {
                _reserved.Add(id);
            }
        }

        public string Take(string? id, string prefix)
        {
            if (id is not null && _taken.Add(id))
            {
                return id;
            }

            string fresh;
            do
            {
                fresh = prefix + _next++;
            }
            while (_reserved.Contains(fresh) || !_taken.Add(fresh));

            return fresh;
        }
    }
}
