using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// The check of the content of a presentation (see <see cref="SlidesHandler"/> for the format): it copies only what it knows into a new object, so nothing
/// unknown survives; numbers are clamped, colours and fonts checked, texts stripped of control characters, pictures only as data: addresses of PNG, JPEG,
/// GIF or WebP (no SVG – it can carry script). Structural errors (no slides array, an element that is not an object, a picture of another type, too much
/// of anything) make the whole content invalid.
/// </summary>
public sealed partial class SlidesNormalizer
{
    public const int MaxSlides = 2000;
    public const int MaxElementsPerSlide = 1500;
    public const int MaxElements = 60000;
    public const int MaxParagraphs = 4000;
    public const int MaxRuns = 4000;
    public const int MaxRunText = 60000;
    public const int MaxNotes = 100000;
    public const int MaxRows = 200;
    public const int MaxCols = 60;
    public const int MaxAnimations = 500;
    public const int MaxDecor = 60;
    public const int MaxCharacters = 8_000_000;

    public static readonly string[] Shapes =
    [
        "rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond", "parallelogram", "trapezoid", "pentagon", "hexagon", "octagon", "plus",
        "star5", "star6", "heart", "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "chevron", "homePlate",
        "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloud", "donut", "can", "snip1Rect",
    ];

    public static readonly string[] Transitions = ["none", "fade", "push", "wipe", "split", "zoom", "cover"];

    public static readonly string[] Effects = ["appear", "fade", "fly", "zoom", "wipe"];

    private static readonly string[] Types = ["text", "shape", "line", "image", "table"];
    private static readonly string[] Arrows = ["none", "triangle", "arrow", "stealth", "oval", "diamond"];
    private static readonly string[] Dashes = ["solid", "dash", "dot", "dashDot", "longDash"];
    private static readonly string[] Directions = ["left", "right", "up", "down"];
    private static readonly string[] AnimDirections = ["bottom", "top", "left", "right"];
    private static readonly string[] Groups = ["cover", "content", "other"];

    private int _elements;
    private long _characters;

    /// <summary>Checks a content; throws <see cref="FormatException"/> with a worded message when it cannot be accepted.</summary>
    public static JsonObject Normalize(JsonObject source) => new SlidesNormalizer().Root(source);

    private static FormatException Invalid(string key = "Slides.Invalid") => new(Text.T(key));

    private JsonObject Root(JsonObject src)
    {
        if (src["slides"] is not JsonArray slides)
        {
            throw Invalid();
        }

        if (slides.Count > MaxSlides)
        {
            throw Invalid("Slides.Invalid.TooLarge");
        }

        var size = src["size"] as JsonObject;
        var w = Num(size?["w"], 120, 6000, 1280);
        var h = Num(size?["h"], 120, 6000, 720);
        var result = new JsonObject
        {
            ["version"] = 1,
            ["size"] = new JsonObject { ["w"] = w, ["h"] = h },
            ["theme"] = Theme(src["theme"] as JsonObject),
            ["footer"] = Footer(src["footer"] as JsonObject),
        };

        var list = new JsonArray();
        var slideIds = new HashSet<string>(StringComparer.Ordinal);
        var n = 0;
        foreach (var node in slides)
        {
            n++;
            if (node is not JsonObject slide)
            {
                throw Invalid();
            }

            list.Add(Slide(slide, Unique(Id(slide["id"]) ?? "s" + n, slideIds)));
        }

        result["slides"] = list;
        return result;
    }

    // ---- design ----------------------------------------------------------------------------------------------------------------

    private JsonObject Theme(JsonObject? src)
    {
        var fallback = SlidesData.Theme(src?["id"] is JsonValue v && v.TryGetValue<string>(out var id) ? id : "paper");
        if (src is null)
        {
            src = fallback;
        }

        var colors = new JsonObject();
        var srcColors = src["colors"] as JsonObject;
        foreach (var name in SlidesData.SchemeNames)
        {
            colors[name] = Hex(srcColors?[name]) ?? Hex(fallback["colors"]![name])!;
        }

        var fonts = src["fonts"] as JsonObject;
        var theme = new JsonObject
        {
            ["id"] = Id(src["id"]) ?? "custom",
            ["name"] = Str(src["name"], 100) ?? "Custom",
            ["fonts"] = new JsonObject
            {
                ["major"] = Font(fonts?["major"]) ?? Font(fallback["fonts"]!["major"])!,
                ["minor"] = Font(fonts?["minor"]) ?? Font(fallback["fonts"]!["minor"])!,
            },
            ["colors"] = colors,
            ["background"] = Fill(src["background"]) ?? new JsonObject { ["type"] = "solid", ["color"] = "lt1" },
            ["heading"] = Color(src["heading"]) ?? "dk1",
            ["text"] = Color(src["text"]) ?? "dk1",
        };
        if (Int(src["variant"], 0, 20, 0) is var variant and > 0)
        {
            theme["variant"] = variant;
        }

        if (src["groups"] is JsonObject groups)
        {
            var g = new JsonObject();
            foreach (var name in Groups)
            {
                if (groups[name] is not JsonObject one)
                {
                    continue;
                }

                var o = new JsonObject();
                if (Fill(one["background"]) is { } bg)
                {
                    o["background"] = bg;
                }

                if (Color(one["heading"]) is { } heading)
                {
                    o["heading"] = heading;
                }

                if (Color(one["text"]) is { } text)
                {
                    o["text"] = text;
                }

                if (o.Count > 0)
                {
                    g[name] = o;
                }
            }

            if (g.Count > 0)
            {
                theme["groups"] = g;
            }
        }

        var decor = new JsonArray();
        if (src["decor"] is JsonArray items)
        {
            foreach (var item in items.OfType<JsonObject>().Take(MaxDecor))
            {
                if (Decor(item) is { } d)
                {
                    decor.Add(d);
                }
            }
        }

        theme["decor"] = decor;
        return theme;
    }

    private JsonObject? Decor(JsonObject src)
    {
        var on = new JsonArray();
        if (src["on"] is JsonArray targets)
        {
            foreach (var target in targets.Take(12))
            {
                var s = Str(target, 20);
                if (s is "all" || (s is not null && (Groups.Contains(s) || SlidesData.IsLayout(s))))
                {
                    on.Add(s);
                }
            }
        }

        if (on.Count == 0)
        {
            on.Add("all");
        }

        var d = new JsonObject
        {
            ["on"] = on,
            ["shape"] = OneOf(src["shape"], Shapes, "rect"),
            ["x"] = Num(src["x"], -2, 3, 0, 4),
            ["y"] = Num(src["y"], -2, 3, 0, 4),
            ["w"] = Num(src["w"], 0, 4, 0.1, 4),
            ["h"] = Num(src["h"], 0, 4, 0.1, 4),
        };
        if (src["src"] is not null)
        {
            d["src"] = Picture(src["src"]);
        }
        else
        {
            d["fill"] = Color(src["fill"]) ?? "accent1";
        }

        CopyFlag(src, d, "flipH");
        CopyFlag(src, d, "flipV");
        if (Num(src["opacity"], 0, 1, 1) is var opacity and < 1)
        {
            d["opacity"] = opacity;
        }

        if (Num(src["rot"], -360, 360, 0) is var rot and not 0)
        {
            d["rot"] = Rotation(rot);
        }

        return d;
    }

    private static JsonObject Footer(JsonObject? src)
    {
        var f = new JsonObject
        {
            ["number"] = Bool(src?["number"]),
            ["date"] = Bool(src?["date"]),
            ["dateText"] = Str(src?["dateText"], 100) ?? string.Empty,
            ["text"] = Str(src?["text"], 300) ?? string.Empty,
            ["skipTitle"] = src?["skipTitle"] is null || Bool(src["skipTitle"]),
        };
        return f;
    }

    // ---- slides -----------------------------------------------------------------------------------------------------------------

    private JsonObject Slide(JsonObject src, string id)
    {
        var layout = Str(src["layout"], 40);
        var slide = new JsonObject
        {
            ["id"] = id,
            ["layout"] = SlidesData.IsLayout(layout) ? layout : "blank",
        };
        if (Bool(src["hidden"]))
        {
            slide["hidden"] = true;
        }

        if (Fill(src["background"]) is { } bg)
        {
            slide["background"] = bg;
        }

        if (src["transition"] is JsonObject tr && OneOf(tr["type"], Transitions, "none") is var type && type != "none")
        {
            var transition = new JsonObject { ["type"] = type, ["duration"] = Num(tr["duration"], 0.1, 10, 0.7) };
            if (OneOf(tr["dir"], Directions, null) is { } dir)
            {
                transition["dir"] = dir;
            }

            if (tr["auto"] is JsonValue)
            {
                transition["auto"] = Num(tr["auto"], 0, 3600, 5);
            }

            slide["transition"] = transition;
        }
        else if (src["transition"] is JsonObject auto && auto["auto"] is JsonValue)
        {
            slide["transition"] = new JsonObject { ["type"] = "none", ["duration"] = 0.7, ["auto"] = Num(auto["auto"], 0, 3600, 5) };
        }

        var notes = TextValue(src["notes"], MaxNotes);
        if (notes.Length > 0)
        {
            slide["notes"] = notes;
        }

        var elements = new JsonArray();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        if (src["elements"] is JsonArray list)
        {
            if (list.Count > MaxElementsPerSlide)
            {
                throw Invalid("Slides.Invalid.TooLarge");
            }

            var n = 0;
            foreach (var node in list)
            {
                n++;
                if (node is not JsonObject e)
                {
                    throw Invalid();
                }

                if (++_elements > MaxElements)
                {
                    throw Invalid("Slides.Invalid.TooLarge");
                }

                elements.Add(Element(e, Unique(Id(e["id"]) ?? "e" + n, ids)));
            }
        }
        else if (src["elements"] is not null)
        {
            throw Invalid();
        }

        DropLonelyGroups(elements);
        slide["elements"] = elements;

        var animations = new JsonArray();
        if (src["animations"] is JsonArray anims)
        {
            foreach (var a in anims.OfType<JsonObject>().Take(MaxAnimations))
            {
                var target = Id(a["el"]);
                if (target is null || !ids.Contains(target))
                {
                    continue;
                }

                var anim = new JsonObject
                {
                    ["el"] = target,
                    ["effect"] = OneOf(a["effect"], Effects, "fade"),
                    ["trigger"] = OneOf(a["trigger"], ["click", "with", "after"], "click"),
                    ["duration"] = Num(a["duration"], 0.05, 20, 0.5),
                };
                if (OneOf(a["dir"], AnimDirections, null) is { } dir)
                {
                    anim["dir"] = dir;
                }

                if (Num(a["delay"], 0, 60, 0) is var delay and > 0)
                {
                    anim["delay"] = delay;
                }

                animations.Add(anim);
            }
        }

        if (animations.Count > 0)
        {
            slide["animations"] = animations;
        }

        return slide;
    }

    private static void DropLonelyGroups(JsonArray elements)
    {
        var counts = elements.Select(e => e!["group"]?.GetValue<string>()).Where(g => g is not null).GroupBy(g => g!).ToDictionary(g => g.Key, g => g.Count());
        foreach (var e in elements.OfType<JsonObject>())
        {
            if (e["group"]?.GetValue<string>() is { } g && counts[g] < 2)
            {
                e.Remove("group");
            }
        }
    }

    // ---- elements ----------------------------------------------------------------------------------------------------------------

    private JsonObject Element(JsonObject src, string id)
    {
        var type = Str(src["type"], 10);
        if (type is null || !Types.Contains(type))
        {
            throw Invalid();
        }

        var e = new JsonObject
        {
            ["id"] = id,
            ["type"] = type,
            ["x"] = Num(src["x"], -20000, 20000, 0),
            ["y"] = Num(src["y"], -20000, 20000, 0),
            ["w"] = Num(src["w"], 0, 20000, 100),
            ["h"] = Num(src["h"], 0, 20000, 100),
        };
        if (Num(src["rot"], -3600, 3600, 0) is var rot && Rotation(rot) is var r and not 0)
        {
            e["rot"] = r;
        }

        CopyFlag(src, e, "flipH");
        CopyFlag(src, e, "flipV");
        if (Str(src["ph"], 20) is { } ph && PhRegex().IsMatch(ph))
        {
            e["ph"] = ph;
        }

        if (Id(src["group"]) is { } group)
        {
            e["group"] = group;
        }

        if (Str(src["name"], 100) is { Length: > 0 } name)
        {
            e["name"] = name;
        }

        switch (type)
        {
            case "text":
            case "shape":
                if (type == "shape")
                {
                    e["shape"] = OneOf(src["shape"], Shapes, "rect");
                    if (src["adj"] is JsonValue)
                    {
                        e["adj"] = Num(src["adj"], -10, 10, 0, 4);
                    }
                }

                ShapeLook(src, e);
                if (src["text"] is JsonObject text)
                {
                    e["text"] = TextBody(text);
                }
                else if (type == "text")
                {
                    e["text"] = TextBody(new JsonObject());
                }

                break;
            case "line":
                e["line"] = Line(src["line"]) ?? new JsonObject { ["color"] = "accent1", ["width"] = 1.5, ["dash"] = "solid" };
                e["arrowStart"] = OneOf(src["arrowStart"], Arrows, "none");
                e["arrowEnd"] = OneOf(src["arrowEnd"], Arrows, "none");
                CopyFlag(src, e, "shadow");
                break;
            case "image":
                e["src"] = Picture(src["src"]);
                if (src["crop"] is JsonObject crop)
                {
                    var l = Num(crop["l"], -1, 0.95, 0, 4);
                    var t = Num(crop["t"], -1, 0.95, 0, 4);
                    var rr = Num(crop["r"], -1, 0.95, 0, 4);
                    var b = Num(crop["b"], -1, 0.95, 0, 4);
                    if (l + rr < 0.98 && t + b < 0.98 && (l != 0 || t != 0 || rr != 0 || b != 0))
                    {
                        e["crop"] = new JsonObject { ["l"] = l, ["t"] = t, ["r"] = rr, ["b"] = b };
                    }
                }

                if (Line(src["line"]) is { } border)
                {
                    e["line"] = border;
                }

                CopyFlag(src, e, "shadow");
                if (Num(src["opacity"], 0, 1, 1) is var op and < 1)
                {
                    e["opacity"] = op;
                }

                if (Str(src["alt"], 2000) is { Length: > 0 } alt)
                {
                    e["alt"] = alt;
                }

                break;
            case "table":
                Table(src, e);
                break;
        }

        return e;
    }

    private void ShapeLook(JsonObject src, JsonObject e)
    {
        if (Fill(src["fill"]) is { } fill)
        {
            e["fill"] = fill;
        }

        if (Line(src["line"]) is { } line)
        {
            e["line"] = line;
        }

        CopyFlag(src, e, "shadow");
        if (Num(src["opacity"], 0, 1, 1) is var opacity and < 1)
        {
            e["opacity"] = opacity;
        }
    }

    private void Table(JsonObject src, JsonObject e)
    {
        if (src["cols"] is not JsonArray cols || cols.Count == 0 || cols.Count > MaxCols || src["rows"] is not JsonArray rows || rows.Count == 0 || rows.Count > MaxRows)
        {
            throw Invalid();
        }

        var widths = new JsonArray();
        foreach (var c in cols)
        {
            widths.Add(Num(c, 4, 20000, 100));
        }

        var outRows = new JsonArray();
        foreach (var rowNode in rows)
        {
            if (rowNode is not JsonObject row)
            {
                throw Invalid();
            }

            var cells = new JsonArray();
            var srcCells = row["cells"] as JsonArray ?? [];
            for (var i = 0; i < cols.Count; i++)
            {
                var cell = i < srcCells.Count ? srcCells[i] as JsonObject : null;
                var c = new JsonObject { ["text"] = TextBody(cell?["text"] as JsonObject ?? new JsonObject()) };
                if (cell is not null && Fill(cell["fill"]) is { } fill)
                {
                    c["fill"] = fill;
                }

                cells.Add(c);
            }

            outRows.Add(new JsonObject { ["h"] = Num(row["h"], 4, 20000, 40), ["cells"] = cells });
        }

        e["cols"] = widths;
        e["rows"] = outRows;
        e["tableStyle"] = OneOf(src["tableStyle"], ["none", "light", "medium", "dark"], "medium");
        e["accent"] = Color(src["accent"]) ?? "accent1";
        e["header"] = src["header"] is null || Bool(src["header"]);
        e["banded"] = src["banded"] is null || Bool(src["banded"]);
        e["firstCol"] = Bool(src["firstCol"]);
        e["lastRow"] = Bool(src["lastRow"]);
    }

    // ---- text ----------------------------------------------------------------------------------------------------------------------

    private JsonObject TextBody(JsonObject src)
    {
        var body = new JsonObject();
        var paras = new JsonArray();
        if (src["paras"] is JsonArray list)
        {
            if (list.Count > MaxParagraphs)
            {
                throw Invalid("Slides.Invalid.TooLarge");
            }

            foreach (var p in list)
            {
                if (p is not JsonObject para)
                {
                    throw Invalid();
                }

                paras.Add(Paragraph(para));
            }
        }

        if (paras.Count == 0)
        {
            paras.Add(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["t"] = string.Empty }) });
        }

        body["paras"] = paras;
        if (OneOf(src["anchor"], ["top", "middle", "bottom"], null) is { } anchor)
        {
            body["anchor"] = anchor;
        }

        if (OneOf(src["autofit"], ["none", "shrink", "grow"], null) is { } autofit)
        {
            body["autofit"] = autofit;
        }

        if (src["inset"] is JsonArray inset && inset.Count == 4)
        {
            body["inset"] = new JsonArray(inset.Select(v => (JsonNode?)Num(v, 0, 1000, 0)).ToArray());
        }

        if (src["wrap"] is JsonValue && !Bool(src["wrap"]))
        {
            body["wrap"] = false;
        }

        return body;
    }

    private JsonObject Paragraph(JsonObject src)
    {
        var runs = new JsonArray();
        if (src["runs"] is JsonArray list)
        {
            if (list.Count > MaxRuns)
            {
                throw Invalid("Slides.Invalid.TooLarge");
            }

            foreach (var r in list)
            {
                if (r is not JsonObject run)
                {
                    throw Invalid();
                }

                runs.Add(Run(run));
            }
        }

        if (runs.Count == 0)
        {
            runs.Add(new JsonObject { ["t"] = string.Empty });
        }

        var p = new JsonObject { ["runs"] = runs };
        if (OneOf(src["align"], ["left", "center", "right", "justify"], null) is { } align)
        {
            p["align"] = align;
        }

        if (Int(src["level"], 0, 8, 0) is var level and > 0)
        {
            p["level"] = level;
        }

        if (Str(src["bullet"], 4) is { Length: > 0 } bullet && !bullet.Any(char.IsControl))
        {
            p["bullet"] = bullet;
        }

        if (src["lineSpacing"] is JsonValue)
        {
            p["lineSpacing"] = Num(src["lineSpacing"], 0.5, 5, 1);
        }

        if (src["spaceBefore"] is JsonValue)
        {
            p["spaceBefore"] = Num(src["spaceBefore"], 0, 400, 0);
        }

        if (src["spaceAfter"] is JsonValue)
        {
            p["spaceAfter"] = Num(src["spaceAfter"], 0, 400, 0);
        }

        return p;
    }

    private JsonObject Run(JsonObject src)
    {
        var text = TextValue(src["t"], MaxRunText);
        _characters += text.Length;
        if (_characters > MaxCharacters)
        {
            throw Invalid("Slides.Invalid.TooLarge");
        }

        var r = new JsonObject { ["t"] = text };
        if (src["b"] is JsonValue bv && bv.TryGetValue<bool>(out var bold))
        {
            r["b"] = bold;
        }

        CopyFlag(src, r, "i");
        CopyFlag(src, r, "u");
        CopyFlag(src, r, "s");
        if (Font(src["font"]) is { } font)
        {
            r["font"] = font;
        }

        if (src["size"] is JsonValue)
        {
            r["size"] = Num(src["size"], 1, 400, 18);
        }

        if (Color(src["color"]) is { } color)
        {
            r["color"] = color;
        }

        if (Color(src["hl"]) is { } hl)
        {
            r["hl"] = hl;
        }

        return r;
    }

    // ---- looks --------------------------------------------------------------------------------------------------------------------

    private JsonObject? Fill(JsonNode? node)
    {
        if (node is not JsonObject src)
        {
            return null;
        }

        switch (Str(src["type"], 10))
        {
            case "none":
                return new JsonObject { ["type"] = "none" };
            case "solid":
                return Color(src["color"]) is { } c ? new JsonObject { ["type"] = "solid", ["color"] = c } : null;
            case "gradient":
                var from = Color(src["from"]);
                var to = Color(src["to"]);
                return from is null || to is null ? null : new JsonObject { ["type"] = "gradient", ["from"] = from, ["to"] = to, ["angle"] = Rotation(Num(src["angle"], -360, 360, 90)) };
            case "image":
                return new JsonObject { ["type"] = "image", ["src"] = Picture(src["src"]) };
            default:
                return null;
        }
    }

    private static JsonObject? Line(JsonNode? node)
    {
        if (node is not JsonObject src || Color(src["color"]) is not { } color)
        {
            return null;
        }

        return new JsonObject { ["color"] = color, ["width"] = Num(src["width"], 0.25, 100, 1), ["dash"] = OneOf(src["dash"], Dashes, "solid") };
    }

    /// <summary>A picture: only a data: address of PNG, JPEG, GIF or WebP, base64 – everything else makes the content invalid.</summary>
    private string Picture(JsonNode? node)
    {
        var src = node is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
        if (src is null || src.Length > Kinds.MaxContentBytes || !PictureRegex().IsMatch(src))
        {
            throw Invalid("Slides.Invalid.Picture");
        }

        var comma = src.IndexOf(',', StringComparison.Ordinal);
        for (var i = comma + 1; i < src.Length; i++)
        {
            var c = src[i];
            if (!(char.IsAsciiLetterOrDigit(c) || c is '+' or '/' or '='))
            {
                throw Invalid("Slides.Invalid.Picture");
            }
        }

        _characters += src.Length / 16;
        return src;
    }

    /// <summary>A colour: #rrggbb, or a colour of the design (dk1 … accent6, hlink) with an optional lighter (+n) / darker (-n) percentage.</summary>
    internal static string? Color(JsonNode? node)
    {
        var s = node is JsonValue v && v.TryGetValue<string>(out var str) ? str.Trim() : null;
        if (s is null)
        {
            return null;
        }

        if (HexRegex().IsMatch(s))
        {
            return s.ToLowerInvariant();
        }

        return SchemeRegex().IsMatch(s) ? s : null;
    }

    private static string? Hex(JsonNode? node)
    {
        var s = node is JsonValue v && v.TryGetValue<string>(out var str) ? str.Trim() : null;
        return s is not null && HexRegex().IsMatch(s) ? s.ToLowerInvariant() : null;
    }

    private static string? Font(JsonNode? node)
    {
        var s = Str(node, 64)?.Trim();
        return s is { Length: > 0 } && FontRegex().IsMatch(s) ? s : null;
    }

    // ---- primitives ----------------------------------------------------------------------------------------------------------------

    private static string? Id(JsonNode? node)
    {
        var s = Str(node, 40);
        return s is not null && IdRegex().IsMatch(s) ? s : null;
    }

    private static string Unique(string id, HashSet<string> seen)
    {
        var candidate = id;
        var n = 2;
        while (!seen.Add(candidate))
        {
            candidate = (id.Length > 34 ? id[..34] : id) + "_" + n++;
        }

        return candidate;
    }

    private static string? Str(JsonNode? node, int max)
    {
        if (node is not JsonValue v || !v.TryGetValue<string>(out var s))
        {
            return null;
        }

        s = new string(s.Where(c => !char.IsControl(c)).ToArray());
        return s.Length > max ? s[..max] : s;
    }

    /// <summary>A text of a run or the notes: control characters out, except the line break and the tab; \r\n becomes \n.</summary>
    private static string TextValue(JsonNode? node, int max)
    {
        if (node is not JsonValue v || !v.TryGetValue<string>(out var s))
        {
            return string.Empty;
        }

        var sb = new StringBuilder(Math.Min(s.Length, max));
        for (var i = 0; i < s.Length && sb.Length < max; i++)
        {
            var c = s[i];
            if (c == '\r')
            {
                if (i + 1 < s.Length && s[i + 1] == '\n')
                {
                    continue;
                }

                c = '\n';
            }

            if (c == '\v')
            {
                c = '\n';
            }

            if (char.IsControl(c) && c is not '\n' and not '\t')
            {
                continue;
            }

            if (char.IsSurrogate(c) && !(char.IsHighSurrogate(c) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1])) && !(char.IsLowSurrogate(c) && i > 0 && char.IsHighSurrogate(s[i - 1])))
            {
                continue;
            }

            sb.Append(c);
        }

        return sb.ToString();
    }

    private static string? OneOf(JsonNode? node, string[] allowed, string? fallback)
    {
        var s = Str(node, 40);
        return s is not null && allowed.Contains(s) ? s : fallback;
    }

    private static bool Bool(JsonNode? node) => node is JsonValue v && v.TryGetValue<bool>(out var b) && b;

    private static void CopyFlag(JsonObject src, JsonObject dst, string name)
    {
        if (Bool(src[name]))
        {
            dst[name] = true;
        }
    }

    internal static double Num(JsonNode? node, double min, double max, double fallback, int digits = 2)
    {
        double value;
        if (node is JsonValue v && v.TryGetValue<double>(out var d))
        {
            value = d;
        }
        else if (node is JsonValue sv && sv.TryGetValue<string>(out var s) && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
        {
            value = parsed;
        }
        else
        {
            return fallback;
        }

        if (!double.IsFinite(value))
        {
            return fallback;
        }

        return Math.Round(Math.Clamp(value, min, max), digits);
    }

    private static int Int(JsonNode? node, int min, int max, int fallback) => (int)Math.Round(Num(node, min, max, fallback, 0));

    private static double Rotation(double degrees)
    {
        var r = degrees % 360;
        if (r < 0)
        {
            r += 360;
        }

        return Math.Round(r, 2) % 360;
    }

    [GeneratedRegex("^#[0-9a-fA-F]{6}$")]
    private static partial Regex HexRegex();

    [GeneratedRegex("^(dk1|lt1|dk2|lt2|accent[1-6]|hlink)(:[+-][0-9]{1,2})?$")]
    private static partial Regex SchemeRegex();

    [GeneratedRegex(@"^[\p{L}\p{N} _.\-]{1,64}$")]
    private static partial Regex FontRegex();

    [GeneratedRegex("^[A-Za-z0-9_-]{1,40}$")]
    private static partial Regex IdRegex();

    [GeneratedRegex("^[A-Za-z][A-Za-z0-9]{0,19}$")]
    private static partial Regex PhRegex();

    [GeneratedRegex("^data:image/(png|jpeg|gif|webp);base64,")]
    private static partial Regex PictureRegex();
}
