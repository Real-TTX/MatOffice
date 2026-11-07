using System.Globalization;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MatOffice.Localization;
using static MatOffice.Services.Formats.Slides.PptxXml;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// Reads a PowerPoint file (.pptx) into the content of Slides: the size, the design of the first master (colours, fonts, background, its shapes as decoration),
/// and every slide with its layout, background, placeholders and text boxes (paragraphs and runs with what the masters say about sizes, colours, bullets),
/// shapes, lines, pictures, tables, groups (flattened, kept as a group), notes, hidden slides, transitions and entrance animations.
/// What it does not know (charts, SmartArt, media, OLE) is left out; a broken element is skipped – the import never fails because of one.
/// </summary>
public sealed class PptxReader
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly HashSet<string> KnownShapes = [.. SlidesNormalizer.Shapes];
    private static readonly string[] FooterTypes = ["dt", "ftr", "sldNum", "hdr", "sldImg"];

    private readonly PresentationPart _presentation;
    private readonly JsonObject _theme = new();
    private readonly Dictionary<string, string> _clrMap = new(StringComparer.Ordinal);
    private readonly JsonObject _footer = new() { ["number"] = false, ["date"] = false, ["dateText"] = string.Empty, ["text"] = string.Empty, ["skipTitle"] = true };
    private double _w = 1280;
    private double _h = 720;
    private long _pictureBytes;
    private XElement? _defaultTextStyle;

    private PptxReader(PresentationPart presentation) => _presentation = presentation;

    /// <summary>The content of a file (not normalized yet). Throws <see cref="Formats.FormatException"/> when it is no presentation.</summary>
    public static JsonObject Read(Stream input)
    {
        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        buffer.Position = 0;
        try
        {
            using var doc = PresentationDocument.Open(buffer, false);
            var presentation = doc.PresentationPart ?? throw new Formats.FormatException(Text.T("Slides.Import.NotPresentation"));
            return new PptxReader(presentation).Content();
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or XmlException or InvalidDataException or IOException or InvalidOperationException or ArgumentException or NotSupportedException or System.FormatException or KeyNotFoundException or NullReferenceException)
        {
            throw new Formats.FormatException(Text.T("Slides.Import.NotPresentation"));
        }
    }

    private JsonObject Content()
    {
        var pres = Load(_presentation);
        var size = pres.Element(P + "sldSz");
        if (size is not null)
        {
            _w = Math.Clamp(Px(LongAttr(size, "cx", 12192000)), 120, 6000);
            _h = Math.Clamp(Px(LongAttr(size, "cy", 6858000)), 120, 6000);
        }

        _defaultTextStyle = pres.Element(P + "defaultTextStyle");
        var slideParts = new List<SlidePart>();
        foreach (var id in pres.Element(P + "sldIdLst")?.Elements(P + "sldId") ?? [])
        {
            var rel = (string?)id.Attribute(R + "id");
            if (rel is not null && _presentation.TryGetPartById(rel, out var part) && part is SlidePart sp)
            {
                slideParts.Add(sp);
            }
        }

        var master = slideParts.FirstOrDefault()?.SlideLayoutPart?.SlideMasterPart ?? _presentation.SlideMasterParts.FirstOrDefault();
        ReadTheme(master, slideParts);

        var slides = new JsonArray();
        var n = 0;
        foreach (var part in slideParts.Take(SlidesNormalizer.MaxSlides))
        {
            n++;
            try
            {
                slides.Add(new SlideReader(this, part, n).Slide());
            }
            catch (Exception ex) when (ex is XmlException or InvalidOperationException or ArgumentException or System.FormatException or NullReferenceException or OverflowException or KeyNotFoundException)
            {
                slides.Add(new JsonObject { ["id"] = "s" + n, ["layout"] = "blank", ["elements"] = new JsonArray() });
            }
        }

        return new JsonObject
        {
            ["version"] = 1,
            ["size"] = new JsonObject { ["w"] = Math.Round(_w, 2), ["h"] = Math.Round(_h, 2) },
            ["theme"] = _theme,
            ["footer"] = _footer,
            ["slides"] = slides,
        };
    }

    private static XElement Load(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream, LoadOptions.None).Root ?? throw new InvalidOperationException();
    }

    private static long LongAttr(XElement? e, string name, long fallback)
        => e?.Attribute(name) is { } a && long.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : fallback;

    private static int IntAttr(XElement? e, string name, int fallback)
        => e?.Attribute(name) is { } a && int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : fallback;

    private static bool BoolAttr(XElement? e, string name) => (string?)e?.Attribute(name) is "1" or "true";

    // ---- design ----------------------------------------------------------------------------------------------------------------------

    private void ReadTheme(SlideMasterPart? master, IReadOnlyList<SlidePart> slideParts)
    {
        var fallback = SlidesData.Theme("paper");
        var colors = new JsonObject();
        var name = "Imported";
        var major = "Calibri Light";
        var minor = "Calibri";
        var themeXml = master?.ThemePart is { } tp ? Load(tp) : null;
        if (themeXml is not null)
        {
            name = (string?)themeXml.Attribute("name") is { Length: > 0 } nm ? nm : name;
            var scheme = themeXml.Element(A + "themeElements")?.Element(A + "clrScheme");
            foreach (var key in SlidesData.SchemeNames)
            {
                var slot = scheme?.Element(A + key);
                var hex = slot?.Element(A + "srgbClr")?.Attribute("val")?.Value ?? slot?.Element(A + "sysClr")?.Attribute("lastClr")?.Value;
                colors[key] = hex is { Length: 6 } ? "#" + hex.ToLowerInvariant() : fallback["colors"]![key]!.GetValue<string>();
            }

            var fonts = themeXml.Element(A + "themeElements")?.Element(A + "fontScheme");
            major = Typeface(fonts?.Element(A + "majorFont")) ?? major;
            minor = Typeface(fonts?.Element(A + "minorFont")) ?? minor;
        }
        else
        {
            foreach (var key in SlidesData.SchemeNames)
            {
                colors[key] = fallback["colors"]![key]!.GetValue<string>();
            }
        }

        _theme["id"] = "imported";
        _theme["name"] = name.Length > 100 ? name[..100] : name;
        _theme["fonts"] = new JsonObject { ["major"] = major, ["minor"] = minor };
        _theme["colors"] = colors;

        var masterXml = master is null ? null : Load(master);
        if (masterXml?.Element(P + "clrMap") is { } map)
        {
            foreach (var attr in map.Attributes())
            {
                _clrMap[attr.Name.LocalName] = attr.Value;
            }
        }

        MasterXml = masterXml;
        var dark = _clrMap.GetValueOrDefault("tx1") == "lt1";
        _theme["background"] = Background(masterXml, master) ?? new JsonObject { ["type"] = "solid", ["color"] = dark ? "dk1" : "lt1" };
        var titleColor = Color(masterXml?.Element(P + "txStyles")?.Element(P + "titleStyle")?.Element(A + "lvl1pPr")?.Element(A + "defRPr")).Color;
        var bodyColor = Color(masterXml?.Element(P + "txStyles")?.Element(P + "bodyStyle")?.Element(A + "lvl1pPr")?.Element(A + "defRPr")).Color;
        _theme["heading"] = titleColor ?? (dark ? "lt1" : "dk1");
        _theme["text"] = bodyColor ?? (dark ? "lt1" : "dk1");
        _theme["decor"] = masterXml is not null && master is not null ? Decorations(master, masterXml, slideParts) : new JsonArray();
    }

    internal XElement? MasterXml { get; private set; }

    private static string? Typeface(XElement? font)
    {
        var face = (string?)font?.Element(A + "latin")?.Attribute("typeface");
        return string.IsNullOrWhiteSpace(face) ? null : face;
    }

    /// <summary>
    /// The decoration of the design: the shapes of the master (on every layout that shows the master's shapes) and of the layouts (on the layout of Slides
    /// that the file's layout becomes; the first file layout per Slides layout counts, the ones the slides use first). The same shape on several layouts is one entry.
    /// </summary>
    private JsonArray Decorations(SlideMasterPart master, XElement masterXml, IReadOnlyList<SlidePart> slideParts)
    {
        var layouts = new Dictionary<string, (SlideLayoutPart Part, XElement Xml)>(StringComparer.Ordinal);
        foreach (var lp in slideParts.Select(s => s.SlideLayoutPart).OfType<SlideLayoutPart>().Concat(master.SlideLayoutParts).Distinct())
        {
            XElement xml;
            try
            {
                xml = Load(lp);
            }
            catch (Exception ex) when (ex is XmlException or InvalidOperationException)
            {
                continue;
            }

            layouts.TryAdd(SlideReader.LayoutOf(xml), (lp, xml));
        }

        var hiding = layouts.Where(l => (string?)l.Value.Xml.Attribute("showMasterSp") is "0" or "false").Select(l => l.Key).ToHashSet(StringComparer.Ordinal);
        var decor = new JsonArray();
        foreach (var d in Decoration(masterXml, master, "all"))
        {
            if (hiding.Count > 0)
            {
                d["on"] = new JsonArray(SlidesData.Layouts.Select(l => l.Id).Where(id => !hiding.Contains(id)).Select(id => (JsonNode)JsonValue.Create(id)!).ToArray());
            }

            decor.Add(d);
        }

        foreach (var (id, (part, xml)) in layouts)
        {
            foreach (var d in Decoration(xml, part, id))
            {
                var same = decor.OfType<JsonObject>().FirstOrDefault(x => SameDecoration(x, d));
                if (same is null)
                {
                    decor.Add(d);
                }
                else if (same["on"] is JsonArray on && !on.Any(o => o?.GetValue<string>() is { } v && (v == "all" || v == id)))
                {
                    on.Add(id);
                }
            }
        }

        return decor;
    }

    private static bool SameDecoration(JsonObject a, JsonObject b)
    {
        var x = (JsonObject)a.DeepClone();
        var y = (JsonObject)b.DeepClone();
        x.Remove("on");
        y.Remove("on");
        return JsonNode.DeepEquals(x, y);
    }

    /// <summary>The shapes of a master or layout that are no placeholders: the decoration of the design.</summary>
    private IEnumerable<JsonObject> Decoration(XElement root, OpenXmlPart part, string on)
    {
        var tree = root.Element(P + "cSld")?.Element(P + "spTree");
        if (tree is null)
        {
            yield break;
        }

        var count = 0;
        foreach (var e in tree.Elements())
        {
            if (count >= 20)
            {
                yield break;
            }

            if (e.Descendants(P + "ph").Any())
            {
                continue;
            }

            var xfrm = e.Element(P + "spPr")?.Element(A + "xfrm");
            if (xfrm is null)
            {
                continue;
            }

            var (x, y, w, h) = Box(xfrm);
            var d = new JsonObject
            {
                ["on"] = new JsonArray(on),
                ["x"] = Math.Round(x / _w, 4), ["y"] = Math.Round(y / _h, 4), ["w"] = Math.Round(w / _w, 4), ["h"] = Math.Round(h / _h, 4),
            };
            if (BoolAttr(xfrm, "flipH"))
            {
                d["flipH"] = true;
            }

            if (BoolAttr(xfrm, "flipV"))
            {
                d["flipV"] = true;
            }

            if (e.Name == P + "pic")
            {
                var rel = (string?)e.Element(P + "blipFill")?.Element(A + "blip")?.Attribute(R + "embed");
                if (Picture(part, rel) is not { } src)
                {
                    continue;
                }

                d["shape"] = "rect";
                d["src"] = src;
            }
            else if (e.Name == P + "sp")
            {
                var spPr = e.Element(P + "spPr")!;
                var fill = Fill(spPr, part, StyleColor(e, "fillRef"));
                if (fill?["type"]?.GetValue<string>() != "solid")
                {
                    continue;
                }

                d["shape"] = ShapeName(spPr);
                d["fill"] = fill["color"]!.DeepClone();
                if (fill["opacity"] is { } op)
                {
                    d["opacity"] = op.DeepClone();
                }
            }
            else
            {
                continue;
            }

            count++;
            yield return d;
        }
    }

    // ---- colours and fills -----------------------------------------------------------------------------------------------------------

    /// <summary>The colour inside an element (srgbClr, schemeClr, sysClr, prstClr …) as a colour of the content, and its alpha.</summary>
    internal (string? Color, double Alpha) Color(XElement? holder)
    {
        if (holder is null)
        {
            return (null, 1);
        }

        foreach (var c in holder.Elements())
        {
            var alpha = c.Element(A + "alpha") is { } al ? Math.Clamp(IntAttr(al, "val", 100000) / 100000.0, 0, 1) : 1;
            switch (c.Name.LocalName)
            {
                case "srgbClr":
                    return (Hex((string?)c.Attribute("val")), alpha);
                case "sysClr":
                    return (Hex((string?)c.Attribute("lastClr")) ?? ((string?)c.Attribute("val") == "window" ? "#ffffff" : "#000000"), alpha);
                case "prstClr":
                    return ((string?)c.Attribute("val") switch { "white" => "#ffffff", "black" => "#000000", "red" => "#ff0000", "blue" => "#0000ff", "green" => "#008000", "yellow" => "#ffff00", "gray" or "grey" => "#808080", _ => "#000000" }, alpha);
                case "scrgbClr":
                    return ($"#{Channel(c, "r")}{Channel(c, "g")}{Channel(c, "b")}", alpha);
                case "schemeClr":
                    var name = SchemeIn((string?)c.Attribute("val") ?? "tx1");
                    if (name is null)
                    {
                        return (null, alpha);
                    }

                    return (WithMods(name, c), alpha);
            }
        }

        return (null, 1);
    }

    private static string Channel(XElement c, string attr) => ((int)Math.Round(Math.Clamp(IntAttr(c, attr, 0) / 100000.0, 0, 1) * 255)).ToString("x2", CultureInfo.InvariantCulture);

    private static string? Hex(string? val) => val is { Length: 6 } && val.All(Uri.IsHexDigit) ? "#" + val.ToLowerInvariant() : null;

    /// <summary>tx1/bg1/tx2/bg2 through the colour map of the master to dk1/lt1/dk2/lt2.</summary>
    private string? SchemeIn(string val)
    {
        if (val is "tx1" or "bg1" or "tx2" or "bg2")
        {
            val = _clrMap.GetValueOrDefault(val) ?? val switch { "tx1" => "dk1", "bg1" => "lt1", "tx2" => "dk2", _ => "lt2" };
        }

        if (val == "folHlink")
        {
            val = "hlink";
        }

        return SlidesData.SchemeNames.Contains(val) ? val : null;
    }

    /// <summary>A scheme colour with its modifications as ":+n" / ":-n" (lumMod/lumOff, tint, shade).</summary>
    private static string WithMods(string name, XElement c)
    {
        var lumMod = c.Element(A + "lumMod") is { } lm ? IntAttr(lm, "val", 100000) : 100000;
        var lumOff = c.Element(A + "lumOff") is { } lo ? IntAttr(lo, "val", 0) : 0;
        var mod = 0;
        if (lumOff > 0)
        {
            mod = (int)Math.Round(lumOff / 1000.0);
        }
        else if (lumMod < 100000)
        {
            mod = -(int)Math.Round((100000 - lumMod) / 1000.0);
        }
        else if (c.Element(A + "tint") is { } tint)
        {
            mod = (int)Math.Round((100000 - IntAttr(tint, "val", 100000)) / 1000.0);
        }
        else if (c.Element(A + "shade") is { } shade)
        {
            mod = -(int)Math.Round((100000 - IntAttr(shade, "val", 100000)) / 1000.0);
        }

        mod = Math.Clamp(mod, -99, 99);
        return mod == 0 ? name : mod > 0 ? $"{name}:+{mod}" : $"{name}:{mod}";
    }

    /// <summary>The colour of a style reference of a shape (p:style/a:fillRef …), used when the shape says nothing itself.</summary>
    internal (string? Color, int Idx) StyleColor(XElement shape, string reference)
    {
        var r = shape.Element(P + "style")?.Element(A + reference);
        return r is null ? (null, 0) : (Color(r).Color, IntAttr(r, "idx", 0));
    }

    /// <summary>The fill of shape properties (or a background): none, solid, gradient (first and last stop), a picture.</summary>
    internal JsonObject? Fill(XElement? props, OpenXmlPart part, (string? Color, int Idx) style = default)
    {
        if (props is null)
        {
            return null;
        }

        if (props.Element(A + "noFill") is not null)
        {
            return new JsonObject { ["type"] = "none" };
        }

        if (props.Element(A + "solidFill") is { } solid)
        {
            var (color, alpha) = Color(solid);
            var f = new JsonObject { ["type"] = "solid", ["color"] = color ?? "accent1" };
            if (alpha < 1)
            {
                f["opacity"] = Math.Round(alpha, 2);
            }

            return f;
        }

        if (props.Element(A + "gradFill") is { } grad)
        {
            var stops = grad.Element(A + "gsLst")?.Elements(A + "gs").OrderBy(g => IntAttr(g, "pos", 0)).ToList() ?? [];
            if (stops.Count == 0)
            {
                return null;
            }

            var from = Color(stops[0]).Color ?? "lt1";
            var to = Color(stops[^1]).Color ?? from;
            var ang = grad.Element(A + "lin") is { } lin ? IntAttr(lin, "ang", 5400000) / 60000.0 : 90;
            return new JsonObject { ["type"] = "gradient", ["from"] = from, ["to"] = to, ["angle"] = Math.Round((ang + 90) % 360, 2) };
        }

        if (props.Element(A + "blipFill") is { } blip)
        {
            var rel = (string?)blip.Element(A + "blip")?.Attribute(R + "embed");
            return Picture(part, rel) is { } src ? new JsonObject { ["type"] = "image", ["src"] = src } : null;
        }

        if (style.Idx > 0 && style.Color is not null && props.Element(A + "grpFill") is null)
        {
            return new JsonObject { ["type"] = "solid", ["color"] = style.Color };
        }

        return null;
    }

    private JsonObject? Background(XElement? root, OpenXmlPart? part)
    {
        var bg = root?.Element(P + "cSld")?.Element(P + "bg");
        if (bg is null || part is null)
        {
            return null;
        }

        if (bg.Element(P + "bgPr") is { } pr)
        {
            var f = Fill(pr, part);
            f?.Remove("opacity");
            return f;
        }

        if (bg.Element(P + "bgRef") is { } reference && Color(reference).Color is { } color)
        {
            return new JsonObject { ["type"] = "solid", ["color"] = color };
        }

        return null;
    }

    /// <summary>A picture of a part as a data: address (PNG, JPEG, GIF, WebP only; at most 40 MB in all).</summary>
    internal string? Picture(OpenXmlPart part, string? rel)
    {
        if (rel is null || !part.TryGetPartById(rel, out var target) || target is not ImagePart image)
        {
            return null;
        }

        var mime = image.ContentType switch
        {
            "image/png" => "image/png",
            "image/jpeg" or "image/jpg" or "image/pjpeg" => "image/jpeg",
            "image/gif" => "image/gif",
            "image/webp" => "image/webp",
            _ => null,
        };
        if (mime is null)
        {
            return null;
        }

        using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        _pictureBytes += ms.Length;
        if (_pictureBytes > Kinds.MaxContentBytes * 3L / 4)
        {
            return null;
        }

        return $"data:{mime};base64,{Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length)}";
    }

    internal static (double X, double Y, double W, double H) Box(XElement xfrm)
    {
        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");
        return (Px(LongAttr(off, "x", 0)), Px(LongAttr(off, "y", 0)), Px(LongAttr(ext, "cx", 0)), Px(LongAttr(ext, "cy", 0)));
    }

    internal static string ShapeName(XElement spPr)
    {
        var prst = (string?)spPr.Element(A + "prstGeom")?.Attribute("prst");
        if (prst is null)
        {
            return "rect";
        }

        if (KnownShapes.Contains(prst))
        {
            return prst;
        }

        return prst switch
        {
            "flowChartAlternateProcess" or "round2SameRect" or "round1Rect" or "round2DiagRect" or "snipRoundRect" => "roundRect",
            "flowChartConnector" or "flowChartOr" or "flowChartSummingJunction" => "ellipse",
            "flowChartDecision" => "diamond",
            "flowChartExtract" or "flowChartMerge" => "triangle",
            "flowChartInputOutput" => "parallelogram",
            "flowChartManualOperation" => "trapezoid",
            "flowChartPreparation" => "hexagon",
            "star4" or "star7" or "star8" or "star10" or "star12" or "star16" or "star24" or "star32" => "star5",
            "notchedRightArrow" or "stripedRightArrow" or "rightArrowCallout" => "rightArrow",
            "leftArrowCallout" => "leftArrow",
            "upArrowCallout" => "upArrow",
            "downArrowCallout" => "downArrow",
            "leftRightArrowCallout" => "leftRightArrow",
            "cloudCallout" => "cloud",
            "borderCallout1" or "borderCallout2" or "callout1" or "callout2" or "accentCallout1" => "wedgeRectCallout",
            "teardrop" or "pie" or "chord" or "blockArc" => "ellipse",
            "snip2SameRect" or "snip2DiagRect" or "snipRect" => "snip1Rect",
            "cube" or "bevel" or "frame" or "plaque" or "foldedCorner" or "flowChartProcess" => "rect",
            _ => "rect",
        };
    }

    // ---- one slide ---------------------------------------------------------------------------------------------------------------------

    /// <summary>The text properties of one level along the chain of a placeholder (shape, layout, master, text styles).</summary>
    private sealed record Chain(List<XElement> Styles);

    private sealed class SlideReader(PptxReader owner, SlidePart part, int number)
    {
        private readonly JsonArray _elements = [];
        private readonly Dictionary<string, string> _ids = new(StringComparer.Ordinal);
        private readonly XElement? _layoutXml = part.SlideLayoutPart is { } lp ? Load(lp) : null;
        private SlideLayout _layout = SlidesData.Layout("blank");
        private int _next;
        private int _groups;

        public JsonObject Slide()
        {
            var root = Load(part);
            _layout = SlidesData.Layout(LayoutOf(_layoutXml));
            var slide = new JsonObject { ["id"] = "s" + number, ["layout"] = _layout.Id };
            if ((string?)root.Attribute("show") is "0" or "false")
            {
                slide["hidden"] = true;
            }

            var bg = owner.Background(root, part) ?? (part.SlideLayoutPart is { } layoutPart ? owner.Background(_layoutXml, layoutPart) : null);
            if (bg is not null)
            {
                slide["background"] = bg;
            }

            var tree = root.Element(P + "cSld")?.Element(P + "spTree");
            if (tree is not null)
            {
                var placeholders = AssignSlots(tree);
                Walk(tree, null, placeholders);
            }

            slide["elements"] = _elements;
            if (Transition(root) is { } transition)
            {
                slide["transition"] = transition;
            }

            if (Animations(root) is { Count: > 0 } anims)
            {
                slide["animations"] = anims;
            }

            if (Notes() is { Length: > 0 } notes)
            {
                slide["notes"] = notes;
            }

            return slide;
        }

        internal static string LayoutOf(XElement? layout)
        {
            var type = (string?)layout?.Attribute("type");
            var mapped = type switch
            {
                "title" => "title",
                "obj" or "tx" or "objOnly" or "txOverObj" or "objOverTx" or "chart" or "tbl" or "dgm" or "txAndObj" or "objAndTx" or "vertTx" or "clipArtAndTx" => "titleContent",
                "secHead" => "section",
                "twoObj" or "twoColTx" => "twoContent",
                "twoTxTwoObj" => "comparison",
                "titleOnly" => "titleOnly",
                "blank" => "blank",
                "picTx" or "objTx" => "pictureCaption",
                _ => null,
            };
            if (mapped is not null)
            {
                return mapped;
            }

            // a custom layout: guess by its placeholders
            var phs = layout?.Descendants(P + "ph").Select(p => (string?)p.Attribute("type") ?? "obj").Where(t => !FooterTypes.Contains(t)).ToList() ?? [];
            if (phs.Contains("ctrTitle"))
            {
                return "title";
            }

            var bodies = phs.Count(t => t is "body" or "obj" or "subTitle" or "pic" or "tbl" or "chart");
            var hasTitle = phs.Contains("title");
            return (hasTitle, bodies) switch
            {
                (true, 0) => "titleOnly",
                (true, 1) => "titleContent",
                (true, 2) => "twoContent",
                (true, >= 3) => "comparison",
                (false, 0) => "blank",
                _ => "titleContent",
            };
        }

        /// <summary>Gives the placeholders of the slide the slots of the chosen layout: titles to "title", the others by their index in order.</summary>
        private Dictionary<XElement, LayoutSlot> AssignSlots(XElement tree)
        {
            var result = new Dictionary<XElement, LayoutSlot>();
            var phs = tree.Elements().Where(e => e.Name == P + "sp" || e.Name == P + "pic")
                .Select(e => (Element: e, Ph: e.Descendants(P + "ph").FirstOrDefault()))
                .Where(x => x.Ph is not null && !FooterTypes.Contains((string?)x.Ph.Attribute("type") ?? "obj"))
                .ToList();
            var free = _layout.Slots.ToList();
            foreach (var (element, _) in phs.Where(x => (string?)x.Ph!.Attribute("type") is "title" or "ctrTitle"))
            {
                if (free.FirstOrDefault(s => s.Role == "title") is { } slot)
                {
                    result[element] = slot;
                    free.Remove(slot);
                }
            }

            foreach (var (element, _) in phs.Where(x => (string?)x.Ph!.Attribute("type") == "subTitle"))
            {
                if ((free.FirstOrDefault(s => s.Ph == "subtitle") ?? free.FirstOrDefault(s => s.Role == "body")) is { } slot)
                {
                    result[element] = slot;
                    free.Remove(slot);
                }
            }

            foreach (var (element, _) in phs.Where(x => (string?)x.Ph!.Attribute("type") == "pic" || x.Element.Name == P + "pic"))
            {
                if (free.FirstOrDefault(s => s.Role == "picture") is { } slot)
                {
                    result[element] = slot;
                    free.Remove(slot);
                }
            }

            foreach (var (element, _) in phs.Where(x => !result.ContainsKey(x.Element) && (string?)x.Ph!.Attribute("type") is not ("title" or "ctrTitle" or "subTitle")).OrderBy(x => IntAttr(x.Ph, "idx", 0)))
            {
                if (free.FirstOrDefault(s => s.Role == "body") is { } slot)
                {
                    result[element] = slot;
                    free.Remove(slot);
                }
            }

            return result;
        }

        private string NewId() => "e" + ++_next;

        /// <summary>The elements of a tree (or a group) into the list; a group maps the coordinates of its children.</summary>
        private void Walk(XElement tree, Func<(double, double, double, double), (double, double, double, double)>? map, Dictionary<XElement, LayoutSlot> slots, string? group = null, double groupRot = 0)
        {
            foreach (var child in tree.Elements())
            {
                if (_elements.Count >= SlidesNormalizer.MaxElementsPerSlide)
                {
                    return;
                }

                var e = child;
                if (e.Name == Mc + "AlternateContent")
                {
                    e = Alternate(e);
                    if (e is null)
                    {
                        continue;
                    }
                }

                try
                {
                    if (e.Name == P + "grpSp")
                    {
                        Group(e, map, slots, groupRot);
                        continue;
                    }

                    var element = e.Name.LocalName switch
                    {
                        "sp" => Shape(e, slots.GetValueOrDefault(e)),
                        "cxnSp" => Line(e),
                        "pic" => Pic(e, slots.GetValueOrDefault(e)),
                        "graphicFrame" => Table(e),
                        _ => null,
                    };
                    if (element is null)
                    {
                        continue;
                    }

                    if (map is not null)
                    {
                        var (x, y, w, h) = map((element["x"]!.GetValue<double>(), element["y"]!.GetValue<double>(), element["w"]!.GetValue<double>(), element["h"]!.GetValue<double>()));
                        element["x"] = Math.Round(x, 2);
                        element["y"] = Math.Round(y, 2);
                        element["w"] = Math.Round(w, 2);
                        element["h"] = Math.Round(h, 2);
                        if (groupRot != 0)
                        {
                            element["rot"] = Math.Round(((element["rot"]?.GetValue<double>() ?? 0) + groupRot) % 360, 2);
                        }
                    }

                    if (group is not null)
                    {
                        element["group"] = group;
                    }

                    var source = (string?)e.Descendants().FirstOrDefault(d => d.Name.LocalName == "cNvPr")?.Attribute("id");
                    if (source is not null)
                    {
                        _ids[source] = element["id"]!.GetValue<string>();
                    }

                    _elements.Add(element);
                }
                catch (Exception ex) when (ex is System.FormatException or InvalidOperationException or ArgumentException or NullReferenceException or OverflowException or KeyNotFoundException)
                {
                    // one broken element does not stop the import
                }
            }
        }

        private static XElement? Alternate(XElement alternate)
        {
            foreach (var choice in alternate.Elements(Mc + "Choice"))
            {
                var requires = ((string?)choice.Attribute("Requires") ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (requires.All(r => r is "p14" or "a14" or "p15" or "v"))
                {
                    return choice.Elements().FirstOrDefault();
                }
            }

            return alternate.Element(Mc + "Fallback")?.Elements().FirstOrDefault();
        }

        private void Group(XElement g, Func<(double, double, double, double), (double, double, double, double)>? outer, Dictionary<XElement, LayoutSlot> slots, double outerRot)
        {
            var xfrm = g.Element(P + "grpSpPr")?.Element(A + "xfrm");
            if (xfrm is null)
            {
                return;
            }

            var (x, y, w, h) = Box(xfrm);
            var chOff = xfrm.Element(A + "chOff");
            var chExt = xfrm.Element(A + "chExt");
            var cx = Px(LongAttr(chOff, "x", 0));
            var cy = Px(LongAttr(chOff, "y", 0));
            var cw = Px(LongAttr(chExt, "cx", 0));
            var ch = Px(LongAttr(chExt, "cy", 0));
            var sx = cw > 0 ? w / cw : 1;
            var sy = ch > 0 ? h / ch : 1;
            (double, double, double, double) Map((double X, double Y, double W, double H) b)
            {
                var r = (x + ((b.X - cx) * sx), y + ((b.Y - cy) * sy), b.W * sx, b.H * sy);
                return outer is null ? r : outer(r);
            }

            var id = "g" + number + "_" + ++_groups;
            Walk(g, Map, slots, id, outerRot + (IntAttr(xfrm, "rot", 0) / 60000.0));
        }

        private JsonObject Geometry(XElement e, XElement? xfrm, LayoutSlot? slot, XElement? ph)
        {
            double x, y, w, h;
            if (xfrm?.Element(A + "off") is not null)
            {
                (x, y, w, h) = Box(xfrm);
            }
            else if (InheritedXfrm(ph) is { } inherited)
            {
                (x, y, w, h) = Box(inherited);
            }
            else if (slot is not null)
            {
                (x, y, w, h) = (slot.X * owner._w, slot.Y * owner._h, slot.W * owner._w, slot.H * owner._h);
            }
            else
            {
                (x, y, w, h) = (0, 0, 100, 100);
            }

            var g = new JsonObject { ["id"] = NewId(), ["x"] = Math.Round(x, 2), ["y"] = Math.Round(y, 2), ["w"] = Math.Round(w, 2), ["h"] = Math.Round(h, 2) };
            if (xfrm is not null)
            {
                var rot = IntAttr(xfrm, "rot", 0) / 60000.0;
                if (rot != 0)
                {
                    g["rot"] = Math.Round(((rot % 360) + 360) % 360, 2);
                }

                if (BoolAttr(xfrm, "flipH"))
                {
                    g["flipH"] = true;
                }

                if (BoolAttr(xfrm, "flipV"))
                {
                    g["flipV"] = true;
                }
            }

            if ((string?)e.Descendants().FirstOrDefault(d => d.Name.LocalName == "cNvPr")?.Attribute("name") is { Length: > 0 } name)
            {
                g["name"] = name.Length > 100 ? name[..100] : name;
            }

            return g;
        }

        /// <summary>The placeholder of the layout (then of the master) that a placeholder of the slide inherits from.</summary>
        private IEnumerable<XElement> Inherited(XElement? ph)
        {
            if (ph is null)
            {
                yield break;
            }

            var type = (string?)ph.Attribute("type") ?? "obj";
            var idx = (string?)ph.Attribute("idx");
            foreach (var root in new[] { _layoutXml, owner.MasterXml })
            {
                var candidates = root?.Element(P + "cSld")?.Element(P + "spTree")?.Elements().Where(x => x.Descendants(P + "ph").Any()).ToList() ?? [];
                var match = candidates.FirstOrDefault(c => idx is not null && (string?)c.Descendants(P + "ph").First().Attribute("idx") == idx && root == _layoutXml)
                            ?? candidates.FirstOrDefault(c => SameType((string?)c.Descendants(P + "ph").First().Attribute("type") ?? "obj", type));
                if (match is not null)
                {
                    yield return match;
                }
            }
        }

        private static bool SameType(string a, string b)
        {
            static string Norm(string t) => t switch { "ctrTitle" => "title", "obj" or "subTitle" => "body", _ => t };
            return Norm(a) == Norm(b);
        }

        private XElement? InheritedXfrm(XElement? ph)
            => Inherited(ph).Select(x => x.Element(P + "spPr")?.Element(A + "xfrm")).FirstOrDefault(x => x?.Element(A + "off") is not null);

        private JsonObject? Shape(XElement sp, LayoutSlot? slot)
        {
            var spPr = sp.Element(P + "spPr");
            var ph = sp.Element(P + "nvSpPr")?.Element(P + "nvPr")?.Element(P + "ph");
            var phType = (string?)ph?.Attribute("type") ?? (ph is null ? null : "obj");
            if (phType is not null && FooterTypes.Contains(phType))
            {
                Footer(sp, phType);
                return null;
            }

            var prst = (string?)spPr?.Element(A + "prstGeom")?.Attribute("prst");
            if (prst is "line" or "straightConnector1" or "bentConnector2" or "bentConnector3" or "curvedConnector3")
            {
                return Line(sp);
            }

            var e = Geometry(sp, spPr?.Element(A + "xfrm"), slot, ph);
            var isText = BoolAttr(sp.Element(P + "nvSpPr")?.Element(P + "cNvSpPr"), "txBox") || (ph is not null && slot?.Role != "picture");
            var type = isText && (prst is null or "rect") ? "text" : "shape";
            e["type"] = type;
            if (slot is not null)
            {
                e["ph"] = slot.Ph;
            }

            if (type == "shape")
            {
                e["shape"] = spPr is null ? "rect" : ShapeName(spPr);
                if (spPr?.Element(A + "prstGeom")?.Element(A + "avLst")?.Elements(A + "gd").FirstOrDefault() is { } gd
                    && ((string?)gd.Attribute("fmla"))?.Split(' ') is [_, var val] && long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var adj))
                {
                    e["adj"] = Math.Round(Math.Clamp(adj / 100000.0, -10, 10), 4);
                }
            }

            var fill = owner.Fill(spPr, part, ph is null ? owner.StyleColor(sp, "fillRef") : default);
            if (fill is not null)
            {
                if (fill["opacity"] is { } op)
                {
                    e["opacity"] = op.DeepClone();
                    fill.Remove("opacity");
                }

                e["fill"] = fill;
            }

            if (LineOf(spPr?.Element(A + "ln"), ph is null ? owner.StyleColor(sp, "lnRef") : default) is { } line)
            {
                e["line"] = line;
            }

            if (spPr?.Element(A + "effectLst")?.Element(A + "outerShdw") is not null)
            {
                e["shadow"] = true;
            }

            var role = slot?.Role ?? (type == "text" ? "text" : "shape");
            if (sp.Element(P + "txBody") is { } body && (TextBody(body, role, slot, ph, sp) is { } text))
            {
                e["text"] = text;
            }
            else if (type == "text")
            {
                e["text"] = new JsonObject { ["paras"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["t"] = string.Empty }) }) };
            }

            return e;
        }

        private void Footer(XElement sp, string type)
        {
            var text = string.Concat(sp.Descendants(A + "t").Select(t => t.Value)).Trim();
            switch (type)
            {
                case "sldNum":
                    owner._footer["number"] = true;
                    break;
                case "dt":
                    owner._footer["date"] = true;
                    if (!sp.Descendants(A + "fld").Any() && text.Length > 0)
                    {
                        owner._footer["dateText"] = text.Length > 100 ? text[..100] : text;
                    }

                    break;
                case "ftr":
                    if (text.Length > 0)
                    {
                        owner._footer["text"] = text.Length > 300 ? text[..300] : text;
                    }

                    break;
            }
        }

        private JsonObject? LineOf(XElement? ln, (string? Color, int Idx) style)
        {
            if (ln?.Element(A + "noFill") is not null)
            {
                return null;
            }

            var color = ln?.Element(A + "solidFill") is { } solid ? owner.Color(solid).Color : null;
            if (color is null && ln?.Element(A + "gradFill")?.Descendants(A + "gs").FirstOrDefault() is { } gs)
            {
                color = owner.Color(gs).Color;
            }

            if (color is null && style.Idx > 0)
            {
                color = style.Color;
            }

            if (color is null)
            {
                return null;
            }

            var width = ln is null ? 1 : Math.Round(LongAttr(ln, "w", 12700) / 12700.0, 2);
            var dash = (string?)ln?.Element(A + "prstDash")?.Attribute("val") switch
            {
                "dash" or "sysDash" => "dash",
                "dot" or "sysDot" => "dot",
                "dashDot" or "sysDashDot" or "lgDashDot" or "lgDashDotDot" or "sysDashDotDot" => "dashDot",
                "lgDash" => "longDash",
                _ => "solid",
            };
            return new JsonObject { ["color"] = color, ["width"] = Math.Max(0.25, width), ["dash"] = dash };
        }

        private JsonObject? Line(XElement e)
        {
            var spPr = e.Element(P + "spPr");
            var xfrm = spPr?.Element(A + "xfrm");
            if (xfrm is null)
            {
                return null;
            }

            var g = Geometry(e, xfrm, null, null);
            g["type"] = "line";
            var ln = spPr?.Element(A + "ln");
            g["line"] = LineOf(ln, owner.StyleColor(e, "lnRef")) ?? new JsonObject { ["color"] = "dk1", ["width"] = 1, ["dash"] = "solid" };
            g["arrowStart"] = Arrow((string?)ln?.Element(A + "headEnd")?.Attribute("type"));
            g["arrowEnd"] = Arrow((string?)ln?.Element(A + "tailEnd")?.Attribute("type"));
            return g;
        }

        private static string Arrow(string? type) => type switch { "triangle" => "triangle", "arrow" => "arrow", "stealth" => "stealth", "oval" => "oval", "diamond" => "diamond", _ => "none" };

        private JsonObject? Pic(XElement pic, LayoutSlot? slot)
        {
            var rel = (string?)pic.Element(P + "blipFill")?.Element(A + "blip")?.Attribute(R + "embed");
            var src = owner.Picture(part, rel);
            if (src is null)
            {
                return null;
            }

            var spPr = pic.Element(P + "spPr");
            var ph = pic.Element(P + "nvPicPr")?.Element(P + "nvPr")?.Element(P + "ph");
            var e = Geometry(pic, spPr?.Element(A + "xfrm"), slot, ph);
            e["type"] = "image";
            e["src"] = src;
            if (slot?.Role == "picture")
            {
                e["ph"] = slot.Ph;
            }

            var crop = pic.Element(P + "blipFill")?.Element(A + "srcRect");
            if (crop is not null && crop.HasAttributes)
            {
                e["crop"] = new JsonObject
                {
                    ["l"] = IntAttr(crop, "l", 0) / 100000.0, ["t"] = IntAttr(crop, "t", 0) / 100000.0, ["r"] = IntAttr(crop, "r", 0) / 100000.0, ["b"] = IntAttr(crop, "b", 0) / 100000.0,
                };
            }

            if (pic.Element(P + "blipFill")?.Element(A + "blip")?.Element(A + "alphaModFix") is { } alpha)
            {
                e["opacity"] = Math.Round(IntAttr(alpha, "amt", 100000) / 100000.0, 2);
            }

            if (LineOf(spPr?.Element(A + "ln"), default) is { } line)
            {
                e["line"] = line;
            }

            if (spPr?.Element(A + "effectLst")?.Element(A + "outerShdw") is not null)
            {
                e["shadow"] = true;
            }

            var descr = (string?)pic.Element(P + "nvPicPr")?.Element(P + "cNvPr")?.Attribute("descr");
            if (!string.IsNullOrEmpty(descr))
            {
                e["alt"] = descr.Length > 2000 ? descr[..2000] : descr;
            }

            return e;
        }

        private JsonObject? Table(XElement frame)
        {
            var tbl = frame.Descendants(A + "tbl").FirstOrDefault();
            var xfrm = frame.Element(P + "xfrm");
            if (tbl is null || xfrm is null)
            {
                return null;
            }

            var cols = tbl.Element(A + "tblGrid")?.Elements(A + "gridCol").Select(c => Math.Max(4, Px(LongAttr(c, "w", 914400)))).ToList() ?? [];
            if (cols.Count == 0 || cols.Count > SlidesNormalizer.MaxCols)
            {
                return null;
            }

            var e = Geometry(frame, xfrm, null, null);
            e["type"] = "table";
            var rows = new JsonArray();
            var anyFill = false;
            foreach (var tr in tbl.Elements(A + "tr").Take(SlidesNormalizer.MaxRows))
            {
                var cells = new JsonArray();
                foreach (var tc in tr.Elements(A + "tc").Take(cols.Count))
                {
                    var cell = new JsonObject();
                    if (!BoolAttr(tc, "hMerge") && !BoolAttr(tc, "vMerge") && tc.Element(A + "txBody") is { } body && TextBody(body, "cell", null, null, null) is { } text)
                    {
                        cell["text"] = text;
                    }

                    if (tc.Element(A + "tcPr") is { } tcPr && tcPr.Element(A + "solidFill") is not null && owner.Fill(tcPr, part) is { } fill)
                    {
                        fill.Remove("opacity");
                        cell["fill"] = fill;
                        anyFill = true;
                    }

                    cells.Add(cell);
                }

                while (cells.Count < cols.Count)
                {
                    cells.Add(new JsonObject());
                }

                rows.Add(new JsonObject { ["h"] = Math.Max(4, Px(LongAttr(tr, "h", 370840))), ["cells"] = cells });
            }

            if (rows.Count == 0)
            {
                return null;
            }

            var tblPr = tbl.Element(A + "tblPr");
            e["cols"] = new JsonArray(cols.Select(c => (JsonNode)Math.Round(c, 2)).ToArray());
            e["rows"] = rows;
            e["tableStyle"] = anyFill ? "none" : "medium";
            e["accent"] = "accent1";
            e["header"] = BoolAttr(tblPr, "firstRow");
            e["banded"] = BoolAttr(tblPr, "bandRow");
            e["firstCol"] = BoolAttr(tblPr, "firstCol");
            e["lastRow"] = BoolAttr(tblPr, "lastRow");
            return e;
        }

        // ---- text ------------------------------------------------------------------------------------------------------------------------

        /// <summary>The lstStyle levels a paragraph inherits from, nearest first.</summary>
        private List<XElement?> Styles(XElement body, string role, XElement? ph)
        {
            var list = new List<XElement?> { body.Element(A + "lstStyle") };
            foreach (var inherited in Inherited(ph))
            {
                list.Add(inherited.Element(P + "txBody")?.Element(A + "lstStyle"));
            }

            var txStyles = owner.MasterXml?.Element(P + "txStyles");
            var phType = (string?)ph?.Attribute("type") ?? (ph is null ? null : "obj");
            if (phType is "title" or "ctrTitle")
            {
                list.Add(txStyles?.Element(P + "titleStyle"));
            }
            else if (phType is not null)
            {
                list.Add(txStyles?.Element(P + "bodyStyle"));
            }
            else
            {
                list.Add(txStyles?.Element(P + "otherStyle"));
                list.Add(owner._defaultTextStyle);
            }

            _ = role;
            return list;
        }

        private static XElement? Level(XElement? style, int level) => style?.Element(A + $"lvl{level + 1}pPr");

        private static T? Find<T>(List<XElement?> styles, int level, Func<XElement, T?> pick)
        {
            foreach (var s in styles)
            {
                if (Level(s, level) is { } l && pick(l) is { } v)
                {
                    return v;
                }
            }

            return default;
        }

        private JsonObject? TextBody(XElement body, string role, LayoutSlot? slot, XElement? ph, XElement? shape)
        {
            var bodyPr = body.Element(A + "bodyPr");
            var styles = Styles(body, role, ph);
            var theme = owner._theme;
            var fontColor = shape is not null && ph is null ? owner.StyleColor(shape, "fontRef").Color : null;
            var paras = new JsonArray();
            foreach (var p in body.Elements(A + "p").Take(SlidesNormalizer.MaxParagraphs))
            {
                var pPr = p.Element(A + "pPr");
                var level = Math.Clamp(IntAttr(pPr, "lvl", 0), 0, 8);
                var para = new JsonObject();
                var runs = new JsonArray();

                // what the chain says, and what the editor would assume by itself
                var size = Find(styles, level, l => l.Element(A + "defRPr")?.Attribute("sz") is { } a ? (double?)(IntAttr(l.Element(A + "defRPr"), "sz", 1800) / 100.0) : null);
                var defaultSize = role switch
                {
                    "title" => slot?.Size ?? 44,
                    "body" => slot?.Size ?? SlidesData.BodySize(level),
                    _ => 18.0,
                };
                var bold = Find(styles, level, l => l.Element(A + "defRPr")?.Attribute("b") is { } b ? (bool?)(b.Value is "1" or "true") : null) ?? false;
                var color = Find(styles, level, l => l.Element(A + "defRPr")?.Element(A + "solidFill") is { } f ? owner.Color(f).Color : null) ?? fontColor;
                var defaultColor = role switch
                {
                    "title" => theme["heading"]!.GetValue<string>(),
                    "shape" => "lt1",
                    "cell" => null,
                    _ => theme["text"]!.GetValue<string>(),
                };
                if (role is "shape" && color is null)
                {
                    color = "dk1";
                }

                var face = Find(styles, level, l => (string?)l.Element(A + "defRPr")?.Element(A + "latin")?.Attribute("typeface"));
                var align = Find(styles, level, l => (string?)l.Attribute("algn"));
                var bullet = Find(styles, level, BulletOf);

                foreach (var node in p.Elements())
                {
                    if (runs.Count >= SlidesNormalizer.MaxRuns)
                    {
                        break;
                    }

                    if (node.Name == A + "r" || node.Name == A + "fld")
                    {
                        var run = Run(node.Element(A + "rPr"), node.Element(A + "t")?.Value ?? string.Empty);
                        Defaults(run, size, defaultSize, bold, color, defaultColor, face, role);
                        runs.Add(run);
                    }
                    else if (node.Name == A + "br")
                    {
                        if (runs.Count > 0 && runs[^1] is JsonObject last)
                        {
                            last["t"] = last["t"]!.GetValue<string>() + "\n";
                        }
                        else
                        {
                            var run = Run(node.Element(A + "rPr"), "\n");
                            Defaults(run, size, defaultSize, bold, color, defaultColor, face, role);
                            runs.Add(run);
                        }
                    }
                }

                if (runs.Count == 0)
                {
                    var run = Run(p.Element(A + "endParaRPr"), string.Empty);
                    Defaults(run, size, defaultSize, bold, color, defaultColor, face, role);
                    runs.Add(run);
                }

                para["runs"] = runs;
                var algn = (string?)pPr?.Attribute("algn") ?? align;
                var defaultAlign = role == "shape" ? "ctr" : slot is not null ? PptxWriter.AlignOut(slot.Align) : "l";
                if (algn is not null && algn != defaultAlign)
                {
                    para["align"] = algn switch { "ctr" => "center", "r" => "right", "just" or "dist" => "justify", _ => "left" };
                }

                if (level > 0)
                {
                    para["level"] = level;
                }

                var effective = BulletOf(pPr ?? new XElement(A + "pPr")) ?? bullet;
                var bulleted = role == "body" && (slot?.Bullet ?? true);
                var defaultChar = level % 2 == 0 ? "•" : "–";
                if (effective == "none")
                {
                    if (bulleted)
                    {
                        para["bullet"] = "none";
                    }
                }
                else if (effective is not null && (!bulleted || effective != defaultChar))
                {
                    para["bullet"] = effective;
                }

                if (pPr?.Element(A + "lnSpc")?.Element(A + "spcPct") is { } lnSpc)
                {
                    para["lineSpacing"] = Math.Round(IntAttr(lnSpc, "val", 100000) / 100000.0, 2);
                }

                if (pPr?.Element(A + "spcBef")?.Element(A + "spcPts") is { } before)
                {
                    para["spaceBefore"] = IntAttr(before, "val", 0) / 100.0;
                }

                if (pPr?.Element(A + "spcAft")?.Element(A + "spcPts") is { } after)
                {
                    para["spaceAfter"] = IntAttr(after, "val", 0) / 100.0;
                }

                paras.Add(para);
            }

            var result = new JsonObject { ["paras"] = paras };
            var anchor = (string?)bodyPr?.Attribute("anchor") ?? Inherited(ph).Select(x => (string?)x.Element(P + "txBody")?.Element(A + "bodyPr")?.Attribute("anchor")).FirstOrDefault(a => a is not null);
            if (anchor is not null)
            {
                result["anchor"] = anchor switch { "ctr" => "middle", "b" => "bottom", _ => "top" };
            }

            if (bodyPr?.Element(A + "normAutofit") is not null)
            {
                result["autofit"] = "shrink";
            }
            else if (bodyPr?.Element(A + "spAutoFit") is not null)
            {
                result["autofit"] = "grow";
            }
            else if (bodyPr?.Element(A + "noAutofit") is not null)
            {
                result["autofit"] = "none";
            }

            if ((string?)bodyPr?.Attribute("wrap") == "none")
            {
                result["wrap"] = false;
            }

            if (bodyPr?.Attribute("lIns") is not null || bodyPr?.Attribute("tIns") is not null || bodyPr?.Attribute("rIns") is not null || bodyPr?.Attribute("bIns") is not null)
            {
                result["inset"] = new JsonArray(Px(LongAttr(bodyPr, "lIns", 91440)), Px(LongAttr(bodyPr, "tIns", 45720)), Px(LongAttr(bodyPr, "rIns", 91440)), Px(LongAttr(bodyPr, "bIns", 45720)));
            }

            return result;
        }

        private void Defaults(JsonObject run, double? size, double defaultSize, bool bold, string? color, string? defaultColor, string? face, string role)
        {
            if (run["size"] is null && size is { } s && Math.Abs(s - defaultSize) > 0.01)
            {
                run["size"] = s;
            }

            if (run["b"] is null && bold && role != "title")
            {
                run["b"] = true;
            }

            if (run["color"] is null && color is not null && color != defaultColor)
            {
                run["color"] = color;
            }

            if (run["font"] is null && face is not null && FontOf(face) is { } font && font != (role == "title" ? owner._theme["fonts"]!["major"]!.GetValue<string>() : owner._theme["fonts"]!["minor"]!.GetValue<string>()))
            {
                run["font"] = font;
            }
        }

        private string? FontOf(string face) => face switch
        {
            "+mj-lt" or "+mj-ea" or "+mj-cs" => owner._theme["fonts"]!["major"]!.GetValue<string>(),
            "+mn-lt" or "+mn-ea" or "+mn-cs" => owner._theme["fonts"]!["minor"]!.GetValue<string>(),
            _ => face,
        };

        private static string? BulletOf(XElement pPr)
        {
            if (pPr.Element(A + "buNone") is not null)
            {
                return "none";
            }

            if (pPr.Element(A + "buChar") is { } ch)
            {
                var c = (string?)ch.Attribute("char");
                if (string.IsNullOrEmpty(c))
                {
                    return "•";
                }

                // Wingdings bullets have no meaning outside the symbol font: the usual ones as plain characters
                var font = (string?)pPr.Element(A + "buFont")?.Attribute("typeface") ?? string.Empty;
                if (font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase) || font.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase))
                {
                    return c switch { "§" => "■", "Ø" => "➢", "ü" => "✓", "q" => "□", "v" => "◆", "n" => "■", "l" => "●", _ => "•" };
                }

                return c.Length > 2 ? c[..2] : c;
            }

            if (pPr.Element(A + "buAutoNum") is { } num)
            {
                return (string?)num.Attribute("type") switch
                {
                    "arabicParenR" or "arabicParenBoth" => "1)",
                    "alphaLcPeriod" => "a.",
                    "alphaLcParenR" or "alphaLcParenBoth" => "a)",
                    "alphaUcPeriod" => "A.",
                    "alphaUcParenR" or "alphaUcParenBoth" => "A)",
                    "romanUcPeriod" or "romanUcParenR" => "I.",
                    "romanLcPeriod" or "romanLcParenR" => "i.",
                    _ => "1.",
                };
            }

            return null;
        }

        private JsonObject Run(XElement? rPr, string text)
        {
            var run = new JsonObject { ["t"] = text.Length > SlidesNormalizer.MaxRunText ? text[..SlidesNormalizer.MaxRunText] : text };
            if (rPr is null)
            {
                return run;
            }

            if (rPr.Attribute("sz") is not null)
            {
                run["size"] = IntAttr(rPr, "sz", 1800) / 100.0;
            }

            if (rPr.Attribute("b") is { } b)
            {
                run["b"] = b.Value is "1" or "true";
            }

            if (BoolAttr(rPr, "i"))
            {
                run["i"] = true;
            }

            if ((string?)rPr.Attribute("u") is { } u && u != "none")
            {
                run["u"] = true;
            }

            if ((string?)rPr.Attribute("strike") is "sngStrike" or "dblStrike")
            {
                run["s"] = true;
            }

            if (rPr.Element(A + "solidFill") is { } fill && owner.Color(fill).Color is { } color)
            {
                run["color"] = color;
            }

            if (rPr.Element(A + "highlight") is { } hl && owner.Color(hl).Color is { } hc)
            {
                run["hl"] = hc;
            }

            if ((string?)rPr.Element(A + "latin")?.Attribute("typeface") is { Length: > 0 } face && FontOf(face) is { } font)
            {
                run["font"] = font;
            }

            return run;
        }

        // ---- notes, transition, animations --------------------------------------------------------------------------------------------

        private string Notes()
        {
            if (part.NotesSlidePart is not { } notes)
            {
                return string.Empty;
            }

            var root = Load(notes);
            var body = root.Descendants(P + "sp").FirstOrDefault(sp => sp.Descendants(P + "ph").Any(ph => (string?)ph.Attribute("type") == "body"));
            if (body is null)
            {
                return string.Empty;
            }

            var lines = body.Descendants(A + "p").Select(p => string.Concat(p.Elements().Select(n => n.Name == A + "br" ? "\n" : n.Element(A + "t")?.Value ?? string.Empty)));
            var text = string.Join('\n', lines).TrimEnd();
            return text.Length > SlidesNormalizer.MaxNotes ? text[..SlidesNormalizer.MaxNotes] : text;
        }

        private static JsonObject? Transition(XElement root)
        {
            var tr = root.Element(P + "transition");
            if (tr is null && root.Element(Mc + "AlternateContent") is { } alt)
            {
                tr = alt.Elements(Mc + "Choice").Concat(alt.Elements(Mc + "Fallback")).SelectMany(x => x.Elements()).FirstOrDefault(x => x.Name.LocalName == "transition");
            }

            if (tr is null)
            {
                return null;
            }

            var effect = tr.Elements().FirstOrDefault(x => x.Name.LocalName is not ("sndAc" or "extLst"));
            var type = effect?.Name.LocalName switch
            {
                "fade" or "dissolve" or "random" or "morph" => "fade",
                "push" or "pull" or "pan" => "push",
                "wipe" or "blinds" or "strips" or "randomBar" or "checker" or "comb" => "wipe",
                "split" or "barn" or "doors" => "split",
                "cover" or "uncover" => "cover",
                "cut" => null,
                null => null,
                _ => "zoom",
            };
            var duration = (string?)tr.Attribute("spd") switch { "fast" => 0.5, "slow" => 1.0, _ => 0.75 };
            var durAttr = tr.Attributes().FirstOrDefault(a => a.Name.LocalName == "dur");
            if (durAttr is not null && int.TryParse(durAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ms))
            {
                duration = Math.Clamp(ms / 1000.0, 0.1, 10);
            }

            var result = new JsonObject { ["type"] = type ?? "none", ["duration"] = duration };
            var dir = (string?)effect?.Attribute("dir");
            var mapped = dir switch { "r" => "left", "d" => "up", "u" => "down", "l" => "right", _ => null };
            if (mapped is not null && type is "push" or "wipe" or "cover")
            {
                result["dir"] = mapped;
            }

            if (tr.Attribute("advTm") is { } adv && int.TryParse(adv.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var advMs))
            {
                result["auto"] = Math.Round(advMs / 1000.0, 2);
            }

            return type is null && result["auto"] is null ? null : result;
        }

        private JsonArray Animations(XElement root)
        {
            var list = new JsonArray();
            var main = root.Element(P + "timing")?.Descendants(P + "cTn").FirstOrDefault(c => (string?)c.Attribute("nodeType") == "mainSeq");
            if (main is null)
            {
                return list;
            }

            foreach (var ctn in main.Descendants(P + "cTn").Where(c => (string?)c.Attribute("presetClass") == "entr"))
            {
                var spid = (string?)ctn.Descendants(P + "spTgt").FirstOrDefault()?.Attribute("spid");
                if (spid is null || !_ids.TryGetValue(spid, out var target) || list.Count >= SlidesNormalizer.MaxAnimations)
                {
                    continue;
                }

                if (list.OfType<JsonObject>().Any(a => a["el"]!.GetValue<string>() == target))
                {
                    continue;
                }

                var preset = IntAttr(ctn, "presetID", 10);
                var subtype = IntAttr(ctn, "presetSubtype", 0);
                var effect = preset switch { 1 => "appear", 2 or 12 or 7 => "fly", 22 or 16 or 18 => "wipe", 23 or 53 or 55 or 31 or 37 => "zoom", _ => "fade" };
                var trigger = (string?)ctn.Attribute("nodeType") switch { "withEffect" => "with", "afterEffect" => "after", _ => "click" };
                var durations = ctn.Descendants(P + "cTn").Select(c => (string?)c.Attribute("dur")).Where(d => d is not null && d != "indefinite" && d != "1")
                    .Select(d => int.TryParse(d, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 500).DefaultIfEmpty(500).Max();
                var a = new JsonObject { ["el"] = target, ["effect"] = effect, ["trigger"] = trigger, ["duration"] = Math.Clamp(durations / 1000.0, 0.05, 20) };
                if (effect is "fly" or "wipe")
                {
                    a["dir"] = subtype switch { 1 => "top", 2 => "right", 8 => "left", _ => "bottom" };
                }

                list.Add(a);
            }

            return list;
        }
    }
}
