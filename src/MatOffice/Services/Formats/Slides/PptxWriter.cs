using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using static MatOffice.Services.Formats.Slides.PptxXml;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// Writes a presentation as a PowerPoint file (.pptx) with DocumentFormat.OpenXml: a slide master with the design (theme colours and fonts, background,
/// decoration, text styles), the eight layouts with their placeholders, and the slides with text boxes and placeholders (runs, bullets, alignment),
/// shapes (preset geometry with fill and line), lines, pictures, tables, groups, notes, backgrounds, transitions and entrance animations.
/// The parts are built with the SDK, their XML with LINQ to XML (checked by <c>OpenXmlValidator</c> in the tests).
/// </summary>
public sealed class PptxWriter
{
    private readonly JsonObject _content;
    private readonly JsonObject _theme;
    private readonly double _w;
    private readonly double _h;
    private readonly Dictionary<string, SlideLayoutPart> _layoutParts = new(StringComparer.Ordinal);

    private PptxWriter(JsonObject content)
    {
        _content = content;
        _theme = content["theme"]!.AsObject();
        _w = content["size"]!["w"]!.GetValue<double>();
        _h = content["size"]!["h"]!.GetValue<double>();
    }

    /// <summary>The file of a (normalized) content.</summary>
    public static byte[] Write(JsonObject content, string name)
    {
        try
        {
            return new PptxWriter(content).Build(name);
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or FormatException or OpenXmlPackageException)
        {
            throw new Formats.FormatException(Localization.Text.T("Slides.Export.Failed"));
        }
    }

    private byte[] Build(string name)
    {
        using var stream = new MemoryStream();
        using (var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentation = doc.AddPresentationPart();
            Feed(doc.AddCoreFilePropertiesPart(), CoreXml(name));
            Feed(doc.AddExtendedFilePropertiesPart(), AppXml());

            var master = presentation.AddNewPart<SlideMasterPart>();
            var theme = master.AddNewPart<ThemePart>();
            Feed(theme, ThemeXml());
            presentation.AddPart(theme);

            var layoutIds = new List<(uint Id, string RelId)>();
            uint nextLayoutId = 2147483649;
            foreach (var layout in SlidesData.Layouts)
            {
                var part = master.AddNewPart<SlideLayoutPart>();
                part.AddPart(master);
                _layoutParts[layout.Id] = part;
                layoutIds.Add((nextLayoutId++, master.GetIdOfPart(part)));
            }

            Feed(master, MasterXml(master, layoutIds));
            foreach (var layout in SlidesData.Layouts)
            {
                var part = _layoutParts[layout.Id];
                Feed(part, LayoutXml(layout, part));
            }

            var notesMaster = presentation.AddNewPart<NotesMasterPart>();
            var notesTheme = notesMaster.AddNewPart<ThemePart>();
            Feed(notesTheme, ThemeXml());
            Feed(notesMaster, NotesMasterXml());

            var slideIds = new List<(uint Id, string RelId)>();
            uint nextSlideId = 256;
            var number = 0;
            foreach (var slide in _content["slides"]!.AsArray().OfType<JsonObject>())
            {
                number++;
                var part = presentation.AddNewPart<SlidePart>();
                var layoutId = slide["layout"]?.GetValue<string>() ?? "blank";
                part.AddPart(_layoutParts[SlidesData.Layout(layoutId).Id]);
                Feed(part, new SlideWriter(this, part, slide, number).Xml());
                var notes = slide["notes"]?.GetValue<string>();
                if (!string.IsNullOrEmpty(notes))
                {
                    var notesPart = part.AddNewPart<NotesSlidePart>();
                    notesPart.AddPart(notesMaster);
                    notesPart.AddPart(part);
                    Feed(notesPart, NotesXml(notes));
                }

                slideIds.Add((nextSlideId++, presentation.GetIdOfPart(part)));
            }

            Feed(presentation.AddNewPart<PresentationPropertiesPart>(), new XDocument(new XElement(P + "presentationPr", Ns())));
            Feed(presentation.AddNewPart<ViewPropertiesPart>(), new XDocument(new XElement(P + "viewPr", Ns(),
                new XElement(P + "gridSpacing", new XAttribute("cx", "76200"), new XAttribute("cy", "76200")))));
            Feed(presentation.AddNewPart<TableStylesPart>(), new XDocument(new XElement(A + "tblStyleLst", new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                new XAttribute("def", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"))));

            Feed(presentation, PresentationXml(presentation.GetIdOfPart(master), presentation.GetIdOfPart(notesMaster), slideIds));
        }

        return stream.ToArray();
    }

    // ---- package parts ---------------------------------------------------------------------------------------------------------------

    private static void Feed(OpenXmlPart part, XDocument xml)
    {
        using var ms = new MemoryStream();
        using (var writer = System.Xml.XmlWriter.Create(ms, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
        {
            xml.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
            xml.Save(writer);
        }

        ms.Position = 0;
        part.FeedData(ms);
    }

    private static object[] Ns() =>
    [
        new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
    ];

    private static XDocument CoreXml(string name)
    {
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        var now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        return new XDocument(new XElement(cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName), new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName), new XAttribute(XNamespace.Xmlns + "xsi", xsi.NamespaceName),
            new XElement(dc + "title", name),
            new XElement(dc + "creator", "MatOffice"),
            new XElement(dcterms + "created", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now),
            new XElement(dcterms + "modified", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now)));
    }

    private XDocument AppXml()
    {
        XNamespace ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        return new XDocument(new XElement(ep + "Properties", new XAttribute("xmlns", ep.NamespaceName),
            new XElement(ep + "Application", "MatOffice Slides"),
            new XElement(ep + "Slides", _content["slides"]!.AsArray().Count),
            new XElement(ep + "PresentationFormat", Math.Abs((_w / _h) - (16.0 / 9)) < 0.01 ? "Widescreen" : Math.Abs((_w / _h) - (4.0 / 3)) < 0.01 ? "On-screen Show (4:3)" : "Custom")));
    }

    private XDocument PresentationXml(string masterRel, string notesMasterRel, List<(uint Id, string RelId)> slides)
    {
        var sldSz = new XElement(P + "sldSz", new XAttribute("cx", S(Emu(_w))), new XAttribute("cy", S(Emu(_h))));
        if (Math.Abs(_w - 960) < 0.5 && Math.Abs(_h - 720) < 0.5)
        {
            sldSz.Add(new XAttribute("type", "screen4x3"));
        }

        var root = new XElement(P + "presentation", Ns(), new XAttribute("saveSubsetFonts", "1"),
            new XElement(P + "sldMasterIdLst", new XElement(P + "sldMasterId", new XAttribute("id", "2147483648"), new XAttribute(R + "id", masterRel))),
            new XElement(P + "notesMasterIdLst", new XElement(P + "notesMasterId", new XAttribute(R + "id", notesMasterRel))));
        if (slides.Count > 0)
        {
            root.Add(new XElement(P + "sldIdLst", slides.Select(s => new XElement(P + "sldId", new XAttribute("id", s.Id.ToString(CultureInfo.InvariantCulture)), new XAttribute(R + "id", s.RelId)))));
        }

        root.Add(sldSz, new XElement(P + "notesSz", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")));
        root.Add(new XElement(P + "defaultTextStyle", LevelStyles(_ => 1800, null, "+mn-lt", level => (long)(level * 457200), _ => 0).Elements()));
        return new XDocument(root);
    }

    // ---- theme -------------------------------------------------------------------------------------------------------------------------

    private XDocument ThemeXml()
    {
        var name = _theme["name"]?.GetValue<string>() ?? "MatOffice";
        var colors = _theme["colors"]!.AsObject();
        XElement Slot(string slot, string key) => new(A + slot, new XElement(A + "srgbClr", new XAttribute("val", colors[key]!.GetValue<string>()[1..].ToUpperInvariant())));
        XElement Ph(params XElement[] mods) => new(A + "schemeClr", new XAttribute("val", "phClr"), mods);
        XElement Mod(string mod, int val) => new(A + mod, new XAttribute("val", S(val)));
        XElement Grad(XElement a, XElement b) => new(A + "gradFill", new XAttribute("rotWithShape", "1"),
            new XElement(A + "gsLst", new XElement(A + "gs", new XAttribute("pos", "0"), a), new XElement(A + "gs", new XAttribute("pos", "100000"), b)),
            new XElement(A + "lin", new XAttribute("ang", "5400000"), new XAttribute("scaled", "0")));
        XElement Ln(int w) => new(A + "ln", new XAttribute("w", S(w)), new XAttribute("cap", "flat"), new XAttribute("cmpd", "sng"), new XAttribute("algn", "ctr"),
            new XElement(A + "solidFill", Ph()), new XElement(A + "prstDash", new XAttribute("val", "solid")), new XElement(A + "miter", new XAttribute("lim", "800000")));
        XElement Font(string kind, string face) => new(A + kind, new XElement(A + "latin", new XAttribute("typeface", face)), new XElement(A + "ea", new XAttribute("typeface", "")), new XElement(A + "cs", new XAttribute("typeface", "")));

        var fonts = _theme["fonts"]!;
        var theme = new XElement(A + "theme", new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName), new XAttribute("name", name),
            new XElement(A + "themeElements",
                new XElement(A + "clrScheme", new XAttribute("name", name),
                    Slot("dk1", "dk1"), Slot("lt1", "lt1"), Slot("dk2", "dk2"), Slot("lt2", "lt2"),
                    Slot("accent1", "accent1"), Slot("accent2", "accent2"), Slot("accent3", "accent3"), Slot("accent4", "accent4"), Slot("accent5", "accent5"), Slot("accent6", "accent6"),
                    Slot("hlink", "hlink"), new XElement(A + "folHlink", new XElement(A + "srgbClr", new XAttribute("val", Lum(colors["hlink"]!.GetValue<string>(), -25)[1..].ToUpperInvariant())))),
                new XElement(A + "fontScheme", new XAttribute("name", name), Font("majorFont", fonts["major"]!.GetValue<string>()), Font("minorFont", fonts["minor"]!.GetValue<string>())),
                new XElement(A + "fmtScheme", new XAttribute("name", name),
                    new XElement(A + "fillStyleLst",
                        new XElement(A + "solidFill", Ph()),
                        Grad(Ph(Mod("lumMod", 110000), Mod("satMod", 105000), Mod("tint", 67000)), Ph(Mod("lumMod", 105000), Mod("satMod", 109000), Mod("tint", 81000))),
                        Grad(Ph(Mod("satMod", 103000), Mod("lumMod", 102000), Mod("tint", 94000)), Ph(Mod("lumMod", 99000), Mod("satMod", 120000), Mod("shade", 78000)))),
                    new XElement(A + "lnStyleLst", Ln(6350), Ln(12700), Ln(19050)),
                    new XElement(A + "effectStyleLst",
                        new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                        new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                        new XElement(A + "effectStyle", new XElement(A + "effectLst", Shadow()))),
                    new XElement(A + "bgFillStyleLst",
                        new XElement(A + "solidFill", Ph()),
                        new XElement(A + "solidFill", Ph(Mod("tint", 95000), Mod("satMod", 170000))),
                        Grad(Ph(Mod("tint", 93000), Mod("satMod", 150000), Mod("shade", 98000), Mod("lumMod", 102000)), Ph(Mod("shade", 63000), Mod("satMod", 120000)))))),
            new XElement(A + "objectDefaults"),
            new XElement(A + "extraClrSchemeLst"));
        return new XDocument(theme);
    }

    internal static XElement Shadow() => new(A + "outerShdw", new XAttribute("blurRad", "50800"), new XAttribute("dist", "38100"), new XAttribute("dir", "2700000"),
        new XAttribute("algn", "tl"), new XAttribute("rotWithShape", "0"), new XElement(A + "prstClr", new XAttribute("val", "black"), new XElement(A + "alpha", new XAttribute("val", "40000"))));

    // ---- master and layouts --------------------------------------------------------------------------------------------------------

    private string BaseColor(string key) => _theme[key]?.GetValue<string>() ?? "dk1";

    private string GroupColor(string group, string key)
        => _theme["groups"]?[group]?[key]?.GetValue<string>() ?? _theme[key]?.GetValue<string>() ?? "dk1";

    private XDocument MasterXml(SlideMasterPart master, List<(uint Id, string RelId)> layouts)
    {
        var ids = new ShapeIds();
        var tree = SpTree();
        foreach (var d in Decor(l => l.Contains("all"), master, ids))
        {
            tree.Add(d);
        }

        var full = new { x = 0.06875, w = 0.8625 };
        tree.Add(PlaceholderSp(ids.Next(), "Title Placeholder", "title", null, Frame(full.x, 0.0532, full.w, 0.1933), "ctr",
            [Para("Click to edit the master title style")]));
        tree.Add(PlaceholderSp(ids.Next(), "Text Placeholder", "body", 1, Frame(full.x, 0.2662, full.w, 0.6345), "t",
            [Para("Click to edit the master text styles"), Para("Second level", 1), Para("Third level", 2)]));
        tree.Add(FooterSp(ids.Next(), "dt", 2, SlidesData.FooterDate, null));
        tree.Add(FooterSp(ids.Next(), "ftr", 3, SlidesData.FooterText, null));
        tree.Add(FooterSp(ids.Next(), "sldNum", 4, SlidesData.FooterNumber, null));

        var fonts = (Major: "+mj-lt", Minor: "+mn-lt");
        var (indent, step) = SlidesData.BodyIndent;
        var heading = BaseColor("heading");
        var text = BaseColor("text");
        var title = new XElement(P + "titleStyle", new XElement(A + "lvl1pPr", new XAttribute("algn", "l"), new XAttribute("defTabSz", "914400"), new XAttribute("rtl", "0"),
            new XAttribute("eaLnBrk", "1"), new XAttribute("latinLnBrk", "0"), new XAttribute("hangingPunct", "1"),
            new XElement(A + "lnSpc", new XElement(A + "spcPct", new XAttribute("val", S((int)(SlidesData.TitleLineSpacing * 100000))))),
            new XElement(A + "spcBef", new XElement(A + "spcPct", new XAttribute("val", "0"))),
            new XElement(A + "buNone"),
            DefRPr(4400, heading, fonts.Major)));
        var body = new XElement(P + "bodyStyle", Enumerable.Range(0, 9).Select(level => new XElement(A + $"lvl{level + 1}pPr",
            new XAttribute("marL", S(Emu(indent + (step * level)))), new XAttribute("indent", S(-Emu(indent))), new XAttribute("algn", "l"), new XAttribute("defTabSz", "914400"),
            new XAttribute("rtl", "0"), new XAttribute("eaLnBrk", "1"), new XAttribute("latinLnBrk", "0"), new XAttribute("hangingPunct", "1"),
            new XElement(A + "lnSpc", new XElement(A + "spcPct", new XAttribute("val", "100000"))),
            new XElement(A + "spcBef", new XElement(A + "spcPts", new XAttribute("val", S((int)(SlidesData.BodySpaceBefore(level) * 100))))),
            new XElement(A + "buFont", new XAttribute("typeface", "Arial"), new XAttribute("panose", "020B0604020202020204"), new XAttribute("pitchFamily", "34"), new XAttribute("charset", "0")),
            new XElement(A + "buChar", new XAttribute("char", level % 2 == 0 ? "•" : "–")),
            DefRPr((int)(SlidesData.BodySize(level) * 100), text, fonts.Minor))));
        var other = new XElement(P + "otherStyle",
            new XElement(A + "defPPr", new XElement(A + "defRPr", new XAttribute("lang", "en-US"))),
            LevelStyles(_ => 1800, "dk1", fonts.Minor, level => (long)(level * 457200), _ => 0).Elements());

        var root = new XElement(P + "sldMaster", Ns(),
            new XElement(P + "cSld", Background(_theme["background"] as JsonObject, master), tree),
            ClrMap(),
            new XElement(P + "sldLayoutIdLst", layouts.Select(l => new XElement(P + "sldLayoutId", new XAttribute("id", l.Id.ToString(CultureInfo.InvariantCulture)), new XAttribute(R + "id", l.RelId)))),
            new XElement(P + "txStyles", title, body, other));
        return new XDocument(root);
    }

    private static XElement LevelStyles(Func<int, int> size, string? color, string font, Func<int, long> marL, Func<int, long> indentOf)
    {
        var holder = new XElement(P + "holder");
        for (var level = 0; level < 9; level++)
        {
            holder.Add(new XElement(A + $"lvl{level + 1}pPr", new XAttribute("marL", S(marL(level))), new XAttribute("indent", S(indentOf(level))), new XAttribute("algn", "l"),
                new XAttribute("defTabSz", "914400"), new XAttribute("rtl", "0"), new XAttribute("eaLnBrk", "1"), new XAttribute("latinLnBrk", "0"), new XAttribute("hangingPunct", "1"),
                DefRPr(size(level), color, font)));
        }

        return holder;
    }

    private static XElement DefRPr(int size, string? color, string font, bool bold = false)
    {
        var r = new XElement(A + "defRPr", new XAttribute("sz", S(size)), new XAttribute("kern", "1200"));
        if (bold)
        {
            r.Add(new XAttribute("b", "1"));
        }

        if (color is not null)
        {
            r.Add(SolidFill(color));
        }

        r.Add(new XElement(A + "latin", new XAttribute("typeface", font)), new XElement(A + "ea", new XAttribute("typeface", font.Replace("-lt", "-ea", StringComparison.Ordinal))),
            new XElement(A + "cs", new XAttribute("typeface", font.Replace("-lt", "-cs", StringComparison.Ordinal))));
        return r;
    }

    private static XElement ClrMap() => new(P + "clrMap",
        new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"), new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
        new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"), new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
        new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"), new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink"));

    private XDocument LayoutXml(SlideLayout layout, SlideLayoutPart part)
    {
        var ids = new ShapeIds();
        var tree = SpTree();
        foreach (var d in Decor(on => on.Contains(layout.Id) || on.Contains(layout.Group), part, ids))
        {
            tree.Add(d);
        }

        var heading = GroupColor(layout.Group, "heading");
        var text = GroupColor(layout.Group, "text");
        foreach (var slot in layout.Slots)
        {
            tree.Add(LayoutPlaceholder(ids.Next(), slot, heading, text));
        }

        var footerColor = text + ":+40";
        tree.Add(FooterSp(ids.Next(), "dt", 10, SlidesData.FooterDate, footerColor));
        tree.Add(FooterSp(ids.Next(), "ftr", 11, SlidesData.FooterText, footerColor));
        tree.Add(FooterSp(ids.Next(), "sldNum", 12, SlidesData.FooterNumber, footerColor));

        var cSld = new XElement(P + "cSld", new XAttribute("name", LayoutName(layout.Id)));
        if (_theme["groups"]?[layout.Group]?["background"] is JsonObject bg)
        {
            cSld.Add(Background(bg, part));
        }

        cSld.Add(tree);
        var root = new XElement(P + "sldLayout", Ns(), new XAttribute("type", layout.Pptx), new XAttribute("preserve", "1"),
            cSld, new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping")));
        return new XDocument(root);
    }

    internal static string LayoutName(string id) => id switch
    {
        "title" => "Title Slide",
        "titleContent" => "Title and Content",
        "section" => "Section Header",
        "twoContent" => "Two Content",
        "comparison" => "Comparison",
        "titleOnly" => "Title Only",
        "pictureCaption" => "Picture with Caption",
        _ => "Blank",
    };

    private XElement LayoutPlaceholder(int id, LayoutSlot slot, string heading, string text)
    {
        var ph = new XElement(P + "ph");
        if (slot.Pptx != "obj")
        {
            ph.Add(new XAttribute("type", slot.Pptx));
        }

        if (slot.Idx > 0)
        {
            ph.Add(new XAttribute("idx", S(slot.Idx)));
        }

        var isTitle = slot.Role == "title";
        var level = new XElement(A + "lvl1pPr");
        if (slot.Align != "left")
        {
            level.Add(new XAttribute("algn", AlignOut(slot.Align)));
        }

        if (!isTitle && !slot.Bullet)
        {
            level.Add(new XAttribute("marL", "0"), new XAttribute("indent", "0"));
            level.Add(new XElement(A + "buNone"));
        }

        var color = isTitle ? heading : slot.Muted ? text + ":+25" : text;
        var defaultColor = isTitle ? BaseColor("heading") : BaseColor("text");
        var rPr = new XElement(A + "defRPr");
        if (slot.Size is { } size)
        {
            rPr.Add(new XAttribute("sz", S((int)(size * 100))));
        }

        if (slot.Bold)
        {
            rPr.Add(new XAttribute("b", "1"));
        }

        if (color != defaultColor)
        {
            rPr.Add(SolidFill(color));
        }

        if (rPr.HasAttributes || rPr.HasElements)
        {
            level.Add(rPr);
        }

        var lst = new XElement(A + "lstStyle");
        if (level.HasAttributes || level.HasElements)
        {
            lst.Add(level);
        }

        var prompt = slot.Role switch
        {
            "title" => "Click to edit the master title style",
            "picture" => "Click the icon to add a picture",
            _ => slot.Ph == "subtitle" ? "Click to edit the master subtitle style" : "Click to edit the master text styles",
        };
        var nameOf = slot.Role switch { "title" => "Title", "picture" => "Picture Placeholder", _ => slot.Ph == "subtitle" ? "Subtitle" : "Text Placeholder" };
        return new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", $"{nameOf} {id}")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))), new XElement(P + "nvPr", ph)),
            new XElement(P + "spPr", Xfrm(slot.X * _w, slot.Y * _h, slot.W * _w, slot.H * _h, 0, false, false)),
            new XElement(P + "txBody", new XElement(A + "bodyPr", new XAttribute("anchor", AnchorOut(slot.Anchor)), new XElement(A + "normAutofit")), lst,
                new XElement(A + "p", new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", prompt)))));
    }

    private XElement Frame(double x, double y, double w, double h) => Xfrm(x * _w, y * _h, w * _w, h * _h, 0, false, false);

    private static XElement Para(string text, int level = 0)
    {
        var p = new XElement(A + "p");
        if (level > 0)
        {
            p.Add(new XElement(A + "pPr", new XAttribute("lvl", S(level))));
        }

        p.Add(new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", text)));
        return p;
    }

    private static XElement PlaceholderSp(int id, string name, string type, int? idx, XElement xfrm, string anchor, XElement[] paras)
    {
        var ph = new XElement(P + "ph", new XAttribute("type", type));
        if (idx is { } i)
        {
            ph.Add(new XAttribute("idx", S(i)));
        }

        return new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", $"{name} {id}")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))), new XElement(P + "nvPr", ph)),
            new XElement(P + "spPr", xfrm, new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))),
            new XElement(P + "txBody", new XElement(A + "bodyPr", new XAttribute("vert", "horz"), new XAttribute("lIns", "91440"), new XAttribute("tIns", "45720"),
                new XAttribute("rIns", "91440"), new XAttribute("bIns", "45720"), new XAttribute("rtlCol", "0"), new XAttribute("anchor", anchor), new XElement(A + "normAutofit")),
                new XElement(A + "lstStyle"), paras));
    }

    private XElement FooterSp(int id, string type, int idx, FooterSlot slot, string? color)
    {
        var rPr = new XElement(A + "defRPr", new XAttribute("sz", S((int)(SlidesData.FooterSize * 100))));
        if (color is not null)
        {
            rPr.Add(SolidFill(color));
        }

        var name = type switch { "dt" => "Date Placeholder", "ftr" => "Footer Placeholder", _ => "Slide Number Placeholder" };
        var p = new XElement(A + "p");
        if (type == "sldNum")
        {
            p.Add(new XElement(A + "fld", new XAttribute("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}"), new XAttribute("type", "slidenum"),
                new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", "‹#›")));
        }
        else
        {
            p.Add(new XElement(A + "endParaRPr", new XAttribute("lang", "en-US")));
        }

        return new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", $"{name} {id}")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", type), new XAttribute("sz", type == "ftr" ? "quarter" : "half"), new XAttribute("idx", S(idx))))),
            new XElement(P + "spPr", Xfrm(slot.X * _w, slot.Y * _h, slot.W * _w, slot.H * _h, 0, false, false)),
            new XElement(P + "txBody", new XElement(A + "bodyPr", new XAttribute("anchor", "ctr")),
                new XElement(A + "lstStyle", new XElement(A + "lvl1pPr", new XAttribute("algn", AlignOut(slot.Align)), rPr)), p));
    }

    /// <summary>The decoration of the design as shapes of the master or a layout.</summary>
    private IEnumerable<XElement> Decor(Func<IReadOnlyList<string>, bool> on, OpenXmlPart part, ShapeIds ids)
    {
        foreach (var d in (_theme["decor"] as JsonArray ?? []).OfType<JsonObject>())
        {
            var targets = (d["on"] as JsonArray ?? []).Select(x => x!.GetValue<string>()).ToList();
            if (!on(targets))
            {
                continue;
            }

            var x = d["x"]!.GetValue<double>() * _w;
            var y = d["y"]!.GetValue<double>() * _h;
            var w = d["w"]!.GetValue<double>() * _w;
            var h = d["h"]!.GetValue<double>() * _h;
            var rot = d["rot"]?.GetValue<double>() ?? 0;
            var flipH = d["flipH"]?.GetValue<bool>() ?? false;
            var flipV = d["flipV"]?.GetValue<bool>() ?? false;
            var id = ids.Next();
            if (d["src"]?.GetValue<string>() is { } src && AddImage(part, src) is { } rel)
            {
                yield return Picture(id, "Decoration " + id, null, rel, null, Xfrm(x, y, w, h, rot, flipH, flipV), d["opacity"]?.GetValue<double>() ?? 1, null, false);
                continue;
            }

            var alpha = d["opacity"]?.GetValue<double>() ?? 1;
            yield return new XElement(P + "sp",
                new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", "Decoration " + id)), new XElement(P + "cNvSpPr"), new XElement(P + "nvPr", new XAttribute("userDrawn", "1"))),
                new XElement(P + "spPr", Xfrm(x, y, w, h, rot, flipH, flipV), Geometry(d["shape"]?.GetValue<string>() ?? "rect", null),
                    SolidFill(d["fill"]?.GetValue<string>() ?? "accent1", alpha), new XElement(A + "ln", new XElement(A + "noFill"))));
        }
    }

    private XDocument NotesMasterXml()
    {
        var tree = SpTree();
        tree.Add(new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Slide Image Placeholder 1")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"), new XAttribute("noRot", "1"), new XAttribute("noChangeAspect", "1"))),
                new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", "sldImg"), new XAttribute("idx", "2")))),
            new XElement(P + "spPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", "685800"), new XAttribute("y", "1143000")), new XElement(A + "ext", new XAttribute("cx", "5486400"), new XAttribute("cy", "3086100"))),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")), new XElement(A + "noFill"),
                new XElement(A + "ln", new XAttribute("w", "12700"), new XElement(A + "solidFill", new XElement(A + "prstClr", new XAttribute("val", "black")))))));
        tree.Add(new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", "3"), new XAttribute("name", "Notes Placeholder 2")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", "body"), new XAttribute("sz", "quarter"), new XAttribute("idx", "3")))),
            new XElement(P + "spPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", "685800"), new XAttribute("y", "4400550")), new XElement(A + "ext", new XAttribute("cx", "5486400"), new XAttribute("cy", "3600450"))),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))),
            new XElement(P + "txBody", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"), Para("Click to edit the master text styles"))));
        var style = new XElement(P + "notesStyle", LevelStyles(_ => 1200, "dk1", "+mn-lt", level => (long)(level * 457200), _ => 0).Elements());
        return new XDocument(new XElement(P + "notesMaster", Ns(), new XElement(P + "cSld", tree), ClrMap(), style));
    }

    private static XDocument NotesXml(string notes)
    {
        var tree = SpTree();
        tree.Add(new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Slide Image Placeholder 1")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"), new XAttribute("noRot", "1"), new XAttribute("noChangeAspect", "1"))),
                new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", "sldImg")))),
            new XElement(P + "spPr")));
        var body = new XElement(P + "txBody", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"));
        foreach (var line in notes.Split('\n'))
        {
            body.Add(line.Length == 0
                ? new XElement(A + "p", new XElement(A + "endParaRPr", new XAttribute("lang", "en-US")))
                : new XElement(A + "p", new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", line))));
        }

        tree.Add(new XElement(P + "sp",
            new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", "3"), new XAttribute("name", "Notes Placeholder 2")),
                new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", "body"), new XAttribute("idx", "1")))),
            new XElement(P + "spPr"), body));
        return new XDocument(new XElement(P + "notes", Ns(), new XElement(P + "cSld", tree), new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping"))));
    }

    // ---- shared pieces ---------------------------------------------------------------------------------------------------------------

    internal static XElement SpTree() => new(P + "spTree",
        new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")),
        new XElement(P + "grpSpPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")), new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
            new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")), new XElement(A + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0")))));

    internal static XElement Xfrm(double x, double y, double w, double h, double rot, bool flipH, bool flipV, XName? name = null)
    {
        var xfrm = new XElement(name ?? A + "xfrm");
        if (rot != 0)
        {
            xfrm.Add(new XAttribute("rot", S((long)Math.Round(rot * 60000))));
        }

        if (flipH)
        {
            xfrm.Add(new XAttribute("flipH", "1"));
        }

        if (flipV)
        {
            xfrm.Add(new XAttribute("flipV", "1"));
        }

        xfrm.Add(new XElement(A + "off", new XAttribute("x", S(Emu(x))), new XAttribute("y", S(Emu(y)))),
            new XElement(A + "ext", new XAttribute("cx", S(Math.Max(0, Emu(w)))), new XAttribute("cy", S(Math.Max(0, Emu(h))))));
        return xfrm;
    }

    /// <summary>The preset geometry of a shape with its one adjustment.</summary>
    internal static XElement Geometry(string shape, double? adj)
    {
        var av = new XElement(A + "avLst");
        if (adj is { } a)
        {
            var name = shape switch
            {
                "rightArrow" or "leftArrow" or "upArrow" or "downArrow" or "leftRightArrow" or "upDownArrow" => "adj2",
                "wedgeRectCallout" or "wedgeRoundRectCallout" or "wedgeEllipseCallout" => null,
                _ => "adj",
            };
            if (name is not null)
            {
                av.Add(new XElement(A + "gd", new XAttribute("name", name), new XAttribute("fmla", "val " + S((long)Math.Round(a * 100000)))));
            }
        }

        return new XElement(A + "prstGeom", new XAttribute("prst", shape), av);
    }

    internal static string AlignOut(string align) => align switch { "center" => "ctr", "right" => "r", "justify" => "just", _ => "l" };

    internal static string AnchorOut(string anchor) => anchor switch { "middle" => "ctr", "bottom" => "b", _ => "t" };

    /// <summary>A background (p:bg) of a fill; pictures become image parts of the part.</summary>
    internal static XElement? Background(JsonObject? fill, OpenXmlPart part)
    {
        var f = FillXml(fill, part, 1);
        return f is null ? null : new XElement(P + "bg", new XElement(P + "bgPr", f, new XElement(A + "effectLst")));
    }

    /// <summary>A fill as DrawingML (noFill, solidFill, gradFill, blipFill).</summary>
    internal static XElement? FillXml(JsonObject? fill, OpenXmlPart part, double alpha)
    {
        switch (fill?["type"]?.GetValue<string>())
        {
            case "none":
                return new XElement(A + "noFill");
            case "solid":
                return SolidFill(fill["color"]!.GetValue<string>(), alpha);
            case "gradient":
                var angle = fill["angle"]?.GetValue<double>() ?? 90;
                return new XElement(A + "gradFill", new XAttribute("rotWithShape", "1"),
                    new XElement(A + "gsLst",
                        new XElement(A + "gs", new XAttribute("pos", "0"), Color(fill["from"]!.GetValue<string>(), alpha)),
                        new XElement(A + "gs", new XAttribute("pos", "100000"), Color(fill["to"]!.GetValue<string>(), alpha))),
                    new XElement(A + "lin", new XAttribute("ang", S((long)Math.Round(((angle + 270) % 360) * 60000))), new XAttribute("scaled", "0")));
            case "image":
                var rel = AddImage(part, fill["src"]!.GetValue<string>());
                return rel is null ? null : new XElement(A + "blipFill", new XAttribute("dpi", "0"), new XAttribute("rotWithShape", "1"),
                    new XElement(A + "blip", new XAttribute(R + "embed", rel)), new XElement(A + "srcRect"), new XElement(A + "stretch", new XElement(A + "fillRect")));
            default:
                return null;
        }
    }

    /// <summary>A picture (data: address) as an image part of the part; answers the relationship id.</summary>
    internal static string? AddImage(OpenXmlPart part, string dataUrl)
    {
        var comma = dataUrl.IndexOf(',', StringComparison.Ordinal);
        if (comma < 0)
        {
            return null;
        }

        var mime = dataUrl[5..dataUrl.IndexOf(';', StringComparison.Ordinal)];
        var type = mime switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            _ => (PartTypeInfo?)null,
        };
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(dataUrl[(comma + 1)..]);
        }
        catch (System.FormatException)
        {
            return null;
        }

        ImagePart image;
        if (type is { } t)
        {
            image = part switch
            {
                SlidePart s => s.AddImagePart(t),
                SlideLayoutPart l => l.AddImagePart(t),
                SlideMasterPart m => m.AddImagePart(t),
                _ => throw new InvalidOperationException(),
            };
        }
        else
        {
            // WebP is not a type of the SDK: added as a part of its own content type
            image = part switch
            {
                SlidePart s => s.AddImagePart(new PartTypeInfo(mime, ".webp")),
                SlideLayoutPart l => l.AddImagePart(new PartTypeInfo(mime, ".webp")),
                SlideMasterPart m => m.AddImagePart(new PartTypeInfo(mime, ".webp")),
                _ => throw new InvalidOperationException(),
            };
        }

        using (var ms = new MemoryStream(bytes))
        {
            image.FeedData(ms);
        }

        return part.GetIdOfPart(image);
    }

    internal static XElement Picture(int id, string name, string? alt, string rel, JsonObject? crop, XElement xfrm, double opacity, XElement? line, bool shadow, XElement? ph = null)
    {
        var cNvPr = new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", name));
        if (!string.IsNullOrEmpty(alt))
        {
            cNvPr.Add(new XAttribute("descr", alt));
        }

        var blip = new XElement(A + "blip", new XAttribute(R + "embed", rel));
        if (opacity < 1)
        {
            blip.Add(new XElement(A + "alphaModFix", new XAttribute("amt", S((int)Math.Round(opacity * 100000)))));
        }

        var blipFill = new XElement(P + "blipFill", blip);
        if (crop is not null)
        {
            var src = new XElement(A + "srcRect");
            foreach (var (key, attr) in new[] { ("l", "l"), ("t", "t"), ("r", "r"), ("b", "b") })
            {
                var v = crop[key]?.GetValue<double>() ?? 0;
                if (v != 0)
                {
                    src.Add(new XAttribute(attr, S((int)Math.Round(v * 100000))));
                }
            }

            blipFill.Add(src);
        }

        blipFill.Add(new XElement(A + "stretch", new XElement(A + "fillRect")));
        var spPr = new XElement(P + "spPr", xfrm, new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")));
        if (line is not null)
        {
            spPr.Add(line);
        }

        if (shadow)
        {
            spPr.Add(new XElement(A + "effectLst", Shadow()));
        }

        return new XElement(P + "pic",
            new XElement(P + "nvPicPr", cNvPr, new XElement(P + "cNvPicPr", new XElement(A + "picLocks", new XAttribute("noChangeAspect", "1"))), new XElement(P + "nvPr", ph)),
            blipFill, spPr);
    }

    // ---- one slide -----------------------------------------------------------------------------------------------------------------

    private sealed class ShapeIds
    {
        private int _next = 1;

        public int Next() => ++_next;
    }

    /// <summary>Writes one slide: its elements, background, transition and animations.</summary>
    private sealed class SlideWriter(PptxWriter owner, SlidePart part, JsonObject slide, int number)
    {
        private readonly ShapeIds _ids = new();
        private readonly Dictionary<string, int> _spids = new(StringComparer.Ordinal);
        private readonly SlideLayout _layout = SlidesData.Layout(slide["layout"]?.GetValue<string>());

        private string Group => _layout.Group;

        public XDocument Xml()
        {
            var tree = SpTree();
            var elements = slide["elements"]!.AsArray().OfType<JsonObject>().ToList();
            for (var i = 0; i < elements.Count;)
            {
                var group = elements[i]["group"]?.GetValue<string>();
                if (group is null)
                {
                    AddElement(tree, elements[i]);
                    i++;
                    continue;
                }

                var members = new List<JsonObject>();
                while (i < elements.Count && elements[i]["group"]?.GetValue<string>() == group)
                {
                    members.Add(elements[i]);
                    i++;
                }

                if (members.Count < 2)
                {
                    members.ForEach(m => AddElement(tree, m));
                    continue;
                }

                tree.Add(GroupXml(members));
            }

            Footers(tree);

            var cSld = new XElement(P + "cSld");
            if (slide["background"] is JsonObject bg && Background(bg, part) is { } bgXml)
            {
                cSld.Add(bgXml);
            }

            cSld.Add(tree);
            var root = new XElement(P + "sld", Ns());
            if (slide["hidden"]?.GetValue<bool>() == true)
            {
                root.Add(new XAttribute("show", "0"));
            }

            root.Add(cSld, new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping")));
            if (TransitionXml() is { } transition)
            {
                root.Add(transition);
            }

            if (PptxAnimations.Timing(slide, _spids) is { } timing)
            {
                root.Add(timing);
            }

            return new XDocument(root);
        }

        private XElement GroupXml(List<JsonObject> members)
        {
            var minX = members.Min(m => m["x"]!.GetValue<double>());
            var minY = members.Min(m => m["y"]!.GetValue<double>());
            var maxX = members.Max(m => m["x"]!.GetValue<double>() + m["w"]!.GetValue<double>());
            var maxY = members.Max(m => m["y"]!.GetValue<double>() + m["h"]!.GetValue<double>());
            var id = _ids.Next();
            var xfrm = Xfrm(minX, minY, maxX - minX, maxY - minY, 0, false, false);
            xfrm.Add(new XElement(A + "chOff", new XAttribute("x", S(Emu(minX))), new XAttribute("y", S(Emu(minY)))),
                new XElement(A + "chExt", new XAttribute("cx", S(Emu(maxX - minX))), new XAttribute("cy", S(Emu(maxY - minY)))));
            var g = new XElement(P + "grpSp",
                new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", "Group " + id)), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")),
                new XElement(P + "grpSpPr", xfrm));
            foreach (var m in members)
            {
                AddElement(g, m);
            }

            return g;
        }

        private void AddElement(XElement tree, JsonObject e)
        {
            var id = _ids.Next();
            _spids[e["id"]!.GetValue<string>()] = id;
            var type = e["type"]!.GetValue<string>();
            var x = e["x"]!.GetValue<double>();
            var y = e["y"]!.GetValue<double>();
            var w = e["w"]!.GetValue<double>();
            var h = e["h"]!.GetValue<double>();
            var rot = e["rot"]?.GetValue<double>() ?? 0;
            var flipH = e["flipH"]?.GetValue<bool>() ?? false;
            var flipV = e["flipV"]?.GetValue<bool>() ?? false;
            var xfrm = Xfrm(x, y, w, h, rot, flipH, flipV);
            var name = e["name"]?.GetValue<string>();
            switch (type)
            {
                case "image":
                {
                    var rel = AddImage(part, e["src"]!.GetValue<string>());
                    if (rel is null)
                    {
                        return;
                    }

                    var slot = _layout.Slot(e["ph"]?.GetValue<string>());
                    var ph = slot?.Role == "picture" ? new XElement(P + "ph", new XAttribute("type", "pic"), new XAttribute("idx", S(slot.Idx))) : null;
                    tree.Add(Picture(id, name ?? "Picture " + id, e["alt"]?.GetValue<string>(), rel, e["crop"] as JsonObject, xfrm, e["opacity"]?.GetValue<double>() ?? 1,
                        LineXml(e["line"] as JsonObject, null, null), e["shadow"]?.GetValue<bool>() == true, ph));
                    return;
                }

                case "line":
                {
                    var spPr = new XElement(P + "spPr", xfrm, Geometry("line", null),
                        LineXml(e["line"] as JsonObject, e["arrowStart"]?.GetValue<string>(), e["arrowEnd"]?.GetValue<string>()));
                    if (e["shadow"]?.GetValue<bool>() == true)
                    {
                        spPr.Add(new XElement(A + "effectLst", Shadow()));
                    }

                    tree.Add(new XElement(P + "cxnSp",
                        new XElement(P + "nvCxnSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", name ?? "Connector " + id)), new XElement(P + "cNvCxnSpPr"), new XElement(P + "nvPr")),
                        spPr));
                    return;
                }

                case "table":
                    tree.Add(TableXml(e, id, name));
                    return;
                default:
                    tree.Add(ShapeXml(e, id, name, xfrm));
                    return;
            }
        }

        private XElement ShapeXml(JsonObject e, int id, string? name, XElement xfrm)
        {
            var type = e["type"]!.GetValue<string>();
            var ph = e["ph"]?.GetValue<string>();
            var slot = _layout.Slot(ph);
            var nvPr = new XElement(P + "nvPr");
            var cNvSpPr = new XElement(P + "cNvSpPr");
            if (slot is not null && (slot.Role != "picture" || e["text"] is not null))
            {
                var phXml = new XElement(P + "ph");
                if (slot.Pptx != "obj")
                {
                    phXml.Add(new XAttribute("type", slot.Pptx));
                }

                if (slot.Idx > 0)
                {
                    phXml.Add(new XAttribute("idx", S(slot.Idx)));
                }

                nvPr.Add(phXml);
                cNvSpPr.Add(new XElement(A + "spLocks", new XAttribute("noGrp", "1")));
            }
            else if (type == "text")
            {
                cNvSpPr.Add(new XAttribute("txBox", "1"));
            }

            var shape = type == "shape" ? e["shape"]?.GetValue<string>() ?? "rect" : "rect";
            var spPr = new XElement(P + "spPr", xfrm, Geometry(shape, e["adj"]?.GetValue<double>()));
            var opacity = e["opacity"]?.GetValue<double>() ?? 1;
            if (FillXml(e["fill"] as JsonObject, part, opacity) is { } fill)
            {
                spPr.Add(fill);
            }
            else if (slot is null)
            {
                spPr.Add(new XElement(A + "noFill"));
            }

            if (LineXml(e["line"] as JsonObject, null, null) is { } line)
            {
                spPr.Add(line);
            }
            else if (slot is null && type == "shape")
            {
                spPr.Add(new XElement(A + "ln", new XElement(A + "noFill")));
            }

            if (e["shadow"]?.GetValue<bool>() == true)
            {
                spPr.Add(new XElement(A + "effectLst", Shadow()));
            }

            var sp = new XElement(P + "sp",
                new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", name ?? (slot is not null ? ph + " " + id : type == "text" ? "TextBox " + id : "Shape " + id))), cNvSpPr, nvPr),
                spPr);
            var role = slot?.Role ?? (type == "text" ? "text" : "shape");
            if (e["text"] is JsonObject body)
            {
                sp.Add(TextBodyXml(body, P + "txBody", role, slot, null));
            }
            else if (type == "shape" && slot is null)
            {
                sp.Add(new XElement(P + "txBody", new XElement(A + "bodyPr", new XAttribute("rtlCol", "0"), new XAttribute("anchor", "ctr")), new XElement(A + "lstStyle"),
                    new XElement(A + "p", new XElement(A + "pPr", new XAttribute("algn", "ctr")), new XElement(A + "endParaRPr", new XAttribute("lang", "en-US")))));
            }

            return sp;
        }

        /// <summary>A text body (p:txBody or a:txBody of a cell). Placeholders inherit from the layout; text boxes, shapes and cells say everything.</summary>
        private XElement TextBodyXml(JsonObject body, XName name, string role, LayoutSlot? slot, (string Text, bool Bold)? cell)
        {
            var anchor = body["anchor"]?.GetValue<string>() ?? (role is "shape" ? "middle" : slot?.Anchor ?? (cell is not null ? "middle" : "top"));
            var bodyPr = new XElement(A + "bodyPr");
            var wrap = body["wrap"]?.GetValue<bool>() ?? true;
            if (!wrap)
            {
                bodyPr.Add(new XAttribute("wrap", "none"));
            }
            else if (slot is null)
            {
                bodyPr.Add(new XAttribute("wrap", "square"));
            }

            if (body["inset"] is JsonArray inset)
            {
                bodyPr.Add(new XAttribute("lIns", S(Emu(inset[0]!.GetValue<double>()))), new XAttribute("tIns", S(Emu(inset[1]!.GetValue<double>()))),
                    new XAttribute("rIns", S(Emu(inset[2]!.GetValue<double>()))), new XAttribute("bIns", S(Emu(inset[3]!.GetValue<double>()))));
            }

            bodyPr.Add(new XAttribute("rtlCol", "0"), new XAttribute("anchor", AnchorOut(anchor)));
            var autofit = body["autofit"]?.GetValue<string>() ?? (slot is not null ? "shrink" : role == "text" ? "grow" : "none");
            if (cell is null)
            {
                bodyPr.Add(autofit switch { "shrink" => new XElement(A + "normAutofit"), "grow" => new XElement(A + "spAutoFit"), _ => new XElement(A + "noAutofit") });
            }

            var tx = new XElement(name, bodyPr, new XElement(A + "lstStyle"));
            var defaultColor = slot is not null ? null : cell is { } c ? c.Text : role == "shape" ? "lt1" : owner.GroupColor(Group, "text");
            var defaultAlign = slot is not null ? null : role == "shape" ? "center" : null;
            foreach (var p in body["paras"]!.AsArray().OfType<JsonObject>())
            {
                tx.Add(ParaXml(p, slot, defaultColor, defaultAlign, cell?.Bold ?? false));
            }

            return tx;
        }

        private XElement ParaXml(JsonObject p, LayoutSlot? slot, string? defaultColor, string? defaultAlign, bool bold)
        {
            var level = p["level"]?.GetValue<int>() ?? 0;
            var pPr = new XElement(A + "pPr");
            var bullet = p["bullet"]?.GetValue<string>();
            var (indent, step) = SlidesData.BodyIndent;
            if (slot is null || slot.Role != "body" || !slot.Bullet)
            {
                // text boxes, shapes and placeholders without bullets: the indent follows the bullet and the level
                if (bullet is not null && bullet != "none")
                {
                    pPr.Add(new XAttribute("marL", S(Emu(indent + (step * level)))), new XAttribute("indent", S(-Emu(indent))));
                }
                else if (level > 0)
                {
                    pPr.Add(new XAttribute("marL", S(Emu(step * level))), new XAttribute("indent", "0"));
                }
            }

            if (level > 0)
            {
                pPr.Add(new XAttribute("lvl", S(level)));
            }

            var align = p["align"]?.GetValue<string>() ?? defaultAlign;
            if (align is not null)
            {
                pPr.Add(new XAttribute("algn", AlignOut(align)));
            }

            if (p["lineSpacing"] is JsonValue ls)
            {
                pPr.Add(new XElement(A + "lnSpc", new XElement(A + "spcPct", new XAttribute("val", S((int)Math.Round(ls.GetValue<double>() * 100000))))));
            }

            if (p["spaceBefore"] is JsonValue sb)
            {
                pPr.Add(new XElement(A + "spcBef", new XElement(A + "spcPts", new XAttribute("val", S((int)Math.Round(sb.GetValue<double>() * 100))))));
            }

            if (p["spaceAfter"] is JsonValue sa)
            {
                pPr.Add(new XElement(A + "spcAft", new XElement(A + "spcPts", new XAttribute("val", S((int)Math.Round(sa.GetValue<double>() * 100))))));
            }

            if (bullet == "none")
            {
                pPr.Add(new XElement(A + "buNone"));
            }
            else if (bullet is not null && NumberingOut(bullet) is { } scheme)
            {
                pPr.Add(new XElement(A + "buFont", new XAttribute("typeface", "+mj-lt")), new XElement(A + "buAutoNum", new XAttribute("type", scheme)));
            }
            else if (bullet is not null)
            {
                pPr.Add(new XElement(A + "buFont", new XAttribute("typeface", "Arial"), new XAttribute("panose", "020B0604020202020204"), new XAttribute("pitchFamily", "34"), new XAttribute("charset", "0")),
                    new XElement(A + "buChar", new XAttribute("char", bullet)));
            }
            else if (slot is null || slot.Role != "body" || !slot.Bullet)
            {
                pPr.Add(new XElement(A + "buNone"));
            }

            var para = new XElement(A + "p");
            if (pPr.HasAttributes || pPr.HasElements)
            {
                para.Add(pPr);
            }

            JsonObject? last = null;
            foreach (var run in p["runs"]!.AsArray().OfType<JsonObject>())
            {
                last = run;
                var text = run["t"]?.GetValue<string>() ?? string.Empty;
                var parts = text.Split('\n');
                for (var i = 0; i < parts.Length; i++)
                {
                    if (i > 0)
                    {
                        para.Add(new XElement(A + "br", RunProps(run, A + "rPr", defaultColor, bold)));
                    }

                    if (parts[i].Length > 0)
                    {
                        para.Add(new XElement(A + "r", RunProps(run, A + "rPr", defaultColor, bold), new XElement(A + "t", parts[i])));
                    }
                }
            }

            para.Add(RunProps(last, A + "endParaRPr", defaultColor, bold));
            return para;
        }

        private XElement RunProps(JsonObject? run, XName name, string? defaultColor, bool bold)
        {
            var r = new XElement(name, new XAttribute("lang", "en-US"));
            if (run?["size"] is JsonValue size)
            {
                r.Add(new XAttribute("sz", S((int)Math.Round(size.GetValue<double>() * 100))));
            }

            var b = run?["b"]?.GetValue<bool>();
            if (b == true || (b is null && bold))
            {
                r.Add(new XAttribute("b", "1"));
            }
            else if (b == false)
            {
                r.Add(new XAttribute("b", "0"));
            }

            if (run?["i"]?.GetValue<bool>() == true)
            {
                r.Add(new XAttribute("i", "1"));
            }

            if (run?["u"]?.GetValue<bool>() == true)
            {
                r.Add(new XAttribute("u", "sng"));
            }

            if (run?["s"]?.GetValue<bool>() == true)
            {
                r.Add(new XAttribute("strike", "sngStrike"));
            }

            r.Add(new XAttribute("dirty", "0"));
            var color = run?["color"]?.GetValue<string>() ?? defaultColor;
            if (color is not null)
            {
                r.Add(SolidFill(color));
            }

            if (run?["hl"]?.GetValue<string>() is { } hl)
            {
                r.Add(new XElement(A + "highlight", Color(hl)));
            }

            if (run?["font"]?.GetValue<string>() is { } font)
            {
                r.Add(new XElement(A + "latin", new XAttribute("typeface", font)), new XElement(A + "cs", new XAttribute("typeface", font)));
            }

            return r;
        }

        private static string? NumberingOut(string bullet) => bullet switch
        {
            "1." => "arabicPeriod",
            "1)" => "arabicParenR",
            "a." => "alphaLcPeriod",
            "a)" => "alphaLcParenR",
            "A." => "alphaUcPeriod",
            "A)" => "alphaUcParenR",
            "I." => "romanUcPeriod",
            "i." => "romanLcPeriod",
            _ => null,
        };

        private static XElement? LineXml(JsonObject? line, string? arrowStart, string? arrowEnd)
        {
            if (line is null)
            {
                return null;
            }

            var ln = new XElement(A + "ln", new XAttribute("w", S((long)Math.Round(line["width"]!.GetValue<double>() * 12700))), SolidFill(line["color"]!.GetValue<string>()));
            var dash = line["dash"]?.GetValue<string>() switch { "dash" => "dash", "dot" => "sysDot", "dashDot" => "dashDot", "longDash" => "lgDash", _ => "solid" };
            ln.Add(new XElement(A + "prstDash", new XAttribute("val", dash)));
            if (arrowStart is not null && arrowStart != "none")
            {
                ln.Add(new XElement(A + "headEnd", new XAttribute("type", arrowStart)));
            }

            if (arrowEnd is not null && arrowEnd != "none")
            {
                ln.Add(new XElement(A + "tailEnd", new XAttribute("type", arrowEnd)));
            }

            return ln;
        }

        private XElement TableXml(JsonObject e, int id, string? name)
        {
            var cols = e["cols"]!.AsArray().Select(c => c!.GetValue<double>()).ToList();
            var rows = e["rows"]!.AsArray().OfType<JsonObject>().ToList();
            var tblPr = new XElement(A + "tblPr");
            if (e["header"]?.GetValue<bool>() == true)
            {
                tblPr.Add(new XAttribute("firstRow", "1"));
            }

            if (e["firstCol"]?.GetValue<bool>() == true)
            {
                tblPr.Add(new XAttribute("firstCol", "1"));
            }

            if (e["lastRow"]?.GetValue<bool>() == true)
            {
                tblPr.Add(new XAttribute("lastRow", "1"));
            }

            if (e["banded"]?.GetValue<bool>() == true)
            {
                tblPr.Add(new XAttribute("bandRow", "1"));
            }

            var tbl = new XElement(A + "tbl", tblPr, new XElement(A + "tblGrid", cols.Select(c => new XElement(A + "gridCol", new XAttribute("w", S(Emu(c)))))));
            for (var r = 0; r < rows.Count; r++)
            {
                var tr = new XElement(A + "tr", new XAttribute("h", S(Emu(rows[r]["h"]!.GetValue<double>()))));
                var cells = rows[r]["cells"]!.AsArray().OfType<JsonObject>().ToList();
                for (var c = 0; c < cols.Count; c++)
                {
                    var cell = c < cells.Count ? cells[c] : new JsonObject();
                    var look = TableLook(e, r, c);
                    var body = cell["text"] as JsonObject ?? new JsonObject { ["paras"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["t"] = "" }) }) };
                    var txBody = TextBodyXml(body, A + "txBody", "cell", null, (look.Text, look.Bold));
                    var anchor = body["anchor"]?.GetValue<string>() ?? "middle";
                    var tcPr = new XElement(A + "tcPr", new XAttribute("marL", "91440"), new XAttribute("marR", "91440"), new XAttribute("marT", "45720"), new XAttribute("marB", "45720"), new XAttribute("anchor", AnchorOut(anchor)));
                    foreach (var side in new[] { "lnL", "lnR", "lnT", "lnB" })
                    {
                        tcPr.Add(look.Border is { } border
                            ? new XElement(A + side, new XAttribute("w", "12700"), SolidFill(border))
                            : new XElement(A + side, new XAttribute("w", "12700"), new XElement(A + "noFill")));
                    }

                    var fill = cell["fill"] is JsonObject own ? FillXml(own, part, 1) : look.Fill is { } lf ? SolidFill(lf) : new XElement(A + "noFill");
                    if (fill is not null)
                    {
                        tcPr.Add(fill);
                    }

                    tr.Add(new XElement(A + "tc", txBody, tcPr));
                }

                tbl.Add(tr);
            }

            var x = e["x"]!.GetValue<double>();
            var y = e["y"]!.GetValue<double>();
            return new XElement(P + "graphicFrame",
                new XElement(P + "nvGraphicFramePr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", name ?? "Table " + id)),
                    new XElement(P + "cNvGraphicFramePr", new XElement(A + "graphicFrameLocks", new XAttribute("noGrp", "1"))), new XElement(P + "nvPr")),
                Xfrm(x, y, cols.Sum(), rows.Sum(rw => rw["h"]!.GetValue<double>()), 0, false, false, P + "xfrm"),
                new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table"), tbl)));
        }

        /// <summary>The slide number, date and footer text of the presentation (not on title slides when "skipTitle").</summary>
        private void Footers(XElement tree)
        {
            var footer = owner._content["footer"] as JsonObject;
            if (footer is null || (_layout.Id == "title" && footer["skipTitle"]?.GetValue<bool>() != false))
            {
                return;
            }

            if (footer["date"]?.GetValue<bool>() == true)
            {
                var text = footer["dateText"]?.GetValue<string>();
                var p = string.IsNullOrEmpty(text)
                    ? new XElement(A + "p", new XElement(A + "fld", new XAttribute("id", "{1C9A2A35-6A1F-4C3B-9C0D-3B7B7C8F6A10}"), new XAttribute("type", "datetime1"),
                        new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", DateTime.Today.ToString("d", CultureInfo.InvariantCulture))))
                    : new XElement(A + "p", new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", text)));
                tree.Add(FooterShape("dt", 10, "Date Placeholder", SlidesData.FooterDate, p));
            }

            if (footer["text"]?.GetValue<string>() is { Length: > 0 } ftr)
            {
                tree.Add(FooterShape("ftr", 11, "Footer Placeholder", SlidesData.FooterText,
                    new XElement(A + "p", new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", ftr)))));
            }

            if (footer["number"]?.GetValue<bool>() == true)
            {
                tree.Add(FooterShape("sldNum", 12, "Slide Number Placeholder", SlidesData.FooterNumber,
                    new XElement(A + "p", new XElement(A + "fld", new XAttribute("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}"), new XAttribute("type", "slidenum"),
                        new XElement(A + "rPr", new XAttribute("lang", "en-US")), new XElement(A + "t", S(number))))));
            }
        }

        private XElement FooterShape(string type, int idx, string name, FooterSlot slot, XElement p)
        {
            var id = _ids.Next();
            return new XElement(P + "sp",
                new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", S(id)), new XAttribute("name", $"{name} {id}")),
                    new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                    new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", type), new XAttribute("sz", type == "ftr" ? "quarter" : "half"), new XAttribute("idx", S(idx))))),
                new XElement(P + "spPr", Xfrm(slot.X * owner._w, slot.Y * owner._h, slot.W * owner._w, slot.H * owner._h, 0, false, false)),
                new XElement(P + "txBody", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"), p));
        }

        private XElement? TransitionXml()
        {
            if (slide["transition"] is not JsonObject t)
            {
                return null;
            }

            var type = t["type"]?.GetValue<string>() ?? "none";
            var duration = t["duration"]?.GetValue<double>() ?? 0.7;
            var dir = t["dir"]?.GetValue<string>();
            var tr = new XElement(P + "transition", new XAttribute("spd", duration <= 0.5 ? "fast" : duration < 0.9 ? "med" : "slow"));
            if (t["auto"] is JsonValue auto)
            {
                tr.Add(new XAttribute("advTm", S((long)Math.Round(auto.GetValue<double>() * 1000))));
            }

            // dir says where the new slide comes FROM (up = the top, down = the bottom); PowerPoint names the direction it moves TO
            var side = (dir ?? (type == "push" ? "down" : "right")) switch { "left" => "r", "up" => "d", "down" => "u", _ => "l" };
            switch (type)
            {
                case "fade":
                    tr.Add(new XElement(P + "fade"));
                    break;
                case "push":
                    tr.Add(new XElement(P + "push", new XAttribute("dir", side)));
                    break;
                case "wipe":
                    tr.Add(new XElement(P + "wipe", new XAttribute("dir", side)));
                    break;
                case "cover":
                    tr.Add(new XElement(P + "cover", new XAttribute("dir", side)));
                    break;
                case "split":
                    tr.Add(new XElement(P + "split", new XAttribute("orient", dir is "up" or "down" ? "horz" : "vert"), new XAttribute("dir", "out")));
                    break;
                case "zoom":
                    tr.Add(new XElement(P + "zoom", new XAttribute("dir", "in")));
                    break;
            }

            return tr;
        }
    }
}
