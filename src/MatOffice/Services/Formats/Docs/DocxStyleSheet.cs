using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The styles of a Word file as the import needs them: the document defaults, paragraph, character and table styles resolved along their
/// <c>basedOn</c> chains, the fonts of the theme, and the properties of paragraphs and runs merged level by level the way Word does
/// (defaults › table style › paragraph style › character style › direct formatting; a paragraph of the default style in a table takes the table
/// style over its own style, like Word). Styles are mapped to Docs by their names, which are English in every Word ("heading 1", "Title" …).
/// </summary>
internal sealed class StyleSheet
{
    private const int MaxChain = 20;

    private readonly Dictionary<string, Style> _byId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Resolved> _paragraphs = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Resolved?> _characters = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Resolved?> _tables = new(StringComparer.Ordinal);
    private readonly Props _defaults = new();
    private readonly string? _defaultParagraph;
    private readonly string? _defaultTable;
    private readonly string _majorFont;
    private readonly string _minorFont;

    public StyleSheet(Styles? styles, string? majorFont, string? minorFont)
    {
        _majorFont = string.IsNullOrWhiteSpace(majorFont) ? "Calibri Light" : majorFont.Trim();
        _minorFont = string.IsNullOrWhiteSpace(minorFont) ? "Calibri" : minorFont.Trim();
        if (styles is null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            if (style.StyleId?.Value is not { Length: > 0 } id)
            {
                continue;
            }

            _byId.TryAdd(id, style);
            if (style.Default?.Value == true)
            {
                switch (TypeOf(style))
                {
                    case "paragraph":
                        _defaultParagraph ??= id;
                        break;
                    case "table":
                        _defaultTable ??= id;
                        break;
                }
            }
        }

        var defaults = styles.DocDefaults;
        if (defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle is { } pPr)
        {
            _defaults.MergeParagraph(pPr);
        }

        if (defaults?.RunPropertiesDefault?.RunPropertiesBaseStyle is { } rPr)
        {
            _defaults.MergeRun(rPr, this);
            Language = LanguageOf(rPr);
        }

        if (Language is null && _defaultParagraph is not null && _byId.TryGetValue(_defaultParagraph, out var normal) && normal.StyleRunProperties is { } normalRun)
        {
            Language = LanguageOf(normalRun);
        }
    }

    /// <summary>The language of the text (<c>de-DE</c>, <c>en</c> …) from the document defaults, or null.</summary>
    public string? Language { get; }

    /// <summary>The paragraph style of an id (or the default style), with its Docs key and its merged properties.</summary>
    public Resolved Paragraph(string? styleId, Resolved? tableStyle)
    {
        var id = styleId is not null && _byId.TryGetValue(styleId, out var style) && TypeOf(style) == "paragraph" ? styleId : _defaultParagraph;
        var cacheKey = (id ?? string.Empty) + "\u0001" + (tableStyle?.Id ?? string.Empty);
        if (_paragraphs.TryGetValue(cacheKey, out var cached))
        {
            return cached;
        }

        var chain = Chain(id, "paragraph");
        var props = _defaults.Clone();
        var isDefault = id is null || id == _defaultParagraph;
        if (tableStyle is not null && !isDefault)
        {
            props.Merge(tableStyle.Props);
        }

        foreach (var s in chain)
        {
            MergeStyle(props, s);
        }

        if (tableStyle is not null && isDefault)
        {
            props.Merge(tableStyle.Props);
        }

        var resolved = new Resolved(id ?? string.Empty, chain.Count > 0 ? NameOf(chain[^1]) : "Normal", KeyOf(chain, props), props);
        _paragraphs[cacheKey] = resolved;
        return resolved;
    }

    /// <summary>A character style with its merged run properties, or null when there is none of that id.</summary>
    public Resolved? Character(string styleId)
    {
        if (_characters.TryGetValue(styleId, out var cached))
        {
            return cached;
        }

        Resolved? result = null;
        if (_byId.TryGetValue(styleId, out var style) && TypeOf(style) == "character")
        {
            var props = new Props();
            foreach (var s in Chain(styleId, "character"))
            {
                if (s.StyleRunProperties is { } rPr)
                {
                    props.MergeRun(rPr, this);
                }
            }

            result = new Resolved(styleId, NameOf(style), "Normal", props);
        }

        _characters[styleId] = result;
        return result;
    }

    /// <summary>A table style (or the default one) with the paragraph/run properties it gives its cells and its borders, or null.</summary>
    public Resolved? Table(string? styleId)
    {
        var id = styleId is not null && _byId.TryGetValue(styleId, out var style) && TypeOf(style) == "table" ? styleId : _defaultTable;
        if (id is null)
        {
            return null;
        }

        if (_tables.TryGetValue(id, out var cached))
        {
            return cached;
        }

        var props = new Props();
        var borders = new Dictionary<string, Border?>(StringComparer.Ordinal);
        var chain = Chain(id, "table");
        foreach (var s in chain)
        {
            MergeStyle(props, s);
            if (s.StyleTableProperties?.GetFirstChild<TableBorders>() is { } tableBorders)
            {
                foreach (var (side, border) in ReadBorders(tableBorders))
                {
                    borders[side] = border;
                }
            }
        }

        var result = chain.Count == 0 ? null : new Resolved(id, NameOf(chain[^1]), "Normal", props) { TableBorders = borders };
        _tables[id] = result;
        return result;
    }

    /// <summary>The numbering instance a numbering style (<c>w:numStyleLink</c>) stands for, or null.</summary>
    public int? NumIdOfStyle(string styleId)
        => _byId.TryGetValue(styleId, out var style) && TypeOf(style) == "numbering"
            ? style.StyleParagraphProperties?.GetFirstChild<NumberingProperties>()?.NumberingId?.Val?.Value
            : null;

    /// <summary>The font of <c>w:rFonts</c>: the theme font when one is named (it wins over the explicit name), else the Latin font.</summary>
    public string? FontOf(RunFonts fonts)
    {
        var theme = fonts.AsciiTheme?.InnerText ?? fonts.HighAnsiTheme?.InnerText;
        if (!string.IsNullOrEmpty(theme))
        {
            return theme.StartsWith("major", StringComparison.Ordinal) ? _majorFont : _minorFont;
        }

        var name = fonts.Ascii?.Value ?? fonts.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
    }

    /// <summary>An on/off property: present without <c>w:val</c> is on; <c>0</c>, <c>false</c>, <c>off</c> are off.</summary>
    public static bool OnOff(OpenXmlElement? element)
    {
        if (element is null)
        {
            return false;
        }

        foreach (var attribute in element.GetAttributes())
        {
            if (attribute.LocalName == "val")
            {
                return attribute.Value?.Trim().ToLowerInvariant() is not ("0" or "false" or "off" or "none");
            }
        }

        return true;
    }

    /// <summary>The sides of a border element (<c>w:pBdr</c>, <c>w:tblBorders</c>, <c>w:tcBorders</c>): top, right, bottom, left, insideH, insideV.</summary>
    public static IEnumerable<(string Side, Border? Border)> ReadBorders(OpenXmlElement borders)
    {
        foreach (var child in borders.ChildElements)
        {
            var side = child switch
            {
                TopBorder => "top",
                BottomBorder => "bottom",
                LeftBorder or StartBorder => "left",
                RightBorder or EndBorder => "right",
                InsideHorizontalBorder => "insideH",
                InsideVerticalBorder => "insideV",
                _ => null,
            };
            if (side is not null && child is BorderType border)
            {
                yield return (side, ReadBorder(border));
            }
        }
    }

    /// <summary>The colour of a shading as <c>RRGGBB</c> (no <c>#</c>), or null for none/automatic.</summary>
    public static string? ShadingColour(Shading shading)
    {
        var pattern = shading.Val?.InnerText;
        if (pattern is "nil")
        {
            return null;
        }

        var fill = shading.Fill?.Value;
        if (pattern is "solid")
        {
            var colour = shading.Color?.Value;
            return IsHex(colour) ? colour : colour is null or "auto" ? (IsHex(fill) ? fill : "000000") : null;
        }

        return IsHex(fill) ? fill : null;
    }

    private static Border ReadBorder(BorderType border)
    {
        var value = border.Val?.InnerText ?? "single";
        if (value is "nil" or "none")
        {
            return new Border(null, 0);
        }

        var width = Math.Clamp((border.Size?.Value ?? 4) / 8.0, 0.25, 6);
        var style = value switch
        {
            "double" or "triple" or "thinThickSmallGap" or "thickThinSmallGap" or "thinThickThinSmallGap" or "thinThickMediumGap" or "thickThinMediumGap"
                or "thinThickThinMediumGap" or "thinThickLargeGap" or "thickThinLargeGap" or "thinThickThinLargeGap" or "doubleWave" => "double",
            "dotted" or "dotDotDash" => "dotted",
            "dashed" or "dashSmallGap" or "dotDash" or "dashDotStroked" => "dashed",
            _ => "solid",
        };
        if (style == "double")
        {
            width = Math.Max(width, 1.5);
        }

        var colour = border.Color?.Value is { } c && IsHex(c) ? "#" + c.ToLowerInvariant() : "#000000";
        return new Border($"{CssList.Pt(width)} {style} {colour}", Math.Clamp((double)(border.Space?.Value ?? 0), 0, 31));
    }

    private static bool IsHex(string? value) => value is { Length: 6 } && value.All(char.IsAsciiHexDigit);

    private static string? LanguageOf(OpenXmlElement rPr)
    {
        var value = rPr.GetFirstChild<Languages>()?.Val?.Value?.Trim();
        if (string.IsNullOrEmpty(value) || value.Length < 2 || !char.IsAsciiLetter(value[0]) || !char.IsAsciiLetter(value[1]))
        {
            return null;
        }

        var parts = value.Split('-', '_');
        var language = parts[0].ToLowerInvariant();
        if (language.Length != 2)
        {
            return null;
        }

        return parts.Length > 1 && parts[1].Length == 2 && parts[1].All(char.IsAsciiLetter) ? language + "-" + parts[1].ToUpperInvariant() : language;
    }

    private static string TypeOf(Style style) => style.Type?.InnerText ?? "paragraph";

    private static string NameOf(Style style) => style.StyleName?.Val?.Value ?? style.StyleId?.Value ?? string.Empty;

    /// <summary>The styles from the root of the basedOn chain to the style itself (only styles of the same type, no loops).</summary>
    private List<Style> Chain(string? id, string type)
    {
        var chain = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (id is not null && chain.Count < MaxChain && seen.Add(id) && _byId.TryGetValue(id, out var style) && TypeOf(style) == type)
        {
            chain.Add(style);
            id = style.BasedOn?.Val?.Value;
        }

        chain.Reverse();
        return chain;
    }

    private void MergeStyle(Props props, Style style)
    {
        if (style.StyleParagraphProperties is { } pPr)
        {
            props.MergeParagraph(pPr);
            if (pPr.GetFirstChild<NumberingProperties>() is { NumberingId: not null } numPr && numPr.NumberingLevelReference is null)
            {
                props.NumStyleId = style.StyleId?.Value;
            }
        }

        if (style.StyleRunProperties is { } rPr)
        {
            props.MergeRun(rPr, this);
        }
    }

    /// <summary>The Docs style of a chain: the nearest style whose name Docs knows; else a heading by its outline level; else Normal.</summary>
    private static string KeyOf(List<Style> chain, Props props)
    {
        for (var i = chain.Count - 1; i >= 0; i--)
        {
            if (DocsStyles.FromWordName(NameOf(chain[i])) is { } key)
            {
                return key;
            }
        }

        return props.OutlineLevel is >= 0 and <= 5 ? "Heading" + (props.OutlineLevel.Value + 1).ToString(CultureInfo.InvariantCulture) : "Normal";
    }

    /// <summary>A border side: its CSS (<c>0.5pt solid #000000</c>, null = explicitly none) and the space to the text in points.</summary>
    internal sealed record Border(string? Css, double Space);

    /// <summary>A resolved style.</summary>
    internal sealed class Resolved(string id, string name, string key, Props props)
    {
        public string Id { get; } = id;

        /// <summary>The name in Word ("heading 1").</summary>
        public string Name { get; } = name;

        /// <summary>The Docs style key (Normal, Heading1, Title …).</summary>
        public string Key { get; } = key;

        public Props Props { get; } = props;

        /// <summary>The borders of a table style (top, right, bottom, left, insideH, insideV).</summary>
        public Dictionary<string, Border?> TableBorders { get; init; } = new(StringComparer.Ordinal);
    }

    /// <summary>The properties of a paragraph and of a run; null = not said at this level (inherited).</summary>
    internal sealed class Props
    {
        public string? Align { get; set; }

        /// <summary>Space before/after in points.</summary>
        public double? Before { get; set; }

        public double? After { get; set; }

        /// <summary>The line spacing as Word writes it (240ths of a line for <c>auto</c>, twips for <c>exact</c>/<c>atLeast</c>).</summary>
        public double? Line { get; set; }

        public string? LineRule { get; set; }

        /// <summary>Indents in points.</summary>
        public double? Left { get; set; }

        public double? Right { get; set; }

        public double? FirstLine { get; set; }

        public double? Hanging { get; set; }

        /// <summary>The shading of the paragraph (<c>RRGGBB</c>; empty = explicitly none).</summary>
        public string? ParagraphShading { get; set; }

        public Dictionary<string, Border?> Borders { get; private set; } = new(StringComparer.Ordinal);

        public int? NumId { get; set; }

        public int? Ilvl { get; set; }

        /// <summary>The paragraph style that brought a numbering without a level (the level comes from the numbering definition).</summary>
        public string? NumStyleId { get; set; }

        public bool? PageBreakBefore { get; set; }

        public int? OutlineLevel { get; set; }

        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public bool? Underline { get; set; }

        public bool? Strike { get; set; }

        public bool? Vanish { get; set; }

        /// <summary>1 = superscript, -1 = subscript, 0 = baseline.</summary>
        public int? Script { get; set; }

        public string? Font { get; set; }

        /// <summary>The font size in points.</summary>
        public double? Size { get; set; }

        /// <summary><c>RRGGBB</c> or <c>auto</c>.</summary>
        public string? Color { get; set; }

        /// <summary>A Word highlight colour name (<c>yellow</c>; empty = explicitly none).</summary>
        public string? Highlight { get; set; }

        /// <summary>The shading of the run (<c>RRGGBB</c>; empty = explicitly none).</summary>
        public string? Shading { get; set; }

        /// <summary>The letter spacing in points.</summary>
        public double? Spacing { get; set; }

        public Props Clone()
        {
            var copy = (Props)MemberwiseClone();
            copy.Borders = new Dictionary<string, Border?>(Borders, StringComparer.Ordinal);
            return copy;
        }

        /// <summary>Takes every property the other one says.</summary>
        public void Merge(Props other)
        {
            Align = other.Align ?? Align;
            Before = other.Before ?? Before;
            After = other.After ?? After;
            if (other.Line is not null)
            {
                Line = other.Line;
                LineRule = other.LineRule;
            }

            Left = other.Left ?? Left;
            Right = other.Right ?? Right;
            if (other.Hanging is not null || other.FirstLine is not null)
            {
                Hanging = other.Hanging;
                FirstLine = other.FirstLine;
            }

            ParagraphShading = other.ParagraphShading ?? ParagraphShading;
            foreach (var (side, border) in other.Borders)
            {
                Borders[side] = border;
            }

            NumId = other.NumId ?? NumId;
            Ilvl = other.Ilvl ?? Ilvl;
            NumStyleId = other.NumStyleId ?? NumStyleId;
            PageBreakBefore = other.PageBreakBefore ?? PageBreakBefore;
            OutlineLevel = other.OutlineLevel ?? OutlineLevel;
            MergeCharacterStyle(other);
        }

        /// <summary>Takes the run properties the other one says (a character style).</summary>
        public void MergeCharacterStyle(Props other)
        {
            Bold = other.Bold ?? Bold;
            Italic = other.Italic ?? Italic;
            Underline = other.Underline ?? Underline;
            Strike = other.Strike ?? Strike;
            Vanish = other.Vanish ?? Vanish;
            Script = other.Script ?? Script;
            Font = other.Font ?? Font;
            Size = other.Size ?? Size;
            Color = other.Color ?? Color;
            Highlight = other.Highlight ?? Highlight;
            Shading = other.Shading ?? Shading;
            Spacing = other.Spacing ?? Spacing;
        }

        /// <summary>Merges paragraph properties (<c>w:pPr</c> of a paragraph, a style, the defaults or a numbering level).</summary>
        public void MergeParagraph(OpenXmlElement pPr)
        {
            foreach (var child in pPr.ChildElements)
            {
                switch (child)
                {
                    case Justification jc when jc.Val is not null:
                        Align = jc.Val.InnerText;
                        break;
                    case SpacingBetweenLines spacing:
                        if (Number(spacing.Before?.Value) is { } before)
                        {
                            Before = Math.Max(0, before) / 20;
                        }

                        if (Number(spacing.After?.Value) is { } after)
                        {
                            After = Math.Max(0, after) / 20;
                        }

                        if (Number(spacing.Line?.Value) is { } line && line > 0)
                        {
                            Line = line;
                            LineRule = spacing.LineRule?.InnerText ?? "auto";
                        }

                        break;
                    case Indentation ind:
                        if (Number(ind.Left?.Value ?? ind.Start?.Value) is { } left)
                        {
                            Left = left / 20;
                        }

                        if (Number(ind.Right?.Value ?? ind.End?.Value) is { } right)
                        {
                            Right = right / 20;
                        }

                        if (Number(ind.Hanging?.Value) is { } hanging)
                        {
                            Hanging = Math.Max(0, hanging) / 20;
                            FirstLine = null;
                        }
                        else if (Number(ind.FirstLine?.Value) is { } first)
                        {
                            FirstLine = first / 20;
                            Hanging = null;
                        }

                        break;
                    case Shading shading:
                        ParagraphShading = ShadingColour(shading) ?? string.Empty;
                        break;
                    case ParagraphBorders borders:
                        foreach (var (side, border) in ReadBorders(borders))
                        {
                            Borders[side] = border;
                        }

                        break;
                    case NumberingProperties numPr:
                        if (numPr.NumberingId?.Val?.Value is { } numId)
                        {
                            NumId = numId;
                            NumStyleId = null;
                        }

                        if (numPr.NumberingLevelReference?.Val?.Value is { } ilvl)
                        {
                            Ilvl = ilvl;
                        }

                        break;
                    case PageBreakBefore pageBreak:
                        PageBreakBefore = OnOff(pageBreak);
                        break;
                    case OutlineLevel outline when outline.Val?.Value is { } level:
                        OutlineLevel = level;
                        break;
                }
            }
        }

        /// <summary>Merges run properties (<c>w:rPr</c> of a run, a style or the defaults).</summary>
        public void MergeRun(OpenXmlElement rPr, StyleSheet styles)
        {
            foreach (var child in rPr.ChildElements)
            {
                switch (child)
                {
                    case RunFonts fonts:
                        Font = styles.FontOf(fonts) ?? Font;
                        break;
                    case Bold bold:
                        Bold = OnOff(bold);
                        break;
                    case Italic italic:
                        Italic = OnOff(italic);
                        break;
                    case Strike strike:
                        Strike = OnOff(strike);
                        break;
                    case DoubleStrike doubleStrike when OnOff(doubleStrike):
                        Strike = true;
                        break;
                    case Vanish vanish:
                        Vanish = OnOff(vanish);
                        break;
                    case Color color when color.Val?.Value is { Length: > 0 } value:
                        Color = value.Equals("auto", StringComparison.OrdinalIgnoreCase) ? "auto" : IsHex(value) ? value : Color;
                        break;
                    case Spacing spacing when spacing.Val?.Value is { } twips:
                        Spacing = Math.Clamp(twips / 20.0, -50, 100);
                        break;
                    case FontSize size when HalfPoints(size.Val?.Value) is { } half:
                        Size = half / 2;
                        break;
                    case Highlight highlight:
                        Highlight = highlight.Val?.InnerText is { } name && name != "none" ? name : string.Empty;
                        break;
                    case Underline underline:
                        Underline = underline.Val?.InnerText is not "none";
                        break;
                    case Shading shading:
                        Shading = ShadingColour(shading) ?? string.Empty;
                        break;
                    case VerticalTextAlignment align:
                        Script = align.Val?.InnerText switch
                        {
                            "superscript" => 1,
                            "subscript" => -1,
                            _ => 0,
                        };
                        break;
                }
            }
        }

        private static double? Number(string? value)
            => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? n : null;

        /// <summary>A font size in half points (<c>22</c>; the strict form <c>11pt</c> too), or null.</summary>
        private static double? HalfPoints(string? value)
        {
            if (value is null)
            {
                return null;
            }

            var v = value.Trim();
            var half = v.EndsWith("pt", StringComparison.OrdinalIgnoreCase) && Number(v[..^2]) is { } pt ? pt * 2 : Number(v);
            return half is > 0 and <= 3276 ? half : null;
        }
    }
}
