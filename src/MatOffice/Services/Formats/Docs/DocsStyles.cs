using System.Collections.Frozen;

namespace MatOffice.Services.Formats.Docs;

/// <summary>A paragraph style of Docs: the same values as the editor's CSS (<c>wwwroot/css/docs.css</c>) and the styles of the Word files.</summary>
internal sealed record ParagraphStyle
{
    /// <summary>The key (and the style id in Word files): Normal, Heading1, Title …</summary>
    public required string Key { get; init; }

    /// <summary>The name in Word (built-in names are English in every Word): "Normal", "heading 1", "Title" …</summary>
    public required string WordName { get; init; }

    public string Font { get; init; } = "Calibri";

    public double Size { get; init; } = 11;

    /// <summary><c>#rrggbb</c>, or null for "automatic".</summary>
    public string? Color { get; init; }

    public bool Italic { get; init; }

    /// <summary>Letter spacing in points.</summary>
    public double LetterSpacing { get; init; }

    public double Before { get; init; }

    public double After { get; init; } = 8;

    /// <summary>The CSS line height L (Word's "multiple" × 1.2).</summary>
    public double Line { get; init; } = 1.3;

    /// <summary>The line spacing in Word (240 = single).</summary>
    public int LineTwips { get; init; } = 259;

    public string Align { get; init; } = "left";

    public double IndentLeft { get; init; }

    public double IndentRight { get; init; }

    public string? Background { get; init; }

    /// <summary>The colour of a top and bottom border (Intense Quote).</summary>
    public string? BorderColor { get; init; }

    public double Padding { get; init; }

    public bool KeepNext { get; init; }

    /// <summary>The outline level of a heading (0 = Heading 1).</summary>
    public int? OutlineLevel { get; init; }

    /// <summary>No space between paragraphs of the same style (List Paragraph).</summary>
    public bool Contextual { get; init; }

    /// <summary>The style in CSS, for a file that has no style sheet (HTML export).</summary>
    public CssList ToCss(bool block = true)
    {
        var css = new CssList();
        css.Set("font-family", DocsCss.FontStack(Font));
        css.Set("font-size", CssList.Pt(Size));
        css.Set("font-weight", "normal");
        if (Italic)
        {
            css.Set("font-style", "italic");
        }

        if (Color is not null)
        {
            css.Set("color", Color);
        }

        if (LetterSpacing != 0)
        {
            css.Set("letter-spacing", CssList.Pt(LetterSpacing));
        }

        if (!block)
        {
            return css;
        }

        css.Set("line-height", CssList.Num(Line));
        css.Set("margin-top", CssList.Pt(Before));
        css.Set("margin-bottom", CssList.Pt(After));
        css.Set("margin-left", CssList.Pt(IndentLeft));
        css.Set("margin-right", CssList.Pt(IndentRight));
        if (Align != "left")
        {
            css.Set("text-align", Align);
        }

        if (Background is not null)
        {
            css.Set("background-color", Background);
        }

        if (BorderColor is not null)
        {
            css.Set("border-top", "0.5pt solid " + BorderColor);
            css.Set("border-bottom", "0.5pt solid " + BorderColor);
            css.Set("padding-top", CssList.Pt(Padding));
            css.Set("padding-bottom", CssList.Pt(Padding));
        }

        return css;
    }
}

/// <summary>The look of a table style: borders, header row, banded rows.</summary>
internal sealed record TableLook(string Key, string? Border, bool AllBorders, string? HeaderFill, string? HeaderText, string? BandFill);

/// <summary>The paragraph styles, the list kinds and the table styles of Docs.</summary>
internal static class DocsStyles
{
    public const string BodyFont = "Calibri";

    public const string HeadingFont = "Calibri Light";

    public const string LinkColour = "#0563c1";

    public static readonly ParagraphStyle Normal = new() { Key = "Normal", WordName = "Normal" };

    public static readonly IReadOnlyList<ParagraphStyle> All =
    [
        Normal,
        new() { Key = "NoSpacing", WordName = "No Spacing", After = 0, Line = 1.2, LineTwips = 240 },
        new() { Key = "Title", WordName = "Title", Font = HeadingFont, Size = 28, LetterSpacing = -0.5, Line = 1.2, LineTwips = 240, After = 4 },
        new() { Key = "Subtitle", WordName = "Subtitle", Color = "#5a5a5a", LetterSpacing = 0.75 },
        Heading(1, 16, "#2f5496", 12),
        Heading(2, 13, "#2f5496", 2),
        Heading(3, 12, "#1f3763", 2),
        Heading(4, 11, "#2f5496", 2, italic: true),
        Heading(5, 11, "#2f5496", 2),
        Heading(6, 11, "#1f3763", 2),
        new() { Key = "Quote", WordName = "Quote", Italic = true, Color = "#404040", Align = "center", IndentLeft = 43.2, IndentRight = 43.2, Before = 10 },
        new()
        {
            Key = "IntenseQuote", WordName = "Intense Quote", Italic = true, Color = "#2f5496", Align = "center", IndentLeft = 43.2, IndentRight = 43.2,
            BorderColor = "#4472c4", Padding = 10, Before = 18, After = 18,
        },
        new() { Key = "ListParagraph", WordName = "List Paragraph", IndentLeft = 36, Contextual = true },
        new() { Key = "Caption", WordName = "caption", Italic = true, Size = 9, Color = "#44546a", After = 10 },
        new() { Key = "Code", WordName = "Code", Font = "Consolas", Size = 10, Background = "#f2f2f2", Line = 1.2, LineTwips = 240 },
        new() { Key = "TOCHeading", WordName = "TOC Heading", Font = HeadingFont, Size = 16, Color = "#2f5496", Before = 12, After = 0, KeepNext = true },
        new() { Key = "TOC1", WordName = "toc 1", After = 5 },
        new() { Key = "TOC2", WordName = "toc 2", After = 5, IndentLeft = 11 },
        new() { Key = "TOC3", WordName = "toc 3", After = 5, IndentLeft = 22 },
        new() { Key = "Note", WordName = "footnote text", Size = 10, After = 0, Line = 1.2, LineTwips = 240 },
    ];

    private static readonly FrozenDictionary<string, ParagraphStyle> ByKey = All.ToFrozenDictionary(s => s.Key, StringComparer.Ordinal);

    private static readonly FrozenDictionary<string, string> ByWordName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["normal"] = "Normal", ["no spacing"] = "NoSpacing", ["title"] = "Title", ["subtitle"] = "Subtitle",
        ["heading 1"] = "Heading1", ["heading 2"] = "Heading2", ["heading 3"] = "Heading3", ["heading 4"] = "Heading4", ["heading 5"] = "Heading5",
        ["heading 6"] = "Heading6", ["heading 7"] = "Heading6", ["heading 8"] = "Heading6", ["heading 9"] = "Heading6",
        ["quote"] = "Quote", ["intense quote"] = "IntenseQuote", ["list paragraph"] = "ListParagraph", ["caption"] = "Caption",
        ["toc heading"] = "TOCHeading", ["toc 1"] = "TOC1", ["toc 2"] = "TOC2", ["toc 3"] = "TOC3", ["toc 4"] = "TOC3", ["toc 5"] = "TOC3",
        ["toc 6"] = "TOC3", ["toc 7"] = "TOC3", ["toc 8"] = "TOC3", ["toc 9"] = "TOC3",
        ["code"] = "Code", ["html preformatted"] = "Code", ["source code"] = "Code", ["macro text"] = "Code",
        ["footnote text"] = "Note", ["endnote text"] = "Note",
    }.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);

    /// <summary>The <c>data-style</c> values of a <c>p</c>.</summary>
    public static readonly FrozenSet<string> ParagraphDataStyles = new[]
    {
        "NoSpacing", "Title", "Subtitle", "ListParagraph", "Caption", "TOCHeading", "TOC1", "TOC2", "TOC3",
    }.ToFrozenSet(StringComparer.Ordinal);

    /// <summary>Every <c>data-style</c> value (IntenseQuote belongs to a blockquote).</summary>
    public static readonly FrozenSet<string> DataStyles = ParagraphDataStyles.Append("IntenseQuote").ToFrozenSet(StringComparer.Ordinal);

    public static readonly FrozenSet<string> BulletKinds = new[] { "bullet", "circle", "square", "dash", "check", "arrow" }.ToFrozenSet(StringComparer.Ordinal);

    public static readonly FrozenSet<string> NumberKinds = new[] { "decimal", "lower-alpha", "upper-alpha", "lower-roman", "upper-roman", "outline" }.ToFrozenSet(StringComparer.Ordinal);

    public static readonly FrozenSet<string> TableStyles = new[] { "grid", "plain", "light", "striped", "accent1", "accent2", "accent6" }.ToFrozenSet(StringComparer.Ordinal);

    public static readonly FrozenSet<string> Fields = new[] { "page", "pages", "date", "time", "title" }.ToFrozenSet(StringComparer.Ordinal);

    public static bool IsListKind(string? kind) => kind is not null && (BulletKinds.Contains(kind) || NumberKinds.Contains(kind));

    public static ParagraphStyle Get(string key) => ByKey.GetValueOrDefault(key) ?? Normal;

    /// <summary>The style key of a Word style name ("heading 1" → Heading1), or null for a style Docs does not have.</summary>
    public static string? FromWordName(string? name) => name is not null && ByWordName.TryGetValue(name.Trim(), out var key) ? key : null;

    /// <summary>Gives a paragraph the tag and data-style of a style key.</summary>
    public static void Apply(DocParagraph p, string key)
    {
        p.Style = null;
        switch (key)
        {
            case "Normal" or "Note":
                p.Tag = "p";
                break;
            case "Quote":
                p.Tag = "blockquote";
                break;
            case "IntenseQuote":
                p.Tag = "blockquote";
                p.Style = "IntenseQuote";
                break;
            case "Code":
                p.Tag = "pre";
                break;
            case var h when h.StartsWith("Heading", StringComparison.Ordinal):
                p.Tag = "h" + h[^1];
                break;
            default:
                p.Tag = "p";
                p.Style = key;
                break;
        }
    }

    /// <summary>
    /// The format of list level <paramref name="level"/> of a kind: Word's numFmt (<c>bullet</c>, <c>decimal</c>, <c>lowerLetter</c> …) and the marker
    /// (the bullet character, or the CSS list-style-type). <c>bullet</c> cycles • ◦ ▪ and <c>decimal</c> 1. a. i. with the level; the other kinds keep their marker.
    /// </summary>
    public static string LevelFormat(string kind, int level) => kind switch
    {
        "bullet" => (level % 3) switch { 0 => "bullet", 1 => "circle", _ => "square" },
        "decimal" => (level % 3) switch { 0 => "decimal", 1 => "lower-alpha", _ => "lower-roman" },
        "outline" => "decimal",
        _ => kind,
    };

    /// <summary>The text of a list marker for a number (plain text exports).</summary>
    public static string Marker(string kind, int level, int number, IReadOnlyList<int>? outline = null)
    {
        var format = LevelFormat(kind, level);
        return format switch
        {
            "bullet" => "•",
            "circle" => "◦",
            "square" => "▪",
            "dash" => "–",
            "check" => "✓",
            "arrow" => "➢",
            "lower-alpha" => Alpha(number) + ".",
            "upper-alpha" => Alpha(number).ToUpperInvariant() + ".",
            "lower-roman" => Roman(number).ToLowerInvariant() + ".",
            "upper-roman" => Roman(number) + ".",
            _ when kind == "outline" && outline is not null => string.Join('.', outline) + ".",
            _ => number + ".",
        };
    }

    public static string Alpha(int n)
    {
        if (n <= 0)
        {
            return "0";
        }

        var s = string.Empty;
        while (n > 0)
        {
            n--;
            s = (char)('a' + (n % 26)) + s;
            n /= 26;
        }

        return s;
    }

    public static string Roman(int n)
    {
        if (n is <= 0 or >= 4000)
        {
            return n.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        (int Value, string Text)[] table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        var s = string.Empty;
        foreach (var (value, text) in table)
        {
            while (n >= value)
            {
                s += text;
                n -= value;
            }
        }

        return s;
    }

    public static TableLook Look(string style) => style switch
    {
        "plain" => new("plain", null, false, null, null, null),
        "light" => new("light", "#bfbfbf", false, null, null, null),
        "striped" => new("striped", null, false, null, null, "#f2f2f2"),
        "accent1" => new("accent1", "#8eaadb", true, "#4472c4", "#ffffff", "#d9e2f3"),
        "accent2" => new("accent2", "#f4b083", true, "#ed7d31", "#ffffff", "#fbe4d5"),
        "accent6" => new("accent6", "#a8d08d", true, "#70ad47", "#ffffff", "#e2efd9"),
        _ => new("grid", "#000000", true, null, null, null),
    };

    /// <summary>The direct look of one cell of a table style (borders and shading) as CSS: what Word files and HTML files get.</summary>
    public static CssList CellCss(TableLook look, bool headerRow, bool firstRow, int bodyIndex)
    {
        var css = new CssList();
        if (look.Border is { } border)
        {
            var line = "0.5pt solid " + border;
            if (look.AllBorders)
            {
                css.Set("border-top", line);
                css.Set("border-right", line);
                css.Set("border-bottom", line);
                css.Set("border-left", line);
            }
            else
            {
                if (firstRow)
                {
                    css.Set("border-top", line);
                }

                css.Set("border-bottom", line);
            }
        }

        if (headerRow && look.HeaderFill is not null)
        {
            css.Set("background-color", look.HeaderFill);
        }
        else if (!headerRow && look.BandFill is not null && bodyIndex % 2 == 0)
        {
            css.Set("background-color", look.BandFill);
        }

        return css;
    }

    private static ParagraphStyle Heading(int level, double size, string colour, double before, bool italic = false) => new()
    {
        Key = "Heading" + level,
        WordName = "heading " + level,
        Font = HeadingFont,
        Size = size,
        Color = colour,
        Italic = italic,
        Before = before,
        After = 0,
        KeepNext = true,
        OutlineLevel = level - 1,
    };
}
