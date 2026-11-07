using System.Collections.Frozen;
using System.Globalization;
using System.Text;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The CSS of Docs HTML: the allowlist of the <c>style</c> attribute (every declaration is checked here, the sanitizer never sees a style),
/// and the helpers of the converters to read lengths, colours and fonts.
/// </summary>
internal static class DocsCss
{
    /// <summary>The properties a style attribute may hold; every other declaration is dropped.</summary>
    public static readonly FrozenSet<string> Allowed = new[]
    {
        "font-family", "font-size", "font-weight", "font-style", "text-decoration", "text-decoration-line", "color", "background-color",
        "text-align", "line-height", "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
        "padding", "padding-top", "padding-right", "padding-bottom", "padding-left", "text-indent", "width", "height", "vertical-align",
        "border", "border-top", "border-right", "border-bottom", "border-left",
        "border-top-width", "border-top-style", "border-top-color", "border-right-width", "border-right-style", "border-right-color",
        "border-bottom-width", "border-bottom-style", "border-bottom-color", "border-left-width", "border-left-style", "border-left-color",
        "border-collapse", "float", "list-style-type", "letter-spacing",
    }.ToFrozenSet(StringComparer.Ordinal);

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    private static readonly string[] Forbidden = ["url(", "expression", "javascript", "vbscript", "@import", "var(", "--", "behavior", "binding", "image(", "image-set("];

    private static readonly FrozenSet<string> Functions = new[] { "rgb", "rgba", "hsl", "hsla" }.ToFrozenSet(StringComparer.Ordinal);

    /// <summary>Reads a style attribute: only allowed properties with safe values, shorthands of margin, padding and border spread to the sides.</summary>
    public static CssList Parse(string? style)
    {
        var css = new CssList();
        if (string.IsNullOrWhiteSpace(style))
        {
            return css;
        }

        foreach (var (rawName, rawValue) in Declarations(style))
        {
            var name = rawName.Trim().ToLowerInvariant();
            if (name == "background")
            {
                // A plain colour in the shorthand (Word writes "background:yellow"); anything else is dropped.
                if (Colour(rawValue) is { } bg)
                {
                    css.Set("background-color", bg);
                }

                continue;
            }

            if (!Allowed.Contains(name))
            {
                continue;
            }

            var value = CleanValue(rawValue);
            if (value is null || !IsSafe(name, value))
            {
                continue;
            }

            switch (name)
            {
                case "margin" or "padding":
                    var parts = SplitSpaces(value);
                    if (parts.Count is >= 1 and <= 4)
                    {
                        var top = parts[0];
                        var right = parts.Count > 1 ? parts[1] : top;
                        var bottom = parts.Count > 2 ? parts[2] : top;
                        var left = parts.Count > 3 ? parts[3] : right;
                        css.Set(name + "-top", top);
                        css.Set(name + "-right", right);
                        css.Set(name + "-bottom", bottom);
                        css.Set(name + "-left", left);
                    }

                    break;
                case "border":
                    foreach (var side in Sides)
                    {
                        css.Set("border-" + side, value);
                    }

                    break;
                default:
                    css.Set(name, value);
                    break;
            }
        }

        return css;
    }

    /// <summary>Splits a style attribute into declarations (semicolons inside quotes or parentheses do not split).</summary>
    public static IEnumerable<(string Name, string Value)> Declarations(string style)
    {
        var start = 0;
        var depth = 0;
        var quote = '\0';
        for (var i = 0; i <= style.Length; i++)
        {
            var c = i < style.Length ? style[i] : ';';
            if (quote != '\0')
            {
                if (c == quote)
                {
                    quote = '\0';
                }

                if (i < style.Length)
                {
                    continue;
                }
            }

            switch (c)
            {
                case '"' or '\'':
                    quote = c;
                    break;
                case '(':
                    depth++;
                    break;
                case ')':
                    depth = Math.Max(0, depth - 1);
                    break;
                case ';' when depth == 0 || i == style.Length:
                    var declaration = style[start..i];
                    start = i + 1;
                    var colon = declaration.IndexOf(':', StringComparison.Ordinal);
                    if (colon > 0)
                    {
                        yield return (declaration[..colon].Trim(), declaration[(colon + 1)..].Trim());
                    }

                    break;
            }
        }
    }

    /// <summary>The value without <c>!important</c>, whitespace collapsed; null when empty.</summary>
    private static string? CleanValue(string value)
    {
        var v = value.Trim();
        var bang = v.LastIndexOf('!');
        if (bang >= 0 && v[(bang + 1)..].Trim().Equals("important", StringComparison.OrdinalIgnoreCase))
        {
            v = v[..bang].Trim();
        }

        if (v.Length == 0)
        {
            return null;
        }

        var sb = new StringBuilder(v.Length);
        var space = false;
        foreach (var c in v)
        {
            if (char.IsWhiteSpace(c))
            {
                space = true;
                continue;
            }

            if (space && sb.Length > 0)
            {
                sb.Append(' ');
            }

            space = false;
            sb.Append(c);
        }

        return sb.ToString();
    }

    /// <summary>
    /// A value is safe when it has no way to load or run anything: no <c>url(</c>, <c>expression(</c>, <c>var(</c>, escapes, comments,
    /// at-rules or markup, functions only <c>rgb</c>/<c>rgba</c>/<c>hsl</c>/<c>hsla</c>, quotes only in font names and list markers.
    /// </summary>
    public static bool IsSafe(string name, string value)
    {
        if (value.Length is 0 or > 300)
        {
            return false;
        }

        foreach (var word in Forbidden)
        {
            if (value.Contains(word, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        var quotesAllowed = name is "font-family" or "list-style-type";
        var quote = '\0';
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (char.IsControl(c) || c is '\\' or '<' or '>' or '{' or '}' or ';' or '@' or '/' or '*' or '!' or '&' or '=' or '`')
            {
                return false;
            }

            if (c is '"' or '\'')
            {
                if (!quotesAllowed)
                {
                    return false;
                }

                quote = quote == '\0' ? c : quote == c ? '\0' : quote;
                continue;
            }

            if (c == '(' && quote == '\0')
            {
                var j = i;
                while (j > 0 && (char.IsAsciiLetter(value[j - 1]) || value[j - 1] == '-'))
                {
                    j--;
                }

                if (!Functions.Contains(value[j..i].ToLowerInvariant()))
                {
                    return false;
                }
            }
        }

        return quote == '\0';
    }

    /// <summary>Splits a value at spaces outside parentheses (<c>1pt solid rgb(0, 0, 0)</c> has three parts).</summary>
    public static List<string> SplitSpaces(string value)
    {
        var parts = new List<string>();
        var depth = 0;
        var start = 0;
        for (var i = 0; i <= value.Length; i++)
        {
            var c = i < value.Length ? value[i] : ' ';
            if (c == '(')
            {
                depth++;
            }
            else if (c == ')')
            {
                depth = Math.Max(0, depth - 1);
            }
            else if (c == ' ' && depth == 0)
            {
                if (i > start)
                {
                    parts.Add(value[start..i]);
                }

                start = i + 1;
            }
        }

        return parts;
    }

    /// <summary>A CSS length in points (<c>pt</c>, <c>px</c>, <c>mm</c>, <c>cm</c>, <c>in</c>, <c>pc</c>, <c>em</c>; a bare number is pixels), or null.</summary>
    public static double? ToPt(string? value, double emPt = 11)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var v = value.Trim().ToLowerInvariant();
        var end = 0;
        while (end < v.Length && (char.IsAsciiDigit(v[end]) || v[end] is '.' or '-' or '+'))
        {
            end++;
        }

        if (end == 0 || !double.TryParse(v[..end], NumberStyles.Float, CultureInfo.InvariantCulture, out var n) || !double.IsFinite(n))
        {
            return v switch
            {
                "thin" => 0.75,
                "medium" => 2.25,
                "thick" => 3.75,
                _ => null,
            };
        }

        return v[end..].Trim() switch
        {
            "pt" => n,
            "px" or "" => n * 0.75,
            "mm" => n * 72 / 25.4,
            "cm" => n * 72 / 2.54,
            "q" => n * 72 / 101.6,
            "in" => n * 72,
            "pc" => n * 12,
            "em" or "rem" => n * emPt,
            "ex" or "ch" => n * emPt / 2,
            _ => null,
        };
    }

    /// <summary>A CSS length in pixels, or null.</summary>
    public static double? ToPx(string? value) => ToPt(value) is { } pt ? pt / 0.75 : null;

    /// <summary>A percentage (<c>30%</c> → 30), or null.</summary>
    public static double? Percent(string? value)
    {
        if (value is null)
        {
            return null;
        }

        var v = value.Trim();
        return v.EndsWith('%') && double.TryParse(v[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var p) && double.IsFinite(p) ? p : null;
    }

    /// <summary>A font size as a CSS value in points (keywords and pixels converted), or null.</summary>
    public static string? FontSizePt(string? value)
    {
        if (value is null)
        {
            return null;
        }

        var keyword = value.Trim().ToLowerInvariant() switch
        {
            "xx-small" => 7.0,
            "x-small" => 7.5,
            "small" => 10.0,
            "medium" => 12.0,
            "large" => 13.5,
            "x-large" => 18.0,
            "xx-large" => 24.0,
            "xxx-large" => 36.0,
            _ => 0.0,
        };
        if (keyword > 0)
        {
            return CssList.Pt(keyword);
        }

        if (Percent(value) is { } p)
        {
            return CssList.Pt(11 * p / 100);
        }

        return ToPt(value) is { } pt && pt is > 0 and <= 1638 ? CssList.Pt(pt) : null;
    }

    /// <summary>A colour as <c>#rrggbb</c> (lower case), or null (transparent, unknown, inherit …).</summary>
    public static string? Colour(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var v = value.Trim().ToLowerInvariant();
        if (v.StartsWith('#'))
        {
            var hex = v[1..];
            if (!hex.All(char.IsAsciiHexDigit))
            {
                return null;
            }

            return hex.Length switch
            {
                3 or 4 => $"#{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}",
                6 or 8 => "#" + hex[..6],
                _ => null,
            };
        }

        if (v.StartsWith("rgb", StringComparison.Ordinal) || v.StartsWith("hsl", StringComparison.Ordinal))
        {
            var open = v.IndexOf('(', StringComparison.Ordinal);
            var close = v.LastIndexOf(')');
            if (open < 0 || close < open)
            {
                return null;
            }

            var args = v[(open + 1)..close].Replace('/', ' ').Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
            if (args.Length < 3)
            {
                return null;
            }

            if (args.Length > 3 && Channel(args[3], 1) is { } alpha && alpha <= 0.01)
            {
                return null;
            }

            if (v.StartsWith("rgb", StringComparison.Ordinal))
            {
                if (Channel(args[0], 255) is { } r && Channel(args[1], 255) is { } g && Channel(args[2], 255) is { } b)
                {
                    return Hex(r, g, b);
                }

                return null;
            }

            if (double.TryParse(args[0].Replace("deg", string.Empty, StringComparison.Ordinal), NumberStyles.Float, CultureInfo.InvariantCulture, out var h)
                && Channel(args[1], 1) is { } s && Channel(args[2], 1) is { } l)
            {
                return FromHsl(h, s, l);
            }

            return null;
        }

        return Named.TryGetValue(v, out var named) ? named : null;
    }

    /// <summary>A colour as <c>RRGGBB</c> for Word, or null.</summary>
    public static string? WordColour(string? value) => Colour(value) is { } c ? c[1..].ToUpperInvariant() : null;

    private static double? Channel(string arg, double max)
    {
        if (arg.EndsWith('%'))
        {
            return double.TryParse(arg[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var p) ? Math.Clamp(p / 100 * max, 0, max) : null;
        }

        return double.TryParse(arg, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? Math.Clamp(n, 0, max) : null;
    }

    private static string Hex(double r, double g, double b)
        => string.Create(CultureInfo.InvariantCulture, $"#{(int)Math.Round(r):x2}{(int)Math.Round(g):x2}{(int)Math.Round(b):x2}");

    private static string FromHsl(double h, double s, double l)
    {
        h = ((h % 360) + 360) % 360 / 360;
        double Hue(double t)
        {
            t = (t + 1) % 1;
            var q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
            var p = (2 * l) - q;
            return t < 1.0 / 6 ? p + ((q - p) * 6 * t) : t < 0.5 ? q : t < 2.0 / 3 ? p + ((q - p) * ((2.0 / 3) - t) * 6) : p;
        }

        return s == 0 ? Hex(l * 255, l * 255, l * 255) : Hex(Hue(h + (1.0 / 3)) * 255, Hue(h) * 255, Hue(h - (1.0 / 3)) * 255);
    }

    private static readonly FrozenDictionary<string, string> Named = new Dictionary<string, string>
    {
        ["black"] = "#000000", ["white"] = "#ffffff", ["red"] = "#ff0000", ["green"] = "#008000", ["blue"] = "#0000ff", ["yellow"] = "#ffff00",
        ["lime"] = "#00ff00", ["aqua"] = "#00ffff", ["cyan"] = "#00ffff", ["fuchsia"] = "#ff00ff", ["magenta"] = "#ff00ff", ["silver"] = "#c0c0c0",
        ["gray"] = "#808080", ["grey"] = "#808080", ["maroon"] = "#800000", ["olive"] = "#808000", ["purple"] = "#800080", ["teal"] = "#008080",
        ["navy"] = "#000080", ["orange"] = "#ffa500", ["darkblue"] = "#00008b", ["darkred"] = "#8b0000", ["darkgreen"] = "#006400",
        ["darkgray"] = "#a9a9a9", ["darkgrey"] = "#a9a9a9", ["lightgray"] = "#d3d3d3", ["lightgrey"] = "#d3d3d3", ["pink"] = "#ffc0cb",
        ["brown"] = "#a52a2a", ["gold"] = "#ffd700", ["violet"] = "#ee82ee", ["indigo"] = "#4b0082", ["crimson"] = "#dc143c",
        ["orangered"] = "#ff4500", ["royalblue"] = "#4169e1", ["steelblue"] = "#4682b4", ["lightblue"] = "#add8e6", ["lightgreen"] = "#90ee90",
        ["lightyellow"] = "#ffffe0", ["whitesmoke"] = "#f5f5f5", ["gainsboro"] = "#dcdcdc", ["dimgray"] = "#696969", ["dimgrey"] = "#696969",
        ["windowtext"] = "#000000", ["window"] = "#ffffff",
    }.ToFrozenDictionary(StringComparer.Ordinal);

    /// <summary>The first family of a font-family list, without quotes.</summary>
    public static string FirstFamily(string fontFamily)
    {
        var quote = '\0';
        for (var i = 0; i < fontFamily.Length; i++)
        {
            var c = fontFamily[i];
            if (quote != '\0')
            {
                quote = c == quote ? '\0' : quote;
            }
            else if (c is '"' or '\'')
            {
                quote = c;
            }
            else if (c == ',')
            {
                return Unquote(fontFamily[..i]);
            }
        }

        return Unquote(fontFamily);
    }

    private static string Unquote(string name) => name.Trim().Trim('"', '\'').Trim();

    private static readonly FrozenSet<string> SerifFonts = new[]
    {
        "times new roman", "times", "cambria", "georgia", "garamond", "book antiqua", "palatino", "palatino linotype", "constantia", "baskerville",
        "baskerville old face", "bookman old style", "century", "century schoolbook", "liberation serif", "caladea", "noto serif", "dejavu serif",
        "sitka text", "sylfaen", "perpetua", "goudy old style", "minion pro", "bodoni mt", "rockwell", "high tower text",
    }.ToFrozenSet(StringComparer.Ordinal);

    private static readonly FrozenSet<string> MonoFonts = new[]
    {
        "consolas", "courier new", "courier", "lucida console", "cascadia code", "cascadia mono", "liberation mono", "dejavu sans mono",
        "source code pro", "fira code", "menlo", "monaco", "andale mono", "noto sans mono", "roboto mono", "ocr a extended",
    }.ToFrozenSet(StringComparer.Ordinal);

    private static readonly FrozenDictionary<string, string> Substitutes = new Dictionary<string, string>
    {
        ["calibri"] = "Carlito", ["calibri light"] = "Carlito", ["cambria"] = "Caladea", ["arial"] = "Liberation Sans", ["helvetica"] = "Arial",
        ["times new roman"] = "Liberation Serif", ["courier new"] = "Liberation Mono", ["consolas"] = "'Courier New'", ["georgia"] = "'Times New Roman'",
        ["segoe ui"] = "'Helvetica Neue'", ["aptos"] = "Calibri", ["aptos display"] = "Calibri",
    }.ToFrozenDictionary(StringComparer.Ordinal);

    /// <summary><c>serif</c>, <c>monospace</c> or <c>sans-serif</c> for a font name.</summary>
    public static string Generic(string family)
    {
        var name = family.Trim().ToLowerInvariant();
        return SerifFonts.Contains(name) ? "serif" : MonoFonts.Contains(name) || name.Contains("mono", StringComparison.Ordinal) ? "monospace" : "sans-serif";
    }

    /// <summary>A font-family value for a font name with fallbacks: <c>Calibri, Carlito, sans-serif</c>.</summary>
    public static string FontStack(string family)
    {
        var name = new string(family.Where(c => c is not ('"' or '\'' or ';' or '\\' or '<' or '>' or '{' or '}')).ToArray()).Trim();
        if (name.Length == 0)
        {
            return "sans-serif";
        }

        var lower = name.ToLowerInvariant();
        if (lower is "serif" or "sans-serif" or "monospace")
        {
            return lower;
        }

        var quoted = name.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-') ? "'" + name + "'" : name;
        return Substitutes.TryGetValue(lower, out var alias) ? $"{quoted}, {alias}, {Generic(name)}" : $"{quoted}, {Generic(name)}";
    }
}
