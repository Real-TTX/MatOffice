using System.Globalization;
using System.Text.Json.Nodes;
using System.Xml.Linq;

namespace MatOffice.Services.Formats.Slides;

/// <summary>Small helpers of the PowerPoint conversion: namespaces, units, colours and the look of tables (the same rules as the editor's render.js).</summary>
internal static class PptxXml
{
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>EMU per px at 96 dpi.</summary>
    public const long EmuPerPx = 9525;

    public static long Emu(double px) => (long)Math.Round(px * EmuPerPx);

    public static double Px(long emu) => Math.Round(emu / (double)EmuPerPx, 2);

    public static string S(long v) => v.ToString(CultureInfo.InvariantCulture);

    public static string S(int v) => v.ToString(CultureInfo.InvariantCulture);

    /// <summary>A colour of the content as DrawingML: srgbClr, or schemeClr with lumMod/lumOff for ":+n" / ":-n"; alpha (0..1) optional.</summary>
    public static XElement Color(string color, double alpha = 1)
    {
        XElement e;
        if (color.StartsWith('#'))
        {
            e = new XElement(A + "srgbClr", new XAttribute("val", color[1..].ToUpperInvariant()));
        }
        else
        {
            var (name, mod) = SplitScheme(color);
            e = new XElement(A + "schemeClr", new XAttribute("val", SchemeOut(name)));
            if (mod > 0)
            {
                e.Add(new XElement(A + "lumMod", new XAttribute("val", S((100 - mod) * 1000))), new XElement(A + "lumOff", new XAttribute("val", S(mod * 1000))));
            }
            else if (mod < 0)
            {
                e.Add(new XElement(A + "lumMod", new XAttribute("val", S((100 + mod) * 1000))));
            }
        }

        if (alpha < 1)
        {
            e.Add(new XElement(A + "alpha", new XAttribute("val", S((int)Math.Round(Math.Clamp(alpha, 0, 1) * 100000)))));
        }

        return e;
    }

    public static XElement SolidFill(string color, double alpha = 1) => new(A + "solidFill", Color(color, alpha));

    /// <summary>"accent1:+40" → (accent1, 40).</summary>
    public static (string Name, int Mod) SplitScheme(string color)
    {
        var i = color.IndexOf(':', StringComparison.Ordinal);
        if (i < 0)
        {
            return (color, 0);
        }

        return (color[..i], int.Parse(color[(i + 1)..], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture));
    }

    /// <summary>The names of the slide (mapped through the colour map of the master): dk1 → tx1 …</summary>
    private static string SchemeOut(string name) => name switch
    {
        "dk1" => "tx1",
        "lt1" => "bg1",
        "dk2" => "tx2",
        "lt2" => "bg2",
        _ => name,
    };

    /// <summary>A colour of the content as #rrggbb (scheme colours resolved through the design, lighter/darker in HSL like the editor).</summary>
    public static string Resolve(string? color, JsonObject theme)
    {
        if (string.IsNullOrEmpty(color))
        {
            return "#000000";
        }

        if (color.StartsWith('#'))
        {
            return color;
        }

        var (name, mod) = SplitScheme(color);
        var hex = theme["colors"]?[name]?.GetValue<string>() ?? "#000000";
        return mod == 0 ? hex : Lum(hex, mod);
    }

    /// <summary>Lighter (+n %) or darker (-n %) in HSL, like lumMod/lumOff.</summary>
    public static string Lum(string hex, int mod)
    {
        var (h, s, l) = ToHsl(hex);
        l = mod > 0 ? (l * (1 - (mod / 100.0))) + (mod / 100.0) : l * (1 + (mod / 100.0));
        return FromHsl(h, s, Math.Clamp(l, 0, 1));
    }

    public static (double H, double S, double L) ToHsl(string hex)
    {
        var r = Convert.ToInt32(hex.Substring(1, 2), 16) / 255.0;
        var g = Convert.ToInt32(hex.Substring(3, 2), 16) / 255.0;
        var b = Convert.ToInt32(hex.Substring(5, 2), 16) / 255.0;
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        if (max == min)
        {
            return (0, 0, l);
        }

        var d = max - min;
        var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        double h;
        if (max == r)
        {
            h = ((g - b) / d) + (g < b ? 6 : 0);
        }
        else if (max == g)
        {
            h = ((b - r) / d) + 2;
        }
        else
        {
            h = ((r - g) / d) + 4;
        }

        return (h / 6, s, l);
    }

    public static string FromHsl(double h, double s, double l)
    {
        double r, g, b;
        if (s == 0)
        {
            r = g = b = l;
        }
        else
        {
            var q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
            var p = (2 * l) - q;
            r = Hue(p, q, h + (1 / 3.0));
            g = Hue(p, q, h);
            b = Hue(p, q, h - (1 / 3.0));
        }

        return $"#{(int)Math.Round(r * 255):x2}{(int)Math.Round(g * 255):x2}{(int)Math.Round(b * 255):x2}";
    }

    private static double Hue(double p, double q, double t)
    {
        if (t < 0)
        {
            t += 1;
        }

        if (t > 1)
        {
            t -= 1;
        }

        if (t < 1 / 6.0)
        {
            return p + ((q - p) * 6 * t);
        }

        if (t < 1 / 2.0)
        {
            return q;
        }

        return t < 2 / 3.0 ? p + ((q - p) * ((2 / 3.0) - t) * 6) : p;
    }

    /// <summary>The look of one cell of a table style: fill, text colour, bold, border colour (null = none). The editor computes the same in render.js.</summary>
    public static (string? Fill, string Text, bool Bold, string? Border) TableLook(JsonObject table, int row, int col)
    {
        var style = table["tableStyle"]?.GetValue<string>() ?? "medium";
        var accent = table["accent"]?.GetValue<string>() ?? "accent1";
        var rows = table["rows"]!.AsArray().Count;
        var header = table["header"]?.GetValue<bool>() ?? true;
        var banded = table["banded"]?.GetValue<bool>() ?? true;
        var firstCol = table["firstCol"]?.GetValue<bool>() ?? false;
        var lastRow = table["lastRow"]?.GetValue<bool>() ?? false;
        var isHeader = header && row == 0;
        var isLast = lastRow && row == rows - 1 && !isHeader;
        var isFirstCol = firstCol && col == 0 && !isHeader && !isLast;
        var bodyIndex = row - (header ? 1 : 0);
        var odd = banded && bodyIndex % 2 == 0;
        var (name, _) = SplitScheme(accent);
        switch (style)
        {
            case "none":
                return (null, "dk1", isHeader || isLast, "dk1:+65");
            case "light":
                if (isHeader)
                {
                    return (null, name + ":-25", true, name);
                }

                return (odd ? name + ":+80" : null, "dk1", isLast || isFirstCol, name + ":+40");
            case "dark":
                if (isHeader)
                {
                    return ("dk1", "lt1", true, "lt1");
                }

                if (isLast || isFirstCol)
                {
                    return (name + ":-50", "lt1", true, "lt1");
                }

                return (odd ? name + ":-25" : name, "lt1", false, "lt1");
            default:
                if (isHeader || isLast || isFirstCol)
                {
                    return (name, "lt1", true, "lt1");
                }

                return (odd ? name + ":+60" : name + ":+80", "dk1", false, "lt1");
        }
    }
}
