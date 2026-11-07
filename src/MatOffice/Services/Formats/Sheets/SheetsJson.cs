using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>Small readers for the content JSON (a wrong type reads as null) and the checks of single values.</summary>
internal static class SheetsJson
{
    /// <summary>The error values of a cell (<c>"t": "e"</c>); <c>#CIRC!</c> is the editor's own (a circular reference).</summary>
    public static readonly IReadOnlySet<string> Errors = new HashSet<string>(StringComparer.Ordinal)
    {
        "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!", "#CIRC!",
    };

    public static string? Str(JsonNode? node)
        => node is JsonValue v && v.GetValueKind() == JsonValueKind.String ? v.GetValue<string>() : null;

    public static double? Num(JsonNode? node)
        => node is JsonValue v && v.GetValueKind() == JsonValueKind.Number && v.TryGetValue(out double d) && double.IsFinite(d) ? d : null;

    public static bool? Bool(JsonNode? node) => node?.GetValueKind() switch
    {
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        _ => null,
    };

    public static bool IsTrue(JsonNode? node) => node?.GetValueKind() == JsonValueKind.True;

    /// <summary>A whole number in [min, max], or null.</summary>
    public static int? Int(JsonNode? node, int min, int max)
        => Num(node) is { } d && d == Math.Floor(d) && d >= min && d <= max ? (int)d : null;

    /// <summary>A number in [min, max], or null.</summary>
    public static double? Range(JsonNode? node, double min, double max)
        => Num(node) is { } d && d >= min && d <= max ? d : null;

    /// <summary>A string, or a number written invariantly (criteria of rules and filters are text).</summary>
    public static string? StrOrNum(JsonNode? node)
        => Str(node) ?? (Num(node) is { } d ? d.ToString("R", CultureInfo.InvariantCulture) : null);

    /// <summary>Text without control characters (tab and line feed stay; CR LF and CR become LF) and without broken surrogates, at most <paramref name="max"/> characters.</summary>
    public static string Clean(string text, int max)
    {
        var clean = true;
        foreach (var c in text)
        {
            if ((char.IsControl(c) && c is not '\t' and not '\n') || char.IsSurrogate(c))
            {
                clean = false;
                break;
            }
        }

        if (!clean)
        {
            var sb = new StringBuilder(text.Length);
            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                if (c == '\r')
                {
                    if (i + 1 >= text.Length || text[i + 1] != '\n')
                    {
                        sb.Append('\n');
                    }
                }
                else if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    sb.Append(c).Append(text[++i]);
                }
                else if (!char.IsSurrogate(c) && (!char.IsControl(c) || c is '\t' or '\n'))
                {
                    sb.Append(c);
                }
            }

            text = sb.ToString();
        }

        return Cut(text, max);
    }

    /// <summary>The first <paramref name="max"/> characters, never half a surrogate pair.</summary>
    public static string Cut(string text, int max)
    {
        if (text.Length <= max)
        {
            return text;
        }

        var end = char.IsHighSurrogate(text[max - 1]) ? max - 1 : max;
        return text[..end];
    }

    /// <summary>A colour <c>#rrggbb</c> (lower case), or null.</summary>
    public static string? Colour(string? text)
    {
        if (text is not { Length: 7 } || text[0] != '#')
        {
            return null;
        }

        for (var i = 1; i < 7; i++)
        {
            if (!char.IsAsciiHexDigit(text[i]))
            {
                return null;
            }
        }

        return text.ToLowerInvariant();
    }

    /// <summary>A font family: letters, digits, spaces and hyphens, 1-64 characters.</summary>
    public static string? Font(string? text)
    {
        text = text?.Trim();
        if (string.IsNullOrEmpty(text) || text.Length > 64)
        {
            return null;
        }

        return text.All(c => char.IsAsciiLetterOrDigit(c) || c is ' ' or '-') ? text : null;
    }

    /// <summary>An id <c>[A-Za-z0-9_-]{1,32}</c>, or null.</summary>
    public static string? Id(string? text)
        => text is { Length: >= 1 and <= 32 } && text.All(c => char.IsAsciiLetterOrDigit(c) || c is '_' or '-') ? text : null;

    /// <summary>
    /// A link the editor may open: <c>http(s)://…</c>, <c>mailto:…</c> or an internal one <c>#Sheet1!A1</c>. Everything else (javascript:, data:, file: …) is null.
    /// </summary>
    public static string? Link(string? text)
    {
        text = text?.Trim();
        if (string.IsNullOrEmpty(text) || text.Length > 2083 || text.Any(char.IsControl))
        {
            return null;
        }

        if (text[0] == '#')
        {
            return text.Length > 1 ? text : null;
        }

        if (text.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
        {
            return text.Length > 7 ? text : null;
        }

        return Uri.TryCreate(text, UriKind.Absolute, out var uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) ? text : null;
    }

    /// <summary>A picture as a data address: png, jpeg, gif or webp in base64 (no SVG: it could hold script).</summary>
    public static string? Image(string? text)
    {
        if (text is null)
        {
            return null;
        }

        var comma = text.IndexOf(',');
        if (comma < 0 || comma == text.Length - 1)
        {
            return null;
        }

        var head = text[..comma];
        if (head is not ("data:image/png;base64" or "data:image/jpeg;base64" or "data:image/gif;base64" or "data:image/webp;base64"))
        {
            return null;
        }

        for (var i = comma + 1; i < text.Length; i++)
        {
            var c = text[i];
            if (!char.IsAsciiLetterOrDigit(c) && c is not '+' and not '/' and not '=')
            {
                return null;
            }
        }

        return text;
    }

    /// <summary>A formula without the leading <c>=</c>, at most 8,192 characters, or null.</summary>
    public static string? Formula(string? text)
    {
        if (text is null)
        {
            return null;
        }

        text = Clean(text, int.MaxValue).Trim();
        if (text.StartsWith('='))
        {
            text = text[1..].TrimStart();
        }

        return text.Length is > 0 and <= SheetsLimits.MaxFormula ? text : null;
    }
}

/// <summary>The limits of a workbook (Excel's where it has one).</summary>
internal static class SheetsLimits
{
    public const int MaxSheets = 200;
    public const int MaxCells = 2_000_000;
    public const int MaxText = 32_767;
    public const int MaxFormula = 8_192;
    public const int MaxCharts = 200;
    public const int MaxImages = 200;
    public const int MaxRules = 500;
    public const int MaxValidations = 500;
    public const int MaxMerges = 10_000;
    public const int MaxNames = 1_000;
    public const int MaxStyles = 65_000;
    public const int DefaultColWidth = 64;
    public const int DefaultRowHeight = 20;
}
