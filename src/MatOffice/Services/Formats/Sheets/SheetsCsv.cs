using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using MatOffice.Localization;
using static MatOffice.Services.Formats.Sheets.SheetsJson;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// .csv and .tsv: one sheet of values. Reading: UTF-8 (BOM optional; UTF-16 with BOM; else Windows-1252), RFC 4180 quotes, the delimiter of a .csv
/// detected among <c>,</c> <c>;</c> and tab, a decimal comma when the delimiter is <c>;</c> or such numbers dominate; numbers, booleans, dates, times and
/// percentages become values with a format. Writing: the active sheet, UTF-8 with BOM, CRLF, the last values of formulas; a text that starts with
/// <c>= + - @</c> gets a protecting <c>'</c> (CSV injection), which the import removes again.
/// </summary>
internal static partial class SheetsCsv
{
    private const string Delimiters = ",;\t";

    // ---- reading ---------------------------------------------------------------------------------------------------------------------

    public static string Read(Stream input, bool tabs)
    {
        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        var text = Decode(buffer.GetBuffer().AsSpan(0, (int)buffer.Length));
        var delimiter = tabs ? '\t' : Detect(text);
        var records = Parse(text, delimiter);
        var decimalComma = DecimalComma(records, delimiter);

        var styles = new JsonArray(new JsonObject());
        var styleOf = new Dictionary<string, int>(StringComparer.Ordinal);
        var cells = new JsonObject();
        var count = 0;
        for (var r = 0; r < records.Count && r < A1.MaxRows; r++)
        {
            var fields = records[r];
            for (var c = 0; c < fields.Count && c < A1.MaxCols; c++)
            {
                if (fields[c].Length == 0)
                {
                    continue;
                }

                if (++count > SheetsLimits.MaxCells)
                {
                    throw new FormatException(Text.T("Sheets.Invalid.TooManyCells"));
                }

                var (value, format) = Value(fields[c], decimalComma);
                var cell = new JsonObject { ["v"] = value };
                if (format is not null)
                {
                    if (!styleOf.TryGetValue(format, out var s))
                    {
                        s = styles.Count;
                        styles.Add(new JsonObject { ["nf"] = format });
                        styleOf[format] = s;
                    }

                    cell["s"] = s;
                }

                cells[A1.Cell(r + 1, c + 1)] = cell;
            }
        }

        return new JsonObject
        {
            ["version"] = 1,
            ["active"] = 0,
            ["styles"] = styles,
            ["sheets"] = new JsonArray(new JsonObject { ["id"] = "s1", ["name"] = Text.T("Sheets.SheetName", 1), ["cells"] = cells }),
        }.ToJsonString();
    }

    /// <summary>The text of the file: a BOM decides (UTF-8, UTF-16); else UTF-8 when the bytes are valid UTF-8, else Windows-1252.</summary>
    public static string Decode(ReadOnlySpan<byte> bytes)
    {
        if (bytes.StartsWith((ReadOnlySpan<byte>)[0xEF, 0xBB, 0xBF]))
        {
            return Encoding.UTF8.GetString(bytes[3..]);
        }

        if (bytes.StartsWith((ReadOnlySpan<byte>)[0xFF, 0xFE]))
        {
            return Encoding.Unicode.GetString(bytes[2..]);
        }

        if (bytes.StartsWith((ReadOnlySpan<byte>)[0xFE, 0xFF]))
        {
            return Encoding.BigEndianUnicode.GetString(bytes[2..]);
        }

        try
        {
            return new UTF8Encoding(false, true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            var windows = CodePagesEncodingProvider.Instance.GetEncoding(1252) ?? Encoding.Latin1;
            return windows.GetString(bytes);
        }
    }

    /// <summary>The delimiter that gives the first lines the same number of columns (more than one), counted outside quotes.</summary>
    public static char Detect(string text)
    {
        var counts = new List<int[]>();
        var current = new int[Delimiters.Length];
        var quoted = false;
        for (var i = 0; i < text.Length && counts.Count < 50; i++)
        {
            var c = text[i];
            if (c == '"')
            {
                quoted = !quoted;
            }
            else if (!quoted && c is '\n' or '\r')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }

                counts.Add(current);
                current = new int[Delimiters.Length];
            }
            else if (!quoted && Delimiters.IndexOf(c) is >= 0 and var d)
            {
                current[d]++;
            }
        }

        if (current.Any(n => n > 0))
        {
            counts.Add(current);
        }

        var best = ',';
        var bestScore = 0.0;
        for (var d = 0; d < Delimiters.Length; d++)
        {
            var perLine = counts.Select(line => line[d]).ToList();
            if (perLine.Count == 0)
            {
                continue;
            }

            var mode = perLine.GroupBy(n => n).OrderByDescending(g => g.Count()).ThenByDescending(g => g.Key).First();
            if (mode.Key == 0)
            {
                continue;
            }

            var score = ((double)mode.Count() / perLine.Count * 1000) + Math.Min(mode.Key, 999);
            if (score > bestScore)
            {
                best = Delimiters[d];
                bestScore = score;
            }
        }

        return best;
    }

    /// <summary>The records of the text (RFC 4180: quoted fields may hold delimiters, line breaks and doubled quotes).</summary>
    public static List<List<string>> Parse(string text, char delimiter)
    {
        var records = new List<List<string>>();
        var fields = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        var hasContent = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (quoted)
            {
                if (c != '"')
                {
                    field.Append(c);
                }
                else if (i + 1 < text.Length && text[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else
                {
                    quoted = false;
                }

                continue;
            }

            if (c == '"' && field.Length == 0)
            {
                quoted = true;
                hasContent = true;
            }
            else if (c == delimiter)
            {
                fields.Add(field.ToString());
                field.Clear();
                hasContent = true;
            }
            else if (c is '\r' or '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }

                fields.Add(field.ToString());
                field.Clear();
                records.Add(fields);
                fields = [];
                hasContent = false;
            }
            else
            {
                field.Append(c);
                hasContent = true;
            }
        }

        if (hasContent || field.Length > 0)
        {
            fields.Add(field.ToString());
            records.Add(fields);
        }

        return records;
    }

    /// <summary>Whether numbers use a decimal comma: with <c>;</c> unless dot decimals dominate, else only when comma decimals dominate.</summary>
    private static bool DecimalComma(List<List<string>> records, char delimiter)
    {
        int comma = 0, dot = 0;
        foreach (var field in records.Take(1000).SelectMany(r => r))
        {
            var t = field.Trim().TrimEnd('%').TrimEnd();
            if (CommaDecimal().IsMatch(t))
            {
                comma++;
            }
            else if (DotDecimal().IsMatch(t))
            {
                dot++;
            }
        }

        return delimiter == ';' ? comma >= dot : comma > dot;
    }

    /// <summary>The value of a field and the number format it brings (dates, times, percentages, thousands separators).</summary>
    public static (JsonNode Value, string? Format) Value(string field, bool decimalComma)
    {
        if (field.Length >= 2 && field[0] == '\'' && field[1] is '=' or '+' or '-' or '@' or '\t' or '\r')
        {
            return (Clean(field[1..], SheetsLimits.MaxText), null);
        }

        var t = field.Trim();
        if (t.Length is > 0 and <= 40)
        {
            if (t.ToUpperInvariant() is "TRUE" or "WAHR")
            {
                return (true, null);
            }

            if (t.ToUpperInvariant() is "FALSE" or "FALSCH")
            {
                return (false, null);
            }

            if (Number(t, decimalComma) is { } number)
            {
                return (number.Value, number.Format);
            }

            if (DateOrTime(t, decimalComma) is { } date)
            {
                return (date.Value, date.Format);
            }
        }

        return (Clean(field, SheetsLimits.MaxText), null);
    }

    private static (double Value, string? Format)? Number(string t, bool decimalComma)
    {
        var percent = t.EndsWith('%');
        if (percent)
        {
            t = t[..^1].TrimEnd();
        }

        var match = (decimalComma ? CommaNumber() : DotNumber()).Match(t);
        if (!match.Success)
        {
            return null;
        }

        var integer = match.Groups["int"].Value;
        var thousands = decimalComma ? '.' : ',';
        var digits = integer.Replace(thousands.ToString(), string.Empty, StringComparison.Ordinal);
        var decimals = match.Groups["dec"].Value;
        if ((digits.Length > 1 && digits[0] == '0') || digits.Length + decimals.Length > 15 || digits.Length + decimals.Length == 0)
        {
            return null; // 00123, phone numbers, IBANs … stay text
        }

        var invariant = match.Groups["sign"].Value + (digits.Length > 0 ? digits : "0") + (decimals.Length > 0 ? "." + decimals : string.Empty) + match.Groups["exp"].Value;
        if (!double.TryParse(invariant, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || !double.IsFinite(value))
        {
            return null;
        }

        var zeros = decimals.Length > 0 ? "." + new string('0', Math.Min(decimals.Length, 10)) : string.Empty;
        if (percent)
        {
            return (value / 100, "0" + zeros + "%");
        }

        return (value, digits.Length != integer.Length ? "#,##0" + zeros : null);
    }

    private static (double Value, string Format)? DateOrTime(string t, bool decimalComma)
    {
        var m = IsoDate().Match(t);
        string? format = null;
        int year = 0, month = 0, day = 0;
        if (m.Success)
        {
            (year, month, day, format) = (Parse(m, "y"), Parse(m, "m"), Parse(m, "d"), "yyyy-mm-dd");
        }
        else if ((m = GermanDate().Match(t)).Success)
        {
            (year, month, day, format) = (Parse(m, "y"), Parse(m, "m"), Parse(m, "d"), "dd.mm.yyyy");
        }
        else if (!decimalComma && (m = UsDate().Match(t)).Success)
        {
            (year, month, day, format) = (Parse(m, "y"), Parse(m, "m"), Parse(m, "d"), "m/d/yyyy");
        }
        else if ((m = TimeOnly().Match(t)).Success)
        {
            var hours = Parse(m, "h");
            var minutes = Parse(m, "min");
            var seconds = m.Groups["s"].Success ? Parse(m, "s") : 0;
            if (hours > 23 || minutes > 59 || seconds > 59)
            {
                return null;
            }

            return (((hours * 3600) + (minutes * 60) + seconds) / 86400.0, m.Groups["s"].Success ? "h:mm:ss" : "h:mm");
        }

        if (format is null || year is < 1900 or > 9999 || month is < 1 or > 12 || day < 1 || day > DateTime.DaysInMonth(year, month))
        {
            return null;
        }

        var date = new DateTime(year, month, day);
        if (m.Groups["h"].Success)
        {
            var hours = Parse(m, "h");
            var minutes = Parse(m, "min");
            var seconds = m.Groups["s"].Success ? Parse(m, "s") : 0;
            if (hours > 23 || minutes > 59 || seconds > 59)
            {
                return null;
            }

            date = date.Add(new TimeSpan(hours, minutes, seconds));
            format += m.Groups["s"].Success ? " hh:mm:ss" : " hh:mm";
        }

        return NumberFormats.ToSerial(date) is { } serial ? (serial, format) : null;
    }

    private static int Parse(Match m, string group) => int.Parse(m.Groups[group].Value, CultureInfo.InvariantCulture);

    // ---- writing ---------------------------------------------------------------------------------------------------------------------

    /// <summary>The active sheet as CSV (<c>,</c>) or TSV (tab): UTF-8 with BOM, CRLF.</summary>
    public static byte[] Write(JsonObject content, bool tabs)
    {
        var delimiter = tabs ? '\t' : ',';
        var sheets = content["sheets"]!.AsArray();
        var sheet = sheets[Math.Clamp(Int(content["active"], 0, sheets.Count - 1) ?? 0, 0, sheets.Count - 1)]!.AsObject();
        var styles = content["styles"] as JsonArray ?? [];
        var rowStyles = Styles(sheet["rows"] as JsonObject, key => A1.TryParseRow(key, out var r) ? r : 0);
        var colStyles = Styles(sheet["cols"] as JsonObject, key => A1.ColumnIndex(key));

        var values = new List<(int Row, int Col, string Text)>();
        foreach (var (address, node) in sheet["cells"] as JsonObject ?? [])
        {
            if (node is not JsonObject cell || !A1.TryParseCell(address, out var r, out var c))
            {
                continue;
            }

            var style = Int(cell["s"], 0, styles.Count - 1) ?? (rowStyles.TryGetValue(r, out var rs) ? rs : colStyles.GetValueOrDefault(c));
            var nf = style > 0 && styles[style] is JsonObject s ? Str(s["nf"]) : null;
            if (Field(cell, nf) is { Length: > 0 } text)
            {
                values.Add((r, c, Quote(text, delimiter)));
            }
        }

        values.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Col.CompareTo(b.Col));
        var columns = values.Count == 0 ? 0 : values.Max(v => v.Col);
        var sb = new StringBuilder();
        var index = 0;
        var lastRow = values.Count == 0 ? 0 : values[^1].Row;
        for (var r = 1; r <= lastRow; r++)
        {
            var c = 1;
            while (index < values.Count && values[index].Row == r)
            {
                for (; c < values[index].Col; c++)
                {
                    sb.Append(delimiter);
                }

                sb.Append(values[index].Text);
                index++;
            }

            for (; c < columns; c++)
            {
                sb.Append(delimiter);
            }

            sb.Append("\r\n");
        }

        var preamble = Encoding.UTF8.GetPreamble();
        var body = Encoding.UTF8.GetBytes(sb.ToString());
        return [.. preamble, .. body];
    }

    private static Dictionary<int, int> Styles(JsonObject? lines, Func<string, int> index)
    {
        var result = new Dictionary<int, int>();
        foreach (var (key, value) in lines ?? [])
        {
            if (index(key) is > 0 and var i && value is JsonObject line && Int(line["s"], 1, int.MaxValue) is { } s)
            {
                result[i] = s;
            }
        }

        return result;
    }

    /// <summary>The text of a cell in a file: numbers invariant (15 digits), dates ISO, booleans TRUE/FALSE, errors as their code, formulas as their last value.</summary>
    public static string? Field(JsonObject cell, string? nf)
    {
        var v = cell["v"];
        if (Str(v) is { } text)
        {
            return Str(cell["t"]) == "e" || text.Length == 0 || text[0] is not ('=' or '+' or '-' or '@' or '\t' or '\r') ? text : "'" + text;
        }

        if (Bool(v) is { } flag)
        {
            return flag ? "TRUE" : "FALSE";
        }

        if (Num(v) is not { } number)
        {
            return null;
        }

        var kind = NumberFormats.Classify(nf);
        if (kind != DateKind.None && NumberFormats.FromSerial(number) is { } date)
        {
            return kind == DateKind.Time && number < 1
                ? date.ToString("HH:mm:ss", CultureInfo.InvariantCulture)
                : date.TimeOfDay == TimeSpan.Zero ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        return number.ToString("G15", CultureInfo.InvariantCulture);
    }

    private static string Quote(string text, char delimiter)
        => text.IndexOfAny([delimiter, '"', '\r', '\n']) < 0 ? text : "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";

    // ---- patterns --------------------------------------------------------------------------------------------------------------------

    [GeneratedRegex(@"^[+-]?(\d{1,3}(\.\d{3})+(,\d+)?|\d+,(\d{1,2}|\d{4,}))$")]
    private static partial Regex CommaDecimal();

    [GeneratedRegex(@"^[+-]?(\d{1,3}(,\d{3})+(\.\d+)?|\d+\.(\d{1,2}|\d{4,}))$")]
    private static partial Regex DotDecimal();

    [GeneratedRegex(@"^(?<sign>[+-]?)(?<int>\d{1,3}(\.\d{3})+|\d*)(,(?<dec>\d+))?(?<exp>[eE][+-]?\d{1,3})?$")]
    private static partial Regex CommaNumber();

    [GeneratedRegex(@"^(?<sign>[+-]?)(?<int>\d{1,3}(,\d{3})+|\d*)(\.(?<dec>\d+))?(?<exp>[eE][+-]?\d{1,3})?$")]
    private static partial Regex DotNumber();

    [GeneratedRegex(@"^(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})(?:[T ](?<h>\d{1,2}):(?<min>\d{2})(?::(?<s>\d{2}))?)?$")]
    private static partial Regex IsoDate();

    [GeneratedRegex(@"^(?<d>\d{1,2})\.(?<m>\d{1,2})\.(?<y>\d{4})(?: (?<h>\d{1,2}):(?<min>\d{2})(?::(?<s>\d{2}))?)?$")]
    private static partial Regex GermanDate();

    [GeneratedRegex(@"^(?<m>\d{1,2})/(?<d>\d{1,2})/(?<y>\d{4})(?: (?<h>\d{1,2}):(?<min>\d{2})(?::(?<s>\d{2}))?)?$")]
    private static partial Regex UsDate();

    [GeneratedRegex(@"^(?<h>\d{1,2}):(?<min>\d{2})(?::(?<s>\d{2}))?$")]
    private static partial Regex TimeOnly();
}
