using System.Text;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>What a number format shows of a number that is a date (Excel serial).</summary>
internal enum DateKind
{
    None,
    Date,
    Time,
    DateTime,
}

/// <summary>Excel number formats: the codes of the built-in ids, whether a code shows a date, and the serial numbers of dates.</summary>
internal static class NumberFormats
{
    /// <summary>The built-in formats of Excel (ids 14 and 22 follow the language: they show the short date of the system).</summary>
    public static string? BuiltIn(int id, bool german) => id switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "#,##0_);(#,##0)",
        6 => "#,##0_);[Red](#,##0)",
        7 => "#,##0.00_);(#,##0.00)",
        8 => "#,##0.00_);[Red](#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 or (>= 27 and <= 36) or (>= 50 and <= 58) => german ? "dd.mm.yyyy" : "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => german ? "dd.mm.yyyy hh:mm" : "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };

    /// <summary>Whether a format code shows a date, a time or both (texts in quotes, escaped characters and [colours] do not count).</summary>
    public static DateKind Classify(string? code)
    {
        if (string.IsNullOrEmpty(code))
        {
            return DateKind.None;
        }

        var section = code.Split(';')[0];
        var plain = new StringBuilder(section.Length);
        var time = false;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            switch (c)
            {
                case '"':
                    var end = section.IndexOf('"', i + 1);
                    i = end < 0 ? section.Length : end;
                    break;
                case '\\' or '_' or '*':
                    i++;
                    break;
                case '[':
                    var close = section.IndexOf(']', i + 1);
                    var inner = close < 0 ? string.Empty : section[(i + 1)..close].ToLowerInvariant();
                    if (inner.Length > 0 && inner.All(ch => ch is 'h' or 'm' or 's'))
                    {
                        time = true;
                    }

                    i = close < 0 ? section.Length : close;
                    break;
                default:
                    plain.Append(char.ToLowerInvariant(c));
                    break;
            }
        }

        var text = plain.ToString();
        if (text.Contains("am/pm", StringComparison.Ordinal) || text.Contains("a/p", StringComparison.Ordinal))
        {
            time = true;
            text = text.Replace("am/pm", string.Empty, StringComparison.Ordinal).Replace("a/p", string.Empty, StringComparison.Ordinal);
        }

        if (text.Contains("general", StringComparison.Ordinal) || text.Contains('e') && !text.Contains('d') && !text.Contains('y'))
        {
            return time ? DateKind.Time : DateKind.None;
        }

        var date = text.Contains('d') || text.Contains('y');
        if (text.Contains('h') || text.Contains('s'))
        {
            time = true;
        }

        // m alone is a month (mmm-yy) unless it stands next to h or s (h:mm, mm:ss)
        if (!date && text.Contains('m') && !time)
        {
            date = true;
        }

        return (date, time) switch
        {
            (true, true) => DateKind.DateTime,
            (true, false) => DateKind.Date,
            (false, true) => DateKind.Time,
            _ => DateKind.None,
        };
    }

    /// <summary>The Excel serial number of a date (1 = 1900-01-01, with Excel's 29 February 1900), or null before 1900.</summary>
    public static double? ToSerial(DateTime date)
    {
        if (date.Year < 1900)
        {
            return null;
        }

        var serial = date.ToOADate();
        return serial < 61 ? serial - 1 : serial;
    }

    /// <summary>The date of an Excel serial number, or null when it is not one (negative, too large).</summary>
    public static DateTime? FromSerial(double serial)
    {
        if (!double.IsFinite(serial) || serial < 0 || serial >= 2_958_466)
        {
            return null;
        }

        var oa = serial < 61 ? serial + 1 : serial;
        if (serial is >= 60 and < 61)
        {
            oa = 61; // Excel's 29 February 1900 does not exist: 1 March
        }

        var date = DateTime.FromOADate(oa);
        // round to whole seconds (serials carry floating-point noise)
        return new DateTime((date.Ticks + (TimeSpan.TicksPerSecond / 2)) / TimeSpan.TicksPerSecond * TimeSpan.TicksPerSecond, DateTimeKind.Unspecified);
    }
}
