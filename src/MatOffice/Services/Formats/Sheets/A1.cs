using System.Globalization;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>A rectangle of cells (1-based, inclusive), written as <c>A1:C3</c> (or <c>A1</c> for one cell).</summary>
internal readonly record struct Area(int Row1, int Col1, int Row2, int Col2)
{
    public bool IsCell => Row1 == Row2 && Col1 == Col2;

    public int Rows => Row2 - Row1 + 1;

    public int Cols => Col2 - Col1 + 1;

    public bool Overlaps(Area other) => Row1 <= other.Row2 && other.Row1 <= Row2 && Col1 <= other.Col2 && other.Col1 <= Col2;

    public bool Contains(int row, int col) => row >= Row1 && row <= Row2 && col >= Col1 && col <= Col2;

    public override string ToString() => IsCell ? A1.Cell(Row1, Col1) : A1.Cell(Row1, Col1) + ":" + A1.Cell(Row2, Col2);
}

/// <summary>Cell addresses in A1 style: <c>A1</c> … <c>XFD1048576</c>, without <c>$</c>.</summary>
internal static class A1
{
    public const int MaxRows = 1_048_576;

    public const int MaxCols = 16_384;

    /// <summary>The letters of a column (1 → A, 27 → AA).</summary>
    public static string Column(int col)
    {
        Span<char> buffer = stackalloc char[3];
        var i = buffer.Length;
        while (col > 0)
        {
            var rest = (col - 1) % 26;
            buffer[--i] = (char)('A' + rest);
            col = (col - 1) / 26;
        }

        return new string(buffer[i..]);
    }

    /// <summary>The number of a column from its letters (case-insensitive), 0 when they are not a column.</summary>
    public static int ColumnIndex(ReadOnlySpan<char> letters)
    {
        if (letters.Length is 0 or > 3)
        {
            return 0;
        }

        var col = 0;
        foreach (var c in letters)
        {
            var u = char.ToUpperInvariant(c);
            if (u is < 'A' or > 'Z')
            {
                return 0;
            }

            col = (col * 26) + (u - 'A' + 1);
        }

        return col <= MaxCols ? col : 0;
    }

    public static string Cell(int row, int col) => Column(col) + row.ToString(CultureInfo.InvariantCulture);

    /// <summary>Reads <c>B12</c> (case-insensitive; with <paramref name="allowDollar"/> also <c>$B$12</c>).</summary>
    public static bool TryParseCell(ReadOnlySpan<char> text, out int row, out int col, bool allowDollar = false)
    {
        row = 0;
        col = 0;
        var i = 0;
        if (allowDollar && i < text.Length && text[i] == '$')
        {
            i++;
        }

        var start = i;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            i++;
        }

        col = ColumnIndex(text[start..i]);
        if (col == 0)
        {
            return false;
        }

        if (allowDollar && i < text.Length && text[i] == '$')
        {
            i++;
        }

        return TryParseRow(text[i..], out row);
    }

    /// <summary>Reads a row number 1 … 1048576 (no sign, no leading zero).</summary>
    public static bool TryParseRow(ReadOnlySpan<char> text, out int row)
    {
        row = 0;
        if (text.Length is 0 or > 7 || text[0] == '0')
        {
            return false;
        }

        foreach (var c in text)
        {
            if (!char.IsAsciiDigit(c))
            {
                return false;
            }

            row = (row * 10) + (c - '0');
        }

        return row <= MaxRows;
    }

    /// <summary>Reads <c>A1:C3</c> or <c>A1</c> (corners in any order) into a normalized area.</summary>
    public static bool TryParseArea(ReadOnlySpan<char> text, out Area area, bool allowDollar = false)
    {
        area = default;
        var colon = text.IndexOf(':');
        if (colon < 0)
        {
            if (!TryParseCell(text, out var r, out var c, allowDollar))
            {
                return false;
            }

            area = new Area(r, c, r, c);
            return true;
        }

        if (!TryParseCell(text[..colon], out var r1, out var c1, allowDollar) || !TryParseCell(text[(colon + 1)..], out var r2, out var c2, allowDollar))
        {
            return false;
        }

        area = new Area(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        return true;
    }

    /// <summary>The canonical form of a cell address (<c>b2</c> → <c>B2</c>), or null.</summary>
    public static string? NormalizeCell(string? text)
        => text is not null && TryParseCell(text, out var r, out var c) ? Cell(r, c) : null;

    /// <summary>The canonical form of a range (<c>c3:a1</c> → <c>A1:C3</c>), or null.</summary>
    public static string? NormalizeArea(string? text)
        => text is not null && TryParseArea(text.Trim(), out var area, allowDollar: true) ? area.ToString() : null;

    /// <summary>
    /// A reference to cells, maybe on another sheet: <c>A1:C6</c>, <c>Sheet2!A1:C6</c>, <c>'My sheet'!A1</c>. Returns the canonical form
    /// (no <c>$</c> in the cell part), or null.
    /// </summary>
    public static string? NormalizeReference(string? text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length > 300)
        {
            return null;
        }

        text = text.Trim();
        var bang = text.LastIndexOf('!');
        if (bang < 0)
        {
            return NormalizeArea(text);
        }

        var sheet = text[..bang];
        var area = NormalizeArea(text[(bang + 1)..]);
        if (area is null || sheet.Length == 0)
        {
            return null;
        }

        var quoted = sheet.Length >= 2 && sheet[0] == '\'' && sheet[^1] == '\'';
        var name = quoted ? sheet[1..^1].Replace("''", "'", StringComparison.Ordinal) : sheet;
        if (name.Length is 0 or > 31 || name.AsSpan().IndexOfAny(":\\/?*[]") >= 0 || name.Any(char.IsControl) || (!quoted && name.Any(c => !char.IsLetterOrDigit(c) && c is not '_' and not '.')))
        {
            return null;
        }

        return sheet + "!" + area;
    }
}
