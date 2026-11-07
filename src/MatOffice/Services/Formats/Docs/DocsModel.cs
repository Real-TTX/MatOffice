using System.Globalization;
using System.Text;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The parsed form of Docs HTML: a flat list of blocks, each paragraph a list of inline items with their formatting. The canonicalizer
/// (<see cref="DocsHtml"/>) builds it from any HTML and writes it back as canonical Docs HTML; the converters read and write it.
/// </summary>
internal abstract class DocBlock;

/// <summary>A paragraph-like block: <c>p</c>, <c>h1</c> … <c>h6</c>, <c>blockquote</c>, <c>pre</c>.</summary>
internal sealed class DocParagraph : DocBlock
{
    public DocParagraph(string tag = "p") => Tag = tag;

    /// <summary><c>p</c>, <c>h1</c> … <c>h6</c>, <c>blockquote</c> or <c>pre</c>.</summary>
    public string Tag { get; set; }

    /// <summary>The <c>data-style</c> (another paragraph style), or null.</summary>
    public string? Style { get; set; }

    /// <summary>The kind of list (<c>data-list</c>), or null when the paragraph is not a list item.</summary>
    public string? List { get; set; }

    /// <summary>The list level 0 … 8.</summary>
    public int Level { get; set; }

    /// <summary>The number the list starts anew with (<c>data-start</c>).</summary>
    public int? Start { get; set; }

    public string? Id { get; set; }

    /// <summary>The paragraph formatting (alignment, spacing, indents, shading, borders).</summary>
    public CssList Css { get; set; } = new();

    public List<DocInline> Inlines { get; } = [];

    /// <summary>The heading level 1 … 6, or 0.</summary>
    public int HeadingLevel => Tag.Length == 2 && Tag[0] == 'h' && Tag[1] is >= '1' and <= '6' ? Tag[1] - '0' : 0;

    /// <summary>The key of the paragraph style (see <see cref="DocsStyles"/>).</summary>
    public string StyleKey => Tag switch
    {
        "blockquote" => Style == "IntenseQuote" ? "IntenseQuote" : "Quote",
        "pre" => "Code",
        "p" => Style ?? (List is null ? "Normal" : "ListParagraph"),
        _ => HeadingLevel > 0 ? "Heading" + HeadingLevel : "Normal",
    };

    /// <summary>A copy of the block attributes without the content.</summary>
    public DocParagraph CloneShell() => new(Tag) { Style = Style, List = List, Level = Level, Start = Start, Css = Css.Clone() };

    /// <summary>The text of the paragraph (fields and pictures left out, line breaks as \n).</summary>
    public string PlainText()
    {
        var sb = new StringBuilder();
        foreach (var inline in Inlines)
        {
            switch (inline)
            {
                case DocText t:
                    sb.Append(t.Text);
                    break;
                case DocLineBreak:
                    sb.Append('\n');
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>The inline items without a single line break at the end (a browser does not show it either).</summary>
    public IEnumerable<DocInline> VisibleInlines()
        => Inlines.Count > 1 && Inlines[^1] is DocLineBreak ? Inlines.Take(Inlines.Count - 1) : Inlines;
}

internal sealed class DocTable : DocBlock
{
    /// <summary>The <c>data-table-style</c>: grid, plain, light, striped, accent1, accent2, accent6.</summary>
    public string TableStyle { get; set; } = "grid";

    /// <summary>The CSS width of the table (<c>100%</c>), or null.</summary>
    public string? Width { get; set; }

    /// <summary>The widths of the columns (<c>30%</c>, or null for "auto").</summary>
    public List<string?> Columns { get; } = [];

    public List<DocRow> Rows { get; } = [];

    /// <summary>The number of grid columns (the widest row, spans counted, rowspans of earlier rows included).</summary>
    public int ColumnCount()
    {
        var max = Columns.Count;
        foreach (var (row, cells) in Layout())
        {
            _ = row;
            foreach (var (_, column, cell) in cells)
            {
                max = Math.Max(max, column + cell.ColSpan);
            }
        }

        return Math.Max(1, max);
    }

    /// <summary>Places every cell on the grid: for each row its cells with their first column (cells spanning rows from above are skipped).</summary>
    public IEnumerable<(DocRow Row, List<(int Index, int Column, DocCell Cell)> Cells)> Layout()
    {
        var taken = new Dictionary<(int Row, int Column), bool>();
        for (var r = 0; r < Rows.Count; r++)
        {
            var row = Rows[r];
            var column = 0;
            var placed = new List<(int, int, DocCell)>();
            for (var i = 0; i < row.Cells.Count; i++)
            {
                while (taken.ContainsKey((r, column)))
                {
                    column++;
                }

                var cell = row.Cells[i];
                placed.Add((i, column, cell));
                for (var dr = 0; dr < cell.RowSpan && r + dr < Rows.Count; dr++)
                {
                    for (var dc = 0; dc < cell.ColSpan; dc++)
                    {
                        taken[(r + dr, column + dc)] = true;
                    }
                }

                column += cell.ColSpan;
            }

            yield return (row, placed);
        }
    }
}

internal sealed class DocRow
{
    /// <summary>A header row (<c>thead</c>, repeated on every page).</summary>
    public bool Header { get; set; }

    public List<DocCell> Cells { get; } = [];
}

internal sealed class DocCell
{
    public bool HeaderCell { get; set; }

    public int ColSpan { get; set; } = 1;

    public int RowSpan { get; set; } = 1;

    /// <summary>background-color, vertical-align, border-*.</summary>
    public CssList Css { get; set; } = new();

    public List<DocBlock> Blocks { get; } = [];
}

/// <summary>A horizontal line (<c>hr</c>).</summary>
internal sealed class DocRule : DocBlock;

/// <summary>A manual page or column break (<c>div data-break</c>).</summary>
internal sealed class DocBreak(string kind) : DocBlock
{
    public string Kind { get; } = kind;
}

internal abstract class DocInline
{
    public RunFormat Format { get; set; } = RunFormat.Plain;
}

internal sealed class DocText(string text) : DocInline
{
    public string Text { get; set; } = text;
}

internal sealed class DocLineBreak : DocInline;

internal sealed class DocImage : DocInline
{
    /// <summary>A <c>data:image/(png|jpeg|gif|webp);base64,</c> address.</summary>
    public string Src { get; set; } = string.Empty;

    public string Alt { get; set; } = string.Empty;

    /// <summary>Width and height in CSS pixels, or null (the size of the picture).</summary>
    public double? Width { get; set; }

    public double? Height { get; set; }

    /// <summary><c>left</c>, <c>right</c>, <c>center</c> or null (inline with the text).</summary>
    public string? Align { get; set; }
}

/// <summary>The reference to a note (<c>sup data-fn</c>).</summary>
internal sealed class DocNoteRef(string id) : DocInline
{
    public string Id { get; } = id;
}

/// <summary>A field (<c>span data-field</c>): page, pages, date, time, title.</summary>
internal sealed class DocField(string kind) : DocInline
{
    public string Kind { get; } = kind;
}

/// <summary>
/// The formatting of inline content, in the canonical nesting of Docs HTML:
/// <c>a</c> › <c>span[data-comment]</c> › <c>span[style]</c> › <c>strong</c> › <c>em</c> › <c>u</c> › <c>s</c> › <c>sub</c>/<c>sup</c>.
/// </summary>
internal sealed record RunFormat
{
    public static readonly RunFormat Plain = new();

    public string? Href { get; init; }

    /// <summary>The ids of the comments over this text, separated by spaces.</summary>
    public string? Comment { get; init; }

    public string? FontFamily { get; init; }

    public string? FontSize { get; init; }

    public string? Color { get; init; }

    public string? Background { get; init; }

    public string? LetterSpacing { get; init; }

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool Underline { get; init; }

    public bool Strike { get; init; }

    /// <summary>1 = superscript, -1 = subscript, 0 = neither.</summary>
    public int Script { get; init; }

    public bool HasSpanStyle => FontFamily is not null || FontSize is not null || Color is not null || Background is not null || LetterSpacing is not null;

    /// <summary>The CSS of the <c>span[style]</c>.</summary>
    public string SpanStyle()
    {
        var css = new CssList();
        css.Set("font-family", FontFamily);
        css.Set("font-size", FontSize);
        css.Set("color", Color);
        css.Set("background-color", Background);
        css.Set("letter-spacing", LetterSpacing);
        return css.ToString();
    }

    /// <summary>Adds the ids of a comment span to the ones already open.</summary>
    public RunFormat WithComment(string ids)
    {
        if (Comment is null)
        {
            return this with { Comment = ids };
        }

        var all = Comment.Split(' ').Concat(ids.Split(' ')).Distinct(StringComparer.Ordinal);
        return this with { Comment = string.Join(' ', all) };
    }
}

/// <summary>An ordered list of CSS declarations (property names in lower case).</summary>
internal sealed class CssList
{
    private readonly List<KeyValuePair<string, string>> _items = [];

    public int Count => _items.Count;

    public IEnumerable<KeyValuePair<string, string>> Items => _items;

    public string? this[string name] => Get(name);

    public string? Get(string name)
    {
        foreach (var item in _items)
        {
            if (item.Key == name)
            {
                return item.Value;
            }
        }

        return null;
    }

    public bool Has(string name) => Get(name) is not null;

    /// <summary>Sets (or with null removes) a declaration; a new one goes to the end.</summary>
    public void Set(string name, string? value)
    {
        var index = _items.FindIndex(i => i.Key == name);
        if (value is null)
        {
            if (index >= 0)
            {
                _items.RemoveAt(index);
            }

            return;
        }

        if (index >= 0)
        {
            _items[index] = new(name, value);
        }
        else
        {
            _items.Add(new(name, value));
        }
    }

    public void Remove(string name) => Set(name, null);

    /// <summary>Sets a declaration only if it is not there yet.</summary>
    public void SetDefault(string name, string? value)
    {
        if (value is not null && !Has(name))
        {
            _items.Add(new(name, value));
        }
    }

    public CssList Clone()
    {
        var copy = new CssList();
        copy._items.AddRange(_items);
        return copy;
    }

    public override string ToString()
    {
        if (_items.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var (name, value) in _items)
        {
            if (sb.Length > 0)
            {
                sb.Append("; ");
            }

            sb.Append(name).Append(": ").Append(value);
        }

        return sb.ToString();
    }

    /// <summary>A length in points written the canonical way (<c>12pt</c>, <c>4.5pt</c>).</summary>
    public static string Pt(double value) => Num(value) + "pt";

    /// <summary>A number with at most two decimals, invariant.</summary>
    public static string Num(double value) => Math.Round(value, 2, MidpointRounding.AwayFromZero).ToString("0.##", CultureInfo.InvariantCulture);
}
