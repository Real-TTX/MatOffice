using System.Text;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// Plain text (<c>.txt</c>). Export: one line per paragraph (list items with their marker, indented by level), table rows as tab-separated cells,
/// notes at the end; UTF-8 without BOM, CRLF (Windows editors). Import: every line a paragraph (UTF-8, UTF-16 with BOM, else Windows-1252).
/// </summary>
internal static class TextFormat
{
    private const string NewLine = "\r\n";

    public static DocsDocument Read(Stream input)
    {
        var text = DocsConvert.Decode(DocsConvert.ReadAll(input));
        var document = new DocsDocument();
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        var count = lines.Length;
        if (count > 0 && lines[^1].Length == 0)
        {
            count--;
        }

        for (var i = 0; i < count; i++)
        {
            var pages = lines[i].Split('\f');
            for (var j = 0; j < pages.Length; j++)
            {
                if (j > 0)
                {
                    document.Body.Add(new DocBreak("page"));
                }

                if (pages.Length > 1 && pages[j].Length == 0)
                {
                    continue;
                }

                var p = new DocParagraph();
                var line = DocsHtml.CleanText(pages[j]);
                if (line.Length > 0)
                {
                    p.Inlines.Add(new DocText(line));
                }

                document.Body.Add(p);
            }
        }

        if (document.Body.Count == 0)
        {
            document.Body.Add(new DocParagraph());
        }

        return document;
    }

    public static byte[] Write(DocsDocument document)
    {
        var sb = new StringBuilder();
        var notes = DocsConvert.NoteNumbers(document);
        WriteBlocks(sb, document.Body, notes, document.Language);
        var ordered = DocsConvert.OrderedNotes(document);
        if (ordered.Count > 0)
        {
            sb.Append(NewLine).Append("____").Append(NewLine);
            foreach (var (note, number) in ordered)
            {
                var text = new StringBuilder();
                WriteBlocks(text, note.Blocks, notes, document.Language);
                sb.Append('[').Append(number).Append("] ").Append(text.ToString().TrimEnd()).Append(NewLine);
            }
        }

        return new UTF8Encoding(false).GetBytes(sb.ToString());
    }

    private static void WriteBlocks(StringBuilder sb, IEnumerable<DocBlock> blocks, Dictionary<string, int> notes, string language)
    {
        var numbering = new ListNumbering();
        foreach (var block in blocks)
        {
            if (block is not DocParagraph { List: not null })
            {
                numbering.Reset();
            }

            switch (block)
            {
                case DocParagraph p when p.List is not null:
                    var (number, path) = numbering.Next(p);
                    var indent = new string(' ', p.Level * 2);
                    var marker = DocsStyles.Marker(p.List, p.Level, number, path) + " ";
                    sb.Append(indent).Append(marker).Append(Inline(p, notes, language).Replace("\n", NewLine + indent + new string(' ', marker.Length), StringComparison.Ordinal)).Append(NewLine);
                    break;
                case DocParagraph p:
                    sb.Append(Inline(p, notes, language).Replace("\n", NewLine, StringComparison.Ordinal)).Append(NewLine);
                    break;
                case DocTable table:
                    foreach (var row in table.Rows)
                    {
                        sb.Append(string.Join('\t', row.Cells.Select(c => string.Join(' ', DocsHtml.Paragraphs(c.Blocks).Select(p => Inline(p, notes, language).Replace('\n', ' ').Replace('\t', ' '))))));
                        sb.Append(NewLine);
                    }

                    break;
                case DocRule:
                    sb.Append("----------").Append(NewLine);
                    break;
                case DocBreak { Kind: "page" }:
                    sb.Append(NewLine);
                    break;
            }
        }
    }

    /// <summary>The text of a paragraph: line breaks as \n, note references as [1], fields as their value.</summary>
    private static string Inline(DocParagraph p, Dictionary<string, int> notes, string language)
    {
        var sb = new StringBuilder();
        foreach (var inline in p.VisibleInlines())
        {
            switch (inline)
            {
                case DocText t:
                    sb.Append(t.Text);
                    break;
                case DocLineBreak:
                    sb.Append('\n');
                    break;
                case DocNoteRef r when notes.TryGetValue(r.Id, out var n):
                    sb.Append('[').Append(n).Append(']');
                    break;
                case DocField f:
                    sb.Append(DocsConvert.FieldText(f.Kind, string.Empty, language));
                    break;
            }
        }

        return sb.ToString();
    }
}
