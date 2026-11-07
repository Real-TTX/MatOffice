using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// Markdown (CommonMark with the GitHub tables, strikethrough and footnotes). Export: headings (Title as <c>#</c>), paragraphs, <c>**bold**</c>,
/// <c>*italic*</c>, <c>~~strike~~</c>, <c>&lt;u&gt;</c>/<c>&lt;sub&gt;</c>/<c>&lt;sup&gt;</c>, links, pictures as <c>data:</c> addresses, nested lists
/// with their numbers, quotes, fenced code, pipe tables, <c>---</c>, hard line breaks, notes as <c>[^1]</c>. Import: the same, and setext headings,
/// <c>*</c>/<c>+</c> bullets, <c>1)</c> numbers, indented code, autolinks, reference links, a few inline tags.
/// </summary>
internal static partial class MarkdownFormat
{
    public static byte[] Write(DocsDocument document) => new UTF8Encoding(false).GetBytes(ToMarkdown(document));

    public static string ToMarkdown(DocsDocument document)
    {
        var writer = new Writer(document);
        writer.Blocks(document.Body);
        writer.Notes();
        return writer.ToString();
    }

    public static DocsDocument Read(Stream input) => FromMarkdown(DocsConvert.Decode(DocsConvert.ReadAll(input)));

    public static DocsDocument FromMarkdown(string text)
    {
        var document = new DocsDocument();
        new Parser(document, text).Run();
        if (document.Body.Count == 0)
        {
            document.Body.Add(new DocParagraph());
        }

        LinksToHeadings(document);
        document.DropUnreferencedNotes();
        return document;
    }

    // ------------------------------------------------------------------ links inside the document

    /// <summary>The headings of Markdown: the headings of Docs and the title (written as <c>#</c>), in the order of the document.</summary>
    private static IEnumerable<DocParagraph> Headings(IEnumerable<DocBlock> blocks)
        => blocks.OfType<DocParagraph>().Where(p => p.HeadingLevel > 0 || p.Style == "Title");

    /// <summary>
    /// The anchors GitHub (and most Markdown viewers) give headings, which the links inside a Markdown file point to: the text in lower case
    /// without punctuation, spaces as hyphens, a number after a repeated one (<c>intro</c>, <c>intro-1</c>).
    /// </summary>
    private static Dictionary<DocParagraph, string> Anchors(IEnumerable<DocParagraph> headings)
    {
        var anchors = new Dictionary<DocParagraph, string>(ReferenceEqualityComparer.Instance);
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var heading in headings)
        {
            var sb = new StringBuilder();
            foreach (var c in heading.PlainText().Trim().ToLowerInvariant())
            {
                if (char.IsLetterOrDigit(c) || c is '-' or '_')
                {
                    sb.Append(c);
                }
                else if (c == ' ')
                {
                    sb.Append('-');
                }
            }

            var anchor = sb.ToString();
            for (var n = 1; !used.Add(anchor); n++)
            {
                anchor = sb + "-" + n.ToString(CultureInfo.InvariantCulture);
            }

            anchors[heading] = anchor;
        }

        return anchors;
    }

    /// <summary>Links to the anchor of a heading point to the heading (it gets an id); links to anything else inside the file are plain text.</summary>
    private static void LinksToHeadings(DocsDocument document)
    {
        var byAnchor = new Dictionary<string, DocParagraph>(StringComparer.Ordinal);
        foreach (var (heading, anchor) in Anchors(Headings(document.Body)))
        {
            byAnchor.TryAdd(anchor, heading);
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var p in DocsHtml.Paragraphs(document.Body.Concat(document.Notes.SelectMany(n => n.Blocks))))
        {
            foreach (var inline in p.Inlines)
            {
                if (inline.Format.Href is not ['#', ..] href)
                {
                    continue;
                }

                string? target = null;
                var anchor = Uri.UnescapeDataString(href[1..]).ToLowerInvariant();
                if (byAnchor.TryGetValue(anchor, out var heading))
                {
                    if (heading.Id is null)
                    {
                        var name = DocsHtml.IsId(anchor) ? anchor : DocsHtml.ToId(anchor);
                        var id = name;
                        for (var n = 2; !ids.Add(id); n++)
                        {
                            id = name[..Math.Min(name.Length, 58)] + "-" + n.ToString(CultureInfo.InvariantCulture);
                        }

                        heading.Id = id;
                    }

                    target = "#" + heading.Id;
                }

                inline.Format = inline.Format with { Href = target };
            }
        }
    }

    // ------------------------------------------------------------------ writer

    private sealed class Writer(DocsDocument document)
    {
        private readonly StringBuilder _sb = new();

        private readonly Dictionary<string, int> _notes = DocsConvert.NoteNumbers(document);

        /// <summary>The anchor of every heading with an id: the links to it point there.</summary>
        private readonly Dictionary<string, string> _anchors = Anchors(Headings(document.Body))
            .Where(a => a.Key.Id is not null).ToDictionary(a => a.Key.Id!, a => a.Value, StringComparer.Ordinal);

        public override string ToString() => _sb.ToString().TrimEnd() + "\n";

        public void Blocks(IReadOnlyList<DocBlock> blocks)
        {
            var numbering = new ListNumbering();
            var offsets = new int[9];
            var inList = false;
            foreach (var block in blocks)
            {
                if (block is DocParagraph { List: not null } item)
                {
                    if (!inList)
                    {
                        numbering.Reset();
                        Array.Clear(offsets);
                        inList = true;
                    }

                    ListItem(item, numbering, offsets);
                    continue;
                }

                if (inList)
                {
                    _sb.Append('\n');
                    inList = false;
                }

                numbering.Reset();
                Block(block);
            }
        }

        private void ListItem(DocParagraph p, ListNumbering numbering, int[] offsets)
        {
            var (number, _) = numbering.Next(p);
            var level = Math.Clamp(p.Level, 0, 8);
            for (var i = 0; i < level; i++)
            {
                if (offsets[i] == 0)
                {
                    offsets[i] = (i == 0 ? 0 : offsets[i - 1]) + 2;
                }
            }

            var indent = level == 0 ? 0 : offsets[level - 1];
            var marker = DocsStyles.NumberKinds.Contains(p.List!) ? number.ToString(CultureInfo.InvariantCulture) + "." : "-";
            offsets[level] = indent + marker.Length + 1;
            for (var i = level + 1; i < 9; i++)
            {
                offsets[i] = 0;
            }

            var continuation = new string(' ', offsets[level]);
            _sb.Append(' ', indent).Append(marker).Append(' ');
            if (p.HeadingLevel > 0)
            {
                _sb.Append('#', p.HeadingLevel).Append(' ').Append(Inline(p.VisibleInlines(), " "));
            }
            else
            {
                _sb.Append(EscapeStart(Inline(p.VisibleInlines(), "\\\n" + continuation)));
            }

            _sb.Append('\n');
        }

        private void Block(DocBlock block)
        {
            switch (block)
            {
                case DocParagraph p when p.HeadingLevel > 0 || p.Style == "Title":
                    _sb.Append('#', Math.Max(1, p.HeadingLevel)).Append(' ').Append(Inline(p.VisibleInlines(), " ")).Append("\n\n");
                    break;
                case DocParagraph { Tag: "blockquote" } p:
                    _sb.Append("> ").Append(EscapeStart(Inline(p.VisibleInlines(), "\\\n> "))).Append("\n\n");
                    break;
                case DocParagraph { Tag: "pre" } p:
                    var code = p.PlainText();
                    var fence = code.Contains("```", StringComparison.Ordinal) ? "~~~~" : "```";
                    _sb.Append(fence).Append('\n').Append(code).Append('\n').Append(fence).Append("\n\n");
                    break;
                case DocParagraph p:
                    var inlines = p.VisibleInlines().ToList();
                    if (inlines.Count == 0)
                    {
                        break;
                    }

                    if (p.Style == "Subtitle")
                    {
                        foreach (var inline in inlines)
                        {
                            inline.Format = inline.Format with { Italic = true };
                        }
                    }

                    var text = EscapeStart(Inline(inlines, "\\\n"));
                    if (text.Trim().Length > 0)
                    {
                        _sb.Append(text).Append("\n\n");
                    }

                    break;
                case DocTable table:
                    Table(table);
                    break;
                case DocRule:
                    _sb.Append("---\n\n");
                    break;
            }
        }

        private void Table(DocTable table)
        {
            var columns = table.ColumnCount();
            var grid = new List<string[]>();
            foreach (var (_, cells) in table.Layout())
            {
                var line = Enumerable.Repeat(string.Empty, columns).ToArray();
                foreach (var (_, column, cell) in cells)
                {
                    if (column < columns)
                    {
                        line[column] = string.Join("<br>", DocsHtml.Paragraphs(cell.Blocks).Select(p => Inline(p.VisibleInlines(), "<br>")).Where(t => t.Length > 0));
                    }
                }

                grid.Add(line);
            }

            if (grid.Count == 0)
            {
                return;
            }

            void Row(string[] cells) => _sb.Append("| ").Append(string.Join(" | ", cells.Select(c => c.Replace("\n", " ", StringComparison.Ordinal)))).Append(" |\n");
            Row(grid[0]);
            _sb.Append('|').Append(string.Concat(Enumerable.Repeat(" --- |", columns))).Append('\n');
            foreach (var line in grid.Skip(1))
            {
                Row(line);
            }

            _sb.Append('\n');
        }

        public void Notes()
        {
            foreach (var (note, number) in DocsConvert.OrderedNotes(document))
            {
                var parts = DocsHtml.Paragraphs(note.Blocks).Select(p => Inline(p.VisibleInlines(), "\\\n    ")).Where(t => t.Length > 0).ToList();
                _sb.Append("[^").Append(number).Append("]: ").Append(string.Join("\n\n    ", parts)).Append('\n');
            }
        }

        private enum Mark
        {
            Link,
            Strong,
            Em,
            Strike,
            Underline,
            Sub,
            Sup,
        }

        private List<(Mark Mark, string? Href)> Chain(RunFormat f)
        {
            var chain = new List<(Mark, string?)>();
            if (f.Href is not null)
            {
                // A link inside the document points to the anchor of its heading.
                chain.Add((Mark.Link, f.Href.StartsWith('#') && _anchors.TryGetValue(f.Href[1..], out var anchor) ? "#" + anchor : f.Href));
            }

            if (f.Bold)
            {
                chain.Add((Mark.Strong, null));
            }

            if (f.Italic)
            {
                chain.Add((Mark.Em, null));
            }

            if (f.Strike)
            {
                chain.Add((Mark.Strike, null));
            }

            if (f.Underline)
            {
                chain.Add((Mark.Underline, null));
            }

            if (f.Script != 0)
            {
                chain.Add((f.Script > 0 ? Mark.Sup : Mark.Sub, null));
            }

            return chain;
        }

        private static string OpenMark(Mark mark) => mark switch
        {
            Mark.Link => "[",
            Mark.Strong => "**",
            Mark.Em => "*",
            Mark.Strike => "~~",
            Mark.Underline => "<u>",
            Mark.Sub => "<sub>",
            _ => "<sup>",
        };

        private static string CloseMark(Mark mark, string? href) => mark switch
        {
            Mark.Link => "](" + Url(href!) + ")",
            Mark.Strong => "**",
            Mark.Em => "*",
            Mark.Strike => "~~",
            Mark.Underline => "</u>",
            Mark.Sub => "</sub>",
            _ => "</sup>",
        };

        private static string Url(string href) => href.Replace(" ", "%20", StringComparison.Ordinal).Replace("(", "%28", StringComparison.Ordinal)
            .Replace(")", "%29", StringComparison.Ordinal).Replace("<", "%3C", StringComparison.Ordinal).Replace(">", "%3E", StringComparison.Ordinal);

        /// <summary>
        /// Inline content as Markdown: the marks open and close as the formatting changes (outer marks stay open as long as possible); spaces at the
        /// edges of formatted text move outside the marks (CommonMark does not close <c>**</c> after a space).
        /// </summary>
        private string Inline(IEnumerable<DocInline> inlines, string lineBreak)
        {
            var sb = new StringBuilder();
            var open = new List<(Mark Mark, string? Href)>();
            var pending = new StringBuilder();

            void Apply(RunFormat format)
            {
                var chain = Chain(format);
                var common = 0;
                while (common < open.Count && common < chain.Count && open[common] == chain[common])
                {
                    common++;
                }

                for (var i = open.Count - 1; i >= common; i--)
                {
                    sb.Append(CloseMark(open[i].Mark, open[i].Href));
                }

                open.RemoveRange(common, open.Count - common);
                sb.Append(pending);
                pending.Clear();
                for (var i = common; i < chain.Count; i++)
                {
                    sb.Append(OpenMark(chain[i].Mark));
                    open.Add(chain[i]);
                }
            }

            foreach (var inline in inlines)
            {
                switch (inline)
                {
                    case DocText t:
                        var text = t.Text;
                        var core = text.Trim(' ', '\t');
                        if (core.Length == 0)
                        {
                            pending.Append(text);
                            break;
                        }

                        var lead = text[..text.IndexOf(core, StringComparison.Ordinal)];
                        var trail = text[(lead.Length + core.Length)..];
                        pending.Append(lead);
                        Apply(t.Format);
                        sb.Append(Escape(core).Replace("\n", lineBreak, StringComparison.Ordinal));
                        pending.Append(trail);
                        break;
                    case DocLineBreak:
                        Apply(inline.Format);
                        sb.Append(lineBreak);
                        break;
                    case DocImage img:
                        Apply(inline.Format);
                        sb.Append("![").Append(Escape(img.Alt)).Append("](").Append(img.Src).Append(')');
                        break;
                    case DocNoteRef r when _notes.TryGetValue(r.Id, out var n):
                        Apply(inline.Format);
                        sb.Append("[^").Append(n).Append(']');
                        break;
                    case DocField f:
                        var value = DocsConvert.FieldText(f.Kind, string.Empty, document.Language);
                        if (value.Length > 0)
                        {
                            Apply(inline.Format);
                            sb.Append(Escape(value));
                        }

                        break;
                }
            }

            Apply(RunFormat.Plain);
            return sb.ToString().TrimEnd();
        }

        private static string Escape(string text)
        {
            var sb = new StringBuilder(text.Length + 8);
            foreach (var c in text)
            {
                if (c is '\\' or '`' or '*' or '_' or '[' or ']' or '<' or '>' or '|' or '~')
                {
                    sb.Append('\\');
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        /// <summary>A paragraph must not start like another block (a heading, a list item, a quote, a rule).</summary>
        private static string EscapeStart(string text)
        {
            if (text.Length == 0)
            {
                return text;
            }

            if (text[0] is '#' or '-' or '+' or '=')
            {
                return "\\" + text;
            }

            var match = OrderedStartRegex().Match(text);
            return match.Success ? text.Insert(match.Length - 1, "\\") : text;
        }
    }

    // ------------------------------------------------------------------ parser

    private sealed partial class Parser(DocsDocument document, string source)
    {
        private readonly Dictionary<string, string> _links = new(StringComparer.Ordinal);

        private readonly Dictionary<string, string> _noteIds = new(StringComparer.Ordinal);

        public void Run()
        {
            var lines = source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n').Select(ExpandTabs).ToList();
            foreach (var line in lines)
            {
                var definition = LinkDefinitionRegex().Match(line);
                if (definition.Success)
                {
                    _links.TryAdd(Key(definition.Groups[1].Value), definition.Groups[2].Value);
                }

                var note = NoteDefinitionRegex().Match(line);
                if (note.Success && !_noteIds.ContainsKey(Key(note.Groups[1].Value)))
                {
                    var id = new string(note.Groups[1].Value.Where(c => char.IsAsciiLetterOrDigit(c) || c is '_' or '-').Take(40).ToArray());
                    if (id.Length == 0 || _noteIds.ContainsValue(id))
                    {
                        id = "n" + (_noteIds.Count + 1).ToString(CultureInfo.InvariantCulture);
                    }

                    _noteIds[Key(note.Groups[1].Value)] = id;
                }
            }

            Blocks(lines, document.Body, 0, 0);
        }

        private static string ExpandTabs(string line)
        {
            if (!line.Contains('\t', StringComparison.Ordinal))
            {
                return line;
            }

            // Only the indentation: tabs inside the text stay.
            var sb = new StringBuilder();
            var i = 0;
            for (; i < line.Length && line[i] is ' ' or '\t'; i++)
            {
                if (line[i] == '\t')
                {
                    sb.Append(' ', 4 - (sb.Length % 4));
                }
                else
                {
                    sb.Append(' ');
                }
            }

            return sb.Append(line, i, line.Length - i).ToString();
        }

        private static string Key(string label) => WhitespaceRegex().Replace(label.Trim(), " ").ToLowerInvariant();

        private static int Indent(string line)
        {
            var n = 0;
            while (n < line.Length && line[n] == ' ')
            {
                n++;
            }

            return n;
        }

        private static bool Blank(string line) => line.Trim().Length == 0;

        /// <summary>Does a line start a block that ends a paragraph?</summary>
        private static bool StartsBlock(string line)
        {
            if (Indent(line) > 3)
            {
                return false;
            }

            var s = line.TrimStart();
            return AtxRegex().IsMatch(s) || FenceRegex().IsMatch(s) || RuleRegex().IsMatch(s) || s.StartsWith('>') || BulletRegex().IsMatch(s);
        }

        private void Blocks(List<string> lines, List<DocBlock> output, int level, int depth)
        {
            var paragraph = new List<string>();
            var listMarker = (string?)null;

            void Flush()
            {
                if (paragraph.Count > 0)
                {
                    var p = new DocParagraph();
                    Inline(string.Join('\n', paragraph), p);
                    output.Add(p);
                    paragraph.Clear();
                }
            }

            var i = 0;
            while (i < lines.Count)
            {
                var line = lines[i];
                if (Blank(line))
                {
                    Flush();
                    i++;
                    continue;
                }

                var indent = Indent(line);
                if (indent >= 4 && paragraph.Count == 0)
                {
                    var code = new List<string>();
                    while (i < lines.Count && (Indent(lines[i]) >= 4 || Blank(lines[i])))
                    {
                        code.Add(lines[i].Length >= 4 ? lines[i][4..] : string.Empty);
                        i++;
                    }

                    while (code.Count > 0 && Blank(code[^1]))
                    {
                        code.RemoveAt(code.Count - 1);
                    }

                    output.Add(Code(code));
                    listMarker = null;
                    continue;
                }

                if (indent >= 4)
                {
                    paragraph.Add(line.Trim());
                    i++;
                    continue;
                }

                var s = line.TrimStart();
                Match m;
                if ((m = FenceRegex().Match(s)).Success)
                {
                    Flush();
                    var fence = m.Groups[1].Value;
                    var code = new List<string>();
                    i++;
                    while (i < lines.Count && !(lines[i].TrimStart().StartsWith(fence, StringComparison.Ordinal) && lines[i].Trim().Trim(fence[0]).Length == 0))
                    {
                        var l = lines[i];
                        code.Add(l[Math.Min(indent, Indent(l))..]);
                        i++;
                    }

                    i++;
                    output.Add(Code(code));
                    listMarker = null;
                    continue;
                }

                if ((m = AtxRegex().Match(s)).Success)
                {
                    Flush();
                    var heading = new DocParagraph("h" + m.Groups[1].Value.Length);
                    Inline(m.Groups[2].Value, heading);
                    output.Add(heading);
                    listMarker = null;
                    i++;
                    continue;
                }

                if (paragraph.Count > 0 && SetextRegex().Match(s) is { Success: true } setext)
                {
                    var heading = new DocParagraph(setext.Groups[1].Value[0] == '=' ? "h1" : "h2");
                    Inline(string.Join('\n', paragraph), heading);
                    paragraph.Clear();
                    output.Add(heading);
                    i++;
                    continue;
                }

                if (RuleRegex().IsMatch(s))
                {
                    Flush();
                    output.Add(new DocRule());
                    listMarker = null;
                    i++;
                    continue;
                }

                if (s.StartsWith('>'))
                {
                    Flush();
                    var quote = new List<string>();
                    while (i < lines.Count && !Blank(lines[i]))
                    {
                        var q = lines[i].TrimStart();
                        if (q.StartsWith('>'))
                        {
                            quote.Add(q.Length > 1 && q[1] == ' ' ? q[2..] : q[1..]);
                        }
                        else if (quote.Count > 0 && !StartsBlock(lines[i]))
                        {
                            quote.Add(q);
                        }
                        else
                        {
                            break;
                        }

                        i++;
                    }

                    var inner = new List<DocBlock>();
                    if (depth < 20)
                    {
                        Blocks(quote, inner, 0, depth + 1);
                    }

                    foreach (var block in inner)
                    {
                        if (block is DocParagraph qp && qp.Tag != "pre")
                        {
                            qp.Tag = "blockquote";
                            qp.Style = null;
                            qp.List = null;
                        }

                        output.Add(block);
                    }

                    listMarker = null;
                    continue;
                }

                if ((m = BulletRegex().Match(s)).Success)
                {
                    Flush();
                    var marker = m.Groups[1].Value;
                    var ordered = char.IsAsciiDigit(marker[0]);
                    var type = ordered ? marker[^1].ToString() : marker;
                    var spaces = m.Groups[2].Value.Length;
                    var offset = indent + marker.Length + (spaces is 0 or > 4 ? 1 : spaces);
                    var item = new List<string> { m.Groups[3].Value };
                    i++;
                    while (i < lines.Count)
                    {
                        var l = lines[i];
                        if (Blank(l))
                        {
                            var next = i + 1;
                            while (next < lines.Count && Blank(lines[next]))
                            {
                                next++;
                            }

                            if (next < lines.Count && Indent(lines[next]) >= offset)
                            {
                                item.Add(string.Empty);
                                i++;
                                continue;
                            }

                            break;
                        }

                        if (Indent(l) >= offset)
                        {
                            item.Add(l[offset..]);
                        }
                        else if (!Blank(item[^1]) && !StartsBlock(l) && !(BulletRegex().IsMatch(l.TrimStart()) && Indent(l) <= 3))
                        {
                            item.Add(l.TrimStart());
                        }
                        else
                        {
                            break;
                        }

                        i++;
                    }

                    var number = ordered && int.TryParse(marker[..^1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 1;
                    var first = listMarker != type;
                    listMarker = type;
                    ListItem(item, output, level, ordered ? "decimal" : "bullet", first && ordered && number != 1 ? Math.Clamp(number, 1, 99999) : null, depth);
                    continue;
                }

                if (s.Contains('|', StringComparison.Ordinal) && i + 1 < lines.Count && DelimiterRow(lines[i + 1]) is { } align && SplitRow(s).Count == align.Count)
                {
                    Flush();
                    var table = new DocTable { TableStyle = "grid", Width = "100%" };
                    table.Rows.Add(TableRow(s, align, header: true));
                    i += 2;
                    while (i < lines.Count && !Blank(lines[i]) && lines[i].Contains('|', StringComparison.Ordinal))
                    {
                        table.Rows.Add(TableRow(lines[i].Trim(), align, header: false));
                        i++;
                    }

                    output.Add(table);
                    listMarker = null;
                    continue;
                }

                if ((m = NoteDefinitionRegex().Match(line)).Success && paragraph.Count == 0)
                {
                    var note = new List<string> { m.Groups[2].Value };
                    i++;
                    while (i < lines.Count && (Indent(lines[i]) >= 4 || (Blank(lines[i]) && i + 1 < lines.Count && Indent(lines[i + 1]) >= 4)))
                    {
                        note.Add(lines[i].Length >= 4 ? lines[i][4..] : string.Empty);
                        i++;
                    }

                    var blocks = new List<DocBlock>();
                    if (depth < 20)
                    {
                        Blocks(note, blocks, 0, depth + 1);
                    }

                    var id = _noteIds.GetValueOrDefault(Key(m.Groups[1].Value));
                    if (id is not null && document.Notes.All(x => x.Id != id) && document.Notes.Count < DocsDocument.MaxNotes)
                    {
                        document.Notes.Add(new DocNote { Id = id, Blocks = blocks });
                    }

                    continue;
                }

                if (paragraph.Count == 0 && LinkDefinitionRegex().IsMatch(line))
                {
                    i++;
                    continue;
                }

                if (paragraph.Count == 0 && HtmlCommentRegex().IsMatch(s))
                {
                    i++;
                    continue;
                }

                listMarker = null;
                paragraph.Add(s);
                i++;
            }

            Flush();
        }

        private void ListItem(List<string> lines, List<DocBlock> output, int level, string kind, int? start, int depth)
        {
            var inner = new List<DocBlock>();
            if (depth < 20)
            {
                Blocks(lines, inner, Math.Min(level + 1, 8), depth + 1);
            }

            var used = false;
            foreach (var block in inner)
            {
                if (block is DocParagraph p && p.List is null && p.Tag is "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                {
                    if (!used)
                    {
                        p.List = kind;
                        p.Level = Math.Min(level, 8);
                        p.Start = start;
                        used = true;
                    }
                    else if (p.Tag == "p")
                    {
                        p.Style = "ListParagraph";
                        if (level > 0)
                        {
                            p.Css.Set("margin-left", CssList.Pt((level + 1) * 36));
                        }
                    }
                }

                output.Add(block);
            }

            if (!used)
            {
                output.Insert(output.Count - inner.Count, new DocParagraph { List = kind, Level = Math.Min(level, 8), Start = start });
            }
        }

        private static DocParagraph Code(List<string> lines)
        {
            var p = new DocParagraph("pre");
            var text = DocsHtml.CleanText(string.Join('\n', lines));
            if (text.Length > 0)
            {
                p.Inlines.Add(new DocText(text));
            }

            return p;
        }

        private static List<string> SplitRow(string row)
        {
            var s = row.Trim();
            if (s.StartsWith('|'))
            {
                s = s[1..];
            }

            if (s.EndsWith('|') && !s.EndsWith("\\|", StringComparison.Ordinal))
            {
                s = s[..^1];
            }

            var cells = new List<string>();
            var sb = new StringBuilder();
            var code = false;
            for (var i = 0; i < s.Length; i++)
            {
                var c = s[i];
                if (c == '\\' && i + 1 < s.Length && s[i + 1] == '|')
                {
                    sb.Append("\\|");
                    i++;
                    continue;
                }

                if (c == '`')
                {
                    code = !code;
                }

                if (c == '|' && !code)
                {
                    cells.Add(sb.ToString().Trim());
                    sb.Clear();
                    continue;
                }

                sb.Append(c);
            }

            cells.Add(sb.ToString().Trim());
            return cells;
        }

        private static List<string?>? DelimiterRow(string line)
        {
            if (!line.Contains('-', StringComparison.Ordinal) || Indent(line) > 3)
            {
                return null;
            }

            var cells = SplitRow(line);
            var align = new List<string?>();
            foreach (var cell in cells)
            {
                if (!DelimiterCellRegex().IsMatch(cell))
                {
                    return null;
                }

                align.Add(cell.StartsWith(':') && cell.EndsWith(':') ? "center" : cell.EndsWith(':') ? "right" : null);
            }

            return align;
        }

        private DocRow TableRow(string line, List<string?> align, bool header)
        {
            var cells = SplitRow(line);
            var row = new DocRow { Header = header };
            for (var c = 0; c < align.Count; c++)
            {
                var cell = new DocCell { HeaderCell = header };
                var p = new DocParagraph();
                if (align[c] is { } a)
                {
                    p.Css.Set("text-align", a);
                }

                Inline(c < cells.Count ? cells[c] : string.Empty, p);
                cell.Blocks.Add(p);
                row.Cells.Add(cell);
            }

            return row;
        }

        // ---------------------------------------------------------------- inline

        private abstract class Node;

        private sealed class TextNode(string text) : Node
        {
            public StringBuilder Text { get; } = new(text);
        }

        private sealed class Delimiter(char c, int count, bool canOpen, bool canClose) : Node
        {
            public char Char { get; } = c;

            public int Count { get; set; } = count;

            public int Original { get; } = count;

            public bool CanOpen { get; } = canOpen;

            public bool CanClose { get; } = canClose;
        }

        private sealed class Bracket(bool image) : Node
        {
            public bool Image { get; } = image;

            public bool Active { get; set; } = true;
        }

        private enum Kind
        {
            Strong,
            Em,
            Strike,
            Link,
        }

        private sealed class Span(Kind kind, List<Node> children, string? href = null) : Node
        {
            public Kind Kind { get; } = kind;

            public List<Node> Children { get; } = children;

            public string? Href { get; } = href;
        }

        private sealed class CodeNode(string text) : Node
        {
            public string Text { get; } = text;
        }

        private sealed class BreakNode : Node;

        private sealed class ImageNode(string src, string alt) : Node
        {
            public string Src { get; } = src;

            public string Alt { get; } = alt;
        }

        private sealed class NoteNode(string id) : Node
        {
            public string Id { get; } = id;
        }

        private sealed class TagNode(string tag, bool close) : Node
        {
            public string Tag { get; } = tag;

            public bool Close { get; } = close;
        }

        private void Inline(string text, DocParagraph p)
        {
            var nodes = ParseInline(text);
            var state = new Dictionary<string, int>(StringComparer.Ordinal);
            Flatten(nodes, RunFormat.Plain, p, state);
            p.Inlines.RemoveAll(x => x is DocText { Text.Length: 0 });
            if (p.Inlines.FirstOrDefault() is DocText first)
            {
                first.Text = first.Text.TrimStart();
            }

            if (p.Inlines.LastOrDefault() is DocText last)
            {
                last.Text = last.Text.TrimEnd();
            }

            p.Inlines.RemoveAll(x => x is DocText { Text.Length: 0 });
        }

        private List<Node> ParseInline(string text)
        {
            var nodes = new List<Node>();
            var brackets = new List<Bracket>();

            void Literal(string s)
            {
                if (nodes.Count > 0 && nodes[^1] is TextNode t)
                {
                    t.Text.Append(s);
                }
                else
                {
                    nodes.Add(new TextNode(s));
                }
            }

            var i = 0;
            while (i < text.Length)
            {
                var c = text[i];
                switch (c)
                {
                    case '\\' when i + 1 < text.Length && char.IsAsciiLetterOrDigit(text[i + 1]) is false && text[i + 1] < 128 && text[i + 1] != '\n' && !char.IsWhiteSpace(text[i + 1]):
                        Literal(text[i + 1].ToString());
                        i += 2;
                        continue;
                    case '\\' when i + 1 < text.Length && text[i + 1] == '\n':
                        nodes.Add(new BreakNode());
                        i += 2;
                        continue;
                    case '\n':
                        if (nodes.Count > 0 && nodes[^1] is TextNode before)
                        {
                            var trailing = before.Text.ToString();
                            var trimmed = trailing.TrimEnd(' ');
                            before.Text.Length = trimmed.Length;
                            if (trailing.Length - trimmed.Length >= 2)
                            {
                                nodes.Add(new BreakNode());
                                i++;
                                continue;
                            }
                        }

                        Literal(" ");
                        i++;
                        while (i < text.Length && text[i] == ' ')
                        {
                            i++;
                        }

                        continue;
                    case '`':
                        var run = Run(text, i, '`');
                        var close = text.IndexOf(new string('`', run), i + run, StringComparison.Ordinal);
                        while (close >= 0 && close + run < text.Length && text[close + run] == '`')
                        {
                            close = text.IndexOf(new string('`', run), close + run + 1, StringComparison.Ordinal);
                        }

                        if (close < 0)
                        {
                            Literal(new string('`', run));
                            i += run;
                            continue;
                        }

                        var code = text[(i + run)..close].Replace('\n', ' ');
                        if (code.Length > 2 && code[0] == ' ' && code[^1] == ' ' && code.Trim().Length > 0)
                        {
                            code = code[1..^1];
                        }

                        nodes.Add(new CodeNode(code));
                        i = close + run;
                        continue;
                    case '*' or '_' or '~':
                        var count = Run(text, i, c);
                        if (c == '~' && count < 2)
                        {
                            Literal("~");
                            i++;
                            continue;
                        }

                        var prev = i > 0 ? text[i - 1] : ' ';
                        var next = i + count < text.Length ? text[i + count] : ' ';
                        var left = !char.IsWhiteSpace(next) && (!IsPunctuation(next) || char.IsWhiteSpace(prev) || IsPunctuation(prev));
                        var right = !char.IsWhiteSpace(prev) && (!IsPunctuation(prev) || char.IsWhiteSpace(next) || IsPunctuation(next));
                        var canOpen = c == '_' ? left && (!right || IsPunctuation(prev)) : left;
                        var canClose = c == '_' ? right && (!left || IsPunctuation(next)) : right;
                        nodes.Add(new Delimiter(c, count, canOpen, canClose));
                        i += count;
                        continue;
                    case '!' when i + 1 < text.Length && text[i + 1] == '[':
                        var image = new Bracket(true);
                        nodes.Add(image);
                        brackets.Add(image);
                        i += 2;
                        continue;
                    case '[':
                        var note = NoteReferenceRegex().Match(text, i);
                        if (note.Success && note.Index == i && _noteIds.TryGetValue(Key(note.Groups[1].Value), out var noteId))
                        {
                            nodes.Add(new NoteNode(noteId));
                            i += note.Length;
                            continue;
                        }

                        var bracket = new Bracket(false);
                        nodes.Add(bracket);
                        brackets.Add(bracket);
                        i++;
                        continue;
                    case ']':
                        if (CloseBracket(text, ref i, nodes, brackets))
                        {
                            continue;
                        }

                        Literal("]");
                        i++;
                        continue;
                    case '<':
                        var auto = AutolinkRegex().Match(text, i);
                        if (auto.Success && auto.Index == i)
                        {
                            var target = auto.Groups[1].Value;
                            var href = target.Contains('@', StringComparison.Ordinal) && !target.Contains(':', StringComparison.Ordinal) ? "mailto:" + target : target;
                            nodes.Add(new Span(Kind.Link, [new TextNode(target)], href));
                            i += auto.Length;
                            continue;
                        }

                        var tag = TagRegex().Match(text, i);
                        if (tag.Success && tag.Index == i)
                        {
                            var name = tag.Groups[2].Value.ToLowerInvariant();
                            nodes.Add(name == "br" ? new BreakNode() : new TagNode(name, tag.Groups[1].Value == "/"));
                            i += tag.Length;
                            continue;
                        }

                        Literal("<");
                        i++;
                        continue;
                    case '&':
                        var entity = EntityRegex().Match(text, i);
                        if (entity.Success && entity.Index == i)
                        {
                            Literal(WebUtility.HtmlDecode(entity.Value));
                            i += entity.Length;
                            continue;
                        }

                        Literal("&");
                        i++;
                        continue;
                    case 'h' or 'w' when (i == 0 || !char.IsLetterOrDigit(text[i - 1])) && BareUrlRegex().Match(text, i) is { Success: true } url && url.Index == i:
                        var link = url.Value.TrimEnd('.', ',', ':', ';', '!', '?', '*', '_', '~', '\'', '"');
                        if (link.EndsWith(')') && link.Count(x => x == '(') < link.Count(x => x == ')'))
                        {
                            link = link[..^1];
                        }

                        nodes.Add(new Span(Kind.Link, [new TextNode(link)], link.StartsWith("www.", StringComparison.OrdinalIgnoreCase) ? "http://" + link : link));
                        i += link.Length;
                        continue;
                }

                var end = i + 1;
                while (end < text.Length && text[end] is not ('\\' or '\n' or '`' or '*' or '_' or '~' or '!' or '[' or ']' or '<' or '&' or 'h' or 'w'))
                {
                    end++;
                }

                Literal(text[i..end]);
                i = end;
            }

            Emphasis(nodes);
            return nodes;
        }

        private static int Run(string text, int i, char c)
        {
            var n = 0;
            while (i + n < text.Length && text[i + n] == c)
            {
                n++;
            }

            return n;
        }

        private static bool IsPunctuation(char c) => char.IsPunctuation(c) || char.IsSymbol(c);

        /// <summary>A <c>]</c>: with the last open bracket and a destination it makes a link or a picture.</summary>
        private bool CloseBracket(string text, ref int i, List<Node> nodes, List<Bracket> brackets)
        {
            if (brackets.Count == 0)
            {
                return false;
            }

            var opener = brackets[^1];
            brackets.RemoveAt(brackets.Count - 1);
            if (!opener.Active)
            {
                return false;
            }

            var index = nodes.IndexOf(opener);
            var after = i + 1;
            string? destination = null;
            var consumed = 0;
            var inline = InlineLinkRegex().Match(text, after);
            if (inline.Success && inline.Index == after)
            {
                destination = inline.Groups[1].Success ? inline.Groups[1].Value : inline.Groups[2].Value;
                consumed = inline.Length;
            }
            else
            {
                var label = string.Concat(nodes.Skip(index + 1).Select(PlainText));
                var reference = ReferenceRegex().Match(text, after);
                if (reference.Success && reference.Index == after)
                {
                    var key = reference.Groups[1].Value.Length > 0 ? reference.Groups[1].Value : label;
                    if (_links.TryGetValue(Key(key), out var url))
                    {
                        destination = url;
                        consumed = reference.Length;
                    }
                }
                else if (_links.TryGetValue(Key(label), out var url))
                {
                    destination = url;
                }
            }

            if (destination is null)
            {
                nodes[index] = new TextNode(opener.Image ? "![" : "[");
                return false;
            }

            var children = nodes.GetRange(index + 1, nodes.Count - index - 1);
            nodes.RemoveRange(index, nodes.Count - index);
            if (opener.Image)
            {
                nodes.Add(new ImageNode(destination.Trim('<', '>'), string.Concat(children.Select(PlainText))));
            }
            else
            {
                Emphasis(children);
                nodes.Add(new Span(Kind.Link, children, destination.Trim('<', '>')));
                foreach (var earlier in brackets.Where(b => !b.Image))
                {
                    earlier.Active = false;
                }
            }

            i = after + consumed;
            return true;
        }

        private static string PlainText(Node node) => node switch
        {
            TextNode t => t.Text.ToString(),
            Delimiter d => new string(d.Char, d.Count),
            CodeNode c => c.Text,
            Span s => string.Concat(s.Children.Select(PlainText)),
            Bracket b => b.Image ? "![" : "[",
            _ => string.Empty,
        };

        /// <summary>CommonMark's "process emphasis": pairs of delimiters (<c>*</c>, <c>_</c>, <c>~~</c>) become strong, emphasis or strikethrough.</summary>
        private static void Emphasis(List<Node> nodes)
        {
            var from = 0;
            while (true)
            {
                var c = -1;
                for (var k = from; k < nodes.Count; k++)
                {
                    if (nodes[k] is Delimiter { CanClose: true, Count: > 0 })
                    {
                        c = k;
                        break;
                    }
                }

                if (c < 0)
                {
                    break;
                }

                var closer = (Delimiter)nodes[c];
                var o = -1;
                for (var k = c - 1; k >= 0; k--)
                {
                    if (nodes[k] is not Delimiter { CanOpen: true, Count: > 0 } d || d.Char != closer.Char)
                    {
                        continue;
                    }

                    if (closer.Char == '~')
                    {
                        if (d.Count < 2 || closer.Count < 2)
                        {
                            continue;
                        }
                    }
                    else if ((d.CanClose || closer.CanOpen) && (d.Original + closer.Original) % 3 == 0 && !(d.Original % 3 == 0 && closer.Original % 3 == 0))
                    {
                        continue;
                    }

                    o = k;
                    break;
                }

                if (o < 0)
                {
                    from = c + 1;
                    continue;
                }

                var opener = (Delimiter)nodes[o];
                var use = closer.Char == '~' ? 2 : opener.Count >= 2 && closer.Count >= 2 ? 2 : 1;
                var kind = closer.Char == '~' ? Kind.Strike : use == 2 ? Kind.Strong : Kind.Em;
                var inner = nodes.GetRange(o + 1, c - o - 1);
                nodes.RemoveRange(o + 1, c - o - 1);
                nodes.Insert(o + 1, new Span(kind, inner));
                opener.Count -= use;
                closer.Count -= use;
                c = o + 2;
                if (opener.Count == 0)
                {
                    nodes.RemoveAt(o);
                    c--;
                }

                if (closer.Count == 0)
                {
                    nodes.RemoveAt(c);
                }

                from = c;
            }
        }

        private static void Flatten(List<Node> nodes, RunFormat format, DocParagraph p, Dictionary<string, int> state)
        {
            foreach (var node in nodes)
            {
                var f = Apply(format, state);
                switch (node)
                {
                    case TextNode t:
                        AddText(p, t.Text.ToString(), f);
                        break;
                    case Delimiter d:
                        AddText(p, new string(d.Char, d.Count), f);
                        break;
                    case Bracket b:
                        AddText(p, b.Image ? "![" : "[", f);
                        break;
                    case CodeNode code:
                        AddText(p, code.Text, f with { FontFamily = "Consolas, 'Courier New', monospace" });
                        break;
                    case BreakNode:
                        p.Inlines.Add(new DocLineBreak { Format = f with { Script = 0 } });
                        break;
                    case ImageNode image when DocsHtml.SafeImage(image.Src) is { } src:
                        p.Inlines.Add(new DocImage { Src = src, Alt = image.Alt, Format = f with { Script = 0 } });
                        break;
                    case NoteNode note:
                        p.Inlines.Add(new DocNoteRef(note.Id) { Format = f with { Script = 0 } });
                        break;
                    case TagNode tag:
                        state[tag.Tag] = Math.Max(0, state.GetValueOrDefault(tag.Tag) + (tag.Close ? -1 : 1));
                        break;
                    case Span span:
                        var inner = span.Kind switch
                        {
                            Kind.Strong => format with { Bold = true },
                            Kind.Em => format with { Italic = true },
                            Kind.Strike => format with { Strike = true },
                            // Links inside the file stay as they are until the headings are known (FromMarkdown resolves them).
                            _ => format with { Href = span.Href is ['#', ..] ? span.Href : DocsHtml.SafeHref(span.Href) ?? format.Href },
                        };
                        Flatten(span.Children, inner, p, state);
                        break;
                }
            }
        }

        private static RunFormat Apply(RunFormat format, Dictionary<string, int> state)
        {
            if (state.Count == 0)
            {
                return format;
            }

            bool On(params string[] tags) => tags.Any(t => state.GetValueOrDefault(t) > 0);
            var f = format;
            if (On("u", "ins"))
            {
                f = f with { Underline = true };
            }

            if (On("s", "del", "strike"))
            {
                f = f with { Strike = true };
            }

            if (On("b", "strong"))
            {
                f = f with { Bold = true };
            }

            if (On("i", "em"))
            {
                f = f with { Italic = true };
            }

            if (On("mark"))
            {
                f = f with { Background = "#ffff00" };
            }

            if (On("sup"))
            {
                f = f with { Script = 1 };
            }
            else if (On("sub"))
            {
                f = f with { Script = -1 };
            }

            return f;
        }

        private static void AddText(DocParagraph p, string text, RunFormat format)
        {
            text = DocsHtml.CleanText(text);
            if (text.Length == 0)
            {
                return;
            }

            if (p.Inlines.Count > 0 && p.Inlines[^1] is DocText last && last.Format == format)
            {
                last.Text += text;
            }
            else
            {
                p.Inlines.Add(new DocText(text) { Format = format });
            }
        }

        [GeneratedRegex(@"^(#{1,6})(?:[ \t]+(.*?))??(?:[ \t]+#+)?[ \t]*$")]
        private static partial Regex AtxRegex();

        [GeneratedRegex(@"^(`{3,}|~{3,})")]
        private static partial Regex FenceRegex();

        [GeneratedRegex(@"^(=+|-+)[ \t]*$")]
        private static partial Regex SetextRegex();

        [GeneratedRegex(@"^([-*_])([ \t]*\1){2,}[ \t]*$")]
        private static partial Regex RuleRegex();

        [GeneratedRegex(@"^([-+*]|\d{1,9}[.)])([ \t]+|$)(.*)$")]
        private static partial Regex BulletRegex();

        [GeneratedRegex(@"^:?-+:?$")]
        private static partial Regex DelimiterCellRegex();

        [GeneratedRegex(@"^ {0,3}\[\^([^\]]+)\]:[ \t]*(.*)$")]
        private static partial Regex NoteDefinitionRegex();

        [GeneratedRegex(@"^ {0,3}\[(?!\^)([^\]]+)\]:[ \t]*<?([^\s>]+)>?(?:[ \t]+(?:""[^""]*""|'[^']*'|\([^)]*\)))?[ \t]*$")]
        private static partial Regex LinkDefinitionRegex();

        [GeneratedRegex(@"\G\[\^([^\]\s]+)\]")]
        private static partial Regex NoteReferenceRegex();

        [GeneratedRegex(@"\G\(\s*(?:<([^<>\n]*)>|((?:[^\s()]|\([^\s()]*\))*))(?:\s+(?:""[^""]*""|'[^']*'|\([^)]*\)))?\s*\)")]
        private static partial Regex InlineLinkRegex();

        [GeneratedRegex(@"\G\[([^\]]*)\]")]
        private static partial Regex ReferenceRegex();

        [GeneratedRegex(@"\G<([A-Za-z][A-Za-z0-9+.\-]{1,31}:[^\s<>]*|[A-Za-z0-9.!#$%&'*+/=?^_`{|}~\-]+@[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)*)>")]
        private static partial Regex AutolinkRegex();

        [GeneratedRegex(@"\G<(/?)(u|ins|sub|sup|s|del|strike|b|strong|i|em|mark|br)\s*/?>", RegexOptions.IgnoreCase)]
        private static partial Regex TagRegex();

        [GeneratedRegex(@"\G&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")]
        private static partial Regex EntityRegex();

        [GeneratedRegex(@"\G(?:https?://|www\.)[^\s<]+", RegexOptions.IgnoreCase)]
        private static partial Regex BareUrlRegex();

        [GeneratedRegex(@"^<!--.*-->\s*$")]
        private static partial Regex HtmlCommentRegex();

        [GeneratedRegex(@"\s+")]
        private static partial Regex WhitespaceRegex();
    }

    [GeneratedRegex(@"^\d{1,9}[.)]")]
    private static partial Regex OrderedStartRegex();
}
