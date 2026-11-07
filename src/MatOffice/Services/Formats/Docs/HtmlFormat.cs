using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// HTML files. Export: a standalone HTML5 page, every style inline (the paragraph styles written out, lists as real <c>ul</c>/<c>ol</c>, tables with their
/// borders and shading, pictures as <c>data:</c> addresses, notes as a numbered list at the end); the class names of Word (<c>MsoTitle</c> …) keep the
/// styles for the way back. Import: any HTML page, also from Word (<c>mso-list</c> paragraphs, footnotes, <c>@page</c>) and Google Docs.
/// </summary>
internal static partial class HtmlFormat
{
    private static readonly Dictionary<string, string> ClassOfStyle = new(StringComparer.Ordinal)
    {
        ["Title"] = "MsoTitle", ["Subtitle"] = "MsoSubtitle", ["NoSpacing"] = "MsoNoSpacing", ["ListParagraph"] = "MsoListParagraph",
        ["Caption"] = "MsoCaption", ["Quote"] = "MsoQuote", ["IntenseQuote"] = "MsoIntenseQuote", ["TOC1"] = "MsoToc1", ["TOC2"] = "MsoToc2",
        ["TOC3"] = "MsoToc3", ["TOCHeading"] = "MsoTocHeading",
    };

    private static readonly Dictionary<string, string> StyleOfClass = ClassOfStyle.ToDictionary(p => p.Value.ToLowerInvariant(), p => p.Key, StringComparer.Ordinal);

    // ------------------------------------------------------------------ import

    public static DocsDocument Read(Stream input)
    {
        var bytes = DocsConvert.ReadAll(input);
        return FromHtml(DocsConvert.Decode(bytes, SniffCharset(bytes)));
    }

    /// <summary>The charset a page declares in its first bytes (<c>&lt;meta charset&gt;</c>), or null.</summary>
    private static string? SniffCharset(byte[] bytes)
    {
        var head = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 4096));
        var match = CharsetRegex().Match(head);
        return match.Success ? match.Groups[1].Value : null;
    }

    public static DocsDocument FromHtml(string html)
    {
        var dom = new HtmlParser(new HtmlParserOptions { IsScripting = false }).ParseDocument(html);
        var isWord = html.Contains("urn:schemas-microsoft-com:office", StringComparison.OrdinalIgnoreCase)
                     || dom.QuerySelector("meta[name=Generator]")?.GetAttribute("content")?.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase) == true;
        var options = HtmlReadOptions.Import with { IsWord = isWord };
        var document = new DocsDocument();
        if (dom.DocumentElement.GetAttribute("lang") is { } lang && LanguageCode(lang) is { } code)
        {
            document.Language = code;
        }

        ReadPage(dom, document.Page);
        var body = dom.Body;
        if (body is null)
        {
            document.Body.Add(new DocParagraph());
            return document;
        }

        CollapseWhitespace(body);
        ReadNotes(body, document, options);
        foreach (var comments in body.QuerySelectorAll("[style*=comment-list], .msocomtxt, .MsoCommentReference").ToList())
        {
            comments.Remove();
        }

        WordLists(body);
        StylesFromClasses(dom, body);
        PageBreaks(dom, body);
        TableStyles(body);
        ImageAlignment(body);
        foreach (var marker in body.QuerySelectorAll("*").Where(e => e.LocalName == "o:p" && string.IsNullOrWhiteSpace(e.TextContent.Replace(' ', ' '))).ToList())
        {
            marker.Remove();
        }

        document.Body = DocsHtml.Parse(dom, body, options);
        CleanUp(document.Body);
        foreach (var note in document.Notes)
        {
            CleanUp(note.Blocks);
        }

        if (document.Body.Count == 0)
        {
            document.Body.Add(new DocParagraph());
        }

        document.DropUnreferencedNotes();
        return document;
    }

    private static string? LanguageCode(string lang)
    {
        var parts = lang.Trim().Split('-', '_');
        if (parts[0].Length != 2 || !parts[0].All(char.IsAsciiLetter))
        {
            return null;
        }

        var code = parts[0].ToLowerInvariant();
        return parts.Length > 1 && parts[1].Length == 2 && parts[1].All(char.IsAsciiLetter) ? code + "-" + parts[1].ToUpperInvariant() : code;
    }

    /// <summary>The paper and the margins of <c>@page</c> (Word writes them, and the export of Docs).</summary>
    private static void ReadPage(IHtmlDocument dom, PageSetup page)
    {
        foreach (var style in dom.QuerySelectorAll("style"))
        {
            var match = PageRuleRegex().Match(style.TextContent);
            if (!match.Success)
            {
                continue;
            }

            var rule = match.Groups[1].Value;
            var size = SizeRegex().Match(rule);
            if (size.Success && DocsCss.ToPt(size.Groups[1].Value) is { } w && DocsCss.ToPt(size.Groups[2].Value) is { } h && w > 0 && h > 0)
            {
                var (wm, hm) = (w / 72 * 25.4, h / 72 * 25.4);
                page.Landscape = wm > hm;
                var (portraitW, portraitH) = page.Landscape ? (hm, wm) : (wm, hm);
                page.Size = PageSetup.SizeOf(portraitW, portraitH);
                page.Width = Math.Round(portraitW, 1);
                page.Height = Math.Round(portraitH, 1);
            }

            var margin = MarginRegex().Match(rule);
            if (margin.Success)
            {
                var values = DocsCss.SplitSpaces(margin.Groups[1].Value.Trim()).Select(v => DocsCss.ToPt(v) is { } pt ? pt / 72 * 25.4 : (double?)null).ToList();
                if (values.Count is >= 1 and <= 4 && values.All(v => v is not null))
                {
                    page.Top = Math.Round(values[0]!.Value, 1);
                    page.Right = Math.Round(values.Count > 1 ? values[1]!.Value : page.Top, 1);
                    page.Bottom = Math.Round(values.Count > 2 ? values[2]!.Value : page.Top, 1);
                    page.Left = Math.Round(values.Count > 3 ? values[3]!.Value : page.Right, 1);
                }
            }

            page.Normalize();
            return;
        }
    }

    /// <summary>Whitespace of the source (line breaks, indentation) becomes single spaces, as a browser shows it; not in pre or pre-wrap.</summary>
    private static void CollapseWhitespace(IElement root)
    {
        var stack = new Stack<(INode Node, bool Keep)>();
        stack.Push((root, false));
        while (stack.Count > 0)
        {
            var (node, keep) = stack.Pop();
            foreach (var child in node.ChildNodes)
            {
                if (child is IText text)
                {
                    if (!keep)
                    {
                        text.Data = WhitespaceRegex().Replace(text.Data, " ");
                    }
                }
                else if (child is IElement element)
                {
                    var style = element.GetAttribute("style") ?? string.Empty;
                    var pre = element.LocalName is "pre" or "textarea" or "xmp" || PreRegex().IsMatch(style);
                    stack.Push((element, keep || pre));
                }
            }
        }
    }

    /// <summary>Notes of the export of Docs (<c>ol.mo-notes</c>) and footnotes of Word's HTML (<c>mso-footnote-id</c>).</summary>
    private static void ReadNotes(IElement body, DocsDocument document, HtmlReadOptions options)
    {
        foreach (var list in body.QuerySelectorAll("ol.mo-notes, hr.mo-notes-rule").ToList())
        {
            if (list.LocalName == "ol")
            {
                foreach (var item in list.Children.Where(c => c.LocalName == "li"))
                {
                    var id = item.GetAttribute("id") is { } fid && fid.StartsWith("fn-", StringComparison.Ordinal) ? fid[3..] : null;
                    if (id is not null && DocsHtml.IsItemId(id))
                    {
                        document.Notes.Add(new DocNote { Id = id, Blocks = DocsHtml.Parse(item.InnerHtml, options) });
                    }
                }
            }

            list.Remove();
        }

        // Word: <a style='mso-footnote-id:ftn1' href="#_ftn1"> in the text, <div style='mso-element:footnote' id=ftn1> at the end.
        var containers = body.QuerySelectorAll("div[style*=footnote-list], div[style*=endnote-list]").ToList();
        foreach (var note in body.QuerySelectorAll("div[style*=mso-element]").Where(d => d.Id is { } i && (i.StartsWith("ftn", StringComparison.Ordinal) || i.StartsWith("edn", StringComparison.Ordinal))))
        {
            foreach (var back in note.QuerySelectorAll("a[href^='#_ftnref'], a[href^='#_ednref']").ToList())
            {
                back.Remove();
            }

            if (DocsHtml.IsItemId(note.Id))
            {
                document.Notes.Add(new DocNote { Id = note.Id!, Blocks = DocsHtml.Parse(note.InnerHtml, options) });
            }
        }

        foreach (var container in containers)
        {
            container.Remove();
        }

        foreach (var reference in body.QuerySelectorAll("a[href^='#_ftn'], a[href^='#_edn']").ToList())
        {
            var target = reference.GetAttribute("href")![2..];
            if (target.StartsWith("ftnref", StringComparison.Ordinal) || target.StartsWith("ednref", StringComparison.Ordinal))
            {
                continue;
            }

            var sup = body.Owner!.CreateElement("sup");
            sup.SetAttribute("data-fn", target);
            reference.Parent!.ReplaceChild(sup, reference);
        }
    }

    /// <summary>Paragraphs of Word's HTML with <c>mso-list:l0 level2 lfo1</c>: list items (the marker in a <c>mso-list:Ignore</c> span tells the kind).</summary>
    private static void WordLists(IElement body)
    {
        string? runKind = null;
        var lastNumber = new int[9];
        foreach (var block in body.QuerySelectorAll("p, h1, h2, h3, h4, h5, h6, li, div"))
        {
            var style = block.GetAttribute("style") ?? string.Empty;
            var match = MsoListRegex().Match(style);
            if (!match.Success)
            {
                if (block.LocalName is "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                {
                    runKind = null;
                    Array.Clear(lastNumber);
                }

                continue;
            }

            var level = Math.Clamp(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) - 1, 0, 8);
            var marker = block.QuerySelectorAll("[style]").FirstOrDefault(e => e.GetAttribute("style")!.Contains("mso-list:Ignore", StringComparison.OrdinalIgnoreCase));
            var text = (marker?.TextContent ?? string.Empty).Replace(' ', ' ').Trim();
            marker?.Remove();
            var (kind, number) = MarkerKind(text);
            if (level == 0 || runKind is null)
            {
                runKind = level == 0 ? kind : runKind ?? kind;
            }
            else if (DocsStyles.LevelFormat(runKind, level) == kind)
            {
                kind = runKind;
            }

            block.SetAttribute("data-list", kind);
            block.SetAttribute("data-level", level.ToString(CultureInfo.InvariantCulture));
            if (number is { } n && n != lastNumber[level] + 1 && n is >= 1 and <= 99999)
            {
                block.SetAttribute("data-start", n.ToString(CultureInfo.InvariantCulture));
            }

            lastNumber[level] = number ?? lastNumber[level] + 1;
            for (var i = level + 1; i < 9; i++)
            {
                lastNumber[i] = 0;
            }
        }
    }

    private static (string Kind, int? Number) MarkerKind(string marker)
    {
        var m = marker.TrimEnd('.', ')').Trim();
        if (m.Length == 0)
        {
            return ("bullet", null);
        }

        if (int.TryParse(m.Split('.')[^1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
        {
            return (marker.Count(c => c == '.') > 1 ? "outline" : "decimal", n);
        }

        if (m.All(c => "ivxlcdm".Contains(c, StringComparison.Ordinal)) && (m.Length > 1 || m is "i"))
        {
            return ("lower-roman", null);
        }

        if (m.All(c => "IVXLCDM".Contains(c, StringComparison.Ordinal)) && (m.Length > 1 || m is "I"))
        {
            return ("upper-roman", null);
        }

        if (m.Length == 1 && char.IsAsciiLetterLower(m[0]) && marker.Length > 1)
        {
            return ("lower-alpha", m[0] - 'a' + 1);
        }

        if (m.Length == 1 && char.IsAsciiLetterUpper(m[0]) && marker.Length > 1)
        {
            return ("upper-alpha", m[0] - 'A' + 1);
        }

        return (m switch
        {
            "o" or "◦" => "circle",
            "§" or "▪" or "■" => "square",
            "-" or "–" or "—" => "dash",
            "ü" or "✓" or "✔" => "check",
            "Ø" or "➢" or "►" or "→" => "arrow",
            _ => "bullet",
        }, null);
    }

    /// <summary>The classes of Word (<c>MsoTitle</c> …, also written by the export of Docs) as paragraph styles.</summary>
    private static void StylesFromClasses(IHtmlDocument dom, IElement body)
    {
        foreach (var element in body.QuerySelectorAll("p[class], div[class]").ToList())
        {
            if (element.HasAttribute("data-list"))
            {
                continue;
            }

            var key = element.ClassList.Select(c => StyleOfClass.GetValueOrDefault(Regex.Replace(c.ToLowerInvariant(), "cxsp(first|middle|last)$", string.Empty))).FirstOrDefault(k => k is not null);
            switch (key)
            {
                case null:
                    break;
                case "Quote" or "IntenseQuote":
                    var quote = dom.CreateElement("blockquote");
                    foreach (var attribute in element.Attributes)
                    {
                        quote.SetAttribute(attribute.Name, attribute.Value);
                    }

                    if (key == "IntenseQuote")
                    {
                        quote.SetAttribute("data-style", "IntenseQuote");
                    }

                    while (element.FirstChild is { } child)
                    {
                        quote.AppendChild(child);
                    }

                    element.Parent!.ReplaceChild(quote, element);
                    break;
                default:
                    element.SetAttribute("data-style", key);
                    break;
            }
        }
    }

    /// <summary><c>page-break-before: always</c> (Word) and <c>break-after: page</c> as page break blocks.</summary>
    private static void PageBreaks(IHtmlDocument dom, IElement body)
    {
        foreach (var element in body.QuerySelectorAll("[style]").ToList())
        {
            var style = element.GetAttribute("style")!;
            var before = BreakBeforeRegex().IsMatch(style);
            var after = BreakAfterRegex().IsMatch(style);
            if (!before && !after)
            {
                continue;
            }

            var pageBreak = dom.CreateElement("div");
            pageBreak.SetAttribute("data-break", "page");
            if (element.LocalName == "br" || (element.LocalName == "div" && !element.HasChildNodes))
            {
                element.Parent!.ReplaceChild(pageBreak, element);
            }
            else if (before)
            {
                element.Parent!.InsertBefore(pageBreak, element);
            }
            else
            {
                element.Parent!.InsertBefore(pageBreak, element.NextSibling);
            }
        }
    }

    /// <summary>Tables of foreign HTML: a border attribute means grid; a table without any border is plain.</summary>
    private static void TableStyles(IElement body)
    {
        foreach (var table in body.QuerySelectorAll("table"))
        {
            if (table.GetAttribute("data-table-style") is { } given && DocsStyles.TableStyles.Contains(given))
            {
                continue;
            }

            if (int.TryParse(table.GetAttribute("border"), out var border) && border > 0)
            {
                table.SetAttribute("data-table-style", "grid");
                continue;
            }

            var bordered = new[] { table }.Concat(table.QuerySelectorAll("td, th")).Any(e => (e.GetAttribute("style") ?? string.Empty).Contains("border", StringComparison.OrdinalIgnoreCase)
                && !BorderNoneRegex().IsMatch(e.GetAttribute("style")!));
            table.SetAttribute("data-table-style", bordered ? "grid" : "plain");
        }
    }

    /// <summary>A picture shown as a block with automatic margins is a centred picture.</summary>
    private static void ImageAlignment(IElement body)
    {
        foreach (var image in body.QuerySelectorAll("img[style]"))
        {
            var style = image.GetAttribute("style")!.Replace(" ", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
            if (!image.HasAttribute("data-align") && style.Contains("display:block", StringComparison.Ordinal) && style.Contains("auto", StringComparison.Ordinal))
            {
                image.SetAttribute("data-align", "center");
            }
        }
    }

    /// <summary>Formatting that only repeats the paragraph style (or the look of the table style) is dropped: the document keeps the style.</summary>
    private static void CleanUp(List<DocBlock> blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case DocParagraph p:
                    CleanParagraph(p);
                    break;
                case DocTable table:
                    CleanTable(table);
                    break;
            }
        }
    }

    private static void CleanParagraph(DocParagraph p)
    {
        var style = DocsStyles.Get(p.StyleKey);
        var css = p.Css;
        DropIf(css, "text-align", v => v == style.Align);
        DropIf(css, "line-height", v => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var l) && Math.Abs(l - style.Line) < 0.035);
        DropIf(css, "margin-top", v => Near(v, style.Before));
        DropIf(css, "margin-bottom", v => Near(v, style.After));
        DropIf(css, "margin-left", v => Near(v, style.IndentLeft));
        DropIf(css, "margin-right", v => Near(v, style.IndentRight));
        DropIf(css, "text-indent", v => Near(v, 0));
        DropIf(css, "background-color", v => v == style.Background);
        DropIf(css, "padding-left", v => Near(v, 0));
        DropIf(css, "padding-right", v => Near(v, 0));
        DropIf(css, "padding-top", v => Near(v, style.BorderColor is null ? 0 : style.Padding));
        DropIf(css, "padding-bottom", v => Near(v, style.BorderColor is null ? 0 : style.Padding));
        if (style.BorderColor is { } border)
        {
            DropIf(css, "border-top", v => v.EndsWith(border, StringComparison.Ordinal));
            DropIf(css, "border-bottom", v => v.EndsWith(border, StringComparison.Ordinal));
        }

        if (p.List is not null)
        {
            css.Remove("margin-left");
            css.Remove("text-indent");
            css.Remove("margin-top");
            css.Remove("margin-bottom");
        }

        foreach (var inline in p.Inlines)
        {
            var f = inline.Format;
            if (f.FontFamily is { } family && string.Equals(DocsCss.FirstFamily(family), style.Font, StringComparison.OrdinalIgnoreCase))
            {
                f = f with { FontFamily = null };
            }

            if (f.FontSize is { } size && Near(size, style.Size))
            {
                f = f with { FontSize = null };
            }

            if (f.Color is { } colour && (colour == style.Color || (style.Color is null && colour == "#000000")))
            {
                f = f with { Color = null };
            }

            if (f.LetterSpacing is { } spacing && Near(spacing, style.LetterSpacing))
            {
                f = f with { LetterSpacing = null };
            }

            if (f.Href is not null)
            {
                f = f with { Underline = false, Color = f.Color is "#0563c1" or "#0000ff" or "#1155cc" or "#0000ee" ? null : f.Color };
            }

            inline.Format = f;
        }

        if (p.Inlines.Count > 0 && p.Inlines[^1] is DocLineBreak)
        {
            p.Inlines.RemoveAt(p.Inlines.Count - 1);
        }
    }

    private static void CleanTable(DocTable table)
    {
        var look = DocsStyles.Look(table.TableStyle);
        var bodyIndex = 0;
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            foreach (var cell in row.Cells)
            {
                var expected = DocsStyles.CellCss(look, row.Header, r == 0, bodyIndex);
                foreach (var (name, value) in expected.Items)
                {
                    DropIf(cell.Css, name, v => v == value);
                }

                CleanUp(cell.Blocks);
                if (row.Header)
                {
                    foreach (var p in DocsHtml.Paragraphs(cell.Blocks))
                    {
                        foreach (var inline in p.Inlines)
                        {
                            inline.Format = inline.Format with
                            {
                                Bold = false,
                                Color = inline.Format.Color == look.HeaderText ? null : inline.Format.Color,
                            };
                        }
                    }
                }
            }

            if (!row.Header)
            {
                bodyIndex++;
            }
        }

        // Word's "Table Grid" in HTML: every cell with the same full border.
        var cells = table.Rows.SelectMany(r => r.Cells).ToList();
        var borders = cells.SelectMany(c => new[] { "border-top", "border-right", "border-bottom", "border-left" }.Select(s => c.Css.Get(s))).ToList();
        if (cells.Count > 0 && borders.All(b => b is not null) && borders.Distinct(StringComparer.Ordinal).Count() == 1
            && borders[0]!.EndsWith("solid #000000", StringComparison.Ordinal) && DocsCss.ToPt(borders[0]!.Split(' ')[0]) <= 1.5)
        {
            table.TableStyle = "grid";
            foreach (var cell in cells)
            {
                foreach (var side in new[] { "border-top", "border-right", "border-bottom", "border-left" })
                {
                    cell.Css.Remove(side);
                }
            }
        }
    }

    private static void DropIf(CssList css, string name, Func<string, bool> equal)
    {
        if (css.Get(name) is { } value && equal(value))
        {
            css.Remove(name);
        }
    }

    private static bool Near(string value, double pt) => DocsCss.ToPt(value) is { } v && Math.Abs(v - pt) < 0.1;

    // ------------------------------------------------------------------ export

    public static byte[] Write(DocsDocument document, string name) => Encoding.UTF8.GetBytes(ToHtml(document, name));

    public static string ToHtml(DocsDocument document, string name)
    {
        var page = document.Page;
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html>\n<html lang=\"").Append(DocsHtml.EscapeAttribute(document.Language)).Append("\">\n<head>\n<meta charset=\"utf-8\">\n");
        sb.Append("<meta name=\"generator\" content=\"MatOffice Docs\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
        sb.Append("<title>");
        DocsHtml.EscapeText(sb, name);
        sb.Append("</title>\n<style>\n");
        sb.Append(CultureInfo.InvariantCulture, $"@page {{ size: {Mm(page.PaperWidth)} {Mm(page.PaperHeight)}; margin: {Mm(page.Top)} {Mm(page.Right)} {Mm(page.Bottom)} {Mm(page.Left)}; }}\n");
        sb.Append(CultureInfo.InvariantCulture, $"body {{ max-width: {Mm(page.TextWidth)}; margin: 2em auto; padding: 0 1em; font-family: {DocsCss.FontStack(DocsStyles.BodyFont)}; font-size: 11pt; line-height: 1.3; }}\n");
        sb.Append("table { border-collapse: collapse; }\ntd, th { padding: 0 5.4pt; vertical-align: top; text-align: left; }\nimg { max-width: 100%; }\n");
        sb.Append("@media print { body { max-width: none; margin: 0; padding: 0; } }\n</style>\n</head>\n<body>\n");
        var exporter = new Exporter(sb, document, name);
        exporter.Blocks(document.Body);
        exporter.Notes();
        sb.Append("</body>\n</html>\n");
        return sb.ToString();
    }

    private static string Mm(double mm) => CssList.Num(mm) + "mm";

    private sealed class Exporter(StringBuilder sb, DocsDocument document, string name)
    {
        private readonly Dictionary<string, int> _notes = DocsConvert.NoteNumbers(document);

        private readonly HashSet<string> _referenced = new(StringComparer.Ordinal);

        public void Blocks(IReadOnlyList<DocBlock> blocks)
        {
            var i = 0;
            while (i < blocks.Count)
            {
                if (blocks[i] is DocParagraph { List: not null })
                {
                    i = ListRun(blocks, i);
                    continue;
                }

                Block(blocks[i]);
                i++;
            }
        }

        private void Block(DocBlock block)
        {
            switch (block)
            {
                case DocParagraph p:
                    Paragraph(p, ParagraphCss(p, list: false));
                    break;
                case DocTable table:
                    Table(table);
                    break;
                case DocRule:
                    sb.Append("<hr>\n");
                    break;
                case DocBreak { Kind: "page" }:
                    sb.Append("<div style=\"page-break-after: always; break-after: page\"></div>\n");
                    break;
            }
        }

        private static CssList ParagraphCss(DocParagraph p, bool list)
        {
            var css = DocsStyles.Get(p.StyleKey).ToCss();
            foreach (var (k, v) in p.Css.Items)
            {
                css.Set(k, v);
            }

            if (list)
            {
                css.Remove("margin-left");
                css.Remove("text-indent");
                css.Set("margin-top", "0pt");
                css.Set("margin-bottom", "0pt");
            }

            css.Set("white-space", "pre-wrap");
            return css;
        }

        private void Paragraph(DocParagraph p, CssList css, string? tag = null)
        {
            tag ??= p.Tag;
            sb.Append('<').Append(tag);
            if (p.Id is not null)
            {
                sb.Append(" id=\"").Append(p.Id).Append('"');
            }

            if (ClassOfStyle.TryGetValue(p.StyleKey, out var cls) && p.List is null)
            {
                sb.Append(" class=\"").Append(cls).Append('"');
            }

            sb.Append(" style=\"").Append(DocsHtml.EscapeAttribute(css.ToString())).Append("\">");
            Inlines(p);
            sb.Append("</").Append(tag).Append(">\n");
        }

        private void Inlines(DocParagraph p)
        {
            var inlines = p.Inlines.ToList();
            foreach (var inline in inlines)
            {
                inline.Format = inline.Format with { Comment = null };
            }

            if (inlines.Count == 0 || (inlines.Count == 1 && inlines[0] is DocLineBreak))
            {
                sb.Append("<br>");
                return;
            }

            DocsHtml.WriteInlines(sb, inlines, Replace);
        }

        private string? Replace(DocInline inline)
        {
            switch (inline)
            {
                case DocNoteRef r:
                    if (!_notes.TryGetValue(r.Id, out var n))
                    {
                        return string.Empty;
                    }

                    var id = _referenced.Add(r.Id) ? $" id=\"fnref-{r.Id}\"" : string.Empty;
                    return $"<sup data-fn=\"{r.Id}\"><a href=\"#fn-{r.Id}\"{id}>{n}</a></sup>";
                case DocField f:
                    return DocsHtml.EscapeText(DocsConvert.FieldText(f.Kind, name, document.Language));
                case DocImage img:
                    var css = new CssList();
                    css.Set("width", img.Width is { } w ? CssList.Num(w) + "px" : null);
                    css.Set("height", img.Height is { } h ? CssList.Num(h) + "px" : null);
                    switch (img.Align)
                    {
                        case "left":
                            css.Set("float", "left");
                            css.Set("margin", "0 9pt 6pt 0");
                            break;
                        case "right":
                            css.Set("float", "right");
                            css.Set("margin", "0 0 6pt 9pt");
                            break;
                        case "center":
                            css.Set("display", "block");
                            css.Set("margin", "6pt auto");
                            break;
                    }

                    var style = css.Count > 0 ? $" style=\"{css}\"" : string.Empty;
                    return $"<img src=\"{img.Src}\" alt=\"{DocsHtml.EscapeAttribute(img.Alt)}\"{style}>";
                default:
                    return null;
            }
        }

        private sealed class OpenList(int level, string kind)
        {
            public int Level { get; } = level;

            public string Kind { get; } = kind;

            public bool ItemOpen { get; set; }
        }

        /// <summary>A run of list paragraphs as nested lists; returns the index after the run.</summary>
        private int ListRun(IReadOnlyList<DocBlock> blocks, int start)
        {
            var stack = new List<OpenList>();
            var i = start;
            for (; i < blocks.Count && blocks[i] is DocParagraph { List: not null } p; i++)
            {
                var level = Math.Clamp(p.Level, 0, 8);
                while (stack.Count > 0 && stack[^1].Level > level)
                {
                    Close(stack);
                }

                if (stack.Count > 0 && stack[^1].Level == level)
                {
                    CloseItem(stack[^1]);
                    if (stack[^1].Kind != p.List || p.Start is not null)
                    {
                        Close(stack);
                    }
                }

                while (stack.Count == 0 || stack[^1].Level < level)
                {
                    var next = stack.Count == 0 ? (level > 0 ? 0 : level) : stack[^1].Level + 1;
                    if (stack.Count > 0 && !stack[^1].ItemOpen)
                    {
                        sb.Append("<li style=\"list-style-type: none\">");
                        stack[^1].ItemOpen = true;
                    }

                    Open(stack, next, p.List!, next == level ? p.Start : null);
                }

                var css = ParagraphCss(p, list: true);
                if (p.Tag == "p")
                {
                    sb.Append("<li");
                    if (p.Id is not null)
                    {
                        sb.Append(" id=\"").Append(p.Id).Append('"');
                    }

                    sb.Append(" style=\"").Append(DocsHtml.EscapeAttribute(css.ToString())).Append("\">");
                    Inlines(p);
                }
                else
                {
                    sb.Append("<li>");
                    Paragraph(p, css);
                }

                stack[^1].ItemOpen = true;
            }

            while (stack.Count > 0)
            {
                Close(stack);
            }

            return i;
        }

        private void Open(List<OpenList> stack, int level, string kind, int? start)
        {
            var format = DocsStyles.LevelFormat(kind, level);
            var numbered = DocsStyles.NumberKinds.Contains(kind);
            var type = format switch
            {
                "bullet" => "disc",
                "dash" => "'– '",
                "check" => "'✓ '",
                "arrow" => "'➢ '",
                _ => format,
            };
            sb.Append(numbered ? "<ol" : "<ul");
            if (kind == "outline")
            {
                sb.Append(" data-list=\"outline\"");
            }

            if (start is { } s)
            {
                sb.Append(" start=\"").Append(s).Append('"');
            }

            sb.Append(" style=\"list-style-type: ").Append(type).Append("; margin-top: 0; margin-bottom: 0; padding-left: 36pt\">");
            stack.Add(new OpenList(level, kind));
        }

        private void CloseItem(OpenList list)
        {
            if (list.ItemOpen)
            {
                sb.Append("</li>\n");
                list.ItemOpen = false;
            }
        }

        private void Close(List<OpenList> stack)
        {
            var list = stack[^1];
            CloseItem(list);
            sb.Append(DocsStyles.NumberKinds.Contains(list.Kind) ? "</ol>" : "</ul>");
            stack.RemoveAt(stack.Count - 1);
            if (stack.Count == 0)
            {
                sb.Append('\n');
            }
        }

        private void Table(DocTable table)
        {
            var look = DocsStyles.Look(table.TableStyle);
            sb.Append("<table data-table-style=\"").Append(table.TableStyle).Append("\" style=\"border-collapse: collapse; width: ")
                .Append(table.Width ?? "100%").Append("\">\n");
            if (table.Columns.Any(c => c is not null))
            {
                sb.Append("<colgroup>");
                foreach (var column in table.Columns)
                {
                    sb.Append(column is null ? "<col>" : $"<col style=\"width: {column}\">");
                }

                sb.Append("</colgroup>\n");
            }

            var headerRows = table.Rows.TakeWhile(r => r.Header).Count();
            var bodyIndex = 0;
            for (var r = 0; r < table.Rows.Count; r++)
            {
                var row = table.Rows[r];
                if (r == 0 && headerRows > 0)
                {
                    sb.Append("<thead>\n");
                }

                if (r == headerRows)
                {
                    sb.Append("<tbody>\n");
                }

                sb.Append("<tr>");
                foreach (var cell in row.Cells)
                {
                    var tag = row.Header || cell.HeaderCell ? "th" : "td";
                    var css = new CssList();
                    css.Set("padding", "0 5.4pt");
                    css.Set("vertical-align", "top");
                    foreach (var (k, v) in DocsStyles.CellCss(look, row.Header, r == 0, bodyIndex).Items)
                    {
                        css.Set(k, v);
                    }

                    foreach (var (k, v) in cell.Css.Items)
                    {
                        css.Set(k, v);
                    }

                    if (row.Header)
                    {
                        css.Set("font-weight", "bold");
                        css.Set("color", look.HeaderText);
                    }

                    sb.Append('<').Append(tag);
                    if (cell.ColSpan > 1)
                    {
                        sb.Append(" colspan=\"").Append(cell.ColSpan).Append('"');
                    }

                    if (cell.RowSpan > 1)
                    {
                        sb.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');
                    }

                    sb.Append(" style=\"").Append(DocsHtml.EscapeAttribute(css.ToString())).Append("\">");
                    if (cell.Blocks.Count == 0)
                    {
                        sb.Append("<br>");
                    }

                    Blocks(cell.Blocks);
                    sb.Append("</").Append(tag).Append('>');
                }

                sb.Append("</tr>\n");
                if (r == headerRows - 1)
                {
                    sb.Append("</thead>\n");
                }

                if (!row.Header)
                {
                    bodyIndex++;
                }
            }

            if (table.Rows.Count > headerRows)
            {
                sb.Append("</tbody>\n");
            }

            sb.Append("</table>\n");
        }

        public void Notes()
        {
            var ordered = DocsConvert.OrderedNotes(document);
            if (ordered.Count == 0)
            {
                return;
            }

            sb.Append("<hr class=\"mo-notes-rule\">\n<ol class=\"mo-notes\" style=\"font-size: 10pt; padding-left: 24pt\">\n");
            foreach (var (note, _) in ordered)
            {
                sb.Append("<li id=\"fn-").Append(note.Id).Append("\">");
                Blocks(note.Blocks);
                sb.Append("</li>\n");
            }

            sb.Append("</ol>\n");
        }
    }

    [GeneratedRegex("""<meta[^>]+charset\s*=\s*["']?([A-Za-z0-9_\-:.]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex CharsetRegex();

    [GeneratedRegex(@"@page\s*[^{]*\{([^}]*)\}", RegexOptions.IgnoreCase)]
    private static partial Regex PageRuleRegex();

    [GeneratedRegex(@"(?:^|;|\s)size\s*:\s*([\d.]+[a-z]+)\s+([\d.]+[a-z]+)", RegexOptions.IgnoreCase)]
    private static partial Regex SizeRegex();

    [GeneratedRegex(@"(?:^|;|\s)margin\s*:\s*([^;]+)", RegexOptions.IgnoreCase)]
    private static partial Regex MarginRegex();

    [GeneratedRegex(@"[ \t\n\r\f]+")]
    private static partial Regex WhitespaceRegex();

    [GeneratedRegex(@"white-space\s*:\s*(pre|pre-wrap|break-spaces)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PreRegex();

    [GeneratedRegex(@"mso-list\s*:\s*l\d+\s+level(\d)", RegexOptions.IgnoreCase)]
    private static partial Regex MsoListRegex();

    [GeneratedRegex(@"(page-break-before|break-before)\s*:\s*(always|page)", RegexOptions.IgnoreCase)]
    private static partial Regex BreakBeforeRegex();

    [GeneratedRegex(@"(page-break-after|break-after)\s*:\s*(always|page)", RegexOptions.IgnoreCase)]
    private static partial Regex BreakAfterRegex();

    [GeneratedRegex(@"border\s*:\s*(none|0)", RegexOptions.IgnoreCase)]
    private static partial Regex BorderNoneRegex();
}
