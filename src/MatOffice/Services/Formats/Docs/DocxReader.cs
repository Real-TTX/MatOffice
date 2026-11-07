using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// Reads Word documents (<c>.docx</c>), also real ones: styles by their (English) names along the basedOn chains, formatting that differs from the
/// Docs style as direct formatting, numbering with its real numbers, tables with merged cells, pictures, links and bookmarks, fields, content
/// controls, tracked changes (inserted text stays, deleted text goes), notes, comments with replies, headers, footers and the page.
/// Anything unknown is skipped; a file that cannot be read throws <see cref="FormatException"/>.
/// </summary>
internal sealed partial class DocxReader
{
    private const int MaxDepth = 40;

    private const long MaxImageBytes = 15L * 1024 * 1024;

    private readonly MainDocumentPart _main;
    private readonly DocsDocument _document = new();
    private readonly StyleSheet _styles;
    private readonly WordNumbering _numbering;
    private readonly HashSet<string> _anchors = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _bookmarks = new(StringComparer.Ordinal);
    private readonly HashSet<string> _ids = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _noteIds = new(StringComparer.Ordinal);
    private readonly List<(string Key, string Id)> _notesToRead = [];
    private readonly HashSet<string> _commentIds = new(StringComparer.Ordinal);
    private readonly Dictionary<DocParagraph, int> _wordNumbers = new(ReferenceEqualityComparer.Instance);
    private readonly List<string?> _sectionTypes = [];
    private readonly List<string> _pendingBookmarks = [];
    private int _sectionIndex;

    private DocxReader(WordprocessingDocument package)
    {
        _main = package.MainDocumentPart ?? throw new FormatException(Localization.Text.T("Docs.Import.Damaged"));
        var theme = _main.ThemePart?.Theme?.ThemeElements?.FontScheme;
        _styles = new StyleSheet(_main.StyleDefinitionsPart?.Styles, theme?.MajorFont?.LatinFont?.Typeface?.Value, theme?.MinorFont?.LatinFont?.Typeface?.Value);
        _numbering = new WordNumbering(_main.NumberingDefinitionsPart?.Numbering, _styles);
    }

    public static DocsDocument Read(Stream input)
    {
        var bytes = DocsConvert.ReadAll(input);
        try
        {
            using var package = WordprocessingDocument.Open(new MemoryStream(bytes, false), false, new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = 100_000_000,
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.NoProcess, FileFormatVersions.Microsoft365),
            });
            return new DocxReader(package).ReadDocument();
        }
        catch (FormatException)
        {
            throw;
        }
        catch (Exception ex) when (ex is not (OutOfMemoryException or InsufficientExecutionStackException))
        {
            throw new FormatException(Localization.Text.T("Docs.Import.Damaged"));
        }
    }

    private DocsDocument ReadDocument()
    {
        var body = _main.Document?.Body ?? throw new FormatException(Localization.Text.T("Docs.Import.Damaged"));
        OpenXmlElement?[] roots =
        [
            body, _main.FootnotesPart?.Footnotes, _main.EndnotesPart?.Endnotes,
            .. _main.HeaderParts.Select(h => (OpenXmlElement?)h.Header), .. _main.FooterParts.Select(f => (OpenXmlElement?)f.Footer),
        ];
        foreach (var root in roots.OfType<OpenXmlElement>())
        {
            CollectAnchors(root);
        }

        var sections = body.Descendants<SectionProperties>().ToList();
        _sectionTypes.AddRange(sections.Select(s => s.GetFirstChild<SectionType>()?.Val?.InnerText));
        ReadComments();

        var story = new Story(_main);
        ReadBlocks(body.ChildElements, _document.Body, story, 0);
        if (_document.Body.Count == 0)
        {
            _document.Body.Add(new DocParagraph());
        }

        var final = body.Elements<SectionProperties>().LastOrDefault() ?? sections.LastOrDefault();
        ReadSection(final, sections);
        ReadNotes();
        ResolveAnchors();
        if (_styles.Language is { } language)
        {
            _document.Language = language;
        }

        _document.Spellcheck = !StyleSheet.OnOff(_main.DocumentSettingsPart?.Settings?.GetFirstChild<HideSpellingErrors>());
        return _document;
    }

    /// <summary>The bookmarks the links of a story point to (<c>w:hyperlink w:anchor</c>, <c>HYPERLINK \l</c>): only they become ids.</summary>
    private void CollectAnchors(OpenXmlElement story)
    {
        foreach (var link in story.Descendants<Hyperlink>())
        {
            if (link.Anchor?.Value is { Length: > 0 } anchor)
            {
                _anchors.Add(anchor);
            }
        }

        foreach (var code in story.Descendants<FieldCode>().Select(f => f.Text).Concat(story.Descendants<SimpleField>().Select(f => f.Instruction?.Value)))
        {
            if (code is not null && AnchorInFieldRegex().Match(code) is { Success: true } m)
            {
                _anchors.Add(m.Groups[1].Value);
            }
        }
    }

    // ------------------------------------------------------------------ stories

    /// <summary>Where blocks come from: the part (for pictures and links), header/footer or note, the table style around, the open fields and comments.</summary>
    private sealed class Story(OpenXmlPart part)
    {
        public OpenXmlPart Part { get; } = part;

        public bool HeaderFooter { get; init; }

        public bool Note { get; init; }

        public StyleSheet.Resolved? TableStyle { get; init; }

        public List<FieldState> Fields { get; init; } = [];

        public List<string> Comments { get; init; } = [];

        public Story InTable(StyleSheet.Resolved? tableStyle) => new(Part)
        {
            HeaderFooter = HeaderFooter, Note = Note, TableStyle = tableStyle, Fields = Fields, Comments = Comments,
        };
    }

    private sealed class FieldState
    {
        public StringBuilder Instruction { get; } = new();

        public bool InResult { get; set; }

        /// <summary>The result is not text of the document (a page number became a field).</summary>
        public bool Hidden { get; set; }

        public string? Kind { get; set; }

        public string? Href { get; set; }
    }

    private void ReadBlocks(IEnumerable<OpenXmlElement> elements, List<DocBlock> output, Story story, int depth)
    {
        if (depth > MaxDepth)
        {
            return;
        }

        foreach (var element in elements)
        {
            switch (element)
            {
                case Paragraph p:
                    ReadParagraph(p, output, story, depth);
                    break;
                case Table t:
                    if (ReadTable(t, story, depth) is { } table)
                    {
                        output.Add(table);
                    }

                    break;
                case SdtBlock sdt:
                    ReadBlocks(sdt.SdtContentBlock?.ChildElements ?? Enumerable.Empty<OpenXmlElement>(), output, story, depth + 1);
                    break;
                case CustomXmlBlock custom:
                    ReadBlocks(custom.ChildElements, output, story, depth + 1);
                    break;
                case BookmarkStart bookmark:
                    Bookmark(bookmark, null);
                    break;
                case CommentRangeStart start:
                    CommentStart(start.Id?.Value, story);
                    break;
                case CommentRangeEnd end:
                    story.Comments.Remove("c" + end.Id?.Value);
                    break;
                case AltChunk chunk:
                    AltChunk(chunk, output, story);
                    break;
                case InsertedRun or MoveToRun:
                    ReadBlocks(element.ChildElements, output, story, depth + 1);
                    break;
                case AlternateContent alternate when Choose(alternate) is { } chosen:
                    ReadBlocks(chosen.ChildElements, output, story, depth + 1);
                    break;
            }
        }

        NumberLists(output);
    }

    /// <summary>An embedded HTML part (<c>w:altChunk</c>) is read like an HTML file.</summary>
    private static void AltChunk(AltChunk chunk, List<DocBlock> output, Story story)
    {
        if (chunk.Id?.Value is not { } id || story.Part.GetPartById(id) is not AlternativeFormatImportPart part
            || !(part.ContentType.Contains("html", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        using var stream = part.GetStream();
        var imported = HtmlFormat.Read(stream);
        output.AddRange(imported.Body);
    }

    /// <summary>The branch of <c>mc:AlternateContent</c> to read: the choice when it holds a drawing, else the fallback.</summary>
    private static OpenXmlElement? Choose(AlternateContent alternate)
    {
        var choice = alternate.GetFirstChild<AlternateContentChoice>();
        var fallback = alternate.GetFirstChild<AlternateContentFallback>();
        if (choice is not null && (choice.Descendants<Drawing>().Any() || fallback is null))
        {
            return choice;
        }

        return fallback;
    }

    // ------------------------------------------------------------------ paragraphs

    private sealed class ParagraphContext
    {
        public required Story Story { get; init; }

        public required ParagraphStyle Style { get; init; }

        public required StyleSheet.Props Base { get; init; }

        public required List<DocBlock> Output { get; init; }

        public required DocParagraph Shell { get; init; }

        public DocParagraph Current { get; set; } = null!;

        /// <summary>The parts of the paragraph in the output (a page break inside a paragraph cuts it in two).</summary>
        public List<DocParagraph> Segments { get; } = [];

        /// <summary>A page or column break was inside the paragraph.</summary>
        public bool Broken { get; set; }

        public string? Href { get; set; }

        public List<DocBlock> After { get; } = [];

        public int Depth { get; init; }
    }

    private void ReadParagraph(Paragraph paragraph, List<DocBlock> output, Story story, int depth)
    {
        var pPr = paragraph.ParagraphProperties;
        var resolved = _styles.Paragraph(pPr?.ParagraphStyleId?.Val?.Value, story.TableStyle);
        var props = resolved.Props.Clone();
        if (pPr is not null)
        {
            props.MergeParagraph(pPr);
        }

        var key = resolved.Key;
        if (story.Note && key == "Normal")
        {
            key = "Note";
        }

        var style = DocsStyles.Get(key);
        var shell = new DocParagraph();
        DocsStyles.Apply(shell, key);
        string? kind = null;
        var level = 0;
        var ilvl = props.NumId is > 0 ? props.Ilvl ?? _numbering.LevelOf(props.NumId.Value, props.NumStyleId) ?? 0 : 0;
        if (props.NumId is > 0 && _numbering.Kind(props.NumId.Value, ilvl) is { } listKind)
        {
            kind = listKind;
            level = Math.Clamp(ilvl, 0, 8);
            if (shell.Tag == "p" || shell.HeadingLevel > 0)
            {
                shell.List = kind;
                shell.Level = level;
                if (shell.Style == "ListParagraph")
                {
                    shell.Style = null;
                    style = DocsStyles.Get("ListParagraph");
                }
            }
        }

        shell.Css = ParagraphCss(props, style, shell.List is not null);
        var context = new ParagraphContext { Story = story, Style = style, Base = props, Output = output, Shell = shell, Depth = depth };
        StartSegment(context, first: true);
        if (props.PageBreakBefore == true && output.Count > 1)
        {
            output.Insert(output.Count - 1, new DocBreak("page"));
        }

        foreach (var name in _pendingBookmarks)
        {
            Bookmark(name, context.Current);
        }

        _pendingBookmarks.Clear();
        Inline(paragraph.ChildElements, context, depth);
        EndParagraph(context);

        // The code of a field ends in its paragraph (its result may go on, like a table of contents): a field of a damaged file that never
        // reaches its result must not hide the rest of the document.
        story.Fields.RemoveAll(f => !f.InResult);
        if (shell.List is not null && props.NumId is { } numId && _numbering.Number(numId, level) is { } number && DocsStyles.NumberKinds.Contains(kind!))
        {
            _wordNumbers[context.Current] = number;
        }

        if (context.Segments is [{ Inlines.Count: 0, List: null, Tag: "p", Id: null } empty] && IsRule(empty.Css) && output[^1] == empty)
        {
            // An empty paragraph with only a line below it: Word's way to draw a horizontal line.
            output[^1] = new DocRule();
        }

        output.AddRange(context.After);
        if (pPr?.SectionProperties is not null)
        {
            var next = _sectionIndex + 1 < _sectionTypes.Count ? _sectionTypes[_sectionIndex + 1] : null;
            if (next is not ("continuous" or "nextColumn"))
            {
                output.Add(new DocBreak("page"));
            }

            _sectionIndex++;
        }
    }

    private static void StartSegment(ParagraphContext context, bool first)
    {
        var p = context.Shell.CloneShell();
        if (!first)
        {
            p.List = null;
            p.Start = null;
        }

        context.Current = p;
        context.Segments.Add(p);
        context.Output.Add(p);
    }

    /// <summary>
    /// A page break in a paragraph: the text before it stays a paragraph, the break becomes a block, the text after it a paragraph of its own.
    /// A part without text is dropped (Word shows the mark of a paragraph that holds only a break next to the break, not on the next page).
    /// </summary>
    private static void BreakParagraph(ParagraphContext context, string kind)
    {
        var before = context.Current;
        if (before.Inlines.Count == 0)
        {
            context.Output.Remove(before);
            context.Segments.Remove(before);
        }

        context.Broken = true;
        context.Output.Add(new DocBreak(kind));
        StartSegment(context, first: false);
        if (before.Inlines.Count == 0)
        {
            context.Current.Id = before.Id;
        }
    }

    private static void EndParagraph(ParagraphContext context)
    {
        var p = context.Current;
        if (p.Inlines.Count == 0 && context.Broken)
        {
            context.Output.Remove(p);
            context.Segments.Remove(p);
            context.Current = context.Segments.LastOrDefault() ?? p;
        }
    }

    /// <summary>The formatting of a horizontal line: a border below, none around.</summary>
    private static bool IsRule(CssList css)
        => css.Has("border-bottom") && !css.Has("border-top") && !css.Has("border-left") && !css.Has("border-right") && !css.Has("background-color");

    private void Inline(IEnumerable<OpenXmlElement> elements, ParagraphContext context, int depth)
    {
        if (depth > MaxDepth)
        {
            return;
        }

        foreach (var element in elements)
        {
            switch (element)
            {
                case Run run:
                    ReadRun(run, context, depth);
                    break;
                case Hyperlink link:
                    var previous = context.Href;
                    context.Href = Href(link, context.Story) ?? previous;
                    Inline(link.ChildElements, context, depth + 1);
                    context.Href = previous;
                    break;
                case SimpleField field:
                    SimpleField(field, context, depth);
                    break;
                case InsertedRun or MoveToRun or CustomXmlRun or BidirectionalOverride or BidirectionalEmbedding or OpenXmlUnknownElement { LocalName: "smartTag" or "customXml" or "ins" or "moveTo" }:
                    Inline(element.ChildElements, context, depth + 1);
                    break;
                case SdtRun sdt:
                    Inline(sdt.SdtContentRun?.ChildElements ?? Enumerable.Empty<OpenXmlElement>(), context, depth + 1);
                    break;
                case BookmarkStart bookmark:
                    Bookmark(bookmark, context.Current);
                    break;
                case CommentRangeStart start:
                    CommentStart(start.Id?.Value, context.Story);
                    break;
                case CommentRangeEnd end:
                    context.Story.Comments.Remove("c" + end.Id?.Value);
                    break;
                case M.OfficeMath or M.Paragraph:
                    var math = string.Concat(element.Descendants<M.Text>().Select(t => t.Text));
                    AddText(context, math, Format(context, RunProps(context, null)));
                    break;
                case AlternateContent alternate when Choose(alternate) is { } chosen:
                    Inline(chosen.ChildElements, context, depth + 1);
                    break;
            }
        }
    }

    private void SimpleField(SimpleField field, ParagraphContext context, int depth)
    {
        var instruction = field.Instruction?.Value ?? string.Empty;
        var (kind, href) = FieldKind(instruction, context.Story);
        if (kind is not null && context.Story.HeaderFooter)
        {
            var rPr = field.Descendants<RunProperties>().FirstOrDefault();
            context.Current.Inlines.Add(new DocField(kind) { Format = Format(context, RunProps(context, rPr)) });
            return;
        }

        var previous = context.Href;
        context.Href = href ?? previous;
        Inline(field.ChildElements, context, depth + 1);
        context.Href = previous;
    }

    /// <summary>The field kinds Docs has (only in headers and footers) and the target of a HYPERLINK field.</summary>
    private (string? Kind, string? Href) FieldKind(string instruction, Story story)
    {
        var words = instruction.Trim().Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            return (null, null);
        }

        var kind = words[0].ToUpperInvariant() switch
        {
            "PAGE" => "page",
            "NUMPAGES" or "SECTIONPAGES" => "pages",
            "DATE" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => "date",
            "TIME" => "time",
            "TITLE" => "title",
            "DOCPROPERTY" when words.Length > 1 && words[1].Trim('"').Equals("Title", StringComparison.OrdinalIgnoreCase) => "title",
            _ => null,
        };
        if (words[0].Equals("HYPERLINK", StringComparison.OrdinalIgnoreCase))
        {
            var target = HyperlinkFieldRegex().Match(instruction);
            if (target.Success)
            {
                var value = target.Groups[2].Value;
                return (null, target.Groups[1].Success && target.Groups[1].Value.Length > 0 ? "#" + value : value);
            }
        }

        return (kind, null);
    }

    private void ReadRun(Run run, ParagraphContext context, int depth)
    {
        var rPr = run.RunProperties;
        var props = RunProps(context, rPr);
        var hidden = props.Vanish == true;
        var format = Format(context, props);
        foreach (var child in run.ChildElements)
        {
            if (child is FieldChar fieldChar)
            {
                FieldChar(fieldChar, context, format);
                continue;
            }

            if (child is FieldCode code)
            {
                if (context.Story.Fields.Count > 0 && !context.Story.Fields[^1].InResult)
                {
                    context.Story.Fields[^1].Instruction.Append(code.Text);
                }

                continue;
            }

            if (hidden || context.Story.Fields.Any(f => !f.InResult || f.Hidden))
            {
                continue;
            }

            var f = format;
            if (context.Story.Fields.LastOrDefault(x => x.Href is not null) is { } fieldLink)
            {
                f = f with { Href = DocsHtml.SafeHref(fieldLink.Href) ?? f.Href };
            }

            switch (child)
            {
                case Text text:
                    AddText(context, text.Text, f);
                    break;
                case TabChar or PositionalTab:
                    AddText(context, "\t", f);
                    break;
                case Break br when br.Type?.InnerText is "page" or "column":
                    BreakParagraph(context, br.Type!.InnerText);
                    break;
                case Break or CarriageReturn:
                    LineBreak(context, f);
                    break;
                case NoBreakHyphen:
                    AddText(context, "\u2011", f);
                    break;
                case SoftHyphen:
                    AddText(context, "\u00AD", f);
                    break;
                case SymbolChar symbol:
                    AddText(context, Symbol(symbol), f);
                    break;
                case Drawing drawing:
                    Drawing(drawing, context, f, depth);
                    break;
                case Picture picture:
                    Vml(picture, context, f, depth);
                    break;
                case EmbeddedObject embedded:
                    Vml(embedded, context, f, depth);
                    break;
                case FootnoteReference footnote when footnote.Id?.Value is { } id && !context.Story.Note:
                    context.Current.Inlines.Add(new DocNoteRef(NoteId("f" + id.ToString(CultureInfo.InvariantCulture))) { Format = f with { Script = 0 } });
                    break;
                case EndnoteReference endnote when endnote.Id?.Value is { } id && !context.Story.Note:
                    context.Current.Inlines.Add(new DocNoteRef(NoteId("e" + id.ToString(CultureInfo.InvariantCulture))) { Format = f with { Script = 0 } });
                    break;
                case Ruby ruby:
                    var baseText = string.Concat(ruby.RubyBase?.Descendants<Text>().Select(t => t.Text) ?? []);
                    AddText(context, baseText, f);
                    break;
                case AlternateContent alternate when Choose(alternate) is { } chosen:
                    var copy = new Run();
                    if (rPr is not null)
                    {
                        copy.Append(rPr.CloneNode(true));
                    }

                    copy.Append(chosen.ChildElements.Select(c => c.CloneNode(true)));
                    ReadRun(copy, context, depth + 1);
                    break;
            }
        }
    }

    private void FieldChar(FieldChar fieldChar, ParagraphContext context, RunFormat format)
    {
        var fields = context.Story.Fields;
        switch (fieldChar.FieldCharType?.InnerText)
        {
            case "begin":
                if (fields.Count < 20)
                {
                    fields.Add(new FieldState());
                }

                break;
            case "separate" when fields.Count > 0:
                Separate(fields[^1], context, format);
                break;
            case "end" when fields.Count > 0:
                var field = fields[^1];
                if (!field.InResult)
                {
                    Separate(field, context, format);
                }

                fields.RemoveAt(fields.Count - 1);
                break;
        }
    }

    private void Separate(FieldState field, ParagraphContext context, RunFormat format)
    {
        field.InResult = true;
        var (kind, href) = FieldKind(field.Instruction.ToString(), context.Story);
        field.Href = href;
        if (kind is not null && context.Story.HeaderFooter && !context.Story.Fields.Take(context.Story.Fields.Count - 1).Any(f => f.Hidden || !f.InResult))
        {
            context.Current.Inlines.Add(new DocField(kind) { Format = format });
            field.Hidden = true;
        }
    }

    private static string Symbol(SymbolChar symbol)
    {
        if (!int.TryParse(symbol.Char?.InnerText, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code))
        {
            return string.Empty;
        }

        if (code >= 0xF000)
        {
            code -= 0xF000;
        }

        var font = symbol.Font?.Value ?? string.Empty;
        if (font.Equals("Symbol", StringComparison.OrdinalIgnoreCase))
        {
            return code switch
            {
                0xB7 => "•", 0xD7 => "×", 0xB1 => "±", 0xB4 => "×", 0xB8 => "÷", 0xA3 => "≤", 0xB3 => "≥", 0xB9 => "≠", 0xA5 => "∞", 0xB0 => "°",
                0xAE => "→", 0xAC => "←", 0xAD => "↑", 0xAF => "↓", 0xDE => "⇒", 0xE3 => "©", 0xE2 => "®", 0xE4 => "™",
                >= 0x41 and <= 0x5A or >= 0x61 and <= 0x7A => Greek(code),
                _ => ((char)code).ToString(),
            };
        }

        if (font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase))
        {
            return code switch
            {
                0xFC => "✓", 0xFB => "✗", 0xFE => "☑", 0xA8 => "☐", 0xD8 => "➢", 0xE0 => "➔", 0xA7 => "▪", 0x6C => "●", 0x6E => "■", 0x4A => "☺",
                _ => "•",
            };
        }

        return code is >= 0x20 and < 0xD800 ? ((char)code).ToString() : string.Empty;
    }

    private static string Greek(int code)
    {
        const string upper = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ";
        const string lower = "αβχδεφγηιϕκλμνοπθρστυϖωξψζ";
        return code <= 0x5A ? upper[code - 0x41].ToString() : lower[code - 0x61].ToString();
    }

    /// <summary>A line break: <c>br</c>, or a new line in the text of code (<c>pre</c> keeps its lines as text).</summary>
    private static void LineBreak(ParagraphContext context, RunFormat format)
    {
        if (context.Shell.Tag == "pre")
        {
            AddText(context, "\n", format);
        }
        else
        {
            context.Current.Inlines.Add(new DocLineBreak { Format = format with { Script = 0 } });
        }
    }

    private static void AddText(ParagraphContext context, string text, RunFormat format)
    {
        text = DocsHtml.CleanText(text);
        if (text.Length == 0)
        {
            return;
        }

        var inlines = context.Current.Inlines;
        if (inlines.Count > 0 && inlines[^1] is DocText last && last.Format == format)
        {
            last.Text += text;
        }
        else
        {
            inlines.Add(new DocText(text) { Format = format });
        }
    }

    // ------------------------------------------------------------------ formatting

    private StyleSheet.Props RunProps(ParagraphContext context, RunProperties? rPr)
    {
        var props = context.Base.Clone();
        var characterStyle = rPr?.RunStyle?.Val?.Value;
        if (characterStyle is not null && _styles.Character(characterStyle) is { } style && !style.Name.Contains("hyperlink", StringComparison.OrdinalIgnoreCase))
        {
            props.MergeCharacterStyle(style.Props);
        }

        if (rPr is not null)
        {
            props.MergeRun(rPr, _styles);
        }

        return props;
    }

    /// <summary>The run formatting where it differs from the paragraph style of Docs.</summary>
    private RunFormat Format(ParagraphContext context, StyleSheet.Props props)
    {
        var style = context.Style;
        var f = RunFormat.Plain with
        {
            Href = context.Href is { } href ? DocsHtml.SafeHref(href.StartsWith('#') ? href : href) ?? (href.StartsWith('#') ? href : null) : null,
            Comment = context.Story.Comments.Count > 0 ? string.Join(' ', context.Story.Comments.Distinct()) : null,
            Bold = props.Bold == true,
            Italic = props.Italic == true && !style.Italic,
            Underline = props.Underline == true,
            Strike = props.Strike == true,
            Script = props.Script ?? 0,
        };
        if (props.Font is { } font && !font.Equals(style.Font, StringComparison.OrdinalIgnoreCase))
        {
            f = f with { FontFamily = DocsCss.FontStack(font) };
        }

        var size = props.Size ?? 10;
        if (Math.Abs(size - style.Size) > 0.01 && size > 0)
        {
            f = f with { FontSize = CssList.Pt(size) };
        }

        var colour = props.Color is { } c && c != "auto" ? DocsCss.Colour("#" + c) : null;
        if (colour is not null && colour != style.Color && !(style.Color is null && colour == "#000000"))
        {
            f = f with { Color = colour };
        }
        else if (colour is null && style.Color is not null)
        {
            f = f with { Color = "#000000" };
        }

        var background = props.Highlight is { Length: > 0 } highlight ? Highlight(highlight) : null;
        background ??= props.Shading is { Length: > 0 } shading ? DocsCss.Colour("#" + shading) : null;
        if (background is not null)
        {
            f = f with { Background = background };
        }

        var spacing = props.Spacing ?? 0;
        if (Math.Abs(spacing - style.LetterSpacing) > 0.01)
        {
            f = f with { LetterSpacing = CssList.Pt(spacing) };
        }

        if (f.Href is not null)
        {
            f = f with { Underline = false, Color = f.Color is "#0563c1" or "#0000ff" or "#000000" ? null : f.Color };
        }

        return f;
    }

    private static string? Highlight(string name) => name switch
    {
        "yellow" => "#ffff00",
        "green" => "#00ff00",
        "cyan" => "#00ffff",
        "magenta" => "#ff00ff",
        "blue" => "#0000ff",
        "red" => "#ff0000",
        "darkBlue" => "#000080",
        "darkCyan" => "#008080",
        "darkGreen" => "#008000",
        "darkMagenta" => "#800080",
        "darkRed" => "#800000",
        "darkYellow" => "#808000",
        "darkGray" => "#808080",
        "lightGray" => "#c0c0c0",
        "black" => "#000000",
        "white" => "#ffffff",
        _ => null,
    };

    /// <summary>The paragraph formatting where it differs from the paragraph style of Docs.</summary>
    private static CssList ParagraphCss(StyleSheet.Props p, ParagraphStyle style, bool list)
    {
        var css = new CssList();
        var align = p.Align switch
        {
            "center" => "center",
            "right" or "end" => "right",
            "both" or "distribute" or "lowKashida" or "mediumKashida" or "highKashida" or "thaiDistribute" => "justify",
            _ => "left",
        };
        if (align != style.Align)
        {
            css.Set("text-align", align);
        }

        var line = p.LineRule is "exact" or "atLeast" && p.Line is { } exact ? CssList.Pt(exact / 20) : null;
        if (line is not null)
        {
            css.Set("line-height", line);
        }
        else
        {
            var multiple = Math.Round((p.Line ?? 240) / 240 * 1.2, 2);
            if (Math.Abs(multiple - style.Line) > 0.02)
            {
                css.Set("line-height", CssList.Num(multiple));
            }
        }

        if (Math.Abs((p.Before ?? 0) - style.Before) > 0.05)
        {
            css.Set("margin-top", CssList.Pt(p.Before ?? 0));
        }

        if (Math.Abs((p.After ?? 0) - style.After) > 0.05)
        {
            css.Set("margin-bottom", CssList.Pt(p.After ?? 0));
        }

        if (!list)
        {
            if (Math.Abs((p.Left ?? 0) - style.IndentLeft) > 0.05)
            {
                css.Set("margin-left", CssList.Pt(p.Left ?? 0));
            }

            if (Math.Abs((p.Right ?? 0) - style.IndentRight) > 0.05)
            {
                css.Set("margin-right", CssList.Pt(p.Right ?? 0));
            }

            var indent = p.Hanging is { } hanging and > 0 ? -hanging : p.FirstLine ?? 0;
            if (Math.Abs(indent) > 0.05)
            {
                css.Set("text-indent", CssList.Pt(indent));
            }
        }

        var shading = p.ParagraphShading is { Length: > 0 } fill ? DocsCss.Colour("#" + fill) : null;
        if (shading != style.Background)
        {
            css.Set("background-color", shading ?? (style.Background is null ? null : "#ffffff"));
        }

        foreach (var side in new[] { "top", "right", "bottom", "left" })
        {
            var border = p.Borders.GetValueOrDefault(side);
            var own = style.BorderColor is { } bc && side is "top" or "bottom" ? "0.5pt solid " + bc : null;
            if (border?.Css != own && !(border is null && own is null))
            {
                if (border is null || own is not null && border.Css is not null && border.Css.EndsWith(style.BorderColor!, StringComparison.Ordinal))
                {
                    continue;
                }

                css.Set("border-" + side, border.Css);
                if (border.Css is not null && border.Space > 0)
                {
                    css.Set("padding-" + side, CssList.Pt(border.Space));
                }
            }
        }

        return css;
    }

    // ------------------------------------------------------------------ links, bookmarks, comments, notes

    private static string? Href(Hyperlink link, Story story)
    {
        string? url = null;
        if (link.Id?.Value is { } id)
        {
            url = story.Part.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString;
        }

        if (link.Anchor?.Value is { Length: > 0 } anchor)
        {
            url = url is null ? "#" + anchor : url + "#" + anchor;
        }

        return url;
    }

    private void Bookmark(BookmarkStart bookmark, DocParagraph? paragraph)
    {
        if (bookmark.Name?.Value is { } name)
        {
            Bookmark(name, paragraph);
        }
    }

    /// <summary>
    /// A bookmark a link points to gives its paragraph an id (Docs keeps ids for the links to them, like the table of contents); the links are
    /// rewritten to the ids at the end.
    /// </summary>
    private void Bookmark(string name, DocParagraph? paragraph)
    {
        if (!_anchors.Contains(name))
        {
            return;
        }

        if (paragraph is null)
        {
            _pendingBookmarks.Add(name);
            return;
        }

        if (paragraph.Id is null)
        {
            var id = DocsHtml.ToId(name);
            for (var n = 2; !_ids.Add(id); n++)
            {
                id = DocsHtml.ToId(name)[..Math.Min(DocsHtml.ToId(name).Length, 58)] + "-" + n.ToString(CultureInfo.InvariantCulture);
            }

            paragraph.Id = id;
        }

        _bookmarks.TryAdd(name, paragraph.Id);
    }

    /// <summary>Links to bookmarks point to the ids of their paragraphs; a link to a bookmark that is not there (or to <c>_top</c>) is plain text.</summary>
    private void ResolveAnchors()
    {
        var blocks = _document.Body.Concat(_document.Notes.SelectMany(n => n.Blocks)).Concat(_document.Header).Concat(_document.Footer)
            .Concat(_document.FirstHeader).Concat(_document.FirstFooter);
        foreach (var p in DocsHtml.Paragraphs(blocks))
        {
            foreach (var inline in p.Inlines)
            {
                if (inline.Format.Href is { } href && href.StartsWith('#'))
                {
                    inline.Format = inline.Format with { Href = _bookmarks.GetValueOrDefault(href[1..]) is { } target ? "#" + target : null };
                }
            }
        }
    }

    private void CommentStart(string? id, Story story)
    {
        if (id is not null && _commentIds.Contains("c" + id) && !story.HeaderFooter && !story.Note)
        {
            story.Comments.Add("c" + id);
        }
    }

    private void ReadComments()
    {
        var comments = _main.WordprocessingCommentsPart?.Comments;
        if (comments is null)
        {
            return;
        }

        var extended = _main.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>().ToList() ?? [];
        var parentOf = extended.Where(e => e.ParaId?.Value is not null && e.ParaIdParent?.Value is not null)
            .GroupBy(e => e.ParaId!.Value!).ToDictionary(g => g.Key, g => g.First().ParaIdParent!.Value!, StringComparer.OrdinalIgnoreCase);
        var doneOf = extended.Where(e => e.ParaId?.Value is not null)
            .GroupBy(e => e.ParaId!.Value!).ToDictionary(g => g.Key, g => g.First().Done?.Value == true, StringComparer.OrdinalIgnoreCase);
        var byPara = new Dictionary<string, DocComment>(StringComparer.OrdinalIgnoreCase);
        var replies = new List<(DocComment Reply, string Parent)>();
        foreach (var comment in comments.Elements<Comment>())
        {
            if (comment.Id?.Value is not { } wordId || _document.Comments.Count >= DocsDocument.MaxComments)
            {
                continue;
            }

            var paragraphs = comment.Elements<Paragraph>().ToList();
            var text = string.Join('\n', paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));
            var item = new DocComment
            {
                Id = "c" + wordId,
                Author = comment.Author?.Value ?? string.Empty,
                Date = comment.Date?.HasValue == true ? comment.Date.Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture) : string.Empty,
                Text = text,
            };
            var paraId = paragraphs.LastOrDefault()?.ParagraphId?.Value;
            if (paraId is not null)
            {
                byPara[paraId] = item;
                item.Done = doneOf.GetValueOrDefault(paraId);
                if (parentOf.TryGetValue(paraId, out var parent))
                {
                    replies.Add((item, parent));
                    continue;
                }
            }

            _document.Comments.Add(item);
            _commentIds.Add(item.Id);
        }

        foreach (var (reply, parent) in replies)
        {
            if (byPara.TryGetValue(parent, out var owner) && _document.Comments.Contains(owner) && owner.Replies.Count < DocsDocument.MaxReplies)
            {
                owner.Replies.Add(reply);
            }
            else
            {
                _document.Comments.Add(reply);
                _commentIds.Add(reply.Id);
            }
        }
    }

    private string NoteId(string key)
    {
        if (!_noteIds.TryGetValue(key, out var id))
        {
            id = "n" + (_noteIds.Count + 1).ToString(CultureInfo.InvariantCulture);
            _noteIds[key] = id;
            _notesToRead.Add((key, id));
        }

        return id;
    }

    private void ReadNotes()
    {
        foreach (var (key, id) in _notesToRead.Take(DocsDocument.MaxNotes))
        {
            var number = key[1..];
            OpenXmlPart? part = key[0] == 'f' ? _main.FootnotesPart : _main.EndnotesPart;
            OpenXmlElement? note = key[0] == 'f'
                ? _main.FootnotesPart?.Footnotes?.Elements<Footnote>().FirstOrDefault(n => n.Id?.InnerText == number)
                : _main.EndnotesPart?.Endnotes?.Elements<Endnote>().FirstOrDefault(n => n.Id?.InnerText == number);
            if (part is null || note is null)
            {
                continue;
            }

            var blocks = new List<DocBlock>();
            ReadBlocks(note.ChildElements, blocks, new Story(part) { Note = true }, 1);
            if (blocks.FirstOrDefault() is DocParagraph { Inlines: [DocText first, ..] })
            {
                first.Text = first.Text.TrimStart();
            }

            _document.Notes.Add(new DocNote { Id = id, Blocks = blocks });
        }
    }

    // ------------------------------------------------------------------ pictures and text boxes

    private void Drawing(Drawing drawing, ParagraphContext context, RunFormat format, int depth)
    {
        OpenXmlElement? container = (OpenXmlElement?)drawing.Inline ?? drawing.Anchor;
        if (container is null)
        {
            return;
        }

        var blip = container.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value is { } embed)
        {
            var extent = container.GetFirstChild<DW.Extent>();
            var docPr = container.GetFirstChild<DW.DocProperties>();
            string? align = null;
            if (container is DW.Anchor anchor)
            {
                align = AnchorAlign(anchor);
            }

            AddImage(context, embed, (extent?.Cx?.Value ?? 0) / 9525.0, (extent?.Cy?.Value ?? 0) / 9525.0, docPr?.Description?.Value ?? docPr?.Title?.Value, align, format);
            return;
        }

        foreach (var box in container.Descendants<TextBoxContent>().Take(10))
        {
            ReadBlocks(box.ChildElements, context.After, context.Story, context.Depth + depth + 1);
        }
    }

    private string? AnchorAlign(DW.Anchor anchor)
    {
        var wrapped = anchor.ChildElements.Any(c => c is DW.WrapSquare or DW.WrapTight or DW.WrapThrough);
        var topBottom = anchor.ChildElements.Any(c => c is DW.WrapTopBottom);
        var position = anchor.GetFirstChild<DW.HorizontalPosition>();
        var align = position?.HorizontalAlignment?.Text;
        if (align is null && position?.PositionOffset?.Text is { } offsetText && long.TryParse(offsetText, out var offset))
        {
            var width = anchor.GetFirstChild<DW.Extent>()?.Cx?.Value ?? 0;
            var text = _document.Page.TextWidth * 36000;
            align = offset + (width / 2) > text / 2 ? "right" : "left";
        }

        if (wrapped)
        {
            return align is "right" or "outside" ? "right" : align is "center" ? "center" : "left";
        }

        return topBottom && align == "center" ? "center" : null;
    }

    private void Vml(OpenXmlElement element, ParagraphContext context, RunFormat format, int depth)
    {
        var image = element.Descendants<V.ImageData>().FirstOrDefault();
        if (image?.RelationshipId?.Value is { } id)
        {
            var shape = image.Parent;
            var style = shape?.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value ?? string.Empty;
            var css = DocsCss.Declarations(style).ToDictionary(d => d.Name.ToLowerInvariant(), d => d.Value, StringComparer.Ordinal);
            var width = DocsCss.ToPx(css.GetValueOrDefault("width")) ?? 0;
            var height = DocsCss.ToPx(css.GetValueOrDefault("height")) ?? 0;
            AddImage(context, id, width, height, image.Title?.Value, null, format);
            return;
        }

        foreach (var box in element.Descendants<TextBoxContent>().Take(10))
        {
            ReadBlocks(box.ChildElements, context.After, context.Story, context.Depth + depth + 1);
        }
    }

    private static void AddImage(ParagraphContext context, string relationshipId, double width, double height, string? alt, string? align, RunFormat format)
    {
        if (context.Story.Part.Parts.FirstOrDefault(p => p.RelationshipId == relationshipId).OpenXmlPart is not ImagePart part)
        {
            return;
        }

        var mime = part.ContentType.ToLowerInvariant() switch
        {
            "image/png" => "png",
            "image/jpeg" or "image/jpg" or "image/pjpeg" => "jpeg",
            "image/gif" => "gif",
            "image/webp" => "webp",
            _ => null,
        };
        if (mime is null)
        {
            return;
        }

        using var stream = part.GetStream();
        if (stream.CanSeek && stream.Length > MaxImageBytes)
        {
            return;
        }

        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        if (memory.Length is 0 or > MaxImageBytes)
        {
            return;
        }

        var bytes = memory.ToArray();
        if (width < 1 || height < 1)
        {
            (width, height) = ImageSize.Of(bytes) ?? (100, 100);
        }

        context.Current.Inlines.Add(new DocImage
        {
            Src = $"data:image/{mime};base64,{Convert.ToBase64String(bytes)}",
            Alt = DocsHtml.CleanText(alt ?? string.Empty).Replace('\n', ' '),
            Width = Math.Round(width, 2),
            Height = Math.Round(height, 2),
            Align = align,

            // Fonts and colours of the run mean nothing for a picture; a link and a comment do.
            Format = RunFormat.Plain with { Href = format.Href, Comment = format.Comment },
        });
    }

    // ------------------------------------------------------------------ tables

    private DocTable? ReadTable(Table table, Story story, int depth)
    {
        if (depth > MaxDepth)
        {
            return null;
        }

        var tblPr = table.GetFirstChild<TableProperties>();
        var tableStyle = _styles.Table(tblPr?.TableStyle?.Val?.Value);
        var result = new DocTable();
        var borders = new Dictionary<string, StyleSheet.Border?>(tableStyle?.TableBorders ?? new Dictionary<string, StyleSheet.Border?>());
        if (tblPr?.TableBorders is { } direct)
        {
            foreach (var (side, border) in StyleSheet.ReadBorders(direct))
            {
                borders[side] = border;
            }
        }

        var fallback = tableStyle?.Name.Equals("Table Grid", StringComparison.OrdinalIgnoreCase) == true || borders.Values.Any(b => b?.Css is not null) ? "grid" : "plain";

        var grid = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Select(c => Twips(c.Width?.Value)).ToList() ?? [];
        var gridTotal = grid.Sum();
        if (gridTotal > 0)
        {
            // A tenth of a percent is far below a millimetre: whole twips do not turn 35% into 34.99%.
            result.Columns.AddRange(grid.Select(w => (string?)(CssList.Num(Math.Round(w / gridTotal * 100, 1)) + "%")));
        }

        var textTwips = _document.Page.TextWidth * 56.6929;
        var width = tblPr?.TableWidth;
        double? percent = width?.Type?.InnerText switch
        {
            "pct" when width.Width?.Value is { } w => w.EndsWith('%') ? DocsCss.Percent(w) : Twips(w) / 50,
            "dxa" when width.Width?.Value is { } w => Twips(w) / textTwips * 100,
            _ => gridTotal > 0 ? gridTotal / textTwips * 100 : null,
        };
        result.Width = percent is > 0 ? CssList.Num(Math.Min(100, percent.Value >= 97 ? 100 : percent.Value)) + "%" : "100%";

        var rows = new List<TableRow>();
        CollectRows(table.ChildElements, rows, 0);
        if (rows.Count > 10000)
        {
            rows.RemoveRange(10000, rows.Count - 10000);
        }

        var mergeOrigin = new Dictionary<int, DocCell>();
        var header = true;
        var look = TableLookOf(tblPr?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableLook>());

        // A first row the table style shows as a header (Word's "Header Row") is the header row of Docs.
        var styledHeader = look.FirstRow && tableStyle?.Conditionals.ContainsKey("firstRow") == true;
        for (var r = 0; r < rows.Count; r++)
        {
            var tr = rows[r];
            var row = new DocRow();
            var trPr = tr.TableRowProperties;
            header &= (trPr?.GetFirstChild<TableHeader>() is { } th && StyleSheet.OnOff(th)) || (r == 0 && styledHeader);
            row.Header = header;
            var column = Int(trPr?.GetFirstChild<GridBefore>()?.Val?.InnerText) ?? 0;
            var cells = new List<TableCell>();
            CollectCells(tr.ChildElements, cells, 0);
            var columns = Math.Max(grid.Count, column + cells.Sum(c => Math.Clamp(Int(c.TableCellProperties?.GridSpan?.Val?.InnerText) ?? 1, 1, 1000)));
            foreach (var tc in cells)
            {
                var tcPr = tc.TableCellProperties;
                var span = Math.Clamp(Int(tcPr?.GridSpan?.Val?.InnerText) ?? 1, 1, 1000);
                var merge = tcPr?.VerticalMerge;
                var mergeValue = merge?.Val?.InnerText;
                if (merge is not null && mergeValue != "restart" && mergeOrigin.TryGetValue(column, out var origin))
                {
                    origin.RowSpan++;
                    column += span;
                    continue;
                }

                var cell = new DocCell { ColSpan = span, HeaderCell = header };
                var types = tableStyle is null ? [] : ConditionalTypes(look, tableStyle, r, rows.Count, row.Header, column, span, columns);
                if (tcPr?.Shading is { } shading && StyleSheet.ShadingColour(shading) is { Length: > 0 } fill)
                {
                    cell.Css.Set("background-color", DocsCss.Colour("#" + fill));
                }
                else if (tableStyle?.FillFor(types) is { } styleFill)
                {
                    cell.Css.Set("background-color", DocsCss.Colour("#" + styleFill));
                }

                if (tcPr?.TableCellVerticalAlignment?.Val?.InnerText is { } valign && valign != "top")
                {
                    cell.Css.Set("vertical-align", valign == "center" ? "middle" : "bottom");
                }

                if (tcPr?.TableCellBorders is { } cellBorders)
                {
                    foreach (var (side, border) in StyleSheet.ReadBorders(cellBorders))
                    {
                        if (border?.Css is { } css && side is "top" or "right" or "bottom" or "left")
                        {
                            cell.Css.Set("border-" + side, css);
                        }
                    }
                }

                ReadBlocks(tc.ChildElements, cell.Blocks, story.InTable(tableStyle?.ForCell(types)), depth + 1);
                if (cell.Blocks.Count == 0)
                {
                    cell.Blocks.Add(new DocParagraph());
                }

                if (mergeValue == "restart")
                {
                    for (var c = column; c < column + span; c++)
                    {
                        mergeOrigin[c] = cell;
                    }
                }
                else
                {
                    for (var c = column; c < column + span; c++)
                    {
                        mergeOrigin.Remove(c);
                    }
                }

                row.Cells.Add(cell);
                column += span;
            }

            if (row.Cells.Count > 0)
            {
                result.Rows.Add(row);
            }
        }

        if (result.Rows.Count == 0)
        {
            return null;
        }

        result.TableStyle = Look(result, borders) ?? fallback;
        HeaderText(result);
        return result;
    }

    /// <summary>Which conditional formatting of its table style a table shows (<c>w:tblLook</c>: the attributes, or the bits of the older <c>w:val</c>).</summary>
    private static (bool FirstRow, bool LastRow, bool FirstColumn, bool LastColumn, bool RowBands, bool ColumnBands) TableLookOf(OpenXmlElement? tblLook)
    {
        var attributes = tblLook?.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value, StringComparer.Ordinal) ?? [];
        var bits = attributes.TryGetValue("val", out var hex) && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var n) ? n : 0;
        bool Flag(string name, int bit) => attributes.TryGetValue(name, out var v) ? v?.Trim().ToLowerInvariant() is "1" or "true" or "on" : (bits & bit) != 0;
        return (Flag("firstRow", 0x20), Flag("lastRow", 0x40), Flag("firstColumn", 0x80), Flag("lastColumn", 0x100), !Flag("noHBand", 0x200), !Flag("noVBand", 0x400));
    }

    /// <summary>The conditional formatting types of a cell, in Word's order (the later ones win): whole table, bands, first/last column, first/last row, corners.</summary>
    private static List<string> ConditionalTypes(
        (bool FirstRow, bool LastRow, bool FirstColumn, bool LastColumn, bool RowBands, bool ColumnBands) look,
        StyleSheet.Resolved style, int row, int rows, bool headerRow, int column, int span, int columns)
    {
        var types = new List<string> { "wholeTable" };
        var firstRow = look.FirstRow && headerRow;
        var lastRow = look.LastRow && row == rows - 1 && !firstRow;
        var firstColumn = look.FirstColumn && column == 0;
        var lastColumn = look.LastColumn && column + span >= columns;
        var bandColumn = column - (look.FirstColumn ? 1 : 0);
        if (look.ColumnBands && bandColumn >= 0 && !firstColumn && !lastColumn)
        {
            types.Add(bandColumn / style.ColumnBand % 2 == 0 ? "band1Vert" : "band2Vert");
        }

        var bandRow = row - (look.FirstRow ? 1 : 0);
        if (look.RowBands && bandRow >= 0 && !firstRow && !lastRow)
        {
            types.Add(bandRow / style.RowBand % 2 == 0 ? "band1Horz" : "band2Horz");
        }

        if (firstColumn)
        {
            types.Add("firstCol");
        }

        if (lastColumn)
        {
            types.Add("lastCol");
        }

        if (firstRow)
        {
            types.Add("firstRow");
        }

        if (lastRow)
        {
            types.Add("lastRow");
        }

        if (firstRow || lastRow)
        {
            var corner = (firstRow ? "n" : "s") + (firstColumn ? "w" : lastColumn ? "e" : string.Empty);
            if (corner.Length == 2)
            {
                types.Add(corner + "Cell");
            }
        }

        return types;
    }

    private static readonly string[] Looks = ["accent1", "accent2", "accent6", "striped", "light", "grid", "plain"];

    /// <summary>
    /// The table style of Docs whose look a table has: the borders of the table fit, and so does the shading of most cells (Docs writes its
    /// table styles into Word files as direct formatting). The shading that only repeats the look is dropped; the rest stays on its cells.
    /// </summary>
    private static string? Look(DocTable table, Dictionary<string, StyleSheet.Border?> borders)
    {
        foreach (var key in Looks)
        {
            var look = DocsStyles.Look(key);
            if (!BordersFit(look, borders))
            {
                continue;
            }

            var (agree, disagree, filled) = (0, 0, 0);
            foreach (var (cell, background) in LookShading(table, look))
            {
                if (cell.Css.Get("background-color") == background)
                {
                    agree++;
                    filled += background is null ? 0 : 1;
                }
                else
                {
                    disagree++;
                }
            }

            if (disagree * 4 > agree || (filled == 0 && (look.HeaderFill is not null || look.BandFill is not null)))
            {
                continue;
            }

            foreach (var (cell, background) in LookShading(table, look))
            {
                if (background is not null && cell.Css.Get("background-color") == background)
                {
                    cell.Css.Remove("background-color");
                }
            }

            return key;
        }

        return null;
    }

    /// <summary>Every cell with the shading a look gives it (header row, banded rows), or null.</summary>
    private static IEnumerable<(DocCell Cell, string? Background)> LookShading(DocTable table, TableLook look)
    {
        var bodyIndex = 0;
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            var background = DocsStyles.CellCss(look, row.Header, r == 0, bodyIndex).Get("background-color");
            foreach (var cell in row.Cells)
            {
                yield return (cell, background);
            }

            if (!row.Header)
            {
                bodyIndex++;
            }
        }
    }

    private static bool BordersFit(TableLook look, Dictionary<string, StyleSheet.Border?> borders)
    {
        string[] sides = ["top", "left", "bottom", "right", "insideH", "insideV"];
        string? Colour(string side) => borders.GetValueOrDefault(side)?.Css?.Split(' ') is [_, "solid", var colour] ? colour : null;
        bool None(string side) => borders.GetValueOrDefault(side)?.Css is null;
        return look switch
        {
            { Border: null } => sides.All(None),
            { AllBorders: true } => sides.All(s => Colour(s) == look.Border),
            _ => Colour("top") == look.Border && Colour("bottom") == look.Border && Colour("insideH") == look.Border && None("left") && None("right") && None("insideV"),
        };
    }

    /// <summary>The text of a header row is bold in every table style (white on the accent colours): that is the look, not formatting.</summary>
    private static void HeaderText(DocTable table)
    {
        var look = DocsStyles.Look(table.TableStyle);
        foreach (var cell in table.Rows.Where(r => r.Header).SelectMany(r => r.Cells))
        {
            foreach (var inline in DocsHtml.Paragraphs(cell.Blocks).SelectMany(p => p.Inlines))
            {
                inline.Format = inline.Format with { Bold = false, Color = inline.Format.Color == look.HeaderText ? null : inline.Format.Color };
            }
        }
    }

    private static void CollectRows(IEnumerable<OpenXmlElement> elements, List<TableRow> rows, int depth)
    {
        foreach (var element in elements)
        {
            switch (element)
            {
                case TableRow row:
                    rows.Add(row);
                    break;
                case SdtRow sdt when depth < MaxDepth:
                    CollectRows(sdt.SdtContentRow?.ChildElements ?? Enumerable.Empty<OpenXmlElement>(), rows, depth + 1);
                    break;
                case CustomXmlRow or InsertedRun when depth < MaxDepth:
                    CollectRows(element.ChildElements, rows, depth + 1);
                    break;
            }
        }
    }

    private static void CollectCells(IEnumerable<OpenXmlElement> elements, List<TableCell> cells, int depth)
    {
        foreach (var element in elements)
        {
            switch (element)
            {
                case TableCell cell:
                    cells.Add(cell);
                    break;
                case SdtCell sdt when depth < MaxDepth:
                    CollectCells(sdt.SdtContentCell?.ChildElements ?? Enumerable.Empty<OpenXmlElement>(), cells, depth + 1);
                    break;
                case CustomXmlCell when depth < MaxDepth:
                    CollectCells(element.ChildElements, cells, depth + 1);
                    break;
            }
        }
    }

    private static double Twips(string? value) => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? Math.Max(0, n) : 0;

    private static int? Int(string? value) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;

    // ------------------------------------------------------------------ lists

    /// <summary>
    /// The numbers of Word against the numbers Docs would show (a list restarts after any other block): where they differ, the item gets
    /// <c>data-start</c> (a list continued after a paragraph, a list starting at 5).
    /// </summary>
    private void NumberLists(List<DocBlock> blocks)
    {
        var numbering = new ListNumbering();
        foreach (var block in blocks)
        {
            if (block is not DocParagraph { List: not null } p)
            {
                numbering.Reset();
                continue;
            }

            var (number, _) = numbering.Next(p);
            if (_wordNumbers.TryGetValue(p, out var word) && word != number && word is >= 1 and <= 99999)
            {
                p.Start = word;
                numbering.Reset();
                numbering.Next(p);
                Reapply(numbering, blocks, p);
            }
        }
    }

    private static void Reapply(ListNumbering numbering, List<DocBlock> blocks, DocParagraph upTo)
    {
        // The counters of the levels above stay where Word had them: replay the run up to this item.
        numbering.Reset();
        var start = blocks.IndexOf(upTo);
        var first = start;
        while (first > 0 && blocks[first - 1] is DocParagraph { List: not null })
        {
            first--;
        }

        for (var i = first; i <= start; i++)
        {
            numbering.Next((DocParagraph)blocks[i]);
        }
    }

    // ------------------------------------------------------------------ section

    private void ReadSection(SectionProperties? section, List<SectionProperties> all)
    {
        var page = _document.Page;
        if (section?.GetFirstChild<PageSize>() is { } size && Twips(size.Width?.InnerText) > 0 && Twips(size.Height?.InnerText) > 0)
        {
            var w = Twips(size.Width?.InnerText) / 56.6929;
            var h = Twips(size.Height?.InnerText) / 56.6929;
            page.Landscape = size.Orient?.InnerText == "landscape" || w > h;
            var (portraitW, portraitH) = (Math.Min(w, h), Math.Max(w, h));
            page.Size = PageSetup.SizeOf(portraitW, portraitH);
            page.Width = Math.Round(portraitW, 1);
            page.Height = Math.Round(portraitH, 1);
        }

        if (section?.GetFirstChild<PageMargin>() is { } margin)
        {
            page.Top = Mm(margin.Top?.InnerText, page.Top);
            page.Right = Mm(margin.Right?.InnerText, page.Right);
            page.Bottom = Mm(margin.Bottom?.InnerText, page.Bottom);
            page.Left = Mm(margin.Left?.InnerText, page.Left);
            page.Header = Mm(margin.Header?.InnerText, page.Header);
            page.Footer = Mm(margin.Footer?.InnerText, page.Footer);
        }

        page.Columns = Math.Clamp(Int(section?.GetFirstChild<Columns>()?.ColumnCount?.InnerText) ?? 1, 1, 3);
        page.DifferentFirst = section?.GetFirstChild<TitlePage>() is { } title && StyleSheet.OnOff(title);
        page.Normalize();

        // A section without its own header takes the one of the section before.
        var chain = all.TakeWhile(s => !ReferenceEquals(s, section)).Append(section).Where(s => s is not null).Reverse().Cast<SectionProperties>().ToList();
        _document.Header = HeaderFooter(chain, header: true, "default");
        _document.Footer = HeaderFooter(chain, header: false, "default");
        if (page.DifferentFirst)
        {
            _document.FirstHeader = HeaderFooter(chain, header: true, "first");
            _document.FirstFooter = HeaderFooter(chain, header: false, "first");
        }
    }

    private static double Mm(string? twips, double fallback)
        => double.TryParse(twips, NumberStyles.Float, CultureInfo.InvariantCulture, out var t) && double.IsFinite(t) ? Math.Round(Math.Abs(t) / 56.6929, 1) : fallback;

    private List<DocBlock> HeaderFooter(List<SectionProperties> chain, bool header, string type)
    {
        foreach (var section in chain)
        {
            HeaderFooterReferenceType? reference = header
                ? section.Elements<HeaderReference>().FirstOrDefault(r => (r.Type?.InnerText ?? "default") == type)
                : section.Elements<FooterReference>().FirstOrDefault(r => (r.Type?.InnerText ?? "default") == type);
            if (reference?.Id?.Value is not { } id)
            {
                continue;
            }

            var part = _main.Parts.FirstOrDefault(p => p.RelationshipId == id).OpenXmlPart;
            OpenXmlElement? root = part switch
            {
                HeaderPart h => h.Header,
                FooterPart f => f.Footer,
                _ => null,
            };
            if (part is null || root is null)
            {
                return [];
            }

            var blocks = new List<DocBlock>();
            ReadBlocks(root.ChildElements, blocks, new Story(part) { HeaderFooter = true }, 1);
            return blocks.All(b => b is DocParagraph { Inlines.Count: 0 }) ? [] : blocks;
        }

        return [];
    }

    [System.Text.RegularExpressions.GeneratedRegex(@"\\l\s+""([^""]+)""")]
    private static partial System.Text.RegularExpressions.Regex AnchorInFieldRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"HYPERLINK\s+(?:(\\l)\s+)?""([^""]+)""", System.Text.RegularExpressions.RegexOptions.IgnoreCase)]
    private static partial System.Text.RegularExpressions.Regex HyperlinkFieldRegex();
}
