using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// Writes a Word document (<c>.docx</c>) that Word opens without repair: the styles of Docs as Word styles, one numbering per list, direct
/// formatting from the inline CSS, tables with the looks of the table styles, pictures, links and bookmarks, headers and footers with page fields,
/// footnotes, comments with replies, the page setup.
/// </summary>
internal sealed class DocxWriter
{
    private const long EmuPerPixel = 9525;

    private static readonly Dictionary<string, HighlightColorValues> Highlights = new(StringComparer.Ordinal)
    {
        ["#ffff00"] = HighlightColorValues.Yellow, ["#00ff00"] = HighlightColorValues.Green, ["#00ffff"] = HighlightColorValues.Cyan,
        ["#ff00ff"] = HighlightColorValues.Magenta, ["#0000ff"] = HighlightColorValues.Blue, ["#ff0000"] = HighlightColorValues.Red,
        ["#000080"] = HighlightColorValues.DarkBlue, ["#008080"] = HighlightColorValues.DarkCyan, ["#008000"] = HighlightColorValues.DarkGreen,
        ["#800080"] = HighlightColorValues.DarkMagenta, ["#800000"] = HighlightColorValues.DarkRed, ["#808000"] = HighlightColorValues.DarkYellow,
        ["#808080"] = HighlightColorValues.DarkGray, ["#c0c0c0"] = HighlightColorValues.LightGray, ["#000000"] = HighlightColorValues.Black,
    };

    private readonly WordprocessingDocument _package;
    private readonly MainDocumentPart _main;
    private readonly DocsDocument _document;
    private readonly string _name;
    private readonly Dictionary<string, string> _bookmarks = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _noteNumbers;
    private readonly HashSet<string> _notesWritten = new(StringComparer.Ordinal);
    private readonly Dictionary<string, AbstractNum> _abstractNums = new(StringComparer.Ordinal);
    private readonly List<NumberingInstance> _nums = [];
    private readonly Dictionary<string, (DocComment Comment, int WordId, List<(DocComment Reply, int WordId)> Replies)> _comments = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _commentRemaining = new(StringComparer.Ordinal);
    private readonly HashSet<string> _commentOpen = new(StringComparer.Ordinal);
    private int _bookmarkId;
    private uint _drawingId;

    private DocxWriter(WordprocessingDocument package, DocsDocument document, string name)
    {
        _package = package;
        _main = package.AddMainDocumentPart();
        _document = document;
        _name = name;
        _noteNumbers = DocsConvert.NoteNumbers(document);
    }

    /// <summary>The page width between the margins in twips (the base of table and picture sizes).</summary>
    private long TextWidthTwips => Math.Max(1440, Twips(_document.Page.TextWidth));

    public static byte[] Write(DocsDocument document, string name)
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            new DocxWriter(package, document, name).Build();
        }

        return stream.ToArray();
    }

    private static long Twips(double mm) => (long)Math.Round(mm * 56.6929);

    private static int TwipsOfPt(double pt) => (int)Math.Round(pt * 20);

    private void Build()
    {
        _package.PackageProperties.Title = _name;
        _package.PackageProperties.Created = DateTime.UtcNow;
        _package.PackageProperties.Modified = DateTime.UtcNow;
        _package.PackageProperties.Language = WordLanguage(_document.Language);

        PrepareBookmarks();
        PrepareComments();

        var body = new Body();
        _main.Document = new Document(body);
        _main.Document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        _main.Document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        _main.Document.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" };

        var story = new Story(_main) { Comments = true };
        Blocks(_document.Body, body, story);
        if (body.LastChild is not Paragraph)
        {
            body.Append(new Paragraph());
        }

        body.Append(BuildSection());

        _main.AddNewPart<StyleDefinitionsPart>().Styles = DocxStyles.Create(WordLanguage(_document.Language));
        _main.AddNewPart<DocumentSettingsPart>().Settings = BuildSettings();
        if (_abstractNums.Count > 0)
        {
            var numbering = new Numbering();
            foreach (var abstractNum in _abstractNums.Values)
            {
                numbering.Append(abstractNum);
            }

            foreach (var num in _nums)
            {
                numbering.Append(num);
            }

            _main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        }

        WriteComments();
    }

    private static string WordLanguage(string language) => language switch
    {
        "de" => "de-DE",
        "en" => "en-US",
        "fr" => "fr-FR",
        "es" => "es-ES",
        "it" => "it-IT",
        "nl" => "nl-NL",
        _ => language,
    };

    private Settings BuildSettings()
    {
        var settings = new Settings(
            new Zoom { Percent = "100" },
            new DefaultTabStop { Val = 720 },
            new CharacterSpacingControl { Val = CharacterSpacingValues.DoNotCompress });
        if (_noteNumbers.Count > 0)
        {
            settings.Append(new FootnoteDocumentWideProperties(new FootnoteSpecialReference { Id = -1 }, new FootnoteSpecialReference { Id = 0 }));
        }

        settings.Append(new Compatibility(new CompatibilitySetting
        {
            Name = CompatSettingNameValues.CompatibilityMode,
            Uri = "http://schemas.microsoft.com/office/word",
            Val = "15",
        }));
        settings.Append(new ThemeFontLanguages { Val = WordLanguage(_document.Language) });
        return settings;
    }

    // ------------------------------------------------------------------ section, headers, footers, notes

    private SectionProperties BuildSection()
    {
        var page = _document.Page;
        var section = new SectionProperties();
        AddHeaderFooter(section, _document.Header, HeaderFooterValues.Default, header: true, force: false);
        AddHeaderFooter(section, _document.Footer, HeaderFooterValues.Default, header: false, force: false);
        if (page.DifferentFirst)
        {
            AddHeaderFooter(section, _document.FirstHeader, HeaderFooterValues.First, header: true, force: true);
            AddHeaderFooter(section, _document.FirstFooter, HeaderFooterValues.First, header: false, force: true);
        }

        if (_noteNumbers.Count > 0)
        {
            WriteFootnotes();
        }

        var size = new PageSize { Width = (UInt32Value)(uint)Twips(page.PaperWidth), Height = (UInt32Value)(uint)Twips(page.PaperHeight) };
        if (page.Landscape)
        {
            size.Orient = PageOrientationValues.Landscape;
        }

        section.Append(size);
        section.Append(new PageMargin
        {
            Top = (int)Twips(page.Top),
            Right = (uint)Twips(page.Right),
            Bottom = (int)Twips(page.Bottom),
            Left = (uint)Twips(page.Left),
            Header = (uint)Twips(page.Header),
            Footer = (uint)Twips(page.Footer),
            Gutter = 0U,
        });
        section.Append(new Columns { Space = "720", ColumnCount = (Int16Value)(short)page.Columns });
        if (page.DifferentFirst)
        {
            section.Append(new TitlePage());
        }

        section.Append(new DocGrid { LinePitch = 360 });
        return section;
    }

    private void AddHeaderFooter(SectionProperties section, List<DocBlock> blocks, HeaderFooterValues type, bool header, bool force)
    {
        if (blocks.Count == 0 && !force)
        {
            return;
        }

        if (header)
        {
            var part = _main.AddNewPart<HeaderPart>();
            var root = new Header();
            Blocks(blocks, root, new Story(part) { HeaderFooter = true });
            if (root.LastChild is not Paragraph)
            {
                root.Append(new Paragraph());
            }

            part.Header = root;
            section.Append(new HeaderReference { Type = type, Id = _main.GetIdOfPart(part) });
        }
        else
        {
            var part = _main.AddNewPart<FooterPart>();
            var root = new Footer();
            Blocks(blocks, root, new Story(part) { HeaderFooter = true });
            if (root.LastChild is not Paragraph)
            {
                root.Append(new Paragraph());
            }

            part.Footer = root;
            section.Append(new FooterReference { Type = type, Id = _main.GetIdOfPart(part) });
        }
    }

    private void WriteFootnotes()
    {
        var part = _main.AddNewPart<FootnotesPart>();
        var footnotes = new Footnotes(
            new Footnote(SeparatorParagraph(new SeparatorMark())) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Footnote(SeparatorParagraph(new ContinuationSeparatorMark())) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
        foreach (var (note, number) in DocsConvert.OrderedNotes(_document))
        {
            var footnote = new Footnote { Id = number };
            Blocks(note.Blocks, footnote, new Story(part) { Note = true });
            if (footnote.GetFirstChild<Paragraph>() is not { } first)
            {
                first = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }));
                footnote.Append(first);
            }

            var mark = new Run(new RunProperties(new RunStyle { Val = "FootnoteReference" }), new FootnoteReferenceMark());
            var space = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
            if (first.ParagraphProperties is { } pPr)
            {
                pPr.InsertAfterSelf(mark);
            }
            else
            {
                first.PrependChild(mark);
            }

            mark.InsertAfterSelf(space);
            if (footnote.LastChild is not Paragraph)
            {
                footnote.Append(new Paragraph());
            }

            footnotes.Append(footnote);
        }

        part.Footnotes = footnotes;
    }

    private static Paragraph SeparatorParagraph(OpenXmlElement mark)
        => new(new ParagraphProperties(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }), new Run(mark));

    // ------------------------------------------------------------------ bookmarks, comments

    /// <summary>Every block id becomes a bookmark (Word's names: letters, digits and underscores, at most 40).</summary>
    private void PrepareBookmarks()
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in DocsHtml.Paragraphs(_document.Body))
        {
            if (p.Id is null)
            {
                continue;
            }

            var name = new string(p.Id.Select(c => char.IsAsciiLetterOrDigit(c) ? c : '_').Take(40).ToArray());
            var unique = name;
            for (var n = 2; !used.Add(unique); n++)
            {
                var suffix = "_" + n.ToString(CultureInfo.InvariantCulture);
                unique = name[..Math.Min(name.Length, 40 - suffix.Length)] + suffix;
            }

            _bookmarks[p.Id] = unique;
        }
    }

    /// <summary>The comments with an anchor in the body, their Word ids, and how many runs each one covers (its range ends after the last).</summary>
    private void PrepareComments()
    {
        foreach (var p in DocsHtml.Paragraphs(_document.Body))
        {
            foreach (var inline in p.VisibleInlines())
            {
                if (inline.Format.Comment is { } ids)
                {
                    foreach (var id in ids.Split(' '))
                    {
                        _commentRemaining[id] = _commentRemaining.GetValueOrDefault(id) + 1;
                    }
                }
            }
        }

        var next = 0;
        foreach (var comment in _document.Comments)
        {
            if (_commentRemaining.ContainsKey(comment.Id))
            {
                _comments[comment.Id] = (comment, next++, comment.Replies.Select(r => (r, next++)).ToList());
            }
        }
    }

    private void WriteComments()
    {
        if (_comments.Count == 0)
        {
            return;
        }

        var part = _main.AddNewPart<WordprocessingCommentsPart>();
        var comments = new Comments();
        comments.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        comments.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        comments.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" };
        var extended = new W15.CommentsEx();
        var paraId = 0x10000000;
        foreach (var (comment, wordId, replies) in _comments.Values.OrderBy(c => c.WordId))
        {
            var parentPara = (++paraId).ToString("X8", CultureInfo.InvariantCulture);
            comments.Append(CommentElement(comment, wordId, parentPara));
            extended.Append(new W15.CommentEx { ParaId = parentPara, Done = comment.Done });
            foreach (var (reply, replyId) in replies)
            {
                var replyPara = (++paraId).ToString("X8", CultureInfo.InvariantCulture);
                comments.Append(CommentElement(reply, replyId, replyPara));
                extended.Append(new W15.CommentEx { ParaId = replyPara, ParaIdParent = parentPara, Done = comment.Done });
            }
        }

        part.Comments = comments;
        _main.AddNewPart<WordprocessingCommentsExPart>().CommentsEx = extended;
    }

    private static Comment CommentElement(DocComment comment, int id, string paraId)
    {
        var author = comment.Author.Length > 0 ? comment.Author : "?";
        var element = new Comment
        {
            Id = id.ToString(CultureInfo.InvariantCulture),
            Author = author,
            Initials = string.Concat(author.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(w => w[0]).Take(3)),
        };
        if (DateTime.TryParse(comment.Date, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var date))
        {
            element.Date = date;
        }

        var lines = comment.Text.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var p = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "CommentText" }));
            if (i == 0)
            {
                p.Append(new Run(new RunProperties(new RunStyle { Val = "CommentReference" }), new AnnotationReferenceMark()));
            }

            p.Append(new Run(new Text(XmlSafe(lines[i])) { Space = SpaceProcessingModeValues.Preserve }));
            if (i == lines.Length - 1)
            {
                p.ParagraphId = paraId;
            }

            element.Append(p);
        }

        return element;
    }

    /// <summary>Opens the ranges of comments that start at this inline.</summary>
    private void CommentStarts(DocInline inline, OpenXmlElement target)
    {
        if (inline.Format.Comment is not { } ids)
        {
            return;
        }

        foreach (var id in ids.Split(' '))
        {
            if (_comments.TryGetValue(id, out var c) && _commentOpen.Add(id))
            {
                target.Append(new CommentRangeStart { Id = c.WordId.ToString(CultureInfo.InvariantCulture) });
                foreach (var (_, replyId) in c.Replies)
                {
                    target.Append(new CommentRangeStart { Id = replyId.ToString(CultureInfo.InvariantCulture) });
                }
            }
        }
    }

    /// <summary>Closes the ranges of comments whose last inline this was (end mark and the reference run).</summary>
    private void CommentEnds(DocInline inline, OpenXmlElement target)
    {
        if (inline.Format.Comment is not { } ids)
        {
            return;
        }

        foreach (var id in ids.Split(' '))
        {
            if (!_comments.TryGetValue(id, out var c) || !_commentOpen.Contains(id))
            {
                continue;
            }

            _commentRemaining[id]--;
            if (_commentRemaining[id] > 0)
            {
                continue;
            }

            _commentOpen.Remove(id);
            foreach (var wordId in new[] { c.WordId }.Concat(c.Replies.Select(r => r.WordId)))
            {
                var text = wordId.ToString(CultureInfo.InvariantCulture);
                target.Append(new CommentRangeEnd { Id = text });
                target.Append(new Run(new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = text }));
            }
        }
    }

    // ------------------------------------------------------------------ blocks

    /// <summary>Where blocks are written: the part (for pictures and links) and what kind of story it is.</summary>
    private sealed class Story(OpenXmlPart part)
    {
        public OpenXmlPart Part { get; } = part;

        public bool HeaderFooter { get; init; }

        public bool Note { get; init; }

        public bool Comments { get; init; }

        /// <summary>The cells of a header row: bold, in the colour of the table style.</summary>
        public TableLook? HeaderRow { get; init; }

        public Story With(TableLook? headerRow) => new(Part) { HeaderFooter = HeaderFooter, Note = Note, Comments = Comments, HeaderRow = headerRow };
    }

    private void Blocks(IReadOnlyList<DocBlock> blocks, OpenXmlCompositeElement parent, Story story)
    {
        var i = 0;
        while (i < blocks.Count)
        {
            if (blocks[i] is DocParagraph { List: not null })
            {
                var end = i;
                while (end < blocks.Count && blocks[end] is DocParagraph { List: not null })
                {
                    end++;
                }

                ListRun(blocks.Skip(i).Take(end - i).Cast<DocParagraph>().ToList(), parent, story);
                i = end;
                continue;
            }

            switch (blocks[i])
            {
                case DocParagraph p:
                    parent.Append(WriteParagraph(p, story, null));
                    break;
                case DocTable table:
                    parent.Append(WriteTable(table, story));
                    break;
                case DocRule:
                    parent.Append(new Paragraph(new ParagraphProperties(
                        new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6U, Space = 1U, Color = "auto" }),
                        new SpacingBetweenLines { Before = "0", After = "160" })));
                    break;
                case DocBreak br:
                    parent.Append(new Paragraph(new Run(new Break { Type = br.Kind == "column" ? BreakValues.Column : BreakValues.Page })));
                    break;
            }

            i++;
        }
    }

    /// <summary>
    /// A run of list paragraphs: one abstract numbering for the kinds of its levels, one numbering instance per segment (a list restarts at
    /// <c>data-start</c>; every instance overrides the start values, so no list continues another one by accident).
    /// </summary>
    private void ListRun(List<DocParagraph> items, OpenXmlCompositeElement parent, Story story)
    {
        var kinds = new string[9];
        foreach (var item in items)
        {
            kinds[Math.Clamp(item.Level, 0, 8)] ??= item.List!;
        }

        var first = items[0].List!;
        for (var l = 0; l < 9; l++)
        {
            kinds[l] ??= l > 0 ? kinds[l - 1] : first;
        }

        var abstractId = AbstractNumFor(kinds);
        var counters = new int[9];
        var numId = 0;
        for (var i = 0; i < items.Count; i++)
        {
            var p = items[i];
            var level = Math.Clamp(p.Level, 0, 8);
            if (i == 0 || p.Start is not null)
            {
                var starts = new int[9];
                for (var l = 0; l < 9; l++)
                {
                    starts[l] = l < level ? counters[l] + 1 : l == level ? p.Start ?? 1 : 1;
                }

                numId = Num(abstractId, starts);
            }

            counters[level] = p.Start ?? counters[level] + 1;
            for (var l = level + 1; l < 9; l++)
            {
                counters[l] = 0;
            }

            parent.Append(WriteParagraph(p, story, (numId, level)));
        }
    }

    private int AbstractNumFor(string[] kinds)
    {
        var key = string.Join('|', kinds);
        if (_abstractNums.TryGetValue(key, out var existing))
        {
            return existing.AbstractNumberId!.Value;
        }

        var id = _abstractNums.Count;
        var multi = kinds.Contains("outline") || kinds.Distinct().Count() > 1;
        var abstractNum = new AbstractNum(new MultiLevelType { Val = multi ? MultiLevelValues.Multilevel : MultiLevelValues.HybridMultilevel })
        {
            AbstractNumberId = id,
        };
        for (var level = 0; level < 9; level++)
        {
            abstractNum.Append(LevelElement(kinds[level], level));
        }

        _abstractNums[key] = abstractNum;
        return id;
    }

    private static Level LevelElement(string kind, int level)
    {
        var format = DocsStyles.LevelFormat(kind, level);
        var (numFmt, text, font) = format switch
        {
            "bullet" => (NumberFormatValues.Bullet, "", "Symbol"),
            "circle" => (NumberFormatValues.Bullet, "o", "Courier New"),
            "square" => (NumberFormatValues.Bullet, "", "Wingdings"),
            "dash" => (NumberFormatValues.Bullet, "–", "Calibri"),
            "check" => (NumberFormatValues.Bullet, "", "Wingdings"),
            "arrow" => (NumberFormatValues.Bullet, "", "Wingdings"),
            "lower-alpha" => (NumberFormatValues.LowerLetter, $"%{level + 1}.", (string?)null),
            "upper-alpha" => (NumberFormatValues.UpperLetter, $"%{level + 1}.", null),
            "lower-roman" => (NumberFormatValues.LowerRoman, $"%{level + 1}.", null),
            "upper-roman" => (NumberFormatValues.UpperRoman, $"%{level + 1}.", null),
            _ => (NumberFormatValues.Decimal, kind == "outline"
                ? string.Join('.', Enumerable.Range(1, level + 1).Select(n => "%" + n.ToString(CultureInfo.InvariantCulture))) + "."
                : $"%{level + 1}.", null),
        };
        var lvl = new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = numFmt },
            new LevelText { Val = text },
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(new Indentation
            {
                Left = ((level + 1) * 720).ToString(CultureInfo.InvariantCulture),
                Hanging = kind == "outline" ? (360 + (level * 144)).ToString(CultureInfo.InvariantCulture) : "360",
            }))
        {
            LevelIndex = level,
        };
        if (font is not null)
        {
            lvl.Append(new NumberingSymbolRunProperties(new RunFonts { Ascii = font, HighAnsi = font, Hint = FontTypeHintValues.Default }));
        }

        return lvl;
    }

    private int Num(int abstractId, int[] starts)
    {
        var id = _nums.Count + 1;
        var num = new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = id };
        for (var level = 0; level < 9; level++)
        {
            num.Append(new LevelOverride(new StartOverrideNumberingValue { Val = Math.Max(0, starts[level]) }) { LevelIndex = level });
        }

        _nums.Add(num);
        return id;
    }

    // ------------------------------------------------------------------ paragraphs and runs

    private Paragraph WriteParagraph(DocParagraph p, Story story, (int NumId, int Level)? numbering)
    {
        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties();
        var key = p.StyleKey;
        if (story.Note && key == "Normal")
        {
            key = "Note";
        }

        var styleId = DocxStyles.StyleId(key);
        if (styleId != "Normal")
        {
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        }

        if (numbering is { } n)
        {
            pPr.NumberingProperties = new NumberingProperties(new NumberingLevelReference { Val = n.Level }, new NumberingId { Val = n.NumId });
        }

        ParagraphFormatting(pPr, p);
        var inlines = p.VisibleInlines().ToList();
        if (!p.Css.Has("text-align") && inlines.Count > 0 && inlines.All(i => i is DocImage { Align: "center" } || i is DocText t && string.IsNullOrWhiteSpace(t.Text)))
        {
            pPr.Justification = new Justification { Val = JustificationValues.Center };
        }

        if (pPr.HasChildren)
        {
            paragraph.Append(pPr);
        }

        int? bookmark = null;
        if (p.Id is not null && _bookmarks.TryGetValue(p.Id, out var name) && !story.HeaderFooter && !story.Note)
        {
            bookmark = ++_bookmarkId;
            paragraph.Append(new BookmarkStart { Name = name, Id = bookmark.Value.ToString(CultureInfo.InvariantCulture) });
        }

        Inlines(inlines, paragraph, story);
        if (bookmark is { } b)
        {
            paragraph.Append(new BookmarkEnd { Id = b.ToString(CultureInfo.InvariantCulture) });
        }

        return paragraph;
    }

    /// <summary>The inline CSS of a block as direct paragraph formatting.</summary>
    private static void ParagraphFormatting(ParagraphProperties pPr, DocParagraph p)
    {
        var css = p.Css;
        if (css.Get("text-align") is { } align)
        {
            pPr.Justification = new Justification
            {
                Val = align switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    _ => JustificationValues.Left,
                },
            };
        }

        var spacing = new SpacingBetweenLines();
        if (css.Get("line-height") is { } line)
        {
            if (double.TryParse(line, NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple))
            {
                spacing.Line = Math.Clamp((int)Math.Round(multiple / 1.2 * 240), 60, 2400).ToString(CultureInfo.InvariantCulture);
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            else if (DocsCss.ToPt(line) is { } pt)
            {
                spacing.Line = Math.Clamp(TwipsOfPt(pt), 20, 31680).ToString(CultureInfo.InvariantCulture);
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }
        }

        if (DocsCss.ToPt(css.Get("margin-top")) is { } before)
        {
            spacing.Before = Math.Clamp(TwipsOfPt(before), 0, 31680).ToString(CultureInfo.InvariantCulture);
        }

        if (DocsCss.ToPt(css.Get("margin-bottom")) is { } after)
        {
            spacing.After = Math.Clamp(TwipsOfPt(after), 0, 31680).ToString(CultureInfo.InvariantCulture);
        }

        if (spacing.HasAttributes)
        {
            pPr.SpacingBetweenLines = spacing;
        }

        var indentation = new Indentation();
        if (DocsCss.ToPt(css.Get("margin-left")) is { } left)
        {
            indentation.Left = Math.Clamp(TwipsOfPt(left), -31680, 31680).ToString(CultureInfo.InvariantCulture);
        }

        if (DocsCss.ToPt(css.Get("margin-right")) is { } right)
        {
            indentation.Right = Math.Clamp(TwipsOfPt(right), -31680, 31680).ToString(CultureInfo.InvariantCulture);
        }

        if (DocsCss.ToPt(css.Get("text-indent")) is { } indent)
        {
            var twips = Math.Clamp(TwipsOfPt(indent), -31680, 31680);
            if (twips >= 0)
            {
                indentation.FirstLine = twips.ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                indentation.Hanging = (-twips).ToString(CultureInfo.InvariantCulture);
            }
        }

        if (indentation.HasAttributes)
        {
            pPr.Indentation = indentation;
        }

        if (DocxStyles.Hex(css.Get("background-color")) is { } fill)
        {
            pPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
        }

        var top = Border<TopBorder>(css.Get("border-top"), css.Get("padding-top"));
        var leftBorder = Border<LeftBorder>(css.Get("border-left"), css.Get("padding-left"));
        var bottom = Border<BottomBorder>(css.Get("border-bottom"), css.Get("padding-bottom"));
        var rightBorder = Border<RightBorder>(css.Get("border-right"), css.Get("padding-right"));
        if (top is not null || leftBorder is not null || bottom is not null || rightBorder is not null)
        {
            var borders = new ParagraphBorders();
            borders.TopBorder = top;
            borders.LeftBorder = leftBorder;
            borders.BottomBorder = bottom;
            borders.RightBorder = rightBorder;
            pPr.ParagraphBorders = borders;
        }
    }

    /// <summary>A CSS border (<c>0.5pt solid #000000</c>) as a Word border; the padding becomes the distance (pt, at most 31).</summary>
    public static T? Border<T>(string? css, string? padding = null)
        where T : BorderType, new()
    {
        if (css is null || DocsHtml.Border(css) is not { } canonical)
        {
            return null;
        }

        var parts = canonical.Split(' ');
        var width = DocsCss.ToPt(parts[0]) ?? 0.5;
        return new T
        {
            Val = parts[1] switch
            {
                "dashed" => BorderValues.Dashed,
                "dotted" => BorderValues.Dotted,
                "double" => BorderValues.Double,
                "groove" => BorderValues.ThreeDEngrave,
                "ridge" => BorderValues.ThreeDEmboss,
                "inset" => BorderValues.Inset,
                "outset" => BorderValues.Outset,
                _ => BorderValues.Single,
            },
            Size = (uint)Math.Clamp(Math.Round(width * 8), 2, 96),
            Space = (uint)Math.Clamp(Math.Round(DocsCss.ToPt(padding) ?? 0), 0, 31),
            Color = parts[2] == "#000000" ? "auto" : parts[2][1..].ToUpperInvariant(),
        };
    }

    private void Inlines(List<DocInline> inlines, Paragraph paragraph, Story story)
    {
        OpenXmlCompositeElement target = paragraph;
        string? href = null;
        foreach (var inline in inlines)
        {
            var format = inline.Format;
            if (story.HeaderRow is { } look)
            {
                format = format with { Bold = true, Color = format.Color ?? look.HeaderText };
            }

            if (format.Href != href)
            {
                target = paragraph;
                href = format.Href;
                if (href is not null && LinkElement(href, story) is { } link)
                {
                    paragraph.Append(link);
                    target = link;
                }
            }

            if (story.Comments)
            {
                CommentStarts(inline, target);
            }

            var rPr = RunPropertiesOf(format, target is Hyperlink);
            switch (inline)
            {
                case DocText t:
                    TextRuns(target, rPr, t.Text);
                    break;
                case DocLineBreak:
                    target.Append(NewRun(rPr, new Break()));
                    break;
                case DocImage image when PictureElement(image, story) is { } drawing:
                    target.Append(NewRun(rPr, drawing));
                    break;
                case DocNoteRef note when !story.HeaderFooter && !story.Note && _noteNumbers.TryGetValue(note.Id, out var number):
                    if (_notesWritten.Add(note.Id))
                    {
                        target.Append(new Run(new RunProperties(new RunStyle { Val = "FootnoteReference" }), new FootnoteReference { Id = number }));
                    }
                    else
                    {
                        var sup = RunPropertiesOf(format with { Script = 1 }, false);
                        target.Append(NewRun(sup, new Text(number.ToString(CultureInfo.InvariantCulture))));
                    }

                    break;
                case DocField field:
                    Field(target, rPr, field.Kind);
                    break;
            }

            if (story.Comments)
            {
                CommentEnds(inline, target);
            }
        }
    }

    private static Run NewRun(RunProperties? rPr, params OpenXmlElement[] children)
    {
        var run = new Run();
        if (rPr is not null)
        {
            run.Append(rPr.CloneNode(true));
        }

        run.Append(children);
        return run;
    }

    private static void TextRuns(OpenXmlCompositeElement target, RunProperties? rPr, string text)
    {
        var run = NewRun(rPr);
        var sb = new StringBuilder();

        void FlushText()
        {
            if (sb.Length > 0)
            {
                run.Append(new Text(XmlSafe(sb.ToString())) { Space = SpaceProcessingModeValues.Preserve });
                sb.Clear();
            }
        }

        foreach (var c in text)
        {
            switch (c)
            {
                case '\t':
                    FlushText();
                    run.Append(new TabChar());
                    break;
                case '\n':
                    FlushText();
                    run.Append(new Break());
                    break;
                default:
                    sb.Append(c);
                    break;
            }
        }

        FlushText();
        if (run.ChildElements.Any(e => e is not RunProperties))
        {
            target.Append(run);
        }
    }

    /// <summary>Text that XML can hold (Word refuses a file with a forbidden character).</summary>
    public static string XmlSafe(string text)
    {
        foreach (var c in text)
        {
            if (!XmlConvert.IsXmlChar(c) && !char.IsSurrogate(c))
            {
                return new string(text.Where(ch => XmlConvert.IsXmlChar(ch) || char.IsSurrogate(ch)).ToArray());
            }
        }

        return text;
    }

    private static RunProperties? RunPropertiesOf(RunFormat f, bool link)
    {
        var rPr = new RunProperties();
        if (link)
        {
            rPr.RunStyle = new RunStyle { Val = "Hyperlink" };
        }

        if (f.FontFamily is not null && DocsCss.FirstFamily(f.FontFamily) is { Length: > 0 and <= 31 } font)
        {
            rPr.RunFonts = new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font };
        }

        if (f.Bold)
        {
            rPr.Bold = new Bold();
            rPr.BoldComplexScript = new BoldComplexScript();
        }

        if (f.Italic)
        {
            rPr.Italic = new Italic();
            rPr.ItalicComplexScript = new ItalicComplexScript();
        }

        if (f.Strike)
        {
            rPr.Strike = new Strike();
        }

        if (DocxStyles.Hex(f.Color) is { } colour)
        {
            rPr.Color = new Color { Val = colour };
        }

        if (DocsCss.ToPt(f.LetterSpacing) is { } spacing && spacing != 0)
        {
            rPr.Spacing = new Spacing { Val = Math.Clamp(TwipsOfPt(spacing), -1584, 1584) };
        }

        if (DocsCss.ToPt(f.FontSize) is { } size && size > 0)
        {
            var half = Math.Clamp((int)Math.Round(size * 2), 2, 3276).ToString(CultureInfo.InvariantCulture);
            rPr.FontSize = new FontSize { Val = half };
            rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = half };
        }

        if (DocsCss.Colour(f.Background) is { } background)
        {
            if (Highlights.TryGetValue(background, out var highlight))
            {
                rPr.Highlight = new Highlight { Val = highlight };
            }
            else
            {
                rPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = background[1..].ToUpperInvariant() };
            }
        }

        if (f.Underline)
        {
            rPr.Underline = new Underline { Val = UnderlineValues.Single };
        }

        if (f.Script != 0)
        {
            rPr.VerticalTextAlignment = new VerticalTextAlignment { Val = f.Script > 0 ? VerticalPositionValues.Superscript : VerticalPositionValues.Subscript };
        }

        return rPr.HasChildren ? rPr : null;
    }

    private Hyperlink? LinkElement(string href, Story story)
    {
        if (href.StartsWith('#'))
        {
            var id = href[1..];
            return new Hyperlink { Anchor = _bookmarks.GetValueOrDefault(id) ?? new string(id.Select(c => char.IsAsciiLetterOrDigit(c) ? c : '_').Take(40).ToArray()), History = true };
        }

        try
        {
            var relationship = story.Part.AddHyperlinkRelationship(new Uri(href, UriKind.Absolute), true);
            return new Hyperlink { Id = relationship.Id, History = true };
        }
        catch (UriFormatException)
        {
            return null;
        }
    }

    private void Field(OpenXmlCompositeElement target, RunProperties? rPr, string kind)
    {
        var german = _document.Language.StartsWith("de", StringComparison.Ordinal);
        var (instruction, result) = kind switch
        {
            "page" => ("PAGE", "1"),
            "pages" => ("NUMPAGES", "1"),
            "date" => (german ? "DATE \\@ \"dd.MM.yyyy\"" : "DATE \\@ \"yyyy-MM-dd\"", DocsConvert.FieldText("date", _name, _document.Language)),
            "time" => ("TIME \\@ \"HH:mm\"", DocsConvert.FieldText("time", _name, _document.Language)),
            _ => ("TITLE", _name),
        };
        target.Append(NewRun(rPr, new FieldChar { FieldCharType = FieldCharValues.Begin }));
        target.Append(NewRun(rPr, new FieldCode(" " + instruction + " ") { Space = SpaceProcessingModeValues.Preserve }));
        target.Append(NewRun(rPr, new FieldChar { FieldCharType = FieldCharValues.Separate }));
        target.Append(NewRun(rPr, new Text(XmlSafe(result)) { Space = SpaceProcessingModeValues.Preserve }));
        target.Append(NewRun(rPr, new FieldChar { FieldCharType = FieldCharValues.End }));
    }

    // ------------------------------------------------------------------ pictures

    private Drawing? PictureElement(DocImage image, Story story)
    {
        var comma = image.Src.IndexOf(',', StringComparison.Ordinal);
        if (comma < 0)
        {
            return null;
        }

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(image.Src[(comma + 1)..]);
        }
        catch (System.FormatException)
        {
            return null;
        }

        var mime = image.Src[5..image.Src.IndexOf(';', StringComparison.Ordinal)];
        var type = mime switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            _ => new PartTypeInfo("image/webp", ".webp"),
        };
        ImagePart? part = story.Part switch
        {
            MainDocumentPart m => m.AddImagePart(type),
            HeaderPart h => h.AddImagePart(type),
            FooterPart f => f.AddImagePart(type),
            FootnotesPart n => n.AddImagePart(type),
            _ => null,
        };
        if (part is null)
        {
            return null;
        }

        using (var data = new MemoryStream(bytes))
        {
            part.FeedData(data);
        }

        var relationshipId = story.Part.GetIdOfPart(part);
        var (width, height) = Size(image, bytes);
        var cx = (long)Math.Round(width * EmuPerPixel);
        var cy = (long)Math.Round(height * EmuPerPixel);
        var id = ++_drawingId;
        var graphic = new A.Graphic(new A.GraphicData(
            new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Picture " + id.ToString(CultureInfo.InvariantCulture) },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
        {
            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
        });
        var docProperties = new DW.DocProperties { Id = id, Name = "Picture " + id.ToString(CultureInfo.InvariantCulture) };
        if (image.Alt.Length > 0)
        {
            docProperties.Description = XmlSafe(image.Alt);
        }

        var locks = new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true });
        if (image.Align is "left" or "right")
        {
            return new Drawing(new DW.Anchor(
                new DW.SimplePosition { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.HorizontalAlignment(image.Align)) { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
                new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
                docProperties,
                locks,
                graphic)
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 114300U,
                DistanceFromRight = 114300U,
                SimplePos = false,
                RelativeHeight = 251658240U + id,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
            });
        }

        return new Drawing(new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            docProperties,
            locks,
            graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        });
    }

    /// <summary>The size of a picture in pixels: as given, else its own size; never wider than the text.</summary>
    private (double Width, double Height) Size(DocImage image, byte[] bytes)
    {
        var natural = ImageSize.Of(bytes) ?? (300, 200);
        double width;
        double height;
        if (image.Width is { } w && image.Height is { } h)
        {
            (width, height) = (w, h);
        }
        else if (image.Width is { } onlyWidth)
        {
            (width, height) = (onlyWidth, onlyWidth * natural.Height / Math.Max(1, natural.Width));
        }
        else if (image.Height is { } onlyHeight)
        {
            (width, height) = (onlyHeight * natural.Width / Math.Max(1, natural.Height), onlyHeight);
        }
        else
        {
            (width, height) = natural;
        }

        var max = TextWidthTwips / 15.0;
        if (width > max)
        {
            height = height * max / width;
            width = max;
        }

        return (Math.Max(1, width), Math.Max(1, height));
    }

    // ------------------------------------------------------------------ tables

    private Table WriteTable(DocTable table, Story story)
    {
        var look = DocsStyles.Look(table.TableStyle);
        var columns = table.ColumnCount();
        var percent = DocsCss.Percent(table.Width) is { } tp and > 0 and <= 100 ? tp : 100;
        var tableTwips = TextWidthTwips * percent / 100;
        var given = table.Columns.Select(c => DocsCss.Percent(c)).ToList();
        var known = given.Take(columns).Where(g => g is > 0).Sum(g => g!.Value);
        var unknown = Enumerable.Range(0, columns).Count(i => i >= given.Count || given[i] is not > 0);
        var rest = unknown > 0 ? Math.Max(5 * unknown, 100 - known) / unknown : 0;
        var shares = Enumerable.Range(0, columns).Select(i => i < given.Count && given[i] is { } g && g > 0 ? g : rest).ToList();
        var total = shares.Sum();
        var widths = shares.Select(s => (long)Math.Max(100, tableTwips * s / total)).ToList();

        var tblPr = new TableProperties();
        if (look.Key == "grid")
        {
            tblPr.TableStyle = new TableStyle { Val = "TableGrid" };
        }

        tblPr.TableWidth = new TableWidth { Width = ((int)Math.Round(percent * 50)).ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Pct };
        if (look.Border is { } colour)
        {
            var color = colour == "#000000" ? "auto" : colour[1..].ToUpperInvariant();
            BorderType Line<T>()
                where T : BorderType, new() => new T { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = color };
            tblPr.TableBorders = look.AllBorders
                ? new TableBorders(Line<TopBorder>(), Line<LeftBorder>(), Line<BottomBorder>(), Line<RightBorder>(), Line<InsideHorizontalBorder>(), Line<InsideVerticalBorder>())
                : new TableBorders(Line<TopBorder>(), Line<BottomBorder>(), Line<InsideHorizontalBorder>());
        }

        tblPr.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
        tblPr.TableCellMarginDefault = new TableCellMarginDefault(
            new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
            new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa });
        tblPr.TableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook { Val = "04A0", FirstRow = table.Rows.Any(r => r.Header), LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

        var result = new Table(tblPr, new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString(CultureInfo.InvariantCulture) })));

        // Cells continuing a rowspan of a row above (Word: vMerge continue).
        var continuations = new Dictionary<int, List<(int Column, DocCell Cell)>>();
        var layout = table.Layout().ToList();
        var bodyIndex = 0;
        for (var r = 0; r < layout.Count; r++)
        {
            var (row, cells) = layout[r];
            var slots = new List<(int Column, DocCell Cell, bool Continue)>();
            foreach (var (_, column, cell) in cells)
            {
                if (column >= columns)
                {
                    continue;
                }

                slots.Add((column, cell, false));
                for (var dr = 1; dr < cell.RowSpan && r + dr < layout.Count; dr++)
                {
                    if (!continuations.TryGetValue(r + dr, out var list))
                    {
                        continuations[r + dr] = list = [];
                    }

                    list.Add((column, cell));
                }
            }

            if (continuations.TryGetValue(r, out var continued))
            {
                slots.AddRange(continued.Select(c => (c.Column, c.Cell, true)));
            }

            slots.Sort((a, b) => a.Column.CompareTo(b.Column));
            var tr = new TableRow();
            if (row.Header)
            {
                tr.Append(new TableRowProperties(new TableHeader()));
            }

            var next = 0;
            foreach (var (column, cell, isContinuation) in slots)
            {
                if (column < next)
                {
                    continue;
                }

                for (; next < column; next++)
                {
                    tr.Append(Cell(null, false, widths[next], 1, look, row.Header, r == 0, bodyIndex, story));
                }

                var span = Math.Min(cell.ColSpan, columns - column);
                var width = widths.Skip(column).Take(span).Sum();
                tr.Append(Cell(cell, isContinuation, width, span, look, row.Header, r == 0, bodyIndex, story));
                next = column + span;
            }

            for (; next < columns; next++)
            {
                tr.Append(Cell(null, false, widths[next], 1, look, row.Header, r == 0, bodyIndex, story));
            }

            result.Append(tr);
            if (!row.Header)
            {
                bodyIndex++;
            }
        }

        return result;
    }

    private TableCell Cell(DocCell? cell, bool continuation, long width, int span, TableLook look, bool headerRow, bool firstRow, int bodyIndex, Story story)
    {
        var tcPr = new TableCellProperties { TableCellWidth = new TableCellWidth { Width = width.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa } };
        if (span > 1)
        {
            tcPr.GridSpan = new GridSpan { Val = span };
        }

        if (cell is not null && (continuation || cell.RowSpan > 1))
        {
            tcPr.VerticalMerge = new VerticalMerge { Val = continuation ? MergedCellValues.Continue : MergedCellValues.Restart };
        }

        var css = DocsStyles.CellCss(look, headerRow, firstRow, bodyIndex);
        if (look.AllBorders || look.Border is null)
        {
            // The table borders draw these already.
            foreach (var side in new[] { "border-top", "border-right", "border-bottom", "border-left" })
            {
                css.Remove(side);
            }
        }
        else
        {
            css.Remove("border-top");
            css.Remove("border-bottom");
        }

        if (cell is not null)
        {
            foreach (var (k, v) in cell.Css.Items)
            {
                css.Set(k, v);
            }
        }

        var top = Border<TopBorder>(css.Get("border-top"));
        var left = Border<LeftBorder>(css.Get("border-left"));
        var bottom = Border<BottomBorder>(css.Get("border-bottom"));
        var right = Border<RightBorder>(css.Get("border-right"));
        if (top is not null || left is not null || bottom is not null || right is not null)
        {
            var borders = new TableCellBorders();
            borders.TopBorder = top;
            borders.LeftBorder = left;
            borders.BottomBorder = bottom;
            borders.RightBorder = right;
            tcPr.TableCellBorders = borders;
        }

        if (DocxStyles.Hex(css.Get("background-color")) is { } fill)
        {
            tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
        }

        if (css.Get("vertical-align") is { } valign && valign != "top")
        {
            tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment { Val = valign == "middle" ? TableVerticalAlignmentValues.Center : TableVerticalAlignmentValues.Bottom };
        }

        var tc = new TableCell(tcPr);
        if (cell is not null && !continuation)
        {
            Blocks(cell.Blocks, tc, story.With(headerRow ? look : null));
        }

        if (tc.LastChild is not Paragraph)
        {
            tc.Append(new Paragraph());
        }

        return tc;
    }
}

/// <summary>The pixel size of a PNG, JPEG, GIF or WebP picture, read from its header.</summary>
internal static class ImageSize
{
    public static (double Width, double Height)? Of(byte[] b)
    {
        try
        {
            if (b.Length > 24 && b[0] == 0x89 && b[1] == 'P' && b[2] == 'N' && b[3] == 'G')
            {
                return (BigEndian(b, 16), BigEndian(b, 20));
            }

            if (b.Length > 10 && b[0] == 'G' && b[1] == 'I' && b[2] == 'F')
            {
                return (b[6] | (b[7] << 8), b[8] | (b[9] << 8));
            }

            if (b.Length > 30 && b[0] == 'R' && b[1] == 'I' && b[2] == 'F' && b[3] == 'F' && b[8] == 'W' && b[9] == 'E' && b[10] == 'B' && b[11] == 'P')
            {
                return b[15] switch
                {
                    (byte)' ' => ((b[26] | (b[27] << 8)) & 0x3FFF, (b[28] | (b[29] << 8)) & 0x3FFF),
                    (byte)'L' => (1 + (((b[22] & 0x3F) << 8) | b[21]), 1 + (((b[24] & 0x0F) << 10) | (b[23] << 2) | ((b[22] & 0xC0) >> 6))),
                    (byte)'X' => (1 + (b[24] | (b[25] << 8) | (b[26] << 16)), 1 + (b[27] | (b[28] << 8) | (b[29] << 16))),
                    _ => null,
                };
            }

            if (b.Length > 4 && b[0] == 0xFF && b[1] == 0xD8)
            {
                var i = 2;
                while (i + 9 < b.Length)
                {
                    if (b[i] != 0xFF)
                    {
                        i++;
                        continue;
                    }

                    var marker = b[i + 1];
                    if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
                    {
                        return ((b[i + 7] << 8) | b[i + 8], (b[i + 5] << 8) | b[i + 6]);
                    }

                    if (marker is 0xD8 or 0x01 or (>= 0xD0 and <= 0xD7) or 0xFF)
                    {
                        i += marker == 0xFF ? 1 : 2;
                        continue;
                    }

                    i += 2 + ((b[i + 2] << 8) | b[i + 3]);
                }
            }
        }
        catch (IndexOutOfRangeException)
        {
            return null;
        }

        return null;
    }

    private static int BigEndian(byte[] b, int i) => (b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3];
}
