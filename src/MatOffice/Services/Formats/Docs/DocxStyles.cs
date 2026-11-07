using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MatOffice.Services.Formats.Docs;

/// <summary>The styles part of the Word files Docs writes: the paragraph styles of <see cref="DocsStyles"/> under their built-in Word names.</summary>
internal static class DocxStyles
{
    /// <summary>The Word style id of a style key (the keys are the ids, except the note text).</summary>
    public static string StyleId(string key) => key == "Note" ? "FootnoteText" : key;

    /// <summary>A CSS colour as Word's <c>RRGGBB</c>, or null.</summary>
    public static string? Hex(string? colour) => DocsCss.WordColour(colour);

    private static string Tw(double pt) => ((int)Math.Round(pt * 20)).ToString(CultureInfo.InvariantCulture);

    public static Styles Create(string language)
    {
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = DocsStyles.BodyFont, HighAnsi = DocsStyles.BodyFont, EastAsia = DocsStyles.BodyFont, ComplexScript = "Times New Roman" },
                    new FontSize { Val = "22" },
                    new FontSizeComplexScript { Val = "22" },
                    new Languages { Val = language, EastAsia = "en-US", Bidi = "ar-SA" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }))));

        foreach (var def in DocsStyles.All)
        {
            styles.Append(Paragraph(def));
        }

        styles.Append(Character("DefaultParagraphFont", "Default Paragraph Font", null, isDefault: true));
        styles.Append(Character("Hyperlink", "Hyperlink", new StyleRunProperties(new Color { Val = "0563C1" }, new Underline { Val = UnderlineValues.Single })));
        styles.Append(Character("FootnoteReference", "footnote reference", new StyleRunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript })));
        styles.Append(Character("CommentReference", "annotation reference", new StyleRunProperties(new FontSize { Val = "16" }, new FontSizeComplexScript { Val = "16" })));

        var commentText = new Style(
            new StyleName { Val = "annotation text" },
            new BasedOn { Val = "Normal" },
            new UIPriority { Val = 99 },
            new StyleParagraphProperties(new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(new FontSize { Val = "20" }, new FontSizeComplexScript { Val = "20" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "CommentText",
        };
        styles.Append(commentText);

        styles.Append(new Style(
            new StyleName { Val = "Normal Table" },
            new UIPriority { Val = 99 },
            new SemiHidden(),
            new UnhideWhenUsed(),
            new StyleTableProperties(
                new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
                new TableCellMarginDefault(
                    new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
                    new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa })))
        {
            Type = StyleValues.Table,
            StyleId = "TableNormal",
            Default = true,
        });

        styles.Append(new Style(
            new StyleName { Val = "Table Grid" },
            new BasedOn { Val = "TableNormal" },
            new UIPriority { Val = 39 },
            new StyleParagraphProperties(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new StyleTableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" },
                new LeftBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" },
                new BottomBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" },
                new RightBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U, Space = 0U, Color = "auto" })))
        {
            Type = StyleValues.Table,
            StyleId = "TableGrid",
        });

        styles.Append(new Style(new StyleName { Val = "No List" }, new UIPriority { Val = 99 }, new SemiHidden(), new UnhideWhenUsed())
        {
            Type = StyleValues.Numbering,
            StyleId = "NoList",
            Default = true,
        });
        return styles;
    }

    private static Style Character(string id, string name, StyleRunProperties? rPr, bool isDefault = false)
    {
        var style = new Style(new StyleName { Val = name }, new UIPriority { Val = isDefault ? 1 : 99 })
        {
            Type = StyleValues.Character,
            StyleId = id,
        };
        if (isDefault)
        {
            style.Default = true;
            style.Append(new SemiHidden(), new UnhideWhenUsed());
        }
        else
        {
            style.Append(new UnhideWhenUsed());
        }

        if (rPr is not null)
        {
            style.Append(rPr);
        }

        return style;
    }

    private static Style Paragraph(ParagraphStyle def)
    {
        var id = StyleId(def.Key);
        var normal = def.Key == "Normal";
        var heading = def.OutlineLevel is not null;
        var style = new Style { Type = StyleValues.Paragraph, StyleId = id };
        if (normal)
        {
            style.Default = true;
        }

        style.Append(new StyleName { Val = def.WordName });
        if (!normal)
        {
            style.Append(new BasedOn { Val = "Normal" });
            style.Append(new NextParagraphStyle { Val = heading || def.Key is "Title" or "Subtitle" or "TOCHeading" or "Caption" ? "Normal" : id });
        }

        style.Append(new UIPriority { Val = def.Key switch { "Normal" => 0, "Title" => 10, "Subtitle" => 11, "Quote" => 29, "IntenseQuote" => 30, "ListParagraph" => 34, "Caption" => 35, "TOCHeading" or "TOC1" or "TOC2" or "TOC3" => 39, "Note" => 99, _ => 9 } });
        if (def.Key is "TOC1" or "TOC2" or "TOC3" or "TOCHeading" or "Note" or "Caption")
        {
            style.Append(new UnhideWhenUsed());
        }

        if (def.Key is not ("TOC1" or "TOC2" or "TOC3" or "Note"))
        {
            style.Append(new PrimaryStyle());
        }

        var pPr = new StyleParagraphProperties();
        if (def.KeepNext)
        {
            pPr.KeepNext = new KeepNext();
            pPr.KeepLines = new KeepLines();
        }

        if (def.BorderColor is { } border)
        {
            var color = border[1..].ToUpperInvariant();
            pPr.ParagraphBorders = new ParagraphBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U, Space = (uint)def.Padding, Color = color },
                new BottomBorder { Val = BorderValues.Single, Size = 4U, Space = (uint)def.Padding, Color = color });
        }

        if (def.Background is { } background)
        {
            pPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = background[1..].ToUpperInvariant() };
        }

        pPr.SpacingBetweenLines = new SpacingBetweenLines
        {
            Before = Tw(def.Before),
            After = Tw(def.After),
            Line = def.LineTwips.ToString(CultureInfo.InvariantCulture),
            LineRule = LineSpacingRuleValues.Auto,
        };
        if (def.IndentLeft != 0 || def.IndentRight != 0)
        {
            pPr.Indentation = new Indentation { Left = Tw(def.IndentLeft), Right = Tw(def.IndentRight) };
        }

        if (def.Contextual)
        {
            pPr.ContextualSpacing = new ContextualSpacing();
        }

        if (def.Align == "center")
        {
            pPr.Justification = new Justification { Val = JustificationValues.Center };
        }

        if (def.OutlineLevel is { } outline)
        {
            pPr.OutlineLevel = new OutlineLevel { Val = outline };
        }
        else if (def.Key == "TOCHeading")
        {
            pPr.OutlineLevel = new OutlineLevel { Val = 9 };
        }

        style.Append(pPr);

        var rPr = new StyleRunProperties();
        if (def.Font != DocsStyles.BodyFont)
        {
            rPr.RunFonts = new RunFonts { Ascii = def.Font, HighAnsi = def.Font, EastAsia = def.Font, ComplexScript = def.Font };
        }

        if (def.Italic)
        {
            rPr.Italic = new Italic();
            rPr.ItalicComplexScript = new ItalicComplexScript();
        }

        if (def.Color is { } colour)
        {
            rPr.Color = new Color { Val = colour[1..].ToUpperInvariant() };
        }

        if (def.LetterSpacing != 0)
        {
            rPr.Spacing = new Spacing { Val = (int)Math.Round(def.LetterSpacing * 20) };
        }

        if (def.Key == "Title")
        {
            rPr.Kern = new Kern { Val = 28U };
        }

        var size = ((int)Math.Round(def.Size * 2)).ToString(CultureInfo.InvariantCulture);
        rPr.FontSize = new FontSize { Val = size };
        rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = size };
        style.Append(rPr);
        return style;
    }
}
