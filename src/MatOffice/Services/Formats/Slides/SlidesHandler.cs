using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Data;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// MatOffice Slides: presentations (like PowerPoint). The editor is <c>wwwroot/js/slides/</c>; this handler checks its content, makes the templates and
/// converts from and to PowerPoint (.pptx, <see cref="PptxWriter"/>, <see cref="PptxReader"/>). PDF is made by printing in the browser.
/// <para>
/// THE CONTENT (version 1) – one JSON object; coordinates in px of a slide at 96 dpi (16:9 = 1280 × 720, 4:3 = 960 × 720; 1 px = 9525 EMU),
/// font sizes and line widths in pt. Text is structured (paragraphs and runs, never HTML), so it is easy to check and to write as DrawingML.
/// <code>
/// {
///   "version": 1,
///   "size": { "w": 1280, "h": 720 },
///   "theme": {                                   the design, kept whole in every presentation (an imported file brings its own)
///     "id": "paper", "name": "Paper", "variant": 0,
///     "fonts": { "major": "Calibri Light", "minor": "Calibri" },          headings / body
///     "colors": { "dk1", "lt1", "dk2", "lt2", "accent1" … "accent6", "hlink" }   (#rrggbb each)
///     "background": Fill,                        of every slide
///     "heading": Color, "text": Color,           colour of titles / of body text
///     "groups": { "cover"|"content"|"other": { "background"?, "heading"?, "text"? } },   overrides per layout group
///     "decor": [ { "on": ["all"|group|layout id], "shape", "x","y","w","h" (fractions of the slide), "fill": Color | "src": picture, "opacity"?, "flipH"?, "flipV"?, "rot"? } ]
///   },
///   "footer": { "number": bool, "date": bool, "dateText": "" (empty = today), "text": "", "skipTitle": true },
///   "slides": [ {
///     "id": "s1", "layout": "title|titleContent|section|twoContent|comparison|titleOnly|blank|pictureCaption",
///     "hidden"?: true, "background"?: Fill, "notes"?: "plain text with \n",
///     "transition"?: { "type": "none|fade|push|wipe|split|zoom|cover", "dir"?: "left|right|up|down" (where the new slide comes from), "duration": s, "auto"?: s },
///     "elements": [ Element … ]                  back to front
///     "animations"?: [ { "el": element id, "effect": "appear|fade|fly|zoom|wipe", "trigger": "click|with|after", "dir"?: "bottom|top|left|right", "duration": s, "delay"?: s } ]
///   } ]
/// }
/// Element: { "id", "type", "x", "y", "w", "h", "rot"?: degrees, "flipH"?, "flipV"?, "ph"?: placeholder slot of the layout ("title", "subtitle", "body",
///            "left", "right", "leftHead", "rightHead", "picture", "caption"), "group"?: id shared by the elements of a group (kept next to each other), "name"? }
///   "text":  a text box or a placeholder – "fill"?, "line"?, "shadow"?, "opacity"?, "text": TextBody
///   "shape": "shape": rect|roundRect|ellipse|triangle|rtTriangle|diamond|parallelogram|trapezoid|pentagon|hexagon|octagon|plus|star5|star6|heart|
///            rightArrow|leftArrow|upArrow|downArrow|leftRightArrow|upDownArrow|chevron|homePlate|wedgeRectCallout|wedgeRoundRectCallout|wedgeEllipseCallout|
///            cloud|donut|can|snip1Rect, "adj"?: the one adjustment (corner radius …, a fraction), "fill"?, "line"?, "shadow"?, "opacity"? (of the fill), "text"?
///   "line":  from (x, y) to (x + w, y + h), mirrored by flipH / flipV; "line": Line, "arrowStart"/"arrowEnd": none|triangle|arrow|stealth|oval|diamond
///   "image": "src": data:image/(png|jpeg|gif|webp);base64,… (no SVG), "crop"?: { "l","t","r","b" fractions }, "line"?, "shadow"?, "opacity"?, "alt"?
///   "table": "cols": [px …], "rows": [ { "h": px, "cells": [ { "text": TextBody, "fill"?: Fill } ] } ],
///            "tableStyle": none|light|medium|dark, "accent": Color, "header", "banded", "firstCol", "lastRow": bool
/// TextBody: { "paras": [ { "runs": [ { "t": "text (\n = line break)", "b"?, "i"?, "u"?, "s"? (strike), "font"?, "size"? (pt), "color"?, "hl"? (highlight) } ],
///             "align"?: left|center|right|justify, "level"?: 0..8, "bullet"?: "none" | a character ("•") | a numbering ("1." "1)" "a." "a)" "A." "I." "i."),
///             "lineSpacing"? (multiple), "spaceBefore"?/"spaceAfter"? (pt) } ],
///             "anchor"?: top|middle|bottom, "autofit"?: none|shrink|grow, "inset"?: [l, t, r, b] px, "wrap"?: false }
///   What a run or a paragraph does not say comes from the placeholder slot of the layout and the design (sizes 44/28/24/20/18 pt, theme fonts and colours).
/// Fill: { "type": "none" } | { "type": "solid", "color" } | { "type": "gradient", "from", "to", "angle" } | { "type": "image", "src" }
/// Line: { "color", "width" (pt), "dash": solid|dash|dot|dashDot|longDash }
/// Color: "#rrggbb", or a colour of the design "dk1|lt1|dk2|lt2|accent1..6|hlink" with an optional ":+40" (40 % lighter) / ":-25" (25 % darker)
/// </code>
/// The layouts (where the placeholders sit) and the designs are data: <c>Resources/Templates/slides/layouts.json</c> and <c>themes.json</c> (the editor has
/// the same in <c>layouts.js</c> / <c>themes.js</c>). The check is <see cref="SlidesNormalizer"/>; the templates are <see cref="SlidesTemplates"/>.
/// </para>
/// </summary>
public sealed class SlidesHandler : IKindHandler
{
    public const string Pptx = ".pptx";

    public DocumentKind Kind => DocumentKind.Slides;

    public string Slug => "slides";

    public string Colour => "#c43e1c";

    public string Icon => "kind-slides";

    public string NativeExtension => ".matslides";

    public IReadOnlyList<FileFormat> Formats =>
        [new(Pptx, "application/vnd.openxmlformats-officedocument.presentationml.presentation", Text.T("Slides.Format.Pptx"), true, true)];

    public IReadOnlyList<KindTemplate> Templates => SlidesTemplates.All;

    public string Create(string templateId, string language) => SlidesTemplates.Build(templateId, language).ToJsonString();

    public ContentCheck Normalize(string json)
    {
        try
        {
            var root = Kinds.ParseObject(json);
            return ContentCheck.Valid(SlidesNormalizer.Normalize(root).ToJsonString());
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (JsonException)
        {
            return ContentCheck.Invalid(Text.T("Documents.Invalid"));
        }
        catch (InvalidOperationException)
        {
            return ContentCheck.Invalid(Text.T("Slides.Invalid"));
        }
    }

    public string Import(Stream input, string extension)
    {
        if (Kinds.NormalizeExtension(extension) != Pptx)
        {
            throw new FormatException(Text.T("Documents.Import.NotSupported"));
        }

        return PptxReader.Read(input).ToJsonString();
    }

    public byte[] Export(string content, string extension, string name)
    {
        if (Kinds.NormalizeExtension(extension) != Pptx)
        {
            throw new FormatException(Text.T("Documents.Export.NotSupported"));
        }

        var root = Kinds.ParseObject(content);
        return PptxWriter.Write(SlidesNormalizer.Normalize(root), name);
    }

    public string PlainText(string content, int max)
    {
        JsonObject root;
        try
        {
            root = Kinds.ParseObject(content);
        }
        catch (Exception ex) when (ex is JsonException or FormatException)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var slide in (root["slides"] as JsonArray ?? []).OfType<JsonObject>())
        {
            foreach (var e in (slide["elements"] as JsonArray ?? []).OfType<JsonObject>())
            {
                AppendBody(sb, e["text"] as JsonObject);
                foreach (var row in (e["rows"] as JsonArray ?? []).OfType<JsonObject>())
                {
                    foreach (var cell in (row["cells"] as JsonArray ?? []).OfType<JsonObject>())
                    {
                        AppendBody(sb, cell["text"] as JsonObject);
                    }
                }

                if (sb.Length >= max)
                {
                    return sb.ToString(0, max);
                }
            }

            if (slide["notes"] is JsonValue notes && notes.TryGetValue<string>(out var n) && n.Length > 0)
            {
                sb.Append(n).Append('\n');
            }

            if (sb.Length >= max)
            {
                return sb.ToString(0, max);
            }
        }

        return sb.ToString().TrimEnd();
    }

    private static void AppendBody(StringBuilder sb, JsonObject? body)
    {
        foreach (var p in (body?["paras"] as JsonArray ?? []).OfType<JsonObject>())
        {
            var line = string.Concat((p["runs"] as JsonArray ?? []).OfType<JsonObject>().Select(r => r["t"] is JsonValue v && v.TryGetValue<string>(out var s) ? s : string.Empty));
            if (line.Trim().Length > 0)
            {
                sb.Append(line).Append('\n');
            }
        }
    }
}
