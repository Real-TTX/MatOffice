using System.Collections.Frozen;
using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Ganss.Xss;

namespace MatOffice.Services.Formats.Docs;

/// <summary>How HTML is read: the editor's own HTML (kept as written) or a foreign file (whitespace of the source collapsed, style values cleaned up).</summary>
internal sealed record HtmlReadOptions
{
    public static readonly HtmlReadOptions Content = new();

    public static readonly HtmlReadOptions Import = new() { IsImport = true };

    /// <summary>Source formatting whitespace is collapsed like a browser renders it, and formatting equal to the paragraph style is dropped.</summary>
    public bool IsImport { get; init; }

    /// <summary>The HTML of Microsoft Word: a line height in percent is Word's "multiple".</summary>
    public bool IsWord { get; init; }
}

/// <summary>
/// Docs HTML (see <c>docs/EDITORS.md</c>): the sanitizer (HtmlSanitizer with a strict allowlist; styles are checked by <see cref="DocsCss"/>), the
/// canonicalizer (any HTML → the flat block list of <see cref="DocBlock"/>) and the writer of canonical Docs HTML.
/// </summary>
internal static class DocsHtml
{
    /// <summary>The deepest nesting of elements that is kept; deeper markup keeps only its text.</summary>
    private const int MaxDepth = 120;

    private static readonly FrozenSet<string> Dropped = new[]
    {
        "script", "style", "template", "noscript", "iframe", "frame", "frameset", "object", "embed", "applet", "svg", "math", "head", "title",
        "meta", "link", "base", "textarea", "select", "option", "optgroup", "button", "input", "canvas", "video", "audio", "source", "track",
        "map", "area", "noembed", "noframes", "xmp", "plaintext", "datalist", "output", "progress", "meter", "dialog", "param", "slot", "portal",
        "rp", "rt",
    }.ToFrozenSet(StringComparer.Ordinal);

    private static readonly FrozenSet<string> Containers = new[]
    {
        "div", "section", "article", "header", "footer", "main", "nav", "aside", "figure", "figcaption", "center", "address", "details",
        "summary", "dl", "dt", "dd", "fieldset", "legend", "hgroup", "form", "body", "html", "search",
    }.ToFrozenSet(StringComparer.Ordinal);

    /// <summary>The paragraph formatting in the order it is written: the canonical form does not depend on the order of the source.</summary>
    private static readonly string[] ParagraphOrder =
    [
        "text-align", "line-height", "margin-top", "margin-bottom", "margin-left", "margin-right", "text-indent", "background-color",
        "border-top", "border-right", "border-bottom", "border-left", "padding-top", "padding-right", "padding-bottom", "padding-left",
    ];

    /// <summary>The formatting of a table cell in the order it is written.</summary>
    private static readonly string[] CellOrder = ["background-color", "vertical-align", "border-top", "border-right", "border-bottom", "border-left"];

    private static readonly FrozenSet<string> ParagraphProperties = ParagraphOrder.ToFrozenSet(StringComparer.Ordinal);

    private static readonly HtmlSanitizer Sanitizer = CreateSanitizer();

    private static HtmlSanitizer CreateSanitizer()
    {
        var sanitizer = new HtmlSanitizer(new HtmlSanitizerOptions
        {
            AllowedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "p", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "table", "caption", "colgroup", "col", "thead", "tbody", "tfoot", "tr",
                "th", "td", "hr", "div", "br", "strong", "em", "u", "s", "sub", "sup", "span", "a", "img", "ul", "ol", "li",
            },
            AllowedAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "id", "href", "src", "alt", "width", "height", "colspan", "rowspan", "start", "type", "value",
                "data-style", "data-list", "data-level", "data-start", "data-break", "data-table-style", "data-align", "data-comment", "data-fn", "data-field",
            },
            AllowedCssProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            AllowedSchemes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "http", "https", "mailto", "data" },
            UriAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "href", "src" },
            AllowedAtRules = new HashSet<AngleSharp.Css.Dom.CssRuleType>(),
            AllowedCssClasses = new HashSet<string>(StringComparer.Ordinal),
        })
        {
            KeepChildNodes = true,
            AllowDataAttributes = false,
        };
        sanitizer.FilterUrl += (_, e) => e.SanitizedUrl = e.Tag.LocalName switch
        {
            "a" => SafeHref(e.OriginalUrl),
            "img" => SafeImage(e.OriginalUrl),
            _ => null,
        };
        return sanitizer;
    }

    /// <summary>Sanitizes and canonicalizes Docs HTML; an empty result is <c>""</c>.</summary>
    public static string Clean(string? html) => Write(Parse(html));

    /// <summary>Reads (sanitizes and canonicalizes) HTML into blocks.</summary>
    public static List<DocBlock> Parse(string? html, HtmlReadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return [];
        }

        var document = new HtmlParser(new HtmlParserOptions { IsScripting = false }).ParseDocument(html);
        return Parse(document, document.Body, options ?? HtmlReadOptions.Content);
    }

    /// <summary>Reads the content of an element of a parsed document (the body of an imported file) into blocks.</summary>
    public static List<DocBlock> Parse(IHtmlDocument document, IElement? root, HtmlReadOptions options)
    {
        if (root is null)
        {
            return [];
        }

        var styles = new Dictionary<IElement, CssList>(ReferenceEqualityComparer.Instance);
        Prepare(document, root, styles, 0);
        if (root is IHtmlElement html)
        {
            Sanitizer.SanitizeDom(document, html);
        }

        return new Builder(styles, options).Body(root);
    }

    // ------------------------------------------------------------------ before the sanitizer

    /// <summary>
    /// Takes the styles off (checked by <see cref="DocsCss"/>, kept aside for the builder), removes elements whose content must not survive,
    /// turns legacy markup into the canonical tags (b → strong, font → span …) and block containers into <c>div</c>.
    /// </summary>
    private static void Prepare(IDocument document, INode node, Dictionary<IElement, CssList> styles, int depth)
    {
        var child = node.FirstChild;
        while (child is not null)
        {
            var next = child.NextSibling;
            switch (child)
            {
                case IText:
                    break;
                case IElement element:
                    var name = element.LocalName;
                    if (Dropped.Contains(name))
                    {
                        element.Remove();
                        break;
                    }

                    if (depth > MaxDepth)
                    {
                        node.ReplaceChild(document.CreateTextNode(TextOf(element)), element);
                        break;
                    }

                    var target = Rename(document, element, name);
                    var css = DocsCss.Parse(target.GetAttribute("style"));
                    LegacyAttributes(target, css);
                    if (css.Count > 0)
                    {
                        styles[target] = css;
                    }

                    target.RemoveAttribute("style");
                    Prepare(document, target, styles, depth + 1);
                    break;
                default:
                    node.RemoveChild(child);
                    break;
            }

            child = next;
        }
    }

    private static IElement Rename(IDocument document, IElement element, string name)
    {
        string? to = name switch
        {
            "b" => "strong",
            "i" or "cite" or "dfn" or "var" => "em",
            "strike" or "del" => "s",
            "ins" => "u",
            "font" or "mark" or "code" or "kbd" or "samp" or "tt" => "span",
            "menu" or "dir" => "ul",
            _ when Containers.Contains(name) && name != "div" => "div",
            _ => null,
        };
        if (to is null)
        {
            return element;
        }

        var replacement = document.CreateElement(to);
        foreach (var attribute in element.Attributes)
        {
            replacement.SetAttribute(attribute.Name, attribute.Value);
        }

        var extra = name switch
        {
            "mark" => "background-color: #ffff00",
            "code" or "kbd" or "samp" or "tt" when !InsidePre(element) => "font-family: Consolas, 'Courier New', monospace",
            "center" => "text-align: center",
            "dd" => "margin-left: 36pt",
            _ => null,
        };
        if (extra is not null)
        {
            replacement.SetAttribute("style", extra + ";" + (element.GetAttribute("style") ?? string.Empty));
        }

        if (name == "font")
        {
            var style = new StringBuilder();
            if (element.GetAttribute("color") is { } color && DocsCss.Colour(color) is { } hex)
            {
                style.Append("color: ").Append(hex).Append(';');
            }

            if (element.GetAttribute("face") is { } face && face.Length < 100)
            {
                style.Append("font-family: ").Append(DocsCss.FontStack(DocsCss.FirstFamily(face))).Append(';');
            }

            if (element.GetAttribute("size") is { } size && int.TryParse(size.Trim().TrimStart('+'), out var n))
            {
                double[] sizes = [7.5, 10, 12, 13.5, 18, 24, 36];
                style.Append("font-size: ").Append(CssList.Pt(sizes[Math.Clamp(n, 1, 7) - 1])).Append(';');
            }

            replacement.SetAttribute("style", style + (element.GetAttribute("style") ?? string.Empty));
        }

        while (element.FirstChild is { } c)
        {
            replacement.AppendChild(c);
        }

        element.Parent!.ReplaceChild(replacement, element);
        return replacement;
    }

    private static bool InsidePre(IElement element)
    {
        for (var e = element.ParentElement; e is not null; e = e.ParentElement)
        {
            if (e.LocalName == "pre")
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The presentational attributes of old HTML (align, bgcolor, valign, a name) as CSS and ids.</summary>
    private static void LegacyAttributes(IElement element, CssList css)
    {
        if (element.GetAttribute("align") is { } align)
        {
            var a = align.Trim().ToLowerInvariant();
            if (element.LocalName == "img")
            {
                if (a is "left" or "right")
                {
                    css.SetDefault("float", a);
                }
            }
            else if (a is "left" or "right" or "center" or "justify")
            {
                css.SetDefault("text-align", a);
            }
        }

        if (element.GetAttribute("bgcolor") is { } bg && DocsCss.Colour(bg) is { } hex)
        {
            css.SetDefault("background-color", hex);
        }

        if (element.GetAttribute("valign")?.Trim().ToLowerInvariant() is { } v && v is "top" or "middle" or "bottom")
        {
            css.SetDefault("vertical-align", v);
        }

        if (element.LocalName == "a" && element.GetAttribute("name") is { } anchor && !element.HasAttribute("id"))
        {
            element.SetAttribute("id", anchor);
        }
    }

    /// <summary>The text of an element, without recursion (for markup nested too deeply).</summary>
    private static string TextOf(INode node)
    {
        var sb = new StringBuilder();
        var stack = new Stack<INode>();
        stack.Push(node);
        while (stack.Count > 0)
        {
            var n = stack.Pop();
            if (n is IText text)
            {
                sb.Append(text.Data);
                continue;
            }

            if (n is IElement e && Dropped.Contains(e.LocalName))
            {
                continue;
            }

            for (var i = n.ChildNodes.Length - 1; i >= 0; i--)
            {
                stack.Push(n.ChildNodes[i]);
            }
        }

        return sb.ToString();
    }

    // ------------------------------------------------------------------ checks of values

    /// <summary>An id of a block: <c>^[A-Za-z][A-Za-z0-9_-]{0,63}$</c>.</summary>
    public static bool IsId(string? id)
    {
        if (id is null || id.Length is 0 or > 64 || !char.IsAsciiLetter(id[0]))
        {
            return false;
        }

        foreach (var c in id)
        {
            if (!char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-'))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>An id of a note, a comment or a reply: <c>^[A-Za-z0-9_-]{1,40}$</c>.</summary>
    public static bool IsItemId(string? id)
    {
        if (id is null || id.Length is 0 or > 40)
        {
            return false;
        }

        foreach (var c in id)
        {
            if (!char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-'))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Turns any name into an id (a bookmark of Word, an anchor): invalid characters become <c>_</c>, a letter first.</summary>
    public static string ToId(string name)
    {
        var sb = new StringBuilder();
        foreach (var c in name)
        {
            sb.Append(char.IsAsciiLetterOrDigit(c) || c is '_' or '-' ? c : '_');
            if (sb.Length >= 63)
            {
                break;
            }
        }

        if (sb.Length == 0 || !char.IsAsciiLetter(sb[0]))
        {
            sb.Insert(0, 'x');
        }

        return sb.ToString();
    }

    /// <summary>A link target that may stay: <c>http:</c>, <c>https:</c>, <c>mailto:</c> or <c>#id</c>; anything else is null.</summary>
    public static string? SafeHref(string? href)
    {
        if (href is null)
        {
            return null;
        }

        var h = href.Trim();
        if (h.Length is 0 or > 4096 || h.Any(char.IsControl))
        {
            return null;
        }

        if (h[0] == '#')
        {
            return IsId(h[1..]) ? h : null;
        }

        var colon = h.IndexOf(':', StringComparison.Ordinal);
        if (colon <= 0)
        {
            return null;
        }

        var scheme = h[..colon].ToLowerInvariant();
        if (scheme is not ("http" or "https" or "mailto"))
        {
            return null;
        }

        if (scheme != "mailto" && (!Uri.TryCreate(h, UriKind.Absolute, out var uri) || string.IsNullOrEmpty(uri.Host)))
        {
            return null;
        }

        return scheme + h[colon..];
    }

    /// <summary>A picture that may stay: <c>data:image/(png|jpeg|gif|webp);base64,…</c> (written canonically), anything else is null.</summary>
    public static string? SafeImage(string? src)
    {
        if (src is null)
        {
            return null;
        }

        var s = src.AsSpan().Trim();
        if (!s.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var semicolon = s.IndexOf(';');
        if (semicolon < 0)
        {
            return null;
        }

        var mime = s[11..semicolon].ToString().ToLowerInvariant() switch
        {
            "png" => "png",
            "jpeg" or "jpg" or "pjpeg" => "jpeg",
            "gif" => "gif",
            "webp" => "webp",
            _ => null,
        };
        var rest = s[(semicolon + 1)..];
        if (mime is null || !rest.StartsWith("base64,", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var data = rest[7..];
        var clean = true;
        foreach (var c in data)
        {
            if (char.IsAsciiLetterOrDigit(c) || c is '+' or '/' or '=')
            {
                continue;
            }

            if (c is ' ' or '\n' or '\r' or '\t')
            {
                clean = false;
                continue;
            }

            return null;
        }

        if (data.Length == 0)
        {
            return null;
        }

        if (clean && semicolon == 11 + mime.Length && s.Length == src.Length && s.StartsWith("data:image/" + mime + ";base64,", StringComparison.Ordinal))
        {
            return src;
        }

        var sb = new StringBuilder(data.Length + 30).Append("data:image/").Append(mime).Append(";base64,");
        foreach (var c in data)
        {
            if (!char.IsWhiteSpace(c))
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    /// <summary>The ids of a comment span (space separated), or null.</summary>
    public static string? CommentIds(string? value)
    {
        if (value is null)
        {
            return null;
        }

        var ids = value.Split([' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries).Where(IsItemId).Distinct(StringComparer.Ordinal).Take(20).ToList();
        return ids.Count == 0 ? null : string.Join(' ', ids);
    }

    /// <summary>Text without control characters (tabs and line breaks stay; CR LF becomes LF) and without lone surrogates.</summary>
    public static string CleanText(string text)
    {
        var dirty = false;
        foreach (var c in text)
        {
            if ((c < ' ' && c is not ('\t' or '\n')) || char.IsSurrogate(c))
            {
                dirty = true;
                break;
            }
        }

        if (!dirty)
        {
            return text;
        }

        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c == '\r')
            {
                if (i + 1 >= text.Length || text[i + 1] != '\n')
                {
                    sb.Append('\n');
                }
            }
            else if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                sb.Append(c).Append(text[++i]);
            }
            else if (char.IsSurrogate(c))
            {
                sb.Append('�');
            }
            else if (c >= ' ' || c is '\t' or '\n')
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    // ------------------------------------------------------------------ the canonical values of styles

    /// <summary>The paragraph formatting of a style attribute in canonical values (lengths in pt, colours as hex).</summary>
    public static CssList ParagraphCss(CssList? css, HtmlReadOptions options)
    {
        var result = new CssList();
        if (css is null)
        {
            return result;
        }

        foreach (var (name, value) in css.Items)
        {
            if (!ParagraphProperties.Contains(name))
            {
                continue;
            }

            var canonical = name switch
            {
                "text-align" => value.ToLowerInvariant() switch
                {
                    "left" or "start" => "left",
                    "right" or "end" => "right",
                    "center" or "-webkit-center" => "center",
                    "justify" => "justify",
                    _ => null,
                },
                "line-height" => LineHeight(value, options),
                "background-color" => DocsCss.Colour(value),
                "text-indent" or "margin-top" or "margin-right" or "margin-bottom" or "margin-left" => Length(value, -1000, 1000),
                "padding-top" or "padding-right" or "padding-bottom" or "padding-left" => Length(value, 0, 200),
                _ => Border(value),
            };
            result.Set(name, canonical);
        }

        foreach (var side in new[] { "top", "right", "bottom", "left" })
        {
            var width = css.Get($"border-{side}-width");
            var style = css.Get($"border-{side}-style");
            var colour = css.Get($"border-{side}-color");
            if (style is not null && !result.Has("border-" + side))
            {
                result.Set("border-" + side, Border($"{width ?? "1px"} {style} {colour ?? "#000000"}"));
            }
        }

        return result;
    }

    /// <summary>A length in pt within limits, or null.</summary>
    public static string? Length(string value, double min, double max)
        => DocsCss.ToPt(value) is { } pt && pt >= min && pt <= max ? CssList.Pt(pt) : null;

    private static string? LineHeight(string value, HtmlReadOptions options)
    {
        var v = value.Trim().ToLowerInvariant();
        if (DocsCss.Percent(v) is { } percent)
        {
            var multiple = percent / 100 * (options.IsWord ? 1.2 : 1);
            return multiple is >= 0.5 and <= 10 ? CssList.Num(multiple) : null;
        }

        if (double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number))
        {
            return number is >= 0.5 and <= 10 ? CssList.Num(number) : null;
        }

        return Length(v, 1, 1000);
    }

    /// <summary>A border in canonical form (<c>0.5pt solid #000000</c>), or null for none or something unreadable.</summary>
    public static string? Border(string value)
    {
        double? width = null;
        string? style = null;
        string? colour = null;
        foreach (var part in DocsCss.SplitSpaces(value))
        {
            var p = part.ToLowerInvariant();
            if (p is "none" or "hidden")
            {
                return null;
            }

            if (p is "solid" or "dashed" or "dotted" or "double" or "groove" or "ridge" or "inset" or "outset")
            {
                style = p;
            }
            else if (DocsCss.ToPt(p) is { } w && (char.IsAsciiDigit(p[0]) || p[0] == '.' || p is "thin" or "medium" or "thick"))
            {
                width = w;
            }
            else if (DocsCss.Colour(p) is { } c)
            {
                colour = c;
            }
        }

        if (style is null || width is <= 0)
        {
            return null;
        }

        return $"{CssList.Pt(Math.Clamp(width ?? 0.75, 0.25, 6))} {style} {colour ?? "#000000"}";
    }

    // ------------------------------------------------------------------ the builder

    /// <summary>What a new paragraph gets when inline content starts one: tag, style, list attributes, paragraph formatting.</summary>
    private sealed class Template
    {
        public string Tag { get; init; } = "p";

        public string? Style { get; init; }

        public string? List { get; init; }

        public int Level { get; init; }

        public int? Start { get; set; }

        public string? Id { get; set; }

        public CssList Css { get; init; } = new();

        /// <summary>The list item this template belongs to: its first paragraph is the item, the others are continuations.</summary>
        public ListSlot? Slot { get; init; }

        /// <summary>Only inherited formatting (text-align, line-height): the template of a container.</summary>
        public bool IsContainer { get; init; }

        public DocParagraph Create()
        {
            var p = new DocParagraph(Tag) { Style = Style, Css = Css.Clone() };
            if (List is not null)
            {
                if (Slot is null || !Slot.Used)
                {
                    p.List = List;
                    p.Level = Level;
                    p.Start = Start;
                    Start = null;
                    if (Slot is not null)
                    {
                        Slot.Used = true;
                    }
                }
                else if (p.Tag == "p" && p.Style is null)
                {
                    p.Style = "ListParagraph";
                    if (Level > 0)
                    {
                        p.Css.SetDefault("margin-left", CssList.Pt((Level + 1) * 36));
                    }
                }
            }

            return p;
        }
    }

    private sealed class ListSlot
    {
        public bool Used { get; set; }
    }

    private sealed class Frame
    {
        public required List<DocBlock> Output { get; init; }

        public required Template Template { get; init; }

        public DocParagraph? Open { get; set; }

        public int ListDepth { get; init; }

        /// <summary>Whitespace alone is content (p, h1–h6, pre); elsewhere it only separates blocks.</summary>
        public bool KeepWhitespace { get; init; }

        public string? ListKind { get; init; }
    }

    private sealed class Builder(Dictionary<IElement, CssList> styles, HtmlReadOptions options)
    {
        private readonly HashSet<string> _ids = new(StringComparer.Ordinal);

        private string? _pendingId;

        public List<DocBlock> Body(IElement root)
        {
            var output = new List<DocBlock>();
            var frame = new Frame { Output = output, Template = new Template { IsContainer = true } };
            Children(root, frame, RunFormat.Plain);
            Close(frame);
            return output;
        }

        private CssList? Css(IElement element) => styles.GetValueOrDefault(element);

        private void Children(INode node, Frame frame, RunFormat format)
        {
            for (var child = node.FirstChild; child is not null; child = child.NextSibling)
            {
                Node(child, frame, format);
            }
        }

        private void Node(INode node, Frame frame, RunFormat format)
        {
            if (node is IText text)
            {
                Text(text.Data, frame, format);
                return;
            }

            if (node is not IElement element)
            {
                return;
            }

            switch (element.LocalName)
            {
                case "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "blockquote" or "pre":
                    Block(element, frame, format);
                    break;
                case "ul" or "ol":
                    List(element, frame, format);
                    break;
                case "li":
                    ListItem(element, frame, format, frame.ListDepth, frame.ListKind ?? "bullet", null);
                    break;
                case "table":
                    Table(element, frame, format);
                    break;
                case "hr":
                    Close(frame);
                    frame.Output.Add(new DocRule());
                    break;
                case "div":
                    if (element.GetAttribute("data-break") is "page" or "column")
                    {
                        Close(frame);
                        frame.Output.Add(new DocBreak(element.GetAttribute("data-break")!));
                    }

                    Container(element, frame, format);
                    break;
                case "br":
                    Add(frame, new DocLineBreak { Format = format with { Script = 0 } });
                    break;
                case "img":
                    Image(element, frame, format);
                    break;
                case "caption":
                    var caption = new Template { Style = "Caption", Css = InheritedCss(frame.Template.Css) };
                    Close(frame);
                    var captionFrame = new Frame { Output = frame.Output, Template = caption, ListDepth = frame.ListDepth };
                    Children(element, captionFrame, format);
                    Close(captionFrame);
                    break;
                case "colgroup" or "col" or "thead" or "tbody" or "tfoot" or "tr" or "td" or "th":
                    // Parts of a table outside a table: only their content counts.
                    Container(element, frame, format);
                    break;
                default:
                    Inline(element, frame, format);
                    break;
            }
        }

        private void Text(string data, Frame frame, RunFormat format)
        {
            var text = CleanText(data);
            if (text.Length == 0 || (frame.Open is null && !frame.KeepWhitespace && string.IsNullOrWhiteSpace(text)))
            {
                return;
            }

            var p = Ensure(frame);
            if (p.Inlines.Count > 0 && p.Inlines[^1] is DocText last && last.Format == format)
            {
                last.Text += text;
            }
            else
            {
                p.Inlines.Add(new DocText(text) { Format = format });
            }
        }

        private void Add(Frame frame, DocInline inline) => Ensure(frame).Inlines.Add(inline);

        private DocParagraph Ensure(Frame frame)
        {
            if (frame.Open is { } open)
            {
                return open;
            }

            var p = frame.Template.Create();
            AssignId(p, frame.Template.Id);
            frame.Template.Id = null;
            frame.Output.Add(p);
            frame.Open = p;
            return p;
        }

        private void AssignId(DocParagraph p, string? id)
        {
            id ??= _pendingId;
            if (p.Id is null && id is not null && IsId(id) && _ids.Add(id))
            {
                p.Id = id;
                if (id == _pendingId)
                {
                    _pendingId = null;
                }
            }
        }

        private void Close(Frame frame)
        {
            if (frame.Open is { } p && options.IsImport)
            {
                TrimWhitespace(p);
            }

            frame.Open = null;
        }

        /// <summary>A block element: closes the open paragraph and reads the content with a template made of the element and the frame.</summary>
        private void Block(IElement element, Frame frame, RunFormat format)
        {
            Close(frame);
            var css = Css(element);
            var outer = frame.Template;
            var name = element.LocalName;
            var dataStyle = element.GetAttribute("data-style");
            var dataList = element.GetAttribute("data-list");
            string tag;
            string? style;
            switch (name)
            {
                case "p":
                    if (dataStyle == "IntenseQuote")
                    {
                        tag = "blockquote";
                        style = "IntenseQuote";
                    }
                    else
                    {
                        tag = outer.IsContainer ? "p" : outer.Tag;
                        style = dataStyle is not null && DocsStyles.ParagraphDataStyles.Contains(dataStyle) && tag == "p" ? dataStyle : outer.Style;
                    }

                    break;
                case "blockquote":
                    tag = "blockquote";
                    style = dataStyle == "IntenseQuote" ? "IntenseQuote" : null;
                    break;
                default:
                    tag = name;
                    style = null;
                    break;
            }

            if (tag != "p" && style is not null && style != "IntenseQuote")
            {
                style = null;
            }

            string? list = null;
            var level = 0;
            int? start = null;
            ListSlot? slot = null;
            if (tag is "p" || tag[0] == 'h')
            {
                if (DocsStyles.IsListKind(dataList))
                {
                    list = dataList;
                    level = Math.Clamp(Int(element.GetAttribute("data-level")) ?? 0, 0, 8);
                    start = Int(element.GetAttribute("data-start")) is { } s and >= 1 and <= 99999 ? s : null;
                }
                else if (outer.List is not null)
                {
                    list = outer.List;
                    level = outer.Level;
                    start = outer.Slot is { Used: false } ? outer.Start : null;
                    slot = outer.Slot;
                }
            }

            var paragraphCss = outer.IsContainer ? InheritedCss(outer.Css) : outer.Css.Clone();
            foreach (var (k, v) in ParagraphCss(css, options).Items)
            {
                paragraphCss.Set(k, v);
            }

            var template = new Template
            {
                Tag = tag, Style = style, List = list, Level = level, Start = start, Css = paragraphCss, Slot = slot,
                Id = element.GetAttribute("id") ?? outer.Id,
            };
            var inner = new Frame { Output = frame.Output, Template = template, ListDepth = frame.ListDepth, ListKind = frame.ListKind, KeepWhitespace = name is not ("blockquote" or "li") };
            var count = frame.Output.Count;
            Children(element, inner, CharFormat(element, css, format));
            Close(inner);
            if (frame.Output.Count == count)
            {
                var p = template.Create();
                AssignId(p, template.Id);
                frame.Output.Add(p);
            }
        }

        /// <summary>A container (div, section …): no block of its own, inline content in it becomes paragraphs.</summary>
        private void Container(IElement element, Frame frame, RunFormat format)
        {
            Close(frame);
            var css = Css(element);
            var outer = frame.Template;
            var templateCss = InheritedCss(outer.Css);
            foreach (var (k, v) in InheritedCss(ParagraphCss(css, options)).Items)
            {
                templateCss.Set(k, v);
            }

            var template = new Template
            {
                Tag = outer.IsContainer ? "p" : outer.Tag, Style = outer.Style, List = outer.List, Level = outer.Level, Start = outer.Start,
                Slot = outer.Slot, Css = outer.IsContainer ? templateCss : outer.Css.Clone(), IsContainer = outer.IsContainer, Id = element.GetAttribute("id"),
            };
            var inner = new Frame { Output = frame.Output, Template = template, ListDepth = frame.ListDepth, ListKind = frame.ListKind };
            Children(element, inner, CharFormat(element, css, format));
            Close(inner);
        }

        private static CssList InheritedCss(CssList css)
        {
            var result = new CssList();
            result.Set("text-align", css.Get("text-align"));
            result.Set("line-height", css.Get("line-height"));
            return result;
        }

        private void List(IElement element, Frame frame, RunFormat format)
        {
            Close(frame);
            var level = Math.Min(frame.ListDepth, 8);
            var kind = ListKind(element, element.LocalName == "ol" ? "decimal" : "bullet", frame.ListKind, level);
            int? start = Int(element.GetAttribute("start")) is { } s and >= 1 and <= 99999 and not 1 ? s : null;
            var listFormat = CharFormat(element, Css(element), format);
            for (var child = element.FirstChild; child is not null; child = child.NextSibling)
            {
                switch (child)
                {
                    case IElement { LocalName: "li" } li:
                        ListItem(li, frame, listFormat, level, kind, start);
                        start = null;
                        break;
                    case IElement { LocalName: "ul" or "ol" } nested:
                        List(nested, new Frame { Output = frame.Output, Template = frame.Template, ListDepth = frame.ListDepth + 1, ListKind = kind }, listFormat);
                        break;
                    case IText text when string.IsNullOrWhiteSpace(text.Data):
                        break;
                    default:
                        // Content directly in a list (not in an item) becomes an item of its own.
                        var template = ItemTemplate(null, frame, level, kind, start);
                        start = null;
                        var inner = new Frame { Output = frame.Output, Template = template, ListDepth = level + 1, ListKind = kind };
                        Node(child, inner, listFormat);
                        Close(inner);
                        break;
                }
            }
        }

        private Template ItemTemplate(IElement? li, Frame frame, int level, string kind, int? start)
        {
            var css = InheritedCss(frame.Template.Css);
            if (li is not null)
            {
                foreach (var (k, v) in ParagraphCss(Css(li), options).Items)
                {
                    if (k is not ("margin-left" or "text-indent" or "padding-left"))
                    {
                        css.Set(k, v);
                    }
                }

                if (Int(li.GetAttribute("value")) is { } value and >= 1 and <= 99999)
                {
                    start = value;
                }
            }

            return new Template { Tag = "p", List = kind, Level = level, Start = start, Css = css, Slot = new ListSlot(), Id = li?.GetAttribute("id") };
        }

        private void ListItem(IElement li, Frame frame, RunFormat format, int level, string kind, int? start)
        {
            Close(frame);
            var itemKind = ListKind(li, kind, kind, level);
            var template = ItemTemplate(li, frame, level, itemKind, start);
            var inner = new Frame { Output = frame.Output, Template = template, ListDepth = level + 1, ListKind = itemKind };
            Children(li, inner, CharFormat(li, Css(li), format));
            Close(inner);
            if (!template.Slot!.Used && !li.ChildNodes.Any(n => n is IElement || (n is IText t && !string.IsNullOrWhiteSpace(t.Data))))
            {
                var p = template.Create();
                AssignId(p, template.Id);
                frame.Output.Add(p);
            }
        }

        /// <summary>The kind of a list or item: data-list, type, list-style-type; a marker that is the parent kind's own marker at this level keeps the parent kind.</summary>
        private string ListKind(IElement element, string fallback, string? parent, int level)
        {
            var own = element.GetAttribute("data-list");
            if (DocsStyles.IsListKind(own))
            {
                return own!;
            }

            string? format = null;
            if (Css(element)?.Get("list-style-type") is { } type)
            {
                format = MarkerKind(type);
            }

            if (format is null && element.GetAttribute("type") is { } t)
            {
                format = t.Trim() switch
                {
                    "1" => "decimal",
                    "a" => "lower-alpha",
                    "A" => "upper-alpha",
                    "i" => "lower-roman",
                    "I" => "upper-roman",
                    var x => MarkerKind(x),
                };
            }

            if (format is null)
            {
                return parent is not null && DocsStyles.BulletKinds.Contains(parent) == DocsStyles.BulletKinds.Contains(fallback) ? parent : fallback;
            }

            if (parent is not null && DocsStyles.LevelFormat(parent, level) == format)
            {
                return parent;
            }

            return format == "decimal" && parent == "outline" ? "outline" : format;
        }

        private static string? MarkerKind(string type)
        {
            var t = type.Trim().Trim('"', '\'').Trim().ToLowerInvariant();
            return t switch
            {
                "disc" or "•" or "·" => "bullet",
                "circle" or "◦" or "o" => "circle",
                "square" or "▪" or "■" or "§" => "square",
                "decimal" or "decimal-leading-zero" => "decimal",
                "lower-alpha" or "lower-latin" => "lower-alpha",
                "upper-alpha" or "upper-latin" => "upper-alpha",
                "lower-roman" => "lower-roman",
                "upper-roman" => "upper-roman",
                "–" or "-" or "—" => "dash",
                "✓" or "✔" or "☑" => "check",
                "➢" or "→" or "➤" or "►" or ">" => "arrow",
                _ => null,
            };
        }

        private void Inline(IElement element, Frame frame, RunFormat format)
        {
            var name = element.LocalName;
            if (name == "sup" && element.GetAttribute("data-fn") is { } fn)
            {
                if (IsItemId(fn))
                {
                    Add(frame, new DocNoteRef(fn) { Format = format with { Script = 0 } });
                }

                return;
            }

            if (name == "span" && element.GetAttribute("data-field") is { } field)
            {
                if (DocsStyles.Fields.Contains(field))
                {
                    Add(frame, new DocField(field) { Format = CharFormat(element, Css(element), format) });
                }

                return;
            }

            if (element.GetAttribute("id") is { } id && IsId(id) && !_ids.Contains(id))
            {
                if (frame.Open is { Id: null } open)
                {
                    AssignId(open, id);
                }
                else if (frame.Open is null)
                {
                    _pendingId = id;
                }
            }

            var f = name switch
            {
                "strong" => format with { Bold = true },
                "em" => format with { Italic = true },
                "u" => format with { Underline = true },
                "s" => format with { Strike = true },
                "sub" => format with { Script = -1 },
                "sup" => format with { Script = 1 },
                "a" when SafeHref(element.GetAttribute("href")) is { } href => format with { Href = href },
                _ => format,
            };
            if (CommentIds(element.GetAttribute("data-comment")) is { } ids)
            {
                f = f.WithComment(ids);
            }

            Children(element, frame, CharFormat(element, Css(element), f));
        }

        /// <summary>The character formatting of an element's style: fonts, colours, bold/italic/underline/strike/super/sub as flags.</summary>
        private RunFormat CharFormat(IElement element, CssList? css, RunFormat format)
        {
            if (css is null || css.Count == 0)
            {
                return format;
            }

            var f = format;
            foreach (var (name, value) in css.Items)
            {
                var v = value.ToLowerInvariant();
                switch (name)
                {
                    case "font-family":
                        f = f with { FontFamily = value };
                        break;
                    case "font-size":
                        f = f with { FontSize = DocsCss.FontSizePt(value) ?? f.FontSize };
                        break;
                    case "color":
                        f = f with { Color = DocsCss.Colour(value) ?? f.Color };
                        break;
                    case "background-color" when element.LocalName is "span" or "strong" or "em" or "u" or "s" or "sub" or "sup" or "a":
                        f = f with { Background = DocsCss.Colour(value) };
                        break;
                    case "letter-spacing":
                        f = f with { LetterSpacing = v == "normal" ? null : Length(value, -50, 50) ?? f.LetterSpacing };
                        break;
                    case "font-weight":
                        f = f with
                        {
                            Bold = v switch
                            {
                                "bold" or "bolder" or "600" or "700" or "800" or "900" => true,
                                "normal" or "lighter" or "100" or "200" or "300" or "400" or "500" => false,
                                _ => f.Bold,
                            },
                        };
                        break;
                    case "font-style":
                        f = f with { Italic = v is "italic" or "oblique" || (v != "normal" && f.Italic) };
                        break;
                    case "text-decoration" or "text-decoration-line":
                        if (v.Contains("none", StringComparison.Ordinal))
                        {
                            f = f with { Underline = false, Strike = false };
                        }

                        if (v.Contains("underline", StringComparison.Ordinal))
                        {
                            f = f with { Underline = true };
                        }

                        if (v.Contains("line-through", StringComparison.Ordinal))
                        {
                            f = f with { Strike = true };
                        }

                        break;
                    case "vertical-align" when element.LocalName is not ("td" or "th"):
                        f = f with { Script = v switch { "super" => 1, "sub" => -1, "baseline" => 0, _ => f.Script } };
                        break;
                }
            }

            return f;
        }

        private void Image(IElement element, Frame frame, RunFormat format)
        {
            if (SafeImage(element.GetAttribute("src")) is not { } src)
            {
                return;
            }

            var css = Css(element);
            var alt = CleanText(element.GetAttribute("alt") ?? string.Empty).Replace('\n', ' ');
            var image = new DocImage
            {
                Src = src,
                Alt = alt.Length > 1000 ? alt[..1000] : alt,
                Width = Pixels(css?.Get("width") ?? element.GetAttribute("width")),
                Height = Pixels(css?.Get("height") ?? element.GetAttribute("height")),
                Format = format with { Script = 0 },
            };
            var align = element.GetAttribute("data-align");
            var floating = css?.Get("float")?.ToLowerInvariant();
            image.Align = align is "left" or "right" or "center" ? align : floating is "left" or "right" ? floating : null;
            Add(frame, image);
        }

        private static double? Pixels(string? value)
        {
            if (value is null || value.Trim().EndsWith('%'))
            {
                return null;
            }

            return DocsCss.ToPx(value) is { } px && px is >= 1 and <= 20000 ? Math.Round(px, 2) : null;
        }

        private void Table(IElement element, Frame frame, RunFormat format)
        {
            Close(frame);
            var css = Css(element);
            var table = new DocTable();
            var style = element.GetAttribute("data-table-style");
            table.TableStyle = style is not null && DocsStyles.TableStyles.Contains(style) ? style : "grid";
            table.Width = TableWidth(css?.Get("width") ?? element.GetAttribute("width"));
            var tableFormat = CharFormat(element, css, format);
            var pixelColumns = new List<double?>();
            foreach (var part in element.Children)
            {
                switch (part.LocalName)
                {
                    case "caption":
                        Node(part, frame, tableFormat);
                        Close(frame);
                        break;
                    case "colgroup":
                        var cols = part.Children.Where(c => c.LocalName == "col").ToList();
                        if (cols.Count == 0)
                        {
                            AddColumns(table, pixelColumns, part, Int(part.GetAttribute("span")) ?? 1);
                        }

                        foreach (var col in cols)
                        {
                            AddColumns(table, pixelColumns, col, Int(col.GetAttribute("span")) ?? 1);
                        }

                        break;
                    case "col":
                        AddColumns(table, pixelColumns, part, Int(part.GetAttribute("span")) ?? 1);
                        break;
                    case "thead":
                        Rows(part, table, header: true, tableFormat);
                        break;
                    case "tbody" or "tfoot":
                        Rows(part, table, header: false, tableFormat);
                        break;
                    case "tr":
                        Row(part, table, header: false, tableFormat);
                        break;
                }
            }

            if (table.Rows.Count == 0)
            {
                return;
            }

            // Header rows stay only at the top.
            var top = true;
            foreach (var row in table.Rows)
            {
                top &= row.Header;
                row.Header = top;
            }

            if (table.Columns.Count == 0)
            {
                ColumnsFromCells(element, table);
            }
            else if (pixelColumns.Any(p => p is not null))
            {
                PixelsToPercent(table, pixelColumns);
            }

            frame.Output.Add(table);
        }

        private void AddColumns(DocTable table, List<double?> pixels, IElement col, int span)
        {
            var width = Css(col)?.Get("width") ?? col.GetAttribute("width");
            for (var i = 0; i < Math.Clamp(span, 1, 1000); i++)
            {
                if (DocsCss.Percent(width) is { } p && p is > 0 and <= 100)
                {
                    table.Columns.Add(CssList.Num(p) + "%");
                    pixels.Add(null);
                }
                else
                {
                    table.Columns.Add(null);
                    pixels.Add(width is null ? null : DocsCss.ToPx(width));
                }
            }
        }

        private static void PixelsToPercent(DocTable table, List<double?> pixels)
        {
            var total = pixels.Sum(p => p ?? 0);
            if (total <= 0 || table.Columns.Any(c => c is not null))
            {
                return;
            }

            for (var i = 0; i < table.Columns.Count; i++)
            {
                table.Columns[i] = pixels[i] is { } px ? CssList.Num(px / total * 100) + "%" : null;
            }
        }

        /// <summary>A table without columns: the widths of the cells of the first row (HTML of Word and of many web pages).</summary>
        private void ColumnsFromCells(IElement element, DocTable table)
        {
            var firstRow = element.QuerySelector("tr");
            if (firstRow is null || options is { IsImport: false })
            {
                return;
            }

            var widths = new List<(double? Percent, double? Pixels, int Span)>();
            foreach (var cell in firstRow.Children.Where(c => c.LocalName is "td" or "th"))
            {
                var width = Css(cell)?.Get("width") ?? cell.GetAttribute("width");
                widths.Add((DocsCss.Percent(width), width is null ? null : DocsCss.ToPx(width), Math.Clamp(Int(cell.GetAttribute("colspan")) ?? 1, 1, 1000)));
            }

            if (widths.Count == 0 || widths.Any(w => w.Span != 1))
            {
                return;
            }

            if (widths.All(w => w.Percent is > 0))
            {
                table.Columns.AddRange(widths.Select(w => (string?)(CssList.Num(w.Percent!.Value) + "%")));
            }
            else if (widths.All(w => w.Pixels is > 0))
            {
                var total = widths.Sum(w => w.Pixels!.Value);
                table.Columns.AddRange(widths.Select(w => (string?)(CssList.Num(w.Pixels!.Value / total * 100) + "%")));
            }
        }

        private static string? TableWidth(string? value)
        {
            if (value is null)
            {
                return null;
            }

            if (DocsCss.Percent(value) is { } p)
            {
                return p is > 0 and <= 100 ? CssList.Num(p) + "%" : null;
            }

            return DocsCss.ToPt(value) is { } pt && pt is > 0 and < 3000 ? CssList.Pt(pt) : null;
        }

        private void Rows(IElement group, DocTable table, bool header, RunFormat format)
        {
            var groupFormat = CharFormat(group, Css(group), format);
            foreach (var row in group.Children.Where(c => c.LocalName == "tr"))
            {
                Row(row, table, header, groupFormat);
            }
        }

        private void Row(IElement tr, DocTable table, bool header, RunFormat format)
        {
            if (table.Rows.Count >= 10000)
            {
                return;
            }

            var row = new DocRow { Header = header };
            var rowCss = Css(tr);
            var rowFormat = CharFormat(tr, rowCss, format);
            foreach (var cellElement in tr.Children.Where(c => c.LocalName is "td" or "th"))
            {
                var css = Css(cellElement);
                var cell = new DocCell
                {
                    HeaderCell = header || cellElement.LocalName == "th",
                    ColSpan = Math.Clamp(Int(cellElement.GetAttribute("colspan")) ?? 1, 1, 1000),
                    RowSpan = Math.Clamp(Int(cellElement.GetAttribute("rowspan")) ?? 1, 1, 1000),
                };
                var background = DocsCss.Colour(css?.Get("background-color")) ?? DocsCss.Colour(rowCss?.Get("background-color"));
                cell.Css.Set("background-color", background);

                // Top is where the text of a cell starts anyway (as in Word).
                var verticalAlign = (css?.Get("vertical-align") ?? rowCss?.Get("vertical-align"))?.ToLowerInvariant();
                cell.Css.Set("vertical-align", verticalAlign is "middle" or "bottom" ? verticalAlign : null);
                foreach (var (k, v) in ParagraphCss(css, options).Items)
                {
                    if (k.StartsWith("border-", StringComparison.Ordinal))
                    {
                        cell.Css.Set(k, v);
                    }
                }

                var cellFrame = new Frame
                {
                    Output = cell.Blocks,
                    Template = new Template { IsContainer = true, Css = InheritedCss(ParagraphCss(css, options)) },
                };
                Children(cellElement, cellFrame, CharFormat(cellElement, css, rowFormat));
                Close(cellFrame);
                row.Cells.Add(cell);
            }

            if (row.Cells.Count > 0)
            {
                table.Rows.Add(row);
            }
        }

        /// <summary>Whitespace of a foreign file: none at the start and the end of a paragraph or around a line break, never two in a row.</summary>
        private static void TrimWhitespace(DocParagraph p)
        {
            if (p.Tag == "pre")
            {
                return;
            }

            var atStart = true;
            for (var i = 0; i < p.Inlines.Count; i++)
            {
                switch (p.Inlines[i])
                {
                    case DocText t:
                        if (atStart)
                        {
                            t.Text = t.Text.TrimStart(' ');
                        }

                        atStart = t.Text.Length == 0 ? atStart : t.Text[^1] == ' ';
                        break;
                    case DocLineBreak:
                        atStart = true;
                        TrimEndBefore(p, i);
                        break;
                    default:
                        atStart = false;
                        break;
                }
            }

            TrimEndBefore(p, p.Inlines.Count);
            p.Inlines.RemoveAll(i => i is DocText { Text.Length: 0 });
        }

        private static void TrimEndBefore(DocParagraph p, int index)
        {
            for (var j = index - 1; j >= 0; j--)
            {
                if (p.Inlines[j] is not DocText t)
                {
                    return;
                }

                t.Text = t.Text.TrimEnd(' ');
                if (t.Text.Length > 0)
                {
                    return;
                }
            }
        }

        private static int? Int(string? value)
            => int.TryParse(value?.Trim(), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    // ------------------------------------------------------------------ the writer

    /// <summary>Writes blocks as canonical Docs HTML (<c>""</c> for none).</summary>
    public static string Write(IEnumerable<DocBlock> blocks)
    {
        var sb = new StringBuilder();
        foreach (var block in blocks)
        {
            WriteBlock(sb, block);
        }

        return sb.ToString();
    }

    private static void WriteBlock(StringBuilder sb, DocBlock block)
    {
        switch (block)
        {
            case DocParagraph p:
                sb.Append('<').Append(p.Tag);
                Attribute(sb, "id", p.Id);
                Attribute(sb, "data-style", p.Style);
                if (p.List is not null)
                {
                    Attribute(sb, "data-list", p.List);
                    Attribute(sb, "data-level", p.Level > 0 ? p.Level.ToString(System.Globalization.CultureInfo.InvariantCulture) : null);
                    Attribute(sb, "data-start", p.Start?.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }

                Attribute(sb, "style", p.Css.Count > 0 ? Ordered(p.Css, ParagraphOrder) : null);
                sb.Append('>');
                if (p.Tag == "pre" && p.Inlines.FirstOrDefault() is DocText { Text: [ '\n', ..] })
                {
                    sb.Append('\n');
                }

                WriteInlines(sb, p.Inlines);
                sb.Append("</").Append(p.Tag).Append('>');
                break;
            case DocTable table:
                WriteTable(sb, table);
                break;
            case DocRule:
                sb.Append("<hr>");
                break;
            case DocBreak br:
                sb.Append("<div data-break=\"").Append(br.Kind).Append("\"></div>");
                break;
        }
    }

    private static void WriteTable(StringBuilder sb, DocTable table)
    {
        sb.Append("<table");
        Attribute(sb, "data-table-style", table.TableStyle);
        Attribute(sb, "style", table.Width is null ? null : "width: " + table.Width);
        sb.Append('>');
        if (table.Columns.Any(c => c is not null))
        {
            sb.Append("<colgroup>");
            foreach (var column in table.Columns)
            {
                sb.Append("<col");
                Attribute(sb, "style", column is null ? null : "width: " + column);
                sb.Append('>');
            }

            sb.Append("</colgroup>");
        }

        var header = table.Rows.TakeWhile(r => r.Header).ToList();
        if (header.Count > 0)
        {
            sb.Append("<thead>");
            foreach (var row in header)
            {
                WriteRow(sb, row, true);
            }

            sb.Append("</thead>");
        }

        if (table.Rows.Count > header.Count)
        {
            sb.Append("<tbody>");
            foreach (var row in table.Rows.Skip(header.Count))
            {
                WriteRow(sb, row, false);
            }

            sb.Append("</tbody>");
        }

        sb.Append("</table>");
    }

    private static void WriteRow(StringBuilder sb, DocRow row, bool header)
    {
        sb.Append("<tr>");
        foreach (var cell in row.Cells)
        {
            var tag = header || cell.HeaderCell ? "th" : "td";
            sb.Append('<').Append(tag);
            Attribute(sb, "colspan", cell.ColSpan > 1 ? cell.ColSpan.ToString(System.Globalization.CultureInfo.InvariantCulture) : null);
            Attribute(sb, "rowspan", cell.RowSpan > 1 ? cell.RowSpan.ToString(System.Globalization.CultureInfo.InvariantCulture) : null);
            Attribute(sb, "style", cell.Css.Count > 0 ? Ordered(cell.Css, CellOrder) : null);
            sb.Append('>');
            if (cell.Blocks.Count == 0)
            {
                sb.Append("<p></p>");
            }

            foreach (var block in cell.Blocks)
            {
                WriteBlock(sb, block);
            }

            sb.Append("</").Append(tag).Append('>');
        }

        sb.Append("</tr>");
    }

    /// <summary>The wrappers of inline content in the canonical nesting, outermost first.</summary>
    private static List<(int Level, string Open, string Close)> Chain(RunFormat f)
    {
        var chain = new List<(int, string, string)>(4);
        if (f.Href is not null)
        {
            chain.Add((0, "<a href=\"" + EscapeAttribute(f.Href) + "\">", "</a>"));
        }

        if (f.Comment is not null)
        {
            chain.Add((1, "<span data-comment=\"" + EscapeAttribute(f.Comment) + "\">", "</span>"));
        }

        if (f.HasSpanStyle)
        {
            chain.Add((2, "<span style=\"" + EscapeAttribute(f.SpanStyle()) + "\">", "</span>"));
        }

        if (f.Bold)
        {
            chain.Add((3, "<strong>", "</strong>"));
        }

        if (f.Italic)
        {
            chain.Add((4, "<em>", "</em>"));
        }

        if (f.Underline)
        {
            chain.Add((5, "<u>", "</u>"));
        }

        if (f.Strike)
        {
            chain.Add((6, "<s>", "</s>"));
        }

        if (f.Script != 0)
        {
            chain.Add(f.Script > 0 ? (7, "<sup>", "</sup>") : (7, "<sub>", "</sub>"));
        }

        return chain;
    }

    /// <summary>Writes inline content: neighbours with the same formatting share their wrappers, outer wrappers as long as possible.</summary>
    public static void WriteInlines(StringBuilder sb, IEnumerable<DocInline> inlines, Func<DocInline, string?>? replace = null)
    {
        var open = new List<(int Level, string Open, string Close)>();
        foreach (var inline in inlines)
        {
            var chain = Chain(inline.Format);
            var common = 0;
            while (common < open.Count && common < chain.Count && open[common].Open == chain[common].Open)
            {
                common++;
            }

            for (var i = open.Count - 1; i >= common; i--)
            {
                sb.Append(open[i].Close);
            }

            open.RemoveRange(common, open.Count - common);
            for (var i = common; i < chain.Count; i++)
            {
                sb.Append(chain[i].Open);
                open.Add(chain[i]);
            }

            if (replace?.Invoke(inline) is { } html)
            {
                sb.Append(html);
                continue;
            }

            switch (inline)
            {
                case DocText t:
                    EscapeText(sb, t.Text);
                    break;
                case DocLineBreak:
                    sb.Append("<br>");
                    break;
                case DocImage img:
                    sb.Append("<img src=\"").Append(img.Src).Append('"');
                    Attribute(sb, "alt", img.Alt);
                    var size = new CssList();
                    size.Set("width", img.Width is { } w ? CssList.Num(w) + "px" : null);
                    size.Set("height", img.Height is { } h ? CssList.Num(h) + "px" : null);
                    Attribute(sb, "style", size.Count > 0 ? size.ToString() : null);
                    Attribute(sb, "data-align", img.Align);
                    sb.Append('>');
                    break;
                case DocNoteRef note:
                    sb.Append("<sup data-fn=\"").Append(note.Id).Append("\"></sup>");
                    break;
                case DocField field:
                    sb.Append("<span data-field=\"").Append(field.Kind).Append("\"></span>");
                    break;
            }
        }

        for (var i = open.Count - 1; i >= 0; i--)
        {
            sb.Append(open[i].Close);
        }
    }

    /// <summary>The declarations in the given order (others after them, as they come).</summary>
    private static string Ordered(CssList css, string[] order)
    {
        if (css.Count == 1)
        {
            return css.ToString();
        }

        var ordered = new CssList();
        foreach (var name in order)
        {
            ordered.Set(name, css.Get(name));
        }

        foreach (var (name, value) in css.Items)
        {
            ordered.SetDefault(name, value);
        }

        return ordered.ToString();
    }

    private static void Attribute(StringBuilder sb, string name, string? value)
    {
        if (value is null)
        {
            return;
        }

        sb.Append(' ').Append(name).Append("=\"").Append(EscapeAttribute(value)).Append('"');
    }

    public static string EscapeAttribute(string value)
    {
        if (value.AsSpan().IndexOfAny('&', '"') < 0)
        {
            return value;
        }

        return value.Replace("&", "&amp;", StringComparison.Ordinal).Replace("\"", "&quot;", StringComparison.Ordinal);
    }

    public static void EscapeText(StringBuilder sb, string text)
    {
        foreach (var c in text)
        {
            switch (c)
            {
                case '&':
                    sb.Append("&amp;");
                    break;
                case '<':
                    sb.Append("&lt;");
                    break;
                case '>':
                    sb.Append("&gt;");
                    break;
                default:
                    sb.Append(c);
                    break;
            }
        }
    }

    public static string EscapeText(string text)
    {
        var sb = new StringBuilder(text.Length + 16);
        EscapeText(sb, text);
        return sb.ToString();
    }

    // ------------------------------------------------------------------ walking the model

    /// <summary>Every paragraph, also the ones in table cells, in the order of the document.</summary>
    public static IEnumerable<DocParagraph> Paragraphs(IEnumerable<DocBlock> blocks)
    {
        foreach (var block in blocks)
        {
            if (block is DocParagraph p)
            {
                yield return p;
            }
            else if (block is DocTable table)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var inner in Paragraphs(cell.Blocks))
                        {
                            yield return inner;
                        }
                    }
                }
            }
        }
    }

    /// <summary>The plain text of blocks: one line per paragraph, table cells separated by tabs.</summary>
    public static void PlainText(StringBuilder sb, IEnumerable<DocBlock> blocks, int max)
    {
        foreach (var block in blocks)
        {
            if (sb.Length >= max)
            {
                return;
            }

            switch (block)
            {
                case DocParagraph p:
                    sb.Append(p.PlainText()).Append('\n');
                    break;
                case DocTable table:
                    foreach (var row in table.Rows)
                    {
                        var first = true;
                        foreach (var cell in row.Cells)
                        {
                            if (!first)
                            {
                                sb.Append('\t');
                            }

                            first = false;
                            sb.Append(string.Join(' ', Paragraphs(cell.Blocks).Select(p => p.PlainText().Replace('\n', ' '))));
                        }

                        sb.Append('\n');
                    }

                    break;
            }
        }
    }
}
