using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace MatOffice.Services.Formats.Docs;

/// <summary>The page of a document; lengths in millimetres, <see cref="Width"/>/<see cref="Height"/> the portrait size.</summary>
internal sealed class PageSetup
{
    public static readonly IReadOnlyDictionary<string, (double Width, double Height)> Sizes = new Dictionary<string, (double, double)>(StringComparer.Ordinal)
    {
        ["A4"] = (210, 297),
        ["A5"] = (148, 210),
        ["A3"] = (297, 420),
        ["Letter"] = (215.9, 279.4),
        ["Legal"] = (215.9, 355.6),
    };

    public string Size { get; set; } = "A4";

    public double Width { get; set; } = 210;

    public double Height { get; set; } = 297;

    public bool Landscape { get; set; }

    public double Top { get; set; } = 25;

    public double Right { get; set; } = 25;

    public double Bottom { get; set; } = 20;

    public double Left { get; set; } = 25;

    public double Header { get; set; } = 12.5;

    public double Footer { get; set; } = 12.5;

    public int Columns { get; set; } = 1;

    public bool DifferentFirst { get; set; }

    /// <summary>The width of the paper as it lies (landscape swapped).</summary>
    public double PaperWidth => Landscape ? Height : Width;

    public double PaperHeight => Landscape ? Width : Height;

    /// <summary>The width of the text between the margins.</summary>
    public double TextWidth => PaperWidth - Left - Right;

    /// <summary>The named size of a portrait paper size (±2 mm), or Custom.</summary>
    public static string SizeOf(double width, double height)
    {
        foreach (var (name, size) in Sizes)
        {
            if (Math.Abs(size.Width - width) <= 2 && Math.Abs(size.Height - height) <= 2)
            {
                return name;
            }
        }

        return "Custom";
    }

    /// <summary>Applies the rules: known sizes have their own measures, the rest is clamped, at least 20 mm of text remain.</summary>
    public void Normalize()
    {
        if (Sizes.TryGetValue(Size, out var size))
        {
            (Width, Height) = size;
        }
        else if (Size == "Custom")
        {
            Width = Round(Math.Clamp(double.IsFinite(Width) ? Width : 210, 50, 1000));
            Height = Round(Math.Clamp(double.IsFinite(Height) ? Height : 297, 50, 1000));
        }
        else
        {
            Size = "A4";
            (Width, Height) = Sizes["A4"];
        }

        Top = Clamp(Top);
        Right = Clamp(Right);
        Bottom = Clamp(Bottom);
        Left = Clamp(Left);
        Header = Clamp(Header);
        Footer = Clamp(Footer);
        (Left, Right) = Fit(Left, Right, PaperWidth);
        (Top, Bottom) = Fit(Top, Bottom, PaperHeight);
        Columns = Math.Clamp(Columns, 1, 3);
    }

    private static (double, double) Fit(double a, double b, double paper)
    {
        var available = paper - 20;
        if (a + b <= available || a + b <= 0)
        {
            return (a, b);
        }

        var factor = Math.Max(0, available) / (a + b);
        return (Math.Floor(a * factor * 100) / 100, Math.Floor(b * factor * 100) / 100);
    }

    private static double Clamp(double value) => Round(Math.Clamp(double.IsFinite(value) ? value : 0, 0, 150));

    private static double Round(double value) => Math.Round(value, 2, MidpointRounding.AwayFromZero);
}

internal sealed class DocNote
{
    public required string Id { get; init; }

    public List<DocBlock> Blocks { get; init; } = [];
}

internal sealed class DocComment
{
    public required string Id { get; init; }

    public string Author { get; set; } = string.Empty;

    public string Date { get; set; } = string.Empty;

    public string Text { get; set; } = string.Empty;

    public bool Done { get; set; }

    public List<DocComment> Replies { get; } = [];
}

/// <summary>
/// A Docs document in memory: what <see cref="DocsHandler.Normalize"/> checks and what the converters read and write. <see cref="FromJson"/> reads
/// (and sanitizes) the content JSON, <see cref="ToJson"/> writes the canonical content JSON.
/// </summary>
internal sealed partial class DocsDocument
{
    public const int MaxNotes = 2000;

    public const int MaxComments = 2000;

    public const int MaxReplies = 500;

    private static readonly JsonSerializerOptions JsonOptions = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    public PageSetup Page { get; set; } = new();

    public List<DocBlock> Body { get; set; } = [];

    public List<DocBlock> Header { get; set; } = [];

    public List<DocBlock> Footer { get; set; } = [];

    public List<DocBlock> FirstHeader { get; set; } = [];

    public List<DocBlock> FirstFooter { get; set; } = [];

    public List<DocNote> Notes { get; set; } = [];

    public List<DocComment> Comments { get; set; } = [];

    public string Language { get; set; } = "en";

    public bool Spellcheck { get; set; } = true;

    /// <summary>Reads content JSON; every HTML is sanitized. Throws <see cref="FormatException"/> for content that is not a Docs document.</summary>
    public static DocsDocument FromJson(string json) => FromJson(Kinds.ParseObject(json));

    public static DocsDocument FromJson(JsonObject root)
    {
        var doc = new DocsDocument
        {
            Page = ReadPage(root["page"]),
            Body = DocsHtml.Parse(Html(root, "html")),
            Header = DocsHtml.Parse(Html(root, "header")),
            Footer = DocsHtml.Parse(Html(root, "footer")),
            FirstHeader = DocsHtml.Parse(Html(root, "firstHeader")),
            FirstFooter = DocsHtml.Parse(Html(root, "firstFooter")),
        };

        if (root["notes"] is JsonArray notes)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in notes)
            {
                if (doc.Notes.Count >= MaxNotes)
                {
                    break;
                }

                if (item is JsonObject note && Str(note["id"]) is { } id && DocsHtml.IsItemId(id) && ids.Add(id))
                {
                    doc.Notes.Add(new DocNote { Id = id, Blocks = DocsHtml.Parse(Str(note["html"])) });
                }
            }
        }

        if (root["comments"] is JsonArray comments)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in comments)
            {
                if (doc.Comments.Count >= MaxComments)
                {
                    break;
                }

                if (ReadComment(item, ids) is { } comment)
                {
                    if (item!["replies"] is JsonArray replies)
                    {
                        foreach (var r in replies)
                        {
                            if (comment.Replies.Count >= MaxReplies)
                            {
                                break;
                            }

                            if (ReadComment(r, ids) is { } reply)
                            {
                                comment.Replies.Add(reply);
                            }
                        }
                    }

                    doc.Comments.Add(comment);
                }
            }
        }

        if (root["settings"] is JsonObject settings)
        {
            if (Str(settings["language"]) is { } language && LanguageRegex().IsMatch(language))
            {
                doc.Language = language;
            }

            if (settings["spellcheck"] is JsonValue spell && spell.TryGetValue<bool>(out var check))
            {
                doc.Spellcheck = check;
            }
        }

        doc.DropUnreferencedNotes();
        return doc;
    }

    /// <summary>Notes that the body does not reference are gone.</summary>
    public void DropUnreferencedNotes()
    {
        if (Notes.Count == 0)
        {
            return;
        }

        var referenced = new HashSet<string>(StringComparer.Ordinal);
        foreach (var p in DocsHtml.Paragraphs(Body))
        {
            foreach (var inline in p.Inlines)
            {
                if (inline is DocNoteRef r)
                {
                    referenced.Add(r.Id);
                }
            }
        }

        Notes.RemoveAll(n => !referenced.Contains(n.Id));
    }

    /// <summary>The canonical content JSON.</summary>
    public string ToJson()
    {
        Page.Normalize();
        var body = DocsHtml.Write(Body);
        var root = new JsonObject
        {
            ["version"] = 1,
            ["page"] = new JsonObject
            {
                ["size"] = Page.Size,
                ["width"] = Page.Width,
                ["height"] = Page.Height,
                ["orientation"] = Page.Landscape ? "landscape" : "portrait",
                ["margins"] = new JsonObject { ["top"] = Page.Top, ["right"] = Page.Right, ["bottom"] = Page.Bottom, ["left"] = Page.Left },
                ["header"] = Page.Header,
                ["footer"] = Page.Footer,
                ["columns"] = Page.Columns,
                ["differentFirst"] = Page.DifferentFirst,
            },
            ["html"] = body.Length == 0 ? "<p></p>" : body,
            ["header"] = DocsHtml.Write(Header),
            ["footer"] = DocsHtml.Write(Footer),
            ["firstHeader"] = DocsHtml.Write(FirstHeader),
            ["firstFooter"] = DocsHtml.Write(FirstFooter),
            ["notes"] = new JsonArray(Notes.Select(n => (JsonNode)new JsonObject
            {
                ["id"] = n.Id,
                ["html"] = DocsHtml.Write(n.Blocks) is { Length: > 0 } html ? html : "<p></p>",
            }).ToArray()),
            ["comments"] = new JsonArray(Comments.Select(c =>
            {
                var o = CommentJson(c);
                o["done"] = c.Done;
                o["replies"] = new JsonArray(c.Replies.Select(r => (JsonNode)CommentJson(r)).ToArray());
                return (JsonNode)o;
            }).ToArray()),
            ["settings"] = new JsonObject { ["language"] = Language, ["spellcheck"] = Spellcheck },
        };
        return root.ToJsonString(JsonOptions);
    }

    private static JsonObject CommentJson(DocComment c) => new() { ["id"] = c.Id, ["author"] = c.Author, ["date"] = c.Date, ["text"] = c.Text };

    private static DocComment? ReadComment(JsonNode? node, HashSet<string> ids)
    {
        if (node is not JsonObject o || Str(o["id"]) is not { } id || !DocsHtml.IsItemId(id) || !ids.Add(id))
        {
            return null;
        }

        var date = Str(o["date"]) ?? string.Empty;
        return new DocComment
        {
            Id = id,
            Author = Limit(DocsHtml.CleanText(Str(o["author"]) ?? string.Empty).Replace('\n', ' ').Replace('\t', ' '), 200),
            Date = date.Length <= 40 && DateRegex().IsMatch(date) ? date : string.Empty,
            Text = Limit(DocsHtml.CleanText(Str(o["text"]) ?? string.Empty), 10_000),
            Done = o["done"] is JsonValue v && v.TryGetValue<bool>(out var done) && done,
        };
    }

    private static string Limit(string text, int max) => text.Length <= max ? text : text[..max];

    /// <summary>An HTML field: missing or null is empty, any other type than a string is an error.</summary>
    private static string? Html(JsonObject root, string name)
    {
        var node = root[name];
        if (node is null)
        {
            return null;
        }

        if (node is JsonValue value && value.TryGetValue<string>(out var html))
        {
            return html;
        }

        throw new FormatException(Localization.Text.T("Documents.Invalid"));
    }

    private static string? Str(JsonNode? node) => node is JsonValue value && value.TryGetValue<string>(out var s) ? s : null;

    private static double? Number(JsonNode? node)
    {
        if (node is not JsonValue value)
        {
            return null;
        }

        if (value.TryGetValue<double>(out var d))
        {
            return double.IsFinite(d) ? d : null;
        }

        return value.TryGetValue<string>(out var s) && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && double.IsFinite(parsed)
            ? parsed
            : null;
    }

    private static PageSetup ReadPage(JsonNode? node)
    {
        var page = new PageSetup();
        if (node is not JsonObject o)
        {
            return page;
        }

        page.Size = Str(o["size"]) is { } size && (PageSetup.Sizes.ContainsKey(size) || size == "Custom") ? size : "A4";
        page.Width = Number(o["width"]) ?? page.Width;
        page.Height = Number(o["height"]) ?? page.Height;
        page.Landscape = Str(o["orientation"]) == "landscape";
        if (o["margins"] is JsonObject m)
        {
            page.Top = Number(m["top"]) ?? page.Top;
            page.Right = Number(m["right"]) ?? page.Right;
            page.Bottom = Number(m["bottom"]) ?? page.Bottom;
            page.Left = Number(m["left"]) ?? page.Left;
        }

        page.Header = Number(o["header"]) ?? page.Header;
        page.Footer = Number(o["footer"]) ?? page.Footer;
        page.Columns = (int)Math.Round(Math.Clamp(Number(o["columns"]) ?? 1, 1, 3));
        page.DifferentFirst = o["differentFirst"] is JsonValue first && first.TryGetValue<bool>(out var f) && f;
        page.Normalize();
        return page;
    }

    [GeneratedRegex("^[a-z]{2}(-[A-Z]{2})?$")]
    private static partial Regex LanguageRegex();

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2}(\.\d{1,9})?)?)?(Z|[+-]\d{2}:?\d{2})?$")]
    private static partial Regex DateRegex();
}
