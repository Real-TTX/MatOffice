using System.Collections.Concurrent;
using System.Text;
using System.Text.Json.Nodes;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The templates of Docs. Each one (except blank) is an embedded JSON file <c>Resources/Templates/docs/&lt;id&gt;.json</c> holding the content
/// twice: <c>{ "en": { content }, "de": { content } }</c>. To keep the files readable, every HTML field of a template (<c>html</c>, <c>header</c>,
/// <c>footer</c>, <c>firstHeader</c>, <c>firstFooter</c>, <c>notes[].html</c>) may be an array of strings (one block per line) instead of one
/// string; <see cref="Create"/> joins them. The names and descriptions are text keys.
/// </summary>
internal static class DocsTemplates
{
    public static IReadOnlyList<KindTemplate> All { get; } =
    [
        new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-docs"),
        new("letter", "Docs.Template.Letter", "Docs.Template.Letter.Text", "mail"),
        new("report", "Docs.Template.Report", "Docs.Template.Report.Text", "book"),
        new("meeting", "Docs.Template.Meeting", "Docs.Template.Meeting.Text", "users"),
        new("resume", "Docs.Template.Resume", "Docs.Template.Resume.Text", "user"),
        new("proposal", "Docs.Template.Proposal", "Docs.Template.Proposal.Text", "briefcase"),
        new("invoice", "Docs.Template.Invoice", "Docs.Template.Invoice.Text", "dollar"),
    ];

    private static readonly string[] HtmlFields = ["html", "header", "footer", "firstHeader", "firstFooter"];

    /// <summary>The content of every template and language read so far (null: the template file is missing).</summary>
    private static readonly ConcurrentDictionary<string, string?> Contents = new(StringComparer.Ordinal);

    /// <summary>The content of a new document: the template in German for <c>de</c>, in English otherwise; blank for an unknown id.</summary>
    public static string Create(string templateId, string language)
    {
        var lang = language.StartsWith("de", StringComparison.OrdinalIgnoreCase) ? "de" : "en";
        if (templateId != "blank" && All.Any(t => t.Id == templateId) && Contents.GetOrAdd(templateId + "/" + lang, Load) is { } content)
        {
            return content;
        }

        return new DocsDocument { Language = lang }.ToJson();
    }

    /// <summary>Reads <c>id/lang</c> from its embedded file (the English content when the language is missing) and joins its HTML lines.</summary>
    private static string? Load(string key)
    {
        var (id, lang) = (key[..key.IndexOf('/', StringComparison.Ordinal)], key[(key.IndexOf('/', StringComparison.Ordinal) + 1)..]);
        var assembly = typeof(DocsTemplates).Assembly;
        var name = assembly.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith($".Resources.Templates.docs.{id}.json", StringComparison.Ordinal));
        if (name is null)
        {
            return null;
        }

        using var stream = assembly.GetManifestResourceStream(name)!;
        if (JsonNode.Parse(stream) is not JsonObject file || (file[lang] ?? file["en"]) is not JsonObject content)
        {
            return null;
        }

        foreach (var field in HtmlFields)
        {
            JoinLines(content, field);
        }

        if (content["notes"] is JsonArray notes)
        {
            foreach (var note in notes.OfType<JsonObject>())
            {
                JoinLines(note, "html");
            }
        }

        return content.ToJsonString();
    }

    /// <summary>An HTML field written as an array of lines becomes one string.</summary>
    private static void JoinLines(JsonObject owner, string field)
    {
        if (owner[field] is not JsonArray lines)
        {
            return;
        }

        var sb = new StringBuilder();
        foreach (var line in lines)
        {
            if (line is JsonValue value && value.TryGetValue<string>(out var text))
            {
                sb.Append(text);
            }
        }

        owner[field] = sb.ToString();
    }
}
