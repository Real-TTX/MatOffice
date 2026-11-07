using System.Collections.Concurrent;
using System.Text.Json.Nodes;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The templates of Docs. Each one (except blank) is an embedded JSON file <c>Resources/Templates/docs/&lt;id&gt;.json</c> holding the content
/// twice: <c>{ "en": { content }, "de": { content } }</c>. The names and descriptions are text keys.
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

    private static readonly ConcurrentDictionary<string, JsonObject?> Files = new(StringComparer.Ordinal);

    /// <summary>The content of a new document: the template in German for <c>de</c>, in English otherwise; blank for an unknown id.</summary>
    public static string Create(string templateId, string language)
    {
        var lang = language.StartsWith("de", StringComparison.OrdinalIgnoreCase) ? "de" : "en";
        if (templateId != "blank" && All.Any(t => t.Id == templateId) && Load(templateId) is { } file && (file[lang] ?? file["en"]) is JsonObject content)
        {
            return content.ToJsonString();
        }

        return new DocsDocument { Language = lang }.ToJson();
    }

    private static JsonObject? Load(string id) => Files.GetOrAdd(id, static key =>
    {
        var assembly = typeof(DocsTemplates).Assembly;
        var name = assembly.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith($".Resources.Templates.docs.{key}.json", StringComparison.Ordinal));
        if (name is null)
        {
            return null;
        }

        using var stream = assembly.GetManifestResourceStream(name)!;
        return JsonNode.Parse(stream) as JsonObject;
    });
}
