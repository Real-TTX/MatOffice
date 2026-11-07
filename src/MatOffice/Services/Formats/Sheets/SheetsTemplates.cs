using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// The templates of Sheets: content JSON files in <c>Resources/Templates/sheets/&lt;id&gt;.json</c> (embedded). Any value that is an object with exactly the
/// keys <c>en</c> and <c>de</c> is the text in both languages and is replaced by the one of the document's language when it is created.
/// </summary>
internal static class SheetsTemplates
{
    /// <summary>The templates of the start page, the blank one first (names and descriptions are keys of the catalog).</summary>
    public static IReadOnlyList<KindTemplate> All { get; } =
    [
        new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-sheets"),
        new("budget", "Sheets.Template.Budget", "Sheets.Template.Budget.Text", "dollar"),
        new("invoice", "Sheets.Template.Invoice", "Sheets.Template.Invoice.Text", "file-text"),
        new("tasks", "Sheets.Template.Tasks", "Sheets.Template.Tasks.Text", "list-checks"),
        new("schedule", "Sheets.Template.Schedule", "Sheets.Template.Schedule.Text", "calendar"),
        new("loan", "Sheets.Template.Loan", "Sheets.Template.Loan.Text", "briefcase"),
        new("sales", "Sheets.Template.Sales", "Sheets.Template.Sales.Text", "bar-chart"),
    ];

    /// <summary>The content of a template in a language (en or de; unknown ids give the blank workbook).</summary>
    public static string Create(string templateId, string language)
    {
        var code = Languages.Normalize(language);
        if (templateId != "blank" && All.Any(t => t.Id == templateId) && Read(templateId) is { } template)
        {
            return Localize(template, code)!.ToJsonString();
        }

        var name = Languages.In(code, () => Text.T("Sheets.SheetName", 1));
        return new JsonObject
        {
            ["version"] = 1,
            ["active"] = 0,
            ["styles"] = new JsonArray(new JsonObject()),
            ["sheets"] = new JsonArray(new JsonObject { ["id"] = "s1", ["name"] = name, ["cells"] = new JsonObject() }),
        }.ToJsonString();
    }

    private static JsonNode? Read(string id)
    {
        var assembly = typeof(SheetsTemplates).Assembly;
        var resource = assembly.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith(".Resources.Templates.sheets." + id + ".json", StringComparison.Ordinal));
        if (resource is null)
        {
            return null;
        }

        using var stream = assembly.GetManifestResourceStream(resource)!;
        return JsonNode.Parse(stream, documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
    }

    /// <summary>A copy of <paramref name="node"/> with every <c>{ "en": …, "de": … }</c> replaced by the value of <paramref name="language"/>.</summary>
    public static JsonNode? Localize(JsonNode? node, string language)
    {
        switch (node)
        {
            case JsonObject o when o.Count == 2 && o.ContainsKey("en") && o.ContainsKey("de"):
                return Localize(o[language], language);
            case JsonObject o:
                var copy = new JsonObject();
                foreach (var (key, value) in o)
                {
                    copy[key] = Localize(value, language);
                }

                return copy;
            case JsonArray a:
                return new JsonArray([.. a.Select(item => Localize(item, language))]);
            default:
                return node?.DeepClone();
        }
    }
}
