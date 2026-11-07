using MatOffice.Data;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>MatOffice Sheets: spreadsheets (like Excel). STUB of the skeleton: the content is taken as it is (one JSON object), no foreign formats yet.</summary>
public sealed class SheetsHandler : IKindHandler
{
    private const string Blank = @"{ ""sheets"": [ { ""id"": ""s1"", ""name"": ""Sheet1"", ""cells"": {} } ], ""active"": 0 }";

    public DocumentKind Kind => DocumentKind.Sheets;

    public string Slug => "sheets";

    public string Colour => "#107c41";

    public string Icon => "kind-sheets";

    public string NativeExtension => ".matsheet";

    public IReadOnlyList<FileFormat> Formats { get; } = [];

    public IReadOnlyList<KindTemplate> Templates { get; } = [new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-sheets")];

    public string Create(string templateId, string language) => Blank;

    public ContentCheck Normalize(string json)
    {
        try
        {
            return ContentCheck.Valid(Kinds.ParseObject(json).ToJsonString());
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (System.Text.Json.JsonException)
        {
            return ContentCheck.Invalid(Localization.Text.T("Documents.Invalid"));
        }
    }

    public string Import(Stream input, string extension) => throw new FormatException(Localization.Text.T("Documents.Import.NotSupported"));

    public byte[] Export(string content, string extension, string name) => throw new FormatException(Localization.Text.T("Documents.Export.NotSupported"));

    public string PlainText(string content, int max) => string.Empty;
}
