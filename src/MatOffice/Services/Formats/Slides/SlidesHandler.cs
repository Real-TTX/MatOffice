using MatOffice.Data;

namespace MatOffice.Services.Formats.Slides;

/// <summary>MatOffice Slides: presentations (like PowerPoint). STUB of the skeleton: the content is taken as it is (one JSON object), no foreign formats yet.</summary>
public sealed class SlidesHandler : IKindHandler
{
    private const string Blank = @"{ ""size"": { ""w"": 1280, ""h"": 720 }, ""slides"": [ { ""id"": ""p1"", ""elements"": [] } ] }";

    public DocumentKind Kind => DocumentKind.Slides;

    public string Slug => "slides";

    public string Colour => "#c43e1c";

    public string Icon => "kind-slides";

    public string NativeExtension => ".matslides";

    public IReadOnlyList<FileFormat> Formats { get; } = [];

    public IReadOnlyList<KindTemplate> Templates { get; } = [new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-slides")];

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
