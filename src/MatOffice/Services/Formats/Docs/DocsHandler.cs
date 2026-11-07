using MatOffice.Data;

namespace MatOffice.Services.Formats.Docs;

/// <summary>MatOffice Docs: text documents (like Word). STUB of the skeleton: the content is taken as it is (one JSON object), no foreign formats yet.</summary>
public sealed class DocsHandler : IKindHandler
{
    private const string Blank = @"{ ""page"": { ""size"": ""A4"", ""orientation"": ""portrait"" }, ""html"": ""<p></p>"" }";

    public DocumentKind Kind => DocumentKind.Docs;

    public string Slug => "docs";

    public string Colour => "#185abd";

    public string Icon => "kind-docs";

    public string NativeExtension => ".matdoc";

    public IReadOnlyList<FileFormat> Formats { get; } = [];

    public IReadOnlyList<KindTemplate> Templates { get; } = [new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-docs")];

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
