using MatOffice.Data;

namespace MatOffice.Services.Formats.Base;

/// <summary>MatOffice Base: desktop databases (like Access). STUB of the skeleton: the content is taken as it is (one JSON object), no foreign formats yet.</summary>
public sealed class BaseHandler : IKindHandler
{
    private const string Blank = @"{ ""tables"": [], ""relations"": [], ""queries"": [], ""forms"": [], ""reports"": [] }";

    public DocumentKind Kind => DocumentKind.Base;

    public string Slug => "base";

    public string Colour => "#a4373a";

    public string Icon => "kind-base";

    public string NativeExtension => ".matbase";

    public IReadOnlyList<FileFormat> Formats { get; } = [];

    public IReadOnlyList<KindTemplate> Templates { get; } = [new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-base")];

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
