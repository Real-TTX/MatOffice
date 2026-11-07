using MatOffice.Data;
using MatOffice.Services.Formats;

namespace MatOffice.Pages;

/// <summary>
/// What the frame of an editor needs from the server (<c>Shared/_LayoutEditor</c>): which app, where the file is read and saved (<see cref="Api"/>), the way back.
/// The editor itself loads the content with a request (the same for a document, a version, a link and a file of a host).
/// </summary>
public sealed record EditorPage(DocumentKind Kind, string Title, string Api, string Mode, string BackUrl, string? ShareUrl = null, string? GuestName = null)
{
    public string Slug => Kinds.SlugOf(Kind);

    public string KindName => Kind.ToString();

    public string Colour => Kinds.Of(Kind).Colour;
}
