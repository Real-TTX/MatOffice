using MatOffice.Services.Formats;
using MatOffice.Services.Hosts;

namespace MatOffice.Tests;

/// <summary>What the registry of the four apps tells a connected app (MatDrive) about new files.</summary>
public sealed class KindsTests
{
    [Theory]
    [InlineData("docs", ".docx")]
    [InlineData("slides", ".pptx")]
    [InlineData("base", ".matbase")]
    public void A_new_file_has_the_format_people_expect(string slug, string extension)
        => Assert.Equal(extension, Kinds.BySlug(slug)!.NewFileExtension);

    [Fact]
    public void Every_app_writes_a_blank_file_in_the_format_of_its_new_files_and_opens_it_again()
    {
        foreach (var kind in Kinds.All)
        {
            var blank = HostSessionService.Blank(kind.Slug, kind.NewFileExtension, "Untitled", "en");
            Assert.True(blank.Ok, $"{kind.Slug}: {blank.Error}");

            // what a host session does with the file: read it back into the content of the app
            var read = Kinds.Read(kind, new MemoryStream(blank.Value!), kind.NewFileExtension);
            Assert.True(read.Ok, $"{kind.Slug} ({kind.NewFileExtension}): {read.Error}");
        }
    }
}
