using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MatOffice.Services.Formats;
using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>Runs when <c>MATOFFICE_DOCX_SAMPLES</c> names a folder of real Word files (they belong to others and never go into the repository).</summary>
public sealed class DocxSamplesFactAttribute : FactAttribute
{
    public DocxSamplesFactAttribute()
    {
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("MATOFFICE_DOCX_SAMPLES")))
        {
            Skip = "MATOFFICE_DOCX_SAMPLES is not set (a folder with .docx files).";
        }
    }
}

/// <summary>Word files as Word writes them (German Word: the style ids are German, the names English) and real Word files.</summary>
public sealed class DocsWordTests
{
    private static readonly DocsHandler Handler = new();

    private static readonly JsonSerializerOptions Plain = new() { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private static readonly byte[] Dot = Convert.FromBase64String(DocsFormatTests.Png[(DocsFormatTests.Png.IndexOf(',', StringComparison.Ordinal) + 1)..]);

    private const string Namespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        + "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
        + "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" mc:Ignorable=\"w14 w15\"";

    /// <summary>The style sheet of a German Word: ids like <c>Standard</c> and <c>berschrift1</c>, the built-in names in English.</summary>
    private const string GermanStyles = """
        <w:docDefaults>
          <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/>
            <w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-DE" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
          <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
        </w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Standard"><w:name w:val="Normal"/><w:qFormat/></w:style>
        <w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/><w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:link w:val="berschrift1Zchn"/>
          <w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>
          <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="berschrift2"><w:name w:val="heading 2"/><w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/>
          <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr>
          <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Titel"><w:name w:val="Title"/><w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:uiPriority w:val="10"/><w:qFormat/>
          <w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>
          <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/>
          <w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Zitat"><w:name w:val="Quote"/><w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:uiPriority w:val="29"/><w:qFormat/>
          <w:pPr><w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Listenabsatz"><w:name w:val="List Paragraph"/><w:basedOn w:val="Standard"/><w:uiPriority w:val="34"/><w:qFormat/>
          <w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
        <w:style w:type="paragraph" w:customStyle="1" w:styleId="Hinweis"><w:name w:val="Hinweis"/><w:basedOn w:val="Standard"/><w:qFormat/>
          <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="E2EFD9"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="375623"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Kopfzeile"><w:name w:val="header"/><w:basedOn w:val="Standard"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
          <w:pPr><w:tabs><w:tab w:val="center" w:pos="4536"/><w:tab w:val="right" w:pos="9072"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Fuzeile"><w:name w:val="footer"/><w:basedOn w:val="Standard"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
          <w:pPr><w:tabs><w:tab w:val="center" w:pos="4536"/><w:tab w:val="right" w:pos="9072"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Funotentext"><w:name w:val="footnote text"/><w:basedOn w:val="Standard"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
          <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Kommentartext"><w:name w:val="annotation text"/><w:basedOn w:val="Standard"/><w:uiPriority w:val="99"/>
          <w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
        <w:style w:type="character" w:default="1" w:styleId="Absatz-Standardschriftart"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
        <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="Absatz-Standardschriftart"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
          <w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Hervorhebung"><w:name w:val="Emphasis"/><w:basedOn w:val="Absatz-Standardschriftart"/><w:uiPriority w:val="20"/><w:qFormat/><w:rPr><w:i/><w:iCs/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Funotenzeichen"><w:name w:val="footnote reference"/><w:basedOn w:val="Absatz-Standardschriftart"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Kommentarzeichen"><w:name w:val="annotation reference"/><w:basedOn w:val="Absatz-Standardschriftart"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>
        <w:style w:type="table" w:default="1" w:styleId="NormaleTabelle"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
          <w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
        <w:style w:type="table" w:styleId="Tabellenraster"><w:name w:val="Table Grid"/><w:basedOn w:val="NormaleTabelle"/><w:uiPriority w:val="39"/>
          <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
          <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
        <w:style w:type="numbering" w:default="1" w:styleId="KeineListe"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
        """;

    private const string Numbering = """
        <w:abstractNum w:abstractNumId="0"><w:nsid w:val="1A2B3C4D"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="0407000F"/>
          <w:lvl w:ilvl="0" w:tplc="04070001"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
          <w:lvl w:ilvl="1" w:tplc="04070003"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/></w:rPr></w:lvl>
          <w:lvl w:ilvl="2" w:tplc="04070005"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="1"><w:nsid w:val="2B3C4D5E"/><w:multiLevelType w:val="hybridMultilevel"/>
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
          <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="2"><w:nsid w:val="3C4D5E6F"/><w:multiLevelType w:val="multilevel"/>
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl>
          <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="792" w:hanging="432"/></w:pPr></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
        <w:num w:numId="3"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
        <w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>
        """;

    private static string Theme(string major, string minor)
        => "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\"><a:themeElements><a:fontScheme name=\"Office\">"
           + $"<a:majorFont><a:latin typeface=\"{major}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
           + $"<a:minorFont><a:latin typeface=\"{minor}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>";

    private static string Picture(string relationshipId, int id, string name, long cx, long cy)
        => "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr>"
           + $"<pic:cNvPr id=\"{id}\" name=\"{name}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{relationshipId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
           + $"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>";

    private static string FieldRuns(string instruction, string result)
        => "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
           + $"<w:r><w:instrText xml:space=\"preserve\"> {instruction} </w:instrText></w:r>"
           + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
           + $"<w:r><w:t>{result}</w:t></w:r>"
           + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static void Feed(OpenXmlPart part, string xml)
    {
        using var data = new MemoryStream(Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + xml));
        part.FeedData(data);
    }

    /// <summary>A Word file made of the XML of its parts (as Word writes them); <paramref name="parts"/> adds parts with the relationship ids the XML uses.</summary>
    private static byte[] WordFile(string body, string styles = GermanStyles, string? numbering = Numbering, string? theme = null, Action<MainDocumentPart>? parts = null)
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            Feed(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {Namespaces}>{styles}</w:styles>");
            if (numbering is not null)
            {
                Feed(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {Namespaces}>{numbering}</w:numbering>");
            }

            Feed(main.AddNewPart<ThemePart>(), theme ?? Theme("Calibri Light", "Calibri"));
            parts?.Invoke(main);
            Feed(main, $"<w:document {Namespaces}><w:body>{body}</w:body></w:document>");
        }

        return stream.ToArray();
    }

    private static JsonObject Read(byte[] docx) => DocsFormatTests.Parse(DocsFormatTests.Import(docx, ".docx"));

    [Fact]
    public void A_document_of_german_word_imports_with_its_styles_lists_tables_pictures_notes_and_comments()
    {
        var body = "<w:p><w:pPr><w:pStyle w:val=\"Titel\"/></w:pPr><w:r><w:t>Projektbericht</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"berschrift1\"/></w:pPr><w:bookmarkStart w:id=\"0\" w:name=\"_Toc1\"/><w:r><w:t>Einleitung</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>"
            + "<w:p><w:r><w:t xml:space=\"preserve\">Normal mit </w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>fett</w:t></w:r><w:r><w:t xml:space=\"preserve\">, </w:t></w:r>"
            + "<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>kursiv</w:t></w:r><w:r><w:t xml:space=\"preserve\">, </w:t></w:r><w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t>unterstrichen</w:t></w:r>"
            + "<w:r><w:t xml:space=\"preserve\">, </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>durchgestrichen</w:t></w:r><w:r><w:t xml:space=\"preserve\">, </w:t></w:r>"
            + "<w:r><w:rPr><w:rStyle w:val=\"Hervorhebung\"/></w:rPr><w:t>betont</w:t></w:r><w:r><w:t xml:space=\"preserve\">, </w:t></w:r>"
            + "<w:r><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\"/><w:sz w:val=\"20\"/></w:rPr><w:t>Code</w:t></w:r><w:r><w:t xml:space=\"preserve\"> und </w:t></w:r>"
            + "<w:r><w:rPr><w:color w:val=\"FF0000\"/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>rot</w:t></w:r><w:r><w:t xml:space=\"preserve\">. </w:t></w:r>"
            + "<w:hyperlink r:id=\"rIdLink\" w:history=\"1\"><w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>MatOffice</w:t></w:r></w:hyperlink><w:r><w:t xml:space=\"preserve\">, </w:t></w:r>"
            + "<w:hyperlink w:anchor=\"_Toc1\" w:history=\"1\"><w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>zur Einleitung</w:t></w:r></w:hyperlink>"
            + "<w:r><w:t xml:space=\"preserve\">, eine Fußnote</w:t></w:r><w:r><w:rPr><w:rStyle w:val=\"Funotenzeichen\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r>"
            + "<w:r><w:t xml:space=\"preserve\"> und </w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:commentRangeStart w:id=\"1\"/><w:r><w:t>kommentiert</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:commentRangeEnd w:id=\"1\"/>"
            + "<w:r><w:rPr><w:rStyle w:val=\"Kommentarzeichen\"/></w:rPr><w:commentReference w:id=\"0\"/></w:r><w:r><w:rPr><w:rStyle w:val=\"Kommentarzeichen\"/></w:rPr><w:commentReference w:id=\"1\"/></w:r>"
            + "<w:ins w:id=\"10\" w:author=\"A\" w:date=\"2026-11-07T10:00:00Z\"><w:r><w:t xml:space=\"preserve\"> eingefügt</w:t></w:r></w:ins>"
            + "<w:del w:id=\"11\" w:author=\"A\" w:date=\"2026-11-07T10:00:00Z\"><w:r><w:delText xml:space=\"preserve\"> gelöscht</w:delText></w:r></w:del>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:t>verborgen</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Hinweis\"/></w:pPr><w:r><w:t>Ein Hinweis.</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Zitat\"/></w:pPr><w:r><w:t>Ein Zitat.</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Punkt eins</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Unterpunkt</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"2\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Unter-Unterpunkt</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Punkt zwei</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Erstens</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Erstens a</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Zweitens</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Dazwischen</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Drittens</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"3\"/></w:numPr></w:pPr><w:r><w:t>Neu eins</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Listenabsatz\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"3\"/></w:numPr></w:pPr><w:r><w:t>Neu zwei</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Gliederung</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"4\"/></w:numPr></w:pPr><w:r><w:t>Kapitel</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"4\"/></w:numPr></w:pPr><w:r><w:t>Abschnitt</w:t></w:r></w:p>"
            + "<w:tbl><w:tblPr><w:tblStyle w:val=\"Tabellenraster\"/><w:tblW w:w=\"5000\" w:type=\"pct\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"3020\"/><w:gridCol w:w=\"3021\"/><w:gridCol w:w=\"3021\"/></w:tblGrid>"
            + "<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"3020\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Kopf A</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"3021\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Kopf B</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"3021\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Kopf C</w:t></w:r></w:p></w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3020\" w:type=\"dxa\"/><w:vMerge w:val=\"restart\"/><w:vAlign w:val=\"center\"/></w:tcPr><w:p><w:r><w:t>Zusammen</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"6042\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Breit</w:t></w:r></w:p></w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3020\" w:type=\"dxa\"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3021\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>b3</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"3021\" w:type=\"dxa\"/><w:tcBorders><w:bottom w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"00B050\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>c3</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"952500\" cy=\"476250\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
            + "<wp:docPr id=\"1\" name=\"Grafik 1\" descr=\"Ein Punkt\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>" + Picture("rIdImg", 1, "Grafik 1", 952500, 476250) + "</wp:inline></w:drawing></w:r></w:p>"
            + "<w:p><w:r><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
            + "<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>"
            + "<wp:extent cx=\"476250\" cy=\"476250\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapSquare wrapText=\"bothSides\"/><wp:docPr id=\"2\" name=\"Grafik 2\" title=\"Rechts\"/><wp:cNvGraphicFramePr/>"
            + Picture("rIdImg", 2, "Grafik 2", 476250, 476250) + "</wp:anchor></w:drawing></w:r><w:r><w:t>Text neben dem Bild.</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t xml:space=\"preserve\">Datum: </w:t></w:r><w:fldSimple w:instr=\" DATE \\@ &quot;dd.MM.yyyy&quot; \"><w:r><w:t>07.11.2026</w:t></w:r></w:fldSimple>"
            + "<w:r><w:t xml:space=\"preserve\">, Seite </w:t></w:r>" + FieldRuns("PAGE", "3") + "</w:p>"
            + "<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t><w:noBreakHyphen/><w:t>D</w:t><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/></w:r></w:p>"
            + "<w:sdt><w:sdtPr><w:alias w:val=\"Feld\"/><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Im Steuerelement</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "<w:p><w:r><w:t>Vor dem Textfeld</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" "
            + "relativeHeight=\"251660288\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH>"
            + "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx=\"1828800\" cy=\"457200\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapNone/>"
            + "<wp:docPr id=\"3\" name=\"Textfeld 3\"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr txBox=\"1\"/>"
            + "<wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1828800\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>"
            + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Im Textfeld</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"
            + "<mc:Fallback><w:pict><v:shape id=\"Textfeld 3\" style=\"width:144pt;height:36pt\"><v:textbox><w:txbxContent><w:p><w:r><w:t>Im Textfeld</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>"
            + "</mc:Fallback></mc:AlternateContent></w:r></w:p>"
            + "<w:p><w:r><w:t>Vor dem Umbruch</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r><w:r><w:t>Nach dem Umbruch</w:t></w:r></w:p>"
            + "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
            + "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Neue Seite</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/></w:pBdr></w:pPr></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1417\" w:right=\"1417\" w:bottom=\"1134\" w:left=\"1417\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>"
            + "<w:cols w:space=\"708\"/></w:sectPr></w:pPr><w:r><w:t>Ende Abschnitt 1</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"berschrift2\"/></w:pPr><w:r><w:t>Abschnitt 2</w:t></w:r></w:p>"
            + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/><w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/><w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>"
            + "<w:pgMar w:top=\"1417\" w:right=\"1417\" w:bottom=\"1134\" w:left=\"1417\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGrid w:linePitch=\"360\"/></w:sectPr>";
        var docx = WordFile(body, parts: main =>
        {
            main.AddHyperlinkRelationship(new Uri("https://example.com/matoffice"), true, "rIdLink");
            using (var image = new MemoryStream(Dot))
            {
                main.AddImagePart(ImagePartType.Png, "rIdImg").FeedData(image);
            }

            Feed(main.AddNewPart<HeaderPart>("rIdHeader"), $"<w:hdr {Namespaces}><w:p><w:pPr><w:pStyle w:val=\"Kopfzeile\"/><w:jc w:val=\"right\"/></w:pPr>"
                + "<w:r><w:t xml:space=\"preserve\">Seite </w:t></w:r>" + FieldRuns("PAGE   \\* MERGEFORMAT", "1") + "<w:r><w:t xml:space=\"preserve\"> von </w:t></w:r>"
                + FieldRuns("NUMPAGES  \\* Arabic  \\* MERGEFORMAT", "2") + "</w:p></w:hdr>");
            Feed(main.AddNewPart<FooterPart>("rIdFooter"), $"<w:ftr {Namespaces}><w:p><w:pPr><w:pStyle w:val=\"Fuzeile\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Vertraulich</w:t></w:r></w:p></w:ftr>");
            Feed(main.AddNewPart<FootnotesPart>("rIdFootnotes"), $"<w:footnotes {Namespaces}>"
                + "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:id=\"1\"><w:p><w:pPr><w:pStyle w:val=\"Funotentext\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Funotenzeichen\"/></w:rPr><w:footnoteRef/></w:r>"
                + "<w:r><w:t xml:space=\"preserve\"> Eine </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>wichtige</w:t></w:r><w:r><w:t xml:space=\"preserve\"> Fußnote.</w:t></w:r></w:p></w:footnote></w:footnotes>");
            Feed(main.AddNewPart<WordprocessingCommentsPart>("rIdComments"), $"<w:comments {Namespaces}>"
                + "<w:comment w:id=\"0\" w:author=\"Erika Mustermann\" w:date=\"2026-11-07T10:00:00Z\" w:initials=\"EM\"><w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\">"
                + "<w:pPr><w:pStyle w:val=\"Kommentartext\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Kommentarzeichen\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Bitte prüfen</w:t></w:r></w:p></w:comment>"
                + "<w:comment w:id=\"1\" w:author=\"Max Mustermann\" w:date=\"2026-11-07T11:30:00Z\" w:initials=\"MM\"><w:p w14:paraId=\"1A2B3C4E\" w14:textId=\"77777777\">"
                + "<w:pPr><w:pStyle w:val=\"Kommentartext\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Kommentarzeichen\"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Erledigt</w:t></w:r></w:p></w:comment></w:comments>");
            Feed(main.AddNewPart<WordprocessingCommentsExPart>("rIdCommentsEx"), "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">"
                + "<w15:commentEx w15:paraId=\"1A2B3C4D\" w15:done=\"1\"/><w15:commentEx w15:paraId=\"1A2B3C4E\" w15:paraIdParent=\"1A2B3C4D\" w15:done=\"0\"/></w15:commentsEx>");
            Feed(main.AddNewPart<DocumentSettingsPart>("rIdSettings"), $"<w:settings {Namespaces}><w:hideSpellingErrors/><w:defaultTabStop w:val=\"708\"/></w:settings>");
        });

        var content = Read(docx);
        var dot = DocsFormatTests.Png;
        var cell = " style=\"line-height: 1.2; margin-bottom: 0pt\"";
        DocsFormatTests.AssertSameHtml(
            "<p data-style=\"Title\">Projektbericht</p><h1 id=\"x_Toc1\">Einleitung</h1>"
            + "<p>Normal mit <strong>fett</strong>, <em>kursiv</em>, <u>unterstrichen</u>, <s>durchgestrichen</s>, <em>betont</em>, "
            + "<span style=\"font-family: 'Courier New', Liberation Mono, monospace; font-size: 10pt\">Code</span> und <span style=\"font-size: 14pt; color: #ff0000; background-color: #ffff00\">rot</span>. "
            + "<a href=\"https://example.com/matoffice\">MatOffice</a>, <a href=\"#x_Toc1\">zur Einleitung</a>, eine Fußnote<sup data-fn=\"n1\"></sup> und <span data-comment=\"c0\">kommentiert</span> eingefügt.</p>"
            + "<p style=\"margin-top: 6pt; margin-bottom: 6pt; background-color: #e2efd9\"><span style=\"color: #375623\"><strong>Ein Hinweis.</strong></span></p>"
            + "<blockquote>Ein Zitat.</blockquote>"
            + "<p data-list=\"bullet\">Punkt eins</p><p data-list=\"bullet\" data-level=\"1\">Unterpunkt</p><p data-list=\"bullet\" data-level=\"2\">Unter-Unterpunkt</p><p data-list=\"bullet\">Punkt zwei</p>"
            + "<p data-list=\"decimal\">Erstens</p><p data-list=\"decimal\" data-level=\"1\">Erstens a</p><p data-list=\"decimal\">Zweitens</p><p>Dazwischen</p>"
            + "<p data-list=\"decimal\" data-start=\"3\">Drittens</p><p data-list=\"decimal\" data-start=\"1\">Neu eins</p><p data-list=\"decimal\">Neu zwei</p>"
            + "<p>Gliederung</p><p data-list=\"outline\">Kapitel</p><p data-list=\"outline\" data-level=\"1\">Abschnitt</p>"
            + "<table data-table-style=\"grid\" style=\"width: 100%\"><colgroup><col style=\"width: 33.3%\"><col style=\"width: 33.3%\"><col style=\"width: 33.3%\"></colgroup>"
            + $"<thead><tr><th><p{cell}>Kopf A</p></th><th><p{cell}>Kopf B</p></th><th><p{cell}>Kopf C</p></th></tr></thead>"
            + $"<tbody><tr><td rowspan=\"2\" style=\"vertical-align: middle\"><p{cell}>Zusammen</p></td><td colspan=\"2\" style=\"background-color: #ffff00\"><p style=\"text-align: center; line-height: 1.2; margin-bottom: 0pt\">Breit</p></td></tr>"
            + $"<tr><td><p{cell}>b3</p></td><td style=\"border-bottom: 1.5pt double #00b050\"><p{cell}>c3</p></td></tr></tbody></table>"
            + $"<p style=\"text-align: center\"><img src=\"{dot}\" alt=\"Ein Punkt\" style=\"width: 100px; height: 50px\"></p>"
            + $"<p><img src=\"{dot}\" alt=\"Rechts\" style=\"width: 50px; height: 50px\" data-align=\"right\">Text neben dem Bild.</p>"
            + "<p>Datum: 07.11.2026, Seite 3</p>"
            + "<p>A\tB<br>C‑D✓</p>"
            + "<p>Im Steuerelement</p>"
            + "<p>Vor dem Textfeld</p><p>Im Textfeld</p>"
            + "<p>Vor dem Umbruch</p><div data-break=\"page\"></div><p>Nach dem Umbruch</p>"
            + "<div data-break=\"page\"></div>"
            + "<div data-break=\"page\"></div><p>Neue Seite</p>"
            + "<hr>"
            + "<p>Ende Abschnitt 1</p><div data-break=\"page\"></div><h2>Abschnitt 2</h2>",
            content["html"]!.GetValue<string>());

        Assert.Equal("{\"size\":\"A4\",\"width\":210,\"height\":297,\"orientation\":\"landscape\",\"margins\":{\"top\":25,\"right\":25,\"bottom\":20,\"left\":25},\"header\":12.5,\"footer\":12.5,\"columns\":1,\"differentFirst\":false}", content["page"]!.ToJsonString());
        Assert.Equal("<p style=\"text-align: right; line-height: 1.2; margin-bottom: 0pt\">Seite <span data-field=\"page\"></span> von <span data-field=\"pages\"></span></p>", content["header"]!.GetValue<string>());
        Assert.Equal("<p style=\"text-align: center; line-height: 1.2; margin-bottom: 0pt\">Vertraulich</p>", content["footer"]!.GetValue<string>());
        Assert.Equal("[{\"id\":\"n1\",\"html\":\"<p>Eine <em>wichtige</em> Fußnote.</p>\"}]", content["notes"]!.ToJsonString(Plain));
        Assert.Equal(
            "[{\"id\":\"c0\",\"author\":\"Erika Mustermann\",\"date\":\"2026-11-07T10:00:00Z\",\"text\":\"Bitte prüfen\",\"done\":true,"
            + "\"replies\":[{\"id\":\"c1\",\"author\":\"Max Mustermann\",\"date\":\"2026-11-07T11:30:00Z\",\"text\":\"Erledigt\"}]}]",
            content["comments"]!.ToJsonString(Plain));
        Assert.Equal("{\"language\":\"de-DE\",\"spellcheck\":false}", content["settings"]!.ToJsonString());
    }

    [Fact]
    public void Theme_fonts_and_document_defaults_become_direct_formatting_where_they_differ_from_docs()
    {
        // Word 365: Aptos 12 pt, line spacing 1.15 – the text keeps its look.
        const string styles = """
            <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="24"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="360" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr>
            <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="0F4761"/><w:sz w:val="40"/></w:rPr></w:style>
            """;
        var aptos = Read(WordFile("<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p><w:p><w:r><w:t>Body text</w:t></w:r></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>",
            styles, numbering: null, theme: Theme("Aptos Display", "Aptos")));
        Assert.Equal(
            "<h1 style=\"line-height: 1.39; margin-top: 18pt; margin-bottom: 4pt\"><span style=\"font-family: 'Aptos Display', Calibri, sans-serif; font-size: 20pt; color: #0f4761\">Heading</span></h1>"
            + "<p style=\"line-height: 1.39\"><span style=\"font-family: Aptos, Calibri, sans-serif; font-size: 12pt\">Body text</span></p>",
            aptos["html"]!.GetValue<string>());
        Assert.Equal("Letter", aptos["page"]!["size"]!.GetValue<string>());
        Assert.Equal(25.4, aptos["page"]!["margins"]!["top"]!.GetValue<double>());
        Assert.Equal("en-GB", aptos["settings"]!["language"]!.GetValue<string>());

        // An old document: Times New Roman 12 pt by name, no theme.
        const string times = """
            <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """;
        Assert.Equal(
            "<p style=\"line-height: 1.2; margin-bottom: 0pt\"><span style=\"font-family: 'Times New Roman', Liberation Serif, serif; font-size: 12pt\">Old text</span></p>",
            Read(WordFile("<w:p><w:r><w:t>Old text</w:t></w:r></w:p>", times, numbering: null))["html"]!.GetValue<string>());
    }

    [Fact]
    public void Broken_and_odd_word_files_still_import()
    {
        // Style loops, missing numbering, unknown styles, merges without a start, spans wider than the grid, a picture of an unknown type,
        // a link without a relationship, fields without an end, empty tables, deeply nested tables.
        const string styles = """
            <w:style w:type="paragraph" w:styleId="A"><w:name w:val="Loop A"/><w:basedOn w:val="B"/></w:style>
            <w:style w:type="paragraph" w:styleId="B"><w:name w:val="Loop B"/><w:basedOn w:val="A"/><w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr></w:style>
            """;
        var nested = string.Concat(Enumerable.Repeat("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>deep</w:t></w:r></w:p>", 60)) + string.Concat(Enumerable.Repeat("</w:tc></w:tr></w:tbl><w:p/>", 60));
        var body = "<w:p><w:pPr><w:pStyle w:val=\"A\"/></w:pPr><w:r><w:t>loop</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Missing\"/><w:numPr><w:ilvl w:val=\"12\"/><w:numId w:val=\"5\"/></w:numPr></w:pPr><w:r><w:t>no numbering</w:t></w:r></w:p>"
            + "<w:tbl><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge/><w:gridSpan w:val=\"9\"/></w:tcPr><w:p><w:r><w:t>odd cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "<w:tbl><w:tr></w:tr></w:tbl>"
            + "<w:p><w:hyperlink r:id=\"rIdNothing\"><w:r><w:t>dead link</w:t></w:r></w:hyperlink><w:r><w:drawing><wp:inline><wp:extent cx=\"10\" cy=\"10\"/><wp:docPr id=\"1\" name=\"x\"/>"
            + Picture("rIdEmf", 1, "x", 10, 10) + "</wp:inline></w:drawing></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:t>never ends</w:t></w:r></w:p>"
            + nested;
        var docx = WordFile(body, styles, numbering: null, parts: main =>
        {
            using var emf = new MemoryStream([1, 2, 3, 4]);
            main.AddImagePart(ImagePartType.Emf, "rIdEmf").FeedData(emf);
        });
        var content = Read(docx);
        var html = content["html"]!.GetValue<string>();
        var text = Handler.PlainText(content.ToJsonString(), 100_000).Split('\n');
        Assert.Equal(["loop", "no numbering", "odd cell", "dead link"], text.Take(4));
        Assert.StartsWith("deep deep deep", text[4], StringComparison.Ordinal);
        Assert.DoesNotContain("data-list", html, StringComparison.Ordinal);
        Assert.Contains("<td colspan=\"9\">", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<img", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<a ", html, StringComparison.Ordinal);
        Assert.True(Handler.Normalize(content.ToJsonString()).Ok);
    }

    /// <summary>Word's "Grid Table 4 - Accent 1": the header row, the bands and the bold first column come from the conditional formatting of the style.</summary>
    private const string GridTable4Accent1 = """
        <w:style w:type="table" w:styleId="Gitternetztabelle4Akzent1"><w:name w:val="Grid Table 4 Accent 1"/><w:basedOn w:val="NormaleTabelle"/><w:uiPriority w:val="49"/>
          <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
          <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/><w:left w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/><w:right w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/>
          </w:tblBorders></w:tblPr>
          <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4" w:themeColor="accent1"/><w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4" w:themeColor="accent1"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4" w:themeColor="accent1"/><w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4" w:themeColor="accent1"/>
            <w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="4472C4" w:themeFill="accent1"/></w:tcPr></w:tblStylePr>
          <w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="4472C4" w:themeColor="accent1"/></w:tcBorders></w:tcPr></w:tblStylePr>
          <w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>
          <w:tblStylePr w:type="lastCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>
          <w:tblStylePr w:type="band1Vert"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr></w:tblStylePr>
          <w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr></w:tblStylePr>
        </w:style>
        """;

    [Theory]
    [InlineData("<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>")]
    [InlineData("<w:tblLook w:val=\"04A0\"/>")]
    public void A_table_in_a_built_in_table_style_gets_its_look(string tblLook)
    {
        string Cell(string text) => $"<w:tc><w:tcPr><w:tcW w:w=\"4531\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var body = $"<w:tbl><w:tblPr><w:tblStyle w:val=\"Gitternetztabelle4Akzent1\"/><w:tblW w:w=\"0\" w:type=\"auto\"/>{tblLook}</w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"4531\"/><w:gridCol w:w=\"4531\"/></w:tblGrid>"
            + $"<w:tr>{Cell("Name")}{Cell("Wert")}</w:tr><w:tr>{Cell("Alpha")}{Cell("1")}</w:tr><w:tr>{Cell("Beta")}{Cell("2")}</w:tr><w:tr>{Cell("Gamma")}{Cell("3")}</w:tr></w:tbl>";
        var p = "<p style=\"line-height: 1.2; margin-bottom: 0pt\">";
        Assert.Equal(
            "<table data-table-style=\"accent1\" style=\"width: 100%\"><colgroup><col style=\"width: 50%\"><col style=\"width: 50%\"></colgroup>"
            + $"<thead><tr><th>{p}Name</p></th><th>{p}Wert</p></th></tr></thead>"
            + $"<tbody><tr><td>{p}<strong>Alpha</strong></p></td><td>{p}1</p></td></tr><tr><td>{p}<strong>Beta</strong></p></td><td>{p}2</p></td></tr>"
            + $"<tr><td>{p}<strong>Gamma</strong></p></td><td>{p}3</p></td></tr></tbody></table>",
            Read(WordFile(body, GermanStyles + GridTable4Accent1))["html"]!.GetValue<string>());
    }

    [Fact]
    public void Only_bookmarks_a_link_points_to_become_ids_and_pictures_take_no_text_formatting()
    {
        var body = "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"target\"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>"
            + "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"bookmark=id.abc\"/><w:r><w:t>Unused</w:t></w:r><w:bookmarkEnd w:id=\"1\"/><w:bookmarkStart w:id=\"2\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"2\"/></w:p>"
            + "<w:p><w:hyperlink w:anchor=\"target\"><w:r><w:t>to target</w:t></w:r></w:hyperlink><w:r><w:t xml:space=\"preserve\"> and </w:t></w:r>"
            + "<w:hyperlink w:anchor=\"_heading=\"><w:r><w:t>to nothing</w:t></w:r></w:hyperlink></w:p>"
            + "<w:p><w:r><w:rPr><w:b/><w:color w:val=\"FF0000\"/><w:sz w:val=\"40\"/></w:rPr><w:drawing><wp:inline><wp:extent cx=\"95250\" cy=\"95250\"/><wp:docPr id=\"1\" name=\"x\"/>"
            + Picture("rIdImg", 1, "x", 95250, 95250) + "</wp:inline></w:drawing></w:r></w:p>";
        var content = Read(WordFile(body, parts: main =>
        {
            using var image = new MemoryStream(Dot);
            main.AddImagePart(ImagePartType.Png, "rIdImg").FeedData(image);
        }));
        Assert.Equal(
            $"<p id=\"target\">Target</p><p>Unused</p><p><a href=\"#target\">to target</a> and to nothing</p><p><img src=\"{DocsFormatTests.Png}\" alt=\"\" style=\"width: 10px; height: 10px\"></p>",
            content["html"]!.GetValue<string>());
    }

    [Fact]
    public void Headers_follow_the_sections_and_the_first_page()
    {
        var body = "<w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdH1\"/><w:headerReference w:type=\"first\" r:id=\"rIdF1\"/><w:titlePg/>"
            + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>two</w:t></w:r></w:p>"
            + "<w:sectPr><w:type w:val=\"continuous\"/><w:pgSz w:w=\"8391\" w:h=\"11906\"/><w:cols w:num=\"2\" w:space=\"708\"/><w:titlePg/></w:sectPr>";
        var content = Read(WordFile(body, parts: main =>
        {
            Feed(main.AddNewPart<HeaderPart>("rIdH1"), $"<w:hdr {Namespaces}><w:p><w:r><w:t>Kopf</w:t></w:r>" + FieldRuns("SECTIONPAGES", "1") + "</w:p></w:hdr>");
            Feed(main.AddNewPart<HeaderPart>("rIdF1"), $"<w:hdr {Namespaces}><w:p><w:r><w:t>Erste Seite</w:t></w:r><w:fldSimple w:instr=\" TITLE \"><w:r><w:t>Titel</w:t></w:r></w:fldSimple></w:p></w:hdr>");
        }));

        // A continuous section break is no page break; the last section takes the headers of the one before; A5 with two columns.
        Assert.Equal("<p>one</p><p>two</p>", content["html"]!.GetValue<string>());
        Assert.Equal("<p>Kopf<span data-field=\"pages\"></span></p>", content["header"]!.GetValue<string>());
        Assert.Equal("<p>Erste Seite<span data-field=\"title\"></span></p>", content["firstHeader"]!.GetValue<string>());
        Assert.Equal("A5", content["page"]!["size"]!.GetValue<string>());
        Assert.Equal(2, content["page"]!["columns"]!.GetValue<int>());
        Assert.True(content["page"]!["differentFirst"]!.GetValue<bool>());
    }

    // ------------------------------------------------------------------ real files

    /// <summary>
    /// Real Word files: every <c>.docx</c> of the folder in <c>MATOFFICE_DOCX_SAMPLES</c> must import into content that normalizes and export to a
    /// valid Word file again. The Docs HTML of each file is written to <c>imported/</c> in that folder, to be looked at.
    /// </summary>
    [DocxSamplesFact]
    public void Real_word_files_import()
    {
        var folder = Environment.GetEnvironmentVariable("MATOFFICE_DOCX_SAMPLES")!;
        var files = Directory.GetFiles(folder, "*.docx");
        Assert.NotEmpty(files);
        var output = Directory.CreateDirectory(Path.Combine(folder, "imported")).FullName;
        var report = new StringBuilder();
        foreach (var file in files)
        {
            var name = Path.GetFileName(file);
            var watch = Stopwatch.StartNew();
            ContentCheck result;
            using (var stream = File.OpenRead(file))
            {
                result = Kinds.Read(Handler, stream, ".docx");
            }

            Assert.True(result.Ok, $"{name}: {result.Error}");
            var read = watch.ElapsedMilliseconds;
            var document = DocsDocument.FromJson(result.Json!);
            var docx = Handler.Export(result.Json!, ".docx", name);
            using (var package = WordprocessingDocument.Open(new MemoryStream(docx), false))
            {
                var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package).Select(e => $"{e.Path?.XPath}: {e.Description}").Take(5).ToList();
                Assert.True(errors.Count == 0, $"{name} exported: " + string.Join("\n", errors));
            }

            var again = Kinds.Read(Handler, new MemoryStream(docx), ".docx");
            Assert.True(again.Ok, $"{name} again: {again.Error}");
            File.WriteAllText(Path.Combine(output, name + ".txt"), Review(result.Json!));
            File.WriteAllBytes(Path.Combine(output, name + ".html"), Handler.Export(result.Json!, ".html", name));
            var paragraphs = DocsHtml.Paragraphs(document.Body).ToList();
            report.Append(name).Append(": ").Append(read).Append(" ms, ")
                .Append(document.Page.Size).Append(document.Page.Landscape ? " landscape" : string.Empty).Append(", ")
                .Append(paragraphs.Count(p => p.HeadingLevel > 0)).Append(" headings, ")
                .Append(paragraphs.Count(p => p.List is not null)).Append(" list items, ")
                .Append(document.Body.OfType<DocTable>().Count()).Append(" tables, ")
                .Append(paragraphs.Sum(p => p.Inlines.OfType<DocImage>().Count())).Append(" pictures, ")
                .Append(document.Notes.Count).Append(" notes, ")
                .Append(document.Comments.Count).Append(" comments\n");
        }

        File.WriteAllText(Path.Combine(output, "report.txt"), report.ToString());
    }

    /// <summary>The content with one block per line, to be read by a person.</summary>
    private static string Review(string json)
    {
        var sb = new StringBuilder();
        foreach (var (key, value) in DocsFormatTests.Parse(json))
        {
            if (value is JsonValue v && v.TryGetValue<string>(out var html))
            {
                var shortened = System.Text.RegularExpressions.Regex.Replace(html, "base64,[A-Za-z0-9+/=]{40,}", m => "base64,…" + m.Length + " chars");
                sb.Append(key).Append(":\n").Append(System.Text.RegularExpressions.Regex.Replace(shortened, "(</(p|h[1-6]|blockquote|pre|tr|thead|tbody|table)>|<hr>|</div>)", "$1\n")).Append("\n\n");
            }
            else
            {
                sb.Append(key).Append(": ").Append(value?.ToJsonString(Plain)).Append("\n\n");
            }
        }

        return sb.ToString();
    }
}
