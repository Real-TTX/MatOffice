using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MatOffice.Data;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// MatOffice Sheets: spreadsheets (like Excel).
/// <para>
/// The content (docs/EDITORS.md, "Sheets content") is one sparse JSON object: <c>{ version: 1, active, styles: [ {}, … ], names: [ { name, ref } ],
/// sheets: [ { id, name, color, hidden, cells: { "A1": { v, f, t, s, link, note } }, cols: { "A": { w, hidden, s } }, rows: { "1": { h, hidden, s } },
/// colW, rowH, merges, freeze, filter, cf, validations, charts, images, grid, headings, zoom, print, protected, sel } ] }</c>. A value <c>v</c> is a number,
/// a string or a boolean (a date is an Excel serial number with a date format); <c>f</c> is a formula without "=" (canonical English, "," between arguments)
/// whose last value is <c>v</c>; <c>t: "e"</c> marks an error value; <c>s</c> is an index into <c>styles</c> (0 = the default <c>{}</c>). The style of a cell
/// is its own <c>s</c>, else the one of its row, else the one of its column. Sizes are pixels at 100 %, font sizes points.
/// </para>
/// <para>
/// <see cref="SheetsNormalizer"/> keeps only these fields with their types and drops what does not fit (links other than http/https/mailto/#, pictures
/// other than png/jpeg/gif/webp data addresses, bad colours, fonts, addresses, style indexes …); formulas are text for the server, the editor calculates.
/// Files: .xlsx (<see cref="SheetsXlsxReader"/>, <see cref="SheetsXlsxWriter"/>), .csv and .tsv (<see cref="SheetsCsv"/>).
/// </para>
/// </summary>
public sealed class SheetsHandler : IKindHandler
{
    public const string XlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public DocumentKind Kind => DocumentKind.Sheets;

    public string Slug => "sheets";

    public string Colour => "#107c41";

    public string Icon => "kind-sheets";

    public string NativeExtension => ".matsheet";

    /// <summary>The formats, labelled in the language of the request.</summary>
    public IReadOnlyList<FileFormat> Formats =>
    [
        new(".xlsx", XlsxMime, Text.T("Sheets.Format.Xlsx"), true, true),
        new(".csv", "text/csv", Text.T("Sheets.Format.Csv"), true, true),
        new(".tsv", "text/tab-separated-values", Text.T("Sheets.Format.Tsv"), true, true),
    ];

    public IReadOnlyList<KindTemplate> Templates => SheetsTemplates.All;

    public string Create(string templateId, string language) => SheetsTemplates.Create(templateId, language);

    public ContentCheck Normalize(string json)
    {
        try
        {
            return ContentCheck.Valid(SheetsNormalizer.Normalize(Kinds.ParseObject(json)).ToJsonString());
        }
        catch (FormatException ex)
        {
            return ContentCheck.Invalid(ex.Message);
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or ArgumentException or System.FormatException)
        {
            return ContentCheck.Invalid(Text.T("Documents.Invalid"));
        }
    }

    public string Import(Stream input, string extension) => Kinds.NormalizeExtension(extension) switch
    {
        ".xlsx" => SheetsXlsxReader.Read(input),
        ".csv" => SheetsCsv.Read(input, tabs: false),
        ".tsv" => SheetsCsv.Read(input, tabs: true),
        _ => throw new FormatException(Text.T("Documents.Import.NotSupported")),
    };

    public byte[] Export(string content, string extension, string name)
    {
        var ext = Kinds.NormalizeExtension(extension);
        if (ext is not (".xlsx" or ".csv" or ".tsv"))
        {
            throw new FormatException(Text.T("Documents.Export.NotSupported"));
        }

        var check = Normalize(content);
        if (!check.Ok)
        {
            throw new FormatException(check.Error!);
        }

        var workbook = JsonNode.Parse(check.Json!)!.AsObject();
        return ext == ".xlsx" ? SheetsXlsxWriter.Write(workbook) : SheetsCsv.Write(workbook, tabs: ext == ".tsv");
    }

    /// <summary>Per sheet its name, then the values row by row (tabs between cells, a line per row); never throws.</summary>
    public string PlainText(string content, int max)
    {
        var sb = new StringBuilder();
        try
        {
            using var doc = JsonDocument.Parse(content, new JsonDocumentOptions { MaxDepth = 128 });
            if (!doc.RootElement.TryGetProperty("sheets", out var sheets) || sheets.ValueKind != JsonValueKind.Array)
            {
                return string.Empty;
            }

            foreach (var sheet in sheets.EnumerateArray())
            {
                if (sb.Length >= max)
                {
                    break;
                }

                if (sheet.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                if (sheet.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
                {
                    sb.Append(name.GetString()).Append('\n');
                }

                if (sheet.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Object)
                {
                    AppendCells(sb, cells, max);
                }

                sb.Append('\n');
            }
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or ArgumentException)
        {
            // what was read so far
        }

        return SheetsJson.Cut(sb.ToString().TrimEnd(), Math.Max(0, max));
    }

    private static void AppendCells(StringBuilder sb, JsonElement cells, int max)
    {
        var values = new List<(int Row, int Col, string Text)>();
        foreach (var cell in cells.EnumerateObject())
        {
            if (cell.Value.ValueKind == JsonValueKind.Object && A1.TryParseCell(cell.Name, out var r, out var c)
                && cell.Value.TryGetProperty("v", out var v) && Display(v) is { Length: > 0 } text)
            {
                values.Add((r, c, text));
            }
        }

        values.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Col.CompareTo(b.Col));
        var row = 0;
        foreach (var (r, _, text) in values)
        {
            if (sb.Length >= max)
            {
                return;
            }

            if (row != 0)
            {
                sb.Append(r == row ? '\t' : '\n');
            }

            row = r;
            sb.Append(text);
        }

        if (row != 0)
        {
            sb.Append('\n');
        }
    }

    private static string? Display(JsonElement v) => v.ValueKind switch
    {
        JsonValueKind.String => v.GetString(),
        JsonValueKind.Number => v.TryGetDouble(out var d) ? d.ToString("G15", CultureInfo.InvariantCulture) : null,
        JsonValueKind.True => "TRUE",
        JsonValueKind.False => "FALSE",
        _ => null,
    };
}
