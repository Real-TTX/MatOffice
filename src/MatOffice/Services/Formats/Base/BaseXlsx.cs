using System.Globalization;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Base;

/// <summary>
/// Excel workbooks and databases: every sheet becomes a table (the first row gives the field names, the types are guessed from the cells, an AutoNumber
/// "ID" becomes the primary key unless a column ID with unique whole numbers is there), and every table becomes a sheet with a header row.
/// </summary>
public static class BaseXlsx
{
    private const int MaxColumns = 255;

    public static string Import(Stream input)
    {
        XLWorkbook book;
        try
        {
            book = new XLWorkbook(input);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new FormatException(Text.T("Documents.Import.Unreadable"));
        }

        using (book)
        {
            var tables = new JsonArray();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var total = 0;
            var t = 0;
            foreach (var sheet in book.Worksheets)
            {
                var used = sheet.RangeUsed();
                if (used is null)
                {
                    continue;
                }

                var firstRow = used.FirstRow().RowNumber();
                var lastRow = used.LastRow().RowNumber();
                var firstCol = used.FirstColumn().ColumnNumber();
                var lastCol = Math.Min(used.LastColumn().ColumnNumber(), firstCol + MaxColumns - 2);
                var rows = lastRow - firstRow;
                total += Math.Max(0, rows);
                if (total > BaseNormalizer.MaxRecords)
                {
                    throw new FormatException(Text.T("Base.Invalid.TooMany", BaseNormalizer.MaxRecords));
                }

                tables.Add(ReadSheet(sheet, "t" + ++t, UniqueName(names, sheet.Name, "Table"), firstRow, lastRow, firstCol, lastCol));
            }

            if (tables.Count == 0)
            {
                throw new FormatException(Text.T("Base.Import.NoSheets"));
            }

            return new JsonObject
            {
                ["version"] = 1,
                ["settings"] = new JsonObject { ["currency"] = Languages.IsGerman ? "EUR" : "USD" },
                ["tables"] = tables,
                ["relationships"] = new JsonArray(),
                ["queries"] = new JsonArray(),
                ["forms"] = new JsonArray(),
                ["reports"] = new JsonArray(),
                ["ui"] = new JsonObject(),
            }.ToJsonString();
        }
    }

    private static string UniqueName(HashSet<string> taken, string wish, string fallback)
    {
        var clean = new string((wish ?? string.Empty).Where(c => !".!`[]<>\"".Contains(c) && !char.IsControl(c)).ToArray()).Trim();
        if (clean.Length > 60)
        {
            clean = clean[..60];
        }

        if (clean.Length == 0)
        {
            clean = fallback;
        }

        var name = clean;
        for (var i = 1; !taken.Add(name); i++)
        {
            name = clean + i.ToString(CultureInfo.InvariantCulture);
        }

        return name;
    }

    private sealed class Column
    {
        public string Id { get; init; } = string.Empty;

        public string Name { get; init; } = string.Empty;

        public bool AllNumbers { get; set; } = true;

        public bool AllIntegers { get; set; } = true;

        public bool AllDates { get; set; } = true;

        public bool AllBools { get; set; } = true;

        public bool AnyCurrency { get; set; }

        public bool AnyValue { get; set; }

        public int Longest { get; set; }

        public HashSet<double> Seen { get; } = [];

        public bool Unique { get; set; } = true;

        public string Type => !AnyValue ? "text" : AllBools ? "yesno" : AllDates ? "datetime" : AllNumbers ? (AnyCurrency ? "currency" : "number") : Longest > 255 ? "memo" : "text";
    }

    private static JsonObject ReadSheet(IXLWorksheet sheet, string id, string name, int firstRow, int lastRow, int firstCol, int lastCol)
    {
        var columns = new List<Column>();
        var fieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var c = firstCol; c <= lastCol; c++)
        {
            var header = sheet.Cell(firstRow, c).GetFormattedString();
            columns.Add(new Column { Id = "f" + (columns.Count + 2).ToString(CultureInfo.InvariantCulture), Name = UniqueName(fieldNames, header, "Field" + (c - firstCol + 1).ToString(CultureInfo.InvariantCulture)) });
        }

        // the types: every cell of a column decides
        for (var r = firstRow + 1; r <= lastRow; r++)
        {
            for (var c = firstCol; c <= lastCol; c++)
            {
                var cell = sheet.Cell(r, c);
                if (cell.IsEmpty())
                {
                    continue;
                }

                var col = columns[c - firstCol];
                col.AnyValue = true;
                var v = cell.Value;
                if (v.IsBoolean)
                {
                    col.AllNumbers = col.AllDates = false;
                    col.Unique = false;
                    continue;
                }

                col.AllBools = false;
                if (v.IsDateTime)
                {
                    col.AllNumbers = false;
                    continue;
                }

                col.AllDates = false;
                if (v.IsNumber)
                {
                    var n = v.GetNumber();
                    if (n != Math.Floor(n) || Math.Abs(n) > int.MaxValue)
                    {
                        col.AllIntegers = false;
                    }

                    if (!col.Seen.Add(n))
                    {
                        col.Unique = false;
                    }

                    var format = cell.Style.NumberFormat.Format ?? string.Empty;
                    if (format.Contains('€') || format.Contains('$') || format.Contains('£') || cell.Style.NumberFormat.NumberFormatId is 5 or 6 or 7 or 8 or 44)
                    {
                        col.AnyCurrency = true;
                    }

                    continue;
                }

                col.AllNumbers = false;
                col.Unique = false;
                col.Longest = Math.Max(col.Longest, cell.GetFormattedString().Length);
            }
        }

        // the key: a column "ID" with unique whole numbers, else a new AutoNumber
        var key = columns.FirstOrDefault(c => string.Equals(c.Name, "ID", StringComparison.OrdinalIgnoreCase) && c.AnyValue && c.AllNumbers && c.AllIntegers && c.Unique);
        var fields = new JsonArray();
        string keyId;
        var counter = 0L;
        if (key is null)
        {
            keyId = "f1";
            fields.Add(Field("f1", UniqueName(fieldNames, "ID", "ID"), "autonumber"));
        }
        else
        {
            keyId = key.Id;
        }

        foreach (var col in columns)
        {
            var type = col.Type;
            var f = Field(col.Id, col.Name, type);
            if (type == "number")
            {
                f["numberSize"] = col.AllIntegers ? "long" : "double";
            }

            if (type == "text")
            {
                f["size"] = 255;
            }

            if (col == key)
            {
                f["required"] = true;
                f["indexed"] = "unique";
            }

            fields.Add(f);
        }

        var rows = new JsonArray();
        var rowId = 0L;
        for (var r = firstRow + 1; r <= lastRow; r++)
        {
            var values = new JsonObject();
            var any = false;
            for (var c = firstCol; c <= lastCol; c++)
            {
                var col = columns[c - firstCol];
                var cell = sheet.Cell(r, c);
                if (cell.IsEmpty())
                {
                    continue;
                }

                var node = CellValue(cell, col.Type);
                if (node is not null)
                {
                    values[col.Id] = node;
                    any = true;
                }
            }

            if (!any)
            {
                continue;
            }

            if (key is null)
            {
                values["f1"] = ++counter;
            }
            else if (values[keyId] is null)
            {
                continue;
            }

            rows.Add(new JsonObject { ["id"] = ++rowId, ["v"] = values });
        }

        if (key is null)
        {
            ((JsonObject)fields[0]!)["next"] = counter + 1;
        }

        return new JsonObject
        {
            ["id"] = id,
            ["name"] = name,
            ["description"] = string.Empty,
            ["fields"] = fields,
            ["primaryKey"] = new JsonArray(keyId),
            ["rows"] = rows,
            ["nextId"] = rowId + 1,
            ["datasheet"] = new JsonObject(),
        };
    }

    private static JsonObject Field(string id, string name, string type)
        => new()
        {
            ["id"] = id,
            ["name"] = name,
            ["type"] = type,
            ["format"] = type == "currency" ? "Currency" : type == "yesno" ? "Yes/No" : string.Empty,
            ["default"] = type == "yesno" ? "No" : string.Empty,
            ["required"] = false,
            ["indexed"] = "no",
        };

    private static JsonNode? CellValue(IXLCell cell, string type)
    {
        var v = cell.Value;
        switch (type)
        {
            case "yesno":
                return v.IsBoolean && v.GetBoolean();
            case "datetime":
                if (!v.IsDateTime)
                {
                    return null;
                }

                var d = v.GetDateTime();
                return d.TimeOfDay == TimeSpan.Zero ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : d.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            case "number":
            case "currency":
                return v.IsNumber ? v.GetNumber() : null;
            default:
            {
                var s = BaseNormalizer.CleanText(cell.GetFormattedString(), type == "memo" ? BaseNormalizer.MaxMemo : 255, type == "memo");
                return s.Length == 0 ? null : s;
            }
        }
    }

    /// <summary>Every table as a sheet: a bold header row, typed cells (numbers, currency, dates, Yes/No), the header frozen and an AutoFilter.</summary>
    public static byte[] Export(string content)
    {
        var root = Kinds.ParseObject(content);
        var currency = (string?)root["settings"]?["currency"] ?? "EUR";
        var currencyFormat = currency switch
        {
            "EUR" => "#,##0.00 \"€\"",
            "USD" => "\"$\"#,##0.00",
            "GBP" => "\"£\"#,##0.00",
            "CHF" => "\"CHF\" #,##0.00",
            _ => "#,##0.00",
        };
        using var book = new XLWorkbook();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in root["tables"] as JsonArray ?? [])
        {
            var name = SheetName(used, (string?)t?["name"] ?? "Table");
            var sheet = book.Worksheets.Add(name);
            var fields = (t?["fields"] as JsonArray ?? []).OfType<JsonObject>().ToList();
            for (var c = 0; c < fields.Count; c++)
            {
                var cell = sheet.Cell(1, c + 1);
                cell.Value = (string?)fields[c]["name"] ?? string.Empty;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F2DCDB");
            }

            var widths = fields.Select(f => Math.Max(8, ((string?)f["name"] ?? string.Empty).Length + 2)).ToArray();
            var r = 1;
            foreach (var row in t?["rows"] as JsonArray ?? [])
            {
                r++;
                var values = row?["v"] as JsonObject;
                for (var c = 0; c < fields.Count; c++)
                {
                    var f = fields[c];
                    var v = values?[(string)f["id"]!];
                    if (v is null)
                    {
                        continue;
                    }

                    var cell = sheet.Cell(r, c + 1);
                    var type = (string?)f["type"];
                    switch (type)
                    {
                        case "yesno":
                            cell.Value = (bool)v;
                            break;
                        case "datetime":
                        {
                            var s = (string)v!;
                            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                            {
                                cell.Value = d;
                                cell.Style.DateFormat.Format = s.Contains('T') ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd";
                            }

                            break;
                        }

                        case "number":
                        case "currency":
                        case "autonumber":
                            cell.Value = (double)v;
                            if (type == "currency")
                            {
                                cell.Style.NumberFormat.Format = currencyFormat;
                            }

                            break;
                        default:
                        {
                            var s = (string)v!;
                            cell.Value = s;
                            if (r < 300)
                            {
                                widths[c] = Math.Max(widths[c], Math.Min(60, s.Length + 2));
                            }

                            break;
                        }
                    }
                }
            }

            for (var c = 0; c < fields.Count; c++)
            {
                sheet.Column(c + 1).Width = (string?)fields[c]["type"] is "datetime" or "currency" ? Math.Max(widths[c], 14) : widths[c];
            }

            if (fields.Count > 0)
            {
                sheet.SheetView.FreezeRows(1);
                sheet.Range(1, 1, Math.Max(1, r), fields.Count).SetAutoFilter();
            }
        }

        if (book.Worksheets.Count == 0)
        {
            book.Worksheets.Add(Text.T("Base.Export.EmptySheet"));
        }

        using var output = new MemoryStream();
        book.SaveAs(output);
        return output.ToArray();
    }

    private static string SheetName(HashSet<string> used, string wish)
    {
        var clean = new string(wish.Where(c => !"[]:*?/\\".Contains(c)).ToArray()).Trim('\'').Trim();
        if (clean.Length > 31)
        {
            clean = clean[..31];
        }

        if (clean.Length == 0)
        {
            clean = "Table";
        }

        var name = clean;
        for (var i = 2; !used.Add(name); i++)
        {
            var suffix = " (" + i.ToString(CultureInfo.InvariantCulture) + ")";
            name = (clean.Length + suffix.Length > 31 ? clean[..(31 - suffix.Length)] : clean) + suffix;
        }

        return name;
    }
}
