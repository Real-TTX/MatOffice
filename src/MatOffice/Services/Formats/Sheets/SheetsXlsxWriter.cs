using System.Globalization;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Parser;
using DocumentFormat.OpenXml.Packaging;
using MatOffice.Localization;
using static MatOffice.Services.Formats.Sheets.SheetsJson;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// Writes (normalized) content as an .xlsx file with ClosedXML: values, formulas, styles, sizes, merges, frozen panes, tab colours, hidden sheets, rows and
/// columns, notes, links, names, validations, conditional formats, the AutoFilter, pictures, view and print settings. Charts are not written. ClosedXML writes
/// formulas without values, so the last values of the editor are put into the file afterwards (OpenXML); Excel calculates everything again when it opens it.
/// A formula ClosedXML cannot parse is written as its last value.
/// </summary>
internal sealed class SheetsXlsxWriter
{
    private readonly XLWorkbook _wb = new();
    private readonly JsonArray _styles;
    private readonly IXLWorksheet _scratch;
    private readonly Dictionary<int, IXLStyle> _protos = [];

    /// <summary>Per sheet name: the cells whose formula was written, with their content (for the cached values).</summary>
    private readonly Dictionary<string, Dictionary<string, JsonObject>> _formulas = new(StringComparer.OrdinalIgnoreCase);

    private SheetsXlsxWriter(JsonObject content)
    {
        _styles = content["styles"] as JsonArray ?? [];
        _scratch = _wb.Worksheets.Add("~" + Guid.NewGuid().ToString("N")[..20]);
    }

    public static byte[] Write(JsonObject content)
    {
        try
        {
            var writer = new SheetsXlsxWriter(content);
            using (writer._wb)
            {
                return writer.Workbook(content);
            }
        }
        catch (FormatException)
        {
            throw;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or IOException or OpenXmlPackageException or NotSupportedException)
        {
            throw new FormatException(Text.T("Sheets.Export.Failed"));
        }
    }

    private byte[] Workbook(JsonObject content)
    {
        var sheets = content["sheets"]!.AsArray().OfType<JsonObject>().ToList();
        var written = new List<IXLWorksheet>();
        foreach (var sheet in sheets)
        {
            written.Add(Sheet(sheet, written.Count));
        }

        foreach (var name in (content["names"] as JsonArray ?? []).OfType<JsonObject>())
        {
            Try(() => _wb.DefinedNames.Add(Str(name["name"])!, Str(name["ref"])!));
        }

        _scratch.Delete();
        var active = Math.Clamp(Int(content["active"], 0, written.Count - 1) ?? 0, 0, written.Count - 1);
        for (var i = 0; i < written.Count; i++)
        {
            written[i].TabActive = i == active;
            written[i].TabSelected = i == active;
        }

        _wb.FullCalculationOnLoad = true;
        using var stream = new MemoryStream();
        _wb.SaveAs(stream);
        if (_formulas.Count > 0)
        {
            WriteCachedValues(stream);
        }

        return stream.ToArray();
    }

    // ---- a sheet ----------------------------------------------------------------------------------------------------------------------

    private IXLWorksheet Sheet(JsonObject sheet, int index)
    {
        var name = Str(sheet["name"]) ?? string.Empty;
        IXLWorksheet ws;
        try
        {
            ws = _wb.Worksheets.Add(name);
        }
        catch (ArgumentException)
        {
            name = Text.T("Sheets.SheetName", index + 1) + " " + Guid.NewGuid().ToString("N")[..6];
            ws = _wb.Worksheets.Add(name);
        }

        if (Colour(Str(sheet["color"])) is { } tab)
        {
            ws.TabColor = XLColor.FromHtml(tab);
        }

        if (IsTrue(sheet["hidden"]))
        {
            ws.Visibility = XLWorksheetVisibility.Hidden;
        }

        var colW = Num(sheet["colW"]) ?? SheetsLimits.DefaultColWidth;
        var rowH = Num(sheet["rowH"]) ?? SheetsLimits.DefaultRowHeight;
        ws.ColumnWidth = Width(colW);
        ws.RowHeight = rowH * 0.75;

        // columns before rows before cells: ClosedXML gives a new row the formats of the existing columns (and the other way round)
        var colStyles = new Dictionary<int, int>();
        var colPx = new Dictionary<int, double>();
        foreach (var (key, value) in sheet["cols"] as JsonObject ?? [])
        {
            if (A1.ColumnIndex(key) is not (> 0 and var c) || value is not JsonObject line)
            {
                continue;
            }

            var col = ws.Column(c);
            if (Num(line["w"]) is { } w)
            {
                col.Width = Width(w);
                colPx[c] = w;
            }

            if (Int(line["s"], 1, _styles.Count - 1) is { } s)
            {
                col.Style = Proto(s);
                colStyles[c] = s;
            }

            if (IsTrue(line["hidden"]))
            {
                col.Hide();
                colPx[c] = 0;
            }
        }

        var rowStyles = new Dictionary<int, int>();
        var rowPx = new SortedDictionary<int, double>();
        foreach (var (key, value) in sheet["rows"] as JsonObject ?? [])
        {
            if (!A1.TryParseRow(key, out var r) || value is not JsonObject line)
            {
                continue;
            }

            var row = ws.Row(r);
            if (Num(line["h"]) is { } h)
            {
                row.Height = h * 0.75;
                rowPx[r] = h;
            }

            if (Int(line["s"], 1, _styles.Count - 1) is { } s)
            {
                row.Style = Proto(s);
                rowStyles[r] = s;
            }

            if (IsTrue(line["hidden"]))
            {
                row.Hide();
                rowPx[r] = 0;
            }
        }

        var formulas = new Dictionary<string, JsonObject>(StringComparer.Ordinal);
        foreach (var (address, value) in sheet["cells"] as JsonObject ?? [])
        {
            if (value is JsonObject cell && A1.TryParseCell(address, out var r, out var c))
            {
                Cell(ws.Cell(r, c), cell, rowStyles.ContainsKey(r) || colStyles.ContainsKey(c)
                    ? rowStyles.GetValueOrDefault(r, colStyles.GetValueOrDefault(c)) : null, formulas);
            }
        }

        if (formulas.Count > 0)
        {
            _formulas[ws.Name] = formulas;
        }

        foreach (var merge in (sheet["merges"] as JsonArray ?? []).Select(Str))
        {
            if (merge is not null && A1.TryParseArea(merge, out var area) && !area.IsCell)
            {
                ws.Range(area.Row1, area.Col1, area.Row2, area.Col2).Merge(false);
            }
        }

        if (sheet["freeze"] is JsonObject freeze)
        {
            var fr = Int(freeze["rows"], 0, A1.MaxRows - 1) ?? 0;
            var fc = Int(freeze["cols"], 0, A1.MaxCols - 1) ?? 0;
            if (fr > 0 || fc > 0)
            {
                ws.SheetView.Freeze(fr, fc);
            }
        }

        Try(() => Filter(ws, sheet["filter"] as JsonObject));
        foreach (var rule in (sheet["cf"] as JsonArray ?? []).OfType<JsonObject>())
        {
            Try(() => Rule(ws, rule));
        }

        foreach (var validation in (sheet["validations"] as JsonArray ?? []).OfType<JsonObject>())
        {
            Try(() => Validation(ws, validation));
        }

        var geometry = new Geometry(colW, rowH, colPx, rowPx);
        foreach (var image in (sheet["images"] as JsonArray ?? []).OfType<JsonObject>())
        {
            Try(() => Picture(ws, image, geometry));
        }

        View(ws, sheet);
        return ws;
    }

    private void Cell(IXLCell cell, JsonObject c, int? inherited, Dictionary<string, JsonObject> formulas)
    {
        if (Int(c["s"], 0, _styles.Count - 1) is { } own)
        {
            cell.Style = Proto(own);
        }
        else if (inherited is { } style)
        {
            cell.Style = Proto(style);
        }

        var formula = Str(c["f"]);
        if (formula is not null && IsValid(formula, cell) && Try(() => cell.FormulaA1 = formula) is not null)
        {
            formulas[cell.Address.ToString()!] = c;
        }
        else
        {
            SetValue(cell, c);
        }

        if (Link(Str(c["link"])) is { } link)
        {
            Try(() => cell.SetHyperlink(link[0] == '#' ? new XLHyperlink(link[1..]) : new XLHyperlink(new Uri(link))));
        }

        if (Str(c["note"]) is { Length: > 0 } note)
        {
            Try(() => cell.CreateComment().AddText(note));
        }
    }

    private static void SetValue(IXLCell cell, JsonObject c)
    {
        var v = c["v"];
        if (Str(v) is { } text)
        {
            if (Str(c["t"]) == "e")
            {
                cell.Value = ErrorOf(text);
            }
            else
            {
                cell.Value = text;
            }
        }
        else if (Num(v) is { } number)
        {
            cell.Value = number;
        }
        else if (Bool(v) is { } flag)
        {
            cell.Value = flag;
        }
    }

    private static XLError ErrorOf(string code) => code switch
    {
        "#NULL!" => XLError.NullValue,
        "#DIV/0!" => XLError.DivisionByZero,
        "#NAME?" => XLError.NameNotRecognized,
        "#NUM!" => XLError.NumberInvalid,
        "#N/A" => XLError.NoValueAvailable,
        "#REF!" or "#CIRC!" => XLError.CellReference,
        _ => XLError.IncompatibleValue,
    };

    /// <summary>Whether Excel (ClosedXML's parser) can read the formula; a broken one would make Excel "repair" the whole file.</summary>
    private static bool IsValid(string formula, IXLCell cell)
    {
        try
        {
            FormulaConverter.ToR1C1(formula, cell.Address.RowNumber, cell.Address.ColumnNumber);
            return true;
        }
        catch (Exception ex) when (ex is ParsingException or ArgumentException or InvalidOperationException or NotSupportedException)
        {
            return false;
        }
    }

    // ---- styles ----------------------------------------------------------------------------------------------------------------------

    /// <summary>The ClosedXML style of a style index: built once on a scratch sheet, then copied to every cell, row or column that uses it.</summary>
    private IXLStyle Proto(int index)
    {
        if (_protos.TryGetValue(index, out var proto))
        {
            return proto;
        }

        var cell = _scratch.Cell(_protos.Count + 1, 1);
        if (index > 0 && index < _styles.Count && _styles[index] is JsonObject style)
        {
            Apply(cell.Style, style);
        }

        proto = cell.Style;
        _protos[index] = proto;
        return proto;
    }

    private static void Apply(IXLStyle target, JsonObject s)
    {
        ApplyFont(target, s);
        if (Str(s["font"]) is { } font)
        {
            target.Font.FontName = font;
        }

        if (Num(s["size"]) is { } size)
        {
            target.Font.FontSize = size;
        }

        target.Alignment.Horizontal = Str(s["ha"]) switch
        {
            "left" => XLAlignmentHorizontalValues.Left,
            "center" => XLAlignmentHorizontalValues.Center,
            "right" => XLAlignmentHorizontalValues.Right,
            "justify" => XLAlignmentHorizontalValues.Justify,
            _ => XLAlignmentHorizontalValues.General,
        };
        target.Alignment.Vertical = Str(s["va"]) switch
        {
            "top" => XLAlignmentVerticalValues.Top,
            "middle" => XLAlignmentVerticalValues.Center,
            _ => XLAlignmentVerticalValues.Bottom,
        };
        if (IsTrue(s["wrap"]))
        {
            target.Alignment.WrapText = true;
        }

        if (Int(s["indent"], 1, 15) is { } indent)
        {
            target.Alignment.Indent = indent;
        }

        if (Str(s["nf"]) is { Length: > 0 } nf)
        {
            target.NumberFormat.Format = nf;
        }

        Border(s["bt"], v => target.Border.TopBorder = v, c => target.Border.TopBorderColor = c);
        Border(s["bb"], v => target.Border.BottomBorder = v, c => target.Border.BottomBorderColor = c);
        Border(s["bl"], v => target.Border.LeftBorder = v, c => target.Border.LeftBorderColor = c);
        Border(s["br"], v => target.Border.RightBorder = v, c => target.Border.RightBorderColor = c);
        if (IsTrue(s["unlocked"]))
        {
            target.Protection.Locked = false;
        }
    }

    /// <summary>The part of a style that conditional formats have too: b, i, u, strike, color, fill.</summary>
    private static void ApplyFont(IXLStyle target, JsonObject s)
    {
        if (IsTrue(s["b"]))
        {
            target.Font.Bold = true;
        }

        if (IsTrue(s["i"]))
        {
            target.Font.Italic = true;
        }

        if (IsTrue(s["u"]))
        {
            target.Font.Underline = XLFontUnderlineValues.Single;
        }

        if (IsTrue(s["strike"]))
        {
            target.Font.Strikethrough = true;
        }

        if (Colour(Str(s["color"])) is { } color)
        {
            target.Font.FontColor = XLColor.FromHtml(color);
        }

        if (Colour(Str(s["fill"])) is { } fill)
        {
            target.Fill.PatternType = XLFillPatternValues.Solid;
            target.Fill.BackgroundColor = XLColor.FromHtml(fill);
        }
    }

    private static void Border(JsonNode? node, Action<XLBorderStyleValues> style, Action<XLColor> color)
    {
        if (SheetsNormalizer.Border(Str(node)) is not { } border)
        {
            return;
        }

        var parts = border.Split(' ');
        style(parts[0] switch
        {
            "medium" => XLBorderStyleValues.Medium,
            "thick" => XLBorderStyleValues.Thick,
            "dashed" => XLBorderStyleValues.Dashed,
            "dotted" => XLBorderStyleValues.Dotted,
            "double" => XLBorderStyleValues.Double,
            "hair" => XLBorderStyleValues.Hair,
            _ => XLBorderStyleValues.Thin,
        });
        color(XLColor.FromHtml(parts[1]));
    }

    // ---- filter, rules, validations, pictures, view --------------------------------------------------------------------------------

    private static void Filter(IXLWorksheet ws, JsonObject? filter)
    {
        if (filter is null || !A1.TryParseArea(Str(filter["range"]) ?? string.Empty, out var area))
        {
            return;
        }

        ws.Range(area.Row1, area.Col1, area.Row2, area.Col2).SetAutoFilter();
        foreach (var (key, value) in filter["cols"] as JsonObject ?? [])
        {
            var col = A1.ColumnIndex(key) - area.Col1 + 1;
            if (col < 1 || value is not JsonObject spec)
            {
                continue;
            }

            var column = ws.AutoFilter.Column(col);
            foreach (var item in (spec["values"] as JsonArray ?? []).Select(Str).OfType<string>())
            {
                column.AddFilter(item, false);
            }

            if (spec["cond"] is JsonObject cond && Str(cond["op"]) is { } op)
            {
                var connector = Condition(column, op, Str(cond["v"]) ?? string.Empty);
                if (connector is not null && Str(cond["op2"]) is { } op2)
                {
                    var next = IsTrue(cond["or"]) ? connector.Or : connector.And;
                    Condition(next, op2, Str(cond["v2"]) ?? string.Empty);
                }
            }
        }
    }

    private static IXLFilterConnector? Condition(IXLFilterColumn column, string op, string value)
    {
        XLCellValue v = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : value;
        return op switch
        {
            "eq" => column.EqualTo(v, false),
            "ne" => column.NotEqualTo(v, false),
            "gt" => column.GreaterThan(v, false),
            "ge" => column.EqualOrGreaterThan(v, false),
            "lt" => column.LessThan(v, false),
            "le" => column.EqualOrLessThan(v, false),
            "begins" => column.BeginsWith(value, false),
            "ends" => column.EndsWith(value, false),
            "contains" => column.Contains(value, false),
            "notContains" => column.NotContains(value, false),
            _ => null,
        };
    }

    private static void Condition(IXLCustomFilteredColumn column, string op, string value)
    {
        XLCellValue v = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : value;
        switch (op)
        {
            case "eq": column.EqualTo(v, false); break;
            case "ne": column.NotEqualTo(v, false); break;
            case "gt": column.GreaterThan(v, false); break;
            case "ge": column.EqualOrGreaterThan(v, false); break;
            case "lt": column.LessThan(v, false); break;
            case "le": column.EqualOrLessThan(v, false); break;
            case "begins": column.BeginsWith(value, false); break;
            case "ends": column.EndsWith(value, false); break;
            case "contains": column.Contains(value, false); break;
            case "notContains": column.NotContains(value, false); break;
        }
    }

    private static void Rule(IXLWorksheet ws, JsonObject r)
    {
        if (!A1.TryParseArea(Str(r["range"]) ?? string.Empty, out var area))
        {
            return;
        }

        var range = ws.Range(area.Row1, area.Col1, area.Row2, area.Col2);
        var cf = range.AddConditionalFormat();
        var v1 = Str(r["v1"]) ?? string.Empty;
        var v2 = Str(r["v2"]) ?? v1;
        IXLStyle? style = Str(r["type"]) switch
        {
            "cell" => Str(r["op"]) switch
            {
                "gt" => cf.WhenGreaterThan(v1),
                "ge" => cf.WhenEqualOrGreaterThan(v1),
                "lt" => cf.WhenLessThan(v1),
                "le" => cf.WhenEqualOrLessThan(v1),
                "eq" => cf.WhenEquals(v1),
                "ne" => cf.WhenNotEquals(v1),
                "between" => cf.WhenBetween(v1, v2),
                _ => cf.WhenNotBetween(v1, v2),
            },
            "text" => Str(r["op"]) switch
            {
                "notContains" => cf.WhenNotContains(v1),
                "begins" => cf.WhenStartsWith(v1),
                "ends" => cf.WhenEndsWith(v1),
                _ => cf.WhenContains(v1),
            },
            "dup" => cf.WhenIsDuplicate(),
            "unique" => cf.WhenIsUnique(),
            "top" => cf.WhenIsTop(Int(r["n"], 1, 1000) ?? 10, IsTrue(r["percent"]) ? XLTopBottomType.Percent : XLTopBottomType.Items),
            "bottom" => cf.WhenIsBottom(Int(r["n"], 1, 1000) ?? 10, IsTrue(r["percent"]) ? XLTopBottomType.Percent : XLTopBottomType.Items),
            "above" or "below" => cf.WhenIsTrue(A1.Cell(area.Row1, area.Col1) + (Str(r["type"]) == "above" ? ">" : "<")
                + $"AVERAGE(${A1.Column(area.Col1)}${area.Row1}:${A1.Column(area.Col2)}${area.Row2})"),
            "formula" => cf.WhenIsTrue(v1),
            _ => null,
        };
        if (style is not null)
        {
            if (r["style"] is JsonObject s)
            {
                ApplyFont(style, s);
            }

            return;
        }

        switch (Str(r["type"]))
        {
            case "bar":
                cf.DataBar(XLColor.FromHtml(Colour(Str(r["color"])) ?? "#638ec6")).LowestValue().HighestValue();
                break;
            case "scale":
                var colors = (r["colors"] as JsonArray ?? []).Select(c => Colour(Str(c))).OfType<string>().ToList();
                var min = cf.ColorScale().LowestValue(XLColor.FromHtml(colors[0]));
                if (colors.Count == 3)
                {
                    min.Midpoint(XLCFContentType.Percentile, 50, XLColor.FromHtml(colors[1])).HighestValue(XLColor.FromHtml(colors[2]));
                }
                else
                {
                    min.HighestValue(XLColor.FromHtml(colors[^1]));
                }

                break;
        }
    }

    private static void Validation(IXLWorksheet ws, JsonObject v)
    {
        if (!A1.TryParseArea(Str(v["range"]) ?? string.Empty, out var area))
        {
            return;
        }

        var dv = ws.Range(area.Row1, area.Col1, area.Row2, area.Col2).CreateDataValidation();
        var v1 = Str(v["v1"]) ?? string.Empty;
        var type = Str(v["type"]);
        if (type == "list")
        {
            dv.List(v1.StartsWith('=') ? v1[1..] : "\"" + v1.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"", Bool(v["dropdown"]) ?? true);
        }
        else if (type == "custom")
        {
            dv.Custom(v1.TrimStart('='));
        }
        else
        {
            dv.AllowedValues = type switch
            {
                "whole" => XLAllowedValues.WholeNumber,
                "decimal" => XLAllowedValues.Decimal,
                "date" => XLAllowedValues.Date,
                "time" => XLAllowedValues.Time,
                _ => XLAllowedValues.TextLength,
            };
            dv.Operator = Str(v["op"]) switch
            {
                "eq" => XLOperator.EqualTo,
                "ne" => XLOperator.NotEqualTo,
                "gt" => XLOperator.GreaterThan,
                "ge" => XLOperator.EqualOrGreaterThan,
                "lt" => XLOperator.LessThan,
                "le" => XLOperator.EqualOrLessThan,
                "notBetween" => XLOperator.NotBetween,
                _ => XLOperator.Between,
            };
            dv.MinValue = Operand(v1);
            if (dv.Operator is XLOperator.Between or XLOperator.NotBetween)
            {
                dv.MaxValue = Operand(Str(v["v2"]) ?? v1);
            }
        }

        dv.IgnoreBlanks = Bool(v["blank"]) ?? true;
        dv.ShowErrorMessage = true;
        dv.ErrorStyle = Str(v["errorStyle"]) switch
        {
            "warning" => XLErrorStyle.Warning,
            "info" => XLErrorStyle.Information,
            _ => XLErrorStyle.Stop,
        };
        dv.ErrorTitle = Str(v["errorTitle"]) ?? string.Empty;
        dv.ErrorMessage = Str(v["error"]) ?? string.Empty;
        if (Str(v["prompt"]) is { } prompt)
        {
            dv.ShowInputMessage = true;
            dv.InputMessage = prompt;
        }
    }

    /// <summary>A value of a validation as a formula: a number as it is, "=…" as the formula, any other text in quotes.</summary>
    private static string Operand(string value)
        => value.StartsWith('=') ? value[1..]
            : double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _) ? value
            : "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";

    private static void Picture(IXLWorksheet ws, JsonObject image, Geometry geometry)
    {
        if (Image(Str(image["src"])) is not { } src)
        {
            return;
        }

        var comma = src.IndexOf(',');
        var format = src[..comma] switch
        {
            "data:image/png;base64" => XLPictureFormat.Png,
            "data:image/jpeg;base64" => XLPictureFormat.Jpeg,
            "data:image/gif;base64" => XLPictureFormat.Gif,
            _ => XLPictureFormat.Webp,
        };
        using var stream = new MemoryStream(Convert.FromBase64String(src[(comma + 1)..]));
        var picture = ws.AddPicture(stream, format);
        var (col, dx) = geometry.Column(Num(image["x"]) ?? 0);
        var (row, dy) = geometry.Row(Num(image["y"]) ?? 0);
        picture.MoveTo(ws.Cell(row, col), (int)Math.Round(dx), (int)Math.Round(dy))
            .WithSize((int)Math.Round(Num(image["w"]) ?? picture.OriginalWidth), (int)Math.Round(Num(image["h"]) ?? picture.OriginalHeight));
    }

    private static void View(IXLWorksheet ws, JsonObject sheet)
    {
        if (Bool(sheet["grid"]) == false)
        {
            ws.ShowGridLines = false;
        }

        if (Bool(sheet["headings"]) == false)
        {
            ws.ShowRowColHeaders = false;
        }

        if (Range(sheet["zoom"], 0.1, 4) is { } zoom)
        {
            ws.SheetView.ZoomScale = (int)Math.Round(zoom * 100);
        }

        if (sheet["print"] is JsonObject print)
        {
            var setup = ws.PageSetup;
            if (IsTrue(print["landscape"]))
            {
                setup.PageOrientation = XLPageOrientation.Landscape;
            }

            switch (Str(print["fit"]))
            {
                case "width":
                    setup.FitToPages(1, 0);
                    break;
                case "page":
                    setup.FitToPages(1, 1);
                    break;
            }

            setup.ShowGridlines = IsTrue(print["gridlines"]);
            setup.ShowRowAndColumnHeadings = IsTrue(print["headings"]);
            if (Str(print["titles"]) is { } titles && SheetsNormalizer.RowSpan(titles) is { } span)
            {
                var parts = span.Split(':');
                setup.SetRowsToRepeatAtTop(int.Parse(parts[0], CultureInfo.InvariantCulture), int.Parse(parts[1], CultureInfo.InvariantCulture));
            }
        }

        if (IsTrue(sheet["protected"]))
        {
            ws.Protect();
        }

        if (A1.NormalizeArea(Str(sheet["sel"])) is { } sel && A1.TryParseArea(sel, out var area))
        {
            Try(() => ws.ActiveCell = ws.Cell(area.Row1, area.Col1));
        }
    }

    // ---- cached values of formulas ---------------------------------------------------------------------------------------------------

    /// <summary>Puts the last values of the editor into the formula cells ClosedXML wrote without one.</summary>
    private void WriteCachedValues(MemoryStream stream)
    {
        using var doc = SpreadsheetDocument.Open(stream, true);
        var workbook = doc.WorkbookPart!;
        foreach (var sheet in workbook.Workbook?.Sheets?.Elements<S.Sheet>() ?? [])
        {
            if (sheet.Name?.Value is not { } name || !_formulas.TryGetValue(name, out var cells) || workbook.GetPartById(sheet.Id!.Value!) is not WorksheetPart part)
            {
                continue;
            }

            if (part.Worksheet is not { } worksheet)
            {
                continue;
            }

            foreach (var cell in worksheet.Descendants<S.Cell>())
            {
                if (cell.CellFormula is null || cell.CellReference?.Value is not { } address || !cells.TryGetValue(address, out var content))
                {
                    continue;
                }

                var v = content["v"];
                if (Str(v) is { } text)
                {
                    var error = Str(content["t"]) == "e";
                    cell.DataType = error ? S.CellValues.Error : S.CellValues.String;
                    cell.CellValue = new S.CellValue(error ? (text == "#CIRC!" ? "#REF!" : text) : text);
                }
                else if (Num(v) is { } number)
                {
                    cell.DataType = null;
                    cell.CellValue = new S.CellValue(number.ToString("R", CultureInfo.InvariantCulture));
                }
                else if (Bool(v) is { } flag)
                {
                    cell.DataType = S.CellValues.Boolean;
                    cell.CellValue = new S.CellValue(flag ? "1" : "0");
                }
            }

            worksheet.Save();
        }
    }

    // ---- helpers ---------------------------------------------------------------------------------------------------------------------

    /// <summary>Pixels in Excel's column width (characters of the default font): px = width × 7 + 5.</summary>
    private static double Width(double px) => Math.Clamp((px - 5) / 7, 0, 255);

    private static object? Try(Action action)
    {
        try
        {
            action();
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException or System.FormatException or NotSupportedException
                                       or IOException or OverflowException or KeyNotFoundException or IndexOutOfRangeException)
        {
            return null;
        }
    }

    /// <summary>Finds the cell under a point of the sheet (pixels), for the anchor of a picture.</summary>
    private sealed class Geometry(double colW, double rowH, Dictionary<int, double> cols, SortedDictionary<int, double> rows)
    {
        public (int Col, double Offset) Column(double x)
        {
            for (var c = 1; c < A1.MaxCols; c++)
            {
                var w = cols.TryGetValue(c, out var px) ? px : colW;
                if (x < w)
                {
                    return (c, x);
                }

                x -= w;
            }

            return (A1.MaxCols, 0);
        }

        public (int Row, double Offset) Row(double y)
        {
            var row = 1;
            foreach (var (r, h) in rows)
            {
                var span = (r - row) * rowH;
                if (y < span)
                {
                    break;
                }

                y -= span;
                row = r;
                if (y < h)
                {
                    return (row, y);
                }

                y -= h;
                row++;
            }

            var more = rowH > 0 ? (int)Math.Floor(y / rowH) : 0;
            return (Math.Min(A1.MaxRows, row + more), y - (more * rowH));
        }
    }
}
