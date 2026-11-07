using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MatOffice.Localization;
using static MatOffice.Services.Formats.Sheets.SheetsJson;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace MatOffice.Services.Formats.Sheets;

/// <summary>
/// Reads an .xlsx file into content (ClosedXML): sheets, values with the cached values of formulas, styles, sizes, merges, frozen panes, notes, links, names,
/// validations, conditional formats (what maps), the AutoFilter range, pictures. The rows and columns that carry sizes or formats are found with a light pass
/// over the XML first, so that nothing is ever asked of ClosedXML for the millions of empty cells a format of a whole row or column covers.
/// </summary>
internal sealed class SheetsXlsxReader
{
    /// <summary>At most this many cells that only carry a format (no value) are taken; the rest of such formats is dropped.</summary>
    private const int FormatOnlyBudget = 300_000;

    private const int Band = 1024;

    private readonly XLWorkbook _wb;
    private readonly Dictionary<string, SheetLayout> _layouts;
    private readonly bool _german = Languages.IsGerman;
    private readonly List<JsonObject> _styles = [new JsonObject()];
    private readonly Dictionary<string, int> _styleByJson = new(StringComparer.Ordinal) { ["{}"] = 0 };
    private readonly Dictionary<IXLStyle, int> _styleByXl = [];
    private readonly string _defaultFont;
    private readonly double _defaultSize;
    private int _cells;
    private int _formatOnly;

    private SheetsXlsxReader(XLWorkbook wb, Dictionary<string, SheetLayout> layouts)
    {
        _wb = wb;
        _layouts = layouts;
        _defaultFont = wb.Style.Font.FontName;
        _defaultSize = wb.Style.Font.FontSize;
    }

    public static string Read(Stream input)
    {
        try
        {
            using var buffer = new MemoryStream();
            input.CopyTo(buffer);
            buffer.Position = 0;
            var layouts = SheetLayout.Read(buffer);
            buffer.Position = 0;
            using var wb = new XLWorkbook(buffer);
            return new SheetsXlsxReader(wb, layouts).Workbook().ToJsonString();
        }
        catch (FormatException)
        {
            throw;
        }
        catch (Exception)
        {
            // a damaged or strange file: whatever ClosedXML or the zip reader threw
            throw new FormatException(Text.T("Documents.Import.Unreadable"));
        }
    }

    private JsonObject Workbook()
    {
        var sheets = new JsonArray();
        var active = 0;
        foreach (var ws in _wb.Worksheets.Take(SheetsLimits.MaxSheets))
        {
            if (ws.TabActive)
            {
                active = sheets.Count;
            }

            sheets.Add(Sheet(ws, sheets.Count));
        }

        if (sheets.Count == 0)
        {
            throw new FormatException(Text.T("Sheets.Invalid.NoSheets"));
        }

        var root = new JsonObject { ["version"] = 1, ["active"] = active };
        var names = Names();
        root["styles"] = new JsonArray([.. _styles]);
        if (names.Count > 0)
        {
            root["names"] = names;
        }

        root["sheets"] = sheets;
        return root;
    }

    private JsonArray Names()
    {
        var names = new JsonArray();
        foreach (var name in _wb.DefinedNames)
        {
            if (name.Visible && !name.Name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase) && Formula(StripPrefixes(name.RefersTo ?? string.Empty)) is { } reference)
            {
                names.Add(new JsonObject { ["name"] = name.Name, ["ref"] = reference });
            }
        }

        return names;
    }

    // ---- a sheet ----------------------------------------------------------------------------------------------------------------------

    private JsonObject Sheet(IXLWorksheet ws, int index)
    {
        var o = new JsonObject { ["id"] = "s" + (index + 1).ToString(CultureInfo.InvariantCulture), ["name"] = ws.Name };
        if (Colour(ws.TabColor) is { } tab)
        {
            o["color"] = tab;
        }

        if (ws.Visibility != XLWorksheetVisibility.Visible)
        {
            o["hidden"] = true;
        }

        var layout = _layouts.GetValueOrDefault(ws.Name) ?? SheetLayout.Empty;
        var geometry = new Geometry(ColPx(ws.ColumnWidth), RowPx(ws.RowHeight));
        var cols = new JsonObject();
        var colStyles = new Dictionary<int, int>();
        foreach (var c in layout.Cols)
        {
            var col = ws.Column(c);
            var entry = Line(col.Width, col.IsHidden, col.Style, ws, isRow: false, geometry.ColW, out var px, out var style);
            if (px is { } w)
            {
                geometry.Cols[c] = w;
            }

            if (col.IsHidden)
            {
                geometry.Cols[c] = 0;
            }

            if (style > 0)
            {
                colStyles[c] = style;
            }

            if (entry.Count > 0)
            {
                cols[A1.Column(c)] = entry;
            }
        }

        var rows = new JsonObject();
        var rowStyles = new Dictionary<int, int>();
        foreach (var r in layout.Rows)
        {
            var row = ws.Row(r);
            var entry = Line(row.Height, row.IsHidden, row.Style, ws, isRow: true, geometry.RowH, out var px, out var style);
            if (px is { } h)
            {
                geometry.Rows[r] = h;
            }

            if (row.IsHidden)
            {
                geometry.Rows[r] = 0;
            }

            if (style > 0)
            {
                rowStyles[r] = style;
            }

            if (entry.Count > 0)
            {
                rows[r.ToString(CultureInfo.InvariantCulture)] = entry;
            }
        }

        o["cells"] = Cells(ws, layout, rowStyles, colStyles);
        if (cols.Count > 0)
        {
            o["cols"] = cols;
        }

        if (rows.Count > 0)
        {
            o["rows"] = rows;
        }

        if (geometry.ColW != SheetsLimits.DefaultColWidth)
        {
            o["colW"] = geometry.ColW;
        }

        if (geometry.RowH != SheetsLimits.DefaultRowHeight)
        {
            o["rowH"] = geometry.RowH;
        }

        var merges = new JsonArray();
        foreach (var range in ws.MergedRanges.Take(SheetsLimits.MaxMerges))
        {
            merges.Add(AreaOf(range.RangeAddress).ToString());
        }

        if (merges.Count > 0)
        {
            o["merges"] = merges;
        }

        if (ws.SheetView.SplitRow > 0 || ws.SheetView.SplitColumn > 0)
        {
            o["freeze"] = new JsonObject { ["rows"] = Math.Max(0, ws.SheetView.SplitRow), ["cols"] = Math.Max(0, ws.SheetView.SplitColumn) };
        }

        if (ws.AutoFilter is { IsEnabled: true, Range: { } filterRange })
        {
            o["filter"] = new JsonObject { ["range"] = AreaOf(filterRange.RangeAddress).ToString() };
        }

        if (Try(() => ConditionalFormats(ws)) is { Count: > 0 } cf)
        {
            o["cf"] = cf;
        }

        if (Try(() => Validations(ws)) is { Count: > 0 } validations)
        {
            o["validations"] = validations;
        }

        if (Try(() => Pictures(ws, geometry)) is { Count: > 0 } images)
        {
            o["images"] = images;
        }

        View(ws, o);
        return o;
    }

    private JsonObject Line(double size, bool hidden, IXLStyle style, IXLWorksheet ws, bool isRow, int defaultPx, out int? px, out int styleIndex)
    {
        var entry = new JsonObject();
        px = isRow ? RowPx(size) : ColPx(size);
        if (px != defaultPx)
        {
            entry[isRow ? "h" : "w"] = px;
        }
        else
        {
            px = null;
        }

        if (hidden)
        {
            entry["hidden"] = true;
        }

        styleIndex = style.Equals(ws.Style) ? 0 : StyleIndex(style);
        if (styleIndex > 0)
        {
            entry["s"] = styleIndex;
        }

        return entry;
    }

    private JsonObject Cells(IXLWorksheet ws, SheetLayout layout, Dictionary<int, int> rowStyles, Dictionary<int, int> colStyles)
    {
        var cells = new JsonObject();
        foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.AllContents))
        {
            if (++_cells > SheetsLimits.MaxCells)
            {
                throw new FormatException(Text.T("Sheets.Invalid.TooManyCells"));
            }

            if (Cell(cell, Inherited(cell, rowStyles, colStyles)) is { } c)
            {
                cells[A1.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber)] = c;
            }
        }

        // cells with a format and nothing else, in bands of rows (a format on millions of cells is cut off by the budget)
        for (var first = 1; first <= layout.MaxRow && _formatOnly < FormatOnlyBudget; first += Band)
        {
            var band = ws.Range(first, 1, Math.Min(first + Band - 1, A1.MaxRows), A1.MaxCols);
            foreach (var cell in band.CellsUsed(XLCellsUsedOptions.NormalFormats, c => c.IsEmpty(XLCellsUsedOptions.AllContents)))
            {
                var style = StyleIndex(cell.Style);
                if (style == Inherited(cell, rowStyles, colStyles))
                {
                    continue;
                }

                if (++_formatOnly > FormatOnlyBudget || ++_cells > SheetsLimits.MaxCells)
                {
                    break;
                }

                cells[A1.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber)] = new JsonObject { ["s"] = style };
            }
        }

        return cells;
    }

    private static int Inherited(IXLCell cell, Dictionary<int, int> rowStyles, Dictionary<int, int> colStyles)
        => rowStyles.TryGetValue(cell.Address.RowNumber, out var r) ? r : colStyles.GetValueOrDefault(cell.Address.ColumnNumber);

    private JsonObject? Cell(IXLCell cell, int inherited)
    {
        var o = new JsonObject();
        var isFormula = false;
        if (cell.HasFormula && (!cell.HasArrayFormula || cell.FormulaReference?.FirstAddress is not { } first
                                || (first.RowNumber == cell.Address.RowNumber && first.ColumnNumber == cell.Address.ColumnNumber)))
        {
            var formula = Formula(StripPrefixes(cell.FormulaA1));
            if (formula is not null && !formula.StartsWith("TABLE(", StringComparison.OrdinalIgnoreCase))
            {
                o["f"] = formula;
                isFormula = true;
            }
        }

        var value = cell.CachedValue;
        switch (value.Type)
        {
            case XLDataType.Boolean:
                o["v"] = value.GetBoolean();
                break;
            case XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan:
                if (double.IsFinite(value.GetUnifiedNumber()))
                {
                    o["v"] = value.GetUnifiedNumber();
                }

                break;
            case XLDataType.Text:
                var text = Clean(value.GetText(), SheetsLimits.MaxText);
                if (text.Length > 0 || isFormula)
                {
                    o["v"] = text;
                }

                break;
            case XLDataType.Error:
                o["v"] = ErrorCode(value.GetError());
                o["t"] = "e";
                break;
        }

        var style = StyleIndex(cell.Style);
        if (style != inherited)
        {
            o["s"] = style;
        }

        if (cell.HasHyperlink && Try(() => LinkOf(cell.GetHyperlink())) is { } link)
        {
            o["link"] = link;
        }

        if (cell.HasComment && Clean(cell.GetComment().Text, SheetsLimits.MaxText) is { Length: > 0 } note)
        {
            o["note"] = note;
        }

        return o.Count > 0 ? o : null;
    }

    private static string? LinkOf(XLHyperlink link)
        => link.IsExternal ? link.ExternalAddress?.OriginalString : link.InternalAddress is { Length: > 0 } address ? "#" + address : null;

    public static string ErrorCode(XLError error) => error switch
    {
        XLError.NullValue => "#NULL!",
        XLError.DivisionByZero => "#DIV/0!",
        XLError.IncompatibleValue => "#VALUE!",
        XLError.CellReference => "#REF!",
        XLError.NameNotRecognized => "#NAME?",
        XLError.NumberInvalid => "#NUM!",
        _ => "#N/A",
    };

    /// <summary>A formula without the prefixes Excel writes for newer functions (<c>_xlfn.XLOOKUP</c> → <c>XLOOKUP</c>); texts in quotes stay as they are.</summary>
    public static string StripPrefixes(string formula)
    {
        if (formula.IndexOf("_xl", StringComparison.OrdinalIgnoreCase) < 0)
        {
            return formula;
        }

        var sb = new StringBuilder(formula.Length);
        var quoted = false;
        for (var i = 0; i < formula.Length; i++)
        {
            var c = formula[i];
            if (c == '"')
            {
                quoted = !quoted;
            }
            else if (!quoted && c == '_' && (i == 0 || !char.IsLetterOrDigit(formula[i - 1]) && formula[i - 1] != '_' && formula[i - 1] != '.'))
            {
                var rest = formula.AsSpan(i);
                if (rest.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) || rest.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase)
                    || rest.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase))
                {
                    i += 5;
                    continue;
                }
            }

            sb.Append(c);
        }

        return sb.ToString();
    }

    // ---- rules, validations, pictures, view ----------------------------------------------------------------------------------------

    private JsonArray ConditionalFormats(IXLWorksheet ws)
    {
        var result = new JsonArray();
        foreach (var cf in ws.ConditionalFormats)
        {
            var rule = Rule(cf);
            if (rule is null)
            {
                continue;
            }

            foreach (var range in cf.Ranges)
            {
                if (result.Count >= SheetsLimits.MaxRules)
                {
                    return result;
                }

                var copy = (JsonObject)rule.DeepClone();
                copy["range"] = AreaOf(range.RangeAddress).ToString();
                result.Add(copy);
            }
        }

        return result;
    }

    private JsonObject? Rule(IXLConditionalFormat cf)
    {
        string? V(int i) => cf.Values.TryGetValue(i, out var f) && f?.Value is { Length: > 0 } text ? text : null;
        var o = new JsonObject { ["range"] = string.Empty };
        switch (cf.ConditionalFormatType)
        {
            case XLConditionalFormatType.CellIs:
                var op = cf.Operator switch
                {
                    XLCFOperator.Equal => "eq",
                    XLCFOperator.NotEqual => "ne",
                    XLCFOperator.GreaterThan => "gt",
                    XLCFOperator.LessThan => "lt",
                    XLCFOperator.EqualOrGreaterThan => "ge",
                    XLCFOperator.EqualOrLessThan => "le",
                    XLCFOperator.Between => "between",
                    XLCFOperator.NotBetween => "notBetween",
                    _ => null,
                };
                if (op is null || V(1) is not { } v1)
                {
                    return null;
                }

                o["type"] = "cell";
                o["op"] = op;
                o["v1"] = RuleValue(v1);
                if (op is "between" or "notBetween")
                {
                    o["v2"] = RuleValue(V(2) ?? v1);
                }

                break;
            case XLConditionalFormatType.ContainsText or XLConditionalFormatType.NotContainsText or XLConditionalFormatType.StartsWith or XLConditionalFormatType.EndsWith:
                o["type"] = "text";
                o["op"] = cf.ConditionalFormatType switch
                {
                    XLConditionalFormatType.ContainsText => "contains",
                    XLConditionalFormatType.NotContainsText => "notContains",
                    XLConditionalFormatType.StartsWith => "begins",
                    _ => "ends",
                };
                o["v1"] = Unquote(V(1) ?? string.Empty);
                break;
            case XLConditionalFormatType.IsDuplicate:
                o["type"] = "dup";
                break;
            case XLConditionalFormatType.IsUnique:
                o["type"] = "unique";
                break;
            case XLConditionalFormatType.Top10:
                o["type"] = cf.Bottom ? "bottom" : "top";
                o["n"] = int.TryParse(V(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? Math.Clamp(n, 1, 1000) : 10;
                if (cf.Percent)
                {
                    o["percent"] = true;
                }

                break;
            case XLConditionalFormatType.Expression:
                if (V(1) is not { } formula)
                {
                    return null;
                }

                o["type"] = "formula";
                o["v1"] = StripPrefixes(formula.TrimStart('='));
                break;
            case XLConditionalFormatType.DataBar:
                o["type"] = "bar";
                o["color"] = cf.Colors.TryGetValue(1, out var bar) ? Colour(bar) ?? "#638ec6" : "#638ec6";
                return o;
            case XLConditionalFormatType.ColorScale:
                var colors = cf.Colors.OrderBy(p => p.Key).Select(p => Colour(p.Value)).ToList();
                if (colors.Count is not (2 or 3) || colors.Any(c => c is null))
                {
                    return null;
                }

                o["type"] = "scale";
                o["colors"] = new JsonArray([.. colors.Select(c => (JsonNode)c!)]);
                return o;
            default:
                return null;
        }

        o["style"] = StyleJson(cf.Style, subset: true);
        return o;
    }

    /// <summary>A value of a rule as the content keeps it: a number or a text as it is, anything else as a formula with "=".</summary>
    private static string RuleValue(string formula)
    {
        formula = StripPrefixes(formula.TrimStart('='));
        if (double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return formula;
        }

        return formula.Length >= 2 && formula[0] == '"' && formula[^1] == '"' ? Unquote(formula) : "=" + formula;
    }

    private static string Unquote(string text)
        => text.Length >= 2 && text[0] == '"' && text[^1] == '"' ? text[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal) : text;

    private static JsonArray Validations(IXLWorksheet ws)
    {
        var result = new JsonArray();
        foreach (var dv in ws.DataValidations)
        {
            var type = dv.AllowedValues switch
            {
                XLAllowedValues.List => "list",
                XLAllowedValues.WholeNumber => "whole",
                XLAllowedValues.Decimal => "decimal",
                XLAllowedValues.Date => "date",
                XLAllowedValues.Time => "time",
                XLAllowedValues.TextLength => "length",
                XLAllowedValues.Custom => "custom",
                _ => null,
            };
            var v1 = StripPrefixes(dv.MinValue ?? string.Empty).TrimStart('=');
            if (type is null || v1.Length == 0)
            {
                continue;
            }

            var o = new JsonObject { ["range"] = string.Empty, ["type"] = type };
            if (type == "list")
            {
                o["v1"] = v1.Length >= 2 && v1[0] == '"' && v1[^1] == '"' ? Unquote(v1) : "=" + v1;
                o["dropdown"] = dv.InCellDropdown;
            }
            else if (type == "custom")
            {
                o["v1"] = v1;
            }
            else
            {
                var op = dv.Operator switch
                {
                    XLOperator.EqualTo => "eq",
                    XLOperator.NotEqualTo => "ne",
                    XLOperator.GreaterThan => "gt",
                    XLOperator.LessThan => "lt",
                    XLOperator.EqualOrGreaterThan => "ge",
                    XLOperator.EqualOrLessThan => "le",
                    XLOperator.NotBetween => "notBetween",
                    _ => "between",
                };
                o["op"] = op;
                o["v1"] = RuleValue(v1);
                if (op is "between" or "notBetween")
                {
                    o["v2"] = RuleValue(StripPrefixes(dv.MaxValue ?? string.Empty) is { Length: > 0 } v2 ? v2 : v1);
                }
            }

            o["blank"] = dv.IgnoreBlanks;
            if (dv.ShowErrorMessage)
            {
                AddText(o, "error", dv.ErrorMessage);
                AddText(o, "errorTitle", dv.ErrorTitle);
                o["errorStyle"] = dv.ErrorStyle switch
                {
                    XLErrorStyle.Warning => "warning",
                    XLErrorStyle.Information => "info",
                    _ => "stop",
                };
            }

            if (dv.ShowInputMessage)
            {
                AddText(o, "prompt", dv.InputMessage);
            }

            foreach (var range in dv.Ranges)
            {
                if (result.Count >= SheetsLimits.MaxValidations)
                {
                    return result;
                }

                var copy = (JsonObject)o.DeepClone();
                copy["range"] = AreaOf(range.RangeAddress).ToString();
                result.Add(copy);
            }
        }

        return result;
    }

    private static void AddText(JsonObject o, string key, string? text)
    {
        if (!string.IsNullOrWhiteSpace(text))
        {
            o[key] = text;
        }
    }

    private static JsonArray Pictures(IXLWorksheet ws, Geometry geometry)
    {
        var result = new JsonArray();
        foreach (var picture in ws.Pictures.Take(SheetsLimits.MaxImages))
        {
            var mime = picture.Format switch
            {
                XLPictureFormat.Png => "png",
                XLPictureFormat.Jpeg => "jpeg",
                XLPictureFormat.Gif => "gif",
                XLPictureFormat.Webp => "webp",
                _ => null,
            };
            if (mime is null)
            {
                continue;
            }

            var cell = picture.TopLeftCell;
            var offset = picture.GetOffset(XLMarkerPosition.TopLeft);
            result.Add(new JsonObject
            {
                ["id"] = "i" + (result.Count + 1).ToString(CultureInfo.InvariantCulture),
                ["src"] = "data:image/" + mime + ";base64," + Convert.ToBase64String(picture.ImageStream.ToArray()),
                ["x"] = geometry.Left(cell.Address.ColumnNumber) + offset.X,
                ["y"] = geometry.Top(cell.Address.RowNumber) + offset.Y,
                ["w"] = Math.Max(1, picture.Width),
                ["h"] = Math.Max(1, picture.Height),
            });
        }

        return result;
    }

    private static void View(IXLWorksheet ws, JsonObject o)
    {
        if (!ws.ShowGridLines)
        {
            o["grid"] = false;
        }

        if (!ws.ShowRowColHeaders)
        {
            o["headings"] = false;
        }

        if (ws.SheetView.ZoomScale is > 0 and not 100)
        {
            o["zoom"] = Math.Clamp(ws.SheetView.ZoomScale, 10, 400) / 100.0;
        }

        var setup = ws.PageSetup;
        var print = new JsonObject();
        if (setup.PageOrientation == XLPageOrientation.Landscape)
        {
            print["landscape"] = true;
        }

        if (setup.PagesWide == 1 && setup.PagesTall is 0 or 1)
        {
            print["fit"] = setup.PagesTall == 1 ? "page" : "width";
        }

        if (setup.ShowGridlines)
        {
            print["gridlines"] = true;
        }

        if (setup.ShowRowAndColumnHeadings)
        {
            print["headings"] = true;
        }

        if (setup.FirstRowToRepeatAtTop > 0)
        {
            print["titles"] = $"{setup.FirstRowToRepeatAtTop}:{Math.Max(setup.FirstRowToRepeatAtTop, setup.LastRowToRepeatAtTop)}";
        }

        if (print.Count > 0)
        {
            o["print"] = print;
        }

        if (ws.IsProtected)
        {
            o["protected"] = true;
        }

        if (Try(() => ws.ActiveCell?.Address) is { } active)
        {
            o["sel"] = A1.Cell(active.RowNumber, active.ColumnNumber);
        }
    }

    // ---- styles ----------------------------------------------------------------------------------------------------------------------

    private int StyleIndex(IXLStyle style)
    {
        if (_styleByXl.TryGetValue(style, out var index))
        {
            return index;
        }

        var json = StyleJson(style, subset: false);
        var key = json.ToJsonString();
        if (!_styleByJson.TryGetValue(key, out index))
        {
            if (_styles.Count >= SheetsLimits.MaxStyles)
            {
                index = 0;
            }
            else
            {
                index = _styles.Count;
                _styles.Add(json);
                _styleByJson[key] = index;
            }
        }

        _styleByXl[style] = index;
        return index;
    }

    private JsonObject StyleJson(IXLStyle st, bool subset)
    {
        var o = new JsonObject();
        var font = st.Font;
        if (!subset && !string.Equals(font.FontName, _defaultFont, StringComparison.OrdinalIgnoreCase) && Font(font.FontName) is { } name)
        {
            o["font"] = name;
        }

        if (!subset && font.FontSize != _defaultSize && font.FontSize is >= 1 and <= 409)
        {
            o["size"] = font.FontSize;
        }

        if (font.Bold)
        {
            o["b"] = true;
        }

        if (font.Italic)
        {
            o["i"] = true;
        }

        if (font.Underline != XLFontUnderlineValues.None)
        {
            o["u"] = true;
        }

        if (font.Strikethrough)
        {
            o["strike"] = true;
        }

        if (Colour(font.FontColor) is { } color && color != "#000000")
        {
            o["color"] = color;
        }

        if (st.Fill.PatternType == XLFillPatternValues.Solid && Colour(st.Fill.BackgroundColor) is { } fill)
        {
            o["fill"] = fill;
        }

        if (subset)
        {
            return o;
        }

        var ha = st.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Left or XLAlignmentHorizontalValues.Fill => "left",
            XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.CenterContinuous => "center",
            XLAlignmentHorizontalValues.Right => "right",
            XLAlignmentHorizontalValues.Justify or XLAlignmentHorizontalValues.Distributed => "justify",
            _ => null,
        };
        if (ha is not null)
        {
            o["ha"] = ha;
        }

        var va = st.Alignment.Vertical switch
        {
            XLAlignmentVerticalValues.Top => "top",
            XLAlignmentVerticalValues.Center or XLAlignmentVerticalValues.Justify or XLAlignmentVerticalValues.Distributed => "middle",
            _ => null,
        };
        if (va is not null)
        {
            o["va"] = va;
        }

        if (st.Alignment.WrapText)
        {
            o["wrap"] = true;
        }

        if (st.Alignment.Indent is > 0 and <= 15)
        {
            o["indent"] = st.Alignment.Indent;
        }

        var nf = st.NumberFormat.Format is { Length: > 0 } code ? code : NumberFormats.BuiltIn(st.NumberFormat.NumberFormatId, _german);
        if (nf is { Length: > 0 and <= 255 } && !nf.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            o["nf"] = nf;
        }

        var border = st.Border;
        AddBorder(o, "bt", border.TopBorder, border.TopBorderColor);
        AddBorder(o, "bb", border.BottomBorder, border.BottomBorderColor);
        AddBorder(o, "bl", border.LeftBorder, border.LeftBorderColor);
        AddBorder(o, "br", border.RightBorder, border.RightBorderColor);
        if (!st.Protection.Locked)
        {
            o["unlocked"] = true;
        }

        return o;
    }

    private void AddBorder(JsonObject o, string key, XLBorderStyleValues style, XLColor color)
    {
        var name = style switch
        {
            XLBorderStyleValues.Thin => "thin",
            XLBorderStyleValues.Medium => "medium",
            XLBorderStyleValues.Thick => "thick",
            XLBorderStyleValues.Dotted => "dotted",
            XLBorderStyleValues.Double => "double",
            XLBorderStyleValues.Hair => "hair",
            XLBorderStyleValues.None => null,
            _ => "dashed",
        };
        if (name is not null)
        {
            o[key] = name + " " + (Colour(color) ?? "#000000");
        }
    }

    /// <summary>A colour as <c>#rrggbb</c>: RGB as it is, indexed from the palette, theme colours from the theme of the workbook with their tint; null for "automatic".</summary>
    private string? Colour(XLColor? color)
    {
        if (color is null || !color.HasValue)
        {
            return null;
        }

        try
        {
            switch (color.ColorType)
            {
                case XLColorType.Theme:
                    var themed = _wb.Theme.ResolveThemeColor(color.ThemeColor);
                    return themed is { HasValue: true, ColorType: XLColorType.Color } ? Hex(Tint(themed.Color, color.ThemeTint)) : null;
                case XLColorType.Indexed:
                    return color.Indexed is 64 or 65 ? null : Hex(color.Color);
                default:
                    return Hex(color.Color);
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or KeyNotFoundException or ArgumentException)
        {
            return null;
        }
    }

    private static string Hex(System.Drawing.Color c) => $"#{c.R:x2}{c.G:x2}{c.B:x2}";

    /// <summary>The tint of a theme colour (ECMA-376: the luminance moves towards black for a negative tint, towards white for a positive one).</summary>
    private static System.Drawing.Color Tint(System.Drawing.Color c, double tint)
    {
        if (tint == 0 || !double.IsFinite(tint))
        {
            return c;
        }

        double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
        double h = 0, s = 0, l = (max + min) / 2;
        if (max != min)
        {
            var d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? ((g - b) / d) + (g < b ? 6 : 0) : max == g ? ((b - r) / d) + 2 : ((r - g) / d) + 4;
            h /= 6;
        }

        l = tint < 0 ? l * (1 + tint) : (l * (1 - tint)) + tint;
        l = Math.Clamp(l, 0, 1);
        double ToRgb(double p, double q, double t)
        {
            t = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
            return t < 1.0 / 6 ? p + ((q - p) * 6 * t) : t < 0.5 ? q : t < 2.0 / 3 ? p + ((q - p) * ((2.0 / 3) - t) * 6) : p;
        }

        if (s == 0)
        {
            r = g = b = l;
        }
        else
        {
            var q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
            var p = (2 * l) - q;
            r = ToRgb(p, q, h + (1.0 / 3));
            g = ToRgb(p, q, h);
            b = ToRgb(p, q, h - (1.0 / 3));
        }

        return System.Drawing.Color.FromArgb((int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));
    }

    // ---- helpers ---------------------------------------------------------------------------------------------------------------------

    /// <summary>Excel's column width (characters of the default font) in pixels.</summary>
    public static int ColPx(double width) => (int)Math.Round(Math.Clamp(width, 0, 255) * 7 + 5);

    /// <summary>Points in pixels (96 dpi).</summary>
    public static int RowPx(double points) => (int)Math.Round(Math.Clamp(points, 0, 409) * 4 / 3);

    private static Area AreaOf(IXLRangeAddress a)
        => new(a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber, a.LastAddress.RowNumber, a.LastAddress.ColumnNumber);

    /// <summary>A part that ClosedXML cannot read is left out instead of failing the whole file.</summary>
    private static T? Try<T>(Func<T> read)
    {
        try
        {
            return read();
        }
        catch (Exception ex) when (ex is not FormatException)
        {
            return default;
        }
    }

    /// <summary>The widths of the columns and heights of the rows that differ from the default, for the place of pictures.</summary>
    private sealed class Geometry(int colW, int rowH)
    {
        public int ColW { get; } = colW;

        public int RowH { get; } = rowH;

        public Dictionary<int, int> Cols { get; } = [];

        public SortedDictionary<int, int> Rows { get; } = [];

        public int Left(int col)
        {
            var x = 0;
            for (var c = 1; c < col; c++)
            {
                x += Cols.TryGetValue(c, out var w) ? w : ColW;
            }

            return x;
        }

        public int Top(int row)
        {
            var y = (row - 1) * RowH;
            foreach (var (r, h) in Rows)
            {
                if (r >= row)
                {
                    break;
                }

                y += h - RowH;
            }

            return y;
        }
    }

    /// <summary>The rows (with a height, hidden or a format) and columns (with a <c>col</c> entry) of a sheet, read from the XML without its cells.</summary>
    private sealed record SheetLayout(SortedSet<int> Rows, SortedSet<int> Cols, int MaxRow)
    {
        public static SheetLayout Empty { get; } = new([], [], 0);

        public static Dictionary<string, SheetLayout> Read(Stream stream)
        {
            var result = new Dictionary<string, SheetLayout>(StringComparer.OrdinalIgnoreCase);
            using var doc = SpreadsheetDocument.Open(stream, false);
            var workbook = doc.WorkbookPart;
            if (workbook?.Workbook?.Sheets is null)
            {
                return result;
            }

            foreach (var sheet in workbook.Workbook.Sheets.Elements<S.Sheet>())
            {
                if (sheet.Name?.Value is { } name && sheet.Id?.Value is { } id && workbook.GetPartById(id) is WorksheetPart part)
                {
                    result[name] = ReadSheet(part);
                }
            }

            return result;
        }

        private static SheetLayout ReadSheet(WorksheetPart part)
        {
            var rows = new SortedSet<int>();
            var cols = new SortedSet<int>();
            var lastRow = 0;
            var maxRow = 0;
            using var reader = OpenXmlReader.Create(part);
            var more = reader.Read();
            while (more)
            {
                if (reader.IsStartElement && reader.ElementType == typeof(S.Row))
                {
                    var r = int.TryParse(Attribute(reader, "r"), NumberStyles.None, CultureInfo.InvariantCulture, out var number) ? number : lastRow + 1;
                    lastRow = r;
                    maxRow = Math.Max(maxRow, r);
                    if (Attribute(reader, "ht") is not null || IsOn(Attribute(reader, "hidden")) || IsOn(Attribute(reader, "customFormat")))
                    {
                        rows.Add(r);
                    }

                    reader.Skip();
                    more = !reader.EOF;
                    continue;
                }

                if (reader.IsStartElement && reader.ElementType == typeof(S.Column))
                {
                    if (int.TryParse(Attribute(reader, "min"), NumberStyles.None, CultureInfo.InvariantCulture, out var min)
                        && int.TryParse(Attribute(reader, "max"), NumberStyles.None, CultureInfo.InvariantCulture, out var max) && max < A1.MaxCols)
                    {
                        for (var c = Math.Max(1, min); c <= max; c++)
                        {
                            cols.Add(c);
                        }
                    }
                }
                else if (reader.IsEndElement && reader.ElementType == typeof(S.SheetData))
                {
                    break;
                }

                more = reader.Read();
            }

            return new SheetLayout(rows, cols, maxRow);
        }

        private static string? Attribute(OpenXmlReader reader, string name)
        {
            foreach (var a in reader.Attributes)
            {
                if (a.LocalName == name)
                {
                    return a.Value;
                }
            }

            return null;
        }

        private static bool IsOn(string? value) => value is "1" or "true";
    }
}
