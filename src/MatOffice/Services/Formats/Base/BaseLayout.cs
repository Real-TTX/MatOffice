using System.Text.Json.Nodes;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Base;

/// <summary>
/// The layouts the Form and Report buttons of Access make, for the templates: a columnar form (labels left, controls right, a title in the header,
/// optional navigation buttons and a subform), a multiple-items form, and a stepped report with group headers, totals and page numbers.
/// The editor builds the same layouts in the browser (wwwroot/js/base/layout.js).
/// </summary>
public static class BaseLayout
{
    /// <summary>A field of a record source as the layouts need it: its name, type and caption.</summary>
    public sealed record SourceField(string Name, string Type, string Caption = "", bool Lookup = false, string Format = "");

    public const int FormWidth = 640;

    /// <summary>The width of the control of a field: a lookup shows the text of the other table, so it is wide whatever the type of its key.</summary>
    public static int ControlWidth(SourceField f) => f.Lookup ? 240 : f.Type switch
    {
        "memo" => 400,
        "number" or "autonumber" => 120,
        "currency" or "datetime" => 140,
        "yesno" => 24,
        "hyperlink" => 300,
        _ => 260,
    };

    private static JsonObject Control(ref int n, string type, int x, int y, int w, int h, Action<JsonObject>? more = null)
    {
        var c = new JsonObject { ["id"] = "c" + ++n, ["type"] = type, ["x"] = x, ["y"] = y, ["w"] = w, ["h"] = h };
        more?.Invoke(c);
        return c;
    }

    private static JsonObject Section(int height, JsonArray controls) => new() { ["height"] = height, ["visible"] = true, ["controls"] = controls };

    private static string ControlType(SourceField f) => f.Type == "yesno" ? "checkbox" : f.Lookup ? "combobox" : "textbox";

    /// <summary>A columnar form ("Form" button): one record, labels and controls, the title in the header.</summary>
    public static JsonObject ColumnarForm(string id, string name, string sourceId, string title, IReadOnlyList<SourceField> fields, Subform? sub, bool buttons)
    {
        var n = 0;
        var width = sub is null ? FormWidth : 720;
        var header = new JsonArray { Control(ref n, "label", 24, 18, width - 48, 38, c => { c["caption"] = title; c["fontSize"] = 22; }) };
        var detail = new JsonArray();
        var y = 16;
        foreach (var f in fields)
        {
            var h = f.Type == "memo" ? 72 : 24;
            var caption = string.IsNullOrEmpty(f.Caption) ? f.Name : f.Caption;
            detail.Add(Control(ref n, "label", 24, y, 150, 24, c => c["caption"] = caption));
            detail.Add(Control(ref n, ControlType(f), 184, y, Math.Min(ControlWidth(f), width - 208), h, c =>
            {
                c["source"] = f.Name;
                c["name"] = f.Name;
                if (f.Type is "number" or "currency" or "autonumber")
                {
                    c["align"] = "right";
                }

                if (f.Type == "autonumber")
                {
                    c["locked"] = true;
                }
            }));
            y += h + 10;
        }

        if (sub is not null)
        {
            y += 6;
            detail.Add(Control(ref n, "label", 24, y, 300, 24, c => { c["caption"] = sub.Caption; c["bold"] = true; }));
            y += 28;
            detail.Add(Control(ref n, "subform", 24, y, width - 48, 240, c =>
            {
                c["name"] = sub.Caption;
                c["subform"] = new JsonObject { ["source"] = sub.SourceId, ["linkMaster"] = sub.LinkMaster, ["linkChild"] = sub.LinkChild };
            }));
            y += 250;
        }

        var footer = new JsonArray();
        if (buttons)
        {
            var x = 24;
            foreach (var (action, key) in new[] { ("previous", "Base.Form.Previous"), ("next", "Base.Form.Next"), ("new", "Base.Form.New"), ("save", "Base.Form.Save"), ("delete", "Base.Form.Delete"), ("close", "Base.Form.Close") })
            {
                var caption = Text.T(key);
                var bx = x;
                footer.Add(Control(ref n, "button", bx, 12, 96, 30, c => { c["caption"] = caption; c["action"] = action; }));
                x += 104;
            }
        }

        return new JsonObject
        {
            ["id"] = id,
            ["name"] = name,
            ["source"] = sourceId,
            ["caption"] = title,
            ["kind"] = "single",
            ["width"] = width,
            ["allowEdits"] = true,
            ["allowAdditions"] = true,
            ["allowDeletions"] = true,
            ["sections"] = new JsonObject
            {
                ["header"] = Section(72, header),
                ["detail"] = Section(y + 8, detail),
                ["footer"] = Section(buttons ? 54 : 0, footer),
            },
        };
    }

    /// <summary>A Multiple Items form: one row of controls per record, the labels in the header.</summary>
    public static JsonObject TabularForm(string id, string name, string sourceId, string title, IReadOnlyList<SourceField> fields)
    {
        var n = 0;
        var labels = new JsonArray();
        var detail = new JsonArray();
        var x = 24;
        foreach (var f in fields)
        {
            var w = f.Lookup ? 160 : f.Type switch { "memo" => 220, "number" or "autonumber" => 90, "currency" or "datetime" => 110, "yesno" => 70, _ => 160 };
            var caption = string.IsNullOrEmpty(f.Caption) ? f.Name : f.Caption;
            labels.Add(Control(ref n, "label", x, 80, w, 24, c => { c["caption"] = caption; c["bold"] = true; }));
            detail.Add(Control(ref n, ControlType(f), f.Type == "yesno" ? x + (w / 2) - 12 : x, 4, f.Type == "yesno" ? 24 : w, 24, c =>
            {
                c["source"] = f.Name;
                c["name"] = f.Name;
                if (f.Type is "number" or "currency" or "autonumber")
                {
                    c["align"] = "right";
                }
            }));
            x += w + 8;
        }

        var width = Math.Max(FormWidth, x + 16);
        var header = new JsonArray { Control(ref n, "label", 24, 18, width - 48, 38, c => { c["caption"] = title; c["fontSize"] = 22; }) };
        foreach (var l in labels.ToList())
        {
            labels.Remove(l);
            header.Add(l);
        }

        return new JsonObject
        {
            ["id"] = id,
            ["name"] = name,
            ["source"] = sourceId,
            ["caption"] = title,
            ["kind"] = "continuous",
            ["width"] = width,
            ["allowEdits"] = true,
            ["allowAdditions"] = true,
            ["allowDeletions"] = true,
            ["sections"] = new JsonObject
            {
                ["header"] = Section(110, header),
                ["detail"] = Section(32, detail),
                ["footer"] = Section(0, []),
            },
        };
    }

    /// <summary>The subform of a columnar form: the related records of a child table.</summary>
    public sealed record Subform(string SourceId, string Caption, string LinkMaster, string LinkChild);

    private static double Weight(SourceField f) => f.Type switch
    {
        "memo" => 5,
        "number" or "autonumber" => 1.4,
        "currency" or "datetime" => 1.8,
        "yesno" => 1,
        _ => 3,
    };

    /// <summary>
    /// A stepped report (Report Wizard): the title and the date in the report header, the labels in the page header, one header per grouping level,
    /// the detail lines, footers with the sums per group, the grand totals in the report footer and "Page n of m" in the page footer.
    /// </summary>
    public static JsonObject SteppedReport(string id, string name, string sourceId, string title, IReadOnlyList<SourceField> fields,
        IReadOnlyList<string> groups, IReadOnlyList<string> sort, IReadOnlyList<string> sums, bool landscape)
    {
        var n = 0;
        var width = landscape ? 1047 : 718;
        var groupFields = groups.Select(g => fields.First(f => f.Name == g)).ToList();
        var detailFields = fields.Where(f => !groups.Contains(f.Name)).ToList();
        var ordered = groupFields.Concat(detailFields).ToList();
        var total = ordered.Sum(Weight);
        var gap = 8;
        var usable = width - (gap * (ordered.Count - 1));
        var columns = new Dictionary<string, (int X, int W)>(StringComparer.Ordinal);
        var x = 0;
        foreach (var f in ordered)
        {
            var w = (int)Math.Floor(usable * Weight(f) / total);
            columns[f.Name] = (x, w);
            x += w + gap;
        }

        string Caption(SourceField f) => string.IsNullOrEmpty(f.Caption) ? f.Name : f.Caption;
        void Align(JsonObject c, SourceField f)
        {
            if (f.Type is "number" or "currency" or "autonumber")
            {
                c["align"] = "right";
            }
        }

        var reportHeader = new JsonArray
        {
            Control(ref n, "label", 0, 6, (int)(width * 0.65), 40, c => { c["caption"] = title; c["fontSize"] = 22; }),
            Control(ref n, "textbox", width - 220, 14, 220, 22, c => { c["source"] = "=Date()"; c["format"] = "Long Date"; c["align"] = "right"; c["border"] = false; }),
        };
        var pageHeader = new JsonArray();
        foreach (var f in ordered)
        {
            var (cx, cw) = columns[f.Name];
            pageHeader.Add(Control(ref n, "label", cx, 6, cw, 22, c => { c["caption"] = Caption(f); c["bold"] = true; Align(c, f); }));
        }

        pageHeader.Add(Control(ref n, "line", 0, 30, width, 1));

        var groupHeaders = new JsonArray();
        var groupFooters = new JsonArray();
        for (var i = 0; i < groupFields.Count; i++)
        {
            var g = groupFields[i];
            var (gx, gw) = columns[g.Name];
            groupHeaders.Add(Section(30, [Control(ref n, "textbox", gx, 4, gw, 22, c => { c["source"] = g.Name; c["bold"] = true; c["fontSize"] = 12; c["border"] = false; })]));
            var foot = new JsonArray();
            if (sums.Count > 0)
            {
                foot.Add(Control(ref n, "textbox", gx, 6, Math.Max(gw, 200), 20, c =>
                {
                    c["source"] = "=\"" + Text.T("Base.Report.SummaryFor").Replace("\"", "\"\"", StringComparison.Ordinal) + " \" & [" + g.Name + "]";
                    c["italic"] = true;
                    c["border"] = false;
                }));
                foreach (var s in sums)
                {
                    var sf = fields.First(f => f.Name == s);
                    var (sx, sw) = columns[s];
                    foot.Add(Control(ref n, "line", sx, 2, sw, 1));
                    foot.Add(Control(ref n, "textbox", sx, 6, sw, 20, c => { c["source"] = "=Sum([" + s + "])"; c["bold"] = true; c["border"] = false; c["format"] = sf.Type == "currency" ? "Currency" : string.Empty; Align(c, sf); }));
                }
            }

            groupFooters.Add(Section(foot.Count > 0 ? 34 : 0, foot));
        }

        var detail = new JsonArray();
        foreach (var f in detailFields)
        {
            var (dx, dw) = columns[f.Name];
            detail.Add(Control(ref n, f.Type == "yesno" ? "checkbox" : "textbox", dx, 2, f.Type == "yesno" ? 18 : dw, f.Type == "yesno" ? 18 : 20, c =>
            {
                c["source"] = f.Name;
                c["border"] = false;
                if (!string.IsNullOrEmpty(f.Format))
                {
                    c["format"] = f.Format;
                }

                Align(c, f);
            }));
        }

        var reportFooter = new JsonArray();
        if (sums.Count > 0)
        {
            var first = ordered[0];
            reportFooter.Add(Control(ref n, "label", columns[first.Name].X, 10, 200, 22, c => { c["caption"] = Text.T("Base.Report.GrandTotal"); c["bold"] = true; }));
            foreach (var s in sums)
            {
                var sf = fields.First(f => f.Name == s);
                var (sx, sw) = columns[s];
                reportFooter.Add(Control(ref n, "line", sx, 4, sw, 1));
                reportFooter.Add(Control(ref n, "textbox", sx, 10, sw, 22, c => { c["source"] = "=Sum([" + s + "])"; c["bold"] = true; c["border"] = false; c["format"] = sf.Type == "currency" ? "Currency" : string.Empty; Align(c, sf); }));
            }
        }
        else
        {
            reportFooter.Add(Control(ref n, "textbox", 0, 10, 300, 22, c => { c["source"] = "=Count(*) & \" " + Text.T("Base.Report.Records").Replace("\"", "\"\"", StringComparison.Ordinal) + "\""; c["bold"] = true; c["border"] = false; }));
        }

        var pageFooter = new JsonArray
        {
            Control(ref n, "textbox", 0, 8, 260, 20, c => { c["source"] = "=Now()"; c["border"] = false; c["fontSize"] = 9; }),
            Control(ref n, "textbox", width - 260, 8, 260, 20, c =>
            {
                c["source"] = "=\"" + Text.T("Base.Report.Page") + " \" & [Page] & \" " + Text.T("Base.Report.Of") + " \" & [Pages]";
                c["align"] = "right";
                c["border"] = false;
                c["fontSize"] = 9;
            }),
        };

        var grouping = new JsonArray();
        foreach (var g in groups)
        {
            grouping.Add(new JsonObject { ["field"] = g, ["header"] = true, ["footer"] = sums.Count > 0, ["desc"] = false, ["keepTogether"] = false });
        }

        var sorting = new JsonArray();
        foreach (var s in sort)
        {
            sorting.Add(new JsonObject { ["field"] = s, ["desc"] = false });
        }

        return new JsonObject
        {
            ["id"] = id,
            ["name"] = name,
            ["source"] = sourceId,
            ["caption"] = title,
            ["orientation"] = landscape ? "landscape" : "portrait",
            ["paper"] = "a4",
            ["width"] = width,
            ["grouping"] = grouping,
            ["sorting"] = sorting,
            ["sections"] = new JsonObject
            {
                ["reportHeader"] = Section(56, reportHeader),
                ["pageHeader"] = Section(34, pageHeader),
                ["groupHeaders"] = groupHeaders,
                ["detail"] = Section(24, detail),
                ["groupFooters"] = groupFooters,
                ["pageFooter"] = Section(32, pageFooter),
                ["reportFooter"] = Section(40, reportFooter),
            },
        };
    }
}
