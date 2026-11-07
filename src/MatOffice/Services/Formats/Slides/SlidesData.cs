using System.Text.Json;
using System.Text.Json.Nodes;

namespace MatOffice.Services.Formats.Slides;

/// <summary>A placeholder of a layout ("Click to add title"): where it sits (fractions of the slide) and how its text looks by default.</summary>
public sealed record LayoutSlot(
    string Ph, string Role, string Pptx, int Idx, double X, double Y, double W, double H,
    double? Size, string Align, string Anchor, bool Bullet, bool Bold, bool Content, bool Muted);

/// <summary>A slide layout (Title Slide, Title and Content ...): its group decides background and decoration of the design.</summary>
public sealed record SlideLayout(string Id, string Group, string Pptx, IReadOnlyList<LayoutSlot> Slots)
{
    public LayoutSlot? Slot(string? ph) => Slots.FirstOrDefault(s => s.Ph == ph);
}

/// <summary>A footer placeholder (date, text, slide number).</summary>
public sealed record FooterSlot(double X, double Y, double W, double H, string Align);

/// <summary>
/// The layouts and designs (themes) of Slides, read from <c>Resources/Templates/slides/layouts.json</c> and <c>themes.json</c> (embedded).
/// The editor has the same data in <c>wwwroot/js/slides/layouts-data.js</c> and <c>themes-data.js</c>; <c>SlidesFormatTests</c> checks that both agree.
/// </summary>
public static class SlidesData
{
    private static readonly Lazy<JsonObject> LayoutsJson = new(() => Load("layouts.json"));
    private static readonly Lazy<JsonObject> ThemesJson = new(() => Load("themes.json"));
    private static readonly Lazy<IReadOnlyList<SlideLayout>> AllLayouts = new(ReadLayouts);

    /// <summary>The scheme colour names of a design.</summary>
    public static readonly string[] SchemeNames = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink"];

    public static IReadOnlyList<SlideLayout> Layouts => AllLayouts.Value;

    public static SlideLayout Layout(string? id) => Layouts.FirstOrDefault(l => l.Id == id) ?? Layouts.First(l => l.Id == "blank");

    public static bool IsLayout(string? id) => Layouts.Any(l => l.Id == id);

    public static FooterSlot FooterDate => Footer("date");

    public static FooterSlot FooterText => Footer("text");

    public static FooterSlot FooterNumber => Footer("number");

    public static double FooterSize => LayoutsJson.Value["footer"]!["size"]!.GetValue<double>();

    /// <summary>The default sizes (pt) of the body text by level 0..8.</summary>
    public static double BodySize(int level)
    {
        var sizes = LayoutsJson.Value["body"]!["sizes"]!.AsArray();
        return sizes[Math.Clamp(level, 0, sizes.Count - 1)]!.GetValue<double>();
    }

    /// <summary>The space before a body paragraph (pt) by level.</summary>
    public static double BodySpaceBefore(int level)
    {
        var sizes = LayoutsJson.Value["body"]!["spaceBefore"]!.AsArray();
        return sizes[Math.Clamp(level, 0, sizes.Count - 1)]!.GetValue<double>();
    }

    /// <summary>The indent of a bullet (px) and the step per level (px).</summary>
    public static (double Indent, double Step) BodyIndent => (LayoutsJson.Value["body"]!["indent"]!.GetValue<double>(), LayoutsJson.Value["body"]!["step"]!.GetValue<double>());

    public static double TitleLineSpacing => LayoutsJson.Value["titleLineSpacing"]!.GetValue<double>();

    public static IEnumerable<string> ThemeIds => ThemesJson.Value["themes"]!.AsArray().Select(t => t!["id"]!.GetValue<string>());

    /// <summary>A copy of a design (the default "paper" when the id is unknown), with a colour variant (0 = its own colours).</summary>
    public static JsonObject Theme(string? id, int variant = 0)
    {
        var themes = ThemesJson.Value["themes"]!.AsArray();
        var theme = (themes.FirstOrDefault(t => t!["id"]!.GetValue<string>() == id) ?? themes[0])!.DeepClone().AsObject();
        var variants = ThemesJson.Value["variants"]!.AsArray();
        if (variant > 0 && variant <= variants.Count)
        {
            var colors = theme["colors"]!.AsObject();
            foreach (var (key, value) in variants[variant - 1]!.AsObject())
            {
                if (key != "id")
                {
                    colors[key] = value!.DeepClone();
                }
            }

            theme["variant"] = variant;
        }

        return theme;
    }

    /// <summary>The raw JSON of the layouts (the test compares it with the editor's copy).</summary>
    internal static JsonObject RawLayouts => LayoutsJson.Value;

    /// <summary>The raw JSON of the designs (the test compares it with the editor's copy).</summary>
    internal static JsonObject RawThemes => ThemesJson.Value;

    private static FooterSlot Footer(string name)
    {
        var f = LayoutsJson.Value["footer"]![name]!;
        return new FooterSlot(f["x"]!.GetValue<double>(), f["y"]!.GetValue<double>(), f["w"]!.GetValue<double>(), f["h"]!.GetValue<double>(), f["align"]!.GetValue<string>());
    }

    private static IReadOnlyList<SlideLayout> ReadLayouts()
    {
        var list = new List<SlideLayout>();
        foreach (var node in LayoutsJson.Value["layouts"]!.AsArray())
        {
            var slots = node!["slots"]!.AsArray().Select(s => new LayoutSlot(
                s!["ph"]!.GetValue<string>(), s["role"]!.GetValue<string>(), s["pptx"]!.GetValue<string>(), s["idx"]!.GetValue<int>(),
                s["x"]!.GetValue<double>(), s["y"]!.GetValue<double>(), s["w"]!.GetValue<double>(), s["h"]!.GetValue<double>(),
                s["size"]?.GetValue<double>(), s["align"]?.GetValue<string>() ?? "left", s["anchor"]?.GetValue<string>() ?? "top",
                s["bullet"]?.GetValue<bool>() ?? s["role"]!.GetValue<string>() == "body", s["bold"]?.GetValue<bool>() ?? false,
                s["content"]?.GetValue<bool>() ?? false, s["muted"]?.GetValue<bool>() ?? false)).ToList();
            list.Add(new SlideLayout(node["id"]!.GetValue<string>(), node["group"]!.GetValue<string>(), node["pptx"]!.GetValue<string>(), slots));
        }

        return list;
    }

    private static JsonObject Load(string name)
    {
        var assembly = typeof(SlidesData).Assembly;
        var resource = assembly.GetManifestResourceNames().First(n => n.EndsWith(".Resources.Templates.slides." + name, StringComparison.Ordinal));
        using var stream = assembly.GetManifestResourceStream(resource)!;
        return JsonNode.Parse(stream, documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true })!.AsObject();
    }
}
