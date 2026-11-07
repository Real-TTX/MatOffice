using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MatOffice.Services.Formats.Docs;

/// <summary>
/// The numbering of a Word file for the import: which Docs list kind a numbering level is (bullets by their character and font, numbers by
/// their format, multilevel "1.1." lists as outline) and the numbers Word shows. Word counts per abstract numbering (instances without overrides
/// continue each other); an instance with start overrides counts on its own from its first use.
/// </summary>
internal sealed class WordNumbering
{
    private readonly Dictionary<int, Instance> _instances = [];
    private readonly Dictionary<int, AbstractList> _abstracts = [];
    private readonly Dictionary<int, LevelDefinition?[]?> _levels = [];
    private readonly Dictionary<string, int?[]> _counters = new(StringComparer.Ordinal);
    private readonly StyleSheet _styles;

    public WordNumbering(Numbering? numbering, StyleSheet styles)
    {
        _styles = styles;
        if (numbering is null)
        {
            return;
        }

        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            if (abstractNum.AbstractNumberId?.Value is not { } id || _abstracts.ContainsKey(id))
            {
                continue;
            }

            var levels = new LevelDefinition?[9];
            foreach (var level in abstractNum.Elements<Level>())
            {
                if (level.LevelIndex?.Value is { } index and >= 0 and <= 8)
                {
                    levels[index] = ReadLevel(level);
                }
            }

            _abstracts[id] = new AbstractList(levels, abstractNum.GetFirstChild<NumberingStyleLink>()?.Val?.Value);
        }

        foreach (var num in numbering.Elements<NumberingInstance>())
        {
            if (num.NumberID?.Value is not { } numId || num.AbstractNumId?.Val?.Value is not { } abstractId || _instances.ContainsKey(numId))
            {
                continue;
            }

            var starts = new int?[9];
            var overrides = new LevelDefinition?[9];
            foreach (var levelOverride in num.Elements<LevelOverride>())
            {
                if (levelOverride.LevelIndex?.Value is not { } index || index is < 0 or > 8)
                {
                    continue;
                }

                starts[index] = levelOverride.StartOverrideNumberingValue?.Val?.Value;
                if (levelOverride.Level is { } level)
                {
                    overrides[index] = ReadLevel(level);
                    starts[index] ??= overrides[index]!.Start;
                }
            }

            _instances[numId] = new Instance(abstractId, starts, overrides);
        }
    }

    /// <summary>The Docs list kind of a numbering level, or null when that level shows no marker (or the numbering is unknown).</summary>
    public string? Kind(int numId, int level)
    {
        if (LevelsOf(numId) is not { } levels)
        {
            return null;
        }

        level = Math.Clamp(level, 0, 8);
        var family = KindOf(levels[0]);
        if (family == "decimal" && levels[1] is { } second && second.Text.Contains("%1", StringComparison.Ordinal) && second.Text.Contains("%2", StringComparison.Ordinal))
        {
            family = "outline";
        }

        if (level == 0)
        {
            return family;
        }

        var own = KindOf(levels[level]);
        if (own is null || family is null)
        {
            return own is null ? null : Cycle(own, level);
        }

        if (family == "outline")
        {
            return DocsStyles.NumberKinds.Contains(own) ? "outline" : Cycle(own, level);
        }

        return DocsStyles.LevelFormat(family, level) == own || (DocsStyles.NumberKinds.Contains(own) && DocsStyles.NumberKinds.Contains(family) && own == "decimal")
            ? family
            : Cycle(own, level);
    }

    /// <summary>
    /// The kind of a level of its own: bullets and numbers whose marker is the one their cycle shows on this level (◦ on level 1 of a bulleted
    /// list, a. on level 1 of a numbered one) are the plain bullet or number kind, as the editor writes a sub-list.
    /// </summary>
    private static string Cycle(string own, int level)
        => DocsStyles.LevelFormat("bullet", level) == own ? "bullet" : DocsStyles.LevelFormat("decimal", level) == own ? "decimal" : own;

    /// <summary>The number Word shows for the next paragraph of this numbering and level (call it once per paragraph, in the order of the document).</summary>
    public int? Number(int numId, int level)
    {
        if (!_instances.TryGetValue(numId, out var instance) || LevelsOf(numId) is not { } levels)
        {
            return null;
        }

        level = Math.Clamp(level, 0, 8);
        var own = instance.Starts.Any(s => s is not null);
        var key = own ? "n" + numId : "a" + instance.AbstractId;
        if (!_counters.TryGetValue(key, out var counters))
        {
            counters = new int?[9];
            _counters[key] = counters;
        }

        var start = instance.Starts[level] ?? levels[level]?.Start ?? 0;
        counters[level] = counters[level] is { } current ? current + 1 : start;
        for (var deeper = level + 1; deeper < 9; deeper++)
        {
            counters[deeper] = null;
        }

        return counters[level];
    }

    /// <summary>The level a paragraph style is bound to in a numbering (<c>w:pStyle</c> inside a level), or null.</summary>
    public int? LevelOf(int numId, string? styleId)
    {
        if (styleId is null || LevelsOf(numId) is not { } levels)
        {
            return null;
        }

        for (var i = 0; i < levels.Length; i++)
        {
            if (levels[i]?.StyleId == styleId)
            {
                return i;
            }
        }

        return null;
    }

    /// <summary>The levels of an instance: the abstract numbering (through a numbering style link if it has one) with the overridden levels.</summary>
    private LevelDefinition?[]? LevelsOf(int numId, int depth = 0)
    {
        if (_levels.TryGetValue(numId, out var cached))
        {
            return cached;
        }

        LevelDefinition?[]? result = null;
        if (depth < 4 && _instances.TryGetValue(numId, out var instance) && _abstracts.TryGetValue(instance.AbstractId, out var list))
        {
            var levels = list.Levels;
            if (levels.All(l => l is null) && list.StyleLink is { } link && _styles.NumIdOfStyle(link) is { } linked && linked != numId)
            {
                levels = LevelsOf(linked, depth + 1) ?? levels;
            }

            result = (LevelDefinition?[])levels.Clone();
            for (var i = 0; i < 9; i++)
            {
                if (instance.Overrides[i] is { } level)
                {
                    result[i] = level;
                }
            }
        }

        _levels[numId] = result;
        return result;
    }

    private static LevelDefinition ReadLevel(Level level)
    {
        var format = level.GetFirstChild<NumberingFormat>()?.Val?.InnerText;
        if (format is null && level.GetFirstChild<AlternateContent>() is { } alternate)
        {
            // Word 2010+ writes custom formats as a choice with a plain fallback
            format = alternate.Descendants<NumberingFormat>().LastOrDefault()?.Val?.InnerText;
        }

        var fonts = level.GetFirstChild<NumberingSymbolRunProperties>()?.GetFirstChild<RunFonts>();
        return new LevelDefinition(
            format ?? "decimal",
            level.GetFirstChild<LevelText>()?.Val?.Value ?? string.Empty,
            level.GetFirstChild<StartNumberingValue>()?.Val?.Value ?? 0,
            fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value,
            level.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value);
    }

    /// <summary>The Docs list kind of one level: the bullets by their character (and symbol font), the numbers by their format.</summary>
    private static string? KindOf(LevelDefinition? level)
    {
        if (level is null)
        {
            return null;
        }

        switch (level.Format)
        {
            case "none":
                return null;
            case "bullet":
                var text = level.Text.Trim();
                if (text.Length == 0)
                {
                    return null;
                }

                var c = text[0];
                var symbol = c >= 0xF000 ? (char)(c - 0xF000) : c;
                var font = level.Font ?? string.Empty;
                if (font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase))
                {
                    return symbol switch
                    {
                        (char)0xFC or (char)0xFB or (char)0xFE => "check",
                        (char)0xA7 or (char)0x6E or (char)0x71 or (char)0xA8 or (char)0x6F => "square",
                        (char)0xD8 or (char)0xE0 or (char)0xE8 or (char)0xF0 or (char)0xC4 => "arrow",
                        _ => "bullet",
                    };
                }

                return c switch
                {
                    'o' or '◦' or '○' => "circle",
                    '▪' or '■' or '□' or '▫' or '§' => "square",
                    '–' or '-' or '—' or '−' => "dash",
                    '✓' or '✔' or '☑' => "check",
                    '➢' or '➤' or '►' or '▶' or '→' or '➔' or '>' or 'Ø' => "arrow",
                    _ => "bullet",
                };
            case "lowerLetter":
                return "lower-alpha";
            case "upperLetter":
                return "upper-alpha";
            case "lowerRoman":
                return "lower-roman";
            case "upperRoman":
                return "upper-roman";
            default:
                return "decimal";
        }
    }

    private sealed record LevelDefinition(string Format, string Text, int Start, string? Font, string? StyleId);

    private sealed record AbstractList(LevelDefinition?[] Levels, string? StyleLink);

    private sealed record Instance(int AbstractId, int?[] Starts, LevelDefinition?[] Overrides);
}
