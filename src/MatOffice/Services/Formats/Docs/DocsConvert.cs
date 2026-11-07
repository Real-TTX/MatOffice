using System.Globalization;
using System.Text;
using MatOffice.Localization;

namespace MatOffice.Services.Formats.Docs;

/// <summary>What the converters share: reading a file into memory, decoding text, the numbers of notes and list items.</summary>
internal static class DocsConvert
{
    /// <summary>The largest file a converter reads.</summary>
    public const long MaxFileBytes = 200L * 1024 * 1024;

    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    /// <summary>The whole file (at most <see cref="MaxFileBytes"/>).</summary>
    public static byte[] ReadAll(Stream input)
    {
        using var memory = new MemoryStream();
        var buffer = new byte[81920];
        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            if (memory.Length + read > MaxFileBytes)
            {
                throw new FormatException(Text.T("Documents.TooLarge"));
            }

            memory.Write(buffer, 0, read);
        }

        return memory.ToArray();
    }

    /// <summary>
    /// Text of a file: a byte order mark (UTF-8, UTF-16) decides, then a declared charset (HTML), then UTF-8 if the bytes are valid UTF-8,
    /// else Windows-1252 (what Windows editors wrote for decades).
    /// </summary>
    public static string Decode(byte[] bytes, string? charset = null)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        }

        if (charset is not null && EncodingOf(charset) is { } declared && declared.CodePage != Encoding.UTF8.CodePage)
        {
            return declared.GetString(bytes);
        }

        try
        {
            return StrictUtf8.GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return (EncodingOf("windows-1252") ?? Encoding.Latin1).GetString(bytes);
        }
    }

    private static Encoding? EncodingOf(string name)
    {
        var n = name.Trim().Trim('"', '\'').ToLowerInvariant();
        if (n is "utf-8" or "utf8")
        {
            return Encoding.UTF8;
        }

        if (n is "iso-8859-1" or "latin1" or "us-ascii" or "ascii")
        {
            // Browsers read these labels as Windows-1252.
            n = "windows-1252";
        }

        try
        {
            return CodePagesEncodingProvider.Instance.GetEncoding(n) ?? Encoding.GetEncoding(n);
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    /// <summary>The number of every note: the position of its first reference in the body.</summary>
    public static Dictionary<string, int> NoteNumbers(DocsDocument document)
    {
        var numbers = new Dictionary<string, int>(StringComparer.Ordinal);
        var known = document.Notes.Select(n => n.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var p in DocsHtml.Paragraphs(document.Body))
        {
            foreach (var inline in p.Inlines)
            {
                if (inline is DocNoteRef r && known.Contains(r.Id) && !numbers.ContainsKey(r.Id))
                {
                    numbers[r.Id] = numbers.Count + 1;
                }
            }
        }

        return numbers;
    }

    /// <summary>The notes in the order of their numbers.</summary>
    public static List<(DocNote Note, int Number)> OrderedNotes(DocsDocument document)
    {
        var numbers = NoteNumbers(document);
        return document.Notes.Where(n => numbers.ContainsKey(n.Id)).Select(n => (n, numbers[n.Id])).OrderBy(x => x.Item2).ToList();
    }

    /// <summary>The text of a field when a file cannot hold the field itself.</summary>
    public static string FieldText(string kind, string name, string language) => kind switch
    {
        "date" => DateTime.Now.ToString(language.StartsWith("de", StringComparison.Ordinal) ? "dd.MM.yyyy" : "yyyy-MM-dd", CultureInfo.InvariantCulture),
        "time" => DateTime.Now.ToString("HH:mm", CultureInfo.InvariantCulture),
        "title" => name,
        _ => string.Empty,
    };
}

/// <summary>Counts list items the way the editor does: consecutive list paragraphs form one list, any other block ends it; a level restarts after a higher one.</summary>
internal sealed class ListNumbering
{
    private readonly int[] _counters = new int[9];

    private bool _inList;

    /// <summary>A block that is not a list item ends the list.</summary>
    public void Reset()
    {
        _inList = false;
        Array.Clear(_counters);
    }

    /// <summary>The number of a list item and the numbers of the levels above it (for outline numbers).</summary>
    public (int Number, int[] Path) Next(DocParagraph p)
    {
        if (!_inList)
        {
            Array.Clear(_counters);
            _inList = true;
        }

        var level = Math.Clamp(p.Level, 0, 8);
        _counters[level] = p.Start ?? _counters[level] + 1;
        for (var i = level + 1; i < 9; i++)
        {
            _counters[i] = 0;
        }

        var path = new int[level + 1];
        for (var i = 0; i <= level; i++)
        {
            path[i] = Math.Max(1, _counters[i]);
        }

        return (_counters[level], path);
    }
}
