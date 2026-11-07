using System.IO.Compression;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// The pictures of the photo album template, drawn in code (no files, nothing from another site): a sunset at a boathouse, the lights of a barbecue,
/// a band on a stage. Each is a small PNG written by a minimal encoder (8-bit RGB, one filter per row chosen by the usual "smallest sum" rule).
/// </summary>
public static class SlidesPictures
{
    private static readonly Lazy<string> SunsetPicture = new(() => DataUrl(Width, Height, Sunset));
    private static readonly Lazy<string> LightsPicture = new(() => DataUrl(Width, Height, (u, v) => Lights(u, v, 1)));
    private static readonly Lazy<string> StagePicture = new(() => DataUrl(Width, Height, Stage));
    private static readonly Lazy<string> CoverPicture = new(() => DataUrl(960, 540, (u, v) => Lights(u, v, 0.45)));

    /// <summary>The size of the album pictures (3:2).</summary>
    public const int Width = 720;

    public const int Height = 480;

    public static string Boathouse => SunsetPicture.Value;

    public static string Barbecue => LightsPicture.Value;

    public static string Band => StagePicture.Value;

    /// <summary>A dim picture of lights (16:9) behind the title of the album.</summary>
    public static string Cover => CoverPicture.Value;

    // ---- the scenes (u, v: 0..1 across and down; answers r, g, b in 0..1) -----------------------------------------------------------------

    private static (double R, double G, double B) Sunset(double u, double v)
    {
        const double horizon = 0.58;
        (double R, double G, double B) Sky(double t) => Stops(t, (0, (0.10, 0.12, 0.32)), (0.45, (0.42, 0.24, 0.47)), (0.8, (0.92, 0.45, 0.36)), (1, (1.0, 0.72, 0.40)));
        double sunU = 0.64, sunV = 0.50, aspect = (double)Width / Height;
        var d = Math.Sqrt(Math.Pow((u - sunU) * aspect, 2) + Math.Pow(v - sunV, 2));
        (double R, double G, double B) c;
        if (v < horizon)
        {
            c = Sky(v / horizon);
            c = Add(c, (1.0, 0.55, 0.30), 0.35 * Math.Exp(-d * d / 0.02));
            if (d < 0.055)
            {
                c = Mix(c, (1.0, 0.88, 0.58), Smooth(0.055, 0.050, d));
            }

            // the trees on the far shore
            var tree = horizon - 0.02 - 0.012 * Math.Sin(u * 23) - (u > 0.78 ? 0.09 * Math.Max(0, 1 - Math.Abs(((u - 0.78) * 30 % 2) - 1) * 1.6) : 0);
            if (u > 0.3 && v > tree)
            {
                c = (0.10, 0.07, 0.13);
            }
        }
        else
        {
            var mirrored = Sky(Math.Max(0, horizon - (v - horizon) * 1.4) / horizon);
            c = Scale(mirrored, 0.55);
            c = Add(c, (0.05, 0.08, 0.16), 1);
            var ripple = Math.Sin(v * Height * 0.9 + Math.Sin(u * 40) * 2);
            c = Scale(c, 1 + 0.08 * ripple);
            var spread = 0.025 + (v - horizon) * 0.3;
            if (Math.Abs(u - sunU) < spread && Math.Sin(v * Height * 1.3 + u * 50) > 0.1)
            {
                c = Add(c, (0.95, 0.62, 0.32), 0.65 * (1 - (v - horizon) / (1 - horizon)) * (1 - Math.Abs(u - sunU) / spread));
            }
        }

        // the land on the left, the boathouse with a lit window, the pier
        var land = 0.60 + 0.3 * Math.Pow(u / 0.45, 2);
        if (u < 0.45 && v > land)
        {
            c = (0.06, 0.05, 0.08);
        }

        var roof = 0.39 + Math.Abs(u - 0.18) * (0.09 / 0.095);
        if ((u >= 0.10 && u <= 0.26 && v >= 0.47 && v <= 0.63) || (u >= 0.085 && u <= 0.275 && v >= roof && v <= 0.48))
        {
            c = (0.07, 0.06, 0.09);
            if (u >= 0.15 && u <= 0.19 && v >= 0.52 && v <= 0.565)
            {
                c = (1.0, 0.76, 0.38);
            }
        }

        if ((u >= 0.25 && u <= 0.52 && v >= 0.652 && v <= 0.664) || (v > 0.664 && v < 0.70 && u > 0.25 && u < 0.52 && (u * 100 % 6) < 0.6))
        {
            c = (0.08, 0.06, 0.08);
        }

        return c;
    }

    /// <summary>The out-of-focus lights of the barbecue picture: centre, radius, strength, colour (always the same: a fixed seed).</summary>
    private static readonly (double U, double V, double R, double A, (double, double, double) Colour)[] Discs = MakeDiscs();

    private static (double, double, double, double, (double, double, double))[] MakeDiscs()
    {
        var random = new Random(7);
        (double, double, double)[] colours = [(1.0, 0.70, 0.30), (1.0, 0.52, 0.20), (1.0, 0.85, 0.45), (1.0, 0.92, 0.75)];
        return Enumerable.Range(0, 42).Select(i => (random.NextDouble(), 0.25 + random.NextDouble() * 0.7, 0.025 + random.NextDouble() * 0.06, 0.10 + random.NextDouble() * 0.22, colours[i % colours.Length])).ToArray();
    }

    private static (double R, double G, double B) Lights(double u, double v, double brightness)
    {
        var aspect = 1.5;
        (double R, double G, double B) c = Stops(v, (0, (0.10, 0.06, 0.05)), (1, (0.24, 0.11, 0.04)));
        c = Add(c, (0.9, 0.42, 0.12), 0.8 * Math.Exp(-(Math.Pow(u - 0.5, 2) / 0.06 + Math.Pow(v - 1.05, 2) / 0.05)));

        // soft discs of out-of-focus lights
        foreach (var (cu, cv, r, a, colour) in Discs)
        {
            var d = Math.Sqrt(Math.Pow((u - cu) * aspect, 2) + Math.Pow(v - cv, 2));
            if (d < r * 1.1)
            {
                var inside = Smooth(r, r * 0.85, d);
                var rim = Math.Exp(-Math.Pow((d - r * 0.9) / (r * 0.08), 2));
                c = Add(c, colour, a * inside + a * 0.5 * rim);
            }
        }

        // a string of lights across the top
        var wire = 0.07 + 0.13 * Math.Sin(Math.PI * u);
        if (Math.Abs(v - wire) < 0.003)
        {
            c = Mix(c, (0.03, 0.02, 0.02), 0.8);
        }

        for (var k = 0; k < 12; k++)
        {
            var bu = 0.04 + k * 0.084;
            var bv = 0.07 + 0.13 * Math.Sin(Math.PI * bu) + 0.022;
            var d = Math.Sqrt(Math.Pow((u - bu) * aspect, 2) + Math.Pow(v - bv, 2));
            c = Add(c, (1.0, 0.75, 0.40), 0.55 * Math.Exp(-d * d / 0.0012));
            if (d < 0.012)
            {
                c = Mix(c, (1.0, 0.95, 0.80), Smooth(0.012, 0.009, d));
            }
        }

        // the edge of a wooden table
        if (v > 0.9 - 0.01 * Math.Sin(u * 7))
        {
            c = Stops((v - 0.88) / 0.12, (0, (0.20, 0.10, 0.05)), (1, (0.10, 0.05, 0.03)));
        }

        return Scale(c, brightness);
    }

    private static (double R, double G, double B) Stage(double u, double v)
    {
        var aspect = 1.5;
        (double R, double G, double B) c = Stops(v, (0, (0.06, 0.05, 0.12)), (0.8, (0.03, 0.03, 0.07)), (1, (0.02, 0.02, 0.04)));
        (double Su, double Tu, (double, double, double) Colour)[] spots = [(0.12, 0.30, (0.25, 0.45, 1.0)), (0.38, 0.46, (0.95, 0.25, 0.75)), (0.62, 0.54, (1.0, 0.65, 0.25)), (0.88, 0.70, (0.20, 0.85, 0.95))];
        foreach (var (su, tu, colour) in spots)
        {
            // a cone from (su, -0.05) to (tu, 0.86)
            double ax = (tu - su) * aspect, ay = 0.91, len = Math.Sqrt(ax * ax + ay * ay);
            double px = (u - su) * aspect, py = v + 0.05;
            var along = (px * ax + py * ay) / len;
            var across = Math.Abs(px * ay - py * ax) / len;
            if (along > 0 && v < 0.86)
            {
                var ratio = across / (along * 0.16 + 0.01);
                if (ratio < 1)
                {
                    c = Add(c, colour, 0.32 * Math.Pow(1 - ratio, 1.6) * Math.Min(1, along * 1.4));
                }
            }

            // the pool of light on the floor
            var pool = Math.Pow((u - tu) * aspect / 0.16, 2) + Math.Pow((v - 0.88) / 0.04, 2);
            if (v > 0.83)
            {
                c = Add(c, colour, 0.25 * Math.Exp(-pool));
            }
        }

        if (v > 0.83 && v < 0.835)
        {
            c = Add(c, (0.3, 0.3, 0.35), 0.4);
        }

        // three musicians and a microphone stand as silhouettes
        foreach (var mu in new[] { 0.32, 0.5, 0.68 })
        {
            var head = Math.Sqrt(Math.Pow((u - mu) * aspect, 2) + Math.Pow(v - 0.5, 2));
            var half = v < 0.58 ? 0.03 + (v - 0.53) * 0.4 : 0.05 - (v - 0.58) * 0.06;
            if (head < 0.035 || (v >= 0.53 && v <= 0.84 && Math.Abs(u - mu) < half / aspect * 1.5))
            {
                c = (0.01, 0.01, 0.02);
            }
        }

        if ((Math.Abs(u - 0.565) < 0.0025 && v > 0.52 && v < 0.84) || Math.Sqrt(Math.Pow((u - 0.565) * aspect, 2) + Math.Pow(v - 0.515, 2)) < 0.012)
        {
            c = (0.01, 0.01, 0.02);
        }

        // the heads of the audience in front
        for (var k = 0; k < 14; k++)
        {
            var hu = k * 0.078 + (k % 2) * 0.02;
            var hv = 0.99 - (k % 3) * 0.02;
            if (Math.Sqrt(Math.Pow((u - hu) * aspect, 2) + Math.Pow(v - hv, 2)) < 0.07)
            {
                c = (0.0, 0.0, 0.01);
            }
        }

        return c;
    }

    // ---- colour helpers -----------------------------------------------------------------------------------------------------------------

    private static (double R, double G, double B) Stops(double t, params (double At, (double R, double G, double B) Colour)[] stops)
    {
        t = Math.Clamp(t, 0, 1);
        for (var i = 1; i < stops.Length; i++)
        {
            if (t <= stops[i].At)
            {
                var k = (t - stops[i - 1].At) / Math.Max(1e-9, stops[i].At - stops[i - 1].At);
                return Mix(stops[i - 1].Colour, stops[i].Colour, k);
            }
        }

        return stops[^1].Colour;
    }

    private static (double R, double G, double B) Mix((double R, double G, double B) a, (double R, double G, double B) b, double k)
        => (a.R + (b.R - a.R) * k, a.G + (b.G - a.G) * k, a.B + (b.B - a.B) * k);

    private static (double R, double G, double B) Add((double R, double G, double B) a, (double R, double G, double B) b, double k)
        => (a.R + b.R * k, a.G + b.G * k, a.B + b.B * k);

    private static (double R, double G, double B) Scale((double R, double G, double B) a, double k) => (a.R * k, a.G * k, a.B * k);

    /// <summary>1 inside (d below inner), 0 outside (d above outer), smooth in between.</summary>
    private static double Smooth(double outer, double inner, double d)
    {
        var t = Math.Clamp((outer - d) / Math.Max(1e-9, outer - inner), 0, 1);
        return t * t * (3 - 2 * t);
    }

    // ---- PNG -----------------------------------------------------------------------------------------------------------------------------

    private static string DataUrl(int width, int height, Func<double, double, (double R, double G, double B)> pixel)
        => "data:image/png;base64," + Convert.ToBase64String(Png(width, height, pixel));

    /// <summary>An 8-bit RGB PNG of a picture given pixel by pixel.</summary>
    public static byte[] Png(int width, int height, Func<double, double, (double R, double G, double B)> pixel)
    {
        var stride = width * 3;
        var rows = new byte[height][];
        for (var y = 0; y < height; y++)
        {
            var row = new byte[stride];
            for (var x = 0; x < width; x++)
            {
                var (r, g, b) = pixel((x + 0.5) / width, (y + 0.5) / height);
                row[x * 3] = Byte(r);
                row[x * 3 + 1] = Byte(g);
                row[x * 3 + 2] = Byte(b);
            }

            rows[y] = row;
        }

        var raw = new byte[(stride + 1) * height];
        var candidate = new byte[stride];
        for (var y = 0; y < height; y++)
        {
            // the filter with the smallest sum of absolute values: none, sub or up
            var best = 0;
            var bestSum = long.MaxValue;
            for (var filter = 0; filter <= 2; filter++)
            {
                long sum = 0;
                for (var i = 0; i < stride; i++)
                {
                    var value = Filter(filter, rows, y, i);
                    sum += value < 128 ? value : 256 - value;
                }

                if (sum < bestSum)
                {
                    bestSum = sum;
                    best = filter;
                }
            }

            raw[y * (stride + 1)] = (byte)best;
            for (var i = 0; i < stride; i++)
            {
                candidate[i] = Filter(best, rows, y, i);
            }

            Buffer.BlockCopy(candidate, 0, raw, y * (stride + 1) + 1, stride);
        }

        using var output = new MemoryStream();
        output.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        var header = new byte[13];
        WriteInt(header, 0, width);
        WriteInt(header, 4, height);
        header[8] = 8;
        header[9] = 2;
        Chunk(output, "IHDR", header);
        using (var compressed = new MemoryStream())
        {
            using (var z = new ZLibStream(compressed, CompressionLevel.SmallestSize, leaveOpen: true))
            {
                z.Write(raw);
            }

            Chunk(output, "IDAT", compressed.ToArray());
        }

        Chunk(output, "IEND", []);
        return output.ToArray();
    }

    private static byte Filter(int filter, byte[][] rows, int y, int i) => filter switch
    {
        1 => (byte)(rows[y][i] - (i >= 3 ? rows[y][i - 3] : 0)),
        2 => (byte)(rows[y][i] - (y > 0 ? rows[y - 1][i] : 0)),
        _ => rows[y][i],
    };

    private static byte Byte(double v) => (byte)Math.Clamp((int)Math.Round(v * 255), 0, 255);

    private static void WriteInt(byte[] buffer, int at, int value)
    {
        buffer[at] = (byte)(value >> 24);
        buffer[at + 1] = (byte)(value >> 16);
        buffer[at + 2] = (byte)(value >> 8);
        buffer[at + 3] = (byte)value;
    }

    private static void Chunk(Stream output, string type, byte[] data)
    {
        var length = new byte[4];
        WriteInt(length, 0, data.Length);
        output.Write(length);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        output.Write(typeBytes);
        output.Write(data);
        var crc = new byte[4];
        WriteInt(crc, 0, (int)Crc32(typeBytes, data));
        output.Write(crc);
    }

    private static readonly uint[] CrcTable = Enumerable.Range(0, 256).Select(n =>
    {
        var c = (uint)n;
        for (var k = 0; k < 8; k++)
        {
            c = (c & 1) != 0 ? 0xedb88320 ^ (c >> 1) : c >> 1;
        }

        return c;
    }).ToArray();

    private static uint Crc32(byte[] type, byte[] data)
    {
        var c = 0xffffffffu;
        foreach (var b in type)
        {
            c = CrcTable[(c ^ b) & 0xff] ^ (c >> 8);
        }

        foreach (var b in data)
        {
            c = CrcTable[(c ^ b) & 0xff] ^ (c >> 8);
        }

        return c ^ 0xffffffffu;
    }
}
