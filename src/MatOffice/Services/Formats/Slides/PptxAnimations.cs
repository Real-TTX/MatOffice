using System.Text.Json.Nodes;
using System.Xml.Linq;
using static MatOffice.Services.Formats.Slides.PptxXml;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// The entrance animations of a slide as PowerPoint timing (p:timing): the main sequence with one group per click; "with previous" effects run in the
/// same block, "after previous" ones start when the block before them ends. Effects: appear, fade, fly in, zoom, wipe.
/// </summary>
internal static class PptxAnimations
{
    public static XElement? Timing(JsonObject slide, Dictionary<string, int> spids)
    {
        var anims = (slide["animations"] as JsonArray ?? []).OfType<JsonObject>().Where(a => spids.ContainsKey(a["el"]?.GetValue<string>() ?? string.Empty)).ToList();
        if (anims.Count == 0)
        {
            return null;
        }

        // click groups → blocks ("after" starts a new block) → effects
        var groups = new List<List<List<JsonObject>>>();
        foreach (var a in anims)
        {
            var trigger = a["trigger"]?.GetValue<string>() ?? "click";
            if (trigger == "click" || groups.Count == 0)
            {
                groups.Add([[a]]);
            }
            else if (trigger == "after")
            {
                groups[^1].Add([a]);
            }
            else
            {
                groups[^1][^1].Add(a);
            }
        }

        var id = 2;
        var mainChildren = new XElement(P + "childTnLst");
        var firstTrigger = anims[0]["trigger"]?.GetValue<string>() ?? "click";
        for (var g = 0; g < groups.Count; g++)
        {
            var cond = new XElement(P + "stCondLst", new XElement(P + "cond", new XAttribute("delay", "indefinite")));
            if (g == 0 && firstTrigger != "click")
            {
                cond.Add(new XElement(P + "cond", new XAttribute("evt", "onBegin"), new XAttribute("delay", "0"), new XElement(P + "tn", new XAttribute("val", "2"))));
            }

            var blocks = new XElement(P + "childTnLst");
            var start = 0.0;
            for (var b = 0; b < groups[g].Count; b++)
            {
                var effects = new XElement(P + "childTnLst");
                var longest = 0.0;
                for (var i = 0; i < groups[g][b].Count; i++)
                {
                    var a = groups[g][b][i];
                    var trigger = a["trigger"]?.GetValue<string>() ?? "click";
                    var node = i > 0 ? "withEffect" : trigger == "click" && b == 0 ? "clickEffect" : trigger == "after" ? "afterEffect" : "withEffect";
                    effects.Add(Effect(a, spids[a["el"]!.GetValue<string>()], node, ref id));
                    longest = Math.Max(longest, (a["duration"]?.GetValue<double>() ?? 0.5) + (a["delay"]?.GetValue<double>() ?? 0));
                }

                blocks.Add(new XElement(P + "par", new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("fill", "hold"),
                    new XElement(P + "stCondLst", new XElement(P + "cond", new XAttribute("delay", S((long)Math.Round(start * 1000))))), effects)));
                start += longest;
            }

            mainChildren.Add(new XElement(P + "par", new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("fill", "hold"), cond, blocks)));
        }

        var seq = new XElement(P + "seq", new XAttribute("concurrent", "1"), new XAttribute("nextAc", "seek"),
            new XElement(P + "cTn", new XAttribute("id", "2"), new XAttribute("dur", "indefinite"), new XAttribute("nodeType", "mainSeq"), mainChildren),
            new XElement(P + "prevCondLst", new XElement(P + "cond", new XAttribute("evt", "onPrev"), new XAttribute("delay", "0"), new XElement(P + "tgtEl", new XElement(P + "sldTgt")))),
            new XElement(P + "nextCondLst", new XElement(P + "cond", new XAttribute("evt", "onNext"), new XAttribute("delay", "0"), new XElement(P + "tgtEl", new XElement(P + "sldTgt")))));
        return new XElement(P + "timing", new XElement(P + "tnLst", new XElement(P + "par",
            new XElement(P + "cTn", new XAttribute("id", "1"), new XAttribute("dur", "indefinite"), new XAttribute("restart", "never"), new XAttribute("nodeType", "tmRoot"),
                new XElement(P + "childTnLst", seq)))));
    }

    private static XElement Effect(JsonObject a, int spid, string nodeType, ref int id)
    {
        var effect = a["effect"]?.GetValue<string>() ?? "fade";
        var dir = a["dir"]?.GetValue<string>() ?? "bottom";
        var ms = S((long)Math.Round((a["duration"]?.GetValue<double>() ?? 0.5) * 1000));
        var delay = S((long)Math.Round((a["delay"]?.GetValue<double>() ?? 0) * 1000));
        var (preset, subtype) = effect switch
        {
            "appear" => (1, 0),
            "fly" => (2, Subtype(dir)),
            "wipe" => (22, Subtype(dir)),
            "zoom" => (53, 16),
            _ => (10, 0),
        };
        var behaviours = new XElement(P + "childTnLst", Set(spid, ref id));
        switch (effect)
        {
            case "fade":
                behaviours.Add(Filter(spid, "fade", ms, ref id));
                break;
            case "wipe":
                behaviours.Add(Filter(spid, dir switch { "top" => "wipe(down)", "left" => "wipe(right)", "right" => "wipe(left)", _ => "wipe(up)" }, ms, ref id));
                break;
            case "fly":
                var (fromX, fromY) = dir switch
                {
                    "top" => ("#ppt_x", "0-#ppt_h/2"),
                    "left" => ("0-#ppt_w/2", "#ppt_y"),
                    "right" => ("1+#ppt_w/2", "#ppt_y"),
                    _ => ("#ppt_x", "1+#ppt_h/2"),
                };
                behaviours.Add(Anim(spid, "ppt_x", fromX, "#ppt_x", ms, ref id), Anim(spid, "ppt_y", fromY, "#ppt_y", ms, ref id));
                break;
            case "zoom":
                behaviours.Add(Anim(spid, "ppt_w", "0", "#ppt_w", ms, ref id), Anim(spid, "ppt_h", "0", "#ppt_h", ms, ref id), Filter(spid, "fade", ms, ref id));
                break;
        }

        return new XElement(P + "par", new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("presetID", S(preset)), new XAttribute("presetClass", "entr"),
            new XAttribute("presetSubtype", S(subtype)), new XAttribute("fill", "hold"), new XAttribute("nodeType", nodeType),
            new XElement(P + "stCondLst", new XElement(P + "cond", new XAttribute("delay", delay))), behaviours));
    }

    private static int Subtype(string dir) => dir switch { "top" => 1, "right" => 2, "left" => 8, _ => 4 };

    private static XElement Target(int spid) => new(P + "tgtEl", new XElement(P + "spTgt", new XAttribute("spid", S(spid))));

    private static XElement Set(int spid, ref int id) => new(P + "set",
        new XElement(P + "cBhvr",
            new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("dur", "1"), new XAttribute("fill", "hold"), new XElement(P + "stCondLst", new XElement(P + "cond", new XAttribute("delay", "0")))),
            Target(spid), new XElement(P + "attrNameLst", new XElement(P + "attrName", "style.visibility"))),
        new XElement(P + "to", new XElement(P + "strVal", new XAttribute("val", "visible"))));

    private static XElement Filter(int spid, string filter, string ms, ref int id) => new(P + "animEffect", new XAttribute("transition", "in"), new XAttribute("filter", filter),
        new XElement(P + "cBhvr", new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("dur", ms)), Target(spid)));

    private static XElement Anim(int spid, string attr, string from, string to, string ms, ref int id) => new(P + "anim", new XAttribute("calcmode", "lin"), new XAttribute("valueType", "num"),
        new XElement(P + "cBhvr", new XAttribute("additive", "base"),
            new XElement(P + "cTn", new XAttribute("id", S(++id)), new XAttribute("dur", ms), new XAttribute("fill", "hold")),
            Target(spid), new XElement(P + "attrNameLst", new XElement(P + "attrName", attr))),
        new XElement(P + "tavLst",
            new XElement(P + "tav", new XAttribute("tm", "0"), new XElement(P + "val", new XElement(P + "strVal", new XAttribute("val", from)))),
            new XElement(P + "tav", new XAttribute("tm", "100000"), new XElement(P + "val", new XElement(P + "strVal", new XAttribute("val", to))))));
}
