using System.Text.Json.Nodes;

namespace MatOffice.Services.Formats.Slides;

/// <summary>
/// The templates of Slides: a blank presentation and realistic decks (pitch, project status, workshop, quarterly review, photo album) in English and German,
/// each with a design and slides of different layouts. Built in code with <see cref="Deck"/> so that the placeholders sit exactly where the layouts put them.
/// </summary>
public static class SlidesTemplates
{
    public static IReadOnlyList<KindTemplate> All { get; } =
    [
        new("blank", "Templates.Blank", "Templates.Blank.Text", "kind-slides"),
        new("pitch", "Slides.Template.Pitch", "Slides.Template.Pitch.Text", "rocket"),
        new("status", "Slides.Template.Status", "Slides.Template.Status.Text", "activity"),
        new("workshop", "Slides.Template.Workshop", "Slides.Template.Workshop.Text", "book"),
        new("review", "Slides.Template.Review", "Slides.Template.Review.Text", "bar-chart"),
        new("album", "Slides.Template.Album", "Slides.Template.Album.Text", "image"),
    ];

    /// <summary>The content of a template in a language (en, de).</summary>
    public static JsonObject Build(string id, string language)
    {
        var de = language == "de";
        return id switch
        {
            "pitch" => Pitch(de),
            "status" => Status(de),
            "workshop" => Workshop(de),
            "review" => Review(de),
            "album" => Album(de),
            _ => Blank(),
        };
    }

    private static JsonObject Blank()
    {
        var deck = new Deck("paper");
        deck.Slide("title");
        return deck.Root;
    }

    private static JsonObject Pitch(bool de)
    {
        var d = new Deck("horizon");
        string L(string en, string ge) => de ? ge : en;
        d.Slide("title").Ph("title", "Lumen Grid").Ph("subtitle", L("Smart energy for every building · Seed round 2027", "Intelligente Energie für jedes Gebäude · Seed-Runde 2027"))
            .Notes(L("Introduce yourself and the company in one sentence. Mention the 94 buildings already live.", "Stell dich und das Unternehmen in einem Satz vor. Erwähne die 94 Gebäude, die schon laufen."));
        d.Slide("titleContent").Ph("title", L("The problem", "Das Problem")).Ph("body",
                L("Buildings waste up to 30 % of the energy they use", "Gebäude verschwenden bis zu 30 % ihrer Energie"),
                L("Owners see the bill – not the cause", "Eigentümer sehen die Rechnung – nicht die Ursache"),
                L("Retrofits are slow, expensive and disruptive", "Sanierungen sind langsam, teuer und störend"),
                "\t" + L("Average project: 14 months, €180 per m²", "Durchschnittliches Projekt: 14 Monate, 180 € pro m²"))
            .Notes(L("Tell the story of the office building in Leipzig that heated empty floors all winter.", "Erzähl die Geschichte des Bürogebäudes in Leipzig, das den ganzen Winter leere Etagen geheizt hat."));
        d.Slide("twoContent").Ph("title", L("Our solution", "Unsere Lösung")).Ph("left",
                L("Plug-in sensors in every room", "Sensoren zum Einstecken in jedem Raum"),
                L("A model that learns the building in 14 days", "Ein Modell, das das Gebäude in 14 Tagen lernt"),
                L("Automatic control of heating and light", "Automatische Steuerung von Heizung und Licht"))
            .Ph("right",
                L("No construction work", "Keine Bauarbeiten"),
                L("Saves 18–27 % from the first month", "Spart 18–27 % ab dem ersten Monat"),
                L("Works with the systems already there", "Arbeitet mit den vorhandenen Anlagen"));
        var market = d.Slide("titleOnly").Ph("title", L("Market", "Markt"));
        var big = market.Shape("ellipse", 120, 230, 330, 330, "accent4", L("€48 bn\nTotal market in Europe", "48 Mrd. €\nGesamtmarkt in Europa"), size: 22, textColor: "dk1");
        var mid = market.Shape("ellipse", 500, 270, 280, 280, "accent2", L("€6.2 bn\nOffices and schools", "6,2 Mrd. €\nBüros und Schulen"), size: 20, textColor: "dk1");
        var small = market.Shape("ellipse", 830, 310, 240, 240, "accent1", L("€410 m\nOur target in 5 years", "410 Mio. €\nUnser Ziel in 5 Jahren"), size: 18, textColor: "dk1");
        market.Animate(big, "zoom", "click").Animate(mid, "zoom", "click").Animate(small, "zoom", "click")
            .Notes(L("Bottom-up: 1.9 million non-residential buildings, €4 per room and month.", "Bottom-up: 1,9 Millionen Nichtwohngebäude, 4 € pro Raum und Monat."));
        d.Slide("titleContent").Ph("title", L("Business model", "Geschäftsmodell")).Ph("body",
            L("Hardware at cost, software as a subscription", "Hardware zum Selbstkostenpreis, Software im Abonnement"),
            L("€4 per room and month, billed yearly", "4 € pro Raum und Monat, jährlich abgerechnet"),
            L("Customers earn the money back in under 9 months", "Kunden haben das Geld in unter 9 Monaten zurück"),
            L("Gross margin of 78 % on software", "Bruttomarge von 78 % auf Software"));
        d.Slide("titleOnly").Ph("title", L("Traction", "Wachstum")).Table(110, 200, 1060, 56,
        [
            [L("Quarter", "Quartal"), L("Buildings", "Gebäude"), L("Rooms", "Räume"), L("Annual revenue", "Jahresumsatz")],
            ["Q1 2026", "12", L("1,450", "1.450"), L("€70k", "70 Tsd. €")],
            ["Q2 2026", "31", L("3,900", "3.900"), L("€187k", "187 Tsd. €")],
            ["Q3 2026", "58", L("7,600", "7.600"), L("€365k", "365 Tsd. €")],
            ["Q4 2026", "94", L("12,300", "12.300"), L("€590k", "590 Tsd. €")],
        ], "accent2");
        var team = d.Slide("titleOnly").Ph("title", "Team");
        string[][] people =
        [
            ["ML", "Mara Lindqvist", L("CEO · 10 years of energy consulting", "CEO · 10 Jahre Energieberatung")],
            ["JO", "Jonas Okafor", L("CTO · building automation, ex-Siemens", "CTO · Gebäudeautomation, ehemals Siemens")],
            ["ED", "Elif Demir", L("Head of sales · 200+ property managers", "Vertrieb · über 200 Hausverwaltungen")],
        ];
        for (var i = 0; i < people.Length; i++)
        {
            var x = 150 + (i * 360);
            team.Shape("ellipse", x + 70, 220, 160, 160, i == 0 ? "accent1" : i == 1 ? "accent2" : "accent4", people[i][0], size: 40, textColor: "dk1", bold: true);
            team.TextBox(x - 20, 400, 340, 44, people[i][1], 24, "lt1", bold: true, align: "center");
            team.TextBox(x - 20, 446, 340, 70, people[i][2], 16, "lt2", align: "center");
        }

        d.Slide("section").Ph("title", L("The ask", "Unser Angebot")).Ph("body", L("€2.5 m seed round to reach 400 buildings by the end of 2028", "2,5 Mio. € Seed-Runde, um bis Ende 2028 400 Gebäude zu erreichen"));
        d.Slide("title").Ph("title", L("Thank you", "Danke")).Ph("subtitle", "mara@lumengrid.example · lumengrid.example");
        d.AllTransitions("fade");
        return d.Root;
    }

    private static JsonObject Status(bool de)
    {
        var d = new Deck("banner");
        string L(string en, string ge) => de ? ge : en;
        d.Slide("title").Ph("title", L("Project Atlas – status report", "Projekt Atlas – Statusbericht")).Ph("subtitle", L("Steering committee · week 42", "Lenkungsausschuss · KW 42"));
        d.Slide("titleContent").Ph("title", L("Agenda", "Agenda")).Ph("body",
            L("Summary and overall status", "Zusammenfassung und Gesamtstatus"),
            L("Milestones", "Meilensteine"),
            L("Risks and issues", "Risiken und Probleme"),
            L("Decisions needed", "Benötigte Entscheidungen"),
            L("Next steps", "Nächste Schritte"));
        var glance = d.Slide("titleOnly").Ph("title", L("Status at a glance", "Status auf einen Blick"));
        var table = glance.Table(90, 210, 1100, 62,
        [
            [L("Area", "Bereich"), "Status", L("Comment", "Kommentar")],
            [L("Scope", "Umfang"), L("On track", "Im Plan"), L("No changes since the last report", "Keine Änderungen seit dem letzten Bericht")],
            [L("Schedule", "Termine"), L("At risk", "Gefährdet"), L("Data migration two weeks late", "Datenmigration zwei Wochen im Verzug")],
            [L("Budget", "Budget"), L("On track", "Im Plan"), L("62 % spent, 58 % done", "62 % ausgegeben, 58 % erledigt")],
            [L("Quality", "Qualität"), L("On track", "Im Plan"), L("214 of 230 test cases passed", "214 von 230 Testfällen bestanden")],
        ], "accent1", widths: [240, 200, 660]);
        Deck.CellFill(table, 1, 1, "#d8f0dc");
        Deck.CellFill(table, 2, 1, "#ffe8c2");
        Deck.CellFill(table, 3, 1, "#d8f0dc");
        Deck.CellFill(table, 4, 1, "#d8f0dc");
        var miles = d.Slide("titleOnly").Ph("title", L("Milestones", "Meilensteine"));
        string[] steps = [L("Design\nJune ✓", "Konzept\nJuni ✓"), L("Build\nAugust ✓", "Umsetzung\nAugust ✓"), L("Migration\nOctober", "Migration\nOktober"), L("Test\nNovember", "Test\nNovember"), L("Go-live\nJanuary", "Start\nJanuar")];
        for (var i = 0; i < steps.Length; i++)
        {
            var fill = i < 2 ? "accent3" : i == 2 ? "accent2" : "dk2:+50";
            var id = miles.Shape(i == 0 ? "homePlate" : "chevron", 90 + (i * 222), 300, 250, 120, fill, steps[i], size: 18, textColor: "lt1", bold: i == 2);
            miles.Animate(id, "fade", i == 0 ? "click" : "after");
        }

        miles.TextBox(90, 470, 1100, 60, L("We are here: migration dry run 2 starts on 24 October.", "Wir sind hier: Probelauf 2 der Migration beginnt am 24. Oktober."), 20, "dk2", align: "center");
        d.Slide("comparison").Ph("title", L("Risks and issues", "Risiken und Probleme"))
            .Ph("leftHead", L("Risks", "Risiken"))
            .Ph("left", L("Old data is of lower quality than expected", "Altdaten sind schlechter als erwartet"), L("Key users are busy at the end of the year", "Key-User sind zum Jahresende ausgelastet"), L("One developer leaves in November", "Ein Entwickler geht im November"))
            .Ph("rightHead", L("What we do about them", "Was wir dagegen tun"))
            .Ph("right", L("Extra cleansing sprint with the data team", "Zusätzlicher Bereinigungssprint mit dem Datenteam"), L("Test plan agreed with the department heads", "Testplan mit den Abteilungsleitungen abgestimmt"), L("Handover and pairing until 15 November", "Übergabe und Pairing bis 15. November"));
        d.Slide("titleContent").Ph("title", L("Decisions needed", "Benötigte Entscheidungen")).Ph("body",
            L("Approve two more weeks for the data migration", "Zwei zusätzliche Wochen für die Datenmigration genehmigen"),
            L("Confirm the go-live date of 12 January", "Starttermin 12. Januar bestätigen"),
            L("Release €40,000 from the contingency budget", "40.000 € aus der Reserve freigeben"));
        d.Slide("titleContent").Ph("title", L("Next steps", "Nächste Schritte")).Ph("body",
            L("Finish migration dry run 2 (31 October)", "Probelauf 2 der Migration abschließen (31. Oktober)"),
            L("Start user acceptance tests (4 November)", "Abnahmetests starten (4. November)"),
            L("Prepare the training for 120 users", "Schulung für 120 Personen vorbereiten"),
            L("Next steering committee: 18 November", "Nächster Lenkungsausschuss: 18. November"));
        d.Root["footer"] = new JsonObject { ["number"] = true, ["date"] = false, ["dateText"] = string.Empty, ["text"] = L("Project Atlas · internal", "Projekt Atlas · intern"), ["skipTitle"] = true };
        d.AllTransitions("push");
        return d.Root;
    }

    private static JsonObject Workshop(bool de)
    {
        var d = new Deck("coral");
        string L(string en, string ge) => de ? ge : en;
        d.Slide("title").Ph("title", L("Writing clear emails", "Klare E-Mails schreiben")).Ph("subtitle", L("Workshop for the customer service team · 90 minutes", "Workshop für das Kundenservice-Team · 90 Minuten"))
            .Notes(L("Welcome everybody, ask who wrote more than 30 emails yesterday.", "Alle begrüßen und fragen, wer gestern mehr als 30 E-Mails geschrieben hat."));
        d.Slide("titleContent").Ph("title", L("What you will learn", "Was du lernst")).Ph("body",
            L("Put the main point into the first two lines", "Das Wichtigste in die ersten zwei Zeilen"),
            L("Structure long emails so they can be scanned", "Lange E-Mails so gliedern, dass man sie überfliegen kann"),
            L("Choose a friendly and clear tone", "Einen freundlichen, klaren Ton wählen"),
            L("Avoid the five most common mistakes", "Die fünf häufigsten Fehler vermeiden"));
        d.Slide("titleOnly").Ph("title", L("Agenda", "Ablauf")).Table(200, 190, 880, 52,
        [
            [L("Time", "Zeit"), L("Topic", "Thema")],
            ["09:00", L("Welcome and goals", "Begrüßung und Ziele")],
            ["09:10", L("Why clear emails matter", "Warum klare E-Mails wichtig sind")],
            ["09:25", L("The first two lines", "Die ersten zwei Zeilen")],
            ["09:50", L("Break", "Pause")],
            ["10:00", L("Structure and tone", "Aufbau und Ton")],
            ["10:20", L("Exercise in pairs", "Übung zu zweit")],
            ["10:25", L("Wrap-up", "Abschluss")],
        ], "accent2", widths: [180, 700]);
        d.Slide("section").Ph("title", L("The first two lines", "Die ersten zwei Zeilen")).Ph("body", L("Say what you want – and what the reader should do", "Sag, was du willst – und was die Leserin tun soll"));
        d.Slide("twoContent").Ph("title", L("Before and after", "Vorher und nachher"))
            .Ph("left", L("Before:", "Vorher:"), "\t" + L("“I am writing to follow up on the matter we discussed in our last meeting regarding the contract …”", "„Ich schreibe Ihnen bezüglich der Angelegenheit, die wir in unserem letzten Termin zum Vertrag besprochen haben …“"))
            .Ph("right", L("After:", "Nachher:"), "\t" + L("“Could you send the signed contract by Friday? Then we can start on 1 March.”", "„Könnten Sie den unterschriebenen Vertrag bis Freitag schicken? Dann können wir am 1. März starten.“"));
        var exercise = d.Slide("titleOnly").Ph("title", L("Exercise", "Übung"));
        var bubble = exercise.Shape("wedgeRectCallout", 140, 220, 700, 260, "accent2", L("Rewrite the email on your handout. Put the request into the first two lines. You have 10 minutes – then swap with your neighbour.", "Schreib die E-Mail auf deinem Arbeitsblatt neu. Die Bitte gehört in die ersten zwei Zeilen. Du hast 10 Minuten – dann tauscht ihr mit dem Nachbarn."), size: 24, textColor: "lt1");
        var timer = exercise.Shape("ellipse", 920, 250, 220, 220, "accent1", L("10\nminutes", "10\nMinuten"), size: 32, textColor: "lt1", bold: true);
        exercise.Animate(bubble, "fly", "click", "bottom").Animate(timer, "zoom", "after");
        d.Slide("titleContent").Ph("title", L("Key takeaways", "Das Wichtigste")).Ph("body",
            L("Main point first", "Das Wichtigste zuerst"),
            L("One email, one topic", "Eine E-Mail, ein Thema"),
            L("Short paragraphs and lists", "Kurze Absätze und Listen"),
            L("Read it aloud before you send it", "Vor dem Senden laut lesen"));
        d.Slide("title").Ph("title", L("Questions?", "Fragen?")).Ph("subtitle", L("Thank you for joining", "Danke fürs Mitmachen"));
        d.AllTransitions("fade");
        return d.Root;
    }

    private static JsonObject Review(bool de)
    {
        var d = new Deck("meridian");
        string L(string en, string ge) => de ? ge : en;
        d.Slide("title").Ph("title", L("Quarterly business review", "Quartalsbericht")).Ph("subtitle", L("Q3 2026 · Sales and operations", "Q3 2026 · Vertrieb und Betrieb"));
        var kpi = d.Slide("titleOnly").Ph("title", L("Highlights", "Höhepunkte"));
        string[][] cards =
        [
            [L("€12.4 m", "12,4 Mio. €"), L("Revenue (+14 %)", "Umsatz (+14 %)")],
            [L("1,284", "1.284"), L("New customers", "Neue Kunden")],
            [L("96.2 %", "96,2 %"), L("On-time delivery", "Pünktliche Lieferung")],
            ["48", L("Net promoter score", "Net Promoter Score")],
        ];
        for (var i = 0; i < cards.Length; i++)
        {
            var x = 88 + (i * 282);
            var card = kpi.Card(x, 230, 258, 250, "accent" + (i + 1), cards[i][0], cards[i][1]);
            kpi.Animate(card, "fly", i == 0 ? "click" : "with", "bottom");
        }

        d.Slide("titleOnly").Ph("title", L("Revenue by region", "Umsatz nach Region")).Table(140, 200, 1000, 58,
        [
            [L("Region", "Region"), "Q3 2025", "Q3 2026", L("Change", "Veränderung")],
            [L("North", "Nord"), L("€3.1 m", "3,1 Mio. €"), L("€3.6 m", "3,6 Mio. €"), "+16 %"],
            [L("South", "Süd"), L("€2.7 m", "2,7 Mio. €"), L("€2.9 m", "2,9 Mio. €"), "+7 %"],
            [L("West", "West"), L("€2.4 m", "2,4 Mio. €"), L("€2.9 m", "2,9 Mio. €"), "+21 %"],
            [L("East", "Ost"), L("€2.7 m", "2,7 Mio. €"), L("€3.0 m", "3,0 Mio. €"), "+11 %"],
            [L("Total", "Gesamt"), L("€10.9 m", "10,9 Mio. €"), L("€12.4 m", "12,4 Mio. €"), "+14 %"],
        ], "accent1", lastRow: true);
        d.Slide("comparison").Ph("title", L("Plan and actual", "Plan und Ist"))
            .Ph("leftHead", L("What went well", "Was gut lief"))
            .Ph("left", L("Strong demand in the West", "Starke Nachfrage im Westen"), L("Partner channel grew by 30 %", "Partnerkanal um 30 % gewachsen"), L("Delivery time down to 2.1 days", "Lieferzeit auf 2,1 Tage gesenkt"))
            .Ph("rightHead", L("What we must improve", "Was besser werden muss"))
            .Ph("right", L("Churn of small accounts up to 4.8 %", "Abwanderung kleiner Kunden auf 4,8 % gestiegen"), L("Hiring for support behind plan", "Einstellungen im Support hinter Plan"), L("Marketing 6 % over budget", "Marketing 6 % über Budget"));
        d.Slide("titleContent").Ph("title", L("Goals for Q4", "Ziele für Q4")).Ph("body",
            L("Revenue of €13.5 m", "Umsatz von 13,5 Mio. €"),
            L("Launch the self-service portal", "Selbstbedienungsportal starten"),
            L("Bring churn of small accounts below 4 %", "Abwanderung kleiner Kunden unter 4 % senken"),
            L("Hire six support agents", "Sechs Personen für den Support einstellen"));
        d.Slide("section").Ph("title", L("Thank you", "Vielen Dank")).Ph("body", L("Questions and discussion", "Fragen und Diskussion"));
        d.Root["footer"] = new JsonObject { ["number"] = true, ["date"] = true, ["dateText"] = string.Empty, ["text"] = L("Confidential", "Vertraulich"), ["skipTitle"] = true };
        d.AllTransitions("fade");
        return d.Root;
    }

    private static JsonObject Album(bool de)
    {
        var d = new Deck("slate");
        string L(string en, string ge) => de ? ge : en;
        d.Slide("title").Ph("title", L("Summer party 2026", "Sommerfest 2026")).Ph("subtitle", L("Pictures of a great evening", "Bilder eines schönen Abends"));
        d.Slide("pictureCaption").Ph("title", L("Arrival", "Ankunft")).Ph("caption", L("Guests arriving at the boathouse at sunset.", "Die Gäste kommen bei Sonnenuntergang am Bootshaus an."));
        d.Slide("pictureCaption").Ph("title", L("The barbecue", "Das Grillen")).Ph("caption", L("Our kitchen team grilled for 140 people.", "Unser Küchenteam hat für 140 Personen gegrillt."));
        d.Slide("pictureCaption").Ph("title", L("The band", "Die Band")).Ph("caption", L("Live music by The Riverside Five until midnight.", "Livemusik von The Riverside Five bis Mitternacht."));
        d.Slide("titleContent").Ph("title", L("Thank you to", "Danke an")).Ph("body",
            L("The organising team", "Das Organisationsteam"),
            L("Our kitchen crew", "Unsere Küchencrew"),
            L("The Riverside Boat Club for the venue", "Den Riverside Bootsclub für den Ort"),
            L("Everyone who brought a dessert", "Alle, die einen Nachtisch mitgebracht haben"));
        d.Slide("title").Ph("title", L("See you next year!", "Bis nächstes Jahr!")).Ph("subtitle", L("Summer party 2027 · save the date: 3 July", "Sommerfest 2027 · schon vormerken: 3. Juli"));
        d.AllTransitions("zoom");
        return d.Root;
    }
}

/// <summary>Builds the content of a presentation slide by slide (used by the templates and the tests).</summary>
public sealed class Deck
{
    public const double Width = 1280;
    public const double Height = 720;

    private readonly JsonArray _slides = [];

    public Deck(string themeId, int variant = 0)
    {
        Root = new JsonObject
        {
            ["version"] = 1,
            ["size"] = new JsonObject { ["w"] = Width, ["h"] = Height },
            ["theme"] = SlidesData.Theme(themeId, variant),
            ["footer"] = new JsonObject { ["number"] = false, ["date"] = false, ["dateText"] = string.Empty, ["text"] = string.Empty, ["skipTitle"] = true },
            ["slides"] = _slides,
        };
    }

    public JsonObject Root { get; }

    /// <summary>A new slide with every placeholder of its layout (empty).</summary>
    public SlideBuilder Slide(string layoutId)
    {
        var layout = SlidesData.Layout(layoutId);
        var n = _slides.Count + 1;
        var slide = new JsonObject { ["id"] = "s" + n, ["layout"] = layout.Id, ["elements"] = new JsonArray() };
        _slides.Add(slide);
        var builder = new SlideBuilder(slide, "s" + n);
        foreach (var slot in layout.Slots)
        {
            builder.Placeholder(slot);
        }

        return builder;
    }

    public void AllTransitions(string type)
    {
        foreach (var s in _slides.OfType<JsonObject>())
        {
            s["transition"] = new JsonObject { ["type"] = type, ["duration"] = 0.7 };
        }
    }

    /// <summary>Gives one cell of a table its own fill.</summary>
    public static void CellFill(JsonObject table, int row, int col, string color)
        => table["rows"]![row]!["cells"]![col]!["fill"] = new JsonObject { ["type"] = "solid", ["color"] = color };

    /// <summary>A text body from lines: a leading tab per level ("\tSub point").</summary>
    public static JsonObject Text(IEnumerable<string> lines, double? size = null, string? color = null, bool bold = false, string? align = null)
    {
        var paras = new JsonArray();
        foreach (var line in lines)
        {
            var level = line.TakeWhile(c => c == '\t').Count();
            var run = new JsonObject { ["t"] = line[level..] };
            if (size is { } s)
            {
                run["size"] = s;
            }

            if (color is not null)
            {
                run["color"] = color;
            }

            if (bold)
            {
                run["b"] = true;
            }

            var p = new JsonObject { ["runs"] = new JsonArray(run) };
            if (level > 0)
            {
                p["level"] = level;
            }

            if (align is not null)
            {
                p["align"] = align;
            }

            paras.Add(p);
        }

        return new JsonObject { ["paras"] = paras };
    }
}

/// <summary>One slide of a <see cref="Deck"/>.</summary>
public sealed class SlideBuilder(JsonObject slide, string prefix)
{
    private int _next;

    public JsonObject Slide => slide;

    private JsonArray Elements => slide["elements"]!.AsArray();

    private string NewId() => prefix + "e" + ++_next;

    internal void Placeholder(LayoutSlot slot)
    {
        var e = new JsonObject
        {
            ["id"] = NewId(),
            ["type"] = slot.Role == "picture" ? "shape" : "text",
            ["ph"] = slot.Ph,
            ["x"] = Math.Round(slot.X * Deck.Width, 2),
            ["y"] = Math.Round(slot.Y * Deck.Height, 2),
            ["w"] = Math.Round(slot.W * Deck.Width, 2),
            ["h"] = Math.Round(slot.H * Deck.Height, 2),
        };
        if (slot.Role == "picture")
        {
            e["shape"] = "rect";
        }

        e["text"] = new JsonObject { ["paras"] = new JsonArray(new JsonObject { ["runs"] = new JsonArray(new JsonObject { ["t"] = string.Empty }) }), ["autofit"] = "shrink" };
        Elements.Add(e);
    }

    /// <summary>Fills a placeholder with lines (a leading tab per level).</summary>
    public SlideBuilder Ph(string ph, params string[] lines)
    {
        var e = Elements.OfType<JsonObject>().First(x => x["ph"]?.GetValue<string>() == ph);
        var text = Deck.Text(lines);
        text["autofit"] = "shrink";
        e["text"] = text;
        return this;
    }

    public SlideBuilder Notes(string text)
    {
        slide["notes"] = text;
        return this;
    }

    /// <summary>A shape with text; answers its id.</summary>
    public string Shape(string shape, double x, double y, double w, double h, string fill, string? text = null, double size = 18, string textColor = "lt1", bool bold = false)
    {
        var id = NewId();
        var e = new JsonObject
        {
            ["id"] = id, ["type"] = "shape", ["shape"] = shape, ["x"] = x, ["y"] = y, ["w"] = w, ["h"] = h,
            ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = fill },
        };
        if (text is not null)
        {
            var body = Deck.Text(text.Split('\n'), size, textColor, bold, "center");
            body["anchor"] = "middle";
            e["text"] = body;
        }

        Elements.Add(e);
        return id;
    }

    /// <summary>A card with a big number and a label under it.</summary>
    public string Card(double x, double y, double w, double h, string accent, string number, string label)
    {
        var id = NewId();
        var body = new JsonObject
        {
            ["paras"] = new JsonArray(
                new JsonObject { ["align"] = "center", ["runs"] = new JsonArray(new JsonObject { ["t"] = number, ["size"] = 40, ["b"] = true, ["color"] = accent + ":-25" }) },
                new JsonObject { ["align"] = "center", ["spaceBefore"] = 6, ["runs"] = new JsonArray(new JsonObject { ["t"] = label, ["size"] = 18, ["color"] = "dk2" }) }),
            ["anchor"] = "middle",
        };
        Elements.Add(new JsonObject
        {
            ["id"] = id, ["type"] = "shape", ["shape"] = "roundRect", ["adj"] = 0.08, ["x"] = x, ["y"] = y, ["w"] = w, ["h"] = h,
            ["fill"] = new JsonObject { ["type"] = "solid", ["color"] = accent + ":+80" },
            ["line"] = new JsonObject { ["color"] = accent, ["width"] = 1.5, ["dash"] = "solid" },
            ["text"] = body,
        });
        return id;
    }

    public string TextBox(double x, double y, double w, double h, string text, double size, string color, bool bold = false, string align = "left")
    {
        var id = NewId();
        var body = Deck.Text(text.Split('\n'), size, color, bold, align);
        body["autofit"] = "grow";
        Elements.Add(new JsonObject { ["id"] = id, ["type"] = "text", ["x"] = x, ["y"] = y, ["w"] = w, ["h"] = h, ["text"] = body });
        return id;
    }

    /// <summary>A table from rows of texts: the first row is the header.</summary>
    public JsonObject Table(double x, double y, double w, double rowHeight, string[][] cells, string accent, double[]? widths = null, bool lastRow = false)
    {
        var cols = cells[0].Length;
        var colWidths = widths ?? Enumerable.Repeat(Math.Round(w / cols, 2), cols).ToArray();
        var rows = new JsonArray();
        foreach (var row in cells)
        {
            var list = new JsonArray();
            foreach (var text in row)
            {
                list.Add(new JsonObject { ["text"] = Deck.Text([text]) });
            }

            rows.Add(new JsonObject { ["h"] = rowHeight, ["cells"] = list });
        }

        var e = new JsonObject
        {
            ["id"] = NewId(), ["type"] = "table", ["x"] = x, ["y"] = y, ["w"] = colWidths.Sum(), ["h"] = rowHeight * cells.Length,
            ["cols"] = new JsonArray(colWidths.Select(c => (JsonNode)c).ToArray()), ["rows"] = rows,
            ["tableStyle"] = "medium", ["accent"] = accent, ["header"] = true, ["banded"] = true, ["firstCol"] = false, ["lastRow"] = lastRow,
        };
        Elements.Add(e);
        return e;
    }

    public SlideBuilder Animate(string elementId, string effect, string trigger, string? dir = null)
    {
        if (slide["animations"] is not JsonArray list)
        {
            list = [];
            slide["animations"] = list;
        }

        var a = new JsonObject { ["el"] = elementId, ["effect"] = effect, ["trigger"] = trigger, ["duration"] = 0.5 };
        if (dir is not null)
        {
            a["dir"] = dir;
        }

        list.Add(a);
        return this;
    }
}
