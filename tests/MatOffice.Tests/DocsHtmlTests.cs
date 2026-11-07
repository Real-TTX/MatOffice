using MatOffice.Services.Formats.Docs;

namespace MatOffice.Tests;

/// <summary>The sanitizer and the canonicalizer of Docs HTML.</summary>
public sealed class DocsHtmlTests
{
    private const string Png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    [Theory]
    [InlineData("<p>a<script>alert(1)</script>b</p>", "<p>ab</p>")]
    [InlineData("<p onclick=\"x()\">a<img src=\"" + Png + "\" onerror=\"alert(1)\"></p>", "<p>a<img src=\"" + Png + "\" alt=\"\"></p>")]
    [InlineData("<p><a href=\"javascript:alert(1)\">x</a></p>", "<p>x</p>")]
    [InlineData("<p><a href=\" JaVaScRiPt:alert(1)\">x</a></p>", "<p>x</p>")]
    [InlineData("<p><a href=\"data:text/html;base64,PHNjcmlwdD4=\">x</a></p>", "<p>x</p>")]
    [InlineData("<p><a href=\"vbscript:x\">x</a><a href=\"/relative\">y</a></p>", "<p>xy</p>")]
    [InlineData("<p><img src=\"http://evil.example/x.png\">a</p>", "<p>a</p>")]
    [InlineData("<p><img src=\"data:image/svg+xml;base64,PHN2Zz4=\">a</p>", "<p>a</p>")]
    [InlineData("<p style=\"background:url(http://x/y.png); color: red\">a</p>", "<p><span style=\"color: #ff0000\">a</span></p>")]
    [InlineData("<p style=\"margin-top: expression(alert(1)); text-align: center\">a</p>", "<p style=\"text-align: center\">a</p>")]
    [InlineData("<p style=\"--x: 1; color: var(--x); text-align: right\">a</p>", "<p style=\"text-align: right\">a</p>")]
    [InlineData("<iframe src=\"https://x\"></iframe><object data=\"x\"></object><embed src=\"x\"><form><input value=\"v\"><button>b</button></form><p>a</p>", "<p>a</p>")]
    [InlineData("<p class=\"x\" data-x=\"1\" data-style=\"Bogus\" title=\"t\">a</p>", "<p>a</p>")]
    [InlineData("<svg><script>alert(1)</script></svg><math><mi>x</mi></math><p>a</p>", "<p>a</p>")]
    [InlineData("<p>a<style>p{}</style><link rel=\"stylesheet\" href=\"x\"><meta http-equiv=\"refresh\"><base href=\"http://x\">b</p>", "<p>ab</p>")]
    public void The_sanitizer_removes_what_could_run_or_load(string html, string expected) => Assert.Equal(expected, DocsHtml.Clean(html));

    [Fact]
    public void The_sanitizer_keeps_the_formatting_of_the_editor()
    {
        const string html = "<h1 id=\"h-intro\" style=\"text-align: center\">Intro</h1>"
            + "<p data-style=\"Title\">T</p>"
            + "<p data-list=\"decimal\" data-level=\"2\" data-start=\"5\">i</p>"
            + "<p><a href=\"https://example.com/x?a=1&amp;b=2\"><span data-comment=\"c1 c2\"><span style=\"font-family: Arial; font-size: 14pt; color: #ff0000; background-color: #ffff00\"><strong><em><u><s><sup>x</sup></s></u></em></strong></span></span></a></p>"
            + "<p><a href=\"#h-intro\">link</a> <a href=\"mailto:a@b.c\">m</a><sup data-fn=\"n1\"></sup><span data-field=\"page\"></span><br>z</p>"
            + "<blockquote data-style=\"IntenseQuote\">q</blockquote><pre>a\n  b</pre><hr><div data-break=\"page\"></div>"
            + "<table data-table-style=\"accent1\" style=\"width: 100%\"><colgroup><col style=\"width: 30%\"><col style=\"width: 70%\"></colgroup>"
            + "<thead><tr><th><p>A</p></th><th><p>B</p></th></tr></thead><tbody><tr><td colspan=\"2\" style=\"background-color: #ff0000; vertical-align: middle\"><p>c</p></td></tr></tbody></table>"
            + "<p><img src=\"" + Png + "\" alt=\"pic\" style=\"width: 10px; height: 20px\" data-align=\"right\"></p>";
        Assert.Equal(html, DocsHtml.Clean(html));
    }

    [Fact]
    public void The_canonical_form_is_stable()
    {
        const string html = "<div><b>a</b><i>b</i></div><ul><li>one<ul><li>two</li></ul></li><li><p>three</p><p>more</p></li></ul>"
            + "<ol type=\"a\" start=\"3\"><li>x</li></ol><font color=\"red\" face=\"Arial\">f</font><section><p>s</p></section>";
        var once = DocsHtml.Clean(html);
        Assert.Equal(once, DocsHtml.Clean(once));
        Assert.Equal(
            "<p><strong>a</strong><em>b</em></p><p data-list=\"bullet\">one</p><p data-list=\"bullet\" data-level=\"1\">two</p>"
            + "<p data-list=\"bullet\">three</p><p data-style=\"ListParagraph\">more</p><p data-list=\"lower-alpha\" data-start=\"3\">x</p>"
            + "<p><span style=\"font-family: Arial, Liberation Sans, sans-serif; color: #ff0000\">f</span></p><p>s</p>",
            once);
    }

    [Theory]
    [InlineData("", "")]
    [InlineData("   \n ", "")]
    [InlineData("plain", "<p>plain</p>")]
    [InlineData("<h2>a<div>b</div>c</h2>", "<h2>a</h2><h2>b</h2><h2>c</h2>")]
    [InlineData("<h1>a<p>b</p></h1>", "<h1>a</h1><h1>b</h1>")]
    [InlineData("<blockquote><p>a</p><p>b</p></blockquote>", "<blockquote>a</blockquote><blockquote>b</blockquote>")]
    [InlineData("<p></p><p> </p>", "<p></p><p> </p>")]
    [InlineData("<table><tr><td>a<p>b</p></td><td></td></tr></table>", "<table data-table-style=\"grid\"><tbody><tr><td><p>a</p><p>b</p></td><td><p></p></td></tr></tbody></table>")]
    [InlineData("<strong>a<p>b</p></strong>", "<p><strong>a</strong></p><p><strong>b</strong></p>")]
    [InlineData("<p><span style=\"font-weight: bold\">a</span><span style=\"font-style: italic; text-decoration: underline line-through\">b</span></p>", "<p><strong>a</strong><em><u><s>b</s></u></em></p>")]
    [InlineData("<p><b style=\"font-weight: normal\">a</b></p>", "<p>a</p>")]
    [InlineData("<p>a</p>\n  <p>b</p>", "<p>a</p><p>b</p>")]
    [InlineData("<p><em><strong>x</strong></em><strong>y</strong></p>", "<p><strong><em>x</em>y</strong></p>")]
    [InlineData("<p><span data-comment=\"c1\"><a href=\"https://x.org\">a</a></span></p>", "<p><a href=\"https://x.org\"><span data-comment=\"c1\">a</span></a></p>")]
    [InlineData("<p><sup data-fn=\"n1\">1</sup><span data-field=\"pages\">7</span></p>", "<p><sup data-fn=\"n1\"></sup><span data-field=\"pages\"></span></p>")]
    [InlineData("<p id=\"1bad\">a</p><p id=\"ok\">b</p><p id=\"ok\">c</p>", "<p>a</p><p id=\"ok\">b</p><p>c</p>")]
    [InlineData("<p style=\"margin: 0 0 8pt; border: 1px solid red; line-height: 150%\">a</p>", "<p style=\"margin-top: 0pt; margin-right: 0pt; margin-bottom: 8pt; margin-left: 0pt; border-top: 0.75pt solid #ff0000; border-right: 0.75pt solid #ff0000; border-bottom: 0.75pt solid #ff0000; border-left: 0.75pt solid #ff0000; line-height: 1.5\">a</p>")]
    public void The_canonicalizer_makes_a_flat_block_list(string html, string expected) => Assert.Equal(expected, DocsHtml.Clean(html));

    [Fact]
    public void Deep_nesting_does_not_overflow()
    {
        var html = string.Concat(Enumerable.Repeat("<span>", 50_000)) + "x" + string.Concat(Enumerable.Repeat("</span>", 50_000));
        Assert.Equal("<p>x</p>", DocsHtml.Clean(html));
    }

    [Fact]
    public void Large_pictures_are_fast()
    {
        var data = Convert.ToBase64String(new byte[6_000_000]);
        var html = string.Concat(Enumerable.Range(0, 200).Select(i => $"<p>Paragraph {i} <strong>bold</strong> <span style=\"color: #ff0000\">red</span></p>"))
                   + $"<p><img src=\"data:image/png;base64,{data}\" alt=\"big\"></p>";
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var clean = DocsHtml.Clean(html);
        watch.Stop();
        Assert.Contains(data, clean, StringComparison.Ordinal);
        Assert.True(watch.ElapsedMilliseconds < 3000, $"{watch.ElapsedMilliseconds} ms");
    }
}
