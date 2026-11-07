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

    [Theory]
    [InlineData("<p><a href=\"javascript&#58;alert(1)\">x</a><a href=\"jav&#x09;ascript:alert(1)\">y</a></p>", "<p>xy</p>")]
    [InlineData("<p><a href=\"#1bad\">a</a><a href=\"http://\">b</a><a href=\"ftp://example.com/f\">c</a><a href=\"//example.com\">d</a></p>", "<p>abcd</p>")]
    [InlineData("<p><a href=\"HTTPS://Example.com/Path\" target=\"_blank\" rel=\"noopener\" onclick=\"x()\">a</a></p>", "<p><a href=\"https://Example.com/Path\">a</a></p>")]
    [InlineData("<p><a href=\"mailto:a@example.com?subject=Hi\">m</a></p>", "<p><a href=\"mailto:a@example.com?subject=Hi\">m</a></p>")]
    [InlineData("<p><img src=\"data:image/png;base64,!!!\">a</p>", "<p>a</p>")]
    [InlineData("<p><img src=\"data:image/png;base64,\">a</p>", "<p>a</p>")]
    [InlineData("<p><img src=\"data:image/jpg;base64,AAAA\"></p>", "<p><img src=\"data:image/jpeg;base64,AAAA\" alt=\"\"></p>")]
    [InlineData("<p><img src=\"DATA:IMAGE/PNG;BASE64,AA AA\n\"></p>", "<p><img src=\"data:image/png;base64,AAAA\" alt=\"\"></p>")]
    [InlineData("<p><img src=\"data:image/gif;base64,AAAA\" width=\"20\" height=\"10\" onload=\"x()\" data-align=\"middle\"></p>", "<p><img src=\"data:image/gif;base64,AAAA\" alt=\"\" style=\"width: 20px; height: 10px\"></p>")]
    [InlineData("<video src=\"x.mp4\"><source src=\"y\"></video><audio src=\"z\"></audio><template><p>t</p></template><noscript><p>n</p></noscript><textarea>z</textarea><p>a</p>", "<p>a</p>")]
    [InlineData("<p>a<!-- comment -->b<custom-element>c</custom-element></p>", "<p>abc</p>")]
    [InlineData("<p style=\"behavior: url(x.htc); -moz-binding: url(x); color: red\">a</p>", "<p><span style=\"color: #ff0000\">a</span></p>")]
    [InlineData("<p style=\"color: re\\64; text-align: center\">a</p>", "<p style=\"text-align: center\">a</p>")]
    [InlineData("<p style=\"color: red /* x */; text-align: right\">a</p>", "<p style=\"text-align: right\">a</p>")]
    [InlineData("<p style=\"color: javascript:alert(1); font-size: 12pt\">a</p>", "<p><span style=\"font-size: 12pt\">a</span></p>")]
    [InlineData("<p style=\"@import 'x.css'; color: blue\">a</p>", "<p><span style=\"color: #0000ff\">a</span></p>")]
    [InlineData("<div style=\"background-image: url(x.png)\"><p>a</p></div>", "<p>a</p>")]
    [InlineData("<p style=\"margin-left: 99999pt; text-indent: 12pt\">a</p>", "<p style=\"text-indent: 12pt\">a</p>")]
    public void The_sanitizer_removes_tricks(string html, string expected) => Assert.Equal(expected, DocsHtml.Clean(html));

    [Theory]
    [InlineData("<p data-list=\"decimal\" data-level=\"99\" data-start=\"0\">a</p>", "<p data-list=\"decimal\" data-level=\"8\">a</p>")]
    [InlineData("<p data-list=\"decimal\" data-level=\"-1\" data-start=\"100000\">a</p>", "<p data-list=\"decimal\">a</p>")]
    [InlineData("<p data-list=\"decimal\" data-level=\"x\" data-start=\"abc\">a</p>", "<p data-list=\"decimal\">a</p>")]
    [InlineData("<p data-list=\"bogus\" data-level=\"2\">a</p>", "<p>a</p>")]
    [InlineData("<p data-level=\"2\" data-start=\"4\">a</p>", "<p>a</p>")]
    [InlineData("<div data-break=\"line\"><p>a</p></div><div data-break=\"column\"></div>", "<p>a</p><div data-break=\"column\"></div>")]
    [InlineData("<table data-table-style=\"fancy\"><tr><td colspan=\"0\" rowspan=\"-3\">a</td><td colspan=\"5000\">b</td></tr></table>",
        "<table data-table-style=\"grid\"><tbody><tr><td><p>a</p></td><td colspan=\"1000\"><p>b</p></td></tr></tbody></table>")]
    [InlineData("<table><tr><td style=\"vertical-align: top; background-color: rgb(255, 0, 0)\">a</td><td style=\"vertical-align: bottom\">b</td></tr></table>",
        "<table data-table-style=\"grid\"><tbody><tr><td style=\"background-color: #ff0000\"><p>a</p></td><td style=\"vertical-align: bottom\"><p>b</p></td></tr></tbody></table>")]
    [InlineData("<p><span data-comment=\"c1 bad! c2 c1\">a</span><sup data-fn=\"bad id!\">1</sup><span data-field=\"evil\">x</span></p>", "<p><span data-comment=\"c1 c2\">a</span></p>")]
    [InlineData("<p data-style=\"IntenseQuote\">q</p><h1 data-style=\"Title\">h</h1><blockquote data-style=\"Title\">b</blockquote><pre data-style=\"Caption\">c</pre>",
        "<blockquote data-style=\"IntenseQuote\">q</blockquote><h1>h</h1><blockquote>b</blockquote><pre>c</pre>")]
    [InlineData("<ol style=\"list-style-type: upper-roman\"><li>a</li></ol><ol><li value=\"5\">b</li></ol><ul style=\"list-style-type: square\"><li>c</li></ul>",
        "<p data-list=\"upper-roman\">a</p><p data-list=\"decimal\" data-start=\"5\">b</p><p data-list=\"square\">c</p>")]
    [InlineData("<P STYLE=\"TEXT-ALIGN: CENTER\">a</P>", "<p style=\"text-align: center\">a</p>")]
    [InlineData("<p id=\"x0123456789012345678901234567890123456789012345678901234567890123456789\">long id</p>", "<p>long id</p>")]
    public void The_sanitizer_checks_the_values_of_the_docs_attributes(string html, string expected) => Assert.Equal(expected, DocsHtml.Clean(html));

    [Theory]
    [InlineData("color: rgb(255, 0, 0)", "color: #ff0000")]
    [InlineData("color: rgba(0, 0, 0, 0)", "")]
    [InlineData("color: hsl(120, 100%, 25%)", "color: #008000")]
    [InlineData("color: #ABC", "color: #aabbcc")]
    [InlineData("color: Red", "color: #ff0000")]
    [InlineData("color: transparent", "")]
    [InlineData("font-size: 16px", "font-size: 12pt")]
    [InlineData("font-size: large", "font-size: 13.5pt")]
    [InlineData("font-size: 150%", "font-size: 16.5pt")]
    [InlineData("font-family: \"Comic Sans MS\", cursive", "font-family: &quot;Comic Sans MS&quot;, cursive")]
    [InlineData("background: yellow", "background-color: #ffff00")]
    [InlineData("background: url(x.png) yellow", "")]
    [InlineData("letter-spacing: 1px", "letter-spacing: 0.75pt")]
    public void The_sanitizer_writes_character_styles_canonically(string style, string expected)
        => Assert.Equal(expected.Length == 0 ? "<p>a</p>" : $"<p><span style=\"{expected}\">a</span></p>", DocsHtml.Clean($"<p><span style='{style}'>a</span></p>"));

    [Fact]
    public void Text_is_escaped_and_cleaned()
    {
        Assert.Equal("<p>a &lt;b&gt; &amp; \"c\" 'd'</p>", DocsHtml.Clean("<p>a &lt;b&gt; &amp; \"c\" 'd'</p>"));
        Assert.Equal("<p>ab</p>", DocsHtml.Clean("<p>a\u0001b</p>"));
        Assert.Equal("<p>a b\tc</p>", DocsHtml.Clean("<p>a&nbsp;b\tc</p>"));
        Assert.Equal("<p>x�y</p>", DocsHtml.Clean("<p>x" + '\uD800' + "y</p>"));
        Assert.Equal("<table data-table-style=\"grid\"><tbody><tr><td><p>a</p><table data-table-style=\"plain\"><tbody><tr><td><p>inner</p></td></tr></tbody></table></td></tr></tbody></table>",
            DocsHtml.Clean("<table><tr><td>a<table data-table-style=\"plain\"><tr><td>inner</td></tr></table></td></tr></table>"));
    }

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
    [InlineData("<p style=\"margin: 0 0 8pt; border: 1px solid red; line-height: 150%\">a</p>", "<p style=\"line-height: 1.5; margin-top: 0pt; margin-bottom: 8pt; margin-left: 0pt; margin-right: 0pt; border-top: 0.75pt solid #ff0000; border-right: 0.75pt solid #ff0000; border-bottom: 0.75pt solid #ff0000; border-left: 0.75pt solid #ff0000\">a</p>")]
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
