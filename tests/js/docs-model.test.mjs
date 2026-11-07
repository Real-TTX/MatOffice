// The model of MatOffice Docs without a browser: node --test tests/js/
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  normalizePage, pageGeometry, sizeName, styleKey, withStyle, levelFormat, marker, alpha, roman, listNumbers, fontStack, firstFamily, stepSize, colour,
  parseStyle, toPt, borderCss, blockCss, inlineFormat, escText, escAttr, cleanText, cleanFormat, sameFormat, normalizeRuns, splitRuns, sliceRuns,
  deleteRuns, insertRuns, formatRuns, formatAt, runsText, rangeHas, rangeValue, splitBlock, mergeBlocks, textBlocks, writeRuns, writeBlocks, newTable,
  tableGrid, columnCount, changeCase, countWords, statistics, findAll, newId, tocBlocks, fieldText, safeHref, linkFromInput, safeImage, isId, isItemId,
  textRun, runsLength,
} from '../../src/MatOffice/wwwroot/js/docs/model.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const p = (runs, props = {}) => ({ type: 'p', css: {}, runs, ...props });

test('the page follows the rules of the server', () => {
  const page = normalizePage({});
  assert.deepEqual(page.margins, { top: 25, right: 25, bottom: 20, left: 25 });
  assert.equal(page.size, 'A4');
  const letter = normalizePage({ size: 'Letter', width: 1, height: 1, orientation: 'landscape' });
  assert.deepEqual([letter.width, letter.height, letter.orientation], [215.9, 279.4, 'landscape']);
  const custom = normalizePage({ size: 'Custom', width: 5, height: 5000 });
  assert.deepEqual([custom.width, custom.height], [50, 1000]);
  const wide = normalizePage({ margins: { left: 150, right: 150, top: -3 } });
  assert.ok(wide.margins.left + wide.margins.right <= 210 - 20);
  assert.equal(wide.margins.top, 0);
  assert.equal(normalizePage({ size: 'B5' }).size, 'A4');
  assert.equal(normalizePage({ columns: 7 }).columns, 3);
  const g = pageGeometry(normalizePage({ orientation: 'landscape' }));
  assert.ok(g.paperWidth > g.paperHeight);
  assert.ok(Math.abs(g.textWidth - (297 - 50) * 96 / 25.4) < 0.01);
  assert.equal(sizeName(215, 280), 'Letter');
  assert.equal(sizeName(100, 100), 'Custom');
});

test('styles map to blocks and back', () => {
  assert.equal(styleKey({ type: 'h2' }), 'Heading2');
  assert.equal(styleKey({ type: 'blockquote', style: 'IntenseQuote' }), 'IntenseQuote');
  assert.equal(styleKey({ type: 'p', style: 'Bogus' }), 'Normal');
  const h = withStyle(p([textRun('x')], { list: 'bullet', level: 1 }), 'Heading1');
  assert.equal(h.type, 'h1');
  assert.equal(h.list, 'bullet');
  const q = withStyle(p([], { list: 'decimal' }), 'Quote');
  assert.equal(q.type, 'blockquote');
  assert.equal(q.list, undefined);
  assert.equal(withStyle({ type: 'h1', id: 'h-a', css: {}, runs: [] }, 'Normal').id, undefined);
});

test('list markers and numbers like the server', () => {
  assert.equal(levelFormat('bullet', 4), 'circle');
  assert.equal(levelFormat('decimal', 2), 'lower-roman');
  assert.equal(marker('decimal', 1, 3), 'c.');
  assert.equal(marker('outline', 2, 1, [2, 1, 1]), '2.1.1.');
  assert.equal(alpha(28), 'ab');
  assert.equal(roman(1994), 'MCMXCIV');
  const blocks = [
    p([], { list: 'decimal', level: 0 }), p([], { list: 'decimal', level: 1 }), p([], { list: 'decimal', level: 0 }),
    p([]), p([], { list: 'decimal', level: 0, start: 5 }), p([], { list: 'decimal', level: 0 }),
    p([], { list: 'bullet', level: 0 }), p([], { list: 'decimal', level: 0 }),
  ];
  const n = listNumbers(blocks);
  assert.deepEqual(blocks.map(b => n.get(b)?.n ?? null), [1, 1, 2, null, 5, 6, 1, 1]);
  assert.deepEqual(n.get(blocks[1]).path, [1, 1]);
});

test('fonts get fallbacks, sizes step like Word', () => {
  assert.equal(fontStack('Calibri'), 'Calibri, Carlito, sans-serif');
  assert.equal(fontStack('Times New Roman'), "'Times New Roman', 'Liberation Serif', Times, serif");
  assert.equal(fontStack('x"; color: red'), "'x color: red', sans-serif");
  assert.equal(firstFamily("'Segoe UI', Arial"), 'Segoe UI');
  assert.equal(stepSize(11, 1), 12);
  assert.equal(stepSize(11, -1), 10.5);
  assert.equal(stepSize(72, 1), 82);
});

test('colours and CSS are read safely', () => {
  assert.equal(colour('#ABC'), '#aabbcc');
  assert.equal(colour('rgb(255, 0, 128)'), '#ff0080');
  assert.equal(colour('rgba(0,0,0,0)'), null);
  assert.equal(colour('windowtext'), '#000000');
  assert.equal(colour('url(x)'), null);
  assert.deepEqual(parseStyle('color: red; background: url(http://x/y.png); --x: 1; font-family: "A;B", serif'), [['color', 'red'], ['font-family', '"A;B", serif']]);
  assert.equal(toPt('16px'), 12);
  assert.equal(toPt('1in'), 72);
  assert.equal(borderCss('1px solid red'), '0.75pt solid #ff0000');
  assert.equal(borderCss('none'), null);
  assert.deepEqual(blockCss(parseStyle('text-align: start; line-height: 150%; margin: 6pt 0 12pt 1in; --mo-sb: 18pt; background: yellow')),
    { 'text-align': 'left', 'line-height': '1.5', 'margin-top': '18pt', 'margin-bottom': '12pt', 'margin-left': '72pt', 'margin-right': '0pt', 'background-color': '#ffff00' });
  assert.deepEqual(inlineFormat(parseStyle('font-weight: 700; font-style: italic; text-decoration: underline line-through; font-size: 16px; vertical-align: super')),
    { b: true, i: true, u: true, s: true, size: '12pt', sup: true });
  assert.deepEqual(inlineFormat(parseStyle('font-weight: normal')), { b: false });
});

test('text is escaped and cleaned', () => {
  assert.equal(escText('a<b>&"'), 'a&lt;b&gt;&amp;"');
  assert.equal(escAttr('a"&<'), 'a&quot;&amp;<');
  assert.equal(cleanText('a\r\nb\u0000c\u0007\td'), 'a\nbc\td');
});

test('runs: split, slice, insert, delete, format', () => {
  const runs = [textRun('Hello '), textRun('bold', { b: true }), textRun(' world')];
  assert.equal(runsLength(runs), 16);
  const [a, b] = splitRuns(runs, 8);
  assert.equal(runsText(a), 'Hello bo');
  assert.equal(runsText(b), 'ld world');
  assert.equal(runsText(sliceRuns(runs, 3, 12)), 'lo bold w');
  assert.equal(runsText(deleteRuns(runs, 6, 11)), 'Hello world');
  assert.equal(deleteRuns(runs, 6, 11).length, 1, 'neighbours with the same format merge');
  const inserted = insertRuns(runs, 6, [textRun('very ', { b: true })]);
  assert.equal(inserted[1].text, 'very bold');
  const formatted = formatRuns(runs, 0, 5, { i: true });
  assert.deepEqual(formatted[0], { text: 'Hello', f: { i: true } });
  assert.deepEqual(formatRuns(runs, 6, 10, f => ({ ...f, b: false })).length, 1);
  assert.ok(rangeHas(runs, 6, 10, f => f.b));
  assert.ok(!rangeHas(runs, 5, 10, f => f.b));
  assert.equal(rangeValue([textRun('a', { size: '12pt' }), textRun('b', { size: '14pt' })], 0, 2, 'size'), '');
  assert.equal(normalizeRuns([textRun(''), { obj: 'br', f: {} }, textRun('x'), textRun('y')]).length, 2);
  assert.ok(sameFormat({ b: true }, cleanFormat({ b: true, i: false })));
});

test('typing continues the format before the caret, but not a link or a comment at its end', () => {
  const runs = [textRun('go '), textRun('here', { href: 'https://example.com/', b: true }), textRun('!')];
  assert.deepEqual(formatAt(runs, 5), { href: 'https://example.com/', b: true });
  assert.deepEqual(formatAt(runs, 7), { b: true });
  assert.deepEqual(formatAt(runs, 3), {});
  assert.deepEqual(formatAt([], 0), {});
  assert.deepEqual(formatAt([textRun('x', { cmt: 'c1', i: true })], 1), { i: true });
});

test('blocks split and merge', () => {
  const h = { type: 'h1', id: 'h-x', css: { 'text-align': 'center' }, runs: [textRun('Title text')] };
  const [first, second] = splitBlock(h, 5);
  assert.equal(first.id, 'h-x');
  assert.equal(second.id, undefined);
  assert.equal(second.css['text-align'], 'center');
  assert.equal(runsText(mergeBlocks(first, second).runs), 'Title text');
  const t = newTable(2, 2);
  t.rows[1].cells[1].blocks = [p([textRun('in a cell')])];
  assert.equal([...textBlocks([p([]), t])].length, 5);
});

test('the writer produces the canonical Docs HTML of the server', () => {
  const runs = [
    textRun('a', { href: 'https://example.com/?a=1&b=2', cmt: 'c1', font: 'Arial, sans-serif', size: '14pt', color: '#ff0000', b: true, i: true }),
    textRun('<b>', { href: 'https://example.com/?a=1&b=2', b: true }),
    { obj: 'br', f: {} }, { obj: 'img', src: PNG, alt: 'A "dot"', width: 10, height: 20.5, align: 'right', f: { b: true } },
    { obj: 'note', id: 'n1', f: {} }, { obj: 'field', kind: 'page', f: { size: '9pt' } },
  ];
  assert.equal(writeRuns(runs),
    '<a href="https://example.com/?a=1&amp;b=2"><span data-comment="c1"><span style="font-family: Arial, sans-serif; font-size: 14pt; color: #ff0000"><strong><em>a</em></strong></span></span><strong>&lt;b&gt;</strong></a>'
    + '<br><img src="' + PNG + '" alt="A &quot;dot&quot;" style="width: 10px; height: 20.5px" data-align="right"><sup data-fn="n1"></sup><span style="font-size: 9pt"><span data-field="page"></span></span>');
  const blocks = [
    { type: 'h2', id: 'h-a', css: { 'text-align': 'center', 'margin-top': '6pt' }, runs: [textRun('H')] },
    { type: 'p', style: 'Title', css: {}, runs: [] },
    { type: 'p', list: 'decimal', level: 2, start: 3, css: {}, runs: [textRun('i')] },
    { type: 'pre', css: {}, runs: [textRun('\nfirst line kept')] },
    { type: 'hr' }, { type: 'break', kind: 'page' },
  ];
  assert.equal(writeBlocks(blocks),
    '<h2 id="h-a" style="text-align: center; margin-top: 6pt">H</h2><p data-style="Title"></p><p data-list="decimal" data-level="2" data-start="3">i</p>'
    + '<pre>\n\nfirst line kept</pre><hr><div data-break="page"></div>');
  const t = newTable(2, 3, { header: true, style: 'accent1' });
  t.rows[1].cells[0].colspan = 2;
  t.rows[1].cells.splice(1, 1);
  t.rows[1].cells[1].css = { 'background-color': '#ff0000' };
  assert.equal(writeBlocks([t]),
    '<table data-table-style="accent1" style="width: 100%"><colgroup><col style="width: 33.33%"><col style="width: 33.33%"><col style="width: 33.33%"></colgroup>'
    + '<thead><tr><th><p></p></th><th><p></p></th><th><p></p></th></tr></thead>'
    + '<tbody><tr><td colspan="2"><p></p></td><td style="background-color: #ff0000"><p></p></td></tr></tbody></table>');
  assert.equal(writeRuns([textRun('x')], { image: () => 'mo-img:1' }), 'x');
});

test('the grid of a table knows merged cells', () => {
  const t = newTable(3, 3);
  t.rows[0].cells[0].rowspan = 2;
  t.rows[1].cells.splice(0, 1);
  t.rows[2].cells[1].colspan = 2;
  t.rows[2].cells.splice(2, 1);
  const grid = tableGrid(t);
  assert.equal(grid[1][0].cell, t.rows[0].cells[0]);
  assert.equal(grid[1][1].cell, t.rows[1].cells[0]);
  assert.equal(grid[2][2].cell, t.rows[2].cells[1]);
  assert.equal(columnCount(t), 3);
});

test('case, words, search', () => {
  assert.equal(changeCase('hello wORLD. next one', 'sentence', 'en'), 'Hello world. Next one');
  assert.equal(changeCase('hello world', 'title', 'en'), 'Hello World');
  assert.equal(changeCase('Hello', 'toggle', 'en'), 'hELLO');
  assert.equal(changeCase('straße', 'upper', 'de'), 'STRASSE');
  assert.equal(countWords("It's a well-known e-mail test, 42 times."), 7);
  assert.deepEqual(statistics(['Hello world', '', 'x￼']), { words: 3, chars: 12, charsNoSpaces: 11, paragraphs: 2 });
  assert.deepEqual(findAll('The cat and the Category', 'cat'), [[4, 7], [16, 19]]);
  assert.deepEqual(findAll('The cat and the Category', 'cat', { wholeWord: true }), [[4, 7]]);
  assert.deepEqual(findAll('The Cat', 'cat', { matchCase: true }), []);
  assert.deepEqual(findAll('a.b a*b', '.'), [[1, 2]]);
  assert.deepEqual(findAll('Grüße grüße', 'grüße', { wholeWord: true }), [[0, 5], [6, 11]]);
});

test('links, pictures and ids are checked', () => {
  assert.equal(safeHref('javascript:alert(1)'), null);
  assert.equal(safeHref(' https://example.com '), 'https://example.com/');
  assert.equal(safeHref('#h-intro'), '#h-intro');
  assert.equal(safeHref('data:text/html,x'), null);
  assert.equal(linkFromInput('example.com/x'), 'https://example.com/x');
  assert.equal(linkFromInput('ada@example.com'), 'mailto:ada@example.com');
  assert.equal(linkFromInput('not a link'), null);
  assert.equal(safeImage(PNG), PNG);
  assert.equal(safeImage('data:image/svg+xml;base64,PHN2Zz4='), null);
  assert.equal(safeImage('https://example.com/x.png'), null);
  assert.ok(isId('h-abc_1') && !isId('1abc') && !isId('a b'));
  assert.ok(isItemId('n1') && !isItemId('n 1'));
  let i = 0;
  assert.equal(newId('h-', () => (i++ % 10) / 10), 'h-adgjor');
});

test('a table of contents and fields', () => {
  const toc = tocBlocks([{ level: 1, text: ' Intro\n ', id: 'h-a' }, { level: 2, text: 'Part', id: 'h-b' }, { level: 4, text: 'deep', id: 'h-c' }], 'Contents');
  assert.equal(writeBlocks(toc), '<p data-style="TOCHeading">Contents</p><p data-style="TOC1"><a href="#h-a">Intro</a></p><p data-style="TOC2"><a href="#h-b">Part</a></p>');
  assert.equal(fieldText('page', { page: 3 }), '3');
  assert.equal(fieldText('title', { name: 'Report' }), 'Report');
  assert.equal(fieldText('date', { lang: 'de', now: new Date(2026, 10, 7) }), '07.11.2026');
});
