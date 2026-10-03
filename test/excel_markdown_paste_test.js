const assert = require('assert');
const api = require('../assets/javascripts/excel_markdown_paste.js');

function excelHtml(inner, head) {
  return '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta name="ProgId" content="Excel.Sheet">' +
    (head || '') + '</head><body>' + inner + '</body></html>';
}

const cases = [];
function test(name, fn) {
  cases.push([name, fn]);
}

test('converts a rectangular Excel range to a markdown table', function () {
  const html = excelHtml(
    '<table><tbody><tr><td>名前</td><td>数量</td></tr><tr><td>りんご</td><td>3</td></tr></tbody></table>'
  );
  assert.strictEqual(api.convertClipboard(html, '名前\t数量\nりんご\t3\n'),
    '| 名前 | 数量 |\n| --- | --- |\n| りんご | 3 |');
});

test('converts a single Excel column when the clipboard is spreadsheet HTML', function () {
  const html = excelHtml('<table><tr><td>りんご</td></tr><tr><td>みかん</td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, 'りんご\nみかん\n'),
    '| りんご |\n| --- |\n| みかん |');
});

test('escapes pipes and keeps in-cell line breaks', function () {
  const html = excelHtml('<table><tr><td>a|b</td><td>line1<br>line2</td></tr><tr><td>A &amp; B</td><td>&lt;tag&gt;</td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, 'a|b\tline1\nline2\nA & B\t<tag>\n'),
    '| a\\|b | line1<br>line2 |\n| --- | --- |\n| A &amp; B | &lt;tag&gt; |');
});

test('expands colspan and rowspan without duplicating the value', function () {
  const html = excelHtml(
    '<table><tr><td colspan="2">見出し</td></tr><tr><td rowspan="2">A</td><td>b1</td></tr><tr><td>b2</td></tr></table>'
  );
  assert.strictEqual(api.convertClipboard(html, ''),
    '| 見出し |  |\n| --- | --- |\n| A | b1 |\n|  | b2 |');
});

test('leaves image markup to Redmine and does not build a table from it', function () {
  assert.strictEqual(api.convertClipboard('<img src="shot.png" alt="screen">', ''), null);
  assert.strictEqual(api.convertClipboard('<table><tr><td><img src="shot.png"></td></tr></table>', ''), null);
  assert.strictEqual(api.shouldIgnorePaste(true, '<img src="shot.png">', ''), true);
  assert.strictEqual(api.shouldIgnorePaste(false, '<img src="shot.png">', ''), true);
  assert.strictEqual(api.shouldIgnorePaste(true, excelHtml('<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>'), 'a\tb\nc\td'), false);
});

test('drops images and objects inside cells', function () {
  const html = excelHtml('<table><tr><td>左<img alt="chart" src="x.png"></td><td>右</td></tr><tr><td>1</td><td>2</td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, '左\t右\n1\t2\n'),
    '| 左 | 右 |\n| --- | --- |\n| 1 | 2 |');
});

test('does not convert a single cell', function () {
  const html = excelHtml('<table><tr><td>メモ</td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, 'メモ'), null);
});

test('does not convert ordinary text or irregular tabs', function () {
  assert.strictEqual(api.convertClipboard('', 'ただの文章です'), null);
  assert.strictEqual(api.convertClipboard('', 'function\tfoo() {\n\treturn 1;\n}'), null);
});

test('reads an Excel table that follows the Windows clipboard header', function () {
  const html = 'Version:1.0\nStartHTML:0000000105\nEndHTML:0000000999\n' +
    excelHtml('<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, 'a\tb\nc\td\n'),
    '| a | b |\n| --- | --- |\n| c | d |');
});

test('does not steal a web page that contains a table plus other text', function () {
  const html = '<html><body><p>説明文です</p><table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table></body></html>';
  assert.strictEqual(api.convertClipboard(html, '説明文です\na\tb\nc\td\n'), null);
});

test('converts a copied HTML table that has no surrounding text', function () {
  const html = '<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>';
  assert.strictEqual(api.convertClipboard(html, 'a\tb\nc\td'),
    '| a | b |\n| --- | --- |\n| c | d |');
});

test('falls back to rectangular TSV when HTML is absent', function () {
  assert.strictEqual(api.convertClipboard('', '名前\t数量\r\nりんご\t3\r\n'),
    '| 名前 | 数量 |\n| --- | --- |\n| りんご | 3 |');
});

test('parses quoted TSV cells that contain newlines', function () {
  assert.strictEqual(api.convertClipboard('', '"line1\nline2"\tnext\nother\tcell\n'),
    '| line1<br>line2 | next |\n| --- | --- |\n| other | cell |');
});

test('trims a trailing empty row and column from the selection', function () {
  const html = excelHtml('<table><tr><td>a</td><td>b</td><td></td></tr><tr><td>c</td><td>d</td><td>&nbsp;</td></tr><tr><td></td><td></td><td></td></tr></table>');
  assert.strictEqual(api.convertClipboard(html, ''),
    '| a | b |\n| --- | --- |\n| c | d |');
});

test('recognizes Google Sheets clipboard HTML', function () {
  const html = '<google-sheets-html-origin><table data-sheets-root="1"><tr><td>a</td><td>b</td></tr></table>';
  assert.strictEqual(api.isSpreadsheetHtml(html), true);
  assert.strictEqual(api.convertClipboard(html, 'a'), '| a | b |\n| --- | --- |');
});

test('keeps a markdown table as its own block', function () {
  assert.strictEqual(api.ensureBlock('hello', ' world', '| a |\n| --- |'),
    '\n\n| a |\n| --- |\n\n');
  assert.strictEqual(api.ensureBlock('', '', '| a |\n| --- |'), '| a |\n| --- |\n');
  assert.strictEqual(api.ensureBlock('hello\n', '', '| a |\n| --- |'),
    '\n| a |\n| --- |\n');
});

test('does not treat text inside a code fence as a paste target', function () {
  const value = 'before\n```\ncode\n```\nafter';
  assert.strictEqual(api.insideFence(value, value.indexOf('code')), true);
  assert.strictEqual(api.insideFence(value, 0), false);
  assert.strictEqual(api.insideFence(value, value.length), false);
});

let failed = 0;
cases.forEach(function (entry) {
  try {
    entry[1]();
    console.log('ok ' + entry[0]);
  } catch (error) {
    failed += 1;
    console.error('FAIL ' + entry[0]);
    console.error(error && error.stack ? error.stack : error);
  }
});

if (typeof DOMParser === 'undefined') {
  console.error('DOMParser is unavailable; HTML table cases were not fully executed');
  process.exit(1);
}

if (failed) process.exit(1);
console.log(cases.length + ' passed');
