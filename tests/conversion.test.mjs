import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';
import JSZip from 'jszip';
import { parseSource, parseHtml, typedValue, columnWeights, slidePages } from '../src/core/model.js';
import { clipboardHtml, wordBlob, excelBlob, pptBlob } from '../src/exporters/office.js';

const dom = new JSDOM('');
globalThis.document = dom.window.document;
const source = await readFile(new URL('../examples/data-types.md', import.meta.url), 'utf8');
const model = parseSource(source);
const sandbox = { console, Blob, Buffer, URL, TextEncoder, TextDecoder, setTimeout, clearTimeout, setImmediate, clearImmediate, atob, btoa };
sandbox.window = sandbox; sandbox.self = sandbox;
vm.createContext(sandbox);
for (const file of ['docx.js', 'exceljs.js', 'pptxgen.js']) vm.runInContext(await readFile(new URL('../src/vendor/' + file, import.meta.url), 'utf8'), sandbox, { filename: file });
globalThis.window = sandbox;
const xml = text => new dom.window.DOMParser().parseFromString(text, 'application/xml');
const unzip = async blob => JSZip.loadAsync(Buffer.from(await blob.arrayBuffer()));

test('Markdown, HTML and TSV share cell values and preserve line breaks', () => {
  assert.deepEqual(model.headers, ['编号', '月费', '比例', '说明', '文本']);
  assert.equal(model.rows[0][3], '第一行\n第二行');
  assert.equal(model.rows[1][3], '');
  const html = parseHtml('<table><tr><th>项目</th><th>说明</th></tr><tr><td>A</td><td>一<br>二</td></tr></table>');
  assert.deepEqual(html.rows, [['A', '一\n二']]);
  assert.deepEqual(parseSource('项目\t说明\nA\t一').rows, [['A', '一']]);
});

test('ambiguous layouts fail visibly instead of silently dropping cells', () => {
  assert.throws(() => parseSource('| A | B |\n| --- | --- |\n| 1 | 2 | 3 |'), /列数/);
  assert.throws(() => parseHtml('<table><tr><td colspan="2">合并</td></tr></table>'), /合并/);
  assert.throws(() => parseHtml('<table></table><table></table>'), /多张/);
});

test('identifiers, dates and formula-like input remain text', () => {
  for (const value of ['00123', '123456789012345678', '2026-09-15', '=1+1', '+SUM(A1:A2)']) {
    assert.deepEqual(typedValue(value), { value, format: '@' });
  }
  assert.deepEqual(typedValue('123', '编号'), { value: '123', format: '@' });
  assert.deepEqual(typedValue('29.90'), { value: 29.9, format: '0.00' });
  assert.deepEqual(typedValue('98.5%'), { value: .985, format: '0.0%' });
  assert.deepEqual(typedValue(''), { value: null, format: 'General' });
});

test('Word clipboard uses relative widths and explicit white header text', () => {
  const fragment = new JSDOM(clipboardHtml(model, 'presentation', 'word')).window.document;
  const table = fragment.querySelector('table');
  assert.equal(table.getAttribute('width'), '100%');
  assert.equal(table.style.width, '100%');
  assert.equal(fragment.querySelector('th span').style.color, 'rgb(255, 255, 255)');
  assert.equal(fragment.querySelector('th p').style.textIndent, '0');
  assert.ok(!table.outerHTML.includes('680px'));
});

test('Excel clipboard declares explicit text values and carries no decoration', () => {
  const html = clipboardHtml(model, 'plain', 'excel');
  const fragment = new JSDOM(html).window.document;
  assert.equal(fragment.querySelectorAll('th').length, 0);
  assert.equal(fragment.querySelectorAll('td').length, 15);
  assert.equal(fragment.querySelector('[x\\:str="00123"]').textContent, '00123');
  assert.ok(html.includes('x:num="0.985"'));
  assert.ok(html.includes('mso-number-format:"0.0%"'));
  assert.ok(!/font-|background|border|padding|width=|height=|text-align/i.test(html));
  const escaped = clipboardHtml({ headers: ['Text'], rows: [['<script>"&</script>']] }, 'plain', 'excel');
  assert.equal(new JSDOM(escaped).window.document.querySelectorAll('script').length, 0);
});

test('plain XLSX has exact values and no title, decoration or dimensions', async () => {
  const zip = await unzip(await excelBlob(model, 'Do not insert this title', 'plain'));
  const sheet = xml(await zip.file('xl/worksheets/sheet1.xml').async('string'));
  const styles = xml(await zip.file('xl/styles.xml').async('string'));
  const strings = xml(await zip.file('xl/sharedStrings.xml').async('string'));
  const allStrings = [...strings.getElementsByTagName('si')].map(si => si.textContent);
  const cells = [...sheet.getElementsByTagName('c')];
  function cell(address) { return cells.find(c => c.getAttribute('r') === address); }
  function value(address) {
    const c = cell(address); const v = c?.getElementsByTagName('v')[0]?.textContent;
    return c?.getAttribute('t') === 's' ? allStrings[Number(v)] : v === undefined ? null : Number(v);
  }
  assert.equal(sheet.getElementsByTagName('row').length, 3);
  assert.equal(value('A2'), '00123');
  assert.equal(value('A3'), '123456789012345678');
  assert.equal(value('B2'), 29.9); assert.equal(value('C2'), .985);
  assert.equal(value('D2'), '第一行\n第二行'); assert.equal(value('D3'), null);
  assert.equal(value('E2'), '=1+1'); assert.equal(value('E3'), '2026-09-15');
  assert.ok(!allStrings.includes('Do not insert this title'));
  for (const tag of ['mergeCells', 'autoFilter', 'pane', 'cols', 'f']) assert.equal(sheet.getElementsByTagName(tag).length, 0);
  for (const row of sheet.getElementsByTagName('row')) assert.equal(row.hasAttribute('ht'), false);
  for (const xf of styles.getElementsByTagName('cellXfs')[0].children) {
    for (const key of ['fontId', 'fillId', 'borderId']) assert.equal(xf.getAttribute(key), '0');
    assert.equal(xf.getElementsByTagName('alignment').length, 0);
  }
});

test('Word file contains a native table with complete text and header color', async () => {
  const zip = await unzip(await wordBlob(model, 'Sample', 'presentation'));
  const doc = xml(await zip.file('word/document.xml').async('string'));
  assert.equal(doc.getElementsByTagName('w:tbl').length, 1);
  assert.equal(doc.getElementsByTagName('w:tr').length, 3);
  assert.ok(doc.documentElement.textContent.includes('123456789012345678'));
  assert.ok([...doc.getElementsByTagName('w:color')].some(c => c.getAttribute('w:val') === 'FFFFFF'));
});

test('long PowerPoint tables retain all rows across editable slides', async () => {
  const long = { headers: ['编号', '说明'], rows: Array.from({ length: 45 }, (_, i) => [String(i + 1), '中文内容 ' + (i + 1)]), align: ['left', 'left'] };
  const plan = slidePages(long);
  assert.ok(plan.pages.length > 1);
  assert.equal(plan.pages.flatMap(p => p.rows).length, 45);
  assert.ok(Math.abs(columnWeights(long).reduce((a, b) => a + b, 0) - 1) < 1e-9);
  for (const page of plan.pages) assert.ok(page.heights.reduce((a, b) => a + b, 0) <= 5.15);
  const zip = await unzip(await pptBlob(long, 'Long table', 'clean'));
  const names = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  assert.equal(names.length, plan.pages.length);
  let dataRows = 0;
  for (const name of names) {
    const doc = xml(await zip.file(name).async('string'));
    assert.equal(doc.getElementsByTagName('a:tbl').length, 1);
    dataRows += doc.getElementsByTagName('a:tr').length - 1;
  }
  assert.equal(dataRows, 45);
});

test('new styles retain cell content, borders and emphasis across clipboard and Office files', async () => {
  for (const name of ['print', 'horizontal', 'compact', 'firstColumn']) {
    const fragment = new JSDOM(clipboardHtml(model, name, 'word')).window.document;
    const first = fragment.querySelector('td');
    assert.equal(first.textContent, '00123');
    assert.equal(fragment.querySelector('table').style.width, '100%');
    assert.equal(first.style.borderLeftWidth, ['print', 'compact'].includes(name) ? '1px' : '0px');
    assert.equal(first.style.padding, name === 'compact' ? '3pt' : name === 'horizontal' ? '9pt' : '6pt');
    assert.equal(first.querySelector('span').style.fontWeight, name === 'firstColumn' ? 'bold' : 'normal');
    if (name === 'print') assert.equal(first.querySelector('span').style.color, 'rgb(0, 0, 0)');

    const word = xml(await (await unzip(await wordBlob(model, '', name))).file('word/document.xml').async('string'));
    const wordCell = word.getElementsByTagName('w:tr')[1].getElementsByTagName('w:tc')[0];
    assert.equal(wordCell.getElementsByTagName('w:t')[0].textContent, '00123');
    assert.equal(wordCell.getElementsByTagName('w:left')[0].getAttribute('w:val'), ['print', 'compact'].includes(name) ? 'single' : 'none');
    assert.equal(wordCell.getElementsByTagName('w:b')[0]?.getAttribute('w:val') !== 'false', name === 'firstColumn');

    const book = new sandbox.ExcelJS.Workbook();
    await book.xlsx.load(Buffer.from(await (await excelBlob(model, '', name)).arrayBuffer()).toString('base64'), { base64: true });
    const sheet = book.worksheets[0];
    assert.equal(sheet.getCell('A2').value, '00123');
    assert.equal(sheet.getCell('C2').value, .985);
    assert.equal(sheet.getCell('A2').border.left?.style, ['print', 'compact'].includes(name) ? 'thin' : undefined);
    assert.equal(!!sheet.getCell('A2').font.bold, name === 'firstColumn');

    const ppt = xml(await (await unzip(await pptBlob(model, '', name))).file('ppt/slides/slide1.xml').async('string'));
    assert.equal(ppt.getElementsByTagName('a:tr').length, model.rows.length + 1);
    const pptCell = ppt.getElementsByTagName('a:tr')[1].getElementsByTagName('a:tc')[0];
    assert.ok(pptCell.textContent.includes('00123'));
    assert.equal(pptCell.getElementsByTagName('a:rPr')[0].getAttribute('b') === '1', name === 'firstColumn');
  }
});

test('style-aware PowerPoint spacing preserves long rows and page bounds', () => {
  const long = { headers: ['编号', '说明'], rows: Array.from({ length: 60 }, (_, i) => [String(i), '合成说明内容']), align: ['left', 'left'] };
  for (const name of ['print', 'horizontal', 'compact', 'firstColumn']) {
    const plan = slidePages(long, name);
    assert.deepEqual(plan.pages.flatMap(p => p.rows), long.rows);
    for (const page of plan.pages) assert.ok(page.heights.reduce((a, b) => a + b, 0) <= 5.15);
  }
  assert.ok(slidePages(long, 'compact').pages[0].rows.length > slidePages(long, 'horizontal').pages[0].rows.length);
});
