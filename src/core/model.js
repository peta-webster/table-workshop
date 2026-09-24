import { marked } from '../vendor/marked.js';

export const SAMPLE = "| 物品编号 | 物品 | 单价 | 可用率 | 备注 |\n| --- | --- | ---: | ---: | --- |\n| 00123 | 文件夹 | 29.90 | 98.5% | 演示用虚构数据，支持中文自动换行 |\n| 00456 | 标签纸 | 0 | 100% | 颜色分类<br>便于查找 |\n| 123456789012345678 | 收纳盒 | 199.00 | 92.0% | 长编号按文本保存 |\n| 00789 | 记事本 |  |  | 空白单元格保持为空 |";
export const THEMES = {
  clean: { name: '简洁办公', head: 'EDF2FA', ink: '334364', band: 'F8FAFF', line: 'DFE5EF', accent: '2C4CF0' },
  academic: { name: '学术三线表', head: 'FFFFFF', ink: '202B40', band: 'FFFFFF', line: '202B40', accent: '202B40' },
  presentation: { name: '汇报演示', head: '263D76', ink: 'FFFFFF', band: 'E7EDF7', line: 'FFFFFF', accent: '2C4CF0' }
};

function fragment(html) {
  const template = document.createElement('template');
  template.innerHTML = html;
  return template.content;
}
function plain(node) {
  const copy = node.cloneNode(true);
  copy.querySelectorAll('script,style,noscript,iframe,object').forEach(n => n.remove());
  copy.querySelectorAll('br').forEach(n => n.replaceWith('\n'));
  copy.querySelectorAll('p,div,li').forEach(n => n.append('\n'));
  copy.querySelectorAll('img').forEach(n => n.replaceWith(n.alt || '[图片]'));
  copy.querySelectorAll('a[href]').forEach(n => {
    const href = n.getAttribute('href');
    if (/^https?:\/\//i.test(href) && n.textContent.trim() !== href) n.append(' (' + href + ')');
  });
  return copy.textContent.replace(/\u00a0/g, ' ').trim();
}
function splitRow(line) {
  line = line.trim();
  if (line.startsWith('|')) line = line.slice(1);
  if (/(?<!\\)\|$/.test(line)) line = line.slice(0, -1);
  const cells = []; let cell = ''; let slashes = 0;
  for (const ch of line) {
    if (ch === '|' && slashes % 2 === 0) { cells.push(cell.trim()); cell = ''; }
    else cell += ch;
    slashes = ch === '\\' ? slashes + 1 : 0;
  }
  cells.push(cell.trim()); return cells;
}
function check(model) {
  const count = model.headers.length;
  if (!count) throw new Error('没有识别到表格列。');
  if (count > 30 || model.rows.length * count > 10000) throw new Error('这张表较大，请分段处理（最多 30 列、10,000 个单元格）。');
  model.rows.forEach((row, i) => {
    if (row.length !== count) throw new Error('第 ' + (i + 1) + ' 行有 ' + row.length + ' 列，表头有 ' + count + ' 列。请补齐空值或检查竖线。');
  });
  return model;
}
export function parseHtml(html) {
  const root = fragment(html); const tables = root.querySelectorAll('table');
  if (tables.length !== 1) throw new Error(tables.length ? '检测到多张表格，请一次只粘贴一张。' : '没有识别到 HTML 表格。');
  const table = tables[0];
  if (Array.from(table.querySelectorAll('td,th')).some(c => c.colSpan > 1 || c.rowSpan > 1)) throw new Error('暂不支持合并单元格，请先取消合并再粘贴。');
  const trs = Array.from(table.rows);
  if (!trs.length) throw new Error('表格没有内容。');
  const rows = trs.map(tr => Array.from(tr.cells).map(plain));
  const align = Array.from(trs[0].cells).map(c => c.style.textAlign || c.getAttribute('align') || 'left');
  return check({ headers: rows[0], rows: rows.slice(1), align, kind: '网页表格' });
}
export function parseSource(text) {
  text = text.trim();
  if (!text) return null;
  if (/<table[\s>]/i.test(text)) return parseHtml(text);
  text = text.replace(/^~~~(?:markdown|md)?\s*\n([\s\S]*?)\n~~~$/i, '$1');
  const fence = String.fromCharCode(96).repeat(3);
  if (text.startsWith(fence) && text.endsWith(fence)) text = text.slice(text.indexOf('\n') + 1, -3).trim();
  const tokens = marked.lexer(text);
  const tables = [];
  function collect(items) { for (const token of items) { if (token.type === 'table') tables.push(token); else if (token.tokens) collect(token.tokens); } }
  collect(tokens);
  if (!tables.length && text.includes('\t')) {
    const rows = text.split(/\r?\n/).map(line => line.split('\t'));
    return check({ headers: rows[0], rows: rows.slice(1), align: rows[0].map(() => 'left'), kind: '制表符表格' });
  }
  if (tables.length !== 1) throw new Error(tables.length ? '检测到多张表格，请一次只保留一张。' : '没有识别到表格。Markdown 表头下面需要一行 | --- | --- |。');
  const token = tables[0]; const rawRows = token.raw.trim().split(/\r?\n/);
  const width = splitRow(rawRows[0]).length;
  rawRows.slice(2).forEach((line, i) => {
    if (splitRow(line).length !== width) throw new Error('第 ' + (i + 1) + ' 行列数与表头不一致。空单元格也需要保留分隔符，内容中的竖线请写成 \\|。');
  });
  const cellText = c => plain(fragment(marked.parseInline(c.text)));
  return check({ headers: token.header.map(cellText), rows: token.rows.map(r => r.map(cellText)), align: token.align.map(a => a || 'left'), kind: 'Markdown' });
}
export function toMarkdown(model) {
  const escape = v => v.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/([*_\[\]~])/g, '\\$1').replace(/`/g, '\\`').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  const row = r => '| ' + r.map(escape).join(' | ') + ' |';
  return [row(model.headers), row(model.align.map(a => a === 'right' ? '---:' : a === 'center' ? ':---:' : '---')), ...model.rows.map(row)].join('\n');
}
export function typedValue(text, header = '') {
  if (text === '') return { value: null, format: 'General' };
  if (/编号|代码|编码|账号|账户|电话|邮编|证件|序号|(^|\s)(id|sku|code|zip|phone)(\s|$)/i.test(header)) return { value: text, format: '@' };
  const pattern = /^-?(?:0|[1-9]\d*|[1-9]\d{0,2}(?:,\d{3})+)(?:\.\d+)?%?$/;
  if (!pattern.test(text)) return { value: text, format: '@' };
  const cleaned = text.replace(/,/g, '').replace(/%$/, '');
  const digits = cleaned.replace(/[-.]/g, '').replace(/^0+/, '');
  if (digits.length > 15 || /^-0(?:\.0+)?$/.test(cleaned)) return { value: text, format: '@' };
  const percent = text.endsWith('%'); const value = Number(cleaned) / (percent ? 100 : 1);
  if (!Number.isFinite(value)) return { value: text, format: '@' };
  const decimals = (cleaned.split('.')[1] || '').length;
  return { value, format: (text.includes(',') ? '#,##0' : '0') + (decimals ? '.' + '0'.repeat(decimals) : '') + (percent ? '%' : '') };
}
export function visibleLength(text) {
  return Array.from(text).reduce((s, ch) => s + (/[\u1100-\uffff]/.test(ch) ? 2 : 1), 0);
}
export function columnWeights(model) {
  const scores = model.headers.map((h, c) => {
    const values = [h, ...model.rows.map(r => r[c])].map(v => Math.max(...v.split('\n').map(visibleLength)));
    values.sort((a, b) => a - b);
    const isIdentifier = /编号|代码|编码|账号|ID|SKU/i.test(h);
    const p = values[isIdentifier ? values.length - 1 : Math.floor((values.length - 1) * .85)] || 0;
    return Math.max(7, Math.min(40, Math.max(visibleLength(h) + 2, p + 2)));
  });
  const sum = scores.reduce((a, b) => a + b, 0);
  return scores.map(n => n / sum);
}
export function escapeHtml(text) { return text.replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch])); }
export function estimateLines(text, width, size) {
  // Conservative CJK-aware wrapping estimate; actual Office font metrics can vary.
  const capacity = Math.max(1, width * 72 / (size * .57));
  return text.split('\n').reduce((n, line) => n + Math.max(1, Math.ceil(visibleLength(line) / capacity)), 0);
}
export function slidePages(model) {
  if (model.headers.length > 10) throw new Error('PowerPoint 当前支持最多 10 列；这张宽表建议导出 Excel 或 Word。');
  const widths = columnWeights(model).map(w => w * 12.05);
  const fontSize = model.headers.length > 7 ? 14 : 17;
  const rowHeight = (row, size = fontSize) => Math.max(.46, ...row.map((v, c) => estimateLines(v, widths[c] - .2, size) * size * 1.3 / 72 + .18));
  const headerHeight = rowHeight(model.headers);
  if (headerHeight > 1.8) throw new Error('表头文字太长，请缩短表头后导出 PowerPoint。');
  const pages = []; let rows = []; let heights = [headerHeight]; let used = headerHeight;
  for (const row of model.rows) {
    const h = rowHeight(row);
    if (h + headerHeight > 5.15) throw new Error('某个单元格内容超过一页幻灯片，请拆分这一行后导出 PowerPoint。');
    if (used + h > 5.15 && rows.length) { pages.push({ rows, heights }); rows = []; heights = [headerHeight]; used = headerHeight; }
    rows.push(row); heights.push(h); used += h;
  }
  pages.push({ rows, heights });
  return { pages, widths, fontSize };
}
