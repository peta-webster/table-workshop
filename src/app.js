import { SAMPLE, parseSource, parseHtml, toMarkdown, columnWeights, slidePages, typedValue } from './core/model.js';
import { wordBlob, excelBlob, pptBlob, clipboardHtml, saveBlob } from './exporters/office.js';

const $ = s => document.querySelector(s);
const state = { model: null, target: 'word', theme: 'clean', title: '办公用品清单（示例）', busy: false };
const targetThemes = { word: 'clean', excel: 'plain', ppt: 'clean' };
const isPlainExcel = () => state.target === 'excel' && state.theme === 'plain';
function selectTarget(target) {
  targetThemes[state.target] = state.theme;
  state.target = target;
  state.theme = targetThemes[target];
}
const meta = {
  word: { ext: 'docx', label: 'Word · A4', hint: '复制的表格随正文宽度调整。粘贴到 Word 后选择“保留源格式”，保留表头颜色、字体与间距。' },
  excel: { ext: 'xlsx', label: 'Excel · 工作表', hint: '数字和百分比可计算；编号、长数字和日期样式文本保留原值。需要精确保留时请下载文件。' },
  ppt: { ext: 'pptx', label: 'PowerPoint · 16:9', hint: '导出可编辑的原生表格，长表自动拆页并重复表头。复制到 PowerPoint 的效果因版本而异，建议下载。' }
};
function message(text, error = false) { $('#message').textContent = text; $('#message').classList.toggle('error', error); }
function makeTable(model, rows = model.rows) {
  const table = document.createElement('table'); table.className = state.theme;
  const group = document.createElement('colgroup');
  columnWeights(model).forEach(w => { const col = document.createElement('col'); col.style.width = w * 100 + '%'; group.append(col); }); table.append(group);
  const head = document.createElement('thead'); const body = document.createElement('tbody');
  [model.headers, ...rows].forEach((row, r) => {
    const tr = document.createElement('tr');
    row.forEach((value, c) => { const cell = document.createElement(r === 0 ? 'th' : 'td'); if (r === 0) cell.scope = 'col'; cell.textContent = value; cell.style.textAlign = isPlainExcel() ? (r && typeof typedValue(value, model.headers[c]).value === 'number' ? 'right' : 'left') : model.align[c] || 'left'; tr.append(cell); });
    (r ? body : head).append(tr);
  });
  table.append(head, body); return table;
}
function render() {
  const m = meta[state.target];
  const plain = isPlainExcel();
  $('#plain-theme').hidden = state.target !== 'excel'; $('#plain-theme').disabled = state.target !== 'excel';
  $('#theme').value = state.theme;
  $('#title-label').textContent = plain ? '文件名' : '表格标题';
  $('#copy').textContent = plain ? '复制数据' : '复制表格';
  $('#download').textContent = state.busy ? '正在生成…' : '下载 .' + m.ext + ' ↓';
  $('#output-hint').textContent = plain ? '只输出数据，保留编号与百分比类型。粘贴到已有 Excel 时选择“匹配目标格式”；下载文件不添加配色、边框、标题行或固定行列尺寸。' : m.hint;
  $('#preview-caption').textContent = m.label + (plain ? ' · 仅数据，网格不导出' : state.target === 'word' && state.model?.headers.length > 6 ? ' 横向' : '');
  $('#download').disabled = !state.model || state.busy; $('#copy').disabled = !state.model || state.busy;
  $('#targets').querySelectorAll('button').forEach(b => { const selected = b.dataset.target === state.target; b.classList.toggle('active', selected); b.setAttribute('aria-pressed', String(selected)); });
  $('#source-kind').textContent = state.model?.kind || 'Markdown / 网页表格';
  $('#dimensions').textContent = state.model ? state.model.rows.length + ' 行 × ' + state.model.headers.length + ' 列' : '';
  const preview = $('#preview'); preview.className = 'preview-sheet ' + state.target + '-sheet'; preview.replaceChildren();
  if (!state.model) { const empty = document.createElement('div'); empty.className = 'empty'; empty.textContent = '粘贴一张表格，在这里预览排版'; preview.append(empty); return; }
  const title = document.createElement('h3'); title.textContent = state.title; if (state.title && !plain) preview.append(title);
  let rows = state.model.rows;
  if (state.target === 'ppt') {
    try {
      const result = slidePages(state.model); rows = result.pages[0].rows;
      $('#preview-caption').textContent = m.label + ' · 第 1 / ' + result.pages.length + ' 页';
    } catch (error) { message(error.message, true); $('#download').disabled = true; }
  }
  const scroll = document.createElement('div'); scroll.className = 'table-scroll'; scroll.append(makeTable(state.model, rows)); preview.append(scroll);
}
function updateSource(text = $('#source').value) {
  $('#source').value = text;
  try { state.model = parseSource(text); message(state.model ? '已识别表格，可以选择样式并导出。' : ''); }
  catch (error) { state.model = null; message(error.message, true); }
  render();
}
function configure({ source, title, theme, target }) {
  if (source !== undefined && typeof source !== 'string') throw new Error('source 必须是字符串。');
  if (title !== undefined && (typeof title !== 'string' || title.length > 80)) throw new Error('标题长度不能超过 80 个字符。');
  if (theme !== undefined && !['plain', 'clean', 'academic', 'presentation'].includes(theme)) throw new Error('未知样式。');
  if (target !== undefined && !Object.keys(meta).includes(target)) throw new Error('未知目标软件。');
  if (theme === 'plain' && (target || state.target) !== 'excel') throw new Error('“无样式（仅数据）”仅适用于 Excel。');
  // Validate before changing visible state.
  const parsed = source !== undefined ? parseSource(source) : state.model;
  if (source !== undefined) { $('#source').value = source; state.model = parsed; }
  if (title !== undefined) { state.title = title; $('#table-title').value = title; }
  if (target !== undefined) selectTarget(target);
  if (theme !== undefined) { state.theme = theme; targetThemes[state.target] = theme; }
  message(state.model ? '已识别表格，可以选择样式并导出。' : ''); render(); return snapshot();
}
function snapshot() { return { target: state.target, theme: state.theme, title: state.title, rows: state.model?.rows.length || 0, columns: state.model?.headers.length || 0, ready: !!state.model }; }
async function exportFile() {
  if (!state.model || state.busy) return;
  const { model, title, theme, target } = state;
  state.busy = true; render();
  try {
    const build = { word: wordBlob, excel: excelBlob, ppt: pptBlob }[target];
    const blob = await build(model, title, theme);
    const name = (title.trim() || '表格').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_');
    saveBlob(blob, name + '.' + meta[target].ext);
    message('文件已生成。打开后可以继续编辑表格。');
  } catch (error) { message('导出失败：' + error.message, true); }
  finally { state.busy = false; render(); }
}
async function copyTable() {
  if (!state.model) return;
  const html = clipboardHtml(state.model, state.theme, state.target);
  // Plain-text fallback keeps line breaks inside a cell unambiguous.
  const plain = [state.model.headers, ...state.model.rows].map(r => r.map(v => /[\t\n"]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v).join('\t')).join('\n');
  try {
    if (navigator.clipboard?.write && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([plain], { type: 'text/plain' }) })]);
    } else {
      let wrote = false;
      const handler = e => { e.preventDefault(); e.clipboardData.setData('text/html', html); e.clipboardData.setData('text/plain', plain); wrote = true; };
      document.addEventListener('copy', handler, { once: true });
      const ok = document.execCommand('copy'); document.removeEventListener('copy', handler);
      if (!ok || !wrote) throw new Error('浏览器不支持带格式复制');
    }
    message(isPlainExcel() ? '已复制数据。粘贴到已有 Excel 后选择“匹配目标格式”，外观和数字显示沿用目标单元格；需要保留原来的百分比、小数位显示时请下载 .xlsx。' : state.target === 'word'
      ? '已复制。粘贴到 Word 后，点击表格旁的粘贴选项，选择“保留源格式”；“匹配目标格式”会覆盖字体颜色。'
      : '已复制。请在目标软件中粘贴并选择保留源格式；如有错位，请使用文件导出。');
  } catch { message('当前浏览器未允许复制，请使用右侧下载按钮导出表格。', true); }
}
$('#source').addEventListener('input', () => updateSource());
$('#source').addEventListener('paste', event => {
  const html = event.clipboardData?.getData('text/html');
  if (!html || !/<table[\s>]/i.test(html)) return;
  event.preventDefault();
  try {
    const model = parseHtml(html); $('#source').value = toMarkdown(model); state.model = model;
    message('已读取网页表格，并转为可修改的 Markdown。'); render();
  } catch (error) {
    const pasted = event.clipboardData.getData('text/plain') || html;
    $('#source').value = pasted; state.model = null; message(error.message, true); render();
  }
});
$('#sample').onclick = () => configure({ source: SAMPLE, title: '办公用品清单（示例）' });
$('#clear').onclick = () => { updateSource(''); $('#source').focus(); };
$('#table-title').oninput = event => { state.title = event.target.value; render(); };
$('#theme').onchange = event => { state.theme = event.target.value; targetThemes[state.target] = state.theme; render(); };
$('#targets').onclick = event => { const button = event.target.closest('[data-target]'); if (button) { selectTarget(button.dataset.target); message(state.model ? '已切换输出格式。' : ''); render(); } };
$('#download').onclick = exportFile; $('#copy').onclick = copyTable;
updateSource(SAMPLE);

// Optional agent interface, sharing the exact same validation and visible state.
const context = document.modelContext;
if (context?.registerTool) {
  const lifecycle = new AbortController();
  const registrations = [
    { name: 'configure_table', title: '设置表格与排版', description: '输入一张表格并设置目标软件、标题或样式，更新页面预览，不下载文件。plain 无样式仅适用于 Excel，此时 title 只作为下载文件名。', inputSchema: { type: 'object', properties: { source: { type: 'string' }, title: { type: 'string', maxLength: 80 }, target: { enum: ['word', 'excel', 'ppt'] }, theme: { enum: ['plain', 'clean', 'academic', 'presentation'] } }, additionalProperties: false }, annotations: { readOnlyHint: false }, execute: configure },
    { name: 'get_table_status', title: '读取表格状态', description: '读取当前表格尺寸、输出软件与样式。', inputSchema: { type: 'object', properties: {}, additionalProperties: false }, annotations: { readOnlyHint: true }, execute: snapshot }
  ];
  for (const tool of registrations) { try { Promise.resolve(context.registerTool(tool, { signal: lifecycle.signal })).catch(() => {}); } catch {} }
  window.addEventListener('pagehide', () => lifecycle.abort(), { once: true });
}
