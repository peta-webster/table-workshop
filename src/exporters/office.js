import { THEMES, typedValue, columnWeights, estimateLines, slidePages, escapeHtml, visibleLength } from '../core/model.js';

export async function wordBlob(model, title, themeName) {
  const D = window.docx; if (!D) throw new Error('Word 导出组件加载失败，请刷新页面重试。');
  const theme = THEMES[themeName]; const landscape = model.headers.length > 6;
  const pageWidth = landscape ? 16838 : 11906; const pageHeight = landscape ? 11906 : 16838;
  const available = pageWidth - 1800; const widths = columnWeights(model).map(w => Math.floor(w * available));
  widths[widths.length - 1] += available - widths.reduce((a, b) => a + b, 0);
  const none = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const border = { style: D.BorderStyle.SINGLE, size: 4, color: theme.line };
  const rows = [model.headers, ...model.rows].map((row, r, all) => new D.TableRow({
    tableHeader: r === 0, cantSplit: false,
    children: row.map((text, c) => {
      let borders = { top: none, left: none, right: none, bottom: border };
      if (themeName === 'academic') borders = { top: r === 0 ? { ...border, size: 10 } : none, left: none, right: none, bottom: r === 0 ? border : r === all.length - 1 ? { ...border, size: 10 } : none };
      return new D.TableCell({
        width: { size: widths[c], type: D.WidthType.DXA }, borders,
        margins: { top: 100, bottom: 100, left: 100, right: 100 },
        shading: { fill: r === 0 ? theme.head : r % 2 === 0 ? theme.band : 'FFFFFF' },
        verticalAlign: D.VerticalAlign.TOP,
        children: text.split('\n').map(line => new D.Paragraph({
          alignment: model.align[c] || 'left',
          spacing: { after: 0, line: 280 },
          children: [new D.TextRun({ text: line, bold: r === 0, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Microsoft YaHei' }, size: 21, color: r === 0 ? theme.ink : '202B40' })]
        }))
      });
    })
  }));
  const table = new D.Table({ rows, columnWidths: widths, width: { size: available, type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED, borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none } });
  const children = [];
  if (title) children.push(new D.Paragraph({ heading: D.HeadingLevel.TITLE, spacing: { after: 240 }, children: [new D.TextRun({ text: title, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Microsoft YaHei' }, size: 32, bold: true, color: '202B40' })] }));
  children.push(table);
  const doc = new D.Document({ creator: '表格工坊', title, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin: { top: 900, bottom: 900, left: 900, right: 900 } } }, children }] });
  return D.Packer.toBlob(doc);
}

export async function excelBlob(model, title, themeName) {
  const E = window.ExcelJS; if (!E) throw new Error('Excel 导出组件加载失败，请刷新页面重试。');
  const theme = THEMES[themeName]; const book = new E.Workbook(); book.creator = '表格工坊';
  if (themeName === 'plain') {
    const sheet = book.addWorksheet('表格');
    [model.headers, ...model.rows].forEach((row, r) => row.forEach((text, c) => {
      const cell = sheet.getCell(r + 1, c + 1);
      const typed = r ? typedValue(text, model.headers[c]) : { value: text, format: '@' };
      cell.value = typed.value;
      // Number formats preserve data meaning (percentages and identifiers), not decoration.
      cell.numFmt = typed.format;
    }));
    return new Blob([await book.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  const sheet = book.addWorksheet('表格', { views: [{ state: 'frozen', ySplit: title ? 3 : 1, showGridLines: false }] });
  const weights = columnWeights(model); const widths = weights.map(w => Math.max(12, Math.min(55, w * 120)));
  widths.forEach((w, c) => sheet.getColumn(c + 1).width = w);
  let headerRow = 1;
  if (title) {
    sheet.mergeCells(1, 1, 1, model.headers.length);
    sheet.getCell(1, 1).value = title; sheet.getCell(1, 1).font = { name: 'Arial', size: 17, bold: true, color: { argb: 'FF202B40' } }; sheet.getRow(1).height = 33;
    headerRow = 3; sheet.getRow(2).height = 10;
  }
  [model.headers, ...model.rows].forEach((row, r, all) => {
    const target = sheet.getRow(r + headerRow);
    row.forEach((text, c) => {
      const cell = target.getCell(c + 1); const typed = r ? typedValue(text, model.headers[c]) : { value: text, format: '@' };
      cell.value = typed.value; cell.numFmt = typed.format;
      cell.font = { name: 'Arial', size: 11, bold: r === 0, color: { argb: 'FF' + (r === 0 ? theme.ink : '202B40') } };
      cell.alignment = { vertical: 'top', horizontal: r && typeof typed.value === 'number' ? 'right' : model.align[c] || 'left', wrapText: true, indent: 1 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + (r === 0 ? theme.head : r % 2 === 0 ? theme.band : 'FFFFFF') } };
      const line = { style: 'thin', color: { argb: 'FF' + theme.line } };
      cell.border = themeName === 'academic'
        ? { ...(r === 0 ? { top: { ...line, style: 'medium' }, bottom: line } : {}), ...(r === all.length - 1 ? { bottom: { ...line, style: 'medium' } } : {}) }
        : { bottom: line };
    });
    const height = Math.max(29, ...row.map((v, c) => estimateLines(v, widths[c] * 7 / 72 - .08, 11) * 15.5 + 10));
    if (height > 409) throw new Error('单行文字超出 Excel 的行高限制，请拆分过长的单元格后重试。');
    target.height = height;
  });
  sheet.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: headerRow + model.rows.length, column: model.headers.length } };
  sheet.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: headerRow + ':' + headerRow };
  const data = await book.xlsx.writeBuffer();
  return new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export async function pptBlob(model, title, themeName) {
  const P = window.PptxGenJS; if (!P) throw new Error('PowerPoint 导出组件加载失败，请刷新页面重试。');
  const ppt = new P(); ppt.layout = 'LAYOUT_WIDE'; ppt.author = '表格工坊'; ppt.subject = '可编辑表格'; ppt.title = title; ppt.lang = 'zh-CN';
  ppt.theme = { headFontFace: 'Microsoft YaHei', bodyFontFace: 'Microsoft YaHei', lang: 'zh-CN' };
  const theme = THEMES[themeName]; const { pages, widths, fontSize } = slidePages(model);
  pages.forEach((page, i) => {
    const slide = ppt.addSlide(); slide.background = { color: 'FFFFFF' };
    slide.addShape(ppt.ShapeType.rect, { x: .64, y: .48, w: .08, h: .34, fill: { color: theme.accent }, line: { color: theme.accent, transparency: 100 } });
    if (title) slide.addText(title, { x: .87, y: .28, w: 11.8, h: .82, fontSize: visibleLength(title) > 90 ? 17 : 25, bold: true, color: '202B40', margin: 0, breakLine: false });
    const rows = [model.headers, ...page.rows].map((row, r, all) => row.map((text, c) => {
      const normal = { type: 'solid', pt: .6, color: theme.line }; const empty = { type: 'solid', pt: 0, color: 'FFFFFF' };
      const borders = themeName === 'academic' ? [r === 0 ? { ...normal, pt: 1.4 } : empty, empty, r === 0 ? normal : r === all.length - 1 ? { ...normal, pt: 1.4 } : empty, empty] : [empty, empty, normal, empty];
      return { text, options: { bold: r === 0, color: r === 0 ? theme.ink : '202B40', fill: { color: r === 0 ? theme.head : r % 2 === 0 ? theme.band : 'FFFFFF' }, align: model.align[c] || 'left', border: borders } };
    }));
    slide.addTable(rows, { x: .64, y: 1.28, w: 12.05, colW: widths, rowH: page.heights, fontFace: 'Microsoft YaHei', fontSize, margin: [6, 7, 6, 7], valign: 'top', paraSpaceAfter: 0, autoPage: false, breakLine: false });
    slide.addText((i + 1) + ' / ' + pages.length, { x: 11.9, y: 6.95, w: .75, h: .2, fontFace: 'Arial', fontSize: 10, color: '788398', align: 'right', margin: 0 });
  });
  return ppt.write({ outputType: 'blob' });
}

export function clipboardHtml(model, themeName, target) {
  if (target === 'excel' && themeName === 'plain') {
    // Keep Excel's cell types without importing fonts, fills, borders or dimensions.
    const rows = [model.headers, ...model.rows].map((row, r) => '<tr>' + row.map((text, c) => {
      const typed = r ? typedValue(text, model.headers[c]) : { value: text, format: '@' };
      // Give x:str an explicit value: browsers serialize a bare attribute as x:str="",
      // which Excel interprets as an empty cell even when the cell contains text.
      const type = typeof typed.value === 'number' ? ' x:num="' + typed.value + '"' : typed.value === null ? '' : ' x:str="' + escapeHtml(text).replace(/\n/g, '&#10;') + '"';
      const format = typed.format === '@' ? '\\@' : typed.format;
      const style = typed.value === null ? '' : ' style=\'mso-number-format:"' + format + '";\'';
      return '<td' + type + style + '>' + escapeHtml(text).replace(/\n/g, '<br style="mso-data-placement:same-cell">') + '</td>';
    }).join('') + '</tr>').join('');
    return '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="UTF-8"></head><body><!--StartFragment--><table>' + rows + '</table><!--EndFragment--></body></html>';
  }
  const theme = THEMES[themeName]; const weights = columnWeights(model);
  const forWord = target === 'word';
  const rows = [model.headers, ...model.rows].map((row, r, all) => '<tr>' + row.map((v, c) => {
    const numeric = r && typeof typedValue(v, model.headers[c]).value === 'number';
    const tag = r === 0 ? 'th' : 'td';
    const border = themeName === 'academic'
      ? (r === 0 ? 'border-top:2px solid #202b40;border-bottom:1px solid #202b40;' : r === all.length - 1 ? 'border-bottom:2px solid #202b40;' : 'border:0;')
      : 'border-bottom:1px solid #' + theme.line + ';';
    const excelFormat = target === 'excel' && !numeric ? 'mso-number-format:"\\@";' : '';
    const width = (weights[c] * 100).toFixed(2) + '%';
    const align = model.align[c] || 'left';
    const textStyle = 'color:#' + (r === 0 ? theme.ink : '202B40') + ';font-family:Arial,Microsoft YaHei,sans-serif;mso-fareast-font-family:Microsoft YaHei;font-size:11pt;font-weight:' + (r === 0 ? 'bold' : 'normal') + ';';
    // Word does not reliably inherit character formatting from table cells.
    let content = '<span style="' + textStyle + '">' + escapeHtml(v).replace(/\n/g, '<br style="mso-data-placement:same-cell">') + '</span>';
    // Reset the destination document's paragraph spacing and first-line indent.
    if (forWord) content = '<p style="margin:0;text-indent:0;line-height:115%;text-align:' + align + ';' + textStyle + '">' + content + '</p>';
    return '<' + tag + ' width="' + width + '" style=\'' + excelFormat + 'width:' + width + ';padding:6pt;border-collapse:collapse;border:0;' + border + 'vertical-align:top;text-align:' + align + ';background:#' + (r === 0 ? theme.head : r % 2 === 0 ? theme.band : 'FFFFFF') + ';' + textStyle + '\'>' + content + '</' + tag + '>';
  }).join('') + '</tr>').join('');
  // A percentage width lets Word use the recipient's text area, including margins.
  const tableWidth = forWord ? '100%' : '680px';
  return '<html><head><meta charset="UTF-8"></head><body><!--StartFragment--><table' + (forWord ? ' width="100%"' : '') + ' cellspacing="0" cellpadding="0" style="border-collapse:collapse;table-layout:fixed;width:' + tableWidth + ';margin:0;mso-table-lspace:0pt;mso-table-rspace:0pt;">' + rows + '</table><!--EndFragment--></body></html>';
}
export function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 30000);
}
