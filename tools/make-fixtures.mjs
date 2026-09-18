/**
 * 生成合成 xlsx 测试样本（P0/P1 回归语料库的种子）。
 *
 * 覆盖用户指定的"必须 100% 保真"六类：
 *   1. 边框/填充/字体/对齐/合并/行高列宽
 *   2. 数字格式（日期/百分比/千分位/货币/自定义）
 *   3. 条件格式（高亮/色阶/数据条/图标集）
 *   4. 批注 + 超链接 + 数据验证
 *   5. 浮动图片
 *   6. Excel 表格(Table) 斑马纹 + 表头样式
 * 另加：7. 启用宏的工作簿（.xlsm，宏部件原样保留但不执行）
 *
 * 运行：npm run fixtures
 * 产出：fixtures/*.xlsx + fixtures/fixture-macro.xlsm + fixtures/manifest.json（期望值清单）
 */
import ExcelJS from 'exceljs';
import { unzipSync, zipSync } from 'fflate';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = join(ROOT, 'fixtures');

mkdirSync(OUT_DIR, { recursive: true });

// ---------------------------------------------------------------- PNG 生成
const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([length, typeBuf, data, crc]);
}

/** 生成一张纯色 PNG（保证是合法 PNG，Excel 与我们的渲染器都能读） */
function makePng(width, height, [r, g, b, a = 255]) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (width * 4 + 1);
    raw[rowStart] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const p = rowStart + 1 + x * 4;
      raw[p] = r;
      raw[p + 1] = g;
      raw[p + 2] = b;
      raw[p + 3] = a;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const FONT = '微软雅黑';
const manifest = {};

async function save(workbook, fileName) {
  const target = join(OUT_DIR, fileName);
  await workbook.xlsx.writeFile(target);
  console.log(`✓ ${fileName}`);
}

// ------------------------------------------------- 1. 样式基础
async function fixtureStyles() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'P0 fixture generator';
  const ws = wb.addWorksheet('样式', {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }],
    properties: { defaultRowHeight: 18, defaultColWidth: 10 },
  });

  ws.columns = [
    { width: 18 },
    { width: 14 },
    { width: 14 },
    { width: 22 },
    { width: 12 },
    { width: 12 },
  ];

  // 合并标题
  ws.mergeCells('A1:F1');
  const title = ws.getCell('A1');
  title.value = '样式保真样本（P0 fixture）';
  title.font = { name: FONT, size: 16, bold: true, color: { argb: 'FF1F2937' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
  ws.getRow(1).height = 32;

  // 表头：四边边框 + 填充 + 白字 + 居中
  const headers = ['区域', '数量', '单价', '生效日期', '占比', '备注'];
  headers.forEach((text, index) => {
    const cell = ws.getCell(2, index + 1);
    cell.value = text;
    cell.font = { name: FONT, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF1E40AF' } },
      bottom: { style: 'thin', color: { argb: 'FF1E40AF' } },
      left: { style: 'thin', color: { argb: 'FF1E40AF' } },
      right: { style: 'thin', color: { argb: 'FF1E40AF' } },
    };
  });

  // 各种边框线型（12 种中的代表）
  const borderStyles = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'dashDot', 'mediumDashed'];
  borderStyles.forEach((style, index) => {
    const cell = ws.getCell(3 + index, 6);
    cell.value = style;
    cell.border = {
      top: { style, color: { argb: 'FFDC2626' } },
      bottom: { style, color: { argb: 'FFDC2626' } },
      left: { style, color: { argb: 'FFDC2626' } },
      right: { style, color: { argb: 'FFDC2626' } },
    };
  });

  // 对齐 / 换行 / 缩进 / 旋转
  ws.getCell('D3').value = '自动换行的长文本，用于验证 wrapText 与行高自适应';
  ws.getCell('D3').alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
  ws.getRow(3).height = 44;
  ws.getCell('D4').value = '缩进2级';
  ws.getCell('D4').alignment = { indent: 2, horizontal: 'left' };
  ws.getCell('D5').value = '旋转45°';
  ws.getCell('D5').alignment = { textRotation: 45 };
  ws.getCell('D6').value = '垂直居中';
  ws.getCell('D6').alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(6).height = 30;

  // 字体样式组合
  ws.getCell('E3').value = '粗体';
  ws.getCell('E3').font = { name: FONT, size: 11, bold: true };
  ws.getCell('E4').value = '斜体';
  ws.getCell('E4').font = { name: FONT, size: 11, italic: true };
  ws.getCell('E5').value = '下划线';
  ws.getCell('E5').font = { name: FONT, size: 11, underline: true };
  ws.getCell('E6').value = '删除线';
  ws.getCell('E6').font = { name: FONT, size: 11, strike: true };
  ws.getCell('E7').value = '红色18号';
  ws.getCell('E7').font = { name: '宋体', size: 18, color: { argb: 'FFDC2626' } };

  // 富文本（单元格内混合格式）
  ws.getCell('F3').value = {
    richText: [
      { font: { name: FONT, size: 11, bold: true, color: { argb: 'FF0000FF' } }, text: '粗蓝' },
      { font: { name: FONT, size: 11, italic: true }, text: ' + 斜体' },
      { font: { name: FONT, size: 11, underline: true, color: { argb: 'FF16A34A' } }, text: ' + 下划线绿' },
    ],
  };

  // 填充色板
  const fills = ['FFFFF3C4', 'FFFDE68A', 'FFBBF7D0', 'FFBFDBFE', 'FFFECACA', 'FFE9D5FF'];
  fills.forEach((argb, index) => {
    const cell = ws.getCell(12 + index, 1);
    cell.value = argb;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  });

  // 行高/列宽/隐藏
  ws.getRow(20).height = 48;
  ws.getCell('A20').value = '行高48';
  ws.getColumn(7).width = 30;
  ws.getCell('G2').value = '列宽30';
  ws.getRow(21).hidden = true;
  ws.getCell('A21').value = '这行被隐藏';
  ws.getColumn(8).hidden = true;
  ws.getCell('H2').value = '这列被隐藏';

  await save(wb, 'fixture-styles.xlsx');

  manifest['fixture-styles.xlsx'] = {
    purpose: '边框/填充/字体/对齐/合并/行高列宽/隐藏/富文本',
    sheets: [{ name: '样式', frozen: { xSplit: 1, ySplit: 1 }, hiddenRows: [21], hiddenCols: [8] }],
    checks: [
      { a1: 'A1', expect: { merged: 'A1:F1', bold: true, fontSize: 16, align: 'center', fill: '#DBEAFE' } },
      { a1: 'A2', expect: { bold: true, color: '#FFFFFF', fill: '#2563EB', align: 'center', border: 'thin #1E40AF' } },
      { a1: 'F3', expect: { borderTop: 'thin' } },
      { a1: 'F11', expect: { borderTop: 'mediumDashed' } },
      { a1: 'D3', expect: { wrapText: true, verticalAlign: 'top' } },
      { a1: 'D4', expect: { indent: 2 } },
      { a1: 'D5', expect: { textRotation: 45 } },
      { a1: 'E5', expect: { underline: true } },
      { a1: 'E6', expect: { strikeThrough: true } },
      { a1: 'F3', expect: { richText: true }, note: 'P0 阶段降级为纯文本，P1 需还原' },
      { a1: 'A20', expect: { rowHeight: 48 } },
      { a1: 'G2', expect: { colWidth: 30 } },
    ],
  };
}

// ------------------------------------------------- 2. 数字格式
async function fixtureNumfmt() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('数字格式');
  ws.columns = [{ width: 20 }, { width: 24 }, { width: 34 }];

  ws.getRow(1).values = ['原始值', '显示效果', 'numFmt'];
  ws.getRow(1).font = { bold: true };

  const cases = [
    [1234.5678, '#,##0.00'],
    [-1234.5678, '#,##0.00;[Red]-#,##0.00'],
    [0.2567, '0.0%'],
    [0.2567, '0.00%'],
    [45658, 'yyyy-mm-dd'],
    [45658, 'yyyy"年"m"月"d"日"'],
    [0.75, 'h:mm AM/PM'],
    [1.5, '[h]:mm:ss'],
    [123456789, '0.00E+00'],
    [1234.5, '¥#,##0.00'],
    [1234.5, '"总计 "#,##0.00" 元"'],
    [12345, '#,##0'],
    [1234.5678, '0.000'],
    [0.5, '# ?/?'],
    ['00123', '@'],
    [45658.5, 'm/d/yy h:mm'],
    [-0.1234, '0.00%;[Red]-0.00%'],
  ];

  cases.forEach(([value, numFmt], index) => {
    const row = index + 2;
    ws.getCell(row, 1).value = value;
    ws.getCell(row, 2).value = value;
    ws.getCell(row, 2).numFmt = numFmt;
    ws.getCell(row, 3).value = numFmt;
  });

  const themeCell = ws.getCell('E2');
  themeCell.value = '主题色 accent1';
  try {
    themeCell.font = { name: FONT, size: 11, color: { theme: 4, tint: -0.25 } };
    themeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: 0.6 } };
    manifest.__themeColorWritten = true;
  } catch (error) {
    manifest.__themeColorWritten = false;
    manifest.__themeColorError = String(error);
  }

  await save(wb, 'fixture-numfmt.xlsx');

  manifest['fixture-numfmt.xlsx'] = {
    purpose: '数字格式显示文本保真（含自定义与内置格式）',
    sheets: [{ name: '数字格式' }],
    cases: cases.map(([value, numFmt], index) => ({ a1: `B${index + 2}`, value, numFmt })),
    themeColor: manifest.__themeColorWritten ? { a1: 'E2', theme: 4, tint: -0.25 } : null,
  };
  delete manifest.__themeColorWritten;
  delete manifest.__themeColorError;
}

// ------------------------------------------------- 3. 条件格式 / 数据验证
async function fixtureRules() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('条件格式');
  ws.columns = [{ width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }];

  ws.getRow(1).values = ['项目', '高亮>100', '色阶', '数据条', '图标集'];
  ws.getRow(1).font = { bold: true };

  const values = [
    ['甲', 120, 30, 55, 90],
    ['乙', 80, 55, 40, 70],
    ['丙', 210, 80, 95, 30],
    ['丁', 45, 12, 20, 55],
    ['戊', 175, 67, 78, 85],
  ];
  values.forEach((row, index) => {
    ws.getRow(index + 2).values = row;
  });

  ws.addConditionalFormatting({
    ref: 'B2:B6',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulae: [100],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
          font: { color: { argb: 'FF9C0006' }, bold: true },
        },
      },
    ],
  });

  ws.addConditionalFormatting({
    ref: 'C2:C6',
    rules: [
      {
        type: 'colorScale',
        priority: 2,
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }],
      },
    ],
  });

  ws.addConditionalFormatting({
    ref: 'D2:D6',
    rules: [
      {
        type: 'dataBar',
        priority: 3,
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: { argb: 'FF638EC6' },
      },
    ],
  });

  ws.addConditionalFormatting({
    ref: 'E2:E6',
    rules: [
      {
        type: 'iconSet',
        priority: 4,
        iconSet: '3TrafficLights1',
        cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }],
      },
    ],
  });

  // 数据验证
  const dvSheet = wb.addWorksheet('数据验证');
  dvSheet.getCell('A1').value = '性别(列表)';
  dvSheet.getCell('B1').value = '年龄(整数1-120)';
  dvSheet.getCell('C1').value = '比率(小数)';
  dvSheet.getCell('D1').value = '日期(2020后)';
  dvSheet.getCell('E1').value = '文本长度<=5';
  dvSheet.getRow(1).font = { bold: true };

  dvSheet.getCell('A2').dataValidation = { type: 'list', allowBlank: true, formulae: ['"男,女,未知"'] };
  dvSheet.getCell('A2').value = '男';
  dvSheet.getCell('B2').dataValidation = { type: 'whole', operator: 'between', allowBlank: true, formulae: [1, 120], showErrorMessage: true, errorTitle: '无效', error: '请输入1~120' };
  dvSheet.getCell('B2').value = 30;
  dvSheet.getCell('C2').dataValidation = { type: 'decimal', operator: 'greaterThan', allowBlank: true, formulae: [0] };
  dvSheet.getCell('C2').value = 0.5;
  dvSheet.getCell('D2').dataValidation = { type: 'date', operator: 'greaterThan', allowBlank: true, formulae: [new Date('2020-01-01')] };
  dvSheet.getCell('D2').value = new Date('2024-03-01');
  dvSheet.getCell('E2').dataValidation = { type: 'textLength', operator: 'lessThanOrEqual', allowBlank: true, formulae: [5] };
  dvSheet.getCell('E2').value = 'abc';

  await save(wb, 'fixture-rules.xlsx');

  manifest['fixture-rules.xlsx'] = {
    purpose: '条件格式四类 + 数据验证五类',
    sheets: [{ name: '条件格式' }, { name: '数据验证' }],
    conditionalFormatting: [
      { ref: 'B2:B6', kind: 'cellIs-greaterThan-100', fill: '#FFC7CE', fontColor: '#9C0006' },
      { ref: 'C2:C6', kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      { ref: 'D2:D6', kind: 'dataBar', color: '#638EC6' },
      { ref: 'E2:E6', kind: 'iconSet', iconSet: '3TrafficLights1' },
    ],
    dataValidations: [
      { a1: 'A2', type: 'list', formulae: '"男,女,未知"' },
      { a1: 'B2', type: 'whole', operator: 'between', formulae: [1, 120] },
      { a1: 'C2', type: 'decimal', operator: 'greaterThan', formulae: [0] },
      { a1: 'D2', type: 'date', operator: 'greaterThan' },
      { a1: 'E2', type: 'textLength', operator: 'lessThanOrEqual', formulae: [5] },
    ],
  };
}

// ------------------------------------------------- 4. 批注 / 超链接 / 图片
async function fixtureExtras() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('批注链接图片');
  ws.columns = [{ width: 20 }, { width: 30 }, { width: 20 }];

  ws.getCell('A1').value = '带批注的单元格';
  ws.getCell('A1').note = '这是一条批注（legacy note）\n第二行';
  ws.getCell('A2').value = '多行批注';
  ws.getCell('A2').note = '第一行\n第二行\n第三行';

  ws.getCell('B1').value = { text: '示例链接', hyperlink: 'https://example.com/path?q=1' };
  ws.getCell('B1').font = { color: { argb: 'FF0563C1' }, underline: true };
  ws.getCell('B2').value = { text: '内部链接', hyperlink: '#条件格式!A1' };
  ws.getCell('B3').value = { text: '邮件链接', hyperlink: 'mailto:test@example.com' };

  const png = makePng(64, 40, [37, 99, 235]);
  const imageId = wb.addImage({ buffer: png, extension: 'png' });
  ws.addImage(imageId, {
    tl: { col: 2.2, row: 4.2 },
    ext: { width: 128, height: 80 },
    editAs: 'oneCell',
  });
  const png2 = makePng(48, 48, [220, 38, 38]);
  const imageId2 = wb.addImage({ buffer: png2, extension: 'png' });
  ws.addImage(imageId2, {
    tl: { col: 2.2, row: 9.2 },
    br: { col: 3.4, row: 11.4 },
    editAs: 'twoCell',
  });

  ws.getCell('A6').value = '图片在右侧 C5 附近（oneCell 锚点）';
  ws.getCell('A11').value = '图片在右侧 C10 附近（twoCell 锚点）';

  await save(wb, 'fixture-extras.xlsx');

  manifest['fixture-extras.xlsx'] = {
    purpose: '批注、超链接、浮动图片',
    sheets: [{ name: '批注链接图片' }],
    notes: [
      { a1: 'A1', text: '这是一条批注（legacy note）\n第二行' },
      { a1: 'A2', text: '第一行\n第二行\n第三行' },
    ],
    hyperlinks: [
      { a1: 'B1', target: 'https://example.com/path?q=1' },
      { a1: 'B2', target: '#条件格式!A1', kind: 'internal' },
      { a1: 'B3', target: 'mailto:test@example.com' },
    ],
    images: [
      { index: 0, anchor: 'oneCell', tl: { col: 2.2, row: 4.2 }, size: { width: 128, height: 80 }, color: '#2563EB' },
      { index: 1, anchor: 'twoCell', tl: { col: 2.2, row: 9.2 }, br: { col: 3.4, row: 11.4 }, color: '#DC2626' },
    ],
  };
}

// ------------------------------------------------- 5. Excel 表格(Table)
async function fixtureTable() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('表格');
  ws.columns = [{ width: 16 }, { width: 12 }, { width: 12 }, { width: 14 }];

  ws.addTable({
    name: 'SalesTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: [{ name: '区域' }, { name: '数量' }, { name: '金额' }, { name: '日期' }],
    rows: [
      ['华东', 120, 3450.5, new Date('2024-01-15')],
      ['华南', 98, 2810.25, new Date('2024-02-20')],
      ['华北', 143, 4120.75, new Date('2024-03-05')],
      ['西南', 76, 1990.0, new Date('2024-04-11')],
    ],
  });

  // 另加一个带汇总行的表格
  ws.addTable({
    name: 'TotalsTable',
    ref: 'A8',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleLight11', showRowStripes: true, showFirstColumn: true },
    columns: [{ name: '项目' }, { name: 'Q1' }, { name: 'Q2' }],
    rows: [
      ['营收', 100, 200],
      ['成本', 40, 80],
    ],
  });
  ws.getCell('B11').value = { formula: 'SUM(B9:B10)', result: 140 };
  ws.getCell('C11').value = { formula: 'SUM(C9:C10)', result: 280 };

  await save(wb, 'fixture-table.xlsx');

  manifest['fixture-table.xlsx'] = {
    purpose: 'Excel 表格(Table)：表头样式 + 斑马纹 + 汇总行（样式需实体化到单元格）',
    sheets: [{ name: '表格' }],
    tables: [
      { name: 'SalesTable', ref: 'A1:D5', style: 'TableStyleMedium9', showRowStripes: true, totalsRow: false },
      { name: 'TotalsTable', ref: 'A8:C11', style: 'TableStyleLight11', showRowStripes: true, totalsRow: true },
    ],
    note: '表格的斑马纹/表头样式**不在单元格样式里**，而是来自 tableStyles.xml —— 这正是需要"实体化"的地方',
  };
}

// ------------------------------------------------- 6. 多表 / 隐藏表 / 公式
async function fixtureMulti() {
  const wb = new ExcelJS.Workbook();
  const first = wb.addWorksheet('第一张');
  const second = wb.addWorksheet('第二张');
  const hidden = wb.addWorksheet('隐藏表');
  const veryHidden = wb.addWorksheet('深度隐藏');

  first.getCell('A1').value = '工作表顺序与文件名不一致的样本';
  first.getCell('A2').value = 10;
  first.getCell('A3').value = 20;
  first.getCell('A4').value = { formula: 'SUM(A2:A3)', result: 30 };
  first.getCell('A5').value = { formula: '第二张!A2*2', result: 200 };
  first.getCell('A6').value = '跨表引用与公式缓存';

  second.getCell('A1').value = '第二张表';
  second.getCell('A2').value = 100;
  second.getCell('B2').value = { formula: 'A2/4', result: 25 };
  second.getCell('C2').value = { formula: 'IF(A2>50,"大","小")', result: '大' };

  hidden.getCell('A1').value = '我是隐藏表';
  veryHidden.getCell('A1').value = '我是深度隐藏表';

  hidden.state = 'hidden';
  veryHidden.state = 'veryHidden';

  await save(wb, 'fixture-multi.xlsx');

  manifest['fixture-multi.xlsx'] = {
    purpose: '多工作表、顺序与文件名不一致、隐藏/深度隐藏、公式与缓存值',
    sheets: [
      { name: '第一张', state: 'visible', index: 0 },
      { name: '第二张', state: 'visible', index: 1 },
      { name: '隐藏表', state: 'hidden', index: 2 },
      { name: '深度隐藏', state: 'veryHidden', index: 3 },
    ],
    formulas: [
      { sheet: '第一张', a1: 'A4', formula: 'SUM(A2:A3)', cached: 30 },
      { sheet: '第一张', a1: 'A5', formula: '第二张!A2*2', cached: 200 },
      { sheet: '第二张', a1: 'B2', formula: 'A2/4', cached: 25 },
      { sheet: '第二张', a1: 'C2', formula: 'IF(A2>50,"大","小")', cached: '大' },
    ],
    note: '导入时默认采用 Excel 缓存值渲染，保证与 Excel 显示一致',
  };
}

/**
 * ------------------------------------------------- 7. 启用宏的工作簿（.xlsm）
 *
 * ExcelJS 只会写 .xlsx，所以这里写完再"改包"：往 OOXML 包里塞一个 `xl/vbaProject.bin`，
 * 并把 `[Content_Types].xml` 改成 macroEnabled（真实 .xlsm 就是这样，工作簿部件内容类型不同）。
 * 目的：验证"我们能打开 .xlsm，宏部件原样保留、不执行"（见 `src/importer/file-kinds.ts`）。
 *
 * 注意：这里的 vbaProject.bin 是**占位字节**，不是可执行的 VBA 工程——
 * 我们既不解析也不执行宏，只要求导出时逐字节不变。
 */
async function fixtureMacro() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('宏工作簿');
  ws.getCell('A1').value = '这是一个启用宏的工作簿（.xlsm）';
  ws.getCell('A2').value = '宏部件原样保留，但本工具不会执行它';
  ws.getCell('A3').value = 2026;
  ws.getCell('B3').value = { formula: 'A3+1', result: 2027 };
  ws.getCell('A1').font = { bold: true, size: 13 };
  ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

  const buffer = await wb.xlsx.writeBuffer();
  const entries = unzipSync(new Uint8Array(buffer));

  // ① 加宏部件（OLE 复合文档的魔术头，够"像"一个 vbaProject.bin 即可）
  const stub = Buffer.alloc(512, 0);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(stub, 0);
  entries['xl/vbaProject.bin'] = new Uint8Array(stub);

  // ② 改 [Content_Types].xml：声明 .bin 部件 + 工作簿改成 macroEnabled
  const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
  const patched = ct
    .replace(
      /<Default Extension="bin"[^>]*\/>/,
      '',
    )
    .replace(
      '</Types>',
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
    )
    .replace(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    );
  entries['[Content_Types].xml'] = new TextEncoder().encode(patched);

  const out = zipSync(entries, { level: 6 });
  writeFileSync(join(OUT_DIR, 'fixture-macro.xlsm'), out);
  console.log('✓ fixture-macro.xlsm');

  manifest['fixture-macro.xlsm'] = {
    purpose: '启用宏的工作簿：能被打开预览；xl/vbaProject.bin 原样保留（不执行）',
    sheets: [{ name: '宏工作簿', state: 'visible', index: 0 }],
    preserved: ['宏(VBA)'],
    note: 'vbaProject.bin 是占位字节（不解析、不执行），导出时逐字节保留',
  };
}

await fixtureStyles();
await fixtureNumfmt();
await fixtureRules();
await fixtureExtras();
await fixtureTable();
await fixtureMulti();
await fixtureMacro();

writeFileSync(join(OUT_DIR, 'manifest.json'), JSON.stringify(manifest, null, 2), 'utf8');
console.log(`\n✓ fixtures/manifest.json（${Object.keys(manifest).length} 个样本）`);
console.log(`输出目录：${OUT_DIR}`);
