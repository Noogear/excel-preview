/**
 * `writeWorkbookPackage`：把 CSV / ODS / XLS 转出来的中性模型打成**合法的最小 xlsx**。
 *
 * 这是"支持更多格式"的地基：只要产物是一份规范 xlsx，下游（解析 → Univer → 编辑 → 工作区 →
 * 撤销 → 外科式导出 → 会话恢复）就一行都不用改。所以本文件的目标是**闭环验证**：
 *   写入 → 我们自己的解析器读回来 → 值/样式/合并/尺寸/冻结逐项对得上。
 * 另外还有一条真机验证（`tools/make-legacy-fixtures.ps1` 生成的真 Excel 样本做往返），
 * 由 tests/unit/legacy-formats.test.ts 覆盖。
 */
import { describe, expect, it } from 'vitest';

import { parseXlsx } from '../../src/parser';
import {
  escapeXml,
  normalizeColor,
  sanitizeSheetName,
  writeWorkbookPackage,
  type WorkbookInput,
} from '../../src/importer/synth-xlsx';

async function roundTrip(input: WorkbookInput) {
  const bytes = writeWorkbookPackage(input);
  const parsed = await parseXlsx(bytes);
  return { bytes, parsed };
}

describe('writeWorkbookPackage：值与公式', () => {
  it('数字/文本/布尔/公式都能原样读回来，空字符串不写单元格', async () => {
    const { parsed } = await roundTrip({
      name: 't',
      sheets: [
        {
          name: '数据',
          cells: [
            { row: 0, col: 0, value: '姓名' },
            { row: 0, col: 1, value: 42 },
            { row: 0, col: 2, value: 3.5 },
            { row: 0, col: 3, value: true },
            { row: 0, col: 4, value: false },
            { row: 0, col: 5, value: '' }, // 空串 → 不写 <c>
            { row: 1, col: 0, value: 3, formula: 'SUM(B1:C1)' },
            { row: 1, col: 1, value: 0 }, // 0 是内容
          ],
        },
      ],
    });

    const sheet = parsed.sheets[0];
    const at = (row: number, col: number) => sheet.cells.find((c) => c.row === row && c.col === col);
    expect(at(0, 0)?.value).toBe('姓名');
    expect(at(0, 1)?.value).toBe(42);
    expect(at(0, 2)?.value).toBe(3.5);
    expect(at(0, 3)?.value).toBe(true);
    expect(at(0, 4)?.value).toBe(false);
    expect(at(0, 5), '空串不该产生单元格').toBeUndefined();
    expect(at(1, 0)?.formula).toBe('SUM(B1:C1)');
    expect(at(1, 1)?.value, '0 必须保留').toBe(0);
  });

  it('文本里的 XML 特殊字符与换行原样保留（Excel 里不会变成乱码）', async () => {
    const text = `A<B & C>"D" 'E'\n第二行`;
    const { parsed } = await roundTrip({ sheets: [{ name: 's', cells: [{ row: 0, col: 0, value: text }] }] });
    expect(parsed.sheets[0].cells.find((c) => c.row === 0 && c.col === 0)?.value).toBe(text);
  });
});

describe('writeWorkbookPackage：样式', () => {
  it('字体/加粗/斜体/下划线/删除线/颜色/填充/对齐/换行/旋转/数字格式都进 cellXfs 并可读回', async () => {
    const { parsed } = await roundTrip({
      sheets: [
        {
          name: 's',
          cells: [{ row: 0, col: 0, value: 'X', style: 0 }],
        },
      ],
      styles: [
        {
          fontFamily: '微软雅黑',
          fontSize: 14,
          bold: true,
          italic: true,
          underline: true,
          strikeThrough: true,
          color: '#CC0000',
          fill: '#FFF2CC',
          horizontalAlign: 'center',
          verticalAlign: 'middle',
          textWrap: true,
          textRotation: 45,
          numberFormat: '0.00"元"',
          border: { top: { style: 'thin', color: '#4F81BD' }, bottom: { style: 'double' } },
        },
      ],
    });

    // styles[0] 是默认样式，我们的样式应落在 styles[1]（字段名与 `ParsedStyle` 一致）
    const style = parsed.styles[1];
    expect(style).toBeTruthy();
    expect(style.fontFamily).toBe('微软雅黑');
    expect(style.fontSize).toBe(14);
    expect(style.bold).toBe(true);
    expect(style.italic).toBe(true);
    expect(style.underline).toBe(true);
    expect(style.strikeThrough).toBe(true);
    expect(style.color).toBe('#CC0000');
    expect(style.fill).toBe('#FFF2CC');
    expect(style.horizontalAlign).toBe('center');
    expect(style.verticalAlign).toBe('middle');
    expect(style.textWrap).toBe(true);
    expect(style.textRotation).toBe(45);
    expect(style.numberFormat).toBe('0.00"元"');
    expect(style.border?.top, '边框线型与颜色都要还原').toEqual({ style: 'thin', color: '#4F81BD' });
    expect(style.border?.bottom?.style).toBe('double');

    // 单元格确实引用了这个样式
    const cell = parsed.sheets[0].cells[0];
    expect(cell.styleIndex).toBe(1);
  });

  it('相同样式只 intern 一份（cellXfs 不膨胀）', async () => {
    const { parsed } = await roundTrip({
      sheets: [
        {
          name: 's',
          cells: [
            { row: 0, col: 0, value: 'a', style: 0 },
            { row: 0, col: 1, value: 'b', style: 0 },
            { row: 1, col: 0, value: 'c', style: 1 },
          ],
        },
      ],
      styles: [{ bold: true }, { bold: true }],
    });
    // 默认 + 两个相同样式 = 2 条
    expect(parsed.styles.length).toBe(2);
    expect(parsed.sheets[0].cells.map((cell) => cell.styleIndex)).toEqual([1, 1, 1]);
  });
});

describe('writeWorkbookPackage：结构（合并/列宽/行高/冻结/隐藏表/多表）', () => {
  it('合并区、列宽、行高、冻结、隐藏工作表与网格线设置都能读回', async () => {
    const { parsed } = await roundTrip({
      sheets: [
        {
          name: '第一张',
          cells: [{ row: 0, col: 0, value: '标题' }, { row: 1, col: 0, value: '正文' }],
          merges: ['A1:C1'],
          colWidths: { 0: 18.5, 2: 30 },
          rowHeights: { 0: 28, 1: 20 },
          freeze: { rows: 1, cols: 0 },
          gridlinesHidden: true,
        },
        { name: '第二张', hidden: true, cells: [{ row: 0, col: 0, value: '隐藏表' }] },
      ],
    });

    expect(parsed.sheets.map((s) => s.name)).toEqual(['第一张', '第二张']);
    const first = parsed.sheets[0];
    expect(first.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }]);
    expect(first.cols[0]?.width).toBe(18.5);
    expect(first.cols[2]?.width).toBe(30);
    expect(first.rows[0]?.height).toBe(28);
    expect(first.rows[1]?.height).toBe(20);
    expect(first.freeze).toEqual({ row: 1, col: 0 });
    expect(first.gridlinesHidden).toBe(true);
    expect(parsed.sheets[1].hidden).toBe(true);
  });

  it('工作表名按 Excel 规则清洗并去重（≤31 字符、去掉非法字符）', async () => {
    const { parsed } = await roundTrip({
      sheets: [
        { name: 'a/b:c*d?e[f]g', cells: [] },
        { name: 'a b c d', cells: [] },
        { name: 'a b c d', cells: [] },
        { name: '', cells: [] },
      ],
    });
    const names = parsed.sheets.map((s) => s.name);
    expect(names[0], '非法字符逐个换成空格').toBe('a b c d e f g');
    expect(new Set(names).size, '不允许重名').toBe(4);
    expect(names.every((name) => name.length <= 31)).toBe(true);
    expect(names[3]).toBe('工作表4');
  });
});

describe('小工具', () => {
  it('escapeXml 覆盖 & < > " \'', () => {
    expect(escapeXml(`&<>"'`)).toBe('&amp;&lt;&gt;&quot;&apos;');
  });

  it('normalizeColor 支持 #RGB/#RRGGBB/#AARRGGBB，非法值返回 undefined', () => {
    expect(normalizeColor('#abc')).toBe('#AABBCC');
    expect(normalizeColor('00FF00')).toBe('#00FF00');
    expect(normalizeColor('#80FF0000')).toBe('#FF0000');
    expect(normalizeColor('red')).toBeUndefined();
    expect(normalizeColor(undefined)).toBeUndefined();
  });

  it('sanitizeSheetName 处理空名/超长/重名（大小写不敏感）', () => {
    const used = new Set<string>();
    expect(sanitizeSheetName('  ', used, '回退')).toBe('回退');
    expect(sanitizeSheetName('x'.repeat(40), used)).toHaveLength(31);
    expect(sanitizeSheetName('Sheet1', used)).toBe('Sheet1');
    expect(sanitizeSheetName('sheet1', used)).toBe('sheet1 (2)');
  });
});
