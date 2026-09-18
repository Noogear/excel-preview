/**
 * P1 导入侧新增模块的单元测试：表格样式实体化 + 行高自适应。
 * 两个模块都是纯函数，适合在 node 环境直接测。
 *
 * 同一函数的相邻场景（A1 区间解析、表头配色、computeAutoRowHeights 规则）合并进单个 `it`，
 * 断言逐条保留。
 */
import { describe, expect, it } from 'vitest';

import {
  computeAutoRowHeights,
  estimateLineCount,
  estimateWrappedRowHeight,
  isWideChar,
  measureTextWidth,
  ptToPx,
} from '../../src/importer/auto-height';
import { buildTableStylePatches, parseA1Range, pickReadableTextColor } from '../../src/importer/table-style';
import { toUniverWorkbook } from '../../src/importer/to-univer';
import type { ParsedTable, ParsedWorkbook } from '../../src/parser/types';

/** 12 个主题色，按写入顺序：dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink */
const THEME = [
  '#000000', // dk1
  '#FFFFFF', // lt1
  '#44546A', // dk2
  '#E7E6E6', // lt2
  '#4472C4', // accent1
  '#ED7D31', // accent2
  '#A5A5A5', // accent3
  '#FFC000', // accent4
  '#5B9BD5', // accent5
  '#70AD47', // accent6
  '#0563C1', // hlink
  '#954F72', // folHlink
];

describe('parseA1Range', () => {
  it('区间 / 单元格 / 绝对引用与多字母列都能解析；反向区间自动纠正；非法输入返回 null', () => {
    expect(parseA1Range('A1:D5')).toEqual({ startRow: 0, startCol: 0, endRow: 4, endCol: 3 });
    expect(parseA1Range('B2')).toEqual({ startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
    expect(parseA1Range('$AA$10:$AB$12')).toEqual({ startRow: 9, startCol: 26, endRow: 11, endCol: 27 });

    expect(parseA1Range('D5:A1')).toEqual({ startRow: 0, startCol: 0, endRow: 4, endCol: 3 });
    expect(parseA1Range('不是范围')).toBeNull();
    expect(parseA1Range('')).toBeNull();
  });
});

describe('pickReadableTextColor', () => {
  it('深底配白字、浅底配深字；缺色时回退到深色字', () => {
    expect(pickReadableTextColor('#000000')).toBe('#FFFFFF');
    expect(pickReadableTextColor('#FFFFFF')).toBe('#1F2937');
    // 橙色 #ED7D31 的相对亮度约 0.59，属于"深底" → 白字（与 Excel 的 Medium 表头一致）
    expect(pickReadableTextColor('#ED7D31')).toBe('#FFFFFF');
    expect(pickReadableTextColor('#FFE699')).toBe('#1F2937');
    expect(pickReadableTextColor(undefined)).toBe('#000000');
  });
});

describe('buildTableStylePatches', () => {
  const baseTable: ParsedTable = {
    name: 'SalesTable',
    ref: 'A1:D5',
    headerRowCount: 1,
    totalsRowCount: 0,
    styleName: 'TableStyleMedium9', // → 强调色 accent2
    showRowStripes: true,
  };

  it('Medium 家族：表头用强调色实底 + 可读字色 + 加粗，整行 4 列都盖到；强调色很浅时表头自动改用深色字', () => {
    const patches = buildTableStylePatches([baseTable], THEME);
    const header = patches.filter((p) => p.row === 0);
    expect(header).toHaveLength(4); // A..D
    for (const patch of header) {
      expect(patch.patch.fill).toBe('#ED7D31');
      expect(patch.patch.bold).toBe(true);
      // Excel 的 Medium 系列表头就是"橙底白字"，与 pickReadableTextColor 的判定一致
      expect(patch.patch.color).toBe('#FFFFFF');
    }

    // 强调色很浅时表头自动改用深色字（避免浅底白字看不清）
    const lightTheme = [...THEME];
    lightTheme[5] = '#FFE699'; // accent2 换成浅黄（TableStyleMedium9 → accent2）
    const lightHeader = buildTableStylePatches([baseTable], lightTheme).filter((p) => p.row === 0);
    expect(lightHeader[0].patch.fill).toBe('#FFE699');
    expect(lightHeader[0].patch.color).toBe('#1F2937');
  });

  it('斑马纹只落在奇数行体行上；关闭后没有条纹', () => {
    const patches = buildTableStylePatches([baseTable], THEME);
    const filledRows = new Set(patches.filter((p) => p.patch.fill && p.row > 0).map((p) => p.row));
    expect([...filledRows].sort()).toEqual([2, 4]); // 体行索引 1、3 → 行 2、4
    expect(buildTableStylePatches([{ ...baseTable, showRowStripes: false }], THEME).filter((p) => p.row > 0))
      .toHaveLength(0);
  });

  it('汇总行加粗 + 上边框；首列/末列强调', () => {
    const withTotals = buildTableStylePatches([{ ...baseTable, ref: 'A1:D6', totalsRowCount: 1 }], THEME);
    const totals = withTotals.filter((p) => p.row === 5);
    expect(totals).toHaveLength(4);
    expect(totals[0].patch.bold).toBe(true);
    expect(totals[0].patch.border?.top?.style).toBe('thin');

    const emphasised = buildTableStylePatches(
      [{ ...baseTable, showFirstColumn: true, showLastColumn: true }],
      THEME,
    );
    const bodyBold = emphasised.filter((p) => p.row > 0 && p.patch.bold).map((p) => p.col);
    expect(bodyBold).toContain(0);
    expect(bodyBold).toContain(3);
  });

  it('Light 家族用浅底深字；缺少主题色时回退到默认强调色；无表格返回空数组', () => {
    const light = buildTableStylePatches([{ ...baseTable, styleName: 'TableStyleLight11' }], THEME);
    expect(light.filter((p) => p.row === 0)[0].patch.fill).not.toBe('#ED7D31');
    expect(light.filter((p) => p.row === 0)[0].patch.color).not.toBe('#FFFFFF');

    expect(buildTableStylePatches([baseTable], undefined).find((p) => p.row === 0)?.patch.fill).toBe('#4472C4');

    expect(buildTableStylePatches(undefined, THEME)).toEqual([]);
    expect(buildTableStylePatches([], THEME)).toEqual([]);
  });
});

describe('auto-height 度量与 computeAutoRowHeights', () => {
  it('全角字符按 1em、半角按 0.55em，isWideChar 能区分中英文与全角标点，pt→px 换算正确', () => {
    expect(measureTextWidth('中文')).toBe(2);
    expect(measureTextWidth('abcd')).toBeCloseTo(2.2, 5);
    expect(isWideChar('中'.codePointAt(0) ?? 0)).toBe(true);
    expect(isWideChar('，'.codePointAt(0) ?? 0)).toBe(true);
    expect(isWideChar('a'.codePointAt(0) ?? 0)).toBe(false);
    expect(ptToPx(11)).toBeCloseTo(14.667, 3);
  });

  it('行数随文本长度增长并统计显式换行；行高估算随行数增长', () => {
    const short = estimateLineCount('短', 11, 100);
    const long = estimateLineCount('中'.repeat(40), 11, 100);
    expect(long).toBeGreaterThan(short);
    expect(estimateLineCount('a\nb\nc', 11, 1000)).toBe(3);
    expect(estimateWrappedRowHeight('中'.repeat(60), 11, 200))
      .toBeGreaterThan(estimateWrappedRowHeight('短', 11, 200));
  });

  it('只处理开启了自动换行的单元格且不给手工设过高度的行加高（与 Excel 行为一致）；同一行取所有单元格的最大需求；短文本与空文本不触发加高', () => {
    const base = { defaultRowHeightPx: 24 };
    expect(computeAutoRowHeights({
      ...base,
      rowsWithCustomHeight: new Set<number>(),
      cells: [{ row: 0, col: 0, text: '中'.repeat(80), fontSizePt: 11, wrap: false, usableWidthPx: 100 }],
    })).toEqual({});
    expect(computeAutoRowHeights({
      ...base,
      rowsWithCustomHeight: new Set([0]),
      cells: [{ row: 0, col: 0, text: '中'.repeat(80), fontSizePt: 11, wrap: true, usableWidthPx: 100 }],
    })).toEqual({});

    // 同一行取所有单元格的最大需求
    const heights = computeAutoRowHeights({
      defaultRowHeightPx: 24,
      rowsWithCustomHeight: new Set<number>(),
      cells: [
        { row: 3, col: 0, text: '中'.repeat(20), fontSizePt: 11, wrap: true, usableWidthPx: 100 },
        { row: 3, col: 1, text: '中'.repeat(120), fontSizePt: 11, wrap: true, usableWidthPx: 100 },
      ],
    });
    expect(heights[3]).toBeGreaterThan(estimateWrappedRowHeight('中'.repeat(20), 11, 100 - 8));

    // 短文本与空文本不触发加高
    expect(computeAutoRowHeights({
      defaultRowHeightPx: 24,
      rowsWithCustomHeight: new Set<number>(),
      cells: [
        { row: 0, col: 0, text: '短', fontSizePt: 11, wrap: true, usableWidthPx: 200 },
        { row: 0, col: 1, text: '', fontSizePt: 11, wrap: true, usableWidthPx: 10 },
      ],
    })).toEqual({});
  });
});

/**
 * 适配层集成：证明"表格样式实体化"与"行高自适应"真的写进了 IWorkbookData
 * （不依赖解析器是否已实现 tables 解析，用手工构造的中立模型驱动）
 */
describe('toUniverWorkbook：表格样式与行高落地', () => {
  function makeParsed(overrides: Partial<ParsedWorkbook> = {}): ParsedWorkbook {
    return {
      sheets: [
        {
          id: 's1',
          name: '表格',
          index: 0,
          rows: {},
          cols: {},
          cells: [
            { row: 0, col: 0, value: '区域' },
            { row: 1, col: 0, value: '华东' },
            { row: 2, col: 0, value: '华南' },
          ],
          merges: [],
          tables: [
            {
              name: 'T1',
              ref: 'A1:B4',
              headerRowCount: 1,
              totalsRowCount: 0,
              styleName: 'TableStyleMedium9',
              showRowStripes: true,
            },
          ],
        },
      ],
      styles: [],
      themeColors: THEME,
      report: { unsupported: [], warnings: [] },
      raw: { entries: {} },
      ...overrides,
    };
  }

  it('表头拿到强调色底与加粗、斑马纹行拿到浅色底、未加斑马纹的体行没有底色，且样式记入降级报告', () => {
    const outcome = toUniverWorkbook(makeParsed());
    const sheet = outcome.workbookData.sheets['s1']!;
    const styles = outcome.workbookData.styles;

    const headerId = sheet.cellData![0][0].s as string;
    const headerStyle = styles[headerId]!;
    expect(headerId, '表头应被赋予实体化样式').toBeTruthy();
    expect(headerStyle.bg).toEqual({ rgb: '#ED7D31' });
    expect(Number(headerStyle.bl)).toBe(1);

    const stripedId = sheet.cellData![2][0].s as string;
    const stripedStyle = styles[stripedId]!;
    expect(stripedStyle?.bg, '斑马纹行必须有底色').toBeTruthy();
    expect((stripedStyle.bg as { rgb?: string }).rgb).not.toBe('#ED7D31');

    // 未加斑马纹的体行不该有底色
    const plainId = sheet.cellData![1][0].s as string | undefined;
    expect(plainId ? styles[plainId]?.bg : undefined).toBeFalsy();

    expect(outcome.report.warnings.join(' ')).toContain('表格');
  });

  it('表格样式补丁会与单元格原有样式合并（底/字色来自表格，字体与居中保留）', () => {
    const parsed = makeParsed({
      sheets: [
        {
          id: 's1',
          name: '表格',
          index: 0,
          rows: {},
          cols: {},
          cells: [{ row: 0, col: 0, value: '区域', styleIndex: 0 }],
          merges: [],
          tables: [
            {
              name: 'T1', ref: 'A1:B2', headerRowCount: 1, totalsRowCount: 0,
              styleName: 'TableStyleMedium9', showRowStripes: true,
            },
          ],
        },
      ],
      // 原样式带字体与居中，实体化后必须保留
      styles: [{ fontFamily: '微软雅黑', fontSize: 14, horizontalAlign: 'center' }],
    });

    const sheet = toUniverWorkbook(parsed).workbookData.sheets['s1']!;
    const style = toUniverWorkbook(parsed).workbookData.styles[sheet.cellData![0][0].s as string]!;
    expect(style.bg).toEqual({ rgb: '#ED7D31' });
    expect(style.ff).toBe('微软雅黑');
    expect(style.fs).toBe(14);
    expect(style.ht).toBe(2);
  });

  it('开启自动换行的长文本行会被加高，普通行保持默认', () => {
    const parsed = makeParsed({
      sheets: [
        {
          id: 's1',
          name: '换行',
          index: 0,
          rows: {},
          cols: {},
          cells: [
            { row: 0, col: 0, value: '短', styleIndex: 0 },
            { row: 1, col: 0, value: '这是一段很长的中文文本，用于触发自动换行并且需要多行才能显示完整内容', styleIndex: 0 },
          ],
          merges: [],
        },
      ],
      styles: [{ textWrap: true, fontSize: 11 }],
    });

    const sheet = toUniverWorkbook(parsed).workbookData.sheets['s1']!;
    expect(sheet.rowData![0]?.h, '短文本行不应被加高').toBeUndefined();
    expect(sheet.rowData![1]?.h ?? 0, '长文本行应被加高').toBeGreaterThan(24);
  });
});
