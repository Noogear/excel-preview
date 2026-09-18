import {
  BooleanNumber,
  HorizontalAlign,
  type IWorkbookData,
  LocaleType,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core';

/**
 * P0 冒烟数据：覆盖"必须 100% 保真"的几类样式的**渲染侧**能力。
 * 注意：这里验证的是 Univer 能否渲染出这些样式，导入保真度由 fixtures 走解析链路验证。
 */
export const P0_STYLE_IDS = {
  header: 'p0-header',
  title: 'p0-title',
  number: 'p0-number',
  percent: 'p0-percent',
  date: 'p0-date',
  bordered: 'p0-bordered',
  fill: 'p0-fill',
  wrapped: 'p0-wrapped',
  rotated: 'p0-rotated',
  boldItalic: 'p0-bold-italic',
} as const;

export function createSampleWorkbook(): IWorkbookData {
  return {
    id: 'p0-workbook',
    name: 'P0 冒烟工作簿',
    appVersion: '0.25.1',
    locale: LocaleType.ZH_CN,
    sheetOrder: ['p0-sheet-1'],
    styles: {
      [P0_STYLE_IDS.title]: {
        ff: 'Microsoft YaHei',
        fs: 16,
        bl: BooleanNumber.TRUE,
        cl: { rgb: '#1F2937' },
        ht: HorizontalAlign.CENTER,
        vt: VerticalAlign.MIDDLE,
      },
      [P0_STYLE_IDS.header]: {
        ff: 'Microsoft YaHei',
        fs: 11,
        bl: BooleanNumber.TRUE,
        cl: { rgb: '#FFFFFF' },
        bg: { rgb: '#2563EB' },
        ht: HorizontalAlign.CENTER,
        vt: VerticalAlign.MIDDLE,
        bd: {
          t: { s: 1, cl: { rgb: '#1E40AF' } },
          b: { s: 1, cl: { rgb: '#1E40AF' } },
          l: { s: 1, cl: { rgb: '#1E40AF' } },
          r: { s: 1, cl: { rgb: '#1E40AF' } },
        },
      },
      [P0_STYLE_IDS.number]: { n: { pattern: '#,##0.00' }, ht: HorizontalAlign.RIGHT },
      [P0_STYLE_IDS.percent]: { n: { pattern: '0.0%' }, ht: HorizontalAlign.RIGHT },
      [P0_STYLE_IDS.date]: { n: { pattern: 'yyyy-mm-dd' }, ht: HorizontalAlign.RIGHT },
      [P0_STYLE_IDS.bordered]: {
        bd: {
          t: { s: 2, cl: { rgb: '#DC2626' } },
          b: { s: 5, cl: { rgb: '#DC2626' } },
          l: { s: 3, cl: { rgb: '#DC2626' } },
          r: { s: 7, cl: { rgb: '#DC2626' } },
        },
      },
      [P0_STYLE_IDS.fill]: { bg: { rgb: '#FDE68A' } },
      [P0_STYLE_IDS.wrapped]: { tb: WrapStrategy.WRAP, vt: VerticalAlign.TOP },
      [P0_STYLE_IDS.rotated]: { tr: { a: 45, v: BooleanNumber.TRUE } },
      [P0_STYLE_IDS.boldItalic]: { bl: BooleanNumber.TRUE, it: BooleanNumber.TRUE },
    },
    sheets: {
      'p0-sheet-1': {
        id: 'p0-sheet-1',
        name: '样式冒烟',
        rowCount: 200,
        columnCount: 30,
        defaultColumnWidth: 96,
        defaultRowHeight: 24,
        freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
        cellData: {
          0: {
            0: { v: 'P0 渲染与保真冒烟表', s: P0_STYLE_IDS.title },
          },
          1: {
            0: { v: '区域', s: P0_STYLE_IDS.header },
            1: { v: '数量', s: P0_STYLE_IDS.header },
            2: { v: '单价', s: P0_STYLE_IDS.header },
            3: { v: '占比', s: P0_STYLE_IDS.header },
            4: { v: '生效日期', s: P0_STYLE_IDS.header },
          },
          2: {
            0: { v: '华东', s: P0_STYLE_IDS.bordered },
            1: { v: 1234, s: P0_STYLE_IDS.number },
            2: { v: 1234.5678, s: P0_STYLE_IDS.number },
            3: { v: 0.2567, s: P0_STYLE_IDS.percent },
            4: { v: 45658, s: P0_STYLE_IDS.date },
          },
          3: {
            0: { v: '华南', s: P0_STYLE_IDS.fill },
            1: { v: 567, s: P0_STYLE_IDS.number },
            2: { v: 89.5, s: P0_STYLE_IDS.number },
            3: { v: 0.118, s: P0_STYLE_IDS.percent },
            4: { v: 45689, s: P0_STYLE_IDS.date },
          },
          4: {
            0: { v: '这是自动换行 + 顶端对齐的长文本，用于验证 wrap/vt 是否生效', s: P0_STYLE_IDS.wrapped },
            1: { v: '旋转45°', s: P0_STYLE_IDS.rotated },
            2: { v: '粗斜体', s: P0_STYLE_IDS.boldItalic },
          },
          6: {
            0: { v: '合计', s: P0_STYLE_IDS.header },
            1: { f: '=SUM(B3:B4)', v: 1801, s: P0_STYLE_IDS.number },
            2: { f: '=SUM(C3:C4)', v: 1324.0678, s: P0_STYLE_IDS.number },
          },
        },
        mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 4 }],
      },
    },
  };
}
