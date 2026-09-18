/**
 * 导入适配层：ParsedWorkbook（中立模型）→ Univer IWorkbookData。
 *
 * 这一层负责所有"语义/单位"转换，是保真度的关键位置：
 *  - Excel 列宽（字符）→ 像素；行高（磅）→ 像素
 *  - OOXML 边框线型 → BorderStyleTypes
 *  - 对齐/换行/旋转 → Univer 枚举
 *  - 缩进：Univer 的 IStyleData 没有 indent 字段，用 pd.l（左内边距）近似（记入 warnings）
 *  - 公式：默认采用 Excel 缓存值渲染，保证与 Excel 显示一致
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  type ICellData,
  type IColumnData,
  type IRowData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
  LocaleType,
} from '@univerjs/core';

import type { ParsedBorder, ParsedMerge, ParsedReport, ParsedStyle, ParsedWorkbook, ParsedSheet } from '../parser/types';
import { computeAutoRowHeights, type AutoHeightCell } from './auto-height';
import { buildTableStylePatches, type StylePatch } from './table-style';

/** Excel 列宽(字符) → 像素（Calibri 11 的经验公式，P1 需按字体度量校准） */
export function excelColWidthToPx(width: number): number {
  if (width <= 0) return 0;
  return Math.round(width * 7 + 5);
}

/** Excel 行高(磅) → 像素 */
export function excelRowHeightToPx(points: number): number {
  return Math.round(points * (96 / 72));
}

const BORDER_STYLE_MAP: Record<ParsedBorder['style'], BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
  dashed: BorderStyleTypes.DASHED,
  dotted: BorderStyleTypes.DOTTED,
  double: BorderStyleTypes.DOUBLE,
  hair: BorderStyleTypes.HAIR,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
};

const H_ALIGN_MAP = { left: 1, center: 2, right: 3 } as const; // HorizontalAlign.LEFT/CENTER/RIGHT
const V_ALIGN_MAP = { top: 1, middle: 2, bottom: 3 } as const; // VerticalAlign.TOP/MIDDLE/BOTTOM
const WRAP = 3; // WrapStrategy.WRAP

export interface ImportOutcome {
  workbookData: IWorkbookData;
  report: ParsedReport;
  /** 样式索引 → Univer 样式 id */
  styleIdByIndex: string[];
}

export interface ImportOptions {
  workbookId?: string;
  name?: string;
  /** 行/列下限，避免空表在 Univer 里太局促 */
  minRows?: number;
  minCols?: number;
}

export function toUniverWorkbook(parsed: ParsedWorkbook, options: ImportOptions = {}): ImportOutcome {
  const warnings = [...parsed.report.warnings];
  const minRows = options.minRows ?? 100;
  const minCols = options.minCols ?? 26;

  const { styles, styleIdByIndex } = convertStyles(parsed.styles, warnings);

  const sheets: IWorkbookData['sheets'] = {};
  const sheetOrder: string[] = [];
  const styleCtx: StyleContext = {
    styles,
    styleIdByIndex,
    parsedStyles: parsed.styles,
    themeColors: parsed.themeColors,
  };

  for (const sheet of parsed.sheets) {
    const data = convertSheet(sheet, styleCtx, minRows, minCols, warnings);
    sheets[data.id] = data;
    sheetOrder.push(data.id);
  }

  const workbookId = options.workbookId ?? `imported-${Date.now()}`;

  return {
    workbookData: {
      id: workbookId,
      name: options.name ?? '导入的工作簿',
      appVersion: '0.25.1',
      locale: LocaleType.ZH_CN,
      sheetOrder,
      styles,
      sheets,
    } as IWorkbookData,
    report: {
      unsupported: parsed.report.unsupported,
      warnings,
      // "已解析并原样保留、不影响预览"的信息（打印设置）——不是缺陷，UI 不该报警
      preserved: parsed.report.preserved ?? [],
    },
    styleIdByIndex,
  };
}

function convertStyles(
  parsedStyles: ParsedStyle[],
  warnings: string[],
): { styles: Record<string, IStyleData>; styleIdByIndex: string[] } {
  const styles: Record<string, IStyleData> = {};
  const styleIdByIndex: string[] = [];
  let indentSeen = 0;
  let rotationSeen = 0;

  parsedStyles.forEach((style, index) => {
    const converted = convertStyle(style, () => {
      indentSeen += 1;
    }, () => {
      rotationSeen += 1;
    });
    if (!converted) {
      styleIdByIndex[index] = '';
      return;
    }
    const id = `xlsx-style-${index}`;
    styles[id] = converted;
    styleIdByIndex[index] = id;
  });

  if (indentSeen > 0) warnings.push(`缩进(alignment.indent)出现 ${indentSeen} 次，已用左内边距近似，P1 需精确化`);
  if (rotationSeen > 0) warnings.push(`文字旋转出现 ${rotationSeen} 次，负角度/255(竖排)语义待校准`);
  return { styles, styleIdByIndex };
}

function convertStyle(style: ParsedStyle, onIndent: () => void, onRotation: () => void): IStyleData | null {
  const result: IStyleData = {};
  let hasAny = false;

  if (style.fontFamily) {
    result.ff = style.fontFamily;
    hasAny = true;
  }
  if (style.fontSize !== undefined) {
    result.fs = style.fontSize;
    hasAny = true;
  }
  if (style.bold) {
    result.bl = BooleanNumber.TRUE;
    hasAny = true;
  }
  if (style.italic) {
    result.it = BooleanNumber.TRUE;
    hasAny = true;
  }
  if (style.underline) {
    result.ul = { s: BooleanNumber.TRUE };
    hasAny = true;
  }
  if (style.strikeThrough) {
    result.st = { s: BooleanNumber.TRUE };
    hasAny = true;
  }
  if (style.color) {
    result.cl = { rgb: style.color };
    hasAny = true;
  }
  if (style.fill) {
    result.bg = { rgb: style.fill };
    hasAny = true;
  }
  if (style.border) {
    const bd: NonNullable<IStyleData['bd']> = {};
    const sides = ['top', 'bottom', 'left', 'right'] as const;
    const keys = { top: 't', bottom: 'b', left: 'l', right: 'r' } as const;
    for (const side of sides) {
      const line = style.border[side];
      if (!line) continue;
      bd[keys[side]] = { s: BORDER_STYLE_MAP[line.style] ?? BorderStyleTypes.THIN, cl: { rgb: line.color ?? '#000000' } };
      hasAny = true;
    }
    if (hasAny) result.bd = bd;
  }
  if (style.horizontalAlign) {
    result.ht = H_ALIGN_MAP[style.horizontalAlign] as IStyleData['ht'];
    hasAny = true;
  }
  if (style.verticalAlign) {
    result.vt = V_ALIGN_MAP[style.verticalAlign] as IStyleData['vt'];
    hasAny = true;
  }
  if (style.textWrap) {
    result.tb = WRAP as IStyleData['tb'];
    hasAny = true;
  }
  if (style.indent && style.indent > 0) {
    onIndent();
    result.pd = { l: style.indent * 8 };
    hasAny = true;
  }
  if (style.textRotation && style.textRotation !== 0) {
    onRotation();
    const raw = style.textRotation;
    // OOXML: 1..90 为逆时针角度；91..180 表示 (angle-90) 顺时针；255 为竖排
    let angle = raw;
    if (raw === 255) angle = 90;
    else if (raw > 90) angle = raw - 90;
    result.tr = { a: angle, v: BooleanNumber.TRUE };
    hasAny = true;
  }
  if (style.numberFormat) {
    result.n = { pattern: style.numberFormat };
    hasAny = true;
  }

  return hasAny ? result : null;
}

interface StyleContext {
  /** Univer 样式表：id -> IStyleData */
  styles: Record<string, IStyleData>;
  /** ParsedStyle 索引 -> Univer 样式 id */
  styleIdByIndex: string[];
  parsedStyles: ParsedStyle[];
  themeColors?: string[];
}

function convertSheet(
  sheet: ParsedSheet,
  ctx: StyleContext,
  minRows: number,
  minCols: number,
  warnings: string[],
): IWorksheetData {
  const { styleIdByIndex, parsedStyles } = ctx;
  const cellData: Record<number, Record<number, ICellData>> = {};
  let maxRow = 0;
  let maxCol = 0;
  let formulaCount = 0;
  let errorCount = 0;

  for (const cell of sheet.cells) {
    const styleId = cell.styleIndex !== undefined ? styleIdByIndex[cell.styleIndex] : undefined;
    const data: ICellData = {};
    if (styleId) data.s = styleId;

    if (cell.formula) {
      formulaCount += 1;
      data.f = `=${cell.formula}`;
      // 默认用 Excel 的缓存值渲染（保证与 Excel 显示一致）
      if (cell.value !== undefined && cell.value !== null) {
        data.v = cell.value;
        data.t = inferCellType(cell.value);
      }
    } else if (cell.error) {
      errorCount += 1;
      data.v = cell.error;
      data.t = CellValueType.STRING;
    } else if (cell.value !== undefined && cell.value !== null) {
      data.v = cell.value;
      data.t = inferCellType(cell.value);
    }

    if (Object.keys(data).length === 0) continue;
    (cellData[cell.row] ??= {})[cell.col] = data;
    maxRow = Math.max(maxRow, cell.row);
    maxCol = Math.max(maxCol, cell.col);
  }

  if (formulaCount > 0) warnings.push(`工作表「${sheet.name}」含 ${formulaCount} 个公式，采用 Excel 缓存值显示`);
  if (errorCount > 0) warnings.push(`工作表「${sheet.name}」含 ${errorCount} 个错误值单元格`);

  const rowData: Record<number, Partial<IRowData>> = {};
  for (const [rowKey, info] of Object.entries(sheet.rows)) {
    const row: Partial<IRowData> = {};
    if (info.height !== undefined) row.h = excelRowHeightToPx(info.height);
    if (info.hidden) row.hd = BooleanNumber.TRUE;
    if (Object.keys(row).length > 0) rowData[Number(rowKey)] = row;
  }

  const columnData: Record<number, Partial<IColumnData>> = {};
  for (const [colKey, info] of Object.entries(sheet.cols)) {
    const col: Partial<IColumnData> = {};
    if (info.width !== undefined) col.w = excelColWidthToPx(info.width);
    if (info.hidden) col.hd = BooleanNumber.TRUE;
    if (Object.keys(col).length > 0) columnData[Number(colKey)] = col;
  }

  // ---- 表格样式实体化：Excel 表格的斑马纹/表头样式不在单元格里，必须按样式名还原后写进单元格 ----
  if (sheet.tables && sheet.tables.length > 0) {
    const patches = buildTableStylePatches(sheet.tables, ctx.themeColors);
    applyStylePatches(patches, cellData, ctx);
    warnings.push(
      `工作表「${sheet.name}」含 ${sheet.tables.length} 个 Excel 表格，斑马纹/表头配色为算法近似（强调色取自主题）`,
    );
  }

  // ---- 自动换行的行高自适应 ----
  applyAutoRowHeights(sheet, cellData, rowData, columnData, parsedStyles, warnings);

  const dimensionRows = sheet.dimension ? sheet.dimension.endRow + 1 : 0;
  const dimensionCols = sheet.dimension ? sheet.dimension.endCol + 1 : 0;

  return {
    id: sheet.id,
    name: sheet.name,
    tabColor: '',
    hidden: sheet.hidden || sheet.veryHidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    freeze: sheet.freeze
      ? { xSplit: sheet.freeze.col, ySplit: sheet.freeze.row, startRow: sheet.freeze.row, startColumn: sheet.freeze.col }
      : { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
    rowCount: Math.max(maxRow + 1, dimensionRows, minRows),
    columnCount: Math.max(maxCol + 1, dimensionCols, minCols),
    zoomRatio: 1,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: sheet.defaultColWidth ? excelColWidthToPx(sheet.defaultColWidth) : 88,
    defaultRowHeight: sheet.defaultRowHeight ? excelRowHeightToPx(sheet.defaultRowHeight) : 24,
    mergeData: sheet.merges.map((merge) => ({
      startRow: merge.startRow,
      startColumn: merge.startCol,
      endRow: merge.endRow,
      endColumn: merge.endCol,
    })),
    cellData,
    rowData,
    columnData,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: sheet.gridlinesHidden ? BooleanNumber.FALSE : BooleanNumber.TRUE,
    rightToLeft: BooleanNumber.FALSE,
  } as IWorksheetData;
}

function inferCellType(value: string | number | boolean): CellValueType {
  if (typeof value === 'number') return CellValueType.NUMBER;
  if (typeof value === 'boolean') return CellValueType.BOOLEAN;
  return CellValueType.STRING;
}

/** 把表格样式补丁合并进单元格样式（基础样式来自该单元格原有样式） */
function applyStylePatches(
  patches: StylePatch[],
  cellData: Record<number, Record<number, ICellData>>,
  ctx: StyleContext,
): void {
  if (patches.length === 0) return;

  const idToIndex = new Map<string, number>();
  ctx.styleIdByIndex.forEach((id, index) => {
    if (id) idToIndex.set(id, index);
  });
  const memo = new Map<string, string>();

  for (const { row, col, patch } of patches) {
    const existing = cellData[row]?.[col];
    const baseIndex = existing?.s ? idToIndex.get(String(existing.s)) : undefined;
    const baseStyle = baseIndex !== undefined ? ctx.parsedStyles[baseIndex] : undefined;

    const key = `${baseIndex ?? -1}|${JSON.stringify(patch)}`;
    let styleId = memo.get(key);
    if (!styleId) {
      styleId = `xlsx-table-${baseIndex ?? 'none'}-${memo.size}`;
      const converted = convertStyle({ ...(baseStyle ?? {}), ...patch }, () => {}, () => {});
      if (converted) {
        ctx.styles[styleId] = converted;
        memo.set(key, styleId);
      } else {
        continue;
      }
    }

    const targetRow = (cellData[row] ??= {});
    targetRow[col] = { ...(existing ?? {}), s: styleId };
  }
}

/** 自动换行的行高自适应（近似估算，只增不减；Excel 手工设过高度的行保持不动） */
function applyAutoRowHeights(
  sheet: ParsedSheet,
  cellData: Record<number, Record<number, ICellData>>,
  rowData: Record<number, Partial<IRowData>>,
  columnData: Record<number, Partial<IColumnData>>,
  parsedStyles: ParsedStyle[],
  warnings: string[],
): void {
  const defaultRowHeightPx = sheet.defaultRowHeight ? excelRowHeightToPx(sheet.defaultRowHeight) : 24;
  const defaultColWidthPx = sheet.defaultColWidth ? excelColWidthToPx(sheet.defaultColWidth) : 88;

  const customRows = new Set<number>();
  for (const [rowKey, info] of Object.entries(sheet.rows)) {
    if (info.customHeight) customRows.add(Number(rowKey));
  }

  const mergeByAnchor = new Map<string, ParsedMerge>();
  for (const merge of sheet.merges) mergeByAnchor.set(`${merge.startRow}:${merge.startCol}`, merge);

  const cells: AutoHeightCell[] = [];
  for (const cell of sheet.cells) {
    const style = cell.styleIndex !== undefined ? parsedStyles[cell.styleIndex] : undefined;
    if (!style?.textWrap) continue;
    if (typeof cell.value !== 'string' || cell.value.length === 0) continue;

    let usableWidthPx = columnData[cell.col]?.w ?? defaultColWidthPx;
    const merge = mergeByAnchor.get(`${cell.row}:${cell.col}`);
    if (merge) {
      usableWidthPx = 0;
      for (let c = merge.startCol; c <= merge.endCol; c++) usableWidthPx += columnData[c]?.w ?? defaultColWidthPx;
    }

    cells.push({
      row: cell.row,
      col: cell.col,
      text: cell.value,
      fontSizePt: style.fontSize ?? 11,
      wrap: true,
      usableWidthPx,
    });
  }

  if (cells.length === 0) return;
  void cellData;

  const heights = computeAutoRowHeights({ cells, rowsWithCustomHeight: customRows, defaultRowHeightPx });
  let grown = 0;
  for (const [rowKey, height] of Object.entries(heights)) {
    const row = Number(rowKey);
    const existing = rowData[row]?.h;
    if (existing !== undefined && existing >= height) continue;
    rowData[row] = { ...(rowData[row] ?? {}), h: height };
    grown += 1;
  }

  if (grown > 0) {
    warnings.push(`工作表「${sheet.name}」有 ${grown} 行按自动换行估算了行高（近似值）`);
  }
}
