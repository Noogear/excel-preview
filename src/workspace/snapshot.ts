/**
 * 快照提取与写回。
 *
 * 约定（产品约束的直接体现）：
 *  - 提取时：样式只用于**预览**（SnapshotStyle），不被当作可编辑对象；
 *  - 写回时：**只写值与公式**，绝不写样式 → 目标单元格原有格式原样保留。
 */
import { BooleanNumber, type ICellData, type IStyleData } from '@univerjs/core';
import type { FWorksheet } from '@univerjs/sheets/facade';

import type { RangeSnapshot, SnapshotCell, SnapshotStyle } from './types';

/** 单个快照允许的最大单元格数（防止误选整表导致内存爆掉） */
export const MAX_SNAPSHOT_CELLS = 4000;

let snapshotSeq = 0;

export interface ExtractResult {
  snapshot: RangeSnapshot | null;
  error?: string;
}

/** 一次操作最多产出多少个工作区条目（安全阀：避免"一键全搬"把超大表塞爆） */
export const MAX_WORKSPACE_ITEMS_PER_ACTION = 500;

/**
 * 「这个格子算不算空」——**唯一判定口径**（所有"搬进工作区"的渠道都走它）。
 *
 * 用户要求原文："默认跳过空单元格，并检查所有的转移单元格到工作区的渠道，确保**绝对**跳过空内容的单元格"。
 * 实测教训（用户座位表 + 编辑清空过的格子）——"空"远不止 `null` 一种形态：
 *  - 单元格从没写过 → `v` 是 `null`（旧口径能跳过）；
 *  - **用户把内容删光**（编辑器里全选删除/退格）→ Univer 存下来的是**空字符串 `''`**，
 *    旧口径只判 `=== null`，于是这种"看起来完全空白"的格子照样变成工作区里的一张白卡片；
 *  - 只打了空格/全角空格（`' '`、`'　'`，中文表格里比比皆是）→ 同样只有空白，视觉上就是空格子；
 *  - 富文本对象 → `normalizeValue` 取纯文本后可能只剩空白。
 *
 * 所以判定是：**没有公式**，且值不是"有意义的标量"（数字/布尔永远算有内容，
 * 因为 0 与 FALSE 都是真内容），且文本 `trim()` 之后为空。
 */
export function isEmptyContent(value: unknown, formula?: string | null): boolean {
  if (formula) return false; // 公式就是内容（哪怕它算出来是空串）
  if (value === null || value === undefined) return true;
  if (typeof value === 'number') return false; // 0 也是内容
  if (typeof value === 'boolean') return false; // false 也是内容
  return String(value).trim() === '';
}

/**
 * 整份快照是不是"空条目"（**所有**格子都没有内容）。
 *
 * 用途：① 单元格 ↔ 工作区条目互换后，换出来的是空格子就不该在工作区里留一张白卡片；
 * ② 恢复上一次会话时，把历史版本（旧口径）存下来的空条目顺手清掉——
 *    否则用户重开应用后看到的还是那些白卡片，"绝对跳过空内容"就不绝对了。
 */
export function isSnapshotEmpty(snapshot: Pick<RangeSnapshot, 'values' | 'formulas'>): boolean {
  const rows = snapshot.values?.length ?? 0;
  if (rows === 0) return true;
  for (let r = 0; r < rows; r += 1) {
    const row = snapshot.values[r];
    // 值矩阵与公式矩阵理论上等宽（同一块区域取出来的），这里按较宽的一侧遍历，
    // 脏数据（长度不一致）也不会被误判成"空"而删掉用户的公式
    const width = Math.max(row?.length ?? 0, snapshot.formulas?.[r]?.length ?? 0);
    for (let c = 0; c < width; c += 1) {
      if (!isEmptyContent(row?.[c], snapshot.formulas?.[r]?.[c])) return false;
    }
  }
  return true;
}

export function extractSnapshot(sheet: FWorksheet, a1?: string): ExtractResult {
  const range = a1 ? sheet.getRange(a1) : sheet.getActiveRange();
  if (!range) return { snapshot: null, error: '没有选中区域' };

  const rect = range.getRange();
  const rows = rect.endRow - rect.startRow + 1;
  const cols = rect.endColumn - rect.startColumn + 1;

  if (rows * cols > MAX_SNAPSHOT_CELLS) {
    return { snapshot: null, error: `选区过大（${rows}×${cols}），请选择 ${MAX_SNAPSHOT_CELLS} 个单元格以内的区域` };
  }

  const displays = range.getDisplayValues();
  // 优先取**原始值**：日期格式的单元格 raw 是序列号；getValues() 可能给 Date 对象
  const values = safeRawValues(range);
  const formulas = safeFormulas(range);

  const mergeLookup = buildMergeLookup(sheet, rect);

  const cells: SnapshotCell[][] = [];
  const plainValues: (string | number | boolean | null)[][] = [];
  const formulaMatrix: (string | null)[][] = [];

  for (let r = 0; r < rows; r++) {
    const cellRow: SnapshotCell[] = [];
    const valueRow: (string | number | boolean | null)[] = [];
    const formulaRow: (string | null)[] = [];

    for (let c = 0; c < cols; c++) {
      const absRow = rect.startRow + r;
      const absCol = rect.startColumn + c;
      const merge = mergeLookup.get(`${absRow}:${absCol}`);

      const rawValue = normalizeValue(values?.[r]?.[c]);
      const formula = normalizeFormula(formulas?.[r]?.[c]);
      const display = displays?.[r]?.[c] ?? '';

      const cell: SnapshotCell = {
        text: typeof display === 'string' ? display : String(display ?? ''),
        value: rawValue,
        style: toSnapshotStyle(readCellStyle(sheet, absRow, absCol)),
      };
      if (formula) cell.formula = formula;
      if (merge) {
        if (merge.covered) cell.covered = true;
        if (merge.rowSpan > 1 || merge.colSpan > 1) cell.merge = { rowSpan: merge.rowSpan, colSpan: merge.colSpan };
      }

      cellRow.push(cell);
      valueRow.push(rawValue);
      formulaRow.push(formula);
    }

    cells.push(cellRow);
    plainValues.push(valueRow);
    formulaMatrix.push(formulaRow);
  }

  snapshotSeq += 1;
  return {
    snapshot: {
      id: `snap-${Date.now()}-${snapshotSeq}`,
      source: {
        sheetId: sheet.getSheetId(),
        sheetName: sheet.getSheetName(),
        a1: rectToA1(rect),
        startRow: rect.startRow,
        startCol: rect.startColumn,
        endRow: rect.endRow,
        endCol: rect.endColumn,
      },
      rows,
      cols,
      cells,
      values: plainValues,
      formulas: formulaMatrix,
      createdAt: Date.now(),
      label: rectToA1(rect),
    },
  };
}

/**
 * 把一个区域**拆成一个个独立单元格快照**（工作区的唯一入口形态）。
 *
 * 用户要求（第 1 条）：工作区相当于一张小表格，放进去的单元格会变成一个个**独立单元格**，
 * 并且**保留它原本的尺寸**；所以无论复制/剪切/导入的是一个格还是一整片区域，
 * 最终都会拆成若干 1×1 的条目。
 *
 * 同时满足第 2 条：**底层默认跳过空内容单元格**——这里按行优先顺序遍历，
 * 只对"有内容"的格子产出条目（判定见 {@link isEmptyContent}：`null`、空字符串、
 * 只有空格的文本都算空；公式、数字 0、布尔 false 都算有内容；空格的样式不搬，因为工作区只管内容）。
 *
 * `max` 是安全阀：一次操作最多产出这么多条目（超大表一键全搬时避免内存与 UI 失控）。
 */
export function extractCellItems(
  sheet: FWorksheet,
  a1?: string,
  max = MAX_WORKSPACE_ITEMS_PER_ACTION,
): { items: RangeSnapshot[]; skippedEmpty: number; truncated: number; error?: string } {
  const range = a1 ? sheet.getRange(a1) : sheet.getActiveRange();
  if (!range) return { items: [], skippedEmpty: 0, truncated: 0, error: '没有选中区域' };

  const rect = range.getRange();
  const rows = rect.endRow - rect.startRow + 1;
  const cols = rect.endColumn - rect.startColumn + 1;

  // 取值口径与 extractSnapshot 一致：**优先 raw 值**。
  // 实测（用户反馈的日期 bug）：单格 Facade 的 `cell.getValue()` 对日期格返回的是
  // 格式化后的**字符串** "2025-01-01"，一旦存进工作区再写回，日期就永久退化成文本；
  // raw 值给的是序列号 45658，配合目标格自带的日期格式，写回仍是日期。
  // 顺带好处：整块只取一次，不再逐格调用取值 API（大范围一键盘点时明显更快）。
  const rawValues = safeRawValues(range);

  const items: RangeSnapshot[] = [];
  let skippedEmpty = 0;
  let truncated = 0;

  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const absRow = rect.startRow + r;
      const absCol = rect.startColumn + c;
      const cell = sheet.getRange(absRow, absCol);
      const rawCell = rawValues?.[r]?.[c];
      // raw 取不到（老版本 Facade）才退回单格取值
      const rawValue = normalizeValue(rawCell === undefined ? cell.getValue() : rawCell);
      const formula = normalizeFormula(safeCellFormula(cell));
      // 底层跳过空内容单元格：没公式、值也是空的（null / '' / 纯空格）→ 不进工作区
      if (isEmptyContent(rawValue, formula)) {
        skippedEmpty += 1;
        continue;
      }
      if (items.length >= max) {
        truncated += 1;
        continue;
      }

      const cellA1 = rectToA1({ startRow: absRow, startColumn: absCol, endRow: absRow, endColumn: absCol });
      const display = cell.getDisplayValue();
      const snapshotCell: SnapshotCell = {
        text: typeof display === 'string' ? display : String(display ?? ''),
        value: rawValue,
        style: toSnapshotStyle(readCellStyle(sheet, absRow, absCol)),
      };
      if (formula) snapshotCell.formula = formula;

      snapshotSeq += 1;
      items.push({
        id: `snap-${Date.now()}-${snapshotSeq}`,
        source: {
          sheetId: sheet.getSheetId(),
          sheetName: sheet.getSheetName(),
          a1: cellA1,
          startRow: absRow,
          startCol: absCol,
          endRow: absRow,
          endCol: absCol,
        },
        rows: 1,
        cols: 1,
        cells: [[snapshotCell]],
        values: [[rawValue]],
        formulas: [[formula]],
        createdAt: Date.now(),
        label: cellA1,
        // 保留来源单元格的尺寸：工作区按这个尺寸展示，方便"一眼认出是哪一格"
        cellSize: readCellSize(sheet, absRow, absCol),
      });
    }
  }

  return { items, skippedEmpty, truncated };
}

/**
 * 「加入工作区」面板的**实时预览**结果（用户要求：面板里别放没用的元素，要能看见会发生什么）。
 *
 * `total` / `empty` 都是**精确值**：`extractCellItems` 即使到了取样上限也会继续数下去
 * （超出的部分记在 `truncated`），所以 `total = items.length + truncated` 与真实加入数一致。
 */
export interface WorkspaceImportPreview {
  /** 块数（多块区域 Ctrl+点选时 > 1） */
  blocks: number;
  /** 会加入的单元格数（= 非空内容格子数） */
  total: number;
  /** 会被跳过的空内容格子数（上限 {@link PREVIEW_EMPTY_CAP}，超出按上限报） */
  empty: number;
  /** 单次上限（真实加入时最多这么多格） */
  limit: number;
  /** 是否超过单次上限（超了只会加入前 limit 格） */
  overLimit: boolean;
  /** 样例格子（最多 PREVIEW_SAMPLE_MAX 个，按行优先顺序） */
  samples: { a1: string; text: string }[];
}

/** 预览最多取多少个样例格子（只影响样例，不影响 total 的精确性） */
export const PREVIEW_SAMPLE_MAX = 12;
/**
 * "跳过多少空内容"的报数上限：选区可以写得很大（例如 `A1:A1048576`），
 * 空格子数会大到没有意义，超过这个数就不再往上加（界面上到上限就不显示这一项）。
 */
export const PREVIEW_EMPTY_CAP = 100_000;

/** 两个矩形（0 基、闭区间）的交集；没有交集返回 null。纯函数，导出供单测直接覆盖 */
export function clipRect(
  a: { startRow: number; startColumn: number; endRow: number; endColumn: number },
  b: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): { startRow: number; startColumn: number; endRow: number; endColumn: number } | null {
  const startRow = Math.max(a.startRow, b.startRow);
  const startColumn = Math.max(a.startColumn, b.startColumn);
  const endRow = Math.min(a.endRow, b.endRow);
  const endColumn = Math.min(a.endColumn, b.endColumn);
  if (endRow < startRow || endColumn < startColumn) return null;
  return { startRow, startColumn, endRow, endColumn };
}

/**
 * 预览一次"加入工作区"的结果：**只读表、不产生任何快照/历史**。
 *
 * 实现上复用 `extractCellItems`（唯一的空内容判定入口），但只扫**可能真有内容的那一小片**：
 * 先和"已用区域"求交集再逐格判定。为什么必须这么做（踩过）：
 * 面板是**边打字边预览**的，用户完全可以写下 `A1:A1048576` 这种写法 ——
 * 逐格扫描整列百万行会把主线程钉死（页面像卡死一样）；而"已用区域之外没有内容"是 `getDataRange()`
 * 的定义本身，所以先求交集**不改变 total**（空内容的报数仍然按原始选区面积算，见下面的 empty）。
 */
export function previewWorkspaceImport(sheet: FWorksheet, a1List: string[]): WorkspaceImportPreview {
  let total = 0;
  let empty = 0;
  const samples: { a1: string; text: string }[] = [];

  let data: { startRow: number; startColumn: number; endRow: number; endColumn: number };
  try {
    data = sheet.getDataRange().getRange();
  } catch {
    data = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 };
  }

  for (const a1 of a1List) {
    let rect: { startRow: number; startColumn: number; endRow: number; endColumn: number };
    try {
      rect = sheet.getRange(a1).getRange();
    } catch {
      continue; // 单块拿不到范围就当它没内容，不阻断整次预览
    }
    const area =
      (rect.endRow - rect.startRow + 1) * (rect.endColumn - rect.startColumn + 1);
    const scanRect = clipRect(rect, data);
    if (!scanRect || area <= 0) {
      empty = Math.min(PREVIEW_EMPTY_CAP, empty + Math.max(0, area)); // 交集为空 = 这一块全是空
      continue;
    }

    const remaining = Math.max(0, PREVIEW_SAMPLE_MAX - samples.length);
    const { items, truncated, error } = extractCellItems(sheet, rectToA1(scanRect), remaining);
    if (error) continue;
    const nonEmpty = items.length + truncated;
    total += nonEmpty;
    // 空内容的报数仍按**用户写的那片区域**算（选区里有多少格是空的），只是不会去逐格扫它
    empty = Math.min(PREVIEW_EMPTY_CAP, empty + Math.max(0, area - nonEmpty));
    for (const item of items) {
      if (samples.length >= PREVIEW_SAMPLE_MAX) break;
      samples.push({ a1: item.source.a1, text: item.cells?.[0]?.[0]?.text ?? '' });
    }
  }

  return {
    blocks: a1List.length,
    total,
    empty,
    limit: MAX_WORKSPACE_ITEMS_PER_ACTION,
    overLimit: total > MAX_WORKSPACE_ITEMS_PER_ACTION,
    samples,
  };
}

/** 读单元格在表里的像素尺寸（拿不到就交给调用方用默认值） */
function readCellSize(sheet: FWorksheet, row: number, col: number): { width: number; height: number } | undefined {
  try {
    const width = sheet.getColumnWidth(col);
    const height = sheet.getRowHeight(row);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return undefined;
    return { width: Math.round(width), height: Math.round(height) };
  } catch {
    return undefined;
  }
}

/** 单元格公式的容错读取（不同版本 Facade 可能没有 getFormula） */
function safeCellFormula(cell: { getFormula?: () => string | null | undefined }): string | null | undefined {
  try {
    return cell.getFormula?.() ?? null;
  } catch {
    return null;
  }
}

/**
 * 把快照写回表格的指定位置。
 * 只写值与公式（f/v），不写 s → 目标单元格格式不受影响。
 */
export function applySnapshot(sheet: FWorksheet, snapshot: RangeSnapshot, target: { row: number; col: number }): void {
  const matrix = buildValueMatrix(snapshot);
  sheet.getRange(target.row, target.col, snapshot.rows, snapshot.cols).setValues(matrix as never);
}

/**
 * "清空内容但保留格式"的单元格表示。
 *
 * 关键坑（实测）：在 `SetRangeValuesMutation` 里传 `null` 表示**整格删除**
 * （Univer 内部走 `realDeleteValue`，会把样式一起删掉），
 * 因此清空内容必须传一个"只把内容字段置空"的对象，让样式 `s` 得以保留。
 */
export function emptyContentCell(): ICellData {
  return { v: null, f: null, si: null, p: null } as unknown as ICellData;
}

/** 构造"只含值与公式"的 ICellData 矩阵（供写回 / 互换共用） */
export function buildValueMatrix(snapshot: RangeSnapshot): ICellData[][] {
  const matrix: ICellData[][] = [];
  for (let r = 0; r < snapshot.rows; r++) {
    const row: ICellData[] = [];
    for (let c = 0; c < snapshot.cols; c++) {
      const formula = snapshot.formulas[r]?.[c] ?? null;
      const value = snapshot.values[r]?.[c] ?? null;
      if (formula) {
        const cell: ICellData = { f: `=${formula}` };
        if (value !== null) cell.v = value;
        row.push(cell);
      } else if (value === null) {
        row.push(emptyContentCell());
      } else {
        row.push({ v: value });
      }
    }
    matrix.push(row);
  }
  return matrix;
}

/** 从"值矩阵"构造快照（用于把 B 区域的值写进 A 区域时的对称操作） */
export function buildValueMatrixFromRect(sheet: FWorksheet, rect: {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}): ICellData[][] {
  const range = sheet.getRange(rect.startRow, rect.startColumn, rect.endRow - rect.startRow + 1, rect.endColumn - rect.startColumn + 1);
  // 与 extractSnapshot 保持同一套取值口径：优先 raw（日期=序列号），兜底 getValues()。
  // 走 getValues() 时日期会变成 Date 对象，normalizeValue 会把它转回序列号（不会退化成字符串）。
  const values = safeRawValues(range) ?? [];
  const formulas = safeFormulas(range);
  const matrix: ICellData[][] = [];
  for (let r = 0; r < values.length; r++) {
    const row: ICellData[] = [];
    for (let c = 0; c < (values[r]?.length ?? 0); c++) {
      const formula = normalizeFormula(formulas?.[r]?.[c]);
      const value = normalizeValue(values[r]?.[c]);
      if (formula) {
        const cell: ICellData = { f: `=${formula}` };
        if (value !== null) cell.v = value;
        row.push(cell);
      } else if (value === null) {
        row.push(emptyContentCell());
      } else {
        row.push({ v: value });
      }
    }
    matrix.push(row);
  }
  return matrix;
}

// ---------------------------------------------------------------- helpers

function readCellStyle(sheet: FWorksheet, row: number, col: number): IStyleData | null {
  try {
    return sheet.getRange(row, col).getCellStyleData('cell') ?? null;
  } catch {
    return null;
  }
}

function safeFormulas(range: ReturnType<FWorksheet['getRange']>): string[][] | null {
  try {
    return range.getFormulas() as string[][];
  } catch {
    return null;
  }
}

/**
 * 取**原始值矩阵**（拿不到就退回 `getValues()`）。
 *
 * 日期格式的单元格在这个 API 下是**序列号**（如 45658），而 `getValues()` 可能返回 `Date` 对象；
 * 用原始值可以彻底绕开"Date → 字符串"的降级路径（见 `dateToExcelSerial` 的注释）。
 */
function safeRawValues(range: ReturnType<FWorksheet['getRange']>): unknown[][] | null {
  try {
    const raw = (range as unknown as { getRawValues?: () => unknown[][] }).getRawValues?.();
    if (Array.isArray(raw)) return raw;
  } catch {
    /* 版本差异：退回 getValues() */
  }
  try {
    return range.getValues() as unknown[][];
  } catch {
    return null;
  }
}

/**
 * 日期 → Excel 序列号（1900 日期系统，基准 1899-12-30）。
 *
 * 为什么需要它：Facade 的 `getValues()` 对**日期格式**的单元格会返回 `Date` 对象，
 * 而早先的归一化会走到"对象分支"→ `String(date)`，写出
 * `"Wed Jan 01 2025 08:00:00 GMT+0800 (中国标准时间)"` 这种字符串——
 * 结果①日期被毁成文本、②Univer 认为"文本格里装了怪值"，弹出
 * `sheets-ui.info.forceStringInfo`（"已强制按文本处理"）提醒（用户实测反馈）。
 * 写回**序列号**才是对的：单元格自带日期格式，显示仍然是日期，类型也不会退化。
 *
 * 用日期的**本地**年月日算（Excel 序列号是日历概念，与时区无关），避免时区导致差一天。
 */
export function dateToExcelSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round(utc / 86_400_000) + 25_569; // 1970-01-01 的 Excel 序列号是 25569
}

function normalizeValue(value: unknown): string | number | boolean | null {
  if (value === undefined || value === null) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  // 日期：转回序列号，保住"日期"语义（见 dateToExcelSerial 的注释）
  if (value instanceof Date) return dateToExcelSerial(value);
  // 富文本等对象：取纯文本
  const asRecord = value as { toPlainText?: () => string; text?: string };
  if (typeof asRecord.toPlainText === 'function') return asRecord.toPlainText();
  return String(value);
}

function normalizeFormula(value: unknown): string | null {
  if (typeof value !== 'string' || value.length === 0) return null;
  return value.startsWith('=') ? value.slice(1) : value;
}

function toSnapshotStyle(style: IStyleData | null): SnapshotStyle | undefined {
  if (!style) return undefined;
  const out: SnapshotStyle = {};
  let has = false;

  if (style.ff) {
    out.fontFamily = style.ff;
    has = true;
  }
  if (typeof style.fs === 'number') {
    out.fontSize = style.fs;
    has = true;
  }
  if (style.bl === BooleanNumber.TRUE) {
    out.bold = true;
    has = true;
  }
  if (style.it === BooleanNumber.TRUE) {
    out.italic = true;
    has = true;
  }
  if (style.ul && typeof style.ul === 'object') {
    out.underline = true;
    has = true;
  }
  if (style.st && typeof style.st === 'object') {
    out.strikeThrough = true;
    has = true;
  }
  if (style.cl?.rgb) {
    out.color = style.cl.rgb;
    has = true;
  }
  if (style.bg?.rgb) {
    out.fill = style.bg.rgb;
    has = true;
  }
  if (style.ht === 1 || style.ht === 2 || style.ht === 3) {
    out.align = style.ht === 1 ? 'left' : style.ht === 2 ? 'center' : 'right';
    has = true;
  }
  if (style.vt === 1 || style.vt === 2 || style.vt === 3) {
    out.vAlign = style.vt === 1 ? 'top' : style.vt === 2 ? 'middle' : 'bottom';
    has = true;
  }
  if (style.tb === 3) {
    out.wrap = true;
    has = true;
  }
  if (style.tr && typeof style.tr === 'object' && typeof style.tr.a === 'number' && style.tr.a !== 0) {
    out.rotate = style.tr.a;
    has = true;
  }
  if (style.n?.pattern) {
    out.numberFormat = style.n.pattern;
    has = true;
  }
  if (style.bd) {
    const border: NonNullable<SnapshotStyle['border']> = {};
    if (style.bd.t?.cl?.rgb) border.top = style.bd.t.cl.rgb;
    if (style.bd.r?.cl?.rgb) border.right = style.bd.r.cl.rgb;
    if (style.bd.b?.cl?.rgb) border.bottom = style.bd.b.cl.rgb;
    if (style.bd.l?.cl?.rgb) border.left = style.bd.l.cl.rgb;
    if (Object.keys(border).length > 0) {
      out.border = border;
      has = true;
    }
  }

  return has ? out : undefined;
}

interface MergeInfo {
  covered?: boolean;
  rowSpan: number;
  colSpan: number;
}

function buildMergeLookup(
  sheet: FWorksheet,
  rect: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): Map<string, MergeInfo> {
  const lookup = new Map<string, MergeInfo>();
  let merges: Array<{ startRow: number; startColumn: number; endRow: number; endColumn: number }> = [];
  try {
    const internal = sheet.getSheet() as unknown as { getMergeData?: () => typeof merges };
    merges = internal.getMergeData?.() ?? [];
  } catch {
    merges = [];
  }

  for (const merge of merges) {
    const inRect =
      merge.startRow <= rect.endRow &&
      merge.endRow >= rect.startRow &&
      merge.startColumn <= rect.endColumn &&
      merge.endColumn >= rect.startColumn;
    if (!inRect) continue;

    const rowSpan = merge.endRow - merge.startRow + 1;
    const colSpan = merge.endColumn - merge.startColumn + 1;
    for (let r = merge.startRow; r <= merge.endRow; r++) {
      for (let c = merge.startColumn; c <= merge.endColumn; c++) {
        lookup.set(`${r}:${c}`, {
          covered: !(r === merge.startRow && c === merge.startColumn),
          rowSpan: r === merge.startRow && c === merge.startColumn ? rowSpan : 1,
          colSpan: r === merge.startRow && c === merge.startColumn ? colSpan : 1,
        });
      }
    }
  }
  return lookup;
}

export function rectToA1(rect: { startRow: number; startColumn: number; endRow: number; endColumn: number }): string {
  const start = `${colToLetter(rect.startColumn)}${rect.startRow + 1}`;
  const end = `${colToLetter(rect.endColumn)}${rect.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

export function colToLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}
