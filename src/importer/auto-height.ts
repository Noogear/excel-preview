/**
 * 自动换行的行高估算。
 *
 * 背景：Excel 里"自动换行 + 未锁定行高"的行会自动变高；而我们按 defaultRowHeight 渲染时，
 * 长文本会被裁切（P2 截图中肉眼可见）。Univer 的 `IRowData.ia/ah` 是"自适应"标记，
 * 但导入时并不会替我们算——所以这里给出**确定性估算**，把需要变高的行算出来。
 *
 * 估算规则（近似，够用且可单测）：
 *  - 半角字符按 0.55 em 计宽，全角/中日韩字符按 1.0 em 计宽
 *  - 每行可容纳的"字宽"= (列宽 - 内边距) / 字号像素宽
 *  - 行高 = 行数 × 字号 × 1.35 + 上下内边距
 *  - 有自定义行高（Excel 里手工设过）的行**不参与**自动变高，与 Excel 行为一致
 */

const PADDING_X = 8;
/**
 * 行高里额外的垂直余量。
 * 注意要小：Excel 的默认行高（11pt 字约 15pt=20px）**已经包含**了上下内边距，
 * 若再叠加 8px 会把"单行文本"也判成需要加高，导致整表行高被无谓撑大（实测踩过）。
 */
const PADDING_Y = 2;
const LINE_HEIGHT_FACTOR = 1.36;

/** 全角字符判定（CJK、全角标点、假名、韩文等） */
export function isWideChar(code: number): boolean {
  return (
    (code >= 0x1100 && code <= 0x115f) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7a3) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe4f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  );
}

/** 文本的"字宽"度量（以 em 为单位） */
export function measureTextWidth(text: string): number {
  let width = 0;
  for (const ch of text) {
    width += isWideChar(ch.codePointAt(0) ?? 0) ? 1 : 0.55;
  }
  return width;
}

/**
 * 估算一个单元格里文本需要的行数。
 * `usableWidthPx` 是单元格可用宽度（已扣内边距）。
 */
export function estimateLineCount(text: string, fontSizePt: number, usableWidthPx: number): number {
  const fontPx = Math.max(1, ptToPx(fontSizePt));
  const emPerLine = Math.max(1, usableWidthPx / fontPx);

  let lines = 0;
  for (const segment of text.split('\n')) {
    const width = measureTextWidth(segment);
    lines += Math.max(1, Math.ceil(width / emPerLine));
  }
  return Math.max(1, lines);
}

export function estimateWrappedRowHeight(text: string, fontSizePt: number, usableWidthPx: number): number {
  const lines = estimateLineCount(text, fontSizePt, usableWidthPx);
  const fontPx = ptToPx(fontSizePt);
  return Math.ceil(lines * fontPx * LINE_HEIGHT_FACTOR + PADDING_Y);
}

export function ptToPx(pt: number): number {
  return (pt * 96) / 72;
}

export interface AutoHeightCell {
  row: number;
  col: number;
  text: string;
  fontSizePt: number;
  /** 该单元格是否设置了自动换行 */
  wrap: boolean;
  /** 合并区宽度（多列合计，px）；未合并时等于所在列宽 */
  usableWidthPx: number;
}

export interface AutoHeightInput {
  cells: AutoHeightCell[];
  /** 已有明确高度的行（Excel 手工设过高度的行）→ 保持不动 */
  rowsWithCustomHeight: Set<number>;
  defaultRowHeightPx: number;
}

/**
 * 计算需要变高的行 → 行高（px）。只返回**需要变高**的行，
 * 且不会低于 defaultRowHeightPx。
 */
export function computeAutoRowHeights(input: AutoHeightInput): Record<number, number> {
  const result: Record<number, number> = {};

  for (const cell of input.cells) {
    if (!cell.wrap) continue;
    if (input.rowsWithCustomHeight.has(cell.row)) continue;
    if (!cell.text) continue;

    const needed = estimateWrappedRowHeight(cell.text, cell.fontSizePt, Math.max(1, cell.usableWidthPx - PADDING_X));
    if (needed <= input.defaultRowHeightPx) continue;

    const current = result[cell.row] ?? input.defaultRowHeightPx;
    if (needed > current) result[cell.row] = needed;
  }

  return result;
}
