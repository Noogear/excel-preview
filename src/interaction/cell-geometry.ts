/**
 * 像素 ↔ 单元格 的**自有**换算（正向在 `swap-flash.ts` 里，这里是反向）。
 *
 * 为什么不用 Univer 渲染服务的 `getCellWithCoordByOffset`：它的滚动换算与提醒框那个
 * 上游公式同源——`content = (css − scroll)/scale + scroll − header`，在 scale=1 时**把滚动量消掉了**。
 * 实测后果：表格滚过之后，拖动落点会解析成"没滚动时"的格子（用户反馈"拖拽起点和落点严重偏移"；
 * 实测：指针在 S74 上，app 解析成 D13）。起点之所以看着正常，是因为我们用的是**选区**而不是命中测试。
 *
 * 坐标约定（对着真实夹具实测校准过，别再凭直觉改）：
 *   - `columnWidthAccumulation[i]` / `rowHeightAccumulation[i]` 是**累计右/下边界**（含第 i 列/行，不含表头）；
 *     例如 `columnWidthAccumulation[18] = 1714` 表示"到第 18 列（S）为止的右边界是 1714"。
 *   - 因此"点在哪个格"要用**第一个 `accum[i] > 内容坐标` 的 i**（不是"最后一个 ≤"）。
 *   - 内容坐标 = 画布 CSS 像素 / 缩放 + 滚动量 − 表头（行头宽 / 列头高）。
 *   - 正向换算（格 → 像素）用 `getCellWithCoordByIndex(row, col, true)` 的 startX/endX（**含表头**），
 *     与这里的反向严格互逆：`css = (startX − scroll) × scale`。
 */

export interface GridMetrics {
  /** 每列/每行的**累计边界**（不含表头，含自身）：`columnOffsets[i]` = 第 i 列右边界 */
  columnOffsets: readonly number[];
  rowOffsets: readonly number[];
  /** 行头宽度 / 列头高度（内容坐标里表头占的偏移） */
  headerWidth: number;
  headerHeight: number;
  /** 视口滚动量（内容坐标单位） */
  scroll: { x: number; y: number };
  /** 渲染缩放 */
  scale: { x: number; y: number };
}

/**
 * 在"累计边界"数组里找 `value` 落在哪一段：
 * 返回**第一个满足 `offsets[i] > value`** 的下标（即该坐标所在的行/列）。
 * 超出末尾返回最后一格（与表格"点空白也落在最后一列"的行为一致）。
 */
export function indexAtOffset(offsets: readonly number[], value: number): number {
  if (offsets.length === 0) return -1;
  let low = 0;
  let high = offsets.length - 1;
  let answer = offsets.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (offsets[mid] > value) {
      answer = mid;
      high = mid - 1;
    } else {
      low = mid + 1;
    }
  }
  return answer;
}

/**
 * 画布内 CSS 像素 → 单元格。
 * 落在表头（行头/列头）里 → 不算单元格，返回 null。
 */
export function cellAtPoint(
  metrics: GridMetrics,
  canvasX: number,
  canvasY: number,
): { row: number; col: number } | null {
  const { x: scaleX, y: scaleY } = metrics.scale;
  if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY) || scaleX === 0 || scaleY === 0) return null;
  if (![canvasX, canvasY, metrics.scroll.x, metrics.scroll.y].every((value) => Number.isFinite(value))) return null;

  // 画布像素 → 内容坐标（去掉表头偏移）
  const contentX = canvasX / scaleX + metrics.scroll.x - metrics.headerWidth;
  const contentY = canvasY / scaleY + metrics.scroll.y - metrics.headerHeight;
  if (contentX < 0 || contentY < 0) return null; // 点在行头/列头上

  const col = indexAtOffset(metrics.columnOffsets, contentX);
  const row = indexAtOffset(metrics.rowOffsets, contentY);
  if (col < 0 || row < 0) return null;
  return { row, col };
}
