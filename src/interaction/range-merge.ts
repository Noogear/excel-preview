/**
 * 选区与**合并单元格**的对齐（纯函数，可单测）。
 *
 * 背景（用户实测反馈）：在课程表这类满是合并单元格的表里，拖选 B4:F8 之后，Univer 会把**模型里的选区
 * 扩成整块**（例如 B4:F11，因为第 9~11 行有竖向合并），于是"右键 → 剪切到工作区"剪走的是 B4:F11，
 * 而不是用户以为的 B4:F8 —— 用户的原话是"剪切的内容并非是我选中的那部分"。
 *
 * 这里的原则很简单：**被操作的范围必须等于高亮显示的范围**。
 * 所以先按合并块把范围补齐（补齐到不再与任何合并块相交为止），再把选区同步成补齐后的范围，
 * 并明确告诉用户"已按整块处理"。
 */

export interface Rect {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

/** 两个矩形是否相交（含边界接触） */
export function rectsIntersect(a: Rect, b: Rect): boolean {
  return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startColumn <= b.endColumn && a.endColumn >= b.startColumn;
}

/**
 * 把 `rect` 扩到"不与任何合并块相交"的最小范围。
 *
 * 为什么要循环：扩大之后可能又压到新的合并块（阶梯状排布的合并很常见），
 * 所以要反复扩，直到一轮下来没有任何变化。`maxPasses` 是防呆上限。
 */
export function expandRectToMerges(
  rect: Rect,
  merges: readonly Rect[],
  maxPasses = 16,
): { rect: Rect; expanded: boolean } {
  let current: Rect = { ...rect };
  let expanded = false;
  for (let pass = 0; pass < maxPasses; pass += 1) {
    let changed = false;
    for (const merge of merges) {
      if (!rectsIntersect(current, merge)) continue;
      const next: Rect = {
        startRow: Math.min(current.startRow, merge.startRow),
        startColumn: Math.min(current.startColumn, merge.startColumn),
        endRow: Math.max(current.endRow, merge.endRow),
        endColumn: Math.max(current.endColumn, merge.endColumn),
      };
      if (
        next.startRow !== current.startRow ||
        next.startColumn !== current.startColumn ||
        next.endRow !== current.endRow ||
        next.endColumn !== current.endColumn
      ) {
        current = next;
        changed = true;
        expanded = true;
      }
    }
    if (!changed) break;
  }
  return { rect: current, expanded };
}

/** 选区的行列尺寸 */
export function rectSize(rect: Rect): { rows: number; cols: number } {
  return {
    rows: Math.max(0, rect.endRow - rect.startRow + 1),
    cols: Math.max(0, rect.endColumn - rect.startColumn + 1),
  };
}
