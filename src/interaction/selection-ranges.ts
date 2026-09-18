/**
 * **不连续多区域选区**（Ctrl+点选）的纯函数工具箱。
 *
 * 用户要求："请实现不连续多区域（Ctrl+点选）"。
 *
 * 事实边界（先读代码再动手，不猜）：
 *  - Univer 的选区模型本身**支持多个区域**：`SheetsSelectionsService.getCurrentSelections()`
 *    返回一个数组，UI 层 `_checkClearPreviousControls()` 在按住 Ctrl/Shift 时**不清掉**已有的选区框，
 *    门面（Facade）也给了 `FWorksheet.getSelection().getActiveRangeList()` 拿全部区域。
 *    所以"能不能选"不用我们造轮子 —— 我们要做的是**让我们自己的功能读懂多块区域**
 *    （以前所有地方都只读 `getActiveRange()` 这一块）。
 *  - 反过来，凡是"天然只有一方"的操作（与工作区条目互换、粘贴、拖动搬运）**只认单块**：
 *    多块时给出明确提示（`describeSingleBlockOnly`），而不是随便挑一块执行。
 *
 * 本模块只做算术与文案，不碰 Univer / DOM，因此可以在 node 环境下逐条验证。
 */
import { parseA1Range } from '../importer/table-style';

/** 一个矩形（0-based，含端点） */
export interface Rect {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** 列号 → 字母（与 `src/workspace/snapshot.ts` 的同名逻辑一致，这里独立一份以免互相牵制） */
export function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** 矩形 → A1 记号（单格给 `B2`，否则给 `B2:D5`） */
export function rectToA1(rect: Rect): string {
  const start = `${colLetter(rect.startCol)}${rect.startRow + 1}`;
  const end = `${colLetter(rect.endCol)}${rect.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/** 面积（格子数） */
export function rectArea(rect: Rect): number {
  if (rect.endRow < rect.startRow || rect.endCol < rect.startCol) return 0;
  return (rect.endRow - rect.startRow + 1) * (rect.endCol - rect.startCol + 1);
}

/** 两块区域是否相交（用于去重/判定"同一块"） */
export function rectsIntersect(a: Rect, b: Rect): boolean {
  return a.startRow <= b.endRow && b.startRow <= a.endRow && a.startCol <= b.endCol && b.startCol <= a.endCol;
}

/** 两个矩形是否完全相等 */
export function rectsEqual(a: Rect, b: Rect): boolean {
  return a.startRow === b.startRow && a.startCol === b.startCol && a.endRow === b.endRow && a.endCol === b.endCol;
}

/** 按行优先给区域排序（状态栏、日志、快照顺序都用它，保证可复现） */
export function sortRects(rects: Rect[]): Rect[] {
  return [...rects].sort(
    (a, b) => a.startRow - b.startRow || a.startCol - b.startCol || a.endRow - b.endRow || a.endCol - b.endCol,
  );
}

/**
 * 把 A1 记号数组收敛成"规范化"的区域表：解析失败丢弃、**被包含的小块丢掉**、
 * 完全相同的去重、最后按行优先排序。
 *
 * 为什么要"丢掉被包含的小块"：Ctrl+点选时用户可能先框一大片、又点里面某一格，
 * 两块内容重叠会让"搬进工作区"出现重复条目；按"大块吃掉小块"处理最符合直觉。
 */
export function normalizeRanges(a1List: readonly string[]): Rect[] {
  const parsed: Rect[] = [];
  for (const a1 of a1List) {
    const rect = parseA1Range(a1 ?? '');
    if (!rect) continue;
    parsed.push({ startRow: rect.startRow, startCol: rect.startCol, endRow: rect.endRow, endCol: rect.endCol });
  }
  const sorted = sortRects(parsed).sort((a, b) => rectArea(b) - rectArea(a) || a.startRow - b.startRow || a.startCol - b.startCol);
  const kept: Rect[] = [];
  for (const rect of sorted) {
    const covered = kept.some((big) => containsRect(big, rect));
    if (!covered) kept.push(rect);
  }
  return sortRects(kept);
}

/** `outer` 是否完全包含 `inner` */
export function containsRect(outer: Rect, inner: Rect): boolean {
  return (
    outer.startRow <= inner.startRow &&
    outer.endRow >= inner.endRow &&
    outer.startCol <= inner.startCol &&
    outer.endCol >= inner.endCol
  );
}

/** 区域表 → A1 记号数组 */
export function rectsToA1(rects: readonly Rect[]): string[] {
  return rects.map((rect) => rectToA1(rect));
}

/** 总格子数 */
export function totalCells(rects: readonly Rect[]): number {
  return rects.reduce((sum, rect) => sum + rectArea(rect), 0);
}

/** 包住所有区域的最小矩形（没有区域时返回 null） */
export function boundingRect(rects: readonly Rect[]): Rect | null {
  if (rects.length === 0) return null;
  return rects.reduce<Rect>(
    (acc, rect) => ({
      startRow: Math.min(acc.startRow, rect.startRow),
      startCol: Math.min(acc.startCol, rect.startCol),
      endRow: Math.max(acc.endRow, rect.endRow),
      endCol: Math.max(acc.endCol, rect.endCol),
    }),
    { ...rects[0] },
  );
}

/**
 * 状态栏用的选区文案。
 *
 * 单块时保持原样（`B2:D5`）——既有习惯不破坏；多块时给"块数 + 总格数 + 前几块"，
 * 否则一屏放不下（用户可能 Ctrl 点十几块）。超过 3 块只列前 3 块并给"…"。
 */
export function summarizeSelection(a1List: readonly string[], maxListed = 3): string {
  const rects = normalizeRanges(a1List);
  if (rects.length === 0) return '—';
  if (rects.length === 1) return rectToA1(rects[0]);
  const listed = rectsToA1(rects);
  const shown = listed.slice(0, maxListed).join(' + ');
  const more = listed.length > maxListed ? ` +…` : '';
  return `${shown}${more}（${rects.length} 块 / ${totalCells(rects)} 格）`;
}

/**
 * 多块区域时"只支持单块的操作"的拒绝文案（互换、粘贴、拖动搬运这类天然只有一方的动作）。
 * 返回 `null` 表示只有一块，可以照常执行。
 */
export function describeSingleBlockOnly(a1List: readonly string[], action: string): string | null {
  const rects = normalizeRanges(a1List);
  if (rects.length <= 1) return null;
  return `已选 ${rects.length} 块区域：${action}一次只能处理一块，请先只选一块（Ctrl+点选可再去掉多选）`;
}

/**
 * 工作区"快速导入"输入框的解析：支持**多块**写法，用空格/逗号/顿号分隔：
 * `A1:B2 D4:E5`、`A1,B2`、`A1、B2`。返回规范化后的区域表（非法片段由调用方报错）。
 */
export function parseRangeList(text: string): { rects: Rect[]; invalid: string[] } {
  const tokens = (text ?? '')
    .split(/[\s,，、;；]+/)
    .map((token) => token.trim())
    .filter(Boolean);
  const rects: Rect[] = [];
  const invalid: string[] = [];
  for (const token of tokens) {
    const rect = parseA1Range(token);
    if (!rect) {
      invalid.push(token);
      continue;
    }
    rects.push({ startRow: rect.startRow, startCol: rect.startCol, endRow: rect.endRow, endCol: rect.endCol });
  }
  return { rects: normalizeRanges(rectsToA1(rects)), invalid };
}
