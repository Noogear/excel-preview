/**
 * 全局"脏单元格"追踪。
 *
 * 为什么需要全局：多标签页下每张工作表会各自安装一次约束锁，
 * 若把脏集合挂在锁实例上，切表/切标签就会丢记录 —— 而外科式导出必须知道**所有**被改过的格子。
 * 因此这里按键 `workbookId → sheetId → "row:col"` 汇总，与锁的生命周期解耦。
 */

export type DirtyKey = string;

const dirtyByWorkbook = new Map<string, Map<string, Set<DirtyKey>>>();

/**
 * 记账开关（默认开）。**应用特性时要关掉**。
 *
 * 为什么需要（真实缺陷）：导入时我们会在"可信代码路径"里应用条件格式 / 数据验证 / 超链接 /
 * 批注 / 图片，其中一些会写单元格（超链接是写进单元格的富文本）。这些写入同样会经过
 * `lock.ts` 的 mutation 包装 → 被记成"脏格"，于是**刚打开的文件立刻显示"未保存"**，
 * 而且在导出时被当成"用户改过的格子"去回写（本来可以原样保留的 XML 被重写）。
 * 暂停后：脏格集合只反映**用户真正改过的格子**，标签状态与外科式导出的范围都回到正确口径。
 *
 * 用**计数**而不是布尔：暂停/恢复可能嵌套（导入与冷标签重建都可能同时进行），
 * 计数器保证"谁暂停谁恢复"不会互相提前打开。
 */
let pauseDepth = 0;

export function pauseDirtyTracking(): void {
  pauseDepth += 1;
}

export function resumeDirtyTracking(): void {
  if (pauseDepth > 0) pauseDepth -= 1;
}

export function isDirtyTrackingPaused(): boolean {
  return pauseDepth > 0;
}

/**
 * 在"暂停记账"的区间里跑一段异步逻辑（异常也保证恢复）。
 * 应用特性用它，避免每处调用都写一遍 try/finally。
 */
export async function withDirtyTrackingPaused<T>(work: () => Promise<T>): Promise<T> {
  pauseDirtyTracking();
  try {
    return await work();
  } finally {
    resumeDirtyTracking();
  }
}

export function recordDirtyCell(workbookId: string, sheetId: string, row: number, col: number): void {
  if (pauseDepth > 0) return;
  let sheets = dirtyByWorkbook.get(workbookId);
  if (!sheets) {
    sheets = new Map<string, Set<DirtyKey>>();
    dirtyByWorkbook.set(workbookId, sheets);
  }
  let cells = sheets.get(sheetId);
  if (!cells) {
    cells = new Set<DirtyKey>();
    sheets.set(sheetId, cells);
  }
  cells.add(`${row}:${col}`);
}

/** 某个工作簿的全部脏单元格（sheetId → 坐标集合） */
export function getDirtyCells(workbookId: string): Map<string, Set<DirtyKey>> {
  return dirtyByWorkbook.get(workbookId) ?? new Map<string, Set<DirtyKey>>();
}

export function hasDirtyCells(workbookId: string): boolean {
  const sheets = dirtyByWorkbook.get(workbookId);
  if (!sheets) return false;
  for (const cells of sheets.values()) {
    if (cells.size > 0) return true;
  }
  return false;
}

export function dirtyCellCount(workbookId: string): number {
  const sheets = dirtyByWorkbook.get(workbookId);
  if (!sheets) return 0;
  let total = 0;
  for (const cells of sheets.values()) total += cells.size;
  return total;
}

/**
 * 展开成导出器需要的"每表一组的编辑坐标列表"。
 * 注意：`parsed.sheets[].id` 就是这里的 sheetId（也是 Univer 的 sheetId）。
 */
export function toSheetEditTargets(workbookId: string): Array<{ sheetId: string; cells: Array<{ row: number; col: number }> }> {
  const result: Array<{ sheetId: string; cells: Array<{ row: number; col: number }> }> = [];
  for (const [sheetId, keys] of getDirtyCells(workbookId)) {
    const cells: Array<{ row: number; col: number }> = [];
    for (const key of keys) {
      const [rowText, colText] = key.split(':');
      const row = Number(rowText);
      const col = Number(colText);
      if (Number.isFinite(row) && Number.isFinite(col)) cells.push({ row, col });
    }
    if (cells.length > 0) result.push({ sheetId, cells });
  }
  return result;
}

/** 关闭标签页时清掉该工作簿的记录（避免内存泄漏） */
export function clearWorkbookDirty(workbookId: string): void {
  dirtyByWorkbook.delete(workbookId);
}

/** 仅用于测试：清空全部 */
export function resetAllDirty(): void {
  dirtyByWorkbook.clear();
}
