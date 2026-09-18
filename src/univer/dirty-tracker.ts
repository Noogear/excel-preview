/**
 * 全局"脏单元格"追踪。
 *
 * 为什么需要全局：多标签页下每张工作表会各自安装一次约束锁，
 * 若把脏集合挂在锁实例上，切表/切标签就会丢记录 —— 而外科式导出必须知道**所有**被改过的格子。
 * 因此这里按键 `workbookId → sheetId → "row:col"` 汇总，与锁的生命周期解耦。
 */

export type DirtyKey = string;

const dirtyByWorkbook = new Map<string, Map<string, Set<DirtyKey>>>();

export function recordDirtyCell(workbookId: string, sheetId: string, row: number, col: number): void {
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
