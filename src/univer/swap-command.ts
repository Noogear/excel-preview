/**
 * 内容互换（swap）：把两个同尺寸区域的内容（值与公式）互相交换。
 *
 * 关键设计：**一次互换 = 一条命令 = 一次撤销**。
 * 如果直接连续调用两次 `setValues`，Univer 会记录两条独立的撤销记录，
 * 用户按一次 Ctrl+Z 只能还原一半，体验上就是个 bug。
 * 这里改为在一个命令的 handler 内用 `sequenceExecute` 执行两条 mutation，
 * Univer 的 UndoRedoService 会把"命令执行期间产生的 mutation"归并为同一条撤销记录。
 */
import {
  CommandType,
  ICommandService,
  type IAccessor,
  type ICellData,
  type ICommand,
  type IMutationInfo,
  type IObjectMatrixPrimitiveType,
  type IRange,
  IUndoRedoService,
  sequenceExecute,
} from '@univerjs/core';
import { SetRangeValuesMutation, type ISetRangeValuesMutationParams } from '@univerjs/sheets';
import type { FWorksheet } from '@univerjs/sheets/facade';

import { buildValueMatrixFromRect, emptyContentCell } from '../workspace/snapshot';

const CLEAR_COMMAND_ID = 'excel-preview.command.clear-range';

export interface ISwapRangesCommandParams {
  unitId: string;
  subUnitId: string;
  /** A 区域应写入的内容（来自 B） */
  matrixForA: ICellData[][];
  /** B 区域应写入的内容（来自 A） */
  matrixForB: ICellData[][];
  rangeA: IRange;
  rangeB: IRange;
}

export interface IClearRangeCommandParams {
  unitId: string;
  subUnitId: string;
  range: IRange;
  /** 清空前的原内容（供撤销写回）；尺寸必须与 range 一致 */
  snapshot: ICellData[][];
}

export const ClearRangeCommand: ICommand<IClearRangeCommandParams, boolean> = {
  id: CLEAR_COMMAND_ID,
  type: CommandType.COMMAND,
  handler(accessor: IAccessor, params?: IClearRangeCommandParams): boolean {
    if (!params) return false;
    const commandService = accessor.get(ICommandService);
    const undoRedoService = accessor.get(IUndoRedoService);
    const { unitId, subUnitId, range, snapshot } = params;

    const rows = range.endRow - range.startRow + 1;
    const cols = range.endColumn - range.startColumn + 1;
    const emptyMatrix: ICellData[][] = Array.from({ length: rows }, () =>
      Array.from({ length: cols }, () => emptyContentCell()),
    );

    // redo = 写空内容（只带 v/f/si/p，不带 s，因此样式原封不动）；undo = 写回原内容
    const redoMutations: IMutationInfo<ISetRangeValuesMutationParams>[] = [
      {
        id: SetRangeValuesMutation.id,
        type: CommandType.MUTATION,
        params: { unitId, subUnitId, cellValue: toSparseMatrix(emptyMatrix, range.startRow, range.startColumn) },
      },
    ];
    const undoMutations: IMutationInfo<ISetRangeValuesMutationParams>[] = [
      {
        id: SetRangeValuesMutation.id,
        type: CommandType.MUTATION,
        params: { unitId, subUnitId, cellValue: toSparseMatrix(snapshot, range.startRow, range.startColumn) },
      },
    ];

    const result = sequenceExecute(redoMutations, commandService);
    if (result.result) {
      undoRedoService.pushUndoRedo({ unitID: unitId, undoMutations, redoMutations, id: CLEAR_COMMAND_ID });
    }
    return result.result;
  },
};

const SWAP_COMMAND_ID = 'excel-preview.command.swap-ranges';

export const SwapRangesCommand: ICommand<ISwapRangesCommandParams, boolean> = {
  id: SWAP_COMMAND_ID,
  type: CommandType.COMMAND,
  handler(accessor: IAccessor, params?: ISwapRangesCommandParams): boolean {
    if (!params) return false;
    const commandService = accessor.get(ICommandService);
    const undoRedoService = accessor.get(IUndoRedoService);
    const { unitId, subUnitId, matrixForA, matrixForB, rangeA, rangeB } = params;

    // 关键事实：Univer 的撤销**不是自动记录**的——每个命令在执行成功后自己 pushUndoRedo。
    // 直接 sequenceExecute(mutation) 不会产生任何撤销记录（实测：撤销完全无效）。
    // 这里由命令自己构造 undo/redo 对，因此"一次互换 = 一条撤销记录"是结构性保证，
    // 不依赖任何批处理 API。
    const redoMutations: IMutationInfo<ISetRangeValuesMutationParams>[] = [
      { id: SetRangeValuesMutation.id, type: CommandType.MUTATION, params: { unitId, subUnitId, cellValue: toSparseMatrix(matrixForA, rangeA.startRow, rangeA.startColumn) } },
      { id: SetRangeValuesMutation.id, type: CommandType.MUTATION, params: { unitId, subUnitId, cellValue: toSparseMatrix(matrixForB, rangeB.startRow, rangeB.startColumn) } },
    ];
    const undoMutations: IMutationInfo<ISetRangeValuesMutationParams>[] = [
      { id: SetRangeValuesMutation.id, type: CommandType.MUTATION, params: { unitId, subUnitId, cellValue: toSparseMatrix(matrixForB, rangeA.startRow, rangeA.startColumn) } },
      { id: SetRangeValuesMutation.id, type: CommandType.MUTATION, params: { unitId, subUnitId, cellValue: toSparseMatrix(matrixForA, rangeB.startRow, rangeB.startColumn) } },
    ];

    const result = sequenceExecute(redoMutations, commandService);

    if (result.result) {
      undoRedoService.pushUndoRedo({
        unitID: unitId,
        undoMutations,
        redoMutations,
        id: SWAP_COMMAND_ID,
      });
    }

    return result.result;
  },
};

export interface SwapOutcome {
  ok: boolean;
  reason?: string;
}

/**
 * 注册互换命令（幂等）。
 * 必须在调用 `swapRanges`/`moveRange` 之前执行一次，否则 `executeCommand` 找不到命令 id。
 */
export function ensureSwapCommandRegistered(sheet: FWorksheet): void {
  const commandService = sheet.getInject().get(ICommandService);
  if (commandService.hasCommand(SWAP_COMMAND_ID)) return;
  commandService.registerCommand(SwapRangesCommand);
  commandService.registerCommand(ClearRangeCommand);
}

/** 业务入口：互换两个 A1 记号的区域（要求同尺寸） */
export function swapRanges(sheet: FWorksheet, a1: string, b1: string): SwapOutcome {
  const rangeA = sheet.getRange(a1);
  const rangeB = sheet.getRange(b1);
  const rectA = rangeA.getRange();
  const rectB = rangeB.getRange();

  const rowsA = rectA.endRow - rectA.startRow + 1;
  const colsA = rectA.endColumn - rectA.startColumn + 1;
  const rowsB = rectB.endRow - rectB.startRow + 1;
  const colsB = rectB.endColumn - rectB.startColumn + 1;

  if (rowsA !== rowsB || colsA !== colsB) {
    return { ok: false, reason: `尺寸不一致：${rowsA}×${colsA} 与 ${rowsB}×${colsB}，无法互换` };
  }

  const commandService = sheet.getInject().get(ICommandService);
  const unitId = sheet.getWorkbook().getUnitId();
  const subUnitId = sheet.getSheetId();

  void commandService.executeCommand<ISwapRangesCommandParams>(SWAP_COMMAND_ID, {
    unitId,
    subUnitId,
    matrixForA: buildValueMatrixFromRect(sheet, rectB),
    matrixForB: buildValueMatrixFromRect(sheet, rectA),
    rangeA: rectA,
    rangeB: rectB,
  });

  return { ok: true };
}

/**
 * 业务入口：清空区域**内容**、保留其格式。
 *
 * 为什么不用 `sheet.command.clear-selection-content`：①它依赖"当前选区"而不是右键的那个区域；
 * ②它被本项目的只读守卫默认拒绝（不在白名单里），执行结果是静默空操作。
 * 这里走自家的 ClearRangeCommand（守卫允许 `excel-preview.command.*`，且命令自带 pushUndoRedo）。
 * 空内容用 emptyContentCell()（{v:null,f:null,si:null,p:null}）——Univer 按 key 合并，
 * 不传 `s` 就不会碰样式；直接传 null 才是"整格删除"（会连样式一起丢），必须避免。
 */
export function clearRange(sheet: FWorksheet, a1: string): SwapOutcome {
  const rect = sheet.getRange(a1).getRange();
  const snapshot = buildValueMatrixFromRect(sheet, rect);

  const commandService = sheet.getInject().get(ICommandService);
  void commandService.executeCommand<IClearRangeCommandParams>(CLEAR_COMMAND_ID, {
    unitId: sheet.getWorkbook().getUnitId(),
    subUnitId: sheet.getSheetId(),
    range: rect,
    snapshot,
  });

  return { ok: true };
}

/** 业务入口：把源区域的内容"移动"到目标位置（源区域内容清空，但保留其格式） */
export function moveRange(sheet: FWorksheet, sourceA1: string, target: { row: number; col: number }): SwapOutcome {
  const source = sheet.getRange(sourceA1);
  const rectA = source.getRange();
  const rows = rectA.endRow - rectA.startRow + 1;
  const cols = rectA.endColumn - rectA.startColumn + 1;

  const sheetRows = sheet.getMaxRows();
  const sheetCols = sheet.getMaxColumns();
  if (target.row + rows > sheetRows || target.col + cols > sheetCols) {
    return { ok: false, reason: '目标位置超出工作表范围' };
  }
  if (target.row === rectA.startRow && target.col === rectA.startColumn) {
    return { ok: false, reason: '目标位置与源位置相同' };
  }

  const rectB: IRange = {
    startRow: target.row,
    startColumn: target.col,
    endRow: target.row + rows - 1,
    endColumn: target.col + cols - 1,
  };

  const emptyMatrix: ICellData[][] = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => emptyContentCell()),
  );

  const commandService = sheet.getInject().get(ICommandService);
  void commandService.executeCommand<ISwapRangesCommandParams>(SWAP_COMMAND_ID, {
    unitId: sheet.getWorkbook().getUnitId(),
    subUnitId: sheet.getSheetId(),
    matrixForA: emptyMatrix,
    matrixForB: buildValueMatrixFromRect(sheet, rectA),
    rangeA: rectA,
    rangeB: rectB,
  });

  return { ok: true };
}

/** 撤销/重做统一走 Facade（`univerAPI.undo()` / `redo()`），不直接操作 IUndoRedoService */

/** ICellData[][] → Univer mutation 期望的稀疏矩阵（key 是**绝对**行列号） */
function toSparseMatrix(matrix: ICellData[][], startRow: number, startCol: number): IObjectMatrixPrimitiveType<ICellData> {
  const out: IObjectMatrixPrimitiveType<ICellData> = {};
  matrix.forEach((row, r) => {
    const targetRow = startRow + r;
    row.forEach((cell, c) => {
      const targetCol = startCol + c;
      if (!out[targetRow]) out[targetRow] = {};
      out[targetRow][targetCol] = cell;
    });
  });
  return out;
}

export { SWAP_COMMAND_ID };
