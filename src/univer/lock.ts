/**
 * P0 探测 ④：只允许编辑内容、不允许编辑格式。
 *
 * 事实核查结论（@univerjs 0.25.1）：
 *  - `ICommandService` **没有** `interceptCommand`（旧版 API，已移除）。
 *    可用的是 `beforeCommandExecuted(listener)`（返回 void，**无法取消执行**）与 `onCommandExecuted`。
 *  - 权限点（`IPermissionService` + `WorksheetSetCellStylePermission` 等）由
 *    `SheetPermissionCheckController` 在**命令执行之后**检查并触发 UI 拦截/提示，
 *    属于"事后拦截"，不能作为唯一防线。
 *  - 因此可靠做法有两层：
 *      A. 样式类命令 unregister + 用 no-op 命令替换（UI 点了也没用）
 *      B. 在 **mutation 层** 包装 `SetRangeValuesMutation`，剥离单元格数据里的 `s`（样式）
 *         —— 无论走打字、粘贴、填充、拖拽哪条路径，样式都写不进去
 */
import { CommandType, ICommandService, type IDisposable, type IAccessor, type IMutation } from '@univerjs/core';
import {
  AddWorksheetMergeAllCommand,
  AddWorksheetMergeCommand,
  AddWorksheetMergeHorizontalCommand,
  AddWorksheetMergeVerticalCommand,
  ClearSelectionFormatCommand,
  ResetBackgroundColorCommand,
  ResetTextColorCommand,
  SetBackgroundColorCommand,
  SetBoldCommand,
  SetBorderBasicCommand,
  SetBorderColorCommand,
  SetBorderCommand,
  SetBorderPositionCommand,
  SetBorderStyleCommand,
  SetColWidthCommand,
  SetFontFamilyCommand,
  SetFontSizeCommand,
  SetHorizontalTextAlignCommand,
  SetItalicCommand,
  SetOverlineCommand,
  SetRangeValuesCommand,
  SetRangeValuesMutation,
  SetRowHeightCommand,
  SetStrikeThroughCommand,
  SetStyleCommand,
  SetTextColorCommand,
  SetTextRotationCommand,
  SetTextWrapCommand,
  SetUnderlineCommand,
  SetVerticalTextAlignCommand,
  SetWorksheetDefaultStyleCommand,
  SetWorksheetRangeThemeStyleCommand,
  type ISetRangeValuesMutationParams,
} from '@univerjs/sheets';
import type { FWorksheet } from '@univerjs/sheets/facade';

import { log } from '../p0/log';
import { recordDirtyCell as recordGlobalDirtyCell } from './dirty-tracker';

/** 全部"格式类"命令：这些必须被拦掉 */
const FORMAT_COMMANDS = [
  SetStyleCommand,
  SetBoldCommand,
  SetItalicCommand,
  SetUnderlineCommand,
  SetOverlineCommand,
  SetStrikeThroughCommand,
  SetFontFamilyCommand,
  SetFontSizeCommand,
  SetTextColorCommand,
  SetBackgroundColorCommand,
  ResetTextColorCommand,
  ResetBackgroundColorCommand,
  SetHorizontalTextAlignCommand,
  SetVerticalTextAlignCommand,
  SetTextWrapCommand,
  SetTextRotationCommand,
  SetBorderCommand,
  SetBorderBasicCommand,
  SetBorderColorCommand,
  SetBorderPositionCommand,
  SetBorderStyleCommand,
  SetColWidthCommand,
  SetRowHeightCommand,
  AddWorksheetMergeCommand,
  AddWorksheetMergeAllCommand,
  AddWorksheetMergeHorizontalCommand,
  AddWorksheetMergeVerticalCommand,
  ClearSelectionFormatCommand,
  SetWorksheetDefaultStyleCommand,
  SetWorksheetRangeThemeStyleCommand,
];

export interface ContentOnlyLock {
  formatCommandIds: string[];
  blockedCount: () => number;
  allowedCount: () => number;
  /** 被 mutation 层剥离掉样式的单元格次数 */
  strippedCount: () => number;
  /**
   * 被写入过的单元格（sheetId → `row:col` 集合）。
   * 外科式导出据此**只回写这些格子**，其余部件字节原样保留。
   */
  getDirtyCells: () => Map<string, Set<string>>;
  restore: () => void;
}

export function installContentOnlyLock(fWorksheet: FWorksheet): ContentOnlyLock {
  const commandService = fWorksheet.getInject().get(ICommandService);
  const disposables: IDisposable[] = [];
  let blocked = 0;
  let allowed = 0;
  let stripped = 0;
  /** sheetId → 被写过的单元格坐标集合（导出用） */
  const dirtyCells = new Map<string, Set<string>>();

  function recordDirty(params?: ISetRangeValuesMutationParams): void {
    if (!params?.cellValue || !params.subUnitId) return;
    let bucket = dirtyCells.get(params.subUnitId);
    if (!bucket) {
      bucket = new Set<string>();
      dirtyCells.set(params.subUnitId, bucket);
    }
    for (const [rowKey, row] of Object.entries(params.cellValue as Record<number, Record<number, unknown>>)) {
      for (const colKey of Object.keys(row ?? {})) {
        bucket.add(`${rowKey}:${colKey}`);
        // 同时写入**全局**追踪（按 workbookId 归档）：多标签/切表后仍然知道哪些格子被改过，
        // 外科式导出与"误关闭恢复"都依赖它
        if (params.unitId) {
          recordGlobalDirtyCell(params.unitId, params.subUnitId, Number(rowKey), Number(colKey));
        }
      }
    }
  }

  // ---------- A. 样式命令：注销后用 no-op 占位 ----------
  const formatCommandIds: string[] = [];
  for (const command of FORMAT_COMMANDS) {
    const id = (command as { id?: string }).id;
    if (!id) continue;
    formatCommandIds.push(id);
    try {
      commandService.unregisterCommand(id);
    } catch (error) {
      log('lock:unregister-error', { id, message: String(error) });
    }
    disposables.push(
      commandService.registerCommand({
        id,
        type: CommandType.COMMAND,
        handler: () => {
          blocked += 1;
          log('lock:blocked', { id, via: 'noop-command-replacement' });
          return false;
        },
      }),
    );
  }

  // ---------- B. 值命令：放行，但记录参数（用于确认样式字段是否混入） ----------
  const valueCommandId = (SetRangeValuesCommand as { id: string }).id;
  disposables.push(
    commandService.beforeCommandExecuted((commandInfo) => {
      if (commandInfo.id !== valueCommandId) return;
      allowed += 1;
      log('lock:value-command', { id: commandInfo.id, params: summarizeParams(commandInfo.params) });
    }),
  );

  // ---------- C. mutation 层兜底：剥离样式字段 ----------
  const mutationId = (SetRangeValuesMutation as { id: string }).id;
  const originalHandler = (SetRangeValuesMutation as IMutation<ISetRangeValuesMutationParams, boolean>).handler;
  try {
    commandService.unregisterCommand(mutationId);
    disposables.push(
      commandService.registerCommand({
        id: mutationId,
        type: CommandType.MUTATION,
        handler: (accessor: IAccessor, params?: ISetRangeValuesMutationParams) => {
          recordDirty(params);
          const { cleaned, removed } = stripStyles(params);
          if (removed > 0) {
            stripped += removed;
            log('lock:stripped-style-from-mutation', { removed, id: mutationId });
          }
          return originalHandler(accessor, cleaned);
        },
      } as IMutation<ISetRangeValuesMutationParams, boolean>),
    );
  } catch (error) {
    log('lock:mutation-wrap-error', { message: String(error) });
  }

  log('lock:installed', {
    formatCommandCount: formatCommandIds.length,
    strategy: ['noop-command-replacement', 'mutation-style-stripping'],
  });

  return {
    formatCommandIds,
    blockedCount: () => blocked,
    allowedCount: () => allowed,
    strippedCount: () => stripped,
    getDirtyCells: () => dirtyCells,
    restore: () => disposables.forEach((d) => d.dispose()),
  };
}

/**
 * 剥离单元格数据里的**样式**字段，只保留内容：
 *  - `s`：单元格样式（字体/加粗/边框/背景/字色/斜体/数字格式…）——**直接丢弃**
 *  - `p`：单元格内富文本 —— **必须保留**，不能拍平！
 *
 * 为什么不能动 `p`（实测教训）：Univer 把单元格**超链接**存成富文本里的 link mark，
 * 一旦把 `p` 拍平成纯文本，链接标记就被销毁 —— 表现为"应用超链接返回成功、批注图片都正常，
 * 但 `getHyperLinks()` 读回为空"。而"超链接/批注等对象只读且必须原样保留"是本产品的硬要求，
 * 因此这里只剥样式 `s`。
 *
 * 富文本格式化的防线改为放在**命令层**：编辑器内的加粗/斜体/字号/颜色等命令
 * 已被只读闸门默认拒绝（见 read-only-guard 的 DENY_PATTERNS），用户无法新增格式。
 */
function stripStyles(params?: ISetRangeValuesMutationParams): { cleaned: ISetRangeValuesMutationParams; removed: number } {
  if (!params?.cellValue) return { cleaned: params as ISetRangeValuesMutationParams, removed: 0 };
  let removed = 0;
  const cellValue: Record<number, Record<number, Record<string, unknown>>> = {};
  for (const [rowKey, row] of Object.entries(params.cellValue as Record<number, Record<number, Record<string, unknown>>>)) {
    const nextRow: Record<number, Record<string, unknown>> = {};
    for (const [colKey, cell] of Object.entries(row ?? {})) {
      if (cell && typeof cell === 'object' && 's' in cell) {
        removed += 1;
        const { s: _style, ...rest } = cell;
        nextRow[Number(colKey)] = rest;
      } else {
        nextRow[Number(colKey)] = cell;
      }
    }
    cellValue[Number(rowKey)] = nextRow;
  }
  return { cleaned: { ...params, cellValue } as ISetRangeValuesMutationParams, removed };
}

function summarizeParams(params: unknown): unknown {
  if (!params || typeof params !== 'object') return params;
  const p = params as Record<string, unknown>;
  const cellValue = p.cellValue as Record<number, Record<number, Record<string, unknown>>> | undefined;
  const styleFieldsSeen: string[] = [];
  let cellCount = 0;
  if (cellValue) {
    for (const row of Object.values(cellValue)) {
      for (const cell of Object.values(row ?? {})) {
        cellCount += 1;
        if (cell && typeof cell === 'object') {
          for (const key of Object.keys(cell)) if (!styleFieldsSeen.includes(key)) styleFieldsSeen.push(key);
        }
      }
    }
  }
  return { keys: Object.keys(p), cellCount, cellFieldNames: styleFieldsSeen };
}

/** P0 探测：权限服务与权限点的真实用法（作为 P1 的备选防线评估） */
export function probePermissionApi(fWorksheet: FWorksheet): void {
  try {
    const permission = fWorksheet.getWorksheetPermission() as unknown as Record<string, unknown>;
    const methods = Object.keys(permission).filter((k) => typeof permission[k] === 'function');
    log('permission:facade-methods', { methods });
  } catch (error) {
    log('permission:facade-error', { message: String(error) });
  }

  try {
    const injector = fWorksheet.getInject() as unknown as { get: (token: unknown) => unknown };
    // IPermissionService 需要从 core 动态取，避免顶层强耦合
    void injector;
    log('permission:route', {
      note: '权限点由 SheetPermissionCheckController 在命令执行后检查，属于事后拦截；已作为 P1 备选防线记录',
    });
  } catch (error) {
    log('permission:probe-error', { message: String(error) });
  }
}
