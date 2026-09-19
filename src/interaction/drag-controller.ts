/**
 * 拖拽会话控制器：把"长按判定 / 幽灵绘制 / 落点命中 / 落点决策"串起来。
 *
 * 依赖的三条已实测事实：
 *  1. 按住鼠标拖动时 Univer 会持续派发 `CellPointerMove`，并给出真实 `{row, column}` → 落点命中靠它；
 *  2. 拖动过程中 Univer 原生的 `SelectionMove*` 不会触发 → 不存在手势冲突；
 *  3. 外部 DOM 拖放（HTML5 DnD）不可靠 → 侧边栏拖回也用 pointer 事件走同一套命中逻辑。
 */
import type { FUniver } from '@univerjs/core/facade';
import type { FWorksheet } from '@univerjs/sheets/facade';

import type { DragPayload, DropTarget, RangeSnapshot } from '../workspace/types';
import { createDragGhost, type DragGhost } from './drag-ghost';
import { renderSnapshotToElement } from './ghost-content';

export type DropHint = 'swap' | 'paste' | 'reject';

export interface DragControllerCallbacks {
  onTargetChange?: (target: DropTarget | null, hint: DropHint) => void;
  onDrop: (payload: DragPayload, target: DropTarget | null, pointer: { x: number; y: number }) => void;
  onDragStateChange?: (dragging: boolean) => void;
  /**
   * 自己算"指针下面是哪个单元格"。
   *
   * 为什么需要它：非选择模式下我们会**屏蔽 pointermove**（否则 Univer 会把按住拖动当成扩选，
   * 出现用户明确不要的多选框），而 Univer 的 `CellPointerMove` 事件正是靠那些 pointermove 派发的。
   * 所以落点命中改由这个回调负责——App 侧用 `ISheetSelectionRenderService.getCellWithCoordByOffset` 实现。
   */
  hitTest?: (x: number, y: number) => DropTarget | null;
}

export interface DragController {
  beginFromSheet(snapshot: RangeSnapshot, origin: { x: number; y: number }): void;
  beginFromWorkspace(snapshot: RangeSnapshot, origin: { x: number; y: number }): void;
  /** 由 App 喂入全局 pointermove */
  updatePointer(x: number, y: number): void;
  finish(x: number, y: number): void;
  cancel(): void;
  readonly active: boolean;
  readonly payload: DragPayload | null;
  dispose(): void;
}

interface Options {
  container: HTMLElement;
}

export function createDragController(
  univerAPI: FUniver,
  options: Options,
  callbacks: DragControllerCallbacks,
): DragController {
  const ghost: DragGhost = createDragGhost({ reducedMotion: prefersReducedMotionSafe() });
  const disposables: Array<{ dispose: () => void }> = [];

  let payload: DragPayload | null = null;
  let lastTarget: DropTarget | null = null;
  let lastHint: DropHint = 'paste';

  const getSheet = (): FWorksheet | null => univerAPI.getActiveWorkbook()?.getActiveSheet() ?? null;

  // ---- 落点命中：Univer 在按住拖动时持续派发 CellPointerMove ----
  const moveKey = (univerAPI.Event as unknown as Record<string, string>).CellPointerMove;
  if (moveKey) {
    disposables.push(
      univerAPI.addEvent(moveKey as never, (params: unknown) => {
        if (!payload) return;
        const p = (params ?? {}) as { row?: number; column?: number; worksheet?: FWorksheet };
        if (typeof p.row !== 'number' || typeof p.column !== 'number') return;
        const sheet = p.worksheet ?? getSheet();
        if (!sheet) return;
        applyTarget({ row: p.row, col: p.column, sheetId: sheet.getSheetId() });
      }) as { dispose: () => void },
    );
  }

  function applyTarget(target: DropTarget | null): void {
    lastTarget = target;
    if (!target) {
      ghost.setAccepting(false);
      callbacks.onTargetChange?.(null, 'reject');
      return;
    }

    const hint = computeHint(target);
    lastHint = hint;
    ghost.setAccepting(hint !== 'reject');
    callbacks.onTargetChange?.(target, hint);

    // 这里**刻意不再**调 setActiveSelection 去画目标框：
    // 选区在拖动期间由上层"钉"在源区域上（Univer 原生会把按住拖动当成扩选，
    // 谁后写谁生效，两边抢就会闪）。目标位置的高亮改由拖拽提示元素（.drag-hint）表达。
  }

  function payloadSize(): { rows: number; cols: number } | null {
    if (!payload) return null;
    return { rows: payload.snapshot.rows, cols: payload.snapshot.cols };
  }

  function computeHint(target: DropTarget): DropHint {
    if (!payload) return 'reject';
    if (payload.kind === 'workspace-item') return 'paste';

    const src = payload.snapshot.source;
    const size = payloadSize();
    if (!size) return 'reject';
    // 同尺寸且不重叠 → 互换；否则只能"粘贴覆盖"
    const sameSize = size.rows === src.endRow - src.startRow + 1 && size.cols === src.endCol - src.startCol + 1;
    const overlap =
      target.row <= src.endRow &&
      target.row + size.rows - 1 >= src.startRow &&
      target.col <= src.endCol &&
      target.col + size.cols - 1 >= src.startCol;

    if (overlap) return 'reject';
    return sameSize ? 'swap' : 'paste';
  }

  function begin(next: DragPayload, origin: { x: number; y: number }): void {
    payload = next;
    lastTarget = null;
    const content = renderSnapshotToElement(next.snapshot, { scale: 1 });
    ghost.show(content, origin);
    ghost.setAccepting(false);
    callbacks.onDragStateChange?.(true);
  }

  return {
    beginFromSheet: (snapshot, origin) => begin({ kind: 'sheet-range', snapshot }, origin),
    beginFromWorkspace: (snapshot, origin) => begin({ kind: 'workspace-item', snapshot }, origin),

    updatePointer: (x, y) => {
      if (!payload) return;
      ghost.moveTo(x, y);

      // 指针离开表格区域（例如移到侧边栏）→ 清掉落点
      const rect = options.container.getBoundingClientRect();
      const inside = x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
      if (!inside) {
        if (lastTarget) applyTarget(null);
        return;
      }

      // 有自建命中就用它（非选择模式下 pointermove 被屏蔽，CellPointerMove 不会派发）；
      // 没有就退回 Univer 事件那条老路（选择模式下仍然有效）。
      const hit = callbacks.hitTest?.(x, y) ?? null;
      if (hit) {
        applyTarget(hit);
      } else if (callbacks.hitTest && lastTarget) {
        applyTarget(null);
      }
    },

    finish: (x, y) => {
      if (!payload) return;
      const droppedPayload = payload;
      const target = lastTarget;
      payload = null;
      lastTarget = null;
      ghost.hide();
      callbacks.onDragStateChange?.(false);
      callbacks.onDrop(droppedPayload, target, { x, y });
    },

    cancel: () => {
      payload = null;
      lastTarget = null;
      ghost.hide();
      callbacks.onDragStateChange?.(false);
    },

    get active() {
      return payload !== null;
    },

    get payload() {
      return payload;
    },

    dispose: () => {
      disposables.forEach((d) => d.dispose());
      ghost.dispose();
    },
  };
}

function prefersReducedMotionSafe(): boolean {
  try {
    return typeof window !== 'undefined' && typeof window.matchMedia === 'function'
      ? window.matchMedia('(prefers-reduced-motion: reduce)').matches
      : false;
  } catch {
    return false;
  }
}
