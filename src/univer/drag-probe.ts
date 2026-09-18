/**
 * P0 探测 ③：长按拖动与 Univer 原生手势的冲突情况。
 *
 * 要回答的问题：
 *  Q1 按住鼠标移动时，Facade 的 CellPointerMove 是否持续触发？（决定命中测试走事件还是自建几何）
 *  Q2 长按（400ms 不动）是否会触发 Univer 自身的编辑/选区行为？（决定长按手势是否可用）
 *  Q3 Univer 是否自带"拖选区边框移动内容"手势？(SelectionMoveStart/Moving/End 何时触发)
 *  Q4 外部 DOM 元素（侧边栏）拖到表格上时，Event.DragOver / Event.Drop 是否给出 row/column？
 */
import type { FUniver } from '@univerjs/core/facade';

import { log } from '../p0/log';

interface Gesture {
  startX: number;
  startY: number;
  downAt: number;
  moves: number;
  cellPointerMoves: number;
  cellHoverMoves: number;
  maxDistance: number;
  longPressFired: boolean;
  selectionEvents: string[];
  lastCell?: { row: number; column: number };
  pointerType: string;
}

let gesture: Gesture | null = null;

const LONG_PRESS_MS = 400;
const MOVE_TOLERANCE_PX = 4;

export interface DragProbeHandle {
  dispose: () => void;
  /** 供测试读取的探测摘要 */
  summary: () => Record<string, unknown>;
}

export function installDragProbe(univerAPI: FUniver, container: HTMLElement): DragProbeHandle {
  const disposables: Array<{ dispose: () => void }> = [];
  const eventNames = [
    'CellPointerDown',
    'CellPointerUp',
    'CellPointerMove',
    'CellHover',
    'CellClicked',
    'SelectionChanged',
    'SelectionMoveStart',
    'SelectionMoving',
    'SelectionMoveEnd',
    'DragOver',
    'Drop',
    'Scroll',
    'BeforeSheetEditStart',
    'SheetEditStarted',
    'SheetEditEnded',
  ] as const;

  for (const name of eventNames) {
    const eventKey = (univerAPI.Event as unknown as Record<string, string>)[name];
    if (!eventKey) {
      log('probe:missing-event', { name });
      continue;
    }
    try {
      const disposable = univerAPI.addEvent(eventKey as never, (params: unknown) => {
        const p = (params ?? {}) as Record<string, unknown>;
        const row = typeof p.row === 'number' ? p.row : undefined;
        const column = typeof p.column === 'number' ? p.column : undefined;

        if (name === 'CellPointerMove' && gesture) {
          gesture.cellPointerMoves += 1;
          if (row !== undefined && column !== undefined) gesture.lastCell = { row, column };
        }
        if (name === 'CellHover' && gesture) gesture.cellHoverMoves += 1;
        if (name.startsWith('Selection')) {
          gesture?.selectionEvents.push(name);
        }
        if (name === 'Drop' || name === 'DragOver') {
          log(`probe:${name}`, {
            row,
            column,
            hasDataTransfer: !!p.dataTransfer,
            types: safeDataTransferTypes(p.dataTransfer),
          });
          return;
        }
        log(`event:${name}`, row === undefined ? undefined : { row, column });
      });
      disposables.push(disposable as { dispose: () => void });
    } catch (error) {
      log('probe:add-event-error', { name, message: String(error) });
    }
  }

  // ---- 手势探测（页面级，捕获阶段，先于 Univer 处理）----
  const onPointerDown = (e: PointerEvent) => {
    if (e.button !== 0) return;
    const target = e.target as HTMLElement | null;
    const inGrid = !!target?.closest('.univer-render-canvas, .univer-canvas, canvas');
    gesture = {
      startX: e.clientX,
      startY: e.clientY,
      downAt: performance.now(),
      moves: 0,
      cellPointerMoves: 0,
      cellHoverMoves: 0,
      maxDistance: 0,
      longPressFired: false,
      selectionEvents: [],
      pointerType: e.pointerType,
    };
    log('gesture:down', { inGrid, x: e.clientX, y: e.clientY, pointerType: e.pointerType });

    const timer = window.setTimeout(() => {
      if (!gesture) return;
      if (gesture.maxDistance <= MOVE_TOLERANCE_PX && gesture.moves === 0) {
        gesture.longPressFired = true;
        log('gesture:longpress', { at: Math.round(performance.now() - gesture.downAt) });
      }
    }, LONG_PRESS_MS);

    const onPointerMove = (ev: PointerEvent) => {
      if (!gesture) return;
      gesture.moves += 1;
      const dx = ev.clientX - gesture.startX;
      const dy = ev.clientY - gesture.startY;
      gesture.maxDistance = Math.max(gesture.maxDistance, Math.hypot(dx, dy));
    };

    const onPointerUp = (ev: PointerEvent) => {
      window.clearTimeout(timer);
      window.removeEventListener('pointermove', onPointerMove, true);
      window.removeEventListener('pointerup', onPointerUp, true);
      window.removeEventListener('pointercancel', onPointerUp, true);
      if (!gesture) return;
      const summary = {
        durationMs: Math.round(performance.now() - gesture.downAt),
        rawPointerMoves: gesture.moves,
        // Q1 的答案在这里：
        cellPointerMovesWhilePressed: gesture.cellPointerMoves,
        cellHoverWhilePressed: gesture.cellHoverMoves,
        maxDistancePx: Math.round(gesture.maxDistance),
        longPressFired: gesture.longPressFired,
        selectionEvents: gesture.selectionEvents,
        lastCell: gesture.lastCell,
        endX: ev.clientX,
        endY: ev.clientY,
        pointerType: gesture.pointerType,
      };
      log('gesture:up', summary);
      lastGestureSummary = summary;
      gesture = null;
    };

    window.addEventListener('pointermove', onPointerMove, true);
    window.addEventListener('pointerup', onPointerUp, true);
    window.addEventListener('pointercancel', onPointerUp, true);
  };

  container.addEventListener('pointerdown', onPointerDown, true);

  // ---- Q4：外部 DOM 拖拽（HTML5 DnD）到表格 ----
  const onDragOver = (e: DragEvent) => {
    // 不 preventDefault，观察 Univer 自己是否处理
    log('dom:dragover', { types: e.dataTransfer ? Array.from(e.dataTransfer.types) : [] });
  };
  const onDrop = (e: DragEvent) => {
    log('dom:drop', { types: e.dataTransfer ? Array.from(e.dataTransfer.types) : [] });
  };
  container.addEventListener('dragover', onDragOver);
  container.addEventListener('drop', onDrop);

  return {
    dispose: () => {
      disposables.forEach((d) => d.dispose());
      container.removeEventListener('pointerdown', onPointerDown, true);
      container.removeEventListener('dragover', onDragOver);
      container.removeEventListener('drop', onDrop);
    },
    summary: () => lastGestureSummary ?? {},
  };
}

let lastGestureSummary: Record<string, unknown> | null = null;

function safeDataTransferTypes(dataTransfer: unknown): string[] {
  try {
    const dt = dataTransfer as DataTransfer | undefined;
    return dt ? Array.from(dt.types) : [];
  } catch {
    return [];
  }
}
