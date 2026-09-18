/**
 * 互换后的"黄色提醒框"（用户要求：**逐渐透明然后消失**，且**不影响操作**）。
 *
 * 为什么不用 Univer 的标记图层 / 选区样式，而自己画 DOM：
 *  1. 选区样式的自定义 `style` 渲染层不采纳（只有主题默认蓝框，看着像"还选中着"）；
 *  2. 标记图层（`highlightRanges`）能出黄色，但**没法做渐隐动画**，而且每次刷新都要重绘画布；
 *  3. DOM 浮层用 CSS 动画做透明度，走合成器、不重绘画布、`pointer-events: none` 完全穿透，
 *     因此"不影响操作"是结构性保证，而不是靠时序躲开。
 *
 * 坐标换算（**实测校正过，别再照抄上游那个反解**）：
 *   `skeleton.getCellWithCoordByIndex(row, col)` 给的是**内容坐标**（含行/列表头偏移，单位是未缩放的
 *   内容像素）；画布内 CSS 像素 = `(内容坐标 − 视口滚动量) × 缩放`。
 *
 *   踩坑记录：先前按上游 `getTransformOffsetX` 的反解写成
 *   `css = (content − scroll) × scale + scroll`——它在 scale=1 时**把滚动量消掉了**，
 *   于是"滚动之后提醒框画在没滚动时的位置"（用户实测反馈），而且会算出画布外的坐标被裁掉。
 *   用"点某个像素 → 读回 Univer 认定的格子 → 该格矩形是否包住这个像素"的对照实测校正成上式。
 */

export interface FlashCellRect {
  /** 单元格内容坐标（含行/列表头偏移），直接来自 skeleton.getCellWithCoordByIndex */
  startX: number;
  startY: number;
  endX: number;
  endY: number;
}

export interface FlashViewport {
  /** 单元格在"画布坐标系"里的矩形（含表头偏移） */
  cellRect: (row: number, col: number) => FlashCellRect | null;
  /** 主视口滚动量（画布像素） */
  scroll: { x: number; y: number };
  /** 渲染缩放（浏览器/表格缩放都会体现在这里） */
  scale: { x: number; y: number };
}

export interface FlashBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** 单元格 → 画布内矩形（纯函数，便于单测） */
export function computeFlashRect(
  viewport: FlashViewport,
  row: number,
  col: number,
): FlashBox | null {
  const rect = viewport.cellRect(row, col);
  if (!rect) return null;
  const { x: scrollX, y: scrollY } = viewport.scroll;
  const { x: scaleX, y: scaleY } = viewport.scale;
  if (![scrollX, scrollY, scaleX, scaleY].every((value) => Number.isFinite(value)) || scaleX === 0 || scaleY === 0) {
    return null;
  }
  // 内容坐标先减滚动量，再乘缩放（见文件头的"实测校正"说明）
  const left = (rect.startX - scrollX) * scaleX;
  const top = (rect.startY - scrollY) * scaleY;
  const width = (rect.endX - rect.startX) * scaleX;
  const height = (rect.endY - rect.startY) * scaleY;
  if (![left, top, width, height].every((value) => Number.isFinite(value))) return null;
  if (width <= 0 || height <= 0) return null;
  return { left, top, width, height };
}

export interface SwapFlashOverlayOptions {
  /** 一个目标（单元格）在**视口**里的矩形；每帧重新调用，滚动/缩放时提醒框会跟着走 */
  measure: (target: FlashTarget) => FlashBox | null;
  /** 淡出时长（毫秒），默认 1600；`persistent` 为 true 时忽略 */
  durationMs?: number;
  /** true = 常亮不淡出（拖动过程中的"落点高亮"用），默认 false */
  persistent?: boolean;
  /** 浮层与方框的 class（默认是互换提醒框的） */
  layerClass?: string;
  boxClass?: string;
  /** 注入点（单测/SSR 友好） */
  doc?: Document | null;
  raf?: ((callback: () => void) => number) | null;
  cancelRaf?: ((id: number) => void) | null;
  now?: () => number;
}

export interface FlashTarget {
  /** 用于诊断日志 */
  a1: string;
  row: number;
  col: number;
}

export interface SwapFlashOverlay {
  /** 显示提醒框（会替换掉上一批）；目标为空时等价于 hide()。`variant: 'reject'` 表示落点无效 */
  show: (targets: FlashTarget[], options?: { variant?: 'normal' | 'reject' }) => void;
  hide: () => void;
  /** 当前是否还有提醒框在屏幕上（e2e/单测断言用） */
  readonly active: boolean;
  dispose: () => void;
}

export const FLASH_LAYER_CLASS = 'swap-flash-layer';
export const FLASH_BOX_CLASS = 'swap-flash-box';
export const DEFAULT_FLASH_DURATION_MS = 1600;
/** 拖动过程中的"落点高亮"（常亮、不淡出） */
export const TARGET_LAYER_CLASS = 'drag-target-layer';
export const TARGET_BOX_CLASS = 'drag-target-box';

/**
 * 创建提醒框浮层。
 *
 * 生命周期由自己管（不占 React 状态）：`show` 时建 DOM + 起 rAF 跟随，到期自动移除。
 * 期间只改自己那几个 div 的 `transform/width/height`，不触碰 Univer 的任何状态。
 */
export function createSwapFlashOverlay(options: SwapFlashOverlayOptions): SwapFlashOverlay {
  const doc = options.doc === undefined ? (typeof document === 'undefined' ? null : document) : options.doc;
  const raf = options.raf === undefined ? (typeof requestAnimationFrame === 'function' ? requestAnimationFrame : null) : options.raf;
  const cancelRaf =
    options.cancelRaf === undefined ? (typeof cancelAnimationFrame === 'function' ? cancelAnimationFrame : null) : options.cancelRaf;
  const now = options.now ?? (() => Date.now());
  const durationMs = Math.max(200, options.durationMs ?? DEFAULT_FLASH_DURATION_MS);
  const persistent = options.persistent === true;
  const layerClass = options.layerClass ?? FLASH_LAYER_CLASS;
  const boxClass = options.boxClass ?? FLASH_BOX_CLASS;

  let layer: HTMLElement | null = null;
  let boxes: Array<{ el: HTMLElement; target: FlashTarget; bornAt: number }> = [];
  let frame: number | null = null;

  const stopLoop = (): void => {
    if (frame === null) return;
    cancelRaf?.(frame);
    frame = null;
  };

  const clearBoxes = (): void => {
    boxes.forEach(({ el }) => el.remove());
    boxes = [];
  };

  const paint = (): void => {
    const expired = !persistent && boxes.length > 0 && now() - Math.max(...boxes.map((b) => b.bornAt)) >= durationMs;
    if (expired) {
      hide();
      return;
    }
    boxes.forEach(({ el, target }) => {
      const box = options.measure(target);
      if (!box) {
        el.style.display = 'none';
        return;
      }
      el.style.display = 'block';
      el.style.transform = `translate3d(${Math.round(box.left)}px, ${Math.round(box.top)}px, 0)`;
      el.style.width = `${Math.round(box.width)}px`;
      el.style.height = `${Math.round(box.height)}px`;
    });
    if (boxes.length > 0) frame = raf?.(paint) ?? null;
  };

  const hide = (): void => {
    stopLoop();
    clearBoxes();
    if (layer) layer.style.display = 'none';
  };

  const show = (targets: FlashTarget[], options2?: { variant?: 'normal' | 'reject' }): void => {
    hide();
    if (doc === null || targets.length === 0) return;
    if (layer === null || !layer.isConnected) {
      layer = doc.createElement('div');
      layer.className = layerClass;
      doc.body.appendChild(layer);
    }
    layer.style.display = 'block';
    const reject = options2?.variant === 'reject';
    const bornAt = now();
    boxes = targets.map((target) => {
      const el = doc.createElement('div');
      el.className = reject ? `${boxClass} is-reject` : boxClass;
      el.setAttribute('data-a1', target.a1);
      if (!persistent) el.style.animationDuration = `${durationMs}ms`;
      layer!.appendChild(el);
      return { el, target, bornAt };
    });
    paint();
  };

  return {
    show,
    hide,
    get active() {
      return boxes.length > 0;
    },
    dispose: () => {
      hide();
      layer?.remove();
      layer = null;
    },
  };
}

/**
 * 拖动过程中的"落点高亮"：常亮不淡出，松手/取消时由调用方 `hide()`。
 * 与互换提醒框共用同一套测量与定位逻辑，只是不自动消失、class 不同。
 */
export function createDragTargetHighlight(options: Omit<SwapFlashOverlayOptions, 'persistent' | 'durationMs'>): SwapFlashOverlay {
  return createSwapFlashOverlay({
    ...options,
    persistent: true,
    layerClass: options.layerClass ?? TARGET_LAYER_CLASS,
    boxClass: options.boxClass ?? TARGET_BOX_CLASS,
  });
}
