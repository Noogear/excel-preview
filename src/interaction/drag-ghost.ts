/**
 * 拖动"幽灵"预览（纯 TS + DOM，不依赖 React / Univer）。
 *
 * 设计要点：
 * - 位置**只**通过 `transform: translate3d(x, y, 0)` 表达，绝不写 left/top（避免每帧 layout）；
 * - `moveTo` 用 rAF 合帧：同一帧内多次调用只写一次 transform（写的是最后一次坐标）；
 * - `content` 直接**搬入**（move，不是 clone）容器——Univer 的表格是 canvas，
 *   clone 出来是空白图；搬入 canvas 后它保留已绘制像素，观感才是"跟着指针的表格碎片"。
 *   调用方在拖动结束后负责把 content 放回原处（本模块不持有它的原父节点语义）；
 * - `accepting` 只切 class，颜色/描边由调用方 CSS 决定，本模块不写死任何颜色。
 */

/** 幽灵默认相对指针的偏移：让指针落在幽灵中心偏左上 */
const DEFAULT_OFFSET = { x: -24, y: -24 } as const;

const BASE_CLASS = 'dsh-drag-ghost';

export interface DragGhostOptions {
  /** 幽灵元素挂载到哪个容器（默认 document.body） */
  container?: HTMLElement;
  /** 相对指针的偏移（默认让指针落在幽灵中心偏左上） */
  offset?: { x: number; y: number };
  /** 禁用动画（prefers-reduced-motion 时由调用方传 true） */
  reducedMotion?: boolean;
  className?: string;
}

export interface DragGhost {
  /** 显示幽灵；content 是要显示的内容（HTMLElement，会被克隆或直接搬入） */
  show(content: HTMLElement, at: { x: number; y: number }): void;
  /** 更新位置（内部用 rAF 合并，同帧多次调用只写一次 transform） */
  moveTo(x: number, y: number): void;
  /** 切换"落点可接受/不可接受"的视觉态 */
  setAccepting(accepting: boolean): void;
  hide(): void;
  readonly visible: boolean;
  dispose(): void;
}

/** rAF / cAF：非浏览器环境下退化为同步执行（幽灵本来就只在浏览器里有意义）。 */
type RafFn = (cb: (time: number) => void) => number;
type CafFn = (id: number) => void;

function resolveRaf(): RafFn | null {
  const g = globalThis as { requestAnimationFrame?: RafFn };
  return typeof g.requestAnimationFrame === 'function' ? g.requestAnimationFrame.bind(globalThis) : null;
}

function resolveCaf(): CafFn | null {
  const g = globalThis as { cancelAnimationFrame?: CafFn };
  return typeof g.cancelAnimationFrame === 'function' ? g.cancelAnimationFrame.bind(globalThis) : null;
}

let baseStyleInjected = false;

/** 基础样式只注入一次（多个幽灵实例共享）。 */
function ensureBaseStyle(doc: Document): void {
  if (baseStyleInjected) return;
  if (doc.getElementById(`${BASE_CLASS}-style`)) {
    baseStyleInjected = true;
    return;
  }
  const style = doc.createElement('style');
  style.id = `${BASE_CLASS}-style`;
  style.textContent = [
    `.${BASE_CLASS}{`,
    '  position:fixed;',
    '  top:0;',
    '  left:0;',
    '  z-index:9999;',
    '  pointer-events:none;',
    '  will-change:transform;',
    '  transform:translate3d(-99999px,-99999px,0);',
    '  contain:layout style;',
    '  overflow:hidden;',
    '}',
    `.${BASE_CLASS}--reduced-motion{`,
    '  transition:none;',
    '}',
  ].join('\n');
  doc.head.appendChild(style);
  baseStyleInjected = true;
}

/**
 * 把 content 的左上角对齐到 `at` 时，幽灵根节点需要的 translate 值。
 * 注意：content 自身的宽高由调用方控制（通常设成与源区域一致），
 * 这里只负责"矩形左上角对齐 + 偏移"。
 */
function placeAt(position: { x: number; y: number }, offset: { x: number; y: number }): string {
  return `translate3d(${position.x + offset.x}px, ${position.y + offset.y}px, 0)`;
}

export function createDragGhost(options: DragGhostOptions = {}): DragGhost {
  const doc = options.container?.ownerDocument ?? globalThis.document;
  const className = options.className ?? BASE_CLASS;
  const offset = options.offset ?? { x: DEFAULT_OFFSET.x, y: DEFAULT_OFFSET.y };
  const reducedMotion = options.reducedMotion === true;

  let root: HTMLDivElement | null = null;
  let content: HTMLElement | null = null;
  let visible = false;
  let disposed = false;
  let accepting = false;
  let position = { x: 0, y: 0 };
  let rafId: number | null = null;

  const raf = resolveRaf();
  const caf = resolveCaf();

  function cancelPendingFrame(): void {
    if (rafId !== null) {
      caf?.(rafId);
      rafId = null;
    }
  }

  /** 立即把当前坐标刷进 transform（hide/dispose 前调用，避免取消 rAF 后位置回跳）。 */
  function flushTransform(): void {
    cancelPendingFrame();
    if (root) root.style.transform = placeAt(position, offset);
  }

  /** 合帧写入：同帧多次 moveTo 只保留最后一次坐标、只写一次 transform。 */
  function scheduleTransform(): void {
    if (raf === null || caf === null) {
      if (root) root.style.transform = placeAt(position, offset);
      return;
    }
    if (rafId !== null) return;
    rafId = raf(() => {
      rafId = null;
      if (disposed || !root) return;
      root.style.transform = placeAt(position, offset);
    });
  }

  function applyAccepting(): void {
    if (!root) return;
    root.classList.toggle(`${className}--accepting`, accepting);
    root.classList.toggle(`${className}--rejecting`, !accepting);
  }

  function ensureRoot(): HTMLDivElement {
    if (root) return root;
    const host = options.container ?? doc.body;
    const el = doc.createElement('div');
    el.className = className;
    if (reducedMotion) el.classList.add(`${BASE_CLASS}--reduced-motion`);
    // 挂载前先摆到屏幕外，杜绝初始位置闪一下。
    el.style.transform = 'translate3d(-99999px,-99999px,0)';
    if (content) el.appendChild(content);
    host.appendChild(el);
    root = el;
    return el;
  }

  return {
    get visible(): boolean {
      return visible;
    },

    show(next: HTMLElement, at: { x: number; y: number }): void {
      if (disposed) return;
      ensureBaseStyle(doc);
      const el = ensureRoot();

      if (content && content !== next) {
        // 旧内容不再属于幽灵，摘掉但不销毁（调用方可能还要用）。
        content.remove();
      }
      content = next;
      if (next.parentElement !== el) el.appendChild(next);

      // 入场：从屏幕外落位到指针附近。reducedMotion 时不做过渡，直接到位。
      position = { x: at.x, y: at.y };
      cancelPendingFrame();
      if (reducedMotion || raf === null) {
        el.style.transition = 'none';
        el.style.transform = placeAt(position, offset);
      } else {
        el.style.transition = 'none';
        el.style.transform = 'translate3d(-99999px,-99999px,0)';
        // 强制取一次布局，确保上面这一帧真的被应用，过渡才会从屏幕外开始。
        void el.offsetWidth;
        el.style.transition = 'transform 140ms cubic-bezier(0.2, 0.8, 0.2, 1)';
        el.style.transform = placeAt(position, offset);
      }

      accepting = false;
      applyAccepting();
      visible = true;
    },

    moveTo(x: number, y: number): void {
      if (disposed || !visible) return;
      position = { x, y };
      scheduleTransform();
    },

    setAccepting(next: boolean): void {
      if (disposed) return;
      accepting = next;
      applyAccepting();
    },

    hide(): void {
      if (disposed) return;
      // 先结算挂起的 rAF，再取消：否则最后一次 moveTo 会丢，隐藏瞬间位置回跳。
      flushTransform();
      visible = false;
      accepting = false;
      if (root) {
        root.style.transition = 'none';
        root.style.transform = 'translate3d(-99999px,-99999px,0)';
        root.classList.remove(`${className}--accepting`, `${className}--rejecting`);
      }
    },

    dispose(): void {
      if (disposed) return;
      disposed = true;
      visible = false;
      // 取消挂起的 rAF —— 这是唯一的异步句柄，取消后不会再有回调持有闭包。
      flushTransform();
      cancelPendingFrame();
      if (root) {
        root.remove();
        root = null;
      }
      content = null;
    },
  };
}

/** 是否偏好减弱动画。非浏览器环境（无 window / matchMedia）返回 false，且不抛异常。 */
export function prefersReducedMotion(): boolean {
  if (typeof window === 'undefined') return false;
  if (typeof window.matchMedia !== 'function') return false;
  try {
    return window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  } catch {
    return false;
  }
}
