/**
 * 「点击交换」的动画原语（纯 TS + DOM，不依赖 React / Univer）。
 * 关键取舍：Univer 是 canvas 渲染，没有"单元格 → 像素"映射，只能拿指针坐标做内容芯片飞行；位移只写 transform（复用 flyIn）绝不逐帧写 left/top；临时节点 / rAF / timer 全部登记，动画结束与 dispose 统一清理，零泄漏。
 * 兜底：swap-animation.css 未加载时（如 e2e 宿主没 import）注入一份同源样式，避免动画退化成无样式裸 div。
 */

import { flyIn } from './fly-in';

export interface SwapAnimationOptions {
  /** 临时节点挂到哪个容器（默认 document.body；层本身 fixed 满屏，不受父级 overflow/stacking 影响） */
  container?: HTMLElement;
  reducedMotion?: boolean;
}

export interface SwapAnimation {
  /** 选中某一方时的反馈：在点击点生成一个脉动圆环（约 420ms 后自动清理） */
  pulseSelection(at: { x: number; y: number }): void;
  /** 交换动画：两个内容芯片 a→b、b→a 对飞；返回 Promise，动画结束（含降级路径）后 resolve。 */
  playSwap(
    a: { x: number; y: number; text: string },
    b: { x: number; y: number; text: string },
    options?: { durationMs?: number },
  ): Promise<void>;
  /** 待交换期间跟随指针的幽灵（显示已选内容摘要） */
  showPendingGhost(text: string, at: { x: number; y: number }): void;
  movePendingGhost(at: { x: number; y: number }): void;
  hidePendingGhost(): void;
  dispose(): void;
}

/* 常量 */

/** 类名前缀（与 swap-animation.css 严格一致） */
export const SWAP_CLASS_PREFIX = 'swap-';

export const SWAP_TEXT_MAX_CHARS = 24;

export const DEFAULT_SWAP_DURATION_MS = 340;
export const DEFAULT_PULSE_DURATION_MS = 420;

const CHIP_FADE_IN_MS = 120;
const CHIP_FADE_OUT_MS = 140;
const GHOST_FADE_IN_MS = 120;
const GHOST_FADE_OUT_MS = 140;

/** 幽灵相对指针的偏移（内容浮在指针右上方，不遮住点击点） */
const GHOST_OFFSET = { x: 16, y: -42 } as const;

const ROOT_CLASS = `${SWAP_CLASS_PREFIX}layer`;
const CHIPS_CLASS = `${SWAP_CLASS_PREFIX}chips`;
const CHIP_CLASS = `${SWAP_CLASS_PREFIX}chip`;
const PULSE_CLASS = `${SWAP_CLASS_PREFIX}pulse`;
const RING_CLASS = `${SWAP_CLASS_PREFIX}pulse-ring`;
const GHOST_CLASS = `${SWAP_CLASS_PREFIX}pending-ghost`;
const GHOST_TEXT_CLASS = `${SWAP_CLASS_PREFIX}pending-ghost-text`;
const BASE_STYLE_ID = `${SWAP_CLASS_PREFIX}layer-style`;
const SVG_NS = 'http://www.w3.org/2000/svg';

const CHIP_FADE_IN_CLASS = `${CHIP_CLASS}--fade-in`;
const CHIP_FADE_OUT_CLASS = `${CHIP_CLASS}--fade-out`;
const GHOST_IN_CLASS = `${GHOST_CLASS}--ghost-in`;
const GHOST_OUT_CLASS = `${GHOST_CLASS}--ghost-out`;

/** 兜底样式，与 swap-animation.css 的变量/取值一致（CSS 已加载时两者等价、互不冲突）。
 *  `transition:none` 必须保留：flyIn 逐帧写 transform，若挂了 transform 过渡，浏览器会把每帧新值当成新的过渡目标（retargeting），元素被拖住永远追不上插值；过渡只用于 opacity。 */
const BASE_CSS = [
  `.${ROOT_CLASS}{position:fixed;inset:0;pointer-events:none;z-index:var(--swap-layer-z-index,9998);}`,
  `.${CHIPS_CLASS}{position:absolute;inset:0;pointer-events:none;}`,
  `.${CHIP_CLASS}{position:absolute;left:0;top:0;box-sizing:border-box;pointer-events:none;transition:none;`,
  'max-width:var(--swap-chip-max-width,240px);',
  'padding:var(--swap-chip-padding-y,5px) var(--swap-chip-padding-x,10px);',
  'background:var(--swap-chip-bg,#fff);color:var(--swap-chip-color,#1f2937);',
  'border:var(--swap-chip-border-width,1px) solid var(--swap-chip-border,#2563eb);',
  'border-radius:var(--swap-chip-radius,999px);',
  'box-shadow:var(--swap-chip-shadow,0 6px 18px rgba(15,23,42,.18));',
  'font-family:var(--swap-chip-font-family,"Microsoft YaHei",system-ui,-apple-system,sans-serif);',
  'font-size:var(--swap-chip-font-size,12px);line-height:var(--swap-chip-line-height,18px);',
  'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;',
  'transform-origin:top left;will-change:transform;contain:layout style;}',
  `.${CHIP_FADE_IN_CLASS}{animation:${SWAP_CLASS_PREFIX}chip-in var(--swap-chip-fade-in-duration,120ms) ease-out both;}`,
  `.${CHIP_FADE_OUT_CLASS}{animation:${SWAP_CLASS_PREFIX}chip-out var(--swap-chip-fade-out-duration,140ms) ease-in both;}`,
  `.${PULSE_CLASS}{position:absolute;left:0;top:0;pointer-events:none;overflow:visible;will-change:transform;}`,
  `.${RING_CLASS}{fill:none;vector-effect:non-scaling-stroke;}`,
  `.${PULSE_CLASS}--pulse{animation:${SWAP_CLASS_PREFIX}ring-pulse var(--swap-pulse-duration,420ms) cubic-bezier(.16,.84,.44,1) both;}`,
  `.${GHOST_CLASS}{position:absolute;left:0;top:0;pointer-events:none;will-change:transform;}`,
  `.${GHOST_IN_CLASS}{animation:${SWAP_CLASS_PREFIX}ghost-in var(--swap-ghost-fade-in-duration,120ms) ease-out both;}`,
  `.${GHOST_OUT_CLASS}{animation:${SWAP_CLASS_PREFIX}ghost-out var(--swap-ghost-fade-out-duration,140ms) ease-in both;}`,
  `.${GHOST_TEXT_CLASS}{display:block;box-sizing:border-box;max-width:var(--swap-chip-max-width,240px);`,
  'padding:var(--swap-chip-padding-y,5px) var(--swap-chip-padding-x,10px);',
  'background:var(--swap-ghost-bg,var(--swap-chip-bg,#fff));color:var(--swap-ghost-color,var(--swap-chip-color,#1f2937));',
  'border:var(--swap-chip-border-width,1px) dashed var(--swap-ghost-border,var(--swap-chip-border,#2563eb));',
  'border-radius:var(--swap-chip-radius,999px);',
  'box-shadow:var(--swap-chip-shadow,0 6px 18px rgba(15,23,42,.18));',
  'font-family:var(--swap-chip-font-family,"Microsoft YaHei",system-ui,-apple-system,sans-serif);',
  'font-size:var(--swap-chip-font-size,12px);line-height:var(--swap-chip-line-height,18px);',
  'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
  `@keyframes ${SWAP_CLASS_PREFIX}chip-in{from{opacity:0}to{opacity:1}}`,
  `@keyframes ${SWAP_CLASS_PREFIX}chip-out{from{opacity:1}to{opacity:0}}`,
  `@keyframes ${SWAP_CLASS_PREFIX}ghost-in{from{opacity:0}to{opacity:1}}`,
  `@keyframes ${SWAP_CLASS_PREFIX}ghost-out{from{opacity:1}to{opacity:0}}`,
  `@keyframes ${SWAP_CLASS_PREFIX}ring-pulse{from{opacity:.9;transform:scale(.55)}to{opacity:0;transform:scale(1.7)}}`,
].join('\n');

/* 小工具 */

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

/** 没有 document（node / SSR）时返回 null：所有 DOM 入口据此降级为 no-op。 */
function resolveDocument(container: HTMLElement | undefined): Document | null {
  if (container) return container.ownerDocument;
  const g = globalThis as { document?: Document };
  return g.document ?? null;
}

function toClassList(className: string): string[] {
  return className.split(/\s+/).filter((part) => part.length > 0);
}

function addClass(el: Element, className: string): void {
  toClassList(className).forEach((part) => el.classList.add(part));
}

function removeClass(el: Element, className: string): void {
  toClassList(className).forEach((part) => el.classList.remove(part));
}

/** 摘要文本：按 code point 截断到 max 个字符（Array.from 避免劈开 emoji 代理对），超出补省略号。 */
export function clampSwapText(text: string, max: number = SWAP_TEXT_MAX_CHARS): string {
  const limit = Number.isFinite(max) && max > 0 ? Math.floor(max) : SWAP_TEXT_MAX_CHARS;
  const chars = Array.from(text);
  return chars.length > limit ? `${chars.slice(0, limit).join('')}…` : text;
}

function safeRect(el: Element): { left: number; top: number } {
  try {
    const r = el.getBoundingClientRect();
    return { left: r.left, top: r.top };
  } catch {
    return { left: 0, top: 0 };
  }
}

function normalizeDuration(value: number | undefined, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback;
  return Math.max(0, value);
}

/* 工厂 */

export function createSwapAnimation(options: SwapAnimationOptions = {}): SwapAnimation {
  const doc = resolveDocument(options.container);
  const reducedMotion = options.reducedMotion === true;
  const raf = resolveRaf();
  const caf = resolveCaf();

  let root: HTMLDivElement | null = null;
  let chipsLayer: HTMLDivElement | null = null;
  let ghost: HTMLDivElement | null = null;
  let ghostPosition = { x: 0, y: 0 };
  let ghostRafId: number | null = null;
  let disposed = false;

  /** 本实例派生的全部异步句柄都登记在这里，dispose 时一次性清空。 */
  const frames = new Set<number>();
  const timers = new Set<number>();
  const pendingRemoval = new Map<Element, number>();
  /** 未 settle 的观察者（playSwap 的 resolver）：dispose 时也必须让它们 settle。 */
  const pendingSettles = new Set<() => void>();

  /* ---------------------------- 异步句柄小工具 ---------------------------- */

  function releaseFrame(id: number | null): void {
    if (id === null) return;
    if (caf !== null) caf(id);
    frames.delete(id);
  }

  function scheduleTimeout(fn: () => void, ms: number): number {
    const raw = globalThis.setTimeout(() => {
      timers.delete(id);
      // dispose 之后不再执行任何回调（dispose 里已清空登记项）。
      if (disposed) return;
      try {
        fn();
      } catch {
        /* 收尾中的异常不允许冒泡成 unhandled error */
      }
    }, Math.max(0, ms));
    const id = typeof raw === 'number' ? raw : Number(raw);
    timers.add(id);
    return id;
  }

  function cancelTimer(id: number | null): void {
    if (id === null) return;
    globalThis.clearTimeout(id);
    timers.delete(id);
  }

  /** 请求移除：等 fade 动画自然结束，超时兜底强制移除（动画被掐断也不会残留节点）。 */
  function requestRemoval(el: Element, fallbackMs: number): void {
    const existing = pendingRemoval.get(el);
    if (existing !== undefined) cancelTimer(existing);
    const timerId = scheduleTimeout(() => {
      pendingRemoval.delete(el);
      el.remove();
    }, fallbackMs);
    pendingRemoval.set(el, timerId);
  }

  function removeNow(el: Element | null): void {
    if (!el) return;
    const timerId = pendingRemoval.get(el);
    if (timerId !== undefined) {
      cancelTimer(timerId);
      pendingRemoval.delete(el);
    }
    removeClass(el, `${CHIP_FADE_OUT_CLASS} ${GHOST_OUT_CLASS}`);
    el.remove();
  }

  /** 淡出 + 移除：先撤掉入场动画再挂出场动画（同属性会互相打断）。 */
  function fadeOutAndRemove(el: HTMLElement, fadeOutClass: string, fadeInClass: string, durationMs: number): void {
    removeClass(el, fadeInClass);
    addClass(el, fadeOutClass);
    requestRemoval(el, durationMs + 160);
  }

  /* ------------------------------- 层与样式 ------------------------------- */

  function ensureBaseStyle(document: Document): void {
    if (document.getElementById(BASE_STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = BASE_STYLE_ID;
    style.textContent = BASE_CSS;
    document.head.appendChild(style);
  }

  function ensureLayers(): HTMLDivElement | null {
    if (disposed || doc === null) return null;
    ensureBaseStyle(doc);
    if (root !== null && chipsLayer !== null) return chipsLayer;
    const host = options.container ?? doc.body;
    const nextRoot = doc.createElement('div');
    nextRoot.className = ROOT_CLASS;
    nextRoot.setAttribute('aria-hidden', 'true');
    const nextChips = doc.createElement('div');
    nextChips.className = CHIPS_CLASS;
    nextRoot.appendChild(nextChips);
    host.appendChild(nextRoot);
    root = nextRoot;
    chipsLayer = nextChips;
    return nextChips;
  }

  /** 把视口坐标换算成层内坐标（层是 fixed inset:0，正常情况下两者同一坐标系）。 */
  function resolveAt(at: { x: number; y: number }): { x: number; y: number } {
    const x = Number.isFinite(at.x) ? at.x : 0;
    const y = Number.isFinite(at.y) ? at.y : 0;
    if (root === null) return { x, y };
    const rect = safeRect(root);
    return { x: x - rect.left, y: y - rect.top };
  }

  function createPulseRing(document: Document, x: number, y: number): SVGSVGElement {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('class', PULSE_CLASS);
    svg.setAttribute('width', '56');
    svg.setAttribute('height', '56');
    svg.setAttribute('viewBox', '0 0 56 56');
    // viewBox 中心对准点击点；脉冲动画缩放的是里面那颗 circle（svg 自身只承载定位 transform）。
    svg.style.transform = `translate3d(${(x - 28).toFixed(1)}px, ${(y - 28).toFixed(1)}px, 0)`;
    const circle = document.createElementNS(SVG_NS, 'circle');
    circle.setAttribute('class', RING_CLASS);
    circle.setAttribute('cx', '28');
    circle.setAttribute('cy', '28');
    circle.setAttribute('r', '14');
    circle.setAttribute('stroke-width', '2');
    svg.appendChild(circle);
    return svg;
  }

  function runPulse(svg: SVGSVGElement, durationMs: number): void {
    addClass(svg, `${PULSE_CLASS}--pulse`);
    const onDone = (): void => {
      svg.removeEventListener('animationend', onDone);
      removeNow(svg);
    };
    svg.addEventListener('animationend', onDone);
    // 兜底：动画被掐断（后台标签页 / 元素被移除 / 无 CSS）也必须移除节点。
    requestRemoval(svg, durationMs + 200);
  }

  function createChip(text: string, x: number, y: number): HTMLDivElement | null {
    if (doc === null) return null;
    const chip = doc.createElement('div');
    chip.className = CHIP_CLASS;
    chip.setAttribute('aria-hidden', 'true');
    chip.textContent = clampSwapText(text);
    // 真实几何位置 = 目标点；入场时用 transform 先"拉回"起点，再交给 flyIn 收敛到恒等变换。
    chip.style.transform = `translate3d(${x.toFixed(1)}px, ${y.toFixed(1)}px, 0)`;
    return chip;
  }

  /** 先摆起点、等两帧再交给 flyIn：rAF 回调发生在绘制之前，只等一帧会让起点与首帧合并，观感上芯片"从终点开始"。 */
  function afterPaint(fn: () => void): void {
    if (raf === null) {
      fn();
      return;
    }
    const id = raf(() => {
      frames.delete(id);
      if (disposed) return;
      const second = raf(() => {
        frames.delete(second);
        if (disposed) return;
        fn();
      });
      frames.add(second);
    });
    frames.add(id);
  }

  function flyChip(
    chip: HTMLDivElement,
    from: { x: number; y: number },
    to: { x: number; y: number },
    durationMs: number,
  ): Promise<void> {
    return new Promise<void>((resolve) => {
      let settled = false;
      let frameId: number | null = null;
      let fallbackTimer: number | null = null;

      const finish = (): void => {
        if (settled) return;
        settled = true;
        pendingSettles.delete(finish);
        releaseFrame(frameId);
        frameId = null;
        cancelTimer(fallbackTimer);
        fallbackTimer = null;
        if (!chip.isConnected) {
          // 已经被移除（dispose / 提前清理）：不碰 DOM，但 promise 必须 settle。
          resolve();
          return;
        }
        const size = chip.getBoundingClientRect();
        flyIn(
          chip,
          { left: from.x, top: from.y, width: size.width, height: size.height },
          { left: to.x, top: to.y, width: size.width, height: size.height },
          { durationMs, reducedMotion, easing: easeForChip },
        )
          .catch(() => {
            /* flyIn 不会 reject；防御性兜底，避免未处理拒绝 */
          })
          .finally(() => {
            if (chip.isConnected) fadeOutAndRemove(chip, CHIP_FADE_OUT_CLASS, CHIP_FADE_IN_CLASS, CHIP_FADE_OUT_MS);
          })
          .finally(() => resolve());
      };

      // 与 flyIn 内部的 2s 兜底同级的"外层死线"：无论如何这套动画都会 settle。
      fallbackTimer = scheduleTimeout(finish, durationMs + 600);
      frameId = raf === null ? null : raf(finish);
      if (frameId !== null) frames.add(frameId);
      pendingSettles.add(finish);
    });
  }

  /* -------------------------------- 对外 API ------------------------------ */

  return {
    pulseSelection(at: { x: number; y: number }): void {
      if (disposed || doc === null) return;
      const layer = ensureLayers();
      if (layer === null) return;
      const point = resolveAt(at);
      const svg = createPulseRing(doc, point.x, point.y);
      layer.appendChild(svg);
      if (reducedMotion) {
        removeNow(svg);
        return;
      }
      runPulse(svg, DEFAULT_PULSE_DURATION_MS);
    },

    playSwap(
      a: { x: number; y: number; text: string },
      b: { x: number; y: number; text: string },
      options2?: { durationMs?: number },
    ): Promise<void> {
      const layer = ensureLayers();
      if (layer === null) return Promise.resolve();

      const durationMs = normalizeDuration(options2?.durationMs, DEFAULT_SWAP_DURATION_MS);
      const fromA = resolveAt(a);
      const fromB = resolveAt(b);
      const chipA = createChip(a.text, fromA.x, fromA.y);
      const chipB = createChip(b.text, fromB.x, fromB.y);
      if (chipA === null || chipB === null) return Promise.resolve();

      // a 在下、b 在上：两者在中途交错时是 b 压着 a，符合"被点中的那一个更靠前"的观感。
      chipB.style.zIndex = '1';
      layer.appendChild(chipA);
      layer.appendChild(chipB);
      addClass(chipA, CHIP_FADE_IN_CLASS);
      addClass(chipB, CHIP_FADE_IN_CLASS);

      if (reducedMotion || raf === null) {
        // 降级路径：不做位移动画，但仍然把临时节点清干净。
        const quick = CHIP_FADE_OUT_MS;
        fadeOutAndRemove(chipA, CHIP_FADE_OUT_CLASS, CHIP_FADE_IN_CLASS, quick);
        fadeOutAndRemove(chipB, CHIP_FADE_OUT_CLASS, CHIP_FADE_IN_CLASS, quick);
        return new Promise<void>((resolve) => {
          let settled = false;
          const finish = (): void => {
            if (settled) return;
            settled = true;
            pendingSettles.delete(finish);
            cancelTimer(timerId);
            removeNow(chipA);
            removeNow(chipB);
            resolve();
          };
          const timerId = scheduleTimeout(finish, quick + 40);
          pendingSettles.add(finish);
        });
      }

      // 起点态：真实几何位置 = 起点，视觉上"从终点被拉回来"（transition 常驻 none，位移只由 rAF 驱动）。
      chipA.style.transform = `translate3d(${fromA.x.toFixed(1)}px, ${fromA.y.toFixed(1)}px, 0)`;
      chipB.style.transform = `translate3d(${fromB.x.toFixed(1)}px, ${fromB.y.toFixed(1)}px, 0)`;

      return new Promise<void>((resolve) => {
        let settled = false;
        let deadline: number | null = null;

        /** 唯一的收尾入口：无论走哪条路，节点都移除、promise 只 resolve 一次。 */
        const settle = (): void => {
          if (settled) return;
          settled = true;
          pendingSettles.delete(settle);
          cancelTimer(deadline);
          deadline = null;
          removeNow(chipA);
          removeNow(chipB);
          resolve();
        };
        pendingSettles.add(settle);

        afterPaint(() => {
          // 真实几何位置改到对方那边 —— flyIn 会用 transform 把我们"拉回"起点再飞过来。
          chipA.style.transform = `translate3d(${fromB.x.toFixed(1)}px, ${fromB.y.toFixed(1)}px, 0)`;
          chipB.style.transform = `translate3d(${fromA.x.toFixed(1)}px, ${fromA.y.toFixed(1)}px, 0)`;
          // 强制结算一次布局，确保"起点观感"真的被绘制过。
          chipA.getBoundingClientRect();
          chipB.getBoundingClientRect();

          Promise.all([flyChip(chipA, fromA, fromB, durationMs), flyChip(chipB, fromB, fromA, durationMs)]).then(
            settle,
            settle,
          );
        });

        // 外层死线：即使 afterPaint 的回调没跑（dispose / 无帧循环），promise 也一定会 settle。
        deadline = scheduleTimeout(settle, durationMs + 900);
      });
    },

    showPendingGhost(text: string, at: { x: number; y: number }): void {
      if (disposed || doc === null) return;
      const layer = ensureLayers();
      if (layer === null) return;

      removeNow(ghost);
      ghost = doc.createElement('div');
      ghost.className = GHOST_CLASS;
      ghost.setAttribute('aria-hidden', 'true');
      const span = doc.createElement('span');
      span.className = GHOST_TEXT_CLASS;
      span.textContent = clampSwapText(text);
      ghost.appendChild(span);

      ghostPosition = resolveAt(at);
      applyGhostTransform();
      if (reducedMotion) {
        removeClass(ghost, GHOST_IN_CLASS);
      } else {
        addClass(ghost, GHOST_IN_CLASS);
      }
      layer.appendChild(ghost);
    },

    movePendingGhost(at: { x: number; y: number }): void {
      if (disposed || ghost === null) return;
      ghostPosition = resolveAt(at);
      if (raf === null) {
        applyGhostTransform();
        return;
      }
      if (ghostRafId !== null) return; // 同帧多次调用只保留最后一次坐标
      const id = raf(() => {
        frames.delete(id);
        ghostRafId = null;
        if (disposed) return;
        applyGhostTransform();
      });
      ghostRafId = id;
      frames.add(id);
    },

    hidePendingGhost(): void {
      if (disposed) return;
      // 先结算挂起的 rAF，否则最后一次 move 会丢、隐藏瞬间位置回跳。
      releaseFrame(ghostRafId);
      ghostRafId = null;
      const target = ghost;
      ghost = null;
      if (target === null) return;
      applyGhostTransformTo(target, ghostPosition);
      fadeOutAndRemove(target, GHOST_OUT_CLASS, GHOST_IN_CLASS, GHOST_FADE_OUT_MS);
    },

    dispose(): void {
      if (disposed) return;
      disposed = true;

      frames.forEach((id) => {
        if (caf !== null) caf(id);
      });
      frames.clear();
      ghostRafId = null;

      timers.forEach((id) => globalThis.clearTimeout(id));
      timers.clear();
      pendingRemoval.clear();

      // 未 settle 的动画 promise 必须立刻 settle（各 resolver 内部有 isConnected 守卫）
      const settles = Array.from(pendingSettles);
      pendingSettles.clear();
      settles.forEach((settle) => {
        try {
          settle();
        } catch {
          /* resolver 抛错不影响 dispose 的清理语义 */
        }
      });

      // 4) 移除全部临时节点（含未登记的漏网节点：整层直接摘掉）
      ghost = null;
      chipsLayer = null;
      const layer = root;
      root = null;
      if (layer) {
        layer.querySelectorAll(`.${CHIP_CLASS}, .${PULSE_CLASS}, .${GHOST_CLASS}`).forEach((el) => el.remove());
        layer.remove();
      }
    },
  };

  function applyGhostTransform(): void {
    if (ghost === null) return;
    applyGhostTransformTo(ghost, ghostPosition);
  }

  function applyGhostTransformTo(el: HTMLElement, position: { x: number; y: number }): void {
    el.style.transform = `translate3d(${(position.x + GHOST_OFFSET.x).toFixed(1)}px, ${(
      position.y + GHOST_OFFSET.y
    ).toFixed(1)}px, 0)`;
  }
}

/** 芯片飞行用的缓动（与 CSS 的 CHIP_EASING 同一条贝塞尔，保证 transform 过渡与 flyIn 不打架）。 */
const easeForChip = ((): ((t: number) => number) => {
  // cubic-bezier(0.22, 0.61, 0.36, 1) 的数值解（Newton + 二分兜底），与 CSS 观感一致。
  const x1 = 0.22;
  const y1 = 0.61;
  const x2 = 0.36;
  const y2 = 1;
  const bezier = (t: number, p1: number, p2: number): number => {
    const u = 1 - t;
    return 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t;
  };
  const slope = (t: number, p1: number, p2: number): number => {
    const u = 1 - t;
    return 3 * u * u * p1 + 6 * u * t * (p2 - p1) + 3 * t * t * (1 - p2);
  };
  return (t: number): number => {
    if (t <= 0) return 0;
    if (t >= 1) return 1;
    let guess = t;
    for (let i = 0; i < 6; i += 1) {
      const x = bezier(guess, x1, x2) - t;
      if (Math.abs(x) < 1e-5) break;
      const d = slope(guess, x1, x2);
      if (d === 0) break;
      guess -= x / d;
    }
    return bezier(guess, y1, y2);
  };
})();
