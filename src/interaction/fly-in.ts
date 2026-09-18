/**
 * 落位动画（纯 TS + DOM，不依赖 React / Univer）。
 *
 * 语义：把 element 从 `from` 矩形**视觉上**动画到 `to` 矩形。
 * 为了让动画不触发 layout，我们**不动** element 的真实几何位置，只写 transform：
 *
 *     translate(dx, dy) scale(s)
 *
 * 即调用方应当已经把 element 的静态位置/尺寸设成 `to`（例如落点格子的位置），
 * `flyIn` 负责先把它"拉回" `from` 的观感，再用 rAF 插值回 `to`。
 * 动画收敛时 `s → to.width/from.width`、`translate → (0,0)`，也就是回到静止状态，
 * 因此结束时会**清空 transform**（`element.style.transform = ''`），不留任何残留。
 */

export interface DOMRectLike {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface FlyInOptions {
  durationMs?: number; // 默认 260
  easing?: (t: number) => number; // 默认 easeOutCubic
  reducedMotion?: boolean;
}

export const DEFAULT_FLY_IN_DURATION_MS = 260;

export function easeOutCubic(t: number): number {
  if (t <= 0) return 0;
  if (t >= 1) return 1;
  const p = 1 - t;
  return 1 - p * p * p;
}

export function easeOutBack(t: number): number {
  // 端点直接短路：多项式形式在 t=0 会留下 2.2e-16 的浮点噪声，
  // 而调用方（以及单测）期望精确的 f(0)=0 / f(1)=1。
  if (t <= 0) return 0;
  if (t >= 1) return 1;
  const c1 = 1.70158;
  const c3 = c1 + 1;
  const p = t - 1;
  return 1 + c3 * p * p * p + c1 * p * p;
}

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

/** h 为 0（例如隐藏行）时退化为 1，杜绝 Infinity / NaN。 */
function safeRatio(toValue: number, fromValue: number): number {
  if (!Number.isFinite(fromValue) || fromValue === 0) return 1;
  if (!Number.isFinite(toValue)) return 1;
  return toValue / fromValue;
}

/**
 * 把 element 从 from 矩形动画到 to 矩形（用 transform 实现，不触发 layout）。
 *
 * - rAF 驱动；结束时清空 transform 并 resolve；
 * - `reducedMotion` 为 true 时立即 resolve（不做动画，也不写 transform）；
 * - 用 `transitionend` + 超时兜底：即使帧循环被浏览器掐断（后台标签页/元素被移除），
 *   promise 也一定会 settle，绝不留下悬挂的 rAF 或永不 resolve 的 promise。
 */
export function flyIn(
  element: HTMLElement,
  from: DOMRectLike,
  to: DOMRectLike,
  options: FlyInOptions = {},
): Promise<void> {
  const durationMs = options.durationMs ?? DEFAULT_FLY_IN_DURATION_MS;
  const easing = options.easing ?? easeOutCubic;

  if (options.reducedMotion === true) {
    return Promise.resolve();
  }

  const raf = resolveRaf();
  const caf = resolveCaf();
  if (raf === null || caf === null) {
    // 非浏览器环境：没有帧循环，直接落位。
    element.style.transform = '';
    return Promise.resolve();
  }

  // 起始观感：真实位置在 to，所以先把"缩放差 + 位移差"补回去，看起来就在 from。
  const startScale = safeRatio(from.width, to.width);
  const startTranslateX = from.left - to.left;
  const startTranslateY = from.top - to.top;
  const endScale = safeRatio(to.width, from.width);

  const lift = (s: number): string =>
    `translate3d(${(startTranslateX * (1 - s)).toFixed(2)}px, ${(startTranslateY * (1 - s)).toFixed(2)}px, 0) scale(${(
      startScale +
      (1 - startScale) * s
    ).toFixed(4)})`;

  element.style.transform = lift(0);

  return new Promise<void>((resolve) => {
    let frameId: number | null = null;
    let fallbackTimer: number | null = null;
    let settleTimer: number | null = null;
    let startTime: number | null = null;
    let settled = false;

    const cleanup = (): void => {
      if (frameId !== null) {
        caf(frameId);
        frameId = null;
      }
      if (fallbackTimer !== null) {
        globalThis.clearTimeout(fallbackTimer);
        fallbackTimer = null;
      }
      if (settleTimer !== null) {
        globalThis.clearTimeout(settleTimer);
        settleTimer = null;
      }
      element.removeEventListener('transitionend', onTransitionEnd);
      element.removeEventListener('transitioncancel', onTransitionEnd);
      // 注意：finish() 之后不再有任何 rAF / 计时器 / 监听器可回调，promise 只 resolve 一次。
    };

    const finish = (): void => {
      if (settled) return;
      settled = true;
      cleanup();
      // 清空 transform：动画结束时它已经是恒等变换，清掉可以避免残留 will-change/层。
      element.style.transform = '';
      resolve();
    };
    function onTransitionEnd(): void {
      finish();
    }

    // 兜底：万一 transitionend 不来（元素被移除、transition 被覆盖），也别永久挂着。
    const totalMs = Math.min(2_000, durationMs + 200);
    // 显式收窄：@types/node 下 setTimeout 返回 Timeout，浏览器下返回 number。
    const rawTimerId = globalThis.setTimeout(() => {
      fallbackTimer = null;
      finish();
    }, totalMs);
    fallbackTimer = typeof rawTimerId === 'number' ? rawTimerId : Number(rawTimerId);

    /**
     * 第二道兜底（纯保险）：收尾本身是**也**通过 rAF 调度的（见下面的 `raf(finish)`）。
     * 如果帧循环在最后一帧之后停摆（后台标签页、宿主主动掐掉 rAF），那个回调永远不会到，
     * promise 就会悬挂。这里挂一个比动画长得多、但一定跑得到的计时器，保证必然 settle。
     * 正常路径下它在 finish() 里被清掉，不会多跑一次。
     */
    const rawSettleId = globalThis.setTimeout(() => {
      settleTimer = null;
      finish();
    }, totalMs + 800);
    settleTimer = typeof rawSettleId === 'number' ? rawSettleId : Number(rawSettleId);

    element.addEventListener('transitionend', onTransitionEnd, { once: true });
    element.addEventListener('transitioncancel', onTransitionEnd, { once: true });

    const step = (time: number): void => {
      frameId = null;
      if (startTime === null) startTime = time;
      const elapsed = time - startTime;
      const raw = durationMs <= 0 ? 1 : Math.min(1, elapsed / durationMs);
      const eased = easing(Math.min(1, Math.max(0, raw)));
      element.style.transform = lift(eased);
      if (raw < 1) {
        frameId = raf(step);
      } else {
        // 让浏览器先按最后一帧完成绘制，再统一收尾（清 transform + resolve）。
        raf(finish);
      }
    };

    frameId = raf(step);
  });
}
