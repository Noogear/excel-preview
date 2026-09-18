/**
 * 「点击交换」交互控制器 —— **纯状态机**。
 *
 * 设计要点（与 `drag-controller` 的职责分界）：
 *  - 本模块**不 touch DOM / Univer / window**，只有 `select` / `getMode` 这类纯入口，
 *    因此可以在 node 环境的 Vitest 里逐条验证转移表（见 `tests/unit/click-swap.test.ts`）。
 *  - "画"的事情全部交给调用方：
 *      · 选中反馈 → `FWorksheet.setActiveSelection(range)`（Univer 自己画第一个选中框）；
 *      · 内容芯片飞行 → `createSwapAnimation().playSwap(a, b)`（坐标来自点击事件的 clientX/clientY）。
 *    本模块只负责"什么时候该发生什么"，不负责"长什么样"。
 *
 * 状态只有三个：
 *
 *   idle ──select(s)──▶ pending(s) ──select(s)──▶ idle            （同一方 → 取消选中）
 *     ▲                    │  └──select(t≠s)──▶ pending(t) | idle  （尺寸不一致 → 换 pending）
 *     └────clear()─────────┘                                       （尺寸一致 → onSwap + 清空）
 *
 * 另外有一个正交的 `enabled` 开关（`setEnabled`）：模式仍是 `click-swap`，
 * 但临时不接受点击（例如正在播放上一次交换的动画时由上层闸住）。
 *
 * 三种交互模式（用户可切换）：
 * - `select`      选择模式：只做原生选择，不长按拖动、不点击互换；
 * - `drag`        拖拽模式：**长按**（默认 400ms 且位移 < 4px）后拖动可搬运/互换内容，
 *                 快速拖动仍是原生框选——这样"框选"和"搬运"不会互相打架；
 * - `click-swap`  点击互换模式：点一方再点另一方即互换（可跨表、可与工作区条目互换）。
 */

export type InteractionMode = 'select' | 'drag' | 'click-swap';

/** 模式顺序即 UI 里的展示顺序 */
export const INTERACTION_MODES: readonly InteractionMode[] = ['select', 'drag', 'click-swap'];

/** 模式的中文名（UI 与日志共用同一份，避免两处写死） */
export const INTERACTION_MODE_LABELS: Readonly<Record<InteractionMode, string>> = {
  select: '选择',
  drag: '拖拽',
  'click-swap': '点击互换',
};

/** 每种模式的一句话说明（鼠标悬停提示用） */
export const INTERACTION_MODE_HINTS: Readonly<Record<InteractionMode, string>> = {
  select: '选择模式：点选 / 框选单元格，不会触发搬运或互换',
  drag: '拖拽模式：按住约 0.4 秒再拖动，即可把内容搬到别处或与目标互换；快速拖动仍是框选',
  'click-swap': '点击互换模式：先点一方，再点另一方，两边内容互换（可与右侧工作区条目互换）',
};

/** 把任意值收敛成合法模式（会话里可能存着旧值 / 脏值） */
export function normalizeInteractionMode(value: unknown): InteractionMode {
  return typeof value === 'string' && (INTERACTION_MODES as readonly string[]).includes(value)
    ? (value as InteractionMode)
    : DEFAULT_INTERACTION_MODE;
}

/** 参与交换的一方：可以是表格单元格区域，也可以是工作区条目 */
export interface SwapSide {
  kind: 'cell' | 'workspace';
  /** cell 时必填 */
  sheetId?: string;
  a1?: string;
  rows: number;
  cols: number;
  /** workspace 时必填 */
  itemId?: string;
  label: string;
}

export interface ClickSwapEvents {
  /** 选中集合变化（0/1/2 个） */
  onSelectionChange?: (sides: SwapSide[]) => void;
  /** 选满两个且允许交换 */
  onSwap?: (a: SwapSide, b: SwapSide) => void;
  /** 不可交换（尺寸不一致 / 同一位置 / 与自身交换） */
  onReject?: (reason: string, pending: SwapSide, incoming: SwapSide) => void;
}

export interface ClickSwapController {
  getMode(): InteractionMode;
  setMode(mode: InteractionMode): void;
  isEnabled(): boolean;
  /** 已选中的一方（0/1 个；选满两个会立即触发 onSwap 并清空） */
  peek(): SwapSide | null;
  /** 点击/选择某一方 */
  select(side: SwapSide): void;
  clear(): void;
  dispose(): void;
}

/** 默认模式是 `drag`：必须显式 `setMode('click-swap')` 才会开始累积选中。 */
export const DEFAULT_INTERACTION_MODE: InteractionMode = 'drag';

/** 尺寸不一致时的拒绝原因（上层可直接拿去 toast，前缀稳定便于断言）。 */
export const REJECT_SIZE_MISMATCH = '尺寸不一致：两边的行列数必须相同才能交换';
/** 同一位置/与自身交换时的拒绝原因。 */
export const REJECT_SAME_SIDE = '同一位置：不能与自己交换';

/** 让调用方不必手写 `setEnabled`（可选增强，接口本身不含它）。 */
export interface ClickSwapControllerWithGate extends ClickSwapController {
  /** 临时闸住/放开点击（模式不变；闸住时清掉未完成的 pending）。 */
  setEnabled(enabled: boolean): void;
}

/** `a1` 的规范键：大小写/空白不敏感，`B2:D5` 与 ` b2:d5 ` 视为同一位置。 */
function normalizeA1(a1: string | undefined): string {
  return (a1 ?? '').trim().toUpperCase();
}

/**
 * 是否是"同一方"：
 * 比较 `kind` + 全部身份字段（`a1` / `sheetId` / `itemId`）。
 * `rows` / `cols` 属于**尺寸**而非身份，不参与这里；缺失的身份字段按空串处理，
 * 因此两个"没有任何身份字段的 cell"会被视为同一方 —— 这符合"与自身交换"的语义，
 * 且不会误判正常数据（正常 cell 必有 `sheetId` + `a1`）。
 */
function isSameSide(a: SwapSide, b: SwapSide): boolean {
  return (
    a.kind === b.kind &&
    normalizeA1(a.a1) === normalizeA1(b.a1) &&
    (a.sheetId ?? '') === (b.sheetId ?? '') &&
    (a.itemId ?? '') === (b.itemId ?? '')
  );
}

/** 尺寸是否一致（工作区条目与单元格**可以**互相交换，所以这里不看 kind）。 */
function isSameSize(a: SwapSide, b: SwapSide): boolean {
  return a.rows === b.rows && a.cols === b.cols;
}

function sizeLabel(side: SwapSide): string {
  return `${side.rows}×${side.cols}`;
}

export function createClickSwapController(events: ClickSwapEvents): ClickSwapControllerWithGate {
  let mode: InteractionMode = DEFAULT_INTERACTION_MODE;
  let enabled = true;
  let pending: SwapSide | null = null;
  let disposed = false;

  /** 统一出口：dispose 之后一个回调都不许再发。 */
  function emitSelectionChange(sides: SwapSide[]): void {
    if (disposed) return;
    events.onSelectionChange?.(sides);
  }

  function emitSwap(a: SwapSide, b: SwapSide): void {
    if (disposed) return;
    if (pending !== null) {
      pending = null;
      emitSelectionChange([]);
    }
    events.onSwap?.(a, b);
  }

  function emitReject(reason: string, prev: SwapSide, incoming: SwapSide): void {
    if (disposed) return;
    events.onReject?.(reason, prev, incoming);
  }

  /** 清空选中（`clear()` / `setEnabled(false)` / `setMode()` 切走时共用）。 */
  function resetSelection(): void {
    if (pending === null) return;
    pending = null;
    emitSelectionChange([]);
  }

  function handleSelect(incoming: SwapSide): void {
    if (pending === null) {
      pending = incoming;
      emitSelectionChange([incoming]);
      return;
    }

    const first = pending;

    // 1) 同一方 → 取消选中（回到 0 个），不触发 onSwap / onReject
    if (isSameSide(first, incoming)) {
      pending = null;
      emitSelectionChange([]);
      return;
    }

    // 2) 尺寸不一致 → 拒绝，并把 pending 换成新的这一方（用户改主意时不必先取消）
    if (!isSameSize(first, incoming)) {
      pending = incoming;
      emitSelectionChange([incoming]);
      emitReject(
        `${REJECT_SIZE_MISMATCH}（已选 ${sizeLabel(first)}，本次 ${sizeLabel(incoming)}）`,
        first,
        incoming,
      );
      return;
    }

    // 3) 尺寸一致 → 交换（emitSwap 内部会先清空选中，保证 onSelectionChange 序列为 1 → 0）
    emitSwap(first, incoming);
  }

  return {
    getMode(): InteractionMode {
      return mode;
    },

    setMode(next: InteractionMode): void {
      if (disposed || next === mode) return;
      mode = next;
      // 切走时丢掉未完成的选中，避免"半途的选择"在新模式里复活。
      resetSelection();
    },

    isEnabled(): boolean {
      return enabled && mode === 'click-swap' && !disposed;
    },

    setEnabled(next: boolean): void {
      if (disposed || next === enabled) return;
      enabled = next;
      if (!enabled) resetSelection();
    },

    peek(): SwapSide | null {
      return pending;
    },

    select(side: SwapSide): void {
      // 模式不对 / 被闸住 / 已销毁 → 一律忽略（不报错、不发回调）。
      if (disposed || mode !== 'click-swap' || !enabled) return;
      handleSelect(side);
    },

    clear(): void {
      if (disposed) return;
      resetSelection();
    },

    dispose(): void {
      if (disposed) return;
      disposed = true;
      pending = null;
    },
  };
}
