/**
 * 右键菜单的纯模型层：**定位夹取**与**键盘导航**。
 *
 * 设计原则：
 * - 本文件不引入 React / DOM / Univer，只有数学与数组遍历 → 可以直接在 node 环境下单测
 *   （`tests/unit/context-menu-model.test.ts`，无需 jsdom）；
 * - 全部是**纯函数**：不改写入参（`items` / `anchor` / `menu` / `viewport` 均只读），
 *   相同输入必定得到相同输出，无隐藏状态、无副作用；
 * - 非有限数（NaN / Infinity）在入口被降级为安全值，避免把 NaN 透传到 `style.left`
 *   上导致整块菜单"消失"。
 *
 * 坐标约定：一律是**视口坐标**（等价于 `clientX/clientY` 与
 * `window.innerWidth/window.innerHeight`）。组件侧用 `position: fixed` + `left/top`
 * 消费 `clampMenuPosition` 的返回值。
 */

export interface MenuItemSpec {
  id: string;
  label: string;
  /** 快捷键提示，如 'Ctrl+C' */
  shortcut?: string;
  disabled?: boolean;
  /** 危险操作（删除/清空），用红色呈现 */
  danger?: boolean;
  /** 分组分隔：与上一项之间画一条分隔线 */
  separatorBefore?: boolean;
  /**
   * 悬停说明（`title`）。主要用于"为什么这一项是灰的"——
   * 例如多块选区时"与工作区互换/粘贴"只支持单块，这里写清原因与做法。
   */
  title?: string;
  /** 次要说明文字，直接显示在菜单项右侧（比 title 更显眼，用于不可用原因的短句） */
  hint?: string;
}

export interface Point {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface Viewport {
  width: number;
  height: number;
}

/** 菜单与视口边缘的最小留白（px），也是 `clampMenuPosition` 的 `margin` 默认值。 */
export const MENU_VIEWPORT_MARGIN = 8;

/* -------------------------------------------------------------------------- */
/* 定位                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * 把菜单定位到点击点，同时保证不超出视口（贴边时向左/向上翻转，留 `margin` 边距）。
 *
 * 单轴策略（x / y 各自独立，互不影响，所以"右侧翻转、下方不翻转"是正常结果）：
 * 1. **优先不翻转**：`锚点 + 菜单尺寸 <= 视口 - margin` 时，菜单左上角就落在点击点上
 *    （点击点自身小于 margin，例如贴左边缘或负坐标，则抬到 margin，宁可整体右/下移）；
 * 2. 该方向放不下 → **翻转**（菜单贴到点击点的左/上侧），并把结果再夹到不超过
 *    `视口 - margin - 菜单尺寸`：点击点本身已经压进边距时，翻转位也可能越过边界；
 * 3. 连翻转都放不下（点击点离两侧都近、菜单超过半屏）→ 取**最大完整可见位置**
 *    `视口 - margin - 菜单尺寸`；若该值仍小于 margin（菜单比"视口 - 2*margin"还大，
 *    完整可见已不可能），退化为贴着左上 margin，至少保证菜单**前缘**
 *    （左侧图标位 + 文字起始处）永远可见。
 *
 * 推论：只要 `菜单尺寸 + 2 * margin <= 视口`，返回值必定让菜单**完整可见**、
 * 且四边都留够 `margin`；空间不足时也绝不把菜单推到视口外。
 */
export function clampMenuPosition(
  anchor: Point,
  menu: Size,
  viewport: Viewport,
  margin: number = MENU_VIEWPORT_MARGIN,
): Point {
  const gap = normalizeMargin(margin);
  const view: Size = {
    width: Math.max(0, finite(viewport.width, 0)),
    height: Math.max(0, finite(viewport.height, 0)),
  };
  const box: Size = {
    width: Math.max(0, finite(menu.width, 0)),
    height: Math.max(0, finite(menu.height, 0)),
  };

  return {
    x: clampAxis(finite(anchor.x, gap), box.width, view.width, gap),
    y: clampAxis(finite(anchor.y, gap), box.height, view.height, gap),
  };
}

/**
 * 单轴夹取：见 `clampMenuPosition` 的三步策略。
 * @param position 锚点在该轴上的坐标（已经过 NaN 降级）
 * @param boxSize 菜单在该轴上的尺寸
 * @param viewSize 视口在该轴上的尺寸
 * @param gap 边距
 */
function clampAxis(position: number, boxSize: number, viewSize: number, gap: number): number {
  // 完整可见时允许的最大起点（可能为负：说明"视口 - 2*gap"装不下菜单）
  const maxStart = viewSize - gap - boxSize;

  // 1) 该方向放得下 → 直接贴着点击点；点击点太靠边则抬到边距处
  if (position + boxSize <= viewSize - gap) return Math.max(gap, position);

  // 2) 放不下 → 向左/向上翻转；翻转位若越过右/下边距，夹回最大起点
  const flipped = position - boxSize;
  if (flipped >= gap) return Math.max(gap, Math.min(flipped, maxStart));

  // 3) 翻转也放不下 → 最大完整可见位置；菜单比视口还大时贴边距，保住前缘
  return Math.max(gap, maxStart);
}

/** 边距降级：负数 / NaN / Infinity 一律回落到默认边距。 */
function normalizeMargin(margin: number): number {
  if (!Number.isFinite(margin) || margin < 0) return MENU_VIEWPORT_MARGIN;
  return margin;
}

/** 数值降级：非有限数替换为 fallback，避免 NaN 在几何计算里扩散。 */
function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback;
}

/* -------------------------------------------------------------------------- */
/* 键盘导航                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 在菜单项里做键盘导航：返回**下一个可选项**的下标（跳过 `disabled`），到边界时循环。
 *
 * 语义约定（单测逐条覆盖）：
 * - 空数组 / 全部 disabled → `-1`（组件据此"不移动焦点"）；
 * - 一定"先移动再看"：`from` 自身可选中也不会原地返回，除非它是**唯一**可选项
 *   （此时循环一整圈后回到它自己，符合"到边界循环"的直觉）；
 * - `direction = 1`（↓）沿数组向下、到头回到开头；`direction = -1`（↑）反之；
 * - `from` 非法（`-1`：尚无焦点、非整数、越界）时：↓ 视作从"开头之前"出发 → 首个可选项，
 *   ↑ 视作从"结尾之后"出发 → 末个可选项。
 */
export function nextEnabledIndex(items: MenuItemSpec[], from: number, direction: 1 | -1): number {
  const count = items.length;
  if (count === 0) return -1;

  const step: 1 | -1 = direction === -1 ? -1 : 1;
  const start = normalizeStart(from, count, step);

  // 最多走一整圈：每个下标恰好被检查一次；全是 disabled 时循环结束后返回 -1
  for (let offset = 1; offset <= count; offset += 1) {
    const index = (((start + step * offset) % count) + count) % count;
    const item = items[index];
    if (item !== undefined && item.disabled !== true) return index;
  }
  return -1;
}

/** 首个可选项下标；没有可选项（空数组 / 全部 disabled）时返回 -1。 */
export function firstEnabledIndex(items: MenuItemSpec[]): number {
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    if (item !== undefined && item.disabled !== true) return index;
  }
  return -1;
}

/**
 * 把 `from` 规整成"可以迈出第一步"的起点下标：
 * - 合法下标原样返回；
 * - 非法/负值 → ↓ 用 `-1`（第一步落在 0）、↑ 用 `0`（第一步落在 `count - 1`）；
 * - 越界（>= count）→ ↓ 用 `count - 1`（第一步回到 0）、↑ 用 `0`（第一步落在末尾）。
 */
function normalizeStart(from: number, count: number, direction: 1 | -1): number {
  if (!Number.isFinite(from)) return direction === 1 ? -1 : 0;
  const value = Math.trunc(from);
  if (value < 0) return direction === 1 ? -1 : 0;
  if (value >= count) return direction === 1 ? count - 1 : 0;
  return value;
}
