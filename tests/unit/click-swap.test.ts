/**
 * 「点击交换」状态机单测（Vitest，node 环境，**无 DOM**）。
 *
 * 这里**只**测纯逻辑：`src/interaction/click-swap.ts` 不允许出现 document / window，
 * 因此不需要 jsdom，也不需要任何桩。视觉部分（脉冲环 / 芯片飞行 / 幽灵）由
 * `swap-animation.ts` 负责，它只依赖指针坐标，属于 e2e 覆盖范围。
 *
 * 覆盖清单（与交付要求逐条对应）：
 *  1. 模式目录与归一化（三种模式的顺序/标签/说明，非法值回落默认）
 *  2. 模式关闭时忽略点击（默认 drag；select 模式；setEnabled(false)；切回 drag 时丢掉未完成选中）
 *  3. 第一次选中 → 第二次交换（onSwap(pending, incoming) 且清空，可立刻开始下一轮）
 *  4. 同一方二次点击 → 取消选中（不触发 onSwap / onReject）
 *  5. 尺寸不一致 → onReject 且 pending 被替换为新的这一方
 *  6. 工作区 ↔ 单元格互交
 *  7. onSelectionChange 序列（0→1→0；尺寸不一致时是 1→1 的"换人"）
 *  8. dispose 与无回调等边界
 */
import { describe, expect, it, vi } from 'vitest';

import {
  DEFAULT_INTERACTION_MODE,
  INTERACTION_MODES,
  INTERACTION_MODE_HINTS,
  INTERACTION_MODE_LABELS,
  createClickSwapController,
  normalizeInteractionMode,
  type ClickSwapControllerWithGate,
  type InteractionMode,
  type SwapSide,
} from '../../src/interaction/click-swap';

/* -------------------------------------------------------------------------- */
/* 脚手架                                                                      */
/* -------------------------------------------------------------------------- */

function cell(a1: string, rows: number, cols: number, sheetId = 'sheet-1', label?: string): SwapSide {
  return { kind: 'cell', sheetId, a1, rows, cols, label: label ?? a1 };
}

function workspace(itemId: string, rows: number, cols: number, label?: string): SwapSide {
  return { kind: 'workspace', itemId, rows, cols, label: label ?? itemId };
}

interface Harness {
  controller: ClickSwapControllerWithGate;
  onSelectionChange: ReturnType<typeof vi.fn>;
  onSwap: ReturnType<typeof vi.fn>;
  onReject: ReturnType<typeof vi.fn>;
  /** 每次 onSelectionChange 的"个数"序列，用来断言 0→1→0 这类形状 */
  counts(): number[];
  /** 每次 onSelectionChange 传进来的 sides 快照（按 label 记录，避免持有可变引用） */
  labels(): string[][];
}

/** 默认建一个**已进入 click-swap 模式**的控制器（模式本身的开关单独测）。 */
function makeHarness(options: { mode?: InteractionMode } = {}): Harness {
  const onSelectionChange = vi.fn();
  const onSwap = vi.fn();
  const onReject = vi.fn();
  const controller = createClickSwapController({ onSelectionChange, onSwap, onReject });
  const mode = options.mode ?? 'click-swap';
  if (mode !== DEFAULT_INTERACTION_MODE) controller.setMode(mode);

  const selectionCalls = (): SwapSide[][] =>
    onSelectionChange.mock.calls.map((call) => call[0] as SwapSide[]);

  return {
    controller,
    onSelectionChange,
    onSwap,
    onReject,
    counts: () => selectionCalls().map((sides) => sides.length),
    labels: () => selectionCalls().map((sides) => sides.map((side) => side.label)),
  };
}

/* -------------------------------------------------------------------------- */
/* 1. 模式目录与归一化 + 2. 模式开关                                              */
/* -------------------------------------------------------------------------- */

describe('模式目录与模式开关', () => {
  it('三种模式顺序、中文名与说明齐全；合法值原样返回，非法/脏值回落到默认', () => {
    expect(INTERACTION_MODES).toEqual(['select', 'drag', 'click-swap']);
    for (const mode of INTERACTION_MODES) {
      expect(INTERACTION_MODE_LABELS[mode], `${mode} 要有中文名`).toBeTruthy();
      expect(INTERACTION_MODE_HINTS[mode], `${mode} 要有说明文案`).toBeTruthy();
      expect(normalizeInteractionMode(mode)).toBe(mode);
    }
    for (const dirty of [undefined, null, 'keyboard', 7, { mode: 'drag' }]) {
      expect(normalizeInteractionMode(dirty)).toBe(DEFAULT_INTERACTION_MODE);
    }
  });

  it('默认 drag 与 select 模式下 select() 一律忽略（不累积、不通知、不报错），切到 click-swap 后立刻累积', () => {
    const dragMode = makeHarness({ mode: 'drag' });
    expect(DEFAULT_INTERACTION_MODE).toBe('drag');
    expect(dragMode.controller.getMode()).toBe('drag');
    expect(dragMode.controller.isEnabled()).toBe(false);

    dragMode.controller.select(cell('A1', 2, 2));
    dragMode.controller.select(cell('B2', 2, 2));

    expect(dragMode.controller.peek()).toBeNull();
    expect(dragMode.onSelectionChange).not.toHaveBeenCalled();
    expect(dragMode.onSwap).not.toHaveBeenCalled();
    expect(dragMode.onReject).not.toHaveBeenCalled();

    const selectMode = makeHarness({ mode: 'select' });
    expect(selectMode.controller.isEnabled(), '选择模式下点击互换必须是关的').toBe(false);
    selectMode.controller.select(cell('A1', 2, 2));
    selectMode.controller.select(cell('C3', 2, 2));
    expect(selectMode.controller.peek()).toBeNull();
    expect(selectMode.counts(), '不应有任何选中变化回调').toEqual([]);
    expect(selectMode.onSwap).not.toHaveBeenCalled();

    // 从 select 切到 click-swap 后立刻开始累积
    selectMode.controller.setMode('click-swap');
    const a = cell('A1', 2, 2);
    selectMode.controller.select(a);
    expect(selectMode.controller.isEnabled()).toBe(true);
    expect(selectMode.controller.peek()).toBe(a);
    expect(selectMode.counts()).toEqual([1]);
  });

  it('setEnabled(false) 闸住点击并清空 pending；切回 drag 也丢掉未完成的选中，之后的点击无效果', () => {
    const gated = makeHarness();
    gated.controller.select(cell('A1', 2, 2));
    gated.controller.setEnabled(false);
    expect(gated.controller.isEnabled()).toBe(false);
    expect(gated.controller.peek()).toBeNull();
    expect(gated.counts()).toEqual([1, 0]);

    gated.controller.select(cell('B2', 2, 2));
    expect(gated.onSwap).not.toHaveBeenCalled();

    gated.controller.setEnabled(true);
    expect(gated.controller.isEnabled()).toBe(true);
    gated.controller.select(cell('A1', 2, 2));
    expect(gated.controller.peek()).not.toBeNull();

    const switched = makeHarness();
    switched.controller.select(cell('A1', 2, 2));
    expect(switched.counts()).toEqual([1]);
    switched.controller.setMode('drag');
    expect(switched.controller.getMode()).toBe('drag');
    expect(switched.controller.peek()).toBeNull();
    expect(switched.counts()).toEqual([1, 0]);

    switched.controller.select(cell('B2', 2, 2));
    expect(switched.counts()).toEqual([1, 0]);
    expect(switched.onSwap).not.toHaveBeenCalled();
  });
});

/* -------------------------------------------------------------------------- */
/* 3. 正常交换                                                                  */
/* -------------------------------------------------------------------------- */

describe('两次点击完成交换', () => {
  it('第一次选中 → 第二次同尺寸交换并清空，可立刻开始下一轮', () => {
    const h = makeHarness();
    const a = cell('A1', 2, 3);
    const b = cell('D4', 2, 3);
    const c = cell('A3', 1, 1);

    h.controller.select(a);
    expect(h.controller.peek()).toBe(a);
    expect(h.onSwap).not.toHaveBeenCalled();
    expect(h.onSelectionChange).toHaveBeenLastCalledWith([a]);

    h.controller.select(b);
    expect(h.onSwap).toHaveBeenCalledTimes(1);
    expect(h.onSwap).toHaveBeenCalledWith(a, b);
    expect(h.onReject).not.toHaveBeenCalled();
    expect(h.controller.peek()).toBeNull();
    expect(h.counts()).toEqual([1, 0]); // 选满两个 → onSwap → 回到 0 个

    // pending 已清空，不残留
    h.controller.select(c);
    expect(h.controller.peek()).toBe(c);
    expect(h.onSwap).toHaveBeenCalledTimes(1);
    expect(h.counts()).toEqual([1, 0, 1]);
  });

  it('跨工作表但尺寸一致可以交换；同一个 a1 但不同 sheetId 视为不同一方', () => {
    const across = makeHarness();
    const s1 = cell('A1', 2, 2, 'sheet-1');
    const s2 = cell('A1', 2, 2, 'sheet-2');
    across.controller.select(s1);
    across.controller.select(s2);
    expect(across.onSwap).toHaveBeenCalledWith(s1, s2);
    expect(across.onReject).not.toHaveBeenCalled();

    const sameA1OtherSheet = makeHarness();
    const b1 = cell('B2', 1, 1, 'sheet-1');
    const b2 = cell('B2', 1, 1, 'sheet-2');
    sameA1OtherSheet.controller.select(b1);
    sameA1OtherSheet.controller.select(b2);
    expect(sameA1OtherSheet.onSwap).toHaveBeenCalledWith(b1, b2);

    // 同一个 cell 连点两次 = 取消；取消之后可以重新选同一方
    const cancelled = makeHarness();
    const a = cell('A1', 2, 2);
    cancelled.controller.select(a);
    cancelled.controller.select(a);
    expect(cancelled.controller.peek()).toBeNull();
    expect(cancelled.counts()).toEqual([1, 0]);
    cancelled.controller.select(a);
    expect(cancelled.controller.peek()).toEqual(a);
    expect(cancelled.counts()).toEqual([1, 0, 1]);
  });
});

/* -------------------------------------------------------------------------- */
/* 4. 同一方二次点击 = 取消选中                                                  */
/* -------------------------------------------------------------------------- */
/* 身份比较的是身份字段（kind/sheetId/a1 或 itemId），不是对象引用；a1 大小写与    */
/* 首尾空白不敏感，label 与尺寸的差异不影响"是不是同一方"。                       */

describe('同一方二次点击 → 取消选中', () => {
  it('结构相同但不同对象、a1 大小写/空白不同都算同一方，不触发 onSwap/onReject', () => {
    const byFields = makeHarness();
    byFields.controller.select(cell('A1', 2, 2, 'sheet-1', '第一处'));
    byFields.controller.select(cell('A1', 2, 2, 'sheet-1', '换个 label 也一样'));
    expect(byFields.controller.peek()).toBeNull();
    expect(byFields.counts()).toEqual([1, 0]);
    expect(byFields.onSwap).not.toHaveBeenCalled();
    expect(byFields.onReject).not.toHaveBeenCalled();

    const caseInsensitive = makeHarness();
    caseInsensitive.controller.select(cell('B2:D5', 4, 4));
    caseInsensitive.controller.select(cell(' b2:d5 ', 4, 4));
    expect(caseInsensitive.controller.peek()).toBeNull();
    expect(caseInsensitive.onSwap).not.toHaveBeenCalled();
  });

  it('同一个 itemId 的 workspace 侧（尺寸被改过）按同一方处理，不触发交换', () => {
    const h = makeHarness();
    const w = workspace('w-1', 2, 2);

    h.controller.select(w);
    h.controller.select(w);
    expect(h.controller.peek()).toBeNull();
    expect(h.counts()).toEqual([1, 0]);
    expect(h.onSwap).not.toHaveBeenCalled();

    const resized = makeHarness();
    resized.controller.select(workspace('w-1', 2, 2));
    resized.controller.select(workspace('w-1', 3, 3));
    expect(resized.controller.peek()).toBeNull();
    expect(resized.onSwap).not.toHaveBeenCalled();
    expect(resized.onReject).not.toHaveBeenCalled();
  });
});

/* -------------------------------------------------------------------------- */
/* 5. 尺寸不一致 → onReject 且 pending 被替换                                    */
/* -------------------------------------------------------------------------- */

describe('尺寸不一致 → onReject 且 pending 换成新的这一方', () => {
  it('行或列任一不同都拒绝：onReject(reason, pending, incoming)，peek 变成 incoming', () => {
    const h = makeHarness();
    const a = cell('A1', 2, 2);
    const b = cell('D4', 3, 5);

    h.controller.select(a);
    h.controller.select(b);

    expect(h.onReject).toHaveBeenCalledTimes(1);
    expect(h.onSwap).not.toHaveBeenCalled();

    const [reason, rejected, incoming] = h.onReject.mock.calls[0] as [string, SwapSide, SwapSide];
    expect(reason).toContain('尺寸不一致');
    expect(reason).toContain('2×2');
    expect(reason).toContain('3×5');
    // 报的是"被顶掉的那一方"与"新来的这一方"，顺序与接口一致
    expect(rejected).toBe(a);
    expect(incoming).toBe(b);
    // 关键行为：用户改主意时不必先取消 —— pending 已经被替换成新的这一方
    expect(h.controller.peek()).toBe(b);

    // 行相同列不同 / 列相同行不同都算不一致
    const byCol = makeHarness();
    byCol.controller.select(cell('A1', 2, 3));
    byCol.controller.select(cell('D4', 2, 4));
    expect(byCol.onReject).toHaveBeenCalledTimes(1);
    expect(byCol.onSwap).not.toHaveBeenCalled();

    const byRow = makeHarness();
    byRow.controller.select(cell('A1', 2, 3));
    byRow.controller.select(cell('D4', 3, 3));
    expect(byRow.onReject).toHaveBeenCalledTimes(1);
    expect(byRow.onSwap).not.toHaveBeenCalled();
  });

  it('被替换后可继续成交（含 1→1 的换人通知）；连续被拒绝两次都各报一次', () => {
    const h = makeHarness();
    const a = cell('A1', 2, 2);
    const b = cell('D4', 3, 5);
    const c = cell('H8', 3, 5);

    h.controller.select(a);
    h.controller.select(b); // 拒绝，pending 变成 b
    expect(h.onReject).toHaveBeenCalledTimes(1);
    // 尺寸不一致时也发 onSelectionChange（1 → 1 的"换人"）
    expect(h.counts()).toEqual([1, 1]);
    expect(h.labels()).toEqual([['A1'], ['D4']]);

    h.controller.select(c); // b 与 c 同尺寸 → 交换
    expect(h.onSwap).toHaveBeenCalledTimes(1);
    expect(h.onSwap).toHaveBeenCalledWith(b, c);
    expect(h.controller.peek()).toBeNull();
    expect(h.counts()).toEqual([1, 1, 0]);

    // 连续被拒绝：pending 始终是最后一个
    const twice = makeHarness();
    twice.controller.select(cell('A1', 1, 1));
    twice.controller.select(cell('B2', 2, 2));
    twice.controller.select(cell('C3', 3, 3));
    expect(twice.onReject).toHaveBeenCalledTimes(2);
    expect(twice.onSwap).not.toHaveBeenCalled();
    expect(twice.controller.peek()?.label).toBe('C3');
    expect(twice.counts()).toEqual([1, 1, 1]);
  });
});

/* -------------------------------------------------------------------------- */
/* 6. 工作区 ↔ 单元格互交                                                        */
/* -------------------------------------------------------------------------- */

describe('工作区条目 ↔ 单元格可以互相交换', () => {
  it('两个方向（尺寸一致）都能交换；跨 kind 尺寸不一致 → onReject（kind 不参与尺寸比较）', () => {
    const toWorkspace = makeHarness();
    const a = cell('A1', 2, 3);
    const w = workspace('w-1', 2, 3, '成绩单');
    toWorkspace.controller.select(a);
    toWorkspace.controller.select(w);
    expect(toWorkspace.onSwap).toHaveBeenCalledWith(a, w);
    expect(toWorkspace.onReject).not.toHaveBeenCalled();

    const toCell = makeHarness();
    const w2 = workspace('w-1', 1, 4);
    const a2 = cell('A1', 1, 4);
    toCell.controller.select(w2);
    toCell.controller.select(a2);
    expect(toCell.onSwap).toHaveBeenCalledWith(w2, a2);

    const rejected = makeHarness();
    const rw = workspace('w-1', 2, 2);
    const ra = cell('A1', 3, 2);
    rejected.controller.select(rw);
    rejected.controller.select(ra);
    expect(rejected.onSwap).not.toHaveBeenCalled();
    expect(rejected.onReject).toHaveBeenCalledTimes(1);
    const [reason, rejectedSide, incoming] = rejected.onReject.mock.calls[0] as [string, SwapSide, SwapSide];
    expect(reason).toContain('尺寸不一致');
    expect(rejectedSide).toBe(rw);
    expect(incoming).toBe(ra);
    expect(rejected.controller.peek()).toBe(ra);

    // 同一个 itemId 的 workspace 侧（身份相同、尺寸被改过）按同一方处理，不触发交换
    const sameItem = makeHarness();
    const sw = workspace('w-2', 2, 2);
    sameItem.controller.select(sw);
    sameItem.controller.select(sw);
    expect(sameItem.controller.peek()).toBeNull();
    expect(sameItem.onSwap).not.toHaveBeenCalled();
    expect(sameItem.onReject).not.toHaveBeenCalled();

    const resized = makeHarness();
    resized.controller.select(workspace('w-3', 2, 2));
    resized.controller.select(workspace('w-3', 3, 3));
    expect(resized.controller.peek()).toBeNull();
    expect(resized.onSwap).not.toHaveBeenCalled();
  });

  it('cell 与 workspace 缺省字段互补，不会被误判成"同一方"', () => {
    const sparse = makeHarness();
    // 两者 a1/sheetId/itemId 全为 undefined：kind 不同 → 不是同一方
    const cellSide: SwapSide = { kind: 'cell', rows: 1, cols: 1, label: '空 cell' };
    const workspaceSide: SwapSide = { kind: 'workspace', rows: 1, cols: 1, label: '空 workspace' };

    sparse.controller.select(cellSide);
    sparse.controller.select(workspaceSide);

    expect(sparse.onSwap).toHaveBeenCalledWith(cellSide, workspaceSide);
  });
});

/* -------------------------------------------------------------------------- */
/* 7. onSelectionChange 序列                                                    */
/* -------------------------------------------------------------------------- */

describe('onSelectionChange 序列', () => {
  it('完整序列是 0 → 1 → 0，交换通知先于 onSwap；取消 / clear 也通知 0 个，且数组不被后续操作改动', () => {
    const h = makeHarness();
    const a = cell('A1', 2, 2);
    const b = cell('C3', 2, 2);

    h.controller.select(a);
    h.controller.select(b);

    // 初始不通知；第一次 select 通知 [a]；交换时先通知 [] 再 onSwap
    expect(h.counts()).toEqual([1, 0]);

    const order: string[] = [];
    const controller = createClickSwapController({
      onSelectionChange: vi.fn(() => order.push('selection')),
      onSwap: vi.fn(() => order.push('swap')),
    });
    controller.setMode('click-swap');
    controller.select(a);
    controller.select(b);
    expect(order).toEqual(['selection', 'selection', 'swap']);

    // 取消选中（同一方二次点击）也发一次 0 个的通知
    const cancelled = makeHarness();
    cancelled.controller.select(a);
    cancelled.controller.select(a);
    expect(cancelled.onSelectionChange).toHaveBeenCalledTimes(2);
    expect(cancelled.onSelectionChange).toHaveBeenLastCalledWith([]);

    cancelled.controller.select(a);
    const first = cancelled.onSelectionChange.mock.calls[0][0] as SwapSide[];
    cancelled.controller.clear();
    expect(cancelled.onSelectionChange).toHaveBeenLastCalledWith([]);
    expect(cancelled.controller.peek()).toBeNull();

    cancelled.controller.clear(); // 已经是空的 → 静默
    expect(cancelled.onSelectionChange).toHaveBeenCalledTimes(4);

    // 每次都是新数组：第一次通知的数组不被后续 select 改动
    cancelled.controller.select(b);
    expect(first).toEqual([a]);
    expect(first).toHaveLength(1);
  });
});

/* -------------------------------------------------------------------------- */
/* 8. dispose 与边界                                                            */
/* -------------------------------------------------------------------------- */

describe('dispose 与边界', () => {
  it('dispose 后不再触发任何回调、peek 归空、isEnabled 为 false，但 getMode() 不变；重复 dispose 是 no-op', () => {
    const h = makeHarness();
    const a = cell('A1', 2, 2);
    const b = cell('C3', 2, 2);

    h.controller.select(a);
    expect(h.counts()).toEqual([1]);

    h.controller.dispose();

    expect(h.controller.peek()).toBeNull();
    expect(h.controller.isEnabled()).toBe(false);
    expect(h.controller.getMode()).toBe('click-swap'); // 模式是配置，不是选中状态

    h.controller.select(b); // 忽略
    h.controller.select(a); // 忽略
    h.controller.clear();
    h.controller.setMode('drag');
    h.controller.setMode('click-swap');
    h.controller.setEnabled(false);
    h.controller.setEnabled(true);

    expect(h.onSelectionChange).toHaveBeenCalledTimes(1);
    expect(h.onSwap).not.toHaveBeenCalled();
    expect(h.onReject).not.toHaveBeenCalled();
    expect(h.controller.peek()).toBeNull();

    expect(() => h.controller.dispose()).not.toThrow();
    expect(h.controller.isEnabled()).toBe(false);
  });

  it('不传任何回调也不会抛异常；零尺寸/负数尺寸只要两边一致就允许交换；peek() 返回传入的对象本身', () => {
    const bare = createClickSwapController({});
    bare.setMode('click-swap');
    expect(() => {
      bare.select(cell('A1', 2, 2));
      bare.select(cell('A1', 2, 2));
      bare.select(cell('B2', 3, 3));
      bare.clear();
      bare.dispose();
    }).not.toThrow();

    const h = makeHarness();
    const a = cell('A1', 0, 0);
    const b = cell('B2', 0, 0);
    h.controller.select(a);
    expect(h.controller.peek()).toBe(a); // 身份判断用得上
    h.controller.select(b);
    expect(h.onSwap).toHaveBeenCalledWith(a, b);
  });
});
