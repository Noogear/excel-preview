/**
 * 历史记录模型单测（node 环境，不需要 jsdom）。
 *
 * 覆盖历史面板真正依赖的语义（同一主题下的整块场景合并为一个 it，断言一条未删）：
 *  - planJump：撤销/重做步数换算（两个方向 + 原地 + 两端边界 + 互斥 / 距离不变量）、
 *    脏入参（负数 / 小数 / NaN / ±Infinity）夹紧且结果恒为非负整数；
 *  - canJump：可用次数是否够走完计划（恰好够 / 富余 / 差一次 / 方向不通 / 原地 + 脏可用计数夹紧）、
 *    与 planJump 的串联；
 *  - appendEntry：追加顺序与纯函数（不改 / 不炸冻结入参）、同 id 原地更新去重、
 *    limit 截断保留最新且去重先于截断；
 *  - formatTime：HH:MM:SS 补零、跨小时 / 跨日 / 1970 前时间戳的本地时区口径、非法时间戳占位符；
 *  - describeKind：7 个 kind 的中文名/tone 映射 + 未知值与原型属性兜底 + 返回新对象；
 *  - clampIndex：非负整数夹紧；jumpBlockReason：可达为 null、不可达的中文文案与 planJump 组合。
 */
import { describe, expect, it } from 'vitest';

import {
  HISTORY_KINDS,
  appendEntry,
  canJump,
  clampIndex,
  describeKind,
  formatTime,
  jumpBlockReason,
  planJump,
  type HistoryEntry,
  type HistoryKind,
  type KindTone,
  type UndoRedoAvailability,
} from '../../src/shell/history-model';

/* ---------------------------------------------------------------- 工具 */

function entry(id: string, over: Partial<HistoryEntry> = {}): HistoryEntry {
  return { id, label: `步骤 ${id}`, kind: 'edit', at: 0, ...over };
}

function ids(entries: HistoryEntry[]): string[] {
  return entries.map((item) => item.id);
}

function labels(entries: HistoryEntry[]): string[] {
  return entries.map((item) => item.label);
}

/** 用本地时间分量构造时间戳：期望值同样用本地 getter 拼，因此与运行时区无关。 */
function localAt(hour: number, minute: number, second: number, day = 2): number {
  return new Date(2024, 0, day, hour, minute, second, 0).getTime();
}

const NO_AVAIL: UndoRedoAvailability = { undos: 0, redos: 0 };

/* ---------------------------------------------------------------- planJump */

describe('planJump - 步数换算', () => {
  it('后退只撤销 / 前进只重做 / 原地与两端边界都不动，且两方向互斥、步数之差等于下标距离', () => {
    expect(planJump(3, 1)).toEqual({ undo: 2, redo: 0 });
    expect(planJump(12, 4)).toEqual({ undo: 8, redo: 0 });
    expect(planJump(1, 3)).toEqual({ undo: 0, redo: 2 });
    expect(planJump(4, 12)).toEqual({ undo: 0, redo: 8 });
    // 原地：含"最初状态 -> 最初状态"
    expect(planJump(2, 2)).toEqual({ undo: 0, redo: 0 });
    expect(planJump(0, 0)).toEqual({ undo: 0, redo: 0 });
    // 从最初状态出发 / 退回最初状态
    expect(planJump(0, 5)).toEqual({ undo: 0, redo: 5 });
    expect(planJump(5, 0)).toEqual({ undo: 5, redo: 0 });

    // 任意一对入参都满足：两个方向互斥、步数之差等于下标距离
    const pairs: Array<[number, number]> = [
      [0, 0], [0, 7], [7, 0], [3, 3], [9, 2], [2, 9], [6, 2], [2, 6], [5, 5],
    ];
    for (const [from, to] of pairs) {
      const plan = planJump(from, to);
      expect(plan.undo === 0 || plan.redo === 0).toBe(true);
      expect(Math.abs(plan.undo - plan.redo)).toBe(Math.abs(from - to));
    }
  });
});

describe('planJump - 脏入参夹紧', () => {
  it('负数 / 小数 / NaN / ±Infinity / 大下标都夹成非负整数步数', () => {
    expect(planJump(-3, 2)).toEqual({ undo: 0, redo: 2 });
    expect(planJump(4, -1)).toEqual({ undo: 4, redo: 0 });
    expect(planJump(-2, -5)).toEqual({ undo: 0, redo: 0 });
    // 小数向下取整（两个方向）
    expect(planJump(2.9, 1.2)).toEqual({ undo: 1, redo: 0 });
    expect(planJump(0.4, 3.7)).toEqual({ undo: 0, redo: 3 });
    // 绝不返回 Infinity 次撤销
    expect(planJump(NaN, 2)).toEqual({ undo: 0, redo: 2 });
    expect(planJump(2, NaN)).toEqual({ undo: 2, redo: 0 });
    expect(planJump(NaN, NaN)).toEqual({ undo: 0, redo: 0 });
    expect(planJump(Infinity, 3)).toEqual({ undo: 0, redo: 3 });
    expect(planJump(3, Infinity)).toEqual({ undo: 3, redo: 0 });
    expect(planJump(-Infinity, -Infinity)).toEqual({ undo: 0, redo: 0 });
    expect(planJump(1_000_000, 999_997)).toEqual({ undo: 3, redo: 0 });

    // 任意脏入参下结果恒为非负整数
    const dirty: Array<[number, number]> = [
      [-1, -1], [-100, 3], [3, -100], [NaN, -Infinity], [-0.5, -0.5],
    ];
    for (const [from, to] of dirty) {
      const plan = planJump(from, to);
      expect(Number.isInteger(plan.undo)).toBe(true);
      expect(Number.isInteger(plan.redo)).toBe(true);
      expect(plan.undo).toBeGreaterThanOrEqual(0);
      expect(plan.redo).toBeGreaterThanOrEqual(0);
    }
  });
});

/* ---------------------------------------------------------------- canJump */

describe('canJump - 可用次数', () => {
  it('恰好够 / 富余 → true；差一次 / 一次都没有 / 方向不通 → false；脏可用计数先夹紧再比', () => {
    expect(canJump({ undo: 2, redo: 0 }, { undos: 2, redos: 0 })).toBe(true);
    expect(canJump({ undo: 0, redo: 3 }, { undos: 0, redos: 3 })).toBe(true);
    expect(canJump({ undo: 1, redo: 0 }, { undos: 99, redos: 99 })).toBe(true);
    // 原地跳不需要任何可用次数
    expect(canJump({ undo: 0, redo: 0 }, NO_AVAIL)).toBe(true);

    expect(canJump({ undo: 2, redo: 0 }, { undos: 1, redos: 0 })).toBe(false);
    expect(canJump({ undo: 0, redo: 3 }, { undos: 0, redos: 2 })).toBe(false);
    expect(canJump({ undo: 1, redo: 0 }, NO_AVAIL)).toBe(false);
    expect(canJump({ undo: 0, redo: 1 }, NO_AVAIL)).toBe(false);
    // 方向不通：只有 redo 可用时撤销计划依然不可行（反之亦然）
    expect(canJump({ undo: 2, redo: 0 }, { undos: 0, redos: 10 })).toBe(false);
    expect(canJump({ undo: 0, redo: 2 }, { undos: 10, redos: 0 })).toBe(false);

    // 脏可用计数（负数 / NaN）先夹紧再比，计划里的脏计数也先夹紧
    expect(canJump({ undo: 1, redo: 0 }, { undos: -1, redos: 0 })).toBe(false);
    expect(canJump({ undo: 0, redo: 0 }, { undos: NaN, redos: NaN })).toBe(true);
    expect(canJump({ undo: 0, redo: 5 }, { undos: 0, redos: -3 })).toBe(false);
    // 1.9 次撤销只需 1 次可用
    expect(canJump({ undo: 1.9, redo: 0 }, { undos: 1, redos: 0 })).toBe(true);
  });

  it('与 planJump 串起来：只剩 1 次重做时只有第 1 步可达', () => {
    const available: UndoRedoAvailability = { undos: 0, redos: 1 };
    expect(canJump(planJump(0, 1), available)).toBe(true);
    expect(canJump(planJump(0, 2), available)).toBe(false);
    expect(canJump(planJump(0, 3), available)).toBe(false);
  });
});

/* ---------------------------------------------------------------- appendEntry */

describe('appendEntry - 追加 / 去重 / limit', () => {
  it('新 id 追加到末尾并保持顺序，且总是返回新数组、不改入参', () => {
    const first = appendEntry([], entry('a', { label: '互换 A3 ⇄ A4', kind: 'swap' }));
    expect(first).toHaveLength(1);
    expect(first[0]?.label).toBe('互换 A3 ⇄ A4');
    expect(first[0]?.kind).toBe('swap');

    let list: HistoryEntry[] = [];
    list = appendEntry(list, entry('a'));
    list = appendEntry(list, entry('b'));
    list = appendEntry(list, entry('c'));
    expect(ids(list)).toEqual(['a', 'b', 'c']);

    const source = [entry('a'), entry('b')];
    const result = appendEntry(source, entry('c'));
    expect(result).not.toBe(source);
    expect(source).toHaveLength(2);
    expect(ids(source)).toEqual(['a', 'b']);
    // 冻结的入参也不会被写坏（无 push / splice 就地修改）
    const frozen = Object.freeze([entry('a')]) as unknown as HistoryEntry[];
    expect(() => appendEntry(frozen, entry('b'))).not.toThrow();
    expect(frozen).toHaveLength(1);
  });

  it('相同 id 视为原地更新：位置不变、字段被改写、旧对象不被修改', () => {
    const old = entry('a', { label: '旧标签', kind: 'edit', at: 1 });
    const source = [old, entry('b'), entry('c')];
    const result = appendEntry(source, entry('a', { label: '新标签', kind: 'swap', at: 2 }));

    expect(ids(result)).toEqual(['a', 'b', 'c']); // 不会被挪到末尾
    expect(labels(result)).toEqual(['新标签', '步骤 b', '步骤 c']);
    expect(result[0]?.kind).toBe('swap');
    expect(result[0]?.at).toBe(2);
    // 旧引用仍是旧值
    expect(old.label).toBe('旧标签');
    expect(source[0]).toBe(old);
    expect(result[0]).not.toBe(old);
  });

  it('limit 只保留最新的 limit 条（丢弃最旧的头部），未给出/非法时不截断；去重发生在截断之前', () => {
    const source = [entry('a'), entry('b'), entry('c')];
    expect(ids(appendEntry(source, entry('d'), 3))).toEqual(['b', 'c', 'd']);
    expect(ids(appendEntry([entry('a'), entry('b')], entry('c'), 1))).toEqual(['c']);
    // 小数向下取整（2.9 视作 2）
    expect(ids(appendEntry(source, entry('c'), 2.9))).toEqual(['b', 'c']);
    // 等于或大于长度 -> 不截断
    expect(ids(appendEntry([entry('a'), entry('b')], entry('c'), 3))).toEqual(['a', 'b', 'c']);
    expect(ids(appendEntry([entry('a')], entry('b'), 50))).toEqual(['a', 'b']);
    // 0 / 负数 / NaN / undefined 视为不限（不返回空数组）
    expect(appendEntry(source, entry('c'), 0)).toHaveLength(3);
    expect(appendEntry(source, entry('c'), -5)).toHaveLength(3);
    expect(appendEntry(source, entry('c'), NaN)).toHaveLength(3);
    // 不传 limit 时也不截断（默认不限）
    const long = Array.from({ length: 40 }, (_, i) => entry(`e${i}`));
    expect(appendEntry(long, entry('last'))).toHaveLength(41);

    // 去重先于截断：更新末尾条目后 limit 仍保留它
    const updated = appendEntry(source, entry('c', { label: '新 C' }), 2);
    expect(ids(updated)).toEqual(['b', 'c']);
    expect(updated[1]?.label).toBe('新 C');
    // 被更新的条目若落在最旧一侧，会被截断掉
    const dropped = appendEntry([entry('a', { label: '旧 A' }), entry('b'), entry('c')], entry('a', { label: '新 A' }), 2);
    expect(ids(dropped)).toEqual(['b', 'c']);
    expect(labels(dropped)).not.toContain('新 A');
  });
});

/* ---------------------------------------------------------------- formatTime */

describe('formatTime - HH:MM:SS（本地时区）', () => {
  it('两位补零，且任意小时都恒为 8 位 HH:MM:SS', () => {
    expect(formatTime(localAt(9, 5, 7))).toBe('09:05:07');
    expect(formatTime(localAt(3, 4, 5))).toBe('03:04:05');
    expect(formatTime(localAt(0, 0, 0))).toBe('00:00:00');
    expect(formatTime(localAt(12, 0, 0))).toBe('12:00:00');
    expect(formatTime(localAt(10, 20, 30))).toBe('10:20:30');
    for (let hour = 0; hour < 24; hour += 1) {
      const text = formatTime(localAt(hour, hour, hour));
      expect(text).toMatch(/^\d{2}:\d{2}:\d{2}$/);
      expect(text).toHaveLength(8);
    }
  });

  it('本地时区口径：跨小时、跨日与 1970 之前的时间戳', () => {
    expect(formatTime(localAt(9, 59, 59))).toBe('09:59:59');
    expect(formatTime(localAt(10, 0, 0))).toBe('10:00:00');
    expect(formatTime(localAt(23, 59, 59))).toBe('23:59:59');
    expect(formatTime(localAt(23, 59, 59) + 1000)).toBe('00:00:00'); // 下一秒是次日零点
    expect(formatTime(new Date(1969, 11, 31, 23, 30, 5).getTime())).toBe('23:30:05');

    const samples = [0, localAt(1, 2, 3), localAt(23, 59, 59), Date.UTC(2024, 0, 2, 9, 5, 7)];
    for (const at of samples) {
      const date = new Date(at);
      const expected = [date.getHours(), date.getMinutes(), date.getSeconds()]
        .map((n) => String(n).padStart(2, '0'))
        .join(':');
      expect(formatTime(at)).toBe(expected);
    }
  });

  it('NaN / ±Infinity / 超出 Date 范围 → 占位符 --:--:--', () => {
    expect(formatTime(NaN)).toBe('--:--:--');
    expect(formatTime(Infinity)).toBe('--:--:--');
    expect(formatTime(-Infinity)).toBe('--:--:--');
    expect(formatTime(8.64e15 + 1)).toBe('--:--:--');
  });
});

/* ---------------------------------------------------------------- describeKind */

describe('describeKind - 分类名与颜色语义键', () => {
  it('7 个 kind 都有非空中文名与合法 tone、互不重复，未知值 / 原型属性 → 其它，返回新对象不污染后续调用', () => {
    const tones: KindTone[] = ['neutral', 'primary', 'ok', 'warn'];
    expect(HISTORY_KINDS).toHaveLength(7);
    for (const kind of HISTORY_KINDS) {
      const described = describeKind(kind);
      expect(described.label.length).toBeGreaterThan(0);
      expect(tones).toContain(described.tone);
    }
    const names = HISTORY_KINDS.map((kind) => describeKind(kind).label);
    expect(new Set(names).size).toBe(names.length);

    expect(describeKind('edit')).toEqual({ label: '编辑', tone: 'neutral' });
    expect(describeKind('swap')).toEqual({ label: '互换', tone: 'primary' });
    expect(describeKind('import')).toEqual({ label: '导入', tone: 'primary' });
    expect(describeKind('export')).toEqual({ label: '导出', tone: 'ok' });
    expect(describeKind('workspace')).toEqual({ label: '工作区', tone: 'ok' });
    expect(describeKind('mode')).toEqual({ label: '模式', tone: 'warn' });
    expect(describeKind('other')).toEqual({ label: '其它', tone: 'neutral' });

    // 未知值 / 原型属性 → 其它
    const dirty = ['rename', '', 'toString', 'constructor', undefined, null] as unknown as HistoryKind[];
    for (const kind of dirty) {
      expect(describeKind(kind)).toEqual({ label: '其它', tone: 'neutral' });
    }
    // 返回新对象：改动它不会污染后续调用
    const first = describeKind('swap');
    first.label = '被改坏了';
    first.tone = 'warn';
    expect(describeKind('swap')).toEqual({ label: '互换', tone: 'primary' });
  });
});

/* ---------------------------------------------------------------- clampIndex */

describe('clampIndex - 非负整数夹紧', () => {
  it('正数原样（小数向下取整），负数 / NaN / ±Infinity → 0，结果恒为非负整数', () => {
    expect(clampIndex(0)).toBe(0);
    expect(clampIndex(3)).toBe(3);
    expect(clampIndex(3.9)).toBe(3);
    expect(clampIndex(-1)).toBe(0);
    expect(clampIndex(-0.5)).toBe(0);
    expect(clampIndex(NaN)).toBe(0);
    expect(clampIndex(Infinity)).toBe(0);
    expect(clampIndex(-Infinity)).toBe(0);
    for (const value of [7.7, -7.7, 0, 1e9, NaN, Infinity]) {
      const clamped = clampIndex(value);
      expect(Number.isInteger(clamped)).toBe(true);
      expect(clamped).toBeGreaterThanOrEqual(0);
    }
  });
});

/* ---------------------------------------------------------------- jumpBlockReason */

describe('jumpBlockReason - 不可达原因文案', () => {
  it('可达 → null（含原地跳）；不足时分别报撤销 / 重做的需求量与可用量（脏计数先夹紧），与 planJump 一致', () => {
    expect(jumpBlockReason({ undo: 2, redo: 0 }, { undos: 2, redos: 0 })).toBeNull();
    expect(jumpBlockReason({ undo: 0, redo: 0 }, NO_AVAIL)).toBeNull();

    expect(jumpBlockReason({ undo: 0, redo: 3 }, { undos: 9, redos: 1 })).toBe('需要重做 3 次，但只有 1 次可用');
    expect(jumpBlockReason({ undo: 4, redo: 0 }, { undos: 2, redos: 9 })).toBe('需要撤销 4 次，但只有 2 次可用');
    expect(jumpBlockReason({ undo: 1, redo: 0 }, { undos: -1, redos: 0 })).toBe('需要撤销 1 次，但只有 0 次可用');

    // 与 planJump 串起来：只剩 1 次重做时第 1 步可达、第 3 步给出文案
    const available: UndoRedoAvailability = { undos: 0, redos: 1 };
    expect(jumpBlockReason(planJump(0, 3), available)).toBe('需要重做 3 次，但只有 1 次可用');
    expect(jumpBlockReason(planJump(0, 1), available)).toBeNull();
  });
});
