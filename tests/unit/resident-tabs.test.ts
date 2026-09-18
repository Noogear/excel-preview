/**
 * 多标签"常驻窗口"决策单测（纯函数）。
 *
 * 这些函数决定"哪个标签的 Univer 工作簿该被冷存（dispose）"以及"切过去要不要重建"。
 * 决策错一个方向就会出事：
 *  - 淘汰了活动标签 → 用户正看着的表突然空白；
 *  - 该淘汰却不淘汰 → 内存继续 O(N) 涨；
 *  - 该重建却不重建 → 切回去是空工作簿（数据看起来丢了）。
 */
import { describe, expect, it } from 'vitest';

import {
  DEFAULT_RESIDENT_TAB_LIMIT,
  dropTab,
  markTabCold,
  markTabUsed,
  needsBuild,
  planEvictions,
  type TabRuntime,
} from '../../src/shell/resident-tabs';

const tab = (id: string, built: boolean, lastUsedAt: number): TabRuntime => ({ id, built, lastUsedAt });

describe('planEvictions', () => {
  it('没超过窗口时谁都不淘汰', () => {
    const tabs = [tab('a', true, 1), tab('b', true, 2), tab('c', true, 3)];
    expect(planEvictions(tabs, 'c', 3)).toEqual([]);
  });

  it('超过窗口时按"最久未用"淘汰非活动标签，绝不动活动标签', () => {
    const tabs = [tab('a', true, 10), tab('b', true, 30), tab('c', true, 20), tab('d', true, 40)];
    // 窗口 3、实体化 4 → 淘汰 1 个：最久未用的非活动标签是 a（10）
    expect(planEvictions(tabs, 'd', 3)).toEqual(['a']);
    // 即使 a 是最久未用的，只要它是活动标签就不能动 → 在剩下的非活动标签里挑最久未用的（c=20）
    expect(planEvictions(tabs, 'a', 3)).toEqual(['c']);
  });

  it('一次可以淘汰多个（开 5 个、窗口 2）', () => {
    const tabs = [tab('a', true, 100), tab('b', true, 200), tab('c', true, 300), tab('d', true, 400), tab('e', true, 500)];
    expect(planEvictions(tabs, 'e', 2)).toEqual(['a', 'b', 'c']);
  });

  it('已经冷着的标签不会被重复淘汰', () => {
    const tabs = [tab('a', true, 1), tab('b', false, 2), tab('c', true, 3), tab('d', true, 4)];
    expect(planEvictions(tabs, 'd', 2)).toEqual(['a']);
  });

  it('lastUsedAt 相同时按 id 稳定排序（同状态决策可重复）', () => {
    const tabs = [tab('b', true, 5), tab('a', true, 5), tab('c', true, 5)];
    expect(planEvictions(tabs, 'c', 2)).toEqual(['a']);
  });

  it('窗口下限是 1（配置写 0/负数也不会把活动标签也淘汰掉）', () => {
    const tabs = [tab('a', true, 1), tab('b', true, 2)];
    expect(planEvictions(tabs, 'b', 0)).toEqual(['a']);
    expect(planEvictions(tabs, 'a', 0)).toEqual(['b']);
  });

  it('默认窗口是 3', () => {
    expect(DEFAULT_RESIDENT_TAB_LIMIT).toBe(3);
    const tabs = [tab('a', true, 1), tab('b', true, 2), tab('c', true, 3), tab('d', true, 4)];
    expect(planEvictions(tabs, 'd')).toEqual(['a']);
  });
});

describe('needsBuild', () => {
  it('已实体化的标签不需要重建；冷标签需要', () => {
    const tabs = [tab('a', true, 1), tab('b', false, 2)];
    expect(needsBuild(tabs, 'a')).toBe(false);
    expect(needsBuild(tabs, 'b')).toBe(true);
  });

  it('记账里没有的标签按"需要重建"处理（宁可多解析一次，也不要给用户一个空表）', () => {
    expect(needsBuild([tab('a', true, 1)], 'zzz')).toBe(true);
  });
});

describe('记账更新', () => {
  it('markTabUsed 记录使用时间并标记为已实体化（新标签会被追加）', () => {
    const before = [tab('a', true, 1)];
    const after = markTabUsed(before, 'a', 99);
    expect(after[0]).toEqual({ id: 'a', built: true, lastUsedAt: 99 });
    expect(before[0].lastUsedAt, '不改原数组').toBe(1);
    expect(markTabUsed(after, 'b', 100)).toHaveLength(2);
  });

  it('markTabCold 只把 built 置 false，保留 lastUsedAt 供后续淘汰参考', () => {
    const after = markTabCold([tab('a', true, 7)], 'a');
    expect(after[0]).toEqual({ id: 'a', built: false, lastUsedAt: 7 });
  });

  it('dropTab 关标签时移除记账（否则记账会随着开关标签无限增长）', () => {
    const after = dropTab([tab('a', true, 1), tab('b', false, 2)], 'a');
    expect(after.map((entry) => entry.id)).toEqual(['b']);
  });
});
