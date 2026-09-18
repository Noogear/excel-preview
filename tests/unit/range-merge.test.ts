/**
 * 选区与合并单元格对齐（纯函数）。
 *
 * 用户实测："选择多个后右键剪切到工作区，剪切的内容并非是我选中的那部分"。
 * 课程表里满是合并块：拖选 B4:F8 时模型选区会被扩成 B4:F11（第 9~11 行是竖向合并），
 * 于是剪走整块。这里把补齐规则钉死，并保证补齐是"最小且收敛"的。
 */
import { describe, expect, it } from 'vitest';

import { expandRectToMerges, rectSize, rectsIntersect } from '../../src/interaction/range-merge';

const rect = (startRow: number, startColumn: number, endRow: number, endColumn: number) => ({
  startRow,
  startColumn,
  endRow,
  endColumn,
});

describe('rectsIntersect', () => {
  it('相交 / 相邻 / 分离', () => {
    expect(rectsIntersect(rect(0, 0, 1, 1), rect(1, 1, 2, 2))).toBe(true);
    expect(rectsIntersect(rect(0, 0, 1, 1), rect(2, 2, 3, 3))).toBe(false);
    expect(rectsIntersect(rect(0, 0, 0, 0), rect(0, 0, 5, 5))).toBe(true);
  });
});

describe('expandRectToMerges', () => {
  it('没有相交的合并块 → 原样返回、expanded=false', () => {
    const result = expandRectToMerges(rect(1, 1, 3, 3), [rect(10, 10, 12, 12)]);
    expect(result).toEqual({ rect: rect(1, 1, 3, 3), expanded: false });
  });

  it('压到竖向合并块 → 扩到整块（用户那个例子的复刻：B4:F8 被第 9~11 行的合并撑到 B4:F11）', () => {
    // B4:F8 = 行 3..7、列 1..5；C 列上有一个 行 5..10 的竖向合并（0-based）
    const merges = [rect(5, 2, 10, 2)];
    const result = expandRectToMerges(rect(3, 1, 7, 5), merges);
    expect(result.expanded).toBe(true);
    expect(result.rect).toEqual(rect(3, 1, 10, 5));
  });

  it('阶梯状合并：扩一次又压到新的合并块 → 反复扩直到收敛', () => {
    const merges = [rect(5, 2, 10, 2), rect(10, 4, 15, 6)];
    const result = expandRectToMerges(rect(3, 1, 7, 5), merges);
    // 第一次扩到行 10，第二次被第二个合并块扩到列 6、行 15
    expect(result.rect).toEqual(rect(3, 1, 15, 6));
    expect(result.expanded).toBe(true);
  });

  it('选区整个包住合并块 → 不需要扩', () => {
    const result = expandRectToMerges(rect(0, 0, 20, 20), [rect(5, 5, 8, 8)]);
    expect(result.expanded).toBe(false);
    expect(result.rect).toEqual(rect(0, 0, 20, 20));
  });

  it('已经对齐（边界正好贴着合并块）→ 不扩（避免无意义地变大）', () => {
    // 选区行 0..4，合并块行 5..8：不相交，保持原样
    const result = expandRectToMerges(rect(0, 0, 4, 4), [rect(5, 0, 8, 4)]);
    expect(result.expanded).toBe(false);
    expect(result.rect).toEqual(rect(0, 0, 4, 4));
  });

  it('只有部分重合的合并块也会把范围补齐到不漏', () => {
    // 选区只覆盖合并块的右下角一点
    const result = expandRectToMerges(rect(4, 4, 6, 6), [rect(2, 2, 5, 5)]);
    expect(result.rect).toEqual(rect(2, 2, 6, 6));
  });

  it('合并块列表为空 → 原样返回', () => {
    expect(expandRectToMerges(rect(1, 1, 2, 2), []).expanded).toBe(false);
  });
});

describe('rectSize', () => {
  it('行列尺寸', () => {
    expect(rectSize(rect(3, 1, 7, 5))).toEqual({ rows: 5, cols: 5 });
    expect(rectSize(rect(0, 0, 0, 0))).toEqual({ rows: 1, cols: 1 });
  });
});
