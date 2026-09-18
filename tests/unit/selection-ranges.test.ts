/**
 * 不连续多区域选区（Ctrl+点选）的纯函数：规范化、去重、状态栏文案、多块导入解析。
 *
 * 用户要求："请实现不连续多区域（Ctrl+点选）"。
 * Univer 的选区模型本来就支持多块（`getCurrentSelections()` 是数组、Facade 有
 * `getActiveRangeList()`），缺的是**我们自己那些只读 `getActiveRange()` 的地方**。
 * 这个模块把"多块"这件事变成可测的算术与文案，App 侧只负责把结果用起来。
 */
import { describe, expect, it } from 'vitest';

import {
  boundingRect,
  colLetter,
  containsRect,
  describeSingleBlockOnly,
  normalizeRanges,
  parseRangeList,
  rectArea,
  rectToA1,
  rectsEqual,
  rectsIntersect,
  rectsToA1,
  sortRects,
  summarizeSelection,
  totalCells,
} from '../../src/interaction/selection-ranges';

const rect = (startRow: number, startCol: number, endRow: number, endCol: number) => ({ startRow, startCol, endRow, endCol });

describe('坐标与面积', () => {
  it('colLetter 覆盖 A/Z/AA/AB', () => {
    expect(colLetter(0)).toBe('A');
    expect(colLetter(25)).toBe('Z');
    expect(colLetter(26)).toBe('AA');
    expect(colLetter(27)).toBe('AB');
  });

  it('rectToA1：单格不带冒号，矩形带冒号', () => {
    expect(rectToA1(rect(1, 1, 1, 1))).toBe('B2');
    expect(rectToA1(rect(1, 1, 4, 3))).toBe('B2:D5');
  });

  it('rectArea / totalCells', () => {
    expect(rectArea(rect(0, 0, 2, 3))).toBe(12);
    expect(rectArea(rect(5, 5, 1, 1)), '反着的矩形算 0 而不是负数').toBe(0);
    expect(totalCells([rect(0, 0, 0, 0), rect(1, 1, 2, 2)])).toBe(1 + 4);
  });

  it('rectsIntersect / rectsEqual / containsRect', () => {
    expect(rectsIntersect(rect(0, 0, 2, 2), rect(2, 2, 4, 4))).toBe(true);
    expect(rectsIntersect(rect(0, 0, 1, 1), rect(2, 2, 4, 4))).toBe(false);
    expect(rectsEqual(rect(0, 0, 1, 1), rect(0, 0, 1, 1))).toBe(true);
    expect(containsRect(rect(0, 0, 4, 4), rect(1, 1, 2, 2))).toBe(true);
    expect(containsRect(rect(0, 0, 1, 1), rect(1, 1, 2, 2))).toBe(false);
  });
});

describe('normalizeRanges：多块区域的规范化', () => {
  it('解析失败/空串直接丢掉，不抛异常', () => {
    expect(normalizeRanges(['B2', '', '不是区域', 'D4'])).toEqual([rect(1, 1, 1, 1), rect(3, 3, 3, 3)]);
    expect(normalizeRanges([])).toEqual([]);
  });

  it('完全相同的块去重', () => {
    expect(normalizeRanges(['B2:D5', 'B2:D5'])).toEqual([rect(1, 1, 4, 3)]);
  });

  it('被包含的小块丢掉（Ctrl+点选点在大片里不该产生重复内容）', () => {
    expect(normalizeRanges(['A1:D4', 'B2'])).toEqual([rect(0, 0, 3, 3)]);
    // 反过来：先点小格再框大片，结果一样
    expect(normalizeRanges(['B2', 'A1:D4'])).toEqual([rect(0, 0, 3, 3)]);
  });

  it('部分重叠但不是包含关系时两块都保留（用户确实选了两个方向）', () => {
    const result = normalizeRanges(['A1:B2', 'B2:C3']);
    expect(result).toHaveLength(2);
  });

  it('结果按行优先排序（日志/快照顺序可复现）', () => {
    expect(rectsToA1(normalizeRanges(['D4', 'A1', 'B2']))).toEqual(['A1', 'B2', 'D4']);
  });

  it('sortRects 按先行后列', () => {
    expect(sortRects([rect(3, 0, 3, 0), rect(0, 1, 0, 1), rect(0, 0, 0, 0)])).toEqual([
      rect(0, 0, 0, 0),
      rect(0, 1, 0, 1),
      rect(3, 0, 3, 0),
    ]);
  });
});

describe('boundingRect', () => {
  it('包住所有区域的最小矩形；空输入返回 null', () => {
    expect(boundingRect([rect(1, 1, 1, 1), rect(3, 2, 4, 5)])).toEqual(rect(1, 1, 4, 5));
    expect(boundingRect([])).toBeNull();
  });
});

describe('summarizeSelection：状态栏文案', () => {
  it('单块保持原样（不破坏既有习惯）', () => {
    expect(summarizeSelection(['B2'])).toBe('B2');
    expect(summarizeSelection(['B2:D5'])).toBe('B2:D5');
    expect(summarizeSelection(['A1:D4', 'B2']), '被包含的小块不影响结论').toBe('A1:D4');
  });

  it('多块给"块数 + 总格数 + 前几块"', () => {
    expect(summarizeSelection(['B2', 'D4'])).toBe('B2 + D4（2 块 / 2 格）');
    expect(summarizeSelection(['A1:B2', 'D4', 'F6'])).toBe('A1:B2 + D4 + F6（3 块 / 6 格）');
  });

  it('块太多时只列前 3 块并给省略号（一屏放不下）', () => {
    const text = summarizeSelection(['A1', 'B1', 'C1', 'D1', 'E1']);
    expect(text).toContain('A1 + B1 + C1 +…');
    expect(text).toContain('（5 块 / 5 格）');
  });

  it('空选区给占位符', () => {
    expect(summarizeSelection([])).toBe('—');
  });
});

describe('describeSingleBlockOnly：只支持单块的操作', () => {
  it('单块返回 null（可以照常执行）', () => {
    expect(describeSingleBlockOnly(['A1:B2'], '与工作区互换')).toBeNull();
    expect(describeSingleBlockOnly(['A1:D4', 'B2'], '粘贴')).toBeNull();
  });

  it('多块给出"为什么不行 + 怎么办"', () => {
    const message = describeSingleBlockOnly(['B2', 'D4'], '与工作区互换');
    expect(message).toContain('2 块');
    expect(message).toContain('与工作区互换');
    expect(message).toContain('Ctrl');
  });
});

describe('parseRangeList：工作区快速导入输入框的多块写法', () => {
  it('空格 / 逗号 / 顿号 / 分号都能分隔', () => {
    for (const text of ['A1:B2 D4:E5', 'A1:B2,D4:E5', 'A1:B2、D4:E5', 'A1:B2;D4:E5']) {
      expect(rectsToA1(parseRangeList(text).rects), text).toEqual(['A1:B2', 'D4:E5']);
    }
  });

  it('单个区域/单格也认（与既有输入兼容）', () => {
    expect(rectsToA1(parseRangeList('B2').rects)).toEqual(['B2']);
    expect(rectsToA1(parseRangeList(' A1:B18 ').rects)).toEqual(['A1:B18']);
    expect(rectsToA1(parseRangeList('$A$1:$B$2').rects)).toEqual(['A1:B2']);
  });

  it('非法片段单独返回（调用方据此报错，而不是整条输入作废）', () => {
    const result = parseRangeList('A1:B2 不是区域 3:');
    expect(rectsToA1(result.rects)).toEqual(['A1:B2']);
    expect(result.invalid).toEqual(['不是区域', '3:']);
  });

  it('空输入 → 空结果（不是错误）', () => {
    expect(parseRangeList('   ').rects).toEqual([]);
    expect(parseRangeList('').invalid).toEqual([]);
  });

  it('重复与被包含的写法会被规范化（搬进工作区不会出现重复条目）', () => {
    expect(rectsToA1(parseRangeList('A1:D4 B2 B2').rects)).toEqual(['A1:D4']);
  });
});
