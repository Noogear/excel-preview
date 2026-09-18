/**
 * 右键菜单纯模型单测（Vitest，node 环境，**无 DOM**）。
 *
 * 只测 `src/shell/context-menu-model.ts`：定位夹取 + 键盘导航。该文件不允许出现
 * document / window，因此不需要 jsdom，也不需要任何桩；组件层（ContextMenu.tsx）的
 * 事件与焦点行为、以及 SSR 结构冒烟属于 e2e 覆盖范围。
 *
 * 用例按"同一主题下的整块场景"合并为 8 个 it（断言一条未删），覆盖：
 *  - clampMenuPosition：常规位置不翻转 / 恰好贴边 / 点击点小于 margin 或为负时抬到 margin；
 *  - clampMenuPosition：margin 语义（默认 8、自定义同时影响"不翻转判定"与"翻转落点"、0 允许贴边、
 *    非法回落 8）与纯函数不改写入参；
 *  - clampMenuPosition：贴边翻转（向右不足向左、向下不足向上、单轴、双轴独立）；
 *  - clampMenuPosition：翻转位侵入边距或放不下时夹到"最大完整可见位置"；菜单大于可视区域时贴左上 margin；
 *  - clampMenuPosition：退化输入（NaN / Infinity / 0 尺寸菜单 / 0 视口）不产生 NaN，网格扫描四边不出界；
 *  - nextEnabledIndex：跳过连续 disabled、相邻可选项走一步、互为反向、边界循环、from = -1、
 *    越界 / 非整数 / NaN 的 from；
 *  - nextEnabledIndex：空数组 / 全 disabled / 单项可选或禁用 / 相邻两项的退化与边界；
 *  - firstEnabledIndex：空数组 / 全 disabled / 首项可选 / 首个可选项位置。
 */
import { describe, expect, it } from 'vitest';

import {
  MENU_VIEWPORT_MARGIN,
  clampMenuPosition,
  firstEnabledIndex,
  nextEnabledIndex,
  type MenuItemSpec,
  type Point,
  type Size,
  type Viewport,
} from '../../src/shell/context-menu-model';

/* -------------------------------------------------------------------------- */
/* 脚手架                                                                      */
/* -------------------------------------------------------------------------- */

/** `'a'` = 可选项，`'-b'` = 禁用项。写用例时一眼能看出禁用位置。 */
function items(...specs: string[]): MenuItemSpec[] {
  return specs.map((spec) =>
    spec.startsWith('-')
      ? { id: spec.slice(1), label: spec.slice(1), disabled: true }
      : { id: spec, label: spec },
  );
}

const MENU: Size = { width: 200, height: 150 };
const VIEW: Viewport = { width: 1000, height: 800 };

function expectPoint(actual: Point, x: number, y: number): void {
  expect({ x: actual.x, y: actual.y }).toEqual({ x, y });
}

/* -------------------------------------------------------------------------- */
/* A. clampMenuPosition                                                        */
/* -------------------------------------------------------------------------- */

describe('clampMenuPosition - 常规位置与 margin', () => {
  it('放得下时贴着点击点（含恰好贴边），点击点小于 margin / 为负时抬到 margin', () => {
    expectPoint(clampMenuPosition({ x: 100, y: 80 }, MENU, VIEW), 100, 80);
    // 1000 - 8 - 200 = 792、800 - 8 - 150 = 642 → 边界取等号仍不翻转
    expectPoint(clampMenuPosition({ x: 792, y: 642 }, MENU, VIEW), 792, 642);
    // 0,0 与负坐标都不会顶出视口
    expectPoint(clampMenuPosition({ x: 0, y: 0 }, MENU, VIEW), MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);
    expectPoint(clampMenuPosition({ x: -50, y: -20 }, MENU, VIEW), MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);
  });

  it('margin 语义：默认 8、自定义同时影响"不翻转判定"与"翻转落点"、0 允许贴边、非法回落 8，且为纯函数', () => {
    expect(MENU_VIEWPORT_MARGIN).toBe(8);
    expectPoint(clampMenuPosition({ x: 999, y: 799 }, MENU, VIEW), 792, 642);
    expectPoint(clampMenuPosition({ x: 999, y: 799 }, MENU, VIEW, MENU_VIEWPORT_MARGIN), 792, 642);
    // margin 24：不翻转要求 x + 200 <= 976；点击 995 不满足 → 翻转 795 → 夹到 976-200 = 776
    expectPoint(clampMenuPosition({ x: 995, y: 795 }, MENU, VIEW, 24), 776, 626);
    // margin 0：800 + 200 = 1000 正好贴右边缘
    expectPoint(clampMenuPosition({ x: 800, y: 650 }, MENU, VIEW, 0), 800, 650);

    // 非法 margin（负数 / NaN）→ 回落默认 8px
    expectPoint(clampMenuPosition({ x: 0, y: 0 }, MENU, VIEW, -20), MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);
    expectPoint(clampMenuPosition({ x: 0, y: 0 }, MENU, VIEW, Number.NaN), MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);

    // 不改写入参（纯函数）
    const anchor: Point = { x: 900, y: 780 };
    const menu: Size = { width: 200, height: 150 };
    const viewport: Viewport = { width: 1000, height: 800 };
    clampMenuPosition(anchor, menu, viewport, 12);
    expect(anchor).toEqual({ x: 900, y: 780 });
    expect(menu).toEqual({ width: 200, height: 150 });
    expect(viewport).toEqual({ width: 1000, height: 800 });
  });
});

describe('clampMenuPosition - 贴边翻转', () => {
  it('右侧 / 下方空间不足各自向左、向上翻转，两轴独立，且一轴翻转不影响另一轴', () => {
    expectPoint(clampMenuPosition({ x: 850, y: 100 }, MENU, VIEW), 650, 100);
    expectPoint(clampMenuPosition({ x: 100, y: 700 }, MENU, VIEW), 100, 550);
    expectPoint(clampMenuPosition({ x: 900, y: 780 }, MENU, VIEW), 700, 630);
    // 一轴翻转、另一轴保持不动
    expectPoint(clampMenuPosition({ x: 950, y: 100 }, MENU, VIEW), 750, 100);
    expectPoint(clampMenuPosition({ x: 100, y: 790 }, MENU, VIEW), 100, 640);
  });

  it('翻转位侵入边距或放不下时夹到"最大完整可见位置"；菜单大于"视口 - 2*margin"时贴左上 margin', () => {
    // 点击点 995 已压进右边缘：翻转位 995-200=795 > 1000-8-200=792 → 用 792
    const p = clampMenuPosition({ x: 995, y: 795 }, MENU, VIEW);
    expectPoint(p, 792, 642);
    expect(p.x + MENU.width).toBe(VIEW.width - MENU_VIEWPORT_MARGIN);
    expect(p.y + MENU.height).toBe(VIEW.height - MENU_VIEWPORT_MARGIN);

    // 菜单超过半屏：翻转位 100-180 = -80 < margin，不可用 → 取 200 - 8 - 180 = 12
    const menu: Size = { width: 180, height: 180 };
    const viewport: Viewport = { width: 200, height: 200 };
    const q = clampMenuPosition({ x: 100, y: 100 }, menu, viewport);
    expectPoint(q, 12, 12);
    expect(q.x + menu.width).toBeLessThanOrEqual(viewport.width - MENU_VIEWPORT_MARGIN);
    expect(q.y + menu.height).toBeLessThanOrEqual(viewport.height - MENU_VIEWPORT_MARGIN);

    // 菜单比"视口 - 2*margin"还大 → 贴左上 margin，保证前缘（图标 + 文字）可见
    const oversized = clampMenuPosition({ x: 50, y: 50 }, { width: 160, height: 200 }, { width: 100, height: 100 });
    expectPoint(oversized, MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);
  });
});

describe('clampMenuPosition - 退化输入与不变量', () => {
  it('NaN / Infinity / 零尺寸 / 零视口都不产生 NaN（0 尺寸退化为点击点），网格扫描四边留够 margin', () => {
    const nan = clampMenuPosition({ x: Number.NaN, y: Number.POSITIVE_INFINITY }, MENU, VIEW);
    expect(Number.isFinite(nan.x)).toBe(true);
    expect(Number.isFinite(nan.y)).toBe(true);
    expectPoint(nan, MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);

    // 尺寸尚未测量出来（0×0）：退化为点击点，仍受 margin 约束
    expectPoint(clampMenuPosition({ x: 300, y: 200 }, { width: 0, height: 0 }, VIEW), 300, 200);
    expectPoint(clampMenuPosition({ x: 2, y: 2 }, { width: 0, height: 0 }, VIEW), MENU_VIEWPORT_MARGIN, MENU_VIEWPORT_MARGIN);

    const zeroView = clampMenuPosition({ x: 50, y: 50 }, MENU, { width: 0, height: 0 });
    expect(Number.isFinite(zeroView.x)).toBe(true);
    expect(Number.isFinite(zeroView.y)).toBe(true);

    // 网格扫描：任何点击点下菜单都完整落在视口内且四边留够 margin
    const menu: Size = { width: 200, height: 150 };
    const viewport: Viewport = { width: 1024, height: 768 };
    for (let ax = -20; ax <= 1044; ax += 37) {
      for (let ay = -20; ay <= 788; ay += 31) {
        const p = clampMenuPosition({ x: ax, y: ay }, menu, viewport, MENU_VIEWPORT_MARGIN);
        expect(p.x).toBeGreaterThanOrEqual(MENU_VIEWPORT_MARGIN);
        expect(p.y).toBeGreaterThanOrEqual(MENU_VIEWPORT_MARGIN);
        expect(p.x + menu.width).toBeLessThanOrEqual(viewport.width - MENU_VIEWPORT_MARGIN);
        expect(p.y + menu.height).toBeLessThanOrEqual(viewport.height - MENU_VIEWPORT_MARGIN);
      }
    }
  });
});

/* -------------------------------------------------------------------------- */
/* B. nextEnabledIndex                                                         */
/* -------------------------------------------------------------------------- */

describe('nextEnabledIndex - 双向导航', () => {
  it('跳过连续 disabled、相邻可选项走一步、互为反向、边界循环、from = -1 与脏 from 的兜底', () => {
    // 两个方向都跳过中间连续的 disabled
    expect(nextEnabledIndex(items('a', '-b', '-c', 'd'), 0, 1)).toBe(3);
    expect(nextEnabledIndex(items('a', '-b', '-c', 'd'), 3, -1)).toBe(0);

    // 相邻可选项之间正常走一步
    const list = items('a', 'b', '-c', 'd');
    expect(nextEnabledIndex(list, 0, 1)).toBe(1);
    expect(nextEnabledIndex(list, 1, 1)).toBe(3);
    // 互为反向：来回一趟回到起点
    const back = items('a', '-b', 'c', '-d', 'e');
    expect(nextEnabledIndex(back, nextEnabledIndex(back, 0, 1), -1)).toBe(0);

    // 到边界后循环：向下回到首个可选项，向上回到末个可选项
    expect(nextEnabledIndex(items('a', '-b', 'c'), 2, 1)).toBe(0);
    expect(nextEnabledIndex(items('-a', '-b', 'c'), 2, 1)).toBe(2); // 循环长度为 1 时回到自己
    expect(nextEnabledIndex(items('a', 'b', '-c'), 0, -1)).toBe(1);
    expect(nextEnabledIndex(items('a', '-b', '-c'), 0, -1)).toBe(0);

    // from = -1 表示"尚无焦点"：向下取首个可选项，向上取末个可选项
    expect(nextEnabledIndex(items('-a', '-b', 'c', 'd'), -1, 1)).toBe(2);
    expect(nextEnabledIndex(items('a', 'b', '-c', 'd'), -1, -1)).toBe(3);

    // from 越界 / 非整数 / NaN 不抛异常，按截断后处理
    expect(nextEnabledIndex(items('-a', 'b', 'c'), 99, 1)).toBe(1);
    expect(nextEnabledIndex(items('a', 'b', '-c'), 99, -1)).toBe(1);
    expect(nextEnabledIndex(items('a', 'b', 'c'), 0.9, 1)).toBe(1);
    expect(nextEnabledIndex(items('a', 'b', 'c'), Number.NaN, 1)).toBe(0);
  });
});

describe('nextEnabledIndex - 边界与退化', () => {
  it('空数组 / 全 disabled → -1；单项可选 → 0、禁用 → -1；相邻两项都可用时不会原地不动', () => {
    expect(nextEnabledIndex([], 0, 1)).toBe(-1);
    expect(nextEnabledIndex([], -1, -1)).toBe(-1);

    const allDisabled = items('-a', '-b', '-c');
    expect(nextEnabledIndex(allDisabled, 0, 1)).toBe(-1);
    expect(nextEnabledIndex(allDisabled, 1, -1)).toBe(-1);
    expect(nextEnabledIndex(allDisabled, -1, 1)).toBe(-1);

    // 只有一项可选时无论方向都回到它自己
    const single = items('-a', 'b', '-c');
    expect(nextEnabledIndex(single, 1, 1)).toBe(1);
    expect(nextEnabledIndex(single, 1, -1)).toBe(1);

    // 单项菜单：可选 → 0，禁用 → -1
    expect(nextEnabledIndex(items('only'), 0, 1)).toBe(0);
    expect(nextEnabledIndex(items('only'), -1, -1)).toBe(0);
    expect(nextEnabledIndex(items('-only'), 0, 1)).toBe(-1);

    // 相邻两项都可用时不会"原地不动"
    expect(nextEnabledIndex(items('a', 'b'), 0, 1)).toBe(1);
    expect(nextEnabledIndex(items('a', 'b'), 1, -1)).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */
/* C. firstEnabledIndex                                                        */
/* -------------------------------------------------------------------------- */

describe('firstEnabledIndex', () => {
  it('空数组 / 全部 disabled → -1；首项可选 → 0；否则返回后面第一个可选项', () => {
    expect(firstEnabledIndex([])).toBe(-1);
    expect(firstEnabledIndex(items('-a', '-b', '-c'))).toBe(-1);
    expect(firstEnabledIndex(items('a', 'b', 'c'))).toBe(0);
    expect(firstEnabledIndex(items('-a', '-b', 'c', 'd'))).toBe(2);
    expect(firstEnabledIndex(items('-a', '-b', '-c', 'd'))).toBe(3);
    expect(firstEnabledIndex(items('only'))).toBe(0);
  });
});
