/**
 * 工作区预览纯函数单测（node 环境，不需要 jsdom）。
 *
 * 覆盖：
 *  - pt→px 换算（11pt ≈ 14.67px）与缩放系数叠加、非法字号/缩放不产出 NaN
 *  - "未提供的字段不输出"（只给 fill 就只出 background；不出现 NaN / undefinedpx）
 *  - 四边边框独立输出、缺边不输出
 *  - rotate 为 0 / undefined / NaN 时不输出 transform
 *  - 对齐（align → textAlign、vAlign → verticalAlign）与换行（wrap）
 *  - 工作区筛选纯函数（filterItems / matchesFilter / listSourceSheets / isFilterActive / itemText）
 *  - WorkspacePanel 的事件契约与 SnapshotPreview 的渲染契约
 */
import { isValidElement, type CSSProperties, type ReactElement, type ReactNode } from 'react';
import { describe, expect, it } from 'vitest';

import { DEFAULT_SETTINGS } from '../../src/persistence/session';
import {
  ALL_SOURCES,
  EMPTY_FILTER,
  filterItems,
  isFilterActive,
  itemText,
  listSourceSheets,
  matchesFilter,
} from '../../src/workspace/filter';
import { snapshotStyleToCss } from '../../src/workspace/preview-style';
import { SnapshotPreview } from '../../src/workspace/SnapshotPreview';
import { WorkspacePanel } from '../../src/workspace/WorkspacePanel';
import type { RangeSnapshot, SnapshotCell, SnapshotStyle } from '../../src/workspace/types';

/** 取排好序的键名，避免依赖对象字面量的书写顺序 */
function keysOf(css: object): string[] {
  return Object.keys(css).sort();
}

/** 所有值拼成字符串，用来断言"没有 NaN / undefinedpx 这类脏值" */
function dump(css: object): string {
  return Object.entries(css)
    .map(([k, v]) => `${k}=${String(v)}`)
    .join(';');
}

/* -------------------------------------------------------------------------- */
/* snapshotStyleToCss                                                          */
/* -------------------------------------------------------------------------- */

describe('snapshotStyleToCss - 字号 pt→px 与缩放', () => {
  it('11pt → 14.67px、12pt → 16px，缩放系数叠加在换算之后，且只影响字号', () => {
    expect(snapshotStyleToCss({ fontSize: 11 }).fontSize).toBe('14.67px');
    expect(snapshotStyleToCss({ fontSize: 12 }).fontSize).toBe('16px'); // 整数不补小数位
    expect(snapshotStyleToCss({ fontSize: 11 }, 0.5).fontSize).toBe('7.33px');
    expect(snapshotStyleToCss({ fontSize: 11 }, 2).fontSize).toBe('29.33px');
    // scale 默认 1，等价于不传
    expect(snapshotStyleToCss({ fontSize: 11 }, 1)).toEqual(snapshotStyleToCss({ fontSize: 11 }));

    const css = snapshotStyleToCss({ fontSize: 11, border: { top: '#111111' } }, 0.25);
    expect(css.fontSize).toBe('3.67px');
    expect(css.borderTop).toBe('1px solid #111111'); // 边框不跟着缩放

    // 字号非法 → 不输出 fontSize，也不产出 NaN 文本
    for (const bad of [0, -12, Number.NaN, Number.POSITIVE_INFINITY]) {
      const badCss = snapshotStyleToCss({ fontSize: bad });
      expect(badCss.fontSize).toBeUndefined();
      expect(keysOf(badCss)).toEqual([]);
      expect(dump(badCss)).not.toMatch(/NaN|Infinity|undefinedpx/);
    }
    // scale 非法时按 1 处理
    for (const bad of [0, -1, Number.NaN]) {
      expect(snapshotStyleToCss({ fontSize: 11 }, bad).fontSize).toBe('14.67px');
    }
  });
});

describe('snapshotStyleToCss - 未提供的字段不输出', () => {
  it('只给 fill 就只出 background；color / fontFamily 各自独立透传；空串 / 全空格不输出', () => {
    const css = snapshotStyleToCss({ fill: '#FFCC00' });
    expect(css).toEqual({ background: '#FFCC00' });
    expect(keysOf(css)).toEqual(['background']);

    expect(keysOf(snapshotStyleToCss({ color: '#123456' }))).toEqual(['color']);
    expect(keysOf(snapshotStyleToCss({ fontFamily: 'Arial' }))).toEqual(['fontFamily']);
    // 避免 1px solid 这类残值
    expect(keysOf(snapshotStyleToCss({ fill: '' }))).toEqual([]);
    expect(keysOf(snapshotStyleToCss({ fill: '   ' }))).toEqual([]);
  });

  it('边框四边独立输出，缺边为空，空对象 / 空串值不输出任何边框', () => {
    const all = snapshotStyleToCss({
      border: { top: '#111111', right: '#222222', bottom: '#333333', left: '#444444' },
    });
    expect(all.borderTop).toBe('1px solid #111111');
    expect(all.borderRight).toBe('1px solid #222222');
    expect(all.borderBottom).toBe('1px solid #333333');
    expect(all.borderLeft).toBe('1px solid #444444');
    expect(keysOf(all)).toEqual(['borderBottom', 'borderLeft', 'borderRight', 'borderTop']);

    const onlyTop = snapshotStyleToCss({ border: { top: '#111111' } });
    expect(keysOf(onlyTop)).toEqual(['borderTop']);
    expect(onlyTop.borderRight).toBeUndefined();
    expect(onlyTop.borderBottom).toBeUndefined();
    expect(onlyTop.borderLeft).toBeUndefined();

    expect(keysOf(snapshotStyleToCss({ border: {} }))).toEqual([]);
    expect(keysOf(snapshotStyleToCss({ border: { top: '', left: '  ' } }))).toEqual([]);

    // rotate：只在非 0 时输出 transform，且角度收敛到 2 位小数
    expect(keysOf(snapshotStyleToCss({}))).toEqual([]);
    expect(keysOf(snapshotStyleToCss(undefined))).toEqual([]);
    expect(keysOf(snapshotStyleToCss({ rotate: 0 }))).toEqual([]);
    expect('transform' in snapshotStyleToCss({ rotate: 0 })).toBe(false);
    expect(keysOf(snapshotStyleToCss({ rotate: Number.NaN }))).toEqual([]);

    expect(snapshotStyleToCss({ rotate: 45 }).transform).toBe('rotate(45deg)');
    expect(snapshotStyleToCss({ rotate: -30 }).transform).toBe('rotate(-30deg)');
    expect(snapshotStyleToCss({ rotate: 12.345 }).transform).toBe('rotate(12.35deg)');
    expect(snapshotStyleToCss({ rotate: 33.3333 }).transform).toBe('rotate(33.33deg)');
  });
});

describe('snapshotStyleToCss - 字形 / 对齐 / 换行', () => {
  it('bold / italic / 下划线 / 删除线映射到字形字段，显式 false 不输出', () => {
    const css = snapshotStyleToCss({ bold: true, italic: true, underline: true, strikeThrough: true });
    expect(css.fontWeight).toBe('bold');
    expect(css.fontStyle).toBe('italic');
    expect(css.textDecoration).toBe('underline line-through');
    expect(snapshotStyleToCss({ underline: true }).textDecoration).toBe('underline');
    expect(snapshotStyleToCss({ strikeThrough: true }).textDecoration).toBe('line-through');

    expect(keysOf(snapshotStyleToCss({ bold: false, italic: false, underline: false, strikeThrough: false }))).toEqual([]);
  });

  it('align / vAlign 只走 textAlign + verticalAlign（不引入 flex，避免 td 失去 table-cell 身份）', () => {
    const aligned = snapshotStyleToCss({ align: 'center' });
    expect(aligned.textAlign).toBe('center');
    expect(aligned.display).toBeUndefined();
    expect(aligned.verticalAlign).toBeUndefined();
    expect(keysOf(aligned)).toEqual(['textAlign']);

    const both = snapshotStyleToCss({ align: 'right', vAlign: 'bottom' });
    expect(both.textAlign).toBe('right');
    expect(both.verticalAlign).toBe('bottom');
    expect(both.display).toBeUndefined();
    expect(both.alignItems).toBeUndefined();
    expect(keysOf(both)).toEqual(['textAlign', 'verticalAlign']);

    expect(snapshotStyleToCss({ vAlign: 'top' }).verticalAlign).toBe('top');
    expect(snapshotStyleToCss({ vAlign: 'middle' }).verticalAlign).toBe('middle');
    expect(snapshotStyleToCss({ vAlign: 'bottom' }).verticalAlign).toBe('bottom');
  });

  it('wrap=true → pre-wrap + break-all；wrap=false → nowrap；numberFormat 不产生任何 CSS；综合脏值不产出 NaN', () => {
    const wrapped = snapshotStyleToCss({ wrap: true });
    expect(wrapped.whiteSpace).toBe('pre-wrap');
    expect(wrapped.wordBreak).toBe('break-all');
    expect(snapshotStyleToCss({ wrap: false }).whiteSpace).toBe('nowrap');
    expect(keysOf(snapshotStyleToCss({ wrap: false }))).toEqual(['whiteSpace']);
    expect(keysOf(snapshotStyleToCss({}))).toEqual([]);
    // numberFormat 只是回溯信息
    expect(keysOf(snapshotStyleToCss({ numberFormat: '0.00%' }))).toEqual([]);
  });

  it('全字段组合输出干净；运行期脏枚举值整体忽略', () => {
    const style: SnapshotStyle = {
      fontFamily: 'Arial',
      fontSize: 11,
      bold: true,
      italic: true,
      underline: true,
      color: '#000000',
      fill: '#FFFFFF',
      align: 'left',
      vAlign: 'middle',
      wrap: true,
      rotate: 90,
      border: { top: '#000000', bottom: '#000000' },
    };
    const text = dump(snapshotStyleToCss(style));
    expect(text).not.toMatch(/NaN|undefinedpx|=undefined/);
    expect(Object.values(snapshotStyleToCss(style)).every((v) => v !== undefined)).toBe(true);

    // 枚举值不在联合类型内（运行期脏数据）→ 不输出无效声明
    const dirty = { align: 'middle', vAlign: 'center', fontSize: '11', rotate: '45' } as unknown as SnapshotStyle;
    expect(keysOf(snapshotStyleToCss(dirty))).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* -------------------------------------------------------------------------- */

/* -------------------------------------------------------------------------- */
/* 工作区筛选（纯函数：搜索内容 / 搜 A1 / 按来源筛）                             */
/* -------------------------------------------------------------------------- */

/** 造一份"某张表某个格子"的条目 */
function cellItem(id: string, sheetName: string, a1: string, text: string): RangeSnapshot {
  return {
    id,
    label: a1,
    rows: 1,
    cols: 1,
    createdAt: 1,
    source: { sheetName, a1, startRow: 0, startCol: 0, endRow: 0, endCol: 0, sheetId: 's1' },
    cells: [[{ text }]],
    values: [[text]],
    formulas: [[null]],
  } as unknown as RangeSnapshot;
}

describe('工作区筛选纯函数', () => {
  const items = [
    cellItem('i1', '成绩表', 'B2', '张三'),
    cellItem('i2', '成绩表', 'B12', '李四'),
    cellItem('i3', '花名册', 'D9', 'Zhang San'),
  ];

  it('itemText 内容优先、空内容退回 label', () => {
    expect(itemText(items[0])).toBe('张三');
    const blank = { ...items[0], cells: [[{ text: '' }]], label: 'H8' } as unknown as RangeSnapshot;
    expect(itemText(blank)).toBe('H8');
    const missing = { ...items[0], cells: [] } as unknown as RangeSnapshot;
    expect(itemText(missing)).toBe(items[0].label);
  });

  it('listSourceSheets 去重且按首次出现顺序', () => {
    expect(listSourceSheets(items)).toEqual(['成绩表', '花名册']);
    expect(listSourceSheets([])).toEqual([]);
    // 表名为空 / 缺失的来源不进列表
    const dirty = [{ ...items[0], source: { ...items[0].source, sheetName: '' } }] as unknown as RangeSnapshot[];
    expect(listSourceSheets(dirty)).toEqual([]);
  });

  it('搜索：不分大小写、忽略首尾空格，命中内容或来源 A1', () => {
    expect(filterItems(items, { query: '张三', source: ALL_SOURCES }).map((i) => i.id)).toEqual(['i1']);
    expect(filterItems(items, { query: '  zhang  ', source: ALL_SOURCES }).map((i) => i.id), '大小写与空格都不影响').toEqual(['i3']);
    expect(filterItems(items, { query: 'b12', source: ALL_SOURCES }).map((i) => i.id), '搜 A1 能定位').toEqual(['i2']);
    expect(filterItems(items, { query: '不存在', source: ALL_SOURCES })).toEqual([]);
    expect(matchesFilter(items[0], { query: '张三', source: '花名册' }), '来源不符就不命中').toBe(false);
  });

  it('来源筛选是精确匹配；条件为空时原样返回同一个数组', () => {
    expect(filterItems(items, { query: '', source: '花名册' }).map((i) => i.id)).toEqual(['i3']);
    expect(filterItems(items, { query: '成绩', source: '花名册' }), '两个条件是"与"').toEqual([]);
    expect(filterItems(items, EMPTY_FILTER), '不过滤时返回同一引用（省一次数组分配）').toBe(items);
    expect(filterItems(items, { query: '   ', source: ALL_SOURCES })).toBe(items);
  });

  it('isFilterActive：搜索或来源任一非默认即为生效', () => {
    expect(isFilterActive(EMPTY_FILTER)).toBe(false);
    expect(isFilterActive({ query: '  ', source: ALL_SOURCES })).toBe(false);
    expect(isFilterActive({ query: '张', source: ALL_SOURCES })).toBe(true);
    expect(isFilterActive({ query: '', source: '花名册' })).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* 测试脚手架（SSR 式：直接调用组件函数拿 React 元素树，node 环境无需 DOM）        */
/* -------------------------------------------------------------------------- */

/** 一份 2×2 的快照：首格带 11pt 字号（换算成 14.67px），第三格是合并区 */
function makeSnapshot(): RangeSnapshot {
  return {
    id: 'snap-1',
    label: 'B2:C3',
    rows: 2,
    cols: 2,
    createdAt: 1,
    source: { sheetName: '成绩表', a1: 'B2:C3', startRow: 1, startCol: 1, endRow: 2, endCol: 2, sheetId: 's1' },
    cells: [
      [{ text: '张三', style: { fontSize: 11 } }, { text: '98.5' }],
      // 第 4 格是被合并覆盖的从属格：不渲染（所以可见 td 恰好 3 个）
      [{ text: '合并' }, { text: '', covered: true }],
    ],
  } as unknown as RangeSnapshot;
}

const NOOP_PROPS = {
  onRemove: () => undefined,
  onClearAll: () => undefined,
  onItemPointerDown: () => undefined,
  onItemContextMenu: () => undefined,
  onAddItem: () => undefined,
  onSettingsChange: () => undefined,
  settings: DEFAULT_SETTINGS,
  // 搜索 / 来源筛选状态在上层；契约测试默认"不过滤"
  filter: EMPTY_FILTER,
  onFilterChange: () => undefined,
};

/** 深度遍历 React 元素树（含函数组件返回的子树） */
function collect(node: ReactNode): ReactElement[] {
  const out: ReactElement[] = [];
  const walk = (value: ReactNode): void => {
    if (Array.isArray(value)) {
      value.forEach(walk);
      return;
    }
    if (!isValidElement(value)) return;
    const el = value as ReactElement<{ children?: ReactNode }>;
    out.push(el);
    walk(el.props.children);
  };
  walk(node);
  return out;
}

function byTestId(node: ReactNode, testId: string): ReactElement<Record<string, unknown>>[] {
  return collect(node).filter((el) => (el.props as Record<string, unknown>)['data-testid'] === testId) as ReactElement<
    Record<string, unknown>
  >[];
}

/** 把元素子树的文本拼起来（用于断言可见文案） */
function textOf(node: ReactNode): string {
  const parts: string[] = [];
  const walk = (value: ReactNode): void => {
    if (typeof value === 'string' || typeof value === 'number') {
      parts.push(String(value));
      return;
    }
    if (Array.isArray(value)) {
      value.forEach(walk);
      return;
    }
    if (isValidElement(value)) walk((value.props as { children?: ReactNode }).children);
  };
  walk(node);
  return parts.join('');
}

describe('WorkspacePanel 事件契约', () => {
  it('卡片根元素带 testid/data-snapshot-id，pointerdown 原样转发且不 preventDefault；点击正文触发 activate', () => {
    const item = makeSnapshot();
    const received: { event: unknown; item: RangeSnapshot }[] = [];
    const activated: RangeSnapshot[] = [];
    let prevented = 0;
    const fakeEvent = {
      type: 'pointerdown',
      preventDefault: () => {
        prevented += 1;
      },
    };

    const tree = WorkspacePanel({
      items: [item],
      ...NOOP_PROPS,
      onItemPointerDown: (event, forwarded) => {
        received.push({ event, item: forwarded });
      },
    });

    const cards = byTestId(tree, 'workspace-item');
    expect(cards).toHaveLength(1);
    expect(cards[0].props['data-snapshot-id']).toBe('snap-1');

    (cards[0].props['onPointerDown'] as (event: unknown) => void)(fakeEvent);
    expect(received).toHaveLength(1);
    expect(received[0].event).toBe(fakeEvent); // 原样透传，未被包装
    expect(received[0].item).toBe(item);
    expect(prevented).toBe(0); // 关键：不能 preventDefault，上层手势层要用完整事件

  });

  it('卡片正文里没有任何按钮（× 角标已移除，删除只走右键菜单）；清空按钮只在有条目时渲染', () => {
    const item = makeSnapshot();
    let cleared = 0;

    const tree = WorkspacePanel({
      items: [item],
      ...NOOP_PROPS,
      onClearAll: () => {
        cleared += 1;
      },
    });

    // 新模型：一格一个方块，卡片里只有单元格预览，没有任何按钮
    expect(byTestId(tree, 'workspace-remove'), '× 角标已移除').toHaveLength(0);
    expect(collect(tree).filter((el) => el.type === 'button' && (el.props as { 'data-testid'?: string })['data-testid'] === undefined)).toHaveLength(0);
    expect(textOf(byTestId(tree, 'workspace-item')[0])).toBe('');

    const clearButtons = byTestId(tree, 'workspace-clear');
    expect(clearButtons).toHaveLength(1);
    // 不传 onRequestClear 时（独立渲染/契约测试）点「清空」直接清空
    (clearButtons[0].props['onClick'] as () => void)();
    expect(cleared).toBe(1);
    expect(byTestId(WorkspacePanel({ items: [], ...NOOP_PROPS }), 'workspace-clear')).toHaveLength(0);
  });

  it('清空的二次确认：确认条只在 clearConfirming 时出现，三个动作各自回调（含"只清显示项"）', () => {
    const a: RangeSnapshot = { ...makeSnapshot(), id: 'snap-a', cells: [[{ text: '张三' }]] } as unknown as RangeSnapshot;
    const b: RangeSnapshot = { ...makeSnapshot(), id: 'snap-b', cells: [[{ text: '李四' }]] } as unknown as RangeSnapshot;

    // 未确认时不渲染确认条
    expect(byTestId(WorkspacePanel({ items: [a, b], ...NOOP_PROPS }), 'workspace-clear-confirm')).toHaveLength(0);

    const requested: number[] = [];
    const cancelled: number[] = [];
    const removed: string[] = [];
    let clearedAll = 0;

    const tree = WorkspacePanel({
      items: [a, b],
      filter: { query: '张三', source: ALL_SOURCES },
      onFilterChange: () => undefined,
      onRemove: (id) => removed.push(id),
      onClearAll: () => {
        clearedAll += 1;
      },
      onItemPointerDown: () => undefined,
      onItemContextMenu: () => undefined,
      onAddItem: () => undefined,
      onSettingsChange: () => undefined,
      settings: DEFAULT_SETTINGS,
      clearConfirming: true,
      onRequestClear: () => requested.push(1),
      onCancelClear: () => cancelled.push(1),
    });

    const confirm = byTestId(tree, 'workspace-clear-confirm')[0];
    expect(confirm.props['role']).toBe('alertdialog');
    expect(textOf(confirm), '筛选生效时文案带上"显示几个 / 全部几个"').toContain('清空显示的 1 个，还是全部 2 个？');

    // 「清空显示的」→ 只对当前可见的那一条调 onRemove，并收起确认条
    (byTestId(tree, 'workspace-clear-visible')[0].props['onClick'] as () => void)();
    expect(removed).toEqual(['snap-a']);
    expect(cancelled).toHaveLength(1);

    // 「清空全部」→ onClearAll + 收起；「取消」→ 只收起
    (byTestId(tree, 'workspace-clear-all')[0].props['onClick'] as () => void)();
    expect(clearedAll).toBe(1);
    expect(cancelled).toHaveLength(2);
    (byTestId(tree, 'workspace-clear-cancel')[0].props['onClick'] as () => void)();
    expect(cancelled).toHaveLength(3);

    // 点顶部的「清空」按钮 → 走 onRequestClear（而不是直接清空）
    (byTestId(tree, 'workspace-clear')[0].props['onClick'] as () => void)();
    expect(requested).toHaveLength(1);
    expect(clearedAll, '点「清空」不应直接清空').toBe(1);

    // 无筛选时：文案与按钮集合都变（没有"只清显示项"）
    const plain = WorkspacePanel({
      items: [a, b],
      ...NOOP_PROPS,
      clearConfirming: true,
      onRequestClear: () => undefined,
      onCancelClear: () => undefined,
    });
    expect(textOf(byTestId(plain, 'workspace-clear-confirm')[0])).toContain('确定清空全部 2 个单元格？');
    expect(byTestId(plain, 'workspace-clear-visible')).toHaveLength(0);
  });

  it('新模型：一格一个方块——按 cellSize 展示（夹在 tileMinWidth−16 … 200 × 18–64）、卡片只留预览 + 悬浮删除，来源收进 title', () => {
    // 一块 1×1 条目：来源 B3，来源单元格尺寸 200×12（下限/上限都要被试到）
    const cell: RangeSnapshot = {
      id: 'snap-cell',
      label: 'B3',
      rows: 1,
      cols: 1,
      createdAt: 1,
      source: { sheetName: '样式', a1: 'B3', startRow: 2, startCol: 1, endRow: 2, endCol: 1, sheetId: 's1' },
      cells: [[{ text: '98.5' }]],
      cellSize: { width: 200, height: 12 },
    } as unknown as RangeSnapshot;

    const tree = WorkspacePanel({ items: [cell], ...NOOP_PROPS });

    // 卡片：正文里**没有**"来源 A1"角标了（角标已删除），来源信息只在 title 提示里
    const card = byTestId(tree, 'workspace-item')[0];
    expect(card.props['title'], 'title 里仍能查到内容与来源').toContain('· 来源 样式!B3');
    expect(card.props['title']).toContain('98.5');
    expect(textOf(card), '正文里既没有来源角标也没有 × 按钮').toBe('');
    expect(collect(card).filter((el) => el.type === SnapshotPreview)).toHaveLength(1);

    // 尺寸夹取：宽度 200 → 200（上限），高度 12 → 18（下限）
    const preview = collect(card).find((el) => el.type === SnapshotPreview)!;
    expect(preview.props['cellWidth']).toBe(200);
    expect(preview.props['cellHeight']).toBe(18);
    // 下限跟着「格宽」设置走：tileMinWidth−16
    expect(preview.props['cellWidth'], 'cellWidth 上限固定 200').toBeLessThanOrEqual(200);

    // 拿不到 cellSize 时回落默认值（104×24），不会产出 NaN / 0
    const noSize: RangeSnapshot = { ...cell, id: 'snap-nosize', cellSize: undefined } as unknown as RangeSnapshot;
    const fallback = collect(WorkspacePanel({ items: [noSize], ...NOOP_PROPS })).find(
      (el) => el.type === SnapshotPreview,
    )!;
    expect(fallback.props['cellWidth']).toBe(104);
    expect(fallback.props['cellHeight']).toBe(24);

    // 窄格子被下限托住：cellSize.width=10 → tileMinWidth−16 = 112（默认 128）
    const tiny: RangeSnapshot = { ...cell, id: 'snap-tiny', cellSize: { width: 10, height: 8 } } as unknown as RangeSnapshot;
    const tinyPreview = collect(WorkspacePanel({ items: [tiny], ...NOOP_PROPS })).find(
      (el) => el.type === SnapshotPreview,
    )!;
    expect(tinyPreview.props['cellWidth']).toBe(DEFAULT_SETTINGS.tileMinWidth - 16);
    expect(tinyPreview.props['cellHeight']).toBe(18);
    // 把「格宽」调大，下限跟着变大
    const wider = collect(
      WorkspacePanel({
        items: [tiny],
        ...NOOP_PROPS,
        settings: { ...DEFAULT_SETTINGS, tileMinWidth: 200 },
      }),
    ).find((el) => el.type === SnapshotPreview)!;
    expect(wider.props['cellWidth']).toBe(184);

    // 「每行放几格」不再由设置直接给，而是**面板宽度 ÷ 格宽**自动决定：
    // 挂 `data-tile-min-width` + `repeat(auto-fill, minmax(Npx, 1fr))` 交给 CSS
    const listEl = collect(WorkspacePanel({ items: [cell], ...NOOP_PROPS })).find((el) => el.type === 'ul');
    expect(listEl?.props['data-tile-min-width']).toBe(DEFAULT_SETTINGS.tileMinWidth);
    expect(String((listEl?.props['style'] as CSSProperties | undefined)?.gridTemplateColumns)).toBe(
      `repeat(auto-fill, minmax(${DEFAULT_SETTINGS.tileMinWidth}px, 1fr))`,
    );
    const custom = collect(
      WorkspacePanel({ items: [cell], ...NOOP_PROPS, settings: { ...DEFAULT_SETTINGS, tileMinWidth: 96 } }),
    ).find((el) => el.type === 'ul');
    expect(custom?.props['data-tile-min-width']).toBe(96);
    expect(custom?.props['data-visible-count'], '可见条目数挂出来供筛选断言').toBe(1);
  });

  it('搜索 / 来源筛选：命中过滤、可见数挂 data-visible-count、筛选生效时出现「清除」', () => {
    const a: RangeSnapshot = {
      ...makeSnapshot(),
      id: 'snap-a',
      cells: [[{ text: '张三' }]],
      source: { ...makeSnapshot().source, sheetName: '成绩表', a1: 'B2' },
    } as unknown as RangeSnapshot;
    const b: RangeSnapshot = {
      ...makeSnapshot(),
      id: 'snap-b',
      cells: [[{ text: '李四' }]],
      source: { ...makeSnapshot().source, sheetName: '花名册', a1: 'D9' },
    } as unknown as RangeSnapshot;

    // 按内容搜：命中 1 条，其余不渲染
    const byQuery = WorkspacePanel({ items: [a, b], ...NOOP_PROPS, filter: { query: '李四', source: ALL_SOURCES } });
    expect(byQuery.props.children).toBeTruthy();
    expect(byTestId(byQuery, 'workspace-item')).toHaveLength(1);
    expect(byTestId(byQuery, 'workspace-item')[0].props['data-snapshot-id']).toBe('snap-b');
    const listByQuery = collect(byQuery).find((el) => el.type === 'ul');
    expect(listByQuery?.props['data-visible-count']).toBe(1);
    expect(byTestId(byQuery, 'workspace-filter-clear'), '筛选生效时出现「清除」').toHaveLength(1);

    // 按来源筛：只留花名册
    const bySource = WorkspacePanel({
      items: [a, b],
      ...NOOP_PROPS,
      filter: { query: '', source: '成绩表' },
    });
    expect(byTestId(bySource, 'workspace-item')).toHaveLength(1);
    expect(byTestId(bySource, 'workspace-item')[0].props['data-snapshot-id']).toBe('snap-a');

    // 搜索框与来源下拉只在有条目时渲染；输入变更回调把 query 抛给上层
    const patches: Array<Record<string, unknown>> = [];
    const withInput = WorkspacePanel({
      items: [a, b],
      ...NOOP_PROPS,
      onFilterChange: (patch) => patches.push(patch),
    });
    const search = byTestId(withInput, 'workspace-search')[0];
    (search.props['onChange'] as (e: { target: { value: string } }) => void)({ target: { value: '张' } });
    expect(patches).toEqual([{ query: '张' }]);
    const select = byTestId(withInput, 'workspace-source-filter')[0];
    (select.props['onChange'] as (e: { target: { value: string } }) => void)({ target: { value: '成绩表' } });
    expect(patches[1]).toEqual({ source: '成绩表' });
    expect(byTestId(WorkspacePanel({ items: [], ...NOOP_PROPS }), 'workspace-search')).toHaveLength(0);
  });

  it('空状态给出拖拽提示；条目是一格一个小方块（正文无角标、来源只在 title）并各挂一个缩略图', () => {
    const item = makeSnapshot();
    const emptyTree = WorkspacePanel({ items: [], ...NOOP_PROPS });
    expect(textOf(byTestId(emptyTree, 'workspace-empty')[0])).toContain('拖拽表格中的单元格到此处暂存');

    const listTree = WorkspacePanel({ items: [item, { ...makeSnapshot(), id: 'snap-2' }], ...NOOP_PROPS });
    const cards = byTestId(listTree, 'workspace-item');
    // 新模型：条目 = 一个个独立单元格；卡片正文里既没有"来源 A1"角标也没有 ×（删除走右键菜单）
    expect(textOf(cards[0])).toBe('');
    expect(textOf(cards[0])).not.toContain('来源');
    // 来源信息仍然可查：title = 「{内容} · 来源 {工作表}!{A1}（右键可复制/剪切/删除）」
    expect(cards[0].props['title']).toBe('张三 · 来源 成绩表!B2:C3（右键可复制/剪切/删除）');

    const previews = collect(listTree).filter((el) => el.type === SnapshotPreview);
    expect(previews.map((el) => (el.props['snapshot'] as RangeSnapshot).id)).toEqual(['snap-1', 'snap-2']);
    expect(cards.map((el) => el.props['data-snapshot-id'])).toEqual(['snap-1', 'snap-2']);
  });
});

describe('SnapshotPreview 渲染契约', () => {
  it('根元素带 testid/data-snapshot-id 与行列元信息；covered 不渲染、文字落进 td、字号换算；合并区带 rowSpan/colSpan', () => {
    const rendered = SnapshotPreview({ snapshot: makeSnapshot() });
    const root = byTestId(rendered, 'snapshot-preview');
    expect(root).toHaveLength(1);
    expect(root[0].props['data-snapshot-id']).toBe('snap-1');
    // 紧凑小表格：不再整体缩放，改为固定单元格尺寸 + 一次显示 N 列（元信息挂在 data-* 上便于断言）
    expect(root[0].props['data-cols']).toBe(2);
    expect(root[0].props['data-visible-cols']).toBe(2);

    const tds = collect(rendered).filter((el) => el.type === 'td');
    expect(tds.map((el) => el.props['children'])).toEqual(['张三', '98.5', '合并']);
    expect((tds[0].props['style'] as CSSProperties).fontSize).toBe('14.67px');
    // 预览是概览：即使来源开了自动换行也强制单行 + 省略号，否则窄格子里会"一字一行"
    expect((tds[0].props['style'] as CSSProperties).whiteSpace).toBe('nowrap');
    expect((tds[0].props['style'] as CSSProperties).textOverflow).toBe('ellipsis');

    // 合并区左上角带 rowSpan/colSpan，被覆盖的从属单元格一个都不渲染
    const merged: RangeSnapshot = {
      ...makeSnapshot(),
      rows: 2,
      cols: 2,
      cells: [
        [{ text: '标题', merge: { rowSpan: 2, colSpan: 2 } }],
        [{ text: '', covered: true }, { text: '', covered: true }],
      ],
    };
    const mergedTds = collect(SnapshotPreview({ snapshot: merged })).filter((el) => el.type === 'td');
    expect(mergedTds).toHaveLength(1);
    expect(mergedTds[0].props['rowSpan']).toBe(2);
    expect(mergedTds[0].props['colSpan']).toBe(2);
    expect(mergedTds[0].props['children']).toBe('标题');
  });

  it('列数限制：一次只渲染 N 列（默认 6），超出时提示还有多少列/行', () => {
    const big: RangeSnapshot = { ...makeSnapshot(), rows: 12, cols: 10, cells: [] };

    // 默认 6 列：10 列的选区只渲染 6 列，提示 +4 列 / +2 行（默认最多 10 行）
    const byDefault = SnapshotPreview({ snapshot: big });
    const root = byTestId(byDefault, 'snapshot-preview')[0];
    expect(root.props['data-visible-cols']).toBe(6);
    expect(textOf(byTestId(byDefault, 'snapshot-preview-more')[0])).toContain('+4 列');
    expect(textOf(byTestId(byDefault, 'snapshot-preview-more')[0])).toContain('+2 行');
    expect(collect(byDefault).filter((el) => el.type === 'col')).toHaveLength(6);

    // 指定 3 列 → 只渲染 3 列，提示 +7 列
    const narrow = SnapshotPreview({ snapshot: big, columns: 3 });
    expect(byTestId(narrow, 'snapshot-preview')[0].props['data-visible-cols']).toBe(3);
    expect(textOf(byTestId(narrow, 'snapshot-preview-more')[0])).toContain('+7 列');

    // 选区本身比限制还窄 → 不渲染"还有更多"提示
    const small = SnapshotPreview({ snapshot: makeSnapshot(), columns: 6 });
    expect(byTestId(small, 'snapshot-preview-more')).toHaveLength(0);
  });
});
