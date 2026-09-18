/**
 * 工作区的**筛选 / 搜索**（纯函数，不碰 React / DOM / Univer）。
 *
 * 用户要求：能按"内容"搜，也能筛"来自哪张表"。
 *
 * 约定：
 *  - 搜索**不区分大小写**、忽略首尾空格；空串 = 不过滤；
 *  - 搜索范围是"这条目在表格里显示的文本"（`cells[0][0].text`），
 *    另外也匹配来源 A1（例如搜 `B12` 能直接找到那一格）；
 *  - 来源筛选是 `sheetName` 精确匹配（`null` / `'all'` = 全部）。
 */
import type { RangeSnapshot } from './types';

export const ALL_SOURCES = 'all' as const;

export interface WorkspaceFilter {
  /** 搜索关键字（空串不过滤） */
  query: string;
  /** 来源工作表名；`ALL_SOURCES` 表示全部 */
  source: string;
}

export const EMPTY_FILTER: WorkspaceFilter = { query: '', source: ALL_SOURCES };

/** 条目的可见文本（内容优先，取不到就退回 label） */
export function itemText(item: RangeSnapshot): string {
  const text = item.cells?.[0]?.[0]?.text;
  return typeof text === 'string' && text !== '' ? text : item.label;
}

/** 当前工作区里出现过的来源工作表（按首次出现顺序，去重） */
export function listSourceSheets(items: readonly RangeSnapshot[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const item of items) {
    const name = item.source?.sheetName;
    if (typeof name !== 'string' || name === '' || seen.has(name)) continue;
    seen.add(name);
    out.push(name);
  }
  return out;
}

/** 该条目是否命中筛选条件 */
export function matchesFilter(item: RangeSnapshot, filter: WorkspaceFilter): boolean {
  if (filter.source !== ALL_SOURCES && item.source?.sheetName !== filter.source) return false;
  const query = filter.query.trim().toLowerCase();
  if (query === '') return true;
  if (itemText(item).toLowerCase().includes(query)) return true;
  // 也允许用 A1 找：搜 "b12" 能定位到 B12
  return item.source?.a1?.toLowerCase().includes(query) ?? false;
}

/** 应用筛选（保持原顺序；筛选条件为空时原样返回同一个数组） */
export function filterItems(items: readonly RangeSnapshot[], filter: WorkspaceFilter): RangeSnapshot[] {
  if (filter.query.trim() === '' && filter.source === ALL_SOURCES) return items as RangeSnapshot[];
  return items.filter((item) => matchesFilter(item, filter));
}

/** 筛选是否处于"生效中"状态（用于 UI 提示与一键清除按钮） */
export function isFilterActive(filter: WorkspaceFilter): boolean {
  return filter.query.trim() !== '' || filter.source !== ALL_SOURCES;
}
