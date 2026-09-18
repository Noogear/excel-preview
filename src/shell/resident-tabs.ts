/**
 * 多标签"常驻窗口"记账（纯函数）。
 *
 * 背景（内存）：Univer 每个**实体化**的工作簿都会建模型 + 渲染器 + 插件资源。
 * 百万格量级下每个标签 ≈ 480–670 MB（见 `P3-内容锁定与工作区-交付说明.md` 的内存章节与
 * `M4-多标签内存分析.md`）。开着 N 个标签就 O(N)——所以要把"打开着的标签"和
 * "当前实体化的标签"分开：标签可以开很多，但**同时实体化的**只保留最近的 K 个，
 * 其余冷存（dispose 掉 Univer 工作簿，字节仍在内存里，切回时重建）。
 *
 * 这里只做**决策**（谁该被冷存、切过去要不要重建），不碰 DOM 也不碰 Univer，便于单测。
 */

/** 同时实体化的标签上限（含当前活动标签）。取 3：够"来回对照两张表"，又不会堆内存 */
export const DEFAULT_RESIDENT_TAB_LIMIT = 3;

export interface TabRuntime {
  id: string;
  /** 该标签的 Univer 工作簿当前是否实体化 */
  built: boolean;
  /** 最近一次成为活动标签的时间（毫秒），用于 LRU */
  lastUsedAt: number;
}

/** 切到某标签时是否需要先重建（冷标签） */
export function needsBuild(tabs: readonly TabRuntime[], id: string): boolean {
  const entry = tabs.find((tab) => tab.id === id);
  // 记账里没有的（理论上不该发生）按"需要重建"处理：重建是幂等的，宁可多花一次解析
  return entry ? !entry.built : true;
}

/**
 * 选出这次要冷存的标签：**非活动**、已实体化、且按"最久未用"排序超出窗口的部分。
 *
 * 规则：
 *  - 绝不动当前活动标签（用户正看着它）；
 *  - 只在实体化数量**超过**上限时才动手，一次可以淘汰多个；
 *  - `lastUsedAt` 相同时按 id 稳定排序，保证同一状态下决策可重复（便于测试与排查）。
 */
export function planEvictions(
  tabs: readonly TabRuntime[],
  activeId: string | null,
  limit: number = DEFAULT_RESIDENT_TAB_LIMIT,
): string[] {
  const safeLimit = Math.max(1, Math.floor(limit));
  const built = tabs.filter((tab) => tab.built);
  if (built.length <= safeLimit) return [];
  const evictable = built
    .filter((tab) => tab.id !== activeId)
    .slice()
    .sort((a, b) => (a.lastUsedAt === b.lastUsedAt ? a.id.localeCompare(b.id) : a.lastUsedAt - b.lastUsedAt));
  const overflow = built.length - safeLimit;
  return evictable.slice(0, overflow).map((tab) => tab.id);
}

/** 更新 LRU 记账：标记 built / 记录使用时间；返回新数组（不改原数组，方便 React 侧比较） */
export function markTabUsed(tabs: readonly TabRuntime[], id: string, at: number): TabRuntime[] {
  const exists = tabs.some((tab) => tab.id === id);
  if (!exists) return [...tabs, { id, built: true, lastUsedAt: at }];
  return tabs.map((tab) => (tab.id === id ? { ...tab, built: true, lastUsedAt: at } : tab));
}

/** 冷存后更新记账（built=false，保留 lastUsedAt 供下次淘汰参考） */
export function markTabCold(tabs: readonly TabRuntime[], id: string): TabRuntime[] {
  return tabs.map((tab) => (tab.id === id ? { ...tab, built: false } : tab));
}

/** 关闭标签：从记账里移除 */
export function dropTab(tabs: readonly TabRuntime[], id: string): TabRuntime[] {
  return tabs.filter((tab) => tab.id !== id);
}
