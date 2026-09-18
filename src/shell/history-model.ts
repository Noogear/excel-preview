/**
 * 历史记录模型 —— **纯函数**，零依赖（不 import React / Univer / DOM）。
 *
 * 为什么是"跳步"而不是"跳到"：
 * Univer 的撤销栈是线性的，只有 `undo()` / `redo()` 一步一动，没有"跳到第 N 步"的 API。
 * 所以"还原到任意一步"必须被翻译成一串连续的撤销/重做 —— 这正是 {@link planJump} 干的事：
 * 它只**算账**（要撤销几次、重做几次），既不碰 Univer，也不关心动画与状态。
 *
 * 下标约定（全模块统一，也是面板 `current` / `onJumpTo` 的口径）：
 * - `0` = 最初状态（一条都没应用）；
 * - 下标 `i` = "已经应用了前 i 条历史"；
 * - 因此 `entries.length` = 全部已应用。合法的 target ∈ [0, entries.length]。
 * - 条目数组下标（0-based）与这个下标差 1：条目 `index` 对应的目标是 `index + 1`。
 *
 * 夹紧规则：所有入参都先过 {@link clampIndex} —— 负数、小数、NaN、±Infinity 一律按
 * "0 / 向下取整"处理，**绝不抛异常也绝不返回负计数**（撤销次数为负会让上层循环失控）。
 */

/** 历史条目的分类；用于面板上的分类徽标与颜色语义键。 */
export type HistoryKind = 'edit' | 'swap' | 'workspace' | 'import' | 'mode' | 'export' | 'other';

/** 全部合法分类（顺序即面板/图例的展示顺序，`other` 作为兜底排最后）。 */
export const HISTORY_KINDS: readonly HistoryKind[] = ['edit', 'swap', 'workspace', 'import', 'mode', 'export', 'other'];

/** 颜色语义键：面板只负责把它翻成类名，具体色值由 CSS 决定。 */
export type KindTone = 'neutral' | 'primary' | 'ok' | 'warn';

/** {@link describeKind} 的返回值：中文分类名 + 颜色语义键。 */
export interface KindDescriptor {
  label: string;
  tone: KindTone;
}

/** 一条历史记录。`id` 是去重键（同 id 视为同一步的更新）。 */
export interface HistoryEntry {
  id: string;
  /** 中文可读标签，如"互换 A3 ⇄ A4" */
  label: string;
  kind: HistoryKind;
  /** 时间戳（ms，本地时区展示） */
  at: number;
  /**
   * 这步动作改的是**表格**还是**工作区**（缺省视为表格）。
   *
   * 撤销/重做要按时间顺序在两者之间切换：工作区动作由我们自己的栈回放，
   * 表格动作交给 Univer 的撤销栈（见 `App.tsx` 的 `stepBack/stepForward`）。
   */
  scope?: 'sheet' | 'workspace';
}

/** 跳步计划：从 current 走到 target 需要撤销/重做多少次。 */
export interface JumpPlan {
  undo: number;
  redo: number;
}

/** Univer 当前可用的撤销/重做次数（来自 `undoRedoStatus$`）。 */
export interface UndoRedoAvailability {
  undos: number;
  redos: number;
}

const KIND_DESCRIPTORS: Readonly<Record<HistoryKind, KindDescriptor>> = {
  edit: { label: '编辑', tone: 'neutral' },
  swap: { label: '互换', tone: 'primary' },
  workspace: { label: '工作区', tone: 'ok' },
  import: { label: '导入', tone: 'primary' },
  mode: { label: '模式', tone: 'warn' },
  export: { label: '导出', tone: 'ok' },
  other: { label: '其它', tone: 'neutral' },
};

const TONES: readonly KindTone[] = ['neutral', 'primary', 'ok', 'warn'];

const FALLBACK_KIND: KindDescriptor = { label: '其它', tone: 'neutral' };

/**
 * 把任意数值夹紧成"非负整数下标/计数"：
 * 负数、小数、NaN、±Infinity、非 number 一律 → 0（小数向下取整）。
 */
export function clampIndex(value: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) return 0;
  return Math.floor(value);
}

/**
 * 跳步计划：从 `current` 走到 `target` 要撤销/重做多少次。
 *
 * - `target < current` → 只撤销（`undo = current - target`）；
 * - `target > current` → 只重做（`redo = target - current`）；
 * - `target === current` → 原地不动（两个都是 0，面板借此把该项当作"当前步"）。
 *
 * 结果里的 undo / redo **不会同时非零**（线性栈的走法唯一）。
 * 两个入参都先按 {@link clampIndex} 夹紧，所以 `planJump(3, 1) => { undo: 2, redo: 0 }`、
 * `planJump(1, 3) => { undo: 0, redo: 2 }`、`planJump(2, 2) => { undo: 0, redo: 0 }`。
 */
export function planJump(current: number, target: number): JumpPlan {
  const from = clampIndex(current);
  const to = clampIndex(target);
  if (to < from) return { undo: from - to, redo: 0 };
  return { undo: 0, redo: to - from };
}

/**
 * 这个跳步计划在当前可用次数下是否走得通。
 *
 * 判定就是字面意思：`plan.undo <= available.undos && plan.redo <= available.redos`。
 * 可用次数同样先夹紧（负数/NaN 视为 0），避免"可用 -1 次"这种脏数据把判定弄反。
 * 注意 `{ undo: 0, redo: 0 }` 永远可行（不需要任何撤销/重做）。
 */
export function canJump(plan: JumpPlan, available: UndoRedoAvailability): boolean {
  return clampIndex(plan.undo) <= clampIndex(available.undos) && clampIndex(plan.redo) <= clampIndex(available.redos);
}

/**
 * 不可达时的中文原因（可达返回 `null`），直接拿去当 `title` / `aria-description`。
 *
 * 例：`jumpBlockReason({ undo: 0, redo: 3 }, { undos: 5, redos: 1 })`
 * → `'需要重做 3 次，但只有 1 次可用'`。
 */
export function jumpBlockReason(plan: JumpPlan, available: UndoRedoAvailability): string | null {
  if (canJump(plan, available)) return null;

  const needUndo = clampIndex(plan.undo);
  const needRedo = clampIndex(plan.redo);
  const haveUndo = clampIndex(available.undos);
  const haveRedo = clampIndex(available.redos);

  // undo / redo 不会同时非零，先撤销后重做只是固定的报告顺序
  if (needUndo > haveUndo) return `需要撤销 ${needUndo} 次，但只有 ${haveUndo} 次可用`;
  if (needRedo > haveRedo) return `需要重做 ${needRedo} 次，但只有 ${haveRedo} 次可用`;
  return '当前无法跳转到这一步';
}

/**
 * 追加一条历史（**不改入参**，总是返回新数组）。
 *
 * - 去重：已存在同 `id` 的条目时**原地更新**（保持它在列表中的位置，不重复追加）——
 *   同一步的标签/时间被改写时不该在历史里裂成两条；
 * - 截断：`limit` 只保留**最新**的 limit 条（丢弃最旧的头部）。`limit` 为
 *   `undefined` / 非有限 / `<= 0` 时不截断；小数向下取整（`2.9 → 2`）。
 */
export function appendEntry(entries: HistoryEntry[], entry: HistoryEntry, limit?: number): HistoryEntry[] {
  const source = Array.isArray(entries) ? entries : [];
  const exists = source.some((item) => item.id === entry.id);
  const next = exists ? source.map((item) => (item.id === entry.id ? entry : item)) : [...source, entry];

  const max = typeof limit === 'number' && Number.isFinite(limit) ? Math.floor(limit) : 0;
  if (max <= 0 || next.length <= max) return next;
  return next.slice(next.length - max);
}

/** 两位补零；只用于 {@link formatTime} 内部的合法 0–59 数值。 */
function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/**
 * 把时间戳格式化成 `HH:MM:SS`（**本地时区**，用 `getHours/getMinutes/getSeconds`）。
 *
 * 纯函数、不依赖 Intl / 系统区域设置，因此单测里用
 * `new Date(2024, 0, 2, 9, 5, 7).getTime()` 构造即可得到与运行时区无关的期望值。
 * 无法格式化的入参（NaN / ±Infinity / 超出 Date 范围）返回占位符 `'--:--:--'`。
 */
export function formatTime(at: number): string {
  if (typeof at !== 'number' || !Number.isFinite(at)) return '--:--:--';
  const date = new Date(at);
  if (!Number.isFinite(date.getTime())) return '--:--:--';
  return `${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(date.getSeconds())}`;
}

function isTone(value: unknown): value is KindTone {
  return typeof value === 'string' && (TONES as readonly string[]).includes(value);
}

function isDescriptor(value: unknown): value is KindDescriptor {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as { label?: unknown; tone?: unknown };
  return typeof candidate.label === 'string' && isTone(candidate.tone);
}

/**
 * 分类 → 中文分类名 + 颜色语义键（面板据此渲染徽标）。
 *
 * 对运行时脏值（持久化数据、JS 调用方塞进来的未知字符串，甚至 `toString` 这类原型属性）
 * 一律回落到"其它 / neutral"，绝不返回 undefined 或半截对象。
 * 每次返回**新对象**，调用方随便改也污染不到内部表。
 */
export function describeKind(kind: HistoryKind): KindDescriptor {
  const found: unknown = (KIND_DESCRIPTORS as Record<string, unknown>)[String(kind)];
  if (isDescriptor(found)) return { label: found.label, tone: found.tone };
  return { label: FALLBACK_KIND.label, tone: FALLBACK_KIND.tone };
}
