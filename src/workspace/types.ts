/**
 * 工作区（暂存区）与拖拽共用的数据模型。
 *
 * 设计要点：
 *  - 快照只保存"渲染预览"需要的样式 + "写回表格"需要的原始值，两者分开：
 *      cells[][].style → 侧边栏预览用（纯数据，DOM 渲染）
 *      values[][]      → 拖回表格时写回用（**只写值，不写样式**，符合"只编辑内容"约束）
 *  - 快照不持有任何 Univer 对象引用（避免实例销毁后悬挂）
 */

export type SnapshotHAlign = 'left' | 'center' | 'right';
export type SnapshotVAlign = 'top' | 'middle' | 'bottom';

/** 预览所需的最小样式集合（已归一化为 CSS 友好值） */
export interface SnapshotStyle {
  fontFamily?: string;
  /** 单位：pt（与 Excel 一致） */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeThrough?: boolean;
  /** #RRGGBB */
  color?: string;
  /** #RRGGBB */
  fill?: string;
  align?: SnapshotHAlign;
  vAlign?: SnapshotVAlign;
  wrap?: boolean;
  /** 文字旋转角度（度） */
  rotate?: number;
  /** 数字格式 pattern（预览时已由 Univer 格式化为文本，这里仅作展示/回溯用） */
  numberFormat?: string;
  /** 边框：CSS 颜色；线宽由 borderWidth 统一表达 */
  border?: {
    top?: string;
    right?: string;
    bottom?: string;
    left?: string;
  };
}

export interface SnapshotCell {
  /** 已格式化的显示文本（由 Univer 的 displayValue 提供，保证与表格所见一致） */
  text: string;
  /** 原始值（写回用） */
  value?: string | number | boolean | null;
  /** 公式（写回用，不含前导 '='） */
  formula?: string;
  /** true 表示这是被合并覆盖的从属单元格（预览时跳过渲染） */
  covered?: boolean;
  /** 合并信息（仅合并区左上角有值） */
  merge?: { rowSpan: number; colSpan: number };
  style?: SnapshotStyle;
}

/** 区域快照：侧边栏里的一个"卡片" */
export interface RangeSnapshot {
  id: string;
  /** 来源信息（用于展示与写回定位） */
  source: {
    sheetId: string;
    sheetName: string;
    /** A1 记号，如 "B2:D5" */
    a1: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };
  rows: number;
  cols: number;
  cells: SnapshotCell[][];
  /** 写回表格用的值矩阵（只含值/公式，不含样式） */
  values: (string | number | boolean | null)[][];
  /** 写回时用公式还是值：true 表示该位置应写公式 */
  formulas: (string | null)[][];
  createdAt: number;
  /** 卡片标题，默认用 A1 记号 */
  label: string;
  /**
   * 来源单元格在表里的像素尺寸（工作区条目按它展示，做到"保留大小"）。
   * 只有"拆成单格"的条目才有；整片区域的快照没有这个字段。
   */
  cellSize?: { width: number; height: number };
}

/** 拖拽会话的载荷类型 */
export type DragPayload =
  | { kind: 'sheet-range'; snapshot: RangeSnapshot }
  | { kind: 'workspace-item'; snapshot: RangeSnapshot };

/** 落点目标 */
export interface DropTarget {
  row: number;
  col: number;
  sheetId: string;
}

/** 落点判定结果（供 UI 提示与动画使用） */
export type DropDecision =
  | { kind: 'swap'; sourceA1: string; targetA1: string }
  | { kind: 'paste'; targetA1: string }
  | { kind: 'snapshot'; targetA1: string }
  | { kind: 'reject'; reason: string };
