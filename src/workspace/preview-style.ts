/**
 * 工作区预览的**纯样式换算**（无 DOM、无副作用、无 React 运行时依赖）。
 *
 * 之所以单独成文件：
 *  - 便于单测（node 环境下直接调用，不需要 jsdom）
 *  - 预览组件只负责"摆放"，所有换算规则集中在这里，改一处即可
 *
 * 数据来源见 `./types`：快照只存 CSS 友好的纯数据，不持有任何 Univer 对象。
 */
import type { CSSProperties } from 'react';
import type { SnapshotStyle } from './types';

/* -------------------------------------------------------------------------- */
/* 常量                                                                        */
/* -------------------------------------------------------------------------- */

/** CSS 绝对长度：1in = 96px = 72pt */
const PX_PER_PT = 96 / 72;

/** 快照不带行列尺寸时的默认单元格尺寸（预览缩略图的口径） */
export const DEFAULT_COL_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 22;

/** 边框统一线宽：缩略图里太细会糊，太粗会脏，固定 1px */
const BORDER_WIDTH = '1px';

/* -------------------------------------------------------------------------- */
/* 小工具                                                                      */
/* -------------------------------------------------------------------------- */

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/** 只接受"能当尺寸用"的正有限数；其它（undefined / NaN / 0 / 负数）一律视为没提供 */
function isPositiveFinite(n: unknown): n is number {
  return typeof n === 'number' && Number.isFinite(n) && n > 0;
}

/** 非空字符串才算有效 CSS 值（避免输出 `undefinedpx` / `1px solid ` 这类垃圾） */
function isCssValue(v: unknown): v is string {
  return typeof v === 'string' && v.trim() !== '';
}

/** 查表取值：key 不在表里（运行期脏数据）时返回 undefined，不抛错、不输出无效 CSS */
function lookup<K extends string, V>(map: Record<K, V>, key: K): V | undefined {
  return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : undefined;
}

/** 水平对齐 → CSS textAlign（与快照枚举同名，仍走查表以便挡住脏数据） */
const H_ALIGN_TO_TEXT: Record<'left' | 'center' | 'right', 'left' | 'center' | 'right'> = {
  left: 'left',
  center: 'center',
  right: 'right',
};

/** 垂直对齐 → CSS vertical-align（td 原生支持，语义完全一致） */
const V_ALIGN_TO_VERTICAL: Record<'top' | 'middle' | 'bottom', 'top' | 'middle' | 'bottom'> = {
  top: 'top',
  middle: 'middle',
  bottom: 'bottom',
};

/* -------------------------------------------------------------------------- */
/* SnapshotStyle → React.CSSProperties                                         */
/* -------------------------------------------------------------------------- */

/**
 * 把快照样式转成内联 CSS。
 *
 * @param style 快照单元格样式（可能为 undefined，此时返回空对象）
 * @param scale 字号缩放系数（默认 1）。**只作用于字号**：快照里的 fontSize 单位是 pt
 *              （与 Excel 一致），换算成 px 后再乘以该系数。
 *              注意：`SnapshotPreview` 用 CSS `transform: scale()` 做整体缩放，
 *              因此它**不传** scale（传了会变成双重缩放）；这个参数是给
 *              "不做 transform、直接把尺寸算进样式"的调用方（导出/打印/固定尺寸卡片）用的。
 *
 * 约定：**未提供的字段不写入返回对象**（不会出现 `undefined` 值 / `NaN` / `undefinedpx`）。
 *
 * 对齐映射（最终方案，理由见正文注释）：
 *  - `align`  → `textAlign`（'left' | 'center' | 'right' 与快照枚举同名）
 *  - `vAlign` → `verticalAlign`（'top' | 'middle' | 'bottom' 与快照枚举同名）
 *  两者都是 table-cell 的原生能力，可同时输出、互不干扰；
 *  **没有**采用 `display:flex + alignItems`：td 一旦变成 flex 容器就不再是 table-cell，
 *  CSS 表格修复会把同一行的多个 flex 单元格塞进同一个匿名 table-cell，整行布局会塌成一列（实测）。
 */
export function snapshotStyleToCss(style: SnapshotStyle | undefined, scale = 1): CSSProperties {
  const css: CSSProperties = {};
  if (!style) return css;

  /* 字体族 --------------------------------------------------------------- */
  if (isCssValue(style.fontFamily)) css.fontFamily = style.fontFamily;

  /* 字号：pt → px → ×scale ---------------------------------------------- */
  if (isPositiveFinite(style.fontSize)) {
    const k = isPositiveFinite(scale) ? scale : 1;
    const px = round2(style.fontSize * PX_PER_PT * k);
    if (isPositiveFinite(px)) css.fontSize = `${px}px`;
  }

  /* 字形 ----------------------------------------------------------------- */
  if (style.bold === true) css.fontWeight = 'bold';
  if (style.italic === true) css.fontStyle = 'italic';

  const decorations: string[] = [];
  if (style.underline === true) decorations.push('underline');
  if (style.strikeThrough === true) decorations.push('line-through');
  if (decorations.length > 0) css.textDecoration = decorations.join(' ');

  /* 颜色 ----------------------------------------------------------------- */
  if (isCssValue(style.color)) css.color = style.color.trim();
  if (isCssValue(style.fill)) css.background = style.fill.trim();

  /* 对齐 ----------------------------------------------------------------- */
  // 方案选择：**水平 textAlign + 垂直 vertical-align**（不用 display:flex）。
  // 理由（实测，不是猜的）：预览是 <table>，给 <td> 加 display:flex 会让它不再是 table-cell，
  // CSS 表格修复规则会把同一行里的多个 flex 兄弟包进**同一个匿名 table-cell**，
  // 结果是这一行的单元格全挤进第一列（Chromium 截图里 "姓名/分数" 直接上下叠在一起、列宽全废）。
  // 而 textAlign / vertical-align 本来就是 table-cell 的原生能力，两者互不冲突、也无副作用，
  // 所以 align 与 vAlign 可以同时输出，不需要"二选一"的妥协。
  const textAlign = style.align ? lookup(H_ALIGN_TO_TEXT, style.align) : undefined;
  const verticalAlign = style.vAlign ? lookup(V_ALIGN_TO_VERTICAL, style.vAlign) : undefined;
  if (textAlign) css.textAlign = textAlign;
  if (verticalAlign) css.verticalAlign = verticalAlign;

  /* 自动换行 ------------------------------------------------------------- */
  if (style.wrap === true) {
    css.whiteSpace = 'pre-wrap';
    css.wordBreak = 'break-all';
  } else if (style.wrap === false) {
    // 明确"不换行"时给 nowrap（Excel 默认单行溢出裁剪），undefined 则不输出
    css.whiteSpace = 'nowrap';
  }

  /* 旋转：0 与 undefined 都不输出 transform ------------------------------ */
  if (typeof style.rotate === 'number' && Number.isFinite(style.rotate) && style.rotate !== 0) {
    css.transform = `rotate(${round2(style.rotate)}deg)`;
  }

  /* 边框：四边各自独立，缺边不输出 --------------------------------------- */
  const border = style.border;
  if (border) {
    const side = (color: string | undefined): string | undefined =>
      isCssValue(color) ? `${BORDER_WIDTH} solid ${color.trim()}` : undefined;

    const top = side(border.top);
    const right = side(border.right);
    const bottom = side(border.bottom);
    const left = side(border.left);
    if (top) css.borderTop = top;
    if (right) css.borderRight = right;
    if (bottom) css.borderBottom = bottom;
    if (left) css.borderLeft = left;
  }

  return css;
}

/* -------------------------------------------------------------------------- */
