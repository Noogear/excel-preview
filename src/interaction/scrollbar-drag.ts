/**
 * 自己实现的滚动条拖拽（**按住即可左右/上下拖动，滑块跟手**）。
 *
 * 为什么要自己实现：Univer 引擎自带的滚动条拖拽**不响应合成指针事件**（CDP 与自定义 `PointerEvent`
 * 都试过：按下不跳、移动不动），既无法在 CI 验证、也难以保证跨浏览器一致。
 *
 * 交互模型（就是普通滚动条的样子，用户要求"按住然后可以左右或者上下拖动"）：
 *  - 滑块长度 = 视口 / 内容 × 轨道长；滑块位置 = 当前滚动 / 上限 × 可移动距离；
 *  - 按住**滑块**：记下"抓取偏移"（指针在滑块内的相对位置），拖动时滑块**严格跟手**，不跳；
 *  - 按住**轨道**：先把滑块移到指针处（居中），随后继续拖动同样跟手；
 *  - 全程按"指针位置 → 滑块位置 → 滚动量"换算，指针拖出滚动条也不影响（自然夹在两端）。
 *
 * 这里只放纯计算（可单测）。
 */

export interface AxisGeometry {
  /** 轨道长度（画布 px） */
  trackLength: number;
  /** 滑块长度（画布 px） */
  thumbLength: number;
  /** 最大滚动量（内容坐标单位，来自 `bar.limitX/limitY`） */
  limit: number;
  /** 轨道起点（画布 px）：把指针位置换算成"轨道内位置"用 */
  trackOrigin: number;
  /** 兼容字段：滑块起始位置（留给调用方按需填，计算不用它） */
  thumbStart: number;
}

const clamp = (value: number, min: number, max: number): number => Math.min(Math.max(value, min), max);

/** 滑块能移动的距离（轨道长 − 滑块长）；防 0 除 */
export function thumbTravel(geo: AxisGeometry): number {
  return Math.max(1, geo.trackLength - geo.thumbLength);
}

/** 当前滚动量 → 滑块在轨道里的起点 */
export function thumbStartForScroll(geo: AxisGeometry, scroll: number): number {
  if (geo.limit <= 0) return 0;
  return clamp((scroll / geo.limit) * thumbTravel(geo), 0, thumbTravel(geo));
}

/** 滑块起点 → 滚动量（与上面互逆） */
export function scrollForThumbStart(geo: AxisGeometry, thumbStart: number): number {
  if (geo.limit <= 0) return 0;
  return clamp((clamp(thumbStart, 0, thumbTravel(geo)) / thumbTravel(geo)) * geo.limit, 0, geo.limit);
}

/** 指针在轨道内的位置（画布 px → 相对轨道起点） */
export function pointerInTrack(geo: AxisGeometry, pointer: number): number {
  return pointer - geo.trackOrigin;
}

/**
 * 按下：决定这次是"抓滑块"还是"点轨道"。
 *
 * @param tolerance 命中滑块的容差（px）：滑块只有几像素宽，偏一点也算抓着它
 * @returns `grabOffset` = 指针相对滑块起点的偏移；`scroll` = 按下后应立即生效的滚动量
 *          （抓滑块时不变；点轨道时先跳到指针处，之后继续拖）
 */
export function beginScrollDrag(
  geo: AxisGeometry,
  pointer: number,
  currentScroll: number,
  tolerance = 4,
): { grabOffset: number; scroll: number; onThumb: boolean } {
  const inTrack = pointerInTrack(geo, pointer);
  const thumbStart = thumbStartForScroll(geo, currentScroll);
  const onThumb = inTrack >= thumbStart - tolerance && inTrack <= thumbStart + geo.thumbLength + tolerance;
  if (onThumb) {
    return { grabOffset: inTrack - thumbStart, scroll: currentScroll, onThumb: true };
  }
  // 点轨道：滑块中心对齐指针，随后从这个偏移继续拖
  const grabOffset = geo.thumbLength / 2;
  return { grabOffset, scroll: scrollForThumbStart(geo, inTrack - grabOffset), onThumb: false };
}

/** 拖动中：指针位置 → 目标滚动量（滑块跟手） */
export function scrollForPointer(geo: AxisGeometry, pointer: number, grabOffset: number): number {
  return scrollForThumbStart(geo, pointerInTrack(geo, pointer) - grabOffset);
}

/** 点轨道跳转（不拖）时的目标滚动量：滑块中心落在指针处 */
export function scrollFromTrackClick(geo: AxisGeometry, pointer: number): number {
  return scrollForThumbStart(geo, pointerInTrack(geo, pointer) - geo.thumbLength / 2);
}

/** 指针是否落在滑块上（含容差） */
export function isOnThumb(geo: AxisGeometry, pointer: number, currentScroll = 0, tolerance = 4): boolean {
  const inTrack = pointerInTrack(geo, pointer);
  const thumbStart = thumbStartForScroll(geo, currentScroll);
  return inTrack >= thumbStart - tolerance && inTrack <= thumbStart + thumbLengthSafe(geo) + tolerance;
}

/** 滑块长度（非法值兜底成 12px，避免"什么都抓不住"） */
export function thumbLengthSafe(geo: AxisGeometry): number {
  return Number.isFinite(geo.thumbLength) && geo.thumbLength > 0 ? geo.thumbLength : 12;
}

/**
 * 指针是不是落在**滚动条的交互条带**里。
 *
 * 实测（真实夹具）：主画布 1110px 宽，而主视口只有 933px——竖向滚动条画在视口右缘，
 * **只有 5~6px 宽**，右边还有 170 多像素"死区"（那里没有单元格）。
 * 早先要求"指针精确压住滚动条（靠引擎 `pick` 命中）"，偏一两像素就毫无反应（用户反馈"完全拖不动"）。
 * 现在按**条带**判定：从"视口右缘往里 24px"一直到"画布右缘"都算滚动条区域；死区本来没内容，
 * 划进来只赚不亏。
 *
 * @param position 指针在**画布坐标**里沿该轴的位置（竖向用 x，横向用 y）
 * @param viewportSize 主视口在该轴上的尺寸（画布坐标）
 * @param canvasSize 画布在该轴上的尺寸
 * @param thickness 条带从视口边缘往里算的厚度
 */
export function isInBarStrip(position: number, viewportSize: number, canvasSize: number, thickness = 24): boolean {
  if (![position, viewportSize, canvasSize].every((value) => Number.isFinite(value))) return false;
  if (position < 0) return false;
  return position >= viewportSize - thickness && position <= canvasSize + 2;
}
