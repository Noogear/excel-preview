/**
 * 把 `RangeSnapshot` 渲染成"紧凑小表格"的缩略图。
 *
 * 设计要点：
 *  - 纯展示、无副作用、无 hooks、不读全局状态：给什么快照就画什么；
 *  - **不做整体缩放**：单元格是固定尺寸（默认 46×20），字号可读、观感与表格一致；
 *  - **一次只显示 N 列**（`columns`，默认 6）：超出部分横向滚动，并在角上提示还有多少列/行；
 *  - `covered`（被合并覆盖的从属单元格）不渲染，位置由左上角单元格的 rowSpan/colSpan 占据。
 */
import type { CSSProperties } from 'react';

import { snapshotStyleToCss } from './preview-style';
import type { RangeSnapshot, SnapshotCell } from './types';
import './workspace.css';

export interface SnapshotPreviewProps {
  snapshot: RangeSnapshot;
  /** 一次显示多少列（默认 6）；超出的列可横向滚动查看 */
  columns?: number;
  /** 单元格宽（默认 46） */
  cellWidth?: number;
  /** 单元格高（默认 20） */
  cellHeight?: number;
  /** 最多显示多少行（默认 10）；超出的行在角上提示 */
  maxRows?: number;
  className?: string;
}

export const DEFAULT_PREVIEW_COLUMNS = 6;
export const DEFAULT_PREVIEW_CELL_WIDTH = 46;
export const DEFAULT_PREVIEW_CELL_HEIGHT = 20;
export const DEFAULT_PREVIEW_MAX_ROWS = 10;
/** 单元格之间的间距（px） */
const CELL_GAP = 1;

/**
 * 结构性基础样式（不属于快照样式，快照样式在其后覆盖）：
 *  - Excel 默认单行显示、溢出裁剪
 *  - 垂直默认居中（快照给了 vAlign 时会被 snapshotStyleToCss 的 verticalAlign 覆盖）
 */
const BASE_CELL_STYLE: CSSProperties = {
  padding: '0 2px',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  verticalAlign: 'middle',
  lineHeight: 1.25,
};

interface PreviewCell {
  key: string;
  text: string;
  style: CSSProperties;
  rowSpan?: number;
  colSpan?: number;
}

interface PreviewRow {
  key: string;
  cells: PreviewCell[];
}

/* -------------------------------------------------------------------------- */
/* 纯计算（组件外，方便阅读与复用）                                             */
/* -------------------------------------------------------------------------- */

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function cls(...parts: (string | undefined | false)[]): string {
  return parts.filter((p): p is string => typeof p === 'string' && p !== '').join(' ');
}

/** 视口尺寸兜底：非正数 / NaN 一律回落默认值 */
function sanitizeBox(value: number | undefined, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : fallback;
}

function toCount(value: number): number {
  return Number.isFinite(value) && value > 0 ? Math.trunc(value) : 0;
}

/** 合并跨度兜底：非法值当 1，并夹在剩余行列内，避免表格被撑破 */
function spanOf(value: number | undefined, limit: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 1;
  const n = Math.trunc(value);
  if (n < 1) return 1;
  return limit > 0 ? Math.min(n, limit) : 1;
}

/** 按行列铺格子：跳过 covered 单元格，给合并区左上角打上 span */
function buildGrid(snapshot: RangeSnapshot, rows: number, cols: number): PreviewRow[] {
  const occupied: boolean[][] = [];
  for (let r = 0; r < rows; r += 1) occupied.push(new Array<boolean>(cols).fill(false));

  const grid: PreviewRow[] = [];
  for (let r = 0; r < rows; r += 1) {
    const cells: PreviewCell[] = [];
    for (let c = 0; c < cols; c += 1) {
      if (occupied[r][c]) continue;

      const cell: SnapshotCell | undefined = snapshot.cells?.[r]?.[c];

      // 被合并覆盖的从属单元格：不渲染（位置由左上角单元格的 span 占据）
      if (cell?.covered === true) {
        occupied[r][c] = true;
        continue;
      }

      const rowSpan = spanOf(cell?.merge?.rowSpan, rows - r);
      const colSpan = spanOf(cell?.merge?.colSpan, cols - c);
      for (let rr = r; rr < r + rowSpan; rr += 1) {
        for (let cc = c; cc < c + colSpan; cc += 1) occupied[rr][cc] = true;
      }

      cells.push({
        key: `${r}-${c}`,
        // 图片 / 数字格式等一律只呈现已格式化好的 text；空文本就是空单元格
        text: typeof cell?.text === 'string' ? cell.text : '',
        // 预览是"概览"而不是保真渲染：即使来源单元格开了自动换行，这里也强制单行 + 省略号。
        // 否则 46px 宽的格子里会变成"一个字一行"，整张卡片被撑成一条竖线（实测见过）。
        style: {
          ...BASE_CELL_STYLE,
          ...snapshotStyleToCss(cell?.style),
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
        },
        rowSpan: rowSpan > 1 ? rowSpan : undefined,
        colSpan: colSpan > 1 ? colSpan : undefined,
      });
    }
    grid.push({ key: `row-${r}`, cells });
  }
  return grid;
}
/* -------------------------------------------------------------------------- */
/* 组件                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * 紧凑的"小表格"预览。
 *
 * 与旧版的区别（用户要求"像表格一样紧凑、留一些间距、可以设置一次显示多少列"）：
 *  - **不再整体缩小**：单元格固定尺寸（默认 46×20），字号可读，观感与表格一致；
 *  - **一次只显示 N 列**（`columns`，默认 6）：超出的部分横向滚动查看，并在角上提示还有多少列/行；
 *  - 单元格之间留 1px 间隙 + 圆角，看起来是"贴在卡片上的小表格"而不是一整块色卡。
 */
export function SnapshotPreview({
  snapshot,
  columns = DEFAULT_PREVIEW_COLUMNS,
  cellWidth = DEFAULT_PREVIEW_CELL_WIDTH,
  cellHeight = DEFAULT_PREVIEW_CELL_HEIGHT,
  maxRows = DEFAULT_PREVIEW_MAX_ROWS,
  className,
}: SnapshotPreviewProps) {
  const cells = Array.isArray(snapshot.cells) ? snapshot.cells : [];
  let maxRowLength = 0;
  for (const row of cells) {
    if (Array.isArray(row) && row.length > maxRowLength) maxRowLength = row.length;
  }
  const rows = Math.max(toCount(snapshot.rows), cells.length);
  const cols = Math.max(toCount(snapshot.cols), maxRowLength);

  const shownCols = Math.max(1, Math.min(toCount(columns), cols || 1));
  const shownRows = Math.max(1, Math.min(toCount(maxRows), rows || 1));
  const hiddenCols = Math.max(0, cols - shownCols);
  const hiddenRows = Math.max(0, rows - shownRows);

  const grid = buildGrid(snapshot, rows, cols);
  const visibleRows = grid.slice(0, shownRows);

  const tableStyle: CSSProperties = {
    borderCollapse: 'separate',
    // 单元格之间留一点间距：既像表格又不至于挤成一片
    borderSpacing: `${CELL_GAP}px`,
    tableLayout: 'fixed',
    width: shownCols * (cellWidth + CELL_GAP) + CELL_GAP,
    backgroundColor: 'var(--ws-sheet-bg, #ffffff)',
    color: 'var(--ws-sheet-fg, #1f2937)',
  };

  return (
    <div
      className={cls('ws-preview', className)}
      data-testid="snapshot-preview"
      data-snapshot-id={snapshot.id}
      data-rows={rows}
      data-cols={cols}
      data-visible-cols={shownCols}
      aria-label={`${snapshot.label}（${snapshot.source.sheetName} ${snapshot.source.a1}）预览`}
    >
      <div className="ws-preview-scroll">
        <table className="ws-preview-table" style={tableStyle}>
          <colgroup>
            {Array.from({ length: shownCols }, (_, c) => (
              <col key={`col-${c}`} style={{ width: cellWidth }} />
            ))}
          </colgroup>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={row.key} style={{ height: cellHeight }}>
                {row.cells
                  // 只渲染落在可见列范围内的单元格（跨列合并的按起点算）
                  .filter((_, index) => index < shownCols)
                  .map((cell) => (
                    <td key={cell.key} rowSpan={cell.rowSpan} colSpan={cell.colSpan} style={cell.style}>
                      {cell.text}
                    </td>
                  ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {hiddenCols > 0 || hiddenRows > 0 ? (
        <span className="ws-preview-more" data-testid="snapshot-preview-more">
          {hiddenCols > 0 ? `+${hiddenCols} 列` : ''}
          {hiddenCols > 0 && hiddenRows > 0 ? ' · ' : ''}
          {hiddenRows > 0 ? `+${hiddenRows} 行` : ''}
        </span>
      ) : null}
    </div>
  );
}