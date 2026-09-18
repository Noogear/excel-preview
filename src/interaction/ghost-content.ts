/**
 * 把快照渲染成一个独立 DOM 元素（用于拖拽"幽灵"）。
 * 与侧边栏预览共用同一个样式映射函数，保证"拖起来的样子"和"卡片里的样子"一致。
 */
import { snapshotStyleToCss } from '../workspace/preview-style';
import type { RangeSnapshot } from '../workspace/types';

export interface GhostContentOptions {
  /** 单元格缩放（幽灵一般比卡片大一点，1 表示按 pt→px 原尺寸） */
  scale?: number;
  /** 最多渲染多少行/列，避免超大选区拖出一个巨物 */
  maxRows?: number;
  maxCols?: number;
}

export function renderSnapshotToElement(snapshot: RangeSnapshot, options: GhostContentOptions = {}): HTMLElement {
  const scale = options.scale ?? 1;
  const maxRows = options.maxRows ?? 12;
  const maxCols = options.maxCols ?? 8;

  const wrapper = document.createElement('div');
  wrapper.className = 'ghost-content';
  wrapper.style.cssText = [
    'position:relative',
    'padding:4px 6px',
    'background:#fff',
    'border:1px solid rgba(37,99,235,.55)',
    'border-radius:8px',
    'box-shadow:0 10px 26px rgba(15,23,42,.22)',
    `transform: scale(${scale})`,
    'transform-origin: top left',
    'pointer-events:none',
  ].join(';');

  const table = document.createElement('table');
  table.style.cssText = 'border-collapse:collapse;table-layout:fixed;font-family:"Microsoft YaHei",system-ui,sans-serif';

  const rows = Math.min(snapshot.rows, maxRows);
  const cols = Math.min(snapshot.cols, maxCols);

  for (let r = 0; r < rows; r++) {
    const tr = document.createElement('tr');
    for (let c = 0; c < cols; c++) {
      const cell = snapshot.cells[r]?.[c];
      if (cell?.covered) continue;

      const td = document.createElement('td');
      const css = snapshotStyleToCss(cell?.style);
      Object.assign(td.style, css, {
        minWidth: '38px',
        maxWidth: '120px',
        padding: '1px 4px',
        fontSize: '12px',
        lineHeight: '16px',
        whiteSpace: 'nowrap',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        border: '1px solid #e5e7eb',
        background: (css.background as string | undefined) ?? '#fff',
      });
      td.textContent = cell?.text ?? '';
      if (cell?.merge) {
        if (cell.merge.colSpan > 1) td.colSpan = Math.min(cell.merge.colSpan, cols - c);
        if (cell.merge.rowSpan > 1) td.rowSpan = Math.min(cell.merge.rowSpan, rows - r);
      }
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  wrapper.appendChild(table);

  if (snapshot.rows > rows || snapshot.cols > cols) {
    const badge = document.createElement('div');
    badge.textContent = `${snapshot.rows}×${snapshot.cols}`;
    badge.style.cssText = 'position:absolute;right:6px;bottom:2px;font-size:10px;color:#6b7280';
    wrapper.appendChild(badge);
  }

  return wrapper;
}
