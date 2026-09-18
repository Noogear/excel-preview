/**
 * 工作区侧边栏（暂存区）。
 *
 * 职责边界：**只负责"呈现 + 抛事件"**（组件内**不使用 hooks**，因此可以在 node 下直接调用做契约测试）。
 *  - 拖拽回表格的手势判定完全在上层手势层，这里拿到 `pointerdown` 原样转发，
 *    不做 `preventDefault`（也不要 setPointerCapture）——上层需要完整事件与坐标；
 *  - 条目右键菜单由上层用共享的 `ContextMenu` 组件渲染（这里只把事件与条目抛出去）。
 *
 * 本轮改造（用户要求）：
 *  - 「+」从顶部工具栏**挪到这里**的右上角；
 *  - 条目 = **一个个独立单元格**（按来源尺寸展示，一格一个方块，每行几格随面板宽度自适应）；
 *  - 顶部**搜索 + 来源筛选**：按内容搜、按"来自哪张表"筛（状态在上层，面板保持无 hooks）；
 *  - 底部一排选项：写回后是否移除条目 + 格子最小宽度；
 *  - 条目上右键可**复制 / 剪切 / 删除**（复制/剪切进内部剪贴板，可在表格里粘贴）。
 */
import type { MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent } from 'react';

import { TILE_MIN_WIDTH_RANGE, type SessionSettings } from '../persistence/session';
import { SnapshotPreview } from './SnapshotPreview';
import { PlusIcon } from '../shell/icons';
import { ALL_SOURCES, filterItems, isFilterActive, itemText, listSourceSheets, type WorkspaceFilter } from './filter';
import type { RangeSnapshot } from './types';
import './workspace.css';

export interface WorkspacePanelProps {
  items: RangeSnapshot[];
  settings: SessionSettings;
  /** 搜索 / 来源筛选状态（放在上层，面板本身保持无 hooks，便于契约测试直接调用） */
  filter: WorkspaceFilter;
  onFilterChange: (patch: Partial<WorkspaceFilter>) => void;
  onRemove: (id: string) => void;
  onClearAll: () => void;
  /** 用户按住某个条目想拖回表格：由上层手势层接管（这里只把事件与条目抛出去） */
  onItemPointerDown: (event: ReactPointerEvent<HTMLElement>, item: RangeSnapshot) => void;
  /** 条目上右键：由上层用共享 ContextMenu 渲染「复制/剪切/删除」 */
  onItemContextMenu: (event: ReactMouseEvent<HTMLElement>, item: RangeSnapshot) => void;
  /** 点击右上角「+」：从当前表快速导入（单输入框自动识别行/列/区域） */
  onAddItem: () => void;
  onSettingsChange: (patch: Partial<SessionSettings>) => void;
  /**
   * 二次确认（清空是不可逆的批量操作）。
   * 状态由上层持有（面板保持无 hooks）；这几个 prop 刻意是**可选**的：
   * 不传时点「清空」直接清空（便于独立渲染/契约测试），生产路径由 App 传入。
   */
  clearConfirming?: boolean;
  onRequestClear?: () => void;
  onCancelClear?: () => void;
  /**
   * 批量移除（"只清空筛选出来的那部分"用）。
   *
   * 必须是**一次原子更新**：早先这里是 `for (const item of visible) onRemove(item.id)`，
   * 一次事件回调里连打 N 次 setState，既多出 N 次渲染、也容易踩"同一份 prev 被覆盖"的坑。
   * 没传时退回逐条调用（独立渲染/契约测试仍可用）。
   */
  onRemoveMany?: (ids: string[]) => void;
  /**
   * 「点击互换」模式下已经点过、正在等待第二次点击的那一条（用高亮告诉用户"我选中它了，
   * 接着点表格里的格子即可互换"）。可选：不传就不显示该状态。
   */
  pendingItemId?: string | null;
  className?: string;
}

function cls(...parts: (string | undefined | false)[]): string {
  return parts.filter((p): p is string => typeof p === 'string' && p !== '').join(' ');
}

/** 把"来源单元格尺寸"夹到可读区间；拿不到就用默认值 */
function clampSize(value: number | undefined, min: number, max: number, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

export function WorkspacePanel({
  items,
  settings,
  filter,
  onFilterChange,
  onRemove,
  onClearAll,
  onItemPointerDown,
  onItemContextMenu,
  onAddItem,
  onSettingsChange,
  clearConfirming = false,
  onRequestClear,
  onCancelClear,
  onRemoveMany,
  pendingItemId = null,
  className,
}: WorkspacePanelProps) {
  const hasItems = items.length > 0;
  const sources = listSourceSheets(items);
  const visible = filterItems(items, filter);
  const filtering = isFilterActive(filter);

  return (
    <aside className={cls('ws-panel', className)} aria-label="工作区暂存区">
      <header className="ws-panel-head">
        <h2 className="ws-panel-title">工作区</h2>
        {hasItems ? <span className="ws-panel-count" data-testid="workspace-count">{items.length}</span> : null}
        <span className="ws-panel-spacer" />
        <button
          type="button"
          className="ws-add"
          data-testid="workspace-add"
          onClick={onAddItem}
          title="从当前表导入（行 / 列 / 区域）"
          aria-label="从当前表导入到工作区"
        >
          <PlusIcon />
        </button>
        {hasItems ? (
          <button
            type="button"
            className="ws-clear"
            data-testid="workspace-clear"
            onClick={onRequestClear ?? onClearAll}
            title="清空工作区（需要二次确认）"
          >
            清空
          </button>
        ) : null}
      </header>

      {/* 二次确认条：清空是不可逆的批量操作，必须问一次；筛选生效时还能只清"当前显示的这些" */}
      {hasItems && clearConfirming ? (
        <div className="ws-confirm" data-testid="workspace-clear-confirm" role="alertdialog" aria-label="确认清空">
          <span className="ws-confirm-text">
            {filtering ? `清空显示的 ${visible.length} 个，还是全部 ${items.length} 个？` : `确定清空全部 ${items.length} 个单元格？`}
          </span>
          <span className="ws-confirm-actions">
            {filtering ? (
              <button
                type="button"
                className="ws-confirm-btn is-primary"
                data-testid="workspace-clear-visible"
                disabled={visible.length === 0}
                onClick={() => {
                  const ids = visible.map((item) => item.id);
                  if (onRemoveMany) onRemoveMany(ids);
                  else for (const id of ids) onRemove(id);
                  onCancelClear?.();
                }}
              >
                清空显示的 {visible.length} 个
              </button>
            ) : null}
            <button
              type="button"
              className={cls('ws-confirm-btn', filtering ? 'is-danger' : 'is-primary')}
              data-testid="workspace-clear-all"
              onClick={() => {
                onClearAll();
                onCancelClear?.();
              }}
            >
              清空全部 {items.length} 个
            </button>
            <button
              type="button"
              className="ws-confirm-btn"
              data-testid="workspace-clear-cancel"
              onClick={() => onCancelClear?.()}
            >
              取消
            </button>
          </span>
        </div>
      ) : null}

      {hasItems ? (
        <div className="ws-filter" data-testid="workspace-filter">
          <input
            type="search"
            className="ws-search"
            data-testid="workspace-search"
            placeholder="搜索内容 / 单元格（如 B12）"
            value={filter.query}
            aria-label="在工作区里搜索内容"
            onChange={(event) => onFilterChange({ query: event.target.value })}
            onKeyDown={(event) => {
              if (event.key === 'Escape') {
                event.stopPropagation();
                onFilterChange({ query: '' });
              }
            }}
          />
          <select
            className="ws-source"
            data-testid="workspace-source-filter"
            aria-label="按来源工作表筛选"
            value={filter.source}
            onChange={(event) => onFilterChange({ source: event.target.value })}
          >
            <option value={ALL_SOURCES}>全部来源（{items.length}）</option>
            {sources.map((name) => (
              <option key={name} value={name}>
                {name}（{items.filter((item) => item.source.sheetName === name).length}）
              </option>
            ))}
          </select>
          {filtering ? (
            <button
              type="button"
              className="ws-filter-clear"
              data-testid="workspace-filter-clear"
              title="清除搜索与筛选"
              onClick={() => onFilterChange({ query: '', source: ALL_SOURCES })}
            >
              {visible.length}/{items.length} · 清除
            </button>
          ) : null}
        </div>
      ) : null}

      {hasItems ? (
        <ul
          className="ws-list"
          // "工作区相当于一张小表格"：每行放几格 = 面板宽度 ÷ 格子最小宽度，**自动**决定
          //（拉宽面板就自动多放几格，不需要用户改任何设置；默认宽度下正好 2 列）
          style={{ gridTemplateColumns: `repeat(auto-fill, minmax(${settings.tileMinWidth}px, 1fr))` }}
          data-tile-min-width={settings.tileMinWidth}
          data-visible-count={visible.length}
        >
          {visible.map((item) => {
            const source = item.source;
            const text = itemText(item);
            return (
              <li
                key={item.id}
                className={cls('ws-item', item.id === pendingItemId && 'is-pending-swap')}
                data-testid="workspace-item"
                data-snapshot-id={item.id}
                data-pending-swap={item.id === pendingItemId ? '1' : undefined}
                title={`${text || '(空)'} · 来源 ${source.sheetName}!${source.a1}（右键可复制/剪切/删除）`}
                onPointerDown={(event) => onItemPointerDown(event, item)}
                onContextMenu={(event) => onItemContextMenu(event, item)}
              >
                {/* 单元格本体：按来源尺寸展示（"保留大小"），并尽量宽到能看清文字 */}
                <div className="ws-item-preview">
                  <SnapshotPreview
                    snapshot={item}
                    cellWidth={clampSize(item.cellSize?.width, settings.tileMinWidth - 16, 200, 104)}
                    cellHeight={clampSize(item.cellSize?.height, 18, 64, 24)}
                  />
                </div>
                {/* 刻意不放角标 ×：格子上只有单元格本身，删除走右键菜单（复制/剪切/删除） */}
              </li>
            );
          })}
          {visible.length === 0 ? (
            <li className="ws-no-match" data-testid="workspace-no-match">
              没有匹配的单元格（共 {items.length} 个）
            </li>
          ) : null}
        </ul>
      ) : (
        <div className="ws-empty" data-testid="workspace-empty">
          <span className="ws-empty-icon" aria-hidden="true">
            ⤓
          </span>
          <span className="ws-empty-text">拖拽表格中的单元格到此处暂存</span>
          <button type="button" className="ws-empty-add" data-testid="workspace-empty-add" onClick={onAddItem}>
            或点这里从当前表导入
          </button>
        </div>
      )}

      {/* 底部选项：与表格的交互语义 */}
      <footer className="ws-panel-foot">
        <label className="ws-opt" title="勾选后，把条目写回表格时保留该条目；取消则写回后从工作区移除">
          <input
            type="checkbox"
            data-testid="workspace-remove-after-paste"
            checked={settings.removeItemAfterPaste}
            onChange={(event) => onSettingsChange({ removeItemAfterPaste: event.target.checked })}
          />
          <span>写回表格后移除条目</span>
        </label>
        <label
          className="ws-opt ws-opt-cols"
          title="格子的最小宽度：调小则每行放更多格，调大则格子更宽、文字更完整（每行几列由面板宽度自动决定）"
        >
          <span>格宽</span>
          <input
            type="number"
            data-testid="workspace-tile-width"
            min={TILE_MIN_WIDTH_RANGE.min}
            max={TILE_MIN_WIDTH_RANGE.max}
            step={4}
            value={settings.tileMinWidth}
            onChange={(event) => onSettingsChange({ tileMinWidth: Number(event.target.value) })}
          />
          <span>px</span>
        </label>
      </footer>
    </aside>
  );
}
