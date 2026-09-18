/**
 * 顶部工具栏（外壳组件，**纯展示**）。
 *
 * 设计约束：
 * - **无 hooks、无副作用**：组件不持有状态、不注册监听。"打开文件"用 `<label>` 包裹隐藏的
 *   `<input type="file">`，靠原生 label→input 关联触发选择框，因此不需要 `useRef`。
 * - **没有品牌区 / 大标题**：最左侧直接就是图标按钮组（用户明确要求去掉"Excel 高保真预览"这类标题）。
 * - 每个按钮 `title` + `aria-label` 双份中文提示（tooltip 给鼠标，aria-label 给读屏）；
 *   所有可点元素都带 `data-testid`，供 e2e 断言。
 * - 高度固定 48px、单行不换行：标签条 `flex:1 + min-width:0 + overflow-x:auto` 吃掉剩余宽度，
 *   其余分组 `flex:0 0 auto`，"+"按钮被挤到最右侧（贴近工作区侧栏那一侧）。
 * - 样式见 `toolbar.css`：全部复用 shell.css 的设计令牌，深色模式自动跟随系统。
 */
import type { ChangeEvent } from 'react';

import {
  CloseIcon,
  DragModeIcon,
  ExportIcon,
  HistoryIcon,
  LogIcon,
  OpenIcon,
  RedoIcon,
  SelectModeIcon,
  SwapModeIcon,
  UndoIcon,
} from './icons';
import {
  INTERACTION_MODE_HINTS,
  INTERACTION_MODE_LABELS,
  INTERACTION_MODES,
  type InteractionMode,
} from '../interaction/click-swap';
import { ACCEPT_ATTR, ACCEPT_ATTR_WITH_BRIDGE, FORMAT_HINT, FORMAT_HINT_WITH_BRIDGE } from '../importer/file-kinds';
import './toolbar.css';

export type { InteractionMode };

/** 三种模式各自的图标（顺序与 INTERACTION_MODES 一致） */
const MODE_ICONS: Record<InteractionMode, (props: { className?: string; size?: number }) => JSX.Element> = {
  select: SelectModeIcon,
  drag: DragModeIcon,
  'click-swap': SwapModeIcon,
};

export interface TabInfo {
  id: string;
  name: string;
  dirty?: boolean;
}

export interface ToolbarProps {
  fileName: string | null;
  tabs: TabInfo[];
  activeTabId: string | null;
  mode: InteractionMode;
  canUndo: boolean;
  canRedo: boolean;
  busy?: boolean;
  historyOpen?: boolean;
  logOpen?: boolean;
  onOpenFile: (file: File) => void;
  /** 点「导出」：把按钮位置交给上层，由上层决定弹菜单还是直接导出 */
  onExport: (anchor: { x: number; y: number }) => void;
  /** 导出菜单是否展开（只用于 aria-expanded 与高亮） */
  exportMenuOpen?: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onToggleHistory: () => void;
  onToggleLog: () => void;
  onModeChange: (mode: InteractionMode) => void;
  onSelectTab: (id: string) => void;
  onCloseTab: (id: string) => void;
  /** 拖到工作区后是否保留表格内容（false = 剪切语义） */
  keepSourceOnWorkspaceDrop: boolean;
  onKeepSourceChange: (keep: boolean) => void;
  /** 本地版的转换桥可用（决定 `accept` 里要不要带上 .xlsb，以及提示文案） */
  bridgeAvailable?: boolean;
}

function cls(...parts: Array<string | false | undefined>): string {
  return parts.filter((part): part is string => typeof part === 'string' && part !== '').join(' ');
}

export function Toolbar({
  fileName,
  tabs,
  activeTabId,
  mode,
  canUndo,
  canRedo,
  busy = false,
  historyOpen = false,
  logOpen = false,
  onOpenFile,
  onExport,
  exportMenuOpen = false,
  onUndo,
  onRedo,
  onToggleHistory,
  onToggleLog,
  onModeChange,
  onSelectTab,
  onCloseTab,
  keepSourceOnWorkspaceDrop,
  onKeepSourceChange,
  bridgeAvailable = false,
}: ToolbarProps): JSX.Element {
  /** 只在真的换了模式时才回调，避免重复点同一个按钮就触发一次"模式变更" */
  const requestMode = (next: InteractionMode): void => {
    if (next !== mode) onModeChange(next);
  };

  /** 可打开的范围随形态走：本地版（且本机 Excel 可用）时连 .xlsb 一起接受 */
  const acceptAttr = bridgeAvailable ? ACCEPT_ATTR_WITH_BRIDGE : ACCEPT_ATTR;
  const formatHint = bridgeAvailable ? FORMAT_HINT_WITH_BRIDGE : FORMAT_HINT;

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>): void => {
    const picked = event.currentTarget.files?.[0];
    if (picked) onOpenFile(picked);
    // 清空 value：否则连续两次选同一个文件不会再触发 change
    event.currentTarget.value = '';
  };

  return (
    <div className={cls('tb-toolbar', busy && 'is-busy')} aria-busy={busy || undefined}>
      {/* 打开表格：label 包 input，点图标即打开原生文件选择框（保持 data-testid="file-input" 契约） */}
      <label
        className={cls('tb-icon-btn', 'tb-file-btn', busy && 'is-busy')}
        data-testid="toolbar-open"
        title={`打开表格（${formatHint}）`}
        aria-disabled={busy || undefined}
      >
        <OpenIcon />
        <input
          className="tb-file-input"
          type="file"
          accept={acceptAttr}
          data-testid="file-input"
          aria-label={`打开表格（${formatHint}）`}
          disabled={busy}
          onChange={handleFileChange}
        />
      </label>

      <div className="tb-group" role="group" aria-label="文件操作">
        <button
          type="button"
          className="tb-icon-btn"
          data-testid="toolbar-export"
          title="导出（默认：保持原格式与原字节；也可另存为其它格式）"
          aria-label="导出（默认：保持原格式与原字节；也可另存为其它格式）"
          aria-haspopup="menu"
          aria-expanded={exportMenuOpen}
          disabled={busy}
          onClick={(event) => {
            // 把按钮位置交给上层：导出菜单就挂在这个按钮旁边
            const rect = event.currentTarget.getBoundingClientRect();
            onExport({ x: rect.left, y: rect.bottom + 4 });
          }}
        >
          <ExportIcon />
        </button>
        <button
          type="button"
          className="tb-icon-btn"
          data-testid="toolbar-undo"
          title="撤销（Ctrl+Z）"
          aria-label="撤销"
          disabled={busy || !canUndo}
          onClick={onUndo}
        >
          <UndoIcon />
        </button>
        <button
          type="button"
          className="tb-icon-btn"
          data-testid="toolbar-redo"
          title="重做（Ctrl+Y）"
          aria-label="重做"
          disabled={busy || !canRedo}
          onClick={onRedo}
        >
          <RedoIcon />
        </button>
      </div>

      {/* 交互模式：三选一（选择 / 拖拽 / 点击互换）。纯图标 + 悬停说明，当前项高亮。 */}
      <div className="tb-seg" role="radiogroup" aria-label="交互模式">
        {INTERACTION_MODES.map((item) => {
          const Icon = MODE_ICONS[item];
          return (
            <button
              key={item}
              type="button"
              role="radio"
              className={cls('tb-seg-btn', mode === item && 'is-active')}
              data-testid={`toolbar-mode-${item}`}
              data-mode={item}
              title={`${INTERACTION_MODE_LABELS[item]}｜${INTERACTION_MODE_HINTS[item]}`}
              aria-label={INTERACTION_MODE_HINTS[item]}
              aria-checked={mode === item}
              disabled={busy}
              onClick={() => requestMode(item)}
            >
              <Icon />
            </button>
          );
        })}
      </div>

      <div className="tb-group" role="group" aria-label="面板">
        <button
          type="button"
          className={cls('tb-icon-btn', historyOpen && 'is-active')}
          data-testid="toolbar-history"
          title="历史记录"
          aria-label="历史记录"
          aria-pressed={historyOpen}
          disabled={busy}
          onClick={onToggleHistory}
        >
          <HistoryIcon />
        </button>
        <button
          type="button"
          className={cls('tb-icon-btn', logOpen && 'is-active')}
          data-testid="toolbar-log"
          title="运行日志"
          aria-label="运行日志"
          aria-pressed={logOpen}
          disabled={busy}
          onClick={onToggleLog}
        >
          <LogIcon />
        </button>
      </div>

      {tabs.length > 0 ? (
        <div className="tb-tabs" role="tablist" aria-label="已打开的标签">
          {tabs.map((tab) => {
            const active = tab.id === activeTabId;
            return (
              // 标签与关闭按钮是兄弟节点：button 不能嵌套 button，同时保证关闭键不会冒泡成"切表"
              <div key={tab.id} className={cls('tb-tab-item', active && 'is-active')}>
                <button
                  type="button"
                  role="tab"
                  className={cls('tb-tab', active && 'is-active')}
                  data-testid={`toolbar-tab-${tab.id}`}
                  title={tab.dirty ? `${tab.name} · 未保存` : tab.name}
                  aria-selected={active}
                  disabled={busy}
                  onClick={() => onSelectTab(tab.id)}
                >
                  {tab.dirty ? <span className="tb-tab-dot" aria-hidden="true" /> : null}
                  <span className="tb-tab-name">{tab.name}</span>
                  {tab.dirty ? <span className="tb-sr-only">（未保存）</span> : null}
                </button>
                <button
                  type="button"
                  className="tb-tab-close"
                  data-testid={`toolbar-tab-close-${tab.id}`}
                  title="关闭标签"
                  aria-label="关闭标签"
                  disabled={busy}
                  // 阻断冒泡：关闭标签绝不能顺带切成当前表
                  onPointerDown={(event) => event.stopPropagation()}
                  onClick={(event) => {
                    event.stopPropagation();
                    onCloseTab(tab.id);
                  }}
                >
                  <CloseIcon size={12} />
                </button>
              </div>
            );
          })}
        </div>
      ) : (
        // 没有任何标签时，用文件名占位（有标签时文件名就是标签本身，不重复展示）
        <div className="tb-tabs tb-tabs-empty">
          {fileName ? (
            <span className="tb-file-name" title={fileName}>
              {fileName}
            </span>
          ) : (
            <span className="tb-file-name is-placeholder">未打开文件</span>
          )}
        </div>
      )}

      {/* 靠工作区（右侧侧栏）一侧：把当前表导入到工作区 */}
      <div className="tb-right">
        {/* 拖到工作区后是否保留表格内容（false = 剪切语义）。「+」已挪到工作区右上角。 */}
        <label className="tb-check" title="勾选：拖到工作区后表格里的内容保留；取消：拖过去后源内容清空（剪切）">
          <input
            type="checkbox"
            data-testid="toolbar-keep-source"
            checked={keepSourceOnWorkspaceDrop}
            onChange={(event) => onKeepSourceChange(event.target.checked)}
          />
          <span>拖到工作区保留内容</span>
        </label>
      </div>
    </div>
  );
}
