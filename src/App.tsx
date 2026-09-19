import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type Dispatch,
  type MouseEvent as ReactMouseEvent,
  type SetStateAction,
} from 'react';

import { ICommandService, IConfigService, IContextService, IUndoRedoService, IUniverInstanceService, LocaleType, type IRange } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { IRenderManagerService, SHEET_VIEWPORT_KEY, Vector2 } from '@univerjs/engine-render';
import { SetStyleCommand, SheetsSelectionsService } from '@univerjs/sheets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { ISheetSelectionRenderService, SheetSkeletonManagerService } from '@univerjs/sheets-ui';

import { createSampleWorkbook } from './data/sample';
import {
  createClickSwapController,
  INTERACTION_MODE_HINTS,
  normalizeInteractionMode,
  type ClickSwapControllerWithGate,
  type InteractionMode,
  type SwapSide,
} from './interaction/click-swap';
import { createDragController, type DragController, type DropHint } from './interaction/drag-controller';
import { flyIn } from './interaction/fly-in';
import { renderSnapshotToElement } from './interaction/ghost-content';
import { createSwapAnimation, type SwapAnimation } from './interaction/swap-animation';
import { cellAtPoint } from './interaction/cell-geometry';
import {
  describeSingleBlockOnly,
  normalizeRanges,
  parseRangeList,
  rectsToA1,
  summarizeSelection,
  totalCells,
} from './interaction/selection-ranges';
import { expandRectToMerges } from './interaction/range-merge';
import {
  beginScrollDrag,
  isInBarStrip,
  isOnThumb,
  scrollForPointer,
  scrollFromTrackClick,
  type AxisGeometry,
} from './interaction/scrollbar-drag';
import {
  computeFlashRect,
  createDragTargetHighlight,
  createSwapFlashOverlay,
  type FlashBox,
  type FlashTarget,
  type SwapFlashOverlay,
} from './interaction/swap-flash';
import { toUniverWorkbook, type ImportOutcome } from './importer/to-univer';
import {
  applyWorkbookFeatures,
  countRetainedImageObjectUrls,
  releaseAllImageObjectUrls,
  releaseWorkbookImageObjectUrls,
  retainedImageObjectUrlsByWorkbook,
  type ApplyFeaturesResult,
} from './importer/apply-features';
import { exportXlsx, type CellEdit, type ExportSource, type SheetEdits } from './exporter/export-xlsx';
import { slimForExport } from './exporter/slim-source';
import { buildDelimited, encodeDelimited } from './exporter/csv-export';
import { convertViaBridge } from './exporter/bridge';
import {
  APP_FORM,
  BRIDGE_EXTENSION,
  bridgeUnavailableHint,
  bridgedFileName,
  initialBridgeStatus,
  probeBridge,
  type BridgeFormat,
  type BridgeStatus,
} from './shell/app-form';
import type { ParsedWorkbook } from './parser/types';
import { parseXlsx } from './parser';
import { clearLog, countKind, findLast, log, p0log, type P0LogEntry } from './p0/log';
import {
  DEFAULT_SETTINGS,
  SIDEBAR_WIDTH_RANGE,
  createAutoSaver,
  createSessionStore,
  normalizeSettings,
  type SessionSettings,
  type SessionState,
  type SessionTab,
  type SessionTabEdit,
} from './persistence/session';
import {
  clearWorkbookDirty,
  dirtyCellCount,
  getDirtyCells,
  pauseDirtyTracking,
  resetAllDirty,
  resumeDirtyTracking,
} from './univer/dirty-tracker';
import { installContentOnlyLock, probePermissionApi, type ContentOnlyLock } from './univer/lock';
import { installReadOnlyGuard, type ReadOnlyGuard } from './univer/read-only-guard';
import { bootUniver, loadWorkbook, type UniverBoot } from './univer/setup';
import { disableFillHandleOn, fillHandleStateOn, installFillHandleOff, type RenderUnitLike } from './univer/fill-handle';
import { clearRange, ensureSwapCommandRegistered, moveRange, swapRanges } from './univer/swap-command';
import {
  MAX_WORKSPACE_ITEMS_PER_ACTION,
  applySnapshot,
  extractCellItems,
  extractSnapshot,
  isSnapshotEmpty,
  previewWorkspaceImport,
  rectToA1,
} from './workspace/snapshot';
import { ImportDialog, type ImportTargetKind } from './workspace/ImportDialog';
// 面板的"目标范围换算"单独一个模块：App.tsx 只导出组件，Fast Refresh 才不会被降级成整页刷新
import { resolveImportRanges } from './workspace/import-target';
import {
  FORMAT_HINT,
  SUPPORTED_EXTENSIONS,
  exportFileNameFor,
  fileKindOf,
  isBridgeOnlyFile,
  isSupportedWorkbookFile,
  unsupportedFileMessage,
} from './importer/file-kinds';
import { openWorkbookBytes } from './importer/open-workbook';
import type { DropTarget, DragPayload, RangeSnapshot } from './workspace/types';
import { EMPTY_FILTER, type WorkspaceFilter } from './workspace/filter';
import { WorkspacePanel } from './workspace/WorkspacePanel';
import { Toolbar, type TabInfo } from './shell/Toolbar';
import { HistoryPanel } from './shell/HistoryPanel';
import { writeClipboardText } from './shell/clipboard-write';
import { ContextMenu } from './shell/ContextMenu';
import type { MenuItemSpec } from './shell/context-menu-model';
import { appendEntry, planJump, type HistoryEntry, type HistoryKind } from './shell/history-model';
import {
  DEFAULT_RESIDENT_TAB_LIMIT,
  dropTab,
  markTabCold,
  markTabUsed,
  needsBuild,
  planEvictions,
  type TabRuntime,
} from './shell/resident-tabs';
import { setUndoShortcutHandler } from './shell/undo-shortcut';

import './shell/shell.css';
import './interaction/swap-animation.css';

const CONTAINER_ID = 'univer-container';
/** "还没有打开任何文件"时的状态栏提示；启动引导与"关掉最后一个标签摆回占位"两处共用，必须一字不差 */
const IDLE_STATUS_TEXT = `拖入表格文件（${SUPPORTED_EXTENSIONS.join(' / ')}）或使用左侧示例开始`;
/** "算不算一次拖动"的位移阈值（px）。取 8 而非 4：手抖常超 4px，会把"点一下"误判成拖放 */
const MOVE_TOLERANCE_PX = 8;
/** 「选中后 3 秒内点工作区空白即加入」的时间窗（毫秒）；超时不动手，避免随手点击莫名多出条目 */
const QUICK_ADD_WINDOW_MS = 3000;
/** 支持上述捷径的交互模式。拖拽模式刻意排除：它的手势语言是"按住拖"，点空白易被当成"放下/取消" */
const QUICK_ADD_MODES: readonly InteractionMode[] = ['select', 'click-swap'];
/** 首帧等待窗口（毫秒）：Univer 绘制异步，刚导入/切标签就断言"画布宽度 0 ⇒ 白屏"会误报 */
const RENDER_SETTLE_MS = 600;
/** 重绑之后**再**给它多久把首帧画出来（超过就记 `render:rebind-failed`，用户按 F5 仍可恢复） */
const RENDER_REBIND_MS = 1200;
/** 抓滚动条滑块的容差（px）：滑块只有几像素宽，偏一点也要算"抓住了" */
const SCROLLBAR_GRAB_TOLERANCE_PX = 6;
/** 滚动条交互条带厚度（px）：从视口右缘/下缘往里算。取 14（滑块约 6px + 8px 容差），再往里会抢走最后一列拖拽 */
const SCROLLBAR_STRIP_PX = 14;
/** 启动时的示例工作簿 id（首个真实文件进来后就释放它，别白占一份工作簿内存） */
const SAMPLE_WORKBOOK_ID = 'p0-workbook';

type Status = 'booting' | 'ready' | 'importing' | 'error';

interface ParsedFeatureCounts {
  cf: number;
  dv: number;
  links: number;
  notes: number;
  tables: number;
  images: number;
}

interface ImportSummary {
  fileName: string;
  sheets: string[];
  unsupported: string[];
  warnings: string[];
  /** 已解析并原样保留、但不影响预览的信息（打印设置）；不是缺陷，状态栏不报警只做说明 */
  preserved: string[];
  parseMs: number;
  adaptMs: number;
  renderMs: number;
  featureMs: number;
  featureIssues: string[];
  featureCounts: ApplyFeaturesResult['counts'] | null;
  /** 解析层实际识别到的特性数量（与"应用结果"分开统计，便于定位是解析没做到还是应用没做到） */
  parsedFeatures: ParsedFeatureCounts;
  cells: number;
}

interface Toast {
  id: number;
  text: string;
  kind: 'info' | 'warn';
}

export function App() {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const sidebarRef = useRef<HTMLDivElement | null>(null);

  /** 启动打点：`app:mounted` = 界面已画出（此时引擎还在引导）；e2e 的启动预算断言这三段 */
  useEffect(() => {
    performance.mark('app:mounted');
  }, []);

  /** 探测本地转换桥：形态决定是否显示入口，结果决定点了能否成功（本地版也可能没装 Excel）。静态形态不发请求 */
  useEffect(() => {
    let cancelled = false;
    void probeBridge().then((status) => {
      if (cancelled) return;
      setBridge(status);
      log('app:bridge-status', {
        form: APP_FORM,
        available: status.available,
        excel: status.excel ?? null,
        reason: status.reason ?? null,
      });
    });
    return () => {
      cancelled = true;
    };
  }, []);
  const apiRef = useRef<FUniver | null>(null);
  const lockRef = useRef<ContentOnlyLock | null>(null);
  const guardRef = useRef<ReadOnlyGuard | null>(null);
  /**
   * 提醒框定位用的 scene/skeleton 缓存（按 unitId）。必须放组件级：释放工作簿的路径在引导 effect 之外，
   * 否则 disposeUnit 后缓存仍指着已释放的渲染单元（大表这一份对象不小）；见 releaseUnitBookkeeping。
   */
  const renderCacheRef = useRef<null | { unitId: string; scene: unknown; skeleton: unknown }>(null);
  /** 分隔条拖动进行中的"摘监听"函数（卸载兜底用，见 startSidebarResize） */
  const sidebarResizeCleanupRef = useRef<(() => void) | null>(null);
  const instanceServiceRef = useRef<IUniverInstanceService | null>(null);
  const controllerRef = useRef<DragController | null>(null);
  const pendingWorkspaceDragRef = useRef<{ snapshot: RangeSnapshot; x: number; y: number; started: boolean } | null>(null);
  /** 这一次按下落在哪个工作区条目上（没拖动就松手 = 点击，用于"点击互换"选边） */
  const pendingWorkspaceClickRef = useRef<RangeSnapshot | null>(null);
  /** 最近一次在表格里动手的时间（点选/按键）；「选中后点空白即加入」靠它判断是否刚选完 */
  const lastGridInteractionAtRef = useRef(0);
  /** 工作区面板上按下的位置：用于把"点一下空白"和"拖到面板里松手"区分开 */
  const workspacePressRef = useRef<{ x: number; y: number } | null>(null);
  /** 最近一次"快速加入"用的选区（同一次选区只加一次，避免连点收两遍） */
  const lastQuickAddRef = useRef<{ key: string; at: number } | null>(null);
  /** **单元级操作的串行闸**：打开文件/冷标签重建/会话恢复都会新建并绑定工作簿，交错执行会让活动单元
   * 与画布绑定的单元错位 → 舞台空白。排进同一条 Promise 链一次只跑一个，只圈"会动工作簿"的异步段 */
  const unitOpsRef = useRef<Promise<unknown>>(Promise.resolve());
  const runUnitOp = useCallback(<T,>(label: string, fn: () => Promise<T> | T): Promise<T> => {
    const next = unitOpsRef.current.then(
      () => fn(),
      () => fn(),
    );
    unitOpsRef.current = next.catch(() => undefined);
    next.then(
      () => log('unit-op:done', { label }),
      (error) => log('unit-op:error', { label, message: String(error) }),
    );
    return next;
  }, []);
  /** `handleWorkspaceItemClick` 的 ref 版本（指针监听装配得比它早） */
  const handleWorkspaceItemClickRef = useRef<(item: RangeSnapshot) => void>(() => {});
  const dropHandlerRef = useRef<(payload: DragPayload, target: DropTarget | null, pointer: { x: number; y: number }) => void>(() => {});
  const summaryRef = useRef<ImportSummary | null>(null);
  const featureResultRef = useRef<ApplyFeaturesResult | null>(null);
  /** 当前标签的导出数据源（瘦身后、含原始 zip 字节）；导出时用它做"外科式修补" */
  const parsedRef = useRef<ExportSource | null>(null);
  const importedFileNameRef = useRef<string>('workbook.xlsx');

  /** 每个标签只留**瘦身后**的导出数据源：完整模型里上百万 ParsedCell 是内存大头（实测百万格堆 543 MB） */
  const tabsDataRef = useRef<Array<{ id: string; fileName: string; source: ExportSource | null; bytes: Uint8Array }>>([]);
  /**
   * 标签的"实体化"记账（多标签内存控制核心）：Univer 每个实体化工作簿 ≈ 数百 MB，
   * 所以打开的标签可以很多，但同时实体化只保留最近用过的 K 个（见 `src/shell/resident-tabs.ts`）。
   */
  const tabRuntimeRef = useRef<TabRuntime[]>([]);
  const activeTabIdRef = useRef<string | null>(null);
  /** 拖动期间被"钉住"的源区域 A1（防止 Univer 把长按拖动当成扩选） */
  const dragSourceA1Ref = useRef<string | null>(null);
  /** 拖动期间逐帧"钉"选区的 rAF 句柄 */
  const pinRafRef = useRef<number | null>(null);
  /** 清空选区（互换/搬运收尾用）；由 attachSheetDeps 装配，避免各处重复取服务 */
  const clearSelectionRef = useRef<() => void>(() => {});
  /** 互换后的黄色提醒框（自动渐隐）；同样由 attachSheetDeps 装配 */
  const flashSwapRef = useRef<(firstA1: string, secondA1?: string) => void>(() => {});
  /** 互换后的黄色提醒框浮层（自动渐隐；见 src/interaction/swap-flash.ts） */
  const flashOverlayRef = useRef<SwapFlashOverlay | null>(null);
  /** 拖动过程中"落点高亮"浮层（常亮，松手即收） */
  const targetHighlightRef = useRef<SwapFlashOverlay | null>(null);
  /** 释放提醒框；由 attachSheetDeps 装配 */
  const clearSwapFlashRef = useRef<() => void>(() => {});
  /** 最近一次指针按下是否带 Shift/Ctrl（带修饰键 = 用户要原生框选：不当选边、也不搬运） */
  const modifierPressRef = useRef(false);
  /** 根/UI 作用域的注入器：选区渲染服务只在这里能取到（sheet 作用域取会抛错） */
  const rootInjectorRef = useRef<{ get: (token: unknown) => unknown } | null>(null);
  /** 撤销/重做服务（冷存标签重建后必须清它的栈，见 buildTabUnit 的注释） */
  const undoRedoServiceRef = useRef<{ clearUndoRedo?: (unitId: string) => void } | null>(null);
  const modeRef = useRef<InteractionMode>('drag');
  const clickSwapRef = useRef<ClickSwapControllerWithGate | null>(null);
  const swapAnimRef = useRef<SwapAnimation | null>(null);
  const lastPointerRef = useRef<{ x: number; y: number }>({ x: 0, y: 0 });
  // ---- 会话持久化（误关闭恢复）----
  const sessionStoreRef = useRef<ReturnType<typeof createSessionStore> | null>(null);
  const sessionSaverRef = useRef<ReturnType<typeof createAutoSaver> | null>(null);
  const lastSessionFingerprintRef = useRef<string>('');
  /** 供测试钩子调用（避免在钩子里直接引用后定义的 const） */
  const snapshotActiveEditsRef = useRef<() => void>(() => {});
  /** 导入文件（由 handleFile 装配）。用 ref：窗口级拖放监听装在引导 effect 里，早于 handleFile 定义 */
  const fileOpenerRef = useRef<(file: File) => void>(() => {});
  /** 按标签 id 快照编辑；用 ref 暴露（coldStoreTab 定义在 snapshotEditsFor 之前，直接引用会踩 TDZ） */
  const snapshotEditsForRef = useRef<(id: string) => number>(() => 0);
  /** 恢复过程中不要触发自动保存（否则会把"恢复了一半"的状态写回去） */
  const restoringRef = useRef(false);
  const itemsStateRef = useRef<RangeSnapshot[]>([]);
  /** 用户设置（拖放语义 / 预览列数 / 工作区宽度）：同步 ref 供手势层即时读取 */
  const settingsRef = useRef<SessionSettings>(DEFAULT_SETTINGS);
  /** 工作区剪贴板：在工作区条目上"复制/剪切"后，可粘贴到表格 */
  const clipboardRef = useRef<RangeSnapshot | null>(null);
  /** 快照的同步存储：React setState 是异步的，测试/程序化流程可能在重渲染前就要按 id 取回条目 */
  const snapshotStoreRef = useRef<Map<string, RangeSnapshot>>(new Map());
  const lastTargetKeyRef = useRef<string>('');

  const [status, setStatus] = useState<Status>('booting');
  const [statusText, setStatusText] = useState('正在启动表格引擎…');
  /** 有文件被拖到窗口上方（显示"松手即打开"的提示；见下方窗口级 drop 处理） */
  const [fileDropActive, setFileDropActive] = useState(false);
  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [items, setItems] = useState<RangeSnapshot[]>([]);
  const [dragging, setDragging] = useState(false);
  const [hint, setHint] = useState<DropHint | 'workspace' | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheetName, setActiveSheetName] = useState('');
  const [selectionText, setSelectionText] = useState('A1');
  const [fps, setFps] = useState(0);
  const [logOpen, setLogOpen] = useState(false);
  const [entries, setEntries] = useState<P0LogEntry[]>([]);
  const [toasts, setToasts] = useState<Toast[]>([]);
  const [reducedMotion, setReducedMotion] = useState(false);
  const [tabs, setTabs] = useState<TabInfo[]>([]);
  const [activeTabId, setActiveTabId] = useState<string | null>(null);
  const [mode, setMode] = useState<InteractionMode>('drag');
  /** 用户设置（拖放语义 / 预览列数 / 工作区宽度），随会话持久化 */
  const [settings, setSettings] = useState<SessionSettings>(DEFAULT_SETTINGS);
  /** 剪贴板里有什么（只用于 UI 提示与菜单可用性；内容在 clipboardRef） */
  const [clipboardLabel, setClipboardLabel] = useState<string | null>(null);
  /** 工作区条目的右键菜单（复用共享的 ContextMenu 组件） */
  const [wsMenu, setWsMenu] = useState<{ x: number; y: number; item: RangeSnapshot } | null>(null);
  /** 工作区的搜索 / 来源筛选（纯视图状态，不持久化） */
  const [wsFilter, setWsFilter] = useState<WorkspaceFilter>(EMPTY_FILTER);
  const updateWsFilter = useCallback((patch: Partial<WorkspaceFilter>) => {
    setWsFilter((prev) => ({ ...prev, ...patch }));
  }, []);
  /** 「清空」的二次确认（不可逆批量操作，必须问一次） */
  const [wsClearConfirming, setWsClearConfirming] = useState(false);
  /** 点击互换模式下"已经点过的另一方"（把工作区条目高亮成待互换）；表格那边由选区框体现，这边必须自己反馈 */
  const [pendingSwapItemId, setPendingSwapItemId] = useState<string | null>(null);
  const [pendingSwapLabel, setPendingSwapLabel] = useState<string | null>(null);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
  const [importBusy, setImportBusy] = useState(false);
  /** 导出菜单锚点（点「导出」时由工具栏给按钮位置） */
  const [exportMenu, setExportMenu] = useState<{ x: number; y: number } | null>(null);
  /** 正在经本机 Excel 转换的格式（转换期间禁用菜单，避免重复提交） */
  const [exportBusy, setExportBusy] = useState<BridgeFormat | null>(null);
  /** 本地转换桥状态（静态形态恒为不可用；本地版探测 /api/bridge/health） */
  const [bridge, setBridge] = useState<BridgeStatus>(() => initialBridgeStatus());
  /** 桥状态的 ref 版本：引导 effect 里的一次性监听器（拖放）要读它 */
  const bridgeStatusRef = useRef(bridge);
  bridgeStatusRef.current = bridge;
  /**
   * 右键菜单打开那一刻**快照**选区（菜单聚焦后画布选区就不可读）。
   * `a1List` 是全部选区（已规范化去重）；`a1` 是主区域，供只支持单块的操作使用。
   */
  interface MenuSelectionSnapshot {
    a1: string;
    a1List: string[];
    rows: number;
    cols: number;
    text: string;
  }
  const menuRef = useRef<{ x: number; y: number; snapshot: MenuSelectionSnapshot } | null>(null);
  const [menu, setMenu] = useState<{
    x: number;
    y: number;
    snapshot: MenuSelectionSnapshot;
  } | null>(null);
  const [canUndo, setCanUndo] = useState(false);
  const [canRedo, setCanRedo] = useState(false);
  // ---- 历史记录：我们自己记账，跳步靠连续撤销/重做 ----
  const [historyEntries, setHistoryEntries] = useState<HistoryEntry[]>([]);
  const [historyIndex, setHistoryIndex] = useState(0);
  /** 账本的**同步**副本：setState 异步，加入工作区后立刻 Ctrl+Z 时读 state 会错位；以 ref 为准，state 只渲染 */
  const entriesRef = useRef<HistoryEntry[]>([]);
  const [undoRedoCounts, setUndoRedoCounts] = useState({ undos: 0, redos: 0 });
  const historyIndexRef = useRef(0);
  const prevUndosRef = useRef(0);
  const pendingHistoryLabelRef = useRef<{ label: string; kind: HistoryKind; at: number } | null>(null);
  /** 正在进行中的表格动作。我们执行的表格命令会同时被 Univer 的 undos 订阅和我们自己的 pushHistory
   * 记账，账本多出"幽灵步骤"（撤销要多按一次）；动作开始时记下已自动入账条数，结束时收敛成一条 */
  const sheetActionRef = useRef<{ label: string; kind: HistoryKind; startAuto: number } | null>(null);
  /** 订阅自动补出来的表格历史 id（按入账顺序；`pushHistory` 认领后移出） */
  const autoSheetEntriesRef = useRef<string[]>([]);
  /** 我们自己的 api.undo()/redo() 引起的 undos 变化不予记账；命令执行期间同步置位，无漏抑制窗口 */
  const suppressAutoEntriesRef = useRef(0);
  /** **长时间窗口**的抑制开关（上一个是"一次性"的）：打开文件时应用特性会连发一串命令，订阅分多次
   * 看到 undos 增长，账本凭空多出幽灵条目、用户按 Ctrl+Z 看着"什么都没撤掉"。用布尔窗口不用计数器 */
  const pauseAutoEntriesRef = useRef(false);
  const undoRedoCountsRef = useRef({ undos: 0, redos: 0 });
  /** 最近一次历史入账时间（用于"打字产生的编辑"去重） */
  const lastHistoryPushAtRef = useRef(0);
  const lastDirtySeenRef = useRef(0);

  // 让测试钩子/回调总能读到最新的工作区条目（它们是引导时一次性安装的，不能闭包捕获 state）
  itemsStateRef.current = items;
  settingsRef.current = settings;

  /** 改设置：逐字段收敛到合法值，并触发一次自动保存（指纹里含 settings） */
  const updateSettings = useCallback((patch: Partial<SessionSettings>) => {
    setSettings((prev) => normalizeSettings({ ...prev, ...patch }));
  }, []);

  // ---------------------------------------------------------------- 小工具
  /** 提示条：3.2 秒后自动消失。定时器句柄要记下并在卸载时清掉（HMR 重挂会对着旧实例 setState） */
  const toastTimersRef = useRef<Set<number>>(new Set());
  const toast = useCallback((text: string, kind: Toast['kind'] = 'info') => {
    const id = Date.now() + Math.random();
    setToasts((prev) => [...prev.slice(-2), { id, text, kind }]);
    const timer = window.setTimeout(() => {
      toastTimersRef.current.delete(timer);
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 3200);
    toastTimersRef.current.add(timer);
  }, []);

  // 与上面成对：卸载时清掉待触发的提示定时器
  useEffect(
    () => () => {
      for (const timer of toastTimersRef.current) window.clearTimeout(timer);
      toastTimersRef.current.clear();
    },
    [],
  );

  const getSheet = useCallback((): FWorksheet | null => apiRef.current?.getActiveWorkbook()?.getActiveSheet() ?? null, []);

  /**
   * 读当前工作表的**全部**选区（支持 Ctrl+点选的多块）并规范化成 A1 列表。
   * `normalizeRanges` 会去重并丢掉被包含的小块；读不到时退回单块，再读不到给空数组（不抛异常）。
   */
  const readSelectionA1List = useCallback((): string[] => {
    try {
      const sheet = apiRef.current?.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return [];
      const selection = sheet.getSelection?.() ?? null;
      const list = selection?.getActiveRangeList?.().map((range) => range.getA1Notation()) ?? [];
      if (list.length > 0) return rectsToA1(normalizeRanges(list));
      const single = sheet.getActiveRange?.()?.getA1Notation?.() ?? null;
      return single ? [single] : [];
    } catch {
      return [];
    }
  }, []);

  const refreshSheetMeta = useCallback(() => {
    const api = apiRef.current;
    const wb = api?.getActiveWorkbook();
    if (!wb) return;
    const sheet = wb.getActiveSheet();
    setSheetNames(wb.getSheets().map((s) => s.getSheetName()));
    setActiveSheetName(sheet?.getSheetName() ?? '');
    const list = readSelectionA1List();
    if (list.length > 0) setSelectionText(summarizeSelection(list));
  }, [readSelectionA1List]);

  // ---------------------------------------------------------------- 状态栏选区
  /** 状态栏的"选区 …"必须跟着真实选区走、多块一起报。文案由 `summarizeSelection` 生成：
   * 单块 `B2:D5`，多块 `B2 + D4（2 块 / 2 格）` */
  useEffect(() => {
    const timer = window.setInterval(() => {
      if (document.hidden) return;
      const list = readSelectionA1List();
      if (list.length === 0) return;
      const text = summarizeSelection(list);
      setSelectionText((prev) => (prev === text ? prev : text));
    }, 250);
    return () => window.clearInterval(timer);
  }, [readSelectionA1List]);

  /** 每次活动工作表变化（首次引导 / 导入 / 切换标签）都要重装格式锁 */
  /** 关掉选区右下角的"填充柄"（见 `src/univer/fill-handle.ts`）：它是自动填充入口，而本产品只允许
   * 逐格改内容；拖它还会被我们的内容拖拽接管、悄悄互换两格内容。控件按选区新建，每次装配都要重扣 */
  const fillHandleLoggedRef = useRef<Set<string>>(new Set());
  const disableFillHandleForCurrentUnit = useCallback((): boolean => {
    try {
      const unitId = apiRef.current?.getActiveWorkbook()?.getId();
      if (!unitId) return false;
      const renderManager = rootInjectorRef.current?.get(IRenderManagerService) as unknown as {
        getRenderById?: (id: string) => RenderUnitLike | null;
      } | null;
      const renderUnit = renderManager?.getRenderById?.(unitId) ?? null;
      const result = disableFillHandleOn(renderUnit, ISheetSelectionRenderService);
      if (result.firstTime && !fillHandleLoggedRef.current.has(unitId)) {
        fillHandleLoggedRef.current.add(unitId);
        log('ui:fill-handle-off', { unitId });
      }
      return result.ok;
    } catch (error) {
      log('ui:fill-handle-error', { message: String(error) });
      return false;
    }
  }, []);

  const attachSheetDeps = useCallback(() => {
    clearSwapFlashRef.current(); // 换了工作表 → 上一张表的黄色提醒框不再适用
    const sheet = getSheet();
    if (!sheet) return;
    lockRef.current?.restore();
    lockRef.current = installContentOnlyLock(sheet);
    ensureSwapCommandRegistered(sheet);
    probePermissionApi(sheet);
    refreshSheetMeta();
    // 同上：填充柄在"只改内容"的约束下是误导入口
    disableFillHandleForCurrentUnit();
    log('app:lock-attached', { sheet: sheet.getSheetName() });
  }, [getSheet, refreshSheetMeta]);

  /** 工作区"撤销/重做"用的快照（键 = 历史条目 id；条目被截断时顺带清理） */
  const workspaceSnapshotsRef = useRef<Map<string, { before: RangeSnapshot[]; after: RangeSnapshot[] }>>(new Map());
  /** pushHistory 的 ref 版本：commitWorkspace 定义在前面，直接引用会踩 TDZ */
  const pushHistoryRef = useRef<(label: string, kind: HistoryKind, scope?: 'sheet' | 'workspace', entryId?: string) => void>(
    () => {},
  );
  /** 撤销/重做的 ref 版本：测试钩子装配得比它们早，直接引用会踩 TDZ */
  const stepBackRef = useRef<() => Promise<boolean>>(async () => false);
  const stepForwardRef = useRef<() => Promise<boolean>>(async () => false);
  /** commitWorkspace 的 ref 版本（同样为了绕开装配顺序） */
  const commitWorkspaceRef = useRef<
    (next: RangeSnapshot[] | ((prev: RangeSnapshot[]) => RangeSnapshot[]), label: string, kind: HistoryKind) => void
  >(() => {});
// ---------------------------------------------------------------- 落点处理
  /** 把 `snapshotStoreRef`（id → 快照）与当前 items 列表对齐（多删少补）。不按当前列表对齐时，
   * 撤销后消失的条目快照永远删不掉（id 永不复用）；撤销"移除条目"又会留下 store 里没有的条目，
   * 读 store 的路径（拖拽写回、粘贴）会取到 undefined */
  const syncSnapshotStore = useCallback((next: RangeSnapshot[]) => {
    const alive = new Set(next.map((item) => item.id));
    for (const id of [...snapshotStoreRef.current.keys()]) if (!alive.has(id)) snapshotStoreRef.current.delete(id);
    for (const item of next) if (!snapshotStoreRef.current.has(item.id)) snapshotStoreRef.current.set(item.id, item);
  }, []);

  /** 工作区状态的**唯一提交入口**：工作区增删只改我们的 React state，Univer 撤销栈不知情，不收口
   * Ctrl+Z 就只会撤表格内容。每次改动记一条带 before/after 快照的 workspace 历史，并同步
   * itemsStateRef 与 snapshotStoreRef，让撤销/重做按时间顺序在两类动作间切换 */
  const commitWorkspace = useCallback(
    (next: RangeSnapshot[] | ((prev: RangeSnapshot[]) => RangeSnapshot[]), label: string, kind: HistoryKind) => {
      const before = itemsStateRef.current;
      const after = typeof next === 'function' ? next(before) : next;
      if (after === before) return;
      itemsStateRef.current = after;
      setItems(after);
      // 同步快照表（按 id 增删，避免整表重建）；撤销/重做路径也走同一个函数，见 syncSnapshotStore
      syncSnapshotStore(after);
      // 记历史（带快照，撤销/重做靠它回放）
      const entryId = `ws-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
      workspaceSnapshotsRef.current.set(entryId, { before, after });
      pushHistoryRef.current(label, kind, 'workspace', entryId);
      log('workspace:commit', { label, before: before.length, after: after.length });
    },
    [syncSnapshotStore],
  );
  commitWorkspaceRef.current = commitWorkspace;

  /**
   * 批量放入工作区（新条目排最前，保持区域内行优先顺序）。
   * 入口级去重（第二道防线）：同一个快照 id 只能存在一份；从表格重拖会生成新 id，不受影响。
   */
  const addWorkspaceItems = useCallback(
    (incoming: RangeSnapshot[], label?: string): number => {
      if (incoming.length === 0) return 0;
      const existing = new Set(itemsStateRef.current.map((item) => item.id));
      const fresh = incoming.filter((item, index) => !existing.has(item.id) && incoming.findIndex((other) => other.id === item.id) === index);
      const dropped = incoming.length - fresh.length;
      if (dropped > 0) log('workspace:duplicate-ignored', { dropped });
      if (fresh.length === 0) return 0;
      commitWorkspaceRef.current(
        (prev) => [...fresh.map((item) => ({ ...item })), ...prev],
        label ?? `加入工作区 ${fresh.length} 格`,
        'workspace',
      );
      for (const item of fresh) {
        log('workspace:add', { id: item.id, a1: item.source.a1, rows: item.rows, cols: item.cols });
      }
      return fresh.length;
    },
    [commitWorkspace],
  );

  /** 放入工作区的**唯一实现**（所有入口共用）：工作区存的是一个个独立单元格、空内容一律跳过，
   * 所以把每块拆成 1×1 条目。多块一次提交（一条历史、一次渲染），返回值表示是否真放进了东西 */
  const addWorkspaceFromRanges = useCallback(
    (a1List: string[], pointer?: { x: number; y: number }): { ok: boolean; items: RangeSnapshot[]; skipped: number; truncated: number } => {
      const sheet = getSheet();
      if (!sheet || a1List.length === 0) return { ok: false, items: [], skipped: 0, truncated: 0 };
      const collected: RangeSnapshot[] = [];
      let skipped = 0;
      let truncated = 0;
      for (const a1 of a1List) {
        const { items, skippedEmpty, truncated: cut, error } = extractCellItems(sheet, a1);
        if (error) {
          toast(error, 'warn');
          return { ok: false, items: [], skipped: 0, truncated: 0 };
        }
        collected.push(...items);
        skipped += skippedEmpty;
        truncated += cut;
      }
      const scope = a1List.join(' ');
      if (collected.length === 0) {
        toast('所选内容全是空的，已跳过', 'warn');
        log('workspace:skip-empty', { skipped, a1: scope, blocks: a1List.length, items: 0 });
        return { ok: false, items: [], skipped, truncated };
      }

      const added = addWorkspaceItems(collected);
      if (added === 0) {
        // 同一个快照又被收了一遍（例如工作区内部拖动被误当成"从表格拖进来"）
        toast('这些内容已经在工作区里了', 'warn');
        log('workspace:skip-duplicate', { a1: scope, items: collected.length });
        return { ok: false, items: collected, skipped, truncated };
      }
      const parts: string[] = [`已放入 ${added} 个单元格`];
      if (a1List.length > 1) parts.unshift(`已从 ${a1List.length} 块区域`);
      if (skipped > 0) parts.push(`跳过 ${skipped} 个空内容单元格`);
      if (truncated > 0) parts.push(`超出上限未放入 ${truncated} 个`);
      toast(parts.join(' · '));
      // 日志字段是 e2e 与用户排错依赖的口径，勿改
      if (skipped > 0) log('workspace:skip-empty', { skipped, a1: scope, blocks: a1List.length });
      if (truncated > 0) log('workspace:truncated', { truncated, a1: scope, max: MAX_WORKSPACE_ITEMS_PER_ACTION });

      // 飞入动画：从指针位置飞到第一个卡片的落点（批量时只飞一次，避免上百个动画卡顿）
      if (pointer) {
        const first = collected[0];
        window.requestAnimationFrame(() => {
          const card = sidebarRef.current?.querySelector(`[data-snapshot-id="${first.id}"]`) as HTMLElement | null;
          if (!card) return;
          const to = card.getBoundingClientRect();
          const flyer = renderSnapshotToElement(first, { scale: 0.62, maxRows: 4, maxCols: 5 });
          Object.assign(flyer.style, {
            position: 'fixed',
            left: `${to.left}px`,
            top: `${to.top}px`,
            width: `${to.width}px`,
            height: `${to.height}px`,
            overflow: 'hidden',
            zIndex: '9998',
            pointerEvents: 'none',
          });
          document.body.appendChild(flyer);
          void flyIn(
            flyer,
            { left: pointer.x - 56, top: pointer.y - 28, width: 112, height: 56 },
            { left: to.left, top: to.top, width: to.width, height: to.height },
            { reducedMotion },
          ).then(() => flyer.remove());
        });
      }
      return { ok: true, items: collected, skipped, truncated };
    },
    [addWorkspaceItems, getSheet, log, reducedMotion, toast],
  );

  /** 单区域版本（拖拽落点用）：转发到多块实现，并把"到底放进去没有"如实返回给调用方 */
  const addWorkspaceItem = useCallback(
    (snapshot: RangeSnapshot, pointer: { x: number; y: number }) => addWorkspaceFromRanges([snapshot.source.a1], pointer),
    [addWorkspaceFromRanges],
  );

  const isInsideSidebar = useCallback((pointer: { x: number; y: number }): boolean => {
    // 先用 DOM 命中判断（能覆盖工作区面板的表头/表尾/滚动条与被面板遮挡的区域），
    // 拿不到元素时再退回"矩形包含"判断。
    try {
      const element = document.elementFromPoint(pointer.x, pointer.y);
      if (element?.closest('.workspace-host, .ws-panel, .body-splitter')) return true;
    } catch {
      /* 指针在窗口外时会抛错，走下面的兜底 */
    }
    const el = sidebarRef.current;
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    return pointer.x >= rect.left && pointer.x <= rect.right && pointer.y >= rect.top && pointer.y <= rect.bottom;
  }, []);

  /**
   * **快速加入**：选择模式下，刚在表格里选过（`QUICK_ADD_WINDOW_MS` 内）时，点工作区面板**空白处**
   * 就把当前选区收进工作区。判定刻意收紧（宁可不动手也不误加）：必须是点击（位移 < `MOVE_TOLERANCE_PX`）、
   * 落点必须是真的空白（不在条目/按钮/输入框/下拉/标签上）、只在支持的交互模式生效、必须在"刚在表格里
   * 动过手"的窗口内、同一次选区只加一次。
   * 落地走唯一入口 `addWorkspaceFromRanges`：逐格拆分、跳过空内容、按 id 去重、可撤销。
   */
  const handleWorkspaceBlankClick = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const pressed = workspacePressRef.current;
      workspacePressRef.current = null;

      const target = event.target as HTMLElement | null;
      if (
        target?.closest(
          '[data-testid="workspace-item"], button, input, select, textarea, label, a, .ws-panel-head, .ws-panel-foot, .ws-splitter',
        )
      ) {
        return;
      }
      if (!pressed || Math.hypot(event.clientX - pressed.x, event.clientY - pressed.y) > MOVE_TOLERANCE_PX) return;
      if (controllerRef.current?.active) return;
      if (!QUICK_ADD_MODES.includes(modeRef.current)) {
        log('workspace:quick-add-skipped', { reason: 'mode', mode: modeRef.current });
        return;
      }

      const since = Date.now() - lastGridInteractionAtRef.current;
      if (lastGridInteractionAtRef.current === 0 || since > QUICK_ADD_WINDOW_MS) {
        log('workspace:quick-add-skipped', { reason: 'stale-selection', since });
        return;
      }

      const a1List = readSelectionA1List();
      if (a1List.length === 0) {
        log('workspace:quick-add-skipped', { reason: 'no-selection' });
        return;
      }

      const key = a1List.join(' ');
      const last = lastQuickAddRef.current;
      if (last && last.key === key && Date.now() - last.at <= QUICK_ADD_WINDOW_MS) {
        log('workspace:quick-add-duplicate', { a1: key });
        return;
      }

      const outcome = addWorkspaceFromRanges(a1List, { x: event.clientX, y: event.clientY });
      lastQuickAddRef.current = { key, at: Date.now() };
      log('workspace:quick-add', {
        a1: key,
        blocks: a1List.length,
        items: outcome.items.length,
        skipped: outcome.skipped,
        ok: outcome.ok,
      });
    },
    [addWorkspaceFromRanges, readSelectionA1List],
  );

  dropHandlerRef.current = (payload, target, pointer) => {
    const sheet = getSheet();
    if (!sheet) return;

    /** 1) 落到侧边栏 → 暂存为工作区条目。**必须区分拖拽来源**：不看 `payload.kind` 的话，工作区条目
     * 拖到面板空白处会被当成"从表格拖进来"，同一个快照再收一遍（新 id、同来源）＝自我复制 */
    if (isInsideSidebar(pointer)) {
      if (payload.kind === 'workspace-item') {
        toast('这个单元格已经在工作区里了；拖到表格里才会写回去');
        log('workspace:drop-self-ignored', { id: payload.snapshot.id, a1: payload.snapshot.source.a1 });
        return;
      }
      const staged = addWorkspaceItem(payload.snapshot, pointer);
      // 用户设置：拖入后是否保留表格内容（false = 剪切语义）
      if (!settingsRef.current.keepSourceOnDrop) {
        const src = payload.snapshot.source;
        /** 只有真的放进去了才动源内容：全空选区会被整块跳过，此时清空源会是"东西没进来、格子却空了" */
        if (!staged.ok) {
          log('workspace:cut-skipped', { a1: src.a1, reason: 'nothing-staged' });
          return;
        }
        beginSheetAction(`剪切 ${src.a1} 到工作区`, 'swap');
        const result = clearRange(sheet, src.a1);
        if (result.ok) {
          pushHistory(`剪切 ${src.a1} 到工作区`, 'swap');
          toast(`已剪切 ${src.a1} 到工作区（源内容已清空，格式保留）`);
          log('workspace:cut-by-drag', { a1: src.a1 });
        } else {
          toast(result.reason ?? '源内容清空失败（内容已复制到工作区）', 'warn');
        }
      } else {
        toast(`已暂存 ${payload.snapshot.source.a1} 到工作区（原内容保留）`);
      }
      return;
    }

    // 2) 没有有效落点
    if (!target) {
      toast('没有落在有效单元格上，操作已取消', 'warn');
      return;
    }

    const snapshot = payload.snapshot;
    const targetA1 = rectToA1({
      startRow: target.row,
      startColumn: target.col,
      endRow: target.row + snapshot.rows - 1,
      endColumn: target.col + snapshot.cols - 1,
    });

    // 3) 工作区条目拖回表格 → 只写内容，保留目标格式
    if (payload.kind === 'workspace-item') {
      /**
       * 写回表格是表格侧的内容编辑，必须走表格通道（beginSheetAction 预置措辞 → 命令 →
       * pushHistory 认领 Univer 那一步）；记成 'workspace' 会让 stepBack 找不到快照、Ctrl+Z 永久失灵。
       */
      const label = `写入工作区条目 ${snapshot.label}`;
      beginSheetAction(label, 'edit');
      applySnapshot(sheet, snapshot, { row: target.row, col: target.col });
      pushHistory(label, 'edit');
      toast(`已写入 ${snapshot.label} → ${targetA1}`);
      log('app:workspace-paste', { from: snapshot.label, to: targetA1 });
      // 用户设置：写回后是否从工作区移除该条目
      if (settingsRef.current.removeItemAfterPaste) {
        removeItem(snapshot.id);
        log('workspace:remove-after-paste', { id: snapshot.id, label: snapshot.label });
      }
      return;
    }

    // 4) 表格内拖动
    const src = snapshot.source;
    const rows = src.endRow - src.startRow + 1;
    const cols = src.endCol - src.startCol + 1;
    const overlap =
      target.row <= src.endRow &&
      target.row + rows - 1 >= src.startRow &&
      target.col <= src.endCol &&
      target.col + cols - 1 >= src.startCol;

    if (overlap) {
      toast('目标区域与源区域重叠，操作已取消', 'warn');
      return;
    }

    const sameSize = rows === snapshot.rows && cols === snapshot.cols;
    if (sameSize) {
      beginSheetAction(`互换 ${src.a1} ⇄ ${targetA1}`, 'swap');
      const result = swapRanges(sheet, src.a1, targetA1);
      if (result.ok) {
        pushHistory(`互换 ${src.a1} ⇄ ${targetA1}`, 'swap');
        toast(`已互换 ${src.a1} ⇄ ${targetA1}（样式保持不变）`);
        log('app:swap', { a: src.a1, b: targetA1 });
        flashSwapRef.current(src.a1, targetA1); // 取消选中 + 黄色提醒框
      } else {
        toast(result.reason ?? '互换失败', 'warn');
      }
    } else {
      beginSheetAction(`移动 ${src.a1} → ${targetA1}`, 'swap');
      const result = moveRange(sheet, src.a1, { row: target.row, col: target.col });
      if (result.ok) {
        toast(`已移动 ${src.a1} → ${targetA1}`);
        pushHistory(`移动 ${src.a1} → ${targetA1}`, 'swap');
        log('app:move', { from: src.a1, to: targetA1 });
      } else {
        toast(result.reason ?? '移动失败', 'warn');
      }
    }
  };

  // ---------------------------------------------------------------- 引导
  useEffect(() => {
    setReducedMotion(window.matchMedia('(prefers-reduced-motion: reduce)').matches);

    /** **热更新重启**（HMR / React Fast Refresh）的识别与处理：Fast Refresh 会重跑本 effect，而
     * refs/state 保留、旧实例已在 cleanup 里 dispose → tabsDataRef 里还记着标签但没有对应 unit
     * （"灵魂标签"：看得见点不动），且启动恢复会把同 id 标签再追加一次。
     * 处理：把所有标签标成**未实体化**（切回按字节重建），恢复只按 id 补充缺失标签 */
    const hotReboot = tabsDataRef.current.length > 0 || tabRuntimeRef.current.length > 0;
    if (hotReboot) {
      tabRuntimeRef.current = tabRuntimeRef.current.map((entry) => ({ ...entry, built: false }));
      log('app:hot-reboot', { tabs: tabsDataRef.current.length, active: activeTabIdRef.current });
    }

    let boot: UniverBoot | null = null;
    const cleanups: Array<() => void> = [];

    /**
     * 填充柄的**类级**关闸，必须在第一个工作簿（示例表）创建之前装好：启动时 loadWorkbook →
     * attachSheetDeps 连着做，那一刻渲染单元还没建出来，"按单元关"会落空，示例表右下角仍会画着小方块。
     */
    log(installFillHandleOff() ? 'ui:fill-handle-off-prototype' : 'ui:fill-handle-off-prototype-missing');

    try {
      boot = bootUniver(CONTAINER_ID);
      apiRef.current = boot.univerAPI as FUniver;

      // 第一道防线：命令层"默认拒绝"闸门（只放行纯内容编辑与不落盘的操作）。
      // 必须在任何业务动作之前装好，且只需装一次（命令服务是全局单例）。
      const rootInjector = (boot.univer as unknown as { __getInjector?: () => { get: (token: unknown) => unknown } }).__getInjector?.();
      rootInjectorRef.current = rootInjector ?? null;
      if (rootInjector) {
        guardRef.current = installReadOnlyGuard(rootInjector);
        instanceServiceRef.current = rootInjector.get(IUniverInstanceService) as IUniverInstanceService;

        // 撤销/重做可用性：订阅官方状态（undos/redos 计数），用于工具栏按钮的禁用态
        try {
          const undoRedo = rootInjector.get(IUndoRedoService) as {
            undoRedoStatus$: { subscribe: (cb: (status: { undos: number; redos: number }) => void) => { unsubscribe: () => void } };
            clearUndoRedo?: (unitId: string) => void;
          };
          undoRedoServiceRef.current = undoRedo;
          const subscription = undoRedo.undoRedoStatus$.subscribe((status) => {
            const undos = status?.undos ?? 0;
            const redos = status?.redos ?? 0;
            setUndoRedoCounts({ undos, redos });
            undoRedoCountsRef.current = { undos, redos };

            /** 历史记账：undos 增加 = 表格侧新的一步（打字等非我们发起的编辑靠这里入账）。不能用 undos
             * 回写 historyIndex（账本里还有工作区动作，条目数与 Univer 步数不再一一对应），撤销/重做
             * 一律走 stepBack/stepForward */
            const previous = prevUndosRef.current;
            if (undos > previous) {
              /** 装载窗口内的增长不属于用户操作，不入账；只更新计数，窗口关闭后新步骤仍能识别 */
              if (pauseAutoEntriesRef.current) {
                prevUndosRef.current = undos;
                return;
              }
              /**
               * 我们自己发起的撤销/重做也会让 undos 变化（重做即 +1）；不跳过的话，刚重做出来的那一步
               * 会被自动入账的"编辑内容"顶掉。状态同步发出，"置 1 → 命令 → 清 0"窗口足够精确。
               */
              if (suppressAutoEntriesRef.current > 0) {
                suppressAutoEntriesRef.current = 0;
                prevUndosRef.current = undos;
                return;
              }
              const added = undos - previous;
              /** 标签只认刚打上的那一个（超过 2 秒的预置标签不能串到用户的下一次编辑上） */
              const marked = pendingHistoryLabelRef.current;
              const pending =
                marked && Date.now() - marked.at < 2000
                  ? { label: marked.label, kind: marked.kind }
                  : { label: '编辑内容', kind: 'edit' as HistoryKind };
              pendingHistoryLabelRef.current = null;
              // 同样写**同步账本**（撤销可能紧接着就来）
              let next = entriesRef.current.slice(0, historyIndexRef.current);
              for (let i = 0; i < added; i += 1) {
                const autoEntry: HistoryEntry = {
                  id: `h-${Date.now()}-${i}-${Math.random().toString(36).slice(2, 7)}`,
                  label: pending.label,
                  kind: pending.kind,
                  at: Date.now(),
                  scope: 'sheet',
                };
                // 记账 id 交给我们自己的表格动作认领（见 beginSheetAction / pushHistory）
                autoSheetEntriesRef.current.push(autoEntry.id);
                next = appendEntry(next, autoEntry, 200);
              }
              // 认领只关心"最近这一小段"，超出的旧 id 没有用了（防止无界增长）
              if (autoSheetEntriesRef.current.length > 400) {
                autoSheetEntriesRef.current = autoSheetEntriesRef.current.slice(-400);
              }
              entriesRef.current = next;
              historyIndexRef.current = next.length;
              setHistoryEntries(next);
              setHistoryIndex(next.length);
            }
            prevUndosRef.current = undos;
          });
          cleanups.push(() => subscription.unsubscribe());
        } catch (error) {
          log('app:undo-status-subscribe-error', { message: String(error) });
        }

        log('app:read-only-guard-installed');
      } else {
        log('app:read-only-guard-missing-injector');
      }

      loadWorkbook(boot.univerAPI, createSampleWorkbook());
      focusSheetUnit(SAMPLE_WORKBOOK_ID);
      attachSheetDeps();

      const container = containerRef.current;
      if (!container) throw new Error('预览容器不存在');

      /**
       * 像素 → 单元格（自建命中测试）。不用 Univer 的 `CellPointerMove`：非选择模式下我们会屏蔽
       * pointermove，它不会再派发；渲染服务的 `getCellWithCoordByOffset` 也不适用。
       */
      /** 提醒框定位用的 scene/skeleton 缓存：声明在组件级（释放单元时要清），见那里的注释 */
      /** 互换后"补清选区"的 rAF 句柄 */
      const selectionClearRafRef = { current: null as number | null };
      /** 互换后"补清选区"的兜底定时器句柄 */
      const selectionClearTimerRef = { current: null as number | null };
      /** "补清选区"拥有的格子（只有这些格子被晚到的选区写入重新选中时我们才去清） */
      const pendingClearOwnedRef = { current: new Set<string>() };
      /** 正在进行的滚动条拖拽（自己实现，见 scrollbar-drag.ts） */
      const scrollDragRef = {
        current: null as null | { axis: 'y' | 'x'; geo: AxisGeometry; grabOffset: number; lastScroll: number },
      };
      /** 拖滚动条时的高亮层（可视反馈，松手即收） */
      const scrollbarGrabRef = { current: null as HTMLElement | null };

      const hitTestCell = (x: number, y: number): { row: number; col: number; sheetId: string } | null => {
        const sheet = getSheet();
        if (!sheet) return null;
        const canvas = activeCanvas();
        if (!canvas) return null;
        const rect = canvas.getBoundingClientRect();
        if (x < rect.left || x > rect.right || y < rect.top || y > rect.bottom) return null;

        /**
         * **自建换算**：上游的 `getCellWithCoordByOffset` 在 scale=1 时会把滚动量消掉，滚过之后落点
         * 会解析成"没滚动时"的格子。这里自己算：css / scale + scroll，再在行列累计偏移里二分查找；
         * 与提醒框的正向换算严格互逆（正向见 `src/interaction/swap-flash.ts`）。
         */
        try {
          const unitId = apiRef.current?.getActiveWorkbook()?.getId();
          if (!unitId) return null;
          let skeleton = renderCacheRef.current?.unitId === unitId ? renderCacheRef.current.skeleton : undefined;
          if (!skeleton) {
            const renderManager = rootInjectorRef.current?.get(IRenderManagerService) as unknown as {
              getRenderById?: (id: string) => { with?: <T>(token: unknown) => T } | null;
            } | null;
            const render = renderManager?.getRenderById?.(unitId) ?? null;
            const skeletonManager = render?.with?.(SheetSkeletonManagerService) as unknown as {
              getCurrentSkeleton?: () => unknown;
            } | null;
            skeleton = skeletonManager?.getCurrentSkeleton?.() ?? null;
            renderCacheRef.current = { unitId, scene: renderCacheRef.current?.scene ?? null, skeleton };
          }
          const typed = skeleton as {
            columnWidthAccumulation?: number[];
            rowHeightAccumulation?: number[];
            rowHeaderWidthAndMarginLeft?: number;
            columnHeaderHeightAndMarginTop?: number;
            scrollX?: number;
            scrollY?: number;
            scaleX?: number;
            scaleY?: number;
          } | null;
          const main = getMainViewport();
          const scene = main?.scene as { getAncestorScale?: () => { scaleX: number; scaleY: number } } | null | undefined;
          const scale = scene?.getAncestorScale?.() ?? { scaleX: typed?.scaleX ?? 1, scaleY: typed?.scaleY ?? 1 };
          const scroll = main?.viewport ? getMainViewportScroll() : { x: typed?.scrollX ?? 0, y: typed?.scrollY ?? 0 };

          const cell = cellAtPoint(
            {
              columnOffsets: typed?.columnWidthAccumulation ?? [],
              rowOffsets: typed?.rowHeightAccumulation ?? [],
              headerWidth: typed?.rowHeaderWidthAndMarginLeft ?? 0,
              headerHeight: typed?.columnHeaderHeightAndMarginTop ?? 0,
              scroll,
              scale: { x: scale.scaleX, y: scale.scaleY },
            },
            x - rect.left,
            y - rect.top,
          );
          if (!cell) return null;
          return { row: cell.row, col: cell.col, sheetId: sheet.getSheetId() };
        } catch (error) {
          log('drag:hit-test-error', { message: String(error) });
          return null;
        }
      };

      /** 清空选区：互换/搬运完成后不该继续留着选中的单元格 */
      const clearCellSelection = (): void => {
        const sheet = getSheet();
        if (!sheet) return;
        try {
          const service = sheet.getInject().get(SheetsSelectionsService) as unknown as {
            clearCurrentSelections?: () => void;
          } | null;
          service?.clearCurrentSelections?.();
        } catch (error) {
          log('selection:clear-error', { message: String(error) });
        }
      };
      clearSelectionRef.current = clearCellSelection;

      /**
       * 再清几次选区（三帧 + 一次延时兜底）：选区写入是异步命令，只清当帧会漏（用户看到"互换后还
       * 选中着最开始那一格"）。但要分清是谁的选区：只收拾这次互换带出来的源格/目标格——无条件清会
       * 抹掉用户在 160ms 内新选的区域；用户一按下指针就立刻收手。
       */
      const clearSelectionSoon = (cells: string[] = []): void => {
        // 累积"这次收尾涉及哪几格"：一次拖拽收尾可能分两步调用，后一次不能覆盖前一次的所有权
        // （否则 Univer 晚到的选区写入就没人收拾了，实测踩过）
        cells.filter(Boolean).forEach((a1) => pendingClearOwnedRef.current.add(a1));
        cancelSelectionClears(false);
        // 第一下**立即清**：这次手势（互换/搬运）刚结束，不该留选中
        clearCellSelection();
        const owned = pendingClearOwnedRef.current;
        let framesLeft = 3;
        const step = (): void => {
          selectionClearRafRef.current = null;
          // 后续几次只为收拾"Univer 异步把互换那两格又写回选中"；选区若已变成别处就立刻收手
          if (owned.size === 0 || !selectionBelongsToGesture(owned)) {
            owned.clear();
            return;
          }
          clearCellSelection();
          framesLeft -= 1;
          if (framesLeft > 0) selectionClearRafRef.current = window.requestAnimationFrame(step);
        };
        if (owned.size > 0) selectionClearRafRef.current = window.requestAnimationFrame(step);
        selectionClearTimerRef.current = window.setTimeout(() => {
          selectionClearTimerRef.current = null;
          if (owned.size === 0) return;
          if (!selectionBelongsToGesture(owned)) {
            owned.clear();
            return;
          }
          clearCellSelection();
        }, 160);
      };

      /** 当前选区是不是"这次手势带出来的"（互换涉及的那几格） */
      const selectionBelongsToGesture = (owned: Set<string>): boolean => {
        try {
          const current = getSheet()?.getActiveRange()?.getA1Notation() ?? null;
          return Boolean(current && owned.has(current));
        } catch {
          return false;
        }
      };

      /** 取消所有"补清选区"的定时器/帧回调（用户开始新的操作时调用；同时放弃所有权） */
      const cancelSelectionClears = (dropOwnership = true): void => {
        if (selectionClearRafRef.current !== null) {
          window.cancelAnimationFrame(selectionClearRafRef.current);
          selectionClearRafRef.current = null;
        }
        if (selectionClearTimerRef.current !== null) {
          window.clearTimeout(selectionClearTimerRef.current);
          selectionClearTimerRef.current = null;
        }
        if (dropOwnership) pendingClearOwnedRef.current.clear();
      };

      /**
       * 单元格 → **视口**矩形（黄色提醒框用）：`getCellWithCoordByIndex` 给内容坐标 → 用 scene 的
       * 滚动/缩放换算到画布 CSS 像素 → 加上画布位置（见 `src/interaction/swap-flash.ts`）。
       * 渲染实例与服务按 unitId 缓存（每帧调用，不能反复查 DI）。
       */
      const measureFlashBox = (row: number, col: number): FlashBox | null => {
        const canvas = activeCanvas();
        const unitId = apiRef.current?.getActiveWorkbook()?.getId();
        if (!canvas || !unitId) return null;
        try {
          let scene = renderCacheRef.current?.unitId === unitId ? renderCacheRef.current.scene : undefined;
          let skeleton = renderCacheRef.current?.unitId === unitId ? renderCacheRef.current.skeleton : undefined;
          if (!scene || !skeleton) {
            const renderManager = rootInjectorRef.current?.get(IRenderManagerService) as unknown as {
              getRenderById?: (id: string) => {
                scene?: unknown;
                with?: <T>(token: unknown) => T;
              } | null;
            } | null;
            const render = renderManager?.getRenderById?.(unitId) ?? null;
            const skeletonManager = render?.with?.(SheetSkeletonManagerService) as unknown as {
              getCurrentSkeleton?: () => unknown;
            } | null;
            scene = render?.scene ?? null;
            skeleton = skeletonManager?.getCurrentSkeleton?.() ?? null;
            renderCacheRef.current = { unitId, scene, skeleton };
          }
          if (!scene || !skeleton) return null;

          const typedScene = scene as {
            getViewport?: (key: unknown) => unknown;
            getViewportScrollXY?: (viewport: unknown) => { x: number; y: number };
            getAncestorScale?: () => { scaleX: number; scaleY: number };
          };
          const typedSkeleton = skeleton as {
            getCellWithCoordByIndex?: (row: number, col: number, header?: boolean) => {
              startX: number;
              startY: number;
              endX: number;
              endY: number;
            } | null;
            scrollX: number;
            scrollY: number;
            scaleX: number;
            scaleY: number;
          };
          const viewport = typedScene.getViewport?.(SHEET_VIEWPORT_KEY.VIEW_MAIN);
          const scroll =
            (viewport ? typedScene.getViewportScrollXY?.(viewport) : null) ?? {
              x: typedSkeleton.scrollX,
              y: typedSkeleton.scrollY,
            };
          const scale = typedScene.getAncestorScale?.() ?? { scaleX: typedSkeleton.scaleX, scaleY: typedSkeleton.scaleY };

          const box = computeFlashRect(
            {
              cellRect: (r, c) => typedSkeleton.getCellWithCoordByIndex?.(r, c, true) ?? null,
              scroll,
              scale: { x: scale.scaleX, y: scale.scaleY },
            },
            row,
            col,
          );
          if (!box) return null;

          const rect = canvas.getBoundingClientRect();
          // 不在可见区域内就不画（否则提醒框会飘到表格外面）
          const left = Math.max(box.left, 0);
          const top = Math.max(box.top, 0);
          const right = Math.min(box.left + box.width, rect.width);
          const bottom = Math.min(box.top + box.height, rect.height);
          if (right - left < 2 || bottom - top < 2) return null;
          return {
            left: rect.left + left,
            top: rect.top + top,
            width: right - left,
            height: bottom - top,
          };
        } catch (error) {
          log('swap:flash-measure-error', { message: String(error) });
          return null;
        }
      };

      /**
       * 互换完成后的"落点提醒"：取消普通选中 + 把换过的两格用黄色框标出（渐隐、自动消失、不挡操作）。
       * 两条走不通的路（别再回去）：往 `setSelections` 的 style 塞自定义样式无效；`highlightRanges()`
       * 做不了渐隐且每帧重绘画布。现在自画 DOM 浮层 + CSS 透明度动画 + `pointer-events: none`。
       */
      const flashSwappedRanges = (firstA1: string, secondA1?: string): void => {
        const sheet = getSheet();
        if (!sheet) return;
        clearCellSelection();
        // 选区写入是异步命令，pointerup 后的 click 可能把它写回来，见 clearSelectionSoon
        clearSelectionSoon([firstA1, secondA1].filter((a1): a1 is string => Boolean(a1)));
        try {
          const targets: FlashTarget[] = [firstA1, secondA1]
            .filter((a1): a1 is string => Boolean(a1))
            .map((a1) => {
              const rect = sheet.getRange(a1).getRange();
              return { a1, row: rect.startRow, col: rect.startColumn };
            });
          // 自检：矩形中心做一次命中测试，落不回原格就说明坐标换算变了（宁可日志暴露，也不画错位的框）
          const mismatch = targets.filter((target) => {
            const box = measureFlashBox(target.row, target.col);
            if (!box) return false;
            const hit = hitTestCell(box.left + box.width / 2, box.top + box.height / 2);
            return !hit || hit.row !== target.row || hit.col !== target.col;
          });
          if (mismatch.length > 0) {
            log('swap:flash-rect-mismatch', { cells: mismatch.map((target) => target.a1) });
          }
          flashOverlayRef.current?.show(targets);
          log('swap:flash', { a: firstA1, b: secondA1 ?? null });
        } catch (error) {
          log('swap:flash-error', { message: String(error) });
        }
      };
      flashSwapRef.current = flashSwappedRanges;

      /** 某 A1 的矩形（测试钩子）；装配在这里而不是钩子里，钩子就不必再摸 sheet */
      const rectOfA1 = (a1: string): FlashBox | null => {
        const sheet = getSheet();
        if (!sheet) return null;
        try {
          const rect = sheet.getRange(a1).getRange();
          return measureFlashBox(rect.startRow, rect.startColumn);
        } catch {
          return null;
        }
      };


      /** 释放提醒框（用户在表格里按下指针时调用；到时间也会自动消失） */
      const clearSwapFlash = (): void => {
        if (!flashOverlayRef.current?.active) return;
        flashOverlayRef.current.hide();
        log('swap:flash-clear', {});
      };
      clearSwapFlashRef.current = clearSwapFlash;

      // 拖拽控制器
      controllerRef.current = createDragController(
        apiRef.current,
        { container },
        {
          onTargetChange: (target, h) => {
            /**
             * 落在**工作区**上方时不算"不能放置"：命中测试只认表格画布，指针在侧边栏上时控制器给 reject，
             * 用户会看到"此处不能放置"（其实能放）。这里按指针位置再判一次。
             */
            const pointer = lastPointerRef.current;
            const overSidebar = isInsideSidebar(pointer);
            const effective: DropHint | 'workspace' = overSidebar ? 'workspace' : h;
            setHint(effective);
            // 落点高亮：把"即将被交换/写入"的那一格框出来；落点无效时换红色虚线，
            // 指针在工作区上时没有网格落点，把高亮收掉
            targetHighlightRef.current?.show(
              target && !overSidebar
                ? [
                    {
                      a1: rectToA1({ startRow: target.row, startColumn: target.col, endRow: target.row, endColumn: target.col }),
                      row: target.row,
                      col: target.col,
                    },
                  ]
                : [],
              { variant: h === 'reject' ? 'reject' : 'normal' },
            );
            // 落点变化记日志：这既是调试信息，也是 e2e 断言"拖动中能解析出目标单元格"的依据
            const key = target ? `${target.row}:${target.col}:${h}` : `none:${h}`;
            if (lastTargetKeyRef.current !== key) {
              lastTargetKeyRef.current = key;
              log('drag:target', target ? { row: target.row, col: target.col, hint: h } : { hint: h });
            }
          },
          onDrop: (payload, target, pointer) => dropHandlerRef.current(payload, target, pointer),
          onDragStateChange: setDragging,
          // 自建命中：非选择模式下 pointermove 被屏蔽，Univer 的 CellPointerMove 不会派发
          hitTest: (x, y) => hitTestCell(x, y),
        },
      );
      cleanups.push(() => controllerRef.current?.dispose());

      // 互换提醒框浮层：自己画 DOM（渐隐 + 穿透），不碰 Univer 的选区/标记图层
      flashOverlayRef.current = createSwapFlashOverlay({ measure: (target) => measureFlashBox(target.row, target.col) });
      // 拖动落点高亮：同一定位逻辑，但常亮不淡出（松手/取消时收掉）
      targetHighlightRef.current = createDragTargetHighlight({ measure: (target) => measureFlashBox(target.row, target.col) });
      cleanups.push(() => {
        flashOverlayRef.current?.dispose();
        flashOverlayRef.current = null;
        targetHighlightRef.current?.dispose();
        targetHighlightRef.current = null;
        // 滚动条"抓取高亮"节点直接挂在 document.body 上：指针若正悬在条带上（hover 态不会自己消失），
        // 卸载后会变成永久游离节点，这里兜底摘掉
        hideScrollbarGrab();
        if (selectionClearRafRef.current !== null) window.cancelAnimationFrame(selectionClearRafRef.current);
        selectionClearRafRef.current = null;
        if (selectionClearTimerRef.current !== null) window.clearTimeout(selectionClearTimerRef.current);
        selectionClearTimerRef.current = null;
      });

      // 点击交换：状态机 + 动画原语（与拖动模式互斥，由 modeRef 控制）
      swapAnimRef.current = createSwapAnimation({ reducedMotion: window.matchMedia('(prefers-reduced-motion: reduce)').matches });
      clickSwapRef.current = createClickSwapController({
        onSelectionChange: (sides) => {
          /** 把"待互换的那一方"同步到界面：工作区条目亮起（`is-pending`），提示"接着点表格里的格子即可互换" */
          const first = sides[0] ?? null;
          setPendingSwapItemId(first?.kind === 'workspace' ? (first.itemId ?? null) : null);
          setPendingSwapLabel(first?.label ?? null);
          if (first) log('swap:selected', { label: first.label, kind: first.kind });
          else log('swap:selection-cleared');
        },
        onSwap: (a, b) => swapHandlerRef.current(a, b),
        onReject: (reason) => toast(reason, 'warn'),
      });
      clickSwapRef.current.setMode(modeRef.current);
      cleanups.push(() => {
        clickSwapRef.current?.dispose();
        swapAnimRef.current?.dispose();
      });

      // 点击模式下：单击单元格即"选中一方"
      const cellClickedKey = (apiRef.current.Event as unknown as Record<string, string>).CellClicked;
      if (cellClickedKey) {
        const disposable = apiRef.current.addEvent(cellClickedKey as never, (params: unknown) => {
          if (modeRef.current !== 'click-swap') return;
          const p = (params ?? {}) as { row?: number; column?: number; worksheet?: FWorksheet };
          if (typeof p.row !== 'number' || typeof p.column !== 'number') return;
          const sheet = p.worksheet ?? getSheet();
          if (!sheet) return;

          /**
           * 带 Shift/Ctrl 的点击是"框选"（见 onPointerDown）：框出单格 = 用户只是在定位，不当选边；
           * 框出一片区域 = 要互换这块区域，照常登记为一方（否则区域互换不可达）。
           */
          if (modifierPressRef.current) {
            const rect = sheet.getActiveRange()?.getRange();
            const single = !rect || (rect.startRow === rect.endRow && rect.startColumn === rect.endColumn);
            if (single) return;
          }

          const active = sheet.getActiveRange()?.getRange();
          const rows = active ? active.endRow - active.startRow + 1 : 1;
          const cols = active ? active.endColumn - active.startColumn + 1 : 1;
          const a1 = active
            ? rectToA1({ startRow: active.startRow, startColumn: active.startColumn, endRow: active.endRow, endColumn: active.endColumn })
            : rectToA1({ startRow: p.row, startColumn: p.column, endRow: p.row, endColumn: p.column });
          const side: SwapSide = {
            kind: 'cell',
            sheetId: sheet.getSheetId(),
            a1,
            rows,
            cols,
            label: `${sheet.getSheetName()}!${a1}`,
          };
          clickSwapRef.current?.select(side);
          swapAnimRef.current?.pulseSelection(lastPointerRef.current);
        });
        cleanups.push(() => (disposable as { dispose: () => void }).dispose());
      }

      /**
       * 指针接线：直接拖拽（无长按）+ 非选择模式下屏蔽原生扩选。
       * 选择模式完全交给 Univer（可框选）；拖拽模式按住即可搬运/互换、点击互换模式点选边，两者都不出现
       * 多选框。实现上 `pointerdown` 不拦截，但从按下起的 `pointermove` 在非选择模式拦掉，Univer 就没机会
       * 扩成一片；代价是 `CellPointerMove` 不再派发，落点命中改由 `hitTestCell` 负责。
       */
      const pressRef = { current: null as null | { x: number; y: number; onGrid: boolean; dragging: boolean; modifier: boolean } };

      /**
       * 这一次按压是不是"落在表格网格上"。**只认主画布**（`univer-sheet-main-canvas`）：
       * 用 `container.contains(target)` 会把缩放条/工作表标签/公式栏/编辑器都算进去，按住它们拖动会被
       * 当成扩选拦掉（表现为"右下角放大缩小失效"）；行/列表头是另外的画布，拖表头仍是原生的选行/选列。
       */
      const isGridTarget = (target: EventTarget | null): boolean => {
        if (!(target instanceof HTMLCanvasElement)) return false;
        return target.id.startsWith('univer-sheet-main-canvas');
      };

      /** 取活动工作簿主视口的 scene / viewport（滚动条命中与滚动量用）。每次现取：切标签/切表都会换渲染实例 */
      const getMainViewport = (): {
        scene: {
          getViewport?: (key: unknown) => unknown;
          getViewportScrollXY?: (viewport: unknown) => { x: number; y: number };
        };
        viewport: unknown;
      } | null => {
        const unitId = apiRef.current?.getActiveWorkbook()?.getId();
        if (!unitId) return null;
        try {
          const renderManager = rootInjectorRef.current?.get(IRenderManagerService) as unknown as {
            getRenderById?: (id: string) => { scene?: unknown } | null;
          } | null;
          const scene = renderManager?.getRenderById?.(unitId)?.scene as
            | { getViewport?: (key: unknown) => unknown; getViewportScrollXY?: (viewport: unknown) => { x: number; y: number } }
            | null
            | undefined;
          if (!scene) return null;
          return { scene, viewport: scene.getViewport?.(SHEET_VIEWPORT_KEY.VIEW_MAIN) ?? null };
        } catch (error) {
          log('scrollbar:viewport-error', { message: String(error) });
          return null;
        }
      };

      /** 主视口滚动量（测试钩子 `getScrollState` 用） */
      const getMainViewportScroll = (): { x: number; y: number } => {
        const main = getMainViewport();
        if (!main?.viewport) return { x: 0, y: 0 };
        try {
          return main.scene.getViewportScrollXY?.(main.viewport) ?? { x: 0, y: 0 };
        } catch {
          return { x: 0, y: 0 };
        }
      };

      /**
       * 滚动条在**页面**坐标里的位置（交互带）：滚动条画在主视口右缘/下缘，而画布比视口大，
       * 所以"画布最右边 4px"并不在滚动条上。用引擎的 `ScrollBar.pick()` 沿边缘扫出命中位置，
       * 不猜几何（track 的 rect 可能还没布局，算出来是 NaN）。
       */
      const getScrollbarBand = (): {
        vertical: { x: number; y: number; width: number; height: number } | null;
        horizon: { x: number; y: number; width: number; height: number } | null;
        verticalThumb: { x: number; y: number; height: number } | null;
        horizonThumb: { y: number; x: number; width: number } | null;
      } | null => {
        const canvas = activeCanvas();
        const main = getMainViewport();
        if (!canvas || !main?.viewport) return null;
        const viewport = main.viewport as { getScrollBar?: () => { pick?: (coord: unknown) => unknown } | null };
        const bar = viewport.getScrollBar?.() ?? null;
        if (!bar?.pick) return null;
        const rect = canvas.getBoundingClientRect();
        const hits = (x: number, y: number): boolean => Boolean(bar.pick!(new Vector2(x, y)));

        /** 沿一个方向扫出命中区间，取区间中点（边缘像素可能在交互区之外，取中点最稳） */
        const scanBand = (collect: (step: number) => boolean, limit: number): number | null => {
          const points: number[] = [];
          for (let step = 2; step <= limit; step += 2) {
            if (collect(step)) points.push(step);
          }
          if (points.length === 0) return null;
          return points[Math.floor(points.length / 2)];
        };

        const verticalInset = scanBand((dx) => hits(rect.width - dx, rect.height * 0.5), Math.min(200, rect.width / 2));
        const horizonInset = scanBand((dy) => hits(rect.width * 0.5, rect.height - dy), Math.min(200, rect.height / 2));
        if (verticalInset === null && horizonInset === null) return null;
        const barX = rect.left + (rect.width - (verticalInset ?? 8));
        const barY = rect.top + (rect.height - (horizonInset ?? 8));

        /** 滑块与轨道都命中滚动条，但只有拖滑块才会连续滚动；pick() 返回命中的 Rect，滑块高度最小，据此挑出 y 区间 */
        const findThumb = (
          sample: (step: number) => { height: number } | null,
          limit: number,
          origin: number,
          toPage: (value: number) => number,
        ): { start: number; end: number } | null => {
          const byHeight = new Map<number, number[]>();
          for (let step = 0; step <= limit; step += 4) {
            const picked = sample(step);
            if (!picked) continue;
            const key = Math.round(picked.height);
            byHeight.set(key, [...(byHeight.get(key) ?? []), step]);
          }
          if (byHeight.size === 0) return null;
          // track 贯穿整条，滑块高度最小 → 取最小高度那一组
          const smallest = [...byHeight.keys()].sort((a, b) => a - b)[0];
          const steps = byHeight.get(smallest) ?? [];
          if (steps.length === 0) return null;
          return { start: toPage(origin + Math.min(...steps)), end: toPage(origin + Math.max(...steps)) };
        };

        const verticalThumb = findThumb(
          (step) => bar.pick!(new Vector2(rect.width - (verticalInset ?? 8), step)) as { height: number } | null,
          rect.height,
          rect.top,
          (value) => value,
        );
        const horizonThumb = findThumb(
          (step) => bar.pick!(new Vector2(step, rect.height - (horizonInset ?? 8))) as { height: number } | null,
          rect.width,
          rect.left,
          (value) => value,
        );

        return {
          vertical:
            verticalInset === null ? null : { x: barX, y: rect.top, width: 2, height: rect.height * 0.8 },
          horizon: horizonInset === null ? null : { x: rect.left, y: barY, width: rect.width * 0.8, height: 2 },
          verticalThumb: verticalThumb ? { x: barX, y: verticalThumb.start, height: verticalThumb.end - verticalThumb.start } : null,
          horizonThumb: horizonThumb ? { y: barY, x: horizonThumb.start, width: horizonThumb.end - horizonThumb.start } : null,
        };
      };

      /**
       * **当前工作簿自己的那块主画布**：多工作簿时页面里有多块同名 canvas，`querySelector` 拿到的
       * 第一块可能属于别的（甚至已隐藏的）工作簿，画布坐标 ↔ 视口几何就全错。按活动 unit id 精确定位。
       */
      const activeCanvas = (): HTMLCanvasElement | null => {
        const unitId = apiRef.current?.getActiveWorkbook()?.getId();
        if (unitId) {
          const exact = container.querySelector(`canvas#univer-sheet-main-canvas_${CSS.escape(unitId)}`);
          if (exact instanceof HTMLCanvasElement) return exact;
        }
        const first = container.querySelector('canvas[id^="univer-sheet-main-canvas"]');
        return first instanceof HTMLCanvasElement ? first : null;
      };

      /**
       * 指针是不是落在滚动条上。非选择模式下我们拦掉 pointermove（阻止扩选/多选框），
       * 但滚动条就画在主画布上，不区分会连滚动条拖拽一起废掉。判定按条带走（见 `resolveScrollbarPress`）。
       */
      const isScrollbarHit = (x: number, y: number): boolean => {
        return resolveScrollbarPress(x, y) !== null;
      };

      /**
       * 解析"这次按下是不是在滚动条上"并给出拖动几何（见 `src/interaction/scrollbar-drag.ts`）。
       * ① `verticalScrollTrack` 这类 Rect 只有 width/height、没有可用的 x/y，所以用引擎的 `pick()`
       * 判断轴向与是否按在滑块上，并沿该轴扫出轨道 origin/length；
       * ② 判定按**条带**走（视口边缘往里 `SCROLLBAR_STRIP_PX` 到画布边缘），不要求精确压住 5~6px 的滚动条；
       * ③ 轨道/滑块长度/滚动上限优先用引擎给的数，拿不到就推算；④ 是否按在滑块上由滚动量推算。
       */
      const resolveScrollbarPress = (
        x: number,
        y: number,
      ): { axis: 'y' | 'x'; pointer: number; geo: AxisGeometry } | null => {
        const canvas = activeCanvas();
        const main = getMainViewport();
        if (!canvas || !main?.viewport) return null;
        const rect = canvas.getBoundingClientRect();
        const localX = x - rect.left;
        const localY = y - rect.top;
        if (localX < 0 || localY < 0 || localX > rect.width || localY > rect.height) return null;

        const viewport = main.viewport as {
          width?: number;
          height?: number;
          getScrollBar?: () => {
            limitX?: number;
            limitY?: number;
            verticalThumbSize?: number;
            horizontalThumbSize?: number;
            hasVerticalThumb?: () => boolean;
            hasHorizonThumb?: () => boolean;
          } | null;
        };
        const viewportW = viewport.width ?? rect.width;
        const viewportH = viewport.height ?? rect.height;
        const bar = viewport.getScrollBar?.() ?? null;

        const inVertical = isInBarStrip(localX, viewportW, rect.width, SCROLLBAR_STRIP_PX);
        const inHorizon = isInBarStrip(localY, viewportH, rect.height, SCROLLBAR_STRIP_PX);
        if (!inVertical && !inHorizon) return null;
        // 两条带都命中（右下角）时，取"离视口边缘更近"的那条
        const verticalDistance = Math.abs(localX - viewportW);
        const horizonDistance = Math.abs(localY - viewportH);
        const axis: 'y' | 'x' = inVertical && (!inHorizon || verticalDistance <= horizonDistance) ? 'y' : 'x';

        // 轨道：竖向沿画布高度、横向沿画布宽度（滚动条铺满视口那一侧）
        const trackOrigin = axis === 'y' ? rect.top - rect.top : 0; // 画布内坐标从 0 起
        const trackLength = Math.max(1, axis === 'y' ? viewportH : viewportW);
        const pointer = axis === 'y' ? localY : localX;

        const contentOffsets: 'rowHeightAccumulation' | 'columnWidthAccumulation' =
          axis === 'y' ? 'rowHeightAccumulation' : 'columnWidthAccumulation';
        const skeleton = renderCacheRef.current?.skeleton as Record<string, unknown> | null;
        const accumulated = (skeleton?.[contentOffsets] as number[] | undefined) ?? [];
        const header = (axis === 'y'
          ? (skeleton?.columnHeaderHeightAndMarginTop as number | undefined)
          : (skeleton?.rowHeaderWidthAndMarginLeft as number | undefined)) ?? 0;
        const viewportSize = axis === 'y' ? viewportH : viewportW;
        const contentSize = accumulated.length > 0 ? accumulated[accumulated.length - 1] + header : viewportSize;

        const engineLimit = (axis === 'y' ? bar?.limitY : bar?.limitX) ?? 0;
        const limit = engineLimit > 0 ? engineLimit : Math.max(0, contentSize - viewportSize);
        // 内容装得下就没有可滚的，不要抢这次按下（交给正常交互）
        if (limit <= 0) return null;

        const engineThumb = ((axis === 'y' ? bar?.verticalThumbSize : bar?.horizontalThumbSize) ?? 0) as number;
        const thumbLength = engineThumb > 0 ? engineThumb : Math.max(12, (trackLength * viewportSize) / Math.max(1, contentSize));

        return {
          axis,
          pointer,
          geo: { trackLength, thumbLength, limit, thumbStart: 0, trackOrigin },
        };
      };

      /** 拖滚动条时的可视反馈：DOM 浮层高亮（穿透、不吃事件），松手即收，不碰 Univer 画布 */
      const showScrollbarGrab = (x: number, y: number, axis: 'y' | 'x'): void => {
        const host = document.createElement('div');
        host.className = 'scrollbar-grab';
        host.dataset.axis = axis;
        document.body.appendChild(host);
        const box = document.createElement('div');
        box.className = 'scrollbar-grab-thumb';
        host.appendChild(box);
        const canvas = activeCanvas();
        const rect = canvas?.getBoundingClientRect();
        if (!rect) return;
        // 竖向：贴在画布右缘（指针那一列）；横向：贴在画布下缘（指针那一行）
        Object.assign(host.style, {
          position: 'fixed',
          left: axis === 'y' ? `${Math.round(x - 8)}px` : `${Math.round(rect.left)}px`,
          top: axis === 'y' ? `${Math.round(rect.top)}px` : `${Math.round(y - 8)}px`,
          width: axis === 'y' ? '16px' : `${Math.round(rect.width)}px`,
          height: axis === 'y' ? `${Math.round(rect.height)}px` : '16px',
        });
        scrollbarGrabRef.current?.remove();
        scrollbarGrabRef.current = host;
      };

      /** 拖拽中：让高亮跟着指针走（竖向只需更新 y，横向只需更新 x） */
      const moveScrollbarGrab = (x: number, y: number): void => {
        const host = scrollbarGrabRef.current;
        if (!host) return;
        if (host.dataset.axis === 'y') host.style.left = `${Math.round(x - 8)}px`;
        else host.style.top = `${Math.round(y - 8)}px`;
      };

      const hideScrollbarGrab = (): void => {
        scrollbarGrabRef.current?.remove();
        scrollbarGrabRef.current = null;
      };

      /** 悬停提示：滚动条只有 5~6px 宽、用户不知道哪里能抓，指针扫到条带就给一条淡高亮 */
      const updateScrollbarHover = (x: number, y: number): void => {
        if (scrollDragRef.current) return; // 拖动中由 showScrollbarGrab 管
        const canvas = activeCanvas();
        const main = getMainViewport();
        const rect = canvas?.getBoundingClientRect();
        if (!canvas || !rect || !main?.viewport) {
          hideScrollbarGrab();
          return;
        }
        const viewport = main.viewport as { width?: number; height?: number };
        const localX = x - rect.left;
        const localY = y - rect.top;
        const inside = localX >= 0 && localY >= 0 && localX <= rect.width && localY <= rect.height;
        const strip = resolveScrollbarPress(x, y);
        if (!inside || !strip) {
          hideScrollbarGrab();
          return;
        }
        if (scrollbarGrabRef.current?.dataset.axis !== strip.axis) {
          showScrollbarGrab(x, y, strip.axis);
          scrollbarGrabRef.current?.classList.add('is-hover');
        } else {
          moveScrollbarGrab(x, y);
        }
        void viewport;
      };

      /** 把目标滚动量落到视口（内容坐标单位，与 `viewportScrollY` 同尺度） */
      const applyScroll = (target: { x?: number; y?: number }): void => {
        const main = getMainViewport();
        if (!main?.viewport) return;
        const viewport = main.viewport as {
          scrollToViewportPos?: (pos: { viewportScrollX?: number; viewportScrollY?: number }) => void;
        };
        const current = getMainViewportScroll();
        try {
          viewport.scrollToViewportPos?.({
            viewportScrollX: target.x ?? current.x,
            viewportScrollY: target.y ?? current.y,
          });
        } catch (error) {
          log('scrollbar:apply-error', { message: String(error) });
        }
      };

      const onPointerDown = (e: PointerEvent) => {
        const onCanvas = isGridTarget(e.target);
        // 滚动条（或行列表头等非内容区）按下：**不能**按"内容拖动"处理，也不能拦 pointermove
        const scrollbar = onCanvas && isScrollbarHit(e.clientX, e.clientY);
        const onGrid = onCanvas && !scrollbar;
        if (scrollbar) {
          // 滚动条拖拽由我们接管：按住即可拖，滑块跟手（见 scrollbar-drag.ts）
          const drag = resolveScrollbarPress(e.clientX, e.clientY);
          if (drag) {
            const current = getMainViewportScroll();
            const currentScroll = drag.axis === 'y' ? current.y : current.x;
            // 抓滑块 → 记下抓取偏移，拖动时严格跟手；点轨道 → 先跳过去，再继续拖
            const begin = beginScrollDrag(drag.geo, drag.pointer, currentScroll, SCROLLBAR_GRAB_TOLERANCE_PX);
            scrollDragRef.current = {
              axis: drag.axis,
              geo: drag.geo,
              grabOffset: begin.grabOffset,
              lastScroll: begin.scroll,
            };
            if (begin.scroll !== currentScroll) applyScroll(drag.axis === 'y' ? { y: begin.scroll } : { x: begin.scroll });
            // 视觉反馈：拖拽期间在滚动条上盖一条高亮，让用户知道"抓住了"
            showScrollbarGrab(e.clientX, e.clientY, drag.axis);
            scrollbarGrabRef.current?.classList.remove('is-hover');
            log('scrollbar:drag-start', { axis: drag.axis, onThumb: begin.onThumb });
            // 不让引擎再处理这次按下（否则可能与我们的滚动叠加）
            e.stopPropagation();
            return;
          }
          log('scrollbar:press-unhandled', {});
        }
        // 在表格里按下 = 用户开始新操作 → 收提醒框、取消"补清选区"
        // （取消很重要：否则紧接着点的那一格会被补清逻辑抢掉）
        if (onGrid) {
          clearSwapFlashRef.current();
          cancelSelectionClears();
          // 表格拿到指针焦点 → 重新置 FOCUSING_SHEET（编辑器/公式栏会抢走焦点，
          // 不重新置位的话"编辑过一次单元格之后滚轮就不滚了"）
          focusSheetUnit(activeTabIdRef.current);
        }
        // 按住 Shift/Ctrl/Cmd = 用户显式要"原生选择"（例如点击互换模式下框出一个区域来互换）。
        // 这条规则让"不出现多选框"与"区域互换"两个需求同时成立
        const modifier = e.shiftKey || e.ctrlKey || e.metaKey;
        pressRef.current = { x: e.clientX, y: e.clientY, onGrid, dragging: false, modifier };
        modifierPressRef.current = modifier;
        if (!onGrid) return;
        lastPointerRef.current = { x: e.clientX, y: e.clientY };
      };

      /**
       * 拖动期间把选区"钉"在源区域上（逐帧 rAF 兜底）。用 queueMicrotask 抢不过 Univer——
       * 它的选区写入是异步命令，会排在微任务之后落地；rAF 在该帧所有微任务之后、绘制之前执行。
       */
      const restorePinnedSelection = (): void => {
        const a1 = dragSourceA1Ref.current;
        if (!a1) return;
        const sheet = getSheet();
        if (!sheet) return;
        try {
          if (sheet.getActiveRange()?.getA1Notation() === a1) return; // 已经是源区域，不必再写
          sheet.setActiveSelection(sheet.getRange(a1));
        } catch (error) {
          log('drag:pin-selection-error', { message: String(error) });
        }
      };

      const pinLoopStep = (): void => {
        pinRafRef.current = null;
        if (!controllerRef.current?.active) return; // 拖动结束，收工
        restorePinnedSelection();
        pinRafRef.current = window.requestAnimationFrame(pinLoopStep);
      };

      const startPinLoop = (): void => {
        if (pinRafRef.current !== null) return;
        pinRafRef.current = window.requestAnimationFrame(pinLoopStep);
      };

      const stopPinLoop = (): void => {
        if (pinRafRef.current === null) return;
        window.cancelAnimationFrame(pinRafRef.current);
        pinRafRef.current = null;
      };
      cleanups.push(stopPinLoop);

      const onPointerMove = (e: PointerEvent) => {
        lastPointerRef.current = { x: e.clientX, y: e.clientY };

        // 滚动条拖拽：自己算目标滚动位置（滑块跟手，见 scrollbar-drag.ts）
        const scrollDrag = scrollDragRef.current;
        if (scrollDrag) {
          const canvas = activeCanvas();
          if (canvas) {
            const rect = canvas.getBoundingClientRect();
            const pointer = scrollDrag.axis === 'y' ? e.clientY - rect.top : e.clientX - rect.left;
            const next = scrollForPointer(scrollDrag.geo, pointer, scrollDrag.grabOffset);
            if (next !== scrollDrag.lastScroll) {
              scrollDrag.lastScroll = next;
              applyScroll(scrollDrag.axis === 'y' ? { y: next } : { x: next });
            }
            moveScrollbarGrab(e.clientX, e.clientY);
          }
          e.stopPropagation();
          e.preventDefault();
          return;
        }

        const press = pressRef.current;
        // 非选择模式：从按下起就拦掉 pointermove，Univer 没机会把选区扩成多选框
        // （pointerdown 不拦，"点哪个格子选中哪个"仍是原生行为）；按了 Shift/Ctrl 则放行（用户要框选）
        if (press?.onGrid && !press.modifier && modeRef.current !== 'select') {
          e.stopPropagation();
          e.preventDefault();
        } else if (!press && e.target === activeCanvas()) {
          // 没按下时在画布上移动：顺手更新"滚动条可抓"的悬停提示
          updateScrollbarHover(e.clientX, e.clientY);
        }

        // 拖拽模式：按住后直接拖（无需长按）——位移超过容差就进入搬运
        if (
          press?.onGrid
          && !press.modifier
          && modeRef.current === 'drag'
          && !press.dragging
          && !controllerRef.current?.active
          && Math.hypot(e.clientX - press.x, e.clientY - press.y) > MOVE_TOLERANCE_PX
        ) {
          const sheet = getSheet();
          const snapshot = sheet ? extractSnapshot(sheet).snapshot : null;
          if (snapshot) {
            press.dragging = true;
            dragSourceA1Ref.current = snapshot.source.a1;
            controllerRef.current?.beginFromSheet(snapshot, { x: press.x, y: press.y });
            log('app:drag-start', { a1: snapshot.source.a1 });
          }
        }

        controllerRef.current?.updatePointer(e.clientX, e.clientY);
        if (controllerRef.current?.active) startPinLoop();

        const pending = pendingWorkspaceDragRef.current;
        if (pending && !pending.started) {
          if (Math.hypot(e.clientX - pending.x, e.clientY - pending.y) > MOVE_TOLERANCE_PX) {
            pending.started = true;
            controllerRef.current?.beginFromWorkspace(pending.snapshot, { x: pending.x, y: pending.y });
            log('app:workspace-drag-start', { a1: pending.snapshot.label });
          }
        }
      };
      const onPointerUp = (e: PointerEvent) => {
        lastPointerRef.current = { x: e.clientX, y: e.clientY };
        if (scrollDragRef.current) {
          log('scrollbar:drag-end', { axis: scrollDragRef.current.axis });
          scrollDragRef.current = null;
          hideScrollbarGrab();
          e.stopPropagation();
          return;
        }
        const wasDragging = pressRef.current?.dragging ?? false;
        const wasGridPress = pressRef.current?.onGrid ?? false;
        /** 记下"刚在表格里点过"，供快速加入判断窗口期；只认没拖动的表格按下（拖动是搬运语义） */
        if (wasGridPress && !wasDragging) lastGridInteractionAtRef.current = Date.now();
        pressRef.current = null;
        const pending = pendingWorkspaceDragRef.current;
        const wasWorkspaceDrag = pending?.started ?? false;
        if (pending && !pending.started) pendingWorkspaceDragRef.current = null;

        /**
         * 收尾顺序很关键：先停"钉选区"的逐帧循环、清掉源标记，**再** finish()。
         * 反过来的话，互换刚清干净的选区会被钉回源单元格（表现为"互换后还选中着最开始那一格"）。
         */
        stopPinLoop();
        dragSourceA1Ref.current = null;
        targetHighlightRef.current?.hide();
        if (controllerRef.current?.active) {
          controllerRef.current.finish(e.clientX, e.clientY);
        }
        // 真正拖动过 → 收尾清选区（补清逻辑会连清几帧并带一次延时兜底，用户一动指针即取消）
        if (wasDragging || wasWorkspaceDrag) clearSelectionSoon();

        /** 工作区条目上"按下 → 没拖动就松手" = 一次点击；点击互换模式下登记为待互换的一方 */
        const clickedItem = pendingWorkspaceClickRef.current;
        pendingWorkspaceClickRef.current = null;
        if (clickedItem && !wasWorkspaceDrag) {
          handleWorkspaceItemClickRef.current(clickedItem);
        }
      };

      window.addEventListener('pointerdown', onPointerDown, true);
      window.addEventListener('pointermove', onPointerMove, true);
      window.addEventListener('pointerup', onPointerUp, true);
      /** 键盘也算"在表格里动手"（快速加入的键盘版）；只在焦点确实落在表格容器里时才算 */
      const onKeyDownInSheet = (event: KeyboardEvent): void => {
        const container = containerRef.current;
        if (!container) return;
        const active = document.activeElement;
        const inside =
          (active instanceof Node && container.contains(active)) || (event.target instanceof Node && container.contains(event.target));
        if (inside) lastGridInteractionAtRef.current = Date.now();
      };
      window.addEventListener('keydown', onKeyDownInSheet, true);
      cleanups.push(() => {
        window.removeEventListener('pointerdown', onPointerDown, true);
        window.removeEventListener('pointermove', onPointerMove, true);
        window.removeEventListener('pointerup', onPointerUp, true);
        window.removeEventListener('keydown', onKeyDownInSheet, true);
      });

      // 活动工作表变化 → 重装格式锁与元信息
      const activeChangedKey = (apiRef.current.Event as unknown as Record<string, string>).ActiveSheetChanged;
      if (activeChangedKey) {
        const disposable = apiRef.current.addEvent(activeChangedKey as never, () => attachSheetDeps());
        cleanups.push(() => (disposable as { dispose: () => void }).dispose());
      }
      const selectionKey = (apiRef.current.Event as unknown as Record<string, string>).SelectionChanged;
      if (selectionKey) {
        const disposable = apiRef.current.addEvent(selectionKey as never, () => refreshSheetMeta());
        cleanups.push(() => (disposable as { dispose: () => void }).dispose());
      }

      installTestHooks({
        apiRef,
        lockRef,
        controllerRef,
        summaryRef,
        featureResultRef,
        guardRef,
        rootInjectorRef,
        getMainViewportScroll,
        getScrollbarBand,
        isScrollbarHit,
        measureFlashBox,
        rectOfA1,
        tabsRef: tabsDataRef,
        tabRuntimeRef,
        activeTabIdRef,
        sessionSaverRef,
        modeRef,
        snapshotActiveEditsRef,
        undoRedoCountsRef,
        pushHistory,
        beginSheetAction,
        commitWorkspace: (next, label, kind) => commitWorkspaceRef.current(next, label, kind),
        fillHandleState: () => {
          try {
            const unitId = apiRef.current?.getActiveWorkbook()?.getId();
            const renderManager = rootInjectorRef.current?.get(IRenderManagerService) as unknown as {
              getRenderById?: (id: string) => RenderUnitLike | null;
            } | null;
            const renderUnit = renderManager?.getRenderById?.(unitId ?? '') ?? null;
            return fillHandleStateOn(renderUnit, ISheetSelectionRenderService);
          } catch {
            return { available: false, controls: 0, enabled: [], visible: [] };
          }
        },
        stepBack: () => stepBackRef.current(),
        stepForward: () => stepForwardRef.current(),
        flashSwap: (a1: string, b1?: string) => flashSwapRef.current(a1, b1),
        itemsRef: { get: () => itemsStateRef.current },
        snapshotStoreRef,
        setItems,
        toast,
        getSheet,
        // 注意：它在本组件里声明得比这个 effect 晚。闭包延迟取值，运行时早就初始化好了。
        ensureActiveUnitRendered,
      });

      // ---------------------------------------------------------------- 文件拖放打开
      /**
       * 把 Excel 工作簿拖进窗口就打开（窗口级接管）：
       * 只处理带 `Files` 的拖放（内部内容拖拽不带 Files，互不干扰）；`dragover` 必须 preventDefault，
       * 否则浏览器不会触发 drop；支持一次拖多个（逐个开成标签页）。
       */
      const dragTypes = (event: DragEvent): string[] => {
        try {
          return Array.from(event.dataTransfer?.types ?? []);
        } catch {
          return [];
        }
      };
      const isFileDrag = (event: DragEvent): boolean => dragTypes(event).includes('Files');

      const onWindowDragOver = (event: DragEvent) => {
        if (!isFileDrag(event)) return;
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
        setFileDropActive(true);
      };
      const onWindowDragLeave = (event: DragEvent) => {
        if (!isFileDrag(event)) return;
        // 只有真正离开窗口才收起提示（在元素之间移动也会触发 dragleave）
        if (event.relatedTarget === null) setFileDropActive(false);
      };
      const onWindowDrop = (event: DragEvent) => {
        if (!isFileDrag(event)) return;
        event.preventDefault();
        setFileDropActive(false);
        const files = Array.from(event.dataTransfer?.files ?? []);
        if (files.length === 0) return;
        /**
         * 可接受范围跟形态走：本地版 + 本机 Excel 可用时 `.xlsb` 也能打开（先借 Excel 转换）。
         * 读 **ref**：引导 effect 只跑一次，闭包里的 state 永远是初值。
         */
        const acceptable = (name: string): boolean =>
          isSupportedWorkbookFile(name) || (bridgeStatusRef.current.available && isBridgeOnlyFile(name));
        const supported = files.filter((file) => acceptable(file.name));
        const rejected = files.filter((file) => !acceptable(file.name));
        log('file:drop', { count: files.length, supported: supported.length, names: files.map((f) => f.name) });
        if (rejected.length > 0) {
          toast(unsupportedFileMessage(rejected.map((file) => file.name)), 'warn');
        }
        // 逐个导入（内部会各自开一个标签页）；串行避免并发解析互相挤压
        void supported.reduce(
          (chain, file) => chain.then(() => new Promise<void>((resolve) => {
            fileOpenerRef.current(file);
            window.setTimeout(resolve, 400);
          })),
          Promise.resolve(),
        );
      };

      window.addEventListener('dragover', onWindowDragOver, true);
      window.addEventListener('dragleave', onWindowDragLeave, true);
      window.addEventListener('drop', onWindowDrop, true);
      cleanups.push(() => {
        window.removeEventListener('dragover', onWindowDragOver, true);
        window.removeEventListener('dragleave', onWindowDragLeave, true);
        window.removeEventListener('drop', onWindowDrop, true);
      });

      setStatus('ready');
      setStatusText(IDLE_STATUS_TEXT);
      log('app:ready');
    } catch (error) {
      log('app:boot-error', { message: String(error), stack: (error as Error)?.stack });
      setStatus('error');
      setStatusText(`启动失败：${String(error)}`);
    }

      // ---- 会话持久化：自动保存（去抖 + 指纹变化才写）与误关闭恢复 ----
      sessionStoreRef.current = createSessionStore();
      sessionSaverRef.current = createAutoSaver(sessionStoreRef.current, () => collectSessionState());
      const sessionTimer = window.setInterval(() => {
        if (restoringRef.current) return;

        // 打字等"不是我们发起"的编辑：脏单元格数增长且近期没有我们自己的入账 → 记一步"编辑内容"
        const activeIdForHistory = activeTabIdRef.current;
        const dirtyNow = activeIdForHistory ? dirtyCellCount(activeIdForHistory) : 0;
        if (dirtyNow > lastDirtySeenRef.current && Date.now() - lastHistoryPushAtRef.current > 1500) {
          pushHistory('编辑内容', 'edit');
        }
        lastDirtySeenRef.current = dirtyNow;

        const fingerprint = sessionFingerprint();
        if (fingerprint === lastSessionFingerprintRef.current) return;
        lastSessionFingerprintRef.current = fingerprint;
        sessionSaverRef.current?.trigger();
      }, 2000);
      const flushSession = () => {
        if (!restoringRef.current) void sessionSaverRef.current?.flush();
      };
      const onVisibilityChange = () => {
        if (document.hidden) flushSession();
      };
      window.addEventListener('pagehide', flushSession);
      // 右键菜单：捕获阶段监听，并在 DOM 层抑制 Univer 自带菜单
      const menuHost = containerRef.current;
      if (menuHost) {
        menuHost.addEventListener('contextmenu', handleContextMenu, true);
        cleanups.push(() => menuHost.removeEventListener('contextmenu', handleContextMenu, true));
      }
      document.addEventListener('visibilitychange', onVisibilityChange);
      cleanups.push(() => {
        window.clearInterval(sessionTimer);
        window.removeEventListener('pagehide', flushSession);
        document.removeEventListener('visibilitychange', onVisibilityChange);
        sessionSaverRef.current?.dispose();
      });

      // 启动恢复：重建每个标签（解析 → 适配 → 建工作簿 → 应用特性 → 回放编辑），再恢复工作区/模式/活动标签
      void (async () => {
        try {
          const state = await sessionStoreRef.current?.load();
          if (!state || state.tabs.length === 0) return;
          restoringRef.current = true;
          log('session:restore-start', { tabs: state.tabs.length, workspace: state.workspace?.length ?? 0 });
          if (state.workspace?.length) {
            /**
             * 恢复时清掉空条目（保证"工作区里绝对没有空条目"跨会话成立），并**按 id 去重**：
             * 同一个快照存两遍会变成同 id 的两张卡片 —— React key 冲突，删一张另一张也跟着消失。
             */
            const seen = new Set<string>();
            const restored = state.workspace.filter((item) => {
              if (isSnapshotEmpty(item)) return false;
              if (seen.has(item.id)) return false;
              seen.add(item.id);
              return true;
            });
            const dropped = state.workspace.length - restored.length;
            if (dropped > 0) log('workspace:skip-on-restore', { dropped, kept: restored.length });
            setItems(restored);
            // 恢复出来的条目也必须进快照表，否则拖拽写回/粘贴按 id 取不到（"条目不存在"）
            syncSnapshotStore(restored);
          }

          // **只实体化当前要看的那个标签**，其余先"冷"着（只记字节与编辑）：全量解析建簿的话，
          // 5 个百万格标签就是 5 份常驻内存；现在内存从 O(N) 变 O(1~K)。
          // **按 id 去重**：热更新重跑引导 effect 时内存里的标签还在（见 `app:hot-reboot`），
          // 无脑追加会出现同 id 同名的"灵魂标签"。这里只补会话里缺的那些。
          const known = new Set(tabsDataRef.current.map((item) => item.id));
          let alreadyKnown = 0;
          for (const tab of state.tabs) {
            if (known.has(tab.id)) {
              alreadyKnown += 1;
              continue;
            }
            known.add(tab.id);
            tabsDataRef.current = [
              ...tabsDataRef.current,
              { id: tab.id, fileName: tab.fileName, source: null, bytes: tab.originalBytes },
            ];
            editsByTabRef.current.set(
              tab.id,
              new Map(tab.edits.map((edit) => [`${edit.sheetId}|${edit.row}:${edit.col}`, edit])),
            );
            tabRuntimeRef.current = [...tabRuntimeRef.current, { id: tab.id, built: false, lastUsedAt: 0 }];
          }
          if (alreadyKnown > 0) {
            log('session:restore-skip-registered', { skipped: alreadyKnown, tabs: tabsDataRef.current.length });
          }
          log('session:tabs-registered', { tabs: tabsDataRef.current.length });

          syncTabs();
          const target =
            state.activeTabId && tabsDataRef.current.some((item) => item.id === state.activeTabId)
              ? state.activeTabId
              : (tabsDataRef.current[tabsDataRef.current.length - 1]?.id ?? null);
          if (target) activateTab(target);
          if (state.mode) {
            // 归一化：会话里可能是旧版本写的值或脏值，收敛到当前合法的三种模式之一
            const restoredMode = normalizeInteractionMode(state.mode);
            modeRef.current = restoredMode;
            setMode(restoredMode);
            clickSwapRef.current?.setMode(restoredMode);
          }
          // 用户设置（旧会话没有此字段时 migrateOrNull 已用默认值补齐）
          setSettings(normalizeSettings(state.settings));
          setStatus('ready');
          log('session:restore-done', { tabs: tabsDataRef.current.length });
        } catch (error) {
          log('session:restore-error', { message: String(error) });
        } finally {
          restoringRef.current = false;
          lastSessionFingerprintRef.current = sessionFingerprint();
        }
      })();

      const logTimer = window.setInterval(() => setEntries([...p0log]), 400);
    return () => {
      window.clearInterval(logTimer);
      /** ① **先把现场落盘**（必须在 `boot.dispose()` 与自动保存器 dispose 之前）：HMR 重跑本 effect 时
       * 页面并不卸载、没有 `pagehide`，而自动保存是"2 秒轮询 + 800ms 去抖"，最近一两秒的编辑还在内存里 */
      if (!restoringRef.current) {
        try {
          const state = collectSessionState();
          void sessionStoreRef.current?.save(state);
        } catch (error) {
          log('session:teardown-save-error', { message: String(error) });
        }
      }
      // ② 再拆监听/定时器（这一步会 dispose 自动保存器，所以上面的写入不能排在它后面）
      cleanups.forEach((fn) => fn());
      lockRef.current?.restore();
      guardRef.current?.restore();
      /**
       * ③ 清掉"按 unitId 记账"的全局账本，并摘掉挂在 window 上的测试钩子：
       * 插图 blob url 要随实例回收；`dirtyByWorkbook` 是模块级的，不清会被下一个实例继承
       * （示例/占位单元用固定 id，最容易被继承）；`window.__p0`/`__app` 的钩子闭包持有
       * `tabsDataRef`（每个标签的整份原始字节）与快照表，不移除就永远回收不掉。
       */
      releaseAllImageObjectUrls();
      resetAllDirty();
      delete (window as unknown as Record<string, unknown>).__p0;
      delete (window as unknown as Record<string, unknown>).__app;
      boot?.dispose();
    };
    // 依赖刻意留空：引导只做一次；后续状态变化通过 ref 读取，避免重建 Univer 实例
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---------------------------------------------------------------- FPS
  useEffect(() => {
    let raf = 0;
    let frames = 0;
    let last = performance.now();
    const tick = () => {
      frames += 1;
      const now = performance.now();
      if (now - last >= 1000) {
        const measured = Math.round((frames * 1000) / (now - last));
        // 页面不可见/被节流时会测到 0，不要用它覆盖掉上一次的有效值
        if (measured > 0) setFps(measured);
        frames = 0;
        last = now;
      }
      raf = window.requestAnimationFrame(tick);
    };
    raf = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(raf);
  }, []);

  // ---------------------------------------------------------------- 多标签页
  const syncTabs = useCallback(() => {
    setTabs(
      tabsDataRef.current.map((tab) => ({
        id: tab.id,
        name: tab.fileName,
        dirty: dirtyCellCount(tab.id) > 0,
      })),
    );
  }, []);

  /**
   * 释放一个工作簿单元时的**统一记账清理**（关标签 / 冷存 / 拆示例簿都走它）。`disposeUnit` 只放掉
   * Univer 的模型，而按 unitId 记账的还有撤销栈（不随 dispose 清）、插图 blob url、渲染单元缓存、
   * 编辑桶、脏格账本、填充柄日志去重表。分散手写必然漏；尤其示例簿/占位簿复用同一个固定 id，
   * 脏格账本跨实例存活会让新实例继承旧计数（`dirtyCellCount` 既是标签红点，也是外科式导出的依据）。
   * @param keepTabState 冷存必须传 `true`：标签还要切回来重建，编辑桶与脏格账本是重建与导出依据。
   * @returns 顺带归还的插图 blob url 个数（日志用）
   */
  const releaseUnitBookkeeping = useCallback((id: string, options: { keepTabState?: boolean } = {}): number => {
    // ① Univer 的撤销栈：不随 disposeUnit 清理（见 clearUndoRedoFor 的说明）
    try {
      undoRedoServiceRef.current?.clearUndoRedo?.(id);
    } catch {
      /* 已经没了就算了 */
    }
    // ② 插图 blob url：每个都钉着一份图片字节
    const urlsReleased = releaseWorkbookImageObjectUrls(id);
    // ③ 按 unitId 缓存的渲染单元（大表这份对象不小）与填充柄日志去重表
    if (renderCacheRef.current?.unitId === id) renderCacheRef.current = null;
    fillHandleLoggedRef.current.delete(id);
    if (!options.keepTabState) {
      editsByTabRef.current.delete(id);
      clearWorkbookDirty(id);
    }
    return urlsReleased;
  }, []);

  /** 冷存一个标签：先把编辑快照下来（否则随模型一起消失），再 dispose 掉工作簿；字节/文件名/编辑桶留着 */
  const coldStoreTab = useCallback(
    (id: string): boolean => {
      const api = apiRef.current;
      if (!api) return false;
      if (activeTabIdRef.current === id) {
        // 活动标签不该被冷存（调用方已保证），真发生就拒绝，避免"看着的表突然没了"
        log('tab:cold-store-skip', { id, reason: 'active' });
        return false;
      }
      try {
        // 先按 id 把它改过的单元格快照下来，否则编辑随模型一起消失
        const captured = snapshotEditsForRef.current(id);
        api.disposeUnit(id);
        tabRuntimeRef.current = markTabCold(tabRuntimeRef.current, id);
        /** 工作簿被释放，它名下那几本账一起还（明细见 releaseUnitBookkeeping）；切回重建会重走
         * `applyWorkbookFeatures` 重新插图。注意**不能**用 `clearUndoRedoFor`：那会连全局历史账本一起重置 */
        const urlsReleased = releaseUnitBookkeeping(id, { keepTabState: true });
        log('tab:cold-store', {
          id,
          captured,
          urlsReleased,
          resident: tabRuntimeRef.current.filter((tab) => tab.built).length,
        });
        return true;
      } catch (error) {
        log('tab:cold-store-error', { id, message: String(error) });
        return false;
      }
    },
    [releaseUnitBookkeeping],
  );

  /** 常驻窗口：实体化的标签超过上限时，把最久未用的非活动标签冷存掉（内存从 O(N) 变 O(K)） */
  const enforceResidentWindow = useCallback(
    (activeId: string | null) => {
      const evict = planEvictions(tabRuntimeRef.current, activeId, DEFAULT_RESIDENT_TAB_LIMIT);
      evict.forEach((id) => coldStoreTab(id));
    },
    [coldStoreTab],
  );

  /** 首个真实文件进来后释放启动时的"示例工作簿"（它不参与标签管理，否则白占一份工作簿内存） */
  const disposeSampleWorkbook = useCallback(() => {
    try {
      const api = apiRef.current;
      if (!api || tabsDataRef.current.length === 0) return;
      // 示例簿一旦被顶掉就取不到它了（`getWorkbook(id)` 返回 null），所以用 try 包住：存在就释放
      if (!api.getWorkbook(SAMPLE_WORKBOOK_ID)) return;
      api.disposeUnit(SAMPLE_WORKBOOK_ID);
      // 同上：示例簿/占位簿复用同一个固定 id，不清记账会让新实例继承旧的脏格计数
      releaseUnitBookkeeping(SAMPLE_WORKBOOK_ID);
      log('tab:sample-disposed', { id: SAMPLE_WORKBOOK_ID });
    } catch (error) {
      log('tab:sample-dispose-error', { message: String(error) });
    }
  }, [releaseUnitBookkeeping]);

  /** 重建一个冷标签：走与导入完全相同的管线（解析 → 适配 → 建簿 → 应用特性 → 回放编辑）。
   * ① 重建后必须**清掉 Univer 的撤销栈**（不随 dispose 清理，沿用原 unitId 会让 Ctrl+Z 把旧模型的
   * mutation 打到新模型上）；② 历史账本按 unitId 切断同类条目；③ 重建期间置 busy */
  const buildTabUnit = useCallback(async (id: string): Promise<boolean> => {
    const api = apiRef.current;
    const tab = tabsDataRef.current.find((item) => item.id === id) ?? null;
    if (!api || !tab) return false;
    if (tabRuntimeRef.current.find((entry) => entry.id === id)?.built) return true;
    setStatus('importing');
    setStatusText(`正在重建标签「${tab.fileName}」…`);
    try {
      const parsed = await parseXlsx(tab.bytes);
      const outcome = toUniverWorkbook(parsed, { name: tab.fileName, workbookId: tab.id });
      api.createWorkbook(outcome.workbookData);
      instanceServiceRef.current?.setCurrentUnitForType(tab.id);
      focusSheetUnit(tab.id);
      guardRef.current?.suspend();
      // 同 handleFile：重建期间的命令不入账（紧接着的 clearUndoRedoFor 会清账本，但别让幽灵 id 留在认领池里）
      pauseAutoEntriesRef.current = true;
      // 同 handleFile：重建时应用特性同样不算"用户改动"，暂停脏格记账
      pauseDirtyTracking();
      try {
        await applyWorkbookFeatures(api, parsed);
      } finally {
        guardRef.current?.resume();
        resumeDirtyTracking();
        pauseAutoEntriesRef.current = false;
      }
      // 更新导出用的瘦身模型（重建后必须刷新，否则导出的是重建前的旧快照）
      tabsDataRef.current = tabsDataRef.current.map((item) =>
        item.id === id ? { ...item, source: slimForExport(parsed, item.bytes) } : item,
      );
      replayEditsFor(id);
      clearUndoRedoFor(id);
      tabRuntimeRef.current = markTabUsed(tabRuntimeRef.current, id, Date.now());
      log('tab:build-done', { id, fileName: tab.fileName });
      return true;
    } catch (error) {
      log('tab:build-error', { id, message: String(error) });
      toast(`标签「${tab.fileName}」重建失败：${String(error)}`, 'warn');
      return false;
    } finally {
      setStatus('ready');
      setStatusText('就绪');
    }
  }, [toast]);

  /** 切换到某个标签：冷标签先重建，再设为当前单元并重装锁与元信息 */
  const activateTab = useCallback(
    (id: string) => {
      /** 走串行闸：冷标签切换要重建工作簿，不能和"打开文件/会话恢复"并行；先 await 重建、成功后再切 */
      void runUnitOp('activate-tab', async () => {
        if (needsBuild(tabRuntimeRef.current, id)) {
          const ok = await buildTabUnit(id);
          if (!ok) {
            log('tab:activate-skip', { id, reason: 'build-failed' });
            return;
          }
        }
        clearSwapFlashRef.current(); // 黄色提醒框属于上一张表的坐标系，切表即收掉
        // 切走前先把即将离开的标签的编辑落一次（冷存发生在切走之后，那时再读"活动簿"就读错对象了）
        snapshotActiveEditsRef.current();
        try {
          instanceServiceRef.current?.setCurrentUnitForType(id);
        } catch (error) {
          log('tab:switch-error', { message: String(error) });
        }
        focusSheetUnit(id); // 焦点标记：不设它滚轮就不滚（见 focusSheetUnit 的注释）
        const tab = tabsDataRef.current.find((item) => item.id === id) ?? null;
        activeTabIdRef.current = id;
        setActiveTabId(id);
        parsedRef.current = tab?.source ?? null;
        importedFileNameRef.current = tab?.fileName ?? 'workbook.xlsx';
        attachSheetDeps();
        refreshSheetMeta();
        tabRuntimeRef.current = markTabUsed(tabRuntimeRef.current, id, Date.now());
        log('tab:activate', { id, fileName: tab?.fileName ?? null });
        enforceResidentWindow(id);
        await ensureActiveUnitRendered(id);
      });
    },
    [attachSheetDeps, buildTabUnit, enforceResidentWindow, refreshSheetMeta, runUnitOp],
  );

  const closeTab = useCallback(
    (id: string) => {
      const list = tabsDataRef.current;
      const index = list.findIndex((item) => item.id === id);
      if (index < 0) return;
      // 脏标签关闭会**永久丢掉编辑**（标签都没了，会话里也不会再恢复它）→ 先确认
      const dirty = dirtyCellCount(id);
      if (dirty > 0) {
        const name = list[index]?.fileName ?? id;
        const ok = window.confirm(`「${name}」有 ${dirty} 个单元格被改过，关闭后这些修改会丢失。确定关闭？`);
        if (!ok) {
          log('tab:close-cancelled', { id, dirty });
          return;
        }
      }
      snapshotActiveEditsRef.current(); // 关闭前先落一次编辑（导出的脏格账本与之一致）
      /**
       * **关掉最后一个标签必须"先摆占位、再释放"**（顺序反了就是白屏）：唯一的活动单元被释放后，
       * Univer 会把当前单元置成 `null` 并摘掉 canvas，渲染根被回收后新建工作簿也画不回来，只能刷新。
       */
      const isLastTab = tabsDataRef.current.length === 1;
      if (isLastTab) restorePlaceholderRef.current('last-tab-closing');
      try {
        apiRef.current?.disposeUnit(id);
      } catch (error) {
        log('tab:dispose-error', { message: String(error) });
      }
      // 关标签的清理全部走统一入口（缺一项就是泄漏，明细见 releaseUnitBookkeeping）
      releaseUnitBookkeeping(id);
      tabRuntimeRef.current = dropTab(tabRuntimeRef.current, id);
      const next = list.filter((item) => item.id !== id);
      tabsDataRef.current = next;
      log('tab:close', { id, remaining: next.length });

      if (activeTabIdRef.current === id) {
        const fallback = next[Math.min(index, next.length - 1)] ?? null;
        if (fallback) {
          activateTab(fallback.id);
        } else {
          activeTabIdRef.current = null;
          setActiveTabId(null);
          parsedRef.current = null;
          // 舞台那边已经摆回占位单元（见上面 isLastTab），状态栏也回到"刚打开工具"的样子：
          // 不清摘要的话，没有文件还挂着「40 单元格 / 降级 3 项」这种上一位用户的信息。
          setSummary(null);
          setStatusText(IDLE_STATUS_TEXT);
        }
      }
      syncTabs();
    },
    [activateTab, releaseUnitBookkeeping, syncTabs],
  );

  // ---------------------------------------------------------------- 会话持久化
  /** 每个标签已收集的编辑（只存内容：值/公式）——键 `sheetId|row:col` */
  const editsByTabRef = useRef<Map<string, Map<string, SessionTabEdit>>>(new Map());

  /**
   * 把某个工作簿设为"焦点单元"。不能只用 `setCurrentUnitForType`：Univer 的滚轮滚动在
   * `SheetsScrollRenderController._wheelEventListener` 首行就检查 `FOCUSING_SHEET`，
   * 而这个标记**只有 `IUniverInstanceService.focusUnit()` 会置位**，不置位滚轮完全失灵。
   * 编辑器/公式栏（docs 单元）拿焦点会把它置回 false，所以在表格里按下指针时要重新置位。
   */
  const focusSheetUnit = useCallback((id: string | null) => {
    if (!id) return;
    try {
      (instanceServiceRef.current as unknown as { focusUnit?: (unitId: string) => void } | null)?.focusUnit?.(id);
    } catch (error) {
      log('focus:error', { id, message: String(error) });
    }
  }, []);

  /** 舞台主画布还在不在：disposeUnit 掉最后一个单元后容器里连 canvas 都没了，那时"切当前单元"是空动作 */
  const hasSheetCanvas = (): boolean => document.querySelector(`#${CONTAINER_ID} canvas`) !== null;

  /**
   * **摆回"占位单元"**：Univer 必须始终至少有一个单元，否则渲染根会被一起回收。
   * 复现：打开文件 → 刷新（会话还原）→ 关掉还原出来的标签 → 再打开文件 ⇒ 白屏（当前单元变 null、
   * canvas 被摘掉，此后模型全对但画布回不来，新建工作簿也救不回来）。所以关最后一个标签时要先建好
   * 占位单元并切成当前单元、再释放。
   * @param reason `last-tab-closing` 关最后一个标签前调用；`empty-tabs` 标签已空（兜底）；
   *   `dead-render-root` 容器里连 canvas 都没了（尽力而为，这种状态救不回来，只留证据）
   */
  const restorePlaceholderUnit = useCallback(
    (reason: 'last-tab-closing' | 'empty-tabs' | 'dead-render-root'): boolean => {
      const api = apiRef.current;
      if (!api) return false;
      if (reason === 'empty-tabs' && tabsDataRef.current.length > 0) return false;
      if (reason === 'dead-render-root' && hasSheetCanvas()) return false;
      try {
        /**
         * 已有占位簿时只切成当前单元、不再新建：会话恢复期间可能已经摆过一个，
         * 同 id 建两份会让 `disposeSampleWorkbook` 只放掉其中一份，留下"看不见的常驻工作簿"。
         */
        const existing = api.getWorkbook(SAMPLE_WORKBOOK_ID);
        if (!existing) {
          loadWorkbook(api, createSampleWorkbook());
        }
        instanceServiceRef.current?.setCurrentUnitForType(SAMPLE_WORKBOOK_ID);
        focusSheetUnit(SAMPLE_WORKBOOK_ID);
        attachSheetDeps();
        log('tab:placeholder-created', { reason, id: SAMPLE_WORKBOOK_ID, created: !existing });
        return !existing;
      } catch (error) {
        log('tab:placeholder-error', { reason, message: String(error) });
        return false;
      }
    },
    [attachSheetDeps, focusSheetUnit],
  );
  /** `closeTab` 之后的兜底 effect 用它判断"是不是真的开过标签"（见那个 effect 的说明） */
  const hadTabsRef = useRef(false);
  /** `closeTab` 比它先定义（两边都要用 `attachSheetDeps`/`focusSheetUnit`），用 ref 转发 */
  const restorePlaceholderRef = useRef(restorePlaceholderUnit);
  restorePlaceholderRef.current = restorePlaceholderUnit;

  /** **兜底自愈**：确认"活动标签 = 画布上真正渲染的那个工作簿"，不是就重绑一次。"页面空白、刷新后
   * 又出现"的共同点都是活动单元与渲染单元错位（并发建簿、标签被冻结后激活…）。
   * ① 先等一会儿再判（`RENDER_SETTLE_MS`，逐帧轮询）；② 不对就重绑 + `resize` 催布局（canvas 尺寸
   * 靠容器尺寸通知驱动，绑定发生在容器 0 宽时不补通知就一直 0×0）；③ 再等 `RENDER_REBIND_MS` 复验 */
  const ensureActiveUnitRendered = useCallback(
    async (id: string): Promise<boolean> => {
      const measure = (): { ok: boolean; active: string | null; canvasWidth: number } => {
        try {
          const active = apiRef.current?.getActiveWorkbook()?.getId?.() ?? null;
          const canvases = Array.from(document.querySelectorAll<HTMLCanvasElement>('#univer-container canvas'));
          const canvasWidth = canvases.reduce((max, canvas) => Math.max(max, Math.round(canvas.getBoundingClientRect().width)), 0);
          return { ok: active === id && canvasWidth > 20, active, canvasWidth };
        } catch (error) {
          log('render:check-error', { id, message: String(error) });
          return { ok: false, active: null, canvasWidth: 0 };
        }
      };
      const settle = async (timeoutMs: number): Promise<{ ok: boolean; active: string | null; canvasWidth: number }> => {
        const deadline = Date.now() + timeoutMs;
        let state = measure();
        while (!state.ok && Date.now() < deadline) {
          await new Promise<void>((resolve) => {
            window.setTimeout(resolve, 32);
          });
          state = measure();
        }
        return state;
      };

      const first = await settle(RENDER_SETTLE_MS);
      if (first.ok) return true;
      log('render:rebind', { id, active: first.active, canvasWidth: first.canvasWidth });
      try {
        /** **渲染根整个没了**（容器里连 canvas 都没有）：先用占位单元把渲染根拉起来，再切目标单元 */
        if (!hasSheetCanvas()) restorePlaceholderUnit('dead-render-root');
        instanceServiceRef.current?.setCurrentUnitForType(id);
        focusSheetUnit(id);
        attachSheetDeps();
        window.dispatchEvent(new Event('resize')); // 见上面 ②：催布局，专治 canvas 停在 0×0
      } catch (error) {
        log('render:rebind-error', { id, message: String(error) });
      }
      const after = await settle(RENDER_REBIND_MS);
      if (!after.ok) {
        log('render:rebind-failed', { id, active: after.active, canvasWidth: after.canvasWidth });
        /** 诚实兜底：渲染根被回收后 JS 救不回来（新建工作簿也没用），只有刷新能恢复；数据在本机不会丢 */
        if (!hasSheetCanvas()) toast('画面没能恢复，请按 F5 刷新页面（已打开的文件与编辑不会丢）', 'warn');
      }
      return after.ok;
    },
    [attachSheetDeps, focusSheetUnit, restorePlaceholderUnit, toast],
  );

  /** **兜底：标签从"有"变成"没有"后再确认一次占位单元还在**（主路径见 `closeTab` 的"先摆后放"）。
   * 只在真的开过标签后才动手：启动与会话恢复没跑完时 `tabs` 也是空的，那时摆占位簿会多出一份没人管的簿 */
  useEffect(() => {
    if (tabs.length > 0) {
      hadTabsRef.current = true;
      return;
    }
    if (!hadTabsRef.current) return;
    hadTabsRef.current = false;
    if (restoringRef.current) return; // 恢复中，`tabs` 空只是暂时的
    if (restorePlaceholderUnit('empty-tabs')) {
      setSummary(null);
      setStatusText(IDLE_STATUS_TEXT);
    }
  }, [tabs.length, restorePlaceholderUnit]);

  /** 把**指定标签**上被改过的单元格读出来累积进 `editsByTab`。必须按 **id** 取工作簿：冷存发生在
   * 切走之后，`getActiveWorkbook()` 已是新标签，按活动簿读会让被冷存的标签一个编辑都没记下 */
  const snapshotEditsFor = useCallback((workbookId: string) => {
    const workbook = apiRef.current?.getWorkbook(workbookId) ?? null;
    if (!workbook) return 0;
    let bucket = editsByTabRef.current.get(workbookId);
    if (!bucket) {
      bucket = new Map<string, SessionTabEdit>();
      editsByTabRef.current.set(workbookId, bucket);
    }
    const target = bucket;
    let captured = 0;
    for (const [sheetId, keys] of getDirtyCells(workbookId)) {
      const sheet = workbook.getSheetBySheetId(sheetId);
      if (!sheet) continue;
      for (const key of keys) {
        const [rowText, colText] = key.split(':');
        const row = Number(rowText);
        const col = Number(colText);
        if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
        const range = sheet.getRange(row, col);
        const formula = range.getFormula();
        target.set(`${sheetId}|${key}`, {
          sheetId,
          row,
          col,
          value: normalizeExportValue(range.getRawValue()) ?? null,
          formula: formula ? formula.replace(/^=/, '') : null,
        });
        captured += 1;
      }
    }
    return captured;
  }, []);

  /** 活动标签的编辑快照（会话自动保存用） */
  const snapshotActiveEdits = useCallback(() => {
    const workbookId = activeTabIdRef.current;
    if (!workbookId) return;
    snapshotEditsFor(workbookId);
  }, [snapshotEditsFor]);
  snapshotEditsForRef.current = snapshotEditsFor;

  /** 把某个标签已收集的编辑**回放**到（刚重建好的）工作簿：只写内容，样式不动 */
  const replayEditsFor = useCallback((id: string) => {
    const workbook = apiRef.current?.getWorkbook(id) ?? null;
    const bucket = editsByTabRef.current.get(id);
    if (!workbook || !bucket) return 0;
    let applied = 0;
    for (const edit of bucket.values()) {
      const sheet = workbook.getSheetBySheetId(edit.sheetId);
      if (!sheet) continue;
      const range = sheet.getRange(edit.row, edit.col);
      if (edit.formula) range.setFormula(`=${edit.formula}`);
      else range.setValue((edit.value ?? '') as never);
      applied += 1;
    }
    log('tab:edits-replayed', { id, applied });
    return applied;
  }, []);

  /** 标签被冷存/重建后重置撤销与历史。撤销栈**不随 disposeUnit 清理**，沿用同一 unitId 重建而不清栈，
   * Ctrl+Z 会把旧模型的 mutation 打到新模型上；历史账本是全局一条，栈清了它也必须清 */
  const clearUndoRedoFor = useCallback((id: string) => {
    try {
      undoRedoServiceRef.current?.clearUndoRedo?.(id);
    } catch (error) {
      log('tab:clear-undo-error', { id, message: String(error) });
    }
    pendingHistoryLabelRef.current = null;
    prevUndosRef.current = 0;
    historyIndexRef.current = 0;
    setHistoryEntries([]);
    setHistoryIndex(0);
    setCanUndo(false);
    setCanRedo(false);
    log('tab:history-reset', { id });
  }, []);

  snapshotActiveEditsRef.current = snapshotActiveEdits;

  const collectSessionState = useCallback((): SessionState => {
    snapshotActiveEdits();
    const tabs: SessionTab[] = tabsDataRef.current.map((tab) => ({
      id: tab.id,
      fileName: tab.fileName,
      originalBytes: tab.bytes,
      snapshot: null,
      edits: [...(editsByTabRef.current.get(tab.id)?.values() ?? [])],
    }));
    return {
      version: 1,
      savedAt: Date.now(),
      activeTabId: activeTabIdRef.current,
      mode: modeRef.current,
      tabs,
      workspace: itemsStateRef.current,
      settings: settingsRef.current,
    };
  }, [snapshotActiveEdits]);

  /** 便宜的指纹：只在"确实变了"时才落盘，避免把整份文件字节反复写进 IndexedDB */
  const sessionFingerprint = useCallback((): string => {
    const tabPart = tabsDataRef.current
      .map((tab) => `${tab.id}:${dirtyCellCount(tab.id)}`)
      .join(',');
    const s = settingsRef.current;
    return `${tabPart}|${itemsStateRef.current.length}|${modeRef.current}|${activeTabIdRef.current ?? ''}`
      + `|${s.keepSourceOnDrop ? 1 : 0}${s.removeItemAfterPaste ? 1 : 0}:${s.tileMinWidth}:${s.sidebarWidth}`;
  }, []);

  const handleFile = useCallback(
    async (file: File) => {
      const api = apiRef.current;
      if (!api) return;
      /** **格式校验的唯一入口**：文件选择框的 `accept` 与拖放过滤都收口到这里，
       * 保证"打不开的格式"给出同一句「为什么 + 怎么办」，而不是解压失败的原始报错 */
      const bridgeStatus = bridgeStatusRef.current;
      const bridgeOnly = isBridgeOnlyFile(file.name);
      if (!isSupportedWorkbookFile(file.name) && !bridgeOnly) {
        log('import:reject', { name: file.name, kind: fileKindOf(file.name) });
        toast(unsupportedFileMessage([file.name]), 'warn');
        setStatus('ready');
        setStatusText(`打不开 ${file.name}：请先另存为 .xlsx`);
        return;
      }
      if (bridgeOnly && !bridgeStatus.available) {
        // `.xlsb` 只有本地版 + 本机 Excel 才打得开：说清原因，别让用户以为文件坏了
        log('import:reject', { name: file.name, kind: 'xlsb', bridge: bridgeStatus.available, form: APP_FORM });
        const hint = bridgeUnavailableHint(bridgeStatus);
        toast(`打不开 .xlsb：${hint}。也可以先在 Excel 里「另存为 .xlsx」再打开`, 'warn');
        setStatus('ready');
        setStatusText(`打不开 ${file.name}：${hint}`);
        return;
      }
      /** 校验通过后进**串行闸**再动手：否则会与"冷标签重建/会话恢复"同时 createWorkbook +
       * attachSheetDeps，把活动单元与画布绑定搞错位（页面空白、刷新才恢复） */
      await runUnitOp('import', async () => {
      // 埋点：用户"选中文件"这一刻，配合 import:done 得到端到端墙钟时间（分段耗时在 summary 里）
      log('import:start', { name: file.name, size: file.size });
      setStatus('importing');
      setStatusText(`正在解析 ${file.name} …`);
      try {
        const originalBytes = new Uint8Array(await file.arrayBuffer());
        const t0 = performance.now();
        /**
         * **统一打开入口**（`src/importer/open-workbook.ts`）：OOXML 直接原字节（导出才逐字节保真）；
         * csv/ods/xls 由自研解析合成规范 xlsx（导出会另存为 .xlsx）；`.xlsb` 只有本地版能先借 Excel 转换。
         */
        let fileBytes: Uint8Array = originalBytes;
        let bridgeNote: string[] = [];
        if (bridgeOnly) {
          setStatusText(`正在用本机 Excel 转换 ${file.name} …`);
          const converted = await convertViaBridge(originalBytes, 'xlsx');
          if (!converted.ok || !converted.bytes) {
            throw new Error(`经本机 Excel 转换失败：${converted.reason ?? '未知原因'}`);
          }
          fileBytes = converted.bytes;
          bridgeNote = [`由 .xlsb 经本机 Excel 转换为 xlsx 后导入（样式、条件格式、批注、图片都保留）`];
          log('import:bridge-converted', { name: file.name, bytes: fileBytes.length });
        }
        const opened = await openWorkbookBytes(fileBytes, bridgeOnly ? file.name.replace(/\.xlsb$/i, '.xlsx') : file.name);
        opened.notes.push(...bridgeNote);
        const bytes = opened.bytes;
        const parsed = opened.parsed;
        const parseMs = Math.round(performance.now() - t0);
        if (opened.converted) {
          log('import:converted', { name: file.name, from: opened.kind, bytes: bytes.length });
        }

        const t1 = performance.now();
        const outcome: ImportOutcome = toUniverWorkbook(parsed, { name: file.name });
        const adaptMs = Math.round(performance.now() - t1);

        // 多标签：新文件**不**顶掉已打开的工作簿，而是新开一个标签页
        const t2 = performance.now();
        api.createWorkbook(outcome.workbookData);
        const renderMs = Math.round(performance.now() - t2);
        const tabId = outcome.workbookData.id;
        // 瘦身再入表：完整 parsed 只在校验/应用特性时需要，之后就不该继续占内存
        const tabEntrySource = slimForExport(parsed, bytes);
        tabsDataRef.current = [
          ...tabsDataRef.current,
          { id: tabId, fileName: file.name, source: tabEntrySource, bytes },
        ];
        // 先把上一个标签的编辑落一次（此刻 active 还是旧标签），再切到新标签
        snapshotActiveEditsRef.current();
        activeTabIdRef.current = tabId;
        setActiveTabId(tabId);
        tabRuntimeRef.current = markTabUsed(tabRuntimeRef.current, tabId, Date.now());
        disposeSampleWorkbook();
        // 显式设为"当前单元"：多标签下 createWorkbook 不足以保证所有内部服务都切过去
        // （不显式切换时超链接等模型会落到旧单元，读回为空）
        try {
          instanceServiceRef.current?.setCurrentUnitForType(tabId);
        } catch (error) {
          log('tab:activate-new-error', { message: String(error) });
        }
        focusSheetUnit(tabId); // 焦点标记：不设它滚轮就不滚

        // 装"仅内容可编辑"锁之前先应用特性；期间挂起只读闸门（闸门默认拒绝，而这是我们自己的可信路径）
        const t3 = performance.now();
        guardRef.current?.suspend();
        /** 装载期间关掉历史记账：应用特性会连发一串命令，不关会在账本里留下"编辑内容"幽灵条目 */
        pauseAutoEntriesRef.current = true;
        /** 同时暂停脏格记账：应用特性也会经过 mutation 包装，不暂停会让刚打开的文件立刻显示"未保存"，
         * 导出时还会把本可原样保留的 XML 当成"改过的格子"回写（见 `dirty-tracker.ts`） */
        pauseDirtyTracking();
        let featureResult: ApplyFeaturesResult;
        try {
          featureResult = await applyWorkbookFeatures(api, parsed);
        } finally {
          guardRef.current?.resume();
          resumeDirtyTracking();
          pauseAutoEntriesRef.current = false;
        }
        const featureMs = Math.round(performance.now() - t3);
        featureResultRef.current = featureResult;
        /** 清空这个新单元的撤销栈：账本此刻是空的，栈里却躺着特性应用的十几步，会错位成
         * "账本说没得撤、Univer 却还能撤"。只清栈、不动账本（里面可能有之前的工作区动作） */
        try {
          undoRedoServiceRef.current?.clearUndoRedo?.(tabId);
          log('import:undo-stack-cleared', { id: tabId });
        } catch (error) {
          log('import:undo-stack-clear-error', { message: String(error) });
        }

        attachSheetDeps();
        syncTabs();
        // 建完簿、挂完依赖再验一次"画布上渲染的确实是这个新标签"，不对就重绑（见 ensureActiveUnitRendered）
        await ensureActiveUnitRendered(tabId);

        const cells = parsed.sheets.reduce((sum, s) => sum + s.cells.length, 0);
        const parsedFeatures: ParsedFeatureCounts = parsed.sheets.reduce<ParsedFeatureCounts>(
          (acc, sheet) => ({
            cf: acc.cf + (sheet.conditionalFormats?.length ?? 0),
            dv: acc.dv + (sheet.dataValidations?.length ?? 0),
            links: acc.links + (sheet.hyperlinks?.length ?? 0),
            notes: acc.notes + (sheet.notes?.length ?? 0),
            tables: acc.tables + (sheet.tables?.length ?? 0),
            images: acc.images + (sheet.images?.length ?? 0),
          }),
          { cf: 0, dv: 0, links: 0, notes: 0, tables: 0, images: 0 },
        );
        const next: ImportSummary = {
          fileName: file.name,
          sheets: parsed.sheets.map((s) => s.name),
          unsupported: outcome.report.unsupported,
          warnings: outcome.report.warnings,
          // 格式转换说明（CSV/ODS/XLS 才有）排在最前：用户最需要先知道"这份文件被翻译过了"
          preserved: [...opened.notes, ...(outcome.report.preserved ?? [])],
          parseMs,
          adaptMs,
          renderMs,
          featureMs,
          featureIssues: featureResult.issues,
          featureCounts: featureResult.counts,
          parsedFeatures,
          cells,
        };
        summaryRef.current = next;
        setSummary(next);
        parsedRef.current = tabEntrySource;
        importedFileNameRef.current = file.name;
        setStatus('ready');
        setStatusText(
          `${file.name} · ${cells.toLocaleString()} 个单元格 · 解析 ${parseMs}ms / 转换 ${adaptMs}ms / 渲染 ${renderMs}ms / 特性 ${featureMs}ms`,
        );
        log('import:done', next);
        // 新标签成了活动标签：顺手把超出常驻窗口的最久未用标签冷存掉（内存封顶）
        enforceResidentWindow(tabId);
      } catch (error) {
        setStatus('error');
        setStatusText(`导入失败：${String(error)}`);
        log('import:error', { message: String(error), stack: (error as Error)?.stack });
      }
      });
    },
    [attachSheetDeps, enforceResidentWindow, ensureActiveUnitRendered, runUnitOp],
  );
  // 窗口级文件拖放走同一条导入路径（见引导 effect 里的 file:drop 处理）
  fileOpenerRef.current = (file: File) => void handleFile(file);

  /**
   * 组装"当前内容"的 xlsx 字节（外科式：只回写脏格，其余部件字节保留）。
   * 保真导出与"交给本机 Excel 另存"都用它：桥的输入必须是当前看到的内容而非重新生成的一份，
   * 这样条件格式/图片/图表才能带到 .ods/.xls 里去。
   */
  const buildCurrentXlsxBytes = useCallback((): { bytes: Uint8Array; editCount: number; sheets: number } | null => {
    const api = apiRef.current;
    const parsed = parsedRef.current;
    const workbookId = activeTabIdRef.current;
    if (!api || !parsed || !workbookId) {
      toast('请先打开一个表格文件（.xlsx / .xlsm）', 'warn');
      return null;
    }
    const workbook = api.getActiveWorkbook();
    if (!workbook) return null;

    // 脏单元格来自**全局**追踪（与锁实例解耦），这样多标签/切表都不会丢记录
    const dirty = getDirtyCells(workbookId);
    const edits: SheetEdits[] = [];
    let editCount = 0;

    for (const [sheetId, keys] of dirty) {
      const sheet = workbook.getSheetBySheetId(sheetId);
      if (!sheet) continue;
      const cells: CellEdit[] = [];
      for (const key of keys) {
        const [rowText, colText] = key.split(':');
        const row = Number(rowText);
        const col = Number(colText);
        if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
        const range = sheet.getRange(row, col);
        const rawValue = range.getRawValue();
        const formula = range.getFormula();
        cells.push({
          row,
          col,
          value: normalizeExportValue(rawValue),
          formula: formula ? formula.replace(/^=/, '') : null,
        });
      }
      if (cells.length > 0) {
        edits.push({ sheetId, cells });
        editCount += cells.length;
      }
    }

    return { bytes: exportXlsx(parsed, edits), editCount, sheets: edits.length };
  }, [toast]);

  /** 触发一次浏览器下载（纯前端） */
  const downloadBytes = useCallback((bytes: Uint8Array, fileName: string, mime: string) => {
    const blob = new Blob([bytes as unknown as BlobPart], { type: mime });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }, []);

  /** 外科式导出：只回写被编辑过的单元格，其余 zip 部件字节原样保留 */
  const handleExport = useCallback(() => {
    const built = buildCurrentXlsxBytes();
    if (!built) return;
    try {
      downloadBytes(
        built.bytes,
        exportFileNameFor(importedFileNameRef.current),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
      toast(built.editCount > 0 ? `已导出（回写 ${built.editCount} 个单元格，其余部件原样保留）` : '已导出（无内容改动）');
      log('export:done', { edits: built.editCount, sheets: built.sheets, bytes: built.bytes.length, form: APP_FORM });
    } catch (error) {
      toast(`导出失败：${String(error)}`, 'warn');
      log('export:error', { message: String(error) });
    }
  }, [buildCurrentXlsxBytes, downloadBytes, toast]);

  /**
   * 交给**本机 Excel** 另存为其它格式（`.ods/.xls/.xlsb`）。不走第三方库：库只认"值"，
   * 字体/填充/边框/条件格式/批注/图片全丢；Excel 是这些格式的原生实现，结果与人工另存等价。
   */
  const handleBridgeExport = useCallback(
    async (format: BridgeFormat) => {
      if (!bridge.available) {
        toast(bridgeUnavailableHint(bridge), 'warn');
        return;
      }
      const built = buildCurrentXlsxBytes();
      if (!built) return;
      setExportBusy(format);
      log('export:bridge-start', { format, bytes: built.bytes.length, edits: built.editCount });
      try {
        const result = await convertViaBridge(built.bytes, format);
        if (!result.ok || !result.bytes) {
          toast(`转换失败：${result.reason ?? '未知原因'}`, 'warn');
          log('export:bridge-error', { format, reason: result.reason ?? null });
          return;
        }
        const name = bridgedFileName(importedFileNameRef.current, format);
        downloadBytes(result.bytes, name, 'application/octet-stream');
        toast(`已导出 ${BRIDGE_EXTENSION[format]}（由本机 Excel 另存，样式与公式都在）`);
        log('export:bridge-done', { format, name, bytes: result.bytes.length });
      } finally {
        setExportBusy(null);
      }
    },
    [bridge, buildCurrentXlsxBytes, downloadBytes, toast],
  );

  /** 导出当前工作表为 CSV（自研：UTF-8 BOM，学号这类前导零不会被吃掉） */
  const handleExportCsv = useCallback(() => {
    const sheet = getSheet();
    if (!sheet) {
      toast('请先打开一个表格文件', 'warn');
      return;
    }
    try {
      const rows = sheet.getDataRange().getDisplayValues().map((row) => row.map((cell) => String(cell ?? '')));
      const text = buildDelimited(rows, ',');
      const bytes = encodeDelimited(text, true);
      const base = importedFileNameRef.current.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlsb|ods|csv|tsv|txt)$/i, '');
      downloadBytes(bytes, `${base}-${sheet.getSheetName()}.csv`, 'text/csv;charset=utf-8');
      toast(`已导出「${sheet.getSheetName()}」为 CSV（${rows.length} 行，带 UTF-8 BOM）`);
      log('export:csv-done', { sheet: sheet.getSheetName(), rows: rows.length, bytes: bytes.length });
    } catch (error) {
      toast(`导出 CSV 失败：${String(error)}`, 'warn');
      log('export:csv-error', { message: String(error) });
    }
  }, [downloadBytes, getSheet, toast]);

  /** 导出菜单的选项：按**当前真的能不能用**决定放不放进来（不可用的直接隐藏）。`.xlsx` 保真导出要有
   * 打开的文件（示例表没有原始字节可回写）；`.csv` 只要有工作表就一直在；`.ods/.xls/.xlsb` 要转换桥 */
  const exportMenuItems = useMemo((): MenuItemSpec[] => {
    const hasFile = tabs.length > 0 && activeTabId !== null;
    const busy = exportBusy !== null;
    const items: MenuItemSpec[] = [];

    if (hasFile) {
      items.push({ id: 'export-xlsx-keep', label: '导出 .xlsx（保真）', shortcut: '默认', disabled: busy });
    }
    items.push({
      id: 'export-csv',
      label: '导出 .csv（当前工作表）',
      shortcut: 'UTF-8 BOM',
      disabled: busy,
    });
    if (bridge.available && hasFile) {
      items.push({
        id: 'export-ods',
        label: '导出 .ods',
        shortcut: '经本机 Excel',
        disabled: busy,
        separatorBefore: true,
      });
      items.push({ id: 'export-xls', label: '导出 .xls', shortcut: '经本机 Excel', disabled: busy });
      items.push({ id: 'export-xlsb', label: '导出 .xlsb', shortcut: '经本机 Excel · 体积小', disabled: busy });
    }
    return items;
  }, [activeTabId, bridge.available, exportBusy, tabs.length]);

  const handleExportMenuSelect = useCallback(
    (id: string) => {
      setExportMenu(null);
      if (id === 'export-xlsx-keep') {
        handleExport();
        return;
      }
      if (id === 'export-csv') {
        handleExportCsv();
        return;
      }
      if (id === 'export-ods' || id === 'export-xls' || id === 'export-xlsb') {
        void handleBridgeExport(id.replace('export-', '') as BridgeFormat);
      }
    },
    [handleBridgeExport, handleExport, handleExportCsv],
  );

  const startWorkspaceDrag = useCallback((event: React.PointerEvent<HTMLElement>, item: RangeSnapshot) => {
    pendingWorkspaceDragRef.current = { snapshot: item, x: event.clientX, y: event.clientY, started: false };
    // 记下"这一次按下落在哪个条目上"：松手时若没拖动，就当成一次点击（见 onPointerUp）
    pendingWorkspaceClickRef.current = item;
  }, []);

  /** 在工作区条目上"点一下"（按下后没拖动就松手）就把它登记成点击互换的一方，让"先点条目、再点表格
   * 单元格"也能互换；再点同一条目 → 控制器判为同一方并取消；尺寸不同 → 提示"尺寸不一致" */
  const handleWorkspaceItemClick = useCallback((item: RangeSnapshot) => {
    if (modeRef.current !== 'click-swap') return;
    if (!clickSwapRef.current?.isEnabled()) return;
    const side: SwapSide = {
      kind: 'workspace',
      itemId: item.id,
      rows: item.rows,
      cols: item.cols,
      label: `工作区「${item.label}」`,
    };
    clickSwapRef.current.select(side);
    swapAnimRef.current?.pulseSelection(lastPointerRef.current);
    log('swap:select-workspace-item', { id: item.id, label: item.label, rows: item.rows, cols: item.cols });
  }, []);
  handleWorkspaceItemClickRef.current = handleWorkspaceItemClick;

  /** 标记"一次表格动作开始了"（互换/移动/写回/清空/剪切…）：命令会让 Univer 的 undos 先涨一次、订阅
   * 补一条历史，动作结束后 `pushHistory` 用我们的措辞覆盖掉，保证"一次动作 = 一条历史" */
  const beginSheetAction = useCallback((label: string, kind: HistoryKind) => {
    sheetActionRef.current = {
      label,
      kind,
      startAuto: autoSheetEntriesRef.current.length,
    };
    pendingHistoryLabelRef.current = { label, kind, at: Date.now() };
  }, []);

  /** 主动记一条历史（我们自己的动作：互换/移动/写回/工作区增删等） */
  const pushHistory = useCallback(
    (label: string, kind: HistoryKind, scope: 'sheet' | 'workspace' = 'sheet', entryId?: string) => {
      const entry: HistoryEntry = {
        id: entryId ?? `h-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        label,
        kind,
        at: Date.now(),
        scope,
      };
      /** 我们执行的表格命令会**再**被 Univer 的 undos 抓到一次，不合并的话一次动作占两条历史、撤销要多
       * 走一步。这里把这次动作期间补出来的那几条收敛成一条并用我们的措辞命名（认领按 id 精确命中） */
      if (scope === 'sheet') {
        pendingHistoryLabelRef.current = null;
        const action = sheetActionRef.current;
        sheetActionRef.current = null;
        const claimed = action ? autoSheetEntriesRef.current.splice(action.startAuto) : [];
        if (claimed.length > 0) {
          const ids = new Set(claimed);
          const live = entriesRef.current.slice(0, historyIndexRef.current);
          const kept: HistoryEntry[] = [];
          let insertAt = -1;
          for (const item of live) {
            if (ids.has(item.id)) {
              if (insertAt < 0) insertAt = kept.length;
              continue;
            }
            kept.push(item);
          }
          if (insertAt < 0) insertAt = kept.length;
          // 沿用第一条自动条目的 id（表格条目没有挂快照，换 id 无副作用）
          const next = appendEntry(
            [...kept.slice(0, insertAt), { ...entry, id: claimed[0] }, ...kept.slice(insertAt)],
            { ...entry, id: claimed[0] },
            200,
          );
          entriesRef.current = next;
          historyIndexRef.current = next.length;
          setHistoryEntries(next);
          setHistoryIndex(next.length);
          lastHistoryPushAtRef.current = Date.now();
          log('history:push-merge', { label, kind, merged: claimed.length });
          return;
        }
      }
      // 同步账本先落定（撤销可能紧接着就来），再让 state 跟上
      const next = appendEntry(entriesRef.current.slice(0, historyIndexRef.current), entry, 200);
      entriesRef.current = next;
      historyIndexRef.current = next.length;
      setHistoryEntries(next);
      setHistoryIndex(next.length);
      // 历史被截断时，把已经不在账本里的工作区快照清掉（避免无限增长）
      const alive = new Set(next.map((item) => item.id));
      for (const id of [...workspaceSnapshotsRef.current.keys()]) {
        if (!alive.has(id)) workspaceSnapshotsRef.current.delete(id);
      }
      lastHistoryPushAtRef.current = Date.now();
      log('history:push', { label, kind, scope });
    },
    [],
  );
  pushHistoryRef.current = pushHistory;

  /** 工具栏"撤销/重做"按钮的可用态：必须以**账本**为准——工作区动作不在 Univer 的栈里，只看 undos 按钮会误灰 */
  useEffect(() => {
    setCanUndo(historyIndex > 0);
    setCanRedo(historyIndex < historyEntries.length);
  }, [historyEntries.length, historyIndex]);

  /** 撤销一步（按时间顺序在表格与工作区之间切换）。账本是唯一顺序来源：
   * 上一步是工作区动作就用我们的快照回放（`before`），否则交给 Univer 的撤销栈 */
  const stepBack = useCallback(async (): Promise<boolean> => {
    const index = historyIndexRef.current;
    if (index <= 0) return false;
    const entry = entriesRef.current[index - 1] ?? null;
    if (entry?.scope === 'workspace') {
      const snap = workspaceSnapshotsRef.current.get(entry.id);
      if (!snap) return false;
      itemsStateRef.current = snap.before;
      setItems(snap.before);
      syncSnapshotStore(snap.before); // 撤销"移除条目"时要把快照补回来（否则后续取不到 → "条目不存在"）
      historyIndexRef.current = index - 1;
      setHistoryIndex(index - 1);
      log('history:undo-workspace', { label: entry.label, items: snap.before.length });
      return true;
    }
    suppressAutoEntriesRef.current = 1; // 这次计数变化是我们自己造成的，不要重复记账
    const ok = (await apiRef.current?.undo()) ?? false;
    suppressAutoEntriesRef.current = 0;
    if (ok) {
      historyIndexRef.current = Math.max(0, index - 1);
      setHistoryIndex(historyIndexRef.current);
      log('history:undo-sheet', { label: entry?.label ?? null });
    }
    return Boolean(ok);
  }, [syncSnapshotStore]);

  /** 重做一步（与 stepBack 对称） */
  const stepForward = useCallback(async (): Promise<boolean> => {
    const index = historyIndexRef.current;
    const list = entriesRef.current;
    if (index >= list.length) return false;
    const entry = list[index] ?? null;
    if (entry?.scope === 'workspace') {
      const snap = workspaceSnapshotsRef.current.get(entry.id);
      if (!snap) return false;
      itemsStateRef.current = snap.after;
      setItems(snap.after);
      syncSnapshotStore(snap.after); // 与撤销对称：重做后 store 必须与 items 一致
      historyIndexRef.current = index + 1;
      setHistoryIndex(index + 1);
      log('history:redo-workspace', { label: entry.label, items: snap.after.length });
      return true;
    }
    suppressAutoEntriesRef.current = 1; // 重做同样会让 undos +1（见订阅里的说明）
    const ok = (await apiRef.current?.redo()) ?? false;
    suppressAutoEntriesRef.current = 0;
    if (ok) {
      historyIndexRef.current = Math.min(list.length, index + 1);
      setHistoryIndex(historyIndexRef.current);
      log('history:redo-sheet', { label: entry?.label ?? null });
    }
    return Boolean(ok);
  }, [syncSnapshotStore]);

      /**
       * 全局撤销/重做快捷键的**处理函数**（转发器由 `main.tsx` 在引导前注册，见 `src/shell/undo-shortcut.ts`）。
       * 不能依赖 Univer 的快捷键服务（只在表格获焦时响应），且它的 `ShortcutService` 在 window capture
       * 阶段注册 Ctrl+Z 且早于我们，`stopPropagation()` 拦不住 —— 必须用 `stopImmediatePropagation()`；
       * 焦点真在编辑器/公式栏/输入框里时让路。
       */
  useEffect(() => {
    /**
     * 焦点是不是在**文本编辑**元素里（在那里让路给原生撤销）。不能简单按 `<input>` 判断：
     * 导入用的 `<input type="file">`、勾选框、按钮都是 input 但不是文本编辑，
     * 粗规则会让选完文件后的 Ctrl+Z 完全没反应。
     */
    const TEXT_INPUT_TYPES = new Set([
      'text',
      'search',
      'url',
      'tel',
      'email',
      'password',
      'number',
      'date',
      'datetime-local',
      'month',
      'time',
      'week',
    ]);
    const isTextEditing = (target: EventTarget | null): boolean => {
      const element = target as HTMLElement | null;
      if (!element || typeof element.closest !== 'function') return false;
      // 我们自己的文本输入框：让原生撤销处理
      const field = element.closest('input, textarea') as HTMLInputElement | HTMLTextAreaElement | null;
      if (field) {
        if (field.tagName === 'TEXTAREA') return true;
        const type = (field as HTMLInputElement).type?.toLowerCase() ?? 'text';
        return TEXT_INPUT_TYPES.has(type);
      }
      if (!element.closest('[contenteditable="true"], [contenteditable=""]')) return false;
      /**
       * 可编辑元素要再分两种：**单元格编辑器/公式栏正在编辑** → 让 Univer 撤这一格里的输入；
       * **Univer 那个常驻的隐藏焦点代理**（contenteditable，用来收键盘事件）→ 必须由我们接管。
       * 用 `EDITOR_ACTIVATED` / `FOCUSING_FX_BAR_EDITOR` 上下文标记区分。
       */
      try {
        const context = rootInjectorRef.current?.get(IContextService) as
          | { getContextValue: (key: string) => unknown }
          | null;
        return Boolean(context?.getContextValue('EDITOR_ACTIVATED')) || Boolean(context?.getContextValue('FOCUSING_FX_BAR_EDITOR'));
      } catch {
        return true; // 读不到上下文时保守一点：不抢编辑器的撤销
      }
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
      const key = event.key.toLowerCase();
      const isUndo = key === 'z' && !event.shiftKey;
      const isRedo = key === 'y' || (key === 'z' && event.shiftKey);
      if (!isUndo && !isRedo) return;
      const editing = isTextEditing(event.target);
      log('history:shortcut', {
        key: event.key,
        undo: isUndo,
        redo: isRedo,
        editing,
        target: (event.target as HTMLElement | null)?.tagName ?? null,
      });
      if (editing) return; // 输入框/单元格编辑器里让给编辑器自己（撤的是那格里的输入）
      /** 账本上确实还有可撤/可重做的步骤时才截下按键；否则放行让 Univer 的快捷键服务去试 */
      const hasStep = isUndo
        ? historyIndexRef.current > 0
        : historyIndexRef.current < entriesRef.current.length;
      if (!hasStep) return;
      event.preventDefault();
      /** 必须用 `stopImmediatePropagation`：Univer 的 ShortcutService 注册在同一阶段且早于我们 */
      event.stopImmediatePropagation();
      void (isUndo ? stepBack() : stepForward());
    };
    setUndoShortcutHandler(onKeyDown);
    return () => setUndoShortcutHandler(null);
  }, [stepBack, stepForward]);
  stepBackRef.current = stepBack;
  stepForwardRef.current = stepForward;

  /** 跳步：按账本顺序一步步撤/重做（表格动作走 Univer，工作区动作走我们的快照） */
  const handleJumpTo = useCallback(
    async (target: number) => {
      const plan = planJump(historyIndexRef.current, target);
      // 逐次串行：两者都可能改变状态，且 Univer 的撤销栈只能连续走
      for (let i = 0; i < plan.undo; i += 1) await stepBack();
      for (let i = 0; i < plan.redo; i += 1) await stepForward();
      log('history:jump', { target, plan });
    },
    [stepBack, stepForward],
  );

  const removeItem = useCallback(
    (id: string, label?: string) => {
      const target = itemsStateRef.current.find((item) => item.id === id) ?? null;
      commitWorkspace((prev) => prev.filter((item) => item.id !== id), label ?? `从工作区移除 ${target?.label ?? id}`, 'workspace');
      log('workspace:remove-item', { id });
    },
    [commitWorkspace],
  );

  /** 批量移除（"只清空筛选出来的那部分"用）。必须一次原子更新：逐条调用会打出 N 次渲染与 N 条日志 */
  const removeItems = useCallback(
    (ids: string[]) => {
      if (ids.length === 0) return;
      const doomed = new Set(ids);
      commitWorkspace((prev) => prev.filter((item) => !doomed.has(item.id)), `从工作区移除 ${ids.length} 格`, 'workspace');
      log('workspace:remove-items', { count: ids.length });
    },
    [commitWorkspace],
  );

  const clearItems = useCallback(() => {
    if (itemsStateRef.current.length === 0) {
      setWsClearConfirming(false);
      return;
    }
    commitWorkspace([], `清空工作区（${itemsStateRef.current.length} 格）`, 'workspace');
    setWsClearConfirming(false); // 清空之后不该还留着确认条
    log('workspace:clear-all', { at: Date.now() });
  }, [commitWorkspace]);

  /**
   * 单元格 ⇄ 工作区条目互换后，把"从单元格换出来的内容"落到条目上。
   * 互换是反向通道：换回来的是那一格原来的内容，若为空，条目会变成白卡片、破坏
   * "工作区不放空内容单元格"这条规则。所以有内容 → 更新条目；空内容 → **移除条目**
   * （内容没丢，已经在单元格里，一次 Ctrl+Z 就能还回来）。
   */
  const absorbSwappedCell = useCallback(
    (item: RangeSnapshot, cellSnapshot: RangeSnapshot | null): 'updated' | 'removed' | 'unchanged' => {
      if (!cellSnapshot) return 'unchanged'; // 取不到快照就不动条目（宁可留着，也不误删）
      if (isSnapshotEmpty(cellSnapshot)) {
        removeItem(item.id, `移除空条目「${item.label}」（换出的单元格是空的）`);
        log('workspace:remove-empty-after-swap', { id: item.id, a1: cellSnapshot.source.a1 });
        return 'removed';
      }
      const updated: RangeSnapshot = {
        ...cellSnapshot,
        id: item.id,
        label: item.label,
        source: item.source,
        createdAt: item.createdAt,
      };
      commitWorkspace(
        (prev) => prev.map((entry) => (entry.id === item.id ? updated : entry)),
        `更新工作区条目「${item.label}」`,
        'workspace',
      );
      return 'updated';
    },
    [commitWorkspace, log, removeItem],
  );

  /* ---- 剪贴板：工作区条目的「复制/剪切」放进内部剪贴板，「粘贴」到表格时只写内容、保留目标格格式 ---- */

  const copyItem = useCallback(
    (item: RangeSnapshot, cut: boolean) => {
      clipboardRef.current = item;
      setClipboardLabel(item.label);
      if (cut) {
        removeItem(item.id);
        toast(`已剪切「${item.label}」到剪贴板（工作区条目已移除）`);
        log('workspace:clipboard', { mode: 'cut', label: item.label });
      } else {
        toast(`已复制「${item.label}」到剪贴板，可在表格中粘贴`);
        log('workspace:clipboard', { mode: 'copy', label: item.label });
      }
    },
    [removeItem, toast],
  );

  /** 把剪贴板内容写进表格的某个 A1（只写内容，样式留在目标格） */
  const pasteClipboardAt = useCallback(
    (a1: string): boolean => {
      const sheet = getSheet();
      const clip = clipboardRef.current;
      if (!sheet || !clip) return false;
      const rect = sheet.getRange(a1).getRange();
      // 同"工作区条目拖回表格"：写进表格就是**表格侧**的一步，必须走表格通道（否则撤销失灵）
      const label = `粘贴工作区条目 ${clip.label}`;
      beginSheetAction(label, 'edit');
      applySnapshot(sheet, clip, { row: rect.startRow, col: rect.startColumn });
      pushHistory(label, 'edit');
      toast(`已粘贴「${clip.label}」→ ${a1}（样式保持不变）`);
      log('workspace:clipboard-paste', { label: clip.label, to: a1 });
      clearSelectionRef.current();
      return true;
    },
    [beginSheetAction, getSheet, pushHistory, toast],
  );

  /** Ctrl/Cmd+V：剪贴板里有工作区内容时优先粘贴它（否则交给浏览器/表格原生行为） */
  const handleClipboardKey = useCallback(
    (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'v') return;
      if (!clipboardRef.current) return;
      const container = containerRef.current;
      if (!container) return;
      if (!container.contains(document.activeElement) && !container.contains(event.target as Node)) return;
      const a1 = selectionText;
      if (!a1) return;
      event.preventDefault();
      event.stopPropagation();
      pasteClipboardAt(a1.split(':')[0]);
    },
    [pasteClipboardAt, selectionText],
  );

  /**
   * 拖分隔条改工作区宽度（同时也是表格区域宽度）：往左拖变宽。
   * ① 必须同时监听 `pointercancel`：手势被取消时不会再发 pointerup，只监听 pointerup 会让
   * 监听赖在 window 上，每次 pointermove 都 setState 重渲染（表现为"拖完还一直卡"）；
   * ② 卸载时要把进行中的监听摘掉（它们不在 effect 里，闭包抓着 `updateSettings`）。
   */
  const startSidebarResize = useCallback(
    (event: React.PointerEvent<HTMLDivElement>) => {
      event.preventDefault();
      const startX = event.clientX;
      const startWidth = settingsRef.current.sidebarWidth;
      const move = (e: PointerEvent): void => {
        updateSettings({ sidebarWidth: startWidth - (e.clientX - startX) });
      };
      const end = (): void => {
        window.removeEventListener('pointermove', move, true);
        window.removeEventListener('pointerup', end, true);
        window.removeEventListener('pointercancel', end, true);
        sidebarResizeCleanupRef.current = null;
      };
      const up = (): void => {
        end();
        log('ui:sidebar-resize', { width: settingsRef.current.sidebarWidth });
      };
      sidebarResizeCleanupRef.current = end;
      window.addEventListener('pointermove', move, true);
      window.addEventListener('pointerup', up, true);
      window.addEventListener('pointercancel', up, true);
    },
    [updateSettings],
  );

  // 卸载时兜底摘掉"拖动中"的监听（与上面成对；见 startSidebarResize 的说明）
  useEffect(() => () => sidebarResizeCleanupRef.current?.(), []);

  /** Esc 取消"清空"的二次确认（确认条是模态意图，键盘也要能退出来） */
  useEffect(() => {
    if (!wsClearConfirming) return;
    const onKey = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape') return;
      event.stopPropagation();
      setWsClearConfirming(false);
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [wsClearConfirming]);

  /** 挂在 window 捕获阶段，先于表格原生粘贴 */
  useEffect(() => {
    window.addEventListener('keydown', handleClipboardKey, true);
    return () => window.removeEventListener('keydown', handleClipboardKey, true);
  }, [handleClipboardKey]);

  /** 一键把"整表里有内容的部分"放进工作区时用的已用区域（空表为 null） */
  const usedRangeA1 = (() => {
    const sheet = getSheet();
    if (!sheet) return null;
    try {
      const rect = sheet.getDataRange().getRange();
      const rows = rect.endRow - rect.startRow + 1;
      const cols = rect.endColumn - rect.startColumn + 1;
      if (rows <= 0 || cols <= 0) return null;
      return rectToA1(rect);
    } catch {
      return null;
    }
  })();

  const swapHandlerRef = useRef<(a: SwapSide, b: SwapSide) => void>(() => {});

  /** 点击交换：两侧都可能是"单元格"或"工作区条目" */
  swapHandlerRef.current = (a: SwapSide, b: SwapSide) => {
    const sheet = getSheet();
    if (!sheet) return;
    const point = lastPointerRef.current;
    const cellSide = a.kind === 'cell' ? a : b.kind === 'cell' ? b : null;
    const itemSide = a.kind === 'workspace' ? a : b.kind === 'workspace' ? b : null;

    // 交换动画：两个"内容芯片"对飞（用点击坐标，不需要单元格像素映射）
    void swapAnimRef.current?.playSwap(
      { x: point.x - 30, y: point.y - 16, text: cellSide?.label ?? a.label },
      { x: point.x + 30, y: point.y + 16, text: itemSide?.label ?? b.label },
    );

    // 工作区条目 ↔ 单元格：**双向**互换内容（样式各自保留）
    if (cellSide && itemSide) {
      const item = snapshotStoreRef.current.get(itemSide.itemId ?? '');
      if (!item || !cellSide.a1) {
        toast('工作区条目不存在', 'warn');
        return;
      }
      const rect = sheet.getRange(cellSide.a1).getRange();
      const { snapshot: cellSnapshot } = extractSnapshot(sheet, cellSide.a1);
      beginSheetAction(`与工作区「${item.label}」互换`, 'swap');
      applySnapshot(sheet, item, { row: rect.startRow, col: rect.startColumn });
      // 条目内容变了 → 进工作区撤销账本；换出来是空格子 → 移除条目（见 absorbSwappedCell）
      const absorbed = absorbSwappedCell(item, cellSnapshot);
      toast(
        absorbed === 'removed'
          ? `已把「${item.label}」写入 ${cellSide.a1}；该格原本是空的，白条目已从工作区移除`
          : `已与工作区条目「${item.label}」互换内容（样式保持不变）`,
      );
      pushHistory(`与工作区「${item.label}」互换`, 'swap');
      log('swap:workspace-cell', { item: item.label, cell: cellSide.a1, absorbed });
      flashSwapRef.current(cellSide.a1); // 取消选中 + 黄色提醒框
      return;
    }

    // 单元格 ↔ 单元格
    if (a.kind === 'cell' && b.kind === 'cell' && a.a1 && b.a1) {
      beginSheetAction(`互换 ${a.a1} ⇄ ${b.a1}`, 'swap');
      const result = swapRanges(sheet, a.a1, b.a1);
      if (result.ok) {
        toast(`已互换 ${a.a1} ⇄ ${b.a1}（样式保持不变）`);
        pushHistory(`互换 ${a.a1} ⇄ ${b.a1}`, 'swap');
        log('swap:cell-cell', { a: a.a1, b: b.a1 });
        flashSwapRef.current(a.a1, b.a1); // 取消选中 + 黄色提醒框
      } else {
        toast(result.reason ?? '互换失败', 'warn');
      }
    }
  };

  const handleModeChange = useCallback(
    (next: InteractionMode) => {
      modeRef.current = next;
      setMode(next);
      clickSwapRef.current?.setMode(next);
      // 切模式时把上一种模式的"半成品状态"清干净：拖动中的幽灵、点击互换的待选
      controllerRef.current?.cancel();
      dragSourceA1Ref.current = null;
      targetHighlightRef.current?.hide();
      clickSwapRef.current?.clear();
      swapAnimRef.current?.hidePendingGhost();
      toast(INTERACTION_MODE_HINTS[next]);
      log('app:mode', { mode: next });
    },
    [toast],
  );

  const handleNewWorkspaceItem = useCallback(() => {
    // 「+」打开快速导入对话框：可手动填写 行号 / 列标 / 区域，也可一键取当前选区
    setImportOpen(true);
  }, []);

  /** 把选区补齐到**整块合并单元格**（返回补齐后的矩形与"是否扩过"）：课程表里满是合并块，拖选
   * B4:F8 时 Univer 已把模型选区扩成 B4:F11，剪切会剪走整块；补齐后同步选区 + 提示保证所见即所得 */
  const normalizeRangeToMerges = useCallback(
    (sheet: FWorksheet, rect: { startRow: number; startColumn: number; endRow: number; endColumn: number }) => {
      try {
        const merges = (sheet as unknown as { getMergeData?: () => Array<{ getRange: () => typeof rect }> })
          .getMergeData?.()
          ?.map((range) => range.getRange()) ?? [];
        return expandRectToMerges(rect, merges);
      } catch (error) {
        log('menu:merge-lookup-error', { message: String(error) });
        return { rect, expanded: false };
      }
    },
    [],
  );

  /** 右键菜单：在 contextmenu 那一刻快照选区，动作一律按快照执行 */
  const handleContextMenu = useCallback(
    (event: MouseEvent) => {
      const container = containerRef.current;
      if (!container || !container.contains(event.target as Node)) return;
      const sheet = getSheet();
      if (!sheet) return;
      event.preventDefault();
      // 必须同时阻断传播：本监听挂在容器捕获阶段，不 stop 的话事件继续下传到 canvas，
      // Univer 自带的右键菜单会弹出来盖住我们的菜单（用户点不动我们的项）
      event.stopPropagation();
      event.stopImmediatePropagation();

      let a1 = 'A1';
      let rows = 1;
      let cols = 1;
      let text = '';
      let note: string | null = null;
      /**
       * 右键作用在**全部**选区上（Ctrl+点选可能有多块）。`a1List` 是规范化后的区域表；
       * `a1` 保留"主区域"，供"与工作区互换/粘贴"这类只支持单块的操作（多块时禁用并说明原因）。
       */
      let a1List: string[] = [];
      try {
        const liveList = readSelectionA1List();
        /** `normalizeRangeToMerges` 用的是 `startColumn/endColumn` 命名（Univer 的 IRange 口径） */
        const rects: Array<{ startRow: number; startColumn: number; endRow: number; endColumn: number }> = [];
        let expandedAny = false;
        for (const candidate of liveList) {
          const live = sheet.getRange(candidate).getRange();
          /** 与合并单元格对齐：每一块都要补，保证"看到的 = 被操作的"（见 normalizeRangeToMerges） */
          const { rect, expanded } = normalizeRangeToMerges(sheet, live);
          if (expanded) expandedAny = true;
          rects.push(rect);
        }
        const normalized = normalizeRanges(
          rects.map((rect) =>
            rectToA1({
              startRow: rect.startRow,
              startColumn: rect.startColumn,
              endRow: rect.endRow,
              endColumn: rect.endColumn,
            }),
          ),
        );
        a1List = rectsToA1(normalized);
        if (normalized.length > 0) {
          const first = normalized[0];
          rows = first.endRow - first.startRow + 1;
          cols = first.endCol - first.startCol + 1;
          a1 = rectToA1({ startRow: first.startRow, startColumn: first.startCol, endRow: first.endRow, endColumn: first.endCol });
          text = sheet.getRange(a1).getDisplayValues().flat().filter(Boolean).slice(0, 3).join(' / ');
        }
        if (expandedAny && a1List.length === 1) {
          note = `选区包含合并单元格，已按整块 ${a1} 处理`;
          log('menu:range-expanded-for-merge', { from: liveList.join(' '), to: a1 });
          toast(note);
          try {
            // 高亮同步：让用户看到的就是会被操作的范围（多块时不动选区，避免把多选压成一块）
            sheet.setActiveSelection(sheet.getRange(a1));
          } catch (error) {
            log('menu:sync-selection-error', { message: String(error) });
          }
        } else if (expandedAny) {
          note = `选区含合并单元格：已按整块处理（当前 ${a1List.length} 块）`;
          toast(note);
          log('menu:range-expanded-for-merge', { from: liveList.join(' '), to: a1List.join(' ') });
        }
      } catch (error) {
        log('menu:selection-error', { message: String(error) });
      }
      const nextMenu = { x: event.clientX, y: event.clientY, snapshot: { a1, a1List, rows, cols, text }, note };
      menuRef.current = nextMenu;
      setMenu(nextMenu);
      log('menu:open', { a1, blocks: a1List.length, rows, cols });
    },
    [getSheet, normalizeRangeToMerges, readSelectionA1List, toast],
  );

  /** 菜单项：只包含"允许的动作"（样式/结构类操作不在此列，因为产品不允许改格式） */
  const menuItems = useCallback((): MenuItemSpec[] => {
    const snapshot = menu?.snapshot;
    if (!snapshot) return [];
    const firstItem = itemsStateRef.current[0] ?? null;
    const blocks = snapshot.a1List.length;
    const multi = blocks > 1;
    const scopeLabel = multi ? `${blocks} 块：${summarizeSelection(snapshot.a1List, 2)}` : snapshot.a1;
    const sizeMatches = !!firstItem && firstItem.rows === snapshot.rows && firstItem.cols === snapshot.cols;
    // 天然只有一方的操作（互换/粘贴）在多块选区下必须说清原因，而不是随便挑一块执行
    const singleBlockReason = describeSingleBlockOnly(snapshot.a1List, '与工作区互换');
    return [
      { id: 'workspace-copy', label: `复制到工作区（${scopeLabel}）` },
      { id: 'workspace-cut', label: `剪切到工作区（${scopeLabel}）` },
      {
        id: 'swap-workspace',
        label: firstItem
          ? `与工作区「${firstItem.label}」互换内容${multi ? '（只支持单块选区）' : ''}`
          : '与工作区条目互换（工作区为空）',
        disabled: !sizeMatches || multi,
        title: singleBlockReason ?? undefined,
        separatorBefore: true,
      },
      {
        // 单块走 Univer 原生复制（同时写 TSV 与带格式 HTML，字体/底色/边框/合并/列宽都会跟过去）；
        // 多块只认纯文本——原生复制只取"最后一个选区"，会丢块，所以标签里如实分开写
        id: 'copy',
        label: multi
          ? `复制内容（${blocks} 块 / ${totalCells(normalizeRanges(snapshot.a1List))} 格，纯文本）`
          : `复制内容（${snapshot.rows}×${snapshot.cols}，含格式）`,
        shortcut: multi ? undefined : 'Ctrl+C',
      },
      {
        // 工作区条目「复制/剪切」后，在表格里右键即可粘贴（只写内容、保留目标格式）
        id: 'paste',
        label: clipboardLabel ? `粘贴「${clipboardLabel}」（保留格式）` : '粘贴（剪贴板为空）',
        disabled: !clipboardLabel || multi,
        title: multi ? describeSingleBlockOnly(snapshot.a1List, '粘贴') ?? undefined : undefined,
      },
      {
        id: 'clear',
        label: multi ? `清空这 ${blocks} 块的内容（保留格式）` : '清空内容（保留格式）',
        danger: true,
        separatorBefore: true,
      },
      { id: 'undo', label: '撤销', shortcut: 'Ctrl+Z', disabled: !canUndo, separatorBefore: true },
      { id: 'redo', label: '重做', shortcut: 'Ctrl+Y', disabled: !canRedo },
    ];
  }, [canRedo, canUndo, clipboardLabel, menu]);

  const handleMenuSelect = useCallback(
    (id: string) => {
      log('menu:select', { id });
      // ContextMenu 把回调存在 ref 里（首帧绑一次），闭包捕获的 menu 是旧值 → 必须从 ref 读当前菜单
      const snapshot = menuRef.current?.snapshot ?? null;
      const sheet = getSheet();
      menuRef.current = null;
      setMenu(null);
      if (!snapshot || !sheet) return;

      if (id === 'workspace-copy' || id === 'workspace-cut') {
        const list = snapshot.a1List.length > 0 ? snapshot.a1List : [snapshot.a1];
        const scope = list.length > 1 ? `${list.length} 块区域` : snapshot.a1;
        // 统一入口：一次提交（一条历史）、跳过空内容、日志口径与拖拽/导入完全一致
        const outcome = addWorkspaceFromRanges(list, { x: window.innerWidth - 150, y: 160 });
        if (!outcome.ok) return;
        const collected = outcome.items;
        log('menu:workspace', { mode: id, a1: list.join(' '), blocks: list.length, items: collected.length });
        if (id === 'workspace-copy') {
          log('workspace:add', { a1: list.join(' '), mode: 'copy', blocks: list.length });
          return;
        }
        /** 剪切：逐块清空内容（保留格式）。每块各记一条历史——它们是各自独立的 mutation，撤销逐块回来 */
        let cleared = 0;
        for (const ref of list) {
          beginSheetAction(`剪切 ${ref} 到工作区`, 'swap');
          const result = clearRange(sheet, ref);
          if (!result.ok) {
            toast(result.reason ?? `剪切 ${ref} 失败（内容已复制到工作区）`, 'warn');
            continue;
          }
          pushHistory(`剪切 ${ref} 到工作区`, 'swap');
          cleared += 1;
        }
        toast(`已剪切 ${scope} 到工作区（${cleared} 块源内容已清空，格式保留）`);
        log('workspace:add', { a1: list.join(' '), mode: 'cut', blocks: list.length, cleared });
        clearSelectionRef.current();
        return;
      }
      if (id === 'swap-workspace') {
        if (describeSingleBlockOnly(snapshot.a1List, '与工作区互换')) {
          toast(describeSingleBlockOnly(snapshot.a1List, '与工作区互换')!, 'warn');
          return;
        }
        const item = itemsStateRef.current[0];
        if (!item) return;
        const rect = sheet.getRange(snapshot.a1).getRange();
        const { snapshot: cellSnapshot } = extractSnapshot(sheet, snapshot.a1);
        beginSheetAction(`与工作区「${item.label}」互换`, 'swap');
        applySnapshot(sheet, item, { row: rect.startRow, col: rect.startColumn });
        // 同"点击互换"：换出来的是空格子就移除条目，工作区里不留白卡片
        const absorbed = absorbSwappedCell(item, cellSnapshot);
        pushHistory(`与工作区「${item.label}」互换`, 'swap');
        toast(
          absorbed === 'removed'
            ? `已把「${item.label}」写入 ${snapshot.a1}；该格原本是空的，白条目已从工作区移除`
            : '已与工作区条目互换内容（样式保持不变）',
        );
        flashSwapRef.current(snapshot.a1); // 取消选中 + 黄色提醒框
        return;
      }
      if (id === 'paste') {
        if (describeSingleBlockOnly(snapshot.a1List, '粘贴')) {
          toast(describeSingleBlockOnly(snapshot.a1List, '粘贴')!, 'warn');
          return;
        }
        if (!pasteClipboardAt(snapshot.a1)) toast('剪贴板为空：先在工作区条目上「复制」或「剪切」', 'warn');
        return;
      }
      if (id === 'copy') {
        const list = snapshot.a1List.length > 0 ? snapshot.a1List : [snapshot.a1];
        /**
         * 单块交给 **Univer 原生复制**（`univer.command.copy`）：它同时写 `text/plain`（TSV，显示值）
         * 与 `text/html`（`<table>` + 内联样式，能带出字体/底色/边框/合并/列宽，Excel 也认）。
         * 在没有 Clipboard API 的环境（http://局域网IP）它会自动降级到 `execCommand('copy')`，
         * 比 `navigator.clipboard.writeText()` 稳（后者在那类环境会静默失败）。失败时退回纯文本路径。
         */
        if (list.length === 1) {
          const commandService = sheet.getInject().get(ICommandService);
          void commandService.executeCommand('univer.command.copy').then(
            (ok) => {
              log('menu:copy', { a1: list[0], blocks: 1, flavor: 'text/plain+text/html', ok: Boolean(ok) });
              if (ok) {
                toast('已复制到剪贴板（含格式，可直接粘到 Excel / WPS）');
                return;
              }
              void copyPlainTextToClipboard(sheet, list, toast);
            },
            (error: unknown) => {
              log('menu:copy-error', { message: String(error) });
              void copyPlainTextToClipboard(sheet, list, toast);
            },
          );
          return;
        }
        /** 多块仍走纯文本：原生复制只认"最后一个选区"（`getCurrentLastSelection()`），会丢块。
         * 块与块之间空一行分隔（与 Excel 多区域复制一致）；代价是不带格式，菜单标签已写明 */
        void copyPlainTextToClipboard(sheet, list, toast);
        return;
      }
      if (id === 'clear') {
        const list = snapshot.a1List.length > 0 ? snapshot.a1List : [snapshot.a1];
        let cleared = 0;
        let failure: string | null = null;
        for (const ref of list) {
          beginSheetAction(`清空内容 ${ref}`, 'edit');
          const result = clearRange(sheet, ref);
          log('menu:clear', { a1: ref, ok: result.ok, reason: result.reason ?? null });
          if (!result.ok) {
            failure = result.reason ?? '清空失败';
            continue;
          }
          pushHistory(`清空内容 ${ref}`, 'edit');
          cleared += 1;
        }
        if (cleared === 0) {
          toast(failure ?? '清空失败', 'warn');
          return;
        }
        toast(
          list.length > 1
            ? `已清空 ${cleared} 块内容（格式保留）${cleared < list.length ? `，${list.length - cleared} 块失败` : ''}`
            : '已清空内容（格式保留）',
        );
        return;
      }
      if (id === 'undo') {
        void stepBack();
        return;
      }
      if (id === 'redo') {
        void stepForward();
      }
    },
    [addWorkspaceItems, commitWorkspace, getSheet, pasteClipboardAt, pushHistory, stepBack, stepForward, toast],
  );
  /** 面板提交：按 行/列/区域（可多块）取出内容 → 跳过空内容单元格 → 放入工作区（一次提交，一条历史）。
   * `keepSource=false` 即剪切：**先提交工作区、再逐块清空源内容**，顺序反了会出现"内容既不在表里
   * 也不在工作区"的窗口；每块各记一条历史，与右键「剪切」同一套通道 */
  const handleImportSubmit = useCallback(
    ({ kind, value, keepSource }: { kind: ImportTargetKind; value: string; keepSource: boolean }) => {
      const sheet = getSheet();
      if (!sheet) return;
      setImportBusy(true);
      try {
        const ranges = resolveImportRanges(sheet, kind, value);
        if ('error' in ranges) {
          toast(ranges.error, 'warn');
          return;
        }
        const a1List = ranges.a1List;

        const collected = addWorkspaceFromRanges(a1List, { x: window.innerWidth - 150, y: 160 });
        if (!collected.ok) return;

        log('import-dialog:submitted', {
          kind,
          value,
          a1: a1List.join(' '),
          blocks: a1List.length,
          items: collected.items.length,
          keepSource,
        });

        if (!keepSource) {
          let cleared = 0;
          for (const ref of a1List) {
            beginSheetAction(`剪切 ${ref} 到工作区`, 'swap');
            const result = clearRange(sheet, ref);
            log('import-dialog:clear', { a1: ref, ok: result.ok, reason: result.reason ?? null });
            if (!result.ok) {
              toast(result.reason ?? `剪切 ${ref} 失败（内容已加入工作区）`, 'warn');
              continue;
            }
            pushHistory(`剪切 ${ref} 到工作区`, 'swap');
            cleared += 1;
          }
          toast(
            cleared > 0
              ? `已把 ${collected.items.length} 格移入工作区，并清空表格里的 ${cleared} 处源内容（格式保留，可撤销）`
              : `已把 ${collected.items.length} 格加入工作区（源内容清空失败，内容仍在表里）`,
            cleared > 0 ? undefined : 'warn',
          );
        }
        setImportOpen(false);
      } finally {
        setImportBusy(false);
      }
    },
    [addWorkspaceFromRanges, beginSheetAction, getSheet, pushHistory, toast],
  );

  /** 面板的实时预览：读表算"这次会加入哪些格子"（只读，不产生快照与历史） */
  const previewImportTarget = useCallback(
    ({ kind, value }: { kind: ImportTargetKind; value: string }) => {
      const sheet = getSheet();
      if (!sheet) return null;
      const ranges = resolveImportRanges(sheet, kind, value);
      if ('error' in ranges) return null;
      return previewWorkspaceImport(sheet, ranges.a1List);
    },
    [getSheet],
  );

  const hintLabel = useMemo(() => {
    if (!dragging) return null;
    if (hint === 'swap') return '松开即互换内容（双方格式保持不变）';
    if (hint === 'paste') return '松开即写入内容（保留目标格式）';
    if (hint === 'workspace') return '松开即暂存到工作区（可随时拖回表格）';
    // 用"告诉你现在能做什么"的正向措辞："此处不能放置"这类否定式提示容易被误解
    if (hint === 'reject') return '这里放不了：拖到单元格上互换，或拖到右侧工作区暂存';
    return '拖到目标单元格，或拖到右侧工作区暂存';
  }, [dragging, hint]);

  return (
    <div className={`app${dragging ? ' is-dragging' : ''}`}>
      {/* 图标工具栏（无大标题）：打开/导出/撤销/重做/模式切换/历史/日志/工作区加号 */}
      <Toolbar
        fileName={null}
        tabs={tabs}
        activeTabId={activeTabId}
        mode={mode}
        canUndo={canUndo}
        canRedo={canRedo}
        busy={status === 'importing'}
        historyOpen={historyOpen}
        logOpen={logOpen}
        onOpenFile={(file) => void handleFile(file)}
        onExport={(anchor) => setExportMenu((current) => (current ? null : anchor))}
        exportMenuOpen={exportMenu !== null}
        bridgeAvailable={bridge.available}
        canExport={exportMenuItems.length > 0}
        onUndo={() => void stepBack()}
        onRedo={() => void stepForward()}
        onToggleHistory={() =>
          setHistoryOpen((open) => {
            // 两个抽屉几何位置相同，避免互相盖住
            if (!open) setLogOpen(false);
            return !open;
          })
        }
        onToggleLog={() =>
          setLogOpen((open) => {
            if (!open) setHistoryOpen(false);
            return !open;
          })
        }
        onModeChange={handleModeChange}
        onSelectTab={activateTab}
        onCloseTab={closeTab}
        keepSourceOnWorkspaceDrop={settings.keepSourceOnDrop}
        onKeepSourceChange={(keep) => {
          updateSettings({ keepSourceOnDrop: keep });
          log('ui:keep-source', { keep });
        }}
      />

      <div
        className="body"
        style={{ ['--sidebar-w' as string]: `${settings.sidebarWidth}px` }}
        data-sidebar-width={settings.sidebarWidth}
      >
        <main className="stage">
          <div id={CONTAINER_ID} ref={containerRef} />
          {status !== 'ready' && status !== 'importing' ? <div className="stage-overlay">{statusText}</div> : null}
          {/* 把 Excel 工作簿拖到窗口上时的提示（松手即打开） */}
          {fileDropActive ? (
            <div className="file-drop-overlay" data-testid="file-drop-overlay">
              <div className="file-drop-card">
                <strong>松开即打开这个表格</strong>
                <span>{FORMAT_HINT}；一次拖多个会分别开成标签页</span>
              </div>
            </div>
          ) : null}
        </main>

        {/* 拖动这条分隔线即可调整工作区（= 表格区域）宽度 */}
        <div
          className="body-splitter"
          data-testid="sidebar-splitter"
          role="separator"
          aria-orientation="vertical"
          aria-label="拖拽调整工作区宽度"
          title="拖动调整工作区宽度"
          onPointerDown={startSidebarResize}
        />

        <div
          className="workspace-host"
          ref={sidebarRef}
          data-testid="workspace-host"
          // 记下按下的位置：松手时用它区分"点一下空白"与"拖到面板里松手"
          onPointerDown={(event) => {
            workspacePressRef.current = { x: event.clientX, y: event.clientY };
          }}
          // 选择模式下：刚在表格里选过（3 秒内）→ 点这片空白就把选区收进工作区
          onClick={handleWorkspaceBlankClick}
        >
          <WorkspacePanel
            items={items}
            settings={settings}
            onRemove={removeItem}
            onRemoveMany={removeItems}
            onClearAll={clearItems}
            filter={wsFilter}
            onFilterChange={updateWsFilter}
            clearConfirming={wsClearConfirming}
            onRequestClear={() => setWsClearConfirming(true)}
            onCancelClear={() => setWsClearConfirming(false)}
            onItemPointerDown={startWorkspaceDrag}
            onItemContextMenu={(event, item) => {
              event.preventDefault();
              event.stopPropagation();
              setWsMenu({ x: event.clientX, y: event.clientY, item });
            }}
            onAddItem={handleNewWorkspaceItem}
            onSettingsChange={updateSettings}
            pendingItemId={pendingSwapItemId}
          />
        </div>

        {/* 统一提示层：放在 `.body` 里是为了拿到它内联的 `--sidebar-w`（真实工作区宽度），
            这样"表格区域底部居中"才是准的；它是 fixed，不参与这里的网格布局。 */}
        <div className="prompt-layer" data-testid="prompt-layer" aria-live="polite">
          {dragging && hintLabel ? (
            <div
              className={`prompt prompt-hint${
                hint === 'swap' ? ' is-ok' : hint === 'reject' ? ' is-danger' : hint === 'workspace' ? '' : ' is-warn'
              }`}
              data-testid="drag-hint"
              data-hint={hint ?? 'none'}
            >
              {hintLabel}
            </div>
          ) : null}
          {toasts.map((t) => (
            <div key={t.id} className={`prompt prompt-toast${t.kind === 'warn' ? ' is-warn' : ''}`} data-testid="toast">
              {t.text}
            </div>
          ))}
        </div>
      </div>

      {/* 工作区条目右键菜单：复制 / 剪切 / 删除（复用与表格同一个菜单组件，定位/键盘/Esc 都一致） */}
      <ContextMenu
        open={wsMenu !== null}
        x={wsMenu?.x ?? 0}
        y={wsMenu?.y ?? 0}
        items={[
          { id: 'ws-copy', label: '复制', shortcut: '可粘贴到表格' },
          { id: 'ws-cut', label: '剪切', shortcut: '复制并移除' },
          { id: 'ws-delete', label: '删除', danger: true, separatorBefore: true },
        ]}
        onSelect={(id) => {
          const target = wsMenu?.item ?? null;
          setWsMenu(null);
          if (!target) return;
          if (id === 'ws-copy') copyItem(target, false);
          else if (id === 'ws-cut') copyItem(target, true);
          else if (id === 'ws-delete') {
            removeItem(target.id);
            log('workspace:remove', { id: target.id, label: target.label });
          }
        }}
        onClose={() => setWsMenu(null)}
      />

      <footer className="statusbar">
        <span className="chip">{activeSheetName || '—'}</span>
        <span className="chip">选区 {selectionText}</span>
        <span className="chip status-text">{statusText}</span>
        {summary ? (
          <>
            <span className="chip">{summary.cells.toLocaleString()} 单元格</span>
            <span className="chip">
              解析 {summary.parseMs}ms · 转换 {summary.adaptMs}ms · 渲染 {summary.renderMs}ms
            </span>
            {summary.unsupported.length > 0 ? (
              <span className="chip warn" data-testid="unsupported-count">
                未支持 {summary.unsupported.length} 项
              </span>
            ) : null}
            {summary.warnings.length > 0 ? <span className="chip warn">降级 {summary.warnings.length} 项</span> : null}
            {summary.featureCounts ? (
              <span className="chip" data-testid="feature-chip">
                条件格式 {summary.featureCounts.conditionalFormats} · 验证 {summary.featureCounts.dataValidations} · 链接{' '}
                {summary.featureCounts.hyperlinks} · 批注 {summary.featureCounts.notes} · 图片 {summary.featureCounts.images}
              </span>
            ) : null}
          </>
        ) : null}
        <span className="statusbar-spacer" />
        <span className={`chip${fps > 0 && fps < 45 ? ' warn' : ''}`} data-testid="fps">
          {fps > 0 ? `${fps} FPS` : '— FPS'}
        </span>
      </footer>

      {logOpen ? (
        <section className="drawer">
          <div className="drawer-head">
            <strong>运行日志</strong>
            <span>{entries.length} 条</span>
            <button type="button" className="btn ghost" onClick={() => clearLog()}>
              清空
            </button>
            <button type="button" className="btn ghost" onClick={() => setLogOpen(false)}>
              收起
            </button>
          </div>
          {summary ? (
            <div className="import-summary" data-testid="import-summary">
              <div>
                <strong>{summary.fileName}</strong> · {summary.sheets.join(' / ')}
              </div>
              {summary.unsupported.length > 0 ? (
                <div className="summary-block">
                  <em>暂不支持（P1 补齐）</em>
                  <ul>
                    {summary.unsupported.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
              {summary.featureCounts ? (
                <div className="summary-block">
                  <em>P1 特性</em>
                  <div className="feature-counts" data-testid="parsed-features">
                    解析到：条件格式 {summary.parsedFeatures.cf} · 验证 {summary.parsedFeatures.dv} · 链接{' '}
                    {summary.parsedFeatures.links} · 批注 {summary.parsedFeatures.notes} · 表格 {summary.parsedFeatures.tables} · 图片{' '}
                    {summary.parsedFeatures.images}
                  </div>
                  <div className="feature-counts" data-testid="feature-counts">
                    已应用：条件格式 {summary.featureCounts.conditionalFormats} · 数据验证 {summary.featureCounts.dataValidations} ·
                    超链接 {summary.featureCounts.hyperlinks} · 批注 {summary.featureCounts.notes} · 图片 {summary.featureCounts.images}
                    {summary.featureCounts.failed > 0 ? ` · 失败 ${summary.featureCounts.failed}` : ''}
                    {summary.featureCounts.skipped > 0 ? ` · 跳过 ${summary.featureCounts.skipped}` : ''}
                  </div>
                  {summary.featureIssues.length > 0 ? (
                    <ul>
                      {summary.featureIssues.slice(0, 10).map((item) => (
                        <li key={item}>{item}</li>
                      ))}
                    </ul>
                  ) : null}
                </div>
              ) : null}
              {summary.preserved.length > 0 ? (
                <div className="summary-block">
                  <em>已原样保留（不影响预览）</em>
                  <ul data-testid="preserved-list">
                    {summary.preserved.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
              {summary.warnings.length > 0 ? (
                <div className="summary-block">
                  <em>降级与警告</em>
                  <ul>
                    {summary.warnings.slice(0, 8).map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
            </div>
          ) : null}
          <ul className="log-list" data-testid="log-list">
            {entries.slice(-200).reverse().map((entry, index) => (
              <li key={`${entry.t}-${index}`}>
                <span className="t">{entry.t}</span>
                <span className="kind">{entry.kind}</span>
                <span className="detail">{formatDetail(entry.detail)}</span>
              </li>
            ))}
          </ul>
        </section>
      ) : null}

      {/* 右键菜单：动作只含"允许的操作"，样式/结构类一律不提供 */}
      <ContextMenu
        open={menu !== null}
        x={menu?.x ?? 0}
        y={menu?.y ?? 0}
        items={menuItems()}
        onSelect={handleMenuSelect}
        onClose={() => setMenu(null)}
      />

      {/* 导出菜单：默认保真导出，另有 CSV（自研）与 .ods/.xls/.xlsb（本地版经本机 Excel）。 */}
      <ContextMenu
        open={exportMenu !== null}
        x={exportMenu?.x ?? 0}
        y={exportMenu?.y ?? 0}
        items={exportMenuItems}
        onSelect={handleExportMenuSelect}
        onClose={() => setExportMenu(null)}
      />

      {/* 历史记录：列出我们自己的动作，可跳回任意一步（靠连续撤销/重做实现） */}
      <HistoryPanel
        open={historyOpen}
        entries={historyEntries}
        current={historyIndex}
        available={{
          // Univer 的状态计数对"我们自己 push 的撤销项"会少报，因此用历史自身的位置作为可达性下限，
          // 避免面板把可达的跳步误判为不可达
          undos: Math.max(undoRedoCounts.undos, historyIndex),
          redos: Math.max(undoRedoCounts.redos, Math.max(0, historyEntries.length - historyIndex)),
        }}
        onJumpTo={(target) => void handleJumpTo(target)}
        onClose={() => setHistoryOpen(false)}
      />

      {/* 工作区「+」面板：一个输入框 + 实时预览 + 「加入工作区后保留表格内容」（与工具栏开关同一设置） */}
      <ImportDialog
        open={importOpen}
        onClose={() => setImportOpen(false)}
        currentSelectionA1={selectionText}
        sheetName={activeSheetName}
        usedRangeA1={usedRangeA1}
        keepSource={settings.keepSourceOnDrop}
        onKeepSourceChange={(keep) => {
          updateSettings({ keepSourceOnDrop: keep });
          log('ui:keep-source', { keep, from: 'import-dialog' });
        }}
        onPreview={previewImportTarget}
        onSubmit={handleImportSubmit}
        busy={importBusy}
      />
    </div>
  );
}

/**
 * 把若干区域以**纯文本 TSV** 写进系统剪贴板（多块之间空一行分隔）。
 * 多块选区与"单块原生复制失败"的兜底都走它。成功/失败必须如实告诉用户：
 * `navigator.clipboard.writeText()` 在 http 局域网下会整条链短路（既不写也不提示）。
 */
async function copyPlainTextToClipboard(
  sheet: FWorksheet,
  list: string[],
  toast: (text: string, kind?: 'info' | 'warn') => void,
): Promise<void> {
  const text = list
    .map((ref) => sheet.getRange(ref).getDisplayValues().map((row) => row.join('\t')).join('\n'))
    .join('\n\n');
  const ok = await writeClipboardText(text);
  log('menu:copy', { a1: list.join(' '), blocks: list.length, flavor: 'text/plain', ok });
  if (!ok) {
    toast('复制失败：浏览器/系统拒绝了剪贴板写入（可用 HTTPS 或 localhost 打开后重试）', 'warn');
    return;
  }
  toast(list.length > 1 ? `已复制 ${list.length} 块到剪贴板（纯文本，不含格式）` : '已复制到剪贴板（纯文本，不含格式）');
}

function formatDetail(detail: unknown): string {
  if (detail === undefined) return '';
  try {
    const text = JSON.stringify(detail);
    return text.length > 200 ? `${text.slice(0, 200)}…` : text;
  } catch {
    return String(detail);
  }
}

/**
 * 把 Univer 取回的原始值收敛成导出器认识的字面量：富文本单元格的 `getRawValue()` 可能返回
 * 富文本对象，直接写进 xlsx 会变成 "[object Object]"，所以取纯文本；无法收敛的降级为 null。
 */
function normalizeExportValue(value: unknown): CellEdit['value'] {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  const richText = value as { toPlainText?: () => string };
  if (typeof richText.toPlainText === 'function') return richText.toPlainText();
  log('export:unexpected-value', { type: typeof value });
  return null;
}

// ================================================================ 测试钩子
/**
 * `window.__p0` 是 Playwright 依赖的测试契约（`window.__app` 为别名）。
 * 钩子只做"读状态 / 调业务函数"，不模拟 UI，避免测试与实现耦合。
 */
interface TestHooksDeps {
  apiRef: { current: FUniver | null };
  lockRef: { current: ContentOnlyLock | null };
  controllerRef: { current: DragController | null };
  summaryRef: { current: ImportSummary | null };
  featureResultRef: { current: ApplyFeaturesResult | null };
  guardRef: { current: ReadOnlyGuard | null };
  rootInjectorRef: { current: { get: (token: unknown) => unknown } | null };
  tabsRef: { current: Array<{ id: string; fileName: string; source: ExportSource | null; bytes: Uint8Array }> };
  tabRuntimeRef: { current: TabRuntime[] };
  activeTabIdRef: { current: string | null };
  sessionSaverRef: { current: { flush: () => Promise<void> } | null };
  modeRef: { current: InteractionMode };
  snapshotActiveEditsRef: { current: () => void };
  undoRedoCountsRef: { current: { undos: number; redos: number } };
  pushHistory: (label: string, kind: HistoryKind) => void;
  /** 标记一次表格动作开始（让"我们自己的命令"只记一条历史，见 pushHistory 的合并逻辑） */
  beginSheetAction: (label: string, kind: HistoryKind) => void;
  /** 撤销/重做一步（表格动作走 Univer、工作区动作走快照；测试钩子也走它） */
  stepBack: () => Promise<boolean>;
  stepForward: () => Promise<boolean>;
  /** 互换后的收尾（清选中 + 黄色提醒框）；测试钩子也走它，保证与真实交互一致 */
  flashSwap: (firstA1: string, secondA1?: string) => void;
  itemsRef: { get: () => RangeSnapshot[] };
  snapshotStoreRef: { current: Map<string, RangeSnapshot> };
  setItems: Dispatch<SetStateAction<RangeSnapshot[]>>;
  toast: (text: string, kind?: 'info' | 'warn') => void;
  getSheet: () => FWorksheet | null;
  /** 主视口滚动量（滚动条用例断言"拖一下真的连续滚了"） */
  getMainViewportScroll: () => { x: number; y: number };
  /** 滚动条交互带在页面坐标里的位置（e2e 按得准）；thumb 是滑块（只有拖它才连续滚动） */
  getScrollbarBand: () => {
    vertical: { x: number; y: number; width: number; height: number } | null;
    horizon: { x: number; y: number; width: number; height: number } | null;
    verticalThumb: { x: number; y: number; height: number } | null;
    horizonThumb: { y: number; x: number; width: number } | null;
  } | null;
  /** 滚动条命中（e2e 与带子位置做一致性自检） */
  isScrollbarHit: (x: number, y: number) => boolean;
  /** 某格在视口里的矩形（页面坐标）：测试断言"矩形必须包住那一格" */
  measureFlashBox: (row: number, col: number) => { left: number; top: number; width: number; height: number } | null;
  /** 某 A1 的矩形；由装配方用 measureFlashBox 实现，钩子里不必再拿 sheet */
  rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null;
  /** 工作区状态的唯一提交入口（含历史记录与快照）；测试钩子也走它 */
  commitWorkspace: (
    next: RangeSnapshot[] | ((prev: RangeSnapshot[]) => RangeSnapshot[]),
    label: string,
    kind: HistoryKind,
  ) => void;
  /** 选区右下角"填充柄"的当前状态（测试断言它确实没在画，见 src/univer/fill-handle.ts） */
  fillHandleState: () => { available: boolean; controls: number; enabled: boolean[]; visible: (boolean | null)[] };
  /** 自愈动作本体（见它的注释）：e2e 造出"绑定错位"后调它，验证真的能接回来 */
  ensureActiveUnitRendered: (id: string) => Promise<boolean>;
}
function installTestHooks(deps: TestHooksDeps): void {
  const { apiRef, getSheet } = deps;

  const hooks = {
    // ---- 既有契约 ----
    log: p0log,
    clearLog,
    countKind,
    findLast,
    getActiveSheet: getSheet,
    getCellStyle: (row: number, col: number) => getSheet()?.getRange(row, col).getCellStyleData('cell') ?? null,
    getCellStyleByA1: (a1: string) => getSheet()?.getRange(a1).getCellStyleData('cell') ?? null,
    getDisplayValue: (row: number, col: number) => getSheet()?.getRange(row, col).getDisplayValue() ?? null,
    getDisplayValueByA1: (a1: string) => getSheet()?.getRange(a1).getDisplayValue() ?? null,
    getValues: (a1: string) => getSheet()?.getRange(a1).getValues() ?? null,
    getSheetNames: () => apiRef.current?.getActiveWorkbook()?.getSheets().map((s) => s.getSheetName()) ?? [],
    /** 拿到全部工作表 Facade（e2e 断言第二张表时用；避免在测试里直接摸 raw Workbook） */
    getSheets: () => apiRef.current?.getActiveWorkbook()?.getSheets() ?? [],
    getSheetByName: (name: string) => apiRef.current?.getActiveWorkbook()?.getSheetByName(name) ?? null,
    /** 切换活动工作表（走 Facade，保证画布会重绘） */
    activateSheet: (name: string) => {
      apiRef.current?.getActiveWorkbook()?.getSheetByName(name)?.activate();
    },
    getImportSummary: () => deps.summaryRef.current,
    /** 条件格式/数据验证/超链接/批注/图片的应用结果 */
    getFeatureReport: () => deps.featureResultRef.current,
    /** 只读闸门：被拦下的命令及次数（用于验证"样式/结构/对象改不动"） */
    getBlockedCommands: () =>
      [...(deps.guardRef.current?.blockedIds() ?? new Map<string, number>())].map(([id, count]) => ({ id, count })),
    isCommandAllowed: (id: string) => deps.guardRef.current?.isAllowed(id) ?? false,
    // ---- 多标签诊断 ----
    getTabs: () => deps.tabsRef.current.map((tab) => ({ id: tab.id, fileName: tab.fileName })),
    /** 多标签"冷/热"记账：e2e 用它断言常驻窗口真的生效（内存从 O(N) 变 O(K)） */
    getTabRuntime: () => deps.tabRuntimeRef.current.map((tab) => ({ ...tab })),
    residentTabLimit: DEFAULT_RESIDENT_TAB_LIMIT,
    /** 滚动条在画布里的位置（画布内 CSS 像素）；给 e2e 按得准用（画布最右边并不在滚动条上） */
    getScrollbarBand: () => deps.getScrollbarBand(),
    /** 某 A1 的高亮矩形（页面坐标）：测试断言"矩形必须包住那一格"（滚动后尤其要成立） */
    rectOfA1: (a1: string) => deps.rectOfA1(a1),
    /** 选区右下角"填充柄"（那个小方块）的状态：应当恒为关闭/不可见 */
    getFillHandleState: () => deps.fillHandleState(),
    /** 主视口滚动量（滚动条/滚轮用例断言"拖一下真的滚了、而且是一格一格连续滚"） */
    getScrollState: () => deps.getMainViewportScroll(),
    getActiveTabId: () => deps.activeTabIdRef.current,
    /** 直接触发自愈动作（`ensureActiveUnitRendered`），供 e2e 验证"绑定错位真的能被修回来" */
    healActiveUnit: (id: string) => deps.ensureActiveUnitRendered(id),
    getMode: () => deps.modeRef.current,
    /** 当前撤销/重做可用次数（历史跳步与诊断用） */
    getUndoRedoCounts: () => deps.undoRedoCountsRef.current,
    /** 强制立刻落盘（e2e 用；真实场景靠定时保存与 pagehide） */
    flushSession: () => {
      deps.snapshotActiveEditsRef.current();
      return deps.sessionSaverRef.current?.flush() ?? Promise.resolve();
    },
    /** **画布此刻真正绑定的工作簿 id**（Univer 的"当前单元"，不是应用层的活动标签）。两者不一致就是
     * "白屏"类故障的核心特征；没有活动簿时返回 `null`，不抛异常 */
    getActiveWorkbookId: () => {
      try {
        return apiRef.current?.getActiveWorkbook()?.getWorkbook().getUnitId() ?? null;
      } catch {
        return null;
      }
    },
    getDirtySummary: () =>
      deps.tabsRef.current.map((tab) => ({ id: tab.id, dirtyCells: dirtyCellCount(tab.id) })),
    /**
     * 插图 blob url 的**未回收计数**（内存泄漏回归用）：每个 url 钉着一份图片字节，
     * 关标签/冷存/整实例拆卸都必须还回去，读到 0 才说明没漏。
     */
    retainedImageObjectUrls: () => ({
      total: countRetainedImageObjectUrls(),
      byWorkbook: retainedImageObjectUrlsByWorkbook(),
    }),
    /** 直接派发任意命令（用于验证"破坏性命令全被拦下"）；返回命令结果 */
    runCommand: (id: string, params?: unknown) => {
      const sheet = getSheet();
      if (!sheet) return Promise.resolve(false);
      const commandService = sheet.getInject().get(ICommandService);
      return commandService.executeCommand(id, (params ?? {}) as never);
    },
    setValue: (row: number, col: number, value: string | number) => {
      getSheet()?.getRange(row, col).setValue(value);
    },
    setValues: (row: number, col: number, values: (string | number)[][]) => {
      getSheet()?.getRange(row, col, values.length, values[0]?.length ?? 0).setValues(values as never);
    },
    setStyledValues: (row: number, col: number, matrix: unknown[][]) => {
      getSheet()?.getRange(row, col, matrix.length, matrix[0]?.length ?? 0).setValues(matrix as never);
    },
    tryFormatCommand: (row: number, col: number) => {
      const sheet = getSheet();
      if (!sheet) return;
      const commandService = sheet.getInject().get(ICommandService);
      void commandService.executeCommand(SetStyleCommand.id, {
        unitId: sheet.getWorkbook().getUnitId(),
        subUnitId: sheet.getSheetId(),
        range: { startRow: row, startColumn: col, endRow: row, endColumn: col },
        style: { bg: { rgb: '#FF00FF' } },
      } as never);
    },
    swapValues: (a1: string, b1: string) => {
      const sheet = getSheet();
      if (sheet) swapRanges(sheet, a1, b1);
    },
    /** "文本格装数字"的弹窗提醒是否已被关掉（`sheets-ui.config.disableForceStringAlert`）：该配置在
     * `bootUniver` 里是 try/catch 兜底设置的，键名变了会被静默吞掉，所以 e2e 直接读配置确认生效 */
    isForceStringAlertDisabled: () => {
      try {
        // 必须走**根** injector：Facade 的 getInject() 是渲染/子表作用域，里面没有 IConfigService
        const config = deps.rootInjectorRef.current?.get(IConfigService) as
          | { getConfig: (id: string) => unknown }
          | undefined;
        const value = config?.getConfig('sheets-ui.config') as { disableForceStringAlert?: boolean } | null | undefined;
        return value?.disableForceStringAlert === true;
      } catch {
        return false;
      }
    },
    /** 测试专用反向对照：把"关弹窗"配置改回去，验证测试真的能发现那个弹窗（否则断言本身抓不到东西） */
    setForceStringAlertDisabled: (value: boolean) => {
      try {
        const config = deps.rootInjectorRef.current?.get(IConfigService) as
          | { getConfig: (id: string) => unknown; setConfig: (id: string, v: unknown, o?: { merge?: boolean }) => void }
          | undefined;
        if (!config) return null;
        config.setConfig('sheets-ui.config', { disableForceStringAlert: value }, { merge: true });
        return (config.getConfig('sheets-ui.config') as { disableForceStringAlert?: boolean } | null)
          ?.disableForceStringAlert ?? null;
      } catch {
        return null;
      }
    },

    /** 造一条轻提示（e2e 验"提示统一在一个层里"用；等价于点界面上的动作弹提示） */
    toast: (text: string, kind?: 'info' | 'warn') => deps.toast(text, kind),

    // ---- 工作区与拖拽 ----
    getWorkspaceItems: () =>
      deps.itemsRef.get().map((item) => ({
        id: item.id,
        label: item.label,
        rows: item.rows,
        cols: item.cols,
        a1: item.source.a1,
        sheetName: item.source.sheetName,
      })),
    clearWorkspace: () => deps.commitWorkspace([], '清空工作区（测试钩子）', 'workspace'),
    /** 把选区放进工作区（测试钩子）：与生产路径一致，走 `extractCellItems` 拆成独立单元格并跳过空格，
     * 且走 `commitWorkspace`，所以这些条目同样可撤销 */
    snapshotSelectionToWorkspace: (a1?: string) => {
      const sheet = getSheet();
      if (!sheet) return { items: [], skippedEmpty: 0, truncated: 0, error: '无活动工作表' };
      const { items: cells, skippedEmpty, truncated, error } = extractCellItems(sheet, a1);
      if (error) return { items: [], skippedEmpty, truncated, error };
      if (cells.length > 0) {
        deps.commitWorkspace(
          (prev) => [...cells.map((item) => ({ ...item })), ...prev],
          `加入工作区 ${cells.length} 格（测试钩子）`,
          'workspace',
        );
      }
      return { items: cells, skippedEmpty, truncated, ids: cells.map((item) => item.id) };
    },
    /** 把工作区条目写回指定位置（只写内容）——走与拖拽写回**同一条**可撤销通道 */
    pasteWorkspaceItem: (id: string, targetA1: string) => {
      const sheet = getSheet();
      const item = deps.snapshotStoreRef.current.get(id) ?? deps.itemsRef.get().find((i) => i.id === id);
      if (!sheet || !item) return { ok: false, reason: '条目不存在' };
      const rect = sheet.getRange(targetA1).getRange();
      const label = `写入工作区条目 ${item.label}（测试钩子）`;
      deps.beginSheetAction(label, 'edit');
      applySnapshot(sheet, item, { row: rect.startRow, col: rect.startColumn });
      deps.pushHistory(label, 'edit');
      return { ok: true };
    },
    /** 表格内两区域互换（走真实命令，含原子撤销） */
    swap: (a1: string, b1: string) => {
      const sheet = getSheet();
      if (!sheet) return { ok: false, reason: '无活动工作表' };
      deps.beginSheetAction(`互换 ${a1} ⇄ ${b1}`, 'swap');
      const result = swapRanges(sheet, a1, b1);
      if (result.ok) {
        deps.pushHistory(`互换 ${a1} ⇄ ${b1}`, 'swap');
        // 与真实交互路径一致：互换后同样清选中 + 画黄色提醒框（否则钩子用例测的行为与用户看到的不一致）
        deps.flashSwap(a1, b1);
      }
      return result;
    },
    /** 表格内移动区域（走真实命令 + 与拖拽同样的记账通道） */
    move: (sourceA1: string, targetA1: string) => {
      const sheet = getSheet();
      if (!sheet) return { ok: false, reason: '无活动工作表' };
      const rect: IRange = sheet.getRange(targetA1).getRange();
      const label = `移动 ${sourceA1} → ${targetA1}`;
      deps.beginSheetAction(label, 'swap');
      const result = moveRange(sheet, sourceA1, { row: rect.startRow, col: rect.startColumn });
      if (result.ok) deps.pushHistory(label, 'swap');
      return result;
    },
    undo: () => deps.stepBack(),
    redo: () => deps.stepForward(),
    getSelectionA1: () => {
      try {
        return getSheet()?.getActiveRange()?.getA1Notation() ?? null;
      } catch {
        return null;
      }
    },
    selectRange: (a1: string) => {
      const sheet = getSheet();
      if (!sheet) return;
      sheet.setActiveSelection(sheet.getRange(a1));
    },
    /** 造一个几何尺寸已知、无冻结的小表，供真实鼠标手势测试使用 */
    createHandsOnSheet: () => {
      const api = apiRef.current;
      if (!api) return null;
      const id = `hands-on-${Date.now()}`;
      api.createWorkbook({
        id,
        name: '手势测试',
        appVersion: '0.25.1',
        locale: LocaleType.ZH_CN,
        sheetOrder: [`${id}-s1`],
        styles: {},
        sheets: {
          [`${id}-s1`]: {
            id: `${id}-s1`,
            name: '手势',
            rowCount: 60,
            columnCount: 20,
            defaultColumnWidth: 100,
            defaultRowHeight: 30,
            cellData: {
              1: { 1: { v: 'A' }, 2: { v: 'B' }, 3: { v: 'C' } },
              2: { 1: { v: 'D' }, 2: { v: 'E' }, 3: { v: 'F' } },
            },
          },
        },
      } as never);
      return id;
    },
    geometry: () => {
      const container = document.getElementById('univer-container');
      const canvas = container?.querySelector('canvas');
      const rect = canvas?.getBoundingClientRect();
      return rect ? { left: rect.left, top: rect.top, width: rect.width, height: rect.height } : null;
    },
    isDragging: () => deps.controllerRef.current?.active ?? false,
  };

  (window as unknown as Record<string, unknown>).__p0 = hooks;
  (window as unknown as Record<string, unknown>).__app = hooks;
}
