/**
 * 会话持久化（Session Persistence）：把"多标签 / 编辑 / 工作区 / 交互模式"整份存进
 * **浏览器本地 IndexedDB**，误关闭页面后重新打开能恢复现场。纯前端，没有任何服务端。
 *
 * 设计约束与取舍：
 *  - **零新依赖**：手写极简 IndexedDB 封装（open/upgrade + 一个 object store + 一个 key）。
 *  - **不存 ParsedWorkbook**：解析结果又大又能从原始字节重建，所以只存原始 xlsx 字节；
 *    恢复时走"重新解析 originalBytes"或"Univer 快照"两条路。
 *  - **不存样式**：本产品不允许改样式，edits 只记录"内容（值 / 公式）"。
 *  - **Uint8Array 直接存**：IndexedDB 走结构化克隆，转 base64 会白白膨胀 33%。
 *  - **持久化失败绝不打断编辑**：所有 IndexedDB 异常一律吞掉（只 console.warn），
 *    load() 永远返回 `SessionState | null`，绝不抛。
 *  - **非浏览器环境（node 单测 / SSR）**：save/clear 静默成功、load 返回 null、
 *    estimateBytes 返回 0，因此本模块可以在 node 环境的 Vitest 里直接跑。
 */

import type { InteractionMode } from '../interaction/click-swap';
import type { RangeSnapshot } from '../workspace/types';

/* ========================================================================== */
/* 公开契约                                                                    */
/* ========================================================================== */

/**
 * 交互模式（`'drag' | 'click-swap'`）**复用 interaction 层的定义**，不在这里另起一套。
 * 用 `import type` + `export type` 转出去：编译后完全消失，所以本模块运行时零依赖
 * （e2e 里可以单独把它 import 进浏览器跑）。
 */
export type { InteractionMode };

export interface SessionTabEdit {
  sheetId: string;
  row: number;
  col: number;
  /** 只存"内容"：值与公式（样式一律不存，因为本产品不允许改样式） */
  value: string | number | boolean | null;
  formula: string | null;
}

export interface SessionTab {
  id: string;
  fileName: string;
  /** 原始 xlsx 字节（导出时做外科式修补要用；恢复时也要靠它重建 ParsedWorkbook） */
  originalBytes: Uint8Array;
  /** Univer 工作簿快照（univerAPI.getActiveWorkbook() 拿不到时允许为 null，恢复时会走"重新导入+回放编辑"路径） */
  snapshot: unknown | null;
  /** 该标签页内被改过的单元格（按顺序回放即可还原编辑） */
  edits: SessionTabEdit[];
}

/**
 * 工作区/交互的**用户设置**（都要持久化，重开浏览器后保持）。
 *
 * 全部给默认值：旧会话里没有这个字段，缺失或字段非法时按默认值补齐（不让整体迁移失败）。
 */
export interface SessionSettings {
  /** 拖到工作区后是否**保留**表格里的内容（false = 剪切语义） */
  keepSourceOnDrop: boolean;
  /** 把条目拖/写回表格后，是否把该条目从工作区移除 */
  removeItemAfterPaste: boolean;
  /**
   * 工作区格子的最小宽度（px）：面板越宽，每行自动放下的格子越多。
   * 默认 128 → 默认宽度（288px）下正好每行 2 列。
   */
  tileMinWidth: number;
  /** 工作区宽度（px） */
  sidebarWidth: number;
}

export const DEFAULT_SETTINGS: SessionSettings = {
  keepSourceOnDrop: true,
  removeItemAfterPaste: false,
  // 128px：默认面板宽度（288px）下正好每行 2 列；拉宽后自动变成 3、4 列
  tileMinWidth: 128,
  sidebarWidth: 288,
};

/** 工作区格子最小宽度的范围（px）。它是"每行放几格"的唯一旋钮：面板越宽、这个值越小，每行放下的越多。 */
export const TILE_MIN_WIDTH_RANGE = { min: 80, max: 220 } as const;
export const SIDEBAR_WIDTH_RANGE = { min: 200, max: 640 } as const;

function clampInt(value: unknown, fallback: number, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

/** 把任意（可能缺失/损坏）的设置收敛成合法值 */
export function normalizeSettings(raw: unknown): SessionSettings {
  if (!isRecord(raw)) return { ...DEFAULT_SETTINGS };
  const r = raw as Record<string, unknown>;
  return {
    keepSourceOnDrop: typeof r.keepSourceOnDrop === 'boolean' ? r.keepSourceOnDrop : DEFAULT_SETTINGS.keepSourceOnDrop,
    removeItemAfterPaste:
      typeof r.removeItemAfterPaste === 'boolean' ? r.removeItemAfterPaste : DEFAULT_SETTINGS.removeItemAfterPaste,
    tileMinWidth: clampInt(r.tileMinWidth, DEFAULT_SETTINGS.tileMinWidth, TILE_MIN_WIDTH_RANGE.min, TILE_MIN_WIDTH_RANGE.max),
    sidebarWidth: clampInt(r.sidebarWidth, DEFAULT_SETTINGS.sidebarWidth, SIDEBAR_WIDTH_RANGE.min, SIDEBAR_WIDTH_RANGE.max),
  };
}

export interface SessionState {
  version: 1;
  savedAt: number;
  activeTabId: string | null;
  mode: InteractionMode;
  tabs: SessionTab[];
  /** 工作区条目（跨表共享，必须持久化） */
  workspace: RangeSnapshot[];
  /** 用户设置（旧会话没有此字段 → 用默认值补齐） */
  settings: SessionSettings;
}

export interface SessionStore {
  save(state: SessionState): Promise<void>;
  load(): Promise<SessionState | null>;
  clear(): Promise<void>;
  /** 估算当前占用的近似字节数（用于 UI 提示） */
  estimateBytes(): Promise<number>;
}

/* ========================================================================== */
/* 常量                                                                        */
/* ========================================================================== */

/** schema 版本。只有 1；将来加字段时在这里 +1 并在 migrateOrNull 里写迁移分支。 */
const DB_VERSION = 1;
const STORE_NAME = 'state';
const RECORD_KEY = 'session';
const DEFAULT_DB_NAME = 'excel-preview-session';

/**
 * open 请求的兜底超时。正常路径上不可能触发（同版本 open 不会被 block），
 * 但"恢复会话"在启动链路上，宁可 5 秒后当作没有会话，也不能让 UI 永远转圈。
 */
const OPEN_TIMEOUT_MS = 5_000;

/** estimateStateBytes 里替换 Uint8Array 的占位符（见该函数注释） */
const BYTES_PLACEHOLDER = '<originalBytes>';

function warn(message: string, error?: unknown): void {
  // 用 console.warn 而不是 throw：本地持久化只是"锦上添花"，坏了不能影响编辑
  console.warn(`[persistence] ${message}`, error === undefined ? '' : error);
}

function noop(): void {
  /* 用于吞掉 promise 结果 */
}

/** 每次调用现查，避免模块加载时机影响判断（也便于单测注入假实现） */
function hasIndexedDB(): boolean {
  return typeof indexedDB !== 'undefined' && indexedDB !== null;
}

/* ========================================================================== */
/* 极简 IndexedDB 封装                                                         */
/* ========================================================================== */

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'));
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error('IndexedDB transaction aborted'));
    tx.onerror = () => reject(tx.error ?? new Error('IndexedDB transaction failed'));
  });
}

/**
 * 打开（必要时创建）数据库。任何失败都以 `null` 收场，绝不 reject。
 * 每次操作都新开连接、用完就关，避免长期持有连接把 deleteDatabase / versionchange 卡住。
 */
function openDatabase(name: string): Promise<IDBDatabase | null> {
  return new Promise<IDBDatabase | null>((resolve) => {
    let settled = false;
    let timer: ReturnType<typeof setTimeout> | null = null;

    const finish = (db: IDBDatabase | null): void => {
      if (settled) return;
      settled = true;
      if (timer !== null) {
        clearTimeout(timer);
        timer = null;
      }
      resolve(db);
    };

    let request: IDBOpenDBRequest;
    try {
      request = indexedDB.open(name, DB_VERSION);
    } catch (error) {
      // 例如无痕模式 / 被策略禁用的 iframe：open 会同步抛 SecurityError
      warn('indexedDB.open 抛异常，按"无本地会话"处理', error);
      finish(null);
      return;
    }

    request.onupgradeneeded = () => {
      const db = request.result;
      try {
        if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
      } catch (error) {
        warn(`创建 object store "${STORE_NAME}" 失败`, error);
      }
    };

    request.onsuccess = () => {
      const db = request.result;
      if (settled) {
        // 超时后才成功：连接没人用了，必须关掉
        try {
          db.close();
        } catch {
          /* 忽略 */
        }
        return;
      }
      // 别的标签页要升级数据库时主动让路，否则会把它 block 住
      db.onversionchange = () => {
        try {
          db.close();
        } catch {
          /* 忽略 */
        }
      };
      finish(db);
    };

    request.onerror = () => {
      // 版本不匹配（VersionError）也走这里：**故意不删库**。
      // 能造成版本不匹配的只有"更新版本的应用写过同一个库名"，这时候删库等于
      // 毁掉新版本的数据；我们只当作"本地没有可用的会话"。
      const error = request.error;
      if (error && error.name === 'VersionError') {
        warn(`本地库 "${name}" 的 schema 版本与应用不一致，忽略已存数据`);
      } else {
        warn(`打开本地库 "${name}" 失败`, error);
      }
      finish(null);
    };

    request.onblocked = () => {
      // 只有在"别的连接持有更旧版本且不肯关闭"时才会触发；同版本下不会发生。
      // 这里不 resolve，等 onsuccess/onerror，超时兜底。
      warn(`本地库 "${name}" 被其它连接占用，等待中`);
    };

    timer = setTimeout(() => {
      warn(`打开本地库 "${name}" 超时（${OPEN_TIMEOUT_MS}ms），按"无本地会话"处理`);
      finish(null);
    }, OPEN_TIMEOUT_MS);
  });
}

async function readRecord(db: IDBDatabase): Promise<unknown> {
  const tx = db.transaction(STORE_NAME, 'readonly');
  const value = await requestToPromise<unknown>(tx.objectStore(STORE_NAME).get(RECORD_KEY));
  await transactionDone(tx);
  return value;
}

async function writeRecord(db: IDBDatabase, value: SessionState): Promise<void> {
  const tx = db.transaction(STORE_NAME, 'readwrite');
  // 结构化克隆：Uint8Array 直接进库，不做 base64
  tx.objectStore(STORE_NAME).put(value, RECORD_KEY);
  await transactionDone(tx);
}

async function deleteRecord(db: IDBDatabase): Promise<void> {
  const tx = db.transaction(STORE_NAME, 'readwrite');
  tx.objectStore(STORE_NAME).delete(RECORD_KEY);
  await transactionDone(tx);
}

/** 统一的"拿连接 → 干活 → 关连接"外壳；失败返回 fallback，绝不抛 */
async function withDatabase<T>(
  dbName: string,
  run: (db: IDBDatabase) => Promise<T>,
  fallback: T,
): Promise<T> {
  if (!hasIndexedDB()) return fallback;
  let db: IDBDatabase | null = null;
  try {
    db = await openDatabase(dbName);
    if (!db) return fallback;
    return await run(db);
  } catch (error) {
    warn('IndexedDB 操作失败（已忽略）', error);
    return fallback;
  } finally {
    if (db) {
      try {
        db.close();
      } catch {
        /* 忽略 */
      }
    }
  }
}

/* ========================================================================== */
/* 数据校验 / 迁移                                                             */
/* ========================================================================== */

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isCellIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0;
}

function isCellValue(value: unknown): value is string | number | boolean | null {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
}

/**
 * 运行时可接受的 mode 值。`satisfies` 保证它不会和 `InteractionMode` 联合类型跑偏
 * （interaction 层改名/删值会在这里编译报错）；新增值需要同步加进来。
 */
const INTERACTION_MODES = ['select', 'drag', 'click-swap'] as const satisfies readonly InteractionMode[];

function isInteractionMode(value: unknown): value is InteractionMode {
  return INTERACTION_MODES.some((mode) => mode === value);
}

/**
 * 把库里读出来的东西变成 Uint8Array。
 * 只接受"字节类"输入（Uint8Array / ArrayBuffer / TypedArray 视图），
 * **不接受 base64 字符串**——我们从不用那种形式存，出现即视为损坏。
 */
function toUint8Array(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  return null;
}

function migrateEditOrNull(raw: unknown): SessionTabEdit | null {
  if (!isRecord(raw)) return null;
  const { sheetId, row, col, value, formula } = raw;
  if (typeof sheetId !== 'string') return null;
  if (!isCellIndex(row) || !isCellIndex(col)) return null;
  if (!isCellValue(value)) return null;
  // formula 缺省与 null 等价（都表示"这个格子没有公式"），故容忍 undefined
  if (!(formula === undefined || formula === null || typeof formula === 'string')) return null;
  return { sheetId, row, col, value, formula: formula ?? null };
}

function migrateTabOrNull(raw: unknown): SessionTab | null {
  if (!isRecord(raw)) return null;
  const { id, fileName, originalBytes, snapshot, edits } = raw;
  if (typeof id !== 'string' || typeof fileName !== 'string') return null;
  const bytes = toUint8Array(originalBytes);
  if (!bytes) return null;
  if (!Array.isArray(edits)) return null;

  const migrated: SessionTabEdit[] = [];
  for (const edit of edits) {
    const parsed = migrateEditOrNull(edit);
    if (!parsed) return null;
    migrated.push(parsed);
  }

  return {
    id,
    fileName,
    originalBytes: bytes,
    // Univer 快照一定是普通对象；不是对象就等于没有快照（恢复时走重新导入路径）
    snapshot: isRecord(snapshot) ? snapshot : null,
    edits: migrated,
  };
}

/**
 * 校验 + 迁移本地数据。任何结构性损坏都返回 `null`（由调用方负责清理坏数据）。
 *
 * 严格程度的分工：
 *  - 顶层字段（version / savedAt / mode / tabs / workspace）**必须齐**，缺一个就算损坏；
 *  - tabs / edits 逐项严格校验（它们直接决定能否还原现场，宁可不恢复也不能还原出错的现场）；
 *  - activeTabId / snapshot / formula 这三个"本来就可以是 null"的字段容忍缺省；
 *  - workspace 项只校验"是对象且有 string id"：RangeSnapshot 归 workspace 模块所有，
 *    在这里做全字段深校验会把那个类型的正常演进误判成本地数据损坏。
 */
export function migrateOrNull(raw: unknown): SessionState | null {
  if (!isRecord(raw)) return null;
  if (raw.version !== 1) return null;

  const { savedAt, activeTabId, mode, tabs, workspace } = raw;

  if (typeof savedAt !== 'number' || !Number.isFinite(savedAt)) return null;
  if (!(activeTabId === undefined || activeTabId === null || typeof activeTabId === 'string')) return null;
  if (!isInteractionMode(mode)) return null;
  if (!Array.isArray(tabs)) return null;
  if (!Array.isArray(workspace)) return null;

  const migratedTabs: SessionTab[] = [];
  for (const tab of tabs) {
    const parsed = migrateTabOrNull(tab);
    if (!parsed) return null;
    migratedTabs.push(parsed);
  }

  const migratedWorkspace: RangeSnapshot[] = [];
  for (const item of workspace) {
    if (!isRecord(item) || typeof item.id !== 'string') return null;
    migratedWorkspace.push(item as unknown as RangeSnapshot);
  }

  return {
    version: 1,
    savedAt,
    activeTabId: activeTabId ?? null,
    mode,
    tabs: migratedTabs,
    workspace: migratedWorkspace,
    // 旧会话没有 settings → 用默认值补齐；有但字段非法 → 逐字段收敛
    settings: normalizeSettings((raw as { settings?: unknown }).settings),
  };
}

/* ========================================================================== */
/* 体积估算                                                                    */
/* ========================================================================== */

/**
 * 近似占用字节数 = 原始 xlsx 字节总长 + 其余数据的 JSON 文本长度 × 2。
 *
 * 两个关键取舍：
 *  1. 序列化前把每个 tab 的 `originalBytes` 换成占位符。否则 `JSON.stringify` 会把
 *     Uint8Array 摊成 `{"0":12,"1":34,...}`，既慢又离谱（同一个字节能膨胀到 4 个字符）。
 *  2. 文本按 2 字节/字符算（UTF-8 下 ASCII 1 字节、中文 3 字节的折中），只要量级对得上，
 *     够 UI 提示用；不做精确计算、也不去真读全部数据。
 *
 * 注意：快照里如果有循环引用或 BigInt，JSON.stringify 会抛——这里吞掉并按"文本部分为 0"算，
 * 保证估算永远不会把 UI 弄崩。
 */
export function estimateStateBytes(state: SessionState): number {
  let rawBytes = 0;
  for (const tab of state.tabs) {
    if (tab.originalBytes instanceof Uint8Array) rawBytes += tab.originalBytes.byteLength;
  }

  let jsonChars = 0;
  try {
    const stripped = {
      ...state,
      tabs: state.tabs.map((tab) => ({ ...tab, originalBytes: BYTES_PLACEHOLDER })),
    };
    jsonChars = (JSON.stringify(stripped) ?? '').length;
  } catch {
    jsonChars = 0;
  }

  return Math.round(rawBytes + jsonChars * 2);
}

/* ========================================================================== */
/* SessionStore                                                                */
/* ========================================================================== */

export function createSessionStore(dbName: string = DEFAULT_DB_NAME): SessionStore {
  /**
   * 最近一次写入/读出的体积缓存：estimateBytes() 是给 UI 频繁调用的，
   * 有缓存就不必为了报个数字把整条记录（含几 MB 的 xlsx 字节）再读一遍。
   * -1 表示"还不知道"，此时才回落到一次读。
   */
  let cachedBytes = -1;

  return {
    async save(state: SessionState): Promise<void> {
      const size = estimateStateBytes(state);
      const ok = await withDatabase(
        dbName,
        async (db) => {
          await writeRecord(db, state);
          return true;
        },
        false,
      );
      if (ok) cachedBytes = size;
    },

    async load(): Promise<SessionState | null> {
      return withDatabase<SessionState | null>(
        dbName,
        async (db) => {
          const raw = await readRecord(db);
          // undefined = 这个 key 从来没写过 → 正常"全新开始"，没什么可清的
          if (raw === undefined) return null;

          // 其余情况（含被别的东西写成 null / 字符串）一律交给 migrateOrNull 判定
          const state = migrateOrNull(raw);
          if (state) {
            cachedBytes = estimateStateBytes(state);
            return state;
          }

          // 结构不合法（版本不符 / 缺字段 / 数据损坏）：返回 null 并清掉，
          // 免得每次启动都踩同一颗雷
          warn('本地会话数据不合法，已清理，按"全新开始"处理');
          await deleteRecord(db);
          cachedBytes = 0;
          return null;
        },
        null,
      );
    },

    async clear(): Promise<void> {
      cachedBytes = 0;
      await withDatabase(
        dbName,
        async (db) => {
          await deleteRecord(db);
          return true;
        },
        false,
      );
    },

    async estimateBytes(): Promise<number> {
      if (cachedBytes >= 0) return cachedBytes;
      return withDatabase(
        dbName,
        async (db) => {
          const raw = await readRecord(db);
          if (raw === undefined || raw === null) return 0;
          const state = migrateOrNull(raw);
          if (!state) return 0;
          const size = estimateStateBytes(state);
          cachedBytes = size;
          return size;
        },
        0,
      );
    },
  };
}

/* ========================================================================== */
/* 自动保存                                                                    */
/* ========================================================================== */

export function createAutoSaver(
  store: SessionStore,
  getState: () => SessionState,
  debounceMs = 800,
): { trigger(): void; flush(): Promise<void>; dispose(): void } {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let disposed = false;
  /** 串行化链：每次写入都挂在上一次之后，保证同一个 key 不会被并发写 */
  let tail: Promise<void> = Promise.resolve();

  function cancelTimer(): void {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  /**
   * 排队一次写入。返回的 promise 在**这次**写入结束时 resolve（永不 reject，
   * 这样调用方不会因为本地存储出错而拿到异常）。
   */
  function enqueue(): Promise<void> {
    const run = tail.then(async () => {
      // 排队期间被 dispose 了就直接放弃，不再落盘
      if (disposed) return;
      try {
        await store.save(getState());
      } catch (error) {
        warn('自动保存失败（已忽略，不影响编辑）', error);
      }
    });
    // tail 必须"永不 reject"，否则一次失败会把后续所有写入一起带走
    tail = run.then(noop, noop);
    return tail;
  }

  return {
    trigger(): void {
      if (disposed) return;
      cancelTimer();
      timer = setTimeout(() => {
        timer = null;
        void enqueue();
      }, Math.max(0, debounceMs));
    },

    async flush(): Promise<void> {
      // flush 的语义是"确保当前状态已经落盘"，所以不等去抖定时器、直接补一次写入
      // （getState 是纯读取，重复写同一个 key 是幂等的）
      cancelTimer();
      if (disposed) return;
      await enqueue();
    },

    dispose(): void {
      disposed = true;
      cancelTimer();
      // 已经在写的那一次无法取消，会正常写完；排队中但还没开始的会被跳过
    },
  };
}
