/**
 * 会话持久化单测（**node 环境，没有 indexedDB**）。
 *
 * 共 **9 个用例**：相关性强的场景合并进同一个 `it`（顺序执行的多个 scenario 块），
 * 但每条行为断言都保留。覆盖四条最容易出错的线：
 *  1. 无 indexedDB 时"静默降级"：load→null、save/clear/estimateBytes 不抛（保证 App 在任何环境都能启动）
 *  2. estimateStateBytes：不把 Uint8Array 摊成逐元素字符串，循环引用不抛
 *  3. migrateOrNull 的容错边界：坏数据一律 null，三种模式都合法，合法数据原样返回、缺省字段归一化
 *  4. createAutoSaver：去抖合并 / flush 立即落盘 / **写入串行化**（不并发写同一个 key）/
 *     dispose 之后彻底闭嘴 / **store 抛异常（配额超限）不影响后续写入**
 *
 * 注：真实 IndexedDB 的"存得进、读得出、Uint8Array 不丢"由 tests/e2e/persistence.spec.ts 在
 * 真浏览器里验证（node 环境无法覆盖）。
 */
import { describe, expect, it } from 'vitest';

import {
  DEFAULT_SETTINGS,
  createAutoSaver,
  createSessionStore,
  estimateStateBytes,
  migrateOrNull,
  type SessionState,
  type SessionStore,
  type SessionTab,
  type SessionTabEdit,
} from '../../src/persistence/session';
import type { RangeSnapshot } from '../../src/workspace/types';

/* -------------------------------------------------------------------------- */
/* 测试夹具                                                                    */
/* -------------------------------------------------------------------------- */

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

/** 轮询等待条件成立（避免用固定 sleep 跟真实定时器赛跑） */
async function until(predicate: () => boolean, timeoutMs = 2_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error('until(): 等待超时');
    await sleep(5);
  }
}

function makeEdit(over: Partial<SessionTabEdit> = {}): SessionTabEdit {
  return { sheetId: 'sheet-1', row: 0, col: 0, value: '已编辑', formula: null, ...over };
}

function makeTab(over: Partial<SessionTab> = {}): SessionTab {
  return {
    id: 'tab-1',
    fileName: '成绩表.xlsx',
    originalBytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04]), // "PK\x03\x04"，xlsx 就是 zip
    snapshot: { id: 'workbook-1', sheetOrder: ['sheet-1'] },
    edits: [makeEdit()],
    ...over,
  };
}

function makeWorkspaceItem(over: Partial<RangeSnapshot> = {}): RangeSnapshot {
  return {
    id: 'ws-1',
    source: { sheetId: 'sheet-1', sheetName: 'Sheet1', a1: 'A1:B2', startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
    rows: 2,
    cols: 2,
    cells: [[{ text: '1', value: 1 }], [{ text: '2', value: 2 }]],
    values: [[1]],
    formulas: [[null]],
    createdAt: 1_700_000_000_000,
    label: 'A1:B2',
    ...over,
  };
}

/** 一份结构完全合法的会话状态 */
function makeState(over: Partial<SessionState> = {}): SessionState {
  return {
    version: 1,
    savedAt: 1_700_000_000_000,
    activeTabId: 'tab-1',
    mode: 'click-swap',
    tabs: [makeTab()],
    workspace: [makeWorkspaceItem()],
    settings: DEFAULT_SETTINGS,
    ...over,
  };
}

/** 记录调用次数 + 并发峰值的假 store：用来验证去抖与串行化 */
class FakeStore implements SessionStore {
  saves = 0;
  /** 当前正在执行的 save 数量 */
  active = 0;
  /** 历史峰值：必须恒为 1，否则说明发生了并发写同一个 key */
  maxActive = 0;
  savedStates: SessionState[] = [];

  constructor(private readonly delayMs = 0) {}

  async save(state: SessionState): Promise<void> {
    this.saves += 1;
    this.active += 1;
    this.maxActive = Math.max(this.maxActive, this.active);
    this.savedStates.push(state);
    try {
      if (this.delayMs > 0) await sleep(this.delayMs);
    } finally {
      this.active -= 1;
    }
  }

  async load(): Promise<SessionState | null> {
    return null;
  }

  async clear(): Promise<void> {
    /* 无关 */
  }

  async estimateBytes(): Promise<number> {
    return 0;
  }
}

/* -------------------------------------------------------------------------- */
/* 1. 无 indexedDB 环境（node）                                                */
/* -------------------------------------------------------------------------- */

describe('createSessionStore - 无 indexedDB 的 node 环境', () => {
  it('load 返回 null、save/clear/estimateBytes 静默成功，不传库名也能用', async () => {
    const store = createSessionStore('persistence-unit-no-idb');
    await expect(store.load()).resolves.toBeNull();
    await expect(store.save(makeState())).resolves.toBeUndefined();
    await expect(store.clear()).resolves.toBeUndefined();
    await expect(store.estimateBytes()).resolves.toBe(0);
    // 默认库名
    await expect(createSessionStore().load()).resolves.toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* 2. estimateStateBytes                                                       */
/* -------------------------------------------------------------------------- */

describe('estimateStateBytes - 体积近似', () => {
  it('含 Uint8Array 时给出贴近真实字节的正数、不逐元素摊开、随数据变大、循环引用不抛', () => {
    // (a) 量级贴近真实字节数：4KB 数据 + 几百字节 JSON → 介于 1x 和 4x 之间
    const bytes = new Uint8Array(4096);
    for (let i = 0; i < bytes.length; i++) bytes[i] = i % 251;
    const state = makeState({ tabs: [makeTab({ originalBytes: bytes })] });

    const estimate = estimateStateBytes(state);

    expect(estimate).toBeGreaterThan(4096);
    expect(estimate).toBeLessThan(4096 * 4); // 4KB 数据 + 几百字节 JSON → 不可能超过 4 倍

    // (b) 不把 Uint8Array 摊成 {"0":..,"1":..} 逐元素字符串
    const plain = new Uint8Array(4096);
    const plainState = makeState({ tabs: [makeTab({ originalBytes: plain })] });
    // 朴素的 JSON.stringify 会膨胀到 ~24KB；我们的估算必须远小于它
    expect(JSON.stringify(plainState).length).toBeGreaterThan(4096 * 4);
    expect(estimateStateBytes(plainState)).toBeLessThan(JSON.stringify(plainState).length);

    // (c) 数据变大时估算随之变大
    const small = estimateStateBytes(makeState({ tabs: [], workspace: [] }));
    expect(small).toBeGreaterThan(0);
    expect(estimateStateBytes(makeState({ tabs: [makeTab({ originalBytes: new Uint8Array(64 * 1024) })] })))
      .toBeGreaterThan(small + 64 * 1024 - 1);

    // (d) 快照里出现循环引用时不抛（走"文本部分算 0"的兜底）
    const circular: Record<string, unknown> = { id: 'wb' };
    circular.self = circular;
    const circularEstimate = estimateStateBytes(
      makeState({ tabs: [makeTab({ originalBytes: new Uint8Array(100), snapshot: circular })] }),
    );
    expect(circularEstimate).toBeGreaterThanOrEqual(100);
    expect(Number.isFinite(circularEstimate)).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* 3. migrateOrNull                                                            */
/* -------------------------------------------------------------------------- */

describe('migrateOrNull - 合法数据', () => {
  it('原样返回（含结构化克隆往返）、Uint8Array 不丢、空会话合法、缺省字段归一化成 null、不改入参', () => {
    const state = makeState();
    const migrated = migrateOrNull(state);
    expect(migrated).toEqual(state);
    expect(migrated?.tabs[0]?.originalBytes).toBeInstanceOf(Uint8Array);
    expect(Array.from(migrated?.tabs[0]?.originalBytes ?? [])).toEqual([0x50, 0x4b, 0x03, 0x04]);
    // 模拟真实的 IndexedDB 往返
    const cloned: unknown = structuredClone(state);
    expect(cloned).not.toBe(state);
    expect(migrateOrNull(cloned)).toEqual(state);

    // 空会话（没有标签、没有工作区）也是合法的
    const empty = makeState({ tabs: [], workspace: [], activeTabId: null, mode: 'drag' });
    expect(migrateOrNull(empty)).toEqual(empty);

    // activeTabId / snapshot / formula 缺省时归一化成 null，且返回新对象不改入参
    const sparse = makeState({
      tabs: [{
        id: 't',
        fileName: 'a.xlsx',
        originalBytes: new Uint8Array(1),
        edits: [{ sheetId: 's', row: 1, col: 2, value: 3 }],
      } as unknown as SessionTab],
    });
    delete (sparse as unknown as Record<string, unknown>).activeTabId;

    const normalized = migrateOrNull(sparse);
    expect(normalized).not.toBeNull();
    expect(normalized).not.toBe(sparse);
    expect(normalized?.activeTabId).toBeNull();
    expect(normalized?.tabs[0]?.snapshot).toBeNull();
    expect(normalized?.tabs[0]?.edits[0]).toEqual({ sheetId: 's', row: 1, col: 2, value: 3, formula: null });
    expect(sparse.version).toBe(1);
  });
});

describe('migrateOrNull - 坏数据一律 null', () => {
  it('版本 / 顶层形态 / 缺字段 / 类型不对都判为损坏', () => {
    expect(migrateOrNull(makeState({ version: 2 as unknown as 1 }))).toBeNull();
    expect(migrateOrNull(makeState({ version: 0 as unknown as 1 }))).toBeNull();
    expect(migrateOrNull(makeState({ version: '1' as unknown as 1 }))).toBeNull();
    const noVersion: Record<string, unknown> = { ...makeState() };
    delete noVersion.version;
    expect(migrateOrNull(noVersion)).toBeNull();

    expect(migrateOrNull(null)).toBeNull();
    expect(migrateOrNull(undefined)).toBeNull();
    expect(migrateOrNull('session')).toBeNull();
    expect(migrateOrNull(42)).toBeNull();
    expect(migrateOrNull([])).toBeNull();
    expect(migrateOrNull([makeState()])).toBeNull();

    for (const key of ['savedAt', 'mode', 'tabs', 'workspace']) {
      const broken: Record<string, unknown> = { ...makeState() };
      delete broken[key];
      expect(migrateOrNull(broken), `缺 ${key} 必须判为损坏`).toBeNull();
    }

    expect(migrateOrNull({ ...makeState(), tabs: null })).toBeNull();
    expect(migrateOrNull({ ...makeState(), tabs: {} })).toBeNull();
    expect(migrateOrNull({ ...makeState(), tabs: 'tab' })).toBeNull();
    expect(migrateOrNull({ ...makeState(), workspace: null })).toBeNull();
    expect(migrateOrNull({ ...makeState(), workspace: 'ws' })).toBeNull();
    expect(migrateOrNull({ ...makeState(), savedAt: 'now' })).toBeNull();
    expect(migrateOrNull({ ...makeState(), savedAt: Number.NaN })).toBeNull();
    expect(migrateOrNull({ ...makeState(), savedAt: Number.POSITIVE_INFINITY })).toBeNull();
    expect(migrateOrNull({ ...makeState(), activeTabId: 7 })).toBeNull();
  });

  it('三种模式都合法、其它 mode → null，且 tab / originalBytes / edit / 工作区条目任一处坏掉 → null', () => {
    for (const mode of ['select', 'drag', 'click-swap'] as const) {
      const state = migrateOrNull({ ...makeState(), mode });
      expect(state?.mode, `${mode} 应被接受`).toBe(mode);
    }
    expect(migrateOrNull({ ...makeState(), mode: 'keyboard' })).toBeNull();
    expect(migrateOrNull({ ...makeState(), mode: null })).toBeNull();

    const withTabs = (tabs: unknown): unknown => ({ ...makeState(), tabs });
    expect(migrateOrNull(withTabs([{ fileName: 'a.xlsx', originalBytes: new Uint8Array(1), edits: [] }]))).toBeNull();
    expect(migrateOrNull(withTabs([{ id: 't', fileName: 'a.xlsx', edits: [] }]))).toBeNull();
    expect(migrateOrNull(withTabs([{ id: 't', fileName: 'a.xlsx', originalBytes: new Uint8Array(1) }]))).toBeNull();
    expect(migrateOrNull(withTabs([null]))).toBeNull();
    // originalBytes 被存成了 base64 字符串 / 普通数组
    expect(migrateOrNull(withTabs([{ ...makeTab(), originalBytes: 'UEsDBA==' }]))).toBeNull();
    expect(migrateOrNull(withTabs([{ ...makeTab(), originalBytes: [1, 2, 3] }]))).toBeNull();

    const withEdit = (edit: unknown): unknown => ({ ...makeState(), tabs: [{ ...makeTab(), edits: [edit] }] });
    expect(migrateOrNull(withEdit(makeEdit({ row: -1 })))).toBeNull();
    expect(migrateOrNull(withEdit(makeEdit({ col: 1.5 })))).toBeNull();
    expect(migrateOrNull(withEdit({ ...makeEdit(), value: { a: 1 } }))).toBeNull();
    expect(migrateOrNull(withEdit({ ...makeEdit(), formula: 42 }))).toBeNull();
    expect(migrateOrNull(withEdit(undefined))).toBeNull();

    expect(migrateOrNull({ ...makeState(), workspace: [{ source: {} }] })).toBeNull();
    expect(migrateOrNull({ ...makeState(), workspace: ['ws-1'] })).toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* 4. createAutoSaver                                                          */
/* -------------------------------------------------------------------------- */

describe('createAutoSaver - 去抖合并与 dispose', () => {
  it('trigger() 连打只落盘 1 次、默认去抖约 800ms；dispose() 取消未到点的定时器，之后 trigger/flush 都不再落盘', async () => {
    // (a) 去抖合并：窗口内连打 10 次 → 只写 1 次
    const store = new FakeStore();
    const saver = createAutoSaver(store, () => makeState(), 25);

    for (let i = 0; i < 10; i++) saver.trigger();
    expect(store.saves, '去抖窗口内不应有任何写入').toBe(0);

    await sleep(150);
    expect(store.saves).toBe(1);
    expect(store.savedStates).toHaveLength(1);
    saver.dispose();

    // (b) 默认窗口约 800ms：500ms 内不会落盘，但最终必然落盘
    const slow = new FakeStore();
    const defaultSaver = createAutoSaver(slow, () => makeState());
    defaultSaver.trigger();
    await sleep(500);
    expect(slow.saves, '默认去抖窗口内不应落盘').toBe(0);
    await until(() => slow.saves === 1, 1_500);
    expect(slow.saves).toBe(1);
    defaultSaver.dispose();

    // (c) dispose() 取消还没到点的去抖定时器，之后什么都不写
    const cancelled = new FakeStore();
    const cancelledSaver = createAutoSaver(cancelled, () => makeState(), 20);
    cancelledSaver.trigger();
    cancelledSaver.dispose();
    await sleep(120);
    expect(cancelled.saves).toBe(0);

    // (d) 已 dispose 的 saver：trigger / flush 都不落盘（且不抛）
    const after = new FakeStore();
    const disposed = createAutoSaver(after, () => makeState(), 10);
    disposed.dispose();
    for (let i = 0; i < 5; i++) disposed.trigger();
    await expect(disposed.flush()).resolves.toBeUndefined();
    await sleep(100);
    expect(after.saves).toBe(0);
  });
});

describe('createAutoSaver - flush', () => {
  it('flush() 立即落盘、写的是调用时的最新状态、不补出第二次写入；无待写内容时也补一次', async () => {
    const store = new FakeStore();
    let current = makeState({ savedAt: 1 });
    const saver = createAutoSaver(store, () => current, 5_000);

    const startedAt = Date.now();
    saver.trigger();
    current = makeState({ savedAt: 2, mode: 'drag' });
    await saver.flush();

    expect(store.saves).toBe(1);
    expect(Date.now() - startedAt, 'flush 不应等 5s 的去抖窗口').toBeLessThan(1_000);
    expect(store.savedStates[0]?.savedAt).toBe(2);
    expect(store.savedStates[0]?.mode).toBe('drag');

    // 去抖定时器已被 flush 取消，之后不该再冒出第二次写入
    await sleep(80);
    expect(store.saves).toBe(1);
    saver.dispose();

    // 没有待写内容时 flush() 也补一次写入（flush = "确保已落盘"）
    const idle = new FakeStore();
    const idleSaver = createAutoSaver(idle, () => makeState(), 5_000);
    await idleSaver.flush();
    expect(idle.saves).toBe(1);
    idleSaver.dispose();
  });
});

describe('createAutoSaver - 串行化（同一个 key 绝不并发写）', () => {
  it('连续 trigger + flush、20 次同步 flush 都串行（maxActive===1），慢写入只排队，排队中 dispose 跳过未开始的写入', async () => {
    // (a) 连续 trigger + flush：写入次数符合预期，且同时只有一次在写
    const store = new FakeStore(20);
    // 去抖窗口设得很大，确保测试期间只有 flush 会触发写入
    const saver = createAutoSaver(store, () => makeState(), 60_000);

    saver.trigger();
    saver.trigger();
    saver.trigger();
    const f1 = saver.flush();
    saver.trigger();
    const f2 = saver.flush();
    saver.trigger();
    const f3 = saver.flush();
    await Promise.all([f1, f2, f3]);

    expect(store.saves).toBe(3);
    expect(store.maxActive, '写入必须串行：任何时刻只能有一次 save 在执行').toBe(1);

    await sleep(100);
    expect(store.saves).toBe(3);
    expect(store.maxActive).toBe(1);
    saver.dispose();

    // (b) 同步连打 20 次 flush：全部串行且无并发重入
    const burst = new FakeStore(5);
    const burstSaver = createAutoSaver(burst, () => makeState(), 60_000);

    const flushes: Array<Promise<void>> = [];
    for (let i = 0; i < 20; i++) flushes.push(burstSaver.flush());
    await Promise.all(flushes);

    expect(burst.saves).toBe(20);
    expect(burst.maxActive).toBe(1);
    burstSaver.dispose();

    // (c) 慢写入期间再次 trigger 只排队不并发
    const slow = new FakeStore(40);
    const slowSaver = createAutoSaver(slow, () => makeState(), 5);

    slowSaver.trigger();
    await until(() => slow.saves === 1); // 第一次写入已经开始
    slowSaver.trigger();

    await sleep(200); // 等去抖定时器 + 两次写入都跑完
    expect(slow.saves).toBe(2);
    expect(slow.maxActive).toBe(1);
    slowSaver.dispose();

    // (d) 排队中 dispose：已经开始的写入写完，还没开始的被跳过
    const queued = new FakeStore(60);
    const queuedSaver = createAutoSaver(queued, () => makeState(), 5);
    queuedSaver.trigger();
    await until(() => queued.saves === 1); // 第一次写入进行中（约 60ms）
    queuedSaver.trigger(); // 它的定时器 5ms 后到点，排到第一次后面
    await sleep(20);
    queuedSaver.dispose(); // 此时第二次还没开始

    await sleep(200);
    expect(queued.saves, '排队中的那次写入必须被 dispose 掉').toBe(1);
    expect(queued.maxActive).toBe(1);
  });

  it('store.save 抛异常（配额超限）不会中断后续写入，tail 不承载失败', async () => {
    let calls = 0;
    const store: SessionStore = {
      async save(): Promise<void> {
        calls += 1;
        if (calls === 1) throw new Error('QuotaExceededError（模拟）');
      },
      async load(): Promise<SessionState | null> {
        return null;
      },
      async clear(): Promise<void> {},
      async estimateBytes(): Promise<number> {
        return 0;
      },
    };
    const saver = createAutoSaver(store, () => makeState(), 60_000);

    await expect(saver.flush()).resolves.toBeUndefined();
    await expect(saver.flush()).resolves.toBeUndefined();
    expect(calls, '第二次写入必须真的发生').toBe(2);

    saver.dispose();
  });
});
