/**
 * 会话持久化（合并自原 persistence.spec.ts 的 5 条 + p3-persistence.spec.ts 的 1 条 → 3 条）。
 *
 *   ① App 层全链路：导入 → 编辑 → 切换三种交互模式 → 放工作区条目 → 重开页面，
 *      标签/编辑/工作区/模式**都恢复**（三种模式值逐一验证，而不只验一种）
 *   ② 存储层真实 IndexedDB 往返：Uint8Array 逐字节一致、多标签/编辑/工作区原样带回来、clear() 归零
 *   ③ 坏数据（version 不符 / 结构损坏 / 非对象）→ load() 返回 null、物理清掉、绝不抛
 *
 * 已删除（属于实现细节或被 tests/unit 覆盖）：
 *   - "模块可由 dev server 直接加载"（只是 typeof 断言）
 *   - createAutoSaver 去抖/flush/串行（tests/unit/persistence.test.ts 用假定时器覆盖）
 *   - "两个库名互不干扰"（同一诉求已由 ② 的往返 + clear 覆盖）
 *
 * 注：① 是"同一 page 内 reload"（同一个 BrowserContext 才共享 IndexedDB）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { clearSession, fixturePath, importFixture, waitForBoot } from './helpers';

/** dev server 直接提供的模块路径（Vite 会即时编译 .ts） */
const MODULE_URL = '/src/persistence/session.ts';

/** 页面里动态 import 到的模块形状：与源码逐字对齐，但**只存在于编译期**（evaluate 回调会被序列化） */
type PersistenceModule = typeof import('../../src/persistence/session');

/** 每次用不同库名，避免同一浏览器 profile 内的残留互相影响 */
function uniqueDbName(label: string): string {
  return `e2e-session-${label}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** 打开页面（存储层不依赖 App 启动，只要 dev server 在） */
async function openPage(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForLoadState('domcontentloaded');
}

interface RoundTripArgs {
  specifier: string;
  dbName: string;
}

test.describe('会话持久化', () => {
  test('① 误关闭后重开：标签/编辑/工作区恢复，且三种交互模式值都能恢复', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures');

    await waitForBoot(page);
    await clearSession(page); // 清掉可能残留的会话，保证用例可重复

    await importFixture(page, 'fixture-styles.xlsx');

    // 编辑内容（A3 原本是空的）
    await page.evaluate(() => {
      (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0.setValue(
        2,
        0,
        '恢复测试-内容',
      );
    });
    // 放一个工作区条目（走「+」对话框：取当前选区 → 确认）
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:F2');
    });
    await page.waitForTimeout(200);
    await page.click('[data-testid="workspace-add"]');
    await page.click('[data-testid="import-use-selection"]');
    await page.click('[data-testid="import-confirm"]');
    await page.waitForTimeout(400);

    // 显式落盘（真实场景靠 2s 定时保存 + pagehide）
    await page.evaluate(async () => {
      await (window as never as { __p0: { flushSession: () => Promise<void> } }).__p0.flushSession();
    });

    const readState = () =>
      page.evaluate(() => {
        const hooks = (window as never as {
          __p0: {
            getTabs: () => Array<{ id: string; fileName: string }>;
            getMode: () => string;
            getWorkspaceItems: () => Array<{ label: string }>;
            getDisplayValue: (r: number, c: number) => string | null;
          };
        }).__p0;
        return {
          tabs: hooks.getTabs(),
          mode: hooks.getMode(),
          workspace: hooks.getWorkspaceItems(),
          value: hooks.getDisplayValue(2, 0),
        };
      });

    const before = await readState();
    expect(before.value).toBe('恢复测试-内容');
    expect(before.workspace.length, '工作区条目应已建立').toBeGreaterThan(0);

    // ---- 三种模式值逐一"切模式 → 落盘 → 重开页面 → 校验" ----
    for (const mode of ['select', 'drag', 'click-swap'] as const) {
      await page.click(`[data-testid="toolbar-mode-${mode}"]`);
      await page.waitForTimeout(250);
      await page.evaluate(async () => {
        await (window as never as { __p0: { flushSession: () => Promise<void> } }).__p0.flushSession();
      });

      // 模拟"误关闭后重新打开"
      await waitForBoot(page);
      await page.waitForFunction(
        () => (window as never as { __p0: { getTabs: () => unknown[] } }).__p0.getTabs().length > 0,
        null,
        { timeout: 30_000 },
      );

      expect(
        await page.evaluate(() => (window as never as { __p0: { getMode: () => string } }).__p0.getMode()),
        `重开页面后交互模式应恢复为 ${mode}`,
      ).toBe(mode);
      await expect(
        page.locator(`[data-testid="toolbar-mode-${mode}"]`),
        `工具条上 ${mode} 按钮应处于选中态`,
      ).toHaveAttribute('aria-checked', 'true');
    }

    // 标签 / 编辑 / 工作区在最后一轮重开之后仍然在
    const after = await readState();
    expect(after.tabs.map((tab) => tab.fileName), '标签页必须恢复').toEqual(before.tabs.map((tab) => tab.fileName));
    expect(after.value, '编辑过的内容必须恢复').toBe('恢复测试-内容');
    expect(after.workspace.length, '工作区条目必须恢复').toBe(before.workspace.length);

    await clearSession(page); // 清理，避免影响后续运行
  });

  test('② 真实 IndexedDB 往返：字节逐字节一致，多标签/编辑/工作区都在，clear() 归零', async ({ page }) => {
    const dbName = uniqueDbName('roundtrip');
    await openPage(page);

    const saved = await page.evaluate(
      async ({ specifier, dbName: name }: RoundTripArgs) => {
        const mod = (await import(/* @vite-ignore */ specifier)) as PersistenceModule;
        const store = mod.createSessionStore(name);

        const big = new Uint8Array(4096);
        for (let i = 0; i < big.length; i++) big[i] = (i * 7) % 251;

        await store.save({
          version: 1,
          savedAt: 1_700_000_000_000,
          activeTabId: 'tab-2',
          mode: 'click-swap',
          settings: {
            keepSourceOnDrop: false,
            removeItemAfterPaste: true,
            // 工作区"每行放几格"由「格子最小宽度」决定（面板宽度 ÷ 这个值）
            tileMinWidth: 160,
            sidebarWidth: 320,
          },
          tabs: [
            {
              id: 'tab-1',
              fileName: '第一个.xlsx',
              originalBytes: big.slice(0, 16),
              snapshot: null,
              edits: [],
            },
            {
              id: 'tab-2',
              fileName: '成绩表.xlsx',
              originalBytes: big,
              snapshot: { id: 'workbook-1', sheetOrder: ['sheet-1'] },
              edits: [
                { sheetId: 'sheet-1', row: 9, col: 3, value: 42, formula: null },
                { sheetId: 'sheet-1', row: 10, col: 0, value: null, formula: 'SUM(A1:A3)' },
                { sheetId: 'sheet-1', row: 11, col: 1, value: true, formula: null },
              ],
            },
          ],
          workspace: [
            {
              id: 'ws-1',
              source: { sheetId: 'sheet-1', sheetName: 'Sheet1', a1: 'A1:B2', startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
              rows: 2,
              cols: 2,
              cells: [[{ text: '1', value: 1 }], [{ text: '2', value: 2 }]],
              values: [[1]],
              formulas: [[null]],
              createdAt: 1_700_000_000_000,
              label: 'A1:B2',
            },
          ],
        });

        return { estimate: await store.estimateBytes() };
      },
      { specifier: MODULE_URL, dbName },
    );

    expect(saved.estimate, '含 4KB xlsx 字节的状态必须报出 >4KB 的占用').toBeGreaterThan(4096);

    // ---- 刷新页面：全新的 JS 环境，只有 IndexedDB 里的数据活下来 ----
    await page.reload();
    await page.waitForLoadState('domcontentloaded');

    const restored = await page.evaluate(
      async ({ specifier, dbName: name }: RoundTripArgs) => {
        const mod = (await import(/* @vite-ignore */ specifier)) as PersistenceModule;
        const store = mod.createSessionStore(name);
        const state = await store.load();
        if (!state) return null;

        const tab = state.tabs[1];
        const bytes = tab?.originalBytes;
        let mismatches = 0;
        if (bytes) {
          for (let i = 0; i < bytes.length; i++) if (bytes[i] !== (i * 7) % 251) mismatches += 1;
        }

        return {
          version: state.version,
          mode: state.mode,
          activeTabId: state.activeTabId,
          tabCount: state.tabs.length,
          firstTabName: state.tabs[0]?.fileName ?? null,
          secondTabName: tab?.fileName ?? null,
          snapshotId: (tab?.snapshot as { id?: string } | null | undefined)?.id ?? null,
          isUint8Array: bytes instanceof Uint8Array,
          byteLength: bytes?.byteLength ?? -1,
          mismatches,
          edits: tab?.edits ?? [],
          workspaceCount: state.workspace.length,
          workspaceFirstId: state.workspace[0]?.id ?? null,
          estimate: await store.estimateBytes(),
        };
      },
      { specifier: MODULE_URL, dbName },
    );

    expect(restored, '刷新后必须能 load 出会话').not.toBeNull();
    if (!restored) return;

    expect(restored.version).toBe(1);
    expect(restored.mode, '交互模式必须恢复').toBe('click-swap');
    expect(restored.activeTabId, '激活标签必须恢复').toBe('tab-2');
    expect(restored.tabCount).toBe(2);
    expect(restored.firstTabName).toBe('第一个.xlsx');
    expect(restored.secondTabName).toBe('成绩表.xlsx');
    expect(restored.snapshotId, 'Univer 快照必须原样带回来').toBe('workbook-1');

    // 原始字节：结构化克隆往返，必须逐字节一致（证明没有转 base64 / 没有截断）
    expect(restored.isUint8Array, 'originalBytes 必须还是 Uint8Array').toBe(true);
    expect(restored.byteLength).toBe(4096);
    expect(restored.mismatches, '4096 个字节必须逐一相同').toBe(0);

    // 编辑（值 / 公式 / 布尔都要对）
    expect(restored.edits).toHaveLength(3);
    expect(restored.edits[0]).toEqual({ sheetId: 'sheet-1', row: 9, col: 3, value: 42, formula: null });
    expect(restored.edits[1]).toEqual({ sheetId: 'sheet-1', row: 10, col: 0, value: null, formula: 'SUM(A1:A3)' });
    expect(restored.edits[2]).toEqual({ sheetId: 'sheet-1', row: 11, col: 1, value: true, formula: null });

    // 工作区
    expect(restored.workspaceCount).toBe(1);
    expect(restored.workspaceFirstId).toBe('ws-1');
    expect(restored.estimate).toBeGreaterThan(4096);

    // ---- clear()：清掉之后回到"全新开始" ----
    const afterClear = await page.evaluate(async ({ specifier, dbName: name }: RoundTripArgs) => {
      const mod = (await import(/* @vite-ignore */ specifier)) as PersistenceModule;
      const store = mod.createSessionStore(name);
      await store.clear();
      return { loaded: await store.load(), bytes: await store.estimateBytes() };
    }, { specifier: MODULE_URL, dbName });

    expect(afterClear.loaded).toBeNull();
    expect(afterClear.bytes).toBe(0);
  });

  test('③ 坏数据（version 不符 / 结构损坏 / 非对象）→ load() 返回 null 且物理清掉', async ({ page }) => {
    await openPage(page);

    const cases = [
      { label: 'version=99', record: { version: 99, savedAt: 1, activeTabId: null, mode: 'drag', tabs: [], workspace: [] } },
      { label: 'tabs 不是数组', record: { version: 1, savedAt: 1, activeTabId: null, mode: 'drag', tabs: 'nope', workspace: [] } },
      { label: '缺 mode', record: { version: 1, savedAt: 1, activeTabId: null, tabs: [], workspace: [] } },
      { label: '非对象', record: 'not-a-session' },
      { label: 'null', record: null },
    ];

    for (const item of cases) {
      const dbName = uniqueDbName('corrupt');
      const result = await page.evaluate(
        async ({ specifier, dbName: name, record }) => {
          const mod = (await import(/* @vite-ignore */ specifier)) as PersistenceModule;

          // 绕过 store，直接往 IndexedDB 里塞坏数据
          const rawPut = await new Promise<boolean>((resolve) => {
            const openReq = indexedDB.open(name, 1);
            openReq.onupgradeneeded = () => {
              if (!openReq.result.objectStoreNames.contains('state')) openReq.result.createObjectStore('state');
            };
            openReq.onerror = () => resolve(false);
            openReq.onsuccess = () => {
              const db = openReq.result;
              const tx = db.transaction('state', 'readwrite');
              tx.objectStore('state').put(record, 'session');
              tx.oncomplete = () => {
                db.close();
                resolve(true);
              };
              tx.onerror = () => {
                db.close();
                resolve(false);
              };
            };
          });

          const store = mod.createSessionStore(name);
          const first = await store.load();
          const bytes = await store.estimateBytes();
          const second = await store.load();

          // 直接读原始 key，确认坏数据被清掉了（而不是每次启动都踩同一颗雷）
          const raw = await new Promise<unknown>((resolve) => {
            const openReq = indexedDB.open(name, 1);
            openReq.onerror = () => resolve('open-failed');
            openReq.onsuccess = () => {
              const db = openReq.result;
              const tx = db.transaction('state', 'readonly');
              const getReq = tx.objectStore('state').get('session');
              getReq.onsuccess = () => {
                const value: unknown = getReq.result;
                db.close();
                resolve(value);
              };
              getReq.onerror = () => {
                db.close();
                resolve('get-failed');
              };
            };
          });

          return { rawPut, firstIsNull: first === null, secondIsNull: second === null, bytes, rawGone: raw === undefined };
        },
        { specifier: MODULE_URL, dbName, record: item.record },
      );

      expect(result.rawPut, `[${item.label}] 坏数据要能塞进去`).toBe(true);
      expect(result.firstIsNull, `[${item.label}] load() 必须返回 null`).toBe(true);
      expect(result.secondIsNull, `[${item.label}] 再 load() 仍是 null`).toBe(true);
      expect(result.bytes, `[${item.label}] estimateBytes() 归零`).toBe(0);
      expect(result.rawGone, `[${item.label}] 坏数据必须被清理`).toBe(true);
    }
  });
});
