/**
 * 多标签：常驻窗口 + 冷存/切回重建（用户问："多个表格能不能同时打开？内存能不能省？"）。
 *
 * 机制：**打开的标签不设上限**（工具栏标签条），但**同时实体化的 Univer 工作簿只保留最近用过的
 * 3 个**（`DEFAULT_RESIDENT_TAB_LIMIT`）；超出的按最久未用冷存（dispose 工作簿），
 * 切回时按需重建（重解析字节 → 建簿 → 应用特性 → 回放编辑）。
 *
 * 这里断言的都是"用户能感知的后果"，不是内部实现细节：
 *  ① 开 4 个标签后，实体化的只有 3 个（内存从 O(N) 变 O(K)）；
 *  ② 冷标签切回去内容仍在（含**改过的单元格**，证明编辑回放正确）；
 *  ③ 冷标签切回后，表格本身可用（活动表名/单元格显示值正确）；
 *  ④ 重建后撤销栈与历史被切断（不会把旧 mutation 打到新模型上）——用日志与计数断言。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, waitForBoot } from './helpers';

const FIXTURES = ['fixture-styles.xlsx', 'fixture-numfmt.xlsx', 'fixture-table.xlsx', 'fixture-rules.xlsx'];

interface TabRuntime {
  id: string;
  built: boolean;
  lastUsedAt: number;
}

const runtime = (page: Page): Promise<TabRuntime[]> =>
  page.evaluate(() => (window as never as { __p0: { getTabRuntime: () => TabRuntime[] } }).__p0.getTabRuntime());

const tabs = (page: Page): Promise<Array<{ id: string; fileName: string }>> =>
  page.evaluate(() => (window as never as { __p0: { getTabs: () => Array<{ id: string; fileName: string }> } }).__p0.getTabs());

/** 切到某个标签（走工具栏真实点击，不是内部函数）；标签 id 从 hooks 里反查，避免猜 testid */
async function clickTab(page: Page, fileName: string): Promise<void> {
  const id = await page.evaluate(
    (name) =>
      (window as never as { __p0: { getTabs: () => Array<{ id: string; fileName: string }> } }).__p0
        .getTabs()
        .find((tab) => tab.fileName === name)?.id ?? null,
    fileName,
  );
  if (!id) throw new Error(`找不到标签：${fileName}`);
  await page.locator(`[data-testid="toolbar-tab-${id}"]`).click();
  await page.waitForFunction(
    (target) => {
      const hooks = (window as never as { __p0: { getActiveTabId: () => string | null } }).__p0;
      return hooks.getActiveTabId() === target;
    },
    id,
    { timeout: 60_000 },
  );
  await page.waitForTimeout(300);
}

test.describe('多标签：内存封顶与冷存重建', () => {
  test.skip(!FIXTURES.every((name) => existsSync(fixturePath(name))), '请先运行 npm run fixtures 生成样本');

  test('① 开 4 个标签：实体化的不超过常驻窗口，冷标签切回后内容与编辑都在', async ({ page }) => {
    test.setTimeout(180_000);
    await waitForBoot(page);

    for (const name of FIXTURES) await importFixture(page, name);

    const all = await tabs(page);
    expect(all, '4 个文件应开成 4 个标签（标签数不设上限）').toHaveLength(4);

    const limit = await page.evaluate(
      () => (window as never as { __p0: { residentTabLimit: number } }).__p0.residentTabLimit,
    );
    const afterImport = await runtime(page);
    const built = afterImport.filter((entry) => entry.built);
    expect(built.length, `实体化标签数应被常驻窗口（${limit}）封顶`).toBeLessThanOrEqual(limit);
    expect(built.length, '至少当前标签是实体化的').toBeGreaterThan(0);
    expect(afterImport.filter((entry) => !entry.built).length, '超出的标签应被冷存').toBeGreaterThan(0);

    // 在最久的那个（第 1 个）标签上改一格，再切走 → 它会经历"冷存 → 重建 → 回放编辑"
    const first = all[0];
    await clickTab(page, first.fileName);
    const edited = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getActiveSheet: () => { getRange: (a1: string) => { getRange: () => { startRow: number; startColumn: number } } };
          setValues: (row: number, col: number, values: (string | number)[][]) => void;
          getDisplayValue: (row: number, col: number) => string | null;
        };
      }).__p0;
      const rect = hooks.getActiveSheet().getRange('A5').getRange();
      hooks.setValues(rect.startRow, rect.startColumn, [['冷存标记']]);
      return { row: rect.startRow, col: rect.startColumn, value: hooks.getDisplayValue(rect.startRow, rect.startColumn) };
    });
    expect(edited.value, '先确认这一格真的写进去了').toBe('冷存标记');

    // 切到别的标签，把第 1 个挤出常驻窗口
    for (const name of [FIXTURES[1], FIXTURES[2], FIXTURES[3]]) await clickTab(page, name);
    const runtimeAfter = await runtime(page);
    const firstRuntime = runtimeAfter.find((entry) => entry.id === first.id);
    expect(firstRuntime?.built, '第 1 个标签这时应已被冷存').toBe(false);

    // 切回去：应重建，并且**编辑被回放**
    await clickTab(page, first.fileName);
    await page.waitForFunction(
      (id) => {
        const hooks = (window as never as { __p0: { getTabRuntime: () => TabRuntime[] } }).__p0.getTabRuntime();
        return hooks.find((entry) => entry.id === id)?.built === true;
      },
      first.id,
      { timeout: 60_000 },
    );
    const rebuilt = await page.evaluate(
      ({ row, col }) =>
        (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0.getDisplayValue(row, col),
      { row: edited.row, col: edited.col },
    );
    expect(rebuilt, '切回后改过的格子内容应还在（编辑回放成功）').toBe('冷存标记');

    const afterRebuild = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { log: Array<{ kind: string }>; getUndoRedoCounts: () => { undos: number; redos: number } };
      }).__p0;
      const kinds = hooks.log.map((entry) => entry.kind);
      return {
        builds: kinds.filter((kind) => kind === 'tab:build-done').length,
        coldStores: kinds.filter((kind) => kind === 'tab:cold-store').length,
        replays: kinds.filter((kind) => kind === 'tab:edits-replayed').length,
        historyResets: kinds.filter((kind) => kind === 'tab:history-reset').length,
        undoCounts: hooks.getUndoRedoCounts(),
      };
    });
    expect(afterRebuild.coldStores, '应确实发生过冷存').toBeGreaterThan(0);
    expect(afterRebuild.builds, '应确实发生过重建').toBeGreaterThan(0);
    expect(afterRebuild.replays, '重建后应回放过编辑').toBeGreaterThan(0);
    expect(afterRebuild.historyResets, '重建后应清撤销栈/历史（否则 Ctrl+Z 会写坏新模型）').toBeGreaterThan(0);
    expect(afterRebuild.undoCounts.undos, '重建后撤销栈应为空').toBe(0);

    // 常驻窗口在多次切换后依然成立
    const finalRuntime = await runtime(page);
    expect(finalRuntime.filter((entry) => entry.built).length).toBeLessThanOrEqual(limit);
  });

  /**
   * ② 堆内存实测（打印，不做脆断言）。
   *
   * 口径与 `perf.spec.ts` / P3 文档一致：`performance.memory.usedJSHeapSize`（不含 Uint8Array 外部内存）。
   * 用中等夹具（1.2 MB）开 4 个标签：第 4 个进来时第 1 个被冷存，
   * 如果常驻窗口生效，堆增长应明显小于"每次导入都常驻"的线性增长。
   */
  test('② 开 4 个标签的堆增长（常驻窗口把内存从 O(N) 拉成 O(K)）', async ({ page }) => {
    test.setTimeout(180_000);
    const medium = 'bench-medium.xlsx';
    test.skip(!existsSync(fixturePath(medium)), '请先运行 npm run bench:fixture 生成中等夹具');

    await waitForBoot(page);
    const samples: Array<{ step: string; heapMB: number; built: number }> = [];
    const readHeap = async (step: string) => {
      const sample = await page.evaluate(() => {
        const hooks = (window as never as {
          __p0: { getTabRuntime: () => TabRuntime[] };
          performance: { memory?: { usedJSHeapSize: number } };
        }).__p0;
        const memory = (performance as unknown as { memory?: { usedJSHeapSize: number } }).memory;
        return {
          heapMB: memory ? Math.round(memory.usedJSHeapSize / 1048576) : -1,
          built: hooks.getTabRuntime().filter((entry) => entry.built).length,
        };
      });
      samples.push({ step, ...sample });
    };

    await readHeap('启动');
    for (let i = 1; i <= 4; i += 1) {
      await importFixture(page, medium);
      await readHeap(`导入第 ${i} 个`);
    }

    // eslint-disable-next-line no-console
    console.log('[tabs-memory] 堆采样（MB）:', JSON.stringify(samples));
    expect(samples.every((sample) => sample.heapMB > 0), '应能读到 performance.memory（Chrome）').toBe(true);
    const limit = await page.evaluate(
      () => (window as never as { __p0: { residentTabLimit: number } }).__p0.residentTabLimit,
    );
    expect(samples[samples.length - 1].built, '第 4 个标签之后实体化的仍应被窗口封顶').toBeLessThanOrEqual(limit);
  });

  /** ③ 脏标签关闭前必须确认（否则编辑静默丢失） */
  test('③ 改过的标签在关闭前会先确认', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await importFixture(page, FIXTURES[0]);

    // 改一格 → 标签变脏
    await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getActiveSheet: () => { getRange: (a1: string) => { getRange: () => { startRow: number; startColumn: number } } };
          setValues: (row: number, col: number, values: (string | number)[][]) => void;
        };
      }).__p0;
      const rect = hooks.getActiveSheet().getRange('B7').getRange();
      hooks.setValues(rect.startRow, rect.startColumn, [['脏了']]);
    });
    await page.waitForTimeout(300);

    const tabId = (await tabs(page))[0].id;
    const closeButton = page.locator(`[data-testid="toolbar-tab-close-${tabId}"]`);

    // 第一次：取消 → 标签还在
    page.once('dialog', (dialog) => void dialog.dismiss());
    await closeButton.click();
    await page.waitForTimeout(300);
    expect(await tabs(page), '取消后标签应保留').toHaveLength(1);

    // 第二次：确认 → 标签关闭
    page.once('dialog', (dialog) => void dialog.accept());
    await closeButton.click();
    await page.waitForTimeout(500);
    expect(await tabs(page), '确认后标签应关闭').toHaveLength(0);

    const closed = await page.evaluate(() =>
      (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
    );
    expect(closed).toContain('tab:close-cancelled');
    expect(closed).toContain('tab:close');
  });
});
