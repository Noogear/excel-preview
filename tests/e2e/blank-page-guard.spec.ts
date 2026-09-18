/**
 * **"打开文件后页面空白、刷新后又好了"** 的回归守卫。
 *
 * 两条独立的成因，各自钉住：
 *  ① **活动单元与画布绑定错位**：打开文件与"冷标签重建 / 会话恢复"同时跑，
 *     会并行 `createWorkbook` + `attachSheetDeps`，把绑定搞错 → 舞台空白、刷新才恢复。
 *     现在三条路径走同一条串行闸（`unitOpsRef`），并在每次操作后自检 + 重绑
 *     （`ensureActiveUnitRendered`，日志 `render:rebind`）。
 *  ② **渲染期抛异常**：React 18 无错误边界时会卸载整棵树 → 整页空白。
 *     现在有 `AppErrorBoundary` 兜底成一张可读的卡片（含「重新加载」）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { clearLog, fixturePath, importFixture, logKinds, waitForBoot } from './helpers';

const FIXTURE = 'fixture-styles.xlsx';

/** 舞台真的画出来了没有：主导航画布必须有可见尺寸，而且 `#root` 不能是空的 */
async function stageState(page: Page): Promise<{ rootChildren: number; canvasWidth: number; crash: number; status: string | null }> {
  return page.evaluate(() => {
    const canvases = Array.from(document.querySelectorAll<HTMLCanvasElement>('#univer-container canvas'));
    return {
      rootChildren: document.getElementById('root')?.children.length ?? -1,
      canvasWidth: canvases.reduce((max, canvas) => Math.max(max, Math.round(canvas.getBoundingClientRect().width)), 0),
      crash: document.querySelectorAll('[data-testid="app-crash"]').length,
      status: document.querySelector('.status-text')?.textContent?.trim() ?? null,
    };
  });
}

test.describe('打开文件不会白屏（回归）', () => {
  test.skip(!existsSync(fixturePath(FIXTURE)), '请先运行 npm run fixtures');

  test('① 冷标签重建进行中打开新文件：仍然正常渲染，且活动标签是新文件', async ({ page }) => {
    test.setTimeout(300_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);

    // 造 4 个标签（常驻窗口 3）→ 最早的会被冷存；用大文件把"重建"拖慢，保证能撞上
    for (const name of ['bench-large.xlsx', 'fixture-numfmt.xlsx', 'fixture-rules.xlsx', 'fixture-extras.xlsx']) {
      await page.setInputFiles('[data-testid="file-input"]', fixturePath(name));
      await page.waitForTimeout(name === 'bench-large.xlsx' ? 9000 : 2500);
    }

    const cold = await page.evaluate(() =>
      (window as never as { __p0: { getTabRuntime: () => Array<{ id: string; built: boolean }> } }).__p0
        .getTabRuntime()
        .find((entry) => !entry.built),
    );
    test.skip(!cold, '常驻窗口没有触发冷存（夹具太小），跳过');

    // 点冷标签 → 开始重建（解析 6MB 要几秒）；紧接着打开新文件 → 以前这里会白屏
    await page.click(`[data-testid="toolbar-tab-${cold!.id}"]`);
    await page.waitForTimeout(80);
    await page.setInputFiles('[data-testid="file-input"]', fixturePath(FIXTURE));
    await page.waitForTimeout(6000);

    const state = await stageState(page);
    console.log('[regression] 重建中导入后:', JSON.stringify(state));
    expect(state.crash, '不该出现崩溃兜底卡').toBe(0);
    expect(state.canvasWidth, '主画布必须有可见宽度（不能白屏）').toBeGreaterThan(200);
    expect(state.rootChildren, '#root 不该是空的').toBeGreaterThan(0);

    const active = await page.evaluate(() =>
      (window as never as { __p0: { getActiveTabId: () => string | null; getTabs: () => Array<{ id: string; fileName: string }> } })
        .__p0.getTabs()
        .find((tab) => tab.id === (window as never as { __p0: { getActiveTabId: () => string | null } }).__p0.getActiveTabId())
        ?.fileName ?? null,
    );
    expect(active, '最后打开的应该是新文件').toBe(FIXTURE);

    // 冷标签那条路径也不能把绑定弄丢：切回去照样有画布
    await page.click(`[data-testid="toolbar-tab-${cold!.id}"]`);
    await page.waitForTimeout(5000);
    const back = await stageState(page);
    expect(back.canvasWidth, '切回冷标签后画布也要在').toBeGreaterThan(200);
  });

  test('② 会话恢复进行中打开新文件：仍然正常渲染', async ({ page }) => {
    test.setTimeout(300_000);
    await page.setViewportSize({ width: 1440, height: 900 });

    // 先存一个"大文件"的会话，让恢复要几秒
    await waitForBoot(page);
    await page.setInputFiles('[data-testid="file-input"]', fixturePath('bench-large.xlsx'));
    await page.waitForTimeout(9000);

    await page.reload();
    await page.waitForFunction(() => typeof (window as never as { __p0?: unknown }).__p0 !== 'undefined', null, { timeout: 60_000 });
    await page.setInputFiles('[data-testid="file-input"]', fixturePath('fixture-numfmt.xlsx'));
    await page.waitForTimeout(9000);

    const state = await stageState(page);
    console.log('[regression] 恢复中导入后:', JSON.stringify(state));
    expect(state.crash).toBe(0);
    expect(state.canvasWidth, '恢复期间打开文件也不能白屏').toBeGreaterThan(200);
    expect(state.rootChildren).toBeGreaterThan(0);
  });

  /**
   * ③ 兜底卡本身（`AppErrorBoundary` + `ErrorFallback`）由单测覆盖
   *（`tests/unit/app-crash.test.ts`：SSR 渲染兜底界面 + 错误状态推导 + 无错时渲染子节点）。
   * e2e 这边只确认**正常路径不会被兜底卡误伤**。
   */
  test('③ 正常打开文件时不会出现兜底卡', async ({ page }) => {
    test.setTimeout(120_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    const state = await stageState(page);
    expect(state.crash, '正常路径不该出现崩溃兜底卡').toBe(0);
    expect(state.canvasWidth).toBeGreaterThan(200);
    expect(await logKinds(page), '不该记录渲染错误').not.toContain('app:render-error');
  });

  /**
   * ④ **自愈动作本身要有用**（不能只是"写了段防御代码"）。
   *
   * 用户报的白屏没法稳定重现，所以这里**人工把故障状态造出来**：用测试钩子
   * `createHandsOnSheet()` 绕过标签体系直接 `createWorkbook`，于是
   * 「Univer 当前单元」≠「应用的活动标签」—— 这正是白屏那一类故障的核心特征
   * （应用以为在看 A，画布画的是 B，或者什么都没画）。
   * 然后调 `healActiveUnit(活动标签)`，断言：绑定真的接回来了、画布还在、并且记了 `render:rebind`。
   */
  test('④ 绑定错位后自愈动作能把画布接回来（人工造故障）', async ({ page }) => {
    test.setTimeout(120_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);

    const tabId = await page.evaluate(
      () => (window as never as { __p0: { getActiveTabId: () => string | null } }).__p0.getActiveTabId(),
    );
    expect(tabId, '打开文件后应当有活动标签').toBeTruthy();
    const heal = (id: string): Promise<boolean> =>
      page.evaluate(
        (target) => (window as never as { __p0: { healActiveUnit: (id: string) => Promise<boolean> } }).__p0.healActiveUnit(target),
        id,
      );

    /**
     * 前半段：**健康状态不许误报**。
     *
     * 自检刚写出来时就是"一测到画布还没画完就判坏了"，于是每次正常导入都重绑一次、还留下
     * `render:rebind` 噪声。现在先等首帧再判，正常路径应当**一次日志都不打**。
     */
    await clearLog(page);
    expect(await heal(tabId!), '健康状态下自愈应当直接通过').toBe(true);
    expect(await logKinds(page), '健康状态不该触发重绑（不能误报）').not.toContain('render:rebind');

    // 后半段：造故障 —— 直接建一个不属于任何标签的工作簿 → 画布绑定的单元不再是活动标签
    const desync = await page.evaluate(() => {
      const p0 = (window as never as {
        __p0: {
          createHandsOnSheet: () => string | null;
          getActiveWorkbookId: () => string | null;
          getActiveTabId: () => string | null;
        };
      }).__p0;
      const id = p0.createHandsOnSheet();
      return { handsOnId: id, unit: p0.getActiveWorkbookId(), tab: p0.getActiveTabId() };
    });
    console.log('[regression] 人工造出的绑定错位:', JSON.stringify(desync));
    test.skip(!desync.handsOnId || desync.unit !== desync.handsOnId, '没能造出「当前单元 ≠ 活动标签」的状态，跳过');
    expect(desync.unit, '此刻画布绑定的确实是那个非标签单元').not.toBe(desync.tab);

    // 自愈：应当发现错位 → 重绑 → 复验通过
    await clearLog(page);
    expect(await heal(tabId!), '自愈动作应当报告"已经接回来了"').toBe(true);

    const now = await page.evaluate(() => {
      const p0 = (window as never as {
        __p0: { getActiveWorkbookId: () => string | null; getActiveTabId: () => string | null };
      }).__p0;
      return { unit: p0.getActiveWorkbookId(), tab: p0.getActiveTabId() };
    });
    expect(now.unit, '自愈后画布绑定的必须就是活动标签').toBe(now.tab);
    expect(now.unit).toBe(tabId);

    const state = await stageState(page);
    expect(state.canvasWidth, '自愈后画布必须有可见宽度').toBeGreaterThan(200);
    expect(state.crash).toBe(0);

    const kinds = await logKinds(page);
    expect(kinds, '应当记录"重绑"').toContain('render:rebind');
    expect(kinds, '不该出现重绑失败').not.toContain('render:rebind-failed');
  });
});
