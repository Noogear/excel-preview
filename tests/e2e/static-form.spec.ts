/**
 * **静态形态**（交付到 GitHub / Gitee Pages 的那一份）的守卫用例。
 *
 * 覆盖三件事（用户要求"实现静态、本地双形态，对加载速度进行优化"）：
 *  ① 静态产物真的能用：挂在**子路径**下也能引导、打开表格、导出下载（没有任何后端接口）；
 *  ② 启动速度：骨架先画出来 → 应用挂载 → 表格可用，三段的耗时都在预算内（并打印实测值）；
 *  ③ 静态形态**不该**出现桥相关入口（本地版才有的 Excel 转换能力）。
 *
 * 依赖 `dist/`：没构建就整组跳过（`npm run build:static`）。可用环境变量换产物目录，
 * 便于做"优化前/优化后"对比：`$env:STATIC_DIST='dist-baseline'`。
 */
import { expect, test } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { startStaticServer } from '../../tools/serve-dist.mjs';

const DIST = resolve(process.env.STATIC_DIST ?? 'dist');
const HAS_DIST = existsSync(join(DIST, 'index.html'));
const FIXTURE = join(process.cwd(), 'fixtures', 'fixture-styles.xlsx');

/** 等元素位置稳定（连续两次测得同一个矩形），避免点到"正在被定位"的菜单项上 */
async function stableBox(page: import('@playwright/test').Page, locator: import('@playwright/test').Locator): Promise<void> {
  let previous = '';
  for (let attempt = 0; attempt < 12; attempt += 1) {
    const box = await locator.boundingBox();
    const key = box ? `${Math.round(box.x)},${Math.round(box.y)}` : '';
    if (key !== '' && key === previous) return;
    previous = key;
    await page.waitForTimeout(60);
  }
}

/** 启动预算（本机实测值留了 3–5 倍余量，避免慢机器假红；真实数字会打印出来） */
const BUDGET = {
  /** 首屏骨架可见（只对含骨架的产物断言） */
  skeletonMs: Number(process.env.BUDGET_SKELETON ?? 1500),
  firstContentfulPaintMs: Number(process.env.BUDGET_FCP ?? 3000),
  appMountedMs: Number(process.env.BUDGET_MOUNTED ?? 8000),
  univerBootedMs: Number(process.env.BUDGET_BOOTED ?? 25_000),
};

test.describe('静态形态：能托管、启动快、桥能力被降级', () => {
  test.skip(!HAS_DIST, `没有找到 ${DIST}/index.html，请先 npm run build:static`);

  let server: Awaited<ReturnType<typeof startStaticServer>>;

  test.beforeAll(async () => {
    server = await startStaticServer({ distDir: DIST, prefix: '/repo/', port: 0 });
  });
  test.afterAll(async () => {
    await server?.close();
  });

  test('① 子路径托管可用：骨架 → 应用 → 打开表格 → 导出下载', async ({ page }) => {
    test.setTimeout(180_000);
    const failed: string[] = [];
    const bad: string[] = [];
    page.on('requestfailed', (request) => failed.push(`${request.url()} :: ${request.failure()?.errorText ?? ''}`));
    page.on('response', (response) => {
      if (response.status() >= 400) bad.push(`${response.status()} ${response.url()}`);
    });

    await page.goto(server.url);

    // 首屏骨架写在 HTML 里（保证"第一帧不是白屏"）：直接读源码断言，避免本地快得看不到
    const html = readFileSync(join(DIST, 'index.html'), 'utf8');
    expect(html, '静态产物的 HTML 里应内联首屏骨架').toContain('data-app-skeleton');

    await page.waitForFunction(() => typeof (window as never as { __p0?: unknown }).__p0 !== 'undefined', null, {
      timeout: 60_000,
    });
    await expect(page.locator('[data-app-skeleton]'), '应用挂载后骨架应被替换').toHaveCount(0);

    // 打开表格（真实文件输入）→ 渲染出内容
    await page.setInputFiles('input[type=file]', FIXTURE);
    await expect
      .poll(
        () =>
          page.evaluate(() =>
            (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('A2'),
          ),
        { timeout: 60_000 },
      )
      .toBe('区域');

    // 样式确实生效（懒加载进来的 App CSS 必须被应用，否则界面是裸的）
    const toolbarStyled = await page.locator('[data-testid="toolbar-undo"]').evaluate((node) => {
      const style = getComputedStyle(node);
      return style.borderRadius !== '' && style.cursor !== '';
    });
    expect(toolbarStyled, '懒加载的应用 CSS 应已生效').toBe(true);

    // 导出：纯前端下载（「导出」现在先出菜单，默认项仍是保真导出 .xlsx）
    //
    // 稳定性：菜单是"先渲染、再按视口测量位置"的（可能微调一次），CPU 忙时单击会落在移动中的元素上 ——
    // 所以先等它位置稳定，再点；万一没触发下载，重试一次（这是驱动方式的抖动，不是应用行为）。
    await page.click('[data-testid="toolbar-export"]');
    const keepItem = page.locator('[data-testid="context-menu-export-xlsx-keep"]');
    await expect(keepItem).toBeVisible();
    await stableBox(page, keepItem);
    let download;
    try {
      [download] = await Promise.all([page.waitForEvent('download', { timeout: 20_000 }), keepItem.click()]);
    } catch {
      if (!(await keepItem.isVisible().catch(() => false))) await page.click('[data-testid="toolbar-export"]');
      await stableBox(page, keepItem);
      [download] = await Promise.all([page.waitForEvent('download', { timeout: 60_000 }), keepItem.click()]);
    }
    const path = await download.path();
    expect(path ? readFileSync(path).length : 0, '导出的 xlsx 应有内容').toBeGreaterThan(1000);

    expect(failed, '静态托管下不应有资源请求失败').toEqual([]);
    expect(bad, '不应有 404/5xx').toEqual([]);
  });

  test('② 启动预算：骨架 → 应用挂载 → 表格可用（打印实测值）', async ({ page }) => {
    test.setTimeout(180_000);
    const html = readFileSync(join(DIST, 'index.html'), 'utf8');
    const hasSkeleton = html.includes('data-app-skeleton');

    // 用 `commit` 而不是 `load`：从"服务器开始回 HTML"计时，才量得到用户实际等多长时间才看到东西
    const navigationStart = Date.now();
    await page.goto(server.url, { waitUntil: 'commit' });

    let skeletonVisibleMs: number | null = null;
    if (hasSkeleton) {
      await expect(page.locator('[data-app-skeleton]'), '首屏骨架应当立刻可见').toBeVisible({ timeout: 15_000 });
      skeletonVisibleMs = Date.now() - navigationStart;
    }

    await page.waitForFunction(() => typeof (window as never as { __p0?: unknown }).__p0 !== 'undefined', null, {
      timeout: 60_000,
    });
    await page.waitForFunction(() => performance.getEntriesByName('univer:booted').length > 0, null, { timeout: 60_000 });
    await page.waitForTimeout(300); // 等 FCP 上报（有些环境会晚一拍）

    const marks = await page.evaluate(() => {
      const read = (name: string): number | null => {
        const entry = performance.getEntriesByName(name)[0] as PerformanceEntry | undefined;
        return entry ? Math.round(entry.startTime) : null;
      };
      const nav = performance.getEntriesByType('navigation')[0] as PerformanceNavigationTiming | undefined;
      const transferred = performance
        .getEntriesByType('resource')
        .reduce((sum, entry) => sum + ((entry as PerformanceResourceTiming).transferSize || 0), 0);
      return {
        fcp: read('first-contentful-paint'),
        bundleLoaded: read('app:bundle-loaded'),
        chunkLoaded: read('app:chunk-loaded'),
        appMounted: read('app:mounted'),
        univerBooted: read('univer:booted'),
        domContentLoaded: nav ? Math.round(nav.domContentLoadedEventEnd) : null,
        transferredKB: Math.round(transferred / 1024),
      };
    });

    console.log(
      `[startup] 骨架可见 ${skeletonVisibleMs ?? '—'}ms · FCP ${marks.fcp ?? '—'}ms · 入口 ${marks.bundleLoaded}ms · ` +
        `应用分块 ${marks.chunkLoaded}ms · 应用挂载 ${marks.appMounted}ms · 表格可用 ${marks.univerBooted}ms · ` +
        `DOMContentLoaded ${marks.domContentLoaded}ms · 传输 ${marks.transferredKB}KB`,
    );

    expect(marks.appMounted, '应用应挂载完成（app:mounted）').not.toBeNull();
    expect(marks.appMounted!, `应用挂载 ${marks.appMounted}ms 超出预算`).toBeLessThan(BUDGET.appMountedMs);
    expect(marks.univerBooted, '表格引擎应就绪（univer:booted）').not.toBeNull();
    expect(marks.univerBooted!, `表格可用 ${marks.univerBooted}ms 超出预算`).toBeLessThan(BUDGET.univerBootedMs);

    if (skeletonVisibleMs !== null) {
      expect(skeletonVisibleMs, `首屏骨架 ${skeletonVisibleMs}ms 才出现，太慢`).toBeLessThan(BUDGET.skeletonMs);
      expect(skeletonVisibleMs, '骨架必须早于应用挂载').toBeLessThan(marks.appMounted!);
    }
    if (marks.fcp !== null) {
      expect(marks.fcp, `首屏绘制 ${marks.fcp}ms 超出预算`).toBeLessThan(BUDGET.firstContentfulPaintMs);
      // 只有"带骨架的产物"才有这条不变量：骨架先画，应用再挂载（无骨架的产物第一帧就是应用本身）
      if (hasSkeleton) expect(marks.fcp, '首屏绘制应早于应用挂载').toBeLessThan(marks.appMounted!);
    }
  });

  /**
   * ③ 形态差异：静态版**没有**本机进程，`/api/bridge/*` 不存在，所以
   * 导出菜单里那三项（.ods/.xls/.xlsb —— 都要经本机 Excel）必须**置灰并写明原因**，
   * 而不是让用户点了等一个永远不会来的响应。
   */
  test('③ 静态形态：桥相关导出项置灰并说明"本地版可用"', async ({ page }) => {
    test.setTimeout(120_000);
    await page.goto(server.url);
    await page.waitForFunction(() => typeof (window as never as { __p0?: unknown }).__p0 !== 'undefined', null, {
      timeout: 60_000,
    });
    await page.setInputFiles('input[type=file]', FIXTURE);
    await expect.poll(() => page.locator('[data-testid="workspace-host"], canvas').count()).toBeGreaterThan(0);

    // 桥接口在静态托管上根本不该存在（这也是"探测不到"的根因）
    const health = await page.evaluate(async () => {
      const response = await fetch('/api/bridge/health');
      return { status: response.status };
    });
    expect(health.status, '静态托管没有桥接口').toBe(404);

    await page.click('[data-testid="toolbar-export"]');
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();

    // 保真导出与 CSV 在静态版照常可用
    await expect(page.locator('[data-testid="context-menu-export-xlsx-keep"]'), '.xlsx 保真导出在静态版可用').toBeEnabled();
    await expect(page.locator('[data-testid="context-menu-export-csv"]'), '.csv 导出在静态版可用').toBeEnabled();

    for (const id of ['export-ods', 'export-xls', 'export-xlsb']) {
      const item = page.locator(`[data-testid="context-menu-${id}"]`);
      await expect(item, `${id} 在静态版应置灰`).toBeDisabled();
      await expect(item, `${id} 应说明原因`).toContainText('本地版');
    }
  });
});
