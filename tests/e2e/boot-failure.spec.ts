/**
 * **"打不开 / 白屏，刷新又好了"** 的第二类成因（能稳定复现的那一类）。
 *
 * 静态托管每次发布都会换掉分块 hash，而用户浏览器里可能还留着**旧的 index.html**
 * （弱网/代理把分块截断也一样）：旧 HTML 指向的分块已经不存在 → 动态 import 直接 reject。
 * 以前这里没有任何兜底，页面会**永远停在骨架屏**；按 F5 拿到新 HTML 就好 —— 正是用户描述的现象。
 *
 * 三条用例把出口钉住：
 *  ① 分块第一次失败 → **自动重来一次** → 应用正常起来（并且清掉重试标记，下次还能再救）
 *  ② 一直失败 → 摆出**可读卡片**（原因 + 「重新加载」），而不是让人对着灰块干等
 *  ③ 连入口脚本都没起来（JS 根本没跑）→ `index.html` 的看门狗摆出「重新加载」
 * 外加 ④ **不许误报**：正常启动时那个"卡住了"的提示不能留在页面上。
 */
import { expect, test, type Page } from '@playwright/test';

import { waitForBoot } from './helpers';

/** 数一数主框架被导航了几次（用来证明"自动重来一次"真的发生了） */
function countNavigations(page: Page): { get: () => number } {
  let navigations = 0;
  page.on('framenavigated', (frame) => {
    if (frame === page.mainFrame()) navigations += 1;
  });
  return { get: () => navigations };
}

test.describe('分块加载失败也有出口（回归）', () => {
  test('① 分块第一次失败：自动重来一次后应用正常起来', async ({ page }) => {
    test.setTimeout(180_000);
    let aborted = 0;
    await page.route('**/src/App.tsx*', async (route) => {
      if (aborted === 0) {
        aborted += 1;
        await route.abort();
        return;
      }
      await route.continue();
    });
    const nav = countNavigations(page);

    await page.goto('/');
    // 等应用真的挂上（说明重试拿到了模块）
    await page.waitForFunction(() => Boolean((window as never as { __p0?: unknown }).__p0), null, { timeout: 90_000 });

    expect(aborted, '第一次请求分块应当被打断（模拟"旧 index.html 指向已删除的分块"）').toBeGreaterThanOrEqual(1);
    expect(nav.get(), '应当自动重来过一次（reload 才会拿到新的 index.html）').toBeGreaterThanOrEqual(2);
    await expect(page.locator('[data-testid="boot-failed"]'), '自救了就不该摆失败卡').toHaveCount(0);
    await expect(page.locator('[data-testid="boot-slow"]'), '起来了就不该留"卡住了"的提示').toHaveCount(0);

    const flag = await page.evaluate(() => sessionStorage.getItem('app:chunk-retry'));
    expect(flag, '成功之后要清掉重试标记（以后真坏了还能再自动救一次）').toBeNull();
  });

  test('② 分块一直失败：摆出可读卡片 + 「重新加载」，不停在骨架屏', async ({ page }) => {
    test.setTimeout(180_000);
    let aborted = 0;
    await page.route('**/src/App.tsx*', async (route) => {
      aborted += 1;
      await route.abort();
    });

    await page.goto('/');
    const card = page.locator('[data-testid="boot-failed"]');
    await expect(card, '失败两次之后必须摆出卡片').toBeVisible({ timeout: 90_000 });
    await expect(page.locator('[data-testid="boot-failed-detail"]')).not.toBeEmpty();
    await expect(page.locator('[data-testid="app-crash-reload"]')).toBeVisible();
    await expect(page.locator('[data-app-skeleton]'), '卡片出现后骨架屏应当让位').toHaveCount(0);
    console.log(`[boot] 分块请求被打断 ${aborted} 次后摆出失败卡`);
    expect(aborted, '不该无限重试（最多自动重来一次）').toBeLessThanOrEqual(6);
  });

  test('③ 入口脚本没起来：看门狗摆出「重新加载」，不留死灰块', async ({ page }) => {
    test.setTimeout(120_000);
    // 把看门狗超时压到 600ms，用例不用等 20 秒
    await page.addInitScript(() => {
      (window as never as { __BOOT_TIMEOUT_MS?: number }).__BOOT_TIMEOUT_MS = 600;
    });
    await page.route('**/src/main.tsx*', async (route) => {
      await route.abort();
    });

    await page.goto('/');
    await expect(page.locator('[data-testid="boot-reload"]'), 'JS 完全没跑起来时也必须给出口').toBeVisible({
      timeout: 30_000,
    });
    await expect(page.locator('[data-testid="boot-slow"]')).toBeVisible();
  });

  test('④ 正常启动不误报：看门狗响了也会随应用挂载一起消失', async ({ page }) => {
    test.setTimeout(120_000);
    await page.addInitScript(() => {
      (window as never as { __BOOT_TIMEOUT_MS?: number }).__BOOT_TIMEOUT_MS = 600;
    });
    await waitForBoot(page);
    await expect(page.locator('[data-testid="boot-slow"]'), '应用起来后不该还挂着"卡住了"').toHaveCount(0);
    await expect(page.locator('[data-testid="boot-failed"]')).toHaveCount(0);
  });
});
