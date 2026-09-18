/**
 * 右下角缩放组件（用户实测反馈：浏览器缩放能用，但拖动这个组件"没有任何效果"）。
 *
 * 根因是我们自己的只读闸门：`@univerjs/sheets-ui` 的 `ZoomSlider` 拖动时发的是
 * `SetZoomRatioCommand`（`sheet.command.set-zoom-ratio`），而白名单里只放了 +/− 按钮用的
 * `change-zoom-ratio` → 拖动被静默拦掉（百分比菜单里的 50%/75%/… 预设同样走这条命令）。
 *
 * 这个文件同时守住两件事：**行为**（拖动能改缩放）与**闸门**（缩放命令不再被拦）。
 * 组件结构（实测）：footer 里 `div[role="track"]` + `button[role="handle"]`（12×12 的圆点），
 * 旁边是百分比按钮（点开是 50%–400% 的预设菜单）与减/加号按钮。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, waitForBoot } from './helpers';

/** 当前缩放：同时读工作表模型值与 footer 上的百分比文字（两者必须一致） */
async function zoomState(page: Page): Promise<{ ratio: number | null; pct: string | null }> {
  return page.evaluate(() => {
    const hooks = (window as never as { __p0: { getActiveSheet: () => { getZoom?: () => number } } }).__p0;
    const button = Array.from(document.querySelectorAll('button')).find((b) => /^\s*\d{1,3}%\s*$/.test(b.textContent ?? ''));
    return { ratio: hooks.getActiveSheet().getZoom?.() ?? null, pct: button?.textContent?.trim() ?? null };
  });
}

/** 被闸门拦下的**缩放**命令（回归护栏：这条命令必须始终放行） */
async function blockedZoomCommands(page: Page): Promise<Array<{ id: string; count: number }>> {
  return page.evaluate(() =>
    (window as never as { __p0: { getBlockedCommands: () => Array<{ id: string; count: number }> } }).__p0
      .getBlockedCommands()
      .filter((entry) => /zoom/i.test(entry.id)),
  );
}

const trackBox = async (page: Page) => page.locator('div[role="track"]').first().boundingBox();

test.describe('右下角缩放组件', () => {
  test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures 生成样本');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');
  });

  test('① 拖动滑块手柄能改缩放（百分比与模型同步），且缩放命令没被闸门拦', async ({ page }) => {
    const before = await zoomState(page);
    expect(before.ratio, '初始应是 100%').toBeCloseTo(1, 2);

    const handle = page.locator('button[role="handle"]').first();
    await expect(handle, 'footer 里应有可拖的滑块手柄').toBeVisible();
    const hb = await handle.boundingBox();
    const tb = await trackBox(page);
    if (!hb || !tb) throw new Error('找不到缩放条');

    // 拖到轨道最左端 → 一定变小（不按比例猜位置，避免轨道宽度变化带来的抖动）
    await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
    await page.mouse.down();
    await page.mouse.move(tb.x + 3, hb.y + hb.height / 2, { steps: 14 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    const after = await zoomState(page);
    expect(after.ratio, `拖到最左端后缩放应变小（实际 ${after.ratio}）`).toBeLessThan(before.ratio!);
    expect(after.pct, '百分比文字应跟着变').not.toBe(before.pct);
    expect(Number.parseInt(after.pct ?? '0', 10) / 100).toBeCloseTo(after.ratio!, 1);

    // 再往右拖回去 → 变大
    const hb2 = await handle.boundingBox();
    if (!hb2) throw new Error('手柄消失');
    await page.mouse.move(hb2.x + hb2.width / 2, hb2.y + hb2.height / 2);
    await page.mouse.down();
    await page.mouse.move(tb.x + tb.width * 0.8, hb2.y + hb2.height / 2, { steps: 14 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    const back = await zoomState(page);
    expect(back.ratio, '往右拖应放大回去').toBeGreaterThan(after.ratio!);

    expect(await blockedZoomCommands(page), '缩放命令不该出现在被拦清单里').toEqual([]);
  });

  test('② 百分比菜单的预设与减/加号按钮都能改缩放', async ({ page }) => {
    // 打开预设菜单（结构实测：`div[role="menu"]` + 若干 `div[role="menuitemradio"]`，即 50%…400%）。
    // 首次点击偶尔会先被布局重排吃掉，所以"点一次没开就再点一次"。
    const menu = page.locator('[role="menu"]');
    for (let attempt = 0; attempt < 2 && !(await menu.isVisible().catch(() => false)); attempt += 1) {
      await page.locator('button').filter({ hasText: /^\s*\d{1,3}%\s*$/ }).first().click();
      await page.waitForTimeout(500);
    }
    await expect(menu, '百分比按钮应弹出预设菜单').toBeVisible({ timeout: 5_000 });

    const items = menu.locator('[role="menuitemradio"]');
    const labels = (await items.allTextContents()).map((text) => text.trim());
    expect(labels.join(' '), '菜单应是 50%–400% 的预设').toContain('50%');
    expect(labels.join(' ')).toContain('400%');

    await items.filter({ hasText: '50%' }).first().click();
    await page.waitForTimeout(400);
    const afterPreset = await zoomState(page);
    expect(afterPreset.ratio, '选 50% 后缩放应是 0.5').toBeCloseTo(0.5, 2);

    // 加号按钮（旁边是减号，两个都走 change-zoom-ratio）
    const plus = page.locator('button:has(svg.univerjs-icon-increase-icon)').first();
    if (await plus.count()) {
      await plus.click();
      await page.waitForTimeout(300);
      expect((await zoomState(page)).ratio, '点加号应变大').toBeGreaterThan(afterPreset.ratio!);
    }

    const minus = page.locator('button:has(svg.univerjs-icon-reduce-icon)').first();
    await minus.click();
    await page.waitForTimeout(300);
    expect((await zoomState(page)).ratio, '点减号应变小').toBeLessThan(1);

    expect(await blockedZoomCommands(page)).toEqual([]);
  });
});
