/**
 * 内存泄漏回归：**插图 blob url 必须随工作簿释放**。
 *
 * 背景（真实泄漏）：导入时插入浮动图片要先 `URL.createObjectURL(blob)`，而图片服务是
 * "插图之后再按 url 取图"的异步流程，所以**不能立刻 revoke**。原实现把它登记进一个模块级
 * `Set` 然后**页面生命周期内永不回收** —— 于是每开一份带图的表就多钉住一批图片字节，
 * 关标签、冷存标签都还不回来（冷存把工作簿 dispose 了，图片字节却还留着，等于白冷存）。
 *
 * 现在的口径：url 按**工作簿**归档，关标签 / 冷存标签 / 整实例拆卸三处都要回收；
 * 冷存标签切回时会重新走 `applyWorkbookFeatures` 重新插图（所以回收是安全的、有回归盯着）。
 *
 * 断言用的是 `__p0.retainedImageObjectUrls()` 这个**确定性计数**，不是堆采样，
 * 所以没有"GC 没跑"这类噪声：数字不对就是真的没回收。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, waitForBoot } from './helpers';

/** 带浮动图片的夹具（批注 / 超链接 / 浮动图片） */
const WITH_IMAGES = 'fixture-extras.xlsx';
const PLAIN = ['fixture-styles.xlsx', 'fixture-numfmt.xlsx', 'fixture-table.xlsx'];

interface Tabs {
  id: string;
  fileName: string;
}

const tabs = (page: Page): Promise<Tabs[]> =>
  page.evaluate(() => (window as never as { __p0: { getTabs: () => Tabs[] } }).__p0.getTabs());

const retainedUrls = (page: Page): Promise<{ total: number; byWorkbook: Record<string, number> }> =>
  page.evaluate(() =>
    (window as never as { __p0: { retainedImageObjectUrls: () => { total: number; byWorkbook: Record<string, number> } } }).__p0.retainedImageObjectUrls(),
  );

/**
 * 走工具栏真实点击关闭标签（与用户操作一致）。
 *
 * 注意：带特性（图片/批注/超链接）的表**导入后就已经算"有改动"**——应用特性时会写单元格，
 * 我们的脏格账本照实记账，标签上因此出现"未保存"，关标签会先弹一次 `window.confirm`。
 * Playwright 默认**取消**对话框，不接住它标签就关不掉（实测就是这么卡的），所以这里明确接受。
 */
async function closeTab(page: Page, id: string): Promise<void> {
  page.once('dialog', (dialog) => void dialog.accept());
  await page.locator(`[data-testid="toolbar-tab-close-${id}"]`).click();
  await page.waitForFunction(
    (target) => !(window as never as { __p0: { getTabs: () => Tabs[] } }).__p0.getTabs().some((tab) => tab.id === target),
    id,
    { timeout: 30_000 },
  );
}

test.describe('内存泄漏：插图 blob url 随工作簿释放', () => {
  test.skip(!existsSync(fixturePath(WITH_IMAGES)), '请先运行 npm run fixtures 生成样本');

  test('① 关标签要把该簿的插图 blob url 全部回收（以前是"页面生命周期内永不回收"）', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);

    const summary = await importFixture(page, WITH_IMAGES, { features: true });
    expect(summary.parsedFeatures?.images ?? 0, `前提：${WITH_IMAGES} 必须真的有浮动图片`).toBeGreaterThan(0);

    const before = await retainedUrls(page);
    expect(before.total, '导入带图的表之后应有未回收的插图 url').toBeGreaterThan(0);

    const [tab] = await tabs(page);
    expect(tab, '前提：应有一个标签').toBeTruthy();
    expect(before.byWorkbook[tab.id], 'url 要登记在**这个工作簿**名下').toBe(before.total);

    await closeTab(page, tab.id);

    const after = await retainedUrls(page);
    expect(after.total, `关标签后必须归还全部插图 url，实际还剩 ${JSON.stringify(after.byWorkbook)}`).toBe(0);
  });

  test('② 被冷存的标签也要回收；切回重建后图片重新登记（不是把图弄没了）', async ({ page }) => {
    test.setTimeout(180_000);
    await waitForBoot(page);

    await importFixture(page, WITH_IMAGES, { features: true });
    const firstTab = (await tabs(page))[0];
    expect((await retainedUrls(page)).byWorkbook[firstTab.id], '第一个标签的图先要有 url').toBeGreaterThan(0);

    // 常驻窗口是 3：再开 3 个 → 第一个被冷存（`planEvictions`）
    for (const name of PLAIN) await importFixture(page, name, { features: true });
    await page.waitForTimeout(500);

    const afterEvict = await retainedUrls(page);
    expect(
      afterEvict.byWorkbook[firstTab.id] ?? 0,
      `第一个标签已被冷存（工作簿 dispose），它的插图 url 必须一起还回去：${JSON.stringify(afterEvict.byWorkbook)}`,
    ).toBe(0);

    // 切回去 → 重建 → 重新插图 → 新的 url 被登记（证明回收没有把"以后还能插图"这条路堵死）
    await page.locator(`[data-testid="toolbar-tab-${firstTab.id}"]`).click();
    await page.waitForFunction(
      (target) => (window as never as { __p0: { getActiveTabId: () => string | null } }).__p0.getActiveTabId() === target,
      firstTab.id,
      { timeout: 60_000 },
    );
    await page.waitForTimeout(800);

    const afterRebuild = await retainedUrls(page);
    expect(
      afterRebuild.byWorkbook[firstTab.id] ?? 0,
      '切回重建后应重新建立该簿的插图 url（否则图片会变空白）',
    ).toBeGreaterThan(0);
  });

  test('③ 关掉全部标签后：插图 url、标签账本都回到空（不留残渣）', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);

    await importFixture(page, WITH_IMAGES, { features: true });
    await importFixture(page, PLAIN[0], { features: true });
    expect((await retainedUrls(page)).total).toBeGreaterThan(0);

    for (const tab of await tabs(page)) await closeTab(page, tab.id);

    expect(await tabs(page), '标签应全部关掉').toHaveLength(0);
    const urls = await retainedUrls(page);
    expect(urls.total, `全部关闭后不该还有未回收的插图 url：${JSON.stringify(urls.byWorkbook)}`).toBe(0);
    expect(
      await page.evaluate(() => (window as never as { __p0: { getDirtySummary: () => unknown[] } }).__p0.getDirtySummary()),
      '脏单元格账本也不该留残渣',
    ).toHaveLength(0);
  });
});
