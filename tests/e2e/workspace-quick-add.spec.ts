/**
 * **快速加入**（用户要求的功能）：
 * "选择模式下，选中单元格后在三秒内点击工作区空白区域时，会把单元格加到空白区域"。
 *
 * 这里把"该生效"和"不该生效"都钉住：
 *  ① 选中一格 → 立刻点工作区空白 → 真的加进去了（并且可撤销、走的是唯一入口）；
 *  ② Ctrl+点选多块 → 点空白 → 多块一起进去；
 *  ③ 选区里有空格子 → 照旧跳过空内容（工作区里永远不放空条目）；
 *  ④ 同一次选区连点两下 → 只加一次；
 *  ⑤ 点到"看起来像空白"的交互元素（搜索框、按钮、条目）→ 不触发；
 *  ⑥ 拖到面板里松手（不是点击）→ 不触发；
 *  ⑦ 拖拽模式下点空白 → 不触发（那一模式的手势语言就是"按住拖"）；
 *     点击互换模式下**要触发**（用户后补的要求："点击交换模式也把这个功能加上去"）；
 *  ⑧ 撤销：Ctrl+Z 能把这次加入撤掉。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, logKinds, switchMode, waitForBoot } from './helpers';

const FIXTURE = 'fixture-styles.xlsx';

const items = (page: Page) =>
  page.evaluate(
    () =>
      (window as never as { __p0: { getWorkspaceItems: () => Array<{ id: string; a1: string }> } }).__p0.getWorkspaceItems(),
  );

/** 单元格在视口里的矩形（渲染服务刚起来时可能量不出 → 轮询） */
async function waitRect(page: Page, a1: string): Promise<{ left: number; top: number; width: number; height: number }> {
  const deadline = Date.now() + 15_000;
  for (;;) {
    const rect = await page.evaluate(
      (ref: string) =>
        (window as never as {
          __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null };
        }).__p0.rectOfA1(ref),
      a1,
    );
    if (rect) return rect;
    if (Date.now() > deadline) throw new Error(`${a1} 一直量不出矩形`);
    await page.waitForTimeout(200);
  }
}

/** 点一下表格里的某个格子（= 选中它，同时开始"3 秒窗口"） */
async function clickCell(page: Page, a1: string): Promise<void> {
  const rect = await waitRect(page, a1);
  await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2);
  await page.waitForTimeout(120);
}

/** 工作区面板底部的空白处 */
async function blankPoint(page: Page): Promise<{ x: number; y: number }> {
  const panel = await page.locator('.ws-panel').boundingBox();
  if (!panel) throw new Error('工作区不可见');
  return { x: panel.x + panel.width / 2, y: panel.y + panel.height - 36 };
}

test.describe('工作区快速加入：选中后点空白就收进来', () => {
  test.skip(!existsSync(fixturePath(FIXTURE)), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await switchMode(page, 'select');
  });

  test('① 选中一格后点工作区空白 → 加入工作区，并可撤销', async ({ page }) => {
    test.setTimeout(120_000);
    await clickCell(page, 'A2');
    const blank = await blankPoint(page);
    await page.mouse.click(blank.x, blank.y);
    await page.waitForTimeout(400);

    const added = await items(page);
    expect(added.length, '应加入 1 个单元格').toBe(1);
    expect(added[0].a1, '来源就是刚选中的那一格').toBe('A2');
    expect(await logKinds(page), '应记录快速加入').toContain('workspace:quick-add');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(1);

    // 可撤销（走的是唯一入口 commitWorkspace）
    await page.keyboard.press('Control+z');
    await expect.poll(() => items(page).then((list) => list.length), { timeout: 8_000 }).toBe(0);
    await page.keyboard.press('Control+y');
    await expect.poll(() => items(page).then((list) => list.length), { timeout: 8_000 }).toBe(1);
  });

  test('② Ctrl+点选多块 → 一次全部加入', async ({ page }) => {
    test.setTimeout(120_000);
    const first = await waitRect(page, 'A2');
    await page.mouse.click(first.left + first.width / 2, first.top + first.height / 2);
    await page.waitForTimeout(120);
    const second = await waitRect(page, 'E4');
    // Playwright 的 mouse.click 没有 modifiers 参数：自己按住 Ctrl 再点（和真实操作一致）
    await page.keyboard.down('Control');
    await page.mouse.click(second.left + second.width / 2, second.top + second.height / 2);
    await page.keyboard.up('Control');
    await page.waitForTimeout(150);
    const selection = await page.evaluate(() =>
      (window as never as { __p0: { getSelectionA1: () => string | null } }).__p0.getSelectionA1(),
    );
    expect(selection, '测试前提：应是多块选区').toBeTruthy();

    const blank = await blankPoint(page);
    await page.mouse.click(blank.x, blank.y);
    await page.waitForTimeout(400);

    const added = await items(page);
    expect(added.length, '两块都该进来（各 1 格）').toBeGreaterThanOrEqual(2);
    expect(new Set(added.map((item) => item.a1)).size, '来源格子不重复').toBe(added.length);
  });

  test('③ 选区里含空格子 → 只收有内容的（工作区不放空条目）', async ({ page }) => {
    test.setTimeout(120_000);
    // A1:H2 里 A1 是合并标题、B1:F1 是空格子；A2:F2 是表头
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:H2');
    });
    // selectRange 是程序化选区，不算"刚动手" → 先真实点一下同一片区域里的格子，打开时间窗
    await clickCell(page, 'A2');
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:H2');
    });
    const blank = await blankPoint(page);
    await page.mouse.click(blank.x, blank.y);
    await page.waitForTimeout(400);

    const added = await items(page);
    expect(added.length, '应加入若干格').toBeGreaterThan(0);
    expect(added.length, '16 格里不可能全有内容 → 空格子被跳过').toBeLessThan(16);
    const texts = await page
      .locator('[data-testid="snapshot-preview"]')
      .evaluateAll((nodes) => nodes.map((node) => (node.textContent ?? '').trim()));
    expect(texts.every((text) => text !== ''), '不该出现空内容条目').toBe(true);
  });

  test('④ 同一次选区连点两下 → 只加一次', async ({ page }) => {
    test.setTimeout(120_000);
    await clickCell(page, 'A2');
    const blank = await blankPoint(page);
    await page.mouse.click(blank.x, blank.y);
    await page.waitForTimeout(300);
    await page.mouse.click(blank.x, blank.y);
    await page.waitForTimeout(300);

    expect(await items(page), '连点两下也只该有 1 条').toHaveLength(1);
    expect(await logKinds(page), '第二次应记"同一次选区重复点击"').toContain('workspace:quick-add-duplicate');
  });

  test('⑤ 点到交互元素（搜索框 / 按钮 / 条目）不触发快速加入', async ({ page }) => {
    test.setTimeout(120_000);
    // 先放一条，让面板出现搜索框与条目
    await page.evaluate(() => {
      (window as never as { __p0: { snapshotSelectionToWorkspace: (a1: string) => unknown } }).__p0.snapshotSelectionToWorkspace('A1');
    });
    await page.waitForTimeout(300);
    await clickCell(page, 'A2');

    await page.locator('[data-testid="workspace-search"]').click();
    await page.waitForTimeout(200);
    await page.locator('[data-testid="workspace-item"]').first().click();
    await page.waitForTimeout(200);
    await page.locator('[data-testid="workspace-add"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toBeVisible();
    await page.keyboard.press('Escape');
    await page.waitForTimeout(200);

    expect(await items(page), '点这些交互元素都不该改变条目数').toHaveLength(1);
  });

  test('⑥ 拖到面板里松手走拖拽（不被记成快速加入）；⑦ 点击互换模式也能快速加入，拖拽模式不动手', async ({ page }) => {
    test.setTimeout(120_000);

    // ⑥ 从表格拖一片到面板空白：这是"拖入工作区"的正常路径，走的是拖拽而不是快速加入
    //（要先把模式切到"拖拽"——选择模式下按住拖是原生框选，本来就不搬内容）
    await switchMode(page, 'drag');
    const rect = await waitRect(page, 'A2');
    const blank = await blankPoint(page);
    await page.mouse.move(rect.left + rect.width / 2, rect.top + rect.height / 2);
    await page.mouse.down();
    await page.mouse.move(blank.x - 60, blank.y - 40, { steps: 6 });
    await page.mouse.move(blank.x, blank.y, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    const afterDrag = await items(page);
    expect(afterDrag.length, '拖拽入工作区照旧生效').toBeGreaterThan(0);
    expect(await logKinds(page), '这一次不该被记成"快速加入"').not.toContain('workspace:quick-add');

    // ⑦ 拖拽模式下点空白：不做任何事（该模式的手势语言是"按住拖"）
    await page.evaluate(() => (window as never as { __p0: { clearWorkspace: () => void } }).__p0.clearWorkspace());
    await page.waitForTimeout(200);
    await clickCell(page, 'A2');
    const dragBlank = await blankPoint(page);
    await page.mouse.click(dragBlank.x, dragBlank.y);
    await page.waitForTimeout(300);
    expect(await items(page), '拖拽模式下点空白不该加入').toHaveLength(0);
    expect(await logKinds(page), '应记录因模式跳过').toContain('workspace:quick-add-skipped');

    // ⑧ 点击互换模式：用户要求"也把这个功能加上去" → 选中后点空白要加入
    await switchMode(page, 'click-swap');
    await clickCell(page, 'A2');
    const swapBlank = await blankPoint(page);
    await page.mouse.click(swapBlank.x, swapBlank.y);
    await page.waitForTimeout(400);
    const inSwapMode = await items(page);
    expect(inSwapMode.length, '点击互换模式下点空白应加入').toBeGreaterThan(0);
    expect(inSwapMode[0].a1, '来源就是刚选中的格子').toBe('A2');
    expect(await logKinds(page), '应记录快速加入').toContain('workspace:quick-add');
  });
});
