/**
 * 工作区拖放的**误操作守卫**（用户实测反馈两条）：
 *
 *  ① "我在工作区拖动单元格到工作区空白地方，会自我复制" —— 根因：落点判断没看拖拽来源，
 *     "工作区条目拖到工作区自己身上"被当成"从表格拖进来"，把同一个快照又收了一份（新 id、同来源）。
 *  ② "选择模式下，选中单元格后点击工作区空白区域时，会把单元格加到工作区" —— 根因：拖动阈值只有 4px，
 *     按下的手抖就算一次拖动，松手落在工作区上就执行了落点动作（点一下变成了"拖进去"）。
 *
 * 这里把两条都钉住，并顺带守住"正常功能没被修坏"：条目拖到**表格单元格**上仍要写回。
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

/** 造一个条目（走生产同款提交入口），返回它的 id */
async function seedItem(page: Page, a1: string): Promise<string> {
  return page.evaluate((ref: string) => {
    const result = (window as never as {
      __p0: { snapshotSelectionToWorkspace: (a1: string) => { ids: string[] } };
    }).__p0.snapshotSelectionToWorkspace(ref);
    return result.ids[0];
  }, a1);
}

/** 取单元格在视口里的矩形（渲染服务刚起来时可能还没量出来 → 轮询等一会） */
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
    if (Date.now() > deadline) throw new Error(`${a1} 一直量不出矩形（不在可见区域？）`);
    await page.waitForTimeout(200);
  }
}

test.describe('工作区拖放：不该发生的"自我复制 / 误加入"', () => {
  test.skip(!existsSync(fixturePath(FIXTURE)), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
  });

  test('① 条目拖到工作区面板的空白处：不复制、不改内容，只提示"已经在工作区里"', async ({ page }) => {
    test.setTimeout(120_000);
    const id = await seedItem(page, 'A1');
    expect(await items(page), '先确认工作区里只有 1 条').toHaveLength(1);

    const tile = await page.locator('[data-testid="workspace-item"]').first().boundingBox();
    const panel = await page.locator('.ws-panel').boundingBox();
    if (!tile || !panel) throw new Error('工作区不可见');

    // 真实手势：按住条目 → 拖到面板下方空白处 → 松手
    await page.mouse.move(tile.x + tile.width / 2, tile.y + tile.height / 2);
    await page.mouse.down();
    await page.mouse.move(tile.x + 24, tile.y + 20, { steps: 5 });
    await page.mouse.move(panel.x + panel.width / 2, panel.y + panel.height - 36, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    const after = await items(page);
    expect(after, '拖到面板空白处不应产生第二条（用户实测的"自我复制"）').toHaveLength(1);
    expect(after[0].id, '还是原来那一条').toBe(id);
    expect(await logKinds(page), '应记录"落点在工作区自身、已忽略"').toContain('workspace:drop-self-ignored');
    expect(await logKinds(page), '不该产生新的加入记录').not.toContain('workspace:add');
    await expect(page.locator('[data-testid="toast"]').last(), '要给一句人话，而不是静默无反应').toContainText(
      '已经在工作区里',
    );
  });

  test('② 条目拖到另一个条目上：同样不复制', async ({ page }) => {
    test.setTimeout(120_000);
    await seedItem(page, 'A1');
    await seedItem(page, 'B2');
    expect(await items(page), '两条待用').toHaveLength(2);

    const tiles = page.locator('[data-testid="workspace-item"]');
    const first = await tiles.nth(0).boundingBox();
    const second = await tiles.nth(1).boundingBox();
    if (!first || !second) throw new Error('条目不可见');

    await page.mouse.move(first.x + first.width / 2, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(second.x + second.width / 2, second.y + second.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    expect(await items(page), '面板内部拖动不应该改变条目数量').toHaveLength(2);
  });

  test('③ 选择模式：选中单元格后"点"工作区空白处（含 5px 手抖）不会把单元格加进去', async ({ page }) => {
    test.setTimeout(120_000);
    await switchMode(page, 'select');
    const panel = await page.locator('.ws-panel').boundingBox();
    if (!panel) throw new Error('工作区不可见');

    const clickBlank = async (): Promise<void> => {
      const rect = await waitRect(page, 'B3');
      await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2);
      await page.waitForTimeout(150);
      // 面板空白处：按下 → 抖 5px（小于 8px 阈值）→ 松手
      const blankX = panel.x + panel.width / 2;
      const blankY = panel.y + panel.height - 36;
      await page.mouse.move(blankX, blankY);
      await page.mouse.down();
      await page.mouse.move(blankX + 3, blankY + 4, { steps: 2 });
      await page.mouse.up();
      await page.waitForTimeout(400);
    };

    await clickBlank();
    expect(await items(page), '空工作区：点空白 + 手抖也不该凭空多出一条').toHaveLength(0);

    // 工作区里有条目时再试一次（点列表下方空白）
    await seedItem(page, 'A1');
    await clickBlank();
    const after = await items(page);
    expect(after, '有 1 条时点空白也不该多出第二条').toHaveLength(1);
    expect(after[0].a1, '而且不该换成刚选中的 B3').toBe('A1');
    expect(await logKinds(page), '不该有新的加入记录').not.toContain('workspace:add');
  });

  test('④ 修好之后正常功能仍在：条目拖到表格单元格上照旧写回', async ({ page }) => {
    test.setTimeout(120_000);
    await seedItem(page, 'A1');
    const target = await waitRect(page, 'D22');

    const tile = await page.locator('[data-testid="workspace-item"]').first().boundingBox();
    if (!tile) throw new Error('条目不可见');
    await page.mouse.move(tile.x + tile.width / 2, tile.y + tile.height / 2);
    await page.mouse.down();
    await page.mouse.move(tile.x - 30, tile.y + 10, { steps: 4 });
    await page.mouse.move(target.left + target.width / 2, target.top + target.height / 2, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(500);

    expect(await logKinds(page), '写回表格仍然要发生').toContain('app:workspace-paste');
    const written = await page.evaluate(() =>
      (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('D22'),
    );
    expect(written ?? '', '目标格应拿到内容').not.toBe('');
    expect(await items(page), '条目不会被顺手删掉（设置未开时）').toHaveLength(1);
  });
});
