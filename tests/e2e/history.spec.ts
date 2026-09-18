/**
 * 撤销与历史（合并自原 p2-interaction.spec.ts 的撤销原子性 + p3-history.spec.ts）。
 *
 * 不变量：
 *   ① 一次互换 = **一次撤销**（两个格子一起还原，不允许只回退一半）
 *   ② 历史面板逐条入账，可回退到任意一步，也能前滚回最新
 *
 * 已删除："没有操作时显示空占位"（纯空态 UI，非关键不变量）。
 *
 * 注意口径（由历史面板定义）：面板第 index 项的 `data-testid="history-item-<index>"`，
 * 点击时回调收到的是**步下标** `index + 1`（0 = 最初状态，N = 全部已应用）。
 */
import { expect, test } from '@playwright/test';

import { waitForBoot } from './helpers';

test.describe('撤销与历史', () => {
  test('① 一次互换 = 一次撤销：两个单元格一起还原', async ({ page }) => {
    await waitForBoot(page);

    const outcome = await page.evaluate(async () => {
      const hooks = (window as never as {
        __p0: {
          getValues: (a1: string) => unknown;
          swap: (a: string, b: string) => { ok: boolean; reason?: string };
          undo: () => Promise<boolean>;
        };
      }).__p0;

      const before = { a: hooks.getValues('A3'), b: hooks.getValues('A4') };
      const swapped = hooks.swap('A3', 'A4');
      await new Promise((r) => setTimeout(r, 300));
      const afterSwap = { a: hooks.getValues('A3'), b: hooks.getValues('A4') };
      await hooks.undo();
      await new Promise((r) => setTimeout(r, 300));
      const afterUndo = { a: hooks.getValues('A3'), b: hooks.getValues('A4') };
      return { before, swapped, afterSwap, afterUndo };
    });

    expect(outcome.swapped.ok, '互换命令必须执行成功').toBe(true);
    expect(outcome.afterSwap.a, '互换后 A3 应变成原 A4 的值').toEqual(outcome.before.b);
    expect(outcome.afterSwap.b, '互换后 A4 应变成原 A3 的值').toEqual(outcome.before.a);
    expect(outcome.afterUndo.a, '一次撤销后 A3 应还原').toEqual(outcome.before.a);
    expect(outcome.afterUndo.b, '一次撤销后 A4 应还原').toEqual(outcome.before.b);
  });

  test('② 动作逐条入账，可跳回任意一步并前滚到最新', async ({ page }) => {
    await waitForBoot(page);

    const original = await page.evaluate(() =>
      (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0.getDisplayValue(2, 0),
    );

    // 造三步可撤销的操作：两次改内容 + 一次互换
    await page.evaluate(async () => {
      const hooks = (window as never as {
        __p0: {
          setValue: (r: number, c: number, v: string) => void;
          swap: (a: string, b: string) => { ok: boolean };
        };
      }).__p0;
      hooks.setValue(2, 0, '第1步');
      await new Promise((r) => setTimeout(r, 350));
      hooks.setValue(4, 0, '第2步');
      await new Promise((r) => setTimeout(r, 350));
      hooks.swap('A3', 'A5');
      await new Promise((r) => setTimeout(r, 450));
    });

    const afterActions = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0;
      return { a3: hooks.getDisplayValue(2, 0), a5: hooks.getDisplayValue(4, 0) };
    });

    await page.click('[data-testid="toolbar-history"]');
    await expect(page.locator('[data-testid="history-panel"]')).toBeVisible();

    const itemCount = await page.locator('[data-testid^="history-item-"]').count();
    const panelText = await page.locator('[data-testid="history-panel"]').innerText();
    expect(itemCount, '历史里应记下我们的操作').toBeGreaterThanOrEqual(2);
    expect(panelText, '互换动作应带着可读标签入账').toContain('互换');

    // 面板第 index 项 = "步下标 index + 1"（0 表示最初状态，只能靠撤销走到）：
    // 点第 1 项就是回到"第 1 步之后"。
    await page.click('[data-testid="history-item-0"]');
    await page.waitForTimeout(700);
    expect(
      await page.evaluate(() =>
        (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0.getDisplayValue(2, 0),
      ),
      '回到第 1 步后 A3 应是第 1 步写进去的内容',
    ).toBe('第1步');

    // 再按一次 Ctrl+Z 才是最初状态（账本与 Univer 的撤销栈必须仍然同步）
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(400);
    expect(
      await page.evaluate(() =>
        (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0.getDisplayValue(2, 0),
      ),
      '再撤一次应回到最初内容',
    ).toBe(original);

    // 前滚到最新：内容应回到三步操作之后的状态（从最初一口气重做全部）
    await page.click(`[data-testid="history-item-${itemCount - 1}"]`);
    await page.waitForTimeout(700);
    const atLatest = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0;
      return { a3: hooks.getDisplayValue(2, 0), a5: hooks.getDisplayValue(4, 0) };
    });
    expect(atLatest.a3, '前滚到最新后应恢复操作后的内容').toBe(afterActions.a3);
    expect(atLatest.a5).toBe(afterActions.a5);

    await page.click('[data-testid="history-close"]');
    await expect(page.locator('[data-testid="history-panel"]')).toHaveCount(0);
  });
});
