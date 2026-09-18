/**
 * 真实课程表（合并单元格密集）下的"选区 = 被操作范围"回归。
 *
 * 用户实测："选择多个后右键剪切到工作区，剪切的内容并非是我选中的那部分"。
 * 根因：拖选 B4:F8 时 Univer 会把**模型选区**扩成整块 B4:F11（第 9~11 行有竖向合并），
 * 剪切自然剪走整块。我们现在的规则是：右键时把范围补齐到整块（若补齐过就提示 + 同步高亮），
 * 保证"高亮看到的 = 实际被操作的"。
 *
 * 用例用用户自己的文件（存在才跑）；同时用夹具做一条不依赖合并的对照。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { importFixture, userFixturePath, waitForBoot } from './helpers';

const USER_FILE = userFixturePath('课程表');

const selection = (page: Page): Promise<string | null> =>
  page.evaluate(() => (window as never as { __p0: { getSelectionA1: () => string | null } }).__p0.getSelectionA1());

const rectOf = (page: Page, a1: string) =>
  page.evaluate(
    (ref) =>
      (window as never as { __p0: { rectOfA1: (a: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1(
        ref,
      ),
    a1,
  );

/** 打开用户文件（不存在则跳过） */
async function openUserFile(page: Page): Promise<void> {
  await page.setInputFiles('[data-testid="file-input"]', USER_FILE!);
  await page.waitForFunction(
    () => Boolean((window as never as { __p0: { getImportSummary: () => unknown } }).__p0.getImportSummary()),
    null,
    { timeout: 120_000 },
  );
  await page.waitForTimeout(800);
}

test.describe('选区与合并单元格对齐', () => {
  test('① 真实课程表：拖选被合并块撑大后，右键会提示并按整块处理，且提示 = 实际选中范围', async ({ page }) => {
    test.setTimeout(240_000);
    test.skip(!USER_FILE || !existsSync(USER_FILE), '未配置用户文件（USER_EXCEL 或 tests/e2e/user-fixture.local），跳过');
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await openUserFile(page);
    await page.locator('[data-testid="toolbar-mode-select"]').click();
    await page.waitForTimeout(200);

    // 拖选 B4 → F8（这一片会被第 9~11 行的竖向合并撑大）
    const from = (await rectOf(page, 'B4'))!;
    const to = (await rectOf(page, 'F8'))!;
    await page.mouse.move(from.left + from.width / 2, from.top + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.left + to.width / 2, to.top + to.height / 2, { steps: 14 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const dragged = await selection(page);
    expect(dragged, '拖选后应有选区').toBeTruthy();

    // 状态栏那个"选区 …"必须跟着真实选区走（以前只在切表时刷新，会一直显示 A1，误导用户）
    await page.waitForTimeout(500);
    const chipText = await page.evaluate(() => {
      const chips = Array.from(document.querySelectorAll('.statusbar .chip'));
      return chips.map((n) => (n.textContent ?? '').trim()).find((text) => text.startsWith('选区 ')) ?? null;
    });
    expect(chipText, `状态栏应显示真实选区（实际 ${chipText}，真实 ${dragged}）`).toBe(`选区 ${dragged}`);

    // 在选区内右键 → 菜单快照必须等于当前真实选区（不许"菜单说一套、选区另一套"）
    const inside = (await rectOf(page, 'C6'))!;
    await page.mouse.click(inside.left + inside.width / 2, inside.top + inside.height / 2, { button: 'right' });
    await page.waitForTimeout(400);

    const state = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          log: Array<{ kind: string; detail?: Record<string, unknown> }>;
          getSelectionA1: () => string | null;
          getBlockedCommands: () => Array<{ id: string }>;
        };
      }).__p0;
      const kinds = hooks.log.map((e) => e.kind);
      return {
        menuSnapshot: hooks.log.filter((e) => e.kind === 'menu:open').pop()?.detail ?? null,
        expanded: hooks.log.filter((e) => e.kind === 'menu:range-expanded-for-merge').pop()?.detail ?? null,
        selection: hooks.getSelectionA1(),
        kinds,
      };
    });

    expect(state.menuSnapshot?.a1, '菜单拿到的范围应等于当前真实选区').toBe(state.selection);
    expect(state.menuSnapshot?.a1, '菜单范围应包含拖选起点 B4').toContain('B4');
    // 若确实被合并块撑大，必须留下日志与提示（用户能看懂"为什么范围变大了"）
    if (state.expanded) {
      // eslint-disable-next-line no-console
      console.log('[merge] 选区被合并块补齐：', JSON.stringify(state.expanded));
      expect(String(state.expanded.to)).toBe(state.selection);
    }
    await page.keyboard.press('Escape');
  });

  test('② 剪切到工作区：工作区条目全部来自选中范围，范围之外的格子内容不变', async ({ page }) => {
    test.setTimeout(240_000);
    test.skip(!USER_FILE || !existsSync(USER_FILE), '未配置用户文件（USER_EXCEL 或 tests/e2e/user-fixture.local），跳过');
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await openUserFile(page);

    // 选一片明确的区域（B4:D6），并记录选区外一格的内容
    const snapshot = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          selectRange: (a1: string) => void;
          getSelectionA1: () => string | null;
          getDisplayValueByA1: (a1: string) => string | null;
          getActiveSheet: () => { getRange: (a: string) => { getDisplayValues: () => string[][] } };
        };
      }).__p0;
      hooks.selectRange('B4:D6');
      return {
        selection: hooks.getSelectionA1(),
        outside: hooks.getActiveSheet().getRange('B7:D11').getDisplayValues(),
        inside: hooks.getActiveSheet().getRange('B4:D6').getDisplayValues(),
      };
    });
    expect(snapshot.selection).toBe('B4:D6');

    const inside = (await rectOf(page, 'C5'))!;
    await page.mouse.click(inside.left + inside.width / 2, inside.top + inside.height / 2, { button: 'right' });
    await page.waitForTimeout(350);
    await page.locator('[data-testid="context-menu"]').getByText('剪切到工作区', { exact: false }).first().click();
    await page.waitForTimeout(800);

    const result = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getWorkspaceItems: () => Array<{ a1: string; label: string }>;
          getActiveSheet: () => { getRange: (a: string) => { getDisplayValues: () => string[][] } };
        };
      }).__p0;
      return {
        items: hooks.getWorkspaceItems(),
        inside: hooks.getActiveSheet().getRange('B4:D6').getDisplayValues(),
        outside: hooks.getActiveSheet().getRange('B7:D11').getDisplayValues(),
      };
    });

    // 条目必须全部落在选区里
    expect(result.items.length, '应产出若干条目').toBeGreaterThan(0);
    const letters = ['B', 'C', 'D'];
    result.items.forEach((item) => {
      const match = /^([A-D])(\d+)$/.exec(item.a1);
      expect(match, `条目 ${item.a1} 应是选区内的单格`).not.toBeNull();
      const col = match![1];
      const row = Number(match![2]);
      expect(letters, `条目列 ${col} 应在 B..D 内`).toContain(col);
      expect(row, `条目行 ${row} 应在 4..6 内`).toBeGreaterThanOrEqual(4);
      expect(row).toBeLessThanOrEqual(6);
    });

    // 选区内被清空；选区外（B7:D11）一字未改
    const insideNow = JSON.stringify(result.inside);
    expect(insideNow.replace(/[",\[\]]/g, '').trim(), '选区内应被清空').toBe('');
    expect(result.outside, '选区之外的内容不能被动到').toEqual(snapshot.outside);
    expect(snapshot.inside.length, '测试前提：选区内原本有内容').toBeGreaterThan(0);
  });

  test('③ 对照：夹具里选区不含合并块时，菜单范围 = 拖选范围（不无谓扩大）', async ({ page }) => {
    test.setTimeout(150_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.locator('[data-testid="toolbar-mode-select"]').click();
    await page.waitForTimeout(200);

    const from = (await rectOf(page, 'B5'))!;
    const to = (await rectOf(page, 'D7'))!;
    await page.mouse.move(from.left + from.width / 2, from.top + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.left + to.width / 2, to.top + to.height / 2, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(250);
    expect(await selection(page)).toBe('B5:D7');

    const inside = (await rectOf(page, 'C6'))!;
    await page.mouse.click(inside.left + inside.width / 2, inside.top + inside.height / 2, { button: 'right' });
    await page.waitForTimeout(350);
    const snapshot = await page.evaluate(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string; detail?: Record<string, unknown> }> } }).__p0.log
          .filter((e) => e.kind === 'menu:open')
          .pop()?.detail ?? null,
    );
    expect(snapshot?.a1, '没有合并块时不该扩大范围').toBe('B5:D7');
    await page.keyboard.press('Escape');
  });
});
