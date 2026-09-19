/**
 * 右键菜单。
 *
 * 断言的是**行为**（不断言"菜单里没有样式项"这类实现细节）：
 *   ① 复制到工作区：区域被**拆成一个个独立单元格**（条目数 = 非空格子数）、原内容保留
 *   ② 剪切到工作区：同样逐格拆分、源内容清空但格式保留、可撤销
 *   ③ 清空内容：内容清掉、格式一字不变、可撤销
 *
 * 工作区条目恒为 1×1：任何入口都按**行优先**把区域拆开、跳过空单元格。
 *
 * 坑（别踩）：右键的坐标会**先把选区设成指针下那一格**，所以选区必须包住右键落点，
 * 否则菜单拿到的是空格快照、动作静默取消。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { a1ToRowCol, canvasBox, fixturePath, importFixture, waitForBoot } from './helpers';

async function openMenuOnCanvas(page: Page): Promise<void> {
  const box = await canvasBox(page);
  await page.mouse.click(box.x + 220, box.y + 120, { button: 'right' });
  await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
}

async function selectRange(page: Page, a1: string): Promise<void> {
  await page.evaluate((ref: string) => {
    (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange(ref);
  }, a1);
  await page.waitForTimeout(200);
}

/** 工作区条目（新模型下每条都是 1×1；来源是一个格子记号） */
async function workspaceItems(page: Page): Promise<Array<{ rows: number; cols: number; a1: string }>> {
  return page.evaluate(() =>
    (window as never as { __p0: { getWorkspaceItems: () => Array<{ rows: number; cols: number; a1: string }> } }).__p0.getWorkspaceItems(),
  );
}

/** 现算 A1:H25（25×8）里的非空格子数——与 `extractCellItems` 的口径一致（有值或有公式） */
async function countNonEmptyCells(page: Page, rows: number, cols: number): Promise<number> {
  return page.evaluate(
    ({ rows: r, cols: c }) => {
      const hooks = (window as never as {
        __p0: {
          getDisplayValue: (row: number, col: number) => string | number | boolean | null;
          getActiveSheet: () => { getRange: (row: number, col: number) => { getFormula: () => string | null } } | null;
        };
      }).__p0;
      const sheet = hooks.getActiveSheet();
      if (!sheet) return 0;
      let count = 0;
      for (let row = 0; row < r; row += 1) {
        for (let col = 0; col < c; col += 1) {
          const display = hooks.getDisplayValue(row, col);
          const text = display === null || display === undefined ? '' : String(display);
          const formula = sheet.getRange(row, col).getFormula();
          if (text !== '' || (typeof formula === 'string' && formula !== '')) count += 1;
        }
      }
      return count;
    },
    { rows, cols },
  );
}

test.describe('右键菜单', () => {
  test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures');

  test('① 复制到工作区：区域逐格拆开（条目数 = 非空格子数），原内容保留', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    await selectRange(page, 'A1:H25'); // 必须包住下面右键的落点
    const valueBefore = await page.evaluate(() =>
      (window as never as { __p0: { getValues: (a1: string) => unknown } }).__p0.getValues('A3'),
    );

    await openMenuOnCanvas(page);
    await page.locator('[data-testid="context-menu-workspace-copy"]').click();
    await page.waitForTimeout(400);
    await expect(page.locator('[data-testid="context-menu"]'), '动作执行后菜单应关闭').toHaveCount(0);

    // 新模型：A1:H25 的 200 格里只有"有值/有公式"的格子会变成条目，其余空格被跳过
    const expected = await countNonEmptyCells(page, 25, 8);
    expect(expected, 'A1:H25 里应有一批非空格子').toBeGreaterThan(0);
    expect(expected, '选的 200 格里必须有空格子').toBeLessThan(25 * 8);

    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(expected));

    const items = await workspaceItems(page);
    expect(items.length, '条目数 = 该区域非空格子数').toBe(expected);
    expect(items.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1 单格').toBe(true);

    const valueAfter = await page.evaluate(() =>
      (window as never as { __p0: { getValues: (a1: string) => unknown } }).__p0.getValues('A3'),
    );
    expect(JSON.stringify(valueAfter), '复制不应改动原内容').toEqual(JSON.stringify(valueBefore));
  });

  test('② 剪切到工作区：逐格拆分、源内容清空、格式保留、可撤销', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    await selectRange(page, 'A1:H25');
    const readA1 = () =>
      page.evaluate(() => {
        const hooks = (window as never as {
          __p0: { getDisplayValue: (r: number, c: number) => string | null; getCellStyle: (r: number, c: number) => unknown };
        }).__p0;
        return { value: hooks.getDisplayValue(0, 0), style: JSON.stringify(hooks.getCellStyle(0, 0) ?? null) };
      });

    const before = await readA1();
    expect(before.value, '测试前提：A1 需有内容').toBeTruthy();
    const expected = await countNonEmptyCells(page, 25, 8);

    await openMenuOnCanvas(page);
    await page.locator('[data-testid="context-menu-workspace-cut"]').click();
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    await page.waitForTimeout(300); // 等落盘/清空收尾

    const after = await readA1();
    const items = await workspaceItems(page);
    expect(items.length, '剪切应把整片区域的非空格子都放进工作区').toBe(expected);
    expect(items.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1 单格').toBe(true);
    expect(after.value ?? '', '剪切后源内容应被清空').not.toBe(before.value);
    expect(after.style, '剪切不得改变源单元格格式').toBe(before.style);

    await page.evaluate(async () => {
      await (window as never as { __p0: { undo: () => Promise<boolean> } }).__p0.undo();
    });
    await page.waitForTimeout(400);
    expect((await readA1()).value, '剪切应可撤销').toBe(before.value);
  });

  test('③ 清空内容：内容清掉、格式一字不变、可撤销', async ({ page }) => {
    await waitForBoot(page);
    await selectRange(page, 'A3');
    await openMenuOnCanvas(page);

    // 菜单动作作用于**打开那一刻的选区快照**，而右键坐标决定快照落在哪个格子。
    // 所以先从日志读出快照 A1，再断言那个格子的前后变化——不写死坐标对应的格子。
    const target = await page.evaluate(
      () =>
        ((window as never as { __p0: { log: Array<{ kind: string; detail?: { a1?: string } }> } }).__p0.log
          .filter((entry) => entry.kind === 'menu:open')
          .pop()?.detail?.a1 ?? '') as string,
    );
    expect(target, '菜单应记录打开时的选区快照').toMatch(/^[A-Z]+\d+$/);
    const { row, col } = a1ToRowCol(target);

    const read = () =>
      page.evaluate(
        ({ r, c }) => {
          const hooks = (window as never as {
            __p0: { getCellStyle: (r: number, c: number) => unknown; getDisplayValue: (r: number, c: number) => string | null };
          }).__p0;
          return { style: JSON.stringify(hooks.getCellStyle(r, c) ?? null), value: hooks.getDisplayValue(r, c) };
        },
        { r: row, c: col },
      );

    const before = await read();
    expect(before.value, `测试前提：${target} 需有内容`).toBeTruthy();

    await page.locator('[data-testid="context-menu-clear"]').click();
    await page.waitForTimeout(500);

    const after = await read();
    expect(after.value, `内容应被清空（${target}）`).not.toBe(before.value);
    expect(after.style, '清空内容不得改变格式').toBe(before.style);

    await page.evaluate(async () => {
      await (window as never as { __p0: { undo: () => Promise<boolean> } }).__p0.undo();
    });
    await page.waitForTimeout(400);
    expect((await read()).value, '清空内容应可撤销').toBe(before.value);
  });
});
