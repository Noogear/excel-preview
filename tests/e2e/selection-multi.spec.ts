/**
 * **不连续多区域选区**（Ctrl+点选 / Ctrl+框选）。
 *
 * 用户要求："请实现不连续多区域（Ctrl+点选）"。
 *
 * 事实边界（先读 Univer 源码确认，再写用例）：
 *  - Univer 的选区模型本来就支持多块：`@univerjs/sheets-ui` 的
 *    `_checkClearPreviousControls(evt)` 在按住 Ctrl/Shift 时**不清掉**已有选区框，
 *    随后 `newSelectionControl(...)` 会**再建一个**选区控件；门面也提供
 *    `FWorksheet.getSelection().getActiveRangeList()` 读全部区域。
 *  - 我们要做的不是"造轮子"，而是让**自己的功能读懂多块**：状态栏、右键菜单、
 *    工作区导入对话框以前都只读 `getActiveRange()` 这一块。
 *  - 天然只有一方的操作（与工作区互换、粘贴）在多块下**明确禁用并说明原因**，
 *    而不是随便挑一块执行。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, switchMode, waitForBoot } from './helpers';

const FIXTURE = 'fixture-styles.xlsx';

const rectOf = (page: Page, a1: string) =>
  page.evaluate(
    (ref) =>
      (
        window as never as {
          __p0: { rectOfA1: (a: string) => { left: number; top: number; width: number; height: number } | null };
        }
      ).__p0.rectOfA1(ref),
    a1,
  );

/** 点某格（可带修饰键） */
async function clickCell(page: Page, a1: string, modifiers: Array<'Control' | 'Shift'> = []): Promise<void> {
  const rect = await rectOf(page, a1);
  expect(rect, `应能量到 ${a1} 的位置`).not.toBeNull();
  if (!rect) return;
  for (const key of modifiers) await page.keyboard.down(key);
  await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2);
  for (const key of modifiers) await page.keyboard.up(key);
  await page.waitForTimeout(250);
}

/** 从 from 拖到 to（可带修饰键） */
async function dragCells(
  page: Page,
  from: string,
  to: string,
  modifiers: Array<'Control' | 'Shift'> = [],
): Promise<void> {
  const a = await rectOf(page, from);
  const b = await rectOf(page, to);
  expect(a && b, `应能量到 ${from} / ${to}`).toBeTruthy();
  if (!a || !b) return;
  for (const key of modifiers) await page.keyboard.down(key);
  await page.mouse.move(a.left + a.width / 2, a.top + a.height / 2);
  await page.mouse.down();
  await page.mouse.move(b.left + b.width / 2, b.top + b.height / 2, { steps: 10 });
  await page.mouse.up();
  for (const key of modifiers) await page.keyboard.up(key);
  await page.waitForTimeout(300);
}

/** 直接读 Univer 的选区模型：有几块、分别是哪些 A1 */
const modelRanges = (page: Page): Promise<string[]> =>
  page.evaluate(() => {
    const sheet = (window as never as { __p0: { getActiveSheet: () => unknown } }).__p0.getActiveSheet() as {
      getSelection?: () => { getActiveRangeList: () => Array<{ getA1Notation: () => string }> } | null;
    } | null;
    const selection = sheet?.getSelection?.() ?? null;
    return selection ? selection.getActiveRangeList().map((range) => range.getA1Notation()) : [];
  });

/** 状态栏里那句"选区 …" */
const statusSelection = (page: Page): Promise<string> =>
  page.evaluate(() => {
    const chips = Array.from(document.querySelectorAll('.statusbar .chip')).map((node) => (node.textContent ?? '').trim());
    return chips.find((text) => text.startsWith('选区')) ?? '';
  });

test.describe('不连续多区域（Ctrl+点选）', () => {
  test.skip(!existsSync(fixturePath(FIXTURE)), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await switchMode(page, 'select');
    await page.waitForTimeout(200);
  });

  test('① Ctrl+点选累积多块；再普通点一下回到单块；状态栏与模型都说实话', async ({ page }) => {
    test.setTimeout(120_000);

    await clickCell(page, 'B2');
    expect(await modelRanges(page), '单击只有一块').toHaveLength(1);
    expect(await statusSelection(page)).toContain('B2');

    await clickCell(page, 'D4', ['Control']);
    expect(await modelRanges(page), 'Ctrl+点选应累积成两块').toEqual(['B2', 'D4']);
    const multi = await statusSelection(page);
    expect(multi, '状态栏要报出块数与总格数').toContain('B2 + D4');
    expect(multi).toContain('2 块');

    await clickCell(page, 'F6', ['Control']);
    expect(await modelRanges(page)).toEqual(['B2', 'D4', 'F6']);
    expect(await statusSelection(page)).toContain('3 块');

    // 不带修饰键的普通点击：回到单块
    await clickCell(page, 'A12');
    expect(await modelRanges(page), '普通点击应重置为单块').toEqual(['A12']);
    expect(await statusSelection(page)).toBe('选区 A12');
  });

  test('② Ctrl+框选同样能追加一块（第二次框出来的矩形成为第二块）', async ({ page }) => {
    test.setTimeout(120_000);
    await clickCell(page, 'B2');
    await dragCells(page, 'D4', 'E5', ['Control']);
    expect(await modelRanges(page), 'Ctrl+拖选应得到 B2 与 D4:E5 两块').toEqual(['B2', 'D4:E5']);
    expect(await statusSelection(page)).toContain('2 块');
    expect(await statusSelection(page)).toContain('5 格'); // 1 + 2×2
  });

  test('③ 右键菜单：追加类操作按"全部块"执行，互换/粘贴在多块下禁用并说明原因', async ({ page }) => {
    test.setTimeout(120_000);
    await clickCell(page, 'B2');
    await clickCell(page, 'D4', ['Control']);

    // 右键落在第一块里（右键落点决定菜单，但菜单要带上全部块）
    const rect = await rectOf(page, 'B2');
    expect(rect).not.toBeNull();
    if (!rect) return;
    await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2, { button: 'right' });
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();

    const copyToWorkspace = page.locator('[data-testid="context-menu-workspace-copy"]');
    await expect(copyToWorkspace, '标签要点明"2 块"').toContainText('2 块');
    // 互换是"两方"操作：多块时必须禁用（并把原因写在 title / hint 里）
    const swap = page.locator('[data-testid="context-menu-swap-workspace"]');
    await expect(swap).toHaveAttribute('aria-disabled', 'true');
    await expect(swap).toHaveAttribute('title', /多块|2 块/);

    // 复制到工作区：条目数 = 两块里非空单元格数（现算，不写死）
    const expected = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => unknown } }).__p0;
      const nonEmpty = (r: number, c: number): number => (String(hooks.getDisplayValue(r, c) ?? '').trim() === '' ? 0 : 1);
      return nonEmpty(1, 1) + nonEmpty(3, 3); // B2 + D4
    });
    expect(expected, '夹具里 B2 / D4 至少有一个有内容').toBeGreaterThan(0);
    await copyToWorkspace.click();
    await page.waitForTimeout(400);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    const sources = await page.evaluate(() =>
      (
        window as never as { __p0: { getWorkspaceItems: () => Array<{ a1: string }> } }
      ).__p0.getWorkspaceItems().map((item) => item.a1),
    );
    expect(sources.every((a1) => a1 === 'B2' || a1 === 'D4'), `条目应只来自这两块：${sources.join(',')}`).toBe(true);
  });

  test('④ 工作区「+」导入：输入多块（A1:B2 D4:E5）也能一次搬进来', async ({ page }) => {
    test.setTimeout(120_000);
    await page.locator('[data-testid="workspace-add"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toBeVisible();
    await page.locator('[data-testid="import-input"]').fill('A1:B2 D4:E5');
    await expect(page.locator('[data-testid="import-detected"]'), '识别行要点明两块').toContainText('2 块');
    await expect(page.locator('[data-testid="import-confirm"]')).toBeEnabled();
    await page.locator('[data-testid="import-confirm"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toHaveCount(0);

    const expected = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => unknown } }).__p0;
      let count = 0;
      for (const [rowFrom, rowTo, colFrom, colTo] of [
        [0, 1, 0, 1],
        [3, 4, 3, 4],
      ]) {
        for (let r = rowFrom; r <= rowTo; r += 1) {
          for (let c = colFrom; c <= colTo; c += 1) {
            if (String(hooks.getDisplayValue(r, c) ?? '').trim() !== '') count += 1;
          }
        }
      }
      return count;
    });
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    expect(expected, '这两块里应有非空内容').toBeGreaterThan(0);
  });
});
