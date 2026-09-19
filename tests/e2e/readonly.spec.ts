/**
 * 只读约束。
 *
 * 产品的核心约束：**只允许改单元格文字内容**，字体/边框/背景/结构/对象一律改不动。
 *
 * 三条：
 *   ① 22 条破坏性命令（样式/结构/对象/子表）逐条被拦下，且拦截前后**文档指纹零变化**
 *   ② mutation 层兜底：夹带样式的写入必须把样式剥离干净、但值要写进去（防"粘贴带格式"绕过闸门）
 *   ③ 真实键盘输入：内容写得进去、样式一字不变、且可撤销（防闸门误伤编辑路径）
 */
import { expect, test, type Page } from '@playwright/test';

import { waitForBoot } from './helpers';

/** 采集"不该被改动"的文档状态指纹 */
async function captureFingerprint(page: Page): Promise<Record<string, unknown>> {
  return page.evaluate(() => {
    const hooks = (window as never as {
      __p0: {
        getActiveSheet: () => {
          getSheetName: () => string;
          getMaxRows: () => number;
          getMaxColumns: () => number;
          getSheet: () => { getMergeData?: () => unknown[] };
          getRange: (a1: string) => { getCellStyleData: (t?: string) => unknown };
          getImages?: () => unknown[];
          getNotes?: () => unknown[];
          getConditionalFormattingRules?: () => unknown[];
          getDataValidations?: () => unknown[];
        } | null;
        getSheetNames: () => string[];
      };
    }).__p0;

    const sheet = hooks.getActiveSheet();
    const style = (a1: string) => JSON.stringify(sheet?.getRange(a1).getCellStyleData('cell') ?? null);
    let mergeCount = -1;
    try {
      mergeCount = sheet?.getSheet().getMergeData?.().length ?? -1;
    } catch {
      mergeCount = -2;
    }
    return {
      sheets: hooks.getSheetNames(),
      rows: sheet?.getMaxRows() ?? -1,
      cols: sheet?.getMaxColumns() ?? -1,
      merges: mergeCount,
      images: sheet?.getImages?.().length ?? -1,
      notes: sheet?.getNotes?.().length ?? -1,
      cfRules: sheet?.getConditionalFormattingRules?.().length ?? -1,
      dvRules: sheet?.getDataValidations?.().length ?? -1,
      styleTitle: style('A1'),
      styleBody: style('A3'),
      styleHeader: style('A2'),
    };
  });
}

test.describe('只读约束：格式/结构改不动，内容可改', () => {
  test('① 样式/结构/对象/子表四类命令全被拦下，且文档指纹零变化', async ({ page }) => {
    await waitForBoot(page);
    const before = await captureFingerprint(page);

    const forbidden = [
      // 样式
      'sheet.command.set-style',
      'sheet.command.set-background-color',
      'sheet.command.set-text-color',
      'sheet.command.set-bold',
      'sheet.command.set-italic',
      'sheet.command.set-border',
      'sheet.command.set-font-size',
      // 结构
      'sheet.command.insert-row',
      'sheet.command.insert-col-before',
      'sheet.command.remove-row',
      'sheet.command.add-worksheet-merge',
      'sheet.command.set-col-width',
      'sheet.command.set-row-height',
      'sheet.command.move-rows',
      // 对象与规则
      'sheet.command.insert-image',
      'sheet.command.delete-note',
      'sheet.command.add-hyper-link',
      'sheet.command.set-data-validation',
      'sheet.command.add-conditional-formatting-rule',
      // 子表
      'sheet.command.insert-sheet',
      'sheet.command.remove-sheet',
      'sheet.command.set-worksheet-name',
    ];

    const results = await page.evaluate(async (ids: string[]) => {
      const hooks = (window as never as {
        __p0: { runCommand: (id: string, params?: unknown) => Promise<unknown> };
      }).__p0;
      const out: Array<{ id: string; result: unknown }> = [];
      for (const id of ids) {
        // 尽量给出合法参数，确保"是被闸门拦下"而不是"因为参数不合法而失败"
        const result = await hooks.runCommand(id, {
          unitId: 'p0-workbook',
          subUnitId: 'p0-sheet-1',
          range: { startRow: 2, startColumn: 0, endRow: 2, endColumn: 0 },
          value: 123,
          style: { bg: { rgb: '#FF00FF' } },
        });
        out.push({ id, result: Boolean(result) });
      }
      return out;
    }, forbidden);

    const after = await captureFingerprint(page);
    const blocked = await page.evaluate(() =>
      (window as never as { __p0: { getBlockedCommands: () => Array<{ id: string }> } }).__p0.getBlockedCommands(),
    );

    const succeeded = results.filter((item) => item.result === true).map((item) => item.id);
    expect(succeeded, `这些破坏性命令居然执行成功了：${succeeded.join(', ')}`).toEqual([]);

    const blockedIds = blocked.map((item) => item.id);
    for (const id of forbidden) {
      expect(blockedIds, `${id} 应被闸门记录为已拦下`).toContain(id);
    }

    // 文档指纹必须一字不变
    expect(after, '文档状态被破坏了').toEqual(before);
  });

  test('② mutation 层兜底：夹带样式的写入被剥离样式，但值要写进去', async ({ page }) => {
    await waitForBoot(page);

    await page.evaluate(() => {
      // 故意夹带样式：紫红背景 + 粗体 + 边框（模拟"粘贴带格式"的路径）
      (window as never as { __p0: { setStyledValues: (r: number, c: number, m: unknown[][]) => void } }).__p0.setStyledValues(
        9,
        0,
        [
          [
            {
              v: '带样式写入',
              s: { bg: { rgb: '#FF00FF' }, bl: 1, bd: { t: { s: 1, cl: { rgb: '#FF00FF' } } } },
            },
          ],
        ],
      );
    });
    await page.waitForTimeout(400);

    const after = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { getCellStyle: (r: number, c: number) => unknown; getDisplayValue: (r: number, c: number) => string | null };
      }).__p0;
      return { style: hooks.getCellStyle(9, 0), value: hooks.getDisplayValue(9, 0) };
    });

    expect(after.value, '值必须写入成功（内容可编辑）').toBe('带样式写入');
    expect(
      JSON.stringify(after.style ?? {}).toUpperCase(),
      '夹带的样式必须被剥离（不允许通过任何路径改格式）',
    ).not.toContain('FF00FF');
  });

  test('③ 真实键盘输入：内容改成功、样式一字不变、撤销可还原', async ({ page }) => {
    await waitForBoot(page);

    const read = () =>
      page.evaluate(() => {
        const hooks = (window as never as {
          __p0: { getCellStyle: (r: number, c: number) => unknown; getDisplayValue: (r: number, c: number) => string | null };
        }).__p0;
        return { style: JSON.stringify(hooks.getCellStyle(2, 0) ?? null), value: hooks.getDisplayValue(2, 0) };
      });

    const before = await read();

    // 先点一下网格让画布获得焦点，再选中目标单元格（模拟真人操作）
    await page.locator('#univer-container canvas[id^="univer-sheet-main-canvas"]').first().click({ position: { x: 240, y: 140 } });
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A3');
    });
    await page.waitForTimeout(200);

    // 真键盘输入：选中后直接打字即进入编辑态。英文用 type（逐键），中文必须用 insertText（逐键拿不到 IME 文本）。
    // 这条路径曾因只读闸门漏放行 docs 子系统而失效（甚至回车清空单元格），本用例是它的回归网。
    await page.keyboard.type('Edited');
    await page.keyboard.insertText('键盘输入测试');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);

    const afterType = await read();
    expect(afterType.value, '键盘输入的内容必须写入').toBe('Edited键盘输入测试');
    expect(afterType.style, '键盘输入不得改变样式').toBe(before.style);

    // 撤销一次应还原内容（证明输入的编辑进入了撤销栈）
    await page.evaluate(async () => {
      await (window as never as { __p0: { undo: () => Promise<boolean> } }).__p0.undo();
    });
    await page.waitForTimeout(400);
    const afterUndo = await read();
    expect(afterUndo.value, '撤销应还原为原内容').toBe(before.value);
    expect(afterUndo.style, '撤销也不得改变样式').toBe(before.style);
  });
});
