/**
 * 日期单元格 / "文本格装数字"提醒 回归（用户实测反馈的 bug）。
 *
 * 现场：拖动日期单元格时弹出两个**原始 i18n key**：
 * `sheets-ui.info.error` + `sheets-ui.info.forceStringInfo`。
 *
 * 查证结论与对应修法（本文件把它钉住）：
 *  ①取值口径：日期格的 raw 就是序列号（如 45658），绝不能变成
 *    `"Wed Jan 01 2025 08:00:00 GMT+0800 (中国标准时间)"` 这种字符串（旧归一化就是这么干的）；
 *  ②弹窗提醒：`sheets-ui.config.disableForceStringAlert` 必须真的生效。上游触发条件是
 *    "单元格 `t` 是 STRING/FORCE_STRING 且值看起来是数字、且数字格式不是文本格式"——
 *    本产品锁死格式，用户没有"转成数字"的补救动作可做，弹窗纯属噪音（且上游连文案 key 都漏配，
 *    任何语言下都只显示 key 本身）。只关弹窗，单元格左上角的角标仍保留。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { canvasBox, fixturePath, importFixture, selectionA1, waitForBoot } from './helpers';

/** 页面里是否存在"强制按文本处理"的提醒（Univer 的提醒是浮层，可能在容器外，所以全文档找） */
async function alertTexts(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const nodes = Array.from(document.body.querySelectorAll('*'));
    const texts = nodes
      .map((node) => (node.textContent ?? '').trim())
      .filter(
        (text) =>
          text.length < 200 &&
          (text.includes('sheets-ui.info') || text.includes('强制') || text.includes('按文本处理') || text.includes('文本格式')),
      );
    return Array.from(new Set(texts)).slice(0, 6);
  });
}

/** 读某个 A1 的原始模型值与显示文本（日期格：raw 应是 number 序列号） */
async function probeCell(page: Page, a1: string) {
  return page.evaluate((ref: string) => {
    const hooks = (window as never as {
      __p0: {
        getActiveSheet: () => { getRange: (a: string) => { getRawValues?: () => unknown[][]; getValues: () => unknown[][] } };
        getDisplayValueByA1: (a1: string) => string | null;
      };
    }).__p0;
    const range = hooks.getActiveSheet().getRange(ref);
    const raw = range.getRawValues?.()?.[0]?.[0];
    const value = range.getValues()?.[0]?.[0];
    return {
      rawNumber: typeof raw === 'number' ? raw : null,
      rawKind: raw === null || raw === undefined ? 'null' : raw instanceof Date ? 'Date' : typeof raw,
      rawText: raw === null || raw === undefined ? 'null' : String(raw).slice(0, 40),
      valueText: value === null || value === undefined ? 'null' : String(value).slice(0, 40),
      display: hooks.getDisplayValueByA1(ref),
    };
  }, a1);
}

test.describe('日期单元格：搬家不毁值、不弹 i18n key', () => {
  test.skip(!existsSync(fixturePath('fixture-numfmt.xlsx')), '请先运行 npm run fixtures 生成样本');

  test.beforeEach(async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-numfmt.xlsx');
  });

  test('① 日期格的 raw 是数字序列号，不是日期字符串', async ({ page }) => {
    const probe = await probeCell(page, 'B6');

    expect(probe.display, 'B6 是 yyyy-mm-dd 的日期格').toMatch(/\d{4}[-/]\d{1,2}[-/]\d{1,2}/);
    expect(probe.rawKind, `日期格 raw 应是 number，实际 ${probe.rawText}`).toBe('number');
    expect(probe.rawNumber, '序列号应与 Excel 语义一致').toBe(45658);
    expect(probe.rawText).not.toContain('GMT');
  });

  test('② 日期格互换：搬过去的是序列号本身（不是被 String() 过的日期）', async ({ page }) => {
    const before = await probeCell(page, 'B6');
    expect(before.rawNumber).toBe(45658);

    // B6（日期，带 yyyy-mm-dd 格式）⇄ B20（空，无格式）
    await page.evaluate(() => {
      (window as never as { __p0: { swap: (a: string, b: string) => unknown } }).__p0.swap('B6', 'B20');
    });
    await page.waitForTimeout(400);

    const atTarget = await probeCell(page, 'B20');
    // 内容搬过去了、格式没搬（产品约定：格式锁死不动）→ 目标格显示 serial 原文是正确行为
    expect(atTarget.rawNumber, `B20 应拿到序列号 45658，实际 ${atTarget.rawText}`).toBe(45658);
    expect(atTarget.rawText, '绝不能是被 String() 的本地时间字符串').not.toContain('GMT');
    expect(atTarget.valueText).not.toContain('GMT');

    const atSource = await probeCell(page, 'B6');
    expect(atSource.rawKind, 'B6 应是空的（内容已换走）').toBe('null');
    expect(atSource.display, '空格显示为空（格式还在，只是没内容）').toBe('');

    expect(await alertTexts(page), '互换日期后不应弹出"强制按文本处理"提醒').toEqual([]);
  });

  test('③ 日期格进工作区再写回：条目里存的是序列号，写回后显示仍是日期', async ({ page }) => {
    const snapshot = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          snapshotSelectionToWorkspace: (a1?: string) => {
            ids?: string[];
            items: Array<{ source: { a1: string }; values: unknown[][] }>;
          };
        };
      }).__p0;
      const result = hooks.snapshotSelectionToWorkspace('B6');
      return {
        ids: result.ids ?? [],
        count: result.items.length,
        a1: result.items[0]?.source.a1 ?? null,
        storedValue: result.items[0]?.values?.[0]?.[0] ?? null,
      };
    });

    expect(snapshot.count, '日期格应产出 1 个工作区条目').toBe(1);
    expect(snapshot.a1).toBe('B6');
    expect(snapshot.storedValue, '工作区条目里存的必须是序列号（存字符串就再也变不回日期了）').toBe(45658);

    await page.evaluate((id: string) => {
      (window as never as { __p0: { pasteWorkspaceItem: (id: string, a1: string) => unknown } }).__p0.pasteWorkspaceItem(id, 'B6');
    }, snapshot.ids[0]);
    await page.waitForTimeout(300);

    const after = await probeCell(page, 'B6');
    expect(after.display, '写回后仍是日期显示（格式没被破坏）').toMatch(/\d{4}[-/]\d{1,2}[-/]\d{1,2}/);
    expect(after.rawNumber).toBe(45658);
    expect(await alertTexts(page)).toEqual([]);
  });

  test('④ "文本格装数字"的悬浮提醒：反向对照能抓到，产品配置下不再出现', async ({ page }) => {
    const disabled = await page.evaluate(() =>
      (window as never as { __p0: { isForceStringAlertDisabled: () => boolean } }).__p0.isForceStringAlertDisabled(),
    );
    expect(disabled, 'sheets-ui.config.disableForceStringAlert 应为 true').toBe(true);

    // 上游触发点是**鼠标悬停**（`IHoverManagerService.currentCell$`），不是选中：
    // 这也解释了用户"拖动时"看到弹窗——拖动过程中指针会扫过一个个格子。
    // 所以这里用真实鼠标移动，而不是 selectRange。
    const box = await canvasBox(page);
    const x = box.x + 175;
    const y = box.y + 110;

    await page.mouse.click(x, y);
    await page.waitForTimeout(200);
    const hovered = await selectionA1(page);
    expect(hovered, '点击后应有活动单元格').toMatch(/^[A-Z]+\d+$/);

    // 造一个**必然满足上游判定**的格子：t=STRING(1) + 值看起来是数字 + 数字格式不是文本格式
    const created = await page.evaluate((a1: string) => {
      const hooks = (window as never as {
        __p0: {
          setStyledValues: (row: number, col: number, matrix: unknown[][]) => void;
          getActiveSheet: () => { getRange: (a: string) => { getRawValues: () => unknown[][] } };
        };
      }).__p0;
      const rect = (window as never as { __p0: { getActiveSheet: () => { getRange: (a: string) => { getRange: () => { startRow: number; startColumn: number } } } } }).__p0
        .getActiveSheet()
        .getRange(a1)
        .getRange();
      hooks.setStyledValues(rect.startRow, rect.startColumn, [[{ v: '20240101', t: 1 }]]);
      return hooks.getActiveSheet().getRange(a1).getRawValues()[0][0];
    }, hovered!);
    expect(created, '造出来的格子应是"看起来是数字"的文本').toBe('20240101');

    /** 悬停到目标格：先移开再移入，确保产生 currentCell 变化 */
    const hoverCell = async (): Promise<void> => {
      await page.mouse.move(box.x + box.width - 40, box.y + box.height - 40);
      await page.waitForTimeout(150);
      await page.mouse.move(x, y, { steps: 4 });
    };

    // ---- 反向对照：把配置改回去，确认这个测试**抓得到**那个弹窗 ----
    const flipped = await page.evaluate(() =>
      (window as never as { __p0: { setForceStringAlertDisabled: (v: boolean) => boolean | null } }).__p0.setForceStringAlertDisabled(false),
    );
    expect(flipped, '测试开关应能把配置改回 false').toBe(false);
    await hoverCell();
    /**
     * 用轮询等弹窗出现，而不是"睡 600ms 再看一眼"：
     * 系统忙时（例如测试并发跑）弹窗可能晚一拍才出来，固定 sleep 会假红。
     * 而且**每次重试都重新悬停一下** —— 上游是按 `currentCell$` 的变化触发的，
     * 只读 DOM 不复现"指针移动"这一下，忙的时候可能一次都没触发出事件。
     * 反向对照必须是"确实抓到过"，否则下面"不再出现"的断言就是假的。
     */
    await expect
      .poll(
        async () => {
          await page.mouse.move(x + 3, y + 2, { steps: 2 });
          await page.waitForTimeout(120);
          return (await alertTexts(page)).length;
        },
        { timeout: 20_000, message: '关掉"关弹窗"配置后应能抓到提醒' },
      )
      .toBeGreaterThan(0);
    const controlAlerts = await alertTexts(page);
    expect(controlAlerts.length, `反向对照抓到的提醒：${controlAlerts.join(' | ')}`).toBeGreaterThan(0);

    // ---- 产品配置：恢复 true，同一个格子不再弹 ----
    const restored = await page.evaluate(() =>
      (window as never as { __p0: { setForceStringAlertDisabled: (v: boolean) => boolean | null } }).__p0.setForceStringAlertDisabled(true),
    );
    expect(restored).toBe(true);
    await hoverCell();
    // 反向断言（"不应再出现"）只能靠"等一会再看"：给足时间，确保不是"还没来得及弹"
    await page.waitForTimeout(1200);
    expect(await alertTexts(page), '产品配置下悬停不应再出现提醒').toEqual([]);
  });
});
