/**
 * 剪贴板两个方向的行为：**复制出去**（到 Excel / WPS / 在线表格）与**粘进来**（从外部）。
 *
 * 缘起（用户提问原文）："复制网页表格的内容后可以粘贴到本地的表格软件或者其他表格软件里？"
 * 以及"对于粘贴行为，是否会覆盖格式？"
 *
 * 这里盯四件曾经是"真实缺陷 / 未验证项"的事：
 *  ① **Ctrl+C 必须真的写系统剪贴板** —— 曾经被只读闸门用错命令 id 拦掉，
 *     表现是"按了完全没反应、也不报错"（`sheet.command.copy` 是不存在的 id，
 *     真 id 是 `univer.command.copy`）。
 *  ② 复制要**带格式**：`text/plain`(TSV) + `text/html`(内联样式/合并/列宽/行高)。
 *  ③ 多块选区只能走纯文本（原生复制只认最后一个选区），菜单要如实写明。
 *  ④ 从外部粘**带格式的 HTML** 进来：值要写进去，但目标格格式**一格都不能变**，也不得引入富文本。
 *
 * 注意：本文件读写的是**真实系统剪贴板**（机器级共享资源），因此在 playwright.config.ts
 * 里归到 `sensitive`（串行）项目，避免与并发用例互相污染。
 */
import { expect, test, type Page } from '@playwright/test';

import { importFixture, switchMode, waitForBoot } from './helpers';

test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

const FIXTURE = 'fixture-styles.xlsx';

interface ClipFlavor {
  types: string[];
  text: string | null;
  html: string | null;
}

/** 读系统剪贴板的第一条 item 的两种口味 */
async function readClipboard(page: Page): Promise<ClipFlavor> {
  return page.evaluate(async () => {
    const items = await navigator.clipboard.read();
    const item = items[0];
    if (!item) return { types: [], text: null, html: null };
    return {
      types: [...item.types],
      text: item.types.includes('text/plain') ? await (await item.getType('text/plain')).text() : null,
      html: item.types.includes('text/html') ? await (await item.getType('text/html')).text() : null,
    };
  });
}

/** 往系统剪贴板写一个哨兵值：这样"复制到底发生没有"才有判据（否则可能读到上一轮的残留） */
async function writeSentinel(page: Page, text: string): Promise<void> {
  await page.evaluate(async (value: string) => {
    await navigator.clipboard.writeText(value);
  }, text);
}

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

/**
 * 等几何量出来再取（`rectOfA1` 依赖 canvas/scene/skeleton 就绪，
 * 刚导入完的那一瞬间还会返回 null —— 直接用会得到"量不到位置"的假失败）。
 */
async function rectOfReady(page: Page, a1: string): Promise<{ left: number; top: number; width: number; height: number }> {
  await page.waitForFunction(
    (ref) =>
      Boolean(
        (window as never as { __p0: { rectOfA1: (a: string) => unknown } }).__p0.rectOfA1(ref),
      ),
    a1,
    { timeout: 20_000 },
  );
  const rect = await rectOf(page, a1);
  if (!rect) throw new Error(`量不到 ${a1} 的位置`);
  return rect;
}

/** 引导 + 导入夹具 + 切到"选择"模式（拖动/点选模式会改掉点格子的语义，用例里不需要） */
async function bootWithFixture(page: Page): Promise<void> {
  await waitForBoot(page);
  await importFixture(page, FIXTURE, { features: true });
  await switchMode(page, 'select');
  await page.waitForTimeout(200);
}

/** 点某格（可带修饰键），用来建立"单块 / 多块"选区 */
async function clickCell(page: Page, a1: string, modifiers: Array<'Control' | 'Shift'> = []): Promise<void> {
  const rect = await rectOfReady(page, a1);
  for (const key of modifiers) await page.keyboard.down(key);
  await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2);
  for (const key of modifiers) await page.keyboard.up(key);
  await page.waitForTimeout(250);
}

/** 右键点某格并等菜单出现 */
async function rightClickCell(page: Page, a1: string): Promise<void> {
  const rect = await rectOfReady(page, a1);
  await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2, { button: 'right' });
  await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
}

/** 表格获得键盘焦点后再按 Ctrl+C（真实用户也是先点一下格子再按） */
async function pressCopy(page: Page): Promise<void> {
  await clickCell(page, 'A1');
  await page.evaluate(() => {
    (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:C3');
  });
  await page.waitForTimeout(150);
  await page.keyboard.press('Control+c');
  await page.waitForTimeout(700);
}

test.describe('复制出去', () => {
  test('① Ctrl+C 写入系统剪贴板：TSV 纯文本 + 带格式的 HTML', async ({ page }) => {
    test.setTimeout(120_000);
    await bootWithFixture(page);

    await page.evaluate(async () => navigator.clipboard.writeText('SENTINEL-CtrlC'));
    await pressCopy(page);

    const clip = await readClipboard(page);
    expect(clip.text, 'Ctrl+C 必须真的写进剪贴板（曾经完全没有反应）').not.toBe('SENTINEL-CtrlC');
    expect(clip.types, '两种口味都要有：纯文本给不支持 HTML 的软件，HTML 给 Excel/WPS').toContain('text/plain');
    expect(clip.types).toContain('text/html');

    // 纯文本必须是 TSV（Tab 分列、换行分行）——Excel/WPS 就是靠这个分格的
    const firstRow = (clip.text ?? '').split('\r\n')[0] ?? (clip.text ?? '').split('\n')[0];
    expect(firstRow.split('\t').length, '第一行应有 3 列（A1:C3）').toBe(3);
    expect(clip.text, '应含表格里的真实内容').toContain('样式保真样本');

    // HTML 必须把格式带出来：字体/加粗/字色/底色/边框/对齐/合并/列宽
    const html = clip.html ?? '';
    for (const marker of ['font-family', 'font-weight', 'color:', 'background', 'border', 'text-align', 'colspan']) {
      expect(html, `HTML 口味里应带出 ${marker}`).toContain(marker);
    }
    expect(html, '列宽也要带（Excel 粘过去列宽才像样）').toContain('<colgroup>');
    expect(html, '外层标记让 Excel / 在线表格认得出这是表格').toContain('<table');

    // 被拦命令里不该再出现复制命令（回归点：白名单 id 写错时这里会 +1）
    const blocked = await page.evaluate(
      () => (window as never as { __p0: { getBlockedCommands: () => Array<{ id: string; count: number }> } }).__p0.getBlockedCommands(),
    );
    expect(blocked.filter((entry) => entry.id.includes('copy')), '复制命令不得再被闸门拦下').toHaveLength(0);
  });

  test('② 右键「复制内容」：单块也带格式，标签写明"含格式"', async ({ page }) => {
    test.setTimeout(120_000);
    await bootWithFixture(page);

    await clickCell(page, 'A1');
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:C3');
    });
    await page.waitForTimeout(150);

    await rightClickCell(page, 'A1');

    const item = page.locator('[data-testid="context-menu-copy"]');
    await expect(item, '标签要如实写明带格式').toContainText('含格式');
    await writeSentinel(page, 'SENTINEL-MENU');
    await item.click();
    await page.waitForTimeout(700);

    const clip = await readClipboard(page);
    expect(clip.text, '右键复制必须真的写进剪贴板').not.toBe('SENTINEL-MENU');
    expect(clip.types, '右键复制也要带 HTML 口味').toContain('text/html');
    expect(clip.html ?? '', 'HTML 里要带格式').toContain('border');
    expect(await page.locator('[data-testid="context-menu"]').count(), '点完菜单应收起').toBe(0);
  });

  test('③ 多块选区：只走纯文本（原生复制会丢块），标签如实写明', async ({ page }) => {
    test.setTimeout(120_000);
    await bootWithFixture(page);

    // Ctrl+点选造两块（Univer 的选区模型原生支持，见 selection-multi.spec.ts）。
    // 刻意避开 A1：它是合并格（A1:F1），点它会选中整块合并区，做"两块"的前提就不干净了。
    await clickCell(page, 'B2');
    await clickCell(page, 'D4', ['Control']);
    const model = await page.evaluate(() => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => unknown } }).__p0.getActiveSheet() as {
        getSelection?: () => { getActiveRangeList: () => Array<{ getA1Notation: () => string }> } | null;
      } | null;
      return (sheet?.getSelection?.()?.getActiveRangeList() ?? []).map((range) => range.getA1Notation());
    });
    expect(model, '前提：必须真的选中两块，否则"多块走纯文本"没被验到').toEqual(['B2', 'D4']);

    await rightClickCell(page, 'B2');

    const item = page.locator('[data-testid="context-menu-copy"]');
    await expect(item, '多块要写明块数').toContainText('块');
    await expect(item, '多块不带格式这件事必须写在标签里，不能让人以为格式会跟过去').toContainText('纯文本');

    await writeSentinel(page, 'SENTINEL-MULTI');
    await item.click();
    await page.waitForTimeout(700);

    const clip = await readClipboard(page);
    expect(clip.text, '多块复制应写出内容').not.toBe('SENTINEL-MULTI');
    expect(clip.types, '多块只写纯文本').toEqual(['text/plain']);
    expect(clip.text ?? '', '多块之间空一行分隔（与 Excel 多区域复制的习惯一致）').toContain('\r\n\r\n');
  });

  test('④ 没有 Clipboard API 时（http 局域网）退回 execCommand，而不是静默失败', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);

    const result = await page.evaluate(async () => {
      // 模拟不安全上下文：把 navigator.clipboard 遮成 undefined（它平时挂在 Navigator.prototype 上）
      Object.defineProperty(navigator, 'clipboard', { value: undefined, configurable: true, writable: true });
      const hadClipboard = typeof navigator.clipboard !== 'undefined';
      // 路径必须走**变量**：写成字符串字面量时 tsc 会去静态解析这个 URL（解析不到 → TS2307），
      // 与 helpers.ts 里 clearSession 的写法保持一致。
      const modulePath = '/src/shell/clipboard-write.ts';
      const mod = (await import(/* @vite-ignore */ modulePath)) as {
        writeClipboardText: (text: string) => Promise<boolean>;
      };
      const ok = await mod.writeClipboardText('FALLBACK-落盘');
      // 撤掉遮罩，恢复浏览器真实实现，好把剪贴板读回来验证
      delete (navigator as { clipboard?: unknown }).clipboard;
      const readBack = typeof navigator.clipboard?.readText === 'function' ? await navigator.clipboard.readText() : null;
      return { hadClipboard, ok, readBack, restored: typeof navigator.clipboard !== 'undefined' };
    });

    expect(result.hadClipboard, '前提：遮罩应生效').toBe(false);
    expect(result.ok, '降级路径也必须返回成功').toBe(true);
    expect(result.restored, '前提：遮罩应能撤掉（否则读不回来）').toBe(true);
    expect(result.readBack, '降级路径写的内容必须真的进了系统剪贴板').toBe('FALLBACK-落盘');
  });
});

test.describe('粘进来', () => {
  test('⑤ 从外部粘带格式的 HTML：值写进去，但目标格格式一格不变、也不引入富文本', async ({ page }) => {
    test.setTimeout(120_000);
    await bootWithFixture(page);

    // 目标选 A2：它本身有样式（蓝底 + 白字 + 四边框），"格式被覆盖"才看得出来；且它不是合并格
    const readCell = () =>
      page.evaluate(() => {
        const range = (
          window as never as {
            __p0: {
              getActiveSheet: () => {
                getRange: (a1: string) => {
                  getValue: (rt?: boolean) => unknown;
                  getDisplayValue: () => unknown;
                  getCellStyleData: (type?: string) => unknown;
                };
              };
            };
          }
        ).__p0
          .getActiveSheet()
          .getRange('A2');
        return {
          display: String(range.getDisplayValue() ?? ''),
          style: JSON.stringify(range.getCellStyleData('cell') ?? null),
          isRichText: typeof range.getValue(true) === 'object' && range.getValue(true) !== null,
        };
      });

    const before = await readCell();
    expect(before.style, '前提：A2 必须是有样式的格子').toContain('bg');

    // 模拟"从网页/Excel 复制带格式内容"：红底、粗体、下划线、蓝字
    await page.evaluate(async () => {
      const html =
        '<table><tbody><tr><td style="background-color:#ff0000;font-weight:bold;text-decoration:underline;color:#0000ff">外部粗体下划线</td></tr></tbody></table>';
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob(['外部粗体下划线'], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        }),
      ]);
    });

    await clickCell(page, 'A2');
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A2');
    });
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(1200);

    const after = await readCell();
    expect(after.display, '前提：值必须真的粘进去了（否则下面的"格式没变"是空断言）').toBe('外部粗体下划线');
    expect(after.style, '粘贴带来的底色/粗体/边框必须被剥掉，目标格格式一格不变').toBe(before.style);
    expect(after.isRichText, '不得因为粘贴而引入单元格内富文本').toBe(false);
  });
});
