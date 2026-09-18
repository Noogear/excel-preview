/**
 * **把 xlsx 拖进窗口就能打开**。
 *
 * 用户实测反馈："无法打开 2502班座位表 单座 2026.9.7.xlsx"——
 * 应用自己的文案一直写着"拖入 xlsx 或使用左侧示例开始"，但**根本没有接文件拖放**：
 * 拖进来毫无反应（浏览器甚至可能直接跳转去打开那个文件）。现在窗口级接管 drag/drop：
 * 只处理带 `Files` 的拖放、`dragover` 里 preventDefault、支持一次多个文件、非 .xlsx 给提示。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { basename } from 'node:path';

import { clearLog, fixturePath, importFixture, userFixturePath, waitForBridgeStatus, waitForBoot } from './helpers';

/**
 * 用户自己的真实表格（本机路径，不进仓库）：见 `helpers.userFixturePath()`。
 * 未配置时相关用例自动跳过 —— 仓库是公开的，不该把"某位老师的本机路径与文件名"写进代码。
 */
const USER_FILE = userFixturePath('座位表');

/** 把本地文件当成"拖进来的文件"派发到窗口（DataTransfer 带 Files） */
async function dropFiles(page: Page, files: Array<{ name: string; path: string; mime?: string }>): Promise<void> {
  const payload = files.map((file) => ({
    name: file.name,
    mime: file.mime ?? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    bytes: Array.from(readFileSync(file.path)),
  }));
  await page.evaluate((items: Array<{ name: string; mime: string; bytes: number[] }>) => {
    const dt = new DataTransfer();
    for (const item of items) {
      const blob = new Blob([new Uint8Array(item.bytes)], { type: item.mime });
      dt.items.add(new File([blob], item.name, { type: item.mime }));
    }
    const target = document.querySelector('.stage') ?? document.body;
    target.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dt }));
    target.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: dt }));
    target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt }));
  }, payload);
}

const summary = (page: Page) =>
  page.evaluate(() => (window as never as { __p0: { getImportSummary: () => { fileName?: string } | null } }).__p0.getImportSummary());

test.describe('拖入文件即打开', () => {
  test('① 拖入用户的座位表 → 真的导入并渲染出内容', async ({ page }) => {
    test.setTimeout(240_000);
    test.skip(!USER_FILE || !existsSync(USER_FILE), '未配置用户文件（USER_EXCEL 或 tests/e2e/user-fixture.local），跳过');
    const userFile = USER_FILE!;
    const userFileName = basename(userFile);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await page.waitForTimeout(500);

    await dropFiles(page, [{ name: userFileName, path: userFile }]);
    await page.waitForFunction(
      (name) =>
        (window as never as { __p0: { getImportSummary: () => { fileName?: string } | null } }).__p0.getImportSummary()?.fileName === name,
      userFileName,
      { timeout: 120_000 },
    );
    await page.waitForTimeout(600);

    const state = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getActiveSheet: () => { getRange: (a1: string) => { getDisplayValues: () => string[][] } };
          log: Array<{ kind: string }>;
        };
      }).__p0;
      const grid = hooks.getActiveSheet().getRange('C1:N22').getDisplayValues();
      const filled = grid.flat().filter((cell) => String(cell ?? '').trim() !== '').length;
      return { filled, kinds: hooks.log.map((e) => e.kind) };
    });
    expect(state.kinds, '应记录一次文件拖放').toContain('file:drop');
    expect(state.kinds, '应走完整导入流程').toContain('import:done');
    expect(state.filled, '拖入后内容应真的渲染出来').toBeGreaterThan(20);
  });

  test('② 拖入时显示"松手即打开"提示，离开后收起', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await page.waitForTimeout(400);

    // 派发 dragenter/dragover（不 drop）
    await page.evaluate(() => {
      const dt = new DataTransfer();
      dt.items.add(new File([new Uint8Array([1, 2, 3])], 'x.xlsx'));
      const target = document.querySelector('.stage') ?? document.body;
      target.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dt }));
      target.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: dt }));
    });
    await expect(page.locator('[data-testid="file-drop-overlay"]'), '拖入时应出现提示').toBeVisible();

    // 离开窗口（relatedTarget 为空）→ 收起
    await page.evaluate(() => {
      const dt = new DataTransfer();
      dt.items.add(new File([new Uint8Array([1])], 'x.xlsx'));
      window.dispatchEvent(new DragEvent('dragleave', { bubbles: true, cancelable: true, dataTransfer: dt, relatedTarget: null }));
    });
    await expect(page.locator('[data-testid="file-drop-overlay"]'), '离开后应收起提示').toHaveCount(0);
  });

  test('③ 拖入打不开的格式 → 逐类给出"为什么 + 怎么办"，且不导入', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await page.waitForTimeout(400);

    const fixture = fixturePath('fixture-styles.xlsx');
    test.skip(!existsSync(fixture), '请先运行 npm run fixtures');

    /**
     * 注意：`.xls` / `.csv` / `.ods` **现在都能打开**（自研解析后翻译成 xlsx），
     * 所以这条用例改用真正不支持的格式：`.numbers`（iWork 私有）。
     * （扩展名判定是唯一入口，字节内容不重要，借用一份真 xlsx 即可。）
     *
     * `.xlsb` 是**双形态**的边界：本地版 + 本机 Excel 时能打开（借 Excel 转 xlsx，完整保真），
     * 静态版或没装 Excel 时按"打不开"处理并说明原因 —— 两种情况都在下面同时断言。
     */
    const bridge = await waitForBridgeStatus(page);

    await dropFiles(page, [{ name: '成绩册.xlsb', path: fixture, mime: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12' }]);
    if (bridge.available) {
      // 经本机 Excel 转换要几秒（Excel COM 冷启动）：等它真的落地，别用固定 sleep 抢跑
      await page.waitForFunction(
        () =>
          (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.some(
            (entry) => entry.kind === 'import:bridge-converted',
          ),
        null,
        { timeout: 120_000 },
      );
    }
    await page.waitForTimeout(800);
    const xlsbState = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getImportSummary: () => { fileName?: string } | null;
          log: Array<{ kind: string }>;
        };
      }).__p0;
      const toasts = Array.from(document.querySelectorAll('[data-testid="toast"]')).map((n) => (n.textContent ?? '').trim());
      return { summary: hooks.getImportSummary()?.fileName ?? null, toasts, kinds: hooks.log.map((e) => e.kind) };
    });
    expect(xlsbState.kinds, '应记录这次拖放').toContain('file:drop');
    if (bridge.available) {
      expect(xlsbState.kinds, '本机 Excel 可用时应借它把 .xlsb 转成 xlsx 再打开').toContain('import:bridge-converted');
      expect(xlsbState.summary, '转换成功后应真的打开这份文件').toBe('成绩册.xlsb');
    } else {
      expect(xlsbState.kinds, '没有本机 Excel 时不该导入 .xlsb').not.toContain('import:start');
      const xlsbToast = xlsbState.toasts.join(' ');
      expect(xlsbToast, '.xlsb 要说清它需要本机 Excel（本地版）').toMatch(/Excel/);
      expect(xlsbToast, '并给出可照做的下一步').toContain('另存为 .xlsx');
      expect(xlsbToast, '要点到具体文件名').toContain('成绩册.xlsb');
    }

    // `.numbers`：两种形态都打不开。先清日志，避免把上面 .xlsb 那条导入的 import:start 算进来
    await clearLog(page);
    await dropFiles(page, [{ name: '名单.numbers', path: fixture, mime: 'application/octet-stream' }]);
    await page.waitForTimeout(800);
    const numbersState = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0;
      const toasts = Array.from(document.querySelectorAll('[data-testid="toast"]')).map((n) => (n.textContent ?? '').trim());
      return {
        toasts,
        kinds: hooks.log.map((e) => e.kind),
        summary: (window as never as { __p0: { getImportSummary: () => { fileName?: string } | null } }).__p0.getImportSummary()?.fileName ?? null,
      };
    });
    expect(numbersState.kinds, '不该导入 .numbers').not.toContain('import:start');
    const numbersToast = numbersState.toasts.join(' ');
    expect(numbersToast, '要说清是别家格式').toContain('另存为 .xlsx');
    expect(numbersToast).toContain('名单.numbers');
    expect(numbersState.summary, '.numbers 不该换掉当前工作簿').toBe(bridge.available ? '成绩册.xlsb' : null);
  });

  /**
   * 用户提问："为什么打开 xlsx 按钮名称为 xlsx，对其他格式的兼容性呢？"
   *
   * 答：能打开的不止 .xlsx。现在共九种（见 `src/importer/file-kinds.ts`）：
   * **OOXML 包**（`.xlsx/.xlsm/.xltx/.xltm`，原字节直用）、**分隔文本**（`.csv/.tsv/.txt`）、
   * **`.ods`**、**`.xls`**（后三种自研解析后翻译成规范 xlsx）。
   * 本条只钉住 `.xlsm`：宏部件**原样保留**（导出逐字节不变）、**不执行**。
   * 打不开的（.xlsb / Numbers / WPS 私有）见 ③，会给出可照做的提示。
   */
  test('④ 打开启用宏的工作簿（.xlsm）：正常预览，宏记在"已原样保留"里', async ({ page }) => {
    test.setTimeout(180_000);
    const macroFile = fixturePath('fixture-macro.xlsm');
    test.skip(!existsSync(macroFile), '请先运行 npm run fixtures');

    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await page.waitForTimeout(400);

    // 文件选择框必须认这些扩展名（按钮文案也跟着改了：现在叫「打开表格」）
    const accept = await page.locator('[data-testid="file-input"]').getAttribute('accept');
    for (const ext of ['.xlsx', '.xlsm', '.xltx', '.xltm', '.csv', '.ods', '.xls']) {
      expect(accept, `accept 应含 ${ext}`).toContain(ext);
    }
    await expect(page.locator('[data-testid="toolbar-open"]')).toHaveAttribute('title', /打开表格/);
    await expect(page.locator('[data-testid="toolbar-open"]')).toHaveAttribute('title', /\.xlsm/);

    await importFixture(page, 'fixture-macro.xlsm');
    await page.waitForTimeout(600);

    const state = await page.evaluate(() => {
      const p0 = (window as never as { __p0: Record<string, (...a: unknown[]) => unknown> }).__p0;
      const summary = p0.getImportSummary() as { fileName: string; unsupported: string[]; preserved: string[] } | null;
      const sheet = p0.getActiveSheet() as { getRange: (a1: string) => { getDisplayValues: () => string[][] } } | null;
      const log = p0.log as unknown as Array<{ kind: string }>;
      return {
        fileName: summary?.fileName ?? null,
        unsupported: summary?.unsupported.join('\n') ?? '',
        preserved: summary?.preserved.join('\n') ?? '',
        grid: sheet ? sheet.getRange('A1:B3').getDisplayValues() : null,
        tabs: (p0.getTabs() as Array<{ fileName: string }>).map((tab) => tab.fileName),
        rejected: log.some((entry) => entry.kind === 'import:reject'),
      };
    });

    expect(state.rejected, '不该被"格式不支持"挡下').toBe(false);
    expect(state.fileName).toBe('fixture-macro.xlsm');
    expect(state.tabs).toEqual(['fixture-macro.xlsm']);
    expect(state.grid?.[0]?.[0], '内容照常渲染').toContain('启用宏的工作簿');
    expect(state.unsupported, '宏不该记成"未支持"（用户会以为宏丢了）').not.toContain('宏');
    expect(state.preserved, '宏应记在"已原样保留"里').toContain('宏(VBA)');
    expect(state.preserved, '并说明不执行').toContain('不执行宏');
  });

  test('⑤ 一次拖入两个 xlsx → 各自开成标签页', async ({ page }) => {
    test.setTimeout(180_000);
    await waitForBoot(page);
    await page.waitForTimeout(400);
    const first = fixturePath('fixture-styles.xlsx');
    const second = fixturePath('fixture-numfmt.xlsx');
    test.skip(!existsSync(first) || !existsSync(second), '请先运行 npm run fixtures');

    await dropFiles(page, [
      { name: 'a-fixture-styles.xlsx', path: first },
      { name: 'b-fixture-numfmt.xlsx', path: second },
    ]);
    await page.waitForFunction(
      () => (window as never as { __p0: { getTabs: () => unknown[] } }).__p0.getTabs().length >= 2,
      null,
      { timeout: 120_000 },
    );
    const tabs = await page.evaluate(() =>
      (window as never as { __p0: { getTabs: () => Array<{ fileName: string }> } }).__p0.getTabs().map((t) => t.fileName),
    );
    expect(tabs.sort(), '两个文件都应开成标签页').toEqual(['a-fixture-styles.xlsx', 'b-fixture-numfmt.xlsx']);
    expect((await summary(page))?.fileName, '最后活动的是第二个文件').toBe('b-fixture-numfmt.xlsx');
  });
});
