/**
 * **常见表格格式的兼容**（用户要求："我希望能支持常见表格格式的兼容……请把其他表格格式都加上去"）。
 *
 * 覆盖面：
 *  - `.csv`（真 Excel 导出的 **GBK/ANSI** 与 **UTF-8** 两种编码）——自研 RFC4180 解析 + 编码识别；
 *  - `.ods`（OpenDocument 表格）与 `.xls`（Excel 97–2003 / BIFF8 二进制）——自研解析；
 *  - 三者都属于"**翻译成 xlsx** 再走既有全链路"，所以：预览/编辑/工作区/撤销/会话恢复全都照旧，
 *    但**导出会另存为 .xlsx**（原来那份文件一个字节都不动）——这条必须在界面上说清楚，用例也钉住。
 *  - 打不开的（`.xlsb` / Numbers / WPS 私有）给出"为什么 + 怎么办"。
 *
 * 夹具由 `npm run fixtures:legacy`（真 Excel COM 转换）生成：同内容的 .xlsx 就是**真值**，
 * 所以 ODS/XLS 的断言都是"与同一份 xlsx 的读数逐格一致"，而不是手写死数字。
 * 夹具缺失时用例自动跳过（与既有约定一致）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';

import { clearLog, fixturePath, importFixture, logKinds, readLog, waitForBridgeStatus, waitForBoot } from './helpers';

const XLSX = 'fixture-styles.xlsx';
const ODS = 'fixture-styles.ods';
const XLS = 'fixture-styles.xls';
const CSV_GBK = 'fixture-styles-gbk.csv';
const CSV_UTF8 = 'fixture-styles-utf8.csv';

interface ImportSummaryLike {
  fileName: string;
  unsupported: string[];
  warnings: string[];
  preserved: string[];
}

const summary = (page: Page): Promise<ImportSummaryLike | null> =>
  page.evaluate(() =>
    (window as never as { __p0: { getImportSummary: () => ImportSummaryLike | null } }).__p0.getImportSummary(),
  );

const cellValue = (page: Page, a1: string): Promise<string | null> =>
  page.evaluate(
    (ref) => (window as never as { __p0: { getDisplayValueByA1: (a: string) => string | null } }).__p0.getDisplayValueByA1(ref),
    a1,
  );

/** 一次读若干个格子的显示值（比对两个格式的同一片内容用） */
async function readCells(page: Page, a1List: string[]): Promise<Array<string | null>> {
  const out: Array<string | null> = [];
  for (const a1 of a1List) out.push(await cellValue(page, a1));
  return out;
}

const clickTab = async (page: Page, fileName: string): Promise<void> => {
  const tabs = await page.evaluate(() =>
    (window as never as { __p0: { getTabs: () => Array<{ id: string; fileName: string }> } }).__p0.getTabs(),
  );
  const tab = tabs.find((item) => item.fileName === fileName);
  expect(tab, `应存在标签页 ${fileName}`).toBeTruthy();
  if (!tab) return;
  await page.locator(`[data-testid="toolbar-tab-${tab.id}"]`).click();
  await page.waitForTimeout(600);
};

test.describe('多格式兼容', () => {
  test.skip(!existsSync(fixturePath(XLSX)), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
  });

  test('① 按钮叫「打开表格」、能被点、悬停是手形，accept 覆盖全部支持的格式', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await page.waitForTimeout(300);

    const open = page.locator('[data-testid="toolbar-open"]');
    await expect(open).toHaveAttribute('title', /打开表格/);
    const accept = await page.locator('[data-testid="file-input"]').getAttribute('accept');
    for (const ext of ['.xlsx', '.xlsm', '.csv', '.tsv', '.ods', '.xls']) {
      expect(accept, `accept 应含 ${ext}`).toContain(ext);
    }

    // 鼠标移到按钮上：指针必须是手形（用户明确要求；后来又反馈过一次"还是箭头"）
    const box = await open.boundingBox();
    expect(box).not.toBeNull();
    if (!box) return;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.waitForTimeout(150);
    const cursor = await page.evaluate(
      ({ x, y }: { x: number; y: number }) => {
        const el = document.elementFromPoint(x, y) as HTMLElement | null;
        return el ? getComputedStyle(el).cursor : null;
      },
      { x: box.x + box.width / 2, y: box.y + box.height / 2 },
    );
    expect(cursor, '打开表格按钮悬停应为手形').toBe('pointer');

    /**
     * 再钉两层（用户第二次反馈"还是不会变手形"，说明只赌一个来源不够稳）：
     *  ① 按钮自己（label）与内部透明 input 的 cursor 都必须是 pointer；
     *  ② 图标区域的四个角也要是手形 —— 指针落在按钮边缘时同样不能退回箭头。
     */
    const cursors = await page.evaluate(() => {
      const label = document.querySelector('[data-testid="toolbar-open"]') as HTMLElement | null;
      const input = document.querySelector('[data-testid="file-input"]') as HTMLElement | null;
      return {
        label: label ? getComputedStyle(label).cursor : null,
        input: input ? getComputedStyle(input).cursor : null,
        inline: label?.style.cursor ?? null,
      };
    });
    expect(cursors.label, 'label 的 cursor').toBe('pointer');
    expect(cursors.input, '透明 input 的 cursor（指针实际停在它上面）').toBe('pointer');
    expect(cursors.inline, '光标写在内联样式上，任何样式表顺序/缓存都盖不掉').toBe('pointer');

    for (const [dx, dy] of [
      [2, 2],
      [box.width - 2, 2],
      [2, box.height - 2],
      [box.width - 2, box.height - 2],
    ] as const) {
      const point = { x: box.x + dx, y: box.y + dy };
      await page.mouse.move(point.x, point.y);
      await page.waitForTimeout(60);
      const corner = await page.evaluate(({ x, y }: { x: number; y: number }) => {
        const el = document.elementFromPoint(x, y) as HTMLElement | null;
        return el ? getComputedStyle(el).cursor : null;
      }, point);
      expect(corner, `按钮内 (${dx},${dy}) 处也应是手形`).toBe('pointer');
    }
  });

  test('② 打开真 Excel 导出的 GBK CSV：中文不乱码，并说明"已转换、导出为 xlsx"', async ({ page }) => {
    test.setTimeout(120_000);
    test.skip(!existsSync(fixturePath(CSV_GBK)), '请先运行 npm run fixtures:legacy（需要本机 Excel）');
    await waitForBoot(page);
    await importFixture(page, CSV_GBK);
    await page.waitForTimeout(400);

    expect(await cellValue(page, 'A1'), 'GBK 中文必须正确解码').toBe('样式保真样本（P0 fixture）');
    expect(await cellValue(page, 'A2')).toBe('区域');
    const tabs = await page.evaluate(() =>
      (window as never as { __p0: { getTabs: () => Array<{ id: string; fileName: string }> } }).__p0.getTabs(),
    );
    expect(tabs.map((tab) => tab.fileName), 'CSV 应作为一个标签页打开').toEqual([CSV_GBK]);

    const info = await summary(page);
    expect(info).not.toBeNull();
    const notes = (info?.preserved ?? []).join('\n');
    expect(notes, '要说清编码识别结果').toContain('GBK');
    expect(notes, '要说清 CSV 不带格式信息').toContain('格式信息');
    expect(notes, '要说清导出会另存为 xlsx').toContain('另存为 .xlsx');
  });

  test('③ UTF-8 CSV 也能打开，且两种编码得到同一张表', async ({ page }) => {
    test.setTimeout(180_000);
    test.skip(
      !existsSync(fixturePath(CSV_UTF8)) || !existsSync(fixturePath(CSV_GBK)),
      '请先运行 npm run fixtures:legacy（需要本机 Excel）',
    );
    await waitForBoot(page);
    await importFixture(page, CSV_GBK);
    await page.waitForTimeout(300);
    const fromGbk = await readCells(page, ['A1', 'A2', 'B2', 'C2']);

    await importFixture(page, CSV_UTF8);
    await page.waitForTimeout(300);
    const fromUtf8 = await readCells(page, ['A1', 'A2', 'B2', 'C2']);
    expect(fromUtf8, '同一张表的两种编码应读到同样的内容').toEqual(fromGbk);

    const info = await summary(page);
    expect((info?.preserved ?? []).join('\n'), 'UTF-8 带 BOM 也要如实报告编码').toMatch(/utf-8/i);
  });

  test('④ .ods 与同内容的 .xlsx 逐格一致（真值对拍）', async ({ page }) => {
    test.setTimeout(180_000);
    test.skip(!existsSync(fixturePath(ODS)), '请先运行 npm run fixtures:legacy（需要本机 Excel）');
    await waitForBoot(page);

    const probes = ['A1', 'A2', 'B2', 'C2', 'D2', 'A3', 'D3', 'E3', 'A12', 'B12'];
    await importFixture(page, XLSX);
    await page.waitForTimeout(400);
    const golden = await readCells(page, probes);

    await importFixture(page, ODS);
    await page.waitForTimeout(600);
    const fromOds = await readCells(page, probes);
    expect(fromOds, '.ods 与 .xlsx 的同一片内容应一致').toEqual(golden);

    const info = await summary(page);
    expect(info?.fileName).toBe(ODS);
    const notes = (info?.preserved ?? []).join('\n');
    expect(notes, '要说明是从 .ods 转换导入的').toContain('ods');
    expect(notes, '要说明导出会另存为 .xlsx').toContain('另存为 .xlsx');
  });

  test('⑤ .xls（BIFF8）与同内容的 .xlsx 逐格一致（真值对拍）', async ({ page }) => {
    test.setTimeout(180_000);
    test.skip(!existsSync(fixturePath(XLS)), '请先运行 npm run fixtures:legacy（需要本机 Excel）');
    await waitForBoot(page);

    const probes = ['A1', 'A2', 'B2', 'C2', 'D2', 'A3', 'D3', 'E3', 'A12', 'B12'];
    await importFixture(page, XLSX);
    await page.waitForTimeout(400);
    const golden = await readCells(page, probes);

    await importFixture(page, XLS);
    await page.waitForTimeout(600);
    const fromXls = await readCells(page, probes);
    expect(fromXls, '.xls 与 .xlsx 的同一片内容应一致').toEqual(golden);

    const info = await summary(page);
    const notes = (info?.preserved ?? []).join('\n');
    expect(notes).toContain('xls');
    expect(notes, '要说明导出会另存为 .xlsx').toContain('另存为 .xlsx');
  });

  test('⑥ 转换来的标签页：导出文件名是 xxx-已编辑.xlsx（原文件不动）', async ({ page }) => {
    test.setTimeout(180_000);
    test.skip(!existsSync(fixturePath(CSV_GBK)), '请先运行 npm run fixtures:legacy（需要本机 Excel）');
    await waitForBoot(page);
    // 先开一份真 xlsx 作为对照：导出后回来看它有没有被动过
    await importFixture(page, XLSX);
    await page.waitForTimeout(300);
    await importFixture(page, CSV_GBK);
    await page.waitForTimeout(400);

    // 改一格再导出，确认导出链路对"转换来的表"同样可用
    await page.evaluate(() => (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0.setValue(5, 0, '转换后编辑'));
    await page.waitForTimeout(300);
    await page.click('[data-testid="toolbar-export"]');
    const [download] = await Promise.all([
      page.waitForEvent('download'),
      page.click('[data-testid="context-menu-export-xlsx-keep"]'),
    ]);
    expect(download.suggestedFilename(), 'csv 导入的表导出为 xlsx').toBe('fixture-styles-gbk-已编辑.xlsx');

    // 导出的是真 xlsx：换回原 xlsx 标签页确认它没被动过
    await clickTab(page, XLSX);
    expect(await cellValue(page, 'A1')).toBe('样式保真样本（P0 fixture）');
  });

  /**
   * 打不开的格式：给出"为什么 + 怎么办"。
   *
   * `.numbers`（iWork 私有）**两种形态都打不开**，所以它是不支持路径的稳定样本。
   * `.xlsb` 则是双形态的边界（本地版 + 本机 Excel 能借 Excel 转成 xlsx 打开），
   * 这里只在"桥不可用"时才要求它被拒 —— 否则等于把"能打开"当成 bug。
   */
  test('⑦ 打不开的格式（Numbers）给出"为什么 + 怎么办"且不导入；.xlsb 按形态决定', async ({ page }) => {
    test.setTimeout(120_000);
    test.skip(!existsSync(fixturePath(XLSX)), '请先运行 npm run fixtures');
    await waitForBoot(page);
    await page.waitForTimeout(300);

    const numbersBytes = Buffer.from('not a real spreadsheet');
    await page.setInputFiles('[data-testid="file-input"]', {
      name: '名单.numbers',
      mimeType: 'application/octet-stream',
      buffer: numbersBytes,
    });
    await page.waitForTimeout(600);

    const toasts = await page.evaluate(() =>
      Array.from(document.querySelectorAll('[data-testid="toast"]')).map((n) => (n.textContent ?? '').trim()),
    );
    const text = toasts.join(' ');
    expect(text, 'Numbers 要说清它不是 Excel/OpenDocument 表格').toContain('另存为 .xlsx');
    expect(await readLog(page).then((entries) => entries.some((e) => e.kind === 'import:start')), '不该真的导入').toBe(false);
    expect(await summary(page), '当前工作簿不应被替换').toBeNull();

    // `.xlsb`：桥可用（本地版 + 本机 Excel）→ 借 Excel 转 xlsx 后照常打开；不可用 → 明确拒绝并指路
    const bridge = await waitForBridgeStatus(page);
    await clearLog(page);
    await page.setInputFiles('[data-testid="file-input"]', {
      name: '成绩册.xlsb',
      mimeType: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
      // 借用一份**真 xlsx** 的字节：Excel 会按内容识别并成功另存为 xlsx，
      // 于是这条用例测的是"双形态的 .xlsb 通路"，不受"文件本身坏了"的噪声干扰。
      buffer: readFileSync(fixturePath(XLSX)),
    });
    if (bridge.available) {
      await page.waitForFunction(
        () =>
          (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.some(
            (entry) => entry.kind === 'import:bridge-converted',
          ),
        null,
        { timeout: 120_000 },
      );
      expect((await summary(page))?.fileName, '转换成功后应真的打开这份文件').toBe('成绩册.xlsb');
    } else {
      await page.waitForTimeout(800);
      const xlsbText = (
        await page.evaluate(() => Array.from(document.querySelectorAll('[data-testid="toast"]')).map((n) => (n.textContent ?? '').trim()))
      ).join(' ');
      expect(xlsbText, '.xlsb 要说明需要本机 Excel（本地版）').toMatch(/Excel/);
      expect(xlsbText, '并给出下一步').toContain('另存为 .xlsx');
      expect(await logKinds(page).then((kinds) => kinds.includes('import:start')), '不该真的导入').toBe(false);
    }
  });
});
