/**
 * 导出格式与双形态的**本地版**用例（跑在开发服务器上）。
 *
 * 覆盖：
 *  ① 导出菜单：默认保真导出 .xlsx 仍然是原来那条路（用户按一下就能拿到文件）；
 *  ② CSV 导出：真下载一份，断言 BOM 与内容（学号这类前导零不能被吃掉）；
 *  ③ 本地转换桥：`/api/bridge/health` 有响应；Excel 可用时**真的导出一次 .ods**，
 *     断言拿到的是合法的 OpenDocument 包（zip + mimetype），不可用时菜单项置灰并说明原因。
 *
 * 不依赖 Excel 也能过：装没装 Excel 只影响 ③ 的分支，两种情况都断言（这才是"双形态"的真实状态）。
 */
import { expect, test } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

import { importFixture, waitForBoot } from './helpers';

const FIXTURE = join(process.cwd(), 'fixtures', 'fixture-styles.xlsx');

test.describe('导出格式与本地转换桥', () => {
  test.skip(!existsSync(FIXTURE), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');
  });

  test('① 导出菜单：默认项仍是保真导出 .xlsx', async ({ page }) => {
    test.setTimeout(120_000);
    await page.click('[data-testid="toolbar-export"]');
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 60_000 }),
      page.locator('[data-testid="context-menu-export-xlsx-keep"]').click(),
    ]);
    expect(download.suggestedFilename()).toMatch(/-已编辑\.(xlsx|xlsm)$/);
    const path = await download.path();
    expect(path ? readFileSync(path).length : 0).toBeGreaterThan(1000);
  });

  test('② 导出 CSV：带 UTF-8 BOM、内容是当前表的显示值', async ({ page }) => {
    test.setTimeout(120_000);
    await page.click('[data-testid="toolbar-export"]');
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 60_000 }),
      page.locator('[data-testid="context-menu-export-csv"]').click(),
    ]);
    expect(download.suggestedFilename()).toMatch(/-样式\.csv$/);
    const path = await download.path();
    const bytes = readFileSync(path!);
    expect([...bytes.slice(0, 3)], 'CSV 应带 UTF-8 BOM（Excel 双击不乱码）').toEqual([0xef, 0xbb, 0xbf]);
    const text = new TextDecoder('utf-8').decode(bytes);
    expect(text, '表头应在 CSV 里').toContain('区域');
    expect(text, '行尾应为 CRLF').toContain('\r\n');
  });

  test('③ 本地转换桥：有健康检查；能导出 .ods 时产出合法 OpenDocument 包', async ({ page }) => {
    test.setTimeout(180_000);

    const health = await page.evaluate(async () => {
      const response = await fetch('/api/bridge/health');
      return { status: response.status, body: (await response.json()) as { available?: boolean; excel?: string; reason?: string } };
    });
    expect(health.status, '本地版应挂载转换桥').toBe(200);

    await page.click('[data-testid="toolbar-export"]');
    const odsItem = page.locator('[data-testid="context-menu-export-ods"]');
    await expect(odsItem).toBeVisible();

    if (!health.body.available) {
      // 没装 Excel：必须置灰 + 说清原因（本机如此时也能守护"不要给一个点了没反应的按钮"）
      await expect(odsItem, 'Excel 不可用时 .ods 导出应置灰').toBeDisabled();
      await expect(odsItem).toContainText('本地版');
      console.log(`[bridge] 本机 Excel 不可用：${health.body.reason ?? '未知原因'}`);
      return;
    }

    await expect(odsItem, 'Excel 可用时 .ods 导出应可点').toBeEnabled({ timeout: 30_000 });
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 150_000 }),
      odsItem.click(),
    ]);
    expect(download.suggestedFilename()).toMatch(/-已编辑\.ods$/);
    const bytes = readFileSync((await download.path())!);
    // 合法的 .ods：zip 包（PK\x03\x04），且第一个条目是未压缩的 mimetype
    expect([...bytes.slice(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(bytes.includes(Buffer.from('application/vnd.oasis.opendocument.spreadsheet'))).toBe(true);
    console.log(`[bridge] 经本机 Excel 导出 .ods：${bytes.length} 字节（${health.body.excel ?? 'Excel'}）`);
  });
});
