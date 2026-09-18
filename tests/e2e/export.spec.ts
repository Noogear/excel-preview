/**
 * 外科式导出（合并自原 p1-import-fidelity.spec.ts ③ 与 p3-constraints.spec.ts 约束②）。
 *
 * 契约只有两条，但都很硬：
 *   ① 改动过的单元格值必须写回导出的 xlsx（并且导出文件能被自研解析器重新读出来）
 *   ② 未改动的部件必须**逐字节保留**——绘图/批注/媒体/样式等只读元素一个字节都不能动、也不能丢
 *
 * 之所以用 fixture-extras.xlsx：它同时带着图片（drawing/media）、批注（comments）与超链接 rels，
 * 是"只读元素原样保留"最严格的样本；白名单之外的一切部件都比对字节。
 */
import { expect, test } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';

import { parseXlsx } from '../../src/parser';
import { fixturePath, importFixture, waitForBoot } from './helpers';

test.describe('导出', () => {
  test('编辑后导出：改动写回、只读元素与其余部件逐字节保留', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-extras.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-extras.xlsx', { features: true });

    // 通过界面路径改一个单元格（A6）
    await page.evaluate(() => {
      (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0.setValue(
        5,
        0,
        '导出测试-已编辑',
      );
    });
    await page.waitForTimeout(400);

    await page.click('[data-testid="toolbar-export"]');
    const [download] = await Promise.all([
      page.waitForEvent('download'),
      page.click('[data-testid="context-menu-export-xlsx-keep"]'),
    ]);
    const downloadPath = await download.path();
    expect(downloadPath, '必须产生下载文件').toBeTruthy();
    if (!downloadPath) return;

    const exported = new Uint8Array(readFileSync(downloadPath));

    // ---- ① 改动写回，未改动的单元格原样 ----
    const reparsed = await parseXlsx(exported);
    const sheet = reparsed.sheets[0];
    expect(sheet, '导出文件必须能被重新解析').toBeTruthy();
    if (!sheet) return;

    expect(sheet.cells.find((c) => c.row === 5 && c.col === 0)?.value, '编辑后的值必须写回').toBe('导出测试-已编辑');
    expect(sheet.cells.find((c) => c.row === 0 && c.col === 0)?.value, '未编辑的单元格必须保持原值').toBe(
      '带批注的单元格',
    );

    // ---- ② 未改动部件逐字节保留 ----
    const before = unzipSync(new Uint8Array(readFileSync(fixturePath('fixture-extras.xlsx'))));
    const after = unzipSync(exported);

    // 只读元素的承载部件必须还在（否则下面的字节比对会变成"空集通过"）
    const readOnlyParts = Object.keys(before).filter(
      (name) => /^xl\/drawings\//.test(name) || /^xl\/media\//.test(name) || /^xl\/comments\d*\.xml$/.test(name),
    );
    expect(readOnlyParts.length, '样本里应带有 drawing/media/comments 部件').toBeGreaterThan(0);

    const whitelist = /^xl\/(worksheets\/sheet\d+\.xml|sharedStrings\.xml|workbook\.xml)$/;
    const changed: string[] = [];
    const missing: string[] = [];
    for (const name of Object.keys(before)) {
      if (whitelist.test(name)) continue;
      if (after[name] === undefined) {
        missing.push(name);
        continue;
      }
      if (Buffer.compare(Buffer.from(after[name]), Buffer.from(before[name])) !== 0) changed.push(name);
    }

    expect(missing, `这些部件在导出文件里不见了：${missing.join(', ')}`).toEqual([]);
    expect(changed, `这些部件不该被改动：${changed.join(', ')}`).toEqual([]);
    expect(Object.keys(after).length, 'zip 条目数不应变化').toBe(Object.keys(before).length);
  });
});
