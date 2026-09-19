/**
 * 导出保真：**导入带特性的表、再导出，那些特性会不会被破坏？**（用户提问）
 *
 * 答案是"不会"，但要分两种情形说清楚，这里就分别钉住：
 *
 *  ① **没动过的格子/部件：一个字节都不重写。**
 *     外科式导出只重写"被编辑工作表"的 `xl/worksheets/sheetN.xml`、**追加** `sharedStrings.xml`、
 *     改 `xl/workbook.xml` 里的 `<calcPr>`。所以"一次都不改就导出"时，连工作表 XML 都不该变
 *     （只允许 workbook.xml 因为 calcPr 变）—— 这是最强的保真形态。
 *     （这也依赖"应用特性不算用户改动"：见 `dirty-tracker.ts` 的暂停记账，否则导入本身就制造了脏格。）
 *
 *  ② **只改了值的那一格所在的部件被重写，但特性元素必须原样留在里面。**
 *     条件格式 / 数据验证 / 超链接 / 绘图 / 表格是工作表 XML 里的**独立元素**（不是塞在单元格里的），
 *     而导出是"按坐标补写 `<c>`"，所以它们必须一字不改地留下来；同时其余 zip 部件（drawing/media/
 *     comments/tables 等）必须**逐字节一致**。
 *
 *  ③ 最后再让**本机 Excel 真的打开一次**导出文件（经转换桥另存为 .ods，能存成说明 Excel 解析无误、
 *     没有触发修复提示）。桥不可用时跳过（CI 上没有 Excel）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';

import { parseXlsx } from '../../src/parser';
import { fixturePath, importFixture, waitForBoot } from './helpers';

type Zip = Record<string, Uint8Array>;

const WHITELIST = /^xl\/(worksheets\/sheet\d+\.xml|sharedStrings\.xml|workbook\.xml)$/;
const ONLY_CALC_PR = /^xl\/workbook\.xml$/;

/** 走界面导出（与用户点击完全一致），返回导出的字节 */
async function exportVia(page: Page): Promise<Uint8Array> {
  await page.click('[data-testid="toolbar-export"]');
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.click('[data-testid="context-menu-export-xlsx-keep"]'),
  ]);
  const downloadPath = await download.path();
  expect(downloadPath, '必须产生下载文件').toBeTruthy();
  return new Uint8Array(readFileSync(downloadPath as string));
}

/** 逐条目比对：返回"变了"和"丢了"的部件名 */
function diffZip(before: Zip, after: Zip, allow: RegExp): { changed: string[]; missing: string[] } {
  const changed: string[] = [];
  const missing: string[] = [];
  for (const name of Object.keys(before)) {
    if (allow.test(name)) continue;
    if (after[name] === undefined) {
      missing.push(name);
      continue;
    }
    if (Buffer.compare(Buffer.from(after[name]), Buffer.from(before[name])) !== 0) changed.push(name);
  }
  return { changed, missing };
}

/** 某个 zip 部件里某个标记出现的次数（用来断言"特性元素没被弄丢"） */
function countMarker(zip: Zip, part: string, marker: string): number {
  const bytes = zip[part];
  if (!bytes) return -1;
  return new TextDecoder().decode(bytes).split(marker).length - 1;
}

/** 自研解析器给出的特性计数（导出的文件必须与原文件一致） */
function featureCounts(parsed: Awaited<ReturnType<typeof parseXlsx>>) {
  return parsed.sheets.reduce(
    (acc, sheet) => ({
      cf: acc.cf + (sheet.conditionalFormats?.length ?? 0),
      dv: acc.dv + (sheet.dataValidations?.length ?? 0),
      links: acc.links + (sheet.hyperlinks?.length ?? 0),
      notes: acc.notes + (sheet.notes?.length ?? 0),
      tables: acc.tables + (sheet.tables?.length ?? 0),
      images: acc.images + (sheet.images?.length ?? 0),
      merges: acc.merges + (sheet.merges?.length ?? 0),
    }),
    { cf: 0, dv: 0, links: 0, notes: 0, tables: 0, images: 0, merges: 0 },
  );
}

test.describe('导出保真：带特性的表', () => {
  test('① 一次都不改就导出：除 calcPr 外逐字节一致（连工作表 XML 都不重写）', async ({ page }) => {
    test.setTimeout(120_000);
    test.skip(!existsSync(fixturePath('fixture-extras.xlsx')), '请先运行 npm run fixtures');
    await waitForBoot(page);
    await importFixture(page, 'fixture-extras.xlsx', { features: true });

    const exported = await exportVia(page);
    const before = unzipSync(new Uint8Array(readFileSync(fixturePath('fixture-extras.xlsx')))) as Zip;
    const after = unzipSync(exported) as Zip;

    const { changed, missing } = diffZip(before, after, ONLY_CALC_PR);
    expect(missing, `这些部件不该消失：${missing.join(', ')}`).toEqual([]);
    expect(changed, `没改过任何内容时，这些部件不该被重写：${changed.join(', ')}`).toEqual([]);
    expect(Object.keys(after).length, 'zip 条目数不应变化').toBe(Object.keys(before).length);
    // workbook.xml 仍然只应为 calcPr 而变（内容长度可以不同，但必须仍是一个合法的工作簿部件）
    expect(countMarker(after, 'xl/workbook.xml', '<sheets>'), 'workbook.xml 结构没被破坏').toBe(1);
  });

  test('② 改了值之后：特性元素仍在工作表 XML 里，只读部件逐字节一致', async ({ page }) => {
    test.setTimeout(240_000);

    /**
     * 每个样本：编辑目标（尽量落在特性范围内，这是最严的情形）+ 必须仍在的标记。
     * 部件名与标记都是**照着夹具实际结构**写的（实测打印出来的，不是猜的）：
     *   · extras：sheet1 有 `<hyperlink`/`<drawing`，另有 comments1/drawings/media
     *   · rules：条件格式在 **sheet1**、数据验证在 **sheet2**
     *   · table：sheet1 有 `<tableParts`，实体在 xl/tables/table1|2.xml
     */
    const cases = [
      {
        file: 'fixture-extras.xlsx',
        // B1 是带超链接的格子：改它的文字，超链接元素必须还在（样式也由 s 索引保留）
        edit: { row: 0, col: 1 },
        expectParts: [{ part: 'xl/worksheets/sheet1.xml', markers: ['<hyperlink', '<drawing'] }],
        readOnly: /^xl\/(drawings\/|media\/|comments\d*\.xml$)/,
        readOnlyMin: 5,
        /** 白名单内但**没有被编辑**的部件：必须逐字节一致 */
        mustStayIdentical: ['xl/worksheets/_rels/sheet1.xml.rels', 'xl/comments1.xml'],
      },
      {
        file: 'fixture-rules.xlsx',
        edit: { row: 1, col: 1 },
        expectParts: [
          { part: 'xl/worksheets/sheet1.xml', markers: ['<conditionalFormatting'] },
          { part: 'xl/worksheets/sheet2.xml', markers: ['<dataValidation'] },
        ],
        readOnly: /^xl\/(worksheets\/sheet2\.xml|styles\.xml|theme\/)/,
        readOnlyMin: 3,
        mustStayIdentical: ['xl/worksheets/sheet2.xml', 'xl/styles.xml'],
      },
      {
        file: 'fixture-table.xlsx',
        edit: { row: 2, col: 1 },
        expectParts: [{ part: 'xl/worksheets/sheet1.xml', markers: ['<tableParts', '<tablePart'] }],
        readOnly: /^xl\/(tables\/|worksheets\/_rels\/)/,
        readOnlyMin: 3,
        mustStayIdentical: ['xl/tables/table1.xml', 'xl/tables/table2.xml', 'xl/worksheets/_rels/sheet1.xml.rels'],
      },
    ];

    for (const item of cases) {
      expect(existsSync(fixturePath(item.file)), `缺少样本 ${item.file}`).toBe(true);
      await waitForBoot(page);
      await importFixture(page, item.file, { features: true });

      const original = new Uint8Array(readFileSync(fixturePath(item.file)));
      const beforeZip = unzipSync(original) as Zip;
      const beforeParsed = await parseXlsx(original);
      const beforeCounts = featureCounts(beforeParsed);

      await page.evaluate(
        (pos: { row: number; col: number }) => {
          (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0.setValue(
            pos.row,
            pos.col,
            '导出保真-已编辑',
          );
        },
        item.edit,
      );
      await page.waitForTimeout(400);

      const exported = await exportVia(page);
      const afterZip = unzipSync(exported) as Zip;

      // ---- a) 白名单之外的一切部件逐字节一致 ----
      const { changed, missing } = diffZip(beforeZip, afterZip, WHITELIST);
      expect(missing, `[${item.file}] 这些部件不该消失：${missing.join(', ')}`).toEqual([]);
      expect(changed, `[${item.file}] 这些部件不该被改动：${changed.join(', ')}`).toEqual([]);

      // ---- b) 白名单内、但没被编辑的部件也必须逐字节一致 ----
      for (const part of item.mustStayIdentical) {
        const same = Buffer.compare(Buffer.from(afterZip[part] ?? []), Buffer.from(beforeZip[part] ?? [])) === 0;
        expect(same, `[${item.file}] ${part} 没有被编辑，必须逐字节一致`).toBe(true);
      }

      // ---- c) 真的存在只读部件（否则上面的比对会"空集通过"） ----
      const readOnlyParts = Object.keys(beforeZip).filter((name) => item.readOnly.test(name));
      expect(readOnlyParts.length, `[${item.file}] 样本应带只读部件`).toBeGreaterThanOrEqual(item.readOnlyMin);

      // ---- d) 被重写的工作表 XML 里，特性元素一个都不能少 ----
      for (const check of item.expectParts) {
        const beforeMarker = countMarker(beforeZip, check.part, check.markers[0]);
        expect(beforeMarker, `[${item.file}] 前提：${check.part} 应含 ${check.markers.join(' / ')}`).toBeGreaterThan(0);
        const beforeMarkers = check.markers.map((marker) => countMarker(beforeZip, check.part, marker));
        const afterMarkers = check.markers.map((marker) => countMarker(afterZip, check.part, marker));
        expect(afterMarkers, `[${item.file}] ${check.part} 被重写后，特性元素必须原样保留`).toEqual(beforeMarkers);
      }

      // ---- e) 语义对拍：导出的文件重新解析后，特性计数与原文件一致 ----
      const afterParsed = await parseXlsx(exported);
      expect(featureCounts(afterParsed), `[${item.file}] 导出后特性数量应与原文件一致`).toEqual(beforeCounts);

      // ---- f) 编辑的那一格确实写回去了 ----
      const editedCell = afterParsed.sheets[0]?.cells.find((c) => c.row === item.edit.row && c.col === item.edit.col);
      expect(editedCell?.value, `[${item.file}] 改过的值必须写回`).toBe('导出保真-已编辑');

      // 每轮之间清掉标签，避免互相影响（脏标签会弹确认，这里显式接受）
      page.once('dialog', (dialog) => void dialog.accept());
      const tabId = await page.evaluate(
        () => (window as never as { __p0: { getTabs: () => Array<{ id: string }> } }).__p0.getTabs()[0]?.id ?? null,
      );
      if (tabId) {
        await page.locator(`[data-testid="toolbar-tab-close-${tabId}"]`).click();
        await page.waitForTimeout(300);
      }
    }
  });

  test('③ 导出的文件本机 Excel 能直接打开（经转换桥另存为 .ods）', async ({ page }) => {
    test.setTimeout(180_000);
    test.skip(!existsSync(fixturePath('fixture-rules.xlsx')), '请先运行 npm run fixtures');
    await waitForBoot(page);
    await importFixture(page, 'fixture-rules.xlsx', { features: true });

    const bridge = await page.evaluate(async () => {
      const res = await fetch('/api/bridge/health');
      return (await res.json()) as { available: boolean; excel?: string; reason?: string };
    });
    test.skip(!bridge.available, `本机没有可用的 Excel 转换桥（${bridge.reason ?? '未知原因'}），跳过`);

    const exported = await exportVia(page);
    const outcome = await page.evaluate(async (bytes: number[]) => {
      const res = await fetch('/api/bridge/convert?to=ods', {
        method: 'POST',
        headers: { 'content-type': 'application/octet-stream' },
        body: new Uint8Array(bytes),
      });
      if (!res.ok) {
        return { ok: false, reason: await res.text() };
      }
      const data = new Uint8Array(await res.arrayBuffer());
      return { ok: true, size: data.length, magic: Array.from(data.slice(0, 2)) };
    }, Array.from(exported));

    expect(outcome.ok, `Excel 应能打开导出的文件并另存（失败原因：${'reason' in outcome ? outcome.reason : ''}）`).toBe(true);
    if (outcome.ok) {
      expect(outcome.size, '桥应回一个非空的 ods').toBeGreaterThan(0);
      expect(outcome.magic, 'ods 是 zip 包（PK 头）').toEqual([0x50, 0x4b]);
    }
  });
});
