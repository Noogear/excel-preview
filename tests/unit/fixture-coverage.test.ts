/**
 * 用真实生成的样本跑一遍解析器 + 适配层，输出"覆盖度报告"。
 *
 * 这是 P0-② 的证据来源：证明自研解析链路能处理 ExcelJS 产出的真实 OOXML，
 * 并量化当前覆盖率（哪些解析到了、哪些进了 unsupported/warnings）。
 */
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { toUniverWorkbook } from '../../src/importer/to-univer';
import { parseXlsx } from '../../src/parser';

const FIXTURE_DIR = join(process.cwd(), 'fixtures');

function listFixtures(): string[] {
  try {
    // 含 .xlsm：启用宏的工作簿走的是同一条 OOXML 解析路径（见 src/importer/file-kinds.ts）
    return readdirSync(FIXTURE_DIR).filter((name) => /\.(xlsx|xlsm)$/i.test(name));
  } catch {
    return [];
  }
}

describe('fixtures 覆盖度', () => {
  const fixtures = listFixtures();

  it('至少存在一个样本（否则先运行 npm run fixtures），且每个样本都可解析并转换为 IWorkbookData（逐文件打印覆盖度报告）', async () => {
    expect(fixtures.length, 'fixtures 目录为空，请先运行 node tools/make-fixtures.mjs').toBeGreaterThan(0);

    for (const fileName of fixtures) {
      const buffer = readFileSync(join(FIXTURE_DIR, fileName));
      const parsed = await parseXlsx(new Uint8Array(buffer));

      expect(parsed.sheets.length, '至少解析出一个工作表').toBeGreaterThan(0);
      expect(parsed.raw.entries['xl/workbook.xml'], '原始 zip 条目必须完整保留（导出要用）').toBeTruthy();

      const outcome = toUniverWorkbook(parsed, { name: fileName });

      const firstModel = parsed.sheets[0];
      const firstSheet = outcome.workbookData.sheets[firstModel?.id ?? ''];
      const a1Cell = firstSheet?.cellData?.[0]?.[0];
      const a1Style = a1Cell?.s ? outcome.workbookData.styles[a1Cell.s as string] : null;

      const coverage = {
        file: fileName,
        // 诊断：第一张表 A1 的最终单元格数据与样式（表格样式实体化是否落到这里）
        a1: {
          value: a1Cell?.v ?? null,
          styleId: (a1Cell?.s as string) ?? null,
          bg: (a1Style?.bg as { rgb?: string } | undefined)?.rgb ?? null,
          bold: a1Style?.bl ?? null,
        },
        sheets: parsed.sheets.map((sheet) => ({
          name: sheet.name,
          cells: sheet.cells.length,
          merges: sheet.merges.length,
          freeze: sheet.freeze ?? null,
          rows: Object.keys(sheet.rows).length,
          cols: Object.keys(sheet.cols).length,
          // ---- P1 特性计数 ----
          cf: sheet.conditionalFormats?.length ?? 0,
          dv: sheet.dataValidations?.length ?? 0,
          links: sheet.hyperlinks?.length ?? 0,
          notes: sheet.notes?.length ?? 0,
          tables: sheet.tables?.length ?? 0,
          images: sheet.images?.length ?? 0,
        })),
        styleCount: parsed.styles.length,
        dxfStyles: parsed.dxfStyles?.length ?? 0,
        nonEmptyStyles: outcome.styleIdByIndex.filter(Boolean).length,
        themeColors: parsed.themeColors?.length ?? 0,
        unsupported: parsed.report.unsupported,
        warnings: parsed.report.warnings.slice(0, 6),
      };
      // eslint-disable-next-line no-console
      console.log(`\n=== ${fileName} ===\n${JSON.stringify(coverage, null, 2)}`);

      // 工作簿必须可用：有工作表、能建 sheet 数据
      expect(outcome.workbookData.sheetOrder.length).toBe(parsed.sheets.length);
    }
  });
});
