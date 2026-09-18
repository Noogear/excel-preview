/**
 * ⑩ 极致性能：Node 侧管线基准（解压 → 解析 → 适配）。
 *
 * 为什么要有这个测试：
 *  - "极致性能"不能只有结构性论据（"没用 DOMParser"），必须给出**数字**与**回归护栏**。
 *  - 这里跑的是产品真正用的同一条管线：`parseXlsx`（fflate 解压 + 索引扫描式 XML 分词）
 *    → `toUniverWorkbook`（单位/枚举映射 + 样式表 + 表格样式 + 行高估算）。
 *  - 断言的阈值故意放得很松（正常机器上约 10 倍余量）：它只拦"复杂度写错导致的雪崩"
 *    （比如某处退化成 O(n²)），不拦机器快慢，避免在 CI/慢机上变成 flaky。
 *
 * 夹具：`npm run bench:fixture`（默认 5 万行 × 20 列 = 100 万格）。夹具不存在时**跳过**而不是失败——
 * 仓库不提交大文件。
 */
import { existsSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { parseXlsx } from '../../src/parser';
import { toUniverWorkbook } from '../../src/importer/to-univer';

const FIXTURE = join(process.cwd(), 'fixtures', 'bench-large.xlsx');
const hasFixture = existsSync(FIXTURE);

/** 宽松阈值：仅用于抓复杂度回归（正常约 1–3s） */
const PARSE_BUDGET_MS = Number(process.env.BENCH_PARSE_BUDGET_MS ?? 30_000);
const ADAPT_BUDGET_MS = Number(process.env.BENCH_ADAPT_BUDGET_MS ?? 30_000);

describe('⑩ 性能：解析/适配管线基准', () => {
  it.skipIf(!hasFixture)('大文件：解析 + 适配的吞吐与内存（打印指标）', async () => {
    const bytes = readFileSync(FIXTURE);
    const sizeMb = statSync(FIXTURE).size / 1024 / 1024;

    // 解压（fflate）单独计时：parseXlsx 内部含解压，这里先做一次冷读
    const t0 = performance.now();
    const parsed = await parseXlsx(bytes);
    const parseMs = performance.now() - t0;
    const heapAfterParse = process.memoryUsage().heapUsed / 1024 / 1024;

    const t1 = performance.now();
    const outcome = toUniverWorkbook(parsed, { name: 'bench-large.xlsx', workbookId: 'bench' });
    const adaptMs = performance.now() - t1;
    const heapAfterAdapt = process.memoryUsage().heapUsed / 1024 / 1024;

    const sheet = outcome.workbookData.sheets[Object.keys(outcome.workbookData.sheets)[0]];
    let cells = 0;
    let nonEmpty = 0;
    for (const row of Object.values(sheet.cellData ?? {}) as Array<Record<string, { v?: unknown } | undefined>>) {
      for (const cell of Object.values(row)) {
        cells += 1;
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') nonEmpty += 1;
      }
    }

    const lines = [
      '',
      '┌─ ⑩ 性能基准：Node 侧管线 ────────────────────────────────',
      `│ 文件            ${sizeMb.toFixed(1)} MB（fixtures/bench-large.xlsx）`,
      `│ 工作表          ${parsed.sheets.length} 个；单元格 ${cells.toLocaleString('en-US')}（非空 ${nonEmpty.toLocaleString('en-US')}）`,
      `│ 解压 + 解析     ${parseMs.toFixed(0)} ms   (${(sizeMb / (parseMs / 1000)).toFixed(1)} MB/s, ${Math.round(cells / (parseMs / 1000)).toLocaleString('en-US')} 格/s)`,
      `│ 适配 IWorkbook  ${adaptMs.toFixed(0)} ms   (${Math.round(cells / (adaptMs / 1000)).toLocaleString('en-US')} 格/s)`,
      `│ 合计            ${(parseMs + adaptMs).toFixed(0)} ms`,
      `│ 堆内存          解析后 ${heapAfterParse.toFixed(0)} MB → 适配后 ${heapAfterAdapt.toFixed(0)} MB`,
      `└──────────────────────────────────────────────────────────`,
      '',
    ];
    // eslint-disable-next-line no-console
    console.log(lines.join('\n'));

    expect(cells, '应真的解析出单元格（否则基准没意义）').toBeGreaterThan(100_000);
    expect(parseMs, `解析耗时应 < ${PARSE_BUDGET_MS}ms`).toBeLessThan(PARSE_BUDGET_MS);
    expect(adaptMs, `适配耗时应 < ${ADAPT_BUDGET_MS}ms`).toBeLessThan(ADAPT_BUDGET_MS);
  }, 300_000);

  it.skipIf(!hasFixture)('重复解析不出现内存/耗时雪崩（跑 3 次看趋势）', async () => {
    const bytes = readFileSync(FIXTURE);
    const timings: number[] = [];
    for (let i = 0; i < 3; i += 1) {
      const t = performance.now();
      await parseXlsx(bytes);
      timings.push(performance.now() - t);
    }
    // eslint-disable-next-line no-console
    console.log(`│ 连续 3 次解析：${timings.map((v) => `${v.toFixed(0)}ms`).join(' / ')}`);
    const first = timings[0];
    const last = timings[timings.length - 1];
    expect(last, '第 3 次不应比第 1 次慢 3 倍以上（排除累积型泄漏）').toBeLessThan(first * 3 + 1000);
  }, 300_000);
});
