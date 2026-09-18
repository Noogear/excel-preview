/**
 * 基准入口（Node 侧管线）。用 `node tools/bench-parse.mjs` 驱动：
 * 该脚本先用 esbuild 把本文件打包成普通 ESM，再用 node 跑——这样既能吃到 TS 源码，
 * 又能对**同一个进程**做 `--cpu-prof` 采样（vitest 的 worker 抓不到剖面）。
 *
 * 内存测量的坑（踩过）：在**同一个函数帧里**把变量置 null 再 GC，堆水位经常不降——
 * 死引用还留在寄存器/栈槽里，保守栈扫描会把它当活对象。所以这里刻意把"持有"与
 * "测量"拆到不同函数：函数返回、栈帧销毁之后再回收、读数，才是可信的数字。
 */
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { parseXlsx } from '../src/parser';
import { toUniverWorkbook } from '../src/importer/to-univer';
import { slimForExport } from '../src/exporter/slim-source';
import type { ParsedWorkbook } from '../src/parser/types';
import type { ImportOutcome } from '../src/importer/to-univer';
import type { ExportSource } from '../src/exporter/export-xlsx';

const FIXTURE = join(process.cwd(), 'fixtures', process.env.BENCH_FILE ?? 'bench-large.xlsx');
const BENCH_DIR = join(process.cwd(), '.bench');

/** 只有带 --expose-gc 时才能强制回收；没有就退化为"读当前水位" */
function collectGarbage(): void {
  const gc = (globalThis as { gc?: () => void }).gc;
  if (!gc) return;
  for (let i = 0; i < 3; i += 1) gc();
}

function heapMb(): number {
  return process.memoryUsage().heapUsed / 1024 / 1024;
}

function countCells(workbookData: ImportOutcome['workbookData']): number {
  let cells = 0;
  for (const sheet of Object.values(workbookData.sheets)) {
    for (const row of Object.values(sheet.cellData ?? {})) cells += Object.keys(row).length;
  }
  return cells;
}

/** 走完整管线，交出"常驻"的三件东西：适配结果 + 瘦身数据源 + 报告 */
async function runPipeline(
  bytes: Uint8Array,
): Promise<{ parseMs: number; adaptMs: number; slimMs: number; outcome: ImportOutcome; slim: ExportSource }> {
  const t0 = performance.now();
  const parsed = await parseXlsx(bytes);
  const parseMs = performance.now() - t0;

  const t1 = performance.now();
  const outcome = toUniverWorkbook(parsed, { name: 'bench', workbookId: 'bench' });
  const adaptMs = performance.now() - t1;

  // 瘦身：只留导出需要的（sheet.id + 行样式映射 + 原始字节）
  const t2 = performance.now();
  const slim = slimForExport(parsed, bytes);
  const slimMs = performance.now() - t2;

  return { parseMs, adaptMs, slimMs, outcome, slim };
}

async function main(): Promise<void> {
  const bytes = readFileSync(FIXTURE);
  const sizeMb = statSync(FIXTURE).size / 1024 / 1024;

  // ---- 内存探针模式：这个进程只干一件事——量"解析模型本身占多少" ----
  //
  // 为什么要单独开进程：同进程里先前那次解析/预热的结果会以死引用的形式留在寄存器与栈槽中，
  // 保守栈扫描会让 GC 以为它还活着（实测：同进程读数忽大忽小，100 MB 与 7 MB 乱跳）。
  // 单独进程 + 不预热，读到的才是干净数字。
  if (process.env.MEM_PROBE === '1') {
    collectGarbage();
    const base = heapMb();
    const model: ParsedWorkbook = await parseXlsx(bytes);
    collectGarbage();
    const held = heapMb();
    let cells = 0;
    for (const sheet of model.sheets) cells += sheet.cells.length;
    mkdirSync(BENCH_DIR, { recursive: true });
    writeFileSync(
      join(BENCH_DIR, 'mem.json'),
      JSON.stringify({ base: Number(base.toFixed(1)), held: Number(held.toFixed(1)), cells }),
      'utf8',
    );
    // 进程退出即释放，不做任何别的测量
    const gc = (globalThis as { gc?: () => void }).gc;
    console.log(`[mem-probe] 基线 ${base.toFixed(0)} MB → 持有解析模型 ${held.toFixed(0)} MB（${cells.toLocaleString('en-US')} 格，${gc ? '' : '未加 --expose-gc '}${((held - base) * 1024 * 1024 / Math.max(1, cells)).toFixed(0)} B/格）`);
    return;
  }

  // 预热一次，避免 JIT 冷启动污染计时（本进程不再报告模型占用，避免上面说的污染）
  await parseXlsx(bytes);

  // ---- 完整管线（计时 + 常驻水位）----
  const run = await runPipeline(bytes); // 函数已返回 → 完整模型无引用
  collectGarbage();
  const retained = heapMb();

  const cells = countCells(run.outcome.workbookData);
  let styleRows = 0;
  for (const sheet of run.slim.sheets) styleRows += sheet.rowStyles?.size ?? 0;

  // 探针进程写下的模型占用（没有就让调用方先跑一次探针）
  let modelMem: { base: number; held: number; cells: number } | null = null;
  try {
    modelMem = JSON.parse(readFileSync(join(BENCH_DIR, 'mem.json'), 'utf8'));
  } catch {
    modelMem = null;
  }

  console.log('┌─ ⑩ 性能基准：Node 侧管线（含预热，计时不含首次冷启动）──');
  console.log(`│ 文件            ${sizeMb.toFixed(1)} MB`);
  console.log(`│ 单元格          ${cells.toLocaleString('en-US')}`);
  console.log(
    `│ 解压 + 解析     ${run.parseMs.toFixed(0)} ms  (${(sizeMb / (run.parseMs / 1000)).toFixed(1)} MB/s, ${Math.round(cells / (run.parseMs / 1000)).toLocaleString('en-US')} 格/s)`,
  );
  console.log(`│ 适配 IWorkbook  ${run.adaptMs.toFixed(0)} ms  (${Math.round(cells / (run.adaptMs / 1000)).toLocaleString('en-US')} 格/s)`);
  console.log(`│ 合计            ${(run.parseMs + run.adaptMs).toFixed(0)} ms`);
  console.log('│ ── 内存（强制 GC 后的 heapUsed；不含 Uint8Array 的外部内存）──');
  if (modelMem) {
    const perCell = ((modelMem.held - modelMem.base) * 1024 * 1024) / Math.max(1, modelMem.cells);
    console.log(`│ 解析模型        ${(modelMem.held - modelMem.base).toFixed(0)} MB（独立进程实测，约 ${perCell.toFixed(0)} B/格）`);
  } else {
    console.log('│ 解析模型        （先跑一次探针进程才有数：npm run bench:parse 会自动跑）');
  }
  console.log(`│ 常驻工作集      ${retained.toFixed(0)} MB（管线跑完、完整模型已断开引用 + GC）`);
  console.log(`│ 模型瘦身        ${run.slimMs.toFixed(0)} ms；只保留 ${styleRows.toLocaleString('en-US')} 条行样式 + 原始 ${sizeMb.toFixed(1)} MB 字节`);
  console.log(`│ 解析报告        警告 ${run.outcome.report.warnings.length} 条 / 不支持特性 ${run.outcome.report.unsupported.length} 条`);
  console.log('└───────────────────────────────────────────────────────');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
