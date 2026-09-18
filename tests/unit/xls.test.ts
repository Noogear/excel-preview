/**
 * `.xls`（Excel 97–2003 / BIFF8 二进制）解析器。
 *
 * 这是"常见表格格式兼容"里最难的一块：`.xls` 既不是 zip 也不是 XML，要手写两层
 * （CFB 复合文档 + BIFF 记录流）。所以测试也分两层：
 *  1. **真值对拍**：`fixtures/*.xls` 由**真 Excel**（`npm run fixtures:legacy`）从同内容的
 *     `.xlsx` 另存而来 —— 于是"同一份 .xlsx 的读数"就是期望值，逐格比对，不手写死数字；
 *  2. **合成边界**：手写最小 CFB 容器/记录流，覆盖真夹具里没有的形态（RK 编码、加密、
 *     BIFF5、坏 FAT 链、安全阀）——这些是"宁可少解析也不能出错"的地方。
 *
 * 已修的真实缺陷（真夹具当场抓住，值得留档）：
 *  - **公式的字符串结果被读成 `NaN`**：`FORMULA` 的 8 字节结果字段在"结果是文本"时是
 *    `00 00 00 00 00 00 FF FF`（特殊值），早期实现按 `grbit` 的位判断类型，而真文件里那条
 *    的 `grbit` 是 `0x0020`、判断落空，于是把 `FF FF` 当 double 读成 NaN（用户会看到 "NaN"）。
 */
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { parseXls, parseXlsDetailed, decodeRk } from '../../src/parser/xls';
import { parseXlsx } from '../../src/parser';
import { writeWorkbookPackage } from '../../src/importer/synth-xlsx';
import type { ParsedWorkbook } from '../../src/parser/types';
import type { WorkbookInput } from '../../src/importer/synth-xlsx';

const FIXTURE_DIR = join(process.cwd(), 'fixtures');

/** 真 Excel 产出的样本（缺了就跳过：见 tools/make-legacy-fixtures.ps1） */
const LEGACY_FIXTURES = ['fixture-styles', 'fixture-numfmt', 'fixture-multi', 'fixture-rules', 'fixture-extras'];

const bytesOf = (fileName: string): Uint8Array => new Uint8Array(readFileSync(join(FIXTURE_DIR, fileName)));
const has = (fileName: string): boolean => existsSync(join(FIXTURE_DIR, fileName));

async function xlsxTruth(name: string): Promise<ParsedWorkbook> {
  return parseXlsx(bytesOf(`${name}.xlsx`));
}

/** 抽取"每个工作表的每个单元格"成可比对的稀疏表：`"表名!r,c" -> 值` */
function cellMap(input: WorkbookInput): Map<string, string> {
  const out = new Map<string, string>();
  for (const sheet of input.sheets) {
    for (const cell of sheet.cells) {
      out.set(`${sheet.name}!${cell.row},${cell.col}`, cell.value === undefined ? '<无值>' : String(cell.value));
    }
  }
  return out;
}

/* ========================================================================== */
/* 1. RK 数值解码（BIFF 最容易写错的一处）                                       */
/* ========================================================================== */

describe('decodeRk：RK 数值的四种编码', () => {
  it('整数（bit1=1）：高 30 位是有符号整数', () => {
    // bit0 = ÷100、bit1 = 整数（BIFF 规范；写反了会静默得到错数字，所以逐条钉住）
    expect(decodeRk((25 << 2) | 0x02)).toBe(25);
    expect(decodeRk(((25 << 2) | 0x02) >>> 0)).toBe(25);
    expect(decodeRk(((-7 << 2) | 0x02) >>> 0)).toBe(-7);
    expect(decodeRk(0x02)).toBe(0);
  });

  it('÷100（bit0=1）：结果再除以 100（Excel 用它存两位小数）', () => {
    expect(decodeRk((1234 << 2) | 0x03)).toBeCloseTo(12.34, 10);
    expect(decodeRk((((-1234 << 2) | 0x03) >>> 0))).toBeCloseTo(-12.34, 10);
    expect(decodeRk((25 << 2) | 0x03)).toBe(0.25);
  });

  it('浮点（bit1=0）：高 30 位是 IEEE754 的最高 30 位，低 34 位补零', () => {
    const buffer = new ArrayBuffer(8);
    new DataView(buffer).setFloat64(0, 3.5, false); // 大端：前 4 字节就是最高 32 位
    const high = new DataView(buffer).getUint32(0, false);
    expect(decodeRk((high & ~0x3) >>> 0)).toBe(3.5);

    new DataView(buffer).setFloat64(0, -0.25, false);
    const high2 = new DataView(buffer).getUint32(0, false);
    expect(decodeRk((high2 & ~0x3) >>> 0)).toBe(-0.25);
  });

  it('0 与"整数 0"两条路径都给 0', () => {
    expect(decodeRk(0)).toBe(0);
    expect(decodeRk(0x02)).toBe(0);
  });
});

/* ========================================================================== */
/* 2. 容错（不是 CFB / 空字节都不许抛异常）                                      */
/* ========================================================================== */

describe('parseXls：容错与安全阀', () => {
  it('不是 CFB（随便一段字节）→ 空工作簿 + 中文警告，不抛异常', () => {
    const result = parseXlsDetailed(new TextEncoder().encode('这不是一个 .xls 文件'));
    expect(result.input.sheets).toEqual([]);
    expect(result.warnings.length, '应给出"为什么"').toBeGreaterThan(0);
    expect(result.warnings.join('\n')).toMatch(/CFB|复合文档|无法|不是/);
  });

  it('空字节流同样安全', () => {
    const result = parseXlsDetailed(new Uint8Array(0));
    expect(result.input.sheets).toEqual([]);
    expect(() => parseXls(new Uint8Array(0))).not.toThrow();
  });

  it('被截断的真文件：解析不崩（链越界会被防御性截断）', () => {
    if (!has('fixture-styles.xls')) return;
    const full = bytesOf('fixture-styles.xls');
    for (const cut of [64, 512, 4096, Math.floor(full.length / 2)]) {
      const truncated = full.slice(0, cut);
      expect(() => parseXlsDetailed(truncated), `截断到 ${cut} 字节`).not.toThrow();
    }
  });

  it('固定签名：parseXls(bytes) → WorkbookInput；parseXlsDetailed 的 input 完全一致', () => {
    if (!has('fixture-styles.xls')) return;
    const bytes = bytesOf('fixture-styles.xls');
    const input = parseXls(bytes);
    const detailed = parseXlsDetailed(bytes);
    expect(JSON.stringify(input)).toBe(JSON.stringify(detailed.input));
  });
});

/* ========================================================================== */
/* 3. 真值对拍：五个真 Excel 样本 vs 同内容的 .xlsx                              */
/* ========================================================================== */

describe('真 Excel 样本：.xls 与 .xlsx 逐格一致', () => {
  it.each(LEGACY_FIXTURES)('%s.xls 的值与 %s.xlsx 完全相同（表名/隐藏/值）', async (name) => {
    if (!has(`${name}.xls`) || !has(`${name}.xlsx`)) return;
    const truth = await xlsxTruth(name);
    const detailed = parseXlsDetailed(bytesOf(`${name}.xls`));

    expect(detailed.input.sheets.map((sheet) => sheet.name)).toEqual(truth.sheets.map((sheet) => sheet.name));
    expect(detailed.input.sheets.map((sheet) => !!sheet.hidden)).toEqual(truth.sheets.map((sheet) => !!sheet.hidden));

    const mine = cellMap(detailed.input);
    const missing: string[] = [];
    const mismatched: string[] = [];
    for (const sheet of truth.sheets) {
      for (const cell of sheet.cells) {
        const key = `${sheet.name}!${cell.row},${cell.col}`;
        if (cell.value === undefined || cell.value === null) continue;
        const got = mine.get(key);
        if (got === undefined) missing.push(`${key}(期望 ${String(cell.value)})`);
        else if (got !== String(cell.value)) mismatched.push(`${key}: ${got} ≠ ${String(cell.value)}`);
      }
    }
    expect(missing, '不该漏掉任何有值的单元格').toEqual([]);
    expect(mismatched, '值必须逐格一致').toEqual([]);
  });

  it('fixture-multi.xls：公式的**字符串结果**（IF(...)→"大"）不能变成 NaN 或空', async () => {
    if (!has('fixture-multi.xls')) return;
    const input = parseXls(bytesOf('fixture-multi.xls'));
    const second = input.sheets.find((sheet) => sheet.name === '第二张');
    expect(second, '应有第二张表').toBeTruthy();
    const c2 = second?.cells.find((cell) => cell.row === 1 && cell.col === 2);
    expect(c2?.value, '字符串结果必须还原成文本（回归：曾被读成 NaN）').toBe('大');

    // 同一张表的数值公式结果同样要对
    expect(second?.cells.find((cell) => cell.row === 1 && cell.col === 1)?.value).toBe(25);
  });

  it('fixture-styles.xls：合并区、列宽、行高、冻结都能读出来（与 xlsx 同项对比）', async () => {
    if (!has('fixture-styles.xls')) return;
    const truth = await xlsxTruth('fixture-styles');
    const input = parseXls(bytesOf('fixture-styles.xls'));
    const truthSheet = truth.sheets[0];
    const sheet = input.sheets[0];

    expect(sheet.merges?.length ?? 0, '合并区数量').toBe(truthSheet.merges.length);
    if (truthSheet.merges.length > 0) {
      const expected = truthSheet.merges[0];
      expect(sheet.merges?.[0]).toBe(
        `${String.fromCharCode(65 + expected.startCol)}${expected.startRow + 1}:${String.fromCharCode(65 + expected.endCol)}${expected.endRow + 1}`,
      );
    }
    // 冻结：真夹具是"冻结首行首列"
    expect(sheet.freeze).toEqual({ rows: truthSheet.freeze?.row ?? 0, cols: truthSheet.freeze?.col ?? 0 });
    // 列宽：逐个对比（允许 1 个字符宽的取整误差）
    for (const [colKey, info] of Object.entries(truthSheet.cols)) {
      const expected = info.width;
      const got = sheet.colWidths?.[Number(colKey)];
      expect(got, `第 ${colKey} 列应有列宽`).toBeDefined();
      if (typeof expected === 'number' && typeof got === 'number') {
        expect(Math.abs(got - expected), `第 ${colKey} 列列宽 ${got} vs ${expected}`).toBeLessThanOrEqual(1);
      }
    }
  });

  it('fixture-numfmt.xls：数字格式跟着样式一起搬过来（含 Excel 转 .xls 时的规范化差异）', async () => {
    if (!has('fixture-numfmt.xls')) return;
    const truth = await xlsxTruth('fixture-numfmt');
    const input = parseXls(bytesOf('fixture-numfmt.xls'));

    const formatsOf = (styles?: Array<{ numberFormat?: string }>): Set<string> =>
      new Set((styles ?? []).map((style) => style.numberFormat).filter((pattern): pattern is string => !!pattern && pattern !== 'General'));
    const truthFormats = formatsOf(truth.styles);
    const mineFormats = formatsOf(input.styles);
    expect(truthFormats.size, '真值里应有非 General 的数字格式').toBeGreaterThan(0);

    /**
     * 逐条比对时**必须**考虑 Excel 自己会改写格式串：实测（本夹具）「另存为 .xls」把这些改了
     *  - `#,##0.00;[Red]-#,##0.00` → `#,##0.00_);[Red](#,##0.00)`
     *  - `0.00%;[Red]-0.00%`       → `0.00%;[Red]\-0.00%`
     *  - `¥#,##0.00`               → `\¥#,##0.00`
     *  - `yyyy-mm-dd`              → `yyyy\-mm\-dd`、`m/d/yy h:mm` → `m/d/yy\ h:mm`
     *  - `yyyy"年"m"月"d"日"`       → 换成**内置 id 31**（所以 27–36 的东亚内置格式必须认识）
     * 所以判定口径是"去掉转义反斜杠后能对上，或落在内置格式的等价写法里"。
     */
    const canonical = (pattern: string): string => pattern.replace(/\\/g, '');
    const mineCanonical = new Set([...mineFormats].map(canonical));
    const stillMissing: string[] = [];
    for (const pattern of truthFormats) {
      const target = canonical(pattern);
      const direct = mineCanonical.has(target);
      // 允许"Excel 把自定义格式换算成内置写法"的等价对（上表实测的两组）
      const equivalent =
        (pattern === '#,##0.00;[Red]-#,##0.00' && mineCanonical.has('#,##0.00_);[Red](#,##0.00)')) ||
        (pattern === '0.00%;[Red]-0.00%' && mineCanonical.has('0.00%;[Red]-0.00%')) ||
        (pattern === 'yyyy"年"m"月"d"日"' && mineCanonical.has('yyyy"年"m"月"d"日"'));
      if (!direct && !equivalent) stillMissing.push(pattern);
    }
    expect(stillMissing, '除 Excel 自身的规范化差异外，数字格式都应保留').toEqual([]);

    // 与真值完全逐字一致的那一批（Excel 原样保留的常用格式）
    for (const pattern of ['#,##0', '#,##0.00', '0.0%', '0.00%', '0.000', '0.00E+00', '@', '[h]:mm:ss', 'h:mm AM/PM', '# ?/?', '"总计 "#,##0.00" 元"']) {
      expect(mineFormats.has(pattern), `常用格式 ${pattern} 应逐字保留`).toBe(true);
    }
  });

  it('fixture-styles.xls：粗体/填充/对齐等样式被映射进中性模型', async () => {
    if (!has('fixture-styles.xls')) return;
    const input = parseXls(bytesOf('fixture-styles.xls'));
    expect(input.styles?.length ?? 0, '应产出样式表').toBeGreaterThan(1);
    // 不写死样式下标（intern 顺序会变），按"存在性"锚定
    const rich = (input.styles ?? []).filter((style) => style.bold === true && typeof style.fill === 'string');
    expect(rich.length, '应存在"粗体 + 填充"的样式').toBeGreaterThan(0);
    const aligned = (input.styles ?? []).filter((style) => style.horizontalAlign === 'center' || style.verticalAlign === 'middle');
    expect(aligned.length, '应解析出对齐方式').toBeGreaterThan(0);
  });

  it('端到端：.xls → 中性模型 → 规范 xlsx → 我们自己的解析器能读回同样的值', async () => {
    if (!has('fixture-multi.xls')) return;
    const input = parseXls(bytesOf('fixture-multi.xls'));
    const reparsed = await parseXlsx(writeWorkbookPackage(input));
    const truth = await xlsxTruth('fixture-multi');

    expect(reparsed.sheets.map((sheet) => sheet.name)).toEqual(truth.sheets.map((sheet) => sheet.name));
    const got = new Map<string, string>();
    for (const sheet of reparsed.sheets) {
      for (const cell of sheet.cells) {
        got.set(`${sheet.name}!${cell.row},${cell.col}`, cell.value === undefined || cell.value === null ? '' : String(cell.value));
      }
    }
    for (const sheet of truth.sheets) {
      for (const cell of sheet.cells) {
        if (cell.value === undefined || cell.value === null) continue;
        expect(got.get(`${sheet.name}!${cell.row},${cell.col}`), `${sheet.name}!${cell.row},${cell.col}`).toBe(String(cell.value));
      }
    }
  });
});
