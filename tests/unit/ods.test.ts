/**
 * `.ods`（OpenDocument Spreadsheet）解析器。
 *
 * 与 `.xls` 一样，测试分两层：
 *  1. **真值对拍**：`fixtures/*.ods` 由**真 Excel**（`npm run fixtures:legacy`）从同内容的
 *     `.xlsx` 另存而来 —— "同一份 .xlsx 的读数"就是期望值，逐格比对，不手写死数字；
 *  2. **合成 content.xml**：覆盖真夹具里没有的形态（重复行列、被覆盖单元格、文本节点、
 *     日期/时长、OpenFormula 翻译与丢弃、安全阀截断）。
 *
 * 已记录的**真实差异**（都来自真夹具，不是猜的）：
 *  - Excel 导出 ODS 会把表格"补满"到整张表大小（`table:number-columns-repeated="16383"`、
 *    百万行空重复），解析器按安全阀**截断并如实告警**：会看到"截断 1032171 行未展开"这类
 *    warning —— 这是源文件的性质，不是解析丢数据（尾部真实内容仍能读到，有用例钉住）；
 *  - **冻结窗格**：真夹具的 `.xlsx` 有冻结（首行首列），而 Excel 存出来的 ODS 里**没有**
 *    `table:table-header-rows` / `table:table-header-columns`（实测 content.xml 两者都不存在），
 *    所以 `.ods` 读回来 `freeze` 为空属于**源文件没写**，不是解析漏项；
 *  - 批注/超链接/图片等"非单元格内容"在 ODS 里另有承载元素，本解析器不导入（会告警）。
 */
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';

import { parseOds, parseOdsResult } from '../../src/parser/ods';
import { parseXlsx } from '../../src/parser';
import { writeWorkbookPackage, type WorkbookInput } from '../../src/importer/synth-xlsx';
import type { ParsedWorkbook } from '../../src/parser/types';

const FIXTURE_DIR = join(process.cwd(), 'fixtures');
const LEGACY_FIXTURES = ['fixture-styles', 'fixture-numfmt', 'fixture-multi', 'fixture-rules', 'fixture-extras'];

const has = (fileName: string): boolean => existsSync(join(FIXTURE_DIR, fileName));
const bytesOf = (fileName: string): Uint8Array => new Uint8Array(readFileSync(join(FIXTURE_DIR, fileName)));

/** 拼一份最小 ODS：只给 `table:table` 内的内容，命名空间与 office:body 骨架由这里补 */
const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'office:version="1.2"',
].join(' ');

function makeOds(tables: string, options: { styles?: string } = {}): Uint8Array {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS}>
  <office:automatic-styles>${options.styles ?? ''}</office:automatic-styles>
  <office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body>
</office:document-content>`;
  return zipSync({ 'content.xml': strToU8(content), mimetype: strToU8('application/vnd.oasis.opendocument.spreadsheet') });
}

const tableOf = (rows: string, name = '表1'): string => `<table:table table:name="${name}">${rows}</table:table>`;
const rowOf = (cells: string): string => `<table:table-row>${cells}</table:table-row>`;
const textCell = (text: string): string =>
  `<table:table-cell office:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
const numberCell = (value: number): string => `<table:table-cell office:value-type="float" office:value="${value}"/>`;

async function xlsxTruth(name: string): Promise<ParsedWorkbook> {
  return parseXlsx(bytesOf(`${name}.xlsx`));
}

/** 稀疏单元格表：`"表名!r,c" -> 字符串值` */
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
/* 1. 签名与容错                                                                */
/* ========================================================================== */

describe('parseOds：签名与容错', () => {
  it('固定签名 parseOds(bytes) → WorkbookInput；parseOdsResult 的 input 与它一致', () => {
    const bytes = makeOds(tableOf(rowOf(textCell('甲') + numberCell(1))));
    const input = parseOds(bytes);
    const detailed = parseOdsResult(bytes);
    expect(JSON.stringify(input)).toBe(JSON.stringify(detailed.input));
    expect(input.sheets[0].cells).toHaveLength(2);
  });

  it('不是 zip / 缺 content.xml → 空工作簿 + 中文警告，不抛异常', () => {
    const notZip = parseOdsResult(new TextEncoder().encode('这不是 ODS'));
    expect(notZip.input.sheets).toEqual([]);
    expect(notZip.warnings.join('\n')).toMatch(/zip|无法/);

    const noContent = parseOdsResult(zipSync({ 'styles.xml': strToU8('<x/>') }));
    expect(noContent.input.sheets).toEqual([]);
    expect(noContent.warnings.join('\n')).toContain('content.xml');
  });

  it('没有 table:name 的表给占位名（中性模型不接受空表名）', () => {
    const bytes = makeOds('<table:table><table:table-row><table:table-cell/></table:table-row></table:table>');
    const input = parseOds(bytes);
    expect(input.sheets).toHaveLength(1);
    expect(input.sheets[0].name.trim().length).toBeGreaterThan(0);
  });
});

/* ========================================================================== */
/* 2. 重复行列 / 合并 / 安全阀                                                   */
/* ========================================================================== */

describe('重复行列与合并', () => {
  it('number-columns-repeated 展开成多个格子；重复的空格子不产生单元格', () => {
    const bytes = makeOds(
      tableOf(rowOf(textCell('甲') + '<table:table-cell table:number-columns-repeated="3"/>' + textCell('乙'))),
    );
    const cells = parseOds(bytes).sheets[0].cells;
    // 重复 3 次 = 光标推进 3 格（第 1/2/3 列），所以"乙"落在第 5 列（0-based 的 4）
    expect(cells.map((cell) => `${cell.col}:${cell.value ?? ''}`)).toEqual(['0:甲', '4:乙']);
  });

  it('number-rows-repeated 按行重复（后面的行坐标不能被压扁）', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(textCell('表头'))
          + '<table:table-row table:number-rows-repeated="2">' + textCell('重复行') + '</table:table-row>'
          + rowOf(textCell('结尾')),
      ),
    );
    const cells = parseOds(bytes).sheets[0].cells;
    expect(cells.map((cell) => `${cell.row}:${cell.value}`)).toContain('0:表头');
    expect(cells.map((cell) => `${cell.row}:${cell.value}`), '结尾应在第 3 行（0-based）').toContain('3:结尾');
  });

  it('合并由"声明者"产生；covered-table-cell 只表示被覆盖、不产生内容', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(
          '<table:table-cell office:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>合并块</text:p></table:table-cell>'
            + '<table:covered-table-cell/>',
        ) + rowOf('<table:covered-table-cell/><table:covered-table-cell/>'),
      ),
    );
    const result = parseOdsResult(bytes);
    expect(result.input.sheets[0].merges, '合并区 = 声明者的范围').toEqual(['A1:B2']);
    expect(result.input.sheets[0].cells.filter((cell) => cell.value === '合并块')).toHaveLength(1);
    expect(result.input.sheets[0].cells, '被覆盖的格子没有内容，不该产出单元格').toHaveLength(1);
    expect(result.stats.merges).toBe(1);
  });

  it('安全阀：百万行的空重复被截断并告警，且**截断不吃掉后面的真实内容**', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(textCell('首行'))
          + '<table:table-row table:number-rows-repeated="1048555"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>'
          + rowOf(textCell('尾行')),
      ),
    );
    const result = parseOdsResult(bytes);
    expect(result.stats.truncated, '应记录发生过截断').toBe(true);
    expect(result.warnings.join('\n')).toMatch(/截断/);
    expect(result.input.sheets[0].cells.some((cell) => cell.value === '尾行'), '尾部内容仍要读到').toBe(true);
  });
});

/* ========================================================================== */
/* 3. 文本 / 日期 / 公式                                                        */
/* ========================================================================== */

describe('文本、日期与公式', () => {
  it('text:s 的空格数、text:line-break、text:span 都还原成普通文本', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(
          '<table:table-cell office:value-type="string"><text:p>甲<text:s text:c="3"/>乙<text:line-break/>丙</text:p></table:table-cell>'
            + '<table:table-cell office:value-type="string"><text:p><text:span>粗蓝</text:span></text:p></table:table-cell>',
        ),
      ),
    );
    const cells = parseOds(bytes).sheets[0].cells;
    expect(cells[0].value, '3 个空格 + 换行').toBe('甲   乙\n丙');
    expect(cells[1].value, 'span 里的文字不能丢').toBe('粗蓝');
  });

  it('日期与时长转成 Excel 序列号（不是文本）；PT36H 是 1.5 天，不取模', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(
          '<table:table-cell office:value-type="date" office:date-value="2025-01-01T00:00:00" office:value="45658"/>'
            + '<table:table-cell office:value-type="time" office:time-value="PT36H0M0S" office:value="1.5"/>',
        ),
      ),
    );
    const cells = parseOds(bytes).sheets[0].cells;
    expect(cells[0].value, '2025-01-01 的序列号是 45658').toBe(45658);
    expect(cells[1].value, 'PT36H = 1.5 天').toBe(1.5);
  });

  it('能可靠翻译的 OpenFormula 变成 Excel 公式；看不懂的丢公式留值并计数', () => {
    const bytes = makeOds(
      tableOf(
        rowOf(
          '<table:table-cell office:value-type="float" office:value="3" table:formula="of:=SUM([.A1:.A2])"/>'
            + '<table:table-cell office:value-type="float" office:value="9" table:formula="of:=COM.MICROSOFT.WEIRD([.A1])"/>',
        ),
      ),
    );
    const result = parseOdsResult(bytes);
    const cells = result.input.sheets[0].cells;
    expect(cells[0].formula, '简单引用与函数要翻译成 Excel 写法').toBe('SUM(A1:A2)');
    expect(cells[1].formula, '翻译不了的公式必须丢掉（绝不写错公式进用户文件）').toBeUndefined();
    expect(cells[1].value, '但缓存值要留着').toBe(9);
    expect(result.stats.formulaDropped).toBeGreaterThan(0);
  });
});

/* ========================================================================== */
/* 4. 真值对拍：五个真 Excel 样本 vs 同内容的 .xlsx                              */
/* ========================================================================== */

describe('真 Excel 样本：.ods 与 .xlsx 逐格一致', () => {
  it.each(LEGACY_FIXTURES)('%s.ods 的值与同内容的 .xlsx 逐格一致（表名/隐藏/值）', async (name) => {
    if (!has(`${name}.ods`) || !has(`${name}.xlsx`)) return;
    const truth = await xlsxTruth(name);
    const result = parseOdsResult(bytesOf(`${name}.ods`));

    expect(result.input.sheets.map((sheet) => sheet.name)).toEqual(truth.sheets.map((sheet) => sheet.name));
    expect(result.input.sheets.map((sheet) => !!sheet.hidden)).toEqual(truth.sheets.map((sheet) => !!sheet.hidden));

    const mine = cellMap(result.input);
    const missing: string[] = [];
    const mismatched: string[] = [];
    for (const sheet of truth.sheets) {
      for (const cell of sheet.cells) {
        if (cell.value === undefined || cell.value === null) continue;
        const key = `${sheet.name}!${cell.row},${cell.col}`;
        const got = mine.get(key);
        if (got === undefined) missing.push(`${key}(期望 ${String(cell.value)})`);
        else if (got !== String(cell.value)) mismatched.push(`${key}: ${got} ≠ ${String(cell.value)}`);
      }
    }
    expect(missing, '不该漏掉任何有值的单元格').toEqual([]);
    expect(mismatched, '值必须逐格一致').toEqual([]);
  });

  it('Excel 导出的 ODS 会把表补满：安全阀截断并告警（不是数据丢失）', () => {
    if (!has('fixture-styles.ods')) return;
    const result = parseOdsResult(bytesOf('fixture-styles.ods'));
    expect(result.stats.truncated, '应记录截断').toBe(true);
    expect(result.warnings.join('\n')).toMatch(/截断|重复/);
  });

  it('合并区、列宽、行高能读出来；冻结为空是"源文件没写"（实测没有 header-rows）', async () => {
    if (!has('fixture-styles.ods')) return;
    const truth = await xlsxTruth('fixture-styles');
    const input = parseOds(bytesOf('fixture-styles.ods'));
    const truthSheet = truth.sheets[0];
    const sheet = input.sheets[0];

    expect(sheet.merges?.length ?? 0, '合并区数量与真值一致').toBe(truthSheet.merges.length);
    expect(sheet.merges?.[0]).toBe('A1:F1');
    expect(Object.keys(sheet.colWidths ?? {}).length, '列宽条数与真值一致').toBe(Object.keys(truthSheet.cols).length);
    expect(
      sheet.freeze ?? null,
      'Excel 的 ODS 导出不写冻结窗格（实测无 table:table-header-rows/columns），因此这里是空',
    ).toBeNull();
  });

  it('样式被映射进中性模型（粗体/填充/对齐至少各有一例）', () => {
    if (!has('fixture-styles.ods')) return;
    const styles = parseOds(bytesOf('fixture-styles.ods')).styles ?? [];
    expect(styles.length, '应产出样式表').toBeGreaterThan(1);
    expect(styles.some((style) => style.bold === true), '应有粗体').toBe(true);
    expect(styles.some((style) => typeof style.fill === 'string'), '应有填充色').toBe(true);
    expect(
      styles.some((style) => style.horizontalAlign === 'center' || style.verticalAlign === 'middle'),
      '应有对齐方式',
    ).toBe(true);
  });

  it('端到端：.ods → 中性模型 → 规范 xlsx → 我们自己的解析器能读回同样的值', async () => {
    if (!has('fixture-multi.ods')) return;
    const input = parseOds(bytesOf('fixture-multi.ods'));
    const reparsed = await parseXlsx(writeWorkbookPackage(input));
    const truth = await xlsxTruth('fixture-multi');

    expect(reparsed.sheets.map((sheet) => sheet.name)).toEqual(truth.sheets.map((sheet) => sheet.name));
    const got = cellMap({
      sheets: reparsed.sheets.map((sheet) => ({
        name: sheet.name,
        cells: sheet.cells.map((cell) => ({ row: cell.row, col: cell.col, value: cell.value })),
      })),
    });
    for (const sheet of truth.sheets) {
      for (const cell of sheet.cells) {
        if (cell.value === undefined || cell.value === null) continue;
        expect(got.get(`${sheet.name}!${cell.row},${cell.col}`), `${sheet.name}!${cell.row},${cell.col}`).toBe(
          String(cell.value),
        );
      }
    }
  });
});
