/**
 * 外科式修补导出（`exportXlsx`）的验收测试。
 *
 * 覆盖点：
 * 1. 往返：编辑 -> 导出 -> 再解析，值与样式都要对，未编辑的格子一个都不能变
 * 2. 字节级保留：只允许 `sheetN.xml` / `sharedStrings.xml` / `workbook.xml` 变，其它条目逐字节相等
 * 3. 表格样本：`xl/tables/table*.xml` 必须原封不动
 * 4. 共享字符串复用：命中已有文本时 `<si>` 不增加
 * 5. 新单元格/新行插入：位置与列序正确
 * 6. `fullCalcOnLoad`：workbook.xml 里确实被加上
 * 7. 边界：XML 转义、非法控制字符、不改入参、sheetId 找不到时抛错
 * 8. 瘦身数据源（`slimForExport`）与完整模型的导出结果**条目级等价**
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

import { exportXlsx, type CellEdit, type SheetEdits } from '../../src/exporter/export-xlsx';
import { slimForExport } from '../../src/exporter/slim-source';
import { parseXlsx } from '../../src/parser';
import { parseSharedStrings } from '../../src/parser/worksheet';
import type { ParsedCell, ParsedWorkbook } from '../../src/parser/types';

const FIXTURE_DIR = join(process.cwd(), 'fixtures');

/** 除这三个部件外，任何条目都不允许变 */
const MUTABLE_ENTRIES = new Set([
  'xl/worksheets/sheet1.xml',
  'xl/sharedStrings.xml',
  'xl/workbook.xml',
]);

/* -------------------------------------------------------------------------- */
/* 小工具                                                                      */
/* -------------------------------------------------------------------------- */

async function loadWorkbook(fileName: string): Promise<ParsedWorkbook> {
  return parseXlsx(new Uint8Array(readFileSync(join(FIXTURE_DIR, fileName))));
}

function bytesOf(fileName: string): Uint8Array {
  return new Uint8Array(readFileSync(join(FIXTURE_DIR, fileName)));
}

function entriesOf(data: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(data);
}

function bytesEqual(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (!a || !b) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function entryText(data: Uint8Array, entry: string): string {
  const bytes = entriesOf(data)[entry];
  expect(bytes, `导出的 xlsx 里应当有 ${entry}`).toBeTruthy();
  return strFromU8(bytes as Uint8Array);
}

function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

function indexCells(cells: readonly ParsedCell[]): Map<string, ParsedCell> {
  const map = new Map<string, ParsedCell>();
  for (const cell of cells) map.set(cellKey(cell.row, cell.col), cell);
  return map;
}

function colLetters(col: number): string {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** 某个工作表里全部单元格的展示值，用于"未编辑的不能变"断言 */
function snapshot(cells: readonly ParsedCell[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const cell of cells) {
    out[`${colLetters(cell.col)}${cell.row + 1}`] = JSON.stringify([
      cell.value ?? null, cell.formula ?? null, cell.styleIndex ?? null, cell.error ?? null,
    ]);
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* 1. 往返                                                                     */
/* -------------------------------------------------------------------------- */

const FIXTURE_STYLES_EDITS: SheetEdits[] = [{
  sheetId: 'rId4', // fixture-styles 的工作表关系 id
  cells: [
    { row: 2, col: 0, value: '外科式修补' },                                  // A3（原本不存在）
    { row: 2, col: 1, value: 12345.678 },                                     // B3（原本不存在）
    { row: 2, col: 2, value: null },                                          // C3（原本不存在）
    { row: 2, col: 4, formula: 'SUM(B3:B4)' },                                // E3（已存在，s="4"）
  ],
}];

describe('exportXlsx 往返（fixture-styles.xlsx）', () => {
  it('编辑后的值 / 公式 / 样式都正确，未编辑的单元格与原文件一致，工作表结构不受影响', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const original = indexCells(source.sheets[0].cells);
    const before = snapshot(source.sheets[0].cells);
    const originalC3Style = original.get(cellKey(2, 2))?.styleIndex;

    const out = exportXlsx(source, FIXTURE_STYLES_EDITS);
    const re = await parseXlsx(out);
    const edited = indexCells(re.sheets[0].cells);

    // A3：新字符串（原文件没有这个值 -> 追加到共享字符串表末尾）
    expect(edited.get(cellKey(2, 0))?.value).toBe('外科式修补');
    // 新格子沿用所在行原有的样式（行 3 的 D3/F3 是 s="3"），避免变成"无样式异类"
    expect(edited.get(cellKey(2, 0))?.styleIndex).toBe(3);
    // B3：新数字
    expect(edited.get(cellKey(2, 1))?.value).toBe(12345.678);
    // C3：清空内容，样式按"所在行原有样式"补齐（原文件里 C3 根本不存在，没有自己的 s）
    expect(originalC3Style, '原文件 C3 不存在').toBeUndefined();
    expect(edited.has(cellKey(2, 2)), 'C3 节点应保留（清空不等于删除样式格子）').toBe(true);
    expect(edited.get(cellKey(2, 2))?.value).toBeUndefined();
    expect(edited.get(cellKey(2, 2))?.styleIndex).toBe(3); // 与同行 D3/F3 的 s="3" 一致
    // E3：新公式写进去了，且原有 s="4" 保留
    expect(edited.get(cellKey(2, 4))?.formula).toBe('SUM(B3:B4)');
    expect(edited.get(cellKey(2, 4))?.styleIndex).toBe(4);
    expect(edited.get(cellKey(2, 4))?.value).toBeUndefined(); // 不写公式缓存值

    // 未编辑的单元格：值/公式/样式一个都不能变
    const after = snapshot(re.sheets[0].cells);
    for (const [key, value] of Object.entries(before)) {
      if (key === 'D3' || key === 'E3' || key === 'F3') continue; // 行 3 被编辑（含新增 A3/B3/C3）
      expect(after[key], `未编辑的 ${key} 不应变化`).toBe(value);
    }
    // 编辑没有意外新增/丢失别的格子（A3/B3/C3 都是原文件里不存在的坐标）
    expect(Object.keys(after).filter((key) => !(key in before)).sort()).toEqual(['A3', 'B3', 'C3']);
    expect(Object.keys(after).length).toBe(Object.keys(before).length + 3);
    for (const key of ['D3', 'F3']) expect(after[key]).toBe(before[key]);

    // 工作表之外的结构（合并/冻结/列宽/行高/dimension 起点）不受影响
    expect(re.sheets[0].merges).toEqual(source.sheets[0].merges);
    expect(re.sheets[0].freeze).toEqual(source.sheets[0].freeze);
    expect(re.sheets[0].cols).toEqual(source.sheets[0].cols);
    expect(re.sheets[0].rows).toEqual(source.sheets[0].rows);
    // 编辑都在 A1:H21 之内 -> dimension 保持原样（"按需扩展"策略）
    expect(re.sheets[0].dimension).toEqual(source.sheets[0].dimension);
    expect(entryText(out, 'xl/worksheets/sheet1.xml')).toContain('<dimension ref="A1:H21"/>');
  });
});

/* -------------------------------------------------------------------------- */
/* 2. 字节级保留                                                               */
/* -------------------------------------------------------------------------- */

describe('exportXlsx 字节级保留', () => {
  it('除 sheet1/sharedStrings/workbook 之外的条目逐字节相等', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const out = exportXlsx(source, [{
      sheetId: 'rId4',
      cells: [{ row: 2, col: 0, value: '只改一个格子' }],
    }]);
    const before = entriesOf(bytesOf('fixture-styles.xlsx'));
    const after = entriesOf(out);

    expect(Object.keys(after)).toEqual(Object.keys(before)); // 条目集合与顺序都不变
    for (const name of Object.keys(before)) {
      if (MUTABLE_ENTRIES.has(name)) {
        expect(bytesEqual(after[name], before[name]), `${name} 本来就是必改部件`).toBe(false);
        continue;
      }
      expect(bytesEqual(after[name], before[name]), `${name} 必须逐字节不变`).toBe(true);
    }
    // 抽查几个"我们完全不解析"的关键部件
    expect(bytesEqual(after['xl/styles.xml'], before['xl/styles.xml'])).toBe(true);
    expect(bytesEqual(after['xl/theme/theme1.xml'], before['xl/theme/theme1.xml'])).toBe(true);
    expect(bytesEqual(after['[Content_Types].xml'], before['[Content_Types].xml'])).toBe(true);
    expect(bytesEqual(after['xl/_rels/workbook.xml.rels'], before['xl/_rels/workbook.xml.rels'])).toBe(true);

    // 空编辑列表：除了默认写入的 calcPr，其它条目字节不变
    const empty = entriesOf(exportXlsx(source, []));
    for (const name of Object.keys(before)) {
      // 默认 fullCalcOnLoad=true 会动 workbook.xml，这是契约行为
      const expected = name !== 'xl/workbook.xml';
      expect(bytesEqual(empty[name], before[name]), `${name} 在空编辑下不应变化`).toBe(expected);
    }
    // 不写 calcPr 时应当一个字节都不动
    const untouched = entriesOf(exportXlsx(source, [], { fullCalcOnLoad: false }));
    for (const name of Object.keys(before)) {
      expect(bytesEqual(untouched[name], before[name]), `${name} 必须逐字节不变`).toBe(true);
    }

    // 多表工作簿里只改被点名的那张表，其它 sheetN.xml 字节不变（值仍能回读）
    {
      const raw = bytesOf('fixture-multi.xlsx');
      const source = await parseXlsx(raw);
      const second = source.sheets[1];
      const out = exportXlsx(source, [{ sheetId: second.id, cells: [{ row: 9, col: 9, value: '第九行' }] }]);
      const before = entriesOf(raw);
      const after = entriesOf(out);

      expect(bytesEqual(after['xl/worksheets/sheet2.xml'], before['xl/worksheets/sheet2.xml'])).toBe(false);
      for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml']) {
        expect(bytesEqual(after[name], before[name]), `${name} 不应被改动`).toBe(true);
      }
      const re = await parseXlsx(out);
      expect(re.sheets).toHaveLength(4);
      expect(re.sheets[1].name).toBe(second.name);
      expect(indexCells(re.sheets[1].cells).get(cellKey(9, 9))?.value).toBe('第九行');
    }
  });

  it('fullCalcOnLoad 默认加上（保留其它 calcPr 属性、不重复写）；显式 false 时 workbook.xml 字节不变', async () => {
    const raw = bytesOf('fixture-styles.xlsx');
    const original = strFromU8(entriesOf(raw)['xl/workbook.xml'] as Uint8Array);
    expect(original).toContain('<calcPr calcId="171027"/>');
    expect(original).not.toContain('fullCalcOnLoad');

    const source = await parseXlsx(raw);
    const out = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [{ row: 2, col: 4, formula: 'SUM(B3:B4)' }],
    }]);
    const workbookXml = entryText(out, 'xl/workbook.xml');
    expect(workbookXml).toContain('fullCalcOnLoad="1"');
    expect(workbookXml).toContain('calcId="171027"'); // 其它属性保留
    expect(workbookXml.indexOf('fullCalcOnLoad')).toBeGreaterThan(workbookXml.indexOf('<calcPr'));
    expect(workbookXml).not.toContain('fullCalcOnLoad="1" fullCalcOnLoad'); // 不重复写

    // fullCalcOnLoad: false -> 不写，但仍会重建 zip（sheet1.xml 已变）
    const off = entriesOf(exportXlsx(source, [{ sheetId: source.sheets[0].id, cells: [{ row: 0, col: 0, value: 'x' }] }], {
      fullCalcOnLoad: false,
    }));
    const before = entriesOf(raw);
    expect(strFromU8(off['xl/workbook.xml'] as Uint8Array)).not.toContain('fullCalcOnLoad');
    expect(bytesEqual(off['xl/workbook.xml'], before['xl/workbook.xml'])).toBe(true);
    expect(bytesEqual(off['xl/worksheets/sheet1.xml'], before['xl/worksheets/sheet1.xml'])).toBe(false);
  });
});

/* -------------------------------------------------------------------------- */
/* 3. 表格样本：tableParts 必须完好                                             */
/* -------------------------------------------------------------------------- */

describe('exportXlsx 与 xl/tables（fixture-table.xlsx）', () => {
  it('编辑单元格后 table1.xml / table2.xml 逐字节不变', async () => {
    const raw = bytesOf('fixture-table.xlsx');
    const source = await parseXlsx(raw);
    // 前提：这是一张真的带 tableParts 的表，且解析器没有重建表格的能力
    const sheetXmlBefore = strFromU8(entriesOf(raw)['xl/worksheets/sheet1.xml'] as Uint8Array);
    expect(sheetXmlBefore).toContain('<tableParts count="2">');
    expect(entriesOf(raw)['xl/tables/table1.xml']).toBeTruthy();

    const out = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [
        { row: 1, col: 1, value: 999 },       // B2 已存在（表内数据）
        { row: 10, col: 2, value: 1234.5 },   // C11 已存在（表内总计行）
      ],
    }]);

    const before = entriesOf(raw);
    const after = entriesOf(out);
    expect(bytesEqual(after['xl/tables/table1.xml'], before['xl/tables/table1.xml'])).toBe(true);
    expect(bytesEqual(after['xl/tables/table2.xml'], before['xl/tables/table2.xml'])).toBe(true);
    expect(entryText(out, 'xl/worksheets/sheet1.xml')).toContain('<tableParts count="2">');

    const re = await parseXlsx(out);
    const cells = indexCells(re.sheets[0].cells);
    expect(cells.get(cellKey(1, 1))?.value).toBe(999);
    expect(cells.get(cellKey(10, 2))?.value).toBe(1234.5);
  });
});

/* -------------------------------------------------------------------------- */
/* 4. 共享字符串复用                                                           */
/* -------------------------------------------------------------------------- */

describe('exportXlsx 共享字符串', () => {
  it('改成文件里已存在的文本时不新增 <si>，count/uniqueCount 不变、已有序号不挪动；新文本追加到末尾并同步 count/uniqueCount', async () => {
    const raw = bytesOf('fixture-styles.xlsx');
    const source = await parseXlsx(raw);
    const sstEntry = entriesOf(raw)['xl/sharedStrings.xml'] as Uint8Array;
    const before = parseSharedStrings(strFromU8(sstEntry), () => {});
    expect(before).toContain('斜体');

    // E3 现在是富文本（P0 降级成拼接后的纯文本），改成文件里已有的纯文本"斜体"(索引 13)
    const out = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [{ row: 2, col: 4, value: '斜体' }],
    }]);

    const sstXml = entryText(out, 'xl/sharedStrings.xml');
    const after = parseSharedStrings(sstXml, () => {});
    expect(after.length, '<si> 数量不应增加').toBe(before.length);
    expect(after).toEqual(before); // 顺序也不能被重排
    expect(sstXml).toContain('count="35"');
    expect(sstXml).toContain('uniqueCount="35"');
    const re = await parseXlsx(out);
    const e3 = indexCells(re.sheets[0].cells).get(cellKey(2, 4));
    expect(e3?.value).toBe('斜体');
    expect(e3?.styleIndex).toBe(4); // s 属性保持
    // E3 的 <v> 复用了已有索引，而不是新增
    const e3Tag = /<c r="E3"[^>]*>.*?<\/c>/.exec(entryText(out, 'xl/worksheets/sheet1.xml'))?.[0] ?? '';
    expect(e3Tag).toContain(`<v>${before.indexOf('斜体')}</v>`);
    // 已存在的索引没有被挪动：A1 仍然指向索引 0
    expect(indexCells(re.sheets[0].cells).get(cellKey(0, 0))?.value).toBe(before[0]);

    // ---- 新文本：追加到末尾并同步 count/uniqueCount，已有序号不变 ----
    const appendOut = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [{ row: 2, col: 0, value: '全新文本' }, { row: 2, col: 1, value: 42 }],
    }]);

    const appendSstXml = entryText(appendOut, 'xl/sharedStrings.xml');
    const appended = parseSharedStrings(appendSstXml, () => {});
    expect(appended.length).toBe(before.length + 1);
    expect(appended.slice(0, before.length)).toEqual(before); // 旧序号一个都没动
    expect(appended[appended.length - 1]).toBe('全新文本');
    expect(appendSstXml).toContain(`count="${before.length + 1}"`);
    expect(appendSstXml).toContain(`uniqueCount="${before.length + 1}"`);

    const appendRe = await parseXlsx(appendOut);
    expect(indexCells(appendRe.sheets[0].cells).get(cellKey(2, 0))?.value).toBe('全新文本');
    expect(indexCells(appendRe.sheets[0].cells).get(cellKey(2, 1))?.value).toBe(42);

    // 空字符串值等同清空，不会写出空 <si>（sharedStrings.xml 整份字节不变）
    const rawSst = entriesOf(raw)['xl/sharedStrings.xml'];
    const cleared = exportXlsx(source, [{ sheetId: source.sheets[0].id, cells: [{ row: 2, col: 4, value: '' }] }]);
    expect(entryText(cleared, 'xl/worksheets/sheet1.xml')).toMatch(/<c r="E3" s="4"\/>/);
    expect(parseSharedStrings(entryText(cleared, 'xl/sharedStrings.xml'), () => {}).length).toBe(before.length);
    expect(bytesEqual(entriesOf(cleared)['xl/sharedStrings.xml'], rawSst)).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* 5. 新单元格 / 新行插入                                                      */
/* -------------------------------------------------------------------------- */

describe('exportXlsx 新单元格插入', () => {
  it('新值能被回读，且行内按列升序、行号顺序不变、原有行属性保留', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const out = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [
        { row: 19, col: 7, value: 'H20' },   // H20：行 20 已存在（A20），按列序插到 A20 之后
        { row: 19, col: 1, value: 'B20' },   // B20：同上，插在 A20 与 H20 之间
        { row: 19, col: 0, value: 'A20-新' }, // A20：已存在，直接改
        { row: 29, col: 2, value: 3.14 },    // C30：整行新建（行 30 原本没有）
        { row: 21, col: 0, value: true },    // A22：夹在 A21 与（无）之间的新行
      ],
    }]);

    const re = await parseXlsx(out);
    const cells = indexCells(re.sheets[0].cells);
    expect(cells.get(cellKey(19, 0))?.value).toBe('A20-新');
    expect(cells.get(cellKey(19, 1))?.value).toBe('B20');
    expect(cells.get(cellKey(19, 7))?.value).toBe('H20');
    expect(cells.get(cellKey(29, 2))?.value).toBe(3.14);
    expect(cells.get(cellKey(21, 0))?.value).toBe(true);

    // XML 层：新行按行号升序插在合适位置，行内 <c> 按列号升序
    const sheetXml = entryText(out, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<row r="22"><c r="A22" t="b"><v>1</v></c></row>');
    expect(sheetXml).toContain('<row r="30"><c r="C30">');
    // 所有 <row r="N"> 必须严格按 N 升序出现
    const rowNumbers = [...sheetXml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]));
    expect(rowNumbers).toEqual([...rowNumbers].sort((a, b) => a - b));
    expect(rowNumbers).toContain(22);
    expect(rowNumbers).toContain(30);
    // 行 20 里 A20/B20/H20 按列号升序
    const rowStart = sheetXml.indexOf('<row r="20"');
    const row20 = sheetXml.slice(rowStart, sheetXml.indexOf('</row>', rowStart));
    expect(row20.indexOf('r="A20"')).toBeGreaterThanOrEqual(0);
    expect(row20.indexOf('r="A20"')).toBeLessThan(row20.indexOf('r="B20"'));
    expect(row20.indexOf('r="B20"')).toBeLessThan(row20.indexOf('r="H20"'));
    // 行 20 原有的 ht/customHeight 等属性必须还在
    expect(row20).toContain('<row r="20" ht="48" customHeight="1"');
    // 越界了 -> dimension 按需扩展
    expect(sheetXml).toContain('<dimension ref="A1:H30"/>');
    // sheetData 之外的部件没被顺带重写
    expect(sheetXml).toContain('<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>');
    expect(sheetXml).toContain('<pageMargins left="0.7"');

    // 清空范围内已有单元格不会改 dimension
    const cleared = exportXlsx(source, [{ sheetId: source.sheets[0].id, cells: [{ row: 0, col: 0, value: null }] }]);
    expect(entryText(cleared, 'xl/worksheets/sheet1.xml')).toContain('<dimension ref="A1:H21"/>');
  });
});

/* -------------------------------------------------------------------------- */
/* 7. .xlsm（启用宏）：编辑后导出，宏部件逐字节不变                                */
/* -------------------------------------------------------------------------- */

describe('exportXlsx：启用宏的工作簿（.xlsm）', () => {
  it('只改单元格内容，vbaProject.bin 与内容类型部件一个字节都不动', async () => {
    const source = await loadWorkbook('fixture-macro.xlsm');
    const before = entriesOf(bytesOf('fixture-macro.xlsm'));

    const out = exportXlsx(
      { ...source, slim: slimForExport(source, bytesOf('fixture-macro.xlsm')) } as unknown as Parameters<typeof exportXlsx>[0],
      [{ sheetId: source.sheets[0].id, cells: [{ row: 9, col: 0, value: '宏工作簿里改一格' }] }],
    );
    const after = entriesOf(out);

    // ① 改动写进去了（字符串走 sharedStrings，重新解析一遍最直观）
    const reparsed = await parseXlsx(out);
    expect(reparsed.sheets.map((sheet) => sheet.name)).toEqual(['宏工作簿']);
    expect(
      reparsed.sheets[0]?.cells.find((cell) => cell.row === 9 && cell.col === 0)?.value,
      '编辑后的值必须写回',
    ).toBe('宏工作簿里改一格');
    expect(entryText(out, 'xl/sharedStrings.xml'), '新字符串应进共享字符串表').toContain('宏工作簿里改一格');
    // ② 宏部件与内容类型逐字节保留（改名成 .xlsx 会让 Excel 认为格式不符，所以扩展名也要保持）
    expect(bytesEqual(before['xl/vbaProject.bin'], after['xl/vbaProject.bin']), '宏部件必须逐字节不变').toBe(true);
    expect(
      bytesEqual(before['[Content_Types].xml'], after['[Content_Types].xml']),
      '内容类型部件（macroEnabled 声明）必须逐字节不变',
    ).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* 6. fullCalcOnLoad                                                           */
/* -------------------------------------------------------------------------- */

describe('exportXlsx 值类型与公式', () => {
  it('数字去 t、布尔写 t="b" 1/0、公式去掉 <f> 转义且不写缓存、删公式留值、字符串结果走 t="str"', async () => {
    const source = await loadWorkbook('fixture-multi.xlsx');
    const out = exportXlsx(source, [{
      sheetId: source.sheets[1].id, // 第二张表 -> sheet2.xml
      cells: [
        { row: 2, col: 0, value: 3.5 },                      // A3
        { row: 2, col: 1, value: false },                    // B3
        { row: 2, col: 2, formula: 'IF(A3>2,"大","小")' },   // C3（含 > 与引号）
        { row: 3, col: 0, formula: null, value: '删掉公式' }, // A4：删公式留值
      ],
    }]);
    const sheetXml = entryText(out, 'xl/worksheets/sheet2.xml');
    expect(sheetXml).toContain('<c r="A3"><v>3.5</v></c>');
    expect(sheetXml).toContain('<c r="B3" t="b"><v>0</v></c>');
    expect(sheetXml).toContain('<c r="C3"><f>IF(A3&gt;2,&quot;大&quot;,&quot;小&quot;)</f></c>');
    // A4 原有 <f>A2/4</f><v>25</v> -> 公式与旧缓存都必须消失
    const row4 = sheetXml.slice(sheetXml.indexOf('<row r="4"'), sheetXml.indexOf('</row>', sheetXml.indexOf('<row r="4"')));
    expect(row4).toMatch(/<c r="A4" t="s"><v>\d+<\/v><\/c>/);
    expect(row4).not.toContain('<f>');

    const re = await parseXlsx(out);
    const cells = indexCells(re.sheets[1].cells);
    expect(cells.get(cellKey(2, 0))?.value).toBe(3.5);
    expect(cells.get(cellKey(2, 1))?.value).toBe(false);
    expect(cells.get(cellKey(2, 2))?.formula).toBe('IF(A3>2,"大","小")');
    expect(cells.get(cellKey(2, 2))?.value).toBeUndefined();
    expect(cells.get(cellKey(3, 0))?.formula).toBeUndefined();
    expect(cells.get(cellKey(3, 0))?.value).toBe('删掉公式');

    // 公式 + 字符串结果走 t="str" 缓存（回读能拿到公式与值）
    const cached = exportXlsx(source, [{
      sheetId: source.sheets[1].id,
      cells: [{ row: 2, col: 2, formula: 'IF(A3>2,"大","小")', value: '大' }],
    }]);
    expect(entryText(cached, 'xl/worksheets/sheet2.xml'))
      .toContain('<c r="C3" t="str"><f>IF(A3&gt;2,&quot;大&quot;,&quot;小&quot;)</f><v>大</v></c>');
    const cachedCell = indexCells((await parseXlsx(cached)).sheets[1].cells).get(cellKey(2, 2));
    expect(cachedCell?.formula).toBe('IF(A3>2,"大","小")');
    expect(cachedCell?.value).toBe('大');

    // ---- 字符串做 XML 转义并剔除非法控制字符；首尾空白用 xml:space="preserve" ----
    const trickySource = await loadWorkbook('fixture-styles.xlsx');
    const tricky = 'a&b<c>d"e\'f\u0007g';
    const trickyOut = exportXlsx(trickySource, [{ sheetId: trickySource.sheets[0].id, cells: [{ row: 0, col: 6, value: tricky }] }]);

    const sstXml = entryText(trickyOut, 'xl/sharedStrings.xml');
    expect(sstXml).toContain('a&amp;b&lt;c&gt;d&quot;e&apos;f');
    expect(sstXml).not.toContain('\u0007');

    const trickyRe = await parseXlsx(trickyOut);
    expect(indexCells(trickyRe.sheets[0].cells).get(cellKey(0, 6))?.value).toBe('a&b<c>d"e\'fg');
    // 转义之后整份 XML 依然能被解析（sheetData 结构没被破坏）
    expect(trickyRe.sheets[0].cells.length).toBeGreaterThan(0);

    // 保留空格的首尾空白
    const spaced = exportXlsx(trickySource, [{ sheetId: trickySource.sheets[0].id, cells: [{ row: 3, col: 0, value: ' 前后有空格 ' }] }]);
    expect(entryText(spaced, 'xl/sharedStrings.xml')).toContain('<t xml:space="preserve"> 前后有空格 </t>');
    const spacedRe = await parseXlsx(spaced);
    expect(indexCells(spacedRe.sheets[0].cells).get(cellKey(3, 0))?.value).toBe(' 前后有空格 ');
  });
});

describe('exportXlsx 边界与错误', () => {
  it('不改动入参（raw.entries 与 sheets 都不被就地修改）；sheetId 找不到时抛出明确错误', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const entriesBefore = Object.entries(source.raw.entries).map(([name, bytes]) => [name, bytes.slice()] as const);
    const cellsBefore = snapshot(source.sheets[0].cells);
    const sheetXmlBefore = strFromU8(source.raw.entries['xl/worksheets/sheet1.xml']);
    const sstBefore = strFromU8(source.raw.entries['xl/sharedStrings.xml']);

    exportXlsx(source, [{ sheetId: source.sheets[0].id, cells: [{ row: 2, col: 0, value: '新值' }] }]);

    expect(Object.keys(source.raw.entries)).toEqual(entriesBefore.map(([name]) => name));
    for (const [name, bytes] of entriesBefore) {
      expect(bytesEqual(source.raw.entries[name], bytes), `${name} 不应被就地修改`).toBe(true);
    }
    expect(snapshot(source.sheets[0].cells)).toEqual(cellsBefore);
    expect(strFromU8(source.raw.entries['xl/worksheets/sheet1.xml'])).toBe(sheetXmlBefore);
    expect(strFromU8(source.raw.entries['xl/sharedStrings.xml'])).toBe(sstBefore);

    // sheetId 找不到时抛出明确错误
    expect(() => exportXlsx(source, [{ sheetId: 'rId99', cells: [{ row: 0, col: 0, value: 'x' }] }]))
      .toThrowError(/rId99/);
  });

  it('同一坐标重复编辑时后者生效且只产生一个 <c>；多个 SheetEdits（含重复 sheetId）能合并', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const cells: CellEdit[] = [
      { row: 9, col: 0, value: '第一次' },
      { row: 9, col: 0, value: '第二次' },
    ];
    const out = exportXlsx(source, [{ sheetId: source.sheets[0].id, cells }]);
    expect(entryText(out, 'xl/worksheets/sheet1.xml').match(/<c r="A10"/g)?.length).toBe(1);
    expect(indexCells((await parseXlsx(out)).sheets[0].cells).get(cellKey(9, 0))?.value).toBe('第二次');

    const merged = indexCells((await parseXlsx(exportXlsx(source, [
      { sheetId: source.sheets[0].id, cells: [{ row: 0, col: 1, value: '来自第一段' }] },
      { sheetId: source.sheets[0].id, cells: [{ row: 0, col: 2, value: '来自第二段' }] },
    ]))).sheets[0].cells);
    expect(merged.get(cellKey(0, 1))?.value).toBe('来自第一段');
    expect(merged.get(cellKey(0, 2))?.value).toBe('来自第二段');
  });
});

/* -------------------------------------------------------------------------- */
/* 8. 组装出来的边界场景（原样本没有的形态）                                    */
/* -------------------------------------------------------------------------- */

/** 只换掉工作表的 XML，其余条目原样带过去 */
async function withSheetXml(xml: string): Promise<ParsedWorkbook> {
  const before = entriesOf(bytesOf('fixture-styles.xlsx'));
  const entries: Record<string, Uint8Array> = { ...before, 'xl/worksheets/sheet1.xml': strToU8(xml) };
  return parseXlsx(zipSync(entries));
}

describe('exportXlsx 边界形态', () => {
  it('共享字符串表缺失时新建一个（导出结果依然可解析，其它条目字节不变）', async () => {
    const before = entriesOf(bytesOf('fixture-styles.xlsx'));
    // 去掉 sharedStrings.xml，并把工作表里所有 t="s" 的格子换成数字，保证原表自洽
    const sheetXml = strFromU8(before['xl/worksheets/sheet1.xml'] as Uint8Array)
      .replace(/<c([^>]*?)t="s"([^>]*?)><v>\d+<\/v><\/c>/g, '<c$1$2><v>1</v></c>');
    const withoutSst: Record<string, Uint8Array> = {};
    for (const [name, bytes] of Object.entries(before)) {
      if (name === 'xl/sharedStrings.xml') continue;
      withoutSst[name] = name === 'xl/worksheets/sheet1.xml' ? strToU8(sheetXml) : bytes;
    }
    const source = await parseXlsx(zipSync(withoutSst));
    expect(source.raw.entries['xl/sharedStrings.xml']).toBeUndefined();

    const out = exportXlsx(source, [{
      sheetId: source.sheets[0].id,
      cells: [{ row: 2, col: 0, value: '需要共享字符串' }],
    }]);
    const sstXml = entryText(out, 'xl/sharedStrings.xml');
    expect(sstXml).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(sstXml).toContain('<si><t>需要共享字符串</t></si>');

    const re = await parseXlsx(out);
    expect(indexCells(re.sheets[0].cells).get(cellKey(2, 0))?.value).toBe('需要共享字符串');
    // 其它条目仍然逐字节不变
    expect(bytesEqual(entriesOf(out)['xl/styles.xml'], before['xl/styles.xml'])).toBe(true);
  });

  it('自闭合 <sheetData/> / <row r="1"/> 与命名空间前缀（<x:row>/<x:c>）都能定位与写入', async () => {
    const empty = await withSheetXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<dimension ref="A1"/><sheetData/></worksheet>',
    );
    const filled = exportXlsx(empty, [{ sheetId: empty.sheets[0].id, cells: [{ row: 0, col: 0, value: 7 }] }]);
    expect(entryText(filled, 'xl/worksheets/sheet1.xml'))
      .toContain('<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>');
    expect(indexCells((await parseXlsx(filled)).sheets[0].cells).get(cellKey(0, 0))?.value).toBe(7);

    const selfClosing = await withSheetXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<dimension ref="A1:B2"/><sheetData><row r="1" ht="30" customHeight="1"/><row r="2"/></sheetData></worksheet>',
    );
    const out = exportXlsx(selfClosing, [{ sheetId: selfClosing.sheets[0].id, cells: [{ row: 0, col: 1, value: 'B1' }] }]);
    const sheetXml = entryText(out, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<row r="1" ht="30" customHeight="1"><c r="B1" t="s">');
    expect(sheetXml).toContain('<row r="2"/>'); // 没碰过的行一个字节都没动
    const re = await parseXlsx(out);
    expect(indexCells(re.sheets[0].cells).get(cellKey(0, 1))?.value).toBe('B1');
    expect(re.sheets[0].rows[0]).toEqual({ height: 30, customHeight: true });

    // 带命名空间前缀的文档
    const prefixed = await withSheetXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<x:dimension ref="A1:B2"/><x:sheetData><x:row r="1">'
      + '<x:c r="A1" s="1"><x:v>5</x:v></x:c>'
      + '</x:row></x:sheetData></x:worksheet>',
    );
    const prefixedOut = exportXlsx(prefixed, [{
      sheetId: prefixed.sheets[0].id,
      cells: [{ row: 0, col: 0, value: '名字空间' }, { row: 0, col: 1, value: 9 }],
    }]);
    const prefixedCells = indexCells((await parseXlsx(prefixedOut)).sheets[0].cells);
    expect(prefixedCells.get(cellKey(0, 0))?.value).toBe('名字空间');
    expect(prefixedCells.get(cellKey(0, 0))?.styleIndex).toBe(1);
    expect(prefixedCells.get(cellKey(0, 1))?.value).toBe(9);
    // 已知取舍：新写入的格子用默认命名空间下的 <c>（Excel 自己也不会写前缀形式）
    expect(entryText(prefixedOut, 'xl/worksheets/sheet1.xml')).toContain('<c r="B1" s="1"><v>9</v></c>');
  });

  it('清空"无属性无内容"的格子会整格移除，行空了连 <row> 一起清掉（有 s 的保留）', async () => {
    const source = await withSheetXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<dimension ref="A1:A2"/><sheetData>'
      + '<row r="1"><c r="A1" s="1"><v>5</v></c></row>'
      + '<row r="2"><c r="A2"><v>9</v></c></row>'
      + '</sheetData></worksheet>',
    );
    const out = exportXlsx(source, [
      { sheetId: source.sheets[0].id, cells: [{ row: 1, col: 0, value: null }] }, // 无属性 -> 整格移除，行也空
      { sheetId: source.sheets[0].id, cells: [{ row: 0, col: 0, value: null }] }, // 有 s -> 保留 <c r="A1" s="1"/>
    ]);
    const sheetXml = entryText(out, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).not.toContain('<row r="2"');
    expect(sheetXml).not.toContain('<c r="A2"');
    expect(sheetXml).toContain('<sheetData><row r="1"><c r="A1" s="1"/></row></sheetData>');
    const cells = indexCells((await parseXlsx(out)).sheets[0].cells);
    expect(cells.get(cellKey(0, 0))?.styleIndex).toBe(1);
    expect(cells.get(cellKey(0, 0))?.value).toBeUndefined();
    expect(cells.has(cellKey(1, 0))).toBe(false);
  });
});

/**
 * 内存优化（`slimForExport`）的核心保证：瘦身后的数据源导出结果必须与完整模型**等价**。
 *
 * 瘦身丢掉的是 `cells` / `rows` / 样式表等"导入那一刻才需要"的东西；导出对这些部件本来就是
 * 字节级原样保留，所以只要"行样式回退"这一条语义没变，产物就应该完全相同。
 *
 * 注意：这里比较的是**zip 内每个条目的字节**，不是整个 zip 文件的字节。
 * 原因：`zipSync` 未指定时间戳时会用当前时间写条目的 DOS 时间（2 秒粒度），
 * 于是"同样内容的两次导出"只要间隔跨过 2 秒，整包字节就不同——整包比较会变成随机失败的用例
 * （实测：单跑必过、满负载跑全套时偶发失败）。zip 外壳的时间戳不属于导出语义。
 */
describe('exportXlsx 瘦身数据源（内存优化）与完整模型等价', () => {
  const fixtures = ['fixture-styles.xlsx', 'fixture-multi.xlsx', 'fixture-table.xlsx'];
  const editSets: Record<string, SheetEdits[]> = {
    // 每个夹具都覆盖"写入新格子 + 改已有格子 + 清空"，正好压到行样式回退那条逻辑
    'fixture-styles.xlsx': [{
      sheetId: 'rId4',
      cells: [
        { row: 2, col: 0, value: '瘦身写入' },
        { row: 2, col: 4, formula: 'SUM(B3:B4)' },
        { row: 3, col: 1, value: null },
      ],
    }],
    'fixture-multi.xlsx': [],
    'fixture-table.xlsx': [],
  };

  /** 条目级等价：键集合一致，且每个条目的字节逐一相等 */
  function expectSameEntries(actual: Uint8Array, expected: Uint8Array, label: string): void {
    const a = entriesOf(actual);
    const b = entriesOf(expected);
    expect(Object.keys(a).sort(), `${label} 条目集合应一致`).toEqual(Object.keys(b).sort());
    for (const name of Object.keys(b)) {
      expect(bytesEqual(a[name], b[name]), `${label} 条目 ${name} 应逐字节相同`).toBe(true);
    }
  }

  for (const fileName of fixtures) {
    it(`${fileName}：瘦身来源 === 完整模型（条目级字节相等）`, async () => {
      const raw = bytesOf(fileName);
      const source = await parseXlsx(raw);

      // 默认编辑集：改第一张表的几个格子（没有专属编辑集时用这个）
      const edits = editSets[fileName].length > 0
        ? editSets[fileName]
        : [{ sheetId: source.sheets[0].id, cells: [{ row: 1, col: 1, value: '瘦身' }] }];

      const fromFull = exportXlsx(source, edits);
      const slim = slimForExport(source, raw); // 只留 sheet.id + 行样式映射 + 原始字节
      const fromSlim = exportXlsx(slim, edits);

      expectSameEntries(fromSlim, fromFull, fileName);

      // 瘦身对象确实不再持有解析模型（否则这个测试就没意义了）
      expect((slim as { raw?: unknown }).raw).toBeUndefined();
      expect(JSON.stringify(slim.sheets[0]).includes('cells')).toBe(false);
    });
  }

  it('既没有 raw 也没有 bytes 时明确报错（不静默产出坏文件）', async () => {
    const source = await loadWorkbook('fixture-styles.xlsx');
    const broken = { ...slimForExport(source), bytes: undefined, raw: undefined };
    expect(() => exportXlsx(broken, [{ sheetId: 'rId4', cells: [{ row: 0, col: 0, value: 1 }] }]))
      .toThrow(/raw\.entries 为空/);
  });

  it('行样式回退语义不变：新格子沿用所在行第一个带样式单元格的 s', async () => {
    const raw = bytesOf('fixture-styles.xlsx');
    const source = await parseXlsx(raw);
    const edits: SheetEdits[] = [{ sheetId: 'rId4', cells: [{ row: 2, col: 0, value: '样式回退' }] }];

    const full = await parseXlsx(exportXlsx(source, edits));
    const slim = await parseXlsx(exportXlsx(slimForExport(source, raw), edits));
    const fullStyle = indexCells(full.sheets[0].cells).get(cellKey(2, 0))?.styleIndex;
    const slimStyle = indexCells(slim.sheets[0].cells).get(cellKey(2, 0))?.styleIndex;

    expect(fullStyle).toBeDefined();
    expect(slimStyle).toBe(fullStyle);
  });
});
