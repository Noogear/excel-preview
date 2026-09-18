/**
 * P0 OOXML 解析器单测。
 *
 * 全部用**自造 xlsx**（手写 XML + fflate.zipSync 打成内存 zip），
 * 不依赖任何外部 fixture，也不读磁盘（只有"fixtures 目录存在才跑"的那两个用例例外）。
 *
 * 覆盖：共享字符串 / 内联字符串 / 布尔 / 数字 / 公式+缓存 / 样式索引（字体·背景·边框·对齐·wrapText）/
 * 自定义 numFmt / 内置 numFmt / theme+tint 颜色 / indexed 颜色 / 合并单元格 / 冻结窗格 /
 * 行高列宽 / 隐藏行列 / 两个工作表且顺序与文件名不一致 / unsupported 记账 / 富文本降级 warning /
 * xml tokenizer 的三类真实回归（自闭合兄弟、前缀同名标签、标签间文本）。
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { zipSync, strToU8, unzipSync, type Zippable } from 'fflate';

import { parseXlsx } from '../../src/parser';
import { applyTint } from '../../src/parser/styles';
import type { ParsedCell, ParsedSheet } from '../../src/parser/types';

/* -------------------------------------------------------------------------- */
/* 自造 xlsx                                                                   */
/* -------------------------------------------------------------------------- */

/** 用给定 XML 部件打一个内存 zip（xlsx 就是一个 zip） */
function makeXlsx(parts: Record<string, string>): Uint8Array {
  const zippable: Zippable = {};
  for (const [path, xml] of Object.entries(parts)) {
    zippable[path] = strToU8(xml);
  }
  return zipSync(zippable, { level: 6 });
}

/** 单元格断言用的小工具：按 row/col 取值 */
function cellAt(sheet: ParsedSheet, row: number, col: number): ParsedCell | undefined {
  return sheet.cells.find((c) => c.row === row && c.col === col);
}

function valueAt(sheet: ParsedSheet, row: number, col: number): unknown {
  return cellAt(sheet, row, col)?.value;
}

/* -------------------------------------------------------------------------- */
/* 手的 XML 部件                                                               */
/* -------------------------------------------------------------------------- */

/** 根关系：officeDocument 指向 xl/workbook.xml */
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

/**
 * 关键点：tab 顺序是 SheetTwo(rId2 -> sheet2.xml) 在前、SheetOne(rId1 -> sheet1.xml) 在后，
 * 文件编号与显示顺序故意错开——解析器不能假设 sheet1.xml 就是第一个工作表。
 * SheetThree 是 hidden。
 */
const WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="SheetTwo" sheetId="2" r:id="rId2"/>
    <sheet name="SheetOne" sheetId="1" r:id="rId1"/>
    <sheet name="SheetThree" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
</workbook>`;

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

/** 索引 0/1/2，其中 2 是富文本（多 run），P0 必须降级为纯文本并记 warning */
const SHARED_STRINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World &amp; Co &lt;P0&gt;</t></si>
  <si>
    <r><rPr><b/><sz val="11"/></rPr><t>Rich</t></r>
    <r><t xml:space="preserve"> Text</t></r>
  </si>
</sst>`;

/** 主题：dk1/lt1 用 sysClr（真实 Excel 就是这样），dk2/lt2 用 srgbClr，其余 accent 用 srgbClr */
const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Probe Theme">
  <a:themeElements>
    <a:clrScheme name="Probe">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;

/**
 * cellXfs（7 个，下标 == 单元格 s 属性）：
 * 0 = 全默认（只写盘 numFmtId=0）
 * 1 = 全套字体/填充/边框/对齐
 * 2 = 自定义 numFmt 176
 * 3 = 内置 numFmt 4
 * 4 = 内置 numFmt 9
 * 5 = indexed 颜色字体
 * 6 = rgb 纯色填充
 *
 * 注意：`<cellStyleXfs>` 故意写在 `<cellXfs>` 前面——`cellStyleXfs` 的后缀正是 `cellXfs`，
 * 解析器必须精确区分这两个容器，否则样式索引整体错位。
 */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="176" formatCode="#,##0.0&quot;元&quot;"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><i/><u val="single"/><strike/><sz val="14"/><color rgb="FFCC0000"/><name val="Arial"/></font>
    <font><sz val="10"/><color indexed="10"/><name val="SimSun"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.3999755851924192"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF4F81BD" tint="-0.25"/></left>
      <right style="mediumDashDot"><color indexed="12"/></right>
      <top style="double"><color rgb="FF000000"/></top>
      <bottom style="hair"><color theme="2" tint="0.5"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1" indent="2" textRotation="45"/>
    </xf>
    <xf numFmtId="176" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>`;

/**
 * SheetOne：样式、主题色、合并、冻结、行高列宽、隐藏行列、dimension，
 * 以及**不解析但要记账**的 extLst。最后两个单元格是畸形数据
 * （错误引用、越界共享字符串），用来验证容错。
 */
const SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:E6"/>
  <sheetViews>
    <sheetView showGridLines="0" tabSelected="1" workbookViewId="0">
      <pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15.5" defaultColWidth="9.140625"/>
  <cols>
    <col min="2" max="3" width="24.71" customWidth="1"/>
    <col min="4" max="4" width="12.5" hidden="1" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="30" customHeight="1">
      <c r="A1" s="1"><v>42</v></c>
      <c r="AA1" s="0"><v>7</v></c>
    </row>
    <row r="3" hidden="1">
      <c r="A3" t="s"><v>0</v></c>
      <c r="B3"><v>3.14159</v></c>
    </row>
    <row r="5">
      <c r="A5" s="2"><v>1234.5</v></c>
      <c r="B5" s="3"><v>1234.5</v></c>
      <c r="C5" s="4"><v>0.25</v></c>
      <c r="D5" s="5"><v>9</v></c>
      <c r="E5" s="6"><v>11</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>99</v></c>
    </row>
    <row r="7">
      <c r="NOTAREF"><v>1</v></c>
    </row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:C1"/>
    <mergeCell ref="B5:B6"/>
  </mergeCells>
  <conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="B1:B3"><formula1>"a,b"</formula1></dataValidation></dataValidations>
  <hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>
  <tableParts count="1"><tablePart r:id="rId2"/></tableParts>
  <extLst><ext uri="{X}"><x14:foo xmlns:x14="http://x"/></ext></extLst>
</worksheet>`;

/** SheetTwo：值类型全覆盖（共享/内联/富文本/布尔/错误/公式）+ 另一个冻结窗格 */
const SHEET2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetViews>
    <sheetView showGridLines="1" workbookViewId="0">
      <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
      <c r="B1" t="s"><v>2</v></c>
      <c r="C1" t="b"><v>1</v></c>
      <c r="D1" t="b"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2" t="e"><v>#DIV/0!</v></c>
      <c r="B2" t="str"><f>CONCATENATE("a","b")</f><v>ab</v></c>
      <c r="C2"><f>SUM(A1:A3)</f><v>6</v></c>
      <c r="D2"><f>SUM(A1:A3)</f></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Inline Text</t></is></c>
      <c r="B3" t="inlineStr"><is><r><t>Inline</t></r><r><t> Rich</t></r></is></c>
      <c r="C3"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`;

/** SheetThree：空表（隐藏表），只验证不抛异常 */
const SHEET3 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`;

const ALL_PARTS: Record<string, string> = {
  '_rels/.rels': ROOT_RELS,
  'xl/workbook.xml': WORKBOOK,
  'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
  'xl/sharedStrings.xml': SHARED_STRINGS,
  'xl/theme/theme1.xml': THEME,
  'xl/styles.xml': STYLES,
  'xl/worksheets/sheet1.xml': SHEET1,
  'xl/worksheets/sheet2.xml': SHEET2,
  'xl/worksheets/sheet3.xml': SHEET3,
};

const BOOK = makeXlsx(ALL_PARTS);

/** tokenizer 与冻结窗格用例共用的最小工作簿（无 styles/theme，避免无关噪声） */
const minimalPartsForTokenizer: Record<string, string> = {
  '_rels/.rels': ROOT_RELS,
  'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="T" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
};

/* -------------------------------------------------------------------------- */
/* 用例                                                                        */
/* -------------------------------------------------------------------------- */

describe('parseXlsx：工作簿骨架', () => {
  it('按 workbook.xml 的顺序输出工作表（而不是按文件名），标注 hidden，完整保留 zip 条目，也接受 ArrayBuffer', async () => {
    const wb = await parseXlsx(BOOK);
    expect(wb.sheets.map((s) => s.name)).toEqual(['SheetTwo', 'SheetOne', 'SheetThree']);
    expect(wb.sheets.map((s) => s.index)).toEqual([0, 1, 2]);
    expect(wb.sheets.map((s) => s.id)).toEqual(['rId2', 'rId1', 'rId3']);

    const [, one, three] = wb.sheets;
    expect(three.hidden).toBe(true);
    expect(three.veryHidden).toBeUndefined();
    expect(one.hidden).toBeUndefined();

    // raw.entries 完整保留所有 zip 条目（供后续外科式修补）
    expect(Object.keys(wb.raw.entries).sort()).toEqual(Object.keys(ALL_PARTS).sort());
    expect(wb.raw.entries['xl/worksheets/sheet2.xml']).toBeInstanceOf(Uint8Array);

    const copy = new Uint8Array(BOOK.byteLength);
    copy.set(BOOK);
    expect((await parseXlsx(copy.buffer)).sheets).toHaveLength(3);
  });
});

describe('parseXlsx：单元格引用与值类型', () => {
  it('任意列字母（含 AA）转成 0-based 行列；共享字符串 / 实体解码 / 富文本降级 / 布尔 / 错误 / 数字都对；公式取 <f> 去掉前导 =，缓存值取 <v>', async () => {
    const wb = await parseXlsx(BOOK);
    const [two, one] = wb.sheets;

    // SheetOne: A1 -> 0,0；AA1 -> 0,26
    expect(cellAt(one, 0, 0)?.value).toBe(42);
    expect(cellAt(one, 0, 26)?.value).toBe(7);

    expect(valueAt(two, 0, 0)).toBe('World & Co <P0>');   // 实体解码 + 共享字符串
    expect(valueAt(two, 0, 1)).toBe('Rich Text');          // 富文本 run 拼接
    expect(valueAt(two, 0, 2)).toBe(true);
    expect(valueAt(two, 0, 3)).toBe(false);
    expect(cellAt(two, 1, 0)?.error).toBe('#DIV/0!');
    expect(cellAt(two, 1, 0)?.value).toBeUndefined();
    expect(valueAt(two, 0, 26)).toBeUndefined();           // 该表没有 AA 列
    expect(valueAt(two, 2, 2)).toBe(0);

    // 公式：取 <f> 且去掉前导 =，缓存值取 <v>（无缓存值时 value 为 undefined）
    expect(cellAt(two, 1, 1)?.formula).toBe('CONCATENATE("a","b")');
    expect(cellAt(two, 1, 1)?.value).toBe('ab');           // t="str"
    expect(cellAt(two, 1, 2)?.formula).toBe('SUM(A1:A3)');
    expect(cellAt(two, 1, 2)?.value).toBe(6);
    expect(cellAt(two, 1, 3)?.formula).toBe('SUM(A1:A3)');
    expect(cellAt(two, 1, 3)?.value).toBeUndefined();

    // 内联字符串：普通 <is><t> 与多 run 富文本
    expect(valueAt(two, 2, 0)).toBe('Inline Text');
    expect(valueAt(two, 2, 1)).toBe('Inline Rich');
  });

  it('行/列信息（行高、隐藏行、customHeight、列宽、隐藏列、默认尺寸、<col min max> 1-based 闭区间）与合并单元格、dimension、冻结窗格（frozen/frozenSplit 算冻结，split 不算）', async () => {
    const wb = await parseXlsx(BOOK);
    const [two, one] = wb.sheets;
    expect(one.rows[0]?.height).toBe(30);
    expect(one.rows[0]?.customHeight).toBe(true);
    expect(one.rows[2]?.hidden).toBe(true);
    // 列宽保持 Excel 字符宽度，不做像素换算
    expect(one.cols[1]?.width).toBe(24.71);
    expect(one.cols[2]?.width).toBe(24.71);
    expect(one.cols[1]?.customWidth).toBe(true);
    expect(one.cols[3]?.width).toBe(12.5);
    expect(one.cols[3]?.hidden).toBe(true);
    expect(one.defaultRowHeight).toBe(15.5);
    expect(one.defaultColWidth).toBe(9.140625);
    expect(two.defaultRowHeight).toBe(15);

    expect(one.merges).toEqual([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
      { startRow: 4, startCol: 1, endRow: 5, endCol: 1 },
    ]);

    // 台面尺寸与冻结窗格：xSplit/ySplit 是"冻结的行列数"
    expect(one.dimension).toEqual({ startRow: 0, startCol: 0, endRow: 5, endCol: 4 });
    expect(one.freeze).toEqual({ row: 1, col: 2 });   // xSplit=2 -> col 2, ySplit=1 -> row 1
    expect(one.gridlinesHidden).toBe(true);
    expect(two.freeze).toEqual({ row: 3, col: 0 });
    expect(two.gridlinesHidden).toBeUndefined();

    const mk = async (pane: string): Promise<{ row: number; col: number } | undefined> => {
      const parsed = await parseXlsx(makeXlsx({
        ...minimalPartsForTokenizer,
        'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews><sheetData/></worksheet>`,
      }));
      return parsed.sheets[0].freeze;
    };
    // 真实样本 fixture-styles.xlsx 里就是 xSplit=1 ySplit=1，冻结的是第 0 行与第 0 列
    expect(await mk('<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'))
      .toEqual({ row: 1, col: 1 });
    expect(await mk('<pane xSplit="2" ySplit="3" state="frozenSplit"/>')).toEqual({ row: 3, col: 2 });
    expect(await mk('<pane xSplit="2" ySplit="3" activePane="bottomRight" state="split"/>')).toBeUndefined();

    // <col min max> 是 1-based 闭区间，展开到 0-based cols 不能差一
    const colWb = await parseXlsx(makeXlsx({
      ...minimalPartsForTokenizer,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet>
  <cols>
    <col min="2" max="3" width="24.71" customWidth="1"/>
    <col min="5" max="5" width="9" hidden="1"/>
  </cols>
  <sheetData/>
</worksheet>`,
    }));
    const cols = colWb.sheets[0].cols;
    // min=2 (B) -> 下标 1；max=3 (C) -> 下标 2
    expect(Object.keys(cols).map(Number).sort((a, b) => a - b)).toEqual([1, 2, 4]);
    expect(cols[1]?.width).toBe(24.71);
    expect(cols[2]?.width).toBe(24.71);
    expect(cols[4]?.hidden).toBe(true);
    expect(cols[0]).toBeUndefined(); // A 列不该被 min=2 波及
  });
});

describe('parseXlsx：样式', () => {
  it('styles 长度 == cellXfs，索引即样式索引（不能把 cellStyleXfs 混进来），引用下标对齐', async () => {
    const wb = await parseXlsx(BOOK);
    expect(wb.styles).toHaveLength(7);
    // 下标 0 就是 cellXfs 的第 0 项（若误把 cellStyleXfs 当容器，这里会整体错位）
    expect(wb.styles[0]).toEqual({
      fontFamily: 'Calibri',
      fontSize: 11,
      color: '#000000',       // <color theme="1"/> -> dk1(Text 1) = 黑
      numberFormat: 'General',
    });

    const one = wb.sheets[1];
    expect(cellAt(one, 0, 0)?.styleIndex).toBe(1);
    expect(cellAt(one, 0, 26)?.styleIndex).toBe(0);
    expect(cellAt(one, 4, 0)?.styleIndex).toBe(2);
    expect(cellAt(one, 2, 1)?.styleIndex).toBeUndefined(); // 无 s 属性

    // 字体/填充/边框/对齐/wrapText 一次性映射
    const s = wb.styles[1];
    expect(s.fontFamily).toBe('Arial');
    expect(s.fontSize).toBe(14);
    expect(s.bold).toBe(true);
    expect(s.italic).toBe(true);
    expect(s.underline).toBe(true);
    expect(s.strikeThrough).toBe(true);
    expect(s.color).toBe('#CC0000');                    // FFCC0000 -> 去 alpha
    expect(s.fill).toBe('#A2B4CA');                     // theme 4 = accent1 + tint 0.3999755851924192
    expect(s.horizontalAlign).toBe('center');
    expect(s.verticalAlign).toBe('middle');             // vertical="center" -> middle
    expect(s.textWrap).toBe(true);
    expect(s.indent).toBe(2);
    expect(s.textRotation).toBe(45);
    expect(s.border?.left).toEqual({ style: 'thin', color: '#376093' });   // FF4F81BD + tint -0.25
    expect(s.border?.right).toEqual({ style: 'mediumDashDot', color: '#0000FF' }); // indexed 12
    expect(s.border?.top).toEqual({ style: 'double', color: '#000000' });
    expect(s.border?.bottom).toEqual({ style: 'hair', color: '#8CA4C2' }); // theme 2 = dk2(#1F497D) + tint 0.5

    // 主题色索引错位规则 + 自定义/内置 numFmt + indexed 与 rgb 颜色
    // 默认字体 <color theme="1"/> 必须是黑色（真实 Excel 也是这样写的）
    expect(wb.styles[0].color).toBe('#000000');
    // 边框那条用的是 theme="2"，对应 dk2(#1F497D) + tint 0.5
    expect(wb.styles[1].border?.bottom?.color).toBe(applyTint('#1F497D', 0.5));
    // themeColors 数组本身保持 clrScheme 的书写顺序（dk1,lt1,dk2,lt2,...,hlink,folHlink）
    expect(wb.themeColors?.[2]).toBe('#1F497D'); // dk2
    expect(wb.themeColors?.[3]).toBe('#EEECE1'); // lt2
    expect(wb.themeColors?.[4]).toBe('#4F81BD'); // accent1
    expect(wb.themeColors).toEqual([
      '#000000', '#FFFFFF', '#1F497D', '#EEECE1',
      '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646',
      '#0000FF', '#800080',
    ]); // sysClr 走 lastClr

    expect(wb.styles[2].numberFormat).toBe('#,##0.0"元"'); // 自定义 176
    expect(wb.styles[3].numberFormat).toBe('#,##0.00');    // 内置 4
    expect(wb.styles[4].numberFormat).toBe('0%');          // 内置 9
    expect(wb.styles[0].numberFormat).toBe('General');     // 内置 0
    expect(wb.styles[5].color).toBe('#FF0000');            // indexed 10
    expect(wb.styles[6].fill).toBe('#00B050');             // rgb 纯色
  });
});
describe('parseXlsx：report 与容错', () => {
  it('未解析的特性按种类记账（含出现次数），已支持的不再计入，且富文本降级必须留痕', async () => {
    const wb = await parseXlsx(BOOK);
    const joined = wb.report.unsupported.join('\n');

    // P1 起：条件格式/数据验证/超链接/表格 已能解析 → **不再**计入 unsupported
    for (const keyword of ['conditionalFormatting', 'dataValidations', 'hyperlink', 'tableParts']) {
      expect(joined, `${keyword} 已支持，不应再出现在未支持清单里`).not.toContain(keyword);
    }
    // 仍未解析的（扩展列表）继续按种类计数，每条都带工作表前缀
    expect(joined).toContain('extLst');
    expect(joined).toContain('1 处');
    expect(wb.report.unsupported.every((u) => u.startsWith('[SheetOne]') || u.startsWith('[SheetTwo]'))).toBe(true);

    // 并且这些特性确实进了模型（不是被静默丢掉）
    const totalCf = wb.sheets.reduce((n, s) => n + (s.conditionalFormats?.length ?? 0), 0);
    const totalDv = wb.sheets.reduce((n, s) => n + (s.dataValidations?.length ?? 0), 0);
    const totalLinks = wb.sheets.reduce((n, s) => n + (s.hyperlinks?.length ?? 0), 0);
    expect(totalCf, '条件格式应被解析进模型').toBe(1);
    expect(totalDv, '数据验证应被解析进模型').toBe(1);
    expect(totalLinks, '超链接应被解析进模型').toBe(1);

    // 富文本降级必须留痕（sharedStrings 与 inlineStr 两条路径）
    const warnings = wb.report.warnings.join('\n');
    expect(warnings).toContain('降级为纯文本');
    expect(warnings).toContain('sharedStrings');
    expect(warnings).toContain('inlineStr');

    // 畸形数据只记 warning 不抛异常（越界共享字符串 → null，非法引用也留痕）
    const malformed = wb.report.warnings.join('\n');
    expect(malformed).toContain('共享字符串索引 99 越界');
    expect(malformed).toContain('NOTAREF');
    expect(cellAt(wb.sheets[1], 5, 0)?.value).toBeNull();
  });

  it('缺 styles/theme/sharedStrings 也能解析并给出 warning；工作表部件缺失按空表处理', async () => {
    const minimal: Record<string, string> = {
      '_rels/.rels': ROOT_RELS,
      'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Only" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
    };

    const wb = await parseXlsx(makeXlsx(minimal));
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe('Only');
    expect(wb.styles).toEqual([]);
    expect(wb.themeColors).toBeUndefined();
    expect(wb.report.warnings.join('\n')).toContain('styles.xml');

    const broken = { ...minimal };
    delete broken['xl/worksheets/sheet1.xml'];
    const missingSheet = await parseXlsx(makeXlsx(broken));
    expect(missingSheet.sheets).toHaveLength(1);
    expect(missingSheet.sheets[0].cells).toEqual([]);
    expect(missingSheet.report.warnings.join('\n')).toContain('不存在');

    // 根本不是 zip 时返回空工作簿 + warning，不抛异常
    const notZip = await parseXlsx(strToU8('this is definitely not a zip file'));
    expect(notZip.sheets).toEqual([]);
    expect(notZip.report.warnings.length).toBeGreaterThan(0);
  });
});

/* -------------------------------------------------------------------------- */
/* xml tokenizer 回归                                                          */
/* -------------------------------------------------------------------------- */

/**
 * 这三个 bug 都真实发生过（在 ExcelJS 产出的 xlsx 上暴露），单独用 XML 部件锁死：
 * 1. `<r>` 与 `<rPr>` 前缀同名 —— `indexOf('</r')` 会命中 `</rPr>`，导致嵌套层级错乱；
 * 2. 标签之间的**文本**被当成元素（`<t>Rich</t>` 会解析出名叫 "Rich" 的元素）；
 * 3. 自闭合元素之后游标推进错误，兄弟元素被整体丢弃。
 */
describe('xml tokenizer 回归', () => {
  const runSheet = async (sheetDataInner: string, sharedStrings = ''): Promise<ParsedSheet> => {
    const parts: Record<string, string> = {
      ...minimalPartsForTokenizer,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet><sheetData>${sheetDataInner}</sheetData></worksheet>`,
    };
    if (sharedStrings) parts['xl/sharedStrings.xml'] = sharedStrings;
    const wb = await parseXlsx(makeXlsx(parts));
    return wb.sheets[0];
  };

  it('连续自闭合兄弟：后面每个元素都要保留（不能吞掉后一个单元格）', async () => {
    const sheet = await runSheet(
      '<row r="1"><c r="A1" s="1"/><c r="B1" s="1"/><c r="C1" s="1"/><c r="D1" s="1"/></row>',
    );
    expect(sheet.cells.map((c) => [c.col, c.styleIndex])).toEqual([
      [0, 1], [1, 1], [2, 1], [3, 1],
    ]);
  });

  it('自闭合元素后紧跟普通兄弟元素', async () => {
    const sheet = await runSheet(
      '<row r="1"><c r="A1" s="2"/><c r="B1"><v>9</v></c><c r="C1" t="b"><v>1</v></c></row>',
    );
    expect(cellAt(sheet, 0, 0)?.styleIndex).toBe(2);
    expect(cellAt(sheet, 0, 0)?.value).toBeUndefined();
    expect(cellAt(sheet, 0, 1)?.value).toBe(9);
    expect(cellAt(sheet, 0, 2)?.value).toBe(true);
  });

  it('嵌套自闭合：容器里的自闭合子元素不能截断容器（富文本 run 降级为纯文本）', async () => {
    const sheet = await runSheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>' +
      '<c r="B1" t="inlineStr"><is><r><rPr><b/><sz val="11"/></rPr><t>Rich</t></r><r><t xml:space="preserve"> Text</t></r></is></c>' +
      '<c r="C1"><v>3</v></c></row>',
    );
    expect(valueAt(sheet, 0, 0)).toBe('a');
    expect(valueAt(sheet, 0, 1)).toBe('Rich Text');
    expect(valueAt(sheet, 0, 2)).toBe(3); // 自闭合 <b/>、<sz/> 之后兄弟仍要在
  });

  it('t 属性缺失但带 <is> 时按内联字符串容错（不丢内容）', async () => {
    const sheet = await runSheet('<row r="1"><c r="A1"><is><t>loose</t></is></c></row>');
    expect(valueAt(sheet, 0, 0)).toBe('loose');
  });

  it('标签名互为前缀：<r>/<rPr>、<si>/<s> 不能互相误配', async () => {
    // 共享字符串里的富文本 run（<r> 与 <rPr> 前缀同名）
    const sst = `<?xml version="1.0"?>
<sst count="1" uniqueCount="1">
  <si><r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r><r><t>+Plain</t></r></si>
</sst>`;
    const sheet = await runSheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', sst);
    expect(valueAt(sheet, 0, 0)).toBe('Bold+Plain');
  });

  it('自闭合元素带命名空间属性（x14ac:dyDescent 这类）', async () => {
    const sheet = await runSheet(
      '<row r="1" ht="30" customHeight="1" spans="1:3" x14ac:dyDescent="0.25">' +
      '<c r="A1" s="3" t="s"><v>0</v></c><c r="B1" s="3"/>' +
      '<c r="C1" s="0"><v>2</v></c></row>',
      '<?xml version="1.0"?><sst count="1" uniqueCount="1"><si><t>NS</t></si></sst>',
    );
    expect(sheet.rows[0]?.height).toBe(30);
    expect(sheet.cells).toHaveLength(3);
    expect(valueAt(sheet, 0, 0)).toBe('NS');
    expect(valueAt(sheet, 0, 2)).toBe(2);
  });

  it('文本内容不会被当成元素（<v> 里的数字与实体）', async () => {
    const sheet = await runSheet(
      '<row r="1"><c r="A1"><v>42</v></c><c r="B1" t="str"><f>IF(1,&quot;a&amp;b&quot;,&quot;&lt;x&gt;&quot;)</f><v>a&amp;b</v></c></row>',
    );
    expect(valueAt(sheet, 0, 0)).toBe(42);
    expect(cellAt(sheet, 0, 1)?.formula).toBe('IF(1,"a&b","<x>")');
    expect(valueAt(sheet, 0, 1)).toBe('a&b');
  });
});

/* -------------------------------------------------------------------------- */
/* 可选 fixture：fixtures/*.xlsx 由另一个任务生成，缺失时跳过                     */
/* -------------------------------------------------------------------------- */

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = resolve(HERE, '../../fixtures');

function listFixtures(): string[] {
  if (!existsSync(FIXTURE_DIR)) return [];
  return readdirSync(FIXTURE_DIR).filter((f) => f.toLowerCase().endsWith('.xlsx'));
}

describe('fixtures/*.xlsx（外部生成，缺失时跳过）', () => {
  const fixtures = listFixtures();

  /**
   * 基于真值的计数断言：直接从 zip 里正则统计 XML，再和解析结果对比。
   * 这是唯一能抓住"tokenizer 悄悄吞掉元素"这类 bug 的断言方式——
   * 只断言"不抛异常"完全测不出少解析（真实案例：40 个单元格只解析出 8 个）。
   */
  it.skipIf(fixtures.length === 0)('行列/cells/merges 数量与 XML 真值一致，且每个样本都至少有一张表', async () => {
    expect(fixtures.length).toBeGreaterThan(0);
    for (const file of fixtures) {
      const bytes = new Uint8Array(readFileSync(join(FIXTURE_DIR, file)));
      const entries = unzipSync(bytes);
      const wb = await parseXlsx(bytes);
      const decode = new TextDecoder();

      expect(wb.sheets.length, `${file} 解析出的工作表为空`).toBeGreaterThan(0);
      expect(wb.raw.entries, `${file} raw.entries 为空`).toBeTruthy();

      // 每个 worksheet 部件都必须被解析到（不允许"部分解析"被平均值掩盖）
      const sheetParts = Object.keys(entries)
        .filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(k))
        .sort();
      let cTotal = 0;
      let rTotal = 0;
      let mergeTotal = 0;
      for (const part of sheetParts) {
        const xml = decode.decode(entries[part]);
        cTotal += (xml.match(/<c[ >/]/g) ?? []).length;
        rTotal += (xml.match(/<row[ >/]/g) ?? []).length;
        mergeTotal += (xml.match(/<mergeCell[ >/]/g) ?? []).length;
      }

      const totalCells = wb.sheets.reduce((sum, s) => sum + s.cells.length, 0);
      const totalMerges = wb.sheets.reduce((sum, s) => sum + s.merges.length, 0);
      expect(totalCells, `${file}: 解析出的 cells 少于 XML 中的 <c> 数量（tokenizer 吞元素）`)
        .toBeGreaterThanOrEqual(cTotal);
      expect(totalMerges, `${file}: 合并单元格数量与真值不一致`).toBe(mergeTotal);
      expect(totalCells, `${file}: 单元格数不该超过 <c> 总数`).toBeLessThanOrEqual(cTotal);
      // rows 只记录带 ht/hidden 等属性的行，所以只需不超过真值
      const totalRows = wb.sheets.reduce((sum, s) => sum + Object.keys(s.rows).length, 0);
      expect(totalRows, `${file}: rows 多于 XML 中的 <row> 数量`).toBeLessThanOrEqual(rTotal);
    }
  });

  it.skipIf(fixtures.length === 0)('styles 数量与 cellXfs 真值一致，且不出现"引用不存在的 id"', async () => {
    for (const file of fixtures) {
      const entries = unzipSync(new Uint8Array(readFileSync(join(FIXTURE_DIR, file))));
      if (!entries['xl/styles.xml']) continue;
      const stylesXml = new TextDecoder().decode(entries['xl/styles.xml']);
      const cellXfs = /<cellXfs[\s\S]*?<\/cellXfs>/.exec(stylesXml)?.[0] ?? '';
      const xfCount = (cellXfs.match(/<xf[ >/]/g) ?? []).length;
      const wb = await parseXlsx(new Uint8Array(readFileSync(join(FIXTURE_DIR, file))));
      expect(wb.styles.length, `${file}: styles 数量应等于 cellXfs 的 xf 数量`).toBe(xfCount);
      const joined = wb.report.warnings.join('\n');
      expect(joined, `${file}: 不应出现越界 id 警告`).not.toContain('不存在的 fillId');
      expect(joined, `${file}: 不应出现越界 id 警告`).not.toContain('不存在的 borderId');
      expect(joined, `${file}: 不应出现越界 id 警告`).not.toContain('不存在的 fontId');
    }
  });
});
