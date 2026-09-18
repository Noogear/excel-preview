/**
 * P1 解析器单测：条件格式 / 数据验证 / dxf 差异样式 / 超链接 / 批注 / Excel 表格 / 浮动图片。
 *
 * 两部分：
 * 1. **fixtures/*.xlsx 真值断言**——真值来自直接读 zip 里的 XML（不是猜的），
 *    断言到字段级，防止"解析器悄悄少解析"被平均值掩盖。
 * 2. 少量手写 XML 的边界用例——fixtures 里没有的形态（multi-range sqref、
 *    absoluteAnchor、只有 `<t>` 的批注、dxf 的 fgColor/bgColor 反转、坏部件容错）。
 */
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strToU8, zipSync, type Zippable } from 'fflate';

import {
  expandSqref,
  isA1RangeToken,
  parseComments,
  parseConditionalFormatting,
  parseDataValidations,
  parseDrawing,
  parseDxfStyles,
  parseTable,
  parseXlsx,
  splitSqref,
} from '../../src/parser';
import { parseStyles } from '../../src/parser/styles';
import { parsePrintSettings } from '../../src/parser/worksheet';
import type { ParsedSheet, ParsedWorkbook } from '../../src/parser/types';

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = resolve(HERE, '../../fixtures');

async function loadFixture(file: string): Promise<ParsedWorkbook> {
  const bytes = new Uint8Array(readFileSync(join(FIXTURE_DIR, file)));
  return parseXlsx(bytes);
}

function sheetOf(wb: ParsedWorkbook, name: string): ParsedSheet {
  const sheet = wb.sheets.find((s) => s.name === name);
  if (!sheet) throw new Error(`fixture 里没有工作表「${name}」，实际有：${wb.sheets.map((s) => s.name).join(', ')}`);
  return sheet;
}

/** 用给定 XML 部件打一个内存 zip（xlsx 就是一个 zip） */
function makeXlsx(parts: Record<string, string>): Uint8Array {
  const zippable: Zippable = {};
  for (const [path, xml] of Object.entries(parts)) zippable[path] = strToU8(xml);
  return zipSync(zippable, { level: 6 });
}

/* ========================================================================== */
/* fixture-rules.xlsx：条件格式四类 + 数据验证五类 + <dxfs>                     */
/* ========================================================================== */

describe('fixture-rules.xlsx：条件格式四类', () => {
  it('4 条规则的顺序 / 范围 / 优先级 / 类型 + 四类（highlight·cellIs / colorScale / dataBar / iconSet）各自的字段；`<dxfs>` 的 hex 归一成 #RRGGBB', async () => {
    const wb = await loadFixture('fixture-rules.xlsx');
    const cf = sheetOf(wb, '条件格式').conditionalFormats ?? [];
    expect(cf).toHaveLength(4);
    expect(cf.map((r) => r.ranges)).toEqual([['B2:B6'], ['C2:C6'], ['D2:D6'], ['E2:E6']]);
    expect(cf.map((r) => r.priority)).toEqual([1, 2, 3, 4]);
    expect(cf.map((r) => r.kind)).toEqual(['highlight', 'colorScale', 'dataBar', 'iconSet']);
    expect(cf.map((r) => r.ruleType)).toEqual(['cellIs', 'colorScale', 'dataBar', 'iconSet']);

    // styles.xml 的 <dxfs count="1"> -> dxfStyles[0]
    expect(wb.dxfStyles).toHaveLength(1);
    expect(wb.dxfStyles?.[0]).toEqual({ bold: true, color: '#9C0006', fill: '#FFC7CE' });

    // highlight（cellIs）：operator/formula1/dxfId，dxfId 顺带挂上 <dxfs> 的样式
    const rule = cf[0];
    expect(rule.operator).toBe('greaterThan');
    expect(rule.formula1).toBe('100');
    expect(rule.dxfId).toBe(0);
    expect(rule.dxf).toEqual({ bold: true, color: '#9C0006', fill: '#FFC7CE' });

    // colorScale / dataBar / iconSet 各自的阈值与颜色
    expect(cf[1].cfvo).toEqual([{ type: 'min' }, { type: 'max' }]);
    expect(cf[1].colors).toEqual(['#F8696B', '#63BE7B']); // 低 -> 高
    expect(cf[2].cfvo).toEqual([{ type: 'min' }, { type: 'max' }]);
    expect(cf[2].color).toBe('#638EC6');
    expect(cf[3].iconSet).toBe('3TrafficLights1');
    expect(cf[3].cfvo).toEqual([
      { type: 'percent', value: '0' },
      { type: 'percent', value: '33' },
      { type: 'percent', value: '67' },
    ]);
  });
});

describe('fixture-rules.xlsx：数据验证五类', () => {
  it('顺序 / 范围 / type / operator / formula1：list / whole / decimal / date / textLength（含 OOXML 默认 between 兜底与 list 无 operator）', async () => {
    const wb = await loadFixture('fixture-rules.xlsx');
    const dv = sheetOf(wb, '数据验证').dataValidations ?? [];
    expect(dv).toHaveLength(5);
    expect(dv.map((r) => r.ranges)).toEqual([['A2'], ['B2'], ['C2'], ['D2'], ['E2']]);
    expect(dv.map((r) => r.type)).toEqual(['list', 'whole', 'decimal', 'date', 'textLength']);
    expect(dv.map((r) => r.operator)).toEqual([undefined, 'between', 'greaterThan', 'greaterThan', 'lessThanOrEqual']);
    expect(dv.map((r) => r.formula1)).toEqual(['"男,女,未知"', '1', '0', '43831', '5']);

    // list：内联列表公式 + allowBlank，且 list 没有 operator（不能硬填规范默认值）
    const list = dv[0];
    expect(list.formula1).toBe('"男,女,未知"');
    expect(list.allowBlank).toBe(true);
    expect(list.ranges).toEqual(['A2']);
    expect(list.operator).toBeUndefined();

    // ⚠️ XML 里 <dataValidation type="whole" allowBlank="1" …> **没有 operator 属性** -> 按 OOXML 默认 between 补齐
    expect(dv[1].operator).toBe('between');
    expect(dv[1].formula1).toBe('1');
    expect(dv[1].formula2).toBe('120');
    expect(dv[1].showErrorMessage).toBe(true);
    expect(dv[1].errorTitle).toBe('无效');
    expect(dv[1].error).toBe('请输入1~120');

    expect(dv[2].operator).toBe('greaterThan');
    expect(dv[2].formula1).toBe('0');
    expect(dv[3].operator).toBe('greaterThan');
    expect(dv[3].formula1).toBe('43831'); // 2020-01-01 的序列值
    expect(dv[4].operator).toBe('lessThanOrEqual');
    expect(dv[4].formula1).toBe('5');
  });
});

/* ========================================================================== */
/* fixture-extras.xlsx：批注 / 超链接 / 图片                                    */
/* ========================================================================== */

describe('fixture-extras.xlsx：超链接 / 批注 / 图片', () => {
  it('超链接 / 批注：外链取 rels 的 Target、内部链接去掉 #、批注保留换行与作者', async () => {
    const wb = await loadFixture('fixture-extras.xlsx');
    const sheet = sheetOf(wb, '批注链接图片');

    expect(sheet.hyperlinks).toEqual([
      { ref: 'B1', target: 'https://example.com/path?q=1' },
      { ref: 'B2', location: '条件格式!A1' },
      { ref: 'B3', target: 'mailto:test@example.com' },
    ]);

    const notes = sheet.notes ?? [];
    expect(notes.map((n) => n.ref)).toEqual(['A1', 'A2']);
    expect(notes[0].text).toBe('这是一条批注（legacy note）\n第二行');
    expect(notes[0].text.split('\n')).toHaveLength(2);
    expect(notes[1].text).toBe('第一行\n第二行\n第三行');
    expect(notes[1].text.split('\n')).toHaveLength(3);
    expect(notes[0].author).toBe('Author');
    expect(notes[1].author).toBe('Author');

    // 2 张图片：oneCell/twoCell 锚点、EMU 偏移、媒体路径已在 zip 内归一化且真的存在
    const images = sheetOf(wb, '批注链接图片').images ?? [];
    expect(images).toHaveLength(2);

    expect(images[0].mediaPath).toBe('xl/media/image1.png');
    expect(images[0].anchorType).toBe('oneCell');
    expect(images[0].from).toEqual({ col: 2, row: 4, colOffEmu: 40000, rowOffEmu: 36000 });
    expect(images[0].extEmu).toEqual({ cx: 1219200, cy: 762000 });
    expect(images[0].to).toBeUndefined();

    expect(images[1].mediaPath).toBe('xl/media/image2.png');
    expect(images[1].anchorType).toBe('twoCell');
    expect(images[1].from).toEqual({ col: 2, row: 9, colOffEmu: 40000, rowOffEmu: 35999 });
    expect(images[1].to).toEqual({ col: 3, row: 11, colOffEmu: 255999, rowOffEmu: 72000 });

    // rels 的 ../media/... 已归一化 -> 媒体路径在 zip 里真的存在
    for (const image of images) {
      expect(wb.raw.entries[image.mediaPath], `${image.mediaPath} 不在 zip 条目里`).toBeInstanceOf(Uint8Array);
    }
  });
});

/* ========================================================================== */
/* fixture-table.xlsx：Excel 表格                                               */
/* ========================================================================== */

describe('fixture-table.xlsx：Excel 表格', () => {
  it('2 个表格顺序与 <tableParts> 一致：SalesTable（Medium9、无汇总行、4 列）', async () => {
    const wb = await loadFixture('fixture-table.xlsx');
    const tables = sheetOf(wb, '表格').tables ?? [];
    expect(tables).toHaveLength(2);
    expect(tables.map((t) => t.name)).toEqual(['SalesTable', 'TotalsTable']);

    const table = tables[0];
    expect(table.displayName).toBe('SalesTable');
    expect(table.ref).toBe('A1:D5');
    expect(table.headerRowCount).toBe(1);
    expect(table.totalsRowCount).toBe(0); // 无 totalsRowCount 属性，且 ref 里没有汇总行
    expect(table.styleName).toBe('TableStyleMedium9');
    expect(table.showRowStripes).toBe(true);
    expect(table.showColumnStripes).toBe(false);
    expect(table.showFirstColumn).toBe(false);
    expect(table.showLastColumn).toBe(false);
    expect(table.columns?.map((c) => c.name)).toEqual(['区域', '数量', '金额', '日期']);

    // 原样保留 totalsRowShown="1"（ExcelJS 的反向写法）+ 一条解释性 warning；
    // totalsRowCount 不因此变成 1
    expect(table.totalsRowShown).toBe(true);
    expect(wb.report.warnings.join('\n')).toContain('totalsRowShown="1"');

    // TotalsTable：有汇总行、Light 家族、首列强调
    const totalsTable = (sheetOf(wb, '表格').tables ?? [])[1];
    expect(totalsTable.name).toBe('TotalsTable');
    expect(totalsTable.displayName).toBe('TotalsTable');
    expect(totalsTable.ref).toBe('A8:C11');
    expect(totalsTable.headerRowCount).toBe(1);
    expect(totalsTable.totalsRowCount).toBe(1); // XML 里是 totalsRowCount="1"
    expect(totalsTable.styleName).toBe('TableStyleLight11');
    expect(totalsTable.showRowStripes).toBe(true);
    expect(totalsTable.showFirstColumn).toBe(true);
    expect(totalsTable.showColumnStripes).toBe(false);
    expect(totalsTable.showLastColumn).toBe(false);
    expect(totalsTable.columns?.map((c) => c.name)).toEqual(['项目', 'Q1', 'Q2']);
  });
});

/* ========================================================================== */
/* report.unsupported：已解析的必须移除，未解析的继续按种类计数                    */
/* ========================================================================== */

describe('report.unsupported 记账', () => {
  it('已解析的特性全部移除，未解析的（x14 扩展 / VML 绘图）继续按种类计数', async () => {
    const rules = await loadFixture('fixture-rules.xlsx');
    const joined = rules.report.unsupported.join('\n');
    expect(joined).not.toContain('条件格式(conditionalFormatting)');
    expect(joined).not.toContain('数据验证(dataValidations)');
    // 反向确认：确实解析出了内容，而不是被整块丢掉
    expect(sheetOf(rules, '条件格式').conditionalFormats).toHaveLength(4);
    expect(sheetOf(rules, '数据验证').dataValidations).toHaveLength(5);
    /* 打印设置自本轮起**不再算"未支持"**（用户实测：导入座位表时状态栏显示"未支持 1 项"，
       点开却是只影响打印的 pageMargins —— 误导）。现在解析成真值并进 report.preserved。 */
    expect(joined, '打印设置不该再出现在未支持清单里').not.toContain('打印设置');
    expect(sheetOf(rules, '条件格式').print, '打印设置应被解析进模型').toBeTruthy();
    expect(
      (rules.report.preserved ?? []).join('\n'),
      '打印设置应记进"已原样保留"（说明它不影响预览）',
    ).toContain('打印设置已读取并原样保留');
    // 仍未解析的继续按一类计数
    expect(joined).toContain('x14 扩展条件格式(x14:conditionalFormattings) · 1 处（未解析）');
    expect(rules.report.unsupported.every((u) => u.startsWith('['))).toBe(true);

    const extras = await loadFixture('fixture-extras.xlsx');
    const extraJoined = extras.report.unsupported.join('\n');
    expect(extraJoined).not.toContain('超链接(hyperlink)');
    expect(extraJoined).not.toContain('批注(comments)');
    expect(extraJoined).not.toContain('绘图/图形(drawing)');
    expect(extraJoined).not.toContain('图片(picture)');
    expect(sheetOf(extras, '批注链接图片').images).toHaveLength(2);
    expect(extraJoined).toContain('旧式 VML 绘图(legacyDrawing)');

    const table = await loadFixture('fixture-table.xlsx');
    expect(table.report.unsupported.join('\n')).not.toContain('表格对象(tableParts)');
    expect(sheetOf(table, '表格').tables).toHaveLength(2);
  });
});

/* ========================================================================== */
/* 启用宏的工作簿（.xlsm）：能打开、宏原样保留、不执行                            */
/* ========================================================================== */

describe('fixture-macro.xlsm：启用宏的工作簿', () => {
  it('解析器按 OPC 关系找部件（不看扩展名）：.xlsm 与 .xlsx 走同一条路径', async () => {
    const wb = await loadFixture('fixture-macro.xlsm');

    expect(wb.sheets.map((sheet) => sheet.name)).toEqual(['宏工作簿']);
    const sheet = sheetOf(wb, '宏工作簿');
    expect(sheet.cells.find((c) => c.row === 0 && c.col === 0)?.value).toContain('启用宏的工作簿');
    expect(sheet.cells.find((c) => c.row === 2 && c.col === 1)?.formula, '公式照样解析').toContain('A3+1');
    // 原始 zip 条目必须完整保留（导出要靠它逐字节回写，宏部件就在里面）
    expect(wb.raw.entries['xl/vbaProject.bin'], '宏部件要在原始条目里').toBeTruthy();
    expect(wb.raw.entries['[Content_Types].xml'], '内容类型部件要在').toBeTruthy();
  });

  it('宏记在"已原样保留"里，而不是"未支持"（用户会以为宏丢了）', async () => {
    const wb = await loadFixture('fixture-macro.xlsm');
    const unsupported = wb.report.unsupported.join('\n');
    const preserved = (wb.report.preserved ?? []).join('\n');

    expect(unsupported, '宏不该出现在"未支持"清单里').not.toContain('宏(VBA)');
    expect(preserved, '应说明宏已原样保留').toContain('宏(VBA)');
    expect(preserved, '还要说明不会执行它').toContain('不执行宏');
  });
});

/* ========================================================================== */
/* 打印设置：从"未支持"改成"已解析并原样保留"（用户实测反馈）                      */
/* ========================================================================== */

describe('打印设置（pageSetup / pageMargins / headerFooter / 分页符）', () => {
  it('解析出真值：纸张/方向/缩放/页边距/页眉页脚/分页符，并给出一行"不影响预览"的说明', () => {
    const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <printOptions horizontalCentered="1"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="2"/>
  <headerFooter><oddHeader>&amp;C&amp;"宋体"座位表</oddHeader><oddFooter>&amp;P/&amp;N</oddFooter></headerFooter>
  <rowBreaks count="1" manualBreakCount="1"><brk id="20" max="16383" man="1"/></rowBreaks>
</worksheet>`;
    const { print, summary } = parsePrintSettings(xml);

    expect(print).toBeTruthy();
    expect(print?.paperSize).toBe(9);
    expect(print?.orientation).toBe('landscape');
    expect(print?.fitToWidth).toBe(1);
    expect(print?.fitToHeight).toBe(2);
    expect(print?.margins).toEqual({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
    expect(print?.headerFooter?.oddHeader).toContain('座位表');
    expect(print?.headerFooter?.oddFooter).toBe('&P/&N');
    expect(print?.breaks).toEqual({ row: 1, col: 0 });
    expect(print?.centered).toBe(true);

    expect(summary).toContain('打印设置已读取并原样保留');
    expect(summary, '纸型/方向/页边距都要说清楚').toContain('纸张 A4');
    expect(summary).toContain('横向');
    expect(summary).toContain('页边距 左0.7/右0.7/上0.75/下0.75');
    expect(summary, '必须点明"不影响预览"，否则用户会以为是缺陷').toContain('不影响预览');
  });

  it('只有空占位（如 <headerFooter/>）时不编默认值，说明写成"空占位"', () => {
    const xml = `<worksheet><sheetData/><headerFooter/></worksheet>`;
    const { print, summary } = parsePrintSettings(xml);
    expect(print).toEqual({});
    expect(summary).toContain('空');
    expect(summary).not.toContain('纸张');
  });

  it('完全没有打印相关标签 → 不进 preserved（不打扰用户）', () => {
    const { print, summary } = parsePrintSettings('<worksheet><sheetData/></worksheet>');
    expect(print).toBeUndefined();
    expect(summary).toBeUndefined();
  });
});

/* ========================================================================== */
/* 填充降级记账：只有**真被引用**的填充才值得报警                                 */
/* ========================================================================== */

describe('填充(paternFill)降级记账', () => {
  const stylesWith = (fills: string, xfs: string): string => `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3">${fills}</fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${xfs.split('<xf ').length - 1}">${xfs}</cellXfs>
</styleSheet>`;

  it('Excel/WPS 的 gray125 占位填充没人引用时**不报警**（用户实测：座位表因此误报"降级 1 项"）', () => {
    const warnings: string[] = [];
    const styles = parseStyles(
      stylesWith(
        '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill>',
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>',
      ),
      { warn: (m) => warnings.push(m) },
    );
    expect(styles).toHaveLength(2);
    expect(styles[1].fill).toBe('#FFCC00');
    expect(warnings.join('\n'), '没被引用的 gray125 不该产生任何降级警告').not.toContain('gray125');
    expect(warnings).toHaveLength(0);
  });

  it('真被单元格引用的非纯色填充才报警，且说清"按无填充显示"', () => {
    const warnings: string[] = [];
    const styles = parseStyles(
      stylesWith(
        '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="lightUp"/></fill>',
        '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>',
      ),
      { warn: (m) => warnings.push(m) },
    );
    expect(styles[0].fill).toBeUndefined();
    expect(warnings.join('\n')).toContain('lightUp');
    expect(warnings.join('\n')).toContain('按无填充显示');
    expect(warnings.join('\n'), '没被引用的 gray125 仍然不报').not.toContain('gray125');
  });
});

/* ========================================================================== */
/* 手写 XML 边界用例（fixtures 里没有的形态）                                    */
/* ========================================================================== */

describe('sqref 展开', () => {
  it('空格分隔的多区域都能展开（含整列 / 整行引用）；非法片段丢弃并记 warning', () => {
    expect(splitSqref('A1:B2 D4:E5')).toEqual(['A1:B2', 'D4:E5']);
    expect(expandSqref('A1:B2 D4:E5')).toEqual(['A1:B2', 'D4:E5']);
    expect(expandSqref('  A2  ')).toEqual(['A2']);
    expect(expandSqref('A:A 1:1')).toEqual(['A:A', '1:1']);
    expect(isA1RangeToken('A:A')).toBe(true);
    expect(isA1RangeToken('A:2')).toBe(false);

    const warnings: string[] = [];
    expect(expandSqref('A1:B2 不是区域', (m) => warnings.push(m), 'dataValidation')).toEqual(['A1:B2']);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('dataValidation');
  });
});

describe('条件格式：多容器 + 多区域 sqref + x14 扩展', () => {
  const XML = `<worksheet><sheetData/>
    <conditionalFormatting sqref="A1:B2 D4:E5"><cfRule type="expression" dxfId="2"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>
    <conditionalFormatting sqref="C1:C9"><cfRule type="containsText" operator="containsText" text="abc" priority="7"><formula>NOT(ISERROR(SEARCH("abc",C1)))</formula></cfRule></conditionalFormatting>
  </worksheet>`;

  it('多个 <conditionalFormatting> 与 multi-range sqref 都展开', () => {
    const rules = parseConditionalFormatting(XML);
    expect(rules).toHaveLength(2);
    expect(rules[0].ranges).toEqual(['A1:B2', 'D4:E5']);
    expect(rules[0].kind).toBe('highlight');
    expect(rules[0].ruleType).toBe('expression');
    expect(rules[0].formula1).toBe('MOD(ROW(),2)=0');
    expect(rules[1].ranges).toEqual(['C1:C9']);
    expect(rules[1].text).toBe('abc');
    expect(rules[1].priority).toBe(7);
  });

  it('缺少 priority 时按出现顺序兜底并记 warning', () => {
    const warnings: string[] = [];
    const rules = parseConditionalFormatting(XML, { warn: (m) => warnings.push(m) });
    expect(rules[0].priority).toBe(1); // 没有 priority 属性 -> 按顺序 1
    expect(warnings.join('\n')).toContain('缺少 priority');
  });

  it('x14 扩展条件格式（无 sqref 属性）被跳过并留 warning', () => {
    const warnings: string[] = [];
    const xml = '<worksheet><extLst><x14:conditionalFormattings xmlns:x14="http://x">'
      + '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{X}"/><xm:sqref>D2:D6</xm:sqref>'
      + '</x14:conditionalFormatting></x14:conditionalFormattings></extLst></worksheet>';
    expect(parseConditionalFormatting(xml, { warn: (m) => warnings.push(m) })).toEqual([]);
    expect(warnings.join('\n')).toContain('x14 扩展条件格式');
  });
});

describe('数据验证：多区域 sqref 与 x14 扩展', () => {
  it('sqref="A1:A5 C1:C5" 展开成两个区域', () => {
    const xml = '<worksheet><dataValidations count="1">'
      + '<dataValidation type="list" allowBlank="1" sqref="A1:A5 C1:C5"><formula1>"x,y"</formula1></dataValidation>'
      + '</dataValidations></worksheet>';
    const rules = parseDataValidations(xml);
    expect(rules).toHaveLength(1);
    expect(rules[0].ranges).toEqual(['A1:A5', 'C1:C5']);
    expect(rules[0].type).toBe('list');
    expect(rules[0].allowBlank).toBe(true);
    expect(rules[0].formula1).toBe('"x,y"');
  });

  it('x14 扩展数据验证（区域写在 <xm:sqref> 子元素）不会被误当成老格式', () => {
    const xml = '<worksheet><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"a"</formula1>'
      + '</dataValidation><extLst><x14:dataValidations xmlns:x14="http://x" count="1">'
      + '<x14:dataValidation type="list"><x14:formula1><xm:f>"b"</xm:f></x14:formula1><xm:sqref>B1</xm:sqref>'
      + '</x14:dataValidation></x14:dataValidations></extLst></dataValidations></worksheet>';
    const rules = parseDataValidations(xml);
    expect(rules).toHaveLength(1);
    expect(rules[0].ranges).toEqual(['A1']);
  });
});

describe('dxf 差异样式：solid 底色以 bgColor 为准', () => {
  it('同时写了 fgColor 与 bgColor 时取 bgColor（与普通 fill 相反），空的 <dxf/> 也占下标', () => {
    const xml = '<styleSheet><dxfs count="2">'
      + '<dxf><font><i/><strike/><color rgb="FF112233"/></font>'
      + '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor rgb="FF00FF00"/></patternFill></fill>'
      + '<border><left style="thin"><color rgb="FF0000FF"/></left></border></dxf>'
      + '<dxf/>'
      + '</dxfs></styleSheet>';
    const dxfs = parseDxfStyles(xml, { warn: () => {} });
    expect(dxfs).toHaveLength(2); // 空的 <dxf/> 也占一个下标，保证 dxfId 对齐
    expect(dxfs[0]).toEqual({
      italic: true,
      strikeThrough: true,
      color: '#112233',
      fill: '#00FF00',
      border: { left: { style: 'thin', color: '#0000FF' } },
    });
    expect(dxfs[1]).toEqual({});
  });
});

describe('批注：<t> 直写与实体换行', () => {
  it('直接 <t>（无 run）与 &#10; 都要正确还原，run 之间直接拼接不插换行', () => {
    const xml = '<comments><authors><author>甲</author><author>乙</author></authors><commentList>'
      + '<comment ref="B2" authorId="1"><text><t>第一行&#10;第二行</t></text></comment>'
      + '<comment ref="C3" authorId="0"><text><r><rPr><b/></rPr><t>a</t></r><r><t xml:space="preserve"> b</t></r></text></comment>'
      + '</commentList></comments>';
    const notes = parseComments(xml);
    expect(notes).toHaveLength(2);
    expect(notes[0].ref).toBe('B2');
    expect(notes[0].text).toBe('第一行\n第二行');
    expect(notes[0].author).toBe('乙');
    expect(notes[1].text).toBe('a b');
    expect(notes[1].author).toBe('甲');

    // 越界 authorId 只记 warning，不抛异常
    const warnings: string[] = [];
    const badXml = '<comments><authors><author>甲</author></authors><commentList>'
      + '<comment ref="A1" authorId="9"><text><t>x</t></text></comment></commentList></comments>';
    const badNotes = parseComments(badXml, { warn: (m) => warnings.push(m) });
    expect(badNotes[0].author).toBeUndefined();
    expect(badNotes[0].text).toBe('x');
    expect(warnings.join('\n')).toContain('越界的 authorId');
  });
});

describe('绘图：absoluteAnchor 与非图片对象', () => {
  const rels = '<Relationships><Relationship Id="rId1" Type="http://x/image" Target="../media/image1.png"/></Relationships>';

  it('absoluteAnchor 用 pos + ext（图片仍能拿到媒体路径）；图表/形状按种类计数，媒体关系缺失时只记 warning', () => {
    const xml = '<xdr:wsDr xmlns:xdr="http://x" xmlns:a="http://a">'
      + '<xdr:absoluteAnchor><xdr:pos x="914400" y="457200"/><xdr:ext cx="609600" cy="304800"/>'
      + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7" name="P"/></xdr:nvPicPr>'
      + '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic><xdr:clientData/>'
      + '</xdr:absoluteAnchor></xdr:wsDr>';
    const result = parseDrawing(xml, { relsXml: rels, partPath: 'xl/drawings/drawing3.xml' });
    expect(result.images).toHaveLength(1);
    expect(result.images[0]).toEqual({
      id: '7',
      mediaPath: 'xl/media/image1.png',
      anchorType: 'absolute',
      from: { col: 0, row: 0 },
      extEmu: { cx: 609600, cy: 304800 },
      posEmu: { x: 914400, y: 457200 },
    });

    // 图表/形状按种类计数，图片照常解析
    const mixedXml = '<xdr:wsDr xmlns:xdr="http://x" xmlns:a="http://a" xmlns:c="http://c">'
      + '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>'
      + '<xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>'
      + '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart"/></xdr:nvGraphicFramePr>'
      + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId9"/></a:graphicData></a:graphic>'
      + '</xdr:graphicFrame></xdr:twoCellAnchor>'
      + '<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:row>3</xdr:row></xdr:from><xdr:ext cx="100" cy="200"/>'
      + '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="4" name="TextBox"/></xdr:nvSpPr></xdr:sp><xdr:clientData/>'
      + '</xdr:oneCellAnchor></xdr:wsDr>';
    const mixed = parseDrawing(mixedXml, { relsXml: rels, partPath: 'xl/drawings/drawing1.xml' });
    expect(mixed.images).toEqual([]);
    expect(mixed.unsupported).toEqual(['图表(chart) · 1 处（未解析）', '形状/文本框(shape) · 1 处（未解析）']);

    // 图片的 r:embed 指向不存在的媒体关系 -> 只记 warning，不抛异常
    const warnings: string[] = [];
    const dangling = '<xdr:wsDr xmlns:xdr="http://x" xmlns:a="http://a">'
      + '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:ext cx="1" cy="2"/>'
      + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="P"/></xdr:nvPicPr>'
      + '<xdr:blipFill><a:blip r:embed="rId404"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>';
    expect(parseDrawing(dangling, { warn: (m) => warnings.push(m) }).images).toEqual([]);
    expect(warnings.join('\n')).toContain('rId404');
  });
});

describe('表格：缺省值与容错', () => {
  it('最小表格：headerRowCount 默认 1、totalsRowCount 默认 0（显式属性优先）', () => {
    const table = parseTable('<table name="T" ref="B2:D4"/>');
    expect(table).toEqual({ name: 'T', ref: 'B2:D4', headerRowCount: 1, totalsRowCount: 0 });

    const explicit = parseTable('<table name="T" ref="B2:D4" totalsRowCount="1" headerRowCount="0"/>');
    expect(explicit?.totalsRowCount).toBe(1);
    expect(explicit?.headerRowCount).toBe(0);
  });

  it('缺 ref 时跳过并记 warning', () => {
    const warnings: string[] = [];
    expect(parseTable('<table name="T"/>', { warn: (m) => warnings.push(m), partPath: 'xl/tables/table9.xml' })).toBeUndefined();
    expect(warnings.join('\n')).toContain('table9.xml');
  });
});

/* ========================================================================== */
/* 部件级容错：一个部件坏掉不能影响其它部件                                       */
/* ========================================================================== */

describe('部件级容错', () => {
  /** 工作表本身完全正常；外部部件从"缺失"到"畸形"各有各的坏法 */
  const PARTS: Record<string, string> = {
    '_rels/.rels': '<Relationships><Relationship Id="rId1" Type="http://x/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns:r="http://r"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
      + '<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>'
      + '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>'
      + '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>'
      // 关系号在 rels 里不存在
      + '<tableParts count="2"><tablePart r:id="rId7"/><tablePart r:id="rId8"/></tableParts>'
      + '<drawing r:id="rId5"/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships>'
      + '<Relationship Id="rId1" Type="http://x/hyperlink" Target="https://example.com/a?b=1" TargetMode="External"/>'
      + '<Relationship Id="rId5" Type="http://x/drawing" Target="../drawings/drawing1.xml"/>'
      + '<Relationship Id="rId6" Type="http://x/comments" Target="../comments9.xml"/>'
      + '<Relationship Id="rId8" Type="http://x/table" Target="../tables/table9.xml"/>'
      // rId9 是"rels 里有、<tableParts> 没引用"的表格
      + '<Relationship Id="rId9" Type="http://x/table" Target="../tables/table8.xml"/>'
      + '</Relationships>',
    'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="http://xdr" xmlns:a="http://a">'
      + '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/>'
      + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="P"/></xdr:nvPicPr>'
      + '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>',
    // 畸形 XML：截断的 <comment>，而且 drawing rels 整个不存在
    'xl/comments9.xml': '<comments><authors><author>甲</author></authors><commentList><comment ref="A1"',
  };

  it('坏部件只记 warning：工作表/条件格式/数据验证/超链接照常解析，坏批注不拖垮整表', async () => {
    const wb = await parseXlsx(makeXlsx(PARTS));
    expect(wb.sheets).toHaveLength(1);
    const sheet = wb.sheets[0];
    expect(sheet.cells).toHaveLength(1);
    expect(sheet.conditionalFormats).toHaveLength(1);
    expect(sheet.dataValidations).toHaveLength(1);
    expect(sheet.hyperlinks).toEqual([{ ref: 'A1', target: 'https://example.com/a?b=1' }]);
    // 坏掉的批注/表格/图片只是没有内容，不是整体失败
    expect(sheet.notes).toBeUndefined();
    expect(sheet.tables).toBeUndefined();
    expect(sheet.images).toBeUndefined();

    const warnings = wb.report.warnings.join('\n');
    expect(warnings).toContain('xl/comments9.xml');                        // 畸形 comments（截断的 <comment>）只降级
    expect(warnings).toContain('rId7');                                    // <tableParts> 里找不到的关系
    expect(warnings).toContain('没有出现在 <tableParts> 里');               // rels 里有、tableParts 没引用的表格
    expect(warnings).toContain('表格部件 xl/tables/table9.xml 不存在');
    expect(warnings).toContain('缺少 xl/drawings/_rels/drawing1.xml.rels');
    expect(warnings).toContain('rId1');                                    // 图片的 r:embed 找不到媒体

    // 外部部件全缺时也不抛异常
    const broken: Record<string, string> = {
      ...PARTS,
      'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships>'
        + '<Relationship Id="rId5" Type="http://x/drawing" Target="../drawings/drawing2.xml"/>'
        + '</Relationships>',
    };
    delete broken['xl/drawings/drawing1.xml'];
    delete broken['xl/comments9.xml'];
    const missingWb = await parseXlsx(makeXlsx(broken));
    expect(missingWb.sheets[0].conditionalFormats).toHaveLength(1);
    expect(missingWb.report.warnings.join('\n')).toContain('xl/drawings/drawing2.xml');
  });
});
