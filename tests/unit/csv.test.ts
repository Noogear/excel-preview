/**
 * CSV / TSV / 分隔文本解析（用户要求："把其他表格格式都加上去"）。
 *
 * 覆盖四类真实情况：
 *  ① **编码**：真 Excel 导出的 GBK 样本（中文 Windows 默认）、带 BOM 的 UTF-8、UTF-16；
 *  ② **分隔符**：逗号 / 分号 / 制表符 / 竖线（欧洲区域把逗号当小数点，于是用分号）；
 *  ③ **RFC 4180 引号**：字段内的分隔符、引号、换行；
 *  ④ **类型判定**：纯数字转数值，但 `007`/`0912` 这类带前导零的必须留作文本。
 *
 * 夹具 `fixtures/fixture-*-gbk.csv` / `-utf8.csv` 由 `npm run fixtures:legacy`（真 Excel COM 转换）
 * 生成；缺失时相关用例自动跳过（与既有约定一致）。
 */
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import {
  decodeDelimited,
  delimiterLabel,
  parseDelimitedRows,
  parseDelimitedScalar,
  parseDelimitedText,
  sheetNameFromFile,
  sniffDelimiter,
} from '../../src/parser/csv';
import { parseXlsx } from '../../src/parser';
import { writeWorkbookPackage } from '../../src/importer/synth-xlsx';

const FIXTURE_DIR = join(process.cwd(), 'fixtures');
const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('编码识别', () => {
  it('无 BOM 的 UTF-8 严格解码成功 → utf-8', () => {
    expect(decodeDelimited(utf8('姓名,分数\n张三,90'))).toEqual({ text: '姓名,分数\n张三,90', encoding: 'utf-8' });
  });

  it('UTF-8 BOM 会去掉后再解，并标明 utf-8-bom', () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...utf8('姓名')]);
    expect(decodeDelimited(bytes)).toEqual({ text: '姓名', encoding: 'utf-8-bom' });
  });

  it('UTF-16 LE/BE（带 BOM）都能解', () => {
    // "SY\r" 的 UTF-16LE：53 00 59 00 0D 00
    const le = new Uint8Array([0xff, 0xfe, 0x53, 0x00, 0x59, 0x00, 0x0d, 0x00]);
    expect(decodeDelimited(le).encoding).toBe('utf-16le');
    expect(decodeDelimited(le).text).toBe('SY\r');

    // 同样的内容用 UTF-16BE：00 53 00 59 00 0D
    const be = new Uint8Array([0xfe, 0xff, 0x00, 0x53, 0x00, 0x59, 0x00, 0x0d]);
    expect(decodeDelimited(be).encoding).toBe('utf-16be');
    expect(decodeDelimited(be).text).toBe('SY\r');
  });

  it('GBK 字节：UTF-8 严格解码失败 → 退回 GBK 并正确还原中文', () => {
    // "姓名" 的 GBK 编码是 D0 D5 C3 FB
    const gbk = new Uint8Array([0xd0, 0xd5, 0xc3, 0xfb, 0x2c, 0x39, 0x30]);
    const decoded = decodeDelimited(gbk);
    expect(decoded.encoding).toBe('gbk');
    expect(decoded.text).toBe('姓名,90');
  });
});

describe('分隔符嗅探', () => {
  it('.tsv 一律用制表符（哪怕内容里逗号更多）', () => {
    expect(sniffDelimiter('a,b\tc\n1,2\t3', 'x.tsv')).toBe('\t');
  });

  it('逗号 / 分号 / 制表符 / 竖线都能认出来', () => {
    expect(sniffDelimiter('a,b,c\n1,2,3')).toBe(',');
    expect(sniffDelimiter('a;b;c\n1;2;3')).toBe(';');
    expect(sniffDelimiter('a\tb\tc\n1\t2\t3')).toBe('\t');
    expect(sniffDelimiter('a|b|c\n1|2|3')).toBe('|');
  });

  it('引号里的分隔符不算数（否则备注列里的逗号会把制表符表格带偏）', () => {
    expect(sniffDelimiter('"甲,乙,丙"\t备注\n"1,2,3"\t说明')).toBe('\t');
  });

  it('每行一致优先于出现次数：偶尔多几个逗号也不会改判', () => {
    // 第 1 行 1 个分号，第 2 行 1 个分号；逗号只在第 3 行出现 3 次
    expect(sniffDelimiter('a;b\n1;2\nx,y,z')).toBe(';');
  });

  it('全是单列时退回逗号（不报错）', () => {
    expect(sniffDelimiter('甲\n乙\n丙')).toBe(',');
  });

  it('delimiterLabel 给中文名（导入摘要里要说人话）', () => {
    expect(delimiterLabel(',')).toBe('逗号');
    expect(delimiterLabel('\t')).toBe('制表符');
  });
});

describe('RFC 4180 解析', () => {
  it('基本切分与 CRLF/LF 混合', () => {
    expect(parseDelimitedRows('a,b\r\nc,d\ne,f', ',')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
      ['e', 'f'],
    ]);
  });

  it('引号内的分隔符、换行、双引号都保留', () => {
    const rows = parseDelimitedRows('"甲,乙","第一行\n第二行","他说""你好"""', ',');
    expect(rows).toEqual([['甲,乙', '第一行\n第二行', '他说"你好"']]);
  });

  it('字段中间的引号按字面量（与 Excel 行为一致）', () => {
    expect(parseDelimitedRows('a"b,c', ',')).toEqual([['a"b', 'c']]);
  });

  it('结尾没有换行也能收下最后一行', () => {
    expect(parseDelimitedRows('a,b', ',')).toEqual([['a', 'b']]);
    expect(parseDelimitedRows('', ',')).toEqual([]);
  });

  it('空字段保留（列对齐靠它）', () => {
    expect(parseDelimitedRows('a,,c', ',')).toEqual([['a', '', 'c']]);
  });
});

describe('类型判定（保守）', () => {
  it('纯数字转数值，包括负数/小数/科学计数法', () => {
    expect(parseDelimitedScalar('42')).toBe(42);
    expect(parseDelimitedScalar('-3.5')).toBe(-3.5);
    expect(parseDelimitedScalar('1e3')).toBe(1000);
    expect(parseDelimitedScalar(' 7 ')).toBe(7);
  });

  it('带前导零的必须留作文本（学号 / 电话 / 编号）', () => {
    expect(parseDelimitedScalar('007')).toBe('007');
    expect(parseDelimitedScalar('0912')).toBe('0912');
    expect(parseDelimitedScalar('-007')).toBe('-007');
    // 0.5 不是"前导零"
    expect(parseDelimitedScalar('0.5')).toBe(0.5);
    expect(parseDelimitedScalar('0')).toBe(0);
  });

  it('日期/百分比/布尔/千分位等一律按文本（不替用户猜）', () => {
    expect(parseDelimitedScalar('2026-09-07')).toBe('2026-09-07');
    expect(parseDelimitedScalar('95%')).toBe('95%');
    expect(parseDelimitedScalar('TRUE')).toBe('TRUE');
    expect(parseDelimitedScalar('1,234')).toBe('1,234');
  });
});

describe('parseDelimitedText：端到端（含真 Excel 导出的夹具）', () => {
  it('小样本：跳过空字段、去掉尾部空行、工作表名取自文件名', async () => {
    const result = parseDelimitedText(utf8('姓名,分数,备注\n张三,90,\n,,\n'), '成绩单.csv');
    expect(result.delimiter).toBe(',');
    expect(result.encoding).toBe('utf-8');
    expect(result.input.sheets).toHaveLength(1);
    expect(result.input.sheets[0].name).toBe('成绩单');
    expect(result.rows).toBe(2);
    expect(result.cols).toBe(3);

    // 能被打成 xlsx 再解析回来（这就是生产链路）
    const parsed = await parseXlsx(writeWorkbookPackage(result.input));
    const cells = parsed.sheets[0].cells;
    expect(cells.find((c) => c.row === 0 && c.col === 0)?.value).toBe('姓名');
    expect(cells.find((c) => c.row === 1 && c.col === 0)?.value).toBe('张三');
    expect(cells.find((c) => c.row === 1 && c.col === 1)?.value).toBe(90);
    expect(cells.find((c) => c.row === 1 && c.col === 2), '空字段不产生单元格').toBeUndefined();
  });

  it('sheetNameFromFile：去扩展名、空名回落', () => {
    expect(sheetNameFromFile('2502班座位表.csv')).toBe('2502班座位表');
    expect(sheetNameFromFile('a.b.c.TSV')).toBe('a.b.c');
    expect(sheetNameFromFile('.csv')).toBe('数据');
  });

  const GBK_FIXTURE = join(FIXTURE_DIR, 'fixture-styles-gbk.csv');
  const UTF8_FIXTURE = join(FIXTURE_DIR, 'fixture-styles-utf8.csv');

  it('真 Excel 导出的 GBK 样本：识别为 gbk、中文正确、表头对得上', async () => {
    if (!existsSync(GBK_FIXTURE)) {
      expect(true, 'fixtures/fixture-styles-gbk.csv 不存在（先跑 npm run fixtures:legacy），跳过').toBe(true);
      return;
    }
    const bytes = new Uint8Array(readFileSync(GBK_FIXTURE));
    const result = parseDelimitedText(bytes, 'fixture-styles.csv');
    expect(result.encoding, '真 Excel 的 ANSI 导出就是 GBK').toBe('gbk');
    expect(result.notes.join('\n')).toContain('GBK');

    const parsed = await parseXlsx(writeWorkbookPackage(result.input));
    const first = parsed.sheets[0];
    const at = (row: number, col: number) => first.cells.find((c) => c.row === row && c.col === col)?.value;
    expect(at(0, 0), '标题里的中文不能乱码').toBe('样式保真样本（P0 fixture）');
    expect(at(1, 0)).toBe('区域');
    expect(at(1, 1)).toBe('数量');
    expect(at(1, 2)).toBe('单价');
  });

  it('真 Excel 导出的 UTF-8 样本（带 BOM）：同样能读，且两条路径结果一致', async () => {
    if (!existsSync(UTF8_FIXTURE)) {
      expect(true, 'fixtures/fixture-styles-utf8.csv 不存在（先跑 npm run fixtures:legacy），跳过').toBe(true);
      return;
    }
    const utf8Result = parseDelimitedText(new Uint8Array(readFileSync(UTF8_FIXTURE)), 'fixture-styles.csv');
    expect(utf8Result.encoding).toBe('utf-8-bom');

    const parsed = await parseXlsx(writeWorkbookPackage(utf8Result.input));
    const at = (row: number, col: number) => parsed.sheets[0].cells.find((c) => c.row === row && c.col === col)?.value;
    expect(at(0, 0)).toBe('样式保真样本（P0 fixture）');
    expect(at(1, 3)).toBe('生效日期');

    if (existsSync(GBK_FIXTURE)) {
      const gbkResult = parseDelimitedText(new Uint8Array(readFileSync(GBK_FIXTURE)), 'fixture-styles.csv');
      expect(gbkResult.rows, '同一张表的两种编码导出行数应一致').toBe(utf8Result.rows);
      expect(gbkResult.cols).toBe(utf8Result.cols);
      expect(gbkResult.input.sheets[0].cells.length).toBe(utf8Result.input.sheets[0].cells.length);
    }
  });

  it('分隔符变体：分号 + 引号字段 + 前导零编号，一起验证', async () => {
    const result = parseDelimitedText(utf8('学号;姓名;备注\n"007";"张三";"含;分号"\n0912;李四;'), '名单.csv');
    expect(result.delimiter).toBe(';');
    const parsed = await parseXlsx(writeWorkbookPackage(result.input));
    const first = parsed.sheets[0];
    const at = (row: number, col: number) => first.cells.find((c) => c.row === row && c.col === col)?.value;
    expect(at(1, 0)).toBe('007');
    expect(at(1, 2)).toBe('含;分号');
    expect(at(2, 0)).toBe('0912');
  });
});
