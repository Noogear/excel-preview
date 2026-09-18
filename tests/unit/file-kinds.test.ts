/**
 * 「能打开哪些格式」的判定与提示文案（纯函数）。
 *
 * 背景（用户两次提问/要求）：
 *  1. "为什么打开 xlsx 按钮名称为 xlsx，对其他格式的兼容性呢？"
 *  2. "我希望能支持常见表格格式的兼容……请把其他表格格式都加上去。"
 *
 * 现在：**OOXML 全家**（.xlsx/.xlsm/.xltx/.xltm）+ **分隔文本**（.csv/.tsv/.txt）+
 * **.ods** + **.xls** 都能打开；打不开的（.xlsb / Numbers / WPS 私有 / 非表格）要给出
 * "为什么 + 请另存为 .xlsx"的可照做提示。
 */
import { describe, expect, it } from 'vitest';

import {
  ACCEPT_ATTR,
  CONVERTED_EXTENSIONS,
  DELIMITED_EXTENSIONS,
  FORMAT_HINT,
  OOXML_WORKBOOK_EXTENSIONS,
  SUPPORTED_EXTENSIONS,
  exportFileNameFor,
  extensionOf,
  fileKindOf,
  isSupportedWorkbookFile,
  kindLabel,
  needsConversion,
  unsupportedFileMessage,
} from '../../src/importer/file-kinds';

describe('fileKindOf：按扩展名分类', () => {
  it('OOXML 工作簿全家（含启用宏与模板）', () => {
    for (const name of ['a.xlsx', 'a.XLSM', '模板.xltx', '模板.xltm', 'C:\\dir\\座位表.Xlsx']) {
      expect(fileKindOf(name), name).toBe('ooxml');
      expect(isSupportedWorkbookFile(name), name).toBe(true);
      expect(needsConversion(name), name).toBe(false);
    }
  });

  it('分隔文本（CSV/TSV/TXT）能打开，但需要转换', () => {
    for (const name of ['名单.csv', '名单.CSV', 'a.tsv', 'a.txt']) {
      expect(fileKindOf(name), name).toBe('csv');
      expect(isSupportedWorkbookFile(name), name).toBe(true);
      expect(needsConversion(name), name).toBe(true);
    }
  });

  it('.ods（OpenDocument）与 .xls（BIFF8）能打开，但需要转换', () => {
    expect(fileKindOf('a.ods')).toBe('ods');
    expect(fileKindOf('a.ODS')).toBe('ods');
    expect(fileKindOf('旧表.xls')).toBe('xls');
    expect(fileKindOf('a.xlt')).toBe('xls');
    for (const name of ['a.ods', '旧表.xls']) {
      expect(isSupportedWorkbookFile(name), name).toBe(true);
      expect(needsConversion(name), name).toBe(true);
    }
  });

  it('.xlsb 单独一类（二进制 OOXML，要另一套解析器）', () => {
    expect(fileKindOf('a.xlsb')).toBe('xlsb');
    expect(isSupportedWorkbookFile('a.xlsb')).toBe(false);
  });

  it('别家的表格格式一类（Numbers / WPS 私有）', () => {
    for (const name of ['a.numbers', 'a.et', 'a.ett']) {
      expect(fileKindOf(name), name).toBe('foreign');
      expect(isSupportedWorkbookFile(name), name).toBe(false);
    }
  });

  it('其它文件（含没有扩展名、以及把表格后缀当中间段的名字）都落到 other', () => {
    for (const name of ['a.pdf', 'a.png', 'README', 'x.xlsx.exe', '']) {
      expect(fileKindOf(name), name).toBe('other');
      expect(isSupportedWorkbookFile(name), name).toBe(false);
    }
  });
});

describe('extensionOf', () => {
  it('只取最后一段扩展名并转小写；没有扩展名给空串', () => {
    expect(extensionOf('A.B.XLSX')).toBe('.xlsx');
    expect(extensionOf('/tmp/a b c.csv')).toBe('.csv');
    expect(extensionOf('noext')).toBe('');
    expect(extensionOf('')).toBe('');
  });
});

describe('unsupportedFileMessage：给"为什么 + 怎么办"', () => {
  it('.xlsb 说清它是二进制 OOXML，并指路"另存为 .xlsx"', () => {
    const message = unsupportedFileMessage(['成绩册.xlsb']);
    expect(message).toContain('二进制');
    expect(message).toContain('另存为 .xlsx');
    expect(message).toContain('成绩册.xlsb');
  });

  it('Numbers/WPS 私有格式说明需要各自解析器', () => {
    expect(unsupportedFileMessage(['a.numbers'])).toContain('另存为 .xlsx');
    expect(unsupportedFileMessage(['a.et'])).toContain('另存为');
  });

  it('其它类型直接列清支持范围', () => {
    const message = unsupportedFileMessage(['photo.png']);
    for (const ext of SUPPORTED_EXTENSIONS) expect(message, ext).toContain(ext);
    expect(message).toContain('photo.png');
  });

  it('多个文件一次列全（用户可能一次拖进来好几个）', () => {
    const message = unsupportedFileMessage(['a.xlsb', 'b.xlsb']);
    expect(message).toContain('a.xlsb');
    expect(message).toContain('b.xlsb');
  });
});

describe('导出的文件名', () => {
  it('OOXML 源保留原扩展名（改名会让 Excel 认为格式与扩展名不符）', () => {
    expect(exportFileNameFor('座位表.xlsx')).toBe('座位表-已编辑.xlsx');
    expect(exportFileNameFor('座位表.xlsm')).toBe('座位表-已编辑.xlsm');
    expect(exportFileNameFor('模板.XLTX')).toBe('模板-已编辑.xltx');
  });

  it('转换源（csv/ods/xls）导出为 .xlsx（字节本来就是合成的 xlsx）', () => {
    expect(exportFileNameFor('名单.csv')).toBe('名单-已编辑.xlsx');
    expect(exportFileNameFor('座位表.ods')).toBe('座位表-已编辑.xlsx');
    expect(exportFileNameFor('旧表.xls')).toBe('旧表-已编辑.xlsx');
  });

  it('没有扩展名 / 只有扩展名时不产生怪名字', () => {
    expect(exportFileNameFor('workbook')).toBe('workbook-已编辑.xlsx');
    expect(exportFileNameFor('.xlsx')).toBe('workbook-已编辑.xlsx');
  });
});

describe('UI 文案常量', () => {
  it('accept 覆盖全部能打开的扩展名（顺序与常量一致）', () => {
    expect(ACCEPT_ATTR.split(',')).toEqual([...SUPPORTED_EXTENSIONS]);
    for (const ext of OOXML_WORKBOOK_EXTENSIONS) expect(ACCEPT_ATTR, ext).toContain(ext);
    for (const ext of DELIMITED_EXTENSIONS) expect(ACCEPT_ATTR, ext).toContain(ext);
  });

  it('提示里说明"宏不执行、原样保留"以及"转换后会另存为 xlsx"', () => {
    expect(FORMAT_HINT).toContain('.xlsm');
    expect(FORMAT_HINT).toContain('.ods');
    expect(FORMAT_HINT).toContain('.csv');
    expect(FORMAT_HINT).toContain('宏');
  });

  it('kindLabel 给中文名（导入摘要里说人话）', () => {
    expect(kindLabel('ooxml')).toContain('Excel');
    expect(kindLabel('csv')).toContain('CSV');
    expect(kindLabel('ods')).toContain('ods');
    expect(kindLabel('xls')).toContain('97-2003');
  });

  it('CONVERTED_EXTENSIONS 与 needsConversion 一致（别漏一种）', () => {
    for (const ext of CONVERTED_EXTENSIONS) {
      expect(needsConversion(`x${ext}`), ext).toBe(true);
    }
  });
});
