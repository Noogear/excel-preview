/**
 * **打开任意受支持格式的统一入口**：字节 + 文件名 → 统一变成"一份规范 xlsx + 它的解析结果"。
 *
 * 为什么统一成 xlsx（而不是每种格式各写一条导入管线）：
 *  - 本工具的预览/编辑/工作区/撤销/导出/会话恢复**全都建立在 OOXML 包之上**：
 *    导出是在原始 zip 条目上做外科式补丁，会话存的是原始字节。
 *  - 所以 `.csv` / `.ods` / `.xls` 的导入思路是"**翻译**"：自研解析 → 中性模型
 *    （`WorkbookInput`，见 `src/importer/synth-xlsx.ts`）→ 写出一份规范 xlsx → 走既有全链路。
 *    好处是这几种格式不会缺任何能力（工作区、撤销、搜索、跨表……都一样），
 *    代价是导入时多一次"写包 + 解析"（毫秒级）。
 *  - OOXML 源（`.xlsx/.xlsm/.xltx/.xltm`）**原样使用用户字节**：一个 bit 都不动，
 *    导出才能继续逐字节保真。
 *
 * 诚实边界：`.csv`/`.ods`/`.xls` 没有"原始 OOXML 字节"可回写，所以**导出会另存为 .xlsx**；
 * 这件事会通过 `notes` 进导入摘要，不藏着。
 */
import { parseDelimitedText } from '../parser/csv';
import { parseXlsx } from '../parser';
import type { ParsedWorkbook } from '../parser/types';
import { fileKindOf, kindLabel, type WorkbookFileKind } from './file-kinds';
import { writeWorkbookPackage, type WorkbookInput } from './synth-xlsx';

export interface OpenedWorkbook {
  /** 标签页要持有的"原始字节"：OOXML 源 = 用户文件；转换源 = 我们合成的 xlsx */
  bytes: Uint8Array;
  /** 已经解析好的模型（OOXML 源在一次解析后复用，不重复解析） */
  parsed: ParsedWorkbook;
  kind: WorkbookFileKind;
  /** 是否发生过格式转换（决定导出文件名与摘要措辞） */
  converted: boolean;
  /** 给导入摘要的说明（不是错误） */
  notes: string[];
}

/** 各转换格式的解析器：返回中性模型 + 说明 */
type Converter = (bytes: Uint8Array, fileName: string) => { input: WorkbookInput; notes: string[] };

/**
 * 把分隔文本（CSV/TSV/TXT）转成中性模型。
 *
 * `encoding` / `delimiter` 都写进 notes —— 中文 Windows 的 Excel 导出的是 GBK，
 * 用户看到"识别为 GBK"才知道我们读对了（而不是碰巧没乱码）。
 */
function convertDelimited(bytes: Uint8Array, fileName: string): { input: WorkbookInput; notes: string[] } {
  const result = parseDelimitedText(bytes, fileName);
  const notes = [
    `由 ${kindLabel('csv')} 转换导入：编码 ${result.encoding}、分隔符「${result.delimiter === '\t' ? '制表符' : result.delimiter}」、` +
      `共 ${result.rows} 行 × ${result.cols} 列`,
    // CSV 只有内容：这件事必须说清楚，否则用户会以为"格式丢了是我们弄坏的"
    'CSV 本身不含字体/边框/合并等格式信息，已按默认样式呈现；导出会另存为 .xlsx（保留原来的 .csv 文件不动）',
    ...result.notes,
  ];
  return { input: result.input, notes };
}

export interface OpenWorkbookOptions {
  /** 注入转换器便于测试（生产路径不传，走内置实现） */
  converters?: Partial<Record<'csv' | 'ods' | 'xls', Converter>>;
}

/**
 * 打开入口。**调用方必须先用 `isSupportedWorkbookFile()` 判过**（这里对不支持的格式会抛错，
 * 免得"打不开"被静默吞掉）。
 */
export async function openWorkbookBytes(
  bytes: Uint8Array,
  fileName: string,
  options: OpenWorkbookOptions = {},
): Promise<OpenedWorkbook> {
  const kind = fileKindOf(fileName);

  if (kind === 'ooxml') {
    return { bytes, parsed: await parseXlsx(bytes), kind, converted: false, notes: [] };
  }

  const converter =
    options.converters?.[kind as 'csv' | 'ods' | 'xls'] ??
    (kind === 'csv' ? convertDelimited : kind === 'ods' ? convertOds : kind === 'xls' ? convertXls : undefined);
  if (!converter) {
    throw new Error(`不支持的格式：${fileName}（${kindLabel(kind)}）`);
  }

  const { input, notes } = await converter(bytes, fileName);  const convertedBytes = writeWorkbookPackage(input);
  const parsed = await parseXlsx(convertedBytes);
  return { bytes: convertedBytes, parsed, kind, converted: true, notes };
}

/* -------------------------------------------------------------------------- */
/* 各格式转换器（动态引入：让主包首屏不必背上三个解析器）                          */
/* -------------------------------------------------------------------------- */

/** 解析器可能同时提供"详细版"（带 warnings/stats）；这里两种形状都接受 */
interface DetailedLike {
  input: WorkbookInput;
  warnings?: string[];
}

function normalizeDetailed(value: unknown): DetailedLike {
  if (value && typeof value === 'object' && 'input' in (value as Record<string, unknown>)) {
    return value as DetailedLike;
  }
  return { input: value as WorkbookInput, warnings: [] };
}

/** `.ods`（OpenDocument 表格） */
async function convertOds(bytes: Uint8Array): Promise<{ input: WorkbookInput; notes: string[] }> {
  const mod = (await import('../parser/ods')) as unknown as {
    parseOds: (data: Uint8Array) => unknown;
    parseOdsResult?: (data: Uint8Array) => unknown;
    parseOdsDetailed?: (data: Uint8Array) => unknown;
  };
  const detailed = mod.parseOdsResult ?? mod.parseOdsDetailed;
  const result = normalizeDetailed(detailed ? detailed(bytes) : mod.parseOds(bytes));
  return {
    input: result.input,
    notes: [
      `由 ${kindLabel('ods')} 转换导入：${result.input.sheets.length} 张表`,
      '导出会另存为 .xlsx（.ods 无法被外科式回写，原文件不动）',
      ...(result.warnings ?? []),
    ],
  };
}

/** `.xls`（Excel 97–2003 / BIFF8） */
async function convertXls(bytes: Uint8Array): Promise<{ input: WorkbookInput; notes: string[] }> {
  const mod = (await import('../parser/xls')) as unknown as {
    parseXls: (data: Uint8Array) => unknown;
    parseXlsDetailed?: (data: Uint8Array) => unknown;
  };
  const result = normalizeDetailed(mod.parseXlsDetailed ? mod.parseXlsDetailed(bytes) : mod.parseXls(bytes));
  return {
    input: result.input,
    notes: [
      `由 ${kindLabel('xls')} 转换导入：${result.input.sheets.length} 张表（BIFF8 二进制）`,
      '导出会另存为 .xlsx（.xls 是二进制格式，无法在外科式回写路径上保留原字节，原文件不动）',
      ...(result.warnings ?? []),
    ],
  };
}
