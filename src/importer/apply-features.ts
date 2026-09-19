/**
 * 把解析出来的"非单元格"特性应用到 Univer 工作表上：
 * 条件格式、数据验证、超链接、批注、浮动图片。
 *
 * 调用时机：`createWorkbook` 之后、安装"仅内容可编辑"锁**之前**。
 * 每一项独立容错：单条规则失败只记入 `failed`，不影响其它规则与整份文件。
 *
 * 实现约定（全部按 `node_modules/@univerjs/*` 的 facade `.d.ts` 与运行时核对过）：
 *  - 条件格式：`fWorksheet.newConditionalFormattingRule()` → `when* / setColorScale /
 *    setDataBar / setIconSet` → `setRanges([IRange])` → `build()` → `addConditionalFormattingRule()`
 *  - 数据验证：`univerAPI.newDataValidation()` → `require*`（或先 `build()` 再
 *    `setCriteria()`，textLength/time 没有专用方法）→ `setAllowBlank/setOptions`
 *    → `fWorksheet.getRange(a1).setDataValidation(rule)`
 *  - 超链接：`fWorksheet.getRange(ref).setHyperLink(url, label)`；表内跳转的 url 用
 *    `FRange.getUrl()` 生成（`#gid=<sheetId>&range=<n>`）
 *  - 批注：`fWorksheet.getRange(ref).createOrUpdateNote(ISheetNote)`
 *  - 图片：`fWorksheet.newOverGridImage().setSource(objectUrl, ImageSourceType.URL)...buildAsync()`
 *    → `insertImages([image])`
 *
 * 计数含义：
 *  - `failed`：本条特性尝试应用但失败（Facade 抛错 / 区域非法 / 数值转不出来 / 媒体字节缺失…）
 *  - `skipped`：本条特性在 Univer API 里无法表达，主动放弃（未识别的图标集名、absoluteAnchor…）
 *  - 两者都会往 `issues` 写一条可读原因；"成功但有降级"（如 dxf 边框无法表达）只写 issue 不计数。
 */
import {
  DataValidationOperator,
  DataValidationType,
  ImageSourceType,
  type IDataValidationRuleOptions,
  type IRange,
} from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import type { FRange, FWorksheet } from '@univerjs/sheets/facade';
// 副作用导入：Facade mixin 通过 `declare module` 增强 FWorksheet/FRange/FEnum 的类型，
// 必须让这些模块进入编译单元；运行时的注册由 src/univer/setup.ts 的 preset 完成（同一模块实例）。
import '@univerjs/sheets-conditional-formatting/facade';
import '@univerjs/sheets-data-validation/facade';
import '@univerjs/sheets-drawing/facade';
import '@univerjs/sheets-hyper-link/facade';
import '@univerjs/sheets-note/facade';

import type { FConditionalFormattingBuilder } from '@univerjs/sheets-conditional-formatting/facade';
import type { FDataValidationBuilder } from '@univerjs/sheets-data-validation/facade';
import type { IColorScale, IIconSet, IValueConfig } from '@univerjs/sheets-conditional-formatting';
import type { ISheetNote } from '@univerjs/sheets-note';

import type {
  CfDxfStyle,
  CfValueObject,
  ConditionalFormatRule,
  DataValidationRule,
  ParsedHyperlink,
  ParsedImage,
  ParsedSheet,
  ParsedWorkbook,
} from '../parser/types';

export interface ApplyFeatureCounts {
  conditionalFormats: number;
  dataValidations: number;
  hyperlinks: number;
  notes: number;
  images: number;
  failed: number;
  skipped: number;
}

export interface ApplyFeaturesResult {
  counts: ApplyFeatureCounts;
  /** 失败/跳过的细节，供"降级报告"与排查使用 */
  issues: string[];
}

/* ==========================================================================
 * 常量与工具
 * ======================================================================== */

/** EMU → px：1 英寸 = 914400 EMU = 96px */
const EMU_PER_PX = 9525;
/** Univer 批注的默认尺寸（sheets-note-ui 的 defaultNoteSize 一致） */
const NOTE_WIDTH = 160;
const NOTE_HEIGHT = 72;
/** 数据条默认色（Excel 经典蓝，cfvo/color 缺失时的兜底） */
const DEFAULT_DATA_BAR_COLOR = '#638EC6';

/**
 * 插图用的 blob url：**按工作簿登记**，工作簿被释放（关标签 / 冷存）时统一回收。
 *
 * 为什么**不能立刻** revoke：图片服务是"插图之后再按 url 取图"的异步流程，
 * 马上 revoke 会让图片变空白 —— 所以 url 的生命周期要**跟着工作簿走**，而不是跟着这次插入走。
 *
 * 为什么**也不能**像以前那样"登记进一个 Set 然后页面生命周期内永不回收"（真实内存泄漏）：
 * 每个 url 都会把它背后的图片字节一直钉在内存里。打开若干份带图的表再关掉，这些字节
 * 一个都还不了（更糟的是冷存标签时，工作簿都 dispose 了、图片字节却还留着）。
 * 释放时机是安全的：`disposeUnit` 之后该簿的图片服务不会再取图；冷存标签切回时会重新走
 * `applyWorkbookFeatures` 重新插图、重新建 url（见 `App.tsx` 的 `buildTabUnit`）。
 */
const objectUrlsByWorkbook = new Map<string, Set<string>>();

/** 拿不到 unitId 时的兜底桶（仍会被 `releaseAllImageObjectUrls` 回收） */
const UNKNOWN_WORKBOOK = '__unknown__';

/** 登记一个"要跟着工作簿活着"的 blob url（由 `applyImages` 调用） */
export function retainImageObjectUrl(workbookId: string, url: string): void {
  let urls = objectUrlsByWorkbook.get(workbookId);
  if (!urls) {
    urls = new Set<string>();
    objectUrlsByWorkbook.set(workbookId, urls);
  }
  urls.add(url);
}

/** 释放单个 url（插图失败时可以立刻作废，不必等整簿释放） */
export function releaseImageObjectUrl(workbookId: string, url: string): void {
  const urls = objectUrlsByWorkbook.get(workbookId);
  if (!urls) return;
  if (!urls.delete(url)) return;
  try {
    URL.revokeObjectURL(url);
  } catch {
    /* 释放失败也不影响其它 url */
  }
  if (urls.size === 0) objectUrlsByWorkbook.delete(workbookId);
}

/**
 * 释放某个工作簿的全部插图 blob url，返回释放个数。
 * 关标签、冷存标签（工作簿被 dispose）时调用。
 */
export function releaseWorkbookImageObjectUrls(workbookId: string): number {
  const urls = objectUrlsByWorkbook.get(workbookId);
  if (!urls) return 0;
  objectUrlsByWorkbook.delete(workbookId);
  let released = 0;
  for (const url of urls) {
    try {
      URL.revokeObjectURL(url);
      released += 1;
    } catch {
      /* 单个 url 释放失败不影响其它；它仍会随页面卸载一起消失 */
    }
  }
  return released;
}

/** 兜底：释放全部（整实例拆卸 / HMR / 测试重置）。返回释放个数。 */
export function releaseAllImageObjectUrls(): number {
  let released = 0;
  for (const workbookId of [...objectUrlsByWorkbook.keys()]) released += releaseWorkbookImageObjectUrls(workbookId);
  return released;
}

/** 诊断/测试用：当前还挂着多少个未被回收的插图 blob url */
export function countRetainedImageObjectUrls(): number {
  let total = 0;
  for (const urls of objectUrlsByWorkbook.values()) total += urls.size;
  return total;
}

/** 诊断/测试用：按工作簿看分布 */
export function retainedImageObjectUrlsByWorkbook(): Record<string, number> {
  const result: Record<string, number> = {};
  for (const [workbookId, urls] of objectUrlsByWorkbook) result[workbookId] = urls.size;
  return result;
}

/**
 * Excel OOXML 的 iconSet 名 → Univer `IIconSetType` 名。
 * Univer 的枚举值就是 Excel 的名字，这里显式列出以便复核；未列出的名字一律跳过并记 issue。
 */
const ICON_SET_NAME_MAP: Record<string, string> = {
  '3Arrows': '3Arrows',
  '3ArrowsGray': '3ArrowsGray',
  '3Flags': '3Flags',
  '3TrafficLights1': '3TrafficLights1',
  '3TrafficLights2': '3TrafficLights2',
  '3Signs': '3Signs',
  '3Symbols': '3Symbols',
  '3Symbols2': '3Symbols2',
  '3Triangles': '3Triangles',
  '3Stars': '3Stars',
  '4Arrows': '4Arrows',
  '4ArrowsGray': '4ArrowsGray',
  '4RedToBlack': '4RedToBlack',
  '4Rating': '4Rating',
  '4TrafficLights': '4TrafficLights',
  '5Arrows': '5Arrows',
  '5ArrowsGray': '5ArrowsGray',
  '5Rating': '5Rating',
  '5Quarters': '5Quarters',
  '5Boxes': '5Boxes',
};

/** dxf 的边框在 Univer 条件格式里只有 bg/cl/bl/it/st/ul 可表达，边框只能降级 */
const DXF_BORDER_NOTE = 'dxf 边框在 Univer 条件格式里无法表达，已忽略边框只保留底色/字体';

interface FeatureContext {
  univerAPI: FUniver;
  fWorksheet: FWorksheet;
  sheet: ParsedSheet;
  parsed: ParsedWorkbook;
  counts: ApplyFeatureCounts;
  issues: string[];
}

function fail(ctx: FeatureContext, message: string): void {
  ctx.counts.failed += 1;
  ctx.issues.push(message);
}

function skip(ctx: FeatureContext, message: string): void {
  ctx.counts.skipped += 1;
  ctx.issues.push(message);
}

function errorText(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}

/**
 * 取当前工作表所属工作簿的 unitId（插图 blob url 按它归档）。
 * 取不到时落到兜底桶：宁可"晚一点释放"，也不能在这里抛异常打断导入。
 */
function workbookIdOf(ctx: FeatureContext): string {
  try {
    return ctx.fWorksheet.getWorkbook().getUnitId() ?? UNKNOWN_WORKBOOK;
  } catch {
    return UNKNOWN_WORKBOOK;
  }
}

/** 去掉 OOXML 里的外层双引号（列表字面量、文本比较值都用它包裹） */
function stripQuotes(raw: string): string {
  const text = raw.trim();
  if (text.length >= 2 && text.startsWith('"') && text.endsWith('"')) return text.slice(1, -1);
  return text;
}

/** 去掉前导 '='（Univer 的公式入参要求带 '='，OOXML 里不带） */
function withEquals(raw: string): string {
  const text = raw.trim();
  if (text === '') return '=';
  return text.startsWith('=') ? text : `=${text}`;
}

/** 字符串 → 数字；转不出来返回 null（调用方决定跳过还是兜底） */
function toNumber(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  const text = stripQuotes(raw).replace(/^=/, '');
  if (text === '') return null;
  const value = Number(text);
  return Number.isFinite(value) ? value : null;
}

/** Excel 颜色（#RRGGBB / RRGGBB / FFRRGGBB）统一成 #RRGGBB */
function normalizeColor(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const text = raw.trim();
  if (/^#[0-9a-f]{6}$/i.test(text)) return text.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(text)) return `#${text.toUpperCase()}`;
  if (/^[0-9a-f]{8}$/i.test(text)) return `#${text.slice(2).toUpperCase()}`;
  return text;
}

/** Excel 1900 日期系统的序列号 → Date（1899-12-30 作原点可自动吸收 1900 闰年 bug） */
export function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}

/** 把 `location`（"Sheet1!A1" / "'我 的表'!A1:B2" / "A1"）拆成表名与 A1 */
export function splitLocation(location: string): { sheetName?: string; a1: string } {
  let text = location.trim().replace(/^#/, '').replace(/^=/, '');
  const bang = text.lastIndexOf('!');
  if (bang < 0) return { a1: text };
  let sheetName = text.slice(0, bang).trim();
  text = text.slice(bang + 1).trim();
  if (sheetName.startsWith("'") && sheetName.endsWith("'") && sheetName.length >= 2) {
    sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
  }
  return { sheetName, a1: text.replace(/\$/g, '') };
}

/** Excel 运算符（不区分大小写、忽略下划线）→ Univer 的数据验证运算符 */
function toValidationOperator(raw: string | undefined, fallback: DataValidationOperator): DataValidationOperator | null {
  const table: Record<string, DataValidationOperator> = {
    between: DataValidationOperator.BETWEEN,
    notbetween: DataValidationOperator.NOT_BETWEEN,
    equal: DataValidationOperator.EQUAL,
    notequal: DataValidationOperator.NOT_EQUAL,
    greaterthan: DataValidationOperator.GREATER_THAN,
    greaterthanorequal: DataValidationOperator.GREATER_THAN_OR_EQUAL,
    lessthan: DataValidationOperator.LESS_THAN,
    lessthanorequal: DataValidationOperator.LESS_THAN_OR_EQUAL,
  };
  if (!raw) return fallback;
  const key = raw.replace(/[\s_-]/g, '').toLowerCase();
  return table[key] ?? null;
}

/** 解析区域列表 → IRange[]（非法区域只记一次 issue，不抛） */
function resolveRanges(ctx: FeatureContext, refs: string[]): IRange[] {
  const ranges: IRange[] = [];
  for (const ref of refs) {
    if (!ref) continue;
    try {
      ranges.push(ctx.fWorksheet.getRange(ref).getRange());
    } catch (error) {
      ctx.issues.push(`区域 ${ref} 无法解析：${errorText(error)}`);
    }
  }
  return ranges;
}

/* ==========================================================================
 * 主入口
 * ======================================================================== */

export async function applySheetFeatures(
  univerAPI: FUniver,
  fWorksheet: FWorksheet,
  sheet: ParsedSheet,
  parsed: ParsedWorkbook,
): Promise<ApplyFeaturesResult> {
  const counts: ApplyFeatureCounts = {
    conditionalFormats: 0,
    dataValidations: 0,
    hyperlinks: 0,
    notes: 0,
    images: 0,
    failed: 0,
    skipped: 0,
  };

  const issues: string[] = [];
  const ctx: FeatureContext = { univerAPI, fWorksheet, sheet, parsed, counts, issues };

  applyConditionalFormats(ctx);
  applyDataValidations(ctx);
  await applyHyperlinks(ctx);
  applyNotes(ctx);
  await applyImages(ctx);

  return { counts, issues };
}

/* ==========================================================================
 * 1) 条件格式
 * ======================================================================== */

/** `whenCellEmpty()` 的返回类型（同一个 builder 类的其它 when / set 分支返回值类型相同） */
type HighlightBuilder = ReturnType<FConditionalFormattingBuilder['whenCellEmpty']>;

function applyConditionalFormats(ctx: FeatureContext): void {
  const rules = ctx.sheet.conditionalFormats ?? [];
  rules.forEach((rule, index) => {
    const label = `条件格式#${index + 1}(${rule.kind}${rule.ruleType ? `/${rule.ruleType}` : ''})`;
    try {
      const ranges = resolveRanges(ctx, rule.ranges);
      if (ranges.length === 0) {
        fail(ctx, `${label}：区域「${rule.ranges.join(' ')}」全部无法解析，已跳过`);
        return;
      }
      const applied = buildConditionalFormat(ctx, rule, ranges, label);
      if (applied) ctx.counts.conditionalFormats += 1;
    } catch (error) {
      fail(ctx, `${label} 应用失败：${errorText(error)}`);
    }
  });
}

/** @returns 是否成功落到工作表 */
function buildConditionalFormat(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  ranges: IRange[],
  label: string,
): boolean {
  switch (rule.kind) {
    case 'highlight':
      return applyHighlightRule(ctx, rule, ranges, label);
    case 'colorScale':
      return applyColorScaleRule(ctx, rule, ranges, label);
    case 'dataBar':
      return applyDataBarRule(ctx, rule, ranges, label);
    case 'iconSet':
      return applyIconSetRule(ctx, rule, ranges, label);
    default:
      skip(ctx, `${label}：Univer 条件格式没有对应的规则类型，已跳过`);
      return false;
  }
}

/** dxf 样式：优先用解析层直接给出的 dxf，退化到 parsed.dxfStyles[dxfId] */
function resolveDxf(ctx: FeatureContext, rule: ConditionalFormatRule): CfDxfStyle | undefined {
  if (rule.dxf) return rule.dxf;
  if (rule.dxfId === undefined) return undefined;
  return ctx.parsed.dxfStyles?.[rule.dxfId];
}

/** 按 ruleType + operator 选择 highlight 分支；返回 null 表示已记 issue 且放弃 */
function createHighlightBuilder(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  label: string,
): HighlightBuilder | null {
  const factory = ctx.fWorksheet.newConditionalFormattingRule();
  const numberOperator = ctx.univerAPI.Enum.ConditionFormatNumberOperatorEnum;
  const timeOperator = ctx.univerAPI.Enum.ConditionFormatTimePeriodOperatorEnum;
  const ruleType = (rule.ruleType ?? '').trim().toLowerCase();
  const operator = (rule.operator ?? '').trim();
  const operatorKey = operator.replace(/[\s_-]/g, '').toLowerCase();
  const text = rule.text !== undefined ? rule.text : rule.formula1 !== undefined ? stripQuotes(rule.formula1) : '';

  // cellIs：数值比较。formula1/formula2 是字符串，转不出数字就跳过并记 issue
  if (ruleType === 'cellis') {
    const first = toNumber(rule.formula1);
    const second = toNumber(rule.formula2);
    const needsTwo = operatorKey === 'between' || operatorKey === 'notbetween';
    if (first === null || (needsTwo && second === null)) {
      fail(
        ctx,
        `${label}：cellIs 的阈值「${rule.formula1 ?? ''}${needsTwo ? ` / ${rule.formula2 ?? ''}` : ''}」无法转成数字，已跳过`,
      );
      return null;
    }
    switch (operatorKey) {
      case 'greaterthan':
        return factory.whenNumberGreaterThan(first);
      case 'greaterthanorequal':
        return factory.whenNumberGreaterThanOrEqualTo(first);
      case 'lessthan':
        return factory.whenNumberLessThan(first);
      case 'lessthanorequal':
        return factory.whenNumberLessThanOrEqualTo(first);
      case 'equal':
        return factory.whenNumberEqualTo(first);
      case 'notequal':
        return factory.whenNumberNotEqualTo(first);
      case 'between':
        return factory.whenNumberBetween(first, second ?? first);
      case 'notbetween':
        return factory.whenNumberNotBetween(first, second ?? first);
      default:
        skip(ctx, `${label}：cellIs 运算符「${operator || '(空)'}」没有对应方法，已跳过`);
        return null;
    }
  }

  // 文本类
  if (ruleType === 'containstext') return factory.whenTextContains(text);
  if (ruleType === 'notcontainstext') return factory.whenTextDoesNotContain(text);
  if (ruleType === 'beginswith') return factory.whenTextStartsWith(text);
  if (ruleType === 'endswith') return factory.whenTextEndsWith(text);

  // 公式
  if (ruleType === 'expression') {
    if (!rule.formula1) {
      fail(ctx, `${label}：expression 规则没有 formula1，已跳过`);
      return null;
    }
    return factory.whenFormulaSatisfied(withEquals(rule.formula1));
  }

  // 重复值 / 唯一值 / 空值
  if (ruleType === 'duplicatevalues') return factory.setDuplicateValues();
  if (ruleType === 'uniquevalues') return factory.setUniqueValues();
  if (ruleType === 'containsblanks') return factory.whenCellEmpty();
  if (ruleType === 'notcontainsblanks') return factory.whenCellNotEmpty();

  // 平均值
  if (ruleType === 'aboveaverage') return factory.setAverage(numberOperator.greaterThan);
  if (ruleType === 'belowaverage') return factory.setAverage(numberOperator.lessThan);

  // 时间段（whenDate 的入参是 CFTimePeriodOperator 枚举）
  if (ruleType === 'timeperiod') {
    const member = Object.keys(timeOperator).find(
      (key) => timeOperator[key as keyof typeof timeOperator] === operator,
    ) as keyof typeof timeOperator | undefined;
    if (!member) {
      skip(ctx, `${label}：时间段运算符「${operator || '(空)'}」没有对应枚举值，已跳过`);
      return null;
    }
    return factory.whenDate(timeOperator[member]);
  }

  if (ruleType === 'containserrors' || ruleType === 'notcontainserrors') {
    skip(ctx, `${label}：Univer 没有「包含错误值」的条件格式分支，已跳过`);
    return null;
  }
  if (ruleType === 'top10') {
    skip(ctx, `${label}：top10 的 rank/percent/bottom 参数不在解析模型里，为避免错误还原已跳过`);
    return null;
  }

  skip(ctx, `${label}：条件格式规则类型「${rule.ruleType || '(空)'}」尚未映射到 Facade 方法，已跳过`);
  return null;
}

function applyHighlightRule(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  ranges: IRange[],
  label: string,
): boolean {
  const builder = createHighlightBuilder(ctx, rule, label);
  if (!builder) return false;

  const dxf = resolveDxf(ctx, rule);
  if (dxf) {
    const fill = normalizeColor(dxf.fill);
    const fontColor = normalizeColor(dxf.color);
    // 注意顺序：先 when*/set* 选分支，再挂样式，最后 setRanges/build
    if (fill) builder.setBackground(fill);
    if (fontColor) builder.setFontColor(fontColor);
    if (dxf.bold) builder.setBold(true);
    if (dxf.italic) builder.setItalic(true);
    if (dxf.strikeThrough) builder.setStrikethrough(true);
    if (dxf.border) ctx.issues.push(`${label}：${DXF_BORDER_NOTE}`);
  }

  const built = builder.setRanges(ranges).build();
  ctx.fWorksheet.addConditionalFormattingRule(built);
  return true;
}

/** Excel cfvo → Univer IValueConfig */
function cfvoToValueConfig(
  ctx: FeatureContext,
  cfvo: CfValueObject | undefined,
  index: number,
  total: number,
): IValueConfig {
  const valueType = ctx.univerAPI.Enum.ConditionFormatValueTypeEnum;
  if (!cfvo) {
    if (index === 0) return { type: valueType.min };
    if (index === total - 1) return { type: valueType.max };
    return { type: valueType.percentile, value: 50 };
  }
  switch (cfvo.type) {
    case 'min':
    case 'autoMin':
      return { type: valueType.min };
    case 'max':
    case 'autoMax':
      return { type: valueType.max };
    case 'num': {
      const value = toNumber(cfvo.value);
      if (value === null) {
        ctx.issues.push(`条件格式阈值 num「${cfvo.value ?? ''}」无法转成数字，已按最小值处理`);
        return { type: valueType.min };
      }
      return { type: valueType.num, value };
    }
    case 'percent': {
      const value = toNumber(cfvo.value);
      if (value === null) ctx.issues.push(`条件格式阈值 percent「${cfvo.value ?? ''}」无法转成数字，已按 0 处理`);
      return { type: valueType.percent, value: value ?? 0 };
    }
    case 'percentile': {
      const value = toNumber(cfvo.value);
      if (value === null) ctx.issues.push(`条件格式阈值 percentile「${cfvo.value ?? ''}」无法转成数字，已按 0 处理`);
      return { type: valueType.percentile, value: value ?? 0 };
    }
    case 'formula':
      return { type: valueType.formula, value: withEquals(cfvo.value ?? '') };
    default:
      ctx.issues.push(`条件格式阈值类型「${cfvo.type}」没有对应枚举值，已按最小值处理`);
      return { type: valueType.min };
  }
}

function applyColorScaleRule(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  ranges: IRange[],
  label: string,
): boolean {
  const colors = (rule.colors ?? []).map((color) => normalizeColor(color)).filter((c): c is string => !!c);
  if (colors.length < 2) {
    skip(ctx, `${label}：色阶至少需要 2 个颜色（解析到 ${colors.length} 个），已跳过`);
    return false;
  }
  if (colors.length > 3) {
    ctx.issues.push(`${label}：解析到 ${colors.length} 个色标，Univer 只演示 2~3 色阶，多余色标已退化`);
  }
  const cfvo = rule.cfvo ?? [];
  const config: IColorScale['config'] = colors.map((color, index) => ({
    index,
    color,
    value: cfvoToValueConfig(ctx, cfvo[index], index, colors.length),
  }));

  const built = ctx.fWorksheet.newConditionalFormattingRule().setColorScale(config).setRanges(ranges).build();
  ctx.fWorksheet.addConditionalFormattingRule(built);
  return true;
}

function applyDataBarRule(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  ranges: IRange[],
  label: string,
): boolean {
  const cfvo = rule.cfvo ?? [];
  const color = normalizeColor(rule.color) ?? DEFAULT_DATA_BAR_COLOR;
  if (!rule.color) ctx.issues.push(`${label}：数据条没有颜色，已用默认色 ${DEFAULT_DATA_BAR_COLOR}`);

  const built = ctx.fWorksheet
    .newConditionalFormattingRule()
    .setDataBar({
      min: cfvoToValueConfig(ctx, cfvo[0], 0, Math.max(cfvo.length, 2)),
      max: cfvoToValueConfig(ctx, cfvo[1], 1, Math.max(cfvo.length, 2)),
      // OOXML 的经典 dataBar 没写渐变信息，Univer 默认非渐变
      isGradient: false,
      positiveColor: color,
      // nativeColor 是"负值条"颜色；Excel 经典数据条只有一个颜色，这里沿用同色
      nativeColor: color,
      isShowValue: rule.showValue !== false,
    })
    .setRanges(ranges)
    .build();
  ctx.fWorksheet.addConditionalFormattingRule(built);
  return true;
}

/** 图标集名字 → Univer 枚举值（用运行时枚举做白名单，映射不到返回 undefined） */
function lookupIconType(ctx: FeatureContext, excelName: string): IIconSet['config'][number]['iconType'] | undefined {
  const mapped = ICON_SET_NAME_MAP[excelName];
  if (!mapped) return undefined;
  const enums = ctx.univerAPI.Enum.ConditionFormatIconSetTypeEnum;
  const key = (Object.keys(enums) as Array<keyof typeof enums>).find((name) => enums[name] === mapped);
  return key ? enums[key] : undefined;
}

function applyIconSetRule(
  ctx: FeatureContext,
  rule: ConditionalFormatRule,
  ranges: IRange[],
  label: string,
): boolean {
  const excelName = (rule.iconSet ?? '').trim();
  const iconType = lookupIconType(ctx, excelName);
  if (!iconType) {
    skip(ctx, `${label}：图标集「${excelName || '(未指定)'}」在 Univer 的图标集枚举里没有对应项，已跳过`);
    return false;
  }

  const cfvo = rule.cfvo ?? [];
  const iconCount = Number(excelName.charAt(0)) || 3;
  const thresholds: CfValueObject[] =
    cfvo.length >= 2
      ? cfvo
      // 阈值缺失时按 Excel 的默认等分百分位兜底（3 档 → 0/33/67，5 档 → 0/20/40/60/80）
      : Array.from({ length: iconCount }, (_, level) => ({
          type: 'percent' as const,
          value: String(Math.round((level * 100) / iconCount)),
        }));
  if (cfvo.length < 2) {
    ctx.issues.push(`${label}：图标集缺少 cfvo 阈值，已按 Excel 默认等分百分位还原`);
  }

  const numberOperator = ctx.univerAPI.Enum.ConditionFormatNumberOperatorEnum;
  // Excel 的图标下标 0 是"最低档"，Univer 的 iconId 0 是图标组里第一个（最高档），所以倒序并翻转 id
  const iconConfigs: IIconSet['config'] = thresholds
    .map((cfvoItem, excelIndex) => ({
      iconType,
      iconId: String(thresholds.length - 1 - excelIndex),
      // OOXML 的 cfvo 默认 gte="1"（大于等于）
      operator: numberOperator.greaterThanOrEqual,
      value: cfvoToValueConfig(ctx, cfvoItem, excelIndex, thresholds.length),
    }))
    .reverse();

  const built = ctx.fWorksheet
    .newConditionalFormattingRule()
    .setIconSet({ iconConfigs, isShowValue: rule.showValue !== false })
    .setRanges(ranges)
    .build();
  ctx.fWorksheet.addConditionalFormattingRule(built);
  return true;
}

/* ==========================================================================
 * 2) 数据验证
 * ======================================================================== */

/** 把 criteria 写进 builder 并产出可提交的规则（textLength/time 走 build 后的 setCriteria） */
type DataValidationPlan = (builder: FDataValidationBuilder) => ReturnType<FDataValidationBuilder['build']>;

function applyDataValidations(ctx: FeatureContext): void {
  const rules = ctx.sheet.dataValidations ?? [];
  rules.forEach((rule, index) => {
    const label = `数据验证#${index + 1}(${rule.type || 'none'})`;
    try {
      if (!rule.ranges || rule.ranges.length === 0) {
        fail(ctx, `${label}：没有区域，已跳过`);
        return;
      }
      const plan = planDataValidation(ctx, rule, label);
      if (!plan) return;

      for (const ref of rule.ranges) {
        // 每条区域单独 build：Univer 用 uid 索引规则，复用同一个 FDataValidation 会让后一条覆盖前一条
        const validation = plan(ctx.univerAPI.newDataValidation());
        validation.setOptions(validationOptions(rule));
        ctx.fWorksheet.getRange(ref).setDataValidation(validation);
      }
      ctx.counts.dataValidations += 1;
    } catch (error) {
      fail(ctx, `${label} 应用失败：${errorText(error)}`);
    }
  });
}

function validationOptions(rule: DataValidationRule): Partial<IDataValidationRuleOptions> {
  const options: Partial<IDataValidationRuleOptions> = {};
  if (rule.showErrorMessage !== undefined) options.showErrorMessage = rule.showErrorMessage;
  if (rule.error !== undefined) options.error = rule.error;
  if (rule.errorTitle !== undefined) options.errorTitle = rule.errorTitle;
  if (rule.showInputMessage !== undefined) options.showInputMessage = rule.showInputMessage;
  if (rule.prompt !== undefined) options.prompt = rule.prompt;
  if (rule.promptTitle !== undefined) options.promptTitle = rule.promptTitle;
  return options;
}

/** 用 setCriteria 兜底：textLength/time 没有专用 require* 方法，先 build 再改 criteria */
function planByCriteria(
  type: DataValidationType,
  rule: DataValidationRule,
  operator: DataValidationOperator,
  formula1: string,
  formula2: string,
): DataValidationPlan {
  return (builder) => {
    const validation = builder.build();
    validation.setCriteria(type, [operator, formula1, formula2], rule.allowBlank ?? false);
    return validation;
  };
}

function planDataValidation(ctx: FeatureContext, rule: DataValidationRule, label: string): DataValidationPlan | null {
  const type = (rule.type ?? '').trim();
  const allowBlank = rule.allowBlank ?? false;

  if (type === 'list') {
    const raw = rule.formula1;
    if (!raw) {
      fail(ctx, `${label}：列表验证没有 formula1，已跳过`);
      return null;
    }
    const text = raw.trim();
    if (text.startsWith('"')) {
      const values = stripQuotes(text)
        .split(',')
        .map((item) => item.trim())
        .filter((item) => item.length > 0);
      if (values.length === 0) {
        fail(ctx, `${label}：列表字面量为空，已跳过`);
        return null;
      }
      // OOXML 的 showDropDown 语义是**取反**的：showDropDown="1" 表示**隐藏**下拉箭头，
      // 所以传给 Univer 的 showDropdown 要取 `!(rule.showDropDown === true)`。
      const showDropdown = rule.showDropDown !== true;
      return (builder) => {
        builder.setAllowBlank(allowBlank);
        return builder.requireValueInList(values, false, showDropdown).build();
      };
    }

    // 区域引用形式（如 $A$1:$A$5 / 基础资料!$A$1:$A$5）
    const range = resolveValidationRange(ctx, text, label);
    if (!range) return null;
    const showDropdown = rule.showDropDown !== true;
    return (builder) => {
      builder.setAllowBlank(allowBlank);
      return builder.requireValueInRange(range, false, showDropdown).build();
    };
  }

  if (type === 'whole' || type === 'decimal') {
    const isInteger = type === 'whole';
    const first = toNumber(rule.formula1);
    const second = toNumber(rule.formula2);
    const operator = toValidationOperator(rule.operator, DataValidationOperator.BETWEEN);
    if (!operator) {
      skip(ctx, `${label}：运算符「${rule.operator ?? '(空)'}」没有对应枚举值，已跳过`);
      return null;
    }
    const needsTwo = operator === DataValidationOperator.BETWEEN || operator === DataValidationOperator.NOT_BETWEEN;
    if (first === null || (needsTwo && second === null)) {
      fail(ctx, `${label}：阈值「${rule.formula1 ?? ''}${needsTwo ? ` / ${rule.formula2 ?? ''}` : ''}」无法转成数字，已跳过`);
      return null;
    }
    const high = second as number;
    return (builder) => {
      builder.setAllowBlank(allowBlank);
      switch (operator) {
        case DataValidationOperator.BETWEEN:
          return builder.requireNumberBetween(first, high, isInteger).build();
        case DataValidationOperator.NOT_BETWEEN:
          return builder.requireNumberNotBetween(first, high, isInteger).build();
        case DataValidationOperator.EQUAL:
          return builder.requireNumberEqualTo(first, isInteger).build();
        case DataValidationOperator.NOT_EQUAL:
          return builder.requireNumberNotEqualTo(first, isInteger).build();
        case DataValidationOperator.GREATER_THAN:
          return builder.requireNumberGreaterThan(first, isInteger).build();
        case DataValidationOperator.GREATER_THAN_OR_EQUAL:
          return builder.requireNumberGreaterThanOrEqualTo(first, isInteger).build();
        case DataValidationOperator.LESS_THAN:
          return builder.requireNumberLessThan(first, isInteger).build();
        default:
          return builder.requireNumberLessThanOrEqualTo(first, isInteger).build();
      }
    };
  }

  if (type === 'date') {
    const firstSerial = toNumber(rule.formula1);
    const secondSerial = toNumber(rule.formula2);
    const operator = toValidationOperator(rule.operator, DataValidationOperator.BETWEEN);
    if (!operator || firstSerial === null) {
      fail(ctx, `${label}：日期序列号「${rule.formula1 ?? ''}」无法解析，已跳过`);
      return null;
    }
    const start = excelSerialToDate(firstSerial);
    const needsTwo = operator === DataValidationOperator.BETWEEN || operator === DataValidationOperator.NOT_BETWEEN;
    if (needsTwo && secondSerial === null) {
      fail(ctx, `${label}：日期区间缺少 formula2，已跳过`);
      return null;
    }
    const end = excelSerialToDate(secondSerial ?? firstSerial);
    if (operator === DataValidationOperator.NOT_EQUAL) {
      // Univer 没有 requireDateNotEqualTo，用 setCriteria 兜底
      return planByCriteria(DataValidationType.DATE, rule, operator, start.toLocaleDateString(), '');
    }
    return (builder) => {
      builder.setAllowBlank(allowBlank);
      switch (operator) {
        case DataValidationOperator.BETWEEN:
          return builder.requireDateBetween(start, end).build();
        case DataValidationOperator.NOT_BETWEEN:
          return builder.requireDateNotBetween(start, end).build();
        case DataValidationOperator.EQUAL:
          return builder.requireDateEqualTo(start).build();
        case DataValidationOperator.GREATER_THAN:
          return builder.requireDateAfter(start).build();
        case DataValidationOperator.GREATER_THAN_OR_EQUAL:
          return builder.requireDateOnOrAfter(start).build();
        case DataValidationOperator.LESS_THAN:
          return builder.requireDateBefore(start).build();
        default:
          return builder.requireDateOnOrBefore(start).build();
      }
    };
  }

  if (type === 'textLength') {
    const operator = toValidationOperator(rule.operator, DataValidationOperator.BETWEEN);
    if (!operator) {
      skip(ctx, `${label}：运算符「${rule.operator ?? '(空)'}」没有对应枚举值，已跳过`);
      return null;
    }
    const needsTwo = operator === DataValidationOperator.BETWEEN || operator === DataValidationOperator.NOT_BETWEEN;
    if (!rule.formula1 || (needsTwo && !rule.formula2)) {
      fail(ctx, `${label}：文本长度缺少 formula1/formula2，已跳过`);
      return null;
    }
    // textLength 没有专用 require* 方法，用 setCriteria(DataValidationType.TEXT_LENGTH)
    return planByCriteria(
      DataValidationType.TEXT_LENGTH,
      rule,
      operator,
      rule.formula1,
      needsTwo ? (rule.formula2 ?? '') : '',
    );
  }

  if (type === 'time') {
    const operator = toValidationOperator(rule.operator, DataValidationOperator.BETWEEN);
    if (!operator || !rule.formula1) {
      skip(ctx, `${label}：时间验证缺少可用运算符/公式，已跳过`);
      return null;
    }
    return planByCriteria(DataValidationType.TIME, rule, operator, rule.formula1, rule.formula2 ?? '');
  }

  if (type === 'custom') {
    if (!rule.formula1) {
      fail(ctx, `${label}：自定义验证没有 formula1，已跳过`);
      return null;
    }
    const formula = withEquals(rule.formula1);
    return (builder) => {
      builder.setAllowBlank(allowBlank);
      return builder.requireFormulaSatisfied(formula).build();
    };
  }

  skip(ctx, `${label}：Univer 数据验证没有类型「${type || '(空)'}」，已跳过`);
  return null;
}

/** 列表验证的"区域引用"形式 */
function resolveValidationRange(ctx: FeatureContext, ref: string, label: string): FRange | null {
  const { sheetName, a1 } = splitLocation(ref);
  const sheet = sheetName ? ctx.univerAPI.getActiveWorkbook()?.getSheetByName(sheetName) : ctx.fWorksheet;
  if (!sheet || !a1) {
    fail(ctx, `${label}：列表来源「${ref}」无法解析成区域，已跳过`);
    return null;
  }
  try {
    return sheet.getRange(a1);
  } catch (error) {
    fail(ctx, `${label}：列表来源「${ref}」解析失败：${errorText(error)}`);
    return null;
  }
}

/* ==========================================================================
 * 3) 超链接
 * ======================================================================== */

async function applyHyperlinks(ctx: FeatureContext): Promise<void> {
  const links = ctx.sheet.hyperlinks ?? [];
  let tooltipLost = 0;

  for (const [index, link] of links.entries()) {
    const label = `超链接#${index + 1}(${link.ref})`;
    try {
      const url = resolveHyperlinkUrl(ctx, link, label);
      if (!url) continue;
      const ok = await ctx.fWorksheet.getRange(link.ref).setHyperLink(url, link.display ?? undefined);
      if (!ok) {
        fail(ctx, `${label}：setHyperLink 返回 false（url=${url}）`);
        continue;
      }
      ctx.counts.hyperlinks += 1;
      if (link.tooltip) tooltipLost += 1;
    } catch (error) {
      fail(ctx, `${label} 应用失败：${errorText(error)}`);
    }
  }

  if (tooltipLost > 0) {
    ctx.issues.push(`超链接的屏幕提示(tooltip)出现 ${tooltipLost} 次，setHyperLink 没有对应参数，已忽略`);
  }
}

/** 外部 URL 优先，其次表内 location（`#Sheet1!A1`）→ 用 FRange.getUrl() 生成内链 url */
function resolveHyperlinkUrl(ctx: FeatureContext, link: ParsedHyperlink, label: string): string | null {
  const target = link.target?.trim();
  if (target && !target.startsWith('#')) return target;

  const location = link.location?.trim() || (target?.startsWith('#') ? target.slice(1) : '');
  if (!location) {
    fail(ctx, `${label}：既没有外部 target 也没有表内 location，已跳过`);
    return null;
  }
  const { sheetName, a1 } = splitLocation(location);
  const sheet = sheetName ? ctx.univerAPI.getActiveWorkbook()?.getSheetByName(sheetName) : ctx.fWorksheet;
  if (!sheet) {
    // 表内跳转指向了本工作簿里不存在的工作表（样本与真实文件里都会出现这种失效引用）：
    // 仍然把 location 原样作为内链负载写进去，保证"链接没丢"，同时记一条降级说明。
    ctx.issues.push(
      `${label}：表内跳转目标工作表「${sheetName ?? ''}」不存在，已保留原始 location 作为链接负载（点击时 Univer 会提示引用无效）`,
    );
    return `#${location}`;
  }
  try {
    return sheet.getRange(a1 || 'A1').getUrl();
  } catch (error) {
    fail(ctx, `${label}：表内跳转「${location}」解析失败：${errorText(error)}`);
    return null;
  }
}

/* ==========================================================================
 * 4) 批注（legacy note）
 * ======================================================================== */

function applyNotes(ctx: FeatureContext): void {
  const notes = ctx.sheet.notes ?? [];
  notes.forEach((note, index) => {
    const label = `批注#${index + 1}(${note.ref})`;
    try {
      const range = ctx.fWorksheet.getRange(note.ref);
      const rect = range.getRange();
      const sheetNote: ISheetNote = {
        id: `xlsx-note-${ctx.sheet.id}-${rect.startRow}-${rect.startColumn}`,
        row: rect.startRow,
        col: rect.startColumn,
        width: NOTE_WIDTH,
        height: NOTE_HEIGHT,
        // 文本原样保留（含 \n），由批注 UI 负责换行渲染
        note: note.text,
        show: false,
      };
      range.createOrUpdateNote(sheetNote);
      if (!range.getNote()) {
        fail(ctx, `${label}：批注写入后读回为空`);
        return;
      }
      ctx.counts.notes += 1;
      // 作者信息 OOXML 里有、Facade 没有对应字段：按约定忽略，不记 issue
    } catch (error) {
      fail(ctx, `${label} 应用失败：${errorText(error)}`);
    }
  });
}

/* ==========================================================================
 * 5) 浮动图片
 * ======================================================================== */

async function applyImages(ctx: FeatureContext): Promise<void> {
  const images = ctx.sheet.images ?? [];
  // 这一批图片的 url 全部登记到**所在工作簿**名下，关标签/冷存时整簿回收
  const workbookId = workbookIdOf(ctx);

  for (const [index, image] of images.entries()) {
    const label = `图片#${index + 1}(${image.mediaPath})`;
    let objectUrl: string | null = null;
    try {
      if (image.anchorType === 'absolute') {
        skip(ctx, `${label}：absoluteAnchor 是绝对 EMU 定位，Facade 只能按单元格锚点插入，已跳过`);
        continue;
      }
      const bytes = readMediaBytes(ctx.parsed, image.mediaPath);
      if (!bytes) {
        fail(ctx, `${label}：zip 条目里找不到该媒体字节，已跳过`);
        continue;
      }
      // Blob 需要 ArrayBuffer（而不是可能共享的 ArrayBufferLike），这里拷一份保证类型与语义都干净
      const buffer = new ArrayBuffer(bytes.byteLength);
      new Uint8Array(buffer).set(bytes);
      objectUrl = URL.createObjectURL(new Blob([buffer], { type: mimeFromPath(image.mediaPath) }));
      // 不 revoke：图片服务在渲染阶段才去取这个 url；但要登记到工作簿名下，随簿释放
      retainImageObjectUrl(workbookId, objectUrl);

      const builder = ctx.fWorksheet
        .newOverGridImage()
        .setSource(objectUrl, ImageSourceType.URL)
        .setColumn(image.from.col)
        .setRow(image.from.row)
        .setColumnOffset(emuToPx(image.from.colOffEmu))
        .setRowOffset(emuToPx(image.from.rowOffEmu))
        .setAnchorType(anchorTypeOf(ctx, image));

      const size = imageSizePx(ctx, image);
      if (size) builder.setWidth(size.width).setHeight(size.height);

      const sheetImage = await builder.buildAsync();
      ctx.fWorksheet.insertImages([sheetImage]);
      ctx.counts.images += 1;
    } catch (error) {
      // 插入失败 → **只**作废这一张图的 url（不能整簿释放，那会把同一簿里已插好的图弄成空白）
      if (objectUrl) releaseImageObjectUrl(workbookId, objectUrl);
      fail(ctx, `${label} 插入失败：${errorText(error)}`);
    }
  }
}

function emuToPx(emu: number | undefined): number {
  if (!emu || !Number.isFinite(emu)) return 0;
  return Math.round(emu / EMU_PER_PX);
}

function anchorTypeOf(
  ctx: FeatureContext,
  image: ParsedImage,
): FUniver['Enum']['SheetDrawingAnchorType'][keyof FUniver['Enum']['SheetDrawingAnchorType']] {
  const anchor = ctx.univerAPI.Enum.SheetDrawingAnchorType;
  // oneCell：只跟随位置（Excel editAs="oneCell"）；twoCell：位置与尺寸都跟随
  return image.anchorType === 'twoCell' ? anchor.Both : anchor.Position;
}

/** oneCell 用 from + extEmu；twoCell 用 from/to 的行列尺寸差 */
function imageSizePx(ctx: FeatureContext, image: ParsedImage): { width: number; height: number } | null {
  if (image.anchorType === 'oneCell' && image.extEmu) {
    const width = emuToPx(image.extEmu.cx);
    const height = emuToPx(image.extEmu.cy);
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (!image.to) return null;

  try {
    const fromColOff = emuToPx(image.from.colOffEmu);
    const fromRowOff = emuToPx(image.from.rowOffEmu);
    const toColOff = emuToPx(image.to.colOffEmu);
    const toRowOff = emuToPx(image.to.rowOffEmu);

    let width = toColOff - fromColOff;
    for (let col = image.from.col; col <= image.to.col; col++) width += ctx.fWorksheet.getColumnWidth(col);
    let height = toRowOff - fromRowOff;
    for (let row = image.from.row; row <= image.to.row; row++) height += ctx.fWorksheet.getRowHeight(row);

    if (width <= 0 || height <= 0) return null;
    return { width: Math.round(width), height: Math.round(height) };
  } catch (error) {
    ctx.issues.push(
      `图片#${image.id}：twoCell 锚点的像素尺寸推算失败（${errorText(error)}），已交给图片服务按原图尺寸插入`,
    );
    return null;
  }
}

function readMediaBytes(parsed: ParsedWorkbook, mediaPath: string): Uint8Array | undefined {
  const entries = parsed.raw.entries;
  const direct: Uint8Array | undefined = entries[mediaPath];
  if (direct) return direct;
  const trimmed: Uint8Array | undefined = entries[mediaPath.replace(/^\/+/, '')];
  return trimmed ?? entries[`/${mediaPath}`];
}

function mimeFromPath(mediaPath: string): string {
  const ext = mediaPath.slice(mediaPath.lastIndexOf('.') + 1).toLowerCase();
  if (ext === 'png') return 'image/png';
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'bmp') return 'image/bmp';
  if (ext === 'webp') return 'image/webp';
  if (ext === 'svg') return 'image/svg+xml';
  return 'application/octet-stream';
}

/* ==========================================================================
 * 工作簿级编排
 * ======================================================================== */

/** 对整份工作簿逐表应用特性并汇总（单表失败不影响其它表） */
export async function applyWorkbookFeatures(api: FUniver, parsed: ParsedWorkbook): Promise<ApplyFeaturesResult> {
  const counts: ApplyFeatureCounts = {
    conditionalFormats: 0,
    dataValidations: 0,
    hyperlinks: 0,
    notes: 0,
    images: 0,
    failed: 0,
    skipped: 0,
  };
  const issues: string[] = [];

  const workbook = api.getActiveWorkbook();
  if (!workbook) return { counts, issues: ['没有活动工作簿'] };

  for (const sheet of parsed.sheets) {
    const fSheet = workbook.getSheetBySheetId(sheet.id) ?? workbook.getSheetByName(sheet.name);
    if (!fSheet) {
      counts.skipped += 1;
      issues.push(`找不到工作表「${sheet.name}」，已跳过其条件格式/数据验证等特性`);
      continue;
    }
    try {
      const result = await applySheetFeatures(api, fSheet, sheet, parsed);
      for (const key of Object.keys(counts) as Array<keyof ApplyFeatureCounts>) {
        counts[key] += result.counts[key];
      }
      issues.push(...result.issues.map((issue) => `[${sheet.name}] ${issue}`));
    } catch (error) {
      counts.failed += 1;
      issues.push(`[${sheet.name}] 应用特性时出错：${String(error)}`);
    }
  }

  return { counts, issues };
}
