/**
 * 条件格式解析：工作表 XML 里的 `<conditionalFormatting>` / `<cfRule>` /
 * `<colorScale>` / `<dataBar>` / `<iconSet>`。
 *
 * 产出契约里的 `ConditionalFormatRule[]`，字段名与 OOXML 一一对应：
 *
 * | OOXML | 契约 |
 * | --- | --- |
 * | `cfRule/@type` | `ruleType`（`colorScale`/`dataBar`/`iconSet` 归类进 `kind`，其余都是 `highlight`） |
 * | `cfRule/@dxfId` | `dxfId`，并把 `styles.xml` 的 `<dxfs>` 对应项放进 `dxf` |
 * | `cfRule/formula` | `formula1` / `formula2`（按出现顺序） |
 * | `colorScale/cfvo` `dataBar/cfvo` `iconSet/cfvo` | `cfvo` |
 * | `colorScale/color` | `colors`（从低到高） |
 * | `dataBar/color` | `color` |
 * | `iconSet/@iconSet` | `iconSet` |
 *
 * 关键实现细节：
 * - 一个工作表可以有**多个** `<conditionalFormatting>`，同一个 `sqref` 里还可以有多个区域
 *   （空格分隔，如 `A1:B2 D4:E5`），两者都展开进 `ranges`。
 * - x14 扩展（`<extLst>` 里的 `<x14:conditionalFormattings>`）**不是**这条老路径：
 *   它没有 `sqref` 属性（用子元素 `<xm:sqref>`），会被识别出来并跳过 + 记 warning，
 *   真正的记账在 worksheet.ts 的 unsupported 扫描里。
 * - 任何单个规则解析失败都只记 warning，不影响其它规则。
 */
import type { CfDxfStyle, CfRuleKind, CfValueObject, ConditionalFormatRule } from './types';
import type { Warn } from './worksheet';
import { expandSqref } from './worksheet';
import { parseColorElement, type ColorContext } from './styles';
import {
  attrBool, attrInt, childElements, elementText, localName, nameMatches, visitElements,
  type ElementRange,
} from './xml';

export interface ParseConditionalFormattingOptions {
  warn?: Warn;
  /** `styles.xml` 的 `<dxfs>`：用于把 `dxfId` 解析成真实样式 */
  dxfStyles?: readonly CfDxfStyle[];
  /** 主题色：cfvo/color 里写 `theme="N"` 时需要（不传则用 Excel 默认主题） */
  themeColors?: readonly string[];
  /**
   * 扫描起点。`<conditionalFormatting>` 一律排在 `<sheetData>` 之后，
   * 从那里起步可以免去再走一遍整张表的数据（百万格时这是主要开销）。
   */
  from?: number;
}

/** cfvo 的 type 白名单（契约里的 8 种） */
const CFVO_TYPES: ReadonlySet<string> = new Set([
  'min', 'max', 'num', 'percent', 'percentile', 'formula', 'autoMin', 'autoMax',
]);

/** 解析工作表的全部条件格式规则（含多容器、多区域） */
export function parseConditionalFormatting(
  xml: string,
  opts: ParseConditionalFormattingOptions = {},
): ConditionalFormatRule[] {
  const warn: Warn = opts.warn ?? (() => {});
  const colorCtx: ColorContext = { themeColors: opts.themeColors, warn };

  const containers: ElementRange[] = [];
  visitElements(xml, (el) => {
    containers.push(el);
    return false; // 子树交给 childElements 处理
  }, { match: (name) => localName(name) === 'conditionalFormatting', from: opts.from ?? 0 });

  const out: ConditionalFormatRule[] = [];
  let order = 0;
  let x14Skipped = 0;

  for (const container of containers) {
    const sqref = container.attrs['sqref'];
    if (sqref === undefined || sqref === '') {
      // x14 扩展条件格式：区域写在子元素 <xm:sqref> 里，不走这里
      x14Skipped++;
      continue;
    }
    const ranges = expandSqref(sqref, warn, 'conditionalFormatting');
    if (ranges.length === 0) {
      warn(`conditionalFormatting 的 sqref="${sqref}" 没有可用的区域，整块规则已跳过`);
      continue;
    }
    for (const ruleEl of childElements(xml, container)) {
      if (localName(ruleEl.name) !== 'cfRule') continue;
      order++;
      try {
        const rule = buildRule(xml, ruleEl, ranges, order, warn, colorCtx);
        attachDxf(rule, opts.dxfStyles, warn);
        out.push(rule);
      } catch (err) {
        warn(`conditionalFormatting 的一条 cfRule 解析失败，已跳过：${errText(err)}`);
      }
    }
  }

  if (x14Skipped > 0) {
    warn(`有 ${x14Skipped} 块 x14 扩展条件格式（无 sqref 属性）暂不解析，已跳过`);
  }
  return out;
}

function attachDxf(
  rule: ConditionalFormatRule,
  dxfStyles: readonly CfDxfStyle[] | undefined,
  warn: Warn,
): void {
  if (rule.dxfId === undefined || !dxfStyles) return;
  const dxf = dxfStyles[rule.dxfId];
  if (dxf) rule.dxf = dxf;
  else warn(`cfRule 引用了不存在的 dxfId=${rule.dxfId}（dxfs 共 ${dxfStyles.length} 项），已忽略差异样式`);
}

function buildRule(
  xml: string,
  el: ElementRange,
  ranges: string[],
  order: number,
  warn: Warn,
  colorCtx: ColorContext,
): ConditionalFormatRule {
  const type = el.attrs['type'] ?? '';
  const priorityAttr = attrInt(el.attrs, 'priority');
  if (priorityAttr === undefined) {
    warn(`cfRule(type="${type}") 缺少 priority，已按出现顺序取 ${order}`);
  }
  const rule: ConditionalFormatRule = {
    kind: kindOf(type),
    ranges,
    priority: priorityAttr ?? order,
  };
  if (type !== '') rule.ruleType = type;
  const operator = el.attrs['operator'];
  if (operator !== undefined && operator !== '') rule.operator = operator;
  const text = el.attrs['text'];
  if (text !== undefined) rule.text = text;
  const dxfId = attrInt(el.attrs, 'dxfId');
  if (dxfId !== undefined) rule.dxfId = dxfId;

  const formulas: string[] = [];
  for (const kid of childElements(xml, el)) {
    const ln = localName(kid.name);
    if (ln === 'formula') {
      formulas.push(elementText(xml, kid));
    } else if (ln === 'colorScale') {
      // <colorScale><cfvo/><cfvo/><color/><color/></colorScale>
      rule.cfvo = readCfvoList(xml, kid, warn, 'colorScale');
      rule.colors = readColors(xml, kid, colorCtx, warn, 'colorScale');
    } else if (ln === 'dataBar') {
      rule.cfvo = readCfvoList(xml, kid, warn, 'dataBar');
      const color = firstColor(xml, kid, colorCtx, warn, 'dataBar');
      if (color !== undefined) rule.color = color;
      const showValue = attrBool(kid.attrs, 'showValue');
      if (showValue !== undefined) rule.showValue = showValue;
    } else if (ln === 'iconSet') {
      rule.iconSet = kid.attrs['iconSet'] ?? '3TrafficLights1'; // 规范默认值
      rule.cfvo = readCfvoList(xml, kid, warn, 'iconSet');
      const showValue = attrBool(kid.attrs, 'showValue');
      if (showValue !== undefined) rule.showValue = showValue;
    }
  }
  if (formulas.length > 0) rule.formula1 = formulas[0];
  if (formulas.length > 1) rule.formula2 = formulas[1];
  return rule;
}

function kindOf(type: string): CfRuleKind {
  if (type === 'colorScale') return 'colorScale';
  if (type === 'dataBar') return 'dataBar';
  if (type === 'iconSet') return 'iconSet';
  return 'highlight';
}

/** `<cfvo type="min" val="0"/>` -> `{ type: 'min' }`（`val` 缺失时不写 `value` 字段） */
function readCfvoList(xml: string, parent: ElementRange, warn: Warn, owner: string): CfValueObject[] {
  const out: CfValueObject[] = [];
  for (const kid of childElements(xml, parent)) {
    if (!nameMatches(kid.name, 'cfvo')) continue;
    const rawType = kid.attrs['type'] ?? 'num';
    let type: CfValueObject['type'] = 'num';
    if (CFVO_TYPES.has(rawType)) {
      type = rawType as CfValueObject['type'];
    } else {
      warn(`${owner} 的 cfvo type="${rawType}" 无法识别，已按 num 处理`);
    }
    const item: CfValueObject = { type };
    const val = kid.attrs['val'];
    if (val !== undefined && val !== '') item.value = val;
    out.push(item);
  }
  return out;
}

function readColors(
  xml: string,
  parent: ElementRange,
  colorCtx: ColorContext,
  warn: Warn,
  owner: string,
): string[] {
  const out: string[] = [];
  for (const kid of childElements(xml, parent)) {
    if (!nameMatches(kid.name, 'color')) continue;
    const hex = parseColorElement(kid.attrs, colorCtx);
    if (hex !== undefined) out.push(hex);
    else warn(`${owner} 的第 ${out.length + 1} 个颜色缺少可用色值，已跳过`);
  }
  return out;
}

function firstColor(
  xml: string,
  parent: ElementRange,
  colorCtx: ColorContext,
  warn: Warn,
  owner: string,
): string | undefined {
  for (const kid of childElements(xml, parent)) {
    if (!nameMatches(kid.name, 'color')) continue;
    const hex = parseColorElement(kid.attrs, colorCtx);
    if (hex !== undefined) return hex;
    warn(`${owner} 的颜色缺少可用色值`);
    return undefined;
  }
  return undefined;
}

function errText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
