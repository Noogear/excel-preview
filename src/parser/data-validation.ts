/**
 * 数据验证解析：工作表 XML 里的 `<dataValidations>` / `<dataValidation>`。
 *
 * 产出契约里的 `DataValidationRule[]`，属性与子元素一一对应：
 * `type` / `operator` / `sqref`(-> `ranges`) / `allowBlank` / `showDropDown` /
 * `showErrorMessage` / `errorTitle` / `error` / `showInputMessage` / `promptTitle` / `prompt`
 * 以及子元素 `<formula1>` / `<formula2>`。
 *
 * 注意点：
 * - `sqref` 可以是多个空格分隔的区域（`A1:A5 C1:C5`），统一展开成数组。
 * - `operator` 在 XML 里可能**被省略**（规范默认值是 `between`）：实测 fixture-rules.xlsx 的
 *   B2 就是 `type="whole"` 而没有 `operator`，语义上是 between。这里只给"真的用运算符"的
 *   类型（whole/decimal/date/time/textLength）补默认值，list/none/custom 不补。
 * - `showDropDown` 的 OOXML 语义是**反的**：`showDropDown="1"` 表示"隐藏单元格下拉箭头"
 *   （Excel 的"提供下拉箭头"取消勾选时写 1）。这里**原样保留属性值**，把语义留给上层决定，
 *   避免在这里悄悄反转一次再被上层反转回来。
 * - x14 扩展数据验证（`<extLst>` 里的 `<x14:dataValidations>`，区域写在子元素 `<xm:sqref>`）
 *   不是这条路径：它的 `<x14:dataValidation>` 没有 `sqref` 属性，会被自然排除；
 *   记账在 worksheet.ts 的 unsupported 扫描里（`x14 扩展数据验证`）。
 */
import type { DataValidationRule } from './types';
import type { Warn } from './worksheet';
import { expandSqref } from './worksheet';
import { attrBool, childElements, elementText, localName, visitElements, type ElementRange } from './xml';

export interface ParseDataValidationsOptions {
  warn?: Warn;
  /**
   * 扫描起点。`<dataValidations>` 一律排在 `<sheetData>` 之后，
   * 从那里起步可以免去再走一遍整张表的数据（百万格时这是主要开销）。
   */
  from?: number;
}

/** 解析工作表的全部数据验证规则 */
export function parseDataValidations(
  xml: string,
  opts: ParseDataValidationsOptions = {},
): DataValidationRule[] {
  const warn: Warn = opts.warn ?? (() => {});

  const containers: ElementRange[] = [];
  visitElements(xml, (el) => {
    containers.push(el);
    return false;
  }, { match: (name) => localName(name) === 'dataValidations', from: opts.from ?? 0 });

  const out: DataValidationRule[] = [];
  for (const container of containers) {
    for (const el of childElements(xml, container)) {
      if (localName(el.name) !== 'dataValidation') continue;
      const sqref = el.attrs['sqref'];
      // x14 版本用子元素 <xm:sqref>，没有这个属性 —— 直接跳过，不产生噪声
      if (sqref === undefined || sqref === '') continue;
      const ranges = expandSqref(sqref, warn, 'dataValidation');
      if (ranges.length === 0) {
        warn(`dataValidation 的 sqref="${sqref}" 没有可用的区域，已跳过该条规则`);
        continue;
      }
      try {
        out.push(buildRule(xml, el, ranges));
      } catch (err) {
        warn(`一条 dataValidation 解析失败，已跳过：${err instanceof Error ? err.message : String(err)}`);
      }
    }
  }
  return out;
}

function buildRule(xml: string, el: ElementRange, ranges: string[]): DataValidationRule {
  const rule: DataValidationRule = {
    ranges,
    // 规范默认 type="none"（只填 type 属性缺失的脏数据也要有值）
    type: el.attrs['type'] ?? 'none',
  };
  copyString(el, 'operator', (v) => { rule.operator = v; });
  copyString(el, 'errorTitle', (v) => { rule.errorTitle = v; });
  copyString(el, 'error', (v) => { rule.error = v; });
  copyString(el, 'promptTitle', (v) => { rule.promptTitle = v; });
  copyString(el, 'prompt', (v) => { rule.prompt = v; });
  copyBool(el, 'allowBlank', (v) => { rule.allowBlank = v; });
  copyBool(el, 'showDropDown', (v) => { rule.showDropDown = v; });
  copyBool(el, 'showErrorMessage', (v) => { rule.showErrorMessage = v; });
  copyBool(el, 'showInputMessage', (v) => { rule.showInputMessage = v; });
  // 规范里 operator 的默认值是 between。实测 fixture-rules.xlsx 的 B2 写的是
  // `type="whole" allowBlank="1" …`——**没有 operator 属性**（ExcelJS 省略了默认值），
  // 语义上它确实是 between（1~120）。只对"真的用运算符"的类型补默认值：
  // list / none / custom 在 Excel 里没有运算符，硬填 between 反而是假信息。
  if (rule.operator === undefined && OPERATOR_TYPES.has(rule.type)) rule.operator = 'between';

  for (const kid of childElements(xml, el)) {
    const ln = localName(kid.name);
    if (ln === 'formula1') rule.formula1 = elementText(xml, kid);
    else if (ln === 'formula2') rule.formula2 = elementText(xml, kid);
  }
  return rule;
}

/** 只有这些 type 会用到 operator（Excel 的数据验证对话框里也只对这些类型显示运算符） */
const OPERATOR_TYPES: ReadonlySet<string> = new Set([
  'whole', 'decimal', 'date', 'time', 'textLength',
]);

function copyString(el: ElementRange, key: string, assign: (v: string) => void): void {
  const v = el.attrs[key];
  if (v !== undefined) assign(v);
}

function copyBool(el: ElementRange, key: string, assign: (v: boolean) => void): void {
  const v = attrBool(el.attrs, key);
  if (v !== undefined) assign(v);
}
