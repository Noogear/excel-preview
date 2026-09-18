/**
 * Excel 表格（ListObject）解析：`xl/tables/tableN.xml`。
 *
 * 表格部件**独立于工作表编号**（`table1.xml` 不一定属于 `sheet1.xml`），
 * 由 `xl/worksheets/_rels/sheetN.xml.rels` 里 Type 以 `/table` 结尾的关系定位，
 * 顺序以工作表 XML 的 `<tableParts>` 为准（Excel 的表格顺序就是它）。
 *
 * 产出的 `ParsedTable` 只描述"表格是什么、覆盖哪块区域、用什么内置样式名"，
 * **不含具体配色**——`TableStyleMedium9` 这类样式目录不在文件里，
 * 需要上层按样式名"实体化"成单元格样式（见 importer/table-style.ts）。
 */
import type { ParsedTable } from './types';
import type { Warn } from './worksheet';
import {
  attrBool, attrInt, childElements, findFirstElement, localName, type ElementRange, type XmlAttributes,
} from './xml';

export interface ParseTableOptions {
  warn?: Warn;
  /** 表格部件在 zip 里的路径，仅用于 warning 文案 */
  partPath?: string;
}

/** 解析一个 `xl/tables/tableN.xml`；不是表格部件（缺 `<table>`）时返回 undefined */
export function parseTable(xml: string, opts: ParseTableOptions = {}): ParsedTable | undefined {
  const warn: Warn = opts.warn ?? (() => {});
  const where = opts.partPath ?? 'table.xml';
  const tableEl = findFirstElement(xml, 'table');
  if (!tableEl) {
    warn(`${where} 里没有 <table> 根元素，已跳过`);
    return undefined;
  }

  const attrs = tableEl.attrs;
  const ref = attrs['ref'];
  if (ref === undefined || ref === '') {
    warn(`${where} 的 <table> 缺少 ref（表格范围），已跳过`);
    return undefined;
  }
  const name = attrs['name'] ?? attrs['displayName'];
  if (name === undefined || name === '') {
    warn(`${where} 的 <table ref="${ref}"> 缺少 name/displayName`);
  }

  const table: ParsedTable = {
    name: name ?? '',
    ref,
    // 规范默认 headerRowCount=1（有表头）、totalsRowCount=0（无汇总行）
    headerRowCount: attrInt(attrs, 'headerRowCount') ?? 1,
    totalsRowCount: readTotalsRowCount(attrs, warn, name ?? ref),
  };
  const displayName = attrs['displayName'];
  if (displayName !== undefined && displayName !== '') table.displayName = displayName;
  const totalsRowShown = attrBool(attrs, 'totalsRowShown');
  if (totalsRowShown !== undefined) table.totalsRowShown = totalsRowShown;

  for (const kid of childElements(xml, tableEl)) {
    const ln = localName(kid.name);
    if (ln === 'tableColumns') {
      const columns = parseColumns(xml, kid);
      if (columns.length > 0) table.columns = columns;
    } else if (ln === 'tableStyleInfo') {
      applyStyleInfo(table, kid);
    }
  }
  return table;
}

/**
 * `totalsRowCount`（契约字段）取值规则：
 * 1. 有显式 `totalsRowCount` 就用它（Excel 有汇总行时一定写这个属性）；
 * 2. 否则视为 **0**。
 *
 * ⚠️ 为什么不按 `totalsRowShown` 推导：实测 fixture-table.xlsx 是 ExcelJS 写出来的，
 * `exceljs/lib/xlsx/xform/table/table-xform.js` 里
 * `totalsRowCount: model.totalsRow ? '1' : undefined` /
 * `totalsRowShown: model.totalsRow ? undefined : '1'`——**ExcelJS 把"没有汇总行"写成
 * `totalsRowShown="1"`**，字面语义与实际相反。fixture-table.xlsx 的 SalesTable
 * （ref=A1:D5、headerRowCount=1、`<autoFilter ref="A1:D5">` 一路覆盖到第 5 行）确实是
 * "表头 + 4 行数据、无汇总行"；对照 TotalsTable（ref=A8:C11 但 `<autoFilter ref="A8:C10">`
 * 恰好少一行，少的正是汇总行）。
 * 因此：`totalsRowCount` 只认显式属性，`totalsRowShown` 原样保留在 `totalsRowShown`
 * 字段里供上层判断，不参与计数（否则会把数据行当汇总行渲染）。
 */
function readTotalsRowCount(attrs: XmlAttributes, warn: Warn, label: string): number {
  const explicit = attrInt(attrs, 'totalsRowCount');
  if (explicit !== undefined) return explicit;
  if (attrBool(attrs, 'totalsRowShown') === true) {
    warn(`表格 ${label} 写了 totalsRowShown="1" 但没有 totalsRowCount，已按"无汇总行"处理`
      + `（ExcelJS 会把"无汇总行"写成 totalsRowShown="1"，字面语义相反）`);
  }
  return 0;
}

function parseColumns(xml: string, holder: ElementRange): Array<{ name: string }> {
  const out: Array<{ name: string }> = [];
  for (const kid of childElements(xml, holder)) {
    if (localName(kid.name) !== 'tableColumn') continue;
    const name = kid.attrs['name'];
    if (name !== undefined) out.push({ name });
  }
  return out;
}

function applyStyleInfo(table: ParsedTable, el: ElementRange): void {
  const name = el.attrs['name'];
  if (name !== undefined && name !== '') table.styleName = name;
  const stripes = attrBool(el.attrs, 'showRowStripes');
  if (stripes !== undefined) table.showRowStripes = stripes;
  const colStripes = attrBool(el.attrs, 'showColumnStripes');
  if (colStripes !== undefined) table.showColumnStripes = colStripes;
  const first = attrBool(el.attrs, 'showFirstColumn');
  if (first !== undefined) table.showFirstColumn = first;
  const last = attrBool(el.attrs, 'showLastColumn');
  if (last !== undefined) table.showLastColumn = last;
}
