/**
 * **OpenDocument Spreadsheet（`.ods`）** 解析器：产出中性的 `WorkbookInput`（见
 * `src/importer/synth-xlsx.ts`），再由 `writeWorkbookPackage` 打成规范 xlsx，于是预览/编辑/
 * 工作区/撤销/导出/会话恢复全部复用既有链路。ODS 只是一条"入口翻译"，与 `csv.ts` 同构。
 *
 * 格式规范里的六个坑（都按真 Excel 产出的样本处理）：
 *  1. **重复展开**：`number-columns-repeated` / `number-rows-repeated` 会把"右侧/下方补齐到整张表"
 *     写成上万列的巨型元素。ODS 用"空行 + 重复"表达行间距，所以必须**截断而不是跳过**（跳过会让
 *     后面的行坐标整体前移）；空重复最多推进 `MAX_EMPTY_ROW_RUN` / `MAX_EMPTY_COL_RUN` 步并记 warning。
 *  2. **合并**：覆盖区域只由左上角的"声明者"产生，遍历时把光标推过被覆盖的格子并记进 `covered`；
 *     绝不按 `table:covered-table-cell` 反推区域——它连盖了几格都不知道。
 *  3. **公式**：ODS 用 OpenFormula（`of:=SUM([.A1:.A3])`）。只翻译 `of:` / `oooc:` 前缀且能逐 token
 *     校验的引用与运算符；命名表达式、实现私有函数、`oooc:` 的正则与通配符扩展一律**丢公式、
 *     留缓存值**并计入 `stats.formulaDropped`——宁可少一个公式也不能写错。
 *  4. **日期与时间**：ODS 用 ISO 文本，必须换算成 Excel 序列号，否则单元格退化成文本、下游数字格式
 *     失去意义；`PT36H0M0S` 这类超过 24 小时的时长按 1.5 天算，不能取模成 0.5。
 *  5. **样式来源两处**：`content.xml` 的 `office:automatic-styles`（就近覆盖，优先）与 `styles.xml`
 *     的 `office:styles`（命名样式 + 全部 `number:*` 定义）；`style:parent-style-name` 链要继承。
 *  6. **列宽换算**：1 字符宽 = 0.2117cm（= 6pt = 8px@96dpi），见 `toCharWidth`。
 *
 * 明确不支持（如实丢弃 + 记 warning，绝不猜值）：条件格式（`style:map`）、数据验证、批注、超链接 URL
 * （文本本身保留）、图片与图形、命名区域、透视表、隐藏行的隐藏状态、`table:table-header-columns` 的冻结列语义。
 */
import { childElements, findFirstElement, localName, nameMatches } from './xml';
import { readText, readZipEntries, type ZipEntries } from './zip';
import type { ElementRange, XmlAttributes } from './xml';
import type { SynthCell, SynthSheet, SynthStyle, WorkbookInput } from '../importer/synth-xlsx';

/* 公开常量与结果类型                                                           */

/** 单次导入的单元格总量上限（与 `csv.ts` 的 `MAX_DELIMITED_CELLS` 同量级、同语义） */
export const MAX_ODS_CELLS = 200_000;
/** 单边上限（与 xlsx 规范一致） */
export const MAX_ODS_ROWS = 1_048_576;
export const MAX_ODS_COLS = 16_384;
/** 连续空列 / 空行的展开上限：16384 就是 ODS 在 Excel 里的自然补齐边界，真文件能完整展开，损坏文件最多多走 16384 步 */
export const MAX_EMPTY_COL_RUN = 16_384;
export const MAX_EMPTY_ROW_RUN = 16_384;
/** 单个 `number-*-repeated` 属性的硬上限（防属性里塞天文数字） */
const MAX_REPEAT = MAX_ODS_ROWS;
/** `text:s text:c` 的展开上限（空格个数） */
const MAX_SPACES = 1_024;
/** 单元格文本长度上限（超长文本对预览无意义，却会撑爆内存） */
const MAX_TEXT_LENGTH = 32_767;

/** 系数依据见文件头第 6 条：1 字符宽 = 0.2117cm */
const CM_PER_CHAR_WIDTH = 0.2117;
const CM_PER_INCH = 2.54;
const PT_PER_INCH = 72;
const MM_PER_INCH = 25.4;

export interface OdsStats {
  /** `content.xml` 里的 `table:table` 个数（含隐藏表） */
  sheetCount: number;
  /** 真正产出到中性模型的单元格数（不含被跳过的空格子） */
  cells: number;
  /** 成功翻译的公式数 */
  formulas: number;
  /** 因无法可靠翻译而"丢公式留值"的个数 */
  formulaDropped: number;
  merges: number;
  /** 被安全阀截断的次数（列重复 + 行重复 + 单元格总量）；`truncated` 表示是否发生过任何截断 */
  truncatedRepeats: number;
  truncated: boolean;
}

export interface OdsParseResult {
  input: WorkbookInput;
  /** 给人看的降级说明（不是异常；进"导入摘要"） */
  warnings: string[];
  stats: OdsStats;
}

/* 主入口                                                                      */

/** 解析 `.ods` 字节，返回中性工作簿 + 降级报告。不抛异常：不是 zip、缺 `content.xml`、结构怪异都退化成"空工作簿 + warning"（局部坏掉不该让整次导入失败） */
export function parseOdsResult(bytes: Uint8Array): OdsParseResult {
  const warnings: string[] = [];
  const stats: OdsStats = {
    sheetCount: 0,
    cells: 0,
    formulas: 0,
    formulaDropped: 0,
    merges: 0,
    truncatedRepeats: 0,
    truncated: false,
  };

  let entries: ZipEntries;
  try {
    entries = readZipEntries(bytes);
  } catch (err) {
    warnings.push(`无法解压 ODS（不是合法 zip）：${err instanceof Error ? err.message : String(err)}`);
    return { input: { sheets: [] }, warnings, stats };
  }

  const contentXml = readText(entries, 'content.xml');
  if (contentXml === undefined) {
    warnings.push('ODS 包内没有 content.xml，无法解析表格内容');
    return { input: { sheets: [] }, warnings, stats };
  }

  const styles = new StyleResolver(contentXml, readText(entries, 'styles.xml'));
  const sheets = parseSheets(contentXml, styles, warnings, stats);
  stats.sheetCount = sheets.length;
  return { input: { sheets, styles: styles.synthStyles() }, warnings, stats };
}

/** 主入口（固定签名）。需要降级报告或统计时用 `parseOdsResult`，两者的 `input` 完全一致 */
export function parseOds(bytes: Uint8Array): WorkbookInput {
  return parseOdsResult(bytes).input;
}

/* 表格遍历                                                                     */

/** 找到 `office:body/office:spreadsheet`；缺任一环返回 undefined，由调用方记 warning */
function findSpreadsheet(contentXml: string): ElementRange | undefined {
  const body = findFirstElement(contentXml, 'office:body');
  if (body) {
    for (const child of childElements(contentXml, body)) {
      if (nameMatches(child.name, 'office:spreadsheet')) return child;
    }
  }
  const root = findFirstElement(contentXml, 'office:document-content');
  if (!root) return undefined;
  for (const child of childElements(contentXml, root)) {
    if (nameMatches(child.name, 'office:spreadsheet')) return child;
  }
  return undefined;
}

/** 跨表共享的预算：一次导入的总量上限，避免"每张表各 20 万"把内存撑爆 */
interface Budget {
  cells: number;
  formulas: number;
  formulaDropped: number;
  exhausted: boolean;
}

interface RowParse {
  /** 单元格；`row` 已经是**绝对**行号（0-based） */
  cells: SynthCell[];
  /** 本行声明的合并（A1 记号，绝对坐标）；跨行合并只有首行声明 */
  merges: string[];
  /** 本行出现过的最大列号（0-based）；裁剪列宽用 */
  maxCol: number;
  /** 本行是否因为单元格预算耗尽而没读完 */
  truncated: boolean;
}

function parseSheets(
  contentXml: string,
  styles: StyleResolver,
  warnings: string[],
  stats: OdsStats,
): SynthSheet[] {
  const spreadsheet = findSpreadsheet(contentXml);
  if (!spreadsheet) {
    warnings.push('content.xml 里找不到 office:body/office:spreadsheet，没有可解析的表格');
    return [];
  }

  const sheets: SynthSheet[] = [];
  const budget: Budget = { cells: 0, formulas: 0, formulaDropped: 0, exhausted: false };
  for (const tableEl of childElements(contentXml, spreadsheet)) {
    if (!nameMatches(tableEl.name, 'table:table')) continue;
    sheets.push(parseTable(contentXml, tableEl, sheets.length, styles, warnings, stats, budget));
  }
  stats.formulas = budget.formulas;
  stats.formulaDropped = budget.formulaDropped;
  return sheets;
}

function parseTable(
  contentXml: string,
  tableEl: ElementRange,
  index: number,
  styles: StyleResolver,
  warnings: string[],
  stats: OdsStats,
  budget: Budget,
): SynthSheet {
  const rawName = tableEl.attrs['table:name'] ?? '';
  const name = rawName.trim() || `工作表${index + 1}`;
  if (!rawName.trim()) warnings.push(`第 ${index + 1} 张表没有 table:name，暂命名为"${name}"`);
  const label = `表「${name}」`;

  const tableProps = styles.resolveTableProps(tableEl.attrs['table:style-name']);
  const cells: SynthCell[] = [];
  const merges: string[] = [];
  /** 列宽先按 `table:table-column` 声明收集，**行都解析完**再裁到"有内容的最后一列"：ODS 把整表补齐
   * 也写成一条 `number-columns-repeated="16376"` 的列定义，全收会输出 16384 条列宽（下游 writer 变成 16384 个 `<col>`） */
  const pendingColWidths = new Map<number, number>();
  const rowHeights: Record<number, number> = {};
  /** 被合并覆盖的绝对坐标（`行 * MAX_ODS_COLS + 列`）；跨行合并会覆盖多行，所以是表级集合 */
  const covered = new Set<number>();

  let row = 0;
  let column = 0;
  let usedColumns = 0;
  let clampedEmptyRows = 0;
  let clampedEmptyCols = 0;
  let hiddenColumnCount = 0;
  let hiddenRowCount = 0;
  let headerRowCount = 0;
  let droppedDecorations = 0;

  for (const child of childElements(contentXml, tableEl)) {
    if (nameMatches(child.name, 'table:table-header-rows')) {
      // 语义只是"这些行是表头"，映射到冻结行；这些行本身仍在下面按普通行解析
      headerRowCount += countHeaderRows(contentXml, child);
      continue;
    }
    if (nameMatches(child.name, 'table:table-column')) {
      const columnInfo = applyColumn(child, styles, pendingColWidths, column);
      column = columnInfo.cursor;
      clampedEmptyCols += columnInfo.clamped;
      hiddenColumnCount += columnInfo.hidden;
      continue;
    }
    if (nameMatches(child.name, 'table:table-columns')) {
      for (const inner of childElements(contentXml, child)) {
        if (!nameMatches(inner.name, 'table:table-column')) continue;
        const columnInfo = applyColumn(inner, styles, pendingColWidths, column);
        column = columnInfo.cursor;
        clampedEmptyCols += columnInfo.clamped;
        hiddenColumnCount += columnInfo.hidden;
      }
      continue;
    }
    if (nameMatches(child.name, 'table:table-row')) {
      if (budget.exhausted) break;
      const repeat = repeatCount(child.attrs['table:number-rows-repeated']);
      const rowProps = styles.resolveRowProps(child.attrs['table:style-name']);
      const parsed = parseRow(contentXml, child, row, styles, warnings, label, budget, covered);
      droppedDecorations += countDecorations(contentXml, child);
      if (rowProps.hidden) hiddenRowCount += 1;

      const hasContent = parsed.cells.length > 0 || parsed.merges.length > 0;
      if (!hasContent && rowProps.hidden !== true) {
        // 纯空行：只推进光标（空行承载"行间距"语义，不能直接丢）。不看行高：空行上的行高是
        // Excel 补齐时写的默认值，照收会让 `rowHeights` 从"用户设过的几行"膨胀成几万行。
        const advance = Math.min(repeat, MAX_EMPTY_ROW_RUN);
        if (advance < repeat) {
          clampedEmptyRows += repeat - advance;
          stats.truncatedRepeats += 1;
          stats.truncated = true;
        }
        row += advance;
        continue;
      }

      let placed = 0;
      for (let instance = 0; instance < repeat; instance += 1) {
        const base = row + instance;
        if (base >= MAX_ODS_ROWS) {
          warnings.push(`${label} 行数超过上限（${MAX_ODS_ROWS}），其余行未导入`);
          stats.truncated = true;
          break;
        }
        if (instance > 0 && budget.cells + parsed.cells.length > MAX_ODS_CELLS) {
          warnings.push(`${label} 单元格数超过上限（${MAX_ODS_CELLS}），其余重复行未导入`);
          budget.exhausted = true;
          stats.truncated = true;
          break;
        }
        for (const cell of parsed.cells) {
          const target = cell.row + instance;
          cells.push(target === cell.row ? cell : { ...cell, row: target });
          budget.cells += 1;
        }
        for (const merge of parsed.merges) merges.push(shiftMergeRows(merge, instance));
        // 只有"有内容的行"的自定义行高才落进模型（空行的高是补齐噪声）
        if (rowProps.height !== undefined && hasContent) rowHeights[base] = rowProps.height;
        placed += 1;
      }
      if (parsed.maxCol > usedColumns) usedColumns = parsed.maxCol;
      row += placed;
      if (parsed.truncated) {
        budget.exhausted = true;
        break;
      }
      continue;
    }
  }

  const colWidths: Record<number, number> = {};
  for (const [col, width] of pendingColWidths) {
    if (col <= usedColumns) colWidths[col] = width;
  }

  if (clampedEmptyRows > 0) {
    warnings.push(`${label} 有连片的空重复行，截断 ${clampedEmptyRows} 行未展开（保留上限 ${MAX_EMPTY_ROW_RUN} 行）`);
  }
  if (clampedEmptyCols > 0) {
    warnings.push(`${label} 有连片的空重复列，截断 ${clampedEmptyCols} 列未展开（保留上限 ${MAX_EMPTY_COL_RUN} 列）`);
  }
  if (hiddenColumnCount > 0) {
    warnings.push(`${label} 有 ${hiddenColumnCount} 个隐藏列未导入（中性模型的列宽表没有隐藏标记）`);
  }
  if (hiddenRowCount > 0) {
    warnings.push(`${label} 有 ${hiddenRowCount} 个隐藏行未导入（中性模型只有行高，没有隐藏标记）`);
  }
  if (droppedDecorations > 0) {
    warnings.push(`${label} 有 ${droppedDecorations} 处批注 / 超链接 / 图片等非单元格内容，未导入`);
  }

  const sheet: SynthSheet = { name, cells };
  if (merges.length > 0) sheet.merges = merges;
  if (Object.keys(colWidths).length > 0) sheet.colWidths = colWidths;
  if (Object.keys(rowHeights).length > 0) sheet.rowHeights = rowHeights;
  if (tableProps.display === false) sheet.hidden = true;
  if (tableProps.gridlines === false) sheet.gridlinesHidden = true;
  if (headerRowCount > 0) sheet.freeze = { rows: Math.min(headerRowCount, MAX_ODS_ROWS), cols: 0 };

  stats.cells += cells.length;
  stats.merges += merges.length;
  return sheet;
}

/* 列                                                                          */

interface ColumnResult {
  cursor: number;
  /** 因为超过空列展开上限而被截掉的列数 */
  clamped: number;
  /** 被标记为隐藏的列数（中性模型没有隐藏列字段，只用于报告） */
  hidden: number;
}

/** 一个 `table:table-column`：展开 `number-columns-repeated`，把宽度写进 `colWidths`。
 * 样式里没有列宽时（例如整表补齐用的裸列）**留空**，让下游用默认列宽——把默认值当"用户设过的
 * 宽度"写出去，Excel 里每一列都会变成"自定义宽度"，用户再也点不回"自动"。 */
function applyColumn(
  colEl: ElementRange,
  styles: StyleResolver,
  colWidths: Map<number, number>,
  cursor: number,
): ColumnResult {
  const repeat = repeatCount(colEl.attrs['table:number-columns-repeated']);
  const visual = styles.resolveColumnVisual(colEl.attrs['table:style-name']);
  const inlineHidden = colEl.attrs['table:visibility'] === 'collapse' || colEl.attrs['table:visibility'] === 'filter';
  const hidden = visual.hidden === true || inlineHidden;

  if (visual.width === undefined && !hidden) {
    const advance = Math.min(repeat, MAX_EMPTY_COL_RUN);
    return { cursor: cursor + advance, clamped: repeat - advance, hidden: 0 };
  }

  const count = Math.min(repeat, Math.max(0, MAX_ODS_COLS - cursor));
  if (visual.width !== undefined) {
    for (let i = 0; i < count; i += 1) colWidths.set(cursor + i, visual.width);
  }
  return { cursor: cursor + count, clamped: repeat - count, hidden: hidden ? count : 0 };
}

/* 行与单元格                                                                   */

function parseRow(
  contentXml: string,
  rowEl: ElementRange,
  rowIndex: number,
  styles: StyleResolver,
  warnings: string[],
  label: string,
  budget: Budget,
  covered: Set<number>,
): RowParse {
  const cells: SynthCell[] = [];
  const merges: string[] = [];
  let column = 0;
  let maxCol = -1;
  let truncated = false;

  for (const cellEl of childElements(contentXml, rowEl)) {
    const isCell = nameMatches(cellEl.name, 'table:table-cell');
    const isCovered = nameMatches(cellEl.name, 'table:covered-table-cell');
    if (!isCell && !isCovered) continue;

    const repeat = repeatCount(cellEl.attrs['table:number-columns-repeated']);
    const span = isCell ? cellSpans(cellEl.attrs) : { colSpan: 1, rowSpan: 1 };

    if (isCovered) {
      // 只有**带位置语义样式**的 covered 格子才占位（代表用户设过底色或对齐）；裸的
      // `<table:covered-table-cell number-columns-repeated="16383"/>` 与纯默认字体的占位符一律跳过。
      const styleRef = styles.internCell(cellEl.attrs['table:style-name']);
      const count = Math.min(repeat, Math.max(0, MAX_ODS_COLS - column));
      if (styleRef.visual) {
        for (let i = 0; i < count; i += 1) {
          if (covered.has(cellKey(rowIndex, column + i))) continue;
          if (budget.cells >= MAX_ODS_CELLS) {
            truncated = true;
            break;
          }
          cells.push({ row: rowIndex, col: column + i, value: null, style: styleRef.index });
          budget.cells += 1;
        }
      }
      column += count;
      if (truncated) break;
      continue;
    }

    // 合并声明用"未截断"的坐标判断，否则会被 16384 列上限扭曲
    if (span.colSpan > 1 || span.rowSpan > 1) {
      const startCol = column;
      const endCol = Math.min(MAX_ODS_COLS - 1, column + span.colSpan - 1);
      const endRow = Math.min(MAX_ODS_ROWS - 1, rowIndex + span.rowSpan - 1);
      merges.push(mergeRef(rowIndex, startCol, endRow, endCol));
      for (let r = rowIndex; r <= endRow; r += 1) {
        for (let c = startCol + 1; c <= endCol; c += 1) covered.add(cellKey(r, c));
      }
    }

    if (column < MAX_ODS_COLS && !budget.exhausted) {
      const styleRef = styles.internCell(cellEl.attrs['table:style-name']);
      const built = isEmptyPlaceholder(cellEl, styleRef)
        ? undefined
        : readCell(contentXml, cellEl, rowIndex, column, styleRef.index, budget, warnings, label);
      // `number-columns-repeated` 是"这个单元格定义重复 N 次"，所以有内容时要复制 N 份
      const copies = span.colSpan > 1 ? 1 : Math.min(repeat, Math.max(0, MAX_ODS_COLS - column));
      if (built) {
        for (let i = 0; i < copies; i += 1) {
          if (budget.cells >= MAX_ODS_CELLS) {
            truncated = true;
            break;
          }
          cells.push(i === 0 ? built : { ...built, col: built.col + i });
          budget.cells += 1;
        }
        const rightmost = column + copies - 1 + (span.colSpan > 1 ? span.colSpan - 1 : 0);
        if (rightmost > maxCol) maxCol = rightmost;
      }
      column += copies;
      if (truncated) break;
      continue;
    }

    // 声明者自身占 1 格；没有内容时只推进光标
    column += span.colSpan > 1 ? 1 : Math.min(repeat, Math.max(1, MAX_ODS_COLS - column));
  }

  return { cells, merges, maxCol, truncated };
}

/** 这个 `table:table-cell` 是不是"纯占位空格子"：ODS（Excel 导出尤其明显）会给每行尾部补齐
 * `number-columns-repeated="16383"` 的裸格子，全收下来真夹具里 30 个真值格子会膨胀成 59 个假格子。
 * 判据是"没有样式引用"或"样式只有字体、无任何位置语义（填充/对齐/边框/数字格式）"——后者是关键，
 * 因为每个格子都带继承自 `Default` 的 `ce1`（宋体 11pt），光看有无样式判不出来；带填充/对齐的空格子
 * 是用户真设过的，必须保留。 */
/** 这些属性的出现说明"这个格子有值"，它就不是占位符 */
const VALUE_ATTRS = [
  'office:value', 'office:boolean-value', 'office:date-value', 'office:time-value',
  'office:string-value', 'table:formula',
] as const;

function isEmptyPlaceholder(
  cellEl: ElementRange,
  styleRef: { index: number; visual: boolean },
): boolean {
  for (const key of VALUE_ATTRS) {
    if (cellEl.attrs[key] !== undefined) return false;
  }
  if (!cellEl.selfClosing) return false;
  if (cellEl.attrs['table:style-name'] === undefined) return true;
  return !styleRef.visual;
}

/** 读一个 `table:table-cell`：值、公式、样式。值优先信 `office:value*` 属性（规范里的权威值），
 * 属性缺失时退回 `<text:p>` 文本——手工写的 ODS 常常只写文本、不写 `office:value-type`。 */
function readCell(
  contentXml: string,
  cellEl: ElementRange,
  row: number,
  col: number,
  styleIndex: number,
  budget: Budget,
  warnings: string[],
  label: string,
): SynthCell | undefined {
  const valueType = cellEl.attrs['office:value-type'] ?? '';
  let value: string | number | boolean | undefined;

  switch (valueType) {
    case 'float':
    case 'currency':
    case 'percentage':
      value = numberFrom(cellEl.attrs['office:value']);
      break;
    case 'boolean':
      value = booleanFrom(cellEl.attrs['office:boolean-value']);
      break;
    case 'date':
      value = serialFromDateText(cellEl.attrs['office:date-value'] ?? '');
      break;
    case 'time':
      value = daysFromDuration(cellEl.attrs['office:time-value'] ?? '');
      break;
    case 'string':
      value = cellEl.attrs['office:string-value'] ?? readCellText(contentXml, cellEl);
      break;
    case '':
    case 'void': {
      const text = readCellText(contentXml, cellEl);
      if (text !== '') value = text;
      break;
    }
    default:
      // 未识别的 value-type：不猜，留空（公式仍可能带上缓存值）
      warnings.push(`${label} 出现未知的 office:value-type="${valueType}"，该单元格只保留公式`);
      break;
  }

  let formula: string | undefined;
  const formulaAttr = cellEl.attrs['table:formula'];
  if (formulaAttr !== undefined && formulaAttr !== '') {
    const translated = odsFormulaToA1(formulaAttr);
    if (translated !== undefined) {
      formula = translated;
      budget.formulas += 1;
    } else {
      budget.formulaDropped += 1;
    }
  }

  // 没有值、没有公式、连样式索引都是 -1（= 该样式解析后无任何可表达字段）→ 不产出单元格。
  if (value === undefined && formula === undefined && styleIndex < 0) return undefined;
  const cell: SynthCell = { row, col };
  // 没有值但带位置语义样式的格子：显式写 `value: null`，与"空单元格"的惯例（见 `SynthCell`）一致
  cell.value = value ?? null;
  if (formula !== undefined) cell.formula = formula;
  if (styleIndex >= 0) cell.style = styleIndex;
  return cell;
}

/* 单元格文本                                                                   */

/** 取单元格的可见文本：拼接所有 `<text:p>`（段间换行），展开 `<text:s text:c="3"/>`、`<text:tab/>`、
 * `<text:line-break/>`，保留 `<text:span>` / `<text:a>` 里的文字。不能复用 `elementText`：ODS 的空白
 * 语义全在这些自闭合元素上，它会把 `<text:s/>` 剥掉、让 "粗蓝 + 斜体" 粘成一个词；这里按标签逐段取，
 * 并跳过元素之间的排版空白（缩进换行）。 */
export function readCellText(contentXml: string, cellEl: ElementRange): string {
  if (cellEl.selfClosing) return '';
  const start = cellEl.openEnd;
  const end = Math.max(cellEl.openEnd, cellEl.closeEnd);
  /** 段文本按"最外层 `<text:p>`"归段：`<text:p>` 可以嵌套（列表项），靠"上一个标签是不是
   * `</text:p>`"判段界会在嵌套时多补或漏补换行；按深度归段（0→1 开一段，回到 0 收一段）是纯局部的。 */
  const paragraphs: string[] = [];
  let current = '';
  /** `<text:p>` 的嵌套深度（> 0 表示正在收集属于单元格的段落文本） */
  let paragraphDepth = 0;
  /** 遮挡容器（`<office:annotation>`、`<draw:frame>`、`<svg:desc>` …）的嵌套深度：它们内部的
   * `<text:p>` 是批注 / 图形说明，拼进来会造出"批注 + 正文"粘在一起的假值。 */
  let blockDepth = 0;
  let i = start;
  let textStart = start;

  const collecting = (): boolean => paragraphDepth > 0 && blockDepth === 0;

  const flush = (upTo: number): void => {
    if (!collecting() || upTo <= textStart) return;
    const raw = contentXml.slice(textStart, upTo);
    if (raw !== '') current += decodeXmlEntities(raw);
  };

  while (i < end) {
    const lt = contentXml.indexOf('<', i);
    if (lt < 0 || lt >= end) break;
    flush(Math.min(lt, end));
    if (contentXml.startsWith('<!--', lt)) {
      const stop = contentXml.indexOf('-->', lt + 4);
      if (stop < 0) break;
      i = stop + 3;
      textStart = i;
      continue;
    }
    if (contentXml.startsWith('<![CDATA[', lt)) {
      const stop = contentXml.indexOf(']]>', lt + 9);
      if (stop < 0) break;
      if (collecting()) current += contentXml.slice(lt + 9, Math.min(stop, end));
      i = stop + 3;
      textStart = i;
      continue;
    }
    const gt = contentXml.indexOf('>', lt);
    if (gt < 0 || gt >= end) break;
    const closing = contentXml.charCodeAt(lt + 1) === 47; // '/'
    const nameEnd = scanTagNameEnd(contentXml, lt, gt);
    const name = localName(contentXml.slice(closing ? lt + 2 : lt + 1, nameEnd));
    const attrs = parseTagAttrs(contentXml, nameEnd, gt);
    const selfClosing = isSelfClosingTag(contentXml, lt, gt);

    if (name === 'p' || name === 'h') {
      if (closing) {
        if (paragraphDepth > 0) paragraphDepth -= 1;
        if (paragraphDepth === 0) {
          paragraphs.push(current);
          current = '';
        }
      } else if (!selfClosing) {
        if (paragraphDepth === 0 && blockDepth === 0) {
          paragraphs.push(current);
          current = '';
        }
        paragraphDepth += 1;
      }
    } else if (BLOCK_CONTAINERS.has(name) && !selfClosing) {
      if (closing) blockDepth = Math.max(0, blockDepth - 1);
      else blockDepth += 1;
    } else if (collecting()) {
      if (name === 's') current += ' '.repeat(Math.min(numberFrom(attrs['text:c']) ?? 1, MAX_SPACES));
      else if (name === 'tab') current += '\t';
      else if (name === 'line-break') current += '\n';
    }
    i = gt + 1;
    textStart = i;
  }
  flush(end);
  paragraphs.push(current);

  // 只裁每段首尾的普通空白（XML 缩进）：用户要的空格由 `<text:s/>` 表达、展开后也落在段首，对整串 trim 会吃掉它
  const trimmed = paragraphs
    .map((part) => part.replace(/^[ \t]+/, '').replace(/[ \t\n]+$/, ''))
    .filter((part) => part !== '');
  const joined = trimmed.join('\n');
  return joined.length > MAX_TEXT_LENGTH ? joined.slice(0, MAX_TEXT_LENGTH) : joined;
}

/** 不算单元格值的块级容器：内部的 `<text:p>` 是批注 / 图形说明 / 索引，拼进来会造出"批注 + 正文"粘在一起的假值 */
const BLOCK_CONTAINERS: ReadonlySet<string> = new Set([
  'annotation', 'frame', 'desc', 'title', 'note', 'index-body', 'ruby',
]);

function isSelfClosingTag(xml: string, lt: number, gt: number): boolean {
  let back = gt - 1;
  while (back > lt && isXmlSpace(xml.charCodeAt(back))) back -= 1;
  return xml.charCodeAt(back) === 47;
}

function scanTagNameEnd(xml: string, lt: number, gt: number): number {
  let i = xml.charCodeAt(lt + 1) === 47 ? lt + 2 : lt + 1;
  while (i < gt && !isXmlSpace(xml.charCodeAt(i)) && xml.charCodeAt(i) !== 47) i += 1;
  return i;
}

function decodeXmlEntities(raw: string): string {
  if (raw.indexOf('&') < 0) return raw;
  return raw.replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|([a-zA-Z]+));/g, (all, dec: string, hex: string, named: string) => {
    if (dec !== undefined) return codePoint(Number.parseInt(dec, 10));
    if (hex !== undefined) return codePoint(Number.parseInt(hex, 16));
    switch (named) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      case 'nbsp': return '\u00a0';
      default: return all;
    }
  });
}

function codePoint(value: number): string {
  if (!Number.isFinite(value) || value < 0 || value > 0x10ffff) return '';
  return String.fromCodePoint(value);
}

function isXmlSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13;
}

/** 在标签自身的 `[start, end)` 区间里解析属性（`xml.ts` 的 `parseAttrs` 坐标是整篇文档） */
function parseTagAttrs(xml: string, start: number, end: number): XmlAttributes {
  const attrs: XmlAttributes = {};
  let i = start;
  while (i < end) {
    while (i < end && (isXmlSpace(xml.charCodeAt(i)) || xml.charCodeAt(i) === 47)) i += 1;
    const nameStart = i;
    while (i < end && !isXmlSpace(xml.charCodeAt(i)) && xml.charCodeAt(i) !== 61 && xml.charCodeAt(i) !== 62) i += 1;
    if (i === nameStart) break;
    const name = xml.slice(nameStart, i);
    while (i < end && isXmlSpace(xml.charCodeAt(i))) i += 1;
    if (xml.charCodeAt(i) !== 61) {
      attrs[name] = 'true';
      continue;
    }
    i += 1;
    while (i < end && isXmlSpace(xml.charCodeAt(i))) i += 1;
    const quote = xml.charCodeAt(i);
    if (quote === 34 || quote === 39) {
      i += 1;
      const valueStart = i;
      while (i < end && xml.charCodeAt(i) !== quote) i += 1;
      attrs[name] = decodeXmlEntities(xml.slice(valueStart, i));
      i += 1;
    } else {
      const valueStart = i;
      while (i < end && !isXmlSpace(xml.charCodeAt(i)) && xml.charCodeAt(i) !== 62) i += 1;
      attrs[name] = decodeXmlEntities(xml.slice(valueStart, i));
    }
  }
  return attrs;
}

/* 值换算                                                                       */

function numberFrom(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim();
  if (text === '') return undefined;
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

/** ODS 的布尔是字面 `true` / `false`（不是 OOXML 的 1 / 0） */
function booleanFrom(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim().toLowerCase();
  if (text === 'true' || text === '1') return true;
  if (text === 'false' || text === '0') return false;
  return undefined;
}

/** ISO 日期（`2025-01-01` / `2025-01-01T12:00:00`，可带小数秒）→ Excel 序列号。基准取
 * 1899-12-30（Excel 1900 闰年 bug 之后一天），1900-03-01 起与 Excel 完全一致；时间部分按 UTC
 * 解析——ODS 的 `date-value` 是不带偏移的本地时间，UTC 解析等于原样取用。 */
export function serialFromDateText(text: string): number | undefined {
  const m = /^(-?\d{4,})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?)?/.exec(text.trim());
  if (!m) return undefined;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return undefined;
  const ms = Date.UTC(year, month - 1, day);
  if (!Number.isFinite(ms)) return undefined;
  const days = ms / 86_400_000 - Date.UTC(1899, 11, 30) / 86_400_000;
  const hours = m[4] === undefined ? 0 : Number(m[4]);
  const minutes = m[5] === undefined ? 0 : Number(m[5]);
  const seconds = m[6] === undefined ? 0 : Number(m[6]);
  const fraction = (hours * 3600 + minutes * 60 + seconds) / 86_400;
  return Math.round((days + fraction) * 1e6) / 1e6;
}

/** ISO 8601 时长（`PT18H0M0S`）→ 天（Excel 的时间值）。`PT36H0M0S` 是 1.5 天而不是 0.5 天：
 * ODS 的时长允许超过 24 小时，取模会把"36 小时"变成"12 小时"。 */
export function daysFromDuration(text: string): number | undefined {
  const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(text.trim());
  if (!m) return undefined;
  const sign = m[1] === '-' ? -1 : 1;
  const days = m[2] === undefined ? 0 : Number(m[2]);
  const hours = m[3] === undefined ? 0 : Number(m[3]);
  const minutes = m[4] === undefined ? 0 : Number(m[4]);
  const seconds = m[5] === undefined ? 0 : Number(m[5]);
  const total = days + (hours * 3600 + minutes * 60 + seconds) / 86_400;
  return Math.round(sign * total * 1e6) / 1e6;
}

/* 公式：OpenFormula → Excel A1 公式                                            */

/** 明确拒绝的函数名：OpenFormula 独有、实现私有、或与 Excel 同名函数的语义不同 */
const FORBIDDEN_FUNCTIONS: ReadonlySet<string> = new Set([
  'REGEX', 'DDE', 'BAHTTEXT', 'MULTIPLE.OPERATIONS', 'COM.MICROSOFT.EXCEL',
]);

/** 把 `table:formula` 的 OpenFormula 表达式翻成 Excel 的 A1 公式（**不含前导 `=`**）；无法可靠
 * 翻译时返回 `undefined`（调用方丢掉公式、保留缓存值）。逐 token 翻译，不做正则整体替换：
 * 引用 `[.A1]` / `[.A1:.A3]` / `[第二张.A2]` / `['我的 表'.A1]` → `A1` / `A1:A3` / `第二张!A2`；
 * 参数分隔 `;` → `,`；`&` 与 `+ - * / ^ = <> <= >= < > ( )` 两边语义一致、原样输出；引号串原样保留。
 * 出现 `!`（OpenFormula 的"不等于"）、`|`（联合）、`~`（通配符）、`{}`（数组）或命名表达式一律拒绝。 */
export function odsFormulaToA1(formula: string): string | undefined {
  let text = formula.trim();
  const colon = text.indexOf(':');
  if (colon > 0) {
    const prefix = text.slice(0, colon).toLowerCase();
    if (prefix !== 'of' && prefix !== 'oooc') return undefined;
    text = text.slice(colon + 1);
  }
  if (text.startsWith('=')) text = text.slice(1).trim();
  if (text === '') return '';

  const out: string[] = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '[') {
      const close = text.indexOf(']', i + 1);
      if (close < 0) return undefined;
      const ref = bracketRefToA1(text.slice(i + 1, close));
      if (ref === undefined) return undefined;
      out.push(ref);
      i = close + 1;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let literal = '"';
      for (;;) {
        if (j >= text.length) return undefined;
        if (text[j] === '"') {
          if (text[j + 1] === '"') {
            literal += '""';
            j += 2;
            continue;
          }
          literal += '"';
          j += 1;
          break;
        }
        literal += text[j];
        j += 1;
      }
      out.push(literal);
      i = j;
      continue;
    }
    if (ch === ';') {
      out.push(',');
      i += 1;
      continue;
    }
    if (ch === '|' || ch === '!' || ch === '~' || ch === '{' || ch === '}') return undefined;
    if (/[A-Za-z_$]/.test(ch)) {
      const m = /^[A-Za-z_$][A-Za-z0-9_$]*/.exec(text.slice(i));
      if (!m) return undefined;
      const word = m[0];
      if (FORBIDDEN_FUNCTIONS.has(word.toUpperCase())) return undefined;
      // `A1` 形状的裸标识符会被 Excel 当引用（命名表达式），语义不明 → 拒绝
      if (/^[A-Za-z]{1,3}\d{1,7}$/.test(word)) return undefined;
      out.push(word);
      i += word.length;
      continue;
    }
    if (/[0-9]/.test(ch)) {
      const m = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(text.slice(i));
      if (!m) return undefined;
      out.push(m[0]);
      i += m[0].length;
      continue;
    }
    if ('+-*/^&=<>(),%'.includes(ch)) {
      out.push(ch);
      i += 1;
      continue;
    }
    if (isXmlSpace(text.charCodeAt(i))) {
      i += 1;
      continue;
    }
    return undefined;
  }
  return out.join('');
}

interface SingleRef {
  /** Excel 写法（跨表已带 `表名!` 前缀） */
  text: string;
  /** 表名 + `!`（必要时带单引号），以及表名原文（用于校验范围两端一致） */
  sheetPrefix?: string;
  sheet?: string;
}

/** 方括号里的引用 → A1 记号（跨表用 Excel 的 `!`）。支持 `.A1`、`.A1:.A3`、`.$A$1:.$B$2`、
 * `第二张.A1`、`'我的 表'.A1`、`.A:A`、`.1:3`；命名表达式（`[销售额]`）返回 undefined。 */
function bracketRefToA1(raw: string): string | undefined {
  const body = raw.trim();
  if (body === '') return undefined;
  const colon = body.indexOf(':');
  const head = colon < 0 ? body : body.slice(0, colon);
  const start = singleRef(head);
  if (start === undefined) return undefined;
  if (colon < 0) return start.text;

  const end = singleRef(body.slice(colon + 1));
  if (end === undefined) return undefined;
  // `[.A1:第二张.A3]` 这种"一半带表名"的写法在 ODS 里不合法，宁可不翻
  if ((start.sheet === undefined) !== (end.sheet === undefined)) return undefined;
  if (start.sheet !== undefined && end.sheet !== undefined && start.sheet !== end.sheet) return undefined;
  const prefix = start.sheetPrefix ?? end.sheetPrefix ?? '';
  return `${prefix}${start.text}:${end.text}`;
}

/** `[.A1]` 的内部形态 → A1（`$` 原样保留；整列 `A:A`、整行 `1:3` 也认） */
function singleRef(raw: string): SingleRef | undefined {
  let body = raw.trim();
  let sheet: string | undefined;
  if (body.startsWith('$')) body = body.slice(1);
  if (body.startsWith("'")) {
    const close = body.indexOf("'", 1);
    if (close < 0) return undefined;
    sheet = body.slice(1, close).replace(/''/g, "'");
    if (body[close + 1] !== '.') return undefined;
    body = body.slice(close + 2);
  } else {
    const dot = body.indexOf('.');
    if (dot >= 0) {
      sheet = body.slice(0, dot);
      body = body.slice(dot + 1);
    }
  }
  const cell = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d{1,7})?$/.exec(body);
  if (!cell) return undefined;
  const col = cell[2];
  const row = cell[4];
  if (col === undefined && row === undefined) return undefined;
  const text = `${cell[1]}${col === undefined ? '' : col.toUpperCase()}${cell[3]}${row ?? ''}`;
  if (sheet === undefined || sheet === '') return { text };
  // 表名不需要引号时不加（与 Excel 一致：中文表名可以直接写）
  const quoted = /[\s'[\]\\/*?:!]/.test(sheet) || /^\d/.test(sheet);
  const prefix = quoted ? `'${sheet.replace(/'/g, "''")}'` : sheet;
  return { text: `${prefix}!${text}`, sheetPrefix: `${prefix}!`, sheet };
}

/* 样式                                                                        */

interface RawStyle {
  styleAttrs: XmlAttributes;
  text?: XmlAttributes;
  paragraph?: XmlAttributes;
  cellProps?: XmlAttributes;
  rowProps?: XmlAttributes;
  colProps?: XmlAttributes;
  tableProps?: XmlAttributes;
}

/** ODS 的子属性元素名 → `RawStyle` 字段 */
const CHILD_PROPS: Readonly<Record<string, keyof RawStyle>> = {
  'text-properties': 'text',
  'paragraph-properties': 'paragraph',
  'table-cell-properties': 'cellProps',
  'table-row-properties': 'rowProps',
  'table-column-properties': 'colProps',
  'table-properties': 'tableProps',
};

/** ODS 边框线型 → 中性线型（`synth-xlsx.ts` 的 `BORDER_STYLE_MAP` 认得的名字） */
const BORDER_STYLE_MAP: Readonly<Record<string, string>> = {
  solid: 'thin',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  // Excel 没有 groove / ridge / inset / outset，退化到最接近的细实线（宁可细，不可粗）
  groove: 'thin',
  ridge: 'thin',
  inset: 'thin',
  outset: 'thin',
};

/** 样式解析：把 ODS 的"样式名 → 属性链"变成中性的 `SynthStyle`。三条关键约定：
 *  - **就近覆盖**：同名样式以 `content.xml` 的 `office:automatic-styles` 为准（它描述"这份文档实际
 *    用的样式"，优先级更高），同一份文档内后出现的覆盖先出现的。
 *  - **继承**：`style:parent-style-name` 链从根到叶逐层覆盖；不继承会丢掉整张表的默认字体。
 *  - **去重**：解析结果按 JSON 串 intern，`SynthCell.style` 只存下标，上万同款单元格共用一条。 */
class StyleResolver {
  private readonly raw = new Map<string, RawStyle>();
  /** 样式名的来源：content 优先于 styles */
  private readonly fromContent = new Set<string>();
  private readonly numberStyles = new Map<string, { xml: string; el: ElementRange }>();
  private readonly list: SynthStyle[] = [];
  private readonly interned = new Map<string, number>();
  private readonly walked = new Map<string, RawStyle | null>();
  private readonly contentXml: string;
  private readonly stylesXml: string;

  constructor(contentXml: string, stylesXml: string | undefined) {
    this.contentXml = contentXml;
    this.stylesXml = stylesXml ?? '';
    this.collect(contentXml, 'content');
    if (stylesXml !== undefined && stylesXml !== this.contentXml) this.collect(stylesXml, 'styles');
  }

  private collect(xml: string, source: 'content' | 'styles'): void {
    const root = findFirstElement(xml, source === 'content' ? 'office:document-content' : 'office:document-styles')
      ?? findFirstElement(xml, 'office:document-content')
      ?? findFirstElement(xml, 'office:document-styles');
    if (!root) return;
    for (const child of childElements(xml, root)) {
      if (nameMatches(child.name, 'office:automatic-styles') || nameMatches(child.name, 'office:styles')) {
        this.collectContainer(xml, child, source);
      } else if (nameMatches(child.name, 'office:body')) {
        for (const inner of childElements(xml, child)) {
          if (nameMatches(inner.name, 'office:automatic-styles')) this.collectContainer(xml, inner, source);
        }
      }
    }
  }

  private collectContainer(xml: string, container: ElementRange, source: 'content' | 'styles'): void {
    for (const el of childElements(xml, container)) {
      const name = localName(el.name);
      const styleName = el.attrs['style:name'];
      if (name === 'style') {
        if (styleName === undefined || styleName === '') continue;
        // 就近覆盖：content.xml 的定义一旦写入，styles.xml 的同名定义不再覆盖
        if (source === 'styles' && this.fromContent.has(styleName)) continue;
        this.raw.set(styleName, this.readRawStyle(xml, el));
        if (source === 'content') this.fromContent.add(styleName);
        this.walked.delete(styleName);
      } else if (name.endsWith('-style')) {
        // number-style / percentage-style / date-style / time-style / text-style / boolean-style /
        // currency-style：数字格式定义，两种 xml 里都可能出现
        if (styleName !== undefined && styleName !== '') {
          const existing = this.numberStyles.get(styleName);
          if (!existing || source === 'content') this.numberStyles.set(styleName, { xml, el });
        }
      }
    }
  }

  private readRawStyle(xml: string, el: ElementRange): RawStyle {
    const style: RawStyle = { styleAttrs: el.attrs };
    for (const child of childElements(xml, el)) {
      const key = CHILD_PROPS[localName(child.name)];
      if (key !== undefined) style[key] = child.attrs;
    }
    return style;
  }

  /** 沿 `style:parent-style-name` 链从根到叶合并（子样式覆盖父样式） */
  private walk(name: string | undefined): RawStyle | null {
    if (name === undefined || name === '') return null;
    const cached = this.walked.get(name);
    if (cached !== undefined) return cached;

    const chain: RawStyle[] = [];
    const seen = new Set<string>();
    let current: string | undefined = name;
    while (current !== undefined && current !== '' && !seen.has(current)) {
      seen.add(current);
      const found = this.raw.get(current);
      if (!found) break;
      chain.push(found);
      current = found.styleAttrs['style:parent-style-name'];
    }
    if (chain.length === 0) {
      this.walked.set(name, null);
      return null;
    }
    const merged: RawStyle = { styleAttrs: {} };
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      const layer = chain[i];
      if (layer.text) merged.text = { ...merged.text, ...layer.text };
      if (layer.paragraph) merged.paragraph = { ...merged.paragraph, ...layer.paragraph };
      if (layer.cellProps) merged.cellProps = { ...merged.cellProps, ...layer.cellProps };
      if (layer.rowProps) merged.rowProps = { ...merged.rowProps, ...layer.rowProps };
      if (layer.colProps) merged.colProps = { ...merged.colProps, ...layer.colProps };
      if (layer.tableProps) merged.tableProps = { ...merged.tableProps, ...layer.tableProps };
      merged.styleAttrs = { ...merged.styleAttrs, ...layer.styleAttrs };
    }
    this.walked.set(name, merged);
    return merged;
  }

  resolveTableProps(name: string | undefined): { display?: boolean; gridlines?: boolean } {
    const props = this.walk(name)?.tableProps ?? {};
    const out: { display?: boolean; gridlines?: boolean } = {};
    if (props['table:display'] !== undefined) out.display = props['table:display'] !== 'false';
    if (props['table:gridlines'] !== undefined) out.gridlines = props['table:gridlines'] !== 'false';
    return out;
  }

  /** 行属性：**只有显式自定义行高才保留**（判据是 `style:use-optimal-row-height="false"`）。
   * ODS 里每行都会带一条行高（哪怕是自动算出的 13.5pt），全收写回 xlsx 会让每行都成
   * `customHeight="1"`，用户再也点不回"自动调整行高"。 */
  resolveRowProps(name: string | undefined): { height?: number; hidden?: boolean } {
    const props = this.walk(name)?.rowProps ?? {};
    const out: { height?: number; hidden?: boolean } = {};
    if (props['style:use-optimal-row-height'] === 'false') {
      const height = toPoints(props['style:row-height']);
      if (height !== undefined) out.height = height;
    }
    if (props['table:visibility'] === 'collapse' || props['table:visibility'] === 'filter') out.hidden = true;
    return out;
  }

  resolveColumnVisual(name: string | undefined): { width?: number; hidden?: boolean } {
    const props = this.walk(name)?.colProps ?? {};
    const out: { width?: number; hidden?: boolean } = {};
    const width = toCharWidth(props['style:column-width']);
    if (width !== undefined) out.width = width;
    if (props['table:visibility'] === 'collapse' || props['table:visibility'] === 'filter') out.hidden = true;
    return out;
  }

  /** 把一个 `table:style-name` 解析成 `SynthStyle` 下标；无任何可表达样式时返回 `-1`，让调用方
   * **不写** `SynthCell.style`（空格子不会被造出来、默认样式不会被写出去）。`visual` 表示是否带
   * "位置语义"（填充/对齐/边框/旋转/数字格式）：`parseRow` 靠它区分"真格子"与"跳格占位符"。 */
  internCell(name: string | undefined): { index: number; visual: boolean } {
    const raw = this.walk(name);
    if (!raw) return { index: -1, visual: false };
    const style = this.toSynthStyle(raw);
    if (!style) return { index: -1, visual: false };
    const visual = style.fill !== undefined
      || style.border !== undefined
      || style.horizontalAlign !== undefined
      || style.verticalAlign !== undefined
      || style.textWrap !== undefined
      || style.numberFormat !== undefined
      || (style.textRotation !== undefined && style.textRotation !== 0);
    const key = JSON.stringify(style);
    const existing = this.interned.get(key);
    if (existing !== undefined) return { index: existing, visual };
    const id = this.list.length;
    this.list.push(style);
    this.interned.set(key, id);
    return { index: id, visual };
  }

  private toSynthStyle(raw: RawStyle): SynthStyle | null {
    const text = raw.text ?? {};
    const cell = raw.cellProps ?? {};
    const paragraph = raw.paragraph ?? {};
    const style: SynthStyle = {};

    const family = firstDefined(text['fo:font-family'], text['style:font-name']);
    if (family !== undefined) {
      const clean = family.replace(/^['"]|['"]$/g, '').trim();
      if (clean !== '') style.fontFamily = clean;
    }
    const size = toPoints(text['fo:font-size'] ?? text['style:font-size-asian']);
    if (size !== undefined) style.fontSize = size;
    if (text['fo:font-weight'] !== undefined) {
      style.bold = text['fo:font-weight'] === 'bold' || text['fo:font-weight'] === '700';
    }
    if (text['fo:font-style'] !== undefined) {
      style.italic = text['fo:font-style'] === 'italic' || text['fo:font-style'] === 'oblique';
    }
    const underline = text['style:text-underline-style'];
    if (underline !== undefined) style.underline = underline !== 'none';
    const strike = text['style:text-line-through-style'];
    if (strike !== undefined) style.strikeThrough = strike !== 'none';
    const color = normalizeOdsColor(text['fo:color']);
    if (color !== undefined) style.color = color;
    const fill = normalizeOdsColor(cell['fo:background-color']);
    if (fill !== undefined) style.fill = fill;

    // 水平对齐在 ODS 里属于**段落**属性（`style:paragraph-properties`），
    // 不在 `table-cell-properties` 上——只查后者会丢掉真夹具里的整套对齐
    const align = paragraph['fo:text-align'] ?? cell['fo:text-align'];
    if (align !== undefined) {
      const mapped = mapHorizontalAlign(align);
      if (mapped !== undefined) style.horizontalAlign = mapped;
    }
    const vertical = cell['style:vertical-align'];
    if (vertical !== undefined) {
      const mapped = mapVerticalAlign(vertical);
      if (mapped !== undefined) style.verticalAlign = mapped;
    }
    if (cell['fo:wrap-option'] !== undefined) style.textWrap = cell['fo:wrap-option'] === 'wrap';
    const rotation = numberFrom(cell['style:rotation-angle'] ?? cell['fo:rotation-angle']);
    if (rotation !== undefined && rotation !== 0) {
      const normalized = ((rotation % 360) + 360) % 360;
      const excelRotation = normalized > 180 ? normalized - 360 : normalized;
      if (excelRotation !== 0) style.textRotation = Math.round(excelRotation * 10) / 10;
    }

    const border = readBorder(cell);
    if (border) style.border = border;

    const dataStyle = raw.styleAttrs['style:data-style-name'];
    if (dataStyle !== undefined && dataStyle !== '' && dataStyle !== 'N0') {
      const format = buildNumberFormat(this.numberStyles.get(dataStyle));
      if (format !== undefined && format !== '' && format !== 'General') style.numberFormat = format;
    }

    return Object.keys(style).length === 0 ? null : style;
  }

  synthStyles(): SynthStyle[] {
    return this.list;
  }
}

/* 数字格式                                                                    */

/** `number:*` 定义 → Excel 格式串；结构对不上就返回 undefined（调用方留空，不瞎编）。
 * `style:map`（ODS 用数据格式做条件分支）**不翻译**：Excel 的分节格式串要求三节是正/负/零，
 * 而 `style:map` 可按任意条件跳转，硬凑会让用户看到与源文件不同的数字。这类只丢格式。 */
function buildNumberFormat(entry: { xml: string; el: ElementRange } | undefined): string | undefined {
  if (!entry) return undefined;
  const { xml, el } = entry;
  const children = childElements(xml, el);
  if (children.length === 0) return undefined;
  if (children.some((child) => nameMatches(child.name, 'style:map'))) return undefined;

  const type = localName(el.name);
  const parts: string[] = [];
  let sawNumber = false;
  let integerDigits = 1;
  let decimals = 0;
  let grouping = false;
  let scientific = false;
  let exponentDigits = 2;
  let fraction = false;

  for (const child of children) {
    switch (localName(child.name)) {
      case 'text': {
        const literal = elementText(xml, child);
        if (literal !== '') parts.push(quoteFormatLiteral(literal));
        break;
      }
      case 'number': {
        sawNumber = true;
        integerDigits = numberFrom(child.attrs['number:min-integer-digits']) ?? 1;
        decimals = numberFrom(child.attrs['number:decimal-places']) ?? 0;
        grouping = child.attrs['number:grouping'] === 'true';
        parts.push(numberBody(integerDigits, decimals, grouping));
        break;
      }
      case 'scientific-number': {
        sawNumber = true;
        scientific = true;
        decimals = numberFrom(child.attrs['number:decimal-places']) ?? 0;
        exponentDigits = Math.max(1, numberFrom(child.attrs['number:min-exponent-digits']) ?? 2);
        break;
      }
      case 'fraction': {
        sawNumber = true;
        fraction = true;
        const maxDenominator = Math.max(1, numberFrom(child.attrs['number:max-denominator-value']) ?? 9);
        parts.push(`# ?/${'?'.repeat(String(maxDenominator).length)}`);
        break;
      }
      case 'currency-symbol': {
        const symbol = elementText(xml, child);
        if (symbol !== '') parts.push(quoteFormatLiteral(symbol));
        break;
      }
      case 'year': {
        sawNumber = true;
        parts.push(child.attrs['number:style'] === 'long' ? 'yyyy' : 'yy');
        break;
      }
      case 'month': {
        sawNumber = true;
        // month 的 `number:style` 只有 textual（月份名）与否之分，`long` 只是补零到两位，对应 Excel 的 `mm`
        parts.push(child.attrs['number:style'] === 'textual' ? 'mmm' : 'mm');
        break;
      }
      case 'day': {
        sawNumber = true;
        parts.push('dd');
        break;
      }
      case 'day-of-week': {
        sawNumber = true;
        parts.push(child.attrs['number:style'] === 'long' ? 'dddd' : 'ddd');
        break;
      }
      case 'hours': {
        sawNumber = true;
        // 时长（`[h]:mm:ss`）与时刻（`h:mm`）的分界就在 truncate-on-overflow
        parts.push(el.attrs['number:truncate-on-overflow'] === 'false' ? '[h]' : 'h');
        break;
      }
      case 'minutes': {
        sawNumber = true;
        const style = child.attrs['number:style'];
        parts.push(style === 'long' || style === 'textual' ? 'mm' : 'm');
        break;
      }
      case 'seconds': {
        sawNumber = true;
        const style = child.attrs['number:style'];
        parts.push(style === 'long' || style === 'textual' ? 'ss' : 's');
        break;
      }
      case 'am-pm': {
        parts.push('AM/PM');
        break;
      }
      case 'text-content': {
        parts.push('@');
        break;
      }
      case 'boolean': {
        return undefined;
      }
      default:
        break;
    }
  }

  if (!sawNumber) return undefined;
  if (scientific) {
    return `0${decimals > 0 ? `.${'0'.repeat(decimals)}` : ''}E+${'0'.repeat(exponentDigits)}`;
  }
  if (fraction) return parts.join('');
  let body = parts.join('');
  if (body === '') return undefined;
  if (type === 'percentage-style' && !body.includes('%')) body += '%';
  return body;
}

/** 数字格式里的字面量（货币符号、中文年月日…）在 Excel 里要加引号，否则会被当格式码 */
function quoteFormatLiteral(text: string): string {
  if (text === '') return '';
  if (text === '%' || text === '-' || text === '+' || text === '/' || text === ':' || text === ' ') return text;
  // 纯格式码（`#,##0`）不该被引起来
  if (/^[#0?,.]+$/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

function numberBody(minIntegerDigits: number, decimals: number, grouping: boolean): string {
  const integer = `${grouping ? '#,##' : ''}${'0'.repeat(Math.max(1, minIntegerDigits))}`;
  return decimals > 0 ? `${integer}.${'0'.repeat(decimals)}` : integer;
}

function elementText(xml: string, el: ElementRange): string {
  if (el.selfClosing) return '';
  const raw = xml.slice(el.openEnd, Math.max(el.openEnd, el.closeEnd));
  const close = raw.lastIndexOf('</');
  const inner = close < 0 ? raw : raw.slice(0, close);
  return decodeXmlEntities(inner.replace(/<[^>]*>/g, ''));
}

/* 样式值换算                                                                   */

/** `#RRGGBB` / `#RGB` / `#AARRGGBB` → `#RRGGBB`；`transparent` 表示无填充 */
function normalizeOdsColor(raw: string | undefined): string | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim();
  if (text === '' || text === 'transparent') return undefined;
  const hex = text.replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) return `#${hex}`;
  if (/^[0-9A-F]{8}$/.test(hex)) return `#${hex.slice(2)}`;
  if (/^[0-9A-F]{3}$/.test(hex)) return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
  return undefined;
}

function mapHorizontalAlign(raw: string): SynthStyle['horizontalAlign'] {
  switch (raw) {
    case 'start':
    case 'left':
      return 'left';
    case 'center':
      return 'center';
    case 'end':
    case 'right':
      return 'right';
    default:
      // justify / justify-all 等在中性模型里没有对应值，如实丢弃
      return undefined;
  }
}

function mapVerticalAlign(raw: string): SynthStyle['verticalAlign'] {
  switch (raw) {
    case 'top':
      return 'top';
    case 'middle':
      return 'middle';
    case 'bottom':
      return 'bottom';
    default:
      return undefined;
  }
}

/** `fo:border` / `fo:border-top` … → 中性边框。简写 `fo:border` 的语义是"四边同一根线"，
 * 四边都要写上，只写一边的话 `fo:border="2pt solid #DC2626"` 的边框就只剩一条边。 */
function readBorder(cell: XmlAttributes): SynthStyle['border'] | undefined {
  const shared = cell['fo:border'];
  const sides: Array<'top' | 'right' | 'bottom' | 'left'> = ['top', 'right', 'bottom', 'left'];
  const border: NonNullable<SynthStyle['border']> = {};
  let any = false;
  for (const side of sides) {
    const line = parseBorderLine(cell[`fo:border-${side}`] ?? shared);
    if (!line) continue;
    border[side] = line;
    any = true;
  }
  return any ? border : undefined;
}

/** ODS 里 `thin` / `medium` / `thick` 是**线宽关键字**，不是线型（真夹具写的是 "thin dashed"） */
const BORDER_WIDTH_WORDS: Readonly<Record<string, number>> = { thin: 0.5, medium: 1, thick: 2.5 };

/** `0.5pt solid #000000` / `thin dashed #DC2626` / `2pt solid #DC2626` → `{ style, color }` */
function parseBorderLine(raw: string | undefined): { style: string; color?: string } | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim();
  if (text === '' || text === 'none' || text === 'hidden') return undefined;
  let widthPt: number | undefined;
  let styleName: string | undefined;
  let color: string | undefined;
  for (const token of text.split(/\s+/)) {
    if (token === '') continue;
    if (token.startsWith('#')) {
      color = normalizeOdsColor(token);
      continue;
    }
    const numeric = /^([\d.]+)(pt|px|in|cm|mm)?$/.exec(token);
    if (numeric) {
      widthPt = toPoints(token);
      continue;
    }
    const widthWord = BORDER_WIDTH_WORDS[token];
    if (widthWord !== undefined && widthPt === undefined) {
      widthPt = widthWord;
      continue;
    }
    if (styleName === undefined) styleName = token;
  }
  if (styleName === undefined) return undefined;
  let style = BORDER_STYLE_MAP[styleName];
  if (style === undefined) return undefined;
  if (styleName === 'solid' && widthPt !== undefined) {
    // solid 的粗细只能从线宽看出来：≥2.5pt → thick，≥1pt → medium，0.5pt 的 "thin solid" → thin
    if (widthPt >= 2.5) style = 'thick';
    else if (widthPt >= 1) style = 'medium';
  }
  return color === undefined ? { style } : { style, color };
}

/** `12pt` / `0.5in` / `1.27cm` / `4mm` / `16px` → 磅（无单位按 pt，ODS 的长度默认单位是 cm，但字号默认 pt） */
function toPoints(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const m = /^(-?[\d.]+)\s*(pt|px|in|cm|mm)?$/.exec(raw.trim());
  if (!m) return undefined;
  const value = Number(m[1]);
  if (!Number.isFinite(value)) return undefined;
  switch (m[2]) {
    case 'px': return round2((value * PT_PER_INCH) / 96);
    case 'in': return round2(value * PT_PER_INCH);
    case 'cm': return round2((value / CM_PER_INCH) * PT_PER_INCH);
    case 'mm': return round2((value / MM_PER_INCH) * PT_PER_INCH);
    default: return round2(value);
  }
}

/** 列宽：绝对长度 → Excel 字符宽度。系数由真夹具反解：xlsx 的 `width="18"` 导出成 ODS 后是
 * `3.81cm`，得 1 字符宽 = 0.2117cm ≈ 8px@96dpi = 6pt，用它验其余各列都吻合到 1e-4cm。 */
function toCharWidth(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|px)?$/.exec(raw.trim());
  if (!m) return undefined;
  const value = Number(m[1]);
  if (!Number.isFinite(value)) return undefined;
  let inches: number;
  switch (m[2]) {
    case 'in': inches = value; break;
    case 'mm': inches = value / MM_PER_INCH; break;
    case 'pt': inches = value / PT_PER_INCH; break;
    case 'px': inches = value / 96; break;
    default: inches = value / CM_PER_INCH; break; // ODS 长度默认单位 cm
  }
  const chars = inches / (CM_PER_CHAR_WIDTH / CM_PER_INCH);
  if (!(chars > 0)) return undefined;
  return round2(chars);
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

function firstDefined(...values: Array<string | undefined>): string | undefined {
  for (const value of values) {
    if (value !== undefined && value !== '') return value;
  }
  return undefined;
}

/* 小工具                                                                      */

function cellSpans(attrs: XmlAttributes): { colSpan: number; rowSpan: number } {
  const colSpan = Math.trunc(Math.min(MAX_ODS_COLS, Math.max(1, numberFrom(attrs['table:number-columns-spanned']) ?? 1)));
  const rowSpan = Math.trunc(Math.min(MAX_ODS_ROWS, Math.max(1, numberFrom(attrs['table:number-rows-spanned']) ?? 1)));
  return { colSpan, rowSpan };
}

const cellKey = (row: number, col: number): number => row * MAX_ODS_COLS + col;

/** `A1` 记号用的列字母（`synth-xlsx.ts` 的同名函数没有导出，这里要自己有一份） */
function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function mergeRef(startRow: number, startCol: number, endRow: number, endCol: number): string {
  return `${colLetter(startCol)}${startRow + 1}:${colLetter(endCol)}${endRow + 1}`;
}

/** 把 `A1:B2` 里的行号整体平移（重复行展开时用） */
function shiftMergeRows(ref: string, delta: number): string {
  if (delta === 0) return ref;
  return ref.replace(/([A-Z]+)(\d+)/g, (_all, letters: string, digits: string) => `${letters}${Number(digits) + delta}`);
}

/** 数一数这一行里有多少"非单元格内容"（批注 / 链接 / 图片），只用于报告 */
function countDecorations(contentXml: string, rowEl: ElementRange): number {
  const raw = contentXml.slice(rowEl.openStart, rowEl.closeEnd);
  return (raw.match(/<office:annotation[\s>]/g)?.length ?? 0)
    + (raw.match(/<draw:frame[\s>]/g)?.length ?? 0)
    + (raw.match(/<text:a[\s>]/g)?.length ?? 0);
}

function countHeaderRows(contentXml: string, container: ElementRange): number {
  let count = 0;
  for (const child of childElements(contentXml, container)) {
    if (nameMatches(child.name, 'table:table-row')) {
      count += repeatCount(child.attrs['table:number-rows-repeated']);
    } else if (nameMatches(child.name, 'table:table-rows') || nameMatches(child.name, 'table:table-row-group')) {
      count += countHeaderRows(contentXml, child);
    }
  }
  return count;
}

function repeatCount(raw: string | undefined): number {
  if (raw === undefined) return 1;
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 1) return 1;
  return Math.min(Math.trunc(value), MAX_REPEAT);
}
