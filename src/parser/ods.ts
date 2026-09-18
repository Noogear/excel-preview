/**
 * **OpenDocument Spreadsheet（`.ods`）** 解析器。
 *
 * 产出中性的 `WorkbookInput`（见 `src/importer/synth-xlsx.ts`），随后由 `writeWorkbookPackage`
 * 打成一份规范 xlsx，于是预览/编辑/工作区/撤销/导出/会话恢复全部复用既有链路——ODS 只是
 * 一条"入口翻译"，与 `csv.ts` 同构。
 *
 * 工程上的六个坑（都按**真 Excel 产出的样本**处理，见 `tools/make-legacy-fixtures.ps1`）：
 *
 *  1. **重复展开**（`table:number-columns-repeated` / `table:number-rows-repeated`）：
 *     Excel 导出 ODS 时会把"右侧 / 下方补齐到整张表"的部分写成
 *     `<table:table-cell table:number-columns-repeated="16383"/>` 与
 *     `<table:table-row table:number-rows-repeated="1048555"><table:table-cell
 *     table:number-columns-repeated="16384"/></table:table-row>`，逐格展开就是几十亿个格子。
 *     策略是**截断而不是跳过**：ODS 用"空行 + 重复"表达行间距，直接丢掉会让后面的行
 *     （真夹具里第 20 行的 `行高48` 正是这种情形）坐标整体前移。所以空重复最多推进
 *     `MAX_EMPTY_ROW_RUN` / `MAX_EMPTY_COL_RUN` 步，被截掉的部分如实记进 `warnings`。
 *
 *  2. **合并**（`table:number-columns-spanned` / `table:number-rows-spanned`）：
 *     覆盖区域**只由左上角那个"声明者"产生**，遍历时把光标推过被覆盖的格子并把它们记进
 *     `covered` 集合。绝不按 `table:covered-table-cell` 反推区域——它只说明"这里被盖住了"，
 *     连盖了几个都不知道（真夹具里它是 `number-columns-repeated="5"` 的裸元素）。
 *
 *  3. **公式**：ODS 用 OpenFormula，`table:formula="of:=SUM([.A1:.A3])"`。只翻译
 *     `of:` / `oooc:` 前缀且**能逐 token 校验**的引用与运算符；任何看不懂的构造
 *     （命名表达式、实现私有函数、`oooc:` 的正则与通配符扩展…）一律**丢掉公式、保留缓存值**，
 *     并计入 `stats.formulaDropped`。宁可少一个公式，也不能把错的公式写进用户文件。
 *
 *  4. **日期与时间**：ODS 用 ISO 文本（`office:date-value="2025-01-01T00:00:00"`、
 *     `office:time-value="PT18H0M0S"`），Excel 用序列号。这里换算成序列号，否则单元格退化成
 *     文本、下游的数字格式也就失去意义。`PT36H0M0S` 这类**超过 24 小时**的时长同样按时长处理
 *     （1.5 天），不能取模成 0.5。
 *
 *  5. **样式来源两处**：`content.xml` 的 `office:automatic-styles`（Excel 给每个用到的样式写一条
 *     定义，**就近覆盖**）与 `styles.xml` 的 `office:styles`（命名样式 + 全部 `number:*` 数字格式
 *     定义）。同名以 `content.xml` 为准；`style:parent-style-name` 链要一路继承——真夹具里
 *     `ce1 → Default` 的字体与字号就来自父样式，不继承就会丢掉整张表的默认字体。
 *
 *  6. **列宽换算**：ODS 存绝对长度（`style:column-width="3.81cm"`），Excel 存字符宽度。
 *     系数由真夹具**反解**得到：xlsx 的 `width="18"`（Excel 字符宽度）导出成 ODS 后是 `3.81cm`，
 *     反解 1 字符宽 = 0.2117cm = 6pt = 8px(96dpi)。用同一系数验其余各列
 *     （14→2.9633、22→4.6567、30→6.35、9→1.905、20→4.2333、16→3.3867、12→2.54cm）全部吻合到
 *     1e-4cm，故取 `chars = cm / 0.2117`。往返误差只有 0.02 字符宽，来源是 Excel 自己导出时
 *     把长度值量化到了 0.01cm 量级。
 *
 * 明确不支持（如实丢弃 + 记 warning，绝不猜值）：条件格式（`style:map`）、数据验证
 * （`table:content-validation`）、批注（`office:annotation`）、超链接 URL（`text:a`，
 * 文本本身保留）、图片与图形（`draw:frame`）、命名区域、透视表、隐藏行的隐藏状态
 * （中性模型只有行高与列宽，没有"隐藏行"字段）、`table:table-header-columns` 的冻结列语义。
 */
import { childElements, findFirstElement, localName, nameMatches } from './xml';
import { readText, readZipEntries, type ZipEntries } from './zip';
import type { ElementRange, XmlAttributes } from './xml';
import type { SynthCell, SynthSheet, SynthStyle, WorkbookInput } from '../importer/synth-xlsx';

/* -------------------------------------------------------------------------- */
/* 公开常量与结果类型                                                           */
/* -------------------------------------------------------------------------- */

/** 单次导入的单元格总量上限（与 `csv.ts` 的 `MAX_DELIMITED_CELLS` 同量级、同语义） */
export const MAX_ODS_CELLS = 200_000;
/** 单边上限（与 xlsx 规范一致） */
export const MAX_ODS_ROWS = 1_048_576;
export const MAX_ODS_COLS = 16_384;
/**
 * 连续空列 / 空行的展开上限。
 *
 * 取的就是 ODS 文本在 Excel 里的自然边界（补齐写法是 16384 列 / 16384 行），于是真文件里的
 * 补齐能**完整且忠实地**展开，而损坏或恶意文件最多也只多走 16384 步。
 */
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
  /** 被安全阀截断的次数（列重复 + 行重复 + 单元格总量） */
  truncatedRepeats: number;
  /** 是否发生了任何截断 */
  truncated: boolean;
}

export interface OdsParseResult {
  input: WorkbookInput;
  /** 给人看的降级说明（不是异常；进"导入摘要"） */
  warnings: string[];
  stats: OdsStats;
}

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 解析 `.ods` 字节，返回中性工作簿 + 降级报告。
 *
 * 不抛异常：不是 zip、缺 `content.xml`、XML 结构怪异……都退化成"空工作簿 + warning"，
 * 与 `parseXlsx` 的容错原则一致（局部坏掉不该让整次导入失败）。
 */
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

/**
 * 主入口（固定签名）：字节 → 中性工作簿。
 * 需要降级报告或统计时用 `parseOdsResult`，两者的 `input` 完全一致。
 */
export function parseOds(bytes: Uint8Array): WorkbookInput {
  return parseOdsResult(bytes).input;
}

/* -------------------------------------------------------------------------- */
/* 表格遍历                                                                     */
/* -------------------------------------------------------------------------- */

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
  /**
   * 列宽先按 `table:table-column` 的声明收集，**等行都解析完**再按实际用到的列裁剪。
   *
   * 为什么必须裁：ODS 把"整表补齐"的列也写成一条 `number-columns-repeated="16376"` 的列定义，
   * 按声明落进 `colWidths` 就会输出 16384 条列宽（真夹具实测），而下游 writer 会把它变成
   * 16384 个 `<col>` 元素。裁剪到"有内容的最后一列"才是用户能看见的那几列。
   */
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
        // 纯空行：只推进光标（见文件头第 1 条：空行承载"行间距"语义，不能直接丢）。
        // 注意这里**不看行高**：空行上的行高是 Excel 补齐时写的默认值，照收会让
        // `rowHeights` 从"用户设过的 4 行"膨胀成"几万行全有行高"。
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

/* -------------------------------------------------------------------------- */
/* 列                                                                          */
/* -------------------------------------------------------------------------- */

interface ColumnResult {
  cursor: number;
  /** 因为超过空列展开上限而被截掉的列数 */
  clamped: number;
  /** 被标记为隐藏的列数（中性模型没有隐藏列字段，只用于报告） */
  hidden: number;
}

/**
 * 一个 `table:table-column`：展开 `number-columns-repeated`，把宽度写进 `colWidths`。
 *
 * 样式里没有列宽时（例如整表补齐用的裸列）**留空**，让下游用默认列宽——把默认值当"用户设过的
 * 宽度"写出去，Excel 里每一列都会变成"自定义宽度"，用户再也点不回"自动"。
 */
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

/* -------------------------------------------------------------------------- */
/* 行与单元格                                                                   */
/* -------------------------------------------------------------------------- */

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
      // 被覆盖的位置不是内容；只有**带位置语义样式**的 covered 格子才占位（那代表用户设过底色
      // 或对齐）。裸的 `<table:covered-table-cell table:number-columns-repeated="16383"/>`
      // 与只带默认字体的占位符一律跳过，否则一处合并就会造出上万个空单元格。
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

/**
 * 这个 `table:table-cell` 是不是一个"纯占位空格子"。
 *
 * ODS（Excel 导出尤其明显）会给每一行的尾部补齐
 * `<table:table-cell table:number-columns-repeated="16383"/>`，也会在正文左侧写
 * `<table:table-cell table:number-columns-repeated="3" table:style-name="ce1"/>` 这种跳格占位符。
 * 把它们当单元格收下，真夹具里 30 个真值格子会膨胀成 59 个"空但带默认样式"的假格子（实测），
 * 预览里就多出一片本来不存在的格子。
 *
 * 判据分两种：
 *  ① 完全没有样式引用 → 一定是占位；
 *  ② 样式只描述了"字体"，**没有任何位置语义**（无填充、无对齐、无边框、无格式）→ 也是占位。
 *     这一条是关键：ODS 里每个格子都会带一个继承自 `Default` 的 `ce1`（字体 宋体 11pt），
 *     光看"有没有样式"判不出来。而带填充 / 对齐的空格子是用户真的设过底色或居中，
 *     必须保留——真夹具 `fixture-rules` 的 C1 就是这种。
 */
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

/**
 * 读一个 `table:table-cell`：值、公式、样式。
 *
 * 值优先信 `office:value*` 属性（那是规范里的权威值），属性缺失时退回 `<text:p>` 文本——
 * 手工写的 ODS 常常只写文本、不写 `office:value-type`。
 */
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

/* -------------------------------------------------------------------------- */
/* 单元格文本                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * 取单元格的可见文本：拼接所有 `<text:p>`（段间换行），展开 `<text:s text:c="3"/>`、
 * `<text:tab/>`、`<text:line-break/>`，保留 `<text:span>` / `<text:a>` 里的文字。
 *
 * 为什么自己扫一遍而不复用 `elementText`：ODS 的空白语义全在这些自闭合元素上，
 * `elementText` 会把它们剥成空串（`<text:s/>` 一个空格就没了），真夹具里
 * "粗蓝 + 斜体 + 下划线绿" 会粘成 "粗蓝+ 斜体+ 下划线绿"。所以这里按标签逐段取，
 * 并**跳过元素之间的排版空白**（缩进换行），只保留元素内部的文本。
 */
export function readCellText(contentXml: string, cellEl: ElementRange): string {
  if (cellEl.selfClosing) return '';
  const start = cellEl.openEnd;
  const end = Math.max(cellEl.openEnd, cellEl.closeEnd);
  /**
   * 段落文本按"最外层 `<text:p>`"分段收集，最后用换行拼起来。
   *
   * 为什么不边扫边补换行：`<text:p>` 是可以嵌套的（列表项、`<text:list>`），
   * 靠"上一个标签是不是 `</text:p>`"来判断段界会在嵌套时多补或漏补换行。
   * 按深度归段是纯局部的：深度从 0 变 1 开一段，回到 0 收一段。
   */
  const paragraphs: string[] = [];
  let current = '';
  /** `<text:p>` 的嵌套深度（> 0 表示正在收集属于单元格的段落文本） */
  let paragraphDepth = 0;
  /**
   * 遮挡容器的嵌套深度（`<office:annotation>`、`<draw:frame>`、`<svg:desc>` …）。
   * 这些容器内部也有 `<text:p>`，但那是批注 / 图形说明，不是单元格的值——真夹具里一条两行的
   * 批注会被拼成 "这是一条批注（legacy note）第二行带批注的单元格" 这种把批注和正文粘起来的假值。
   */
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

  // 只裁每一段首尾的**普通空白**：那是 XML 缩进。用户真正想要的空格由 `<text:s/>` 表达，
  // 而它展开后同样落在段内（甚至段首），所以不能对整串做 trim，否则会吃掉用户的前导空格。
  const trimmed = paragraphs
    .map((part) => part.replace(/^[ \t]+/, '').replace(/[ \t\n]+$/, ''))
    .filter((part) => part !== '');
  const joined = trimmed.join('\n');
  return joined.length > MAX_TEXT_LENGTH ? joined.slice(0, MAX_TEXT_LENGTH) : joined;
}

/**
 * 不算单元格值的块级容器：它们内部也有 `<text:p>`，但那是批注 / 图形说明 / 索引，
 * 拼进来会造出"批注 + 正文"粘在一起的假值（真夹具验证过）。
 */
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

/* -------------------------------------------------------------------------- */
/* 值换算                                                                       */
/* -------------------------------------------------------------------------- */

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

/**
 * ISO 日期（`2025-01-01` / `2025-01-01T12:00:00`，可带小数秒）→ Excel 序列号。
 *
 * 基准取 1899-12-30（Excel 的 1900 闰年 bug 之后一天），于是 1900-03-01 起的日期与 Excel
 * 完全一致——Excel 自己导出的 ODS 不会早于 1900，这个基准足够且可测。
 * 时间部分按 UTC 解析：ODS 的 `date-value` 是不带偏移的本地时间，UTC 解析等于原样取用。
 */
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

/**
 * ISO 8601 时长（`PT18H0M0S`）→ 天（Excel 的时间值）。
 * `PT36H0M0S` 是 1.5 天而不是 0.5 天：ODS 的时长允许超过 24 小时，Excel 的 `[h]:mm:ss`
 * 也这么用，取模会把"36 小时"变成"12 小时"。
 */
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

/* -------------------------------------------------------------------------- */
/* 公式：OpenFormula → Excel A1 公式                                            */
/* -------------------------------------------------------------------------- */

/** 明确拒绝的函数名：OpenFormula 独有、实现私有、或与 Excel 同名函数的语义不同 */
const FORBIDDEN_FUNCTIONS: ReadonlySet<string> = new Set([
  'REGEX', 'DDE', 'BAHTTEXT', 'MULTIPLE.OPERATIONS', 'COM.MICROSOFT.EXCEL',
]);

/**
 * 把 `table:formula` 的 OpenFormula 表达式翻成 Excel 的 A1 公式（**不含前导 `=`**）；
 * 无法可靠翻译时返回 `undefined`（调用方丢掉公式、保留缓存值）。
 *
 * 翻译规则（逐 token，不做正则整体替换）：
 *  - 引用 `[.A1]` / `[.A1:.A3]` / `[第二张.A2]` / `['我的 表'.A1]` → `A1` / `A1:A3` / `第二张!A2`；
 *  - 参数分隔 `;` → `,`（真夹具里 `IF([.A2]>50;"大";"小")` 就走这条）；
 *  - `&`（拼接）、`+ - * / ^ = <> <= >= < > ( )` 两边语义一致，原样输出；
 *  - 引号字符串原样保留（两边都用 `"` 且都用 `""` 转义）；
 *  - 出现 `!`（OpenFormula 的"不等于"）、`|`（联合）、`~`（通配符）、`{}`（数组）、
 *    含 `.` 或形如单元格引用的裸标识符（命名表达式）一律判定为不可翻译。
 */
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
  /** 表名 + `!`（必要时带单引号） */
  sheetPrefix?: string;
  /** 表名原文，用于校验范围两端一致 */
  sheet?: string;
}

/**
 * 方括号里的引用 → A1 记号（跨表用 Excel 的 `!`）。
 *
 * 支持 `.A1`、`.A1:.A3`、`.$A$1:.$B$2`、`第二张.A1`、`'我的 表'.A1`、`.A:A`、`.1:3`；
 * 命名表达式（`[销售额]`）返回 undefined，由调用方判定为不可翻译。
 */
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

/* -------------------------------------------------------------------------- */
/* 样式                                                                        */
/* -------------------------------------------------------------------------- */

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

/**
 * 样式解析：把 ODS 的"样式名 → 属性链"变成中性的 `SynthStyle`。
 *
 * 三条关键约定：
 *  - **就近覆盖**：同名样式在 `content.xml` 与 `styles.xml` 都有定义时用 `content.xml` 的
 *    （ODS 的 `office:automatic-styles` 描述"这份文档实际用的样式"，优先级更高）；
 *    同一份文档里后出现的覆盖先出现的。
 *  - **继承**：`style:parent-style-name` 链从根到叶逐层覆盖；不继承会丢掉整张表的默认字体
 *    （真夹具里 `ce1` 自己不写字体，全靠父样式 `Default`）。
 *  - **去重**：解析结果按 JSON 串 intern，`SynthCell.style` 只存下标，
 *    上万个同款单元格共用一条样式，样式表不会膨胀。
 */
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

  /**
   * 行属性：**只有显式自定义行高才保留**。
   *
   * 为什么看 `style:use-optimal-row-height`：ODS 里每一行都会带一条行高（哪怕是自动算出来的
   * 13.5pt）。把它们全当成"用户设过的行高"写回 xlsx，每一行都会变成 `customHeight="1"`，
   * 用户在 Excel 里再也点不回"自动调整行高"。真夹具验证：`use-optimal-row-height="false"`
   * 的 4 行恰好就是 xlsx 真值里有 `ht` 的 4 行（32 / 44 / 30 / 48）。
   */
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

  /**
   * 把一个 `table:style-name` 解析成 `SynthStyle` 下标；没有任何可表达样式时 `index = -1`。
   * 返回 `-1` 让调用方**不写** `SynthCell.style`，于是空格子不会被造出来、默认样式不会被写出去。
   *
   * `visual` 表示这条样式是否带"位置语义"（填充 / 对齐 / 边框 / 旋转 / 数字格式）。
   * 只带字体的样式等价于 ODS 的默认单元格样式，`parseRow` 靠它区分"真格子"与"跳格占位符"。
   */
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

/* -------------------------------------------------------------------------- */
/* 数字格式                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * `number:*` 定义 → Excel 格式串；结构对不上就返回 undefined（调用方留空，不瞎编）。
 *
 * 已知取舍：`style:map`（ODS 用数据格式做条件分支，例如"负值套另一个样式"）**不翻译**——
 * Excel 的分节格式串要求三节分别是正/负/零，而 `style:map` 可以按任意条件跳转，
 * 硬凑出来的串会让用户看到与源文件不同的数字。这类单元格只丢格式、值照旧。
 */
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
        // ODS 的 `number:style` 在 month 上只有 textual（月份名）与否之分；
        // `long` 只是"补零到两位"，对应 Excel 的 `mm`（真夹具 N49 验证）
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

/* -------------------------------------------------------------------------- */
/* 样式值换算                                                                   */
/* -------------------------------------------------------------------------- */

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

/**
 * `fo:border` / `fo:border-top` … → 中性边框。
 *
 * 简写 `fo:border` 的语义是"四边同一根线"，所以四边都要写上——只写一边的话，
 * 真夹具里 `fo:border="2pt solid #DC2626"` 的 medium 边框就只剩一条边。
 */
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
    // solid 的粗细只能从线宽看出来。分档与真夹具对齐：0.5pt → thin（ODS 的 "thin solid"）、
    // 2pt → medium（`fixture-styles` 的 ce8，xlsx 真值是 medium）、3pt 及以上 → thick。
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

/**
 * 列宽：绝对长度 → Excel 字符宽度。
 *
 * 系数由真夹具**反解**：xlsx 真值里的 `width="18"`（Excel 字符宽度单位）导出成 ODS 后是
 * `3.81cm`，两者相除得 1 字符宽 = 0.211667cm；用同一系数验其余各列
 * （14→2.9633、22→4.6567、30→6.35、9→1.905、20→4.2333、16→3.3867、12→2.54）全部吻合，
 * 而这个系数正好等于 **0.2117cm ≈ 8px at 96dpi ≈ 6pt** —— 即 1 字符宽 = 6pt。
 */
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

/* -------------------------------------------------------------------------- */
/* 小工具                                                                      */
/* -------------------------------------------------------------------------- */

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
