/**
 * 外科式修补导出：把预览里的单元格编辑写回原 xlsx。
 *
 * 设计原则
 * --------
 * 1. **只重写必须改的部件**：被编辑工作表的 `xl/worksheets/sheetN.xml`、
 *    `xl/sharedStrings.xml`（且只**追加** `<si>`）、`xl/workbook.xml`（只动 `<calcPr>`）。
 *    其余 zip 条目**逐字节原样写回**——图表、透视表、形状、条件格式、VBA 等我们没解析的
 *    部件因此不会被"重新序列化"破坏。
 * 2. **按 row / cell 定位的区间拼接**：不做整篇 XML 的正则替换，`sheetData` 之外的所有
 *    字节（`<cols>`、`<mergeCells>`、`<pageSetup>`、`<extLst>`…）保持原文。
 * 3. **不改动入参**：`source.raw.entries` / `source.sheets` 只读，输出是全新的 `Uint8Array`。
 *
 * `<dimension>` 策略（题目允许二选一，这里选"按需扩展"）
 * ------------------------------------------------------
 * 当且仅当存在**落在原 dimension 之外**的编辑目标时，才把 `ref` 扩成
 * 「原范围 ∪ 本次编辑坐标」的最小外接矩形；只改/清空范围内单元格时 `<dimension>` 一个
 * 字节都不动。理由：dimension 是给读取方的提示，写小了会让 Excel 误判工作表尺寸
 * （滚动条、选区、个别版本提示"需要修复"）；而每次都重算又会让范围内的普通编辑白白
 * 多改一段本来不必改的字节。
 *
 * 已知取舍
 * --------
 * - 新增单元格不带 `s` 样式属性（编辑模型里没有"新单元格样式"这一项，凭空造一个会污染
 *   styles.xml；如需样式请由调用方在导入侧先把样式格子建出来）。
 * - 写公式时不写缓存值：属于该格子的旧 `<v>` 缓存已失效、会被删除，重算交给
 *   `<calcPr fullCalcOnLoad="1"/>`。
 * - 字符串一律走共享字符串表（`t="s"`）；命中已有相同文本则复用其索引，否则追加到末尾。
 */
import { zipSync, strFromU8, strToU8 } from 'fflate';

import { parseCellRef } from '../parser/worksheet';
import { dirName, findEntry, resolvePath, type ZipEntries } from '../parser/zip';
import {
  attrInt, childElements, findFirstElement, localName, nameMatches,
} from '../parser/xml';

import type { ElementRange } from '../parser/xml';
import type { ParsedCell, ParsedSheet, ParsedWorkbook } from '../parser/types';
import { readZipEntries } from '../parser/zip';

/* -------------------------------------------------------------------------- */
/* 对外契约                                                                    */
/* -------------------------------------------------------------------------- */

export interface CellEdit {
  row: number;            // 0-based
  col: number;            // 0-based
  /** 新值；null 表示清空内容（**必须保留该单元格原有的 s 样式属性**） */
  value?: string | number | boolean | null;
  /** 新公式，不含前导 '='；null 表示删除公式 */
  formula?: string | null;
  /**
   * 字符串是否强制走共享字符串表（默认自动：能复用到就复用，否则追加）。
   *
   * 说明：本实现里字符串**始终**写入共享字符串表（`t="s"`）——这样复用与去重
   * 是同一个逻辑，也不会为了一个格子去改写 `styles.xml` 之外的部件。
   * 唯一不走共享字符串表的情况是"公式的字符串结果缓存"（`t="str"`，值必须内联），
   * 那种情况传 `true` 也不会改变行为（规范不允许把公式结果放进共享字符串表）。
   */
  forceSharedString?: boolean;
}

export interface SheetEdits { sheetId: string; cells: CellEdit[] }

/**
 * 导出所需的**最小**数据源。
 *
 * 为什么要有这个类型：完整 `ParsedWorkbook` 里百万个 `ParsedCell` 对象是常驻内存的大头，
 * 而外科式导出真正用到的只有三样——zip 条目、`sheet.id`、以及"每行第一个带样式的 s"
 * （`collectRowStyles` 用给新格子补样式）。所以导入完成后可以把模型瘦身成这个形状
 * （见 `slimForExport`），内存降下来，导出结果一字不差。
 *
 * 两种给条目的方式：
 * - `raw.entries`：已经在内存里（直接把完整 `ParsedWorkbook` 传进来时走这条）；
 * - `bytes`：原始 xlsx 字节，导出时**惰性解压**。瘦身路径用这个——省下几十 MB 常驻内存，
 *   代价是每次导出多花一次解压（6 MB 文件实测约 0.2 s）。
 */
export interface ExportSource {
  raw?: { entries: ZipEntries };
  bytes?: Uint8Array;
  sheets: readonly ExportSourceSheet[];
}

export interface ExportSourceSheet {
  id: string;
  /** 行号 → 该行第一个带样式单元格的 `s` */
  rowStyles?: ReadonlyMap<number, number>;
  /** 兼容直接传完整 `ParsedSheet` 的调用方：只读它现算 `rowStyles` */
  cells?: readonly ParsedCell[];
}

export interface ExportOptions {
  /** 是否写入 <calcPr fullCalcOnLoad="1"/> 让 Excel 打开时重算（默认 true） */
  fullCalcOnLoad?: boolean;
}

/** 归一化后的编辑；同一坐标重复出现时后者覆盖前者 */
interface ResolvedEdit {
  row: number;
  col: number;
  value: string | number | boolean | null | undefined;
  formula: string | null | undefined;
}

const DEFAULT_WORKBOOK_PATH = 'xl/workbook.xml';
const DEFAULT_SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';
/** Excel 上限，仅用于挡住明显越界的输入 */
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const SST_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

/** 返回新的 xlsx 字节；不改动入参 */
export function exportXlsx(
  source: ExportSource,
  edits: SheetEdits[],
  options: ExportOptions = {},
): Uint8Array {
  const fullCalcOnLoad = options.fullCalcOnLoad !== false;
  // 瘦身路径（只有 bytes）时在这里惰性解压；两种来源的条目内容完全一致
  const entries = source.raw?.entries ?? (source.bytes ? readZipEntries(source.bytes) : undefined);
  if (!entries || Object.keys(entries).length === 0) {
    throw new Error('exportXlsx：source.raw.entries 为空，没有可修补的原始 xlsx 条目');
  }

  // 1) 先确定要改哪几张表（sheetId 找不到对应部件时抛错，绝不静默跳过）
  const targets = new Map<string, { entryName: string; edits: ResolvedEdit[] }>();
  for (const sheetEdits of edits) {
    const resolved = resolveEdits(sheetEdits);
    if (resolved.length === 0) continue;
    const known = targets.get(sheetEdits.sheetId);
    if (known) {
      known.edits = dedupeEdits([...known.edits, ...resolved]);
      continue;
    }
    targets.set(sheetEdits.sheetId, {
      entryName: resolveSheetEntry(entries, sheetEdits.sheetId),
      edits: dedupeEdits(resolved),
    });
  }

  // 2) 共享字符串表：按需追加（已有索引与顺序一律不动）
  const sst = createSharedStringEditor(entries, targets.size > 0);

  // 3) 逐表修补 sheet XML（内容没变的表不登记，保持原字节）
  const replaced = new Map<string, Uint8Array>();
  for (const { entryName, edits: sheetCellEdits } of targets.values()) {
    const bytes = findEntry(entries, entryName);
    if (!bytes) throw new Error(`exportXlsx：工作表部件 "${entryName}" 在 zip 中不存在`);
    const next = strToU8(
      patchWorksheetXml(strFromU8(bytes), sheetCellEdits, sst, source.sheets),
    );
    if (!sameBytes(next, bytes)) replaced.set(entryName, next);
  }
  if (sst.changed && sst.entryName && sst.bytes) replaced.set(sst.entryName, sst.bytes);

  // 4) workbook.xml：只动 <calcPr>
  const workbookEntry = resolveWorkbookEntry(entries);
  const workbookBytes = findEntry(entries, workbookEntry);
  if (!workbookBytes) throw new Error(`exportXlsx：找不到工作簿部件 "${workbookEntry}"`);
  if (fullCalcOnLoad) {
    const next = strToU8(patchCalcPr(strFromU8(workbookBytes)));
    if (!sameBytes(next, workbookBytes)) replaced.set(workbookEntry, next);
  }

  // 5) 按原始条目顺序重建 zip：没命中的条目原字节写回，命中的换成新字节
  const out: ZipEntries = {};
  for (const name of Object.keys(entries)) out[name] = replaced.get(name) ?? entries[name];
  // 兜底：条目名大小写/前导斜杠写法与替换表 key 不一致时补进去（不改顺序）；
  // 只有 findEntry 确认"压根没有这个条目"时才追加，避免写出重名条目
  for (const [name, bytes] of replaced) {
    if (name in out || findEntry(entries, name)) continue;
    out[name] = bytes;
  }
  return zipSync(out);
}

/* -------------------------------------------------------------------------- */
/* 部件定位                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * `ParsedSheet.id` 是 workbook.xml 里的 `r:id`（如 `rId4`）；
 * 这里顺着 `_rels/.rels` -> workbook.xml -> workbook.xml.rels 把它解析成 zip 条目名。
 */
function resolveSheetEntry(entries: ZipEntries, sheetId: string): string {
  const workbookEntry = resolveWorkbookEntry(entries);
  const workbookBytes = findEntry(entries, workbookEntry);
  if (!workbookBytes) throw new Error(`exportXlsx：找不到工作簿部件 "${workbookEntry}"`);
  const workbookXml = strFromU8(workbookBytes);

  const relsEntry = `${dirName(workbookEntry)}/_rels/${baseName(workbookEntry)}.rels`;
  const relsBytes = findEntry(entries, relsEntry);
  const relsXml = relsBytes ? strFromU8(relsBytes) : undefined;

  const target = sheetIdToTarget(workbookXml, relsXml, sheetId);
  if (!target) {
    throw new Error(
      `exportXlsx：找不到 sheetId "${sheetId}" 对应的工作表部件`
      + `（已查 ${workbookEntry}${relsXml ? ` 与 ${relsEntry}` : '，且缺少 rels'}）`,
    );
  }
  const entryName = resolvePath(dirName(workbookEntry), target);
  if (!findEntry(entries, entryName)) {
    throw new Error(`exportXlsx：sheetId "${sheetId}" 指向的部件 "${entryName}" 不在 zip 内`);
  }
  return entryName;
}

/** workbook.xml 的 `<sheet r:id>` -> rels 里的 target（相对 workbook 部件的路径） */
function sheetIdToTarget(
  workbookXml: string,
  relsXml: string | undefined,
  sheetId: string,
): string | undefined {
  void workbookXml; // sheetId 本身来自 ParsedSheet.id，无需再回读 workbook.xml
  if (!relsXml) return undefined;
  const root = findFirstElement(relsXml, 'Relationships');
  if (!root) return undefined;
  for (const rel of childElements(relsXml, root)) {
    if (localName(rel.name) !== 'Relationship') continue;
    if (rel.attrs['Id'] !== sheetId) continue;
    if (rel.attrs['TargetMode'] === 'External') return undefined; // 外链工作表不是本地部件
    return rel.attrs['Target'];
  }
  return undefined;
}

/** `_rels/.rels` 的 officeDocument 关系 -> workbook 部件（退化为约定路径） */
function resolveWorkbookEntry(entries: ZipEntries): string {
  const rootRelsBytes = findEntry(entries, '_rels/.rels');
  if (rootRelsBytes) {
    const rootRels = strFromU8(rootRelsBytes);
    const root = findFirstElement(rootRels, 'Relationships');
    if (root) {
      for (const rel of childElements(rootRels, root)) {
        if (localName(rel.name) !== 'Relationship') continue;
        const type = rel.attrs['Type'] ?? '';
        if (!type.endsWith('officeDocument')) continue;
        const target = rel.attrs['Target'];
        if (!target) continue;
        const candidate = normalizeEntryPath(target);
        if (findEntry(entries, candidate)) return candidate;
      }
    }
  }
  return DEFAULT_WORKBOOK_PATH;
}

/* -------------------------------------------------------------------------- */
/* 编辑归一化                                                                  */
/* -------------------------------------------------------------------------- */

function resolveEdits(sheetEdits: SheetEdits): ResolvedEdit[] {
  const out: ResolvedEdit[] = [];
  for (const cell of sheetEdits.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.col) || cell.row < 0 || cell.col < 0) {
      throw new Error(
        `exportXlsx：编辑坐标非法（row=${String(cell.row)}, col=${String(cell.col)}）`,
      );
    }
    if (cell.row >= MAX_ROW || cell.col >= MAX_COL) {
      throw new Error(`exportXlsx：编辑坐标超出 xlsx 上限（row=${cell.row}, col=${cell.col}）`);
    }
    out.push({
      row: cell.row,
      col: cell.col,
      value: typeof cell.value === 'undefined' ? undefined : cell.value,
      formula: typeof cell.formula === 'undefined' ? undefined : cell.formula,
    });
  }
  return out;
}

/** 同坐标去重：后面的覆盖前面的 */
function dedupeEdits(list: ResolvedEdit[]): ResolvedEdit[] {
  const byKey = new Map<string, ResolvedEdit>();
  for (const edit of list) byKey.set(cellKey(edit.row, edit.col), edit);
  return [...byKey.values()];
}

/* -------------------------------------------------------------------------- */
/* 工作表 XML 修补                                                             */
/* -------------------------------------------------------------------------- */

function patchWorksheetXml(
  xml: string,
  edits: ResolvedEdit[],
  sst: SharedStringEditor,
  sourceSheets: readonly ExportSourceSheet[],
): string {
  const sheetData = findFirstElement(xml, 'sheetData');
  if (!sheetData) throw new Error('exportXlsx：工作表缺少 <sheetData>，无法写入单元格');

  // 0) 样式来源：本次编辑落在的行里，原有单元格的样式索引（用于"新格子"补齐 s）
  const rowStyles = collectRowStyles(sourceSheets);
  const styleFor = (row: number): number | undefined => rowStyles.get(row);

  // 1) 先算好每个目标单元格的新内容（此时还不碰原文）
  const planned = new Map<string, CellContent>();
  for (const edit of edits) {
    planned.set(`${colLetters(edit.col)}${edit.row + 1}`, buildCellContent(edit, sst));
  }

  // 2) 逐行修补已存在的 <row>
  const touchedRows = new Set<number>();
  const rowResults: Array<{ rowEl: ElementRange; text: string }> = [];
  for (const rowEl of childElements(xml, sheetData)) {
    if (!nameMatches(rowEl.name, 'row')) continue;
    const rowIndex = rowNumberOf(rowEl);
    const text = rowIndex === undefined
      ? xml.slice(rowEl.openStart, rowEl.closeEnd)
      : patchRow(xml, rowEl, rowIndex, planned, touchedRows, styleFor(rowIndex));
    rowResults.push({ rowEl, text });
  }

  // 3) 目标行原本不存在 -> 整行新建，按行号升序插进 sheetData
  const newRows: Array<{ rowIndex: number; cells: Array<{ col: number; text: string }> }> = [];
  for (const [ref, content] of planned) {
    const coord = parseCellRef(ref);
    if (!coord || touchedRows.has(coord.row)) continue;
    let row = newRows.find((item) => item.rowIndex === coord.row);
    if (!row) {
      row = { rowIndex: coord.row, cells: [] };
      newRows.push(row);
    }
    row.cells.push({ col: coord.col, text: serializeRefCell(ref, content, styleFor(coord.row)) });
  }
  newRows.sort((a, b) => a.rowIndex - b.rowIndex);

  const inner = xml.slice(sheetData.openEnd, innerEnd(xml, sheetData));
  // 注意：rowResults 里的区间是**整篇 xml** 的坐标，而 inner 是从 openEnd 开始的切片，
  // 所以要减去基准偏移再交给 rebuildSheetData（否则会按错位置切片，导致整行重复）。
  const sheetDataInner = rebuildSheetData(inner, newRows, rowResults, sheetData.openEnd);

  // 4) 拼接：sheetData 之外的所有字节原样保留；dimension 只在明显不匹配时才动
  //    注意 `<sheetData/>` 要先展开成 `<sheetData>`，否则内容会被写到标签外面
  const open = sheetData.selfClosing
    ? openTag(xml, sheetData).replace(/\/\s*>$/, '>')
    : openTag(xml, sheetData);
  const rebuilt = open + sheetDataInner
    + (sheetData.selfClosing ? '</sheetData>' : xml.slice(innerEnd(xml, sheetData), sheetData.closeEnd));
  return updateDimension(
    xml.slice(0, sheetData.openStart) + rebuilt + xml.slice(sheetData.closeEnd),
    planned,
  );
}

/** 一个目标单元格算出来的内容 */
interface CellContent {
  value: string | number | boolean | null;
  /** `t` 属性：s=共享字符串 / str=公式字符串结果 / b=布尔 / undefined=数字或空 */
  type: 's' | 'str' | 'b' | undefined;
  formula: string | undefined;
  /** 字符串走共享字符串表时记下的索引 */
  sstIndex: number | undefined;
}

function buildCellContent(edit: ResolvedEdit, sst: SharedStringEditor): CellContent {
  const formula = typeof edit.formula === 'string' ? stripFormula(edit.formula) : undefined;
  const hasFormula = formula !== undefined && formula.length > 0;
  const value = edit.value === undefined ? null : edit.value;

  if (typeof value === 'string') {
    if (value.length === 0) {
      // 空串等同于清空内容（Excel 里空字符串单元格本来就是空）
      return { value: null, type: undefined, formula: hasFormula ? formula : undefined, sstIndex: undefined };
    }
    if (hasFormula) {
      // 公式的字符串结果缓存：规范写法是 t="str" + <v>文本</v>
      return { value, type: 'str', formula, sstIndex: undefined };
    }
    return { value, type: 's', formula: undefined, sstIndex: sst.intern(value) };
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new Error(`exportXlsx：单元格值不是有限数字（row=${edit.row}, col=${edit.col}）`);
    }
    return { value, type: undefined, formula: hasFormula ? formula : undefined, sstIndex: undefined };
  }
  if (typeof value === 'boolean') {
    return { value, type: 'b', formula: hasFormula ? formula : undefined, sstIndex: undefined };
  }
  return { value: null, type: undefined, formula: hasFormula ? formula : undefined, sstIndex: undefined };
}

/* -------------------------------------------------------------------------- */
/* 单个 <row> 修补                                                             */
/* -------------------------------------------------------------------------- */

function patchRow(
  xml: string,
  rowEl: ElementRange,
  rowIndex: number,
  planned: Map<string, CellContent>,
  touchedRows: Set<number>,
  rowStyle: number | undefined,
): string {
  const inner = xml.slice(rowEl.openEnd, innerEnd(xml, rowEl));
  /** 行内容的改动：原文区间 `[start, end)`（相对 inner）替换成 `text`；`start === end` 即纯插入 */
  interface RowPatch { start: number; end: number; text: string }
  const patches: RowPatch[] = [];
  /** 同一插入点上的多处改动（例如"改写某个 <c>"与"在它前面插一个新 <c>"）要合并 */
  const push = (start: number, end: number, text: string): void => {
    const last = patches[patches.length - 1];
    if (last && last.start === start) {
      last.end = Math.max(last.end, end);
      last.text += text;
      return;
    }
    patches.push({ start, end, text });
  };
  /** 行内每个格子的列号与它在 inner 中的区间 */
  const cells: Array<{ col: number; start: number; end: number }> = [];

  let implicitCol = 0;
  for (const cellEl of childElements(xml, rowEl)) {
    if (!nameMatches(cellEl.name, 'c')) continue;
    const rawRef = cellEl.attrs['r'];
    const parsed = rawRef ? parseCellRef(rawRef) : undefined;
    const row = parsed ? parsed.row : rowIndex;
    const col = parsed ? parsed.col : implicitCol;
    implicitCol = col + 1;
    const start = cellEl.openStart - rowEl.openEnd;
    const end = cellEl.closeEnd - rowEl.openEnd;
    cells.push({ col, start, end });
    const target = planned.get(`${colLetters(col)}${row + 1}`);
    if (!target) continue;
    const rewritten = rewriteCell(xml.slice(cellEl.openStart, cellEl.closeEnd), cellEl, target, xml, rowStyle);
    push(start, end, rewritten ?? ''); // rewritten === null = 整格删除（替换成空串）
  }

  // 新增的格子：按列号升序，插到行内第一个列号更大的 <c> 之前；都不大于则追加到行尾
  const additions = collectRowAdditions(planned, rowIndex)
    .filter((addition) => !cells.some((cell) => cell.col === addition.col));
  for (const addition of additions) {
    const next = cells.find((cell) => cell.col > addition.col);
    const at = next ? next.start : inner.length;
    push(at, at, serializeRefCell(addition.ref, addition.content, rowStyle));
  }

  if (patches.length === 0) return xml.slice(rowEl.openStart, rowEl.closeEnd);
  touchedRows.add(rowIndex);

  patches.sort((a, b) => a.start - b.start);
  let body = '';
  let cursor = 0;
  for (const patch of patches) {
    body += inner.slice(cursor, patch.start) + patch.text;
    cursor = Math.max(cursor, patch.end);
  }
  body += inner.slice(cursor);

  // 行内容已被清空：只有"位置/选区提示"这类附属属性（r、spans）时整行删掉，
  // 否则保留空行（行高、隐藏等行级信息必须留着）
  if (body.trim().length === 0 && !hasRowLevelInfo(rowEl)) return '';
  if (rowEl.selfClosing) {
    // 原本 <row .../>：要有新格子才展开（展开时要去掉结尾的 `/`）
    return body.length === 0
      ? xml.slice(rowEl.openStart, rowEl.closeEnd)
      : `${openTag(xml, rowEl).replace(/\/\s*>$/, '>')}${body}</row>`;
  }
  return `${openTag(xml, rowEl)}${body}${closeTagOf(xml, rowEl)}`;
}

function collectRowAdditions(
  planned: Map<string, CellContent>,
  rowIndex: number,
): Array<{ ref: string; col: number; content: CellContent }> {
  const out: Array<{ ref: string; col: number; content: CellContent }> = [];
  for (const [ref, content] of planned) {
    const coord = parseCellRef(ref);
    if (!coord || coord.row !== rowIndex) continue;
    out.push({ ref, col: coord.col, content });
  }
  return out.sort((a, b) => a.col - b.col);
}

/* -------------------------------------------------------------------------- */
/* 单元格原文改写                                                              */
/* -------------------------------------------------------------------------- */

/**
 * 用新内容改写一个**已存在**单元格的原文。
 *
 * - 保留 `r`、`s`（样式）与 `cm`/`vm`（元数据索引，与值同生共死）；
 * - `t` 按新值类型重写：字符串 -> `t="s"`（公式字符串结果 -> `t="str"`）、布尔 -> `t="b"`、
 *   数字 -> 去掉 `t`；清空内容时保留原 `t` 以免影响同格其它属性语义（此时无 `<v>`）；
 * - `<f>` / `<v>` 换成新内容，`<is>`、`<extLst>` 等其它子元素原样保留；
 * - 返回 `null` 表示整格删除（清空后既无属性也无内容）。
 */
function rewriteCell(
  original: string,
  el: ElementRange,
  content: CellContent,
  xml: string,
  fallbackStyle: number | undefined,
): string | null {
  const openEnd = openTagEnd(original, 0);
  const open = original.slice(0, openEnd);
  // 属性正文 = 标签名之后到 `>`（或 `/>`）之前
  const bodyStart = 1 + el.name.length;
  const bodyEnd = openEnd - (el.selfClosing ? 2 : 1);
  const attrs = parseAttrText(open.slice(bodyStart, Math.max(bodyStart, bodyEnd)));
  fillMissingStyle(attrs, fallbackStyle);

  if (content.type === 's') setAttr(attrs, 't', 's');
  else if (content.type === 'str') setAttr(attrs, 't', 'str');
  else if (content.type === 'b') setAttr(attrs, 't', 'b');
  else setAttr(attrs, 't', undefined); // 数字 / 清空：t 不再有意义，去掉

  const kept: string[] = [];
  if (!el.selfClosing) {
    for (const child of childElements(xml, el)) {
      // t 不是 s 却还留着 <v> 会让 Excel 报修复；公式缓存同理，一律丢掉重写
      const ln = localName(child.name);
      if (ln === 'v' || ln === 'f') continue;
      kept.push(xml.slice(child.openStart, child.closeEnd));
    }
  }
  if (content.formula !== undefined) kept.push(`<f>${escapeXml(content.formula)}</f>`);
  const valuePart = renderValue(content);
  if (valuePart !== undefined) kept.push(valuePart);

  if (kept.length === 0) {
    // 清空后剩下的只是空壳：只有 r（或什么都没有）就整格删掉，留着 s 之类的才写成自闭合
    const meaningful = attrs.entries.filter(([name]) => name !== 'r');
    if (meaningful.length === 0) return null;
    return `<c${serializeAttrs(attrs)}/>`;
  }
  return `<c${serializeAttrs(attrs)}>${kept.join('')}</c>`;
}

/** 目标行原本不存在时整格新造一个 `<c r=".." .../>`；`style` 用于补齐同行样式 */
function serializeRefCell(ref: string, content: CellContent, style?: number): string {
  const attrs: AttrCollector = { entries: [['r', ref]] };
  fillMissingStyle(attrs, style);
  if (content.type === 's') setAttr(attrs, 't', 's');
  else if (content.type === 'str') setAttr(attrs, 't', 'str');
  else if (content.type === 'b') setAttr(attrs, 't', 'b');
  const kept: string[] = [];
  if (content.formula !== undefined) kept.push(`<f>${escapeXml(content.formula)}</f>`);
  const valuePart = renderValue(content);
  if (valuePart !== undefined) kept.push(valuePart);
  return `<c${serializeAttrs(attrs)}${kept.length === 0 ? '/>' : `>${kept.join('')}</c>`}`;
}

/**
 * 补齐样式属性。原始单元格本来没有 `s`、但同一行其它格子有样式时，
 * 沿用它（行 3 的 C/D/E/F 都该是同一档样式），否则新格子会变成"无样式的异类"。
 * 编辑模型里没有"新单元格样式"这一项，所以这是唯一能保持视觉一致的办法。
 */
function fillMissingStyle(attrs: AttrCollector, fallback: number | undefined): void {
  if (fallback === undefined || fallback < 0) return;
  if (attrs.entries.some(([name]) => name === 's')) return;
  attrs.entries.push(['s', String(fallback)]);
}

/** 每张表里：行号 -> 该行第一个带样式的单元格的 s（用于新格子补齐样式） */
function collectRowStyles(sheets: readonly ExportSourceSheet[]): Map<number, number> {
  const out = new Map<number, number>();
  for (const sheet of sheets) {
    // 瘦身来源：直接读预算好的映射（与下面现算的结果完全一致）
    if (sheet.rowStyles) {
      for (const [row, style] of sheet.rowStyles) {
        if (!out.has(row)) out.set(row, style);
      }
      continue;
    }
    for (const cell of sheet.cells ?? []) {
      if (cell.styleIndex === undefined) continue;
      if (!out.has(cell.row)) out.set(cell.row, cell.styleIndex);
    }
  }
  return out;
}

/** `<v>` 片段；`undefined` 表示不写 `<v>`（清空/无缓存值） */
function renderValue(content: CellContent): string | undefined {
  const value = content.value;
  if (value === null) return undefined;
  if (typeof value === 'boolean') return `<v>${value ? '1' : '0'}</v>`;
  if (typeof value === 'number') return `<v>${String(value)}</v>`;
  if (content.type === 's') {
    if (content.sstIndex === undefined) {
      throw new Error('exportXlsx：内部错误，共享字符串索引缺失');
    }
    return `<v>${content.sstIndex}</v>`;
  }
  return `<v>${escapeXml(value)}</v>`;
}

/* -------------------------------------------------------------------------- */
/* 属性文本处理（只动必要的键，其它属性原样保留）                                */
/* -------------------------------------------------------------------------- */

interface AttrCollector { entries: Array<[string, string]> }

/**
 * 解析开始标签里的属性（`t` 在原文里也可能写作 `type`，都认）。
 *
 * `bound` 用于限定解析范围：自闭合标签要排除结尾的 `/`，否则会被当成一个"无值属性"
 * 而写出 `c="true"` 这种垃圾属性。
 */
function parseAttrText(text: string, bound: number = text.length): AttrCollector {
  const entries: Array<[string, string]> = [];
  const end = Math.max(0, Math.min(bound, text.length));
  let i = 0;
  while (i < end) {
    while (i < end && isWs(text.charCodeAt(i))) i++;
    if (i >= end) break;
    if (text.charCodeAt(i) === 47 /* / */) break;
    const nameStart = i;
    while (i < end && !isNameStop(text.charCodeAt(i))) i++;
    const name = text.slice(nameStart, i);
    if (name.length === 0) break;
    while (i < end && isWs(text.charCodeAt(i))) i++;
    if (text.charCodeAt(i) !== 61 /* = */) {
      entries.push([name, 'true']);
      continue;
    }
    i++;
    while (i < end && isWs(text.charCodeAt(i))) i++;
    const quote = text.charCodeAt(i);
    if (quote === 34 || quote === 39) {
      i++;
      const start = i;
      while (i < end && text.charCodeAt(i) !== quote) i++;
      entries.push([name, text.slice(start, i)]);
      i++;
      continue;
    }
    const start = i;
    while (i < end && !isWs(text.charCodeAt(i))) i++;
    entries.push([name, text.slice(start, i)]);
  }
  return { entries };
}

/** 设置/删除 `t`（或 `type`）属性：已有就改原位置，避免属性顺序大搬家 */
function setAttr(attrs: AttrCollector, key: 't', value: string | undefined): void {
  const index = attrs.entries.findIndex(([name]) => name === 't' || name === 'type');
  if (value === undefined) {
    if (index >= 0) attrs.entries.splice(index, 1);
    return;
  }
  if (index >= 0) attrs.entries[index] = [attrs.entries[index][0], value];
  else attrs.entries.push([key, value]);
}

function serializeAttrs(attrs: AttrCollector): string {
  let out = '';
  for (const [name, value] of attrs.entries) out += ` ${name}="${escapeAttr(value)}"`;
  return out;
}

/* -------------------------------------------------------------------------- */
/* sheetData 重建                                                              */
/* -------------------------------------------------------------------------- */

/**
 * 把 `sheetData` 的内容拼回去：保留原有 `<row>` 的区间（只替换被修补的那些），
 * 并把新建的行按行号升序插到正确位置。
 *
 * `rowResults` 里的区间是整篇 XML 的坐标，`base` 是 `inner` 在整篇 XML 中的起点。
 */
function rebuildSheetData(
  inner: string,
  newRows: Array<{ rowIndex: number; cells: Array<{ col: number; text: string }> }>,
  rowResults: Array<{ rowEl: ElementRange; text: string }>,
  base: number,
): string {
  if (newRows.length === 0 && rowResults.length === 0) return inner;

  // 行号 -> 该行在 inner 中的区间
  const spans: Array<{ index: number; start: number; end: number; text: string }> = [];
  let implicit = 0;
  for (const { rowEl, text } of rowResults) {
    const r = attrInt(rowEl.attrs, 'r');
    const index = r !== undefined && r > 0 ? r - 1 : implicit;
    implicit = index + 1;
    spans.push({
      index,
      start: rowEl.openStart - base,
      end: rowEl.closeEnd - base,
      text,
    });
  }

  let out = '';
  let cursor = 0;
  let lastIndex = -1;
  let pending = 0;
  /** 把行号小于 `before` 的新行全部吐出 */
  const flushBefore = (before: number): void => {
    while (pending < newRows.length && newRows[pending].rowIndex < before) {
      out += serializeRow(newRows[pending]);
      pending++;
    }
  };
  for (const span of spans) {
    out += inner.slice(cursor, span.start);
    cursor = span.end;
    if (span.index >= lastIndex) {
      // 正常情况：新行插在"第一个行号更大的原有行"之前
      flushBefore(span.index);
    } else {
      // 行号倒退（异常数据）：插到这一行之前，至少不破坏已有的单调顺序
      const next = newRows[pending];
      if (next) {
        out += serializeRow(next);
        pending++;
      }
    }
    out += span.text;
    if (span.index > lastIndex) lastIndex = span.index;
  }
  out += inner.slice(cursor);
  while (pending < newRows.length) {
    out += serializeRow(newRows[pending]);
    pending++;
  }
  return out;
}

function serializeRow(row: { rowIndex: number; cells: Array<{ col: number; text: string }> }): string {
  const cells = [...row.cells].sort((a, b) => a.col - b.col).map((cell) => cell.text).join('');
  return `<row r="${row.rowIndex + 1}">${cells}</row>`;
}

function rowNumberOf(rowEl: ElementRange): number | undefined {
  const r = attrInt(rowEl.attrs, 'r');
  return r !== undefined && r > 0 ? r - 1 : undefined;
}

/** 行上是否有"必须保留"的信息：行高 / 隐藏 / 自定义格式 / 大纲级别等（`r`、`spans` 只是附属提示） */
const ROW_AUX_ATTRS: ReadonlySet<string> = new Set(['r', 'spans', 'x14ac:dyDescent', 'dyDescent']);

function hasRowLevelInfo(rowEl: ElementRange): boolean {
  return Object.keys(rowEl.attrs).some((name) => !ROW_AUX_ATTRS.has(name));
}

/* -------------------------------------------------------------------------- */
/* dimension                                                                   */
/* -------------------------------------------------------------------------- */

/** 只有编辑目标越出原 `<dimension>` 时才扩展（策略见文件头） */
function updateDimension(xml: string, planned: Map<string, CellContent>): string {
  if (planned.size === 0) return xml;
  const dim = findFirstElement(xml, 'dimension');
  if (!dim) return xml;
  const ref = dim.attrs['ref'];
  const dot = ref ? ref.indexOf(':') : -1;
  if (!ref || dot < 0) return xml;
  const start = parseCellRef(ref.slice(0, dot));
  const end = parseCellRef(ref.slice(dot + 1));
  if (!start || !end) return xml;

  let minRow = Math.min(start.row, end.row);
  let minCol = Math.min(start.col, end.col);
  let maxRow = Math.max(start.row, end.row);
  let maxCol = Math.max(start.col, end.col);

  let outside = false;
  const coords: Array<{ row: number; col: number }> = [];
  for (const key of planned.keys()) {
    const coord = parseCellRef(key);
    if (!coord) continue;
    coords.push(coord);
    if (coord.row < minRow || coord.row > maxRow || coord.col < minCol || coord.col > maxCol) outside = true;
  }
  if (!outside) return xml;

  for (const coord of coords) {
    minRow = Math.min(minRow, coord.row);
    minCol = Math.min(minCol, coord.col);
    maxRow = Math.max(maxRow, coord.row);
    maxCol = Math.max(maxCol, coord.col);
  }
  const next = `${colLetters(minCol)}${minRow + 1}:${colLetters(maxCol)}${maxRow + 1}`;
  if (next === ref) return xml;

  // 只替换 <dimension> 开始标签里的 ref 属性值，其余字节保持原样
  const openEnd = openTagEnd(xml, dim.openStart);
  const tag = xml.slice(dim.openStart, openEnd);
  const replacedTag = tag.replace(/ref\s*=\s*"[^"]*"/, `ref="${escapeAttr(next)}"`);
  if (replacedTag === tag) return xml;
  return xml.slice(0, dim.openStart) + replacedTag + xml.slice(openEnd);
}

/* -------------------------------------------------------------------------- */
/* workbook.xml: <calcPr>                                                      */
/* -------------------------------------------------------------------------- */

/** 只改 `<calcPr>`：有就补/改 fullCalcOnLoad，没有就插在 `<sheets>` 之后（schema 允许的位置） */
function patchCalcPr(xml: string): string {
  const calcPr = findFirstElement(xml, 'calcPr');
  if (calcPr) {
    const openEnd = openTagEnd(xml, calcPr.openStart);
    const tag = xml.slice(calcPr.openStart, openEnd);
    let next: string;
    if (/\bfullCalcOnLoad\s*=/.test(tag)) {
      next = tag.replace(/fullCalcOnLoad\s*=\s*"[^"]*"/, 'fullCalcOnLoad="1"');
    } else {
      const trimmed = tag.replace(/\s*\/?>$/, '');
      next = `${trimmed} fullCalcOnLoad="1"${calcPr.selfClosing ? '/>' : '>'}`;
    }
    if (next === tag) return xml;
    return xml.slice(0, calcPr.openStart) + next + xml.slice(openEnd);
  }

  const anchor = findFirstElement(xml, 'sheets');
  if (anchor) return `${xml.slice(0, anchor.closeEnd)}<calcPr fullCalcOnLoad="1"/>${xml.slice(anchor.closeEnd)}`;
  const workbook = findFirstElement(xml, 'workbook');
  if (workbook && !workbook.selfClosing) {
    return `${xml.slice(0, workbook.openEnd)}<calcPr fullCalcOnLoad="1"/>${xml.slice(workbook.openEnd)}`;
  }
  return xml;
}

/* -------------------------------------------------------------------------- */
/* sharedStrings                                                               */
/* -------------------------------------------------------------------------- */

interface SharedStringEditor {
  /** 原表文本 + 本次追加的文本（索引即单元格 `<v>` 里的值） */
  readonly strings: string[];
  changed: boolean;
  entryName?: string;
  bytes?: Uint8Array;
  /** 文本 -> 索引：命中已有条目就复用，否则追加（返回新索引） */
  intern(text: string): number;
}

function createSharedStringEditor(entries: ZipEntries, needed: boolean): SharedStringEditor {
  const entryName = findSharedStringsEntry(entries);
  const bytes = entryName ? findEntry(entries, entryName) : undefined;
  const source = bytes ? strFromU8(bytes) : undefined;
  const nodes = source ? parseSharedStringNodes(source) : [];
  /** 当前（可能已经被追加过的）表 XML；每次追加后都要更新，否则下一次追加会覆盖上一次 */
  let current = source;

  const editor: SharedStringEditor = {
    strings: nodes.map((node) => node.text),
    changed: false,
    ...(entryName ? { entryName } : {}),
    intern(text: string): number {
      const known = editor.strings.indexOf(text);
      if (known >= 0) return known;
      const index = editor.strings.length;
      if (!needed) {
        // 没有实际写入需求时只登记到内存表，不产生任何输出改动
        editor.strings.push(text);
        return index;
      }
      editor.strings.push(text);
      // `<si>` 条数以追加后的表长为权威值（count / uniqueCount 都要用它）
      const total = editor.strings.length;
      current = current === undefined
        ? createSharedStrings(nodes, text, total)
        : appendSharedString(current, total, text);
      if (source === undefined) editor.entryName = editor.entryName ?? DEFAULT_SHARED_STRINGS_PATH;
      editor.bytes = strToU8(current);
      editor.changed = true;
      return index;
    },
  };
  return editor;
}

function findSharedStringsEntry(entries: ZipEntries): string | undefined {
  if (findEntry(entries, DEFAULT_SHARED_STRINGS_PATH)) return DEFAULT_SHARED_STRINGS_PATH;
  return Object.keys(entries).find((name) => {
    const lower = name.toLowerCase();
    return lower.endsWith('sharedstrings.xml') && !lower.includes('_rels');
  });
}

interface SharedStringNode { text: string; preserved: boolean }

/** 只读地取出所有 `<si>` 的纯文本（富文本取各 run 的 `<t>` 拼接，与解析器一致） */
function parseSharedStringNodes(xml: string): SharedStringNode[] {
  const out: SharedStringNode[] = [];
  for (const si of directElements(xml, 'si')) {
    const kids = childElements(xml, si);
    const runs = kids.filter((kid) => nameMatches(kid.name, 'r'));
    if (runs.length > 0) {
      let text = '';
      let preserved = false;
      for (const run of runs) {
        const t = childElements(xml, run).find((kid) => nameMatches(kid.name, 't'));
        if (!t) continue;
        text += innerTextOf(xml, t);
        if (/\bxml:space\s*=\s*"preserve"/.test(xml.slice(t.openStart, openTagEnd(xml, t.openStart)))) {
          preserved = true;
        }
      }
      out.push({ text, preserved });
      continue;
    }
    const texts = kids.filter((kid) => nameMatches(kid.name, 't'));
    let text = '';
    let preserved = false;
    for (const t of texts) {
      text += innerTextOf(xml, t);
      if (/\bxml:space\s*=\s*"preserve"/.test(xml.slice(t.openStart, openTagEnd(xml, t.openStart)))) {
        preserved = true;
      }
    }
    out.push({ text, preserved });
  }
  return out;
}

/**
 * 追加一个 `<si><t>…</t></si>`，并同步 `<sst>` 上的 count / uniqueCount（存在才同步）。
 * `total` 是追加后的 `<si>` 总数。
 */
function appendSharedString(xml: string, total: number, text: string): string {
  const root = findFirstElement(xml, 'sst');
  if (!root) return xml;
  const node = sharedStringNode(text);
  let out: string;
  if (root.selfClosing) {
    const tag = xml.slice(root.openStart, openTagEnd(xml, root.openStart)).replace(/\/\s*>$/, '>');
    out = `${xml.slice(0, root.openStart)}${tag}${node}</sst>${xml.slice(root.closeEnd)}`;
  } else {
    const closeAt = xml.lastIndexOf('</', innerEnd(xml, root));
    const at = closeAt >= root.openEnd ? closeAt : innerEnd(xml, root);
    out = `${xml.slice(0, at)}${node}${xml.slice(at)}`;
  }
  return bumpSstCounts(out, total);
}

/** 原文件没有 sharedStrings.xml 时才走这里：新建一个只含本次字符串的最小表 */
function createSharedStrings(existing: SharedStringNode[], text: string, total: number): string {
  const nodes = [...existing.map((node) => sharedStringNode(node.text, node.preserved)), sharedStringNode(text)];
  return `${XML_DECL}<sst xmlns="${SST_NAMESPACE}" count="${total}" uniqueCount="${total}">`
    + `${nodes.join('')}</sst>`;
}

function sharedStringNode(text: string, preserved = needsPreserve(text)): string {
  return `<si><t${preserved ? ' xml:space="preserve"' : ''}>${escapeXml(text)}</t></si>`;
}

function bumpSstCounts(xml: string, total: number): string {
  const root = findFirstElement(xml, 'sst');
  if (!root) return xml;
  const openEnd = openTagEnd(xml, root.openStart);
  let tag = xml.slice(root.openStart, openEnd);
  if (/\bcount\s*=/.test(tag)) tag = tag.replace(/count\s*=\s*"[^"]*"/, `count="${total}"`);
  if (/\buniqueCount\s*=/.test(tag)) tag = tag.replace(/uniqueCount\s*=\s*"[^"]*"/, `uniqueCount="${total}"`);
  return xml.slice(0, root.openStart) + tag + xml.slice(openEnd);
}

/* -------------------------------------------------------------------------- */
/* 局部工具                                                                    */
/* -------------------------------------------------------------------------- */

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/** 字节级比较：内容没变就不登记替换，原条目字节原样写回 */
function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function colLetters(col: number): string {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    const rest = (n - 1) % 26;
    out = String.fromCharCode(65 + rest) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function baseName(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash < 0 ? path : path.slice(slash + 1);
}

function normalizeEntryPath(path: string): string {
  let p = path.replace(/\\/g, '/');
  if (p.startsWith('/')) p = p.slice(1);
  while (p.startsWith('./')) p = p.slice(2);
  return p;
}

/**
 * 元素内容结束位置（自闭合元素等于 `openEnd`，否则等于结束标签的起点）。
 *
 * 不能用 `lastIndexOf('</', el.closeEnd)`：那是"从 closeEnd 往前找最后一个 `</`"，
 * 当元素内容以子元素结尾时（`<row r="1"><c .../></row>`）会命中**父元素的**结束标签，
 * 把 `</row>` 当成内容切出来。这里从 closeEnd 往前找匹配的 `<`，再校验标签名。
 */
function innerEnd(xml: string, el: ElementRange): number {
  if (el.selfClosing || el.closeEnd <= el.openEnd) return el.openEnd;
  let at = xml.lastIndexOf('<', el.closeEnd - 1);
  while (at >= el.openEnd) {
    if (xml.charCodeAt(at + 1) === 47 /* / */ && tagNameAt(xml, at + 2) === el.name) return at;
    const next = xml.lastIndexOf('<', at - 1);
    if (next >= at) break;
    at = next;
  }
  return el.closeEnd;
}

/** 从 `from` 开始读一个标签名（到空白 / `/` / `>` 为止） */
function tagNameAt(xml: string, from: number): string {
  let i = from;
  while (i < xml.length && !isNameStop(xml.charCodeAt(i))) i++;
  return xml.slice(from, i);
}

function openTag(xml: string, el: ElementRange): string {
  return xml.slice(el.openStart, openTagEnd(xml, el.openStart));
}

function closeTagOf(xml: string, el: ElementRange): string {
  return el.selfClosing ? '' : xml.slice(innerEnd(xml, el), el.closeEnd);
}

/** 从 `<` 扫描到开始标签的 `>`（跳过引号内的 `>`） */
function openTagEnd(text: string, from: number): number {
  let i = from + 1;
  let quote = 0;
  while (i < text.length) {
    const c = text.charCodeAt(i);
    if (quote !== 0) {
      if (c === quote) quote = 0;
    } else if (c === 34 || c === 39) {
      quote = c;
    } else if (c === 62 /* > */) {
      return i + 1;
    }
    i++;
  }
  return text.length;
}

/**
 * 逐次取出同名的顶层元素（`<si>` 这种可能有很多个的）。
 *
 * 注意：这里**不能**用解析层的 `findFirstElement(xml, 'si')`。它的名字匹配是
 * 「按 localName 前缀无关」的，`<sst>` / `<si>` 在标签名切分上会被混为一谈
 * （`si` 是 `sst` 的前缀），于是只会拿到根元素一个节点，索引就全错了。
 * 所以这里自己按标签名**精确匹配**扫描，并且只取最外层。
 */
function directElements(xml: string, name: string): ElementRange[] {
  const out: ElementRange[] = [];
  for (const range of collectElements(xml, name)) {
    const last = out[out.length - 1];
    if (last && range.openStart < last.closeEnd) continue; // 嵌套的同名元素忽略
    out.push(range);
  }
  return out;
}

/** 扫描全部标签，取出标签名**恰好等于** `name` 的非自闭合元素 */
function collectElements(xml: string, name: string): ElementRange[] {
  const out: ElementRange[] = [];
  const open = new RegExp(`<${name}(?=[\\s/>])`, 'g');
  const closeTag = `</${name}>`;
  for (let match = open.exec(xml); match; match = open.exec(xml)) {
    const end = xml.indexOf(closeTag, match.index + match[0].length);
    if (end < 0) break;
    out.push(makeRange(xml, name, match.index, end + closeTag.length));
  }
  return out;
}

/** 用解析层同一套词法重建一个元素区间 */
function makeRange(xml: string, name: string, openStart: number, closeEnd: number): ElementRange {
  const openEnd = openTagEnd(xml, openStart);
  const selfClosing = xml.charCodeAt(openEnd - 2) === 47;
  const attrs: Record<string, string> = {};
  const bodyStart = openStart + 1 + name.length;
  const bodyEnd = openEnd - (selfClosing ? 2 : 1);
  for (const [key, value] of parseAttrText(xml.slice(bodyStart, Math.max(bodyStart, bodyEnd))).entries) {
    attrs[key] = value;
  }
  return { name, openStart, openEnd, closeEnd, selfClosing, attrs, start: openStart };
}

/** 子元素内的纯文本（解码 XML 实体；富文本场景只用于叶子 `<t>`） */
function innerTextOf(xml: string, el: ElementRange): string {
  if (el.selfClosing) return '';
  const at = xml.lastIndexOf(`</${el.name}>`, el.closeEnd);
  const from = at >= el.openEnd ? at : el.closeEnd;
  return decodeXmlEntities(xml.slice(el.openEnd, from));
}

function decodeXmlEntities(text: string): string {
  if (text.indexOf('&') < 0) return text;
  return text.replace(/&(?:#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match) => {
    switch (match) {
      case '&amp;': return '&';
      case '&lt;': return '<';
      case '&gt;': return '>';
      case '&quot;': return '"';
      case '&apos;': return "'";
      default: break;
    }
    if (match.startsWith('&#x') || match.startsWith('&#X')) {
      return safeFromCodePoint(Number.parseInt(match.slice(3, -1), 16));
    }
    if (match.startsWith('&#')) return safeFromCodePoint(Number.parseInt(match.slice(2, -1), 10));
    return match; // 未知实体原样保留
  });
}

function safeFromCodePoint(cp: number): string {
  if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff) return '';
  try {
    return String.fromCodePoint(cp);
  } catch {
    return '';
  }
}

/**
 * 公式正文清洗：
 * - 去掉可能混进来的前导 `=`（契约要求不含它，这里容错）；
 * - 调用方若把 `<` / `>` 写成了 XML 转义（`&lt;` / `&gt;`），这里先还原成字面量再输出，
 *   输出阶段统一由 `escapeXml` 负责转义，避免出现 `&amp;gt;` 这种双重转义。
 */
function stripFormula(formula: string): string {
  const trimmed = formula.trim();
  return decodeXmlEntities(trimmed.startsWith('=') ? trimmed.slice(1) : trimmed);
}

/** XML 文本转义：& < > " '，并剔除 XML 1.0 非法控制字符 / 落单代理项 */
function escapeXml(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += text[i] + text[i + 1];
        i++;
        continue;
      }
      out += '\ufffd';
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) {
      out += '\ufffd';
      continue;
    }
    if (code === 0x09 || code === 0x0a || code === 0x0d) {
      out += text[i];
      continue;
    }
    if (code < 0x20 || code === 0xfffe || code === 0xffff) continue;
    switch (code) {
      case 38: out += '&amp;'; break;
      case 60: out += '&lt;'; break;
      case 62: out += '&gt;'; break;
      case 34: out += '&quot;'; break;
      case 39: out += '&apos;'; break;
      default: out += text[i];
    }
  }
  return out;
}

/** 属性值转义（换行/制表符按规范写成字符引用） */
function escapeAttr(text: string): string {
  return escapeXml(text).replace(/[\t\n\r]/g, (ch) => {
    if (ch === '\t') return '&#9;';
    if (ch === '\n') return '&#10;';
    return '&#13;';
  });
}

function needsPreserve(text: string): boolean {
  if (text.length === 0) return false;
  return isWs(text.charCodeAt(0)) || isWs(text.charCodeAt(text.length - 1));
}

function isWs(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13;
}

function isNameStop(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13
    || code === 47 || code === 62 || code === 61;
}
