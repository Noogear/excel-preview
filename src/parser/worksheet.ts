/**
 * 单工作表解析：`xl/worksheets/sheetN.xml`。
 *
 * 覆盖范围：
 * - `sheetData` 的行/列引用（A1 / AB12）、行高、隐藏行、自定义行高
 * - 单元格值类型：共享字符串 / 公式字符串 / 内联字符串 / 布尔 / 错误 / 数字
 * - 公式 `<f>` 与缓存值 `<v>`
 * - `mergeCells`、冻结窗格、`showGridLines`、`cols` 列宽、`dimension`、`sheetFormatPr`
 * - P1：`<hyperlinks>`（ref / r:id / location / display / tooltip）、`<tableParts>`（收集 r:id）、
 *   `<drawing r:id>`；`sqref` 展开成 A1 区域数组（`expandSqref`，条件格式与数据验证共用）
 * - 打印设置（`pageSetup`/`pageMargins`/`headerFooter`/`printOptions`/分页符）**已解析**进
 *   `sheet.print`，只影响打印 → 进 `report.preserved`，不再列为"未支持"
 * - 仍**不解析**但按种类记账的功能进 `unsupported`（迷你图 / 排序状态 / VML 绘图 / 保护 / …）：
 *   条件格式、数据验证、超链接、表格、批注、浮动图片自 P1 起已解析，不再记账。
 *
 * 本文件只读**工作表自身**的 XML；批注/表格/绘图等外部部件由编排层（index.ts）通过
 * `xl/worksheets/_rels/sheetN.xml.rels` 定位、解析后经 `ParseWorksheetInput.parts` 注入，
 * 因此这里不依赖 zip，也不依赖任何 IO。
 */
import type {
  ParsedCell, ParsedCellValue, ParsedColInfo, ParsedHyperlink, ParsedImage, ParsedMerge,
  ParsedNote, ParsedPrintSettings, ParsedRowInfo, ParsedSheet, ParsedTable,
} from './types';
import {
  attrBool, attrInt, attrNumber, childElements, decodeEntities, earliestNameAt, elementText,
  findCloseTag, findFirstElement, findElements, localName, nameMatches, parseAttrs, stripTags,
  visitElements, type ElementRange, type XmlAttributes,
} from './xml';

/** 解析层的统一降级回调（不抛异常，问题都变成可读文本） */
export type Warn = (msg: string) => void;

// 热路径内联用的字符码（避免每格都做函数调用）
const SLASH = 47;
const GT = 62;
const SPACE = 32;
const TAB = 9;
const LF = 10;
const CR = 13;
const LOWER_V = 118;
const LOWER_F = 102;

/* -------------------------------------------------------------------------- */
/* 引用解析                                                                    */
/* -------------------------------------------------------------------------- */

/** "AB12" -> { row: 11, col: 27 }（0-based；行 1-based -> 0-based，列字母 -> 0-based） */
export function parseCellRef(ref: string): { row: number; col: number } | undefined {
  const s = ref.trim();
  let i = 0;
  let col = 0;
  let letters = 0;
  while (i < s.length) {
    const c = s.charCodeAt(i);
    let v: number;
    if (c >= 65 && c <= 90) v = c - 64;       // A-Z
    else if (c >= 97 && c <= 122) v = c - 96; // a-z（OOXML 要求大写，容错小写）
    else break;
    col = col * 26 + v;
    letters++;
    i++;
    if (letters > 3) return undefined; // 超过 XFD 已不合法
  }
  if (letters === 0) return undefined;
  let digits = 0;
  let row = 0;
  while (i < s.length) {
    const c = s.charCodeAt(i);
    if (c < 48 || c > 57) break;
    row = row * 10 + (c - 48);
    digits++;
    i++;
    if (digits > 7) return undefined;
  }
  // 必须整串消费完（`A1:B2`、`A1 ` 之外的脏数据一律判为不可解析）
  if (digits === 0 || i !== s.length || row < 1) return undefined;
  return { row: row - 1, col: col - 1 };
}

/** "A1:D10" / "A1" -> 合并区间（0-based，包含端点） */
export function parseRangeRef(ref: string): ParsedMerge | undefined {
  const parts = ref.split(':');
  const start = parseCellRef(parts[0]);
  if (!start) return undefined;
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start;
  if (!end) return undefined;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

/* -------------------------------------------------------------------------- */
/* sqref：'A1:B2 D4:E5' -> ['A1:B2', 'D4:E5']                                   */
/* -------------------------------------------------------------------------- */

/**
 * 按空白切分 sqref（`<conditionalFormatting sqref>`、`<dataValidation sqref>`、
 * `<mergeCell>` 之外的引用型属性都用空白分隔多个区域）。不做合法性校验，也不建中间字符串。
 */
export function splitSqref(sqref: string): string[] {
  const out: string[] = [];
  const len = sqref.length;
  let i = 0;
  while (i < len) {
    const c = sqref.charCodeAt(i);
    if (c === 32 || c === 9 || c === 10 || c === 13) { i++; continue; }
    const start = i;
    while (i < len) {
      const c2 = sqref.charCodeAt(i);
      if (c2 === 32 || c2 === 9 || c2 === 10 || c2 === 13) break;
      i++;
    }
    out.push(sqref.slice(start, i));
  }
  return out;
}

/** 单个 A1 记号是否合法：`A1` / `A1:B2` / 整列 `A:A` / 整行 `1:1`（可带 `$`） */
export function isA1RangeToken(token: string): boolean {
  if (token.length === 0) return false;
  const colon = token.indexOf(':');
  if (colon < 0) return endpointKind(token) !== 0;
  const head = token.slice(0, colon);
  const tail = token.slice(colon + 1);
  if (tail.indexOf(':') >= 0) return false;
  const headKind = endpointKind(head);
  return headKind !== 0 && headKind === endpointKind(tail);
}

/** 端点种类：1 = 单元格，2 = 整列，3 = 整行，0 = 非法（`A:B`、`1:2` 合法，`A:2` 非法） */
function endpointKind(part: string): 0 | 1 | 2 | 3 {
  const p = part.replace(/\$/g, '');
  if (p.length === 0) return 0;
  if (/^[A-Za-z]{1,3}$/.test(p)) return 2;
  if (/^[0-9]{1,7}$/.test(p)) return 3;
  return parseCellRef(p) !== undefined ? 1 : 0;
}

/**
 * sqref -> A1 区域数组：`'A1:B2 D4:E5'` -> `['A1:B2', 'D4:E5']`。
 * 非法片段被丢弃；给了 `warn` 时逐条记录（`owner` 用于指明来源，如 `dataValidation`）。
 */
export function expandSqref(sqref: string, warn?: Warn, owner?: string): string[] {
  const out: string[] = [];
  for (const token of splitSqref(sqref)) {
    if (isA1RangeToken(token)) out.push(token);
    else if (warn) warn(`${owner ?? 'sqref'} 的引用片段 "${token}" 不是合法的 A1 区域，已跳过`);
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* sharedStrings                                                               */
/* -------------------------------------------------------------------------- */

/**
 * `xl/sharedStrings.xml` -> 纯文本数组（索引 == `t="s"` 单元格的 `<v>` 值）。
 * 富文本（多个 `<r>` run）在 P0 里直接拼接为纯文本，格式化信息丢弃。
 */
export function parseSharedStrings(xml: string, warn: (msg: string) => void): string[] {
  const out: string[] = [];
  let richTextCount = 0;
  for (const si of findElements(xml, 'si')) {
    const runs = childElements(xml, si).filter((c) => nameMatches(c.name, 'r'));
    if (runs.length > 0) {
      richTextCount++;
      out.push(concatRuns(xml, runs));
      continue;
    }
    const texts = childElements(xml, si).filter((c) => nameMatches(c.name, 't'));
    out.push(texts.length > 0 ? texts.map((t) => elementText(xml, t)).join('') : '');
  }
  if (richTextCount > 0) {
    warn(`sharedStrings 中有 ${richTextCount} 条富文本（多 run）已降级为纯文本，run 级格式丢失`);
  }
  return out;
}

/** 把若干 `<r>` run 拼成纯文本（只取每个 run 的 `<t>`，跳过 `rPr`） */
function concatRuns(xml: string, runs: ElementRange[]): string {
  let out = '';
  for (const run of runs) {
    for (const kid of childElements(xml, run)) {
      if (nameMatches(kid.name, 't')) {
        out += elementText(xml, kid);
        break;
      }
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* 单元格                                                                      */
/* -------------------------------------------------------------------------- */

interface CellContext {
  xml: string;
  shared: readonly string[];
  warn: (msg: string) => void;
}

interface CellParts {
  vText?: string;
  fText?: string;
  inline?: { text: string; rich: boolean };
}

/**
 * 单元格子元素读取——**热路径**。
 *
 * 绝大多数单元格只有 `<v>`（值）和 `<f>`（公式）两种子元素，而百万格规模下
 * "通用访问器 + 每子元素建 ElementRange/attrs 对象 + 数组"的开销是主要成本（实测占解析时间大头）。
 * 这里用 indexOf 直接取 `<v>`/`<f>` 的内部文本，**零对象分配**；返回 `after` 让调用方直接跳到
 * 单元格内容之后，从而**不必为每个 `<c>` 再做一次 findCloseTag**（那正是 21% 的自身耗时来源）。
 *
 * 一旦遇到不认识的东西（`<is>` 内联字符串、`<extLst>` 等扩展），立刻退回通用实现，
 * 保证行为与通用路径完全一致——快路径只覆盖"确定安全"的子集。
 */
function readCellPartsFast(ctx: CellContext, contentStart: number, limit: number): { parts: CellParts; after: number } {
  const xml = ctx.xml;
  const parts: CellParts = {};
  let i = contentStart;

  while (i < limit) {
    const lt = xml.indexOf('<', i);
    if (lt < 0 || lt >= limit) return { parts, after: limit };
    const nx = xml.charCodeAt(lt + 1);
    if (nx === SLASH) return { parts, after: lt }; // 到 `</c` 了（调用方自行跳过）

    const isV = nx === LOWER_V;
    const isF = nx === LOWER_F;
    if (!isV && !isF) return { parts, after: -1 }; // <is>/<extLst>… 交给通用实现

    const afterName = lt + 2;
    const boundary = xml.charCodeAt(afterName);
    if (boundary !== GT && boundary !== SPACE && boundary !== TAB && boundary !== LF && boundary !== CR) {
      return { parts, after: -1 }; // 如 `<vFoo>`，不是我们要的标签
    }

    const gt = xml.indexOf('>', afterName);
    if (gt < 0 || gt >= limit) return { parts, after: -1 };

    // 自闭合：`<v/>`、共享公式的 `<f t="shared" si="0"/>`
    if (xml.charCodeAt(gt - 1) === SLASH) {
      if (isV) {
        if (parts.vText === undefined) parts.vText = '';
      } else if (parts.fText === undefined) {
        parts.fText = '';
      }
      i = gt + 1;
      continue;
    }

    const name = isV ? 'v' : 'f';
    const close = findCloseTag(xml, name, gt + 1, limit);
    if (close < 0) return { parts, after: -1 };

    const inner = xml.slice(gt + 1, close);
    const text = inner.indexOf('<') < 0 ? decodeEntities(inner) : decodeEntities(stripTags(inner));
    if (isV) {
      if (parts.vText === undefined) parts.vText = text;
    } else if (parts.fText === undefined) {
      parts.fText = text;
    }

    const gt2 = xml.indexOf('>', close + name.length + 2);
    if (gt2 < 0 || gt2 >= limit) return { parts, after: limit };
    i = gt2 + 1;
  }

  return { parts, after: limit };
}

/** 通用实现：认识所有子元素（`<is>` 富文本等），代价是每层都要建区间对象 */
function readCellPartsGeneric(ctx: CellContext, cell: CellNode): CellParts {
  const parts: CellParts = {};
  for (const kid of childElements(ctx.xml, cell as ElementRange)) {
    const ln = localName(kid.name);
    if (ln === 'v' && parts.vText === undefined) {
      parts.vText = elementText(ctx.xml, kid);
    } else if (ln === 'f' && parts.fText === undefined) {
      parts.fText = elementText(ctx.xml, kid);
    } else if (ln === 'is' && parts.inline === undefined) {
      parts.inline = readInlineString(ctx, kid);
    }
  }
  return parts;
}

/** 走通用路径读一个单元格的子元素（先自己算 `</c>`） */
function readCellParts(ctx: CellContext, cell: CellNode): CellParts {
  const fast = readCellPartsFast(ctx, cell.openEnd, cell.closeEnd);
  if (fast.after >= 0) return fast.parts;
  return readCellPartsGeneric(ctx, cell);
}

function readInlineString(ctx: CellContext, isEl: ElementRange): { text: string; rich: boolean } {
  const kids = childElements(ctx.xml, isEl);
  const runs = kids.filter((k) => nameMatches(k.name, 'r'));
  if (runs.length > 0) return { text: concatRuns(ctx.xml, runs), rich: true };
  const texts = kids.filter((k) => nameMatches(k.name, 't'));
  return { text: texts.map((t) => elementText(ctx.xml, t)).join(''), rich: false };
}

function cellValueFrom(cellType: string, parts: CellParts, ctx: CellContext): ParsedCellValue | undefined {
  // 内联字符串：规范要求 t="inlineStr"，但 `<is>` 只可能出现在内联字符串单元格里，
  // 因此即使 t 属性缺失/写错也按内联字符串取值（容错优先，避免整格内容丢失）。
  if (cellType === 'inlineStr' || (parts.inline !== undefined && parts.vText === undefined)) {
    const inline = parts.inline;
    if (!inline) return '';
    if (inline.rich) ctx.warn('内联富文本（inlineStr 含多个 <r> run）已降级为纯文本，run 级格式丢失');
    return inline.text;
  }

  const raw = parts.vText;
  if (raw === undefined) return undefined;

  switch (cellType) {
    case 's': {
      const idx = Number.parseInt(raw, 10);
      if (!Number.isFinite(idx)) {
        ctx.warn(`共享字符串索引 "${raw}" 不是合法整数，单元格按空值处理`);
        return null;
      }
      if (idx < 0 || idx >= ctx.shared.length) {
        ctx.warn(`共享字符串索引 ${idx} 越界（sharedStrings 共 ${ctx.shared.length} 条），单元格按空值处理`);
        return null;
      }
      return ctx.shared[idx];
    }
    case 'str':
      return raw;
    case 'b': {
      const t = raw.trim();
      if (t === '1' || t.toLowerCase() === 'true') return true;
      if (t === '0' || t.toLowerCase() === 'false' || t === '') return false;
      ctx.warn(`布尔单元格值 "${raw}" 无法识别，按 false 处理`);
      return false;
    }
    case 'e':
      // 错误值走 error 字段，value 保持 undefined
      return undefined;
    case 'n':
    case '':
    default: {
      const t = raw.trim();
      if (t === '') return undefined;
      const n = Number(t);
      if (!Number.isFinite(n)) {
        ctx.warn(`公式/数值缓存值 "${raw}" 不是合法数字，原始文本按字符串保留`);
        return raw;
      }
      return n;
    }
  }
}

/**
 * 单元格节点的**最小视图**：热路径用它替代完整的 `ElementRange`，
 * 以便复用一个 scratch 对象、不给百万个单元格各建一份区间对象。
 * `ElementRange` 天然满足这个结构，所以旧调用点无需改动。
 */
interface CellNode {
  attrs: XmlAttributes;
  openEnd: number;
  closeEnd: number;
  selfClosing: boolean;
}

function buildCell(ctx: CellContext, cell: CellNode, rowFallback: number, colFallback: number): ParsedCell | undefined {
  return buildCellFrom(ctx, cell, rowFallback, colFallback, readCellParts(ctx, cell));
}

function buildCellFrom(
  ctx: CellContext,
  cell: CellNode,
  rowFallback: number,
  colFallback: number,
  parts: CellParts,
): ParsedCell | undefined {
  const ref = cell.attrs['r'];
  let row = rowFallback;
  let col = colFallback;
  if (ref) {
    const parsed = parseCellRef(ref);
    if (parsed) {
      row = parsed.row;
      col = parsed.col;
    } else {
      ctx.warn(`单元格引用 "${ref}" 无法解析，已退化为按位置计数`);
    }
  }
  const cellType = cell.attrs['t'] ?? '';
  const out: ParsedCell = { row, col };
  const value = cellValueFrom(cellType, parts, ctx);
  if (value !== undefined) out.value = value;
  if (parts.fText !== undefined) {
    const f = parts.fText;
    out.formula = f.startsWith('=') ? f.slice(1) : f;
  }
  if (cellType === 'e' && parts.vText !== undefined) out.error = parts.vText.trim();
  const styleIndex = attrInt(cell.attrs, 's');
  if (styleIndex !== undefined && styleIndex >= 0) out.styleIndex = styleIndex;
  return out;
}

/* -------------------------------------------------------------------------- */
/* 行 / 列 / 合并 / 冻结                                                       */
/* -------------------------------------------------------------------------- */

/** 热路径复用的单元格视图（解析是同步单线程，不存在重入） */
const CELL_SCRATCH: CellNode = { attrs: {}, openEnd: 0, closeEnd: 0, selfClosing: false };

function parseSheetRows(
  xml: string,
  sheetData: ElementRange,
  ctx: CellContext,
  cells: ParsedCell[],
): Record<number, ParsedRowInfo> {
  const rows: Record<number, ParsedRowInfo> = {};
  let implicitRow = 0;
  const sheetEnd = sheetData.closeEnd;

  // 带命名空间前缀的写法（`<x:row>`）走不了内联扫描，直接回退通用实现
  if (xml.indexOf('<row', sheetData.openEnd) < 0 || sheetData.selfClosing) {
    return parseSheetRowsGeneric(xml, sheetData, ctx, cells);
  }

  let i = sheetData.openEnd;
  while (i < sheetEnd) {
    const lt = xml.indexOf('<row', i);
    if (lt < 0 || lt >= sheetEnd) break;
    const rb = xml.charCodeAt(lt + 4);
    if (rb !== GT && rb !== SPACE && rb !== TAB && rb !== LF && rb !== CR) {
      i = lt + 4;
      continue;
    }
    const gt = xml.indexOf('>', lt + 4);
    if (gt < 0 || gt >= sheetEnd) break;

    const rowAttrs = parseAttrs(xml, lt + 4, gt);
    const r = attrInt(rowAttrs, 'r');
    const rowIdx = r !== undefined && r > 0 ? r - 1 : implicitRow;
    implicitRow = rowIdx + 1;

    const info: ParsedRowInfo = {};
    const height = attrNumber(rowAttrs, 'ht');
    if (height !== undefined) info.height = height;
    const hidden = attrBool(rowAttrs, 'hidden');
    if (hidden !== undefined) info.hidden = hidden;
    const customHeight = attrBool(rowAttrs, 'customHeight');
    if (customHeight !== undefined) info.customHeight = customHeight;
    if (Object.keys(info).length > 0) rows[rowIdx] = info;

    const closeRow = xml.indexOf('</row', gt + 1);
    const rowLimit = closeRow < 0 || closeRow > sheetEnd ? sheetEnd : closeRow;

    let implicitCol = 0;
    let before = cells.length;
    let k = gt + 1;
    while (k < rowLimit) {
      const clt = xml.indexOf('<c', k);
      if (clt < 0 || clt >= rowLimit) break;
      const cb = xml.charCodeAt(clt + 2);
      if (cb !== GT && cb !== SPACE && cb !== TAB && cb !== LF && cb !== CR) {
        k = clt + 2;
        continue;
      }
      const cgt = xml.indexOf('>', clt + 2);
      if (cgt < 0 || cgt >= rowLimit) break;

      CELL_SCRATCH.attrs = parseAttrs(xml, clt + 2, cgt);
      CELL_SCRATCH.openEnd = cgt + 1;
      CELL_SCRATCH.selfClosing = xml.charCodeAt(cgt - 1) === SLASH;

      let parts: CellParts;
      let next: number;
      if (CELL_SCRATCH.selfClosing) {
        // `<c r="A1" s="1"/>`：没有子元素，单元格到开始标签就结束。
        // 必须显式处理——否则会把"下一个单元格"的内容当成自己的子元素读进来（吞格）。
        CELL_SCRATCH.closeEnd = cgt + 1;
        parts = {};
        next = cgt + 1;
      } else {
        const fast = readCellPartsFast(ctx, cgt + 1, rowLimit);
        parts = fast.parts;
        next = fast.after;
        if (next < 0) {
          // 复杂子元素（内联富文本 / 扩展）：补上 `</c>` 后走通用实现
          const close = findCloseTag(xml, 'c', cgt + 1, rowLimit);
          const gt2 = close < 0 ? -1 : xml.indexOf('>', close + 2);
          CELL_SCRATCH.closeEnd = gt2 < 0 ? rowLimit : gt2 + 1;
          parts = readCellPartsGeneric(ctx, CELL_SCRATCH);
          next = CELL_SCRATCH.closeEnd;
        } else {
          CELL_SCRATCH.closeEnd = next;
        }
      }

      const cell = buildCellFrom(ctx, CELL_SCRATCH, rowIdx, implicitCol, parts);
      if (cell) {
        cells.push(cell);
        implicitCol = cell.col + 1;
      }
      k = next > k ? next : k + 1;
    }

    // 安全网：这一行按 `<c>` 一个都没扫到、但行内容非空 → 说明写法不是标准无前缀形式，
    // 退回通用实现重扫这一行（宁可慢，也不能静默丢数据）
    if (cells.length === before && rowLimit > gt + 1) {
      for (const cellEl of childElements(xml, { ...rowAttrsOwner(xml, lt, gt, rowLimit) })) {
        if (!nameMatches(cellEl.name, 'c')) continue;
        const cell = buildCell(ctx, cellEl, rowIdx, implicitCol);
        if (!cell) continue;
        cells.push(cell);
        implicitCol = cell.col + 1;
      }
    }

    if (closeRow < 0 || closeRow >= sheetEnd) break;
    i = closeRow + 6;
  }

  return rows;
}

/** 构造一个"仅用于通用回退扫描"的行区间对象 */
function rowAttrsOwner(xml: string, lt: number, gt: number, rowLimit: number): ElementRange {
  return {
    name: 'row',
    openStart: lt,
    openEnd: gt + 1,
    closeEnd: rowLimit,
    selfClosing: false,
    attrs: {},
    start: lt,
  };
}

/** 通用实现：支持任意前缀与嵌套写法（热路径的安全网） */
function parseSheetRowsGeneric(
  xml: string,
  sheetData: ElementRange,
  ctx: CellContext,
  cells: ParsedCell[],
): Record<number, ParsedRowInfo> {
  const rows: Record<number, ParsedRowInfo> = {};
  let implicitRow = 0;
  for (const rowEl of childElements(xml, sheetData)) {
    if (!nameMatches(rowEl.name, 'row')) continue;
    const r = attrInt(rowEl.attrs, 'r');
    const rowIdx = r !== undefined && r > 0 ? r - 1 : implicitRow;
    implicitRow = rowIdx + 1;

    const info: ParsedRowInfo = {};
    const height = attrNumber(rowEl.attrs, 'ht');
    if (height !== undefined) info.height = height;
    const hidden = attrBool(rowEl.attrs, 'hidden');
    if (hidden !== undefined) info.hidden = hidden;
    const customHeight = attrBool(rowEl.attrs, 'customHeight');
    if (customHeight !== undefined) info.customHeight = customHeight;
    if (Object.keys(info).length > 0) rows[rowIdx] = info;

    let implicitCol = 0;
    for (const cellEl of childElements(xml, rowEl)) {
      if (!nameMatches(cellEl.name, 'c')) continue;
      const cell = buildCell(ctx, cellEl, rowIdx, implicitCol);
      if (!cell) continue;
      cells.push(cell);
      implicitCol = cell.col + 1;
    }
  }
  return rows;
}

function parseCols(xml: string, colsEl: ElementRange): Record<number, ParsedColInfo> {
  const cols: Record<number, ParsedColInfo> = {};
  for (const colEl of childElements(xml, colsEl)) {
    if (!nameMatches(colEl.name, 'col')) continue;
    const min = attrInt(colEl.attrs, 'min');
    const max = attrInt(colEl.attrs, 'max');
    if (min === undefined || max === undefined || min < 1 || max < min) continue;
    const info: ParsedColInfo = {};
    const width = attrNumber(colEl.attrs, 'width');
    // 原样保留 Excel 字符宽度，不转像素
    if (width !== undefined) info.width = width;
    const hidden = attrBool(colEl.attrs, 'hidden');
    if (hidden !== undefined) info.hidden = hidden;
    const customWidth = attrBool(colEl.attrs, 'customWidth');
    if (customWidth !== undefined) info.customWidth = customWidth;
    if (Object.keys(info).length === 0) continue;
    const upper = Math.min(max, 20000);
    for (let c = min; c <= upper; c++) cols[c - 1] = { ...info };
  }
  return cols;
}

function parseMerges(xml: string, warn: (msg: string) => void, from = 0): ParsedMerge[] {
  const holder = findFirstElement(xml, 'mergeCells', from);
  if (!holder) return [];
  const merges: ParsedMerge[] = [];
  for (const mergeEl of childElements(xml, holder)) {
    if (!nameMatches(mergeEl.name, 'mergeCell')) continue;
    const ref = mergeEl.attrs['ref'];
    if (!ref) continue;
    const range = parseRangeRef(ref);
    if (range) merges.push(range);
    else warn(`合并单元格引用 "${ref}" 无法解析，已跳过`);
  }
  return merges;
}

/* -------------------------------------------------------------------------- */
/* 冻结窗格 / 视图 / 维度                                                      */
/* -------------------------------------------------------------------------- */

interface SheetViewInfo {
  freeze?: { row: number; col: number };
  gridlinesHidden?: boolean;
}

function parseSheetViews(xml: string, warn: (msg: string) => void): SheetViewInfo {
  const info: SheetViewInfo = {};
  const viewsEl = findFirstElement(xml, 'sheetViews');
  if (!viewsEl) return info;
  for (const viewEl of childElements(xml, viewsEl)) {
    if (!nameMatches(viewEl.name, 'sheetView')) continue;
    const showGridLines = attrBool(viewEl.attrs, 'showGridLines');
    if (showGridLines === false && info.gridlinesHidden === undefined) info.gridlinesHidden = true;
    for (const kid of childElements(xml, viewEl)) {
      if (!nameMatches(kid.name, 'pane')) continue;
      const state = kid.attrs['state'] ?? 'split';
      if (state !== 'frozen' && state !== 'frozenSplit') continue; // split 不是冻结
      const xSplit = attrNumber(kid.attrs, 'xSplit') ?? 0;
      const ySplit = attrNumber(kid.attrs, 'ySplit') ?? 0;
      const col = Math.max(0, Math.trunc(xSplit));
      const row = Math.max(0, Math.trunc(ySplit));
      if (row === 0 && col === 0) {
        const topLeft = kid.attrs['topLeftCell'];
        const derived = topLeft ? parseCellRef(topLeft) : undefined;
        if (derived) {
          info.freeze = { row: derived.row, col: derived.col };
        } else {
          warn('冻结窗格 pane 的 xSplit/ySplit 均为 0 且 topLeftCell 无法解析，已忽略冻结设置');
        }
        continue;
      }
      if (info.freeze === undefined) info.freeze = { row, col };
    }
  }
  return info;
}

function parseDimension(xml: string): ParsedSheet['dimension'] {
  const el = findFirstElement(xml, 'dimension');
  if (!el) return undefined;
  const ref = el.attrs['ref'];
  if (!ref) return undefined;
  const range = parseRangeRef(ref);
  if (!range) return undefined;
  return range;
}

interface FormatPr {
  defaultRowHeight?: number;
  defaultColWidth?: number;
}

function parseSheetFormatPr(xml: string): FormatPr {
  const el = findFirstElement(xml, 'sheetFormatPr');
  if (!el) return {};
  const out: FormatPr = {};
  const rowHeight = attrNumber(el.attrs, 'defaultRowHeight');
  if (rowHeight !== undefined) out.defaultRowHeight = rowHeight;
  const colWidth = attrNumber(el.attrs, 'defaultColWidth');
  if (colWidth !== undefined) out.defaultColWidth = colWidth;
  return out;
}

/* -------------------------------------------------------------------------- */
/* 超链接 / 表格 / 绘图 的**引用**（关系 id 由编排层结合 sheet rels 解析）          */
/* -------------------------------------------------------------------------- */

/** `<hyperlink>` 的原始属性（`r:id` 尚未解析成目标） */
export interface RawHyperlink {
  ref: string;
  relId?: string;
  location?: string;
  display?: string;
  tooltip?: string;
}

/** 工作表里指向外部部件的引用 */
export interface SheetPartRefs {
  /** `<hyperlinks>` 里的超链接（含空容器：`hasHyperlinks` 为 true 但数组为空） */
  hyperlinks: RawHyperlink[];
  hasHyperlinks: boolean;
  /** `<tableParts><tablePart r:id>` 的顺序即 Excel 里的表格顺序 */
  tableRelIds: string[];
  /** `<drawing r:id>` */
  drawingRelId?: string;
}

/**
 * 单遍扫描工作表 XML，取出超链接 / 表格 / 绘图三处**引用**。
 * 与 `collectUnsupported` 一样是纯 tokenizer 级扫描：只对目标元素建对象，不建整棵树。
 */
export function parseSheetPartRefs(xml: string, from = 0): SheetPartRefs {
  const refs: SheetPartRefs = { hyperlinks: [], hasHyperlinks: false, tableRelIds: [] };
  // 只对这三种元素回调：无 `match` 时**每个**元素都会回调并建区间/属性对象，
  // 百万格的表就是几百万次无用对象分配（实测占解析时间可观）。
  // `from` 由编排层给出（这些都排在 sheetData 之后），避免再走一遍整个表数据。
  visitElements(xml, (el) => {
    const ln = localName(el.name);
    if (ln === 'hyperlinks') {
      refs.hasHyperlinks = true;
      for (const link of childElements(xml, el)) {
        if (localName(link.name) !== 'hyperlink') continue;
        const ref = link.attrs['ref'];
        if (!ref) continue;
        const raw: RawHyperlink = { ref };
        const relId = link.attrs['r:id'] ?? link.attrs['id'];
        if (relId) raw.relId = relId;
        const location = link.attrs['location'];
        if (location !== undefined) raw.location = location;
        const display = link.attrs['display'];
        if (display !== undefined) raw.display = display;
        const tooltip = link.attrs['tooltip'];
        if (tooltip !== undefined) raw.tooltip = tooltip;
        refs.hyperlinks.push(raw);
      }
      return false; // 子树已处理完
    }
    if (ln === 'tableParts') {
      for (const part of childElements(xml, el)) {
        if (localName(part.name) !== 'tablePart') continue;
        const relId = part.attrs['r:id'] ?? part.attrs['id'];
        if (relId) refs.tableRelIds.push(relId);
      }
      return false;
    }
    if (ln === 'drawing') {
      const relId = el.attrs['r:id'] ?? el.attrs['id'];
      if (relId && refs.drawingRelId === undefined) refs.drawingRelId = relId;
      return false;
    }
    return undefined;
  }, { match: (name) => isPartRefName(name), from });
  return refs;
}

function isPartRefName(name: string): boolean {
  const ln = localName(name);
  return ln === 'hyperlinks' || ln === 'tableParts' || ln === 'drawing';
}

/**
 * 把 `<hyperlink>` + sheet rels 的目标合成契约里的 `ParsedHyperlink`。
 *
 * `location` 以 `#` 开头（Excel 内部链接两种写法）时去掉 `#`；
 * rels 的 `Target` 若本身是 `#Sheet!A1`，它表达的是"文档内位置"而不是外部 URL，
 * 因此并入 `location`，不写进 `target`（否则下游会把 `#...` 当成可点击的外链）。
 */
function buildHyperlinks(
  raw: readonly RawHyperlink[],
  targets: ReadonlyMap<string, string> | undefined,
  warn: Warn,
): ParsedHyperlink[] {
  const out: ParsedHyperlink[] = [];
  for (const link of raw) {
    const item: ParsedHyperlink = { ref: link.ref };
    let location = link.location;
    const relId = link.relId;
    const relTarget = relId !== undefined ? targets?.get(relId) : undefined;
    if (relTarget !== undefined) {
      if (relTarget.startsWith('#')) {
        if (location === undefined) location = relTarget;
      } else {
        item.target = relTarget;
      }
    } else if (relId !== undefined && targets !== undefined && !targets.has(relId)) {
      warn(`超链接 ${link.ref} 的关系 ${relId} 在 sheet rels 里找不到，已只保留显示信息`);
    }
    if (location !== undefined) item.location = location.startsWith('#') ? location.slice(1) : location;
    if (link.display !== undefined) item.display = link.display;
    if (link.tooltip !== undefined) item.tooltip = link.tooltip;
    out.push(item);
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* 不支持特性记账                                                              */
/* -------------------------------------------------------------------------- */

/**
 * 仍**不解析**的元素 -> 记账标签（按种类计数，只统计每类一条，避免 report 被上万条撑爆）。
 *
 * P1 变更：conditionalFormatting / dataValidations / hyperlinks / tableParts / comments /
 * drawing（浮动图片）已经从这张表里移除——它们现在真的被解析进中立模型了。
 * 剩下的是确实没还原的：打印设置、迷你图、排序状态、VML 绘图、OLE、保护、透视表等。
 */
const UNSUPPORTED_LABELS: Readonly<Record<string, string>> = {
  extLst: '扩展列表(extLst)',
  autoFilter: '自动筛选(autoFilter)',
  picture: '工作表背景图片(picture)',
  oleObjects: '嵌入对象(oleObjects)',
  legacyDrawing: '旧式 VML 绘图(legacyDrawing)',
  pivotTableDefinition: '数据透视表(pivotTableDefinition)',
  sheetProtection: '工作表保护(sheetProtection)',
  protectedRanges: '保护区域(protectedRanges)',
  sparklineGroups: '迷你图(sparklineGroups)',
  sortState: '排序状态(sortState)',
  dataConsolidate: '合并计算(dataConsolidate)',
  scenarios: '方案(scenarios)',
  cellWatches: '监视窗口(cellWatches)',
  ignoredErrors: '忽略错误标记(ignoredErrors)',
};

/**
 * 打印相关标签：**不再计入"未支持"**，改成解析成真值 + 进 `report.preserved`。
 *
 * 理由（用户实测反馈）：导入座位表时状态栏显示"未支持 1 项"，点开是"打印设置"——
 * 而打印设置只影响打印，屏幕预览与内容编辑完全不受影响，导出又走字节级保真路线原样保留它，
 * 把它列为"未支持"是误导。现在的口径：**能读到值就记下来，并在导入摘要里说明"不影响预览"**。
 */
const PRINT_TAGS: ReadonlySet<string> = new Set([
  'pageSetup',
  'pageMargins',
  'printOptions',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
]);

/**
 * x14/x15 扩展里"同一功能的现代写法"。它们用**带前缀的标签名**出现
 * （`<x14:dataValidations>`、`<x14:conditionalFormattings>`），
 * 而带前缀就意味着走的不是我们已经支持的那条老路径，因此单独记账。
 */
const X14_LABELS: Readonly<Record<string, string>> = {
  dataValidations: 'x14 扩展数据验证(x14:dataValidations)',
  conditionalFormattings: 'x14 扩展条件格式(x14:conditionalFormattings)',
  sparklineGroups: 'x14 迷你图(x14:sparklineGroups)',
};

/**
 * 单遍扫描工作表 XML，把**不解析**的特性按种类计数成 `标签 · N 处（未解析）`。
 *
 * 只扫标签（跳过注释 / PI / 结束标签），不建元素对象、不复制文本；
 * 命中即记账，但**不跳过子树**——`<extLst>` 里还可能嵌着 x14 版本的另一类特性，
 * 跳过会漏记（实测 fixture-rules.xlsx 的 `<extLst>` 里就有 x14 条件格式）。
 */
function collectUnsupported(xml: string): string[] {
  const counts = new Map<string, number>();
  const len = xml.length;
  let i = 0;
  while (i < len) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    const nx = xml.charCodeAt(lt + 1);
    if (nx === 63 || nx === 33) { // <? ?> / <!-- -->
      const end = xml.indexOf('>', lt + 2);
      if (end < 0) break;
      i = end + 1;
      continue;
    }
    if (nx === 47) { // </...>
      const end = xml.indexOf('>', lt + 2);
      if (end < 0) break;
      i = end + 1;
      continue;
    }
    let p = lt + 1;
    const nameStart = p;
    while (p < len && !isNameStop(xml.charCodeAt(p))) p++;
    const name = xml.slice(nameStart, p);
    let q = p;
    let quote = 0;
    let gt = -1;
    while (q < len) {
      const c = xml.charCodeAt(q);
      if (quote !== 0) { if (c === quote) quote = 0; }
      else if (c === 34 || c === 39) quote = c;
      else if (c === 62) { gt = q; break; }
      q++;
    }
    if (gt < 0) break;
    const ln = localName(name);
    // 打印相关标签走"已解析并原样保留"那条路（见 PRINT_TAGS 的说明），不算未支持
    if (PRINT_TAGS.has(ln)) {
      i = gt + 1;
      continue;
    }
    // 带前缀（x14:/x15:/xr:…）的现代写法优先按扩展表记账，避免与老路径混淆
    const label = (name.length !== ln.length ? X14_LABELS[ln] : undefined) ?? UNSUPPORTED_LABELS[ln];
    if (label) counts.set(label, (counts.get(label) ?? 0) + 1);
    i = gt + 1;
  }
  return [...counts.entries()].map(([label, n]) => `${label} · ${n} 处（未解析）`);
}

/** OOXML paperSize -> 常见纸型名（只列常见的几种，其余直接显示数字） */
const PAPER_SIZES: Readonly<Record<number, string>> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  8: 'A3',
  9: 'A4',
  11: 'A5',
  12: 'B4',
  13: 'B5',
};

/** 去掉浮点尾巴：0.75 → "0.75"、1.0 → "1" */
function trimNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(3)));
}

/**
 * 解析一页工作表的打印设置（`pageSetup` / `pageMargins` / `headerFooter` / `printOptions` / 分页符）。
 *
 * 这些字段只影响打印，不影响屏幕预览；解析它们的目的**不是**在预览里还原，
 * 而是：① 让导入摘要能说清"文件里有什么、我们怎么处理的"；
 * ② 不再把它们误报成"未支持"。导出时它们由字节级保真路线原样保留。
 */
export function parsePrintSettings(xml: string): { print?: ParsedPrintSettings; summary?: string } {
  const print: ParsedPrintSettings = {};
  let touched = false;

  const setup = findFirstElement(xml, 'pageSetup');
  if (setup) {
    touched = true;
    const paperSize = attrInt(setup.attrs, 'paperSize');
    if (paperSize !== undefined) print.paperSize = paperSize;
    const orientation = setup.attrs['orientation'];
    if (orientation === 'landscape' || orientation === 'portrait') print.orientation = orientation;
    const scale = attrInt(setup.attrs, 'scale');
    if (scale !== undefined) print.scale = scale;
    const fitToWidth = attrInt(setup.attrs, 'fitToWidth');
    if (fitToWidth !== undefined) print.fitToWidth = fitToWidth;
    const fitToHeight = attrInt(setup.attrs, 'fitToHeight');
    if (fitToHeight !== undefined) print.fitToHeight = fitToHeight;
  }

  const margins = findFirstElement(xml, 'pageMargins');
  if (margins) {
    touched = true;
    const read = (name: string): number | undefined => attrNumber(margins.attrs, name);
    const left = read('left') ?? 0;
    const right = read('right') ?? 0;
    const top = read('top') ?? 0;
    const bottom = read('bottom') ?? 0;
    print.margins = { left, right, top, bottom };
    const header = read('header');
    const footer = read('footer');
    if (header !== undefined) print.margins.header = header;
    if (footer !== undefined) print.margins.footer = footer;
  }

  const hf = findFirstElement(xml, 'headerFooter');
  if (hf) {
    touched = true;
    const texts: NonNullable<ParsedPrintSettings['headerFooter']> = {};
    for (const kid of childElements(xml, hf)) {
      const ln = localName(kid.name);
      const text = elementText(xml, kid)?.trim();
      if (!text) continue;
      if (ln === 'oddHeader') texts.oddHeader = text;
      else if (ln === 'oddFooter') texts.oddFooter = text;
      else if (ln === 'evenHeader') texts.evenHeader = text;
      else if (ln === 'evenFooter') texts.evenFooter = text;
    }
    if (Object.keys(texts).length > 0) print.headerFooter = texts;
  }

  const options = findFirstElement(xml, 'printOptions');
  if (options) {
    touched = true;
    if (attrBool(options.attrs, 'horizontalCentered') === true || attrBool(options.attrs, 'verticalCentered') === true) {
      print.centered = true;
    }
  }

  const rowBreaks = findFirstElement(xml, 'rowBreaks');
  const colBreaks = findFirstElement(xml, 'colBreaks');
  if (rowBreaks || colBreaks) {
    touched = true;
    print.breaks = {
      row: rowBreaks ? countBrk(xml, rowBreaks) : 0,
      col: colBreaks ? countBrk(xml, colBreaks) : 0,
    };
  }

  if (!touched) return {};
  return { print, summary: `打印设置已读取并原样保留：${describePrint(print)}（只影响打印，不影响预览与编辑）` };
}

/** 数一数 `<rowBreaks>` / `<colBreaks>` 里有多少个 `<brk>`（拿不到 count 属性时兜底） */
function countBrk(xml: string, holder: ElementRange): number {
  const declared = attrInt(holder.attrs, 'count');
  if (declared !== undefined) return declared;
  let n = 0;
  for (const kid of childElements(xml, holder)) if (localName(kid.name) === 'brk') n += 1;
  return n;
}

/** 打印设置 -> 一行中文说明（只写文件里真的有的项，不编默认值） */
function describePrint(print: ParsedPrintSettings): string {
  const parts: string[] = [];
  if (print.paperSize !== undefined) parts.push(`纸张 ${PAPER_SIZES[print.paperSize] ?? `#${print.paperSize}`}`);
  if (print.orientation) parts.push(print.orientation === 'landscape' ? '横向' : '纵向');
  if (print.scale !== undefined) parts.push(`缩放 ${print.scale}%`);
  if (print.fitToWidth !== undefined || print.fitToHeight !== undefined) {
    parts.push(`打印适应 ${print.fitToWidth ?? 1} 页宽 × ${print.fitToHeight ?? 1} 页高`);
  }
  if (print.margins) {
    const m = print.margins;
    parts.push(`页边距 左${trimNumber(m.left)}/右${trimNumber(m.right)}/上${trimNumber(m.top)}/下${trimNumber(m.bottom)}`);
  }
  if (print.headerFooter) {
    const which: string[] = [];
    if (print.headerFooter.oddHeader || print.headerFooter.evenHeader) which.push('页眉');
    if (print.headerFooter.oddFooter || print.headerFooter.evenFooter) which.push('页脚');
    if (which.length > 0) parts.push(`含${which.join('与')}文字`);
  }
  if (print.centered) parts.push('居中打印');
  if (print.breaks && print.breaks.row + print.breaks.col > 0) {
    parts.push(`手工分页符 ${print.breaks.row + print.breaks.col} 处`);
  }
  return parts.length > 0 ? parts.join(' · ') : '文件里只有空的打印设置占位（无需处理）';
}

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

export interface WorksheetDiagnostics {
  warnings: string[];
  unsupported: string[];
  /** 已解析并原样保留、但不影响预览的信息（打印设置）；**不是**缺陷 */
  preserved: string[];
}

export interface ParseWorksheetInput {
  id: string;
  name: string;
  index: number;
  hidden?: boolean;
  veryHidden?: boolean;
  xml: string;
  sharedStrings: readonly string[];
  /* ---- P1 新增（全部可选，缺省时行为与 P0 完全一致） ---- */
  /** 已扫描好的部件引用；给了就不必再扫一遍工作表 XML */
  refs?: SheetPartRefs;
  /** sheet rels 解析出的超链接目标：relId -> Target（外部链接原样，内部 `#A1` 形式会被并进 location） */
  hyperlinkTargets?: ReadonlyMap<string, string>;
  /** 编排层（index.ts）已解析好的外部部件：批注 / 表格 / 图片 */
  parts?: {
    notes?: ParsedNote[];
    tables?: ParsedTable[];
    images?: ParsedImage[];
  };
}

export interface ParseWorksheetResult {
  sheet: ParsedSheet;
  diagnostics: WorksheetDiagnostics;
}

export function parseWorksheet(input: ParseWorksheetInput): ParseWorksheetResult {
  const warnings: string[] = [];
  const warn = (msg: string): void => { warnings.push(msg); };
  const ctx: CellContext = { xml: input.xml, shared: input.sharedStrings, warn };

  const sheet: ParsedSheet = {
    id: input.id,
    name: input.name,
    index: input.index,
    rows: {},
    cols: {},
    cells: [],
    merges: [],
  };
  if (input.hidden) sheet.hidden = true;
  if (input.veryHidden) sheet.veryHidden = true;

  const sheetData = findFirstElement(input.xml, 'sheetData');
  if (!sheetData) {
    warn('工作表缺少 <sheetData>，按空表处理');
  } else {
    sheet.rows = parseSheetRows(input.xml, sheetData, ctx, sheet.cells);
  }

  // `sheetData` 之后的"尾巴"起点：合并区/超链接/表格/绘图等一律排在表数据之后，
  // 从这里起步就不必为每个特性各走一遍整个表数据（百万格时这就是主要开销）。
  const tail = sheetData ? sheetData.closeEnd : 0;
  const atLeast = (localName: string): number => earliestNameAt(input.xml, localName, tail);

  const colsEl = findFirstElement(input.xml, 'cols');
  if (colsEl) sheet.cols = parseCols(input.xml, colsEl);

  sheet.merges = parseMerges(input.xml, warn, sheetData ? atLeast('mergeCells') : 0);

  const view = parseSheetViews(input.xml, warn);
  if (view.freeze) sheet.freeze = view.freeze;
  if (view.gridlinesHidden) sheet.gridlinesHidden = true;

  const dim = parseDimension(input.xml);
  if (dim) sheet.dimension = dim;

  const formatPr = parseSheetFormatPr(input.xml);
  if (formatPr.defaultRowHeight !== undefined) sheet.defaultRowHeight = formatPr.defaultRowHeight;
  if (formatPr.defaultColWidth !== undefined) sheet.defaultColWidth = formatPr.defaultColWidth;

  // 打印设置：解析成真值（只影响打印 → 不报警，进 report.preserved，见 PRINT_TAGS 说明）
  const printInfo = parsePrintSettings(input.xml);
  if (printInfo.print) sheet.print = printInfo.print;

  // 超链接：本表的 `<hyperlink>` + sheet rels 的目标
  const refs = input.refs ?? parseSheetPartRefs(input.xml, sheetData ? atLeast('hyperlinks') : 0);
  if (refs.hasHyperlinks) {
    const links = buildHyperlinks(refs.hyperlinks, input.hyperlinkTargets, warn);
    if (links.length > 0) sheet.hyperlinks = links;
  }

  // 批注 / 表格 / 图片由编排层解析后注入（本文件不碰 zip）
  const parts = input.parts;
  if (parts) {
    if (parts.notes && parts.notes.length > 0) sheet.notes = parts.notes;
    if (parts.tables && parts.tables.length > 0) sheet.tables = parts.tables;
    if (parts.images && parts.images.length > 0) sheet.images = parts.images;
  }

  return {
    sheet,
    diagnostics: {
      warnings,
      unsupported: collectUnsupported(input.xml),
      // 不算缺陷、但要跟用户说清楚的（目前是打印设置）
      preserved: printInfo.summary ? [printInfo.summary] : [],
    },
  };
}

/* -------------------------------------------------------------------------- */
/* 局部工具                                                                    */
/* -------------------------------------------------------------------------- */

function isNameStop(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13 || c === 47 || c === 62 || c === 61;
}
