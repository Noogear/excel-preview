/**
 * **把非 OOXML 的表格（CSV / ODS / XLS）先变成一份我们自己的 xlsx，再走既有全链路。**
 *
 * 为什么是这个思路（而不是"为每种格式各写一条导入管线"）：
 *  - 本工具的价值在"逐格保真预览 + 外科式回写导出 + 会话恢复"，这三件事**全都建立在
 *    OOXML 包（zip + XML）之上**：导出是拿原始 zip 条目打补丁、会话存的是原始字节。
 *  - 所以只要能把别的格式**翻译成一份合法的最小 xlsx**，下游（解析 → 适配 → Univer → 编辑 →
 *    工作区 → 撤销 → 导出 → 会话恢复）一行都不用改，也不会出现"某种格式少了某个能力"。
 *  - 代价是导入时多一次"写包 + 再解析"（毫秒级，且只在小文件上发生）。
 *
 * 这个 writer 的输出必须**同时**被三方接受，因此有对应的三组测试：
 *  ① 我们自己的 `parseXlsx`（一致性与后续所有功能）；
 *  ② 会话恢复（字节要能反复解析）；
 *  ③ 真 Excel（导出的产物要能打开）——由 `tools/make-legacy-fixtures.ps1` 的往返样例间接保证。
 *
 * 只写"够用且规范"的部分：一个工作表 = sheetData + cols + mergeCells + sheetViews(冻结) +
 * dimension + pageMargins；样式表 = numFmts/fonts/fills/borders/cellXfs（去重后intern）。
 * 字符串一律用 **inlineStr**（不建 sharedStrings 部件，少一个能写错的地方；Excel 与我们的解析器都认）。
 */
import { strToU8, zipSync, type Zippable } from 'fflate';

/* -------------------------------------------------------------------------- */
/* 输入模型（CSV / ODS / XLS 三个解析器的公共出口）                              */
/* -------------------------------------------------------------------------- */

/** 单元格值：只保留"内容"三态（数字 / 文本 / 布尔）与公式 */
export interface SynthCell {
  row: number;
  col: number;
  value?: string | number | boolean | null;
  /** 不带前导 `=` */
  formula?: string;
  /** 指向 `WorkbookInput.styles` 的下标（缺省 = 默认样式） */
  style?: number;
}

/** 一份样式（各家格式都往这个中性形状上映射；字段名与 `ParsedStyle` 对齐） */
export interface SynthStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeThrough?: boolean;
  /** `#RRGGBB` */
  color?: string;
  /** `#RRGGBB` 纯色填充 */
  fill?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  textWrap?: boolean;
  /** 0–180，Excel 顺时针为正 */
  textRotation?: number;
  /** Excel 数字格式串，如 `0.00`、`yyyy-mm-dd`、`#,##0.00` */
  numberFormat?: string;
  border?: {
    top?: { style: string; color?: string };
    right?: { style: string; color?: string };
    bottom?: { style: string; color?: string };
    left?: { style: string; color?: string };
  };
}

export interface SynthSheet {
  name: string;
  hidden?: boolean;
  cells: SynthCell[];
  /** A1 记号，如 `A1:C1` */
  merges?: string[];
  /** 列宽（Excel 字符宽度）：键是 0-based 列号 */
  colWidths?: Record<number, number>;
  /** 行高（磅）：键是 0-based 行号 */
  rowHeights?: Record<number, number>;
  /** 冻结：切分行/列数（0 表示不冻） */
  freeze?: { rows: number; cols: number };
  gridlinesHidden?: boolean;
  defaultRowHeight?: number;
  defaultColWidth?: number;
}

export interface WorkbookInput {
  /** 工作簿名（我们自己用；xlsx 里不存） */
  name?: string;
  sheets: SynthSheet[];
  /** 样式表（下标即 `SynthCell.style`） */
  styles?: SynthStyle[];
}

/* -------------------------------------------------------------------------- */
/* 小工具                                                                      */
/* -------------------------------------------------------------------------- */

const NAMESPACE_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NAMESPACE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** XML 文本转义（属性与文本共用；`"` 也转掉，属性里就安全了） */
export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** 去掉颜色里的 alpha（`#AARRGGBB` → `#RRGGBB`），非法值返回 undefined */
export function normalizeColor(color: string | undefined): string | undefined {
  if (typeof color !== 'string') return undefined;
  const hex = color.trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{8}$/.test(hex)) return `#${hex.slice(2)}`;
  if (/^[0-9A-F]{6}$/.test(hex)) return `#${hex}`;
  if (/^[0-9A-F]{3}$/.test(hex)) return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
  return undefined;
}

function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function cellRef(row: number, col: number): string {
  return `${colLetter(col)}${row + 1}`;
}

/**
 * 工作表名的合法化（Excel 的硬规则）：≤31 字符、不能含 `[ ] : * ? / \`、
 * 不能以单引号开头结尾、不能为空、不能重名。
 */
export function sanitizeSheetName(raw: string, used: Set<string>, fallback = '工作表'): string {
  let name = (raw ?? '').replace(/[[\]:*?/\\]/g, ' ').replace(/\s+/g, ' ').trim();
  if (name.startsWith("'")) name = name.slice(1);
  if (name.endsWith("'")) name = name.slice(0, -1);
  name = name.trim().slice(0, 31);
  if (!name) name = fallback;
  let candidate = name;
  let n = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${n})`;
    candidate = name.slice(0, 31 - suffix.length) + suffix;
    n += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

/* -------------------------------------------------------------------------- */
/* 样式表：把 `SynthStyle[]` intern 成 fonts/fills/borders/numFmts/cellXfs      */
/* -------------------------------------------------------------------------- */

interface XfRecord {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  applyAlignment?: boolean;
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  textRotation?: number;
}

/** 边框线型：我们的中性名 → OOXML 的 style 名（未知的退回 thin，绝不写非法值） */
const BORDER_STYLE_MAP: Readonly<Record<string, string>> = {
  thin: 'thin',
  hair: 'hair',
  dotted: 'dotted',
  dashed: 'dashed',
  dashDot: 'dashDot',
  dashDotDot: 'dashDotDot',
  double: 'double',
  medium: 'medium',
  mediumDashed: 'mediumDashed',
  mediumDashDot: 'mediumDashDot',
  mediumDashDotDot: 'mediumDashDotDot',
  slantDashDot: 'slantDashDot',
  thick: 'thick',
};

class StyleTable {
  private readonly fonts: string[] = [];
  private readonly fills: string[] = [];
  private readonly borders: string[] = [];
  private readonly numFmts = new Map<string, number>();
  private readonly xfs: XfRecord[] = [];
  private readonly fontKeys = new Map<string, number>();
  private readonly fillKeys = new Map<string, number>();
  private readonly borderKeys = new Map<string, number>();
  private readonly xfKeys = new Map<string, number>();

  constructor() {
    // 规范要求 0/1 号填充分别是 none 与 gray125（我们自己的解析器把两者都当"无填充"）
    this.fills.push('<fill><patternFill patternType="none"/></fill>');
    this.fills.push('<fill><patternFill patternType="gray125"/></fill>');
    this.borders.push('<border><left/><right/><top/><bottom/><diagonal/></border>');
    this.fonts.push('<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>');
    this.xfs.push({ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 });
  }

  /** 把一个中性样式登记进表里，返回 cellXfs 下标 */
  intern(style: SynthStyle | undefined): number {
    const fontId = this.internFont(style);
    const fillId = this.internFill(style?.fill);
    const borderId = this.internBorder(style?.border);
    const { numFmtId, custom } = this.internNumFmt(style?.numberFormat);
    if (custom) this.customNumFmtIds.add(numFmtId);
    const xf: XfRecord = { fontId, fillId, borderId, numFmtId };
    if (style?.bold || style?.italic || style?.underline || style?.strikeThrough || style?.color || style?.fontFamily || style?.fontSize !== undefined) {
      xf.applyFont = true;
    }
    if (fillId !== 0) xf.applyFill = true;
    if (borderId !== 0) xf.applyBorder = true;
    if (style?.numberFormat) xf.applyNumberFormat = true;
    if (style?.horizontalAlign) {
      xf.horizontal = style.horizontalAlign;
      xf.applyAlignment = true;
    }
    if (style?.verticalAlign) {
      xf.vertical = style.verticalAlign === 'middle' ? 'center' : style.verticalAlign;
      xf.applyAlignment = true;
    }
    if (style?.textWrap) {
      xf.wrapText = true;
      xf.applyAlignment = true;
    }
    if (typeof style?.textRotation === 'number' && style.textRotation !== 0) {
      // OOXML 的 textRotation：1–90 顺时针，91–180 表示逆时针 1–90
      const rotation = Math.max(-90, Math.min(90, Math.round(style.textRotation)));
      xf.textRotation = rotation < 0 ? 90 - rotation : rotation;
      xf.applyAlignment = true;
    }
    const key = JSON.stringify(xf);
    const existing = this.xfKeys.get(key);
    if (existing !== undefined) return existing;
    const id = this.xfs.length;
    this.xfs.push(xf);
    this.xfKeys.set(key, id);
    return id;
  }

  private readonly customNumFmtIds = new Set<number>();

  private internFont(style: SynthStyle | undefined): number {
    if (!style) return 0;
    const size = typeof style.fontSize === 'number' && style.fontSize > 0 ? Math.round(style.fontSize * 2) / 2 : 11;
    const color = normalizeColor(style.color);
    const name = style.fontFamily?.trim() || 'Calibri';
    const key = `${name}|${size}|${style.bold ? 1 : 0}${style.italic ? 1 : 0}${style.underline ? 1 : 0}${style.strikeThrough ? 1 : 0}|${color ?? ''}`;
    const existing = this.fontKeys.get(key);
    if (existing !== undefined) return existing;
    const parts = [
      '<font>',
      style.bold ? '<b/>' : '',
      style.italic ? '<i/>' : '',
      style.strikeThrough ? '<strike/>' : '',
      style.underline ? '<u/>' : '',
      `<sz val="${size}"/>`,
      color ? `<color rgb="FF${color.slice(1)}"/>` : '<color theme="1"/>',
      `<name val="${escapeXml(name)}"/>`,
      '<family val="2"/>',
      '</font>',
    ];
    const id = this.fonts.length;
    this.fonts.push(parts.join(''));
    this.fontKeys.set(key, id);
    return id;
  }

  private internFill(fill: string | undefined): number {
    const color = normalizeColor(fill);
    if (!color) return 0;
    const existing = this.fillKeys.get(color);
    if (existing !== undefined) return existing;
    const id = this.fills.length;
    this.fills.push(
      `<fill><patternFill patternType="solid"><fgColor rgb="FF${color.slice(1)}"/><bgColor indexed="64"/></patternFill></fill>`,
    );
    this.fillKeys.set(color, id);
    return id;
  }

  private internBorder(border: SynthStyle['border']): number {
    if (!border) return 0;
    const sides: Array<'left' | 'right' | 'top' | 'bottom'> = ['left', 'right', 'top', 'bottom'];
    if (!sides.some((side) => border[side]?.style)) return 0;
    const key = sides
      .map((side) => `${side}:${border[side]?.style ?? ''}:${normalizeColor(border[side]?.color) ?? ''}`)
      .join('|');
    const existing = this.borderKeys.get(key);
    if (existing !== undefined) return existing;
    const body = sides
      .map((side) => {
        const line = border[side];
        if (!line?.style) return `<${side}/>`;
        const style = BORDER_STYLE_MAP[line.style] ?? 'thin';
        const color = normalizeColor(line.color);
        const colorXml = color ? `<color rgb="FF${color.slice(1)}"/>` : '<color auto="1"/>';
        return `<${side} style="${style}">${colorXml}</${side}>`;
      })
      .join('');
    const id = this.borders.length;
    this.borders.push(`<border>${body}<diagonal/></border>`);
    this.borderKeys.set(key, id);
    return id;
  }

  /** 数字格式：内置 id 直接用（0 = General），自定义 id 从 164 起（规范约定） */
  private internNumFmt(pattern: string | undefined): { numFmtId: number; custom: boolean } {
    if (!pattern || pattern === 'General') return { numFmtId: 0, custom: false };
    const builtin = BUILTIN_BY_PATTERN.get(pattern);
    if (builtin !== undefined) return { numFmtId: builtin, custom: false };
    const existing = this.numFmts.get(pattern);
    if (existing !== undefined) return { numFmtId: existing, custom: true };
    const id = 164 + this.numFmts.size;
    this.numFmts.set(pattern, id);
    return { numFmtId: id, custom: true };
  }

  toXml(): string {
    const numFmtXml = [...this.numFmts.entries()]
      .map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`)
      .join('');
    const xfXml = this.xfs
      .map((xf) => {
        const attrs = [
          `numFmtId="${xf.numFmtId}"`,
          `fontId="${xf.fontId}"`,
          `fillId="${xf.fillId}"`,
          `borderId="${xf.borderId}"`,
          'xfId="0"',
          xf.applyNumberFormat ? 'applyNumberFormat="1"' : '',
          xf.applyFont ? 'applyFont="1"' : '',
          xf.applyFill ? 'applyFill="1"' : '',
          xf.applyBorder ? 'applyBorder="1"' : '',
          xf.applyAlignment ? 'applyAlignment="1"' : '',
        ]
          .filter(Boolean)
          .join(' ');
        const alignment: string[] = [];
        if (xf.horizontal) alignment.push(`horizontal="${xf.horizontal}"`);
        if (xf.vertical) alignment.push(`vertical="${xf.vertical}"`);
        if (xf.wrapText) alignment.push('wrapText="1"');
        if (xf.textRotation !== undefined) alignment.push(`textRotation="${xf.textRotation}"`);
        const body = alignment.length > 0 ? `<alignment ${alignment.join(' ')}/>` : '';
        return `<xf ${attrs}>${body}</xf>`;
      })
      .join('');
    return [
      XML_HEAD,
      `<styleSheet xmlns="${NAMESPACE_MAIN}">`,
      numFmtXml ? `<numFmts count="${this.numFmts.size}">${numFmtXml}</numFmts>` : '',
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`,
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`,
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`,
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
      `<cellXfs count="${this.xfs.length}">${xfXml}</cellXfs>`,
      '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>',
      '</styleSheet>',
    ].join('');
  }
}

/** 内置数字格式（id → 格式串）：反向表用于"能被内置 id 表达的就别写自定义格式" */
const BUILTIN_BY_PATTERN = new Map<string, number>([
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['# ?/?', 12],
  ['# ??/??', 13],
  ['m/d/yy', 14],
  ['d-mmm-yy', 15],
  ['d-mmm', 16],
  ['mmm-yy', 17],
  ['h:mm AM/PM', 18],
  ['h:mm:ss AM/PM', 19],
  ['h:mm', 20],
  ['h:mm:ss', 21],
  ['m/d/yy h:mm', 22],
  ['#,##0 ;(#,##0)', 37],
  ['#,##0 ;[Red](#,##0)', 38],
  ['#,##0.00;(#,##0.00)', 39],
  ['#,##0.00;[Red](#,##0.00)', 40],
  ['mm:ss', 45],
  ['[h]:mm:ss', 46],
  ['mm:ss.0', 47],
  ['@', 49],
]);

/* -------------------------------------------------------------------------- */
/* 工作表 XML                                                                  */
/* -------------------------------------------------------------------------- */

function cellXml(cell: SynthCell, styleId: number): string {
  const ref = cellRef(cell.row, cell.col);
  const styleAttr = styleId > 0 ? ` s="${styleId}"` : '';
  const formula = cell.formula?.replace(/^=/, '');
  const value = cell.value;

  if (formula) {
    const cached =
      typeof value === 'number' ? `<v>${value}</v>`
        : typeof value === 'boolean' ? `<v>${value ? 1 : 0}</v>`
          : typeof value === 'string' && value !== '' ? `<v>${escapeXml(value)}</v>`
            : '';
    const typeAttr = typeof value === 'string' && value !== '' ? ' t="str"' : '';
    return `<c r="${ref}"${styleAttr}${typeAttr}><f>${escapeXml(formula)}</f>${cached}</c>`;
  }

  if (value === null || value === undefined || value === '') return styleAttr ? `<c r="${ref}"${styleAttr}/>` : '';
  if (typeof value === 'number') {
    return Number.isFinite(value) ? `<c r="${ref}"${styleAttr}><v>${value}</v></c>` : `<c r="${ref}"${styleAttr}/>`;
  }
  if (typeof value === 'boolean') return `<c r="${ref}"${styleAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(value))}</t></is></c>`;
}

function sheetXml(sheet: SynthSheet, styleIds: number[]): string {
  // 行优先分组（同一行的单元格必须挨在一起，且行号升序）
  const byRow = new Map<number, SynthCell[]>();
  let maxRow = 0;
  let maxCol = 0;
  for (const cell of sheet.cells) {
    if (cell.row < 0 || cell.col < 0) continue;
    const list = byRow.get(cell.row);
    if (list) list.push(cell);
    else byRow.set(cell.row, [cell]);
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.col > maxCol) maxCol = cell.col;
  }

  const rowsXml: string[] = [];
  for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
    const cells = byRow.get(row)!.slice().sort((a, b) => a.col - b.col);
    const height = sheet.rowHeights?.[row];
    const rowAttrs = [`r="${row + 1}"`];
    if (typeof height === 'number' && height > 0) rowAttrs.push(`ht="${Math.round(height * 100) / 100}"`, 'customHeight="1"');
    rowsXml.push(`<row ${rowAttrs.join(' ')}>${cells.map((cell) => cellXml(cell, styleIds[cell.style ?? -1] ?? 0)).join('')}</row>`);
  }

  const colsXml =
    sheet.colWidths && Object.keys(sheet.colWidths).length > 0
      ? `<cols>${Object.entries(sheet.colWidths)
          .map(([col, width]) => ({ col: Number(col), width }))
          .filter((entry) => Number.isFinite(entry.col) && entry.col >= 0 && entry.width > 0)
          .sort((a, b) => a.col - b.col)
          .map(
            (entry) =>
              `<col min="${entry.col + 1}" max="${entry.col + 1}" width="${Math.round(entry.width * 100) / 100}" customWidth="1"/>`,
          )
          .join('')}</cols>`
      : '';

  const mergeXml =
    sheet.merges && sheet.merges.length > 0
      ? `<mergeCells count="${sheet.merges.length}">${sheet.merges
          .map((ref) => `<mergeCell ref="${escapeXml(ref)}"/>`)
          .join('')}</mergeCells>`
      : '';

  const viewAttrs = sheet.gridlinesHidden ? ' showGridLines="0"' : '';
  const freeze = sheet.freeze && (sheet.freeze.rows > 0 || sheet.freeze.cols > 0) ? sheet.freeze : null;
  const paneXml = freeze
    ? (() => {
        const parts: string[] = [];
        if (freeze.cols > 0) parts.push(`xSplit="${freeze.cols}"`);
        if (freeze.rows > 0) parts.push(`ySplit="${freeze.rows}"`);
        const topLeft = `${colLetter(freeze.cols)}${freeze.rows + 1}`;
        const activePane = freeze.cols > 0 && freeze.rows > 0 ? 'bottomRight' : freeze.cols > 0 ? 'topRight' : 'bottomLeft';
        parts.push(`topLeftCell="${topLeft}"`, `activePane="${activePane}"`, 'state="frozen"');
        return `<pane ${parts.join(' ')}/>`;
      })()
    : '';

  const sheetFormatPr = [
    typeof sheet.defaultRowHeight === 'number' && sheet.defaultRowHeight > 0 ? `defaultRowHeight="${sheet.defaultRowHeight}"` : 'defaultRowHeight="15"',
    typeof sheet.defaultColWidth === 'number' && sheet.defaultColWidth > 0 ? `defaultColWidth="${sheet.defaultColWidth}"` : '',
    'outlineLevelRow="0" outlineLevelCol="0"',
  ]
    .filter(Boolean)
    .join(' ');

  const dimension = `A1:${cellRef(Math.max(maxRow, 0), Math.max(maxCol, 0))}`;

  return [
    XML_HEAD,
    `<worksheet xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_REL}">`,
    `<dimension ref="${dimension}"/>`,
    `<sheetViews><sheetView workbookViewId="0"${viewAttrs}>${paneXml}</sheetView></sheetViews>`,
    `<sheetFormatPr ${sheetFormatPr}/>`,
    colsXml,
    `<sheetData>${rowsXml.join('')}</sheetData>`,
    mergeXml,
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    '</worksheet>',
  ].join('');
}

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 把中性的 `WorkbookInput` 打成一份**合法的最小 xlsx**（zip 字节）。
 *
 * 返回的字节会：① 立刻被 `parseXlsx` 解析成模型；② 作为该标签页的"原始字节"用于
 * 外科式导出；③ 存进会话以便恢复。三者都要求它是一份规范文件，所以这里宁可少写特性、
 * 也不写任何"我们自己的解析器能容忍、Excel 却打不开"的东西。
 */
export function writeWorkbookPackage(input: WorkbookInput): Uint8Array {
  const table = new StyleTable();
  const styleIds = (input.styles ?? []).map((style) => table.intern(style));

  const used = new Set<string>();
  const sheets = input.sheets.length > 0 ? input.sheets : [{ name: '工作表', cells: [] }];
  const sheetNames = sheets.map((sheet, index) => sanitizeSheetName(sheet.name, used, `工作表${index + 1}`));

  const files: Zippable = {};

  const sheetOverrides = sheetNames
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('');

  files['[Content_Types].xml'] = strToU8(
    [
      XML_HEAD,
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      sheetOverrides,
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
      '</Types>',
    ].join(''),
  );

  files['_rels/.rels'] = strToU8(
    [
      XML_HEAD,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      `<Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="xl/workbook.xml"/>`,
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
      '</Relationships>',
    ].join(''),
  );

  const sheetEntries = sheetNames
    .map((name, index) => {
      const state = sheets[index]?.hidden ? ' state="hidden"' : '';
      return `<sheet name="${escapeXml(name)}" sheetId="${index + 1}"${state} r:id="rId${index + 1}"/>`;
    })
    .join('');

  files['xl/workbook.xml'] = strToU8(
    [
      XML_HEAD,
      `<workbook xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_REL}">`,
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="12000"/></bookViews>',
      `<sheets>${sheetEntries}</sheets>`,
      // 让 Excel 打开时重算公式（我们写的是"缓存值 + 公式"，重算能保证与源文件一致）
      '<calcPr calcId="191029" fullCalcOnLoad="1"/>',
      '</workbook>',
    ].join(''),
  );

  const workbookRels = sheetNames
    .map((_, index) => `<Relationship Id="rId${index + 1}" Type="${REL_WORKSHEET}" Target="worksheets/sheet${index + 1}.xml"/>`)
    .join('');
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    [
      XML_HEAD,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      workbookRels,
      `<Relationship Id="rId${sheetNames.length + 1}" Type="${REL_STYLES}" Target="styles.xml"/>`,
      '</Relationships>',
    ].join(''),
  );

  files['xl/styles.xml'] = strToU8(table.toXml());

  sheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(sheetXml(sheet, styleIds));
  });

  const workbookName = escapeXml(input.name ?? '工作簿');
  files['docProps/core.xml'] = strToU8(
    [
      XML_HEAD,
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
      `<dc:title>${workbookName}</dc:title>`,
      '<cp:lastModifiedBy>Excel 预览工具</cp:lastModifiedBy>',
      '</cp:coreProperties>',
    ].join(''),
  );
  files['docProps/app.xml'] = strToU8(
    [
      XML_HEAD,
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
      '<Application>Excel 预览工具</Application>',
      '</Properties>',
    ].join(''),
  );

  return zipSync(files, { level: 6 });
}
