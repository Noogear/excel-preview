/**
 * OOXML(.xlsx) 解析器的中立数据模型。
 *
 * 这里是解析层与上层适配层（ParsedWorkbook -> Univer IWorkbookData）之间的契约，
 * 上层只依赖本文件的类型，不依赖任何具体解析实现。
 */

export type BorderLineStyle =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'dashDot' | 'dashDotDot' | 'mediumDashed' | 'mediumDashDot' | 'mediumDashDotDot' | 'slantDashDot';

export interface ParsedBorder { style: BorderLineStyle; color?: string }

export interface ParsedStyle {
  fontFamily?: string; fontSize?: number; bold?: boolean; italic?: boolean;
  underline?: boolean; strikeThrough?: boolean; color?: string;
  fill?: string;
  border?: { top?: ParsedBorder; bottom?: ParsedBorder; left?: ParsedBorder; right?: ParsedBorder };
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  textWrap?: boolean; indent?: number; textRotation?: number;
  numberFormat?: string;
}

export type ParsedCellValue = string | number | boolean | null;

export interface ParsedCell {
  row: number; col: number;
  value?: ParsedCellValue;
  formula?: string;        // 不含前导 '='，如 "SUM(A1:A3)"
  styleIndex?: number;
  error?: string;          // t="e" 时的错误文本，如 "#DIV/0!"
}

export interface ParsedRowInfo { height?: number; hidden?: boolean; customHeight?: boolean }
export interface ParsedColInfo { width?: number; hidden?: boolean; customWidth?: boolean }
export interface ParsedMerge { startRow: number; startCol: number; endRow: number; endCol: number }
export interface ParsedFreeze { row: number; col: number }

/* ============================================================================
 * 附属特性：条件格式 / 数据验证 / 超链接 / 批注 / 表格 / 图片
 * 字段名与 OOXML 一一对应，便于上层映射，也便于对外解释"哪些没还原"
 * ========================================================================== */

export type CfRuleKind = 'highlight' | 'colorScale' | 'dataBar' | 'iconSet';

/** 条件格式阈值（cfvo） */
export interface CfValueObject {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax';
  value?: string;
}

/** differential formatting：条件格式命中的单元格样式（styles.xml 的 <dxfs>） */
export interface CfDxfStyle {
  bold?: boolean;
  italic?: boolean;
  strikeThrough?: boolean;
  color?: string;
  fill?: string;
  border?: { top?: ParsedBorder; bottom?: ParsedBorder; left?: ParsedBorder; right?: ParsedBorder };
}

export interface ConditionalFormatRule {
  kind: CfRuleKind;
  /** sqref 展开后的区域（A1 记号，可能多个） */
  ranges: string[];
  priority: number;
  /** highlight：cellIs / containsText / duplicateValues / expression … */
  ruleType?: string;
  operator?: string;
  text?: string;
  formula1?: string;
  formula2?: string;
  dxfId?: number;
  dxf?: CfDxfStyle;
  cfvo?: CfValueObject[];
  /** colorScale 的颜色序列（从低到高） */
  colors?: string[];
  /** dataBar 颜色 */
  color?: string;
  /** iconSet 名称，如 3TrafficLights1 */
  iconSet?: string;
  showValue?: boolean;
}

export interface DataValidationRule {
  /** sqref 展开后的区域（A1 记号） */
  ranges: string[];
  /** list / whole / decimal / date / time / textLength / custom / none */
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
}

export interface ParsedHyperlink {
  /** 单元格引用，如 B1 */
  ref: string;
  /** 外部链接目标（来自 sheet rels） */
  target?: string;
  /** 工作表内位置（location 属性，如 "条件格式!A1"） */
  location?: string;
  display?: string;
  tooltip?: string;
}

/** 批注（legacy note，来自 xl/comments*.xml） */
export interface ParsedNote {
  ref: string;
  text: string;
  author?: string;
}

/** Excel 表格（ListObject，来自 xl/tables/table*.xml） */
export interface ParsedTable {
  name: string;
  displayName?: string;
  /** 表格范围（A1 记号，含表头与汇总行） */
  ref: string;
  headerRowCount: number;
  totalsRowCount: number;
  /** `totalsRowShown` 原样保留（**注意**：ExcelJS 把"无汇总行"写成 `totalsRowShown="1"`，语义与字面相反，不要直接当"有汇总行"用） */
  totalsRowShown?: boolean;
  /** 表格样式名，如 TableStyleMedium9 */
  styleName?: string;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  columns?: Array<{ name: string }>;
}

/** 浮动图片（来自 xl/drawings/drawing*.xml）
 *  锚点偏移单位是 EMU（1 英寸 = 914400 EMU）；col/row 是 0-based 单元格索引 */
export interface ParsedImage {
  id: string;
  /** zip 内的媒体路径，如 xl/media/image1.png */
  mediaPath: string;
  anchorType: 'oneCell' | 'twoCell' | 'absolute';
  from: { col: number; row: number; colOffEmu?: number; rowOffEmu?: number };
  to?: { col: number; row: number; colOffEmu?: number; rowOffEmu?: number };
  /** oneCell 锚点用；EMU */
  extEmu?: { cx: number; cy: number };
  /** absoluteAnchor 的 `xdr:pos`（EMU，相对工作表左上角）；oneCell/twoCell 锚点没有该值 */
  posEmu?: { x: number; y: number };
}

/**
 * 打印设置（`pageSetup` / `pageMargins` / `headerFooter` / `printOptions` / 分页符）。
 *
 * 为什么它不再算"未支持"：这些字段**只影响打印**，对屏幕预览与内容编辑毫无影响，
 * 而导出是"在原始 XML 上打补丁"的字节级保真路线 —— 它们本来就原样保留着。
 * 以前把它们记进 `unsupported`，用户导入自己的座位表时会看到"未支持 1 项"，
 * 点开却是"打印设置"，属于**误导**（实测反馈）。现在改成解析出真值 + 归入
 * `report.preserved`（已原样保留，不影响预览）。
 */
export interface ParsedPrintSettings {
  /** OOXML paperSize（9 = A4、8 = A3、1 = Letter…） */
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  /** 缩放百分比（与 fitToPage 互斥） */
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  /** 单位：英寸（与文件里一致，不换算） */
  margins?: { left: number; right: number; top: number; bottom: number; header?: number; footer?: number };
  headerFooter?: { oddHeader?: string; oddFooter?: string; evenHeader?: string; evenFooter?: string };
  centered?: boolean;
  /** 手工分页符处数（行/列） */
  breaks?: { row: number; col: number };
}

export interface ParsedSheet {
  id: string; name: string; index: number; hidden?: boolean; veryHidden?: boolean;
  rows: Record<number, ParsedRowInfo>;
  cols: Record<number, ParsedColInfo>;
  cells: ParsedCell[];
  merges: ParsedMerge[];
  freeze?: ParsedFreeze;
  gridlinesHidden?: boolean;
  defaultRowHeight?: number; defaultColWidth?: number;
  dimension?: { startRow: number; startCol: number; endRow: number; endCol: number };
  /* ---- 附属特性 ---- */
  conditionalFormats?: ConditionalFormatRule[];
  dataValidations?: DataValidationRule[];
  hyperlinks?: ParsedHyperlink[];
  notes?: ParsedNote[];
  tables?: ParsedTable[];
  images?: ParsedImage[];
  print?: ParsedPrintSettings;
}

export interface ParsedReport {
  unsupported: string[];
  warnings: string[];
  /**
   * "已解析并原样保留、但不影响预览"的信息（目前是打印设置）。
   * 与 `unsupported` 的区别：这里**不是**缺陷，状态栏不该因此报警。
   */
  preserved?: string[];
}

export interface ParsedWorkbook {
  sheets: ParsedSheet[];
  styles: ParsedStyle[];
  /** 条件格式命中的差异样式（styles.xml 的 <dxfs>，索引对应 cfRule 的 dxfId） */
  dxfStyles?: CfDxfStyle[];
  themeColors?: string[];
  report: ParsedReport;
  raw: { entries: Record<string, Uint8Array> };
}
