/**
 * 编排层：把一个 xlsx（内存字节）解析成中立的 `ParsedWorkbook`。
 *
 * 解析链路：
 *   bytes -> zip 条目
 *        -> `_rels/.rels` 找 workbook part（不假设固定叫 `xl/workbook.xml`）
 *        -> `xl/_rels/workbook.xml.rels` + `xl/workbook.xml` 得到工作表顺序/target
 *        -> sharedStrings / theme / styles（含 `<dxfs>` 条件格式差异样式）/ 每个工作表
 *        -> 每个工作表再通过 `xl/worksheets/_rels/sheetN.xml.rels` 找到它自己的
 *           批注(comments) / 表格(table) / 绘图(drawing) / 超链接(hyperlink) 部件
 *
 * 关键约定：**不假设任何文件名编号与 sheet 编号一致**——sheet1.xml 可以引用
 * comments3.xml / tables/table7.xml / drawings/drawing2.xml，一切都从 rels 来。
 *
 * 容错原则：缺部件用默认值，坏部件只记 warning，绝不因为局部问题整体抛异常；
 * 一个部件解析失败不影响其它部件（每个部件各自 try/catch）。
 * 只有"根本不是 zip"或"zip 里没有任何可用部件"才返回空工作簿 + warning。
 */
import type {
  CfDxfStyle, ParsedImage, ParsedNote, ParsedReport, ParsedSheet, ParsedStyle, ParsedTable, ParsedWorkbook,
} from './types';
import { dirName, readText, readZipEntries, resolvePath, type ZipEntries } from './zip';
import { childElements, earliestNameAt, findFirstElement, localName } from './xml';
import { parseDxfStyles, parseStyles, parseThemeColors } from './styles';
import { parseSharedStrings, parseSheetPartRefs, parseWorksheet, type SheetPartRefs } from './worksheet';
import { parseConditionalFormatting } from './conditional-formatting';
import { parseDataValidations } from './data-validation';
import { parseComments } from './comments';
import { parseTable } from './table';
import { parseDrawing } from './drawing';

export * from './types';
export { readZipEntries, readText as readEntryText, readBytes as readEntryBytes, decodeText } from './zip';
export {
  parseStyles, parseDxfStyles, parseThemeColors, applyTint, parseHexColor,
  BUILTIN_NUM_FMTS, INDEXED_COLORS,
} from './styles';
export {
  parseWorksheet, parseSharedStrings, parseCellRef, parseRangeRef,
  parseSheetPartRefs, expandSqref, splitSqref, isA1RangeToken,
} from './worksheet';
export { parseConditionalFormatting } from './conditional-formatting';
export { parseDataValidations } from './data-validation';
export { parseComments } from './comments';
export { parseTable } from './table';
export { parseDrawing } from './drawing';
export {
  visitElements, findElements, findFirstElement, childElements,
  elementText, decodeEntities, localName, nameMatches,
} from './xml';
export type { ElementRange, XmlAttributes, VisitOptions, ElementVisitor } from './xml';
export type { SheetPartRefs, RawHyperlink, Warn } from './worksheet';
export type { ParseConditionalFormattingOptions } from './conditional-formatting';
export type { ParseDataValidationsOptions } from './data-validation';
export type { ParseCommentsOptions } from './comments';
export type { ParseTableOptions } from './table';
export type { ParseDrawingOptions, ParseDrawingResult } from './drawing';

const DEFAULT_WORKBOOK_PATH = 'xl/workbook.xml';
const DEFAULT_WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels';
const DEFAULT_STYLES_PATH = 'xl/styles.xml';
const DEFAULT_SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';
const DEFAULT_THEME_PATH = 'xl/theme/theme1.xml';

/** 主入口：字节进，`ParsedWorkbook` 出（不抛异常，问题都进 `report`） */
export async function parseXlsx(data: Uint8Array | ArrayBuffer): Promise<ParsedWorkbook> {
  const report: ParsedReport = { unsupported: [], warnings: [], preserved: [] };
  const warn = (msg: string): void => { report.warnings.push(msg); };

  let entries: ZipEntries;
  try {
    entries = readZipEntries(data);
  } catch (err) {
    warn(`无法解压 xlsx（可能不是合法 zip）：${errText(err)}`);
    return { sheets: [], styles: [], report, raw: { entries: {} } };
  }

  if (Object.keys(entries).length === 0) {
    warn('zip 内没有任何条目');
    return { sheets: [], styles: [], report, raw: { entries } };
  }

  // 1) 定位 workbook part（先信 _rels/.rels，失败再退回约定路径）
  const relsRoot = readText(entries, '_rels/.rels');
  let workbookPath = DEFAULT_WORKBOOK_PATH;
  if (relsRoot) {
    const found = findWorkbookPath(relsRoot);
    if (found) workbookPath = found;
    else warn('_rels/.rels 中没有 officeDocument 关系，回退到 xl/workbook.xml');
  } else {
    warn('缺少 _rels/.rels，回退到 xl/workbook.xml');
  }

  const workbookXml = readText(entries, workbookPath);
  if (!workbookXml) {
    warn(`找不到工作簿部件 ${workbookPath}，无法解析工作表清单`);
    return { sheets: [], styles: [], report, raw: { entries } };
  }

  // 2) 工作表清单：名字/顺序/隐藏状态 + rels 里的 target
  const sheetRefs = parseWorkbookSheets(workbookXml, warn);
  const relsPath = `${dirName(workbookPath)}/_rels/${baseName(workbookPath)}.rels`;
  const workbookRels = readText(entries, relsPath)
    ?? readText(entries, DEFAULT_WORKBOOK_RELS_PATH);
  if (!workbookRels) warn(`缺少 ${relsPath}，工作表将按约定路径 xl/worksheets/sheetN.xml 尝试加载`);
  const relTargets = workbookRels ? parseRelationships(workbookRels, dirName(workbookPath)) : new Map<string, string>();

  // 3) 共享字符串 / 主题 / 样式（styles.xml 只读一次，样式表与 dxfs 都从它出）
  const sharedStrings = loadSharedStrings(entries, warn);
  const themeColors = loadThemeColors(entries, workbookPath, warn);
  const stylesPart = readStylesPart(entries, workbookPath, warn);
  const styles = stylesPart ? safeParseStyles(stylesPart.xml, themeColors, warn) : [];
  const dxfStyles = stylesPart ? safeParseDxfStyles(stylesPart.xml, themeColors, warn) : [];

  // 4) 逐个工作表（含它自己的批注/表格/绘图部件）
  const sheets: ParsedSheet[] = [];
  sheetRefs.forEach((ref, index) => {
    const target = relTargets.get(ref.relId) ?? fallbackSheetTarget(index);
    if (!relTargets.has(ref.relId)) {
      warn(`工作表 "${ref.name}" 的关系 ${ref.relId} 在 workbook.xml.rels 中缺失，按 ${target} 猜测`);
    }
    /** 本表所有降级信息都带上工作表名，便于定位 */
    const sheetWarn = (msg: string): void => { report.warnings.push(`[${ref.name}] ${msg}`); };
    const unsupported = (msg: string): void => { report.unsupported.push(`[${ref.name}] ${msg}`); };

    const sheetXml = readText(entries, target);
    if (!sheetXml) {
      warn(`工作表 "${ref.name}" 的部件 ${target} 不存在，已按空表处理`);
      sheets.push({
        id: ref.relId, name: ref.name, index,
        ...(ref.hidden ? { hidden: true } : {}),
        ...(ref.veryHidden ? { veryHidden: true } : {}),
        rows: {}, cols: {}, cells: [], merges: [],
      });
      return;
    }

    // 4a) 本表引用（超链接 / 表格 / 绘图）与 sheet rels —— 文件名编号与 sheet 编号无关
    //
    // 尾段起点：合并区/超链接/条件格式/数据验证/表格/绘图等一律排在 `<sheetData>` 之后。
    // 只算一次 `</sheetData>`（这一步本身要扫完整张表数据），然后各特性都从尾段起步，
    // 免去"每个特性各走一遍整张表"。`earliestNameAt` 会在元素实际出现在更前面时自动前移起点，
    // 所以不依赖 schema 顺序也能保证不漏。
    const sheetTail = findFirstElement(sheetXml, 'sheetData')?.closeEnd ?? 0;
    const refs = parseSheetPartRefs(sheetXml, earliestNameAt(sheetXml, 'hyperlinks', sheetTail));
    const sheetRelsPart = `${dirName(target)}/_rels/${baseName(target)}.rels`;
    const sheetRelsXml = readText(entries, sheetRelsPart);
    const sheetRels = sheetRelsXml ? parseSheetRelations(sheetRelsXml, dirName(target)) : [];
    if (!sheetRelsXml && hasPartRefs(refs)) {
      sheetWarn(`引用了外部部件（超链接/表格/绘图）但缺少 ${sheetRelsPart}，相关目标无法解析`);
    }
    for (const kind of collectRelKinds(sheetRels)) {
      const label = SHEET_REL_LABELS[kind.kind];
      if (label) unsupported(`${label} · ${kind.count} 处（未解析）`);
    }

    // 4b) 各部件分别解析：任何一个坏掉都不影响其它
    const notes = loadComments(entries, sheetRels, sheetWarn);
    const tables = loadTables(entries, refs, sheetRels, sheetWarn);
    const drawing = loadDrawing(entries, refs, sheetRels, sheetWarn);
    if (drawing) for (const u of drawing.unsupported) unsupported(u);

    const hyperlinkTargets = new Map<string, string>();
    for (const rel of sheetRels) {
      if (rel.kind === 'hyperlink') hyperlinkTargets.set(rel.id, rel.target);
    }

    // 4c) 工作表本体（单元格 / 合并 / 冻结 / 超链接 / 注入的批注·表格·图片）
    const result = parseWorksheet({
      id: ref.relId,
      name: ref.name,
      index,
      ...(ref.hidden ? { hidden: true } : {}),
      ...(ref.veryHidden ? { veryHidden: true } : {}),
      xml: sheetXml,
      sharedStrings,
      refs,
      hyperlinkTargets,
      parts: {
        ...(notes ? { notes } : {}),
        ...(tables ? { tables } : {}),
        ...(drawing && drawing.images.length > 0 ? { images: drawing.images } : {}),
      },
    });
    for (const w of result.diagnostics.warnings) report.warnings.push(`[${ref.name}] ${w}`);
    for (const u of result.diagnostics.unsupported) report.unsupported.push(`[${ref.name}] ${u}`);
    for (const p of result.diagnostics.preserved) (report.preserved ??= []).push(`[${ref.name}] ${p}`);

    // 4d) 条件格式 / 数据验证（纯工作表 XML 解析，dxf 已就绪）
    try {
      const rules = parseConditionalFormatting(sheetXml, {
        warn: sheetWarn,
        dxfStyles,
        from: earliestNameAt(sheetXml, 'conditionalFormatting', sheetTail),
        ...(themeColors ? { themeColors } : {}),
      });
      if (rules.length > 0) result.sheet.conditionalFormats = rules;
    } catch (err) {
      sheetWarn(`条件格式解析失败，已忽略：${errText(err)}`);
    }
    try {
      const validations = parseDataValidations(sheetXml, {
        warn: sheetWarn,
        from: earliestNameAt(sheetXml, 'dataValidations', sheetTail),
      });
      if (validations.length > 0) result.sheet.dataValidations = validations;
    } catch (err) {
      sheetWarn(`数据验证解析失败，已忽略：${errText(err)}`);
    }

    sheets.push(result.sheet);
  });

  if (sheets.length === 0) warn('工作簿里没有解析出任何工作表');

  // 5) 工作簿级部件（宏 / 图表工作表 / 表单控件 …）——这些没有工作表归属
  const workbookParts = collectWorkbookParts(entries);
  for (const [label, count] of workbookParts.unsupported) {
    report.unsupported.push(`[工作簿] ${label} · ${count} 个部件（未解析）`);
  }
  for (const [label, count] of workbookParts.preserved) {
    (report.preserved ??= []).push(
      `[工作簿] ${label} · ${count} 个部件已原样保留（导出时逐字节不变；本工具只预览与编辑单元格内容，不执行宏）`,
    );
  }

  validateStyleIndices(sheets, styles, warn);
  dedupeWarnings(report);

  return {
    sheets,
    styles,
    ...(dxfStyles.length > 0 ? { dxfStyles } : {}),
    ...(themeColors ? { themeColors } : {}),
    report,
    raw: { entries },
  };
}

/* -------------------------------------------------------------------------- */
/* 工作表关系（xl/worksheets/_rels/sheetN.xml.rels）                            */
/* -------------------------------------------------------------------------- */

interface SheetRel {
  id: string;
  /** Type 的末段，如 comments / table / drawing / hyperlink / pivotTable */
  kind: string;
  /** 内部部件为 zip 内绝对路径；外部（TargetMode="External"）原样保留 */
  target: string;
  external: boolean;
}

/** 仍不解析的**工作表级关系**（按种类计数） */
const SHEET_REL_LABELS: Readonly<Record<string, string>> = {
  pivotTable: '数据透视表(pivotTable)',
  slicer: '切片器(slicer)',
  timeline: '日程表(timeline)',
  threadedComment: '线程化批注(threadedComment)',
};

/**
 * 解析 sheet rels。内部部件按 `baseDir` 归一化成 zip 条目名
 * （`../comments1.xml` + `xl/worksheets` -> `xl/comments1.xml`），
 * 外部目标（超链接、外部图片）**原样保留**，否则 URL 会被当路径改写。
 */
function parseSheetRelations(relsXml: string, baseDir: string): SheetRel[] {
  const out: SheetRel[] = [];
  const root = findFirstElement(relsXml, 'Relationships');
  if (!root) return out;
  for (const el of childElements(relsXml, root)) {
    if (localName(el.name) !== 'Relationship') continue;
    const id = el.attrs['Id'];
    const target = el.attrs['Target'];
    if (!id || !target) continue;
    const external = el.attrs['TargetMode'] === 'External';
    out.push({
      id,
      kind: relKind(el.attrs['Type'] ?? ''),
      target: external ? target : resolvePath(baseDir, target),
      external,
    });
  }
  return out;
}

/** `…/relationships/comments` -> `comments` */
function relKind(type: string): string {
  const slash = type.lastIndexOf('/');
  return slash < 0 ? type : type.slice(slash + 1);
}

function collectRelKinds(rels: readonly SheetRel[]): Array<{ kind: string; count: number }> {
  const counts = new Map<string, number>();
  for (const rel of rels) counts.set(rel.kind, (counts.get(rel.kind) ?? 0) + 1);
  return [...counts.entries()].map(([kind, count]) => ({ kind, count }));
}

function relsOfKind(rels: readonly SheetRel[], kind: string): SheetRel[] {
  return rels.filter((rel) => rel.kind === kind);
}

function hasPartRefs(refs: SheetPartRefs): boolean {
  return refs.hasHyperlinks || refs.tableRelIds.length > 0 || refs.drawingRelId !== undefined;
}

/* -------------------------------------------------------------------------- */
/* 各工作表部件装载（失败只记 warning）                                          */
/* -------------------------------------------------------------------------- */

function loadComments(
  entries: ZipEntries,
  sheetRels: readonly SheetRel[],
  sheetWarn: (msg: string) => void,
): ParsedNote[] | undefined {
  const rel = relsOfKind(sheetRels, 'comments')[0];
  if (!rel) return undefined;
  const xml = readText(entries, rel.target);
  if (xml === undefined) {
    sheetWarn(`批注部件 ${rel.target} 不存在，已忽略批注`);
    return undefined;
  }
  try {
    const notes = parseComments(xml, { warn: sheetWarn, partPath: rel.target });
    return notes.length > 0 ? notes : undefined;
  } catch (err) {
    sheetWarn(`批注部件 ${rel.target} 解析失败，已忽略：${errText(err)}`);
    return undefined;
  }
}

function loadTables(
  entries: ZipEntries,
  refs: SheetPartRefs,
  sheetRels: readonly SheetRel[],
  sheetWarn: (msg: string) => void,
): ParsedTable[] | undefined {
  const tableRels = relsOfKind(sheetRels, 'table');
  if (tableRels.length === 0 && refs.tableRelIds.length === 0) return undefined;

  const byId = new Map(tableRels.map((rel) => [rel.id, rel]));
  // `<tableParts>` 的顺序就是 Excel 里的表格顺序
  const ordered: SheetRel[] = [];
  for (const relId of refs.tableRelIds) {
    const rel = byId.get(relId);
    if (rel) {
      ordered.push(rel);
      byId.delete(relId);
    } else {
      sheetWarn(`<tableParts> 引用的表格关系 ${relId} 在 sheet rels 里找不到，已跳过`);
    }
  }
  if (refs.tableRelIds.length === 0) ordered.push(...tableRels);
  else if (byId.size > 0) {
    sheetWarn(`有 ${byId.size} 个表格部件没有出现在 <tableParts> 里，已忽略`);
  }

  const tables: ParsedTable[] = [];
  for (const rel of ordered) {
    const xml = readText(entries, rel.target);
    if (xml === undefined) {
      sheetWarn(`表格部件 ${rel.target} 不存在，已忽略该表格`);
      continue;
    }
    try {
      const table = parseTable(xml, { warn: sheetWarn, partPath: rel.target });
      if (table) tables.push(table);
    } catch (err) {
      sheetWarn(`表格部件 ${rel.target} 解析失败，已忽略：${errText(err)}`);
    }
  }
  return tables.length > 0 ? tables : undefined;
}

function loadDrawing(
  entries: ZipEntries,
  refs: SheetPartRefs,
  sheetRels: readonly SheetRel[],
  sheetWarn: (msg: string) => void,
): { images: ParsedImage[]; unsupported: string[] } | undefined {
  const drawingRels = relsOfKind(sheetRels, 'drawing');
  // 优先用工作表 XML 的 <drawing r:id>，它才是这张表真正引用的那张绘图
  let rel = refs.drawingRelId !== undefined ? drawingRels.find((r) => r.id === refs.drawingRelId) : undefined;
  if (!rel && refs.drawingRelId !== undefined && drawingRels.length === 0) {
    sheetWarn(`<drawing> 引用的关系 ${refs.drawingRelId} 在 sheet rels 里找不到，图片无法解析`);
  }
  if (!rel) rel = drawingRels[0];
  if (!rel) return undefined;
  if (rel.external) {
    sheetWarn(`绘图部件 ${rel.target} 是外部引用，已忽略`);
    return undefined;
  }

  const xml = readText(entries, rel.target);
  if (xml === undefined) {
    sheetWarn(`绘图部件 ${rel.target} 不存在，已忽略图片`);
    return undefined;
  }
  const drawingRelsPath = `${dirName(rel.target)}/_rels/${baseName(rel.target)}.rels`;
  const drawingRelsXml = readText(entries, drawingRelsPath);
  if (!drawingRelsXml) {
    sheetWarn(`缺少 ${drawingRelsPath}，图片的媒体路径无法解析`);
  }
  try {
    return parseDrawing(xml, {
      ...(drawingRelsXml !== undefined ? { relsXml: drawingRelsXml } : {}),
      partPath: rel.target,
      warn: sheetWarn,
    });
  } catch (err) {
    sheetWarn(`绘图部件 ${rel.target} 解析失败，已忽略：${errText(err)}`);
    return undefined;
  }
}

/* -------------------------------------------------------------------------- */
/* 工作簿级"仍不支持"的部件                                                      */
/* -------------------------------------------------------------------------- */

const WORKBOOK_PART_LABELS: ReadonlyArray<readonly [RegExp, string]> = [
  [/^xl\/macrosheets\//i, '宏工作表(macrosheet)'],
  [/^xl\/dialogsheets\//i, '对话框工作表(dialogsheet)'],
  [/^xl\/chartsheets\//i, '图表工作表(chartsheet)'],
  [/^xl\/externalLinks\//i, '外部链接缓存(externalLink)'],
  [/^xl\/ctrlProps\//i, '表单控件(ctrlProps)'],
  [/^xl\/activeX\//i, 'ActiveX 控件(activeX)'],
];

/**
 * **原样保留、但我们不会去用**的工作簿级部件。
 *
 * 目前只有 VBA 宏（`.xlsm`/`.xltm` 里的 `xl/vbaProject.bin`）：
 * 导出是"在原字节上打补丁"，这个部件我们一个字节都不动 → 宏在导出文件里**完好无损**；
 * 浏览器里也不可能执行 VBA，所以把它列进"未支持"会让用户以为宏丢了（实测反馈过同类误报：
 * 打印设置、未被引用的占位填充）。这里改成如实说明：**保留但不执行**。
 */
const PRESERVED_PART_LABELS: ReadonlyArray<readonly [RegExp, string]> = [
  [/^xl\/vbaProject\.bin$/i, '宏(VBA)'],
];

function collectWorkbookParts(
  entries: ZipEntries,
): { unsupported: Array<[string, number]>; preserved: Array<[string, number]> } {
  const unsupported = new Map<string, number>();
  const preserved = new Map<string, number>();
  for (const path of Object.keys(entries)) {
    let matched = false;
    for (const [pattern, label] of PRESERVED_PART_LABELS) {
      if (pattern.test(path)) {
        preserved.set(label, (preserved.get(label) ?? 0) + 1);
        matched = true;
        break;
      }
    }
    if (matched) continue;
    for (const [pattern, label] of WORKBOOK_PART_LABELS) {
      if (pattern.test(path)) {
        unsupported.set(label, (unsupported.get(label) ?? 0) + 1);
        break;
      }
    }
  }
  return { unsupported: [...unsupported.entries()], preserved: [...preserved.entries()] };
}

/* -------------------------------------------------------------------------- */
/* 部件定位                                                                    */
/* -------------------------------------------------------------------------- */

interface SheetRef {
  relId: string;
  name: string;
  hidden?: boolean;
  veryHidden?: boolean;
}

/** `_rels/.rels` -> officeDocument 关系指向的 workbook part */
export function findWorkbookPath(relsXml: string): string | undefined {
  for (const rel of relationships(relsXml)) {
    const type = rel.type ?? '';
    if (type.endsWith('/officeDocument') || type.endsWith('officeDocument')) {
      return normalizeTarget(rel.target);
    }
  }
  return undefined;
}

interface RelInfo { id: string; type?: string; target: string; mode?: string }

function relationships(relsXml: string): RelInfo[] {
  const out: RelInfo[] = [];
  const root = findFirstElement(relsXml, 'Relationships');
  if (!root) return out;
  for (const el of childElements(relsXml, root)) {
    if (localName(el.name) !== 'Relationship') continue;
    const id = el.attrs['Id'];
    const target = el.attrs['Target'];
    if (!id || !target) continue;
    out.push({
      id,
      target,
      ...(el.attrs['Type'] !== undefined ? { type: el.attrs['Type'] } : {}),
      ...(el.attrs['TargetMode'] !== undefined ? { mode: el.attrs['TargetMode'] } : {}),
    });
  }
  return out;
}

/** `xl/_rels/workbook.xml.rels` -> relId 到**绝对条目名**的映射 */
export function parseRelationships(relsXml: string, baseDir: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const rel of relationships(relsXml)) {
    if (rel.mode === 'External') continue; // 外链不解析
    map.set(rel.id, resolvePath(baseDir, rel.target));
  }
  return map;
}

function normalizeTarget(target: string): string {
  let t = target.replace(/\\/g, '/');
  if (t.startsWith('/')) t = t.slice(1);
  return t;
}

/* -------------------------------------------------------------------------- */
/* workbook.xml                                                                */
/* -------------------------------------------------------------------------- */

export function parseWorkbookSheets(workbookXml: string, warn: (msg: string) => void): SheetRef[] {
  const sheetsEl = findFirstElement(workbookXml, 'sheets');
  if (!sheetsEl) {
    warn('workbook.xml 缺少 <sheets>，无法确定工作表清单');
    return [];
  }
  const out: SheetRef[] = [];
  for (const el of childElements(workbookXml, sheetsEl)) {
    if (localName(el.name) !== 'sheet') continue;
    const name = el.attrs['name'] ?? `Sheet${out.length + 1}`;
    const relId = el.attrs['r:id'] ?? el.attrs['id'];
    if (!relId) {
      warn(`工作表 "${name}" 缺少 r:id，已跳过`);
      continue;
    }
    const state = el.attrs['state'] ?? 'visible';
    out.push({
      relId,
      name,
      ...(state === 'hidden' || state === 'veryHidden' ? { hidden: true } : {}),
      ...(state === 'veryHidden' ? { veryHidden: true } : {}),
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* 支撑部件                                                                    */
/* -------------------------------------------------------------------------- */

function loadSharedStrings(entries: ZipEntries, warn: (msg: string) => void): string[] {
  const xml = readText(entries, DEFAULT_SHARED_STRINGS_PATH);
  if (!xml) {
    // 没有 sharedStrings 是合法的（所有字符串都内联或没有字符串）
    return [];
  }
  try {
    return parseSharedStrings(xml, warn);
  } catch (err) {
    warn(`sharedStrings 解析失败，已忽略：${errText(err)}`);
    return [];
  }
}

function loadThemeColors(
  entries: ZipEntries,
  workbookPath: string,
  warn: (msg: string) => void,
): string[] | undefined {
  const candidates = [
    resolvePath(dirName(workbookPath), 'theme/theme1.xml'),
    DEFAULT_THEME_PATH,
  ];
  for (const path of candidates) {
    const xml = readText(entries, path);
    if (!xml) continue;
    try {
      return parseThemeColors(xml, warn);
    } catch (err) {
      warn(`theme1.xml 解析失败，主题色回退为默认主题：${errText(err)}`);
      return undefined;
    }
  }
  warn('缺少 xl/theme/theme1.xml，theme 颜色将回退为 Excel 默认主题色');
  return undefined;
}

/** styles.xml 只读一次，`<cellXfs>` 与 `<dxfs>` 共用同一份 XML */
function readStylesPart(
  entries: ZipEntries,
  workbookPath: string,
  warn: (msg: string) => void,
): { path: string; xml: string } | undefined {
  const candidates = [
    resolvePath(dirName(workbookPath), 'styles.xml'),
    DEFAULT_STYLES_PATH,
  ];
  for (const path of candidates) {
    const xml = readText(entries, path);
    if (xml) return { path, xml };
  }
  warn('缺少 xl/styles.xml，所有单元格按默认样式处理');
  return undefined;
}

function safeParseStyles(
  stylesXml: string,
  themeColors: string[] | undefined,
  warn: (msg: string) => void,
): ParsedStyle[] {
  try {
    return parseStyles(stylesXml, { ...(themeColors ? { themeColors } : {}), warn });
  } catch (err) {
    warn(`styles.xml 解析失败，样式回退为空表（全部按默认样式渲染）：${errText(err)}`);
    return [];
  }
}

/** `<dxfs>` -> 条件格式差异样式；失败只降级 dxfs，不影响单元格样式 */
function safeParseDxfStyles(
  stylesXml: string,
  themeColors: string[] | undefined,
  warn: (msg: string) => void,
): CfDxfStyle[] {
  try {
    return parseDxfStyles(stylesXml, { ...(themeColors ? { themeColors } : {}), warn });
  } catch (err) {
    warn(`styles.xml 的 <dxfs> 解析失败，条件格式差异样式已忽略：${errText(err)}`);
    return [];
  }
}

/** 单元格引用了越界 styleIndex 时统一提醒（不逐格刷屏） */
function validateStyleIndices(sheets: ParsedSheet[], styles: ParsedStyle[], warn: (msg: string) => void): void {
  let maxSeen = -1;
  let offenders = 0;
  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      if (cell.styleIndex === undefined) continue;
      if (cell.styleIndex > maxSeen) maxSeen = cell.styleIndex;
      if (cell.styleIndex >= styles.length) offenders++;
    }
  }
  if (offenders > 0) {
    warn(`共有 ${offenders} 个单元格引用了越界的样式索引（最大 s=${maxSeen}，styles 长度 ${styles.length}），已按默认样式处理`);
  }
}

function dedupeWarnings(report: ParsedReport): void {
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const w of report.warnings) {
    if (seen.has(w)) continue;
    seen.add(w);
    unique.push(w);
  }
  report.warnings = unique;
  report.unsupported = [...new Set(report.unsupported)];
  report.preserved = [...new Set(report.preserved ?? [])];
  // raw.entries 不在这里裁剪，保留完整性交给调用方
}

function baseName(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash < 0 ? path : path.slice(slash + 1);
}

/** rels 缺失时的路径猜测（0 -> xl/worksheets/sheet1.xml） */
function fallbackSheetTarget(index: number): string {
  return `xl/worksheets/sheet${index + 1}.xml`;
}

function errText(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}
