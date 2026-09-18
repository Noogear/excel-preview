/**
 * 浮动图片 / 绘图解析：`xl/drawings/drawingN.xml`（+ `xl/drawings/_rels/drawingN.xml.rels`）。
 *
 * 覆盖三种锚点：
 * - `xdr:oneCellAnchor`：`from`（col/row + colOff/rowOff，EMU）+ `xdr:ext`（cx/cy，EMU）
 * - `xdr:twoCellAnchor`：`from` + `to`
 * - `xdr:absoluteAnchor`：`xdr:pos`（x/y，EMU，相对工作表左上角）+ `xdr:ext`
 *
 * 媒体路径：图片本体通过 `<a:blip r:embed="rId1"/>` 引用，`rId1` 要查
 * `xl/drawings/_rels/drawing1.xml.rels` 的 `Target="../media/image1.png"`，
 * 再按 drawing 部件所在目录归一化成 zip 内的路径 `xl/media/image1.png`
 * （直接看 Target 会得到带 `../` 的相对路径，zip 里没有这个条目名）。
 *
 * 非图片的绘图对象（图表 / 形状 / 文本框 / SmartArt / 分组）不解析，但**按种类计数**
 * 返回给编排层进 `report.unsupported`——不能因为"这个文件里有图也有图表"就把图表悄悄丢掉。
 */
import type { ParsedImage } from './types';
import type { Warn } from './worksheet';
import {
  attrNumber, childElements, elementText, findFirstElement, localName, visitElements,
  type ElementRange,
} from './xml';
import { dirName, resolvePath } from './zip';

const DEFAULT_PART_PATH = 'xl/drawings/drawing1.xml';

export interface ParseDrawingOptions {
  /** `xl/drawings/_rels/drawingN.xml.rels` 的原文 */
  relsXml?: string;
  /** 或直接给 relId -> zip 内路径 的映射（与 `relsXml` 同时给时以本映射为准） */
  mediaTargets?: ReadonlyMap<string, string>;
  /** drawing 部件在 zip 内的路径；决定 `../media/x.png` 如何归一化（默认 xl/drawings/drawing1.xml） */
  partPath?: string;
  warn?: Warn;
}

export interface ParseDrawingResult {
  images: ParsedImage[];
  /** 未支持特性，形如 `图表(chart) · 1 处（未解析）` */
  unsupported: string[];
}

type AnchorType = ParsedImage['anchorType'];

const ANCHOR_TYPES: ReadonlyMap<string, AnchorType> = new Map<string, AnchorType>([
  ['oneCellAnchor', 'oneCell'],
  ['twoCellAnchor', 'twoCell'],
  ['absoluteAnchor', 'absolute'],
]);

/** 锚点里出现即代表"这类绘图对象不解析"的元素 -> 记账标签 */
const OBJECT_LABELS: ReadonlyMap<string, string> = new Map<string, string>([
  ['sp', '形状/文本框(shape)'],
  ['cxnSp', '连接线(cxnSp)'],
  ['grpSp', '组合形状(grpSp)'],
  ['contentPart', '内容部件(contentPart)'],
]);

/** 解析绘图部件 */
export function parseDrawing(xml: string, opts: ParseDrawingOptions = {}): ParseDrawingResult {
  const warn: Warn = opts.warn ?? (() => {});
  const partPath = opts.partPath ?? DEFAULT_PART_PATH;
  const targets = buildTargetMap(opts, partPath, warn);

  const anchors: ElementRange[] = [];
  visitElements(xml, (el) => {
    anchors.push(el);
    return false; // 锚点内部按需再扫，避免整棵树都被建出来
  }, { match: (name) => ANCHOR_TYPES.has(localName(name)) });

  const images: ParsedImage[] = [];
  const counts = new Map<string, number>();
  const bump = (label: string): void => { counts.set(label, (counts.get(label) ?? 0) + 1); };

  let index = 0;
  for (const anchor of anchors) {
    index++;
    const anchorType = ANCHOR_TYPES.get(localName(anchor.name));
    if (!anchorType) continue;

    let from: ParsedImage['from'] | undefined;
    let to: ParsedImage['to'] | undefined;
    let ext: { cx: number; cy: number } | undefined;
    let pos: { x: number; y: number } | undefined;
    let object: DrawingObject | undefined;

    for (const kid of childElements(xml, anchor)) {
      const ln = localName(kid.name);
      if (ln === 'from') from = readMarker(xml, kid);
      else if (ln === 'to') to = readMarker(xml, kid);
      else if (ln === 'ext') ext = readExt(kid);
      else if (ln === 'pos') pos = { x: attrNumber(kid.attrs, 'x') ?? 0, y: attrNumber(kid.attrs, 'y') ?? 0 };
      else if (ln === 'pic') object = readPicture(xml, kid);
      else if (ln === 'graphicFrame') object = readGraphicFrame(xml, kid);
      else {
        const kind = nonPictureKindOf(ln);
        if (kind !== undefined) object = { kind, cNvId: readCNvId(xml, kid) };
      }
    }

    if (object === undefined) {
      warn(`${partPath}: 第 ${index} 个锚点(${anchorType})里没有可识别的绘图对象，已跳过`);
      continue;
    }
    if (object.kind !== 'picture') {
      // graphicFrame 里若是图表引用，按"图表"记账，否则是 SmartArt/其它图形框架
      if (object.kind === 'graphicFrame') {
        bump(object.chartRelId !== undefined ? '图表(chart)' : 'SmartArt/图形框架(graphicFrame)');
      } else {
        bump(OBJECT_LABELS.get(object.kind) ?? object.kind);
      }
      continue;
    }

    const id = object.cNvId ?? `img${index}`;
    if (object.embed === undefined) {
      warn(`${partPath}: 图片 ${id} 没有 <a:blip r:embed>（可能是链接图片），已跳过`);
      continue;
    }
    const rel = targets.get(object.embed);
    if (rel === undefined) {
      warn(`${partPath}: 图片 ${id} 的关系 ${object.embed} 在 drawing rels 里找不到，已跳过`);
      continue;
    }
    if (rel.external) {
      warn(`${partPath}: 图片 ${id} 指向外部资源 ${rel.target}（未下载），已跳过`);
      continue;
    }

    const image: ParsedImage = {
      id,
      mediaPath: rel.target,
      anchorType,
      from: from ?? { col: 0, row: 0 },
    };
    if (from === undefined) warn(`${partPath}: 图片 ${id} 缺少 <xdr:from>，锚点位置按 A1 处理`);
    if (to !== undefined) image.to = to;
    if (ext !== undefined) image.extEmu = ext;
    if (pos !== undefined) image.posEmu = pos;
    images.push(image);
  }

  return {
    images,
    unsupported: [...counts.entries()].map(([label, n]) => `${label} · ${n} 处（未解析）`),
  };
}

/* -------------------------------------------------------------------------- */
/* 关系（媒体路径）                                                             */
/* -------------------------------------------------------------------------- */

interface RelTarget { target: string; external: boolean }

function buildTargetMap(
  opts: ParseDrawingOptions,
  partPath: string,
  warn: Warn,
): Map<string, RelTarget> {
  const map = new Map<string, RelTarget>();
  if (opts.relsXml) {
    const root = findFirstElement(opts.relsXml, 'Relationships');
    if (!root) {
      warn(`${partPath} 的 rels 缺少 <Relationships>，媒体路径无法解析`);
    } else {
      const baseDir = dirName(partPath);
      for (const rel of childElements(opts.relsXml, root)) {
        if (localName(rel.name) !== 'Relationship') continue;
        const id = rel.attrs['Id'];
        const target = rel.attrs['Target'];
        if (!id || !target) continue;
        const external = rel.attrs['TargetMode'] === 'External';
        // `../media/image1.png` 相对 drawing 部件所在目录 -> `xl/media/image1.png`
        map.set(id, { target: external ? target : resolvePath(baseDir, target), external });
      }
    }
  }
  if (opts.mediaTargets) {
    for (const [id, target] of opts.mediaTargets) map.set(id, { target, external: false });
  }
  return map;
}

/* -------------------------------------------------------------------------- */
/* 锚点 / 对象                                                                  */
/* -------------------------------------------------------------------------- */

function readMarker(xml: string, el: ElementRange): ParsedImage['from'] {
  const marker: ParsedImage['from'] = { col: 0, row: 0 };
  for (const kid of childElements(xml, el)) {
    const ln = localName(kid.name);
    if (ln === 'col') marker.col = intText(xml, kid) ?? 0;
    else if (ln === 'row') marker.row = intText(xml, kid) ?? 0;
    else if (ln === 'colOff') marker.colOffEmu = intText(xml, kid) ?? 0;
    else if (ln === 'rowOff') marker.rowOffEmu = intText(xml, kid) ?? 0;
  }
  return marker;
}

function readExt(el: ElementRange): { cx: number; cy: number } {
  return { cx: attrNumber(el.attrs, 'cx') ?? 0, cy: attrNumber(el.attrs, 'cy') ?? 0 };
}

interface DrawingObject {
  kind: 'picture' | 'graphicFrame' | 'sp' | 'cxnSp' | 'grpSp' | 'contentPart';
  cNvId?: string;
  embed?: string;
  chartRelId?: string;
}

/** 非图片的绘图对象元素名 -> 对象种类（`OBJECT_LABELS` 的键与之保持一致） */
function nonPictureKindOf(name: string): DrawingObject['kind'] | undefined {
  if (name === 'sp' || name === 'cxnSp' || name === 'grpSp' || name === 'contentPart') return name;
  return undefined;
}

function readPicture(xml: string, picEl: ElementRange): DrawingObject {
  const found = scanObject(xml, picEl);
  const object: DrawingObject = { kind: 'picture' };
  if (found.cnv) {
    const id = found.cnv.attrs['id'];
    if (id !== undefined && id !== '') object.cNvId = id;
  }
  if (found.blip) {
    const embed = found.blip.attrs['r:embed'] ?? found.blip.attrs['embed'];
    if (embed !== undefined && embed !== '') object.embed = embed;
  }
  return object;
}

function readGraphicFrame(xml: string, frameEl: ElementRange): DrawingObject {
  const found = scanObject(xml, frameEl);
  const object: DrawingObject = { kind: 'graphicFrame' };
  if (found.cnv) {
    const id = found.cnv.attrs['id'];
    if (id !== undefined && id !== '') object.cNvId = id;
  }
  if (found.chart) {
    const relId = found.chart.attrs['r:id'] ?? found.chart.attrs['id'];
    if (relId !== undefined && relId !== '') object.chartRelId = relId;
  }
  return object;
}

function readCNvId(xml: string, el: ElementRange): string | undefined {
  const found = scanObject(xml, el);
  const id = found.cnv?.attrs['id'];
  return id !== undefined && id !== '' ? id : undefined;
}

interface ScannedObject {
  blip?: ElementRange;
  cnv?: ElementRange;
  chart?: ElementRange;
}

const SCAN_NAMES: ReadonlySet<string> = new Set(['blip', 'cNvPr', 'chart']);

/** 在某个绘图对象子树里找 `<a:blip>` / `<xdr:cNvPr>` / `<c:chart>`（各取第一个） */
function scanObject(xml: string, root: ElementRange): ScannedObject {
  const found: ScannedObject = {};
  visitSubtree(xml, root, (el) => {
    const ln = localName(el.name);
    if (ln === 'blip') { if (!found.blip) found.blip = el; }
    else if (ln === 'cNvPr') { if (!found.cnv) found.cnv = el; }
    else if (ln === 'chart') { if (!found.chart) found.chart = el; }
  }, SCAN_NAMES);
  return found;
}

/**
 * 在 `root` 的**子树内**做单遍扫描，并把偏移量换算回原始 XML 的坐标系
 * （`visitElements` 只认整串坐标；这里沿用 `childElements` 的换算方式）。
 */
function visitSubtree(
  xml: string,
  root: ElementRange,
  visit: (el: ElementRange) => void,
  names: ReadonlySet<string>,
): void {
  if (root.selfClosing) return;
  const base = root.openEnd;
  const inner = xml.slice(base, Math.max(base, root.closeEnd));
  visitElements(inner, (child) => {
    visit({
      ...child,
      openStart: child.openStart + base,
      openEnd: child.openEnd + base,
      closeEnd: child.closeEnd + base,
      start: child.openStart + base,
    });
  }, { match: (name) => names.has(localName(name)) });
}

function intText(xml: string, el: ElementRange): number | undefined {
  const text = elementText(xml, el).trim();
  if (text === '') return undefined;
  const n = Number(text);
  return Number.isFinite(n) ? Math.trunc(n) : undefined;
}
