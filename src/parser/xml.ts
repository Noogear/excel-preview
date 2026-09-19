/**
 * 极简 XML 词法工具（无 DOM、无第三方依赖）：基于索引的单遍扫描，子树只记录原文区间
 * `[openStart, closeEnd)`，对外 API 都在该区间上做子串操作。名字保留原始写法（含命名空间前缀）。
 * 容错：标签不平衡/未闭合、注释、PI、CDATA 都不抛异常，最坏只是提前结束扫描。
 * 不用 `DOMParser`：需在无 DOM 环境运行，且要保留 raw 原文（导出时按字节区间改 XML）。
 */

/** 常用字符 charCode；用普通对象而非 `const enum`（后者在 isolatedModules / 原生 TS 下不可用） */
const Ch = {
  Lt: 60, // <
  Gt: 62, // >
  Amp: 38, // &
  Semi: 59, // ;
  Hash: 35, // #
  Slash: 47, // /
  Bang: 33, // !
  Question: 63, // ?
  Eq: 61, // =
  Single: 39, // '
  Double: 34, // "
  Dash: 45, // -
  OpenBracket: 91, // [
  Space: 32,
  Tab: 9,
  Lf: 10,
  Cr: 13,
  LowerX: 120, // x
  UpperX: 88, // X
} as const;

/** 属性的**原始写法**（含前缀），不做命名空间展开 */
export type XmlAttributes = Record<string, string>;

const NAME_STOP = new Set<number>([
  Ch.Space, Ch.Tab, Ch.Lf, Ch.Cr, Ch.Slash, Ch.Gt, Ch.Eq,
]);

/** 单次遍历的最大步数保护（脏数据兜底，正常 XML 远达不到） */
const MAX_VISIT_STEPS = 5_000_000;

/** 标签名停止符内联判定：名字扫描是逐字符热循环，直接比较比 `Set.has` 快一个档次（须与 `NAME_STOP` 一致） */
function isNameStop(c: number): boolean {
  return c === Ch.Space || c === Ch.Tab || c === Ch.Lf || c === Ch.Cr || c === Ch.Slash || c === Ch.Gt || c === Ch.Eq;
}

function isWs(c: number): boolean {
  return c === Ch.Space || c === Ch.Tab || c === Ch.Lf || c === Ch.Cr;
}

/**
 * 从 `from` 开始找 `</name ...>` 的起始位置。
 *
 * 不能直接用 `indexOf('</' + name)`：那会把 `</rPr>` 当成 `</r>` 的结束标签
 * （`<r>` 是 `<rPr>` 的前缀，ExcelJS 富文本 `<r><rPr>` 正好踩这个坑）；
 * 必须要求名字之后紧跟 `>` 或空白才算命中。
 */

/** `'</' + name` 的缓存：消除上百万次临时字符串分配；上限防脏数据撑爆内存 */
const CLOSE_TAG_CACHE = new Map<string, string>();
const CLOSE_TAG_CACHE_MAX = 256;

function closeTagOf(name: string): string {
  const cached = CLOSE_TAG_CACHE.get(name);
  if (cached !== undefined) return cached;
  const built = `</${name}`;
  if (CLOSE_TAG_CACHE.size < CLOSE_TAG_CACHE_MAX) CLOSE_TAG_CACHE.set(name, built);
  return built;
}

export function findCloseTag(xml: string, name: string, from: number, limit: number): number {
  const closeTag = closeTagOf(name);
  const len = Math.min(limit, xml.length);
  let at = from;
  for (;;) {
    const found = xml.indexOf(closeTag, at);
    if (found < 0 || found >= len) return -1;
    const after = found + closeTag.length;
    const c = after >= len ? Ch.Gt : xml.charCodeAt(after);
    if (c === Ch.Gt) return found;
    if (isWs(c)) {
      // `</name >`、`</name foo="1">` 等带空白的写法
      const gt = xml.indexOf('>', after);
      if (gt >= 0 && gt < len) return found;
    }
    at = found + 2; // 继续找下一个候选
  }
}

const ENTITY_TABLE: ReadonlyMap<string, string> = new Map<string, string>([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
  ['nbsp', '\u00a0'],
]);

function codePointToString(cp: number): string {
  if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff) return '';
  try {
    return String.fromCodePoint(cp);
  } catch {
    return '';
  }
}

function decodeEntity(raw: string): string {
  if (raw.length === 0) return '';
  if (raw.charCodeAt(0) === Ch.Hash) {
    const isHex = raw.length > 1 && (raw.charCodeAt(1) === Ch.LowerX || raw.charCodeAt(1) === Ch.UpperX);
    const digits = raw.slice(isHex ? 2 : 1);
    const cp = isHex ? Number.parseInt(digits, 16) : Number.parseInt(digits, 10);
    return codePointToString(cp);
  }
  const known = ENTITY_TABLE.get(raw);
  // 未知实体原样保留（`&` + raw + `;`），避免静默丢内容
  return known ?? `&${raw};`;
}

/** 解码 XML 实体；没有 `&` 时零分配直接返回原串 */
export function decodeEntities(text: string): string {
  if (text.indexOf('&') < 0) return text;
  let out = '';
  let i = 0;
  while (i < text.length) {
    const amp = text.indexOf('&', i);
    if (amp < 0) {
      out += text.slice(i);
      break;
    }
    const semi = text.indexOf(';', amp + 1);
    // 限制长度，避免把 `a & b ; c` 这种当成实体拼命扫描
    if (semi < 0 || semi - amp > 12) {
      out += text.slice(i, amp + 1);
      i = amp + 1;
      continue;
    }
    out += text.slice(i, amp) + decodeEntity(text.slice(amp + 1, semi));
    i = semi + 1;
  }
  return out;
}

export function parseAttrs(xml: string, start: number, end: number): XmlAttributes {
  const attrs: XmlAttributes = {};
  let i = start;
  while (i < end) {
    while (i < end && (isWs(xml.charCodeAt(i)) || xml.charCodeAt(i) === Ch.Slash)) i++;
    const nameStart = i;
    while (i < end && !isNameStop(xml.charCodeAt(i))) i++;
    if (i === nameStart) break; // 既不是名字也不是空白，防止死循环
    const name = xml.slice(nameStart, i);
    while (i < end && isWs(xml.charCodeAt(i))) i++;
    if (xml.charCodeAt(i) !== Ch.Eq) {
      // 无值属性：按 XML 规范不合法，但容错为 `true`
      if (name.length > 0) attrs[name] = 'true';
      continue;
    }
    i++; // 跳过 '='
    while (i < end && isWs(xml.charCodeAt(i))) i++;
    const quote = xml.charCodeAt(i);
    if (quote === Ch.Double || quote === Ch.Single) {
      i++;
      const valueStart = i;
      while (i < end && xml.charCodeAt(i) !== quote) i++;
      attrs[name] = decodeEntities(xml.slice(valueStart, i));
      if (i < end) i++; // 跳过收尾引号
    } else {
      // 未加引号（非法 OOXML，但容错）：读到空白为止
      const valueStart = i;
      while (i < end && !isWs(xml.charCodeAt(i)) && xml.charCodeAt(i) !== Ch.Gt) i++;
      attrs[name] = decodeEntities(xml.slice(valueStart, i));
    }
  }
  return attrs;
}

export interface ElementRange {
  name: string;
  openStart: number;
  /** 开始标签 `>` 之后一位 */
  openEnd: number;
  /** 整棵子树结束后的偏移；自闭合元素等于 `openEnd` */
  closeEnd: number;
  selfClosing: boolean;
  attrs: XmlAttributes;
  /** 该元素在原文中的起始偏移（等同于 openStart），便于继续做子扫描 */
  readonly start: number;
}

export interface VisitCtx {
  /** 当前元素闭区间之后的下一个偏移，用于跳过整棵子树 */
  readonly afterEnd: number;
}

/** 访问控制：返回 `false` = 跳过该元素（不深入其子树），继续扫后面的兄弟 */
export type ElementVisitor = (el: ElementRange, ctx: VisitCtx) => boolean | void;

export interface VisitOptions {
  pruneOnFalse?: boolean;
  /** 只访问这些名字的元素；前缀无关（`x:sheetData` 也能被 `sheetData` 命中）；缺省访问全部 */
  names?: ReadonlySet<string>;
  /** 与 `names` 联合判断是否命中（给了 `match` 就忽略 `names`） */
  match?: (name: string) => boolean;
  /** 扫描起点（默认 0）。在**原始 xml 坐标系**里只扫某个区间，免去切片 */
  from?: number;
  /** 扫描终点（开区间，默认 xml.length） */
  to?: number;
  /** 只扫**直接子元素**（深度 1）：遇到开始标签整棵跳过，遇到结束标签收工 */
  depthLimit?: boolean;
  /** 命中第一个元素并回调完就**结束整趟遍历**（否则返回 false 只跳过该子树，仍会走完整份文档） */
  stopOnFirstMatch?: boolean;
}

/**
 * 单遍扫描的事件式遍历：深度优先，回调返回 `false` 时跳过该元素整棵子树。
 * 只在"标签"上落点，每处理完一个标签就把游标挪到它开头之后，同一个 `<` 不会被处理两次。
 */
export function visitElements(xml: string, visit: ElementVisitor, opts: VisitOptions = {}): boolean {
  const len = xml.length;
  const { pruneOnFalse = true, names, match, depthLimit = false, stopOnFirstMatch = false } = opts;

  const visitRange = (s: number, e: number): void => {
    let i = s;
    let guard = 0;
    while (i < e) {
      if (guard++ > MAX_VISIT_STEPS) return; // 防御：脏数据不应导致死循环
      const lt = xml.indexOf('<', i);
      if (lt < 0 || lt >= e) return;
      const nx = xml.charCodeAt(lt + 1);

      if (nx === Ch.Bang) {
        if (xml.startsWith('<![CDATA[', lt)) {
          const end = xml.indexOf(']]>', lt + 9);
          if (end < 0) return;
          i = end + 3;
        } else {
          const end = xml.indexOf('>', lt + 2);
          if (end < 0) return;
          i = end + 1;
        }
        continue;
      }
      if (nx === Ch.Question) {
        const end = xml.indexOf('?>', lt + 2);
        if (end < 0) return;
        i = end + 2;
        continue;
      }
      // 结束标签：深度 1 模式下它意味着"父元素到此为止"，直接收工
      if (nx === Ch.Slash) {
        if (depthLimit) return;
        const end = xml.indexOf('>', lt + 2);
        if (end < 0) return;
        i = end + 1;
        continue;
      }

      // 标签名保留原始写法（含命名空间前缀）
      const nameStart = lt + 1;
      let nameEnd = nameStart;
      while (nameEnd < e && !isNameStop(xml.charCodeAt(nameEnd))) nameEnd++;
      const name = xml.slice(nameStart, nameEnd);

      // 跳过引号内的内容，找到开始标签的 `>`
      let q = nameEnd;
      let quote = 0;
      let gt = -1;
      while (q < e) {
        const c = xml.charCodeAt(q);
        if (quote !== 0) {
          if (c === quote) quote = 0;
        } else if (c === Ch.Double || c === Ch.Single) {
          quote = c;
        } else if (c === Ch.Gt) {
          gt = q;
          break;
        }
        q++;
      }
      if (gt < 0) return; // 标签未闭合

      let back = gt - 1;
      while (back > nameEnd && isWs(xml.charCodeAt(back))) back--;
      const selfClosing = xml.charCodeAt(back) === Ch.Slash;
      const attrEnd = selfClosing ? back : gt;

      let closeEnd = gt + 1;
      if (!selfClosing) {
        const found = findCloseTag(xml, name, gt + 1, e);
        if (found < 0) {
          i = gt + 1; // 容错：标签不平衡，继续往后扫
          continue;
        }
        const gt2 = xml.indexOf('>', found + name.length + 2);
        closeEnd = gt2 < 0 ? found + name.length + 2 : gt2 + 1;
      }

      const wanted = match ? match(name) : !names || names.has(name) || names.has(localName(name));
      if (wanted) {
        const el: ElementRange = {
          name,
          openStart: lt,
          openEnd: gt + 1,
          closeEnd,
          selfClosing,
          attrs: parseAttrs(xml, nameEnd, attrEnd),
          start: lt,
        };
        if (visit(el, { afterEnd: closeEnd }) === false && pruneOnFalse) {
          if (stopOnFirstMatch) return; // 命中即收工，不再扫后面
          i = closeEnd; // 不深入该元素，继续扫后面的兄弟
          continue;
        }
        if (stopOnFirstMatch) return;
      }

      // 深度 1：不论命中与否都跳过整棵子树，只留直接子元素
      if (depthLimit) {
        i = closeEnd;
        continue;
      }

      // 无论是否命中都继续向下扫：目标元素可能嵌在任意深度的非目标元素里
      i = lt + 1;
    }
  };

  visitRange(opts.from ?? 0, Math.min(opts.to ?? len, len));
  return true;
}

export function findElements(xml: string, name: string): ElementRange[] {
  const out: ElementRange[] = [];
  visitElements(xml, (el) => { out.push(el); }, { names: new Set([name]) });
  return out;
}

/**
 * 找某个名字（含带前缀写法）在 XML 里**最早**出现的位置；两种写法都没有时返回 `fallback`。
 * 带前缀的写法（`<x:conditionalFormatting>`）不会被 `indexOf('<conditionalFormatting')` 命中，
 * 所以两种都要查；代价只有两次原生 indexOf，省掉"为每个特性各走一趟全文档"。
 */
export function earliestNameAt(xml: string, localName: string, fallback: number): number {
  // 快路径：元素就在尾段里（正常文件都如此）——只扫尾段，代价几乎为零
  const plainInTail = xml.indexOf(`<${localName}`, fallback);
  if (plainInTail >= 0) return plainInTail;
  const prefixedInTail = xml.indexOf(`:${localName}`, fallback);
  if (prefixedInTail >= 0) return prefixedInTail - 1;

  // 尾段没有：再确认它是否出现在更前面（少见，但必须支持，不能漏）
  const plain = xml.indexOf(`<${localName}`);
  if (plain >= 0) return plain;
  const prefixed = xml.indexOf(`:${localName}`);
  return prefixed < 0 ? fallback : prefixed - 1;
}

/** 找出第一个名字等于 `name` 的元素（命中即停，不做无谓的全文档扫描） */
export function findFirstElement(xml: string, name: string, from = 0): ElementRange | undefined {
  let found: ElementRange | undefined;
  visitElements(xml, (el) => { found = el; return false; }, {
    names: new Set([name]),
    stopOnFirstMatch: true,
    from,
  });
  return found;
}

/**
 * 某个元素的**直接**子元素（不含孙子）。用 `depthLimit` 在原始 xml 上直接扫：不切子串、
 * 不把偏移折算回外层坐标系（每格都调一次，切片 + 展开对象在百万格规模下开销可观）。
 */
export function childElements(xml: string, el: ElementRange): ElementRange[] {
  const out: ElementRange[] = [];
  if (el.selfClosing) return out;
  visitElements(xml, (child) => { out.push(child); }, {
    depthLimit: true,
    from: el.openEnd,
    to: Math.max(el.openEnd, el.closeEnd),
  });
  return out;
}

/** 元素名匹配器：支持 `a:clrScheme` 与 `clrScheme` 互认（前缀无关） */
export function localName(name: string): string {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
}

export function nameMatches(name: string, wanted: string): boolean {
  return name === wanted || localName(name) === wanted;
}

/** 取元素原文（含标记） */
export function rawText(xml: string, el: ElementRange): string {
  return xml.slice(el.openStart, el.closeEnd);
}

/** 取元素**内部**原文（不含开始/结束标签），保留内层标记 */
export function innerRaw(xml: string, el: ElementRange): string {
  if (el.selfClosing) return '';
  const end = Math.max(el.openEnd, el.closeEnd);
  return stripCloseTag(xml.slice(el.openEnd, end), el.name);
}

function stripCloseTag(inner: string, name: string): string {
  const closeIdx = inner.lastIndexOf('</');
  if (closeIdx < 0) return inner;
  const tail = inner.slice(closeIdx);
  const m = /^<\/\s*([^\s>]+)\s*>$/.exec(tail);
  if (!m) return inner;
  return nameMatches(m[1], localName(name)) ? inner.slice(0, closeIdx) : inner;
}

/**
 * 读取元素文本：剥离所有内层标记（含 CDATA）并解码实体；`<v>123</v>` 这类叶子元素走快路径。
 */
export function elementText(xml: string, el: ElementRange): string {
  if (el.selfClosing) return '';
  const raw = innerRaw(xml, el);
  if (raw.indexOf('<') < 0) return decodeEntities(raw);
  return decodeEntities(stripTags(raw));
}

/** 去掉字符串里的所有 `<...>` 标记，并把 CDATA 内容取出 */
export function stripTags(raw: string): string {
  if (raw.indexOf('<') < 0) return raw;
  let out = '';
  let i = 0;
  while (i < raw.length) {
    const lt = raw.indexOf('<', i);
    if (lt < 0) {
      out += raw.slice(i);
      break;
    }
    out += raw.slice(i, lt);
    if (raw.startsWith('<![CDATA[', lt)) {
      const end = raw.indexOf(']]>', lt + 9);
      if (end < 0) {
        out += raw.slice(lt + 9);
        break;
      }
      out += raw.slice(lt + 9, end);
      i = end + 3;
      continue;
    }
    const gt = raw.indexOf('>', lt);
    if (gt < 0) break;
    i = gt + 1;
  }
  return out;
}

export function firstElementText(xml: string, name: string): string | undefined {
  const el = findFirstElement(xml, name);
  return el ? elementText(xml, el) : undefined;
}

/** 属性布尔语义：`1/true/on` 为真（XOXML 用 `1`/`0`） */
export function attrBool(attrs: XmlAttributes, key: string): boolean | undefined {
  const v = attrs[key];
  if (v === undefined) return undefined;
  return isTruthyAttr(v);
}

export function isTruthyAttr(v: string): boolean {
  if (v === '' || v === '1' || v === 'true' || v === 'on') return true;
  if (v === '0' || v === 'false' || v === 'off') return false;
  return true;
}

export function attrNumber(attrs: XmlAttributes, key: string): number | undefined {
  const v = attrs[key];
  if (v === undefined || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

export function attrInt(attrs: XmlAttributes, key: string): number | undefined {
  const n = attrNumber(attrs, key);
  return n === undefined ? undefined : Math.trunc(n);
}
