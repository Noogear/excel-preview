/**
 * 极简 XML 词法工具（无 DOM、无第三方依赖）。
 *
 * 设计要点：
 * - 基于索引的单遍扫描。游标只前进，元素树按需构建（`visitElements` 提供零分配遍历）。
 * - 每棵元素子树记录其原文区间 `[openStart, closeEnd)`，几乎全部对外 API 都在这个区间上做
 *   子串操作，因此不需要为每个节点复制字符串，也不会产生海量临时字符串。
 * - 标签名/属性名**保留原始写法**（含命名空间前缀，如 `x:worksheet`、`r:id`）。
 * - 支持 `&amp; &lt; &gt; &quot; &apos;` 与十进制/十六进制数字实体。
 * - 容错：不匹配的结束标签、未闭合的标签、注释/PI/CDATA 都不会抛异常，最坏情况是提前结束扫描。
 *
 * 之所以不用 `DOMParser`：解析器要能在 Node / Worker / 无 DOM 环境运行，且需要保留 raw 原文
 * （后续"外科式修补导出"要按字节区间改写 XML）。
 */

/**
 * 常用字符的 charCode 常量。
 * 刻意用普通对象而不是 `const enum`：`const enum` 在 isolatedModules / Node 原生 TS
 * 等"只擦除类型"的工具链下不可用，而扫描器只需要几个内联常量。
 */
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

/**
 * 标签名停止符的内联判定。
 *
 * 原来用 `NAME_STOP.has(c)`（Set 查找）：名字扫描是**每个元素**都要跑的逐字符循环，
 * 百万级单元格下这里会被调用上千万次，Set 的哈希查找比直接比较慢一个档次（实测占比可观）。
 */
function isNameStop(c: number): boolean {
  return c === Ch.Space || c === Ch.Tab || c === Ch.Lf || c === Ch.Cr || c === Ch.Slash || c === Ch.Gt || c === Ch.Eq;
}

function isWs(c: number): boolean {
  return c === Ch.Space || c === Ch.Tab || c === Ch.Lf || c === Ch.Cr;
}

/**
 * 从 `from` 开始找 `</name ...>` 的起始位置。
 *
 * 注意：**不能**直接用 `indexOf('</' + name)`——那会把 `</rPr>` 当成 `</r>` 的结束标签
 * （`<r>` 是 `<rPr>` 的前缀，ExcelJS 产出的富文本 `<r><rPr>` 正好踩这个坑）。
 * 这里要求名字之后紧跟 `>` 或空白，才能算命中。
 */
/**
 * `'</' + name` 的缓存。
 *
 * 每个非自闭合元素都要构造一次结束标签前缀；百万级单元格下这是上百万次临时字符串分配。
 * 标签名种类是个位数，缓存后彻底消除这部分分配与 GC 压力（上限纯粹是防脏数据撑爆内存）。
 */
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

/* -------------------------------------------------------------------------- */
/* 实体解码                                                                    */
/* -------------------------------------------------------------------------- */

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
    // &#10; / &#x41;
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

/* -------------------------------------------------------------------------- */
/* 属性解析                                                                    */
/* -------------------------------------------------------------------------- */

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

/* -------------------------------------------------------------------------- */
/* 零分配遍历                                                                  */
/* -------------------------------------------------------------------------- */

export interface ElementRange {
  name: string;
  /** 开始标签的 `<` 在原文中的偏移 */
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
  /** true = 回调返回 false 时跳过该元素的子树（默认 true） */
  pruneOnFalse?: boolean;
  /** 只访问这些名字的元素；前缀无关（`x:sheetData` 也能被 `sheetData` 命中）；缺省访问全部 */
  names?: ReadonlySet<string>;
  /** 与 `names` 联合判断是否命中（给了 `match` 就忽略 `names`） */
  match?: (name: string) => boolean;
  /** 扫描起点（默认 0）。用于在**原始 xml 坐标系**里只扫某个区间，免去切片 */
  from?: number;
  /** 扫描终点（开区间，默认 xml.length） */
  to?: number;
  /**
   * 只扫**直接子元素**（深度 1）：
   * 遇到开始标签就整棵跳过，遇到结束标签就收工。
   * 这样 `childElements` 不必先 `slice` 出一段子串再折算偏移。
   */
  depthLimit?: boolean;
  /**
   * 命中第一个元素并回调完就**结束整趟遍历**。
   *
   * `findFirstElement` 必须开这个：否则回调返回 false 只是"跳过该子树"，
   * 遍历仍会把整份文档走完——对 50MB 的 sheet XML（四百万个元素）来说，
   * 每次"找第一个 X"都变成一次全量扫描，实测这是解析耗时的大头。
   */
  stopOnFirstMatch?: boolean;
}

/**
 * 单遍扫描 XML 的事件式遍历：深度优先，回调返回 `false` 时跳过该元素整棵子树。
 * 不构建树、不为每个元素建对象（只构造一个轻量 `ElementRange`）。
 *
 * 索引约定：扫描器只在"标签"上落点——每处理完一个标签就把游标挪到它开头之后，
 * 下一轮 `indexOf('<')` 自然会越过文本内容。同一个 `<` 不会被处理两次（唯一的
 * 例外是上一轮 `indexOf` 找不到下一个标签时 `lt` 变成 -1，循环随即结束）。
 */
export function visitElements(xml: string, visit: ElementVisitor, opts: VisitOptions = {}): boolean {
  const len = xml.length;
  const { pruneOnFalse = true, names, match, depthLimit = false, stopOnFirstMatch = false } = opts;

  // 起止位置都在**原始 xml** 坐标系里，不做任何折算
  const visitRange = (s: number, e: number): void => {
    let i = s;
    let guard = 0;
    while (i < e) {
      if (guard++ > MAX_VISIT_STEPS) return; // 防御：脏数据不应导致死循环
      const lt = xml.indexOf('<', i);
      if (lt < 0 || lt >= e) return;
      const nx = xml.charCodeAt(lt + 1);

      // 注释 / 处理指令 / DOCTYPE / CDATA
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

      // 标签名（保留原始写法，含命名空间前缀）
      const nameStart = lt + 1;
      let nameEnd = nameStart;
      while (nameEnd < e && !isNameStop(xml.charCodeAt(nameEnd))) nameEnd++;
      const name = xml.slice(nameStart, nameEnd);

      // 开始标签的 `>`（跳过引号内的内容）
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

      // `/>` -> 自闭合
      let back = gt - 1;
      while (back > nameEnd && isWs(xml.charCodeAt(back))) back--;
      const selfClosing = xml.charCodeAt(back) === Ch.Slash;
      const attrEnd = selfClosing ? back : gt;

      // 结束标签位置（自闭合元素直接等于 openEnd）
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

/** 找出所有名字等于 `name` 的元素（任意深度），返回轻量区间对象数组 */
export function findElements(xml: string, name: string): ElementRange[] {
  const out: ElementRange[] = [];
  visitElements(xml, (el) => { out.push(el); }, { names: new Set([name]) });
  return out;
}

/**
 * 找某个名字（含带前缀写法）在 XML 里**最早**出现的位置；两种写法都没有时返回 `fallback`。
 *
 * 用途：带前缀的现代写法（`<x:conditionalFormatting>`）不会被 `indexOf('<conditionalFormatting')`
 * 命中，所以两种都要查；查得到就从那里开始扫，查不到就说明这份 XML 里根本没有该元素，
 * 直接从 `fallback` 起步即可（`fallback` 之后必然也不会有）。
 * 代价只有两次原生 indexOf，换来的是"不必为每个特性各走一趟全文档"。
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
 * 某个元素的**直接**子元素（子元素的子元素不会出现在结果里）。
 *
 * 用 `depthLimit` 在**原始 xml** 上直接扫：不再 `slice` 出子串，也不需要把偏移折算回外层
 * 坐标系。每个单元格都调一次 childElements，旧实现的"切片 + 每子节点展开对象 + 重新折算"
 * 在百万格规模下是实打实的开销。
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
  // 只有形如 </name ...> 或 </x:name> 才剥离
  const m = /^<\/\s*([^\s>]+)\s*>$/.exec(tail);
  if (!m) return inner;
  return nameMatches(m[1], localName(name)) ? inner.slice(0, closeIdx) : inner;
}

/**
 * 读取元素的文本内容：剥离所有内层标记（含 CDATA），解码实体。
 * 对 `<v>123</v>`、`<t>a&amp;b</t>` 这类叶子元素是快路径。
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

/** 便捷读取：XML 中第一个 `name` 元素的文本 */
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
