/**
 * 样式解析：`xl/styles.xml` + `xl/theme/theme1.xml`。
 *
 * 产出 `ParsedStyle[]`，索引与 `cellXfs` 顺序一一对应（即工作表里 `c/@s` 的取值）。
 * 全部颜色统一归一化为 `#RRGGBB`；无法解析的颜色记 warning 并省略该字段（不抛异常）。
 *
 * 引用的规范：ECMA-376 Part 1 §18.8（styles）、§20.1.6（theme clrScheme）、
 * §18.8.27（numFmt 内置格式）。
 */
import type { BorderLineStyle, CfDxfStyle, ParsedBorder, ParsedStyle } from './types';
import {
  attrBool, attrInt, attrNumber, childElements, decodeEntities, elementText,
  findFirstElement, localName, nameMatches, visitElements,
  type ElementRange, type XmlAttributes,
} from './xml';

/* -------------------------------------------------------------------------- */
/* 内置 numFmt 表（id 0–49）                                                   */
/* -------------------------------------------------------------------------- */

/**
 * 内置格式 id -> pattern。
 * 说明：ECMA-376 里 14–22、45–47 的**具体 pattern 由区域设置决定**，
 * 绝大多数区域与 Excel 的经典实现一致，这里采用其通行写法。
 */
export const BUILTIN_NUM_FMTS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  /**
   * 23–36：**区域相关**的内置格式（ECMA-376 §18.8.30 把它们留给东亚/其它区域）。
   *
   * 这里按**中日韩区域的通行实现**填上（27–36 是中日文日期/时间写法）。
   * 实测来历：`fixture-numfmt` 的 `yyyy"年"m"月"d"日"` 在 xlsx 里是自定义格式，
   * 而 Excel「另存为 .xls」时把它换成了**内置 id 31** —— 早期表里 31 写的是 `General`，
   * 于是 .xls 里这个日期格式整条丢掉（用户看到日期变成"常规"）。
   */
  23: 'General',
  24: 'General',
  25: 'General',
  26: 'General',
  27: 'yyyy"年"m"月"',
  28: 'm"月"d"日"',
  29: 'm"月"d"日"',
  30: 'm-d-yy',
  31: 'yyyy"年"m"月"d"日"',
  32: 'h"时"mm"分"',
  33: 'h"时"mm"分"ss"秒"',
  34: '上午/下午h"时"mm"分"',
  35: '上午/下午h"时"mm"分"ss"秒"',
  36: 'yyyy"年"m"月"',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

/** `numFmtId="164"` 起才是自定义格式；小于该值优先查内置表 */
export const CUSTOM_NUM_FMT_BASE_ID = 164;

/* -------------------------------------------------------------------------- */
/* indexex 调色板（ECMA-376 §18.8.27 的继承表 + Excel 经典 64 色调色板）          */
/* -------------------------------------------------------------------------- */

/**
 * 索引调色板（ECMA-376 §18.8.27 继承的 Excel 经典 64 色调色板）。
 * 0/1 是系统前景/背景，2–7 是经典八色，8–15 是重复的"半亮"档，16 起是扩展色。
 */
export const INDEXED_COLORS: Readonly<Record<number, string>> = {
  0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00', 4: '#0000FF',
  5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF',
  8: '#000000', 9: '#FFFFFF', 10: '#FF0000', 11: '#00FF00', 12: '#0000FF',
  13: '#FFFF00', 14: '#FF00FF', 15: '#00FFFF',
  16: '#800000', 17: '#008000', 18: '#000080', 19: '#808000',
  20: '#800080', 21: '#008080', 22: '#C0C0C0', 23: '#808080',
  24: '#9999FF', 25: '#993366', 26: '#FFFFCC', 27: '#CCFFFF',
  28: '#660066', 29: '#FF8080', 30: '#0066CC', 31: '#CCCCFF',
  32: '#000080', 33: '#FF00FF', 34: '#FFFF00', 35: '#00FFFF',
  36: '#800080', 37: '#800000', 38: '#008080', 39: '#0000FF',
  40: '#00CCFF', 41: '#CCFFFF', 42: '#CCFFCC', 43: '#FFFF99',
  44: '#99CCFF', 45: '#FF99CC', 46: '#CC99FF', 47: '#FFCC99',
  48: '#3366FF', 49: '#33CCCC', 50: '#99CC00', 51: '#FFCC00',
  52: '#FF9900', 53: '#FF6600', 54: '#666699', 55: '#969696',
  56: '#003366', 57: '#339966', 58: '#003300', 59: '#333300',
  60: '#993300', 61: '#993366', 62: '#333399', 63: '#333333',
};

/** indexed == 64（`indexed="64"`）在 Excel 里表示"系统前景色"，取黑 */
export const SYSTEM_FOREGROUND = '#000000';

/* -------------------------------------------------------------------------- */
/* 主题色                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * `a:clrScheme` 的**书写**顺序（也是 `ParsedWorkbook.themeColors` 数组的顺序）：
 * dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink。
 */
export const THEME_SLOT_ORDER = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const;

/**
 * ⚠️ 样式里 `theme="N"` 的 N **不是** `clrScheme` 的书写下标：前两个槽位整体错位。
 *
 * | `theme="N"` | 主题槽位 | Office 名称 |
 * | --- | --- | --- |
 * | 0 | lt1 | Background 1 |
 * | 1 | dk1 | Text 1 |
 * | 2 | lt2 | Background 2 |
 * | 3 | dk2 | Text 2 |
 * | 4..9 | accent1..6 | Accent 1..6 |
 * | 10 | hlink | Hyperlink |
 * | 11 | folHlink | Followed Hyperlink |
 *
 * 即：下标顺序是 `lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink`，
 * 而 `clrScheme` 的**书写**顺序是 `dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink`
 * —— 只有第 1、2 位互换，其余一一对应。
 *
 * 依据（全部来自实测真实 Excel 文件，不是推测）：
 * - 默认字体写 `<color theme="1"/>` 且渲染为**黑** → 1 = dk1；
 * - 白底纯色填充写 `<fgColor theme="0"/>` → 0 = lt1；
 * - 实测本解析器输出：`theme=2` -> lt2(#EEECE1)、`theme=3` -> dk2(#1F497D)、
 *   `theme=9` 出现在 `<fgColor>`（accent6 填充）→ 4..11 与书写顺序一一对应。
 *
 * 忽略第 1、2 位的互换会导致"黑字变白、白底变黑"这类灾难性颜色反转。
 */
const THEME_INDEX_TO_SLOT: readonly number[] = [1, 0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];

/** 取色失败时的兜底（Excel 默认主题） */
export const FALLBACK_THEME_COLORS: readonly string[] = [
  '#000000', '#FFFFFF', '#1F497D', '#EEECE1',
  '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646',
  '#0000FF', '#800080',
];

/**
 * `a:clrScheme` 的子元素顺序是 dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink。
 * 每个槽位内部取第一个 `a:srgbClr/@val` 或 `a:sysClr/@lastClr`。
 */
export function parseThemeColors(themeXml: string, warn: (msg: string) => void): string[] {
  const schemeEl = findFirstElement(themeXml, 'clrScheme');
  if (!schemeEl) {
    warn('theme1.xml 缺少 a:clrScheme，主题色回退为 Excel 默认主题');
    return [...FALLBACK_THEME_COLORS];
  }
  const children = childElements(themeXml, schemeEl);
  const byName = new Map<string, string>();
  for (const child of children) {
    const slot = localName(child.name);
    const color = readSchemeSlot(themeXml, child);
    if (color && !byName.has(slot)) byName.set(slot, color);
  }
  const colors = THEME_SLOT_ORDER.map((slot, i) => {
    const found = byName.get(slot);
    if (found) return found;
    const positional = children[i] ? readSchemeSlot(themeXml, children[i]) : undefined;
    return positional ?? FALLBACK_THEME_COLORS[i];
  });
  if (colors.length < THEME_SLOT_ORDER.length) {
    warn('theme1.xml 的 a:clrScheme 槽位不足 12 个，缺失项已用默认主题色补齐');
  }
  return colors;
}

/** 在槽位元素内部找 srgbClr / sysClr / scrgbClr，返回 `#RRGGBB` */
function readSchemeSlot(themeXml: string, slotEl: ElementRange): string | undefined {
  const kids = childElements(themeXml, slotEl);
  for (const kid of kids) {
    const ln = localName(kid.name);
    if (ln === 'srgbClr') {
      const rgb = parseHexColor(kid.attrs['val'] ?? '');
      if (rgb) return rgb;
    } else if (ln === 'sysClr') {
      // 老主题用 sysClr（如 windowText/window），lastClr 是上次渲染的缓存值
      const last = parseHexColor(kid.attrs['lastClr'] ?? '');
      if (last) return last;
      const mapped = SYSTEM_COLOR_MAP[kid.attrs['val'] ?? ''];
      if (mapped) return mapped;
    } else if (ln === 'scrgbClr') {
      const pct = (key: string): number | undefined => {
        const v = kid.attrs[key];
        if (v === undefined) return undefined;
        const n = Number(v);
        return Number.isFinite(n) ? n : undefined;
      };
      const r = pct('r'), g = pct('g'), b = pct('b');
      if (r !== undefined && g !== undefined && b !== undefined) {
        return rgbToHex(r * 255, g * 255, b * 255);
      }
    }
  }
  return undefined;
}

/** `sysClr/@val` 在没有 lastClr 时的语义回退 */
const SYSTEM_COLOR_MAP: Readonly<Record<string, string>> = {
  windowText: '#000000',
  window: '#FFFFFF',
  captionText: '#000000',
  caption: '#FFFFFF',
  highlight: '#0078D4',
  highlightText: '#FFFFFF',
  btnFace: '#F0F0F0',
  btnText: '#000000',
  grayText: '#808080',
  menuText: '#000000',
  menu: '#FFFFFF',
  infoText: '#000000',
  infoBk: '#FFFFE1',
};

/* -------------------------------------------------------------------------- */
/* 颜色工具                                                                    */
/* -------------------------------------------------------------------------- */

export function parseHexColor(raw: string): string | undefined {
  let s = raw.trim();
  if (s.startsWith('#')) s = s.slice(1);
  if (s.length === 8) s = s.slice(2); // AARRGGBB -> RRGGBB（忽略 alpha）
  if (s.length === 6 && /^[0-9a-fA-F]{6}$/.test(s)) return `#${s.toUpperCase()}`;
  if (s.length === 3 && /^[0-9a-fA-F]{3}$/.test(s)) {
    return `#${s[0]}${s[0]}${s[1]}${s[1]}${s[2]}${s[2]}`.toUpperCase();
  }
  return undefined;
}

function clamp255(v: number): number {
  return v < 0 ? 0 : v > 255 ? 255 : Math.round(v);
}

export function rgbToHex(r: number, g: number, b: number): string {
  const h = (n: number): string => clamp255(n).toString(16).padStart(2, '0').toUpperCase();
  return `#${h(r)}${h(g)}${h(b)}`;
}

export interface Hsl { h: number; s: number; l: number }

export function rgbToHsl(r: number, g: number, b: number): Hsl {
  const rn = clamp255(r) / 255;
  const gn = clamp255(g) / 255;
  const bn = clamp255(b) / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return { h: 0, s: 0, l: Math.round(l * 255) };
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === rn) h = ((gn - bn) / d) % 6;
  else if (max === gn) h = (bn - rn) / d + 2;
  else h = (rn - gn) / d + 4;
  h *= 60;
  if (h < 0) h += 360;
  // Excel 内部用 0..255 的 H/L，S 用 0..255 的整数
  return { h: Math.round(h), s: Math.round(s * 255), l: Math.round(l * 255) };
}

export function hslToHex(h: number, s: number, l: number): string {
  const hn = ((h % 360) + 360) % 360;
  const sn = Math.min(255, Math.max(0, s)) / 255;
  const ln = Math.min(255, Math.max(0, l)) / 255;
  if (sn === 0) {
    const v = Math.round(ln * 255);
    return rgbToHex(v, v, v);
  }
  const q = ln < 0.5 ? ln * (1 + sn) : ln + sn - ln * sn;
  const p = 2 * ln - q;
  const channel = (tRaw: number): number => {
    let t = tRaw;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const r = channel(hn / 360 + 1 / 3);
  const g = channel(hn / 360);
  const b = channel(hn / 360 - 1 / 3);
  return rgbToHex(r * 255, g * 255, b * 255);
}

/**
 * Excel 的 tint 变换（与 OOXML `@tint` 语义一致）：
 * - tint < 0：亮度按 (1 + tint) 缩放（变暗）
 * - tint > 0：亮度向 1（白）插值 lum * (1 - tint) + tint（变亮）
 * - 色相不变，饱和度在变亮时按 (1 - tint) 缩放
 */
export function applyTint(hex: string, tint: number): string {
  if (!Number.isFinite(tint) || tint === 0) return hex;
  const m = /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/.exec(hex.trim());
  if (!m) return hex;
  const r = Number.parseInt(m[1], 16);
  const g = Number.parseInt(m[2], 16);
  const b = Number.parseInt(m[3], 16);
  const hsl = rgbToHsl(r, g, b);
  const t = Math.max(-1, Math.min(1, tint));
  if (t < 0) {
    hsl.l = Math.round(hsl.l * (1 + t));
  } else {
    hsl.l = Math.round(hsl.l * (1 - t) + 255 * t);
    hsl.s = Math.round(hsl.s * (1 - t));
  }
  return hslToHex(hsl.h, hsl.s, hsl.l);
}

export interface ColorContext {
  themeColors?: readonly string[];
  warn: (msg: string) => void;
}
/**
 * 解析一个颜色元素（`<color>` / `<fgColor>` / `<bgColor>` / `<font><color>`）的属性。
 * 支持 rgb / indexed / theme+tint / auto。
 */
export function parseColorElement(attrs: XmlAttributes, ctx: ColorContext): string | undefined {
  const rgb = attrs['rgb'];
  if (rgb !== undefined && rgb !== '') {
    const parsed = parseHexColor(rgb);
    if (parsed) return applyTint(parsed, numericTint(attrs));
    ctx.warn(`无法解析颜色 rgb="${rgb}"，已忽略该颜色`);
    return undefined;
  }

  const indexed = attrs['indexed'];
  if (indexed !== undefined && indexed !== '') {
    const idx = Number.parseInt(indexed, 10);
    if (!Number.isFinite(idx)) {
      ctx.warn(`无法解析颜色 indexed="${indexed}"，已忽略该颜色`);
      return undefined;
    }
    if (idx === 64) return applyTint(SYSTEM_FOREGROUND, numericTint(attrs));
    const palette = INDEXED_COLORS[idx];
    if (palette) return applyTint(palette, numericTint(attrs));
    ctx.warn(`indexed 颜色索引 ${idx} 不在内置调色板(0-63)中，颜色回退为 #000000`);
    return applyTint('#000000', numericTint(attrs));
  }

  const theme = attrs['theme'];
  if (theme !== undefined && theme !== '') {
    const idx = Number.parseInt(theme, 10);
    const palette = ctx.themeColors ?? FALLBACK_THEME_COLORS;
    // 注意：样式的 theme 下标需要经过 THEME_INDEX_TO_SLOT 折算（0/1、2/3 与书写顺序互换）
    const slot = Number.isFinite(idx) && idx >= 0 && idx < THEME_INDEX_TO_SLOT.length
      ? THEME_INDEX_TO_SLOT[idx]
      : undefined;
    const base = slot !== undefined && slot < palette.length ? palette[slot] : undefined;
    if (!base) {
      ctx.warn(`theme 颜色索引 ${theme} 超出主题色范围，颜色回退为 #000000`);
      return applyTint('#000000', numericTint(attrs));
    }
    return applyTint(base, numericTint(attrs));
  }

  if (attrs['auto'] !== undefined) return '#000000';
  return undefined;
}

function numericTint(attrs: XmlAttributes): number {
  const raw = attrs['tint'];
  if (raw === undefined || raw === '') return 0;
  const t = Number(raw);
  return Number.isFinite(t) ? t : 0;
}

/* -------------------------------------------------------------------------- */
/* 边框 / 字体 / 填充 / 对齐                                                   */
/* -------------------------------------------------------------------------- */

/** OOXML 边框线型 -> 契约里的 13 种（其中 hair/slantDashDot 等保持原样） */
const BORDER_STYLE_MAP: Readonly<Record<string, BorderLineStyle>> = {
  thin: 'thin',
  medium: 'medium',
  thick: 'thick',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  hair: 'hair',
  dashDot: 'dashDot',
  dashDotDot: 'dashDotDot',
  mediumDashed: 'mediumDashed',
  mediumDashDot: 'mediumDashDot',
  mediumDashDotDot: 'mediumDashDotDot',
  slantDashDot: 'slantDashDot',
};

export function mapBorderStyle(raw: string | undefined): BorderLineStyle | undefined {
  if (!raw) return undefined;
  return BORDER_STYLE_MAP[raw];
}

function parseBorderChild(xml: string, el: ElementRange, ctx: ColorContext): ParsedBorder | undefined {
  const style = mapBorderStyle(el.attrs['style']);
  const colorEl = childElements(xml, el).find((c) => nameMatches(c.name, 'color'));
  const color = colorEl ? parseColorElement(colorEl.attrs, ctx) : undefined;
  if (!style) {
    // 有颜色没线型（Excel 里常见）——契约要求 style 必填，只能降级为 thin
    if (color) {
      ctx.warn(`边框线型 "${el.attrs['style'] ?? ''}" 无法识别，已按 thin 处理`);
      return { style: 'thin', color };
    }
    return undefined;
  }
  return color ? { style, color } : { style };
}

function parseBorder(xml: string, el: ElementRange, ctx: ColorContext): ParsedStyle['border'] | undefined {
  const border: NonNullable<ParsedStyle['border']> = {};
  let any = false;
  for (const side of childElements(xml, el)) {
    const ln = localName(side.name);
    if (ln !== 'left' && ln !== 'right' && ln !== 'top' && ln !== 'bottom') continue;
    const parsed = parseBorderChild(xml, side, ctx);
    if (!parsed) continue;
    border[ln] = parsed;
    any = true;
  }
  return any ? border : undefined;
}

/** 字体在 `cellXfs` 之前先摊平成 `ParsedStyle` 的子集，最后再合并到 xf */
type FontDef = Omit<
  ParsedStyle,
  'fill' | 'border' | 'horizontalAlign' | 'verticalAlign' | 'textWrap' | 'indent' | 'textRotation' | 'numberFormat'
>;

function parseFont(xml: string, el: ElementRange, ctx: ColorContext): FontDef {
  const font: FontDef = {};
  for (const kid of childElements(xml, el)) {
    const ln = localName(kid.name);
    const val = kid.attrs['val'];
    switch (ln) {
      case 'name':
        if (val) font.fontFamily = decodeEntities(val);
        break;
      case 'sz': {
        const n = attrNumber(kid.attrs, 'val');
        if (n !== undefined) font.fontSize = n;
        break;
      }
      case 'b': {
        const raw = val;
        if (raw === undefined) font.bold = true;
        else font.bold = attrBool(kid.attrs, 'val') ?? false;
        break;
      }
      case 'i': {
        const raw = val;
        if (raw === undefined) font.italic = true;
        else font.italic = attrBool(kid.attrs, 'val') ?? false;
        break;
      }
      case 'u': {
        const raw = val;
        if (raw === undefined) font.underline = true;
        else font.underline = (attrBool(kid.attrs, 'val') ?? false) || raw === 'single' || raw === 'double';
        break;
      }
      case 'strike': {
        const raw = val;
        if (raw === undefined) font.strikeThrough = true;
        else font.strikeThrough = attrBool(kid.attrs, 'val') ?? false;
        break;
      }
      case 'color':
        font.color = parseColorElement(kid.attrs, ctx);
        break;
      default:
        break;
    }
  }
  return font;
}

/** `parseFill` 的结果：能画的填充色，以及"这个填充我们画不出来"的原因（延迟到确认被引用后才记 warning） */
interface ParsedFillResult {
  fill?: string;
  issue?: string;
}

function parseFill(xml: string, el: ElementRange, ctx: ColorContext): ParsedFillResult {
  for (const kid of childElements(xml, el)) {
    const ln = localName(kid.name);
    if (ln === 'patternFill') {
      const pattern = kid.attrs['patternType'] ?? 'none';
      if (pattern === 'none') return {};
      if (pattern === 'solid') {
        const kids = childElements(xml, kid);
        const fg = kids.find((c) => nameMatches(c.name, 'fgColor'));
        const bg = kids.find((c) => nameMatches(c.name, 'bgColor'));
        // Excel 写出的 solid 里，真正生效的是 fgColor；fgColor 缺失时退回 bgColor
        const color = (fg ? parseColorElement(fg.attrs, ctx) : undefined)
          ?? (bg ? parseColorElement(bg.attrs, ctx) : undefined);
        return color ? { fill: color } : {};
      }
      /**
       * `gray125` 是 **ECMA-376 里与 `none` 并列的默认占位填充**：Excel/WPS 几乎每个文件都会
       * 写一个 `fills[1] = gray125`，而绝大多数文件根本没有单元格引用它。
       * 以前这里当场记 warning，于是用户导入座位表时状态栏无端出现"降级 1 项"（实测反馈）。
       * 现在只**记录**原因，等 cellXfs 解析完、确认真的被某个单元格样式引用时才上报。
       */
      if (pattern === 'gray125') {
        return { issue: '填充 patternType="gray125"（12.5% 灰点阵）未渲染，该格按无填充显示（本工具只画纯色填充）' };
      }
      return { issue: `填充 patternType="${pattern}" 未渲染，该格按无填充显示（本工具只画 solid 纯色填充）` };
    }
    if (ln === 'gradientFill') {
      return { issue: '渐变填充(gradientFill)未渲染，该格按无填充显示（本工具只画 solid 纯色填充）' };
    }
  }
  return {};
}

function parseAlignment(xml: string, xfEl: ElementRange, style: ParsedStyle): void {
  const alignEl = childElements(xml, xfEl).find((c) => nameMatches(c.name, 'alignment'));
  if (!alignEl) return;
  const a = alignEl.attrs;

  const h = a['horizontal'];
  if (h === 'left' || h === 'center' || h === 'right') style.horizontalAlign = h;
  else if (h === 'start') style.horizontalAlign = 'left';
  else if (h === 'end') style.horizontalAlign = 'right';
  else if (h === 'centerContinuous') style.horizontalAlign = 'center';
  else if (h === 'justify' || h === 'distributed' || h === 'fill') style.horizontalAlign = 'left';

  const v = a['vertical'];
  if (v === 'top' || v === 'bottom') style.verticalAlign = v;
  else if (v === 'center') style.verticalAlign = 'middle';
  else if (v === 'justify' || v === 'distributed') style.verticalAlign = 'top';

  const wrap = attrBool(a, 'wrapText');
  if (wrap !== undefined) style.textWrap = wrap;

  const indent = attrInt(a, 'indent');
  if (indent !== undefined && indent !== 0) style.indent = indent;

  const rotation = attrInt(a, 'textRotation');
  if (rotation !== undefined && rotation !== 0) style.textRotation = rotation;
}

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 按**本地名精确**查找第一个元素。
 *
 * `findFirstElement` 用 `indexOf` 顺序扫描，而 `cellStyleXfs` 的后缀恰好是 `cellXfs`，
 * 因此像 `<cellStyleXfs>` 排在 `<cellXfs>` 之前的文档会命中错的容器。这里多校验一次精确名。
 */
function findExactElement(xml: string, local: string): ElementRange | undefined {
  let found: ElementRange | undefined;
  visitElements(xml, (el) => {
    if (localName(el.name) !== local) return; // 继续找
    found = el;
    return false; // 命中即剪枝
  });
  return found;
}

/* -------------------------------------------------------------------------- */
/* 条件格式差异样式：<dxfs>                                                     */
/* -------------------------------------------------------------------------- */

/**
 * `xl/styles.xml` 的 `<dxfs>` -> `CfDxfStyle[]`（下标 == `cfRule/@dxfId`）。
 *
 * 与普通 `<fill>` 最容易踩的坑：**dxf 的 solid 填充底色写在 `<bgColor>`**，
 * 而普通 fill 的 solid 底色写在 `<fgColor>`（两者语义相反）。
 * 实测 fixture-rules.xlsx：`<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/>`
 * 在 Excel 里渲染成浅红底；若照普通 fill 取 fgColor，会拿到"未写 = 黑色"，
 * 于是"浅红底红字"变成"黑底红字"。因此这里**以 bgColor 为准**（bgColor 缺失才退回 fgColor）。
 *
 * 只保留契约里有的字段（bold/italic/strikeThrough/color/fill/border），
 * 其余（`<numFmt>`、`<alignment>`、`<protection>`、`<u>`）忽略；空的 `<dxf/>` 也占一个下标，
 * 保证 `dxfId` 与数组下标严格对齐。
 */
export function parseDxfStyles(stylesXml: string | undefined, opts: ParseStylesOptions): CfDxfStyle[] {
  if (!stylesXml) return [];
  const holder = findExactElement(stylesXml, 'dxfs');
  if (!holder) return [];
  const ctx: ColorContext = { themeColors: opts.themeColors, warn: opts.warn };
  const out: CfDxfStyle[] = [];
  for (const dxf of childElements(stylesXml, holder)) {
    if (localName(dxf.name) !== 'dxf') continue;
    out.push(parseDxf(stylesXml, dxf, ctx));
  }
  return out;
}

function parseDxf(xml: string, dxf: ElementRange, ctx: ColorContext): CfDxfStyle {
  const style: CfDxfStyle = {};
  for (const kid of childElements(xml, dxf)) {
    const ln = localName(kid.name);
    if (ln === 'font') {
      const font = parseFont(xml, kid, ctx);
      if (font.bold !== undefined) style.bold = font.bold;
      if (font.italic !== undefined) style.italic = font.italic;
      if (font.strikeThrough !== undefined) style.strikeThrough = font.strikeThrough;
      if (font.color !== undefined) style.color = font.color;
    } else if (ln === 'fill') {
      const fill = parseDxfFill(xml, kid, ctx);
      if (fill !== undefined) style.fill = fill;
    } else if (ln === 'border') {
      const border = parseBorder(xml, kid, ctx);
      if (border !== undefined) style.border = border;
    }
  }
  return style;
}

/** dxf 的 `<fill>`：solid 时底色取 `<bgColor>`（与普通 fill 相反） */
function parseDxfFill(xml: string, fillEl: ElementRange, ctx: ColorContext): string | undefined {
  for (const kid of childElements(xml, fillEl)) {
    const ln = localName(kid.name);
    if (ln === 'patternFill') {
      const pattern = kid.attrs['patternType'] ?? 'none';
      if (pattern === 'none') return undefined;
      if (pattern !== 'solid') {
        ctx.warn(`条件格式 dxf 的 patternFill patternType="${pattern}" 暂不支持（只处理 solid 纯色）`);
        return undefined;
      }
      const kids = childElements(xml, kid);
      const bg = kids.find((c) => nameMatches(c.name, 'bgColor'));
      const fg = kids.find((c) => nameMatches(c.name, 'fgColor'));
      return (bg ? parseColorElement(bg.attrs, ctx) : undefined)
        ?? (fg ? parseColorElement(fg.attrs, ctx) : undefined);
    }
    if (ln === 'gradientFill') {
      ctx.warn('条件格式 dxf 的渐变填充(gradientFill)暂不支持');
      return undefined;
    }
  }
  return undefined;
}

export interface ParseStylesOptions {
  themeColors?: readonly string[];
  warn: (msg: string) => void;
}

/** `xl/styles.xml` -> `ParsedStyle[]`（索引 == cellXfs 下标） */
export function parseStyles(stylesXml: string | undefined, opts: ParseStylesOptions): ParsedStyle[] {
  if (!stylesXml) return [];
  const ctx: ColorContext = { themeColors: opts.themeColors, warn: opts.warn };
  const customFmts = new Map<number, string>();
  const fonts: Array<Omit<ParsedStyle, 'fill' | 'border'>> = [];
  const fills: Array<string | undefined> = [];
  const borders: Array<ParsedStyle['border']> = [];
  const xfDefs: Array<{ rawStyle: XfRawStyle; source: string }> = [];

  const numFmtEl = findFirstElement(stylesXml, 'numFmts');
  if (numFmtEl) {
    for (const fmt of childElements(stylesXml, numFmtEl)) {
      if (!nameMatches(fmt.name, 'numFmt')) continue;
      const id = attrInt(fmt.attrs, 'numFmtId');
      const code = fmt.attrs['formatCode'];
      if (id === undefined || code === undefined) {
        opts.warn(`styles.xml 中存在缺少 numFmtId/formatCode 的 numFmt，已跳过`);
        continue;
      }
      customFmts.set(id, decodeEntities(code));
    }
  }

  const fontsEl = findFirstElement(stylesXml, 'fonts');
  if (fontsEl) {
    for (const f of childElements(stylesXml, fontsEl)) {
      if (!nameMatches(f.name, 'font')) continue;
      fonts.push(parseFont(stylesXml, f, ctx));
    }
  }

  const fillsEl = findFirstElement(stylesXml, 'fills');
  /** 与 `fills` 同下标：该填充"解析不了"的原因（只有真被 cellXfs 引用时才上报，见文件末尾） */
  const fillIssues: Array<string | undefined> = [];
  if (fillsEl) {
    for (const f of childElements(stylesXml, fillsEl)) {
      if (!nameMatches(f.name, 'fill')) continue;
      const parsed = parseFill(stylesXml, f, ctx);
      fills.push(parsed.fill);
      fillIssues.push(parsed.issue);
    }
  }

  const bordersEl = findFirstElement(stylesXml, 'borders');
  if (bordersEl) {
    for (const b of childElements(stylesXml, bordersEl)) {
      if (!nameMatches(b.name, 'border')) continue;
      borders.push(parseBorder(stylesXml, b, ctx));
    }
  }

  /**
   * 注意：**不能**用 `findFirstElement(xml, 'cellXfs')` —— 它会先命中 `cellStyleXfs`
   * （`cellStyleXfs` 的后缀正是 `cellXfs`）。必须精确区分这两个容器：
   * 只有 `cellXfs` 的下标才是单元格 `c/@s` 引用的样式索引。
   */
  for (const holderName of ['cellStyleXfs', 'cellXfs'] as const) {
    const holder = findExactElement(stylesXml, holderName);
    if (!holder) continue;
    for (const xf of childElements(stylesXml, holder)) {
      if (localName(xf.name) !== 'xf') continue;
      xfDefs.push({ rawStyle: buildXfStyle(stylesXml, xf, customFmts), source: holderName });
    }
  }

  const cellXfDefs = xfDefs.filter((d) => d.source === 'cellXfs');

  /**
   * 只对**真的被用到**的填充上报"画不出来"。
   *
   * 用户实测反馈：导入座位表时状态栏出现"降级 1 项"，点开一看是
   * `patternFill patternType="gray125" 暂不支持` —— 而这个 gray125 是 Excel/WPS 写进
   * `fills[1]` 的规范占位填充，全表没有任何 `xf` 引用它（实测：25 个 cellXfs 全是 fillId=0）。
   * 也就是说这条"降级"完全不影响预览，纯属噪音。现在改成按引用关系上报：
   * 没被引用的填充只是躺在 styles.xml 里，既不画也不报警。
   */
  const referencedFills = new Set<number>();
  for (const def of cellXfDefs) {
    if (def.rawStyle.fillId !== undefined) referencedFills.add(def.rawStyle.fillId);
  }
  const reported = new Set<string>();
  for (const id of [...referencedFills].sort((a, b) => a - b)) {
    const issue = fillIssues[id];
    if (issue && !reported.has(issue)) {
      reported.add(issue);
      opts.warn(`cellXfs 引用了填充 #${id}：${issue}`);
    }
  }

  return cellXfDefs.map(({ rawStyle }) => applyParts(rawStyle, fonts, fills, borders, opts.warn));
}

/** 把 `xf` 的引用关系摊平成中间形态（仍保留 numFmtId 等原始字段） */
interface XfRawStyle extends ParsedStyle {
  fontId?: number;
  fillId?: number;
  borderId?: number;
}

function buildXfStyle(xml: string, xf: ElementRange, customFmts: Map<number, string>): XfRawStyle {
  const style: XfRawStyle = {};
  const fontId = attrInt(xf.attrs, 'fontId');
  const fillId = attrInt(xf.attrs, 'fillId');
  const borderId = attrInt(xf.attrs, 'borderId');
  if (fontId !== undefined) style.fontId = fontId;
  if (fillId !== undefined) style.fillId = fillId;
  if (borderId !== undefined) style.borderId = borderId;
  const numFmtId = attrInt(xf.attrs, 'numFmtId');
  if (numFmtId !== undefined) {
    const pattern = customFmts.get(numFmtId) ?? BUILTIN_NUM_FMTS[numFmtId];
    if (pattern !== undefined) style.numberFormat = pattern;
  }
  parseAlignment(xml, xf, style);
  return style;
}

function applyParts(
  raw: XfRawStyle,
  fonts: Array<Omit<ParsedStyle, 'fill' | 'border'>>,
  fills: Array<string | undefined>,
  borders: Array<ParsedStyle['border']>,
  warn: (msg: string) => void,
): ParsedStyle {
  const out: ParsedStyle = {};
  const { fontId, fillId, borderId } = raw;

  if (fontId !== undefined) {
    const font = fonts[fontId];
    if (font) Object.assign(out, font);
    else if (fontId !== 0) warn(`cellXfs 引用了不存在的 fontId=${fontId}，已忽略字体`);
  }
  if (fillId !== undefined) {
    if (fillId >= 2) {
      const fill = fills[fillId];
      if (fill) out.fill = fill;
      else warn(`cellXfs 引用了不存在的 fillId=${fillId}，已忽略填充`);
    }
    // fillId 0/1 是规范定义的无填充（none / gray125），不需要记 warning
  }
  if (borderId !== undefined) {
    const border = borders[borderId];
    if (border) out.border = border;
    else if (borderId !== 0) warn(`cellXfs 引用了不存在的 borderId=${borderId}，已忽略边框`);
  }
  if (raw.numberFormat !== undefined) out.numberFormat = raw.numberFormat;
  if (raw.horizontalAlign !== undefined) out.horizontalAlign = raw.horizontalAlign;
  if (raw.verticalAlign !== undefined) out.verticalAlign = raw.verticalAlign;
  if (raw.textWrap !== undefined) out.textWrap = raw.textWrap;
  if (raw.indent !== undefined) out.indent = raw.indent;
  if (raw.textRotation !== undefined) out.textRotation = raw.textRotation;
  return out;
}
