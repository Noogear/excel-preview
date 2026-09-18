/**
 * **CSV / TSV / 任意分隔文本** 解析器（用户要求："把其他表格格式都加上去"）。
 *
 * 产出中性的 `WorkbookInput`（见 `src/importer/synth-xlsx.ts`），随后被打成一份规范 xlsx，
 * 于是预览/编辑/工作区/撤销/导出/会话恢复全部复用既有链路——CSV 只是一条"入口翻译"。
 *
 * 工程上的三个坑（都按真实文件处理，不做假设）：
 *  1. **编码**：中文 Windows 的 Excel「另存为 CSV」写出的是 **GBK（ANSI）**，网页里直接按 UTF-8 解会变乱码。
 *     顺序：BOM（UTF-8 / UTF-16）→ 严格 UTF-8 试解 → 退回 GBK。解不出来时宁可给"识别为 GBK"，
 *     也不静默产出乱码（`encoding` 会写进导入摘要）。
 *  2. **分隔符**：逗号 / 分号 / 制表符 / 竖线都有人用（欧洲区域设置把逗号当小数点，于是用分号）。
 *     策略：`.tsv` 直接制表符；否则在前若干非空行里统计"引号外"的候选符出现次数，
 *     取"每行次数一致且最多"的那个。
 *  3. **引号规则（RFC 4180）**：字段可被 `"` 包起来，内部 `""` 表示一个 `"`，
 *     字段里可以有换行与分隔符 —— 用状态机一次扫完，不用正则（正则处理不了嵌套引号）。
 *
 * 类型判定（保守、可预期）：
 *  - 纯数字（含负号、小数、科学计数法）→ 数值；
 *  - **带前导零**（`007`、`0912`、`0571…`）→ **保留为文本**（学号/电话/编号的零不能被吃掉）；
 *  - `TRUE/FALSE`、日期文本等一律**按文本**保留（CSV 本来没有类型，不替用户猜）。
 */
import type { SynthCell, SynthSheet, WorkbookInput } from '../importer/synth-xlsx';

/** 单次导入的安全阀：单元格总量上限（超出截断并如实报告） */
export const MAX_DELIMITED_CELLS = 200_000;
/** 单边上限（与 xlsx 规范一致） */
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

export interface DelimitedParseResult {
  input: WorkbookInput;
  /** 实际使用的分隔符（中文名给摘要用） */
  delimiter: string;
  /** 实际使用的编码（utf-8 / utf-8-bom / utf-16le / gbk） */
  encoding: string;
  rows: number;
  cols: number;
  /** 是否因为超过上限被截断 */
  truncated: boolean;
  /** 给"导入摘要"用的人话（不是错误，只是说明） */
  notes: string[];
}

/* -------------------------------------------------------------------------- */
/* 编码                                                                        */
/* -------------------------------------------------------------------------- */

function decodeWith(label: string, bytes: Uint8Array, fatal: boolean): string | null {
  try {
    return new TextDecoder(label, { fatal }).decode(bytes);
  } catch {
    return null;
  }
}

/**
 * 猜编码并解码。
 *
 * 为什么"严格 UTF-8"这一招管用：GBK 字节序列极少能通过 UTF-8 的非法序列检查
 * （实测 `fixture-styles-gbk.csv` 直接抛错），所以"能严格解开就是 UTF-8"是可靠的判据；
 * 反过来（把 UTF-8 当 GBK 解）才会乱码，而我们不会先试 GBK。
 */
export function decodeDelimited(bytes: Uint8Array): { text: string; encoding: string } {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { text: decodeWith('utf-8', bytes.subarray(3), false) ?? '', encoding: 'utf-8-bom' };
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { text: decodeWith('utf-16le', bytes.subarray(2), false) ?? '', encoding: 'utf-16le' };
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    // TextDecoder 没有 utf-16be：交换字节序后按 LE 解（BOM 已去掉）
    const swapped = new Uint8Array(bytes.length - 2);
    for (let i = 2; i + 1 < bytes.length; i += 2) {
      swapped[i - 2] = bytes[i + 1];
      swapped[i - 1] = bytes[i];
    }
    return { text: decodeWith('utf-16le', swapped, false) ?? '', encoding: 'utf-16be' };
  }
  const utf8 = decodeWith('utf-8', bytes, true);
  if (utf8 !== null) return { text: utf8, encoding: 'utf-8' };
  const gbk = decodeWith('gbk', bytes, false);
  if (gbk !== null) return { text: gbk, encoding: 'gbk' };
  // 连 GBK 都读不了（例如 latin1 的二进制噪声）：按 latin1 兜住，至少不崩
  return { text: decodeWith('latin1', bytes, false) ?? '', encoding: 'latin1' };
}

/* -------------------------------------------------------------------------- */
/* 分隔符                                                                      */
/* -------------------------------------------------------------------------- */

/** 数一数"引号外"某个字符出现几次（引号内的分隔符不算分隔符） */
function countOutsideQuotes(line: string, char: string): number {
  let count = 0;
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && ch === char) count += 1;
  }
  return count;
}

const DELIMITER_NAMES: Readonly<Record<string, string>> = {
  ',': '逗号',
  ';': '分号',
  '\t': '制表符',
  '|': '竖线',
};

export function delimiterLabel(delimiter: string): string {
  return DELIMITER_NAMES[delimiter] ?? delimiter;
}

/**
 * 猜分隔符：候选里挑"每行出现次数一致、且次数最多"的那个。
 * 一致性的权重高于次数本身——否则一个含大量逗号的备注列会把制表符表格带偏。
 */
export function sniffDelimiter(text: string, fileName = ''): string {
  if (/\.tsv$/i.test(fileName)) return '\t';
  const lines = text
    .split(/\r\n|\n|\r/)
    .filter((line) => line.trim() !== '')
    .slice(0, 20);
  if (lines.length === 0) return ',';

  let best = ',';
  let bestScore = -1;
  for (const candidate of [',', ';', '\t', '|']) {
    const counts = lines.map((line) => countOutsideQuotes(line, candidate));
    const positive = counts.filter((n) => n > 0);
    if (positive.length === 0) continue;
    const first = counts[0];
    const consistent = counts.every((n) => n === first);
    // 一致 → 以第一行次数为主分；不一致 → 用"出现次数 >0 的行数"做弱分
    const score = consistent ? 1000 + first : positive.length;
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return best;
}

/* -------------------------------------------------------------------------- */
/* 解析（RFC 4180 状态机）                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 按分隔符切成"行 × 字段"。引号内的换行/分隔符原样保留；
 * `""` 在引号内表示一个字面量 `"`；CRLF / LF / CR 都当行分隔。
 */
export function parseDelimitedRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"' && field === '') {
      // 只有"字段开头"的引号才是包裹引号（`a"b` 里的引号按字面量处理，与 Excel 一致）
      inQuotes = true;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      continue;
    }
    field += ch;
  }
  // 收尾：最后一行没有换行符时也要收进来（空文件除外）
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** 纯数字（含负号/小数/科学计数法）→ number；带前导零、`+`、千分位、日期等一律留作文本 */
export function parseDelimitedScalar(text: string): string | number {
  const trimmed = text.trim();
  if (trimmed === '') return '';
  if (!/^-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) return text;
  // 前导零（007 / 0912 / 0.5 不算）→ 文本：学号、电话、编号的零不能被吃掉
  if (/^-?0\d/.test(trimmed)) return text;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : text;
}

/** 从文件名取工作表名（去扩展名；空则给"数据"） */
export function sheetNameFromFile(fileName: string): string {
  const stem = (fileName ?? '').replace(/\.[^./\\]+$/, '').trim();
  return stem || '数据';
}

/**
 * 主入口：字节 + 文件名 → 中性工作簿（单表）。
 *
 * 说明：CSV 没有"第二个工作表"的概念，所以永远只有一张表；多表的需求请用 .xlsx/.ods/.xls。
 */
export function parseDelimitedText(bytes: Uint8Array, fileName = ''): DelimitedParseResult {
  const { text, encoding } = decodeDelimited(bytes);
  const delimiter = sniffDelimiter(text, fileName);
  const rows = parseDelimitedRows(text, delimiter);

  // 去掉尾部整行为空的记录（Excel 导出的 CSV 常带一个空尾行）
  let end = rows.length;
  while (end > 0 && rows[end - 1].every((cell) => cell.trim() === '')) end -= 1;
  const trimmedRows = rows.slice(0, end);

  const width = trimmedRows.reduce((max, row) => Math.max(max, row.length), 0);
  const notes: string[] = [];
  let truncated = false;

  const cells: SynthCell[] = [];
  let written = 0;
  for (let r = 0; r < trimmedRows.length && r < MAX_ROWS; r += 1) {
    const row = trimmedRows[r];
    for (let c = 0; c < Math.min(width, MAX_COLS); c += 1) {
      const raw = row[c] ?? '';
      if (raw === '') continue; // 空字段不写单元格（下游也会跳过空格子，这里省一遍）
      if (written >= MAX_DELIMITED_CELLS) {
        truncated = true;
        break;
      }
      const value = parseDelimitedScalar(raw);
      if (value === '') continue;
      cells.push({ row: r, col: c, value });
      written += 1;
    }
    if (truncated) break;
  }

  if (encoding === 'gbk') {
    notes.push('CSV 编码识别为 GBK（中文 Windows 的 Excel 默认用 ANSI/GBK 导出），已按 GBK 正确解码');
  }
  if (truncated) {
    notes.push(`CSV 单元格数超过上限（${MAX_DELIMITED_CELLS}），超出部分未导入`);
  }

  const sheet: SynthSheet = { name: sheetNameFromFile(fileName), cells };
  return {
    input: { name: sheetNameFromFile(fileName), sheets: [sheet] },
    delimiter,
    encoding,
    rows: Math.min(trimmedRows.length, MAX_ROWS),
    cols: Math.min(width, MAX_COLS),
    truncated,
    notes,
  };
}
