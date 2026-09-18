/**
 * CSV / TSV 导出（**自研**，不走轮子也不走桥）。
 *
 * 为什么自研：这三种格式本来就不承载样式，而"给别的系统吃"的场合最在乎的是
 * ① 编码（Excel 打开中文不错行 → 加 UTF-8 BOM）② 前导零（学号/电话不能被吃）③ 分隔符与引号规则。
 * 这些恰好是我们已经写熟的东西（导入侧的 `parseDelimitedText` 是同一套口径）。
 */

/** CSV 里一个字段要不要加引号：含分隔符/引号/换行就必须加（RFC4180） */
function needsQuotes(text: string, delimiter: string): boolean {
  return text.includes(delimiter) || text.includes('"') || /[\r\n]/.test(text);
}

/** 单个字段的转义（内部引号翻倍） */
export function csvField(text: string, delimiter = ','): string {
  const value = text ?? '';
  if (!needsQuotes(value, delimiter)) return value;
  return `"${value.replace(/"/g, '""')}"`;
}

/**
 * 把二维文本拼成 CSV/TSV 文本。
 * 行尾统一 `\r\n`（Excel 与多数系统都认），最后一行也带行尾 —— 与 Excel 自己的导出一致。
 */
export function buildDelimited(rows: string[][], delimiter = ','): string {
  return rows.map((row) => row.map((cell) => csvField(cell ?? '', delimiter)).join(delimiter)).join('\r\n') + '\r\n';
}

/** 编码成字节：默认带 UTF-8 BOM（Excel 双击打开不乱码） */
export function encodeDelimited(text: string, withBom = true): Uint8Array {
  const body = new TextEncoder().encode(text);
  if (!withBom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf], 0);
  out.set(body, 3);
  return out;
}
