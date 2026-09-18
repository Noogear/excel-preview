/**
 * Excel 表格（ListObject）样式实体化。
 *
 * 背景（实测确认）：表格的**表头样式与斑马纹不在单元格里**，而是由
 * `xl/tables/tableN.xml` 的 `tableStyleInfo@name` 指向 Excel 内置表格样式目录
 * （该目录不在文件里，`tableStyles.xml` 的 count 为 0）。
 * 因此要"看起来和 Excel 一样"，只能我们自己按样式名还原出配色并写进单元格样式。
 *
 * 还原策略（**算法近似**，已记入导入报告的降级项）：
 *  - 强调色：从 theme1.xml 的 accent1..6 取，序号映射 `accent = floor((N-1)/7) + 1`
 *    （Excel 的 Medium/Dark 每 7 个变体共用一个强调色）
 *  - Medium/Dark：表头 = 强调色实底 + 白字加粗；斑马纹 = 强调色极浅色调
 *  - Light：表头 = 强调色极浅底 + 强调色深字；斑马纹 = 强调色极浅色调（Light 家族本就淡雅）
 *  - 汇总行：加粗 + 上边框；首列/末列：加粗
 * 后续若要 100% 保真，需要引入 Excel 内置表格样式目录（约 60 个样式 × 每样式若干配色）作为数据表。
 */
import type { ParsedStyle, ParsedTable } from '../parser/types';
import { applyTint } from '../parser/styles';

export interface StylePatch {
  row: number;
  col: number;
  patch: ParsedStyle;
}

const FALLBACK_ACCENT = '#4472C4';

/** 解析 A1 记号（支持 `A1:D5` 与 `A1`） */
export function parseA1Range(ref: string): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(ref.trim());
  if (!match) return null;
  const startCol = columnToIndex(match[1]);
  const startRow = Number(match[2]) - 1;
  const endCol = match[3] ? columnToIndex(match[3]) : startCol;
  const endRow = match[4] ? Number(match[4]) - 1 : startRow;
  return {
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

function columnToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function buildTableStylePatches(tables: ParsedTable[] | undefined, themeColors: string[] | undefined): StylePatch[] {
  if (!tables || tables.length === 0) return [];

  const patches: StylePatch[] = [];
  for (const table of tables) {
    const range = parseA1Range(table.ref);
    if (!range) continue;
    patches.push(...buildOneTable(table, range, themeColors));
  }
  return patches;
}

function buildOneTable(
  table: ParsedTable,
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
  themeColors: string[] | undefined,
): StylePatch[] {
  const styleName = table.styleName ?? 'TableStyleMedium2';
  const family = /^TableStyle(Light|Medium|Dark)/.exec(styleName)?.[1]?.toLowerCase() ?? 'medium';
  const variantNumber = Number(/(\d+)$/.exec(styleName)?.[1] ?? '2');
  const accents = collectAccents(themeColors);
  const accentIndex = Math.min(accents.length - 1, Math.max(0, Math.floor((variantNumber - 1) / 7)));
  const accent = accents[accentIndex] ?? FALLBACK_ACCENT;

  const headerRows = table.headerRowCount ?? 1;
  const totalsRows = table.totalsRowCount ?? 0;
  const bodyStart = range.startRow + headerRows;
  const bodyEnd = range.endRow - totalsRows;

  const headerFill = family === 'light' ? applyTint(accent, 0.82) : accent;
  const headerColor = family === 'light' ? applyTint(accent, -0.3) : pickReadableTextColor(headerFill);
  const bandFill = applyTint(accent, 0.86);
  const totalsFill = family === 'dark' ? applyTint(accent, 0.1) : applyTint(accent, 0.72);

  const patches: StylePatch[] = [];

  for (let row = range.startRow; row <= range.endRow; row++) {
    const isHeader = row < bodyStart;
    const isTotals = totalsRows > 0 && row > bodyEnd;
    const bodyIndex = row - bodyStart;
    const isBanded = !!table.showRowStripes && !isHeader && !isTotals && bodyIndex % 2 === 1;

    for (let col = range.startCol; col <= range.endCol; col++) {
      const isFirst = table.showFirstColumn && col === range.startCol;
      const isLast = table.showLastColumn && col === range.endCol;
      const patch: ParsedStyle = {};

      if (isHeader) {
        patch.fill = headerFill;
        patch.color = headerColor;
        patch.bold = true;
        if (family !== 'light') patch.bold = true;
      } else if (isTotals) {
        patch.fill = totalsFill;
        patch.bold = true;
        patch.border = { top: { style: 'thin', color: accent } };
      } else if (isBanded) {
        patch.fill = bandFill;
      }

      if ((isFirst || isLast) && !isHeader) patch.bold = true;

      if (Object.keys(patch).length > 0) patches.push({ row, col, patch });
    }
  }

  return patches;
}

/** themeColors 按写入顺序（dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink） */
function collectAccents(themeColors: string[] | undefined): string[] {
  if (!themeColors || themeColors.length < 10) return [FALLBACK_ACCENT];
  return themeColors.slice(4, 10);
}

/** 依据底色亮度选择黑/白字，避免"浅底白字"看不清 */
export function pickReadableTextColor(background: string | undefined): string {
  if (!background) return '#000000';
  const hex = background.replace('#', '');
  if (hex.length < 6) return '#000000';
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  // 相对亮度（sRGB 近似）
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.62 ? '#1F2937' : '#FFFFFF';
}
