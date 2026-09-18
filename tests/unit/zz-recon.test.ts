/** 临时侦察脚本（稍后删除）：xls 解析结果 vs xlsx 真值 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, it } from 'vitest';
import { parseXlsx } from '../../src/parser';
import { parseXlsDetailed } from '../../src/parser/xls';

const FIX = join(process.cwd(), 'fixtures');
const FILES = ['fixture-styles', 'fixture-numfmt', 'fixture-rules', 'fixture-extras', 'fixture-multi'];

const colLetter = (c: number): string => {
  let n = c;
  let s = '';
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
};

describe('侦察', () => {
  it('对比 xls 与 xlsx', async () => {
    for (const name of FILES) {
      const xls = parseXlsDetailed(new Uint8Array(readFileSync(join(FIX, `${name}.xls`))));
      const truth = await parseXlsx(new Uint8Array(readFileSync(join(FIX, `${name}.xlsx`))));
      console.log(`\n########## ${name} ##########`);
      console.log('xls stats', JSON.stringify(xls.stats));
      console.log('xls warnings', JSON.stringify(xls.warnings, null, 1));
      console.log('xls styles', JSON.stringify(xls.input.styles));
      console.log(
        'xls sheets',
        xls.input.sheets
          .map((s) => `${s.name}(hidden=${!!s.hidden}, freeze=${JSON.stringify(s.freeze)}, merges=${JSON.stringify(s.merges)}, cols=${JSON.stringify(s.colWidths)}, rows=${JSON.stringify(s.rowHeights)}, defRowH=${s.defaultRowHeight}, defColW=${s.defaultColWidth}, grid=${s.gridlinesHidden})`)
          .join(' | '),
      );
      for (const sheet of xls.input.sheets) {
        for (const cell of sheet.cells) {
          const st = cell.style !== undefined ? xls.input.styles?.[cell.style] : undefined;
          console.log(`  ${sheet.name} ${colLetter(cell.col)}${cell.row + 1} = ${JSON.stringify(cell.value)} ${st ? JSON.stringify(st) : ''}`);
        }
      }
      console.log('--- xlsx 真值 ---');
      for (const sheet of truth.sheets) {
        console.log(
          `  ${sheet.name}(hidden=${!!sheet.hidden}, veryHidden=${!!sheet.veryHidden}, freeze=${JSON.stringify(sheet.freeze)}, merges=${JSON.stringify(sheet.merges)}, cols=${JSON.stringify(sheet.cols)}, rows=${JSON.stringify(sheet.rows)}, defRowH=${sheet.defaultRowHeight})`,
        );
        for (const cell of sheet.cells) {
          const st = cell.styleIndex !== undefined ? truth.styles[cell.styleIndex] : undefined;
          console.log(`    ${colLetter(cell.col)}${cell.row + 1} = ${JSON.stringify(cell.value)} f=${JSON.stringify(cell.formula)} ${st ? JSON.stringify(st) : ''}`);
        }
      }
    }
  });
});
