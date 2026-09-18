/**
 * 打印样本里"P1 待还原"部件的真实 XML 结构，用于对齐解析器类型定义。
 * 运行：node tools/inspect-p1.mjs [文件名...]
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';

const files = process.argv.slice(2);
const targets = files.length > 0 ? files : ['fixture-rules.xlsx', 'fixture-extras.xlsx', 'fixture-table.xlsx'];

function slice(xml, marker, len = 700) {
  const idx = xml.indexOf(marker);
  if (idx < 0) return `（未找到 ${marker}）`;
  return xml.slice(idx, idx + len);
}

for (const file of targets) {
  const buffer = readFileSync(join(process.cwd(), 'fixtures', file));
  const entries = unzipSync(new Uint8Array(buffer));
  console.log(`\n${'='.repeat(78)}\n# ${file}\n${'='.repeat(78)}`);
  console.log('条目：', Object.keys(entries).filter((n) => !n.endsWith('/')).join('  '));

  const sheet1 = entries['xl/worksheets/sheet1.xml'] ? strFromU8(entries['xl/worksheets/sheet1.xml']) : '';
  const sheet2 = entries['xl/worksheets/sheet2.xml'] ? strFromU8(entries['xl/worksheets/sheet2.xml']) : '';

  console.log('\n--- conditionalFormatting ---');
  console.log(slice(sheet1, '<conditionalFormatting', 900));
  console.log('\n--- dataValidations ---');
  console.log(slice(sheet2 || sheet1, '<dataValidations', 700));
  console.log('\n--- hyperlinks ---');
  console.log(slice(sheet1, '<hyperlinks', 500));
  console.log('\n--- tableParts ---');
  console.log(slice(sheet1, '<tableParts', 300));
  console.log('\n--- legacyDrawing / drawing ---');
  console.log(slice(sheet1, '<legacyDrawing', 200));
  console.log(slice(sheet1, '<drawing ', 200));

  const styles = entries['xl/styles.xml'] ? strFromU8(entries['xl/styles.xml']) : '';
  if (styles.includes('<dxfs')) {
    console.log('\n--- styles.xml dxfs ---');
    console.log(slice(styles, '<dxfs', 800));
  }

  const tableEntry = Object.keys(entries).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
  if (tableEntry) {
    console.log(`\n--- ${tableEntry} ---`);
    console.log(strFromU8(entries[tableEntry]).slice(0, 900));
  }

  const commentEntry = Object.keys(entries).find((n) => /^xl\/comments\d+\.xml$/.test(n));
  if (commentEntry) {
    console.log(`\n--- ${commentEntry} ---`);
    console.log(strFromU8(entries[commentEntry]).slice(0, 700));
  }

  const drawingEntry = Object.keys(entries).find((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
  if (drawingEntry) {
    console.log(`\n--- ${drawingEntry} ---`);
    console.log(strFromU8(entries[drawingEntry]).slice(0, 1400));
    const rels = Object.keys(entries).find((n) => /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(n));
    if (rels) {
      console.log(`\n--- ${rels} ---`);
      console.log(strFromU8(entries[rels]).slice(0, 500));
    }
  }

  const sheetRels = Object.keys(entries).find((n) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n));
  if (sheetRels) {
    console.log(`\n--- ${sheetRels} ---`);
    console.log(strFromU8(entries[sheetRels]).slice(0, 900));
  }

  const media = Object.keys(entries).filter((n) => n.startsWith('xl/media/'));
  if (media.length) console.log('\n媒体文件：', media.join('  '));
}
