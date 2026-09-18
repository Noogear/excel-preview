/**
 * 调试脚本：打印样本 xlsx 内部 XML 的"真值"统计，用于核对解析器覆盖度。
 * 运行：node tools/inspect-fixture.mjs [文件名]
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

const fileName = process.argv[2] ?? 'fixture-styles.xlsx';
const buffer = readFileSync(join(process.cwd(), 'fixtures', fileName));
const entries = unzipSync(new Uint8Array(buffer));

console.log(`# ${fileName}`);
console.log('zip 条目：', Object.keys(entries).join('\n  '));

const sheetEntry = Object.keys(entries).find((name) => /^xl\/worksheets\/sheet1\.xml$/.test(name));
if (!sheetEntry) process.exit(0);

const xml = strFromU8(entries[sheetEntry]);
console.log(`\n--- ${sheetEntry} (${xml.length} 字符) ---`);
console.log('真值统计：');
console.log('  <row 出现次数      :', (xml.match(/<row[\s>]/g) ?? []).length);
console.log('  <c 出现次数        :', (xml.match(/<c[\s>]/g) ?? []).length);
console.log('  <f> 公式数         :', (xml.match(/<f[\s>]/g) ?? []).length);
console.log('  <mergeCell 数      :', (xml.match(/<mergeCell[\s>]/g) ?? []).length);
console.log('  <pane 出现         :', (xml.match(/<pane[\s>]/g) ?? []).length);
console.log('  自闭合 <c .../> 数 :', (xml.match(/<c[^>]*\/>/g) ?? []).length);

const head = xml.slice(0, 1200);
console.log('\n--- 前 1200 字符 ---');
console.log(head);

const idx = xml.indexOf('<mergeCells');
if (idx >= 0) console.log('\n--- mergeCells 附近 ---\n' + xml.slice(idx, idx + 300));

const paneIdx = xml.indexOf('<pane');
if (paneIdx >= 0) console.log('\n--- pane 附近 ---\n' + xml.slice(paneIdx, paneIdx + 300));
else console.log('\n（无 <pane>：说明 ExcelJS 未写入冻结窗格）');

const stylesXml = entries['xl/styles.xml'] ? strFromU8(entries['xl/styles.xml']) : '';
if (stylesXml) {
  console.log('\n--- styles.xml 统计 ---');
  console.log('  <fill 数    :', (stylesXml.match(/<fill[\s>]/g) ?? []).length);
  console.log('  <border 数  :', (stylesXml.match(/<border[\s>]/g) ?? []).length);
  console.log('  <xf 数      :', (stylesXml.match(/<xf[\s>]/g) ?? []).length);
  console.log('  <numFmt 数  :', (stylesXml.match(/<numFmt[\s>]/g) ?? []).length);
  const fillsIdx = stylesXml.indexOf('<fills');
  console.log('\n--- fills 片段 ---\n' + stylesXml.slice(fillsIdx, fillsIdx + 600));
  const bordersIdx = stylesXml.indexOf('<borders');
  console.log('\n--- borders 片段 ---\n' + stylesXml.slice(bordersIdx, bordersIdx + 700));
}
