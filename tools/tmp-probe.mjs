import { parseXlsx } from './src/parser/index.ts';
import { writeWorkbookPackage } from './src/importer/synth-xlsx.ts';
import { unzipSync, strFromU8 } from 'fflate';

const bytes = writeWorkbookPackage({
  sheets: [{ name: 's', cells: [{ row: 0, col: 0, value: 'X', style: 0 }] }],
  styles: [{ fontFamily: '微软雅黑', fontSize: 14, bold: true, italic: true, fill: '#FFF2CC', color: '#CC0000', numberFormat: '0.00"元"', horizontalAlign: 'center', textWrap: true, textRotation: 45, verticalAlign: 'middle', underline: true, strikeThrough: true, border: { top: { style: 'thin', color: '#4F81BD' }, bottom: { style: 'double' } } }],
});
const zip = unzipSync(bytes);
console.log(strFromU8(zip['xl/styles.xml']).slice(0, 2000));
const parsed = await parseXlsx(bytes);
console.log('styles:', JSON.stringify(parsed.styles, null, 1));
console.log('cell:', JSON.stringify(parsed.sheets[0].cells));
