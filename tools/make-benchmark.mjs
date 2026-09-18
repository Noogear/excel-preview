/**
 * 大文件基准夹具生成器（⑩ 极致性能）。
 *
 * 目标：造一个"真实感"的大 xlsx——不是纯数字灌水，而是混合
 *   ① 共享字符串（走 sharedStrings，考解析器的字符串表）
 *   ② 数值 + 公式（考公式与数字格式）
 *   ③ 多种样式（考 styles.xml 解析与单元格样式指纹）
 *   ④ 若干空单元格（考稀疏性处理）
 * 这样测出来的数字才对得上真实表格。
 *
 * 规模可用环境变量覆盖：
 *   BENCH_ROWS=100000 BENCH_COLS=20 node tools/make-benchmark.mjs
 *
 * 用 exceljs 的流式写入器（WorkbookWriter）：100k 行 × 20 列 = 200 万格，
 * 一次性 writeFile 会把整份工作表堆在内存里，流式写入才稳。
 */
import ExcelJS from 'exceljs';
import { mkdirSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = join(ROOT, 'fixtures');
const ROWS = Number(process.env.BENCH_ROWS ?? 50_000);
const COLS = Number(process.env.BENCH_COLS ?? 20);
const OUT_NAME = process.env.BENCH_OUT ?? 'bench-large.xlsx';

const HEADERS = [
  '学号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '生物', '政治',
  '历史', '地理', '总分', '平均分', '排名', '等级', '备注', '联系电话', '宿舍', '状态',
];

async function main() {
  mkdirSync(OUT_DIR, { recursive: true });
  const target = join(OUT_DIR, OUT_NAME);

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: target,
    useSharedStrings: true,
    useStyles: true,
  });
  workbook.creator = 'benchmark generator';

  const sheet = workbook.addWorksheet('成绩总表', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  sheet.columns = HEADERS.slice(0, COLS).map((header, index) => ({
    header,
    key: `c${index}`,
    width: index < 3 ? 14 : 10,
  }));

  // 标题行加个样式，确保 styles.xml 不是空的
  sheet.getRow(1).font = { bold: true, size: 11, name: '微软雅黑' };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
  sheet.getRow(1).commit();

  const start = Date.now();
  for (let r = 0; r < ROWS; r += 1) {
    const row = sheet.getRow(r + 2);
    const studentId = 20240000 + r;
    row.getCell(1).value = `S${studentId}`;
    row.getCell(2).value = r % 7 === 0 ? '' : `学生${(r % 9973) + 1}`; // 混入空单元格
    row.getCell(3).value = `${(r % 12) + 1}班`;
    for (let c = 4; c <= Math.min(12, COLS); c += 1) {
      row.getCell(c).value = 60 + ((r * 7 + c * 13) % 41); // 60..100
    }
    if (COLS >= 13) row.getCell(13).value = { formula: `SUM(D${r + 2}:L${r + 2})`, result: 0 };
    if (COLS >= 14) row.getCell(14).value = { formula: `ROUND(M${r + 2}/9,2)`, result: 0 };
    if (COLS >= 15) row.getCell(15).value = (r % ROWS) + 1;
    if (COLS >= 16) row.getCell(16).value = ['A', 'B', 'C', 'D'][r % 4];
    if (COLS >= 17) row.getCell(17).value = r % 5 === 0 ? '' : '表现稳定';
    if (COLS >= 18) row.getCell(18).value = `138${String(10000000 + (r % 9000000)).slice(0, 8)}`;
    if (COLS >= 19) row.getCell(19).value = `${(r % 30) + 101}室`;
    if (COLS >= 20) row.getCell(20).value = r % 3 === 0 ? '在读' : '休学';
    // 每 100 行给整行来个底色，考"样式指纹"与解析器的样式表规模
    if (r % 100 === 0) {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      });
    }
    row.commit();
    if (r > 0 && r % 20_000 === 0) {
      console.log(`  … 已写 ${r} 行（${((Date.now() - start) / 1000).toFixed(1)}s）`);
    }
  }
  await workbook.commit();

  const size = statSync(target).size;
  const cells = ROWS * COLS;
  console.log(`✓ ${OUT_NAME}`);
  console.log(`  规模：${ROWS} 行 × ${COLS} 列 = ${cells.toLocaleString('en-US')} 格`);
  console.log(`  体积：${(size / 1024 / 1024).toFixed(1)} MB`);
  console.log(`  生成耗时：${((Date.now() - start) / 1000).toFixed(1)}s`);
}

main().catch((error) => {
  console.error('生成失败：', error);
  process.exit(1);
});
