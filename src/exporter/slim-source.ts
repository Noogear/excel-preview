/**
 * 导入完成后的**模型瘦身**：把常驻内存的 `ParsedWorkbook` 压缩成导出真正需要的最小形状。
 *
 * 背景（实测）：100 万格的表，完整解析模型里光是 `ParsedCell` 对象就有上百万个，
 * 浏览器堆水位到 543 MB。但外科式导出只用到：
 *   ① 原始 zip 条目（这里改用**原始字节 + 导出时惰性解压**，省掉几十 MB 常驻）
 *   ② `sheet.id`（把编辑定位到对应工作表部件）
 *   ③ "每行第一个带样式的 s"（给新写入的格子补样式，见 exporter 的 collectRowStyles）
 *
 * 所以瘦身后只需要保留原始字节（本来就要留着——会话恢复与导出都靠它）与一份
 * `行号 → s` 的小映射（几万条，几 MB）。`cells` / `rows` / 样式表 / 条件格式等一律丢掉：
 * 它们只服务于"导入那一刻的渲染与特性应用"，那一步早就做完了，
 * 而导出对这些部件是**字节级原样保留**，根本不需要在内存里重建它们。
 *
 * 这样做的前提是：**没有任何运行期路径再读完整模型**（已核对：只有导出用它，
 * 标签切换用的是 Univer 内存里的工作簿，会话只存原始字节与编辑）。
 */
import type { ParsedWorkbook } from '../parser/types';
import type { ExportSource, ExportSourceSheet } from './export-xlsx';

/**
 * 把完整解析结果压成导出所需的最小数据源（不修改原对象）。
 *
 * `bytes` 是原始 xlsx 字节；给了它导出就会惰性解压（推荐——常驻内存只留这几 MB，
 * 代价是每次导出多一次解压，6 MB 文件实测约 0.2 s）。不传则导出仍然可用，
 * 但那时数据源里既没有 `raw.entries` 也没有 `bytes`，`exportXlsx` 会明确报错而不是静默出错。
 */
export function slimForExport(parsed: ParsedWorkbook, bytes?: Uint8Array): ExportSource {
  const sheets: ExportSourceSheet[] = parsed.sheets.map((sheet) => {
    const rowStyles = new Map<number, number>();
    for (const cell of sheet.cells) {
      if (cell.styleIndex === undefined) continue;
      if (!rowStyles.has(cell.row)) rowStyles.set(cell.row, cell.styleIndex);
    }
    return { id: sheet.id, rowStyles };
  });

  return bytes ? { bytes, sheets } : { sheets };
}
