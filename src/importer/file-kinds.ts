/**
 * 「我们能打开哪些文件」——**唯一口径**（工具栏按钮、拖放、导入校验都读它）。
 *
 * 用户提问："为什么打开 xlsx 按钮名称为 xlsx，对其他格式的兼容性呢？"
 * 用户要求（后续）："我希望能支持常见表格格式的兼容……请把其他表格格式都加上去，
 * 并确保按钮名称为打开表格，鼠标移上去的时候要变手形。"
 *
 * 现在能打开的（都是**真解析**，不是"能选但打不开"）：
 *  - `.xlsx` 普通工作簿、`.xlsm` 启用宏的工作簿（宏原样保留、**不执行**）、`.xltx`/`.xltm` 模板
 *    —— **OOXML 包**，按 OPC 关系找部件，解析器根本不看扩展名；
 *  - `.csv` / `.tsv` / `.txt` —— 分隔文本：自研 RFC4180 解析 + 编码识别（UTF-8/UTF-16/GBK），
 *    再合成为一份规范 xlsx 走既有全链路（见 `src/parser/csv.ts`）；
 *  - `.ods` —— OpenDocument 表格（zip + XML，自研解析，见 `src/parser/ods.ts`）；
 *  - `.xls` —— Excel 97–2003 的 **BIFF8 二进制**（自研 CFB + 记录流解析，见 `src/parser/xls.ts`）。
 *
 * 后三种格式本身**无法被"外科式回写"**（我们的导出是在 OOXML 字节上打补丁），
 * 所以导入时会被翻译成一份规范 xlsx：预览、编辑、工作区、撤销、会话恢复全部照旧，
 * **导出会另存为 .xlsx**（原名 + `-已编辑.xlsx`），这条会在导入摘要里明确写出来。
 *
 * 仍然打不开的（明确拒绝 + 指路，不做"半支持"）：
 *  - `.xlsb`：Excel 的**二进制** OOXML（zip 里是 binary parts，不是 XML）。**本地版**可以借本机 Excel
 *    先转成 xlsx 再打开（见 `src/shell/app-form.ts` 与 `build/vite-plugin-excel-bridge.ts`）；
 *    静态托管版没有本机进程，仍按"打不开"处理并说明原因；
 *  - `.numbers` / `.et` / `.ett`：iWork / WPS 私有格式；
 *  - 其它（PDF/图片/纯文本日志…）根本不是表格。
 */

/** 扩展名 → 处理方式 */
export type WorkbookFileKind = 'ooxml' | 'csv' | 'ods' | 'xls' | 'xlsb' | 'foreign' | 'other';

/** OOXML 工作簿包（同一套解析路径；写进 `<input accept>`） */
export const OOXML_WORKBOOK_EXTENSIONS = ['.xlsx', '.xlsm', '.xltx', '.xltm'] as const;
/** 分隔文本 */
export const DELIMITED_EXTENSIONS = ['.csv', '.tsv', '.txt'] as const;
/** 需要"翻译成 xlsx"才能导入的格式（导出会另存为 .xlsx） */
export const CONVERTED_EXTENSIONS = ['.csv', '.tsv', '.txt', '.ods', '.xls'] as const;
/** **只有本地版能打开**的格式（要借本机 Excel 先转成 xlsx） */
export const BRIDGE_ONLY_EXTENSIONS = ['.xlsb'] as const;

/** 能打开的**全部**扩展名（顺序即 UI 提示里的顺序；`accept` 直接由它生成，避免两处漂移） */
export const SUPPORTED_EXTENSIONS = [...OOXML_WORKBOOK_EXTENSIONS, ...DELIMITED_EXTENSIONS, '.ods', '.xls'] as const;

/** 直接喂给 `<input type="file" accept>` 的字符串 */
export const ACCEPT_ATTR = SUPPORTED_EXTENSIONS.join(',');

/**
 * 本地版的 `accept`：多出 `.xlsb`。
 * 形态不同、可打开的范围就不同，所以 accept 也必须跟着形态走（否则文件选择框里选不到 .xlsb）。
 */
export const ACCEPT_ATTR_WITH_BRIDGE = [...SUPPORTED_EXTENSIONS, ...BRIDGE_ONLY_EXTENSIONS].join(',');

/** 给用户看的一句话（按钮 tooltip、拖放提示共用），避免各处各写一份 */
export const FORMAT_HINT =
  '支持 .xlsx / .xlsm / .xltx / .xltm / .ods / .xls / .csv / .tsv（xls·ods·csv 会转换为 xlsx，宏不执行、原样保留）';

/** 本地版额外的提示 */
export const FORMAT_HINT_WITH_BRIDGE = `${FORMAT_HINT}；本地版还能借本机 Excel 打开 .xlsb`;

/** 取小写扩展名（没有扩展名时返回空串） */
export function extensionOf(fileName: string): string {
  const match = /\.[^./\\]+$/.exec(fileName ?? '');
  return match ? match[0].toLowerCase() : '';
}

export function fileKindOf(fileName: string): WorkbookFileKind {
  const ext = extensionOf(fileName);
  if ((OOXML_WORKBOOK_EXTENSIONS as readonly string[]).includes(ext)) return 'ooxml';
  if ((DELIMITED_EXTENSIONS as readonly string[]).includes(ext)) return 'csv';
  if (ext === '.ods' || ext === '.fods') return 'ods';
  if (ext === '.xls' || ext === '.xlt' || ext === '.xla' || ext === '.xlw') return 'xls';
  if (ext === '.xlsb') return 'xlsb';
  if (ext === '.numbers' || ext === '.et' || ext === '.ett') return 'foreign';
  return 'other';
}

/** 这个文件我们能不能打开（能不能打开 ≠ 能不能原样导出，后者见 `exportKindOf`） */
export function isSupportedWorkbookFile(fileName: string): boolean {
  const kind = fileKindOf(fileName);
  return kind === 'ooxml' || kind === 'csv' || kind === 'ods' || kind === 'xls';
}

/** 这个格式是不是"只有本地版（带本机 Excel 转换桥）才打得开" */
export function isBridgeOnlyFile(fileName: string): boolean {
  return fileKindOf(fileName) === 'xlsb';
}

/** 导入时是不是"需要先翻译成 xlsx"（导入摘要里要说明，导出也会变成 .xlsx） */
export function needsConversion(fileName: string): boolean {
  const kind = fileKindOf(fileName);
  return kind === 'csv' || kind === 'ods' || kind === 'xls';
}

/** 格式的中文名（导入摘要、日志里说人话） */
export function kindLabel(kind: WorkbookFileKind): string {
  switch (kind) {
    case 'ooxml':
      return 'Excel 工作簿';
    case 'csv':
      return 'CSV/文本表格';
    case 'ods':
      return 'OpenDocument 表格(.ods)';
    case 'xls':
      return 'Excel 97-2003(.xls)';
    case 'xlsb':
      return 'Excel 二进制工作簿(.xlsb)';
    case 'foreign':
      return '其它表格格式';
    default:
      return '未知格式';
  }
}

/**
 * 打不开时给用户的**原因 + 下一步**（不是一句"不支持"打发）。
 *
 * 每类格式一句话，末尾统一附上"请先另存为 .xlsx"这条可照做的动作；
 * 多个文件时把文件名列出来（用户一次可能拖进来好几个）。
 */
export function unsupportedFileMessage(fileNames: string[]): string {
  const names = fileNames.filter(Boolean);
  const list = names.join('、');
  const kind = names.length > 0 ? fileKindOf(names[0]) : 'other';
  switch (kind) {
    case 'xlsb':
      return (
        `.xlsb 是 Excel 的二进制工作簿：本工具的解析器读不了它。` +
        `本地版可以借本机 Excel 先转成 .xlsx 再打开（装好 Excel 后重开本地版即可）；` +
        `静态托管版没有本机进程，请先「另存为 .xlsx」再拖进来：${list}`
      );
    case 'foreign':
      return `这不是 Excel/OpenDocument 表格（Numbers、WPS 私有格式等需要各自的解析器）。请先另存为 .xlsx：${list}`;
    default:
      return `只支持表格文件（${SUPPORTED_EXTENSIONS.join(' / ')}）：${list}`;
  }
}

/**
 * 导出文件名。
 *
 *  - **OOXML 源**：保留原扩展名 —— 导出是"在原字节上做外科式修补"，`.xlsm` 出来的字节里仍带着
 *    `vbaProject.bin` 与 macroEnabled 的内容类型，硬改名叫 `.xlsx` 会让 Excel 认为"格式与扩展名不符"；
 *  - **转换源**（.csv/.ods/.xls）：标签页里的字节已经是我们合成的 xlsx，所以导出叫 `-已编辑.xlsx`。
 */
export function exportFileNameFor(originalName: string, kind: WorkbookFileKind = fileKindOf(originalName)): string {
  const ext = extensionOf(originalName);
  const stem = originalName.slice(0, originalName.length - ext.length) || 'workbook';
  if (kind === 'ooxml') return `${stem}-已编辑${ext || '.xlsx'}`;
  return `${stem}-已编辑.xlsx`;
}
