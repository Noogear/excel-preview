/**
 * 「导入到工作区」对话框的输入解析（**纯函数**，无副作用、零依赖、可单测）。
 *
 * 边界与约定（刻意对齐 Excel 的真实限制，避免生成出表格装不下的 A1 记号）：
 *  - 行号：1 ~ 1048576（Excel 最大行），只接受十进制正整数；可写单行（`3`）或行区间（`3:5`）；
 *  - 列标：A ~ XFD（Excel 最大列，第 16384 列），1~3 个字母；可写单列（`B`）或列区间（`B:D`）；
 *  - 区域：`A1` 或 `A1:C10`，两端坐标**按轴各自取 min/max 归一化**，
 *    所以 `C10:A1` / `A10:C1` 都会规范化成 `A1:C10`（反转的区间不是错误，而是被摆正）；
 *    行 / 列区间同理（`5:3`→`3:5`、`D:B`→`B:D`）；
 *  - `$` 绝对引用记号会被忽略：`$A$1:$B$18`→`A1:B18`、`$B:$D`→`B:D`（用户从 Excel 名称框 / 公式里
 *    复制出来就是这个样子）；但 `$` 只允许紧贴在列标前与行号前，`A1$` / `A$$1` / `$$A1` 仍然非法；
 *  - 大小写不敏感；只去掉**首尾**空格（`' B:D '` 合法，`'B : D'` 非法：那是格式错误）；
 *  - 全角字符（`Ａ１`、`３`、`：`）一律按非法处理：Univer / A1 记号只认 ASCII。
 *
 * 错误信息是面向用户的中文短句，UI 会把它们原样显示在输入框下方。
 *
 * 三个解析函数（行 / 列 / 区域）的**字符集互不相交**（行=数字+冒号、列=字母与`$`+冒号、
 * 区域=字母+数字的单元格记号），所以「自动识别」（见 `detectImportKind`）的结果永远唯一，
 * 且"识别得出来" ⟺ "对应解析一定成功"。
 */

/** 解析结果：判别联合（`ok` 收窄），失败时 `error` 一定非空 */
export type ParseResult = { ok: true; value: string } | { ok: false; error: string };

/** 「导入到工作区」的目标类型：整行 / 整列 / 矩形区域（也是 `onSubmit` 的 `kind`） */
export type ImportTargetKind = 'row' | 'column' | 'range';

/** Excel 单表最大行号（1 基） */
export const MAX_ROW = 1_048_576;
/** Excel 单表最大列标（1 基，即第 16384 列） */
export const MAX_COLUMN_LETTERS = 'XFD';

const MAX_COLUMN = 16_384;

/** 行号类错误文案（`'0'` / `'abc'` / `''` 都归到这里：空串同样"不是大于 0 的整数"） */
const ROW_FORMAT_ERROR = '行号必须是大于 0 的整数';
const ROW_RANGE_ERROR = `行号超出范围（最大 ${MAX_ROW}）`;
const COLUMN_FORMAT_ERROR = '列标应为字母（例如 B 或 B:D）';
const COLUMN_RANGE_ERROR = `列标超出范围（最大 ${MAX_COLUMN_LETTERS}）`;
const RANGE_FORMAT_ERROR = '区域格式应为 A1:C10';

function ok(value: string): ParseResult {
  return { ok: true, value };
}

function fail(error: string): ParseResult {
  return { ok: false, error };
}

/** 去掉首尾空格（非字符串按空串处理，防止 JS 调用方传进 undefined） */
function trim(raw: string): string {
  return typeof raw === 'string' ? raw.trim() : '';
}

/** 1 基列标 → 列号；`A`→1、`Z`→26、`AA`→27、`XFD`→16384 */
function lettersToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** 列号 → 1 基列标（与 snapshot.ts 的 colToLetter 同算法，但入参是 1 基） */
function indexToLetters(index: number): string {
  let n = index;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

type ColumnResult = { ok: true; letters: string; index: number } | { ok: false; error: string };

/**
 * 校验单个列标：非字母 → 格式错误；字母但超长 / 超出 XFD → 范围错误。
 * 允许一个**紧贴列标前**的 `$`（Excel 绝对引用，`$B`→`B`）；`$$B` / `B$` 仍然非法。
 */
function parseColumnToken(token: string): ColumnResult {
  const match = /^\$?([A-Za-z]+)$/.exec(token);
  if (!match) return { ok: false, error: COLUMN_FORMAT_ERROR };

  const letters = match[1].toUpperCase();
  if (letters.length > MAX_COLUMN_LETTERS.length) return { ok: false, error: COLUMN_RANGE_ERROR };
  const index = lettersToIndex(letters);
  if (index > MAX_COLUMN) return { ok: false, error: COLUMN_RANGE_ERROR };
  return { ok: true, letters, index };
}

type RowResult = { ok: true; row: number } | { ok: false; error: string };

/** 校验单个行号：非纯数字 → 格式错误；0 / 越界 → 相应的错误文案 */
function parseRowToken(token: string): RowResult {
  if (!/^[0-9]+$/.test(token)) return { ok: false, error: ROW_FORMAT_ERROR };

  const row = Number(token);
  if (!Number.isSafeInteger(row) || row < 1) return { ok: false, error: ROW_FORMAT_ERROR };
  if (row > MAX_ROW) return { ok: false, error: ROW_RANGE_ERROR };

  return { ok: true, row };
}

type CellResult = { ok: true; colIndex: number; row: number } | { ok: false; error: string };

/**
 * 校验单个单元格记号（`A1`）：列、行分别给出具体错误。
 * 列标前 / 行号前各允许一个 `$`（`$A$1`、`A$1`、`$A1` 都等价于 `A1`），
 * 但 `$` 不能出现在别处（`A1$`、`A$$1` 是格式错误）。
 */
function parseCellToken(token: string): CellResult {
  const match = /^\$?([A-Za-z]+)\$?([0-9]+)$/.exec(token);
  if (!match) return { ok: false, error: RANGE_FORMAT_ERROR };

  const column = parseColumnToken(match[1]);
  if (!column.ok) return { ok: false, error: column.error };

  const row = parseRowToken(match[2]);
  if (!row.ok) return { ok: false, error: row.error };

  return { ok: true, colIndex: column.index, row: row.row };
}

/**
 * 单行 / 行区间输入 → 规范化行号（1 基，去前导零；区间两端取 min/max 摆正）。
 * `'3'`→`'3'`；`' 3 '`→`'3'`；`'007'`→`'7'`；`'5:3'`→`'3:5'`；
 * `'0'` / `'abc'` / `''` / `'3.5'` / `'-1'` / `'３'` / `'3 : 5'` / `'3:5:7'` → 错误。
 */
export function parseRowInput(raw: string): ParseResult {
  const text = trim(raw);
  if (text === '') return fail(ROW_FORMAT_ERROR);

  const parts = text.split(':');
  if (parts.length > 2) return fail(ROW_FORMAT_ERROR);

  const first = parseRowToken(parts[0]);
  if (!first.ok) return fail(first.error);
  if (parts.length === 1) return ok(String(first.row));

  const second = parseRowToken(parts[1]);
  if (!second.ok) return fail(second.error);

  const start = Math.min(first.row, second.row);
  const end = Math.max(first.row, second.row);
  return ok(`${start}:${end}`);
}

/**
 * 单列输入 → 规范化列标，支持单列与列区间。
 * `'b'`→`'B'`；`'b:d'`→`'B:D'`；`'D:B'`→`'B:D'`（反转摆正）；`'$B:$D'`→`'B:D'`（绝对引用记号忽略）；`'XFE'`→错误。
 */
export function parseColumnInput(raw: string): ParseResult {
  const text = trim(raw);
  if (text === '') return fail(COLUMN_FORMAT_ERROR);

  const parts = text.split(':');
  if (parts.length > 2) return fail(COLUMN_FORMAT_ERROR);

  const first = parseColumnToken(parts[0]);
  if (!first.ok) return fail(first.error);
  if (parts.length === 1) return ok(indexToLetters(first.index));

  const second = parseColumnToken(parts[1]);
  if (!second.ok) return fail(second.error);

  const start = Math.min(first.index, second.index);
  const end = Math.max(first.index, second.index);
  return ok(`${indexToLetters(start)}:${indexToLetters(end)}`);
}

/**
 * 区域输入 → 规范化的 A1 记号（两端坐标按轴取 min/max）。
 * `'a1:c10'`→`'A1:C10'`；`'b2'`→`'B2'`；`'C10:A1'`→`'A1:C10'`；`'$A$1:$B$18'`→`'A1:B18'`（绝对引用记号忽略）；
 * `'A0'` / `''` / `'Ａ１'` / `'A1$'` → 错误。
 */
export function parseRangeInput(raw: string): ParseResult {
  const text = trim(raw);
  if (text === '') return fail(RANGE_FORMAT_ERROR);

  const parts = text.split(':');
  if (parts.length > 2) return fail(RANGE_FORMAT_ERROR);

  const a = parseCellToken(parts[0]);
  if (!a.ok) return fail(a.error);
  if (parts.length === 1) return ok(`${indexToLetters(a.colIndex)}${a.row}`);

  const b = parseCellToken(parts[1]);
  if (!b.ok) return fail(b.error);

  const startRow = Math.min(a.row, b.row);
  const endRow = Math.max(a.row, b.row);
  const startCol = Math.min(a.colIndex, b.colIndex);
  const endCol = Math.max(a.colIndex, b.colIndex);
  return ok(`${indexToLetters(startCol)}${startRow}:${indexToLetters(endCol)}${endRow}`);
}

/**
 * **自动识别**输入的是行、列还是区域（对话框唯一输入框用），识别不出返回 `null`。
 *
 * 规则（大小写不敏感、允许首尾空格；**内部空格 / 全角字符一律非法**，与三个解析函数同严格度）：
 *  - 行：`3`、`3:5`（纯数字，可带一个冒号）；
 *  - 列：`B`、`B:D`、`$B:$D`（纯字母，可带一个冒号）；
 *  - 区域：`A1`、`A1:B18`、`$A$1:$B$18`（字母 + 数字的单元格记号，可带一个冒号）；
 *  - 其余（`''`、`'B : D'`、`'３'`、`'Ａ１：Ｃ１０'`、`'0'`、`'XFE'`、`'A1C10'`、`'3:5:7'` …）→ `null`。
 *
 * 实现就是"按顺序试三个解析函数"：三者的字符集互不相交，所以结果唯一，顺序不影响判定；
 * 这样也保证了「识别得出类型」⟺「该类型一定能解析成功」，UI 的识别行与校验永远不会打架。
 */
export function detectImportKind(raw: string): ImportTargetKind | null {
  const text = trim(raw);
  if (text === '') return null;
  if (parseRowInput(text).ok) return 'row';
  if (parseColumnInput(text).ok) return 'column';
  if (parseRangeInput(text).ok) return 'range';
  return null;
}

/**
 * 行号 → 整行 A1 记号（供上层把"单行"拼成区域）：`1` → `'1:1'`。
 * 非法行号（非整数 / < 1 / > 最大行）返回空串，调用方用 `!== ''` 判断即可。
 */
export function a1OfRow(row: number): string {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) return '';
  return `${row}:${row}`;
}

/**
 * 列标 → 整列 A1 记号：`'B'` → `'B:B'`。
 * 也接受已经是区间的输入（`'b:d'` → `'B:D'`，`'D:B'` → `'B:D'`），方便上层直接透传
 * `parseColumnInput` 的结果；非法输入返回空串。
 */
export function a1OfColumn(letters: string): string {
  const parsed = parseColumnInput(letters);
  if (!parsed.ok) return '';
  return parsed.value.includes(':') ? parsed.value : `${parsed.value}:${parsed.value}`;
}
