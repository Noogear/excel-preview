/**
 * 「导入到工作区」对话框输入解析单测（**纯函数**，node 环境，不需要 jsdom）。
 *
 * 覆盖 import-parse.ts 的 6 个导出函数：parseRowInput / parseColumnInput / parseRangeInput /
 * detectImportKind / a1OfRow / a1OfColumn；外加 ImportDialog.tsx 里两个纯函数
 * （resolveImportDraft / describeImportTarget：识别 + 归一化 + 文案，UI 只调它们）。
 *
 * 关键边界（曾经出过问题的地方）：
 *  - 大小写不敏感、只去首尾空格（`' B:D '` 合法，`'B : D'` 非法）
 *  - 反转区间要"摆正"：`'5:3'`→`'3:5'`、`'D:B'`→`'B:D'`、`'C10:A1'`→`'A1:C10'`、`'A10:C1'`→`'A1:C10'`
 *  - `$` 绝对引用记号忽略（`'$A$1:$B$18'`→`'A1:B18'`），但只能紧贴列标前 / 行号前
 *  - Excel 真实边界：行 1~1048576、列 A~XFD（`'XFE'` 超范围）
 *  - 自动识别（`detectImportKind`）：3/3:5→行、B/B:D→列、A1/A1:B18→区域，其余（含内部空格、
 *    全角、空串、非法值）→ null，且"识别得出" ⟺ "对应解析一定成功"
 *  - 异常：`''` / `'0'` / `'A0'` / 全角字符（`'３'`、`'Ａ１'`、`'B：D'`）/ 小数点 / 负号 / 多冒号
 *
 * 末尾另附一组 `ImportDialog` 的 SSR 渲染契约冒烟（node 环境无 jsdom，用 renderToStaticMarkup）。
 */
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import {
  ImportDialog,
  describeImportPreview,
  describeImportTarget,
  importConfirmLabel,
  resolveImportDraft,
} from '../../src/workspace/ImportDialog';
import {
  MAX_COLUMN_LETTERS,
  MAX_ROW,
  a1OfColumn,
  a1OfRow,
  detectImportKind,
  parseColumnInput,
  parseRangeInput,
  parseRowInput,
} from '../../src/workspace/import-parse';
import type { ImportTargetKind, ParseResult } from '../../src/workspace/import-parse';
import { resolveImportRanges } from '../../src/workspace/import-target';
import { PREVIEW_EMPTY_CAP, clipRect } from '../../src/workspace/snapshot';
import type { FWorksheet } from '@univerjs/sheets/facade';

/** 断言"成功且规范化结果等于 value" */
function expectOk(result: ParseResult, value: string): void {
  expect(result).toEqual({ ok: true, value });
}

/** 断言"失败、错误信息非空"，可选再断言具体文案 */
function expectFail(result: ParseResult, message?: string): void {
  expect(result.ok).toBe(false);
  if (result.ok) return; // 类型收窄（上面已断言失败，这里只是给 TS 看的）
  expect(typeof result.error).toBe('string');
  expect(result.error.length).toBeGreaterThan(0);
  if (message !== undefined) expect(result.error).toBe(message);
}

const ROW_FORMAT_ERROR = '行号必须是大于 0 的整数';
const ROW_RANGE_ERROR = `行号超出范围（最大 ${MAX_ROW}）`;
const COLUMN_FORMAT_ERROR = '列标应为字母（例如 B 或 B:D）';
const COLUMN_RANGE_ERROR = `列标超出范围（最大 ${MAX_COLUMN_LETTERS}）`;
const RANGE_FORMAT_ERROR = '区域格式应为 A1:C10';

/* -------------------------------------------------------------------------- */
/* parseRowInput                                                               */
/* -------------------------------------------------------------------------- */

describe('parseRowInput - 单行 / 行区间', () => {
  it('正整数规范化：去首尾空格、去前导零', () => {
    expectOk(parseRowInput('3'), '3');
    expectOk(parseRowInput(' 3 '), '3');
    expectOk(parseRowInput('007'), '7');
  });

  it('行区间与列区间同构：两端都规范化，反转的区间被摆正（不是错误）', () => {
    expectOk(parseRowInput('3:5'), '3:5');
    expectOk(parseRowInput(' 3:5 '), '3:5');
    expectOk(parseRowInput('007:010'), '7:10');
    expectOk(parseRowInput('3:3'), '3:3');
    expectOk(parseRowInput('5:3'), '3:5');
    expectOk(parseRowInput('1:1048576'), `1:${MAX_ROW}`);
    expectOk(parseRowInput(`${MAX_ROW}:${MAX_ROW}`), `${MAX_ROW}:${MAX_ROW}`);
  });

  it(`1 与最大行 ${MAX_ROW} 合法，超出范围报错；非正整数一律拒绝（0 / 负数 / 小数 / 科学计数法 / 字母 / 空串）`, () => {
    expectOk(parseRowInput('1'), '1');
    expectOk(parseRowInput(String(MAX_ROW)), String(MAX_ROW));
    expectFail(parseRowInput(String(MAX_ROW + 1)), ROW_RANGE_ERROR);
    expectFail(parseRowInput(`1:${MAX_ROW + 1}`), ROW_RANGE_ERROR);

    for (const bad of ['0', '-1', '3.5', '3e2', 'abc', '', '   ']) {
      expectFail(parseRowInput(bad), ROW_FORMAT_ERROR);
    }
  });

  it('行区间的两端同样严格：任一端非法 / 多冒号 / 冒号两侧有空格 / 空端点 都拒绝', () => {
    expectFail(parseRowInput('0:5'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3:0'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3:'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput(':5'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3:5:7'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3 : 5'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3: 5'), ROW_FORMAT_ERROR);
    expectFail(parseRowInput('3.5:4'), ROW_FORMAT_ERROR);
  });

  it('只认 ASCII 阿拉伯数字：全角、千分位、溢出成 Infinity 的超长数字串都拒绝', () => {
    expectFail(parseRowInput('３'));
    expectFail(parseRowInput('１２'));
    expectFail(parseRowInput('3,000'));
    expectFail(parseRowInput('999999999999999999999999'));
  });
});

/* -------------------------------------------------------------------------- */
/* parseColumnInput                                                            */
/* -------------------------------------------------------------------------- */

describe('parseColumnInput - 单列', () => {
  it('单列与区间都转大写、去首尾空格，区间形式保留（含单列区间 A:A）；反转的区间被摆正，不是错误', () => {
    expectOk(parseColumnInput('B'), 'B');
    expectOk(parseColumnInput('b'), 'B');
    expectOk(parseColumnInput('xfd'), MAX_COLUMN_LETTERS);
    expectOk(parseColumnInput('b:d'), 'B:D');
    expectOk(parseColumnInput(' B:D '), 'B:D');
    expectOk(parseColumnInput('A:A'), 'A:A');
    expectOk(parseColumnInput('Z:AA'), 'Z:AA');           // 跨进位
    expectOk(parseColumnInput('XFD:XFD'), 'XFD:XFD');     // 最大列区间
    expectOk(parseColumnInput('D:B'), 'B:D');
  });

  it('$ 绝对引用记号被忽略（Excel 名称框里复制出来就是这样），但仍必须紧贴列标', () => {
    expectOk(parseColumnInput('$B'), 'B');
    expectOk(parseColumnInput('$b:$d'), 'B:D');
    expectOk(parseColumnInput('$D:$B'), 'B:D');
    expectOk(parseColumnInput('XFD'), MAX_COLUMN_LETTERS);

    expectFail(parseColumnInput('$$B'), COLUMN_FORMAT_ERROR);
    expectFail(parseColumnInput('B$'), COLUMN_FORMAT_ERROR);
    expectFail(parseColumnInput('$'), COLUMN_FORMAT_ERROR);
    expectFail(parseColumnInput('B:$'), COLUMN_FORMAT_ERROR);
  });

  it('超范围（任意一端）/ 带数字 / 多冒号 / 冒号两侧有空格 / 全角 都拒绝', () => {
    expectFail(parseColumnInput('XFE'), COLUMN_RANGE_ERROR);
    expectFail(parseColumnInput('xfe:zzz'), COLUMN_RANGE_ERROR);
    expectFail(parseColumnInput('AAAA'), COLUMN_RANGE_ERROR);
    expectFail(parseColumnInput('$XFE'), COLUMN_RANGE_ERROR);

    expectFail(parseColumnInput('B2'), COLUMN_FORMAT_ERROR);
    expectFail(parseColumnInput('1'), COLUMN_FORMAT_ERROR);
    expectFail(parseColumnInput('B : D'), COLUMN_FORMAT_ERROR);

    expectFail(parseColumnInput('B:'));
    expectFail(parseColumnInput(':D'));
    expectFail(parseColumnInput('B:D:E'));
    expectFail(parseColumnInput(''));
    expectFail(parseColumnInput(' \t '));
    expectFail(parseColumnInput('Ｂ'));
    expectFail(parseColumnInput('B：D'));
  });
});

/* -------------------------------------------------------------------------- */
/* parseRangeInput                                                             */
/* -------------------------------------------------------------------------- */

describe('parseRangeInput - 区域', () => {
  it('大小写不敏感、去首尾空格；单个单元格不加冒号，显式写了冒号就保留区间形式', () => {
    expectOk(parseRangeInput('a1:c10'), 'A1:C10');
    expectOk(parseRangeInput(' A1:C10 '), 'A1:C10');
    expectOk(parseRangeInput('b2'), 'B2');
    expectOk(parseRangeInput('B2:B2'), 'B2:B2');
    expectOk(parseRangeInput(`a1:xfd${MAX_ROW}`), `A1:XFD${MAX_ROW}`);
  });

  it('$ 绝对引用记号被忽略：列标前 / 行号前各一个（$A$1 / A$1 / $A1 等价），但别处不行', () => {
    expectOk(parseRangeInput('$A$1'), 'A1');
    expectOk(parseRangeInput('$a$1:$b$18'), 'A1:B18');
    expectOk(parseRangeInput('A$1:$B$18'), 'A1:B18');
    expectOk(parseRangeInput('$A1:$B18'), 'A1:B18');
    expectOk(parseRangeInput('$C$10:$A$1'), 'A1:C10'); // 摆正 + 忽略 $

    for (const bad of ['A1$', 'A$$1', '$$A$1', '$A$', '$1', '$A$1$']) {
      expectFail(parseRangeInput(bad), RANGE_FORMAT_ERROR);
    }
  });

  it('反转区间按轴各自摆正：行列同时反转 / 只反转行 / 只反转列', () => {
    expectOk(parseRangeInput('C10:A1'), 'A1:C10');
    expectOk(parseRangeInput('A10:C1'), 'A1:C10');
    expectOk(parseRangeInput('C1:A10'), 'A1:C10');
  });

  it('越界与格式错误都拒绝，错误文案指明原因（含冒号两侧必须完整、全角）', () => {
    expectFail(parseRangeInput(`A${MAX_ROW + 1}`), ROW_RANGE_ERROR);
    expectFail(parseRangeInput('A0'), ROW_FORMAT_ERROR);
    expectFail(parseRangeInput('XFE1'), COLUMN_RANGE_ERROR);
    expectFail(parseRangeInput('A1:XFE9'), COLUMN_RANGE_ERROR);

    for (const bad of ['A1C10', '1A', 'A', '', '  ', 'Ａ１', 'Ａ１：Ｃ１０']) {
      expectFail(parseRangeInput(bad), RANGE_FORMAT_ERROR);
    }
    expectFail(parseRangeInput('A1:'));
    expectFail(parseRangeInput(':C10'));
    expectFail(parseRangeInput('A1:C10:D20'));
  });
});

/* -------------------------------------------------------------------------- */
/* detectImportKind - 自动识别                                                  */
/* -------------------------------------------------------------------------- */

/** 识别规则表：`[输入, 期望类型]`（`null` = 识别不出，UI 应当禁用「导入」并显示原因） */
const DETECT_CASES: readonly (readonly [string, ImportTargetKind | null])[] = [
  // 行：单行与行区间
  ['3', 'row'],
  [' 3 ', 'row'],
  ['007', 'row'],
  ['1', 'row'],
  [String(MAX_ROW), 'row'],
  ['3:5', 'row'],
  [' 3:5 ', 'row'],
  ['5:3', 'row'],
  [`1:${MAX_ROW}`, 'row'],
  // 列：单列与列区间
  ['B', 'column'],
  ['b', 'column'],
  ['xfd', 'column'],
  ['B:D', 'column'],
  [' b:d ', 'column'],
  ['D:B', 'column'],
  ['$B:$D', 'column'],
  // 区域：单元格与矩形
  ['A1', 'range'],
  ['a1', 'range'],
  ['A1:B18', 'range'],
  [' a1:b18 ', 'range'],
  ['B2', 'range'],
  ['$A$1:$B$18', 'range'],
  [`A1:XFD${MAX_ROW}`, 'range'],
  // 识别不出：空 / 内部空格 / 全角 / 越界 / 格式错误
  ['', null],
  ['   ', null],
  ['\t', null],
  ['B : D', null],
  ['3 : 5', null],
  ['A1 : B18', null],
  ['３', null],
  ['Ｂ', null],
  ['Ａ１', null],
  ['Ａ１：Ｃ１０', null],
  ['0', null],
  ['-1', null],
  ['3.5', null],
  ['3,000', null],
  [`${MAX_ROW + 1}`, null],
  ['XFE', null],
  ['AAAA', null],
  [`A${MAX_ROW + 1}`, null],
  ['A0', null],
  ['A1C10', null],
  ['1A', null],
  ['A1:', null],
  [':D', null],
  ['3:5:7', null],
  ['A1:C10:D20', null],
  ['abc', 'column'],   // 3 个字母以内都是合法列标
  ['$', null],
  ['::', null],
];

describe('detectImportKind - 自动识别行 / 列 / 区域', () => {
  it('规则表：3/3:5→行、B/B:D→列、A1/A1:B18/$A$1:$B$18→区域，其余→null', () => {
    for (const [input, expected] of DETECT_CASES) {
      expect(detectImportKind(input), `${JSON.stringify(input)} 应识别为 ${String(expected)}`).toBe(expected);
    }
  });

  it('大小写不敏感、允许首尾空格；但内部空格 / 全角字符一律非法', () => {
    expect(detectImportKind('b:d')).toBe('column');
    expect(detectImportKind('  B:D  ')).toBe('column');
    expect(detectImportKind('$a$1:$b$18')).toBe('range');
    expect(detectImportKind('B : D')).toBeNull();
    expect(detectImportKind('B\u3000D')).toBeNull(); // 全角空格
  });

  it('"识别得出类型" ⟺ "该类型一定能解析成功"（识别与校验永不打架）', () => {
    const inputs = [...DETECT_CASES.map(([input]) => input), 'B:D', 'A1:B18', 'abc', 'Z:AA'];
    for (const input of inputs) {
      const kind = detectImportKind(input);
      if (kind === null) continue;
      const parsed =
        kind === 'row' ? parseRowInput(input) : kind === 'column' ? parseColumnInput(input) : parseRangeInput(input);
      expect(parsed.ok, `${JSON.stringify(input)} 识别为 ${kind} 却解析失败`).toBe(true);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* resolveImportDraft / describeImportTarget（对话框的纯逻辑）                    */
/* -------------------------------------------------------------------------- */

describe('resolveImportDraft - 识别 + 归一化 + 提交值', () => {
  it('行：3→kind=row value=3；3:5→kind=row value=3:5（提交值不再补冒号）', () => {
    expect(resolveImportDraft('3')).toEqual({
      ok: true,
      kind: 'row',
      value: '3',
      submitValue: '3',
      label: '第 3 行',
    });
    expect(resolveImportDraft(' 3:5 ')).toEqual({
      ok: true,
      kind: 'row',
      value: '3:5',
      submitValue: '3:5',
      label: '第 3:5 行',
    });
  });

  it('列：B→B；b:d→B:D（提交值与显示文案都规范化）', () => {
    expect(resolveImportDraft('B')).toEqual({
      ok: true,
      kind: 'column',
      value: 'B',
      submitValue: 'B',
      label: 'B 列',
    });
    expect(resolveImportDraft('b:d')).toEqual({
      ok: true,
      kind: 'column',
      value: 'B:D',
      submitValue: 'B:D',
      label: 'B:D 列',
    });
  });

  it('区域：单个单元格提交时补成 A1:A1（老契约）；区间原样；$ 被忽略', () => {
    expect(resolveImportDraft('A1')).toEqual({
      ok: true,
      kind: 'range',
      value: 'A1',
      submitValue: 'A1:A1',
      label: '区域 A1',
    });
    expect(resolveImportDraft('A1:B18')).toEqual({
      ok: true,
      kind: 'range',
      value: 'A1:B18',
      submitValue: 'A1:B18',
      label: '区域 A1:B18',
    });
    expect(resolveImportDraft('$A$1:$B$18')).toEqual({
      ok: true,
      kind: 'range',
      value: 'A1:B18',
      submitValue: 'A1:B18',
      label: '区域 A1:B18',
    });
  });

  it('识别不出时给出中文原因（空 / 内部空格 / 全角 / 越界），且一定不是 ok', () => {
    for (const bad of ['', '   ', 'B : D', 'Ａ１', '３', '0', 'XFE', 'XFE1', 'A1C10', 'A1:', '3:5:7']) {
      const result = resolveImportDraft(bad);
      expect(result.ok, `${JSON.stringify(bad)} 不该被识别`).toBe(false);
      if (result.ok) continue;
      expect(result.error).toMatch(/[\u4e00-\u9fa5]/);
      expect(result.error).not.toMatch(/undefined|NaN/);
    }

    const spaced = resolveImportDraft('B : D');
    if (!spaced.ok) expect(spaced.error).toContain('空格');
    const empty = resolveImportDraft('  ');
    if (!empty.ok) expect(empty.error).toContain('请输入');
  });

  it('describeImportTarget 的三种文案（识别结果那行）', () => {
    expect(describeImportTarget('row', '3')).toBe('第 3 行');
    expect(describeImportTarget('row', '3:5')).toBe('第 3:5 行');
    expect(describeImportTarget('column', 'B:D')).toBe('B:D 列');
    expect(describeImportTarget('range', 'A1:B18')).toBe('区域 A1:B18');
  });
});

/* -------------------------------------------------------------------------- */
/* a1OfRow / a1OfColumn                                                        */
/* -------------------------------------------------------------------------- */

describe('a1OfRow / a1OfColumn', () => {
  it('合法行号拼成 N:N（最大行也合法）；非法行号（0 / 负数 / 小数 / NaN / Infinity / 超范围）→ 空串', () => {
    expect(a1OfRow(1)).toBe('1:1');
    expect(a1OfRow(3)).toBe('3:3');
    expect(a1OfRow(MAX_ROW)).toBe(`${MAX_ROW}:${MAX_ROW}`);
    for (const bad of [0, -3, 1.5, Number.NaN, Number.POSITIVE_INFINITY, MAX_ROW + 1]) {
      expect(a1OfRow(bad), `${bad} 不是合法行号`).toBe('');
    }
  });

  it('列标拼成 B:B（大小写不敏感、最大列合法），已是区间时原样规范化；非法列标 → 空串', () => {
    expect(a1OfColumn('B')).toBe('B:B');
    expect(a1OfColumn('b')).toBe('B:B');
    expect(a1OfColumn(MAX_COLUMN_LETTERS)).toBe('XFD:XFD');
    expect(a1OfColumn('b:d')).toBe('B:D');
    expect(a1OfColumn('D:B')).toBe('B:D');
    for (const bad of ['', 'XFE', 'B2', 'Ｂ']) {
      expect(a1OfColumn(bad), `${bad} 不是合法列标`).toBe('');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* 交付契约（类型 + 文案 + 边界常量）                                            */
/* -------------------------------------------------------------------------- */

describe('导出契约', () => {
  it('ParseResult 是判别联合，错误信息是中文短句且不含 NaN/undefined 脏值；边界常量与 Excel 一致', () => {
    expect(parseRowInput('3')).toHaveProperty('ok', true);
    expect(parseRowInput('0')).toHaveProperty('ok', false);
    expect(parseRowInput('3')).toHaveProperty('value');
    expect(parseRowInput('0')).toHaveProperty('error');

    const samples = [parseRowInput('abc'), parseColumnInput('XFE'), parseRangeInput('A1C10')];
    for (const sample of samples) {
      if (sample.ok) throw new Error('这里应当全部失败');
      expect(sample.error).toMatch(/[\u4e00-\u9fa5]/);
      expect(sample.error).not.toMatch(/undefined|NaN/);
    }

    expect(MAX_ROW).toBe(1_048_576);
    expect(MAX_COLUMN_LETTERS).toBe('XFD');
  });
});

/* -------------------------------------------------------------------------- */
/* ImportDialog 渲染契约（SSR 静态渲染冒烟）                                    */
/* -------------------------------------------------------------------------- */
/* 本工程 vitest 是 node 环境、没有 jsdom，所以用 renderToStaticMarkup 把组件     */
/* 渲成 HTML 字符串来守"渲染契约"（testid / 默认值 / 无选区禁用）。              */
/* 交互（打字、自动识别、ESC、焦点、取选区）由 e2e 覆盖；组件没有独立的 .tsx     */
/* 测试文件，因为本次交付只允许新建被点名的文件。                                */
/* 注意：本文件是 .ts，不能用 JSX，只能用 createElement。                        */

function renderDialog(open: boolean, currentSelectionA1: string | null, keepSource = true): string {
  return renderToStaticMarkup(
    createElement(ImportDialog, {
      open,
      onClose: () => undefined,
      currentSelectionA1,
      sheetName: '成绩表',
      onSubmit: () => undefined,
      usedRangeA1: 'A1:H14',
      keepSource,
      onKeepSourceChange: () => undefined,
    }),
  );
}

/** 取出带某个 testid 的那个标签（用来看 disabled / checked 这类布尔属性） */
function tagOf(html: string, testId: string): string {
  const index = html.indexOf(`data-testid="${testId}"`);
  if (index < 0) return '';
  return html.slice(html.lastIndexOf('<', index), html.indexOf('>', index) + 1);
}

describe('ImportDialog 渲染契约（SSR）', () => {
  it('open=false 时不渲染任何内容', () => {
    expect(renderDialog(false, null)).toBe('');
  });

  it('open=true 时渲染面板骨架与全部 testid；无用元素「跳过空单元格」已删除', () => {
    const html = renderDialog(true, null);
    expect(tagOf(html, 'import-dialog')).not.toBe('');
    expect(html).toContain('role="dialog"');
    expect(html).toContain('aria-modal="true"');
    for (const testId of [
      'import-close', 'import-input', 'import-detected', 'import-error',
      'import-cancel', 'import-confirm', 'import-use-selection',
      // 一键把整表已用区域填进输入框（空内容在底层仍会被跳过）
      'import-whole-sheet',
      // 新增：实时预览 + "加入工作区后保留表格内容"选项
      'import-preview', 'import-preview-summary', 'import-keep-source', 'import-keep-hint',
    ]) {
      expect(html, `缺少 ${testId}`).toContain(`data-testid="${testId}"`);
    }
    /**
     * 用户点名要求删掉的元素：那个"跳过空单元格（固定行为，不可关闭）"的禁用复选框。
     * 它不可交互、又容易被误读成"能关"，属于纯占位的无用元素。
     */
    expect(html, '「跳过空单元格」复选框必须已被移除').not.toContain('import-skip-empty');
    expect(html, '不该再出现"不可关闭"这种说明书式文案').not.toContain('不可关闭');
    for (const gone of ['import-tab-row', 'import-tab-column', 'import-tab-range']) {
      expect(html, `${gone} 应当已被移除`).not.toContain(gone);
    }
  });

  it('只有一个输入框且默认空：识别行为占位、确认禁用、保留选项默认勾选、无选区时取选区禁用', () => {
    const html = renderDialog(true, null);
    expect(html.match(/data-testid="import-input"/g)).toHaveLength(1);
    expect(html).toContain('例如 3、B:D 或 A1:B18');
    expect(html).toContain('识别为：—');
    expect(tagOf(html, 'import-confirm')).toContain('disabled');
    expect(html).toContain('加入工作区');
    /* 预览空态：说清楚这一块是干什么的，而不是一句空保证 */
    expect(html).toContain('识别出范围后，这里显示会加入哪些单元格');
    /* "加入工作区后保留表格内容"默认勾选（等于复制），并且可以随时改 */
    expect(tagOf(html, 'import-keep-source')).toContain('checked');
    expect(tagOf(html, 'import-keep-source'), '这个选项是可交互的，不能禁用').not.toContain('disabled');
    expect(html).toContain('加入工作区后保留表格内容');
    expect(tagOf(html, 'import-use-selection')).toContain('disabled');
    expect(html).toContain('无选区');
    expect(html).toContain('成绩表'); // 工作表名出现在说明文案里
  });

  it('取消勾选"保留表格内容"＝剪切语义：复选框不勾选、提示文案改口', () => {
    const html = renderDialog(true, null, false);
    expect(tagOf(html, 'import-keep-source')).not.toContain('checked');
    expect(html).toContain('已取消勾选');
    expect(html).toContain('剪切');
    expect(html).toContain('撤销一次即可恢复');
  });

  it('有选区时"当前选区"可用且显示 A1', () => {
    const html = renderDialog(true, 'B2:D5');
    expect(tagOf(html, 'import-use-selection')).not.toContain('disabled');
    expect(html).toContain('B2:D5');
  });

  it('「整个已用区域」只填输入框（不再一键直提）：有已用区域时可用并标出区域，空表时禁用', () => {
    const html = renderDialog(true, null);
    expect(tagOf(html, 'import-whole-sheet')).not.toContain('disabled');
    expect(html).toContain('整个已用区域');
    expect(html).toContain('A1:H14'); // usedRangeA1 显示在按钮上
    // 只是"填范围"的按钮：必须是 type="button"（提交只走底部的确认键）
    expect(tagOf(html, 'import-whole-sheet')).toContain('type="button"');

    const emptySheet = renderToStaticMarkup(
      createElement(ImportDialog, {
        open: true,
        onClose: () => undefined,
        currentSelectionA1: null,
        sheetName: '空表',
        onSubmit: () => undefined,
        usedRangeA1: null,
        keepSource: true,
        onKeepSourceChange: () => undefined,
      }),
    );
    expect(tagOf(emptySheet, 'import-whole-sheet'), '空表应禁用').toContain('disabled');
    expect(emptySheet).toContain('空表');
  });
});

/* -------------------------------------------------------------------------- */
/* 面板预览文案（纯函数；渲染态由 e2e 覆盖）                                     */
/* -------------------------------------------------------------------------- */

/**
 * 面板的"目标范围换算"（`src/workspace/import-target.ts`）。
 *
 * 这段换算**提交与预览共用**，所以它错了会出现两种症状：预览的数字与实际加入的格子不一致，
 * 或者"行/列"直接展开成整列 100 万行把上限打爆。这里用一张假的表（已用区域 A1:H21）钉住口径。
 */
describe('「加入工作区」面板的目标范围换算（resolveImportRanges）', () => {
  const sheet = {
    getDataRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0, endRow: 20, endColumn: 7 }) }),
  } as unknown as FWorksheet;

  it('行 → 收敛到已用区域的列范围（不会取整行 16384 列）', () => {
    expect(resolveImportRanges(sheet, 'row', '3')).toEqual({ a1List: ['A3:H3'] });
    expect(resolveImportRanges(sheet, 'row', '3:5')).toEqual({ a1List: ['A3:H5'] });
  });

  it('列 → 收敛到已用区域的行范围（不会取整列 1048576 行）', () => {
    expect(resolveImportRanges(sheet, 'column', 'B')).toEqual({ a1List: ['B1:B21'] });
    expect(resolveImportRanges(sheet, 'column', 'B:D')).toEqual({ a1List: ['B1:D21'] });
  });

  it('区域 → 单块补成 A1:A1，多块原样展开', () => {
    expect(resolveImportRanges(sheet, 'range', 'A1')).toEqual({ a1List: ['A1'] });
    expect(resolveImportRanges(sheet, 'range', 'A1:B2 D4:E5')).toEqual({ a1List: ['A1:B2', 'D4:E5'] });
  });

  it('认不出的区域 → 报错而不是静默返回空表', () => {
    expect(resolveImportRanges(sheet, 'range', '')).toHaveProperty('error');
  });
});

/**
 * 预览的"只扫可能真有内容的那一片"（`clipRect`）。
 *
 * 为什么必须有它：面板是边打字边预览的，`A1:A1048576` 这种写法如果逐格扫会直接钉死主线程
 * （实测风险：百万次取值）。已用区域之外按定义没有内容，所以先求交集不改变"将加入多少格"。
 */
describe('预览范围裁剪（clipRect）', () => {
  const data = { startRow: 0, startColumn: 0, endRow: 20, endColumn: 7 }; // A1:H21

  it('完全在已用区域内 → 原样返回', () => {
    expect(clipRect({ startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 }, data)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 5,
      endColumn: 3,
    });
  });

  it('越过已用区域 → 裁到交集（整列写法不会去扫 100 万行）', () => {
    expect(clipRect({ startRow: 0, startColumn: 0, endRow: 1_048_575, endColumn: 0 }, data)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 20,
      endColumn: 0,
    });
    expect(clipRect({ startRow: 5, startColumn: 2, endRow: 1_048_575, endColumn: 16_383 }, data)).toEqual({
      startRow: 5,
      startColumn: 2,
      endRow: 20,
      endColumn: 7,
    });
  });

  it('完全没有交集 → null（这一块全是空）', () => {
    expect(clipRect({ startRow: 30, startColumn: 0, endRow: 40, endColumn: 3 }, data)).toBeNull();
    expect(clipRect({ startRow: 0, startColumn: 9, endRow: 3, endColumn: 12 }, data)).toBeNull();
  });
});

describe('ImportDialog 预览文案', () => {
  const preview = {
    blocks: 2,
    total: 12,
    empty: 156,
    limit: 500,
    overLimit: false,
    samples: [{ a1: 'A1', text: '张三' }],
  };

  it('describeImportPreview：报出加入数 / 跳过数 / 块数', () => {
    expect(describeImportPreview(null)).toBeNull();
    expect(describeImportPreview({ ...preview, blocks: 1 })).toBe('将加入 12 个单元格 · 跳过 156 个空内容');
    expect(describeImportPreview(preview)).toBe('将加入 12 个单元格 · 跳过 156 个空内容 · 2 块区域');
    expect(describeImportPreview({ ...preview, blocks: 1, empty: 0 })).toBe('将加入 12 个单元格');
    // 选区写得极大时（空格子数是上亿级别的）不再展示那一项，避免吓人的大数字
    expect(describeImportPreview({ ...preview, blocks: 1, empty: PREVIEW_EMPTY_CAP })).toBe('将加入 12 个单元格');
  });

  it('importConfirmLabel：确认键上带格数，没算出预览时退回通用文案', () => {
    expect(importConfirmLabel(null)).toBe('加入工作区');
    expect(importConfirmLabel({ ...preview, total: 0 })).toBe('加入工作区');
    expect(importConfirmLabel(preview)).toBe('加入 12 格');
  });
});
