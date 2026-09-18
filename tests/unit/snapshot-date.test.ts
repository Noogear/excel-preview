/**
 * 日期单元格"搬家不毁值"回归测试。
 *
 * 背景（用户实测反馈）：把日期单元格拖进工作区再拖回去 / 与别的格互换时，会弹出
 * `sheets-ui.info.error` + `sheets-ui.info.forceStringInfo` 两个**原始 i18n key**。
 * 追根因时发现真正的元凶是我们自己的取值口径：日期格的 `getValues()` 返回 `Date` 对象，
 * 旧的归一化走到"对象分支" → `String(date)`，写回的是
 * `"Wed Jan 01 2025 08:00:00 GMT+0800 (中国标准时间)"` 这种字符串——
 * ①日期语义被毁成文本；②Univer 认为"文本格里塞了怪值"，于是弹提醒。
 *
 * 修法：①优先取 raw 值（日期格 raw 就是序列号）；②万一只能拿到 `Date`，也转回序列号。
 * 这里把两条路都钉死，避免以后有人"顺手"改回归一化。
 */
import { describe, expect, it } from 'vitest';

import { buildValueMatrixFromRect, dateToExcelSerial, extractCellItems, isEmptyContent, isSnapshotEmpty } from '../../src/workspace/snapshot';

/* -------------------------------------------------------------------------- */
/* 假 worksheet：只实现 snapshot.ts 真正调用到的那几个方法                        */
/* -------------------------------------------------------------------------- */

interface FakeCell {
  /** 单格取值（Facade 的 getValue：日期格可能给 Date） */
  getValue: () => unknown;
  display?: string;
  formula?: string;
  style?: Record<string, unknown>;
}

function fakeSheet(cells: Record<string, FakeCell>, options: { raw?: unknown[][] } = {}) {
  const keys = Object.keys(cells);
  const rows = keys.map((k) => Number(k.split(':')[0]));
  const cols = keys.map((k) => Number(k.split(':')[1]));
  const rect = {
    startRow: Math.min(...rows),
    startColumn: Math.min(...cols),
    endRow: Math.max(...rows),
    endColumn: Math.max(...cols),
  };

  const single = (row: number, col: number) => {
    const cell = cells[`${row}:${col}`];
    return {
      getValue: () => cell?.getValue() ?? null,
      getDisplayValue: () => cell?.display ?? '',
      getFormula: () => cell?.formula ?? null,
      getCellStyleData: () => cell?.style ?? null,
      getRange: () => ({ startRow: row, startColumn: col, endRow: row, endColumn: col }),
    };
  };

  const area = {
    getRange: () => rect,
    getValues: () =>
      Array.from({ length: rect.endRow - rect.startRow + 1 }, (_, r) =>
        Array.from({ length: rect.endColumn - rect.startColumn + 1 }, (_, c) => cells[`${rect.startRow + r}:${rect.startColumn + c}`]?.getValue() ?? null),
      ),
    getFormulas: () =>
      Array.from({ length: rect.endRow - rect.startRow + 1 }, (_, r) =>
        Array.from({ length: rect.endColumn - rect.startColumn + 1 }, (_, c) => cells[`${rect.startRow + r}:${rect.startColumn + c}`]?.formula ?? null),
      ),
    // raw 只在显式提供时存在，用来验证"优先 raw，没有才退回 getValues()"
    ...(options.raw ? { getRawValues: () => options.raw } : {}),
  };

  return {
    getSheetId: () => 'sheet-1',
    getSheetName: () => 'Sheet1',
    getColumnWidth: () => 100,
    getRowHeight: () => 24,
    getActiveRange: () => area,
    // 1 参（'A1'）= 整块区域；2 参（row,col）= 单格；4 参 = 指定矩形区域
    getRange: (...args: unknown[]) => (args.length >= 3 || typeof args[0] === 'string' ? area : single(Number(args[0]), Number(args[1]))),
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
}

const asSheet = (sheet: unknown) => sheet as never;

/* -------------------------------------------------------------------------- */
/* dateToExcelSerial                                                           */
/* -------------------------------------------------------------------------- */

describe('dateToExcelSerial', () => {
  it('对齐 Excel 1900 日期系统的已知锚点', () => {
    expect(dateToExcelSerial(new Date(1970, 0, 1))).toBe(25569); // Unix 纪元
    expect(dateToExcelSerial(new Date(2025, 0, 1))).toBe(45658);
    expect(dateToExcelSerial(new Date(1900, 2, 1))).toBe(61); // 1900-03-01（跳过 Excel 的假闰日）
  });

  it('按本地年月日计算：同一天的不同时刻得到同一个序列号', () => {
    const midnight = new Date(2025, 0, 1, 0, 0, 0);
    const lateNight = new Date(2025, 0, 1, 23, 59, 59);
    expect(dateToExcelSerial(midnight)).toBe(dateToExcelSerial(lateNight));
    // 时区漂移会把日期推早/推晚一天，这里要求严格等于锚点
    expect(dateToExcelSerial(lateNight)).toBe(45658);
  });

  it('输出整数且逐日递增 1', () => {
    expect(Number.isInteger(dateToExcelSerial(new Date(2025, 5, 15, 13, 45)))).toBe(true);
    expect(dateToExcelSerial(new Date(2025, 5, 16)) - dateToExcelSerial(new Date(2025, 5, 15))).toBe(1);
  });
});

/* -------------------------------------------------------------------------- */
/* extractCellItems：日期格进工作区不能退化成字符串                              */
/* -------------------------------------------------------------------------- */

describe('extractCellItems 的日期处理', () => {
  it('getValue() 返回 Date 时写回序列号，绝不写 "Wed Jan 01 2025 ..." 这种字符串', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => new Date(2025, 0, 1), display: '2025/1/1' },
    });

    const { items } = extractCellItems(asSheet(sheet), 'A1');

    expect(items).toHaveLength(1);
    expect(items[0].values[0][0]).toBe(45658);
    expect(typeof items[0].values[0][0]).toBe('number');
    // 预览文本仍用显示值（给用户看的是"2025/1/1"而不是 45658）
    expect(items[0].cells[0][0].text).toBe('2025/1/1');
    expect(JSON.stringify(items[0])).not.toContain('GMT');
  });

  it('优先 raw 值：单格 getValue() 给的是格式化字符串 "2025-01-01"，raw 才是序列号', () => {
    // 这正是用户遇到的场景：Facade 单格取值对日期格返回显示文本，
    // 用它写回就会把日期永久变成文本；raw 值是序列号 45658。
    const sheet = fakeSheet(
      { '0:0': { getValue: () => '2025-01-01', display: '2025-01-01' } },
      { raw: [[45658]] },
    );

    const { items } = extractCellItems(asSheet(sheet), 'A1');

    expect(items[0].values[0][0]).toBe(45658);
    expect(typeof items[0].values[0][0]).toBe('number');
    expect(items[0].cells[0][0].text, '预览文本仍用显示值').toBe('2025-01-01');
  });

  it('空单元格不进工作区（底层跳过）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => null },
      '0:1': { getValue: () => '张三' },
    });

    const { items, skippedEmpty } = extractCellItems(asSheet(sheet), 'A1:B1');

    expect(items.map((i) => i.source.a1)).toEqual(['B1']);
    expect(skippedEmpty).toBe(1);
  });

  /**
   * 用户实测反馈："多选后复制/剪切到工作区以后，不会默认跳过空单元格了"。
   *
   * 根因：旧口径只判 `v === null`，而 Univer 里"看起来是空的"格子有三种形态：
   * ①用户把内容删光 → `v` 是**空字符串**；②只打了空格/全角空格；③从没写过 → `null`。
   * 前两种都会变成工作区里的一张白卡片。
   */
  it('"空"的三种形态都不进工作区：null / 空字符串 / 只有空格（含全角）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => '', display: '' }, // 用户在编辑器里把内容删光
      '0:1': { getValue: () => ' ', display: ' ' }, // 只有半角空格
      '0:2': { getValue: () => '\u3000', display: '\u3000' }, // 全角空格（中文表格常见）
      '0:3': { getValue: () => '\n', display: '\n' }, // 只有换行
      '0:4': { getValue: () => null }, // 从没写过
      '0:5': { getValue: () => '张三', display: '张三' }, // 唯一有内容的
    });

    const { items, skippedEmpty } = extractCellItems(asSheet(sheet), 'A1:F1');

    expect(items.map((i) => i.source.a1), '只有 F1 应该进工作区').toEqual(['F1']);
    expect(skippedEmpty).toBe(5);
  });

  it('0 与 false 是内容，不能被当成空（数学/判断题的答案）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => 0, display: '0' },
      '0:1': { getValue: () => false, display: 'FALSE' },
      '0:2': { getValue: () => '0', display: '0' },
    });

    const { items, skippedEmpty } = extractCellItems(asSheet(sheet), 'A1:C1');

    expect(items.map((i) => i.source.a1)).toEqual(['A1', 'B1', 'C1']);
    expect(skippedEmpty).toBe(0);
  });

  it('公式算出来是空串的格子仍要保留（公式本身就是内容）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => '', formula: '=IF(1>2,"x","")', display: '' },
    });

    const { items, skippedEmpty } = extractCellItems(asSheet(sheet), 'A1');

    expect(items).toHaveLength(1);
    expect(items[0].formulas[0][0]).toBe('IF(1>2,"x","")');
    expect(skippedEmpty).toBe(0);
  });

  it('公式格带上公式（写回时不会被当成纯值丢掉）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => 3, formula: '=1+2', display: '3' },
    });

    const { items } = extractCellItems(asSheet(sheet), 'A1');

    // 模型里的公式**不带前导 `=`**（buildValueMatrix 写回时再加回来），这是既有约定
    expect(items[0].formulas[0][0]).toBe('1+2');
    expect(items[0].values[0][0]).toBe(3);
  });
});

/* -------------------------------------------------------------------------- */
/* isEmptyContent / isSnapshotEmpty：全应用唯一的"空内容"口径                     */
/* -------------------------------------------------------------------------- */

describe('空内容判定（工作区不放空条目）', () => {
  it('isEmptyContent：null/undefined/空串/纯空白算空，其余算有内容', () => {
    for (const empty of [null, undefined, '', ' ', '\u3000', '\n', '\t  \u3000']) {
      expect(isEmptyContent(empty), JSON.stringify(empty)).toBe(true);
    }
    for (const value of [0, -1, 3.14, false, true, '张三', ' 0 ']) {
      expect(isEmptyContent(value), JSON.stringify(value)).toBe(false);
    }
    // 有公式就一定有内容（哪怕算出来是空串）
    expect(isEmptyContent('', 'IF(1>2,"x","")')).toBe(false);
    expect(isEmptyContent(null, '1+1')).toBe(false);
  });

  it('isSnapshotEmpty：整份快照全空才算空条目；只要有一格有内容就保留', () => {
    expect(isSnapshotEmpty({ values: [[null, '']], formulas: [[null, null]] })).toBe(true);
    expect(isSnapshotEmpty({ values: [[null, '李四']], formulas: [[null, null]] })).toBe(false);
    expect(isSnapshotEmpty({ values: [[null]], formulas: [[null, '1+1']] })).toBe(false);
    expect(isSnapshotEmpty({ values: [], formulas: [] })).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* buildValueMatrixFromRect：互换/粘贴的取值口径                                 */
/* -------------------------------------------------------------------------- */

describe('buildValueMatrixFromRect 的日期处理', () => {
  it('有 raw 值（序列号）时优先用 raw，不产生 Date → 字符串降级', () => {
    const sheet = fakeSheet(
      {
        '0:0': { getValue: () => new Date(2025, 0, 1), display: '2025/1/1' },
        '0:1': { getValue: () => 45658, display: '2025/1/1' },
      },
      { raw: [[45658, 45658]] },
    );

    const matrix = buildValueMatrixFromRect(asSheet(sheet), {
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 1,
    });

    expect(matrix).toEqual([[{ v: 45658 }, { v: 45658 }]]);
  });

  it('拿不到 raw 时把 Date 归一化成序列号（兜底路径）', () => {
    const sheet = fakeSheet({
      '0:0': { getValue: () => new Date(2025, 0, 1), display: '2025/1/1' },
      '0:1': { getValue: () => null },
    });

    const matrix = buildValueMatrixFromRect(asSheet(sheet), {
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 1,
    });

    expect(matrix[0][0]).toEqual({ v: 45658 });
    // 空值走 emptyContentCell（只清内容、保留格式），不是 null 整格删除
    expect(matrix[0][1]).toMatchObject({ v: null, f: null });
  });
});
