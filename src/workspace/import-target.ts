/**
 * 「加入工作区」面板的**目标范围换算**：把面板识别出来的输入（行 / 列 / 区域，区域可多块）
 * 换算成一张 A1 区域表（纯函数，只读表，不产生任何快照与历史）。
 *
 * 为什么单独一个文件（而不是放在 `App.tsx` 里）：
 * `App.tsx` 只导出 `App` 这一个组件时，Vite 的 React Fast Refresh 才能热替换它；
 * 一旦多出一个"非组件导出"，插件会判定"有不可热替换的导出"，于是**整页刷新**——
 * 用户在开发态改一行代码就会看到页面重载（并触发 `tests/e2e/hot-reload.spec.ts` 里那条
 * "改代码后不该出现灵魂标签"的守卫从"热更新"退化成"整页重载"）。
 * 顺带好处：这段换算逻辑可以脱离 Univer 启动流程被单测直接覆盖。
 *
 * 提交与**预览共用**同一个换算：预览显示的数字和真正加入的格子必须出自同一段代码，
 * 否则就会出现"预览说 12 格、实际加入 13 格"这种最招人烦的不一致。
 */
import type { FWorksheet } from '@univerjs/sheets/facade';

import { parseA1Range } from '../importer/table-style';
import { parseRangeList, rectsToA1 } from '../interaction/selection-ranges';
import type { ImportTargetKind } from './import-parse';
import { rectToA1 } from './snapshot';

/**
 * 行 / 列要**收敛到已用数据范围**：直接取 `A:A` 是整列 100 万行，必然超上限。
 * 区域则原样展开（多块用空格分隔的写法由面板负责规范化）。
 */
export function resolveImportRanges(
  sheet: FWorksheet,
  kind: ImportTargetKind,
  value: string,
): { a1List: string[] } | { error: string } {
  const data = sheet.getDataRange().getRange();
  const a1List: string[] = [];
  if (kind === 'row') {
    // 行号可以是单行 `3`，也可以是行区间 `3:5`（面板的自动识别允许区间）
    const [fromText, toText] = value.includes(':') ? value.split(':') : [value, value];
    const from = Math.max(0, Number(fromText) - 1);
    const to = Math.max(0, Number(toText) - 1);
    a1List.push(
      rectToA1({
        startRow: Math.min(from, to),
        startColumn: data.startColumn,
        endRow: Math.max(from, to),
        endColumn: data.endColumn,
      }),
    );
  } else if (kind === 'column') {
    const [fromText, toText] = value.includes(':') ? value.split(':') : [value, value];
    const from = parseA1Range(`${fromText}1`)?.startCol ?? 0;
    const to = parseA1Range(`${toText}1`)?.startCol ?? from;
    a1List.push(rectToA1({ startRow: data.startRow, startColumn: from, endRow: data.endRow, endColumn: to }));
  } else {
    const { rects } = parseRangeList(value);
    if (rects.length === 0) return { error: '无法识别所选区域' };
    a1List.push(...rectsToA1(rects));
  }
  return { a1List };
}
