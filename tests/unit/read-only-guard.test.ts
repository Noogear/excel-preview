/**
 * 只读约束闸门的规则单测。
 *
 * 这是"只允许改文字内容"的**第一道防线**（默认拒绝），规则错一个就可能放行破坏性命令，
 * 所以这里把放行/拒绝两侧都钉死，包括"未知命令一律拒绝"。
 */
import { describe, expect, it } from 'vitest';

import { isCommandAllowed } from '../../src/univer/read-only-guard';

describe('放行：不写入文档的操作与纯内容编辑', () => {
  const allowed = [
    // 撤销重做
    'univer.command.undo',
    'univer.command.redo',
    // 打字 / 公式栏 / 我们的互换与写回
    'sheet.command.set-range-values',
    'sheet.mutation.set-range-values',
    // 剪贴板（样式由 mutation 层剥离）
    'sheet.command.paste',
    'sheet.command.paste-value',
    'sheet.command.paste-by-short-key',
    'sheet.command.copy',
    'sheet.command.cut',
    // 内容清理与传播
    'sheet.command.clear-selection-content',
    'sheet.command.auto-clear-content',
    'sheet.command.copy-down',
    'sheet.command.copy-right',
    // 选择与导航
    'sheet.command.move-selection',
    'sheet.command.move-selection-enter-tab',
    'sheet.command.select-all',
    'sheet.command.expand-selection',
    'sheet.command.scroll-to-cell',
    'sheet.command.scroll-view',
    // 滚轮滚动（漏了它滚轮就完全不动——用户实测反馈）
    'sheet.command.set-scroll-relative',
    // 视图与切表
    'sheet.command.change-zoom-ratio',
    // 缩放滑块/百分比菜单用的就是这条（只放行 change-zoom-ratio 时滑块拖不动——用户实测反馈）
    'sheet.command.set-zoom-ratio',
    'sheet.command.set-worksheet-activate',
    // 不落盘的 operation
    'sheet.operation.set-selections',
    'sheet.operation.set-scroll',
    'sheet.operation.set-cell-edit-visible',
    'sheet.operation.set-activate-cell-edit',
    'sheet.operation.scroll-to-range',
    'sheet.operation.mark-dirty-row-auto-height',
    'sheet.operation.cancel-mark-dirty-row-auto-height',
    // 单元格编辑器（docs 子系统）：两种 id 写法都要放行，否则打字进不去编辑器
    'doc.command.insert-text',
    'doc.command.delete-text',
    'doc.command-replace-snapshot',
    'doc.operation.set-selections',
    'doc.operation.scroll-to-cursor',
    // 公式重算记账
    'formula.mutation.set-formula-calculation-start',
    'formula.mutation.set-trigger-formula-calculation-start',
  ];

  for (const id of allowed) {
    it(`放行 ${id}`, () => {
      expect(isCommandAllowed(id)).toBe(true);
    });
  }
});

describe('拒绝：样式 / 结构 / 对象 / 子表', () => {
  const denied = [
    // 样式家族
    'sheet.command.set-style',
    'sheet.command.set-bold',
    'sheet.command.set-background-color',
    'sheet.command.set-text-color',
    'sheet.command.set-italic',
    'sheet.command.set-border',
    'sheet.command.set-font-size',
    'sheet.command.set-text-wrap',
    'sheet.command.reset-background-color',
    'sheet.command.clear-selection-format',
    'sheet.command.paste-format',
    'sheet.operation.set-format-painter',
    // 结构与尺寸
    'sheet.command.set-col-width',
    'sheet.command.set-row-height',
    'sheet.command.delta-column-width',
    'sheet.command.delta-row-height',
    'sheet.command.paste-col-width',
    'sheet.command.insert-row',
    'sheet.command.insert-col-before',
    'sheet.command.delete-range-move-left',
    'sheet.command.remove-row',
    'sheet.command.append-row',
    'sheet.command.add-worksheet-merge',
    'sheet.command.remove-worksheet-merge',
    'sheet.command.move-rows',
    'sheet.command.move-range',
    'sheet.command.reorder-range',
    'sheet.command.auto-fill',
    'sheet.command.refill',
    // 子表与工作簿结构
    'sheet.command.insert-sheet',
    'sheet.command.remove-sheet',
    'sheet.command.copy-sheet',
    'sheet.command.set-worksheet-name',
    'sheet.operation.rename-sheet',
    'sheet.command.set-tab-color',
    'sheet.command.set-frozen',
    'sheet.command.toggle-gridlines',
    'sheet.command.set-worksheet-protection',
    // 对象与规则
    'sheet.command.insert-image',
    'sheet.command.delete-note',
    'sheet.command.create-or-update-note',
    'sheet.command.add-hyper-link',
    'sheet.command.set-data-validation',
    'sheet.command.add-conditional-formatting-rule',
    'sheet.command.delete-table',
    'sheet.command.set-filter-criteria',
    'sheet.command.sort-range',
    // 其它危险动作
    'sheet.command.clear-selection-all',
    'sheet.command.repeat-last-action',
    // 编辑器内的富文本格式化（"只改文字"不允许加入格式）
    'doc.command.set-bold',
    'doc.command.set-italic',
    'doc.command.set-underline',
    'doc.command.set-font-size',
    'doc.command.set-fore-color',
    'doc.command.set-back-color',
    'doc.command.set-paragraph-align',
  ];

  for (const id of denied) {
    it(`拒绝 ${id}`, () => {
      expect(isCommandAllowed(id)).toBe(false);
    });
  }
});

describe('默认拒绝', () => {
  it('未知命名空间一律拒绝', () => {
    expect(isCommandAllowed('some.unknown.command')).toBe(false);
    expect(isCommandAllowed('')).toBe(false);
    expect(isCommandAllowed('univer.command.something-new')).toBe(false);
  });

  it('未列入白名单的 sheet 命令一律拒绝', () => {
    expect(isCommandAllowed('sheet.command.some-future-feature')).toBe(false);
    expect(isCommandAllowed('sheet.mutation.set-something-else')).toBe(false);
  });

  it('拒绝优先级高于放行（含 format 的 operation 也拦掉）', () => {
    expect(isCommandAllowed('sheet.operation.set-format-painter')).toBe(false);
    // 反例：普通 operation 仍然放行
    expect(isCommandAllowed('sheet.operation.set-selections')).toBe(true);
  });
});
