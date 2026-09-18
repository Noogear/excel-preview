/**
 * 撤销/重做快捷键的"按键识别"（纯函数）。
 *
 * 这个模块存在的理由（实测踩坑，详见模块头注释）：Univer 的 ShortcutService 在引导时就把
 * Ctrl+Z / Ctrl+Y 注册到 window 的 capture 阶段，我们晚注册就永远排在它后面——
 * 结果"Univer 先撤一步、我们的账本原地不动"，用户看到的是"工作区撤不掉"。
 * 因此我们在引导前抢注第一顺位，并用 `stopImmediatePropagation()` 截下这次按键。
 * 这里只测"哪些组合归我们管"，避免误伤别的快捷键（例如 Ctrl+Shift+Z 是重做、Ctrl+Alt+Z 不是我们的）。
 */
import { describe, expect, it } from 'vitest';

import { hasUndoShortcutHandler, isUndoRedoKey, setUndoShortcutHandler } from '../../src/shell/undo-shortcut';

describe('isUndoRedoKey: 哪些按键组合由我们接管', () => {
  it('Ctrl+Z / Cmd+Z 是撤销', () => {
    expect(isUndoRedoKey({ key: 'z', ctrlKey: true })).toBe(true);
    expect(isUndoRedoKey({ key: 'z', metaKey: true })).toBe(true);
  });

  it('大写 Z 也认（按住 Shift 时 key 会是大写，重做由上层再判断 shiftKey）', () => {
    expect(isUndoRedoKey({ key: 'Z', ctrlKey: true })).toBe(true);
  });

  it('Ctrl+Y 是重做', () => {
    expect(isUndoRedoKey({ key: 'y', ctrlKey: true })).toBe(true);
    expect(isUndoRedoKey({ key: 'Y', metaKey: true })).toBe(true);
  });

  it('不带修饰键、或带 Alt 的组合不归我们管（保留浏览器/系统行为）', () => {
    expect(isUndoRedoKey({ key: 'z' })).toBe(false);
    expect(isUndoRedoKey({ key: 'z', ctrlKey: true, altKey: true })).toBe(false);
  });

  it('其它按键一律放行（避免给每次敲键做多余判断）', () => {
    for (const key of ['a', 'c', 'v', 'x', 'Enter', 'ArrowUp', '']) {
      expect(isUndoRedoKey({ key, ctrlKey: true }), key).toBe(false);
    }
  });
});

describe('setUndoShortcutHandler: 挂载/卸载时的接管状态', () => {
  it('塞入处理函数后标记为"已接管"，卸载（传 null）后恢复', () => {
    const before = hasUndoShortcutHandler();
    setUndoShortcutHandler(() => {});
    expect(hasUndoShortcutHandler()).toBe(true);
    setUndoShortcutHandler(null);
    expect(hasUndoShortcutHandler()).toBe(before);
  });
});
