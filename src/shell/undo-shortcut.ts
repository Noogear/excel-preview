/**
 * 撤销/重做快捷键的**优先级**（谁先拿到 Ctrl+Z）。
 *
 * 背景（用户反馈"无法撤回对工作区的操作，或工作区的操作没被记录"的最后一层根因）：
 *
 * 1. Univer 的 `ShortcutService` 在引导时用
 *    `window.addEventListener('keydown', handler, { capture: true })` 注册了 Ctrl+Z / Ctrl+Y
 *    （`@univerjs/ui` 的 `UndoShortcutItem`/`RedoShortcutItem`，绑定 `Ctrl(4096)|90`/`Ctrl|89`）。
 * 2. 我们的快捷键接管是 React effect 里注册的，**必然晚于**它。
 * 3. 同一个 target（window）同一个阶段（capture）的多个监听器按**注册顺序**依次执行，
 *    `stopPropagation()` 同节点无效，`preventDefault()` 也拦不住后面那个监听器。
 *
 * 于是实测出现这种"看起来像玄学"的现象：按一次 Ctrl+Z，Univer 先撤了表格（画布变了），
 * 我们的账本却还在原位——因为等我们执行 `api.undo()` 时，Univer 的撤销栈已经空了，
 * 命令返回 false，账本索引原地不动。再按一次就"什么都没发生"，用户看到的就是
 * "工作区撤不掉 / 撤销乱了"。
 *
 * 修法：**在 Univer 引导之前**就把我们的监听器挂上去（本模块由 `main.tsx` 在 render 前调用），
 * 抢到第一顺位；真正处理时用 `stopImmediatePropagation()` 把这次按键彻底截下来，
 * 由我们按账本顺序统一撤销（表格动作交给 Univer 的撤销栈，工作区动作走我们自己的快照）。
 *
 * 这里刻意做成一个极小的模块（不依赖 React / Univer）：
 *  - `install()` 只注册一个"转发器"，真正的处理函数由 App 在挂载后 `setHandler` 塞进来；
 *  - 没有处理函数（未挂载 / 已卸载）时完全不干预，事件照常往下走。
 */

/** 一次撤销/重做按键请求的处理器；返回 `true` 表示"这次按键我接管了"。 */
export type UndoShortcutHandler = (event: KeyboardEvent) => void;

let installed = false;
let handler: UndoShortcutHandler | null = null;

/** 只认这几种组合（其余按键连判断都不做，避免给每次敲键加开销） */
export function isUndoRedoKey(event: {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
}): boolean {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return false;
  const key = event.key?.toLowerCase?.() ?? '';
  return key === 'z' || key === 'y';
}

/**
 * 在 **Univer 引导之前**调用（`main.tsx`）：把我们的转发器注册成第一顺位的
 * window 捕获监听器。重复调用无副作用。
 */
export function installUndoShortcutPriority(): void {
  if (installed) return;
  installed = true;
  window.addEventListener(
    'keydown',
    (event: KeyboardEvent) => {
      if (!handler) return;
      if (!isUndoRedoKey(event)) return;
      handler(event);
    },
    { capture: true },
  );
}

/** App 挂载后塞入真正的处理函数（卸载时传 `null`） */
export function setUndoShortcutHandler(next: UndoShortcutHandler | null): void {
  handler = next;
}

/** 仅供测试/诊断：当前是否已接管 */
export function hasUndoShortcutHandler(): boolean {
  return handler !== null;
}
