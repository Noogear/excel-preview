import { createRoot } from 'react-dom/client';

import { installUndoShortcutPriority } from './shell/undo-shortcut';
import './shell/shell.css';

const container = document.getElementById('root');
if (!container) throw new Error('#root not found');

/**
 * 抢在 Univer 引导之前占住 Ctrl+Z / Ctrl+Y 的第一顺位。
 *
 * 必须在这里（而不是 App 的 effect 里）：Univer 的 ShortcutService 引导时就往 window 上
 * 注册了 capture 监听器，同 target 同阶段按注册顺序执行，晚注册的只能"看着它先撤"——
 * 那正是"工作区撤不掉"的根因，详见 `src/shell/undo-shortcut.ts` 的说明。
 */
installUndoShortcutPriority();

/**
 * 注意：这里**故意不使用 `StrictMode`**。
 *
 * P0 实测（Playwright 诊断）：在 React 18 StrictMode 下 effect 会 mount → cleanup → mount，
 * 而 Univer 的渲染容器（含 docs 编辑器容器这类全局单例 DOM）在 dispose 后无法被第二个实例正确接管，
 * 结果是页面上 `document.querySelectorAll('canvas').length === 0`——表格完全不渲染，
 * 但应用状态却显示"已就绪"（因为 API 层调用都成功了），极难排查。
 *
 * 因此 Univer 宿主必须由我们自己控制生命周期：单实例、显式创建与销毁。
 */

/**
 * **懒加载应用主体**（加载速度优化）。
 *
 * 以前这里静态 `import { App } from './App'`，于是"整个表格引擎（分包后仍有 vendor-univer 数 MB）"
 * 都在**入口 chunk** 里：浏览器必须先下完、解析完这一大坨，才能画出第一帧 —— 用户看到的就是长时间白屏。
 * 现在入口只做两件小事（接管快捷键 + 挂载 React），`App`（连同 Univer）走动态 import：
 *  - `index.html` 里的骨架先画出来（"正在加载表格引擎…"），用户立刻有反馈；
 *  - 大块并行下载，下完立即挂载，不需要用户刷新或等待额外步骤。
 *
 * 打点：`app:bundle-loaded`（入口脚本开始执行）、`app:chunk-loaded`（App 分块到位、即将挂载），
 * 配合 `univer:booted` 日志就能算出"骨架 → 应用 → 可编辑"三段耗时（e2e 的启动预算就是用它们断言的）。
 */
performance.mark('app:bundle-loaded');

const { App } = await import('./App');
performance.mark('app:chunk-loaded');

createRoot(container).render(<App />);
