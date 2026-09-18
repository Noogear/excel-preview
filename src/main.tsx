import { createRoot } from 'react-dom/client';

import { installUndoShortcutPriority } from './shell/undo-shortcut';
import './shell/shell.css';
// 兜底卡片（崩溃卡 / 加载失败卡）的样式**静态引入**：加载失败那条路径上 `ErrorBoundary` 模块
// 可能正是没下来的那一块，样式不能跟着它一起缺席（否则卡片是裸文本，按钮也点不准）。
import './shell/error-boundary.css';

declare global {
  interface Window {
    /** `index.html` 里的看门狗（应用挂载成功后撤掉它，避免"卡住了"的误报） */
    __bootWatchdog?: () => void;
  }
}

const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('#root not found');
/** 非空版本（下面的闭包里也要用，显式标好类型省得跟 narrow 较劲） */
const container: HTMLElement = rootElement;

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

/**
 * **分块加载失败要能自救**（用户反馈："打开文件后，页面是空白的，但是网页刷新后又会出现"）。
 *
 * 这一类"刷新就好"的现象里，有一个**能稳定复现**的成因：静态托管（GitHub Pages/Gitee Pages）
 * 每次发布都会换掉分块的 hash，而用户浏览器里可能还留着**旧的 index.html**（或弱网/代理把分块截断）。
 * 旧 HTML 指向的 `App-<旧hash>.js` 已经不存在 → 动态 import 直接 reject。以前这里没有任何兜底：
 * 骨架屏会**永远停在"正在加载表格引擎…"**，用户看到的就是"白屏/打不开"，而按 F5 拿到新 HTML 就好了
 * —— 和用户描述的一模一样。
 *
 * 现在：① 先**自动重来一次**（reload 会重新取 index.html，指向新 hash，通常这一步就好了）；
 * ② 还不行就摆一张可读的卡片（含原因与「重新加载」），而不是让用户对着灰块干等。
 * 标记写在 `sessionStorage`：**起来之后清掉**，所以以后真坏了还能再自动救一次；存储不可用时**不**自动
 * 重试（否则可能刷成死循环）。
 */
const RETRY_FLAG = 'app:chunk-retry';

function readRetryFlag(): boolean | null {
  try {
    return sessionStorage.getItem(RETRY_FLAG) === '1';
  } catch {
    return null; // 隐私模式等场景拿不到存储：当"已经试过"，直接给卡片，绝不死循环
  }
}

function writeRetryFlag(value: boolean): void {
  try {
    if (value) sessionStorage.setItem(RETRY_FLAG, '1');
    else sessionStorage.removeItem(RETRY_FLAG);
  } catch {
    /* 存储不可用就算了：不影响主流程 */
  }
}

/** 连应用主体都下不来时的兜底卡片（不依赖 React —— 它可能正是没下来的那一块） */
function showBootFailure(detail: string): void {
  window.__bootWatchdog?.();
  const card = document.createElement('div');
  card.className = 'app-crash';
  card.setAttribute('data-testid', 'boot-failed');
  card.setAttribute('role', 'alert');
  const box = document.createElement('div');
  box.className = 'app-crash-card';
  const title = document.createElement('h1');
  title.className = 'app-crash-title';
  title.textContent = '没能加载应用';
  const text = document.createElement('p');
  text.className = 'app-crash-text';
  text.textContent = '通常是网络中断，或者浏览器缓存了旧版本的页面（发布后地址里的版本号会变）。点下面重新加载即可。';
  const pre = document.createElement('pre');
  pre.className = 'app-crash-detail';
  pre.setAttribute('data-testid', 'boot-failed-detail');
  pre.textContent = detail;
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'app-crash-btn is-primary';
  button.setAttribute('data-testid', 'app-crash-reload');
  button.textContent = '重新加载';
  button.addEventListener('click', () => window.location.reload());
  const actions = document.createElement('div');
  actions.className = 'app-crash-actions';
  actions.append(button);
  box.append(title, text, pre, actions);
  card.append(box);
  container.replaceChildren(card);
}

async function boot(): Promise<void> {
  try {
    const { App } = await import('./App');
    const { AppErrorBoundary } = await import('./shell/ErrorBoundary');
    performance.mark('app:chunk-loaded');
    writeRetryFlag(false); // 起来了就清标记：下次再坏还能自动救一次
    createRoot(container).render(
      <AppErrorBoundary>
        <App />
      </AppErrorBoundary>,
    );
    window.__bootWatchdog?.();
  } catch (error) {
    const detail = String((error as { message?: string } | null)?.message ?? error);
    if (readRetryFlag() === false) {
      writeRetryFlag(true);
      window.location.reload();
      return;
    }
    showBootFailure(detail);
  }
}

void boot();
