/**
 * **渲染错误兜底**（用户反馈："打开文件后，页面是空白的，但是网页刷新后又会出现"）。
 *
 * 白屏有两种可能，这一层治的是其中一种、也是后果最难看的一种：
 *  - **渲染期抛异常**：React 18 在没有错误边界时会把整棵树卸载 → `#root` 变空 → 整页空白。
 *    刷新后因为走"会话恢复"这条干净路径，看起来"自己好了"，用户就以为是随机的。
 *    现在任何未捕获的渲染错误都会落成一个可读的提示卡（含错误摘要 + 「重新加载」按钮），
 *    数据不受影响：会话在 IndexedDB 里，重新加载会把它恢复回来。
 *  - 另一种是"活动单元与画布绑定错位"（页面在、表格区空白）——那个由 `App` 里的
 *    `ensureActiveUnitRendered` 自愈处理。
 *
 * 注意：这里**只**兜渲染错误；事件回调/异步里的异常不会走到这里（那些各自有 try/catch 与日志）。
 */
import { Component, type ErrorInfo, type ReactNode } from 'react';

import { log } from '../p0/log';
import './error-boundary.css';

interface Props {
  children: ReactNode;
}

interface State {
  error: Error | null;
}

/** 兜底界面（导出是为了单测能直接渲染它） */
export function ErrorFallback({ error }: { error: Error | null }): JSX.Element {
  return (
    <div className="app-crash" data-testid="app-crash" role="alert">
      <div className="app-crash-card">
        <h1 className="app-crash-title">界面出错了</h1>
        <p className="app-crash-text">
          已经打开的文件与编辑都保存在本机（浏览器数据库），不会丢。点下面重新加载即可回到现场。
        </p>
        <pre className="app-crash-detail" data-testid="app-crash-detail">
          {error?.message ?? '未知错误'}
        </pre>
        <div className="app-crash-actions">
          <button
            type="button"
            className="app-crash-btn is-primary"
            data-testid="app-crash-reload"
            onClick={() => window.location.reload()}
          >
            重新加载
          </button>
        </div>
      </div>
    </div>
  );
}

export class AppErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    log('app:render-error', {
      message: String(error?.message ?? error),
      componentStack: (info?.componentStack ?? '').slice(0, 400),
    });
  }

  render(): ReactNode {
    if (this.state.error) return <ErrorFallback error={this.state.error} />;
    return this.props.children;
  }
}
