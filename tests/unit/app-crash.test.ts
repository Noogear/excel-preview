/**
 * 渲染错误兜底（`AppErrorBoundary` / `ErrorFallback`）单测。
 *
 * 用户反馈过"打开文件后页面是空白的，刷新后又出现"：渲染期抛异常时，React 18 在没有错误边界
 * 的情况下会**卸载整棵树**，`#root` 变空 → 整页空白；刷新后走"会话恢复"这条干净路径又正常了。
 * 这里钉住三件事：① 有错时渲染成可读的兜底卡（含「重新加载」按钮与错误摘要）；
 * ② 错误状态推导只认第一手错误对象；③ 没错时原样渲染子节点（不能把正常路径吃掉）。
 *
 * node 环境没有 jsdom，所以用 `renderToStaticMarkup` 做 SSR 冒烟（同 import-parse.test.ts 的做法）。
 */
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { AppErrorBoundary, ErrorFallback } from '../../src/shell/ErrorBoundary';

describe('渲染错误兜底', () => {
  it('ErrorFallback：给出人话、错误摘要与「重新加载」按钮', () => {
    const html = renderToStaticMarkup(createElement(ErrorFallback, { error: new Error('boom: 单元格越界') }));
    expect(html).toContain('data-testid="app-crash"');
    expect(html).toContain('界面出错了');
    expect(html, '要说明数据没丢').toContain('不会丢');
    expect(html).toContain('data-testid="app-crash-reload"');
    expect(html).toContain('重新加载');
    expect(html, '错误摘要要显示出来，便于回报').toContain('boom: 单元格越界');
  });

  it('ErrorFallback：错误为空时也不炸（显示"未知错误"）', () => {
    const html = renderToStaticMarkup(createElement(ErrorFallback, { error: null }));
    expect(html).toContain('未知错误');
  });

  it('getDerivedStateFromError：把错误对象收进 state（只做状态推导，不产生副作用）', () => {
    const error = new Error('render failed');
    const state = AppErrorBoundary.getDerivedStateFromError(error);
    expect(state.error).toBe(error);
  });

  it('无错误时原样渲染子节点（正常路径不被兜底吃掉）', () => {
    const html = renderToStaticMarkup(
      createElement(AppErrorBoundary, null, createElement('span', { 'data-testid': 'child' }, '正常内容')),
    );
    expect(html).toContain('data-testid="child"');
    expect(html).toContain('正常内容');
    expect(html).not.toContain('app-crash');
  });
});
