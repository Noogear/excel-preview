/**
 * 交互原语单测（node 环境，无 DOM）。
 *
 * 只覆盖可以在 node 下验证的**纯逻辑**：
 * - `easeOutCubic` / `easeOutBack`：端点、单调性、回弹过冲；
 * - `prefersReducedMotion()`：无 window 环境返回 false 且不抛异常。
 *
 * 长按判定（`long-press.ts`）已随交互改版移除——搬运改成"按住直接拖"，不再需要长按计时器。
 * DOM 相关的 `drag-ghost` / `drag-controller` 不在这里测（由 Playwright e2e 覆盖真实手势）。
 */
import { describe, expect, it } from 'vitest';

import { easeOutBack, easeOutCubic } from '../../src/interaction/fly-in';
import { prefersReducedMotion } from '../../src/interaction/drag-ghost';

describe('缓动函数', () => {
  it('easeOutCubic 端点与单调性；easeOutBack 端点与过冲（>1 但有限）', () => {
    expect(easeOutCubic(0)).toBe(0);
    expect(easeOutCubic(1)).toBe(1);

    let prev = -Infinity;
    for (let i = 0; i <= 100; i += 1) {
      const v = easeOutCubic(i / 100);
      expect(Number.isFinite(v)).toBe(true);
      expect(v).toBeGreaterThanOrEqual(prev);
      expect(v).toBeGreaterThanOrEqual(0);
      expect(v).toBeLessThanOrEqual(1);
      prev = v;
    }

    expect(easeOutBack(0)).toBe(0);
    expect(easeOutBack(1)).toBe(1);

    let maxValue = -Infinity;
    for (let i = 0; i <= 100; i += 1) {
      const v = easeOutBack(i / 100);
      expect(Number.isFinite(v)).toBe(true);
      maxValue = Math.max(maxValue, v);
    }
    // 回弹曲线的特征：中后段超过 1，最后收敛回 1。
    expect(maxValue).toBeGreaterThan(1);
    expect(maxValue).toBeLessThan(1.15);
  });
});

describe('prefersReducedMotion', () => {
  it('无 window 环境（node / SSR）返回 false 且不抛异常', () => {
    expect(typeof window).toBe('undefined');
    expect(() => prefersReducedMotion()).not.toThrow();
    expect(prefersReducedMotion()).toBe(false);
  });
});
