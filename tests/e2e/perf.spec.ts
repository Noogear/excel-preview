/**
 * ⑩ 极致性能：**浏览器侧**实测（用户真正感知的指标）。
 *
 * 与 `tests/unit/perf-pipeline.test.ts`（Node 侧管线吞吐）互补，这里量的是端到端体验：
 *   ① 导入墙钟时间：从选中文件 → 解析 → 适配 → Univer 建工作簿 → 首帧出现内容
 *   ② 长任务（longtask）：导入期间主线程有没有被单次 >50ms 的任务卡住
 *   ③ JS 堆占用：百万格文件导入后的内存水位
 *   ④ 拖动帧率：按住拖动过程中的 rAF 间隔 p50/p95
 *
 * 夹具：`npm run bench:fixture`（默认 5 万行 × 20 列 = 100 万格）；缺失时跳过。
 * 阈值刻意宽松：只拦"数量级劣化"，不拦机器快慢（避免在慢机/CI 上变 flaky）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const FIXTURE = join(process.cwd(), 'fixtures', process.env.PERF_FIXTURE ?? 'bench-large.xlsx');
const hasFixture = existsSync(FIXTURE);

interface LogEntry {
  t: number;
  kind: string;
  detail?: Record<string, unknown>;
}

async function waitForBoot(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForFunction(() => Boolean((window as never as { __p0?: unknown }).__p0), null, { timeout: 30_000 });
  await page.waitForFunction(
    () => Boolean((window as never as { __p0?: { findLast: (k: string) => unknown } }).__p0?.findLast('app:ready')),
    null,
    { timeout: 30_000 },
  );
}

test.describe('⑩ 性能：浏览器端实测', () => {
  test.skip(!hasFixture, '请先运行 npm run bench:fixture 生成 fixtures/bench-large.xlsx');

  test('百万格文件：导入墙钟 / 首帧 / 长任务 / 堆占用', async ({ page }) => {
    test.setTimeout(300_000);
    await waitForBoot(page);

    // 先挂上"主线程静默"采样器与长任务观察器。
    //
    // 注意：**不要**用 `getImageData` 轮询画面来判断首帧——超大 canvas 的像素回读本身
    // 就是几百毫秒到数秒的长任务，测量工具会把被测对象污染掉（实测过：数据完全不可信）。
    // 这里用 rAF 间隔衡量"主线程什么时候空出来"：连续 3 帧 < 32ms 即认为可交互。
    await page.evaluate(() => {
      const w = window as never as {
        __longtasks: number[];
        __frames: Array<{ t: number; d: number }>;
        __sampling: boolean;
        __last: number;
      };
      w.__longtasks = [];
      w.__frames = [];
      w.__sampling = true;
      w.__last = performance.now();
      try {
        new PerformanceObserver((list) => {
          for (const entry of list.getEntries()) w.__longtasks.push(Math.round(entry.duration));
        }).observe({ entryTypes: ['longtask'] });
      } catch {
        /* 浏览器不支持 longtask 时不阻断测量 */
      }
      const tick = (now: number): void => {
        w.__frames.push({ t: now, d: now - w.__last });
        w.__last = now;
        if (w.__sampling) requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });

    await page.setInputFiles('[data-testid="file-input"]', FIXTURE);
    await page.waitForFunction(
      () => Boolean((window as never as { __p0?: { findLast: (k: string) => unknown } }).__p0?.findLast('import:done')),
      null,
      { timeout: 180_000 },
    );

    // 等到主线程连续 3 帧都 < 32ms（≈30FPS），作为"真正可交互"的时刻
    const quietMs = await page.evaluate(async () => {
      const hooks = (window as never as { __p0: { log: LogEntry[] } }).__p0;
      const start = hooks.log.filter((e) => e.kind === 'import:start').pop()?.t ?? 0;
      const w = window as never as { __frames: Array<{ t: number; d: number }> };
      const deadline = performance.now() + 60_000;
      while (performance.now() < deadline) {
        await new Promise((resolve) => requestAnimationFrame(resolve));
        const frames = w.__frames;
        const n = frames.length;
        if (n >= 3) {
          const last3 = frames.slice(n - 3);
          if (last3.every((f) => f.d < 32)) return Math.round(last3[0].t - start);
        }
      }
      return -1;
    });

    const metrics = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { log: LogEntry[] } }).__p0;
      const start = hooks.log.filter((e) => e.kind === 'import:start').pop()?.t ?? 0;
      const done = hooks.log.filter((e) => e.kind === 'import:done').pop();
      const detail = (done?.detail ?? {}) as Record<string, number>;
      const memory = (performance as unknown as { memory?: { usedJSHeapSize: number } }).memory;
      const longtasks = (window as never as { __longtasks?: number[] }).__longtasks ?? [];
      return {
        totalMs: Math.round((done?.t ?? 0) - start),
        parseMs: detail.parseMs ?? -1,
        adaptMs: detail.adaptMs ?? -1,
        renderMs: detail.renderMs ?? -1,
        featureMs: detail.featureMs ?? -1,
        cells: detail.cells ?? -1,
        heapMb: memory ? Math.round(memory.usedJSHeapSize / 1024 / 1024) : -1,
        longtasks,
      };
    });

    const longtaskTotal = metrics.longtasks.reduce((sum, v) => sum + v, 0);
    const worst = metrics.longtasks.length > 0 ? Math.max(...metrics.longtasks) : 0;

    // eslint-disable-next-line no-console
    console.log(
      [
        '',
        '┌─ ⑩ 性能实测：浏览器端导入 ──────────────────────────────',
        `│ 夹具            ${FIXTURE.replace(/^.*[\\/]/, '')}`,
        `│ 单元格          ${metrics.cells.toLocaleString('en-US')}`,
        `│ 端到端墙钟      ${metrics.totalMs} ms（选中文件 → 解析/适配/建簿完成）`,
        `│   ├ 解析        ${metrics.parseMs} ms`,
        `│   ├ 适配        ${metrics.adaptMs} ms`,
        `│   ├ 建工作簿    ${metrics.renderMs} ms`,
        `│   └ 特性应用    ${metrics.featureMs} ms`,
        `│ 主线程静默      ${quietMs} ms（连续 3 帧 < 32ms，即真正可交互）`,
        `│ 长任务          ${metrics.longtasks.length} 次，合计 ${longtaskTotal} ms，最长 ${worst} ms`,
        `│ JS 堆           ${metrics.heapMb} MB`,
        '└─────────────────────────────────────────────────────────',
        '',
      ].join('\n'),
    );

    expect(metrics.cells, '应导入十万级以上的单元格').toBeGreaterThan(100_000);
    expect(metrics.totalMs, '端到端导入墙钟应 < 60s').toBeLessThan(60_000);
    expect(quietMs, '主线程应在 60s 内空出来').toBeGreaterThan(0);
    expect(quietMs, '主线程静默应 < 60s').toBeLessThan(60_000);
    expect(worst, '单次长任务不应超过 20s（超出说明主线程被整段阻塞）').toBeLessThan(20_000);
  });

  test('百万格文件：拖动过程帧率', async ({ page }) => {
    test.setTimeout(300_000);
    await waitForBoot(page);
    await page.setInputFiles('[data-testid="file-input"]', FIXTURE);
    await page.waitForFunction(
      () => Boolean((window as never as { __p0?: { findLast: (k: string) => unknown } }).__p0?.findLast('import:done')),
      null,
      { timeout: 180_000 },
    );
    await page.waitForTimeout(500);

    const canvas = page.locator('#univer-container canvas[id^="univer-sheet-main-canvas"]').first();
    const box = await canvas.boundingBox();
    if (!box) throw new Error('canvas 不可见');

    await page.evaluate(() => {
      const w = window as never as { __frames: number[]; __sampling: boolean; __last: number };
      w.__frames = [];
      w.__sampling = true;
      w.__last = performance.now();
      const tick = (now: number): void => {
        w.__frames.push(now - w.__last);
        w.__last = now;
        if (w.__sampling) requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });

    // 用真实鼠标按住并拖动（长按拖动是产品的核心交互之一）
    const startX = box.x + 180;
    const startY = box.y + 140;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    for (let i = 1; i <= 60; i += 1) {
      await page.mouse.move(startX + i * 4, startY + Math.round(Math.sin(i / 5) * 30));
      await page.waitForTimeout(16);
    }
    await page.mouse.up();

    const stats = await page.evaluate(() => {
      const w = window as never as { __frames: number[]; __sampling: boolean };
      w.__sampling = false;
      const frames = w.__frames.filter((v) => v > 0 && v < 2000).slice(2);
      const sorted = [...frames].sort((a, b) => a - b);
      const pick = (q: number): number => sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * q))] ?? 0;
      return {
        count: sorted.length,
        p50: Math.round(pick(0.5) * 10) / 10,
        p95: Math.round(pick(0.95) * 10) / 10,
        worst: Math.round((sorted[sorted.length - 1] ?? 0) * 10) / 10,
      };
    });

    const fps = (ms: number): number => (ms > 0 ? Math.round(1000 / ms) : 0);
    // eslint-disable-next-line no-console
    console.log(
      [
        '',
        '┌─ ⑩ 性能实测：拖动帧率（百万格） ──────────────────────',
        `│ 采样帧数        ${stats.count}`,
        `│ 帧间隔 p50      ${stats.p50} ms（≈ ${fps(stats.p50)} FPS）`,
        `│ 帧间隔 p95      ${stats.p95} ms（≈ ${fps(stats.p95)} FPS）`,
        `│ 最差帧          ${stats.worst} ms`,
        '└─────────────────────────────────────────────────────────',
        '',
      ].join('\n'),
    );

    expect(stats.count, '应采到足够帧').toBeGreaterThan(20);
    expect(stats.p95, 'p95 帧间隔应 < 200ms').toBeLessThan(200);
  });
});
