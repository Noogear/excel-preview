/**
 * e2e 共享工具（不是测试文件：不含 test()，不会被 Playwright 收集）。
 *
 * 所有断言仍然走"行为面"：`window.__p0` 暴露的 Facade 数据 / 事件日志 / 真实键鼠手势，
 * 尽量不碰 DOM 类名与内部实现细节。
 */
import { expect, type Page } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const FIXTURE_DIR = join(process.cwd(), 'fixtures');

export const fixturePath = (name: string): string => join(FIXTURE_DIR, name);

export interface LogEntry {
  t: number;
  kind: string;
  detail?: Record<string, unknown>;
}

export interface FeatureCounts {
  conditionalFormats: number;
  dataValidations: number;
  hyperlinks: number;
  notes: number;
  images: number;
  failed: number;
  skipped: number;
}

export interface ImportSummary {
  fileName: string;
  sheets: string[];
  unsupported: string[];
  warnings: string[];
  /** 已解析并原样保留、但不影响预览（打印设置）；不是缺陷 */
  preserved: string[];
  featureIssues?: string[];
  featureCounts?: FeatureCounts | null;
  parsedFeatures?: { cf: number; dv: number; links: number; notes: number; tables: number; images: number };
  cells?: number;
}

/**
 * 引导：等 __p0 钩子就位 + app:ready 落日志。
 *
 * 超时给到 60 秒（原来 30 秒）：本机 6 并发跑主项目时，开发服务器要为**每个 worker** 第一次
 * 变换整张模块图（Univer 那一坨 5MB 的 vendor），偶发会超过 30 秒，于是出现"启动超时"的**假红**
 * （实测 `blank-page-guard` ⑤ / `context-menu` ③ / `coordinates` ⓪ 三条单跑都绿、并跑偶红）。
 * 用例本身的预算本来就是 60–180 秒，这里放宽只是不再让"机器忙"表现成"功能坏了"。
 */
export async function waitForBoot(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForFunction(() => Boolean((window as never as { __p0?: unknown }).__p0), null, { timeout: 60_000 });
  await page.waitForFunction(
    () => {
      const hooks = (window as never as { __p0?: { findLast: (k: string) => unknown } }).__p0;
      return Boolean(hooks?.findLast('app:ready') ?? hooks?.findLast('p0:ready'));
    },
    null,
    { timeout: 60_000 },
  );
}

/** 走界面上的文件输入框导入夹具，并等到导入摘要换成这份文件 */
export async function importFixture(
  page: Page,
  fileName: string,
  options: { features?: boolean } = {},
): Promise<ImportSummary> {
  await page.setInputFiles('[data-testid="file-input"]', fixturePath(fileName));
  await page.waitForFunction(
    (name) => {
      const hooks = (window as never as { __p0: { getImportSummary: () => ImportSummary | null } }).__p0;
      return hooks.getImportSummary()?.fileName === name;
    },
    fileName,
    { timeout: 60_000 },
  );
  if (options.features) {
    // 特性应用（条件格式/批注/图片…）在导入摘要之后完成，等报告出现再读，避免读到上一份的残留
    await page.waitForFunction(
      (name) => {
        const hooks = (window as never as {
          __p0: { getImportSummary: () => ImportSummary | null; getFeatureReport: () => unknown };
        }).__p0;
        return hooks.getImportSummary()?.fileName === name && hooks.getFeatureReport() !== null;
      },
      fileName,
      { timeout: 60_000 },
    );
  }
  const summary = await page.evaluate(() =>
    (window as never as { __p0: { getImportSummary: () => ImportSummary | null } }).__p0.getImportSummary(),
  );
  if (!summary) throw new Error('未拿到导入摘要');
  return summary;
}

/** 用 Vite 模块图直接清库，保证持久化用例可重复（路径用变量传入，绕过静态模块解析） */
export async function clearSession(page: Page): Promise<void> {
  await page.evaluate(async (modulePath: string) => {
    const mod = (await import(/* @vite-ignore */ modulePath)) as { createSessionStore: () => { clear: () => Promise<void> } };
    await mod.createSessionStore().clear();
  }, '/src/persistence/session.ts');
}

/** 工作表主画布（Univer 内部另有一个 0×0 的 docs 画布，裸 `canvas` 会命中它） */
export function sheetCanvas(page: Page) {
  return page.locator('#univer-container canvas[id^="univer-sheet-main-canvas"]').first();
}

export async function canvasBox(page: Page): Promise<{ x: number; y: number; width: number; height: number }> {
  const locator = sheetCanvas(page);
  await locator.waitFor({ state: 'visible' });
  const box = await locator.boundingBox();
  if (!box) throw new Error('canvas 不可见');
  return box;
}

/**
 * 取"用户自己的真实表格"路径（几条可选用例会拿它做真机验证：真实课程表/座位表里的合并块、
 * 打印设置这类东西，合成夹具覆盖不到）。
 *
 * **为什么不在代码里写死路径**：那份文件属于用户本机、且含真实姓名学号；仓库是要公开的，
 * 路径与文件名都不该跟着代码走。所以按下面的顺序找，找不到就**跳过**相关用例（不会假红）：
 *   ① 环境变量 `USER_EXCEL`；
 *   ② `tests/e2e/user-fixture.local`（一行路径；`*.local` 已在 .gitignore 里，不进仓库）。
 * 本机要跑这几条用例：
 *   `"$env:USER_EXCEL='D:\路径\我的表.xlsx'; npm run e2e"` 或把路径写进 user-fixture.local。
 */
export function userFixturePath(hint?: string): string | null {
  const candidates: string[] = [];
  const fromEnv = process.env.USER_EXCEL?.trim();
  if (fromEnv) candidates.push(fromEnv);
  try {
    const local = join(dirname(fileURLToPath(import.meta.url)), 'user-fixture.local');
    if (existsSync(local)) {
      for (const line of readFileSync(local, 'utf8').split(/\r?\n/)) {
        const text = line.trim();
        if (text !== '' && !text.startsWith('#')) candidates.push(text);
      }
    }
  } catch {
    /* 读不到就当没配置 */
  }
  if (candidates.length === 0) return null;
  if (!hint) return candidates[0];
  // 文件名里带提示词的优先（一个本地文件可以同时登记"座位表"和"课程表"两份）
  return candidates.find((candidate) => basename(candidate).includes(hint)) ?? candidates[0];
}

export async function readLog(page: Page): Promise<LogEntry[]> {
  return page.evaluate(() =>
    (window as never as { __p0: { log: LogEntry[] } }).__p0.log.map((entry) => ({
      t: entry.t,
      kind: entry.kind,
      detail: entry.detail,
    })),
  );
}

export async function logKinds(page: Page): Promise<string[]> {
  return page.evaluate(() => (window as never as { __p0: { log: LogEntry[] } }).__p0.log.map((entry) => entry.kind));
}

export async function lastDetail(page: Page, kind: string): Promise<Record<string, unknown> | null> {
  return page.evaluate(
    (k) =>
      ((window as never as { __p0: { log: LogEntry[] } }).__p0.log.filter((entry) => entry.kind === k).pop()?.detail ??
        null) as Record<string, unknown> | null,
    kind,
  );
}

/**
 * 等**应用自己**探测完本地转换桥，并返回它的结论。
 *
 * 为什么不能让用例自己去 `fetch('/api/bridge/health')`：那是"服务端能不能用"，
 * 而界面上点得动点不动看的是**应用探测完之后的 state**（探测要启动一次 Excel，2–6 秒）。
 * 用例里再 fetch 一次，就会在应用还没探测完时误判"桥可用"，然后奇怪为什么走了"打不开"分支（踩过）。
 */
export async function waitForBridgeStatus(
  page: Page,
  timeoutMs = 45_000,
): Promise<{ available: boolean; excel?: string; reason?: string }> {
  await page.waitForFunction(
    () => (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.some((entry) => entry.kind === 'app:bridge-status'),
    null,
    { timeout: timeoutMs },
  );
  const detail = await lastDetail(page, 'app:bridge-status');
  return {
    available: Boolean(detail?.available),
    excel: typeof detail?.excel === 'string' ? detail.excel : undefined,
    reason: typeof detail?.reason === 'string' ? detail.reason : undefined,
  };
}

export async function clearLog(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
  });
}

export async function selectionA1(page: Page): Promise<string | null> {
  return page.evaluate(() => (window as never as { __p0: { getSelectionA1: () => string | null } }).__p0.getSelectionA1());
}

export function a1ToRowCol(a1: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!match) throw new Error(`非法 A1 记号：${a1}`);
  let col = 0;
  for (const ch of match[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(match[2]) - 1, col: col - 1 };
}

export async function valueAt(page: Page, a1: string): Promise<string | null> {
  const { row, col } = a1ToRowCol(a1);
  return page.evaluate(
    ({ r, c }) =>
      (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0.getDisplayValue(r, c),
    { r: row, c: col },
  );
}

export async function styleAt(page: Page, a1: string): Promise<string> {
  const { row, col } = a1ToRowCol(a1);
  return page.evaluate(
    ({ r, c }) =>
      JSON.stringify(
        (window as never as { __p0: { getCellStyle: (r: number, c: number) => unknown } }).__p0.getCellStyle(r, c) ?? null,
      ),
    { r: row, c: col },
  );
}

export async function activateSheet(page: Page, name: string): Promise<void> {
  await page.evaluate((sheetName: string) => {
    (window as never as { __p0: { activateSheet: (n: string) => void } }).__p0.activateSheet(sheetName);
  }, name);
  await page.waitForFunction(
    (sheetName) => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => { getSheetName: () => string } | null } }).__p0.getActiveSheet();
      return sheet?.getSheetName() === sheetName;
    },
    name,
    { timeout: 15_000 },
  );
}

export type InteractionMode = 'select' | 'drag' | 'click-swap';

export async function switchMode(page: Page, mode: InteractionMode): Promise<void> {
  const button = page.locator(`[data-testid="toolbar-mode-${mode}"]`);
  await button.click();
  await expect(button).toHaveAttribute('aria-checked', 'true');
}
