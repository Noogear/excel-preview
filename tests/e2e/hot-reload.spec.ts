/**
 * **热更新（Vite HMR / React Fast Refresh）不该多出"灵魂标签"**。
 *
 * 用户实测反馈："每次你更新代码后，标签页都会出现一个名字相同的'灵魂'标签页，
 * 这可能会影响到正常使用吗，有修复方案吗？"
 *
 * 复现出来的根因（诊断日志：`session:restore-start {tabs:1}` → `session:tabs-registered {tabs:2}`
 * → 控制台一串 React "two children with the same key"）：
 * Fast Refresh 会把 App 的引导 effect **重跑一遍**，而 refs/state 是**保留**的，
 * 但旧 Univer 实例已经在 cleanup 里 dispose 了。于是：
 *  ① `tabsDataRef` 里还记着那些标签，新实例里却没有对应 unit（点它会报 no document with unitId）——
 *     这就是"灵魂标签"：看得见、点不动；
 *  ② 启动恢复把会话里的标签**又追加了一遍**（同 id、同名）→ 标签栏出现两个同名标签。
 *
 * 修法（三处，缺一不可）：
 *  1. 识别"重跑"（`app:hot-reboot`）：把所有标签标成未实体化（切回时按字节重建）→ 不再是死标签；
 *  2. 恢复时**按 id 去重**（`session:restore-skip-registered`）→ 不可能出现同名重复；
 *  3. cleanup 里**先把现场同步写库**再拆监听 → 热更新不会吞掉"最近 1~2 秒"的编辑。
 *
 * 第二个用例另外钉住"自我修复"：旧版本可能已经把重复条目写进了会话，
 * 现在恢复时按 id 去重，读一次就自动清干净。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { clearSession, fixturePath, importFixture, waitForBoot } from './helpers';

const APP_SOURCE = resolve(process.cwd(), 'src/App.tsx');

interface TabInfo {
  id: string;
  fileName: string;
}

const tabs = (page: Page): Promise<TabInfo[]> =>
  page.evaluate(() => (window as never as { __p0: { getTabs: () => TabInfo[] } }).__p0.getTabs());

const logKinds = (page: Page): Promise<string[]> =>
  page.evaluate(() => (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind));

/** 改一次 App.tsx（触发热更新），并等到 `app:hot-reboot` 落日志；等不到就返回 false（环境不支持 HMR） */
async function hotReload(page: Page, original: string, round: number): Promise<boolean> {
  writeFileSync(APP_SOURCE, `${original}\n// hot-reload-probe ${round} ${Date.now()}\n`, 'utf8');
  try {
    await page.waitForFunction(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.some((entry) => entry.kind === 'app:hot-reboot'),
      null,
      { timeout: 15_000 },
    );
    return true;
  } catch {
    return false;
  }
}

test.describe('热更新后标签页不乱', () => {
  test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures');

  test('① 改代码（Fast Refresh）后仍然只有一个标签，且刚敲进去的内容不丢', async ({ page }) => {
    test.setTimeout(180_000);
    const original = readFileSync(APP_SOURCE, 'utf8');
    const consoleErrors: string[] = [];
    page.on('console', (msg) => {
      if (msg.type() === 'error') consoleErrors.push(msg.text());
    });

    try {
      await waitForBoot(page);
      await importFixture(page, 'fixture-styles.xlsx');
      await page.waitForTimeout(2600); // 会话轮询周期（2s）跑过一轮

      const before = await tabs(page);
      expect(before, '导入后应恰好一个标签').toHaveLength(1);

      // 造一个"刚敲进去、可能还没落盘"的编辑，再触发热更新
      await page.evaluate(() => (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0.setValue(2, 0, '热更新前写入'));
      await page.waitForTimeout(150);

      const applied = await hotReload(page, original, 1);
      test.skip(!applied, '当前环境没有触发 HMR（可能不是 vite dev 服务），跳过');
      await page.waitForTimeout(2500); // 等恢复 + 重建完成

      const after = await tabs(page);
      expect(after, '热更新后仍然只有一个标签（不能出现同名"灵魂标签"）').toHaveLength(1);
      expect(after[0].id, '还是原来那个标签（按 id 去重）').toBe(before[0].id);
      expect(new Set(after.map((tab) => tab.id)).size, 'id 不能重复').toBe(after.length);

      // 死标签的现场证据：新实例里这个 unit 必须真的被重建出来
      const state = await page.evaluate(() => {
        const p0 = (window as never as { __p0: Record<string, (...a: unknown[]) => unknown> }).__p0;
        const sheet = p0.getActiveSheet() as { getRange: (a1: string) => { getDisplayValues: () => string[][] } } | null;
        return {
          active: p0.getActiveTabId(),
          a3: sheet ? sheet.getRange('A3').getDisplayValues()[0][0] : null,
          built: (p0.getTabRuntime() as Array<{ built: boolean }>).map((entry) => entry.built),
        };
      });
      expect(state.built, '唯一那个标签应当是已实体化的').toEqual([true]);
      expect(state.a3, '热更新前刚写进去的内容不能丢').toBe('热更新前写入');

      // 再热更新一次：不能累加
      const applied2 = await hotReload(page, original, 2);
      expect(applied2).toBe(true);
      await page.waitForTimeout(2500);
      expect(await tabs(page), '连续热更新也不该累加').toHaveLength(1);

      const kinds = await logKinds(page);
      expect(kinds, '应留下"热更新重启"的日志').toContain('app:hot-reboot');
      expect(kinds, '应按 id 去重（跳过已注册的标签）').toContain('session:restore-skip-registered');
      expect(
        consoleErrors.filter((text) => text.includes('same key')),
        '不该再有"重复 key"的 React 报错',
      ).toHaveLength(0);
    } finally {
      writeFileSync(APP_SOURCE, original, 'utf8');
    }
  });

  test('② 会话里已被写坏（同一标签存了两遍）时，恢复会自动去重（自我修复）', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await clearSession(page); // 从干净状态开始，避免上一轮残留
    await page.reload();
    await waitForBoot(page);

    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(2600);
    expect(await tabs(page)).toHaveLength(1);

    // 模拟旧版本留下的坏会话：把同一个标签在 tabs 里写两遍
    const duplicated = await page.evaluate(async (modulePath: string) => {
      const mod = (await import(/* @vite-ignore */ modulePath)) as {
        createSessionStore: () => {
          load: () => Promise<{ tabs: unknown[] } | null>;
          save: (state: unknown) => Promise<void>;
        };
      };
      const store = mod.createSessionStore();
      const state = await store.load();
      if (!state) return 0;
      const broken = { ...state, tabs: [...state.tabs, ...state.tabs] };
      await store.save(broken);
      return broken.tabs.length;
    }, '/src/persistence/session.ts');
    expect(duplicated, '应成功写入一份"重复条目"的会话').toBe(2);

    await page.reload();
    await waitForBoot(page);
    await page.waitForTimeout(2500);

    const restored = await tabs(page);
    expect(restored, '读回来必须只剩一个标签').toHaveLength(1);
    expect(new Set(restored.map((tab) => tab.id)).size).toBe(1);
    expect(await logKinds(page), '应记下跳过的重复条目').toContain('session:restore-skip-registered');
  });
});
