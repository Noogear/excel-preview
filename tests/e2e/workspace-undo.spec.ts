/**
 * 工作区动作的撤销/重做 + 快捷键焦点无关性。
 *
 * 用户实测两条：
 *  1. "无法撤回对工作区的操作（撤回时工作区未改变，或工作区的操作未被记录）"；
 *  2. "鼠标点击工作区后，撤回重做这类快捷键失效"。
 *
 * 第一条的根因：工作区增删只改我们的 React state，Univer 的撤销栈完全不知情。
 * 现在所有工作区改动都走 `commitWorkspace`（记带 before/after 快照的历史），
 * 撤销/重做按账本时间顺序在"表格动作"与"工作区动作"之间切换。
 * 第二条的根因：以前完全依赖 Univer 的快捷键服务，而它只在表格有焦点时响应；
 * 现在窗口层接管 Ctrl/Cmd+Z、Ctrl+Y、Ctrl+Shift+Z（焦点在输入框里时让路）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { fixturePath, importFixture, lastDetail, readLog, valueAt, waitForBoot } from './helpers';

const items = (page: Page) =>
  page.evaluate(
    () =>
      (window as never as { __p0: { getWorkspaceItems: () => Array<{ id: string; a1: string }> } }).__p0.getWorkspaceItems().length,
  );

/** 把某个区域放进工作区（走测试钩子，它和生产路径共用 commitWorkspace） */
async function fillWorkspace(page: Page, a1: string): Promise<number> {
  await page.evaluate((ref: string) => {
    (window as never as { __p0: { snapshotSelectionToWorkspace: (a1?: string) => unknown } }).__p0.snapshotSelectionToWorkspace(ref);
  }, a1);
  await page.waitForTimeout(250);
  return items(page);
}

/**
 * 等工作区条目数变成期望值（**轮询**，不是睡一觉再看）。
 *
 * 为什么必须这样（测试提速后暴露的问题）：并发跑测试时机器更忙，而撤销/重做链路上有异步命令
 * （Univer 的 undo 是异步执行的），固定 `waitForTimeout(300)` 在忙时会假红。
 * 轮询不影响"到底撤没撤对"的判定，又能把等待压到刚好够 —— 测试反而更准也更快。
 */
async function expectItems(page: Page, expected: number, message: string): Promise<void> {
  await expect.poll(() => items(page), { timeout: 10_000, message }).toBe(expected);
}

/**
 * 按一次撤销/重做，并等应用**真的处理完**这一次按键。
 *
 * 为什么不能只 `press` 完睡一会（并发跑时实测偶发假红："重做第一步"拿到 0）：
 * 快捷键是同步分发的，但**这一步的动作是异步的**（Univer 的 undo/redo 是异步命令，
 * 工作区回放也依赖 React 提交）。账本每处理完一步都会记 `history:undo-*` / `history:redo-*`，
 * 所以"等日志出现新的一条"就是"这一步真的落地了"的准确信号；到了边界（没得撤）时不会有新日志，
 * 这种情况给 4 秒后放行（由调用方的断言去判断对错）。
 */
async function pressAndWait(page: Page, key: 'Control+z' | 'Control+y'): Promise<void> {
  const before = (await readLog(page)).filter((entry) => /^history:(undo|redo)/.test(entry.kind)).length;
  await page.keyboard.press(key);
  await page
    .waitForFunction(
      (count: number) =>
        (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter((entry) =>
          /^history:(undo|redo)/.test(entry.kind),
        ).length > count,
      before,
      { timeout: 4_000 },
    )
    .catch(() => undefined);
  await page.waitForTimeout(80);
}

/** 焦点在 body 上时按 Ctrl+Z（模拟"点了工作区之后"） */
async function pressUndo(page: Page): Promise<void> {
  await pressAndWait(page, 'Control+z');
}
async function pressRedo(page: Page): Promise<void> {
  await pressAndWait(page, 'Control+y');
}

test.describe('工作区动作可撤销/重做', () => {
  test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures');

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(300);
  });

  test('① 加入工作区后 Ctrl+Z 撤销、Ctrl+Y 重做', async ({ page }) => {
    test.setTimeout(120_000);
    await expectItems(page, 0, '工作区条目数应为 0');

    const added = await fillWorkspace(page, 'A1:F6');
    expect(added, '应先放进若干条目').toBeGreaterThan(0);

    await pressUndo(page);
    await expectItems(page, 0, 'Ctrl+Z 应把工作区恢复成加入之前');

    await pressRedo(page);
    await expectItems(page, added, 'Ctrl+Y 应把条目放回来');

    // 连续性：反复撤/重做都不乱
    await pressUndo(page);
    await pressUndo(page);
    await expectItems(page, 0, '工作区条目数应为 0');
    await pressRedo(page);
    await expectItems(page, added, '工作区条目数应为 added');
  });

  test('② 删除/清空工作区也能撤销回来', async ({ page }) => {
    test.setTimeout(120_000);
    const added = await fillWorkspace(page, 'A1:F6');
    expect(added).toBeGreaterThan(1);

    // 右键第一张卡片 → 删除（工作区条目菜单与表格右键共用同一个 ContextMenu 组件）
    await page.locator('[data-testid="workspace-item"]').first().click({ button: 'right' });
    await page.waitForTimeout(300);
    await page.locator('[data-testid="context-menu-ws-delete"]').click();
    await page.waitForTimeout(300);
    await expectItems(page, added - 1, '删除后少一个');

    await pressUndo(page);
    await expectItems(page, added, 'Ctrl+Z 应把删掉的条目还回来');

    // 清空（带二次确认）→ 撤销
    await page.locator('[data-testid="workspace-clear"]').click();
    await page.waitForTimeout(200);
    await page.locator('[data-testid="workspace-clear-all"]').click();
    await page.waitForTimeout(300);
    await expectItems(page, 0, '清空后应为 0');

    await pressUndo(page);
    await expectItems(page, added, 'Ctrl+Z 应把清空恢复');
  });

  test('③ 点击工作区之后快捷键依然有效（焦点不在表格上）', async ({ page }) => {
    test.setTimeout(120_000);
    const added = await fillWorkspace(page, 'A1:F6');
    expect(added).toBeGreaterThan(0);

    // 点一下工作区面板（卡片与空白处各点一次）——以前这一下之后 Ctrl+Z 就没反应了
    await page.locator('[data-testid="workspace-item"]').first().click();
    await page.waitForTimeout(150);
    await page.locator('.ws-panel-title').click();
    await page.waitForTimeout(150);

    await pressUndo(page);
    await expectItems(page, 0, '点了工作区之后 Ctrl+Z 仍应撤销');

    await pressRedo(page);
    await expectItems(page, added, 'Ctrl+Y 同样有效');
  });

  test('④ 表格动作与工作区动作按时间顺序交替撤销', async ({ page }) => {
    test.setTimeout(120_000);
    const added = await fillWorkspace(page, 'A1:F6');
    expect(added).toBeGreaterThan(0);

    // 再做一个表格动作：把 B4 与 D6 互换
    const beforeSwap = await page.evaluate(() => ({
      b4: (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('B4'),
      d6: (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('D6'),
    }));
    await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { swap: (a: string, b: string) => unknown; getDisplayValueByA1: (a1: string) => string | null };
      }).__p0;
      hooks.swap('B4', 'D6');
    });
    await page.waitForTimeout(400);
    const afterSwap = await page.evaluate(() => ({
      b4: (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('B4'),
      d6: (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('D6'),
    }));
    expect(afterSwap.b4, '互换后 B4 应拿到 D6 的值').toBe(beforeSwap.d6);

    // 第一次撤销：撤掉"表格互换"（最新），工作区不动
    await pressUndo(page);
    const afterUndo1 = await page.evaluate(() => ({
      b4: (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('B4'),
      items: (window as never as { __p0: { getWorkspaceItems: () => unknown[] } }).__p0.getWorkspaceItems().length,
    }));
    expect(afterUndo1.b4, '最新的是表格动作 → 先撤表格').toBe(beforeSwap.b4);
    expect(afterUndo1.items, '工作区这一步不该被动到').toBe(added);

    // 第二次撤销：撤掉"加入工作区"
    await pressUndo(page);
    await expectItems(page, 0, '再撤一次才是工作区那一步');

    // 两次重做按原顺序回来
    await pressRedo(page);
    await expectItems(page, added, '重做第一步：工作区条目回来');
    await pressRedo(page);
    const afterRedo2 = await page.evaluate(
      () => (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('B4'),
    );
    expect(afterRedo2, '重做第二步：表格互换回来').toBe(afterSwap.b4);
  });

  test('⑤ 工具栏撤销/重做按钮对工作区动作同样可用（不再是灰的）', async ({ page }) => {
    test.setTimeout(120_000);
    const undoButton = page.locator('[data-testid="toolbar-undo"]');
    const redoButton = page.locator('[data-testid="toolbar-redo"]');

    const added = await fillWorkspace(page, 'A1:F6');
    await expect(undoButton, '加入工作区之后"撤销"应可用').toBeEnabled();
    await undoButton.click();
    await page.waitForTimeout(300);
    await expectItems(page, 0, '按钮撤销应生效');
    await expect(redoButton, '撤销之后"重做"应可用').toBeEnabled();
    await redoButton.click();
    await page.waitForTimeout(300);
    await expectItems(page, added, '按钮重做应生效');
  });

  /**
   * 回归：一次表格动作**只能记一条**历史。
   *
   * 实测过的坑：我们自己的互换命令既被 Univer 的 `undos` 计数抓到（订阅自动补一条"编辑内容"），
   * 又被 `pushHistory` 记了一条"互换 …"，于是账本里凭空多出一个"幽灵步骤"——
   * 按一次 Ctrl+Z 只撤掉表格、索引却退了两格，用户看到的就是"工作区那一步撤不掉"。
   */
  test('⑥ 一次表格动作只记一条历史（不多出幽灵步骤）', async ({ page }) => {
    test.setTimeout(120_000);

    await page.evaluate(() => {
      (window as never as { __p0: { swap: (a: string, b: string) => unknown } }).__p0.swap('B4', 'D6');
    });
    await page.waitForTimeout(400);

    await page.click('[data-testid="toolbar-history"]');
    await expect(page.locator('[data-testid="history-panel"]')).toBeVisible();
    const rows = page.locator('[data-testid^="history-item-"]');
    await expect(rows, '一次互换 = 一条历史').toHaveCount(1);
    await expect(rows.first(), '留的是我们自己的措辞，而不是自动记账的"编辑内容"').toContainText('互换 B4 ⇄ D6');
    await page.click('[data-testid="history-close"]');

    // 一条历史 = 一次撤销就能回到互换之前
    const before = await page.evaluate(() =>
      (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('D6'),
    );
    await pressUndo(page);
    expect(
      await page.evaluate(() =>
        (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0.getDisplayValueByA1('B4'),
      ),
      '一次 Ctrl+Z 就该撤掉这次互换',
    ).toBe(before);
  });

  /**
   * 回归（用户："确保把所有对工作区、表格的操作都统一走可撤销的通道"）。
   *
   * 曾经的根因：把工作区条目拖回表格时记的是 `scope: 'workspace'` 的历史，而工作区快照只有
   * `commitWorkspace` 会建 —— 撤销时 `stepBack` 找不到快照，直接 `return false`：
   * **账本索引卡在最后一步，Ctrl+Z 从此永久失灵**（不是"这一步撤不掉"，是后面全都撤不掉）。
   * 现在这条路径走表格通道（`beginSheetAction` + `pushHistory(..., 'sheet')`）。
   */
  test('⑦ 工作区条目拖回表格：写入进账本且 Ctrl+Z 能撤销（回归：作用域记错导致撤销失灵）', async ({ page }) => {
    test.setTimeout(120_000);
    expect(await fillWorkspace(page, 'A1')).toBe(1);

    /**
     * 落点选**已用区域之外**的 D22（夹具已用区域是 A1:H21）：
     * 空目标格让"撤销后应恢复成空"成为可断言的事实，不必猜它原来是什么。
     * 像素位置从 `__p0.rectOfA1` 反查（它给的是视口坐标，正好喂给真实鼠标手势）。
     */
    const rect = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null };
      }).__p0;
      return hooks.rectOfA1('D22');
    });
    expect(rect, 'D22 应落在可见区域内（否则手势点不到）').not.toBeNull();

    const tile = await page.locator('[data-testid="workspace-item"]').first().boundingBox();
    if (!tile) throw new Error('工作区条目不可见');

    // 真实手势：从条目上按住，拖到 D22 松手（先小步越过拖动阈值，再走大位移）
    await page.mouse.move(tile.x + tile.width / 2, tile.y + tile.height / 2);
    await page.mouse.down();
    await page.mouse.move(tile.x - 40, tile.y + 10, { steps: 4 });
    await page.mouse.move(rect!.left + rect!.width / 2, rect!.top + rect!.height / 2, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    const detail = await lastDetail(page, 'app:workspace-paste');
    expect(detail, '拖到单元格上应产生一次写入').not.toBeNull();
    expect(detail?.to, '落点应与手势瞄的格子一致').toBe('D22');

    const written = (await valueAt(page, 'D22')) ?? '';
    expect(written, '写入后目标格应该有内容了').not.toBe('');

    // 撤销：目标格回到写入之前（以前这一步是"按了没反应"）
    await pressUndo(page);
    expect((await valueAt(page, 'D22')) ?? '', 'Ctrl+Z 应把这次写入撤掉（D22 本来就是空的）').toBe('');
    await expectItems(page, 1, '条目本身不该被动到（写回不移除条目）');

    // 重做：内容回来
    await pressRedo(page);
    expect((await valueAt(page, 'D22')) ?? '', 'Ctrl+Y 应把写入重做出来').toBe(written);
  });

  /**
   * 新面板（用户："加号按钮打开的面板请进行重构…要求有'加入工作区后保留表格内容'的选项"）。
   *
   * 取消勾选 = 剪切语义：内容先进工作区，再清空源内容。**顺序不能反** ——
   * 撤销时先回滚表格内容、再回滚工作区，任何一步之后内容都还在某一边，不会丢。
   */
  test('⑧ 「+」面板取消「保留表格内容」＝剪切：Ctrl+Z 先还原表格内容、再撤工作区', async ({ page }) => {
    test.setTimeout(120_000);
    await page.locator('[data-testid="workspace-add"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toBeVisible();

    // A12:A17 是夹具里 6 个连续有内容的格子（探针核过）
    await page.locator('[data-testid="import-input"]').fill('A12:A17');
    await expect(page.locator('[data-testid="import-preview-summary"]'), '预览要如实报数').toContainText('将加入 6 个单元格');
    await expect(page.locator('[data-testid="import-confirm"]'), '确认键上带格数').toContainText('加入 6 格');

    // 取消勾选「加入工作区后保留表格内容」→ 剪切
    const keep = page.locator('[data-testid="import-keep-source"]');
    await expect(keep, '默认勾选（复制语义）').toBeChecked();
    await keep.uncheck();
    await expect(page.locator('[data-testid="import-keep-hint"]')).toContainText('剪切');

    const before = await valueAt(page, 'A12');
    expect(before, 'A12 本来有内容').not.toBe('');

    await page.locator('[data-testid="import-confirm"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toHaveCount(0);

    await expectItems(page, 6, '6 格都进了工作区');
    expect(await valueAt(page, 'A12'), '剪切语义：源内容应被清空').toBe('');

    // 第一次撤销：表格内容回来（内容永远不会两头都不在）
    await pressUndo(page);
    expect(await valueAt(page, 'A12'), 'Ctrl+Z 应先还原表格里的内容').toBe(before);
    await expectItems(page, 6, '这一步不该动工作区');

    // 第二次撤销：工作区回到加入之前
    await pressUndo(page);
    await expectItems(page, 0, '再撤一次才移除工作区条目');

    // 两次重做按原顺序回来
    await pressRedo(page);
    await expectItems(page, 6, '工作区条目数应为 6');
    await pressRedo(page);
    expect(await valueAt(page, 'A12'), '重做到位：源内容再次被清空').toBe('');
  });

  test('⑨ 「+」面板默认「保留表格内容」＝复制：表格内容不动，撤销只撤工作区', async ({ page }) => {
    test.setTimeout(120_000);
    await page.locator('[data-testid="workspace-add"]').click();
    await page.locator('[data-testid="import-input"]').fill('A12:A17');
    await expect(page.locator('[data-testid="import-keep-source"]')).toBeChecked();

    const before = await valueAt(page, 'A12');
    await page.locator('[data-testid="import-confirm"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toHaveCount(0);

    await expectItems(page, 6, '工作区条目数应为 6');
    expect(await valueAt(page, 'A12'), '复制语义：表格内容必须原样保留').toBe(before);

    // 一次撤销就回到"工作区为空"，表格依旧不动
    await pressUndo(page);
    await expectItems(page, 0, '工作区条目数应为 0');
    expect(await valueAt(page, 'A12')).toBe(before);
  });

  /**
   * 回归（用户："该面板添加的数据不会被历史记录所记录而导致无法撤销"）。
   *
   * 打开文件时应用特性（条件格式/数据验证/超链接/批注/图片）会跑一串 Univer 命令，
   * 以前这些 `undos` 增长会被订阅当成"用户编辑"记进账本，于是刚打开文件就躺着几条
   * "编辑内容"幽灵条目：用户按 Ctrl+Z 撤不动任何看得见的东西，还以为撤销坏了。
   * 现在装载窗口整个不记账，并且顺手清掉新单元的撤销栈。
   */
  test('⑩ 打开文件不在账本里留幽灵条目：历史面板为空，之后的编辑照常可撤', async ({ page }) => {
    test.setTimeout(120_000);

    const pushes = (await readLog(page)).filter((entry) => entry.kind === 'history:push');
    expect(pushes.map((entry) => entry.detail?.label), '打开文件不该往账本里写任何一步').toEqual([]);
    expect(
      (await readLog(page)).some((entry) => entry.kind === 'import:undo-stack-cleared'),
      '导入后应清掉新单元的撤销栈（装载不进撤销栈）',
    ).toBe(true);

    await page.click('[data-testid="toolbar-history"]');
    await expect(page.locator('[data-testid="history-panel"]')).toBeVisible();
    await expect(page.locator('[data-testid^="history-item-"]'), '刚打开文件时历史是空的').toHaveCount(0);
    await page.click('[data-testid="history-close"]');

    // 之后的一次真实编辑必须照常入账、照常能撤
    const before = await valueAt(page, 'B9');
    await page.evaluate(() => {
      (window as never as { __p0: { setValue: (row: number, col: number, value: string) => void } }).__p0.setValue(8, 1, '幽灵检查');
    });
    await page.waitForTimeout(400);
    expect(await valueAt(page, 'B9')).toBe('幽灵检查');

    await pressUndo(page);
    expect(await valueAt(page, 'B9'), '编辑照常可撤销').toBe(before ?? '');
  });
});
