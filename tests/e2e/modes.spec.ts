/**
 * 三种交互模式（合并自原 p4-modes.spec.ts 的 4 条 → 3 条）。
 *
 * 交互约定（用户要求，也是本文件守护的不变量）：
 * - 选择模式：完全原生，允许框选、允许出现多选框；
 * - 拖拽模式：**按住即可拖动**（不需要长按），且**全程不出现多选框**，松手后内容互换、互换后不再选中任何单元格；
 * - 点击互换模式：依次点两格互换，互换后不选中、不出现多选框。
 *
 * 早先的实现把"按住 + 移动"交给 Univer 当扩选处理，实测选区会从 B3 涨成 B3:C5 并保留，
 * 用户看到的就是"按住以后变成多选了"——本文件就是那条回归的护栏。
 *
 * 已删除：原"拖拽模式：按住拖动全程不出现多选框"独立用例（与原第 1 条的拖动循环完全重复，
 * 其"按住不动也不出多选框"的检查已并入第 1 条）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import {
  canvasBox,
  clearLog,
  fixturePath,
  importFixture,
  lastDetail,
  logKinds,
  selectionA1,
  switchMode,
  valueAt,
  waitForBoot,
} from './helpers';

const FIXTURE = 'fixture-styles.xlsx';

/** 拖动期间选区必须是"源单元格本身"或"已经空掉"，绝不能是一片区域 */
function assertNoMultiSelect(selection: string | null, source: string | null, step: string): void {
  expect(selection === source || !selection, `${step} 出现了多选框：${selection}`).toBe(true);
}

/** 拖到 (x,y) 之前先把一片区域的值快照下来（落点只有松手后才知道是哪个格子） */
async function snapshotGrid(page: Page): Promise<Record<string, string | null>> {
  return page.evaluate(() => {
    const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => string | null } }).__p0;
    const colName = (index: number): string => {
      let n = index + 1;
      let name = '';
      while (n > 0) {
        const remainder = (n - 1) % 26;
        name = String.fromCharCode(65 + remainder) + name;
        n = Math.floor((n - 1) / 26);
      }
      return name;
    };
    const out: Record<string, string | null> = {};
    for (let r = 0; r < 16; r += 1) {
      for (let c = 0; c < 12; c += 1) out[`${colName(c)}${r + 1}`] = hooks.getDisplayValue(r, c);
    }
    return out;
  });
}

const norm = (v: string | null | undefined): string | null => (v === '' || v === undefined ? null : v);

test.describe('三种交互模式', () => {
  test.skip(!existsSync(fixturePath(FIXTURE)), '请先运行 npm run fixtures');

  /**
   * 选区右下角那个小方块（Univer 的**填充柄**）在本工具里必须**不出现**。
   *
   * 用户实测反馈："选中单元格的时候，选中的单元格右下角会有个角标，但是这个角标毫无用处，
   * 能否利用起来或者直接不要显示这个角标？"
   *
   * 它没用是**设计使然**：填充柄服务于"自动填充"，而自动填充会批量改写一片单元格（还可能带上格式），
   * 本产品只允许逐格改内容，命令闸门永久拒绝 `/auto-fill|refill/`。
   * 更糟的是默认的拖拽模式下捏着它一拖，会被我们自己的内容搬运接管 —— 实测
   * "已互换 A12 ⇄ A15"：看着像填充，实际换了内容。所以直接不画（见 `src/univer/fill-handle.ts`）。
   */
  test('选区右下角不再画"填充柄"小方块；角上拖动 = 普通格子拖动', async ({ page }) => {
    test.setTimeout(120_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await page.waitForTimeout(400);

    const handleState = () =>
      page.evaluate(() =>
        (
          window as never as {
            __p0: { getFillHandleState: () => { available: boolean; controls: number; enabled: boolean[]; visible: (boolean | null)[] } };
          }
        ).__p0.getFillHandleState(),
      );

    // 选中一格 → 那个小方块既不启用、也不在画
    await page.evaluate(() => (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A12'));
    await page.waitForTimeout(400);
    const state = await handleState();
    expect(state.available, '应能读到选区控件').toBe(true);
    expect(state.controls, '至少有一个选区控件').toBeGreaterThan(0);
    expect(state.enabled.every((on) => on === false), '填充柄开关必须全关').toBe(true);
    expect(state.visible.every((visible) => visible === false), '填充柄都不该在画').toBe(true);

    // 新建的选区控件（换格、多选）也必须是关的
    await page.evaluate(() => (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('C5:E9'));
    await page.waitForTimeout(400);
    const multi = await handleState();
    expect(multi.enabled.every((on) => on === false) && multi.visible.every((v) => v === false), '多选后同样不该出现填充柄').toBe(true);

    // 切到另一张工作表也一样（控件是按"工作表 × 选区"新建的，闸门要在每次装配时重新扣上）
    const sheetNames = await page.evaluate(() =>
      (window as never as { __p0: { getSheetNames: () => string[] } }).__p0.getSheetNames(),
    );
    if (sheetNames.length > 1) {
      await page.evaluate((name: string) => {
        (window as never as { __p0: { activateSheet: (n: string) => void } }).__p0.activateSheet(name);
      }, sheetNames[1]);
      await page.waitForTimeout(600);
      await page.evaluate(() => (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('B2'));
      await page.waitForTimeout(400);
      const other = await handleState();
      expect(other.visible.every((v) => v === false) && other.enabled.every((on) => on === false), `另一张表（${sheetNames[1]}）也不该有填充柄`).toBe(true);
      // 切回第一张，后面的手势断言仍在原表上进行
      await page.evaluate((name: string) => {
        (window as never as { __p0: { activateSheet: (n: string) => void } }).__p0.activateSheet(name);
      }, sheetNames[0]);
      await page.waitForTimeout(600);
    }

    // 默认拖拽模式：从"角上"起拖 = 普通内容搬运（说明这里已经没有填充语义了）
    await switchMode(page, 'drag');
    await page.evaluate(() => (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A12'));
    await page.waitForTimeout(300);
    const rect = await page.evaluate(() =>
      (window as never as { __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1('A12'),
    );
    expect(rect, '应能定位 A12').not.toBeNull();
    if (!rect) return;
    await clearLog(page);
    const before = { a12: await valueAt(page, 'A12'), a15: await valueAt(page, 'A15') };
    const hx = rect.left + rect.width - 1;
    const hy = rect.top + rect.height - 1;
    await page.mouse.move(hx, hy);
    await page.mouse.down();
    await page.mouse.move(hx, hy + 70, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(500);
    expect(await logKinds(page), '角上起拖走的是我们自己的互换').toContain('app:swap');
    expect((await lastDetail(page, 'app:swap'))?.a, '从 A12 起拖').toBe('A12');
    expect(await valueAt(page, 'A12'), '内容确实换过（和从格子中间拖一样）').toBe(before.a15);

    // 选择模式：从角上（内 3px）起拖仍是正常框选，不是死区、更不是"填充"
    await switchMode(page, 'select');
    await page.evaluate(() => (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A12'));
    await page.waitForTimeout(300);
    await page.mouse.move(rect.left + rect.width - 3, rect.top + rect.height - 3);
    await page.mouse.down();
    await page.mouse.move(rect.left + rect.width - 3, rect.top + rect.height + 45, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    expect(await selectionA1(page), '角上起拖应正常框选到 A14').toBe('A12:A14');
  });

  test('拖拽模式：按住即可拖动、全程不出现多选框、松手互换且互换后不再选中', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await switchMode(page, 'drag');
    await clearLog(page);

    const box = await canvasBox(page);
    const x = box.x + 220;
    const y = box.y + 120;

    // 先轻点一下确定这个像素点落在哪个格子，再用真实键盘写入内容——
    // 样本文件左上角有些格子是空的，用键盘写一个值比"赌它非空"稳。
    await page.mouse.click(x, y);
    await page.waitForTimeout(200);
    const source = await selectionA1(page);
    expect(source, '点击后应有活动单元格').toMatch(/^[A-Z]+\d+$/);

    const marker = '拖拽源';
    await page.keyboard.type('DRAGSRC');
    await page.keyboard.insertText(marker);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(250);
    await page.mouse.click(x, y); // Enter 会把选区下移，点回来
    await page.waitForTimeout(200);
    expect(await selectionA1(page), '应回到源单元格').toBe(source);
    const sourceValue = await valueAt(page, source!);
    expect(sourceValue, '测试前提：源单元格需有内容').toBeTruthy();

    const gridBefore = await snapshotGrid(page);

    /** 松手前装好"提醒框出现过吗"的观察器（它会渐隐，回头再数 DOM 会漏判，见下方说明） */
    await page.evaluate(() => {
      const store = window as never as { __flashSeen?: { count: number; a1: (string | null)[] } };
      store.__flashSeen = { count: 0, a1: [] };
      const record = (): void => {
        const seen = store.__flashSeen;
        if (!seen || seen.count > 0) return;
        const boxes = Array.from(document.querySelectorAll('.swap-flash-box'));
        if (boxes.length === 0) return;
        seen.count = boxes.length;
        seen.a1 = boxes.map((box) => box.getAttribute('data-a1'));
      };
      record();
      new MutationObserver(record).observe(document.body, { childList: true, subtree: true });
    });

    // 拖拽模式：**按住直接拖**，不需要长按
    await page.mouse.move(x, y);
    await page.mouse.down();

    // 按住不动的阶段（旧实现的长按窗口）也不能出现多选框
    await page.waitForTimeout(120);
    assertNoMultiSelect(await selectionA1(page), source, '按住阶段');

    // 第一步移动就应进入搬运（位移过容差即触发），而不是等长按计时
    await page.mouse.move(x + 20, y + 12, { steps: 2 });
    await page.waitForTimeout(80);
    expect(await logKinds(page), '按住拖动应立刻触发搬运').toContain('app:drag-start');
    await expect(page.locator('.dsh-drag-ghost'), '按住拖动应立刻出现拖动幽灵').toHaveCount(1);

    for (let i = 2; i <= 6; i += 1) {
      await page.mouse.move(x + i * 20, y + i * 12, { steps: 2 });
      await page.waitForTimeout(30);
      assertNoMultiSelect(await selectionA1(page), source, `第 ${i} 步拖动时`);
    }

    /**
     * 落点高亮（用户要求）：拖动过程中把"即将被交换"的那一格框出来。
     * 断言三件事：浮层存在、落在日志记录的落点上、且整层穿透不吃鼠标事件。
     */
    const highlight = await page.evaluate(() => {
      const boxes = Array.from(document.querySelectorAll('.drag-target-box'));
      const layer = document.querySelector('.drag-target-layer');
      const rect = boxes[0]?.getBoundingClientRect();
      return {
        count: boxes.length,
        a1: boxes[0]?.getAttribute('data-a1') ?? null,
        size: rect ? { w: Math.round(rect.width), h: Math.round(rect.height) } : null,
        pointerEvents: layer ? getComputedStyle(layer).pointerEvents : null,
      };
    });
    expect(highlight.count, '拖动中应出现落点高亮').toBe(1);
    expect(highlight.size?.w ?? 0, '高亮框要有尺寸').toBeGreaterThan(10);
    expect(highlight.pointerEvents, '高亮层必须穿透点击').toBe('none');
    const lastTarget = await lastDetail(page, 'drag:target');
    expect(lastTarget, '拖动中应解析出落点').not.toBeNull();

    await page.mouse.up();
    await page.waitForTimeout(400);
    await expect(page.locator('.drag-target-box'), '松手后落点高亮应立刻收掉').toHaveCount(0);

    const swap = await lastDetail(page, 'app:swap');
    expect(swap, '拖动落点应触发互换').not.toBeNull();
    if (!swap) return;
    const target = String(swap.b);
    expect(target).not.toBe(source);

    // 互换语义：两边内容对调（目标格可能原本是空的 → 效果等于把内容搬过去）
    expect(norm(await valueAt(page, target)), `目标格 ${target} 应拿到源内容`).toBe(norm(sourceValue));
    expect(norm(await valueAt(page, source!)), `源格 ${source} 应拿到目标格原值`).toBe(norm(gridBefore[target]));

    /**
     * 互换完成后不该继续选中单元格——**而且要持续为空**。
     *
     * 用户实测反馈："互换后依旧会选中最开始的那一格"。根因是收尾顺序：先在 `finish()` 里互换
     * （内部清了一次选区），紧接着 `restorePinnedSelection()` 又把**源格**钉回选中；
     * 加上 Univer 的选区写入是异步命令，偶尔会落在我们清完之后。
     * 现在：先停"钉选区"循环、清掉源标记，再 finish；并连清几帧 + 一次 160ms 兜底（用户一动指针即取消）。
     * 这里按时间采样三次，专门防"清完又被写回来"这种竞态。
     */
    for (const waitMs of [50, 300, 600]) {
      await page.waitForTimeout(waitMs === 50 ? 50 : waitMs - 50);
      expect(await selectionA1(page), `互换后 ${waitMs}ms 采样：不应还选中着（源格是 ${source}）`).toBeFalsy();
    }
    /**
     * 黄色提醒框：**用观察器记录"它出现过"**，而不是在这一刻去数 DOM。
     *
     * 为什么（并发跑测试后暴露）：提醒框只活 ~1.6 秒就渐隐，而上面三次选区采样已经花掉近 1 秒；
     * 机器忙时等我们回头数 DOM，它早消失了 —— 断言会假红（而且红得没有信息量）。
     * 观察器在松手**之前**装好，只要它出现过就记为真，既保留"真的画出来了"这个事实，
     * 又不依赖"何时去看"。
     */
    const flash = await page.evaluate(
      () => (window as never as { __flashSeen?: { count: number; a1: (string | null)[] } }).__flashSeen ?? { count: 0, a1: [] },
    );
    expect(flash.count, '互换后应出现黄色提醒框（渐隐前它确实被画出来了）').toBeGreaterThan(0);
    expect(flash.a1.filter(Boolean).length, '提醒框应标出被互换的格子').toBeGreaterThan(0);
  });

  test('选择模式：按住拖动仍是原生框选，不搬内容也不出现拖拽幽灵', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await switchMode(page, 'select');
    await clearLog(page);

    const box = await canvasBox(page);
    const x = box.x + 220;
    const y = box.y + 120;

    await page.mouse.click(x, y);
    await page.waitForTimeout(200);
    const source = await selectionA1(page);
    const before = await valueAt(page, source!);

    await page.mouse.move(x, y);
    await page.mouse.down();
    await expect(page.locator('.dsh-drag-ghost'), '选择模式不应出现拖动幽灵').toHaveCount(0);
    for (let i = 1; i <= 6; i += 1) {
      await page.mouse.move(x + i * 20, y + i * 12, { steps: 2 });
      await page.waitForTimeout(20);
    }
    // 选择模式是唯一允许出现多选框的模式：按住拖动应当照旧框选
    expect(await selectionA1(page), '选择模式按住拖动应当框选出一个区域').toContain(':');
    await page.mouse.up();
    await page.waitForTimeout(400);

    const kinds = await logKinds(page);
    expect(kinds, '选择模式不应触发搬运').not.toContain('app:drag-start');
    expect(kinds, '选择模式不应触发互换').not.toContain('app:swap');
    expect(kinds, '选择模式不应触发移动').not.toContain('app:move');
    expect(await valueAt(page, source!), '选择模式内容不应变化').toBe(before);
  });

  test('点击互换模式：依次点两格即互换，互换后不选中', async ({ page }) => {
    test.setTimeout(120_000);
    await waitForBoot(page);
    await importFixture(page, FIXTURE);
    await switchMode(page, 'click-swap');
    await clearLog(page);

    const box = await canvasBox(page);

    // 点第一格（这一步本身就把它登记为"待互换的一方"），再用真实键盘写入内容，避免赌样本里这格非空。
    // 注意：**不能再点一次同一格**——点击互换里"再点同一格"是取消选中，那样后面永远不会触发互换。
    await page.mouse.click(box.x + 220, box.y + 120);
    await page.waitForTimeout(250);
    const first = await selectionA1(page);
    expect(first, '点击后应有活动单元格').toMatch(/^[A-Z]+\d+$/);
    await page.keyboard.type('SWAPSRC');
    await page.keyboard.insertText('互换源');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(250);
    const firstValue = await valueAt(page, first!);
    expect(firstValue, '测试前提：第一格需有内容').toBeTruthy();

    // 点第二格（另一处）→ 应立即互换
    await page.mouse.click(box.x + 400, box.y + 220);
    await page.waitForTimeout(600);
    // 注意：点击互换走的是 `swap:cell-cell` 日志（拖动落点走的是 `app:swap`），两者别混
    const swap = await lastDetail(page, 'swap:cell-cell');
    expect(swap, '两次点击应触发互换').not.toBeNull();
    if (!swap) return;
    expect(String(swap.a)).toBe(first);

    const second = String(swap.b);
    expect(second).not.toBe(first);
    expect(await valueAt(page, second), '第二格应拿到第一格的内容').toBe(firstValue);
    // 互换完成后不该继续选中单元格，也不该出现多选框
    expect(await selectionA1(page), '互换完成后不应继续选中单元格').toBeFalsy();
  });
});
