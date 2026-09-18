/**
 * 坐标一致性：**"点哪一格"与"那一格画在哪"必须一致**。
 *
 * 这条链路同时决定三件事：拖动落点命中、落点高亮、互换后的黄色提醒框。
 * 用户实测反馈过"滚动之后提醒框渲染位置不对"，根因就是这里的换算把滚动量消掉了
 * （见 `src/interaction/swap-flash.ts` 文件头的"实测校正"）。
 *
 * 断言方式不猜几何：**点一个像素 → 读回 Univer 自己认定的那一格（真值）→
 * 该格的高亮矩形必须包住这个像素**。滚动前后、横向滚动、缩放后都要成立。
 */
import { expect, test, type Page } from '@playwright/test';

import { canvasBox, importFixture, waitForBoot } from './helpers';

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

const selection = (page: Page): Promise<string | null> =>
  page.evaluate(() => (window as never as { __p0: { getSelectionA1: () => string | null } }).__p0.getSelectionA1());

const rectOf = (page: Page, a1: string): Promise<Box | null> =>
  page.evaluate(
    (ref) => (window as never as { __p0: { rectOfA1: (a: string) => Box | null } }).__p0.rectOfA1(ref),
    a1,
  );

const scroll = (page: Page): Promise<{ x: number; y: number }> =>
  page.evaluate(() => (window as never as { __p0: { getScrollState: () => { x: number; y: number } } }).__p0.getScrollState());

/** 点一个像素 → 读回 Univer 认定的格子；再检查该格矩形是否包住这个像素 */
async function assertPixelBelongsToCell(page: Page, label: string, x: number, y: number): Promise<string> {
  await page.mouse.click(x, y);
  await page.waitForTimeout(180);
  const a1 = await selection(page);
  expect(a1, `${label}：点像素后应有选中格`).toBeTruthy();
  const cell = a1!.split(':')[0];
  const box = await rectOf(page, cell);
  expect(box, `${label}：应能算出 ${cell} 的矩形（滚动后为 null 就说明换算把滚动量丢了）`).not.toBeNull();
  const inside =
    x >= box!.left - 1 && x <= box!.left + box!.width + 1 && y >= box!.top - 1 && y <= box!.top + box!.height + 1;
  expect(
    inside,
    `${label}：像素(${Math.round(x)},${Math.round(y)}) 落在 ${cell}，但算出的矩形是 ${JSON.stringify(box)}`,
  ).toBe(true);
  return cell;
}

/** 滚轮滚动（滚轮本身已由 scrollbar.spec 覆盖，这里只当"制造滚动"的手段） */
async function wheel(page: Page, dx: number, dy: number, times = 4): Promise<void> {
  const box = await canvasBox(page);
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  for (let i = 0; i < times; i += 1) {
    await page.mouse.wheel(dx, dy);
    await page.waitForTimeout(80);
  }
}

async function makeScrollableSheet(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as never as { __p0: { createHandsOnSheet: () => string | null } }).__p0.createHandsOnSheet();
  });
  await page.waitForTimeout(400);
}

/** 用百分比菜单设置缩放（结构见 zoom.spec） */
async function setZoom(page: Page, percent: string): Promise<void> {
  const menu = page.locator('[role="menu"]');
  for (let attempt = 0; attempt < 2 && !(await menu.isVisible().catch(() => false)); attempt += 1) {
    await page.locator('button').filter({ hasText: /^\s*\d{1,3}%\s*$/ }).first().click();
    await page.waitForTimeout(500);
  }
  await menu.locator('[role="menuitemradio"]').filter({ hasText: percent }).first().click();
  await page.waitForTimeout(400);
}

test.describe('坐标一致性（格 ↔ 像素）', () => {
  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await makeScrollableSheet(page);
  });

  /**
   * **用户实测反馈的主场景**：真实夹具（带冻结窗格）滚动之后，拖动的起点与落点必须还是
   * 指针底下那两格。
   *
   * 之前用的是 Univer 渲染服务的命中测试，它的滚动换算在 scale=1 时把滚动量消掉了：
   * 指针在 S74、app 认成 D13（起点看着正常只是因为起点用的是**选区**而不是命中测试）。
   * 现在命中测试改成自建换算（`src/interaction/cell-geometry.ts`，实测校准）。
   */
  test('⓪ 真实夹具（冻结窗格）滚动后：拖动起点/落点仍是指针底下那两格', async ({ page }) => {
    test.setTimeout(120_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);

    const box = await canvasBox(page);
    // 先滚到中间（纵横都滚）
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    for (let i = 0; i < 5; i += 1) {
      await page.mouse.wheel(300, 300);
      await page.waitForTimeout(90);
    }
    const scrolled = await scroll(page);
    expect(scrolled.x, '测试前提：应已横向滚动').toBeGreaterThan(0);
    expect(scrolled.y, '测试前提：应已纵向滚动').toBeGreaterThan(0);

    const startX = box.x + 260;
    const startY = box.y + 220;
    const dropX = startX + 240;
    const dropY = startY + 160;

    // 真值：分别点一下，读 Univer 自己认定的格
    const truthOf = async (x: number, y: number): Promise<string> => {
      await page.mouse.click(x, y);
      await page.waitForTimeout(180);
      const a1 = await selection(page);
      expect(a1, '点像素应选中一格').toBeTruthy();
      return a1!.split(':')[0];
    };
    const truthStart = await truthOf(startX, startY);
    const truthDrop = await truthOf(dropX, dropY);
    expect(truthStart, '两个像素应指向不同的格').not.toBe(truthDrop);

    // 回到起点格，开始拖动到落点像素
    await page.mouse.click(startX, startY);
    await page.waitForTimeout(200);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    await page.waitForTimeout(200);
    await page.evaluate(() => {
      (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
    });

    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX + 30, startY + 20, { steps: 3 });
    await page.waitForTimeout(150);

    const start = await page.evaluate(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string; detail?: { a1?: string } }> } }).__p0.log
          .filter((e) => e.kind === 'app:drag-start')
          .pop()?.detail?.a1 ?? null,
    );
    expect(start, `拖动起点应是 ${truthStart}（指针底下那一格）`).toBe(truthStart);

    await page.mouse.move(dropX, dropY, { steps: 10 });
    await page.waitForTimeout(200);
    await page.mouse.up();
    await page.waitForTimeout(400);

    const swap = await page.evaluate(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string; detail?: { a?: string; b?: string } }> } }).__p0.log
          .filter((e) => e.kind === 'app:swap')
          .pop()?.detail ?? null,
    );
    expect(swap, '拖到另一格应触发互换').not.toBeNull();
    expect(swap!.a, `互换的源应是 ${truthStart}`).toBe(truthStart);
    expect(swap!.b, `互换的落点应是 ${truthDrop}（指针底下那一格）`).toBe(truthDrop);

    // 提醒框也要画在这两格上（滚动后的位置换算同样由上面那套几何负责）
    const boxes = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.swap-flash-box')).map((node) => node.getAttribute('data-a1')),
    );
    expect(boxes.sort(), '提醒框应覆盖互换的两格').toEqual([truthStart, truthDrop].sort());
  });

  test('① 未滚动 / 已滚动 / 横向滚动后：点哪一格，该格矩形就包住那个像素', async ({ page }) => {
    test.setTimeout(120_000);
    const box = await canvasBox(page);
    const px = box.x + 200;
    const py = box.y + 150;

    await assertPixelBelongsToCell(page, '未滚动', px, py);
    expect((await scroll(page)).y, '测试前提：应能滚动').toBe(0);

    await wheel(page, 0, 300);
    expect((await scroll(page)).y, '测试前提：滚轮应已滚动').toBeGreaterThan(0);
    const afterScroll = await assertPixelBelongsToCell(page, '纵向滚动后', px, py);
    const afterScroll2 = await assertPixelBelongsToCell(page, '纵向滚动后（另一点）', box.x + 260, box.y + 220);
    expect(afterScroll, '滚动后同一个像素应指向更靠下的行').not.toBe(afterScroll2);

    await wheel(page, 300, 0);
    expect((await scroll(page)).x, '测试前提：应能横向滚动').toBeGreaterThan(0);
    await assertPixelBelongsToCell(page, '横向滚动后', px, py);
  });

  test('② 缩放后同样成立（缩放的矩形也要跟着放大）', async ({ page }) => {
    test.setTimeout(120_000);
    await setZoom(page, '150%');
    const box = await canvasBox(page);
    const first = await assertPixelBelongsToCell(page, '150% 缩放', box.x + 220, box.y + 160);
    const box150 = await rectOf(page, first);
    // 缩放后单元格在屏幕上的尺寸应变大（行高 30 → 45 左右）
    expect(box150!.height, '150% 缩放下单元格应更高').toBeGreaterThan(35);

    await setZoom(page, '100%');
    const second = await assertPixelBelongsToCell(page, '回到 100%', box.x + 220, box.y + 160);
    const box100 = await rectOf(page, second);
    expect(box100!.height, '回到 100% 后恢复原尺寸').toBeLessThan(box150!.height);
  });

  test('③ 滚动之后互换：黄色提醒框必须落在互换的那两格上', async ({ page }) => {
    test.setTimeout(120_000);
    await wheel(page, 0, 300);
    expect((await scroll(page)).y, '测试前提：应已滚动').toBeGreaterThan(0);

    // 先各点一下读出**当前可见**的两格（不猜坐标，避免选到滚动区外的格子）
    const box = await canvasBox(page);
    const first = await assertPixelBelongsToCell(page, '滚动后第一格', box.x + 200, box.y + 150);
    const second = await assertPixelBelongsToCell(page, '滚动后第二格', box.x + 340, box.y + 260);
    expect(first).not.toBe(second);

    // 用真实命令互换（钩子与真实交互路径一致：会清选中并画提醒框）
    await page.evaluate(
      ({ a, b }) => {
        (window as never as { __p0: { swap: (x: string, y: string) => unknown } }).__p0.swap(a, b);
      },
      { a: first, b: second },
    );
    await page.waitForTimeout(300);

    const boxes = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.swap-flash-box')).map((node) => {
        const rect = node.getBoundingClientRect();
        return { a1: node.getAttribute('data-a1'), left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      }),
    );
    expect(boxes.length, '互换后应有两个提醒框').toBe(2);
    for (const a1 of [first, second]) {
      const expected = await rectOf(page, a1);
      expect(expected, `应能算出 ${a1} 的矩形`).not.toBeNull();
      const drawn = boxes.find((entry) => entry.a1 === a1);
      expect(drawn, `提醒框应覆盖 ${a1}`).toBeTruthy();
      expect(Math.abs(drawn!.left - expected!.left), `${a1} 的提醒框横向位置应与该格一致`).toBeLessThan(2);
      expect(Math.abs(drawn!.top - expected!.top), `${a1} 的提醒框纵向位置应与该格一致（滚动后尤其要成立）`).toBeLessThan(2);
      expect(Math.abs(drawn!.width - expected!.width)).toBeLessThan(2);
      expect(Math.abs(drawn!.height - expected!.height)).toBeLessThan(2);
    }
  });
});
