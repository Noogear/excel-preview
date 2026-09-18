/**
 * 滚动条（用户实测："滚动条无法正常使用，有明显顿挫感"）。
 *
 * 根因：非选择模式下我们会拦掉 `pointermove`（用来阻止 Univer 把拖动当成扩选、画出多选框），
 * 而 Univer 的滚动条**就画在主画布上**——事件在捕获阶段被我们 `stopPropagation`，
 * 滚动条拖拽收不到移动事件，于是"拖不动 / 一跳一跳"。
 *
 * 修法：按下时先问引擎 `Viewport.getScrollBar().pick(coord)`，命中滚动条就不按"内容拖动"处理、
 * 也不拦 pointermove（见 `src/App.tsx` 的 `isScrollbarHit`）。
 *
 * 断言方式：拖动过程中**逐点采样**滚动量——真能用的话每一小步都会跟着滚（连续、单调）；
 * 被拦掉的话滚动量在整个拖拽过程中几乎不动，只在松手时跳一次。
 *
 * 用 `createHandsOnSheet()` 造表：60 行 × 20 列、行高 30，几何确定，一定有滚动条
 * （夹具样本太小，整个表一屏就放得下，反而测不到滚动）。
 */
import { expect, test, type Page } from '@playwright/test';

import { canvasBox, importFixture, waitForBoot } from './helpers';

const scrollState = (page: Page): Promise<{ x: number; y: number }> =>
  page.evaluate(() => (window as never as { __p0: { getScrollState: () => { x: number; y: number } } }).__p0.getScrollState());

/** 滚动条交互带（页面坐标）。注意它**不在**画布最右/最下边缘：主画布比视口大，四周有表头与留白 */
const scrollbarBand = (page: Page) =>
  page.evaluate(() =>
    (
      window as never as {
        __p0: {
          getScrollbarBand: () => {
            vertical: { x: number; y: number; width: number; height: number } | null;
            horizon: { x: number; y: number; width: number; height: number } | null;
            verticalThumb: { x: number; y: number; height: number } | null;
            horizonThumb: { y: number; x: number; width: number } | null;
          } | null;
        };
      }
    ).__p0.getScrollbarBand(),
  );

/** 滚轮往下滚若干次（同时用来确认滚轮本身没被我们的手势逻辑影响） */
async function wheelDown(page: Page, box: { x: number; y: number; width: number; height: number }, times = 6) {
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  for (let i = 0; i < times; i += 1) {
    await page.mouse.wheel(0, 300);
    await page.waitForTimeout(70);
  }
}

/** 造一张有滚动空间的表（60 行 × 20 列、行高 30）；夹具样本一屏放得下，测不到滚动 */
async function makeScrollableSheet(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as never as { __p0: { createHandsOnSheet: () => string | null } }).__p0.createHandsOnSheet();
  });
  await page.waitForTimeout(400);
}

/** 画布中心 x（横向滚动条上取一点用） */
function box_center_x(box: { x: number; width: number }): number {
  return box.x + box.width * 0.5;
}

test.describe('滚动条可用性', () => {
  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
  });

  /**
   * 滚轮滚动（用户实测："滚动条无法正常使用"里最常被踩的一种）。
   *
   * 根因是我们自己的只读闸门：滚轮走的是 `SetScrollRelativeCommand`
   * （`sheet.command.set-scroll-relative`），而白名单里**漏了它** → 每滚一次被拦一次，滚轮完全不动。
   * 修法见 `src/univer/read-only-guard.ts`（另加 `focusSheetUnit` 保证 FOCUSING_SHEET 为真）。
   */
  /**
   * **用户实测反馈的主场景**：真实夹具（带冻结窗格）的滚动条。
   * 关键断言：按在滚动条上（含**按偏一两像素**）绝不能被当成内容拖动 ——
   * 而且拖完之后表格里不能出现任何互换/搬运。
   */
  test('⓪ 真实夹具：拖滚动条（含按偏一点）不会变成内容拖拽', async ({ page }) => {
    test.setTimeout(150_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    await page.waitForTimeout(200);

    const band = await scrollbarBand(page);
    expect(band?.vertical, '应能定位到竖向滚动条').not.toBeNull();
    const track = band!.vertical!;
    const thumb = band!.verticalThumb;
    const barX = thumb ? thumb.x : track.x;

    // 三种落点：正中、偏左 3px、偏右 3px（后者以前会落到画布上被当成内容拖动）
    for (const offset of [0, -3, 3]) {
      await page.evaluate(() => {
        (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
      });
      const startY = (thumb ? thumb.y + thumb.height / 2 : track.y + track.height * 0.4) + 40;
      await page.mouse.move(barX + offset, startY);
      await page.mouse.down();
      await page.mouse.move(barX + offset, startY + 60, { steps: 6 });
      await page.mouse.up();
      await page.waitForTimeout(250);

      const kinds = await page.evaluate(() =>
        (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
      );
      expect(kinds, `偏移 ${offset}px：应按滚动条处理`).toContain('scrollbar:drag-start');
      expect(kinds, `偏移 ${offset}px：不该变成内容拖动`).not.toContain('app:drag-start');
      expect(
        kinds.filter((kind) => kind === 'app:swap' || kind === 'app:move'),
        `偏移 ${offset}px：不该互换/搬运内容`,
      ).toEqual([]);
    }

    // 而且真的滚动了（不是"识别成滚动条但不动"）
    const moved = await scrollState(page);
    expect(moved.y, '拖滚动条应让视图滚动').toBeGreaterThan(0);

    // 横向滚动条同理（按偏 3px）
    const hBand = await scrollbarBand(page);
    if (hBand?.horizon) {
      await page.evaluate(() => {
        (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
      });
      const hY = hBand.horizon.y;
      await page.mouse.move(box_center_x(await canvasBox(page)) , hY + 3);
      await page.mouse.down();
      await page.mouse.move(box_center_x(await canvasBox(page)) + 60, hY + 3, { steps: 6 });
      await page.mouse.up();
      await page.waitForTimeout(250);
      const kinds = await page.evaluate(() =>
        (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
      );
      expect(kinds, '横向滚动条按偏一点也不该变成内容拖动').not.toContain('app:drag-start');
    }
  });

  test('① 滚轮能滚动（横向/纵向都行），且不误触发内容搬运', async ({ page }) => {
    test.setTimeout(120_000);
    await makeScrollableSheet(page);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    const box = await canvasBox(page);
    await page.evaluate(() => {
      (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
    });

    await wheelDown(page, box);
    const afterWheelY = await scrollState(page);
    expect(afterWheelY.y, '纵向滚轮应能滚动').toBeGreaterThan(0);

    // 横向滚轮
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    for (let i = 0; i < 4; i += 1) {
      await page.mouse.wheel(300, 0);
      await page.waitForTimeout(70);
    }
    expect((await scrollState(page)).x, '横向滚轮应能滚动').toBeGreaterThan(0);

    const kinds = await page.evaluate(() =>
      (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
    );
    expect(kinds, '滚轮滚动不该被当成拖动').not.toContain('app:drag-start');
    expect(kinds.filter((kind) => kind === 'app:swap' || kind === 'app:move')).toEqual([]);

    // 回归护栏：滚动命令不允许被闸门拦下
    const blocked = await page.evaluate(() =>
      (window as never as { __p0: { getBlockedCommands: () => Array<{ id: string }> } }).__p0
        .getBlockedCommands()
        .map((entry) => entry.id)
        .filter((id) => /scroll/i.test(id)),
    );
    expect(blocked, '滚动命令不该出现在被拦清单里').toEqual([]);
  });

  test('② 编辑过单元格之后滚轮仍然可用（焦点不会被编辑器永久抢走）', async ({ page }) => {
    test.setTimeout(120_000);
    await makeScrollableSheet(page);
    const box = await canvasBox(page);

    // 双击进入编辑 → 输入 → 回车提交（这个过程会把焦点交给 docs 编辑器）
    await page.mouse.dblclick(box.x + 150, box.y + 150);
    await page.waitForTimeout(300);
    await page.keyboard.type('焦点测试');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(300);

    // 在表格里点一下拿回焦点，然后滚轮
    await page.mouse.click(box.x + 150, box.y + 150);
    await page.waitForTimeout(200);
    await wheelDown(page, box, 4);
    expect((await scrollState(page)).y, '编辑后滚轮仍应能滚动').toBeGreaterThan(0);
  });

  /**
   * **普通滚动条的手感**（用户要求："按住然后可以左右或者上下拖动"）。
   *
   * 断言的就是"跟手"这件事：
   *  ① 按住滑块**不跳**（滚动量不变）；
   *  ② 往下拖 → 滚动量逐步变大；再原路拖回 → 回到起点（1:1 跟手，不是按比例飘）；
   *  ③ 拖动期间有抓取高亮（可视反馈），松手即收；
   *  ④ 拖到两端会夹住，不会滚过头。
   */
  test('③ 按住滑块拖动：不跳、跟手、松手回到原处（真实夹具）', async ({ page }) => {
    test.setTimeout(150_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);

    const band = await scrollbarBand(page);
    expect(band?.verticalThumb, '应能定位到竖向滑块').not.toBeNull();
    const thumb = band!.verticalThumb!;
    const barX = thumb.x;
    const startY = thumb.y + thumb.height / 2;

    // ① 按住滑块：不该跳
    const before = await scrollState(page);
    await page.mouse.move(barX, startY);
    await page.mouse.down();
    await page.waitForTimeout(150);
    const afterDown = await scrollState(page);
    expect(afterDown.y, '按住滑块不该让视图跳动').toBeCloseTo(before.y, 0);

    // ② 往下拖：逐步变大，且滑块跟手
    const samples: number[] = [];
    const step = 18;
    for (let i = 1; i <= 8; i += 1) {
      await page.mouse.move(barX, startY + i * step, { steps: 2 });
      await page.waitForTimeout(50);
      samples.push((await scrollState(page)).y);
    }
    const distinct = samples.filter((value, index) => index === 0 || value !== samples[index - 1]).length;
    expect(distinct, `往下拖应逐步滚动（采样=${JSON.stringify(samples)}）`).toBeGreaterThanOrEqual(6);
    for (let i = 1; i < samples.length; i += 1) {
      expect(samples[i], '应单调向下').toBeGreaterThanOrEqual(samples[i - 1]);
    }
    const movedDown = samples[samples.length - 1];
    expect(movedDown, '往下拖应确实滚下去了').toBeGreaterThan(before.y);

    // ③ 拖动期间应有抓取高亮
    expect(await page.locator('.scrollbar-grab').count(), '拖动期间应有抓取高亮').toBe(1);

    // ②b 原路拖回：应回到起点附近（跟手 = 同样的位移回到同样的滚动量）
    for (let i = 7; i >= 0; i -= 1) {
      await page.mouse.move(barX, startY + i * step, { steps: 2 });
      await page.waitForTimeout(40);
    }
    await page.mouse.up();
    await page.waitForTimeout(200);
    const back = await scrollState(page);
    expect(Math.abs(back.y - before.y), `原路拖回应回到起点（${before.y} → ${back.y}）`).toBeLessThan(12);
    expect(await page.locator('.scrollbar-grab').count(), '松手后高亮应立刻收掉').toBe(0);

    // ④ 拖到最上/最下会夹住：不会滚过头，也不会变成负数
    const thumbNow = (await scrollbarBand(page))!.verticalThumb!;
    await page.mouse.move(thumbNow.x, thumbNow.y + thumbNow.height / 2);
    await page.mouse.down();
    await page.mouse.move(thumbNow.x, -500, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(200);
    const atTop = await scrollState(page);
    expect(atTop.y, '拖出顶部应夹在 0').toBe(0);

    const thumbTop = (await scrollbarBand(page))!.verticalThumb!;
    await page.mouse.move(thumbTop.x, thumbTop.y + thumbTop.height / 2);
    await page.mouse.down();
    await page.mouse.move(thumbTop.x, 5000, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(200);
    const atBottom = await scrollState(page);
    expect(atBottom.y, '拖出底部应停在滚动上限（不会无限滚）').toBeGreaterThan(atTop.y);

    // 全程不该被当成内容拖动
    const kinds = await page.evaluate(() =>
      (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
    );
    expect(kinds, '拖滚动条不该被当成内容拖动').not.toContain('app:drag-start');
    expect(kinds.filter((kind) => kind === 'app:swap' || kind === 'app:move')).toEqual([]);
  });

  test('③b 横向滚动条同样"按住拖动"（左右跟手）', async ({ page }) => {
    test.setTimeout(150_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);

    const band = await scrollbarBand(page);
    expect(band?.horizonThumb, '应能定位到横向滑块').not.toBeNull();
    const thumb = band!.horizonThumb!;
    const barY = thumb.y;
    const startX = thumb.x + thumb.width / 2;

    const before = await scrollState(page);
    await page.mouse.move(startX, barY);
    await page.mouse.down();
    await page.waitForTimeout(150);
    expect((await scrollState(page)).x, '按住横向滑块不该跳动').toBeCloseTo(before.x, 0);

    const samples: number[] = [];
    for (let i = 1; i <= 8; i += 1) {
      await page.mouse.move(startX + i * 20, barY, { steps: 2 });
      await page.waitForTimeout(50);
      samples.push((await scrollState(page)).x);
    }
    await page.mouse.up();
    await page.waitForTimeout(200);
    const distinct = samples.filter((value, index) => index === 0 || value !== samples[index - 1]).length;
    expect(distinct, `往右拖应逐步滚动（采样=${JSON.stringify(samples)}）`).toBeGreaterThanOrEqual(6);
    expect(samples[samples.length - 1], '往右拖应确实滚过去了').toBeGreaterThan(before.x);
  });

  /**
   * 拖动竖向滚动条时**不能**变成内容拖动。
   * 断言：拖滚动条不产生 `app:drag-start` / 互换 / 搬运；横向同理。
   */
  /**
   * **"能抓住"是滚动条可用的前提**（用户反馈"完全拖不动"）。
   *
   * 真实夹具的几何（实测）：主画布 1110px 宽、主视口只有 933px → 竖向滚动条画在视口右缘
   * **只有 5~6px 宽**，它右边还有 170 多像素死区。早先的实现要求"精确压住滚动条"，
   * 偏一两像素就毫无反应。现在按**条带**判定：视口右缘往里 14px + 视口外到画布边缘的整段死区。
   *
   * 用例：在条带里的若干位置按下并拖动都应滚动、都不能变成内容拖动；网格内部则仍应正常拖单元格。
   */
  test('⑥ 滚动条"抓得住"：条带内任意位置按下都能拖，网格内部仍拖单元格', async ({ page }) => {
    test.setTimeout(150_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);

    const box = await canvasBox(page);
    const clearLog = () =>
      page.evaluate(() => {
        (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
      });
    const kinds = () =>
      page.evaluate(() => (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((e) => e.kind));

    // 条带内：贴着右边缘、往里一点、以及视口外的死区
    for (const offset of [4, 12, 60, 120]) {
      await clearLog();
      const before = await scrollState(page);
      const x = box.x + box.width - offset;
      const y = box.y + 200;
      await page.mouse.move(x, y);
      await page.mouse.down();
      await page.mouse.move(x, y + 60, { steps: 5 });
      await page.mouse.up();
      await page.waitForTimeout(200);
      const after = await scrollState(page);
      const seen = await kinds();
      expect(seen, `距右边缘 ${offset}px：应被当成滚动条`).toContain('scrollbar:drag-start');
      expect(seen, `距右边缘 ${offset}px：不该变成内容拖动`).not.toContain('app:drag-start');
      expect(Math.abs(after.y - before.y), `距右边缘 ${offset}px：拖动应真的滚动了`).toBeGreaterThan(10);
    }

    // 网格内部：仍然是拖单元格（不能误判成滚动条）
    await clearLog();
    await page.mouse.move(box.x + 200, box.y + 200);
    await page.mouse.down();
    await page.mouse.move(box.x + 260, box.y + 240, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(250);
    const gridSeen = await kinds();
    expect(gridSeen, '网格内部应是内容拖动').toContain('app:drag-start');
    expect(gridSeen, '网格内部不该被当成滚动条').not.toContain('scrollbar:drag-start');
  });

  test('⑥b 悬停在滚动条条带上会显示"可抓"提示（不按下也给反馈）', async ({ page }) => {
    test.setTimeout(120_000);
    await importFixture(page, 'fixture-styles.xlsx');
    await page.waitForTimeout(400);

    const box = await canvasBox(page);
    // 悬停在右侧条带
    await page.mouse.move(box.x + box.width - 8, box.y + 200);
    await page.waitForTimeout(250);
    expect(await page.locator('.scrollbar-grab.is-hover').count(), '条带上悬停应有淡提示').toBe(1);

    // 移到网格内部 → 提示消失
    await page.mouse.move(box.x + 200, box.y + 200);
    await page.waitForTimeout(250);
    expect(await page.locator('.scrollbar-grab').count(), '离开条带应收起提示').toBe(0);

    // 横向条带同理
    await page.mouse.move(box.x + 300, box.y + box.height - 8);
    await page.waitForTimeout(250);
    expect(await page.locator('.scrollbar-grab.is-hover').count(), '底部条带悬停应有提示').toBe(1);
  });

  test('④ 拖滚动条不会误触发内容搬运', async ({ page }) => {
    test.setTimeout(120_000);
    await makeScrollableSheet(page);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    await page.evaluate(() => {
      (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
    });

    const band = await scrollbarBand(page);
    const thumb = band!.verticalThumb ?? band!.vertical!;
    const barX = 'x' in thumb ? thumb.x : 0;
    const thumbTop = 'y' in thumb ? thumb.y : 0;
    const thumbHeight = 'height' in thumb ? thumb.height : 20;
    const startY = thumbTop + thumbHeight / 2;
    await page.mouse.move(barX, startY);
    await page.mouse.down();
    await page.mouse.move(barX, startY + 120, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const kinds = await page.evaluate(() =>
      (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
    );
    expect(kinds, '拖滚动条不该被当成内容拖动').not.toContain('app:drag-start');
    expect(kinds.filter((kind) => kind === 'app:swap' || kind === 'app:move'), '拖滚动条不该搬运内容').toEqual([]);
    expect(kinds, '应记录"按在滚动条上"以便诊断').toContain('scrollbar:drag-start');
  });

  test('⑤ 横向滚动条：按下被识别为滚动条，点轨道能跳转', async ({ page }) => {
    test.setTimeout(120_000);
    await makeScrollableSheet(page);
    const box = await canvasBox(page);
    // 先横向滚一点，让横向滑块离开最左端
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    for (let i = 0; i < 5; i += 1) {
      await page.mouse.wheel(300, 0);
      await page.waitForTimeout(70);
    }
    const before = await scrollState(page);
    expect(before.x, '测试前提：应能横向滚动').toBeGreaterThan(0);

    const band = await scrollbarBand(page);
    expect(band?.horizon, '应能定位到横向滚动条').not.toBeNull();
    const track = band!.horizon!;
    const thumb = band!.horizonThumb;
    const barY = track.y;
    const pressX = thumb ? thumb.x + thumb.width / 2 : track.x + track.width / 2;

    await page.evaluate(() => {
      (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
    });
    await page.mouse.move(pressX, barY);
    await page.mouse.down();
    await page.mouse.move(pressX - 60, barY, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(250);

    const kinds = await page.evaluate(() =>
      (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.map((entry) => entry.kind),
    );
    expect(kinds, '横向滚动条按下也应被识别为滚动条').toContain('scrollbar:drag-start');
    expect(kinds, '横向滚动条上拖动不该被当成内容拖动').not.toContain('app:drag-start');
    expect(await scrollState(page)).toBeTruthy();
  });
});
