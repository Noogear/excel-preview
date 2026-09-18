/**
 * 提示的统一性（用户实测反馈）：
 *  - "拖动到工作区时提示'此处不能放置'，但其实能放" → 提示要按**落点语义**走；
 *  - "提示有多个组件、风格与位置不统一、没在最上层会被遮挡" → 所有瞬时提示统一进
 *    同一个 `prompt-layer`（同一位置/同一套样式），层级要压过 Univer 的弹层（上游 popup 到 1070）。
 */
import { expect, test, type Page } from '@playwright/test';

import { canvasBox, importFixture, waitForBoot } from './helpers';

/** 提示层的层级与位置（要能压住 Univer 的弹层） */
const promptLayerStyle = (page: Page) =>
  page.evaluate(() => {
    const layer = document.querySelector('[data-testid="prompt-layer"]');
    if (!layer) return null;
    const style = getComputedStyle(layer);
    const rect = layer.getBoundingClientRect();
    return {
      zIndex: Number.parseInt(style.zIndex, 10),
      position: style.position,
      pointerEvents: style.pointerEvents,
      rect: { x: Math.round(rect.x), y: Math.round(rect.y), w: Math.round(rect.width), h: Math.round(rect.height) },
    };
  });

/** Univer 自带弹层的最高层级（用来对照：我们的提示必须更高） */
const univerMaxZ = (page: Page) =>
  page.evaluate(() => {
    const nodes = Array.from(document.querySelectorAll('#univer-container *'));
    let max = 0;
    nodes.forEach((node) => {
      const value = Number.parseInt(getComputedStyle(node).zIndex, 10);
      if (Number.isFinite(value) && value > max) max = value;
    });
    return max;
  });

const dragHint = (page: Page) =>
  page.evaluate(() => {
    const hint = document.querySelector('[data-testid="drag-hint"]');
    return hint
      ? { text: (hint.textContent ?? '').trim(), hint: hint.getAttribute('data-hint'), cls: hint.className }
      : null;
  });

test.describe('提示统一性', () => {
  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');
  });

  test('① 拖到工作区时提示"可放置"，而不是"此处不能放置"', async ({ page }) => {
    test.setTimeout(120_000);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    await page.waitForTimeout(200);
    const box = await canvasBox(page);

    // 先确认这一格有内容（拖空格的语义不同）：点一下读出实际落点，再往那一格写值
    await page.mouse.click(box.x + 200, box.y + 150);
    await page.waitForTimeout(200);
    const cell = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getSelectionA1: () => string | null;
          getActiveSheet: () => { getRange: (a1: string) => { getRange: () => { startRow: number; startColumn: number } } };
          setValues: (row: number, col: number, values: (string | number)[][]) => void;
        };
      }).__p0;
      const a1 = hooks.getSelectionA1();
      if (!a1) return null;
      const range = hooks.getActiveSheet().getRange(a1.split(':')[0]).getRange();
      hooks.setValues(range.startRow, range.startColumn, [['拖到工作区']]);
      return a1;
    });
    expect(cell, '应能点到一个单元格并写入内容').toBeTruthy();

    // 按住 B5 开始拖动：**等 app:drag-start 真的出现**再继续（避免"没拖动就断言提示"的假失败）
    const startedDrag = async (): Promise<boolean> => {
      const before = await page.evaluate(
        () => (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter((e) => e.kind === 'app:drag-start').length,
      );
      await page.mouse.move(box.x + 200, box.y + 150);
      await page.mouse.down();
      await page.mouse.move(box.x + 240, box.y + 170, { steps: 4 });
      await page.waitForTimeout(200);
      const after = await page.evaluate(
        () => (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter((e) => e.kind === 'app:drag-start').length,
      );
      return after > before;
    };
    let started = await startedDrag();
    if (!started) {
      await page.mouse.up();
      await page.waitForTimeout(200);
      started = await startedDrag();
    }
    expect(started, '应进入拖动状态（app:drag-start）').toBe(true);
    await expect(page.locator('[data-testid="drag-hint"]'), '拖动中应出现提示').toBeVisible({ timeout: 5_000 });

    // 移到工作区上方
    const sidebar = await page.locator('.ws-panel').first().boundingBox();
    if (!sidebar) throw new Error('找不到工作区面板');
    await page.mouse.move(sidebar.x + sidebar.width / 2, sidebar.y + sidebar.height / 2, { steps: 12 });
    await page.waitForTimeout(300);

    const overWorkspace = await dragHint(page);
    expect(overWorkspace, '工作区上方应有提示').not.toBeNull();
    expect(overWorkspace!.hint, '应识别为"落到工作区"').toBe('workspace');
    expect(overWorkspace!.text, '措辞要是"可放置"').toContain('暂存到工作区');
    expect(overWorkspace!.text, '不该再出现"不能放置"').not.toContain('不能放置');
    // 工作区上方没有网格落点 → 不应画单元格高亮
    expect(await page.locator('.drag-target-box').count(), '工作区上方不该有单元格高亮').toBe(0);

    // 松手真的进了工作区（提示与实际行为一致）
    const before = await page.locator('[data-testid="workspace-item"]').count();
    await page.mouse.up();
    await page.waitForTimeout(600);
    expect(await page.locator('[data-testid="workspace-item"]').count(), '松手应真的放进工作区').toBeGreaterThan(before);
  });

  test('①b 拖动全程都不会出现"不能放置"这种否定式提示', async ({ page }) => {
    test.setTimeout(120_000);
    await page.locator('[data-testid="toolbar-mode-drag"]').click();
    await page.waitForTimeout(200);
    const box = await canvasBox(page);

    // 找一个有内容的格子并选中
    await page.mouse.click(box.x + 200, box.y + 150);
    await page.waitForTimeout(200);
    await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getSelectionA1: () => string | null;
          getActiveSheet: () => { getRange: (a1: string) => { getRange: () => { startRow: number; startColumn: number } } };
          setValues: (row: number, col: number, values: (string | number)[][]) => void;
        };
      }).__p0;
      const a1 = hooks.getSelectionA1()?.split(':')[0];
      if (!a1) return;
      const rect = hooks.getActiveSheet().getRange(a1).getRange();
      hooks.setValues(rect.startRow, rect.startColumn, [['全程提示']]);
    });

    await page.mouse.move(box.x + 200, box.y + 150);
    await page.mouse.down();
    const seen: string[] = [];
    const sample = async (): Promise<void> => {
      const text = await page.evaluate(
        () => (document.querySelector('[data-testid="drag-hint"]')?.textContent ?? '').trim(),
      );
      if (text) seen.push(text);
    };
    // 依次扫过：网格 → 网格右下（仍有效）→ 工作区 → 顶部工具栏 → 底部状态栏
    const path: Array<[number, number]> = [
      [box.x + 240, box.y + 180],
      [box.x + box.width - 60, box.y + box.height - 80],
      [box.x + box.width + 120, box.y + 260],
      [box.x + 320, box.y - 30],
      [box.x + 320, box.y + box.height + 20],
    ];
    for (const [x, y] of path) {
      await page.mouse.move(x, y, { steps: 6 });
      await page.waitForTimeout(150);
      await sample();
    }
    await page.mouse.up();
    await page.waitForTimeout(200);

    expect(seen.length, '全程应采到提示文案').toBeGreaterThan(0);
    seen.forEach((text) => {
      expect(text, `不该出现否定式提示："${text}"`).not.toContain('不能放置');
    });
    // 至少出现过一次正向引导（无效落点时的措辞）
    expect(seen.some((text) => text.includes('拖到单元格上互换') || text.includes('暂存到工作区')), `提示应给出可行做法：${JSON.stringify(seen)}`).toBe(true);
  });

  test('② 所有提示都在同一个层里，位置统一、层级压过 Univer 弹层', async ({ page }) => {
    test.setTimeout(120_000);
    const layer = await promptLayerStyle(page);
    expect(layer, '应有统一的提示层').not.toBeNull();
    expect(layer!.position).toBe('fixed');
    expect(layer!.pointerEvents, '提示层不能吃鼠标事件').toBe('none');

    const univerZ = await univerMaxZ(page);
    expect(layer!.zIndex, `提示层级(${layer!.zIndex}) 必须高于 Univer 弹层(${univerZ})`).toBeGreaterThan(univerZ);

    // 触发一条轻提示：确认它也渲染在同一个层里，且用同一套 prompt 样式
    await page.evaluate(() => {
      (window as never as { __p0: { toast: (text: string, kind?: string) => void } }).__p0.toast('统一提示层自检');
    });
    await page.waitForTimeout(300);
    const inside = await page.evaluate(() => {
      const layer = document.querySelector('[data-testid="prompt-layer"]');
      const toast = layer?.querySelector('[data-testid="toast"]') ?? null;
      return {
        inLayer: Boolean(toast),
        text: (toast?.textContent ?? '').trim(),
        cls: toast?.className ?? '',
      };
    });
    expect(inside.inLayer, '轻提示必须渲染在统一提示层里').toBe(true);
    expect(inside.text).toContain('统一提示层自检');
    expect(inside.cls, '轻提示要用统一的 prompt 样式').toContain('prompt');
  });
});
