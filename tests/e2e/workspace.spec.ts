/**
 * 工作区与内容搬运（P4：工作区 = 一个个**独立单元格**）。
 *
 * 新数据模型（与 `src/workspace/snapshot.ts` 的 `extractCellItems` 一一对应）：
 *   - **任何入口**（拖拽入侧边栏 / 右键「复制·剪切到工作区」/「+」导入 / 一键全搬）都会把区域
 *     **拆成 1×1 条目**：按**行优先**顺序，新条目排在列表**最前**；
 *   - 条目带来源单元格的像素尺寸（`cellSize`，展示时夹在 28–140 × 16–64）；
 *   - **底层默认跳过空单元格**：没值也没公式的格子不产生条目（日志 `workspace:skip-empty`）；
 *   - 一次操作上限 `MAX_WORKSPACE_ITEMS_PER_ACTION = 500`（超出记 `workspace:truncated`）。
 *
 * 不变量：**搬运的永远只是文字内容，格式留在原单元格**。
 *   ① 两个入口（真实拖动 / 右键复制）都拆成单格条目、空状态消失、行优先顺序
 *   ② 条目写回单元格 → 只改内容，目标格格式一字不变
 *   ③ 工作区**跨表共享**：一张表里的条目，切到另一张表仍能写回（复制语义，源表不变）
 *   ④ 两区域互换 → 值与格式各自留在原格（确定性路径，不依赖鼠标落点）
 *   ⑤ 「+」面板：非法输入报错、合法区域按非空格子数拆条目、不生成空条目；
 *      实时预览的数字（将加入 N / 跳过 M）必须与真正加入的条目数一致
 *   ⑥ 测试钩子新契约：返回 1×1 条目数组，条目数 = 非空格子数、skippedEmpty = 空格子数
 *   ⑦ **一键全搬**（整个已用区域）：点一下只"填范围 + 给预览"，确认后条目数 = 已用区域非空格子数、跳过全部空格、无空条目
 *   ⑧ 单格条目恒 1 列；「格宽」+ 面板宽度自动决定每行几格；分隔条拖动"变宽且右边界不动"
 *   ⑨ 拖放语义开关 + 条目右键复制/粘贴/删除
 *   ⑩ 清空二次确认（全部 / 只清显示项 / 取消）+ 搜索与来源筛选
 *   ⑪ 点击互换后取消选中并给出提醒框，点别处提醒消失
 *
 * 断言口径说明：所有"应该有几个条目"的数字都**在页面里现算**（统计源区域非空格子数），
 * 不写死数字——既证明"条目数 = 非空格子数"，也不会因为样本微调而假红；
 * 只有夹具真值（已用区域 A1:H21、A1:F2 里 7 个非空、A1:F6 里 19 个非空、A3 空 / A4='medium'）
 * 写死，注释里都标了来源。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';

import { activateSheet, canvasBox, fixturePath, importFixture, switchMode, valueAt, waitForBoot } from './helpers';

/**
 * 夹具真值（`fixtures/fixture-styles.xlsx` 的「样式」表，由 `tools/make-fixtures.mjs` 生成，
 * 已用**真实探针**核对过）：
 * - 已用区域 **A1:H21** = 168 格，其中 **35 格**有内容（其余 133 格是空）——
 *   这是"跳过空单元格"最直接的证据；
 * - 有内容的格子：`A1`、`A2:F2`（表头 6 格）、`D3:F3`、`D4:F4`、`D5:F5`、`D6:F6`、
 *   `F7`–`F11`、`A12:A17`、`A20`、`A21` —— 共 35 格；
 * - 所以 `A1:F6` = 36 格里只有 **19 格**有内容（第 1 行只有 A1，其余 5 格是空）；
 *   `A1:F2` = 12 格里只有 **7 格**有内容（同样因为 A1:F1 是合并标题）。
 * 需要"整片区域都非空"时必须挑 **A1:H25 之外**的场景，别拿第 1 行当满格用。
 */
const STYLES_USED_RANGE = { rows: 21, cols: 8 };
/** A1:F6 里真正有内容的 19 个格子（行优先，探针读数） */
const A1_F6_NON_EMPTY = [
  'A1', 'A2', 'B2', 'C2', 'D2', 'E2', 'F2',
  'D3', 'E3', 'F3', 'D4', 'E4', 'F4', 'D5', 'E5', 'F5', 'D6', 'E6', 'F6',
];

/** 读工作区条目（走 __p0 契约，返回每个条目的来源 A1 与行列数） */
async function workspaceItems(
  page: Page,
): Promise<Array<{ id: string; label: string; rows: number; cols: number; a1: string; sheetName: string }>> {
  return page.evaluate(() =>
    (
      window as never as {
        __p0: {
          getWorkspaceItems: () => Array<{
            id: string;
            label: string;
            rows: number;
            cols: number;
            a1: string;
            sheetName: string;
          }>;
        };
      }
    ).__p0.getWorkspaceItems(),
  );
}

/**
 * 调 `__p0.snapshotSelectionToWorkspace(a1)`（新契约：**返回拆出来的条目数组**）。
 *
 * 返回形状与 src 逐字对齐：
 *   `{ items: RangeSnapshot[]; skippedEmpty: number; truncated: number; ids: string[]; error?: string }`
 * 只取断言需要的字段（跨 evaluate 边界传对象，保持可序列化）。
 */
async function extractViaHook(
  page: Page,
  a1: string,
): Promise<{
  count: number;
  skippedEmpty: number;
  truncated: number;
  ids: string[];
  error?: string;
  shapes: number[][];
  a1s: string[];
  labels: string[];
}> {
  return page.evaluate((range: string) => {
    const hooks = (window as never as {
      __p0: {
        snapshotSelectionToWorkspace: (a1?: string) => {
          items: Array<{ id: string; rows: number; cols: number; source: { a1: string }; label: string }>;
          skippedEmpty: number;
          truncated: number;
          ids: string[];
          error?: string;
        };
      };
    }).__p0;
    const result = hooks.snapshotSelectionToWorkspace(range);
    return {
      count: result.items.length,
      skippedEmpty: result.skippedEmpty,
      truncated: result.truncated,
      ids: result.ids,
      error: result.error,
      shapes: result.items.map((item) => [item.rows, item.cols]),
      a1s: result.items.map((item) => item.source.a1),
      labels: result.items.map((item) => item.label),
    };
  }, a1);
}

/**
 * 现算"某个矩形区域里有几个非空格子"。
 *
 * 与 `extractCellItems` 的口径一致：**有值或有公式**才算非空（显示文本非空即视为有值）。
 * 用它将"条目数 = 非空格子数"写成**不写死数字**的断言。
 */
async function countNonEmptyCells(page: Page, rows: number, cols: number): Promise<number> {
  return page.evaluate(
    ({ rows: r, cols: c }) => {
      const hooks = (window as never as {
        __p0: {
          getDisplayValue: (row: number, col: number) => string | number | boolean | null;
          getActiveSheet: () => { getRange: (row: number, col: number) => { getFormula: () => string | null } } | null;
        };
      }).__p0;
      const sheet = hooks.getActiveSheet();
      if (!sheet) return 0;
      let count = 0;
      for (let row = 0; row < r; row += 1) {
        for (let col = 0; col < c; col += 1) {
          const display = hooks.getDisplayValue(row, col);
          const text = display === null || display === undefined ? '' : String(display);
          const formula = sheet.getRange(row, col).getFormula();
          if (text !== '' || (typeof formula === 'string' && formula !== '')) count += 1;
        }
      }
      return count;
    },
    { rows, cols },
  );
}

/** 读当前表「已用区域」的行列数（A1:H21 → { rows: 21, cols: 8 }） */
async function usedRangeSize(page: Page): Promise<{ rows: number; cols: number; a1: string }> {
  return page.evaluate(() => {
    const sheet = (window as never as {
      __p0: {
        getActiveSheet: () => {
          getDataRange: () => { getRange: () => { startRow: number; startColumn: number; endRow: number; endColumn: number } };
        } | null;
      };
    }).__p0.getActiveSheet();
    const rect = sheet?.getDataRange().getRange();
    if (!rect) return { rows: 0, cols: 0, a1: 'A1' };
    const letter = (n: number): string => {
      let out = '';
      let value = n;
      do {
        out = String.fromCharCode(65 + (value % 26)) + out;
        value = Math.floor(value / 26) - 1;
      } while (value >= 0);
      return out;
    };
    const start = `${letter(rect.startColumn)}${rect.startRow + 1}`;
    const end = `${letter(rect.endColumn)}${rect.endRow + 1}`;
    return {
      rows: rect.endRow - rect.startRow + 1,
      cols: rect.endColumn - rect.startColumn + 1,
      a1: start === end ? start : `${start}:${end}`,
    };
  });
}

/** 在画布上某个坐标右键，等到自家右键菜单出现（右键会先把选区设成指针下那一格） */
async function openMenuOnCanvasAt(page: Page, x: number, y: number): Promise<void> {
  await page.mouse.click(x, y, { button: 'right' });
  await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
}

/**
 * 在画布**左上角 200×200** 内右键（避开冻结列的分区换算）。
 *
 * Univer 把冻结列单独渲染在一个分区里：canvas 的 x 坐标不能直接当"滚动内容偏移"用
 * （实测点 canvas+70px 会落到 B 列）。所以这里**不写死目标格**，而是从 `menu:open`
 * 日志里读回"这次右键实际落在哪个 A1"，再据此断言拆分结果——既不猜坐标，也不写死格子。
 * 调用前用 `selectRange` 把选区铺满一整片区域，右键的落点自然落在区域内部。
 */
async function rightClickInsideSelection(page: Page): Promise<string> {
  const box = await canvasBox(page);
  // 探针实测：canvas 左上角 (+70, +120) 稳定落在内容网格里（冻结列/行会让像素偏移 ≠ 内容偏移，
  // 例如 (+5, +120) 会选中整行 A3:Z3）——所以固定用这个偏移，再从日志读回真实落点。
  await openMenuOnCanvasAt(page, box.x + 70, box.y + 120);
  const a1 = await page.evaluate(
    () =>
      ((window as never as { __p0: { log: Array<{ kind: string; detail?: { a1?: string } }> } }).__p0.log
        .filter((entry) => entry.kind === 'menu:open')
        .pop()?.detail?.a1 ?? '') as string,
  );
  expect(a1, '右键菜单应记录打开那一刻的选区快照 A1（单格或区域）').toMatch(/^[A-Z]+\d+(:[A-Z]+\d+)?$/);
  return a1;
}

/** 全部条目的快照预览都应满足的条件收集器 */
async function visibleColsOfAllPreviews(page: Page): Promise<string[]> {
  return page.locator('[data-testid="snapshot-preview"]').evaluateAll((nodes) =>
    nodes.map((node) => node.getAttribute('data-visible-cols') ?? ''),
  );
}

/**
 * 数"工作区这张小表格每行放了几个格子"。
 *
 * 每行几格**不再由某个设置直接给出**，而是"面板宽度 ÷ 格子最小宽度"自动决定的，
 * 所以只能从渲染结果数：按条目的 y 坐标（第二行起格子高度一致）分组，取最大的一组。
 */
async function tilesPerRow(page: Page): Promise<number> {
  return page.locator('[data-testid="workspace-item"]').evaluateAll((nodes) => {
    const buckets = new Map<number, number>();
    for (const node of nodes) {
      const top = Math.round(node.getBoundingClientRect().top / 4) * 4; // 4px 容差归组
      buckets.set(top, (buckets.get(top) ?? 0) + 1);
    }
    return buckets.size === 0 ? 0 : Math.max(...buckets.values());
  });
}

/** 清空工作区（现在要二次确认：点「清空」→ 点「清空全部 N 个」） */
async function clearWorkspace(page: Page): Promise<void> {
  await page.locator('[data-testid="workspace-clear"]').click();
  await expect(page.locator('[data-testid="workspace-clear-confirm"]')).toBeVisible();
  await page.locator('[data-testid="workspace-clear-all"]').click();
  await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(0);
}

/**
 * 往工作区搜索框里**真的敲字**（而不是 `fill()` 直接写值）。
 *
 * 为什么这么写：`fill()` 是"直接改 DOM 值 + 派发一个 input 事件"，长跑（多个 spec 连着跑）里
 * 偶发过一次 **React 没处理那个事件**——受控输入随即被重渲染回空值，断言就红了。
 * 排查记录：单独跑、整文件跑都复现不出来；失败 trace 里页面没有重载（`[vite] connecting` 只出现一次）、
 * 输入框节点也没被替换（Playwright 的 target 标记还在），所以判定为测试驱动方式的偶发问题，不是应用行为。
 * 真实按键每个字符都会带上"当前完整文本"，即使丢一次事件，后面的按键也会把状态纠正过来；
 * 顺带更接近用户真实操作。
 */
async function typeSearch(page: Page, text: string): Promise<void> {
  const search = page.locator('[data-testid="workspace-search"]');
  await search.click();
  await search.press('Control+a');
  await search.press('Delete');
  if (text !== '') await search.pressSequentially(text, { delay: 10 });
  await expect(search, `搜索框里应当是「${text}」`).toHaveValue(text);
}

/**
 * 直接把工作区清空（走 `__p0.clearWorkspace` 钩子）。
 *
 * 用于**用例内部**重置状态：不依赖 DOM 按钮，也就不会被"面板此刻有没有重渲染完"影响。
 * 面向用户的清空路径（二次确认、只清显示项）另有用例专门守。
 */
async function resetWorkspace(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as never as { __p0: { clearWorkspace: () => void } }).__p0.clearWorkspace();
  });
  await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(0);
  await expect(page.locator('[data-testid="workspace-empty"]')).toBeVisible();
}

test.describe('工作区与内容搬运', () => {
  test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures 生成样本');

  test('① 两个入口都拆成独立单元格条目：真实拖动 + 右键复制，行优先且新条目在最前', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    await expect(page.locator('[data-testid="workspace-empty"]')).toBeVisible();

    // ---- 入口 A：真实鼠标拖动（拖拽模式下按住即可拖，不需要长按）----
    // 选区先铺成整张已用区域，再从**有内容的格子**（探针实测：x+175 起落在 D/E/F 列、
    // y+100 落在第 3 行）按下拖走。落点必须是"有值/有公式"的格子，否则新模型下
    // "跳过空单元格"会让工作区一条都不产生（旧的 (+120,+90) 正好落在空格 A3 上）。
    const box = await canvasBox(page);
    const sidebar = await page.locator('.workspace-host').boundingBox();
    if (!sidebar) throw new Error('侧边栏不可见');

    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:H25');
    });
    await page.waitForTimeout(200);
    const expectedDragged = await countNonEmptyCells(page, 25, 8);
    // 先确保是"保留内容"语义（设置随会话持久化，可能被别的用例改成剪切过）
    const keepSource = page.locator('[data-testid="toolbar-keep-source"]');
    if (!(await keepSource.isChecked())) await keepSource.check();

    await page.mouse.move(box.x + 175, box.y + 100);
    await page.mouse.down();
    await page.mouse.move(sidebar.x + sidebar.width / 2, sidebar.y + sidebar.height / 2, { steps: 14 });
    await page.waitForTimeout(60);
    await page.mouse.up();

    await expect(page.locator('[data-testid="workspace-item"]').first()).toBeVisible({ timeout: 5000 });
    await expect(page.locator('[data-testid="workspace-empty"]')).toHaveCount(0);

    const dragged = await workspaceItems(page);
    expect(dragged.length, '拖到侧边栏应把整片区域的非空格子都拆成条目').toBe(expectedDragged);
    // 新模型：**一条 = 一个单元格**；旧模型下这里会是一条 data-cols=8 的"整片区域"条目
    expect(dragged.every((item) => item.rows === 1 && item.cols === 1), '条目必须都是 1×1 单格').toBe(true);
    expect(dragged[0].a1, '条目的来源 A1 是单个格子记号').toMatch(/^[A-Z]+\d+$/);
    expect(await countNonEmptyCells(page, 25, 8), '"保留内容"语义下源数据不受影响').toBe(expectedDragged);

    const draggedCols = await visibleColsOfAllPreviews(page);
    expect(draggedCols.length, '每个条目都应挂一个快照预览').toBe(dragged.length);
    expect(draggedCols.every((value) => value === '1'), '单格条目的可见列数恒为 1').toBe(true);

    await resetWorkspace(page);

    // ---- 入口 B：右键「复制到工作区」（确定性入口，用于断言拆分规则）----
    // （入口 A 的条目已在上面清掉：否则两批条目叠在一起，"最后一个条目是 A1"这类顺序断言就没意义）
    // 探针实测：canvas 左上角 (+70, +120) 落点稳定落在 A1:H25 这块选区内部（冻结列/行会让
    // 像素偏移 ≠ 内容偏移，所以这里不猜格子，而是从 menu:open 日志读回真实落点再推导期望）。
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:H25');
    });
    await page.waitForTimeout(200);

    const menuA1 = await rightClickInsideSelection(page);
    expect(menuA1, '右键落点应落在选区内部（否则菜单拿到的是单格快照）').toMatch(/^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/);

    // 下拉起点：选区快照是 A1:H25 时从 A1 开始；万一只选中一行，则从该行的落点列开始
    const menuMatch = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(menuA1)!;
    const startCol = [...menuMatch[1]].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
    const startRow = Number(menuMatch[2]) - 1;
    const endCol = menuMatch[3] ? [...menuMatch[3]].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1 : startCol;
    const endRow = menuMatch[4] ? Number(menuMatch[4]) - 1 : startRow;

    await page.locator('[data-testid="context-menu-workspace-copy"]').click();
    await page.waitForTimeout(500);
    await expect(page.locator('[data-testid="context-menu"]'), '动作执行后菜单应关闭').toHaveCount(0);

    // 期望条目数 = 该快照区域里的非空格子数（现算，不写死）
    const expectedCount = await countNonEmptyCells(page, endRow + 1, endCol + 1);
    expect(expectedCount, '选区里应有一批非空格子').toBeGreaterThan(0);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expectedCount);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(expectedCount));
    expect((await visibleColsOfAllPreviews(page)).length, 'DOM 里的条目数 = 非空格子数').toBe(expectedCount);

    const copied = await workspaceItems(page);
    expect(copied.length, '条目数 = 该区域非空格子数').toBe(expectedCount);
    expect(copied.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1').toBe(true);

    // 行优先：从落点开始，先横着走完这一行再换下一行（跳过空格子）
    const expectedOrder = await page.evaluate(
      ({ startRow, endRow, startCol, endCol }: { startRow: number; endRow: number; startCol: number; endCol: number }) => {
        const hooks = (window as never as { __p0: { getDisplayValue: (r: number, c: number) => unknown } }).__p0;
        const out: string[] = [];
        for (let row = startRow; row <= endRow; row += 1) {
          for (let col = startCol; col <= endCol; col += 1) {
            const value = hooks.getDisplayValue(row, col);
            const text = value === null || value === undefined ? '' : String(value);
            if (text !== '') out.push(`${String.fromCharCode(65 + col)}${row + 1}`);
          }
        }
        return out;
      },
      { startRow, endRow, startCol, endCol },
    );
    expect(copied.map((item) => item.a1), '拆出来的条目必须按行优先顺序排列').toEqual(expectedOrder);
    expect(copied[0].a1, '新条目排在最前').toBe(expectedOrder[0]);
    expect(expectedOrder[0], '探针实测：A1:H25 的第一个非空格子就是 A1').toBe('A1');

    // 每个条目的 title 都是「{内容} · 来源 {工作表}!{A1}（右键可复制/剪切/删除）」
    const domTitles = await page
      .locator('[data-testid="workspace-item"]')
      .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('title') ?? ''));
    expect(domTitles.length, 'DOM 里的条目数应与快照一致').toBe(expectedCount);
    const domRefs = domTitles.map((title) => /来源 [^!]+!([A-Z]+\d+)/.exec(title)?.[1] ?? '');
    expect(domRefs, 'DOM 里的条目顺序也要行优先').toEqual(expectedOrder);
    const firstItem = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { getDisplayValueByA1: (a1: string) => unknown; getActiveSheet: () => { getSheetName: () => string } | null };
      }).__p0;
      return {
        text: String(hooks.getDisplayValueByA1('A1') ?? ''),
        sheet: hooks.getActiveSheet()?.getSheetName() ?? '',
      };
    });
    expect(domTitles[0], '第一个条目是来源 A1 的那一格').toBe(
      `${firstItem.text} · 来源 ${firstItem.sheet}!A1（右键可复制/剪切/删除）`,
    );
    expect(domTitles[domTitles.length - 1], '最后一个条目是行优先的最后一格 A21').toContain('来源 样式!A21');
  });

  test('② 条目写回单元格：只改内容，目标单元格格式一字不变', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    const result = await page.evaluate(async () => {
      const hooks = (window as never as {
        __p0: {
          snapshotSelectionToWorkspace: (a1?: string) => {
            items: Array<{ id: string; rows: number; cols: number; source: { a1: string } }>;
            skippedEmpty: number;
            error?: string;
          };
          getWorkspaceItems: () => Array<{ id: string; a1: string; rows: number; cols: number }>;
          pasteWorkspaceItem: (id: string, a1: string) => { ok: boolean };
          getCellStyle: (r: number, c: number) => unknown;
          getDisplayValue: (r: number, c: number) => string | null;
        };
      }).__p0;

      // D5 是夹具里的「旋转45°」格（探针实测在 D3:F6 这批有内容的格子里）
      const source = 'D5';
      const sourceText = String(hooks.getDisplayValue(4, 3) ?? '');
      if (sourceText === '') return { error: `${source} 应该是非空的` as const };

      const extracted = hooks.snapshotSelectionToWorkspace(source);
      await new Promise((r) => setTimeout(r, 300)); // 等 React 把条目提交到 DOM
      if (extracted.error) return { error: `提取失败：${extracted.error}` as const };
      if (extracted.items.length !== 1) return { error: `${source} 应拆出 1 条，实际 ${extracted.items.length}` as const };
      const item = extracted.items[0];
      if (item.rows !== 1 || item.cols !== 1) return { error: `条目不是 1×1：${item.rows}×${item.cols}` as const };
      if (hooks.getWorkspaceItems().length !== 1) return { error: '工作区条目数不对' as const };

      const styleBefore = JSON.stringify(hooks.getCellStyle(2, 0) ?? null); // A3 有红色四边边框
      const valueBefore = hooks.getDisplayValue(2, 0);
      const paste = hooks.pasteWorkspaceItem(item.id, 'A3');
      await new Promise((r) => setTimeout(r, 400));
      return {
        paste,
        source,
        sourceText,
        itemA1: item.source.a1,
        styleBefore,
        styleAfter: JSON.stringify(hooks.getCellStyle(2, 0) ?? null),
        valueBefore,
        valueAfter: hooks.getDisplayValue(2, 0),
      };
    });

    expect('error' in result ? result.error : undefined).toBeUndefined();
    if ('error' in result) return;
    expect(result.itemA1, '单格条目的来源就是一个格子').toBe('D5');
    expect(result.paste.ok).toBe(true);
    expect(result.valueAfter, '内容应被写入').not.toBe(result.valueBefore);
    expect(result.valueAfter, '写回的正是来源格的内容').toBe(result.sourceText);
    expect(result.styleAfter, '目标单元格格式必须一字不变').toBe(result.styleBefore);
    await expect(page.locator('[data-testid="workspace-item"]'), '写回是复制语义，条目仍在').toHaveCount(1);
  });

  test('③ 工作区跨表共享：另一张表也能写回条目，且源表内容不被清掉', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-rules.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-rules.xlsx', { features: true });

    // 在第一张表写入可辨识内容，再取该格放进工作区（hook 新契约：返回条目数组）
    const created = await page.evaluate(async () => {
      const hooks = (window as never as {
        __p0: {
          setValue: (r: number, c: number, v: string) => void;
          snapshotSelectionToWorkspace: (a1?: string) => {
            items: Array<{ id: string; rows: number; cols: number; source: { a1: string } }>;
            skippedEmpty: number;
          };
        };
      }).__p0;
      hooks.setValue(0, 0, '跨表源内容');
      await new Promise((r) => setTimeout(r, 300));
      const result = hooks.snapshotSelectionToWorkspace('A1');
      const item = result.items[0] ?? null;
      return item
        ? { id: item.id, a1: item.source.a1, rows: item.rows, cols: item.cols, count: result.items.length, skipped: result.skippedEmpty }
        : null;
    });
    expect(created?.id, '工作区条目创建失败').toBeTruthy();
    if (!created?.id) return;
    expect(created.count, 'A1 有内容 → 只应拆出 1 个条目').toBe(1);
    expect(created.rows, '条目是 1×1 单格').toBe(1);
    expect(created.cols, '条目是 1×1 单格').toBe(1);
    // 收窄成 string：钩子返回的 id 是可选字段，直接传进 evaluate 会与形参类型不符（tsc 会报 TS2769）
    const createdId: string = created.id;
    const createdA1: string = created.a1;

    // 切到另一张表后写回：条目仍然在（跨表共享），内容写进当前表，格式不变
    await activateSheet(page, '数据验证');

    const pasted = await page.evaluate(
      async ({ id, a1 }: { id: string; a1: string }) => {
        const hooks = (window as never as {
          __p0: {
            pasteWorkspaceItem: (id: string, a1: string) => { ok: boolean };
            getCellStyle: (r: number, c: number) => unknown;
            getDisplayValue: (r: number, c: number) => string | null;
            getDisplayValueByA1: (a1: string) => unknown;
            getWorkspaceItems: () => unknown[];
          };
        }).__p0;
        const styleBefore = JSON.stringify(hooks.getCellStyle(0, 0) ?? null); // A1 是表头
        const valueBefore = hooks.getDisplayValue(0, 0);
        const paste = hooks.pasteWorkspaceItem(id, a1);
        await new Promise((r) => setTimeout(r, 400));
        return {
          paste,
          styleBefore,
          styleAfter: JSON.stringify(hooks.getCellStyle(0, 0) ?? null),
          valueBefore,
          valueAfter: hooks.getDisplayValue(0, 0),
          writtenAtA1: String(hooks.getDisplayValueByA1(a1) ?? ''),
          itemCount: hooks.getWorkspaceItems().length,
        };
      },
      { id: createdId, a1: createdA1 },
    );

    expect(pasted.paste.ok).toBe(true);
    expect(pasted.writtenAtA1, `另一张表的 ${createdA1} 应拿到条目内容`).toBe('跨表源内容');
    expect(pasted.valueAfter).not.toBe(pasted.valueBefore);
    expect(pasted.styleAfter, '跨表写回也不得改格式').toBe(pasted.styleBefore);
    expect(pasted.itemCount, '条目不应因为切表而消失或减少').toBe(1);

    // 写回是复制语义：源表内容必须原样还在
    await activateSheet(page, '条件格式');
    expect(await valueAt(page, 'A1'), '源表内容不应被清掉').toBe('跨表源内容');
    expect((await workspaceItems(page)).length, '切回源表后条目仍在').toBe(1);
  });

  test('④ 两区域互换：内容对调，两侧格式各自留在原格', async ({ page }) => {
    await waitForBoot(page);

    const probe = await page.evaluate(async () => {
      const hooks = (window as never as {
        __p0: {
          getCellStyle: (r: number, c: number) => unknown;
          getDisplayValue: (r: number, c: number) => string | null;
          swap: (a: string, b: string) => { ok: boolean };
        };
      }).__p0;

      const styleOf = (r: number, c: number) => JSON.stringify(hooks.getCellStyle(r, c) ?? null);
      const before = {
        a3: hooks.getDisplayValue(2, 0),
        a4: hooks.getDisplayValue(3, 0),
        styleA3: styleOf(2, 0),
        styleA4: styleOf(3, 0),
      };
      const swapped = hooks.swap('A3', 'A4');
      await new Promise((r) => setTimeout(r, 400));
      return {
        swapped,
        before,
        after: {
          a3: hooks.getDisplayValue(2, 0),
          a4: hooks.getDisplayValue(3, 0),
          styleA3: styleOf(2, 0),
          styleA4: styleOf(3, 0),
        },
      };
    });

    // 前提：两个格子内容/格式确实不同，否则"互换"与"没动"无法区分
    expect(probe.before.a3).toBeTruthy();
    expect(probe.before.a4).toBeTruthy();
    expect(probe.before.a3).not.toBe(probe.before.a4);
    expect(probe.before.styleA3).not.toBe(probe.before.styleA4);

    expect(probe.swapped.ok, '互换命令必须执行成功').toBe(true);
    expect(probe.after.a3, '内容应互换').toBe(probe.before.a4);
    expect(probe.after.a4, '内容应互换').toBe(probe.before.a3);
    expect(probe.after.styleA3, 'A3 的格式必须原样留在 A3').toBe(probe.before.styleA3);
    expect(probe.after.styleA4, 'A4 的格式必须原样留在 A4').toBe(probe.before.styleA4);
  });

  /**
   * 需求 ⑨：「+」从当前表快速导入（单个输入框自动识别）。
   *
   * 对话框的解析细节（行/列/区域的裁剪与非法输入）已有单元测试覆盖，这里守住**接线与拆分口径**：
   * 按钮能开对话框、确认能真的生成条目、**条目数 = 该区域非空格子数**、非法输入报错且不生成条目。
   */
  test('⑤ 「+」导入：非法输入报错；合法区域按非空格子数拆成单格条目（不再是一整片）', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // 前提自检：夹具的已用区域是 A1:H21（下面"跳过空格"的断言要靠它）
    const used = await usedRangeSize(page);
    expect(used.rows, '夹具已用区域行数').toBe(STYLES_USED_RANGE.rows);
    expect(used.cols, '夹具已用区域列数').toBe(STYLES_USED_RANGE.cols);
    expect(used.a1).toBe('A1:H21');

    // 「+」现在在工作区右上角（不在顶部工具栏）
    await page.locator('[data-testid="workspace-add"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toBeVisible();

    // 先试一个非法输入：应给出内联错误、禁用「导入」，且不生成条目
    await page.locator('[data-testid="import-input"]').fill('B : D');
    await expect(page.locator('[data-testid="import-error"]')).toBeVisible();
    await expect(page.locator('[data-testid="import-confirm"]'), '非法输入应禁用导入').toBeDisabled();
    expect((await workspaceItems(page)).length, '非法输入不应生成条目').toBe(0);

    // 再导入整整 25 行（比已用区域还大）：空行/空格子必须被跳过
    await page.locator('[data-testid="import-input"]').fill('A1:H25');
    await expect(page.locator('[data-testid="import-confirm"]'), '合法输入应可导入').toBeEnabled();

    /**
     * 新面板：**实时预览**替代了原来那个不可点的"跳过空单元格"复选框。
     * 预览里的数字必须与真正加入的条目数一致（"跳过 N 个空内容"就是 200 − 非空数）。
     */
    const expected = await countNonEmptyCells(page, 25, 8);
    expect(expected, 'A1:H25 里应有一批非空格子').toBeGreaterThan(0);
    const summary = page.locator('[data-testid="import-preview-summary"]');
    await expect(summary, '预览要报出将加入的格数').toContainText(`将加入 ${expected} 个单元格`);
    await expect(summary, '预览要报出被跳过的空格数').toContainText(`跳过 ${25 * 8 - expected} 个空内容`);
    await expect(page.locator('[data-testid="import-preview-samples"]'), '预览要列出前几格内容').toBeVisible();
    await expect(page.locator('[data-testid="import-confirm"]'), '确认键上带格数').toContainText(`加入 ${expected} 格`);
    await expect(page.locator('[data-testid="import-skip-empty"]'), '无用元素已被移除').toHaveCount(0);

    await page.locator('[data-testid="import-confirm"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toHaveCount(0);

    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(expected));
    expect(expected, '选的 25×8=200 格里必须有空格子被跳过').toBeLessThan(25 * 8);

    const items = await workspaceItems(page);
    expect(items.length, '条目数 = 该区域非空格子数').toBe(expected);
    expect(items.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1').toBe(true);
    expect(items.map((item) => item.a1)).toEqual([...items.map((item) => item.a1)].sort(compareA1RowMajor));

    // 不产生空条目：每条预览的单元格文本都必须非空
    const texts = await page
      .locator('[data-testid="snapshot-preview"]')
      .evaluateAll((nodes) => nodes.map((node) => (node.textContent ?? '').trim()));
    expect(texts.length).toBe(expected);
    expect(texts.every((text) => text !== ''), '跳过空单元格：不应出现空内容的条目').toBe(true);

    const cols = await visibleColsOfAllPreviews(page);
    expect(cols.every((value) => value === '1'), '单格条目的可见列数恒为 1').toBe(true);
  });

  /**
   * 新契约：`__p0.snapshotSelectionToWorkspace(a1)` **返回拆出来的条目数组**。
   *
   * 断言口径（用户要求）：**条目数 = 该区域非空格子数**；`skippedEmpty` = 区域总格子数 − 条目数；
   * 每个条目都是 1×1 且带 `cellSize`（来源单元格的像素尺寸，工作区按它展示）。
   */
  test('⑥ 测试钩子新契约：返回 1×1 条目数组，条目数 = 非空格子数、skippedEmpty = 空格子数', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // 探针实测：A1:F2 十二格里只有 7 格有内容 —— A1（合并标题）+ A2:F2（表头）
    // 第 1 行的 B1:F1 属于 A1:F1 合并区，**没有值** → 按"跳过空单元格"的规则必须被跳过
    const dense = await extractViaHook(page, 'A1:F2');
    expect(dense.error).toBeUndefined();
    expect(dense.count, 'A1:F2 里 7 个非空格子 → 7 个条目').toBe(7);
    expect(dense.skippedEmpty, 'A1:F1 合并后 B1:F1 是空格子 → 跳过 5 个').toBe(5);
    expect(dense.count + dense.skippedEmpty, '条目数 + 跳过数 = 区域总格子数').toBe(2 * 6);
    expect(dense.truncated, '远低于 500 上限').toBe(0);
    expect(dense.shapes.every(([rows, cols]) => rows === 1 && cols === 1), '每个条目都是 1×1').toBe(true);
    expect(new Set(dense.ids).size, '条目 id 必须唯一').toBe(7);
    expect(dense.a1s, '行优先：先横着走完第一行（B1:F1 是空 → 只剩 A1），再换行').toEqual([
      'A1', 'A2', 'B2', 'C2', 'D2', 'E2', 'F2',
    ]);
    expect(dense.labels).toEqual(dense.a1s); // 单格条目的 label 就是它自己的 A1

    // 条目带来源单元格尺寸（工作区按 cellSize 展示，夹在 28–140 × 16–64）
    const sizes = await page.evaluate(() => {
      const items = (window as never as {
        __p0: {
          snapshotSelectionToWorkspace: (a1?: string) => {
            items: Array<{ cellSize?: { width: number; height: number } }>;
          };
        };
      }).__p0.snapshotSelectionToWorkspace('A20');
      return items.items.map((item) => item.cellSize ?? null);
    });
    expect(sizes, '单格条目必须带 cellSize').toHaveLength(1);
    expect(sizes[0], 'A20 的来源单元格尺寸应当读得到').not.toBeNull();
    // A20 在夹具里被单独设成"行高 48"（`tools/make-fixtures.mjs` 的 `ws.getRow(20).height = 48`）
    // 注意：Univer 的 `getRowHeight()` 会做设备像素换算并夹到自己的上下限，所以这里按"明显高于
    // 默认行高"来断言来源尺寸确实被保留下来（精确的 48 由解析层的行高用例守）。
    expect(sizes[0]!.height, 'A20 的行高应明显高于默认行高').toBeGreaterThan(30);
    expect(sizes[0]!.height, '行高不应离谱（展示时还要夹到 64）').toBeLessThanOrEqual(96);
    expect(sizes[0]!.width, '列宽应读得到').toBeGreaterThan(0);

    await resetWorkspace(page);

    // A1:H25：25×8=200 格，其中只有一批非空 → 条目数必须正好等于非空格子数，其余全部跳过
    const big = await extractViaHook(page, 'A1:H25');
    const expected = await countNonEmptyCells(page, 25, 8);
    expect(expected, 'A1:H25 里应有一批非空格子').toBeGreaterThan(0);
    expect(big.count, '条目数 = 该区域非空格子数').toBe(expected);
    expect(big.skippedEmpty, 'skippedEmpty = 区域总格子数 − 条目数').toBe(25 * 8 - expected);
    expect(big.count, '不是"一整片 1 条"，而是逐格拆开').toBeGreaterThan(1);
    expect(big.shapes.every(([rows, cols]) => rows === 1 && cols === 1), '每个条目都是 1×1').toBe(true);
    expect(big.a1s, '按行优先顺序').toEqual([...big.a1s].sort(compareA1RowMajor));
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
  });

  /**
   * 需求 ⑨ 新增入口：**一键全搬**（`data-testid="import-whole-sheet"`，文案「整个已用区域」）。
   *
   * 新面板里它变成"帮我填范围"的按钮：点一下把已用区域填进输入框（并立刻给出预览），
   * 再由用户按确认键提交 —— 用户不必先按下去才知道会搬进来多少格。
   *
   * 断言：
   *   - 条目数 > 5 且**恰好等于已用区域里的非空格子数**（证明没有把空格子也搬进来）；
   *   - 条目数 < 已用区域总格子数（证明空格子确实被跳过，日志里也应有 `workspace:skip-empty`）；
   *   - **没有任何条目是空内容**（逐条读预览文本）；
   *   - 每条都是 1×1、来源 A1 互不重复、行优先排列。
   */
  test('⑦ 一键全搬「整个已用区域」：条目数 = 非空格子数、跳过全部空格、无空条目', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    await page.locator('[data-testid="workspace-add"]').click();
    const whole = page.locator('[data-testid="import-whole-sheet"]');
    await expect(whole, '「整个已用区域」按钮必须存在').toBeVisible();
    await expect(whole, '有内容时按钮可用').toBeEnabled();
    await expect(whole, '按钮上应标出已用区域').toContainText('A1:H21');

    const totalCells = STYLES_USED_RANGE.rows * STYLES_USED_RANGE.cols; // 168
    const expected = await countNonEmptyCells(page, STYLES_USED_RANGE.rows, STYLES_USED_RANGE.cols); // 35

    // 点一下：只**填范围**（输入框拿到已用区域，预览立刻报数），不直接提交
    await whole.click();
    await expect(page.locator('[data-testid="import-input"]')).toHaveValue('A1:H21');
    await expect(page.locator('[data-testid="import-preview-summary"]')).toContainText(
      `将加入 ${expected} 个单元格 · 跳过 ${totalCells - expected} 个空内容`,
    );
    await expect(page.locator('[data-testid="import-confirm"]')).toContainText(`加入 ${expected} 格`);

    await page.locator('[data-testid="import-confirm"]').click();
    await expect(page.locator('[data-testid="import-dialog"]')).toHaveCount(0);

    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(expected));

    const items = await workspaceItems(page);
    expect(items.length, '一键全搬应搬进来一批条目').toBeGreaterThan(5);
    expect(items.length, '条目数 = 已用区域非空格子数').toBe(expected);
    expect(items.length, '必须少于区域总格子数 → 空格子被跳过').toBeLessThan(totalCells);
    expect(new Set(items.map((item) => item.a1)).size, '每个来源格子只应产生一个条目').toBe(expected);
    expect(items.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1 单格').toBe(true);
    expect(items.map((item) => item.a1), '按行优先顺序拆出').toEqual([...items.map((item) => item.a1)].sort(compareA1RowMajor));

    // 逐条读预览文本：不允许出现空内容条目
    const previews = page.locator('[data-testid="snapshot-preview"]');
    await expect(previews).toHaveCount(expected);
    const texts = await previews.evaluateAll((nodes) => nodes.map((node) => (node.textContent ?? '').trim()));
    expect(texts.length).toBe(expected);
    expect(texts.every((text) => text !== ''), '一键全搬的结果里不能有任何空条目').toBe(true);
    expect(
      texts.filter((text) => text === '').length,
      '空条目条数',
    ).toBe(0);

    // 跳过空单元格的证据：日志里出现 workspace:skip-empty，且跳过数 = 区域总格子数 − 条目数
    const skipLog = await page.evaluate(
      () =>
        ((window as never as {
          __p0: { log: Array<{ kind: string; detail?: { skipped?: number } }> };
        }).__p0.log
          .filter((entry) => entry.kind === 'workspace:skip-empty')
          .pop()?.detail ?? null) as { skipped?: number } | null,
    );
    expect(skipLog, '一键全搬应记录 workspace:skip-empty').not.toBeNull();
    expect(skipLog?.skipped, '跳过数应等于区域空格子数').toBe(totalCells - expected);

    // 再验一次：换一张表（另一份夹具）后，「整个已用区域」按钮上的区域记号跟着当前表走
    await importFixture(page, 'fixture-rules.xlsx');
    await activateSheet(page, '数据验证');
    await page.waitForTimeout(300);
    await page.locator('[data-testid="workspace-add"]').click();
    const wholeOnOtherSheet = page.locator('[data-testid="import-whole-sheet"]');
    await expect(wholeOnOtherSheet, '换表后按钮仍可用（该表有内容）').toBeEnabled();
    await expect(wholeOnOtherSheet).toContainText(/^整个已用区域[A-E][0-9]/);
    await page.keyboard.press('Escape');
  });

  /**
   * 需求：工作区外观与宽度（新模型口径）。
   * - 条目 = 一格一个小方块：**可见列数恒为 1**，不再有"还有 N 列"提示；
   * - 「每行放几格」由 **面板宽度 ÷ 格子最小宽度（`tileMinWidth`）自动决定**，
   *   所以拉宽面板 / 改"格宽"都会让每行的格子数变化；
   * - 表格区与工作区之间的分隔条可拖动改宽度（变宽且右边界不动）。
   */
  test('⑧ 单格条目恒 1 列；「格宽」与面板宽度自动决定每行几格；分隔条可拖动改宽度', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // 先走新契约的 hook 拿一次 A1:F6（探针实测：36 格里只有 19 格有内容）
    const viaHook = await extractViaHook(page, 'A1:F6');
    expect(viaHook.error).toBeUndefined();
    expect(viaHook.count, 'A1:F6 里 19 个非空格子 → 19 个条目').toBe(A1_F6_NON_EMPTY.length);
    expect(viaHook.skippedEmpty, '其余 17 格是空格 → 全部跳过').toBe(36 - A1_F6_NON_EMPTY.length);
    expect(viaHook.a1s, '行优先（第 1 行只有 A1 有值，B1:F1 是空）').toEqual(A1_F6_NON_EMPTY);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(A1_F6_NON_EMPTY.length);
    await resetWorkspace(page);

    // 再走真实入口（右键「复制到工作区」）——两者语义必须一致
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:F6');
    });
    await page.waitForTimeout(200);
    const menuA1 = await rightClickInsideSelection(page);
    expect(menuA1, '右键落点应落在 A1:F6 这块选区内部').toBe('A1:F6');
    await page.locator('[data-testid="context-menu-workspace-copy"]').click();
    await page.waitForTimeout(500);

    const expected = await countNonEmptyCells(page, 6, 6);
    expect(expected, 'A1:F6 里的非空格子数').toBe(A1_F6_NON_EMPTY.length);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
    expect((await workspaceItems(page)).map((item) => item.a1), '真实入口的拆分顺序也要行优先').toEqual(
      A1_F6_NON_EMPTY,
    );

    // 单格条目：data-cols / data-rows / data-visible-cols 全是 1；也不该出现"还有 N 列"提示
    const preview = page.locator('[data-testid="snapshot-preview"]').first();
    await expect(preview).toHaveAttribute('data-rows', '1');
    await expect(preview).toHaveAttribute('data-cols', '1');
    await expect(preview).toHaveAttribute('data-visible-cols', '1');
    await expect(page.locator('[data-testid="snapshot-preview-more"]'), '单格条目已无"还有 N 列/行"提示').toHaveCount(0);

    // "保留来源尺寸"落地：预览的列宽/行高夹在 [tileMinWidth−16, 200] × [18, 64] 内
    const cellBox = await page
      .locator('[data-testid="snapshot-preview"]')
      .first()
      .evaluate((node) => {
        const col = node.querySelector('col');
        const width = Number((col?.getAttribute('style') ?? '').replace(/[^\d.]/g, ''));
        const rowHeight = Number((node.querySelector('tr')?.getAttribute('style') ?? '').replace(/[^\d.]/g, ''));
        return { width, rowHeight };
      });
    expect(cellBox.width, '单元格宽度应 ≥ tileMinWidth−16（默认 112）').toBeGreaterThanOrEqual(128 - 16);
    expect(cellBox.width, '单元格宽度最多 200').toBeLessThanOrEqual(200);
    expect(cellBox.rowHeight, '行高应落在 18–64 的夹取区间内').toBeGreaterThanOrEqual(18);
    expect(cellBox.rowHeight).toBeLessThanOrEqual(64);

    // 「格宽」默认 128：挂到 .ws-list 的 data-tile-min-width 与栅格样式上
    const list = page.locator('.ws-list');
    await expect(list).toHaveAttribute('data-tile-min-width', '128');
    await expect(list).toHaveAttribute('data-visible-count', String(expected));
    const tileWidth = page.locator('[data-testid="workspace-tile-width"]');
    await expect(tileWidth, '默认格宽').toHaveValue('128');

    // 每行几格是"面板宽度 ÷ 格宽"自动算的：默认宽度下先量一次基线
    const colsAtDefault = await tilesPerRow(page);
    expect(colsAtDefault, '默认宽度下每行应能放下格子（≥1）').toBeGreaterThanOrEqual(1);
    expect(colsAtDefault * 128, '每行格子数 × 格宽不应超过面板宽度太多').toBeLessThanOrEqual(288 + 128);

    // 把「格宽」拉到最大（220）→ 每行放得下的格子变少（通常只剩 1 格），且立即生效
    await tileWidth.fill('220');
    await expect(list).toHaveAttribute('data-tile-min-width', '220');
    await expect(tileWidth).toHaveValue('220');
    const colsAtWide = await tilesPerRow(page);
    expect(colsAtWide, '格宽变大 → 每行格子数应减少或持平').toBeLessThanOrEqual(colsAtDefault);
    await expect(page.locator('[data-testid="workspace-item"]'), '改格宽不影响条目数').toHaveCount(expected);
    await tileWidth.fill('128');
    await expect(list).toHaveAttribute('data-tile-min-width', '128');

    // 每个条目的预览仍然是 1 列（"每行几格"只是排布）
    const cols = await visibleColsOfAllPreviews(page);
    expect(cols.length).toBe(expected);
    expect(cols.every((value) => value === '1'), '每行几格只是排布，单格条目仍然 1 列').toBe(true);

    // 把面板拉宽 → 每行自动放得下更多格子（不需要改任何设置）
    const body = page.locator('.body');
    const panel = page.locator('.ws-panel');
    const widthBefore = Number(await body.getAttribute('data-sidebar-width'));
    const rectBefore = await panel.boundingBox();
    const box = await page.locator('[data-testid="sidebar-splitter"]').boundingBox();
    if (!box || !rectBefore) throw new Error('分隔条或工作区不可见');
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x - 80, box.y + box.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(250);

    const widthAfter = Number(await body.getAttribute('data-sidebar-width'));
    const rectAfter = await panel.boundingBox();
    expect(widthAfter, '往左拖应把工作区加宽').toBeGreaterThan(widthBefore + 40);
    expect(rectAfter, '工作区面板仍应可见').not.toBeNull();
    expect(rectAfter!.width, '工作区面板本身也要真的变宽（不能只位移）').toBeGreaterThan(rectBefore.width + 40);
    expect(
      Math.abs(rectAfter!.x + rectAfter!.width - (rectBefore.x + rectBefore.width)),
      '右边界应保持吸附（位移 < 4px）',
    ).toBeLessThan(4);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(expected);
  });

  /**
   * 需求：两个语义开关 + 条目剪贴板。
   * - 顶部「拖到工作区保留内容」取消后，拖到工作区 = 剪切（源内容清空、格式保留、可撤销）；
   * - 条目右键「复制」后可粘贴进表格（只写内容、目标格式不变）；右键「删除」移除条目；
   * - 工作区「写回表格后移除条目」勾选后，条目写回即消失。
   */
  test('⑨ 拖放语义开关 + 条目右键复制/粘贴/删除', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // --- 1) 取消"保留内容"：拖到工作区应当是剪切 ---
    const keep = page.locator('[data-testid="toolbar-keep-source"]');
    await expect(keep).toBeChecked();
    await keep.uncheck();

    // 真实手势：先选中整张已用区域，再从有内容的格子（x+175 / y+100 → D3 一带）按住拖进工作区
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A1:H25');
    });
    await page.waitForTimeout(200);
    const valueBefore = await valueAt(page, 'A1');
    const expectedDragged = await countNonEmptyCells(page, 25, 8);

    const box = await canvasBox(page);
    const sidebar = await page.locator('.ws-panel').boundingBox();
    if (!sidebar) throw new Error('工作区不可见');
    await page.mouse.move(box.x + 175, box.y + 100);
    await page.mouse.down();
    await page.mouse.move(sidebar.x + sidebar.width / 2, sidebar.y + 160, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    const dragged = await workspaceItems(page);
    expect(dragged.length, '剪切语义下应把整片区域的非空格子都拖进工作区').toBe(expectedDragged);
    expect(dragged.every((item) => item.rows === 1 && item.cols === 1), '每条都必须是 1×1').toBe(true);
    expect(await valueAt(page, 'A1'), '取消"保留内容"后源内容应被清空').not.toBe(valueBefore);

    // 撤销应把剪切还原
    await page.evaluate(async () => {
      await (window as never as { __p0: { undo: () => Promise<boolean> } }).__p0.undo();
    });
    await page.waitForTimeout(400);
    expect(await valueAt(page, 'A1'), '剪切应可撤销').toBe(valueBefore);
    await keep.check();

    // --- 2) 条目右键复制 → 表格里右键粘贴（只写内容、目标格式不变） ---
    await page.locator('[data-testid="workspace-item"]').first().click({ button: 'right' });
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
    await page.locator('[data-testid="context-menu-ws-copy"]').click();
    await page.waitForTimeout(200);

    // 粘贴目标是**一个空格**（A3）：这样"写进去了"与"什么都没发生"才区分得开
    //（若拿一个已有内容的格子当目标，恰好条目内容相同就会假红）。选区必须先铺住右键落点。
    const target = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          selectRange: (a1: string) => void;
          getDisplayValueByA1: (a1: string) => unknown;
          getCellStyleByA1: (a1: string) => unknown;
        };
      }).__p0;
      const candidates = ['A3', 'B3', 'C3', 'A5', 'C6', 'B7'];
      const a1 = candidates.find((ref) => String(hooks.getDisplayValueByA1(ref) ?? '') === '') ?? 'A3';
      hooks.selectRange(a1);
      return { a1, value: String(hooks.getDisplayValueByA1(a1) ?? ''), style: JSON.stringify(hooks.getCellStyleByA1(a1) ?? null) };
    });
    expect(target.value, '测试前提：目标格应当是空的').toBe('');
    expect(target.a1, '探针实测：A3 落在右键落点 A1:H25 内部').toMatch(/^[A-C][3-7]$/);

    await rightClickInsideSelection(page);
    await page.locator('[data-testid="context-menu-paste"]').click();
    await page.waitForTimeout(600);

    const afterPaste = await page.evaluate(
      (a1: string) => {
        const hooks = (window as never as {
          __p0: { getDisplayValueByA1: (a1: string) => unknown; getCellStyleByA1: (a1: string) => unknown };
        }).__p0;
        return {
          value: String(hooks.getDisplayValueByA1(a1) ?? ''),
          style: JSON.stringify(hooks.getCellStyleByA1(a1) ?? null),
        };
      },
      target.a1,
    );
    expect(afterPaste.value, '粘贴应把条目内容写进目标格').not.toBe('');
    expect(afterPaste.style, '粘贴不得改变目标格格式').toBe(target.style);
    const clipboardLabel = await page.evaluate(
      () =>
        ((window as never as { __p0: { log: Array<{ kind: string; detail?: { label?: string; to?: string } }> } }).__p0.log
          .filter((entry) => entry.kind === 'workspace:clipboard-paste')
          .pop()?.detail ?? null) as { label?: string; to?: string } | null,
    );
    expect(clipboardLabel?.to, '粘贴日志应记下真正写到的格子').toBe(target.a1);

    // --- 3) 右键删除条目（卡片上刻意没有 × 按钮，删除只走右键菜单） ---
    const beforeDelete = (await workspaceItems(page)).length;
    await expect(page.locator('[data-testid="workspace-remove"]'), '卡片上不应再有删除按钮').toHaveCount(0);
    await page.locator('[data-testid="workspace-item"]').first().click({ button: 'right' });
    await page.locator('[data-testid="context-menu-ws-delete"]').click();
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(beforeDelete - 1);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(beforeDelete - 1));
  });

  /**
   * 工作区新能力（本轮追加）：
   * - **右上角「清空」要二次确认**（不可逆的批量操作），确认条里能选"清空全部 / 取消"；
   * - **搜索**：按内容、按来源 A1 命中；按 **来源工作表** 筛选；「清除」与 Esc 都能复位；
   * - 筛选生效时清空还能只清"当前显示的这些"。
   */
  test('⑩ 清空二次确认 + 搜索/来源筛选 + 只清显示项', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // 取**整个已用区域**（A1:H21 → 35 个非空格子）放进工作区：这样才包含 A12 这类"靠下的格子"
    const expected = await extractViaHook(page, 'A1:H21');
    expect(expected.error).toBeUndefined();
    const total = expected.count;
    expect(total, '夹具已用区域里有 35 个非空格子').toBe(35);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);

    // ---- 1) 搜索：按内容命中 ----
    const search = page.locator('[data-testid="workspace-search"]');
    await expect(search).toBeVisible();
    await typeSearch(page, '粗体'); // E3 的内容
    await expect(page.locator('.ws-list')).toHaveAttribute('data-visible-count', '1');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(1);
    await expect(page.locator('[data-testid="workspace-item"]').first()).toHaveAttribute('title', /粗体.*来源 样式!E3/);
    await expect(page.locator('[data-testid="workspace-filter-clear"]'), '筛选生效时出现「清除」').toBeVisible();

    // ---- 2) 搜索：按来源 A1 命中（搜 A12 能定位到第 12 行的填充色格子）----
    await typeSearch(page, 'A12');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(1);
    await expect(page.locator('[data-testid="workspace-item"]').first()).toHaveAttribute('title', /来源 样式!A12/);
    // 大小写不敏感
    await typeSearch(page, 'a12');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(1);

    // ---- 3) 搜不到 → 空态提示 ----
    await typeSearch(page, '绝对不存在的文本');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(0);
    await expect(page.locator('[data-testid="workspace-no-match"]')).toBeVisible();

    // ---- 4) Esc 清空搜索 → 全部回来 ----
    await search.press('Escape');
    await expect(search).toHaveValue('');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);

    // ---- 5) 按来源筛选：这张表只有一个来源，选中它后条目数不变（证明选项与计数正确）----
    const sourceFilter = page.locator('[data-testid="workspace-source-filter"]');
    await expect(sourceFilter).toBeVisible();
    const optionText = await sourceFilter.locator('option').allTextContents();
    expect(optionText[0], '第一项是"全部来源（N）"').toBe(`全部来源（${total}）`);
    expect(optionText).toContain(`样式（${total}）`);
    await sourceFilter.selectOption('样式');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);
    await sourceFilter.selectOption('all');
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);
    await expect(page.locator('[data-testid="workspace-filter-clear"]'), '复位后「清除」消失').toHaveCount(0);

    // ---- 6) 清空要二次确认：点「清空」不能直接清光，可以取消 ----
    await page.locator('[data-testid="workspace-clear"]').click();
    const confirm = page.locator('[data-testid="workspace-clear-confirm"]');
    await expect(confirm).toBeVisible();
    await expect(confirm).toContainText(`确定清空全部 ${total} 个单元格？`);
    await expect(page.locator('[data-testid="workspace-clear-all"]')).toContainText(`清空全部 ${total} 个`);
    await expect(page.locator('[data-testid="workspace-clear-visible"]'), '无筛选时没有"只清显示项"').toHaveCount(0);
    await page.locator('[data-testid="workspace-clear-cancel"]').click();
    await expect(confirm).toHaveCount(0);
    await expect(page.locator('[data-testid="workspace-item"]'), '取消后条目一个都不能少').toHaveCount(total);

    // Esc 也能取消确认
    await page.locator('[data-testid="workspace-clear"]').click();
    await expect(confirm).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);

    // ---- 7) 筛选生效时，可以只清"当前显示的这些" ----
    await search.fill('区域'); // A2 表头
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(1);
    await page.locator('[data-testid="workspace-clear"]').click();
    await expect(confirm).toContainText(`清空显示的 1 个，还是全部 ${total} 个？`);
    await page.locator('[data-testid="workspace-clear-visible"]').click();
    await expect(confirm).toHaveCount(0);
    await expect(page.locator('[data-testid="workspace-no-match"]'), '命中的那 1 个应被清掉').toBeVisible();
    await search.fill('');
    await expect(page.locator('[data-testid="workspace-item"]'), '其余条目一个不少').toHaveCount(total - 1);
    await expect(page.locator('[data-testid="workspace-count"]')).toHaveText(String(total - 1));

    // ---- 8) 最后用「清空全部」收尾 ----
    await page.locator('[data-testid="workspace-clear"]').click();
    await page.locator('[data-testid="workspace-clear-all"]').click();
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(0);
    await expect(page.locator('[data-testid="workspace-empty"]')).toBeVisible();
  });

  /**
   * 互换后的"落点提醒"（本轮追加）：互换完成**不再留下选中**，改用**黄色**提醒框标出刚换的两格，
   * 在表格里按下指针提醒就消失（选区回到用户新点的那一格）。
   *
   * 颜色走的是 Univer 的标记图层 API（`highlightRanges`），渲染层真的会采纳自定义 style；
   * 每次运行都会把画面存到 `test-results/swap-flash.png` 供人工复核（颜色本身不做像素断言）。
   */
  test('⑪ 点击互换：取消选中并给出黄色提醒，点别处提醒消失', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    await page.locator('[data-testid="toolbar-mode-click-swap"]').click();
    await expect(page.locator('[data-testid="toolbar-mode-click-swap"]')).toHaveAttribute('aria-checked', 'true');
    await page.waitForTimeout(200);

    // 点击互换模式下单击一次即"选中一方"：这里把**实际选中的那个格子**从日志读回来，
    // 不用像素坐标硬猜（冻结列/行会让像素偏移 ≠ 内容偏移），下面的断言都用这个真实 A1。
    const box = await canvasBox(page);
    const selectedRef = async (): Promise<string> =>
      page.evaluate(
        () =>
          ((window as never as { __p0: { log: Array<{ kind: string; detail?: { label?: string } }> } }).__p0.log
            .filter((entry) => entry.kind === 'swap:selected')
            .pop()?.detail?.label ?? '') as string,
      );

    const clickFirst = async (): Promise<void> => {
      await page.mouse.click(box.x + 70, box.y + 120);
      await page.waitForTimeout(250);
    };

    await clickFirst();
    await page.waitForTimeout(100);
    const firstLabel = await selectedRef();
    expect(firstLabel, '第一次点击应选中一个单元格作为互换的一方').toMatch(/^.+![A-Z]+\d+$/);
    const firstRef = firstLabel.split('!')[1];
    expect(await valueAt(page, firstRef), '测试前提：第一个格子应当是空的').toBe('');

    // 再点另一个格子：逐个试几个落点直到**真的换成了另一方**（日志里出现 swap:cell-cell）。
    // 只把"从这次互换开始"的日志算进来：清空日志后再点，避免把上一次选边也算进来。
    await page.evaluate(() => {
      (window as never as { __p0: { log: unknown[] } }).__p0.log.length = 0;
    });
    const candidates: Array<[number, number]> = [
      [110, 245],
      [130, 270],
      [150, 280],
      [70, 135],
    ];
    let secondRef = '';
    for (const [dx, dy] of candidates) {
      await page.mouse.click(box.x + dx, box.y + dy);
      await page.waitForTimeout(250);
      const swappedRef = await page.evaluate(
        () =>
          ((window as never as { __p0: { log: Array<{ kind: string; detail?: { b?: string } }> } }).__p0.log
            .filter((entry) => entry.kind === 'swap:cell-cell')
            .pop()?.detail?.b ?? '') as string,
      );
      if (swappedRef !== '') {
        secondRef = swappedRef;
        break;
      }
      // 没换成功说明这次点到的还是同一格（或落点无效）→ 重新选中第一方再试下一个落点
      await clickFirst();
    }
    expect(secondRef, '第二次点击应选中另一方并完成互换').toMatch(/^[A-Z]+\d+$/);
    expect(secondRef).not.toBe(firstRef);

    const afterSwap = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          log: Array<{ kind: string; detail?: { a?: string; b?: string | null } }>;
          getSelectionA1: () => string | null;
        };
      }).__p0;
      const swaps = hooks.log.filter((entry) => entry.kind === 'swap:cell-cell');
      const flashes = hooks.log.filter((entry) => entry.kind === 'swap:flash');
      const kinds = hooks.log.map((entry) => entry.kind);
      return {
        swap: swaps[swaps.length - 1]?.detail ?? null,
        flash: flashes[flashes.length - 1]?.detail ?? null,
        swapped: swaps.length,
        flashCount: flashes.length,
        flashErrors: kinds.filter((kind) => kind === 'swap:flash-error').length,
        rectMismatch: kinds.filter((kind) => kind === 'swap:flash-rect-mismatch').length,
        selection: hooks.getSelectionA1(),
      };
    });

    expect(afterSwap.swapped, '两次点击应触发一次单元格互换').toBeGreaterThan(0);
    expect(
      afterSwap.flashCount,
      `互换后应记一条 swap:flash（提醒框已画出；flash-error=${afterSwap.flashErrors}）`,
    ).toBeGreaterThan(0);
    expect(afterSwap.swap?.a, '互换的是一方').toBe(firstRef);
    expect(afterSwap.swap?.b, '互换的是另一方').toBe(secondRef);
    expect(afterSwap.flash?.a, '提醒框覆盖互换的一方').toBe(firstRef);
    expect(afterSwap.flash?.b, '提醒框覆盖另一方').toBe(secondRef);
    expect(afterSwap.selection, '互换后不应还留着选中').toBeFalsy();
    expect(
      afterSwap.rectMismatch,
      '提醒框定位自检：矩形中心必须命中原单元格（坐标换算错了会在这里暴露）',
    ).toBe(0);

    // 提醒框：DOM 浮层（不是选区、不是标记图层），两个框、位置正确、颜色是黄的
    const boxes = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.swap-flash-box')).map((node) => {
        const rect = node.getBoundingClientRect();
        const style = getComputedStyle(node);
        return {
          a1: node.getAttribute('data-a1'),
          left: Math.round(rect.left),
          top: Math.round(rect.top),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
          border: style.borderTopColor,
          pointerEvents: style.pointerEvents,
        };
      }),
    );
    expect(boxes, '互换后应出现两个黄色提醒框').toHaveLength(2);
    expect(boxes.map((b) => b.a1).sort()).toEqual([firstRef, secondRef].sort());
    boxes.forEach((b) => {
      expect(b.width, `${b.a1} 的框要有宽度`).toBeGreaterThan(10);
      expect(b.height).toBeGreaterThan(5);
      // #facc15 → rgb(250, 204, 21)
      expect(b.border, `${b.a1} 的描边应是黄色，实际 ${b.border}`).toContain('250, 204, 21');
    });
    // 整层不吃鼠标事件（"不影响操作"的结构性保证）
    const layerPointerEvents = await page.evaluate(
      () => getComputedStyle(document.querySelector('.swap-flash-layer') as Element).pointerEvents,
    );
    expect(layerPointerEvents, '提醒框整层必须穿透点击').toBe('none');

    // 提醒框会**自己逐渐消失**（不需要用户点任何地方）
    await expect(page.locator('.swap-flash-box'), '淡出后提醒框应自动消失').toHaveCount(0, { timeout: 6_000 });

    // 再互换一次，专门验证"点别处立即收掉"这条路径
    await clickFirst();
    for (const [dx, dy] of candidates) {
      await page.mouse.click(box.x + dx, box.y + dy);
      await page.waitForTimeout(250);
      const again = await page.evaluate(
        () =>
          (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter((e) => e.kind === 'swap:flash')
            .length,
      );
      if (again > 1) break;
      await clickFirst();
    }
    await page.mouse.click(box.x + 230, box.y + 300);
    await page.waitForTimeout(250);
    const afterClick = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { getSelectionA1: () => string | null; log: Array<{ kind: string }> };
      }).__p0;
      return {
        selection: hooks.getSelectionA1(),
        cleared: hooks.log.filter((entry) => entry.kind === 'swap:flash-clear').length,
        boxes: document.querySelectorAll('.swap-flash-box').length,
      };
    });
    expect(afterClick.cleared, '点别处应立刻收掉提醒框（日志 swap:flash-clear）').toBeGreaterThan(0);
    expect(afterClick.boxes, '点别处后提醒框应立刻消失').toBe(0);
    expect(afterClick.selection, '点别处后选区应是新点的那一格').toBeTruthy();
    expect(afterClick.selection, '不再是刚互换的区域').not.toBe(firstRef);
  });

  /**
   * 条目卡片上**没有任何按钮**：
   * - 卡片正文只有"单元格本身"（预览），没有 × 角标、没有来源角标；
   * - 单个删除走**右键菜单**（复制 / 剪切 / 删除）；
   * - 来源信息只在 `title` 悬浮提示里（`{内容} · 来源 {工作表}!{A1}（右键可复制/剪切/删除）`）。
   */
  test('⑫ 条目卡片不带 × 按钮：删除走右键菜单，来源只在 title 里', async ({ page }) => {
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    const expected = await extractViaHook(page, 'A1:F6');
    expect(expected.error).toBeUndefined();
    const total = expected.count;
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total);

    // 卡片内部：只有预览，没有任何按钮 / 输入框
    const internals = await page.locator('[data-testid="workspace-item"]').first().evaluate((node) => ({
      buttons: node.querySelectorAll('button').length,
      inputs: node.querySelectorAll('input, select, textarea').length,
      previews: node.querySelectorAll('[data-testid="snapshot-preview"]').length,
      text: (node.textContent ?? '').trim(),
    }));
    expect(internals.buttons, '卡片上不应再有 × 按钮').toBe(0);
    expect(internals.inputs).toBe(0);
    expect(internals.previews, '每张卡片只有一个单元格预览').toBe(1);
    expect(internals.text, '正文里不再显示来源角标').not.toContain('来源');
    await expect(page.locator('[data-testid="workspace-remove"]')).toHaveCount(0);

    // 来源信息仍然可查：title 里带"内容 · 来源 工作表!A1"
    const title = await page.locator('[data-testid="workspace-item"]').first().getAttribute('title');
    expect(title).toContain('· 来源 样式!');
    expect(title).toMatch(/右键可复制\/剪切\/删除/);

    // 右键菜单是唯一入口：复制 / 剪切 / 删除都在（快捷说明文字挂在菜单项里）
    await page.locator('[data-testid="workspace-item"]').first().click({ button: 'right' });
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
    await expect(page.locator('[data-testid="context-menu-ws-copy"]')).toContainText('复制');
    await expect(page.locator('[data-testid="context-menu-ws-cut"]')).toContainText('剪切');
    await expect(page.locator('[data-testid="context-menu-ws-delete"]')).toContainText('删除');
    const removedLabel = (await workspaceItems(page))[0].a1;
    await page.locator('[data-testid="context-menu-ws-delete"]').click();
    await expect(page.locator('[data-testid="workspace-item"]')).toHaveCount(total - 1);
    expect((await workspaceItems(page)).some((item) => item.a1 === removedLabel), '右键删除应只删掉那一个').toBe(false);
    const removeLogs = await page.evaluate(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string; detail?: { id?: string } }> } }).__p0.log.filter(
          (entry) => entry.kind === 'workspace:remove-item',
        ).length,
    );
    expect(removeLogs, '删除应留下 workspace:remove-item 日志').toBeGreaterThan(0);
  });

  /**
   * 用户实测反馈："多选后复制/剪切到工作区以后，不会默认跳过空单元格了"——
   * 要求"默认跳过空单元格，并检查**所有**的转移单元格到工作区的渠道，确保**绝对**跳过空内容的单元格"。
   *
   * 根因：旧口径只判 `v === null`。但"看起来是空的格子"有三种形态：
   *  ① 用户在编辑器里把内容删光 → `v` 是**空字符串 `''`**（旧口径漏掉的就是这种）；
   *  ② 只打了空格 / 全角空格（中文表格常见）；
   *  ③ 从没写过 → `null`（旧口径能跳过）。
   * 现在统一由 `isEmptyContent` 判定（无公式 + 文本 trim 后为空；`0`/`false`/公式都算内容），
   * 且所有渠道都收敛到同一个入口 `extractCellItems`。
   */
  test('⑬ 空内容绝不放进去：null / 空字符串 / 只有空格，所有渠道一致', async ({ page }) => {
    test.setTimeout(120_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    /**
     * 在 **A6:A11**（夹具里只有 A10 被我写成内容，其余本来就是空的）造出"空"的各种形态。
     * 选可见范围是必须的：`rectOfA1` 要靠画布上的真实位置，屏幕外的格子量不到。
     */
    await page.evaluate(async () => {
      const hooks = (window as never as { __p0: { setValue: (r: number, c: number, v: string) => void } }).__p0;
      const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
      hooks.setValue(5, 0, ''); // A6 编辑器清空 → 空字符串（旧口径漏掉的就是这种）
      await wait(150);
      hooks.setValue(6, 0, ' '); // A7 半角空格
      await wait(150);
      hooks.setValue(7, 0, '\u3000'); // A8 全角空格
      await wait(150);
      hooks.setValue(9, 0, '有内容'); // A10 真内容
      await wait(400);
    });

    /**
     * 渠道①（**生产路径**）：多选 A6:A11 → 右键「复制到工作区」。
     * 落点必须在这片选区**里面**（右键会先把选区设成指针下那一格）。
     */
    await page.evaluate(() => {
      (window as never as { __p0: { selectRange: (a1: string) => void } }).__p0.selectRange('A6:A11');
    });
    await page.waitForTimeout(250);
    const inside = await page.evaluate(() =>
      (window as never as { __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1('A8'),
    );
    expect(inside, '应能定位 A8 的位置').not.toBeNull();
    if (!inside) return;
    await page.mouse.click(inside.left + inside.width / 2, inside.top + inside.height / 2, { button: 'right' });
    await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
    await expect(page.locator('[data-testid="context-menu-workspace-copy"]')).toContainText('A6:A11');
    await page.locator('[data-testid="context-menu-workspace-copy"]').click();
    await page.waitForTimeout(500);

    const skipLogs = await page.evaluate(
      () =>
        (window as never as { __p0: { log: Array<{ kind: string; detail?: { skipped?: number } }> } }).__p0.log.filter(
          (entry) => entry.kind === 'workspace:skip-empty',
        ),
    );
    expect(skipLogs.length, '应留下"跳过了 N 个空单元格"的日志').toBeGreaterThan(0);
    expect(skipLogs[skipLogs.length - 1].detail?.skipped, '6 格里 5 格是空的（含空字符串/空格/全角空格）').toBe(5);

    /** 渠道②：面板里真的只有一张卡片（用户看到的东西），没有白卡片 */
    const cards = page.locator('[data-testid="workspace-item"]');
    await expect(cards, '工作区只应有 1 张卡片').toHaveCount(1);
    const cardTexts = await cards.allInnerTexts();
    expect(cardTexts.join('\n')).toContain('有内容');
    expect(cardTexts.join('\n').trim(), '卡片里不该只有空白').not.toBe('');

    /** 渠道①b：同一个底层入口（拖拽 / 「+」导入 / 一键全搬 都走它）的精确读数 */
    const viaHook = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          snapshotSelectionToWorkspace: (a1: string) => {
            items: Array<{ source: { a1: string } }>;
            skippedEmpty: number;
          };
        };
      }).__p0;
      const out = hooks.snapshotSelectionToWorkspace('A6:A11');
      return { a1s: out.items.map((item) => item.source.a1), skippedEmpty: out.skippedEmpty };
    });
    expect(viaHook.a1s, '只有 A10 有内容').toEqual(['A10']);
    expect(viaHook.skippedEmpty, '五个空格子全部跳过').toBe(5);

    /** 渠道③：单元格 ↔ 工作区条目互换——换出来的是空格子时不能留白卡片 */
    const itemsBefore = await workspaceItems(page);
    // A5 是从没写过、也没有样式的空格子；右键落点在哪一格，菜单拿到的就是哪一格
    const rect = await page.evaluate(() =>
      (window as never as { __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1('A5'),
    );
    expect(rect, '应能定位 A5 的位置').not.toBeNull();
    if (rect) {
      await page.mouse.click(rect.left + rect.width / 2, rect.top + rect.height / 2, { button: 'right' });
      await expect(page.locator('[data-testid="context-menu"]')).toBeVisible();
      const swapItem = page.locator('[data-testid="context-menu-swap-workspace"]');
      await expect(swapItem).toBeEnabled();
      await swapItem.click();
      await page.waitForTimeout(400);

      const after = await workspaceItems(page);
      expect(
        after.some((item) => item.a1 === 'A5'),
        '换出来的是空格子 → 条目应被移除，而不是变成一张白卡片',
      ).toBe(false);
      expect(after.length, '条目数应少一张').toBe(itemsBefore.length - 1);
      expect(
        await page.evaluate(
          () =>
            (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter(
              (e) => e.kind === 'workspace:remove-empty-after-swap',
            ).length,
        ),
        '应留下"换出空格子 → 移除条目"的日志',
      ).toBeGreaterThan(0);
    }
  });

  /**
   * 用户要求（第 4 条）："点击互换模式请支持点击工作区选中后再点击表格进行互换（支持与工作区进行互换）"。
   *
   * 以前只有"先点单元格 → 再点单元格"能累积选中：工作区条目上只接了拖拽（拖回表格写回），
   * 单击既不高亮也不登记，于是"先点条目再点格子"完全没反应。
   * 现在：条目上按下→没拖动就松手 = 一次点击 → 登记为待互换的一方（并高亮成"待互换"），
   * 再点表格里尺寸相同的格子即互换；两个顺序都成立。
   */
  test('⑭ 点击互换：先点工作区条目、再点单元格也能互换（两个顺序都行）', async ({ page }) => {
    test.setTimeout(120_000);
    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    // 准备 2 张卡片（A1 / A2 都有内容，A3 是空的会被跳过）
    const seeded = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: { snapshotSelectionToWorkspace: (a1: string) => { items: Array<{ source: { a1: string } }> } };
      }).__p0;
      return hooks.snapshotSelectionToWorkspace('A1:A3').items.map((item) => item.source.a1);
    });
    expect(seeded.length, '应有 2 个条目').toBe(2);
    await page.waitForTimeout(300);

    await switchMode(page, 'click-swap');
    await page.waitForTimeout(200);

    /** 顺序一：先点工作区条目，再点单元格 */
    const tile = page.locator('[data-testid="workspace-item"]').first();
    await tile.click();
    await expect(tile, '点过的条目要标记为"待互换"').toHaveAttribute('data-pending-swap', '1');
    expect(
      await page.evaluate(
        () =>
          (window as never as { __p0: { log: Array<{ kind: string }> } }).__p0.log.filter(
            (e) => e.kind === 'swap:select-workspace-item',
          ).length,
      ),
      '应留下"把工作区条目选为一方"的日志',
    ).toBeGreaterThan(0);

    // 再点表格里一个**有内容**的格子（A12 在夹具里有内容）
    const targetRect = await page.evaluate(() =>
      (window as never as { __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1('A12'),
    );
    expect(targetRect, '应能量到 A12').not.toBeNull();
    if (targetRect) {
      const before = await valueAt(page, 'A12');
      const tileId = (await workspaceItems(page))[0].id;
      await page.mouse.click(targetRect.left + targetRect.width / 2, targetRect.top + targetRect.height / 2);
      await page.waitForTimeout(500);

      const after = await valueAt(page, 'A12');
      expect(after, '互换后单元格应拿到条目的内容').not.toBe(before);
      expect(
        await page.evaluate(
          (id: string) =>
            (window as never as { __p0: { getWorkspaceItems: () => Array<{ id: string; label: string }> } }).__p0
              .getWorkspaceItems()
              .find((item) => item.id === id)?.label ?? null,
          tileId,
        ),
        '条目还在（换回来的是 A12 原来的内容，非空）',
      ).not.toBeNull();
      await expect(page.locator('[data-testid="workspace-item"][data-pending-swap="1"]'), '互换后不应残留待互换标记').toHaveCount(0);
    }

    /** 顺序二：先点单元格，再点工作区条目 */
    const cellRect = await page.evaluate(() =>
      (window as never as { __p0: { rectOfA1: (a1: string) => { left: number; top: number; width: number; height: number } | null } }).__p0.rectOfA1('B12'),
    );
    if (cellRect) {
      const beforeAtCell = await valueAt(page, 'B12');
      await page.mouse.click(cellRect.left + cellRect.width / 2, cellRect.top + cellRect.height / 2);
      await page.waitForTimeout(300);
      const secondTile = page.locator('[data-testid="workspace-item"]').first();
      const labelBefore = (await workspaceItems(page))[0].label;
      await secondTile.click();
      await page.waitForTimeout(500);
      expect(await valueAt(page, 'B12'), '反过来点同样要互换').not.toBe(beforeAtCell);
      expect((await workspaceItems(page))[0].label, '条目仍应存在（换回来的是 B12 的内容）').not.toBe(labelBefore);
    }

    /** 同一个条目点两次 = 取消选中（状态机里"同一方"的语义） */
    const cancelTile = page.locator('[data-testid="workspace-item"]').first();
    await cancelTile.click();
    await page.waitForTimeout(250);
    await expect(cancelTile, '第一次点击应标记为待互换').toHaveAttribute('data-pending-swap', '1');
    await cancelTile.click();
    await page.waitForTimeout(250);
    await expect(page.locator('[data-testid="workspace-item"][data-pending-swap="1"]'), '再点一次应取消').toHaveCount(0);
  });
});

/** A1 记号的行优先比较（先比行号，再比列标） */
function compareA1RowMajor(a: string, b: string): number {
  const parse = (a1: string): { row: number; col: number } => {
    const match = /^([A-Z]+)(\d+)$/.exec(a1);
    if (!match) return { row: 0, col: 0 };
    const col = [...match[1]].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
    return { row: Number(match[2]) - 1, col };
  };
  const left = parse(a);
  const right = parse(b);
  return left.row - right.row || left.col - right.col;
}
