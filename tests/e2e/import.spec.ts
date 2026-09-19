/**
 * 导入与保真。
 *
 * 链路：真实 xlsx（tools/make-fixtures.mjs 生成）→ 自研 OOXML 解析器 → 适配层
 *       → Univer IWorkbookData → canvas 渲染。
 *
 * 保留的 6 条都是"导入不可退让的契约"：
 *   ① 引导 + 导入全程**不发起任何外部网络请求**，且关键样式真的进了数据模型
 *   ② 数字格式的显示语义与 Excel 一致
 *   ③ Excel 表格（ListObject）的斑马纹/表头/汇总行被实体化到单元格
 *   ④ 自动换行的长文本行被加高（否则文本被裁切）
 *   ⑤ 条件格式 4 类 + 数据验证 5 类真的进入工作表模型（不只是计数）
 *   ⑥ 超链接 / 批注 / 图片真的进入工作表模型
 *
 * 已删除：多工作表表名顺序、fixture-table 的特性层"无事可做"、条件格式截图单独用例
 * （前两条属于解析器细节，已由 tests/unit 覆盖；截图并入 ⑤ 内）。
 */
import { expect, test, type Page } from '@playwright/test';
import { existsSync } from 'node:fs';
import { basename } from 'node:path';

import { fixturePath, importFixture, sheetCanvas, userFixturePath, waitForBoot } from './helpers';

interface FacadeRange {
  getHyperLinks(): Array<{ row: number; column: number; url: string; label: string }>;
}

interface FacadeSheet {
  getSheetName(): string;
  getConditionalFormattingRules(): Array<{
    cfId: string;
    ranges: Array<{ startRow: number; startColumn: number; endRow: number; endColumn: number }>;
    rule: { type?: string; style?: unknown; config?: unknown };
  }>;
  getDataValidations(): Array<{
    getCriteriaType(): string;
    getCriteriaValues(): [string | undefined, string | undefined, string | undefined];
    rule: { type: string; formula1?: string; formula2?: string; error?: string };
  }>;
  getNotes(): Array<{ id: string; row: number; col: number; note: string }>;
  getImages(): Array<{ getId(): string }>;
  getRange(ref: string): FacadeRange;
}

/** 主画布画完才截图：canvas 渲染没有 DOM 可等，只能等画布可见 + 让出几帧 */
async function settleCanvas(page: Page): Promise<void> {
  await sheetCanvas(page).waitFor({ state: 'visible' });
  await page.waitForTimeout(600);
}

test.describe('导入与保真', () => {
  test('① 导入真实 xlsx：渲染成功、全程零外部网络请求、关键样式进入数据模型', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures 生成样本');

    const external: string[] = [];
    page.on('request', (request) => {
      const url = request.url();
      if (!/^https?:\/\/(127\.0\.0\.1|localhost)/.test(url) && !url.startsWith('data:') && !url.startsWith('blob:')) {
        external.push(url);
      }
    });
    const consoleErrors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') consoleErrors.push(message.text());
    });

    await waitForBoot(page);

    // 泄漏检测：必须恰好存在一个"工作表主画布"实例（StrictMode 双挂载会变成 0 个或 2 个）
    await expect(sheetCanvas(page)).toBeVisible();
    await expect(page.locator('#univer-container canvas[id^="univer-sheet-main-canvas"]')).toHaveCount(1, {
      timeout: 15_000,
    });

    const summary = await importFixture(page, 'fixture-styles.xlsx');
    expect(summary.sheets).toContain('样式');

    const probe = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getDisplayValueByA1: (a1: string) => string | null;
          getCellStyleByA1: (a1: string) => Record<string, unknown> | null;
        };
      }).__p0;
      return {
        title: hooks.getDisplayValueByA1('A1'),
        header: hooks.getDisplayValueByA1('A2'),
        headerStyle: hooks.getCellStyleByA1('A2'),
        borderCellStyle: hooks.getCellStyleByA1('F3'),
        wrapCellStyle: hooks.getCellStyleByA1('D3'),
      };
    });

    expect(probe.title).toContain('样式保真样本');
    expect(probe.header).toBe('区域');
    // 表头：粗体 + 蓝底 + 白字（说明样式真的进了单元格模型）
    expect(Number(probe.headerStyle?.bl), '表头应加粗').toBe(1);
    expect(JSON.stringify(probe.headerStyle?.bg), '表头应有强调色底').toContain('2563EB');
    expect(JSON.stringify(probe.headerStyle?.cl), '表头应有可读字色').toContain('FFFFFF');
    // 边框 / 自动换行
    expect(JSON.stringify(probe.borderCellStyle?.bd)).toContain('t');
    expect(Number(probe.wrapCellStyle?.tb), '自动换行标记应落地').toBeGreaterThan(0);

    expect(consoleErrors, `控制台错误：${consoleErrors.join(' | ')}`).toEqual([]);
    expect(external, `导入过程发起了外部网络请求：${external.join(' | ')}`).toEqual([]);

    await page.screenshot({ path: 'test-results/import-styles.png' });
  });

  test('② 数字格式样本：显示文本与 Excel 语义一致', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-numfmt.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-numfmt.xlsx');

    const displays = await page.evaluate(() => {
      const hooks = (window as never as { __p0: { getDisplayValueByA1: (a1: string) => string | null } }).__p0;
      const out: Record<string, string | null> = {};
      for (const a1 of ['B2', 'B4', 'B6', 'B10', 'B13']) out[a1] = hooks.getDisplayValueByA1(a1);
      return out;
    });

    // 不做逐字面量断言（Excel 本地化有差异），只锁"格式确实生效"
    expect(displays.B2, '#,##0.00 应显示千分位与两位小数').toMatch(/1[,，]?234\.57/);
    expect(displays.B4, '0.0% 应显示百分号').toContain('%');
    expect(displays.B6, 'yyyy-mm-dd 应显示为日期').toMatch(/2025|2024/);
    expect(displays.B10, '0.00E+00 应显示科学计数').toMatch(/E\+0/i);
    expect(displays.B13, '#,##0 应显示千分位整数').toMatch(/12[,，]?345/);
  });

  test('③ Excel 表格：表头强调色、斑马纹交替、汇总行加粗带上边框', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-table.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-table.xlsx');

    const probe = await page.evaluate(() => {
      const hooks = (window as never as {
        __p0: {
          getCellStyleByA1: (a1: string) => Record<string, unknown> | null;
          getDisplayValueByA1: (a1: string) => string | null;
        };
      }).__p0;
      const read = (a1: string) => {
        const style = hooks.getCellStyleByA1(a1) ?? {};
        const border = style.bd as Record<string, { s?: number }> | undefined;
        return {
          a1,
          text: hooks.getDisplayValueByA1(a1),
          background: (style.bg as { rgb?: string } | undefined)?.rgb ?? null,
          bold: Number(style.bl ?? 0) === 1,
          topBorder: border?.t?.s ?? 0,
          color: (style.cl as { rgb?: string } | undefined)?.rgb ?? null,
        };
      };
      // 表格 1：A1:D5（表头 + 4 行体行，斑马纹）；表格 2：A8:C11（带汇总行）
      return { header: read('A1'), body1: read('A2'), body2: read('A3'), totals: read('B11') };
    });

    expect(probe.header.background, '表头必须有强调色底（说明表格样式已实体化）').toBeTruthy();
    expect(probe.header.bold, '表头加粗').toBe(true);
    expect(probe.header.color, '表头文字颜色必须被显式设置，避免浅底浅字').toBeTruthy();
    expect(probe.body1.background, '斑马纹：体行 1 与体行 2 的底色必须不同').not.toBe(probe.body2.background);
    expect(probe.totals.bold, '汇总行加粗').toBe(true);
    expect(probe.totals.topBorder, '汇总行有上边框').toBeGreaterThan(0);

    await page.screenshot({ path: 'test-results/import-table.png' });
  });

  test('④ 自动换行的长文本行被加高，普通行不受影响', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-styles.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-styles.xlsx');

    const heights = await page.evaluate(() => {
      const sheet = (window as never as {
        __p0: { getActiveSheet: () => { getRowHeight: (r: number) => number } | null };
      }).__p0.getActiveSheet();
      if (!sheet) return null;
      // 样本：D3 是"自动换行的长文本"（第 3 行，0-based 2）
      return { row3: sheet.getRowHeight(2), row4: sheet.getRowHeight(3) };
    });

    expect(heights, '必须能读到行高').not.toBeNull();
    if (!heights) return;
    expect(heights.row3, '自动换行的长文本行必须比普通行更高').toBeGreaterThan(heights.row4);
    expect(heights.row3).toBeGreaterThan(24);
  });

  test('⑤ fixture-rules：4 条条件格式 + 5 条数据验证真的进了工作表模型', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-rules.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-rules.xlsx', { features: true });

    const report = await page.evaluate(() =>
      (window as never as { __p0: { getFeatureReport: () => { counts: Record<string, number>; issues: string[] } | null } }).__p0.getFeatureReport(),
    );
    expect(report).not.toBeNull();
    if (!report) return;

    expect(report.counts.conditionalFormats).toBe(4);
    expect(report.counts.dataValidations).toBe(5);
    // P0 阶段的"尚未接入"桩必须彻底消失
    expect(report.issues.join(' | ')).not.toContain('尚未接入');

    // ---- 条件格式：Facade 直读（映射对照表的可执行版本）----
    const cf = await page.evaluate(() => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => FacadeSheet | null } }).__p0.getActiveSheet();
      if (!sheet) throw new Error('没有活动工作表');
      return {
        sheetName: sheet.getSheetName(),
        rules: sheet.getConditionalFormattingRules().map((item) => ({
          type: item.rule?.type ?? '',
          ranges: item.ranges.map((r) => `${r.startRow},${r.startColumn}-${r.endRow},${r.endColumn}`),
          config: JSON.stringify(item.rule?.config ?? null),
          style: JSON.stringify(item.rule?.style ?? null),
        })),
      };
    });

    expect(cf.sheetName).toBe('条件格式');
    expect(cf.rules.map((r) => r.type)).toEqual(
      expect.arrayContaining(['highlightCell', 'colorScale', 'dataBar', 'iconSet']),
    );

    const highlight = cf.rules.find((r) => r.type === 'highlightCell');
    expect(highlight?.ranges).toEqual(['1,1-5,1']); // B2:B6
    // dxf：fill=#FFC7CE、字体色 #9C0006、加粗（univer 的 ColorKit 落库形态是 rgb(r,g,b)）
    expect(highlight?.style ?? '').toContain('rgb(255,199,206)');
    expect(highlight?.style ?? '').toContain('rgb(156,0,6)');
    expect(highlight?.style ?? '').toContain('"bl":1');

    const colorScale = cf.rules.find((r) => r.type === 'colorScale');
    expect(colorScale?.ranges).toEqual(['1,2-5,2']); // C2:C6
    expect(colorScale?.config ?? '').toContain('F8696B');
    expect(colorScale?.config ?? '').toContain('63BE7B');

    const dataBar = cf.rules.find((r) => r.type === 'dataBar');
    expect(dataBar?.ranges).toEqual(['1,3-5,3']); // D2:D6
    expect(dataBar?.config ?? '').toContain('638EC6');

    const iconSet = cf.rules.find((r) => r.type === 'iconSet');
    expect(iconSet?.ranges).toEqual(['1,4-5,4']); // E2:E6
    expect(iconSet?.config ?? '').toContain('3TrafficLights1');
    expect(iconSet?.config ?? '').toContain('"value":67'); // percent 阈值原样进入

    // 四类规则都在这一屏内（视觉证据）
    await settleCanvas(page);
    await page.screenshot({ path: 'test-results/import-rules.png' });

    // ---- 数据验证：切到第二张表后 Facade 直读 ----
    await page.evaluate(() => {
      (window as never as { __p0: { activateSheet: (name: string) => void } }).__p0.activateSheet('数据验证');
    });
    await page.waitForFunction(
      () =>
        (window as never as { __p0: { getActiveSheet: () => { getSheetName: () => string } | null } }).__p0
          .getActiveSheet()
          ?.getSheetName() === '数据验证',
      null,
      { timeout: 15_000 },
    );

    const dv = await page.evaluate(() => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => FacadeSheet | null } }).__p0.getActiveSheet();
      if (!sheet) throw new Error('没有活动工作表');
      return {
        sheetName: sheet.getSheetName(),
        rules: sheet.getDataValidations().map((rule) => ({
          type: rule.getCriteriaType(),
          formula1: rule.rule.formula1,
          error: rule.rule.error,
        })),
      };
    });

    expect(dv.sheetName).toBe('数据验证');
    expect(dv.rules.length).toBe(5);
    const dvTypes = dv.rules.map((r) => r.type);
    for (const expected of ['list', 'whole', 'decimal', 'date', 'textLength']) {
      expect(dvTypes, `数据验证类型应包含 ${expected}：${dvTypes.join(',')}`).toContain(expected);
    }
    // 列表：OOXML 的 `"男,女,未知"` 剥引号后按逗号切分，Univer 侧存成 JSON 数组字符串
    expect(dv.rules.find((r) => r.type === 'list')?.formula1).toBe('["男","女","未知"]');
    // 错误提示文案也要落地
    expect(dv.rules.find((r) => r.type === 'whole')?.error).toBe('请输入1~120');
  });

  test('⑥ fixture-extras：3 条超链接 + 2 条批注 + 2 张图片进入工作表模型', async ({ page }) => {
    test.skip(!existsSync(fixturePath('fixture-extras.xlsx')), '请先运行 npm run fixtures 生成样本');

    await waitForBoot(page);
    await importFixture(page, 'fixture-extras.xlsx', { features: true });

    const report = await page.evaluate(() =>
      (window as never as { __p0: { getFeatureReport: () => { counts: Record<string, number> } | null } }).__p0.getFeatureReport(),
    );
    expect(report).not.toBeNull();
    if (!report) return;
    expect(report.counts.hyperlinks).toBe(3);
    expect(report.counts.notes).toBe(2);
    expect(report.counts.images).toBe(2);

    const probe = await page.evaluate(() => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => FacadeSheet | null } }).__p0.getActiveSheet();
      if (!sheet) throw new Error('没有活动工作表');
      return {
        sheetName: sheet.getSheetName(),
        links: ['B1', 'B2', 'B3'].flatMap((a1) =>
          sheet.getRange(a1).getHyperLinks().map((link) => ({ a1, url: link.url })),
        ),
        notes: sheet.getNotes().map((note) => ({ row: note.row, col: note.col, note: note.note })),
        imageIds: sheet.getImages().map((image) => image.getId()),
      };
    });

    // 外部 URL 超链接
    expect(probe.links.find((l) => l.a1 === 'B1')?.url).toBe('https://example.com/path?q=1');
    // B2 是表内跳转（location="#条件格式!A1"）：目标表不存在时按"保留原始 location"降级，但链接不能丢
    expect(probe.links.find((l) => l.a1 === 'B2')?.url ?? '').toMatch(/gid=|条件格式!A1/);
    // B3 是 mailto
    expect(probe.links.find((l) => l.a1 === 'B3')?.url ?? '').toContain('mailto:test@example.com');

    expect(probe.notes).toHaveLength(2);
    expect(probe.notes.map((n) => n.note).join('\n')).toContain('批注');
    // 批注文本里的换行必须原样保留
    expect(probe.notes.some((n) => n.note.includes('\n'))).toBe(true);

    expect(probe.imageIds).toHaveLength(2);

    await settleCanvas(page);
    await page.screenshot({ path: 'test-results/import-extras.png' });
  });

  /**
   * 用户实测反馈：导入 `2502班座位表 单座 2026.9.7.xlsx` 时状态栏出现
   * 「未支持 1 项　降级 1 项」，看着像文件有什么毛病。
   *
   * 查出来的两条其实都**不是缺陷**：
   *  ① "未支持 1 项" = `[Sheet1] 打印设置(pageSetup/pageMargins/headerFooter) · 2 处`——
   *     文件里就是 `<pageMargins left="0.75" …/>` 与一个空的 `<headerFooter/>`，只影响打印，
   *     屏幕预览与编辑完全不受影响，导出又原样保留字节；
   *  ② "降级 1 项" = `patternFill patternType="gray125" 暂不支持`——这是 Excel/WPS 写进
   *     `fills[1]` 的**规范占位填充**，全表 25 个 cellXfs 没有一个引用它（实测 fillId 全是 0）。
   *
   * 现在：① 打印设置被解析成真值并归入"已原样保留（不影响预览）"；② 只有真被单元格引用的
   * 填充才算降级。于是这份文件的状态栏不再报警，抽屉里仍能查到打印设置的真值。
   */
  test('⑦ 用户座位表：状态栏不再误报"未支持/降级"，打印设置改列在"已原样保留"里', async ({ page }) => {
    test.setTimeout(240_000);
    const userFile = userFixturePath('座位表');
    test.skip(!userFile || !existsSync(userFile), '未配置用户文件（USER_EXCEL 或 tests/e2e/user-fixture.local），跳过');
    const userFileName = basename(userFile!);

    await page.setViewportSize({ width: 1440, height: 900 });
    await waitForBoot(page);
    await page.setInputFiles('[data-testid="file-input"]', userFile!);
    await page.waitForFunction(
      (expected: string) =>
        (window as never as { __p0: { getImportSummary: () => { fileName?: string } | null } }).__p0.getImportSummary()?.fileName ===
        expected,
      userFileName,
      { timeout: 150_000 },
    );
    await page.waitForTimeout(800);

    const summary = await page.evaluate(() =>
      (window as never as { __p0: { getImportSummary: () => { unsupported: string[]; warnings: string[]; preserved: string[] } | null } }).__p0.getImportSummary(),
    );
    expect(summary).not.toBeNull();
    if (!summary) return;

    // ① 打印设置不再算"未支持"，而是"已原样保留"（并说明不影响预览）
    expect(summary.unsupported.join('\n'), '打印设置不该再出现在未支持里').not.toContain('打印设置');
    expect(summary.preserved.join('\n'), '应记进"已原样保留"').toContain('打印设置');
    expect(summary.preserved.join('\n'), '要说明不影响预览').toContain('不影响预览');
    expect(summary.preserved.join('\n'), '页边距真值应被解析出来').toContain('页边距');

    // ② gray125 是未被引用的占位填充，不该产生"降级"
    expect(summary.warnings.join('\n'), 'gray125 占位填充不该报降级').not.toContain('gray125');

    // ③ 状态栏：这一份文件不该出现任何"未支持 / 降级"警告胶囊
    await expect(page.locator('[data-testid="unsupported-count"]'), '状态栏不该再显示"未支持 N 项"').toHaveCount(0);
    const statusText = (await page.locator('.statusbar').innerText()).replace(/\s+/g, ' ');
    expect(statusText, '状态栏里不应出现"降级"字样').not.toContain('降级');

    // ④ 但信息不能丢：打开日志抽屉能看到打印设置的原值与出处
    await page.locator('[data-testid="toolbar-log"]').click();
    await expect(page.locator('[data-testid="preserved-list"]')).toBeVisible();
    await expect(page.locator('[data-testid="preserved-list"]')).toContainText('打印设置');
    await expect(page.locator('[data-testid="preserved-list"]')).toContainText('不影响预览');

    // ⑤ 内容照常渲染（座位表标题与姓名都在）
    const filled = await page.evaluate(() => {
      const sheet = (window as never as { __p0: { getActiveSheet: () => { getRange: (a1: string) => { getDisplayValues: () => string[][] } } } }).__p0.getActiveSheet();
      return sheet.getRange('C1:N22').getDisplayValues().flat().filter((v) => String(v ?? '').trim() !== '').length;
    });
    expect(filled, '座位表内容应正常渲染').toBeGreaterThan(20);
  });
});
