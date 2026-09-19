import { defineConfig, devices } from '@playwright/test';

/**
 * Univer 是 canvas 渲染，断言分两层：
 * 1) 通过 window.__p0（测试钩子）读取 Facade 数据 / 事件日志
 * 2) 截图对比（toHaveScreenshot）做视觉回归
 *
 * ── 两个项目：为什么不能全都并发（用户要求"缩短测试时间但不影响测试效果"）──
 *
 * `main`（并发）：绝大多数用例。每个 test 自带独立浏览器上下文，彼此不共享状态，可以放心并发。
 * `sensitive`（串行）：少数**会互相干扰**或用例本身在测"机器负载下的时序"的用例：
 *   · `hot-reload.spec.ts` 会**改写 src/App.tsx** 触发真实热更新 —— 开发服务器会把这次更新广播给
 *     **所有**已连接的页面，并跑时会把别的 worker 的页面一起重启（测出来的失败不是真失败）；
 *   · `perf.spec.ts` 断言 p95 帧间隔 < 200ms、`tabs-memory.spec.ts` 采堆 —— 并跑时数字会被别的 canvas 抢走；
 *   · `static-form.spec.ts` 有自己的静态服务器（不受热更新影响），但它测的是**启动时间预算**；
 *   · `date-cells.spec.ts` ④ 依赖"悬停 → 上游弹提醒"这条**时序**（实测：它自己一个文件并发 6 也能过，
 *     但和别的 spec 抢机器时悬停事件就一直不触发，属于"测的是时序"的用例）。
 *
 * 组合方式（`npm run e2e` = 全量一遍）：
 *   playwright test --project=main --workers=6  &&  playwright test --project=sensitive --workers=1
 * 直接 `playwright test`（不带 --project）则是**保守路径**：顶层 workers=1，先 main 后 sensitive，慢但绝对稳。
 */
/** 串行项目里的用例（会互相干扰，或本身在测"负载下的时序"） */
const SENSITIVE = [
  '**/perf.spec.ts',
  '**/tabs-memory.spec.ts',
  '**/static-form.spec.ts',
  '**/date-cells.spec.ts',
  /**
   * 剪贴板用例读写的是**真实系统剪贴板**（机器级共享资源）：
   * 并发跑时别人的复制会盖掉它的哨兵值，测出来的失败不是真失败。
   */
  '**/clipboard.spec.ts',
];

/**
 * **单独一个进程**跑的用例（不是"串行"能解决的）。
 *
 * 为什么必须独立进程（实测查出来的）：`hot-reload.spec.ts` 会**改写 `src/App.tsx`** 触发真实热更新，
 * 它和 `perf.spec.ts`（百万格）**同处一个 playwright 进程**时，用例全部通过、但**进程永远不退出**
 * （实测：各自单跑都秒退；`hot-reload + perf` 一起跑就挂住；`clipboard + date-cells` 这种组合正常）。
 * 二分过程见 `P3-内容锁定与工作区-交付说明.md` 的「修复 AA」。
 *
 * 所以这里是**三次独立调用**（`npm run e2e` = main → sensitive → hot-reload，各自一个新进程），
 * 而不是把它塞进某个项目里 —— 换项目也还在同一个进程里，解决不了。
 */
const HOT_RELOAD = ['**/hot-reload.spec.ts'];

export default defineConfig({
  testDir: './tests/e2e',
  /**
   * 注意：Playwright 每次运行会**清空** outputDir，因此两个进程同时跑会互相删掉对方的 trace/截图
   * （表现为 `browserContext.close: ENOENT ... traces/...`）。
   * 需要并发跑时用 `PW_OUTPUT_DIR=test-results-b pnpm e2e` 指定不同目录即可。
   */
  outputDir: process.env.PW_OUTPUT_DIR ?? 'test-results',
  fullyParallel: false,
  /** 保守默认值：不带 --project 直接跑时是"一条一条来"，绝不互相干扰；快跑见 npm run e2e */
  workers: Number(process.env.PW_WORKERS ?? 1),
  reporter: [['list']],
  timeout: 60_000,
  expect: {
    timeout: 15_000,
    toHaveScreenshot: { maxDiffPixelRatio: 0.02 },
  },
  use: {
    baseURL: 'http://127.0.0.1:5273',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    viewport: { width: 1440, height: 900 },
  },
  projects: [
    {
      name: 'main',
      // 文件内也并发：每个 test 独立上下文，互不共享状态
      fullyParallel: true,
      // 串行用例与"改写源码"的热更新用例都不进 main
      testIgnore: [...SENSITIVE, ...HOT_RELOAD],
      // 直接用系统已安装的 Chrome，避免下载 Playwright 自带 Chromium（内网/受限网络下更快）
      use: { ...devices['Desktop Chrome'], channel: 'chrome' },
    },
    {
      name: 'sensitive',
      fullyParallel: false,
      testMatch: SENSITIVE,
      use: { ...devices['Desktop Chrome'], channel: 'chrome' },
    },
    {
      /** 单独一个进程跑（见 HOT_RELOAD 的说明）：`npm run e2e:hot-reload` */
      name: 'hot-reload',
      fullyParallel: false,
      testMatch: HOT_RELOAD,
      use: { ...devices['Desktop Chrome'], channel: 'chrome' },
    },
  ],
  webServer: {
    command: 'npx vite --port 5273 --strictPort',
    url: 'http://127.0.0.1:5273',
    reuseExistingServer: true,
    timeout: 180_000,
  },
});
