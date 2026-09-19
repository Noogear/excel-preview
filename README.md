# excel-preview — 纯前端 Excel 高保真预览与编辑工具

给班主任用的表格工具：**打开 xlsx 看到的就是 Excel 里的样子**，但**只允许改单元格内容**——
字体/边框/底色/合并/行列宽/条件格式/批注/图片等一概动不了，导出时原文件里没动过的部件**逐字节保留**。

纯前端、可离线：文件全程在浏览器里解析与导出（IndexedDB 本地存储），**不上传任何服务器**。

> 工程实现（技术选型、关键决策与理由、实测数据、已知边界、测试与部署）在 [`工程纪要.md`](./工程纪要.md)。
> 这份 README 只讲"怎么用、怎么跑、怎么发"。

## 能力概览

**导入**：`.xlsx` / `.xlsm`（宏原样保留、不执行）/ `.xltx` / `.xltm` / `.csv` / `.tsv` / `.txt`（含 GBK）/ `.ods` / `.xls`；
`.xlsb` 需本机 Excel（本地版）；`.numbers`、`.et` 明确不支持并说明原因与出路。

**保真**：边框/填充/字体/对齐/合并/行列尺寸、数字格式（含日期货币百分比与自定义）、条件格式四类、
数据验证、超链接、批注、浮动图片、Excel 表格（斑马纹/表头/汇总行）、冻结窗格、自动换行行高自适应。

**三种交互模式**（工具栏切换）：

| 模式 | 用法 |
| --- | --- |
| **选择** | 完全原生：点选 / 框选 / **Ctrl+点选多区域** |
| **拖拽** | **按住直接拖**即可搬运或与目标互换内容（全程不出现多选框） |
| **点击互换** | 依次点两格即互换；也可以**先点工作区条目、再点表格里的格子** |

任何互换/搬运完成后自动取消选中。三种模式下按住 Shift/Ctrl 都退回原生框选（用于"区域互换"这种先选一片的场景）。

**侧边栏工作区**：把选中的格子**逐格拆成一张张小卡片**暂存（含搜索、按来源表筛选、可拖宽），
需要时拖回表格任意位置或右键写回；支持「+」面板按 `A1:B2` / `3:5` / `B:D` / 多块 `A1:B2 D4:E5` 批量加入、
一键全搬"整表已用区域"、以及"选择模式下选中后 3 秒内点一下工作区空白"的快速加入。

**只读约束**：属性面板类能力（字体/边框等）**永久不提供**；自动填充（选区右下角小方块）已关闭；
粘贴外来内容只写值，目标格格式一格不变。

## 快速开始

### 方式一：一键运行（推荐，Windows）

**双击 `run.cmd`**：检查环境 → 首次自动装依赖 → 起服务 → 打开浏览器（不跑测试）。

```
run.cmd                    启动（默认端口 5273，被占用会自动顺延）
run.cmd -Port 5300         换端口
run.cmd -NoBrowser         不自动打开浏览器
run.cmd -SkipInstall       跳过依赖检查
```

关掉窗口（或 Ctrl+C）服务即结束、端口释放；窗口被强杀时也有看门狗进程兜底清理，不留后台 node。
端口上已经跑着本工具时，再次双击只会**复用**它并打开浏览器。

### 方式二：手动

```bash
npm install                 # 环境限制 npm 缓存目录时：npm install --cache .npm-cache
npm run fixtures            # 生成回归样本（合成数据，不含任何真实学生信息）
npm run dev                 # http://localhost:5273
```

页面点「打开表格」导入 `fixtures/` 下任意样本即可；右侧「事件日志」实时显示 Univer 事件、
格式锁拦截与导入报告（含"由 X 格式转换导入"的说明与降级项）。

### 双形态

| | **本地版**（默认） | **静态版**（部署用） |
| --- | --- | --- |
| 怎么起 | `npm run dev` / `run.cmd`；打包 `npm run build:local` | `npm run build:static` |
| 产物 | `dist/`，`base: '/'` | `dist/`，`base: './'`（可挂子路径） |
| 本机 Excel 桥 | ✅ `.ods/.xls/.xlsb` 导出与 `.xlsb` 导入（由 Excel 自己另存，样式/公式/条件格式/批注/图片都在） | ❌ 没有本机进程：这几项在菜单里**置灰并写明"本地版可用"** |
| 数据 | 全程本机（IndexedDB） | 全程在浏览器，**任何文件都不上传** |

### 本地运行要点

- **浏览器**：Chromium 内核（Chrome / Edge 105+），用的是 Canvas 2D 与 IndexedDB。
- **端口**：默认 `5273`（`strictPort`）；改端口用 `npm run dev -- --port 5274`。
- **生产构建**：`npm run build:local` / `npm run build:static`（先 `tsc --noEmit` 再 `vite build`），
  产物是纯静态文件；本地看效果用 `npm run serve:dist`（刻意挂在 `http://127.0.0.1:5399/repo/` 子路径）。
- **端到端测试**用**系统已装的 Chrome**（不下载 Chromium），Playwright 会自动起/复用 5273 的 dev 服务。
- **大文件**：性能夹具默认 5 万行 × 20 列（100 万格），可选，不影响功能验证。

## 部署（EdgeOne Pages）

部署用**静态版**：`npm run build:static`，把 `dist/` 传到 EdgeOne Pages 即可（`base: './'`，子路径也能用）。

- **GitHub Pages 部署已停用**：`.github/workflows/static-pages.yml` 已删除，推到 `main` 不会再触发发布。
  如需彻底关闭，到 GitHub 仓库 `Settings → Pages` 把 Source 设为 `None`（已发布的那份站点会随之下线）。
- 静态形态没有本机进程，所以 `.ods/.xls/.xlsb` 导出与 `.xlsb` 导入是置灰的（界面会写明"本地版可用"）。

### 首次慢、后续快吗？—— 是，机制如下

| | 说明 |
| --- | --- |
| **首次访问** | 要下全部产物 **≈6.8 MB**，其中 Univer 引擎 `vendor-univer` 一块就 **5.3 MB**（gzip ≈1.4 MB）。这是最慢的一次 |
| **后续访问** | 产物都是**内容哈希命名**（`App-<hash>.js`、`vendor-univer-<hash>.js`）。内容没变时浏览器直接用缓存，**一个字节都不重下**；只有 `index.html` 需要校验（没变就只回一个 304） |
| 缓存过期之后 | 靠 `ETag` 走 304 协商，仍然**不会**重下那 5 MB（实测 GitHub Pages 对 `index.html` 与 `/assets/*` 都给 `max-age=600`，即 10 分钟内零请求） |
| EdgeOne 上建议 | 在缓存规则里把 `/assets/*` 设成**长 TTL**（哈希命名的文件可以 `immutable`），`index.html` 设**短 TTL 或不缓存** —— 这样发新版能立刻生效，而 5 MB 的引擎只在首次或版本更新时下 |
| 应用自身的"快" | 上次打开的标签、编辑与工作区都存在 IndexedDB 里，**重开页面会直接恢复现场**，不用重新选文件 |
| 注意 | 没有 Service Worker，所以**离线打不开**；强制刷新（Ctrl+F5）会把所有资源重下一次。解析/导出是每次操作都要做的，跟缓存无关 |

## 测试

```bash
npm run typecheck      # tsc --noEmit（零报错）
npm test               # Vitest 单元 467 例（解析器 / 适配 / 导出 / 几何 / 模型 / 只读门禁 …）
npm run e2e            # Playwright 端到端 131 例，三段独立进程：main(112) → sensitive(17) → hot-reload(2)
npm run e2e:fast       # 只跑主项目 112 例（并发）
npm run e2e:sensitive  # 只跑串行 17 例
npm run e2e:hot-reload # 只跑热更新 2 例（必须是独立进程）
npm run bridge:clean   # 收拾"转换桥"可能残留的无窗口 Excel（-DryRun 只看不杀）
```

| 命令 | 耗时（本机 16 核 / 23 GB） |
| --- | --- |
| `npm run typecheck` | ~8 s |
| `npm test` | ~15–30 s |
| `npm run e2e` | ~3–4 分钟 |
| `npm run e2e:fast` / `:sensitive` / `:hot-reload` | ~2 分钟 / ~1.5 分钟 / ~20 s |

**为什么分三段**（并发能把 8 分钟压到 3 分钟，但这几类不能和别人抢机器，否则会**测出假失败**）：

- `sensitive`（串行）：`perf`（帧率断言）、`tabs-memory`（采堆）、`static-form`（启动预算）、
  `date-cells`（悬停时序）、`clipboard`（读写**真实系统剪贴板**，机器级共享资源）。
- `hot-reload`（**独立进程**）：它会改写 `src/App.tsx` 触发真实 HMR；与 `perf` 同进程时用例全过但
  Playwright **永不退出**，所以单独一次调用。

依赖机器的用例会**自动跳过**而不是假红：转换桥相关（CI 上没有 Excel）、真实用户样本
（用 `USER_EXCEL` 环境变量或 `tests/e2e/user-fixture.local` 指定）。

想临时调并发：`PW_WORKERS=8 npm run e2e`（6 与 8 耗时几乎一样，6 更稳）。
只跑一两个文件：`node node_modules/@playwright/test/cli.js test tests/e2e/workspace.spec.ts --project=main --workers=6`。

**性能基准**（夹具与产物不进仓库）：

```bash
npm run bench:fixture  # 生成 fixtures/bench-large.xlsx（5 万行 × 20 列 ≈ 6 MB）
npm run bench:parse    # Node 侧管线基准：解析/适配/吞吐/堆内存
```

自定义规模：`$env:BENCH_ROWS='100000'; $env:BENCH_COLS='30'; npm run bench:fixture`；
换夹具跑端到端性能：`$env:PERF_FIXTURE='bench-medium.xlsx'; npm run e2e -- tests/e2e/perf.spec.ts`。

## 目录结构

```
src/
  App.tsx          主组件：引导 / 多标签 / 导入导出 / 工作区 / 交互装配（全项目最大文件）
  parser/          自研解析：OOXML（工作表/样式/条件格式/数据验证/批注/表格/图片/打印设置）、
                   CSV/TSV/文本（GBK）、ODS、XLS（BIFF8）
  importer/        → Univer：数据适配、特性应用（CF/DV/链接/批注/图片）、翻译成规范 xlsx
  exporter/        外科式回写 xlsx / CSV 导出 / 瘦身导出源
  univer/          引导、只读闸门、样式命令锁、脏格账本、互换命令、填充柄关闭
  interaction/     三种模式的手势与动画（拖拽/点击互换/选区/滚动条）
  workspace/       侧边栏工作区（逐格快照、导入面板、预览渲染）
  persistence/     IndexedDB 会话（自动保存 + 误关闭恢复）
  shell/           工具栏 / 标签条 / 右键菜单 / 历史面板 / 主题
tests/unit（纯逻辑）  tests/e2e（真实浏览器，分三段）  tools/（夹具、基准、启动器、Excel 桥）
fixtures/          脚本合成的样本（**不含任何真实学生数据**）
```

模块职责与内部细节见 [`工程纪要.md`](./工程纪要.md)。

## 维护者注意

- `.cmd` 必须保持**纯 ASCII**（cmd.exe 按 OEM 代码页读它，中文会被误解析成命令）；
  `tools/*.ps1` 必须存成 **UTF-8 with BOM**（否则 Windows PowerShell 5.1 读中文乱码）。
- 改样式/结构相关命令的白名单时注意：Univer 命令的 `name` 与 `id` **不是一回事**
  （`sheet.command.copy` 只是 name，真 id 是 `univer.command.copy`），写错会"空放行 + 真命令被拦"。
- 视觉基线在 `tests/e2e/*-snapshots/`，布局改动后用 `--update-snapshots` 重建。
