# excel-preview — 纯前端 Excel 高保真预览与编辑工具

> **在线版（静态托管，打开即用）**：<https://noogear.github.io/excel-preview/> ——
> 纯前端、无需安装；文件都在你自己的浏览器里解析，**不上传任何服务器**。
> 在线版为了免安装牺牲了"经本机 Excel 另存为 .ods/.xls/.xlsb"这类能力（菜单里会置灰并说明）；
> 需要它们就用本地版：`run.cmd`（或 `npm run dev`）。
>
> **形态**：同一套代码出两份产物 —— **本地版**（`run.cmd` / `npm run dev`，带"本机 Excel 转换桥"：
> `.ods/.xls/.xlsb` 导出与 `.xlsb` 导入都完整保真）与 **静态版**（`npm run build:static`，可直接托管到
> GitHub / Gitee / EdgeOne Pages，桥相关入口自动置灰并说明"本地版可用"）。详见下文「双形态」。
>
> 当前进度：**P0 技术验证 ✅ · P2 交互与 UI ✅ · P1 保真与导出 ✅ · P3 内容锁定/工作区/历史/持久化/性能 ✅ · P5 多格式导入导出（第一刀）✅**
> - P0 结论：[`p0-findings.md`](./p0-findings.md)
> - P2 交付说明：[`P2-交互与UI-交付说明.md`](./P2-交互与UI-交付说明.md)
> - P1 交付说明：[`P1-保真与导出-交付说明.md`](./P1-保真与导出-交付说明.md)
> - P3 交付说明（10 项需求逐条 + 性能实测 + 已知边界）：[`P3-内容锁定与工作区-交付说明.md`](./P3-内容锁定与工作区-交付说明.md)
> - P5 方案（多格式导入导出 / 库选型 / 双形态 / 静态托管）：[`P5-导入导出多格式与库选型-方案.md`](./P5-导入导出多格式与库选型-方案.md)
> - 需求可行性与方案（仓库外）：上级目录的 `Excel预览工具-可行性评估.md`
>
> 已可用的能力：导入 xlsx / xlsm / ods / xls / csv（宏原样保留、不执行；后三种自动转换为 xlsx）并渲染（边框/填充/字体/对齐/合并/行列尺寸/数字格式/冻结/公式缓存值、
> **条件格式四类 / 数据验证 / 超链接 / 批注 / 浮动图片 / Excel 表格斑马纹**、自动换行行高自适应）、
> 只改内容不改格式的编辑锁定、**三种交互模式（选择 / 拖拽 / 点击互换，工具栏图标切换）**、
> **不连续多区域选区（Ctrl+点选 / Ctrl+框选：状态栏报块数、右键菜单按全部块执行）**、
> 多标签页与跨表工作区（**工作区 = 一个个独立单元格**：任何入口都逐格拆分、按行优先排、
> **绝对跳过空内容单元格**（空字符串 / 空格 / 全角空格都算空；`0`/`false`/公式都算内容）、
> 保留来源单元格尺寸、一格一个小方块、**搜索内容 / 按来源表筛选**、可拖宽度、条目右键复制/剪切/删除 + 粘贴进表格、
> **快速加入：选择模式下选中单元格后 3 秒内点一下工作区空白就能收进来**、
> 「+」导入行/列/区域与**一键全搬整表已用区域**、清空要二次确认、拖放守卫（条目在工作区里拖不会自我复制））、
> 右键菜单（**复制到工作区 / 剪切到工作区** / 与工作区互换 / 复制 / 粘贴 / 清空内容（**保留格式且可撤销**）/ 撤销重做）、
> 原子撤销与可回退的历史面板、互换后不再留下选中而是给出"落点提醒"、
> 关闭后重开恢复现场（IndexedDB：标签、编辑、工作区、模式、界面设置）、
> **外科式导出回 xlsx**（只回写改动过的单元格，其余 zip 部件字节原样保留，图表/透视表等不会丢）、深浅色主题与动画。
>
> 交互模式怎么用（工具栏三个图标按钮，悬停有说明）：
> **选择** = 完全原生（可框选多格，**Ctrl+点选可累加不连续的多块区域**）；
> **拖拽** = **按住直接拖**即可搬运/互换内容，全程不出现多选框；
> **点击互换** = 依次点两个单元格即互换（**也可以先点右侧工作区条目、再点表格里的格子**）。
> 任何互换/搬运完成后都会自动取消选中。
> 任一模式下按住 **Shift / Ctrl** 都退回原生选择（可框选），用于"区域互换"这类需要先选一片区域的场景。
>
> 工作区怎么用：右上角 **「+」** 打开面板 —— 一个输入框（`A1:B18` / `3:5` / `B:D` / 多块 `A1:B2 D4:E5` 自动识别）
> + 一块**实时预览**（`将加入 12 个单元格 · 跳过 156 个空内容`，并列出前几格长什么样，确认键上直接写「加入 12 格」）
> + 一个**「加入工作区后保留表格内容」**开关（取消勾选＝剪切：内容进工作区、源格内容清空、保留格式，一次撤销就能回来）；
> 里面还有一颗 **「整个已用区域」** 按钮（把整表有内容的部分填进输入框，看清预览再确认）；
> 放进工作区的**每一个格子都是一条独立条目**（按行优先排序、保留它原本的尺寸、**空内容一律不进**），
> 上方搜索框可按**内容**或**单元格地址**（如 `B12`）找，右侧下拉可按**来源工作表**筛；
> 每行放几格由面板宽度自动决定，底部「格宽 N px」可以调；拖表格区与工作区之间的**分隔条**可改宽度；
> 顶部「拖到工作区保留内容」与底部「写回表格后移除条目」两个开关决定搬运语义（前者与面板里那个开关是**同一个设置**）；
> 条目上**右键**可复制/剪切/删除，之后在表格里右键「粘贴」或按 Ctrl/Cmd+V 写回；
> 点「清空」会先问一次（可只清当前筛选出来的那些）。
>
> **最快的加入方式（快速加入）**：在**选择模式**下点一下表格里的单元格（Ctrl+点选可多块），
> **3 秒内**再点工作区面板的**空白处**，选区就整片收进来了（走的是同一个入口：逐格拆分、跳过空内容、
> 按 id 去重、**可撤销**）。规则很克制，避免误加：只在选择模式生效；拖到面板里松手不算（那是"拖入工作区"）；
> 点搜索框/按钮/条目不算；同一次选区连点两下也只加一次；超过 3 秒后点空白只是点空白。
>
> 性能（实测，100 万格 / 6.1 MB）：解析 **1180 ms**、端到端导入 **1439 ms**、拖动帧间隔 p50 **4.2 ms**；
> 导入后只保留"导出真正需要的东西"（原始字节 + 行样式映射），解析模型不再常驻内存。
>
> **能打开哪些格式**（按钮写的是「打开表格」，悬停是手形）：
>
> | 格式 | 能不能开 | 说明 |
> | --- | --- | --- |
> | `.xlsx` 工作簿 | ✅ | 主用格式（原字节直用，导出逐字节保真） |
> | `.xlsm` 启用宏的工作簿 | ✅ | 同一条 OOXML 路径；**宏部件原样保留、不执行**，导出仍是 `.xlsm` |
> | `.xltx` / `.xltm` 模板 | ✅ | 包结构相同，照常预览与编辑 |
> | `.csv` / `.tsv` / `.txt` | ✅ | 自研 RFC4180 解析：**GBK/UTF-8/UTF-16 编码识别**、分隔符嗅探、引号规则；数字转数值，`007` 这类带前导零的留文本 |
> | `.ods`（OpenDocument） | ✅ | 自研 zip+XML 解析：重复行列、合并、文本/日期/时长、样式、列宽行高、冻结、隐藏表 |
> | `.xls`（Excel 97–2003） | ✅ | 自研 **BIFF8 二进制** 解析：手写 CFB（OLE2）容器 + 记录流（SST/RK/XF/FONT/FORMAT/MERGEDCELLS…） |
> | `.xlsb` | ❌ | Excel 二进制 OOXML（zip 里是 binary parts），需要另一套解析器；提示"另存为 .xlsx" |
> | `.numbers` / `.et` | ❌ | iWork / WPS 私有格式；提示"另存为 .xlsx" |
>
> **`.csv` / `.ods` / `.xls` 的导入方式**：自研解析 → 翻译成一份**规范 xlsx** → 再走既有全链路
> （预览/编辑/工作区/撤销/多标签/会话恢复一个能力都不缺），所以**导出会另存为 `.xlsx`**
> （`原名-已编辑.xlsx`），**原文件一个字节都不动** —— 这条会写在导入摘要里。
> 真 Excel 产出的 `.xls/.ods/.csv` 样本由 `npm run fixtures:legacy` 生成（需要本机 Excel），
> 它们是解析器的**真值**：单元与端到端用例都拿"同内容的 .xlsx"逐格对拍。

技术栈：React 18 + Vite 6 + TypeScript（strict）+ Univer 0.25.1 + Vitest + Playwright
**不含任何服务端**：xlsx 解析、转换、渲染、编辑、拖拽全部在浏览器内完成。

---

## 快速开始

### 方式一：一键运行（推荐，Windows）

**双击 `run.cmd`** 即可：检查环境 → 首次自动装依赖 → 起服务 → 打开浏览器。
它**只负责把工具跑起来**，不会跑测试或基准。

```
run.cmd                    启动工具（默认端口 5273，被占用会自动顺延）
run.cmd -Port 5300         换端口
run.cmd -NoBrowser         不自动打开浏览器
run.cmd -SkipInstall       跳过依赖检查
```

**关闭那个窗口（或按 Ctrl+C）服务就结束**，端口会被释放；即使窗口被强杀来不及清理，
也有一个脱离控制台的看门狗进程兜底，在脚本进程消失后释放端口，不会留下后台 node。
若端口上已经跑着本工具，再次双击只会**复用**它并打开浏览器，不会起第二份。

> 测试与基准不在这个启动器里，用 npm 脚本：
> `npm test`（类型检查 + 单元测试）、`npm run e2e`（端到端）、`npm run bench:parse`（性能基准）。

### 方式二：手动

```bash
cd excel-preview

# 1) 安装依赖
npm install
#    （若你的环境限制 npm 缓存目录，用：npm install --cache .npm-cache）

# 2) 生成回归样本（6 个覆盖关键样式的 xlsx + 1 个启用宏的 .xlsm + manifest.json）
npm run fixtures
#    可选：再用本机 Excel 另存出真 .xls / .ods / .csv 样本（解析器真值；没装 Excel 会打印提示并跳过）
npm run fixtures:legacy

# 3) 启动
npm run dev          # http://localhost:5273
```

页面上点「打开表格」导入 `fixtures/` 下的任意样本即可看到渲染效果（`.xlsx/.xlsm/.ods/.xls/.csv` 都行）；
右侧「事件日志」会实时显示 Univer 事件、格式锁拦截、导入报告（含"由 X 格式转换导入"的说明）。

### 双形态：本地版 与 静态版

| | **本地版**（默认） | **静态版** |
| --- | --- | --- |
| 怎么起 | `npm run dev`（或 `run.cmd`）；打包 `npm run build:local` | `npm run build:static` |
| 产物 | `dist/`（`base: '/'`） | `dist/`（`base: './'`，可直接丢到 GitHub / Gitee / EdgeOne Pages 的项目子路径下） |
| 本机 Excel 转换桥 | ✅ 有：`.ods/.xls/.xlsb` 导出、`.xlsb` 导入都**完整保真**（由 Excel 自己另存，样式/公式/条件格式/批注/图片都在） | ❌ 没有本机进程：这几项在导出菜单里**置灰并写明"本地版可用"**，`.xlsb` 仍按"打不开"处理并指路 |
| 速度 | 一样（同一套产物结构与启动优化） | 一样 |
| 数据 | 全程在本机（浏览器 IndexedDB） | 全程在浏览器里，**任何文件都不上传服务器** |

- **静态版怎么发**：`npm run build:static` → 把 `dist/` 传到托管。仓库里已经带了一份
  GitHub Actions 工作流（`.github/workflows/static-pages.yml`：`npm ci → build:static → 产物体检 → 单测 → deploy-pages`），
  推到 `main` 即自动发布；项目站点在 `/<repo>/` 子路径下也能用（已实测）。`dist/.nojekyll` 已自动带上（跳过 Jekyll 处理）。
- **本地桥怎么工作**：`npm run dev` 时 Vite 中间件挂出 `GET /api/bridge/health` 与
  `POST /api/bridge/convert?to=ods|xls|xlsb|xlsx`；收到浏览器发来的 xlsx 字节后，调
  `tools/excel-bridge.ps1`（Excel COM：不可见、关警告、禁宏）另存为目标格式再回传。
  浏览器**不能传路径**、单请求有体积上限、转换串行、临时目录用完即删。
- **没装 Excel 的机器**：本地版照样能用，只是那三项导出/`.xlsb` 导入会明确说"本机 Excel 不可用"。

### 启动速度（本机实测，同一台机器、同一份产物）

| 指标 | 优化前 | 优化后 | 怎么做到的 |
| --- | --- | --- | --- |
| 首屏可见（骨架/首次内容绘制） | 556 ms（此前是**白屏**，一直等到应用挂载） | **44 ms**（骨架可见 79 ms） | 骨架直接写在 `index.html` + 内联关键 CSS；`App` 改懒加载 |
| 入口 chunk | 6792 KB（gzip 1900 KB） | **1.4 KB（gzip 0.8 KB）** | 入口只做"接管 Ctrl+Z + 挂载 React + 动态 import App" |
| 分块 | 1 个巨块 + 69 个断词词典 | 8 个：入口 / App / vendor-react / vendor / vendor-univer / 解析器… | `manualChunks` 分包 + 剔除非必要分块 |
| 产物总量 | 11.17 MB / 82 个文件 | **6.79 MB / 13 个文件** | 剔除 Univer 的 69 个断词词典分块（表格工具用不到，见下） |
| 到"表格可用" | 425 ms | 445 ms | 代码总量没变（本地磁盘快，差别在噪声内）；真正的收益是首屏不再白屏、重复访问只重下业务分块 |

> **关于"只保留 zh-CN 语言包"**：构建日志里曾出现 `de-1901 / hu / th / ru / uk …` 几十个"语言文件"，
> 它们其实是 `@univerjs/engine-render` 里给**文档排版**用的**断词词典**（69 种语言，动态 import）。
> 本工具只用表格、从不创建文档排版器，所以构建期由 `vite-plugins/vite-plugin-drop-hyphenation.ts` 把那张表清空
> （`loadPattern()` 里本来就有 `if (!loader) return;`），产物里**不再有任何非中文语言数据**
> （已核：无 `Arbeitsblätter` / `Feuilles de calcul` / `"de-DE"` 等字面量，只有中文文案）。
> 用 `npm run measure:bundle` 可以随时体检（断词词典数量必须为 0）；e2e 的启动预算用例也会断言。

### 本地运行要点

- **纯前端**：不需要后端、不需要数据库；xlsx 全程在浏览器里解析与导出，文件不会离开本机。
- **浏览器要求**：需要 Chromium 内核（Chrome / Edge 105+）。用的是 Canvas 2D 与 IndexedDB。
- **端口**：默认 `5273`（`vite.config.ts` 里 `strictPort`）。被占用时改配置或 `npm run dev -- --port 5274`。
- **生产构建**：`npm run build:local` / `npm run build:static`（先 `tsc --noEmit` 再 `vite build`），
  产物在 `dist/`，是纯静态文件；本地看效果用 `npm run serve:dist`（默认挂在 `http://127.0.0.1:5399/repo/`，
  刻意带子路径，与 Pages 项目站点一致）。
- **端到端测试**用**系统已装的 Chrome**（`playwright.config.ts` 里 `channel: 'chrome'`），
  不会去下载 Chromium；所以本机装过 Chrome 就能跑 `npm run e2e`。Playwright 会自动起/复用 5273 端口的 dev 服务。
- **大文件**：性能夹具默认 5 万行 × 20 列（100 万格），生成与测试都是可选的，不影响功能验证。

> 维护者注意（`tools/start.ps1` / `tools/run-helper.ps1` 里都有注释，这里列个索引）：
> ① `.cmd` 文件必须保持纯 ASCII——cmd.exe 按 OEM 代码页读它，中文会被误解析成命令；
> ② 两个 `.ps1` 必须存成 **UTF-8 with BOM**，否则 Windows PowerShell 5.1 读中文会乱码；
> ③ PowerShell 函数不要把 native 子进程的输出留在返回值里（会和退出码拼成数组，导致"全绿判失败"）；
> ④ 函数参数名必须与调用处一致：简单函数会把绑不上的参数悄悄塞进 `$args`，静默传空值。

## 测试

```bash
npm run typecheck      # tsc --noEmit（零报错）
npm test               # Vitest：461 个用例 / 29 个文件（解析器（OOXML/CSV/ODS/XLS）/ 样本真值 /
                       #   导入适配与格式判定 / xlsx 合成写入器 / 只读门禁 / 交互与三种模式 /
                       #   像素↔格换算 / 多区域选区 / 选区与合并块对齐 / 工作区逐格拆分与空内容跳过 /
                       #   「+」面板的识别与预览文案、目标范围换算 /
                       #   日期取值口径 / 提醒框与滚动条几何 / 多标签常驻窗口 / 持久化 / 导出保真 /
                       #   右键菜单 / 历史 / 撤销快捷键归属 / 性能护栏）
npm run e2e            # Playwright：111 个端到端用例 / 25 个文件（导入保真与导入摘要 / 多格式兼容 /
                       #   只读约束 / 导出与导出菜单（含本地 Excel 桥导出 .ods）/ 三种交互模式与选区外观 /
                       #   不连续多区域 / 工作区与跨表 / 工作区快速加入（选中后点空白）/ 工作区拖放守卫（不自我复制、过期选区不误加）/
                       #   工作区撤销与重做（含"拖回表格可撤销""面板剪切/复制两条语义""打开文件不留幽灵条目"）/
                       #   静态形态（子路径托管可用 + 启动预算 + 桥能力降级）/
                       #   右键菜单 / 历史 / 持久化 / 日期单元格 / 缩放组件 /
                       #   滚动与滚动条 / 坐标一致性 / 提示统一 / 选区与合并块 / 文件拖放与打开格式 /
                       #   多标签内存 / 热更新后标签不乱 / 百万格性能）
```

### 测试要跑多久（以及怎么再快一点）

本机（16 逻辑核 / 23 GB）实测：

| 命令 | 内容 | 耗时 |
| --- | --- | --- |
| `npm run typecheck` | `tsc --noEmit` | ~8 s |
| `npm test` | 单元 461 例（Vitest 多线程跑文件） | ~15–30 s（看机器忙不忙） |
| `npm run e2e` | 端到端 **111 例全量**（两阶段串起来） | **~3 分钟** |
| `npm run e2e:fast` | 只跑"主项目"97 例（跳过 14 例串行用例） | ~2 分钟 |
| `npm run e2e:sensitive` | 只跑那 14 例串行用例 | ~1.2 分钟 |

**为什么是两阶段**（`main` 并发 + `sensitive` 串行）—— 并发能把 8 分钟压到 3 分钟，但有几类用例
天生不能和别人抢机器，放一起会**测出假失败**（不是放宽断言，而是让它们各自在干净环境里跑）：

| 串行用例 | 为什么必须串行 |
| --- | --- |
| `hot-reload.spec.ts` | 它会**改写 `src/App.tsx`** 触发真实热更新；开发服务器会把更新广播给所有页面，并跑会把别的 worker 的页面一起重启 |
| `perf.spec.ts` | 断言 p95 帧间隔 < 200ms —— 并跑时数字被别的 canvas 抢走 |
| `tabs-memory.spec.ts` | 采堆 + 冷存重建，同样吃"机器安静"这个前提 |
| `static-form.spec.ts` | 测的是启动时间预算（骨架/FCP/可用时间） |
| `date-cells.spec.ts` | ④ 依赖"悬停 → 上游弹提醒"这条时序（自己一个文件并发 6 也能过，但和别的 spec 抢机器时悬停一直不触发） |

顺带把"忙时假红"的断言改成**轮询/观察器**（这类改法**更严**，不是更松）：

- `workspace-undo.spec.ts`：条目数断言改成 `expect.poll`（撤销走的是异步命令，固定 sleep 在忙时会假红）；
- `modes.spec.ts`：黄色提醒框只活 ~1.6 秒，改成**松手前装 MutationObserver 记录"它出现过"**，
  不再依赖"回头数 DOM 时它还在"；
- `date-cells.spec.ts`：反向对照（"关掉配置后确实会弹"）改成轮询 + 每次重新悬停一下。

想临时调并发度：`PW_WORKERS=8 npm run e2e`（本项目在 6 与 8 之间耗时几乎一样，6 更稳）。
只想跑一两个文件时，直接点名即可（`--project=main` 走并发池）：

```bash
node node_modules/@playwright/test/cli.js test tests/e2e/workspace.spec.ts --project=main --workers=6
```
> 被删掉的是重复覆盖、探索性用例与"一个行为拆七八条"的穷举；关键断言（22 条禁令命令、
> 导出字节级保真、tokenizer 回归、三种模式、工作区逐格拆分与跨表共享、
> **工作区绝不放空内容条目**、持久化、
> 撤销（含"一次动作 = 一条历史"与工作区动作按时间顺序撤销）、
> 日期搬家不退化、滚动与缩放可用、滚动后拖动起点落点仍准、提示统一、多标签冷存重建、
> **热更新后标签不多不少**、**选区没有没用的填充柄**、**格式判定与宏部件原样保留**、
> **`.xls/.ods/.csv` 与同内容 xlsx 逐格一致**、**Ctrl+点选多区域**、
> **所有表格/工作区操作都走同一条可撤销通道**（含"工作区条目拖回表格可撤销"与"打开文件不留幽灵条目"））一条没少。

**多标签与内存**：标签数不设上限；**同时实体化的工作簿最多 3 个**（常驻窗口），
超出的非活动标签会被冷存（释放工作簿与渲染器），切回时按需重建并回放编辑。
实测开 4 个中等标签时，第 4 个进来后堆从 347 MB 回落到 194 MB（`tests/e2e/tabs-memory.spec.ts`）。
现状与边界见 `M4-多标签内存分析.md`。

性能基准（夹具与产物不进仓库）：

```bash
npm run bench:fixture  # 生成 fixtures/bench-large.xlsx（50k 行 × 20 列 = 100 万格，约 6 MB）
npm run bench:parse    # Node 侧管线基准，打印 解析/适配/吞吐/堆内存 指标表
```

自定义规模（PowerShell）：`$env:BENCH_ROWS='100000'; $env:BENCH_COLS='30'; npm run bench:fixture`；
CPU 剖面：`$env:BENCH_PROFILE='1'; node tools/bench-parse.mjs`；
换夹具跑端到端性能：`$env:PERF_FIXTURE='bench-medium.xlsx'; npm run e2e -- tests/e2e/perf.spec.ts`。

> 端到端默认使用**系统已安装的 Chrome**（`playwright.config.ts` 里 `channel: 'chrome'`），
> 避免下载 Playwright 自带 Chromium。CI 环境可改回 `devices['Desktop Chrome']`。
> 视觉基线在 `tests/e2e/*-snapshots/`，布局改动后用 `--update-snapshots` 重建。

## 目录结构

```
src/
  parser/        自研 OOXML 解析器（中立模型，不依赖 Univer，可在 Node 里单测）
    xml.ts        索引扫描式流式 tokenizer（性能关键路径）
    styles.ts     styles.xml + theme1.xml（字体/填充/边框/数字格式/主题色+tint）
    worksheet.ts  单元格值/公式/行高列宽/合并/冻结/未支持项记账
    index.ts      parseXlsx(Uint8Array) -> ParsedWorkbook
  parser/          自研解析：OOXML（工作表/样式/条件格式/数据验证/批注/表格/图片/打印设置）、
                   CSV/TSV/文本表格（含 GBK 识别）、ODS、XLS（BIFF8 二进制）
  importer/
    to-univer.ts  ParsedWorkbook -> IWorkbookData（单位换算/枚举映射都在这里）
    file-kinds.ts 能打开哪些格式（含 .csv/.ods/.xls）与打不开时的"为什么+怎么办"
    synth-xlsx.ts 把非 OOXML 的表格翻译成一份规范 xlsx（于是下游能力一个不少）
    open-workbook.ts 统一打开入口：按格式分派 → 一律产出"字节 + 解析结果"
  univer/
    setup.ts      createUniver 引导（已关闭 ribbon / 格式工具栏）
    lock.ts       「只允许编辑内容、不允许改格式」两层防线
    swap-command.ts  原子互换/移动命令（自己推送 undo/redo 记录）
    fill-handle.ts   关掉选区右下角那个"填充柄"（自动填充被永久拒绝，留着就是假入口）
  interaction/
    drag-ghost.ts     rAF 合并的拖拽幽灵
    drag-controller.ts 拖拽会话：命中 → 落点决策 → 回调（命中可注入，非选择模式下用自建命中测试）
    click-swap.ts     三种交互模式的状态机（选择 / 拖拽 / 点击互换）
    fly-in.ts         落位动画（rAF + transform）
    ghost-content.ts  把快照渲染成幽灵 DOM（与侧边栏共用样式映射）
  workspace/
    types.ts         快照模型（预览样式与写回值分离；单格条目带 cellSize）
    snapshot.ts      Univer 选区 ⇄ 快照：extractCellItems 逐格拆分 / previewWorkspaceImport 预览 / applySnapshot 只写内容地写回
    filter.ts        工作区的搜索与来源筛选（纯函数）
    ImportDialog.tsx / import-parse.ts / import-target.ts / import-dialog.css
                     「+」面板：自动识别行/列/区域、实时预览、保留/剪切选项；import-target 是提交与预览共用的范围换算
                     （单独成文件是为了让 App.tsx 只导出组件 —— 否则 Fast Refresh 会退化成整页刷新）
    WorkspacePanel.tsx / SnapshotPreview.tsx / preview-style.ts / workspace.css
  shell/shell.css    外壳设计系统（令牌 / 布局 / 动效 / 深浅色）
  shell/history-model.ts  历史账本模型（跳步计划，纯函数）
  shell/undo-shortcut.ts  撤销/重做快捷键的第一顺位接管（必须在引导前注册，见下方"坑"第 7 条）
  parser/          自研解析：OOXML（工作表/样式/条件格式/数据验证/批注/表格/图片/打印设置）、
                   CSV/TSV/文本表格（含 GBK 识别）、ODS、XLS（BIFF8 二进制）
  p0/log.ts       事件日志（Playwright 通过这些日志做断言）
tools/
  make-fixtures.mjs   生成合成样本（xlsx + .xlsm）
  make-legacy-fixtures.ps1 用本机 Excel 另存出真 .xls / .ods / .csv（解析器真值）
  inspect-fixture.mjs 打印样本内部 XML 真值（排查解析问题用）
fixtures/             6 个 xlsx + 1 个 .xlsm +（可选）.xls/.ods/.csv 样本 + manifest.json（期望值清单）
tests/
  unit/         解析器单测 + 样本覆盖度
  e2e/          端到端（引导/离线/格式锁/互换/手势/导入渲染）
```

## 开发时必须知道的坑（都是实测踩出来的，详见 P2 交付说明）

1. **不要给 Univer 宿主加 `StrictMode`**：React 18 双执行 effect 会让 Univer 完全不渲染
   （`canvas` 数量为 0），但 API 层却一切正常，极难排查。详见 `src/main.tsx` 注释。
2. **测试里不要用裸 `canvas` 选择器**：Univer 内部有 0×0 的 docs 画布会排在最前，
   坐标类手势测试会全线错位；请用 `canvas[id^="univer-sheet-main-canvas"]`。
3. **Univer 的撤销不是自动的**：命令必须在成功后自己 `pushUndoRedo({undoMutations, redoMutations})`，
   否则 Ctrl+Z 无效。
4. **传 `null` 给 `SetRangeValuesMutation` 是"整格删除"**（样式一起没）；
   清空内容要用 `{ v: null, f: null, si: null, p: null }`，即 `emptyContentCell()`。
5. **拦截/观测写入要挂 mutation 层**，不要挂 `beforeCommandExecuted`——后者对 Facade 发起的写入不触发。
6. **`ICommandService` 没有 `interceptCommand`**（旧版 API 已移除），别照抄网上旧文章。
7. **Ctrl+Z / Ctrl+Y 的第一顺位在 Univer 手里**：它的 `ShortcutService` 引导时就用
   `window.addEventListener('keydown', …, { capture: true })` 注册了撤销/重做，
   同 target 同阶段按**注册顺序**执行——晚注册的监听器 `stopPropagation()` 拦不住它，
   `preventDefault()` 也不被它理会（实测现象：Univer 先撤了一步、我们的账本原地不动，
   接着再按就"没反应"）。所以我们的转发器必须在 `main.tsx` 里、**引导之前**注册，
   处理时用 `stopImmediatePropagation()` 截下按键；详见 `src/shell/undo-shortcut.ts`。
8. **`undos` 计数会把同一件事记两次**：我们自己的表格命令既会让 `undos` 涨（订阅自动补一条历史），
   我们又手工记一条。要么在动作期间"认领并合并"，要么在自发的 `undo/redo` 期间抑制
   （`beginSheetAction` / `suppressAutoEntriesRef`），否则账本会多出"幽灵步骤"。
9. **"空单元格"不止 `null` 一种**：用户在编辑器里把内容删光后，Univer 存下来的是**空字符串 `''`**；
   中文表格里还常见只有全角空格 `'　'` 的格子。只判 `=== null` 会让这些"看着是空的"格子
   混进工作区变成白卡片（用户实测反馈）。统一用 `src/workspace/snapshot.ts` 的
   `isEmptyContent()`（无公式 + `trim()` 后为空；`0`/`false`/公式都算内容）。
10. **Excel/WPS 的 `fills[1] = gray125` 是规范占位**，几乎每个文件都有、但通常没有任何 `cellXfs`
   引用它。解析时不要"看到就警告"，否则用户会以为文件有毛病（实测：状态栏因此显示"降级 1 项"）。
   按引用关系上报（`parseStyles` 末尾），未引用的填充保持沉默。
11. **改代码（Vite HMR / React Fast Refresh）会把 App 的引导 effect 重跑一遍，
   但 `refs`/`state` 是保留的**：旧 Univer 实例已在 cleanup 里 dispose，可内存里还记着那些标签。
   若此时照常"启动恢复"，会话里的标签会被再追加一遍 → 标签栏多出一个同 id 同名的"灵魂标签"
   （点它报 `no document with unitId …`），控制台还会刷 React "two children with the same key"。
   处理见 `App.tsx` 引导 effect 开头的 `app:hot-reboot`（把标签标成未实体化）、恢复循环里的
   **按 id 去重**（`session:restore-skip-registered`）、以及 cleanup 里"先把现场同步写库再拆监听"。
   回归用例：`tests/e2e/hot-reload.spec.ts`（会临时往 `App.tsx` 末尾追加一行注释来真实触发 HMR，`finally` 还原）。

## 已确认的关键事实

- Univer 开源包（`@univerjs/*`）为 **Apache-2.0**；xlsx 导入导出属于 **Pro 付费且需服务端**，本项目**不用**它，改为自研解析/修补。
- Univer 自带渲染性能指标管道：`@univerjs/sheets-ui` 的 `SheetRenderController.renderMetric$`
  （公开导出，每 60 帧汇总 FPS / 各渲染阶段 max-min-avg），性能面板接它即可；
  `@univerjs/telemetry` 只是上报出口，注入是 `Optional`，**不注册实现就完全不出网**。
- 实测证据图见 [`evidence/`](./evidence/)：条件格式四类（高亮/色阶/数据条/图标集）、
  浮动图片锚点、超链接与批注的渲染效果。

## 并发跑测试

```bash
PW_OUTPUT_DIR=test-results-b npm run e2e   # Playwright 每次运行会清空 outputDir，并发时需分开
```
