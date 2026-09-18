# P0 技术验证结论（Univer 0.25.1 实测）

> 验证方式：在 `excel-preview/` 建立最小工程（Vite + React + TS + Univer 0.25.1），
> 通过**读源码/类型定义**与**运行真实代码**两条路径核实，不依赖文档猜测。
> 状态：技术栈核实完成 ✅ · 解析链路跑通 ✅ · 端到端手势测量见文末。

---

## 一、结论速览

| 待验证问题 | 结论 | 影响 |
| --- | --- | --- |
| 工程能否跑起来（Vite+React+TS+Univer） | ✅ 可以，`tsc --noEmit` 全绿 | 无 |
| 自研 OOXML 解析 → Univer 渲染 是否成立 | ✅ 成立，6 个真实样本全部解析成 `IWorkbookData` | 自研导入层路线确认 |
| 长按拖动所需的目标单元格命中 | ✅ Facade 事件已提供，**不需要**内部 API | 拖拽方案可行 |
| "只允许编辑内容、不允许改格式" | ⚠️ 可用，但**不是**文档里常见的那种做法 | 见第三节 |
| `@univerjs/telemetry` 能否做性能监控 | ⚠️ 不能直接用；但有更好的官方指标管道 | 见第二节 |
| 完全离线（无任何外部请求） | ✅ 静态检查通过（后续 e2e 再断言一次） | 无 |

---

## 二、`@univerjs/telemetry` 的真相（修正你给的技术栈）

**它不是性能监控工具，但你的直觉不算错——Univer 内部确实用它上报渲染耗时。**

实测证据：

1. `@univerjs/telemetry` 只定义接口：README 原文 "defines the telemetry service interface used by Univer packages"，导出物只有 `ITelemetryService`（注入标记 + 接口）。没有任何采集、统计、可视化能力。
2. `@univerjs/sheets-ui` 依赖它，并在 `SheetRenderController` 里做了这件事（`lib/es/index.js:15112-15232`）：
   - 订阅引擎的 `renderFrameTimeMetric$` / `renderFrameTags$`，按帧累积指标
   - 每 **60 帧**（`FRAME_STACK_THRESHOLD = 60`）汇总一次：对每个数值字段算 `max/min/avg`，并附上 `elapsedTimeToStart`、`FPS`、`frameTime`
   - 推给 `renderMetric$`，同时调用 `this._telemetryService?.capture('sheet_render_cost', telemetryData)`
3. 注入方式是 `Optional(ITelemetryService)` → **不注册实现就是 no-op，不会发任何网络请求**（离线要求天然满足；`@univerjs/preset-sheets-core` 的产物里也没有任何硬编码 http(s) 端点）。

### 因此性能监控的正确设计是

```ts
// 1) 官方渲染指标（公开导出：@univerjs/sheets-ui 导出 SheetRenderController / ITelemetryData）
const render = injector.get(IRenderManagerService).getRenderById(unitId);
render.with(SheetRenderController).renderMetric$.subscribe((m: ITelemetryData) => {
  // m: { FPS, frameTime, elapsedTimeToStart, 各渲染阶段: { max, min, avg } }
});

// 2) 我们自己补"导入链路"的分阶段计时（Univer 不覆盖这块）
//    解压 → SAX 解析 → 适配转换 → createWorkbook → 首帧
//    已实现：src/App.tsx 的 import 流程里记录 parseMs / adaptMs

// 3) ITelemetryService 只作为"可选上报出口"：默认注册 no-op 实现（保证离线）
```

---

## 三、"只允许编辑内容、不允许改格式"：正确姿势（重要修正）

### 3.1 被证伪的做法

- ❌ **`ICommandService.interceptCommand(...)`**：0.25.1 的 `ICommandService` **没有这个方法**（旧版 API，已移除）。实测其接口只有：
  `hasCommand / registerCommand / unregisterCommand / registerMultipleCommand / executeCommand / syncExecuteCommand / onCommandExecuted / beforeCommandExecuted / onMutationExecutedForCollab`
  ——其中 `beforeCommandExecuted(listener)` 的回调返回 `void`，**无法取消执行**。
- ⚠️ **仅靠权限点**：`IPermissionService` + `WorksheetSetCellStylePermission` / `WorksheetSetCellValuePermission` 等权限点确实存在（`@univerjs/sheets` 公开导出），但执行者是 `SheetPermissionCheckController`，其内部是 `_commandExecutedListener` —— 即**命令执行之后**再检查并触发 UI 拦截/提示。属于"事后拦截"，不能作为唯一防线（数据已经被改了）。
- ❌ `fWorksheet.getWorksheetPermission().setMode('readOnly')`：一刀切，会连内容编辑一起锁掉。

### 3.2 P0 采用的两层防线（已实现：`src/univer/lock.ts`）

| 层 | 做法 | 覆盖的风险路径 | 实测 |
| --- | --- | --- | --- |
| A. 命令层 | 把 30 个样式类命令（`SetStyleCommand`、`SetBoldCommand`、`SetBorderCommand`、`SetColWidthCommand`、`AddWorksheetMergeCommand`…）`unregisterCommand` 后用 no-op 命令占位 | 工具栏点按、快捷键、菜单 | ✅ 拦截生效、样式未变 |
| B. **mutation 层** | `unregisterCommand(SetRangeValuesMutation.id)` → 注册包装版：先把 `cellData` 里的 `s`（样式）剥离，再交给原始 handler | 打字、粘贴、格式刷、填充、拖拽移动——**所有写入路径的兜底** | ✅ 夹带样式被剥离、值正常写入 |

B 层是关键：Univer 单元格模型是 `{ v, s, f, p }`，值与样式分离，因此"只写 `v`、剥掉 `s`"既满足"能编辑内容"，又**从数据层保证格式永不改变**。

> 附带发现：`IRowData`/`IColumnData` 也有 `s`（行/列级样式），P1 需一并纳入剥离范围。

---

## 四、拖拽与命中测试：Facade 事件已足够

`@univerjs/sheets-ui` 的 Facade 事件（`lib/types/facade/f-event.d.ts`）实测提供：

| 事件 | 参数 | 用途 |
| --- | --- | --- |
| `CellPointerMove` / `CellHover` / `CellPointerDown` / `CellPointerUp` / `CellClicked` | `{ row, column, worksheet, workbook }` | **坐标→单元格**，拖动中实时命中 |
| `DragOver` / `Drop` | `{ row, column, dataTransfer, ... }` | 外部 DOM（侧边栏）拖入表格，原生给落点 |
| `SelectionMoveStart` / `SelectionMoving` / `SelectionMoveEnd` | `{ selections }` | 选区拖动生命周期 |
| `Scroll` | `{ scrollX, scrollY }` | 自建几何索引时的滚动补偿 |
| `SheetZoomChanged` / `BeforeSheetZoomChange` | `{ zoom }` | 缩放补偿 |
| `BeforeSheetEditStart` / `BeforeClipboardChange` / `BeforeClipboardPaste` | 支持 `params.cancel = true` | 可阻断编辑/粘贴 |
| `SheetSkeletonChanged` | `{ skeleton, effectedRanges }` | 几何信息兜底 |

另外，`@univerjs/sheets` 原生就有 `MoveRangeCommand` / `MoveRangeMutation` / `ReorderRangeCommand`（"拖动移动内容"），因此：
- **"移动"** 可直接复用原生命令；
- **"互换"** 需自研（读两边值 → 交叉写回），实现已在 `src/App.tsx` 的 `swapValues` 中验证通过。

---

## 五、自研解析链路：跑通，但发现真实文件上的缺陷

### 5.1 已跑通

`fixtures/` 下 6 个用 ExcelJS 生成的真实样本，经 `parseXlsx → toUniverWorkbook` 全部转成 `IWorkbookData`（`tests/unit/fixture-coverage.test.ts`，7 passed）：

| 样本 | 覆盖内容 |
| --- | --- |
| `fixture-styles.xlsx` | 边框/填充/字体/对齐/合并/行高列宽/隐藏行列/富文本/冻结 |
| `fixture-numfmt.xlsx` | 17 种数字格式（日期/百分比/千分位/货币/自定义/科学计数/[h]:mm:ss…）+ 主题色+tint |
| `fixture-rules.xlsx` | 条件格式四类（高亮/色阶/数据条/图标集）+ 数据验证五类 |
| `fixture-extras.xlsx` | 批注、超链接（外部/内部/邮件）、浮动图片（oneCell/twoCell 锚点） |
| `fixture-table.xlsx` | Excel 表格(Table)：表头样式 + 斑马纹 + 汇总行 |
| `fixture-multi.xlsx` | 多工作表、顺序与文件名不一致、隐藏/深度隐藏、公式与缓存值 |

解析器还把**未解析的部件显式记入 `report.unsupported`**（如"条件格式 4 处""数据验证 1 处""表格对象 1 处"），而不是静默丢弃——这正是"降级报告"产品能力的底座。

### 5.2 发现并已委派修复的缺陷

对 `fixture-styles.xlsx` 做 XML 真值统计：`<row>` 19 个、`<c>` 40 个、`<fill>` 10 个、`<border>` 11 个、`<pane state="frozen">` 存在。
解析器实际只得到：单元格 **8**、行 **5**、fill **2**、border **1**、`freeze: null`。

**根因**：自闭合子元素打断容器遍历（`<patternFill patternType="none"/>`、`<left/>`、`<c r="B1" s="1"/>` 之后兄弟元素被整体丢弃）——即 `xml.ts` 的 tokenizer 在 `/>` 处理上索引推进有误。
已连同真值对照表委派修复，并要求补充"自闭合标签"回归用例。

> 这正是 P0 的价值：**如果直接进入 P1，这个 bug 会让"边框/填充保真"整体失效**，而它只在真实 OOXML 上暴露。

---

## 六、本机环境事实（影响后续开发流程）

| 项 | 实测 | 影响 |
| --- | --- | --- |
| Node / npm | `v24.13.1` / `11.4.1`，**无 pnpm** | 用 npm 即可；Univer 官方示例用 pnpm，不冲突 |
| npm registry | `registry.npmmirror.com`，ping 189ms | 安装快（472 包 / 20s） |
| fs 沙箱 | 全局 npm 缓存目录（`D:\software\nodejs\node_cache`）在工作区外，**写入被拒** | 必须加 `--cache .npm-cache` 指到工作区内 |
| 子进程沙箱 | 受限模式禁止命名管道 → `npm` 生命周期脚本、`esbuild`（Vite/Vitest 依赖）、Playwright 均 `spawn EPERM` | 跑 `npm install` / `vitest` / `vite` / `playwright` 需提权 |

---

## 七、端到端测量结果（Playwright：**13 passed / 0 failed / 54s**）

| 验证项 | 实测结果 |
| --- | --- |
| 引导 | 17ms 完成；**外部网络请求 0 条** → "完全前端 + 离线"达成 |
| 渲染保真（渲染侧） | 字体/字号/粗体/字色/底色/对齐/四边边框值可从数据模型精确读回；截图见 `test-results/diagnose.png` |
| 内容编辑 | 改值成功，且**边框样式一字不变** |
| 格式锁 A 层（命令） | 派发 `sheet.command.set-style` 被 no-op 命令拦截，样式未变 |
| 格式锁 B 层（mutation） | 故意写入夹带 `bg:#FF00FF` 的单元格：`sheet.mutation.set-range-values` 剥离 1 个样式字段，**值写入成功、`styleLeaked: false`** |
| **长按拖动命中** | 长按 400ms 被正确识别；**按住鼠标拖动时 `CellPointerMove` 连续触发 10 次**，并给出真实 `{row: 11, column: 25}` → **命中测试可纯走 Facade 事件，无需内部 API、无需自建几何索引** |
| 与原生手势冲突 | 整个拖动过程中 `SelectionMoveStart/Moving/End` **一次都没触发** → 从单元格内部拖动不会与 Univer 原生"拖边框移动内容"手势打架，可安全自建长按手势 |
| 内容互换 | 两格值互换成功，双方样式各自保持不变 |
| 外部拖放（HTML5 DnD） | 自定义 `dataTransfer` 能到达容器（`application/x-p0-univer-range` 可见），但 **Univer 的 `Event.DragOver` / `Event.Drop` 对合成事件不触发**（真实用户拖拽下是否可用未定；不值得赌）→ **侧边栏拖回表格改用 pointer 事件 + 已验证的 `CellPointerMove` 命中** |
| 数字格式保真 | 17 种格式显示全部正确：`#,##0.00`→1,234.57、`0.0%`→25.7%、`yyyy-mm-dd`→2025-01-01、`yyyy"年"m"月"d"日"`→2025年1月1日、`[h]:mm:ss`→36:00:00、`¥#,##0.00`→¥1,234.50、`# ?/?`→ 1/2、`0.00E+00`→1.23E+08、格式内字面量→`总计 1,234.50 元`、`m/d/yy h:mm`→1/1/25 12:00 |
| 降级报告 | 条件格式 / 数据验证 / 表格对象等未支持项按种类计数进入 `unsupported`，**不是静默丢弃** |
| 单元测试 | **44 passed**（解析器 37 + 样本覆盖 7），`tsc --noEmit` 全绿 |

### 附带发现的两个工程陷阱（都已修，且都是"极具迷惑性"的那类）

1. **React 18 `StrictMode` 会让 Univer 完全不渲染。**
   effect 双执行（mount → cleanup → mount）后：第一个实例被 cleanup `dispose()`，而渲染容器里的全局单例 DOM（`univer-doc-selection-container-__INTERNAL_EDITOR__DOCS_NORMAL`）无法被第二个实例正确接管，最终 `document.querySelectorAll('canvas').length === 0`——**表格一个像素都不画，但应用状态却显示"已就绪"**（因为所有 API 调用都成功了，日志也正常）。排查花了很久，最后靠 Playwright 诊断用例 dump DOM 才定位。
   → 结论：**Univer 宿主必须单实例、显式管理生命周期**；已从 `main.tsx` 移除 `StrictMode` 并在代码里写明原因（同时保留"canvas 层数必须恰为 3"的断言防止实例残留）。

2. **不能用 `beforeCommandExecuted` 观测 Facade 发起的写入。**
   `fRange.setValue()` 走完后，`beforeCommandExecuted` 监听器**未触发**（`allowedCount: 0`），但 mutation 层**一定触发**（B 层验证里 `sheet.mutation.set-range-values` 正常命中）。
   → 结论：要观测/拦截写入，一律挂 **mutation 层**，不要挂命令层。

---

## 八、P0 结论：可以进入 P1

四个不确定点全部有了确定答案，且都不需要额外的架构改动：

| P0 待验证问题 | 答案 |
| --- | --- |
| ① 工程能否跑通 | ✅ Vite+React+TS+Univer 0.25.1，类型检查与 44 个单测全绿 |
| ② 自研解析→渲染是否成立 | ✅ 6 个真实样本全链路渲染成功，数字格式 17/17 正确，缺陷已修 |
| ③ 长按拖动与原生手势是否冲突 | ✅ 不冲突；`CellPointerMove` 在按住时持续触发，命中方案确定 |
| ④ 权限点/拦截能否锁住格式 | ✅ 两层防线实测生效（命令层拦截 + mutation 层剥离） |

**修订后的风险清单（相对评估报告）**：原风险 #4（长按手势冲突）与 #6 之外的多数风险已降级；
新增需要注意的一条：**Univer 与 React StrictMode 不兼容**，这在 P1 搭正式 UI 时必须遵守。
