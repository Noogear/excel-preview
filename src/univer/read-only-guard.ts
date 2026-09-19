/**
 * 全局只读约束闸门（**默认拒绝**）。
 *
 * 产品约束（用户需求原文）：只允许修改单元格的**文字内容**；
 * 字体/加粗/边框/背景色/字色/斜体等一切样式、以及结构（增删行列、合并、行列尺寸）、
 * 图片、批注、条件格式、数据验证、超链接、子表创建……统统不允许编辑。
 *
 * 为什么用"默认拒绝"而不是"枚举坏命令"：
 *   Univer 的命令分散在 10+ 个包里、且每个命令都可能触发若干 mutation，
 *   枚举"坏命令"必然漏（漏一个就是数据被改坏）。
 *   因此这里反过来做——**只放行一份白名单**，其余一律拦掉。
 *
 * 与既有防线的关系（三层，互补）：
 *   1. 本闸门：命令层默认拒绝（拦结构/对象/样式/子表）
 *   2. `lock.ts` 命令层：把 30 个样式命令注销并替换为 no-op（UI 点了也没反应）
 *   3. `lock.ts` mutation 层：剥离 cellData 里的 `s`（兜底——任何路径写样式都写不进去）
 *
 * 实现方式：包装 `ICommandService` 单例上的 `executeCommand` / `syncExecuteCommand`。
 * 这是**唯一**能实现"默认拒绝"的位置：0.25.1 的 ICommandService 没有拦截器
 * （P0 实测：`interceptCommand` 已被移除，`beforeCommandExecuted` 无法取消执行）。
 */
import { ICommandService, type IExecutionOptions } from '@univerjs/core';

export interface ReadOnlyGuard {
  /** 被拦下的命令 id 及次数（用于调试与"为什么点不动"的解释） */
  blockedIds: () => Map<string, number>;
  blockedCount: () => number;
  isAllowed: (id: string) => boolean;
  /**
   * 临时放行**所有**命令（仅用于导入期由我们自己的可信代码应用条件格式/数据验证/图片等）。
   * 必须成对使用 `resume()`（建议 try/finally）；用户交互期间绝不能处于挂起状态。
   */
  suspend: () => void;
  resume: () => void;
  isSuspended: () => boolean;
  restore: () => void;
}

/**
 * 放行规则（**必须同时满足"命中放行"且"未命中拒绝"**）。
 *
 * 放行的理由分别是：
 *  - `sheet.operation.*`：操作**不写入文档快照**（选区、滚动、编辑器显隐、行列脏标记等），
 *    拦掉它们会让表格无法点选/滚动/输入
 *  - `undo/redo`：撤销重做本身必须可用（它回放的 mutation 会再走一遍本闸门）
 *  - 值与公式：`set-range-values`（打字/公式栏/我们的互换与写回）
 *  - 剪贴板：`univer.command.copy`（真实 id 不在 `sheet.command.*` 下，见下方长注释）；粘贴走 `sheet.command.paste*`
 *  - 内容清理与传播：`clear-selection-content`/`auto-clear-content`/`copy-down`/`copy-right`
 *  - 选择与导航：`move-selection*`/`select-all`/`expand-selection`/`scroll-*`
 *  - 视图：`change-zoom-ratio`
 *  - 切表：`set-worksheet-activate`（多标签必需）
 */
const ALLOW_PATTERNS: RegExp[] = [
  /^univer\.command\.(undo|redo)$/,
  /**
   * 我们自己的命令命名空间。
   * 注意：导入期"应用条件格式/数据验证/超链接/批注/图片"走的是 **Univer 自己的命令**，
   * 那些会被闸门拦掉——因此导入期用 `suspend()` 临时放行（见 App 的 handleFile），
   * 而不是把那些命令加进白名单（否则用户也能手动改 CF/DV 了）。
   */
  /^excel-preview\.command\./,
  /**
   * 单元格内编辑走的是 **docs 子系统**（编辑器本身是一份富文本文档）：
   * 打字是 `doc.command.insert-text`、提交快照是 `doc.command-replace-snapshot`、
   * 编辑器选区是 `doc.operation.set-selections`。
   * 实测教训（三个坑，别只放行 command）：
   *  1) 一开始只放行 `sheet.*`，结果打字进不了编辑器、回车把原单元格**清空**了（数据丢失）；
   *  2) docs 的 id 有两种写法——`doc.command.xxx` 与 `doc.command-xxx`（连字符），不能带尾点匹配；
   *  3) 编辑器真正的文本变更落在 **`doc.mutation.rich-text-editing`** 上，
   *     只放行 command/operation 仍然打不字（实测被拦 43 次）。
   * 富文本格式类命令由 DENY 拦掉，且 mutation 层会把 `p` 拍平成纯文本做兜底。
   */
  /^doc\.(command|operation|mutation)/,
  /** 公式引擎的重算记账（只写计算结果，不碰样式） */
  /^formula\.(mutation|command)/,
  /^sheet\.operation\./,
  /**
   * 复制：**真实 id 在 `univer.command.*` 命名空间里**，不是 `sheet.command.*`。
   *
   * 实测教训（真实缺陷，e2e 抓到的）：`@univerjs/ui` 里
   * `CopyCommandName = 'univer.command.copy'`，而 `@univerjs/sheets-ui` 的
   * `SheetCopyCommand = { id: CopyCommand.id, name: 'sheet.command.copy' }` ——
   * **name 与 id 不是一回事**，命令注册表认的是 id。
   *
   * 之前这里白名单写的是 `sheet.command.(copy|cut)`，那两个 id 在 0.25.1 里
   * **根本不存在**（派发会报 "command is not registered"），于是这条"放行"是空放行，
   * 真命令反被"默认拒绝"拦掉 —— 表现就是 **Ctrl+C 完全没反应**（剪贴板纹丝不动、
   * 也没有任何提示），而右键菜单里我们自己实现的「复制内容」是好的，所以一直没被发现。
   *
   * **`cut` 故意不放行**：`univer.command.cut` 本身只做标记，真正清空源发生在粘贴时，
   * 而那条路径可能落到被本闸门拦掉的 `move-range` / `reorder-range` 上，
   * 会变成"剪切看着成功、却粘不出来"的半坏状态。产品里的剪切由「剪切到工作区」承担。
   * （`sheet.command.copy` 只作为将来 id 改名的兼容别名保留，0.25.1 下不存在。）
   */
  /^univer\.command\.copy$/,
  /^sheet\.command\.copy$/,
  /^sheet\.command\.set-range-values$/,
  /^sheet\.command\.paste(-value|-by-short-key)?$/,
  /^sheet\.command\.optional-paste$/,
  /^sheet\.command\.clear-selection-content$/,
  /^sheet\.command\.auto-clear-content$/,
  /^sheet\.command\.(copy-down|copy-right)$/,
  /^sheet\.command\.(move-selection|move-selection-enter-tab)$/,
  /^sheet\.command\.(select-all|expand-selection)$/,
  /^sheet\.command\.scroll-(to-cell|view|view-reset)$/,
  /**
   * 滚轮滚动走的就是这条（`SheetsScrollRenderController._wheelEventListener` →
   * `SetScrollRelativeCommand`）。实测教训：漏了它，**滚轮完全滚不动**（一次滚轮一条被拦记录），
   * 用户看到的就是"滚动条/滚动无法正常使用"（用户实测反馈）。
   * 它只改视口滚动位置（operation 级），不写文档快照、不改样式。
   */
  /^sheet\.command\.set-scroll-relative$/,
  /^sheet\.command\.change-zoom-ratio$/,
  /**
   * 缩放滑块/百分比下拉走的是这个命令（`@univerjs/sheets-ui` 的 ZoomSlider → `SetZoomRatioCommand`），
   * 而 +/− 按钮走的是上面的 `change-zoom-ratio`。
   * 实测教训：只放行 `change-zoom-ratio` 时，"右下角缩放条拖不动、百分比菜单点了没反应"——
   * 因为这条被本闸门拦掉了（用户实测反馈）。它只写 `set-zoom-ratio` operation（纯视图、不进撤销栈、
   * 不改文档快照与样式），与放行 `change-zoom-ratio` 是同一类，因此一并放行。
   */
  /^sheet\.command\.set-zoom-ratio$/,
  /^sheet\.command\.set-worksheet-activate$/,
  /^sheet\.command\.set-cell-edit-[a-z-]+$/,
  /^sheet\.mutation\.set-range-values$/,
];

/**
 * 拒绝规则（优先级**高于**放行规则）。
 * 这里只列"即使命中放行也要拦"的高危家族，其余靠"默认拒绝"兜住。
 */
const DENY_PATTERNS: RegExp[] = [
  // 样式家族（含编辑器内的富文本格式化：加粗/斜体/下划线/删除线/字体/字号/前后景/对齐/缩进/项目符号）
  /format/i,
  /style/i,
  /border/i,
  /bold|italic|underline|strike|overline/i,
  /font|typeface/i,
  /(fore|back|background|foreground)-?color/i,
  /align|indent|heading|bullet|list/i,
  /reset-(background|text)-color/i,
  // 结构与尺寸
  /merge/i,
  /(column|col|row)-(width|height)/i,
  /delta-(column|row)/i,
  /(insert|remove|delete)-(row|rows|col|cols|column|columns|range|sheet|sheets|cell|cells)/i,
  /reorder-range|move-(rows|cols|range)/i,
  /append-row/i,
  /auto-fill|refill/i,
  // 对象与规则
  /table|filter|sort|pivot|sparkline|drawing|image|float-dom|note|comment|hyperlink|validation|conditional/i,
  /protection|frozen|gridline|theme/i,
  /defined-name/i,
  // 子表与工作簿结构
  /(insert|remove|copy)-sheet/i,
  /(sheet|worksheet)-(hide|show|order|rename|name|row-count|column-count|default-style|right-to-left)/i,
  /^sheet\.operation\.rename-sheet$/,
  /^sheet\.operation\.set-format-painter$/,
  // 其它
  /repeat-last-action/i,
  /clear-selection-all/i,
];

export function isCommandAllowed(id: string): boolean {
  if (DENY_PATTERNS.some((pattern) => pattern.test(id))) return false;
  return ALLOW_PATTERNS.some((pattern) => pattern.test(id));
}

interface CommandServiceLike {
  executeCommand: (id: string, params?: unknown, options?: IExecutionOptions) => Promise<unknown>;
  syncExecuteCommand: (id: string, params?: unknown, options?: IExecutionOptions) => unknown;
}

/**
 * 已经装过闸门的命令服务 → 句柄。
 *
 * 为什么需要：`installReadOnlyGuard` 是**猴补**（直接改实例上的方法），装两次就会包两层，
 * 而且第一层的 `restore()` 会把第二层的包装一起抹掉（两边都想还原"自己记下的原始方法"）。
 * 以前只靠调用方"每次 boot 只装一次"的自觉；现在这里兜住：同一个服务重复安装直接返回同一个句柄。
 * 用 `WeakMap` 而不是给服务挂属性：不给上游对象留我们的痕迹，也不会阻止它被回收。
 */
const installedGuards = new WeakMap<object, ReadOnlyGuard>();

/**
 * 安装闸门。传入**根 injector**（`univer.__getInjector()`）——命令服务是全局单例，
 * 装一次即可覆盖所有工作表与所有标签页。
 */
export function installReadOnlyGuard(injector: { get: (token: unknown) => unknown }): ReadOnlyGuard {
  const commandService = injector.get(ICommandService) as CommandServiceLike;
  const existing = installedGuards.get(commandService);
  if (existing) return existing;

  const blocked = new Map<string, number>();
  let suspended = false;
  let restored = false;

  /** 包装前的**自有属性描述符**（没有自有属性说明方法来自原型，restore 时要把我们加的自有属性删掉） */
  const ownExecute = Object.getOwnPropertyDescriptor(commandService, 'executeCommand');
  const ownSyncExecute = Object.getOwnPropertyDescriptor(commandService, 'syncExecuteCommand');
  const originalExecute = commandService.executeCommand.bind(commandService);
  const originalSyncExecute = commandService.syncExecuteCommand.bind(commandService);

  const record = (id: string): void => {
    blocked.set(id, (blocked.get(id) ?? 0) + 1);
  };

  commandService.executeCommand = (id: string, params?: unknown, options?: IExecutionOptions) => {
    if (!suspended && !isCommandAllowed(id)) {
      record(id);
      return Promise.resolve(false);
    }
    return originalExecute(id, params, options);
  };

  commandService.syncExecuteCommand = (id: string, params?: unknown, options?: IExecutionOptions) => {
    if (!suspended && !isCommandAllowed(id)) {
      record(id);
      return false;
    }
    return originalSyncExecute(id, params, options);
  };

  const guard: ReadOnlyGuard = {
    blockedIds: () => new Map(blocked),
    blockedCount: () => [...blocked.values()].reduce((sum, n) => sum + n, 0),
    isAllowed: isCommandAllowed,
    suspend: () => {
      suspended = true;
    },
    resume: () => {
      suspended = false;
    },
    isSuspended: () => suspended,
    /**
     * 拆掉猴补。
     *
     * 这里按"包装前是不是自有属性"分别处理：原来是自有属性就还原那份描述符，
     * 原来在原型上就**删掉**我们加的自有属性（以前是写成 bind 后的副本，等于永远留了一层自有属性
     * 盖住原型方法 —— 有界但不是"干净还原"，而且重复安装时会互相打架）。
     */
    restore: () => {
      if (restored) return;
      restored = true;
      suspended = false;
      if (ownExecute) Object.defineProperty(commandService, 'executeCommand', ownExecute);
      else delete (commandService as unknown as Record<string, unknown>).executeCommand;
      if (ownSyncExecute) Object.defineProperty(commandService, 'syncExecuteCommand', ownSyncExecute);
      else delete (commandService as unknown as Record<string, unknown>).syncExecuteCommand;
      installedGuards.delete(commandService);
    },
  };
  installedGuards.set(commandService, guard);
  return guard;
}
