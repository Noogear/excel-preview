/**
 * 关掉选区右下角那个小方块 —— Univer 的**填充柄**（auto-fill handle）。
 *
 * 用户实测反馈："选中单元格的时候，选中的单元格右下角会有个角标，但是这个角标毫无用处，
 * 能否利用起来或者直接不要显示这个角标？"
 *
 * 为什么它在本工具里一定是"没用"的：
 *  - 填充柄是 Univer 给**自动填充（拖拽填充序列/复制）**用的入口，而自动填充会**批量改写一片单元格**
 *    （既复制内容、也可能连格式一起铺开）。本产品只允许"逐格改内容"，命令闸门
 *    （`src/univer/read-only-guard.ts` 的 `/auto-fill|refill/i`）**永久拒绝**这条命令 ——
 *    也就是说这个手柄永远不会成功，纯属误导。
 *  - 更糟的是：默认的"拖拽"模式下，我们自己的内容搬运会接管指针 —— 用户捏着这个手柄一拖，
 *    实际上是**把这一格的内容和落点那一格互换了**（实测日志："已互换 A12 ⇄ A15"）。
 *    一个看起来像"填充"的手柄却悄悄换了内容，这是必须消掉的风险。
 *
 * 于是这里直接让它**不画**：`SelectionControl` 的 `_enableAutoFill = false` 会让
 * `_updateControl()` 里的 `this._autoFillControl.hide()` 永远生效（它本来就是这个开关）。
 * 但控件是 Univer 内部按"渲染单元 × 选区"创建的，没有公开的开关，所以：
 *  ① 包一层 `newSelectionControl()`（在服务实例上，幂等），保证**以后新建的**控件都是关的；
 *  ② 顺手把**已经存在**的控件关掉（切换工作表/多选会新建控件，所以两步都要）。
 *
 * 与 `read-only-guard` 的分工：闸门负责"就算有人绕过 UI 也改不动"，
 * 这里负责"不给用户一个点了没用的假入口"。
 */
import { SelectionControl } from '@univerjs/sheets-ui';

/** `SelectionControl` 里我们真正要用的那几个字段（0.25.1 的私有字段名，升级时看这里） */
interface SelectionControlLike {
  /** 为 false 时 `_updateControl()` 不会画填充柄 —— 官方自己的开关 */
  _enableAutoFill?: boolean;
  _showAutoFill?: boolean;
  /** 填充柄那个 Rect（公开 getter），`visible` 是它当前是否在画 */
  fillControl?: { visible?: boolean; hide?: () => void } | null;
}

interface SelectionRenderServiceLike {
  newSelectionControl?: (...args: unknown[]) => SelectionControlLike | null | undefined;
  getSelectionControls?: () => SelectionControlLike[];
  /** 我们自己打的标记：同一个服务实例只包一次 */
  __excelPreviewFillHandleOff?: boolean;
}

/** 渲染单元（`renderManager.getRenderById(unitId)`）的最小形态 */
export interface RenderUnitLike {
  with?: (token: unknown) => unknown;
}

/** 一个控件：不再显示填充柄（已经画出来的那一帧也立刻收掉） */
function muteControl(control: SelectionControlLike | null | undefined): void {
  if (!control) return;
  control._enableAutoFill = false;
  control._showAutoFill = false;
  control.fillControl?.hide?.();
}

/**
 * **类级**关闸（引导时调一次，必须在第一个工作簿创建之前）。
 *
 * 为什么还需要这一层（用户实测反馈："选区右下角那个小方块目前依旧绘制"）：
 * 下面那个"按渲染单元处理"的办法，前提是**渲染单元已经存在**。
 * 但启动时我们是 `loadWorkbook(示例) → attachSheetDeps()` 连着做的，那一刻渲染单元还没建出来，
 * 于是示例工作簿的选区控件是用**原版工厂**造的，右下角的小方块照样画（实测：`enabled:[true]`、
 * 截图里那个蓝色小方块清晰可见）。用户第一次打开应用看到的就是它。
 *
 * 这一层不依赖任何单元：直接包 `SelectionControl.prototype.updateRangeBySelectionWithCoord`——
 * 选区控件每次更新选区都会走它（构造时也会走一次），在委托之前先把开关按死，
 * 覆盖"现在与将来、所有单元、所有选区"。
 */
export function installFillHandleOff(): boolean {
  const proto = (SelectionControl as unknown as { prototype?: SelectionControlProto } | undefined)?.prototype;
  if (!proto || typeof proto.updateRangeBySelectionWithCoord !== 'function') return false;
  if (proto.__excelPreviewFillHandleOff) return true;
  const original = proto.updateRangeBySelectionWithCoord;
  proto.updateRangeBySelectionWithCoord = function patched(this: SelectionControlLike, ...args: unknown[]): unknown {
    muteControl(this);
    return original.apply(this, args);
  };
  proto.__excelPreviewFillHandleOff = true;
  return true;
}

interface SelectionControlProto {
  updateRangeBySelectionWithCoord?: (...args: unknown[]) => unknown;
  __excelPreviewFillHandleOff?: boolean;
}

/**
 * 关掉某个工作簿（渲染单元）的填充柄。
 *
 * - `ok`：拿到了渲染服务并处理过（含"之前就处理过"）；
 * - `firstTime`：**这次**才第一次给它打上开关（调用方据此只记一条日志，避免刷屏）。
 *
 * 拿不到渲染服务时返回 `{ ok: false }` 并**不做任何事**（宁可留着那个小方块，也不能在这里抛异常
 * 影响引导链路）。
 */
export function disableFillHandleOn(
  renderUnit: RenderUnitLike | null | undefined,
  selectionToken: unknown,
): { ok: boolean; firstTime: boolean } {
  let service: SelectionRenderServiceLike | null = null;
  try {
    service = (renderUnit?.with?.(selectionToken) ?? null) as SelectionRenderServiceLike | null;
  } catch {
    return { ok: false, firstTime: false };
  }
  if (!service) return { ok: false, firstTime: false };

  // ② 已经存在的控件（当前选区、多选产生的其它控件）立即生效
  try {
    for (const control of service.getSelectionControls?.() ?? []) muteControl(control);
  } catch {
    /* 拿不到就算了，下面的包装仍会覆盖后续新建的控件 */
  }

  // ① 包一层工厂，保证以后新建的控件也是关的（幂等）
  if (service.__excelPreviewFillHandleOff) return { ok: true, firstTime: false };
  if (typeof service.newSelectionControl !== 'function') return { ok: true, firstTime: false };
  const original = service.newSelectionControl.bind(service);
  service.newSelectionControl = (...args: unknown[]) => {
    const control = original(...args);
    muteControl(control);
    return control;
  };
  service.__excelPreviewFillHandleOff = true;
  return { ok: true, firstTime: true };
}

/**
 * 诊断/测试用：当前活动单元的填充柄状态。
 *
 * `visible` 直接读那个 Rect 的公开 getter —— 它才是"画面上到底有没有那个小方块"的真身；
 * `enabled` 是我们的开关（应当恒为 false）。
 */
export function fillHandleStateOn(
  renderUnit: RenderUnitLike | null | undefined,
  selectionToken: unknown,
): { available: boolean; controls: number; enabled: boolean[]; visible: (boolean | null)[] } {
  let service: SelectionRenderServiceLike | null = null;
  try {
    service = (renderUnit?.with?.(selectionToken) ?? null) as SelectionRenderServiceLike | null;
  } catch {
    service = null;
  }
  if (!service) return { available: false, controls: 0, enabled: [], visible: [] };
  const controls = service.getSelectionControls?.() ?? [];
  return {
    available: true,
    controls: controls.length,
    enabled: controls.map((control) => control._enableAutoFill !== false),
    visible: controls.map((control) => control.fillControl?.visible ?? null),
  };
}
