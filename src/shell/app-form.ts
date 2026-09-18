/**
 * **双形态**（用户要求"静态、本地双形态"）。
 *
 * | 形态 | 构建 | 转换桥 | 场景 |
 * | --- | --- | --- | --- |
 * | `local`（默认） | `npm run build:local` / `npm run dev` | ✅ | 本机跑，带 Excel 转换桥：`.ods/.xls/.xlsb` 导出、`.xlsb` 导入都能**完整保真** |
 * | `static` | `npm run build:static` | ❌ | 丢到 GitHub/Gitee Pages 等静态托管，任何老师打开即用；桥相关入口隐藏/置灰并说明"本地版可用" |
 *
 * 形态在构建时由 `__APP_FORM__` 编进产物（见 `vite.config.ts`）；桥的可用性在**运行时**探测
 * （`GET /api/bridge/health`）—— 本地版也可能因为没装 Excel 而用不了桥，所以两件事分开判断：
 * `isLocalForm` 决定"要不要显示这个入口"，`bridge.available` 决定"点了能不能成"。
 */

/** 构建期注入的形态（`vite.config.ts` 的 define） */
declare const __APP_FORM__: 'local' | 'static';

export type AppForm = 'local' | 'static';

/** 产物形态；测试/运行时兜底成 local（裸跑 TS 时不会炸） */
export const APP_FORM: AppForm = typeof __APP_FORM__ === 'string' ? __APP_FORM__ : 'local';

/** 本地版（含转换桥的那一份） */
export const isLocalForm = APP_FORM === 'local';

/** 桥的 HTTP 入口（本地版由 Vite 中间件提供；静态版没有这个路由） */
export const BRIDGE_BASE = '/api/bridge';

export interface BridgeStatus {
  /** 探测中 */
  checking: boolean;
  /** 能不能用（本地版 + 本机装了 Excel + 中间件在） */
  available: boolean;
  /** 本机 Excel 版本（能拿到就显示给用户看，排错方便） */
  excel?: string;
  /** 不可用时的原因（直接展示，不粉饰） */
  reason?: string;
  /** 目标格式 → 该格式的扩展名（桥能产出的格式） */
  formats: BridgeFormat[];
}

/** 桥能产出的文件格式（与 `tools/excel-bridge.ps1` 的 FileFormat 映射保持一致） */
export type BridgeFormat = 'xlsx' | 'ods' | 'xls' | 'xlsb';

const EMPTY_BRIDGE: BridgeStatus = { checking: false, available: false, formats: [] };

/** 静态形态下桥永远不可用（连探测都不发请求，避免控制台出现 404 噪音） */
export function initialBridgeStatus(): BridgeStatus {
  return isLocalForm
    ? { checking: true, available: false, formats: ['xlsx', 'ods', 'xls', 'xlsb'] }
    : { ...EMPTY_BRIDGE, reason: '静态托管版没有本机进程，Excel 转换能力只在本地版可用' };
}

/**
 * 探测转换桥。**永不抛错**：拿不到就当不可用，并把原因写进 `reason` 给用户看。
 *
 * 超时给 20 秒（与中间件的探测预算一致）：本机 Excel 的 COM 探测要**启动一次 Excel**，
 * 冷启动 2–6 秒很正常 —— 一开始只给 2.5 秒，结果"脚本能探测到、界面却说不可用"（踩过）。
 * 真正的转换另有 90 秒预算（见 `src/exporter/bridge.ts`）。
 */
export async function probeBridge(timeoutMs = 20_000): Promise<BridgeStatus> {
  const formats: BridgeFormat[] = ['xlsx', 'ods', 'xls', 'xlsb'];
  if (!isLocalForm) return initialBridgeStatus();
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const response = await fetch(`${BRIDGE_BASE}/health`, { signal: controller.signal });
    clearTimeout(timer);
    if (!response.ok) {
      return { checking: false, available: false, formats, reason: `转换桥无响应（HTTP ${response.status}）` };
    }
    const body = (await response.json()) as { available?: boolean; excel?: string; reason?: string };
    return {
      checking: false,
      available: Boolean(body.available),
      excel: body.excel,
      reason: body.reason,
      formats,
    };
  } catch (error) {
    const aborted = (error as Error)?.name === 'AbortError';
    return {
      checking: false,
      available: false,
      formats,
      reason: aborted
        ? `探测本机 Excel 超时（超过 ${Math.round(timeoutMs / 1000)} 秒）`
        : `连不上本地转换桥：${String((error as Error)?.message ?? error)}`,
    };
  }
}

/** 桥不可用时，给用户看的一句人话（各入口共用，避免文案各不相同） */
export function bridgeUnavailableHint(status: BridgeStatus): string {
  if (isLocalForm) return status.reason ?? '本机 Excel 转换桥不可用（本地版需要本机安装 Excel）';
  return '这一项只在本地版可用（静态托管版没有本机 Excel 进程）';
}

/** 桥能负责的格式 → 扩展名 */
export const BRIDGE_EXTENSION: Record<BridgeFormat, string> = {
  xlsx: '.xlsx',
  ods: '.ods',
  xls: '.xls',
  xlsb: '.xlsb',
};

/** 把"原名-已编辑.xlsx"换成目标格式的文件名 */
export function bridgedFileName(originalName: string, format: BridgeFormat): string {
  const base = originalName.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlsb|ods|csv|tsv|txt)$/i, '');
  return `${base}-已编辑${BRIDGE_EXTENSION[format]}`;
}
