/**
 * 本地转换桥的**前端客户端**：把"当前内容的 xlsx 字节"交给本机 Excel，换回目标格式的字节。
 *
 * 两件必须做对的事：
 *  ① 失败要给出**能读懂的原因**（没装 Excel / 被别的程序占用 / 超时），而不是一个空按钮；
 *  ② 超时兜底：Excel 冷启动 + 大表可能慢，但不能让用户对着转圈的按钮无限等。
 *
 * 静态托管形态下这个模块不会被调用（入口已按 `isLocalForm` 隐藏/置灰）。
 */
import { BRIDGE_BASE, type BridgeFormat } from '../shell/app-form';

export interface BridgeConvertResult {
  ok: boolean;
  bytes?: Uint8Array;
  reason?: string;
}

/** 一次转换最多等多久（毫秒）：Excel 冷启动 + 保存大表，给足 90 秒 */
export const BRIDGE_CONVERT_TIMEOUT_MS = 90_000;

export async function convertViaBridge(
  input: Uint8Array,
  to: BridgeFormat,
  timeoutMs = BRIDGE_CONVERT_TIMEOUT_MS,
): Promise<BridgeConvertResult> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(`${BRIDGE_BASE}/convert?to=${encodeURIComponent(to)}`, {
      method: 'POST',
      headers: { 'content-type': 'application/octet-stream' },
      body: input as unknown as BodyInit,
      signal: controller.signal,
    });
    if (!response.ok) {
      const detail = (await response.json().catch(() => null)) as { reason?: string } | null;
      return { ok: false, reason: detail?.reason ?? `转换桥返回 HTTP ${response.status}` };
    }
    return { ok: true, bytes: new Uint8Array(await response.arrayBuffer()) };
  } catch (error) {
    const message = String((error as Error)?.message ?? error);
    const aborted = (error as Error)?.name === 'AbortError';
    return {
      ok: false,
      reason: aborted ? `转换超时（超过 ${Math.round(timeoutMs / 1000)} 秒）` : `转换失败：${message}`,
    };
  } finally {
    clearTimeout(timer);
  }
}
