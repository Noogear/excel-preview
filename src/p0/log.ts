/**
 * P0 事件日志：把 Univer 触发的事件、我们注入的手势探测结果记录下来，
 * 供 Playwright 通过 window.__p0log 读取断言，也方便人工在页面上观察。
 */
export interface P0LogEntry {
  /** 相对页面启动的毫秒数 */
  t: number;
  kind: string;
  detail?: unknown;
}

const MAX_ENTRIES = 3000;
const startedAt = performance.now();

export const p0log: P0LogEntry[] = [];

export function log(kind: string, detail?: unknown): void {
  p0log.push({ t: Math.round(performance.now() - startedAt), kind, detail });
  if (p0log.length > MAX_ENTRIES) p0log.splice(0, p0log.length - MAX_ENTRIES);
  // 便于人工在 DevTools 里观察
  // eslint-disable-next-line no-console
  console.debug('[p0]', kind, detail ?? '');
}

export function clearLog(): void {
  p0log.length = 0;
}

/** 统计某类事件出现次数 */
export function countKind(kind: string): number {
  return p0log.filter((e) => e.kind === kind).length;
}

export function findLast(kind: string): P0LogEntry | undefined {
  for (let i = p0log.length - 1; i >= 0; i--) {
    if (p0log[i].kind === kind) return p0log[i];
  }
  return undefined;
}
