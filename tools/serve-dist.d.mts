/**
 * `tools/serve-dist.mjs` 的类型声明（给 e2e 直接 import 用）。
 */
export interface StaticServerHandle {
  url: string;
  port: number;
  close: () => Promise<void>;
}

export function startStaticServer(options?: {
  distDir?: string;
  prefix?: string;
  port?: number;
}): Promise<StaticServerHandle>;
