/**
 * 构建期插件：**本地转换桥**（只在「本地版」形态下挂载）。
 *
 * 提供两个接口（都只监听本机回环地址，浏览器只能通过它们间接使用本机 Excel）：
 *   GET  /api/bridge/health   → `{ available, excel?, reason? }`
 *   POST /api/bridge/convert?to=xlsx|ods|xls|xlsb  （body = 文件字节）→ 转换后的字节
 *
 * 为什么桥在 Vite 里（而不是独立进程）：项目本来就是"双击 run.cmd 起一个本地服务再打开浏览器"，
 * 复用这个服务零额外进程/端口；`vite preview` 也走同一套中间件，本地版两种起法行为一致。
 * 静态版（`APP_FORM=static`）**不挂**这个中间件 —— 静态托管上根本没有本机进程，前端会据此降级。
 *
 * 安全边界（与 `tools/excel-bridge.ps1` 一起）：
 *   ① 只接受**字节流 + 格式枚举**，浏览器不能传路径（路径全在服务端临时目录里生成）；
 *   ② 体积上限、格式白名单、请求串行（Excel COM 不能并发）；
 *   ③ 每次请求建独立临时目录，结束即删；超时杀掉自己启动的 EXCEL.EXE；
 *   ④ 只绑定 127.0.0.1（由 Vite 开发服务器本身决定）。
 */
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import type { Connect, Plugin } from 'vite';

/** 目标格式 → Excel SaveAs 的 FileFormat 号 */
const SAVE_AS: Record<string, number> = {
  xlsx: 51,
  xlsm: 52,
  xlsb: 50,
  xls: 56,
  ods: 60,
};

const EXTENSION: Record<string, string> = {
  xlsx: 'xlsx',
  xlsm: 'xlsm',
  xlsb: 'xlsb',
  xls: 'xls',
  ods: 'ods',
};

/** 单次转换的请求体上限（Excel 侧也吃不下更大的；正常班级表远小于它） */
const MAX_BYTES = 64 * 1024 * 1024;
/** 单次转换超时（毫秒）：大表 + Excel 冷启动，给足 60 秒 */
const CONVERT_TIMEOUT_MS = 60_000;
/** 健康探测超时（毫秒） */
const PROBE_TIMEOUT_MS = 20_000;

interface ProbeResult {
  available: boolean;
  excel?: string;
  reason?: string;
}

interface BridgeOptions {
  /** 应用形态：静态版不挂桥 */
  form: 'local' | 'static';
  /** 仓库根目录（用来定位 tools/excel-bridge.ps1） */
  root: string;
  /** 日志（默认走 Vite 的 config.logger） */
  log?: (message: string) => void;
}

function runPowerShell(args: string[], timeoutMs: number): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((fulfil) => {
    const child = spawn('pwsh', ['-NoProfile', '-NonInteractive', '-File', ...args], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      fulfil({ code: -1, stdout, stderr: `${stderr}\n[bridge] 超时 ${timeoutMs}ms，已结束` });
    }, timeoutMs);
    child.stdout?.on('data', (chunk) => {
      stdout += String(chunk);
    });
    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });
    child.on('error', (error) => {
      clearTimeout(timer);
      fulfil({ code: -1, stdout, stderr: `${stderr}\n${String(error)}` });
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      fulfil({ code: code ?? -1, stdout, stderr });
    });
  });
}

function sendJson(res: Parameters<Connect.NextHandleFunction>[1], status: number, body: unknown): void {
  const payload = JSON.stringify(body);
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  res.setHeader('cache-control', 'no-store');
  res.end(payload);
}

async function readBody(req: Connect.IncomingMessage, limit = MAX_BYTES): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    const buffer = chunk as Buffer;
    size += buffer.length;
    if (size > limit) throw new Error(`请求体超过上限 ${Math.round(limit / 1024 / 1024)}MB`);
    chunks.push(buffer);
  }
  return Buffer.concat(chunks);
}

export function excelBridgePlugin(options: BridgeOptions): Plugin {
  const script = resolve(options.root, 'tools', 'excel-bridge.ps1');
  const note = options.log ?? (() => {});
  /** 探测结果缓存（Excel 装没装不会中途变） */
  let probe: ProbeResult | null = null;
  /** Excel COM 不能并发：转换串行排队 */
  let queue: Promise<unknown> = Promise.resolve();

  async function probeExcel(): Promise<ProbeResult> {
    if (probe) return probe;
    if (!existsSync(script)) {
      probe = { available: false, reason: `找不到转换脚本 ${script}` };
      return probe;
    }
    const result = await runPowerShell([script, '-Probe'], PROBE_TIMEOUT_MS);
    const line = result.stdout.trim().split(/\r?\n/).filter(Boolean).pop() ?? '';
    try {
      const parsed = JSON.parse(line) as ProbeResult;
      probe = parsed.available ? parsed : { available: false, reason: parsed.reason ?? '本机 Excel 不可用' };
    } catch {
      probe = {
        available: false,
        reason: '探测本机 Excel 失败（未安装 Excel，或 PowerShell 无法启动 COM）',
      };
    }
    note(`[excel-bridge] 探测结果：${JSON.stringify(probe)}`);
    return probe;
  }

  async function convert(bytes: Buffer, to: string): Promise<{ ok: true; data: Buffer } | { ok: false; reason: string }> {
    if (!(to in SAVE_AS)) return { ok: false, reason: `不支持的目标格式：${to}` };
    const status = await probeExcel();
    if (!status.available) return { ok: false, reason: status.reason ?? '本机 Excel 不可用' };

    const dir = await mkdtemp(join(tmpdir(), 'excel-bridge-'));
    const input = join(dir, `in-${Date.now()}.xlsx`);
    const output = join(dir, `out-${Date.now()}.${EXTENSION[to]}`);
    try {
      await writeFile(input, bytes);
      const result = await runPowerShell(
        [script, '-In', input, '-Out', output, '-FileFormat', String(SAVE_AS[to])],
        CONVERT_TIMEOUT_MS,
      );
      const line = result.stdout.trim().split(/\r?\n/).filter(Boolean).pop() ?? '';
      let payload: { ok?: boolean; reason?: string } = {};
      try {
        payload = JSON.parse(line) as { ok?: boolean; reason?: string };
      } catch {
        /* 下面按"没有 ok"处理 */
      }
      if (!payload.ok || !existsSync(output)) {
        return { ok: false, reason: payload.reason ?? result.stderr.trim() ?? 'Excel 转换失败' };
      }
      return { ok: true, data: await readFile(output) };
    } catch (error) {
      return { ok: false, reason: String((error as Error)?.message ?? error) };
    } finally {
      await rm(dir, { recursive: true, force: true }).catch(() => undefined);
    }
  }

  const middleware: Connect.NextHandleFunction = (req, res, next) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    if (!url.pathname.startsWith('/api/bridge')) {
      next();
      return;
    }

    if (url.pathname === '/api/bridge/health') {
      void probeExcel().then((status) => sendJson(res, 200, status));
      return;
    }

    if (url.pathname === '/api/bridge/convert' && req.method === 'POST') {
      const to = (url.searchParams.get('to') ?? '').toLowerCase();
      // 串行：Excel COM 单实例，并发只会互相抢
      queue = queue
        .then(async () => {
          const bytes = await readBody(req);
          const outcome = await convert(bytes, to);
          if (!outcome.ok) {
            sendJson(res, 502, { ok: false, reason: outcome.reason });
            return;
          }
          res.statusCode = 200;
          res.setHeader('content-type', 'application/octet-stream');
          res.setHeader('cache-control', 'no-store');
          res.setHeader('content-length', outcome.data.length);
          res.end(outcome.data);
        })
        .catch((error: unknown) => {
          sendJson(res, 500, { ok: false, reason: String((error as Error)?.message ?? error) });
        });
      return;
    }

    sendJson(res, 404, { ok: false, reason: `未知的桥接口：${url.pathname}` });
  };

  return {
    name: 'excel-preview:excel-bridge',
    apply: () => options.form === 'local',
    configureServer(server) {
      server.middlewares.use(middleware);
      // vitest 也会起一个 Vite server，这里不必刷屏
      if (!process.env.VITEST) {
        note(
          '[excel-bridge] 本地转换桥已挂载：GET /api/bridge/health · POST /api/bridge/convert?to=ods|xls|xlsb|xlsx',
        );
      }
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware);
    },
  };
}
