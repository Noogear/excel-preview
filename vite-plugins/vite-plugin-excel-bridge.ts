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
/**
 * 单次转换超时（毫秒）。
 *
 * 为什么从 60 秒放宽到 180 秒（实测踩到的坑）：Excel 冷启动 + 大表 SaveAs 在慢机器上会超过 60 秒
 * （实测一次 60.36 秒，**刚好越过旧上限**）。而超时是"强杀"路径：杀了 PowerShell，
 * 脚本里的 `finally { $excel.Quit() }` 就没机会执行 → 留下一个看不见的 EXCEL.EXE。
 * 于是"超时 → 残留 → 机器更慢 → 更容易超时"形成正反馈（实测攒到过 57 个残留进程）。
 * 现在双管齐下：上限放宽 + 超时按 PID 精确收拾（见 killExcelByPid）。
 */
const CONVERT_TIMEOUT_MS = 180_000;
/** 健康探测超时（毫秒）：冷启动 Excel 也可能十几秒，给足 */
const PROBE_TIMEOUT_MS = 60_000;
/** 探测**失败**结果的缓存时长：只挡住"连点两下"的重复探测，不至于把瞬时失败钉死一整个会话 */
const NEGATIVE_PROBE_TTL_MS = 5_000;

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

/**
 * 我们这次启动的 Excel 进程登记表（PID）。
 *
 * 用途：超时强杀 PowerShell 时，`finally` 里的 `Quit()` 跑不到，必须由中间件按 PID 收拾；
 * 另外服务器退出时（`httpServer` 的 close）也用它兜一次底。
 * 只登记"我们确认新起出来"的 PID（脚本用创建前后差集判断），所以不会误杀用户自己开的 Excel。
 */
const spawnedExcelPids = new Set<number>();

/** 按 PID 强杀一个 Excel（超时路径 / 退出兜底）。杀不掉就忽略——不能让它影响请求结果 */
function killExcelByPid(pid: number): void {
  if (!Number.isInteger(pid) || pid <= 0) return;
  try {
    const killer = spawn('taskkill', ['/PID', String(pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' });
    killer.on('error', () => undefined);
    killer.on('close', () => undefined);
  } catch {
    /* 杀不掉就算了：脚本侧还有一次 Stop-Process 兜底 */
  }
  spawnedExcelPids.delete(pid);
}

/** 退出兜底：把我们起过的 Excel 全部收掉（正常情况下它们早就被脚本 Quit 了） */
function killAllSpawnedExcel(): void {
  for (const pid of [...spawnedExcelPids]) killExcelByPid(pid);
}

function runPowerShell(
  args: string[],
  timeoutMs: number,
): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((fulfil) => {
    const child = spawn('pwsh', ['-NoProfile', '-NonInteractive', '-File', ...args], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    /** 脚本起来后会把 Excel 的 PID 作为**第一行**报出来（见 excel-bridge.ps1 的说明） */
    let excelPid: number | null = null;
    let settled = false;
    const capturePid = (chunk: string): void => {
      if (excelPid !== null) return;
      for (const line of chunk.split(/\r?\n/)) {
        if (!line.includes('excel-pid')) continue;
        try {
          const parsed = JSON.parse(line) as { bridge?: string; pid?: number };
          if (parsed.bridge === 'excel-pid' && typeof parsed.pid === 'number') {
            excelPid = parsed.pid;
            spawnedExcelPids.add(parsed.pid);
          }
        } catch {
          /* 半行/脏行，等下一块数据 */
        }
      }
    };
    const timer = setTimeout(() => {
      settled = true;
      child.kill();
      /**
       * **关键**：进程被杀 = 脚本的 `finally` 不会执行 → Excel 不会被 Quit。
       * 所以这里按 PID 精确补一刀，否则每次超时都留下一个看不见的 EXCEL.EXE。
       */
      if (excelPid !== null) killExcelByPid(excelPid);
      fulfil({ code: -1, stdout, stderr: `${stderr}\n[bridge] 超时 ${timeoutMs}ms，已结束（并已清理 Excel${excelPid ? ` #${excelPid}` : ''}）` });
    }, timeoutMs);
    child.stdout?.on('data', (chunk) => {
      const text = String(chunk);
      stdout += text;
      capturePid(text);
    });
    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });
    child.on('error', (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      fulfil({ code: -1, stdout, stderr: `${stderr}\n${String(error)}` });
    });
    child.on('close', (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      // 正常结束：脚本自己已经 Quit 了，这里只把登记表清掉
      if (excelPid !== null) spawnedExcelPids.delete(excelPid);
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
  /**
   * 探测结果缓存。
   *
   * 成功结果**长期缓存**（"本机装没装 Excel"不会中途变）；
   * 失败结果只缓存 `NEGATIVE_PROBE_TTL_MS` 一小会儿 —— 实测踩过：探测会因为"Excel 正忙 / 冷启动慢"
   * 这类**瞬时**原因失败，而以前失败也被永久缓存，于是这一整个开发服务器会话里桥都显示"不可用"，
   * 只能重启服务才能恢复（本次开发中就撞上一次）。给个短 TTL，让它自己缓过来。
   */
  let probe: ProbeResult | null = null;
  let probeFailedAt = 0;
  /** Excel COM 不能并发：转换串行排队 */
  let queue: Promise<unknown> = Promise.resolve();

  async function probeExcel(): Promise<ProbeResult> {
    if (probe?.available) return probe;
    if (probe && !probe.available && Date.now() - probeFailedAt < NEGATIVE_PROBE_TTL_MS) return probe;
    if (!existsSync(script)) {
      probe = { available: false, reason: `找不到转换脚本 ${script}` };
      probeFailedAt = Date.now();
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
    if (!probe.available) probeFailedAt = Date.now();
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
      /**
       * 服务器退出时兜一次底：把我们起过、但还活着的 Excel 收掉。
       * （正常路径下脚本自己已经 Quit 了，这里只防"超时强杀"那种漏网。）
       */
      server.httpServer?.on('close', killAllSpawnedExcel);
      // vitest 也会起一个 Vite server，这里不必刷屏
      if (!process.env.VITEST) {
        note(
          '[excel-bridge] 本地转换桥已挂载：GET /api/bridge/health · POST /api/bridge/convert?to=ods|xls|xlsb|xlsx',
        );
      }
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware);
      server.httpServer?.on('close', killAllSpawnedExcel);
    },
  };
}
