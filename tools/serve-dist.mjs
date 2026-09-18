/**
 * 把 `dist/` 当静态站点起来（默认挂在 `/repo/` 子路径下，模拟 GitHub / Gitee Pages 的项目站点）。
 *
 * 两个用途：
 *  ① 人手验证：`node tools/serve-dist.mjs`（默认 5399 端口、前缀 /repo/，对应 `npm run build:static`）；
 *  ② e2e：`tests/e2e/static-form.spec.ts` 直接 `import { startStaticServer }` 起一个临时服务，
 *     验证"静态形态"能引导、能打开表格、能导出，以及桥相关能力被正确降级。
 *
 * 只依赖 node 内置模块（不引第三方静态服务器），保证离线可用。
 */
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize, resolve } from 'node:path';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.woff2': 'font/woff2',
  '.ico': 'image/x-icon',
  '.map': 'application/json; charset=utf-8',
};

/**
 * 起一个只服务 `distDir` 的静态服务器。
 * @param {{ distDir?: string, prefix?: string, port?: number }} [options]
 * @returns {Promise<{ url: string, port: number, close: () => Promise<void> }>}
 */
export async function startStaticServer(options = {}) {
  const distDir = resolve(options.distDir ?? join(process.cwd(), 'dist'));
  const prefix = options.prefix ?? '/repo/';
  const wanted = options.port ?? 5399;

  const server = createServer(async (req, res) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    let pathname = decodeURIComponent(url.pathname);
    if (!pathname.startsWith(prefix)) {
      res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      res.end(`not found（本站点挂在 ${prefix} 下）`);
      return;
    }
    pathname = pathname.slice(prefix.length);
    if (pathname === '' || pathname.endsWith('/')) pathname += 'index.html';
    const file = join(distDir, normalize(pathname).replace(/^([.][.][/\\])+/, ''));
    try {
      const info = await stat(file);
      if (!info.isFile()) throw new Error('not a file');
      const body = await readFile(file);
      res.writeHead(200, {
        'content-type': MIME[extname(file).toLowerCase()] ?? 'application/octet-stream',
        'content-length': body.length,
        // 刻意**不加** `cache-control: no-store`：Chromium 对 no-store 的文档不上报
        // first-contentful-paint，会把"启动预算"用例的度量口径搞没（踩过）。
        // 本地验证场景本来也不需要缓存控制。
        'cache-control': 'no-cache',
      });
      res.end(body);
    } catch {
      res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      res.end('not found');
    }
  });

  await new Promise((fulfil, reject) => {
    server.once('error', reject);
    server.listen(wanted, '127.0.0.1', fulfil);
  });
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : wanted;
  return {
    url: `http://127.0.0.1:${port}${prefix}`,
    port,
    close: () =>
      new Promise((fulfil) => {
        server.close(() => fulfil());
      }),
  };
}

// 直接 `node tools/serve-dist.mjs` 时当命令行工具用
if (process.argv[1] && resolve(process.argv[1]).endsWith('serve-dist.mjs')) {
  const prefix = process.argv[2] ?? '/repo/';
  const port = Number(process.argv[3] ?? 5399);
  const handle = await startStaticServer({ prefix, port });
  console.log(`[serve-dist] ${handle.url} → ${resolve('dist')}`);
  console.log('[serve-dist] Ctrl+C 结束');
}
