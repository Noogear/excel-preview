/**
 * 构建产物体检：体积 / 分块 / gzip 后大小 / 是否混进无关语言文件。
 *
 * 用法：
 *   node tools/measure-bundle.mjs            # 体检现有 dist/
 *   node tools/measure-bundle.mjs --json     # 只输出 JSON（给测试或 CI 用）
 *
 * 为什么需要它（用户要求"只保留 zh-CN 语言包，对加载速度进行优化"）：
 * 优化前后必须有**同一口径**的数字，否则"变快了"只是感觉。这个脚本就是那把尺子：
 *  - 入口 chunk 与 CSS 的原始/ gzip 体积；
 *  - 全部产物的总体积与文件数；
 *  - **分块清单**：Univer 的 engine-render 会把"断词词典"按语言拆成几十个懒加载分块
 *    （af/as/de-1901/hu/th/…），本项目是纯表格工具，用不到它们 —— 数量必须为 0。
 */
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { extname, join, resolve } from 'node:path';
import { gzipSync } from 'node:zlib';

const DIST = resolve(process.argv.includes('--dist') ? process.argv[process.argv.indexOf('--dist') + 1] : 'dist');
const JSON_ONLY = process.argv.includes('--json');

/** 断词词典分块的文件名特征（engine-render 的 hyphenation patterns） */
const HYPHENATION_HINT = /^(af|as|be|bg|bn|ca|cop|cs|cu|cy|da|de-1901|de-1996|de-ch-1901|el-monoton|el-polyton|en-gb|es|et|eu|fi|fr|ga|grc|gu|hi|hr|hsb|hu|hy|ia|id|is|it|ka|km|kn|la|lt|lv|ml|mn-cyrl|mr|nb|nl|nn|no|or|pa|pi|pl|pt|rm|ro|ru|sa|sk|sl|sr-cyrl|sv|ta|te|th|tk|tr|uk|zh-latn-pinyin)(-[A-Za-z0-9_-]+)?\.js$/;

function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    const info = statSync(full);
    if (info.isDirectory()) walk(full, out);
    else out.push({ path: full.replace(DIST + '\\', '').replace(DIST + '/', ''), size: info.size });
  }
  return out;
}

function gzipSize(path) {
  try {
    return gzipSync(readFileSync(path), { level: 9 }).length;
  } catch {
    return null;
  }
}

const indexHtml = readFileSync(join(DIST, 'index.html'), 'utf8');
const entryAssets = [...indexHtml.matchAll(/(?:src|href)="\.?\/?(assets\/[^"]+)"/g)].map((match) => match[1]);

const files = walk(DIST);
const js = files.filter((file) => extname(file.path) === '.js');
const css = files.filter((file) => extname(file.path) === '.css');
const totalBytes = files.reduce((sum, file) => sum + file.size, 0);

const entry = entryAssets
  .filter((path) => path.endsWith('.js'))
  .map((path) => {
    const found = files.find((file) => file.path.replace(/\\/g, '/') === path);
    if (!found) return null;
    return { path, size: found.size, gzip: gzipSize(join(DIST, found.path)) };
  })
  .filter(Boolean);

const entryCss = entryAssets
  .filter((path) => path.endsWith('.css'))
  .map((path) => {
    const found = files.find((file) => file.path.replace(/\\/g, '/') === path);
    if (!found) return null;
    return { path, size: found.size, gzip: gzipSize(join(DIST, found.path)) };
  })
  .filter(Boolean);

/** 懒加载分块（不在 index.html 里直接引用、也不是入口的 JS） */
const lazyChunks = js.filter((file) => !entryAssets.includes(file.path.replace(/\\/g, '/')));
const hyphenation = lazyChunks.filter((file) => HYPHENATION_HINT.test(file.path.split(/[\\/]/).pop() ?? ''));

/**
 * 分组（口径要写清楚，否则"变快了"说不明白）：
 *  - `entry`：HTML 直接引用的（首屏骨架就靠它，越小越好）；
 *  - `appCritical`：入口 + 应用主体 + 框架/引擎分包 —— 到"表格能用"为止必须下完的；
 *  - `lazy`：用到才下的（CSV/ODS/XLS 解析器等）。
 */
const appCritical = [...entry, ...entryCss, ...lazyChunks.filter((file) => /^(App|vendor)/.test(file.path.split(/[\\/]/).pop() ?? ''))];
const lazy = lazyChunks.filter((file) => !appCritical.some((item) => item.path === file.path));
const gzipSum = (items) => items.reduce((sum, item) => sum + (gzipSize(join(DIST, item.path)) ?? 0), 0);

const report = {
  dist: DIST,
  totalBytes,
  totalFiles: files.length,
  jsChunks: js.length,
  cssFiles: css.length,
  entry,
  entryCss,
  /** 首屏骨架就绪所需的字节（入口 JS + 入口 CSS，gzip） */
  firstLoadGzip: [...entry, ...entryCss].reduce((sum, item) => sum + (item.gzip ?? 0), 0),
  /** 到"表格可用"为止必须下完的字节（gzip） */
  appCriticalGzip: gzipSum(appCritical),
  appCritical: appCritical.map((item) => ({ path: item.path, size: item.size, gzip: gzipSize(join(DIST, item.path)) })),
  lazyGzip: gzipSum(lazy),
  lazy: lazy.map((item) => ({ path: item.path, size: item.size, gzip: gzipSize(join(DIST, item.path)) })),
  largest: files
    .slice()
    .sort((a, b) => b.size - a.size)
    .slice(0, 8)
    .map((file) => ({ path: file.path, size: file.size, gzip: gzipSize(join(DIST, file.path)) })),
  hyphenationChunks: hyphenation.map((file) => file.path),
};

if (JSON_ONLY) {
  console.log(JSON.stringify(report, null, 2));
} else {
  const mb = (bytes) => `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  const kb = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;
  console.log('── 构建产物体检 ─────────────────────────────────────────────');
  console.log(`产物总量      ${mb(totalBytes)}（${files.length} 个文件）`);
  console.log(`JS 分块       ${js.length} 个；CSS ${css.length} 个`);
  for (const item of entry) console.log(`入口 JS       ${kb(item.size)}（gzip ${kb(item.gzip ?? 0)}）  ${item.path}`);
  for (const item of entryCss) console.log(`入口 CSS      ${kb(item.size)}（gzip ${kb(item.gzip ?? 0)}）  ${item.path}`);
  console.log(`首屏骨架字节  ${kb(report.firstLoadGzip)}（gzip，入口 JS + CSS —— 这段一到就能画出骨架）`);
  console.log(`到可用字节    ${kb(report.appCriticalGzip)}（gzip，入口 + 应用 + 框架/引擎分包）`);
  console.log(`惰性分块字节  ${kb(report.lazyGzip)}（gzip，用到才下）`);
  console.log(`断词词典分块  ${hyphenation.length} 个${hyphenation.length ? '  ← 应为 0（本项目用不到）' : ' ✓'}`);
  if (hyphenation.length) {
    for (const file of hyphenation.slice(0, 8)) console.log(`   · ${file.path}`);
    if (hyphenation.length > 8) console.log(`   · …还有 ${hyphenation.length - 8} 个`);
  }
  console.log('最大的 8 个文件：');
  for (const file of report.largest) console.log(`   ${kb(file.size).padStart(9)}（gzip ${kb(file.gzip ?? 0).padStart(9)}）  ${file.path}`);
  console.log('─────────────────────────────────────────────────────────────');
}

// 体检失败（混进断词词典）时给非 0 退出码：CI / 测试可以直接用
if (hyphenation.length > 0 && !JSON_ONLY && process.argv.includes('--strict')) process.exit(1);
