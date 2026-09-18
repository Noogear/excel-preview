/**
 * 跑 Node 侧性能基准：esbuild 打包 tools/bench-entry.ts → node 执行。
 *
 *   node tools/bench-parse.mjs                 # 出数字
 *   BENCH_PROFILE=1 node tools/bench-parse.mjs # 额外产出 V8 CPU 剖面（.bench/*.cpuprofile）
 *   BENCH_FILE=bench-medium.xlsx node tools/bench-parse.mjs
 *
 * 为什么绕一层打包：基准要吃 TS 源码，而 vitest 的 worker 进程抓不到 `--cpu-prof` 剖面。
 * 打包成一个普通 ESM 后用 node 直接跑，源码与剖面就都在同一个进程里了。
 */
import { build } from 'esbuild';
import { mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = join(ROOT, '.bench');
const BUNDLE = join(OUT_DIR, 'bench-entry.mjs');
const profile = process.env.BENCH_PROFILE === '1';

rmSync(OUT_DIR, { recursive: true, force: true });
mkdirSync(OUT_DIR, { recursive: true });

await build({
  entryPoints: [join(ROOT, 'tools', 'bench-entry.ts')],
  outfile: BUNDLE,
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node20',
  logLevel: 'warning',
});

const args = ['--expose-gc', BUNDLE];
if (profile) args.unshift('--cpu-prof', '--cpu-prof-dir=.bench', '--cpu-prof-interval=100');

// 1) 先跑内存探针进程：它只量"解析模型占多少"，跑完把结果写进 .bench/mem.json。
//    必须单独进程——同进程里先前解析留下的死引用会被保守栈扫描当成活的，读数会乱跳。
const probe = spawnSync(process.execPath, ['--expose-gc', BUNDLE], {
  cwd: ROOT,
  stdio: 'inherit',
  env: { ...process.env, MEM_PROBE: '1' },
});
if (probe.status !== 0) {
  console.warn('（内存探针进程未正常结束，模型占用一栏会缺失）');
}

// 2) 再跑主进程：计时 + 常驻工作集
const started = Date.now();
const result = spawnSync(process.execPath, args, { cwd: ROOT, stdio: 'inherit' });
console.log(`（进程总耗时 ${((Date.now() - started) / 1000).toFixed(1)}s，退出码 ${result.status}）`);

if (profile) {
  const files = readdirSync(OUT_DIR).filter((name) => name.endsWith('.cpuprofile'));
  for (const file of files) {
    const profileJson = JSON.parse(readFileSync(join(OUT_DIR, file), 'utf8'));
    const selfTime = new Map();
    const byId = new Map(profileJson.nodes.map((node) => [node.id, node]));
    // 按样本自身的时长（timeDeltas，单位 µs）归因，避免"样本数 × 固定间隔"的估算误差
    const samples = profileJson.samples ?? [];
    const deltas = profileJson.timeDeltas ?? [];
    let totalUs = 0;
    samples.forEach((id, index) => {
      const us = deltas[index] ?? 0;
      totalUs += us;
      const node = byId.get(id);
      if (!node) return;
      const { functionName, url, lineNumber } = node.callFrame;
      const key = `${functionName || '(anonymous)'}  ${(url || '').replace(/^.*[\\/]/, '')}:${lineNumber + 1}`;
      selfTime.set(key, (selfTime.get(key) ?? 0) + us);
    });
    const total = totalUs || 1;
    const top = [...selfTime.entries()].sort((a, b) => b[1] - a[1]).slice(0, 14);
    console.log(`\n┌─ CPU 剖面（自身耗时 Top，采样总时长 ${(totalUs / 1000).toFixed(0)}ms）`);
    for (const [name, us] of top) {
      console.log(`│ ${((us / total) * 100).toFixed(1).padStart(5)}%  ${(us / 1000).toFixed(0).padStart(6)}ms  ${name}`);
    }
    console.log('└──────────────────────────────────────────────────────');
  }
}
