/**
 * 构建期插件：把 Univer 引擎里**本项目用不到的断词词典**从产物中剔除。
 *
 * 背景（用户要求"只保留 zh-CN 语言包，对加载速度进行优化"）：
 * `@univerjs/engine-render` 里有一张 `PATTERN_LOADERS` 表，把 69 种语言的**断词（hyphenation）词典**
 * 写成动态 import（`af`/`de-1901`/`hu`/`th`/`uk`/`zh-latn-pinyin`…）。这些词典只服务于
 * **文档（Word）排版**的分行算法；本项目是纯表格工具，从不创建文档排版器，也就永远不会加载它们。
 * 但打包器看到动态 import 就会为每种语言产出一个分块 —— 实测 **69 个分块、约 4.3 MB**，
 * 构建日志里满屏 `de-1901 / hu / th` 之类的"语言文件"，看起来像语言包，其实是断词词典。
 *
 * 做法：只在构建时把那张表替换成空对象（`loadPattern()` 里本来就写着 `if (!loader) return;`，
 * 空表是安全的；需要断词时它只是不加载，不会抛错）。
 *
 * 为什么不用 `resolve.alias` / `manualChunks`：前者要逐个别名 69 个相对路径，后者只能"合并"、
 * 不能去掉字节。这里直接把动态 import 从图里摘掉，才是真的不产分块。
 */
import type { Plugin } from 'vite';

const ENGINE_RENDER_INDEX = /[\\/]node_modules[\\/]@univerjs[\\/]engine-render[\\/]lib[\\/](es|cjs)[\\/]index\.js$/;

export function dropUnusedHyphenationPatterns(): Plugin {
  let dropped = 0;
  return {
    name: 'excel-preview:drop-unused-hyphenation-patterns',
    // 只在**构建**时介入：开发态由 esbuild 预打包按需提供，且我们从不触发放这类词典的加载路径
    apply: 'build',
    // 必须早于 Vite 内置的依赖预打包/转换
    enforce: 'pre',
    transform(code, id) {
      if (!ENGINE_RENDER_INDEX.test(id)) return null;
      if (!code.includes('PATTERN_LOADERS')) return null;
      const next = code.replace(/const PATTERN_LOADERS = \{[\s\S]*?\n\};/, 'const PATTERN_LOADERS = {};');
      if (next === code) return null;
      dropped += 1;
      this.warn(
        '[excel-preview] 已剔除 engine-render 的断词词典分块（69 种语言，本项目只用表格、不做文档排版）；' +
          '如需恢复，删掉 vite.config.ts 里的 dropUnusedHyphenationPatterns()。',
      );
      return { code: next, map: null };
    },
    buildEnd() {
      if (dropped === 0) {
        this.warn(
          '[excel-preview] 没有匹配到 engine-render 的 PATTERN_LOADERS —— 上游包结构可能变了，' +
            '请用 `node tools/measure-bundle.mjs` 确认产物里是否又混进了语言/断词分块。',
        );
      }
    },
  };
}
