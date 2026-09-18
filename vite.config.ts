import react from '@vitejs/plugin-react';
import { defineConfig } from 'vitest/config';

import { dropUnusedHyphenationPatterns } from './build/vite-plugin-drop-hyphenation';
import { excelBridgePlugin } from './build/vite-plugin-excel-bridge';

/**
 * 双形态（用户要求"静态、本地双形态"）：
 *
 * | 形态 | 怎么构建 | base | 转换桥 | 用途 |
 * | --- | --- | --- | --- | --- |
 * | `local`（默认） | `npm run build:local` / `npm run dev` | `/` | ✅ 有 | 本机 Excel 转换（.ods/.xls/.xlsb 导出与 .xlsb 导入） |
 * | `static` | `npm run build:static`（= `vite build --mode static`） | `./` | ❌ 无 | 丢到 GitHub/Gitee Pages 等静态托管，任何老师打开即用 |
 *
 * 用 Vite 的 `--mode` 而不是环境变量：Windows 上 `APP_FORM=static vite build` 这种写法在 npm script 里
 * 不通用（要额外装 cross-env），而 `--mode` 天生跨平台。
 *
 * `base: './'`：项目站点挂在 `https://<user>.github.io/<repo>/` 子路径下，绝对路径 `/assets/...` 会 404。
 * 形态通过 `__APP_FORM__` 编进产物（`src/shell/app-form.ts` 读取），静态形态下桥相关能力会被隐藏/置灰。
 */
export default defineConfig(({ mode }) => {
  const APP_FORM: 'local' | 'static' = mode === 'static' ? 'static' : 'local';

  return {
    plugins: [
      react(),
      dropUnusedHyphenationPatterns(),
      // 本地版才挂转换桥；静态版（--mode static）连插件都不启用
      excelBridgePlugin({ form: APP_FORM, root: __dirname, log: (message) => console.log(message) }),
    ],
    base: APP_FORM === 'static' ? './' : '/',
    define: {
      __APP_FORM__: JSON.stringify(APP_FORM),
    },
    server: {
      port: 5273,
      strictPort: false,
    },
    build: {
      target: 'es2022',
      chunkSizeWarningLimit: 4096,
      rollupOptions: {
        output: {
          /**
           * 分包：Univer / React 各自成块。
           *
           * 为什么（加载速度）：以前是**一个 6.8 MB 的入口 chunk**，改一行代码用户就要重下整块；
           * 分包后 ① 框架与业务代码分开缓存（发新版只重下业务那块）② 浏览器可并行下载
           * ③ 入口 chunk 变小，`App` 还能再走懒加载，首屏骨架不必等它。
           */
          manualChunks(id) {
            if (!id.includes('node_modules')) return undefined;
            if (/[\\/]node_modules[\\/](react|react-dom|scheduler)[\\/]/.test(id)) return 'vendor-react';
            if (id.includes('@univerjs')) return 'vendor-univer';
            return 'vendor';
          },
        },
      },
    },
    esbuild: {
      // 只丢"纯调试"的调用：console.error/warn 保留（排错要用）
      pure: ['console.log', 'console.debug', 'console.info'],
    },
    test: {
      environment: 'node',
      include: ['tests/unit/**/*.test.ts'],
      testTimeout: 30_000,
    },
  };
});
