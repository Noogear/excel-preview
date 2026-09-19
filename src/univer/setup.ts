/**
 * Univer 引导（P0）：用官方 preset 以最少代码把表格跑起来，
 * 并把 univerAPI 暴露到 window，供 Playwright 断言与人工调试。
 */
import { IConfigService, LocaleType, type IWorkbookData } from '@univerjs/core';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note';
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';
import { createUniver, merge } from '@univerjs/presets';
import { defaultTheme } from '@univerjs/themes';

import '@univerjs/preset-sheets-core/lib/index.css';

import { log } from '../p0/log';

export interface UniverBoot {
  univerAPI: ReturnType<typeof createUniver>['univerAPI'];
  univer: ReturnType<typeof createUniver>['univer'];
  dispose: () => void;
}

/**
 * 我们自己补的中文文案。
 *
 * 起因（用户实测反馈）：拖动日期单元格时弹出 `sheets-ui.info.error` /
 * `sheets-ui.info.forceStringInfo` 两个**原始 i18n key**。查证后确认这是 Univer 上游的漏配——
 * 代码里 `t('sheets-ui.info.forceStringInfo')` 有调用，但**所有语言包（含 en-US）都没有这个 key**，
 * 所以任何语言下都只会显示 key 本身。我们在语言合并时补上，用户至少能读到人话。
 *
 * （触发这次提醒的根因是"日期被降级成字符串再写回"，已在本仓的 snapshot 归一化里修掉；
 *   这里补文案是兜底：以后任何路径再弹出这个提醒，也不会是一串 key。）
 */
const EXTRA_ZH_CN = {
  'sheets-ui': {
    info: {
      error: '提示',
      forceStringInfo: '这个单元格是文本格式，但内容看起来是数字，已按文本处理（不参与计算）。',
    },
  },
};

export function bootUniver(containerId: string): UniverBoot {
  const started = performance.now();

  const { univerAPI, univer } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: merge({}, UniverPresetSheetsCoreZhCN, EXTRA_ZH_CN),
    },
    theme: defaultTheme,
    presets: [
      UniverSheetsCorePreset({
        container: containerId,
        // 关掉 Univer 自带的 ribbon / 格式工具栏：
        // 本产品只允许编辑内容、不允许改格式，留着格式按钮会出现"点了没反应"的矛盾体验。
        // 公式栏保留（编辑内容是核心能力），底部 footer 保留（工作表标签 + 缩放）。
        // 只关掉 ribbon / 格式工具栏：本产品只允许编辑内容、不允许改格式，
        // 留着格式按钮会出现"点了没反应"的矛盾体验。
        // header 保留（公式栏在里面，编辑内容是核心能力），footer 保留（工作表标签 + 缩放）。
        header: true,
        toolbar: false,
        formulaBar: true,
        // footer 的类型是对象或 false：保留工作表标签 + 缩放 + 统计
        // addSheetButtonConfig.show = false：**不要子表创建功能**（用户明确要求）
        footer: {
          sheetBar: true,
          zoomSlider: true,
          statisticBar: true,
          addSheetButtonConfig: { show: false },
        },
        // 关掉 Univer 自带右键菜单：①它的菜单项大多是格式/结构操作，与本产品"只允许改内容"的
        // 约束直接矛盾（点了没反应）；②它和本项目自己的右键菜单（⑦）会同时弹出、且层级更高，
        // 实测把自家菜单项完全盖住导致用户点不动。关掉后右键只出自家菜单。
        contextMenu: false,
      }),
      // ---- P1 保真所需的渲染能力 ----
      // 条件格式四类（高亮/色阶/数据条/图标集）
      UniverSheetsConditionalFormattingPreset(),
      // 数据验证（下拉列表、数值/日期范围等）
      UniverSheetsDataValidationPreset(),
      // 超链接（外部链接/单元格跳转）
      UniverSheetsHyperLinkPreset(),
      // 批注（legacy note）
      UniverSheetsNotePreset(),
      // 浮动图片（drawing 锚点）
      UniverSheetsDrawingPreset(),
      // 自动筛选（Excel 表格自带筛选按钮）
      UniverSheetsFilterPreset(),
    ],
  });

  disableForceStringAlert(univer);
  // 记下"这个实例"的配置服务引用（dispose 时按身份摘掉模块级引用，见下面的 dispose）
  const bootedConfigService = configServiceRef;

  // 启动打点：配合入口的 `app:bundle-loaded` / `app:chunk-loaded` / `app:mounted`
  // 就能量出"骨架 → 应用 → 表格可用"三段耗时（e2e 的启动预算断言读这些 mark）
  performance.mark('univer:booted');
  log('univer:booted', { ms: Math.round(performance.now() - started) });

  return {
    univerAPI,
    univer,
    dispose: () => {
      /**
       * 摘掉模块级引用再拆实例：`configServiceRef` 是模块级的，而配置服务属于**这一个** Univer 实例。
       * 不清的话，卸载/HMR 之后模块仍强引用那个已 dispose 的 `IConfigService`
       * （连带它的配置 Map 与 `_configChanged$` Subject 及其订阅者闭包），要等下一次 `bootUniver` 才被覆盖。
       * 按身份比对，避免把"下一个实例刚写进去的引用"误清掉。
       */
      if (configServiceRef === bootedConfigService) configServiceRef = null;
      univer.dispose();
    },
  };
}

/**
 * 关掉"文本格式但看起来是数字"的**弹窗**提醒（保留单元格角标）。
 *
 * 触发条件（读上游源码确认）：当前激活单元格 `t === FORCE_STRING | STRING`
 * 且 `isRealNum(v)` 成立时，只要**选中/移动到该单元格**就弹一个悬浮提示。
 *
 * 为什么本产品要关：①本工具**锁死格式**，用户没有"把文本改成数字"这条路，
 * 提醒里说的补救动作根本做不到，属于纯噪音；②班主任场景里学号/身份证/手机号
 * 这类"文本存数字"极其常见，留着弹窗等于每点一格弹一次；
 * ③我们在 snapshot 归一化里已经修掉了真正的元凶（日期被降级成字符串再写回）。
 * 信息不丢：只关弹窗，`disableForceStringMark` 保持默认 false，单元格左上角的
 * 小三角（ForceString 角标）照常显示，需要的人依然看得到。
 */
type ConfigServiceLike = {
  getConfig: (id: string) => unknown;
  setConfig: (id: string, value: unknown, options?: { merge?: boolean }) => void;
};

const SHEETS_UI_CONFIG_KEY = 'sheets-ui.config';

/** 引导时抓到的配置服务（配置服务是全局单例，之后重复写配置要用它） */
let configServiceRef: ConfigServiceLike | null = null;

/**
 * 把"关弹窗"这条配置写进 `sheets-ui.config`。
 *
 * 实测坑：`createUniver()` 返回时插件生命周期**还没跑完**，sheets-ui 插件随后会用
 * `setConfig(key, config)`（非 merge）把整个 `sheets-ui.config` 覆盖成它自己的配置对象，
 * 于是"引导时写一次"会被静默抹掉（配置读回来只剩 formulaBar/statusBarStatistic）。
 * 所以除了引导时写，**每次新建工作簿后再补一次**（见 `loadWorkbook`）。
 * 也不能靠 `subscribeConfigValue$` 监听回写：上游那个实现用 `hasOwnProperty` 判断 Map 的键，
 * 永远不会触发。
 */
function enforceForceStringAlertOff(): void {
  const configService = configServiceRef;
  if (!configService) return;
  try {
    const current = configService.getConfig(SHEETS_UI_CONFIG_KEY) as { disableForceStringAlert?: boolean } | null;
    if (current?.disableForceStringAlert === true) return;
    configService.setConfig(SHEETS_UI_CONFIG_KEY, { disableForceStringAlert: true }, { merge: true });
    log('univer:force-string-alert-disabled', {});
  } catch (error) {
    // 关不掉也不该影响启动：本地化文案已补齐，最差情况只是多一个能看懂的弹窗。
    log('univer:force-string-alert-disable-failed', { error: String(error) });
  }
}

function disableForceStringAlert(univer: ReturnType<typeof createUniver>['univer']): void {
  try {
    type Injector = { get: (token: unknown) => unknown } | undefined;
    const injector = (univer as unknown as { __getInjector?: () => Injector }).__getInjector?.();
    const configService = injector?.get(IConfigService) as ConfigServiceLike | undefined;
    if (!configService) return;
    configServiceRef = configService;
    enforceForceStringAlertOff();
  } catch (error) {
    log('univer:force-string-alert-disable-failed', { error: String(error) });
  }
}

export function loadWorkbook(univerAPI: UniverBoot['univerAPI'], data: IWorkbookData): void {
  const started = performance.now();
  univerAPI.createWorkbook(data);
  // 工作簿创建会触发插件的 ready 阶段，插件的 `sheets-ui.config` 注册可能在这之后落地，
  // 把引导时写的"关弹窗"覆盖掉 → 这里补写一次（幂等）。
  enforceForceStringAlertOff();
  log('univer:workbook-created', { ms: Math.round(performance.now() - started), id: data.id });
}
