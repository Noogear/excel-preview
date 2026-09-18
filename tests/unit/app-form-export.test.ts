/**
 * 双形态与导出格式的单测（纯函数，node 环境）。
 *
 * 覆盖用户这一轮的三点要求里**能被单测钉住**的部分：
 *  - "只保留 zh-CN 语言包"：语言相关的表在构建期被剔除这件事由 `tools/measure-bundle.mjs --strict`
 *    与 e2e 的产物断言守卫（见 `tests/e2e/static-form.spec.ts`），这里只钉住"形态/文件名"这类纯逻辑；
 *  - "双形态"：形态常量、桥的可用性判断、桥不可用时的文案、目标格式 → 文件名；
 *  - "导出"：CSV 的引号与换行规则、BOM 编码（学号前导零与中文不乱码都靠它）。
 */
import { describe, expect, it } from 'vitest';

import { buildDelimited, csvField, encodeDelimited } from '../../src/exporter/csv-export';
import {
  APP_FORM,
  BRIDGE_EXTENSION,
  bridgeUnavailableHint,
  bridgedFileName,
  initialBridgeStatus,
  isLocalForm,
  type BridgeStatus,
} from '../../src/shell/app-form';

describe('形态（静态 / 本地）', () => {
  it('单测环境里形态默认是 local（vite define 未注入时的兜底）', () => {
    expect(['local', 'static']).toContain(APP_FORM);
    expect(isLocalForm).toBe(APP_FORM === 'local');
  });

  it('initialBridgeStatus：本地版"探测中"，静态版直接"不可用 + 说明原因"', () => {
    const status = initialBridgeStatus();
    if (isLocalForm) {
      expect(status.checking).toBe(true);
      expect(status.formats).toEqual(['xlsx', 'ods', 'xls', 'xlsb']);
    } else {
      expect(status.checking).toBe(false);
      expect(status.available).toBe(false);
      expect(status.reason).toContain('本地版');
      expect(status.formats).toEqual([]);
    }
  });

  it('bridgeUnavailableHint：把"为什么不能用"说清楚，而不是一句"不可用"', () => {
    const unavailable: BridgeStatus = { checking: false, available: false, formats: [], reason: '本机没装 Excel' };
    expect(bridgeUnavailableHint(unavailable)).toBe('本机没装 Excel');
    const noReason: BridgeStatus = { checking: false, available: false, formats: [] };
    expect(bridgeUnavailableHint(noReason).length).toBeGreaterThan(6);
  });

  it('目标格式 → 扩展名与文件名（沿用"原名-已编辑"的约定）', () => {
    expect(BRIDGE_EXTENSION).toEqual({ xlsx: '.xlsx', ods: '.ods', xls: '.xls', xlsb: '.xlsb' });
    expect(bridgedFileName('成绩表.xlsx', 'ods')).toBe('成绩表-已编辑.ods');
    expect(bridgedFileName('fixture-styles.xls', 'xlsb')).toBe('fixture-styles-已编辑.xlsb');
    expect(bridgedFileName('没有扩展名', 'xls')).toBe('没有扩展名-已编辑.xls');
  });
});

describe('CSV 导出（自研）', () => {
  it('csvField：含分隔符/引号/换行才加引号，内部引号翻倍', () => {
    expect(csvField('张三')).toBe('张三');
    expect(csvField('张三,李四')).toBe('"张三,李四"');
    expect(csvField('他说"你好"')).toBe('"他说""你好"""');
    expect(csvField('第一行\n第二行')).toBe('"第一行\n第二行"');
    // 制表符当分隔符时，逗号不再是"需要引号"的字符
    expect(csvField('张三,李四', '\t')).toBe('张三,李四');
    expect(csvField('甲\t乙', '\t')).toBe('"甲\t乙"');
  });

  it('buildDelimited：行尾统一 CRLF，末尾也带行尾（与 Excel 导出一致）', () => {
    const text = buildDelimited(
      [
        ['学号', '姓名'],
        ['007', '张三'],
      ],
      ',',
    );
    expect(text).toBe('学号,姓名\r\n007,张三\r\n');
    expect(buildDelimited([['a', 'b']], '\t')).toBe('a\tb\r\n');
  });

  it('encodeDelimited：默认带 UTF-8 BOM（Excel 双击打开中文不乱码）', () => {
    const withBom = encodeDelimited('学号');
    expect([...withBom.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder('utf-8').decode(withBom)).toBe('学号');
    const withoutBom = encodeDelimited('学号', false);
    expect(withoutBom.length).toBe(withBom.length - 3);
  });
});
