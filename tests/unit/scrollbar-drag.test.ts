/**
 * 滚动条拖拽换算（纯函数）。
 *
 * 交互模型就是普通滚动条：**按住即可左右/上下拖，滑块跟手**（用户明确要求）。
 * 上游引擎的拖拽不响应合成指针事件、也无法在 CI 验证，所以按下后由我们接管（见模块头说明）。
 */
import { describe, expect, it } from 'vitest';

import {
  beginScrollDrag,
  isOnThumb,
  scrollForPointer,
  scrollForThumbStart,
  scrollFromTrackClick,
  thumbStartForScroll,
  thumbTravel,
  type AxisGeometry,
} from '../../src/interaction/scrollbar-drag';

/** 典型几何：轨道 700px、滑块 200px、最大滚动 1000（内容坐标单位）、轨道起点 100 */
const geo: AxisGeometry = { trackLength: 700, thumbLength: 200, limit: 1000, thumbStart: 0, trackOrigin: 100 };

describe('thumbTravel / 位置换算', () => {
  it('滑块可移动距离 = 轨道长 − 滑块长', () => {
    expect(thumbTravel(geo)).toBe(500);
  });

  it('滑块比轨道还长时不会除零（至少 1）', () => {
    expect(thumbTravel({ ...geo, thumbLength: 900 })).toBe(1);
  });

  it('滚动量 ↔ 滑块起点 互逆，且夹在两端', () => {
    expect(thumbStartForScroll(geo, 0)).toBe(0);
    expect(thumbStartForScroll(geo, 1000)).toBe(500);
    expect(thumbStartForScroll(geo, 500)).toBe(250);
    expect(thumbStartForScroll(geo, 99999)).toBe(500);
    expect(scrollForThumbStart(geo, 250)).toBe(500);
    expect(scrollForThumbStart(geo, -50)).toBe(0);
    expect(scrollForThumbStart(geo, 9999)).toBe(1000);
  });

  it('内容装得下（limit=0）时永远停在 0', () => {
    expect(thumbStartForScroll({ ...geo, limit: 0 }, 100)).toBe(0);
    expect(scrollForThumbStart({ ...geo, limit: 0 }, 100)).toBe(0);
  });
});

describe('beginScrollDrag（按下：抓滑块还是点轨道）', () => {
  it('抓滑块：记下抓取偏移、滚动量不变（不跳）', () => {
    // 当前滚动 500 → 滑块起点 250（轨道内坐标）；指针在 250+50=300 处
    const begin = beginScrollDrag(geo, 100 + 300, 500);
    expect(begin.onThumb).toBe(true);
    expect(begin.grabOffset).toBe(50);
    expect(begin.scroll, '抓滑块时不应该跳').toBe(500);
  });

  it('抓滑块带容差：偏出滑块 6px 也算抓住', () => {
    const start = thumbStartForScroll(geo, 500); // 250
    const pointerAt = 100 + start + geo.thumbLength + 5;
    expect(isOnThumb(geo, pointerAt, 500, 6)).toBe(true);
    expect(beginScrollDrag(geo, pointerAt, 500, 6).onThumb).toBe(true);
  });

  it('点轨道：滑块中心对齐指针（先跳过去），随后可继续拖', () => {
    // 点在轨道内 400 处 → 滑块起点 400-100=300 → 滚动 600
    const begin = beginScrollDrag(geo, 100 + 400, 0);
    expect(begin.onThumb).toBe(false);
    expect(begin.grabOffset).toBe(100); // 半个滑块
    expect(begin.scroll).toBe(600);
  });
});

describe('scrollForPointer（拖动中：滑块跟手）', () => {
  it('指针严格对应滑块位置：拖到轨道底部 = 滚到底', () => {
    // 抓滑块**中心**（轨道内 250+100=350 处）→ 抓取偏移 100
    const begin = beginScrollDrag(geo, 100 + 350, 500);
    expect(begin.grabOffset).toBe(100);
    // 指针到轨道最上（100）→ 滑块起点被夹到 0 → 滚到顶
    expect(scrollForPointer(geo, 100 + 100, begin.grabOffset), '拖到最上').toBe(0);
    // 指针到轨道最下（600）→ 滑块起点被夹到 500 → 滚到底
    expect(scrollForPointer(geo, 100 + 600, begin.grabOffset), '拖到最下').toBe(1000);
  });

  it('跟手 = 指针位移多少，滑块就走多少（1:1，不缩放）', () => {
    const begin = beginScrollDrag(geo, 100 + 350, 500); // 抓滑块中心，偏移 100
    // 指针往右 50px → 滑块起点 250+50=300 → 滚动 = 300/500*1000 = 600
    expect(scrollForPointer(geo, 100 + 400, begin.grabOffset)).toBe(600);
    // 再往右 50px → 700
    expect(scrollForPointer(geo, 100 + 450, begin.grabOffset)).toBe(700);
  });

  it('指针拖出轨道两端也会夹住（不会滚过头）', () => {
    const begin = beginScrollDrag(geo, 100 + 350, 500);
    expect(scrollForPointer(geo, -9999, begin.grabOffset)).toBe(0);
    expect(scrollForPointer(geo, 99999, begin.grabOffset)).toBe(1000);
  });

  it('从"点轨道跳过来的位置"继续拖也一致', () => {
    const begin = beginScrollDrag(geo, 100 + 400, 0);
    expect(begin.scroll).toBe(600);
    // 从同一点继续往右 50px → 650
    expect(scrollForPointer(geo, 100 + 450, begin.grabOffset)).toBe(700);
  });
});

describe('scrollFromTrackClick（点一下不拖）', () => {
  it('滑块中心落在指针处，并夹在两端', () => {
    expect(scrollFromTrackClick(geo, 100)).toBe(0);
    expect(scrollFromTrackClick(geo, 100 + 700)).toBe(1000);
    expect(scrollFromTrackClick(geo, 100 + 350)).toBe(500);
  });

  it('limit 为 0 时永远是 0', () => {
    expect(scrollFromTrackClick({ ...geo, limit: 0 }, 500)).toBe(0);
  });
});
