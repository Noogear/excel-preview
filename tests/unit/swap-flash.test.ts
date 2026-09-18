/**
 * 互换提醒框的坐标换算（纯函数）与浮层生命周期单测。
 *
 * 背景：提醒框要"画在正确的单元格上、逐渐透明消失、且不吃鼠标事件"。
 * 这里守住最容易错的两块：①单元格内容坐标 → 画布坐标的换算（含滚动/缩放/表头偏移）；
 * ②浮层的 DOM 生命周期（show → 自动消失 → hide/dispose 收干净）。
 * 浮层用注入的假 document / 假 rAF 驱动，不依赖 jsdom。
 */
import { describe, expect, it } from 'vitest';

import {
  DEFAULT_FLASH_DURATION_MS,
  FLASH_BOX_CLASS,
  FLASH_LAYER_CLASS,
  computeFlashRect,
  createSwapFlashOverlay,
  type FlashCellRect,
  type FlashViewport,
} from '../../src/interaction/swap-flash';

/* -------------------------------------------------------------------------- */
/* 坐标换算                                                                    */
/* -------------------------------------------------------------------------- */

/** 造一个"第 0 列从 x=46 开始、宽 100、高 30"的表（46 = 行头宽度） */
function viewport(overrides: Partial<FlashViewport> = {}): FlashViewport {
  const cellRect = (row: number, col: number): FlashCellRect => ({
    startX: 46 + col * 100,
    startY: 24 + row * 30,
    endX: 46 + (col + 1) * 100,
    endY: 24 + (row + 1) * 30,
  });
  return {
    cellRect,
    scroll: { x: 0, y: 0 },
    scale: { x: 1, y: 1 },
    ...overrides,
  };
}

describe('computeFlashRect', () => {
  it('无滚动、无缩放：表头偏移直接体现在位置里', () => {
    expect(computeFlashRect(viewport(), 0, 0)).toEqual({ left: 46, top: 24, width: 100, height: 30 });
    expect(computeFlashRect(viewport(), 2, 1)).toEqual({ left: 146, top: 84, width: 100, height: 30 });
  });

  it('滚动后位置**跟着滚动量平移**（这是修过的坑：旧公式在 scale=1 时把滚动消掉了）', () => {
    const scrolled = viewport({ scroll: { x: 100, y: 60 } });
    // left = 46 - 100 = -54（该列被滚出左边），top = 24 - 60 = -36（该行被滚出上边）
    expect(computeFlashRect(scrolled, 0, 0)).toEqual({ left: -54, top: -36, width: 100, height: 30 });
    // 第 1 列在滚动后回到可视区：left = 146 - 100 = 46
    expect(computeFlashRect(scrolled, 0, 1)).toEqual({ left: 46, top: -36, width: 100, height: 30 });
    // 往下滚 1 行的高度（30）→ 第 1 行正好贴到内容顶部
    const oneRow = viewport({ scroll: { x: 0, y: 30 } });
    expect(computeFlashRect(oneRow, 1, 0)).toEqual({ left: 46, top: 24, width: 100, height: 30 });
  });

  it('缩放（表格缩放/浏览器缩放）后尺寸与位置一起放大', () => {
    const zoomed = computeFlashRect(viewport({ scale: { x: 1.5, y: 1.5 } }), 0, 1);
    // left = 146 * 1.5 = 219；width = 100 * 1.5 = 150
    expect(zoomed).toEqual({ left: 219, top: 36, width: 150, height: 45 });
    const zoomedScrolled = computeFlashRect(viewport({ scale: { x: 2, y: 2 }, scroll: { x: 46, y: 24 } }), 0, 1);
    // left = (146 - 46) * 2 = 200；top = (24 - 24) * 2 = 0
    expect(zoomedScrolled).toEqual({ left: 200, top: 0, width: 200, height: 60 });
  });

  it('拿不到单元格 / 非法几何 / 零尺寸 → 返回 null（宁可不画，也不画错位置）', () => {
    expect(computeFlashRect({ ...viewport(), cellRect: () => null }, 0, 0)).toBeNull();
    expect(computeFlashRect(viewport({ scale: { x: 0, y: 1 } }), 0, 0)).toBeNull();
    expect(computeFlashRect(viewport({ scroll: { x: Number.NaN, y: 0 } }), 0, 0)).toBeNull();
    expect(
      computeFlashRect({ ...viewport(), cellRect: () => ({ startX: 10, startY: 10, endX: 10, endY: 10 }) }, 0, 0),
    ).toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* 浮层生命周期                                                                */
/* -------------------------------------------------------------------------- */

interface FakeNode {
  className: string;
  style: Record<string, string>;
  children: FakeNode[];
  isConnected: boolean;
  parent?: FakeNode;
  attributes: Record<string, string>;
  setAttribute: (name: string, value: string) => void;
  appendChild: (child: FakeNode) => FakeNode;
  remove: () => void;
}

function fakeDoc(): { doc: Document; created: FakeNode[]; body: FakeNode } {
  const created: FakeNode[] = [];
  const makeNode = (): FakeNode => {
    const node: FakeNode = {
      className: '',
      style: {} as Record<string, string>,
      children: [],
      isConnected: true,
      attributes: {},
      setAttribute(name, value) {
        node.attributes[name] = value;
      },
      appendChild(child) {
        child.parent = node;
        node.children.push(child);
        return child;
      },
      // 真 DOM 的 remove() 会把自己从父节点摘掉；假节点也必须一致，
      // 否则"到期后页面上不留方框"这条断言就是假的
      remove() {
        node.isConnected = false;
        const parent = node.parent;
        if (parent) parent.children = parent.children.filter((child) => child !== node);
      },
    };
    created.push(node);
    return node;
  };
  const body = makeNode();
  const doc = {
    body,
    createElement: () => makeNode(),
  } as unknown as Document;
  return { doc, created, body };
}

/** 手动驱动的 rAF：把回调排进队列，测试里显式 flush */
function manualRaf() {
  const queue = new Map<number, () => void>();
  let seq = 0;
  return {
    raf: (callback: () => void) => {
      seq += 1;
      queue.set(seq, callback);
      return seq;
    },
    cancelRaf: (id: number) => {
      queue.delete(id);
    },
    flush: () => {
      const pending = [...queue.entries()];
      queue.clear();
      pending.forEach(([, callback]) => callback());
    },
    get pending(): number {
      return queue.size;
    },
  };
}

describe('createSwapFlashOverlay', () => {
  it('show 后建出浮层与方框，并把矩形写进 transform/宽高', () => {
    const { doc, body } = fakeDoc();
    const raf = manualRaf();
    const overlay = createSwapFlashOverlay({
      measure: (target) => ({ left: target.col * 100, top: target.row * 30, width: 100, height: 30 }),
      doc,
      raf: raf.raf,
      cancelRaf: raf.cancelRaf,
    });

    overlay.show([
      { a1: 'B2', row: 1, col: 1 },
      { a1: 'C3', row: 2, col: 2 },
    ]);

    const layer = body.children[0];
    expect(layer.className).toBe(FLASH_LAYER_CLASS);
    expect(layer.children).toHaveLength(2);
    expect(layer.children[0].className).toBe(FLASH_BOX_CLASS);
    expect(layer.children[0].attributes['data-a1']).toBe('B2');
    expect(layer.children[0].style.transform).toBe('translate3d(100px, 30px, 0)');
    expect(layer.children[1].style.transform).toBe('translate3d(200px, 60px, 0)');
    expect(overlay.active).toBe(true);
    overlay.dispose();
  });

  it('到期自动消失（到时间后没有任何方框留在页面上）', () => {
    const { doc, body } = fakeDoc();
    const raf = manualRaf();
    let clock = 0;
    const overlay = createSwapFlashOverlay({
      measure: () => ({ left: 0, top: 0, width: 10, height: 10 }),
      doc,
      raf: raf.raf,
      cancelRaf: raf.cancelRaf,
      now: () => clock,
      durationMs: 400,
    });

    overlay.show([{ a1: 'A1', row: 0, col: 0 }]);
    expect(overlay.active).toBe(true);

    clock = 200;
    raf.flush();
    expect(overlay.active, '没到期时还在').toBe(true);

    clock = 500;
    raf.flush();
    expect(overlay.active, '到期后自动收掉').toBe(false);
    expect(body.children[0].children).toHaveLength(0);
    expect(body.children[0].style.display).toBe('none');
    overlay.dispose();
  });

  it('目标为空 / 量不出矩形时的表现：不建框、不报错', () => {
    const { doc, body } = fakeDoc();
    const raf = manualRaf();
    const overlay = createSwapFlashOverlay({ measure: () => null, doc, raf: raf.raf, cancelRaf: raf.cancelRaf });

    overlay.show([]);
    expect(overlay.active).toBe(false);
    expect(body.children).toHaveLength(0);

    overlay.show([{ a1: 'A1', row: 0, col: 0 }]);
    expect(overlay.active).toBe(true);
    expect(body.children[0].children[0].style.display).toBe('none');

    overlay.hide();
    expect(overlay.active).toBe(false);
    overlay.dispose();
  });

  it('默认时长是 1600ms（用户要"逐渐透明"，太快看不见、太慢挡视线）', () => {
    expect(DEFAULT_FLASH_DURATION_MS).toBe(1600);
  });
});
