/**
 * 像素 ↔ 单元格 换算（纯函数）。
 *
 * 这套换算是"拖动落点、落点高亮、互换提醒框"共同的底座，且是**实测校准**出来的
 * （对着真实夹具：点一个像素 → 读 Univer 认定的格 → 反推公式）。踩过的坑：
 *  1. 上游反解在 scale=1 时把滚动量消掉 → 滚动后落点偏移（指针在 S74，解析成 D13）；
 *  2. `columnWidthAccumulation` 是**累计边界**（含自身、不含表头），所以查找必须是
 *     "第一个 > 坐标 的下标"，写成"最后一个 ≤"就会少一格。
 */
import { describe, expect, it } from 'vitest';

import { cellAtPoint, indexAtOffset, type GridMetrics } from '../../src/interaction/cell-geometry';

/** 真实夹具里量到的数字：行头 46、列头 20；第 18 列右边界 1714、第 73 行下边界 1862 */
const columnOffsets = [
  // 前 6 列与 16..20 列是实测值，中间按等差补齐（只为让下标与真实数组一致）
  131, 234, 337, 496, 585, 674, 755, 836, 917, 998, 1079, 1160, 1241, 1322, 1403, 1484, 1564, 1639, 1714, 1789, 1864,
];
const rowOffsets = [
  ...Array.from({ length: 70 }, (_, index) => (index + 1) * 25),
  1790,
  1814,
  1838,
  1862,
  1886,
  1910,
];

const metrics: GridMetrics = {
  columnOffsets,
  rowOffsets,
  headerWidth: 46,
  headerHeight: 20,
  scroll: { x: 1186, y: 1500 },
  scale: { x: 1, y: 1 },
};

describe('indexAtOffset（累计边界查找）', () => {
  it('返回第一个"边界 > 坐标"的下标（该坐标落在这一段里）', () => {
    expect(indexAtOffset([131, 234, 337], 0)).toBe(0);
    expect(indexAtOffset([131, 234, 337], 130)).toBe(0);
    // 正好压在边界上：属于**下一格**（与上游 searchArray 的语义一致）
    expect(indexAtOffset([131, 234, 337], 131)).toBe(1);
    expect(indexAtOffset([131, 234, 337], 200)).toBe(1);
    expect(indexAtOffset([131, 234, 337], 337)).toBe(2);
  });

  it('超出末尾算最后一格（点空白处仍落在最后一列）', () => {
    expect(indexAtOffset([131, 234, 337], 99999)).toBe(2);
  });

  it('空数组返回 -1（拿不到几何就别猜）', () => {
    expect(indexAtOffset([], 10)).toBe(-1);
  });
});

describe('cellAtPoint（滚动后的命中换算）', () => {
  it('实测校准点：滚动 (1186,1500) 时，像素 (500,380) 应是第 18 列第 73 行（S74）', () => {
    // 真值来自"点这个像素 → Univer 选中的是 S74"
    expect(cellAtPoint(metrics, 500, 380)).toEqual({ row: 73, col: 18 });
  });

  it('同一个像素在未滚动时指向左上角附近的格（说明滚动真的参与了换算）', () => {
    const atOrigin = cellAtPoint({ ...metrics, scroll: { x: 0, y: 0 } }, 500, 380);
    // contentX = 500-46 = 454 → 第 3 列（D）；contentY = 380-20 = 360 → 第 14 行
    expect(atOrigin).toEqual({ row: 14, col: 3 });
  });

  it('滚动量平移一格的高度 → 命中行随之加一', () => {
    const one = cellAtPoint({ ...metrics, scroll: { x: 1186, y: 1500 + 24 } }, 500, 380);
    expect(one?.row).toBe(74);
    expect(one?.col).toBe(18);
  });

  it('横向滚动同理', () => {
    const shifted = cellAtPoint({ ...metrics, scroll: { x: 1186 + 75, y: 1500 } }, 500, 380);
    expect(shifted?.col).toBe(19);
  });

  it('缩放下换算互逆：像素坐标要按缩放折算', () => {
    // contentX = 250/2 + 1186 - 46 = 1265 → 第 13 列；contentY = 190/2 + 1500 - 20 = 1575 → 第 63 行
    const zoomed = cellAtPoint({ ...metrics, scale: { x: 2, y: 2 } }, 250, 190);
    expect(zoomed).toEqual({ row: 63, col: 13 });
  });

  it('点在行头/列头上 → 不是单元格（返回 null）', () => {
    const atOrigin = { ...metrics, scroll: { x: 0, y: 0 } };
    expect(cellAtPoint(atOrigin, 10, 300), '行头（x < 46）').toBeNull();
    expect(cellAtPoint(atOrigin, 300, 5), '列头（y < 20）').toBeNull();
  });

  it('缩放为 0 / 非有限值 → null（不产出非法坐标）', () => {
    expect(cellAtPoint({ ...metrics, scale: { x: 0, y: 1 } }, 100, 100)).toBeNull();
    expect(cellAtPoint({ ...metrics, scroll: { x: Number.NaN, y: 0 } }, 100, 100)).toBeNull();
    expect(cellAtPoint({ ...metrics, columnOffsets: [] }, 100, 100)).toBeNull();
  });
});
