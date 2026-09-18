/**
 * 右键菜单（外壳组件，**纯展示 + 受控回调**）。
 *
 * 职责边界（刻意划清，便于上层接线）：
 * - **只管画和导航**：定位、焦点、键盘、鼠标外的关闭时机；
 * - **不碰业务**：不知道什么是单元格、不知道选区、不监听 canvas、不调用 Univer。
 *   选中某项只回调 `onSelect(id)`，**动作执行与关闭菜单都由父层负责**
 *   （父层通常写：`onSelect={(id) => { runAction(id); setMenu(null); }}`）；
 * - 不持有"菜单开没开"的状态：`open` 由父层控制，本组件在 `open === false` 时返回 `null`。
 *
 * 定位：`position: fixed` + `left/top`，数值来自纯函数 `clampMenuPosition`
 * （贴边翻转、8px 边距）。因为要先知道菜单自身尺寸才能夹取，挂载后用 `useLayoutEffect`
 * 量一次 `getBoundingClientRect()`（量之前 `visibility: hidden`，用户看不到跳动），
 * 若位置没变则跳过 setState，避免"渲染 → 测量 → setState"自激循环。
 *
 * 键盘（WAI-ARIA menu 模式）：
 *   ↑/↓ 在可选项之间移动（跳过 disabled，到边界循环）、Home/End 首/末可选项、
 *   Enter/Space 触发当前项、Esc 关闭、Tab 关闭（不拦截默认行为，焦点正常移出）。
 * 关闭时机：菜单外 mousedown（捕获阶段，即使画布 stopPropagation 也能收到）、
 *   窗口失焦、滚动、resize（滚动 / resize 均以 passive 注册）。
 *
 * 样式见 `context-menu.css`：`.cm-*` 类 + shell.css 设计令牌，深色自动跟随系统。
 */
import { Fragment, useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';

import { PlusIcon } from './icons';
import {
  MENU_VIEWPORT_MARGIN,
  clampMenuPosition,
  firstEnabledIndex,
  nextEnabledIndex,
  type MenuItemSpec,
  type Point,
  type Size,
  type Viewport,
} from './context-menu-model';
import './context-menu.css';

/** 转发模型层类型，父层可以只从组件文件 import。 */
export type { MenuItemSpec, Point, Size, Viewport } from './context-menu-model';

export interface ContextMenuProps {
  open: boolean;
  /** 鼠标位置（clientX/clientY） */
  x: number;
  y: number;
  items: MenuItemSpec[];
  /** 选择某项（父层负责执行动作并关闭） */
  onSelect: (id: string) => void;
  onClose: () => void;
}

function cls(...parts: Array<string | false | undefined>): string {
  return parts.filter((part): part is string => typeof part === 'string' && part !== '').join(' ');
}

/**
 * 菜单内容的"指纹"：只有它变化时才重新测量 / 重新夹取位置。
 * 父层几乎每次渲染都会传一个新的 `items` 数组字面量，用引用做依赖会每次都重算，
 * 因此这里按字段拼一个稳定字符串（顺序敏感，字段无关）。
 */
function menuItemsKey(items: MenuItemSpec[]): string {
  return items
    .map((item) =>
      [
        item.id,
        item.label,
        item.shortcut ?? '',
        item.disabled === true ? '1' : '0',
        item.danger === true ? '1' : '0',
        item.separatorBefore === true ? '1' : '0',
      ].join('\u0001'),
    )
    .join('\u0002');
}

/** 当前视口尺寸；拿不到（非浏览器环境）时用菜单自身尺寸兜底，保证不产生 NaN。 */
function readViewport(fallback: Size): Viewport {
  if (typeof window === 'undefined') return { width: fallback.width, height: fallback.height };
  return {
    width: window.innerWidth > 0 ? window.innerWidth : fallback.width,
    height: window.innerHeight > 0 ? window.innerHeight : fallback.height,
  };
}

export function ContextMenu({ open, x, y, items, onSelect, onClose }: ContextMenuProps): JSX.Element | null {
  const menuRef = useRef<HTMLDivElement | null>(null);
  const itemRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const [position, setPosition] = useState<Point | null>(null);
  const [activeIndex, setActiveIndex] = useState(-1);

  /** activeIndex 的镜像：键盘处理里读它，避免把 state 塞进回调依赖造成频繁重建。 */
  const activeIndexRef = useRef(-1);
  const setActive = useCallback((index: number) => {
    activeIndexRef.current = index;
    setActiveIndex(index);
  }, []);

  /** 已按哪个"指纹"完成过定位（用于跳过重复测量）。 */
  const layoutKeyRef = useRef<string | null>(null);
  /** 上一帧是否已打开：用来识别"刚打开 → 自动聚焦首项"这个边沿。 */
  const wasOpenRef = useRef(false);
  /** onClose 的最新引用：让事件订阅只依赖 `open`，父层重渲染不会反复解绑/绑定。 */
  const onCloseRef = useRef(onClose);
  useEffect(() => {
    onCloseRef.current = onClose;
  }, [onClose]);

  const itemsKey = menuItemsKey(items);
  const layoutKey = `${x}|${y}|${itemsKey}`;

  const focusItem = useCallback(
    (index: number) => {
      setActive(index);
      // preventScroll：键盘导航不该让页面/画布跟着滚（滚动会触发关闭监听）
      if (index >= 0) itemRefs.current[index]?.focus({ preventScroll: true });
    },
    [setActive],
  );

  /* ------------------------------------------------------------ 定位 + 自动聚焦 */
  useLayoutEffect(() => {
    if (!open) {
      wasOpenRef.current = false;
      layoutKeyRef.current = null;
      setPosition(null);
      setActive(-1);
      return;
    }

    const node = menuRef.current;
    if (node === null) return;

    if (layoutKeyRef.current !== layoutKey) {
      layoutKeyRef.current = layoutKey;
      // 这一帧还没挂上 .is-positioned（入场动画尚未开始），量到的是**未缩放**的布局盒；
      // 否则动画的 transform: scale() 会让这里量到"缩小版"尺寸，夹取结果偏出 2~3px。
      const rect = node.getBoundingClientRect();
      const size: Size = { width: rect.width, height: rect.height };
      setPosition(clampMenuPosition({ x, y }, size, readViewport(size), MENU_VIEWPORT_MARGIN));
    }

    // 打开的第一帧、或原当前项被移除/禁用时，把焦点交给第一个可选项
    const current = activeIndexRef.current;
    const currentItem = current >= 0 ? items[current] : undefined;
    const needsFocus = !wasOpenRef.current || currentItem === undefined || currentItem.disabled === true;
    wasOpenRef.current = true;

    if (needsFocus) {
      const index = firstEnabledIndex(items);
      setActive(index);
      // preventScroll：聚焦绝不触发滚动 —— 滚动会被本组件的 scroll 监听当成"该关闭了"
      if (index >= 0) itemRefs.current[index]?.focus({ preventScroll: true });
      else node.focus({ preventScroll: true }); // 全是禁用项：焦点落在菜单容器上，Esc 仍然可用
    }
  }, [open, layoutKey, x, y, items, setActive]);

  /* ------------------------------------------------------------ 关闭时机 */
  useEffect(() => {
    if (!open) return undefined;

    // 菜单外 mousedown：捕获阶段监听，画布 / 其他组件 stopPropagation 也拦不住
    const handlePointerDown = (event: MouseEvent): void => {
      const node = menuRef.current;
      const target = event.target;
      if (node !== null && target instanceof Node && node.contains(target)) return;
      onCloseRef.current();
    };
    // 滚动（含内部滚动容器）/ resize：菜单没有跟随锚点的语义，直接关闭
    const handleReflow = (): void => onCloseRef.current();
    // 窗口失焦（切到别的应用 / 打开原生菜单）：不留悬空浮层
    const handleBlur = (): void => onCloseRef.current();

    document.addEventListener('mousedown', handlePointerDown, true);
    document.addEventListener('scroll', handleReflow, { capture: true, passive: true });
    window.addEventListener('resize', handleReflow, { passive: true });
    window.addEventListener('blur', handleBlur);

    return () => {
      document.removeEventListener('mousedown', handlePointerDown, true);
      document.removeEventListener('scroll', handleReflow, true);
      window.removeEventListener('resize', handleReflow);
      window.removeEventListener('blur', handleBlur);
    };
  }, [open]);

  /* ------------------------------------------------------------ 键盘 */
  const handleKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      const { key } = event;

      if (key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        onClose();
        return;
      }

      if (key === 'Tab') {
        // 焦点即将离开菜单：关闭但不拦截默认行为，Tab 照常移动
        onClose();
        return;
      }

      if (key === 'ArrowDown' || key === 'ArrowUp') {
        event.preventDefault();
        focusItem(nextEnabledIndex(items, activeIndexRef.current, key === 'ArrowDown' ? 1 : -1));
        return;
      }

      if (key === 'Home' || key === 'End') {
        event.preventDefault();
        // End = "从第 0 项往前走一步" → 循环到末个可选项；无需为此多导出一个函数
        focusItem(key === 'Home' ? firstEnabledIndex(items) : nextEnabledIndex(items, 0, -1));
        return;
      }

      if (key === 'Enter' || key === ' ' || key === 'Spacebar') {
        const index = activeIndexRef.current;
        const item = index >= 0 ? items[index] : undefined;
        if (item === undefined || item.disabled === true) return;
        // 阻止 <button> 的原生激活，避免同一次按键触发两遍
        event.preventDefault();
        onSelect(item.id);
      }
    },
    [focusItem, items, onClose, onSelect],
  );

  if (!open) return null;

  return (
    <div
      ref={menuRef}
      className={cls('cm-menu', position !== null && 'is-positioned')}
      data-testid="context-menu"
      role="menu"
      aria-orientation="vertical"
      tabIndex={-1}
      // 尺寸量完之前先按点击点摆好并置为全透明：等定位算完（同一帧内、paint 之前）再显示，
      // 用户不会看到"先出现在点击点、再跳到夹取位置"的抖动；.is-positioned 也是在这一帧
      // 才挂上，入场动画因此不会干扰上面的测量。
      // 注意：这里**不能**用 visibility: hidden —— visibility:hidden 的元素不可聚焦，
      // 下面那段"自动聚焦第一个可选项"会静默失败（真实浏览器实测过）。
      style={{
        left: position?.x ?? x,
        top: position?.y ?? y,
        opacity: position === null ? 0 : undefined,
      }}
      onKeyDown={handleKeyDown}
      onContextMenu={(event) => event.preventDefault()}
    >
      {items.map((item, index) => {
        const Icon = resolveItemIcon(item.id);
        const disabled = item.disabled === true;
        const active = index === activeIndex;
        return (
          <Fragment key={item.id}>
            {item.separatorBefore === true && index > 0 ? (
              <div className="cm-sep" role="separator" aria-orientation="horizontal" />
            ) : null}
            <button
              type="button"
              role="menuitem"
              data-testid={`context-menu-${item.id}`}
              className={cls(
                'cm-item',
                active && 'is-active',
                item.danger === true && 'is-danger',
                disabled && 'is-disabled',
              )}
              // 悬停说明：主要用来解释"这一项为什么是灰的"（如多块选区只支持单块操作）
              title={item.title ?? (item.hint ? `${item.label}：${item.hint}` : undefined)}
              // 用 aria-disabled（而非原生 disabled）保留可访问树里的播报，点击由下面的守卫拦掉
              aria-disabled={disabled ? true : undefined}
              tabIndex={active && !disabled ? 0 : -1}
              ref={(node) => {
                itemRefs.current[index] = node;
              }}
              onClick={() => {
                if (disabled) return;
                onSelect(item.id);
              }}
              onMouseEnter={() => {
                if (!disabled) setActive(index);
              }}
              onFocus={() => setActive(index)}
            >
              <span className="cm-icon" aria-hidden="true">
                <Icon />
              </span>
              <span className="cm-label">{item.label}</span>
              {item.hint !== undefined && item.hint !== '' ? (
                <span className="cm-shortcut" data-testid={`context-menu-${item.id}-hint`}>{item.hint}</span>
              ) : item.shortcut !== undefined && item.shortcut !== '' ? (
                <span className="cm-shortcut">{item.shortcut}</span>
              ) : null}
            </button>
          </Fragment>
        );
      })}
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/* 图标：菜单项左侧图标位                                                       */
/* -------------------------------------------------------------------------- */

/**
 * 图标外壳（与 `icons.tsx` 同一套约定：16×16、currentColor、线宽 1.6、纯装饰）。
 * 菜单专用图标就地定义——`icons.tsx` 属于"不许改"的文件，因此只**读**它的通用图标
 * （如 `PlusIcon`），新增图形写在这里。
 */
function Glyph({ children }: { children: ReactNode }): JSX.Element {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width={16}
      height={16}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.6}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  );
}

/** 复制：两张叠起来的纸 */
function CopyIcon(): JSX.Element {
  return (
    <Glyph>
      <rect x="5.4" y="5.4" width="8.2" height="8.2" rx="1.4" />
      <path d="M10.6 5.4V3.8a1.4 1.4 0 0 0-1.4-1.4H3.8a1.4 1.4 0 0 0-1.4 1.4v5.4a1.4 1.4 0 0 0 1.4 1.4h1.6" />
    </Glyph>
  );
}

/** 剪切：剪刀 */
function CutIcon(): JSX.Element {
  return (
    <Glyph>
      <circle cx="4.3" cy="4.3" r="1.9" />
      <circle cx="4.3" cy="11.7" r="1.9" />
      <path d="M6 5.5 13 12.4" />
      <path d="M6 10.5 13 3.6" />
    </Glyph>
  );
}

/** 粘贴：带夹子的写字板 */
function PasteIcon(): JSX.Element {
  return (
    <Glyph>
      <rect x="3.2" y="3.2" width="9.6" height="10.6" rx="1.6" />
      <path d="M6 3.2v-.4a.6.6 0 0 1 .6-.6h2.8a.6.6 0 0 1 .6.6v.4" />
      <path d="M6.2 7.4h3.6" />
      <path d="M6.2 10h3.6" />
    </Glyph>
  );
}

/** 删除：垃圾桶 */
function TrashIcon(): JSX.Element {
  return (
    <Glyph>
      <path d="M2.9 4.4h10.2" />
      <path d="M6.4 4.4V3.2a.9.9 0 0 1 .9-.9h1.4a.9.9 0 0 1 .9.9v1.2" />
      <path d="M4.3 4.4l.6 7.9a1.3 1.3 0 0 0 1.3 1.2h3.6a1.3 1.3 0 0 0 1.3-1.2l.6-7.9" />
      <path d="M6.8 7v3.9" />
      <path d="M9.2 7v3.9" />
    </Glyph>
  );
}

/** 清空：橡皮擦 */
function EraserIcon(): JSX.Element {
  return (
    <Glyph>
      <path d="M9.1 13.1H13" />
      <path d="M6.3 13.1 2.9 9.7a1.4 1.4 0 0 1 0-2l4.8-4.8a1.4 1.4 0 0 1 2 0l2.4 2.4a1.4 1.4 0 0 1 0 2l-4.9 4.9H6.3z" />
      <path d="M6.1 4.5 11 9.4" />
    </Glyph>
  );
}

/** 行：横向切分的表格 */
function RowIcon(): JSX.Element {
  return (
    <Glyph>
      <rect x="2.4" y="3.4" width="11.2" height="9.2" rx="1.4" />
      <path d="M2.4 8h11.2" />
    </Glyph>
  );
}

/** 列：纵向切分的表格 */
function ColumnIcon(): JSX.Element {
  return (
    <Glyph>
      <rect x="2.4" y="3.4" width="11.2" height="9.2" rx="1.4" />
      <path d="M8 3.4v9.2" />
    </Glyph>
  );
}

/** 排序：下箭头 + 由长到短的短线 */
function SortIcon(): JSX.Element {
  return (
    <Glyph>
      <path d="M4.2 3.2v9.6" />
      <path d="M2.2 10.8 4.2 12.8 6.2 10.8" />
      <path d="M8.8 4.6h4.8" />
      <path d="M8.8 8h3.4" />
      <path d="M8.8 11.4h2" />
    </Glyph>
  );
}

/** 合并：两个箭头汇入下方一块 */
function MergeIcon(): JSX.Element {
  return (
    <Glyph>
      <path d="M8 2.4v4" />
      <path d="M5.4 4.4 8 7l2.6-2.6" />
      <rect x="3" y="9" width="10" height="4.4" rx="1.2" />
    </Glyph>
  );
}

/** 兜底图标：未识别的 id 用一个中性圆点，保持左侧图标位对齐 */
function DotIcon(): JSX.Element {
  return (
    <Glyph>
      <circle cx="8" cy="8" r="2.1" fill="currentColor" stroke="none" />
    </Glyph>
  );
}

/**
 * 菜单图标渲染器：本地图标（无参）与 `icons.tsx` 的通用图标（可选 className/size）
 * 都能满足这个签名，因此可以放进同一张映射表。
 */
type MenuIconRenderer = (props: { className?: string; size?: number }) => JSX.Element;

/**
 * 按 id 关键字挑图标（大小写不敏感）。菜单项 id 由父层定义，这里只做"尽力而为"的映射：
 * 命中不了就用中性圆点，绝不因为没有图标而破坏左侧对齐。
 * 注意判定顺序：`insert-row-above` 这类 id 先命中 `row`，符合"行/列"更具体的直觉。
 */
function resolveItemIcon(id: string): MenuIconRenderer {
  const key = id.toLowerCase();
  if (key.includes('cut')) return CutIcon;
  if (key.includes('copy') || key.includes('duplicate')) return CopyIcon;
  if (key.includes('paste')) return PasteIcon;
  if (key.includes('delete') || key.includes('remove') || key.includes('trash')) return TrashIcon;
  if (key.includes('clear') || key.includes('erase') || key.includes('reset')) return EraserIcon;
  if (key.includes('row')) return RowIcon;
  if (key.includes('col')) return ColumnIcon;
  if (key.includes('sort') || key.includes('order')) return SortIcon;
  if (key.includes('merge')) return MergeIcon;
  if (key.includes('insert') || key.includes('add') || key.includes('new')) return PlusIcon;
  return DotIcon;
}
