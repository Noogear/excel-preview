/**
 * 外壳图标集（内联 SVG，零依赖）。
 *
 * 统一约定：
 * - `viewBox="0 0 16 16"`，默认 `16×16`，颜色取 `currentColor`，线宽 `1.6`，圆头圆角；
 * - `aria-hidden="true"` + `focusable="false"`：图标只是装饰，可访问名称由宿主按钮的
 *   `aria-label` / 可见文本提供（避免读屏重复播报）；
 * - 组件是纯函数、无 hooks、无副作用，签名固定为 `{ className?, size? } => JSX.Element`，
 *   **刻意不使用 `React.FC`**（FC 隐含 children 与可选返回类型，不利于 strict 下的显式契约）。
 */
import type { ReactNode } from 'react';

/** 所有图标共用的 props：只允许调样式与尺寸，不允许塞 children。 */
export interface IconProps {
  className?: string;
  size?: number;
}

interface GlyphProps extends IconProps {
  children: ReactNode;
}

/**
 * 图标外壳：集中承载 viewBox / 描边 / 无障碍属性，具体图形由各图标以 path 传入。
 * 这样新增图标时不会漏掉 strokeWidth 或 aria-hidden。
 */
function Glyph({ className, size, children }: GlyphProps): JSX.Element {
  return (
    <svg
      className={className}
      xmlns="http://www.w3.org/2000/svg"
      width={size ?? 16}
      height={size ?? 16}
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

/** 打开文件：张开的文件夹 */
export function OpenIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M2.4 12.1V4.3a1.2 1.2 0 0 1 1.2-1.2h2.6l1.4 1.7h4.6a1.2 1.2 0 0 1 1.2 1.2v1.1" />
      <path d="M2.4 12.1l1.7-4.6a1.1 1.1 0 0 1 1.04-.72h8.1a1.1 1.1 0 0 1 1.04 1.48l-1.4 3.9a1.1 1.1 0 0 1-1.04.72H3.6a1.2 1.2 0 0 1-1.2-1.2z" />
    </Glyph>
  );
}

/** 导出：向下箭头落入托盘（下载语义） */
export function ExportIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M8 2.6v7.4" />
      <path d="M4.9 6.9 8 10l3.1-3.1" />
      <path d="M2.9 11.2v1.2a1.4 1.4 0 0 0 1.4 1.4h7.4a1.4 1.4 0 0 0 1.4-1.4v-1.2" />
    </Glyph>
  );
}

/** 撤销：左向回勾箭头 */
export function UndoIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M3.2 6.6h6.6a3.6 3.6 0 0 1 0 7.2H5.3" />
      <path d="M5.9 3.9 3.2 6.6l2.7 2.7" />
    </Glyph>
  );
}

/** 重做：右向回勾箭头（撤销的镜像） */
export function RedoIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M12.8 6.6H6.2a3.6 3.6 0 0 0 0 7.2h4.5" />
      <path d="M10.1 3.9l2.7 2.7-2.7 2.7" />
    </Glyph>
  );
}

/** 历史记录：时钟 */
export function HistoryIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <circle cx="8" cy="8" r="5.6" />
      <path d="M8 4.6V8l2.5 1.5" />
    </Glyph>
  );
}

/** 运行日志：终端窗口 */
export function LogIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <rect x="2.2" y="3.4" width="11.6" height="9.2" rx="1.6" />
      <path d="M4.9 6.7l1.7 1.5-1.7 1.5" />
      <path d="M8.6 10.1h2.6" />
    </Glyph>
  );
}

/** 交互模式「选择」：指针 + 选择框（表示"只做选择，不搬运"） */
export function SelectModeIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M3.6 3.2 9.1 8.4 6.5 8.9 8 11.8 6.7 12.5 5.2 9.6 3.6 11Z" />
      <path d="M10.8 9.4h5.2" />
      <path d="M10.8 12.4h5.2" />
      <path d="M10.8 15.4h5.2" />
    </Glyph>
  );
}

/** 交互模式「按住拖拽」：手 */
export function DragModeIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M12 7.3V4a1.33 1.33 0 0 0-1.33-1.33 1.33 1.33 0 0 0-1.33 1.33" />
      <path d="M9.33 6.7V2.67a1.33 1.33 0 0 0-1.33-1.34 1.33 1.33 0 0 0-1.33 1.34v1.33" />
      <path d="M6.67 7V4a1.33 1.33 0 0 0-1.34-1.33A1.33 1.33 0 0 0 4 4v5.33" />
      <path d="M12 5.33a1.33 1.33 0 1 1 2.67 0v4a5.33 5.33 0 0 1-5.34 5.34H8c-1.87 0-3-.58-3.99-1.56l-2.4-2.4a1.33 1.33 0 0 1 1.89-1.88L4.67 10" />
    </Glyph>
  );
}

/** 交互模式「点击交换」：双向箭头 */
export function SwapModeIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M3 5.7h9.2" />
      <path d="M9.9 3.4 12.2 5.7 9.9 8" />
      <path d="M13 10.3H3.8" />
      <path d="M6.1 8 3.8 10.3 6.1 12.6" />
    </Glyph>
  );
}

/** 新增：加号 */
export function PlusIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M8 3.4v9.2" />
      <path d="M3.4 8h9.2" />
    </Glyph>
  );
}

/** 关闭：叉号 */
export function CloseIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <path d="M4.3 4.3l7.4 7.4" />
      <path d="M11.7 4.3l-7.4 7.4" />
    </Glyph>
  );
}

/** 工作区侧栏：右侧带分栏的方框（与本项目侧栏在右侧的布局一致） */
export function WorkspaceIcon({ className, size }: IconProps): JSX.Element {
  return (
    <Glyph className={className} size={size}>
      <rect x="2.2" y="3.4" width="11.6" height="9.2" rx="1.6" />
      <path d="M9.8 3.4v9.2" />
    </Glyph>
  );
}
