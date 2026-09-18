/**
 * 「加入工作区」面板（**一个输入框 + 实时预览 + 一个保留/剪切选项**）。
 *
 * 用户要求（原话）："工作区的加号按钮打开的面板请进行重构 … 要求有'加入工作区后保留表格内容'的选项，
 * 务必保证没有无用的元素比如'跳过空单元格'，而且该面板添加的数据不会被历史记录所记录而导致无法撤销"。
 * 因此这一版做了三件事：
 *  ① **删掉"跳过空单元格"那个禁用复选框**——它是不可交互的说明文字，占地方还容易被误读成"能关"；
 *     跳过空内容本来就是底层唯一入口 `extractCellItems` 的固定行为，改成在**预览**里如实报数
 *     （"跳过 156 个空内容"），用户看到的是结果而不是一句空洞的保证；
 *  ② **新增"加入工作区后保留表格内容"选项**：勾选＝复制（表格不动），取消＝剪切（加入后清空这些格子的
 *     内容、保留格式）。它和顶部工具栏"拖到工作区保留内容"是**同一个设置**（`keepSourceOnDrop`），
 *     面板里改一下就会被记住，下次打开还是它 —— 一处语义，两处入口；
 *  ③ **实时预览**：输入框下面直接显示"将加入多少个单元格、跳过多少个空内容、前几格是什么"，
 *     确认键上也带数字（「加入 12 格」）。用户不必先按下去才知道会发生什么。
 *
 * 撤销口径（用户重点强调）：本面板**不直接改任何状态**——它只把
 * `{ kind, value, keepSource }` 抛给上层；上层用工作区的**唯一提交入口**（`commitWorkspace`，
 * 带 before/after 快照的历史条目）和表格侧的统一通道（`beginSheetAction` + `pushHistory`）落地。
 * 所以这里的每一次确认都正好是**一条可撤销的历史**：撤销先回滚表格（剪切的情形），再回滚工作区。
 *
 * 职责边界（纯展示 + 受控回调）：
 *  - 不取数据、不碰 Univer、不读全局状态、不发任何请求；
 *  - 识别与校验全部来自 `import-parse.ts`（`detectImportKind` + 三个 parse 函数）与
 *    `selection-ranges.ts`（多块写法），组件只负责呈现与交互；`resolveImportDraft` 是唯一把
 *    "识别结果 + 提交值 + 文案"揉在一起的纯函数（导出是为了单测能直接覆盖，UI 只调它一个）；
 *    预览统计由上层通过 `onPreview` 注入（它需要读表），本组件只渲染；
 *  - 只有 `useState` 管草稿 + 必要的焦点管理副作用（自动聚焦、ESC、焦点陷阱、关闭后还原焦点）。
 *
 * testid 契约（e2e 依赖，勿改名）：import-dialog / import-close / import-input / import-error /
 * import-detected / import-use-selection / import-whole-sheet / import-preview /
 * import-preview-summary / import-preview-samples / import-keep-source / import-cancel / import-confirm。
 * （`import-skip-empty` 已随那个无用复选框一起删除。）
 */
import { useCallback, useEffect, useId, useRef, useState } from 'react';
import type { ChangeEvent, FormEvent, MouseEvent as ReactMouseEvent } from 'react';

import { CloseIcon } from '../shell/icons';
import { detectImportKind, parseColumnInput, parseRangeInput, parseRowInput } from './import-parse';
import type { ImportTargetKind, ParseResult } from './import-parse';
// 多块区域（Ctrl+点选）的写法与规范化：与表格里的多选共用同一套纯函数
import { parseRangeList, rectsToA1, totalCells } from '../interaction/selection-ranges';
import { PREVIEW_EMPTY_CAP, type WorkspaceImportPreview } from './snapshot';
import './import-dialog.css';

export type { ImportTargetKind };

/** 一次"加入工作区"的用户意图：范围 + 是否保留表格内容 */
export interface ImportSubmitPayload {
  kind: ImportTargetKind;
  value: string;
  /** true = 复制（表格内容保留）；false = 剪切（加入后清空源内容，可撤销） */
  keepSource: boolean;
}

export interface ImportDialogProps {
  open: boolean;
  /** 关闭（点遮罩/ESC/取消） */
  onClose: () => void;
  /** 当前选区对应的 A1 记号，用于"当前选区"按钮；无选区传 null */
  currentSelectionA1: string | null;
  /** 当前工作表名，用于文案 */
  sheetName: string;
  /** 当前表「已用区域」的 A1 记号（用于"整个已用区域"一键填范围）；空表传 null */
  usedRangeA1: string | null;
  /** 加入工作区后是否保留表格内容（与顶部工具栏开关共用同一个设置） */
  keepSource: boolean;
  onKeepSourceChange: (keep: boolean) => void;
  /** 预览钩子：由上层读表算出"这次会加入哪些格子"；不传则不显示预览块 */
  onPreview?: (target: { kind: ImportTargetKind; value: string }) => WorkspaceImportPreview | null;
  /** 用户确认后的结果（上层负责落地成一条可撤销的历史） */
  onSubmit: (result: ImportSubmitPayload) => void;
  /** 提交中（禁用按钮） */
  busy?: boolean;
}

/** 输入框标题 / 占位符 / 提示：三种写法都写得下，用户不必先选类型 */
const FIELD_LABEL = '数据范围';
const PLACEHOLDER = '例如 3、B:D 或 A1:B18';
const HINT = '同一个输入框自动识别：行号 3 / 3:5、列标 B / B:D、区域 A1:B18（$A$1:$B$18 也认）。多块区域用空格分开，如 A1:B2 D4:E5。';

/** 空输入时识别行显示的占位文案（不飘红：刚打开就满屏红字太吓人） */
const DETECTED_EMPTY = '识别为：—';
/** 非空但识别不出时识别行显示的内容（具体原因由下方的 import-error 给出） */
const DETECTED_INVALID = '无法识别';
const GENERIC_ERROR = '无法识别：请输入行号（3 或 3:5）、列标（B 或 B:D）或区域（A1:B18）';

/** 预览里最多列几个样例格子（其余只报数，避免面板被撑爆） */
const PREVIEW_SAMPLES = 6;
/** 预览防抖（毫秒）：边打字边读表太浪费，停一下再算 */
const PREVIEW_DEBOUNCE_MS = 120;

/** 可聚焦元素选择器（焦点陷阱用）：排除 disabled / tabindex="-1" */
const FOCUSABLE_SELECTOR =
  'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled),' +
  ' a[href], [tabindex]:not([tabindex="-1"])';

function cls(...parts: (string | undefined | false)[]): string {
  return parts.filter((p): p is string => typeof p === 'string' && p !== '').join(' ');
}

/** 按已识别出的类型调用对应的解析函数（识别与校验用的是同一批函数，二者永不打架） */
function parseByKind(kind: ImportTargetKind, raw: string): ParseResult {
  switch (kind) {
    case 'row':
      return parseRowInput(raw);
    case 'column':
      return parseColumnInput(raw);
    case 'range':
      return parseRangeInput(raw);
  }
}

/** 提交给上层的 value：区域统一成 `A1:A1` 形式（单个单元格补成 B2:B2），行 / 列原样（解析时已规范化） */
function toSubmitValue(kind: ImportTargetKind, value: string): string {
  if (kind === 'range' && !value.includes(':')) return `${value}:${value}`;
  return value;
}

/**
 * 识别结果那行的文案（不含"识别为："前缀）：
 * `('row', '3')`→`'第 3 行'`、`('row', '3:5')`→`'第 3:5 行'`、
 * `('column', 'B:D')`→`'B:D 列'`、`('range', 'A1:B18')`→`'区域 A1:B18'`。
 * 导出仅为单测能直接覆盖这行 UI 文案。
 */
export function describeImportTarget(kind: ImportTargetKind, value: string): string {
  switch (kind) {
    case 'row':
      return `第 ${value} 行`;
    case 'column':
      return `${value} 列`;
    case 'range':
      return `区域 ${value}`;
  }
}

/**
 * 识别失败时，尽量按用户的"本来意图"给一句具体原因（纯文案，不参与解析）：
 * 空 → 内部空格 → 非半角字符 → 按形状挑一个解析函数取其错误文案。
 * 走到这里时三个解析函数必然全部失败，所以一定拿得到具体的 `error`。
 */
function explainImportError(raw: string): string {
  const text = typeof raw === 'string' ? raw.trim() : '';
  if (text === '') return '请输入行号、列标或区域';
  if (/\s/.test(text)) return '输入中间不能有空格（列标写作 B:D，区域写作 A1:B18）';
  if (/[^\x21-\x7e]/.test(text)) return '只支持半角字符（例如 A1:B18，全角请切换输入法）';

  const hasLetter = /[A-Za-z]/.test(text);
  const hasDigit = /[0-9]/.test(text);
  if (!hasLetter && !hasDigit) return GENERIC_ERROR; // ':'、'$'、'$$' 这类"什么都没有"的输入

  // 按输入形状挑一个"最像用户本来意图"的解析函数，借它的错误文案（例如"列标超出范围（最大 XFD）"）
  const parsed = hasLetter && hasDigit
    ? parseRangeInput(text)
    : hasLetter
      ? parseColumnInput(text)
      : parseRowInput(text);
  return parsed.ok ? GENERIC_ERROR : parsed.error;
}

/** 识别成功的结果 */
export interface ResolvedImport {
  kind: ImportTargetKind;
  /** 解析函数规范化后的输入（行 `'3'` / `'3:5'`、列 `'B'` / `'B:D'`、区域 `'A1'` / `'A1:B18'`） */
  value: string;
  /** 抛给上层 `onSubmit` 的 value（区域补全成 `A1:A1`） */
  submitValue: string;
  /** 识别结果那行的文案（不含"识别为："前缀） */
  label: string;
}

/** 草稿的识别 + 校验结果：判别联合（`ok` 收窄） */
export type ResolvedDraft = ({ ok: true } & ResolvedImport) | { ok: false; error: string };

/**
 * 把输入框里的原始草稿变成"加入目标"：**纯函数**，识别 + 校验 + 归一化 + 文案一步到位。
 * 识别不出（含空输入）时给出面向用户的原因，调用方据此禁用确认键并显示错误。
 */
export function resolveImportDraft(raw: string): ResolvedDraft {
  /**
   * **多块区域**（用户要求"不连续多区域 Ctrl+点选"）：`A1:B2 D4:E5` 也认。
   *
   * 判定放在最前面：多块写法用空格/逗号/顿号分隔，与单块的"中间不能有空格"规则天然互斥，
   * 所以不会和既有识别打架。提交值原样带上全部块（空格分隔），由上层逐块取内容。
   */
  const multi = parseRangeList(raw);
  if (multi.rects.length > 1 && multi.invalid.length === 0) {
    const a1List = rectsToA1(multi.rects);
    const cells = totalCells(multi.rects);
    return {
      ok: true,
      kind: 'range',
      value: a1List.join(' '),
      submitValue: a1List.join(' '),
      label: `区域 ${a1List.join(' + ')}（${a1List.length} 块 / ${cells} 格）`,
    };
  }
  if (multi.rects.length === 1 && multi.invalid.length > 0) {
    return { ok: false, error: `有无法识别的片段：${multi.invalid.join('、')}（区域写作 A1:B18，多块用空格分隔）` };
  }

  const kind = detectImportKind(raw);
  if (kind === null) return { ok: false, error: explainImportError(raw) };

  const parsed = parseByKind(kind, raw);
  // 兜底：识别得出来就一定解析得动（两者用的是同一批函数），这里只是给 TS 收窄
  if (!parsed.ok) return { ok: false, error: parsed.error };

  return {
    ok: true,
    kind,
    value: parsed.value,
    submitValue: toSubmitValue(kind, parsed.value),
    label: describeImportTarget(kind, parsed.value),
  };
}

/**
 * 预览摘要那一行的文案（导出仅为单测能直接覆盖这行 UI 文案）：
 * `将加入 12 个单元格 · 跳过 156 个空内容`；没有预览（未识别 / 未注入钩子）时返回 null。
 *
 * 空格子数到 `PREVIEW_EMPTY_CAP` 就不再展示：用户可以把选区写得极大（`A1:A1048576`），
 * 报一个上亿的数字只会让人困惑（"到底要搬多少"看的是 `total`）。
 */
export function describeImportPreview(preview: WorkspaceImportPreview | null): string | null {
  if (!preview) return null;
  const parts: string[] = [`将加入 ${preview.total} 个单元格`];
  if (preview.empty > 0 && preview.empty < PREVIEW_EMPTY_CAP) parts.push(`跳过 ${preview.empty} 个空内容`);
  if (preview.blocks > 1) parts.push(`${preview.blocks} 块区域`);
  return parts.join(' · ');
}

/** 确认键文案：带上会加入的格数，用户不必先按下去才知道会发生什么（导出仅为单测） */
export function importConfirmLabel(preview: WorkspaceImportPreview | null): string {
  return preview && preview.total > 0 ? `加入 ${preview.total} 格` : '加入工作区';
}

/** 收集容器内当前可聚焦的元素（顺序即 DOM 顺序） */
function focusableIn(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter((el) => el.tabIndex >= 0);
}

export function ImportDialog({
  open,
  onClose,
  currentSelectionA1,
  sheetName,
  usedRangeA1,
  keepSource,
  onKeepSourceChange,
  onPreview,
  onSubmit,
  busy = false,
}: ImportDialogProps): JSX.Element | null {
  const [draft, setDraft] = useState('');
  const [preview, setPreview] = useState<WorkspaceImportPreview | null>(null);

  const boxRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  /** 关闭后要把焦点还给"打开对话框的那个按钮" */
  const restoreRef = useRef<HTMLElement | null>(null);
  /** 遮罩上是否按下了鼠标（避免在输入框里拖选文字、松手落在遮罩上时误关） */
  const pressedOnOverlay = useRef(false);

  const baseId = useId();
  const titleId = `${baseId}-title`;
  const descId = `${baseId}-desc`;
  const inputId = `${baseId}-input`;
  const detectId = `${baseId}-detect`;
  const hintId = `${baseId}-hint`;
  const errorId = `${baseId}-error`;
  const previewId = `${baseId}-preview`;
  const keepId = `${baseId}-keep`;

  /* ------------------------------------------------------------------ 副作用 */

  // 关闭时就把草稿与预览复位：这样"下次打开的第一帧"已经是干净的空输入框，
  // 下面的自动聚焦不会因为"打开瞬间还要清空"而把焦点丢在刚被卸载的输入框上。
  useEffect(() => {
    if (open) return;
    setDraft('');
    setPreview(null);
  }, [open]);

  // 记住打开前的焦点元素，关闭时还回去（键盘用户不会"掉到页面开头"）
  useEffect(() => {
    if (!open) return;
    if (typeof document === 'undefined') return;
    const active = document.activeElement;
    restoreRef.current = active instanceof HTMLElement ? active : null;
    return () => {
      const target = restoreRef.current;
      restoreRef.current = null;
      if (target && target.isConnected) target.focus();
    };
  }, [open]);

  // 打开时焦点自动落到输入框并整段选中（草稿已在关闭时复位，所以重新打开就是从空开始）
  useEffect(() => {
    if (!open) return;
    const input = inputRef.current;
    if (!input) return;
    input.focus();
    input.select();
  }, [open]);

  // ESC 关闭 + Tab 焦点陷阱（capture 阶段监听：即使表格里某个组件吞掉了 keydown，ESC 依然有效）
  useEffect(() => {
    if (!open) return;
    if (typeof document === 'undefined') return;

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        if (busy) return; // 提交中不允许用 ESC 中断
        event.preventDefault();
        onClose();
        return;
      }
      if (event.key !== 'Tab') return;

      const root = boxRef.current;
      if (!root) return;
      const focusables = focusableIn(root);
      if (focusables.length === 0) return;

      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      const active = document.activeElement;
      const inside = active instanceof HTMLElement && root.contains(active);

      if (event.shiftKey) {
        if (!inside || active === first) {
          event.preventDefault();
          last.focus();
        }
        return;
      }
      if (!inside || active === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener('keydown', onKeyDown, true);
    return () => document.removeEventListener('keydown', onKeyDown, true);
  }, [open, busy, onClose]);

  /* ------------------------------------------------------------------ 识别 */

  const trimmed = draft.trim();
  const isEmpty = trimmed === '';
  const resolved = resolveImportDraft(draft);
  // 空输入不飘红（占位符和提示已经说清楚了）；非空且识别不出时立刻给出原因
  const error = !resolved.ok && !isEmpty ? resolved.error : null;
  const detected = resolved.ok
    ? `识别为：${resolved.label}`
    : isEmpty
      ? DETECTED_EMPTY
      : DETECTED_INVALID;

  const selection = typeof currentSelectionA1 === 'string' ? currentSelectionA1.trim() : '';
  const hasSelection = selection !== '';

  /** 预览的依赖键：只在"真正换了目标范围"时重算（对象每次渲染都是新的，不能直接当依赖） */
  const targetKey = resolved.ok ? `${resolved.kind}|${resolved.submitValue}` : '';

  // 预览：防抖后调上层注入的读表钩子；目标不可识别 / 未注入钩子时清空
  useEffect(() => {
    if (!open || !onPreview || targetKey === '') {
      setPreview(null);
      return;
    }
    const [kind, value] = targetKey.split('|') as [ImportTargetKind, string];
    let cancelled = false;
    const timer = window.setTimeout(() => {
      if (cancelled) return;
      try {
        setPreview(onPreview({ kind, value }));
      } catch {
        setPreview(null);
      }
    }, PREVIEW_DEBOUNCE_MS);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [open, onPreview, targetKey]);

  const previewLine = describeImportPreview(preview);

  /* ------------------------------------------------------------------ 交互 */

  const focusInput = (select: boolean): void => {
    const input = inputRef.current;
    if (!input) return;
    input.focus();
    if (select) input.select();
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>): void => {
    setDraft(event.target.value);
  };

  /** 把某个范围填进输入框（不直接提交：先让用户看到预览，再确认） */
  const fillDraft = useCallback((value: string) => {
    setDraft(value);
    // 填入后把焦点送回输入框，但**不整段选中**：用户多半想接着改这个范围记号
    inputRef.current?.focus();
  }, []);

  const handleSubmit = (event: FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    if (busy) return;
    const result = resolveImportDraft(draft);
    if (!result.ok) {
      // 兜底：正常情况下识别不出时确认键已禁用，走不到这里
      focusInput(true);
      return;
    }
    onSubmit({ kind: result.kind, value: result.submitValue, keepSource });
  };

  const handleOverlayMouseDown = (event: ReactMouseEvent<HTMLDivElement>): void => {
    pressedOnOverlay.current = event.target === event.currentTarget;
  };

  const handleOverlayClick = (event: ReactMouseEvent<HTMLDivElement>): void => {
    const pressed = pressedOnOverlay.current;
    pressedOnOverlay.current = false;
    if (!pressed || busy) return;
    if (event.target !== event.currentTarget) return;
    onClose();
  };

  if (!open) return null;

  const samples = preview?.samples ?? [];
  const shown = samples.slice(0, PREVIEW_SAMPLES);
  const restCount = preview ? Math.max(0, preview.total - shown.length) : 0;
  const confirmLabel = importConfirmLabel(preview);

  return (
    <div
      className="imp-overlay"
      data-testid="import-dialog"
      onMouseDown={handleOverlayMouseDown}
      onClick={handleOverlayClick}
    >
      <div
        className="imp-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={descId}
        aria-busy={busy}
        ref={boxRef}
      >
        <header className="imp-head">
          <div className="imp-head-text">
            <h2 className="imp-title" id={titleId}>
              加入工作区
            </h2>
            <p className="imp-subtitle" id={descId}>
              把「
              <span className="imp-sheet">{sheetName}</span>
              」里的内容拆成一个个单元格放进工作区，随后可以逐个拖回表格
            </p>
          </div>
          <button
            type="button"
            className="imp-icon-btn"
            data-testid="import-close"
            aria-label="关闭"
            title="关闭（Esc）"
            disabled={busy}
            onClick={onClose}
          >
            <CloseIcon size={14} />
          </button>
        </header>

        <form className="imp-form" onSubmit={handleSubmit} noValidate>
          <div className="imp-body">
            <div className="imp-field">
              <label className="imp-label" htmlFor={inputId}>
                {FIELD_LABEL}
              </label>
              <input
                id={inputId}
                ref={inputRef}
                className={cls('imp-input', error !== null && 'is-invalid')}
                data-testid="import-input"
                data-kind={resolved.ok ? resolved.kind : undefined}
                type="text"
                value={draft}
                placeholder={PLACEHOLDER}
                autoComplete="off"
                spellCheck={false}
                aria-invalid={error !== null}
                aria-describedby={cls(
                  error !== null ? errorId : '',
                  detectId,
                  previewId,
                  hintId,
                )}
                onChange={handleInputChange}
              />
              <p className="imp-hint" id={detectId} data-testid="import-detected" aria-live="polite">
                {detected}
              </p>
              <p className="imp-error" id={errorId} data-testid="import-error" aria-live="polite">
                {error ?? ''}
              </p>

              {/* 两个"帮我填范围"的快捷入口（都只填输入框，不直接提交：先看预览再确认） */}
              <div className="imp-quick">
                <button
                  type="button"
                  className="imp-btn is-ghost imp-quick-btn"
                  data-testid="import-use-selection"
                  disabled={busy || !hasSelection}
                  title={hasSelection ? `把当前选区 ${selection} 填入输入框` : '当前没有选区'}
                  onClick={() => {
                    if (busy || !hasSelection) return;
                    fillDraft(selection);
                  }}
                >
                  当前选区
                  <span className={cls('imp-sel-value', !hasSelection && 'is-empty')}>
                    {hasSelection ? selection : '无选区'}
                  </span>
                </button>
                <button
                  type="button"
                  className="imp-btn is-ghost imp-quick-btn"
                  data-testid="import-whole-sheet"
                  disabled={busy || !usedRangeA1}
                  title={
                    usedRangeA1
                      ? `把整表已用区域 ${usedRangeA1} 填入输入框（空内容仍会被跳过）`
                      : '当前表没有内容'
                  }
                  onClick={() => {
                    if (busy || !usedRangeA1) return;
                    fillDraft(usedRangeA1);
                  }}
                >
                  整个已用区域
                  <span className={cls('imp-sel-value', !usedRangeA1 && 'is-empty')}>
                    {usedRangeA1 ?? '空表'}
                  </span>
                </button>
              </div>
              <p className="imp-hint" id={hintId}>
                {HINT}
              </p>
            </div>

            {/* 实时预览：确认之前就把"会发生什么"如实摆出来（替代原来那个不可点的"跳过空单元格"） */}
            <section
              className={cls('imp-preview', preview && preview.total > 0 ? 'is-ready' : undefined)}
              data-testid="import-preview"
              data-total={preview?.total ?? 0}
              id={previewId}
              aria-live="polite"
            >
              <p className="imp-preview-line" data-testid="import-preview-summary">
                {previewLine ?? '识别出范围后，这里显示会加入哪些单元格'}
              </p>
              {preview && preview.total > 0 ? (
                <div className="imp-preview-samples" data-testid="import-preview-samples">
                  {shown.map((sample) => (
                    <span className="imp-sample" key={`${sample.a1}-${sample.text}`}>
                      <span className="imp-sample-a1">{sample.a1}</span>
                      {sample.text}
                    </span>
                  ))}
                  {restCount > 0 ? <span className="imp-sample-more">…还有 {restCount} 格</span> : null}
                </div>
              ) : null}
              {preview && preview.overLimit ? (
                <p className="imp-preview-warn">
                  超过单次上限 {preview.limit} 格：只会加入前 {preview.limit} 格
                </p>
              ) : null}
            </section>

            {/* 保留 / 剪切：与顶部工具栏"拖到工作区保留内容"是同一个设置 */}
            <div className="imp-keep-row">
              <label className="imp-check" htmlFor={keepId}>
                <input
                  id={keepId}
                  type="checkbox"
                  data-testid="import-keep-source"
                  checked={keepSource}
                  disabled={busy}
                  onChange={(event) => onKeepSourceChange(event.target.checked)}
                />
                <span>加入工作区后保留表格内容</span>
              </label>
              <p className="imp-hint" data-testid="import-keep-hint">
                {keepSource
                  ? '勾选中：表格里的原内容保持不动（等于复制）。'
                  : '已取消勾选：加入后清空这些格子的内容、保留格式（等于剪切）；撤销一次即可恢复。'}
              </p>
            </div>
          </div>

          <footer className="imp-foot">
            <button type="button" className="imp-btn" data-testid="import-cancel" disabled={busy} onClick={onClose}>
              取消
            </button>
            <button
              type="submit"
              className="imp-btn is-primary"
              data-testid="import-confirm"
              disabled={busy || !resolved.ok}
            >
              {busy ? '加入中…' : confirmLabel}
            </button>
          </footer>
        </form>
      </div>
    </div>
  );
}
