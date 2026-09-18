/**
 * 历史记录面板（底部抽屉）。
 *
 * 职责边界：**纯展示 + 抛事件**。
 * - 不 import Univer、不调 `undo()` / `redo()`、不自己执行任何跳步；
 *   "撤销 N 次 + 重做 M 次"由上层拿 `onJumpTo` 的目标下标自己算（见 `./history-model`）。
 * - 无 hooks、无副作用：`open === false` 直接返回 `null`。
 *
 * 下标语义（与 `history-model` 完全一致，接线时最容易踩的一点）：
 * - `current` / `onJumpTo(index)` 用的都是"**已应用条数**"下标：0 = 最初状态，
 *   `i` = 已应用前 i 条，`entries.length` = 全部已应用；
 * - 列表项的**数组下标** `index` 对应的目标是 `index + 1`（第 index+1 步），
 *   因此点击 `history-item-2` 回调收到的是 `3`：跳转完成后这一条本身处于"已应用"。
 *   这样每一条的序号、摘要里的"第 N 步"、以及可跳到的状态才是同一个数。
 */
import { CloseIcon } from './icons';
import {
  canJump,
  clampIndex,
  describeKind,
  formatTime,
  jumpBlockReason,
  planJump,
  type HistoryEntry,
  type UndoRedoAvailability,
} from './history-model';
import './history-panel.css';

export interface HistoryPanelProps {
  open: boolean;
  entries: HistoryEntry[];
  /** 当前处于第几条（0 = 最初状态，entries.length = 全部已应用） */
  current: number;
  /** Univer 当前可用的撤销/重做次数（用于禁用不可达的跳步） */
  available: UndoRedoAvailability;
  onJumpTo: (index: number) => void;
  onClose: () => void;
}

/** 拼接类名，过滤空值（与本仓库其它组件同样的写法，不引第三方 clsx）。 */
function cls(...parts: (string | undefined | false)[]): string {
  return parts.filter((p): p is string => typeof p === 'string' && p !== '').join(' ');
}

export function HistoryPanel({
  open,
  entries,
  current,
  available,
  onJumpTo,
  onClose,
}: HistoryPanelProps): JSX.Element | null {
  if (!open) return null;

  const list = Array.isArray(entries) ? entries : [];
  const total = list.length;
  // 摘要里的 N 夹紧到 [0, total]：外部传进来的越界值不该让"第 7 / 5 步"这种文案露出去
  const at = Math.min(clampIndex(current), total);
  const avail: UndoRedoAvailability = { undos: clampIndex(available.undos), redos: clampIndex(available.redos) };

  return (
    <section className="hp-panel" data-testid="history-panel" role="region" aria-label="历史记录">
      <div className="hp-head">
        <strong className="hp-title">历史记录</strong>
        <span className="hp-summary" data-testid="history-summary">
          当前：第 {at} / {total} 步
        </span>
        <span className="hp-spacer" />
        <button
          type="button"
          className="hp-close"
          data-testid="history-close"
          aria-label="关闭历史记录"
          title="关闭"
          onClick={onClose}
        >
          <CloseIcon size={14} />
        </button>
      </div>

      {total === 0 ? (
        <div className="hp-empty" data-testid="history-empty">
          还没有可还原的操作
        </div>
      ) : (
        <ol className="hp-list">
          {list.map((entry, index) => {
            // 数组下标 index(0-based) ↔ 步下标 index + 1，见文件头"下标语义"
            const target = index + 1;
            const plan = planJump(at, target);
            const reason = jumpBlockReason(plan, avail);
            const reachable = canJump(plan, avail);
            const applied = target <= at;
            const isCurrent = target === at;
            const kind = describeKind(entry.kind);
            const verb = isCurrent ? '当前步' : applied ? '回到' : '重做到';

            return (
              <li className="hp-row" key={entry.id}>
                <button
                  type="button"
                  className={cls('hp-item', applied ? 'is-applied' : 'is-undone', isCurrent && 'is-current')}
                  data-testid={`history-item-${index}`}
                  data-step={target}
                  data-state={applied ? 'applied' : 'undone'}
                  aria-current={isCurrent ? 'step' : undefined}
                  disabled={!reachable}
                  // 不可达时 title 说明原因（如"需要重做 3 次，但只有 1 次可用"），可达时说明这一步做什么
                  title={reason ?? `${verb} · 第 ${target} 步 · ${entry.label}`}
                  onClick={() => onJumpTo(target)}
                >
                  <span className="hp-step">{target}</span>
                  <span className="hp-time">{formatTime(entry.at)}</span>
                  <span className={cls('hp-badge', `hp-badge-${kind.tone}`)}>{kind.label}</span>
                  <span className="hp-label">{entry.label}</span>
                </button>
              </li>
            );
          })}
        </ol>
      )}
    </section>
  );
}
