/**
 * 批注（legacy note）解析：`xl/commentsN.xml`。
 *
 * 结构：`<comments><authors><author>…</author></authors>
 *        <commentList><comment ref="A1" authorId="0"><text>…</text></comment>…`
 *
 * 文本还原要点（实测 fixture-extras.xlsx）：
 * - `<text>` 下既可能是**多个 `<r>` run**（每个 run 一个 `<t>`，可能还带 `<rPr>`），
 *   也可能是直接的 `<t>`；两种都要处理，且 run 之间**直接拼接**（run 是格式分段，不是段落）。
 * - 换行是 `<t xml:space="preserve">第一行\n第二行</t>` 里的**字面换行**（Excel 用 `&#10;`
 *   或裸 `\n`），必须原样保留——`xml:space="preserve"` 只是"别裁空白"，不是换行来源。
 *   解码实体（`&#10;` -> `\n`）由 `elementText`/`decodeEntities` 负责，这里不做 trim。
 */
import type { ParsedNote } from './types';
import type { Warn } from './worksheet';
import { childElements, elementText, findFirstElement, localName, type ElementRange } from './xml';

export interface ParseCommentsOptions {
  warn?: Warn;
  /** 部件路径，仅用于 warning 文案 */
  partPath?: string;
}

/** 解析一个 `xl/commentsN.xml` -> `ParsedNote[]`（按文件里的顺序） */
export function parseComments(xml: string, opts: ParseCommentsOptions = {}): ParsedNote[] {
  const warn: Warn = opts.warn ?? (() => {});
  const where = opts.partPath ?? 'comments.xml';

  const authors = parseAuthors(xml);
  const listEl = findFirstElement(xml, 'commentList');
  if (!listEl) {
    warn(`${where} 缺少 <commentList>，没有可用的批注`);
    return [];
  }

  const out: ParsedNote[] = [];
  for (const el of childElements(xml, listEl)) {
    if (localName(el.name) !== 'comment') continue;
    const ref = el.attrs['ref'];
    if (ref === undefined || ref === '') {
      warn(`${where} 有一条批注缺少 ref，已跳过`);
      continue;
    }
    const note: ParsedNote = { ref, text: readCommentText(xml, el, warn, where, ref) };
    const authorId = el.attrs['authorId'];
    if (authorId !== undefined) {
      const idx = Number.parseInt(authorId, 10);
      if (Number.isFinite(idx) && idx >= 0 && idx < authors.length) {
        note.author = authors[idx];
      } else {
        warn(`${where} 的批注 ${ref} 引用了越界的 authorId=${authorId}（authors 共 ${authors.length} 项）`);
      }
    }
    out.push(note);
  }
  return out;
}

function parseAuthors(xml: string): string[] {
  const authorsEl = findFirstElement(xml, 'authors');
  if (!authorsEl) return [];
  const out: string[] = [];
  for (const el of childElements(xml, authorsEl)) {
    if (localName(el.name) !== 'author') continue;
    out.push(elementText(xml, el));
  }
  return out;
}

/** `<text>` -> 纯文本；`<r><t>` 与直接 `<t>` 都支持，**不做 trim**（保留换行与首尾空格） */
function readCommentText(
  xml: string,
  commentEl: ElementRange,
  warn: Warn,
  where: string,
  ref: string,
): string {
  const textEl = childElements(xml, commentEl).find((c) => localName(c.name) === 'text');
  if (!textEl) {
    warn(`${where} 的批注 ${ref} 没有 <text>，按空文本处理`);
    return '';
  }
  let out = '';
  for (const kid of childElements(xml, textEl)) {
    const ln = localName(kid.name);
    if (ln === 't') {
      out += elementText(xml, kid);
    } else if (ln === 'r') {
      // run：取第一个 <t>，跳过 <rPr>（字体信息不属于中立模型）
      for (const runKid of childElements(xml, kid)) {
        if (localName(runKid.name) === 't') {
          out += elementText(xml, runKid);
          break;
        }
      }
    }
  }
  return out;
}
