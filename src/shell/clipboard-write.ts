/**
 * 往系统剪贴板写**纯文本**，带"老式降级"兜底。
 *
 * 为什么需要这个文件（实测教训）：
 *  `navigator.clipboard` 只在**安全上下文**（HTTPS / localhost / 127.0.0.1）存在。
 *  用 `http://局域网IP` 打开给同事用时它是 `undefined`，而
 *  `navigator.clipboard?.writeText(text).then(...)` 会**整条短路** ——
 *  既不写剪贴板，也不弹"复制失败"，用户看到的就是"点了复制没反应"（静默失败）。
 *
 * 降级手段是 `textarea + document.execCommand('copy')`：它不要求安全上下文，
 * 也是 Univer 自己（`BrowserClipboardService._legacyCopyText`）在无 Clipboard API 时的兜底做法。
 *
 * 返回 `true/false` 而不是抛异常：调用方要据此给用户**如实**的提示，不能假装成功。
 */
export async function writeClipboardText(text: string): Promise<boolean> {
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      /* 权限被拒 / 文档失焦 → 继续走老式兜底 */
    }
  }
  return legacyCopyText(text);
}

/** 无 Clipboard API（http 局域网、老浏览器）时的兜底：临时 textarea + execCommand */
function legacyCopyText(text: string): boolean {
  try {
    const area = document.createElement('textarea');
    area.value = text;
    area.setAttribute('readonly', '');
    // 必须在视口内且可选中，否则 execCommand('copy') 拿不到选区；
    // 用 fixed + 负偏移把它挪出可见区域，同时避免页面跳动。
    area.style.position = 'fixed';
    area.style.top = '-1000px';
    area.style.left = '-1000px';
    area.style.opacity = '0';
    const active = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    document.body.appendChild(area);
    try {
      area.select();
      area.setSelectionRange(0, text.length);
      return document.execCommand('copy');
    } finally {
      area.remove();
      active?.focus?.();
    }
  } catch {
    return false;
  }
}
