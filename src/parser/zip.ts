/**
 * xlsx 容器（ZIP / OPC 包）解包。
 *
 * 只依赖 fflate，不碰 `fs`，输入输出都是内存中的 `Uint8Array`，
 * 因此同一份代码可以跑在 Node、Worker、浏览器里。
 */
import { unzipSync, strFromU8, decompressSync, inflateSync } from 'fflate';

/** 条目名 -> 原始字节（条目名保留 xlsx 内的原始写法，如 `xl/worksheets/sheet1.xml`） */
export type ZipEntries = Record<string, Uint8Array>;

const utf8Decoder = new TextDecoder('utf-8');

/** 解压 xlsx 字节流，返回全部 zip 条目（含我们暂不解析的部件，供后续外科式修补用） */
export function readZipEntries(data: Uint8Array | ArrayBuffer): ZipEntries {
  const bytes = toUint8Array(data);
  const raw = unzipSync(bytes);
  const entries: ZipEntries = {};
  for (const name of Object.keys(raw)) {
    entries[name] = raw[name];
  }
  return entries;
}

export function toUint8Array(data: Uint8Array | ArrayBuffer): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/** 条目内容按 UTF-8 解码为文本；条目不存在返回 `undefined` */
export function readText(entries: ZipEntries, path: string): string | undefined {
  const bytes = findEntry(entries, path);
  return bytes === undefined ? undefined : decodeText(bytes);
}

/** 条目内容为二进制；条目不存在返回 `undefined` */
export function readBytes(entries: ZipEntries, path: string): Uint8Array | undefined {
  return findEntry(entries, path);
}

export function decodeText(bytes: Uint8Array): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return utf8Decoder.decode(bytes.subarray(3)); // 去 BOM
  }
  return utf8Decoder.decode(bytes);
}

export function hasEntry(entries: ZipEntries, path: string): boolean {
  return findEntry(entries, path) !== undefined;
}

/**
 * 查找条目：xlsx 条目名大小写不敏感匹配（少数写出的工具会改大小写），
 * 同时容忍前导 `/` 与 `./`。
 */
export function findEntry(entries: ZipEntries, path: string): Uint8Array | undefined {
  const direct = entries[path];
  if (direct !== undefined) return direct;
  const normalized = normalizePath(path);
  const directNorm = entries[normalized];
  if (directNorm !== undefined) return directNorm;
  for (const key of Object.keys(entries)) {
    if (normalizePath(key) === normalized) return entries[key];
  }
  return undefined;
}

export function normalizePath(path: string): string {
  let p = path.replace(/\\/g, '/');
  if (p.startsWith('/')) p = p.slice(1);
  while (p.startsWith('./')) p = p.slice(2);
  return p;
}

/** Excel 相对路径解析（`xl/workbook.xml` + `worksheets/sheet1.xml` -> `xl/worksheets/sheet1.xml`） */
export function resolvePath(baseDir: string, target: string): string {
  const t = normalizePath(target);
  if (t.startsWith('/')) return t.slice(1);
  const baseParts = normalizePath(baseDir).split('/').filter((s) => s.length > 0);
  const parts = t.split('/');
  for (const part of parts) {
    if (part === '' || part === '.') continue;
    if (part === '..') baseParts.pop();
    else baseParts.push(part);
  }
  return baseParts.join('/');
}

export function dirName(path: string): string {
  const p = normalizePath(path);
  const slash = p.lastIndexOf('/');
  return slash < 0 ? '' : p.slice(0, slash);
}

/* -------------------------------------------------------------------------- */
/* 单条目兜底解压（用于 readZipEntries 之外的场景，如从 raw bytes 直接取工作表）    */
/* -------------------------------------------------------------------------- */

/**
 * 兜底解压：正常情况下统一走 `readZipEntries`，这里只处理"已知是 deflate 流"的场景。
 */
export function inflateMaybe(bytes: Uint8Array): Uint8Array {
  try {
    return decompressSync(bytes);
  } catch {
    try {
      return inflateSync(bytes);
    } catch {
      return bytes;
    }
  }
}

export { strFromU8 };
