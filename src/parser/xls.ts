/**
 * **`.xls`（Excel 97–2003，BIFF8 二进制）解析器。**
 *
 * 产出中性的 `WorkbookInput`（见 `src/importer/synth-xlsx.ts`），随后被打成一份规范 xlsx，
 * 于是预览/编辑/工作区/撤销/导出/会话恢复全部复用既有链路 —— 与 CSV 一样，`.xls` 只是一条
 * "入口翻译"。**不引第三方库**：`.xls` 既不是 zip 也不是 XML，`fflate` / `xml.ts` 都帮不上忙，
 * 所以这里手写两层：
 *
 *  1. **CFB（Compound File Binary / OLE2 复合文档）**：512 字节头 → FAT（`_sectFat` + DIFAT）→
 *     目录项（128 字节/项）→ 找名为 `Workbook`（BIFF8）或 `Book`（BIFF5）的流。
 *     小于 4096 字节（`_ulMiniSectorCutoff`）的流走**迷你流 + 迷你 FAT**，所以"读一条流"必须
 *     分两路：普通 FAT 扇区链、迷你扇区链。**扇区链必须按 FAT 逐扇区拼接**：大工作簿的流一定
 *     跨很多扇区，而且可以是碎片化的非连续链（真夹具就是 `1..38 → 39..46 → 47..54 → 56`）。
 *  2. **BIFF8 记录流**：`[u16 id][u16 len][payload]`，再按 `BOF`/`EOF` 切成 globals 与各工作表子流。
 *
 * 三个真踩过的坑（都写进了测试）：
 *  - **CFB 头里"首个目录扇区"在偏移 48，不是 44**：偏移 44 是 `_csectDir`（目录扇区数，
 *    BIFF8 里恒为 0），偏移 48 才是 `_sectDir`。读错会把 BIFF 数据当目录项，表现为
 *    "文件里明明有 `Workbook` 流却找不到"。
 *  - **FAT 链上的特殊值不能当"链结束"**：`0xFFFFFFFD` 是 FATSECT（该扇区装 FAT 本身）、
 *    `0xFFFFFFFC` 是 DIFSECT，它们只说明"这个扇区是 FAT/DIFAT"，**不表示链到此为止**。
 *    只有 `0xFFFFFFFE`（ENDOFCHAIN）与 `0xFFFFFFFF`（FREESECT）才是结束。
 *    真夹具 `fixture-styles.xls` 的目录链就是 `…→54→（跳过 FAT 扇区 55）→56`，
 *    把 FATSECT 当结束会直接丢掉整个目录。
 *  - **SST 的 CONTINUE 与普通记录的 CONTINUE 规则不同**：普通记录跨 CONTINUE 只是把字节接起来；
 *    而 `SST` 里一条字符串如果正好跨过 CONTINUE 边界，**新 CONTINUE 的第一个字节不是字符数据，
 *    而是这一段的"压缩/非压缩"标志位**，必须每段重新读一次。SST 头部的 `cstTotal`/`cstUnique`
 *    也只在第一段里。
 *
 * 已知不支持（明确检测 + 如实报告，绝不硬猜）：
 *  - **加密工作簿**（`FILEPASS 0x002F`）：加密后 SST / 样式 / 单元格记录全被混淆，没有"部分解析"
 *    的可能 —— 直接返回空工作簿 + 中文警告，由上层提示用户；
 *  - **BIFF5/BIFF7/BIFF4**（`BOF` 版本 < 0x0600）：字符串是 8 位代码页而非 Unicode，产物会大面积
 *    乱码，因此不解析，返回空工作簿 + 中文警告；
 *  - **图表工作表 / 宏工作表 / 对话框工作表**（子流类型不是 0x0010）：跳过并逐个报告；
 *  - **公式体**：`FORMULA` 的 `rgce` 是 RPN 字节码，本模块**不反编译**，只搬缓存值
 *    （与既有"导入时用 Excel 缓存值渲染"的约定一致）；结果是错误值（`#DIV/0!` 等）时
 *    中性模型没有表达形式，如实丢弃并计数；
 *  - **富文本分段**：SST 字符串的 rich/phonetic 段按长度跳过，只取纯文本；
 *  - **条件格式 / 数据验证 / 批注 / 超链接 / 图片 / 透视表**：`.xls` 里分别由
 *    `CONDFMT`+`CF`、`DVAL`+`DV`、`NOTE`、`HLINK`、`MSODRAWING`+`OBJ`、`SX*` 承载，本模块不解析。
 */
import { BUILTIN_NUM_FMTS, INDEXED_COLORS } from './styles';
import type { SynthCell, SynthSheet, SynthStyle, WorkbookInput } from '../importer/synth-xlsx';

/* -------------------------------------------------------------------------- */
/* 常量与安全阀                                                                */
/* -------------------------------------------------------------------------- */

/** 单元格总量上限：超出后停止写入并如实报告（与 CSV 解析器同一口径） */
export const MAX_XLS_CELLS = 200_000;
/** 单表行列上限（与 xlsx 规范一致） */
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
/** CFB 头的长度；扇区数据从 512 字节之后开始 */
const CFB_HEADER_SIZE = 512;
/** 目录项长度 */
const DIR_ENTRY_SIZE = 128;
/** 一条扇区链最多走多少步（防御环状链；远大于任何真实文件） */
const MAX_CHAIN_STEPS = 1 << 20;
/** 默认扇区 512 字节、默认迷你扇区 64 字节 */
const DEFAULT_SECTOR_SHIFT = 9;
const DEFAULT_MINI_SECTOR_SHIFT = 6;

/* CFB 特殊扇区号 */
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
/** 所有 >= 该值（且不是上面两个）的取值都不是普通数据扇区号 */
const MAXREGSECT = 0xfffffffa;

/* BIFF 记录号（只列本模块认识的） */
const REC = {
  BOF: 0x0809,
  EOF: 0x000a,
  FILEPASS: 0x002f,
  BOUNDSHEET: 0x0085,
  SST: 0x00fc,
  CONTINUE: 0x003c,
  DATEMODE: 0x0022,
  FONT: 0x0031,
  FORMAT: 0x041e,
  XF: 0x00e0,
  PALETTE: 0x0092,
  LABELSST: 0x00fd,
  LABEL: 0x0204,
  RSTRING: 0x00d6,
  RK: 0x027e,
  MULRK: 0x00bd,
  NUMBER: 0x0203,
  BOOLERR: 0x0205,
  FORMULA: 0x0006,
  STRING: 0x0207,
  BLANK: 0x0201,
  MULBLANK: 0x00be,
  MERGEDCELLS: 0x00e5,
  ROW: 0x0208,
  COLINFO: 0x007d,
  PANE: 0x0041,
  WINDOW2: 0x023e,
  DEFCOLWIDTH: 0x0055,
  DEFAULTROWHEIGHT: 0x0225,
} as const;

/** BOF 子流类型（`BOF` 偏移 2 的 u16） */
const SUBSTREAM = { WORKBOOK_GLOBALS: 0x0005, WORKSHEET: 0x0010, CHART: 0x0020, MACRO: 0x0040 } as const;

/* -------------------------------------------------------------------------- */
/* RK 解码                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * 解码一个 `RK` 值（4 字节，按 i32 存放）。
 *
 * 为什么单独导出并配单元测试：RK 把"整数/浮点 × 原值/百分之一"四种情况压进 32 位，
 * 位运算写错一位就**静默**得到错数字（不抛异常、不越界），只能靠逐条断言兜住。
 *  - bit0（`0x01`）置位：真实值是解码结果的 **1/100**（Excel 用它存百分比与两位小数）；
 *  - bit1（`0x02`）置位：高 30 位是**有符号整数**（`i32 >> 2`，算术右移保住负号）；
 *  - bit1 清零：高 30 位是 **IEEE754 双精度的最高 30 位**，低 34 位补 0 ——
 *    把小端 i32 放到 8 字节缓冲的**高 4 字节**（低 4 字节保持 0）再按双精度读，即可无损还原。
 */
export function decodeRk(raw: number): number {
  const isDiv100 = (raw & 0x01) !== 0;
  const isInteger = (raw & 0x02) !== 0;
  let value: number;
  if (isInteger) {
    value = raw >> 2; // 算术右移：负数也正确
  } else {
    const view = new DataView(new ArrayBuffer(8));
    view.setUint32(4, raw & 0xfffffffc, true);
    value = view.getFloat64(0, true);
  }
  return isDiv100 ? value / 100 : value;
}

/* -------------------------------------------------------------------------- */
/* 越界安全的字节读取                                                          */
/* -------------------------------------------------------------------------- */

function readU8(bytes: Uint8Array, offset: number): number | undefined {
  if (offset < 0 || offset >= bytes.length) return undefined;
  return bytes[offset];
}

function readU16(bytes: Uint8Array, offset: number): number | undefined {
  if (offset < 0 || offset + 2 > bytes.length) return undefined;
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes: Uint8Array, offset: number): number | undefined {
  if (offset < 0 || offset + 4 > bytes.length) return undefined;
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

function readF64(bytes: Uint8Array, offset: number): number | undefined {
  if (offset < 0 || offset + 8 > bytes.length) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset + offset, 8);
  return view.getFloat64(0, true);
}

function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
  let total = 0;
  for (const chunk of chunks) total += chunk.length;
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

/** UTF-16LE 解码（CFB 目录项名与 BIFF8 非压缩字符串共用） */
function decodeUtf16Le(bytes: Uint8Array, offset: number, length: number): string {
  const end = Math.min(offset + length, bytes.length);
  let out = '';
  for (let i = offset; i + 1 < end; i += 2) out += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
  return out;
}

/**
 * 8 位字符解码。
 *
 * 说明：BIFF8 的"压缩"字符串（flags 高位 = 0）**按规范就是 Latin-1**（每字节一个码位），
 * 不是代码页 —— 所以这里用一一映射而不是 `TextDecoder('gbk')`：字符集信息在 FONT 记录的
 * `grbit`/`chs` 里，逐串判代码页只会引入歧义。
 */
function decodeLatin1(bytes: Uint8Array, offset: number, length: number): string {
  const end = Math.min(offset + length, bytes.length);
  let out = '';
  for (let i = offset; i < end; i += 1) out += String.fromCharCode(bytes[i]);
  return out;
}

/* -------------------------------------------------------------------------- */
/* CFB（复合文档）层                                                            */
/* -------------------------------------------------------------------------- */

interface CfbDirectoryEntry {
  name: string;
  type: number;
  startSector: number;
  size: number;
}

interface CfbContainer {
  sectorSize: number;
  /** 文件里实际存在的扇区数（越界检测用） */
  sectorCount: number;
  fat: number[];
  miniFat: number[];
  miniStream: Uint8Array;
  miniSectorSize: number;
  miniCutoff: number;
  entries: CfbDirectoryEntry[];
}

/**
 * 走一条**普通**扇区链，把各扇区原样拼接。
 *
 * 防御四件事：越界扇区号（文件被截断）、环状链、超长链、FAT 缺项 ——
 * 全部是"截断 + 中文警告"，绝不 `while (true)`。
 */
function readSectorChain(
  container: CfbContainer,
  startSector: number,
  raw: Uint8Array,
  warnings: string[],
  label: string,
): Uint8Array {
  const chunks: Uint8Array[] = [];
  const seen = new Set<number>();
  let sector = startSector;
  let steps = 0;
  let stopReason = '';

  while (sector <= MAXREGSECT) {
    if (steps >= MAX_CHAIN_STEPS) {
      stopReason = `链长超过 ${MAX_CHAIN_STEPS} 步`;
      break;
    }
    if (seen.has(sector)) {
      stopReason = `扇区链成环（回到扇区 ${sector}）`;
      break;
    }
    seen.add(sector);
    if (sector >= container.sectorCount) {
      stopReason = `扇区 ${sector} 越界（文件只有 ${container.sectorCount} 个扇区）`;
      break;
    }
    const begin = CFB_HEADER_SIZE + sector * container.sectorSize;
    chunks.push(raw.subarray(begin, Math.min(begin + container.sectorSize, raw.length)));
    const next = container.fat[sector];
    if (next === undefined) {
      stopReason = `FAT 里没有扇区 ${sector} 的项`;
      break;
    }
    // 只有 ENDOFCHAIN / FREESECT 才是"链结束"；FATSECT/DIFSECT 只说明该扇区装的是 FAT/DIFAT，
    // 出现在链上时必须跳过并继续（真夹具 fixture-styles.xls 的目录链是 …→54→55(FAT)→56）。
    if (next === ENDOFCHAIN || next === FREESECT) break;
    sector = next;
    steps += 1;
  }

  if (stopReason) warnings.push(`CFB ${label}的扇区链不完整：${stopReason}，已按已读到的部分处理`);
  return concatBytes(chunks);
}

/** 走一条**迷你**扇区链（迷你流内部的链） */
function readMiniChain(
  container: CfbContainer,
  startSector: number,
  warnings: string[],
  label: string,
): Uint8Array {
  const size = container.miniSectorSize;
  const mini = container.miniStream;
  const totalSectors = Math.floor(mini.length / size);
  const chunks: Uint8Array[] = [];
  const seen = new Set<number>();
  let sector = startSector;
  let steps = 0;
  let stopReason = '';

  while (sector <= MAXREGSECT) {
    if (steps >= MAX_CHAIN_STEPS) {
      stopReason = '迷你扇区链过长';
      break;
    }
    if (seen.has(sector)) {
      stopReason = `迷你扇区链成环（回到迷你扇区 ${sector}）`;
      break;
    }
    seen.add(sector);
    if (sector >= totalSectors) {
      stopReason = `迷你扇区 ${sector} 越界（迷你流只有 ${totalSectors} 个迷你扇区）`;
      break;
    }
    chunks.push(mini.subarray(sector * size, (sector + 1) * size));
    const next = container.miniFat[sector];
    if (next === undefined) {
      stopReason = `迷你 FAT 里没有迷你扇区 ${sector} 的项`;
      break;
    }
    if (next === ENDOFCHAIN || next === FREESECT) break;
    sector = next;
    steps += 1;
  }

  if (stopReason) warnings.push(`CFB ${label}的迷你扇区链不完整：${stopReason}，已按已读到的部分处理`);
  return concatBytes(chunks);
}

/**
 * 解析 CFB 容器：头 → DIFAT/FAT → 目录 → 迷你流/迷你 FAT。
 *
 * 返回 `undefined` 表示"根本不是 CFB"（签名不对或文件太短）；其余问题都是
 * "尽量把能读的读出来 + 中文警告"，因为真文件偶尔带冗余扇区。
 */
function readCfb(bytes: Uint8Array, warnings: string[]): CfbContainer | undefined {
  if (bytes.length < CFB_HEADER_SIZE) {
    warnings.push(`文件只有 ${bytes.length} 字节，不足 CFB 头的 512 字节，不是 OLE2 复合文档`);
    return undefined;
  }
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  for (let i = 0; i < signature.length; i += 1) {
    if (bytes[i] !== signature[i]) {
      warnings.push('文件头不是 OLE2 复合文档签名（D0 CF 11 E0 A1 B1 1A E1），不是 .xls');
      return undefined;
    }
  }

  const sectorShift = readU16(bytes, 30) ?? DEFAULT_SECTOR_SHIFT;
  const miniSectorShift = readU16(bytes, 32) ?? DEFAULT_MINI_SECTOR_SHIFT;
  const validSectorShift = sectorShift >= 7 && sectorShift <= 20;
  const validMiniShift = miniSectorShift >= 2 && miniSectorShift <= 20;
  if (!validSectorShift) warnings.push(`CFB 头的扇区大小移位 ${sectorShift} 不在 7..20，按 512 字节处理`);
  if (!validMiniShift) warnings.push(`CFB 头的迷你扇区移位 ${miniSectorShift} 不在 2..20，按 64 字节处理`);

  const sectorSize = 1 << (validSectorShift ? sectorShift : DEFAULT_SECTOR_SHIFT);
  const miniSectorSize = 1 << (validMiniShift ? miniSectorShift : DEFAULT_MINI_SECTOR_SHIFT);
  const sectorCount = Math.floor((bytes.length - CFB_HEADER_SIZE) / sectorSize);
  const miniCutoff = readU32(bytes, 56) ?? 4096;
  // 注意：48 才是 _sectDir（首个目录扇区），44 是 _csectDir（目录扇区数，BIFF8 恒为 0）
  const firstDirSector = readU32(bytes, 48) ?? ENDOFCHAIN;
  const firstMiniFatSector = readU32(bytes, 60) ?? ENDOFCHAIN;
  const firstDifatSector = readU32(bytes, 68) ?? ENDOFCHAIN;
  const numDifatSectors = readU32(bytes, 72) ?? 0;

  /** 只按扇区号取原始字节，不解析链（FAT/DIFAT 自身用） */
  const rawSector = (sector: number): Uint8Array | undefined => {
    if (sector < 0 || sector >= sectorCount) return undefined;
    const begin = CFB_HEADER_SIZE + sector * sectorSize;
    return bytes.subarray(begin, Math.min(begin + sectorSize, bytes.length));
  };

  // 1) DIFAT：头里 109 项 + 后续 DIFAT 扇区
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i += 1) {
    const value = readU32(bytes, 76 + i * 4);
    if (value === undefined || value === FREESECT) break;
    if (value <= MAXREGSECT) fatSectors.push(value);
  }
  let difatSector = firstDifatSector;
  let difatSteps = 0;
  const difatSeen = new Set<number>();
  while (difatSector <= MAXREGSECT && difatSteps <= numDifatSectors + 8) {
    if (difatSeen.has(difatSector)) {
      warnings.push(`CFB 的 DIFAT 链成环（扇区 ${difatSector}），已停止读取 DIFAT`);
      break;
    }
    difatSeen.add(difatSector);
    const data = rawSector(difatSector);
    if (!data) {
      warnings.push(`CFB 的 DIFAT 扇区 ${difatSector} 越界，已停止读取 DIFAT`);
      break;
    }
    const perSector = Math.floor(sectorSize / 4) - 1;
    for (let i = 0; i < perSector; i += 1) {
      const value = readU32(data, i * 4);
      if (value === undefined) break;
      if (value <= MAXREGSECT) fatSectors.push(value);
    }
    const next = readU32(data, perSector * 4);
    if (next === undefined || next === ENDOFCHAIN || next === FREESECT) break;
    difatSector = next;
    difatSteps += 1;
  }

  // 2) FAT
  const fat: number[] = [];
  for (const sector of fatSectors) {
    const data = rawSector(sector);
    if (!data) {
      warnings.push(`CFB 的 FAT 扇区 ${sector} 越界，已跳过`);
      continue;
    }
    for (let i = 0; i + 4 <= data.length; i += 4) fat.push(readU32(data, i) ?? FREESECT);
  }
  if (fat.length === 0) warnings.push('CFB 没有任何可用的 FAT 扇区，无法定位数据流');

  const container: CfbContainer = {
    sectorSize,
    sectorCount,
    fat,
    miniFat: [],
    miniStream: new Uint8Array(0),
    miniSectorSize,
    miniCutoff,
    entries: [],
  };

  // 3) 目录
  const dirData = readSectorChain(container, firstDirSector, bytes, warnings, '目录');
  const entries: CfbDirectoryEntry[] = [];
  for (let at = 0; at + DIR_ENTRY_SIZE <= dirData.length; at += DIR_ENTRY_SIZE) {
    const type = dirData[at + 66];
    if (type !== 1 && type !== 2 && type !== 5) continue; // 0 = 空项
    const nameLength = readU16(dirData, at + 64) ?? 0;
    if (nameLength < 2 || nameLength > 64) continue;
    entries.push({
      name: decodeUtf16Le(dirData, at, nameLength - 2),
      type,
      startSector: readU32(dirData, at + 116) ?? ENDOFCHAIN,
      size: readU32(dirData, at + 120) ?? 0,
    });
  }
  container.entries = entries;
  if (entries.length === 0) warnings.push('CFB 目录里没有读到任何有效目录项');

  // 4) 根项（根存储）的流 = 迷你流的容器；再读迷你 FAT
  const root = entries.find((entry) => entry.type === 5);
  if (root && root.size > 0) {
    container.miniStream = readSectorChain(container, root.startSector, bytes, warnings, '迷你流').subarray(0, root.size);
  }
  if (firstMiniFatSector <= MAXREGSECT) {
    const miniFatData = readSectorChain(container, firstMiniFatSector, bytes, warnings, '迷你 FAT');
    for (let i = 0; i + 4 <= miniFatData.length; i += 4) container.miniFat.push(readU32(miniFatData, i) ?? FREESECT);
  }

  return container;
}

/** 读一条流：`size >= miniCutoff` 走普通 FAT 链，否则走迷你流 + 迷你 FAT 链 */
function readCfbStream(container: CfbContainer, entry: CfbDirectoryEntry, bytes: Uint8Array, warnings: string[]): Uint8Array {
  if (entry.size === 0) return new Uint8Array(0);
  const data = entry.size >= container.miniCutoff
    ? readSectorChain(container, entry.startSector, bytes, warnings, `流 "${entry.name}"`)
    : readMiniChain(container, entry.startSector, warnings, `流 "${entry.name}"`);
  if (data.length < entry.size) {
    warnings.push(`CFB 流 "${entry.name}" 声明 ${entry.size} 字节，实际只读到 ${data.length} 字节（文件可能被截断）`);
  }
  return data.length > entry.size ? data.subarray(0, entry.size) : data;
}

/* -------------------------------------------------------------------------- */
/* BIFF 记录流                                                                 */
/* -------------------------------------------------------------------------- */

interface BiffRecord {
  id: number;
  /** 记录头在流里的偏移（`BOUNDSHEET` 的 `lbPlyPos` 指向它） */
  offset: number;
  dataOffset: number;
  dataLength: number;
}

/**
 * 扫描记录表。遇到长度越界的记录（文件被截断 / 记录头是垃圾）就停下并记录原因 ——
 * 任何畸形输入都只会"少读一点"，不会死循环也不会抛异常。
 */
function scanRecords(stream: Uint8Array, warnings: string[]): BiffRecord[] {
  const records: BiffRecord[] = [];
  let at = 0;
  while (at + 4 <= stream.length) {
    const id = readU16(stream, at) ?? 0;
    const length = readU16(stream, at + 2) ?? 0;
    if (at + 4 + length > stream.length) {
      warnings.push(
        `BIFF 记录流在偏移 ${at} 处被截断（记录 0x${id.toString(16)} 声明长度 ${length}，`
        + `剩余 ${stream.length - at - 4} 字节），后续记录已忽略`,
      );
      break;
    }
    records.push({ id, offset: at, dataOffset: at + 4, dataLength: length });
    if (records.length > 400_000) {
      warnings.push('BIFF 记录数超过 40 万条，疑似记录头损坏，已停止扫描');
      break;
    }
    at += 4 + length;
  }
  return records;
}

function recordData(stream: Uint8Array, record: BiffRecord): Uint8Array {
  return stream.subarray(record.dataOffset, record.dataOffset + record.dataLength);
}

/** "主记录 + 紧随其后的 CONTINUE 段"：各段保持独立，因为 SST 的续接要按段处理 */
interface ChunkedPayload {
  chunks: Uint8Array[];
  /** CONTINUE 的条数（主记录不算） */
  continues: number;
}

function readChunked(stream: Uint8Array, records: readonly BiffRecord[], index: number): ChunkedPayload {
  const chunks = [recordData(stream, records[index])];
  let continues = 0;
  let next = index + 1;
  while (next < records.length && records[next].id === REC.CONTINUE) {
    chunks.push(recordData(stream, records[next]));
    continues += 1;
    next += 1;
  }
  return { chunks, continues };
}

/* -------------------------------------------------------------------------- */
/* BIFF8 字符串                                                                */
/* -------------------------------------------------------------------------- */

interface UnicodeStringResult {
  text: string;
  /** 读完后落在哪一段、段内偏移（SST 需要接着往下读） */
  chunkIndex: number;
  position: number;
  truncated: boolean;
}

/**
 * 读一条 **BIFF8 Unicode 字符串**（SST 用；rich/phonetic 段读完后自行跳过）。
 *
 * 布局（`XLUnicodeRichExtendedString`）：
 * `u16 cch` + `u8 flags`（`0x01` = 16 位字符、`0x04` = 有 `cExtRst`、`0x08` = 有 `cRun`）
 * + 可选 `u16 cRun` + 可选 `u32 cbExtRst` + 字符数据 + `cRun*4` 字节 rich 段 + `cbExtRst` 字节 phonetic 段。
 *
 * **跨 CONTINUE 的坑**：字符数据一旦跨段，新段的**首字节是这一段的编码标志**
 * （`0x01` = 16 位、`0x00` = 8 位），必须每段重新读一次，不能沿用上一段的编码。
 * 这也是 SST 的 CONTINUE 与普通记录 CONTINUE 的唯一区别。
 */
function readUnicodeString(chunks: readonly Uint8Array[], startChunk: number, startPosition: number): UnicodeStringResult {
  let ci = startChunk;
  let pos = startPosition;
  const done = (text: string, truncated: boolean): UnicodeStringResult => ({ text, chunkIndex: ci, position: pos, truncated });
  const advance = (need: number): boolean => {
    while (ci < chunks.length && pos + need > chunks[ci].length) {
      ci += 1;
      pos = 0;
    }
    return ci < chunks.length;
  };

  if (!advance(3)) return done('', true);
  const cch = readU16(chunks[ci], pos) ?? 0;
  const flags = chunks[ci][pos + 2];
  pos += 3;
  let runCount = 0;
  let extSize = 0;
  if ((flags & 0x08) !== 0) {
    if (!advance(2)) return done('', true);
    runCount = readU16(chunks[ci], pos) ?? 0;
    pos += 2;
  }
  if ((flags & 0x04) !== 0) {
    if (!advance(4)) return done('', true);
    extSize = readU32(chunks[ci], pos) ?? 0;
    pos += 4;
  }

  let highByte = (flags & 0x01) !== 0;
  let text = '';
  let remaining = cch;
  let truncated = false;
  while (remaining > 0) {
    if (ci >= chunks.length) {
      truncated = true;
      break;
    }
    const chunk = chunks[ci];
    if (pos >= chunk.length) {
      // 字符数据正好用尽当前段：下一段首字节是编码标志
      ci += 1;
      pos = 0;
      if (ci >= chunks.length) {
        truncated = true;
        break;
      }
      highByte = (chunks[ci][pos] & 0x01) !== 0;
      pos += 1;
      continue;
    }
    if (highByte) {
      if (pos + 2 > chunk.length) {
        // 一个 16 位字符被段边界劈开：BIFF 不允许这种切法，只能按截断处理
        truncated = true;
        break;
      }
      text += String.fromCharCode(readU16(chunk, pos) ?? 0);
      pos += 2;
    } else {
      text += String.fromCharCode(chunk[pos]);
      pos += 1;
    }
    remaining -= 1;
  }

  // 跳过 rich run 与 phonetic 段（同样可能跨 CONTINUE，按段推进）
  let skip = runCount * 4 + extSize;
  while (skip > 0 && ci < chunks.length) {
    const available = chunks[ci].length - pos;
    if (available <= 0) {
      ci += 1;
      pos = 0;
      continue;
    }
    const take = Math.min(available, skip);
    pos += take;
    skip -= take;
  }
  return done(text, truncated);
}

/**
 * 读短 Unicode 字符串（`u8 cch` + `u8 flags` + 字符），用于 `BOUNDSHEET` 表名。
 * `cch` 最高位是"非压缩"标志位（BIFF8 的 `ShortXLUnicodeString` 允许两种写法）。
 */
function readShortUnicodeString(bytes: Uint8Array, offset: number): string {
  const rawLength = readU8(bytes, offset);
  if (rawLength === undefined) return '';
  const cch = rawLength & 0x7f;
  // 最高位置位表示"这是个 Unicode 串"（无论低 7 位怎么算），此时后面的 flags 字节可能不存在；
  // 但 BOUNDSHEET 一律带 flags 字节，所以这里按"cch + flags + 字符"的规范布局读。
  const flags = readU8(bytes, offset + 1) ?? 0;
  return (flags & 0x01) !== 0 || (rawLength & 0x80) !== 0
    ? decodeUtf16Le(bytes, offset + 2, cch * 2)
    : decodeLatin1(bytes, offset + 2, cch);
}

/**
 * 读 `FONT` 记录里的字体名。
 *
 * 为什么单独一个函数：`FONT` 的定长头在实现里是 **14 字节**（到 `unused3` 为止），
 * 紧接着是 `u8 cch` + `u8 flags` + 字符数据；真夹具上 `cch = <字符数>` 落在偏移 14、
 * `flags = 0x01` 落在偏移 15、字符从偏移 16 开始。若按"[MS-XLS] 的文字字段表"把
 * `bCharSet` 当成偏移 14 的独立字节去读，就会把"字符数"误当字符集、把名字整体错位一格，
 * 得到「卛」这样的乱码（真踩过）。这里按"数据自洽"来判：先用 `cch @14 + flags @15`
 * 试解，若长度对不上就在 `cch @15 + flags @16` 再试一次，两次都拿不出可打印字符串时留空。
 */
function readFontName(data: Uint8Array): string {
  const layouts: Array<[number, number]> = [
    [14, 15], // 真夹具的实际布局
    [15, 16], // [MS-XLS] 字段表所描述的布局
  ];
  for (const [countAt, flagAt] of layouts) {
    const rawCount = readU8(data, countAt);
    const flags = readU8(data, flagAt);
    if (rawCount === undefined || flags === undefined) continue;
    const cch = rawCount & 0x7f;
    if (cch === 0) continue;
    const textOffset = flagAt + 1;
    const needed = (flags & 0x01) !== 0 ? cch * 2 : cch;
    if (textOffset + needed > data.length) continue;
    const text = (flags & 0x01) !== 0
      ? decodeUtf16Le(data, textOffset, cch * 2)
      : decodeLatin1(data, textOffset, cch);
    if (text.length === cch && isPrintableText(text)) return text;
  }
  return '';
}

/** 字体名只可能是可打印字符；用来判定"这个偏移读出来的到底是不是名字" */
function isPrintableText(text: string): boolean {
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code < 0x20 || code === 0x7f) return false;
    if (code >= 0xfdd0 && code <= 0xfdef) return false;
    if (code >= 0xfff0 && code <= 0xffff) return false;
  }
  return true;
}

/** 读一条内联字符串（`u16 cch` + `u8 flags` + 字符）：`FORMAT` / `STRING` / `LABEL` 用 */
function readInlineString(bytes: Uint8Array, offset: number): string {
  const cch = readU16(bytes, offset) ?? 0;
  const flags = readU8(bytes, offset + 2) ?? 0;
  return (flags & 0x01) !== 0
    ? decodeUtf16Le(bytes, offset + 3, cch * 2)
    : decodeLatin1(bytes, offset + 3, cch);
}

/* -------------------------------------------------------------------------- */
/* 解析结果类型                                                                */
/* -------------------------------------------------------------------------- */

export interface XlsParseResult {
  input: WorkbookInput;
  /** 中文说明（不是错误，只是"哪些没解析/被跳过"），可直接展示给用户 */
  warnings: string[];
  /** 计数统计：给报告与测试用，便于说清"到底丢了多少东西" */
  stats: Record<string, number>;
}

interface SheetDescriptor {
  name: string;
  /** 0 = 可见、1 = 隐藏、2 = 深度隐藏 */
  visibility: number;
  /** BIFF 流内偏移（该表 BOF 记录头的位置） */
  position: number;
  /** 子流类型（0x0010 = 工作表） */
  substreamType: number;
}

/**
 * 一条 XF 记录（BIFF8 的 `XF` 里内嵌 `CellXF`）。
 *
 * **边框与填充就在 XF 里**（不是靠下标去查另一张"边框表"）：`dgLeft/dgRight/dgTop/dgBottom`
 * 各 4 位是线型枚举，`icv*` 各 7 位是颜色索引，`fls` 6 位是填充图案、`icvFore/icvBack` 是填充色。
 * 曾按"BIFF8 用 BORDER 0x2085 独立记录集"实现，结果整张表都读不出边框与填充 —— 那是 BIFF5/7 的
 * 组织方式，BIFF8 已经内联进 `CellXF` 了。
 */
interface XfRecord {
  fontIndex: number;
  formatIndex: number;
  /** 1 = 单元格样式 XF（不是给单元格用的） */
  isStyleXf: boolean;
  /** 水平对齐（低 3 位） */
  horizontal: number;
  /** 是否自动换行 */
  wrap: boolean;
  /** 垂直对齐（3 位） */
  vertical: number;
  rotation: number;
  indent: number;
  border: { left: number; right: number; top: number; bottom: number };
  borderColor: { left: number; right: number; top: number; bottom: number };
  fillPattern: number;
  fillForeground: number;
  fillBackground: number;
}

interface FontRecord {
  heightTwips: number;
  options: number;
  colorIndex: number;
  boldWeight: number;
  underline: number;
  name: string;
}

interface SheetLayout {
  merges: string[];
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  freeze?: { rows: number; cols: number };
  gridlinesHidden?: boolean;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  /** 只统计用 */
  customRows: number;
  customCols: number;
}

/** 单元格记录先收"最小事实"，样式在最后统一翻译（XF 表要等 globals 读完才完整） */
interface PendingCell {
  row: number;
  col: number;
  value?: string | number | boolean | null;
  xf: number;
  /** 是否来自 BLANK / MULBLANK（只有样式的空单元格） */
  blank: boolean;
}

interface GlobalsInfo {
  dateMode1904: boolean;
  fontTable: FontRecord[];
  formatTable: Map<number, string>;
  xfTable: XfRecord[];
  palette: Map<number, string>;
  sst: string[];
  sheets: SheetDescriptor[];
  encrypted: boolean;
  layouts: SheetLayout[];
}

/* -------------------------------------------------------------------------- */
/* 主入口                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * **主入口：`.xls` 字节 → 中性工作簿。**
 *
 * 永远不抛异常：任何一层失败都退化成"空工作簿 + 中文警告"（上层据此提示用户）。
 * 需要"为什么没解析出来 / 丢了多少"时用 `parseXlsDetailed`，它返回同样的 `input` 加诊断信息。
 */
export function parseXls(bytes: Uint8Array): WorkbookInput {
  return parseXlsDetailed(bytes).input;
}

/**
 * 带诊断的入口：多返回 `warnings`（中文说明）与 `stats`（计数）。
 *
 * 为什么与 `parseXls` 分开：`parseXls` 的签名是给生产链路用的（只关心模型），
 * 而"哪些字段没解析出来、丢了多少"必须能对用户讲清楚，这部分不该塞进 `WorkbookInput`。
 */
export function parseXlsDetailed(bytes: Uint8Array): XlsParseResult {
  const warnings: string[] = [];
  const stats: Record<string, number> = {
    sheets: 0,
    cells: 0,
    styles: 0,
    merges: 0,
    truncatedCells: 0,
  };
  const empty: XlsParseResult = { input: { sheets: [] }, warnings, stats };

  const container = readCfb(bytes, warnings);
  if (!container) return empty;

  const streamEntry =
    container.entries.find((entry) => entry.type === 2 && entry.name === 'Workbook')
    ?? container.entries.find((entry) => entry.type === 2 && entry.name === 'Book');
  if (!streamEntry) {
    const names = container.entries.filter((entry) => entry.type === 2).map((entry) => entry.name).join('、');
    warnings.push(
      `CFB 里找不到 Workbook/Book 流（现有流：${names || '无'}），这不是 Excel 工作簿（可能是 .doc/.ppt 等其它 OLE2 文档）`,
    );
    return empty;
  }

  const stream = readCfbStream(container, streamEntry, bytes, warnings);
  if (stream.length < 8) {
    warnings.push(`Workbook 流只有 ${stream.length} 字节，太短，无法解析`);
    return empty;
  }
  stats.streamBytes = stream.length;

  const records = scanRecords(stream, warnings);
  if (records.length === 0) {
    warnings.push('Workbook 流里没有解析出任何 BIFF 记录');
    return empty;
  }
  stats.records = records.length;

  // ---- 版本判定：BIFF8 = 0x0600；更早的版本用 8 位代码页存字符串，产物会乱码，因此不解析 ----
  const biffVersion = readU16(stream, records[0].dataOffset) ?? 0;
  stats.biffVersion = biffVersion;
  if (biffVersion < 0x0600) {
    warnings.push(
      `这是 BIFF${biffVersion >= 0x0500 ? '5/7' : '4 或更早'} 工作簿（BOF 版本 0x${biffVersion.toString(16)}），`
      + '本模块只支持 BIFF8（Excel 97–2003）：老版本用 8 位代码页存字符串，解析出来会是乱码，所以未导入',
    );
    return empty;
  }
  if (streamEntry.name === 'Book') warnings.push('CFB 流名是 Book（BIFF5 的习惯），但 BOF 版本是 BIFF8，已按 BIFF8 解析');

  // ---- globals 子流 ----
  const globals = parseGlobals(stream, records, warnings, stats);
  if (globals.encrypted) {
    warnings.push(
      '工作簿已加密（FILEPASS 记录）：加密后共享字符串表、样式表与单元格记录全被混淆，无法做部分解析；'
      + '请在 Excel 里取消密码（文件 → 信息 → 保护工作簿 → 用密码进行加密，清空密码）后重新导入',
    );
    return empty;
  }

  // ---- 逐个子流：单元格与版式 ----
  const indexByOffset = new Map<number, number>();
  records.forEach((record, index) => indexByOffset.set(record.offset, index));

  const cellsBySheet: PendingCell[][] = globals.sheets.map(() => []);
  globals.sheets.forEach((sheet, index) => {
    const startIndex = indexByOffset.get(sheet.position);
    if (startIndex === undefined) {
      warnings.push(`工作表 "${sheet.name}" 的 BOUNDSHEET 位置 ${sheet.position} 不是任何记录起点，已按空表处理`);
      return;
    }
    if (sheet.substreamType !== SUBSTREAM.WORKSHEET) {
      const label =
        sheet.substreamType === SUBSTREAM.CHART ? '图表工作表(chartsheet)'
          : sheet.substreamType === SUBSTREAM.MACRO ? '宏工作表(macrosheet)'
            : `类型 0x${sheet.substreamType.toString(16)} 的子流`;
      stats.skippedSubstreams = (stats.skippedSubstreams ?? 0) + 1;
      warnings.push(`工作表 "${sheet.name}" 是${label}，本模块不支持，已跳过（其内容不会出现在预览里）`);
      return;
    }
    cellsBySheet[index] = parseSheetCells(stream, records, startIndex, globals, warnings);
    globals.layouts[index] = parseSheetLayout(stream, records, startIndex, sheet.name, warnings);
  });

  const input = buildWorkbook(globals, cellsBySheet, warnings, stats);
  stats.sheets = input.sheets.length;
  stats.cells = input.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0);
  return { input, warnings, stats };
}

/* -------------------------------------------------------------------------- */
/* globals 子流                                                                */
/* -------------------------------------------------------------------------- */

function parseGlobals(
  stream: Uint8Array,
  records: readonly BiffRecord[],
  warnings: string[],
  stats: Record<string, number>,
): GlobalsInfo {
  const info: GlobalsInfo = {
    dateMode1904: false,
    fontTable: [],
    formatTable: new Map(),
    xfTable: [],
    palette: new Map(),
    sst: [],
    sheets: [],
    encrypted: false,
    layouts: [],
  };

  // globals 到第一个 EOF 为止
  let end = records.length;
  for (let i = 1; i < records.length; i += 1) {
    if (records[i].id === REC.EOF) {
      end = i;
      break;
    }
  }

  for (let i = 1; i < end; i += 1) {
    const record = records[i];
    const data = recordData(stream, record);
    switch (record.id) {
      case REC.FILEPASS:
        info.encrypted = true;
        break;
      case REC.DATEMODE:
        info.dateMode1904 = (readU16(data, 0) ?? 0) === 1;
        break;
      case REC.BOUNDSHEET: {
        const position = readU32(data, 0) ?? 0;
        const visibility = readU8(data, 4) ?? 0;
        // 先记名字与偏移，子流类型等定位到 BOF 再定（BOUNDSHEET 里没有这个信息）
        info.sheets.push({ name: readShortUnicodeString(data, 6), visibility, position, substreamType: -1 });
        break;
      }
      case REC.SST: {
        const payload = readChunked(stream, records, i);
        i += payload.continues;
        info.sst = parseSst(payload, warnings, stats);
        break;
      }
      case REC.FONT:
        info.fontTable.push(parseFont(data));
        break;
      case REC.FORMAT:
        info.formatTable.set(readU16(data, 0) ?? 0, readInlineString(data, 2));
        break;
      case REC.XF:
        info.xfTable.push(parseXf(data));
        break;
      case REC.PALETTE: {
        const count = readU16(data, 0) ?? 0;
        for (let k = 0; k < count; k += 1) {
          const at = 2 + k * 4;
          const red = readU8(data, at);
          const green = readU8(data, at + 1);
          const blue = readU8(data, at + 2);
          if (red === undefined || green === undefined || blue === undefined) break;
          // 调色板从索引 8 开始覆盖
          info.palette.set(k + 8, toHexColor(red, green, blue));
        }
        break;
      }
      default:
        break;
    }
  }

  // 定位各子流的 BOF 并读类型
  const indexByOffset = new Map<number, number>();
  records.forEach((record, index) => indexByOffset.set(record.offset, index));
  for (const sheet of info.sheets) {
    const index = indexByOffset.get(sheet.position);
    if (index === undefined) continue;
    const record = records[index];
    sheet.substreamType = record.id === REC.BOF ? readU16(stream, record.dataOffset + 2) ?? -1 : -1;
  }

  stats.globalsRecords = end;
  stats.fonts = info.fontTable.length;
  stats.xfs = info.xfTable.length;
  stats.formats = info.formatTable.size;
  stats.sstStrings = info.sst.length;
  stats.boundsheets = info.sheets.length;
  stats.dateMode1904 = info.dateMode1904 ? 1 : 0;
  if (info.palette.size > 0) stats.paletteColors = info.palette.size;
  if (info.sheets.length === 0) warnings.push('globals 里没有任何 BOUNDSHEET 记录，工作簿里没有工作表');
  return info;
}

/**
 * 解析 SST（共享字符串表）。
 *
 * 两条规则必须分开处理（最容易写错的地方）：
 *  1. `cstTotal`（含重复的引用数）与 `cstUnique`（去重后的条数）**只在第一段**里，
 *     后续 CONTINUE 直接续字符串数据；**只读 `cstUnique` 条**（按 `cstTotal` 读会越过字符串区）；
 *  2. 字符串跨 CONTINUE 时，新段首字节是编码标志 —— 这条由 `readUnicodeString` 负责。
 */
function parseSst(payload: ChunkedPayload, warnings: string[], stats: Record<string, number>): string[] {
  const out: string[] = [];
  const first = payload.chunks[0];
  if (first.length < 8) {
    warnings.push('SST 记录的第一段不足 8 字节，共享字符串表已忽略');
    return out;
  }
  const total = readU32(first, 0) ?? 0;
  const unique = readU32(first, 4) ?? 0;
  stats.sstTotal = total;
  stats.sstUnique = unique;
  if (total > unique) stats.sstDuplicated = total - unique;

  const wanted = Math.min(total, unique);
  let chunkIndex = 0;
  let position = 8;
  for (let k = 0; k < wanted; k += 1) {
    while (chunkIndex < payload.chunks.length && position >= payload.chunks[chunkIndex].length) {
      chunkIndex += 1;
      position = 0;
    }
    if (chunkIndex >= payload.chunks.length) break;
    const result = readUnicodeString(payload.chunks, chunkIndex, position);
    out.push(result.text);
    chunkIndex = result.chunkIndex;
    position = result.position;
    if (result.truncated) break;
  }
  if (out.length < wanted) {
    warnings.push(`SST 声明 ${wanted} 条共享字符串，实际只读到 ${out.length} 条（记录被截断或长度不自洽），缺失部分已忽略`);
    stats.sstMissing = wanted - out.length;
  }
  return out;
}

function parseFont(data: Uint8Array): FontRecord {
  return {
    heightTwips: readU16(data, 0) ?? 200,
    options: readU16(data, 2) ?? 0,
    colorIndex: readU16(data, 4) ?? 0x7fff,
    boldWeight: readU16(data, 6) ?? 400,
    underline: readU8(data, 12) ?? 0,
    name: readFontName(data),
  };
}

/**
 * 解析 `XF (0x00E0)`。
 *
 * 布局（[MS-XLS] 2.4.353 `XF` + 2.5.20 `CellXF`，共 20 字节）：
 * ```
 *  0..1  ifnt        字体索引
 *  2..3  ifmt        数字格式索引
 *  4     fLocked(1) fHidden(1) fStyle(1) f123Prefix(1) + ixfParent 低 4 位
 *  5     ixfParent 高 8 位
 *  6     alc(3) fWrap(1) alcV(3) fJustLast(1)
 *  7     trot        文字旋转
 *  8     cIndent(4) fShrinkToFit(1) reserved(3)
 *  9..15 边框：dgLeft/dgRight/dgTop/dgBottom 各 4 位 + icvLeft/icvRight/icvTop/icvBottom 各 7 位
 * 16..17 grbitDiag(2) dgDiag(4) fHasXFExt(1) fls(6) fsxButton(1) reserved(1)
 * 18..19 icvFore(7) icvBack(7) reserved(2)
 * ```
 * 注意 `alc`/`alcV` **不是**同一个字节里的两个 4 位字段：`alc` 占 3 位、`fWrap` 占 1 位，
 * 所以 `alcV` 从第 4 位才开始（把 `0x12` 当成"左对齐+居中"就会把居中读成"自动换行关闭"）。
 */
function parseXf(data: Uint8Array): XfRecord {
  const byte4 = readU8(data, 4) ?? 0;
  const byte6 = readU8(data, 6) ?? 0;
  const byte8 = readU8(data, 8) ?? 0;

  /**
   * 从**偏移 9** 起的连续位流读取任意位宽的字段（小端位序）。
   *
   * 为什么必须按位读：边框区 16+7+7+2+7+7 位 + 填充区 6 位一共 54 位，
   * 字段之间**没有字节对齐**（`icvRight` 从第 23 位开始，`fls` 从第 52 位开始）。
   * 按字节/半字节去取会得到"边框线型大致对、颜色全乱"的结果 —— 真踩过。
   */
  const bits = (bitOffset: number, bitCount: number): number => {
    let value = 0;
    for (let i = 0; i < bitCount; i += 1) {
      const absolute = 9 * 8 + bitOffset + i;
      const byte = readU8(data, absolute >> 3) ?? 0;
      value |= ((byte >> (absolute & 7)) & 1) << i;
    }
    return value;
  };

  return {
    fontIndex: readU16(data, 0) ?? 0,
    formatIndex: readU16(data, 2) ?? 0,
    isStyleXf: (byte4 & 0x04) !== 0,
    horizontal: byte6 & 0x07,
    wrap: (byte6 & 0x08) !== 0,
    vertical: (byte6 & 0x70) >> 4,
    rotation: readU8(data, 7) ?? 0,
    indent: byte8 & 0x0f,
    border: {
      left: bits(0, 4),
      right: bits(4, 4),
      top: bits(8, 4),
      bottom: bits(12, 4),
    },
    borderColor: {
      left: bits(16, 7),
      right: bits(23, 7),
      // 第 30..31 位是 grbitDiag（对角边框），不取
      top: bits(32, 7),
      bottom: bits(39, 7),
    },
    fillPattern: bits(52, 6),
    fillForeground: bits(58, 7),
    fillBackground: bits(65, 7),
  };
}

/* -------------------------------------------------------------------------- */
/* 工作表：单元格                                                              */
/* -------------------------------------------------------------------------- */

function parseSheetCells(
  stream: Uint8Array,
  records: readonly BiffRecord[],
  startIndex: number,
  globals: GlobalsInfo,
  warnings: string[],
): PendingCell[] {
  const cells: PendingCell[] = [];
  let truncated = false;
  /** 上一条 FORMULA 在等它的 STRING 缓存结果 */
  let waitingString: PendingCell | undefined;

  for (let i = startIndex + 1; i < records.length; i += 1) {
    const record = records[i];
    if (record.id === REC.EOF) break;
    const data = recordData(stream, record);

    // STRING 只对紧跟在 FORMULA 之后的那一条有效（中间隔了别的记录就作废）
    if (waitingString) {
      if (record.id === REC.STRING) {
        waitingString.value = readInlineString(data, 0);
        waitingString = undefined;
        continue;
      }
      waitingString = undefined;
    }

    if (cells.length >= MAX_XLS_CELLS) {
      if (!truncated) {
        truncated = true;
        warnings.push(`单个工作表的单元格数超过上限（${MAX_XLS_CELLS}），超出部分未导入`);
      }
      break;
    }

    switch (record.id) {
      case REC.LABELSST: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        const index = readU32(data, 6);
        if (row === undefined || col === undefined || xf === undefined || index === undefined) break;
        if (index >= globals.sst.length) {
          warnings.push(`有单元格引用了越界的共享字符串下标 ${index}（SST 只有 ${globals.sst.length} 条），已按空文本处理`);
          pushCell(cells, row, col, xf, '', false);
        } else {
          pushCell(cells, row, col, xf, globals.sst[index], false);
        }
        break;
      }
      case REC.LABEL:
      case REC.RSTRING: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        if (row === undefined || col === undefined || xf === undefined) break;
        // LABEL / RSTRING 的文本都是"u16 cch + u8 flags + 字符"，RSTRING 只是后面多了 rich 段
        pushCell(cells, row, col, xf, readInlineString(data, 6), false);
        break;
      }
      case REC.RK: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        const raw = readU32(data, 6);
        if (row === undefined || col === undefined || xf === undefined || raw === undefined) break;
        pushCell(cells, row, col, xf, decodeRk(raw | 0), false);
        break;
      }
      case REC.MULRK: {
        const row = readU16(data, 0);
        const firstCol = readU16(data, 2);
        if (row === undefined || firstCol === undefined) break;
        const count = Math.floor((record.dataLength - 6) / 6);
        for (let k = 0; k < count; k += 1) {
          const xf = readU16(data, 4 + k * 6);
          const raw = readU32(data, 6 + k * 6);
          if (xf === undefined || raw === undefined) break;
          pushCell(cells, row, firstCol + k, xf, decodeRk(raw | 0), false);
        }
        break;
      }
      case REC.NUMBER: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        const value = readF64(data, 6);
        if (row === undefined || col === undefined || xf === undefined || value === undefined) break;
        pushCell(cells, row, col, xf, value, false);
        break;
      }
      case REC.BOOLERR: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        const value = readU8(data, 6);
        const isError = readU8(data, 7);
        if (row === undefined || col === undefined || xf === undefined || value === undefined) break;
        if (isError) {
          // 错误值（#DIV/0! 等）在 SynthCell 里没有表达形式：如实丢弃 + 只保留样式
          pushCell(cells, row, col, xf, undefined, false);
        } else {
          pushCell(cells, row, col, xf, value !== 0, false);
        }
        break;
      }
      case REC.FORMULA: {
        const row = readU16(data, 0);
        const col = readU16(data, 2);
        const xf = readU16(data, 4);
        if (row === undefined || col === undefined || xf === undefined) break;
        const cell: PendingCell = { row, col, xf, blank: false };
        /**
         * 8 字节"结果字段"（记录内偏移 6..13）有两种编码（BIFF8 规范；实测 `fixture-multi.xls`
         * 的 `第二张!C2`（`IF(...)` 结果是文本 "大"）走的就是特殊值这条）：
         *  - **普通数值**：IEEE754 小端 double；
         *  - **特殊值**：结果字段的**最后两个字节**是 `FF FF`，此时**第一个字节**才是类型 ——
         *    `0` 字符串（真文本在下一条 `STRING` 记录里）、`1` 布尔（取值在第三个字节）、
         *    `2` 错误值（不猜文本，只保留样式）、`3` 空串。
         * 实测踩过的两个坑，都由 `fixture-multi.xls` 当场抓住：
         *  ① 早先按 `grbit` 的位判断结果类型，而真文件里那条字符串结果的 `grbit` 是 `0x0020`，
         *     判断落空 → 把 `FF FF` 当 double 读成了 **NaN**（用户会在格子里看到 "NaN"）；
         *  ② 换成按字节标记判断后，偏移写错了一格（读成记录开头而不是结果字段），
         *     于是"字符串结果"这条又没识别到，值变成空。
         * 现在：只有 `结果字段[6..7] === FF FF` 才算特殊值，类型看 `结果字段[0]`。
         */
        const resultField = 6;
        const special = data[resultField + 6] === 0xff && data[resultField + 7] === 0xff;
        if (special) {
          const kind = data[resultField];
          if (kind === 0) waitingString = cell;
          else if (kind === 1) cell.value = data[resultField + 2] !== 0;
          else if (kind === 3) cell.value = '';
          // kind === 2（错误值）：中性模型没有错误值的表达形式，丢掉值、保留样式
        } else {
          const cached = readF64(data, resultField);
          if (cached === undefined) break;
          cell.value = cached;
        }
        if (row < MAX_ROWS && col < MAX_COLS) cells.push(cell);
        break;
      }
      case REC.BLANK:
      case REC.MULBLANK: {
        if (record.id === REC.BLANK) {
          const row = readU16(data, 0);
          const col = readU16(data, 2);
          const xf = readU16(data, 4);
          if (row === undefined || col === undefined || xf === undefined) break;
          pushCell(cells, row, col, xf, undefined, true);
        } else {
          const row = readU16(data, 0);
          const firstCol = readU16(data, 2);
          if (row === undefined || firstCol === undefined) break;
          const count = Math.floor((record.dataLength - 6) / 2);
          for (let k = 0; k < count; k += 1) {
            const xf = readU16(data, 4 + k * 2);
            if (xf === undefined) break;
            pushCell(cells, row, firstCol + k, xf, undefined, true);
          }
        }
        break;
      }
      default:
        // 行高/列宽/合并/冻结等在 parseSheetLayout 里单独扫；其余记录本模块不使用
        break;
    }
  }

  if (waitingString) {
    // FORMULA 声明了字符串结果却没有后续 STRING 记录：如实丢掉这个缓存值（不编造空串）
    cells.pop();
  }
  return cells;
}

/** 统一的"写一个单元格"入口：集中做行列越界检查，避免每个 case 各写一遍 */
function pushCell(
  cells: PendingCell[],
  row: number,
  col: number,
  xf: number,
  value: string | number | boolean | undefined,
  blank: boolean,
): void {
  if (row >= MAX_ROWS || col >= MAX_COLS) return;
  cells.push(value === undefined ? { row, col, xf, blank } : { row, col, xf, blank, value });
}

/* -------------------------------------------------------------------------- */
/* 工作表：版式（合并/列宽/行高/冻结/网格线）                                     */
/* -------------------------------------------------------------------------- */

function emptyLayout(): SheetLayout {
  return { merges: [], colWidths: {}, rowHeights: {}, customRows: 0, customCols: 0 };
}

function parseSheetLayout(
  stream: Uint8Array,
  records: readonly BiffRecord[],
  startIndex: number,
  sheetName: string,
  warnings: string[],
): SheetLayout {
  const layout = emptyLayout();
  let mergeOverflow = 0;

  for (let i = startIndex + 1; i < records.length; i += 1) {
    const record = records[i];
    if (record.id === REC.EOF) break;
    const data = recordData(stream, record);
    switch (record.id) {
      case REC.MERGEDCELLS: {
        const count = readU16(data, 0) ?? 0;
        for (let k = 0; k < count; k += 1) {
          const firstRow = readU16(data, 2 + k * 8);
          const lastRow = readU16(data, 4 + k * 8);
          const firstCol = readU16(data, 6 + k * 8);
          const lastCol = readU16(data, 8 + k * 8);
          if (firstRow === undefined || lastRow === undefined || firstCol === undefined || lastCol === undefined) break;
          if (
            firstRow >= MAX_ROWS || lastRow >= MAX_ROWS || firstCol >= MAX_COLS || lastCol >= MAX_COLS
            || lastRow < firstRow || lastCol < firstCol
          ) {
            mergeOverflow += 1;
            continue;
          }
          layout.merges.push(`${colLetter(firstCol)}${firstRow + 1}:${colLetter(lastCol)}${lastRow + 1}`);
        }
        break;
      }
      case REC.COLINFO: {
        const firstCol = readU16(data, 0);
        const lastCol = readU16(data, 2);
        const widthRaw = readU16(data, 4);
        const flags = readU16(data, 8) ?? 0;
        if (firstCol === undefined || lastCol === undefined || widthRaw === undefined) break;
        // fUserSet（0x0002）没置位时，宽度只是"默认宽度"的复述，不是用户设定值；fHidden（0x0001）表示隐藏列
        if ((flags & 0x0002) === 0) break;
        const width = Math.round((widthRaw / 256) * 100) / 100;
        if (width <= 0) break;
        const end = Math.min(lastCol, MAX_COLS - 1);
        for (let col = firstCol; col <= end; col += 1) layout.colWidths[col] = width;
        layout.customCols += end - firstCol + 1;
        break;
      }
      case REC.ROW: {
        const row = readU16(data, 0);
        const height = readU16(data, 6);
        const flags = readU16(data, 12);
        if (row === undefined || height === undefined || flags === undefined) break;
        if (row >= MAX_ROWS) break;
        // 只有 fCustomHeight（0x0020）才说明用户设过行高；否则 ht 是默认行高的复述
        if ((flags & 0x0020) !== 0 && height > 0) {
          layout.rowHeights[row] = Math.round((height / 20) * 100) / 100;
          layout.customRows += 1;
        }
        break;
      }
      case REC.DEFAULTROWHEIGHT: {
        const height = readU16(data, 2);
        if (height !== undefined && height > 0) layout.defaultRowHeight = Math.round((height / 20) * 100) / 100;
        break;
      }
      case REC.DEFCOLWIDTH: {
        const chars = readU16(data, 0);
        if (chars !== undefined && chars > 0) layout.defaultColWidth = chars;
        break;
      }
      case REC.PANE: {
        const splitX = readU16(data, 0) ?? 0;
        const splitY = readU16(data, 2) ?? 0;
        const topRow = readU16(data, 4) ?? 0;
        const leftCol = readU16(data, 6) ?? 0;
        // 只有真的"冻结"（split 值 > 0）才算；0 表示没有窗格分割
        const rows = splitY > 0 ? topRow : 0;
        const cols = splitX > 0 ? leftCol : 0;
        if (rows > 0 || cols > 0) layout.freeze = { rows, cols };
        break;
      }
      case REC.WINDOW2: {
        const flags = readU16(data, 0) ?? 0;
        // bit1（0x0002）= fGridLines：清零表示隐藏网格线
        if ((flags & 0x0002) === 0) layout.gridlinesHidden = true;
        break;
      }
      default:
        break;
    }
  }

  if (mergeOverflow > 0) warnings.push(`工作表 "${sheetName}" 有 ${mergeOverflow} 个合并区越界或非法，已忽略`);
  return layout;
}

/* -------------------------------------------------------------------------- */
/* 样式翻译                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 把 XF / FONT / FORMAT / PALETTE / BORDER 翻译成"XF 下标 → 中性样式"的表。
 *
 * 为什么做成表而不是逐格翻译：`SynthCell.style` 是**下标**语义，同一张表里成百上千个格子
 * 往往只有十几种 XF；先去重能让下游 writer 的 font/fill/border intern 命中同一批对象。
 */
function buildXfStyles(
  globals: GlobalsInfo,
  warnings: string[],
  stats: Record<string, number>,
): Array<SynthStyle | undefined> {
  const table: Array<SynthStyle | undefined> = [];
  let fontMissing = 0;

  for (const xf of globals.xfTable) {
    // 样式 XF（fStyle = 1）只用来描述"单元格样式"，不直接作用于单元格；单元格引用它们时
    // 继承属性，但真夹具里单元格直接引用带完整属性的 XF，所以这里如实翻译每一条即可。
    const style: SynthStyle = {};
    const font = globals.fontTable[xf.fontIndex];
    if (font) {
      if (font.name) style.fontFamily = font.name;
      if (font.heightTwips > 0) style.fontSize = Math.round((font.heightTwips / 20) * 2) / 2;
      if (font.boldWeight >= 600) style.bold = true;
      if ((font.options & 0x0002) !== 0) style.italic = true;
      if ((font.options & 0x0008) !== 0) style.strikeThrough = true;
      // 下划线枚举：1 = 单线、2 = 双线、0x21/0x22 = 会计用单线/双线（中性模型都归为"有下划线"）
      const underline = font.underline;
      if (underline === 1 || underline === 2 || underline === 0x21 || underline === 0x22) style.underline = true;
      const color = colorFor(globals, font.colorIndex);
      if (color) style.color = color;
    } else if (xf.fontIndex !== 0) {
      fontMissing += 1;
    }

    const horizontal = xf.horizontal;
    if (horizontal === 1) style.horizontalAlign = 'left';
    else if (horizontal === 2) style.horizontalAlign = 'center';
    else if (horizontal === 3) style.horizontalAlign = 'right';
    else if (horizontal >= 4) {
      // 4=fill 5=justify 6=centerAcross 7=distributed：中性模型只有三档，按最接近的落位并计数
      style.horizontalAlign = horizontal === 6 ? 'center' : horizontal === 4 ? 'left' : 'center';
      stats.approximatedAlignments = (stats.approximatedAlignments ?? 0) + 1;
    }
    const vertical = xf.vertical;
    if (vertical === 0) style.verticalAlign = 'top';
    else if (vertical === 1) style.verticalAlign = 'middle';
    else if (vertical === 2) style.verticalAlign = 'bottom';
    else if (vertical === 3) {
      style.verticalAlign = 'middle'; // vjustify 按居中近似
      stats.approximatedAlignments = (stats.approximatedAlignments ?? 0) + 1;
    }
    if (xf.wrap) style.textWrap = true;
    if (xf.indent > 0) stats.xfIndent = (stats.xfIndent ?? 0) + 1;

    // 文字旋转：1..90 顺时针、91..180 表示逆时针 (90 - 值)、255 是竖排
    const rotation = xf.rotation;
    if (rotation >= 1 && rotation <= 90) style.textRotation = rotation;
    else if (rotation > 90 && rotation < 180) style.textRotation = 90 - rotation;
    else if (rotation === 255) stats.verticalText = (stats.verticalText ?? 0) + 1;

    // 填充：fls = 1 是纯色（只渲染 icvFore）；2..18 是条纹/网格图案
    const pattern = xf.fillPattern;
    if (pattern === 1) {
      const color = colorFor(globals, xf.fillForeground);
      if (color) style.fill = color;
    } else if (pattern !== 0) {
      // 图案填充在中性模型里只有"纯色填充"一种表达，如实丢弃并计数
      stats.patternFillsDropped = (stats.patternFillsDropped ?? 0) + 1;
    }

    const side = (line: number, colorIndex: number): { style: string; color?: string } | undefined => {
      const lineStyle = BORDER_STYLES[line];
      if (!lineStyle) return undefined;
      const color = colorFor(globals, colorIndex);
      return color ? { style: lineStyle, color } : { style: lineStyle };
    };
    const top = side(xf.border.top, xf.borderColor.top);
    const bottom = side(xf.border.bottom, xf.borderColor.bottom);
    const left = side(xf.border.left, xf.borderColor.left);
    const right = side(xf.border.right, xf.borderColor.right);
    if (top || bottom || left || right) {
      style.border = {
        ...(top ? { top } : {}),
        ...(bottom ? { bottom } : {}),
        ...(left ? { left } : {}),
        ...(right ? { right } : {}),
      };
    }

    const format = formatCodeFor(globals, xf.formatIndex);
    if (format && format !== 'General') style.numberFormat = format;

    table.push(Object.keys(style).length > 0 ? style : undefined);
  }

  stats.xfStyles = table.length;
  if (fontMissing > 0) warnings.push(`有 ${fontMissing} 个单元格样式引用了查不到的字体记录，这些样式的字体留空`);
  return table;
}

/**
 * BIFF8 边框线型（边框项的低 4 位）→ 中性线型名。
 * 该枚举与 OOXML 的 `ST_BorderStyle` 同构，因此这一步翻译是无损的。
 */
const BORDER_STYLES: Readonly<Record<number, string | undefined>> = {
  0: undefined, // 无线
  1: 'thin',
  2: 'medium',
  3: 'dashed',
  4: 'dotted',
  5: 'thick',
  6: 'double',
  7: 'hair',
  8: 'mediumDashed',
  9: 'dashDot',
  10: 'mediumDashDot',
  11: 'dashDotDot',
  12: 'mediumDashDotDot',
  13: 'slantDashDot',
};

/** 颜色索引 → `#RRGGBB`；`0x7FFF`（自动）返回 undefined（不猜颜色），`0x40` 是系统前景色（黑） */
function colorFor(globals: GlobalsInfo, index: number): string | undefined {
  if (index === 0x7fff) return undefined;
  if (index === 0x40) return '#000000';
  return globals.palette.get(index) ?? INDEXED_COLORS[index];
}

/** 数字格式索引 → 格式串：文件里的 FORMAT 记录优先，其次查内置表（与 OOXML 的内置表同一份） */
function formatCodeFor(globals: GlobalsInfo, index: number): string | undefined {
  return globals.formatTable.get(index) ?? BUILTIN_NUM_FMTS[index];
}

function toHexColor(red: number, green: number, blue: number): string {
  const hex = (value: number): string => Math.max(0, Math.min(255, value)).toString(16).padStart(2, '0').toUpperCase();
  return `#${hex(red)}${hex(green)}${hex(blue)}`;
}

function colLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/* -------------------------------------------------------------------------- */
/* 组装中性工作簿                                                              */
/* -------------------------------------------------------------------------- */

function buildWorkbook(
  globals: GlobalsInfo,
  cellsBySheet: readonly PendingCell[][],
  warnings: string[],
  stats: Record<string, number>,
): WorkbookInput {
  const xfStyles = buildXfStyles(globals, warnings, stats);
  const styleIds = new Map<string, number>();
  const styles: SynthStyle[] = [];
  let blankDropped = 0;
  let outOfRangeXf = 0;

  const sheets: SynthSheet[] = globals.sheets.map((descriptor, index) => {
    const layout = globals.layouts[index] ?? emptyLayout();
    const cells: SynthCell[] = [];

    for (const item of cellsBySheet[index] ?? []) {
      let styleId: number | undefined;
      if (item.xf >= xfStyles.length) {
        outOfRangeXf += 1;
      } else {
        const xfStyle = xfStyles[item.xf];
        if (xfStyle) {
          const key = JSON.stringify(xfStyle);
          const existing = styleIds.get(key);
          if (existing !== undefined) styleId = existing;
          else {
            styleId = styles.length;
            styles.push(xfStyle);
            styleIds.set(key, styleId);
          }
        }
      }
      if (item.blank && styleId === undefined) {
        // "只有样式的空单元格"在样式就是默认样式时不含任何信息，不写（否则整表到处是空格子）
        blankDropped += 1;
        continue;
      }
      const cell: SynthCell = { row: item.row, col: item.col };
      if (item.value !== undefined) cell.value = item.value;
      if (styleId !== undefined) cell.style = styleId;
      cells.push(cell);
    }

    const sheet: SynthSheet = { name: descriptor.name, cells };
    if (descriptor.visibility === 1 || descriptor.visibility === 2) sheet.hidden = true;
    if (layout.merges.length > 0) sheet.merges = layout.merges;
    if (Object.keys(layout.colWidths).length > 0) sheet.colWidths = layout.colWidths;
    if (Object.keys(layout.rowHeights).length > 0) sheet.rowHeights = layout.rowHeights;
    if (layout.freeze && (layout.freeze.rows > 0 || layout.freeze.cols > 0)) sheet.freeze = layout.freeze;
    if (layout.gridlinesHidden) sheet.gridlinesHidden = true;
    if (layout.defaultRowHeight !== undefined) sheet.defaultRowHeight = layout.defaultRowHeight;
    if (layout.defaultColWidth !== undefined) sheet.defaultColWidth = layout.defaultColWidth;
    stats.merges = (stats.merges ?? 0) + layout.merges.length;
    stats.customRows = (stats.customRows ?? 0) + layout.customRows;
    stats.customCols = (stats.customCols ?? 0) + layout.customCols;
    if (descriptor.visibility === 2) stats.veryHiddenSheets = (stats.veryHiddenSheets ?? 0) + 1;
    return sheet;
  });

  stats.styles = styles.length;
  if (blankDropped > 0) stats.blankCellsDropped = blankDropped;
  if (outOfRangeXf > 0) warnings.push(`有 ${outOfRangeXf} 个单元格引用了越界的样式索引（XF），已按默认样式处理`);

  return styles.length > 0 ? { sheets, styles } : { sheets };
}
