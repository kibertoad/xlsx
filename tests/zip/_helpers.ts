// Shared zip fixture helpers. Malformed-archive tests build a real archive
// through the project's own writer and then bend one header field, so the
// fixture differs from a valid one in exactly the way under test. The record
// offsets live here so a new test does not re-derive the layout from APPNOTE.

import { toBuffer } from '../../src/io/node.js';
import { createZipWriter } from '../../src/zip/writer.js';

const SIG_LFH = 0x04034b50;
const SIG_CD = 0x02014b50;
const SIG_EOCD = 0x06054b50;

// Offsets within each record, relative to its signature (APPNOTE 4.3.7,
// 4.3.12, 4.3.16).
const LFH_COMP_SIZE = 18;
const LFH_UNCOMP_SIZE = 22;
const CD_COMP_SIZE = 20;
const CD_UNCOMP_SIZE = 24;
const CD_NAME_LEN = 28;
const CD_EXTRA_LEN = 30;
const CD_COMMENT_LEN = 32;
const CD_LFH_OFFSET = 42;
const CD_NAME = 46;
const CD_RECORD_SIZE = 46;
const EOCD_TOTAL_ENTRIES = 10;
const EOCD_CD_OFFSET = 16;
const EOCD_SIZE = 22;
const MAX_EOCD_COMMENT = 0xffff;

const u16 = (b: Uint8Array, off: number): number => (b[off] ?? 0) | ((b[off + 1] ?? 0) << 8);

const u32 = (b: Uint8Array, off: number): number =>
  ((b[off] ?? 0) | ((b[off + 1] ?? 0) << 8) | ((b[off + 2] ?? 0) << 16) | ((b[off + 3] ?? 0) << 24)) >>> 0;

const writeU32 = (b: Uint8Array, off: number, v: number): void => {
  b[off] = v & 0xff;
  b[off + 1] = (v >>> 8) & 0xff;
  b[off + 2] = (v >>> 16) & 0xff;
  b[off + 3] = (v >>> 24) & 0xff;
};

/** Build a zip archive from in-memory entries. `compress: false` stores an entry. */
export const buildArchive = async (
  entries: ReadonlyArray<{ path: string; bytes: Uint8Array; compress?: boolean }>,
): Promise<Uint8Array> => {
  const sink = toBuffer();
  const w = createZipWriter(sink);
  for (const e of entries) {
    await w.addEntry(e.path, e.bytes, e.compress === false ? { compress: false } : undefined);
  }
  await w.finalize();
  return sink.result();
};

const findEocd = (b: Uint8Array): number => {
  const minStart = Math.max(0, b.length - EOCD_SIZE - MAX_EOCD_COMMENT);
  for (let i = b.length - EOCD_SIZE; i >= minStart; i--) {
    if (u32(b, i) === SIG_EOCD) return i;
  }
  throw new Error('no EOCD signature in fixture archive');
};

/** Byte offset of the central-directory record for `name`. */
const findCdRecord = (b: Uint8Array, name: string): number => {
  const eocd = findEocd(b);
  const total = u16(b, eocd + EOCD_TOTAL_ENTRIES);
  let p = u32(b, eocd + EOCD_CD_OFFSET);
  const decoder = new TextDecoder();
  for (let i = 0; i < total; i++) {
    if (u32(b, p) !== SIG_CD) throw new Error(`no central-directory record at byte ${p}`);
    const nameLen = u16(b, p + CD_NAME_LEN);
    if (decoder.decode(b.subarray(p + CD_NAME, p + CD_NAME + nameLen)) === name) return p;
    p += CD_RECORD_SIZE + nameLen + u16(b, p + CD_EXTRA_LEN) + u16(b, p + CD_COMMENT_LEN);
  }
  throw new Error(`no entry named "${name}" in fixture archive`);
};

/**
 * Rewrite an entry's declared sizes in both the central directory and the
 * local file header, returning a copy. The reader trusts the central
 * directory; patching the local header too keeps the fixture consistent with
 * what a cross-checking reader would see.
 */
export const patchEntrySizes = (
  bytes: Uint8Array,
  name: string,
  sizes: { compSize?: number; uncompSize?: number },
): Uint8Array => {
  const out = new Uint8Array(bytes);
  const cd = findCdRecord(out, name);
  const lfh = u32(out, cd + CD_LFH_OFFSET);
  if (u32(out, lfh) !== SIG_LFH) throw new Error(`no local file header for "${name}" at byte ${lfh}`);
  if (sizes.compSize !== undefined) {
    writeU32(out, cd + CD_COMP_SIZE, sizes.compSize);
    writeU32(out, lfh + LFH_COMP_SIZE, sizes.compSize);
  }
  if (sizes.uncompSize !== undefined) {
    writeU32(out, cd + CD_UNCOMP_SIZE, sizes.uncompSize);
    writeU32(out, lfh + LFH_UNCOMP_SIZE, sizes.uncompSize);
  }
  return out;
};
