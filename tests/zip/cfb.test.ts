import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlNotImplementedError } from '../../src/utils/exceptions.js';
import { classifyCfb } from '../../src/zip/cfb.js';
import { openZip } from '../../src/zip/reader.js';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const FREE_SECTOR = 0xffffffff;
const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;
const TYPE_ROOT = 5;

interface Entry {
  name: string;
  type: number;
}

interface CfbLayout {
  entries: Entry[];
  /** 9 for a version 3 file (512-byte sectors), 12 for version 4 (4096-byte). */
  sectorShift?: number;
  /** Overrides the FAT slot of the last directory sector, to build a bad chain. */
  lastDirectorySectorNext?: number;
}

const stream = (name: string): Entry => ({ name, type: TYPE_STREAM });

// Lays the file out as: header, sector 0 = FAT, then the directory on the odd
// sectors 1, 3, 5, ... with the even ones left free. The gaps are the point: a
// reader that assumes the directory is contiguous instead of following the FAT
// reads a free sector and misses every entry after the first four.
const buildCfb = ({ entries, sectorShift = 9, lastDirectorySectorNext = END_OF_CHAIN }: CfbLayout): Uint8Array => {
  const sectorSize = 2 ** sectorShift;
  const entriesPerSector = sectorSize / 128;
  const directorySectors = Math.max(1, Math.ceil(entries.length / entriesPerSector));
  const lastSector = 2 * directorySectors - 1;
  const bytes = new Uint8Array((lastSector + 2) * sectorSize);
  const view = new DataView(bytes.buffer);
  const sectorOffset = (sector: number): number => (sector + 1) * sectorSize;

  bytes.set(CFB_MAGIC, 0);
  view.setUint16(0x1a, sectorShift === 12 ? 4 : 3, true);
  view.setUint16(0x1e, sectorShift, true);
  view.setUint32(0x30, 1, true);
  for (let i = 0; i < 109; i++) view.setUint32(0x4c + i * 4, i === 0 ? 0 : FREE_SECTOR, true);

  const fat = sectorOffset(0);
  for (let i = 0; i < sectorSize / 4; i++) view.setUint32(fat + i * 4, FREE_SECTOR, true);
  view.setUint32(fat, FAT_SECTOR, true);
  for (let i = 0; i < directorySectors; i++) {
    const sector = 2 * i + 1;
    view.setUint32(fat + sector * 4, sector === lastSector ? lastDirectorySectorNext : sector + 2, true);
  }

  entries.forEach((entry, index) => {
    const sector = 2 * Math.floor(index / entriesPerSector) + 1;
    const at = sectorOffset(sector) + (index % entriesPerSector) * 128;
    for (let i = 0; i < entry.name.length; i++) view.setUint16(at + i * 2, entry.name.charCodeAt(i), true);
    view.setUint16(at + 0x40, (entry.name.length + 1) * 2, true);
    view.setUint8(at + 0x42, entry.type);
  });
  return bytes;
};

const ROOT: Entry = { name: 'Root Entry', type: TYPE_ROOT };
const PROPERTY_STREAMS = [stream('\u0005SummaryInformation'), stream('\u0005DocumentSummaryInformation')];

const biff8 = (): Uint8Array => buildCfb({ entries: [ROOT, stream('Workbook'), ...PROPERTY_STREAMS] });

// The stream layout Excel writes for a password-protected xlsx. With 512-byte
// sectors `EncryptedPackage` lands in the third directory sector.
const encryptedXlsx = (sectorShift?: number): Uint8Array =>
  buildCfb({
    entries: [
      ROOT,
      { name: '\u0006DataSpaces', type: TYPE_STORAGE },
      stream('Version'),
      stream('DataSpaceMap'),
      { name: 'DataSpaceInfo', type: TYPE_STORAGE },
      stream('StrongEncryptionDataSpace'),
      { name: 'TransformInfo', type: TYPE_STORAGE },
      stream('\u0006Primary'),
      stream('EncryptionInfo'),
      stream('EncryptedPackage'),
    ],
    ...(sectorShift === undefined ? {} : { sectorShift }),
  });

describe('classifyCfb', () => {
  it('reads a BIFF8 Workbook stream as a legacy workbook', () => {
    expect(classifyCfb(biff8())).toBe('legacy-workbook');
  });

  it('reads a BIFF5 Book stream as a legacy workbook', () => {
    expect(classifyCfb(buildCfb({ entries: [ROOT, stream('Book')] }))).toBe('legacy-workbook');
  });

  it('follows the FAT chain to an EncryptedPackage stream in a later directory sector', () => {
    expect(classifyCfb(encryptedXlsx())).toBe('encrypted-ooxml');
  });

  it('handles the 4096-byte sectors of a version 4 file', () => {
    expect(classifyCfb(encryptedXlsx(12))).toBe('encrypted-ooxml');
  });

  it('matches stream names case-insensitively, as CFB compares them', () => {
    expect(classifyCfb(buildCfb({ entries: [ROOT, stream('WORKBOOK')] }))).toBe('legacy-workbook');
  });

  it('prefers EncryptedPackage when a Workbook stream is present too', () => {
    const both = buildCfb({ entries: [ROOT, stream('Workbook'), stream('EncryptedPackage')] });
    expect(classifyCfb(both)).toBe('encrypted-ooxml');
  });

  it('ignores a storage that happens to be called Workbook', () => {
    const storage = buildCfb({ entries: [ROOT, { name: 'Workbook', type: TYPE_STORAGE }] });
    expect(classifyCfb(storage)).toBe('unknown');
  });

  it('is unknown for another Office format', () => {
    expect(classifyCfb(buildCfb({ entries: [ROOT, stream('WordDocument'), stream('1Table')] }))).toBe('unknown');
  });

  it('is unknown when only the magic bytes are present', () => {
    expect(classifyCfb(new Uint8Array(CFB_MAGIC))).toBe('unknown');
    const headerOnly = new Uint8Array(512);
    headerOnly.set(CFB_MAGIC, 0);
    expect(classifyCfb(headerOnly)).toBe('unknown');
  });

  it('is unknown when the directory sector lies past the end of the file', () => {
    expect(classifyCfb(biff8().subarray(0, 1024))).toBe('unknown');
  });

  it('terminates on a directory chain that loops back on itself', () => {
    const looping = buildCfb({ entries: [ROOT, stream('WordDocument')], lastDirectorySectorNext: 1 });
    expect(classifyCfb(looping)).toBe('unknown');
  });

  it('reads a view into a larger buffer from its own offset', () => {
    const file = biff8();
    const padded = new Uint8Array(file.length + 64);
    padded.set(file, 64);
    expect(classifyCfb(padded.subarray(64))).toBe('legacy-workbook');
  });
});

describe('CFB input is rejected with advice that fits the file', () => {
  it('tells the owner of a legacy .xls to convert it', async () => {
    const attempt = openZip(fromBuffer(biff8()));
    await expect(attempt).rejects.toThrowError(OpenXmlNotImplementedError);
    await expect(attempt).rejects.toThrowError(/legacy \.xls .* Convert it to \.xlsx/);
    await expect(attempt).rejects.not.toThrowError(/decrypt/i);
  });

  it('tells the owner of an encrypted xlsx to decrypt it', async () => {
    const attempt = openZip(fromBuffer(encryptedXlsx()));
    await expect(attempt).rejects.toThrowError(OpenXmlNotImplementedError);
    await expect(attempt).rejects.toThrowError(/Encrypted xlsx is not supported/);
  });

  it('names both possibilities when the container cannot be classified', async () => {
    const headerOnly = new Uint8Array(512);
    headerOnly.set(CFB_MAGIC, 0);
    await expect(openZip(fromBuffer(headerOnly))).rejects.toThrowError(/password-protected xlsx .* legacy Office format/);
  });

  it('reaches loadWorkbook and loadWorkbookStream callers', async () => {
    await expect(loadWorkbook(fromBuffer(biff8()))).rejects.toThrowError(/legacy \.xls/);
    await expect(loadWorkbookStream(fromBuffer(biff8()))).rejects.toThrowError(/legacy \.xls/);
  });
});
