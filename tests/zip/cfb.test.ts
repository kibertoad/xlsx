import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlNotImplementedError, OpenXmlUnsupportedFormatError } from '../../src/utils/exceptions.js';
import { classifyCfb } from '../../src/zip/cfb.js';
import { openZip } from '../../src/zip/reader.js';

// Written by LibreOffice 25.8 from a two-cell sheet: `biff8.xls` through the
// "MS Excel 97" filter, `biff8-password.xls` the same with the password
// `secret`, and `encrypted.xlsx` through "Calc MS Excel 2007 XML" with the same
// password. They anchor the synthetic containers below to a real writer's
// layout, which `buildCfb` cannot do because it shares the reader's reading of
// the spec.
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/cfb/${name}`, import.meta.url)));

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const FREE_SECTOR = 0xffffffff;
const NO_STREAM = 0xffffffff;
const HEADER_DIFAT_ENTRIES = 109;
const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;
const TYPE_ROOT = 5;

interface Entry {
  name: string;
  type: number;
  children?: Entry[];
}

interface CfbLayout {
  /** What the root storage holds. The root entry itself is added by the builder. */
  children: Entry[];
  /** 9 for a version 3 file (512-byte sectors), 12 for version 4 (4096-byte). */
  sectorShift?: number;
  /** Where the directory chain starts. Defaults to the sector after the FAT. */
  firstDirectorySector?: number;
  /** Overrides the FAT slot of the last directory sector, to build a bad chain. */
  lastDirectorySectorNext?: number;
  /** Overrides the right-sibling id of the root's last child, to build a bad tree. */
  lastChildRightSibling?: number;
}

const stream = (name: string): Entry => ({ name, type: TYPE_STREAM });
const storage = (name: string, children: Entry[]): Entry => ({ name, type: TYPE_STORAGE, children });

interface FlatEntry {
  name: string;
  type: number;
  child: number;
  right: number;
}

// Depth-first ids, with each storage's children strung together as a chain of
// right siblings. A real writer balances the tree, which the fixtures cover.
const flatten = (rootChildren: Entry[], lastChildRightSibling: number): FlatEntry[] => {
  const root: FlatEntry = { name: 'Root Entry', type: TYPE_ROOT, child: NO_STREAM, right: NO_STREAM };
  const flat = [root];
  const add = (children: Entry[], parent: FlatEntry, lastRight: number): void => {
    let previous: FlatEntry | undefined;
    for (const child of children) {
      const entry: FlatEntry = { name: child.name, type: child.type, child: NO_STREAM, right: lastRight };
      if (previous === undefined) parent.child = flat.length;
      else previous.right = flat.length;
      flat.push(entry);
      add(child.children ?? [], entry, NO_STREAM);
      previous = entry;
    }
  };
  add(rootChildren, root, lastChildRightSibling);
  return flat;
};

// Lays the file out as: header, the FAT sectors, the DIFAT sectors if the FAT
// outgrows the header's 109 slots, then the directory on every other sector
// from `firstDirectorySector` with the ones between left free. The gaps are the
// point: a reader that assumes the directory is contiguous instead of following
// the FAT reads a free sector and misses every entry after the first four.
const buildCfb = ({
  children,
  sectorShift = 9,
  firstDirectorySector = 1,
  lastDirectorySectorNext = END_OF_CHAIN,
  lastChildRightSibling = NO_STREAM,
}: CfbLayout): Uint8Array => {
  const entries = flatten(children, lastChildRightSibling);
  const sectorSize = 2 ** sectorShift;
  const idsPerSector = sectorSize / 4;
  const entriesPerSector = sectorSize / 128;
  const directorySectors = Math.ceil(entries.length / entriesPerSector);
  const lastSector = firstDirectorySector + 2 * (directorySectors - 1);
  const fatSectors = Math.ceil((lastSector + 1) / idsPerSector);
  const difatSectors = Math.ceil(Math.max(0, fatSectors - HEADER_DIFAT_ENTRIES) / (idsPerSector - 1));
  const bytes = new Uint8Array((lastSector + 2) * sectorSize);
  const view = new DataView(bytes.buffer);
  const sectorOffset = (sector: number): number => (sector + 1) * sectorSize;
  const setFat = (sector: number, next: number): void =>
    view.setUint32(sectorOffset(Math.floor(sector / idsPerSector)) + (sector % idsPerSector) * 4, next, true);

  bytes.set(CFB_MAGIC, 0);
  view.setUint16(0x1a, sectorShift === 12 ? 4 : 3, true);
  view.setUint16(0x1e, sectorShift, true);
  view.setUint32(0x2c, fatSectors, true);
  view.setUint32(0x30, firstDirectorySector, true);
  view.setUint32(0x44, difatSectors === 0 ? END_OF_CHAIN : fatSectors, true);
  view.setUint32(0x48, difatSectors, true);

  // FAT sectors are 0..fatSectors-1 and DIFAT sectors follow them directly.
  for (let sector = 0; sector < fatSectors * idsPerSector; sector++) setFat(sector, FREE_SECTOR);
  for (let sector = 0; sector < fatSectors; sector++) setFat(sector, FAT_SECTOR);
  for (let i = 0; i < difatSectors; i++) setFat(fatSectors + i, DIFAT_SECTOR);
  for (let sector = firstDirectorySector; sector <= lastSector; sector += 2) {
    setFat(sector, sector === lastSector ? lastDirectorySectorNext : sector + 2);
  }

  for (let i = 0; i < HEADER_DIFAT_ENTRIES; i++) {
    view.setUint32(0x4c + i * 4, i < fatSectors ? i : FREE_SECTOR, true);
  }
  for (let i = 0; i < difatSectors; i++) {
    const at = sectorOffset(fatSectors + i);
    for (let slot = 0; slot < idsPerSector - 1; slot++) {
      const fatSector = HEADER_DIFAT_ENTRIES + i * (idsPerSector - 1) + slot;
      view.setUint32(at + slot * 4, fatSector < fatSectors ? fatSector : FREE_SECTOR, true);
    }
    view.setUint32(at + (idsPerSector - 1) * 4, i === difatSectors - 1 ? END_OF_CHAIN : fatSectors + i + 1, true);
  }

  entries.forEach((entry, index) => {
    const sector = firstDirectorySector + 2 * Math.floor(index / entriesPerSector);
    const at = sectorOffset(sector) + (index % entriesPerSector) * 128;
    for (let i = 0; i < entry.name.length; i++) view.setUint16(at + i * 2, entry.name.charCodeAt(i), true);
    view.setUint16(at + 0x40, (entry.name.length + 1) * 2, true);
    view.setUint8(at + 0x42, entry.type);
    view.setUint32(at + 0x44, NO_STREAM, true);
    view.setUint32(at + 0x48, entry.right, true);
    view.setUint32(at + 0x4c, entry.child, true);
  });
  return bytes;
};

const PROPERTY_STREAMS = [stream('\u0005SummaryInformation'), stream('\u0005DocumentSummaryInformation')];

const biff8 = (): Uint8Array => buildCfb({ children: [stream('Workbook'), ...PROPERTY_STREAMS] });

// The stream layout Excel writes for a password-protected xlsx. With 512-byte
// sectors `EncryptedPackage` lands in the third directory sector.
const ENCRYPTED_XLSX_CHILDREN = [
  storage('\u0006DataSpaces', [
    stream('Version'),
    stream('DataSpaceMap'),
    storage('DataSpaceInfo', [stream('StrongEncryptionDataSpace')]),
    storage('TransformInfo', [storage('StrongEncryptionTransform', [stream('\u0006Primary')])]),
  ]),
  stream('EncryptionInfo'),
  stream('EncryptedPackage'),
];

describe('classifyCfb on files written by LibreOffice', () => {
  it('reads a BIFF8 .xls as a legacy workbook', () => {
    expect(classifyCfb(fixture('biff8.xls'))).toBe('legacy-xls');
  });

  it('reads a password-protected xlsx as encrypted', () => {
    expect(classifyCfb(fixture('encrypted.xlsx'))).toBe('encrypted-xlsx');
  });

  // BIFF encryption is a FILEPASS record inside the Workbook stream, so the
  // container looks like any other .xls.
  it('reads a password-protected .xls as a legacy workbook', () => {
    expect(classifyCfb(fixture('biff8-password.xls'))).toBe('legacy-xls');
  });
});

describe('classifyCfb', () => {
  it('reads a BIFF8 Workbook stream as a legacy workbook', () => {
    expect(classifyCfb(biff8())).toBe('legacy-xls');
  });

  it('reads a BIFF5 Book stream as a legacy workbook', () => {
    expect(classifyCfb(buildCfb({ children: [stream('Book')] }))).toBe('legacy-xls');
  });

  it('follows the FAT chain to an EncryptedPackage stream in a later directory sector', () => {
    expect(classifyCfb(buildCfb({ children: ENCRYPTED_XLSX_CHILDREN }))).toBe('encrypted-xlsx');
  });

  it('handles the 4096-byte sectors of a version 4 file', () => {
    expect(classifyCfb(buildCfb({ children: ENCRYPTED_XLSX_CHILDREN, sectorShift: 12 }))).toBe('encrypted-xlsx');
  });

  // Sector 14080 is the first one whose FAT slot lives in FAT sector 110, past
  // the 109 the header lists, so reaching the second directory sector means
  // reading the DIFAT chain.
  it('follows the DIFAT chain when the directory sits past the FAT sectors the header lists', () => {
    const large = buildCfb({ children: ENCRYPTED_XLSX_CHILDREN, firstDirectorySector: 110 * 128 });
    expect(classifyCfb(large)).toBe('encrypted-xlsx');
  });

  it('matches stream names case-insensitively, as CFB compares them', () => {
    expect(classifyCfb(buildCfb({ children: [stream('WORKBOOK')] }))).toBe('legacy-xls');
  });

  it('prefers EncryptedPackage when the root storage holds a Workbook stream too', () => {
    const both = buildCfb({ children: [stream('Workbook'), stream('EncryptedPackage')] });
    expect(classifyCfb(both)).toBe('encrypted-xlsx');
  });

  it('ignores the EncryptedPackage of a protected document embedded in a .xls', () => {
    const embedded = storage('MBD0001A2B3', [stream('EncryptionInfo'), stream('EncryptedPackage')]);
    expect(classifyCfb(buildCfb({ children: [embedded, stream('Workbook')] }))).toBe('legacy-xls');
  });

  it('ignores the Workbook stream of a sheet embedded in a .doc', () => {
    const embedded = storage('ObjectPool', [storage('_1234567890', [stream('Workbook')])]);
    expect(classifyCfb(buildCfb({ children: [stream('WordDocument'), embedded] }))).toBe('compound-file');
  });

  it('ignores a storage that happens to be called Workbook', () => {
    expect(classifyCfb(buildCfb({ children: [storage('Workbook', [])] }))).toBe('compound-file');
  });

  it('is compound-file for another Office format', () => {
    expect(classifyCfb(buildCfb({ children: [stream('WordDocument'), stream('1Table')] }))).toBe('compound-file');
  });

  it('is compound-file when only the magic bytes are present', () => {
    expect(classifyCfb(new Uint8Array(CFB_MAGIC))).toBe('compound-file');
    const headerOnly = new Uint8Array(512);
    headerOnly.set(CFB_MAGIC, 0);
    expect(classifyCfb(headerOnly)).toBe('compound-file');
  });

  it('is compound-file when the directory sector lies past the end of the file', () => {
    expect(classifyCfb(biff8().subarray(0, 1024))).toBe('compound-file');
  });

  it('terminates on a directory chain that loops back on itself', () => {
    const looping = buildCfb({ children: [stream('WordDocument')], lastDirectorySectorNext: 1 });
    expect(classifyCfb(looping)).toBe('compound-file');
  });

  it('terminates on a sibling tree that loops back on itself', () => {
    const looping = buildCfb({ children: [stream('WordDocument'), stream('1Table')], lastChildRightSibling: 1 });
    expect(classifyCfb(looping)).toBe('compound-file');
  });

  it('terminates on a DIFAT chain that loops back on itself', () => {
    // 301 FAT sectors need two DIFAT sectors, 301 and 302. The first is full, so
    // the reader does follow its last slot, which is pointed back at itself here.
    const looping = buildCfb({ children: [stream('WordDocument')], firstDirectorySector: 300 * 128 });
    new DataView(looping.buffer).setUint32((301 + 1) * 512 + 127 * 4, 301, true);
    expect(classifyCfb(looping)).toBe('compound-file');
  });

  it('reads a view into a larger buffer from its own offset', () => {
    const file = biff8();
    const padded = new Uint8Array(file.length + 64);
    padded.set(file, 64);
    expect(classifyCfb(padded.subarray(64))).toBe('legacy-xls');
  });
});

describe('input in another format is rejected with advice that fits the file', () => {
  it('tells the owner of a legacy .xls to convert it', async () => {
    const attempt = openZip(fromBuffer(fixture('biff8.xls')));
    await expect(attempt).rejects.toThrowError(OpenXmlUnsupportedFormatError);
    await expect(attempt).rejects.toMatchObject({ format: 'legacy-xls' });
    await expect(attempt).rejects.toThrowError(/legacy \.xls .* Convert it to \.xlsx/);
    await expect(attempt).rejects.not.toThrowError(/decrypt/i);
  });

  it('tells the owner of an encrypted xlsx to decrypt it', async () => {
    const attempt = openZip(fromBuffer(fixture('encrypted.xlsx')));
    await expect(attempt).rejects.toThrowError(OpenXmlUnsupportedFormatError);
    await expect(attempt).rejects.toMatchObject({ format: 'encrypted-xlsx' });
    await expect(attempt).rejects.toThrowError(/Encrypted xlsx is not supported/);
  });

  it('names both possibilities when the container cannot be classified', async () => {
    const headerOnly = new Uint8Array(512);
    headerOnly.set(CFB_MAGIC, 0);
    const attempt = openZip(fromBuffer(headerOnly));
    await expect(attempt).rejects.toMatchObject({ format: 'compound-file' });
    await expect(attempt).rejects.toThrowError(/password-protected xlsx .* legacy Office format/);
  });

  it('stays catchable as OpenXmlNotImplementedError', async () => {
    await expect(openZip(fromBuffer(biff8()))).rejects.toThrowError(OpenXmlNotImplementedError);
  });

  // Excel 2.x to 4.0 wrote the BIFF records with no compound file around them.
  // The four bytes are the BOF record id and its body length.
  it.each([
    ['BIFF2', [0x09, 0x00, 0x04, 0x00]],
    ['BIFF3', [0x09, 0x02, 0x06, 0x00]],
    ['BIFF4', [0x09, 0x04, 0x06, 0x00]],
  ])('recognises a raw %s .xls that has no compound file around it', async (_version, bof) => {
    const raw = new Uint8Array(64);
    raw.set(bof, 0);
    await expect(openZip(fromBuffer(raw))).rejects.toMatchObject({ format: 'legacy-xls' });
  });

  it('leaves other non-zip input to the zip reader', async () => {
    const notBof = new Uint8Array(64);
    notBof.set([0x09, 0x00, 0x20, 0x00], 0);
    await expect(openZip(fromBuffer(notBof))).rejects.not.toThrowError(OpenXmlUnsupportedFormatError);
  });

  it('reaches loadWorkbook and loadWorkbookStream callers', async () => {
    await expect(loadWorkbook(fromBuffer(biff8()))).rejects.toThrowError(/legacy \.xls/);
    await expect(loadWorkbookStream(fromBuffer(biff8()))).rejects.toThrowError(/legacy \.xls/);
  });
});
