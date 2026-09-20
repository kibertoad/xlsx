// Just enough of an OLE Compound File Binary reader ([MS-CFB]) to say what a
// container holds. Two unrelated kinds of workbook arrive in one: a
// password-protected xlsx, which Excel stores as an `EncryptedPackage` stream,
// and a legacy BIFF `.xls`, whose cells live in a `Workbook` (BIFF8) or `Book`
// (BIFF5) stream. They need opposite advice, decrypt versus convert, and the
// streams in the root storage are the only cheap way to tell them apart.

import type { UnsupportedFormat } from '../utils/exceptions.js';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const HEADER_SIZE = 512;
const SECTOR_SHIFT_OFFSET = 0x1e;
const FIRST_DIRECTORY_SECTOR_OFFSET = 0x30;
const FIRST_DIFAT_SECTOR_OFFSET = 0x44;
const HEADER_DIFAT_OFFSET = 0x4c;
const HEADER_DIFAT_ENTRIES = 109;
// Version 3 files use 512-byte sectors and version 4 files 4096-byte ones.
const SECTOR_SHIFTS: ReadonlySet<number> = new Set([9, 12]);
const SECTOR_ID_BYTES = 4;
// Every id above this is a marker (DIFSECT, FATSECT, ENDOFCHAIN, FREESECT).
const MAX_REGULAR_SECTOR = 0xfffffffa;

const DIRECTORY_ENTRY_SIZE = 128;
const ROOT_ENTRY_ID = 0;
const ENTRY_NAME_LENGTH_OFFSET = 0x40;
const ENTRY_TYPE_OFFSET = 0x42;
const ENTRY_LEFT_SIBLING_OFFSET = 0x44;
const ENTRY_RIGHT_SIBLING_OFFSET = 0x48;
const ENTRY_CHILD_OFFSET = 0x4c;
const ENTRY_NAME_MAX_BYTES = 64;
const ENTRY_TYPE_STREAM = 2;
const UTF16_UNIT_BYTES = 2;

// [MS-CFB] §2.6.4 compares entry names case-insensitively.
const ENCRYPTED_PACKAGE_STREAM = 'ENCRYPTEDPACKAGE';
const LEGACY_WORKBOOK_STREAMS: ReadonlySet<string> = new Set(['WORKBOOK', 'BOOK']);

export const isCfbCompoundDocument = (bytes: Uint8Array): boolean => {
  if (bytes.length < CFB_MAGIC.length) return false;
  for (let i = 0; i < CFB_MAGIC.length; i++) {
    if (bytes[i] !== CFB_MAGIC[i]) return false;
  }
  return true;
};

const streamName = (view: DataView, entryOffset: number): string | undefined => {
  if (view.getUint8(entryOffset + ENTRY_TYPE_OFFSET) !== ENTRY_TYPE_STREAM) return undefined;
  // The stored length is in bytes and counts the UTF-16 null terminator.
  const nameBytes = view.getUint16(entryOffset + ENTRY_NAME_LENGTH_OFFSET, true);
  if (nameBytes < UTF16_UNIT_BYTES || nameBytes > ENTRY_NAME_MAX_BYTES) return undefined;
  let name = '';
  for (let i = 0; i < nameBytes - UTF16_UNIT_BYTES; i += UTF16_UNIT_BYTES) {
    name += String.fromCharCode(view.getUint16(entryOffset + i, true));
  }
  return name.toUpperCase();
};

/**
 * Classify a CFB container by the streams directly inside its root storage.
 * Streams in nested storages are skipped on purpose: a `.xls` that embeds a
 * password-protected document carries that document's `EncryptedPackage`
 * inside an `MBD...` storage, and a `.doc` that embeds a sheet carries a
 * nested `Workbook`, and neither says what the outer file is.
 *
 * Anything the walk cannot make sense of (a truncated file, a sector id past
 * the end, a looping chain or sibling tree) ends the walk instead of throwing,
 * and what was seen up to that point decides the answer, `compound-file` if
 * nothing was. The caller is already on its way to rejecting the file and only
 * wants the best available explanation.
 */
export function classifyCfb(bytes: Uint8Array): UnsupportedFormat {
  if (bytes.length < HEADER_SIZE) return 'compound-file';
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorShift = view.getUint16(SECTOR_SHIFT_OFFSET, true);
  if (!SECTOR_SHIFTS.has(sectorShift)) return 'compound-file';
  const sectorSize = 2 ** sectorShift;
  const idsPerSector = sectorSize / SECTOR_ID_BYTES;
  const sectorCount = Math.ceil(bytes.length / sectorSize);

  // Sector 0 starts after the header, which occupies one full sector.
  // `undefined` for a marker id and for a sector the file is too short to hold.
  const sectorStart = (sector: number): number | undefined => {
    if (sector > MAX_REGULAR_SECTOR) return undefined;
    const start = (sector + 1) * sectorSize;
    return start + sectorSize <= bytes.length ? start : undefined;
  };

  // The header lists the first 109 FAT sectors. A file past 109 * idsPerSector
  // sectors (about 6.8 MB at 512 bytes) continues the list in a chain of DIFAT
  // sectors, each ending in the id of the next. A writer that puts the
  // directory after the data leaves it out of the header's reach in such a file.
  const fatSectors: number[] = [];
  // A FAT maps every sector of the file, so it cannot need more sectors than
  // this. The cap is what ends a DIFAT chain that loops back on itself.
  const maxFatSectors = Math.ceil(sectorCount / idsPerSector);
  // False once the list has ended: FAT sectors are listed without gaps, so the
  // first marker id is the end.
  const collectFatSectors = (offset: number, slots: number): boolean => {
    for (let i = 0; i < slots; i++) {
      const id = view.getUint32(offset + i * SECTOR_ID_BYTES, true);
      if (id > MAX_REGULAR_SECTOR || fatSectors.length >= maxFatSectors) return false;
      fatSectors.push(id);
    }
    return true;
  };
  if (collectFatSectors(HEADER_DIFAT_OFFSET, HEADER_DIFAT_ENTRIES)) {
    const nextDifatSlot = (idsPerSector - 1) * SECTOR_ID_BYTES;
    let difat = sectorStart(view.getUint32(FIRST_DIFAT_SECTOR_OFFSET, true));
    while (difat !== undefined && collectFatSectors(difat, idsPerSector - 1)) {
      difat = sectorStart(view.getUint32(difat + nextDifatSlot, true));
    }
  }

  const nextInChain = (sector: number): number | undefined => {
    const fatSector = fatSectors[Math.floor(sector / idsPerSector)];
    const start = fatSector === undefined ? undefined : sectorStart(fatSector);
    if (start === undefined) return undefined;
    return view.getUint32(start + (sector % idsPerSector) * SECTOR_ID_BYTES, true);
  };

  // Byte offsets of the directory's sectors in chain order, so an entry id
  // resolves to an offset without walking the chain again.
  const directory: number[] = [];
  let sector: number | undefined = view.getUint32(FIRST_DIRECTORY_SECTOR_OFFSET, true);
  // A chain cannot hold more sectors than the file does, so this bound ends a
  // FAT that loops back on itself.
  let remaining = sectorCount;
  while (sector !== undefined && remaining-- > 0) {
    const start = sectorStart(sector);
    if (start === undefined) break;
    directory.push(start);
    sector = nextInChain(sector);
  }

  const entriesPerSector = sectorSize / DIRECTORY_ENTRY_SIZE;
  // `undefined` for NOSTREAM and for any other id the directory does not reach.
  const entryOffset = (id: number): number | undefined => {
    const start = directory[Math.floor(id / entriesPerSector)];
    return start === undefined ? undefined : start + (id % entriesPerSector) * DIRECTORY_ENTRY_SIZE;
  };

  const root = entryOffset(ROOT_ENTRY_ID);
  if (root === undefined) return 'compound-file';
  // A storage's children form a binary tree: the storage points at one child,
  // and the rest hang off that child's sibling ids. Following sibling ids only,
  // never an entry's own child id, keeps the walk inside the root storage.
  const pending = [view.getUint32(root + ENTRY_CHILD_OFFSET, true)];
  const visited = new Set<number>();
  let format: UnsupportedFormat = 'compound-file';
  for (let id = pending.pop(); id !== undefined; id = pending.pop()) {
    const entry = entryOffset(id);
    if (entry === undefined || visited.has(id)) continue;
    visited.add(id);
    const name = streamName(view, entry);
    if (name === ENCRYPTED_PACKAGE_STREAM) return 'encrypted-xlsx';
    if (name !== undefined && LEGACY_WORKBOOK_STREAMS.has(name)) format = 'legacy-xls';
    pending.push(
      view.getUint32(entry + ENTRY_LEFT_SIBLING_OFFSET, true),
      view.getUint32(entry + ENTRY_RIGHT_SIBLING_OFFSET, true),
    );
  }
  return format;
}
