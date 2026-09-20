// Just enough of an OLE Compound File Binary reader ([MS-CFB]) to say what a
// container holds. Two unrelated kinds of workbook arrive in one: a
// password-protected xlsx, which Excel stores as an `EncryptedPackage` stream,
// and a legacy BIFF `.xls`, whose cells live in a `Workbook` (BIFF8) or `Book`
// (BIFF5) stream. They need opposite advice, decrypt versus convert, and the
// stream names in the directory are the only cheap way to tell them apart.

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const HEADER_SIZE = 512;
const SECTOR_SHIFT_OFFSET = 0x1e;
const FIRST_DIRECTORY_SECTOR_OFFSET = 0x30;
const DIFAT_OFFSET = 0x4c;
const HEADER_DIFAT_ENTRIES = 109;
// Version 3 files use 512-byte sectors and version 4 files 4096-byte ones.
const SECTOR_SHIFTS: ReadonlySet<number> = new Set([9, 12]);
// Every id above this is a marker (DIFSECT, FATSECT, ENDOFCHAIN, FREESECT).
const MAX_REGULAR_SECTOR = 0xfffffffa;

const DIRECTORY_ENTRY_SIZE = 128;
const ENTRY_NAME_LENGTH_OFFSET = 0x40;
const ENTRY_TYPE_OFFSET = 0x42;
const ENTRY_NAME_MAX_BYTES = 64;
const ENTRY_TYPE_STREAM = 2;

// [MS-CFB] §2.6.4 compares entry names case-insensitively.
const ENCRYPTED_PACKAGE_STREAM = 'ENCRYPTEDPACKAGE';
const LEGACY_WORKBOOK_STREAMS: ReadonlySet<string> = new Set(['WORKBOOK', 'BOOK']);

export type CfbContent = 'encrypted-ooxml' | 'legacy-workbook' | 'unknown';

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
  if (nameBytes < 2 || nameBytes > ENTRY_NAME_MAX_BYTES) return undefined;
  let name = '';
  for (let i = 0; i < nameBytes - 2; i += 2) {
    name += String.fromCharCode(view.getUint16(entryOffset + i, true));
  }
  return name.toUpperCase();
};

/**
 * Classify a CFB container by the streams its directory lists. The scan is
 * flat: it does not track which storage owns a stream, so a `.doc` with an
 * embedded spreadsheet also reads as `legacy-workbook`. Anything the walk
 * cannot make sense of (truncated header, FAT beyond the header's 109 DIFAT
 * slots, a looping chain) is `unknown` rather than an error, because the
 * caller is already on its way to rejecting the file and only wants the best
 * available explanation.
 */
export function classifyCfb(bytes: Uint8Array): CfbContent {
  if (bytes.length < HEADER_SIZE) return 'unknown';
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorShift = view.getUint16(SECTOR_SHIFT_OFFSET, true);
  if (!SECTOR_SHIFTS.has(sectorShift)) return 'unknown';
  const sectorSize = 2 ** sectorShift;
  const fatEntriesPerSector = sectorSize / 4;
  // Sector 0 starts after the header, which occupies one full sector.
  const sectorOffset = (sector: number): number => (sector + 1) * sectorSize;

  const nextInChain = (sector: number): number | undefined => {
    const fatIndex = Math.floor(sector / fatEntriesPerSector);
    if (fatIndex >= HEADER_DIFAT_ENTRIES) return undefined;
    const fatSector = view.getUint32(DIFAT_OFFSET + fatIndex * 4, true);
    if (fatSector > MAX_REGULAR_SECTOR) return undefined;
    const at = sectorOffset(fatSector) + (sector % fatEntriesPerSector) * 4;
    return at + 4 <= bytes.length ? view.getUint32(at, true) : undefined;
  };

  let content: CfbContent = 'unknown';
  let sector: number | undefined = view.getUint32(FIRST_DIRECTORY_SECTOR_OFFSET, true);
  // A chain cannot hold more sectors than the file does, so this bound ends a
  // FAT that loops back on itself.
  let remaining = Math.ceil(bytes.length / sectorSize);
  while (sector !== undefined && sector <= MAX_REGULAR_SECTOR && remaining-- > 0) {
    const start = sectorOffset(sector);
    if (start + sectorSize > bytes.length) break;
    for (let entry = start; entry < start + sectorSize; entry += DIRECTORY_ENTRY_SIZE) {
      const name = streamName(view, entry);
      if (name === ENCRYPTED_PACKAGE_STREAM) return 'encrypted-ooxml';
      if (name !== undefined && LEGACY_WORKBOOK_STREAMS.has(name)) content = 'legacy-workbook';
    }
    sector = nextInChain(sector);
  }
  return content;
}
