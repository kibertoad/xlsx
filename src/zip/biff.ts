// Excel 2.x to 4.0 wrote a `.xls` as a bare BIFF record stream with no OLE
// compound file around it, so `./cfb.ts` never sees one. Such a file opens with
// a BOF record, and the record's 4-byte header is enough to recognise it.

const BOF_RECORD_ID_LOW_BYTE = 0x09;
// The BOF record id is 0x0009, 0x0209, 0x0409 or 0x0809 depending on the BIFF
// version. The high byte is the key and the value is the record body lengths
// that version writes (BIFF5 writes 8 bytes under 0x0809 and BIFF8 writes 16).
const BOF_BODY_LENGTHS: ReadonlyMap<number, ReadonlySet<number>> = new Map([
  [0x00, new Set([4])],
  [0x02, new Set([6])],
  [0x04, new Set([6])],
  [0x08, new Set([8, 16])],
]);
const RECORD_HEADER_BYTES = 4;

export const isRawBiffWorkbook = (bytes: Uint8Array): boolean => {
  if (bytes.length < RECORD_HEADER_BYTES || bytes[0] !== BOF_RECORD_ID_LOW_BYTE) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const bodyLength = view.getUint16(2, true);
  return BOF_BODY_LENGTHS.get(view.getUint8(1))?.has(bodyLength) === true;
};
