// Tests for SaveOptions.mtime / compressionLevel.
//
// ZIP has no "no timestamp" encoding, so fflate stamps the wall clock into each
// entry by default and two saves of the same workbook differ in bytes. Pinning
// mtime is what makes golden-file tests and content-addressed caching possible.

import { afterEach, describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { saveWorkbook, workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { OpenXmlIoError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { appendRows, getCell } from '../../src/worksheet/worksheet.js';
import { createZipWriter } from '../../src/zip/writer.js';

const STAMP = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));
const STAMP_ISO = STAMP.toISOString();

/**
 * Decode the DOS date/time out of the first local file header. Layout per
 * APPNOTE 4.3.7: signature(4) version(2) flags(2) method(2) time(2) date(2),
 * each little-endian, the pair packed as year-1980/month/day and
 * hour/minute/second-over-two.
 */
const firstEntryStamp = (zip: Uint8Array): string => {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  expect(view.getUint32(0, true)).toBe(0x04034b50);
  const time = view.getUint16(10, true);
  const date = view.getUint16(12, true);
  const pad = (n: number): string => String(n).padStart(2, '0');
  const y = 1980 + (date >>> 9);
  return `${y}-${pad((date >>> 5) & 0xf)}-${pad(date & 0x1f)} ${pad(time >>> 11)}:${pad((time >>> 5) & 0x3f)}:${pad((time & 0x1f) * 2)}`;
};

const buildReport = () => {
  const wb = createWorkbook();
  // Caller-controlled stamps: the wall clock has no business in a report built
  // from a payload.
  wb.properties = { creator: 'renderer', created: STAMP_ISO, modified: STAMP_ISO };
  const ws = addWorksheet(wb, 'Leverage');
  appendRows(ws, [
    ['Language', 'Words'],
    ['de', 71_579],
    ['fr', 12_004],
  ]);
  return wb;
};

describe('deterministic output', () => {
  it('a pinned mtime makes two saves byte-identical', async () => {
    const first = await workbookToBytes(buildReport(), { mtime: STAMP });
    const second = await workbookToBytes(buildReport(), { mtime: STAMP });
    expect(second).toEqual(first);
  });

  it('without a pinned mtime the bytes carry the wall clock', async () => {
    const pinned = await workbookToBytes(buildReport(), { mtime: STAMP });
    const unpinned = await workbookToBytes(buildReport());
    expect(unpinned).not.toEqual(pinned);
  });

  it('a different stamp is a different archive', async () => {
    // The assertion above can only see a drifting clock. This one fails
    // outright if mtime stops reaching the entries, whatever the clock says.
    const early = await workbookToBytes(buildReport(), { mtime: STAMP });
    const late = await workbookToBytes(buildReport(), { mtime: new Date(Date.UTC(2026, 5, 7, 8, 9, 0)) });
    expect(late).not.toEqual(early);
  });

  it('the entry stamp is the pinned time read in UTC', async () => {
    const bytes = await workbookToBytes(buildReport(), { mtime: STAMP });
    expect(firstEntryStamp(bytes)).toBe('2026-01-02 03:04:00');
  });

  it('the pinned archive still loads', async () => {
    const bytes = await workbookToBytes(buildReport(), { mtime: STAMP });
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = getSheet(wb, 'Leverage');
    expect(ws).toBeDefined();
    if (!ws) return;
    expect(getCell(ws, 2, 2)?.value).toBe(71_579);
    expect(wb.properties?.creator).toBe('renderer');
  });

  it('compressionLevel reaches the deflate stream', async () => {
    const stored = await workbookToBytes(buildReport(), { mtime: STAMP, compressionLevel: 0 });
    const squeezed = await workbookToBytes(buildReport(), { mtime: STAMP, compressionLevel: 9 });
    expect(stored.byteLength).toBeGreaterThan(squeezed.byteLength);
    // Level 0 still has to round-trip.
    const wb = await loadWorkbook(fromBuffer(stored));
    expect(getSheet(wb, 'Leverage')).toBeDefined();
  });

  it('saveWorkbook threads the options through to the sink', async () => {
    const sinkA = toBuffer();
    const sinkB = toBuffer();
    await saveWorkbook(buildReport(), sinkA, { mtime: STAMP });
    await saveWorkbook(buildReport(), sinkB, { mtime: STAMP });
    expect(sinkB.result()).toEqual(sinkA.result());
  });

  it('the streaming writer honours compressionLevel too', async () => {
    const render = async (compressionLevel: 0 | 9): Promise<Buffer> => {
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink, { mtime: STAMP, compressionLevel });
      const ws = await wb.addWorksheet('Data');
      for (let i = 0; i < 200; i++) await ws.appendRow(['de', 71_579]);
      await ws.close();
      await wb.finalize();
      return sink.result();
    };
    expect((await render(0)).byteLength).toBeGreaterThan((await render(9)).byteLength);
  });

  it('the streaming writer takes the same option', async () => {
    const render = async (): Promise<Buffer> => {
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink, { mtime: STAMP });
      const ws = await wb.addWorksheet('Data');
      await ws.appendRow(['de', 71_579]);
      await ws.close();
      await wb.finalize();
      return sink.result();
    };
    expect(await render()).toEqual(await render());
  });
});

// Pinned headers must agree even when UTC components describe a local DST gap.
describe('a pinned stamp survives the machine it was rendered on', () => {
  const original = process.env['TZ'];

  afterEach(() => {
    if (original === undefined) delete process.env['TZ'];
    else process.env['TZ'] = original;
  });

  it.each([
    ['America/New_York', '2026-03-08T02:30:00Z'],
    ['America/New_York', '2026-11-01T02:30:00Z'],
    ['Europe/Berlin', '2026-03-29T02:30:00Z'],
    ['Pacific/Apia', '2011-12-30T12:00:00Z'],
  ])('preserves UTC components across clock transitions in %s (%s)', async (zone, iso) => {
    const mtime = new Date(iso);
    process.env['TZ'] = 'UTC';
    const utc = await workbookToBytes(buildReport(), { mtime });
    process.env['TZ'] = zone;
    const local = await workbookToBytes(buildReport(), { mtime });
    expect(firstEntryStamp(local)).toBe(iso.replace('T', ' ').replace('Z', ''));
    expect(local).toEqual(utc);
  });

  it('pins local and central headers without modifying stored payloads', async () => {
    process.env['TZ'] = 'America/New_York';
    const mtime = new Date('2026-03-08T02:30:00Z');
    const payload = new Uint8Array(40);
    new DataView(payload.buffer).setUint32(0, 0x04034b50, true);
    const sink = toBuffer();
    const zip = createZipWriter(sink, { mtime });
    await zip.addEntry('stored', payload, { compress: false });
    const entry = zip.addStreamingEntry('streamed', { compress: false });
    entry.write(payload);
    await entry.end();
    await zip.finalize();
    const bytes = sink.result();
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const end = bytes.byteLength - 22;
    let central = view.getUint32(end + 16, true);
    for (let i = 0; i < 2; i++) {
      expect(view.getUint32(central, true)).toBe(0x02014b50);
      const local = view.getUint32(central + 42, true);
      expect(firstEntryStamp(bytes.subarray(local))).toBe('2026-03-08 02:30:00');
      expect(view.getUint32(central + 12, true)).toBe(view.getUint32(local + 10, true));
      const data = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      expect(new Uint8Array(bytes.subarray(data, data + payload.length))).toEqual(payload);
      central += 46 + view.getUint16(central + 28, true) + view.getUint16(central + 30, true) + view.getUint16(central + 32, true);
    }
  });

  it.each([
    ['1980-01-01T00:00:00Z', '1980-01-01 00:00:00'],
    ['2099-12-31T23:59:59Z', '2099-12-31 23:59:58'],
  ])('preserves boundary dates at ZIP resolution: %s', async (iso, expected) => {
    process.env['TZ'] = 'Pacific/Kiritimati';
    const bytes = await workbookToBytes(buildReport(), { mtime: new Date(iso) });
    expect(firstEntryStamp(bytes)).toBe(expected);
  });

  it('renders the same bytes in two timezones', async () => {
    process.env['TZ'] = 'UTC';
    const utc = await workbookToBytes(buildReport(), { mtime: STAMP });
    process.env['TZ'] = 'Asia/Tokyo';
    const tokyo = await workbookToBytes(buildReport(), { mtime: STAMP });
    expect(tokyo).toEqual(utc);
    expect(firstEntryStamp(tokyo)).toBe('2026-01-02 03:04:00');
  });
});

describe('both options are checked before a byte is written', () => {
  it('rejects a date ZIP cannot record', async () => {
    // The Unix epoch is the reflex choice for a reproducible build, and it is
    // 10 years before the DOS date field begins.
    await expect(workbookToBytes(buildReport(), { mtime: new Date(0) })).rejects.toThrow(/1980-2099/);
    await expect(workbookToBytes(buildReport(), { mtime: new Date('nonsense') })).rejects.toThrow(
      /invalid Date/,
    );
  });

  it('rejects a level fflate would quietly replace with its default', async () => {
    // @ts-expect-error deliberately wrong: the union covers TypeScript callers, not JS ones
    await expect(workbookToBytes(buildReport(), { compressionLevel: 12 })).rejects.toThrow(/\[0, 9\]/);
    // @ts-expect-error deliberately wrong
    await expect(workbookToBytes(buildReport(), { compressionLevel: 2.5 })).rejects.toThrow(OpenXmlIoError);
  });

  it('leaves the sink empty when an option is rejected', async () => {
    const sink = toBuffer();
    await expect(saveWorkbook(buildReport(), sink, { mtime: new Date(0) })).rejects.toThrow(OpenXmlIoError);
    expect(sink.result().byteLength).toBe(0);
  });
});
