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

// fflate encodes the DOS stamp with local-time getters, so the writer offsets
// the date before handing it over. Without that, a golden file committed from a
// laptop does not match the one CI renders from the same input.
describe('a pinned stamp survives the machine it was rendered on', () => {
  const original = process.env['TZ'];

  afterEach(() => {
    if (original === undefined) delete process.env['TZ'];
    else process.env['TZ'] = original;
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
