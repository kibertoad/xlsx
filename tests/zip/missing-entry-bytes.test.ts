// A central directory can describe an entry the archive does not contain: a
// DEFLATE entry with no compressed bytes, or one whose declared span runs past
// EOF. Neither is decodable, both are reachable from `loadWorkbook` and
// `loadWorkbookStream` on untrusted input, and neither may hang the reader or
// yield whatever bytes happened to be there.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlIoError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { buildArchive, patchEntrySizes } from './_helpers.js';

const PAYLOAD = new TextEncoder().encode('a payload long enough for deflate to do something '.repeat(20));

/** Drain a stream to a byte count, releasing the reader lock either way. */
const drain = async (stream: ReadableStream<Uint8Array>): Promise<number> => {
  const reader = stream.getReader();
  let total = 0;
  try {
    for (let next = await reader.read(); !next.done; next = await reader.read()) {
      total += next.value.byteLength;
    }
  } finally {
    reader.releaseLock();
  }
  return total;
};

describe('DEFLATE entry with no compressed bytes', () => {
  const craft = async (): Promise<Uint8Array> =>
    patchEntrySizes(await buildArchive([{ path: 'a.xml', bytes: PAYLOAD }]), 'a.xml', { compSize: 0 });

  it('rejects a streaming read', async () => {
    const archive = await openZip(fromBuffer(await craft()));
    try {
      expect(() => archive.readStream('a.xml')).toThrowError(OpenXmlIoError);
      expect(() => archive.readStream('a.xml')).toThrowError('carries no compressed bytes');
    } finally {
      archive.close();
    }
  });

  it('rejects a buffered read rather than returning an empty part', async () => {
    const archive = await openZip(fromBuffer(await craft()));
    try {
      expect(() => archive.read('a.xml')).toThrowError(OpenXmlIoError);
      expect(() => archive.read('a.xml')).toThrowError('carries no compressed bytes');
    } finally {
      archive.close();
    }
  });

  it('fails a workbook load instead of running forever', async () => {
    // The vector the guard exists for: `loadWorkbookStream` reads worksheet
    // parts through `readStream`, so a sheet part shaped like this reaches the
    // inflater loop. The whole iteration has to settle, not just start.
    const wb = createWorkbook();
    const sheet = addWorksheet(wb, 'Alpha');
    for (let r = 1; r <= 20; r++) setCell(sheet, r, 1, `row-${r}`);
    const bytes = patchEntrySizes(await workbookToBytes(wb), 'xl/worksheets/sheet1.xml', { compSize: 0 });

    await expect(
      (async () => {
        const streamed = await loadWorkbookStream(fromBuffer(bytes));
        try {
          for await (const _row of streamed.openWorksheet('Alpha').iterRows()) {
            /* drain */
          }
        } finally {
          await streamed.close();
        }
      })(),
    ).rejects.toThrowError(OpenXmlIoError);
  }, 20_000);

  it('still reads an entry whose deflate payload decodes to nothing', async () => {
    // The legitimate neighbour: a zero-byte file does get a deflate stream,
    // just one that emits no bytes. It has to close cleanly on both paths.
    const archive = await openZip(fromBuffer(await buildArchive([{ path: 'a.xml', bytes: new Uint8Array(0) }])));
    try {
      expect(await drain(archive.readStream('a.xml'))).toBe(0);
      expect(archive.read('a.xml').byteLength).toBe(0);
    } finally {
      archive.close();
    }
  });
});

describe('entry whose declared compressed size runs past EOF', () => {
  // `subarray` clamps, so the reader has to compare the declared span against
  // the archive length itself. A truncated binary part has no parser to catch
  // it downstream: passthrough parts go back out on save as-is.
  it.each([
    { label: 'DEFLATE', compress: true },
    { label: 'STORE', compress: false },
  ])('rejects both read paths ($label)', async ({ compress }) => {
    const honest = await buildArchive([{ path: 'a.bin', bytes: PAYLOAD, compress }]);
    const archive = await openZip(fromBuffer(patchEntrySizes(honest, 'a.bin', { compSize: honest.length * 2 })));
    try {
      expect(() => archive.read('a.bin')).toThrowError('past the end of the');
      expect(() => archive.readStream('a.bin')).toThrowError(OpenXmlIoError);
    } finally {
      archive.close();
    }
  });
});
