// The re-read cache is bounded so a full load does not end up holding every
// uncompressed part. Object identity is the observable: a cached entry hands
// back the same array, an uncached one re-inflates into a fresh array with the
// same contents.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

const CACHE_MAX_ENTRY_BYTES = 64 * 1024;

// Deflate-friendly but not degenerate, so compressed size stays a sane fraction
// of the payload and the ratio guard has nothing to complain about.
const filler = (n: number): Uint8Array => {
  const out = new Uint8Array(n);
  for (let i = 0; i < n; i++) out[i] = 32 + ((i * 7 + (i >> 5)) % 90);
  return out;
};

const buildArchive = async (
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

describe('inflate cache bounds', () => {
  it('keeps a small entry so a second read does not re-inflate', async () => {
    const small = filler(4 * 1024);
    const archive = await openZip(fromBuffer(await buildArchive([{ path: 'a/_rels/a.xml.rels', bytes: small }])));
    const first = archive.read('a/_rels/a.xml.rels');
    const second = archive.read('a/_rels/a.xml.rels');
    expect(second).toBe(first);
    expect(second).toEqual(small);
    archive.close();
  });

  it('does not retain an entry larger than the cache ceiling', async () => {
    const big = filler(CACHE_MAX_ENTRY_BYTES + 1024);
    const archive = await openZip(fromBuffer(await buildArchive([{ path: 'xl/worksheets/sheet1.xml', bytes: big }])));
    const first = archive.read('xl/worksheets/sheet1.xml');
    const second = archive.read('xl/worksheets/sheet1.xml');
    // Same bytes, different array: the payload was re-inflated rather than held
    // for the lifetime of the archive.
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    archive.close();
  });

  it('applies the ceiling to STORE entries too', async () => {
    const big = filler(CACHE_MAX_ENTRY_BYTES + 1024);
    const archive = await openZip(
      fromBuffer(await buildArchive([{ path: 'xl/media/image1.png', bytes: big, compress: false }])),
    );
    const first = archive.read('xl/media/image1.png');
    const second = archive.read('xl/media/image1.png');
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    archive.close();
  });

  it('evicts oldest-first once the cached total passes its ceiling', async () => {
    // 4 MB ceiling, 40 KB per entry: 128 entries overshoot it comfortably.
    const entries = Array.from({ length: 128 }, (_, i) => ({
      path: `part${i}.xml`,
      bytes: filler(40 * 1024),
    }));
    const archive = await openZip(fromBuffer(await buildArchive(entries)));
    const firstRead = archive.read('part0.xml');
    for (const e of entries) archive.read(e.path);
    // part0 was evicted to make room, so it comes back as a fresh array.
    expect(archive.read('part0.xml')).not.toBe(firstRead);
    // The most recent entry is still resident.
    expect(archive.read('part127.xml')).toBe(archive.read('part127.xml'));
    archive.close();
  });

  it('decodes non-ASCII entry names through the shared decoder', async () => {
    const payload = new TextEncoder().encode('ok');
    const archive = await openZip(
      fromBuffer(await buildArchive([{ path: 'xl/media/画像1.png', bytes: payload }, { path: 'b/ü.xml', bytes: payload }])),
    );
    expect(archive.has('xl/media/画像1.png')).toBe(true);
    expect(archive.has('b/ü.xml')).toBe(true);
    expect(archive.read('xl/media/画像1.png')).toEqual(payload);
    archive.close();
  });
});

describe('decompression budget across repeated reads', () => {
  it('charges an entry once however many times it is inflated', async () => {
    const payload = filler(200 * 1024);
    const bytes = await buildArchive([{ path: 'big.xml', bytes: payload }]);
    // Room for one copy of the entry and change, but not two. Before the budget
    // tracked charges per path, a stream followed by a read counted the payload
    // twice and tripped the archive-total guard on a legitimate file.
    const archive = await openZip(fromBuffer(bytes), {
      decompressionLimits: { maxTotalUncompressedBytes: Math.floor(payload.byteLength * 1.5) },
    });

    const reader = archive.readStream('big.xml').getReader();
    for (;;) {
      const { done } = await reader.read();
      if (done) break;
    }
    expect(archive.read('big.xml')).toEqual(payload);
    expect(archive.read('big.xml')).toEqual(payload);
    archive.close();
  });

  it('still rejects an archive whose distinct entries exceed the total', async () => {
    const payload = filler(200 * 1024);
    const bytes = await buildArchive([
      { path: 'one.xml', bytes: payload },
      { path: 'two.xml', bytes: payload },
    ]);
    // Two distinct entries still sum against the archive ceiling, so refunding a
    // re-read has not weakened the guard. The declared-totals check catches this
    // at open time, before a single byte is inflated.
    await expect(
      openZip(fromBuffer(bytes), {
        decompressionLimits: { maxTotalUncompressedBytes: Math.floor(payload.byteLength * 1.5) },
      }),
    ).rejects.toThrow(/decompression-bomb guard/);
  });
});
