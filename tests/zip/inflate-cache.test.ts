// The re-read cache is bounded, so a full load does not end up holding every
// uncompressed part, and `read` hands each caller an array of its own whether
// the bytes came from the cache or from a fresh inflate.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import {
  CACHE_MAX_ENTRY_BYTES,
  CACHE_MAX_TOTAL_BYTES,
  createInflateCache,
} from '../../src/zip/inflate-cache.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

// Small enough to cache, and 128 of them come to exactly the total ceiling, so
// the eviction cases below sit on an exact boundary instead of near one.
const ENTRY_BYTES = CACHE_MAX_TOTAL_BYTES / 128;
const ENTRIES_TO_FILL = 128;

const block = (n: number, fill: number): Uint8Array => new Uint8Array(n).fill(fill);

const fillCache = (cache: ReturnType<typeof createInflateCache>, count: number): void => {
  for (let i = 0; i < count; i++) cache.set(`part${i}`, block(ENTRY_BYTES, 7));
};

describe('inflate cache', () => {
  it('keeps an entry at the size ceiling and turns away the next byte', () => {
    const cache = createInflateCache();
    cache.set('at-ceiling', block(CACHE_MAX_ENTRY_BYTES, 1));
    cache.set('over-ceiling', block(CACHE_MAX_ENTRY_BYTES + 1, 2));
    expect(cache.get('at-ceiling')).toEqual(block(CACHE_MAX_ENTRY_BYTES, 1));
    expect(cache.get('over-ceiling')).toBeUndefined();
  });

  it('shares no array with its callers', () => {
    const cache = createInflateCache();
    const handedIn = block(16, 1);
    cache.set('a', handedIn);
    handedIn[0] = 9;
    const handedOut = cache.get('a');
    expect(handedOut?.[0]).toBe(1);
    if (handedOut) handedOut[1] = 9;
    expect(cache.get('a')?.[1]).toBe(1);
  });

  it('holds a full set of entries up to the total ceiling', () => {
    const cache = createInflateCache();
    fillCache(cache, ENTRIES_TO_FILL);
    expect(cache.get('part0')).toBeDefined();
    expect(cache.get(`part${ENTRIES_TO_FILL - 1}`)).toBeDefined();
  });

  it('evicts the least recently used entry once the total is exceeded', () => {
    const cache = createInflateCache();
    fillCache(cache, ENTRIES_TO_FILL);
    cache.set('one-too-many', block(ENTRY_BYTES, 7));
    expect(cache.get('part0')).toBeUndefined();
    expect(cache.get('part1')).toBeDefined();
    expect(cache.get('one-too-many')).toBeDefined();
  });

  it('spares an entry that has been read since it was cached', () => {
    const cache = createInflateCache();
    fillCache(cache, ENTRIES_TO_FILL);
    // Reading part0 makes it the most recent, so the next eviction has to fall
    // on part1 instead. Without that, the parts a load keeps coming back to are
    // precisely the ones it drops.
    expect(cache.get('part0')).toBeDefined();
    cache.set('one-too-many', block(ENTRY_BYTES, 7));
    expect(cache.get('part0')).toBeDefined();
    expect(cache.get('part1')).toBeUndefined();
  });

  it('counts an entry cached twice once', () => {
    const cache = createInflateCache();
    fillCache(cache, ENTRIES_TO_FILL - 1);
    cache.set('part0', block(ENTRY_BYTES, 7));
    cache.set('last', block(ENTRY_BYTES, 7));
    // 128 distinct entries, exactly the ceiling: counting part0 twice would
    // have pushed the total over it and evicted part0 as the oldest.
    expect(cache.get('part0')).toBeDefined();
  });

  it('gives back its whole allowance on clear', () => {
    const cache = createInflateCache();
    fillCache(cache, ENTRIES_TO_FILL);
    cache.clear();
    expect(cache.get('part0')).toBeUndefined();
    fillCache(cache, ENTRIES_TO_FILL);
    expect(cache.get('part0')).toBeDefined();
  });
});

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

const drain = async (stream: ReadableStream<Uint8Array>): Promise<number> => {
  const reader = stream.getReader();
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) return total;
    total += value?.byteLength ?? 0;
  }
};

describe('archive reads', () => {
  it('hands every read an array the caller owns', async () => {
    // Under the cache ceiling and deflated, over it and deflated, and stored:
    // whichever way the bytes reach the caller, they are the caller's.
    const parts = [
      { path: 'a/_rels/a.xml.rels', bytes: filler(4 * 1024) },
      { path: 'xl/worksheets/sheet1.xml', bytes: filler(CACHE_MAX_ENTRY_BYTES + 1024) },
      { path: 'xl/media/image1.png', bytes: filler(4 * 1024), compress: false },
    ];
    const archive = await openZip(fromBuffer(await buildArchive(parts)));
    for (const part of parts) {
      const first = archive.read(part.path);
      first[0] = 0xff;
      expect(archive.read(part.path)).toEqual(part.bytes);
    }
    archive.close();
  });

  it('decodes non-ASCII entry names through the shared decoder', async () => {
    const payload = new TextEncoder().encode('ok');
    const archive = await openZip(
      fromBuffer(
        await buildArchive([
          { path: 'xl/media/画像1.png', bytes: payload },
          { path: 'b/ü.xml', bytes: payload },
        ]),
      ),
    );
    expect(archive.has('xl/media/画像1.png')).toBe(true);
    expect(archive.has('b/ü.xml')).toBe(true);
    expect(archive.read('xl/media/画像1.png')).toEqual(payload);
    archive.close();
  });
});

describe('decompression budget across repeated reads', () => {
  const payload = filler(200 * 1024);
  // Room for one copy of the entry and change, but not two. Before the budget
  // charged each path its high-water mark, inflating the same entry twice
  // counted the payload twice and tripped the archive-total guard on a
  // legitimate file.
  const openWithRoomForOne = async (): ReturnType<typeof openZip> =>
    openZip(fromBuffer(await buildArchive([{ path: 'big.xml', bytes: payload }])), {
      decompressionLimits: { maxTotalUncompressedBytes: Math.floor(payload.byteLength * 1.5) },
    });

  it('charges an entry once however many times it is inflated', async () => {
    const archive = await openWithRoomForOne();
    expect(await drain(archive.readStream('big.xml'))).toBe(payload.byteLength);
    expect(archive.read('big.xml')).toEqual(payload);
    expect(archive.read('big.xml')).toEqual(payload);
    archive.close();
  });

  it('charges an entry once when both streams open before either inflates', async () => {
    const archive = await openWithRoomForOne();
    // What two `iterRows` calls over one sheet do: neither stream has inflated
    // a byte by the time the second one is created.
    const first = archive.readStream('big.xml');
    const second = archive.readStream('big.xml');
    expect(await drain(first)).toBe(payload.byteLength);
    expect(await drain(second)).toBe(payload.byteLength);
    archive.close();
  });

  it('charges an entry once when two reads of it interleave', async () => {
    const archive = await openWithRoomForOne();
    const a = archive.readStream('big.xml').getReader();
    const b = archive.readStream('big.xml').getReader();
    let fromA = 0;
    let fromB = 0;
    for (let aDone = false, bDone = false; !aDone || !bDone; ) {
      if (!aDone) {
        const chunk = await a.read();
        aDone = chunk.done;
        fromA += chunk.value?.byteLength ?? 0;
      }
      if (!bDone) {
        const chunk = await b.read();
        bDone = chunk.done;
        fromB += chunk.value?.byteLength ?? 0;
      }
    }
    expect(fromA).toBe(payload.byteLength);
    expect(fromB).toBe(payload.byteLength);
    archive.close();
  });

  it('still rejects an archive whose distinct entries exceed the total', async () => {
    const bytes = await buildArchive([
      { path: 'one.xml', bytes: payload },
      { path: 'two.xml', bytes: payload },
    ]);
    // Two distinct entries still sum against the archive ceiling, so charging a
    // re-read once has not weakened the guard. The declared-totals check
    // catches this at open time, before a single byte is inflated.
    await expect(
      openZip(fromBuffer(bytes), {
        decompressionLimits: { maxTotalUncompressedBytes: Math.floor(payload.byteLength * 1.5) },
      }),
    ).rejects.toThrow(/decompression-bomb guard/);
  });
});
