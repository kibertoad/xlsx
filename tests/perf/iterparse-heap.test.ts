// Heap metric for the SAX iterator. `iterParse` feeds saxes in fixed-size
// chunks and drains the event queue between writes, so peak heap tracks the
// chunk size rather than the document size. Feeding saxes more than that in one
// `write()` queues every event that argument produces before the consumer sees
// the first, which scales with whatever it was handed.
//
// Both input shapes are covered because they fail differently. Bytes and
// strings are ours to chunk. A ReadableStream's chunk size belongs to the
// producer: the zip reader pushes 64 KB of compressed bytes per pull and
// worksheet XML inflates about sevenfold, so the default full-sheet walk sees
// chunks around 450 KB and needs the same cap applied on arrival.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//   pnpm test:perf
//   PERF_HEAP_GATE=1 pnpm test:perf   # asserts the ceiling
//
// Numbers depend on Node version and V8 GC mood, so the gate is off by default
// and each ceiling is set well above its observed figure.

import { describe, expect, it } from 'vitest';
import { iterParse, type SaxInput } from '../../src/xml/iterparse.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';

const PERF_HEAP_GATE = process.env['PERF_HEAP_GATE'] === '1';

const ROWS = 200_000;
const COLS = 5;
const EXPECTED_EVENTS = ROWS * (2 + COLS * 5) + 2;

/** Inflated worksheet chunk size the zip reader's stream actually emits. */
const ZIP_STREAM_CHUNK_BYTES = 450 * 1024;

// Ceilings are per shape because the two regress to very different figures.
// On this input, chunked feeding measures around 65 MB from bytes and around
// 95 MB from a 450 KB stream; unchunked it measures around 975 MB and around
// 220 MB. A single loose ceiling would clear the stream regression entirely,
// so each sits above its own observed figure and below its own regression.
const BYTES_CEILING_MB = 200;
const STREAM_CEILING_MB = 150;

const sheetBody = (rows: number, cols: number): Uint8Array => {
  const parts = [`<?xml version="1.0" encoding="UTF-8"?><sheetData xmlns="${SHEET_MAIN_NS}">`];
  for (let r = 1; r <= rows; r++) {
    parts.push(`<row r="${r}">`);
    for (let c = 1; c <= cols; c++) parts.push(`<c r="A${r}" t="n"><v>${r * c}</v></c>`);
    parts.push('</row>');
  }
  parts.push('</sheetData>');
  return new TextEncoder().encode(parts.join(''));
};

const inChunksOf = (bytes: Uint8Array, chunkBytes: number): ReadableStream<Uint8Array> => {
  let offset = 0;
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.byteLength) {
        controller.close();
        return;
      }
      const end = Math.min(offset + chunkBytes, bytes.byteLength);
      controller.enqueue(bytes.subarray(offset, end));
      offset = end;
    },
  });
};

const walk = async (label: string, ceilingMb: number, inputBytes: number, input: SaxInput): Promise<void> => {
  const gc = (globalThis as { gc?: () => void }).gc;
  if (typeof gc === 'function') gc();

  const before = process.memoryUsage().heapUsed;
  let peak = before;
  let events = 0;
  const t0 = performance.now();
  for await (const _ev of iterParse(input)) {
    events++;
    // Sampling every event would dominate the run; every 64k is frequent
    // enough to catch a queue that grows with the input.
    if ((events & 0xffff) === 0) {
      const used = process.memoryUsage().heapUsed;
      if (used > peak) peak = used;
    }
  }
  const used = process.memoryUsage().heapUsed;
  if (used > peak) peak = used;
  const seconds = (performance.now() - t0) / 1000;
  const peakMb = (peak - before) / (1024 * 1024);

  console.log(
    `[perf-iterparse] ${label}: ${(inputBytes / (1024 * 1024)).toFixed(0)} MB input, ` +
      `${events.toLocaleString()} events, +${peakMb.toFixed(0)} MB peak heap, ${seconds.toFixed(2)}s`,
  );

  // Guards the walk itself, gate or no gate: per row a start/end pair, per
  // cell a `<c>` pair, a `<v>` pair and one text event, plus `<sheetData>`.
  expect(events).toBe(EXPECTED_EVENTS);
  if (PERF_HEAP_GATE) {
    expect(peakMb).toBeLessThan(ceilingMb);
  }
};

describe('perf: iterParse peak heap tracks the chunk, not the input', () => {
  it(
    'walks a 40 MB sheet body handed over as bytes',
    async () => {
      const bytes = sheetBody(ROWS, COLS);
      await walk('bytes', BYTES_CEILING_MB, bytes.byteLength, bytes);
    },
    600_000,
  );

  it(
    'walks the same body arriving as 450 KB stream chunks',
    async () => {
      const bytes = sheetBody(ROWS, COLS);
      await walk('stream', STREAM_CEILING_MB, bytes.byteLength, inChunksOf(bytes, ZIP_STREAM_CHUNK_BYTES));
    },
    600_000,
  );
});
