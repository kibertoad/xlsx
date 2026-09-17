// Heap metric for the SAX iterator. `iterParse` feeds saxes in fixed-size
// chunks and drains the event queue between writes, so peak heap tracks the
// chunk size rather than the document size. Feeding a whole document in one
// `write()` instead queues every event the document produces before the
// consumer sees the first, which scales with the document.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//   pnpm test:perf
//   PERF_HEAP_GATE=1 pnpm test:perf   # asserts the ceiling
//
// Numbers depend on Node version and V8 GC mood, so the gate is off by default
// and the ceiling is set well above the observed figure.

import { describe, expect, it } from 'vitest';
import { iterParse } from '../../src/xml/iterparse.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';

const PERF_HEAP_GATE = process.env['PERF_HEAP_GATE'] === '1';

const ROWS = 200_000;
const COLS = 5;

// Chunked feeding lands near 90 MB on Node 22 regardless of document size;
// whole-document feeding of this input lands near 975 MB. 300 MB separates the
// two by a wide margin in both directions.
const CEILING_MB = 300;

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

describe('perf: iterParse peak heap tracks the chunk, not the document', () => {
  it(
    'walks a 40 MB sheet body without buffering its whole event stream',
    async () => {
      const bytes = sheetBody(ROWS, COLS);
      const gc = (globalThis as { gc?: () => void }).gc;
      if (typeof gc === 'function') gc();

      const before = process.memoryUsage().heapUsed;
      let peak = before;
      let events = 0;
      const t0 = performance.now();
      for await (const _ev of iterParse(bytes)) {
        events++;
        // Sampling every event would dominate the run; every 64k is frequent
        // enough to catch a queue that grows with the document.
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
        `[perf-iterparse] ${(bytes.byteLength / (1024 * 1024)).toFixed(0)} MB input, ` +
          `${events.toLocaleString()} events, +${peakMb.toFixed(0)} MB peak heap, ${seconds.toFixed(2)}s`,
      );

      // Guards the walk itself, gate or no gate: per row a start/end pair, per
      // cell a `<c>` pair, a `<v>` pair and one text event, plus `<sheetData>`.
      expect(events).toBe(ROWS * (2 + COLS * 5) + 2);
      if (PERF_HEAP_GATE) {
        expect(peakMb).toBeLessThan(CEILING_MB);
      }
    },
    600_000,
  );
});
