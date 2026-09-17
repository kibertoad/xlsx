// Row-offset-index sanity checks for the streaming reader.
// Compares iterRows({minRow}) wall-time on a tail-of-sheet read
// against a full-walk baseline. The index path should be dramatically
// faster for large sheets accessed near the bottom.
//
// The second case covers repeat band queries on one worksheet handle. A
// worksheet part is far too large for the archive's re-read cache, so the
// handle has to hold the bytes it indexed: without that, every query inflates
// the whole part again.
//
// Excluded from the default `pnpm test` run via vitest.config.ts.
// Run with: pnpm test:perf

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';

const PERF_ROW_INDEX_GATE = process.env['PERF_ROW_INDEX_GATE'] === '1';
// Row-index path should beat the no-min walk by a substantial margin
// on this shape; if it ever flips the wrong way, something has
// regressed in the slice-and-wrap logic.
const MIN_SPEEDUP = 2;
// A repeat band query re-uses the inflated part and its index, so all it does
// is SAX over the 100 rows it yields. The first query inflates ~1.7 MB and
// scans it; anything close to parity means the part is being re-read.
const MIN_REPEAT_SPEEDUP = 3;

describe('phase-4 perf — row-index speedup', () => {
  it(
    'iterRows({minRow: ~lastRow}) is faster than walking from start',
    async () => {
      const ROWS = 10_000;
      const COLS = 5;
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink);
      const ws = await wb.addWorksheet('Big');
      for (let r = 0; r < ROWS; r++) {
        const row = new Array<number>(COLS);
        for (let c = 0; c < COLS; c++) row[c] = r * COLS + c;
        await ws.appendRow(row);
      }
      await ws.close();
      await wb.finalize();
      const archive = sink.result();

      const fresh = async (): Promise<{ ms: number; count: number }> => {
        const reader = await loadWorkbookStream(fromBuffer(archive));
        const sheet = reader.openWorksheet('Big');
        const t0 = performance.now();
        let count = 0;
        for await (const _row of sheet.iterRows({ minRow: ROWS - 99, maxRow: ROWS })) {
          count++;
        }
        const t1 = performance.now();
        await reader.close();
        return { ms: t1 - t0, count };
      };

      // Baseline: no-min walk → full-sheet iter (count cap to drop late rows).
      const baseline = async (): Promise<{ ms: number; count: number }> => {
        const reader = await loadWorkbookStream(fromBuffer(archive));
        const sheet = reader.openWorksheet('Big');
        const t0 = performance.now();
        let count = 0;
        let lastRow = 0;
        for await (const row of sheet.iterRows()) {
          lastRow = row[0]?.row ?? lastRow;
          if (lastRow >= ROWS - 99) count++;
        }
        const t1 = performance.now();
        await reader.close();
        return { ms: t1 - t0, count };
      };

      // Warm-up to stabilise V8 + the row-index cache for `fresh` runs.
      await fresh();
      await baseline();

      // Best-of-3 because perf tests are noisy on shared CPUs.
      const indexRuns: Array<{ ms: number; count: number }> = [];
      const fullRuns: Array<{ ms: number; count: number }> = [];
      for (let i = 0; i < 3; i++) {
        indexRuns.push(await fresh());
        fullRuns.push(await baseline());
      }
      const indexMs = Math.min(...indexRuns.map((r) => r.ms));
      const fullMs = Math.min(...fullRuns.map((r) => r.ms));
      const speedup = fullMs / Math.max(indexMs, 0.001);

      process.stderr.write(
        `[perf-row-index] tail-100-of-${ROWS.toLocaleString()} rows: index ${indexMs.toFixed(2)}ms, full ${fullMs.toFixed(2)}ms, speedup ${speedup.toFixed(1)}x\n`,
      );

      // Both paths must report the same row count (50k − 100 + 1 = 100).
      expect(indexRuns[0]?.count).toBe(100);
      expect(fullRuns[0]?.count).toBe(100);
      if (PERF_ROW_INDEX_GATE) {
        expect(speedup).toBeGreaterThan(MIN_SPEEDUP);
      }
    },
    /* timeout */ 5 * 60_000,
  );

  it(
    'a second band query on the same worksheet does not inflate the part again',
    async () => {
      const ROWS = 20_000;
      const COLS = 5;
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink);
      const ws = await wb.addWorksheet('Big');
      for (let r = 0; r < ROWS; r++) {
        const row = new Array<number>(COLS);
        for (let c = 0; c < COLS; c++) row[c] = r * COLS + c;
        await ws.appendRow(row);
      }
      await ws.close();
      await wb.finalize();

      const reader = await loadWorkbookStream(fromBuffer(sink.result()));
      const sheet = reader.openWorksheet('Big');
      const band = async (): Promise<{ ms: number; count: number }> => {
        const t0 = performance.now();
        let count = 0;
        for await (const _row of sheet.iterRows({ minRow: ROWS - 99 })) count++;
        return { ms: performance.now() - t0, count };
      };

      const first = await band();
      // Best-of-3 on the repeats, for the same reason the case above takes one.
      const repeats = [await band(), await band(), await band()];
      const repeatMs = Math.min(...repeats.map((r) => r.ms));
      const speedup = first.ms / Math.max(repeatMs, 0.001);
      await reader.close();

      process.stderr.write(
        `[perf-row-index] repeat band query on ${ROWS.toLocaleString()} rows: first ${first.ms.toFixed(2)}ms, repeat ${repeatMs.toFixed(2)}ms, speedup ${speedup.toFixed(1)}x\n`,
      );

      expect(first.count).toBe(100);
      for (const r of repeats) expect(r.count).toBe(100);
      if (PERF_ROW_INDEX_GATE) {
        expect(speedup).toBeGreaterThan(MIN_REPEAT_SPEEDUP);
      }
    },
    /* timeout */ 5 * 60_000,
  );
});
