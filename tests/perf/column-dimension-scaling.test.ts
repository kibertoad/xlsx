// Scaling gate for the bulk column-dimension helpers.
//
// `hideColumns` / `unhideColumns` / `setColumnWidths` and friends used to call
// `getColumnDimension` + `setColumnDimension` once per column, and each of
// those scans the whole `columnDimensions` map. Hiding N columns therefore
// scanned a map that grew as it went, so the cost rose with the square of the
// band. They now pair the whole band against the existing runs in one pass.
//
// This measures the shape of the curve rather than an absolute time, so it
// does not encode one machine's speed. Each measurement repeats the band until
// it is well clear of the timer's noise floor: a ratio between two
// single-millisecond samples turns on one GC pause, which would fail a correct
// build on a shared CI runner.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//   pnpm test:perf
//   PERF_GATE=1 pnpm test:perf   # asserts the ratio, as CI does

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { hideColumns } from '../../src/worksheet/worksheet.js';

const PERF_GATE = process.env['PERF_GATE'] === '1';

const SMALL_BAND = 4_000;
const LARGE_BAND = 8_000;
const REPEATS = 40;
/// Linear work doubles when the band doubles. The ceiling leaves room for timer
// noise and for the allocation the wider band does per entry, while still
// failing if the per-column scan comes back and squares the ratio.
const MAX_RATIO = 3;

/** Total time to hide `count` columns on each of `REPEATS` fresh worksheets. */
const timeHideColumns = (count: number): number => {
  const sheets = Array.from({ length: REPEATS }, () => addWorksheet(createWorkbook(), 'S'));
  const start = performance.now();
  for (const ws of sheets) hideColumns(ws, 1, count);
  const elapsed = performance.now() - start;
  for (const ws of sheets) {
    if (ws.columnDimensions.size !== count) {
      throw new Error(`expected ${count} entries, got ${ws.columnDimensions.size}`);
    }
  }
  return elapsed;
};

describe('bulk column-dimension scaling', () => {
  it('grows linearly with the band width', () => {
    timeHideColumns(SMALL_BAND); // warm-up, so the first measured run is not the JIT's
    const small = timeHideColumns(SMALL_BAND);
    const large = timeHideColumns(LARGE_BAND);
    const ratio = large / small;
    console.log(
      `hideColumns x${REPEATS}: ${SMALL_BAND} cols ${small.toFixed(1)}ms, ${LARGE_BAND} cols ${large.toFixed(1)}ms, ratio ${ratio.toFixed(2)}`,
    );
    if (PERF_GATE) expect(ratio).toBeLessThan(MAX_RATIO);
  });
});
