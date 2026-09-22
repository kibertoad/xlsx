// Scaling gate for the bulk column-dimension helpers.
//
// `hideColumns` / `unhideColumns` / `groupColumns` and friends used to call
// `getColumnDimension` + `setColumnDimension` once per column, and each of
// those scans the whole `columnDimensions` map. Hiding N columns therefore
// scanned a map that grew as it went, which is quadratic: 2000 columns took
// 23 ms and 4000 took 61 ms on the machine this was written on. They now pair
// the whole band against the existing runs in one pass.
//
// This measures the shape of the curve rather than an absolute time, so it does
// not encode one machine's speed. Excluded from the default `pnpm test` run
// (see vitest.config.ts). Run explicitly:
//
//   pnpm test:perf
//   PERF_GATE=1 pnpm test:perf   # asserts the ratio, as CI does
//
// PERF_GATE off by default: a loaded machine can skew a ratio between two
// millisecond-scale measurements badly enough to fail a correct build.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { hideColumns } from '../../src/worksheet/worksheet.js';

const PERF_GATE = process.env['PERF_GATE'] === '1';

const SMALL_BAND = 4_000;
const LARGE_BAND = 8_000;
// Linear work doubles when the band doubles. A quadratic term put the observed
// ratio near 2.7; the ceiling leaves room for timer noise and allocation
// effects while still failing if the per-column scan comes back.
const MAX_RATIO = 2.4;

const timeHideColumns = (count: number): number => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  const start = performance.now();
  hideColumns(ws, 1, count);
  const elapsed = performance.now() - start;
  if (ws.columnDimensions.size !== count) {
    throw new Error(`expected ${count} entries, got ${ws.columnDimensions.size}`);
  }
  return elapsed;
};

describe('bulk column-dimension scaling', () => {
  it('grows linearly with the band width', () => {
    timeHideColumns(SMALL_BAND); // warm-up, so the first measured run is not the JIT's
    const small = timeHideColumns(SMALL_BAND);
    const large = timeHideColumns(LARGE_BAND);
    // Guard against a divide-by-almost-zero when the small run is too fast to
    // time; the ratio only means something once there is something to measure.
    const ratio = large / Math.max(small, 0.5);
    // eslint-disable-next-line no-console
    console.log(
      `hideColumns: ${SMALL_BAND} cols ${small.toFixed(1)}ms, ${LARGE_BAND} cols ${large.toFixed(1)}ms, ratio ${ratio.toFixed(2)}`,
    );
    if (PERF_GATE) expect(ratio).toBeLessThan(MAX_RATIO);
  });
});
