// Retention metric for the modelled save path.
//
// Sheets are serialised straight into their ZIP entry and released, so by the
// time the last sheet is being written none of the earlier ones is still held.
// Collecting every sheet's bytes first and writing them in a second pass
// retains the whole set instead, and that sum is what decides whether a
// multi-sheet workbook fits in a container's heap.
//
// Measured after a forced GC from inside the sink, so this reads retention and
// not allocation churn. The streaming writer trades a large retained buffer for
// short-lived garbage, which means a sampled `heapUsed` peak moves the wrong way
// while the memory ceiling drops; retention is the honest measure.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). CI runs it
// via `PERF_GATE=1 pnpm test:perf`. Needs `--expose-gc` to assert; without it
// the forced GC is unavailable and the gate skips rather than reporting noise.

import { describe, expect, it } from 'vitest';
import type { XlsxSink } from '../../src/io/sink.js';
import { saveWorkbook } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { appendRow } from '../../src/worksheet/worksheet.js';

const SHEETS = 8;
const ROWS_PER_SHEET = 40_000;
const PERF_GATE = process.env['PERF_GATE'] === '1';

// Each sheet's part is roughly 7 MB uncompressed, so holding all eight costs
// about 60 MB. Streaming holds none of them; the reading sits near zero and the
// gate leaves generous room for allocator slack.
const MAX_RETAINED_MB = 20;

const gcFn = (globalThis as { gc?: () => void }).gc;
const forceGc = (): void => {
  if (typeof gcFn === 'function') {
    for (let i = 0; i < 3; i++) gcFn();
  }
};

/**
 * Drops every chunk, and once the archive is mostly written takes a
 * post-GC reading of what the save is still holding. Sampling from the sink is
 * what puts the probe inside the save rather than after it.
 */
const probingSink = (sampleAfterBytes: number): XlsxSink & { retained(): number } => {
  let total = 0;
  let sample = 0;
  return {
    toBytes() {
      return {
        write(chunk: Uint8Array) {
          total += chunk.byteLength;
          if (total >= sampleAfterBytes && sample === 0) {
            forceGc();
            const m = process.memoryUsage();
            sample = m.heapUsed + m.external;
          }
        },
        async finish() {
          return new Uint8Array(0);
        },
      };
    },
    retained() {
      return sample;
    },
  };
};

describe('modelled save retention', () => {
  it('does not hold earlier sheets while writing later ones', async () => {
    const wb = createWorkbook();
    for (let s = 0; s < SHEETS; s++) {
      const ws = addWorksheet(wb, `S${s}`);
      for (let r = 0; r < ROWS_PER_SHEET; r++) appendRow(ws, [r, `row-${s}-${r}`, r * 1.5, true, r % 7]);
    }

    forceGc();
    const base = process.memoryUsage();
    const baseline = base.heapUsed + base.external;
    // Sample once the deflated output is well past the first sheet, so every
    // earlier sheet would still be resident under a collect-then-write scheme.
    const sink = probingSink(1_500_000);
    await saveWorkbook(wb, sink);

    const retainedMb = (sink.retained() - baseline) / 1048576;
    console.log(`save retention over the workbook model: ${retainedMb.toFixed(1)} MB`);
    if (PERF_GATE && typeof gcFn === 'function') {
      expect(retainedMb).toBeLessThan(MAX_RETAINED_MB);
    }
  }, 300_000);
});
