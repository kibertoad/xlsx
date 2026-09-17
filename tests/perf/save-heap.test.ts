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
// via `PERF_GATE=1 pnpm test:perf`, whose config passes `--expose-gc` to the
// worker. Under the gate a missing `globalThis.gc` fails the test: a forced GC
// is what separates retention from allocation churn, so without it the number
// is noise and a silent skip would leave the gate green forever.

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

    // A sink that never crossed sampleAfterBytes reports 0, which would make
    // the delta hugely negative and pass the gate on a measurement that never
    // happened. Check the probe fired before reading anything into the number.
    const sampled = sink.retained();
    expect(sampled, 'the sink never reached sampleAfterBytes, so nothing was measured').toBeGreaterThan(0);

    const retainedMb = (sampled - baseline) / 1048576;
    // stderr, matching heap.test.ts: the runner swallows console output from a
    // passing test, and the number is the whole point of the run.
    process.stderr.write(`[perf-save-heap] save retention over the workbook model: ${retainedMb.toFixed(1)} MB\n`);
    if (PERF_GATE) {
      expect(typeof gcFn, 'PERF_GATE needs --expose-gc; vitest.perf.config.ts passes it').toBe('function');
      expect(retainedMb).toBeLessThan(MAX_RETAINED_MB);
    }
  }, 300_000);
});
