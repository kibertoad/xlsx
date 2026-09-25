// Throughput gate for the read path, the mirror of throughput.test.ts. Reading
// is the direction most callers spend their time in, and `loadWorkbook` is
// ~98% of the cost of a load-then-walk, so a regression in the `<sheetData>`
// walk is invisible to the write-side gate.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//     pnpm test:perf
//     PERF_GATE=1 pnpm test:perf  # also assert the throughput floor
//
// 50k rows × 6 text columns is the shape the sheetData walk is tuned for: wide
// enough that per-cell cost dominates, small enough that one run is a few
// seconds.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { getNonEmptyCellCount } from '../../src/worksheet/worksheet.js';
import { measureAllocationProbe } from './machine-probe.js';

const ROWS = 50_000;
const COLS = 6;
const TOTAL_CELLS = ROWS * COLS;

// PERF_GATE off by default: a laptop idling or CI under load can dip below the
// floor without the code being wrong. Set the env var when you want a hard
// assertion (release branches, perf-regression PRs).
const PERF_GATE = process.env['PERF_GATE'] === '1';
// The gate is a ratio, not a cells/s number: read throughput over what
// {@link measureAllocationProbe} says this machine does with the same shape of
// allocation. A fixed cells/s floor cannot hold here, because the CI fleet's
// read throughput spans 88k to 239k across consecutive runs of unchanged code,
// which is wider than any regression worth catching. Dividing by the probe
// takes the machine back out of the number.
//
// Recalibrate by reading the `[perf]` line off a few CI runs, not by scaling a
// laptop measurement: the ratio is meant to hold across machines, but the
// margin under it has to come from CI's own spread. 0.01 is deliberately loose
// against the 0.017 this laptop reports, because no slow-runner ratio has been
// observed yet; tighten it once the CI band is known.
const FLOOR_CELLS_PER_PROBE_UNIT = 0.01;

const ITERATIONS = 3;

const buildArchive = async (): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Data');
  for (let r = 0; r < ROWS; r++) {
    const row = new Array<string>(COLS);
    for (let c = 0; c < COLS; c++) row[c] = `r${r}c${c}`;
    await ws.appendRow(row);
  }
  await ws.close();
  await wb.finalize();
  return sink.result();
};

const measureOnce = async (bytes: Uint8Array): Promise<{ seconds: number; cells: number }> => {
  const t0 = performance.now();
  const wb = await loadWorkbook(fromBuffer(bytes));
  const t1 = performance.now();
  const ws = getSheet(wb, 'Data');
  if (ws === undefined) throw new Error('load produced no Data sheet');
  return { seconds: (t1 - t0) / 1000, cells: getNonEmptyCellCount(ws) };
};

describe('perf: loadWorkbook read throughput', () => {
  it(
    `reads ${ROWS} × ${COLS} = ${TOTAL_CELLS.toLocaleString()} cells and reports cells/s`,
    async () => {
      const bytes = await buildArchive();
      // Warmup, discarded: the first pass pays JIT compile and heap growth for
      // the whole read path and reads systematically low.
      await measureOnce(bytes);
      // Best-of-N: shared CPUs and thermal throttling create wide variance per
      // run, but the best run reflects the pipeline's real ceiling.
      const runs: Array<{ seconds: number; cells: number }> = [];
      for (let i = 0; i < ITERATIONS; i++) runs.push(await measureOnce(bytes));
      const bestSeconds = Math.min(...runs.map((r) => r.seconds));
      const bestCellsPerSec = Math.round(TOTAL_CELLS / bestSeconds);
      // Probed after the reads, on the heap they left behind, so the probe
      // meets the allocator in the state the measurement did.
      const probeOpsPerSec = measureAllocationProbe();
      const perProbeUnit = bestCellsPerSec / probeOpsPerSec;

      const summaries = runs
        .map((r, i) => `#${i + 1} ${(TOTAL_CELLS / r.seconds).toFixed(0)} cells/s (${r.seconds.toFixed(2)}s)`)
        .join(' · ');
      process.stderr.write(
        `[perf] ${TOTAL_CELLS.toLocaleString()} cells × ${ITERATIONS} runs → best ${bestCellsPerSec.toLocaleString()} cells/s; archive ${bytes.byteLength.toLocaleString()} bytes\n        runs: ${summaries}\n        machine probe ${probeOpsPerSec.toLocaleString()} allocs/s → ${perProbeUnit.toFixed(4)} cells per probe unit (floor ${FLOOR_CELLS_PER_PROBE_UNIT})\n`,
      );

      // A load that dropped cells would post a flattering cells/s, so the gate
      // only means anything alongside the count.
      for (const run of runs) expect(run.cells).toBe(TOTAL_CELLS);
      if (PERF_GATE) {
        expect(perProbeUnit).toBeGreaterThanOrEqual(FLOOR_CELLS_PER_PROBE_UNIT);
      }
    },
    /* timeout */ 10 * 60_000,
  );
});
