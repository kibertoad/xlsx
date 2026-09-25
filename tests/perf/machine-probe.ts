// Machine-speed probe for the throughput gates.
//
// A fixed cells/s floor assumes every runner reads at roughly the same speed.
// The CI fleet does not: across 15 consecutive runs the read gate reported
// between 88k and 239k cells/s, and the slow end sat 40% under the median
// while the write gate on the same job stayed within 2% of its own median. So
// the spread is not "the whole machine was busy", it is that allocation-heavy
// work varies far more across the fleet than compute-heavy work does, and only
// the read path is allocation-heavy.
//
// This probe measures that axis directly. It builds the same shape the
// worksheet model builds (a Map per row holding one small object per column
// with a string field) and reports how fast this machine does it. The read
// gate divides its throughput by the probe, so a runner with a slower memory
// subsystem scales both sides and the ratio holds.
//
// It deliberately touches no library code. If it called into `src/`, a
// regression in the read path would slow the probe too, the ratio would not
// move, and the gate would be blind to the thing it exists to catch.

const PROBE_ROWS = 50_000;
const PROBE_COLS = 6;
const PROBE_OPS = PROBE_ROWS * PROBE_COLS;

interface ProbeCell {
  value: string;
  index: number;
}

/**
 * One pass, returning seconds. The built structure is walked before the timer
 * is read so V8 cannot treat the allocation as dead.
 */
const probeOnce = (): number => {
  const t0 = performance.now();
  const rows = new Map<number, Map<number, ProbeCell>>();
  for (let r = 0; r < PROBE_ROWS; r++) {
    const row = new Map<number, ProbeCell>();
    for (let c = 0; c < PROBE_COLS; c++) {
      row.set(c, { value: `r${r}c${c}`, index: r * PROBE_COLS + c });
    }
    rows.set(r, row);
  }
  let reachable = 0;
  for (const row of rows.values()) reachable += row.size;
  const seconds = (performance.now() - t0) / 1000;
  if (reachable !== PROBE_OPS) {
    throw new Error(`machine probe built ${reachable} cells, expected ${PROBE_OPS}`);
  }
  return seconds;
};

/**
 * Cells-shaped allocations per second on this machine. Warmed up and
 * best-of-N for the same reason the gates are: the first pass pays JIT and
 * heap growth, and a shared CPU makes any single pass a lower bound.
 */
export const measureAllocationProbe = (iterations = 3): number => {
  probeOnce();
  let bestSeconds = Number.POSITIVE_INFINITY;
  for (let i = 0; i < iterations; i++) bestSeconds = Math.min(bestSeconds, probeOnce());
  return Math.round(PROBE_OPS / bestSeconds);
};
