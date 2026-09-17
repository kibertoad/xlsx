import { defineConfig } from 'vitest/config';

// Perf-only config. `pnpm test:perf` runs tests under tests/perf/ which the
// default config explicitly excludes (so `pnpm test` stays fast).
//
// The write-only throughput gate measures ≥500k cells/s for the 100k×30 shape
// on an M1 baseline. Set PERF_GATE=1 to fail the run on regressions.

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/perf/**/*.test.ts'],
    // Bench files run via `pnpm bench` (vitest bench), not the test runner.
    exclude: ['tests/perf/**/*.bench.ts', 'node_modules', 'dist', 'reference'],
    // The retention gate forces a GC before sampling, so it needs
    // `globalThis.gc`. Vitest spawns its workers with its own execArgv, so
    // neither NODE_OPTIONS nor the parent's flags reach them: the flag has to
    // be declared here or the gate has no way to tell retention from garbage.
    pool: 'forks',
    execArgv: ['--expose-gc'],
  },
});
