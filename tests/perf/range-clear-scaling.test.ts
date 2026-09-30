// Cost gate for the range walks that delete cells.
//
// A range can name far more coordinates than the sheet holds cells: `'A:A'` is
// 1_048_576 coordinates and `'A1:XFD1048576'` is seventeen billion. The sparse
// cases fail if the walk visits every coordinate of the rectangle; the last
// case fails if it visits every row of the store for a small range.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//   pnpm test:perf
//   PERF_GATE=1 pnpm test:perf   # asserts the ceilings, as CI does

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { clearRange, mergeCells, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const PERF_GATE = process.env['PERF_GATE'] === '1';

/** Comfortably above a sparse walk, far below an area walk of these bands. */
const CEILING_MS = 50;

/** Far above a two-by-two walk, below one pass over `MANY_ROWS` stored rows. */
const SMALL_RANGE_CEILING_MS = 2;
const MANY_ROWS = 200_000;

const sparseSheet = (): Worksheet => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  for (let r = 1; r <= 20; r++) setCell(ws, r, 1, r);
  return ws;
};

const time = (label: string, run: () => void): number => {
  const start = performance.now();
  run();
  const elapsed = performance.now() - start;
  console.log(`${label}: ${elapsed.toFixed(1)}ms`);
  return elapsed;
};

describe('range walks follow the cells, not the rectangle', () => {
  it('merges a whole-column band over a sparse sheet', () => {
    const ws = sparseSheet();
    // A hundred columns, so an area walk costs about a hundred million lookups.
    const elapsed = time("mergeCells('A:CV') over 20 cells", () => mergeCells(ws, 'A:CV'));
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('merges the whole grid over a sparse sheet', () => {
    // Seventeen billion coordinates. An area walk never finishes this.
    const ws = sparseSheet();
    const elapsed = time("mergeCells('A1:XFD1048576') over 20 cells", () =>
      mergeCells(ws, 'A1:XFD1048576'),
    );
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('clears the whole grid over a sparse sheet', () => {
    const ws = sparseSheet();
    let removed = 0;
    const elapsed = time("clearRange('A1:XFD1048576') over 20 cells", () => {
      removed = clearRange(ws, 'A1:XFD1048576');
    });
    expect(removed).toBe(20);
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('stays cheap for a small range on a sheet with many rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    for (let r = 1; r <= MANY_ROWS; r++) setCell(ws, r, 1, r);
    const elapsed = time(`clearRange('A1:B2') over ${MANY_ROWS} rows`, () => clearRange(ws, 'A1:B2'));
    if (PERF_GATE) expect(elapsed).toBeLessThan(SMALL_RANGE_CEILING_MS);
  });
});
