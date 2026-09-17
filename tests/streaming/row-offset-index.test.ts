// Phase 4 §2.x — row-offset index correctness for sub-sheet
// streaming reads. iterRows({minRow}) builds a `<row r="N">` byte
// offset map on first use and binary-searches it instead of
// SAX-walking from start; this file pins the visible behaviour so a
// regression in either the index builder or the saxes wrap doesn't
// silently corrupt the iter output.

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';

const here = dirname(fileURLToPath(import.meta.url));
// Excel-emitted, unlike everything this library writes: the worksheet root
// declares x14ac and all 142 rows carry `x14ac:dyDescent`.
const EXCEL_SAMPLE = resolve(here, '../../reference/openpyxl/openpyxl/reader/tests/data/sample.xlsx');

const buildSheet = async (rows: number): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'A');
  for (let r = 1; r <= rows; r++) {
    setCell(ws, r, 1, r);
    setCell(ws, r, 2, `row-${r}`);
  }
  return workbookToBytes(wb);
};

describe('phase-4 — row-offset index for sub-sheet iter', () => {
  it('rejects cached band queries after the workbook closes', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await buildSheet(10)));
    const ws = wb.openWorksheet('A');
    for await (const _row of ws.iterRows({ minRow: 2 })) { /* prime the cache */ }
    await wb.close();
    expect(() => ws.iterRows({ minRow: 2 })).toThrow(/archive is closed/);
  });

  it('iterRows({ minRow: K }) yields only rows ≥ K', async () => {
    const bytes = await buildSheet(50);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 40 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]);
    await wb.close();
  });

  it('iterRows({ minRow, maxRow }) yields the band only', async () => {
    const bytes = await buildSheet(20);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 5, maxRow: 8 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([5, 6, 7, 8]);
    await wb.close();
  });

  it('iterRows({ minRow > maxKnownRow }) yields no rows without throwing', async () => {
    const bytes = await buildSheet(10);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 999 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([]);
    await wb.close();
  });

  it('the index is reused — second band query on the same worksheet still works', async () => {
    const bytes = await buildSheet(30);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const a: number[] = [];
    for await (const row of ws.iterRows({ minRow: 10, maxRow: 12 })) {
      const r = row[0]?.row;
      if (r !== undefined) a.push(r);
    }
    const b: number[] = [];
    for await (const row of ws.iterRows({ minRow: 25, maxRow: 27 })) {
      const r = row[0]?.row;
      if (r !== undefined) b.push(r);
    }
    expect(a).toEqual([10, 11, 12]);
    expect(b).toEqual([25, 26, 27]);
    await wb.close();
  });

  it('the band matches the full walk on an Excel-emitted sheet', async () => {
    // The band is replayed as a standalone document, which has to carry the
    // namespace declarations the worksheet root made. Rebuilding the envelope
    // from scratch drops them and every prefixed row attribute Excel writes
    // then fails to resolve. The fixtures above cannot catch it: our own
    // writer emits no prefixed attributes on <row>.
    const wb = await loadWorkbookStream(fromBuffer(readFileSync(EXCEL_SAMPLE)));
    const first = wb.sheetNames[0];
    if (first === undefined) throw new Error('fixture has no sheets');
    const ws = wb.openWorksheet(first);

    const cells = async (opts?: { minRow: number }) => {
      const out: Array<{ row: number; col: number; value: unknown }> = [];
      for await (const row of ws.iterRows(opts)) {
        for (const c of row) out.push({ row: c.row, col: c.col, value: c.value });
      }
      return out;
    };

    const all = await cells();
    const band = await cells({ minRow: 3 });
    expect(all.length).toBeGreaterThan(100);
    expect(band).toEqual(all.filter((c) => c.row >= 3));
    await wb.close();
  });

  it('preserves declarations on sheetData and processing instructions during band replay', async () => {
    const parts = unzipSync(await buildSheet(3));
    const path = 'xl/worksheets/sheet1.xml';
    const original = parts[path];
    if (!original) throw new Error('missing sheet');
    parts[path] = strToU8(strFromU8(original)
      .replace('<worksheet', '<?probe > <fake> ?><worksheet')
      .replace('<sheetData>', '<sheetData xmlns:custom="urn:test">')
      .replace(/<row /g, '<row custom:flag="yes" '));
    const wb = await loadWorkbookStream(fromBuffer(zipSync(parts)));
    try {
      const ws = wb.openWorksheet('A');
      const full = [];
      for await (const row of ws.iterRows()) full.push(row);
      const band = [];
      for await (const row of ws.iterRows({ minRow: 2 })) band.push(row);
      expect(band).toEqual(full.slice(1));
    } finally {
      await wb.close();
    }
  });

  it('cell values stay correct in the sliced path (sharedStrings + numeric mix)', async () => {
    const bytes = await buildSheet(10);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const last: ReadonlyArray<{ col: number; value: unknown }> = await (async () => {
      for await (const row of ws.iterRows({ minRow: 7, maxRow: 7 })) {
        return row.map((c) => ({ col: c.col, value: c.value }));
      }
      return [];
    })();
    expect(last).toEqual([
      { col: 1, value: 7 },
      { col: 2, value: 'row-7' },
    ]);
    await wb.close();
  });
});
