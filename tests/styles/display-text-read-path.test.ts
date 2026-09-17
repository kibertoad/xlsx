// End-to-end cover for the read path: build a workbook, write real bytes,
// load them back, and read the text off the reloaded cells. The number formats
// have to survive the stylesheet round-trip for these to pass, which is the
// part a same-process test on a freshly built workbook cannot show.

import { describe, expect, it } from 'vitest';
import type { Cell } from '../../src/cell/cell.js';
import { cellValueAsString } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getCellDate, getCellDisplayText, setCellNumberFormat } from '../../src/styles/index.js';
import type { Workbook } from '../../src/workbook/workbook.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { isWorksheetEmpty, iterRows, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

/** 2024-03-14 under the Windows epoch. */
const MARCH_14_2024 = 45_365;

const ROWS: Array<[string, number, string]> = [
  ['0.0%', 0.5, '50.0%'],
  ['#,##0.00', 1_234_567.891, '1,234,567.89'],
  ['General', 1.1 + 2.2, '3.3'],
  ['yyyy-mm-dd', MARCH_14_2024, '2024-03-14'],
  ['$#,##0.00;($#,##0.00)', -99.5, '($99.50)'],
  ['[h]:mm:ss', 1.5, '36:00:00'],
];

const buildBytes = async (date1904 = false): Promise<Uint8Array> => {
  const wb = createWorkbook({ date1904 });
  const ws = addWorksheet(wb, 'Report');
  setCell(ws, 1, 1, 'Format');
  setCell(ws, 1, 2, 'Value');
  ROWS.forEach(([code, value], index) => {
    const row = index + 2;
    setCell(ws, row, 1, code);
    setCellNumberFormat(wb, setCell(ws, row, 2, value), code);
  });
  return workbookToBytes(wb);
};

const worksheetOf = (wb: Workbook, title: string): Worksheet => {
  const ws = getSheet(wb, title);
  if (ws === undefined) throw new Error(`no worksheet ${title}`);
  return ws;
};

const textRows = (wb: Workbook, ws: Worksheet): string[][] =>
  [...iterRows(ws)].map((row) => row.map((cell) => (cell === undefined ? '' : getCellDisplayText(wb, cell))));

describe('reading back a workbook this library wrote', () => {
  it('renders each cell the way the format says', async () => {
    const wb = await loadWorkbook(fromBuffer(await buildBytes()));
    const ws = worksheetOf(wb, 'Report');
    expect(isWorksheetEmpty(ws)).toBe(false);
    expect(textRows(wb, ws)).toEqual([
      ['Format', 'Value'],
      ...ROWS.map(([code, , expected]) => [code, expected]),
    ]);
  });

  it('is what cellValueAsString cannot give, because it never sees the format', async () => {
    const wb = await loadWorkbook(fromBuffer(await buildBytes()));
    const ws = worksheetOf(wb, 'Report');
    const percent = cellAt(ws, 2, 2);
    expect(getCellDisplayText(wb, percent)).toBe('50.0%');
    expect(cellValueAsString(percent.value)).toBe('0.5');
  });

  it('reads a date-formatted serial back as a Date', async () => {
    const wb = await loadWorkbook(fromBuffer(await buildBytes()));
    const ws = worksheetOf(wb, 'Report');
    // The loader hands back the serial; the format is the only evidence that
    // it is a date, which is what getCellDate consults.
    expect(cellAt(ws, 5, 2).value).toBe(MARCH_14_2024);
    expect(getCellDate(wb, cellAt(ws, 5, 2))?.toISOString()).toBe('2024-03-14T00:00:00.000Z');
    // The elapsed row measures a span, so it has no date reading.
    expect(getCellDate(wb, cellAt(ws, 7, 2))).toBeUndefined();
  });

  it('honours the epoch recorded in the reloaded workbook', async () => {
    const wb = await loadWorkbook(fromBuffer(await buildBytes(true)));
    expect(wb.date1904).toBe(true);
    const ws = worksheetOf(wb, 'Report');
    expect(getCellDisplayText(wb, cellAt(ws, 5, 2))).toBe('2028-03-15');
    expect(getCellDate(wb, cellAt(ws, 5, 2))?.toISOString()).toBe('2028-03-15T00:00:00.000Z');
  });
});

function cellAt(ws: Worksheet, row: number, col: number): Cell {
  const cell = ws.rows.get(row)?.get(col);
  if (cell === undefined) throw new Error(`no cell at ${row},${col}`);
  return cell;
}
