// A loaded `<col min="1" max="16384" width="12"/>` is one entry covering every
// column, which is what Excel writes when the user sets a sheet-wide width.
// Editing one column inside such a run has to leave the other columns with the
// width they had; the run used to be deleted wholesale.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import {
  autofitColumns,
  collapseColumnGroup,
  expandColumnGroup,
  getColumnDimension,
  groupColumns,
  hideColumn,
  hideColumns,
  setCell,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  ungroupColumns,
  unhideColumn,
  unhideColumns,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

const withRun = (): Worksheet => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  ws.columnDimensions.set(1, { min: 1, max: 10, width: 12, customWidth: true });
  return ws;
};

const widths = (ws: Worksheet, from: number, to: number): Array<number | undefined> => {
  const out: Array<number | undefined> = [];
  for (let c = from; c <= to; c++) out.push(getColumnDimension(ws, c)?.width);
  return out;
};

describe('editing one column of a multi-column run', () => {
  it('splits the run around the column setColumnWidth touches', () => {
    const ws = withRun();
    setColumnWidth(ws, 3, 20);
    expect(widths(ws, 1, 11)).toEqual([12, 12, 20, 12, 12, 12, 12, 12, 12, 12, undefined]);
    expect([...ws.columnDimensions.values()].map((d) => [d.min, d.max])).toEqual(
      expect.arrayContaining([
        [1, 2],
        [3, 3],
        [4, 10],
      ]),
    );
  });

  it('keeps the left neighbours when the first column of a run is edited', () => {
    const ws = withRun();
    setColumnWidth(ws, 1, 20);
    expect(widths(ws, 1, 10)).toEqual([20, 12, 12, 12, 12, 12, 12, 12, 12, 12]);
  });

  it('keeps the right neighbours when the last column of a run is edited', () => {
    const ws = withRun();
    setColumnWidth(ws, 10, 20);
    expect(widths(ws, 1, 10)).toEqual([12, 12, 12, 12, 12, 12, 12, 12, 12, 20]);
  });

  it('carries the run fields onto the edited column', () => {
    const ws = withRun();
    hideColumn(ws, 5);
    const edited = getColumnDimension(ws, 5);
    expect(edited).toMatchObject({ min: 5, max: 5, width: 12, customWidth: true, hidden: true });
  });

  it('splits the run when a field is dropped rather than added', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.columnDimensions.set(1, { min: 1, max: 10, width: 12, hidden: true });
    unhideColumn(ws, 4);
    expect(getColumnDimension(ws, 4)).toMatchObject({ width: 12 });
    expect(getColumnDimension(ws, 4)?.hidden).toBeUndefined();
    expect(getColumnDimension(ws, 3)?.hidden).toBe(true);
    expect(getColumnDimension(ws, 5)?.hidden).toBe(true);
  });

  it('drops the entry when the last field is dropped, leaving the rest of the run', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.columnDimensions.set(1, { min: 1, max: 5, hidden: true });
    unhideColumn(ws, 3);
    expect(getColumnDimension(ws, 3)).toBeUndefined();
    expect(getColumnDimension(ws, 2)?.hidden).toBe(true);
    expect(getColumnDimension(ws, 4)?.hidden).toBe(true);
  });

  it('splits around a whole band, not just one column', () => {
    const ws = withRun();
    hideColumns(ws, 4, 6);
    expect(widths(ws, 1, 10)).toEqual([12, 12, 12, 12, 12, 12, 12, 12, 12, 12]);
    for (let c = 4; c <= 6; c++) expect(getColumnDimension(ws, c)?.hidden).toBe(true);
    for (const c of [3, 7]) expect(getColumnDimension(ws, c)?.hidden).toBeUndefined();
  });

  it('survives a load, an edit and a save', async () => {
    const source = createWorkbook();
    const sheet = addWorksheet(source, 'S');
    sheet.columnDimensions.set(1, { min: 1, max: 10, width: 12, customWidth: true });
    setCell(sheet, 1, 1, 'x');

    const loaded = await loadWorkbook(fromBuffer(await workbookToBytes(source)));
    const ws = getSheet(loaded, 'S');
    if (!ws) throw new Error('expected sheet S');
    // The reader keeps the run as one entry, which is the shape that used to
    // lose its siblings.
    expect(getColumnDimension(ws, 7)?.max).toBe(10);
    setColumnWidth(ws, 3, 20);

    const again = await loadWorkbook(fromBuffer(await workbookToBytes(loaded)));
    const reread = getSheet(again, 'S');
    if (!reread) throw new Error('expected sheet S');
    expect(widths(reread, 1, 10)).toEqual([12, 12, 20, 12, 12, 12, 12, 12, 12, 12]);
  });
});

describe('bulk column edits', () => {
  it('group then ungroup leaves the run fields intact', () => {
    const ws = withRun();
    groupColumns(ws, 2, 4, 2);
    expect(getColumnDimension(ws, 3)?.outlineLevel).toBe(2);
    ungroupColumns(ws, 2, 4);
    expect(getColumnDimension(ws, 3)?.outlineLevel).toBeUndefined();
    expect(widths(ws, 1, 10)).toEqual([12, 12, 12, 12, 12, 12, 12, 12, 12, 12]);
  });

  it('collapse then expand leaves the run fields intact', () => {
    const ws = withRun();
    collapseColumnGroup(ws, 2, 4);
    expect(getColumnDimension(ws, 3)?.hidden).toBe(true);
    expandColumnGroup(ws, 2, 4);
    expect(getColumnDimension(ws, 3)?.hidden).toBeUndefined();
    expect(widths(ws, 1, 10)).toEqual([12, 12, 12, 12, 12, 12, 12, 12, 12, 12]);
  });

  it('setColumnWidths splits the run once for the whole band', () => {
    const ws = withRun();
    setColumnWidths(ws, [30, 31, 32], 2);
    expect(widths(ws, 1, 6)).toEqual([12, 30, 31, 32, 12, 12]);
    expect(getColumnDimension(ws, 2)?.customWidth).toBe(true);
  });

  it('autofitColumns keeps the rest of a run when it widens one column', () => {
    const ws = withRun();
    setCell(ws, 1, 3, 'a considerably longer value');
    autofitColumns(ws);
    expect(getColumnDimension(ws, 3)?.width).toBeGreaterThan(12);
    expect(widths(ws, 1, 2)).toEqual([12, 12]);
    expect(widths(ws, 4, 10)).toEqual([12, 12, 12, 12, 12, 12, 12]);
  });

  it('unhideColumns over a band that holds no entries is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    unhideColumns(ws, 1, 5);
    expect(ws.columnDimensions.size).toBe(0);
  });
});

describe('column entries the helpers write', () => {
  it('setColumnDimension with no fields still registers the column', () => {
    const ws = withRun();
    const entry = setColumnDimension(ws, 5, {});
    expect(entry).toEqual({ min: 5, max: 5 });
    // The caller gets the entry the sheet holds, so mutating it is a write.
    expect(getColumnDimension(ws, 5)).toBe(entry);
    expect(widths(ws, 4, 6)).toEqual([12, undefined, 12]);
  });

  it('rejects a band past the last column before building it', () => {
    const ws = withRun();
    expect(() => hideColumns(ws, 1, 50_000_000)).toThrow(OpenXmlSchemaError);
    expect(() => hideColumns(ws, 1, 50_000_000)).toThrow(/out of range/);
    expect(ws.columnDimensions.size).toBe(1);
  });

  it('unhideColumn rejects a column outside the sheet under its own name', () => {
    const ws = withRun();
    expect(() => unhideColumn(ws, 0)).toThrow(/Worksheet col 0 out of range/);
  });
});

describe('runs that overlap each other', () => {
  it('keeps the columns only the second run covers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.columnDimensions.set(1, { min: 1, max: 10, width: 11 });
    ws.columnDimensions.set(5, { min: 5, max: 20, width: 22 });
    setColumnWidth(ws, 7, 99);
    // 1-10 read as the first run before the edit and still do; 11-20 keep the
    // second run's width instead of losing it with the split.
    expect(widths(ws, 1, 20)).toEqual([
      11, 11, 11, 11, 11, 11, 99, 11, 11, 11, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22,
    ]);
  });

  it('gives up only the columns an entry filed under a foreign key claims', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.columnDimensions.set(1, { min: 1, max: 10, width: 11 });
    // A caller-written entry whose key is not its `min`. Column 3 is the one
    // key the split cannot have, so it is the only one that loses its width.
    ws.columnDimensions.set(3, { min: 50, max: 60, width: 33 });
    setColumnWidth(ws, 2, 99);
    expect(widths(ws, 1, 10)).toEqual([11, 99, undefined, 11, 11, 11, 11, 11, 11, 11]);
    expect(getColumnDimension(ws, 55)?.width).toBe(33);
  });
});
