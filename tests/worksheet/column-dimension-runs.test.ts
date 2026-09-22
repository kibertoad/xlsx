// A loaded `<col min="1" max="16384" width="12"/>` is one entry covering every
// column, which is what Excel writes when the user sets a sheet-wide width.
// Editing one column inside such a run has to leave the other columns with the
// width they had; the run used to be deleted wholesale.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import {
  collapseColumnGroup,
  expandColumnGroup,
  getColumnDimension,
  groupColumns,
  hideColumn,
  hideColumns,
  setCell,
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

  it('setColumnWidths assigns a different width per column in one pass', () => {
    const ws = withRun();
    setColumnWidths(ws, [30, 31, 32], 2);
    expect(widths(ws, 1, 6)).toEqual([12, 30, 31, 32, 12, 12]);
  });

  it('unhideColumns over a band that holds no entries is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    unhideColumns(ws, 1, 5);
    expect(ws.columnDimensions.size).toBe(0);
  });
});
