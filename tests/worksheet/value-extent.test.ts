// getValueExtent and iterRows({ extent: 'values' }) against the shape that
// motivates them: a sheet formatted past its content, which every reading that
// counts cell objects reports as taller and wider than its data.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { registerCellStyle } from '../../src/styles/cell-style.js';
import { addWorksheet, createWorkbook, getSheetByIndex } from '../../src/workbook/workbook.js';
import {
  getDataExtent,
  getMaxCol,
  getMaxRow,
  getValueExtent,
  iterValues,
  setCell,
} from '../../src/worksheet/worksheet.js';

/** Four rows of data, then two rows carrying nothing but a style id. */
const makeFormattedPastItsData = () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  for (let r = 1; r <= 4; r++) {
    setCell(ws, r, 1, `a${r}`);
    setCell(ws, r, 2, r);
  }
  const shaded = registerCellStyle(wb, { numberFormat: '#,##0' });
  setCell(ws, 5, 1, null, shaded);
  setCell(ws, 6, 1, null, shaded);
  return { wb, ws };
};

describe('getValueExtent', () => {
  it('stops at the last cell holding a value', () => {
    const { ws } = makeFormattedPastItsData();
    expect(getValueExtent(ws)).toEqual({ minRow: 1, maxRow: 4, minCol: 1, maxCol: 2 });
  });

  it('getDataExtent still counts the formatting-only cells', () => {
    const { ws } = makeFormattedPastItsData();
    expect(getDataExtent(ws)).toEqual({ minRow: 1, maxRow: 6, minCol: 1, maxCol: 2 });
    expect(getMaxRow(ws)).toBe(6);
    expect(getMaxCol(ws)).toBe(2);
  });

  it('narrows the column bound too, not just the row bound', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 9, null, registerCellStyle(wb, { numberFormat: '0.00' }));
    expect(getValueExtent(ws)).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
    expect(getMaxCol(ws)).toBe(9);
  });

  it('reports the leading blank band too, so the box is a true bounding box', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, null, registerCellStyle(wb, { numberFormat: '0.00' }));
    setCell(ws, 7, 3, 'only value');
    expect(getValueExtent(ws)).toEqual({ minRow: 7, maxRow: 7, minCol: 3, maxCol: 3 });
  });

  it('returns undefined when the sheet holds formatting but no value', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const shaded = registerCellStyle(wb, { numberFormat: '0.00' });
    setCell(ws, 1, 1, null, shaded);
    setCell(ws, 20, 5, null, shaded);
    expect(getValueExtent(ws)).toBeUndefined();
    expect(getDataExtent(ws)).toEqual({ minRow: 1, maxRow: 20, minCol: 1, maxCol: 5 });
  });

  it('returns undefined for a sheet with no cells at all', () => {
    const wb = createWorkbook();
    expect(getValueExtent(addWorksheet(wb, 'S'))).toBeUndefined();
  });
});

describe("iterValues with extent: 'values'", () => {
  it('yields one row per row of data, with no phantom tail', () => {
    const { ws } = makeFormattedPastItsData();
    expect([...iterValues(ws, { extent: 'values' })]).toEqual([
      ['a1', 1],
      ['a2', 2],
      ['a3', 3],
      ['a4', 4],
    ]);
  });

  it("defaults to extent: 'cells', which keeps the formatted tail", () => {
    const { ws } = makeFormattedPastItsData();
    const rows = [...iterValues(ws)];
    expect(rows.length).toBe(6);
    expect(rows.slice(4)).toEqual([
      [null, null],
      [null, null],
    ]);
  });

  it('keeps position i on column minCol + i, so a blank leading band still pads', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 3, 3, 'c3');
    setCell(ws, 9, 1, null, registerCellStyle(wb, { numberFormat: '0.00' }));
    // Rows 1..3 x cols 1..3: the value extent trims the styled row 9 but the
    // iteration still starts at A1, so 'c3' lands at index 2 of the third row.
    expect([...iterValues(ws, { extent: 'values' })]).toEqual([
      [null, null, null],
      [null, null, null],
      [null, null, 'c3'],
    ]);
  });

  it('an explicit maxRow still wins over the selected extent', () => {
    const { ws } = makeFormattedPastItsData();
    expect([...iterValues(ws, { extent: 'values', maxRow: 2 })]).toEqual([
      ['a1', 1],
      ['a2', 2],
    ]);
  });

  it('yields nothing for a sheet whose cells are all formatting', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const shaded = registerCellStyle(wb, { numberFormat: '0.00' });
    setCell(ws, 1, 1, null, shaded);
    setCell(ws, 50, 4, null, shaded);
    expect([...iterValues(ws, { extent: 'values' })]).toEqual([]);
    expect([...iterValues(ws)].length).toBe(50);
  });
});

describe('formatting-only cells survive a save / load round trip', () => {
  it('a reloaded sheet reports the inflated extent and the value extent', async () => {
    const { wb } = makeFormattedPastItsData();
    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const ws = getSheetByIndex(reloaded, 0);
    if (ws === undefined) throw new Error('expected a worksheet at index 0');
    expect(getMaxRow(ws)).toBe(6);
    expect(getValueExtent(ws)).toEqual({ minRow: 1, maxRow: 4, minCol: 1, maxCol: 2 });
    expect([...iterValues(ws, { extent: 'values' })].length).toBe(4);
  });
});
