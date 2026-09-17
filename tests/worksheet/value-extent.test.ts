// getValueExtent against the shape that motivates it: a sheet formatted past
// its content, which every reading that counts cell objects reports as taller
// and wider than its data.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { registerCellStyle } from '../../src/styles/cell-style.js';
import { addWorksheet, createWorkbook, getSheetByIndex } from '../../src/workbook/workbook.js';
import {
  ensureCell,
  getCellExtent,
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

  it('getCellExtent still counts the formatting-only cells', () => {
    const { ws } = makeFormattedPastItsData();
    expect(getCellExtent(ws)).toEqual({ minRow: 1, maxRow: 6, minCol: 1, maxCol: 2 });
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

  it('starts at the first cell holding a value, so a leading blank band is outside the box', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, null, registerCellStyle(wb, { numberFormat: '0.00' }));
    setCell(ws, 7, 3, 'only value');
    expect(getValueExtent(ws)).toEqual({ minRow: 7, maxRow: 7, minCol: 3, maxCol: 3 });
  });

  it('counts the empty string as no value, the shape a CSV converter leaves behind', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'header');
    for (let r = 2; r <= 200; r++) setCell(ws, r, 1, '');
    expect(getValueExtent(ws)).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
    expect(getCellExtent(ws)?.maxRow).toBe(200);
  });

  it('excludes a cell whose only payload is a hyperlink or a comment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    ensureCell(ws, 4, 1).hyperlinkId = 0;
    ensureCell(ws, 1, 6).commentId = 0;
    expect(getValueExtent(ws)).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
    expect(getCellExtent(ws)).toEqual({ minRow: 1, maxRow: 4, minCol: 1, maxCol: 6 });
  });

  it('returns undefined when the sheet holds formatting but no value', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const shaded = registerCellStyle(wb, { numberFormat: '0.00' });
    setCell(ws, 1, 1, null, shaded);
    setCell(ws, 20, 5, null, shaded);
    expect(getValueExtent(ws)).toBeUndefined();
    expect(getCellExtent(ws)).toEqual({ minRow: 1, maxRow: 20, minCol: 1, maxCol: 5 });
  });

  it('returns undefined for a sheet with no cells at all', () => {
    const wb = createWorkbook();
    expect(getValueExtent(addWorksheet(wb, 'S'))).toBeUndefined();
  });
});

describe('iterValues bounded by the value extent', () => {
  it('yields one row per row of data, with no phantom tail', () => {
    const { ws } = makeFormattedPastItsData();
    const box = getValueExtent(ws);
    if (box === undefined) throw new Error('expected a value extent');
    expect([...iterValues(ws, box)]).toEqual([
      ['a1', 1],
      ['a2', 2],
      ['a3', 3],
      ['a4', 4],
    ]);
  });

  it('the default bound keeps the formatted tail', () => {
    const { ws } = makeFormattedPastItsData();
    const rows = [...iterValues(ws)];
    expect(rows.length).toBe(6);
    expect(rows.slice(4)).toEqual([
      [null, null],
      [null, null],
    ]);
  });

  it('trims the leading blank band as well, so the first yielded row holds data', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 3, 3, 'c3');
    setCell(ws, 3, 4, 'd3');
    setCell(ws, 9, 1, null, registerCellStyle(wb, { numberFormat: '0.00' }));
    const box = getValueExtent(ws);
    if (box === undefined) throw new Error('expected a value extent');
    expect([...iterValues(ws, box)]).toEqual([['c3', 'd3']]);
    // Unbounded, the same sheet pads out to the styled row 9 and column A.
    expect([...iterValues(ws)].length).toBe(9);
  });

  it('an explicit maxRow narrows the box further', () => {
    const { ws } = makeFormattedPastItsData();
    const box = getValueExtent(ws);
    if (box === undefined) throw new Error('expected a value extent');
    expect([...iterValues(ws, { ...box, maxRow: 2 })]).toEqual([
      ['a1', 1],
      ['a2', 2],
    ]);
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
  });

  it('a reloaded empty-string cell is still outside the value extent', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'header');
    setCell(ws, 2, 1, '');
    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const sheet = getSheetByIndex(reloaded, 0);
    if (sheet === undefined) throw new Error('expected a worksheet at index 0');
    expect(getValueExtent(sheet)).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
  });
});
