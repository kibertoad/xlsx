// Tests for getCellExtent / getCellExtentRef bounding-box helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  getCellExtent,
  getCellExtentRef,
  setCell,
} from '../../src/worksheet/worksheet.js';

describe('getCellExtent', () => {
  it('returns undefined for an empty sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getCellExtent(ws)).toBeUndefined();
  });

  it('1×1 sheet → {1,1,1,1}', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect(getCellExtent(ws)).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
  });

  it('walks both axes correctly with sparse data', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 3, 5, 'tl');
    setCell(ws, 7, 9, 'br');
    expect(getCellExtent(ws)).toEqual({ minRow: 3, maxRow: 7, minCol: 5, maxCol: 9 });
  });

  it('a single populated cell anywhere → that one cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 100, 50, 'lonely');
    expect(getCellExtent(ws)).toEqual({ minRow: 100, maxRow: 100, minCol: 50, maxCol: 50 });
  });

  it('ignores empty row maps (deleted-then-emptied rows)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 5, 'x');
    // Force an empty row map at row 99 to simulate post-delete state.
    ws.rows.set(99, new Map());
    expect(getCellExtent(ws)).toEqual({ minRow: 5, maxRow: 5, minCol: 5, maxCol: 5 });
  });
});

describe('getCellExtentRef', () => {
  it('returns canonical A1 range string', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    setCell(ws, 5, 3, 'br');
    expect(getCellExtentRef(ws)).toBe('A1:C5');
  });

  it('1×1 → just "A1" (no colon)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 4, 7, 'x');
    expect(getCellExtentRef(ws)).toBe('G4');
  });

  it('empty sheet → undefined', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getCellExtentRef(ws)).toBeUndefined();
  });
});
