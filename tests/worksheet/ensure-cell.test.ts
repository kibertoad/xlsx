// Tests for ensureCell and the mandatory-value contract on setCell.
//
// The pair exists because a styling / formula pass that walks already-populated
// rows must be able to reach a cell without wiping it.

import { describe, expect, it } from 'vitest';
import { getFormulaText, setFormula } from '../../src/cell/cell.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getCellFill, setCellBackgroundColor } from '../../src/styles/cell-style.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { deleteCell, ensureCell, getCell, setCell } from '../../src/worksheet/worksheet.js';

describe('ensureCell', () => {
  it('allocates a blank cell at an unpopulated coordinate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 3, 2);
    expect(c.row).toBe(3);
    expect(c.col).toBe(2);
    expect(c.value).toBeNull();
    expect(c.styleId).toBe(0);
    expect(getCell(ws, 3, 2)).toBe(c);
  });

  it('returns the existing cell with its value and styleId intact', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'keep me', 4);
    const c = ensureCell(ws, 1, 1);
    expect(c.value).toBe('keep me');
    expect(c.styleId).toBe(4);
  });

  it('leaves formulas alone where setCell would blank them', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFormula(ensureCell(ws, 1, 1), 'SUM(B1:D1)');
    expect(getFormulaText(ensureCell(ws, 1, 1))).toBe('SUM(B1:D1)');
    setCell(ws, 1, 1, null);
    expect(getCell(ws, 1, 1)?.value).toBeNull();
  });

  it('advances the append cursor so appendRow does not land on it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    ensureCell(ws, 5, 1);
    expect(ws._appendRowCursor).toBe(5);
  });
});

describe('blanking a cell on purpose', () => {
  it('null clears the value and keeps the cell and its formatting', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'draft');
    setCellBackgroundColor(wb, c, 'FFFFFF00');
    const styled = c.styleId;
    expect(styled).not.toBe(0);

    setCell(ws, 1, 1, null);
    expect(getCell(ws, 1, 1)?.value).toBeNull();
    expect(getCell(ws, 1, 1)?.styleId).toBe(styled);
  });

  it('a blanked cell survives a save / load round-trip with its fill', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Blank');
    setCellBackgroundColor(wb, setCell(ws, 2, 3, 'draft'), 'FFFFFF00');
    setCell(ws, 2, 3, null);

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const sheet = getSheet(reloaded, 'Blank');
    expect(sheet).toBeDefined();
    if (!sheet) return;
    const cell = getCell(sheet, 2, 3);
    expect(cell).toBeDefined();
    if (!cell) return;
    expect(cell.value).toBeNull();
    const fill = getCellFill(reloaded, cell);
    expect(fill.kind === 'pattern' && fill.fgColor?.rgb).toBe('FFFFFF00');
  });

  it('deleteCell is the other meaning: no cell, no formatting', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCellBackgroundColor(wb, setCell(ws, 1, 1, 'draft'), 'FFFFFF00');
    deleteCell(ws, 1, 1);
    expect(getCell(ws, 1, 1)).toBeUndefined();
  });
});

describe('setCell with an explicit value', () => {
  it('replaces the value and keeps the styleId unless overridden', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'first', 3);
    setCell(ws, 1, 1, 'second');
    expect(getCell(ws, 1, 1)?.value).toBe('second');
    expect(getCell(ws, 1, 1)?.styleId).toBe(3);
    setCell(ws, 1, 1, 'third', 7);
    expect(getCell(ws, 1, 1)?.styleId).toBe(7);
  });
});
