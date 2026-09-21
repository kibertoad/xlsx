// Smoke tests for the docs/migrate-from-sheetjs.md API map. The map is the
// first thing a porting user reads, so every name in it has to exist with the
// signature the table implies. A rename now fails CI instead of stranding the
// guide.

import { describe, expect, it } from 'vitest';

describe('migrate-from-sheetjs: API map', () => {
  it('the reading and writing rows', async () => {
    const io = await import('../../src/io/index.js');
    const node = await import('../../src/node.js');
    expect(typeof io.loadWorkbook).toBe('function');
    expect(typeof io.saveWorkbook).toBe('function');
    expect(typeof io.workbookToBytes).toBe('function');
    expect(typeof io.fromBlob).toBe('function');
    expect(typeof io.fromResponse).toBe('function');
    expect(typeof io.fromStream).toBe('function');
    expect(typeof node.fromBuffer).toBe('function');
    expect(typeof node.fromFile).toBe('function');
    expect(typeof node.toFile).toBe('function');
  });

  it('the workbook and worksheet rows', async () => {
    const workbook = await import('../../src/workbook/index.js');
    const worksheet = await import('../../src/worksheet/index.js');
    expect(typeof workbook.createWorkbook).toBe('function');
    expect(typeof workbook.addWorksheet).toBe('function');
    expect(typeof workbook.sheetNames).toBe('function');
    expect(typeof workbook.getSheet).toBe('function');
    expect(typeof worksheet.ensureCellByCoord).toBe('function');
    expect(typeof worksheet.getCellByCoord).toBe('function');
    expect(typeof worksheet.iterRows).toBe('function');
    expect(typeof worksheet.iterValues).toBe('function');
    expect(typeof worksheet.appendRows).toBe('function');
    expect(typeof worksheet.getMergedCells).toBe('function');
    expect(typeof worksheet.setColumnWidth).toBe('function');
    expect(typeof worksheet.setFreezePanes).toBe('function');
    expect(typeof worksheet.isWorksheetEmpty).toBe('function');
  });

  it('the coordinate rows', async () => {
    const utils = await import('../../src/utils/index.js');
    expect(utils.tupleToCoordinate(2, 3)).toBe('B3');
    expect(utils.coordinateToTuple('B3')).toEqual({ col: 2, row: 3 });
    expect(typeof utils.rangeBoundaries).toBe('function');
  });

  it('the `cell.w`, `cell.f` and `cell.z` rows', async () => {
    const cell = await import('../../src/cell/index.js');
    const styles = await import('../../src/styles/index.js');
    expect(typeof cell.getFormulaText).toBe('function');
    expect(typeof cell.cellValueAsString).toBe('function');
    expect(typeof styles.getCellDisplayText).toBe('function');
    expect(typeof styles.getCellNumberFormat).toBe('function');
    expect(typeof styles.getCellDate).toBe('function');
  });

  it('the `raw: false` walk the guide spells out', async () => {
    const styles = await import('../../src/styles/index.js');
    const workbook = await import('../../src/workbook/index.js');
    const worksheet = await import('../../src/worksheet/index.js');

    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Data');
    worksheet.setCell(ws, 1, 1, 'Share');
    styles.setCellNumberFormat(wb, worksheet.setCell(ws, 1, 2, 0.5), '0.0%');

    const text = [...worksheet.iterRows(ws)].map((row) =>
      row.map((c) => (c === undefined ? '' : styles.getCellDisplayText(wb, c))),
    );
    expect(text).toEqual([['Share', '50.0%']]);
  });

  it('the formatting-only cell walk the guide spells out', async () => {
    const cell = await import('../../src/cell/index.js');
    const styles = await import('../../src/styles/index.js');
    const workbook = await import('../../src/workbook/index.js');
    const worksheet = await import('../../src/worksheet/index.js');

    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Data');
    worksheet.setCell(ws, 1, 1, 'a');
    worksheet.setCell(ws, 2, 1, null, styles.registerCellStyle(wb, { numberFormat: '#,##0' }));

    const kept: unknown[] = [];
    for (const c of worksheet.iterCells(ws)) {
      if (cell.isEmptyCell(c)) continue;
      kept.push(c.value);
    }
    expect(kept).toEqual(['a']);

    const box = worksheet.getValueExtent(ws);
    expect(box).toEqual({ minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 });
    if (box === undefined) throw new Error('expected a value extent');
    expect([...worksheet.iterValues(ws, box)]).toEqual([['a']]);
  });

  it('the cellDates replacement the guide points at', async () => {
    const styles = await import('../../src/styles/index.js');
    const workbook = await import('../../src/workbook/index.js');
    const worksheet = await import('../../src/worksheet/index.js');

    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Data');
    const cell = worksheet.setCell(ws, 1, 1, 45_365);
    styles.setCellNumberFormat(wb, cell, 'yyyy-mm-dd');
    expect(styles.getCellDate(wb, cell)?.toISOString()).toBe('2024-03-14T00:00:00.000Z');
  });
});
