// A `<col width>` / `<row ht>` is an `xsd:double`. `NaN` and `Infinity` have no
// lexical form there, so `width="NaN"` reaches Excel as a part that fails
// schema validation and gets offered for repair. A size the caller supplies is
// refused at the setter; the serializer is the backstop for the dimension maps,
// which the public API lets a caller write into directly.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook, type WriteOnlyWorksheet } from '../../src/streaming/write-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getSheet, type Workbook } from '../../src/workbook/workbook.js';
import { makeColumnDimension, makeRowDimension } from '../../src/worksheet/dimensions.js';
import {
  autofitColumn,
  autofitColumns,
  collapseColumnGroup,
  getColumnDimension,
  getRowDimension,
  groupColumns,
  hideColumn,
  hideRow,
  setCell,
  setColumnWidth,
  setColumnWidths,
  setRowHeight,
  setRowHeights,
  unhideColumn,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

const sheet = (): Worksheet => addWorksheet(createWorkbook(), 'S');

const BAD_SIZES = [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1];

const worksheetPart = async (wb: Workbook): Promise<string> => {
  const part = unzipSync(await workbookToBytes(wb))['xl/worksheets/sheet1.xml'];
  if (!part) throw new Error('no worksheet part');
  return new TextDecoder().decode(part);
};

describe('a size the caller passes in', () => {
  it('is refused by setColumnWidth when it is non-finite or negative', () => {
    for (const bad of BAD_SIZES) {
      expect(() => setColumnWidth(sheet(), 1, bad)).toThrow(OpenXmlSchemaError);
    }
  });

  it('is refused by setRowHeight on the same rule', () => {
    for (const bad of BAD_SIZES) {
      expect(() => setRowHeight(sheet(), 1, bad)).toThrow(OpenXmlSchemaError);
    }
  });

  it('is accepted at zero and at a fraction', () => {
    expect(setColumnWidth(sheet(), 1, 0).width).toBe(0);
    expect(setColumnWidth(sheet(), 1, 12.5).width).toBe(12.5);
    expect(setRowHeight(sheet(), 1, 0).height).toBe(0);
    expect(setRowHeight(sheet(), 1, 24.75).height).toBe(24.75);
  });

  it('is accepted past Excel ceilings of 255 characters and 409 points', () => {
    // A value past them is a well-formed double that Excel clamps on open, so
    // refusing it would refuse a file that works.
    expect(setColumnWidth(sheet(), 1, 1000).width).toBe(1000);
    expect(setRowHeight(sheet(), 1, 500).height).toBe(500);
  });
});

describe('a size already in the model', () => {
  /** What a part carrying `width="-1"` and `ht="-3"` loads as. */
  const withNegativeSizes = async (): Promise<{ wb: Workbook; ws: Worksheet }> => {
    const source = createWorkbook();
    const authored = addWorksheet(source, 'S');
    setCell(authored, 1, 1, 'x');
    authored.columnDimensions.set(1, makeColumnDimension(1, { width: -1 }));
    authored.rowDimensions.set(1, makeRowDimension({ height: -3 }));
    const wb = await loadWorkbook(fromBuffer(await workbookToBytes(source)));
    const ws = getSheet(wb, 'S');
    if (!ws) throw new Error('no sheet');
    return { wb, ws };
  };

  it('survives the round-trip that produced it', async () => {
    const { ws } = await withNegativeSizes();
    expect(getColumnDimension(ws, 1)?.width).toBe(-1);
    expect(getRowDimension(ws, 1)?.height).toBe(-3);
  });

  it('does not block a mutator that leaves it alone', async () => {
    // Each of these re-submits the existing entry through
    // set{Column,Row}Dimension, so a guard on that funnel would reject a value
    // the caller never supplied.
    const { ws } = await withNegativeSizes();
    expect(() => hideColumn(ws, 1)).not.toThrow();
    expect(() => unhideColumn(ws, 1)).not.toThrow();
    expect(() => groupColumns(ws, 1, 1)).not.toThrow();
    expect(() => collapseColumnGroup(ws, 1, 1)).not.toThrow();
    expect(() => hideRow(ws, 1)).not.toThrow();
    expect(getColumnDimension(ws, 1)?.width).toBe(-1);
  });

  it('is written back rather than rejected at save', async () => {
    const { wb, ws } = await withNegativeSizes();
    hideColumn(ws, 1);
    const xml = await worksheetPart(wb);
    expect(xml).toContain('width="-1"');
    expect(xml).toContain('ht="-3"');
  });
});

describe('the bulk setters', () => {
  it('skip an unusable width and keep going', () => {
    const ws = sheet();
    setColumnWidths(ws, [10, Number.NaN, 12, -1, 14]);
    expect(ws.columnDimensions.get(1)?.width).toBe(10);
    expect(ws.columnDimensions.get(2)).toBeUndefined();
    expect(ws.columnDimensions.get(3)?.width).toBe(12);
    expect(ws.columnDimensions.get(4)).toBeUndefined();
    expect(ws.columnDimensions.get(5)?.width).toBe(14);
  });

  it('skip an unusable height and keep going', () => {
    const ws = sheet();
    setRowHeights(ws, [10, Number.NaN, 12, -1, 14]);
    expect(ws.rowDimensions.get(1)?.height).toBe(10);
    expect(ws.rowDimensions.get(2)).toBeUndefined();
    expect(ws.rowDimensions.get(3)?.height).toBe(12);
    expect(ws.rowDimensions.get(4)).toBeUndefined();
    expect(ws.rowDimensions.get(5)?.height).toBe(14);
  });

  it('skip an unusable entry in the record form too', () => {
    const ws = sheet();
    setColumnWidths(ws, { 1: 10, 2: -1, 3: 12 });
    setRowHeights(ws, { 1: 10, 2: Number.NEGATIVE_INFINITY, 3: 12 });
    expect(ws.columnDimensions.get(2)).toBeUndefined();
    expect(ws.columnDimensions.get(3)?.width).toBe(12);
    expect(ws.rowDimensions.get(2)).toBeUndefined();
    expect(ws.rowDimensions.get(3)?.height).toBe(12);
  });
});

describe('autofit', () => {
  const populated = (): Worksheet => {
    const ws = sheet();
    setCell(ws, 1, 1, 'first');
    setCell(ws, 1, 2, 'second');
    return ws;
  };

  it('names the option it was handed rather than the setter it calls', () => {
    const ws = populated();
    expect(() => autofitColumns(ws, { padding: Number.NaN })).toThrow(/autofitColumns: padding/);
    expect(() => autofitColumn(ws, 1, { max: Number.NaN })).toThrow(/autofitColumn: max/);
    expect(() => autofitColumns(ws, { min: -1 })).toThrow(OpenXmlSchemaError);
    expect(ws.columnDimensions.size).toBe(0);
  });

  it('still sizes a column when the options are usable', () => {
    const ws = populated();
    autofitColumns(ws, { padding: 2 });
    expect(ws.columnDimensions.get(1)?.width).toBeGreaterThan(0);
  });
});

describe('the emitted worksheet part', () => {
  it('carries the sizes the setters accepted', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'x');
    setColumnWidth(ws, 1, 12);
    setRowHeight(ws, 1, 24);
    const xml = await worksheetPart(wb);
    expect(xml).toContain('<col min="1" max="1" width="12"');
    expect(xml).toContain('ht="24"');
  });

  it('refuses a width written straight into columnDimensions', async () => {
    // setColumnDimension's docblock points callers at this map for
    // range-spanning entries, so the setter guard alone cannot keep NaN out.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'x');
    ws.columnDimensions.set(1, makeColumnDimension(1, { width: Number.NaN }));
    await expect(workbookToBytes(wb)).rejects.toThrow(/<col min="1" max="1"> width must be a finite number/);
  });

  it('refuses a height written straight into rowDimensions, naming the row', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 3, 1, 'x');
    ws.rowDimensions.set(3, makeRowDimension({ height: Number.POSITIVE_INFINITY }));
    await expect(workbookToBytes(wb)).rejects.toThrow(/<row r="3"> ht must be a finite number/);
  });
});

describe('the write-only worksheet', () => {
  const writeOnlySheet = async (): Promise<{ ws: WriteOnlyWorksheet; finish: () => Promise<string> }> => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const ws = await wb.addWorksheet('S');
    return {
      ws,
      finish: async () => {
        await ws.close();
        await wb.finalize();
        const part = unzipSync(sink.result())['xl/worksheets/sheet1.xml'];
        if (!part) throw new Error('no worksheet part');
        return new TextDecoder().decode(part);
      },
    };
  };

  it('refuses the sizes the modelled setter refuses', async () => {
    const { ws, finish } = await writeOnlySheet();
    for (const bad of BAD_SIZES) {
      expect(() => ws.setColumnWidth(1, bad)).toThrow(OpenXmlSchemaError);
    }
    await finish();
  });

  it('refuses a column index outside the grid', async () => {
    const { ws, finish } = await writeOnlySheet();
    for (const col of [0, 16_385, 1.5]) {
      expect(() => ws.setColumnWidth(col, 12)).toThrow(OpenXmlSchemaError);
    }
    await finish();
  });

  it('emits a usable width into the cols header', async () => {
    const { ws, finish } = await writeOnlySheet();
    ws.setColumnWidth(1, 12);
    ws.setColumnWidth(16_384, 20);
    await ws.appendRow(['x']);
    const xml = await finish();
    expect(xml).toContain('<col min="1" max="1" width="12" customWidth="1"/>');
    expect(xml).toContain('<col min="16384" max="16384" width="20" customWidth="1"/>');
  });

  it('throws the error type the modelled setter throws for the same value', async () => {
    // A caller writing through both writers gets one catch, not two.
    const { ws, finish } = await writeOnlySheet();
    expect(() => ws.setColumnWidth(1, Number.NaN)).toThrow(OpenXmlSchemaError);
    expect(() => setColumnWidth(sheet(), 1, Number.NaN)).toThrow(OpenXmlSchemaError);
    await finish();
  });
});
