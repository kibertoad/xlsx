// A `<col width>` / `<row ht>` is an `xsd:double` in the part. `NaN` and
// `Infinity` have no lexical form there, so a non-finite value reached Excel as
// `width="NaN"`: a worksheet no schema validator accepts and Excel offers to
// repair. Both writers now refuse the value instead of emitting it.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { OpenXmlIoError, OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  setCell,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  setRowDimension,
  setRowHeight,
} from '../../src/worksheet/worksheet.js';

const sheet = () => {
  const wb = createWorkbook();
  return addWorksheet(wb, 'S');
};

const BAD_SIZES = [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1];

describe('column width validation', () => {
  it('rejects a non-finite or negative width', () => {
    for (const bad of BAD_SIZES) {
      expect(() => setColumnWidth(sheet(), 1, bad)).toThrow(OpenXmlSchemaError);
    }
  });

  it('rejects one through the low-level setColumnDimension too', () => {
    expect(() => setColumnDimension(sheet(), 1, { width: Number.NaN })).toThrow(OpenXmlSchemaError);
  });

  it('accepts zero and a fractional width', () => {
    expect(setColumnWidth(sheet(), 1, 0).width).toBe(0);
    expect(setColumnWidth(sheet(), 1, 12.5).width).toBe(12.5);
  });

  it('leaves a width past Excel ceiling alone, since Excel clamps it on open', () => {
    // 255 is Excel's widest column. A larger value is still a well-formed
    // double, so refusing it would refuse a file that opens fine.
    expect(setColumnWidth(sheet(), 1, 1000).width).toBe(1000);
  });

  it('setColumnWidths still skips a non-numeric entry rather than throwing', () => {
    // Its documented contract: entries that are not finite numbers are ignored,
    // which is what lets a caller pass a sparse array.
    const ws = sheet();
    setColumnWidths(ws, [10, Number.NaN, 12]);
    expect(ws.columnDimensions.get(1)?.width).toBe(10);
    expect(ws.columnDimensions.get(2)).toBeUndefined();
    expect(ws.columnDimensions.get(3)?.width).toBe(12);
  });
});

describe('row height validation', () => {
  it('rejects a non-finite or negative height', () => {
    for (const bad of BAD_SIZES) {
      expect(() => setRowHeight(sheet(), 1, bad)).toThrow(OpenXmlSchemaError);
    }
  });

  it('rejects one through the low-level setRowDimension too', () => {
    expect(() => setRowDimension(sheet(), 1, { height: Number.NaN })).toThrow(OpenXmlSchemaError);
  });

  it('accepts zero and a fractional height', () => {
    expect(setRowHeight(sheet(), 1, 0).height).toBe(0);
    expect(setRowHeight(sheet(), 1, 24.75).height).toBe(24.75);
  });
});

describe('the emitted part', () => {
  it('carries no NaN once the setters refuse one', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'x');
    setColumnWidth(ws, 1, 12);
    setRowHeight(ws, 1, 24);
    const part = unzipSync(await workbookToBytes(wb))['xl/worksheets/sheet1.xml'];
    if (!part) throw new Error('no worksheet part');
    expect(new TextDecoder().decode(part)).not.toContain('NaN');
  });
});

describe('the write-only path', () => {
  const collectingSink = () => {
    const chunks: Uint8Array[] = [];
    return {
      sink: {
        toBytes: () => ({
          write: (c: Uint8Array) => {
            chunks.push(c);
          },
          finish: () => new Uint8Array(0),
        }),
      },
      chunks,
    };
  };

  it('rejects a non-finite width, which used to reach the part as width="NaN"', async () => {
    const { sink } = collectingSink();
    const wb = await createWriteOnlyWorkbook(sink as never);
    const ws = await wb.addWorksheet('S');
    expect(() => ws.setColumnWidth(1, Number.NaN)).toThrow(OpenXmlIoError);
    wb.abort();
  });

  it('rejects an off-grid column index, which used to reach the part as min="0"', async () => {
    const { sink } = collectingSink();
    const wb = await createWriteOnlyWorkbook(sink as never);
    const ws = await wb.addWorksheet('S');
    // `<col min="0">` and a column past XFD are both outside the grid; the
    // modelled writer has always rejected them and this path did not.
    expect(() => ws.setColumnWidth(0, 12)).toThrow(OpenXmlIoError);
    expect(() => ws.setColumnWidth(16_385, 12)).toThrow(OpenXmlIoError);
    expect(() => ws.setColumnWidth(1.5, 12)).toThrow(OpenXmlIoError);
    wb.abort();
  });

  it('still accepts a valid column width', async () => {
    const { sink } = collectingSink();
    const wb = await createWriteOnlyWorkbook(sink as never);
    const ws = await wb.addWorksheet('S');
    expect(() => ws.setColumnWidth(16_384, 12)).not.toThrow();
    wb.abort();
  });
});
