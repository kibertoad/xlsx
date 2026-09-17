// `loadWorkbookStream` answers a corrupt numeric cell the same way
// `loadWorkbook` does. Yielding `null` instead would be indistinguishable from
// an empty cell, and the two entry points would disagree about the same bytes.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream, type ReadOnlyWorksheet } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';

const SHEET_PART = 'xl/worksheets/sheet1.xml';

/** Four rows of numbers, with row `at`'s value replaced by `text`. */
const savedWithCellText = async (text: string, at = 1): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  for (let row = 1; row <= 4; row++) setCell(ws, row, 1, row * 10);
  const entries = unzipSync(await workbookToBytes(wb));
  const part = entries[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  const patched = new TextDecoder().decode(part).replace(`<v>${at * 10}</v>`, `<v>${text}</v>`);
  entries[SHEET_PART] = new TextEncoder().encode(patched);
  return zipSync(entries);
};

const drain = async (ws: ReadOnlyWorksheet, minRow?: number): Promise<unknown[][]> => {
  const rows: unknown[][] = [];
  for await (const row of ws.iterValues(minRow === undefined ? {} : { minRow })) rows.push(row);
  return rows;
};

describe('loadWorkbookStream on a numeric cell that is not a finite number', () => {
  it('throws, naming the sheet and the cell', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await savedWithCellText('oops')));
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(
        'worksheet: <v>oops</v> at Data!A1 is not a finite number',
      );
    } finally {
      await wb.close();
    }
  });

  it('throws on the indexed band-query path too', async () => {
    // minRow > 1 replays from the row-offset index instead of streaming the
    // part, which is a second route into the row iterator.
    const wb = await loadWorkbookStream(fromBuffer(await savedWithCellText('1e400', 3)));
    try {
      await expect(drain(wb.openWorksheet('Data'), 3)).rejects.toThrow(
        'worksheet: <v>1e400</v> at Data!A3 is not a finite number',
      );
    } finally {
      await wb.close();
    }
  });

  it('agrees with loadWorkbook about the same bytes', async () => {
    const bytes = await savedWithCellText('oops');
    await expect(loadWorkbook(fromBuffer(bytes))).rejects.toThrow(OpenXmlSchemaError);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(OpenXmlSchemaError);
    } finally {
      await wb.close();
    }
  });

  it('still reads a blank value as an empty cell', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await savedWithCellText('\n    ')));
    try {
      expect(await drain(wb.openWorksheet('Data'))).toEqual([[null], [20], [30], [40]]);
    } finally {
      await wb.close();
    }
  });
});
