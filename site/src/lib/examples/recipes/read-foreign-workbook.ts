// Read a workbook somebody else produced: take the bytes as a Blob, find the
// first worksheet that has anything in it, and turn its rows into the text a
// person would see in Excel.

import { fromBlob, loadWorkbook } from '@office-kit/xlsx/io';
import { getCellDate, getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Workbook } from '@office-kit/xlsx/workbook';
import { isWorksheetEmpty, iterRows, type Worksheet } from '@office-kit/xlsx/worksheet';

export async function readSheetAsText(file: Blob): Promise<string[][]> {
  const wb = await loadWorkbook(fromBlob(file));

  // wb.sheets mixes worksheets and chartsheets, and a producer often leaves a
  // blank sheet in front of the data.
  const first = wb.sheets.find((s) => s.kind === 'worksheet' && !isWorksheetEmpty(s.sheet));
  if (first?.kind !== 'worksheet') return [];

  return [...iterRows(first.sheet)].map((row) =>
    row.map((cell) => (cell === undefined ? '' : getCellDisplayText(wb, cell))),
  );
}

// `getCellDisplayText` gives you the text, which is what a CSV or an HTML table
// wants. When you need the value, ask for the value: a date cell holds a plain
// serial number, and its number format is the only thing that says so.
export function readDueDates(wb: Workbook, ws: Worksheet, column: number): Date[] {
  const due: Date[] = [];
  for (const row of iterRows(ws, { minCol: column, maxCol: column })) {
    const cell = row[0];
    if (cell === undefined) continue;
    const date = getCellDate(wb, cell);
    if (date !== undefined) due.push(date);
  }
  return due;
}
