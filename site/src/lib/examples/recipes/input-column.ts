// A column the recipient is meant to fill in: Excel's "Input" style marks it
// as editable, and a decimal validation keeps what they type usable.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { applyBuiltinStyle } from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { addDataValidation, appendRows, ensureCell, makeDataValidation } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Quote');

// C2 carries a rate that is already agreed; C3 is the one to fill in.
appendRows(ws, [
  ['Language', 'Words', 'Rate per word'],
  ['de', 71_579, 0.12],
  ['fr', 12_004],
]);

// ensureCell styles both rows the same way: it allocates the blank C3 and
// hands back C2 with its 0.12 intact. setCell would need the existing value
// passed back in to avoid wiping it.
for (let row = 2; row <= 3; row++) {
  applyBuiltinStyle(wb, ensureCell(ws, row, 3), 'Input');
}

addDataValidation(
  ws,
  makeDataValidation({
    type: 'decimal',
    operator: 'between',
    sqref: 'C2:C3',
    formula1: '0',
    formula2: '10',
    // Both flags default to false in ECMA-376: leave them out and Excel shows
    // neither the prompt nor the error, and takes any entry.
    showInputMessage: true,
    prompt: 'Rate per word, 0 to 10',
    showErrorMessage: true,
    errorTitle: 'Out of range',
    error: 'Enter a rate between 0 and 10.',
  }),
);

await saveWorkbook(wb, toFile('quote.xlsx'));
