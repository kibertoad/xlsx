// A cell whose text already looks like the `_xHHHH_` escape convention has to
// survive a save and a load on every carrier that decodes the convention: a
// shared string, a rich-text run and an inline string. The writer escapes the
// underscore that opens such a sequence; readers decode left to right without
// overlapping, so the character that closes one sequence has to stay available
// to open the next.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { expectSheet, partText, SHEET_PART, SST_PART } from './_helpers.js';

// Sequences that share an underscore, a lone sequence, a literal `_x005F_`
// that itself opens one, and sequences closed by a character the writer
// escapes in turn: `SKU_x0041` + `\n` becomes `SKU_x0041_x000A_` unless the
// leading underscore is protected as well.
const SAMPLES = [
  '_x0041_x0042_',
  'a_x0044_x0045_x0046_b',
  '_x005F_x0041_',
  'foo_x0041_bar',
  '_x005F_',
  '_x0041__x0042_',
  'SKU_x0041\nrest',
  'col_x0009\tval',
  'note_x000D\rend',
];

// A write-only workbook keeps strings in the shared table until 8 MB of
// payload, then emits every later value inline. 66 rows of 32,000 characters
// clear that budget; the `t="inlineStr"` assertions below fail if the
// accounting ever changes and the samples land in the table instead.
const ROWS_TO_FILL_STRING_TABLE = 66;

describe('cell text that already looks escaped', () => {
  it('round-trips shared strings through a save and a load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    SAMPLES.forEach((text, i) => setCell(ws, i + 1, 1, text));

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const sheet = expectSheet(reloaded.sheets[0]?.sheet);
    SAMPLES.forEach((text, i) => {
      expect(getCell(sheet, i + 1, 1)?.value).toBe(text);
    });
  });

  it('round-trips rich-text runs', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, {
      kind: 'rich-text',
      runs: makeRichText([
        { text: '_x0041_x0042_', font: { b: true } },
        { text: '_x005F_x0041_' },
        { text: 'SKU_x0041\nrest' },
      ]),
    });

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    expect(getCell(expectSheet(reloaded.sheets[0]?.sheet), 1, 1)?.value).toEqual({
      kind: 'rich-text',
      runs: [
        { text: '_x0041_x0042_', font: { b: true } },
        { text: '_x005F_x0041_' },
        { text: 'SKU_x0041\nrest' },
      ],
    });
  });

  it('escapes every opening underscore in the emitted part', async () => {
    // The reload assertions above would also pass if the writer emitted the
    // text verbatim and the reader happened to leave it alone. Read the part
    // to pin the stored form, which is what makes Excel agree with us about
    // the cell.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, '_x0041_x0042_');
    setCell(ws, 2, 1, 'SKU_x0041\nrest');

    const sst = partText(unzipSync(await workbookToBytes(wb)), SST_PART);
    expect(sst).toContain('<t>_x005F_x0041_x005F_x0042_</t>');
    expect(sst).toContain('<t xml:space="preserve">SKU_x005F_x0041_x000A_rest</t>');
  });

  it('round-trips inline strings, which the streaming writer emits once the table is full', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink, { compressionLevel: 0 });
    const ws = await wb.addWorksheet('S');
    for (let i = 0; i < ROWS_TO_FILL_STRING_TABLE; i++) await ws.appendRow([`${i}`.padEnd(32_000, 'x')]);
    await ws.appendRow(SAMPLES);
    await ws.close();
    await wb.finalize();
    const bytes = sink.result();

    const sheet = partText(unzipSync(bytes), SHEET_PART);
    expect(sheet).toContain('<is><t>_x005F_x0041_x005F_x0042_</t></is>');

    const streamed = await loadWorkbookStream(fromBuffer(bytes));
    try {
      const rows = [];
      for await (const row of streamed.openWorksheet('S').iterValues()) rows.push(row);
      expect(rows[ROWS_TO_FILL_STRING_TABLE]).toEqual(SAMPLES);
    } finally {
      await streamed.close();
    }

    const reloaded = await loadWorkbook(fromBuffer(bytes));
    const loadedSheet = expectSheet(reloaded.sheets[0]?.sheet);
    SAMPLES.forEach((text, i) => {
      expect(getCell(loadedSheet, ROWS_TO_FILL_STRING_TABLE + 1, i + 1)?.value).toBe(text);
    });
  });
});
