// A cell whose text already looks like the `_xHHHH_` escape convention has to
// survive a save and a load. The writer protects the sequence by escaping its
// leading underscore; two sequences can share that underscore, and the reader
// decodes left to right without overlapping, so the protection has to leave
// every opening underscore in place rather than consume the one it matched.

import { type Unzipped, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const decoder = new TextDecoder();
const SST_PART = 'xl/sharedStrings.xml';

// Sequences that share an underscore, a lone sequence, and a literal
// `_x005F_` that itself opens one.
const SAMPLES = [
  '_x0041_x0042_',
  'a_x0044_x0045_x0046_b',
  '_x005F_x0041_',
  'foo_x0041_bar',
  '_x005F_',
  '_x0041__x0042_',
];

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected worksheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

const partText = (archive: Unzipped, path: string): string => {
  const entry = archive[path];
  if (!entry) throw new Error(`no ${path} in the package`);
  return decoder.decode(entry);
};

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
      runs: makeRichText([{ text: '_x0041_x0042_', font: { b: true } }, { text: '_x005F_x0041_' }]),
    });

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    expect(getCell(expectSheet(reloaded.sheets[0]?.sheet), 1, 1)?.value).toEqual({
      kind: 'rich-text',
      runs: [{ text: '_x0041_x0042_', font: { b: true } }, { text: '_x005F_x0041_' }],
    });
  });

  it('leaves no decodable sequence in the emitted part', async () => {
    // The reload assertions above would also pass if the writer emitted the
    // text verbatim and the reader happened to leave it alone. Read the part
    // to pin that every opening underscore really was escaped, which is what
    // makes Excel agree with us about the cell.
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, '_x0041_x0042_');
    const sst = partText(unzipSync(await workbookToBytes(wb)), SST_PART);
    expect(sst).toContain('<t>_x005F_x0041_x005F_x0042_</t>');
  });
});
