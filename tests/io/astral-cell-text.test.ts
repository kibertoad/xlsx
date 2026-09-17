// Astral-plane text (emoji, CJK Extension B, musical notation, …) reached the
// sharedStrings part as `_xD83D__xDE00_`: `escapeCellString` matched UTF-16
// code units, so both halves of a surrogate pair looked like illegal
// characters. Our own reader undid the escape, which is why a round-trip test
// never caught it, but Excel, LibreOffice and Sheets all render the escape
// text verbatim. These assertions look at the emitted part, not at what we can
// read back.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const decoder = new TextDecoder();
const EMOJI = 'hi \u{1F600} there';

const partText = (bytes: Uint8Array, path: string): string => {
  const entry = unzipSync(bytes)[path];
  if (!entry) throw new Error(`no ${path} in the package`);
  return decoder.decode(entry);
};

const sheetOf = (wb: Awaited<ReturnType<typeof loadWorkbook>>): Worksheet => {
  const sheet = wb.sheets[0]?.sheet;
  if (!sheet || !('rows' in sheet)) throw new Error('expected a worksheet');
  return sheet;
};

describe('astral characters in cell text', () => {
  it('writes a shared string with the character intact', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, EMOJI);
    const bytes = await workbookToBytes(wb);

    const sst = partText(bytes, 'xl/sharedStrings.xml');
    expect(sst).toContain(`<t>${EMOJI}</t>`);
    expect(sst).not.toContain('_xD83D_');

    expect(getCell(sheetOf(await loadWorkbook(fromBuffer(bytes))), 1, 1)?.value).toBe(EMOJI);
  });

  it('writes rich-text runs with the character intact', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, {
      kind: 'rich-text',
      runs: makeRichText([{ text: EMOJI }]),
    });
    const bytes = await workbookToBytes(wb);
    expect(partText(bytes, 'xl/sharedStrings.xml')).toContain(EMOJI);
  });

  it('writes formula text and its cached string result with the character intact', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, {
      kind: 'formula',
      t: 'normal',
      formula: `CONCAT("${EMOJI}")`,
      cachedValue: EMOJI,
    });
    const bytes = await workbookToBytes(wb);
    const sheet = partText(bytes, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain(`<f>CONCAT("${EMOJI}")</f>`);
    expect(sheet).toContain(`<v>${EMOJI}</v>`);
  });

  it('reads back the split form older versions of this writer produced', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, 'placeholder');
    const bytes = await workbookToBytes(wb);
    const archive = unzipSync(bytes);
    archive['xl/sharedStrings.xml'] = new TextEncoder().encode(
      partText(bytes, 'xl/sharedStrings.xml').replace('placeholder', 'hi _xD83D__xDE00_ there'),
    );
    const { zipSync } = await import('fflate');

    const reloaded = await loadWorkbook(fromBuffer(zipSync(archive)));
    expect(getCell(sheetOf(reloaded), 1, 1)?.value).toBe(EMOJI);
  });
});
