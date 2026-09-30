// Every assertion here reads the emitted part rather than the reloaded model.
// `unescapeCellString` rebuilds a split surrogate pair on the way in, so a
// load-save-load comparison reports success whatever the file on disk says.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { expectSheet, partText, SHEET_PART, SST_PART } from './_helpers.js';

const encoder = new TextEncoder();
const EMOJI = 'hi \u{1F600} there';
const SPLIT = 'hi _xD83D__xDE00_ there';

describe('astral characters in cell text', () => {
  it('writes a shared string with the character intact', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, EMOJI);
    const bytes = await workbookToBytes(wb);

    const sst = partText(unzipSync(bytes), SST_PART);
    expect(sst).toContain(`<t>${EMOJI}</t>`);
    expect(sst).not.toContain('_xD83D_');

    const reloaded = await loadWorkbook(fromBuffer(bytes));
    expect(getCell(expectSheet(reloaded.sheets[0]?.sheet), 1, 1)?.value).toBe(EMOJI);
  });

  it('writes rich-text runs with the character intact', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, {
      kind: 'rich-text',
      runs: makeRichText([{ text: EMOJI }]),
    });
    const bytes = await workbookToBytes(wb);
    expect(partText(unzipSync(bytes), SST_PART)).toContain(EMOJI);
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
    const sheet = partText(unzipSync(bytes), SHEET_PART);
    expect(sheet).toContain(`<f>CONCAT("${EMOJI}")</f>`);
    expect(sheet).toContain(`<v>${EMOJI}</v>`);
  });

  it('rebuilds the split form from a sharedStrings part an earlier version wrote', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, 'placeholder');
    const archive = unzipSync(await workbookToBytes(wb));
    archive[SST_PART] = encoder.encode(partText(archive, SST_PART).replace('placeholder', SPLIT));

    const reloaded = await loadWorkbook(fromBuffer(zipSync(archive)));
    expect(getCell(expectSheet(reloaded.sheets[0]?.sheet), 1, 1)?.value).toBe(EMOJI);
  });

  it('hands back the split form in formula text verbatim, since no reader unescapes <f>', async () => {
    // The `_xHHHH_` convention only round-trips where the reader inverts it,
    // which is the shared-string and inline-string path. Formula text and a
    // cached string result reach the caller as written, so an emoji an earlier
    // version split there stays split. Pinned so the next changelog entry
    // about this does not overstate what reading such a file recovers.
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, {
      kind: 'formula',
      t: 'normal',
      formula: 'PLACEHOLDER_F',
      cachedValue: 'PLACEHOLDER_V',
    });
    const archive = unzipSync(await workbookToBytes(wb));
    archive[SHEET_PART] = encoder.encode(
      partText(archive, SHEET_PART)
        .replace('PLACEHOLDER_F', `CONCAT("${SPLIT}")`)
        .replace('PLACEHOLDER_V', SPLIT),
    );

    const reloaded = await loadWorkbook(fromBuffer(zipSync(archive)));
    const value = getCell(expectSheet(reloaded.sheets[0]?.sheet), 1, 1)?.value;
    expect(value).toEqual({
      kind: 'formula',
      t: 'normal',
      formula: `CONCAT("${SPLIT}")`,
      cachedValue: SPLIT,
    });
  });
});
