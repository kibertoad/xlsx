// Excel accepts 32_767 UTF-16 code units in one cell. A longer string used to
// be written without complaint, and Excel then reported the workbook as
// needing repair. The writer refuses it instead, so the failure lands where
// the caller can still do something about it.

import { type Unzipped, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
// Through the barrel the subpath publishes, so a dropped re-export fails here
// rather than only in a consumer's build.
import { MAX_CELL_TEXT_LENGTH } from '../../src/utils/index.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const AT_LIMIT = 'x'.repeat(MAX_CELL_TEXT_LENGTH);
const OVER_LIMIT = 'x'.repeat(MAX_CELL_TEXT_LENGTH + 1);
/** One character, two UTF-16 code units. */
const EMOJI = '\u{1F600}';
const SHEET_PART = 'xl/worksheets/sheet1.xml';
const SST_PART = 'xl/sharedStrings.xml';
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const save = async (value: Parameters<typeof setCell>[3]): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'S'), 1, 1, value);
  return workbookToBytes(wb);
};

const partText = (archive: Unzipped, path: string): string => {
  const entry = archive[path];
  if (!entry) throw new Error(`no ${path} in the package`);
  return decoder.decode(entry);
};

const expectSheet = (
  sheet: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined,
): Worksheet => {
  if (!sheet || !('rows' in sheet)) throw new Error('expected a worksheet');
  return sheet;
};

describe('cell text length', () => {
  it('accepts a string exactly at the limit', async () => {
    await expect(save(AT_LIMIT)).resolves.toBeInstanceOf(Uint8Array);
  });

  it('rejects a string one code unit past it', async () => {
    await expect(save(OVER_LIMIT)).rejects.toThrow(OpenXmlSchemaError);
  });

  it('names the cell and both lengths', async () => {
    await expect(save(OVER_LIMIT)).rejects.toThrow(/A1 is 32768 UTF-16 code units, past the 32767/);
  });

  it('counts a rich-text cell across its runs, not per run', async () => {
    // Each run is comfortably inside the limit; the cell they make is one unit
    // past it.
    const half = 'y'.repeat(Math.ceil((MAX_CELL_TEXT_LENGTH + 1) / 2));
    expect(half.length * 2).toBe(MAX_CELL_TEXT_LENGTH + 1);
    await expect(
      save({ kind: 'rich-text', runs: makeRichText([{ text: half }, { text: half }]) }),
    ).rejects.toThrow(OpenXmlSchemaError);
  });

  it('accepts rich-text runs that together sit at the limit', async () => {
    const most = 'y'.repeat(MAX_CELL_TEXT_LENGTH - 1);
    await expect(
      save({ kind: 'rich-text', runs: makeRichText([{ text: most }, { text: 'z' }]) }),
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  // The ceiling is on code units, not characters. Checked against Excel with a
  // pair of files holding the same 16_384 characters: the one measuring 32_767
  // code units opens, and the one measuring 32_768 is reported as needing
  // repair. A character count would have passed both.
  it('accepts astral text that fills the limit in code units', async () => {
    const text = EMOJI.repeat((MAX_CELL_TEXT_LENGTH - 1) / 2) + 'x';
    expect(text.length).toBe(MAX_CELL_TEXT_LENGTH);
    expect([...text]).toHaveLength((MAX_CELL_TEXT_LENGTH + 1) / 2);
    await expect(save(text)).resolves.toBeInstanceOf(Uint8Array);
  });

  it('rejects astral text one code unit past the limit, though it is half that in characters', async () => {
    const text = EMOJI.repeat((MAX_CELL_TEXT_LENGTH + 1) / 2);
    expect(text.length).toBe(MAX_CELL_TEXT_LENGTH + 1);
    await expect(save(text)).rejects.toThrow(OpenXmlSchemaError);
  });

  it('rejects an over-long cached formula result', async () => {
    await expect(
      save({ kind: 'formula', t: 'normal', formula: 'A2', cachedValue: OVER_LIMIT }),
    ).rejects.toThrow(OpenXmlSchemaError);
  });

  it('refuses the cell on the write-only path without consuming the row number', async () => {
    // A refused cell throws out of appendRow before any of the row reaches the
    // stream, so the caller that shortens the value and appends again has to
    // land on the row the failed call was building.
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const ws = await wb.addWorksheet('S');
    await ws.appendRow(['first']);
    await expect(ws.appendRow(['keep', OVER_LIMIT])).rejects.toThrow(OpenXmlSchemaError);
    await ws.appendRow(['keep', 'short']);
    await ws.close();
    await wb.finalize();

    const sheet = partText(unzipSync(sink.result()), SHEET_PART);
    expect(sheet).toContain('<row r="2">');
    expect(sheet).not.toContain('<row r="3">');
  });

  it('refuses on save a cell the reader accepted, naming it so the caller can shorten it', async () => {
    // The reader has no ceiling of its own: a file from a producer that does
    // not check loads, and the refusal lands on the way back out. Excel cannot
    // open that file either, so writing it back would hand the caller a
    // workbook Excel reports as needing repair.
    const seed = createWorkbook();
    setCell(addWorksheet(seed, 'S'), 1, 1, 'PLACEHOLDER');
    const archive = unzipSync(await workbookToBytes(seed));
    archive[SST_PART] = encoder.encode(partText(archive, SST_PART).replace('PLACEHOLDER', OVER_LIMIT));

    const wb = await loadWorkbook(fromBuffer(zipSync(archive)));
    const ws = expectSheet(wb.sheets[0]?.sheet);
    expect(getCell(ws, 1, 1)?.value).toBe(OVER_LIMIT);
    await expect(workbookToBytes(wb)).rejects.toThrow(/string at A1 is 32768/);

    setCell(ws, 1, 1, OVER_LIMIT.slice(0, MAX_CELL_TEXT_LENGTH));
    await expect(workbookToBytes(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});
