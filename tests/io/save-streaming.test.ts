// Guards the save path's incremental worksheet emission and the part ordering
// that depends on it. Both properties exist so a sheet never has to be held in
// memory in full; see tests/perf/save-heap.test.ts for the heap metric itself.

import { describe, expect, it } from 'vitest';
import { fromArrayBuffer } from '../../src/io/browser.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { makeStylesheet } from '../../src/styles/stylesheet.js';
import { makeSharedStrings } from '../../src/workbook/shared-strings.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { appendRow, setCell } from '../../src/worksheet/worksheet.js';
import { worksheetToBytes, writeWorksheetXml } from '../../src/worksheet/writer.js';
import { openZip } from '../../src/zip/reader.js';

const ROWS = 5_000;
const COLS = 5;

const sheetWithRows = (rows: number) => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  for (let r = 0; r < rows; r++) appendRow(ws, [r, `row-${r}`, r * 1.5, true, r % 7]);
  return { wb, ws };
};

const reload = async (bytes: Uint8Array) => loadWorkbook(fromArrayBuffer(bytes.slice().buffer as ArrayBuffer));

describe('writeWorksheetXml', () => {
  it('emits the part as many small fragments rather than one string', () => {
    const { ws } = sheetWithRows(ROWS);
    const sizes: number[] = [];
    writeWorksheetXml(ws, { sharedStrings: makeSharedStrings(), styles: makeStylesheet() }, (chunk) => {
      sizes.push(chunk.length);
    });

    // One fragment per cell plus the row and document scaffolding. A serialiser
    // that built the part in one piece would emit a single huge fragment, which
    // is exactly the shape this path exists to avoid.
    expect(sizes.length).toBeGreaterThan(ROWS * COLS);
    expect(Math.max(...sizes)).toBeLessThan(1024);
  });

  it('concatenates to byte-identical output with worksheetToBytes', () => {
    const { ws } = sheetWithRows(200);
    const parts: string[] = [];
    writeWorksheetXml(ws, { sharedStrings: makeSharedStrings(), styles: makeStylesheet() }, (chunk) => {
      parts.push(chunk);
    });
    const streamed = new TextEncoder().encode(parts.join(''));
    const whole = worksheetToBytes(ws, { sharedStrings: makeSharedStrings(), styles: makeStylesheet() });
    expect(streamed).toEqual(whole);
  });
});

describe('saveWorkbook part ordering', () => {
  it('writes workbook.xml before the sheet parts and its rels after them', async () => {
    const { wb } = sheetWithRows(50);
    const archive = await openZip(fromArrayBuffer((await workbookToBytes(wb)).slice().buffer as ArrayBuffer));
    // `list()` sorts, so read entry order off the raw archive bytes instead.
    const raw = new TextDecoder('latin1').decode(await workbookToBytes(wb));
    const at = (path: string) => raw.indexOf(path);
    expect(at('xl/workbook.xml')).toBeGreaterThanOrEqual(0);
    expect(at('xl/workbook.xml')).toBeLessThan(at('xl/worksheets/sheet1.xml'));
    expect(at('xl/_rels/workbook.xml.rels')).toBeGreaterThan(at('xl/worksheets/sheet1.xml'));
    archive.close();
  });

  it('records the sharedStrings relationship discovered while serialising sheets', async () => {
    const { wb } = sheetWithRows(20);
    const archive = await openZip(fromArrayBuffer((await workbookToBytes(wb)).slice().buffer as ArrayBuffer));
    const rels = new TextDecoder().decode(archive.read('xl/_rels/workbook.xml.rels'));
    // The table is only populated while sheets serialise, so the rels part has
    // to be built after them or this relationship goes missing and Excel opens
    // the file with every string cell blank.
    expect(rels).toContain('sharedStrings.xml');
    archive.close();

    const reloaded = await reload(await workbookToBytes(wb));
    const cell = reloaded.sheets[0]?.sheet;
    expect(cell && 'rows' in cell ? cell.rows.get(1)?.get(2)?.value : undefined).toBe('row-0');
  });

  it('omits the sharedStrings relationship when no cell holds a string', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Numbers');
    setCell(ws, 1, 1, 41);
    setCell(ws, 1, 2, 1.5);
    const archive = await openZip(fromArrayBuffer((await workbookToBytes(wb)).slice().buffer as ArrayBuffer));
    const rels = new TextDecoder().decode(archive.read('xl/_rels/workbook.xml.rels'));
    expect(rels).not.toContain('sharedStrings.xml');
    expect(archive.has('xl/sharedStrings.xml')).toBe(false);
    archive.close();
  });
});
