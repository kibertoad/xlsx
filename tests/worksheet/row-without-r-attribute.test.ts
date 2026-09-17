// `@r` is optional on CT_Row (ECMA-376 section 18.3.1.73). Excel always
// writes it, so a generator that leaves it off produced two different
// failures: `loadWorkbook` rejected the file outright, and
// `loadWorkbookStream` numbered every such row 0, dropped it for falling
// below `minRow`, and reported an empty sheet without raising anything.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { getCell } from '../../src/worksheet/worksheet.js';

const enc = new TextEncoder();
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SPREADSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

const sheet = (body: string): string =>
  `${XML_DECL}<worksheet xmlns="${MAIN_NS}"><sheetData>${body}</sheetData></worksheet>`;

const readSheet = (body: string) => parseWorksheetXml(sheet(body), 'S', { sharedStrings: [] });

const pkg = (body: string): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="${CT_NS}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/xl/workbook.xml" ContentType="${SPREADSHEET_CT}.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${SPREADSHEET_CT}.worksheet+xml"/>` +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    'xl/workbook.xml': enc.encode(
      `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': enc.encode(sheet(body)),
  });

const streamRows = async (body: string, opts?: { minRow?: number }): Promise<unknown[][]> => {
  const wb = await loadWorkbookStream(fromBuffer(Buffer.from(pkg(body))));
  try {
    const rows: unknown[][] = [];
    for await (const row of wb.openWorksheet('S').iterValues(opts)) rows.push(row);
    return rows;
  } finally {
    await wb.close();
  }
};

// Three rows, none carrying @r: they are rows 1, 2 and 3.
const BARE = '<row><c r="A1" t="str"><v>a</v></c></row>' +
  '<row><c r="A2" t="str"><v>b</v></c></row>' +
  '<row><c r="A3" t="str"><v>c</v></c></row>';

describe('<row> without @r', () => {
  it('loadWorkbook numbers the rows in document order', () => {
    const ws = readSheet(BARE);
    expect([1, 2, 3].map((row) => getCell(ws, row, 1)?.value)).toEqual(['a', 'b', 'c']);
  });

  it('loadWorkbookStream yields every row', async () => {
    expect(await streamRows(BARE)).toEqual([['a'], ['b'], ['c']]);
  });

  it('continues numbering after a row that does carry @r', () => {
    // `<row r="5">` then a bare row, which is row 6.
    const ws = readSheet(
      '<row><c r="A1" t="str"><v>a</v></c></row>' +
        '<row r="5"><c r="A5" t="str"><v>e</v></c></row>' +
        '<row><c r="A6" t="str"><v>f</v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toBe('a');
    expect(getCell(ws, 5, 1)?.value).toBe('e');
    expect(getCell(ws, 6, 1)?.value).toBe('f');
  });

  it('applies a streaming band query to the derived row numbers', async () => {
    expect(await streamRows(BARE, { minRow: 2 })).toEqual([['b'], ['c']]);
  });

  it('still rejects a row whose @r is present but not a positive integer', () => {
    expect(() => readSheet('<row r="0"><c r="A1" t="str"><v>a</v></c></row>')).toThrow();
    expect(() => readSheet('<row r="nope"><c r="A1" t="str"><v>a</v></c></row>')).toThrow();
  });
});
