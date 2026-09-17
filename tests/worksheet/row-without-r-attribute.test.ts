// `@r` is optional on CT_Row (ECMA-376 §18.3.1.73). Excel always writes it, so
// a sheet that leaves it off is where the three places that number a row can
// disagree: the DOM parser, the streaming SAX walk, and the byte-level
// row-offset index a band query seeks with.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
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

const openStream = (body: string) => loadWorkbookStream(fromBuffer(Buffer.from(pkg(body))));

/** Every cell the streaming reader yields, as `row:col=value`. */
const streamCells = async (body: string, opts?: { minRow?: number; maxRow?: number; minCol?: number }): Promise<string[]> => {
  const wb = await openStream(body);
  try {
    const out: string[] = [];
    for await (const row of wb.openWorksheet('S').iterRows(opts)) {
      for (const c of row) out.push(`${c.row}:${c.col}=${String(c.value)}`);
    }
    return out;
  } finally {
    await wb.close();
  }
};

/** The row number the streaming reader labels each yielded batch with. */
const streamRowNumbers = async (body: string): Promise<number[]> => {
  const wb = await openStream(body);
  try {
    const out: number[] = [];
    for await (const row of wb.openWorksheet('S').iterRows()) {
      const first = row[0];
      if (first) out.push(first.row);
    }
    return out;
  } finally {
    await wb.close();
  }
};

/** Every cell `loadWorkbook` placed, in the same `row:col=value` shape. */
const parsedCells = (body: string): string[] => {
  const ws = readSheet(body);
  const out: string[] = [];
  for (const [row, cols] of [...ws.rows].sort((a, b) => a[0] - b[0])) {
    for (const [col, cell] of [...cols].sort((a, b) => a[0] - b[0])) {
      out.push(`${row}:${col}=${String(cell.value)}`);
    }
  }
  return out;
};

/** `<row>` with no `@r`, holding one string cell that carries no ref either. */
const bareRow = (text: string): string => `<row><c t="str"><v>${text}</v></c></row>`;
/** `<row>` with no `@r`, holding one string cell that locates itself. */
const locatedRow = (ref: string, text: string): string =>
  `<row><c r="${ref}" t="str"><v>${text}</v></c></row>`;
/** `<row r="…">` holding one string cell that carries no ref. */
const numberedRow = (r: string, text: string): string =>
  `<row r="${r}"><c t="str"><v>${text}</v></c></row>`;

const BARE = bareRow('a') + bareRow('b') + bareRow('c');
const LOCATED = locatedRow('A1', 'a') + locatedRow('A2', 'b') + locatedRow('A3', 'c');
const MIXED_HEAD = numberedRow('1', 'a') + bareRow('b') + bareRow('c');
const MIXED_TAIL = BARE + numberedRow('4', 'd');
const SHAPES = [BARE, LOCATED, MIXED_HEAD, MIXED_TAIL, numberedRow('5', 'e') + bareRow('f')];

describe('<row> without @r: loadWorkbook', () => {
  it('numbers the rows in document order when the cells are unlocated too', () => {
    const ws = readSheet(BARE);
    expect([1, 2, 3].map((row) => getCell(ws, row, 1)?.value)).toEqual(['a', 'b', 'c']);
  });

  it('numbers the rows in document order when the cells carry refs', () => {
    const ws = readSheet(LOCATED);
    expect([1, 2, 3].map((row) => getCell(ws, row, 1)?.value)).toEqual(['a', 'b', 'c']);
  });

  it('keys a row dimension by the derived row', () => {
    const ws = readSheet('<row r="1"/><row ht="30" customHeight="1"/>');
    expect([...ws.rowDimensions.keys()]).toEqual([2]);
  });

  it('continues numbering after a row that does carry @r', () => {
    const ws = readSheet(numberedRow('5', 'e') + bareRow('f'));
    expect(getCell(ws, 5, 1)?.value).toBe('e');
    expect(getCell(ws, 6, 1)?.value).toBe('f');
  });

  it('takes the row from the first cell that locates itself, gap and all', () => {
    const ws = readSheet(locatedRow('A1', 'a') + locatedRow('A3', 'c'));
    expect(getCell(ws, 1, 1)?.value).toBe('a');
    expect(getCell(ws, 3, 1)?.value).toBe('c');
    expect(getCell(ws, 2, 1)).toBeUndefined();
  });

  it('cannot land on a row an out-of-order @r already filled', () => {
    const ws = readSheet(numberedRow('2', 'two') + numberedRow('1', 'one') + bareRow('third'));
    expect([1, 2, 3].map((row) => getCell(ws, row, 1)?.value)).toEqual(['one', 'two', 'third']);
  });

  it('reads an @r padded with the whitespace xsd:unsignedInt collapses', () => {
    const ws = readSheet(numberedRow(' 3 ', 'c') + bareRow('d'));
    expect(getCell(ws, 3, 1)?.value).toBe('c');
    expect(getCell(ws, 4, 1)?.value).toBe('d');
  });

  it('rejects an @r that is not a row number', () => {
    expect(() => readSheet(numberedRow('0', 'a'))).toThrow(OpenXmlSchemaError);
    expect(() => readSheet(numberedRow('0', 'a'))).toThrow(/<row r="0"> is not a row number/);
    expect(() => readSheet(numberedRow('nope', 'a'))).toThrow(/<row r="nope"> is not a row number/);
    expect(() => readSheet(numberedRow('12abc', 'a'))).toThrow(/<row r="12abc"> is not a row number/);
    expect(() => readSheet(numberedRow('1048577', 'a'))).toThrow(/is not a row number in \[1, 1048576\]/);
  });

  it('rejects a derived row past the last row', () => {
    const body = '<row r="1048576" ht="30" customHeight="1"/><row ht="31" customHeight="1"/>';
    expect(() => readSheet(body)).toThrow(OpenXmlSchemaError);
    expect(() => readSheet(body)).toThrow(/falls past the last row \(1048576\)/);
  });

  it('still rejects a cell whose ref names a row other than its own', () => {
    expect(() => readSheet('<row r="2"><c r="A3" t="str"><v>c</v></c></row>')).toThrow(
      /<c r="A3"> row 3 disagrees with row 2/,
    );
    expect(() =>
      readSheet('<row><c r="A1" t="str"><v>a</v></c><c r="B2" t="str"><v>b</v></c></row>'),
    ).toThrow(/<c r="B2"> row 2 disagrees with row 1/);
  });
});

describe('<row> without @r: loadWorkbookStream', () => {
  it('yields every row when the cells are unlocated too', async () => {
    expect(await streamCells(BARE)).toEqual(['1:1=a', '2:1=b', '3:1=c']);
  });

  it('yields every row when the cells carry refs', async () => {
    expect(await streamCells(LOCATED)).toEqual(['1:1=a', '2:1=b', '3:1=c']);
  });

  it('numbers a row that holds no cell at all', async () => {
    expect(await streamRowNumbers('<row ht="30" customHeight="1"/>' + bareRow('b'))).toEqual([2]);
  });

  it('applies a band query to the derived row numbers', async () => {
    expect(await streamCells(BARE, { minRow: 2 })).toEqual(['2:1=b', '3:1=c']);
  });

  it('applies a band query when the first row is the only one carrying @r', async () => {
    expect(await streamCells(MIXED_HEAD, { minRow: 2 })).toEqual(['2:1=b', '3:1=c']);
  });

  it('applies a band query when the bare rows come before an @r row', async () => {
    expect(await streamCells(MIXED_TAIL, { minRow: 2 })).toEqual(['2:1=b', '3:1=c', '4:1=d']);
  });

  it('keeps a band query over a gap only the cell refs record', async () => {
    const body = numberedRow('1', 'a') + locatedRow('A3', 'c') + numberedRow('4', 'd');
    expect(await streamCells(body)).toEqual(['1:1=a', '3:1=c', '4:1=d']);
    expect(await streamCells(body, { minRow: 3 })).toEqual(['3:1=c', '4:1=d']);
  });

  it('matches the full walk at every minRow, on every shape of sheet', async () => {
    for (const body of SHAPES) {
      const all = await streamCells(body);
      for (const minRow of [1, 2, 3, 4, 5, 6]) {
        const band = await streamCells(body, { minRow });
        expect(band).toEqual(all.filter((c) => Number(c.split(':')[0]) >= minRow));
      }
    }
  });

  it('seeks a band by an @r padded with the whitespace xsd:unsignedInt collapses', async () => {
    const body = numberedRow(' 1 ', 'a') + numberedRow(' 2 ', 'b') + numberedRow(' 3 ', 'c');
    expect(await streamCells(body, { minRow: 2 })).toEqual(['2:1=b', '3:1=c']);
  });

  it('rejects an @r that is not a row number, as loadWorkbook does', async () => {
    await expect(streamCells(numberedRow('nope', 'a'))).rejects.toThrow(OpenXmlSchemaError);
    await expect(streamCells(numberedRow('nope', 'a'))).rejects.toThrow(
      /<row r="nope"> is not a row number/,
    );
    await expect(streamCells(numberedRow('0', 'a'))).rejects.toThrow(/<row r="0"> is not a row number/);
  });

  it('agrees with loadWorkbook on where every cell lands', async () => {
    for (const body of SHAPES) {
      expect(await streamCells(body)).toEqual(parsedCells(body));
    }
  });
});


describe('mixed located and unlocated cells', () => {
  const body = '<row><c t="str"><v>a</v></c><c r="B3" t="str"><v>b</v></c><c t="str"><v>c</v></c></row>' + bareRow('d');
  it('uses the first located cell even when it is not the first cell', async () => {
    expect(await streamCells(body)).toEqual(parsedCells(body));
  });
  it('keeps the whole derived row in a band query', async () => {
    expect(await streamCells(body, { minRow: 3, maxRow: 3 })).toEqual(['3:1=a', '3:2=b', '3:3=c']);
  });
  it('derives columns independently of the column filter', async () => {
    const unlocated = '<row><c t="str"><v>a</v></c><c t="str"><v>b</v></c></row>';
    expect(await streamCells(unlocated, { minCol: 2 })).toEqual(['1:2=b']);
  });
  it('accepts an explicit plus sign in unsignedInt row attributes', async () => {
    const signed = numberedRow('+1', 'a') + numberedRow(' +2 ', 'b');
    expect(parsedCells(signed)).toEqual(['1:1=a', '2:1=b']);
    expect(await streamCells(signed, { minRow: 2 })).toEqual(['2:1=b']);
  });
});


it('reports the resolved row for an invalid number before the first located cell', async () => {
  const body = '<row><c><v>oops</v></c><c r="B3"><v>2</v></c></row>';
  await expect(streamCells(body, { minRow: 3 })).rejects.toThrow('at S!A3 is not a finite number');
  await expect(streamCells(body, { minRow: 4 })).resolves.toEqual([]);
});
