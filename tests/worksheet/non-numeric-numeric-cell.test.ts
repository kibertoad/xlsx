// `<v>` under `t="n"` went through `Number.parseFloat` with no check on the
// result, so text became NaN and an overflowing exponent became Infinity, and
// both were stored on the cell. The load reported success. The failure landed
// at the other end of the pipeline, on save, as "cannot serialise non-finite
// number" naming a cell the caller never wrote.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
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

describe('a numeric cell whose value is not a finite number', () => {
  it('loadWorkbook rejects text', async () => {
    expect(() => readSheet('<row r="1"><c r="A1"><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
    await expect(
      loadWorkbook(fromBuffer(Buffer.from(pkg('<row r="1"><c r="A1"><v>oops</v></c></row>')))),
    ).rejects.toThrow(OpenXmlSchemaError);
  });

  it('loadWorkbook rejects an overflowing exponent', () => {
    expect(() => readSheet('<row r="1"><c r="A1"><v>1e400</v></c></row>')).toThrow(OpenXmlSchemaError);
  });

  it('loadWorkbook rejects it in a cached formula result too', () => {
    expect(() => readSheet('<row r="1"><c r="A1"><f>A2</f><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
  });

  it('loadWorkbookStream drops it rather than yielding NaN', async () => {
    const wb = await loadWorkbookStream(
      fromBuffer(Buffer.from(pkg('<row r="1"><c r="A1"><v>oops</v></c><c r="B1"><v>1e400</v></c></row>'))),
    );
    try {
      const rows: unknown[][] = [];
      for await (const row of wb.openWorksheet('S').iterValues()) rows.push(row);
      expect(rows).toEqual([[null, null]]);
    } finally {
      await wb.close();
    }
  });

  it('leaves ordinary numbers and empty cells alone', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>-3.5e-7</v></c><c r="C1"><v></v></c><c r="D1"/></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toBe(42);
    expect(getCell(ws, 1, 2)?.value).toBeCloseTo(-3.5e-7);
    expect(getCell(ws, 1, 3)?.value).toBeNull();
    expect(getCell(ws, 1, 4)?.value).toBeNull();
  });
});
