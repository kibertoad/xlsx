// `<c t="b">` carries an xsd:boolean: its lexical space is `1` / `0` / `true` /
// `false`, and its whiteSpace facet is `collapse`, so the padding a
// pretty-printer leaves around the text is insignificant. The XML parser runs
// with `trimValues: false`, which is why the padded forms are worth pinning.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { getCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';

const enc = new TextEncoder();
const dec = new TextDecoder();
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

const loadRowValues = async (body: string, cols: number): Promise<unknown[]> => {
  const wb = await loadWorkbook(fromBuffer(pkg(body)));
  const ws = getSheet(wb, 'S');
  if (!ws) throw new Error('fixture has no worksheet "S"');
  return Array.from({ length: cols }, (_, i) => getCell(ws, 1, i + 1)?.value);
};

const streamValues = async (body: string): Promise<unknown[][]> => {
  const wb = await loadWorkbookStream(fromBuffer(pkg(body)));
  try {
    const rows: unknown[][] = [];
    for await (const row of wb.openWorksheet('S').iterValues()) rows.push(row);
    return rows;
  } finally {
    await wb.close();
  }
};

const sheetXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/worksheets/sheet1.xml'));
  } finally {
    archive.close();
  }
};

const ALL_FORMS =
  '<row r="1">' +
  '<c r="A1" t="b"><v>true</v></c>' +
  '<c r="B1" t="b"><v>false</v></c>' +
  '<c r="C1" t="b"><v>1</v></c>' +
  '<c r="D1" t="b"><v>0</v></c>' +
  '</row>';

// A producer that indents its parts, or a file run through `xmllint --format`.
const PADDED_FORMS =
  '<row r="1">' +
  '<c r="A1" t="b"><v>\n    1\n  </v></c>' +
  '<c r="B1" t="b"><v> true </v></c>' +
  '<c r="C1" t="b"><v>TRUE</v></c>' +
  '<c r="D1" t="b"><v>\tfalse\t</v></c>' +
  '</row>';

describe('boolean cells over the xsd:boolean lexical space', () => {
  it('loadWorkbook reads all four lexical forms', async () => {
    expect(await loadRowValues(ALL_FORMS, 4)).toEqual([true, false, true, false]);
  });

  it('loadWorkbook ignores insignificant whitespace and case', async () => {
    expect(await loadRowValues(PADDED_FORMS, 4)).toEqual([true, true, true, false]);
  });

  it('loadWorkbookStream reads all four lexical forms', async () => {
    expect(await streamValues(ALL_FORMS)).toEqual([[true, false, true, false]]);
  });

  it('loadWorkbookStream ignores insignificant whitespace and case', async () => {
    expect(await streamValues(PADDED_FORMS)).toEqual([[true, true, true, false]]);
  });

  it('reads a cached formula result in any of the forms', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1" t="b"><f>A2</f><v>true</v></c><c r="B1" t="b"><f>B2</f><v> 0 </v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ kind: 'formula', cachedValue: true });
    expect(getCell(ws, 1, 2)?.value).toMatchObject({ kind: 'formula', cachedValue: false });
  });

  it('rejects a cached result outside the lexical space instead of dropping it', () => {
    // Dropping it would take the `t="b"` the writer derives from it with it,
    // rewriting the cell as an untyped formula on the next save.
    expect(() => readSheet('<row r="1"><c r="A1" t="b"><f>A2</f><v>yes</v></c></row>')).toThrow(OpenXmlSchemaError);
  });

  it('treats a boolean cell with no value as empty, and saves it untyped', async () => {
    const wb = await loadWorkbook(fromBuffer(pkg('<row r="1"><c r="A1" t="b"/><c r="B1" t="b"><v>true</v></c></row>')));
    const ws = getSheet(wb, 'S');
    if (!ws) throw new Error('fixture has no worksheet "S"');
    expect(getCell(ws, 1, 1)?.value).toBeNull();
    const xml = await sheetXml(await workbookToBytes(wb));
    expect(xml).toContain('<c r="A1"/>');
    expect(xml).toContain('<c r="B1" t="b"><v>1</v></c>');
  });

  it('loadWorkbook rejects a value outside the lexical space rather than calling it false', async () => {
    await expect(loadWorkbook(fromBuffer(pkg('<row r="1"><c r="A1" t="b"><v>yes</v></c></row>')))).rejects.toThrow(
      OpenXmlSchemaError,
    );
  });

  it('names the offending cell and bounds the text it quotes', () => {
    const junk = `yes\n${'x'.repeat(5000)}`;
    let message = '';
    try {
      readSheet(`<row r="3"><c r="B3" t="b"><v>${junk}</v></c></row>`);
    } catch (err) {
      message = err instanceof Error ? err.message : '';
    }
    expect(message).toContain('r="B3"');
    expect(message).not.toContain('\n');
    expect(message.length).toBeLessThan(120);
  });

  it('loadWorkbookStream stays lenient on the same input', async () => {
    expect(await streamValues('<row r="1"><c r="A1" t="b"><v>yes</v></c><c r="B1" t="b"/></row>')).toEqual([
      [null, null],
    ]);
  });
});
