// `xl/sharedStrings.xml` and `xl/styles.xml` are a convention, not a rule:
// what binds either part to the workbook is its relationship. `loadWorkbook`
// has always followed the rels when the conventional path is absent;
// `loadWorkbookStream` only looked at the conventional path, so every
// `t="s"` cell in such a workbook came back as `null` and every cell style
// resolved against an empty pool, both without raising anything.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const enc = new TextEncoder();
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SPREADSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

// Deliberately not the paths either loader hardcodes.
const SST_PART = 'xl/strings.xml';
const STYLES_PART = 'xl/cellformats.xml';

const fixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="${CT_NS}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/xl/workbook.xml" ContentType="${SPREADSHEET_CT}.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${SPREADSHEET_CT}.worksheet+xml"/>` +
        `<Override PartName="/${SST_PART}" ContentType="${SPREADSHEET_CT}.sharedStrings+xml"/>` +
        `<Override PartName="/${STYLES_PART}" ContentType="${SPREADSHEET_CT}.styles+xml"/>` +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
        '</Relationships>',
    ),
    'xl/workbook.xml': enc.encode(
      `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL_NS}/sharedStrings" Target="strings.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_NS}/styles" Target="cellformats.xml"/>` +
        '</Relationships>',
    ),
    [SST_PART]: enc.encode(
      `${XML_DECL}<sst xmlns="${MAIN_NS}" count="1" uniqueCount="1"><si><t>hello</t></si></sst>`,
    ),
    [STYLES_PART]: enc.encode(
      `${XML_DECL}<styleSheet xmlns="${MAIN_NS}">` +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
        '</styleSheet>',
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}"><sheetData>` +
        '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>' +
        '</sheetData></worksheet>',
    ),
  });

describe('sharedStrings and styles at a path the rels choose', () => {
  const bytes = fixture();

  it('loadWorkbook resolves both through the workbook rels', async () => {
    const wb = await loadWorkbook(fromBuffer(Buffer.from(bytes)));
    const sheet = wb.sheets[0]?.sheet as Worksheet;
    expect(getCell(sheet, 1, 1)?.value).toBe('hello');
    expect(wb.styles.cellXfs.length).toBe(2);
  });

  it('loadWorkbookStream resolves both through the workbook rels', async () => {
    const wb = await loadWorkbookStream(fromBuffer(Buffer.from(bytes)));
    try {
      const rows: unknown[][] = [];
      for await (const row of wb.openWorksheet('S').iterValues()) rows.push(row);
      expect(rows).toEqual([['hello']]);
      expect(wb.styles.cellXfs.length).toBe(2);
    } finally {
      await wb.close();
    }
  });
});
