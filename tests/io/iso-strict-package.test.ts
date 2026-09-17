// ISO 29500 strict packages. Excel's Save As dialog offers "Strict Open XML
// Spreadsheet", it writes a file with the .xlsx extension, and every part
// inside it carries a purl.oclc.org namespace instead of a
// schemas.openxmlformats.org one. The reader is built on the transitional
// namespaces, so it has to say that rather than report a missing relationship.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlNotImplementedError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

/**
 * Transitional namespace to its strict counterpart. Two of the pairs rename
 * the local part as well (`extended-properties`, `custom-properties`), which
 * is why this is a table rather than a prefix swap.
 */
const STRICT_NAMESPACES: ReadonlyArray<readonly [string, string]> = [
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://purl.oclc.org/ooxml/officeDocument/relationships',
  ],
  [
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'http://purl.oclc.org/ooxml/spreadsheetml/main',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'http://purl.oclc.org/ooxml/drawingml/main',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'http://purl.oclc.org/ooxml/officeDocument/extendedProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
    'http://purl.oclc.org/ooxml/officeDocument/customProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
  ],
];

/** Namespaces the package layer keeps in both variants (ECMA-376 part 2). */
const PACKAGE_NAMESPACES = [
  'http://schemas.openxmlformats.org/package/2006/relationships',
  'http://schemas.openxmlformats.org/package/2006/content-types',
];

const rezipWithNamespaces = async (
  bytes: Uint8Array,
  pairs: ReadonlyArray<readonly [string, string]>,
): Promise<Uint8Array> => {
  const archive = await openZip(fromBuffer(bytes));
  const sink = toBuffer();
  const writer = createZipWriter(sink);
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();
  for (const path of archive.list()) {
    const payload = archive.read(path);
    if (!path.endsWith('.xml') && !path.endsWith('.rels')) {
      await writer.addEntry(path, payload);
      continue;
    }
    let text = decoder.decode(payload);
    for (const [from, to] of pairs) text = text.split(from).join(to);
    await writer.addEntry(path, encoder.encode(text));
  }
  archive.close();
  await writer.finalize();
  return sink.result();
};

const transitionalWorkbookBytes = async (): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 'header');
  setCell(ws, 2, 1, 42);
  return workbookToBytes(wb);
};

describe('loadWorkbook on an ISO 29500 strict package', () => {
  it('names the format instead of reporting a missing relationship', async () => {
    const strict = await rezipWithNamespaces(
      await transitionalWorkbookBytes(),
      STRICT_NAMESPACES,
    );
    const load = loadWorkbook(fromBuffer(strict));
    await expect(load).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
    await expect(load).rejects.toThrow(/ISO 29500 strict/);
    await expect(load).rejects.toThrow(/Strict Open XML Spreadsheet/);
    await expect(load).rejects.not.toThrow(/missing officeDocument relationship/);
  });

  it('says how to get a readable file', async () => {
    const strict = await rezipWithNamespaces(
      await transitionalWorkbookBytes(),
      STRICT_NAMESPACES,
    );
    await expect(loadWorkbook(fromBuffer(strict))).rejects.toThrow(/Excel Workbook \(\.xlsx\)/);
  });

  it('catches strict parts behind transitional package relationships', async () => {
    // A converter can leave _rels/.rels transitional while writing strict
    // parts. Without the second check that package reads as a workbook with no
    // sheets, or fails somewhere deep in a part reader.
    const partsOnly = STRICT_NAMESPACES.filter(
      ([from]) => !from.endsWith('/relationships'),
    );
    const mixed = await rezipWithNamespaces(await transitionalWorkbookBytes(), partsOnly);
    await expect(loadWorkbook(fromBuffer(mixed))).rejects.toThrow(/ISO 29500 strict/);
  });

  it('leaves a transitional package loading as before', async () => {
    const wb = await loadWorkbook(fromBuffer(await transitionalWorkbookBytes()));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Data']);
  });

  it('does not fire on the package namespaces, which strict keeps', async () => {
    // Guards against a detector written as "any purl.oclc.org URI": the OPC
    // layer is the same in both variants, so these must stay untouched.
    const bytes = await transitionalWorkbookBytes();
    const archive = await openZip(fromBuffer(bytes));
    const rels = new TextDecoder().decode(archive.read('_rels/.rels'));
    archive.close();
    expect(rels).toContain(PACKAGE_NAMESPACES[0]);
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.sheets.length).toBe(1);
  });
});
