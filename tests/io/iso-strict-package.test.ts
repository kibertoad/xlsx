// ISO 29500 strict packages. Excel's Save As dialog offers "Strict Open XML
// Spreadsheet", it writes a file with the .xlsx extension, and every part
// inside it carries a purl.oclc.org namespace instead of a
// schemas.openxmlformats.org one. The readers are built on the transitional
// namespaces, so they have to say that rather than report a missing
// relationship, an unexpected root element, or nothing at all.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlNotImplementedError, OpenXmlSchemaError } from '../../src/utils/exceptions.js';
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
const OPC_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OPC_CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/** Dublin Core, which docProps/core.xml carries in both variants. */
const DCORE_NS = 'http://purl.org/dc/elements/1.1/';

const WORKSHEET_PART = 'xl/worksheets/sheet1.xml';
const WORKBOOK_PART = 'xl/workbook.xml';

/** Rewrite the namespace URIs of every XML part `appliesTo` accepts. */
const rezipWithNamespaces = async (
  bytes: Uint8Array,
  pairs: ReadonlyArray<readonly [string, string]>,
  appliesTo: (path: string) => boolean = () => true,
): Promise<Uint8Array> => {
  const archive = await openZip(fromBuffer(bytes));
  const sink = toBuffer();
  const writer = createZipWriter(sink);
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();
  try {
    for (const path of archive.list()) {
      const payload = archive.read(path);
      const isXml = path.endsWith('.xml') || path.endsWith('.rels');
      if (!isXml || !appliesTo(path)) {
        await writer.addEntry(path, payload);
        continue;
      }
      let text = decoder.decode(payload);
      for (const [from, to] of pairs) text = text.split(from).join(to);
      await writer.addEntry(path, encoder.encode(text));
    }
    await writer.finalize();
  } catch (cause) {
    await writer.abort(cause);
    throw cause;
  } finally {
    archive.close();
  }
  return sink.result();
};

/** Build once: each call is a save, an inflate, a rewrite and a deflate. */
const memo = <T>(build: () => Promise<T>): (() => Promise<T>) => {
  let pending: Promise<T> | undefined;
  return () => (pending ??= build());
};

const transitionalPackage = memo(async (): Promise<Uint8Array> => {
  const wb = createWorkbook();
  // core.xml pins the Dublin Core URIs the detector must ignore.
  wb.properties = { creator: 'iso-strict fixture' };
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 'header');
  setCell(ws, 2, 1, 42);
  return workbookToBytes(wb);
});

const strictPackage = memo(async (): Promise<Uint8Array> =>
  rezipWithNamespaces(await transitionalPackage(), STRICT_NAMESPACES),
);

const expectNamedAsStrict = async (load: Promise<unknown>): Promise<void> => {
  await expect(load).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
  await expect(load).rejects.toThrow(/ISO 29500 strict/);
  await expect(load).rejects.toThrow(/Excel Workbook \(\.xlsx\)/);
};

const partText = async (bytes: Uint8Array, path: string): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return new TextDecoder().decode(archive.read(path));
  } finally {
    archive.close();
  }
};

describe('loadWorkbook on an ISO 29500 strict package', () => {
  it('names the format instead of reporting a missing relationship', async () => {
    const load = loadWorkbook(fromBuffer(await strictPackage()));
    await expectNamedAsStrict(load);
    await expect(load).rejects.toThrow(/Strict Open XML Spreadsheet/);
    await expect(load).rejects.not.toThrow(/missing officeDocument relationship/);
  });

  it('catches strict parts behind transitional package relationships', async () => {
    // A converter can leave _rels/.rels transitional while writing strict
    // parts. Without detection in the part readers this surfaces as
    // `parseSharedStringsXml: root is "{…strict…}sst", expected sst`.
    const partsOnly = STRICT_NAMESPACES.filter(([from]) => !from.endsWith('/relationships'));
    const mixed = await rezipWithNamespaces(await transitionalPackage(), partsOnly);
    await expectNamedAsStrict(loadWorkbook(fromBuffer(mixed)));
    await expectNamedAsStrict(loadWorkbookStream(fromBuffer(mixed)));
  });

  it('catches a strict workbook part, which otherwise reads as a sheetless workbook', async () => {
    // The `<sheets>` lookups are transitional QNames, so a strict workbook.xml
    // matches none of them: before the root-element check this package loaded
    // as a workbook with no sheets and no error at all.
    const mixed = await rezipWithNamespaces(
      await transitionalPackage(),
      STRICT_NAMESPACES,
      (path) => path === WORKBOOK_PART,
    );
    await expectNamedAsStrict(loadWorkbook(fromBuffer(mixed)));
    await expectNamedAsStrict(loadWorkbookStream(fromBuffer(mixed)));
  });

  it('catches a strict worksheet part', async () => {
    const mixed = await rezipWithNamespaces(
      await transitionalPackage(),
      STRICT_NAMESPACES,
      (path) => path === WORKSHEET_PART,
    );
    await expectNamedAsStrict(loadWorkbook(fromBuffer(mixed)));
  });

  it('names the format from the streaming reader too', async () => {
    await expectNamedAsStrict(loadWorkbookStream(fromBuffer(await strictPackage())));
  });

  it('keeps the OPC package namespaces, so detection cannot key off them', async () => {
    // Strict rewrites the markup namespaces and leaves ECMA-376 part 2 alone:
    // the container still parses, which is why the rel *type* is what gives it
    // away in the root rels.
    const strict = await strictPackage();
    expect(await partText(strict, '_rels/.rels')).toContain(OPC_RELS_NS);
    expect(await partText(strict, '[Content_Types].xml')).toContain(OPC_CONTENT_TYPES_NS);
  });
});

describe('loadWorkbook on a transitional package', () => {
  it('loads it, Dublin Core purl.org URIs and all', async () => {
    // Guards against a detector written as "any purl URI": docProps/core.xml
    // is purl.org/dc in both variants.
    const bytes = await transitionalPackage();
    expect(await partText(bytes, 'docProps/core.xml')).toContain(DCORE_NS);
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Data']);
    const streamed = await loadWorkbookStream(fromBuffer(bytes));
    expect(streamed.sheetNames).toEqual(['Data']);
  });

  it('rejects a workbook root in an unrelated namespace rather than reading it as sheetless', async () => {
    const wrongNs = await rezipWithNamespaces(
      await transitionalPackage(),
      [
        [
          'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'http://schemas.openxmlformats.org/spreadsheetml/2099/main',
        ],
      ],
      (path) => path === WORKBOOK_PART,
    );
    await expect(loadWorkbook(fromBuffer(wrongNs))).rejects.toBeInstanceOf(OpenXmlSchemaError);
    await expect(loadWorkbook(fromBuffer(wrongNs))).rejects.toThrow(/expected workbook/);
    await expect(loadWorkbookStream(fromBuffer(wrongNs))).rejects.toThrow(/expected workbook/);
  });
});
