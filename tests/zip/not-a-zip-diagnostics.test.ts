// What openZip says about input that is not a zip. Upload paths validate by
// file extension, so a CSV or a PDF renamed to .xlsx reaches the loader and
// "archive is not a valid zip" is true but tells its user nothing.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import {
  OpenXmlContentLimitError,
  OpenXmlError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
  OpenXmlUnsupportedFormatError,
} from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

const ascii = (s: string): Uint8Array => new TextEncoder().encode(s);

const openError = async (bytes: Uint8Array): Promise<Error> => {
  try {
    await openZip(fromBuffer(bytes));
  } catch (err) {
    if (err instanceof Error) return err;
    throw err;
  }
  throw new Error('expected openZip to reject');
};

const EOCD_SIG = 0x06054b50;

/**
 * A complete archive: local file headers at the front, an intact EOCD at the
 * back, and the central directory between them overwritten. Nothing is missing
 * from the file, so the reader finds the trailer and only then discovers that
 * the index it points at is gibberish.
 */
const zipWithCorruptCentralDirectory = async (): Promise<Uint8Array> => {
  const sink = toBuffer();
  const writer = createZipWriter(sink);
  await writer.addEntry('a.txt', ascii('first entry'));
  await writer.addEntry('b.txt', ascii('second entry'));
  await writer.finalize();
  const bytes = sink.result();

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocdOff = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocdOff = i;
      break;
    }
  }
  if (eocdOff < 0) throw new Error('fixture zip has no EOCD');

  const cdOffset = view.getUint32(eocdOff + 16, true);
  bytes.fill(0xff, cdOffset, eocdOff);
  return bytes;
};

describe('openZip names what the leading bytes look like', () => {
  it('CSV content reads as plain text', async () => {
    const err = await openError(ascii('id,name,amount\n1,Widget,4.50\n2,Gadget,9.00\n'));
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/not a valid zip/);
    expect(err.message).toMatch(/plain text/);
    expect(err.message).toMatch(/CSV/);
  });

  it('a UTF-8 byte-order mark is named as such', async () => {
    const bom = new Uint8Array([0xef, 0xbb, 0xbf, ...ascii('id,name\n1,Widget\n')]);
    const err = await openError(bom);
    expect(err.message).toMatch(/UTF-8 byte-order mark/);
  });

  it('a UTF-16 byte-order mark is named as such', async () => {
    const utf16 = new Uint8Array([0xff, 0xfe, 0x69, 0x00, 0x64, 0x00, 0x0a, 0x00]);
    const err = await openError(utf16);
    expect(err.message).toMatch(/UTF-16 byte-order mark/);
  });

  it('a PDF is named as a PDF', async () => {
    const pdf = new Uint8Array([...ascii('%PDF-1.7\n'), 0x80, 0x01, 0x02, 0x03]);
    const err = await openError(pdf);
    expect(err.message).toMatch(/a PDF/);
  });

  it('a zip prefix with no central directory reads as truncated', async () => {
    const truncated = new Uint8Array(64);
    truncated.set([0x50, 0x4b, 0x03, 0x04]);
    const err = await openError(truncated);
    expect(err.message).toMatch(/truncated or partially uploaded/);
  });

  // A whole archive whose index rotted is a different repair from a transfer
  // that stopped early, and the local file header at offset 0 is identical in
  // both. Sending the owner of a complete file off to re-upload it wastes the
  // one thing the message is for.
  it('a whole zip with a corrupt central directory does not read as truncated', async () => {
    const err = await openError(await zipWithCorruptCentralDirectory());
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/end-of-central-directory signature was found/);
    expect(err.message).not.toMatch(/truncated or partially uploaded/);
  });

  it('an EOCD signature alone does not prove the central directory is present', async () => {
    const bytes = new Uint8Array(64);
    const view = new DataView(bytes.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint32(42, EOCD_SIG, true);
    view.setUint16(50, 1, true);
    view.setUint16(52, 1, true);
    view.setUint32(58, 0, true); // Points at a local header; no directory exists.
    view.setUint16(10, 99, true); // Also prevents permissive fallback decoding.
    const err = await openError(bytes);
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/end-of-central-directory signature was found/);
    expect(err.message).not.toMatch(/central directory is present/);
  });

  it('says nothing extra about bytes no magic number identifies', async () => {
    const random = new Uint8Array(64);
    for (let i = 0; i < random.length; i++) random[i] = (i * 37 + 129) & 0xff;
    const err = await openError(random);
    expect(err.message).toMatch(/not a valid zip/);
    expect(err.message).not.toMatch(/leading bytes/);
  });

  it('an empty file reports its length against the minimum record size', async () => {
    const err = await openError(new Uint8Array(0));
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/is 0 bytes/);
    expect(err.message).toMatch(/22-byte minimum/);
  });

  it('a too-short text file is still named as text', async () => {
    const err = await openError(ascii('a,b\n'));
    expect(err.message).toMatch(/22-byte minimum/);
    expect(err.message).toMatch(/plain text/);
  });

  it('the diagnosis survives the loadWorkbook path, not just openZip', async () => {
    await expect(loadWorkbook(fromBuffer(ascii('id,name\n1,Widget\n')))).rejects.toThrow(
      /plain text/,
    );
  });
});

// Pins the classes the README's error table and the loadWorkbook docstring
// promise, so the documented contract cannot drift without a test failing.
describe('the class loadWorkbook throws is the contract', () => {
  it('input that is not a zip is an OpenXmlIoError', async () => {
    await expect(loadWorkbook(fromBuffer(ascii('id,name\n1,Widget\n')))).rejects.toBeInstanceOf(
      OpenXmlIoError,
    );
  });

  // The wording and the `format` discriminant belong to tests/zip/cfb.test.ts;
  // this only pins the row of the table, including that the narrower class
  // still answers a catch written against the wider one.
  it('another Office format is an OpenXmlUnsupportedFormatError', async () => {
    const biff2 = new Uint8Array(64);
    biff2.set([0x09, 0x00, 0x04, 0x00]);
    const attempt = loadWorkbook(fromBuffer(biff2));
    await expect(attempt).rejects.toBeInstanceOf(OpenXmlUnsupportedFormatError);
    await expect(attempt).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
  });

  it('a readable zip that is not an OPC package is an OpenXmlSchemaError', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('hello.txt', ascii('a zip, but not a package'));
    await writer.finalize();
    await expect(loadWorkbook(fromBuffer(sink.result()))).rejects.toBeInstanceOf(
      OpenXmlSchemaError,
    );
  });

  // The one row that hangs off OpenXmlError directly. A catch ladder built
  // from the rows above it drops through, which is the whole reason the table
  // has to list it.
  it('a workbook past contentLimits is an OpenXmlContentLimitError and nothing narrower', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 1);
    setCell(ws, 2, 1, 2);
    const bytes = await workbookToBytes(wb);

    const attempt = loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 1 } });
    await expect(attempt).rejects.toBeInstanceOf(OpenXmlContentLimitError);
    await expect(attempt).rejects.toBeInstanceOf(OpenXmlError);
    await expect(attempt).rejects.not.toBeInstanceOf(OpenXmlIoError);
    await expect(attempt).rejects.not.toBeInstanceOf(OpenXmlSchemaError);
    await expect(attempt).rejects.not.toBeInstanceOf(OpenXmlNotImplementedError);
  });

  it('every one of them is an OpenXmlError, which is what a caller catches', async () => {
    const err = await openError(ascii('id,name\n1,Widget\n'));
    expect(err).toBeInstanceOf(OpenXmlError);
  });
});
