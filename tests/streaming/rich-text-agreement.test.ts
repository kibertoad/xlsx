// `loadWorkbook` and `loadWorkbookStream` have to report the same cell value
// for the same bytes. The streaming reader used to join a rich string's runs
// into plain text, so per-run formatting was reachable through one entry point
// and not the other, even though both are typed `CellValue`.

import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import type { CellValue } from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const RUNS = makeRichText([
  { text: 'bo', font: { b: true } },
  { text: 'ld', font: { i: true, sz: 14 } },
  { text: ' plain' },
]);

const modelledValues = async (bytes: Uint8Array): Promise<CellValue[]> => {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ws = getSheet(wb, 'S');
  if (!ws) throw new Error('fixture has no worksheet "S"');
  const out: CellValue[] = [];
  for (let row = 1; row <= 2; row++) out.push(getCell(ws, row, 1)?.value ?? null);
  return out;
};

const streamedValues = async (bytes: Uint8Array): Promise<CellValue[]> => {
  const wb = await loadWorkbookStream(fromBuffer(bytes));
  const out: CellValue[] = [];
  for await (const row of wb.openWorksheet('S').iterRows()) {
    out.push(row[0]?.value ?? null);
  }
  await wb.close();
  return out;
};

/** A package whose strings live in the shared-strings table. */
const sharedStringPackage = async (): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  setCell(ws, 1, 1, { kind: 'rich-text', runs: RUNS });
  setCell(ws, 2, 1, 'plain shared');
  return workbookToBytes(wb);
};

/**
 * A package whose strings are inline rather than shared. Built by rewriting the
 * modelled writer's output, since no public API asks for inline strings
 * directly: the write-only path uses them only once its bounded string table is
 * full, which takes 100_000 entries to reach. The `<si>` bodies are already the
 * CT_Rst the `<is>` branch wants, so each cell's `t="s"` index is swapped for
 * the body it pointed at.
 */
const inlineStringPackage = async (): Promise<Uint8Array> => {
  const { unzipSync, zipSync } = await import('fflate');
  const archive = unzipSync(await sharedStringPackage());
  const dec = new TextDecoder();
  const enc = new TextEncoder();
  const sst = dec.decode(archive['xl/sharedStrings.xml'] as Uint8Array);
  // Lift each `<si>` body into the cell that referenced it by index.
  const bodies = [...sst.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => m[1] as string);
  let sheet = dec.decode(archive['xl/worksheets/sheet1.xml'] as Uint8Array);
  sheet = sheet.replace(/<c r="(A\d)" t="s"><v>(\d+)<\/v><\/c>/g, (_full, ref: string, idx: string) => {
    return `<c r="${ref}" t="inlineStr"><is>${bodies[Number(idx)] as string}</is></c>`;
  });
  archive['xl/worksheets/sheet1.xml'] = enc.encode(sheet);
  // The shared-strings part stays in the package: no cell references it any
  // more, and removing it would leave the workbook rels pointing at nothing.
  return zipSync(archive);
};

describe('the two readers agree on rich text', () => {
  it('shared strings keep their runs in both readers', async () => {
    const bytes = await sharedStringPackage();
    const expected: CellValue[] = [{ kind: 'rich-text', runs: RUNS }, 'plain shared'];
    expect(await modelledValues(bytes)).toEqual(expected);
    expect(await streamedValues(bytes)).toEqual(expected);
  });

  it('inline strings keep their runs in both readers', async () => {
    const bytes = await inlineStringPackage();
    const streamed = await streamedValues(bytes);
    expect(streamed).toEqual(await modelledValues(bytes));
    expect(streamed[0]).toEqual({ kind: 'rich-text', runs: RUNS });
    expect(streamed[1]).toBe('plain shared');
  });

  it('decodes an _xHHHH_ escape inside an inline run', async () => {
    // The old path unescaped each `<t>` as it closed. Building the subtree has
    // to keep that property: one run's text is decoded on its own, so a chunk
    // boundary cannot split a sequence and two runs cannot form one between
    // them.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, { kind: 'rich-text', runs: makeRichText([{ text: 'a\tb' }, { text: '_x0041_' }]) });
    setCell(ws, 2, 1, 'x');
    const { unzipSync, zipSync } = await import('fflate');
    const archive = unzipSync(await workbookToBytes(wb));
    const dec = new TextDecoder();
    const enc = new TextEncoder();
    const sst = dec.decode(archive['xl/sharedStrings.xml'] as Uint8Array);
    const body = (/<si>([\s\S]*?)<\/si>/.exec(sst) as RegExpExecArray)[1] as string;
    let sheet = dec.decode(archive['xl/worksheets/sheet1.xml'] as Uint8Array);
    sheet = sheet.replace(
      /<c r="A1" t="s"><v>0<\/v><\/c>/,
      `<c r="A1" t="inlineStr"><is>${body}</is></c>`,
    );
    archive['xl/worksheets/sheet1.xml'] = enc.encode(sheet);
    const bytes = zipSync(archive);
    expect((await streamedValues(bytes))[0]).toEqual({
      kind: 'rich-text',
      runs: [{ text: 'a\tb' }, { text: '_x0041_' }],
    });
  });

  it('reads a CDATA section inside an inline string the same way', async () => {
    const { unzipSync, zipSync } = await import('fflate');
    const archive = unzipSync(await sharedStringPackage());
    const dec = new TextDecoder();
    const enc = new TextEncoder();
    let sheet = dec.decode(archive['xl/worksheets/sheet1.xml'] as Uint8Array);
    sheet = sheet
      .replace(/<c r="A1" t="s"><v>0<\/v><\/c>/, '<c r="A1" t="inlineStr"><is><t><![CDATA[a<b]]></t></is></c>')
      .replace(
        /<c r="A2" t="s"><v>1<\/v><\/c>/,
        '<c r="A2" t="inlineStr"><is><r><rPr><b/></rPr><t>x<![CDATA[&y]]></t></r></is></c>',
      );
    archive['xl/worksheets/sheet1.xml'] = enc.encode(sheet);
    const bytes = zipSync(archive);
    const streamed = await streamedValues(bytes);
    expect(streamed).toEqual(await modelledValues(bytes));
    expect(streamed[0]).toBe('a<b');
    expect(streamed[1]).toEqual({ kind: 'rich-text', runs: [{ text: 'x&y', font: { b: true } }] });
  });

  it('round-trips a write-only export through the streaming reader', async () => {
    // Write with one entry point and read with the other, which is the shape a
    // streaming export-then-verify pipeline has. These two strings fit the
    // write-only string table, so they land in the shared-strings part; the
    // inline fallback is covered by the cases above.
    const chunks: Uint8Array[] = [];
    const sink = {
      toBytes: () => ({
        write: (c: Uint8Array) => {
          chunks.push(c);
        },
        finish: () => {
          const total = chunks.reduce((n, c) => n + c.byteLength, 0);
          const out = new Uint8Array(total);
          let at = 0;
          for (const c of chunks) {
            out.set(c, at);
            at += c.byteLength;
          }
          return out;
        },
      }),
    };
    const wb = await createWriteOnlyWorkbook(sink as never);
    const ws = await wb.addWorksheet('S');
    await ws.appendRow([{ kind: 'rich-text', runs: RUNS }]);
    await ws.appendRow(['plain shared']);
    await ws.close();
    await wb.finalize();
    const total = chunks.reduce((n, c) => n + c.byteLength, 0);
    const bytes = new Uint8Array(total);
    let at = 0;
    for (const c of chunks) {
      bytes.set(c, at);
      at += c.byteLength;
    }
    expect(await streamedValues(bytes)).toEqual(await modelledValues(bytes));
  });
});
