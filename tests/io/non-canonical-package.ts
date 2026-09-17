// Packages for the loader tests that cover where a workbook-level part lives.
//
// Every fixture starts from bytes this library wrote and edits them, so it
// carries the parts a real package carries and follows the writer when its
// output changes. A hand-written skeleton would pin the loaders against a
// package no producer emits.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { workbookToBytes } from '../../src/io/save.js';
import { setCellNumberFormat } from '../../src/styles/cell-style.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';

export type Parts = Record<string, Uint8Array>;

export const WORKBOOK_PATH = 'xl/workbook.xml';
export const WB_RELS_PATH = 'xl/_rels/workbook.xml.rels';
const CONTENT_TYPES_PATH = '[Content_Types].xml';

/** Paths neither loader hardcodes. */
const SST_PART = 'xl/strings.xml';
const STYLES_PART = 'xl/cellformats.xml';
const THEME_PART = 'xl/branding.xml';

export const THEME_BYTES = strToU8('<theme/>');
/** Custom format on A2, so a stylesheet that was read is distinguishable from an empty one. */
export const NUMBER_FORMAT = '0.000';

export const zipParts = (parts: Parts): Uint8Array => zipSync(parts);

/** Entry bytes, with a clear failure when the package stopped carrying the part. */
export function part(parts: Parts, path: string): Uint8Array {
  const bytes = parts[path];
  if (!bytes) throw new Error(`fixture: no "${path}" in the package`);
  return bytes;
}

export function movePart(parts: Parts, from: string, to: string): void {
  parts[to] = part(parts, from);
  delete parts[from];
}

/** Edit a part's XML in place. The edit runs on text, which is how these packages differ. */
export function rewrite(parts: Parts, path: string, edit: (xml: string) => string): void {
  parts[path] = strToU8(edit(strFromU8(part(parts, path))));
}

/** Substitute once, failing loudly when writer output no longer holds the needle. */
export function replaceOnce(xml: string, needle: string, replacement: string): string {
  if (!xml.includes(needle)) throw new Error(`fixture: expected "${needle}" in the written package`);
  return xml.replace(needle, replacement);
}

const write = async (text: string, numberFormat?: string): Promise<Parts> => {
  const wb = createWorkbook();
  wb.themeXml = THEME_BYTES;
  const ws = addWorksheet(wb, 'S');
  setCell(ws, 1, 1, text);
  const numeric = setCell(ws, 2, 1, 42);
  if (numberFormat !== undefined) setCellNumberFormat(wb, numeric, numberFormat);
  return unzipSync(await workbookToBytes(wb));
};

interface Move {
  /** Entry the writer emits. */
  from: string;
  /** Where the fixture puts it instead. */
  to: string;
  /** Rel target the writer emits, relative to `xl/`. */
  target: string;
}

const MOVES: ReadonlyArray<Move> = [
  { from: 'xl/sharedStrings.xml', to: SST_PART, target: 'sharedStrings.xml' },
  { from: 'xl/styles.xml', to: STYLES_PART, target: 'styles.xml' },
  { from: 'xl/theme/theme1.xml', to: THEME_PART, target: 'theme/theme1.xml' },
];

const relTargetFor = (path: string): string => path.slice('xl/'.length);

/**
 * `hello` in A1 as a shared string and a custom number format on A2, with
 * sharedStrings, styles and theme moved off the conventional paths and the
 * rels plus the manifest pointed at where they went.
 */
export async function relocatedPackage(): Promise<Parts> {
  const parts = await write('hello', NUMBER_FORMAT);
  for (const move of MOVES) movePart(parts, move.from, move.to);
  rewrite(parts, WB_RELS_PATH, (xml) =>
    MOVES.reduce((out, m) => replaceOnce(out, `Target="${m.target}"`, `Target="${relTargetFor(m.to)}"`), xml),
  );
  rewrite(parts, CONTENT_TYPES_PATH, (xml) =>
    MOVES.reduce((out, m) => replaceOnce(out, `PartName="/${m.from}"`, `PartName="/${m.to}"`), xml),
  );
  return parts;
}

/**
 * A second workbook's parts, to leave at the conventional paths of a package
 * whose rels point elsewhere. Different text, and no custom number format, so
 * either part is recognisable in what a loader returns.
 */
export async function stalePackage(): Promise<Parts> {
  return write('STALE');
}
