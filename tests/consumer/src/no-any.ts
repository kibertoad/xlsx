// Type-level half of the packaging gate.
//
// The failure this guards against is silent. When a shipped `.d.ts` carries a
// relative specifier Node's ESM rules reject (`from './load'` instead of
// `from './load.js'`), a consumer on `moduleResolution: node16` gets TS2834
// instead. That error lands inside `node_modules`, so the near-universal
// `skipLibCheck: true` swallows it and the affected declarations degrade to
// TypeScript's error type. A fixture that merely imports and calls the API
// still compiles clean, which is why the assertions below exist instead.
//
// They work because a conditional type whose checked side is `any` or the
// error type takes both branches and collapses to `boolean`. `Expect` only
// accepts `true`, so any declaration that stopped resolving fails to compile
// here, in the consumer's own source, where nothing suppresses it.
//
// Note that `keyof`- and `NotAny`-style probes do NOT work: `keyof any` is
// `string | number | symbol`, and `0 extends 1 & T` distinguishes `any` but
// not the error type. Only the collapse-to-`boolean` behaviour catches both.

import type { Cell } from '@office-kit/xlsx/cell';
import { makeRichText } from '@office-kit/xlsx/cell';
import type { ChartSpace } from '@office-kit/xlsx/chart';
import { makeBarChart } from '@office-kit/xlsx/chart';
import type { Chartsheet } from '@office-kit/xlsx/chartsheet';
import { makeChartsheet } from '@office-kit/xlsx/chartsheet';
import type { Drawing } from '@office-kit/xlsx/drawing';
import { makeOneCellAnchor } from '@office-kit/xlsx/drawing';
import type { LoadOptions } from '@office-kit/xlsx/io';
import { loadWorkbook } from '@office-kit/xlsx/io';
import { fromFile } from '@office-kit/xlsx/node';
import type { CoreProperties } from '@office-kit/xlsx/packaging';
import { makeCoreProperties } from '@office-kit/xlsx/packaging';
import type { Schema } from '@office-kit/xlsx/schema';
import { defineSchema } from '@office-kit/xlsx/schema';
import type { ReadOnlyWorkbook } from '@office-kit/xlsx/streaming';
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';
import type { Alignment } from '@office-kit/xlsx/styles';
import { makeAlignment } from '@office-kit/xlsx/styles';
import { inferCellType, OpenXmlError } from '@office-kit/xlsx/utils';
import type { Workbook } from '@office-kit/xlsx/workbook';
import { createWorkbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { setCell } from '@office-kit/xlsx/worksheet';
import type { XmlNode } from '@office-kit/xlsx/xml';
import { parseXml } from '@office-kit/xlsx/xml';
import type { DecompressionLimits } from '@office-kit/xlsx/zip';
import { DEFAULT_DECOMPRESSION_LIMITS, openZip } from '@office-kit/xlsx/zip';

type Expect<T extends true> = T;

type IsInterface<T> = T extends object ? true : false;
type IsFunction<T> = T extends (...args: never[]) => unknown ? true : false;
type IsClass<T> = T extends abstract new (...args: never[]) => unknown ? true : false;

export type SubpathTypesResolve = [
  Expect<IsInterface<Cell>>,
  Expect<IsFunction<typeof makeRichText>>,
  Expect<IsInterface<ChartSpace>>,
  Expect<IsFunction<typeof makeBarChart>>,
  Expect<IsInterface<Chartsheet>>,
  Expect<IsFunction<typeof makeChartsheet>>,
  Expect<IsInterface<Drawing>>,
  Expect<IsFunction<typeof makeOneCellAnchor>>,
  Expect<IsInterface<LoadOptions>>,
  Expect<IsFunction<typeof loadWorkbook>>,
  Expect<IsFunction<typeof fromFile>>,
  Expect<IsInterface<CoreProperties>>,
  Expect<IsFunction<typeof makeCoreProperties>>,
  Expect<IsInterface<Schema<unknown>>>,
  Expect<IsFunction<typeof defineSchema>>,
  Expect<IsInterface<ReadOnlyWorkbook>>,
  Expect<IsFunction<typeof loadWorkbookStream>>,
  Expect<IsInterface<Alignment>>,
  Expect<IsFunction<typeof makeAlignment>>,
  Expect<IsClass<typeof OpenXmlError>>,
  Expect<IsFunction<typeof inferCellType>>,
  Expect<IsInterface<Workbook>>,
  Expect<IsFunction<typeof createWorkbook>>,
  Expect<IsInterface<Worksheet>>,
  Expect<IsFunction<typeof setCell>>,
  Expect<IsInterface<XmlNode>>,
  Expect<IsFunction<typeof parseXml>>,
  Expect<IsInterface<DecompressionLimits>>,
  Expect<IsFunction<typeof openZip>>,
];

// Sharper checks on the shapes a consumer actually reaches into, so the suite
// says more than "these names resolved to some object type".
export type PublicShapesResolve = [
  Expect<Workbook['sheets'] extends readonly unknown[] ? true : false>,
  Expect<ReadOnlyWorkbook['sheetNames'] extends readonly string[] ? true : false>,
  Expect<XmlNode['name'] extends string ? true : false>,
  Expect<Cell extends { value: unknown } ? true : false>,
  Expect<typeof DEFAULT_DECOMPRESSION_LIMITS extends { maxTotalUncompressedBytes: number } ? true : false>,
];
