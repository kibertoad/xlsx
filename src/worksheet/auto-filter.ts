// AutoFilter.
//
// **Stage 1**: ref + filterColumns where each entry is the `kind: 'filters'`
// variant — the value-list dropdown filter that covers >95% of real-world
// spreadsheets. customFilters / top10 / dynamicFilter / colorFilter /
// iconFilter / SortState are reserved for later iterations.

import type { RangeRef } from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { parseRange, rangeToString } from './cell-range.js';

export type FilterColumn = {
  kind: 'filters';
  colId: number;
  /** Discrete values that pass the filter. Stored as strings to match the wire format. */
  values: string[];
  /** Whether blanks are visible. */
  blank?: boolean;
};

export interface AutoFilter {
  /** Excel range the filter covers (`"A1:E100"`), as stored on the wire. */
  ref: string;
  filterColumns: FilterColumn[];
}

/**
 * `ref` takes numeric bounds as well as an A1 string, matching the range-taking
 * helpers elsewhere. A string is validated and then stored as written, since
 * `"A:E"` is a legal filter ref that normalising would expand to the full grid.
 */
export function makeAutoFilter(opts: { ref: RangeRef; filterColumns?: FilterColumn[] }): AutoFilter {
  return { ref: resolveFilterRef(opts.ref), filterColumns: opts.filterColumns ?? [] };
}

const resolveFilterRef = (ref: RangeRef): string => {
  if (typeof ref !== 'string') return rangeToString(parseRange(ref));
  // Parse to validate, then keep the caller's spelling: "A:E" is a legal filter
  // ref and the normalised form would name all 1 048 576 rows of each column.
  parseRange(ref);
  return ref;
};

export function makeFilterColumn(opts: {
  colId: number;
  values: ReadonlyArray<string>;
  blank?: boolean;
}): FilterColumn {
  return {
    kind: 'filters',
    colId: opts.colId,
    values: [...opts.values],
    ...(opts.blank !== undefined ? { blank: opts.blank } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet.js';

/** Add an AutoFilter dropdown header strip to the given range. */
export const addAutoFilter = (ws: Worksheet, ref: RangeRef): AutoFilter => {
  ws.autoFilter = makeAutoFilter({ ref });
  return ws.autoFilter;
};

/**
 * Add a value-list dropdown filter to a column inside the existing AutoFilter
 * range. `colId` is 0-based relative to the AutoFilter left edge.
 */
export const addAutoFilterColumn = (
  ws: Worksheet,
  colId: number,
  values: ReadonlyArray<string>,
  opts: { blank?: boolean } = {},
): FilterColumn => {
  if (!ws.autoFilter) {
    throw new OpenXmlSchemaError('addAutoFilterColumn: call addAutoFilter(ws, ref) first');
  }
  const fc = makeFilterColumn({ colId, values, ...(opts.blank !== undefined ? { blank: opts.blank } : {}) });
  ws.autoFilter.filterColumns.push(fc);
  return fc;
};

/** Drop the worksheet's AutoFilter entirely. */
export const removeAutoFilter = (ws: Worksheet): void => {
  delete (ws as { autoFilter?: AutoFilter }).autoFilter;
};
