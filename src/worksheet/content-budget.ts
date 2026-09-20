// Caps on the content a read will model, and the running totals that enforce
// them. The companion to the zip layer's decompression guard: that one bounds
// the bytes an archive inflates to, this one bounds the cells those bytes turn
// into, which is what decides how long a read spends modelling them and how
// large the model it hands back gets.

import { formatSheetQualifiedRef, tupleToCoordinate } from '../utils/coordinate.js';
import { OpenXmlContentLimitError, OpenXmlError } from '../utils/exceptions.js';

/**
 * Caps on the content one read will model. Both are unlimited when absent, so
 * a caller that asks for nothing gets what it always got.
 */
export interface ContentLimits {
  /** Maximum cells the read will place before it refuses to continue. */
  maxCells?: number;
  /** Maximum rows the read will place before it refuses to continue. */
  maxRows?: number;
}

/** {@link ContentLimits} with every field settled; `Infinity` is "unlimited". */
export interface ResolvedContentLimits {
  readonly maxCells: number;
  readonly maxRows: number;
}

/** What a caller that passed no `contentLimits` gets. */
export const UNLIMITED_CONTENT_LIMITS: ResolvedContentLimits = {
  maxCells: Number.POSITIVE_INFINITY,
  maxRows: Number.POSITIVE_INFINITY,
};

// A cap has to be a positive integer to mean anything: a NaN, a negative or a
// zero turns `count > cap` into a gate that either never opens or never
// closes. Reject those where the caller hands them over rather than at the
// first cell, where the message would name a cell instead of the mistake.
const requirePositiveInteger = (field: string, value: number): void => {
  if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
    throw new OpenXmlError(`contentLimits.${field} must be a positive integer; got ${String(value)}`);
  }
};

/**
 * Validate what the caller passed and settle the absent fields.
 *
 * Kept apart from {@link makeContentBudget} because the two happen at
 * different times: an entry point validates once, where the caller can see the
 * error, while each pass over the content builds its own counters from the
 * result.
 */
export function resolveContentLimits(input: ContentLimits | undefined): ResolvedContentLimits {
  if (input === undefined) return UNLIMITED_CONTENT_LIMITS;
  if (input.maxCells !== undefined) requirePositiveInteger('maxCells', input.maxCells);
  if (input.maxRows !== undefined) requirePositiveInteger('maxRows', input.maxRows);
  return {
    maxCells: input.maxCells ?? Number.POSITIVE_INFINITY,
    maxRows: input.maxRows ?? Number.POSITIVE_INFINITY,
  };
}

/** Running totals for one pass over the content. */
export interface ContentBudget {
  readonly limits: ResolvedContentLimits;
  cells: number;
  rows: number;
}

/** Fresh counters against already-resolved limits. */
export function makeContentBudget(limits: ResolvedContentLimits): ContentBudget {
  return { limits, cells: 0, rows: 0 };
}

/**
 * Charge one cell to the budget. Called before the cell is built, so a read
 * that is going to be refused does not first allocate what it is refusing.
 *
 * Takes the coordinate as a column and a row rather than a formatted ref: this
 * runs once per cell of the sheet, and the ref is only needed on the one call
 * that throws.
 */
export function chargeCell(budget: ContentBudget, sheet: string, col: number, row: number): void {
  budget.cells++;
  if (budget.cells > budget.limits.maxCells) {
    const at = formatSheetQualifiedRef(sheet, tupleToCoordinate(col, row));
    throw new OpenXmlContentLimitError(
      `worksheet: reading ${at} passes contentLimits.maxCells of ${budget.limits.maxCells}`,
    );
  }
}

/** Charge one row to the budget. */
export function chargeRow(budget: ContentBudget, sheet: string, row: number): void {
  budget.rows++;
  if (budget.rows > budget.limits.maxRows) {
    throw new OpenXmlContentLimitError(
      `worksheet: reading row ${row} of ${sheet} passes contentLimits.maxRows of ${budget.limits.maxRows}`,
    );
  }
}
