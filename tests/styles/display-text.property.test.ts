// Invariants that have to hold over arbitrary values, rather than the specific
// rows in display-text.test.ts.

import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { getCellDate, getCellDisplayText, setCellNumberFormat } from '../../src/styles/index.js';
import { dateToExcel } from '../../src/utils/datetime.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/index.js';
import { setCell } from '../../src/worksheet/index.js';

const displayNumber = (code: string, value: number, date1904 = false): string => {
  const wb = createWorkbook({ date1904 });
  const ws = addWorksheet(wb, 'Sheet1');
  const cell = setCell(ws, 1, 1, value);
  setCellNumberFormat(wb, cell, code);
  return getCellDisplayText(wb, cell);
};

const MAX_DECIMALS = 6;
/**
 * Excel carries 15 significant digits and the renderer follows it, so the
 * generators stay inside that budget: a wider value is rounded on purpose and
 * would not round-trip.
 */
const MAX_MAGNITUDE = 1e6;
const MAX_INTEGER_15_DIGITS = 999_999_999_999_999;

describe('fixed-decimal codes round-trip through Number()', () => {
  it('prints exactly the requested decimals, and parses back to the rounded value', () => {
    fc.assert(
      fc.property(
        fc.double({ min: -MAX_MAGNITUDE, max: MAX_MAGNITUDE, noNaN: true, noDefaultInfinity: true }),
        fc.integer({ min: 0, max: MAX_DECIMALS }),
        (value, decimals) => {
          const code = decimals === 0 ? '0' : `0.${'0'.repeat(decimals)}`;
          const text = displayNumber(code, value);
          const dot = text.indexOf('.');
          expect(dot).toBe(decimals === 0 ? -1 : text.length - decimals - 1);
          // Rounding is half away from zero at `decimals` places, so the text
          // can never sit further from the value than half of the last place.
          expect(Math.abs(Number(text) - value)).toBeLessThanOrEqual(0.5 * 10 ** -decimals);
        },
      ),
    );
  });
});

describe('thousands grouping keeps the digits it groups', () => {
  it('#,##0 prints the integer part with separators and nothing else', () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: MAX_INTEGER_15_DIGITS }), (value) => {
        const text = displayNumber('#,##0', value);
        expect(text.replaceAll(',', '')).toBe(String(value));
        const groups = text.split(',');
        expect(groups.slice(1).every((group) => group.length === 3)).toBe(true);
        expect(groups[0]?.length).toBeGreaterThanOrEqual(1);
      }),
    );
  });
});

/** The coarsest denominator below is 16, so half a step of it bounds them all. */
const FRACTION_TOLERANCE = 1 / 32;

/** Read back `2  3/4 `, `1/02` or a blanked `3     ` as a number. */
const fractionValue = (text: string): number => {
  const slash = text.indexOf('/');
  if (slash === -1) return Number(text.trim());
  const left = text
    .slice(0, slash)
    .trim()
    .split(/\s+/)
    .filter((piece) => piece.length > 0);
  const numerator = Number(left[left.length - 1]);
  const whole = left.length > 1 ? Number(left[0]) : 0;
  return whole + numerator / Number(text.slice(slash + 1).trim());
};

describe('a rendered fraction is worth what the cell holds', () => {
  it('reads back to within half a step of the denominator the code allows', () => {
    fc.assert(
      fc.property(
        fc.double({ min: 0, max: MAX_MAGNITUDE, noNaN: true, noDefaultInfinity: true }),
        fc.constantFrom('# ??/??', '# ??/00', '?/00', '# ?/16'),
        (value, code) => {
          expect(Math.abs(fractionValue(displayNumber(code, value)) - value)).toBeLessThanOrEqual(FRACTION_TOLERANCE);
        },
      ),
    );
  });
});

describe('the rendered date agrees with getCellDate', () => {
  it('yyyy-mm-dd hh:mm:ss prints the fields of the Date the reader hands back', () => {
    fc.assert(
      fc.property(
        // From 1904 so the serial stays non-negative under either epoch.
        fc.date({ min: new Date(Date.UTC(1904, 0, 1)), max: new Date(Date.UTC(9999, 11, 31)), noInvalidDate: true }),
        fc.boolean(),
        (date, date1904) => {
          const wb = createWorkbook({ date1904 });
          const ws = addWorksheet(wb, 'Sheet1');
          // Excel's display precision under this code is one second.
          const wholeSecond = new Date(Math.floor(date.getTime() / 1000) * 1000);
          const serial = dateToExcel(wholeSecond, { epoch: date1904 ? 'mac' : 'windows' });
          const cell = setCell(ws, 1, 1, serial);
          setCellNumberFormat(wb, cell, 'yyyy-mm-dd hh:mm:ss');

          const read = getCellDate(wb, cell);
          expect(read).toBeDefined();
          if (read === undefined) return;
          const pad = (n: number, width = 2): string => String(n).padStart(width, '0');
          expect(getCellDisplayText(wb, cell)).toBe(
            `${pad(read.getUTCFullYear(), 4)}-${pad(read.getUTCMonth() + 1)}-${pad(read.getUTCDate())} ` +
              `${pad(read.getUTCHours())}:${pad(read.getUTCMinutes())}:${pad(read.getUTCSeconds())}`,
          );
        },
      ),
    );
  });
});
