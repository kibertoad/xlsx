import { describe, expect, it } from 'vitest';
import { fromTree } from '../../src/schema/serialize.js';
import { makeSide } from '../../src/styles/borders.js';
import { SideSchema } from '../../src/styles/borders.schema.js';
import { parseXml } from '../../src/xml/parser.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';

// https://github.com/office-kit/xlsx/issues/99
// OnlyOffice writes explicit `<left style="none"/>` sides. `none` is the first
// value of ECMA-376 §18.18.3 ST_BorderStyle, so it must parse, not throw
// "expected one of [...]; got \"none\"".
describe('issue #99: border style "none"', () => {
  it('parses a side with style="none" from XML', () => {
    const xml = `<side xmlns="${SHEET_MAIN_NS}" style="none"/>`;
    const side = fromTree(parseXml(xml), SideSchema);
    expect(side.style).toBe('none');
  });

  it('makeSide accepts style="none"', () => {
    expect(makeSide({ style: 'none' }).style).toBe('none');
  });
});
