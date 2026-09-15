import { describe, expect, it } from 'vitest';
import { serializeFill } from '../../src/drawing/dml/dml-xml.js';
import type { Fill } from '../../src/drawing/dml/fill.js';

describe('issue #54 — malformed Fill kind is not silently dropped', () => {
  it('throws when serializing a Fill with an unknown kind', () => {
    // Caller mistypes `kind: 'solid'` instead of `kind: 'solidFill'`. Previously
    // the serializer silently emitted nothing, producing an empty <c:spPr></c:spPr>
    // that Excel rendered as the default colour.
    const bogus = { kind: 'solid', color: { kind: 'srgb', val: 'FF0000' } } as unknown as Fill;
    expect(() => serializeFill(bogus)).toThrow(/Fill/);
  });
});
