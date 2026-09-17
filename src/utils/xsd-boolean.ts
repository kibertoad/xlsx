/**
 * Parse an xsd:boolean, or `undefined` when the text is outside its lexical
 * space. The type's whiteSpace facet is `collapse`, so the padding a
 * pretty-printer leaves around the text is insignificant and a reformatted
 * part reads the same as a compact one. `t` / `f` and mixed case are outside
 * the lexical space proper, but openpyxl accepts them, so the readers here do
 * too rather than rejecting a file another tool opens.
 */
export const parseXsdBoolean = (raw: string | undefined): boolean | undefined => {
  // Excel writes `1` / `0`; answer the common spellings without allocating.
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  if (raw === undefined) return undefined;
  const text = raw.trim().toLowerCase();
  if (text === '1' || text === 'true' || text === 't') return true;
  if (text === '0' || text === 'false' || text === 'f') return false;
  return undefined;
};
