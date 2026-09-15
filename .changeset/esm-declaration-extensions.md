---
'@office-kit/xlsx': patch
---

fix: types were silently lost on `moduleResolution: node16` / `nodenext`

Relative imports inside the shipped `.d.ts` files had no file extension
(`from './load'`), which Node's ESM rules reject. Consumers on
`moduleResolution: node16` or `nodenext` got TS2834 inside `node_modules`,
where the usual `skipLibCheck: true` discarded it, so every symbol imported
from `@office-kit/xlsx/*` degraded to an error type: no autocomplete, and no
type errors reported against the library's API.

The declarations now carry `.js` extensions, so all of `node16`, `nodenext` and
`bundler` resolve the full type graph with `skipLibCheck: false`. No runtime
behaviour, export name or type signature changed.

`moduleResolution: node10` remains unsupported, since the subpaths are declared
only through `exports`.
