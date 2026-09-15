# tests/consumer

A stand-in downstream project. `pnpm check:consumer` packs the library, installs
the tarball here, then compiles and runs this code against it. Nothing in this
directory is published, imported by `src/`, or run by vitest.

## Why it exists

The rest of the suite imports `src/` directly, so it never sees the artifact npm
would publish. One class of defect only exists in that artifact: a relative
import inside a shipped `.d.ts` that carries no file extension.

`from './load'` resolves fine for bundlers and for `moduleResolution: bundler`,
so `pnpm build`, `pnpm test` and `pnpm size` all stay green. A consumer on
`moduleResolution: node16` or `nodenext` gets TS2834 instead. That error is
raised inside `node_modules`, where the near-universal `skipLibCheck: true`
discards it, and the declarations it came from degrade to TypeScript's error
type. The consumer's build passes, autocomplete comes back empty, and no type
error fires anywhere in their own code.

Telling consumers to switch to `moduleResolution: bundler` is not a fix. A
project that compiles with plain `tsc` and runs `node dist/server.js` unbundled
is still subject to Node's ESM rules at runtime, so `bundler` would let
TypeScript accept specifiers Node then rejects, trading a contained typing bug
for `ERR_MODULE_NOT_FOUND` in production.

## What runs

`pnpm typecheck` at the repo root resolves with `nodenext`, so an extensionless
relative specifier in `src/` fails there first, with TS2835 naming the fix.
`pnpm check:attw` checks the `exports` map and the shipped declarations from
outside the package. This fixture covers what neither can: a consumer compiling
and then executing against the installed layout.

| Pass | Catches |
| --- | --- |
| `tsc -p tsconfig.<mode>.json` (`skipLibCheck: false`) | Any error inside the shipped `.d.ts` tree, including TS2834. |
| the same with `--skipLibCheck` | Declarations that degraded silently, via `src/no-any.ts`. This is the configuration real consumers use. |
| `node dist/<mode>/main.js` | Node resolving the library for itself, unbundled. TypeScript accepting an import is not evidence that Node will. |

Modes: `node16`, `nodenext`, `bundler`. The runtime pass covers `node16` and
`nodenext`; `bundler` is type-check only, since it has no unbundled Node
equivalent.

## Files

- `src/no-any.ts` has the type-level assertions. They rely on a conditional type
  over `any`, or over the error type TypeScript substitutes for an unresolved
  import, taking both branches and collapsing to `boolean`. `Expect<T extends
  true>` rejects that. `keyof`-based and `0 extends 1 & T` probes both fail to
  detect the error type, which is why the assertions are written structurally.
- `src/main.ts` is the runtime entry. It static-imports one export from each
  published subpath so an `exports` map typo fails here, then round-trips a
  workbook through memory and through a file on disk.
- `tsconfig.base.json` sets `skipLibCheck: false`; the three `tsconfig.<mode>.json`
  files add the resolution mode and an `outDir`.

`package.json` exists mainly so the name differs from `@office-kit/xlsx`.
Without that boundary, Node and TypeScript would both resolve
`@office-kit/xlsx/io` through package self-reference, straight back to the
repo's own `dist/`, and this fixture would never touch the installed tarball.

The install runs through npm, not pnpm, so the fixture gets an ordinary
third-party `node_modules` rather than a workspace link back to `src/`. Both
`node_modules/` and `dist/` here are generated and gitignored.
