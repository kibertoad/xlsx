// Pack the library, install the tarball into tests/consumer, and check that a
// downstream project can consume it under Node's own ESM rules.
//
// `pnpm typecheck` already catches a missing `.js` on a relative specifier in
// src/, because the root tsconfig resolves with `nodenext`, and `pnpm
// check:attw` checks the `exports` map from the outside. What is left, and
// what this covers, is a real consumer compiling and then running against the
// published layout.
//
// Three passes per resolution mode, each catching something the others miss:
//   1. `skipLibCheck: false`: surfaces errors inside node_modules that a real
//      consumer would have suppressed.
//   2. `skipLibCheck: true`: the near-universal consumer setting, where a
//      broken declaration graph degrades silently instead of erroring.
//      tests/consumer/src/no-any.ts is what makes this pass fail.
//   3. plain `tsc` emit, then `node dist/<mode>/main.js`: Node resolving the
//      library for itself, unbundled. TypeScript accepting an import is not
//      evidence that Node will; that gap is a published ERR_MODULE_NOT_FOUND.

import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const fixtureDir = join(repoRoot, 'tests', 'consumer');
const tsc = join(repoRoot, 'node_modules', 'typescript', 'bin', 'tsc');

const failures = [];

function section(title) {
  console.log(`\n=== ${title} ===`);
}

function record(label, result) {
  if (result.error) throw result.error;
  if (result.status !== 0) {
    failures.push(`${label} (exit ${result.status})`);
    return false;
  }
  return true;
}

function runNode(label, args) {
  return record(label, spawnSync(process.execPath, args, { stdio: 'inherit', cwd: repoRoot }));
}

// `npm` is a shell script on POSIX and a `.cmd` shim on Windows, so it can
// only be spawned through a shell. Node 24 deprecates `shell: true` alongside
// an argv array (DEP0190), so these calls pass one pre-quoted command string.
function quote(value) {
  return process.platform === 'win32' ? `"${value}"` : `'${value}'`;
}

function runShell(label, command, options = {}) {
  return record(
    label,
    spawnSync(command, { stdio: 'inherit', cwd: repoRoot, shell: true, ...options }),
  );
}

function pack(destination) {
  const result = spawnSync(`npm pack --json --pack-destination ${quote(destination)}`, {
    cwd: repoRoot,
    encoding: 'utf8',
    shell: true,
  });
  if (result.status !== 0) throw new Error(`npm pack failed:\n${result.stderr}`);
  // npm prints notices on stdout ahead of the JSON payload.
  const [{ filename }] = JSON.parse(result.stdout.slice(result.stdout.indexOf('[')));
  return join(destination, filename);
}

if (!existsSync(join(repoRoot, 'dist', 'io', 'index.d.ts'))) {
  console.error('dist/ is missing or has no declarations. Run `pnpm build` first.');
  process.exit(1);
}

const staging = mkdtempSync(join(tmpdir(), 'xlsx-pack-'));
try {
  section('pack');
  const tarball = pack(staging);
  console.log(tarball);

  section('install tarball into tests/consumer');
  rmSync(join(fixtureDir, 'node_modules'), { recursive: true, force: true });
  rmSync(join(fixtureDir, 'dist'), { recursive: true, force: true });
  rmSync(join(fixtureDir, 'package-lock.json'), { force: true });
  // npm rather than pnpm: the fixture has to look like an ordinary third-party
  // install of the tarball, with the library's own runtime deps resolved for
  // it, not a workspace link back to src/.
  runShell(
    'npm install',
    `npm install --no-save --no-package-lock --no-audit --no-fund ${quote(tarball)}`,
    { cwd: fixtureDir },
  );

  for (const mode of ['node16', 'nodenext', 'bundler']) {
    const project = join(fixtureDir, `tsconfig.${mode}.json`);

    section(`tsc --moduleResolution ${mode} (skipLibCheck: false)`);
    runNode(`${mode} typecheck with skipLibCheck: false`, [tsc, '-p', project]);

    section(`tsc --moduleResolution ${mode} (skipLibCheck: true)`);
    runNode(`${mode} typecheck with skipLibCheck: true`, [
      tsc,
      '-p',
      project,
      '--skipLibCheck',
      '--noEmit',
    ]);
  }

  for (const mode of ['node16', 'nodenext']) {
    section(`node dist/${mode}/main.js`);
    runNode(`${mode} runtime`, [join(fixtureDir, 'dist', mode, 'main.js')]);
  }
} finally {
  rmSync(staging, { recursive: true, force: true });
}

if (failures.length > 0) {
  console.error(`\n${failures.length} packaging check(s) failed:`);
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}
console.log('\nAll packaging checks passed.');
