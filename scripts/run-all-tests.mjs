import { readdirSync } from 'node:fs';
import { execFile } from 'node:child_process';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const exclusions = new Map([
  ['tests/setup-xml-provider.mjs', 'XML provider setup module; exercised by suites'],
  ['tests/evals/run-evals.mjs', 'Live Gemini API evaluation; requires credentials and incurs provider usage'],
  ['tests/word-desktop/list-regression.mjs', 'Word Desktop fixture lane; run with the documented PowerShell harness'],
  ['tests/phase4/perf-harness.mjs', 'Observational performance lane; use npm run benchmark:ooxml']
]);

function discover(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    const full = resolve(directory, entry.name);
    if (entry.isDirectory()) return discover(full);
    return /\.(mjs|js)$/.test(entry.name) ? [relative(root, full).replaceAll('\\', '/')] : [];
  });
}

function positiveInteger(name, fallback) {
  const value = process.env[name] ?? String(fallback);
  if (!/^[1-9]\d*$/.test(value)) throw new Error(`${name} must be a positive integer`);
  return Number(value);
}

const timeout = positiveInteger('DOCX_TEST_TIMEOUT', 180000);
const concurrency = positiveInteger('DOCX_TEST_CONCURRENCY', 4);
const entries = [...discover(resolve(root, 'tests')), ...discover(resolve(root, 'mcp/docx-server/tests'))].sort();
if (process.argv.includes('--list')) {
  for (const file of entries) console.log(`${exclusions.has(file) ? 'SKIP' : 'RUN '} ${file}${exclusions.has(file) ? `: ${exclusions.get(file)}` : ''}`);
} else {
  const suites = entries.filter(file => !exclusions.has(file));
  const results = new Array(suites.length);
  let next = 0;
  async function worker() {
    while (next < suites.length) {
      const index = next++;
      const file = suites[index];
      results[index] = await new Promise(done => {
        const args = [resolve(root, file), ...(file === 'tests/phase4/golden-guardrail.mjs' ? ['--verify'] : [])];
        execFile(process.execPath, args, { cwd: root, timeout, maxBuffer: 4 * 1024 * 1024, windowsHide: true }, (error, stdout = '', stderr = '') => {
          const failureMarker = /(?:❌\s*(?:FAIL|FAILED|FAILURE)|\bTEST FAILED\b|\bINTEGRATION TEST FAILURE\b|^FAIL:)/im.test(`${stdout}\n${stderr}`);
          const skips = `${stdout}\n${stderr}`.split(/\r?\n/).filter(line => /\bSKIP(?:PED)?\b/.test(line));
          const knownDefects = stdout.split(/\r?\n/).filter(line => line.startsWith('KNOWN_LIBRARY_DEFECT:'));
          done({ file, error, stdout, stderr, failureMarker, skips, knownDefects });
        });
      });
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, suites.length) }, worker));
  let failed = 0;
  for (const result of results) {
    const pass = !result.error && !result.failureMarker;
    if (!pass) failed++;
    console.log(`${pass ? 'PASS' : 'FAIL'} ${result.file}${result.skips.length ? ' (contains reported skips)' : ''}`);
    for (const skip of result.skips) console.log(`  ${skip.trim()}`);
    for (const defect of result.knownDefects) console.log(`  ${defect.trim()}`);
    if (!pass) console.error([result.stdout, result.stderr, result.error?.message].filter(Boolean).join('\n'));
  }
  for (const file of entries.filter(file => exclusions.has(file))) console.log(`SKIP ${file}: ${exclusions.get(file)}`);
  console.log(`\n${results.length - failed} suites passed, ${failed} failed; ${entries.length - suites.length} entrypoints excluded. Suite-internal skips are reported above and are not counted as passes.`);
  if (failed) process.exitCode = 1;
}
