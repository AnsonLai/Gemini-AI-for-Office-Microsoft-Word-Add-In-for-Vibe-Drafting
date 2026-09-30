import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { writeFile } from 'node:fs/promises';
import { cpus } from 'node:os';
import { openDocx, configureLogger } from '@ansonlai/docx-redline-js';
import { zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

configureLogger({ info() {}, warn() {}, error() {} });
function count(name, fallback) {
  const value = process.env[name] ?? String(fallback);
  if (!/^[1-9]\d*$/.test(value)) throw new Error(`${name} must be a positive integer`);
  return Number(value);
}
const iterations = count('DOCX_BENCH_ITERATIONS', 15);
const warmups = count('DOCX_BENCH_WARMUPS', 3);
const operationCount = count('DOCX_BENCH_OPERATIONS', 10);
const paragraphCounts = process.env.DOCX_BENCH_PARAGRAPHS ? [count('DOCX_BENCH_PARAGRAPHS', 100)] : [100, 1000];
const encoder = new TextEncoder();
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

function fixture(paragraphCount) {
  assert.ok(operationCount <= paragraphCount, 'Operations must target distinct paragraphs');
  const text = number => `Clause ${number}: The receiving party shall retain confidential records for thirty days and return all copies on request.`;
  const xml = `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${Array.from({ length: paragraphCount }, (_, i) => `<w:p w14:paraId="${(i + 1).toString(16).padStart(8, '0').toUpperCase()}"><w:r><w:t>${text(i + 1)}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`;
  const operations = Array.from({ length: operationCount }, (_, i) => {
    const number = Math.floor(i * paragraphCount / operationCount) + 1;
    return { type: 'replace', target: { exactText: text(number) }, replacements: [{ find: 'thirty', replace: 'sixty' }] };
  });
  const bytes = zipDocx(new Map([
    ['[Content_Types].xml', encoder.encode(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ['_rels/.rels', encoder.encode(`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', encoder.encode(xml)]
  ]));
  return { xml, operations, bytes };
}

function statistics(samples) {
  const sorted = [...samples].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  const rounded = value => Number(value.toFixed(3));
  return { medianMs: rounded(sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2), p95Ms: rounded(sorted[Math.ceil(sorted.length * .95) - 1]), minMs: rounded(sorted[0]), maxMs: rounded(sorted.at(-1)), samplesMs: samples };
}
async function measure(action) {
  for (let i = 0; i < warmups; i++) await action();
  const samples = [];
  for (let i = 0; i < iterations; i++) {
    const start = performance.now();
    await action();
    samples.push(performance.now() - start);
  }
  return statistics(samples);
}

const results = [];
for (const paragraphs of paragraphCounts) {
  const source = fixture(paragraphs);
  const core = await measure(async () => {
    const result = await applyOperationsToDocumentXml(source.xml, source.operations, 'Benchmark', null, { atomic: true, generateRedlines: true, structuredContent: true, pairReplacements: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.results.length, operationCount);
    assert.ok(result.results.every(item => item.status === 'applied'));
    assert.equal((result.documentXml.match(/>sixty</g) || []).length, operationCount);
  });
  const packageLifecycle = await measure(async () => {
    const doc = openDocx(source.bytes);
    const result = await doc.applyOperations(source.operations, { author: 'Benchmark', atomic: true, generateRedlines: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.ok(doc.toUint8Array().length > 0);
  });
  const openOnly = await measure(() => { assert.ok(openDocx(source.bytes)); });
  const edited = openDocx(source.bytes);
  assert.equal((await edited.applyOperations(source.operations, { author: 'Benchmark', atomic: true, generateRedlines: true })).status, 'ok');
  const saveOnly = await measure(() => { assert.ok(edited.toUint8Array().length > 0); });
  results.push({ paragraphs, operations: operationCount, sourceXmlBytes: encoder.encode(source.xml).length, sourceDocxBytes: source.bytes.length, coreBatch: core, openOnly, saveOnly, openApplySave: packageLifecycle, coreMedianUnder100ms: core.medianMs < 100 });
}
const report = { measuredAt: new Date().toISOString(), nodeVersion: process.version, platform: process.platform, architecture: process.arch, cpuModel: cpus()[0]?.model, warmups, iterations, workload: 'Unique legal-style paragraphs; independent localized thirty-to-sixty replacements; atomic tracked changes', excludes: ['Word getOoxml/insertOoxml and context.sync latency', 'LLM/network/transport latency', 'disk reads/writes'], results };
const outputArg = process.argv.find(arg => arg.startsWith('--output='));
if (outputArg) await writeFile(outputArg.slice('--output='.length), `${JSON.stringify(report, null, 2)}\n`);
console.log(JSON.stringify(report, null, 2));
if (process.argv.includes('--require-sub100ms') && results.some(result => !result.coreMedianUnder100ms)) process.exitCode = 1;
