import { requestGemini } from '../../src/taskpane/modules/chat/gemini-client.js';
/**
 * run-evals.mjs — model-in-the-loop eval harness for the redline diff generator.
 *
 * MANUAL ONLY — this calls the paid Gemini API. It is NOT part of the default
 * regression suite. Usage:
 *
 *   GEMINI_API_KEY=... node tests/evals/run-evals.mjs \
 *       --model gemini-2.5-pro --model gemini-flash-latest [--case <name>]
 *
 * For each (model, case) it:
 *   1. Loads the case fixture .docx and builds `[P#] text` anchored content the
 *      same shape the add-in feeds the diff generator.
 *   2. Calls the model with the SHARED prompt + schema (redline-prompt.js) so the
 *      eval exercises the exact production prompt.
 *   3. Runs the returned change set through the SAME validation pipeline used in
 *      production (sanitizeChangeSet + verifyAnchor) to compute how many changes
 *      would actually be applied vs. rejected. (There is no host-free engine that
 *      applies the index-based change set to a docx, so scoring is done against
 *      the validated change set + its proposed content — which is exactly what
 *      gates application in the add-in.)
 *   4. Scores against the case's structural `expect` block.
 *
 * Exit code 0 only if every requested (model, case) passes.
 */

import './../setup-xml-provider.mjs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import AdmZip from 'adm-zip';
import { parseOoxml, getDocumentParagraphNodes, getParagraphText } from '@ansonlai/docx-redline-js';

import { buildRedlineDiffPrompt, REDLINE_DIFF_SCHEMA } from '../../src/taskpane/modules/commands/redline-prompt.js';
import { sanitizeChangeSet, verifyAnchor, parseAnchoredParagraphs } from '../../src/taskpane/modules/commands/change-validation.js';
import { getModelProfile } from '../../src/taskpane/modules/config/model-profiles.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CASES_DIR = path.join(__dirname, 'cases');

// ---------------------------------------------------------------------------
// Arg parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const models = [];
  const cases = [];
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--model') models.push(argv[++i]);
    else if (argv[i] === '--case') cases.push(argv[++i]);
  }
  return { models, cases };
}

// ---------------------------------------------------------------------------
// Fixture loading
// ---------------------------------------------------------------------------

export function buildAnchoredText(docxPath) {
  const zip = new AdmZip(docxPath);
  const entry = zip.getEntry('word/document.xml');
  if (!entry) throw new Error(`No word/document.xml in ${docxPath}`);
  const xml = entry.getData().toString('utf8');
  const doc = parseOoxml(xml);
  const nodes = getDocumentParagraphNodes(doc);
  const lines = nodes.map((n, i) => `[P${i + 1}] ${getParagraphText(n) || ''}`);
  return lines.join('\n');
}

function loadCases(filterNames) {
  const files = fs.readdirSync(CASES_DIR).filter((f) => f.endsWith('.json'));
  const cases = files.map((f) => {
    const c = JSON.parse(fs.readFileSync(path.join(CASES_DIR, f), 'utf8'));
    c._file = f;
    return c;
  });
  if (filterNames.length === 0) return cases;
  return cases.filter((c) => filterNames.includes(c.name));
}

// ---------------------------------------------------------------------------
// Model call
// ---------------------------------------------------------------------------

async function callGemini(model, apiKey, instruction, anchoredText) {
  const profile = getModelProfile(model);
  const payload = {
    contents: [{ parts: [{ text: buildRedlineDiffPrompt(instruction, anchoredText) }] }],
    generationConfig: {
      temperature: profile.temperature,
      maxOutputTokens: profile.maxOutputTokens,
      responseMimeType: 'application/json',
      responseSchema: REDLINE_DIFF_SCHEMA
    }
  };
  const json = await requestGemini({ model, apiKey, payload });
  const text = json?.candidates?.[0]?.content?.parts?.find((p) => typeof p?.text === 'string')?.text;
  if (!text) throw new Error('No JSON text part in model response');
  return JSON.parse(text);
}

// ---------------------------------------------------------------------------
// Scoring (pure)
// ---------------------------------------------------------------------------

export function scoreChangeSet(expect, rawChanges, paragraphTexts) {
  const { changes: sanitized } = sanitizeChangeSet(rawChanges, paragraphTexts.length, paragraphTexts);
  const applied = sanitized.filter((c) => verifyAnchor(c, paragraphTexts).ok);
  const changesApplied = applied.length;
  const proposed = Array.isArray(rawChanges) ? rawChanges.length : 0;
  const rejected = proposed - changesApplied;

  // Concatenate the proposed document text from the applied changes.
  const resultText = applied
    .map((c) => [c.newContent, c.content, c.replacementText].filter((v) => typeof v === 'string').join('\n'))
    .join('\n');

  const failures = [];
  const e = expect || {};
  if (typeof e.minChangesApplied === 'number' && changesApplied < e.minChangesApplied) {
    failures.push(`changesApplied ${changesApplied} < min ${e.minChangesApplied}`);
  }
  if (typeof e.maxChangesApplied === 'number' && changesApplied > e.maxChangesApplied) {
    failures.push(`changesApplied ${changesApplied} > max ${e.maxChangesApplied}`);
  }
  for (const needle of e.mustContainText || []) {
    if (!resultText.includes(needle)) failures.push(`missing text: ${JSON.stringify(needle)}`);
  }
  for (const needle of e.mustNotContainText || []) {
    if (resultText.includes(needle)) failures.push(`forbidden text present: ${JSON.stringify(needle)}`);
  }

  return { pass: failures.length === 0, changesApplied, rejected, failures };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const { models, cases: caseFilter } = parseArgs(process.argv.slice(2));
  const apiKey = process.env.GEMINI_API_KEY;

  if (!apiKey) {
    console.error('GEMINI_API_KEY is not set. This harness calls the paid Gemini API.');
    process.exit(2);
  }
  if (models.length === 0) {
    console.error('No --model specified. Example: --model gemini-2.5-pro');
    process.exit(2);
  }

  const cases = loadCases(caseFilter);
  if (cases.length === 0) {
    console.error('No matching cases found in tests/evals/cases/.');
    process.exit(2);
  }

  let allPass = true;
  const rows = [];

  for (const model of models) {
    for (const c of cases) {
      const fixturePath = path.resolve(CASES_DIR, c.documentFixture);
      let result;
      try {
        const anchoredText = buildAnchoredText(fixturePath);
        const paragraphTexts = parseAnchoredParagraphs(anchoredText);
        const rawChanges = await callGemini(model, apiKey, c.instruction, anchoredText);
        result = scoreChangeSet(c.expect, rawChanges, paragraphTexts);
      } catch (err) {
        result = { pass: false, changesApplied: '-', rejected: '-', failures: [err.message] };
      }
      if (!result.pass) allPass = false;
      rows.push({
        model,
        name: c.name,
        pass: result.pass,
        changesApplied: result.changesApplied,
        rejected: result.rejected,
        reason: result.failures.join('; ')
      });
    }
  }

  // Print a per-model table.
  console.log('\nmodel | case | pass | applied | rejected | reason');
  console.log('------|------|------|---------|----------|-------');
  for (const r of rows) {
    console.log(`${r.model} | ${r.name} | ${r.pass ? 'PASS' : 'FAIL'} | ${r.changesApplied} | ${r.rejected} | ${r.reason}`);
  }
  console.log(`\n${rows.filter((r) => r.pass).length}/${rows.length} passed.`);

  // Set exitCode (rather than process.exit) and let the event loop drain so we
  // don't race undici's socket teardown on Windows.
  process.exitCode = allPass ? 0 : 1;
}

// Only run when invoked directly (so scoreChangeSet can be unit-tested).
if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  main();
}
