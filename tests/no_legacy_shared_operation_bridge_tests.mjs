import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';

const retiredModulePaths = [
  'src/taskpane/modules/commands/shared-operation-bridge.js',
  'src/taskpane/modules/commands/agentic-tools.js',
  'src/taskpane/modules/commands/mutation-outcome.js',
  'src/taskpane/modules/commands/table-request-validation.js',
  'src/taskpane/modules/commands/redline-operation-converter.js',
  'src/taskpane/modules/docx-redline-js-integration/integration.js',
  'src/taskpane/modules/docx-redline-js-integration/word-route-change.js',
  'src/taskpane/modules/docx-redline-js-integration/word-structured-list.js',
  'src/taskpane/modules/utils/markdown-utils.js',
  'src/taskpane/modules/utils/text-statistics.js',
];

const retiredSymbols = [
  'applyReconciliationToParagraph',
  'applyReconciliationToParagraphBatch',
  'shouldUseOoxmlReconciliation',
  'getAuthorForTracking',
  'routeWordParagraphChange',
  'applyStructuredListDirectOoxml',
];

const typedPathFiles = [
  'src/taskpane/taskpane.js',
  'src/taskpane/modules/chat/document-agent-loop.js',
  'src/taskpane/modules/chat/tool-registry.js',
  'src/taskpane/modules/chat/document-tools.js',
  'src/taskpane/modules/chat/document-edit-intents.js',
  'src/taskpane/modules/chat/document-inspection-session.js',
  'src/taskpane/modules/chat/word-document-source.js',
];

for (const retiredModulePath of retiredModulePaths) {
  assert.equal(existsSync(retiredModulePath), false, `Retired module should be absent: ${retiredModulePath}`);
}

const integrationBarrel = readFileSync('src/taskpane/modules/docx-redline-js-integration/index.js', 'utf8');
const documentTools = readFileSync('src/taskpane/modules/chat/document-tools.js', 'utf8');
const editIntents = readFileSync('src/taskpane/modules/chat/document-edit-intents.js', 'utf8');
for (const retiredSymbol of retiredSymbols) {
  assert.equal(integrationBarrel.includes(retiredSymbol), false,
    `Integration barrel should not export retired symbol: ${retiredSymbol}`);
  assert.equal(documentTools.includes(retiredSymbol), false,
    `Typed document tools should not use retired symbol: ${retiredSymbol}`);
  assert.equal(editIntents.includes(retiredSymbol), false,
    `Typed edit compiler should not use retired symbol: ${retiredSymbol}`);
}

assert.equal(documentTools.includes('ReconciliationPipeline'), false,
  'typed document edits must use the public library operation path without a local pipeline fallback');
assert.match(documentTools, /getAuthor\(\)/,
  'typed edits receive the tracked-change author through their request-scoped dependency');
for (const path of typedPathFiles) {
  const source = readFileSync(path, 'utf8');
  assert.equal(source.includes('modules/commands/agentic-tools.js'), false,
    `${path} must not import the retired general executor`);
}

console.log('PASS: retired executor/bridge modules are absent from the typed document path');
