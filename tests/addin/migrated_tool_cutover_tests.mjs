import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { parse } from 'acorn';

const read = path => readFileSync(path, 'utf8');
const typedPaths = [
  'src/taskpane/taskpane.js',
  'src/taskpane/modules/chat/document-agent-loop.js',
  'src/taskpane/modules/chat/tool-registry.js',
  'src/taskpane/modules/chat/document-tools.js',
  'src/taskpane/modules/chat/document-edit-intents.js',
  'src/taskpane/modules/chat/document-inspection-session.js',
  'src/taskpane/modules/chat/word-document-source.js',
];
const retiredImports = /(?:commands\/)?(?:agentic-tools|table-request-validation|mutation-outcome)\.js|utils\/(?:markdown-utils|text-statistics)\.js/;
const noProviderPaths = typedPaths.slice(1);
const purePaths = [
  'src/taskpane/modules/chat/document-agent-loop.js',
  'src/taskpane/modules/chat/tool-registry.js',
  'src/taskpane/modules/chat/document-edit-intents.js',
  'src/taskpane/modules/chat/document-inspection-session.js',
];

for (const path of typedPaths) {
  const source = read(path);
  const imports = parse(source, { ecmaVersion: 'latest', sourceType: 'module' }).body
    .filter(node => node.type === 'ImportDeclaration')
    .map(node => node.source.value);
  assert.equal(source.includes('modules/commands/agentic-tools.js'), false,
    `${path}: the retired general executor must not be in the typed production path`);
  assert.equal(imports.some(specifier => retiredImports.test(specifier)), false,
    `${path}: obsolete executor helpers must not be imported by the typed production path`);
}

for (const path of noProviderPaths) {
  const source = read(path);
  const imports = parse(source, { ecmaVersion: 'latest', sourceType: 'module' }).body
    .filter(node => node.type === 'ImportDeclaration')
    .map(node => node.source.value);
  assert.equal(imports.some(specifier => /gemini|generative-language|googleapis/i.test(specifier)), false,
    `${path}: typed document tools receive the model request through the injected boundary`);
  assert.doesNotMatch(source, /\bfetch\s*\(/, `${path}: provider transport belongs to the shared client`);
}

for (const path of purePaths) {
  const source = read(path);
  assert.doesNotMatch(source, /\bWord\.run\s*\(|\bOffice\./,
    `${path}: pure dispatch/compiler/session code must not call Word or Office.js`);
  assert.doesNotMatch(source, /\.insertParagraph\s*\(|\.insertText\s*\(|\.insertOoxml\s*\(|\.addRows\s*\(/,
    `${path}: document mutations belong to the document-tools host boundary`);
}

const documentTools = read('src/taskpane/modules/chat/document-tools.js');
assert.doesNotMatch(documentTools, /\.insertParagraph\s*\(|\.insertText\s*\(|\.addRows\s*\(|\.insertComment\s*\(/,
  'typed edits must not use the retired per-object native Word mutation routes');
const sourceReader = read('src/taskpane/modules/chat/word-document-source.js');
assert.doesNotMatch(sourceReader, /\.insertParagraph\s*\(|\.insertText\s*\(|\.insertOoxml\s*\(/,
  'the Office.js source reader must remain read-only');

console.log('PASS: typed document path excludes retired executors, provider bypasses and per-object native edits');
