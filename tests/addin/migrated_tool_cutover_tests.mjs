import assert from 'assert';
import fs from 'fs';
import { parse } from 'acorn';

const agenticToolsPath = 'src/taskpane/modules/commands/agentic-tools.js';
const source = fs.readFileSync(agenticToolsPath, 'utf8');
const program = parse(source, { ecmaVersion: 'latest', sourceType: 'module' });

function extractFunctionBody(functionName) {
    const declaration = program.body.map(node => node.type === 'ExportNamedDeclaration' ? node.declaration : node)
        .find(node => node?.type === 'FunctionDeclaration' && node.id?.name === functionName);
    assert.ok(declaration, `Missing function: ${functionName}`);
    return source.slice(declaration.body.start + 1, declaration.body.end - 1);
}

function assertContains(text, token, message) {
    assert.strictEqual(text.includes(token), true, message);
}

function assertNotContains(text, token, message) {
    assert.strictEqual(text.includes(token), false, message);
}

function testRedlineCutover() {
    const body = extractFunctionBody('executeRedline');
    const applyBody = extractFunctionBody('applyRedlineChangeSet');
    assertContains(
        body,
        'applyRedlineChangeSet(',
        'executeRedline should route through applyRedlineChangeSet'
    );
    assertContains(
        applyBody,
        'applyRedlineChangesToWordContext(',
        'applyRedlineChangeSet should route through shared redline runner'
    );
    assertNotContains(
        body,
        'routeChangeOperation(',
        'executeRedline should not route through legacy command-level routeChangeOperation logic'
    );
}

function testCommentCutover() {
    const body = extractFunctionBody('executeComment');
    assertContains(
        body,
        'applyObservedParagraphOperation(',
        'executeComment should route through the observed shared Word operation bridge'
    );
    assertNotContains(
        body,
        'insertComment(',
        'executeComment should not use legacy Word search/insertComment path'
    );
}

function testHighlightCutover() {
    const body = extractFunctionBody('executeHighlight');
    assertContains(
        body,
        'applyObservedParagraphOperation(',
        'executeHighlight should route through the observed shared Word operation bridge'
    );
    assertNotContains(
        body,
        'applyHighlightToOoxml(',
        'executeHighlight should not call legacy local OOXML highlight helper'
    );
}

function testNoLegacyBridgeImport() {
    const observedBody = extractFunctionBody('applyObservedParagraphOperation');
    assertContains(observedBody, 'executePureOoxmlBatch(', 'Observed operations should use the atomic production batch runner');
    assertNotContains(
        source,
        'shared-operation-bridge',
        'agentic tools should not import legacy command shared-operation bridge module'
    );
}

function run() {
    testRedlineCutover();
    testCommentCutover();
    testHighlightCutover();
    testNoLegacyBridgeImport();
    console.log('PASS: migrated tool cutover tests');
}

run();
