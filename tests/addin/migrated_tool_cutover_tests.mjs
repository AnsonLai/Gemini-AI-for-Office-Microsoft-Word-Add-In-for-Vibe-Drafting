import assert from 'assert';
import fs from 'fs';

const agenticToolsPath = 'src/taskpane/modules/commands/agentic-tools.js';
const source = fs.readFileSync(agenticToolsPath, 'utf8');

function extractFunctionBody(functionName) {
    const marker = `async function ${functionName}(`;
    const start = source.indexOf(marker);
    assert.notStrictEqual(start, -1, `Missing function: ${functionName}`);

    const openBrace = source.indexOf('{', start);
    assert.notStrictEqual(openBrace, -1, `Missing opening brace for: ${functionName}`);

    let depth = 0;
    for (let i = openBrace; i < source.length; i += 1) {
        const ch = source[i];
        if (ch === '{') depth += 1;
        if (ch === '}') {
            depth -= 1;
            if (depth === 0) {
                return source.slice(openBrace + 1, i);
            }
        }
    }

    throw new Error(`Missing closing brace for: ${functionName}`);
}

function assertContains(text, token, message) {
    assert.strictEqual(text.includes(token), true, message);
}

function assertNotContains(text, token, message) {
    assert.strictEqual(text.includes(token), false, message);
}

function testRedlineCutover() {
    const body = extractFunctionBody('executeRedline');
    assertContains(
        body,
        'applyRedlineChangesToWordContext(',
        'executeRedline should route through shared redline runner'
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
        'applySharedOperationToWordParagraph(',
        'executeComment should route through shared word operation bridge'
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
        'applySharedOperationToWordParagraph(',
        'executeHighlight should route through shared word operation bridge'
    );
    assertNotContains(
        body,
        'applyHighlightToOoxml(',
        'executeHighlight should not call legacy local OOXML highlight helper'
    );
}

function testNoLegacyBridgeImport() {
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
