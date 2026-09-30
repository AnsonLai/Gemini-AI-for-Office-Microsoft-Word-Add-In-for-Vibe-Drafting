import assert from 'assert';
import { existsSync, readFileSync } from 'fs';

const retiredModulePaths = [
    'src/taskpane/modules/commands/shared-operation-bridge.js',
    'src/taskpane/modules/docx-redline-js-integration/integration.js',
    'src/taskpane/modules/docx-redline-js-integration/word-route-change.js',
    'src/taskpane/modules/docx-redline-js-integration/word-structured-list.js'
];

const retiredSymbols = [
    'applyReconciliationToParagraph',
    'applyReconciliationToParagraphBatch',
    'shouldUseOoxmlReconciliation',
    'getAuthorForTracking',
    'routeWordParagraphChange',
    'applyStructuredListDirectOoxml'
];

function run() {
    for (const retiredModulePath of retiredModulePaths) {
        assert.strictEqual(
            existsSync(retiredModulePath),
            false,
            `Retired integration module should be removed: ${retiredModulePath}`
        );
    }

    const integrationBarrel = readFileSync(
        'src/taskpane/modules/docx-redline-js-integration/index.js',
        'utf8'
    );
    const agenticTools = readFileSync(
        'src/taskpane/modules/commands/agentic-tools.js',
        'utf8'
    );

    for (const retiredSymbol of retiredSymbols) {
        assert.ok(
            !integrationBarrel.includes(retiredSymbol),
            `Integration barrel should not export retired symbol: ${retiredSymbol}`
        );
        assert.ok(
            !agenticTools.includes(retiredSymbol),
            `Agentic tools should not use retired symbol: ${retiredSymbol}`
        );
    }

    assert.ok(
        !agenticTools.includes('ReconciliationPipeline'),
        'Agentic list edits should use the package reconciliation entry point without a direct pipeline fallback'
    );

    assert.ok(
        agenticTools.includes('const authorName = loadRedlineAuthor();'),
        'Highlight flow should use the injected redline-author provider'
    );

    console.log('PASS: retired integration modules remain absent');
}

run();
