import './setup-xml-provider.mjs';

import assert from 'assert';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { ReconciliationPipeline } from '../src/taskpane/modules/reconciliation/pipeline/pipeline.js';

const ORIGINAL_OOXML = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Seed</w:t></w:r></w:p>';
const ORIGINAL_TEXT = 'Seed';
const LIST_TEXT = '1. First item\n2. Second item';

async function runCase(label, includeNumbering, expectedHasNumberingRelationship) {
    const originalExecute = ReconciliationPipeline.prototype.execute;
    ReconciliationPipeline.prototype.execute = async function mockedExecute() {
        return {
            ooxml: '<w:p><w:r><w:t>Mock List Output</w:t></w:r></w:p>',
            isValid: true,
            warnings: [],
            includeNumbering,
            numberingXml: null
        };
    };

    try {
        const result = await applyRedlineToOxml(ORIGINAL_OOXML, ORIGINAL_TEXT, LIST_TEXT, {
            author: 'TestUser',
            generateRedlines: false
        });

        assert(result.hasChanges, `${label}: expected changes`);
        const hasNumberingRelationship = result.oxml.includes('/relationships/numbering');
        assert.strictEqual(
            hasNumberingRelationship,
            expectedHasNumberingRelationship,
            `${label}: unexpected numbering relationship presence`
        );
    } finally {
        ReconciliationPipeline.prototype.execute = originalExecute;
    }
}

async function run() {
    await runCase('includeNumbering=false', false, false);
    await runCase('includeNumbering=true', true, true);
    await runCase('includeNumbering=undefined', undefined, true);
    console.log('PASS: includeNumbering behavior');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
