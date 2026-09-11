import assert from 'assert';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';

const PARAGRAPH_OOXML = '<w:p><w:r><w:t>Mock List Output</w:t></w:r></w:p>';

function runCase(label, includeNumbering, expectedHasNumberingRelationship) {
    const packageOoxml = buildDocumentFragmentPackage(PARAGRAPH_OOXML, {
        includeNumbering,
        numberingXml: null
    });
    const hasNumberingRelationship = packageOoxml.includes('/relationships/numbering');
    assert.strictEqual(
        hasNumberingRelationship,
        expectedHasNumberingRelationship,
        `${label}: unexpected numbering relationship presence`
    );
}

function run() {
    runCase('includeNumbering=false', false, false);
    runCase('includeNumbering=true', true, true);
    // Numbering packaging remains opt-in when the pipeline does not explicitly
    // request it.
    runCase('includeNumbering=undefined', undefined, false);
    console.log('PASS: includeNumbering behavior');
}

run();

