import './setup-xml-provider.mjs';

import assert from 'assert';

import {
    setDefaultAuthor,
    getDefaultAuthor,
    setPlatform,
    getPlatform,
    applyRedlineToOxml
} from '@gsd/docx-reconciliation/index.js';
import { createRevisionMetadata, resetRevisionIdCounter } from '@gsd/docx-reconciliation/core/types.js';
import { injectCommentsIntoOoxml, resetRevisionIdCounter as resetCommentRevisionIdCounter } from '@gsd/docx-reconciliation/services/comment-engine.js';
import { ReconciliationPipeline } from '@gsd/docx-reconciliation/pipeline/pipeline.js';

function testCreateRevisionMetadataUsesConfiguredDefaultAuthor() {
    resetRevisionIdCounter(1000);
    setDefaultAuthor('Configured Author');
    const metadata = createRevisionMetadata();
    assert.strictEqual(metadata.author, 'Configured Author');
}

async function testApplyRedlineUsesConfiguredDefaultAuthor() {
    setDefaultAuthor('Configured Author');
    const originalOoxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:t>Hello world</w:t></w:r>
        </w:p>
    `;

    const result = await applyRedlineToOxml(originalOoxml, 'Hello world', 'Hello brave world');
    assert.strictEqual(result.hasChanges, true);
    assert.match(result.oxml, /w:author="Configured Author"/);
}

function testCommentEngineUsesConfiguredDefaultAuthor() {
    setDefaultAuthor('Configured Author');
    resetCommentRevisionIdCounter(1000);
    const originalOoxml = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>Hello world</w:t></w:r></w:p>
            </w:body>
        </w:document>
    `;
    const result = injectCommentsIntoOoxml(originalOoxml, [
        { paragraphIndex: 1, textToFind: 'Hello', commentContent: 'Check this' }
    ]);
    assert.strictEqual(result.commentsApplied, 1);
    assert.match(result.commentsXml || '', /w:author="Configured Author"/);
}

function testPipelineUsesConfiguredPlatformDefaults() {
    setPlatform('OfficeOnline');
    assert.strictEqual(getPlatform(), 'OfficeOnline');
    const webPipeline = new ReconciliationPipeline();
    assert.strictEqual(webPipeline.platform, 'OfficeOnline');
    assert.strictEqual(webPipeline.isWebPlatform, true);

    setPlatform('Win32');
    const winPipeline = new ReconciliationPipeline();
    assert.strictEqual(winPipeline.platform, 'Win32');
    assert.strictEqual(winPipeline.isWebPlatform, false);
}

async function run() {
    try {
        testCreateRevisionMetadataUsesConfiguredDefaultAuthor();
        await testApplyRedlineUsesConfiguredDefaultAuthor();
        testCommentEngineUsesConfiguredDefaultAuthor();
        testPipelineUsesConfiguredPlatformDefaults();
        console.log('PASS: reconciliation author/platform defaults');
    } finally {
        setDefaultAuthor('Author');
        setPlatform('Unknown');
    }
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});


