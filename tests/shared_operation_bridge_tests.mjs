import './setup-xml-provider.mjs';

import assert from 'assert';
import { applySharedOperationToParagraphOoxml } from '../src/taskpane/modules/commands/shared-operation-bridge.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function buildParagraphXml(text) {
    return `<w:p xmlns:w="${NS_W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function testHighlightBridge() {
    const paragraphXml = buildParagraphXml('Alpha target span.');
    const result = await applySharedOperationToParagraphOoxml(
        paragraphXml,
        {
            type: 'highlight',
            targetRef: 'P1',
            target: 'Alpha target span.',
            textToHighlight: 'target',
            color: 'yellow'
        },
        {
            author: 'BridgeTest',
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'highlight bridge should report changes');
    assert.ok(result.paragraphOoxml.includes('w:highlight'), 'highlight bridge should return highlighted paragraph OOXML');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('pkg:package'), 'highlight bridge should return package OOXML for Word insertion');
    assert.strictEqual(result.commentsXml, null, 'highlight bridge should not emit comments XML');
}

async function testCommentBridge() {
    const paragraphXml = buildParagraphXml('Comment target span.');
    const result = await applySharedOperationToParagraphOoxml(
        paragraphXml,
        {
            type: 'comment',
            targetRef: 'P1',
            target: 'Comment target span.',
            textToComment: 'target',
            commentContent: 'Please verify this term.'
        },
        {
            author: 'BridgeTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'comment bridge should report changes');
    assert.ok(result.paragraphOoxml.includes('commentRangeStart'), 'comment bridge should return paragraph OOXML with comment markers');
    assert.ok(result.commentsXml && result.commentsXml.includes('Please verify this term.'), 'comment bridge should emit comments XML');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('/word/comments.xml'), 'comment bridge should return package OOXML for Word insertion');
}

async function run() {
    await testHighlightBridge();
    await testCommentBridge();
    console.log('PASS: shared operation bridge tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
