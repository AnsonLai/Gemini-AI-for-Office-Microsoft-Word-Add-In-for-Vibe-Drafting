import assert from 'assert';
import {
    applySubstringSearchReplace,
    toScopedSharedRedlineOperation
} from '../../src/taskpane/modules/commands/redline-operation-converter.js';

function testSubstringReplacementExact() {
    const result = applySubstringSearchReplace(
        'Alpha target paragraph.',
        'target',
        'updated'
    );
    assert.strictEqual(result.applied, true, 'exact replacement should apply');
    assert.strictEqual(result.matchMode, 'exact', 'exact replacement should report exact match mode');
    assert.strictEqual(result.modifiedText, 'Alpha updated paragraph.');
}

function testSubstringReplacementCaseInsensitive() {
    const result = applySubstringSearchReplace(
        'Alpha TARGET paragraph.',
        'target',
        'updated'
    );
    assert.strictEqual(result.applied, true, 'case-insensitive replacement should apply');
    assert.strictEqual(result.matchMode, 'case_insensitive', 'case-insensitive replacement should report mode');
    assert.strictEqual(result.modifiedText, 'Alpha updated paragraph.');
}

function testEditParagraphConversion() {
    const converted = toScopedSharedRedlineOperation(
        {
            operation: 'edit_paragraph',
            newContent: 'Rewritten content.'
        },
        {
            scopeStartText: 'Original content.',
            scopeParagraphCount: 1
        }
    );

    assert.strictEqual(converted.ok, true, 'edit_paragraph conversion should succeed');
    assert.deepStrictEqual(
        converted.operation,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'Original content.',
            modified: 'Rewritten content.'
        },
        'edit_paragraph conversion should produce scoped redline operation'
    );
}

function testReplaceRangeConversion() {
    const converted = toScopedSharedRedlineOperation(
        {
            operation: 'replace_range',
            content: '1. item a\n2. item b'
        },
        {
            scopeStartText: 'Item a',
            scopeParagraphCount: 3
        }
    );

    assert.strictEqual(converted.ok, true, 'replace_range conversion should succeed');
    assert.strictEqual(converted.operation.targetRef, 'P1', 'replace_range conversion should set scoped targetRef');
    assert.strictEqual(converted.operation.targetEndRef, 'P3', 'replace_range conversion should set scoped targetEndRef');
}

function testReplaceRangeInsertionBeforeStartNormalization() {
    const converted = toScopedSharedRedlineOperation(
        {
            paragraphIndex: 1,
            endParagraphIndex: 0,
            operation: 'replace_range',
            content: 'Please fill in the bracketed information throughout this document.'
        },
        {
            scopeStartText: 'NON-DISCLOSURE AGREEMENT',
            scopeParagraphCount: 1,
            insertionBeforeStart: true
        }
    );

    assert.strictEqual(converted.ok, true, 'replace_range insertion-before-start conversion should succeed');
    assert.deepStrictEqual(
        converted.operation,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'NON-DISCLOSURE AGREEMENT',
            modified: 'Please fill in the bracketed information throughout this document.\nNON-DISCLOSURE AGREEMENT'
        },
        'replace_range insertion-before-start should normalize into prefixed scoped redline text'
    );
}

function testModifyTextConversionNotFound() {
    const converted = toScopedSharedRedlineOperation(
        {
            operation: 'modify_text',
            originalText: 'missing',
            replacementText: 'updated'
        },
        {
            scopeStartText: 'Original content.',
            scopeParagraphCount: 1
        }
    );
    assert.strictEqual(converted.ok, false, 'modify_text conversion should fail if search text is missing');
}

function run() {
    testSubstringReplacementExact();
    testSubstringReplacementCaseInsensitive();
    testEditParagraphConversion();
    testReplaceRangeConversion();
    testReplaceRangeInsertionBeforeStartNormalization();
    testModifyTextConversionNotFound();
    console.log('PASS: redline operation converter tests');
}

run();

