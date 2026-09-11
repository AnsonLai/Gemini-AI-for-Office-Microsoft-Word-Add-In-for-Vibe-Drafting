import assert from 'assert';
import {
    assertRedlineResult,
    prepareOperationInput,
    RedlineOperationError
} from '../src/taskpane/modules/docx-redline-js-integration/redline-result.js';

function testSuccessAndNoOpPassThrough() {
    const success = { status: 'ok', hasChanges: true };
    const noOp = { status: 'no-op', hasChanges: false };

    assert.strictEqual(assertRedlineResult(success), success);
    assert.strictEqual(assertRedlineResult(noOp), noOp);
    assert.strictEqual(assertRedlineResult({ hasChanges: false }).hasChanges, false);
}

function testStructuredErrorsThrowWithStableDetails() {
    for (const code of [
        'PARSE_ERROR',
        'TARGET_NOT_FOUND',
        'PARTIAL_TARGET',
        'EXISTING_REVISIONS',
        'DIFF_TOKEN_LIMIT',
        'BATCH_OPERATION_FAILED',
        'FUTURE_ERROR_CODE'
    ]) {
        assert.throws(
            () => assertRedlineResult({
                status: 'error',
                hasChanges: false,
                error: { code, message: 'Synthetic failure' },
                warnings: ['diagnostic warning']
            }, 'contract test'),
            error => {
                assert.ok(error instanceof RedlineOperationError);
                assert.strictEqual(error.code, code);
                assert.strictEqual(error.message, `[${code}] Synthetic failure`);
                assert.deepStrictEqual(error.details, {
                    context: 'contract test',
                    warnings: ['diagnostic warning'],
                    packageError: { code, message: 'Synthetic failure' }
                });
                return true;
            }
        );
    }
}

function testErrorFieldAloneIsStillFailure() {
    assert.throws(
        () => assertRedlineResult({ error: { message: 'Missing status and code' } }),
        error => error.code === 'OPERATION_ERROR'
    );
}

function testFailureDiagnosticsArePreserved() {
    const receipt = { operationIndex: 2, finalDisposition: 'rolled_back', committed: false };
    const validationSummary = { errors: 1, warnings: 0 };
    assert.throws(
        () => assertRedlineResult({
            status: 'error',
            error: {
                code: 'GENERATED_OOXML_INVALID',
                message: 'Generated output failed validation',
                stage: 'mutation-envelope'
            },
            warnings: ['unsafe output rejected'],
            receipt,
            rolledBack: true,
            validationSummary
        }, 'diagnostic contract'),
        error => {
            assert.strictEqual(error.code, 'GENERATED_OOXML_INVALID');
            assert.deepStrictEqual(error.details.receipt, receipt);
            assert.strictEqual(error.details.rolledBack, true);
            assert.deepStrictEqual(error.details.validationSummary, validationSummary);
            assert.strictEqual(error.details.packageError.stage, 'mutation-envelope');
            return true;
        }
    );
}

function testOperationSanitizationIsExplicitAndLiteralSafe() {
    const source = {
        type: 'redline',
        modified: 'Here is the redline:\nPay $1,000 under $term$.'
    };
    const disabled = prepareOperationInput(source, false);
    const enabled = prepareOperationInput(source, true);

    assert.strictEqual(disabled.operation, source);
    assert.strictEqual(disabled.sanitized, false);
    assert.notStrictEqual(enabled.operation, source);
    assert.strictEqual(enabled.sanitized, true);
    assert.strictEqual(enabled.operation.modified, 'Pay $1,000 under $term$.');
    assert.strictEqual(source.modified, 'Here is the redline:\nPay $1,000 under $term$.');
}

function run() {
    testSuccessAndNoOpPassThrough();
    testStructuredErrorsThrowWithStableDetails();
    testErrorFieldAloneIsStillFailure();
    testFailureDiagnosticsArePreserved();
    testOperationSanitizationIsExplicitAndLiteralSafe();
    console.log('PASS: redline result contract tests');
}

run();
