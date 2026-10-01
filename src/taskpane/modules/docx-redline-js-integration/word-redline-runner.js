import { executePureOoxmlBatch } from './word-operation-runner.js';
import { planRedlineBatchOperationsWithMapping, assertSourceBaseline } from './redline-plan.js';
export { planRedlineBatchOperations, assertSourceBaseline } from './redline-plan.js';
export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    if (changes.length === 0) return { changesApplied: 0, skipped: [], written: false, writeAttempted: false, mutationOutcome: 'noop' };
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = options.onInfo || (() => {});
    const onWarn = options.onWarn || ((_message, diagnostic) => console.warn(
        `[${logPrefix}] Batch warning`, diagnostic || { code: 'ENGINE_WARNING' }
    ));
    const batchRunner = options.batchRunner || executePureOoxmlBatch;
    let changeOperationIndexes = changes.map((_, index) => index + 1);
    try {
        const result = await batchRunner(
            context,
            context.document.body,
            source => {
                if (Object.prototype.hasOwnProperty.call(options, 'sourceBaseline')) {
                    assertSourceBaseline(changes, source.paragraphs, options.sourceBaseline);
                }
                const plan = planRedlineBatchOperationsWithMapping(changes, source.paragraphs, {
                    ...options,
                    sourceDocumentXml: source.documentXml
                });
                changeOperationIndexes = plan.changeOperationIndexes;
                return plan.operations;
            },
            {
                author: options.author,
                generateRedlines: options.generateRedlines,
                sanitizeInput: options.sanitizeInput === true,
                disableNativeTracking: options.disableNativeTracking,
                baseTrackingMode: options.baseTrackingMode ?? null,
                onInfo,
                onWarn
            }
        );
        const receipts = Array.isArray(result?.receipts) ? result.receipts : [];
        const changesApplied = result?.written === true
            ? changeOperationIndexes.filter(operationIndex => receipts.some(receipt => (
                receipt.operationIndex === operationIndex && receipt.committed === true
            ))).length
            : 0;
        const skipped = changes.flatMap((change, index) => {
            const operationIndex = changeOperationIndexes[index] ?? index + 1;
            const receipt = receipts.find(item => item.operationIndex === operationIndex);
            if (result?.written === true && receipt?.committed) return [];
            const item = result?.results?.[operationIndex - 1];
            const error = item?.error || result?.error;
            return [{
                paragraphIndex: change?.paragraphIndex,
                operation: change?.operation,
                reason: error?.message || (result?.hasChanges === false ? 'no changes produced' : 'batch was rolled back'),
                ...(error?.code ? { code: error.code } : {}),
                ...(receipt ? { receipt } : {}),
                ...(result?.rolledBack !== undefined ? { rolledBack: result.rolledBack } : {})
            }];
        });
        if (result?.status === 'error') onWarn(result.error?.message || 'Redline batch failed.', {
            code: result.error?.code || 'ENGINE_REFUSED',
            mutationOutcome: result.mutationOutcome,
            written: result.written === true,
            writeAttempted: result.writeAttempted === true,
            rolledBack: result.rolledBack === true,
            operationErrors: (result.results || []).flatMap((item, index) => item?.error ? [{
                operationIndex: index + 1,
                code: item.error.code || 'OPERATION_ERROR'
            }] : [])
        });
        onInfo(`Total changes applied: ${changesApplied}`);
        return {
            changesApplied, skipped, batchResult: result,
            status: result?.status, error: result?.error,
            receipts, written: result?.written === true,
            writeAttempted: result?.writeAttempted === true,
            mutationOutcome: result?.mutationOutcome || (result?.written ? 'applied' : result?.rolledBack ? 'rolled_back' : result?.hasChanges === false ? 'noop' : 'refused')
        };
    } catch (error) {
        onWarn(`Redline batch failed: ${error?.message || error}`, {
            code: error?.code || 'REDLINE_BATCH_FAILED', written: false, writeAttempted: false
        });
        return {
            changesApplied: 0,
            skipped: changes.map(change => ({
                paragraphIndex: change?.paragraphIndex,
                operation: change?.operation,
                reason: error?.message || String(error),
                ...(error?.code ? { code: error.code } : {})
            })),
            error,
            written: false,
            writeAttempted: false,
            mutationOutcome: ['STALE_DOCUMENT_CONTEXT', 'UNSUPPORTED_TABLE_FORMATTING'].includes(error?.code)
                ? 'refused'
                : 'failed'
        };
    }
}
