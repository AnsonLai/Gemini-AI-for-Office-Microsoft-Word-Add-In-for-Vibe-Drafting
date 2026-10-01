import { executePureOoxmlBatch } from './word-operation-runner.js';
import { planRedlineBatchOperations, assertSourceBaseline } from './redline-plan.js';
export { planRedlineBatchOperations, assertSourceBaseline } from './redline-plan.js';
export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    if (changes.length === 0) return { changesApplied: 0, skipped: [], written: false, writeAttempted: false, mutationOutcome: 'noop' };
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = options.onInfo || (() => {});
    const onWarn = options.onWarn || (() => console.warn(`[${logPrefix}] Batch warning; consult the structured result.`));
    const batchRunner = options.batchRunner || executePureOoxmlBatch;
    try {
        const result = await batchRunner(
            context,
            context.document.body,
            source => {
                if (Object.prototype.hasOwnProperty.call(options, 'sourceBaseline')) {
                    assertSourceBaseline(changes, source.paragraphs, options.sourceBaseline);
                }
                return planRedlineBatchOperations(changes, source.paragraphs, options);
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
            ? receipts.filter(receipt => receipt.committed === true).length
            : 0;
        const skipped = changes.flatMap((change, index) => {
            const receipt = receipts.find(item => item.operationIndex === index + 1);
            if (result?.written === true && receipt?.committed) return [];
            const item = result?.results?.[index];
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
        if (result?.status === 'error') onWarn(result.error?.message || 'Redline batch failed.');
        onInfo(`Total changes applied: ${changesApplied}`);
        return {
            changesApplied, skipped, batchResult: result,
            status: result?.status, error: result?.error,
            receipts, written: result?.written === true,
            writeAttempted: result?.writeAttempted === true,
            mutationOutcome: result?.mutationOutcome || (result?.written ? 'applied' : result?.rolledBack ? 'rolled_back' : result?.hasChanges === false ? 'noop' : 'refused')
        };
    } catch (error) {
        onWarn(`Redline batch failed: ${error?.message || error}`);
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
            mutationOutcome: error?.code === 'STALE_DOCUMENT_CONTEXT' ? 'refused' : 'failed'
        };
    }
}
