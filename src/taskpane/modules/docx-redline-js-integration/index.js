// Re-export the standalone package surface.
export * from '@ansonlai/docx-redline-js';

// Add-in-only integration bridge exports.
export {
    getParagraphOoxmlWithFallback,
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';
export {
    assertRedlineResult,
    INPUT_SANITIZED_WARNING,
    prepareOperationInput,
    RedlineOperationError
} from './redline-result.js';
export {
    applyWordOperation,
    applySharedOperationToWordParagraph,
    applySharedOperationToWordScope,
    applySharedOperationToParagraphOoxml,
    applySharedOperationToScopeOoxml,
    executePureOoxmlBatch
} from './word-operation-runner.js';
export {
    applyRedlineChangesToWordContext,
    planRedlineBatchOperations
} from './word-redline-runner.js';

