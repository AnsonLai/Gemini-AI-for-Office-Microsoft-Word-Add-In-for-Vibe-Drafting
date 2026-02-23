// Re-export the standalone package surface.
export * from '@ansonlai/docx-redline-js';

// Add-in-only integration bridge exports.
export {
    applyReconciliationToParagraph,
    applyReconciliationToParagraphBatch,
    shouldUseOoxmlReconciliation,
    getAuthorForTracking
} from './integration.js';
export {
    getParagraphOoxmlWithFallback,
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';
export { applyStructuredListDirectOoxml } from './word-structured-list.js';
export { routeWordParagraphChange } from './word-route-change.js';
export {
    applyWordOperation,
    applySharedOperationToWordParagraph,
    applySharedOperationToWordScope,
    applySharedOperationToParagraphOoxml,
    applySharedOperationToScopeOoxml
} from './word-operation-runner.js';
export {
    applyRedlineChangesToWordContext,
    findNearbyParagraphIndexForModifyText
} from './word-redline-runner.js';

