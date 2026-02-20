/**
 * Backward-compatible command-layer bridge exports.
 *
 * Canonical implementation now lives in reconciliation integration adapter.
 */

export {
    applySharedOperationToParagraphOoxml,
    applySharedOperationToScopeOoxml
} from '../reconciliation/integration/word-operation-runner.js';
