/**
 * Backward-compatible command-layer export surface.
 *
 * Canonical converter implementation now lives in reconciliation orchestration.
 */

export {
    applySubstringSearchReplace,
    toScopedSharedRedlineOperation
} from '../reconciliation/orchestration/redline-operation-converter.js';
