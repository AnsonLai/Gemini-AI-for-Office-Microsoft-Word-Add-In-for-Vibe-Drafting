/**
 * Standalone reconciliation entrypoint (no Word JS API dependencies).
 */

// Adapters
export { configureXmlProvider } from './adapters/xml-adapter.js';
export { configureLogger } from './adapters/logger.js';

// Engine
import {
    applyRedlineToOxml as applyRedlineToOxmlEngine,
    sanitizeAiResponse,
    parseOoxml,
    serializeOoxml
} from './engine/oxml-engine.js';
import { wrapInDocumentFragment as wrapInDocumentFragmentShared } from './pipeline/serialization.js';
import {
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback
} from './orchestration/list-structural-fallback.js';

/**
 * Standalone-safe redline wrapper.
 *
 * In non-Word runtimes, the engine can return `{ useNativeApi: true, hasChanges: true }`
 * without an OOXML payload for some format-only operations. Standalone callers cannot
 * complete that native fallback path, so normalize to a no-op with warnings.
 */
export async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const result = await applyRedlineToOxmlEngine(oxml, originalText, modifiedText, options);
    if (result?.useNativeApi && typeof result?.oxml !== 'string') {
        const existingWarnings = Array.isArray(result?.warnings) ? result.warnings : [];
        return {
            ...result,
            oxml,
            hasChanges: false,
            warnings: [
                ...existingWarnings,
                'Standalone mode cannot execute native Word API fallback for this operation.'
            ]
        };
    }
    return result;
}

/**
 * Applies redline reconciliation, then forces single-line structural list
 * conversion when the redline is a no-op on marker-prefixed list text.
 *
 * This is useful for inputs like `1. HEADER` where text diff is unchanged but
 * OOXML should convert plain text markers into real Word list structure.
 *
 * @param {string} oxml - Original OOXML
 * @param {string} originalText - Original visible text
 * @param {string} modifiedText - Proposed modified text
 * @param {Object} [options={}] - Reconciliation options
 * @param {boolean} [options.listFallbackAllowExistingList=true] - Allow fallback even when paragraph is already list-bound
 * @returns {Promise<{ oxml: string, hasChanges: boolean } & Record<string, any>>}
 */
export async function applyRedlineToOxmlWithListFallback(oxml, originalText, modifiedText, options = {}) {
    const allowExistingListForFallback = options.listFallbackAllowExistingList !== false;
    const plan = buildSingleLineListStructuralFallbackPlan({
        oxml,
        originalText,
        modifiedText,
        allowExistingList: allowExistingListForFallback
    });
    const preferListFallback = options.preferListStructuralFallback !== false;
    let preflightFallbackWarnings = [];

    if (plan && preferListFallback) {
        const fallbackResult = await executeSingleLineListStructuralFallback(plan, {
            author: options.author,
            generateRedlines: options.generateRedlines,
            pipeline: options.listFallbackPipeline
        });
        if (fallbackResult?.hasChanges && fallbackResult?.oxml) {
            const wrappedOxml = wrapInDocumentFragmentShared(fallbackResult.oxml, {
                includeNumbering: fallbackResult.includeNumbering ?? true,
                numberingXml: fallbackResult.numberingXml
            });
            const fallbackWarnings = Array.isArray(fallbackResult?.warnings) ? fallbackResult.warnings : [];
            return {
                oxml: wrappedOxml,
                hasChanges: true,
                warnings: fallbackWarnings,
                listStructuralFallbackApplied: true,
                listStructuralFallbackKey: fallbackResult.listStructuralFallbackKey || null,
                listStructuralFallbackNumberingXml: fallbackResult.numberingXml || null
            };
        }
        preflightFallbackWarnings = Array.isArray(fallbackResult?.warnings) ? fallbackResult.warnings : [];
    }

    const baseResult = await applyRedlineToOxml(oxml, originalText, modifiedText, options);

    if (!plan) {
        return {
            ...baseResult,
            warnings: [
                ...(Array.isArray(baseResult?.warnings) ? baseResult.warnings : []),
                ...preflightFallbackWarnings
            ],
            listStructuralFallbackApplied: false
        };
    }

    if (preferListFallback) {
        return {
            ...baseResult,
            warnings: [
                ...(Array.isArray(baseResult?.warnings) ? baseResult.warnings : []),
                ...preflightFallbackWarnings
            ],
            listStructuralFallbackApplied: false
        };
    }

    if (baseResult?.hasChanges) {
        return {
            ...baseResult,
            listStructuralFallbackApplied: false
        };
    }

    const fallbackResult = await executeSingleLineListStructuralFallback(plan, {
        author: options.author,
        generateRedlines: options.generateRedlines,
        pipeline: options.listFallbackPipeline
    });
    if (!fallbackResult?.hasChanges || !fallbackResult?.oxml) {
        const existingWarnings = Array.isArray(baseResult?.warnings) ? baseResult.warnings : [];
        const fallbackWarnings = Array.isArray(fallbackResult?.warnings) ? fallbackResult.warnings : [];
        return {
            ...baseResult,
            warnings: [...existingWarnings, ...fallbackWarnings],
            listStructuralFallbackApplied: false
        };
    }

    const wrappedOxml = wrapInDocumentFragmentShared(fallbackResult.oxml, {
        includeNumbering: fallbackResult.includeNumbering ?? true,
        numberingXml: fallbackResult.numberingXml
    });
    const existingWarnings = Array.isArray(baseResult?.warnings) ? baseResult.warnings : [];
    const fallbackWarnings = Array.isArray(fallbackResult?.warnings) ? fallbackResult.warnings : [];

    return {
        ...baseResult,
        oxml: wrappedOxml,
        hasChanges: true,
        warnings: [...existingWarnings, ...preflightFallbackWarnings, ...fallbackWarnings],
        listStructuralFallbackApplied: true,
        listStructuralFallbackKey: fallbackResult.listStructuralFallbackKey || null,
        listStructuralFallbackNumberingXml: fallbackResult.numberingXml || null
    };
}

export { sanitizeAiResponse, parseOoxml, serializeOoxml };

// Pipeline components
export { ReconciliationPipeline } from './pipeline/pipeline.js';
export { ingestOoxml } from './pipeline/ingestion.js';
export { preprocessMarkdown } from './pipeline/markdown-processor.js';
export { serializeToOoxml, wrapInDocumentFragment } from './pipeline/serialization.js';

// Comment engine
export {
    injectCommentsIntoOoxml,
    injectCommentsIntoPackage,
    buildCommentElement,
    buildCommentsPartXml
} from './services/comment-engine.js';

// Formatting removal utilities (outside reconciliation folder)
export {
    removeFormattingFromRPr,
    applyFormattingRemovalToOoxml,
    applyHighlightToOoxml
} from '../../ooxml-formatting-removal.js';

// Table/list tools
export { generateTableOoxml } from './services/table-reconciliation.js';
export { NumberingService } from './services/numbering-service.js';
export { buildReconciliationPlan, RoutePlanKind, normalizeContentEscapesForRouting } from './orchestration/route-plan.js';
export { parseMarkdownListContent, hasListItems } from './orchestration/list-parsing.js';
export { buildListMarkdown, inferNumberingStyleFromMarker, normalizeListItemsWithLevels } from './orchestration/list-markdown.js';
export {
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback
} from './orchestration/list-structural-fallback.js';

// Core types/constants
export { DiffOp, RunKind, ContainerKind, ContentType, NS_W, escapeXml } from './core/types.js';
export { extractParagraphIdFromOoxml } from './core/ooxml-identifiers.js';
export {
    WORD_MAIN_NS,
    getParagraphText,
    getDocumentParagraphNodes,
    normalizeWhitespaceForTargeting,
    isMarkdownTableText,
    parseParagraphReference,
    stripLeadingParagraphMarker,
    splitLeadingParagraphMarker,
    findContainingWordElement,
    findParagraphByReference,
    findParagraphByStrictText,
    findParagraphByBestTextMatch,
    resolveTargetParagraph,
    buildTargetReferenceSnapshot,
    resolveTargetParagraphWithSnapshot
} from './core/paragraph-targeting.js';
export { synthesizeTableMarkdownFromMultilineCellEdit } from './core/table-targeting.js';
export {
    getParagraphListInfo,
    collectContiguousListParagraphBlock,
    synthesizeExpandedListScopeEdit,
    planListInsertionOnlyEdit,
    stripRedundantLeadingListMarkers
} from './core/list-targeting.js';
