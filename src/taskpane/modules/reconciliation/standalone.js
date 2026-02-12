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
    resolveTargetParagraph
} from './core/paragraph-targeting.js';
export { synthesizeTableMarkdownFromMultilineCellEdit } from './core/table-targeting.js';
