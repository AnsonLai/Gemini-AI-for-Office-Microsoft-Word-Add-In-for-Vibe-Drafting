/**
 * OOXML Reconciliation Pipeline - Module Entry Point
 * 
 * Exports the public API for the reconciliation system.
 */

// Adapters
export { configureXmlProvider, createParser, createSerializer, parseXml, serializeXml } from './adapters/xml-adapter.js';
export { configureLogger, setLogLevel, getLogLevel, log, warn, error } from './adapters/logger.js';

// Main pipeline
export { ReconciliationPipeline, detectContentType, parseListItems, parseTable } from './pipeline/pipeline.js';
export { NumberingService } from './services/numbering-service.js';

// Core types
export { DiffOp, RunKind, ContainerKind, ContentType, NS_W, NS_R, escapeXml, getNextRevisionId, resetRevisionIdCounter } from './core/types.js';
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
    planListInsertionOnlyEdit
} from './core/list-targeting.js';

// Individual stage functions (for advanced usage)
export { ingestOoxml, ingestTableToVirtualGrid } from './pipeline/ingestion.js';
export { preprocessMarkdown, getApplicableFormatHints, mergeFormats } from './pipeline/markdown-processor.js';
export { computeWordLevelDiffOps, computeWordDiffs, wordsToChars, charsToWords, collectDiffBoundaries } from './pipeline/diff-engine.js';
export { splitRunsAtDiffBoundaries, applyPatches } from './pipeline/patching.js';
export { serializeToOoxml, wrapInDocumentFragment } from './pipeline/serialization.js';

// Integration helpers (for Word Add-in)
export { applyReconciliationToParagraph, applyReconciliationToParagraphBatch, shouldUseOoxmlReconciliation, getAuthorForTracking } from './integration/integration.js';
export { getParagraphOoxmlWithFallback, insertOoxmlWithRangeFallback, withNativeTrackingDisabled } from './integration/word-ooxml.js';
export { applyStructuredListDirectOoxml } from './integration/word-structured-list.js';
export { routeWordParagraphChange } from './integration/word-route-change.js';

// OOXML Engine V5.1 - Hybrid Mode (DOM-based manipulation)
export { applyRedlineToOxml, sanitizeAiResponse, parseOoxml, serializeOoxml } from './engine/oxml-engine.js';

// Table Reconciliation
export { generateTableOoxml, diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from './services/table-reconciliation.js';

// Comment Engine
export { injectCommentsIntoOoxml, injectCommentsIntoPackage, buildCommentElement, buildCommentsPartXml } from './services/comment-engine.js';

// Command-layer orchestration helpers (Word-agnostic planners)
export { buildReconciliationPlan, RoutePlanKind, normalizeContentEscapesForRouting } from './orchestration/route-plan.js';
export { parseMarkdownListContent, hasListItems } from './orchestration/list-parsing.js';
export { buildListMarkdown, inferNumberingStyleFromMarker, normalizeListItemsWithLevels } from './orchestration/list-markdown.js';

