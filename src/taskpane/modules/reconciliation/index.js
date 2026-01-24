/**
 * OOXML Reconciliation Pipeline - Module Entry Point
 * 
 * Exports the public API for the reconciliation system.
 */

// Main pipeline
export { ReconciliationPipeline, detectContentType, parseListItems, parseTable } from './pipeline.js';
export { NumberingService } from './numbering-service.js';

// Core types
export { DiffOp, RunKind, ContainerKind, ContentType, NS_W, NS_R, escapeXml, getNextRevisionId, resetRevisionIdCounter } from './types.js';

// Individual stage functions (for advanced usage)
export { ingestOoxml, ingestTableToVirtualGrid } from './ingestion.js';
export { preprocessMarkdown, getApplicableFormatHints, mergeFormats } from './markdown-processor.js';
export { computeWordLevelDiffOps, wordsToChars, charsToWords, collectDiffBoundaries } from './diff-engine.js';
export { splitRunsAtDiffBoundaries, applyPatches } from './patching.js';
export { serializeToOoxml, wrapInDocumentFragment } from './serialization.js';

// Integration helpers (for Word Add-in)
export { applyReconciliationToParagraph, shouldUseOoxmlReconciliation, getAuthorForTracking } from './integration.js';

// OOXML Engine V5.1 - Hybrid Mode (DOM-based manipulation)
export { applyRedlineToOxml, sanitizeAiResponse } from './oxml-engine.js';

// Table Reconciliation
export { generateTableOoxml, diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from './table-reconciliation.js';

