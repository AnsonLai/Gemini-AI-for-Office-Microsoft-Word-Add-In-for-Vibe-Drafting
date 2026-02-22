/**
 * Standalone reconciliation entrypoint (no Word JS API dependencies).
 */

// Adapters
export { configureXmlProvider } from './adapters/xml-adapter.js';
export { configureLogger } from './adapters/logger.js';
export { setDefaultAuthor, getDefaultAuthor, setPlatform, getPlatform } from './adapters/config.js';

// Engine
import {
    applyRedlineToOxml as applyRedlineToOxmlEngine,
    sanitizeAiResponse,
    parseOoxml,
    serializeOoxml
} from './engine/oxml-engine.js';
import { parseTable as parseMarkdownTable } from './pipeline/pipeline.js';
import { wrapInDocumentFragment as wrapInDocumentFragmentShared } from './pipeline/serialization.js';
import {
    WORD_MAIN_NS,
    getParagraphText as getParagraphTextCore,
    resolveTargetParagraphWithSnapshot as resolveTargetParagraphWithSnapshotCore
} from './core/paragraph-targeting.js';
import {
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback,
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    stripSingleLineListMarkerPrefix
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
 * Reconciles a Markdown table against an OOXML scope.
 *
 * This centralizes table-specific validation + reconciliation so Word add-in
 * and browser modules can share the same entrypoint.
 *
 * @param {string} oxml - OOXML scope to reconcile (paragraph/range/table package)
 * @param {string} originalText - Original visible text in that scope
 * @param {string} markdownTable - Markdown table text
 * @param {Object} [options={}] - Reconciliation options forwarded to applyRedlineToOxml
 * @returns {Promise<{ oxml: string, hasChanges: boolean, warnings?: string[], isMarkdownTable: boolean, tableData?: Object }>}
 */
export async function reconcileMarkdownTableOoxml(oxml, originalText, markdownTable, options = {}) {
    const sourceOoxml = typeof oxml === 'string' ? oxml : '';
    const tableText = typeof markdownTable === 'string' ? markdownTable : String(markdownTable || '');
    let tableData;

    try {
        tableData = parseMarkdownTable(tableText);
    } catch {
        tableData = { headers: [], rows: [] };
    }

    const hasTableData = (tableData?.headers?.length || 0) > 0 || (tableData?.rows?.length || 0) > 0;
    if (!hasTableData) {
        return {
            oxml: sourceOoxml,
            hasChanges: false,
            isMarkdownTable: false,
            warnings: ['Could not parse Markdown table from input.']
        };
    }

    const result = await applyRedlineToOxml(
        sourceOoxml,
        originalText || '',
        tableText,
        options
    );

    return {
        ...result,
        isMarkdownTable: true,
        tableData
    };
}

/**
 * Heuristic detector for paragraphs likely belonging to a table-source block.
 *
 * @param {string} text - Paragraph text
 * @returns {boolean}
 */
export function isLikelyStructuredTableSourceParagraph(text) {
    const normalized = String(text || '').trim();
    if (!normalized) return false;
    if (/^and$/i.test(normalized)) return true;
    if (/^\[.*\]$/.test(normalized)) return true;
    if (/^\(.*\)$/.test(normalized)) return true;
    if (/:\s*$/.test(normalized)) return true;
    if (normalized.length <= 90 && !/[.!?]$/.test(normalized) && /[:\[\]()]/.test(normalized)) return true;
    if (/^[\[(]/.test(normalized)) return true;
    return false;
}

/**
 * Infers a contiguous paragraph block for table conversion starting from a paragraph.
 *
 * @param {Element|null} startParagraph - Starting w:p node
 * @param {Object} [options={}] - Inference options
 * @param {number} [options.maxScan=10] - Max sibling paragraphs to inspect
 * @param {(paragraph: Element) => string} [options.getParagraphText] - Optional text getter
 * @returns {Element[]|null}
 */
export function inferTableReplacementParagraphBlock(startParagraph, options = {}) {
    const maxScan = Number.isInteger(options?.maxScan) && options.maxScan > 0 ? options.maxScan : 10;
    const paragraphTextGetter = typeof options?.getParagraphText === 'function'
        ? options.getParagraphText
        : getParagraphTextCore;

    if (!startParagraph || !startParagraph.parentNode) return null;

    const block = [startParagraph];
    let cursor = startParagraph.nextSibling;
    let scanned = 0;

    while (cursor && scanned < maxScan) {
        scanned += 1;
        const nextCursor = cursor.nextSibling;
        if (cursor.nodeType !== 1 || cursor.namespaceURI !== WORD_MAIN_NS || cursor.localName !== 'p') {
            cursor = nextCursor;
            continue;
        }

        const text = String(paragraphTextGetter(cursor) || '').trim();
        if (!text) {
            if (block.length > 1) break;
            cursor = nextCursor;
            continue;
        }

        if (!isLikelyStructuredTableSourceParagraph(text)) break;
        block.push(cursor);
        cursor = nextCursor;
    }

    return block.length > 1 ? block : null;
}

/**
 * Resolves a contiguous paragraph range using paragraph references.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string|number|null} startRef - Start paragraph reference (e.g. P12)
 * @param {string|number|null} endRef - End paragraph reference (e.g. P15)
 * @param {Object} [options={}] - Resolution options
 * @param {string} [options.opType='redline'] - Operation type hint
 * @param {Array|null} [options.targetRefSnapshot=null] - Optional target snapshot
 * @param {(message: string) => void} [options.onInfo] - Optional info logger
 * @param {(message: string) => void} [options.onWarn] - Optional warn logger
 * @returns {Element[]|null}
 */
export function resolveParagraphRangeByRefs(xmlDoc, startRef, endRef, options = {}) {
    if (!xmlDoc || !startRef || !endRef) return null;

    const opType = options?.opType || 'redline';
    const targetRefSnapshot = options?.targetRefSnapshot || null;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };

    const start = resolveTargetParagraphWithSnapshotCore(xmlDoc, {
        targetRef: startRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!start) return null;

    const end = resolveTargetParagraphWithSnapshotCore(xmlDoc, {
        targetRef: endRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!end) return null;

    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS('*', 'p'));
    const startIdx = allParagraphs.indexOf(start);
    const endIdx = allParagraphs.indexOf(end);
    if (startIdx < 0 || endIdx < startIdx) return null;

    const range = allParagraphs.slice(startIdx, endIdx + 1);
    if (range.length === 0) return null;

    const parent = range[0]?.parentNode || null;
    if (!parent) return null;
    if (!range.every(node => node && node.parentNode === parent)) return null;
    return range;
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

export {
    createDynamicNumberingIdState,
    reserveNextNumberingId,
    reserveNextNumberingIdPair,
    overwriteParagraphNumIds,
    extractFirstParagraphNumId,
    buildExplicitDecimalMultilevelNumberingXml,
    remapNumberingPayloadForDocument,
    mergeNumberingXmlBySchemaOrder
} from './services/numbering-helpers.js';

// Pipeline components
export { ReconciliationPipeline } from './pipeline/pipeline.js';
export { ingestOoxml } from './pipeline/ingestion.js';
export { ingestWordOoxmlToPlainText, ingestWordOoxmlToMarkdown } from './pipeline/ingestion-export.js';
export { preprocessMarkdown } from './pipeline/markdown-processor.js';
export { serializeToOoxml, wrapInDocumentFragment } from './pipeline/serialization.js';

// Comment engine
export {
    injectCommentsIntoOoxml,
    injectCommentsIntoPackage,
    buildCommentElement,
    buildCommentsPartXml
} from './services/comment-engine.js';

// Formatting removal utilities
export {
    removeFormattingFromRPr,
    applyFormattingRemovalToOoxml,
    applyHighlightToOoxml
} from './engine/formatting-removal.js';

// Table/list tools
export { generateTableOoxml } from './services/table-reconciliation.js';
export { NumberingService } from './services/numbering-service.js';
export {
    parseXmlStrictStandalone,
    getBodyElementFromDocument,
    insertBodyElementBeforeSectPr,
    normalizeBodySectionOrderStandalone,
    sanitizeNestedParagraphsInTables,
    getPackagePartName,
    extractReplacementNodesFromOoxml,
    ensureNumberingArtifactsInZip,
    ensureCommentsArtifactsInZip,
    validateDocxPackage
} from './services/standalone-docx-plumbing.js';
export { buildReconciliationPlan, RoutePlanKind, normalizeContentEscapesForRouting } from './orchestration/route-plan.js';
export { parseMarkdownListContent, hasListItems } from './orchestration/list-parsing.js';
export { buildListMarkdown, inferNumberingStyleFromMarker, normalizeListItemsWithLevels } from './orchestration/list-markdown.js';
export {
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback,
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    stripSingleLineListMarkerPrefix
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

