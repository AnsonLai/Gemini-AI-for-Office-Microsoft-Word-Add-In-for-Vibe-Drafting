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

function parseIntegerAttribute(element, names) {
    if (!element || !Array.isArray(names)) return null;
    for (const name of names) {
        const raw = element.getAttribute(name);
        if (raw == null || raw === '') continue;
        const parsed = Number.parseInt(String(raw), 10);
        if (Number.isFinite(parsed)) return parsed;
    }
    return null;
}

function nextAvailableId(startId, occupiedIds, maxPreferred = null) {
    let candidate = Number.isInteger(startId) && startId > 0 ? startId : 1;
    const occupied = occupiedIds instanceof Set ? occupiedIds : new Set();

    while (occupied.has(candidate)) {
        candidate += 1;
    }

    if (Number.isInteger(maxPreferred) && maxPreferred > 0 && candidate > maxPreferred) {
        for (let probe = 1; probe <= maxPreferred; probe += 1) {
            if (!occupied.has(probe)) return probe;
        }
    }

    return candidate;
}

/**
 * Builds a dynamic numbering-id state from existing numbering XML.
 *
 * This avoids hardcoded ID floors and keeps IDs deterministic relative to the
 * current document numbering definitions.
 *
 * @param {string} numberingXml - Existing `word/numbering.xml` content
 * @param {{
 *   minId?: number,
 *   maxPreferred?: number
 * }} [options={}] - Optional ID preferences
 * @returns {{
 *   nextNumId: number,
 *   nextAbstractNumId: number,
 *   usedNumIds: Set<number>,
 *   usedAbstractNumIds: Set<number>,
 *   minId: number,
 *   maxPreferred: number
 * }}
 */
export function createDynamicNumberingIdState(numberingXml, options = {}) {
    const minId = Number.isInteger(options?.minId) && options.minId > 0 ? options.minId : 1;
    const maxPreferred = Number.isInteger(options?.maxPreferred) && options.maxPreferred >= minId
        ? options.maxPreferred
        : 32767;

    const usedNumIds = new Set();
    const usedAbstractNumIds = new Set();

    if (String(numberingXml || '').trim()) {
        try {
            const numberingDoc = parseOoxml(numberingXml);
            const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
            const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));

            for (const abstractNum of abstractNums) {
                const id = parseIntegerAttribute(abstractNum, ['w:abstractNumId', 'abstractNumId']);
                if (id != null) usedAbstractNumIds.add(id);
            }
            for (const num of nums) {
                const id = parseIntegerAttribute(num, ['w:numId', 'numId']);
                if (id != null) usedNumIds.add(id);
            }
        } catch {
            // Ignore malformed numbering XML and fall back to empty sets.
        }
    }

    const maxUsedNumId = usedNumIds.size > 0 ? Math.max(...usedNumIds) : 0;
    const maxUsedAbstractNumId = usedAbstractNumIds.size > 0 ? Math.max(...usedAbstractNumIds) : 0;
    const baseNumId = Math.max(minId, maxUsedNumId + 1);
    const baseAbstractNumId = Math.max(minId, maxUsedAbstractNumId + 1);

    return {
        nextNumId: nextAvailableId(baseNumId, usedNumIds, maxPreferred),
        nextAbstractNumId: nextAvailableId(baseAbstractNumId, usedAbstractNumIds, maxPreferred),
        usedNumIds,
        usedAbstractNumIds,
        minId,
        maxPreferred
    };
}

function normalizeNumberingIdState(state) {
    if (!state || typeof state !== 'object') return null;

    if (!(state.usedNumIds instanceof Set)) state.usedNumIds = new Set();
    if (!(state.usedAbstractNumIds instanceof Set)) state.usedAbstractNumIds = new Set();

    if (!Number.isInteger(state.minId) || state.minId < 1) {
        state.minId = 1;
    }
    if (!Number.isInteger(state.maxPreferred) || state.maxPreferred < state.minId) {
        state.maxPreferred = 32767;
    }

    if (!Number.isInteger(state.nextNumId) || state.nextNumId < state.minId) {
        state.nextNumId = state.minId;
    }
    if (!Number.isInteger(state.nextAbstractNumId) || state.nextAbstractNumId < state.minId) {
        state.nextAbstractNumId = state.minId;
    }

    state.nextNumId = nextAvailableId(state.nextNumId, state.usedNumIds, state.maxPreferred);
    state.nextAbstractNumId = nextAvailableId(state.nextAbstractNumId, state.usedAbstractNumIds, state.maxPreferred);
    return state;
}

/**
 * Reserves the next available ID from a mutable numbering-id state.
 *
 * @param {ReturnType<typeof createDynamicNumberingIdState>} state
 * @param {'num'|'abstract'} [kind='num']
 * @returns {number|null}
 */
export function reserveNextNumberingId(state, kind = 'num') {
    const normalized = normalizeNumberingIdState(state);
    if (!normalized) return null;

    const useAbstract = kind === 'abstract';
    const id = useAbstract ? normalized.nextAbstractNumId : normalized.nextNumId;
    if (!Number.isInteger(id) || id < 1) return null;

    if (useAbstract) {
        normalized.usedAbstractNumIds.add(id);
        normalized.nextAbstractNumId = nextAvailableId(id + 1, normalized.usedAbstractNumIds, normalized.maxPreferred);
    } else {
        normalized.usedNumIds.add(id);
        normalized.nextNumId = nextAvailableId(id + 1, normalized.usedNumIds, normalized.maxPreferred);
    }

    return id;
}

/**
 * Reserves the next available numbering IDs on a mutable numbering-id state.
 *
 * @param {ReturnType<typeof createDynamicNumberingIdState>} state
 * @returns {{ numId: number, abstractNumId: number } | null}
 */
export function reserveNextNumberingIdPair(state) {
    const numId = reserveNextNumberingId(state, 'num');
    const abstractNumId = reserveNextNumberingId(state, 'abstract');
    if (numId == null || abstractNumId == null) return null;

    return { numId, abstractNumId };
}

const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function hasXmlParseError(doc) {
    if (!doc || !doc.documentElement) return true;
    if (doc.documentElement.localName === 'parsererror') return true;
    return doc.getElementsByTagName('parsererror').length > 0;
}

function isDirectWordChild(node, localName) {
    return !!(
        node &&
        node.nodeType === 1 &&
        node.namespaceURI === WORD_MAIN_NS &&
        node.localName === localName
    );
}

function insertNumberingNodeInSchemaOrder(root, node, kind) {
    if (!root || !node) return;
    const directChildren = Array.from(root.childNodes || []).filter(
        child => child && child.nodeType === 1 && child.namespaceURI === WORD_MAIN_NS
    );

    let anchor = null;
    if (kind === 'abstract') {
        anchor = directChildren.find(
            child => child.localName === 'num' || child.localName === 'numIdMacAtCleanup'
        ) || null;
    } else {
        anchor = directChildren.find(child => child.localName === 'numIdMacAtCleanup') || null;
    }

    if (anchor) root.insertBefore(node, anchor);
    else root.appendChild(node);
}

/**
 * Merges incoming numbering definitions into an existing numbering part while
 * preserving schema child order (`abstractNum*` before `num*`).
 *
 * @param {string} existingNumberingXml
 * @param {string} incomingNumberingXml
 * @returns {string}
 */
export function mergeNumberingXmlBySchemaOrder(existingNumberingXml, incomingNumberingXml) {
    const existingText = String(existingNumberingXml || '');
    const incomingText = String(incomingNumberingXml || '');
    if (!incomingText.trim()) return existingText;
    if (!existingText.trim()) return incomingText;

    try {
        const existingDoc = parseOoxml(existingText);
        const incomingDoc = parseOoxml(incomingText);
        if (hasXmlParseError(existingDoc) || hasXmlParseError(incomingDoc)) {
            return existingText;
        }

        const existingRoot = existingDoc.documentElement;
        const incomingRoot = incomingDoc.documentElement;
        if (!existingRoot || !incomingRoot) return existingText;

        const existingAbstractIds = new Set(
            Array.from(existingRoot.childNodes || [])
                .filter(node => isDirectWordChild(node, 'abstractNum'))
                .map(node => parseIntegerAttribute(node, ['w:abstractNumId', 'abstractNumId']))
                .filter(id => id != null)
        );
        const existingNumIds = new Set(
            Array.from(existingRoot.childNodes || [])
                .filter(node => isDirectWordChild(node, 'num'))
                .map(node => parseIntegerAttribute(node, ['w:numId', 'numId']))
                .filter(id => id != null)
        );

        for (const incomingNode of Array.from(incomingRoot.childNodes || [])) {
            if (!isDirectWordChild(incomingNode, 'abstractNum')) continue;
            const incomingId = parseIntegerAttribute(incomingNode, ['w:abstractNumId', 'abstractNumId']);
            if (incomingId == null || existingAbstractIds.has(incomingId)) continue;
            insertNumberingNodeInSchemaOrder(existingRoot, existingDoc.importNode(incomingNode, true), 'abstract');
            existingAbstractIds.add(incomingId);
        }

        for (const incomingNode of Array.from(incomingRoot.childNodes || [])) {
            if (!isDirectWordChild(incomingNode, 'num')) continue;
            const incomingId = parseIntegerAttribute(incomingNode, ['w:numId', 'numId']);
            if (incomingId == null || existingNumIds.has(incomingId)) continue;
            insertNumberingNodeInSchemaOrder(existingRoot, existingDoc.importNode(incomingNode, true), 'num');
            existingNumIds.add(incomingId);
        }

        return serializeOoxml(existingDoc);
    } catch {
        return existingText;
    }
}

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
