/**
 * Shared fallback helpers for structural list conversion when a redline is a
 * text no-op but the target is marker-prefixed plain text (for example `1. X`).
 */

import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import { getXmlParseError } from '../core/xml-query.js';
import {
    getDocumentParagraphNodes,
    normalizeWhitespaceForTargeting
} from '../core/paragraph-targeting.js';
import { getParagraphListInfo } from '../core/list-targeting.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { parseMarkdownListContent, hasListItems } from './list-parsing.js';
import { inferNumberingStyleFromMarker } from './list-markdown.js';

function parseSingleLineListCandidate(text) {
    const rawText = String(text || '');
    if (!rawText.trim()) return null;
    if (rawText.includes('\n')) return null;

    const parsed = parseMarkdownListContent(rawText);
    if (!parsed || !hasListItems(parsed) || !Array.isArray(parsed.items) || parsed.items.length !== 1) {
        return null;
    }

    const item = parsed.items[0];
    if (!item || (item.type !== 'numbered' && item.type !== 'bullet')) return null;

    const marker = String(item.marker || '').trim();
    const numberingStyle = item.type === 'numbered'
        ? inferNumberingStyleFromMarker(marker || '1.')
        : 'bullet';

    return {
        type: item.type,
        marker,
        numberingStyle,
        startAt: parseMarkerStart(marker, numberingStyle),
        contentText: String(item.text || '').trim(),
        normalizedContent: normalizeWhitespaceForTargeting(String(item.text || ''))
    };
}

function parseMarkerStart(marker, numberingStyle) {
    if (numberingStyle !== 'decimal') return null;
    const match = String(marker || '').trim().match(/^(\d+)\.?$/);
    if (!match) return null;
    const value = Number.parseInt(match[1], 10);
    return Number.isFinite(value) && value > 0 ? value : null;
}

/**
 * Resolves how a caller should handle numbering ownership for a single-line
 * structural list fallback across multi-operation runs.
 *
 * @param {{
 *   numberingKey?: string|null,
 *   startAt?: number|null
 * }|null} plan - Single-line fallback plan
 * @param {{
 *   explicitByNumberingKey?: Map<string, { numId: string, nextStartAt: number }>
 * }|null} sequenceState - Optional mutable sequence state
 * @returns {{
 *   type: 'none'|'sharedByStyle'|'explicitReuse'|'explicitStartNew'|'explicitIsolated',
 *   numberingKey: string|null,
 *   startAt: number|null,
 *   numId: string|null
 * }}
 */
export function resolveSingleLineListFallbackNumberingAction(plan, sequenceState = null) {
    const numberingKey = plan?.numberingKey ? String(plan.numberingKey) : null;
    const startAt = Number.isInteger(plan?.startAt) && plan.startAt > 0 ? plan.startAt : null;

    if (!numberingKey) {
        return {
            type: 'none',
            numberingKey: null,
            startAt,
            numId: null
        };
    }

    if (startAt == null) {
        return {
            type: 'sharedByStyle',
            numberingKey,
            startAt: null,
            numId: null
        };
    }

    const explicitMap = sequenceState?.explicitByNumberingKey;
    if (!(explicitMap instanceof Map)) {
        return {
            type: 'explicitIsolated',
            numberingKey,
            startAt,
            numId: null
        };
    }

    const existing = explicitMap.get(numberingKey) || null;
    if (
        existing &&
        existing.numId != null &&
        Number.isInteger(existing.nextStartAt) &&
        existing.nextStartAt === startAt
    ) {
        return {
            type: 'explicitReuse',
            numberingKey,
            startAt,
            numId: String(existing.numId)
        };
    }

    return {
        type: 'explicitStartNew',
        numberingKey,
        startAt,
        numId: null
    };
}

/**
 * Records/advances explicit numeric single-line sequence state.
 *
 * @param {{
 *   explicitByNumberingKey?: Map<string, { numId: string, nextStartAt: number }>
 * }|null} sequenceState - Mutable sequence state
 * @param {string|null} numberingKey - Numbering signature key
 * @param {string|number|null} numId - Sequence numId
 * @param {number|null} startAt - Current explicit start marker
 */
export function recordSingleLineListFallbackExplicitSequence(sequenceState, numberingKey, numId, startAt) {
    if (!sequenceState || !(sequenceState.explicitByNumberingKey instanceof Map)) return;
    if (!numberingKey || numId == null || !Number.isInteger(startAt) || startAt < 1) return;

    sequenceState.explicitByNumberingKey.set(String(numberingKey), {
        numId: String(numId),
        nextStartAt: startAt + 1
    });
}

/**
 * Clears explicit sequence tracking for a numbering signature.
 *
 * @param {{
 *   explicitByNumberingKey?: Map<string, { numId: string, nextStartAt: number }>
 * }|null} sequenceState - Mutable sequence state
 * @param {string|null} numberingKey - Numbering signature key
 */
export function clearSingleLineListFallbackExplicitSequence(sequenceState, numberingKey) {
    if (!sequenceState || !(sequenceState.explicitByNumberingKey instanceof Map)) return;
    if (!numberingKey) return;
    sequenceState.explicitByNumberingKey.delete(String(numberingKey));
}

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node =>
            node &&
            node.nodeType === 1 &&
            node.namespaceURI === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' &&
            node.localName === localName
    ) || null;
}

/**
 * Forces explicit list bindings on paragraph nodes.
 *
 * Useful when callers need deterministic `w:numPr` assignment regardless of
 * incoming generator quirks or paragraph property tracked-change payloads.
 *
 * @param {Node[]|null|undefined} nodes - Candidate nodes containing paragraphs
 * @param {{
 *   numId: string|number,
 *   ilvl?: number,
 *   clearParagraphPropertyChanges?: boolean,
 *   removeListPropertyNode?: boolean
 * }} options - Binding options
 * @returns {number} Number of paragraph nodes updated
 */
export function enforceListBindingOnParagraphNodes(nodes, options = {}) {
    const numId = options?.numId;
    if (numId == null) return 0;
    const ilvl = Number.isInteger(options?.ilvl) ? Math.max(0, options.ilvl) : 0;
    const clearParagraphPropertyChanges = options?.clearParagraphPropertyChanges !== false;
    const removeListPropertyNode = options?.removeListPropertyNode !== false;

    const paragraphs = (Array.isArray(nodes) ? nodes : [])
        .filter(node => node && node.nodeType === 1 && node.localName === 'p');
    let updated = 0;

    for (const paragraph of paragraphs) {
        const ownerDoc = paragraph.ownerDocument;
        if (!ownerDoc) continue;

        let pPr = getDirectWordChild(paragraph, 'pPr');
        if (!pPr) {
            pPr = ownerDoc.createElementNS(
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w:pPr'
            );
            paragraph.insertBefore(pPr, paragraph.firstChild);
        }

        if (clearParagraphPropertyChanges) {
            const pPrChange = getDirectWordChild(pPr, 'pPrChange');
            if (pPrChange) pPr.removeChild(pPrChange);
        }

        if (removeListPropertyNode) {
            const listPr = getDirectWordChild(pPr, 'listPr');
            if (listPr) pPr.removeChild(listPr);
        }

        let numPr = getDirectWordChild(pPr, 'numPr');
        if (!numPr) {
            numPr = ownerDoc.createElementNS(
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w:numPr'
            );
            pPr.appendChild(numPr);
        }

        let ilvlEl = getDirectWordChild(numPr, 'ilvl');
        if (!ilvlEl) {
            ilvlEl = ownerDoc.createElementNS(
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w:ilvl'
            );
            numPr.appendChild(ilvlEl);
        }
        ilvlEl.setAttribute('w:val', String(ilvl));

        let numIdEl = getDirectWordChild(numPr, 'numId');
        if (!numIdEl) {
            numIdEl = ownerDoc.createElementNS(
                'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w:numId'
            );
            numPr.appendChild(numIdEl);
        }
        numIdEl.setAttribute('w:val', String(numId));
        updated++;
    }

    return updated;
}

function getFirstParagraphFromOxml(oxml) {
    const parser = createParser();
    const doc = parser.parseFromString(String(oxml || ''), 'application/xml');
    const parseError = getXmlParseError(doc);
    if (parseError) return null;
    const paragraphs = getDocumentParagraphNodes(doc);
    return paragraphs[0] || null;
}

function trimTrailingBlankParagraph(oxml) {
    if (!oxml) return '';
    return String(oxml).replace(/<w:p>\s*<w:pPr>\s*<\/w:pPr>\s*<\/w:p>\s*$/i, '');
}

function getAttributeFirst(element, names) {
    for (const name of names) {
        const value = element.getAttribute(name);
        if (value != null && value !== '') return value;
    }
    return null;
}

function getElementId(element, names) {
    const raw = getAttributeFirst(element, names);
    const parsed = Number.parseInt(String(raw || ''), 10);
    return Number.isFinite(parsed) ? parsed : null;
}

function setElementVal(element, value) {
    element.setAttribute('w:val', String(value));
}

function extractFirstParagraphNumIdFromOxml(oxml) {
    const parser = createParser();
    const doc = parser.parseFromString(String(oxml || ''), 'application/xml');
    const parseError = getXmlParseError(doc);
    if (parseError) return null;

    const paragraphs = getDocumentParagraphNodes(doc);
    const firstParagraph = paragraphs[0] || null;
    if (!firstParagraph) return null;

    const numIdNodes = Array.from(firstParagraph.getElementsByTagNameNS('*', 'numId'));
    for (const numIdNode of numIdNodes) {
        const numId = getElementId(numIdNode, ['w:val', 'val']);
        if (numId != null) return String(numId);
    }
    return null;
}

function applyStartOverrideToNumberingXml(numberingXml, targetNumId, startAt, options = {}) {
    if (!numberingXml || !targetNumId || !Number.isInteger(startAt) || startAt < 1) return numberingXml;
    const setAbstractStartOverride = options.setAbstractStartOverride !== false;
    const parser = createParser();
    const serializer = createSerializer();
    const numberingDoc = parser.parseFromString(String(numberingXml || ''), 'application/xml');
    const parseError = getXmlParseError(numberingDoc);
    if (parseError) return numberingXml;

    const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));
    const target = nums.find(node => {
        const id = getElementId(node, ['w:numId', 'numId']);
        return id != null && String(id) === String(targetNumId);
    });
    if (!target) return numberingXml;

    const abstractNumIdNode = Array.from(target.getElementsByTagNameNS('*', 'abstractNumId'))[0] || null;
    const abstractNumId = getElementId(abstractNumIdNode, ['w:val', 'val']);

    let lvlOverride = Array.from(target.getElementsByTagNameNS('*', 'lvlOverride'))
        .find(node => {
            const ilvl = getElementId(node, ['w:ilvl', 'ilvl']);
            return ilvl === 0;
        }) || null;
    if (!lvlOverride) {
        lvlOverride = numberingDoc.createElementNS(
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w:lvlOverride'
        );
        lvlOverride.setAttribute('w:ilvl', '0');
        target.appendChild(lvlOverride);
    }

    let startOverride = Array.from(lvlOverride.getElementsByTagNameNS('*', 'startOverride'))[0] || null;
    if (!startOverride) {
        startOverride = numberingDoc.createElementNS(
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w:startOverride'
        );
        lvlOverride.appendChild(startOverride);
    }
    setElementVal(startOverride, startAt);

    // Optional compatibility mode: set abstract-level <w:start> in addition to
    // num-level <w:startOverride>. This can influence other lists sharing that
    // abstract definition in some renderers, so callers may disable it.
    if (setAbstractStartOverride && abstractNumId != null) {
        const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
        const abstractNum = abstractNums.find(node => {
            const id = getElementId(node, ['w:abstractNumId', 'abstractNumId']);
            return id != null && id === abstractNumId;
        }) || null;
        if (abstractNum) {
            let lvl = Array.from(abstractNum.getElementsByTagNameNS('*', 'lvl'))
                .find(node => {
                    const ilvl = getElementId(node, ['w:ilvl', 'ilvl']);
                    return ilvl === 0;
                }) || null;
            if (!lvl) {
                lvl = numberingDoc.createElementNS(
                    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'w:lvl'
                );
                lvl.setAttribute('w:ilvl', '0');
                abstractNum.appendChild(lvl);
            }

            let startNode = Array.from(lvl.getElementsByTagNameNS('*', 'start'))[0] || null;
            if (!startNode) {
                startNode = numberingDoc.createElementNS(
                    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'w:start'
                );
                lvl.insertBefore(startNode, lvl.firstChild);
            }
            setElementVal(startNode, startAt);
        }
    }

    return serializer.serializeToString(numberingDoc);
}

/**
 * Builds a plan for single-line structural list fallback.
 *
 * @param {{
 *   oxml: string,
 *   originalText: string,
 *   modifiedText: string,
 *   allowExistingList?: boolean
 * }} options - Input values for fallback detection
 * @returns {{ listInput: string, numberingKey: string, originalText: string, wasListParagraph: boolean, startAt: number|null }|null}
 */
export function buildSingleLineListStructuralFallbackPlan(options = {}) {
    const oxml = String(options.oxml || '');
    const originalText = String(options.originalText || '');
    const modifiedText = String(options.modifiedText || '');
    const allowExistingList = options.allowExistingList === true;

    if (!oxml.trim() || !modifiedText.trim()) return null;

    const paragraph = getFirstParagraphFromOxml(oxml);
    if (!paragraph) return null;

    const existingListInfo = getParagraphListInfo(paragraph);
    if (existingListInfo && !allowExistingList) return null;

    const modifiedCleanText = preprocessMarkdown(modifiedText).cleanText || modifiedText;
    const modifiedCandidate = parseSingleLineListCandidate(modifiedText)
        || parseSingleLineListCandidate(modifiedCleanText);
    if (!modifiedCandidate) return null;

    const currentCandidate = parseSingleLineListCandidate(originalText);
    const sameRawText =
        normalizeWhitespaceForTargeting(originalText) === normalizeWhitespaceForTargeting(modifiedCleanText);
    const sameListText = !!currentCandidate
        && currentCandidate.type === modifiedCandidate.type
        && currentCandidate.normalizedContent === modifiedCandidate.normalizedContent;
    if (!sameRawText && !sameListText) return null;

    return {
        listInput: `${modifiedCandidate.marker} ${modifiedCandidate.contentText}`.trim(),
        numberingKey: `${modifiedCandidate.type}:${modifiedCandidate.numberingStyle}:single`,
        originalText,
        wasListParagraph: !!existingListInfo,
        startAt: modifiedCandidate.startAt
    };
}

/**
 * Executes a single-line structural list fallback plan.
 *
 * Returns list fragment OOXML (not wrapped package), plus `numberingXml`.
 *
 * @param {{ listInput: string, numberingKey: string, originalText?: string, startAt?: number|null }} plan - Fallback plan
 * @param {{
 *   author?: string,
 *   generateRedlines?: boolean,
 *   pipeline?: ReconciliationPipeline,
 *   setAbstractStartOverride?: boolean
 * }} [options={}] - Execution options
 * @returns {Promise<{
 *   hasChanges: boolean,
 *   oxml: string,
 *   numberingXml: string|null,
 *   includeNumbering: boolean,
 *   listStructuralFallbackApplied: boolean,
 *   listStructuralFallbackKey: string|null,
 *   warnings: string[]
 * }>}
 */
export async function executeSingleLineListStructuralFallback(plan, options = {}) {
    if (!plan || !plan.listInput) {
        return {
            hasChanges: false,
            oxml: '',
            numberingXml: null,
            includeNumbering: false,
            listStructuralFallbackApplied: false,
            listStructuralFallbackKey: null,
            warnings: ['Single-line list fallback plan missing']
        };
    }

    const author = options.author || 'AI';
    const generateRedlines = options.generateRedlines ?? true;
    const pipeline = options.pipeline || new ReconciliationPipeline({ author, generateRedlines });

    const result = await pipeline.executeListGeneration(
        plan.listInput,
        null,
        null,
        String(plan.originalText || '')
    );

    const rawOxml = result?.oxml || result?.ooxml || '';
    const oxml = trimTrailingBlankParagraph(rawOxml);
    const generatedNumId = extractFirstParagraphNumIdFromOxml(oxml);
    const numberingXmlWithStart = applyStartOverrideToNumberingXml(
        result?.numberingXml || null,
        generatedNumId,
        Number.isInteger(plan?.startAt) ? plan.startAt : null,
        {
            setAbstractStartOverride: options.setAbstractStartOverride
        }
    );
    const isValid = result?.isValid !== false;
    if (!oxml || !isValid) {
        return {
            hasChanges: false,
            oxml: '',
            numberingXml: null,
            includeNumbering: false,
            listStructuralFallbackApplied: false,
            listStructuralFallbackKey: plan.numberingKey || null,
            warnings: ['Single-line list fallback produced no valid OOXML']
        };
    }

    return {
        hasChanges: true,
        oxml,
        numberingXml: numberingXmlWithStart,
        includeNumbering: true,
        listStructuralFallbackApplied: true,
        listStructuralFallbackKey: plan.numberingKey || null,
        listStructuralFallbackStartAt: Number.isInteger(plan?.startAt) ? plan.startAt : null,
        warnings: ['Single-line list structural fallback applied']
    };
}
