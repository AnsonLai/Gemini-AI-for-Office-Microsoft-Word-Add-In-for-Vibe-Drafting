/**
 * Shared list-targeting helpers for per-paragraph redline callers.
 */

import {
    WORD_MAIN_NS,
    getParagraphText,
    normalizeWhitespaceForTargeting
} from './paragraph-targeting.js';

function getFirstDescendantByLocalName(node, localName) {
    if (!node || typeof node.getElementsByTagNameNS !== 'function') return null;
    const namespaced = node.getElementsByTagNameNS(WORD_MAIN_NS, localName);
    if (namespaced.length > 0) return namespaced[0];
    const anyNs = node.getElementsByTagNameNS('*', localName);
    return anyNs.length > 0 ? anyNs[0] : null;
}

function readValAttribute(element) {
    if (!element) return null;
    if (typeof element.getAttributeNS === 'function') {
        const namespaced = element.getAttributeNS(WORD_MAIN_NS, 'val');
        if (namespaced) return namespaced;
    }
    return element.getAttribute('w:val') || element.getAttribute('val') || null;
}

function parseOutlineLevelFromMarker(marker) {
    const normalized = String(marker || '').trim();
    if (!/^\d+(?:\.\d+)+\.?$/.test(normalized)) return null;
    const parts = normalized.replace(/\.$/, '').split('.');
    return Math.max(0, parts.length - 1);
}

const REDUNDANT_LIST_PREFIX_REGEX = /^(?:(?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|[ivxlcIVXLC]+\.|[-*+\u2022]))\s+/;

/**
 * Strips redundant manual list markers from the start of list item text.
 *
 * Useful when model output contains doubled markers like:
 * - "2.1. - Item text"
 * - "- 2.1. Item text"
 *
 * @param {string} text - Candidate list item text
 * @returns {string}
 */
export function stripRedundantLeadingListMarkers(text) {
    let value = String(text || '').trim();
    let passes = 0;
    while (passes < 4 && REDUNDANT_LIST_PREFIX_REGEX.test(value)) {
        value = value.replace(REDUNDANT_LIST_PREFIX_REGEX, '').trimStart();
        passes++;
    }
    return value.trim();
}

function parseModifiedListItems(modifiedText) {
    const rawLines = String(modifiedText || '').split(/\r?\n/g);
    const items = [];
    let hasListMarkers = false;

    for (const rawLine of rawLines) {
        const line = rawLine.trimEnd();
        if (!line.trim()) continue;

        const markerMatch = line.match(/^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|[ivxlcIVXLC]+\.|[-*+\u2022]))\s+(.*)$/);
        if (markerMatch) {
            hasListMarkers = true;
            const marker = markerMatch[2];
            const markerType = /^[-*+\u2022]$/.test(marker) ? 'bullet' : 'numbered';
            const level = Math.floor((markerMatch[1] || '').length / 2);
            items.push({
                kind: 'list',
                markerType,
                level,
                marker,
                outlineLevel: markerType === 'numbered' ? parseOutlineLevelFromMarker(marker) : null,
                text: stripRedundantLeadingListMarkers(markerMatch[3])
            });
            continue;
        }

        items.push({
            kind: 'text',
            text: line.trim()
        });
    }

    return { items, hasListMarkers };
}

function toMarkdownLine(level, markerType, text) {
    const indent = '  '.repeat(Math.max(0, level));
    const marker = markerType === 'numbered' ? '1.' : '-';
    return `${indent}${marker} ${String(text || '').trim()}`.trimEnd();
}

function isNormalizedTextEqual(a, b) {
    return normalizeWhitespaceForTargeting(a) === normalizeWhitespaceForTargeting(b);
}

function resolveInsertionLevel(item, anchorLevel, baselineLevel) {
    if (Number.isInteger(item?.outlineLevel)) {
        return Math.max(0, item.outlineLevel);
    }
    return Math.max(0, anchorLevel + ((item?.level || 0) - baselineLevel));
}

function shouldPromoteBulletInsertionsToChildDepth(parsedItems, normalizedTargetText, anchorLevel) {
    if (!Array.isArray(parsedItems) || parsedItems.length < 2) return false;
    if (!Number.isInteger(anchorLevel) || anchorLevel < 1) return false;

    const firstItem = parsedItems[0];
    const trailingListItems = parsedItems.slice(1).filter(item => item.kind === 'list');
    if (trailingListItems.length === 0) return false;
    if (trailingListItems.some(item => item.markerType !== 'bullet')) return false;
    if (trailingListItems.some(item => Number.isInteger(item.outlineLevel))) return false;

    if (firstItem?.kind === 'text') {
        return isNormalizedTextEqual(firstItem.text, normalizedTargetText);
    }

    if (firstItem?.kind === 'list' && firstItem.markerType === 'numbered') {
        return isNormalizedTextEqual(firstItem.text, normalizedTargetText);
    }

    return false;
}

function promoteBulletInsertionsToChildDepth(entries, anchorLevel) {
    return entries.map(entry => {
        const relativeDepth = Math.max(0, (entry.ilvl || 0) - anchorLevel);
        return {
            ...entry,
            ilvl: Math.min(8, anchorLevel + 1 + relativeDepth)
        };
    });
}

function buildListEntriesForInsertion(parsedItems, normalizedTargetText, anchorLevel, defaultMarkerType) {
    const firstItem = parsedItems[0];
    const trailingListItems = parsedItems.slice(1).filter(item => item.kind === 'list');

    // Pattern A: plain anchor text + list item lines (target line + inserted lines)
    if (
        firstItem?.kind === 'text' &&
        isNormalizedTextEqual(firstItem.text, normalizedTargetText) &&
        trailingListItems.length > 0
    ) {
        const firstTrailingLevel = trailingListItems[0].level;
        return trailingListItems.map(item => ({
            ilvl: resolveInsertionLevel(item, anchorLevel, firstTrailingLevel),
            markerType: item.markerType || defaultMarkerType,
            text: item.text
        }));
    }

    // Pattern B: list lines where first line repeats target item.
    if (parsedItems.every(item => item.kind === 'list')) {
        const firstList = parsedItems[0];
        if (!firstList || !isNormalizedTextEqual(firstList.text, normalizedTargetText)) {
            return null;
        }
        const firstLevel = firstList.level;
        return parsedItems
            .slice(1)
            .map(item => ({
                ilvl: resolveInsertionLevel(item, anchorLevel, firstLevel),
                markerType: item.markerType || defaultMarkerType,
                text: item.text
            }))
            .filter(item => item.text);
    }

    return null;
}

/**
 * Reads list numbering metadata from paragraph OOXML.
 *
 * @param {Element} paragraph - Paragraph element
 * @returns {{ numId: string, ilvl: number }|null}
 */
export function getParagraphListInfo(paragraph) {
    if (!paragraph) return null;

    const pPr = getFirstDescendantByLocalName(paragraph, 'pPr');
    if (!pPr) return null;
    const numPr = getFirstDescendantByLocalName(pPr, 'numPr');
    if (!numPr) return null;
    const numIdEl = getFirstDescendantByLocalName(numPr, 'numId');
    if (!numIdEl) return null;

    const numId = readValAttribute(numIdEl);
    if (!numId) return null;

    const ilvlEl = getFirstDescendantByLocalName(numPr, 'ilvl');
    const ilvlRaw = readValAttribute(ilvlEl);
    const ilvl = Number.parseInt(ilvlRaw || '0', 10);

    return {
        numId: String(numId),
        ilvl: Number.isFinite(ilvl) ? ilvl : 0
    };
}

/**
 * Collects contiguous sibling list paragraphs sharing the same `numId`.
 *
 * @param {Element} targetParagraph - Target paragraph
 * @returns {Element[]|null}
 */
export function collectContiguousListParagraphBlock(targetParagraph) {
    const targetInfo = getParagraphListInfo(targetParagraph);
    if (!targetInfo) return null;

    const parent = targetParagraph.parentNode;
    if (!parent) return null;

    const siblings = Array.from(parent.childNodes || []).filter(
        node =>
            node &&
            node.nodeType === 1 &&
            node.namespaceURI === WORD_MAIN_NS &&
            node.localName === 'p'
    );
    const targetIndex = siblings.indexOf(targetParagraph);
    if (targetIndex < 0) return null;

    let start = targetIndex;
    while (start > 0) {
        const prevInfo = getParagraphListInfo(siblings[start - 1]);
        if (!prevInfo || prevInfo.numId !== targetInfo.numId) break;
        start--;
    }

    let end = targetIndex;
    while (end < siblings.length - 1) {
        const nextInfo = getParagraphListInfo(siblings[end + 1]);
        if (!nextInfo || nextInfo.numId !== targetInfo.numId) break;
        end++;
    }

    return siblings.slice(start, end + 1);
}

/**
 * Synthesizes block-level list markdown edits when a single list item receives
 * multiline list content (for example insert-between-item intent).
 *
 * @param {Element} targetParagraph - Resolved target paragraph
 * @param {string} modifiedText - Proposed replacement text
 * @param {{
 *   currentParagraphText?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} [options] - Optional logging/context options
 * @returns {{ paragraphs: Element[], originalText: string, modifiedText: string }|null}
 */
export function synthesizeExpandedListScopeEdit(targetParagraph, modifiedText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};

    const rawModified = String(modifiedText || '');
    if (!rawModified.includes('\n')) return null;

    const targetListInfo = getParagraphListInfo(targetParagraph);
    if (!targetListInfo) return null;

    const blockParagraphs = collectContiguousListParagraphBlock(targetParagraph);
    if (!blockParagraphs || blockParagraphs.length === 0) return null;

    const parsed = parseModifiedListItems(rawModified);
    if (!parsed.hasListMarkers || parsed.items.length < 2) return null;

    const normalizedTargetText = normalizeWhitespaceForTargeting(
        options.currentParagraphText || getParagraphText(targetParagraph)
    );
    const listItemsOnly = parsed.items.filter(item => item.kind === 'list');
    const firstListType = listItemsOnly[0]?.markerType || 'bullet';

    const blockInfos = blockParagraphs.map(paragraph => ({
        paragraph,
        list: getParagraphListInfo(paragraph),
        text: String(getParagraphText(paragraph) || '').trim()
    }));
    const targetIndex = blockParagraphs.indexOf(targetParagraph);
    if (targetIndex < 0) return null;

    const baseLevel = Math.min(...blockInfos.map(info => info.list?.ilvl ?? 0));
    const originalMarkdownLines = blockInfos.map(info =>
        toMarkdownLine((info.list?.ilvl ?? 0) - baseLevel, firstListType, info.text)
    );

    let replacementEntries = null;
    const firstItem = parsed.items[0];
    const trailingListItems = parsed.items.slice(1).filter(item => item.kind === 'list');

    if (firstItem?.kind === 'text' && isNormalizedTextEqual(firstItem.text, normalizedTargetText) && trailingListItems.length > 0) {
        const anchorLevel = Math.max(0, (blockInfos[targetIndex].list?.ilvl ?? 0) - baseLevel);
        const firstTrailingLevel = trailingListItems[0].level;
        replacementEntries = [
            {
                level: anchorLevel,
                markerType: firstListType,
                text: blockInfos[targetIndex].text
            },
            ...trailingListItems.map(item => ({
                level: resolveInsertionLevel(item, anchorLevel, firstTrailingLevel),
                markerType: item.markerType || firstListType,
                text: item.text
            }))
        ];
    } else if (parsed.items.every(item => item.kind === 'list')) {
        const anchorLevel = Math.max(0, (blockInfos[targetIndex].list?.ilvl ?? 0) - baseLevel);
        const firstLevel = parsed.items[0].level;
        replacementEntries = parsed.items.map(item => ({
            level: resolveInsertionLevel(item, anchorLevel, firstLevel),
            markerType: item.markerType || firstListType,
            text: item.text
        }));
    } else {
        onWarn('[List] Multiline list edit did not match supported insertion/replace patterns; skipping list-block synthesis.');
        return null;
    }

    const replacementLines = replacementEntries.map(entry => toMarkdownLine(entry.level, entry.markerType, entry.text));
    const modifiedMarkdownLines = originalMarkdownLines
        .slice(0, targetIndex)
        .concat(replacementLines)
        .concat(originalMarkdownLines.slice(targetIndex + 1));

    const originalText = originalMarkdownLines.join('\n');
    const nextModifiedText = modifiedMarkdownLines.join('\n');
    if (nextModifiedText === originalText) return null;

    onInfo('[List] Expanded single-item list edit to contiguous list block for stable middle insertion.');
    return {
        paragraphs: blockParagraphs,
        originalText,
        modifiedText: nextModifiedText
    };
}

/**
 * Plans insertion-only list edits for multiline middle-insert requests.
 *
 * This returns only new list items to insert after the target paragraph, so
 * callers can emit insertion-only redlines instead of deleting/reinserting
 * whole list blocks.
 *
 * @param {Element} targetParagraph - Resolved target list paragraph
 * @param {string} modifiedText - Proposed replacement text
 * @param {{
 *   currentParagraphText?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} [options] - Optional context/log callbacks
 * @returns {{ targetParagraph: Element, numId: string, entries: Array<{ ilvl: number, text: string, markerType: 'bullet'|'numbered' }> }|null}
 */
export function planListInsertionOnlyEdit(targetParagraph, modifiedText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};

    const rawModified = String(modifiedText || '');
    if (!rawModified.includes('\n')) return null;

    const targetListInfo = getParagraphListInfo(targetParagraph);
    if (!targetListInfo) return null;

    const parsed = parseModifiedListItems(rawModified);
    if (!parsed.hasListMarkers || parsed.items.length < 2) return null;

    const normalizedTargetText = normalizeWhitespaceForTargeting(
        options.currentParagraphText || getParagraphText(targetParagraph)
    );
    const listItemsOnly = parsed.items.filter(item => item.kind === 'list');
    const defaultMarkerType = listItemsOnly[0]?.markerType || 'bullet';
    const anchorLevel = Math.max(0, targetListInfo.ilvl);
    let entries = buildListEntriesForInsertion(parsed.items, normalizedTargetText, anchorLevel, defaultMarkerType);

    if (!entries || entries.length === 0) {
        onWarn('[List] Could not derive insertion-only entries from multiline list edit.');
        return null;
    }

    if (shouldPromoteBulletInsertionsToChildDepth(parsed.items, normalizedTargetText, anchorLevel)) {
        entries = promoteBulletInsertionsToChildDepth(entries, anchorLevel);
        onInfo('[List] Promoted bullet insertion to child depth for nested numbered-list intent.');
    }

    onInfo('[List] Planned insertion-only list redline entries (no block rewrite).');
    return {
        targetParagraph,
        numId: targetListInfo.numId,
        entries
    };
}
