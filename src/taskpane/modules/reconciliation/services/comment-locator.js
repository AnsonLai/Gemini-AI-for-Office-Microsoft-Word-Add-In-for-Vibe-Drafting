/**
 * Comment text location and marker injection helpers.
 */

import { NS_W } from '../core/types.js';
import { getElementsByTag, getFirstElementByTag } from '../core/xml-query.js';

/**
 * Builds a paragraph text index in a single pass for repeated lookups.
 *
 * @param {Element} paragraph - w:p element
 * @returns {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }}
 */
export function createParagraphTextIndex(paragraph) {
    const runs = getElementsByTag(paragraph, 'w:r');
    const runOffsets = [];
    let fullText = '';

    for (const run of runs) {
        const start = fullText.length;
        const textNodes = getElementsByTag(run, 'w:t');
        for (const textNode of textNodes) {
            fullText += textNode.textContent || '';
        }
        runOffsets.push({ run, start, end: fullText.length });
    }

    return { fullText, runOffsets };
}

/**
 * Finds text within a prebuilt paragraph text index.
 *
 * @param {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }} paragraphIndex - Prebuilt index
 * @param {string} searchText - Text to locate
 * @returns {{ found: boolean, startRun?: Element, startOffset?: number, endRun?: Element, endOffset?: number }}
 */
export function findTextInParagraphIndex(paragraphIndex, searchText) {
    const searchIndex = paragraphIndex.fullText.indexOf(searchText);
    if (searchIndex === -1) {
        return { found: false };
    }

    const searchEnd = searchIndex + searchText.length;
    let startRun = null;
    let endRun = null;
    let startOffset = 0;
    let endOffset = 0;

    for (const { run, start, end } of paragraphIndex.runOffsets) {
        if (searchIndex >= start && searchIndex < end) {
            startRun = run;
            startOffset = searchIndex - start;
        }
        if (searchEnd > start && searchEnd <= end) {
            endRun = run;
            endOffset = searchEnd - start;
        }
    }

    return {
        found: true,
        startRun,
        startOffset,
        endRun,
        endOffset
    };
}

function cloneRunWithText(xmlDoc, rPr, newText) {
    const newRun = xmlDoc.createElementNS(NS_W, 'w:r');
    if (rPr) {
        newRun.appendChild(rPr.cloneNode(true));
    }

    const newTextNode = xmlDoc.createElementNS(NS_W, 'w:t');
    newTextNode.setAttribute('xml:space', 'preserve');
    newTextNode.textContent = newText;
    newRun.appendChild(newTextNode);
    return newRun;
}

/**
 * Injects comment markers around text in a paragraph.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} paragraph - w:p element
 * @param {string} textToFind - Target text
 * @param {number} commentId - Comment id
 * @param {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }|null} [paragraphIndex=null] - Optional prebuilt index
 * @returns {boolean}
 */
export function injectMarkersIntoParagraph(xmlDoc, paragraph, textToFind, commentId, paragraphIndex = null) {
    const activeIndex = paragraphIndex || createParagraphTextIndex(paragraph);
    const location = findTextInParagraphIndex(activeIndex, textToFind);
    if (!location.found || !location.startRun) {
        return false;
    }

    const startMarker = xmlDoc.createElementNS(NS_W, 'w:commentRangeStart');
    startMarker.setAttribute('w:id', String(commentId));

    const endMarker = xmlDoc.createElementNS(NS_W, 'w:commentRangeEnd');
    endMarker.setAttribute('w:id', String(commentId));

    const referenceRun = xmlDoc.createElementNS(NS_W, 'w:r');
    const reference = xmlDoc.createElementNS(NS_W, 'w:commentReference');
    reference.setAttribute('w:id', String(commentId));
    referenceRun.appendChild(reference);

    if (location.startRun === location.endRun) {
        const run = location.startRun;
        const textNode = getFirstElementByTag(run, 'w:t');
        if (!textNode) {
            run.parentNode.insertBefore(startMarker, run);
            if (run.nextSibling) {
                run.parentNode.insertBefore(endMarker, run.nextSibling);
                endMarker.parentNode.insertBefore(referenceRun, endMarker.nextSibling);
            } else {
                run.parentNode.appendChild(endMarker);
                run.parentNode.appendChild(referenceRun);
            }
            return true;
        }

        const fullText = textNode.textContent || '';
        const beforeText = fullText.substring(0, location.startOffset);
        const highlightedText = fullText.substring(location.startOffset, location.endOffset);
        const afterText = fullText.substring(location.endOffset);
        const rPr = getFirstElementByTag(run, 'w:rPr');
        const parent = run.parentNode;

        if (beforeText) {
            parent.insertBefore(cloneRunWithText(xmlDoc, rPr, beforeText), run);
        }

        parent.insertBefore(startMarker, run);
        textNode.textContent = highlightedText;

        if (run.nextSibling) {
            parent.insertBefore(endMarker, run.nextSibling);
        } else {
            parent.appendChild(endMarker);
        }
        parent.insertBefore(referenceRun, endMarker.nextSibling || null);

        if (afterText) {
            parent.insertBefore(cloneRunWithText(xmlDoc, rPr, afterText), referenceRun.nextSibling || null);
        }

        return true;
    }

    const startTextNode = getFirstElementByTag(location.startRun, 'w:t');
    if (startTextNode && location.startOffset > 0) {
        const fullText = startTextNode.textContent || '';
        const beforeText = fullText.substring(0, location.startOffset);
        const highlightedStart = fullText.substring(location.startOffset);

        if (beforeText) {
            const rPr = getFirstElementByTag(location.startRun, 'w:rPr');
            location.startRun.parentNode.insertBefore(cloneRunWithText(xmlDoc, rPr, beforeText), location.startRun);
        }
        startTextNode.textContent = highlightedStart;
    }

    location.startRun.parentNode.insertBefore(startMarker, location.startRun);

    const endRun = location.endRun || location.startRun;
    const endTextNode = getFirstElementByTag(endRun, 'w:t');
    if (endTextNode && location.endOffset < (endTextNode.textContent || '').length) {
        const fullText = endTextNode.textContent || '';
        const highlightedEnd = fullText.substring(0, location.endOffset);
        const afterText = fullText.substring(location.endOffset);

        endTextNode.textContent = highlightedEnd;

        if (afterText) {
            const rPr = getFirstElementByTag(endRun, 'w:rPr');
            if (endRun.nextSibling) {
                endRun.parentNode.insertBefore(cloneRunWithText(xmlDoc, rPr, afterText), endRun.nextSibling);
            } else {
                endRun.parentNode.appendChild(cloneRunWithText(xmlDoc, rPr, afterText));
            }
        }
    }

    if (endRun.nextSibling) {
        endRun.parentNode.insertBefore(endMarker, endRun.nextSibling);
        endMarker.parentNode.insertBefore(referenceRun, endMarker.nextSibling);
    } else {
        endRun.parentNode.appendChild(endMarker);
        endRun.parentNode.appendChild(referenceRun);
    }

    return true;
}
