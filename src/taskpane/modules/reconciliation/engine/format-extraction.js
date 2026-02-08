/**
 * OOXML format extraction utilities.
 *
 * Provides shared paragraph filtering, text-span extraction, and run-format
 * extraction for formatting-aware reconciliation flows.
 */

import { extractFormatFromRPr } from './rpr-helpers.js';
import { advanceOffsetForParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { log } from '../adapters/logger.js';
import { getElementsByTag, getFirstElementByTag } from '../core/xml-query.js';

/**
 * Returns only document body paragraphs, excluding comments/footnotes/endnotes.
 *
 * @param {Document} xmlDoc - XML document
 * @returns {Element[]}
 */
export function getDocumentParagraphs(xmlDoc) {
    const excludedContainers = ['w:comment', 'w:footnote', 'w:endnote'];
    const allParagraphs = getElementsByTag(xmlDoc, 'w:p');

    return allParagraphs.filter(p => {
        let node = p.parentNode;
        while (node && node.nodeName) {
            if (excludedContainers.includes(node.nodeName)) {
                return false;
            }
            node = node.parentNode;
        }
        return true;
    });
}

/**
 * Builds linear text spans from paragraph elements.
 *
 * @param {Element[]} paragraphs - Paragraph nodes
 * @returns {{ textSpans: Array, charOffset: number }}
 */
export function buildTextSpansFromParagraphs(paragraphs) {
    const textSpans = [];
    let charOffset = 0;

    for (let pIndex = 0; pIndex < paragraphs.length; pIndex++) {
        const p = paragraphs[pIndex];
        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === 'w:r') {
                const rPr = getFirstElementByTag(child, 'w:rPr');
                Array.from(child.childNodes).forEach(rc => {
                    if (rc.nodeName === 'w:t') {
                        const text = rc.textContent || '';
                        if (text.length > 0) {
                            textSpans.push({
                                charStart: charOffset,
                                charEnd: charOffset + text.length,
                                textElement: rc,
                                runElement: child,
                                paragraph: p,
                                container: p.parentNode,
                                rPr
                            });
                            charOffset += text.length;
                        }
                    } else if (rc.nodeName === 'w:br' || rc.nodeName === 'w:cr' || rc.nodeName === 'w:tab' || rc.nodeName === 'w:noBreakHyphen') {
                        textSpans.push({
                            charStart: charOffset,
                            charEnd: charOffset + 1,
                            textElement: rc,
                            runElement: child,
                            paragraph: p,
                            container: p.parentNode,
                            rPr
                        });
                        charOffset += 1;
                    }
                });
            } else if (child.nodeName === 'w:hyperlink') {
                Array.from(child.childNodes).forEach(hc => {
                    if (hc.nodeName === 'w:r') {
                        const rPr = getFirstElementByTag(hc, 'w:rPr');
                        Array.from(hc.childNodes).forEach(rc => {
                            if (rc.nodeName === 'w:t') {
                                const text = rc.textContent || '';
                                if (text.length > 0) {
                                    textSpans.push({
                                        charStart: charOffset,
                                        charEnd: charOffset + text.length,
                                        textElement: rc,
                                        runElement: hc,
                                        paragraph: p,
                                        container: child,
                                        rPr
                                    });
                                    charOffset += text.length;
                                }
                            } else if (rc.nodeName === 'w:br' || rc.nodeName === 'w:cr' || rc.nodeName === 'w:tab' || rc.nodeName === 'w:noBreakHyphen') {
                                textSpans.push({
                                    charStart: charOffset,
                                    charEnd: charOffset + 1,
                                    textElement: rc,
                                    runElement: hc,
                                    paragraph: p,
                                    container: child,
                                    rPr
                                });
                                charOffset += 1;
                            }
                        });
                    }
                });
            }
        });
        charOffset = advanceOffsetForParagraphBoundary(charOffset, pIndex, paragraphs.length);
    }

    return { textSpans, charOffset };
}

/**
 * Processes a single run element and appends extracted spans/hints.
 *
 * @param {Element} run - `w:r` element
 * @param {Element} paragraph - Parent paragraph
 * @param {number} charOffset - Start offset
 * @param {Array} textSpans - Span collection (mutated)
 * @param {Array} formatHints - Format hint collection (mutated)
 * @param {Object|null} pFormat - Paragraph-level run format defaults
 * @returns {number}
 */
export function processRunForFormatting(run, paragraph, charOffset, textSpans, formatHints, pFormat = null) {
    let rPr = null;
    for (const child of Array.from(run.childNodes)) {
        if (child.nodeName === 'w:rPr') {
            rPr = child;
            break;
        }
    }

    const format = extractFormatFromRPr(rPr);

    if (pFormat) {
        if (pFormat.bold && !format.bold) format.bold = true;
        if (pFormat.italic && !format.italic) format.italic = true;
        if (pFormat.underline && !format.underline) format.underline = true;
        if (pFormat.strikethrough && !format.strikethrough) format.strikethrough = true;
    }

    format.hasFormatting = format.bold || format.italic || format.underline || format.strikethrough;

    let currentOffset = charOffset;
    for (const child of Array.from(run.childNodes)) {
        if (child.nodeName === 'w:t') {
            const text = child.textContent || '';
            if (text.length > 0) {
                const start = currentOffset;
                const end = currentOffset + text.length;

                textSpans.push({
                    charStart: start,
                    charEnd: end,
                    textElement: child,
                    runElement: run,
                    paragraph,
                    rPr,
                    format: { ...format }
                });

                if (format.hasFormatting) {
                    formatHints.push({
                        start,
                        end,
                        format: { ...format },
                        run,
                        rPr
                    });
                }

                currentOffset = end;
            }
        }
    }

    return currentOffset;
}

/**
 * Extracts existing formatting hints from OOXML paragraph runs.
 *
 * @param {Document} xmlDoc - XML document
 * @returns {{ existingFormatHints: Array, textSpans: Array, paragraphs: Element[] }}
 */
export function extractFormattingFromOoxml(xmlDoc) {
    const existingFormatHints = [];
    const textSpans = [];
    let charOffset = 0;

    const paragraphs = getDocumentParagraphs(xmlDoc);

    for (let pIndex = 0; pIndex < paragraphs.length; pIndex++) {
        const p = paragraphs[pIndex];
        let pRPr = null;
        for (const child of Array.from(p.childNodes)) {
            if (child.nodeName === 'w:pPr') {
                for (const pChild of Array.from(child.childNodes)) {
                    if (pChild.nodeName === 'w:rPr') {
                        pRPr = pChild;
                        break;
                    }
                }
                break;
            }
        }

        const pFormat = extractFormatFromRPr(pRPr);
        if (pFormat.hasFormatting) {
            log(`[OxmlEngine] Found paragraph-level formatting: ${JSON.stringify(pFormat)}`);
        }

        for (const child of Array.from(p.childNodes)) {
            if (child.nodeName === 'w:r') {
                charOffset = processRunForFormatting(child, p, charOffset, textSpans, existingFormatHints, pFormat);
            } else if (child.nodeName === 'w:hyperlink') {
                for (const hc of Array.from(child.childNodes)) {
                    if (hc.nodeName === 'w:r') {
                        charOffset = processRunForFormatting(hc, p, charOffset, textSpans, existingFormatHints, pFormat);
                    }
                }
            }
        }
        charOffset = advanceOffsetForParagraphBoundary(charOffset, pIndex, paragraphs.length);
    }

    log(`[OxmlEngine] Extracted ${textSpans.length} text spans, ${existingFormatHints.length} format hints`);
    return { existingFormatHints, textSpans, paragraphs };
}
