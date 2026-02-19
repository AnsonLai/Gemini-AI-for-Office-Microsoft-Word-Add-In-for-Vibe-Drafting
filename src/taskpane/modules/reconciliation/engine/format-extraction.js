/**
 * OOXML format extraction utilities.
 *
 * Provides shared paragraph filtering, text-span extraction, and run-format
 * extraction for formatting-aware reconciliation flows.
 */

import { extractFormatFromRPr } from './rpr-helpers.js';
import { advanceOffsetForParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { log } from '../adapters/logger.js';
import { getElementsByTag, getElementsByTagNS, getFirstElementByTag } from '../core/xml-query.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function isWordElement(node, localName) {
    if (!node || node.nodeType !== 1) return false;
    if (node.namespaceURI === NS_W && node.localName === localName) return true;
    const nodeName = String(node.nodeName || '');
    return nodeName === `w:${localName}` || nodeName === localName;
}

function isExcludedRevisionContainer(node) {
    if (!node || node.nodeType !== 1 || node.namespaceURI !== NS_W) return false;
    return node.localName === 'del' || node.localName === 'moveFrom';
}

function collectParagraphRuns(paragraph) {
    const runs = [];
    const stack = Array.from(paragraph?.childNodes || []).reverse();

    while (stack.length > 0) {
        const node = stack.pop();
        if (!node || node.nodeType !== 1) continue;
        if (isExcludedRevisionContainer(node)) continue;

        if (isWordElement(node, 'r')) {
            runs.push(node);
            continue;
        }

        const children = Array.from(node.childNodes || []);
        for (let i = children.length - 1; i >= 0; i -= 1) {
            stack.push(children[i]);
        }
    }

    return runs;
}

/**
 * Returns only document body paragraphs, excluding comments/footnotes/endnotes.
 *
 * @param {Document} xmlDoc - XML document
 * @returns {Element[]}
 */
export function getDocumentParagraphs(xmlDoc) {
    const excludedContainers = new Set(['comment', 'footnote', 'endnote']);
    const allParagraphs = getElementsByTagNS(xmlDoc, '*', 'p');

    return allParagraphs.filter(p => {
        let node = p.parentNode;
        while (node && node.nodeName) {
            const localName = String(node.localName || '').toLowerCase();
            if (excludedContainers.has(localName)) {
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
        const runs = collectParagraphRuns(p);
        for (const run of runs) {
            const rPr = getFirstElementByTag(run, 'w:rPr');
            Array.from(run.childNodes || []).forEach(rc => {
                if (isWordElement(rc, 't')) {
                    const text = rc.textContent || '';
                    if (text.length > 0) {
                        textSpans.push({
                            charStart: charOffset,
                            charEnd: charOffset + text.length,
                            textElement: rc,
                            runElement: run,
                            paragraph: p,
                            container: run.parentNode,
                            rPr
                        });
                        charOffset += text.length;
                    }
                } else if (isWordElement(rc, 'br') || isWordElement(rc, 'cr') || isWordElement(rc, 'tab') || isWordElement(rc, 'noBreakHyphen')) {
                    textSpans.push({
                        charStart: charOffset,
                        charEnd: charOffset + 1,
                        textElement: rc,
                        runElement: run,
                        paragraph: p,
                        container: run.parentNode,
                        rPr
                    });
                    charOffset += 1;
                }
            });
        }
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
        if (isWordElement(child, 'rPr')) {
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
        if (isWordElement(child, 't')) {
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
            if (isWordElement(child, 'pPr')) {
                for (const pChild of Array.from(child.childNodes)) {
                    if (isWordElement(pChild, 'rPr')) {
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

        const runs = collectParagraphRuns(p);
        for (const run of runs) {
            charOffset = processRunForFormatting(run, p, charOffset, textSpans, existingFormatHints, pFormat);
        }
        charOffset = advanceOffsetForParagraphBoundary(charOffset, pIndex, paragraphs.length);
    }

    log(`[OxmlEngine] Extracted ${textSpans.length} text spans, ${existingFormatHints.length} format hints`);
    return { existingFormatHints, textSpans, paragraphs };
}
