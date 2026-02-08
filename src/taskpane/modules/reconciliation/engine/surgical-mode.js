/**
 * Surgical reconciliation mode.
 *
 * This mode performs in-place run-level edits and preserves existing structure,
 * making it safe for tables and other complex OOXML containers.
 */

import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import { computeWordDiffs } from '../pipeline/diff-engine.js';
import { appendParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { NS_W, getNextRevisionId } from '../core/types.js';
import { createSerializer } from '../adapters/xml-adapter.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { buildOverrideRPrXml } from './rpr-helpers.js';
import { getFirstElementByTag } from '../core/xml-query.js';
import {
    createTrackChange,
    createTextRun,
    createFormattedRuns,
    createTextRunWithRPrElement,
    injectFormattingToRPr
} from './run-builders.js';

/**
 * Builds minimal OOXML for a surgical text replacement with track changes.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} originalRun - Original run
 * @param {string} textContent - Replacement text
 * @param {string} author - Author name
 * @param {string} dateStr - ISO date
 * @param {Object} formatToRemove - Format flags to force off
 * @returns {string}
 */
function buildSurgicalReplacementOoxml(xmlDoc, originalRun, textContent, author, dateStr, formatToRemove) {
    const authorName = author || 'Gemini AI';
    const delId = getNextRevisionId();
    const insId = getNextRevisionId();

    const serializer = createSerializer();

    let rPrXml = '';
    const rPr = getFirstElementByTag(originalRun, 'w:rPr');
    if (rPr) {
        rPrXml = serializer.serializeToString(rPr);
        rPrXml = rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '');
    }

    const unformattedRPrXml = buildOverrideRPrXml(xmlDoc, originalRun, formatToRemove, serializer);

    const escapedText = textContent
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

    const delFragment = `<w:del xmlns:w="${NS_W}" w:id="${delId}" w:author="${authorName}" w:date="${dateStr}">` +
        `<w:r>${rPrXml}<w:delText xml:space="preserve">${escapedText}</w:delText></w:r>` +
        `</w:del>`;

    const insFragment = `<w:ins xmlns:w="${NS_W}" w:id="${insId}" w:author="${authorName}" w:date="${dateStr}">` +
        `<w:r>${unformattedRPrXml}<w:t xml:space="preserve">${escapedText}</w:t></w:r>` +
        `</w:ins>`;

    return `${delFragment}${insFragment}`;
}

/**
 * Builds minimal OOXML for an unformatted run (without track wrappers).
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} originalRun - Original run
 * @param {string} textContent - Replacement text
 * @param {Object} formatToRemove - Format flags to force off
 * @returns {string}
 */
function buildUnformattedRunOoxml(xmlDoc, originalRun, textContent, formatToRemove) {
    const serializer = createSerializer();
    const unformattedRPrXml = buildOverrideRPrXml(xmlDoc, originalRun, formatToRemove, serializer);

    const escapedText = textContent
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

    return `<w:r xmlns:w="${NS_W}">${unformattedRPrXml}<w:t xml:space="preserve">${escapedText}</w:t></w:r>`;
}

/**
 * Applies surgical mode reconciliation.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original text
 * @param {string} modifiedText - Modified text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @param {Element|null} [targetParagraph=null] - Optional scope paragraph
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applySurgicalMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true, targetParagraph = null) {
    let fullText = '';
    const textSpans = [];

    const allParagraphs = targetParagraph
        ? [targetParagraph]
        : getDocumentParagraphs(xmlDoc);

    allParagraphs.forEach((p, pIndex) => {
        const container = p.parentNode;

        for (let child = p.firstChild; child; child = child.nextSibling) {
            if (child.nodeName === 'w:r') {
                const runResult = processRunElement(child, p, container, fullText.length, textSpans);
                fullText += runResult.text;
            } else if (child.nodeName === 'w:hyperlink') {
                for (let hc = child.firstChild; hc; hc = hc.nextSibling) {
                    if (hc.nodeName === 'w:r') {
                        const runResult = processRunElement(hc, p, container, fullText.length, textSpans);
                        fullText += runResult.text;
                    }
                }
            }
        }

        fullText = appendParagraphBoundary(fullText, pIndex, allParagraphs.length);
    });

    const diffs = computeWordDiffs(fullText, modifiedText);
    const spanIndex = buildSpanIndex(textSpans);

    let originalPos = 0;
    let newPos = 0;
    const processedSpans = new Set();

    for (const [op, text] of diffs) {
        if (op === 0) {
            const len = text.length;
            const startPos = originalPos;
            const endPos = originalPos + len;

            forEachOverlappingSpan(spanIndex, startPos, endPos, span => {
                const overlapStartOriginal = Math.max(span.charStart, startPos);
                const overlapEndOriginal = Math.min(span.charEnd, endPos);
                const segmentLen = overlapEndOriginal - overlapStartOriginal;
                const relativeOffset = overlapStartOriginal - startPos;
                const overlapStartNew = newPos + relativeOffset;
                const overlapEndNew = overlapStartNew + segmentLen;
                const applicableHints = getApplicableFormatHints(formatHints, overlapStartNew, overlapEndNew);
                reconcileFormattingForTextSpan(xmlDoc, span, overlapStartOriginal, overlapEndOriginal, applicableHints, author, generateRedlines);
            });

            originalPos += len;
            newPos += len;
        } else if (op === -1) {
            processDelete(xmlDoc, spanIndex, originalPos, originalPos + text.length, processedSpans, author, generateRedlines);
            originalPos += text.length;
        } else if (op === 1) {
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.trim().length > 0) {
                processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, processedSpans, author, formatHints, newPos, generateRedlines);
            }
            newPos += text.length;
        }
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Reconciles formatting for a text span segment.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Object} span - Text span
 * @param {number} start - Absolute start
 * @param {number} end - Absolute end
 * @param {Array} applicableHints - Applicable format hints
 * @param {string} author - Author name
 * @param {boolean} generateRedlines - Track change toggle
 */
function reconcileFormattingForTextSpan(xmlDoc, span, start, end, applicableHints, author, generateRedlines) {
    const desiredFormat = {};
    if (applicableHints.length > 0) {
        applicableHints.forEach(h => Object.assign(desiredFormat, h.format));
    }

    const rPr = span.rPr;
    const hasElement = (tagName) => {
        if (!rPr) return false;
        for (let node = rPr.firstChild; node; node = node.nextSibling) {
            if (node.nodeName === tagName) {
                return true;
            }
        }
        return false;
    };

    const existingFormat = {
        bold: hasElement('w:b'),
        italic: hasElement('w:i'),
        underline: hasElement('w:u'),
        strikethrough: hasElement('w:strike')
    };

    const formatsToCheck = ['bold', 'italic', 'underline', 'strikethrough'];
    const changesNeeded = formatsToCheck.some(f => !!desiredFormat[f] !== existingFormat[f]);

    if (!changesNeeded) return;

    const parent = span.runElement.parentNode;
    if (!parent) return;

    const fullText = span.textElement.textContent || '';
    const runStart = span.charStart;

    const localStart = start - runStart;
    const localEnd = end - runStart;

    const beforeText = fullText.substring(0, localStart);
    const affectedText = fullText.substring(localStart, localEnd);
    const afterText = fullText.substring(localEnd);

    if (beforeText.length > 0) {
        const beforeRun = createTextRun(xmlDoc, beforeText, rPr, false);
        parent.insertBefore(beforeRun, span.runElement);
    }

    const newRPr = injectFormattingToRPr(xmlDoc, rPr, desiredFormat, author, generateRedlines);
    const newRun = createTextRunWithRPrElement(xmlDoc, affectedText, newRPr, false);
    parent.insertBefore(newRun, span.runElement);

    if (afterText.length > 0) {
        const afterRun = createTextRun(xmlDoc, afterText, rPr, false);
        parent.insertBefore(afterRun, span.runElement);
    }

    parent.removeChild(span.runElement);
}

/**
 * Processes a run and appends text span metadata.
 *
 * @param {Element} r - Run element
 * @param {Element} p - Paragraph element
 * @param {Element} container - Parent container
 * @param {number} currentOffset - Absolute offset
 * @param {Array} textSpans - Span collection
 * @returns {{ nextOffset: number, text: string }}
 */
function processRunElement(r, p, container, currentOffset, textSpans) {
    const rPr = getFirstElementByTag(r, 'w:rPr');
    let localOffset = currentOffset;
    const textParts = [];

    for (let rc = r.firstChild; rc; rc = rc.nextSibling) {
        if (rc.nodeName === 'w:t') {
            const text = rc.textContent || '';
            if (text.length > 0) {
                textSpans.push({
                    charStart: localOffset,
                    charEnd: localOffset + text.length,
                    textElement: rc,
                    runElement: r,
                    paragraph: p,
                    container,
                    rPr
                });
                localOffset += text.length;
                textParts.push(text);
            }
        } else if (rc.nodeName === 'w:br' || rc.nodeName === 'w:cr') {
            textSpans.push({
                charStart: localOffset,
                charEnd: localOffset + 1,
                textElement: rc,
                runElement: r,
                paragraph: p,
                container,
                rPr
            });
            localOffset += 1;
            textParts.push('\n');
        } else if (rc.nodeName === 'w:tab') {
            textSpans.push({
                charStart: localOffset,
                charEnd: localOffset + 1,
                textElement: rc,
                runElement: r,
                paragraph: p,
                container,
                rPr
            });
            localOffset += 1;
            textParts.push('\t');
        } else if (rc.nodeName === 'w:noBreakHyphen') {
            textSpans.push({
                charStart: localOffset,
                charEnd: localOffset + 1,
                textElement: rc,
                runElement: r,
                paragraph: p,
                container,
                rPr
            });
            localOffset += 1;
            textParts.push('\u2011');
        }
    }
    return { nextOffset: localOffset, text: textParts.join('') };
}

/**
 * Applies deletion over affected spans.
 *
 * @param {Document} xmlDoc - XML document
 * @param {{ spans: Array, starts: number[], ends: number[] }} spanIndex - Indexed spans
 * @param {number} startPos - Delete start
 * @param {number} endPos - Delete end
 * @param {Set<Node>} processedSpans - Processed marker set
 * @param {string} author - Author name
 * @param {boolean} generateRedlines - Track change toggle
 */
function processDelete(xmlDoc, spanIndex, startPos, endPos, processedSpans, author, generateRedlines) {
    forEachOverlappingSpan(spanIndex, startPos, endPos, span => {
        if (processedSpans.has(span.textElement)) return;

        const deleteStart = Math.max(0, startPos - span.charStart);
        const deleteEnd = Math.min(span.charEnd - span.charStart, endPos - span.charStart);

        const originalText = span.textElement.textContent || '';
        const beforeText = originalText.substring(0, deleteStart);
        const deletedText = originalText.substring(deleteStart, deleteEnd);
        const afterText = originalText.substring(deleteEnd);

        if (deletedText.length === 0) return;

        const parent = span.runElement.parentNode;
        if (!parent) return;

        if (beforeText.length === 0 && afterText.length === 0) {
            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            if (generateRedlines) {
                const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
                parent.insertBefore(delWrapper, span.runElement);
            }
            parent.removeChild(span.runElement);
        } else {
            const oldRun = span.runElement;

            if (beforeText.length > 0) {
                const beforeRun = createTextRun(xmlDoc, beforeText, span.rPr, false);
                parent.insertBefore(beforeRun, oldRun);
            }

            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            if (generateRedlines) {
                const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
                parent.insertBefore(delWrapper, oldRun);
            }

            if (afterText.length > 0) {
                const afterRun = createTextRun(xmlDoc, afterText, span.rPr, false);
                parent.insertBefore(afterRun, oldRun);
                span.runElement = afterRun;
                span.textElement = getFirstElementByTag(afterRun, 'w:t') || getFirstElementByTag(afterRun, 't');
            }

            parent.removeChild(oldRun);
        }

        processedSpans.add(span.textElement);
    });
}

/**
 * Inserts new text at an absolute position.
 *
 * @param {Document} xmlDoc - XML document
 * @param {{ spans: Array, starts: number[], ends: number[] }} spanIndex - Indexed spans
 * @param {number} pos - Insertion position
 * @param {string} text - Text to insert
 * @param {Set<Node>} processedSpans - Processed marker set (unused, kept for signature)
 * @param {string} author - Author name
 * @param {Array} [formatHints=[]] - Format hints
 * @param {number} [insertOffset=0] - Offset in modified text
 * @param {boolean} [generateRedlines=true] - Track change toggle
 */
function processInsert(xmlDoc, spanIndex, pos, text, processedSpans, author, formatHints = [], insertOffset = 0, generateRedlines = true) {
    void processedSpans;

    let targetSpan = findContainingSpan(spanIndex, pos);

    if (!targetSpan && pos > 0) {
        targetSpan = findFirstSpanEndingAt(spanIndex, pos);
    }

    if (!targetSpan && pos > 0) {
        targetSpan = findLastSpanEndingBeforeOrAt(spanIndex, pos);
    }

    if (!targetSpan && spanIndex.spans.length > 0) {
        targetSpan = spanIndex.spans[spanIndex.spans.length - 1];
    }

    if (targetSpan) {
        const applicableHints = getApplicableFormatHints(formatHints, insertOffset, insertOffset + text.length);
        const baseRPr = targetSpan.rPr;
        const parent = targetSpan.runElement.parentNode;

        if (parent) {
            const referenceNode = (pos === targetSpan.charStart) ? targetSpan.runElement : targetSpan.runElement.nextSibling;

            if (applicableHints.length === 0) {
                const insRun = createTextRun(xmlDoc, text, baseRPr, false);
                if (generateRedlines) {
                    const insWrapper = createTrackChange(xmlDoc, 'ins', insRun, author);
                    parent.insertBefore(insWrapper, referenceNode);
                } else {
                    parent.insertBefore(insRun, referenceNode);
                }
            } else {
                const runs = createFormattedRuns(xmlDoc, text, baseRPr, applicableHints, insertOffset, author, generateRedlines);

                if (generateRedlines) {
                    const insWrapper = createTrackChange(xmlDoc, 'ins', null, author);
                    runs.forEach(run => insWrapper.appendChild(run));
                    parent.insertBefore(insWrapper, referenceNode);
                } else {
                    runs.forEach(run => parent.insertBefore(run, referenceNode));
                }
            }
        }
    }
}

function buildSpanIndex(textSpans) {
    const spans = textSpans
        .slice()
        .sort((a, b) => a.charStart - b.charStart || a.charEnd - b.charEnd);

    const starts = spans.map(span => span.charStart);
    const ends = spans.map(span => span.charEnd);

    return { spans, starts, ends };
}

function upperBound(values, target) {
    let left = 0;
    let right = values.length;

    while (left < right) {
        const middle = (left + right) >> 1;
        if (values[middle] <= target) {
            left = middle + 1;
        } else {
            right = middle;
        }
    }

    return left;
}

function lowerBound(values, target) {
    let left = 0;
    let right = values.length;

    while (left < right) {
        const middle = (left + right) >> 1;
        if (values[middle] < target) {
            left = middle + 1;
        } else {
            right = middle;
        }
    }

    return left;
}

function forEachOverlappingSpan(spanIndex, startPos, endPos, callback) {
    if (endPos <= startPos || spanIndex.spans.length === 0) {
        return;
    }

    let index = upperBound(spanIndex.ends, startPos);
    while (index < spanIndex.spans.length) {
        const span = spanIndex.spans[index];
        if (span.charStart >= endPos) {
            break;
        }
        callback(span);
        index++;
    }
}

function findContainingSpan(spanIndex, pos) {
    if (spanIndex.spans.length === 0) return null;

    const index = upperBound(spanIndex.starts, pos) - 1;
    if (index < 0) return null;

    const span = spanIndex.spans[index];
    return pos >= span.charStart && pos < span.charEnd ? span : null;
}

function findFirstSpanEndingAt(spanIndex, pos) {
    const index = lowerBound(spanIndex.ends, pos);
    if (index < spanIndex.spans.length && spanIndex.ends[index] === pos) {
        return spanIndex.spans[index];
    }
    return null;
}

function findLastSpanEndingBeforeOrAt(spanIndex, pos) {
    const index = upperBound(spanIndex.ends, pos) - 1;
    if (index >= 0) {
        return spanIndex.spans[index];
    }
    return null;
}
