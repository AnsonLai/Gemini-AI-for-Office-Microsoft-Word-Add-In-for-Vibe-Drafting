/**
 * Surgical reconciliation mode.
 *
 * This mode performs in-place run-level edits and preserves existing structure,
 * making it safe for tables and other complex OOXML containers.
 */

import { diff_match_patch } from 'diff-match-patch';
import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import { wordsToChars, charsToWords } from '../pipeline/diff-engine.js';
import { NS_W } from '../core/types.js';
import { createSerializer } from '../adapters/xml-adapter.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { buildOverrideRPrXml } from './rpr-helpers.js';
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
    const delId = Math.floor(Math.random() * 1000000);
    const insId = Math.floor(Math.random() * 1000000);

    const serializer = createSerializer();

    let rPrXml = '';
    const rPr = originalRun.getElementsByTagName('w:rPr')[0];
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

        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === 'w:r') {
                processRunElement(child, p, container, fullText.length, textSpans);
                fullText = getUpdatedFullText(child, fullText);
            } else if (child.nodeName === 'w:hyperlink') {
                Array.from(child.childNodes).forEach(hc => {
                    if (hc.nodeName === 'w:r') {
                        processRunElement(hc, p, container, fullText.length, textSpans);
                        fullText = getUpdatedFullText(hc, fullText);
                    }
                });
            }
        });

        if (pIndex < allParagraphs.length - 1) {
            fullText += '\n';
        }
    });

    const dmp = new diff_match_patch();
    const { chars1, chars2, wordArray } = wordsToChars(fullText, modifiedText);
    const charDiffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(charDiffs);
    const diffs = charsToWords(charDiffs, wordArray);

    let originalPos = 0;
    let newPos = 0;
    const processedSpans = new Set();

    for (const [op, text] of diffs) {
        if (op === 0) {
            const len = text.length;
            const startPos = originalPos;
            const endPos = originalPos + len;

            const affectedSpans = textSpans.filter(s =>
                s.charEnd > startPos && s.charStart < endPos
            );

            for (const span of affectedSpans) {
                const overlapStartOriginal = Math.max(span.charStart, startPos);
                const overlapEndOriginal = Math.min(span.charEnd, endPos);
                const segmentLen = overlapEndOriginal - overlapStartOriginal;
                const relativeOffset = overlapStartOriginal - startPos;
                const overlapStartNew = newPos + relativeOffset;
                const overlapEndNew = overlapStartNew + segmentLen;
                const applicableHints = getApplicableFormatHints(formatHints, overlapStartNew, overlapEndNew);
                reconcileFormattingForTextSpan(xmlDoc, span, overlapStartOriginal, overlapEndOriginal, applicableHints, author, generateRedlines);
            }

            originalPos += len;
            newPos += len;
        } else if (op === -1) {
            processDelete(xmlDoc, textSpans, originalPos, originalPos + text.length, processedSpans, author, generateRedlines);
            originalPos += text.length;
        } else if (op === 1) {
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.trim().length > 0) {
                processInsert(xmlDoc, textSpans, originalPos, textWithoutNewlines, processedSpans, author, formatHints, newPos, generateRedlines);
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
        return rPr && Array.from(rPr.childNodes).some(n => n.nodeName === tagName);
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
 * @returns {number}
 */
function processRunElement(r, p, container, currentOffset, textSpans) {
    const rPr = r.getElementsByTagName('w:rPr')[0] || null;
    let localOffset = currentOffset;

    Array.from(r.childNodes).forEach(rc => {
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
        }
    });
    return localOffset;
}

/**
 * Returns full text with current run text appended.
 *
 * @param {Element} r - Run element
 * @param {string} currentFullText - Current full text
 * @returns {string}
 */
function getUpdatedFullText(r, currentFullText) {
    let fullText = currentFullText;
    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            fullText += rc.textContent || '';
        } else if (rc.nodeName === 'w:br' || rc.nodeName === 'w:cr') {
            fullText += '\n';
        } else if (rc.nodeName === 'w:tab') {
            fullText += '\t';
        } else if (rc.nodeName === 'w:noBreakHyphen') {
            fullText += '\u2011';
        }
    });
    return fullText;
}

/**
 * Applies deletion over affected spans.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Span collection
 * @param {number} startPos - Delete start
 * @param {number} endPos - Delete end
 * @param {Set<Node>} processedSpans - Processed marker set
 * @param {string} author - Author name
 * @param {boolean} generateRedlines - Track change toggle
 */
function processDelete(xmlDoc, textSpans, startPos, endPos, processedSpans, author, generateRedlines) {
    const affectedSpans = textSpans.filter(s =>
        s.charEnd > startPos && s.charStart < endPos
    );

    for (const span of affectedSpans) {
        if (processedSpans.has(span.textElement)) continue;

        const deleteStart = Math.max(0, startPos - span.charStart);
        const deleteEnd = Math.min(span.charEnd - span.charStart, endPos - span.charStart);

        const originalText = span.textElement.textContent || '';
        const beforeText = originalText.substring(0, deleteStart);
        const deletedText = originalText.substring(deleteStart, deleteEnd);
        const afterText = originalText.substring(deleteEnd);

        if (deletedText.length === 0) continue;

        const parent = span.runElement.parentNode;
        if (!parent) continue;

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
                span.textElement = afterRun.getElementsByTagName('w:t')[0] || afterRun.getElementsByTagName('t')[0];
            }

            parent.removeChild(oldRun);
        }

        processedSpans.add(span.textElement);
    }
}

/**
 * Inserts new text at an absolute position.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Span collection
 * @param {number} pos - Insertion position
 * @param {string} text - Text to insert
 * @param {Set<Node>} processedSpans - Processed marker set (unused, kept for signature)
 * @param {string} author - Author name
 * @param {Array} [formatHints=[]] - Format hints
 * @param {number} [insertOffset=0] - Offset in modified text
 * @param {boolean} [generateRedlines=true] - Track change toggle
 */
function processInsert(xmlDoc, textSpans, pos, text, processedSpans, author, formatHints = [], insertOffset = 0, generateRedlines = true) {
    let targetSpan = textSpans.find(s => pos >= s.charStart && pos < s.charEnd);

    if (!targetSpan && pos > 0) {
        targetSpan = textSpans.find(s => pos === s.charEnd);
    }

    if (!targetSpan && pos > 0) {
        const before = textSpans.filter(s => s.charEnd <= pos);
        if (before.length > 0) {
            targetSpan = before[before.length - 1];
        }
    }

    if (!targetSpan && textSpans.length > 0) {
        targetSpan = textSpans[textSpans.length - 1];
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
