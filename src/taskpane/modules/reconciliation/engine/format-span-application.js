/**
 * Span-level formatting helpers.
 *
 * Handles boundary splitting and robust format application on already-extracted
 * text spans without owning paragraph targeting concerns.
 */

import { mergeFormats } from '../pipeline/markdown-processor.js';
import { injectFormattingToRPr, createTextRun } from './run-builders.js';

/**
 * Splits spans at all supplied absolute boundaries.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Input spans
 * @param {number[]} boundaries - Absolute boundaries
 * @returns {Array}
 */
export function splitSpansAtBoundaries(xmlDoc, textSpans, boundaries) {
    const sortedBoundaries = Array.from(new Set(boundaries)).sort((a, b) => a - b);
    let currentSpans = [...textSpans];
    let splitsOccurred = true;

    while (splitsOccurred) {
        splitsOccurred = false;
        const nextPassSpans = [];

        for (const span of currentSpans) {
            let splitThisSpan = false;
            for (const boundary of sortedBoundaries) {
                if (boundary > span.charStart && boundary < span.charEnd) {
                    const splitResult = splitSpanAtOffset(xmlDoc, span, boundary);
                    if (splitResult) {
                        nextPassSpans.push(splitResult[0], splitResult[1]);
                        splitsOccurred = true;
                        splitThisSpan = true;
                        break;
                    }
                }
            }
            if (!splitThisSpan) {
                nextPassSpans.push(span);
            }
        }

        currentSpans = nextPassSpans;
    }

    return currentSpans;
}

/**
 * Robust formatting application.
 * Identifies all boundaries, splits all runs first, then applies merged formats.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Text spans
 * @param {Array} formatHints - Format hints
 * @param {string} author - Change author
 * @param {boolean} generateRedlines - Track change toggle
 * @returns {void}
 */
export function applyFormatHintsToSpansRobust(xmlDoc, textSpans, formatHints, author, generateRedlines) {
    if (textSpans.length === 0) return;

    const boundaries = [];
    for (const hint of formatHints) {
        boundaries.push(hint.start, hint.end);
    }

    const currentSpans = splitSpansAtBoundaries(xmlDoc, textSpans, boundaries);

    for (const span of currentSpans) {
        const applicableHints = formatHints.filter(h => h.start < span.charEnd && h.end > span.charStart);
        if (applicableHints.length > 0) {
            const targetFormat = mergeFormats(...applicableHints.map(h => h.format));
            addFormattingToRun(xmlDoc, span.runElement, targetFormat, author, generateRedlines);
        }
    }
}

/**
 * Splits a text span at a specific absolute character offset.
 * Modifies the DOM and returns the two new span objects.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Object} span - Text span
 * @param {number} absoluteOffset - Absolute split offset
 * @returns {Array|null}
 */
export function splitSpanAtOffset(xmlDoc, span, absoluteOffset) {
    const run = span.runElement;
    const parent = run.parentNode;
    if (!parent) return null;

    const fullText = span.textElement.textContent || '';
    const localSplitPoint = absoluteOffset - span.charStart;

    const textBefore = fullText.substring(0, localSplitPoint);
    const textAfter = fullText.substring(localSplitPoint);

    if (textBefore.length === 0 || textAfter.length === 0) return null;

    const runBefore = createTextRun(xmlDoc, textBefore, span.rPr, false);
    const runAfter = createTextRun(xmlDoc, textAfter, span.rPr, false);

    parent.insertBefore(runBefore, run);
    parent.insertBefore(runAfter, run);
    parent.removeChild(run);

    const tBefore = runBefore.getElementsByTagName('w:t')[0];
    const tAfter = runAfter.getElementsByTagName('w:t')[0];

    return [
        { ...span, charEnd: absoluteOffset, textElement: tBefore, runElement: runBefore },
        { ...span, charStart: absoluteOffset, textElement: tAfter, runElement: runAfter }
    ];
}

/**
 * Adds formatting elements to a run's rPr, with track change support.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} run - Run element
 * @param {Object} format - Format flags
 * @param {string} author - Change author
 * @param {boolean} generateRedlines - Track change toggle
 * @returns {void}
 */
function addFormattingToRun(xmlDoc, run, format, author, generateRedlines) {
    let rPr = run.getElementsByTagName('w:rPr')[0];
    const baseRPr = rPr ? rPr.cloneNode(true) : null;

    if (!rPr) {
        rPr = xmlDoc.createElement('w:rPr');
        run.insertBefore(rPr, run.firstChild);
    }

    const newRPr = injectFormattingToRPr(xmlDoc, baseRPr, format, author, generateRedlines);

    while (rPr.firstChild) rPr.removeChild(rPr.firstChild);
    Array.from(newRPr.childNodes).forEach(child => rPr.appendChild(child));
}
