/**
 * Formatting application orchestration utilities.
 *
 * Applies format-only changes and surgical formatting synchronization over
 * existing OOXML runs.
 */

import { mergeFormats } from '../pipeline/markdown-processor.js';
import { applyFormatOverridesToRPr, extractFormatFromRPr } from './rpr-helpers.js';
import { snapshotAndAttachRPrChange, injectFormattingToRPr } from './run-builders.js';
import { getDocumentParagraphs, buildTextSpansFromParagraphs } from './format-extraction.js';
import { buildParagraphInfos, findTargetParagraphInfo } from './format-paragraph-targeting.js';
import { splitSpansAtBoundaries, applyFormatHintsToSpansRobust } from './format-span-application.js';
import { getRevisionTimestamp } from '../core/types.js';
import { warn, log } from '../adapters/logger.js';
import { getFirstElementByTag } from '../core/xml-query.js';

function normalizePrecomputedFormatContext(precomputedContext) {
    if (Array.isArray(precomputedContext)) {
        return {
            textSpans: precomputedContext,
            paragraphs: null,
            paragraphInfos: null
        };
    }

    if (!precomputedContext || typeof precomputedContext !== 'object') {
        return {
            textSpans: null,
            paragraphs: null,
            paragraphInfos: null
        };
    }

    return {
        textSpans: Array.isArray(precomputedContext.textSpans) ? precomputedContext.textSpans : null,
        paragraphs: Array.isArray(precomputedContext.paragraphs) ? precomputedContext.paragraphs : null,
        paragraphInfos: Array.isArray(precomputedContext.paragraphInfos) ? precomputedContext.paragraphInfos : null
    };
}

/**
 * Removes existing core formatting via `w:rPrChange` snapshots and explicit overrides.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Text spans (unused, kept for compatibility)
 * @param {Array} existingFormatHints - Existing formatting hints
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Change author
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyFormatRemovalAsSurgicalReplacement(xmlDoc, textSpans, existingFormatHints, serializer, author, generateRedlines = true) {
    void textSpans;
    let hasAnyChanges = false;
    const processedRuns = new Set();
    const dateStr = getRevisionTimestamp();

    log(`[OxmlEngine] Surgical format removal: ${existingFormatHints.length} hints to process (using w:rPrChange)`);

    for (const hint of existingFormatHints) {
        const run = hint.run;
        if (processedRuns.has(run)) continue;
        processedRuns.add(run);
        if (!run.parentNode) continue;

        log('[OxmlEngine] Processing run for surgical format removal, format:', hint.format);

        let rPr = getFirstElementByTag(run, 'w:rPr');
        if (!rPr) {
            rPr = xmlDoc.createElement('w:rPr');
            run.insertBefore(rPr, run.firstChild);
        }

        if (generateRedlines) {
            snapshotAndAttachRPrChange(xmlDoc, rPr, author || 'Gemini AI', dateStr);
        }

        applyFormatOverridesToRPr(xmlDoc, rPr, hint.format);
        hasAnyChanges = true;
    }

    if (hasAnyChanges) {
        log('[OxmlEngine] Surgical format removal completed successfully (Pure Format Mode)');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
    }

    log('[OxmlEngine] No format changes were applied');
    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
}

/**
 * Applies format additions by synchronizing run-level target state after boundary splitting.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array} textSpans - Extracted text spans
 * @param {Array} formatHints - Format hints
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Change author
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyFormatAdditionsAsSurgicalReplacement(xmlDoc, textSpans, formatHints, serializer, author, generateRedlines = true) {
    let hasAnyChanges = false;
    const processedRuns = new Set();

    if (!textSpans || textSpans.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const boundaries = [];
    for (const hint of formatHints) {
        boundaries.push(hint.start, hint.end);
    }
    const currentSpans = splitSpansAtBoundaries(xmlDoc, textSpans, boundaries);
    const orderedSpans = currentSpans
        .slice()
        .sort((a, b) => a.charStart - b.charStart || a.charEnd - b.charEnd);
    const getOverlappingHints = createFormatHintOverlapLookup(formatHints);

    for (const span of orderedSpans) {
        if (!span || !span.textElement || span.textElement.nodeName !== 'w:t') continue;

        const applicableHints = getOverlappingHints(span.charStart, span.charEnd);
        if (applicableHints.length === 0) continue;

        const mergedDesiredFormat = mergeFormats(...applicableHints.map(h => h.format));
        const desiredFormat = {
            bold: !!mergedDesiredFormat.bold,
            italic: !!mergedDesiredFormat.italic,
            underline: !!mergedDesiredFormat.underline,
            strikethrough: !!mergedDesiredFormat.strikethrough
        };

        const existingFormatRaw = span.format || extractFormatFromRPr(span.rPr);
        const existingFormat = {
            bold: !!existingFormatRaw.bold,
            italic: !!existingFormatRaw.italic,
            underline: !!existingFormatRaw.underline,
            strikethrough: !!existingFormatRaw.strikethrough
        };

        const formatsToCheck = ['bold', 'italic', 'underline', 'strikethrough'];
        const needsSync = formatsToCheck.some(f => desiredFormat[f] !== existingFormat[f]);
        if (!needsSync) continue;

        if (processedRuns.has(span.runElement)) continue;
        processedRuns.add(span.runElement);

        const textContent = span.textElement.textContent || '';
        if (!textContent) continue;

        const parentNode = span.runElement.parentNode;
        if (!parentNode) continue;

        const run = span.runElement;
        const baseRPr = getFirstElementByTag(run, 'w:rPr');
        const syncedRPr = injectFormattingToRPr(
            xmlDoc,
            baseRPr,
            desiredFormat,
            author || 'Gemini AI',
            generateRedlines
        );

        let rPr = baseRPr;
        if (!rPr) {
            run.insertBefore(syncedRPr, run.firstChild);
        } else {
            while (rPr.firstChild) {
                rPr.removeChild(rPr.firstChild);
            }
            Array.from(syncedRPr.childNodes).forEach(child => rPr.appendChild(child));
        }

        hasAnyChanges = true;
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: hasAnyChanges };
}

/**
 * Applies formatting changes to existing text without modifying content.
 * Used when markdown formatting is applied to unchanged text.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original plain text
 * @param {Array} formatHints - Format hints
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Change author
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @param {Array|Object|null} [precomputedContext=null] - Optional precomputed format context
 * @returns {{ oxml?: string, hasChanges: boolean, useNativeApi?: boolean, formatHints?: Array, originalText?: string }}
 */
export function applyFormatOnlyChanges(xmlDoc, originalText, formatHints, serializer, author, generateRedlines = true, precomputedContext = null) {
    const precomputed = normalizePrecomputedFormatContext(precomputedContext);
    const allParagraphs = precomputed.paragraphs || getDocumentParagraphs(xmlDoc);

    let textSpans = precomputed.textSpans || [];
    if (!precomputed.textSpans) {
        ({ textSpans } = buildTextSpansFromParagraphs(allParagraphs));
    }

    if (!textSpans || textSpans.length === 0) {
        warn('[OxmlEngine] No spans available for format-only change; deferring to native API');
        return {
            hasChanges: true,
            useNativeApi: true,
            formatHints,
            originalText
        };
    }

    if (!formatHints || formatHints.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const paragraphInfos = precomputed.paragraphInfos || buildParagraphInfos(xmlDoc, allParagraphs, textSpans);
    const { targetInfo, matchOffset } = findTargetParagraphInfo(paragraphInfos, originalText);

    if (!targetInfo || !targetInfo.spans || targetInfo.spans.length === 0) {
        warn('[OxmlEngine] Unable to pinpoint target paragraph for format-only change; deferring to native API');
        return {
            hasChanges: true,
            useNativeApi: true,
            formatHints,
            originalText
        };
    }

    const baseOffset = targetInfo.spans[0].charStart;
    const localizedSpans = targetInfo.spans.map(span => ({
        ...span,
        charStart: span.charStart - baseOffset,
        charEnd: span.charEnd - baseOffset
    }));

    const adjustedHints = formatHints.map(hint => ({
        ...hint,
        start: hint.start + matchOffset,
        end: hint.end + matchOffset
    }));

    applyFormatHintsToSpansRobust(xmlDoc, localizedSpans, adjustedHints, author, generateRedlines);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Surgical variant of format-only changes.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original plain text
 * @param {Array} formatHints - Format hints
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Change author
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @param {Array|Object|null} [precomputedContext=null] - Optional precomputed format context
 * @returns {{ oxml?: string, hasChanges: boolean, useNativeApi?: boolean, formatHints?: Array, originalText?: string }}
 */
export function applyFormatOnlyChangesSurgical(xmlDoc, originalText, formatHints, serializer, author, generateRedlines = true, precomputedContext = null) {
    const precomputed = normalizePrecomputedFormatContext(precomputedContext);
    const allParagraphs = precomputed.paragraphs || getDocumentParagraphs(xmlDoc);

    let textSpans = precomputed.textSpans || [];
    if (!precomputed.textSpans) {
        ({ textSpans } = buildTextSpansFromParagraphs(allParagraphs));
    }

    if (!textSpans || textSpans.length === 0) {
        warn('[OxmlEngine] No spans available for surgical format-only change; deferring to native API');
        return {
            hasChanges: true,
            useNativeApi: true,
            formatHints,
            originalText
        };
    }

    if (!formatHints || formatHints.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const paragraphInfos = precomputed.paragraphInfos || buildParagraphInfos(xmlDoc, allParagraphs, textSpans);
    const { targetInfo, matchOffset } = findTargetParagraphInfo(paragraphInfos, originalText);

    if (!targetInfo || !targetInfo.spans || targetInfo.spans.length === 0) {
        warn('[OxmlEngine] Unable to pinpoint target paragraph for surgical format-only change; deferring to native API');
        return {
            hasChanges: true,
            useNativeApi: true,
            formatHints,
            originalText
        };
    }

    const baseOffset = targetInfo.spans[0].charStart;
    const localizedSpans = targetInfo.spans.map(span => ({
        ...span,
        charStart: span.charStart - baseOffset,
        charEnd: span.charEnd - baseOffset
    }));

    const adjustedHints = formatHints.map(hint => ({
        ...hint,
        start: hint.start + matchOffset,
        end: hint.end + matchOffset
    }));

    return applyFormatAdditionsAsSurgicalReplacement(
        xmlDoc,
        localizedSpans,
        adjustedHints,
        serializer,
        author,
        generateRedlines
    );
}

/**
 * Builds a sweep-line overlap lookup for format hints.
 *
 * @param {Array} formatHints - Format hints
 * @returns {(start:number, end:number)=>Array}
 */
function createFormatHintOverlapLookup(formatHints) {
    const sortedHints = (formatHints || [])
        .slice()
        .sort((a, b) => a.start - b.start || a.end - b.end);

    const activeHints = [];
    let nextHintIndex = 0;

    return (start, end) => {
        while (nextHintIndex < sortedHints.length && sortedHints[nextHintIndex].start < end) {
            activeHints.push(sortedHints[nextHintIndex]);
            nextHintIndex++;
        }

        for (let i = activeHints.length - 1; i >= 0; i--) {
            if (activeHints[i].end <= start) {
                activeHints.splice(i, 1);
            }
        }

        return activeHints.filter(hint => hint.start < end && hint.end > start);
    };
}
