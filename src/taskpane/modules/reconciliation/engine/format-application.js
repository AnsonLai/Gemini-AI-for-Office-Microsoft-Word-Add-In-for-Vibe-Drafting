/**
 * Formatting application utilities.
 *
 * Applies format-only changes and surgical formatting edits over existing OOXML runs.
 */

import { mergeFormats } from '../pipeline/markdown-processor.js';
import { applyFormatOverridesToRPr, extractFormatFromRPr } from './rpr-helpers.js';
import { snapshotAndAttachRPrChange, injectFormattingToRPr, createTextRun, createTextRunWithRPrElement } from './run-builders.js';
import { getDocumentParagraphs, buildTextSpansFromParagraphs } from './format-extraction.js';
import { log, warn } from '../adapters/logger.js';
export function applyFormatRemovalAsSurgicalReplacement(xmlDoc, textSpans, existingFormatHints, serializer, author, generateRedlines = true) {
    let hasAnyChanges = false;
    const processedRuns = new Set();
    const dateStr = new Date().toISOString();

    log(`[OxmlEngine] Surgical format removal: ${existingFormatHints.length} hints to process (using w:rPrChange)`);

    for (const hint of existingFormatHints) {
        const run = hint.run;

        // Skip if already processed this run
        if (processedRuns.has(run)) continue;
        processedRuns.add(run);

        // Skip if run has no parent (already removed)
        if (!run.parentNode) continue;

        log(`[OxmlEngine] Processing run for surgical format removal, format:`, hint.format);

        // Get or create w:rPr
        let rPr = run.getElementsByTagName('w:rPr')[0];
        if (!rPr) {
            rPr = xmlDoc.createElement('w:rPr');
            run.insertBefore(rPr, run.firstChild);
        }

        // If generating redlines, track the change
        if (generateRedlines) {
            snapshotAndAttachRPrChange(xmlDoc, rPr, author || 'Gemini AI', dateStr);
        }

        // Apply format overrides (unbold, unitalic, etc.)
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
 * Format ADDITION via surgical text replacement (pure OOXML approach).
 * UPDATED: Uses w:rPr modification with w:rPrChange instead of text replacement (w:del/w:ins).
 * This prevents the "strikeout + new text" visualization and shows a proper formatting change.
 */
export function applyFormatAdditionsAsSurgicalReplacement(xmlDoc, textSpans, formatHints, serializer, author, generateRedlines = true) {
    let hasAnyChanges = false;
    const processedRuns = new Set();

    if (!textSpans || textSpans.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const boundaries = new Set();
    for (const hint of formatHints) {
        boundaries.add(hint.start);
        boundaries.add(hint.end);
    }
    const sortedBoundaries = Array.from(boundaries).sort((a, b) => a - b);

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

    for (const span of currentSpans) {
        if (!span || !span.textElement || span.textElement.nodeName !== 'w:t') continue;

        const applicableHints = formatHints.filter(h => h.start < span.charEnd && h.end > span.charStart);
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
        const baseRPr = run.getElementsByTagName('w:rPr')[0] || null;
        const syncedRPr = injectFormattingToRPr(
            xmlDoc,
            baseRPr,
            desiredFormat,
            author || 'Gemini AI',
            generateRedlines
        );

        // Replace or insert run properties with fully synchronized target format.
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
 * Removes formatting from specified spans using the pre-extracted data.
 * Handles both direct formatting (w:b tags) and inherited formatting (from paragraph).
 * For inherited formatting, adds explicit override elements (w:b w:val="0").
 */
export function applyFormatRemovalWithSpans(xmlDoc, textSpans, existingFormatHints, serializer, author, generateRedlines = true) {
    let hasAnyChanges = false;
    const processedRuns = new Set();
    const processedParagraphs = new Set();

    log(`[OxmlEngine] applyFormatRemovalWithSpans: ${existingFormatHints.length} hints to process`);

    // 1. Check and strip paragraph-level formatting first
    for (const span of textSpans) {
        const paragraph = span.paragraph;
        if (processedParagraphs.has(paragraph)) continue;
        processedParagraphs.add(paragraph);

        // Find pPr/rPr
        let pPr = null;
        let pRPr = null;
        for (const child of Array.from(paragraph.childNodes)) {
            if (child.nodeName === 'w:pPr') {
                pPr = child;
                for (const pChild of Array.from(child.childNodes)) {
                    if (pChild.nodeName === 'w:rPr') {
                        pRPr = pChild;
                        break;
                    }
                }
                break;
            }
        }

        if (pRPr) {
            const pToRemove = [];
            for (const child of Array.from(pRPr.childNodes)) {
                if (['w:b', 'w:i', 'w:u', 'w:strike'].includes(child.nodeName)) {
                    pToRemove.push(child);
                }
            }

            if (pToRemove.length > 0) {
                hasAnyChanges = true;
                log(`[OxmlEngine] Removing paragraph-level formatting: ${pToRemove.map(e => e.nodeName).join(', ')}`);
                for (const el of pToRemove) {
                    pRPr.removeChild(el);
                }
            }
        }
    }

    // 2. Process each format hint - handle both direct and inherited formatting
    for (const hint of existingFormatHints) {
        const run = hint.run;
        let rPr = hint.rPr;

        // Skip if already processed this run
        if (processedRuns.has(run)) continue;
        processedRuns.add(run);

        log(`[OxmlEngine] Processing format hint: bold=${hint.format.bold}, italic=${hint.format.italic}, rPr=${rPr ? 'exists' : 'null'}`);

        // Case 1: Run has rPr - check for direct formatting OR style-based formatting
        if (rPr && rPr.parentNode) {
            const toRemove = [];
            let hasStyleRef = false;

            for (const child of Array.from(rPr.childNodes)) {
                if (['w:b', 'w:i', 'w:u', 'w:strike'].includes(child.nodeName)) {
                    toRemove.push(child);
                }
                if (child.nodeName === 'w:rStyle') {
                    hasStyleRef = true;
                }
            }

            // Case 1a: Direct formatting tags exist - remove them
            if (toRemove.length > 0) {
                hasAnyChanges = true;
                log(`[OxmlEngine] Removing direct formatting from run: ${toRemove.map(e => e.nodeName).join(', ')}`);

                // Create rPrChange for track changes
                if (generateRedlines) {
                    snapshotAndAttachRPrChange(xmlDoc, rPr, author || 'Gemini AI', new Date().toISOString());
                }

                for (const el of toRemove) {
                    rPr.removeChild(el);
                }
            }
            // Case 1b: No direct tags but formatting detected (from style or paragraph) - add overrides
            else if (hint.format.hasFormatting) {
                hasAnyChanges = true;
                log(`[OxmlEngine] Adding format overrides for style-based/inherited formatting (rPr exists)`);

                // Create rPrChange before modifying
                if (generateRedlines) {
                    snapshotAndAttachRPrChange(xmlDoc, rPr, author || 'Gemini AI', new Date().toISOString());
                }

                // Add explicit override elements to turn OFF the formatting
                if (hint.format.bold) {
                    const b = xmlDoc.createElement('w:b');
                    b.setAttribute('w:val', '0');
                    rPr.insertBefore(b, rPr.firstChild);
                }
                if (hint.format.italic) {
                    const i = xmlDoc.createElement('w:i');
                    i.setAttribute('w:val', '0');
                    rPr.insertBefore(i, rPr.firstChild);
                }
                if (hint.format.underline) {
                    const u = xmlDoc.createElement('w:u');
                    u.setAttribute('w:val', 'none');
                    rPr.insertBefore(u, rPr.firstChild);
                }
                if (hint.format.strikethrough) {
                    const strike = xmlDoc.createElement('w:strike');
                    strike.setAttribute('w:val', '0');
                    rPr.insertBefore(strike, rPr.firstChild);
                }
            }
        }
        // Case 2: Run has no rPr but inherits formatting - add override elements
        else if (!rPr && hint.format.hasFormatting) {
            hasAnyChanges = true;
            log(`[OxmlEngine] Adding format overrides for inherited formatting`);

            // Create rPr for this run
            rPr = xmlDoc.createElement('w:rPr');

            // Add override elements to turn OFF inherited formatting
            if (hint.format.bold) {
                const b = xmlDoc.createElement('w:b');
                b.setAttribute('w:val', '0');
                rPr.appendChild(b);
            }
            if (hint.format.italic) {
                const i = xmlDoc.createElement('w:i');
                i.setAttribute('w:val', '0');
                rPr.appendChild(i);
            }
            if (hint.format.underline) {
                const u = xmlDoc.createElement('w:u');
                u.setAttribute('w:val', 'none');
                rPr.appendChild(u);
            }
            if (hint.format.strikethrough) {
                const strike = xmlDoc.createElement('w:strike');
                strike.setAttribute('w:val', '0');
                rPr.appendChild(strike);
            }

            // Create rPrChange for track changes (previous state was empty)
            if (generateRedlines) {
                const emptyOriginalRPr = xmlDoc.createElement('w:rPr');
                snapshotAndAttachRPrChange(xmlDoc, rPr, author || 'Gemini AI', new Date().toISOString(), emptyOriginalRPr);
            }

            // Insert rPr as first child of run
            run.insertBefore(rPr, run.firstChild);
        }
    }

    if (hasAnyChanges) {
        log('[OxmlEngine] Format removal applied successfully');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
    }

    log('[OxmlEngine] No formatting elements were removed');
    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
}

// ============================================================================
// FORMAT-ONLY MODE (applies formatting without text changes)
// ============================================================================

/**
 * Applies formatting to a single paragraph only.
 * Used for table cell edits where format hints are relative to the target paragraph's text only.
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {Element} paragraph - The target paragraph element
 * @param {Array} formatHints - Format hints with start/end offsets
 * @param {string} author - Author for track changes
 * @param {boolean} generateRedlines - Whether to generate track changes
 */
export function applyFormatToSingleParagraph(xmlDoc, paragraph, formatHints, author, generateRedlines) {
    const { textSpans, charOffset } = buildTextSpansFromParagraphs([paragraph]);

    log(`[OxmlEngine] Single paragraph has ${textSpans.length} text spans, total chars: ${charOffset}`);

    // Apply each format hint to the corresponding text spans robustly
    applyFormatHintsToSpansRobust(xmlDoc, textSpans, formatHints, author, generateRedlines);
}

/**
 * Applies formatting changes to existing text without modifying content.
 * Used when markdown formatting is applied to unchanged text.
 */
export function applyFormatOnlyChanges(xmlDoc, originalText, formatHints, serializer, author, generateRedlines = true, precomputedSpans = null) {
    const allParagraphs = getDocumentParagraphs(xmlDoc);

    // Reuse precomputed spans when available to keep character offsets aligned
    let textSpans = Array.isArray(precomputedSpans) ? precomputedSpans : [];

    if (!Array.isArray(precomputedSpans)) {
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

    // Group spans by paragraph element for precise offset alignment
    const paragraphInfos = buildParagraphInfos(xmlDoc, allParagraphs, textSpans);
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

    // Adjust spans so char offsets are relative to the target paragraph
    const baseOffset = targetInfo.spans[0].charStart;
    const localizedSpans = targetInfo.spans.map(span => ({
        ...span,
        charStart: span.charStart - baseOffset,
        charEnd: span.charEnd - baseOffset
    }));

    // Shift format hints if the match occurs after the paragraph start
    const adjustedHints = formatHints.map(hint => ({
        ...hint,
        start: hint.start + matchOffset,
        end: hint.end + matchOffset
    }));

    applyFormatHintsToSpansRobust(xmlDoc, localizedSpans, adjustedHints, author, generateRedlines);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

export function applyFormatOnlyChangesSurgical(xmlDoc, originalText, formatHints, serializer, author, generateRedlines = true, precomputedSpans = null) {
    const allParagraphs = getDocumentParagraphs(xmlDoc);

    let textSpans = Array.isArray(precomputedSpans) ? precomputedSpans : [];

    if (!Array.isArray(precomputedSpans)) {
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

    const paragraphInfos = buildParagraphInfos(xmlDoc, allParagraphs, textSpans);
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
 * Builds paragraph metadata (text, spans, offsets) used for format-only changes.
 */
export function buildParagraphInfos(xmlDoc, paragraphs, textSpans) {
    const spansByParagraph = new Map();
    for (const span of textSpans) {
        if (!span || !span.paragraph) continue;
        if (!spansByParagraph.has(span.paragraph)) {
            spansByParagraph.set(span.paragraph, []);
        }
        spansByParagraph.get(span.paragraph).push(span);
    }

    const infos = [];
    let runningOffset = 0;

    paragraphs.forEach((p, index) => {
        const spans = (spansByParagraph.get(p) || []).slice().sort((a, b) => a.charStart - b.charStart);
        const text = buildParagraphTextFromSpans(spans);

        infos.push({
            paragraph: p,
            spans,
            text,
            startOffset: runningOffset
        });

        runningOffset += normalizeParagraphComparisonText(text).length;
        if (index < paragraphs.length - 1) {
            runningOffset += 1; // Account for implicit newline between paragraphs
        }
    });

    return infos;
}

/**
 * Reconstructs human-readable text from spans to align with Word paragraph text.
 */
export function buildParagraphTextFromSpans(spans) {
    if (!spans || spans.length === 0) return '';

    let text = '';
    for (const span of spans) {
        if (!span || !span.textElement) continue;
        const nodeName = span.textElement.nodeName;
        if (nodeName === 'w:t') {
            text += span.textElement.textContent || '';
        } else if (nodeName === 'w:tab') {
            text += '\t';
        } else if (nodeName === 'w:br' || nodeName === 'w:cr') {
            text += '\n';
        } else if (nodeName === 'w:noBreakHyphen') {
            text += '\u2011';
        }
    }
    return text;
}

/**
 * Finds a paragraph info entry that matches the provided text.
 */
export function findMatchingParagraphInfo(paragraphInfos, originalText) {
    if (!originalText) return null;

    const normalizedOriginal = normalizeParagraphComparisonText(originalText);
    const normalizedTrim = normalizedOriginal.trim();
    if (!normalizedTrim) return null;

    for (const info of paragraphInfos) {
        const candidate = normalizeParagraphComparisonText(info.text);
        if (candidate === normalizedOriginal) {
            return info;
        }
    }

    for (const info of paragraphInfos) {
        const candidate = normalizeParagraphComparisonText(info.text);
        if (candidate.trim() === normalizedTrim) {
            return info;
        }
    }

    return null;
}

/**
 * Finds the best target paragraph and offset for format-only operations.
 *
 * @param {Array} paragraphInfos - Paragraph metadata
 * @param {string} originalText - Original text used for matching
 * @returns {{ targetInfo: Object|null, matchOffset: number }}
 */
export function findTargetParagraphInfo(paragraphInfos, originalText) {
    const normalizedOriginalFull = normalizeParagraphComparisonText(originalText);
    const normalizedOriginalTrim = normalizedOriginalFull.trim();

    let targetInfo = null;
    let matchOffset = 0;

    for (const info of paragraphInfos) {
        const candidateFull = normalizeParagraphComparisonText(info.text);
        if (candidateFull === normalizedOriginalFull) {
            targetInfo = info;
            return { targetInfo, matchOffset };
        }
    }

    if (normalizedOriginalTrim.length > 0) {
        for (const info of paragraphInfos) {
            const candidateFull = normalizeParagraphComparisonText(info.text);
            if (candidateFull.trim() === normalizedOriginalTrim) {
                targetInfo = info;
                return { targetInfo, matchOffset };
            }
        }
    }

    if (normalizedOriginalTrim.length > 0) {
        const docPlain = paragraphInfos.map(info => normalizeParagraphComparisonText(info.text)).join('\n');
        const idx = docPlain.indexOf(normalizedOriginalTrim);
        if (idx !== -1) {
            for (const info of paragraphInfos) {
                const start = info.startOffset;
                const length = normalizeParagraphComparisonText(info.text).length;
                if (idx >= start && idx <= start + length) {
                    targetInfo = info;
                    matchOffset = idx - start;
                    break;
                }
            }
        }
    }

    return { targetInfo, matchOffset };
}

/**
 * Walks up the DOM to find the containing paragraph for a run.
 */
export function getContainingParagraph(node) {
    let current = node;
    while (current) {
        if (current.nodeName === 'w:p') return current;
        current = current.parentNode;
    }
    return null;
}

/**
 * Normalizes paragraph text for comparisons (handles carriage returns and NBSP).
 */
export function normalizeParagraphComparisonText(text) {
    if (!text) return '';
    return text
        .replace(/\r/g, '\n')
        .replace(/\u00a0/g, ' ');
}

/**
 * Robust version of formatting application.
 * Identifies all boundaries, splits ALl runs FIRST, then applies merged formats.
 */
export function applyFormatHintsToSpansRobust(xmlDoc, textSpans, formatHints, author, generateRedlines) {
    if (textSpans.length === 0) return;

    // 1. Identify all boundaries
    const boundaries = new Set();
    for (const hint of formatHints) {
        boundaries.add(hint.start);
        boundaries.add(hint.end);
    }
    const sortedBoundaries = Array.from(boundaries).sort((a, b) => a - b);

    // 2. Split-First: ensure all runs are broken at all boundaries
    let currentSpans = [...textSpans];
    let splitsOccurred = true;
    while (splitsOccurred) {
        splitsOccurred = false;
        let nextPassSpans = [];
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

    // 3. Apply formatting ONLY to spans that have applicable hints
    // Spans without hints are LEFT UNTOUCHED to preserve existing formatting
    for (const span of currentSpans) {
        const applicableHints = formatHints.filter(h => h.start < span.charEnd && h.end > span.charStart);

        // Only apply formatting if there are hints for this span
        // This preserves existing formatting on spans not targeted by AI
        if (applicableHints.length > 0) {
            const targetFormat = mergeFormats(...applicableHints.map(h => h.format));
            addFormattingToRun(xmlDoc, span.runElement, targetFormat, author, generateRedlines);
        }
        // If no hints apply, leave the span unchanged (preserve original formatting)
    }
}


/**
 * Splits a text span at a specific absolute character offset.
 * Modifies the DOM and returns the two new span objects.
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

    // Create new runs
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
 * Applies a single format hint to affected text spans.
 * Splits runs when only partial formatting is needed.
 */
// Deprecated: use applyFormatHintsToSpansRobust instead
export function applyFormatHintToSpans(xmlDoc, textSpans, hint, author, generateRedlines) {
    applyFormatHintsToSpansRobust(xmlDoc, textSpans, [hint], author, generateRedlines);
}

/**
 * Creates a text run with formatting applied directly.
 */
export function createFormattedRunWithElement(xmlDoc, text, baseRPr, format, author, generateRedlines) {
    const run = xmlDoc.createElement('w:r');

    // Create rPr with formatting (and track changes if author provided AND enabled)
    const rPr = injectFormattingToRPr(xmlDoc, baseRPr, format, author, generateRedlines);

    run.appendChild(rPr);

    // Add text element
    const textEl = xmlDoc.createElement('w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Adds formatting elements to a run's rPr, with track change support.
 */
export function addFormattingToRun(xmlDoc, run, format, author, generateRedlines) {
    let rPr = run.getElementsByTagName('w:rPr')[0];
    const baseRPr = rPr ? rPr.cloneNode(true) : null;

    // Create rPr if it doesn't exist
    if (!rPr) {
        rPr = xmlDoc.createElement('w:rPr');
        run.insertBefore(rPr, run.firstChild);
    }

    // Synchronize formatting using robust helper
    const newRPr = injectFormattingToRPr(xmlDoc, baseRPr, format, author, generateRedlines);

    // Replace old rPr children with new ones
    while (rPr.firstChild) rPr.removeChild(rPr.firstChild);
    Array.from(newRPr.childNodes).forEach(child => rPr.appendChild(child));
}
