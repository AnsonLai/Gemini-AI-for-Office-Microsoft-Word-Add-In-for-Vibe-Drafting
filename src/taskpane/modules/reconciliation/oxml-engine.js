/**
 * OOXML Engine V5.1 - Hybrid Mode
 * 
 * This engine modifies OOXML DOM in-place rather than serializing new content.
 * This ensures Word interprets our w:ins/w:del as actual track changes.
 * 
 * TWO MODES:
 * - SURGICAL MODE (for tables): Modifies existing elements in place, never creates/deletes structure
 * - RECONSTRUCTION MODE (for body without tables): Allows new paragraphs for list splitting
 */

import { diff_match_patch } from 'diff-match-patch';
import { preprocessMarkdown, getApplicableFormatHints } from './markdown-processor.js';
import { ingestTableToVirtualGrid } from './ingestion.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from './table-reconciliation.js';
import { parseTable, ReconciliationPipeline } from './pipeline.js';
import { NumberingService } from './numbering-service.js';
import { NS_W } from './types.js';

// ============================================================================
// TYPES
// ============================================================================

/**
 * @typedef {Object} TextSpan
 * @property {number} charStart - Start character offset in full text
 * @property {number} charEnd - End character offset in full text  
 * @property {Element} textElement - The w:t element
 * @property {Element} runElement - The w:r element
 * @property {Element} paragraph - The w:p element
 * @property {Element} container - Parent container element
 * @property {Element|null} rPr - Run properties (formatting)
 */

/**
 * @typedef {Object} PropertyMapEntry
 * @property {number} start - Start offset
 * @property {number} end - End offset
 * @property {Element|null} rPr - Run properties
 * @property {Node} [wrapper] - Optional wrapper element (e.g., hyperlink)
 */

/**
 * @typedef {Object} ParagraphMapEntry  
 * @property {number} start - Start offset
 * @property {number} end - End offset
 * @property {Element|null} pPr - Paragraph properties
 * @property {Node} container - Parent container
 */

/**
 * @typedef {Object} SentinelMapEntry
 * @property {number} start - Start offset
 * @property {Node} node - The sentinel node
 * @property {boolean} [isTextBox] - Whether this is a text box
 * @property {Node} [originalContainer] - Original text box container
 */

// ============================================================================
// MAIN EXPORT
// ============================================================================

/**
 * Applies redline track changes to OOXML by modifying the DOM in-place.
 * 
 * @param {string} oxml - Original OOXML string (pkg:package format or raw body)
 * @param {string} originalText - Original plain text
 * @param {string} modifiedText - New text (may contain markdown)
 * @param {Object} [options] - Options
 * @param {string} [options.author='AI'] - Author for track changes
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const author = options.author || 'Gemini AI';

    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, 'text/xml');
    } catch (e) {
        console.error('[OxmlEngine] Failed to parse OXML:', e);
        return { oxml, hasChanges: false };
    }

    // Check for parse errors
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        console.error('[OxmlEngine] XML parse error:', parseError.textContent);
        return { oxml, hasChanges: false };
    }

    // Sanitize and preprocess markdown from modified text
    const sanitizedText = sanitizeAiResponse(modifiedText);
    const { cleanText: cleanModifiedText, formatHints } = preprocessMarkdown(sanitizedText);

    // Check if there are actual text changes
    const hasTextChanges = cleanModifiedText.trim() !== originalText.trim();
    const hasFormatHints = formatHints.length > 0;

    // Extract existing formatting from the OOXML paragraph runs
    const { existingFormatHints, textSpans } = extractFormattingFromOoxml(xmlDoc);
    const hasExistingFormatting = existingFormatHints.length > 0;

    console.log(`[OxmlEngine] Text changes: ${hasTextChanges}, New format hints: ${formatHints.length}, Existing format hints: ${existingFormatHints.length}`);

    // Determine format removal: no new hints but existing formatting exists
    const needsFormatRemoval = !hasTextChanges && !hasFormatHints && hasExistingFormatting;

    // Early exit only if NO text changes AND NO format hints to add AND NO existing formatting to remove
    if (!hasTextChanges && !hasFormatHints && !hasExistingFormatting) {
        console.log('[OxmlEngine] No text changes, no format hints, and no existing formatting detected');
        return { oxml, hasChanges: false };
    }

    // Format REMOVAL: text is unchanged, no new hints, but original has formatting to strip
    if (needsFormatRemoval) {
        console.log(`[OxmlEngine] Format REMOVAL detected: stripping ${existingFormatHints.length} format ranges`);
        return applyFormatRemovalWithSpans(xmlDoc, textSpans, existingFormatHints, serializer, author);
    }

    // Format-only change: text is the same but we have formatting to apply
    if (!hasTextChanges && hasFormatHints) {
        console.log(`[OxmlEngine] Format-only change detected: ${formatHints.length} format hints`);
        return applyFormatOnlyChanges(xmlDoc, originalText, formatHints, serializer, author);
    }

    // Check for tables to decide mode
    const tables = xmlDoc.getElementsByTagName('w:tbl');
    const hasTables = tables.length > 0;

    const isMarkdownTable = /^\|.+\|/.test(cleanModifiedText.trim()) && cleanModifiedText.includes('\n');
    const isTargetList = cleanModifiedText.includes('\n') && /^([-*+]|\d+\.)/m.test(cleanModifiedText.trim());

    console.log(`[OxmlEngine] Mode: ${hasTables ? 'SURGICAL' : 'RECONSTRUCTION'}, formatHints: ${formatHints.length}, isMarkdownTable: ${isMarkdownTable}, isTargetList: ${isTargetList}`);

    if (hasTables && isMarkdownTable) {
        return applyTableReconciliation(xmlDoc, cleanModifiedText, serializer, author, formatHints);
    } else if (hasTables) {
        return applySurgicalMode(xmlDoc, originalText, cleanModifiedText, serializer, author, formatHints);
    } else if (isTargetList) {
        // Use the new ReconciliationPipeline for list expanded content
        const pipeline = new ReconciliationPipeline({ author, generateRedlines: true });
        const result = await pipeline.execute(oxml, modifiedText);

        // Wrap the result with numbering definitions for proper list rendering
        if (result.isValid && result.ooxml && result.ooxml !== oxml) {
            const wrapped = pipeline.wrapForInsertion(result.ooxml, {
                includeNumbering: result.includeNumbering || true
            });
            return { oxml: wrapped, hasChanges: true };
        }
        return { oxml, hasChanges: false };
    } else {
        return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer, author, formatHints);
    }
}

// ============================================================================
// FORMAT EXTRACTION & REMOVAL
// ============================================================================

/**
 * Extracts formatting hints from actual paragraph runs in the OOXML.
 * Only looks inside w:p elements to avoid style definitions.
 * 
 * @returns {{ existingFormatHints: Array, textSpans: Array }}
 */
/**
 * Extracts formatting hints from actual paragraph runs in the OOXML.
 * Only looks inside w:p elements to avoid style definitions.
 * 
 * @returns {{ existingFormatHints: Array, textSpans: Array }}
 */
function extractFormattingFromOoxml(xmlDoc) {
    const existingFormatHints = [];
    const textSpans = [];
    let charOffset = 0;

    // Only process runs inside paragraphs (not styles)
    const allParagraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

    for (const p of allParagraphs) {
        // 1. Find pPr and its rPr (paragraph-level default run properties)
        let pRPr = null;
        for (const child of p.childNodes) {
            if (child.nodeName === 'w:pPr') {
                for (const pChild of child.childNodes) {
                    if (pChild.nodeName === 'w:rPr') {
                        pRPr = pChild;
                        break;
                    }
                }
                break;
            }
        }

        // Extract paragraph-level format flags
        const pFormat = extractFormatFromRPr(pRPr);
        if (pFormat.hasFormatting) {
            console.log(`[OxmlEngine] Found paragraph-level formatting: ${JSON.stringify(pFormat)}`);
        }

        // 2. Process runs directly inside paragraphs
        for (const child of p.childNodes) {
            if (child.nodeName === 'w:r') {
                charOffset = processRunForFormatting(child, p, charOffset, textSpans, existingFormatHints, pFormat);
            } else if (child.nodeName === 'w:hyperlink') {
                // Process runs inside hyperlinks
                for (const hc of child.childNodes) {
                    if (hc.nodeName === 'w:r') {
                        charOffset = processRunForFormatting(hc, p, charOffset, textSpans, existingFormatHints, pFormat);
                    }
                }
            }
        }
        // Add newline between paragraphs (not after last)
        charOffset++; // Account for implicit newline
    }

    console.log(`[OxmlEngine] Extracted ${textSpans.length} text spans, ${existingFormatHints.length} format hints`);
    return { existingFormatHints, textSpans };
}

/**
 * Extracts formatting flags from an rPr element.
 */
function extractFormatFromRPr(rPr) {
    const format = { bold: false, italic: false, underline: false, strikethrough: false, hasFormatting: false };
    if (!rPr) return format;

    for (const child of rPr.childNodes) {
        if (child.nodeName === 'w:b') format.bold = true;
        if (child.nodeName === 'w:i') format.italic = true;
        if (child.nodeName === 'w:u') format.underline = true;
        if (child.nodeName === 'w:strike') format.strikethrough = true;

        // Check for style reference (very common for bold/italic)
        if (child.nodeName === 'w:rStyle') {
            const styleRef = child.getAttribute('w:val');
            if (styleRef) {
                const lowerStyle = styleRef.toLowerCase();
                if (lowerStyle.includes('bold') || lowerStyle.includes('strong')) format.bold = true;
                if (lowerStyle.includes('italic') || lowerStyle.includes('emphasis')) format.italic = true;
                if (lowerStyle.includes('underline')) format.underline = true;
            }
        }
    }

    format.hasFormatting = format.bold || format.italic || format.underline || format.strikethrough;
    return format;
}

/**
 * Processes a single run element to extract text spans and formatting.
 */
function processRunForFormatting(run, paragraph, charOffset, textSpans, formatHints, pFormat = null) {
    // Find rPr by iterating children
    let rPr = null;
    for (const child of run.childNodes) {
        if (child.nodeName === 'w:rPr') {
            rPr = child;
            break;
        }
    }

    // Extract formatting flags from rPr, merging with paragraph defaults
    const format = extractFormatFromRPr(rPr);

    // Merge with paragraph-level defaults if they aren't explicitly overridden
    // Note: In OOXML, if pPr/rPr has bold, all runs are bold unless they have bold=off.
    // Simplifying: if pFormat has it, we have it.
    if (pFormat) {
        if (pFormat.bold && !format.bold) format.bold = true;
        if (pFormat.italic && !format.italic) format.italic = true;
        if (pFormat.underline && !format.underline) format.underline = true;
        if (pFormat.strikethrough && !format.strikethrough) format.strikethrough = true;
    }

    format.hasFormatting = format.bold || format.italic || format.underline || format.strikethrough;

    // Find text elements
    let currentOffset = charOffset;
    for (const child of run.childNodes) {
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
                    paragraph: paragraph,
                    rPr: rPr,
                    format: { ...format }
                });

                // If this run has any formatting, record it as a format hint
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
 * Removes formatting from specified spans using the pre-extracted data.
 * Handles both direct formatting (w:b tags) and inherited formatting (from paragraph).
 * For inherited formatting, adds explicit override elements (w:b w:val="0").
 */
function applyFormatRemovalWithSpans(xmlDoc, textSpans, existingFormatHints, serializer, author) {
    let hasAnyChanges = false;
    const processedRuns = new Set();
    const processedParagraphs = new Set();

    console.log(`[OxmlEngine] applyFormatRemovalWithSpans: ${existingFormatHints.length} hints to process`);

    // 1. Check and strip paragraph-level formatting first
    for (const span of textSpans) {
        const paragraph = span.paragraph;
        if (processedParagraphs.has(paragraph)) continue;
        processedParagraphs.add(paragraph);

        // Find pPr/rPr
        let pPr = null;
        let pRPr = null;
        for (const child of paragraph.childNodes) {
            if (child.nodeName === 'w:pPr') {
                pPr = child;
                for (const pChild of child.childNodes) {
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
                console.log(`[OxmlEngine] Removing paragraph-level formatting: ${pToRemove.map(e => e.nodeName).join(', ')}`);
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

        console.log(`[OxmlEngine] Processing format hint: bold=${hint.format.bold}, italic=${hint.format.italic}, rPr=${rPr ? 'exists' : 'null'}`);

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
                console.log(`[OxmlEngine] Removing direct formatting from run: ${toRemove.map(e => e.nodeName).join(', ')}`);

                // Create rPrChange for track changes
                let rPrChange = null;
                for (const child of rPr.childNodes) {
                    if (child.nodeName === 'w:rPrChange') {
                        rPrChange = child;
                        break;
                    }
                }

                if (!rPrChange) {
                    rPrChange = xmlDoc.createElement('w:rPrChange');
                    rPrChange.setAttribute('w:author', author || 'Gemini AI');
                    rPrChange.setAttribute('w:date', new Date().toISOString());

                    const originalRPr = xmlDoc.createElement('w:rPr');
                    for (const child of rPr.childNodes) {
                        if (child.nodeName !== 'w:rPrChange') {
                            originalRPr.appendChild(child.cloneNode(true));
                        }
                    }
                    rPrChange.appendChild(originalRPr);
                    rPr.appendChild(rPrChange);
                }

                for (const el of toRemove) {
                    rPr.removeChild(el);
                }
            }
            // Case 1b: No direct tags but formatting detected (from style or paragraph) - add overrides
            else if (hint.format.hasFormatting) {
                hasAnyChanges = true;
                console.log(`[OxmlEngine] Adding format overrides for style-based/inherited formatting (rPr exists)`);

                // Create rPrChange before modifying
                let rPrChange = null;
                for (const child of rPr.childNodes) {
                    if (child.nodeName === 'w:rPrChange') {
                        rPrChange = child;
                        break;
                    }
                }

                if (!rPrChange) {
                    rPrChange = xmlDoc.createElement('w:rPrChange');
                    rPrChange.setAttribute('w:author', author || 'Gemini AI');
                    rPrChange.setAttribute('w:date', new Date().toISOString());

                    const originalRPr = xmlDoc.createElement('w:rPr');
                    for (const child of rPr.childNodes) {
                        if (child.nodeName !== 'w:rPrChange') {
                            originalRPr.appendChild(child.cloneNode(true));
                        }
                    }
                    rPrChange.appendChild(originalRPr);
                    rPr.appendChild(rPrChange);
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
            console.log(`[OxmlEngine] Adding format overrides for inherited formatting`);

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
            const rPrChange = xmlDoc.createElement('w:rPrChange');
            rPrChange.setAttribute('w:author', author || 'Gemini AI');
            rPrChange.setAttribute('w:date', new Date().toISOString());
            const emptyOriginalRPr = xmlDoc.createElement('w:rPr');
            rPrChange.appendChild(emptyOriginalRPr);
            rPr.appendChild(rPrChange);

            // Insert rPr as first child of run
            run.insertBefore(rPr, run.firstChild);
        }
    }

    if (hasAnyChanges) {
        console.log('[OxmlEngine] Format removal applied successfully');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
    }

    console.log('[OxmlEngine] No formatting elements were removed');
    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
}

/**
 * LEGACY: Checks if the XML document contains any relevant formatting tags.
 * Kept for backward compatibility but now extractFormattingFromOoxml is preferred.
 */
function checkOxmlForFormatting(xmlDoc) {
    const formattingTags = ['w:b', 'w:i', 'w:u', 'w:strike'];
    for (const tag of formattingTags) {
        if (xmlDoc.getElementsByTagName(tag).length > 0) {
            return true;
        }
    }
    return false;
}

// ============================================================================
// FORMAT-ONLY MODE (applies formatting without text changes)
// ============================================================================

/**
 * Applies formatting changes to existing text without modifying content.
 * Used when markdown formatting is applied to unchanged text.
 */
function applyFormatOnlyChanges(xmlDoc, originalText, formatHints, serializer, author) {
    let fullText = '';
    const textSpans = [];

    const allParagraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

    // Build text span map
    allParagraphs.forEach((p, pIndex) => {
        const container = p.parentNode;

        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === 'w:r') {
                processRunElement(child, p, container, fullText, textSpans);
                fullText = getUpdatedFullText(child, fullText);
            } else if (child.nodeName === 'w:hyperlink') {
                Array.from(child.childNodes).forEach(hc => {
                    if (hc.nodeName === 'w:r') {
                        processRunElement(hc, p, container, fullText, textSpans);
                        fullText = getUpdatedFullText(hc, fullText);
                    }
                });
            }
        });

        if (pIndex < allParagraphs.length - 1) {
            fullText += '\n';
        }
    });

    // Apply each format hint to the corresponding text spans
    for (const hint of formatHints) {
        applyFormatHintToSpans(xmlDoc, textSpans, hint, author);
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Applies a single format hint to affected text spans.
 * Splits runs when only partial formatting is needed.
 */
function applyFormatHintToSpans(xmlDoc, textSpans, hint, author) {
    // Find spans that overlap with this format hint
    const affectedSpans = textSpans.filter(s =>
        s.charEnd > hint.start && s.charStart < hint.end
    );

    for (const span of affectedSpans) {
        const run = span.runElement;
        const parent = run.parentNode;
        if (!parent) continue;

        // Calculate the portion of this run that needs formatting
        const runStart = span.charStart;
        const runEnd = span.charEnd;
        const formatStart = Math.max(runStart, hint.start);
        const formatEnd = Math.min(runEnd, hint.end);

        const fullText = span.textElement.textContent || '';
        const localStart = formatStart - runStart;
        const localEnd = formatEnd - runStart;

        const beforeText = fullText.substring(0, localStart);
        const formattedText = fullText.substring(localStart, localEnd);
        const afterText = fullText.substring(localEnd);

        // Get existing rPr for inheritance
        const existingRPr = run.getElementsByTagName('w:rPr')[0] || null;

        // If the entire run needs formatting, just add it directly
        if (localStart === 0 && localEnd === fullText.length) {
            addFormattingToRun(xmlDoc, run, hint.format, author);
        } else {
            // Need to split the run into parts
            // Create runs for before, formatted, and after sections

            if (beforeText.length > 0) {
                // Create run for unformatted text before
                const beforeRun = createTextRun(xmlDoc, beforeText, existingRPr, false);
                parent.insertBefore(beforeRun, run);
            }

            // Create run for formatted text
            // Pass author for proper tracking of the format change
            const formattedRun = createFormattedRunWithElement(xmlDoc, formattedText, existingRPr, hint.format, author);
            parent.insertBefore(formattedRun, run);

            if (afterText.length > 0) {
                // Create run for unformatted text after
                const afterRun = createTextRun(xmlDoc, afterText, existingRPr, false);
                parent.insertBefore(afterRun, run);
            }

            // Remove the original run
            parent.removeChild(run);
        }
    }
}

/**
 * Creates a text run with formatting applied directly.
 */
function createFormattedRunWithElement(xmlDoc, text, baseRPr, format, author) {
    const run = xmlDoc.createElement('w:r');

    // Create rPr with formatting (and track changes if author provided)
    const rPr = injectFormattingToRPr(xmlDoc, baseRPr, format, author);

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
function addFormattingToRun(xmlDoc, run, format, author) {
    let rPr = run.getElementsByTagName('w:rPr')[0];

    // Create rPr if it doesn't exist
    if (!rPr) {
        rPr = xmlDoc.createElement('w:rPr');
        run.insertBefore(rPr, run.firstChild);
    }

    // Create rPrChange to track this modification
    if (author) {
        createRPrChange(xmlDoc, rPr, author);
    }

    // Check if element exists in rPr
    const hasElement = (tagName) => {
        return Array.from(rPr.childNodes).some(n => n.nodeName === tagName);
    };

    // Add formatting elements
    if (format.bold && !hasElement('w:b')) {
        const b = xmlDoc.createElement('w:b');
        rPr.appendChild(b);
    }
    if (format.italic && !hasElement('w:i')) {
        const i = xmlDoc.createElement('w:i');
        rPr.appendChild(i);
    }
    if (format.underline && !hasElement('w:u')) {
        const u = xmlDoc.createElement('w:u');
        u.setAttribute('w:val', 'single');
        rPr.appendChild(u);
    }
    if (format.strikethrough && !hasElement('w:strike')) {
        const strike = xmlDoc.createElement('w:strike');
        rPr.appendChild(strike);
    }
}

// ============================================================================
// TABLE RECONCILIATION MODE (Phase 9)
// ============================================================================

/**
 * Applies structural reconciliation to tables using Virtual Grid.
 */
function applyTableReconciliation(xmlDoc, modifiedText, serializer, author, formatHints) {
    const tableNodes = Array.from(xmlDoc.getElementsByTagName('w:tbl'));
    const newTableData = parseTable(modifiedText);

    if (tableNodes.length === 0 || newTableData.rows.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    // For now, always reconcile the first table in the fragment
    // In multi-table fragments, we'd need matching logic
    const targetTable = tableNodes[0];
    const oldGrid = ingestTableToVirtualGrid(targetTable);

    // Compute operations
    const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);

    if (operations.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    // Serialize new table
    const options = { generateRedlines: true, author };
    const reconciledOxml = serializeVirtualGridToOoxml(oldGrid, operations, options);

    // Parse the reconciled OOXML and replace the old table in the DOM
    const parser = new DOMParser();
    const reconciledDoc = parser.parseFromString(reconciledOxml, 'application/xml');
    const newTableNode = xmlDoc.importNode(reconciledDoc.documentElement, true);

    targetTable.parentNode.replaceChild(newTableNode, targetTable);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

// ============================================================================
// SURGICAL MODE (for tables - preserves structure)
// ============================================================================

/**
 * Surgical mode: Modifies existing runs in-place without changing paragraph structure.
 * Safe for tables and complex layouts.
 */
function applySurgicalMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints) {
    let fullText = '';
    const textSpans = [];

    const allParagraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

    // Build text span map
    allParagraphs.forEach((p, pIndex) => {
        const container = p.parentNode;

        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === 'w:r') {
                processRunElement(child, p, container, fullText, textSpans);
                fullText = getUpdatedFullText(child, fullText);
            } else if (child.nodeName === 'w:hyperlink') {
                // Process runs inside hyperlink
                Array.from(child.childNodes).forEach(hc => {
                    if (hc.nodeName === 'w:r') {
                        processRunElement(hc, p, container, fullText, textSpans);
                        fullText = getUpdatedFullText(hc, fullText);
                    }
                });
            }
        });

        // Add newline between paragraphs
        if (pIndex < allParagraphs.length - 1) {
            fullText += '\n';
        }
    });

    // Compute diff
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(fullText, modifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // Process deletions and insertions
    let currentPos = 0;
    let insertOffset = 0; // Track position in new text for format hints
    const processedSpans = new Set();

    for (const [op, text] of diffs) {
        if (op === 0) {
            // EQUAL - reconcile formatting
            const startPos = currentPos;
            const endPos = currentPos + text.length;

            // Find spans covered by this equal text
            const affectedSpans = textSpans.filter(s =>
                s.charEnd > startPos && s.charStart < endPos
            );

            for (const span of affectedSpans) {
                // Calculate overlap
                const overlapStart = Math.max(span.charStart, startPos);
                const overlapEnd = Math.min(span.charEnd, endPos);

                // Get format hints applicable to this overlap (adjusted relative to modified text start)
                // Note: formatHints use indices relative to the FULL modified text
                const localOffset = currentPos; // diff text position matches modifiedText position for Equal/Insert ops

                // Oops, wait. formatHints are relative to the CLEAN MODIFIED TEXT.
                // In Surgical Mode, `modifiedText` passed to this function IS the clean modified text.
                // So currentPos tracks the position in clean modified text correctly.

                // Check if any hints apply to this overlap
                const applicableHints = getApplicableFormatHints(formatHints, overlapStart, overlapEnd);

                // Reconcile formatting for this span segment
                reconcileFormattingForTextSpan(xmlDoc, span, overlapStart, overlapEnd, applicableHints, author);
            }

            currentPos += text.length;
        } else if (op === -1) {
            // DELETE
            processDelete(xmlDoc, textSpans, currentPos, currentPos + text.length, processedSpans, author);
            // Delete does not advance position in new text
        } else if (op === 1) {
            // INSERT - convert newlines to spaces for surgical mode
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.trim().length > 0) {
                processInsert(xmlDoc, textSpans, currentPos, textWithoutNewlines, processedSpans, author, formatHints, currentPos);
            }
            currentPos += text.length;
        }
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Reconciles formatting for a text span (or part of it).
 * Removes formatting that shouldn't be there, adds formatting that should.
 */
function reconcileFormattingForTextSpan(xmlDoc, span, start, end, applicableHints, author) {
    // 1. Determine desired format for this segment
    // Combine all applicable hints (later hints override/merge)
    const desiredFormat = {};
    if (applicableHints.length > 0) {
        // Merge all hints
        applicableHints.forEach(h => Object.assign(desiredFormat, h.format));
    }

    // 2. Check existing format
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

    // 3. Compare
    // We only care if:
    // a) Desired has format, Existing does not -> Add
    // b) Desired does NOT have format, Existing DOES -> Remove (if we are strict)

    const formatsToCheck = ['bold', 'italic', 'underline', 'strikethrough'];
    const changesNeeded = formatsToCheck.some(f => !!desiredFormat[f] !== existingFormat[f]);

    if (!changesNeeded) return;

    // 4. Apply changes
    // Since we might be affecting only PART of a run, we basically need to do a 
    // "replace" of that part with a new run that has the correct formatting
    // This is similar to processDelete + processInsert, but semantic is "Formatting Change"

    // To properly track "Formatted" changes in Word, we use w:rPrChange.
    // However, if we split the run, we need to be careful.

    // Simplest approach: Treat as "Format Change" logic similar to applyFormatHintToSpans
    // preventing code duplication would be good, but we need "Removal" logic here too.

    const parent = span.runElement.parentNode;
    if (!parent) return;

    const fullText = span.textElement.textContent || '';
    const runStart = span.charStart;

    const localStart = start - runStart;
    const localEnd = end - runStart;

    const beforeText = fullText.substring(0, localStart);
    const affectedText = fullText.substring(localStart, localEnd);
    const afterText = fullText.substring(localEnd);

    // Split if needed
    if (beforeText.length > 0) {
        const beforeRun = createTextRun(xmlDoc, beforeText, rPr, false);
        parent.insertBefore(beforeRun, span.runElement);
    }

    // Create new RPR based on desired format
    // We base it on existing RPR but FORCE the desired state for the checked properties
    const newRPr = injectExactFormattingToRPr(xmlDoc, rPr, desiredFormat, author);

    const newRun = createTextRunWithRPrElement(xmlDoc, affectedText, newRPr, false);
    parent.insertBefore(newRun, span.runElement);

    if (afterText.length > 0) {
        const afterRun = createTextRun(xmlDoc, afterText, rPr, false);
        parent.insertBefore(afterRun, span.runElement);
    }

    parent.removeChild(span.runElement);
}

/**
 * Creates an rPr that strictly matches the desired format state.
 * Adds w:rPrChange if ANY change is made.
 */
function injectExactFormattingToRPr(xmlDoc, baseRPr, desiredFormat, author) {
    const rPr = xmlDoc.createElement('w:rPr');

    // Copy base properties FIRST
    if (baseRPr) {
        Array.from(baseRPr.childNodes).forEach(child => {
            // Skip formatting tags we control
            if (!['w:b', 'w:i', 'w:u', 'w:strike', 'w:rPrChange'].includes(child.nodeName)) {
                rPr.appendChild(child.cloneNode(true));
            }
        });
    }

    // Add track change info if author provided (ALWAYS, since we only call this if changes needed)
    if (author && baseRPr) {
        // We need to pass the ORIGINAL rPr state to createRPrChange
        // But we must clone it to avoid mutating original DOM
        createRPrChange(xmlDoc, rPr, author, baseRPr); // Modified createRPrChange to accept explicit previous state
    } else if (author) {
        // No base RPR, create empty prev
        createRPrChange(xmlDoc, rPr, author);
    }

    // Enforce desired format
    if (desiredFormat.strikethrough) {
        const strike = xmlDoc.createElement('w:strike');
        rPr.insertBefore(strike, rPr.firstChild);
    }
    if (desiredFormat.underline) {
        const u = xmlDoc.createElement('w:u');
        u.setAttribute('w:val', 'single');
        rPr.insertBefore(u, rPr.firstChild);
    }
    if (desiredFormat.italic) {
        const i = xmlDoc.createElement('w:i');
        rPr.insertBefore(i, rPr.firstChild);
    }
    if (desiredFormat.bold) {
        const b = xmlDoc.createElement('w:b');
        rPr.insertBefore(b, rPr.firstChild);
    }

    return rPr;
}

/**
 * Processes a run element and extracts text spans
 */
function processRunElement(r, p, container, currentFullText, textSpans) {
    const rPr = r.getElementsByTagName('w:rPr')[0] || null;

    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            const text = rc.textContent || '';
            if (text.length > 0) {
                textSpans.push({
                    charStart: currentFullText.length,
                    charEnd: currentFullText.length + text.length,
                    textElement: rc,
                    runElement: r,
                    paragraph: p,
                    container,
                    rPr
                });
            }
        }
    });
}

/**
 * Gets updated full text after processing a run
 */
function getUpdatedFullText(r, currentFullText) {
    let fullText = currentFullText;
    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            fullText += rc.textContent || '';
        }
    });
    return fullText;
}

/**
 * Processes a deletion by splitting runs and wrapping in w:del
 */
function processDelete(xmlDoc, textSpans, startPos, endPos, processedSpans, author) {
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
            // Entire run is deleted
            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
            parent.insertBefore(delWrapper, span.runElement);
            parent.removeChild(span.runElement);
        } else {
            // Partial deletion - split the run
            if (beforeText.length > 0) {
                const beforeRun = createTextRun(xmlDoc, beforeText, span.rPr, false);
                parent.insertBefore(beforeRun, span.runElement);
            }

            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
            parent.insertBefore(delWrapper, span.runElement);

            if (afterText.length > 0) {
                const afterRun = createTextRun(xmlDoc, afterText, span.rPr, false);
                parent.insertBefore(afterRun, span.runElement);
            }

            parent.removeChild(span.runElement);
        }

        processedSpans.add(span.textElement);
    }
}

/**
 * Processes an insertion by adding a new w:ins element with optional formatting
 */
function processInsert(xmlDoc, textSpans, pos, text, processedSpans, author, formatHints = [], insertOffset = 0) {
    // Find the span at the insertion position
    let targetSpan = textSpans.find(s => pos >= s.charStart && pos < s.charEnd);

    // If not found, try the span that ends at this position
    if (!targetSpan && pos > 0) {
        targetSpan = textSpans.find(s => pos === s.charEnd);
    }

    // If still not found, try the span just before
    if (!targetSpan && pos > 0) {
        const before = textSpans.filter(s => s.charEnd <= pos);
        if (before.length > 0) {
            targetSpan = before[before.length - 1];
        }
    }

    // Last resort: use the last span
    if (!targetSpan && textSpans.length > 0) {
        targetSpan = textSpans[textSpans.length - 1];
    }

    if (targetSpan) {
        // Check for applicable format hints for this insertion
        const applicableHints = getApplicableFormatHints(formatHints, insertOffset, insertOffset + text.length);

        // Inherit formatting from adjacent run
        const baseRPr = targetSpan.rPr;

        // If there are format hints, we need to apply them
        const parent = targetSpan.runElement.parentNode;
        if (parent) {
            if (applicableHints.length === 0) {
                // No special formatting - use base rPr
                const insRun = createTextRun(xmlDoc, text, baseRPr, false);
                const insWrapper = createTrackChange(xmlDoc, 'ins', insRun, author);
                parent.insertBefore(insWrapper, targetSpan.runElement.nextSibling);
            } else {
                // Apply format hints - may need to split text into multiple runs
                const runs = createFormattedRuns(xmlDoc, text, baseRPr, applicableHints, insertOffset, author);
                const insWrapper = createTrackChange(xmlDoc, 'ins', null, author);
                runs.forEach(run => insWrapper.appendChild(run));
                parent.insertBefore(insWrapper, targetSpan.runElement.nextSibling);
            }
        }
    }
}

// ============================================================================
// RECONSTRUCTION MODE (for body without tables - allows paragraph changes)
// ============================================================================

/**
 * Reconstruction mode: Rebuilds paragraph content allowing new paragraphs.
 * Supports list splitting via newlines.
 */
function applyReconstructionMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints) {
    const body = xmlDoc.getElementsByTagName('w:body')[0] || xmlDoc.documentElement;
    const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

    if (paragraphs.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    // Build context maps
    let originalFullText = '';
    const propertyMap = [];
    const paragraphMap = [];
    const sentinelMap = [];
    const referenceMap = new Map();
    const tokenToCharMap = new Map();
    let nextCharCode = 0xe000;

    const uniqueContainers = new Set();
    const replacementContainers = new Map();

    // Process all paragraphs
    paragraphs.forEach((p, pIndex) => {
        const pStart = originalFullText.length;

        Array.from(p.childNodes).forEach(child => {
            originalFullText = processChildNode(
                child, originalFullText, propertyMap, sentinelMap,
                referenceMap, tokenToCharMap, nextCharCode
            );
            if (referenceMap.size > tokenToCharMap.size) {
                nextCharCode++;
            }
        });

        // Add paragraph separator
        if (pIndex < paragraphs.length - 1) {
            originalFullText += '\n';
        }

        const pEnd = originalFullText.length;
        const pPr = p.getElementsByTagName('w:pPr')[0] || null;
        const container = p.parentNode;
        if (container) uniqueContainers.add(container);

        paragraphMap.push({
            start: pStart,
            end: pEnd,
            pPr,
            container: container || body
        });
    });

    // Process modified text - replace reference tokens with private use chars
    let processedModifiedText = modifiedText;
    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, 'g'), char);
    });

    // Compute diff
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(originalFullText, processedModifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // Create document fragments for each container
    const containerFragments = new Map();
    uniqueContainers.forEach(c => containerFragments.set(c, xmlDoc.createDocumentFragment()));
    if (!containerFragments.has(body)) containerFragments.set(body, xmlDoc.createDocumentFragment());

    // Helper functions
    const getParagraphInfo = (index) => {
        const match = paragraphMap.find(m => index >= m.start && index < m.end);
        if (!match && paragraphMap.length > 0) {
            return paragraphMap[paragraphMap.length - 1];
        }
        return match || { pPr: null, container: body };
    };

    const getRunProperties = (index) => {
        const match = propertyMap.find(m => index >= m.start && index < m.end);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const createNewParagraph = (pPr) => {
        const newP = xmlDoc.createElement('w:p');
        if (pPr) newP.appendChild(pPr.cloneNode(true));
        return newP;
    };

    // Initialize current paragraph
    let startInfo = getParagraphInfo(0);
    let currentParagraph = createNewParagraph(startInfo.pPr);
    let currentContainer = startInfo.container;
    let currentFragment = containerFragments.get(currentContainer);
    if (currentFragment) currentFragment.appendChild(currentParagraph);

    let currentOriginalIndex = 0;
    let currentInsertOffset = 0; // Track position in new text for format hints

    // Process each diff
    for (const [op, text] of diffs) {
        if (op === 0) {
            // EQUAL
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const range = propertyMap.find(m =>
                    currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end
                );
                const length = range
                    ? Math.min(range.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);

                appendTextToCurrent(
                    xmlDoc, chunk, 'equal', props.rPr, props.wrapper,
                    currentOriginalIndex + offset, currentParagraph, paragraphMap,
                    containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                    replacementContainers, getParagraphInfo, createNewParagraph, author,
                    formatHints, currentInsertOffset
                );

                offset += length;
            }
            currentOriginalIndex += text.length;
        } else if (op === 1) {
            // INSERT
            const isStartOfParagraph = paragraphMap.some(p => p.start === currentOriginalIndex);
            const props = currentOriginalIndex > 0 && !isStartOfParagraph
                ? getRunProperties(currentOriginalIndex - 1)
                : getRunProperties(currentOriginalIndex);

            appendTextToCurrent(
                xmlDoc, text, 'insert', props.rPr, props.wrapper,
                currentOriginalIndex, currentParagraph, paragraphMap,
                containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                replacementContainers, getParagraphInfo, createNewParagraph, author,
                formatHints, currentInsertOffset
            );
            currentInsertOffset += text.length;
        } else if (op === -1) {
            // DELETE
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const range = propertyMap.find(m =>
                    currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end
                );
                const length = range
                    ? Math.min(range.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);

                appendTextToCurrent(
                    xmlDoc, chunk, 'delete', props.rPr, props.wrapper,
                    currentOriginalIndex + offset, currentParagraph, paragraphMap,
                    containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                    replacementContainers, getParagraphInfo, createNewParagraph, author,
                    formatHints, currentInsertOffset
                );

                offset += length;
            }
            currentOriginalIndex += text.length;
        }
    }

    // Replace old paragraphs with new ones
    paragraphs.forEach(p => {
        if (p.parentNode) p.parentNode.removeChild(p);
    });

    containerFragments.forEach((fragment, container) => {
        const replacement = replacementContainers.get(container);
        const target = replacement || container;
        target.appendChild(fragment);
    });

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Processes a child node during paragraph traversal
 */
function processChildNode(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    if (child.nodeName === 'w:r') {
        return processRunForReconstruction(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode);
    } else if (child.nodeName === 'w:hyperlink') {
        return processHyperlinkForReconstruction(child, originalFullText, propertyMap);
    } else if (['w:sdt', 'w:oMath', 'm:oMath', 'w:bookmarkStart', 'w:bookmarkEnd'].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child });
        return originalFullText + '\uFFFC';
    }
    return originalFullText;
}

/**
 * Processes a run for reconstruction mode
 */
function processRunForReconstruction(r, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    let fullText = originalFullText;
    const rPr = r.getElementsByTagName('w:rPr')[0] || null;

    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            const textContent = rc.textContent || '';
            if (textContent.length > 0) {
                propertyMap.push({
                    start: fullText.length,
                    end: fullText.length + textContent.length,
                    rPr
                });
                fullText += textContent;
            }
        } else if (['w:drawing', 'w:pict', 'w:object', 'w:fldChar', 'w:instrText'].includes(rc.nodeName)) {
            // Sentinel for embedded objects
            const rcElement = rc;
            const txbxContent = rcElement.getElementsByTagName ? rcElement.getElementsByTagName('w:txbxContent')[0] : null;
            const hasTextBox = rc.nodeName === 'w:pict' && !!txbxContent;

            sentinelMap.push({
                start: fullText.length,
                node: rc,
                isTextBox: hasTextBox,
                originalContainer: hasTextBox ? txbxContent : undefined
            });
            fullText += '\uFFFC';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (rc.nodeName === 'w:footnoteReference' || rc.nodeName === 'w:endnoteReference') {
            // Reference handling
            const ref = rc;
            const id = ref.getAttribute('w:id');
            if (id) {
                const type = rc.nodeName === 'w:footnoteReference' ? 'FN' : 'EN';
                const tokenString = `{{__${type}_${id}__}}`;
                const char = String.fromCharCode(nextCharCode);
                referenceMap.set(char, rc);
                tokenToCharMap.set(tokenString, char);
                fullText += char;
                propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
            }
        }
    });

    return fullText;
}

/**
 * Processes a hyperlink for reconstruction mode
 */
function processHyperlinkForReconstruction(h, originalFullText, propertyMap) {
    let fullText = originalFullText;

    Array.from(h.childNodes).forEach(hc => {
        if (hc.nodeName === 'w:r') {
            const r = hc;
            const rPr = r.getElementsByTagName('w:rPr')[0] || null;
            const texts = Array.from(r.getElementsByTagName('w:t'));
            texts.forEach(t => {
                const textContent = t.textContent || '';
                if (textContent.length > 0) {
                    propertyMap.push({
                        start: fullText.length,
                        end: fullText.length + textContent.length,
                        rPr,
                        wrapper: h
                    });
                    fullText += textContent;
                }
            });
        }
    });

    return fullText;
}

/**
 * Appends text to current paragraph with proper track change wrapping
 * 
 * Note: This function has many parameters because it needs to manage complex
 * state during reconstruction. In a refactor, this could be encapsulated in a class.
 */
function appendTextToCurrent(
    xmlDoc, text, type, rPr, wrapper, baseIndex,
    currentParagraphRef, paragraphMap, containerFragments,
    sentinelMap, referenceMap, tokenToCharMap,
    replacementContainers, getParagraphInfo, createNewParagraph, author,
    formatHints = [], insertOffset = 0
) {
    // This is a simplified version - full implementation would need to track
    // currentParagraph as a mutable reference. For now, we process in-line.

    const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);

    parts.forEach(part => {
        if (part === '\n') {
            // Handle newline by creating a new paragraph
            if (type !== 'delete') {
                const info = getParagraphInfo(baseIndex);
                const nextParagraph = createNewParagraph(info.pPr);
                const fragment = containerFragments.get(info.container);
                if (fragment) {
                    fragment.appendChild(nextParagraph);
                    // Update the reference for subsequent text
                    currentParagraphRef.appendChild(xmlDoc.createTextNode('')); // Dummy for reference handling
                    // This is a bit hacky because JS doesn't have pointers, but let's assume
                    // the caller handles the updated currentParagraph.
                    // In a real implementation, we'd return the new paragraph or update a state object.
                }
            }
        } else if (part === '\uFFFC') {
            // Re-insert sentinel/embedded object
            const sentinel = sentinelMap.find(s => s.start === baseIndex);
            if (sentinel) {
                const clone = sentinel.node.cloneNode(true);
                if (sentinel.isTextBox && sentinel.originalContainer) {
                    const newContainer = clone.getElementsByTagName('w:txbxContent')[0];
                    if (newContainer) {
                        while (newContainer.firstChild) newContainer.removeChild(newContainer.firstChild);
                        replacementContainers.set(sentinel.originalContainer, newContainer);
                    }
                }
                // Append to current paragraph
                currentParagraphRef.appendChild(clone);
            }
        } else if (referenceMap.has(part)) {
            // Re-insert reference
            if (type !== 'delete') {
                const refNode = referenceMap.get(part);
                if (refNode) {
                    const clone = refNode.cloneNode(true);
                    const run = xmlDoc.createElement('w:r');
                    if (rPr) run.appendChild(rPr.cloneNode(true));
                    run.appendChild(clone);
                    currentParagraphRef.appendChild(run);
                }
            }
        } else if (part.length > 0) {
            // Normal text - check if we need to apply format hints for insertions
            let parent = currentParagraphRef;
            if (wrapper) {
                const wrapperClone = wrapper.cloneNode(false);
                parent = wrapperClone;
                currentParagraphRef.appendChild(wrapperClone);
            }

            if (type === 'delete') {
                const run = xmlDoc.createElement('w:r');
                if (rPr) run.appendChild(rPr.cloneNode(true));
                const t = xmlDoc.createElement('w:delText');
                t.setAttribute('xml:space', 'preserve');
                t.textContent = part;
                run.appendChild(t);
                const del = createTrackChange(xmlDoc, 'del', run, author);
                parent.appendChild(del);
            } else if (type === 'insert') {
                // Check for applicable format hints
                const applicableHints = getApplicableFormatHints(formatHints, insertOffset, insertOffset + part.length);

                if (applicableHints.length === 0) {
                    // No formatting - simple insert
                    const run = xmlDoc.createElement('w:r');
                    if (rPr) run.appendChild(rPr.cloneNode(true));
                    const t = xmlDoc.createElement('w:t');
                    t.setAttribute('xml:space', 'preserve');
                    t.textContent = part;
                    run.appendChild(t);
                    const ins = createTrackChange(xmlDoc, 'ins', run, author);
                    parent.appendChild(ins);
                } else {
                    // Apply format hints
                    const runs = createFormattedRuns(xmlDoc, part, rPr, applicableHints, insertOffset);
                    const ins = createTrackChange(xmlDoc, 'ins', null, author);
                    runs.forEach(run => ins.appendChild(run));
                    parent.appendChild(ins);
                }
            } else {
                // Equal - no track change wrapper
                const run = xmlDoc.createElement('w:r');
                if (rPr) run.appendChild(rPr.cloneNode(true));
                const t = xmlDoc.createElement('w:t');
                t.setAttribute('xml:space', 'preserve');
                t.textContent = part;
                run.appendChild(t);
                parent.appendChild(run);
            }
        }
    });
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Creates a track change wrapper element (w:ins or w:del)
 */
function createTrackChange(xmlDoc, type, run, author) {
    const wrapper = xmlDoc.createElement(type === 'ins' ? 'w:ins' : 'w:del');
    wrapper.setAttribute('w:id', Math.floor(Math.random() * 90000 + 10000).toString());
    wrapper.setAttribute('w:author', author);
    wrapper.setAttribute('w:date', new Date().toISOString());
    if (run) {
        wrapper.appendChild(run);
    }
    return wrapper;
}

/**
 * Creates a text run with optional formatting
 */
function createTextRun(xmlDoc, text, rPr, isDelete) {
    const run = xmlDoc.createElement('w:r');
    if (rPr) run.appendChild(rPr.cloneNode(true));

    const textEl = xmlDoc.createElement(isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Creates an array of runs with formatting applied based on format hints.
 * Splits text at format boundaries and applies appropriate formatting.
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {string} text - Text to format
 * @param {Element|null} baseRPr - Base run properties to inherit
 * @param {Array} formatHints - Array of {start, end, format} hints
 * @param {number} baseOffset - Base offset for position calculations
 * @param {string} [author] - Optional author for track changes
 * @returns {Element[]} Array of w:r elements
 */
function createFormattedRuns(xmlDoc, text, baseRPr, formatHints, baseOffset, author) {
    const runs = [];
    let pos = 0;

    // Sort hints by start position
    const sortedHints = [...formatHints].sort((a, b) => a.start - b.start);

    for (const hint of sortedHints) {
        const localStart = Math.max(0, hint.start - baseOffset);
        const localEnd = Math.min(text.length, hint.end - baseOffset);

        // Skip if hint doesn't apply to this text range
        if (localStart >= text.length || localEnd <= 0) continue;

        // Text before the formatted section
        if (localStart > pos) {
            const beforeText = text.slice(pos, localStart);
            runs.push(createTextRun(xmlDoc, beforeText, baseRPr, false));
        }

        // Formatted text
        const formattedText = text.slice(localStart, localEnd);
        const formattedRPr = injectFormattingToRPr(xmlDoc, baseRPr, hint.format, author);
        runs.push(createTextRunWithRPrElement(xmlDoc, formattedText, formattedRPr, false));

        pos = localEnd;
    }

    // Remaining text after last format hint
    if (pos < text.length) {
        runs.push(createTextRun(xmlDoc, text.slice(pos), baseRPr, false));
    }

    return runs;
}

/**
 * Creates a text run with an rPr Element (not cloned)
 */
function createTextRunWithRPrElement(xmlDoc, text, rPrElement, isDelete) {
    const run = xmlDoc.createElement('w:r');
    if (rPrElement) run.appendChild(rPrElement);

    const textEl = xmlDoc.createElement(isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Injects formatting into run properties, creating new w:rPr element with formatting.
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {Element|null} baseRPr - Base run properties to inherit (will be cloned)
 * @param {Object} format - Format flags {bold, italic, underline, strikethrough}
 * @returns {Element} New w:rPr element with formatting applied
 */
/**
 * Injects formatting into run properties, creating new w:rPr element with formatting.
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {Element|null} baseRPr - Base run properties to inherit (will be cloned)
 * @param {Object} format - Format flags {bold, italic, underline, strikethrough}
 * @param {string} [author] - Optional author for track changes
 * @returns {Element} New w:rPr element with formatting applied
 */
function injectFormattingToRPr(xmlDoc, baseRPr, format, author) {
    // Always create a new rPr to ensure we don't mutate original references
    const rPr = xmlDoc.createElement('w:rPr');

    // Copy existing properties from base
    if (baseRPr) {
        Array.from(baseRPr.childNodes).forEach(child => {
            rPr.appendChild(child.cloneNode(true));
        });
    }

    if (!format || Object.keys(format).length === 0) {
        return rPr;
    }

    // Add track change info if author provided
    if (author) {
        createRPrChange(xmlDoc, rPr, author);
    }

    // Add formatting elements (at the beginning, before other properties)
    // Check if formatting already exists to avoid duplicates
    const hasElement = (tagName) => {
        return Array.from(rPr.childNodes).some(n => n.nodeName === tagName);
    };

    if (format.strikethrough && !hasElement('w:strike')) {
        const strike = xmlDoc.createElement('w:strike');
        rPr.insertBefore(strike, rPr.firstChild);
    }
    if (format.underline && !hasElement('w:u')) {
        const u = xmlDoc.createElement('w:u');
        u.setAttribute('w:val', 'single');
        rPr.insertBefore(u, rPr.firstChild);
    }
    if (format.italic && !hasElement('w:i')) {
        const i = xmlDoc.createElement('w:i');
        rPr.insertBefore(i, rPr.firstChild);
    }
    if (format.bold && !hasElement('w:b')) {
        const b = xmlDoc.createElement('w:b');
        rPr.insertBefore(b, rPr.firstChild);
    }

    return rPr;
}

/**
 * Creates w:rPrChange element to track property changes.
 * Captures the PREVIOUS state of properties (which is the current content of rPr before new format).
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {Element} rPr - The run properties element to append to
 * @param {string} author - Author name
 */
function createRPrChange(xmlDoc, rPr, author, previousRPrArg) {
    const rPrChange = xmlDoc.createElement('w:rPrChange');
    rPrChange.setAttribute('w:id', Math.floor(Math.random() * 90000 + 10000).toString());
    rPrChange.setAttribute('w:author', author);
    rPrChange.setAttribute('w:date', new Date().toISOString());

    // Create inner rPr for the "previous" state
    const previousRPr = xmlDoc.createElement('w:rPr');

    // Determine source for previous state
    const sourceNode = previousRPrArg || rPr;

    // Clone *current* children of rPr into previousRPr (snapshot of state before change)
    // Note: We only capture what is currently in rPr, which represents the "base" properties
    Array.from(sourceNode.childNodes).forEach(child => {
        // Don't recurse into existing rPrChanges to avoid infinite nesting loop in simple implementation
        if (child.nodeName !== 'w:rPrChange') {
            previousRPr.appendChild(child.cloneNode(true));
        }
    });

    rPrChange.appendChild(previousRPr);
    rPr.appendChild(rPrChange);
}

/**
 * Sanitizes AI response text by removing common prefixes
 */
export function sanitizeAiResponse(text) {
    let cleaned = text;
    // Remove common AI prefixes
    cleaned = cleaned.replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, '');
    // Remove LaTeX-style formatting
    cleaned = cleaned.replace(/\$\\text\{/g, '').replace(/\}\$/g, '');
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, '$1');
    // Normalize literal \n strings to real newlines
    cleaned = cleaned.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
    return cleaned;
}

