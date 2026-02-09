/**
 * Word paragraph change routing/apply helper.
 *
 * Extracted from command layer so other Word-facing modules can reuse the
 * same reconciliation-driven route behavior.
 */

import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { wrapInDocumentFragment } from '../pipeline/serialization.js';
import { extractParagraphIdFromOoxml } from '../core/ooxml-identifiers.js';
import { buildReconciliationPlan, RoutePlanKind } from '../orchestration/route-plan.js';
import { applyStructuredListDirectOoxml } from './word-structured-list.js';
import {
    getParagraphOoxmlWithFallback,
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';

/**
 * Applies a single paragraph change using reconciliation route planning.
 *
 * @param {Object} change - Tool change payload
 * @param {Word.Paragraph} targetParagraph - Target paragraph proxy
 * @param {Word.RequestContext} context - Word request context
 * @param {Object} [options={}] - Route options
 * @param {boolean} [options.propertiesPreloaded=false] - Whether required paragraph properties are preloaded
 * @param {Object} options.services - Runtime service callbacks
 * @param {() => boolean} options.services.loadRedlineSetting - Redline enabled provider
 * @param {() => string} options.services.loadRedlineAuthor - Redline author provider
 * @param {(text: string) => string} options.services.markdownToWordHtml - Markdown->Word HTML transformer
 * @param {(text: string) => Promise<{ cleanText: string, formatHints: Array }>} options.services.preprocessMarkdownForParagraph - Markdown preprocessing helper
 * @param {(paragraph: Word.Paragraph, cleanText: string, hints: Array, context: Word.RequestContext) => Promise<void>} options.services.applyFormatHintsToRanges - Native format apply helper
 * @param {(paragraph: Word.Paragraph, originalText: string, hints: Array, context: Word.RequestContext) => Promise<void>} options.services.applyFormatRemovalToRanges - Native format removal helper
 * @returns {Promise<void>}
 */
export async function routeWordParagraphChange(change, targetParagraph, context, options = {}) {
    const {
        propertiesPreloaded = false,
        services = {}
    } = options;

    const {
        loadRedlineSetting,
        loadRedlineAuthor,
        markdownToWordHtml,
        preprocessMarkdownForParagraph,
        applyFormatHintsToRanges,
        applyFormatRemovalToRanges
    } = services;

    assertRequiredService(loadRedlineSetting, 'loadRedlineSetting');
    assertRequiredService(loadRedlineAuthor, 'loadRedlineAuthor');
    assertRequiredService(markdownToWordHtml, 'markdownToWordHtml');
    assertRequiredService(preprocessMarkdownForParagraph, 'preprocessMarkdownForParagraph');
    assertRequiredService(applyFormatHintsToRanges, 'applyFormatHintsToRanges');
    assertRequiredService(applyFormatRemovalToRanges, 'applyFormatRemovalToRanges');

    // Properties text, style, parentTableCellOrNullObject, parentTableOrNullObject
    // should ideally be pre-loaded by the caller to avoid syncs here.
    if (!propertiesPreloaded) {
        targetParagraph.load('text, style, parentTableCellOrNullObject, parentTableOrNullObject');
        await context.sync();
    }

    const originalText = targetParagraph.text;
    const plan = buildReconciliationPlan({
        originalText,
        newContent: change.newContent || change.content || ''
    });
    const newContent = plan.normalizedContent;

    // Treat multi-line marker content as list structure (including A./B./C.
    // and roman markers), even when the model chose edit_paragraph.
    if (plan.kind === RoutePlanKind.STRUCTURED_LIST_DIRECT) {
        console.log('[routeWordParagraphChange] Detected structured list content, using reconciliation list generation');
        const redlineEnabled = loadRedlineSetting();
        const redlineAuthor = loadRedlineAuthor();
        const parsedListData = plan.parsedListData;

        const listItems = (parsedListData.items || []).filter((item) => item && (item.type === 'numbered' || item.type === 'bullet'));
        console.log(`[routeWordParagraphChange] Structured list items parsed: ${listItems.length}`);

        if (listItems.length > 0) {
            let paragraphFont = null;
            if (targetParagraph.font && targetParagraph.font.name) {
                paragraphFont = targetParagraph.font.name;
            } else if (!propertiesPreloaded) {
                targetParagraph.load('font/name');
                await context.sync();
                paragraphFont = targetParagraph.font?.name || null;
            }

            try {
                const pipeline = new ReconciliationPipeline({
                    generateRedlines: redlineEnabled,
                    author: redlineAuthor,
                    font: paragraphFont || 'Calibri'
                });

                const result = await pipeline.executeListGeneration(
                    newContent,
                    null,
                    null,
                    originalText
                );

                const listOoxml = result?.ooxml || result?.oxml || '';
                const listIsValid = result?.isValid !== false;
                if (!listOoxml || !listIsValid) {
                    const warnings = Array.isArray(result?.warnings) ? result.warnings.join('; ') : 'none';
                    throw new Error(`Structured list generation produced invalid OOXML (isValid=${listIsValid}, warnings=${warnings})`);
                }

                const wrappedOoxml = wrapInDocumentFragment(listOoxml, {
                    includeNumbering: true,
                    numberingXml: result.numberingXml
                });

                await withNativeTrackingDisabled(context, async () => {
                    await insertOoxmlWithRangeFallback(targetParagraph, wrappedOoxml, 'Replace', context, 'routeWordParagraphChange/StructuredList');
                }, {
                    enabled: redlineEnabled,
                    logPrefix: 'routeWordParagraphChange/StructuredList'
                });
                return;
            } catch (structuredListError) {
                console.warn('[routeWordParagraphChange] Reconciliation structured-list path failed, falling back to direct list insertion:', structuredListError);
                await withNativeTrackingDisabled(context, async () => {
                    await applyStructuredListDirectOoxml(context, targetParagraph, parsedListData);
                }, {
                    enabled: redlineEnabled,
                    logPrefix: 'routeWordParagraphChange/StructuredListFallback'
                });
                return;
            }
        }

        throw new Error('[routeWordParagraphChange] Structured list conversion failed: no list items parsed.');
    }

    // Empty original text - try native APIs first.
    if (plan.kind === RoutePlanKind.EMPTY_FORMATTED_TEXT || plan.kind === RoutePlanKind.EMPTY_HTML) {
        console.log('Empty paragraph detected');
        if (plan.flags?.hasMarkdownTable) {
            console.log('Detected table in empty paragraph, using OOXML Hybrid Mode');
        }

        if (plan.kind === RoutePlanKind.EMPTY_FORMATTED_TEXT) {
            console.log('Empty paragraph with formatting - using insertText for simplicity');
            const { cleanText, formatHints } = await preprocessMarkdownForParagraph(newContent);
            targetParagraph.insertText(cleanText, 'Replace');
            await context.sync();

            if (formatHints.length > 0) {
                try {
                    await applyFormatHintsToRanges(targetParagraph, cleanText, formatHints, context);
                } catch (formatError) {
                    console.warn('Could not apply formatting:', formatError);
                }
            }
            return;
        }

        console.log('Using HTML insertion for empty paragraph');
        const htmlContent = markdownToWordHtml(newContent);
        targetParagraph.insertHtml(htmlContent, 'Replace');
        return;
    }

    // Block elements (headings, mixed content, etc.).
    if (plan.kind === RoutePlanKind.BLOCK_HTML) {
        console.log('Block elements detected, using HTML replacement');
        const htmlContent = markdownToWordHtml(newContent);
        targetParagraph.insertHtml(htmlContent, 'Replace');
        return;
    }

    // Use OOXML engine for text edits.
    console.log('Attempting OOXML Hybrid Mode for text edit');
    const redlineEnabled = loadRedlineSetting();

    if (!propertiesPreloaded) {
        targetParagraph.load('text');
        await context.sync();
    }

    const paragraphOriginalText = targetParagraph.text;
    const paragraphOoxmlRead = await getParagraphOoxmlWithFallback(targetParagraph, context, {
        propertiesPreloaded,
        logPrefix: 'OxmlEngine'
    });
    const paragraphOoxmlValue = paragraphOoxmlRead.ooxml;

    if (!paragraphOoxmlValue) {
        console.warn('[OxmlEngine] Unable to retrieve OOXML for paragraph; skipping OOXML edit');
        return;
    }

    console.log('[OxmlEngine] Original text:', paragraphOriginalText.length > 500 ? `${paragraphOriginalText.substring(0, 500)}...` : paragraphOriginalText);
    console.log('[OxmlEngine] Original text length:', paragraphOriginalText.length);
    const targetParagraphId = extractParagraphIdFromOoxml(paragraphOoxmlValue);

    const redlineAuthor = loadRedlineAuthor();
    const result = await applyRedlineToOxml(
        paragraphOoxmlValue,
        paragraphOriginalText,
        newContent,
        {
            author: redlineEnabled ? redlineAuthor : undefined,
            generateRedlines: redlineEnabled,
            targetParagraphId
        }
    );

    if (!result.hasChanges) {
        console.log('[OxmlEngine] No changes detected by engine');
        return;
    }

    if (result.useNativeApi && result.formatHints) {
        console.log('[OxmlEngine] Using native Font API for table cell formatting');
        await applyFormatHintsToRanges(targetParagraph, result.originalText, result.formatHints, context);
        console.log('Native API formatting successful');
        return;
    }

    if (result.isSurgicalFormatChange && result.surgicalChanges) {
        console.log(`[OxmlEngine] Applying ${result.surgicalChanges.length} surgical format changes`);

        let successfulSurgicalChanges = 0;

        await withNativeTrackingDisabled(context, async () => {
            const paragraphRange = targetParagraph.getRange();
            paragraphRange.load('text');
            await context.sync();

            for (const changeEntry of result.surgicalChanges) {
                try {
                    console.log(`[OxmlEngine] Surgical: searching for "${changeEntry.searchText}"`);

                    const searchResults = paragraphRange.search(changeEntry.searchText, {
                        matchCase: true,
                        matchWholeWord: false
                    });
                    searchResults.load('items/text');
                    await context.sync();

                    if (searchResults.items.length > 0) {
                        const targetRange = searchResults.items[0];
                        targetRange.insertOoxml(changeEntry.replacementOoxml, 'Replace');
                        await context.sync();
                        console.log(`[OxmlEngine] Surgical replacement applied for "${changeEntry.searchText}"`);
                        successfulSurgicalChanges++;
                    } else {
                        console.warn(`[OxmlEngine] Text not found for surgical replacement: "${changeEntry.searchText}"`);
                    }
                } catch (changeError) {
                    console.warn(`[OxmlEngine] Failed to apply surgical change: ${changeError.message}`);
                }
            }

            if (successfulSurgicalChanges === result.surgicalChanges.length) {
                console.log('Surgical format changes completed');
            } else {
                console.warn(`[OxmlEngine] Surgical replacements applied: ${successfulSurgicalChanges}/${result.surgicalChanges.length}`);
            }
        }, {
            enabled: true,
            logPrefix: 'OxmlEngine/Surgical'
        });

        if (successfulSurgicalChanges === result.surgicalChanges.length) {
            return;
        }

        console.warn('[OxmlEngine] Surgical format removal incomplete; no fallback configured');
        return;
    }

    if (result.useNativeApi && result.formatRemovalHints) {
        console.log('[OxmlEngine] Using native Font API for format removal');
        await applyFormatRemovalToRanges(targetParagraph, result.originalText, result.formatRemovalHints, context);
        console.log('Native API format removal successful');
        return;
    }

    console.log('[OxmlEngine] Generated OOXML with track changes, length:', result.oxml.length);

    try {
        await withNativeTrackingDisabled(context, async ({ originalMode, trackingDisabled }) => {
            console.log(`[OxmlEngine] Current track changes mode: ${originalMode}, redlineEnabled: ${redlineEnabled}, isFormatOnly: ${result.isFormatOnly}, shouldDisableTracking: ${trackingDisabled}`);
            if (trackingDisabled) {
                console.log('[OxmlEngine] Temporarily disabling Word track changes for text-based OOXML insertion');
            }

            targetParagraph.insertOoxml(result.oxml, 'Replace');
            await context.sync();
            console.log('OOXML Hybrid Mode reconciliation successful');

            if (result.oxml.includes('<w:numPr>') || result.oxml.includes('ListParagraph')) {
                try {
                    console.log('[OxmlEngine] Detected list in result, applying spacing workaround');

                    const pCount = (result.oxml.match(/<w:p>/g) || []).length;

                    if (pCount > 1) {
                        const paragraphs = context.document.body.paragraphs;
                        paragraphs.load('items/text');
                        await context.sync();

                        const targetIdx = targetParagraph.index || 0;

                        if (targetIdx + pCount - 1 < paragraphs.items.length) {
                            const lastListItem = paragraphs.items[targetIdx + pCount - 1];
                            lastListItem.insertParagraph('', 'After');
                            await context.sync();

                            console.log(`[OxmlEngine] Inserted dummy spacing paragraph after ${pCount} list items`);
                            await context.sync();
                            console.log('[OxmlEngine] Left dummy spacing paragraph for testing');
                        }
                    }
                } catch (spacingError) {
                    console.warn('[OxmlEngine] Spacing workaround failed (non-critical):', spacingError.message);
                }
            }
            if (trackingDisabled) {
                console.log(`[OxmlEngine] Restoring track changes mode to: ${originalMode}`);
            }
        }, {
            enabled: true,
            logPrefix: 'OxmlEngine/Insert'
        });
    } catch (insertError) {
        console.error('OOXML insertion failed:', insertError.message);
        console.log('Falling back to simple text replacement');
        targetParagraph.insertText(newContent, 'Replace');
        await context.sync();
    }
}

function assertRequiredService(fn, name) {
    if (typeof fn === 'function') return;
    throw new Error(`[routeWordParagraphChange] Missing required service: ${name}`);
}
