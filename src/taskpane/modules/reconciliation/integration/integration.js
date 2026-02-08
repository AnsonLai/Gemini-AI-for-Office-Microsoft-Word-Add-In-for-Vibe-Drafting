/**
 * OOXML Reconciliation Pipeline - Integration Helper
 * 
 * Provides a bridge between the Word Add-in's existing flow and 
 * the new OOXML-based reconciliation pipeline.
 */

import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { log, warn, error } from '../adapters/logger.js';

/**
 * Applies OOXML-level reconciliation to a paragraph.
 * This is a drop-in replacement for applyWordLevelDiffs that works
 * at the OOXML level for proper track changes.
 * 
 * @param {Word.Paragraph} paragraph - The paragraph to modify
 * @param {string} newText - New text with optional markdown formatting
 * @param {Word.RequestContext} context - Word API context
 * @param {Object} options - Options
 * @param {boolean} options.generateRedlines - Whether to generate track changes
 * @param {string} options.author - Author for track changes
 * @returns {Promise<{success: boolean, message: string}>}
 */
export async function applyReconciliationToParagraph(paragraph, newText, context, options = {}) {
    const batchResult = await applyReconciliationToParagraphBatch(
        [{ paragraph, newText }],
        context,
        options
    );

    const first = batchResult.results[0];
    return {
        success: !!first?.success,
        message: first?.message || batchResult.message
    };
}

/**
 * Applies OOXML reconciliation to multiple paragraphs with batched `context.sync()` calls.
 *
 * @param {Array<{ paragraph: Word.Paragraph, newText: string }>} edits - Paragraph edits
 * @param {Word.RequestContext} context - Word API context
 * @param {Object} options - Options
 * @param {boolean} [options.generateRedlines=true] - Whether to generate track changes
 * @param {string} [options.author='Gemini AI'] - Author for track changes
 * @param {boolean} [options.disableNativeTracking=false] - Temporarily disable native Word tracking while inserting OOXML
 * @param {Word.ChangeTrackingMode|null} [options.nativeTrackingMode=null] - Preloaded native tracking mode (avoids extra load/sync)
 * @returns {Promise<{ success: boolean, message: string, results: Array<{ index: number, success: boolean, changed?: boolean, message: string }> }>}
 */
export async function applyReconciliationToParagraphBatch(edits, context, options = {}) {
    const {
        generateRedlines = true,
        author = 'Gemini AI',
        disableNativeTracking = false,
        nativeTrackingMode = null
    } = options;

    if (!Array.isArray(edits) || edits.length === 0) {
        return {
            success: true,
            message: 'No paragraph edits supplied',
            results: []
        };
    }

    try {
        const pipeline = new ReconciliationPipeline({
            generateRedlines,
            author,
            validateOutput: true
        });

        // Batch-read all OOXML payloads first (1 sync).
        const ooxmlResults = edits.map(edit => edit.paragraph.getOoxml());
        await context.sync();

        const results = [];
        let queuedInsertions = 0;
        let trackingModeToRestore = null;
        let shouldDisableNativeTracking = false;

        if (disableNativeTracking) {
            if (nativeTrackingMode !== null && nativeTrackingMode !== undefined) {
                trackingModeToRestore = nativeTrackingMode;
            } else {
                context.document.load('changeTrackingMode');
                await context.sync();
                trackingModeToRestore = context.document.changeTrackingMode;
            }
            shouldDisableNativeTracking = trackingModeToRestore !== Word.ChangeTrackingMode.off;
        }

        for (let index = 0; index < edits.length; index++) {
            const edit = edits[index];
            try {
                const originalOoxml = ooxmlResults[index]?.value || '';
                log('[Integration] Got paragraph OOXML, length:', originalOoxml.length);

                if (!originalOoxml) {
                    results.push({
                        index,
                        success: false,
                        changed: false,
                        message: 'Empty OOXML payload'
                    });
                    continue;
                }

                const reconcileResult = await pipeline.execute(originalOoxml, edit.newText);
                if (!reconcileResult.isValid && reconcileResult.warnings.length > 0) {
                    warn('[Integration] Pipeline warnings:', reconcileResult.warnings);
                }

                if (!reconcileResult.ooxml || reconcileResult.ooxml === originalOoxml) {
                    results.push({
                        index,
                        success: true,
                        changed: false,
                        message: 'No changes detected'
                    });
                    continue;
                }

                const wrappedOoxml = pipeline.wrapForInsertion(reconcileResult.ooxml);
                edit.paragraph.insertOoxml(wrappedOoxml, 'Replace');
                queuedInsertions++;

                results.push({
                    index,
                    success: true,
                    changed: true,
                    message: 'Changes queued'
                });
            } catch (itemError) {
                error('[Integration] Reconciliation failed for paragraph index', index, itemError);
                results.push({
                    index,
                    success: false,
                    changed: false,
                    message: `Reconciliation error: ${itemError.message}`
                });
            }
        }

        // Batch-write queued operations (tracking toggle + paragraph insertions).
        if (queuedInsertions > 0) {
            if (shouldDisableNativeTracking) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            }
            await context.sync();

            if (shouldDisableNativeTracking) {
                context.document.changeTrackingMode = trackingModeToRestore;
                await context.sync();
            }
        }

        log('[Integration] Successfully applied batched OOXML reconciliation');

        const successCount = results.filter(result => result.success).length;
        const changedCount = results.filter(result => result.changed).length;
        return {
            success: successCount === results.length,
            message: `Applied ${changedCount} change(s) across ${successCount}/${results.length} paragraph edit(s)`,
            results
        };
    } catch (errorObj) {
        error('[Integration] Batch reconciliation failed:', errorObj);
        return {
            success: false,
            message: `Reconciliation error: ${errorObj.message}`,
            results: edits.map((_, index) => ({
                index,
                success: false,
                message: `Reconciliation error: ${errorObj.message}`
            }))
        };
    }
}

/**
 * Checks if OOXML reconciliation should be used for a given change.
 * Currently returns false (disabled) until pipeline is validated.
 * 
 * Enable by setting USE_OOXML_RECONCILIATION = true after testing.
 * 
 * @param {Object} change - The change object from AI
 * @returns {boolean}
 */
export function shouldUseOoxmlReconciliation(change) {
    // Feature flag - disabled by default until validated
    const USE_OOXML_RECONCILIATION = false;

    if (!USE_OOXML_RECONCILIATION) {
        return false;
    }

    // Only use for edit_paragraph operations initially
    if (change.operation !== 'edit_paragraph') {
        return false;
    }

    // Check if the change involves complex content that benefits from OOXML
    const hasMarkdownFormatting = /\*\*|__|~~|\+\+/.test(change.newContent);

    return hasMarkdownFormatting;
}

/**
 * Gets the current author setting for track changes.
 * 
 * @returns {string}
 */
export function getAuthorForTracking() {
    // Try to get from settings, fallback to 'Gemini AI'
    try {
        const stored = localStorage.getItem('redlineAuthor');
        return stored || 'Gemini AI';
    } catch {
        return 'Gemini AI';
    }
}
