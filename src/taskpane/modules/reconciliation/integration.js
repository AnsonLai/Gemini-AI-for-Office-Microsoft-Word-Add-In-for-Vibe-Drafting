/**
 * OOXML Reconciliation Pipeline - Integration Helper
 * 
 * Provides a bridge between the Word Add-in's existing flow and 
 * the new OOXML-based reconciliation pipeline.
 */

import { ReconciliationPipeline } from './index.js';

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
    const { generateRedlines = true, author = 'AI' } = options;

    try {
        // Step 1: Get the paragraph's OOXML
        const ooxmlResult = paragraph.getOoxml();
        await context.sync();

        const originalOoxml = ooxmlResult.value;
        console.log('[Integration] Got paragraph OOXML, length:', originalOoxml.length);

        // Step 2: Run the reconciliation pipeline
        const pipeline = new ReconciliationPipeline({
            generateRedlines,
            author,
            validateOutput: true
        });

        const result = await pipeline.execute(originalOoxml, newText);

        if (!result.isValid && result.warnings.length > 0) {
            console.warn('[Integration] Pipeline warnings:', result.warnings);
        }

        // Step 3: Wrap for insertion if needed
        const wrappedOoxml = pipeline.wrapForInsertion(result.ooxml);

        // Step 4: Insert the reconciled OOXML
        paragraph.insertOoxml(wrappedOoxml, 'Replace');
        await context.sync();

        console.log('[Integration] Successfully applied OOXML reconciliation');

        return {
            success: true,
            message: 'Changes applied via OOXML reconciliation'
        };

    } catch (error) {
        console.error('[Integration] Reconciliation failed:', error);
        return {
            success: false,
            message: `Reconciliation error: ${error.message}`
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
    // Try to get from settings, fallback to 'AI'
    try {
        const stored = localStorage.getItem('gemini_author_name');
        return stored || 'AI';
    } catch {
        return 'AI';
    }
}
