/**
 * Reconstruction reconciliation mode orchestration.
 */

import { computeWordDiffs } from '../pipeline/diff-engine.js';
import { buildReconstructionMapping } from './reconstruction-mapper.js';
import { applyReconstructionDiffs } from './reconstruction-writer.js';

/**
 * Applies reconstruction mode reconciliation.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original text (kept for signature compatibility)
 * @param {string} modifiedText - Modified text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyReconstructionMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true) {
    const mapping = buildReconstructionMapping(xmlDoc, modifiedText);
    if (mapping.paragraphs.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const diffs = computeWordDiffs(mapping.originalFullText, mapping.processedModifiedText);

    return applyReconstructionDiffs(
        xmlDoc,
        diffs,
        mapping,
        serializer,
        author,
        formatHints,
        generateRedlines
    );
}
