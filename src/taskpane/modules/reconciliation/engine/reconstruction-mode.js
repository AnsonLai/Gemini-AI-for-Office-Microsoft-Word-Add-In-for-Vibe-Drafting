/**
 * Reconstruction reconciliation mode orchestration.
 */

import { diff_match_patch } from 'diff-match-patch';
import { wordsToChars, charsToWords } from '../pipeline/diff-engine.js';
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

    const dmp = new diff_match_patch();
    const { chars1, chars2, wordArray } = wordsToChars(mapping.originalFullText, mapping.processedModifiedText);
    const charDiffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(charDiffs);
    const diffs = charsToWords(charDiffs, wordArray);

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
