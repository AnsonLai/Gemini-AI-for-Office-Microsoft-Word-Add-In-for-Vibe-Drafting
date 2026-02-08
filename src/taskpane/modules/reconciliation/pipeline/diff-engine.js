/**
 * OOXML Reconciliation Pipeline - Diff Engine
 * 
 * Word-level diffing with offset tracking for precise run splitting.
 */

import { diff_match_patch } from 'diff-match-patch';
import { DiffOp } from '../core/types.js';

const DMP = new diff_match_patch();

/**
 * Converts text into word tokens represented as unique characters.
 * This allows DMP to diff at word-level instead of character-level.
 * 
 * @param {string} text1 - First text to tokenize
 * @param {string} text2 - Second text to tokenize
 * @returns {{ chars1: string, chars2: string, wordArray: string[] }}
 */
export function wordsToChars(text1, text2) {
    const wordArray = [];
    const wordHash = new Map();

    function tokenize(text) {
        const tokens = [];
        const regex = /(\S+)(\s*)/g;
        let match;
        while ((match = regex.exec(text)) !== null) {
            if (match[1]) tokens.push(match[1]);
            if (match[2]) tokens.push(match[2]);
        }
        return tokens;
    }

    function mapTokensToChars(tokens) {
        let chars = '';
        for (const token of tokens) {
            if (wordHash.has(token)) {
                chars += String.fromCharCode(wordHash.get(token));
            } else {
                const charCode = wordArray.length;
                wordArray.push(token);
                wordHash.set(token, charCode);
                chars += String.fromCharCode(charCode);
            }
        }
        return chars;
    }

    const tokens1 = tokenize(text1);
    const tokens2 = tokenize(text2);

    return {
        chars1: mapTokensToChars(tokens1),
        chars2: mapTokensToChars(tokens2),
        wordArray
    };
}

/**
 * Converts character-encoded diffs back to actual word diffs.
 * 
 * @param {Array} diffs - DMP diff array with character codes
 * @param {string[]} wordArray - Array mapping char codes to words
 * @returns {Array} DMP-style diff array with actual words
 */
export function charsToWords(diffs, wordArray) {
    const wordDiffs = [];

    for (const [op, chars] of diffs) {
        const parts = [];
        for (let i = 0; i < chars.length; i++) {
            const charCode = chars.charCodeAt(i);
            if (charCode < wordArray.length) {
                parts.push(wordArray[charCode]);
            }
        }
        wordDiffs.push([op, parts.join('')]);
    }

    return wordDiffs;
}

/**
 * Computes word-level diff tuples using a shared diff engine instance.
 *
 * @param {string} originalText - Original text
 * @param {string} newText - New text
 * @param {{ cleanupSemantic?: boolean }} [options={}] - Diff options
 * @returns {Array<[number, string]>}
 */
export function computeWordDiffs(originalText, newText, options = {}) {
    if (originalText === newText) {
        return [[0, originalText]];
    }

    if (!originalText) {
        return [[1, newText]];
    }

    if (!newText) {
        return [[-1, originalText]];
    }

    const { cleanupSemantic = true } = options;

    const { chars1, chars2, wordArray } = wordsToChars(originalText, newText);
    const charDiffs = DMP.diff_main(chars1, chars2);
    if (cleanupSemantic) {
        DMP.diff_cleanupSemantic(charDiffs);
    }

    return charsToWords(charDiffs, wordArray);
}

/**
 * Computes word-level diff operations with offset tracking.
 * 
 * @param {string} originalText - Original text
 * @param {string} newText - New text
 * @param {{ cleanupSemantic?: boolean }} [options={}] - Diff options
 * @returns {import('../core/types.js').DiffOperation[]}
 */
export function computeWordLevelDiffOps(originalText, newText, options = {}) {
    // Handle edge cases
    if (originalText === newText) {
        return [{
            type: DiffOp.EQUAL,
            startOffset: 0,
            endOffset: originalText.length,
            text: originalText
        }];
    }

    if (!originalText) {
        return [{
            type: DiffOp.INSERT,
            startOffset: 0,
            endOffset: 0,
            text: newText
        }];
    }

    if (!newText) {
        return [{
            type: DiffOp.DELETE,
            startOffset: 0,
            endOffset: originalText.length,
            text: originalText
        }];
    }

    const wordDiffs = computeWordDiffs(originalText, newText, options);

    // Convert to operations with offsets
    const operations = [];
    let originalOffset = 0;

    for (const [op, text] of wordDiffs) {
        if (op === 0) { // EQUAL
            operations.push({
                type: DiffOp.EQUAL,
                startOffset: originalOffset,
                endOffset: originalOffset + text.length,
                text
            });
            originalOffset += text.length;
        } else if (op === -1) { // DELETE
            operations.push({
                type: DiffOp.DELETE,
                startOffset: originalOffset,
                endOffset: originalOffset + text.length,
                text
            });
            originalOffset += text.length;
        } else if (op === 1) { // INSERT
            operations.push({
                type: DiffOp.INSERT,
                startOffset: originalOffset,
                endOffset: originalOffset, // Insertions don't span original text
                text
            });
            // Don't advance originalOffset for insertions
        }
    }

    return operations;
}

/**
 * Collects all unique boundary offsets from diff operations.
 * 
 * @param {import('../core/types.js').DiffOperation[]} diffOps - Diff operations
 * @returns {Set<number>}
 */
export function collectDiffBoundaries(diffOps) {
    const boundaries = new Set();
    for (const op of diffOps) {
        boundaries.add(op.startOffset);
        boundaries.add(op.endOffset);
    }
    return boundaries;
}

