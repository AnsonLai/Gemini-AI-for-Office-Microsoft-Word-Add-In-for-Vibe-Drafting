/**
 * Paragraph offset policy for text-model alignment.
 *
 * The reconciliation stack treats paragraph boundaries as a single newline
 * separator inserted between adjacent paragraphs (never after the last one).
 */

export const PARAGRAPH_BOUNDARY_TEXT = '\n';
export const PARAGRAPH_BOUNDARY_LENGTH = PARAGRAPH_BOUNDARY_TEXT.length;

/**
 * Returns true when a paragraph boundary separator should be inserted.
 *
 * @param {number} paragraphIndex - Zero-based paragraph index
 * @param {number} paragraphCount - Total paragraph count
 * @returns {boolean}
 */
export function hasParagraphBoundaryAfter(paragraphIndex, paragraphCount) {
    return paragraphIndex < paragraphCount - 1;
}

/**
 * Appends paragraph boundary text when policy requires one.
 *
 * @param {string} text - Source text
 * @param {number} paragraphIndex - Zero-based paragraph index
 * @param {number} paragraphCount - Total paragraph count
 * @returns {string}
 */
export function appendParagraphBoundary(text, paragraphIndex, paragraphCount) {
    if (!hasParagraphBoundaryAfter(paragraphIndex, paragraphCount)) {
        return text;
    }
    return text + PARAGRAPH_BOUNDARY_TEXT;
}

/**
 * Advances an offset by paragraph-boundary length when policy requires one.
 *
 * @param {number} offset - Current offset
 * @param {number} paragraphIndex - Zero-based paragraph index
 * @param {number} paragraphCount - Total paragraph count
 * @returns {number}
 */
export function advanceOffsetForParagraphBoundary(offset, paragraphIndex, paragraphCount) {
    if (!hasParagraphBoundaryAfter(paragraphIndex, paragraphCount)) {
        return offset;
    }
    return offset + PARAGRAPH_BOUNDARY_LENGTH;
}
