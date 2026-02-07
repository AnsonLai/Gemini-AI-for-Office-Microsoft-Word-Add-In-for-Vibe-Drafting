/**
 * OOXML Reconciliation Pipeline - Core Types
 * 
 * Data types and enums for the reconciliation system.
 */

// WordprocessingML namespace
export const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Diff operation types from word-level diffing
 */
export const DiffOp = Object.freeze({
    EQUAL: 'equal',
    DELETE: 'delete',
    INSERT: 'insert'
});

/**
 * Run types in the run model
 */
export const RunKind = Object.freeze({
    TEXT: 'run',
    DELETION: 'deletion',
    INSERTION: 'insertion',
    HYPERLINK: 'hyperlink',
    BOOKMARK: 'bookmark',
    FIELD: 'field',
    // Container delimiters for preserving hierarchy
    CONTAINER_START: 'container_start',
    CONTAINER_END: 'container_end',
    // Multi-paragraph support
    PARAGRAPH_START: 'paragraph_start'
});

/**
 * Container types that wrap runs
 */
export const ContainerKind = Object.freeze({
    SDT: 'sdt',                 // Content Control
    SMART_TAG: 'smartTag',      // Smart Tag
    CUSTOM_XML: 'customXml',    // Custom XML
    FIELD_COMPLEX: 'fldComplex' // Complex field (fldChar-based)
});

/**
 * Content types for block-level detection
 */
export const ContentType = Object.freeze({
    PARAGRAPH: 'paragraph',
    BULLET_LIST: 'bullet_list',
    NUMBERED_LIST: 'numbered_list',
    TABLE: 'table'
});

/**
 * Supported numbering formats
 */
export const NumberFormat = Object.freeze({
    DECIMAL: 'decimal',           // 1, 2, 3
    LOWER_ALPHA: 'lowerLetter',   // a, b, c
    UPPER_ALPHA: 'upperLetter',   // A, B, C
    LOWER_ROMAN: 'lowerRoman',    // i, ii, iii
    UPPER_ROMAN: 'upperRoman',    // I, II, III
    BULLET: 'bullet',             // •
    OUTLINE: 'outline'            // 1.1.2.3
});

/**
 * Numbering suffixes/formats
 */
export const NumberSuffix = Object.freeze({
    PERIOD: 'period',             // 1.
    PAREN_RIGHT: 'parenRight',    // 1)
    PAREN_BOTH: 'parenBoth',      // (1)
    NONE: 'none'
});

/**
 * @typedef {Object} RunEntry
 * @property {string} kind - RunKind value
 * @property {string} text - Text content of the run
 * @property {string} rPrXml - Serialized run properties (formatting)
 * @property {number} startOffset - Start offset in accepted text
 * @property {number} endOffset - End offset in accepted text
 * @property {string} [author] - Author for track changes
 * @property {string} [nodeXml] - Original XML for special elements
 */

/**
 * @typedef {Object} DiffOperation
 * @property {string} type - DiffOp value
 * @property {number} startOffset - Start offset in original text
 * @property {number} endOffset - End offset in original text
 * @property {string} text - Text content of the operation
 */

/**
 * @typedef {Object} FormatHint
 * @property {number} start - Start offset in clean text
 * @property {number} end - End offset in clean text
 * @property {Object} format - Format flags (bold, italic, etc.)
 */

/**
 * @typedef {Object} IngestionResult
 * @property {RunEntry[]} runModel - Array of run entries
 * @property {string} acceptedText - Reconstructed text from runs
 * @property {Element} pPr - Paragraph properties element
 */

/**
 * @typedef {Object} PreprocessResult
 * @property {string} cleanText - Text with markdown stripped
 * @property {FormatHint[]} formatHints - Position-based format information
 */

/**
 * @typedef {Object} ReconciliationResult
 * @property {string} ooxml - The reconciled OOXML output
 * @property {boolean} isValid - Whether validation passed
 * @property {string[]} warnings - Any warnings during processing
 */

/**
 * @typedef {Object} SerializationOptions
 * @property {string} [author='Gemini AI'] - Author for generated track changes
 * @property {boolean} [generateRedlines=true] - Toggle track-change wrappers
 * @property {string|null} [font=null] - Optional font override for generated runs
 */

/**
 * @typedef {Object} DocumentFragmentOptions
 * @property {boolean} [includeNumbering=false] - Include numbering relationship/part
 * @property {string|null} [numberingXml=null] - Custom numbering part payload
 * @property {boolean} [appendTrailingParagraph=true] - Append trailing blank paragraph
 */

/**
 * Escapes XML special characters
 * @param {string} str - String to escape
 * @returns {string} Escaped string
 */
export function escapeXml(str) {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Global revision ID counter for track changes
let revisionIdCounter = 1000;

/**
 * Gets the next unique revision ID for track changes
 * @returns {number} Unique revision ID
 */
export function getNextRevisionId() {
    return revisionIdCounter++;
}

/**
 * Returns the canonical ISO timestamp used for track-change metadata.
 *
 * @param {Date} [date] - Optional date source (for tests)
 * @returns {string}
 */
export function getRevisionTimestamp(date = new Date()) {
    return date.toISOString();
}

/**
 * Creates shared revision metadata for OOXML track-change tags.
 *
 * @param {string} [author='Gemini AI'] - Track-change author
 * @returns {{ id: number, author: string, date: string }}
 */
export function createRevisionMetadata(author = 'Gemini AI') {
    return {
        id: getNextRevisionId(),
        author,
        date: getRevisionTimestamp()
    };
}

/**
 * Resets the revision ID counter (for testing)
 * @param {number} [startValue=1000] - Value to reset to
 */
export function resetRevisionIdCounter(startValue = 1000) {
    revisionIdCounter = startValue;
}
