/**
 * Pure routing plan builder used by command-layer Word adapters.
 *
 * This module must remain Word-agnostic.
 */

import { parseTable } from '../pipeline/content-analysis.js';
import { parseMarkdownListContent } from './list-parsing.js';

export const RoutePlanKind = Object.freeze({
    STRUCTURED_LIST_DIRECT: 'structured_list_direct',
    EMPTY_FORMATTED_TEXT: 'empty_formatted_text',
    EMPTY_HTML: 'empty_html',
    BLOCK_HTML: 'block_html',
    OOXML_ENGINE: 'ooxml_engine'
});

/**
 * Converts literal escape sequences (for example "\\n") into actual characters.
 *
 * @param {string} content - Raw model content
 * @returns {string}
 */
export function normalizeContentEscapesForRouting(content) {
    if (!content || typeof content !== 'string') return content || '';
    return content
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r');
}

/**
 * Builds a deterministic route plan for command-layer application.
 *
 * @param {Object} params - Plan input
 * @param {string} params.originalText - Current paragraph text
 * @param {string} params.newContent - Requested new content
 * @returns {Object}
 */
export function buildReconciliationPlan(params = {}) {
    const originalText = params.originalText || '';
    const normalizedContent = normalizeContentEscapesForRouting(params.newContent || '');

    const parsedListData = parseMarkdownListContent(normalizedContent) || { type: 'text', items: [] };
    const hasStructuredListContent = normalizedContent.includes('\n') && parsedListData.type !== 'text';
    if (hasStructuredListContent) {
        return {
            kind: RoutePlanKind.STRUCTURED_LIST_DIRECT,
            normalizedContent,
            parsedListData,
            flags: {
                hasStructuredListContent: true,
                isOriginalEmpty: isEmpty(originalText),
                hasInlineFormatting: hasInlineMarkdownFormatting(normalizedContent),
                hasBlockElements: hasBlockElements(normalizedContent),
                hasMarkdownTable: hasMarkdownTable(normalizedContent)
            }
        };
    }

    const isOriginalEmpty = isEmpty(originalText);
    if (isOriginalEmpty) {
        const hasInlineFormatting = hasInlineMarkdownFormatting(normalizedContent);
        if (hasInlineFormatting) {
            return {
                kind: RoutePlanKind.EMPTY_FORMATTED_TEXT,
                normalizedContent,
                parsedListData,
                flags: {
                    hasStructuredListContent: false,
                    isOriginalEmpty: true,
                    hasInlineFormatting: true,
                    hasBlockElements: hasBlockElements(normalizedContent),
                    hasMarkdownTable: hasMarkdownTable(normalizedContent)
                }
            };
        }

        return {
            kind: RoutePlanKind.EMPTY_HTML,
            normalizedContent,
            parsedListData,
            flags: {
                hasStructuredListContent: false,
                isOriginalEmpty: true,
                hasInlineFormatting: false,
                hasBlockElements: hasBlockElements(normalizedContent),
                hasMarkdownTable: hasMarkdownTable(normalizedContent)
            }
        };
    }

    const hasBlocks = hasBlockElements(normalizedContent);
    if (hasBlocks) {
        return {
            kind: RoutePlanKind.BLOCK_HTML,
            normalizedContent,
            parsedListData,
            flags: {
                hasStructuredListContent: false,
                isOriginalEmpty: false,
                hasInlineFormatting: hasInlineMarkdownFormatting(normalizedContent),
                hasBlockElements: true,
                hasMarkdownTable: hasMarkdownTable(normalizedContent)
            }
        };
    }

    return {
        kind: RoutePlanKind.OOXML_ENGINE,
        normalizedContent,
        parsedListData,
        flags: {
            hasStructuredListContent: false,
            isOriginalEmpty: false,
            hasInlineFormatting: hasInlineMarkdownFormatting(normalizedContent),
            hasBlockElements: false,
            hasMarkdownTable: hasMarkdownTable(normalizedContent)
        }
    };
}

function isEmpty(text) {
    return !text || text.trim().length === 0;
}

function hasMarkdownTable(content) {
    if (!content || !content.includes('|')) return false;
    const tableData = parseTable(content);
    return tableData.rows.length > 0 || tableData.headers.length > 0;
}

function hasInlineMarkdownFormatting(text) {
    if (!text) return false;
    return /(\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_|`.+?`|~~.+?~~|\+\+.+?\+\+)/.test(text);
}

function hasBlockElements(content) {
    if (!content) return false;

    const hasUnorderedList = /^[\s]*[-*+]\s+/m.test(content);
    const hasOrderedList = /^[\s]*\d+\.\s+/m.test(content);
    const hasOutlineList = /^[\s]*\d+\.\d+(?:\.\d+)*\.?\s+/m.test(content);
    const hasAlphaDotList = /^[\s]*[A-Za-z]\.\s+/m.test(content);
    const hasRomanDotList = /^[\s]*[ivxlcIVXLC]+\.\s+/m.test(content);
    const hasAlphaList = /^[\s]*\([a-z]\)\s+/m.test(content);
    const hasTable = /\|.*\|.*\n/.test(content);
    const hasHeading = /^#{1,9}\s/m.test(content);
    const hasMultipleLineBreaks = content.includes('\n\n');

    return hasUnorderedList
        || hasOrderedList
        || hasOutlineList
        || hasAlphaDotList
        || hasRomanDotList
        || hasAlphaList
        || hasTable
        || hasHeading
        || hasMultipleLineBreaks;
}
