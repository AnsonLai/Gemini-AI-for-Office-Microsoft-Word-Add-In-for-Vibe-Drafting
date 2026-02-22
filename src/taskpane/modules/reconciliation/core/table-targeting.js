/**
 * Shared table-targeting helpers for per-paragraph redline callers.
 */

import {
    WORD_MAIN_NS,
    getParagraphText,
    isMarkdownTableText,
    findContainingWordElement,
    normalizeWhitespaceForTargeting
} from './paragraph-targeting.js';

function getDirectWordChildren(element, localName) {
    if (!element) return [];
    return Array.from(element.childNodes || []).filter(
        node =>
            node &&
            node.nodeType === 1 &&
            node.namespaceURI === WORD_MAIN_NS &&
            node.localName === localName
    );
}

function escapeMarkdownCell(text) {
    return String(text || '')
        .replace(/\|/g, '\\|')
        .replace(/\r?\n/g, '<br>');
}

function extractTableMatrix(tableElement) {
    const rowElements = getDirectWordChildren(tableElement, 'tr');
    const matrix = rowElements.map(row => {
        const cellElements = getDirectWordChildren(row, 'tc');
        return cellElements.map(cell => {
            const paragraphs = getDirectWordChildren(cell, 'p');
            if (paragraphs.length === 0) return normalizeWhitespaceForTargeting(getParagraphText(cell));
            const lines = paragraphs
                .map(p => normalizeWhitespaceForTargeting(getParagraphText(p)))
                .filter(Boolean);
            return lines.join('\n');
        });
    });

    const columnCount = matrix.reduce((max, row) => Math.max(max, row.length), 0);
    matrix.forEach(row => {
        while (row.length < columnCount) row.push('');
    });

    return { matrix, rowElements, columnCount };
}

function tableMatrixToMarkdown(matrix, columnCount) {
    if (!Array.isArray(matrix) || matrix.length === 0 || columnCount <= 0) return null;
    const normalized = matrix.map(row => {
        const copy = Array.isArray(row) ? row.slice(0, columnCount) : [];
        while (copy.length < columnCount) copy.push('');
        return copy;
    });

    const header = normalized[0];
    const separator = new Array(columnCount).fill('---');
    const bodyRows = normalized.slice(1);
    const toLine = row => `| ${row.map(cell => escapeMarkdownCell(cell)).join(' | ')} |`;

    return [toLine(header), toLine(separator), ...bodyRows.map(toLine)].join('\n');
}

function isSymmetricLabelRow(rowValues) {
    if (!Array.isArray(rowValues) || rowValues.length < 2) return false;
    const normalized = rowValues.map(value => normalizeWhitespaceForTargeting(value)).filter(Boolean);
    if (normalized.length < 2) return false;
    return normalized.every(value => value === normalized[0]);
}

/**
 * Heuristic detector for paragraphs likely belonging to a table-source block.
 *
 * @param {string} text - Paragraph text
 * @returns {boolean}
 */
export function isLikelyStructuredTableSourceParagraph(text) {
    const normalized = String(text || '').trim();
    if (!normalized) return false;
    if (/^and$/i.test(normalized)) return true;
    if (/^\[.*\]$/.test(normalized)) return true;
    if (/^\(.*\)$/.test(normalized)) return true;
    if (/:\s*$/.test(normalized)) return true;
    if (normalized.length <= 90 && !/[.!?]$/.test(normalized) && /[:\[\]()]/.test(normalized)) return true;
    if (/^[\[(]/.test(normalized)) return true;
    return false;
}

/**
 * Infers a contiguous paragraph block for table conversion starting from a paragraph.
 *
 * @param {Element|null} startParagraph - Starting w:p node
 * @param {Object} [options={}] - Inference options
 * @param {number} [options.maxScan=10] - Max sibling paragraphs to inspect
 * @param {(paragraph: Element) => string} [options.getParagraphText] - Optional text getter
 * @returns {Element[]|null}
 */
export function inferTableReplacementParagraphBlock(startParagraph, options = {}) {
    const maxScan = Number.isInteger(options?.maxScan) && options.maxScan > 0 ? options.maxScan : 10;
    const paragraphTextGetter = typeof options?.getParagraphText === 'function'
        ? options.getParagraphText
        : getParagraphText;

    if (!startParagraph || !startParagraph.parentNode) return null;

    const block = [startParagraph];
    let cursor = startParagraph.nextSibling;
    let scanned = 0;

    while (cursor && scanned < maxScan) {
        scanned += 1;
        const nextCursor = cursor.nextSibling;
        if (cursor.nodeType !== 1 || cursor.namespaceURI !== WORD_MAIN_NS || cursor.localName !== 'p') {
            cursor = nextCursor;
            continue;
        }

        const text = String(paragraphTextGetter(cursor) || '').trim();
        if (!text) {
            if (block.length > 1) break;
            cursor = nextCursor;
            continue;
        }

        if (!isLikelyStructuredTableSourceParagraph(text)) break;
        block.push(cursor);
        cursor = nextCursor;
    }

    return block.length > 1 ? block : null;
}

/**
 * Builds full-table markdown when a table-cell redline uses multiline text.
 *
 * Heuristic:
 * - If target paragraph is inside a table cell
 * - and `modifiedText` is multiline but not markdown-table syntax
 * - and first line matches current paragraph text
 * Then treat extra lines as row insertions in the target column.
 * For symmetric label rows (for example `Title:` in both signature columns),
 * inserted values are mirrored across columns.
 *
 * @param {Element} targetParagraph - Resolved target paragraph
 * @param {string} modifiedText - User/model modified text
 * @param {{
 *   tableElement?: Element|null,
 *   currentParagraphText?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} [options] - Optional context/log callbacks
 * @returns {string|null}
 */
export function synthesizeTableMarkdownFromMultilineCellEdit(targetParagraph, modifiedText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};

    const rawModified = String(modifiedText || '');
    if (!rawModified.includes('\n')) return null;
    if (isMarkdownTableText(rawModified)) return null;

    const lines = rawModified
        .split(/\r?\n/g)
        .map(line => line.trim())
        .filter(Boolean);
    if (lines.length < 2) return null;

    const tableElement = options.tableElement || findContainingWordElement(targetParagraph, 'tbl');
    const rowElement = findContainingWordElement(targetParagraph, 'tr');
    const cellElement = findContainingWordElement(targetParagraph, 'tc');
    if (!tableElement || !rowElement || !cellElement) return null;

    const currentParagraphText = normalizeWhitespaceForTargeting(
        options.currentParagraphText || getParagraphText(targetParagraph)
    );
    const firstLine = normalizeWhitespaceForTargeting(lines[0]);
    if (currentParagraphText && firstLine && firstLine !== currentParagraphText) {
        // Avoid rewriting full tables from ambiguous multiline content.
        onWarn('[Table] Multiline cell text did not anchor to original cell text; skipping table-row synthesis heuristic.');
        return null;
    }

    const { matrix, rowElements, columnCount } = extractTableMatrix(tableElement);
    if (matrix.length === 0 || columnCount === 0) return null;

    const rowIndex = rowElements.indexOf(rowElement);
    const cellElements = getDirectWordChildren(rowElement, 'tc');
    const colIndex = cellElements.indexOf(cellElement);
    if (rowIndex < 0 || colIndex < 0 || colIndex >= columnCount) return null;

    matrix[rowIndex][colIndex] = lines[0];
    const mirrorAcrossColumns = isSymmetricLabelRow(matrix[rowIndex]);
    if (mirrorAcrossColumns) {
        onInfo('[Table] Symmetric row detected; mirroring inserted row values across columns.');
    }

    for (let i = 1; i < lines.length; i++) {
        const insertIndex = rowIndex + i;
        const extraValue = lines[i];
        if (
            insertIndex < matrix.length &&
            !normalizeWhitespaceForTargeting(matrix[insertIndex][colIndex])
        ) {
            if (mirrorAcrossColumns) {
                for (let col = 0; col < columnCount; col++) {
                    if (!normalizeWhitespaceForTargeting(matrix[insertIndex][col])) {
                        matrix[insertIndex][col] = extraValue;
                    }
                }
            } else {
                matrix[insertIndex][colIndex] = extraValue;
            }
        } else {
            const newRow = new Array(columnCount).fill('');
            if (mirrorAcrossColumns) {
                for (let col = 0; col < columnCount; col++) newRow[col] = extraValue;
            } else {
                newRow[colIndex] = extraValue;
            }
            matrix.splice(Math.min(insertIndex, matrix.length), 0, newRow);
        }
    }

    const markdown = tableMatrixToMarkdown(matrix, columnCount);
    if (!markdown) return null;

    onInfo('[Table] Synthesized full markdown table from multiline cell edit for table-scope reconciliation.');
    return markdown;
}
