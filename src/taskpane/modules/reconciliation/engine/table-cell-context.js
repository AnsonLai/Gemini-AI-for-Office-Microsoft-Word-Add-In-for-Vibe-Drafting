/**
 * Table-cell context helpers.
 *
 * Handles detection of table-wrapped paragraph OOXML and extraction of only
 * target paragraph content for safe `insertOoxml` replacement.
 */

import { getDocumentParagraphs } from './format-extraction.js';
import { log } from '../adapters/logger.js';
import { buildParagraphOnlyPackage } from '../services/package-builder.js';

/**
 * Detects whether the current XML is table-wrapped and resolves target paragraph context.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Source paragraph text
 * @returns {{
 *   hasTableWrapper: boolean,
 *   isTableCellParagraph: boolean,
 *   targetParagraph?: Element|null,
 *   paragraphs: Element[],
 *   paragraph: Element|null,
 *   tableElement: Element|null
 * }}
 */
export function detectTableCellContext(xmlDoc, originalText) {
    const tables = xmlDoc.getElementsByTagName('w:tbl');
    if (tables.length === 0) {
        return { hasTableWrapper: false, isTableCellParagraph: false, paragraphs: [], paragraph: null, tableElement: null };
    }

    const allParagraphs = getDocumentParagraphs(xmlDoc);
    const paragraphsInCells = allParagraphs.filter(p => {
        let parent = p.parentNode;
        while (parent) {
            if (parent.nodeName === 'w:tc') return true;
            parent = parent.parentNode;
        }
        return false;
    });

    log(`[OxmlEngine] Table wrapper detected: ${tables.length} tables, ${paragraphsInCells.length} paragraphs in cells`);

    let targetParagraph = null;
    if (originalText && originalText.trim()) {
        const normalizedTarget = originalText.trim();
        for (const p of paragraphsInCells) {
            const textNodes = p.getElementsByTagName('w:t');
            let paragraphText = '';
            for (const t of Array.from(textNodes)) {
                paragraphText += t.textContent || '';
            }

            if (paragraphText.trim() === normalizedTarget) {
                targetParagraph = p;
                log(`[OxmlEngine] Found target paragraph by text match: "${normalizedTarget.substring(0, 30)}..."`);
                break;
            }
        }
    }

    return {
        hasTableWrapper: true,
        isTableCellParagraph: paragraphsInCells.length > 0,
        targetParagraph,
        paragraphs: paragraphsInCells,
        paragraph: targetParagraph || paragraphsInCells[0] || null,
        tableElement: tables[0]
    };
}

/**
 * Serializes one or more paragraphs without surrounding table wrappers.
 *
 * @param {Document} xmlDoc - XML document (unused, kept for signature compatibility)
 * @param {Element|Element[]} paragraphs - Paragraph or paragraph array
 * @param {XMLSerializer} serializer - Serializer instance
 * @returns {string}
 */
export function serializeParagraphOnly(xmlDoc, paragraphs, serializer) {
    const paragraphArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];

    let combinedXml = '';
    for (const p of paragraphArray) {
        if (!p) continue;
        let pXml = serializer.serializeToString(p);
        pXml = pXml.replace(/\s+xmlns:w="[^"]*"/g, '');
        pXml = pXml.replace(/\s+xmlns:r="[^"]*"/g, '');
        pXml = pXml.replace(/\s+xmlns:wp="[^"]*"/g, '');
        combinedXml += pXml;
    }

    log(`[OxmlEngine] Stripping table wrapper, serializing ${paragraphArray.length} paragraphs`);
    log(`[OxmlEngine] Paragraph XML preview: ${combinedXml.substring(0, 200)}...`);

    return wrapParagraphInPackage(combinedXml);
}

/**
 * Wraps paragraph XML in a complete OOXML package.
 *
 * @param {string} paragraphXml - Paragraph-only XML
 * @returns {string}
 */
export function wrapParagraphInPackage(paragraphXml) {
    return buildParagraphOnlyPackage(paragraphXml);
}
