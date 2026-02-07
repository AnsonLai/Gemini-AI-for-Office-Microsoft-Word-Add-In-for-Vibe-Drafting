/**
 * Table-cell context helpers.
 *
 * Handles detection of table-wrapped paragraph OOXML and extraction of only
 * target paragraph content for safe `insertOoxml` replacement.
 */

import { getDocumentParagraphs } from './format-extraction.js';
import { log } from '../adapters/logger.js';
import { buildParagraphOnlyPackage } from '../services/package-builder.js';
import { getElementsByTag } from '../core/xml-query.js';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

/**
 * Detects whether the current XML is table-wrapped and resolves target paragraph context.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Source paragraph text
 * @param {Object} [options={}] - Detection options
 * @param {string|null} [options.targetParagraphId=null] - Preferred paragraph id (`w14:paraId`)
 * @returns {{
 *   hasTableWrapper: boolean,
 *   isTableCellParagraph: boolean,
 *   targetParagraph?: Element|null,
 *   paragraphs: Element[],
 *   paragraph: Element|null,
 *   tableElement: Element|null
 * }}
 */
export function detectTableCellContext(xmlDoc, originalText, options = {}) {
    const { targetParagraphId = null } = options;
    const tables = getElementsByTag(xmlDoc, 'w:tbl');
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

    // Most reliable selector when available (avoids ambiguous duplicate text matches in tables).
    if (targetParagraphId) {
        const normalizedTargetId = String(targetParagraphId).toUpperCase();
        targetParagraph = paragraphsInCells.find(p => {
            const paragraphId = getParagraphId(p);
            return paragraphId && paragraphId.toUpperCase() === normalizedTargetId;
        }) || null;

        if (targetParagraph) {
            log(`[OxmlEngine] Found target paragraph by paraId: "${targetParagraphId}"`);
        } else {
            log(`[OxmlEngine] paraId "${targetParagraphId}" not found in wrapper, falling back to text match`);
        }
    }

    if (originalText && originalText.trim()) {
        const normalizedTarget = originalText.trim();
        if (!targetParagraph) {
            for (const p of paragraphsInCells) {
                const textNodes = getElementsByTag(p, 'w:t');
                let paragraphText = '';
                for (const t of textNodes) {
                    paragraphText += t.textContent || '';
                }

                if (paragraphText.trim() === normalizedTarget) {
                    targetParagraph = p;
                    log(`[OxmlEngine] Found target paragraph by text match: "${normalizedTarget.substring(0, 30)}..."`);
                    break;
                }
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

/**
 * Reads the best available paragraph identity from OOXML attributes.
 *
 * @param {Element} paragraph - Paragraph element
 * @returns {string|null}
 */
function getParagraphId(paragraph) {
    if (!paragraph) return null;

    const namespacedId = typeof paragraph.getAttributeNS === 'function'
        ? paragraph.getAttributeNS(W14_NS, 'paraId')
        : null;
    if (namespacedId) return namespacedId;

    return paragraph.getAttribute('w14:paraId')
        || paragraph.getAttribute('w:paraId')
        || paragraph.getAttribute('paraId')
        || null;
}
