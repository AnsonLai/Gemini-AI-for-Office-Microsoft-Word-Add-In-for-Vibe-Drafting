/**
 * OOXML Reconciliation Pipeline - Numbering Service
 * 
 * Manages list numbering to ensure continuation and consistent formatting.
 */

import { NS_W } from './types.js';

export class NumberingService {
    constructor() {
        this.contextMap = new Map(); // Cache for numIds found in the document
        this.nextNumId = 1000;
    }

    /**
     * Preserves an existing numId for a given format signature.
     * 
     * @param {string} signature - Format signature (e.g., 'bullet' or 'decimal')
     * @param {string} numId - Existing numId from Word
     */
    registerExistingNumId(signature, numId) {
        this.contextMap.set(signature, numId);
    }

    /**
     * Resolves the best numId to use for a requested list format.
     * 
     * @param {Object} formatConfig - Requested format (type, level)
     * @param {Object} existingContext - Context from adjacent paragraph
     * @returns {string} The numId to use
     */
    getOrCreateNumId(formatConfig, existingContext = null) {
        const requestedType = formatConfig.type || 'bullet';

        // Priority 1: Use existing context if it matches the requested type
        if (existingContext && existingContext.numId) {
            // If we have type info from context, check if it matches
            if (existingContext.type === requestedType || existingContext.type === 'unknown') {
                return existingContext.numId;
            }
        }

        // Priority 2: Use cached numId for this format
        if (this.contextMap.has(requestedType)) {
            return this.contextMap.get(requestedType);
        }

        // Priority 3: Fallback based on type
        // Note: In a real document, we should verify these exist in numbering.xml
        return requestedType === 'numbered' ? '2' : '1';
    }

    /**
     * Builds paragraph properties for a list item.
     * 
     * @param {string} numId - The numId
     * @param {number} ilvl - Indentation level
     * @returns {string} Serialized w:pPr XML
     */
    buildListPPr(numId, ilvl) {
        return `
            <w:pPr>
                <w:pStyle w:val="ListParagraph"/>
                <w:numPr>
                    <w:ilvl w:val="${ilvl}"/>
                    <w:numId w:val="${numId}"/>
                </w:numPr>
            </w:pPr>
        `;
    }
}
