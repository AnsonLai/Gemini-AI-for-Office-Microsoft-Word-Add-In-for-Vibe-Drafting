/**
 * OOXML identity extraction helpers.
 */

/**
 * Extracts a paragraph identity token from OOXML (`w14:paraId` when present).
 *
 * @param {string} ooxml - OOXML payload
 * @returns {string|null}
 */
export function extractParagraphIdFromOoxml(ooxml) {
    if (!ooxml || typeof ooxml !== 'string') return null;
    const match = ooxml.match(/\b(?:w14:paraId|w:paraId|paraId)="([^"]+)"/i);
    return match ? match[1] : null;
}
