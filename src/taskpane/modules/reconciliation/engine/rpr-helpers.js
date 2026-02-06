/**
 * Run property (w:rPr) helper utilities.
 *
 * This module owns low-level formatting element operations, including
 * schema-order insertion, format extraction, and format add/remove transforms.
 */

/**
 * Canonical OOXML run-property schema ordering.
 * Shared by all rPr synchronizers.
 */
export const RPR_SCHEMA_ORDER = [
    'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
    'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof',
    'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern',
    'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd',
    'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout',
    'w:specVanish', 'w:oMath'
];

/**
 * Inserts an rPr child node in schema order.
 *
 * @param {Element} rPr - Run properties element
 * @param {Element} el - Child element to insert
 */
export function insertRPrChildInOrder(rPr, el) {
    const myIndex = RPR_SCHEMA_ORDER.indexOf(el.nodeName);
    const myPriority = myIndex === -1 ? 999 : myIndex;

    let inserted = false;
    for (const child of Array.from(rPr.childNodes)) {
        if (child.nodeType !== 1) continue;
        const childIndex = RPR_SCHEMA_ORDER.indexOf(child.nodeName);
        const childPriority = childIndex === -1 ? 999 : childIndex;
        if (childPriority > myPriority) {
            rPr.insertBefore(el, child);
            inserted = true;
            break;
        }
    }
    if (!inserted) rPr.appendChild(el);
}

/**
 * Shared override routine used by add/remove format transforms.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties target
 * @param {{bold?: boolean, italic?: boolean, underline?: boolean, strikethrough?: boolean}} formatFlags - Flags to apply
 * @param {'remove'|'add'} mode - Override mode
 */
function _applyOverrides(xmlDoc, rPr, formatFlags, mode) {
    if (!rPr || !formatFlags) return;

    const applyBold = !!formatFlags.bold;
    const applyItalic = !!formatFlags.italic;
    const applyUnderline = !!formatFlags.underline;
    const applyStrike = !!formatFlags.strikethrough;

    const removalSet = new Set();
    if (applyBold) {
        removalSet.add('w:b');
        removalSet.add('w:bCs');
    }
    if (applyItalic) {
        removalSet.add('w:i');
        removalSet.add('w:iCs');
    }
    if (applyUnderline) removalSet.add('w:u');
    if (applyStrike) removalSet.add('w:strike');

    if (removalSet.size > 0) {
        const toRemove = [];
        for (const child of Array.from(rPr.childNodes)) {
            if (removalSet.has(child.nodeName)) {
                toRemove.push(child);
            }
        }
        for (const el of toRemove) {
            rPr.removeChild(el);
        }
    }

    if (applyBold) {
        const b = xmlDoc.createElement('w:b');
        b.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, b);

        const bCs = xmlDoc.createElement('w:bCs');
        bCs.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, bCs);
    }
    if (applyItalic) {
        const i = xmlDoc.createElement('w:i');
        i.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, i);

        const iCs = xmlDoc.createElement('w:iCs');
        iCs.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, iCs);
    }
    if (applyUnderline) {
        const u = xmlDoc.createElement('w:u');
        u.setAttribute('w:val', mode === 'add' ? 'single' : 'none');
        insertRPrChildInOrder(rPr, u);
    }
    if (applyStrike) {
        const strike = xmlDoc.createElement('w:strike');
        strike.setAttribute('w:val', mode === 'add' ? '1' : '0');
        insertRPrChildInOrder(rPr, strike);
    }
}

/**
 * Builds an rPr XML snippet that explicitly removes formatting while preserving other properties.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} originalRun - Source run
 * @param {Object} formatToRemove - Format flags to remove
 * @param {XMLSerializer} serializer - Serializer instance
 * @returns {string}
 */
export function buildOverrideRPrXml(xmlDoc, originalRun, formatToRemove, serializer) {
    const baseRPr = originalRun.getElementsByTagName('w:rPr')[0] || null;
    const rPr = baseRPr ? baseRPr.cloneNode(true) : xmlDoc.createElement('w:rPr');
    _applyOverrides(xmlDoc, rPr, formatToRemove, 'remove');

    let rPrXml = serializer.serializeToString(rPr);
    rPrXml = rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '');
    return rPrXml === '<w:rPr/>' ? '' : rPrXml;
}

/**
 * Removes formatting tags and adds explicit off overrides for the specified flags.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties
 * @param {Object} formatToRemove - Format flags to remove
 */
export function applyFormatOverridesToRPr(xmlDoc, rPr, formatToRemove) {
    _applyOverrides(xmlDoc, rPr, formatToRemove, 'remove');
}

/**
 * Adds formatting tags to rPr for the specified flags without forcing other flags off.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties
 * @param {Object} formatToAdd - Format flags to add
 */
export function applyFormatAdditionsToRPr(xmlDoc, rPr, formatToAdd) {
    _applyOverrides(xmlDoc, rPr, formatToAdd, 'add');
}

/**
 * Removes all w:rPrChange children from run properties.
 *
 * @param {Element} rPr - Run properties
 */
export function stripRPrChangeNodes(rPr) {
    if (!rPr) return;
    const toRemove = [];
    for (const child of Array.from(rPr.childNodes)) {
        if (child.nodeName === 'w:rPrChange') {
            toRemove.push(child);
        }
    }
    for (const el of toRemove) {
        rPr.removeChild(el);
    }
}

/**
 * Builds a cloned rPr with format additions applied.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} originalRun - Source run
 * @param {Object} formatToAdd - Format flags to add
 * @returns {Element}
 */
export function buildAddedFormatRPr(xmlDoc, originalRun, formatToAdd) {
    const baseRPr = originalRun.getElementsByTagName('w:rPr')[0] || null;
    const rPr = baseRPr ? baseRPr.cloneNode(true) : xmlDoc.createElement('w:rPr');
    stripRPrChangeNodes(rPr);
    applyFormatAdditionsToRPr(xmlDoc, rPr, formatToAdd);
    return rPr;
}

/**
 * Extracts format flags from a run properties element.
 *
 * @param {Element|null} rPr - Run properties element
 * @returns {{ bold: boolean, italic: boolean, underline: boolean, strikethrough: boolean, hasFormatting: boolean }}
 */
export function extractFormatFromRPr(rPr) {
    const format = { bold: false, italic: false, underline: false, strikethrough: false, hasFormatting: false };
    if (!rPr) return format;

    for (const child of Array.from(rPr.childNodes)) {
        if (child.nodeName === 'w:b') format.bold = true;
        if (child.nodeName === 'w:i') format.italic = true;
        if (child.nodeName === 'w:u') format.underline = true;
        if (child.nodeName === 'w:strike') format.strikethrough = true;

        if (child.nodeName === 'w:rStyle') {
            const styleRef = child.getAttribute('w:val');
            if (styleRef) {
                const lowerStyle = styleRef.toLowerCase();
                if (lowerStyle.includes('bold') || lowerStyle.includes('strong')) format.bold = true;
                if (lowerStyle.includes('italic') || lowerStyle.includes('emphasis')) format.italic = true;
                if (lowerStyle.includes('underline')) format.underline = true;
            }
        }
    }

    format.hasFormatting = format.bold || format.italic || format.underline || format.strikethrough;
    return format;
}
