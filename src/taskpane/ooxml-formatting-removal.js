/**
 * OOXML Formatting Removal Utilities
 * 
 * Provides functions to surgically remove formatting from OOXML runs
 * while preserving the text content and other properties.
 */

import { parseOoxml, serializeOoxml } from './ooxml-utils.js';

/**
 * Removes specific formatting properties from a run properties (w:rPr) element.
 * This allows surgical removal of bold, italic, underline, color, etc. from OOXML.
 * 
 * @param {Element} rPr - The w:rPr element to modify
 * @param {string[]} formatTypes - Array of format types to remove: ['bold', 'italic', 'underline', 'strikethrough', 'color', 'highlight', 'fontSize', 'fontFamily', 'all']
 * @returns {Element|null} Modified rPr element, or null if all formatting removed
 */
export function removeFormattingFromRPr(rPr, formatTypes = ['all']) {
    if (!rPr) return null;

    const rPrClone = rPr.cloneNode(true);

    if (formatTypes.includes('all')) {
        // Remove all character formatting properties
        const toRemove = ['w:b', 'w:i', 'w:u', 'w:strike', 'w:dstrike', 'w:color',
            'w:sz', 'w:szCs', 'w:rFonts', 'w:highlight', 'w:vertAlign',
            'w:spacing', 'w:w', 'w:kern', 'w:position'];
        toRemove.forEach(tag => {
            // Handle both namespaced and non-namespaced versions
            const elements = rPrClone.querySelectorAll(`${tag}, ${tag.replace('w:', '')}`);
            elements.forEach(el => el.remove());
        });
    } else {
        // Remove specific properties
        const tagMap = {
            'bold': 'w:b',
            'italic': 'w:i',
            'underline': 'w:u',
            'strikethrough': 'w:strike',
            'doubleStrike': 'w:dstrike',
            'color': 'w:color',
            'highlight': 'w:highlight',
            'fontSize': 'w:sz',
            'fontSizeCs': 'w:szCs', // Complex script font size
            'fontFamily': 'w:rFonts',
            'superscript': 'w:vertAlign',
            'subscript': 'w:vertAlign'
        };

        formatTypes.forEach(type => {
            const tag = tagMap[type];
            if (tag) {
                // Handle both namespaced and non-namespaced versions
                const elements = rPrClone.querySelectorAll(`${tag}, ${tag.replace('w:', '')}`);
                elements.forEach(el => el.remove());
            }
        });
    }

    // Return null if rPr is now empty (no children)
    return rPrClone.children.length > 0 ? rPrClone : null;
}

/**
 * Applies formatting removal to OOXML containing the specified text.
 * Searches for text in runs and removes specified formatting properties.
 * 
 * @param {string} ooxmlString - OOXML string (paragraph or larger structure)
 * @param {string} targetText - Text to find and remove formatting from
 * @param {string[]} formatTypes - Array of format types to remove
 * @returns {string} Modified OOXML string
 */
export function applyFormattingRemovalToOoxml(ooxmlString, targetText, formatTypes) {
    if (!targetText || !ooxmlString) return ooxmlString;

    const doc = parseOoxml(ooxmlString);
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Find all text runs
    const runs = doc.getElementsByTagNameNS(NS_W, 'r');

    // Also handle runs inside w:ins (insertions)
    const insertions = doc.getElementsByTagNameNS(NS_W, 'ins');
    const allRuns = [...Array.from(runs)];

    for (const ins of insertions) {
        const insideRuns = ins.getElementsByTagNameNS(NS_W, 'r');
        allRuns.push(...Array.from(insideRuns));
    }

    for (const run of allRuns) {
        // Extract text from this run
        const textNodes = run.getElementsByTagNameNS(NS_W, 't');
        const runText = Array.from(textNodes).map(t => t.textContent).join('');

        // If this run contains the target text (or equals it)
        if (runText.includes(targetText) || runText === targetText) {
            // Find the rPr element
            const rPrElements = run.getElementsByTagNameNS(NS_W, 'rPr');

            if (rPrElements.length > 0) {
                const rPr = rPrElements[0];
                const newRPr = removeFormattingFromRPr(rPr, formatTypes);

                if (newRPr) {
                    // Replace with modified rPr
                    rPr.parentNode.replaceChild(newRPr, rPr);
                } else {
                    // Remove entire rPr if empty
                    rPr.remove();
                }
            }
        }
    }

    return serializeOoxml(doc);
}

// ==================== HIGHLIGHT INJECTION ====================

/**
 * Word API color names → OOXML w:highlight values
 */
const HIGHLIGHT_COLOR_MAP = {
    'yellow': 'yellow', 'green': 'green', 'cyan': 'cyan',
    'magenta': 'magenta', 'blue': 'blue', 'red': 'red',
    'darkblue': 'darkBlue', 'darkcyan': 'darkCyan',
    'darkgreen': 'darkGreen', 'darkmagenta': 'darkMagenta',
    'darkred': 'darkRed', 'darkyellow': 'darkYellow',
    'gray25': 'lightGray', 'gray50': 'darkGray',
    'black': 'black', 'white': 'white'
};

/**
 * Injects a highlight color into a run properties (w:rPr) element.
 * If rPr is null, creates a new rPr element with the highlight.
 * 
 * @param {Document} doc - The OOXML document (for creating new elements)
 * @param {Element|null} rPr - The w:rPr element to modify (or null to create new)
 * @param {string} color - Highlight color name (default: 'yellow')
 * @returns {Element} Modified or new rPr element with highlight
 */
export function injectHighlightIntoRPr(doc, rPr, color = 'yellow') {
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const ooxmlColor = HIGHLIGHT_COLOR_MAP[color.toLowerCase()] || 'yellow';

    let rPrElement = rPr;
    if (!rPrElement) {
        // Create new rPr element
        rPrElement = doc.createElementNS(NS_W, 'w:rPr');
    } else {
        rPrElement = rPr.cloneNode(true);
    }

    // Remove any existing highlight
    const existingHighlight = rPrElement.getElementsByTagNameNS(NS_W, 'highlight');
    Array.from(existingHighlight).forEach(el => el.remove());

    // Create and add new highlight element
    const highlightEl = doc.createElementNS(NS_W, 'w:highlight');
    highlightEl.setAttributeNS(NS_W, 'w:val', ooxmlColor);
    rPrElement.appendChild(highlightEl);

    return rPrElement;
}

/**
 * Applies highlight formatting to OOXML runs containing the specified text.
 * This is the inverse of applyFormattingRemovalToOoxml for highlights.
 * 
 * @param {string} ooxmlString - OOXML string (paragraph or package)
 * @param {string} targetText - Text to find and highlight
 * @param {string} color - Highlight color (default: 'yellow')
 * @returns {string} Modified OOXML string with highlights applied
 */
export function applyHighlightToOoxml(ooxmlString, targetText, color = 'yellow') {
    if (!targetText || !ooxmlString) return ooxmlString;

    const doc = parseOoxml(ooxmlString);
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Find all text runs (including inside w:ins)
    const runs = doc.getElementsByTagNameNS(NS_W, 'r');
    const insertions = doc.getElementsByTagNameNS(NS_W, 'ins');
    const allRuns = [...Array.from(runs)];

    for (const ins of insertions) {
        const insideRuns = ins.getElementsByTagNameNS(NS_W, 'r');
        allRuns.push(...Array.from(insideRuns));
    }

    for (const run of allRuns) {
        // Extract text from this run
        const textNodes = run.getElementsByTagNameNS(NS_W, 't');
        const runText = Array.from(textNodes).map(t => t.textContent).join('');

        // If this run contains the target text
        if (runText.includes(targetText) || runText === targetText) {
            // Get or create rPr
            const rPrElements = run.getElementsByTagNameNS(NS_W, 'rPr');
            const existingRPr = rPrElements.length > 0 ? rPrElements[0] : null;
            const newRPr = injectHighlightIntoRPr(doc, existingRPr, color);

            if (existingRPr) {
                run.replaceChild(newRPr, existingRPr);
            } else {
                // Insert rPr as first child of run
                run.insertBefore(newRPr, run.firstChild);
            }
        }
    }

    return serializeOoxml(doc);
}
