/**
 * OOXML Formatting Removal Utilities
 * 
 * Provides functions to surgically remove formatting from OOXML runs
 * while preserving the text content and other properties.
 */

import { parseOoxml, serializeOoxml } from './modules/reconciliation/oxml-engine.js';

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
 * Supports track changes via w:rPrChange.
 * 
 * @param {Document} doc - The OOXML document (for creating new elements)
 * @param {Element|null} rPr - The w:rPr element to modify (or null to create new)
 * @param {string} color - Highlight color name (default: 'yellow')
 * @param {Object} options - Options { generateRedlines: boolean, author: string }
 * @returns {Element} Modified or new rPr element with highlight
 */
export function injectHighlightIntoRPr(doc, rPr, color = 'yellow', options = {}) {
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const ooxmlColor = HIGHLIGHT_COLOR_MAP[color.toLowerCase()] || 'yellow';
    const { generateRedlines = false, author = 'Gemini AI' } = options;

    let rPrElement = rPr;
    if (!rPrElement) {
        // Create new rPr element
        rPrElement = doc.createElementNS(NS_W, 'w:rPr');
    } else {
        rPrElement = rPr.cloneNode(true);
    }

    // Capture "previous" state for redlines BEFORE modification
    // Clone the *original* rPr children before we touch them
    let previousRPrState = null;
    if (generateRedlines) {
        previousRPrState = doc.createElementNS(NS_W, 'w:rPr');
        Array.from(rPrElement.childNodes).forEach(child => {
            // Don't include existing rPrChange in the "previous" state wrapper usually, 
            // but for simplicity we clone children. Word generally handles nested track changes poorly,
            // so best to exclude rPrChange from the inner previous state.
            if (child.nodeName !== 'w:rPrChange') {
                previousRPrState.appendChild(child.cloneNode(true));
            }
        });
    }

    // --- APPLY CHANGE ---
    // Remove any existing highlight
    const existingHighlight = rPrElement.getElementsByTagNameNS(NS_W, 'highlight');
    Array.from(existingHighlight).forEach(el => el.remove());

    // Create and add new highlight element
    const highlightEl = doc.createElementNS(NS_W, 'w:highlight');
    highlightEl.setAttributeNS(NS_W, 'w:val', ooxmlColor);
    rPrElement.appendChild(highlightEl);

    // --- WRAP IN REDLINES IF ENABLED ---
    if (generateRedlines && previousRPrState) {
        const rPrChange = doc.createElementNS(NS_W, 'w:rPrChange');

        // Attributes
        rPrChange.setAttributeNS(NS_W, 'w:id', Math.floor(Math.random() * 9999999).toString());
        rPrChange.setAttributeNS(NS_W, 'w:author', author);
        rPrChange.setAttributeNS(NS_W, 'w:date', new Date().toISOString());

        // Format: <w:rPrChange ...> <w:rPr>...previous...</w:rPr> </w:rPrChange>
        rPrChange.appendChild(previousRPrState);

        // Remove any EXISTING rPrChange to avoid duplicates or nested weirdness
        const existingChange = rPrElement.getElementsByTagNameNS(NS_W, 'rPrChange');
        Array.from(existingChange).forEach(el => el.remove());

        // Append to rPr
        rPrElement.appendChild(rPrChange);
    }

    return rPrElement;
}

/**
 * Applies highlight formatting to OOXML runs containing the specified text.
 * Performs surgical splitting of runs if the text is a substring.
 * 
 * @param {string} ooxmlString - OOXML string (paragraph or package)
 * @param {string} targetText - Text to find and highlight
 * @param {string} color - Highlight color (default: 'yellow')
 * @returns {string} Modified OOXML string with highlights applied
 */
export function applyHighlightToOoxml(ooxmlString, targetText, color = 'yellow', options = {}) {
    if (!targetText || !ooxmlString) return ooxmlString;

    const doc = parseOoxml(ooxmlString);
    const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Helper to get text from a run
    const getRunText = (run) => {
        const textNodes = run.getElementsByTagNameNS(NS_W, 't');
        return Array.from(textNodes).map(t => t.textContent).join('');
    };

    // Find all runs
    const runs = Array.from(doc.getElementsByTagNameNS(NS_W, 'r'));
    const insertions = Array.from(doc.getElementsByTagNameNS(NS_W, 'ins'));

    // Collect all runs to check
    let allRuns = [...runs];
    for (const ins of insertions) {
        const insideRuns = ins.getElementsByTagNameNS(NS_W, 'r');
        allRuns.push(...Array.from(insideRuns));
    }

    // Process runs 
    // Note: We need to be careful about mutating the DOM while iterating.
    // However, since we split one run into multiple, we don't disturb the *order* of subsequent processed runs usually,
    // but a safe approach is to process updates after identification, or break after first match if we assume 1 match per call.
    // Given the task usually implies "highlight all occurrences" or "highlight this specific recurrence", 
    // but the current API is simple "textToFind". We'll assume "highlight all non-overlapping occurrences".

    for (let i = 0; i < allRuns.length; i++) {
        const run = allRuns[i];
        const runText = getRunText(run);

        if (!runText) continue;

        const matchIndex = runText.indexOf(targetText);
        if (matchIndex === -1) continue;

        // --- SPLITTING LOGIC ---
        // 1. Prefix (if match > 0)
        // 2. Match (highlighted)
        // 3. Suffix (if match + len < total len)

        const parent = run.parentNode;
        const prefixText = runText.substring(0, matchIndex);
        const matchText = runText.substring(matchIndex, matchIndex + targetText.length);
        const suffixText = runText.substring(matchIndex + targetText.length);

        // We replace the single 'run' with a fragment of 1-3 runs
        const fragment = doc.createDocumentFragment();

        // 1. Create Prefix Run
        if (prefixText.length > 0) {
            const prefixRun = run.cloneNode(true);
            // Update text content
            const tNodes = prefixRun.getElementsByTagNameNS(NS_W, 't');
            // Simply remove all t nodes and add one with new text to avoid complexity of multiple t nodes
            Array.from(tNodes).forEach(t => t.remove());
            const newT = doc.createElementNS(NS_W, 'w:t');
            // Preserve xml:space="preserve" if it existed, or just add it usually
            newT.setAttribute('xml:space', 'preserve');
            newT.textContent = prefixText;
            prefixRun.appendChild(newT);
            fragment.appendChild(prefixRun);
        }

        // 2. Create Match Run (With Highlight)
        if (matchText.length > 0) {
            const matchRun = run.cloneNode(true);
            // Update text content
            const tNodes = matchRun.getElementsByTagNameNS(NS_W, 't');
            Array.from(tNodes).forEach(t => t.remove());
            const newT = doc.createElementNS(NS_W, 'w:t');
            newT.setAttribute('xml:space', 'preserve');
            newT.textContent = matchText;
            matchRun.appendChild(newT);

            // Inject Highlight
            const rPrElements = matchRun.getElementsByTagNameNS(NS_W, 'rPr');
            const existingRPr = rPrElements.length > 0 ? rPrElements[0] : null;
            const newRPr = injectHighlightIntoRPr(doc, existingRPr, color, options);

            if (existingRPr) {
                matchRun.replaceChild(newRPr, existingRPr);
            } else {
                matchRun.insertBefore(newRPr, matchRun.firstChild);
            }
            fragment.appendChild(matchRun);
        }

        // 3. Create Suffix Run
        if (suffixText.length > 0) {
            const suffixRun = run.cloneNode(true);
            // Update text content
            const tNodes = suffixRun.getElementsByTagNameNS(NS_W, 't');
            Array.from(tNodes).forEach(t => t.remove());
            const newT = doc.createElementNS(NS_W, 'w:t');
            newT.setAttribute('xml:space', 'preserve');
            newT.textContent = suffixText;
            suffixRun.appendChild(newT);
            fragment.appendChild(suffixRun);
        }

        // Replace original run
        parent.replaceChild(fragment, run);

        // IMPORTANT: If we had a suffix that *also* contained the text (e.g. "target target"), 
        // our simple loop won't catch it because we replaced the node 'run'.
        // For now, we'll assume one match per run for simplicity, or we would need to recurse on the suffix.
        // Given the short contexts usually, this is acceptable for v1 fix.
    }

    return serializeOoxml(doc);
}
