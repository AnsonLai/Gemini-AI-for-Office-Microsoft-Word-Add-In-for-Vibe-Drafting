/**
 * OOXML run/track-change builders.
 *
 * This module centralizes creation of `w:r`, `w:ins`, `w:del`, `w:rPrChange`,
 * and `w:pPrChange` elements used by surgical and reconstruction modes.
 */

import { RPR_SCHEMA_ORDER } from './rpr-helpers.js';

/**
 * Creates an insertion/deletion wrapper.
 *
 * @param {Document} xmlDoc - XML document
 * @param {'ins'|'del'} type - Wrapper type
 * @param {Element|null} run - Optional run to append
 * @param {string} author - Change author
 * @returns {Element}
 */
export function createTrackChange(xmlDoc, type, run, author) {
    const wrapper = xmlDoc.createElement(type === 'ins' ? 'w:ins' : 'w:del');
    wrapper.setAttribute('w:id', Math.floor(Math.random() * 90000 + 10000).toString());
    wrapper.setAttribute('w:author', author);
    wrapper.setAttribute('w:date', new Date().toISOString());
    if (run) {
        wrapper.appendChild(run);
    }
    return wrapper;
}

/**
 * Creates a text run with optional formatting.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text content
 * @param {Element|null} rPr - Run properties
 * @param {boolean} isDelete - Use `w:delText` instead of `w:t`
 * @returns {Element}
 */
export function createTextRun(xmlDoc, text, rPr, isDelete) {
    const run = xmlDoc.createElement('w:r');
    if (rPr) run.appendChild(rPr.cloneNode(true));

    const textEl = xmlDoc.createElement(isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Creates an array of runs with formatting applied from hints.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text to split and format
 * @param {Element|null} baseRPr - Base run properties
 * @param {Array} formatHints - Formatting hints
 * @param {number} baseOffset - Absolute base offset
 * @param {string} [author] - Change author
 * @param {boolean} [generateRedlines] - Whether to create rPrChange
 * @returns {Element[]}
 */
export function createFormattedRuns(xmlDoc, text, baseRPr, formatHints, baseOffset, author, generateRedlines) {
    if (!text) return [];

    const breaks = new Set([0, text.length]);
    for (const hint of formatHints) {
        const localStart = Math.max(0, hint.start - baseOffset);
        const localEnd = Math.min(text.length, hint.end - baseOffset);
        if (localStart >= 0 && localStart < text.length) breaks.add(localStart);
        if (localEnd > 0 && localEnd <= text.length) breaks.add(localEnd);
    }

    const sortedBreaks = Array.from(breaks).sort((a, b) => a - b);
    const runs = [];

    for (let i = 0; i < sortedBreaks.length - 1; i++) {
        const start = sortedBreaks[i];
        const end = sortedBreaks[i + 1];
        const segment = text.slice(start, end);
        if (!segment) continue;

        const segmentBaseOffset = baseOffset + start;
        const segmentEndOffset = baseOffset + end;

        const applicableHints = formatHints.filter(h =>
            h.start <= segmentBaseOffset && h.end >= segmentEndOffset
        );

        const combinedFormat = {};
        applicableHints.forEach(h => {
            if (h.format) Object.assign(combinedFormat, h.format);
        });

        const formattedRPr = injectFormattingToRPr(xmlDoc, baseRPr, combinedFormat, author, generateRedlines);
        runs.push(createTextRunWithRPrElement(xmlDoc, segment, formattedRPr, false));
    }

    return runs;
}

/**
 * Creates a text run with an existing rPr element (no clone).
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text content
 * @param {Element|null} rPrElement - Run properties element
 * @param {boolean} isDelete - Use `w:delText` instead of `w:t`
 * @returns {Element}
 */
export function createTextRunWithRPrElement(xmlDoc, text, rPrElement, isDelete) {
    const run = xmlDoc.createElement('w:r');
    if (rPrElement) run.appendChild(rPrElement);

    const textEl = xmlDoc.createElement(isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Creates a new rPr synchronized to the requested core formatting flags.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element|null} baseRPr - Base run properties
 * @param {Object|null} format - Format flags
 * @param {string} [author] - Change author
 * @param {boolean} [generateRedlines] - Whether to create rPrChange
 * @returns {Element}
 */
export function injectFormattingToRPr(xmlDoc, baseRPr, format, author, generateRedlines) {
    const rPr = xmlDoc.createElement('w:rPr');

    if (baseRPr) {
        Array.from(baseRPr.childNodes).forEach(child => {
            if (!['w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u', 'w:strike', 'w:rPrChange'].includes(child.nodeName)) {
                rPr.appendChild(child.cloneNode(true));
            }
        });
    }

    const activeFormat = format || { bold: false, italic: false, underline: false, strikethrough: false };

    if (author && generateRedlines) {
        createRPrChange(xmlDoc, rPr, author, baseRPr);
    }

    const syncElement = (tagName, isOn, valOn = null, valOff = '0') => {
        const el = xmlDoc.createElement(tagName);
        if (isOn) {
            if (valOn) el.setAttribute('w:val', valOn);
        } else if (valOff) {
            el.setAttribute('w:val', valOff);
        }

        const myIndex = RPR_SCHEMA_ORDER.indexOf(tagName);
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
    };

    syncElement('w:b', !!activeFormat.bold, '1', '0');
    syncElement('w:bCs', !!activeFormat.bold, '1', '0');
    syncElement('w:i', !!activeFormat.italic, '1', '0');
    syncElement('w:iCs', !!activeFormat.italic, '1', '0');
    syncElement('w:u', !!activeFormat.underline, 'single', 'none');
    syncElement('w:strike', !!activeFormat.strikethrough, '1', '0');

    return rPr;
}

/**
 * Creates and attaches a `w:rPrChange` snapshot.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Target run properties
 * @param {string} author - Change author
 * @param {string} dateStr - ISO date string
 * @param {Element} [sourceNode] - Optional source for previous state snapshot
 * @returns {Element}
 */
export function snapshotAndAttachRPrChange(xmlDoc, rPr, author, dateStr, sourceNode) {
    const rPrChange = xmlDoc.createElement('w:rPrChange');
    rPrChange.setAttribute('w:id', Math.floor(Math.random() * 90000 + 10000).toString());
    rPrChange.setAttribute('w:author', author);
    rPrChange.setAttribute('w:date', dateStr);

    const previousRPr = xmlDoc.createElement('w:rPr');
    const source = sourceNode || rPr;

    Array.from(source.childNodes).forEach(child => {
        if (child.nodeName !== 'w:rPrChange') {
            previousRPr.appendChild(child.cloneNode(true));
        }
    });

    rPrChange.appendChild(previousRPr);

    const existing = rPr.getElementsByTagName('w:rPrChange')[0];
    if (existing) {
        rPr.removeChild(existing);
    }

    rPr.appendChild(rPrChange);
    return rPrChange;
}

/**
 * Creates `w:rPrChange` for track formatting changes.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties target
 * @param {string} author - Change author
 * @param {Element} [previousRPrArg] - Optional explicit previous-state source
 * @returns {void}
 */
function createRPrChange(xmlDoc, rPr, author, previousRPrArg) {
    snapshotAndAttachRPrChange(xmlDoc, rPr, author, new Date().toISOString(), previousRPrArg || rPr);
}
