/**
 * Shared list markdown construction/parsing helpers for command adapters.
 */

/**
 * Infers numbering style from a list marker.
 *
 * @param {string} marker - Marker text
 * @returns {'decimal'|'lowerAlpha'|'upperAlpha'|'lowerRoman'|'upperRoman'}
 */
export function inferNumberingStyleFromMarker(marker) {
    const m = (marker || '').trim();
    if (!m) return 'decimal';
    if (/^\d+(?:\.\d+)*\.?$/.test(m) || /^\(\d+\)$/.test(m)) return 'decimal';
    if (/^[ivxlcdm]+\.$/.test(m)) return 'lowerRoman';
    if (/^[IVXLCDM]{2,}\.$/.test(m)) return 'upperRoman';
    if (/^[a-z]\.$/.test(m)) return 'lowerAlpha';
    if (/^[A-Z]\.$/.test(m)) return 'upperAlpha';
    return 'decimal';
}

/**
 * Builds list markdown from normalized item+level input.
 *
 * @param {Array<{ text: string, level: number }>} itemsWithLevels - Items with indentation levels
 * @param {'bullet'|'numbered'} listType - List kind
 * @param {'decimal'|'lowerAlpha'|'upperAlpha'|'lowerRoman'|'upperRoman'} numberingStyle - Number style for numbered lists
 * @returns {string}
 */
export function buildListMarkdown(itemsWithLevels, listType, numberingStyle) {
    const levelCounters = new Map();
    const lines = [];

    for (const item of itemsWithLevels) {
        const level = Math.max(0, Number(item?.level) || 0);
        for (const key of Array.from(levelCounters.keys())) {
            if (key > level) {
                levelCounters.delete(key);
            }
        }

        const nextCounter = (levelCounters.get(level) || 0) + 1;
        levelCounters.set(level, nextCounter);

        const marker = buildListMarker(nextCounter, listType, numberingStyle);
        const indent = ' '.repeat(level * 4);
        lines.push(`${indent}${marker} ${item.text || ''}`.trimEnd());
    }

    return lines.join('\n');
}

/**
 * Normalizes list item text/indentation into markdown-ready model.
 *
 * @param {Array<string>} rawItems - Raw tool items
 * @param {Object} [options={}] - Normalize options
 * @param {number} [options.indentSpaces=4] - Spaces per indent level
 * @returns {Array<{ text: string, level: number, removedMarker: string|null }>}
 */
export function normalizeListItemsWithLevels(rawItems, options = {}) {
    const indentSpaces = Math.max(1, Number(options.indentSpaces) || 4);
    const markersRegex = /^((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)/;

    return (rawItems || []).map((rawItem) => {
        const item = String(rawItem ?? '');
        const indentMatch = item.match(/^(\s*)/);
        const indentSize = indentMatch ? indentMatch[1].length : 0;
        const level = Math.floor(indentSize / indentSpaces);

        let stripped = item.trim();
        let removedMarker = null;
        const markerMatch = stripped.match(markersRegex);
        if (markerMatch) {
            removedMarker = markerMatch[1].trim() || null;
            stripped = stripped.replace(markersRegex, '');
        }

        return {
            text: stripped.trim(),
            level,
            removedMarker
        };
    });
}

function buildListMarker(counter, listType, numberingStyle) {
    if (listType === 'bullet') return '-';

    switch (numberingStyle) {
        case 'lowerAlpha':
            return `${toAlphaSequence(counter, false)}.`;
        case 'upperAlpha':
            return `${toAlphaSequence(counter, true)}.`;
        case 'lowerRoman':
            return `${toRoman(counter, false)}.`;
        case 'upperRoman':
            return `${toRoman(counter, true)}.`;
        case 'decimal':
        default:
            return `${counter}.`;
    }
}

function toAlphaSequence(value, upper = false) {
    let n = Math.max(1, Number(value) || 1);
    let out = '';
    while (n > 0) {
        n -= 1;
        out = String.fromCharCode(97 + (n % 26)) + out;
        n = Math.floor(n / 26);
    }
    return upper ? out.toUpperCase() : out;
}

function toRoman(value, upper = false) {
    let n = Math.max(1, Number(value) || 1);
    const romanPairs = [
        [1000, 'M'],
        [900, 'CM'],
        [500, 'D'],
        [400, 'CD'],
        [100, 'C'],
        [90, 'XC'],
        [50, 'L'],
        [40, 'XL'],
        [10, 'X'],
        [9, 'IX'],
        [5, 'V'],
        [4, 'IV'],
        [1, 'I']
    ];
    let out = '';
    for (const [num, sym] of romanPairs) {
        while (n >= num) {
            out += sym;
            n -= num;
        }
    }
    return upper ? out : out.toLowerCase();
}
