/**
 * Shared markdown list parsing for command adapters.
 *
 * Keeps command-layer list parsing aligned with reconciliation marker logic.
 */

import { matchListMarker, stripListMarker } from '../pipeline/list-markers.js';

/**
 * Parses markdown list-like content into structured items.
 *
 * Output shape is compatible with command-layer expectations from `parseMarkdownList`.
 *
 * @param {string} content - Raw markdown/text content
 * @returns {{ type: 'numbered'|'bullet'|'text', items: Array<{ type: 'numbered'|'bullet'|'text', level: number, text: string, marker?: string }> }|null}
 */
export function parseMarkdownListContent(content) {
    if (!content) return null;

    const normalized = String(content).trim();
    if (!normalized) return null;

    const lines = normalized.split('\n');
    const items = [];

    for (const line of lines) {
        if (!line.trim()) continue;

        const markerMatch = matchListMarker(line, { allowZeroSpaceAfterMarker: false });
        if (markerMatch) {
            const indent = markerMatch[1] || '';
            const marker = markerMatch[2].trim();
            const text = stripListMarker(line, { allowZeroSpaceAfterMarker: false }).trim();
            const level = Math.floor(indent.length / 2);
            const isBullet = /^[-*+\u2022]$/.test(marker);

            items.push({
                type: isBullet ? 'bullet' : 'numbered',
                level,
                text,
                marker
            });
            continue;
        }

        items.push({
            type: 'text',
            level: 0,
            text: line.trim()
        });
    }

    if (items.length === 0) return null;

    const hasNumbered = items.some(item => item.type === 'numbered');
    const hasBullet = items.some(item => item.type === 'bullet');

    return {
        type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
        items
    };
}

/**
 * Checks whether parsed list data includes at least one real list item.
 *
 * @param {{ items?: Array<{ type?: string }> }|null} parsedListData - Parsed list data
 * @returns {boolean}
 */
export function hasListItems(parsedListData) {
    if (!parsedListData || !Array.isArray(parsedListData.items)) return false;
    return parsedListData.items.some(item => item?.type === 'numbered' || item?.type === 'bullet');
}
