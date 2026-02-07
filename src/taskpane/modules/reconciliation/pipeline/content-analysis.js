/**
 * Content analysis helpers for paragraph/list/table classification.
 */

import { ContentType } from '../core/types.js';
import { matchListMarker, stripListMarker } from './list-markers.js';

/**
 * Parses table from markdown-style table text.
 *
 * @param {string} text - Table text
 * @returns {{ headers: string[], rows: string[][], hasHeader: boolean }}
 */
export function parseTable(text) {
    const lines = text.split('\n').filter(line => line.trim().startsWith('|'));
    if (lines.length === 0) {
        return { headers: [], rows: [], hasHeader: false };
    }

    const dataLines = lines.filter(line => !line.includes('---'));
    const rows = dataLines.map(line =>
        line
            .split('|')
            .slice(1, -1)
            .map(cell => cell.trim())
    );

    return {
        headers: rows[0] || [],
        rows: rows.slice(1),
        hasHeader: lines.some(line => line.includes('---'))
    };
}

/**
 * Parses list items from markdown-style list text.
 *
 * @param {string} text - List text
 * @returns {Array<{ line: string, text: string, marker: string, indent: number, level: number, listType: 'bullet'|'numbered' }>}
 */
export function parseListItems(text) {
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    const items = [];

    lines.forEach(line => {
        const markerMatch = matchListMarker(line, { allowZeroSpaceAfterMarker: true });
        if (!markerMatch) return;

        const marker = markerMatch[2].trim();
        const indent = (line.match(/^(\s*)/)?.[1].length) || 0;
        const listType = /^[-*+•]/.test(marker) ? 'bullet' : 'numbered';
        const outlineDepth = (marker.match(/\./g) || []).length;
        const level = outlineDepth > 1 ? Math.min(8, outlineDepth - 1) : Math.min(8, Math.floor(indent / 2));

        items.push({
            line,
            text: stripListMarker(line, { allowZeroSpaceAfterMarker: true }),
            marker,
            indent,
            level,
            listType
        });
    });

    return items;
}

/**
 * Detects content type from text.
 *
 * @param {string} text - Text to classify
 * @returns {ContentType}
 */
export function detectContentType(text) {
    const normalized = (text || '').trim();
    if (!normalized) return ContentType.PARAGRAPH;

    const table = parseTable(normalized);
    if (table.headers.length > 0 || table.rows.length > 0) {
        return ContentType.TABLE;
    }

    const listItems = parseListItems(normalized);
    if (listItems.length > 0) {
        const hasBullet = listItems.some(item => item.listType === 'bullet');
        return hasBullet ? ContentType.BULLET_LIST : ContentType.NUMBERED_LIST;
    }

    return ContentType.PARAGRAPH;
}
