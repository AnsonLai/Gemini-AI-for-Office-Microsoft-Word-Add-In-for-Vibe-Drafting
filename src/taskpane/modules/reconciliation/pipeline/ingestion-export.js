/**
 * Word OOXML ingestion helpers for plain-text/markdown export.
 *
 * These helpers are intentionally conservative and only map obvious signals:
 * - Paragraph structure
 * - Heading styles
 * - List paragraph properties
 * - Run-level bold/italic formatting
 */

import { createParser } from '../adapters/xml-adapter.js';
import { NS_W } from '../core/types.js';
import { getXmlParseError } from '../core/xml-query.js';

function hasParserError(doc) {
    if (!doc || !doc.documentElement) return true;
    if (doc.documentElement.localName === 'parsererror') return true;
    return !!getXmlParseError(doc);
}

function parseWordOoxml(ooxml) {
    const source = typeof ooxml === 'string' ? ooxml : String(ooxml || '');
    if (!source.trim()) {
        return null;
    }

    try {
        const parser = createParser();
        const doc = parser.parseFromString(source, 'application/xml');
        if (hasParserError(doc)) {
            return null;
        }
        return doc;
    } catch {
        return null;
    }
}

function getWordParagraphs(doc) {
    if (!doc) return [];
    const namespaced = Array.from(doc.getElementsByTagNameNS(NS_W, 'p'));
    if (namespaced.length > 0) return namespaced;
    return Array.from(doc.getElementsByTagNameNS('*', 'p')).filter(node => node?.localName === 'p');
}

function getDirectWordChild(node, localName) {
    const children = Array.from(node?.childNodes || []);
    for (const child of children) {
        if (child?.nodeType !== 1) continue;
        if (child.namespaceURI !== NS_W) continue;
        if (child.localName === localName) return child;
    }
    return null;
}

function getWordDescendants(node, localName) {
    return Array.from(node?.getElementsByTagNameNS?.(NS_W, localName) || []);
}

function getWordAttribute(element, names) {
    if (!element) return '';
    for (const name of names) {
        const value = element.getAttribute(name);
        if (value != null && value !== '') return value;
    }
    return '';
}

function hasWordAncestorWithin(node, localName, boundary) {
    let cursor = node?.parentNode || null;
    while (cursor && cursor !== boundary) {
        if (cursor.nodeType === 1 && cursor.namespaceURI === NS_W && cursor.localName === localName) {
            return true;
        }
        cursor = cursor.parentNode;
    }
    return false;
}

function readRunText(run) {
    let text = '';
    for (const child of Array.from(run?.childNodes || [])) {
        if (!child || child.nodeType !== 1 || child.namespaceURI !== NS_W) continue;
        if (child.localName === 't') {
            text += child.textContent || '';
        } else if (child.localName === 'tab') {
            text += '\t';
        } else if (child.localName === 'br' || child.localName === 'cr') {
            text += '\n';
        } else if (child.localName === 'noBreakHyphen') {
            text += '\u2011';
        }
    }
    return text;
}

function getRunFormatting(run) {
    const rPr = getDirectWordChild(run, 'rPr');
    if (!rPr) return { bold: false, italic: false };
    return {
        bold: getWordDescendants(rPr, 'b').length > 0,
        italic: getWordDescendants(rPr, 'i').length > 0
    };
}

function collectParagraphSegments(paragraph) {
    const segments = [];
    const runs = Array.from(paragraph?.getElementsByTagNameNS?.(NS_W, 'r') || []);
    for (const run of runs) {
        if (hasWordAncestorWithin(run, 'del', paragraph)) continue;
        const text = readRunText(run);
        if (!text) continue;
        segments.push({
            text,
            ...getRunFormatting(run)
        });
    }
    return segments;
}

function normalizeInlineWhitespace(text) {
    const normalizedLines = String(text || '')
        .replace(/\r/g, '')
        .split('\n')
        .map(line => line.replace(/[ \t]+/g, ' ').trim());
    return normalizedLines.join('\n').trim();
}

function wrapRunMarkdown(text, format) {
    const raw = String(text || '');
    if (!raw) return '';
    if (!format?.bold && !format?.italic) return raw;

    const match = raw.match(/^(\s*)([\s\S]*?)(\s*)$/);
    if (!match) return raw;
    const leading = match[1] || '';
    const core = match[2] || '';
    const trailing = match[3] || '';
    if (!core.trim()) return raw;

    let wrapped = core;
    if (format.bold && format.italic) wrapped = `***${core}***`;
    else if (format.bold) wrapped = `**${core}**`;
    else if (format.italic) wrapped = `*${core}*`;
    return `${leading}${wrapped}${trailing}`;
}

function parseHeadingLevel(paragraph) {
    const pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) return null;

    const pStyle = getWordDescendants(pPr, 'pStyle')[0] || null;
    const styleVal = getWordAttribute(pStyle, ['w:val', 'val']);
    if (styleVal) {
        const match = styleVal.match(/^heading\s*([1-9])$/i);
        if (match) {
            const level = Number.parseInt(match[1], 10);
            if (Number.isInteger(level)) return Math.min(Math.max(level, 1), 6);
        }
    }

    const outlineLvl = getWordDescendants(pPr, 'outlineLvl')[0] || null;
    const outlineRaw = getWordAttribute(outlineLvl, ['w:val', 'val']);
    const outline = Number.parseInt(outlineRaw, 10);
    if (Number.isInteger(outline) && outline >= 0) {
        return Math.min(outline + 1, 6);
    }

    return null;
}

function parseListInfo(paragraph) {
    const pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) return null;

    const numPr = getWordDescendants(pPr, 'numPr')[0] || null;
    if (!numPr) return null;

    const ilvlEl = getWordDescendants(numPr, 'ilvl')[0] || null;
    const numIdEl = getWordDescendants(numPr, 'numId')[0] || null;
    const ilvl = Number.parseInt(getWordAttribute(ilvlEl, ['w:val', 'val']), 10);
    const numId = getWordAttribute(numIdEl, ['w:val', 'val']);

    const safeLevel = Number.isInteger(ilvl) && ilvl >= 0 ? ilvl : 0;
    const marker = numId === '1' ? '-' : '1.';
    return { level: safeLevel, marker };
}

function paragraphToPlainText(paragraph) {
    const text = collectParagraphSegments(paragraph).map(segment => segment.text).join('');
    return normalizeInlineWhitespace(text);
}

function paragraphToMarkdown(paragraph) {
    const segments = collectParagraphSegments(paragraph);
    const inline = segments.map(segment => wrapRunMarkdown(segment.text, segment)).join('');
    const normalizedInline = normalizeInlineWhitespace(inline);
    if (!normalizedInline) return '';

    const headingLevel = parseHeadingLevel(paragraph);
    if (headingLevel != null) {
        return `${'#'.repeat(headingLevel)} ${normalizedInline}`;
    }

    const list = parseListInfo(paragraph);
    if (list) {
        const indent = '  '.repeat(list.level);
        return `${indent}${list.marker} ${normalizedInline}`;
    }

    return normalizedInline;
}

/**
 * Ingests Word OOXML and returns readable plain text.
 *
 * @param {string} ooxml
 * @returns {string}
 */
export function ingestWordOoxmlToPlainText(ooxml) {
    const doc = parseWordOoxml(ooxml);
    if (!doc) return '';

    const paragraphs = getWordParagraphs(doc);
    if (paragraphs.length === 0) {
        const fallback = normalizeInlineWhitespace(doc.documentElement?.textContent || '');
        return fallback;
    }

    const lines = paragraphs.map(paragraphToPlainText);
    return lines.join('\n\n').trim();
}

/**
 * Ingests Word OOXML and returns basic markdown.
 *
 * @param {string} ooxml
 * @returns {string}
 */
export function ingestWordOoxmlToMarkdown(ooxml) {
    const doc = parseWordOoxml(ooxml);
    if (!doc) return '';

    const paragraphs = getWordParagraphs(doc);
    if (paragraphs.length === 0) {
        return '';
    }

    const lines = paragraphs.map(paragraphToMarkdown);
    return lines.join('\n\n').trim();
}
