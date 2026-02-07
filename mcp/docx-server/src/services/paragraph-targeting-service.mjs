import {
    NS_W14,
    allByTag,
    getLocalName,
    paragraphText,
    parseXmlStrict,
    serializeXml
} from './xml-utils.mjs';

/**
 * @param {Element} paragraph
 * @param {number} index
 * @returns {{ id: string, source: 'paraId'|'index', paraId: string|null }}
 */
function paragraphHandle(paragraph, index) {
    const paraId =
        paragraph.getAttributeNS?.(NS_W14, 'paraId')
        || paragraph.getAttribute('w14:paraId')
        || paragraph.getAttribute('w:paraId')
        || paragraph.getAttribute('paraId')
        || null;

    if (paraId) {
        return { id: String(paraId), source: 'paraId', paraId: String(paraId) };
    }

    return { id: `idx:${index + 1}`, source: 'index', paraId: null };
}

/**
 * @param {string} documentXml
 * @param {{ start?: number, limit?: number }} [windowing]
 */
export function listParagraphs(documentXml, windowing = {}) {
    const start = Math.max(0, Number(windowing.start ?? 0));
    const limit = Math.max(1, Number(windowing.limit ?? 50));

    const doc = parseXmlStrict(documentXml, 'word/document.xml');
    const paragraphs = allByTag(doc, 'p');

    const end = Math.min(paragraphs.length, start + limit);
    const entries = [];

    for (let i = start; i < end; i += 1) {
        const p = paragraphs[i];
        const handle = paragraphHandle(p, i);
        entries.push({
            id: handle.id,
            source: handle.source,
            index: i + 1,
            text: paragraphText(p)
        });
    }

    return {
        total: paragraphs.length,
        start,
        limit,
        items: entries
    };
}

/**
 * @param {string} documentXml
 * @param {string} paragraphId
 */
export function resolveParagraph(documentXml, paragraphId) {
    const doc = parseXmlStrict(documentXml, 'word/document.xml');
    const paragraphs = allByTag(doc, 'p');

    const asIndex = parseIndexId(paragraphId);
    if (asIndex !== null) {
        const idx = asIndex - 1;
        if (idx < 0 || idx >= paragraphs.length) {
            throw new Error(`Paragraph index out of range: ${paragraphId}`);
        }
        const paragraph = paragraphs[idx];
        const handle = paragraphHandle(paragraph, idx);
        return {
            doc,
            paragraph,
            index: idx + 1,
            id: handle.id,
            paraId: handle.paraId,
            source: handle.source
        };
    }

    for (let i = 0; i < paragraphs.length; i += 1) {
        const p = paragraphs[i];
        const handle = paragraphHandle(p, i);
        if (handle.id === paragraphId) {
            return {
                doc,
                paragraph: p,
                index: i + 1,
                id: handle.id,
                paraId: handle.paraId,
                source: handle.source
            };
        }
    }

    throw new Error(`Paragraph not found: ${paragraphId}`);
}

/**
 * @param {Document} doc
 * @param {Element} paragraph
 * @param {Element[]} replacementNodes
 * @returns {string}
 */
export function replaceParagraph(doc, paragraph, replacementNodes) {
    if (!paragraph.parentNode) {
        throw new Error('Cannot replace paragraph without parent node');
    }

    const parent = paragraph.parentNode;
    for (const node of replacementNodes) {
        if (getLocalName(node) === 'sectPr') {
            continue;
        }
        parent.insertBefore(doc.importNode(node, true), paragraph);
    }
    parent.removeChild(paragraph);

    return serializeXml(doc);
}

/**
 * @param {Element} paragraph
 * @returns {string}
 */
export function serializeParagraph(paragraph) {
    return serializeXml(paragraph);
}

/**
 * @param {string} value
 * @returns {number|null}
 */
function parseIndexId(value) {
    if (!value) return null;
    const byPrefix = String(value).match(/^idx:(\d+)$/i);
    if (byPrefix) {
        return Number(byPrefix[1]);
    }
    if (/^\d+$/.test(String(value))) {
        return Number(value);
    }
    return null;
}
