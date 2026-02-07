import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

export const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

export const REL_OFFICE_DOCUMENT_TYPE =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
export const REL_COMMENTS_TYPE =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const REL_NUMBERING_TYPE =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

export const CONTENT_TYPE_DOCUMENT =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
export const CONTENT_TYPE_COMMENTS =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const CONTENT_TYPE_NUMBERING =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

const parser = new DOMParser();
const serializer = new XMLSerializer();

/**
 * @param {string} xmlText
 * @param {string} label
 * @returns {Document}
 */
export function parseXmlStrict(xmlText, label = 'xml') {
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = firstByTag(xmlDoc, 'parsererror');
    if (parseError) {
        const details = parseError.textContent || 'Unknown parser error';
        throw new Error(`Invalid XML (${label}): ${details}`);
    }
    return xmlDoc;
}

/**
 * @param {Node} node
 * @returns {string}
 */
export function serializeXml(node) {
    return serializer.serializeToString(node);
}

/**
 * @param {NodeList|ArrayLike<any>|null|undefined} list
 * @returns {any[]}
 */
export function toArray(list) {
    if (!list) return [];
    const result = [];
    for (let i = 0; i < list.length; i += 1) {
        result.push(list.item ? list.item(i) : list[i]);
    }
    return result;
}

/**
 * @param {Node|null|undefined} node
 * @returns {string}
 */
export function getLocalName(node) {
    if (!node) return '';
    return node.localName || String(node.nodeName || '').replace(/^.*:/, '');
}

/**
 * @param {Node|null|undefined} node
 * @param {string} localName
 * @returns {boolean}
 */
export function isElement(node, localName = '') {
    if (!node || node.nodeType !== 1) return false;
    if (!localName) return true;
    return getLocalName(node) === localName;
}

/**
 * @param {ParentNode} node
 * @param {string} localName
 * @returns {Element[]}
 */
export function allByTag(node, localName) {
    return toArray(node.getElementsByTagNameNS('*', localName));
}

/**
 * @param {ParentNode} node
 * @param {string} localName
 * @returns {Element|null}
 */
export function firstByTag(node, localName) {
    const entries = allByTag(node, localName);
    return entries.length > 0 ? entries[0] : null;
}

/**
 * @param {Element} body
 * @returns {Element[]}
 */
export function directChildElements(body) {
    const result = [];
    for (const child of toArray(body.childNodes)) {
        if (child.nodeType === 1) {
            result.push(child);
        }
    }
    return result;
}

/**
 * @param {Document} doc
 * @returns {Element|null}
 */
export function getBody(doc) {
    return firstByTag(doc, 'body');
}

/**
 * Ensures `w:sectPr` is the last direct child in `w:body`.
 *
 * @param {Document} doc
 * @returns {Document}
 */
export function normalizeBodySectionOrder(doc) {
    const body = getBody(doc);
    if (!body) {
        throw new Error('Document has no w:body element');
    }

    const direct = directChildElements(body);
    const sectNodes = direct.filter(node => getLocalName(node) === 'sectPr');

    let sectPr = sectNodes.length > 0 ? sectNodes[0] : null;
    if (!sectPr) {
        sectPr = doc.createElementNS(NS_W, 'w:sectPr');
        body.appendChild(sectPr);
    }

    // Remove duplicate sectPr nodes if they exist.
    for (let i = 1; i < sectNodes.length; i += 1) {
        const node = sectNodes[i];
        if (node.parentNode) {
            node.parentNode.removeChild(node);
        }
    }

    // Move sectPr to the end if needed.
    if (body.lastChild !== sectPr) {
        body.appendChild(sectPr);
    }

    return doc;
}

/**
 * @param {string} value
 * @returns {string}
 */
export function escapeXml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * @param {Element} paragraph
 * @returns {string}
 */
export function paragraphText(paragraph) {
    const textNodes = allByTag(paragraph, 't');
    let text = '';
    for (const node of textNodes) {
        text += node.textContent || '';
    }
    return text;
}

