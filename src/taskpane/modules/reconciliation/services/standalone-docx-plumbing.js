/**
 * Standalone OOXML/Docx plumbing helpers shared by browser and Node hosts.
 */

import { createParser, createSerializer } from '../adapters/xml-adapter.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

const DOCUMENT_PATH = 'word/document.xml';
const NUMBERING_PATH = 'word/numbering.xml';
const COMMENTS_PATH = 'word/comments.xml';
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels';

export function parseXmlStrictStandalone(xmlText, label = 'xml') {
    const parser = createParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function isSectionPropertiesElement(node) {
    return !!node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'sectPr';
}

export function getBodyElementFromDocument(xmlDoc) {
    return xmlDoc.getElementsByTagNameNS('*', 'body')[0] || null;
}

function getDirectSectionProperties(body) {
    for (const child of Array.from(body.childNodes || [])) {
        if (isSectionPropertiesElement(child)) return child;
    }
    return null;
}

export function insertBodyElementBeforeSectPr(body, element) {
    const sectPr = getDirectSectionProperties(body);
    if (sectPr) {
        body.insertBefore(element, sectPr);
    } else {
        body.appendChild(element);
    }
}

export function normalizeBodySectionOrderStandalone(xmlDoc) {
    const body = getBodyElementFromDocument(xmlDoc);
    if (!body) return;
    const sectPr = getDirectSectionProperties(body);
    if (!sectPr) return;
    let cursor = sectPr.nextSibling;
    while (cursor) {
        const next = cursor.nextSibling;
        if (cursor.nodeType === 1) {
            body.insertBefore(cursor, sectPr);
        }
        cursor = next;
    }
}

/**
 * Flattens nested table-cell paragraphs (`w:tc > w:p > w:p`) by promoting
 * inner paragraphs to direct `w:tc` children.
 *
 * @param {Document} xmlDoc
 * @param {{ onInfo?: (message: string) => void }} [options]
 * @returns {number} number of nested paragraphs fixed
 */
export function sanitizeNestedParagraphsInTables(xmlDoc, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => {};
    const tcs = xmlDoc.getElementsByTagNameNS(NS_W, 'tc');
    let fixed = 0;
    for (const tc of Array.from(tcs)) {
        const outerParagraphs = Array.from(tc.childNodes || []).filter(
            node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p'
        );
        for (const outerParagraph of outerParagraphs) {
            const innerParagraphs = Array.from(outerParagraph.childNodes || []).filter(
                node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p'
            );
            for (const innerParagraph of innerParagraphs) {
                tc.insertBefore(innerParagraph, outerParagraph);
                fixed += 1;
            }
        }
    }
    if (fixed > 0) {
        onInfo(`[Sanitize] Fixed ${fixed} nested w:p element(s) in table cells`);
    }
    return fixed;
}

export function getPackagePartName(partElement) {
    return partElement.getAttribute('pkg:name') || partElement.getAttribute('name') || '';
}

function extractFromPackageXml(packageXml) {
    const parser = createParser();
    const serializer = createSerializer();
    const pkgDoc = parser.parseFromString(packageXml, 'application/xml');
    const parts = Array.from(pkgDoc.getElementsByTagNameNS('*', 'part'));
    const documentPart = parts.find(part => getPackagePartName(part) === '/word/document.xml');
    if (!documentPart) {
        throw new Error('Package output missing /word/document.xml part');
    }

    const xmlData = documentPart.getElementsByTagNameNS('*', 'xmlData')[0];
    if (!xmlData) {
        throw new Error('Package document part missing pkg:xmlData');
    }
    const documentNode = Array.from(xmlData.childNodes || []).find(node => node.nodeType === 1);
    if (!documentNode) {
        throw new Error('Package document part missing XML payload');
    }

    const body = documentNode.getElementsByTagNameNS('*', 'body')[0];
    const replacementNodes = body
        ? Array.from(body.childNodes || []).filter(node => node.nodeType === 1 && !isSectionPropertiesElement(node))
        : [documentNode];

    const numberingPart = parts.find(part => getPackagePartName(part) === '/word/numbering.xml');
    let numberingXml = null;
    if (numberingPart) {
        const numberingXmlData = numberingPart.getElementsByTagNameNS('*', 'xmlData')[0];
        const numberingNode = numberingXmlData
            ? Array.from(numberingXmlData.childNodes || []).find(node => node.nodeType === 1)
            : null;
        if (numberingNode) {
            numberingXml = serializer.serializeToString(numberingNode);
        }
    }

    return {
        replacementNodes,
        numberingXml,
        sourceType: 'package'
    };
}

/**
 * Extracts replacement nodes and optional numbering payload from engine output.
 *
 * @param {string} outputOxml
 * @returns {{ replacementNodes: Element[], numberingXml: string|null, sourceType: 'package'|'document'|'fragment' }}
 */
export function extractReplacementNodesFromOoxml(outputOxml) {
    if (typeof outputOxml !== 'string' || !outputOxml.trim()) {
        throw new Error('Reconciliation engine returned no OOXML payload for this operation');
    }

    if (outputOxml.includes('<pkg:package')) {
        return extractFromPackageXml(outputOxml);
    }

    if (outputOxml.includes('<w:document')) {
        const parser = createParser();
        const doc = parser.parseFromString(outputOxml, 'application/xml');
        const body = doc.getElementsByTagNameNS('*', 'body')[0];
        const replacementNodes = body
            ? Array.from(body.childNodes || []).filter(node => node.nodeType === 1 && !isSectionPropertiesElement(node))
            : Array.from(doc.childNodes || []).filter(node => node.nodeType === 1);
        return { replacementNodes, numberingXml: null, sourceType: 'document' };
    }

    const wrapped = `<root xmlns:w="${NS_W}">${outputOxml}</root>`;
    const parser = createParser();
    const fragmentDoc = parser.parseFromString(wrapped, 'application/xml');
    const replacementNodes = Array.from(fragmentDoc.documentElement.childNodes || []).filter(node => node.nodeType === 1);
    return { replacementNodes, numberingXml: null, sourceType: 'fragment' };
}

function upsertContentTypeOverride(ctDoc, partName, contentType) {
    const overrides = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override'));
    const hasOverride = overrides.some(
        override => (override.getAttribute('PartName') || '').toLowerCase() === String(partName).toLowerCase()
    );
    if (hasOverride) return false;

    const override = ctDoc.createElementNS(NS_CT, 'Override');
    override.setAttribute('PartName', partName);
    override.setAttribute('ContentType', contentType);
    ctDoc.documentElement.appendChild(override);
    return true;
}

function upsertDocumentRelationship(relsDoc, relType, target) {
    const relsRoot = relsDoc.getElementsByTagNameNS('*', 'Relationships')[0] || relsDoc.documentElement;
    const rels = Array.from(relsRoot.getElementsByTagNameNS('*', 'Relationship'));
    const hasRel = rels.some(rel => (rel.getAttribute('Type') || '') === relType);
    if (hasRel) return false;

    let maxId = 0;
    for (const rel of rels) {
        const idValue = rel.getAttribute('Id') || '';
        const idNum = Number.parseInt(idValue.replace(/^rId/i, ''), 10);
        if (Number.isFinite(idNum)) {
            maxId = Math.max(maxId, idNum);
        }
    }

    const rel = relsDoc.createElementNS(NS_RELS, 'Relationship');
    rel.setAttribute('Id', `rId${maxId + 1}`);
    rel.setAttribute('Type', relType);
    rel.setAttribute('Target', target);
    relsRoot.appendChild(rel);
    return true;
}

async function readZipText(zip, filePath) {
    const entry = zip.file(filePath);
    if (!entry) return null;
    return entry.async('string');
}

/**
 * Ensures numbering part + package metadata exist and merges numbering payloads.
 *
 * @param {any} zip
 * @param {string|string[]|null|undefined} numberingXmlList
 * @param {{
 *   mergeNumberingXml?: ((existingXml: string, incomingXml: string) => string),
 *   onInfo?: (message: string) => void
 * }} [options]
 */
export async function ensureNumberingArtifactsInZip(zip, numberingXmlList, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => {};
    const mergeNumberingXml = typeof options?.mergeNumberingXml === 'function'
        ? options.mergeNumberingXml
        : null;
    const incomingPayloads = (Array.isArray(numberingXmlList) ? numberingXmlList : [numberingXmlList]).filter(Boolean);
    if (incomingPayloads.length === 0) return;

    const existing = await readZipText(zip, NUMBERING_PATH);
    let mergedNumberingXml = existing || null;
    for (const incomingNumberingXml of incomingPayloads) {
        if (!mergedNumberingXml) {
            mergedNumberingXml = incomingNumberingXml;
            continue;
        }
        mergedNumberingXml = mergeNumberingXml
            ? mergeNumberingXml(mergedNumberingXml, incomingNumberingXml)
            : incomingNumberingXml;
    }

    if (!existing) {
        onInfo('[Demo] Adding numbering.xml');
    } else {
        onInfo('[Demo] Merging numbering.xml payload(s) into existing numbering definitions');
    }
    zip.file(NUMBERING_PATH, mergedNumberingXml);

    const parser = createParser();
    const serializer = createSerializer();

    const ctText = await readZipText(zip, CONTENT_TYPES_PATH);
    if (ctText) {
        const ctDoc = parser.parseFromString(ctText, 'application/xml');
        if (upsertContentTypeOverride(ctDoc, '/word/numbering.xml', NUMBERING_CONTENT_TYPE)) {
            zip.file(CONTENT_TYPES_PATH, serializer.serializeToString(ctDoc));
        }
    }

    const relsText = await readZipText(zip, DOCUMENT_RELS_PATH);
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        if (upsertDocumentRelationship(relsDoc, NUMBERING_REL_TYPE, 'numbering.xml')) {
            zip.file(DOCUMENT_RELS_PATH, serializer.serializeToString(relsDoc));
        }
    }
}

/**
 * Ensures comments part + package metadata exist and merges incoming comments.
 *
 * @param {any} zip
 * @param {string|null|undefined} commentsXml
 * @param {{ onInfo?: (message: string) => void }} [options]
 */
export async function ensureCommentsArtifactsInZip(zip, commentsXml, options = {}) {
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => {};
    if (!commentsXml) return;

    const parser = createParser();
    const serializer = createSerializer();
    const existingText = await readZipText(zip, COMMENTS_PATH);
    if (!existingText) {
        onInfo('[Demo] Adding comments.xml');
        zip.file(COMMENTS_PATH, commentsXml);
    } else {
        const existingDoc = parseXmlStrictStandalone(existingText, 'word/comments.xml (existing)');
        const incomingDoc = parseXmlStrictStandalone(commentsXml, 'word/comments.xml (incoming)');
        const existingRoot = existingDoc.documentElement;
        const existingIds = new Set(
            Array.from(existingRoot.getElementsByTagNameNS(NS_W, 'comment'))
                .map(comment => comment.getAttribute('w:id') || comment.getAttribute('id'))
                .filter(Boolean)
        );
        for (const incomingComment of Array.from(incomingDoc.documentElement.getElementsByTagNameNS(NS_W, 'comment'))) {
            const id = incomingComment.getAttribute('w:id') || incomingComment.getAttribute('id');
            if (id && existingIds.has(id)) {
                throw new Error(`Duplicate comment id: ${id}`);
            }
            existingRoot.appendChild(existingDoc.importNode(incomingComment, true));
        }
        zip.file(COMMENTS_PATH, serializer.serializeToString(existingDoc));
    }

    const ctText = await readZipText(zip, CONTENT_TYPES_PATH);
    if (ctText) {
        const ctDoc = parser.parseFromString(ctText, 'application/xml');
        if (upsertContentTypeOverride(ctDoc, '/word/comments.xml', COMMENTS_CONTENT_TYPE)) {
            zip.file(CONTENT_TYPES_PATH, serializer.serializeToString(ctDoc));
        }
    }

    const relsText = await readZipText(zip, DOCUMENT_RELS_PATH);
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        if (upsertDocumentRelationship(relsDoc, COMMENTS_REL_TYPE, 'comments.xml')) {
            zip.file(DOCUMENT_RELS_PATH, serializer.serializeToString(relsDoc));
        }
    }
}

/**
 * Validates core package integrity for document/comments/numbering artifacts.
 *
 * @param {any} zip
 */
export async function validateDocxPackage(zip) {
    const documentXml = await readZipText(zip, DOCUMENT_PATH);
    if (!documentXml) {
        throw new Error('Validation failed: missing word/document.xml');
    }

    const documentDoc = parseXmlStrictStandalone(documentXml, DOCUMENT_PATH);
    normalizeBodySectionOrderStandalone(documentDoc);
    const body = getBodyElementFromDocument(documentDoc);
    if (!body) {
        throw new Error('Validation failed: word/document.xml has no w:body');
    }

    const directBodyElements = Array.from(body.childNodes || []).filter(node => node.nodeType === 1);
    const sectPrIndexes = directBodyElements
        .map((node, index) => ({ node, index }))
        .filter(entry => isSectionPropertiesElement(entry.node))
        .map(entry => entry.index);

    if (sectPrIndexes.length > 1) {
        throw new Error('Validation failed: multiple body-level w:sectPr');
    }
    if (sectPrIndexes.length === 1 && sectPrIndexes[0] !== directBodyElements.length - 1) {
        throw new Error('Validation failed: w:sectPr not last');
    }

    const tcs = documentDoc.getElementsByTagNameNS(NS_W, 'tc');
    for (const tc of Array.from(tcs)) {
        for (const child of Array.from(tc.childNodes || []).filter(node => node.nodeType === 1)) {
            if (child.namespaceURI === NS_W && child.localName === 'p') {
                const hasNestedParagraph = Array.from(child.childNodes || []).some(
                    node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p'
                );
                if (hasNestedParagraph) {
                    throw new Error('Validation failed: nested w:p');
                }
            }
        }
    }

    const hasNumberingUsage = documentDoc.getElementsByTagNameNS(NS_W, 'numPr').length > 0;
    const hasCommentUsage =
        documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length > 0
        || documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeEnd').length > 0
        || documentDoc.getElementsByTagNameNS(NS_W, 'commentReference').length > 0;

    const numberingXml = await readZipText(zip, NUMBERING_PATH);
    const commentsXml = await readZipText(zip, COMMENTS_PATH);

    if (numberingXml) {
        parseXmlStrictStandalone(numberingXml, NUMBERING_PATH);
    } else if (hasNumberingUsage) {
        throw new Error('Validation failed: numbering used but part missing');
    }

    if (commentsXml) {
        parseXmlStrictStandalone(commentsXml, COMMENTS_PATH);
    } else if (hasCommentUsage) {
        throw new Error('Validation failed: comments used but part missing');
    }

    const ctXml = await readZipText(zip, CONTENT_TYPES_PATH);
    if (!ctXml) {
        throw new Error(`Validation failed: missing ${CONTENT_TYPES_PATH}`);
    }
    const ctDoc = parseXmlStrictStandalone(ctXml, CONTENT_TYPES_PATH);

    const relsXml = await readZipText(zip, DOCUMENT_RELS_PATH);
    if (!relsXml) {
        throw new Error(`Validation failed: missing ${DOCUMENT_RELS_PATH}`);
    }
    const relsDoc = parseXmlStrictStandalone(relsXml, DOCUMENT_RELS_PATH);

    if (numberingXml) {
        const hasNumberingContentType = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override')).some(override =>
            (override.getAttribute('PartName') || '').toLowerCase() === '/word/numbering.xml'
            && (override.getAttribute('ContentType') || '') === NUMBERING_CONTENT_TYPE
        );
        const hasNumberingRel = Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).some(rel =>
            (rel.getAttribute('Type') || '') === NUMBERING_REL_TYPE
        );
        if (!hasNumberingContentType) {
            throw new Error('Validation failed: numbering CT override missing');
        }
        if (!hasNumberingRel) {
            throw new Error('Validation failed: numbering rel missing');
        }
    }

    if (commentsXml) {
        const hasCommentsContentType = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override')).some(override =>
            (override.getAttribute('PartName') || '').toLowerCase() === '/word/comments.xml'
            && (override.getAttribute('ContentType') || '') === COMMENTS_CONTENT_TYPE
        );
        const hasCommentsRel = Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).some(rel =>
            (rel.getAttribute('Type') || '') === COMMENTS_REL_TYPE
        );
        if (!hasCommentsContentType) {
            throw new Error('Validation failed: comments CT override missing');
        }
        if (!hasCommentsRel) {
            throw new Error('Validation failed: comments rel missing');
        }
    }
}
