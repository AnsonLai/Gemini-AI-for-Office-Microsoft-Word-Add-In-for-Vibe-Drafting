/** Flat OPC source inspection and insertion-package assembly around public library operations. */
import {
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone,
    getDefaultAuthor,
    inspectDocumentParts,
    mergeNumberingXmlBySchemaOrder,
    computeDocumentPartsRevisionToken
} from '@ansonlai/docx-redline-js';
import { createParser, createSerializer } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { reconcileCommentSiblingParts } from '@ansonlai/docx-redline-js/services/comment-thread-parts.js';
import { getPartSpec } from '@ansonlai/docx-redline-js/services/package-parts.js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PART_TYPES = Object.fromEntries(
    ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible', 'numbering']
        .map(kind => getPartSpec(kind)).map(spec => [`/${spec.path}`, spec.contentType])
);
function parseXmlStrict(xmlText, label) {
    const parser = createParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown parse error'}`);
    }
    return xmlDoc;
}

function wrapParagraphNodesAsDocument(paragraphNodes) {
    const serializer = createSerializer();
    const bodyXml = (paragraphNodes || [])
        .map(node => serializer.serializeToString(node))
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`;
}

function extractBodyChildElements(xmlDoc) {
    if (!xmlDoc) return [];
    const body = xmlDoc.getElementsByTagNameNS(NS_W, 'body')[0] || xmlDoc.getElementsByTagNameNS('*', 'body')[0];
    if (!body) return [];

    return Array.from(body.childNodes || []).filter(
        node => node
            && node.nodeType === 1
            && !(node.namespaceURI === NS_W && node.localName === 'sectPr')
    );
}

function packagePart(packageDoc, name) {
    return Array.from(packageDoc.getElementsByTagNameNS(NS_PKG, 'part')).find(
        part => (part.getAttribute('pkg:name') || part.getAttribute('name')) === name
    ) || null;
}

function packagePartXml(packageDoc, name) {
    const xmlData = packagePart(packageDoc, name)?.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    const root = Array.from(xmlData?.childNodes || []).find(node => node.nodeType === 1);
    return root ? createSerializer().serializeToString(root) : null;
}

function replacePackagePartXml(packageDoc, name, xml) {
    if (!xml) return;
    let part = packagePart(packageDoc, name);
    if (!part) {
        part = packageDoc.createElementNS(NS_PKG, 'pkg:part');
        part.setAttribute('pkg:name', name);
        part.setAttribute('pkg:contentType', PART_TYPES[name]);
        packageDoc.documentElement.appendChild(part);
    }
    if (PART_TYPES[name]) part.setAttribute('pkg:contentType', PART_TYPES[name]);
    let xmlData = part.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    if (!xmlData) {
        xmlData = packageDoc.createElementNS(NS_PKG, 'pkg:xmlData');
        part.appendChild(xmlData);
    }
    while (xmlData.firstChild) xmlData.removeChild(xmlData.firstChild);
    xmlData.appendChild(packageDoc.importNode(parseXmlStrict(xml, name).documentElement, true));
}

function ensurePackageRelationship(packageDoc, kind) {
    const name = '/word/_rels/document.xml.rels';
    let relsXml = packagePartXml(packageDoc, name);
    if (!relsXml) relsXml = `<Relationships xmlns="${NS_REL}"/>`;
    const relsDoc = parseXmlStrict(relsXml, name);
    const relationships = Array.from(relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship'));
    const spec = getPartSpec(kind);
    const type = spec.relType;
    if (!relationships.some(rel => rel.getAttribute('Type') === type)) {
        const used = new Set(relationships.map(rel => rel.getAttribute('Id')));
        let next = 1;
        while (used.has(`rId${next}`)) next++;
        const rel = relsDoc.createElementNS(NS_REL, 'Relationship');
        rel.setAttribute('Id', `rId${next}`);
        rel.setAttribute('Type', type);
        rel.setAttribute('Target', spec.relTarget);
        relsDoc.documentElement.appendChild(rel);
    }
    // Relationship parts have a different content type from the Word XML parts.
    let part = packagePart(packageDoc, name);
    if (!part) {
        part = packageDoc.createElementNS(NS_PKG, 'pkg:part');
        part.setAttribute('pkg:name', name);
        part.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-package.relationships+xml');
        packageDoc.documentElement.appendChild(part);
    }
    let xmlData = part.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    if (!xmlData) {
        xmlData = packageDoc.createElementNS(NS_PKG, 'pkg:xmlData');
        part.appendChild(xmlData);
    }
    while (xmlData.firstChild) xmlData.removeChild(xmlData.firstChild);
    xmlData.appendChild(packageDoc.importNode(relsDoc.documentElement, true));
}

function readBatchSource(scopeOoxml) {
    if (scopeOoxml.includes('<pkg:package')) {
        const packageDoc = parseXmlStrict(scopeOoxml, 'Word scope package');
        const documentXml = packagePartXml(packageDoc, '/word/document.xml');
        if (!documentXml) throw new Error('Word scope package has no document.xml part');
        const parts = {
            documentXml,
            commentsXml: packagePartXml(packageDoc, '/word/comments.xml'),
            commentsExtendedXml: packagePartXml(packageDoc, '/word/commentsExtended.xml'),
            commentsIdsXml: packagePartXml(packageDoc, '/word/commentsIds.xml'),
            commentsExtensibleXml: packagePartXml(packageDoc, '/word/commentsExtensible.xml'),
            numberingXml: packagePartXml(packageDoc, '/word/numbering.xml'),
            stylesXml: packagePartXml(packageDoc, '/word/styles.xml')
        };
        // Bind every remaining transported dependency through the public token API.
        // Package assembly preserves these parts, so freshness must cover them too.
        const covered = new Set(['/word/document.xml', '/word/comments.xml', '/word/commentsExtended.xml', '/word/numbering.xml', '/word/styles.xml']);
        parts.additionalParts = Object.fromEntries(Array.from(packageDoc.getElementsByTagNameNS(NS_PKG, 'part'))
            .map(part => [part.getAttribute('pkg:name'), part])
            .filter(([name]) => name && !covered.has(name))
            .map(([name, part]) => [name, packagePartXml(packageDoc, name)
                || part.getElementsByTagNameNS(NS_PKG, 'binaryData')[0]?.textContent || '']));
        const inspection = inspectDocumentParts(parts);
        if (inspection.status === 'error') throw new Error(inspection.error?.message || 'Cannot inspect Word scope');
        return {
            ...parts,
            paragraphs: inspection.paragraphs,
            commentCount: inspection.comments?.length || 0,
            packageDoc
        };
    }
    const extracted = extractReplacementNodesFromOoxml(scopeOoxml);
    if (extracted.status === 'error' || !extracted.replacementNodes?.length) {
        throw new Error(extracted.error?.message || 'Word scope has no editable OOXML nodes');
    }
    const documentXml = extracted.sourceType === 'document'
        ? scopeOoxml
        : wrapParagraphNodesAsDocument(extracted.replacementNodes);
    const parts = { documentXml, numberingXml: extracted.numberingXml || null };
    const inspection = inspectDocumentParts(parts);
    if (inspection.status === 'error') throw new Error(inspection.error?.message || 'Cannot inspect Word scope');
    return {
        ...parts,
        paragraphs: inspection.paragraphs,
        commentCount: inspection.comments?.length || 0,
        packageDoc: null
    };
}

export function inspectWordSource(scopeOoxml, options = {}) {
    const source = readBatchSource(scopeOoxml);
    const parts = Object.fromEntries([
        'documentXml', 'commentsXml', 'commentsExtendedXml', 'commentsIdsXml',
        'commentsExtensibleXml', 'numberingXml', 'stylesXml'
    ].map(name => [name, source[name] || null]));
    parts.additionalParts = source.additionalParts || {};
    const inspection = inspectDocumentParts(parts, { ...options, revisionView: 'accepted' });
    if (inspection.status !== 'ok') throw Object.assign(new Error(inspection.error?.message || 'Canonical inspection failed.'), { code: inspection.error?.code || 'CANONICAL_CONTEXT_UNAVAILABLE' });
    return { parts, inspection };
}

/**
 * Parse one immutable OOXML source, build and execute a canonical operation batch,
 * and prepare the OOXML payload without reading or writing a host document.
 *
 * `operations` may be an array or a function receiving the inspected source.
 * The returned source omits its internal DOM package object and is safe to pass
 * between portable consumers.
 */
export async function prepareCanonicalBatch(scopeOoxml, operations, options = {}) {
    const source = readBatchSource(scopeOoxml);
    const publicSource = {
        documentXml: source.documentXml,
        paragraphs: source.paragraphs,
        commentsXml: source.commentsXml || null,
        commentsExtendedXml: source.commentsExtendedXml || null,
        commentsIdsXml: source.commentsIdsXml || null,
        commentsExtensibleXml: source.commentsExtensibleXml || null,
        numberingXml: source.numberingXml || null,
        stylesXml: source.stylesXml || null
    };
    publicSource.additionalParts = source.additionalParts || {};
    const batch = typeof operations === 'function'
        ? await operations(publicSource)
        : operations;
    if (!Array.isArray(batch)) throw new TypeError('Batch operations must be an array');
    if (batch.length === 0) {
        return {
            status: 'noop',
            source: publicSource,
            operations: batch,
            result: { status: 'ok', hasChanges: false, results: [], receipts: [] }
        };
    }

    const runner = typeof options.runner === 'function' ? options.runner : applyOperationsToDocumentXml;
    const runtimeContext = {
        commentsXml: source.commentsXml || null,
        commentsExtendedXml: source.commentsExtendedXml || null,
        numberingXml: source.numberingXml || null,
        stylesXml: source.stylesXml || null
    };
    const result = await runner(
        source.documentXml,
        batch,
        options.author || getDefaultAuthor(),
        runtimeContext,
        {
            atomic: true,
            structuredContent: true,
            pairReplacements: true,
            generateRedlines: options.generateRedlines !== false,
            existingRevisions: options.existingRevisions,
            sanitizeInput: options.sanitizeInput === true,
            strictTargets: options.strictTargets !== false,
            expectedRevision: options.expectedRevision === true
                ? await computeDocumentPartsRevisionToken({ documentXml: source.documentXml, ...runtimeContext })
                : options.expectedRevision,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    if (!result || result.status === 'error' || result.status === 'partial' || result.error || result.rolledBack) {
        return { status: 'refused', source: publicSource, operations: batch, result };
    }
    if (result.commentsXmlMode === 'merge'
        && typeof result.commentsXml === 'string'
        && result.commentsXml.trim()
        && source.commentCount > 0) {
        return {
            status: 'refused',
            source: publicSource,
            operations: batch,
            result,
            error: {
                code: 'UNSUPPORTED_EXISTING_COMMENTS',
                message: 'The library returned a comments delta, but this consumer cannot safely merge it with existing comments. No insertion package was prepared.'
            }
        };
    }
    if (!result.hasChanges) {
        return { status: 'noop', source: publicSource, operations: batch, result };
    }
    try {
        return {
            status: 'ready',
            source: publicSource,
            operations: batch,
            result,
            insertionPayload: buildBatchInsertionPackage(source, result)
        };
    } catch (error) {
        return { status: 'error', source: publicSource, operations: batch, result, error };
    }
}

function buildBatchInsertionPackage(source, result) {
    const serializer = createSerializer();
    const outputDoc = parseXmlStrict(result.documentXml, 'batch output');
    normalizeBodySectionOrderStandalone(outputDoc);
    const outputXml = serializer.serializeToString(outputDoc.documentElement);
    const commentSiblings = reconcileCommentSiblingParts({
        commentsXml: result.commentsXml,
        commentsIdsXml: source.commentsIdsXml,
        commentsExtensibleXml: source.commentsExtensibleXml
    });
    let numberingXml = source.numberingXml || null;
    for (const part of result.numberingXmlParts || []) {
        numberingXml = numberingXml ? mergeNumberingXmlBySchemaOrder(numberingXml, part) : part;
    }
    if (source.packageDoc) {
        replacePackagePartXml(source.packageDoc, '/word/document.xml', outputXml);
        if (result.commentsXml) {
            replacePackagePartXml(source.packageDoc, '/word/comments.xml', result.commentsXml);
            ensurePackageRelationship(source.packageDoc, 'comments');
        }
        if (result.commentsExtendedXml) {
            replacePackagePartXml(source.packageDoc, '/word/commentsExtended.xml', result.commentsExtendedXml);
            ensurePackageRelationship(source.packageDoc, 'commentsExtended');
        }
        for (const kind of ['commentsIds', 'commentsExtensible']) {
            const xml = commentSiblings[`${kind}Xml`];
            if (xml) replacePackagePartXml(source.packageDoc, `/${getPartSpec(kind).path}`, xml);
        }
        // A body edit must also repair the commentsExtended type written before 0.8.1.
        const extendedPart = packagePart(source.packageDoc, '/word/commentsExtended.xml');
        if (extendedPart) extendedPart.setAttribute('pkg:contentType', getPartSpec('commentsExtended').contentType);
        if (numberingXml && (result.numberingXmlParts?.length || !source.numberingXml)) {
            replacePackagePartXml(source.packageDoc, '/word/numbering.xml', numberingXml);
            ensurePackageRelationship(source.packageDoc, 'numbering');
        }
        return serializer.serializeToString(source.packageDoc);
    }
    const bodyXml = extractBodyChildElements(outputDoc).map(node => serializer.serializeToString(node)).join('');
    const fragmentPackage = buildDocumentFragmentPackage(bodyXml, {
        includeNumbering: !!numberingXml,
        numberingXml,
        appendTrailingParagraph: false
    });
    if (!result.commentsXml) return fragmentPackage;
    const packageDoc = parseXmlStrict(fragmentPackage, 'batch fragment package');
    replacePackagePartXml(packageDoc, '/word/comments.xml', result.commentsXml);
    ensurePackageRelationship(packageDoc, 'comments');
    if (result.commentsExtendedXml) {
        replacePackagePartXml(packageDoc, '/word/commentsExtended.xml', result.commentsExtendedXml);
        ensurePackageRelationship(packageDoc, 'commentsExtended');
    }
    return serializer.serializeToString(packageDoc);
}
